"""End-to-end scenarios: the ``doc_*`` tools chained the way an agent chains them.

Every other test module in this suite pins one tool, or one engine layer, in
isolation. This one instead walks the loop an agent actually runs on a
document somebody else wrote -- inspect, find, edit, comment, format, list,
table style, audit, compare -- on ``combined``, the fixture that carries every
OOXML feature this server knows about at once, and checks that the *whole*
trip leaves the package internally consistent (``validate_package``), reopens
in LibreOffice when it is available, and that ``doc_compare`` against an
untouched copy of the same fixture reports exactly the edits this test asked
for and nothing else.

The four refusal scenarios below are not an edge case of one tool: they are
the answer this server gives instead of guessing, on constructs the engine
Risks call out by name -- a field cut in half, a locator naming a paragraph
that has moved under it, a search that does not say which of several matches
it means, and an edit that would land inside a tracked deletion (D-005,
D-006/D-009/D-016 do not apply here; the fixtures below stay in the main
story's V2 index space).
"""

from __future__ import annotations

import asyncio
import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from tests.fixtures.builders import build
from tests.support.libreoffice import opens_in_libreoffice, requires_libreoffice
from tests.support.package_check import validate_package
from word_document_server.tools.comment_write_tools import add_comment
from word_document_server.tools.v2.registry import v2_tools


def call(name: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Call a V2 tool the way an MCP client reaches it, and return its report."""
    return asyncio.run(v2_tools()[name](*args, **kwargs))


# --------------------------------------------------------------------------
# The full agent loop
# --------------------------------------------------------------------------


def _run_agent_edits(path: Path) -> dict[str, Any]:
    """Run the writing half of the agent loop against `path`, in place.

    A tracked-change edit, a comment, a character style, a two-item numbered
    list and a table style -- the six families of write this server offers,
    minus the ones J06-P3's sibling reviews already cover on their own
    fixtures. Returns what was actually done (paragraph indices, before/after
    text) so the two tests below build their expectations from that instead
    of a second, independently-typed copy of the same numbers.
    """
    matches = call(
        "doc_find", str(path), "First paragraph of the simple fixture."
    )["matches"]
    assert len(matches) == 1
    revised_paragraph = matches[0]["paragraph"]

    edited = call(
        "doc_edit_text",
        str(path),
        {"paragraph": revised_paragraph},
        "replace",
        "First paragraph, revised.",
        track_changes=True,
        author="E2E Reviewer",
    )
    assert edited["status"] == "ok"
    assert edited["saved"] is True

    anchor = call("doc_find", str(path), "Custom body style.")["matches"]
    assert len(anchor) == 1
    comment = json.loads(
        asyncio.run(
            add_comment(
                str(path),
                "Custom body style.",
                "Please confirm this figure.",
                author="E2E Reviewer",
            )
        )
    )
    assert comment["success"] is True

    styled = call(
        "doc_format_range",
        str(path),
        {"find": "Plain, then "},
        {"char_style": "FixtureEmphasis"},
    )
    assert styled["status"] == "ok"
    assert styled["runs"] == 1

    listed = call(
        "doc_apply_list",
        str(path),
        [
            {"find": "Body text under the second heading."},
            {"find": "Custom note style."},
        ],
        kind="decimal",
    )
    assert listed["status"] == "ok"
    assert listed["warnings"] == []
    assert len(listed["changes"]) == 2

    tabled = call("doc_apply_table_style", str(path), 0, "TableGrid")
    assert tabled["status"] == "ok"
    assert tabled["style"] == "TableGrid"

    return {
        "revised": (revised_paragraph, matches[0]["text"], "First paragraph, revised."),
        "commented": anchor[0]["paragraph"],
        "styled": styled["changes"][0],
        "listed": listed["changes"],
    }


def test_the_full_agent_loop_on_combined(fixture_docx, tmp_path):
    """inspect -> find -> tracked edit -> comment -> character style -> list ->
    table style -> audit -> compare -> validate.

    ``doc_compare`` against an untouched copy of the same fixture is the check
    that closes the loop: every paragraph, table and counter it reports as
    changed is one this test meant to touch, and nothing else moved.
    """
    original_path = tmp_path / "combined-original.docx"
    original_path.write_bytes(build("combined"))
    path = fixture_docx("combined")

    # 1. doc_inspect: the lay of the land.
    document = call("doc_inspect", str(path))["document"]
    assert document["counts"]["tables"] == 2

    done = _run_agent_edits(path)

    # doc_audit: read-only.
    before_audit = path.read_bytes()
    audit = call("doc_audit", str(path))["audit"]
    assert audit["counts"]["total"] > 0
    assert path.read_bytes() == before_audit

    # doc_compare against the untouched original.
    comparison = call("doc_compare", str(original_path), str(path))
    assert comparison["identical"] is False
    assert comparison["paragraphs"]["removed"] == []
    assert comparison["paragraphs"]["added"] == [
        {"story": "comments", "paragraph": None, "text": "Please confirm this figure."}
    ]

    expected_changes = {
        done["revised"],
        (done["commented"], "Custom body style.", "Custom body style."),
        (done["styled"]["paragraph"], done["styled"]["before"], done["styled"]["after"]),
        *(
            (change["paragraph"], change["before"], change["after"])
            for change in done["listed"]
        ),
    }
    actual_changes = {
        (change["paragraph"], change["text_before"], change["text_after"])
        for change in comparison["paragraphs"]["changed"]
    }
    assert actual_changes == expected_changes

    fields_by_paragraph = {
        change["paragraph"]: set(change["fields"])
        for change in comparison["paragraphs"]["changed"]
    }
    assert "text" in fields_by_paragraph[done["revised"][0]]
    assert "markers" in fields_by_paragraph[done["commented"]]
    assert "runs" in fields_by_paragraph[done["styled"]["paragraph"]]
    for change in done["listed"]:
        assert "num_pr" in fields_by_paragraph[change["paragraph"]]

    assert comparison["tables"] == {
        "added": [],
        "removed": [],
        "changed": [{"story": "document", "table": 0, "fields": ["tbl_pr"]}],
    }
    assert comparison["sections_changed"] == []
    assert {item["field"] for item in comparison["counters_changed"]} == {
        "comments",
        "comment_references",
        "revisions",
    }

    # No part outside the ones these six edits necessarily touch was rewritten.
    assert comparison["parts"]["removed"] == []
    untouched_parts = {"word/styles.xml", "word/theme/theme1.xml", "word/settings.xml"}
    with zipfile.ZipFile(path) as archive:
        package_names = set(archive.namelist())
    assert untouched_parts <= package_names, untouched_parts - package_names
    assert not untouched_parts & set(comparison["parts"]["changed"])

    # validate_package: no dangling relationship either way.
    assert validate_package(path) == []


@requires_libreoffice
@pytest.mark.libreoffice
def test_the_edited_document_reopens_in_libreoffice(fixture_docx):
    """The output of the same agent loop is still a document LibreOffice can open.

    Optional, and never a source of truth (see `tests/support/libreoffice.py`):
    skipped outright when `soffice` is not on `PATH`.
    """
    path = fixture_docx("combined")

    _run_agent_edits(path)

    assert opens_in_libreoffice(path) is True


# --------------------------------------------------------------------------
# One refusal scenario per reason
# --------------------------------------------------------------------------


def test_an_edit_that_would_cut_a_field_in_half_is_refused(fixture_docx):
    """A range that overlaps a field without containing it is `crosses-field`."""
    path = fixture_docx("fields")
    before = path.read_bytes()

    report = call(
        "doc_edit_text", str(path), {"paragraph": 2}, "replace", "X", start=0, end=5
    )

    assert report["status"] == "error"
    assert report["code"] == "crosses-field"
    assert path.read_bytes() == before


def test_an_edit_behind_a_stale_anchor_is_refused(fixture_docx):
    """`expect_text` that no longer matches fails loudly instead of editing blind."""
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call(
        "doc_edit_text",
        str(path),
        {"paragraph": 1, "expect_text": "This is not what is there"},
        "replace",
        "X",
    )

    assert report["status"] == "error"
    assert report["code"] == "stale_anchor"
    assert path.read_bytes() == before


def test_an_ambiguous_search_is_refused(fixture_docx):
    """A `find` locator with more than one match and no `occurrence` is refused."""
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call("doc_edit_text", str(path), {"find": "paragraph"}, "replace", "X")

    assert report["status"] == "error"
    assert report["code"] == "ambiguous"
    assert path.read_bytes() == before


def test_an_edit_across_an_existing_revision_is_refused(fixture_docx):
    """A range that strictly contains a tracked deletion is `hidden-content`.

    `doc_find` matches on visible text, which does not include text hidden
    under a `w:del`; the match it returns here spans across the deletion
    without knowing it, exactly the situation `doc_edit_text` has to catch on
    its own.
    """
    path = fixture_docx("tracked_changes")
    before = path.read_bytes()

    matches = call("doc_find", str(path), "inserted text, and kept")["matches"]
    assert len(matches) == 1
    match = matches[0]

    report = call(
        "doc_edit_text",
        str(path),
        {"paragraph": match["paragraph"]},
        "replace",
        "X",
        start=match["start"],
        end=match["end"],
    )

    assert report["status"] == "error"
    assert report["code"] == "hidden-content"
    assert path.read_bytes() == before
