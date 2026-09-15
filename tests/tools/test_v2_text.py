"""The V2 text surface: the shape of every answer, and what an agent can trust.

Four tools, one envelope.  What is pinned here is less the four operations --
the engine layers below already have their own tests -- than the three promises
the tool layer makes on top of them:

*the report is always the same shape*
    ``status``, ``dry_run``, ``changes`` and ``warnings`` are present whatever
    happened, so a caller never has to guess whether a key exists before
    reading it, and a failure is a value, not an exception.

*an error keeps the engine's own code*
    D-022: ``invalid``, ``not_found``, ``ambiguous`` and ``stale_anchor`` come
    back untouched.  A tool layer that renamed them would force every caller to
    learn a second vocabulary for the same four situations.

*one index space*
    D-016: the ``paragraph`` of a ``changes`` entry, the ``paragraph`` of a
    ``doc_find`` match and the ``index`` of a ``doc_inspect`` block are the same
    number for the same paragraph, and that number is what a ``{"paragraph": i}``
    locator takes.  ``null`` when the paragraph is in a table cell or a text
    box, which have no place in that space.

The tools are called through :func:`~word_document_server.tools.v2.registry.v2_tools`
rather than by importing the functions: that is the wrapper an MCP client
reaches, and the envelope and the error mapping only exist there.

``simple`` is the fixture of choice for the fidelity assertions: it holds no
table, no text box and no content control, so the snapshot's paragraph index
space and the V2 one coincide there and an allowed paragraph can be named by a
bare integer.
"""

from __future__ import annotations

import asyncio
import zipfile
from pathlib import Path
from typing import Any

import pytest
from fastmcp import FastMCP

from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.engine.errors import (
    InvalidText,
    LocatorError,
    PackageError,
    UnsupportedRange,
)
from word_document_server.tools.v2.registry import (
    discover_tool_specs,
    error_code,
    register_v2_tools,
    v2_tools,
)

#: The tools this module is about.
TEXT_TOOLS = ("doc_inspect", "doc_find", "doc_edit_text", "doc_format_range")

#: Keys the registry fills in on every successful call.
ENVELOPE = {"status", "dry_run", "changes", "warnings"}


def call(name: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Call a V2 tool the way an MCP client reaches it, and return its report."""
    return asyncio.run(v2_tools()[name](*args, **kwargs))


def texts(report: dict[str, Any]) -> list[str]:
    """The ``after`` text of every change in `report`."""
    return [change["after"] for change in report["changes"]]


def document_xml(path: Path) -> str:
    """The main document part of `path`, as text."""
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


# --------------------------------------------------------------------------
# The envelope
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("tool", "extra"),
    [
        ("doc_inspect", {}),
        ("doc_find", {"pattern": "paragraph"}),
        ("doc_edit_text", {"locator": {"paragraph": 1}, "text": "x", "dry_run": True}),
        (
            "doc_format_range",
            {"locator": {"paragraph": 1}, "patch": {"bold": True}, "dry_run": True},
        ),
    ],
)
def test_every_report_carries_the_four_shared_keys(fixture_docx, tool, extra):
    report = call(tool, str(fixture_docx("simple")), **extra)

    assert report["status"] == "ok"
    assert ENVELOPE <= set(report)
    assert isinstance(report["changes"], list)
    assert isinstance(report["warnings"], list)
    assert report["dry_run"] is extra.get("dry_run", False)


@pytest.mark.parametrize(
    ("locator", "code"),
    [
        ({"paragraph": "one"}, "invalid"),
        ({"chapter": 1}, "invalid"),
        ({"paragraph": 99}, "not_found"),
        ({"story": "header9", "paragraph": 0}, "not_found"),
        ({"find": "paragraph"}, "ambiguous"),
        ({"paragraph": 1, "expect_text": "Not what is there"}, "stale_anchor"),
    ],
)
def test_a_locator_failure_keeps_the_engine_code_and_writes_nothing(
    fixture_docx, locator, code
):
    """D-022: the four locator codes are the tool layer's codes too."""
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call("doc_edit_text", str(path), locator, "replace", "rewritten")

    assert report == {"status": "error", "code": code, "message": report["message"]}
    assert report["message"]
    assert path.read_bytes() == before


def test_an_engine_refusal_keeps_its_own_reason_as_the_code(fixture_docx):
    path = fixture_docx("simple")

    report = call("doc_edit_text", str(path), {"paragraph": 1}, "delete", start=0, end=9_999)

    assert report["status"] == "error"
    assert report["code"] == "out-of-range"


def test_a_missing_file_is_a_report_not_an_exception(tmp_path):
    report = call("doc_inspect", str(tmp_path / "absent.docx"))

    assert report["status"] == "error"
    assert report["code"] == "file_not_found"


def test_a_file_that_is_not_a_package_is_reported(tmp_path):
    path = tmp_path / "not-a-docx.docx"
    path.write_bytes(b"this is not a zip")

    report = call("doc_inspect", str(path))

    assert report["status"] == "error"
    assert report["code"] == "package_error"


@pytest.mark.parametrize(
    ("exception", "code"),
    [
        (LocatorError("ambiguous", "two of them"), "ambiguous"),
        (UnsupportedRange("crosses-table-cell", "no"), "crosses-table-cell"),
        (InvalidText("a control character"), "invalid_text"),
        (PackageError("no main part"), "package_error"),
        (FileNotFoundError(2, "No such file", "x.docx"), "file_not_found"),
        (PermissionError(13, "Permission denied", "x.docx"), "io_error"),
        (ValueError("empty pattern"), "invalid_argument"),
        (TypeError("not a mapping"), "invalid_argument"),
    ],
)
def test_each_failure_family_maps_to_a_stable_code(exception, code):
    assert error_code(exception) == code


# --------------------------------------------------------------------------
# Discovery and registration
# --------------------------------------------------------------------------


def test_the_text_tools_are_discovered_without_being_named_anywhere():
    discovered = [spec.name for spec in discover_tool_specs()]

    assert set(TEXT_TOOLS) <= set(discovered)
    assert len(discovered) == len(set(discovered))


def test_registering_derives_name_schema_and_description_from_the_function():
    server = FastMCP("test")

    registered = register_v2_tools(server)

    assert set(TEXT_TOOLS) <= set(registered)
    tools = {tool.name: tool for tool in asyncio.run(server.list_tools())}
    find = tools["doc_find"]
    assert find.parameters["required"] == ["filename", "pattern"]
    assert find.parameters["properties"]["regex"]["type"] == "boolean"
    assert find.description.startswith(
        "Find where a string occurs in a Word document"
    )
    # The signature, not a hand-written schema: an enum comes from the Literal.
    edit = tools["doc_edit_text"]
    assert edit.parameters["properties"]["action"]["enum"] == [
        "replace",
        "insert",
        "delete",
    ]
    assert tools["doc_inspect"].annotations.read_only_hint is True
    assert edit.annotations.destructive_hint is True


# --------------------------------------------------------------------------
# One index space -- D-016
# --------------------------------------------------------------------------


def test_inspect_find_and_edit_agree_on_the_index_of_a_paragraph(fixture_docx):
    """The three surfaces number the same paragraph the same way.

    ``content_controls`` is the fixture that used to tell the spaces apart: it
    holds a block ``w:sdt``, whose paragraph counts in the V2 space and does not
    count in python-docx's body list.
    """
    path = fixture_docx("content_controls")
    anchor = "Paragraph inside a block content control."

    found = call("doc_find", str(path), anchor)["matches"]
    assert len(found) == 1
    index = found[0]["paragraph"]

    blocks = call("doc_inspect", str(path))["document"]["blocks"]
    assert blocks[index]["index"] == index
    assert blocks[index]["text"] == anchor

    reached = call(
        "doc_edit_text", str(path), {"paragraph": index}, "delete", dry_run=True
    )
    assert reached["changes"][0]["paragraph"] == index
    assert reached["changes"][0]["before"] == anchor


@pytest.mark.parametrize(
    ("fixture", "pattern"),
    [
        ("tables", "Row two, column two"),
        ("text_boxes", "Text inside the VML text box."),
    ],
)
def test_a_paragraph_outside_the_index_space_reports_a_null_index(
    fixture_docx, fixture, pattern
):
    matches = call("doc_find", str(fixture_docx(fixture)), pattern)["matches"]

    assert [match["paragraph"] for match in matches] == [None]


def test_an_edit_in_a_table_cell_reports_a_null_index_too(fixture_docx):
    path = fixture_docx("tables")

    report = call(
        "doc_edit_text",
        str(path),
        {"table": 0, "row": 1, "col": 1},
        "replace",
        "Rewritten cell",
    )

    assert report["changes"] == [
        {
            "story": "document",
            "paragraph": None,
            "before": "Row two, column two",
            "after": "Rewritten cell",
        }
    ]
    assert validate_package(path) == []


# --------------------------------------------------------------------------
# doc_inspect
# --------------------------------------------------------------------------


def test_inspect_describes_the_document_without_touching_it(fixture_docx):
    path = fixture_docx("tables")
    before = path.read_bytes()

    document = call("doc_inspect", str(path))["document"]

    assert {"stories", "counts", "blocks", "tables", "sections"} <= set(document)
    assert document["counts"]["tables"] == len(document["tables"])
    assert [table["index"] for table in document["tables"]] == [0, 1]
    assert path.read_bytes() == before


# --------------------------------------------------------------------------
# doc_find
# --------------------------------------------------------------------------


def test_find_defaults_to_the_body_and_reports_the_real_story_id(fixture_docx):
    """D-006: ``body`` goes in, ``document`` comes out."""
    path = str(fixture_docx("simple"))

    implicit = call("doc_find", path, "paragraph")
    explicit = call("doc_find", path, "paragraph", stories=["body"])

    assert implicit["matches"] == explicit["matches"]
    assert {match["story"] for match in implicit["matches"]} == {"document"}


def test_find_searches_the_stories_it_is_given_and_skips_the_absent_ones(fixture_docx):
    path = str(fixture_docx("headers_footers"))

    matches = call(
        "doc_find", path, "Fixture", stories=["body", "header1", "footnotes"]
    )["matches"]

    assert [match["story"] for match in matches] == ["document", "header1"]


def test_max_results_truncates_and_says_so(fixture_docx):
    path = str(fixture_docx("simple"))

    full = call("doc_find", path, "paragraph")
    cut = call("doc_find", path, "paragraph", max_results=1)

    assert len(full["matches"]) == 2
    assert full["truncated"] is False
    assert full["warnings"] == []
    assert cut["matches"] == full["matches"][:1]
    assert cut["truncated"] is True
    assert cut["warnings"]


def test_a_max_results_equal_to_the_number_of_matches_is_not_truncated(fixture_docx):
    report = call("doc_find", str(fixture_docx("simple")), "paragraph", max_results=2)

    assert len(report["matches"]) == 2
    assert report["truncated"] is False


@pytest.mark.parametrize(
    ("pattern", "kwargs", "expected"),
    [
        ("Secon", {}, 2),
        ("Secon", {"case": False}, 3),
        ("Secon", {"whole_word": True}, 0),
        ("S[ei]cond", {}, 0),
        ("S[ei]cond", {"regex": True}, 2),
    ],
)
def test_each_matching_flag_reaches_the_engine(fixture_docx, pattern, kwargs, expected):
    """``simple`` reads ``Second`` twice and ``second`` once, and ``Secon`` alone
    is never a whole word -- so each flag moves the count on its own."""
    path = str(fixture_docx("simple"))

    matches = call("doc_find", path, pattern, **kwargs)["matches"]

    assert len(matches) == expected


@pytest.mark.parametrize("pattern", ["", "(?:)"])
def test_a_pattern_that_cannot_be_searched_is_refused(fixture_docx, pattern):
    report = call("doc_find", str(fixture_docx("simple")), pattern, regex=True)

    assert report["status"] == "error"
    assert report["code"] == "invalid_argument"


def test_a_negative_max_results_is_refused(fixture_docx):
    report = call("doc_find", str(fixture_docx("simple")), "paragraph", max_results=-1)

    assert report["status"] == "error"
    assert report["code"] == "invalid_argument"


# --------------------------------------------------------------------------
# doc_edit_text
# --------------------------------------------------------------------------


def test_replace_rewrites_the_located_span_and_leaves_the_rest_alone(fixture_docx):
    path = fixture_docx("simple")
    before = snapshot(path)

    report = call(
        "doc_edit_text",
        str(path),
        {"find": "Second paragraph", "expect_text": "Second paragraph,"},
        "replace",
        "Deuxieme paragraphe",
    )

    assert report["status"] == "ok"
    assert report["saved"] is True
    assert texts(report) == ["Deuxieme paragraphe, with a trailing sentence."]
    assert validate_package(path) == []
    assert_unchanged_except(before, snapshot(path), paragraphs=[2])


def test_insert_and_delete_act_at_the_offsets_find_reported(fixture_docx):
    path = fixture_docx("simple")
    match = call("doc_find", str(path), "trailing")["matches"][0]

    inserted = call(
        "doc_edit_text",
        str(path),
        {"paragraph": match["paragraph"]},
        "insert",
        "very ",
        start=match["start"],
    )
    assert texts(inserted) == ["Second paragraph, with a very trailing sentence."]

    deleted = call(
        "doc_edit_text",
        str(path),
        {"paragraph": match["paragraph"]},
        "delete",
        start=match["start"],
        end=match["start"] + len("very "),
    )
    assert texts(deleted) == ["Second paragraph, with a trailing sentence."]
    assert validate_package(path) == []


def test_without_offsets_the_edit_covers_the_span_the_locator_resolved(fixture_docx):
    path = fixture_docx("simple")

    whole = call(
        "doc_edit_text", str(path), {"paragraph": 1}, "replace", "All of it",
        dry_run=True,
    )
    span = call(
        "doc_edit_text",
        str(path),
        {"find": "First", "occurrence": 1},
        "replace",
        "Premier",
        dry_run=True,
    )

    assert texts(whole) == ["All of it"]
    assert texts(span) == ["Premier paragraph of the simple fixture."]


def test_dry_run_reports_the_whole_edit_and_writes_nothing(fixture_docx):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call(
        "doc_edit_text", str(path), {"paragraph": 1}, "replace", "Rewritten",
        dry_run=True,
    )

    assert report["dry_run"] is True
    assert report["saved"] is False
    assert report["changes"] == [
        {
            "story": "document",
            "paragraph": 1,
            "before": "First paragraph of the simple fixture.",
            "after": "Rewritten",
        }
    ]
    assert path.read_bytes() == before


def test_a_tracked_edit_keeps_the_replaced_text_in_the_package(fixture_docx):
    path = fixture_docx("simple")

    report = call(
        "doc_edit_text",
        str(path),
        {"find": "Second paragraph", "occurrence": 1},
        "replace",
        "Deuxieme paragraphe",
        track_changes=True,
        author="Reviewer",
    )

    assert texts(report) == ["Deuxieme paragraphe, with a trailing sentence."]
    xml = document_xml(path)
    assert "w:del" in xml
    assert "w:ins" in xml
    assert 'w:author="Reviewer"' in xml
    # The old text is hidden, not dropped: rejecting the revision restores it.
    assert "Second paragraph" in xml
    assert validate_package(path) == []


def test_a_tracked_insert_folded_into_the_authors_own_revision_warns(fixture_docx):
    path = fixture_docx("simple")
    call(
        "doc_edit_text",
        str(path),
        {"paragraph": 1},
        "insert",
        "Added. ",
        start=0,
        track_changes=True,
        author="Reviewer",
    )

    again = call(
        "doc_edit_text",
        str(path),
        {"paragraph": 1},
        "insert",
        "More. ",
        start=len("Added. "),
        track_changes=True,
        author="Reviewer",
    )

    assert again["status"] == "ok"
    assert again["warnings"]
    assert texts(again) == ["Added. More. First paragraph of the simple fixture."]


def test_an_empty_insert_changes_nothing_says_so_and_does_not_rewrite_the_file(
    fixture_docx,
):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call("doc_edit_text", str(path), {"paragraph": 1}, "insert", "")

    assert report["status"] == "ok"
    assert report["saved"] is False
    assert report["warnings"]
    assert texts(report) == ["First paragraph of the simple fixture."]
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    "kwargs",
    [
        {"action": "delete", "text": "something"},
        {"action": "insert", "text": "something", "end": 3},
    ],
)
def test_an_argument_the_action_cannot_use_is_refused(fixture_docx, kwargs):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call("doc_edit_text", str(path), {"paragraph": 1}, **kwargs)

    assert report["status"] == "error"
    assert report["code"] == "invalid_argument"
    assert path.read_bytes() == before


def test_an_edit_reaches_a_story_other_than_the_body(fixture_docx):
    path = fixture_docx("headers_footers")

    report = call(
        "doc_edit_text",
        str(path),
        {"story": "header1", "paragraph": 0},
        "replace",
        "Rewritten header",
    )

    assert report["changes"][0]["story"] == "header1"
    assert texts(report) == ["Rewritten header"]
    assert validate_package(path) == []


# --------------------------------------------------------------------------
# doc_format_range
# --------------------------------------------------------------------------


def test_formatting_patches_the_runs_of_the_range_without_changing_the_text(
    fixture_docx,
):
    path = fixture_docx("simple")

    report = call(
        "doc_format_range",
        str(path),
        {"find": "trailing"},
        {"bold": True, "size_pt": 14},
    )

    assert report["status"] == "ok"
    assert report["runs"] == 1
    change = report["changes"][0]
    assert change["before"] == change["after"] == (
        "Second paragraph, with a trailing sentence."
    )
    assert "<w:b/>" in document_xml(path)
    assert validate_package(path) == []


def test_format_dry_run_writes_nothing(fixture_docx):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call(
        "doc_format_range", str(path), {"paragraph": 1}, {"italic": True}, dry_run=True
    )

    assert report["dry_run"] is True
    assert report["saved"] is False
    assert report["runs"] >= 1
    assert path.read_bytes() == before


@pytest.mark.parametrize("patch", [{}, {"glowing": True}, {"size_pt": 11.3}])
def test_a_patch_the_schema_cannot_store_is_refused(fixture_docx, patch):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = call("doc_format_range", str(path), {"paragraph": 1}, patch)

    assert report["status"] == "error"
    assert report["code"] == "invalid_argument"
    assert path.read_bytes() == before
