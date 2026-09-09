"""Tests for the hyperlink tool, once it runs on the OOXML engine.

The tool used to rebuild the linked text into a brand-new run, which lost
whatever formatting the original runs carried and allocated its ``rId`` by
counting the existing ones.  It now wraps the runs that are already there and
lets the package layer allocate and deduplicate the relationship, so the three
things checked below are: the text and its formatting survive a link and an
unlink, the relationships stay in step with the links that reference them, and
nothing outside the target paragraph moves.
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from typing import Any

import pytest

from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.core.hyperlink_writer import (
    HYPERLINK_STYLE,
    add_hyperlink_to_doc,
    list_hyperlinks_in_doc,
    remove_hyperlink_from_doc,
)
from word_document_server.engine.find import find
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import qn
from word_document_server.tools.hyperlink_tools import manage_hyperlinks

W_HYPERLINK = qn("w:hyperlink")
W_R = qn("w:r")
W_T = qn("w:t")
R_ID = qn("r:id")

PROBE_URL = "https://example.org/probe"


def run(coro: Any) -> Any:
    """Drive an ``async`` tool call to completion."""
    return asyncio.run(coro)


def links_of(pkg: DocxPackage) -> list:
    """Every ``w:hyperlink`` of the main story, in document order."""
    return list(pkg.document.iter(W_HYPERLINK))


def external_rels_to(pkg: DocxPackage, url: str) -> list[str]:
    """The rIds of the document part's external relationships pointing at `url`."""
    return [
        rId
        for rId, rel in pkg.document_part.rels.items()
        if rel.is_external and rel.target_ref == url
    ]


def paragraph_index_of(path: Path, text: str) -> int:
    """V2 index of the first paragraph of the body containing `text`."""
    matches = find(DocxPackage.open(path), text, max_results=1)
    assert matches, f"fixture has no paragraph containing {text!r}"
    index = matches[0].index
    assert index is not None
    return index


@pytest.fixture
def simple(fixture_docx) -> Path:
    return fixture_docx("simple")


# --------------------------------------------------------------------------
# add
# --------------------------------------------------------------------------


def test_add_wraps_the_existing_runs_without_touching_the_text(simple: Path) -> None:
    before = visible_text(find(DocxPackage.open(simple), "First paragraph")[0].paragraph)

    result = add_hyperlink_to_doc(str(simple), "First paragraph", PROBE_URL)
    assert result["success"] is True

    pkg = DocxPackage.open(simple)
    (link,) = links_of(pkg)
    assert link.get(R_ID) == result["relationship_id"]
    assert "".join(t.text or "" for t in link.iter(W_T)) == "First paragraph"

    paragraph = next(link.iterancestors(qn("w:p")))
    assert visible_text(paragraph) == before, "the text was rebuilt instead of wrapped"
    assert external_rels_to(pkg, PROBE_URL) == [result["relationship_id"]]
    assert validate_package(simple) == []


def test_add_gives_the_link_the_hyperlink_style_when_the_package_defines_it(
    fixture_docx,
) -> None:
    path = fixture_docx("hyperlinks")  # this fixture defines the Hyperlink style
    result = add_hyperlink_to_doc(str(path), "Anchor target paragraph", PROBE_URL)
    assert result["success"] is True

    pkg = DocxPackage.open(path)
    (added,) = [link for link in links_of(pkg) if link.get(R_ID) == result["relationship_id"]]
    assert [style.get(qn("w:val")) for style in added.iter(qn("w:rStyle"))] == [
        HYPERLINK_STYLE
    ]
    assert validate_package(path) == []


def test_add_falls_back_to_direct_formatting_without_the_hyperlink_style(
    simple: Path,
) -> None:
    # The simple fixture defines no Hyperlink character style: writing a
    # w:rStyle pointing at it would dangle, so the look is written directly.
    assert add_hyperlink_to_doc(str(simple), "First paragraph", PROBE_URL)["success"] is True

    pkg = DocxPackage.open(simple)
    (link,) = links_of(pkg)
    assert list(link.iter(qn("w:rStyle"))) == []
    assert [c.get(qn("w:val")) for c in link.iter(qn("w:color"))] == ["0563C1"]
    assert [u.get(qn("w:val")) for u in link.iter(qn("w:u"))] == ["single"]
    assert validate_package(simple) == []


def test_add_reuses_one_relationship_for_one_url(simple: Path) -> None:
    first = add_hyperlink_to_doc(str(simple), "First paragraph", PROBE_URL)
    second = add_hyperlink_to_doc(str(simple), "Second paragraph", PROBE_URL)
    assert first["success"] and second["success"]
    assert first["relationship_id"] == second["relationship_id"]

    pkg = DocxPackage.open(simple)
    assert len(links_of(pkg)) == 2
    assert len(external_rels_to(pkg, PROBE_URL)) == 1, "the relationship was duplicated"
    assert validate_package(simple) == []


def test_add_restricted_to_one_paragraph_links_that_paragraph(simple: Path) -> None:
    index = paragraph_index_of(simple, "Second paragraph")
    result = add_hyperlink_to_doc(str(simple), "paragraph", PROBE_URL, paragraph_index=index)
    assert result["success"] is True
    assert result["paragraph_index"] == index

    pkg = DocxPackage.open(simple)
    (link,) = links_of(pkg)
    paragraph = next(link.iterancestors(qn("w:p")))
    assert visible_text(paragraph).startswith("Second paragraph")


def test_add_refuses_an_index_out_of_range_and_leaves_the_file_alone(simple: Path) -> None:
    before = simple.read_bytes()
    result = add_hyperlink_to_doc(str(simple), "paragraph", PROBE_URL, paragraph_index=999)
    assert result["success"] is False
    assert "out of range" in result["error"]
    assert simple.read_bytes() == before


def test_add_reports_a_missing_target_and_leaves_the_file_alone(simple: Path) -> None:
    before = simple.read_bytes()
    result = add_hyperlink_to_doc(str(simple), "no such text", PROBE_URL)
    assert result["success"] is False
    assert "not found" in result["error"]
    assert simple.read_bytes() == before


def test_add_refuses_to_nest_a_link_inside_a_link(fixture_docx) -> None:
    path = fixture_docx("hyperlinks")
    before = path.read_bytes()
    result = add_hyperlink_to_doc(str(path), "example.org", PROBE_URL)
    assert result["success"] is False
    assert "nested" in result["error"]
    assert path.read_bytes() == before, "a refused link still rewrote the document"


def test_add_touches_nothing_but_the_target_paragraph_and_the_rels(fixture_docx) -> None:
    path = fixture_docx("combined")
    before = snapshot(path.read_bytes())
    index = next(
        i
        for i, sig in enumerate(before.story_paragraphs("document"))
        if sig.text == "Second heading"
    )

    assert add_hyperlink_to_doc(str(path), "Second heading", PROBE_URL)["success"] is True

    after = snapshot(path.read_bytes())
    assert_unchanged_except(
        before,
        after,
        paragraphs=[index],
        parts=["word/_rels/document.xml.rels"],
        counters=["hyperlinks"],
    )
    assert validate_package(path) == []


# --------------------------------------------------------------------------
# remove
# --------------------------------------------------------------------------


def test_remove_keeps_the_runs_and_drops_the_relationship(fixture_docx) -> None:
    path = fixture_docx("hyperlinks")
    pkg = DocxPackage.open(path)
    (linked,) = [link for link in links_of(pkg) if link.get(R_ID) is not None]
    rId = linked.get(R_ID)
    target = pkg.rel_target(pkg.document, rId)

    result = remove_hyperlink_from_doc(str(path), "example.org")
    assert result["success"] is True
    assert result["removed"] == [
        {"text": "example.org", "url": target, "relationship_id": rId}
    ]
    assert result["relationships_removed"] == [rId]

    after = DocxPackage.open(path)
    assert [link.get(R_ID) for link in links_of(after)] == [None], "only the anchor link is left"
    assert rId not in after.document_part.rels
    # The run stayed where it was, with its text and its character style.
    (unlinked,) = [
        element
        for element in after.document.iter(W_R)
        if "".join(t.text or "" for t in element.iter(W_T)) == "example.org"
    ]
    assert unlinked.getparent().tag == qn("w:p")
    assert [style.get(qn("w:val")) for style in unlinked.iter(qn("w:rStyle"))] == [
        HYPERLINK_STYLE
    ]
    assert validate_package(path) == []


def test_remove_keeps_a_relationship_another_link_still_uses(simple: Path) -> None:
    add_hyperlink_to_doc(str(simple), "First paragraph", PROBE_URL)
    add_hyperlink_to_doc(str(simple), "Second paragraph", PROBE_URL)

    result = remove_hyperlink_from_doc(str(simple), "First paragraph")
    assert result["success"] is True
    assert result["relationships_removed"] == [], "a still-referenced relationship was dropped"

    pkg = DocxPackage.open(simple)
    assert len(links_of(pkg)) == 1
    assert len(external_rels_to(pkg, PROBE_URL)) == 1
    assert validate_package(simple) == []


def test_remove_reports_a_missing_link_and_leaves_the_file_alone(fixture_docx) -> None:
    path = fixture_docx("hyperlinks")
    before = path.read_bytes()
    result = remove_hyperlink_from_doc(str(path), "no such link")
    assert result["success"] is False
    assert "No hyperlink found" in result["error"]
    assert path.read_bytes() == before


def test_remove_needs_something_to_aim_at(simple: Path) -> None:
    result = remove_hyperlink_from_doc(str(simple))
    assert result["success"] is False
    assert "required" in result["error"]


# --------------------------------------------------------------------------
# list
# --------------------------------------------------------------------------


def test_list_reports_external_and_anchor_links_with_their_story(fixture_docx) -> None:
    listed = list_hyperlinks_in_doc(str(fixture_docx("hyperlinks")))
    assert listed["success"] is True
    assert listed["total_hyperlinks"] == len(listed["hyperlinks"]) == 2

    external, internal = listed["hyperlinks"]
    assert external["text"] == "example.org"
    assert external["url"] == "https://example.org/fixture"
    assert external["anchor"] is None
    assert external["story"] == "document"
    assert isinstance(external["paragraph_index"], int)

    assert internal["text"] == "jump to the anchor"
    assert internal["url"] == "#FixtureAnchor"
    assert internal["anchor"] == "FixtureAnchor"
    assert internal["relationship_id"] is None


# --------------------------------------------------------------------------
# The MCP tool
# --------------------------------------------------------------------------


def test_manage_hyperlinks_keeps_the_add_response_shape(simple: Path) -> None:
    payload = json.loads(
        run(manage_hyperlinks(str(simple), action="add", text="First paragraph", url=PROBE_URL))
    )
    assert payload["success"] is True
    assert payload["text"] == "First paragraph"
    assert payload["url"] == PROBE_URL
    assert payload["relationship_id"].startswith("rId")
    assert PROBE_URL in payload["message"]


def test_manage_hyperlinks_round_trips_add_then_remove(simple: Path) -> None:
    assert json.loads(
        run(manage_hyperlinks(str(simple), action="add", text="First paragraph", url=PROBE_URL))
    )["success"]
    payload = json.loads(
        run(manage_hyperlinks(str(simple), action="remove", text="First paragraph"))
    )
    assert payload["success"] is True
    assert payload["removed_count"] == 1

    pkg = DocxPackage.open(simple)
    assert links_of(pkg) == []
    assert external_rels_to(pkg, PROBE_URL) == []
    assert validate_package(simple) == []


def test_manage_hyperlinks_lists_through_the_tool(fixture_docx) -> None:
    payload = json.loads(run(manage_hyperlinks(str(fixture_docx("hyperlinks")), action="list")))
    assert payload["success"] is True
    assert payload["total_hyperlinks"] == 2


def test_manage_hyperlinks_rejects_an_unknown_action(simple: Path) -> None:
    payload = json.loads(run(manage_hyperlinks(str(simple), action="explode")))
    assert payload["success"] is False
    assert "Unknown action" in payload["error"]
    assert "'add'" in payload["error"] and "'remove'" in payload["error"]


def test_manage_hyperlinks_reports_a_missing_document(tmp_path: Path) -> None:
    payload = json.loads(run(manage_hyperlinks(str(tmp_path / "absent.docx"), action="list")))
    assert payload["success"] is False
    assert "does not exist" in payload["error"]
