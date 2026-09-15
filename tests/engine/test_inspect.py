"""Tests for :mod:`word_document_server.engine.inspect`.

The report exists to be read and then acted on, so most cases below check both
halves at once: the number inspection prints is the number a locator takes back.
"""

from __future__ import annotations

import json

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.snapshot import diff, snapshot
from word_document_server.engine.find import find
from word_document_server.engine.inspect import TEXT_PREVIEW, inspect
from word_document_server.engine.locators import indexed_paragraphs, resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.revisions import list_revisions
from word_document_server.engine.xmlns import qn

W_NAME = qn("w:name")
W_ID = qn("w:id")


def _pkg(name: str) -> DocxPackage:
    return DocxPackage.open(build(name))


def _by_index(report: dict, key: str = "blocks") -> dict[int, dict]:
    return {entry["index"]: entry for entry in report[key]}


# --------------------------------------------------------------------------------------
# Shape of the report
# --------------------------------------------------------------------------------------


def test_the_report_has_the_documented_keys() -> None:
    report = inspect(_pkg("combined"))
    assert set(report) == {
        "stories",
        "counts",
        "blocks",
        "tables",
        "sections",
        "styles_used",
        "bookmarks",
        "fields",
    }


def test_the_report_is_plain_json_ready_data() -> None:
    # No lxml element may leak out: the report crosses the MCP boundary, and a
    # live element there would read as a handle the caller could keep.
    report = inspect(_pkg("combined"))
    assert json.loads(json.dumps(report)) == report


def test_inspecting_never_touches_the_package() -> None:
    source = build("combined")
    pkg = DocxPackage.open(source)
    inspect(pkg)
    assert diff(snapshot(source), snapshot(pkg.to_bytes())).is_empty()


# --------------------------------------------------------------------------------------
# Blocks
# --------------------------------------------------------------------------------------


def test_every_block_index_is_a_usable_locator() -> None:
    pkg = _pkg("combined")
    for block in inspect(pkg)["blocks"]:
        target = resolve(pkg, {"paragraph": block["index"]})
        assert target.index == block["index"]
        assert target.text.startswith(block["text"])
        assert block["truncated"] == (len(target.text) > TEXT_PREVIEW)


def test_blocks_cover_the_v2_space_of_the_main_story_only() -> None:
    pkg = _pkg("combined")
    report = inspect(pkg)
    assert len(report["blocks"]) == len(indexed_paragraphs(pkg.document))
    assert [block["index"] for block in report["blocks"]] == list(
        range(len(report["blocks"]))
    )


def test_a_text_box_paragraph_is_not_a_block() -> None:
    report = inspect(_pkg("text_boxes"))
    assert [block["text"] for block in report["blocks"]] == [
        "Fixture: text boxes",
        "Body text before the boxes.",
        "Anchor of the VML box.",
        "Anchor of the DrawingML box.",
        "Body text after the boxes.",
    ]


def test_a_block_content_control_paragraph_is_a_block() -> None:
    report = inspect(_pkg("content_controls"))
    assert report["blocks"][1]["text"] == "Paragraph inside a block content control."


def test_a_cell_paragraph_is_not_a_block() -> None:
    report = inspect(_pkg("tables"))
    assert [block["text"] for block in report["blocks"]] == [
        "Fixture: tables",
        "Paragraph after the table.",
    ]


def test_headings_report_their_style_and_level() -> None:
    blocks = _by_index(inspect(_pkg("simple")))
    assert (blocks[0]["kind"], blocks[0]["style"], blocks[0]["heading_level"]) == (
        "heading",
        "Heading1",
        1,
    )
    assert (blocks[3]["kind"], blocks[3]["style"], blocks[3]["heading_level"]) == (
        "heading",
        "Heading2",
        2,
    )
    assert blocks[1]["kind"] == "paragraph"
    assert blocks[1]["heading_level"] is None


def test_list_items_report_their_numbering_level() -> None:
    blocks = _by_index(inspect(_pkg("complex_numbering")))
    assert [(block["kind"], block["list_level"]) for block in blocks.values()] == [
        ("heading", None),
        ("list_item", 0),
        ("list_item", 1),
        ("list_item", 2),
        ("list_item", 0),
        ("list_item", 0),
        ("list_item", 0),
    ]


def test_long_text_is_truncated_and_says_so() -> None:
    pkg = _pkg("simple")
    paragraph = indexed_paragraphs(pkg.document)[1]
    run = etree.SubElement(paragraph, qn("w:r"))
    etree.SubElement(run, qn("w:t")).text = "x" * 200

    block = _by_index(inspect(pkg))[1]
    assert len(block["text"]) == TEXT_PREVIEW
    assert block["truncated"] is True
    assert _by_index(inspect(pkg))[2]["truncated"] is False


def test_a_block_reports_the_fields_it_holds() -> None:
    blocks = _by_index(inspect(_pkg("fields")))
    assert [block["has_fields"] for block in blocks.values()] == [
        False,
        True,
        True,
        True,
        False,
        True,
    ]


def test_a_block_reports_the_comments_anchored_in_it() -> None:
    blocks = _by_index(inspect(_pkg("comments")))
    assert [block["has_comments"] for block in blocks.values()] == [
        False,
        True,
        True,
        True,
    ]


def test_a_block_reports_the_revisions_it_holds() -> None:
    pkg = _pkg("tracked_changes")
    blocks = _by_index(inspect(pkg))
    revised = {
        revision.paragraph_index
        for revision in list_revisions(pkg)
        if revision.story == "document"
    }
    assert {index for index, block in blocks.items() if block["has_revisions"]} == revised
    assert blocks[0]["has_revisions"] is False


# --------------------------------------------------------------------------------------
# Tables
# --------------------------------------------------------------------------------------


def test_tables_are_described_the_way_the_table_locator_addresses_them() -> None:
    report = inspect(_pkg("tables"))
    assert report["tables"] == [
        {
            "index": 0,
            "rows": 3,
            # The widest row: the first one merges two grid columns into a
            # single cell, so it holds two cells where the others hold three.
            "columns": 3,
            "first_cell": "Merged header across two columns",
            "nested": False,
        },
        {
            "index": 1,
            "rows": 1,
            "columns": 2,
            "first_cell": "Nested A",
            "nested": True,
        },
    ]


def test_every_table_index_is_a_usable_locator() -> None:
    pkg = _pkg("tables")
    for table in inspect(pkg)["tables"]:
        target = resolve(pkg, {"table": table["index"], "row": 0, "col": 0})
        assert target.text.startswith(table["first_cell"][:20])
        assert target.index is None


# --------------------------------------------------------------------------------------
# Sections
# --------------------------------------------------------------------------------------


def test_sections_report_page_setup_in_twips() -> None:
    sections = inspect(_pkg("sections"))["sections"]
    assert len(sections) == 2
    first, second = sections
    assert first["orientation"] == "portrait"
    assert first["format"] == {"width": 12240, "height": 15840, "name": "Letter"}
    assert first["margins"]["left"] == 1800
    assert second["orientation"] == "landscape"
    assert (second["format"]["width"], second["format"]["height"]) == (15840, 12240)
    # A landscape Letter page is still Letter: the name does not depend on which
    # side is the long one.
    assert second["format"]["name"] == "Letter"
    assert second["margins"]["left"] == 720


def test_sections_name_the_header_and_footer_stories() -> None:
    report = inspect(_pkg("headers_footers"))
    [section] = report["sections"]
    assert section["headers"] == {"default": "header1", "first": "header2"}
    assert section["footers"] == {"default": "footer1"}
    # Those ids are stories, so they can be inspected and addressed as such.
    assert {story["id"] for story in report["stories"]} >= {
        "header1",
        "header2",
        "footer1",
    }


def test_an_unknown_page_size_is_reported_without_a_name() -> None:
    pkg = _pkg("simple")
    size = next(pkg.document.iter(qn("w:pgSz")))
    size.set(qn("w:w"), "9000")
    size.set(qn("w:h"), "9000")
    [section] = inspect(pkg)["sections"]
    assert section["format"] == {"width": 9000, "height": 9000, "name": None}
    assert section["orientation"] == "portrait"


# --------------------------------------------------------------------------------------
# Stories and counts
# --------------------------------------------------------------------------------------


def test_stories_report_the_size_of_their_own_index_space() -> None:
    pkg = _pkg("headers_footers")
    roots = dict(pkg.stories())
    for story in inspect(pkg)["stories"]:
        assert story["paragraphs"] == len(indexed_paragraphs(roots[story["id"]]))
    assert inspect(pkg)["stories"][0]["id"] == "document"


def test_counts_agree_with_what_the_report_lists() -> None:
    report = inspect(_pkg("combined"))
    counts = report["counts"]
    assert counts["paragraphs"] == len(report["blocks"])
    assert counts["tables"] == len(report["tables"])
    assert counts["sections"] == len(report["sections"])
    assert counts["bookmarks"] == len(report["bookmarks"])
    assert counts["fields"] == len(report["fields"])
    assert counts["stories"] == len(report["stories"])


def test_counts_agree_with_the_engine_itself() -> None:
    pkg = _pkg("combined")
    counts = inspect(pkg)["counts"]
    assert counts["revisions"] == len(list_revisions(pkg))
    assert counts["comments"] == 3
    assert counts["words"] > 0
    assert counts["characters"] > counts["words"]


def test_a_document_without_comments_counts_none() -> None:
    assert inspect(_pkg("simple"))["counts"]["comments"] == 0


# --------------------------------------------------------------------------------------
# Styles, bookmarks, fields
# --------------------------------------------------------------------------------------


def test_styles_used_counts_references_not_definitions() -> None:
    used = inspect(_pkg("paragraph_styles"))["styles_used"]
    assert used["paragraph"]["FixtureBody"] == 2
    assert used["paragraph"]["Quote"] == 1
    # Defined by the template and by the fixture, referenced by neither.
    assert "Normal" not in used["paragraph"]


def test_styles_used_separates_the_three_kinds() -> None:
    used = inspect(_pkg("combined"))["styles_used"]
    assert set(used) == {"paragraph", "character", "table"}
    assert "Heading1" in used["paragraph"]
    assert "Hyperlink" in used["character"]
    assert used["table"] == {"TableGrid": 2}


def test_bookmarks_are_listed_with_their_span() -> None:
    bookmarks = inspect(_pkg("bookmarks"))["bookmarks"]
    assert [entry["name"] for entry in bookmarks] == [
        "FixtureInline",
        "FixtureSpanning",
        "FixtureEmpty",
        "_Hidden_Word_Bookmark",
    ]
    inline = bookmarks[0]
    assert (inline["story"], inline["paragraph"]) == ("document", 1)
    assert (inline["start"], inline["end"]) == (8, 23)
    assert resolve(_pkg("bookmarks"), {"bookmark": "FixtureInline"}).start == 8


def test_a_bookmark_in_a_table_cell_reports_no_paragraph_index() -> None:
    pkg = _pkg("tables")
    cell_paragraph = resolve(pkg, {"table": 0, "row": 1, "col": 1}).paragraph
    start = etree.SubElement(cell_paragraph, qn("w:bookmarkStart"))
    start.set(W_ID, "8100")
    start.set(W_NAME, "InACell")

    [entry] = inspect(pkg)["bookmarks"]
    assert entry["name"] == "InACell"
    assert entry["paragraph"] is None
    assert entry["story"] == "document"


def test_fields_are_listed_with_their_instruction_and_span() -> None:
    fields = inspect(_pkg("fields"))["fields"]
    assert [entry["instr"] for entry in fields] == [
        "PAGE   \\* MERGEFORMAT",
        'TOC \\o "1-3" \\h \\z \\u',
        'IF  = 1 "first page" "later page"',
        "PAGE",
        "REF FixtureRefTarget \\h",
    ]
    page = fields[0]
    assert (page["story"], page["paragraph"]) == ("document", 1)
    assert page["start"] < page["end"]


def test_fields_of_a_footer_report_their_own_story() -> None:
    fields = inspect(_pkg("headers_footers"))["fields"]
    assert [(entry["story"], entry["instr"]) for entry in fields] == [
        ("footer1", "PAGE"),
        ("footer1", "NUMPAGES"),
    ]


@pytest.mark.parametrize("name", sorted({"simple", "combined", "tables", "text_boxes"}))
def test_inspection_runs_on_every_shape_of_document(name: str) -> None:
    pkg = _pkg(name)
    report = inspect(pkg)
    assert report["counts"]["paragraphs"] == len(report["blocks"])
    # Whatever the document holds, a search result and the report agree on the
    # index space they number in.
    for match in find(pkg, "Fixture"):
        if match.index is not None:
            assert report["blocks"][match.index]["index"] == match.index
