"""Tests for :mod:`word_document_server.engine.locators`.

The cases below are organised by locator form, then by failure code.  Two
properties are checked throughout rather than in one place: the V2 index a
target reports is the index the ``paragraph`` form takes back (round trip), and
that index space is the one -- and only one -- the rest of the codebase numbers
in.
"""

from __future__ import annotations

import copy

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.snapshot import diff, snapshot
from word_document_server.engine.errors import LocatorError
from word_document_server.engine.find import Match, find
from word_document_server.engine.locators import Target, indexed_paragraphs, resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import qn

W_NAME = qn("w:name")
W_ID = qn("w:id")
W_BOOKMARK_START = qn("w:bookmarkStart")


def _pkg(name: str) -> DocxPackage:
    return DocxPackage.open(build(name))


def _texts(pkg: DocxPackage, story: str = "document") -> list[str]:
    root = dict(pkg.stories())[story]
    return [visible_text(paragraph) for paragraph in indexed_paragraphs(root)]


def _code(excinfo: pytest.ExceptionInfo[LocatorError]) -> str:
    return excinfo.value.code


# --------------------------------------------------------------------------------------
# Target
# --------------------------------------------------------------------------------------


def test_target_has_the_same_shape_as_a_match() -> None:
    # A caller must be able to treat a search result and a resolved locator
    # alike; the two dataclasses are only useful if they stay aligned.
    assert list(Target.__dataclass_fields__) == [
        "story",
        "paragraph",
        "index",
        "start",
        "end",
    ]
    assert set(Target.__dataclass_fields__) <= set(Match.__dataclass_fields__)


def test_target_text_is_read_from_the_live_paragraph() -> None:
    pkg = _pkg("simple")
    target = resolve(pkg, {"find": "simple fixture"})
    assert target.text == "simple fixture"
    assert visible_text(target.paragraph)[target.start : target.end] == target.text


def test_resolving_never_touches_the_package() -> None:
    source = build("combined")
    pkg = DocxPackage.open(source)
    for locator in (
        {"paragraph": 3},
        {"find": "Fixture: tables"},
        {"bookmark": "FixtureInline"},
        {"heading": "Fixture: bookmarks"},
        {"table": 0, "row": 1, "col": 1},
    ):
        resolve(pkg, locator)
    assert diff(snapshot(source), snapshot(pkg.to_bytes())).is_empty()


# --------------------------------------------------------------------------------------
# The paragraph form, and the index space it addresses
# --------------------------------------------------------------------------------------


def test_paragraph_locator_addresses_the_v2_index() -> None:
    pkg = _pkg("simple")
    target = resolve(pkg, {"paragraph": 2})
    assert target.story == "document"
    assert target.index == 2
    assert target.text == "Second paragraph, with a trailing sentence."
    assert (target.start, target.end) == (0, len(target.text))


def test_every_paragraph_index_round_trips() -> None:
    pkg = _pkg("combined")
    for index, text in enumerate(_texts(pkg)):
        target = resolve(pkg, {"paragraph": index})
        assert target.index == index
        assert target.text == text


def test_paragraph_index_skips_table_cells() -> None:
    # The "tables" fixture holds a heading, a table of eight cell paragraphs,
    # then one body paragraph: the cells must consume no index at all.
    pkg = _pkg("tables")
    assert _texts(pkg) == ["Fixture: tables", "Paragraph after the table."]
    assert resolve(pkg, {"paragraph": 1}).text == "Paragraph after the table."


def test_paragraph_index_skips_text_boxes() -> None:
    # Five body paragraphs, three text-box paragraphs (VML, DrawingML and the
    # DrawingML fallback) that must not shift the ones that follow.
    pkg = _pkg("text_boxes")
    assert _texts(pkg) == [
        "Fixture: text boxes",
        "Body text before the boxes.",
        "Anchor of the VML box.",
        "Anchor of the DrawingML box.",
        "Body text after the boxes.",
    ]
    assert resolve(pkg, {"paragraph": 4}).text == "Body text after the boxes."


def test_paragraph_index_counts_block_content_control_content() -> None:
    # The opposite rule: a block w:sdt is part of the flow, so its paragraph
    # takes an index and everything after it is shifted by one.
    pkg = _pkg("content_controls")
    assert resolve(pkg, {"paragraph": 1}).text == "Paragraph inside a block content control."
    assert resolve(pkg, {"paragraph": 3}).text == "Alpha"


def test_the_index_space_is_the_one_the_tool_layer_numbers() -> None:
    # D-016: one filter, two wrappers (engine, utils) -- never a third.  A
    # locator index and a paragraph_index of the tool layer must be the same
    # number, on a document holding every construct that could split them.
    from word_document_server.utils.document_utils import (
        indexed_paragraphs as tool_indexed_paragraphs,
    )

    pkg = _pkg("combined")
    root = pkg.document
    assert indexed_paragraphs(root) == tool_indexed_paragraphs(root)


def test_the_index_space_is_the_one_find_reports() -> None:
    pkg = _pkg("combined")
    matches = find(pkg, "Fixture: ")
    assert matches and all(match.index is not None for match in matches)
    for match in matches:
        assert resolve(pkg, {"paragraph": match.index}).paragraph is match.paragraph


def test_paragraph_index_past_the_end_is_not_found() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"paragraph": 5})
    assert _code(excinfo) == "not_found"
    assert "5 indexed paragraph(s)" in str(excinfo.value)


# --------------------------------------------------------------------------------------
# Stories
# --------------------------------------------------------------------------------------


def test_a_locator_can_address_another_story() -> None:
    pkg = _pkg("headers_footers")
    target = resolve(pkg, {"paragraph": 0, "story": "header1"})
    assert target.story == "header1"
    assert target.text.startswith("Fixture header")


def test_body_is_an_accepted_alias_that_is_never_reported() -> None:
    pkg = _pkg("simple")
    target = resolve(pkg, {"paragraph": 0, "story": "body"})
    assert target.story == "document"


def test_an_absent_story_is_not_found_and_the_error_lists_the_real_ones() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"paragraph": 0, "story": "header7"})
    assert _code(excinfo) == "not_found"
    assert "'document'" in str(excinfo.value)


# --------------------------------------------------------------------------------------
# expect_text: the stale anchor
# --------------------------------------------------------------------------------------


def test_expect_text_accepts_the_exact_text_and_a_prefix() -> None:
    pkg = _pkg("simple")
    exact = resolve(pkg, {"paragraph": 1, "expect_text": "First paragraph of the simple fixture."})
    prefix = resolve(pkg, {"paragraph": 1, "expect_text": "First paragraph"})
    assert exact.index == prefix.index == 1


def test_a_stale_anchor_names_the_paragraphs_that_do_match() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"paragraph": 1, "expect_text": "Second paragraph"})
    assert _code(excinfo) == "stale_anchor"
    message = str(excinfo.value)
    assert "found 'First paragraph of the simple fixture.'" in message
    assert "[2]" in message


def test_a_stale_anchor_says_so_when_nothing_matches() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"paragraph": 1, "expect_text": "a sentence this document never had"})
    assert _code(excinfo) == "stale_anchor"
    assert "no paragraph of that story carries that text" in str(excinfo.value)


def test_an_index_shifted_by_an_edit_is_caught_by_its_anchor() -> None:
    # The failure the anchor exists for: a paragraph inserted above index 2
    # silently turns it into another paragraph.
    pkg = _pkg("simple")
    anchor = {"paragraph": 2, "expect_text": "Second paragraph"}
    assert resolve(pkg, anchor).index == 2

    paragraphs = indexed_paragraphs(pkg.document)
    paragraphs[0].addprevious(paragraphs[0].makeelement(qn("w:p"), {}))
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, anchor)
    assert _code(excinfo) == "stale_anchor"
    assert "[3]" in str(excinfo.value)


def test_expect_text_is_checked_on_every_form() -> None:
    pkg = _pkg("tables")
    assert resolve(pkg, {"table": 0, "row": 0, "col": 1, "expect_text": "Third column"})
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"table": 0, "row": 0, "col": 1, "expect_text": "Fourth column"})
    assert _code(excinfo) == "stale_anchor"
    assert "the located paragraph" in str(excinfo.value)


def test_a_stale_anchor_suggests_candidates_after_a_within_search_too() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(
            pkg,
            {
                "find": "paragraph",
                "within": {"paragraph": 1},
                "expect_text": "Second paragraph",
            },
        )
    assert _code(excinfo) == "stale_anchor"
    assert "[2]" in str(excinfo.value)


# --------------------------------------------------------------------------------------
# The find form
# --------------------------------------------------------------------------------------


def test_find_locator_returns_the_span_of_the_match() -> None:
    pkg = _pkg("simple")
    target = resolve(pkg, {"find": "simple fixture"})
    assert (target.index, target.start, target.end) == (1, 23, 37)
    assert target.text == "simple fixture"


def test_find_locator_without_occurrence_refuses_to_guess() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"find": "paragraph"})
    assert _code(excinfo) == "ambiguous"
    assert "occurs 2 times" in str(excinfo.value)
    assert "'occurrence'" in str(excinfo.value)


def test_find_locator_counts_occurrences_from_one() -> None:
    pkg = _pkg("simple")
    first = resolve(pkg, {"find": "paragraph", "occurrence": 1})
    second = resolve(pkg, {"find": "paragraph", "occurrence": 2})
    assert (first.index, second.index) == (1, 2)


def test_an_occurrence_past_the_last_one_is_not_found() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"find": "paragraph", "occurrence": 3})
    assert _code(excinfo) == "not_found"
    assert "occurs 2 time(s)" in str(excinfo.value)


def test_text_that_does_not_occur_is_not_found() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"find": "no such text"})
    assert _code(excinfo) == "not_found"


def test_within_restricts_the_search_to_one_span() -> None:
    pkg = _pkg("simple")
    inside = resolve(pkg, {"find": "paragraph", "within": {"paragraph": 2}})
    assert inside.index == 2
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"find": "simple fixture", "within": {"paragraph": 2}})
    assert _code(excinfo) == "not_found"
    assert "the 'within' span" in str(excinfo.value)


def test_within_can_narrow_a_search_down_to_a_bookmark() -> None:
    pkg = _pkg("bookmarks")
    target = resolve(pkg, {"find": "span", "within": {"bookmark": "FixtureInline"}})
    assert target.index == 1
    # "After." sits in the same paragraph but outside the bookmark's span.
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"find": "After", "within": {"bookmark": "FixtureInline"}})
    assert _code(excinfo) == "not_found"


def test_within_and_an_explicit_story_must_agree() -> None:
    pkg = _pkg("headers_footers")
    with pytest.raises(LocatorError) as excinfo:
        resolve(
            pkg,
            {"find": "Fixture", "story": "header1", "within": {"paragraph": 0}},
        )
    assert _code(excinfo) == "invalid"
    assert "contradicts" in str(excinfo.value)


def test_find_is_the_only_form_that_reaches_a_text_box() -> None:
    pkg = _pkg("text_boxes")
    target = resolve(pkg, {"find": "Text inside the VML text box."})
    assert target.index is None
    assert target.story == "document"
    assert target.text == "Text inside the VML text box."


def test_find_searches_the_requested_story() -> None:
    pkg = _pkg("headers_footers")
    target = resolve(pkg, {"find": "Page", "story": "footer1"})
    assert target.story == "footer1"
    assert target.index == 0


# --------------------------------------------------------------------------------------
# The bookmark form
# --------------------------------------------------------------------------------------


def test_bookmark_locator_spans_from_start_to_end() -> None:
    pkg = _pkg("bookmarks")
    target = resolve(pkg, {"bookmark": "FixtureInline"})
    assert target.index == 1
    assert target.text == "Bookmarked span"


def test_a_bookmark_spanning_two_paragraphs_stops_at_the_first_one() -> None:
    pkg = _pkg("bookmarks")
    target = resolve(pkg, {"bookmark": "FixtureSpanning"})
    assert target.index == 2
    assert target.text == "A bookmark opens here"
    assert target.end == len(visible_text(target.paragraph))


def test_an_empty_bookmark_resolves_to_an_empty_span() -> None:
    pkg = _pkg("bookmarks")
    target = resolve(pkg, {"bookmark": "FixtureEmpty"})
    assert target.start == target.end
    assert target.text == ""


def test_an_unknown_bookmark_is_not_found_and_the_error_lists_the_real_ones() -> None:
    pkg = _pkg("bookmarks")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"bookmark": "Nowhere"})
    assert _code(excinfo) == "not_found"
    assert "FixtureInline" in str(excinfo.value)


def test_a_duplicated_bookmark_name_is_ambiguous() -> None:
    # Word does not write two bookmarks under one name, but a document that has
    # been through a merge or a hand-written tool can carry them.
    pkg = _pkg("bookmarks")
    paragraph = indexed_paragraphs(pkg.document)[0]
    clone = etree.SubElement(paragraph, W_BOOKMARK_START)
    clone.set(W_ID, "8001")
    clone.set(W_NAME, "FixtureInline")

    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"bookmark": "FixtureInline"})
    assert _code(excinfo) == "ambiguous"
    assert "2 bookmarks" in str(excinfo.value)


# --------------------------------------------------------------------------------------
# The heading form
# --------------------------------------------------------------------------------------


def test_heading_locator_matches_the_whole_heading() -> None:
    pkg = _pkg("simple")
    target = resolve(pkg, {"heading": "Second heading"})
    assert target.index == 3
    assert target.text == "Second heading"


def test_heading_locator_falls_back_to_a_prefix() -> None:
    pkg = _pkg("simple")
    assert resolve(pkg, {"heading": "Second"}).index == 3


def test_an_exact_heading_wins_over_a_longer_one_it_prefixes() -> None:
    # A document holding both "Second" and "Second heading" must not be
    # ambiguous for the caller who typed the short one exactly.
    pkg = _pkg("simple")
    paragraphs = indexed_paragraphs(pkg.document)
    short = copy.deepcopy(paragraphs[3])
    for text in short.iter(qn("w:t")):
        text.text = "Second"
    paragraphs[3].addnext(short)

    target = resolve(pkg, {"heading": "Second"})
    assert target.text == "Second"
    assert target.index == 4


def test_an_ambiguous_heading_names_the_candidates() -> None:
    pkg = _pkg("combined")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"heading": "Fixture"})
    assert _code(excinfo) == "ambiguous"
    assert "address one by its index" in str(excinfo.value)


def test_an_unknown_heading_is_not_found_and_the_error_lists_real_headings() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"heading": "Conclusion"})
    assert _code(excinfo) == "not_found"
    assert "Fixture: simple" in str(excinfo.value)


def test_a_body_paragraph_is_not_a_heading() -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"heading": "First paragraph of the simple fixture."})
    assert _code(excinfo) == "not_found"


# --------------------------------------------------------------------------------------
# The table form
# --------------------------------------------------------------------------------------


def test_table_locator_reaches_a_cell_paragraph() -> None:
    pkg = _pkg("tables")
    target = resolve(pkg, {"table": 0, "row": 1, "col": 1})
    assert target.text == "Row two, column two"
    # A cell paragraph is out of the V2 space by construction.
    assert target.index is None


def test_table_locator_defaults_to_the_first_paragraph_of_the_cell() -> None:
    pkg = _pkg("tables")
    explicit = resolve(pkg, {"table": 0, "row": 0, "col": 0, "paragraph": 0})
    implicit = resolve(pkg, {"table": 0, "row": 0, "col": 0})
    assert implicit.paragraph is explicit.paragraph
    assert implicit.text == "Merged header across two columns"


def test_a_merged_cell_occupies_one_column_position() -> None:
    # The first row holds two cells: one spanning two grid columns, then the
    # third column.  "col" counts cells, so the second one is col 1, not col 2.
    pkg = _pkg("tables")
    assert resolve(pkg, {"table": 0, "row": 0, "col": 1}).text == "Third column"
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"table": 0, "row": 0, "col": 2})
    assert _code(excinfo) == "not_found"
    assert "2 cell(s)" in str(excinfo.value)


def test_a_nested_table_is_numbered_after_its_host() -> None:
    pkg = _pkg("tables")
    target = resolve(pkg, {"table": 1, "row": 0, "col": 1})
    assert target.text == "Nested B"
    assert target.index is None


@pytest.mark.parametrize(
    ("locator", "expected"),
    [
        ({"table": 2, "row": 0, "col": 0}, "2 table(s)"),
        ({"table": 0, "row": 3, "col": 0}, "3 row(s)"),
        ({"table": 0, "row": 1, "col": 3}, "3 cell(s)"),
        ({"table": 0, "row": 1, "col": 1, "paragraph": 1}, "1 paragraph(s)"),
    ],
)
def test_every_table_coordinate_is_bounds_checked(
    locator: dict[str, int], expected: str
) -> None:
    pkg = _pkg("tables")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, locator)
    assert _code(excinfo) == "not_found"
    assert expected in str(excinfo.value)


def test_the_table_form_needs_its_three_coordinates() -> None:
    pkg = _pkg("tables")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"table": 0, "row": 0})
    assert _code(excinfo) == "invalid"
    assert "'col'" in str(excinfo.value)


# --------------------------------------------------------------------------------------
# Malformed locators
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "locator",
    [
        {},
        {"story": "document"},
        {"paragraph": 1, "find": "x"},
        {"paragraph": 1, "colour": "red"},
        {"paragraph": -1},
        {"paragraph": "1"},
        {"paragraph": True},
        {"paragraph": 1.0},
        {"find": ""},
        {"find": 3},
        {"find": "x", "occurrence": 0},
        {"bookmark": ""},
        {"heading": None},
        {"paragraph": 0, "story": 3},
        {"paragraph": 0, "expect_text": 7},
        [],
        "paragraph 1",
        None,
    ],
)
def test_a_malformed_locator_is_invalid(locator: object) -> None:
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, locator)  # type: ignore[arg-type]
    assert _code(excinfo) == "invalid"


def test_an_invalid_locator_looks_nothing_up() -> None:
    # The distinction matters to a caller: "invalid" means fix the locator,
    # "not_found" means the document does not hold that place.
    pkg = _pkg("simple")
    with pytest.raises(LocatorError) as excinfo:
        resolve(pkg, {"table": 0, "row": 0, "col": 0, "story": "nowhere", "extra": 1})
    assert _code(excinfo) == "invalid"
    assert "extra" in str(excinfo.value)
