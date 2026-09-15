"""Tests for :mod:`word_document_server.engine.find`."""

from __future__ import annotations

import pytest
from lxml import etree

from tests.fixtures.builders import build
from word_document_server.engine.find import Match, find, iter_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import qn

W_P = qn("w:p")


def _pkg(name: str) -> DocxPackage:
    return DocxPackage.open(build(name))


# --------------------------------------------------------------------------------------
# iter_paragraphs
# --------------------------------------------------------------------------------------


def test_iter_paragraphs_covers_cells_and_sdt_in_document_order() -> None:
    pkg = _pkg("tables")
    root = dict(pkg.stories())["document"]
    paragraphs = iter_paragraphs(root)

    # Same order and count as a plain lxml walk, cells and nested tables included.
    assert paragraphs == list(root.iter(W_P))
    texts = [visible_text(p) for p in paragraphs]
    assert texts == [
        "Fixture: tables",
        "Merged header across two columns",
        "Third column",
        "Vertically merged cell",
        "Row two, column two",
        "Nested A",
        "Nested B",
        "",
        "",
        "Row three, column two",
        "Row three, column three",
        "Paragraph after the table.",
    ]


def test_iter_paragraphs_includes_block_sdt_content() -> None:
    pkg = _pkg("content_controls")
    root = dict(pkg.stories())["document"]
    texts = [visible_text(p) for p in iter_paragraphs(root)]
    assert "Paragraph inside a block content control." in texts


def test_iter_paragraphs_returns_a_list_not_a_one_shot_iterator() -> None:
    pkg = _pkg("simple")
    root = dict(pkg.stories())["document"]
    paragraphs = iter_paragraphs(root)
    assert len(paragraphs) == len(list(paragraphs))


# --------------------------------------------------------------------------------------
# Basic matching
# --------------------------------------------------------------------------------------


def test_find_literal_match_in_body() -> None:
    pkg = _pkg("simple")
    matches = find(pkg, "simple fixture")
    assert len(matches) == 1
    match = matches[0]
    assert isinstance(match, Match)
    assert match.story == "document"
    assert match.paragraph.tag == W_P
    assert match.index == 1  # second body paragraph, 0-based
    assert (match.start, match.end) == (23, 37)
    assert match.text == "simple fixture"
    assert visible_text(match.paragraph)[match.start : match.end] == match.text


def test_find_context_is_a_window_around_the_match() -> None:
    pkg = _pkg("simple")
    [match] = find(pkg, "simple fixture")
    text = visible_text(match.paragraph)
    assert match.context == text
    assert match.text in match.context


def test_find_matches_are_returned_in_document_order() -> None:
    pkg = _pkg("simple")
    matches = find(pkg, "paragraph")
    starts = [(m.index, m.start) for m in matches]
    assert starts == sorted(starts)
    assert len(matches) >= 2


def test_find_match_straddling_two_runs() -> None:
    # "One senten" | "ce cut in thr" | "ee runs." -- the pattern below spans the
    # boundary between the first and second run.
    pkg = _pkg("mixed_runs")
    matches = find(pkg, "sentence cut in three")
    assert len(matches) == 1
    assert matches[0].text == "sentence cut in three"


def test_find_no_match_across_paragraph_boundary() -> None:
    # "Fixture: simple" is immediately followed, in the next paragraph, by
    # "First paragraph...". Concatenating them must not be found: each
    # paragraph is searched on its own.
    pkg = _pkg("simple")
    assert find(pkg, "simpleFirst") == []


# --------------------------------------------------------------------------------------
# Visibility: tracked changes, hyperlinks, insertions
# --------------------------------------------------------------------------------------


def test_find_ignores_text_under_deletion() -> None:
    pkg = _pkg("tracked_changes")
    assert find(pkg, "deleted text") == []
    # The surrounding, non-deleted text is still found.
    assert len(find(pkg, "Kept text")) == 1


def test_find_finds_inserted_text() -> None:
    pkg = _pkg("tracked_changes")
    matches = find(pkg, "inserted text")
    assert len(matches) == 1
    assert matches[0].text == "inserted text"


def test_find_finds_text_inside_a_hyperlink() -> None:
    pkg = _pkg("hyperlinks")
    matches = find(pkg, "example.org")
    assert len(matches) == 1
    assert matches[0].story == "document"


# --------------------------------------------------------------------------------------
# whole_word, regex, case
# --------------------------------------------------------------------------------------


def test_find_whole_word_excludes_partial_matches() -> None:
    pkg = _pkg("simple")
    # "graph" is a substring of "paragraph" but never a whole word on its own.
    assert find(pkg, "graph") != []
    assert find(pkg, "graph", whole_word=True) == []
    assert find(pkg, "paragraph", whole_word=True) != []


def test_find_regex_mode() -> None:
    pkg = _pkg("simple")
    matches = find(pkg, r"para\w+", regex=True)
    assert matches
    assert all(m.text.startswith("para") for m in matches)


def test_find_regex_mode_respects_whole_word() -> None:
    pkg = _pkg("simple")
    assert find(pkg, r"gra\w+", regex=True, whole_word=True) == []


def test_find_case_sensitivity() -> None:
    pkg = _pkg("simple")
    assert find(pkg, "FIXTURE") == []
    assert find(pkg, "FIXTURE", case=False) != []


# --------------------------------------------------------------------------------------
# stories
# --------------------------------------------------------------------------------------


def test_find_default_stories_is_body_only() -> None:
    pkg = _pkg("headers_footers")
    # "Fixture header " only lives in header1; the default must not see it.
    assert find(pkg, "Fixture header") == []
    matches = find(pkg, "Fixture header", stories=("header1",))
    assert len(matches) == 1
    assert matches[0].story == "header1"


def test_find_stories_accepts_multiple_explicit_ids() -> None:
    pkg = _pkg("headers_footers")
    matches = find(pkg, "Fixture", stories=("header1", "header2", "footer1"))
    assert {m.story for m in matches} == {"header1", "header2"}


def test_find_unknown_story_is_silently_skipped() -> None:
    pkg = _pkg("simple")
    assert find(pkg, "simple", stories=("footnotes",)) == []


# --------------------------------------------------------------------------------------
# V2 paragraph index: table cells excluded, sdt included
# --------------------------------------------------------------------------------------


def test_find_index_excludes_table_cell_paragraphs() -> None:
    pkg = _pkg("tables")
    [match] = find(pkg, "Row three, column two")
    assert match.index is None


def test_find_index_counts_only_top_level_paragraphs_around_a_table() -> None:
    pkg = _pkg("tables")
    [heading] = find(pkg, "Fixture: tables")
    [after] = find(pkg, "Paragraph after the table.")
    assert heading.index == 0
    assert after.index == 1


def test_find_index_includes_block_sdt_content() -> None:
    pkg = _pkg("content_controls")
    [match] = find(pkg, "Paragraph inside a block content control.")
    assert match.index is not None


# --------------------------------------------------------------------------------------
# V2 paragraph index: text boxes excluded
# --------------------------------------------------------------------------------------

#: The three paragraphs the ``text_boxes`` fixture hides in a shape: the VML
#: box, and the two branches -- modern and fallback -- of the DrawingML one,
#: which describe the same box twice.
TEXT_BOX_TEXTS = (
    "Text inside the VML text box.",
    "Text inside the DrawingML text box.",
    "Text inside the DrawingML fallback.",
)

#: Every paragraph of the ``text_boxes`` body, in document order, with the V2
#: index it must report.  The boxes sit between anchors 2 and 3.
TEXT_BOX_BODY = (
    "Fixture: text boxes",
    "Body text before the boxes.",
    "Anchor of the VML box.",
    "Anchor of the DrawingML box.",
    "Body text after the boxes.",
)


def test_iter_paragraphs_still_lists_text_box_paragraphs() -> None:
    pkg = _pkg("text_boxes")
    root = dict(pkg.stories())["document"]
    paragraphs = iter_paragraphs(root)
    assert paragraphs == list(root.iter(W_P))
    texts = [visible_text(p) for p in paragraphs]
    assert [text for text in texts if text in TEXT_BOX_TEXTS] == list(TEXT_BOX_TEXTS)


def test_find_gives_no_index_to_a_paragraph_inside_a_text_box() -> None:
    pkg = _pkg("text_boxes")
    for text in TEXT_BOX_TEXTS:
        [match] = find(pkg, text)
        assert match.index is None, text


def test_find_numbers_the_paragraphs_around_a_text_box_without_a_gap() -> None:
    pkg = _pkg("text_boxes")
    found = []
    for text in TEXT_BOX_BODY:
        [match] = find(pkg, text)
        found.append(match.index)
    assert found == [0, 1, 2, 3, 4]


# --------------------------------------------------------------------------------------
# max_results
# --------------------------------------------------------------------------------------


def test_find_max_results_truncates_deterministically() -> None:
    pkg = _pkg("simple")
    full = find(pkg, "paragraph")
    limited = find(pkg, "paragraph", max_results=1)
    assert limited == full[:1]


def test_find_max_results_zero_returns_nothing() -> None:
    pkg = _pkg("simple")
    assert find(pkg, "paragraph", max_results=0) == []


# --------------------------------------------------------------------------------------
# Errors
# --------------------------------------------------------------------------------------


def test_find_empty_pattern_raises() -> None:
    pkg = _pkg("simple")
    with pytest.raises(ValueError):
        find(pkg, "")


def test_find_regex_with_empty_match_raises() -> None:
    pkg = _pkg("simple")
    with pytest.raises(ValueError):
        find(pkg, r"x?", regex=True)


def test_find_invalid_regex_raises_value_error() -> None:
    pkg = _pkg("simple")
    with pytest.raises(ValueError):
        find(pkg, "(unclosed", regex=True)


# --------------------------------------------------------------------------------------
# Read-only
# --------------------------------------------------------------------------------------


def test_find_does_not_mutate_the_tree() -> None:
    pkg = _pkg("combined")
    root = dict(pkg.stories())["document"]
    before = etree.tostring(root)
    find(pkg, "Fixture", stories=("body", "header1", "footer1", "footnotes", "endnotes"))
    after = etree.tostring(root)
    assert before == after
