"""Tests for ``format_text`` and for the cell formatting reached through
``format_table_cell_text``.

Both tools used to work by *rebuilding* what they touched: ``format_text``
cleared every run of the paragraph and re-created three runs from the text it
had read back, and ``format_cell_text`` assigned ``cell.text``, which empties
the cell of everything it holds.  What that lost is not visible in the text --
it is the run properties of the runs nobody asked about, the hyperlink the runs
lived in, the table nested in the cell -- so the assertions here are written on
the XML the tools leave behind, and on the package-scale diff, not on the
paragraph text alone.

Paragraph indices are stated in the space they belong to.  ``format_text``
takes a **V2** index (every paragraph of the body except those in a table cell
or a text box, see ``engine/find.py``), while ``assert_unchanged_except`` takes
the wider **snapshot** index space (``story-root//w:p``).  The two coincide only
in a document without tables; :func:`snapshot_index` resolves the second one by
text so the difference is never guessed.
"""

from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

from tests.support.package_check import validate_package
from tests.support.snapshot import Snapshot, assert_unchanged_except, snapshot
from word_document_server.engine.find import _v2_index_map, iter_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import qn
from word_document_server.tools.format_tools import format_table_cell_text, format_text

W_R = qn("w:r")
W_RPR = qn("w:rPr")
W_T = qn("w:t")
W_TBL = qn("w:tbl")
W_GRID_COL = qn("w:gridCol")
W_HYPERLINK = qn("w:hyperlink")
W_VAL = qn("w:val")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def run(coro: Any) -> Any:
    """Drive an ``async`` tool call to completion."""
    return asyncio.run(coro)


def story_paragraphs(path: Path) -> list[etree._Element]:
    """Every ``w:p`` of the main story of the document at `path`."""
    _, story = DocxPackage.open(path).stories()[0]
    return iter_paragraphs(story)


def v2_paragraphs(path: Path) -> dict[int, etree._Element]:
    """The main story's paragraphs, keyed by their V2 index."""
    paragraphs = story_paragraphs(path)
    return {index: paragraph for paragraph, index in _v2_index_map(paragraphs).items()}


def snapshot_index(snap: Snapshot, text: str, story: str = "document") -> int:
    """Snapshot-space index of the one paragraph whose text is exactly `text`."""
    matches = [i for i, sig in enumerate(snap.story_paragraphs(story)) if sig.text == text]
    assert len(matches) == 1, (
        f"expected exactly one {story!r} paragraph with text {text!r} "
        f"(snapshot space), found at indices {matches}"
    )
    return matches[0]


def runs_of(paragraph: etree._Element) -> list[etree._Element]:
    """Every ``w:r`` of the paragraph, in document order, containers included."""
    return list(paragraph.iter(W_R))


def rpr_names(run_element: etree._Element) -> list[str]:
    """Local names of the run's ``w:rPr`` children, in order."""
    properties = run_element.find(W_RPR)
    return [] if properties is None else [etree.QName(c).localname for c in properties]


def run_text(run_element: etree._Element) -> str:
    """Concatenated ``w:t`` text of one run."""
    return "".join(node.text or "" for node in run_element.iter(W_T))


def formatting_of(paragraph: etree._Element) -> list[tuple[str, tuple[str, ...]]]:
    """``(text, rPr children)`` for every run of the paragraph that shows text."""
    return [
        (run_text(element), tuple(rpr_names(element)))
        for element in runs_of(paragraph)
        if run_text(element)
    ]


def grid_of(table: etree._Element) -> list[str | None]:
    """The ``w:w`` of every ``w:gridCol`` of `table`.

    ``TableSignature`` only counts the columns, so a rewritten grid width is
    invisible to ``assert_unchanged_except`` (decision D-004): a test that
    touches a table checks the grid itself.
    """
    return [column.get(qn("w:w")) for column in table.iter(W_GRID_COL)]


def tables_of(path: Path) -> list[etree._Element]:
    """Every ``w:tbl`` of the main story, nested ones included, in document order."""
    _, story = DocxPackage.open(path).stories()[0]
    return list(story.iter(W_TBL))


# --------------------------------------------------------------------------
# What the range covers, and only that
# --------------------------------------------------------------------------


def test_untouched_runs_keep_their_own_formatting(fixture_docx):
    """The regression the rebuild caused: formatting outside the range is kept."""
    path = fixture_docx("mixed_runs")

    result = run(format_text(str(path), 1, 0, 5, bold=True))
    assert "formatted successfully" in result

    paragraph = v2_paragraphs(path)[1]
    assert visible_text(paragraph) == "Plain then bold bold italic et la fin."
    assert formatting_of(paragraph) == [
        ("Plain", ("b",)),
        (" then ", ()),
        ("bold ", ("b",)),
        ("bold italic ", ("b", "i")),
        ("et la fin.", ("lang",)),
    ]


def test_a_character_style_survives_direct_formatting(fixture_docx):
    """``w:rStyle`` is not something a run can be rebuilt from its text with."""
    path = fixture_docx("character_styles")

    result = run(format_text(str(path), 1, 12, 22, bold=True))
    assert "formatted successfully" in result

    paragraph = v2_paragraphs(path)[1]
    emphasised = [element for element in runs_of(paragraph) if run_text(element) == "emphasised"]
    assert len(emphasised) == 1
    # rStyle before b: w:rPr is a sequence, not a bag.
    assert rpr_names(emphasised[0]) == ["rStyle", "b"]
    assert emphasised[0].find(W_RPR).find(qn("w:rStyle")).get(W_VAL) == "FixtureEmphasis"


def test_a_hyperlink_is_not_duplicated_into_the_paragraph(fixture_docx):
    """The proof case: runs inside ``w:hyperlink`` are not the paragraph's own.

    Rebuilding the paragraph from its full text re-emitted the hyperlink's text
    as a plain run while the hyperlink kept its own, so the reader saw it twice.
    """
    path = fixture_docx("hyperlinks")
    before = visible_text(v2_paragraphs(path)[1])
    assert before == "An external link: example.org."

    result = run(format_text(str(path), 1, 3, 11, bold=True))
    assert "formatted successfully" in result

    paragraph = v2_paragraphs(path)[1]
    assert visible_text(paragraph) == before
    links = paragraph.findall(W_HYPERLINK)
    assert len(links) == 1
    assert visible_text(links[0].getparent()).count("example.org") == 1
    assert formatting_of(paragraph) == [
        ("An ", ()),
        ("external", ("b",)),
        (" link: ", ()),
        ("example.org", ("rStyle",)),
        (".", ()),
    ]


def test_formatting_inside_a_hyperlink_keeps_the_run_in_it(fixture_docx):
    """A run never leaves its ``w:hyperlink``."""
    path = fixture_docx("hyperlinks")

    result = run(format_text(str(path), 1, 18, 29, bold=True))
    assert "formatted successfully" in result

    paragraph = v2_paragraphs(path)[1]
    assert visible_text(paragraph) == "An external link: example.org."
    link = paragraph.find(W_HYPERLINK)
    assert [run_text(element) for element in link.iter(W_R)] == ["example.org"]
    assert rpr_names(link.find(W_R)) == ["rStyle", "b"]
    assert link.get(qn("r:id")) is not None


def test_the_package_is_unchanged_outside_the_paragraph(fixture_docx):
    path = fixture_docx("mixed_runs")
    before = snapshot(path.read_bytes())
    index = snapshot_index(before, "Plain then bold bold italic et la fin.")

    assert "formatted successfully" in run(
        format_text(str(path), 1, 0, 5, bold=True, italic=True, color="red",
                    font_size=14, font_name="Arial")
    )

    after = snapshot(path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[index])
    assert validate_package(path) == []


def test_a_second_call_only_adds_what_it_names(fixture_docx):
    """The patch is a set of changes: what a previous call set is still there."""
    path = fixture_docx("mixed_runs")

    run(format_text(str(path), 1, 0, 5, bold=True))
    run(format_text(str(path), 1, 0, 5, italic=True))

    paragraph = v2_paragraphs(path)[1]
    assert formatting_of(paragraph)[0] == ("Plain", ("b", "i"))


# --------------------------------------------------------------------------
# The index space
# --------------------------------------------------------------------------


def test_table_cell_paragraphs_take_no_index(fixture_docx):
    """Ten of the twelve paragraphs of the fixture live in a table cell."""
    path = fixture_docx("tables")
    assert len(story_paragraphs(path)) == 12
    assert sorted(v2_paragraphs(path)) == [0, 1]

    result = run(format_text(str(path), 1, 0, 9, bold=True))
    assert "Text 'Paragraph' formatted successfully in paragraph 1." == result

    paragraph = v2_paragraphs(path)[1]
    assert formatting_of(paragraph) == [("Paragraph", ("b",)), (" after the table.", ())]


def test_an_index_past_the_last_paragraph_is_refused(fixture_docx):
    path = fixture_docx("tables")
    before = path.read_bytes()

    result = run(format_text(str(path), 2, 0, 4, bold=True))
    assert result == "Invalid paragraph index. Document has 2 paragraphs (0-1)."
    assert path.read_bytes() == before


def test_positions_outside_the_visible_text_are_refused(fixture_docx):
    path = fixture_docx("mixed_runs")
    before = path.read_bytes()

    result = run(format_text(str(path), 1, 0, 500, bold=True))
    assert result == "Invalid text positions. Paragraph has 38 characters."
    assert path.read_bytes() == before


# --------------------------------------------------------------------------
# Colours
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("given", "expected"),
    [("red", "FF0000"), ("FF0000", "FF0000"), ("#00ff00", "00FF00"), ("auto", "auto")],
)
def test_a_colour_is_written_as_the_schema_stores_it(fixture_docx, given, expected):
    path = fixture_docx("mixed_runs")

    assert "formatted successfully" in run(format_text(str(path), 1, 0, 5, color=given))

    paragraph = v2_paragraphs(path)[1]
    coloured = runs_of(paragraph)[0]
    assert run_text(coloured) == "Plain"
    assert coloured.find(W_RPR).find(qn("w:color")).get(W_VAL) == expected


def test_an_unreadable_colour_is_refused_rather_than_written_black(fixture_docx):
    """Falling back to black wrote a colour the caller never asked for."""
    path = fixture_docx("mixed_runs")
    before = path.read_bytes()

    result = run(format_text(str(path), 1, 0, 5, color="chartreuse"))
    assert result.startswith("Invalid color 'chartreuse'.")
    assert "FF0000" in result or "red" in result
    assert path.read_bytes() == before


# --------------------------------------------------------------------------
# Table cells
# --------------------------------------------------------------------------


def test_new_cell_text_keeps_the_table_nested_in_the_cell(fixture_docx):
    """``cell.text = ...`` used to empty the cell of its nested table."""
    path = fixture_docx("tables")
    before = snapshot(path.read_bytes())
    # The cell's own paragraph sits after the table nested in it, and is empty:
    # it cannot be named by its text, so it is named by its neighbour.
    index = snapshot_index(before, "Nested B") + 1
    grid_before = [grid_of(table) for table in tables_of(path)]

    result = run(format_table_cell_text(str(path), 0, 1, 2, text_content="Replaced"))
    assert "formatted successfully" in result

    texts = [visible_text(paragraph) for paragraph in story_paragraphs(path)]
    assert texts.count("Nested A") == 1
    assert texts.count("Nested B") == 1
    assert "Replaced" in texts
    assert [grid_of(table) for table in tables_of(path)] == grid_before

    after = snapshot(path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[index])
    assert validate_package(path) == []


def test_cell_formatting_keeps_what_the_runs_already_carried(fixture_docx):
    path = fixture_docx("tables")

    result = run(format_table_cell_text(str(path), 0, 0, 0, italic=True))
    assert "formatted successfully" in result

    header = story_paragraphs(path)[1]
    assert visible_text(header) == "Merged header across two columns"
    assert formatting_of(header) == [("Merged header across two columns", ("b", "i"))]


def test_cell_formatting_leaves_the_nested_table_alone(fixture_docx):
    """``cell.paragraphs`` is the cell's own paragraphs, not the nested table's."""
    path = fixture_docx("tables")
    before = snapshot(path.read_bytes())

    result = run(format_table_cell_text(str(path), 0, 1, 2, bold=True))
    assert "formatted successfully" in result

    nested = [p for p in story_paragraphs(path) if visible_text(p) in ("Nested A", "Nested B")]
    assert [formatting_of(paragraph) for paragraph in nested] == [
        [("Nested A", ())],
        [("Nested B", ())],
    ]
    # The cell's own paragraph is empty, so there is no run to patch at all.
    after = snapshot(path.read_bytes())
    assert_unchanged_except(before, after)


def test_an_unreadable_cell_colour_is_refused_rather_than_written_black(fixture_docx):
    path = fixture_docx("tables")
    before = path.read_bytes()

    result = run(format_table_cell_text(str(path), 0, 0, 0, color="chartreuse"))
    assert result.startswith("Failed to format cell text: Invalid color 'chartreuse'.")
    assert path.read_bytes() == before
