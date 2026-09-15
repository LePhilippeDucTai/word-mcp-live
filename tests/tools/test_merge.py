"""Tests for ``merge_documents``, now backed by the OOXML engine.

The tool used to rebuild its target in a blank ``Document()`` from
``paragraph.text``; it now opens the first source and appends the body of every
other one through
:func:`word_document_server.engine.merge.append_document`. What is pinned here
is what a caller sees: the return message keeps its shape (``"Successfully
merged N documents into X"``, ``"Cannot merge documents. ..."``, ``"Failed to
merge documents: ..."``), the parameters keep their names and their meaning, and
what used to be silently destroyed is now either carried over or reported.

The one deliberate behaviour change is the base: the merged document *is* the
first source, so it keeps its own styles, sections, headers and footers instead
of the empty defaults of a blank document.
"""

from __future__ import annotations

import asyncio
import zipfile
from pathlib import Path
from typing import Any

import pytest

from tests.support.package_check import validate_package
from tests.support.snapshot import diff, snapshot
from word_document_server.core import tables as tables_module
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import qn
from word_document_server.tools.document_tools import merge_documents


def _run(coro: Any) -> Any:
    """Drive an ``async`` tool call to completion."""
    return asyncio.run(coro)


def _texts(path: Path) -> list[str]:
    """Visible text of every paragraph of the body, in document order."""
    pkg = DocxPackage.open(str(path))
    return [visible_text(paragraph) for paragraph in pkg.document.iter(qn("w:p"))]


def _members(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as archive:
        return set(archive.namelist())


# --------------------------------------------------------------------------------------
# The message a caller gets
# --------------------------------------------------------------------------------------


def test_merging_reports_the_number_of_documents_and_the_target(tmp_path, fixture_docx) -> None:
    first = fixture_docx("simple")
    second = fixture_docx("tables")
    target = tmp_path / "merged.docx"

    result = _run(merge_documents(str(target), [str(first), str(second)]))

    assert result == f"Successfully merged 2 documents into {target}"


def test_the_docx_extension_is_still_added_for_the_caller(tmp_path, fixture_docx) -> None:
    source = fixture_docx("simple")
    target = tmp_path / "merged"

    result = _run(merge_documents(str(target), [str(source)]))

    assert result == f"Successfully merged 1 documents into {target}.docx"
    assert (tmp_path / "merged.docx").exists()


def test_a_missing_source_is_reported_without_writing_anything(tmp_path, fixture_docx) -> None:
    source = fixture_docx("simple")
    target = tmp_path / "merged.docx"
    absent = tmp_path / "nowhere.docx"

    result = _run(merge_documents(str(target), [str(source), str(absent)]))

    assert result == (
        f"Cannot merge documents. The following source files do not exist: {absent}"
    )
    assert not target.exists()


def test_an_unusable_target_directory_is_reported(tmp_path, fixture_docx) -> None:
    source = fixture_docx("simple")
    target = tmp_path / "nowhere" / "merged.docx"

    result = _run(merge_documents(str(target), [str(source)]))

    assert result.startswith("Cannot create target document: Directory ")


# --------------------------------------------------------------------------------------
# What survives the merge
# --------------------------------------------------------------------------------------


def test_a_single_source_is_reproduced_exactly(tmp_path, fixture_docx) -> None:
    source = fixture_docx("combined")
    target = tmp_path / "merged.docx"

    _run(merge_documents(str(target), [str(source)], add_page_breaks=False))

    before, after = snapshot(source.read_bytes()), snapshot(target.read_bytes())
    # Diff.is_empty is a method: without the call this assertion would be vacuous.
    assert diff(before, after).is_empty(), diff(before, after).describe()
    assert "word/comments.xml" in after.parts


def test_the_first_source_keeps_its_headers_footers_and_sections(tmp_path, fixture_docx) -> None:
    first = fixture_docx("headers_footers")
    second = fixture_docx("simple")
    target = tmp_path / "merged.docx"

    _run(merge_documents(str(target), [str(first), str(second)]))

    members = _members(target)
    assert {"word/header1.xml", "word/header2.xml", "word/footer1.xml"} <= members
    assert "word/media/image1.png" in members  # the picture inside the header
    assert snapshot(target.read_bytes()).sect_pr["document"] == (
        snapshot(first.read_bytes()).sect_pr["document"]
    )


def test_the_appended_bodies_keep_their_runs_and_direct_formatting(tmp_path, fixture_docx) -> None:
    first = fixture_docx("simple")
    second = fixture_docx("mixed_runs")
    target = tmp_path / "merged.docx"

    _run(merge_documents(str(target), [str(first), str(second)], add_page_breaks=False))

    merged = snapshot(target.read_bytes()).story_paragraphs("document")
    expected = snapshot(second.read_bytes()).story_paragraphs("document")
    assert [signature.runs for signature in merged][-len(expected) :] == [
        signature.runs for signature in expected
    ]


def test_images_hyperlinks_and_bookmarks_are_carried_over(tmp_path, fixture_docx) -> None:
    first = fixture_docx("simple")
    target = tmp_path / "merged.docx"
    sources = [str(first), str(fixture_docx("drawings")), str(fixture_docx("hyperlinks"))]

    _run(merge_documents(str(target), sources))

    assert "word/media/image1.png" in _members(target)
    merged = DocxPackage.open(str(target))
    external = [
        relationship.target_ref
        for relationship in merged.document_part.rels.values()
        if relationship.is_external
    ]
    assert "https://example.org/fixture" in external
    names = [
        element.get(qn("w:name")) for element in merged.document.iter(qn("w:bookmarkStart"))
    ]
    assert "FixtureAnchor" in names


def test_tables_keep_their_geometry_and_every_cell_paragraph(tmp_path, fixture_docx) -> None:
    first = fixture_docx("simple")
    second = fixture_docx("tables")
    target = tmp_path / "merged.docx"

    _run(merge_documents(str(target), [str(first), str(second)]))

    merged = snapshot(target.read_bytes())
    expected = snapshot(second.read_bytes())
    assert [
        (signature.rows, signature.grid_cols, signature.grid, signature.cells)
        for _, signature in sorted(merged.tables.items())
    ] == [
        (signature.rows, signature.grid_cols, signature.grid, signature.cells)
        for _, signature in sorted(expected.tables.items())
    ]


def test_page_breaks_are_added_between_documents_only_when_asked(tmp_path, fixture_docx) -> None:
    sources = [str(fixture_docx("simple")), str(fixture_docx("tables"))]

    with_breaks = tmp_path / "with.docx"
    _run(merge_documents(str(with_breaks), sources, add_page_breaks=True))
    without = tmp_path / "without.docx"
    _run(merge_documents(str(without), sources, add_page_breaks=False))

    def breaks(path: Path) -> list[str]:
        pkg = DocxPackage.open(str(path))
        return [
            element.get(qn("w:type"))
            for element in pkg.document.iter(qn("w:br"))
            if element.get(qn("w:type")) == "page"
        ]

    assert breaks(with_breaks) == ["page"]
    assert breaks(without) == []


def test_the_merged_document_is_a_valid_package(tmp_path, fixture_docx) -> None:
    sources = [
        str(fixture_docx("headers_footers")),
        str(fixture_docx("tables")),
        str(fixture_docx("drawings")),
        str(fixture_docx("complex_numbering")),
        str(fixture_docx("style_inheritance")),
    ]
    target = tmp_path / "merged.docx"

    _run(merge_documents(str(target), sources))

    assert validate_package(target) == []
    texts = _texts(target)
    for marker in (
        "Fixture: headers and footers",
        "Fixture: tables",
        "Fixture: drawings",
        "Fixture: complex numbering",
        "Fixture: style inheritance",
    ):
        assert marker in texts


# --------------------------------------------------------------------------------------
# What is refused, and what is reported
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("fixture", ["comments", "footnotes"])
def test_a_source_carrying_comments_or_notes_is_refused(
    tmp_path, fixture_docx, fixture: str
) -> None:
    first = fixture_docx("simple")
    second = fixture_docx(fixture)
    target = tmp_path / "merged.docx"

    result = _run(merge_documents(str(target), [str(first), str(second)]))

    assert result.startswith("Failed to merge documents: ")
    assert fixture in result
    # Nothing is written when a source is refused: the merge happens in memory.
    assert not target.exists()


def test_a_refused_source_leaves_an_existing_target_alone(tmp_path, fixture_docx) -> None:
    first = fixture_docx("simple")
    second = fixture_docx("comments")
    target = tmp_path / "merged.docx"
    target.write_bytes(fixture_docx("tables").read_bytes())
    before = target.read_bytes()

    result = _run(merge_documents(str(target), [str(first), str(second)]))

    assert result.startswith("Failed to merge documents: ")
    assert target.read_bytes() == before


def test_what_could_not_be_carried_over_is_appended_to_the_message(
    tmp_path, fixture_docx
) -> None:
    first = fixture_docx("simple")
    second = fixture_docx("headers_footers")
    target = tmp_path / "merged.docx"

    result = _run(merge_documents(str(target), [str(first), str(second)]))

    head, _, warnings = result.partition("\n")
    assert head == f"Successfully merged 2 documents into {target}"
    assert warnings.startswith("Warnings:\n- headers_footers.docx: ")
    assert "headers and footers" in warnings


def test_no_warning_section_when_nothing_was_lost(tmp_path, fixture_docx) -> None:
    sources = [str(fixture_docx("simple")), str(fixture_docx("tables"))]
    target = tmp_path / "merged.docx"

    result = _run(merge_documents(str(target), sources))

    assert "\n" not in result


# --------------------------------------------------------------------------------------
# The helper the old implementation leaned on
# --------------------------------------------------------------------------------------


def test_copy_table_is_gone() -> None:
    # It rebuilt a table from cell text alone -- one paragraph per cell, no runs,
    # no merges. merge_documents was its last caller.
    assert not hasattr(tables_module, "copy_table")
