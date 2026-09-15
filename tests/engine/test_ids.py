"""Tests for :mod:`word_document_server.engine.ids`.

Expected values are recomputed from the raw package with ``zipfile`` and lxml
rather than with the engine, so that a bug in the engine's traversal cannot make
the assertion agree with it.
"""

from __future__ import annotations

import re
import zipfile
from collections.abc import Iterable
from pathlib import Path

import pytest
from lxml import etree

from word_document_server.engine.errors import IdExhausted
from word_document_server.engine.ids import (
    ANNOTATION_ID_TAGS,
    MAX_DECIMAL_ID,
    PARA_ID_MAX,
    PARA_ID_MIN,
    next_annotation_id,
    next_comment_id,
    next_footnote_id,
    next_para_id,
)
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.xmlns import W14, W, qn

PARA_ID_RE = re.compile(r"^[0-9A-F]{8}$")


def _word_roots(path: Path) -> Iterable[etree._Element]:
    """Every ``word/*.xml`` part of the package, parsed independently."""
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("word/") and name.endswith(".xml") and "/_rels/" not in name:
                yield etree.fromstring(archive.read(name))


def _max_w_id(path: Path, tags: Iterable[str]) -> int | None:
    """Largest ``w:id`` carried by `tags` anywhere under ``word/``.

    `tags` may be given prefixed (``"w:ins"``) or bare (``"ins"``).
    """
    wanted = tuple(f"{{{W}}}{tag.rpartition(':')[2]}" for tag in tags)
    values = [
        int(element.get(f"{{{W}}}id"))
        for root in _word_roots(path)
        for element in root.iter(*wanted)
        if element.get(f"{{{W}}}id") is not None
    ]
    return max(values) if values else None


def _used_para_ids(path: Path) -> set[int]:
    attribute = f"{{{W14}}}paraId"
    return {
        int(raw, 16)
        for root in _word_roots(path)
        for element in root.iter("*")
        if (raw := element.get(attribute)) is not None
    }


def _first_paragraph(root: etree._Element) -> etree._Element:
    return next(root.iter(qn("w:p")))


def _add_bookmark(paragraph: etree._Element, id_value: int, name: str) -> None:
    start = etree.SubElement(paragraph, qn("w:bookmarkStart"))
    start.set(qn("w:id"), str(id_value))
    start.set(qn("w:name"), name)
    end = etree.SubElement(paragraph, qn("w:bookmarkEnd"))
    end.set(qn("w:id"), str(id_value))


# ----------------------------------------------------------------------------
# next_annotation_id
# ----------------------------------------------------------------------------


def test_next_annotation_id_starts_at_one(fixture_docx) -> None:
    # 0 is the id Word gives the implicit _GoBack bookmark, so it is never handed out.
    assert next_annotation_id(DocxPackage.open(fixture_docx("simple"))) == 1


@pytest.mark.parametrize(
    "name", ["tracked_changes", "bookmarks", "hyperlinks", "fields", "comments", "combined"]
)
def test_next_annotation_id_clears_every_annotation_in_the_package(name: str, fixture_docx) -> None:
    path = fixture_docx(name)
    expected = _max_w_id(path, ANNOTATION_ID_TAGS)
    assert expected is not None, f"{name} carries no annotation id to clear"
    assert next_annotation_id(DocxPackage.open(path)) == expected + 1


def test_next_annotation_id_is_pure(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("tracked_changes"))
    assert next_annotation_id(pkg) == next_annotation_id(pkg)


def test_next_annotation_id_advances_once_the_id_is_used(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("bookmarks"))
    first = next_annotation_id(pkg)
    _add_bookmark(_first_paragraph(pkg.document), first, "Allocated")
    assert next_annotation_id(pkg) == first + 1


def test_next_annotation_id_scans_headers_and_footers(fixture_docx) -> None:
    # An id allocated from the body alone would collide with a header bookmark.
    pkg = DocxPackage.open(fixture_docx("headers_footers"))
    header = dict(pkg.stories())["header1"]
    _add_bookmark(_first_paragraph(header), 500, "InHeader")
    assert next_annotation_id(pkg) == 501


def test_next_annotation_id_scans_the_comments_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("comments"))
    comments = pkg.root_of(pkg.part("/word/comments.xml"))
    _add_bookmark(_first_paragraph(comments), 700, "InComment")
    assert next_annotation_id(pkg) == 701


def test_next_annotation_id_ignores_a_non_numeric_id(fixture_docx) -> None:
    # A malformed id written by another producer must not stop the engine from
    # allocating; it is skipped, not parsed.
    pkg = DocxPackage.open(fixture_docx("simple"))
    paragraph = _first_paragraph(pkg.document)
    _add_bookmark(paragraph, 12, "Numeric")
    _add_bookmark(paragraph, 0, "Malformed")
    paragraph.findall(qn("w:bookmarkStart"))[-1].set(qn("w:id"), "not-a-number")
    assert next_annotation_id(pkg) == 13


def test_next_annotation_id_raises_when_the_space_is_full(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    _add_bookmark(_first_paragraph(pkg.document), MAX_DECIMAL_ID, "Last")
    with pytest.raises(IdExhausted):
        next_annotation_id(pkg)


# ----------------------------------------------------------------------------
# next_comment_id
# ----------------------------------------------------------------------------


def test_next_comment_id_starts_at_one_without_a_comments_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    assert pkg.find_part("/word/comments.xml") is None
    assert next_comment_id(pkg) == 1


@pytest.mark.parametrize("name", ["comments", "combined"])
def test_next_comment_id_clears_every_comment(name: str, fixture_docx) -> None:
    path = fixture_docx(name)
    expected = _max_w_id(path, ("comment", "commentRangeStart", "commentReference"))
    assert expected is not None
    assert next_comment_id(DocxPackage.open(path)) == expected + 1


def test_next_comment_id_clears_an_orphaned_anchor(fixture_docx) -> None:
    # The comment is gone from comments.xml but the body still references it:
    # reusing its id would attach the new comment to the stale anchor.
    pkg = DocxPackage.open(fixture_docx("comments"))
    comments_root = pkg.root_of(pkg.part("/word/comments.xml"))
    for comment in list(comments_root.iter(qn("w:comment"))):
        comments_root.remove(comment)
    assert next_comment_id(pkg) == _max_w_id(fixture_docx("comments"), ("commentReference",)) + 1


# ----------------------------------------------------------------------------
# next_footnote_id
# ----------------------------------------------------------------------------


def test_next_footnote_id_starts_at_one_without_a_footnotes_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    assert next_footnote_id(pkg) == 1


@pytest.mark.parametrize("name", ["footnotes", "combined"])
def test_next_footnote_id_clears_every_footnote(name: str, fixture_docx) -> None:
    path = fixture_docx(name)
    expected = _max_w_id(path, ("footnote", "footnoteReference"))
    assert expected is not None
    result = next_footnote_id(DocxPackage.open(path))
    assert result == expected + 1
    assert result >= 1


def test_next_footnote_id_ignores_the_reserved_separator_ids(fixture_docx) -> None:
    # Word numbers the separators -1 and 0; a document holding only those must
    # still allocate 1, never 0.
    pkg = DocxPackage.open(fixture_docx("footnotes"))
    footnotes = dict(pkg.stories())["footnotes"]
    for footnote in list(footnotes.iter(qn("w:footnote"))):
        if footnote.get(qn("w:type")) is None:
            footnotes.remove(footnote)
    assert {f.get(qn("w:id")) for f in footnotes.iter(qn("w:footnote"))} <= {"-1", "0"}
    for reference in list(pkg.document.iter(qn("w:footnoteReference"))):
        reference.getparent().remove(reference)
    assert next_footnote_id(pkg) == 1


# ----------------------------------------------------------------------------
# next_para_id
# ----------------------------------------------------------------------------


@pytest.mark.parametrize("name", ["simple", "comments", "combined", "tracked_changes"])
def test_next_para_id_is_a_valid_unused_value(name: str, fixture_docx) -> None:
    path = fixture_docx(name)
    result = next_para_id(DocxPackage.open(path))

    assert PARA_ID_RE.match(result), result
    value = int(result, 16)
    assert PARA_ID_MIN <= value <= PARA_ID_MAX
    assert value not in _used_para_ids(path)


def test_next_para_id_is_pure(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("comments"))
    assert next_para_id(pkg) == next_para_id(pkg)


def test_next_para_id_advances_once_the_id_is_used(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("comments"))
    first = next_para_id(pkg)
    _first_paragraph(pkg.document).set(qn("w14:paraId"), first)
    second = next_para_id(pkg)
    assert second != first
    assert int(second, 16) not in {int(first, 16)}


def test_next_para_id_scans_the_comments_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("comments"))
    comments = pkg.root_of(pkg.part("/word/comments.xml"))
    _first_paragraph(comments).set(qn("w14:paraId"), "7FFFFFF0")
    assert next_para_id(pkg) == "7FFFFFF1"


def test_next_para_id_wraps_instead_of_overflowing(fixture_docx) -> None:
    # 7FFFFFFF is the last value Word accepts; the next one restarts at 1 rather
    # than producing 80000000, which Word rejects.
    pkg = DocxPackage.open(fixture_docx("simple"))
    _first_paragraph(pkg.document).set(qn("w14:paraId"), f"{PARA_ID_MAX:08X}")
    assert next_para_id(pkg) == f"{PARA_ID_MIN:08X}"


def test_next_para_id_skips_a_used_value_after_wrapping(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    paragraphs = list(pkg.document.iter(qn("w:p")))
    paragraphs[0].set(qn("w14:paraId"), f"{PARA_ID_MAX:08X}")
    paragraphs[1].set(qn("w14:paraId"), "00000001")
    assert next_para_id(pkg) == "00000002"


def test_next_para_id_ignores_a_non_hexadecimal_value(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    _first_paragraph(pkg.document).set(qn("w14:paraId"), "not-hex")
    assert next_para_id(pkg) == f"{PARA_ID_MIN:08X}"


# ----------------------------------------------------------------------------
# Allocated ids survive a save
# ----------------------------------------------------------------------------


def test_allocated_ids_survive_a_save(fixture_docx, tmp_path: Path) -> None:
    pkg = DocxPackage.open(fixture_docx("combined"))
    annotation = next_annotation_id(pkg)
    para = next_para_id(pkg)
    paragraph = _first_paragraph(pkg.document)
    paragraph.set(qn("w14:paraId"), para)
    _add_bookmark(paragraph, annotation, "Allocated")

    target = tmp_path / "out.docx"
    pkg.save(target)

    reopened = DocxPackage.open(target)
    assert next_annotation_id(reopened) == annotation + 1
    assert int(para, 16) in _used_para_ids(target)
    assert next_para_id(reopened) != para
