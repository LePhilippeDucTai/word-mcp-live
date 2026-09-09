"""Tests for tests/support/package_check.py.

Two kinds of coverage:
- structural checks exercised directly on hand-built (and hand-corrupted)
  .docx packages, which never require LibreOffice;
- a sanity check that LibreOffice-produced fixtures are themselves clean,
  guarded by `requires_libreoffice` since LibreOffice stays optional.
"""

import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from tests.support.libreoffice import fodt_to_docx, requires_libreoffice
from tests.support.package_check import CT_NS, REL_NS, W_NS, validate_package

_FIXTURE_NAMES = ["rich", "tables_lists", "notes_fields"]


def _make_base_docx(tmp_path: Path, name: str = "base.docx") -> Path:
    doc = Document()
    doc.add_paragraph("Hello world.")
    doc.add_paragraph("Second paragraph.")
    path = tmp_path / name
    doc.save(path)
    return path


def _read_zip(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _write_zip(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_bookmark(parts: dict[str, bytes], bookmark_id: str, name: str) -> None:
    """Insert a bookmarkStart/bookmarkEnd pair into the first paragraph."""
    root = etree.fromstring(parts["word/document.xml"])
    first_p = root.find(f"{{{W_NS}}}body/{{{W_NS}}}p")
    start = etree.SubElement(first_p, f"{{{W_NS}}}bookmarkStart")
    start.set(f"{{{W_NS}}}id", bookmark_id)
    start.set(f"{{{W_NS}}}name", name)
    end = etree.SubElement(first_p, f"{{{W_NS}}}bookmarkEnd")
    end.set(f"{{{W_NS}}}id", bookmark_id)
    parts["word/document.xml"] = _serialize(root)


def _add_comment(parts: dict[str, bytes], comment_id: str) -> None:
    """Add a commentReference in the body plus a matching word/comments.xml part."""
    root = etree.fromstring(parts["word/document.xml"])
    body = root.find(f"{{{W_NS}}}body")
    p = etree.SubElement(body, f"{{{W_NS}}}p")
    r = etree.SubElement(p, f"{{{W_NS}}}r")
    ref = etree.SubElement(r, f"{{{W_NS}}}commentReference")
    ref.set(f"{{{W_NS}}}id", comment_id)
    sect_pr = body.find(f"{{{W_NS}}}sectPr")
    if sect_pr is not None:
        body.remove(p)
        sect_pr.addprevious(p)
    parts["word/document.xml"] = _serialize(root)

    comments_root = etree.Element(f"{{{W_NS}}}comments", nsmap={"w": W_NS})
    comment_el = etree.SubElement(comments_root, f"{{{W_NS}}}comment")
    comment_el.set(f"{{{W_NS}}}id", comment_id)
    comment_el.set(f"{{{W_NS}}}author", "Tester")
    cp = etree.SubElement(comment_el, f"{{{W_NS}}}p")
    cr = etree.SubElement(cp, f"{{{W_NS}}}r")
    ct_el = etree.SubElement(cr, f"{{{W_NS}}}t")
    ct_el.text = "A comment."
    parts["word/comments.xml"] = _serialize(comments_root)

    ct_root = etree.fromstring(parts["[Content_Types].xml"])
    override = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
    override.set("PartName", "/word/comments.xml")
    override.set(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    )
    parts["[Content_Types].xml"] = _serialize(ct_root)

    rels_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
    rel = etree.SubElement(rels_root, f"{{{REL_NS}}}Relationship")
    rel.set("Id", "rIdComments")
    rel.set(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    )
    rel.set("Target", "comments.xml")
    parts["word/_rels/document.xml.rels"] = _serialize(rels_root)


def test_validate_package_clean_document_has_no_issues(tmp_path):
    path = _make_base_docx(tmp_path)
    assert validate_package(path) == []


def test_validate_package_bookmark_and_comment_are_coherent(tmp_path):
    path = _make_base_docx(tmp_path)
    parts = _read_zip(path)
    _add_bookmark(parts, bookmark_id="0", name="MyBookmark")
    _add_comment(parts, comment_id="0")
    _write_zip(path, parts)

    assert validate_package(path) == []


def test_validate_package_detects_missing_relationship_target(tmp_path):
    path = _make_base_docx(tmp_path)
    parts = _read_zip(path)
    rels_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
    styles_rel = next(
        rel
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship")
        if rel.get("Target") == "styles.xml"
    )
    styles_rel.set("Target", "styles-does-not-exist.xml")
    parts["word/_rels/document.xml.rels"] = _serialize(rels_root)
    _write_zip(path, parts)

    issues = validate_package(path)
    assert any(i.code == "REL-TARGET-MISSING" for i in issues), issues


def test_validate_package_detects_duplicate_id(tmp_path):
    path = _make_base_docx(tmp_path)
    parts = _read_zip(path)
    _add_bookmark(parts, bookmark_id="0", name="BM1")
    _add_bookmark(parts, bookmark_id="0", name="BM2")
    _write_zip(path, parts)

    issues = validate_package(path)
    assert any(i.code == "ID-DUPLICATE" for i in issues), issues


def test_validate_package_detects_part_without_content_type(tmp_path):
    path = _make_base_docx(tmp_path)
    parts = _read_zip(path)
    parts["word/extra.bin"] = b"\x00\x01\x02"
    _write_zip(path, parts)

    issues = validate_package(path)
    assert any(i.code == "CT-MISSING" and i.part == "word/extra.bin" for i in issues), issues


def test_validate_package_missing_file_reports_issue(tmp_path):
    issues = validate_package(tmp_path / "does-not-exist.docx")
    assert len(issues) == 1
    assert issues[0].code == "ZIP-UNREADABLE"


def test_validate_package_flags_stray_temp_file(tmp_path):
    path = _make_base_docx(tmp_path)
    (tmp_path / "~$base.docx").write_bytes(b"lock file")

    issues = validate_package(path)
    assert any(i.code == "TEMP-FILE-PRESENT" for i in issues), issues


@requires_libreoffice
@pytest.mark.libreoffice
@pytest.mark.parametrize("name", _FIXTURE_NAMES)
def test_validate_package_libreoffice_fixtures_are_clean(name, tmp_path_factory):
    cache_dir = tmp_path_factory.mktemp("odt_docx_cache_pkgcheck")
    docx_path = fodt_to_docx(name, cache_dir)

    issues = validate_package(docx_path)

    assert issues == [], issues
