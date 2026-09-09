"""Tests for :mod:`word_document_server.engine.package`.

The headline test is :func:`test_open_then_save_is_lossless`: every fixture is
opened and saved without a single edit, and the canonical snapshot of the result
must be identical to the snapshot of the source, with no package validation
issue. That is the floor the whole engine builds on -- an operation cannot be
non-degrading if merely opening the document already degrades it.
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part, XmlPart

from tests.fixtures.builders import (
    ALL_FIXTURES,
    CT_COMMENTS_EXTENDED,
    RT_COMMENTS_EXTENDED,
    build,
)
from tests.support.package_check import validate_package
from tests.support.snapshot import diff, snapshot
from word_document_server.engine import package as package_module
from word_document_server.engine.errors import PackageError
from word_document_server.engine.package import (
    MAIN_STORY,
    DocxPackage,
    atomic_write_bytes,
    story_name,
)
from word_document_server.engine.xmlns import qn

FIXTURE_NAMES = sorted(ALL_FIXTURES)

COMMENTS_EXTENDED_PARTNAME = "/word/commentsExtended.xml"
COMMENTS_EXTENDED_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)


@pytest.fixture
def dest(tmp_path: Path) -> Path:
    """An empty directory, so "no temporary file was left" is a crisp assertion."""
    directory = tmp_path / "dest"
    directory.mkdir()
    return directory


# ----------------------------------------------------------------------------
# Opening
# ----------------------------------------------------------------------------


def test_open_from_path_exposes_the_document_root(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    assert pkg.document.tag == qn("w:document")
    assert pkg.document.find(qn("w:body")) is not None


def test_open_from_bytes_exposes_the_same_document(fixture_docx) -> None:
    path = fixture_docx("simple")
    from_path = DocxPackage.open(path)
    from_bytes = DocxPackage.open(path.read_bytes())
    assert from_bytes.document.find(qn("w:body")) is not None
    difference = diff(snapshot(from_path.to_bytes()), snapshot(from_bytes.to_bytes()))
    assert difference.is_empty, difference.describe()


def test_open_from_path_records_the_source_and_bytes_do_not(fixture_docx) -> None:
    path = fixture_docx("simple")
    assert DocxPackage.open(path).source_path == path
    assert DocxPackage.open(path.read_bytes()).source_path is None


def test_open_rejects_a_file_that_is_not_a_package(tmp_path: Path) -> None:
    junk = tmp_path / "not-a-docx.docx"
    junk.write_bytes(b"PK\x03\x04 truncated nonsense")
    with pytest.raises(PackageError):
        DocxPackage.open(junk)


def test_open_rejects_a_zip_without_a_main_document_part(tmp_path: Path) -> None:
    # A valid OPC package for another application: readable zip, no w:document.
    empty = tmp_path / "empty.docx"
    empty.write_bytes(b"")
    with pytest.raises(PackageError):
        DocxPackage.open(empty)


def test_open_missing_file_raises_the_plain_os_error(tmp_path: Path) -> None:
    # OS failures stay OS failures: PackageError is about package semantics.
    with pytest.raises(FileNotFoundError):
        DocxPackage.open(tmp_path / "absent.docx")


# ----------------------------------------------------------------------------
# Round trip
# ----------------------------------------------------------------------------


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_open_then_save_is_lossless(name: str, fixture_docx, dest: Path) -> None:
    source = fixture_docx(name)
    target = dest / f"{name}.docx"

    DocxPackage.open(source).save(target)

    difference = diff(snapshot(source), snapshot(target))
    assert difference.is_empty, difference.describe()
    assert validate_package(target) == []


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_to_bytes_is_what_save_writes(name: str, dest: Path) -> None:
    # Compared canonically, not byte for byte: the OPC writer stamps every zip
    # member with the current time, so two serializations of the same package
    # differ in their archive metadata and in nothing else.
    pkg = DocxPackage.open(build(name))
    target = dest / "out.docx"
    pkg.save(target)

    difference = diff(snapshot(target), snapshot(pkg.to_bytes()))
    assert difference.is_empty, difference.describe()


# ----------------------------------------------------------------------------
# Parts and stories
# ----------------------------------------------------------------------------


def test_part_accepts_both_partname_forms(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    assert pkg.part("word/styles.xml") is pkg.part("/word/styles.xml")
    assert pkg.part("/word/styles.xml").content_type == CT.WML_STYLES


def test_part_raises_on_an_absent_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    with pytest.raises(PackageError):
        pkg.part("/word/footnotes.xml")
    assert pkg.find_part("/word/footnotes.xml") is None


def test_root_of_refuses_a_binary_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("drawings"))
    image = next(
        part for part in pkg.package.iter_parts() if part.content_type.startswith("image/")
    )
    with pytest.raises(PackageError):
        pkg.root_of(image)


def test_stories_cover_body_headers_footers_and_notes(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("combined"))
    names = [name for name, _ in pkg.stories()]

    assert names[0] == MAIN_STORY
    assert names == sorted(names, key=lambda n: (n != MAIN_STORY, n))
    assert {"document", "header1", "footer1", "footnotes", "endnotes"} <= set(names)
    # Comments are annotations attached to a story, never a story of their own.
    assert "comments" not in names


def test_stories_match_the_snapshot_story_ids(fixture_docx) -> None:
    path = fixture_docx("combined")
    pkg = DocxPackage.open(path)
    snapshot_stories = set(snapshot(path).stories)
    assert {name for name, _ in pkg.stories()} == snapshot_stories - {"comments"}


def test_story_roots_are_the_live_elements(fixture_docx, dest: Path) -> None:
    # word/footnotes.xml is one of the parts python-docx loads as an opaque blob;
    # editing a parsed copy of it would be silently dropped on save.
    pkg = DocxPackage.open(fixture_docx("footnotes"))
    footnotes = dict(pkg.stories())["footnotes"]
    text = footnotes.iter(qn("w:t")).__next__()
    text.text = "Edited through the story root."

    target = dest / "out.docx"
    pkg.save(target)

    assert "Edited through the story root." in snapshot(target).text("footnotes")


def test_every_xml_part_is_loaded_live(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("comments"))
    for part in pkg.package.iter_parts():
        if part.content_type.endswith("+xml"):
            assert isinstance(part, XmlPart), part.partname


def test_story_name_maps_part_names_to_snapshot_ids() -> None:
    assert story_name("/word/document.xml") == "document"
    assert story_name("word/header2.xml") == "header2"
    assert story_name("/docProps/core.xml") == "docProps/core.xml"


# ----------------------------------------------------------------------------
# ensure_part
# ----------------------------------------------------------------------------


def test_ensure_part_creates_a_related_part(fixture_docx, dest: Path) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    root = pkg.ensure_part(
        COMMENTS_EXTENDED_PARTNAME,
        CT_COMMENTS_EXTENDED,
        RT_COMMENTS_EXTENDED,
        COMMENTS_EXTENDED_XML,
    )
    assert root.tag == qn("w15:commentsEx")

    target = dest / "out.docx"
    pkg.save(target)

    reopened = DocxPackage.open(target)
    part = reopened.part(COMMENTS_EXTENDED_PARTNAME)
    assert part.content_type == CT_COMMENTS_EXTENDED
    assert validate_package(target) == []
    # The content type override made it into [Content_Types].xml, and the part is
    # reachable from word/_rels/document.xml.rels (validate_package checks both ways).
    after = snapshot(target)
    assert ("word/commentsExtended.xml", CT_COMMENTS_EXTENDED) in after.content_types
    assert "word/commentsExtended.xml" in after.parts
    assert (RT_COMMENTS_EXTENDED, "commentsExtended.xml") in after.relationships["word/document.xml"]


def test_ensure_part_is_idempotent(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    rels_before = len(pkg.document_part.rels)

    first = pkg.ensure_part(
        COMMENTS_EXTENDED_PARTNAME,
        CT_COMMENTS_EXTENDED,
        RT_COMMENTS_EXTENDED,
        COMMENTS_EXTENDED_XML,
    )
    first.set(qn("w15:paraId"), "0000BEEF")
    rels_after_create = len(pkg.document_part.rels)

    second = pkg.ensure_part(
        COMMENTS_EXTENDED_PARTNAME,
        CT_COMMENTS_EXTENDED,
        RT_COMMENTS_EXTENDED,
        COMMENTS_EXTENDED_XML,
    )

    assert second is first, "a second call must return the live root, not a fresh one"
    assert second.get(qn("w15:paraId")) == "0000BEEF", "existing content must survive"
    assert rels_after_create == rels_before + 1
    assert len(pkg.document_part.rels) == rels_after_create, "the relationship was duplicated"


def test_ensure_part_returns_the_root_of_an_existing_part(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    root = pkg.ensure_part("/word/styles.xml", CT.WML_STYLES, RT.STYLES, "<w:styles/>")
    assert root is pkg.root_of(pkg.part("/word/styles.xml"))
    assert root.find(qn("w:style")) is not None, "the existing styles were replaced"


def test_ensure_part_refuses_a_content_type_conflict(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    with pytest.raises(PackageError, match="content type"):
        pkg.ensure_part("/word/styles.xml", CT_COMMENTS_EXTENDED, RT.STYLES, "<w:styles/>")


def test_ensure_part_refuses_malformed_initial_xml(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    with pytest.raises(PackageError, match="well formed"):
        pkg.ensure_part(
            COMMENTS_EXTENDED_PARTNAME, CT_COMMENTS_EXTENDED, RT_COMMENTS_EXTENDED, "<w15:oops>"
        )


# ----------------------------------------------------------------------------
# Relationships
# ----------------------------------------------------------------------------


def test_rel_target_resolves_an_external_relationship(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("hyperlinks"))
    hyperlink = next(pkg.document.iter(qn("w:hyperlink")))
    r_id = hyperlink.get(qn("r:id"))
    assert pkg.rel_target(pkg.document_part, r_id) == "https://example.org/fixture"


def test_rel_target_accepts_a_story_root(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("hyperlinks"))
    r_id = next(pkg.document.iter(qn("w:hyperlink"))).get(qn("r:id"))
    assert pkg.rel_target(pkg.document, r_id) == pkg.rel_target(pkg.document_part, r_id)


def test_rel_target_resolves_an_internal_relationship(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("headers_footers"))
    reference = next(pkg.document.iter(qn("w:headerReference")))
    target = pkg.rel_target(pkg.document, reference.get(qn("r:id")))
    assert isinstance(target, Part)
    assert str(target.partname).startswith("/word/header")


def test_rel_target_raises_on_an_unknown_rid(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    with pytest.raises(PackageError, match="rId9999"):
        pkg.rel_target(pkg.document, "rId9999")


def test_rel_target_raises_on_a_foreign_element(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    other = DocxPackage.open(fixture_docx("mixed_runs"))
    with pytest.raises(PackageError, match="does not belong"):
        pkg.rel_target(other.document, "rId1")


def test_add_external_rel_deduplicates_by_target(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    first = pkg.add_external_rel(pkg.document, "https://example.org/a")
    again = pkg.add_external_rel(pkg.document_part, "https://example.org/a")
    other = pkg.add_external_rel(pkg.document, "https://example.org/b")

    assert first == again
    assert other != first
    targets = [rel.target_ref for rel in pkg.document_part.rels.values() if rel.is_external]
    assert targets.count("https://example.org/a") == 1


def test_add_external_rel_survives_a_save(fixture_docx, dest: Path) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    r_id = pkg.add_external_rel(pkg.document, "https://example.org/saved")
    target = dest / "out.docx"
    pkg.save(target)

    reopened = DocxPackage.open(target)
    assert reopened.rel_target(reopened.document, r_id) == "https://example.org/saved"
    assert validate_package(target) == []


def test_add_external_rel_refuses_an_empty_target(fixture_docx) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    with pytest.raises(PackageError):
        pkg.add_external_rel(pkg.document, "")


# ----------------------------------------------------------------------------
# Atomic save
# ----------------------------------------------------------------------------


def test_save_replaces_the_destination(fixture_docx, dest: Path) -> None:
    target = dest / "out.docx"
    target.write_bytes(b"previous content")
    returned = DocxPackage.open(fixture_docx("simple")).save(target)

    assert returned == target
    assert validate_package(target) == []
    assert list(dest.iterdir()) == [target]


def test_save_keeps_the_destination_intact_when_the_write_fails(
    fixture_docx, dest: Path, monkeypatch
) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    target = dest / "out.docx"
    target.write_bytes(b"previous content")

    def refuse(*_args, **_kwargs):
        raise OSError("no space left on device")

    monkeypatch.setattr(package_module.os, "replace", refuse)
    with pytest.raises(OSError, match="no space left"):
        pkg.save(target)

    assert target.read_bytes() == b"previous content"
    assert list(dest.iterdir()) == [target], "a temporary file was left behind"


def test_save_does_not_touch_the_destination_when_serialization_fails(
    fixture_docx, dest: Path, monkeypatch
) -> None:
    pkg = DocxPackage.open(fixture_docx("simple"))
    target = dest / "out.docx"
    target.write_bytes(b"previous content")

    def explode(_self):
        raise RuntimeError("serialization blew up")

    monkeypatch.setattr(DocxPackage, "to_bytes", explode)
    with pytest.raises(RuntimeError):
        pkg.save(target)

    assert target.read_bytes() == b"previous content"
    assert list(dest.iterdir()) == [target]


def test_save_creates_a_new_file(fixture_docx, dest: Path) -> None:
    target = dest / "new.docx"
    DocxPackage.open(fixture_docx("simple")).save(target)
    assert validate_package(target) == []
    assert list(dest.iterdir()) == [target]


def test_atomic_write_bytes_replaces_content(dest: Path) -> None:
    target = dest / "value.bin"
    atomic_write_bytes(target, b"first")
    atomic_write_bytes(target, b"second")
    assert target.read_bytes() == b"second"
    assert list(dest.iterdir()) == [target]


def test_atomic_write_bytes_removes_the_temporary_on_failure(dest: Path) -> None:
    target = dest / "value.bin"
    target.write_bytes(b"first")
    with pytest.raises(TypeError):
        atomic_write_bytes(target, "not bytes")  # type: ignore[arg-type]
    assert target.read_bytes() == b"first"
    assert list(dest.iterdir()) == [target]


def test_atomic_write_bytes_accepts_a_bare_filename(dest: Path, monkeypatch) -> None:
    monkeypatch.chdir(dest)
    atomic_write_bytes("relative.bin", b"payload")
    assert (dest / "relative.bin").read_bytes() == b"payload"
    assert os.listdir(dest) == ["relative.bin"]
