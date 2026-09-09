"""Tests for :mod:`word_document_server.engine.merge`.

What is pinned here is the contract of an *import*: the blocks the source
carried arrive in the target with their runs, their formatting, their images and
their ids intact, and everything the import could not carry faithfully is either
refused or reported. The target's own content -- and in particular its final
``w:sectPr`` -- comes out untouched.

Two instrument traps of this repository apply throughout (see
``docs/plans/v2-semantic-engine``): ``Diff.is_empty`` is a *method*, so it is
always called with parentheses; and ``TableSignature`` only counts the columns of
a ``w:tblGrid``, so a test about an imported table's geometry checks the
``w:gridCol`` widths itself rather than relying on ``assert_unchanged_except``.
"""

from __future__ import annotations

import copy
import io
import struct
import zlib

import pytest
from docx import Document
from docx.shared import Inches
from lxml import etree

from tests.fixtures.builders import build
from tests.support.package_check import validate_package
from tests.support.snapshot import (
    ParagraphSignature,
    Snapshot,
    assert_unchanged_except,
    diff,
    snapshot,
)
from word_document_server.engine.errors import PackageError
from word_document_server.engine.merge import MergeRefused, append_document
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import R, qn

R_NAMESPACE = "{" + R + "}"
W_BODY = qn("w:body")
W_SECT_PR = qn("w:sectPr")
W_P = qn("w:p")
W_TBL = qn("w:tbl")
W_VAL = qn("w:val")
W_ID = qn("w:id")


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------


def _pkg(name: str) -> DocxPackage:
    """A fresh package built from the named fixture."""
    return DocxPackage.open(build(name))


def _body(pkg: DocxPackage) -> etree._Element:
    return pkg.document.find(W_BODY)


def _blocks(pkg: DocxPackage) -> list[etree._Element]:
    """Body children of `pkg` but its final ``w:sectPr``."""
    body = _body(pkg)
    sect_pr = body[-1] if len(body) and body[-1].tag == W_SECT_PR else None
    return [child for child in body if child is not sect_pr]


def _texts(pkg: DocxPackage) -> list[str]:
    return [visible_text(paragraph) for paragraph in pkg.document.iter(W_P)]


def _root_of(pkg: DocxPackage, partname: str) -> etree._Element:
    return pkg.root_of(pkg.part(partname))


def _style(pkg: DocxPackage, style_id: str) -> etree._Element | None:
    for element in _root_of(pkg, "/word/styles.xml").iter(qn("w:style")):
        if element.get(qn("w:styleId")) == style_id:
            return element
    return None


def _num_ids(pkg: DocxPackage) -> list[str]:
    """``w:numPr/w:numId`` values of the body, in document order."""
    return [element.get(W_VAL) for element in _body(pkg).iter(qn("w:numId"))]


def _attribute_values(root: etree._Element, tag: str, attribute: str) -> list[str]:
    return [element.get(attribute) for element in root.iter(qn(tag))]


def _relationship_references(root: etree._Element) -> list[str]:
    """Every relationship id referenced under `root`, whatever attribute holds it."""
    return [
        value
        for element in root.iter()
        if isinstance(element.tag, str)
        for name, value in element.attrib.items()
        if name.startswith(R_NAMESPACE)
    ]


def _canonical(element: etree._Element) -> bytes:
    return etree.tostring(element, method="c14n", exclusive=True)


def _shape(signature: ParagraphSignature) -> tuple:
    """Everything of a paragraph signature but where it sits."""
    return (
        signature.style,
        signature.text,
        signature.runs,
        signature.deleted_runs,
        signature.markers,
        signature.field_instructions,
        signature.ppr,
        signature.inline_controls,
        signature.enclosing_controls,
    )


def _shapes(snap: Snapshot) -> list[tuple]:
    return [_shape(signature) for signature in snap.story_paragraphs("document")]


def _png(red: int, green: int, blue: int) -> bytes:
    """A deterministic 1x1 PNG of the given colour."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        payload = tag + data
        return struct.pack(">I", len(data)) + payload + struct.pack(">I", zlib.crc32(payload))

    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    idat = zlib.compress(bytes([0, red, green, blue]), 9)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


#: The exact bytes the ``drawings`` and ``headers_footers`` fixtures embed.
FIXTURE_PNG = _png(255, 0, 0)


def _document_with_image(image: bytes) -> DocxPackage:
    """A one-picture package, so that the image bytes are known to the test."""
    document = Document()
    document.add_paragraph("Picture below.")
    document.add_picture(io.BytesIO(image), width=Inches(1))
    buffer = io.BytesIO()
    document.save(buffer)
    return DocxPackage.open(buffer.getvalue())


def _emptied(pkg: DocxPackage) -> DocxPackage:
    """Strip every block of the body, keeping the final ``w:sectPr``."""
    body = _body(pkg)
    for child in list(body):
        if child.tag != W_SECT_PR:
            body.remove(child)
    return pkg


# --------------------------------------------------------------------------------------
# Where the blocks land
# --------------------------------------------------------------------------------------


def test_blocks_are_appended_before_the_final_sect_pr() -> None:
    target, source = _pkg("simple"), _pkg("tables")
    kept = [child.tag for child in _blocks(target)]
    imported = [child.tag for child in _blocks(source)]

    report = append_document(target, source)

    body = _body(target)
    assert body[-1].tag == W_SECT_PR, "the target must keep closing on its section properties"
    assert [child.tag for child in _blocks(target)] == kept + imported
    assert report.blocks == len(imported)


def test_the_target_keeps_its_paragraphs_and_its_section_properties() -> None:
    target, source = _pkg("simple"), _pkg("simple")
    before = snapshot(target.to_bytes())

    append_document(target, source)

    after = snapshot(target.to_bytes())
    added = [key for key in after.paragraphs if key not in before.paragraphs]
    assert len(added) == len(before.story_paragraphs("document"))
    # Nothing but the appended paragraphs moved: the sect_pr of the story stays
    # guarded here, because only naming its part in `parts` would relax it.
    assert_unchanged_except(before, after, paragraphs=added)


def test_an_imported_table_keeps_its_grid_widths() -> None:
    # D-004: TableSignature only counts w:gridCol, so the widths are checked here.
    target, source = _pkg("simple"), _pkg("tables")
    expected = [
        _attribute_values(table, "w:gridCol", qn("w:w"))
        for table in _body(source).iter(W_TBL)
    ]

    append_document(target, source)

    assert expected  # sanity: the fixture does carry a grid
    assert [
        _attribute_values(table, "w:gridCol", qn("w:w"))
        for table in _body(target).iter(W_TBL)
    ] == expected


@pytest.mark.parametrize(
    "fixture",
    ["mixed_runs", "tables", "fields", "content_controls", "bookmarks", "tracked_changes"],
)
def test_imported_paragraphs_are_the_source_paragraphs(fixture: str) -> None:
    target, source = _pkg("simple"), _pkg(fixture)
    expected = _shapes(snapshot(source.to_bytes()))
    kept = len(_shapes(snapshot(target.to_bytes())))

    append_document(target, source)

    assert _shapes(snapshot(target.to_bytes()))[kept:] == expected


# --------------------------------------------------------------------------------------
# Page breaks
# --------------------------------------------------------------------------------------


def test_page_break_is_inserted_between_the_two_documents() -> None:
    target, source = _pkg("simple"), _pkg("simple")
    kept = len(_blocks(target))

    report = append_document(target, source, page_break=True)

    blocks = _blocks(target)
    separator = blocks[kept]
    assert [element.get(qn("w:type")) for element in separator.iter(qn("w:br"))] == ["page"]
    assert list(separator.iter(qn("w:t"))) == [], "the separator carries the break and nothing else"
    # The separator is the tool's own doing, not a block of the source.
    assert report.blocks == len(blocks) - kept - 1


def test_no_page_break_without_it() -> None:
    target = _pkg("simple")
    append_document(target, _pkg("simple"), page_break=False)

    assert [
        element.get(qn("w:type")) for element in _body(target).iter(qn("w:br"))
    ] == []


def test_no_page_break_when_the_target_has_nothing_to_break_away_from() -> None:
    target = _emptied(_pkg("simple"))
    source = _pkg("simple")

    append_document(target, source, page_break=True)

    assert _texts(target) == _texts(source)


# --------------------------------------------------------------------------------------
# Relationships: media and external links
# --------------------------------------------------------------------------------------


def test_a_referenced_image_is_copied_into_the_target() -> None:
    target = _pkg("drawings")
    green = _png(0, 255, 0)

    report = append_document(target, _document_with_image(green))

    assert report.parts == ("/word/media/image2.png",)
    assert target.part("/word/media/image2.png").blob == green
    # The image the target already had is untouched, name and bytes alike.
    assert target.part("/word/media/image1.png").blob == FIXTURE_PNG


def test_an_image_the_target_already_holds_is_not_duplicated() -> None:
    target = _pkg("drawings")

    report = append_document(target, _document_with_image(FIXTURE_PNG))

    assert report.parts == ()
    assert [
        str(part.partname)
        for part in target.package.iter_parts()
        if str(part.partname).startswith("/word/media/")
    ] == ["/word/media/image1.png"]


def test_every_imported_image_reference_resolves_in_the_target() -> None:
    target = _pkg("simple")
    report = append_document(target, _pkg("drawings"))

    assert report.parts == ("/word/media/image1.png",)
    references = _relationship_references(_body(target))
    assert len(references) == 2  # the fixture embeds the same picture twice
    for rId in references:
        assert target.rel_target(target.document, rId) is target.part("/word/media/image1.png")


def test_an_external_hyperlink_is_recreated_as_an_external_relationship() -> None:
    target = _pkg("simple")
    append_document(target, _pkg("hyperlinks"))

    links = [
        element
        for element in _body(target).iter(qn("w:hyperlink"))
        if element.get(qn("r:id")) is not None
    ]
    assert len(links) == 1
    assert target.rel_target(target.document, links[0].get(qn("r:id"))) == (
        "https://example.org/fixture"
    )
    # The anchor-based link has no relationship and must keep its anchor.
    anchors = [element.get(qn("w:anchor")) for element in _body(target).iter(qn("w:hyperlink"))]
    assert "FixtureAnchor" in anchors


def test_a_dangling_relationship_reference_is_refused() -> None:
    source = _pkg("hyperlinks")
    link = next(
        element
        for element in _body(source).iter(qn("w:hyperlink"))
        if element.get(qn("r:id")) is not None
    )
    link.set(qn("r:id"), "rIdNotThere")

    with pytest.raises(PackageError, match="rIdNotThere"):
        append_document(_pkg("simple"), source)


# --------------------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------------------


def test_missing_styles_are_copied_with_their_based_on_chain() -> None:
    target, source = _pkg("simple"), _pkg("style_inheritance")
    assert _style(target, "FixtureLeaf") is None  # sanity

    report = append_document(target, source)

    # FixtureLeaf is the only style the body names; Branch, Root and the linked
    # character style come with it, through basedOn and link.
    assert set(report.styles) == {
        "FixtureRoot",
        "FixtureBranch",
        "FixtureLeaf",
        "FixtureBranchChar",
    }
    for style_id in report.styles:
        assert _canonical(_style(target, style_id)) == _canonical(_style(source, style_id))
    assert [
        signature.style for signature in snapshot(target.to_bytes()).story_paragraphs("document")
    ][-4:] == ["FixtureRoot", "FixtureBranch", "FixtureLeaf", None]


def test_a_style_the_target_defines_differently_is_kept_and_reported() -> None:
    target, source = _pkg("style_inheritance"), _pkg("style_inheritance")
    size = _style(target, "FixtureRoot").find(qn("w:rPr")).find(qn("w:sz"))
    size.set(W_VAL, "40")

    report = append_document(target, source)

    assert report.styles == ()
    assert any("FixtureRoot" in warning for warning in report.warnings)
    assert _style(target, "FixtureRoot").find(qn("w:rPr")).find(qn("w:sz")).get(W_VAL) == "40"


def test_a_style_defined_nowhere_is_reported_rather_than_invented() -> None:
    source = _pkg("simple")
    paragraph = _body(source).find(W_P)
    paragraph.find(qn("w:pPr")).find(qn("w:pStyle")).set(W_VAL, "NoSuchStyle")

    report = append_document(_pkg("simple"), source)

    assert report.styles == ()
    assert any("NoSuchStyle" in warning for warning in report.warnings)


# --------------------------------------------------------------------------------------
# Numbering
# --------------------------------------------------------------------------------------


def test_imported_lists_are_renumbered_and_keep_their_definition() -> None:
    target, source = _pkg("complex_numbering"), _pkg("complex_numbering")
    before = _num_ids(target)

    report = append_document(target, source)

    assert set(report.numbering) == {900, 901, 902}
    assert set(report.numbering.values()).isdisjoint({900, 901, 902})
    # The target's own paragraphs keep pointing at their own lists.
    assert _num_ids(target)[: len(before)] == before
    assert _num_ids(target)[len(before) :] == [
        str(report.numbering[int(value)]) for value in before
    ]

    numbering = _root_of(target, "/word/numbering.xml")
    abstract_of = {
        element.get(qn("w:numId")): element.find(qn("w:abstractNumId")).get(W_VAL)
        for element in numbering.iter(qn("w:num"))
    }
    # 900 and 901 shared one abstract definition in the source; their copies
    # must still share one, and it must not be the one 902 uses.
    assert abstract_of[str(report.numbering[900])] == abstract_of[str(report.numbering[901])]
    assert abstract_of[str(report.numbering[900])] != abstract_of[str(report.numbering[902])]


def test_an_imported_list_keeps_its_levels_and_its_start_override() -> None:
    target, source = _pkg("complex_numbering"), _pkg("complex_numbering")

    report = append_document(target, source)

    numbering = _root_of(target, "/word/numbering.xml")
    by_num_id = {element.get(qn("w:numId")): element for element in numbering.iter(qn("w:num"))}
    override = by_num_id[str(report.numbering[901])].find(qn("w:lvlOverride"))
    assert override.find(qn("w:startOverride")).get(W_VAL) == "5"

    by_abstract = {
        element.get(qn("w:abstractNumId")): element
        for element in numbering.iter(qn("w:abstractNum"))
    }
    imported = by_abstract[
        by_num_id[str(report.numbering[900])].find(qn("w:abstractNumId")).get(W_VAL)
    ]
    original = by_abstract["900"]
    assert [_canonical(level) for level in imported.findall(qn("w:lvl"))] == [
        _canonical(level) for level in original.findall(qn("w:lvl"))
    ]


def test_a_list_the_source_never_defined_loses_its_numbering_rather_than_stealing_one() -> None:
    target, source = _pkg("complex_numbering"), _pkg("complex_numbering")
    numbering = _root_of(source, "/word/numbering.xml")
    doomed = next(
        element for element in numbering.iter(qn("w:num")) if element.get(qn("w:numId")) == "900"
    )
    doomed.getparent().remove(doomed)

    report = append_document(target, source)

    assert 900 not in report.numbering
    assert any("900" in warning for warning in report.warnings)
    assert "900" not in _num_ids(target)[len(_num_ids(_pkg("complex_numbering"))) :]


# --------------------------------------------------------------------------------------
# What is refused, what is warned about
# --------------------------------------------------------------------------------------


def _with_endnote_reference(pkg: DocxPackage) -> DocxPackage:
    paragraph = _body(pkg).find(W_P)
    run = etree.SubElement(paragraph, qn("w:r"))
    etree.SubElement(run, qn("w:endnoteReference")).set(W_ID, "2")
    return pkg


@pytest.mark.parametrize(
    ("source", "reason"),
    [
        (lambda: _pkg("comments"), "comments"),
        (lambda: _pkg("footnotes"), "footnotes"),
        (lambda: _with_endnote_reference(_pkg("simple")), "endnotes"),
    ],
)
def test_a_source_carrying_annotations_is_refused_before_the_target_is_touched(
    source, reason: str
) -> None:
    target = _pkg("simple")
    before = snapshot(target.to_bytes())

    with pytest.raises(MergeRefused) as raised:
        append_document(target, source())

    assert raised.value.reason == reason
    assert reason in str(raised.value)
    # Diff.is_empty is a method: calling it is what makes this assertion real.
    assert diff(before, snapshot(target.to_bytes())).is_empty()


def test_headers_and_footers_of_the_source_are_reported_not_imported() -> None:
    target, source = _pkg("simple"), _pkg("headers_footers")
    before = sorted(story for story, _ in target.stories())

    report = append_document(target, source)

    assert any("headers and footers" in warning for warning in report.warnings)
    assert sorted(story for story, _ in target.stories()) == before
    assert "Body text; the header and footer carry the interesting parts." in _texts(target)


def test_an_imported_section_break_loses_its_header_references() -> None:
    source = _pkg("headers_footers")
    body = _body(source)
    section = body[-1]
    assert section.find(qn("w:headerReference")) is not None  # sanity
    paragraph = body.find(W_P)
    paragraph.find(qn("w:pPr")).append(copy.deepcopy(section))

    target = _pkg("simple")
    report = append_document(target, source)

    imported = [
        element
        for element in _body(target).iter(W_SECT_PR)
        if element.getparent().tag == qn("w:pPr")
    ]
    assert len(imported) == 1, "the section break itself must survive the import"
    assert imported[0].find(qn("w:headerReference")) is None
    assert imported[0].find(qn("w:footerReference")) is None
    assert imported[0].find(qn("w:pgSz")) is not None, "the page setup must survive"
    assert any("section break" in warning for warning in report.warnings)


# --------------------------------------------------------------------------------------
# Identifiers
# --------------------------------------------------------------------------------------


def test_colliding_bookmark_ids_are_renumbered_start_and_end_together() -> None:
    target, source = _pkg("bookmarks"), _pkg("bookmarks")

    report = append_document(target, source)

    starts = _attribute_values(_body(target), "w:bookmarkStart", W_ID)
    ends = _attribute_values(_body(target), "w:bookmarkEnd", W_ID)
    assert len(starts) == 2 * len(_attribute_values(_body(_pkg("bookmarks")), "w:bookmarkStart", W_ID))
    assert len(set(starts)) == len(starts)
    assert sorted(ends) == sorted(starts)
    assert any("FixtureInline" in warning for warning in report.warnings)


def test_ids_that_do_not_collide_are_left_alone() -> None:
    target, source = _pkg("simple"), _pkg("bookmarks")
    expected = _attribute_values(_body(source), "w:bookmarkStart", W_ID)

    append_document(target, source)

    assert _attribute_values(_body(target), "w:bookmarkStart", W_ID) == expected


def test_colliding_revision_ids_are_renumbered() -> None:
    target, source = _pkg("tracked_changes"), _pkg("tracked_changes")
    before = _attribute_values(_body(target), "w:ins", W_ID) + _attribute_values(
        _body(target), "w:del", W_ID
    )

    append_document(target, source)

    after = _attribute_values(_body(target), "w:ins", W_ID) + _attribute_values(
        _body(target), "w:del", W_ID
    )
    assert len(after) == 2 * len(before)
    assert len(set(after)) == len(after)


def test_colliding_para_ids_are_renumbered() -> None:
    target, source = _pkg("simple"), _pkg("simple")
    for pkg in (target, source):
        _body(pkg).find(W_P).set(qn("w14:paraId"), "0000ABCD")

    append_document(target, source)

    used = [
        element.get(qn("w14:paraId"))
        for element in _body(target).iter()
        if element.get(qn("w14:paraId")) is not None
    ]
    assert len(used) == 2
    assert len(set(used)) == 2
    assert all(len(value) == 8 for value in used)


# --------------------------------------------------------------------------------------
# The result is a package Word can open
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "fixture",
    [
        "simple",
        "tables",
        "drawings",
        "hyperlinks",
        "style_inheritance",
        "complex_numbering",
        "bookmarks",
        "tracked_changes",
        "content_controls",
        "fields",
        "text_boxes",
        "sections",
        "headers_footers",
    ],
)
def test_the_merged_package_is_valid(tmp_path, fixture: str) -> None:
    target = _pkg("simple")
    append_document(target, _pkg(fixture), page_break=True)
    path = target.save(tmp_path / f"{fixture}.docx")

    assert validate_package(path) == []


@pytest.mark.parametrize(
    "fixture",
    ["drawings", "bookmarks", "complex_numbering", "tracked_changes", "tables", "hyperlinks"],
)
def test_merging_a_document_into_itself_stays_valid(tmp_path, fixture: str) -> None:
    target = _pkg(fixture)
    append_document(target, _pkg(fixture), page_break=True)
    path = target.save(tmp_path / f"self-{fixture}.docx")

    assert validate_package(path) == []
