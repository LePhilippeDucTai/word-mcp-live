"""Tests for :mod:`word_document_server.engine.format`.

Most cases are written on hand-built paragraphs whose XML is visible in the test:
what matters about this layer is the tree it leaves behind -- the order of the
``w:rPr`` children, the attributes it clears, the runs it does *not* touch --
and that cannot be read from the text.  The last section then does the same work
on a real fixture package, saves it, and checks at package scale that nothing
outside the two edited paragraphs moved.
"""

from __future__ import annotations

from functools import cache
from pathlib import Path

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.engine.errors import LocatorError, UnsupportedRange
from word_document_server.engine.format import (
    PPR_ORDER,
    RPR_ORDER,
    apply_rpr,
    comment,
    hyperlink,
    set_ppr_child,
    unwrap,
    wrap,
)
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import resolve
from word_document_server.engine.textmodel import segments, visible_text
from word_document_server.engine.xmlns import NAMESPACES, qn

W_P = qn("w:p")
W_R = qn("w:r")
W_RPR = qn("w:rPr")
W_PPR = qn("w:pPr")
W_VAL = qn("w:val")
W_ID = qn("w:id")


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------


def paragraph_from(children: str) -> etree._Element:
    """Return a standalone ``w:p`` whose children are the given XML fragment."""
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in NAMESPACES.items())
    return etree.fromstring(f"<w:p {declarations}>{children}</w:p>")


def names(element: etree._Element | None) -> list[str]:
    """Local names of the element's children, in order."""
    return [] if element is None else [etree.QName(child).localname for child in element]


def rpr_names(run: etree._Element) -> list[str]:
    """Local names of the run's ``w:rPr`` children, in order."""
    return names(run.find(W_RPR))


def runs(paragraph: etree._Element) -> list[etree._Element]:
    """Every ``w:r`` of the paragraph, in document order."""
    return list(paragraph.iter(W_R))


def canonical(element: etree._Element | None) -> str:
    if element is None:
        return ""
    return etree.tostring(element, method="c14n", exclusive=True).decode()


def formats(paragraph: etree._Element) -> list[str]:
    """The canonical ``w:rPr`` in force for each visible character."""
    found: list[str] = []
    for segment in segments(paragraph):
        if not segment.text:
            continue
        run = segment.run
        found.extend([canonical(None if run is None else run.find(W_RPR))] * len(segment.text))
    return found


def marker_ids(paragraph: etree._Element, tag: str) -> list[str]:
    return [element.get(W_ID) for element in paragraph.iter(qn(tag))]


@cache
def fixture_bytes(name: str) -> bytes:
    """Build a fixture once per session; the builders are deterministic."""
    return build(name)


@cache
def untouched_bytes(name: str) -> bytes:
    """The fixture as :class:`DocxPackage` rewrites it, with no edit at all.

    The reference of every package-scale comparison below, for the reason
    ``tests/engine/test_ranges_properties.py`` states: opening and saving is not
    byte-neutral today, and that belongs to ``engine/package.py``, not here.
    """
    return DocxPackage.open(fixture_bytes(name)).to_bytes()


def body_paragraphs(package: DocxPackage) -> list[etree._Element]:
    """The paragraphs of the main story, in snapshot-index order."""
    _, root = package.stories()[0]
    return list(root.iter(W_P))


def paragraph_starting_with(package: DocxPackage, prefix: str) -> tuple[int, etree._Element]:
    """The first body paragraph whose visible text starts with `prefix`."""
    for index, paragraph in enumerate(body_paragraphs(package)):
        if visible_text(paragraph).startswith(prefix):
            return index, paragraph
    raise AssertionError(f"no paragraph starting with {prefix!r}")


# --------------------------------------------------------------------------------------
# Schema order
# --------------------------------------------------------------------------------------


def test_the_two_schema_tables_hold_no_duplicate() -> None:
    assert len(set(RPR_ORDER)) == len(RPR_ORDER)
    assert len(set(PPR_ORDER)) == len(PPR_ORDER)


def test_rpr_children_are_written_in_schema_order() -> None:
    paragraph = paragraph_from('<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>abc</w:t></w:r>')
    pieces = resolve(paragraph, 0, 3)

    # Deliberately out of schema order: the patch is a set, not a sequence.
    apply_rpr(pieces, {"highlight": "yellow", "bold": True, "color": "FF0000",
                       "size_pt": 12, "font": "Consolas"})

    assert rpr_names(paragraph[0]) == [
        "rFonts",
        "b",
        "color",
        "sz",
        "szCs",
        "highlight",
        "u",
    ]


def test_a_new_property_lands_before_the_property_revision() -> None:
    paragraph = paragraph_from(
        "<w:r><w:rPr>"
        '<w:rPrChange w:id="7" w:author="A" w:date="2024-01-01T00:00:00Z"><w:rPr/></w:rPrChange>'
        "</w:rPr><w:t>abc</w:t></w:r>"
    )
    apply_rpr(resolve(paragraph, 0, 3), {"bold": True})

    assert rpr_names(paragraph[0]) == ["b", "rPrChange"]


def test_an_unknown_child_of_the_rpr_is_left_where_it_is() -> None:
    paragraph = paragraph_from(
        '<w:r><w:rPr><w:snapToGrid/><w:lang w:val="fr-FR"/></w:rPr><w:t>abc</w:t></w:r>'
    )
    apply_rpr(resolve(paragraph, 0, 3), {"italic": True})

    assert rpr_names(paragraph[0]) == ["i", "snapToGrid", "lang"]


# --------------------------------------------------------------------------------------
# What a patch touches, and what it does not
# --------------------------------------------------------------------------------------


THREE_RUNS = (
    '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">one </w:t></w:r>'
    '<w:r><w:rPr><w:rStyle w:val="FixtureEmphasis"/></w:rPr><w:t xml:space="preserve">two </w:t></w:r>'
    "<w:r><w:t>three</w:t></w:r>"
)


def test_formatting_a_subrange_leaves_the_other_runs_untouched() -> None:
    paragraph = paragraph_from(THREE_RUNS)
    before = [canonical(run) for run in runs(paragraph)]

    patched = apply_rpr(resolve(paragraph, 4, 8), {"bold": True})

    after = runs(paragraph)
    assert visible_text(paragraph) == "one two three"
    assert patched == (after[1],)
    assert canonical(after[0]) == before[0]
    assert canonical(after[2]) == before[2]
    # The run keeps its character style: the patch adds, it does not rebuild.
    assert rpr_names(after[1]) == ["rStyle", "b"]


def test_formatting_part_of_a_run_splits_it_and_spares_the_rest() -> None:
    paragraph = paragraph_from('<w:r><w:rPr><w:i/></w:rPr><w:t>abcdef</w:t></w:r>')

    apply_rpr(resolve(paragraph, 2, 4), {"bold": True})

    assert visible_text(paragraph) == "abcdef"
    assert formats(paragraph) == (
        [canonical(runs(paragraph)[0].find(W_RPR))] * 2
        + [canonical(runs(paragraph)[1].find(W_RPR))] * 2
        + [canonical(runs(paragraph)[2].find(W_RPR))] * 2
    )
    # ``w:b`` outranks ``w:i`` in CT_RPr, so the added property comes first.
    assert [rpr_names(run) for run in runs(paragraph)] == [["i"], ["b", "i"], ["i"]]


def test_every_covered_run_is_patched() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    patched = apply_rpr(resolve(paragraph, 0, 13), {"italic": True})

    assert len(patched) == 3
    assert all("i" in rpr_names(run) for run in runs(paragraph))


def test_an_empty_range_formats_nothing() -> None:
    paragraph = paragraph_from(THREE_RUNS)
    before = etree.tostring(paragraph)

    assert apply_rpr(resolve(paragraph, 4, 4), {"bold": True}) == ()
    assert etree.tostring(paragraph) == before


def test_removing_the_last_property_removes_the_rpr() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")
    before = etree.tostring(paragraph)

    apply_rpr(resolve(paragraph, 0, 3), {"bold": True})
    assert rpr_names(paragraph[0]) == ["b"]

    apply_rpr(resolve(paragraph, 0, 3), {"bold": None})
    assert etree.tostring(paragraph) == before


def test_removing_one_property_keeps_the_others() -> None:
    paragraph = paragraph_from(
        '<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>abc</w:t></w:r>'
    )

    apply_rpr(resolve(paragraph, 0, 3), {"underline": None})

    assert rpr_names(paragraph[0]) == ["b", "i"]


def test_a_false_toggle_is_written_as_an_explicit_off() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    apply_rpr(resolve(paragraph, 0, 3), {"bold": False, "small_caps": False})
    properties = paragraph[0].find(W_RPR)

    assert properties.find(qn("w:b")).get(W_VAL) == "0"
    assert properties.find(qn("w:smallCaps")).get(W_VAL) == "0"

    apply_rpr(resolve(paragraph, 0, 3), {"bold": True})
    assert properties.find(qn("w:b")).get(W_VAL) is None


@pytest.mark.parametrize(
    ("value", "expected"),
    [(True, "single"), (False, "none"), ("dotted", "dotted")],
)
def test_underline_accepts_a_flag_or_a_style(value: object, expected: str) -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    apply_rpr(resolve(paragraph, 0, 3), {"underline": value})

    assert paragraph[0].find(W_RPR).find(qn("w:u")).get(W_VAL) == expected


@pytest.mark.parametrize(
    ("size", "half_points"), [(12, "24"), (10.5, "21"), (8.0, "16")]
)
def test_size_is_stored_in_half_points_on_both_scripts(
    size: float, half_points: str
) -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    apply_rpr(resolve(paragraph, 0, 3), {"size_pt": size})
    properties = paragraph[0].find(W_RPR)

    assert properties.find(qn("w:sz")).get(W_VAL) == half_points
    assert properties.find(qn("w:szCs")).get(W_VAL) == half_points

    apply_rpr(resolve(paragraph, 0, 3), {"size_pt": None})
    assert paragraph[0].find(W_RPR) is None


def test_superscript_and_subscript_share_one_element() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")
    pieces = resolve(paragraph, 0, 3)

    apply_rpr(pieces, {"superscript": True})
    assert paragraph[0].find(W_RPR).find(qn("w:vertAlign")).get(W_VAL) == "superscript"

    apply_rpr(pieces, {"subscript": True})
    assert rpr_names(paragraph[0]) == ["vertAlign"]
    assert paragraph[0].find(W_RPR).find(qn("w:vertAlign")).get(W_VAL) == "subscript"

    apply_rpr(pieces, {"superscript": True, "subscript": False})
    assert paragraph[0].find(W_RPR).find(qn("w:vertAlign")).get(W_VAL) == "superscript"

    apply_rpr(pieces, {"superscript": False, "subscript": False})
    assert paragraph[0].find(W_RPR).find(qn("w:vertAlign")).get(W_VAL) == "baseline"

    apply_rpr(pieces, {"superscript": None})
    assert paragraph[0].find(W_RPR) is None


@pytest.mark.parametrize(
    ("patch", "error"),
    [
        ({"superscript": True, "subscript": True}, ValueError),
        ({"underline": "squiggly"}, ValueError),
        ({"highlight": "puce"}, ValueError),
        ({"color": "12345"}, ValueError),
        ({"color": "GGGGGG"}, ValueError),
        ({"size_pt": 10.3}, ValueError),
        ({"size_pt": 0}, ValueError),
        ({"size_pt": 2000}, ValueError),
        ({"font": ""}, ValueError),
        ({"weight": True}, ValueError),
        ({"bold": "yes"}, TypeError),
        ({"underline": 3}, TypeError),
        ({"highlight": 3}, TypeError),
        ({"color": 0xFF0000}, TypeError),
        ({"size_pt": "12"}, TypeError),
        ({"font": 12}, TypeError),
        ({"superscript": "up"}, TypeError),
    ],
)
def test_an_unusable_patch_is_refused_before_anything_is_written(
    patch: dict, error: type[Exception]
) -> None:
    paragraph = paragraph_from(THREE_RUNS)
    before = etree.tostring(paragraph)
    pieces = resolve(paragraph, 0, 13)

    with pytest.raises(error):
        apply_rpr(pieces, patch)
    assert etree.tostring(paragraph) == before


def test_a_valid_property_next_to_an_invalid_one_is_not_applied() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")
    pieces = resolve(paragraph, 0, 3)

    with pytest.raises(ValueError):
        apply_rpr(pieces, {"bold": True, "highlight": "puce"})

    assert paragraph[0].find(W_RPR) is None


def test_apply_rpr_wants_the_pieces_of_a_resolved_range() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    with pytest.raises(TypeError):
        apply_rpr(runs(paragraph), {"bold": True})


# --------------------------------------------------------------------------------------
# Theme properties
# --------------------------------------------------------------------------------------


THEMED = (
    "<w:r><w:rPr>"
    '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" '
    'w:eastAsia="MS Mincho" w:cstheme="majorBidi"/>'
    '<w:color w:val="4472C4" w:themeColor="accent1" w:themeTint="66" w:themeShade="BF"/>'
    "</w:rPr><w:t>abc</w:t></w:r>"
)


def test_an_explicit_font_replaces_the_theme_fonts() -> None:
    paragraph = paragraph_from(THEMED)

    apply_rpr(resolve(paragraph, 0, 3), {"font": "Consolas"})
    fonts = paragraph[0].find(W_RPR).find(qn("w:rFonts"))

    assert fonts.get(qn("w:ascii")) == "Consolas"
    assert fonts.get(qn("w:hAnsi")) == "Consolas"
    assert fonts.get(qn("w:cs")) == "Consolas"
    assert fonts.get(qn("w:asciiTheme")) is None
    assert fonts.get(qn("w:hAnsiTheme")) is None
    assert fonts.get(qn("w:cstheme")) is None
    # A script this layer does not write is left alone.
    assert fonts.get(qn("w:eastAsia")) == "MS Mincho"


def test_an_explicit_colour_replaces_the_theme_colour() -> None:
    paragraph = paragraph_from(THEMED)

    apply_rpr(resolve(paragraph, 0, 3), {"color": "#ff0000"})
    color = paragraph[0].find(W_RPR).find(qn("w:color"))

    assert color.get(W_VAL) == "FF0000"
    assert color.get(qn("w:themeColor")) is None
    assert color.get(qn("w:themeTint")) is None
    assert color.get(qn("w:themeShade")) is None


def test_removing_the_font_drops_the_rfonts_only_when_nothing_is_left() -> None:
    themed = paragraph_from(THEMED)
    apply_rpr(resolve(themed, 0, 3), {"font": None})
    fonts = themed[0].find(W_RPR).find(qn("w:rFonts"))
    assert fonts is not None
    assert dict(fonts.attrib) == {qn("w:eastAsia"): "MS Mincho"}

    plain = paragraph_from(
        '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>abc</w:t></w:r>'
    )
    apply_rpr(resolve(plain, 0, 3), {"font": None})
    assert plain[0].find(W_RPR) is None


def test_removing_the_colour_removes_the_whole_element() -> None:
    paragraph = paragraph_from(THEMED)

    apply_rpr(resolve(paragraph, 0, 3), {"color": None})

    assert rpr_names(paragraph[0]) == ["rFonts"]


def test_auto_is_a_colour_value_of_its_own() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    apply_rpr(resolve(paragraph, 0, 3), {"color": "auto"})

    assert paragraph[0].find(W_RPR).find(qn("w:color")).get(W_VAL) == "auto"


# --------------------------------------------------------------------------------------
# Character styles
# --------------------------------------------------------------------------------------


def test_a_character_style_the_package_defines_is_written() -> None:
    package = DocxPackage.open(fixture_bytes("character_styles"))
    _, paragraph = paragraph_starting_with(package, "Plain, then ")

    apply_rpr(resolve(paragraph, 0, 5), {"char_style": "FixtureEmphasis"}, pkg=package)

    assert runs(paragraph)[0].find(W_RPR).find(qn("w:rStyle")).get(W_VAL) == "FixtureEmphasis"


def test_an_unknown_character_style_is_refused() -> None:
    package = DocxPackage.open(fixture_bytes("character_styles"))
    _, paragraph = paragraph_starting_with(package, "Plain, then ")
    # Taken after the split: ``resolve`` splits, ``apply_rpr`` is what must not
    # write anything once the style has been refused.
    pieces = resolve(paragraph, 0, 5)
    before = etree.tostring(paragraph)

    with pytest.raises(LocatorError) as refusal:
        apply_rpr(pieces, {"char_style": "NoSuchStyle"}, pkg=package)

    assert refusal.value.code == "unknown-style"
    assert etree.tostring(paragraph) == before


def test_a_paragraph_style_is_not_a_character_style() -> None:
    package = DocxPackage.open(fixture_bytes("character_styles"))
    _, paragraph = paragraph_starting_with(package, "Plain, then ")

    with pytest.raises(LocatorError) as refusal:
        apply_rpr(resolve(paragraph, 0, 5), {"char_style": "Normal"}, pkg=package)

    assert refusal.value.code == "wrong-style-type"


def test_a_character_style_cannot_be_checked_without_the_package() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    with pytest.raises(ValueError):
        apply_rpr(resolve(paragraph, 0, 3), {"char_style": "FixtureEmphasis"})


def test_removing_a_character_style_needs_no_package() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    apply_rpr(resolve(paragraph, 4, 8), {"char_style": None})

    assert runs(paragraph)[1].find(W_RPR) is None


# --------------------------------------------------------------------------------------
# Envelopes: hyperlink
# --------------------------------------------------------------------------------------


def test_wrapping_moves_the_runs_into_the_hyperlink() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    link = wrap(resolve(paragraph, 4, 8), hyperlink("rId42"))

    assert link.tag == qn("w:hyperlink")
    assert link.get(qn("r:id")) == "rId42"
    assert link.get(qn("w:history")) == "1"
    assert names(paragraph) == ["r", "hyperlink", "r"]
    assert names(link) == ["r"]
    assert visible_text(paragraph) == "one two three"


def test_wrapping_then_unwrapping_restores_the_paragraph() -> None:
    paragraph = paragraph_from(THREE_RUNS)
    before = canonical(paragraph)

    link = wrap(resolve(paragraph, 4, 8), hyperlink(anchor="FixtureAnchor"))
    assert canonical(paragraph) != before

    moved = unwrap(link)

    assert canonical(paragraph) == before
    assert [etree.QName(element).localname for element in moved] == ["r"]


def test_wrapping_takes_the_markers_that_sit_inside_the_range_along() -> None:
    paragraph = paragraph_from(
        '<w:r><w:t xml:space="preserve">one </w:t></w:r>'
        '<w:bookmarkStart w:id="9" w:name="Inner"/>'
        "<w:r><w:t>two</w:t></w:r>"
    )

    link = wrap(resolve(paragraph, 0, 7), hyperlink("rId7"))

    assert names(link) == ["r", "bookmarkStart", "r"]
    assert marker_ids(paragraph, "w:bookmarkStart") == ["9"]


def test_wrapping_refuses_a_range_that_crosses_a_container() -> None:
    paragraph = paragraph_from(
        '<w:r><w:t xml:space="preserve">one </w:t></w:r>'
        '<w:hyperlink r:id="rId1"><w:r><w:t>two</w:t></w:r></w:hyperlink>'
    )
    before = visible_text(paragraph)

    with pytest.raises(UnsupportedRange) as refusal:
        wrap(resolve(paragraph, 2, 6), hyperlink("rId2"))

    assert refusal.value.reason == "crosses-container"
    assert visible_text(paragraph) == before


def test_wrapping_refuses_a_hyperlink_inside_a_hyperlink() -> None:
    paragraph = paragraph_from(
        '<w:hyperlink r:id="rId1"><w:r><w:t>linked text</w:t></w:r></w:hyperlink>'
    )

    with pytest.raises(UnsupportedRange) as refusal:
        wrap(resolve(paragraph, 0, 6), hyperlink("rId2"))

    assert refusal.value.reason == "nested-hyperlink"


def test_wrapping_refuses_an_empty_range() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    with pytest.raises(UnsupportedRange) as refusal:
        wrap(resolve(paragraph, 4, 4), hyperlink("rId1"))

    assert refusal.value.reason == "empty-range"


def test_a_hyperlink_needs_a_target() -> None:
    with pytest.raises(ValueError):
        hyperlink()


# --------------------------------------------------------------------------------------
# Envelopes: comment
# --------------------------------------------------------------------------------------


def test_a_comment_frames_the_range_with_its_three_marks() -> None:
    paragraph = paragraph_from(THREE_RUNS)

    start = wrap(resolve(paragraph, 4, 8), comment(5))

    assert start.tag == qn("w:commentRangeStart")
    assert names(paragraph) == ["r", "commentRangeStart", "r", "commentRangeEnd", "r", "r"]
    assert marker_ids(paragraph, "w:commentRangeStart") == ["5"]
    assert marker_ids(paragraph, "w:commentRangeEnd") == ["5"]
    assert marker_ids(paragraph, "w:commentReference") == ["5"]
    assert visible_text(paragraph) == "one two three"


def test_the_reference_takes_the_comment_style_when_the_package_defines_it() -> None:
    package = DocxPackage.open(fixture_bytes("comments"))
    _, paragraph = paragraph_starting_with(package, "This sentence carries")

    wrap(resolve(paragraph, 0, 4), comment(9, pkg=package))
    reference = next(paragraph.iter(qn("w:commentReference")))
    run = reference.getparent()

    assert run.find(W_RPR).find(qn("w:rStyle")).get(W_VAL) == "CommentReference"


def test_the_reference_stays_unstyled_when_the_package_lacks_the_style() -> None:
    package = DocxPackage.open(fixture_bytes("simple"))
    _, paragraph = paragraph_starting_with(package, "First paragraph")

    wrap(resolve(paragraph, 0, 5), comment(1, pkg=package))
    reference = next(paragraph.iter(qn("w:commentReference")))

    assert reference.getparent().find(W_RPR) is None


def test_a_comment_can_be_anchored_inside_a_hyperlink() -> None:
    paragraph = paragraph_from(
        '<w:hyperlink r:id="rId1"><w:r><w:t>linked text</w:t></w:r></w:hyperlink>'
    )

    wrap(resolve(paragraph, 0, 6), comment(3))

    assert names(paragraph[0]) == ["commentRangeStart", "r", "commentRangeEnd", "r", "r"]
    assert visible_text(paragraph) == "linked text"


def test_a_comment_id_must_be_a_number() -> None:
    with pytest.raises(ValueError):
        comment("first")


# --------------------------------------------------------------------------------------
# Envelopes: unwrap
# --------------------------------------------------------------------------------------


def test_unwrapping_a_content_control_keeps_its_runs() -> None:
    paragraph = paragraph_from(
        "<w:sdt>"
        '<w:sdtPr><w:alias w:val="Fixture"/><w:id w:val="1"/></w:sdtPr>'
        "<w:sdtContent><w:r><w:t>held text</w:t></w:r></w:sdtContent>"
        "</w:sdt>"
    )

    moved = unwrap(paragraph[0])

    assert names(paragraph) == ["r"]
    assert moved == (paragraph[0],)
    assert visible_text(paragraph) == "held text"


def test_unwrapping_an_insertion_keeps_its_runs() -> None:
    paragraph = paragraph_from(
        '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:t>inserted</w:t></w:r></w:ins>"
    )

    unwrap(paragraph[0])

    assert names(paragraph) == ["r"]
    assert visible_text(paragraph) == "inserted"


def test_unwrapping_a_deletion_is_refused() -> None:
    paragraph = paragraph_from(
        '<w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:delText>gone</w:delText></w:r></w:del>"
    )
    before = etree.tostring(paragraph)

    with pytest.raises(UnsupportedRange) as refusal:
        unwrap(paragraph[0])

    assert refusal.value.reason == "unsupported-envelope"
    assert etree.tostring(paragraph) == before


def test_unwrapping_a_run_is_refused() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    with pytest.raises(UnsupportedRange):
        unwrap(paragraph[0])


# --------------------------------------------------------------------------------------
# Paragraph properties
# --------------------------------------------------------------------------------------


def test_the_ppr_is_created_as_the_first_child_of_the_paragraph() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    element = set_ppr_child(paragraph, "w:jc", {"w:val": "center"})

    assert names(paragraph) == ["pPr", "r"]
    assert element.get(W_VAL) == "center"


def test_a_ppr_child_lands_at_its_schema_rank() -> None:
    paragraph = paragraph_from('<w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>abc</w:t></w:r>')

    set_ppr_child(paragraph, "w:numPr", {})
    set_ppr_child(paragraph, "w:pStyle", {"w:val": "FixtureBody"})
    set_ppr_child(paragraph, "w:spacing", {"w:before": "120"})

    assert names(paragraph.find(W_PPR)) == ["pStyle", "numPr", "spacing", "jc"]


def test_setting_a_child_again_replaces_its_attributes_and_keeps_its_children() -> None:
    paragraph = paragraph_from(
        "<w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"3\"/></w:numPr></w:pPr>"
        "<w:r><w:t>abc</w:t></w:r>"
    )

    element = set_ppr_child(paragraph, "w:numPr", {})
    assert names(element) == ["ilvl", "numId"]

    spacing = set_ppr_child(paragraph, "w:spacing", {"w:before": "120", "w:after": "120"})
    spacing = set_ppr_child(paragraph, "w:spacing", {"w:after": "240"})
    assert dict(spacing.attrib) == {qn("w:after"): "240"}


def test_setting_then_removing_a_child_leaves_the_paragraph_as_it_was() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")
    before = etree.tostring(paragraph)

    set_ppr_child(paragraph, "w:keepNext", {})
    assert names(paragraph) == ["pPr", "r"]

    assert set_ppr_child(paragraph, "w:keepNext", None) is None
    assert etree.tostring(paragraph) == before


def test_removing_a_child_keeps_a_ppr_that_still_says_something() -> None:
    paragraph = paragraph_from(
        '<w:pPr><w:pStyle w:val="Quote"/><w:jc w:val="both"/></w:pPr><w:r><w:t>abc</w:t></w:r>'
    )

    set_ppr_child(paragraph, "w:jc", None)

    assert names(paragraph.find(W_PPR)) == ["pStyle"]


def test_removing_a_child_from_a_paragraph_without_ppr_is_a_no_op() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")
    before = etree.tostring(paragraph)

    assert set_ppr_child(paragraph, "w:jc", None) is None
    assert etree.tostring(paragraph) == before


def test_a_tag_the_ppr_cannot_hold_is_refused() -> None:
    paragraph = paragraph_from("<w:r><w:t>abc</w:t></w:r>")

    with pytest.raises(ValueError):
        set_ppr_child(paragraph, "w:b", {})
    assert paragraph.find(W_PPR) is None


# --------------------------------------------------------------------------------------
# At package scale
# --------------------------------------------------------------------------------------


def test_formatting_and_linking_disturb_nothing_else_in_the_package(tmp_path: Path) -> None:
    """Two edits on a real package: what moves must be the two paragraphs only."""
    name = "character_styles"
    package = DocxPackage.open(fixture_bytes(name))
    story, root = package.stories()[0]

    formatted_index, formatted = paragraph_starting_with(package, "Plain, then ")
    linked_index, linked = paragraph_starting_with(package, "Style plus direct")

    apply_rpr(
        resolve(formatted, 0, 5),
        {"bold": True, "size_pt": 11, "color": "#0563C1", "char_style": "FixtureCode"},
        pkg=package,
    )
    rId = package.add_external_rel(root, "https://example.org/format")
    wrap(resolve(linked, 0, 5), hyperlink(rId))

    path = package.save(tmp_path / f"{name}.docx")
    assert validate_package(path) == []

    after = snapshot(path)
    assert after.text(story)[formatted_index].startswith("Plain, then ")
    assert_unchanged_except(
        snapshot(untouched_bytes(name)),
        snapshot(path),
        paragraphs=[(story, formatted_index), (story, linked_index)],
        # The link needs a relationship, which lives in a part of its own.
        parts=["word/_rels/document.xml.rels"],
        # Wrapping a range in a hyperlink is exactly what this counter counts.
        counters={"hyperlinks"},
    )
