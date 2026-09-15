"""Tests for :mod:`word_document_server.engine.styles`.

The style sheet is read, never written, so what is pinned here is what a reader
is entitled to believe.

*A level says only what it sets.*
    The whole point of ``chain`` is to answer "where does this bold come from".
    That only works if an inherited property is absent from the level that
    inherits it -- so the absence of a key is asserted as carefully as its
    value.

*``resolved`` applies the levels in Word's order.*
    Defaults, then the ``basedOn`` chain from its root down, then the style.
    And it merges ``w:ind``, ``w:spacing`` and ``w:numPr`` attribute by
    attribute, the way Word inherits them: a style setting only "space after"
    must not silently drop the line spacing it inherits.

*A usage carries a locator that actually resolves.*
    Every locator reported here is fed back through
    :func:`~word_document_server.engine.locators.resolve` and checked to land on
    the very element the usage was found on.  A report of where a style is used
    is worthless if the addresses in it are one off.

Malformed style sheets -- a ``basedOn`` pointing nowhere, a chain that loops --
are built by editing the live styles part of an opened package rather than by
adding a fixture: they are properties of this reader, not documents anyone else
needs.
"""

from __future__ import annotations

import io
import json
import re
import zipfile

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.libreoffice import fodt_to_docx, requires_libreoffice
from word_document_server.engine.errors import PackageError
from word_document_server.engine.locators import resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.styles import (
    STYLE_FAMILIES,
    STYLES_PARTNAME,
    USAGE_PREVIEW,
    StyleUsage,
    decode_ppr,
    decode_rpr,
    find_style_usage,
    get_style,
    list_styles,
)
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import NAMESPACES, qn

W_STYLE = qn("w:style")
W_STYLE_ID = qn("w:styleId")
W_VAL = qn("w:val")
W_PPR = qn("w:pPr")
W_PSTYLE = qn("w:pStyle")


def _pkg(name: str) -> DocxPackage:
    return DocxPackage.open(build(name))


def _fragment(xml: str) -> etree._Element:
    """Parse a ``w:``-prefixed fragment, declaring the namespaces it uses.

    The declarations go right after the element *name*, so that a self-closing
    tag (``<w:pStyle w:val="Quote"/>``) does not get them appended after its
    slash.
    """
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in NAMESPACES.items())
    stripped = xml.strip()
    match = re.match(r"<[A-Za-z0-9_:.-]+", stripped)
    assert match is not None, f"not a tag: {stripped[:40]!r}"
    return etree.fromstring(
        f"{stripped[: match.end()]} {declarations}{stripped[match.end() :]}"
    )


def _styles_element(pkg: DocxPackage) -> etree._Element:
    return pkg.root_of(pkg.part(STYLES_PARTNAME))


def _style_element(pkg: DocxPackage, style_id: str) -> etree._Element:
    for style in _styles_element(pkg).findall(W_STYLE):
        if style.get(W_STYLE_ID) == style_id:
            return style
    raise AssertionError(f"the fixture has no style {style_id!r}")


def _without_styles(blob: bytes) -> bytes:
    """`blob` with the styles part, its content-type override and its rel removed."""
    source = zipfile.ZipFile(io.BytesIO(blob))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            if name == "word/styles.xml":
                continue
            data = source.read(name)
            if name == "[Content_Types].xml":
                data = re.sub(rb'<Override PartName="/word/styles\.xml"[^>]*/>', b"", data)
            if name == "word/_rels/document.xml.rels":
                data = re.sub(rb'<Relationship[^>]*styles\.xml"[^>]*/>', b"", data)
            target.writestr(name, data)
    return buffer.getvalue()


def _info(pkg: DocxPackage, style_id: str):
    for info in list_styles(pkg):
        if info.style_id == style_id:
            return info
    raise AssertionError(f"list_styles did not report {style_id!r}")


# --------------------------------------------------------------------------------------
# list_styles
# --------------------------------------------------------------------------------------


def test_every_style_of_the_part_is_listed_in_document_order() -> None:
    pkg = _pkg("style_inheritance")
    listed = [info.style_id for info in list_styles(pkg)]
    in_part = [
        style.get(W_STYLE_ID) for style in _styles_element(pkg).findall(W_STYLE)
    ]
    assert listed == in_part
    assert listed[:1] == ["Normal"]


def test_the_families_partition_the_style_sheet() -> None:
    pkg = _pkg("style_inheritance")
    total = len(list_styles(pkg))
    assert total == sum(len(list_styles(pkg, family)) for family in STYLE_FAMILIES)
    assert all(
        info.family == "character" for info in list_styles(pkg, "character")
    )


def test_an_unknown_family_is_refused_rather_than_matching_nothing() -> None:
    # Returning [] for a typo would read as "this document defines no character
    # style", which is a different answer from "there is no such family".
    with pytest.raises(ValueError, match="unknown style family"):
        list_styles(_pkg("simple"), "charater")


def test_the_identity_and_wiring_of_a_style_are_reported() -> None:
    info = _info(_pkg("style_inheritance"), "FixtureBranch")
    assert info.name == "Fixture Branch"
    assert info.family == "paragraph"
    assert info.based_on == "FixtureRoot"
    assert info.next_style == "FixtureRoot"
    assert info.link == "FixtureBranchChar"


def test_a_style_without_wiring_reports_none_not_its_own_id() -> None:
    info = _info(_pkg("style_inheritance"), "Normal")
    assert info.based_on is None
    assert info.link is None


def test_builtin_reads_the_custom_style_flag_and_does_not_judge_the_name() -> None:
    pkg = _pkg("style_inheritance")
    # FixtureBranchChar declares w:customStyle="1"; FixtureBranch, just as
    # hand-written, does not.  The report follows the document, not the name.
    assert _info(pkg, "FixtureBranchChar").builtin is False
    assert _info(pkg, "FixtureBranch").builtin is True
    assert _info(pkg, "Normal").builtin is True


def test_the_default_style_of_a_family_is_flagged() -> None:
    pkg = _pkg("style_inheritance")
    assert _info(pkg, "Normal").default is True
    assert _info(pkg, "FixtureBranch").default is False


def test_the_three_visibility_flags_and_the_priority_are_reported() -> None:
    pkg = _pkg("style_inheritance")
    assert _info(pkg, "FixtureRoot").q_format is True
    assert _info(pkg, "FixtureBranch").q_format is False
    heading = _info(pkg, "Heading1")
    assert heading.q_format is True
    assert isinstance(heading.ui_priority, int)
    assert _info(pkg, "FixtureRoot").ui_priority is None


def test_a_flag_explicitly_turned_off_is_off() -> None:
    # <w:qFormat w:val="0"/> is not the same as a bare <w:qFormat/>, and a
    # reader that only checks for the element's presence gets it backwards.
    pkg = _pkg("style_inheritance")
    style = _style_element(pkg, "FixtureRoot")
    style.find(qn("w:qFormat")).set(W_VAL, "0")
    assert _info(pkg, "FixtureRoot").q_format is False


def test_a_style_without_a_name_falls_back_to_its_id() -> None:
    pkg = _pkg("style_inheritance")
    style = _style_element(pkg, "FixtureRoot")
    style.remove(style.find(qn("w:name")))
    assert _info(pkg, "FixtureRoot").name == "FixtureRoot"


def test_a_style_without_an_id_is_skipped_because_nothing_can_reference_it() -> None:
    pkg = _pkg("style_inheritance")
    _styles_element(pkg).append(
        _fragment('<w:style w:type="paragraph"><w:name w:val="Nameless"/></w:style>')
    )
    assert all(info.style_id for info in list_styles(pkg))
    assert "Nameless" not in {info.name for info in list_styles(pkg)}


def test_a_package_without_a_style_sheet_reads_as_defining_no_style() -> None:
    # Not an error: a document may lean entirely on Word's own defaults.
    pkg = DocxPackage.open(_without_styles(build("simple")))
    assert pkg.find_part(STYLES_PARTNAME) is None
    assert list_styles(pkg) == []
    with pytest.raises(PackageError, match="no style with id or name"):
        get_style(pkg, "Heading1")


def test_usage_is_read_from_the_body_not_from_the_style_sheet() -> None:
    # find_style_usage reports where an id is *referenced*; it does not check
    # that the style sheet defines it.  A document whose styles part is gone
    # still has paragraphs pointing at Heading1, and hiding them would make the
    # function's answer depend on a part it never reads.
    pkg = DocxPackage.open(_without_styles(build("simple")))
    usages = find_style_usage(pkg, "Heading1")
    assert [usage.kind for usage in usages] == ["paragraph"]
    assert usages[0].text == "Fixture: simple"


# --------------------------------------------------------------------------------------
# get_style: the chain
# --------------------------------------------------------------------------------------


def test_a_style_is_found_by_id_and_by_name() -> None:
    pkg = _pkg("style_inheritance")
    assert get_style(pkg, "FixtureLeaf").info.style_id == "FixtureLeaf"
    assert get_style(pkg, "Fixture Leaf").info.style_id == "FixtureLeaf"


def test_the_id_wins_over_the_name() -> None:
    pkg = _pkg("style_inheritance")
    # Give FixtureRoot the *name* "FixtureLeaf": a w:pStyle saying FixtureLeaf
    # still means the style whose id is FixtureLeaf, so this lookup must too.
    _style_element(pkg, "FixtureRoot").find(qn("w:name")).set(W_VAL, "FixtureLeaf")
    assert get_style(pkg, "FixtureLeaf").info.style_id == "FixtureLeaf"


def test_the_chain_runs_from_the_style_to_the_document_defaults() -> None:
    detail = get_style(_pkg("style_inheritance"), "FixtureLeaf")
    assert [level.style_id for level in detail.chain] == [
        "FixtureLeaf",
        "FixtureBranch",
        "FixtureRoot",
        "Normal",
        "",
    ]
    assert detail.chain[-1].name == "Document defaults"
    assert detail.warnings == ()


def test_a_level_reports_only_what_it_sets() -> None:
    detail = get_style(_pkg("style_inheritance"), "FixtureLeaf")
    levels = {level.style_id: level for level in detail.chain}
    assert levels["FixtureLeaf"].run_props == {"italic": True}
    assert levels["FixtureBranch"].run_props == {"bold": True}
    assert levels["FixtureRoot"].run_props == {"size_pt": 12.0}
    # The inherited bold is *not* on the leaf: that is what makes the chain
    # readable as "where does this come from".
    assert "bold" not in levels["FixtureLeaf"].run_props
    assert "italic" not in levels["FixtureBranch"].run_props


def test_the_style_s_own_properties_are_the_first_level() -> None:
    detail = get_style(_pkg("style_inheritance"), "FixtureLeaf")
    assert detail.run_props == detail.chain[0].run_props
    assert detail.paragraph_props == detail.chain[0].paragraph_props
    assert detail.paragraph_props == {"indent": {"left": 720}}


def test_an_unknown_style_names_a_few_of_the_ones_that_exist() -> None:
    with pytest.raises(PackageError, match="no style with id or name 'Nope'") as caught:
        get_style(_pkg("style_inheritance"), "Nope")
    assert "Normal" in str(caught.value) or "BodyText" in str(caught.value)


@pytest.mark.parametrize("bad", ["", "   "])
def test_an_empty_style_lookup_is_refused(bad: str) -> None:
    with pytest.raises(ValueError, match="non-empty"):
        get_style(_pkg("simple"), bad)


def test_a_based_on_pointing_nowhere_stops_the_chain_and_warns() -> None:
    pkg = _pkg("style_inheritance")
    _style_element(pkg, "FixtureRoot").find(qn("w:basedOn")).set(W_VAL, "Ghost")
    detail = get_style(pkg, "FixtureLeaf")
    assert [level.style_id for level in detail.chain] == [
        "FixtureLeaf",
        "FixtureBranch",
        "FixtureRoot",
        "",
    ]
    assert any("Ghost" in warning for warning in detail.warnings)


def test_a_based_on_chain_that_loops_stops_instead_of_recursing() -> None:
    pkg = _pkg("style_inheritance")
    # FixtureRoot -> FixtureLeaf -> FixtureBranch -> FixtureRoot: legal to
    # write, fatal to follow.
    _style_element(pkg, "FixtureRoot").find(qn("w:basedOn")).set(W_VAL, "FixtureLeaf")
    detail = get_style(pkg, "FixtureLeaf")
    visited = [level.style_id for level in detail.chain]
    assert visited == ["FixtureLeaf", "FixtureBranch", "FixtureRoot", ""]
    assert len(visited) == len(set(visited))
    assert any("loops" in warning for warning in detail.warnings)


def test_a_style_that_is_its_own_parent_terminates() -> None:
    pkg = _pkg("style_inheritance")
    _style_element(pkg, "FixtureRoot").find(qn("w:basedOn")).set(W_VAL, "FixtureRoot")
    detail = get_style(pkg, "FixtureRoot")
    assert [level.style_id for level in detail.chain] == ["FixtureRoot", ""]
    assert detail.warnings


# --------------------------------------------------------------------------------------
# get_style: resolution
# --------------------------------------------------------------------------------------


def test_resolution_applies_the_chain_from_the_defaults_up() -> None:
    resolved = get_style(_pkg("style_inheritance"), "FixtureLeaf").resolved
    assert resolved["run"]["italic"] is True  # from the leaf
    assert resolved["run"]["bold"] is True  # inherited from the branch
    assert resolved["run"]["size_pt"] == 12.0  # from the root, over the defaults' 11.0
    assert resolved["paragraph"]["indent"] == {"left": 720}  # leaf over branch's 360


def test_the_nearest_level_wins() -> None:
    pkg = _pkg("style_inheritance")
    # The root says 12 pt, the defaults 11 pt; make the leaf say 20 and it wins.
    _style_element(pkg, "FixtureLeaf").find(qn("w:rPr")).append(
        _fragment('<w:sz w:val="40"/>')
    )
    assert get_style(pkg, "FixtureLeaf").resolved["run"]["size_pt"] == 20.0


def test_composite_paragraph_properties_are_inherited_attribute_by_attribute() -> None:
    # The document defaults say <w:spacing w:after="200" w:line="276"
    # w:lineRule="auto"/> and FixtureRoot says only w:after.  Word keeps the
    # inherited line spacing; replacing the element wholesale would lose it.
    resolved = get_style(_pkg("style_inheritance"), "FixtureLeaf").resolved
    assert resolved["paragraph"]["spacing"] == {
        "after": 200,
        "line": 276,
        "line_rule": "auto",
    }


def test_a_font_that_follows_the_theme_keeps_the_reference() -> None:
    resolved = get_style(_pkg("style_inheritance"), "FixtureLeaf").resolved
    # The document defaults follow the theme and cache no literal: flattening
    # the two halves would leave a caller unable to tell which is which.
    assert resolved["run"]["font"] == {"value": None, "theme": "minorHAnsi"}


def test_the_resolved_view_is_plain_json_ready_data() -> None:
    detail = get_style(_pkg("combined"), "Heading1")
    payload = {"resolved": detail.resolved, "chain": [level.run_props for level in detail.chain]}
    assert json.loads(json.dumps(payload)) == payload


# --------------------------------------------------------------------------------------
# Decoding
# --------------------------------------------------------------------------------------


def test_a_character_style_decodes_its_run_properties() -> None:
    detail = get_style(_pkg("character_styles"), "FixtureEmphasis")
    assert detail.info.family == "character"
    assert detail.run_props == {
        "italic": True,
        "color": {"value": "C00000", "theme": None, "tint": None, "shade": None},
    }


def test_a_pinned_font_is_reported_with_no_theme_reference() -> None:
    detail = get_style(_pkg("character_styles"), "FixtureCode")
    assert detail.run_props["font"] == {"value": "Consolas", "theme": None}


def test_an_absent_toggle_is_absent_and_an_explicit_off_is_false() -> None:
    # Three states, not two: "inherit", "on" and "cancel what I inherit".
    assert decode_rpr(_fragment("<w:rPr/>")) == {}
    assert decode_rpr(_fragment("<w:rPr><w:b/></w:rPr>"))["bold"] is True
    assert decode_rpr(_fragment('<w:rPr><w:b w:val="0"/></w:rPr>'))["bold"] is False
    assert decode_rpr(_fragment('<w:rPr><w:b w:val="false"/></w:rPr>'))["bold"] is False


def test_a_size_is_reported_in_points_from_the_half_points_stored() -> None:
    assert decode_rpr(_fragment('<w:rPr><w:sz w:val="23"/></w:rPr>'))["size_pt"] == 11.5


def test_a_themed_colour_keeps_its_tint() -> None:
    decoded = decode_rpr(
        _fragment(
            '<w:rPr><w:color w:val="B4C6E7" w:themeColor="accent1" w:themeTint="66"/></w:rPr>'
        )
    )
    assert decoded["color"] == {
        "value": "B4C6E7",
        "theme": "accent1",
        "tint": "66",
        "shade": None,
    }


def test_vertical_alignment_decodes_into_the_two_keys_the_patch_uses() -> None:
    superscript = decode_rpr(_fragment('<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'))
    assert superscript["superscript"] is True
    assert superscript["subscript"] is False
    baseline = decode_rpr(_fragment('<w:rPr><w:vertAlign w:val="baseline"/></w:rPr>'))
    assert baseline == {"superscript": False, "subscript": False}


def test_the_three_font_slots_are_decoded_separately() -> None:
    decoded = decode_rpr(
        _fragment(
            '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" '
            'w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/></w:rPr>'
        )
    )
    assert decoded["font"] == {"value": None, "theme": "majorHAnsi"}
    assert decoded["font_east_asia"] == {"value": None, "theme": "majorEastAsia"}
    assert decoded["font_cs"] == {"value": None, "theme": "majorBidi"}


def test_hansi_alone_still_reads_as_the_latin_font() -> None:
    decoded = decode_rpr(_fragment('<w:rPr><w:rFonts w:hAnsi="Consolas"/></w:rPr>'))
    assert decoded["font"] == {"value": "Consolas", "theme": None}


def test_paragraph_measurements_stay_in_twips() -> None:
    decoded = decode_ppr(
        _fragment(
            '<w:pPr><w:jc w:val="both"/><w:ind w:left="720" w:hanging="360"/>'
            '<w:spacing w:before="120" w:after="240" w:lineRule="auto"/>'
            "<w:keepNext/><w:outlineLvl w:val=\"2\"/></w:pPr>"
        )
    )
    assert decoded["alignment"] == "both"
    assert decoded["indent"] == {"left": 720, "hanging": 360}
    assert decoded["spacing"] == {"before": 120, "after": 240, "line_rule": "auto"}
    assert decoded["keep_next"] is True
    assert decoded["outline_level"] == 2


def test_the_start_and_end_spellings_of_an_indent_are_read() -> None:
    # Word writes w:left/w:right; the strict-conformance spelling is
    # w:start/w:end, and a document may carry either.
    decoded = decode_ppr(_fragment('<w:pPr><w:ind w:start="360" w:end="180"/></w:pPr>'))
    assert decoded["indent"] == {"left": 360, "right": 180}


def test_the_numbering_link_reports_only_what_it_states() -> None:
    both = decode_ppr(
        _fragment('<w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr>')
    )
    assert both["numbering"] == {"num_id": 3, "level": 1}
    # w:numId and w:ilvl are separate children and are inherited separately, so
    # a level that states one must not blank the other.
    partial = decode_ppr(_fragment('<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>'))
    assert partial["numbering"] == {"num_id": 3}


def test_the_numbering_link_is_inherited_attribute_by_attribute() -> None:
    pkg = _pkg("style_inheritance")
    _style_element(pkg, "FixtureRoot").find(qn("w:pPr")).append(
        _fragment('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>')
    )
    _style_element(pkg, "FixtureLeaf").find(qn("w:pPr")).append(
        _fragment('<w:numPr><w:ilvl w:val="2"/></w:numPr>')
    )
    resolved = get_style(pkg, "FixtureLeaf").resolved["paragraph"]
    assert resolved["numbering"] == {"num_id": 7, "level": 2}


def test_a_property_outside_the_minimal_model_is_not_invented() -> None:
    # w:bdr is legal and unread here.  Reporting nothing is what makes a
    # decoded set a reading of the style rather than a replacement for the XML.
    decoded = decode_rpr(_fragment('<w:rPr><w:bdr w:val="single"/><w:b/></w:rPr>'))
    assert decoded == {"bold": True}


# --------------------------------------------------------------------------------------
# find_style_usage
# --------------------------------------------------------------------------------------


def _resolves_to(pkg: DocxPackage, usage: StyleUsage) -> etree._Element:
    """Resolve a usage's locator and return the paragraph it lands on."""
    assert usage.locator is not None
    return resolve(pkg, usage.locator).paragraph


def test_a_paragraph_usage_carries_a_locator_that_lands_on_it() -> None:
    pkg = _pkg("style_inheritance")
    usages = find_style_usage(pkg, "FixtureRoot")
    assert [usage.kind for usage in usages] == ["paragraph"]
    usage = usages[0]
    assert usage.story == "document"
    assert usage.text == "Root of the chain."
    assert visible_text(_resolves_to(pkg, usage)) == usage.text


def test_a_run_usage_carries_the_offsets_of_the_run_in_its_paragraph() -> None:
    pkg = _pkg("character_styles")
    usages = find_style_usage(pkg, "FixtureEmphasis")
    assert [usage.kind for usage in usages] == ["run", "run"]
    first = usages[0]
    paragraph = _resolves_to(pkg, first)
    # The offsets are in the visible text of the located paragraph, which is
    # exactly what doc_format_range takes back.
    assert visible_text(paragraph)[first.start : first.end] == "emphasised"
    assert first.text == "emphasised"


def test_inheritance_is_not_followed() -> None:
    # FixtureLeaf is based on FixtureBranch, which is based on FixtureRoot.
    # A paragraph in FixtureLeaf is a usage of FixtureLeaf, not of FixtureRoot.
    pkg = _pkg("style_inheritance")
    assert len(find_style_usage(pkg, "FixtureRoot")) == 1
    assert len(find_style_usage(pkg, "FixtureLeaf")) == 1


def test_a_style_name_finds_nothing_because_the_document_references_ids() -> None:
    assert find_style_usage(_pkg("style_inheritance"), "Fixture Root") == []


def test_an_unused_or_undefined_style_is_simply_empty() -> None:
    pkg = _pkg("simple")
    assert find_style_usage(pkg, "NoSuchStyleAnywhere") == []


def test_an_empty_style_id_is_refused() -> None:
    with pytest.raises(ValueError, match="non-empty"):
        find_style_usage(_pkg("simple"), "")


def test_a_table_usage_reports_the_table_index_and_no_paragraph_index() -> None:
    pkg = _pkg("tables")
    usages = find_style_usage(pkg, "TableGrid")
    assert [usage.kind for usage in usages] == ["table", "table"]
    assert [usage.table for usage in usages] == [0, 1]
    assert all(usage.index is None for usage in usages)
    # The table address, not a fabricated cell: resolving it as it stands fails
    # loudly rather than silently hitting some arbitrary paragraph.
    assert usages[0].locator == {"story": "document", "table": 0}


def test_the_table_address_becomes_a_locator_once_a_cell_is_named() -> None:
    pkg = _pkg("tables")
    usage = find_style_usage(pkg, "TableGrid")[0]
    assert usage.locator is not None
    target = resolve(pkg, {**usage.locator, "row": 0, "col": 0})
    assert target.story == "document"
    assert target.paragraph.tag == qn("w:p")


def test_a_usage_in_a_table_cell_is_addressed_by_the_table_form() -> None:
    pkg = _pkg("tables")
    cell_paragraph = next(
        paragraph
        for paragraph in pkg.stories()[0][1].iter(qn("w:p"))
        if any(ancestor.tag == qn("w:tc") for ancestor in paragraph.iterancestors())
    )
    properties = cell_paragraph.find(W_PPR)
    if properties is None:
        properties = etree.SubElement(cell_paragraph, W_PPR)
        cell_paragraph.insert(0, properties)
    properties.append(_fragment('<w:pStyle w:val="Quote"/>'))

    usages = find_style_usage(pkg, "Quote")
    assert [usage.kind for usage in usages] == ["paragraph"]
    usage = usages[0]
    # D-016: a cell paragraph has no V2 index, so it reports None rather than a
    # number that would address a different paragraph.
    assert usage.index is None
    assert set(usage.locator) == {"story", "table", "row", "col", "paragraph"}
    assert _resolves_to(pkg, usage) is cell_paragraph


def test_a_usage_in_a_text_box_has_no_locator_at_all() -> None:
    pkg = _pkg("text_boxes")
    box_paragraph = next(
        paragraph
        for paragraph in pkg.stories()[0][1].iter(qn("w:p"))
        if any(
            ancestor.tag == qn("w:txbxContent") for ancestor in paragraph.iterancestors()
        )
    )
    properties = box_paragraph.find(W_PPR)
    if properties is None:
        properties = etree.SubElement(box_paragraph, W_PPR)
        box_paragraph.insert(0, properties)
    properties.append(_fragment('<w:pStyle w:val="Quote"/>'))

    usages = find_style_usage(pkg, "Quote")
    assert [usage.kind for usage in usages] == ["paragraph"]
    # Neither in the V2 index space nor in a cell: there is no locator form for
    # it, and None says so instead of pointing somewhere plausible and wrong.
    assert usages[0].locator is None
    assert usages[0].index is None
    assert usages[0].text == visible_text(box_paragraph)


def test_every_story_is_searched_and_each_usage_names_its_own() -> None:
    pkg = _pkg("combined")
    stories = {usage.story for usage in find_style_usage(pkg, "Header")}
    assert stories
    assert stories <= {name for name, _ in pkg.stories()}
    assert "document" not in stories  # the Header style lives in the header parts


def test_a_usage_locator_resolves_in_the_story_it_names() -> None:
    pkg = _pkg("combined")
    for usage in find_style_usage(pkg, "FootnoteText"):
        target = resolve(pkg, usage.locator)
        assert target.story == usage.story


def test_usages_come_back_in_document_order_within_a_story() -> None:
    pkg = _pkg("combined")
    indices = [
        usage.index
        for usage in find_style_usage(pkg, "Heading1")
        if usage.story == "document" and usage.index is not None
    ]
    assert indices == sorted(indices)
    assert len(indices) > 1


def test_a_long_usage_is_previewed_and_says_so() -> None:
    pkg = _pkg("style_inheritance")
    paragraph = next(
        p
        for p in pkg.stories()[0][1].iter(qn("w:p"))
        if (pr := p.find(W_PPR)) is not None
        and (st := pr.find(W_PSTYLE)) is not None
        and st.get(W_VAL) == "FixtureRoot"
    )
    paragraph.find(qn("w:r")).find(qn("w:t")).text = "x" * (USAGE_PREVIEW + 40)

    usage = find_style_usage(pkg, "FixtureRoot")[0]
    assert len(usage.text) == USAGE_PREVIEW
    assert usage.truncated is True


def test_a_short_usage_is_not_flagged_as_truncated() -> None:
    assert find_style_usage(_pkg("style_inheritance"), "FixtureRoot")[0].truncated is False


def test_finding_usages_does_not_modify_the_document() -> None:
    blob = build("combined")
    pkg = DocxPackage.open(blob)
    before = pkg.to_bytes()
    assert find_style_usage(pkg, "Heading1")
    assert get_style(pkg, "Heading1")
    assert list_styles(pkg)
    after = pkg.to_bytes()

    # Compared part by part: the zip stamps a fresh timestamp on every member,
    # so the archive bytes differ even when nothing was touched.
    opened_before = zipfile.ZipFile(io.BytesIO(before))
    opened_after = zipfile.ZipFile(io.BytesIO(after))
    assert opened_before.namelist() == opened_after.namelist()
    for name in opened_before.namelist():
        assert opened_before.read(name) == opened_after.read(name), name


# --------------------------------------------------------------------------------------
# A style sheet this project did not write
# --------------------------------------------------------------------------------------


@requires_libreoffice
def test_a_libreoffice_style_sheet_reads_and_its_usages_resolve(tmp_path_factory) -> None:
    path = fodt_to_docx("rich", tmp_path_factory.mktemp("soffice-styles"))
    pkg = DocxPackage.open(path)
    styles = list_styles(pkg)
    assert styles

    used = [
        info
        for info in styles
        if info.family == "paragraph" and find_style_usage(pkg, info.style_id)
    ]
    assert used, "the rich fixture applies at least one paragraph style"
    for info in used:
        detail = get_style(pkg, info.style_id)
        assert detail.chain[-1].name == "Document defaults"
        for usage in find_style_usage(pkg, info.style_id):
            if usage.locator is not None and usage.kind != "table":
                assert resolve(pkg, usage.locator).story == usage.story
