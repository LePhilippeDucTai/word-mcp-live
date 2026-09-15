"""Tests for :mod:`word_document_server.engine.effective`.

What is pinned here, and why each of them is worth a test.

*The provenance is the product.*
    An effective value without its source is what ``doc_inspect`` already gives.
    Every layer of the cascade is therefore made to win at least once --
    ``docDefaults``, a table style, each step of a ``basedOn`` chain, a
    character style, direct formatting -- and the reported `source` is asserted,
    not just the value.

*Toggles XOR, and direct formatting does not.*
    ECMA-376 §17.7.3 is the single most counter-intuitive rule in the format
    model: a bold style based on a bold style renders *not* bold.  A reader that
    implements plain override gets the right answer on every one-level document
    and the wrong one on every real template, so the two-level case is tested
    directly, both down a ``basedOn`` chain and across the paragraph/character
    boundary.  The ``sources`` list that explains the flip is pinned too: without
    it the answer reads as a bug.

*"I cannot say" is not "nothing".*
    Three values are answers rather than data -- ``"unresolved"`` for a table
    style, ``"mixed"`` for a range that is not uniform, and an absent key for a
    property no layer sets.  Collapsing any two of them would make an agent
    either edit blindly or give up wrongly, so each is produced on purpose.

*Reading does not write.*
    Asking what a range looks like must not split a run.  The obvious
    implementation goes through :func:`...ranges.resolve`, which does split, and
    the damage is invisible in the answer -- so it is asserted on the package.

*Theme resolution stays a hint (D-028).*
    The slot lookup and the ``w:clrSchemeMapping`` indirection are pinned
    exactly; the tint is pinned by its direction and by the fact that it reached
    the resolver at all, never by a recalibrated constant.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Sequence
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.snapshot import diff, snapshot
from word_document_server.engine.effective import (
    EXTRA_RUN_TOGGLES,
    MIXED,
    TOGGLE_PROPERTIES,
    UNRESOLVED,
    effective_format,
)
from word_document_server.engine.errors import LocatorError
from word_document_server.engine.locators import Target, indexed_paragraphs, resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.styles import STYLES_PARTNAME
from word_document_server.engine.theme import read_theme
from word_document_server.tools.v2.effective import TOOLS, doc_get_effective_format
from word_document_server.tools.v2.registry import discover_tool_specs

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _element(xml: str) -> etree._Element:
    """Parse one WordprocessingML element written without its namespace prologue."""
    return etree.fromstring(f'<wrap xmlns:w="{W}">{xml}</wrap>')[0]


def _package(
    name: str = "simple",
    *,
    styles: Sequence[str] = (),
    body: Sequence[str] = (),
) -> DocxPackage:
    """A fixture package with extra styles and paragraphs grafted into it.

    The graft happens on the live elements of an opened package rather than in
    ``tests/fixtures/builders.py``: the cases below need styles that exist to
    exercise one rule each, and they have no business in a fixture every other
    test pays for.  Nothing is saved, so nothing is committed.
    """
    pkg = DocxPackage.open(build(name))
    if styles:
        root = pkg.root_of(pkg.part(STYLES_PARTNAME))
        for xml in styles:
            root.append(_element(xml))
    if body:
        body_element = pkg.document.find(f"{{{W}}}body")
        section = body_element.find(f"{{{W}}}sectPr")
        for xml in body:
            element = _element(xml)
            if section is None:
                body_element.append(element)
            else:
                section.addprevious(element)
    return pkg


def _report(pkg: DocxPackage, locator: dict[str, Any], **span: int) -> dict[str, Any]:
    """Resolve `locator` and report the effective format of what it names."""
    target = resolve(pkg, locator)
    if span:
        target = dataclasses.replace(target, **span)
    return effective_format(pkg, target)


def _channels(color: str) -> tuple[int, int, int]:
    return tuple(int(color[index : index + 2], 16) for index in (0, 2, 4))  # type: ignore[return-value]


# --------------------------------------------------------------------------------------
# Every layer wins at least once
# --------------------------------------------------------------------------------------


def test_a_basedon_chain_reports_which_step_set_each_property():
    """The interesting answer is not "bold", it is "bold, from FixtureBranch"."""
    pkg = DocxPackage.open(build("style_inheritance"))
    report = _report(pkg, {"find": "Leaf, based on branch."})

    assert report["paragraph_style"] == "FixtureLeaf"
    assert report["char_styles"] == []
    assert report["table_style"] is None
    # italic is the leaf's own, bold comes from its parent, the size from its
    # grandparent: three steps of one chain, each named.
    assert report["run"]["italic"] == {"value": True, "source": "style:FixtureLeaf"}
    assert report["run"]["bold"] == {"value": True, "source": "style:FixtureBranch"}
    assert report["run"]["size_pt"] == {"value": 12.0, "source": "style:FixtureRoot"}


def test_a_character_style_beats_the_paragraph_style_and_direct_beats_both():
    pkg = _package(
        "character_styles",
        styles=[
            """
            <w:style w:type="paragraph" w:styleId="BlueParagraph">
              <w:name w:val="Blue Paragraph"/>
              <w:rPr><w:color w:val="0000FF"/></w:rPr>
            </w:style>
            """
        ],
        body=[
            """
            <w:p>
              <w:pPr><w:pStyle w:val="BlueParagraph"/></w:pPr>
              <w:r><w:t>from the paragraph style</w:t></w:r>
              <w:r><w:rPr><w:rStyle w:val="FixtureEmphasis"/></w:rPr><w:t>from the character style</w:t></w:r>
              <w:r><w:rPr><w:rStyle w:val="FixtureEmphasis"/><w:color w:val="00FF00"/></w:rPr><w:t>from the run</w:t></w:r>
            </w:p>
            """
        ],
    )
    paragraph = _report(pkg, {"find": "from the paragraph style"})
    assert paragraph["char_styles"] == []
    assert paragraph["run"]["color"] == {
        "value": "0000FF",
        "source": "style:BlueParagraph",
    }

    character = _report(pkg, {"find": "from the character style"})
    assert character["char_styles"] == ["FixtureEmphasis"]
    assert character["run"]["color"] == {
        "value": "C00000",
        "source": "style:FixtureEmphasis",
    }

    direct = _report(pkg, {"find": "from the run"})
    assert direct["run"]["color"] == {"value": "00FF00", "source": "direct"}


def test_the_default_paragraph_style_applies_when_the_paragraph_names_none():
    """A paragraph without a ``w:pStyle`` is not unstyled; Word applies the default."""
    pkg = DocxPackage.open(build("simple"))
    report = _report(pkg, {"paragraph": 1})

    assert report["paragraph_style"] == "Normal"
    # docDefaults is reachable only through a paragraph nothing else sets.
    assert report["run"]["size_pt"]["source"] == "docDefaults"


def test_a_property_no_layer_sets_is_absent_rather_than_null():
    """Absent means "nothing sets it", which is not "set to nothing"."""
    report = _report(DocxPackage.open(build("simple")), {"paragraph": 1})

    assert "highlight" not in report["run"]
    assert "alignment" not in report["paragraph"]


def test_a_style_the_document_does_not_define_is_a_warning_not_a_crash():
    pkg = _package(
        body=[
            (
                '<w:p><w:pPr><w:pStyle w:val="NoSuchStyle"/></w:pPr>'
                "<w:r><w:t>Orphan style.</w:t></w:r></w:p>"
            )
        ]
    )
    report = _report(pkg, {"find": "Orphan style."})

    assert report["paragraph_style"] == "NoSuchStyle"
    assert any("NoSuchStyle" in warning for warning in report["warnings"])
    # The rest of the cascade still answers.
    assert report["run"]["size_pt"]["source"] == "docDefaults"


# --------------------------------------------------------------------------------------
# ECMA-376 §17.7.3: toggles
# --------------------------------------------------------------------------------------

_XOR_STYLES = (
    """
    <w:style w:type="paragraph" w:styleId="XorRoot">
      <w:name w:val="Xor Root"/>
      <w:rPr><w:b/><w:vanish/><w:outline/></w:rPr>
    </w:style>
    """,
    """
    <w:style w:type="paragraph" w:styleId="XorLeaf">
      <w:name w:val="Xor Leaf"/>
      <w:basedOn w:val="XorRoot"/>
      <w:rPr><w:b/></w:rPr>
    </w:style>
    """,
    """
    <w:style w:type="character" w:styleId="XorChar">
      <w:name w:val="Xor Char"/>
      <w:rPr><w:b/></w:rPr>
    </w:style>
    """,
)

_XOR_BODY = (
    (
        '<w:p><w:pPr><w:pStyle w:val="XorLeaf"/></w:pPr>'
        "<w:r><w:t>Bold on bold.</w:t></w:r></w:p>"
    ),
    (
        '<w:p><w:pPr><w:pStyle w:val="XorLeaf"/></w:pPr>'
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Bold on bold, forced.</w:t></w:r></w:p>"
    ),
    (
        '<w:p><w:pPr><w:pStyle w:val="XorRoot"/></w:pPr>'
        '<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Bold cancelled.</w:t></w:r></w:p>'
    ),
    (
        '<w:p><w:pPr><w:pStyle w:val="XorRoot"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="XorChar"/></w:rPr><w:t>Bold across families.</w:t></w:r></w:p>'
    ),
)


@pytest.fixture
def xor_package() -> DocxPackage:
    return _package(styles=_XOR_STYLES, body=_XOR_BODY)


def test_a_toggle_stated_twice_down_a_chain_cancels_itself(xor_package):
    """The rule that makes real templates behave unlike anyone's intuition."""
    report = _report(xor_package, {"find": "Bold on bold."})

    assert report["run"]["bold"]["value"] is False
    assert report["run"]["bold"]["source"] == "style:XorLeaf"
    assert report["run"]["bold"]["sources"] == ["style:XorRoot", "style:XorLeaf"]


def test_a_toggle_stated_once_carries_no_sources_list(xor_package):
    """``sources`` explains a flip; it is noise when there was nothing to flip."""
    vanish = _report(xor_package, {"find": "Bold on bold."})["run"]["vanish"]

    assert vanish == {"value": True, "source": "style:XorRoot"}


def test_direct_formatting_is_outside_the_xor_rule(xor_package):
    forced = _report(xor_package, {"find": "Bold on bold, forced."})["run"]["bold"]
    cancelled = _report(xor_package, {"find": "Bold cancelled."})["run"]["bold"]

    assert forced == {"value": True, "source": "direct"}
    assert cancelled == {"value": False, "source": "direct"}


def test_the_xor_crosses_the_paragraph_and_character_style_boundary(xor_package):
    """Both chains feed one XOR; folding them separately would keep the bold."""
    report = _report(xor_package, {"find": "Bold across families."})

    assert report["char_styles"] == ["XorChar"]
    assert report["run"]["bold"]["value"] is False
    assert report["run"]["bold"]["sources"] == ["style:XorRoot", "style:XorChar"]


def test_every_toggle_of_the_section_is_read_including_the_six_outside_the_model():
    """``decode_rpr`` stops at the minimal model; §17.7.3 does not."""
    pkg = _package(
        body=[
            (
                "<w:p><w:r><w:rPr>"
                "<w:b/><w:i/><w:caps/><w:smallCaps/><w:strike/>"
                "<w:dstrike/><w:outline/><w:shadow/><w:emboss/><w:imprint/><w:vanish/>"
                "</w:rPr><w:t>All toggles.</w:t></w:r></w:p>"
            )
        ]
    )
    report = _report(pkg, {"find": "All toggles."})

    assert set(EXTRA_RUN_TOGGLES) <= set(TOGGLE_PROPERTIES)
    for name in TOGGLE_PROPERTIES:
        assert report["run"][name] == {"value": True, "source": "direct"}, name


def test_a_toggle_from_a_style_outside_the_minimal_model_still_reaches_the_report(
    xor_package,
):
    """``w:outline`` lives on a style element ``decode_rpr`` does not decode."""
    outline = _report(xor_package, {"find": "Bold on bold."})["run"]["outline"]

    assert outline == {"value": True, "source": "style:XorRoot"}


@pytest.mark.parametrize(
    ("attr", "expected"),
    [
        ('w:val="0"', False),
        ('w:val="false"', False),
        ('w:val="off"', False),
        ('w:val="1"', True),
        ('w:val="true"', True),
        ('w:val="on"', True),
        ("", True),
    ],
)
def test_the_extra_toggles_read_off_the_way_the_model_toggles_do(attr, expected):
    """§17.7.3's own toggles (``vanish``, outside the minimal model) must decode
    exactly as the model's own toggles (``bold``) do: both now go through
    :func:`~word_document_server.engine.styles._toggle_value`, so there is one
    three-state vocabulary rather than two lists maintained separately."""
    pkg = _package(
        body=[
            (
                f"<w:p><w:r><w:rPr><w:b {attr}/><w:vanish {attr}/></w:rPr>"
                "<w:t>Toggle probe.</w:t></w:r></w:p>"
            )
        ]
    )
    report = _report(pkg, {"find": "Toggle probe."})

    assert report["run"]["bold"] == {"value": expected, "source": "direct"}
    assert report["run"]["vanish"] == {"value": expected, "source": "direct"}


# --------------------------------------------------------------------------------------
# Paragraph properties, inherited one setting at a time
# --------------------------------------------------------------------------------------


def test_a_multipart_property_keeps_one_source_per_key():
    """``indent.left`` may come from the style while ``indent.hanging`` is direct."""
    pkg = _package(
        "style_inheritance",
        body=[
            (
                '<w:p><w:pPr><w:pStyle w:val="FixtureLeaf"/><w:ind w:hanging="180"/></w:pPr>'
                "<w:r><w:t>Leaf plus a hanging indent.</w:t></w:r></w:p>"
            )
        ],
    )
    indent = _report(pkg, {"find": "Leaf plus a hanging indent."})["paragraph"]["indent"]

    assert indent["left"] == {"value": 720, "source": "style:FixtureLeaf"}
    assert indent["hanging"] == {"value": 180, "source": "direct"}


def test_spacing_keys_come_from_the_layers_that_state_them():
    pkg = DocxPackage.open(build("style_inheritance"))
    spacing = _report(pkg, {"find": "Leaf, based on branch."})["paragraph"]["spacing"]

    assert spacing["after"] == {"value": 200, "source": "style:FixtureRoot"}
    assert spacing["line"]["source"] == "docDefaults"


def test_the_numbering_link_is_reported_with_its_own_provenance():
    """The link, not the level's own properties -- see the module docstring."""
    pkg = _package(
        body=[
            (
                "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
                "<w:r><w:t>A numbered paragraph.</w:t></w:r></w:p>"
            )
        ]
    )
    numbering = _report(pkg, {"find": "A numbered paragraph."})["paragraph"]["numbering"]

    assert numbering["num_id"] == {"value": 1, "source": "direct"}
    assert numbering["level"] == {"value": 0, "source": "direct"}


def test_a_direct_paragraph_property_beats_the_style_that_also_states_it():
    pkg = DocxPackage.open(build("paragraph_styles"))
    styled = _report(pkg, {"find": "Custom body style."})["paragraph"]["alignment"]
    overridden = _report(pkg, {"find": "Styled and directly centered."})["paragraph"][
        "alignment"
    ]

    assert styled == {"value": "both", "source": "style:FixtureBody"}
    assert overridden == {"value": "center", "source": "direct"}


# --------------------------------------------------------------------------------------
# Theme references
# --------------------------------------------------------------------------------------


def test_a_themed_font_reports_the_typeface_the_reference_and_the_layer():
    pkg = DocxPackage.open(build("themes"))
    theme = read_theme(pkg)
    assert theme is not None
    report = _report(pkg, {"find": "Major theme font."})

    assert report["run"]["font"] == {
        "value": theme.major.latin,
        "source": "theme",
        "theme": "majorHAnsi",
        "from": "direct",
    }
    # The script slots are addressed separately, which is the point of reading
    # w:eastAsiaTheme and w:cstheme rather than only w:asciiTheme.
    assert report["run"]["font_east_asia"]["theme"] == "majorEastAsia"
    assert report["run"]["font_cs"]["theme"] == "majorBidi"


def test_an_empty_theme_slot_is_an_answer_not_a_failure_to_resolve():
    """``<a:ea typeface=""/>`` says "no font for this script"; that is not a warning."""
    pkg = DocxPackage.open(build("themes"))
    theme = read_theme(pkg)
    assert theme is not None
    assert theme.minor.ea == ""
    report = _report(pkg, {"paragraph": 0})

    assert report["run"]["font_east_asia"]["value"] == ""
    assert report["run"]["font_east_asia"]["source"] == "theme"
    assert not [w for w in report["warnings"] if "minorEastAsia" in w]


def test_a_theme_colour_resolves_through_the_scheme_and_the_mapping():
    pkg = _package(
        "themes",
        body=[
            (
                '<w:p><w:r><w:rPr><w:color w:themeColor="accent1"/></w:rPr>'
                "<w:t>Plain accent.</w:t></w:r></w:p>"
            ),
            (
                '<w:p><w:r><w:rPr><w:color w:themeColor="text1"/></w:rPr>'
                "<w:t>Mapped text slot.</w:t></w:r></w:p>"
            ),
        ],
    )
    theme = read_theme(pkg)
    assert theme is not None

    plain = _report(pkg, {"find": "Plain accent."})["run"]["color"]
    assert plain == {
        "value": theme.colors["accent1"],
        "source": "theme",
        "theme": "accent1",
        "from": "direct",
    }
    # text1 is not dk1 by definition: it is whatever w:clrSchemeMapping sends
    # t1 to.  Short-circuiting the indirection passes here and fails on a
    # document that swapped its slots.
    mapped = _report(pkg, {"find": "Mapped text slot."})["run"]["color"]
    assert mapped["value"] == theme.colors[theme.scheme_slot("text1")]


def test_a_tint_reaches_the_resolver_and_only_ever_lightens():
    """D-028: the direction is pinned, the constant is not recalibrated here."""
    pkg = DocxPackage.open(build("themes"))
    theme = read_theme(pkg)
    assert theme is not None
    base = theme.colors["accent1"]
    tinted = _report(pkg, {"find": "Accent 1 tinted."})["run"]["color"]
    shaded = _report(pkg, {"find": "Accent 1 shaded."})["run"]["color"]

    assert tinted["source"] == "theme" and tinted["theme"] == "accent1"
    assert shaded["source"] == "theme" and shaded["theme"] == "accent1"
    # A tint that was dropped on the way to the resolver would give the base back.
    assert tinted["value"] != base
    assert shaded["value"] != base
    assert all(
        light >= plain for light, plain in zip(_channels(tinted["value"]), _channels(base))
    )
    assert all(
        dark <= plain for dark, plain in zip(_channels(shaded["value"]), _channels(base))
    )


def test_a_theme_reference_that_does_not_resolve_falls_back_to_the_cached_value():
    pkg = _package(
        "themes",
        body=[
            (
                '<w:p><w:r><w:rPr><w:rFonts w:ascii="Cached" w:asciiTheme="notATheme"/></w:rPr>'
                "<w:t>Broken reference.</w:t></w:r></w:p>"
            )
        ],
    )
    report = _report(pkg, {"find": "Broken reference."})

    assert report["run"]["font"] == {
        "value": "Cached",
        "source": "direct",
        "theme": "notATheme",
    }
    assert any("notATheme" in warning for warning in report["warnings"])


# --------------------------------------------------------------------------------------
# Table styles: named, never guessed
# --------------------------------------------------------------------------------------

_TABLE_STYLES = (
    """
    <w:style w:type="table" w:styleId="XorTable">
      <w:name w:val="Xor Table"/>
      <w:pPr><w:jc w:val="right"/></w:pPr>
      <w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>
    </w:style>
    """,
)

_TABLE_BODY = (
    """
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="XorTable"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:tc><w:p><w:r><w:t>Cell left to the table style.</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Cell that says otherwise.</w:t></w:r>
      </w:p></w:tc></w:tr>
    </w:tbl>
    """,
)


def test_a_table_style_names_the_properties_it_owns_and_resolves_none_of_them():
    pkg = _package(styles=_TABLE_STYLES, body=_TABLE_BODY)
    report = _report(pkg, {"find": "Cell left to the table style."})

    assert report["table_style"] == "XorTable"
    assert report["index"] is None  # D-016: a cell paragraph has no V2 index
    assert report["run"]["bold"] == {
        "value": UNRESOLVED,
        "source": "table_style:XorTable",
    }
    assert report["run"]["color"] == {
        "value": UNRESOLVED,
        "source": "table_style:XorTable",
    }
    assert report["paragraph"]["alignment"] == {
        "value": UNRESOLVED,
        "source": "table_style:XorTable",
    }
    assert any("XorTable" in warning for warning in report["warnings"])


def test_a_higher_layer_resolves_what_the_table_style_could_not():
    """``unresolved`` is the answer only where nothing more specific speaks."""
    pkg = _package(styles=_TABLE_STYLES, body=_TABLE_BODY)
    report = _report(pkg, {"find": "Cell that says otherwise."})

    assert report["run"]["bold"] == {"value": False, "source": "direct"}
    assert report["paragraph"]["alignment"] == {"value": "center", "source": "direct"}
    # The one nobody overrode is still honest about being unknown.
    assert report["run"]["color"]["value"] == UNRESOLVED


def test_a_paragraph_outside_a_table_reports_no_table_style_and_no_warning():
    report = _report(DocxPackage.open(build("simple")), {"paragraph": 1})

    assert report["table_style"] is None
    assert report["warnings"] == []


# --------------------------------------------------------------------------------------
# Ranges: uniform, mixed, narrowed
# --------------------------------------------------------------------------------------

_MIXED_BODY = (
    """
    <w:p>
      <w:r><w:rPr><w:b/><w:color w:val="111111"/></w:rPr><w:t>AAAA</w:t></w:r>
      <w:r><w:rPr><w:color w:val="222222"/></w:rPr><w:t>BBBB</w:t></w:r>
      <w:r><w:rPr><w:b/><w:color w:val="111111"/></w:rPr><w:t>CCCC</w:t></w:r>
    </w:p>
    """,
)


def test_a_range_whose_runs_disagree_answers_mixed():
    pkg = _package(body=_MIXED_BODY)
    report = _report(pkg, {"find": "AAAABBBB"})

    # bold is stated on one run and absent from the other: still a disagreement.
    assert report["run"]["bold"] == {"value": MIXED, "source": MIXED}
    # Both runs state a colour directly, so the source survives the disagreement.
    assert report["run"]["color"] == {"value": MIXED, "source": "direct"}


def test_narrowing_the_range_to_one_run_resolves_the_disagreement():
    pkg = _package(body=_MIXED_BODY)
    whole = _report(pkg, {"find": "AAAABBBBCCCC"})
    narrowed = _report(pkg, {"find": "AAAABBBBCCCC"}, start=0, end=4)

    assert whole["run"]["color"]["value"] == MIXED
    assert narrowed["run"]["color"] == {"value": "111111", "source": "direct"}
    assert narrowed["start"] == 0 and narrowed["end"] == 4


def test_runs_that_agree_are_not_reported_as_mixed():
    pkg = _package(body=_MIXED_BODY)
    report = _report(pkg, {"find": "AAAABBBBCCCC"}, start=0, end=4)
    other = _report(pkg, {"find": "AAAABBBBCCCC"}, start=8, end=12)

    assert report["run"]["bold"] == other["run"]["bold"] == {
        "value": True,
        "source": "direct",
    }


def test_an_empty_paragraph_answers_from_its_paragraph_mark():
    """What a reader would see if they typed there -- not "no formatting at all"."""
    pkg = _package(
        body=[
            (
                '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:b/>'
                '<w:color w:val="ABCDEF"/></w:rPr></w:pPr></w:p>'
            )
        ]
    )
    root = dict(pkg.stories())["document"]
    index = len(indexed_paragraphs(root)) - 1
    report = _report(pkg, {"paragraph": index})

    assert report["start"] == report["end"] == 0
    assert report["run"]["bold"] == {"value": True, "source": "direct"}
    assert report["run"]["color"] == {"value": "ABCDEF", "source": "direct"}


# --------------------------------------------------------------------------------------
# Reading does not write
# --------------------------------------------------------------------------------------


def test_reading_a_range_does_not_split_a_run(tmp_path: Path):
    """The obvious implementation goes through ``ranges.resolve``, which splits."""
    pkg = _package(body=_MIXED_BODY)
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    pkg.save(before)

    # A span whose boundaries fall inside runs: exactly what a splitter would cut.
    _report(pkg, {"find": "AAAABBBBCCCC"}, start=2, end=10)
    pkg.save(after)

    delta = diff(snapshot(before), snapshot(after))
    assert delta.is_empty(), delta.describe()


def test_a_target_is_required():
    pkg = DocxPackage.open(build("simple"))
    with pytest.raises(TypeError):
        effective_format(pkg, {"paragraph": 0})  # type: ignore[arg-type]


def test_the_report_names_where_it_applies():
    pkg = DocxPackage.open(build("simple"))
    target = resolve(pkg, {"paragraph": 1})
    report = effective_format(pkg, target)

    assert isinstance(target, Target)
    assert report["story"] == target.story
    assert report["index"] == target.index == 1
    assert (report["start"], report["end"]) == (target.start, target.end)


# --------------------------------------------------------------------------------------
# The tool
# --------------------------------------------------------------------------------------


def test_the_tool_module_exports_a_registered_spec():
    """D-024: a module without a non-empty ``TOOLS`` is discovered as nothing."""
    assert [spec.name for spec in TOOLS] == ["doc_get_effective_format"]
    assert "doc_get_effective_format" in {spec.name for spec in discover_tool_specs()}


def test_the_tool_answers_for_a_locator(fixture_docx):
    path = fixture_docx("style_inheritance")
    report = doc_get_effective_format(str(path), {"find": "Leaf, based on branch."})

    assert report["effective"]["paragraph_style"] == "FixtureLeaf"
    assert report["effective"]["run"]["bold"]["source"] == "style:FixtureBranch"
    assert report["warnings"] == report["effective"]["warnings"]


def test_the_tool_accepts_explicit_offsets(fixture_docx):
    path = fixture_docx("character_styles")
    whole = doc_get_effective_format(str(path), {"paragraph": 1})["effective"]
    narrowed = doc_get_effective_format(
        str(path), {"paragraph": 1}, start=12, end=22
    )["effective"]

    assert whole["run"]["color"]["value"] == MIXED
    assert narrowed["run"]["color"] == {
        "value": "C00000",
        "source": "style:FixtureEmphasis",
    }


def test_the_tool_refuses_offsets_outside_the_paragraph(fixture_docx):
    path = fixture_docx("simple")
    with pytest.raises(LocatorError) as caught:
        doc_get_effective_format(str(path), {"paragraph": 1}, start=0, end=10_000)

    assert caught.value.code == "out-of-range"


def test_the_tool_leaves_the_file_alone(fixture_docx):
    path = fixture_docx("style_inheritance")
    before = path.read_bytes()
    doc_get_effective_format(str(path), {"paragraph": 1})

    assert path.read_bytes() == before
