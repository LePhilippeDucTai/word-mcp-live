"""Tests for :mod:`word_document_server.engine.theme`.

Three things are pinned here.

*The theme is read, never parsed live.*
    ``word/theme/theme1.xml`` is outside
    :data:`~word_document_server.engine.package.LIVE_CONTENT_TYPES` (D-007) and
    must stay there: reading it has to go through ``Part.blob``, so that a
    package opened, read and saved still carries the part byte for byte.  That
    is asserted on the archive itself rather than through
    ``assert_unchanged_except``, whose ``parts=`` must never name a blob part.

*The indirection is followed all the way.*
    ``w:themeColor="text1"`` does not name the theme's ``dk1``; it names
    whatever ``w:clrSchemeMapping`` sends ``t1`` to, which happens to be
    ``dark1``, which is ``dk1``.  A reader that short-circuits any of the three
    steps gets the right answer on a default document and the wrong one on a
    document that swapped its slots -- so the swapped document is tested too.

*Tint and shade are a hint, and say so.*
    The fixture caches the values Word itself computed for
    ``accent1 + themeTint="66"`` and ``accent1 + themeShade="BF"``.  They are
    checked against a reimplementation in floating point, within a tolerance,
    because that is the honest claim -- see the module docstring of
    ``engine/theme.py``.  What *is* exact is the algebra around them: ``FF``
    changes nothing, ``00`` saturates, a tint never darkens and a shade never
    lightens.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from tests.fixtures.builders import build
from tests.support.libreoffice import fodt_to_docx, requires_libreoffice
from word_document_server.engine.errors import PackageError
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.theme import (
    COLOR_SLOTS,
    DEFAULT_COLOR_MAP,
    THEME_PARTNAME,
    Theme,
    ThemeFonts,
    apply_tint_shade,
    read_theme,
)

THEME_MEMBER = "word/theme/theme1.xml"

#: The colour scheme of the template every fixture is built from.
TEMPLATE_COLORS = {
    "dk1": "000000",
    "lt1": "FFFFFF",
    "dk2": "1F497D",
    "lt2": "EEECE1",
    "accent1": "4F81BD",
    "accent2": "C0504D",
    "accent3": "9BBB59",
    "accent4": "8064A2",
    "accent5": "4BACC6",
    "accent6": "F79646",
    "hlink": "0000FF",
    "folHlink": "800080",
}


def _pkg(name: str = "themes") -> DocxPackage:
    return DocxPackage.open(build(name))


def _theme(name: str = "themes") -> Theme:
    theme = read_theme(_pkg(name))
    assert theme is not None
    return theme


def _rewritten(blob: bytes, **replacements: bytes | None) -> bytes:
    """Rebuild `blob` with members replaced, or dropped when given ``None``."""
    source = zipfile.ZipFile(io.BytesIO(blob))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            key = name.replace("/", "__").replace(".", "_")
            if key in replacements:
                if replacements[key] is None:
                    continue
                target.writestr(name, replacements[key])
            else:
                target.writestr(name, source.read(name))
    return buffer.getvalue()


def _without_theme(blob: bytes) -> bytes:
    """`blob` with the theme part, its content-type override and its rel removed."""
    source = zipfile.ZipFile(io.BytesIO(blob))
    types = re.sub(
        rb'<Override PartName="/word/theme/theme1\.xml"[^>]*/>',
        b"",
        source.read("[Content_Types].xml"),
    )
    rels = re.sub(
        rb"<Relationship[^>]*theme1\.xml\"[^>]*/>",
        b"",
        source.read("word/_rels/document.xml.rels"),
    )
    return _rewritten(
        blob,
        word__theme__theme1_xml=None,
        __Content_Types___xml=types,
        word___rels__document_xml_rels=rels,
    )


# --------------------------------------------------------------------------------------
# Reading without disturbing
# --------------------------------------------------------------------------------------


def test_reading_the_theme_leaves_the_part_byte_for_byte_identical() -> None:
    # D-007: theme1.xml is not in LIVE_CONTENT_TYPES, so it is never reparsed
    # and never reindented.  Reading it must not change that.
    blob = build("themes")
    before = zipfile.ZipFile(io.BytesIO(blob)).read(THEME_MEMBER)

    pkg = DocxPackage.open(blob)
    assert read_theme(pkg) is not None
    after = zipfile.ZipFile(io.BytesIO(pkg.to_bytes())).read(THEME_MEMBER)

    assert after == before


def test_the_theme_part_is_still_refused_by_root_of() -> None:
    # The guard that makes the assertion above hold: asking for a live root on
    # the theme part fails loudly rather than handing out one that save drops.
    pkg = _pkg()
    with pytest.raises(PackageError, match="not loaded live"):
        pkg.root_of(pkg.part(THEME_PARTNAME))


def test_a_theme_is_plain_data_detached_from_the_package() -> None:
    theme = _theme()
    # Nothing here is an lxml element, so nothing can be mistaken for a handle.
    assert isinstance(theme.name, str)
    assert isinstance(theme.major, ThemeFonts)
    assert all(isinstance(value, str) for value in theme.colors.values())


# --------------------------------------------------------------------------------------
# Fonts
# --------------------------------------------------------------------------------------


def test_the_two_font_schemes_are_read() -> None:
    theme = _theme()
    assert theme.name == "Office Theme"
    assert theme.major.latin == "Calibri"
    assert theme.minor.latin == "Cambria"


def test_an_empty_typeface_is_reported_as_empty_not_missing() -> None:
    # <a:ea typeface=""/> is what the template writes; it means "no font of its
    # own for this script", which is not the same as "the theme says nothing".
    theme = _theme()
    assert theme.major.ea == ""
    assert theme.major.cs == ""


@pytest.mark.parametrize(
    ("reference", "expected"),
    [
        ("majorHAnsi", "Calibri"),
        ("majorAscii", "Calibri"),
        ("majorEastAsia", ""),
        ("majorBidi", ""),
        ("minorHAnsi", "Cambria"),
        ("minorAscii", "Cambria"),
        ("minorEastAsia", ""),
        ("minorBidi", ""),
    ],
)
def test_every_st_theme_value_resolves_to_a_typeface(reference: str, expected: str) -> None:
    assert _theme().resolve_font(reference) == expected


@pytest.mark.parametrize("reference", ["", "minorhansi", "majorLatin", "nonsense"])
def test_an_unknown_font_reference_resolves_to_none(reference: str) -> None:
    # None means "cannot say", which a caller can branch on; a fallback to the
    # minor latin font would hide a reference this module does not understand.
    assert _theme().resolve_font(reference) is None


# --------------------------------------------------------------------------------------
# Colours
# --------------------------------------------------------------------------------------


def test_every_scheme_slot_is_read() -> None:
    assert dict(_theme().colors) == TEMPLATE_COLORS
    assert set(TEMPLATE_COLORS) == set(COLOR_SLOTS)


def test_a_sys_clr_slot_is_read_from_its_last_clr() -> None:
    # dk1 and lt1 are <a:sysClr val="windowText" lastClr="000000"/>: the usable
    # value is in lastClr, and a reader that only knows a:srgbClr loses both.
    theme = _theme()
    assert theme.colors["dk1"] == "000000"
    assert theme.colors["lt1"] == "FFFFFF"


@pytest.mark.parametrize(
    ("theme_color", "slot"),
    [
        ("accent1", "accent1"),
        ("accent6", "accent6"),
        ("hyperlink", "hlink"),
        ("followedHyperlink", "folHlink"),
        ("text1", "dk1"),
        ("background1", "lt1"),
        ("text2", "dk2"),
        ("background2", "lt2"),
        ("dark1", "dk1"),
        ("light1", "lt1"),
        ("dark2", "dk2"),
        ("light2", "lt2"),
    ],
)
def test_every_theme_colour_reaches_its_scheme_slot(theme_color: str, slot: str) -> None:
    theme = _theme()
    assert theme.scheme_slot(theme_color) == slot
    assert theme.resolve_color(theme_color) == TEMPLATE_COLORS[slot]


@pytest.mark.parametrize("theme_color", ["none", "", "Accent1", "accent7", "major"])
def test_a_colour_this_module_cannot_name_resolves_to_none(theme_color: str) -> None:
    assert _theme().resolve_color(theme_color) is None


def test_a_slot_the_theme_does_not_define_resolves_to_none() -> None:
    theme = Theme(name="t", major=ThemeFonts(), minor=ThemeFonts(), colors={})
    assert theme.scheme_slot("accent1") == "accent1"
    # The slot is named, the theme just has no colour for it: "cannot say",
    # never a fabricated black.
    assert theme.resolve_color("accent1") is None


def test_the_colour_map_is_followed_rather_than_assumed() -> None:
    # A document may swap its text and background slots without touching the
    # theme.  text1 must then resolve to white, not to the theme's dk1.
    swapped = Theme(
        name="t",
        major=ThemeFonts(),
        minor=ThemeFonts(),
        colors=TEMPLATE_COLORS,
        color_map={**DEFAULT_COLOR_MAP, "t1": "light1", "bg1": "dark1"},
    )
    assert swapped.resolve_color("text1") == "FFFFFF"
    assert swapped.resolve_color("background1") == "000000"
    # dark1 names the scheme slot outright and is untouched by the mapping.
    assert swapped.resolve_color("dark1") == "000000"


def test_the_document_color_map_is_read_from_the_settings_part() -> None:
    assert dict(_theme().color_map) == dict(DEFAULT_COLOR_MAP)


def test_a_document_without_a_colour_map_falls_back_to_word_s_default() -> None:
    blob = build("themes")
    settings = zipfile.ZipFile(io.BytesIO(blob)).read("word/settings.xml")
    stripped = re.sub(rb"<w:clrSchemeMapping[^>]*/>", b"", settings)
    assert stripped != settings

    theme = read_theme(DocxPackage.open(_rewritten(blob, word__settings_xml=stripped)))
    assert theme is not None
    assert dict(theme.color_map) == dict(DEFAULT_COLOR_MAP)
    assert theme.resolve_color("text1") == "000000"


# --------------------------------------------------------------------------------------
# Tint and shade
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("color", ["4472C4", "000000", "FFFFFF", "F79646"])
def test_neither_tint_nor_shade_leaves_the_colour_alone(color: str) -> None:
    assert apply_tint_shade(color) == color
    # FF is the identity on both scales, which is what makes "no tint" and
    # "tint of FF" the same document.
    assert apply_tint_shade(color, tint="FF") == color
    assert apply_tint_shade(color, shade="FF") == color


@pytest.mark.parametrize("color", ["4472C4", "1F497D", "F79646"])
def test_the_extremes_saturate(color: str) -> None:
    assert apply_tint_shade(color, tint="00") == "FFFFFF"
    assert apply_tint_shade(color, shade="00") == "000000"


@pytest.mark.parametrize("amount", ["11", "66", "99", "CC"])
def test_a_tint_never_darkens_and_a_shade_never_lightens(amount: str) -> None:
    base = "4472C4"

    def channels(value: str) -> tuple[int, ...]:
        return tuple(int(value[index : index + 2], 16) for index in (0, 2, 4))

    original = channels(base)
    assert all(a >= b for a, b in zip(channels(apply_tint_shade(base, tint=amount)), original))
    assert all(a <= b for a, b in zip(channels(apply_tint_shade(base, shade=amount)), original))


@pytest.mark.parametrize(
    ("kwargs", "cached"),
    [
        ({"tint": "66"}, "B4C6E7"),
        ({"shade": "BF"}, "2F5597"),
    ],
)
def test_tint_and_shade_land_where_word_cached_them(kwargs: dict, cached: str) -> None:
    # The pairs the `themes` fixture carries are Word's own output for the
    # Office 2013 accent1.  The reimplementation is floating point where Word's
    # is fixed point, so the claim is "within a unit or two per channel", and
    # the module docstring says so rather than pretending to be exact.
    produced = apply_tint_shade("4472C4", **kwargs)
    distance = [
        abs(int(produced[i : i + 2], 16) - int(cached[i : i + 2], 16)) for i in (0, 2, 4)
    ]
    assert max(distance) <= 2, f"{produced} is too far from Word's {cached}"


def test_resolve_color_applies_the_tint_it_is_given() -> None:
    theme = _theme()
    plain = theme.resolve_color("accent1")
    assert theme.resolve_color("accent1", tint="66") != plain
    assert theme.resolve_color("accent1", tint="66") == apply_tint_shade(plain, tint="66")


def test_a_tint_may_be_given_as_an_int_or_as_hex_digits() -> None:
    assert apply_tint_shade("4472C4", tint=0x66) == apply_tint_shade("4472C4", tint="66")
    assert apply_tint_shade("4472C4", shade=191) == apply_tint_shade("4472C4", shade="bf")


def test_a_colour_carrying_both_a_tint_and_a_shade_is_refused() -> None:
    # The schema makes them exclusive; a document with both is telling two
    # stories about one colour, and picking one silently would hide that.
    with pytest.raises(ValueError, match="either a tint or a shade"):
        apply_tint_shade("4472C4", tint="66", shade="BF")
    with pytest.raises(ValueError, match="either a tint or a shade"):
        _theme().resolve_color("accent1", tint="66", shade="BF")


@pytest.mark.parametrize("bad", ["", "6", "666", "GG", "0x66"])
def test_a_tint_that_is_not_two_hex_digits_is_refused(bad: str) -> None:
    with pytest.raises(ValueError, match="ST_UcharHexNumber"):
        apply_tint_shade("4472C4", tint=bad)


@pytest.mark.parametrize("bad", [True, 3.5, [], b"66"])
def test_a_tint_of_the_wrong_type_is_a_type_error(bad: object) -> None:
    # bool is an int in Python and would otherwise sail through as 0 or 1.
    with pytest.raises(TypeError):
        apply_tint_shade("4472C4", tint=bad)


@pytest.mark.parametrize("bad", [-1, 256])
def test_a_tint_outside_a_byte_is_refused(bad: int) -> None:
    with pytest.raises(ValueError, match="between 0 and 255"):
        apply_tint_shade("4472C4", tint=bad)


@pytest.mark.parametrize("bad", ["", "auto", "4472C", "#GGGGGG", "44 72 C4"])
def test_a_base_colour_that_is_not_rrggbb_is_refused(bad: str) -> None:
    with pytest.raises(ValueError, match="RRGGBB"):
        apply_tint_shade(bad, tint="66")


def test_a_leading_hash_is_accepted_and_the_answer_is_upper_case() -> None:
    assert apply_tint_shade("#4472c4") == "4472C4"


# --------------------------------------------------------------------------------------
# Documents without a usable theme
# --------------------------------------------------------------------------------------


def test_a_document_without_a_theme_part_reads_as_none() -> None:
    # Legal: Word falls back to its own built-in theme.  None is "this document
    # says nothing", which resolves references to None rather than to guesses.
    pkg = DocxPackage.open(_without_theme(build("themes")))
    assert pkg.find_part(THEME_PARTNAME) is None
    assert read_theme(pkg) is None


def test_a_theme_part_that_does_not_parse_is_a_package_error() -> None:
    blob = _rewritten(build("themes"), word__theme__theme1_xml=b"<a:theme><not closed")
    with pytest.raises(PackageError, match="not well-formed"):
        read_theme(DocxPackage.open(blob))


def test_a_theme_without_theme_elements_reads_as_empty_rather_than_failing() -> None:
    minimal = (
        b'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        b'name="Bare"/>'
    )
    theme = read_theme(
        DocxPackage.open(_rewritten(build("themes"), word__theme__theme1_xml=minimal))
    )
    assert theme is not None
    assert theme.name == "Bare"
    assert dict(theme.colors) == {}
    assert theme.major == ThemeFonts()
    assert theme.resolve_color("accent1") is None
    assert theme.resolve_font("minorHAnsi") == ""


def test_a_slot_expressed_in_a_form_this_module_does_not_read_is_left_out() -> None:
    # a:schemeClr and a:prstClr are legal in a:clrScheme.  Skipping the slot is
    # the honest answer; approximating it would put a wrong colour in the map.
    blob = build("themes")
    original = zipfile.ZipFile(io.BytesIO(blob)).read(THEME_MEMBER)
    patched = original.replace(
        b'<a:accent2>\n        <a:srgbClr val="C0504D"/>\n      </a:accent2>',
        b'<a:accent2>\n        <a:prstClr val="red"/>\n      </a:accent2>',
    )
    assert patched != original

    theme = read_theme(
        DocxPackage.open(_rewritten(blob, word__theme__theme1_xml=patched))
    )
    assert theme is not None
    assert "accent2" not in theme.colors
    assert theme.resolve_color("accent2") is None
    # The slots around it are untouched.
    assert theme.colors["accent1"] == TEMPLATE_COLORS["accent1"]


# --------------------------------------------------------------------------------------
# A theme this project did not write
# --------------------------------------------------------------------------------------


@requires_libreoffice
def test_a_libreoffice_theme_is_read_the_same_way(tmp_path_factory) -> None:
    # LibreOffice is a fixture producer, never a source of truth: what is
    # checked is that its theme goes through the same reader, not that its
    # colours are the right ones.
    path = fodt_to_docx("rich", tmp_path_factory.mktemp("soffice-theme"))
    pkg = DocxPackage.open(path)
    theme = read_theme(pkg)
    if theme is None:
        pytest.skip("this LibreOffice build writes no theme part")
    assert set(theme.colors) <= set(COLOR_SLOTS)
    for slot, value in theme.colors.items():
        assert re.fullmatch(r"[0-9A-F]{6}", value), f"{slot} is {value!r}"
    assert theme.resolve_font("minorHAnsi") is not None
