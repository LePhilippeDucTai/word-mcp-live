"""The document theme: the fonts and colours styles refer to without naming.

A ``w:rPr`` rarely says ``Aptos``.  It says ``w:asciiTheme="minorHAnsi"``, and
the font that resolves to lives in ``word/theme/theme1.xml``.  Colours work the
same way: ``w:themeColor="accent1"`` names a slot of the theme's colour scheme,
optionally darkened by ``w:themeShade`` or lightened by ``w:themeTint``.  A
reader that stops at the ``w:val`` cached next to the reference reads whatever
the producing application happened to compute last -- which is stale as soon as
the theme changes, and absent altogether when the producer did not bother.  This
module is the other half: it reads the theme and resolves the references.

Reading, never parsing live
---------------------------
``word/theme/theme1.xml`` is *not* in
:data:`~word_document_server.engine.package.LIVE_CONTENT_TYPES` and must not be
added to it (D-007): the engine never edits the theme, and loading it through
python-docx's parser would reindent it, so a plain open/save would stop
reproducing it byte for byte.  :func:`read_theme` therefore goes through
``Part.blob`` and parses a detached, read-only copy.  Nothing here writes.

What resolves to what
---------------------
The theme's colour scheme names its slots ``dk1``, ``lt1``, ``dk2``, ``lt2``,
``accent1``..``accent6``, ``hlink`` and ``folHlink``.  WordprocessingML names
them differently -- ``text1``, ``background1``, ``accent1``, ``hyperlink``, ... --
and routes most of them through the ``w:clrSchemeMapping`` of
``word/settings.xml``, which is how a document can swap its text and background
slots without touching the theme.  :meth:`Theme.resolve_color` follows that
indirection; :func:`read_theme` reads the mapping, and falls back to Word's own
default when the document carries none.

Fidelity of tint and shade
--------------------------
``w:themeTint`` and ``w:themeShade`` scale the *luminance* of the slot colour in
HSL space, which is what Word does and what :func:`apply_tint_shade` reproduces.
The arithmetic is done in floating point here and in fixed point there, so a
channel may land one or two units away from the value Word itself caches in the
neighbouring ``w:val``.  A resolved colour is therefore a **rendering hint** --
good enough to tell an agent that a run is a dark blue -- and never something to
write back into the document: the reference and its cached value are the
document's own business, and this module only reads them.
"""

from __future__ import annotations

import colorsys
import string
from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType

from lxml import etree

from word_document_server.engine.errors import PackageError
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.xmlns import qn

__all__ = [
    "COLOR_SLOTS",
    "DEFAULT_COLOR_MAP",
    "THEME_PARTNAME",
    "A",
    "Theme",
    "ThemeFonts",
    "apply_tint_shade",
    "read_theme",
]

#: DrawingML main namespace.  Declared here rather than in
#: :mod:`word_document_server.engine.xmlns`, which holds the prefixes the engine
#: *writes*: the theme is read-only, and nothing else in the engine touches
#: DrawingML.
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

#: Conventional part name of the theme, used when no relationship points at one.
THEME_PARTNAME = "/word/theme/theme1.xml"

#: Relationship type of the theme part.
_THEME_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

#: Slots of ``a:clrScheme``, in schema order.  The names are DrawingML's, not
#: WordprocessingML's -- see :data:`DEFAULT_COLOR_MAP` for the bridge.
COLOR_SLOTS = (
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)

#: ``w:clrSchemeMapping`` attribute -> the ``ST_ColorSchemeIndex`` Word uses when
#: the document declares no mapping.  This *is* the mapping Word writes into a
#: new document, so a document without one behaves like a document with this one.
DEFAULT_COLOR_MAP: Mapping[str, str] = MappingProxyType(
    {
        "bg1": "light1",
        "t1": "dark1",
        "bg2": "light2",
        "t2": "dark2",
        "accent1": "accent1",
        "accent2": "accent2",
        "accent3": "accent3",
        "accent4": "accent4",
        "accent5": "accent5",
        "accent6": "accent6",
        "hyperlink": "hyperlink",
        "followedHyperlink": "followedHyperlink",
    }
)

#: ``ST_ThemeColor`` values that go through ``w:clrSchemeMapping`` -> the
#: attribute of that element which governs them.
_MAPPED_THEME_COLORS: Mapping[str, str] = MappingProxyType(
    {
        "text1": "t1",
        "background1": "bg1",
        "text2": "t2",
        "background2": "bg2",
        "accent1": "accent1",
        "accent2": "accent2",
        "accent3": "accent3",
        "accent4": "accent4",
        "accent5": "accent5",
        "accent6": "accent6",
        "hyperlink": "hyperlink",
        "followedHyperlink": "followedHyperlink",
    }
)

#: ``ST_ColorSchemeIndex`` -- and the four ``ST_ThemeColor`` values that name a
#: scheme slot outright rather than through the mapping -> the ``a:clrScheme``
#: slot.  ``dark1`` is the scheme's ``dk1`` whatever the mapping says; it is
#: ``text1`` that the mapping may send elsewhere.
_SCHEME_INDEX_SLOT: Mapping[str, str] = MappingProxyType(
    {
        "dark1": "dk1",
        "light1": "lt1",
        "dark2": "dk2",
        "light2": "lt2",
        "accent1": "accent1",
        "accent2": "accent2",
        "accent3": "accent3",
        "accent4": "accent4",
        "accent5": "accent5",
        "accent6": "accent6",
        "hyperlink": "hlink",
        "followedHyperlink": "folHlink",
    }
)

#: ``ST_Theme`` values (``w:asciiTheme``, ``w:cstheme``, ...) -> the font scheme
#: and the script slot inside it.  ``ascii`` and ``hAnsi`` share the latin slot,
#: which is why a document that sets only one of them still renders in it.
_THEME_FONTS: Mapping[str, tuple[str, str]] = MappingProxyType(
    {
        "majorAscii": ("major", "latin"),
        "majorHAnsi": ("major", "latin"),
        "majorEastAsia": ("major", "ea"),
        "majorBidi": ("major", "cs"),
        "minorAscii": ("minor", "latin"),
        "minorHAnsi": ("minor", "latin"),
        "minorEastAsia": ("minor", "ea"),
        "minorBidi": ("minor", "cs"),
    }
)

_A_THEME_ELEMENTS = f"{{{A}}}themeElements"
_A_CLR_SCHEME = f"{{{A}}}clrScheme"
_A_FONT_SCHEME = f"{{{A}}}fontScheme"
_A_MAJOR_FONT = f"{{{A}}}majorFont"
_A_MINOR_FONT = f"{{{A}}}minorFont"
_A_SRGB_CLR = f"{{{A}}}srgbClr"
_A_SYS_CLR = f"{{{A}}}sysClr"

_W_CLR_SCHEME_MAPPING = qn("w:clrSchemeMapping")

_SETTINGS_PARTNAME = "/word/settings.xml"


def _hex6(value: str | None) -> str | None:
    """Normalise an ``RRGGBB`` value, or ``None`` when it is not one."""
    if value is None:
        return None
    raw = value.strip().removeprefix("#")
    if len(raw) != 6 or not all(character in string.hexdigits for character in raw):
        return None
    return raw.upper()


def _hex_byte(value: str | int | None, name: str) -> int | None:
    """Read a ``ST_UcharHexNumber`` -- ``"BF"``, or the int 191.

    Returns:
        The value as 0..255, or ``None`` when `value` is ``None``.

    Raises:
        TypeError: if it is neither a string nor an int.  ``True`` counts as
            neither: ``bool`` is an ``int`` in Python, but never a colour
            component here.
        ValueError: if it is a string that is not two hexadecimal digits, or an
            int outside a byte.
    """
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (str, int)):
        raise TypeError(
            f"{name} takes two hexadecimal digits or an int, got {type(value).__name__}"
        )
    if isinstance(value, int):
        if not 0 <= value <= 255:
            raise ValueError(f"{name} must be between 0 and 255, got {value}")
        return value
    raw = value.strip()
    if len(raw) != 2 or not all(character in string.hexdigits for character in raw):
        raise ValueError(
            f"{name} is a two-digit hexadecimal number (ST_UcharHexNumber), got {value!r}"
        )
    return int(raw, 16)


def apply_tint_shade(
    color: str, tint: str | int | None = None, shade: str | int | None = None
) -> str:
    """Lighten or darken an ``RRGGBB`` colour the way Word's tint and shade do.

    Both scale the colour's *luminance* in HSL space, on a 0..255 scale where
    ``FF`` is "leave it alone": a tint moves the luminance towards white by
    ``1 - tint/255``, a shade multiplies it by ``shade/255``.  Applying neither
    returns `color` unchanged.

    See the module docstring on why the result may sit a unit or two from the
    value Word caches for the same pair.

    Args:
        color: the base colour, ``RRGGBB``, with or without a leading ``#``.
        tint: ``w:themeTint``, as its two hexadecimal digits or as an int.
        shade: ``w:themeShade``, likewise.

    Returns:
        The resulting colour as six upper-case hexadecimal digits.

    Raises:
        TypeError: if `tint` or `shade` is neither a string nor an int.
        ValueError: if `color` is not ``RRGGBB``, if `tint` or `shade` is out of
            range or malformed, or if both are given -- ``w:themeTint`` and
            ``w:themeShade`` are mutually exclusive in the schema, and a
            document carrying both is telling two different stories about the
            same colour.
    """
    base = _hex6(color)
    if base is None:
        raise ValueError(f"expected an RRGGBB colour, got {color!r}")
    tint_value = _hex_byte(tint, "tint")
    shade_value = _hex_byte(shade, "shade")
    if tint_value is not None and shade_value is not None:
        raise ValueError("a colour carries either a tint or a shade, never both")
    if tint_value is None and shade_value is None:
        return base

    red, green, blue = (int(base[index : index + 2], 16) / 255.0 for index in (0, 2, 4))
    hue, luminance, saturation = colorsys.rgb_to_hls(red, green, blue)
    if tint_value is not None:
        factor = tint_value / 255.0
        luminance = luminance * factor + (1.0 - factor)
    else:
        luminance = luminance * (shade_value / 255.0)
    channels = colorsys.hls_to_rgb(hue, min(max(luminance, 0.0), 1.0), saturation)
    return "".join(f"{min(255, max(0, round(channel * 255))):02X}" for channel in channels)


@dataclass(frozen=True)
class ThemeFonts:
    """One font scheme: the typeface per script slot.

    A slot the theme leaves empty is the empty string, not ``None``: that is
    what the part stores (``<a:ea typeface=""/>``), and it means "no font of its
    own for this script", which a renderer answers by falling back to `latin`.

    Attributes:
        latin: ``a:latin``, the typeface ``ascii`` and ``hAnsi`` both resolve to.
        ea: ``a:ea``, for East Asian text.
        cs: ``a:cs``, for complex scripts.
    """

    latin: str = ""
    ea: str = ""
    cs: str = ""


@dataclass(frozen=True)
class Theme:
    """The parts of ``theme1.xml`` a WordprocessingML document refers to.

    Plain data, read once: holding a :class:`Theme` does not keep the package
    alive and does not go stale in a way that could be mistaken for a handle.

    Attributes:
        name: the ``name`` of ``a:theme`` (``"Office Theme"``), or ``""``.
        major: the heading font scheme.
        minor: the body font scheme.
        colors: scheme slot (:data:`COLOR_SLOTS`) -> ``RRGGBB``.  A slot the
            theme expresses in a form this module does not read -- anything but
            ``a:srgbClr`` and ``a:sysClr`` -- is absent rather than guessed.
        color_map: the ``w:clrSchemeMapping`` of the document, or
            :data:`DEFAULT_COLOR_MAP` when it declares none.
    """

    name: str
    major: ThemeFonts
    minor: ThemeFonts
    colors: Mapping[str, str]
    color_map: Mapping[str, str] = DEFAULT_COLOR_MAP

    def scheme_slot(self, theme_color: str) -> str | None:
        """The ``a:clrScheme`` slot a ``w:themeColor`` value names.

        ``"accent1"`` -> ``"accent1"``, ``"text1"`` -> ``"dk1"`` under the
        default mapping.  Returns ``None`` for ``"none"``, and for any value
        outside ``ST_ThemeColor``.
        """
        if not isinstance(theme_color, str):
            return None
        name = theme_color.strip()
        if name in _SCHEME_INDEX_SLOT and name not in _MAPPED_THEME_COLORS:
            # dark1, light1, dark2, light2: the scheme slot, not through the map.
            return _SCHEME_INDEX_SLOT[name]
        attribute = _MAPPED_THEME_COLORS.get(name)
        if attribute is None:
            return None
        index = self.color_map.get(attribute, DEFAULT_COLOR_MAP.get(attribute))
        return None if index is None else _SCHEME_INDEX_SLOT.get(index)

    def resolve_color(
        self,
        theme_color: str,
        tint: str | int | None = None,
        shade: str | int | None = None,
    ) -> str | None:
        """Resolve a ``w:themeColor`` reference to an ``RRGGBB`` colour.

        Args:
            theme_color: the ``w:themeColor`` value -- ``accent1``, ``text1``,
                ``hyperlink``, ...
            tint: the ``w:themeTint`` sitting next to it, if any.
            shade: the ``w:themeShade`` sitting next to it, if any.

        Returns:
            Six upper-case hexadecimal digits, or ``None`` when `theme_color` is
            ``none``, is not a ``ST_ThemeColor`` value, or names a slot this
            theme does not define.  ``None`` means "this module cannot say",
            never "black".

        Raises:
            TypeError: if `tint` or `shade` is neither a string nor an int.
            ValueError: if either is out of range, or if both are given.
        """
        slot = self.scheme_slot(theme_color)
        if slot is None:
            return None
        base = self.colors.get(slot)
        if base is None:
            return None
        return apply_tint_shade(base, tint, shade)

    def resolve_font(self, theme_font: str) -> str | None:
        """Resolve a ``ST_Theme`` reference (``minorHAnsi``) to a typeface.

        Returns:
            The typeface the theme stores for that slot -- possibly ``""``, see
            :class:`ThemeFonts` -- or ``None`` when `theme_font` is not a
            ``ST_Theme`` value.
        """
        if not isinstance(theme_font, str):
            return None
        target = _THEME_FONTS.get(theme_font.strip())
        if target is None:
            return None
        scheme, slot = target
        fonts = self.major if scheme == "major" else self.minor
        return getattr(fonts, slot)


def _slot_color(slot: etree._Element) -> str | None:
    """The ``RRGGBB`` a ``a:clrScheme`` slot holds, or ``None``.

    Two forms are read: ``a:srgbClr`` carries the value in ``val``, and
    ``a:sysClr`` -- how ``dk1`` and ``lt1`` are almost always written -- carries
    the system colour in ``val`` and the concrete one in ``lastClr``.  A slot
    expressed as ``a:schemeClr`` or ``a:prstClr`` is left unresolved rather than
    approximated.
    """
    srgb = slot.find(_A_SRGB_CLR)
    if srgb is not None:
        return _hex6(srgb.get("val"))
    system = slot.find(_A_SYS_CLR)
    if system is not None:
        return _hex6(system.get("lastClr"))
    return None


def _read_color_scheme(elements: etree._Element | None) -> dict[str, str]:
    """Slot -> ``RRGGBB`` for the ``a:clrScheme`` of `elements`."""
    if elements is None:
        return {}
    scheme = elements.find(_A_CLR_SCHEME)
    if scheme is None:
        return {}
    colors: dict[str, str] = {}
    for name in COLOR_SLOTS:
        slot = scheme.find(f"{{{A}}}{name}")
        if slot is None:
            continue
        value = _slot_color(slot)
        if value is not None:
            colors[name] = value
    return colors


def _read_fonts(scheme: etree._Element | None) -> ThemeFonts:
    """The three script slots of an ``a:majorFont`` or ``a:minorFont``."""
    if scheme is None:
        return ThemeFonts()
    slots: dict[str, str] = {}
    for attribute, tag in (("latin", "latin"), ("ea", "ea"), ("cs", "cs")):
        element = scheme.find(f"{{{A}}}{tag}")
        slots[attribute] = "" if element is None else (element.get("typeface") or "")
    return ThemeFonts(**slots)


def _theme_blob(pkg: DocxPackage) -> bytes | None:
    """The bytes of the theme part, found by relationship then by convention."""
    for relationship in pkg.document_part.rels.values():
        if relationship.reltype == _THEME_RELTYPE and not relationship.is_external:
            return relationship.target_part.blob
    part = pkg.find_part(THEME_PARTNAME)
    return None if part is None else part.blob


def _read_color_map(pkg: DocxPackage) -> Mapping[str, str]:
    """The ``w:clrSchemeMapping`` of the settings part, or the Word default.

    Only the attributes the document actually declares override the default, so
    a partial mapping stays a partial override rather than blanking the rest.
    """
    part = pkg.find_part(_SETTINGS_PARTNAME)
    if part is None:
        return DEFAULT_COLOR_MAP
    try:
        root = pkg.root_of(part)
    except PackageError:  # pragma: no cover - python-docx registers settings.xml
        return DEFAULT_COLOR_MAP
    mapping = root.find(_W_CLR_SCHEME_MAPPING)
    if mapping is None:
        return DEFAULT_COLOR_MAP
    found = dict(DEFAULT_COLOR_MAP)
    for attribute in DEFAULT_COLOR_MAP:
        value = mapping.get(qn(f"w:{attribute}"))
        if value is not None:
            found[attribute] = value
    return MappingProxyType(found)


def read_theme(pkg: DocxPackage) -> Theme | None:
    """Read the theme of `pkg`, or ``None`` when it carries none.

    A document without a theme part is legal -- Word falls back to its own
    built-in theme -- so the absence is reported as ``None`` rather than raised:
    a caller resolving a reference against ``None`` gets ``None`` back, which is
    the same "cannot say" as an unknown slot.

    The part is read through ``Part.blob`` and parsed detached: nothing this
    function returns is connected to the package, and saving `pkg` afterwards
    still reproduces ``theme1.xml`` byte for byte (D-007).

    Raises:
        PackageError: if the theme part exists but does not parse.
    """
    blob = _theme_blob(pkg)
    if blob is None:
        return None
    try:
        root = etree.fromstring(blob)
    except etree.XMLSyntaxError as exc:
        raise PackageError(f"theme part is not well-formed XML: {exc}") from exc

    elements = root.find(_A_THEME_ELEMENTS)
    font_scheme = None if elements is None else elements.find(_A_FONT_SCHEME)
    return Theme(
        name=root.get("name") or "",
        major=_read_fonts(None if font_scheme is None else font_scheme.find(_A_MAJOR_FONT)),
        minor=_read_fonts(None if font_scheme is None else font_scheme.find(_A_MINOR_FONT)),
        colors=MappingProxyType(_read_color_scheme(elements)),
        color_map=_read_color_map(pkg),
    )
