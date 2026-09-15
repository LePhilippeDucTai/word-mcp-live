"""Direct run formatting, and the envelopes a range can be put into.

This is the layer that writes ``w:rPr``.  It exists because the way the current
tools change formatting -- ``format_text`` in ``tools/format_tools.py`` -- is to
throw the runs of the range away and build new ones from the text it read back,
which silently drops everything the old runs carried: the character style, the
language, the theme font, the tracked property revision, the proofing state.
Nothing here ever creates or removes a run.  A patch is applied *onto* the
``w:rPr`` of the runs a range already resolved to, property by property, so a
property the caller did not name is left exactly as it was.

Schema order
------------
``w:rPr`` and ``w:pPr`` are ``xsd:sequence``, not ``xsd:all``: their children
have one legal order and Word repairs -- or refuses to open -- a document that
gets it wrong.  Appending is therefore never an option (the defect at
``utils/document_utils.py:440``, ``pPr.append(numPr)``, which puts ``w:numPr``
after ``w:jc``).  :data:`RPR_ORDER` and :data:`PPR_ORDER` transcribe the two
sequences from ECMA-376 (§17.3.2.28 ``CT_RPr`` and §17.3.1.26 ``CT_PPr``), and
every element this module writes is inserted at its rank, whatever order the
caller listed the properties in.  A child whose tag is not in the table is left
where it is rather than moved: an element this module has not been taught about
is not one it can position.

Themes
------
A theme reference and a direct value are two different ways of saying the same
thing, and Word gives the theme one priority: setting ``w:ascii="Consolas"`` on a
``w:rFonts`` that still carries ``w:asciiTheme="majorHAnsi"`` changes nothing on
screen.  So an explicit ``font`` clears ``asciiTheme``/``hAnsiTheme``/``cstheme``
and an explicit ``color`` clears ``themeColor``/``themeTint``/``themeShade``.
The theme attributes of the scripts this module does not write --
``eastAsiaTheme`` and the ``w:eastAsia`` font -- are left alone.

Envelopes
---------
:func:`wrap` puts a resolved range inside a hyperlink, or between the markers of
a comment.  Both need the covered runs to be contiguous siblings; when they are
not -- a range that starts outside a hyperlink and ends inside it -- there is no
single element to create, and the answer is
:class:`~word_document_server.engine.errors.UnsupportedRange` rather than an
envelope that covers more or less than what was asked for.  :func:`unwrap` is the
inverse and keeps the runs: it splices the children of the container back where
the container was, so ``unwrap(wrap(...))`` restores the paragraph.
"""

from __future__ import annotations

import string
from collections.abc import Callable, Mapping
from dataclasses import dataclass

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

from word_document_server.engine.errors import LocatorError, UnsupportedRange
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import Pieces
from word_document_server.engine.xmlns import qn

__all__ = [
    "COMMENT_REFERENCE_STYLE",
    "HIGHLIGHT_COLORS",
    "PPR_ORDER",
    "RPR_ORDER",
    "RUN_PROPERTIES",
    "UNDERLINE_STYLES",
    "Block",
    "Envelope",
    "apply_rpr",
    "character_style_exists",
    "comment",
    "hyperlink",
    "set_ppr_child",
    "unwrap",
    "wrap",
]

# --------------------------------------------------------------------------
# The two schema sequences
# --------------------------------------------------------------------------

#: Children of ``w:rPr`` in schema order (ECMA-376 §17.3.2.28).  The four
#: revision elements at the front belong to ``CT_ParaRPr`` -- the ``w:rPr`` of a
#: paragraph mark -- and are listed so that such an ``w:rPr`` is ordered
#: correctly too, even though this module only ever edits a run's own.
RPR_ORDER: tuple[str, ...] = (
    "w:ins",
    "w:del",
    "w:moveFrom",
    "w:moveTo",
    "w:rStyle",
    "w:rFonts",
    "w:b",
    "w:bCs",
    "w:i",
    "w:iCs",
    "w:caps",
    "w:smallCaps",
    "w:strike",
    "w:dstrike",
    "w:outline",
    "w:shadow",
    "w:emboss",
    "w:imprint",
    "w:noProof",
    "w:snapToGrid",
    "w:vanish",
    "w:webHidden",
    "w:color",
    "w:spacing",
    "w:w",
    "w:kern",
    "w:position",
    "w:sz",
    "w:szCs",
    "w:highlight",
    "w:u",
    "w:effect",
    "w:bdr",
    "w:shd",
    "w:fitText",
    "w:vertAlign",
    "w:rtl",
    "w:cs",
    "w:em",
    "w:lang",
    "w:eastAsianLayout",
    "w:specVanish",
    "w:oMath",
    "w:rPrChange",
)

#: Children of ``w:pPr`` in schema order (ECMA-376 §17.3.1.26).
PPR_ORDER: tuple[str, ...] = (
    "w:pStyle",
    "w:keepNext",
    "w:keepLines",
    "w:pageBreakBefore",
    "w:framePr",
    "w:widowControl",
    "w:numPr",
    "w:suppressLineNumbers",
    "w:pBdr",
    "w:shd",
    "w:tabs",
    "w:suppressAutoHyphens",
    "w:kinsoku",
    "w:wordWrap",
    "w:overflowPunct",
    "w:topLinePunct",
    "w:autoSpaceDE",
    "w:autoSpaceDN",
    "w:bidi",
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)

_RPR_RANK: dict[str, int] = {qn(tag): rank for rank, tag in enumerate(RPR_ORDER)}
_PPR_RANK: dict[str, int] = {qn(tag): rank for rank, tag in enumerate(PPR_ORDER)}

# --------------------------------------------------------------------------
# Vocabularies
# --------------------------------------------------------------------------

#: The properties :func:`apply_rpr` understands, in the order it applies them.
RUN_PROPERTIES: tuple[str, ...] = (
    "char_style",
    "font",
    "bold",
    "italic",
    "caps",
    "small_caps",
    "strike",
    "color",
    "size_pt",
    "highlight",
    "underline",
    "superscript",
    "subscript",
)

#: ``ST_Underline`` (ECMA-376 §17.18.99).  ``none`` is a value, not an absence:
#: it switches off an underline inherited from a style.
UNDERLINE_STYLES: frozenset[str] = frozenset(
    {
        "single",
        "words",
        "double",
        "thick",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashedHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dashDotHeavy",
        "dotDotDash",
        "dashDotDotHeavy",
        "wave",
        "wavyHeavy",
        "wavyDouble",
        "none",
    }
)

#: ``ST_HighlightColor`` (ECMA-376 §17.18.40).  Word's highlighter is a closed
#: list of 16 colours plus ``none``; it is not an RGB value.
HIGHLIGHT_COLORS: frozenset[str] = frozenset(
    {
        "black",
        "blue",
        "cyan",
        "darkBlue",
        "darkCyan",
        "darkGray",
        "darkGreen",
        "darkMagenta",
        "darkRed",
        "darkYellow",
        "green",
        "lightGray",
        "magenta",
        "none",
        "red",
        "white",
        "yellow",
    }
)

#: Style id Word gives the character style of a comment reference mark.
COMMENT_REFERENCE_STYLE = "CommentReference"

_TOGGLES: dict[str, str] = {
    "bold": "w:b",
    "italic": "w:i",
    "caps": "w:caps",
    "small_caps": "w:smallCaps",
    "strike": "w:strike",
}

#: Half-point bounds of ``w:sz``: Word's own limits, 0.5 pt to 1638 pt.
_MIN_HALF_POINTS = 1
_MAX_HALF_POINTS = 3276

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_RPR = qn("w:rPr")
_W_PPR = qn("w:pPr")
_W_VAL = qn("w:val")
_W_ID = qn("w:id")
_W_RFONTS = qn("w:rFonts")
_W_HYPERLINK = qn("w:hyperlink")
_W_SDT = qn("w:sdt")
_W_SDT_CONTENT = qn("w:sdtContent")

#: Attributes an explicit ``font`` writes, and the theme attributes it clears.
_FONT_ATTRIBUTES = ("w:ascii", "w:hAnsi", "w:cs")
_FONT_THEME_ATTRIBUTES = ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme")

#: Attributes an explicit ``color`` clears, on top of the value it writes.
_COLOR_THEME_ATTRIBUTES = ("w:themeColor", "w:themeTint", "w:themeShade")

#: Containers :func:`unwrap` opens.  ``w:del`` and ``w:moveFrom`` are absent on
#: purpose: their runs hold ``w:delText``, and turning that back into visible
#: text is a revision operation, owned by ``engine/revisions.py``.
_UNWRAPPABLE: frozenset[str] = frozenset(
    qn(tag)
    for tag in (
        "w:hyperlink",
        "w:ins",
        "w:moveTo",
        "w:sdt",
        "w:smartTag",
        "w:customXml",
        "w:dir",
        "w:bdo",
    )
)

#: Property children an unwrapped container leaves behind with itself.
_ENVELOPE_PROPERTIES: frozenset[str] = frozenset(
    qn(tag) for tag in ("w:sdtPr", "w:sdtEndPr", "w:smartTagPr", "w:customXmlPr")
)

#: One step of a patch: mutate the ``w:rPr`` it is handed.  Built during
#: validation, run afterwards, so a rejected value cannot leave a half-applied
#: patch behind.
_Directive = Callable[[etree._Element], None]

#: What :func:`wrap` calls: it receives the contiguous block a range resolved to
#: and returns the element identifying the envelope it built.
Envelope = Callable[["Block"], etree._Element]


# --------------------------------------------------------------------------
# Placing children in schema order
# --------------------------------------------------------------------------


def _place(parent: etree._Element, child: etree._Element, rank: dict[str, int]) -> None:
    """Move `child` to its schema position among `parent`'s children.

    Children whose tag is unknown to `rank` are ignored by the comparison: the
    element goes before the first *known* child that outranks it, and to the end
    when there is none.
    """
    position = rank[child.tag]
    for index, existing in enumerate(parent):
        if existing is child:
            continue
        other = rank.get(existing.tag)
        if other is not None and other > position:
            parent.insert(index, child)
            return


def _ensure_child(parent: etree._Element, tag: str, rank: dict[str, int]) -> etree._Element:
    """Return the first `tag` child of `parent`, creating it in order if absent."""
    found = parent.find(tag)
    if found is not None:
        return found
    child = etree.SubElement(parent, tag)
    _place(parent, child, rank)
    return child


def _drop_children(parent: etree._Element, tag: str) -> None:
    """Remove every `tag` child of `parent`."""
    for child in parent.findall(tag):
        parent.remove(child)


# --------------------------------------------------------------------------
# Reading the patch
# --------------------------------------------------------------------------


def _setter(tag: str, attributes: dict[str, str], drop: tuple[str, ...] = ()) -> _Directive:
    """A directive writing `attributes` on the `tag` child, after clearing `drop`."""

    def directive(properties: etree._Element) -> None:
        element = _ensure_child(properties, qn(tag), _RPR_RANK)
        for name in drop:
            element.attrib.pop(qn(name), None)
        for name, value in attributes.items():
            element.set(qn(name), value)

    return directive


def _remover(*tags: str) -> _Directive:
    """A directive removing every child named by `tags`."""

    def directive(properties: etree._Element) -> None:
        for tag in tags:
            _drop_children(properties, qn(tag))

    return directive


def _toggle(name: str, tag: str, value: object) -> list[_Directive]:
    """``True`` writes the bare element, ``False`` writes ``w:val="0"``."""
    if value is None:
        return [_remover(tag)]
    if not isinstance(value, bool):
        raise TypeError(f"{name} takes True, False or None, got {value!r}")
    return [_setter(tag, {} if value else {"w:val": "0"}, drop=("w:val",))]


def _underline(value: object) -> list[_Directive]:
    if value is None:
        return [_remover("w:u")]
    if isinstance(value, bool):
        style = "single" if value else "none"
    elif isinstance(value, str):
        style = value
    else:
        raise TypeError(f"underline takes True, False, a style name or None, got {value!r}")
    if style not in UNDERLINE_STYLES:
        raise ValueError(
            f"unknown underline style {style!r}; known: {', '.join(sorted(UNDERLINE_STYLES))}"
        )
    return [_setter("w:u", {"w:val": style}, drop=("w:val",))]


def _highlight(value: object) -> list[_Directive]:
    if value is None:
        return [_remover("w:highlight")]
    if not isinstance(value, str):
        raise TypeError(f"highlight takes a colour name or None, got {value!r}")
    if value not in HIGHLIGHT_COLORS:
        raise ValueError(
            f"unknown highlight colour {value!r}; "
            f"known: {', '.join(sorted(HIGHLIGHT_COLORS))}"
        )
    return [_setter("w:highlight", {"w:val": value}, drop=("w:val",))]


def _color(value: object) -> list[_Directive]:
    """``w:color`` plus the removal of the theme colour it would lose to."""
    if value is None:
        return [_remover("w:color")]
    if not isinstance(value, str):
        raise TypeError(f"color takes 'auto', an RRGGBB value or None, got {value!r}")
    raw = value.strip().removeprefix("#")
    if raw.lower() == "auto":
        setting = "auto"
    elif len(raw) == 6 and all(character in string.hexdigits for character in raw):
        setting = raw.upper()
    else:
        raise ValueError(f"color must be 'auto' or six hexadecimal digits, got {value!r}")
    return [
        _setter("w:color", {"w:val": setting}, drop=("w:val", *_COLOR_THEME_ATTRIBUTES))
    ]


def _font(value: object) -> list[_Directive]:
    """``w:rFonts`` for the ascii, high-ansi and complex-script slots."""
    if value is not None and not isinstance(value, str):
        raise TypeError(f"font takes a font name or None, got {value!r}")
    if value is not None and not value.strip():
        raise ValueError("font must be a font name, not an empty string")

    def directive(properties: etree._Element) -> None:
        if value is None:
            fonts = properties.find(_W_RFONTS)
            if fonts is None:
                return
            for name in (*_FONT_ATTRIBUTES, *_FONT_THEME_ATTRIBUTES):
                fonts.attrib.pop(qn(name), None)
            # An rFonts with nothing left says nothing; keep it only if another
            # script (eastAsia, or a hint) still uses it.
            if not fonts.attrib:
                properties.remove(fonts)
            return
        fonts = _ensure_child(properties, _W_RFONTS, _RPR_RANK)
        for name in _FONT_THEME_ATTRIBUTES:
            fonts.attrib.pop(qn(name), None)
        for name in _FONT_ATTRIBUTES:
            fonts.set(qn(name), value)

    return [directive]


def _size(value: object) -> list[_Directive]:
    """``w:sz`` and ``w:szCs``, in half-points as the schema requires."""
    if value is None:
        return [_remover("w:sz", "w:szCs")]
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"size_pt takes a number of points or None, got {value!r}")
    doubled = value * 2
    if doubled != int(doubled):
        raise ValueError(
            f"size_pt is stored in half-points and must be a multiple of 0.5, got {value!r}"
        )
    half_points = int(doubled)
    if not _MIN_HALF_POINTS <= half_points <= _MAX_HALF_POINTS:
        raise ValueError(
            f"size_pt must be between {_MIN_HALF_POINTS / 2} and {_MAX_HALF_POINTS / 2} points, "
            f"got {value!r}"
        )
    return [
        _setter("w:sz", {"w:val": str(half_points)}, drop=("w:val",)),
        _setter("w:szCs", {"w:val": str(half_points)}, drop=("w:val",)),
    ]


def _vertical_alignment(patch: Mapping[str, object]) -> list[_Directive]:
    """``w:vertAlign``, which ``superscript`` and ``subscript`` share.

    One element, two keys: the two properties are not independent, so they are
    read together.  Asking for both at once is a contradiction and is refused;
    a ``False`` next to the other's ``True`` is not (it says "not that one"),
    and the ``True`` wins.
    """
    if "superscript" not in patch and "subscript" not in patch:
        return []
    values: dict[str, object] = {
        name: patch.get(name) for name in ("superscript", "subscript") if name in patch
    }
    for name, value in values.items():
        if value is not None and not isinstance(value, bool):
            raise TypeError(f"{name} takes True, False or None, got {value!r}")
    if values.get("superscript") and values.get("subscript"):
        raise ValueError("superscript and subscript cannot both be requested")
    if values.get("superscript"):
        setting = "superscript"
    elif values.get("subscript"):
        setting = "subscript"
    elif any(value is False for value in values.values()):
        setting = "baseline"
    else:
        return [_remover("w:vertAlign")]
    return [_setter("w:vertAlign", {"w:val": setting}, drop=("w:val",))]


def _char_style(value: object, pkg: DocxPackage | None) -> list[_Directive]:
    if value is None:
        return [_remover("w:rStyle")]
    if not isinstance(value, str):
        raise TypeError(f"char_style takes a style id or None, got {value!r}")
    if not value.strip():
        raise ValueError("char_style must be a style id, not an empty string")
    if pkg is None:
        raise ValueError(
            "char_style needs the package the runs belong to, to check the style exists: "
            "call apply_rpr(..., pkg=package)"
        )
    _require_character_style(pkg, value)
    return [_setter("w:rStyle", {"w:val": value}, drop=("w:val",))]


def _parse_patch(patch: Mapping[str, object], pkg: DocxPackage | None) -> list[_Directive]:
    """Turn a patch into directives, rejecting every bad value before any write."""
    if not isinstance(patch, Mapping):
        raise TypeError(f"patch must be a mapping of property names, got {type(patch).__name__}")
    unknown = sorted(set(patch) - set(RUN_PROPERTIES))
    if unknown:
        raise ValueError(
            f"unknown run properties: {', '.join(unknown)}; "
            f"known: {', '.join(sorted(RUN_PROPERTIES))}"
        )
    directives: list[_Directive] = []
    for name in RUN_PROPERTIES:
        if name in ("superscript", "subscript"):
            continue
        if name not in patch:
            continue
        value = patch[name]
        if name in _TOGGLES:
            directives += _toggle(name, _TOGGLES[name], value)
        elif name == "underline":
            directives += _underline(value)
        elif name == "highlight":
            directives += _highlight(value)
        elif name == "color":
            directives += _color(value)
        elif name == "font":
            directives += _font(value)
        elif name == "size_pt":
            directives += _size(value)
        else:  # char_style
            directives += _char_style(value, pkg)
    return directives + _vertical_alignment(patch)


# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------


def _styles_root(pkg: DocxPackage) -> etree._Element | None:
    """Root of the styles part, or ``None`` when the package has none."""
    try:
        part = pkg.document_part.part_related_by(RT.STYLES)
    except KeyError:
        part = pkg.find_part("/word/styles.xml")
    if part is None:
        return None
    return pkg.root_of(part)


def _style_type(pkg: DocxPackage, style_id: str) -> str | None:
    """The ``w:type`` of the style `style_id`, or ``None`` if it is not defined."""
    root = _styles_root(pkg)
    if root is None:
        return None
    for style in root.iter(qn("w:style")):
        if style.get(qn("w:styleId")) == style_id:
            return style.get(qn("w:type")) or ""
    return None


def character_style_exists(pkg: DocxPackage, style_id: str) -> bool:
    """Whether `pkg` defines `style_id` as a character style."""
    return _style_type(pkg, style_id) == "character"


def _require_character_style(pkg: DocxPackage, style_id: str) -> None:
    """Raise unless `style_id` is a character style of `pkg`.

    A ``w:rStyle`` pointing at a style the package does not define is ignored by
    Word and lost by the next round trip, and one pointing at a *paragraph*
    style is a different kind of mistake with the same silent outcome.  Both are
    refused here rather than written and forgotten.

    Raises:
        LocatorError: ``unknown-style`` if no such style exists,
            ``wrong-style-type`` if it exists but is not a character style.
    """
    kind = _style_type(pkg, style_id)
    if kind is None:
        raise LocatorError(
            "unknown-style", f"styles.xml defines no style with the id {style_id!r}"
        )
    if kind != "character":
        raise LocatorError(
            "wrong-style-type",
            f"style {style_id!r} is a {kind or 'untyped'} style; w:rStyle needs a character style",
        )


# --------------------------------------------------------------------------
# Applying a patch
# --------------------------------------------------------------------------


def _run_properties(run: etree._Element) -> etree._Element:
    """Return the run's ``w:rPr``, creating it as its first child if absent."""
    properties = run.find(_W_RPR)
    if properties is None:
        properties = etree.SubElement(run, _W_RPR)
        run.insert(0, properties)
    return properties


def apply_rpr(
    pieces: Pieces, patch: Mapping[str, object], *, pkg: DocxPackage | None = None
) -> tuple[etree._Element, ...]:
    """Apply `patch` to the run properties of every run `pieces` covers.

    The patch is read as a set of *changes*, not as a description of the final
    formatting: a property the patch does not mention is left untouched, a
    property mapped to ``None`` is removed (so what the paragraph style says
    applies again), and ``False`` on a toggle writes an explicit ``w:val="0"``,
    which is how a run switches off something a style turned on.

    Accepted keys, all optional -- see :data:`RUN_PROPERTIES`:

    ``bold``, ``italic``, ``caps``, ``small_caps``, ``strike``
        ``True`` / ``False`` / ``None``.
    ``underline``
        ``True`` (``single``), ``False`` (``none``), any of
        :data:`UNDERLINE_STYLES`, or ``None``.
    ``superscript``, ``subscript``
        ``True`` / ``False`` / ``None``.  They share one ``w:vertAlign``, so
        asking for both at once raises.
    ``size_pt``
        points, a multiple of ``0.5`` -- ``w:sz`` counts half-points and cannot
        store anything else.  Written to ``w:szCs`` too, as Word does.
    ``font``
        a font name, written to the ascii, high-ansi and complex-script slots of
        ``w:rFonts``; the matching theme fonts are cleared.
    ``color``
        ``"auto"`` or ``RRGGBB`` (a leading ``#`` is accepted); the theme colour
        attributes are cleared.
    ``highlight``
        one of :data:`HIGHLIGHT_COLORS`.
    ``char_style``
        the id of a character style, which `pkg` must define.

    An ``w:rPr`` left with no children at all is removed, so applying a property
    and then removing it gives the run back exactly as it was.

    Args:
        pieces: what :func:`word_document_server.engine.ranges.resolve` returned.
            An empty range covers no run and is a no-op.
        patch: the properties to change.
        pkg: the package the runs belong to; needed only for ``char_style``.

    Returns:
        The runs that were patched, in document order.

    Raises:
        TypeError: if `pieces` is not a :class:`~...ranges.Pieces`, if `patch` is
            not a mapping, or if a value is of the wrong type.
        ValueError: on an unknown property name or a value the schema cannot
            store.  Every value is checked before the first write, so a refused
            patch leaves the runs exactly as they were.
        LocatorError: ``unknown-style`` / ``wrong-style-type`` for a
            ``char_style`` the package does not define as a character style.
    """
    if not isinstance(pieces, Pieces):
        raise TypeError(f"expected the Pieces of a resolved range, got {type(pieces).__name__}")
    directives = _parse_patch(patch, pkg)
    if not directives:
        return pieces.runs
    for run in pieces.runs:
        properties = _run_properties(run)
        for directive in directives:
            directive(properties)
        if len(properties) == 0 and not properties.attrib:
            run.remove(properties)
    return pieces.runs


# --------------------------------------------------------------------------
# Paragraph properties
# --------------------------------------------------------------------------


def _paragraph_element(paragraph: object) -> etree._Element:
    """Coerce `paragraph` to a live ``w:p`` element.

    Accepts the element itself or a python-docx ``Paragraph``, like
    :mod:`word_document_server.engine.textmodel` and
    :mod:`word_document_server.engine.ranges`.

    Raises:
        TypeError: if `paragraph` is neither.
    """
    element = paragraph
    if not isinstance(element, etree._Element):
        for attribute in ("_p", "_element"):
            candidate = getattr(element, attribute, None)
            if isinstance(candidate, etree._Element):
                element = candidate
                break
    if not isinstance(element, etree._Element):
        raise TypeError(
            f"expected a w:p element or a python-docx paragraph, got {type(paragraph).__name__}"
        )
    if element.tag != _W_P:
        raise TypeError(f"expected a w:p element, got {element.tag!r}")
    return element


def set_ppr_child(
    paragraph: object, tag: str, attrs: Mapping[str, str] | None = None
) -> etree._Element | None:
    """Set -- or remove -- one child of the paragraph's ``w:pPr``.

    The child is created at its rank in :data:`PPR_ORDER`, and ``w:pPr`` itself
    at the front of the paragraph, which is where the schema puts it.

    ``attrs`` is the whole attribute set of the element: an empty mapping means
    "this child, with no attribute" (``w:numPr``, ``w:keepNext``), and ``None``
    means "remove it".  The *children* of an element that already exists are
    kept, so re-setting ``w:numPr`` does not drop its ``w:ilvl`` and ``w:numId``.
    Removing the last child of a ``w:pPr`` removes the ``w:pPr`` too, so setting
    a property and removing it leaves the paragraph as it was found.

    Args:
        paragraph: a ``w:p`` element or a python-docx paragraph.
        tag: a prefixed name from :data:`PPR_ORDER`, such as ``"w:jc"``.
        attrs: prefixed attribute names to values, or ``None`` to remove.

    Returns:
        The element, or ``None`` when it was removed.

    Raises:
        TypeError: if `paragraph` is not a paragraph.
        ValueError: if `tag` is not a child ``w:pPr`` may hold.
    """
    element = _paragraph_element(paragraph)
    if tag not in PPR_ORDER:
        raise ValueError(f"{tag!r} is not a child of w:pPr; known: {', '.join(PPR_ORDER)}")
    qualified = qn(tag)

    properties = element.find(_W_PPR)
    if attrs is None:
        if properties is None:
            return None
        _drop_children(properties, qualified)
        if len(properties) == 0 and not properties.attrib:
            element.remove(properties)
        return None

    if properties is None:
        properties = etree.SubElement(element, _W_PPR)
        element.insert(0, properties)
    child = _ensure_child(properties, qualified, _PPR_RANK)
    for name in list(child.attrib):
        del child.attrib[name]
    for name, value in attrs.items():
        child.set(qn(name), value)
    return child


# --------------------------------------------------------------------------
# Envelopes
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Block:
    """The contiguous stretch of siblings an envelope is built around.

    Attributes:
        parent: the element holding the block -- a ``w:p``, or a container of
            one such as an existing ``w:hyperlink`` or ``w:ins``.
        index: index of the first element of the block among `parent`'s
            children, read just before the envelope is built.
        elements: the block itself, in document order: the runs the range covers
            and the markers that sit between them.
    """

    parent: etree._Element
    index: int
    elements: tuple[etree._Element, ...]


def _block(pieces: Pieces) -> Block:
    """The block `pieces` covers, or a refusal.

    Raises:
        UnsupportedRange: ``empty-range`` when the range covers no run,
            ``crosses-container`` when the covered runs do not share a parent,
            ``discontiguous-runs`` when something the range does not cover sits
            between them.
    """
    runs = pieces.runs
    if not runs:
        raise UnsupportedRange(
            "empty-range",
            f"range {pieces.start}..{pieces.end} covers no run to put in an envelope",
        )
    parents = {id(run.getparent()): run.getparent() for run in runs}
    if len(parents) != 1:
        raise UnsupportedRange(
            "crosses-container",
            f"range {pieces.start}..{pieces.end} spans several containers "
            "(a hyperlink, a tracked insertion or a content control boundary); "
            "an envelope must wrap siblings",
        )
    parent = next(iter(parents.values()))
    if parent is None:
        raise UnsupportedRange("crosses-container", "the covered runs have no parent")

    positions = [parent.index(run) for run in runs]
    first, last = min(positions), max(positions)
    covered = {id(element) for element in (*runs, *pieces.markers)}
    block = list(parent)[first : last + 1]
    for element in block:
        if id(element) not in covered:
            name = etree.QName(element).localname
            raise UnsupportedRange(
                "discontiguous-runs",
                f"range {pieces.start}..{pieces.end} leaves <{name}> between the runs it "
                "covers; an envelope would take in content the range does not cover",
            )
    return Block(parent=parent, index=first, elements=tuple(block))


def wrap(pieces: Pieces, factory: Envelope) -> etree._Element:
    """Put the runs `pieces` covers inside the envelope `factory` builds.

    The range must resolve to contiguous siblings: an envelope is one element
    (or one pair of markers) around one stretch of content, and a range that
    starts outside a hyperlink and ends inside it has no such stretch.

    Args:
        pieces: what :func:`word_document_server.engine.ranges.resolve` returned.
        factory: :func:`hyperlink` or :func:`comment`, or any callable taking the
            :class:`Block` and returning the element it created.

    Returns:
        Whatever `factory` returns: the ``w:hyperlink`` for :func:`hyperlink`,
        the ``w:commentRangeStart`` for :func:`comment`.

    Raises:
        TypeError: if `pieces` is not a :class:`~...ranges.Pieces`.
        UnsupportedRange: see :func:`_block`, plus ``nested-hyperlink``.
    """
    if not isinstance(pieces, Pieces):
        raise TypeError(f"expected the Pieces of a resolved range, got {type(pieces).__name__}")
    return factory(_block(pieces))


def _move_into(block: Block, container: etree._Element) -> None:
    """Move the block's elements into `container`, then `container` into place."""
    for element in block.elements:
        container.append(element)
    # The index is read from the caller's Block, taken before the elements were
    # detached: everything before the block kept its position, so it still holds.
    block.parent.insert(block.index, container)


def hyperlink(
    rId: str | None = None,
    *,
    anchor: str | None = None,
    tooltip: str | None = None,
    history: bool = True,
) -> Envelope:
    """An envelope wrapping the range in a ``w:hyperlink``.

    `rId` points at a relationship of the *part the paragraph lives in* -- get
    one from :meth:`~word_document_server.engine.package.DocxPackage.add_external_rel`
    -- and `anchor` names a bookmark of the document instead.  At least one of
    the two is required; passing both is how Word links to a fragment of an
    external target.

    Raises:
        ValueError: if neither `rId` nor `anchor` is given.
    """
    if not rId and not anchor:
        raise ValueError("a hyperlink needs a relationship id, an anchor, or both")

    def build(block: Block) -> etree._Element:
        if block.parent.tag == _W_HYPERLINK or any(
            ancestor.tag == _W_HYPERLINK for ancestor in block.parent.iterancestors(_W_HYPERLINK)
        ):
            raise UnsupportedRange(
                "nested-hyperlink", "a hyperlink cannot be nested inside another hyperlink"
            )
        link = etree.SubElement(block.parent, _W_HYPERLINK)
        if rId:
            link.set(qn("r:id"), rId)
        if anchor:
            link.set(qn("w:anchor"), anchor)
        if tooltip:
            link.set(qn("w:tooltip"), tooltip)
        if history:
            link.set(qn("w:history"), "1")
        _move_into(block, link)
        return link

    return build


def comment(comment_id: int | str, *, pkg: DocxPackage | None = None) -> Envelope:
    """An envelope anchoring the comment `comment_id` on the range.

    A comment is not a container: it is a ``w:commentRangeStart`` before the
    range, a ``w:commentRangeEnd`` after it, and a run holding the
    ``w:commentReference`` mark -- the three carrying the same id, which is what
    ``word/comments.xml`` keys the comment text on.  Nothing is moved, so a
    comment can be anchored on a range that is already inside a hyperlink or a
    tracked insertion.

    The reference run takes the ``CommentReference`` character style when `pkg`
    is given and defines it, since that is what makes Word draw the mark at the
    right size; a package without the style gets an unstyled reference rather
    than a dangling ``w:rStyle``.

    Raises:
        ValueError: if `comment_id` is not a decimal id.
    """
    try:
        identifier = str(int(comment_id))
    except (TypeError, ValueError):
        raise ValueError(f"a comment id must be a decimal number, got {comment_id!r}") from None
    styled = pkg is not None and character_style_exists(pkg, COMMENT_REFERENCE_STYLE)

    def build(block: Block) -> etree._Element:
        parent = block.parent
        start = etree.SubElement(parent, qn("w:commentRangeStart"))
        start.set(_W_ID, identifier)
        end = etree.SubElement(parent, qn("w:commentRangeEnd"))
        end.set(_W_ID, identifier)
        run = etree.SubElement(parent, _W_R)
        if styled:
            properties = etree.SubElement(run, _W_RPR)
            etree.SubElement(properties, qn("w:rStyle")).set(_W_VAL, COMMENT_REFERENCE_STYLE)
        etree.SubElement(run, qn("w:commentReference")).set(_W_ID, identifier)

        after = block.index + len(block.elements)
        parent.insert(block.index, start)
        parent.insert(after + 1, end)
        parent.insert(after + 2, run)
        return start

    return build


def unwrap(element: object) -> tuple[etree._Element, ...]:
    """Replace a container by its content, keeping the runs where they were.

    The inverse of :func:`wrap` for a container envelope: the children of
    `element` -- the children of its ``w:sdtContent`` for a content control --
    are spliced back into its parent at its position, in order, and `element`
    itself, with the property children it owns (``w:sdtPr``, ``w:smartTagPr``,
    ...), disappears.  The runs are the very same elements, so their text,
    properties and markers are untouched by definition.

    Returns:
        The elements that took the container's place, in document order.

    Raises:
        TypeError: if `element` is not an element, or has no parent.
        UnsupportedRange: ``unsupported-envelope`` for an element that is not a
            container this layer opens -- ``w:del`` and ``w:moveFrom`` in
            particular, whose content is hidden text that only
            ``engine/revisions.py`` may bring back.
    """
    if not isinstance(element, etree._Element):
        raise TypeError(f"expected an element, got {type(element).__name__}")
    if element.tag not in _UNWRAPPABLE:
        name = etree.QName(element).localname
        raise UnsupportedRange(
            "unsupported-envelope",
            f"<{name}> is not a container this layer can open without changing what "
            "the document says",
        )
    parent = element.getparent()
    if parent is None:
        raise TypeError("cannot unwrap an element that has no parent")

    holder = element
    if element.tag == _W_SDT:
        content = element.find(_W_SDT_CONTENT)
        if content is not None:
            holder = content
    moved = [child for child in holder if child.tag not in _ENVELOPE_PROPERTIES]

    index = parent.index(element)
    for offset, child in enumerate(moved):
        parent.insert(index + offset, child)
    parent.remove(element)
    return tuple(moved)
