"""The style sheet: what a style is, what it inherits, where it is used, how it
is written.

Direct formatting is only half of what a run looks like; the other half is the
style it names and everything that style inherits.  This module is that half.
:mod:`word_document_server.engine.format` writes ``w:rPr`` on runs; this one
writes ``w:style`` -- the definition every one of those runs may point at.

Three questions
---------------
:func:`list_styles`
    what styles does this document define?  One :class:`StyleInfo` per
    ``w:style``, with the metadata Word's own style pane shows: the family, the
    ``basedOn``/``next``/``link`` wiring, the UI priority and the three flags
    that decide whether a style is offered to a human.
:func:`get_style`
    what does *this* style actually say?  The style's own decoded properties,
    the same for every ancestor of its ``basedOn`` chain, and the `resolved`
    merge of all of them on top of ``w:docDefaults``.
:func:`find_style_usage`
    where is it used?  Every paragraph, run and table that names it, each with
    the locator that reaches it.

Decoded, not raw
----------------
Properties come back in the vocabulary
:func:`~word_document_server.engine.format.apply_rpr` takes -- ``bold``,
``size_pt``, ``font``, ``color``, ``alignment``, ``indent`` -- so what is read
here can be handed back there without a translation table in between.  Only a
property the minimal model covers is decoded: identity, inheritance, font,
paragraph layout, the numbering link and the metadata.  A ``w:rPr`` child
outside it (borders, effects, east-asian typography) is *not* reported, which is
why a decoded property set is a reading of the style and never a replacement for
the XML.

Theme references survive
------------------------
A style that says ``w:asciiTheme="minorHAnsi"`` is not saying "Aptos"; it is
saying "whatever the theme calls minorHAnsi".  Flattening the two would make a
reader unable to tell a style that pins a font from one that follows the theme,
so a themed property is reported as ``{"value": ..., "theme": ...}`` with both
halves: `value` is the literal the document caches, `theme` the reference.
:mod:`word_document_server.engine.theme` resolves the reference when a caller
wants the concrete font or colour.

Resolution order
----------------
``resolved`` merges, in this order and each step overriding the previous:
``w:docDefaults``, then the ``basedOn`` chain from its root down, then the style
itself.  That is the order Word applies, minus the two things this layer cannot
know: the numbering level's own properties and the direct formatting of the run.
So ``resolved`` is what the style contributes, not what the reader sees.

A ``basedOn`` chain that loops -- legal to write, fatal to follow -- stops at the
first style it revisits rather than recursing, and the cycle is reported in
:attr:`StyleDetail.warnings`.

Writing
-------
:func:`create_style`, :func:`update_style`, :func:`clone_style` and
:func:`delete_style` take and return the vocabulary the readers above report, so
what :func:`get_style` hands out can be handed straight back.  Four properties
of the writers are worth stating up front, because they are what makes editing a
style sheet safe:

*The schema order is not this module's to invent.*
    ``w:style``, ``w:rPr`` and ``w:pPr`` are ``xsd:sequence``: Word repairs -- or
    refuses -- a document whose children are out of order.  :data:`STYLE_ORDER`
    transcribes ``CT_Style`` (ECMA-376 §17.7.4.17); the other two are
    :data:`~word_document_server.engine.format.RPR_ORDER` and
    :data:`~word_document_server.engine.format.PPR_ORDER`, reused from the
    formatting layer rather than restated, so there is one transcription of each
    sequence in the engine and not two.

*A reference is checked before it is written.*
    ``basedOn``, ``next`` and ``link`` may name a style by id or by name, and are
    stored as the id; one that names nothing is refused rather than written, so
    the writers cannot produce the dangling chain :func:`get_style` warns about.

*A theme reference stays a theme reference.*
    ``run_props`` accepts the ``{"value": ..., "theme": ...}`` form
    :func:`get_style` reports, and writes both halves back as they were.  A
    colour resolved against the theme is never written into the style sheet
    (D-028): what goes in is the reference the document already carried.

*A style in use is not deleted by surprise.*
    :func:`delete_style` refuses a style anything points at unless the caller
    says what to reassign those places to, and repoints the ``basedOn``/``next``/
    ``link`` of the other styles itself -- the two halves of "delete" that a
    caller cannot see from the outside.
"""

from __future__ import annotations

import copy
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from docx.oxml.parser import parse_xml
from lxml import etree

from word_document_server.engine.errors import LocatorError, PackageError
from word_document_server.engine.format import (
    _PPR_RANK,
    _RPR_RANK,
    RUN_PROPERTIES,
    _drop_children,
    _ensure_child,
    _parse_patch,
)
from word_document_server.engine.locators import _tables, indexed_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import segments, visible_text
from word_document_server.engine.xmlns import W, qn

if TYPE_CHECKING:  # pragma: no cover - typing only
    from docx.document import Document

__all__ = [
    "ALIGNMENTS",
    "LINE_RULES",
    "PARAGRAPH_PROPERTIES",
    "STYLES_PARTNAME",
    "STYLE_ORDER",
    "STYLE_SPEC_KEYS",
    "USAGE_PREVIEW",
    "WRITABLE_FAMILIES",
    "StyleDeletion",
    "StyleDetail",
    "StyleInfo",
    "StyleLevel",
    "StyleUsage",
    "clone_style",
    "create_style",
    "delete_style",
    "derive_style_id",
    "find_style_usage",
    "get_style",
    "list_styles",
    "style_info",
    "styles_root",
    "update_style",
]

#: Part name of the style sheet.
STYLES_PARTNAME = "/word/styles.xml"

#: Characters of text a :class:`StyleUsage` quotes.  A usage report is a map of
#: where a style is used, not a copy of the text it is used on.
USAGE_PREVIEW = 80

#: The four values of ``w:type`` on ``w:style``.
STYLE_FAMILIES = ("paragraph", "character", "table", "numbering")

_W_STYLES = qn("w:styles")
_W_STYLE = qn("w:style")
_W_STYLE_ID = qn("w:styleId")
_W_TYPE = qn("w:type")
_W_DEFAULT = qn("w:default")
_W_CUSTOM_STYLE = qn("w:customStyle")
_W_NAME = qn("w:name")
_W_BASED_ON = qn("w:basedOn")
_W_NEXT = qn("w:next")
_W_LINK = qn("w:link")
_W_UI_PRIORITY = qn("w:uiPriority")
_W_Q_FORMAT = qn("w:qFormat")
_W_HIDDEN = qn("w:hidden")
_W_SEMI_HIDDEN = qn("w:semiHidden")
_W_VAL = qn("w:val")

_W_DOC_DEFAULTS = qn("w:docDefaults")
_W_RPR_DEFAULT = qn("w:rPrDefault")
_W_PPR_DEFAULT = qn("w:pPrDefault")
_W_RPR = qn("w:rPr")
_W_PPR = qn("w:pPr")

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_TR = qn("w:tr")
_W_TC = qn("w:tc")
_W_TBL_PR = qn("w:tblPr")
_W_TBL_STYLE = qn("w:tblStyle")
_W_PSTYLE = qn("w:pStyle")
_W_RSTYLE = qn("w:rStyle")

_W_RFONTS = qn("w:rFonts")
_W_COLOR = qn("w:color")
_W_SZ = qn("w:sz")
_W_U = qn("w:u")
_W_HIGHLIGHT = qn("w:highlight")
_W_VERT_ALIGN = qn("w:vertAlign")

_W_JC = qn("w:jc")
_W_IND = qn("w:ind")
_W_SPACING = qn("w:spacing")
_W_OUTLINE_LVL = qn("w:outlineLvl")
_W_NUMPR = qn("w:numPr")
_W_NUM_ID = qn("w:numId")
_W_ILVL = qn("w:ilvl")

#: Decoded name -> ``w:rPr`` child, for the properties that are a bare on/off
#: toggle.  Same names as the ``patch`` of
#: :func:`~word_document_server.engine.format.apply_rpr`.
_RUN_TOGGLES = {
    "bold": qn("w:b"),
    "italic": qn("w:i"),
    "caps": qn("w:caps"),
    "small_caps": qn("w:smallCaps"),
    "strike": qn("w:strike"),
}

#: Decoded name -> ``w:pPr`` child, for the paragraph toggles.
_PARAGRAPH_TOGGLES = {
    "keep_next": qn("w:keepNext"),
    "keep_lines": qn("w:keepLines"),
    "page_break_before": qn("w:pageBreakBefore"),
    "widow_control": qn("w:widowControl"),
    "contextual_spacing": qn("w:contextualSpacing"),
}

#: ``w:ind`` attribute -> decoded name.  Twips throughout, like
#: :mod:`word_document_server.engine.inspect`: the unit the package stores.
_INDENT_ATTRIBUTES = {
    "left": "left",
    "start": "left",
    "right": "right",
    "end": "right",
    "firstLine": "first_line",
    "hanging": "hanging",
}

#: ``w:spacing`` attribute -> decoded name.
_SPACING_ATTRIBUTES = {
    "before": "before",
    "after": "after",
    "line": "line",
    "lineRule": "line_rule",
}

#: What ``w:val`` spells when it means "off".  Word writes all four.
_OFF = frozenset({"0", "false", "off"})

#: Decoded properties whose parts are inherited one by one rather than as a
#: block.  ``w:ind``, ``w:spacing`` and ``w:numPr`` each hold several
#: independent settings, so a style that sets ``w:spacing w:after`` keeps the
#: line spacing it inherits instead of clearing it.  Everything else -- ``font``,
#: ``color`` -- is one property expressed over several attributes, and a level
#: that states it replaces it whole.
_MERGED_PROPERTIES = frozenset({"indent", "spacing", "numbering"})


# --------------------------------------------------------------------------------------
# Reading the part
# --------------------------------------------------------------------------------------


def _styles_root(pkg: DocxPackage) -> etree._Element | None:
    """The live ``w:styles`` root, or ``None`` when the package has no style sheet."""
    part = pkg.find_part(STYLES_PARTNAME)
    if part is None:
        return None
    try:
        root = pkg.root_of(part)
    except PackageError:  # pragma: no cover - styles.xml is in LIVE_CONTENT_TYPES
        return None
    return root if root.tag == _W_STYLES else None


def _flag(element: etree._Element | None, tag: str) -> bool:
    """Whether the toggle child `tag` of `element` is on.

    A bare ``<w:qFormat/>`` is on; ``<w:qFormat w:val="0"/>`` is off.  An absent
    child is off, which is the schema's own default for all three flags here.
    """
    if element is None:
        return False
    child = element.find(tag)
    if child is None:
        return False
    return child.get(_W_VAL) not in _OFF


def _child_val(element: etree._Element | None, tag: str) -> str | None:
    """The ``w:val`` of the child `tag`, or ``None`` when there is no such child."""
    if element is None:
        return None
    child = element.find(tag)
    return None if child is None else child.get(_W_VAL)


def _int_val(element: etree._Element | None, tag: str) -> int | None:
    """The ``w:val`` of the child `tag` read as an int, or ``None``."""
    return _as_int(_child_val(element, tag))


def _as_int(value: str | None) -> int | None:
    """An attribute read as an int, or ``None`` when absent or not a number."""
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


# --------------------------------------------------------------------------------------
# Identity and metadata
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class StyleInfo:
    """One ``w:style``: who it is, what it inherits, how Word offers it.

    Attributes:
        style_id: the ``w:styleId``, which is what ``w:pStyle`` and ``w:rStyle``
            reference and what every function here takes.
        name: the ``w:name``, which is what Word's style pane shows.  Falls back
            to `style_id` when the style declares none.
        family: ``w:type`` -- one of :data:`STYLE_FAMILIES`, or the raw value for
            a document that writes something else.
        builtin: the *declared* status: ``w:customStyle="1"`` makes a style
            custom, its absence makes it built-in.  This reads the flag, it does
            not judge the name: a hand-written style that omits the flag is
            reported as built-in because that is what the document claims.
        default: whether ``w:default="1"`` makes it the default of its family.
        based_on: ``w:basedOn``, the style this one inherits from.
        next_style: ``w:next``, the style Word applies to the paragraph typed
            after one in this style.  Named with a suffix because ``next`` is a
            builtin.
        link: ``w:link``, the paragraph style a character style pairs with (or
            the other way round).
        ui_priority: ``w:uiPriority``, the sort key of Word's style gallery.
        q_format: ``w:qFormat``, whether the style is offered in that gallery.
        hidden: ``w:hidden``, whether Word hides it from the user entirely.
        semi_hidden: ``w:semiHidden``, whether it is hidden until used.
    """

    style_id: str
    name: str
    family: str
    builtin: bool
    default: bool
    based_on: str | None
    next_style: str | None
    link: str | None
    ui_priority: int | None
    q_format: bool
    hidden: bool
    semi_hidden: bool


def _style_info(style: etree._Element) -> StyleInfo:
    """Read one ``w:style`` element into a :class:`StyleInfo`."""
    style_id = style.get(_W_STYLE_ID) or ""
    return StyleInfo(
        style_id=style_id,
        name=_child_val(style, _W_NAME) or style_id,
        family=style.get(_W_TYPE) or "paragraph",
        builtin=style.get(_W_CUSTOM_STYLE) in (None, *_OFF),
        default=style.get(_W_DEFAULT) not in (None, *_OFF),
        based_on=_child_val(style, _W_BASED_ON),
        next_style=_child_val(style, _W_NEXT),
        link=_child_val(style, _W_LINK),
        ui_priority=_int_val(style, _W_UI_PRIORITY),
        q_format=_flag(style, _W_Q_FORMAT),
        hidden=_flag(style, _W_HIDDEN),
        semi_hidden=_flag(style, _W_SEMI_HIDDEN),
    )


def _styles_in(root: etree._Element | None) -> list[etree._Element]:
    """Every ``w:style`` of a style sheet root, in document order.

    A ``w:style`` without a ``w:styleId`` is skipped: nothing can reference it,
    so reporting it would only offer an id no caller can use.
    """
    if root is None:
        return []
    return [style for style in root.findall(_W_STYLE) if style.get(_W_STYLE_ID)]


def _style_elements(pkg: DocxPackage) -> list[etree._Element]:
    """Every ``w:style`` the package defines, in document order."""
    return _styles_in(_styles_root(pkg))


def list_styles(pkg: DocxPackage, family: str | None = None) -> list[StyleInfo]:
    """Every style the document defines, in the order the style sheet lists them.

    Args:
        pkg: the package to read.
        family: keep only this ``w:type`` -- ``"paragraph"``, ``"character"``,
            ``"table"`` or ``"numbering"``.  ``None`` keeps them all.

    Returns:
        One :class:`StyleInfo` per ``w:style``.  Empty when the package has no
        style sheet, which is legal and means the document leans entirely on
        Word's own defaults.

    Raises:
        ValueError: if `family` is not one of :data:`STYLE_FAMILIES`.  A typo
            silently matching nothing would read as "this document defines no
            character style", which is a different answer.
    """
    if family is not None and family not in STYLE_FAMILIES:
        raise ValueError(
            f"unknown style family {family!r}; known: {', '.join(STYLE_FAMILIES)}"
        )
    found = [_style_info(style) for style in _style_elements(pkg)]
    if family is None:
        return found
    return [info for info in found if info.family == family]


# --------------------------------------------------------------------------------------
# Decoding properties
# --------------------------------------------------------------------------------------


def _themed(value: str | None, theme: str | None, **extra: Any) -> dict[str, Any] | None:
    """A ``{"value", "theme"}`` property, or ``None`` when neither half is set.

    See the module docstring: the literal and the theme reference are kept side
    by side rather than collapsed, because only the pair says whether the style
    pins the property or follows the theme.
    """
    if value is None and theme is None and not any(v is not None for v in extra.values()):
        return None
    return {"value": value, "theme": theme, **extra}


def _toggle_value(properties: etree._Element, tag: str) -> bool | None:
    """A toggle child read as ``True``/``False``, or ``None`` when it is absent.

    The three states matter here: absent means "inherit", which is not the same
    as the explicit ``w:val="0"`` a style writes to *cancel* an inherited bold.
    """
    child = properties.find(tag)
    if child is None:
        return None
    return child.get(_W_VAL) not in _OFF


def decode_rpr(properties: etree._Element | None) -> dict[str, Any]:
    """Decode a ``w:rPr`` into the vocabulary ``apply_rpr`` takes.

    Only a property the minimal model covers appears, and only when the element
    sets it: an absent key means "this level says nothing", which is what makes
    the merge in :func:`get_style` meaningful.
    """
    if properties is None:
        return {}
    decoded: dict[str, Any] = {}
    for name, tag in _RUN_TOGGLES.items():
        value = _toggle_value(properties, tag)
        if value is not None:
            decoded[name] = value

    underline = properties.find(_W_U)
    if underline is not None:
        decoded["underline"] = underline.get(_W_VAL)

    half_points = _as_int(_child_val(properties, _W_SZ))
    if half_points is not None:
        decoded["size_pt"] = half_points / 2

    highlight = _child_val(properties, _W_HIGHLIGHT)
    if highlight is not None:
        decoded["highlight"] = highlight

    vertical = _child_val(properties, _W_VERT_ALIGN)
    if vertical is not None:
        decoded["superscript"] = vertical == "superscript"
        decoded["subscript"] = vertical == "subscript"

    style = _child_val(properties, _W_RSTYLE)
    if style is not None:
        decoded["char_style"] = style

    fonts = properties.find(_W_RFONTS)
    if fonts is not None:
        # ascii and hAnsi address the same latin text and are written together;
        # reading one as "the font" and reporting the other separately would
        # invent a distinction no producer makes.
        latin = _themed(
            fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi")),
            fonts.get(qn("w:asciiTheme")) or fonts.get(qn("w:hAnsiTheme")),
        )
        if latin is not None:
            decoded["font"] = latin
        east_asian = _themed(fonts.get(qn("w:eastAsia")), fonts.get(qn("w:eastAsiaTheme")))
        if east_asian is not None:
            decoded["font_east_asia"] = east_asian
        complex_script = _themed(fonts.get(qn("w:cs")), fonts.get(qn("w:cstheme")))
        if complex_script is not None:
            decoded["font_cs"] = complex_script

    color = properties.find(_W_COLOR)
    if color is not None:
        decoded["color"] = _themed(
            color.get(_W_VAL),
            color.get(qn("w:themeColor")),
            tint=color.get(qn("w:themeTint")),
            shade=color.get(qn("w:themeShade")),
        )
    return decoded


def decode_ppr(properties: etree._Element | None) -> dict[str, Any]:
    """Decode a ``w:pPr`` into the minimal paragraph model.

    Measurements stay in twips (1/1440 inch), the unit the package stores and
    the one :mod:`word_document_server.engine.inspect` already reports.
    """
    if properties is None:
        return {}
    decoded: dict[str, Any] = {}
    for name, tag in _PARAGRAPH_TOGGLES.items():
        value = _toggle_value(properties, tag)
        if value is not None:
            decoded[name] = value

    alignment = _child_val(properties, _W_JC)
    if alignment is not None:
        decoded["alignment"] = alignment

    outline = _as_int(_child_val(properties, _W_OUTLINE_LVL))
    if outline is not None:
        decoded["outline_level"] = outline

    indent = properties.find(_W_IND)
    if indent is not None:
        values: dict[str, int] = {}
        for attribute, name in _INDENT_ATTRIBUTES.items():
            amount = _as_int(indent.get(qn(f"w:{attribute}")))
            if amount is not None:
                values.setdefault(name, amount)
        if values:
            decoded["indent"] = values

    spacing = properties.find(_W_SPACING)
    if spacing is not None:
        values_any: dict[str, Any] = {}
        for attribute, name in _SPACING_ATTRIBUTES.items():
            raw = spacing.get(qn(f"w:{attribute}"))
            if raw is None:
                continue
            values_any[name] = raw if name == "line_rule" else _as_int(raw)
        if values_any:
            decoded["spacing"] = values_any

    numbering = properties.find(_W_NUMPR)
    if numbering is not None:
        # The link to the numbering definition, not the numbering itself: which
        # bullet or number that level draws is engine/numbering.py's answer.
        # ``w:numId`` and ``w:ilvl`` are separate children, so a level that sets
        # only one of them leaves the other to be inherited -- which is why the
        # absent one is omitted rather than reported as ``None``.
        link: dict[str, Any] = {}
        num_id = _int_val(numbering, _W_NUM_ID)
        if num_id is not None:
            link["num_id"] = num_id
        level = _int_val(numbering, _W_ILVL)
        if level is not None:
            link["level"] = level
        decoded["numbering"] = link
    return decoded


# --------------------------------------------------------------------------------------
# One style, resolved
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class StyleLevel:
    """One step of a ``basedOn`` chain, decoded on its own.

    Attributes:
        style_id: the id of the style at this level, or ``""`` for the
            ``w:docDefaults`` level, which has none.
        name: its ``w:name``, or ``"Document defaults"`` for that level.
        run_props: what this level alone says about runs.
        paragraph_props: what it alone says about paragraphs.
    """

    style_id: str
    name: str
    run_props: dict[str, Any]
    paragraph_props: dict[str, Any]


@dataclass(frozen=True)
class StyleDetail:
    """A style, its inheritance chain and the merge of the two.

    Attributes:
        info: the identity and metadata of the style asked for.
        run_props: what the style itself says about runs, nothing inherited.
        paragraph_props: likewise, about paragraphs.
        chain: the levels that feed `resolved`, **most specific first**: the
            style, then its ``basedOn`` ancestors, then the ``w:docDefaults``
            level last.  Reading it top to bottom answers "where does this
            property come from".
        resolved: ``{"run": {...}, "paragraph": {...}}``, the chain applied in
            reverse -- defaults first, the style last.  See the module docstring
            on what it deliberately leaves out.
        warnings: what made the reading incomplete: a ``basedOn`` pointing at a
            style the document does not define, a chain that loops.
    """

    info: StyleInfo
    run_props: dict[str, Any]
    paragraph_props: dict[str, Any]
    chain: tuple[StyleLevel, ...]
    resolved: dict[str, dict[str, Any]]
    warnings: tuple[str, ...] = ()


def _doc_defaults(pkg: DocxPackage) -> StyleLevel:
    """The ``w:docDefaults`` level -- the floor every chain stands on."""
    root = _styles_root(pkg)
    defaults = None if root is None else root.find(_W_DOC_DEFAULTS)
    run = None
    paragraph = None
    if defaults is not None:
        run_default = defaults.find(_W_RPR_DEFAULT)
        if run_default is not None:
            run = run_default.find(_W_RPR)
        paragraph_default = defaults.find(_W_PPR_DEFAULT)
        if paragraph_default is not None:
            paragraph = paragraph_default.find(_W_PPR)
    return StyleLevel("", "Document defaults", decode_rpr(run), decode_ppr(paragraph))


def _merge_into(resolved: dict[str, Any], level: dict[str, Any]) -> None:
    """Apply one level's properties on top of what lower levels resolved to.

    A property of :data:`_MERGED_PROPERTIES` is merged key by key so that a
    level stating one of its parts leaves the others inherited; every other
    property is replaced outright.
    """
    for name, value in level.items():
        if name in _MERGED_PROPERTIES and isinstance(value, dict):
            existing = resolved.get(name)
            merged = dict(existing) if isinstance(existing, dict) else {}
            merged.update(value)
            resolved[name] = merged
        else:
            resolved[name] = value


def _by_id_and_name(
    styles: list[etree._Element],
) -> tuple[dict[str, etree._Element], dict[str, etree._Element]]:
    """Index the styles by id and by name; the first of a duplicate name wins."""
    by_id: dict[str, etree._Element] = {}
    by_name: dict[str, etree._Element] = {}
    for style in styles:
        style_id = style.get(_W_STYLE_ID)
        if style_id is None:
            continue  # pragma: no cover - filtered by _style_elements
        by_id.setdefault(style_id, style)
        name = _child_val(style, _W_NAME)
        if name is not None:
            by_name.setdefault(name, style)
    return by_id, by_name


def get_style(pkg: DocxPackage, id_or_name: str) -> StyleDetail:
    """Read one style, its ``basedOn`` chain and the properties they resolve to.

    Args:
        pkg: the package to read.
        id_or_name: the ``w:styleId`` (``"Heading1"``) or the ``w:name``
            (``"heading 1"``).  The id is tried first, so a document where one
            style's name is another style's id resolves the way a ``w:pStyle``
            would.

    Returns:
        A :class:`StyleDetail`.

    Raises:
        ValueError: if `id_or_name` is empty.
        PackageError: if no style of the document answers to it.  The message
            lists a few ids the document does define.
    """
    if not isinstance(id_or_name, str) or not id_or_name.strip():
        raise ValueError("a style is looked up by a non-empty id or name")
    styles = _style_elements(pkg)
    by_id, by_name = _by_id_and_name(styles)
    # Explicit ``is None``: an lxml element is falsy when it has no child, so
    # ``by_id.get(...) or by_name.get(...)`` would skip a childless style.
    element = by_id.get(id_or_name)
    if element is None:
        element = by_name.get(id_or_name)
    if element is None:
        known = sorted(by_id)[:10]
        raise PackageError(
            f"this document defines no style with id or name {id_or_name!r}; "
            f"it defines {known}"
        )

    levels: list[StyleLevel] = []
    warnings: list[str] = []
    seen: set[str] = set()
    current: etree._Element | None = element
    while current is not None:
        style_id = current.get(_W_STYLE_ID) or ""
        if style_id in seen:
            warnings.append(
                f"the basedOn chain of {id_or_name!r} loops back to {style_id!r}; "
                "it was followed once and stopped there"
            )
            break
        seen.add(style_id)
        levels.append(
            StyleLevel(
                style_id=style_id,
                name=_child_val(current, _W_NAME) or style_id,
                run_props=decode_rpr(current.find(_W_RPR)),
                paragraph_props=decode_ppr(current.find(_W_PPR)),
            )
        )
        parent_id = _child_val(current, _W_BASED_ON)
        if parent_id is None:
            break
        current = by_id.get(parent_id)
        if current is None:
            warnings.append(
                f"style {style_id!r} is based on {parent_id!r}, which this document "
                "does not define; the chain stops there"
            )

    levels.append(_doc_defaults(pkg))

    resolved_run: dict[str, Any] = {}
    resolved_paragraph: dict[str, Any] = {}
    for level in reversed(levels):
        _merge_into(resolved_run, level.run_props)
        _merge_into(resolved_paragraph, level.paragraph_props)

    return StyleDetail(
        info=_style_info(element),
        run_props=dict(levels[0].run_props),
        paragraph_props=dict(levels[0].paragraph_props),
        chain=tuple(levels),
        resolved={"run": resolved_run, "paragraph": resolved_paragraph},
        warnings=tuple(warnings),
    )


# --------------------------------------------------------------------------------------
# Where a style is used
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class StyleUsage:
    """One place a style is applied.

    Attributes:
        kind: ``"paragraph"`` (``w:pStyle``), ``"run"`` (``w:rStyle``) or
            ``"table"`` (``w:tblStyle``).
        story: the story id, as
            :meth:`~word_document_server.engine.package.DocxPackage.stories`
            names it -- never the ``"body"`` alias.
        locator: a locator naming the place, ready to hand to
            :func:`~word_document_server.engine.locators.resolve`.  Two
            exceptions, both loud rather than approximate:

            * for a ``table``, it is ``{"story", "table"}`` -- the table's
              address.  The locator vocabulary has no form for a whole table, so
              a caller that wants a cell adds ``row`` and ``col``; passing it as
              it stands fails as ``invalid``, which is the correct answer, and
              not the silent hit on some arbitrary cell that a fabricated
              ``row: 0, col: 0`` would give.
            * ``None`` for a paragraph that is in neither the V2 index space nor
              a table cell -- a text box.  Reach it with a ``find`` locator.
        index: the V2 paragraph index, or ``None`` outside that space (D-016).
        table: the table index for a ``table`` usage, ``None`` otherwise.
        start: for a ``run``, the offset of its first character in the visible
            text of its paragraph; ``None`` for the other kinds.
        end: offset just past its last character, so that `start` and `end` can
            be pasted into ``doc_format_range``.  ``None`` likewise.
        text: the visible text of the place, truncated to
            :data:`USAGE_PREVIEW`.
    """

    kind: str
    story: str
    locator: dict[str, Any] | None
    index: int | None
    table: int | None = None
    start: int | None = None
    end: int | None = None
    text: str = ""
    truncated: bool = False


def _preview(text: str) -> tuple[str, bool]:
    """`text` cut to :data:`USAGE_PREVIEW`, and whether cutting happened."""
    if len(text) <= USAGE_PREVIEW:
        return text, False
    return text[:USAGE_PREVIEW], True


def _cell_locators(story_id: str, root: etree._Element) -> dict[etree._Element, dict[str, Any]]:
    """Paragraph element -> its ``table`` locator, for every table cell paragraph.

    Built from the same two spaces
    :func:`~word_document_server.engine.locators.resolve` walks: the table order
    of :func:`~word_document_server.engine.locators._tables`, and the direct
    ``w:tr``/``w:tc``/``w:p`` children beneath it.  Nested tables are in that
    order too, so a paragraph in a nested cell gets the nested table's index.
    """
    found: dict[etree._Element, dict[str, Any]] = {}
    for table_index, table in enumerate(_tables(root)):
        for row_index, row in enumerate(table.findall(_W_TR)):
            for column_index, cell in enumerate(row.findall(_W_TC)):
                for paragraph_index, paragraph in enumerate(cell.findall(_W_P)):
                    found[paragraph] = {
                        "story": story_id,
                        "table": table_index,
                        "row": row_index,
                        "col": column_index,
                        "paragraph": paragraph_index,
                    }
    return found


def _run_spans(paragraph: etree._Element) -> dict[etree._Element, tuple[int, int]]:
    """Run element -> its ``(start, end)`` in the paragraph's visible text.

    Computed from :func:`~word_document_server.engine.textmodel.segments`, so a
    run inside a hyperlink, an insertion or a content control is measured like
    any other, and a run whose text is hidden under a deletion collapses to an
    empty span at the offset where it sits rather than disappearing.
    """
    spans: dict[etree._Element, tuple[int, int]] = {}
    for segment in segments(paragraph):
        run = segment.run
        if run is None:
            continue
        low, high = spans.get(run, (segment.start, segment.end))
        spans[run] = (min(low, segment.start), max(high, segment.end))
    return spans


def find_style_usage(pkg: DocxPackage, style_id: str) -> list[StyleUsage]:
    """Every paragraph, run and table that applies `style_id`, in document order.

    The match is on the ``w:styleId`` exactly, which is what ``w:pStyle``,
    ``w:rStyle`` and ``w:tblStyle`` carry -- a style *name* finds nothing, so
    look the id up with :func:`get_style` first if that is all you have.
    Inheritance is not followed: a paragraph in a style *based on* `style_id`
    is not a usage of `style_id`, it is a usage of its own style.

    Every story is searched -- body, headers, footers, footnotes, endnotes --
    and each usage says which one it was found in.

    Args:
        pkg: the package to read.
        style_id: the ``w:styleId`` to look for.

    Returns:
        One :class:`StyleUsage` per place, stories in the order
        :meth:`~word_document_server.engine.package.DocxPackage.stories` returns
        them and document order within each.  Empty when the style is unused,
        which is also the answer for a style id the document does not define:
        this function reports usage, it does not validate the id.

    Raises:
        ValueError: if `style_id` is empty.
    """
    if not isinstance(style_id, str) or not style_id:
        raise ValueError("a style id to look for must be a non-empty string")

    usages: list[StyleUsage] = []
    for story_id, root in pkg.stories():
        indices = {
            paragraph: index for index, paragraph in enumerate(indexed_paragraphs(root))
        }
        cells = _cell_locators(story_id, root)
        tables = {table: index for index, table in enumerate(_tables(root))}

        for element in root.iter(_W_P, _W_TBL_PR):
            if element.tag == _W_TBL_PR:
                if _child_val(element, _W_TBL_STYLE) != style_id:
                    continue
                table = element.getparent()
                index = None if table is None else tables.get(table)
                if index is None:
                    # A w:tblPr outside the table index space: a table in a text
                    # box, or a w:tblPrEx-like stray.  It has no address.
                    continue
                usages.append(
                    StyleUsage(
                        kind="table",
                        story=story_id,
                        locator={"story": story_id, "table": index},
                        index=None,
                        table=index,
                    )
                )
                continue

            paragraph = element
            index = indices.get(paragraph)
            locator: dict[str, Any] | None = cells.get(paragraph)
            if locator is None and index is not None:
                locator = {"story": story_id, "paragraph": index}

            if _child_val(paragraph.find(_W_PPR), _W_PSTYLE) == style_id:
                text, truncated = _preview(visible_text(paragraph))
                usages.append(
                    StyleUsage(
                        kind="paragraph",
                        story=story_id,
                        locator=locator,
                        index=index,
                        text=text,
                        truncated=truncated,
                    )
                )

            spans: dict[etree._Element, tuple[int, int]] | None = None
            for run in paragraph.iter(_W_R):
                if _child_val(run.find(_W_RPR), _W_RSTYLE) != style_id:
                    continue
                if spans is None:
                    spans = _run_spans(paragraph)
                start, end = spans.get(run, (None, None))
                whole = visible_text(paragraph)
                text, truncated = _preview(
                    whole if start is None else whole[start:end]
                )
                usages.append(
                    StyleUsage(
                        kind="run",
                        story=story_id,
                        locator=locator,
                        index=index,
                        start=start,
                        end=end,
                        text=text,
                        truncated=truncated,
                    )
                )
    return usages


# --------------------------------------------------------------------------------------
# The part, for writing
# --------------------------------------------------------------------------------------

#: Content of a style sheet that defines nothing yet.
_EMPTY_STYLES = f'<w:styles xmlns:w="{W}"/>'

#: Children of ``w:style`` in schema order (ECMA-376 §17.7.4.17, ``CT_Style``).
#: The table-only children are listed even though this module writes paragraph
#: and character styles: a table style edited elsewhere keeps its own children in
#: the right place when one of these is added next to them.
STYLE_ORDER: tuple[str, ...] = (
    "w:name",
    "w:aliases",
    "w:basedOn",
    "w:next",
    "w:link",
    "w:autoRedefine",
    "w:hidden",
    "w:uiPriority",
    "w:semiHidden",
    "w:unhideWhenUsed",
    "w:qFormat",
    "w:locked",
    "w:personal",
    "w:personalCompose",
    "w:personalReply",
    "w:rsid",
    "w:pPr",
    "w:rPr",
    "w:tblPr",
    "w:trPr",
    "w:tcPr",
    "w:tblStylePr",
)

_STYLE_RANK: dict[str, int] = {qn(tag): rank for rank, tag in enumerate(STYLE_ORDER)}

#: The families this module creates.  ``table`` and ``numbering`` styles are read
#: by :func:`list_styles` and reassigned by :func:`delete_style`, but writing one
#: means writing ``w:tblPr``/``w:tblStylePr`` or a ``w:numStyleLink``, which is a
#: different vocabulary and lives elsewhere.
WRITABLE_FAMILIES: tuple[str, ...] = ("paragraph", "character")

#: The keys a style spec may carry.  `name`, `family`, `based_on`, `next`,
#: `link`, `q_format`, `run_props` and `paragraph_props` are the vocabulary of
#: the plan; `style_id`, `ui_priority` and `builtin` are there because a caller
#: recreating one of Word's own styles needs all three -- see
#: :func:`create_style`.
STYLE_SPEC_KEYS: tuple[str, ...] = (
    "name",
    "style_id",
    "family",
    "based_on",
    "next",
    "link",
    "q_format",
    "ui_priority",
    "builtin",
    "run_props",
    "paragraph_props",
)

#: ``ST_Jc`` (ECMA-376 §17.18.44), the values ``w:jc`` may take.
ALIGNMENTS: frozenset[str] = frozenset(
    {
        "start",
        "end",
        "left",
        "right",
        "center",
        "both",
        "distribute",
        "numTab",
        "highKashida",
        "lowKashida",
        "mediumKashida",
        "thaiDistribute",
    }
)

#: ``ST_LineSpacingRule`` (ECMA-376 §17.18.48).
LINE_RULES: frozenset[str] = frozenset({"auto", "exact", "atLeast"})

#: The properties a `paragraph_props` mapping may name -- the vocabulary
#: :func:`decode_ppr` reports, in the order they are applied.
PARAGRAPH_PROPERTIES: tuple[str, ...] = (
    *_PARAGRAPH_TOGGLES,
    "alignment",
    "outline_level",
    "indent",
    "spacing",
    "numbering",
)

#: Decoded name -> (``w:rFonts`` attributes carrying the literal, attributes
#: carrying the theme reference).  Mirrors what :func:`decode_rpr` reads back.
_FONT_SLOTS: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
    "font": (("w:ascii", "w:hAnsi"), ("w:asciiTheme", "w:hAnsiTheme")),
    "font_east_asia": (("w:eastAsia",), ("w:eastAsiaTheme",)),
    "font_cs": (("w:cs",), ("w:cstheme",)),
}

#: ``w:ind`` attribute of each decoded indent key.  ``start``/``end`` are read
#: back by :func:`decode_ppr` as aliases of ``left``/``right``, but only the
#: latter pair is written: Word writes those and a document carrying both says
#: the same thing twice.
_INDENT_SLOTS: dict[str, str] = {
    "left": "w:left",
    "right": "w:right",
    "first_line": "w:firstLine",
    "hanging": "w:hanging",
}

#: ``w:spacing`` attribute of each decoded spacing key.
_SPACING_SLOTS: dict[str, str] = {
    "before": "w:before",
    "after": "w:after",
    "line": "w:line",
    "line_rule": "w:lineRule",
}

#: Children of ``w:numPr`` in schema order.  The last two belong to a tracked
#: numbering change; they are never written here, only kept in place.
_NUMPR_RANK: dict[str, int] = {
    qn(tag): rank
    for rank, tag in enumerate(("w:ilvl", "w:numId", "w:numberingChange", "w:ins"))
}

#: Highest ``w:outlineLvl``: 0 to 8 are the nine heading levels, 9 is body text.
_MAX_OUTLINE_LEVEL = 9

#: One step of a spec: mutate the property element it is handed.  Built while the
#: spec is validated and run afterwards, so a refused value cannot leave a
#: half-written style behind.
_Directive = Callable[[etree._Element], None]

#: The ``w:style`` children naming another style, and what each one means.
_REFERENCE_CHILDREN: dict[str, str] = {
    "based_on": "w:basedOn",
    "next": "w:next",
    "link": "w:link",
}


def _document_part(pkg: DocxPackage | Document) -> Part:
    """The main document part of `pkg`, which may be either kind of handle.

    Both are accepted for the same reason
    :mod:`word_document_server.engine.numbering` accepts both: the V2 tools hold
    a :class:`DocxPackage` and the older ``core``/``utils`` code holds a
    python-docx ``Document``, and a style sheet is only consistent if both edit
    the same part of the same package.

    Raises:
        TypeError: if `pkg` is neither.
    """
    if isinstance(pkg, DocxPackage):
        return pkg.document_part
    part = getattr(pkg, "part", None)
    if isinstance(part, Part):
        return part
    raise TypeError(
        f"expected a DocxPackage or a python-docx Document, got {type(pkg).__name__}"
    )


def _live_styles_root(part: Part) -> etree._Element:
    """The live root of a style part, refusing a blob-backed one."""
    if not isinstance(part, XmlPart):
        raise PackageError(
            f"part {part.partname} holds the styles but is not loaded as live XML "
            f"({part.content_type}); edits made on it would be dropped at save time"
        )
    return part.element


def styles_root(
    pkg: DocxPackage | Document, *, create: bool = False
) -> etree._Element | None:
    """Return the root ``w:styles`` element, creating the part on demand.

    The part is found through the document part's ``styles`` relationship rather
    than by name, because that is what Word follows.  When `create` is true and
    no such relationship exists, an empty style sheet is added and related -- and
    if a part already sits under :data:`STYLES_PARTNAME` without being related,
    it is related rather than duplicated.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
        create: add the part when the package has none.

    Returns:
        The live root element, or ``None`` when the package has no style sheet
        and `create` is false.

    Raises:
        PackageError: if the style part is not loaded as live XML, which would
            make every edit made here vanish at save time.
    """
    part = _document_part(pkg)
    for relationship in part.rels.values():
        if relationship.reltype == RT.STYLES and not relationship.is_external:
            return _live_styles_root(relationship.target_part)
    if not create:
        return None

    package = part.package
    for existing in package.iter_parts():
        if str(existing.partname) == STYLES_PARTNAME:
            part.relate_to(existing, RT.STYLES)
            return _live_styles_root(existing)

    element = parse_xml(_EMPTY_STYLES.encode("utf-8"))
    created = XmlPart(PackURI(STYLES_PARTNAME), CT.WML_STYLES, element, package)
    part.relate_to(created, RT.STYLES)
    return element


def _require_styles_root(pkg: DocxPackage | Document) -> etree._Element:
    """The style sheet root, or a :class:`PackageError` if there is none."""
    root = styles_root(pkg)
    if root is None:
        raise PackageError("the package has no style part, so it defines no style")
    return root


def _ensure_styles_root(pkg: DocxPackage | Document) -> etree._Element:
    """The style sheet root, created if the package has none."""
    root = styles_root(pkg, create=True)
    if root is None:  # unreachable: with create=True the part is added or it raises
        raise PackageError("the style part could not be created")
    return root


# --------------------------------------------------------------------------------------
# Identity of a new style
# --------------------------------------------------------------------------------------


def derive_style_id(name: str) -> str:
    """The ``w:styleId`` Word derives from a style name.

    Word's rule is to keep the letters and digits of the name and drop
    everything else, so ``"Fixture Body"`` becomes ``"FixtureBody"``.  It does
    *not* change the case, which is why recreating one of Word's own styles --
    whose name is ``"heading 1"`` but whose id is ``"Heading1"`` -- means passing
    the id explicitly rather than deriving it.

    Raises:
        ValueError: if `name` holds no letter or digit at all, leaving nothing to
            derive an id from.
    """
    if not isinstance(name, str):
        raise TypeError(f"a style name is a string, got {type(name).__name__}")
    derived = "".join(character for character in name if character.isalnum())
    if not derived:
        raise ValueError(
            f"no style id can be derived from {name!r}: a style id is the name "
            "without its spaces and punctuation, and that leaves nothing"
        )
    return derived


def _check_free(
    root: etree._Element, style_id: str, name: str, *, excluding: etree._Element | None = None
) -> None:
    """Refuse `style_id` or `name` if another style already answers to it.

    Ids are compared exactly, names case-insensitively: Word's style pane treats
    "Body Text" and "body text" as one name, and letting both exist produces a
    document whose style list has two entries a human cannot tell apart.

    Raises:
        LocatorError: ``already_exists``.
    """
    for style in _styles_in(root):
        if style is excluding:
            continue
        existing_id = style.get(_W_STYLE_ID) or ""
        if existing_id == style_id:
            raise LocatorError(
                "already_exists",
                f"this document already defines a style with the id {style_id!r} "
                f"(named {_child_val(style, _W_NAME) or existing_id!r}); "
                "update it, clone it, or choose another name",
            )
        existing_name = _child_val(style, _W_NAME)
        if existing_name is not None and existing_name.casefold() == name.casefold():
            raise LocatorError(
                "already_exists",
                f"this document already defines a style named {existing_name!r} "
                f"(id {existing_id!r}); update it, clone it, or choose another name",
            )


def _find_style(root: etree._Element, id_or_name: str) -> etree._Element:
    """The ``w:style`` `id_or_name` addresses, by id first then by name.

    Raises:
        ValueError: if `id_or_name` is not a non-empty string.
        LocatorError: ``not_found`` if no style answers to it.
    """
    if not isinstance(id_or_name, str) or not id_or_name.strip():
        raise ValueError("a style is looked up by a non-empty id or name")
    by_id, by_name = _by_id_and_name(_styles_in(root))
    # Explicit ``is None``: an lxml element with no child is falsy.
    element = by_id.get(id_or_name)
    if element is None:
        element = by_name.get(id_or_name)
    if element is None:
        raise LocatorError(
            "not_found",
            f"this document defines no style with id or name {id_or_name!r}; "
            f"it defines {sorted(by_id)[:10]}",
        )
    return element


def style_info(pkg: DocxPackage | Document, id_or_name: str) -> StyleInfo:
    """The identity and metadata of one style, without its chain.

    :func:`get_style` answers the same question and much more, but it needs a
    :class:`DocxPackage`; this one reads the style sheet alone, so it also works
    on the python-docx ``Document`` the pre-V2 tools hold.

    Raises:
        ValueError: if `id_or_name` is not a non-empty string.
        LocatorError: ``not_found`` if no style answers to it.
        PackageError: if the package has no style part.
    """
    return _style_info(_find_style(_require_styles_root(pkg), id_or_name))


def _reference_id(
    root: etree._Element, field: str, value: str, self_id: str
) -> str:
    """The id to write for a ``basedOn``/``next``/``link`` reference.

    `value` may be an id or a name; the id is what gets written, because that is
    what Word resolves.  A style referring to itself is legal and common
    (``w:next`` on a body style), so `self_id` is accepted before the lookup --
    the style being created is not in `root` yet.

    Raises:
        ValueError: if `value` is not a non-empty string.
        LocatorError: ``not_found`` if `value` names no style of the document.
    """
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field!r} takes a style id or name, got {value!r}")
    if value == self_id:
        return value
    try:
        target = _find_style(root, value)
    except LocatorError as exc:
        raise LocatorError(
            "not_found",
            f"{field!r} names {value!r}, which this document does not define; "
            "a style reference pointing at nothing is ignored by Word and lost "
            f"on the next round trip ({exc})",
        ) from exc
    return target.get(_W_STYLE_ID) or value


# --------------------------------------------------------------------------------------
# Writing properties
# --------------------------------------------------------------------------------------


def _themed_parts(field: str, value: Mapping[str, Any]) -> tuple[Any, Any]:
    """The ``value``/``theme`` halves of a themed property, validated."""
    unknown = sorted(set(value) - {"value", "theme", "tint", "shade"})
    if unknown:
        raise ValueError(
            f"{field!r} takes {{'value', 'theme'}} as get_style reports it, "
            f"plus 'tint' and 'shade' for a colour; unknown: {', '.join(unknown)}"
        )
    return value.get("value"), value.get("theme")


def _font_directive(field: str, value: Any) -> list[_Directive]:
    """Directives writing one ``w:rFonts`` slot pair, theme reference included.

    A plain string is the literal font with no theme; the mapping form is what
    :func:`decode_rpr` reports, and both halves are written back as they were.
    Only the slots this property owns are touched, so setting ``font`` leaves the
    east-asian font of the style alone.
    """
    literal: Any
    theme: Any
    if value is None:
        literal = theme = None
    elif isinstance(value, Mapping):
        literal, theme = _themed_parts(field, value)
    elif isinstance(value, str):
        literal, theme = value, None
    else:
        raise TypeError(
            f"{field!r} takes a font name, a {{'value', 'theme'}} mapping or None, "
            f"got {value!r}"
        )
    for part in (literal, theme):
        if part is not None and (not isinstance(part, str) or not part.strip()):
            raise ValueError(f"{field!r} must be a non-empty font or theme name, got {part!r}")

    slots, theme_slots = _FONT_SLOTS[field]

    def directive(properties: etree._Element) -> None:
        if literal is None and theme is None:
            fonts = properties.find(_W_RFONTS)
            if fonts is None:
                return
            for attribute in (*slots, *theme_slots):
                fonts.attrib.pop(qn(attribute), None)
            # An rFonts with nothing left says nothing; keep it only if another
            # script still uses it.
            if not fonts.attrib:
                properties.remove(fonts)
            return
        fonts = _ensure_child(properties, _W_RFONTS, _RPR_RANK)
        for attribute, setting in ((slots, literal), (theme_slots, theme)):
            for name in attribute:
                if setting is None:
                    fonts.attrib.pop(qn(name), None)
                else:
                    fonts.set(qn(name), setting)

    return [directive]


def _themed_color_directives(value: Mapping[str, Any], pkg: Any) -> list[_Directive]:
    """Directives writing a ``w:color`` that follows the theme.

    The literal goes through
    :func:`~word_document_server.engine.format.apply_rpr`'s own validation --
    there is one definition of what a colour may be in this engine -- and the
    theme attributes are put back on top of it afterwards, since that validation
    clears them by design (an explicit colour beats a theme one in Word, so the
    two are only written together on purpose, which is exactly this case).
    """
    unknown = sorted(set(value) - {"value", "theme", "tint", "shade"})
    if unknown:
        raise ValueError(
            "'color' takes {'value', 'theme', 'tint', 'shade'} as get_style "
            f"reports it; unknown: {', '.join(unknown)}"
        )
    literal = value.get("value")
    theme = value.get("theme")
    extras = {"w:themeColor": theme, "w:themeTint": value.get("tint"),
              "w:themeShade": value.get("shade")}
    if literal is None and all(extra is None for extra in extras.values()):
        return _parse_patch({"color": None}, pkg)
    for name, extra in extras.items():
        if extra is not None and not isinstance(extra, str):
            raise TypeError(f"{name} takes a string or None, got {extra!r}")
    # ``w:val`` is required on ``w:color``; a style that only names a theme
    # colour caches "auto" beside it, which is what Word writes too.
    directives = _parse_patch({"color": literal if literal is not None else "auto"}, pkg)

    def directive(properties: etree._Element) -> None:
        color = properties.find(_W_COLOR)
        if color is None:  # pragma: no cover - the directive above just made it
            return
        for name, extra in extras.items():
            if extra is not None:
                color.set(qn(name), extra)

    return [*directives, directive]


def _run_directives(patch: Mapping[str, Any], pkg: Any) -> list[_Directive]:
    """Turn a `run_props` mapping into directives over a ``w:rPr``.

    The vocabulary is the one :func:`decode_rpr` reports and
    :func:`~word_document_server.engine.format.apply_rpr` takes: what is read
    from a style can be written back to another without a translation table.  Two
    kinds of value need handling here rather than there -- the themed form of
    ``font`` and ``color``, which a *run* never carries in the minimal model but a
    style very much does.

    Raises:
        TypeError: if `patch` is not a mapping, or a value is of the wrong type.
        ValueError: on an unknown property or a value the schema cannot store.
    """
    if not isinstance(patch, Mapping):
        raise TypeError(
            f"'run_props' must be a mapping of property names, got {type(patch).__name__}"
        )
    known = (*RUN_PROPERTIES, *_FONT_SLOTS)
    unknown = sorted(set(patch) - set(known))
    if unknown:
        raise ValueError(
            f"unknown run properties: {', '.join(unknown)}; known: {', '.join(sorted(known))}"
        )
    if "char_style" in patch and not isinstance(pkg, DocxPackage):
        # apply_rpr checks a char_style against the style sheet, and reaching it
        # needs the package handle.  Saying so beats the AttributeError a
        # python-docx Document would produce three frames down.
        raise ValueError(
            "'char_style' is checked against the styles the package defines, which "
            "needs a DocxPackage; open the document with DocxPackage.open to set it"
        )
    plain: dict[str, Any] = {}
    themed: list[_Directive] = []
    for name, value in patch.items():
        if name in _FONT_SLOTS and (name != "font" or isinstance(value, Mapping)):
            themed += _font_directive(name, value)
        elif name == "color" and isinstance(value, Mapping):
            themed += _themed_color_directives(value, pkg)
        else:
            plain[name] = value
    return [*_parse_patch(plain, pkg), *themed]


def _int_setting(field: str, value: Any, *, minimum: int | None = None,
                 maximum: int | None = None) -> str:
    """`value` as the decimal string an attribute stores, bounds checked."""
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"{field!r} takes a whole number or None, got {value!r}")
    if minimum is not None and value < minimum:
        raise ValueError(f"{field!r} must be {minimum} or more, got {value}")
    if maximum is not None and value > maximum:
        raise ValueError(f"{field!r} must be at most {maximum}, got {value}")
    return str(value)


def _attribute_child_directive(
    field: str, tag: str, slots: Mapping[str, str], value: Any,
    *, string_keys: frozenset[str] = frozenset(),
    allowed: Mapping[str, frozenset[str]] | None = None,
) -> list[_Directive]:
    """Directives for a ``w:pPr`` child whose settings are its attributes.

    ``w:ind`` and ``w:spacing`` each hold several independent settings, and
    :func:`get_style` inherits them one by one; writing them the same way -- the
    named attributes set, the others left as they were -- is what makes "add a
    space after" not silently drop the line spacing the style already had.  The
    whole child is removed by passing ``None``, and one of its attributes by
    mapping that key to ``None``.
    """
    qualified = qn(tag)
    if value is None:
        def remove(properties: etree._Element) -> None:
            _drop_children(properties, qualified)
        return [remove]
    if not isinstance(value, Mapping):
        raise TypeError(f"{field!r} takes a mapping of settings or None, got {value!r}")
    unknown = sorted(set(value) - set(slots))
    if unknown:
        raise ValueError(
            f"unknown {field} settings: {', '.join(unknown)}; known: {', '.join(slots)}"
        )
    settings: dict[str, str | None] = {}
    for name, raw in value.items():
        if raw is None:
            settings[slots[name]] = None
        elif name in string_keys:
            choices = (allowed or {}).get(name)
            if not isinstance(raw, str):
                raise TypeError(f"{field}.{name} takes a string or None, got {raw!r}")
            if choices is not None and raw not in choices:
                raise ValueError(
                    f"unknown {field}.{name} {raw!r}; known: {', '.join(sorted(choices))}"
                )
            settings[slots[name]] = raw
        else:
            settings[slots[name]] = _int_setting(f"{field}.{name}", raw)

    def directive(properties: etree._Element) -> None:
        child = _ensure_child(properties, qualified, _PPR_RANK)
        for attribute, setting in settings.items():
            if setting is None:
                child.attrib.pop(qn(attribute), None)
            else:
                child.set(qn(attribute), setting)
        if not child.attrib and len(child) == 0:
            properties.remove(child)

    return [directive]


def _numbering_directive(value: Any) -> list[_Directive]:
    """Directives for ``w:numPr`` -- the link to a numbering definition.

    Which glyph or number that definition draws is
    :mod:`word_document_server.engine.numbering`'s answer, not this module's; a
    style only points at it.
    """
    numpr = qn("w:numPr")
    if value is None:
        def remove(properties: etree._Element) -> None:
            _drop_children(properties, numpr)
        return [remove]
    if not isinstance(value, Mapping):
        raise TypeError(f"'numbering' takes {{'num_id', 'level'}} or None, got {value!r}")
    unknown = sorted(set(value) - {"num_id", "level"})
    if unknown:
        raise ValueError(
            f"unknown numbering settings: {', '.join(unknown)}; known: level, num_id"
        )
    children: dict[str, str | None] = {}
    if "level" in value:
        raw = value["level"]
        children["w:ilvl"] = (
            None if raw is None else _int_setting("numbering.level", raw, minimum=0, maximum=8)
        )
    if "num_id" in value:
        raw = value["num_id"]
        children["w:numId"] = (
            None if raw is None else _int_setting("numbering.num_id", raw, minimum=0)
        )

    def directive(properties: etree._Element) -> None:
        link = _ensure_child(properties, numpr, _PPR_RANK)
        for tag, setting in children.items():
            if setting is None:
                _drop_children(link, qn(tag))
            else:
                # w:ilvl then w:numId, the schema's order for CT_NumPr.
                _ensure_child(link, qn(tag), _NUMPR_RANK).set(_W_VAL, setting)
        if len(link) == 0:
            properties.remove(link)

    return [directive]


def _paragraph_directives(patch: Mapping[str, Any]) -> list[_Directive]:
    """Turn a `paragraph_props` mapping into directives over a ``w:pPr``.

    The vocabulary is :func:`decode_ppr`'s, measurements included: twips, the
    unit the package stores.

    Raises:
        TypeError: if `patch` is not a mapping, or a value is of the wrong type.
        ValueError: on an unknown property or a value the schema cannot store.
    """
    if not isinstance(patch, Mapping):
        raise TypeError(
            f"'paragraph_props' must be a mapping of property names, got {type(patch).__name__}"
        )
    unknown = sorted(set(patch) - set(PARAGRAPH_PROPERTIES))
    if unknown:
        raise ValueError(
            f"unknown paragraph properties: {', '.join(unknown)}; "
            f"known: {', '.join(sorted(PARAGRAPH_PROPERTIES))}"
        )
    directives: list[_Directive] = []
    for name in PARAGRAPH_PROPERTIES:
        if name not in patch:
            continue
        value = patch[name]
        if name in _PARAGRAPH_TOGGLES:
            directives += _toggle_directive(name, _PARAGRAPH_TOGGLES[name], value)
        elif name == "alignment":
            directives += _val_directive(
                name, qn("w:jc"), value, allowed=ALIGNMENTS
            )
        elif name == "outline_level":
            directives += _val_directive(
                name,
                qn("w:outlineLvl"),
                value
                if value is None
                else _int_setting(name, value, minimum=0, maximum=_MAX_OUTLINE_LEVEL),
            )
        elif name == "indent":
            directives += _attribute_child_directive(name, "w:ind", _INDENT_SLOTS, value)
        elif name == "spacing":
            directives += _attribute_child_directive(
                name,
                "w:spacing",
                _SPACING_SLOTS,
                value,
                string_keys=frozenset({"line_rule"}),
                allowed={"line_rule": LINE_RULES},
            )
        else:  # numbering
            directives += _numbering_directive(value)
    return directives


def _toggle_directive(field: str, tag: str, value: Any) -> list[_Directive]:
    """``True`` writes the bare element, ``False`` an explicit ``w:val="0"``.

    The explicit off matters on a style as much as on a run: it is how a style
    cancels something the style it is based on turned on.
    """
    if value is not None and not isinstance(value, bool):
        raise TypeError(f"{field!r} takes True, False or None, got {value!r}")

    def directive(properties: etree._Element) -> None:
        if value is None:
            _drop_children(properties, tag)
            return
        child = _ensure_child(properties, tag, _PPR_RANK)
        child.attrib.pop(_W_VAL, None)
        if not value:
            child.set(_W_VAL, "0")

    return [directive]


def _val_directive(
    field: str, tag: str, value: Any, *, allowed: frozenset[str] | None = None
) -> list[_Directive]:
    """A ``w:pPr`` child that is nothing but its ``w:val``."""
    if value is not None:
        if not isinstance(value, str):
            raise TypeError(f"{field!r} takes a string or None, got {value!r}")
        if allowed is not None and value not in allowed:
            raise ValueError(
                f"unknown {field} {value!r}; known: {', '.join(sorted(allowed))}"
            )

    def directive(properties: etree._Element) -> None:
        if value is None:
            _drop_children(properties, tag)
            return
        _ensure_child(properties, tag, _PPR_RANK).set(_W_VAL, value)

    return [directive]


def _apply_properties(
    style: etree._Element, tag: str, directives: list[_Directive]
) -> None:
    """Run `directives` on the ``w:rPr``/``w:pPr`` of `style`, creating it once.

    A property block left empty is removed, so setting a property and removing it
    gives the style back exactly as it was found.
    """
    if not directives:
        return
    properties = _ensure_child(style, tag, _STYLE_RANK)
    for directive in directives:
        directive(properties)
    if len(properties) == 0 and not properties.attrib:
        style.remove(properties)


# --------------------------------------------------------------------------------------
# Creating, updating, cloning
# --------------------------------------------------------------------------------------


def _check_spec(spec: Any) -> Mapping[str, Any]:
    """Validate the shape of a style spec and return it.

    Raises:
        TypeError: if `spec` is not a mapping.
        ValueError: on a key outside :data:`STYLE_SPEC_KEYS`.  A misspelled key
            silently ignored would report a style created to the caller's
            specification and write something else.
    """
    if not isinstance(spec, Mapping):
        raise TypeError(f"a style spec is a mapping, got {type(spec).__name__}")
    unknown = sorted(set(spec) - set(STYLE_SPEC_KEYS))
    if unknown:
        raise ValueError(
            f"unknown style spec keys: {', '.join(unknown)}; "
            f"known: {', '.join(STYLE_SPEC_KEYS)}"
        )
    return spec


def _apply_spec(
    root: etree._Element,
    style: etree._Element,
    spec: Mapping[str, Any],
    *,
    pkg: Any,
) -> None:
    """Write the metadata and the properties of `spec` onto `style`.

    Every key is optional and means "change this": a key the spec leaves out is
    left as it stands, which is what makes the same function serve
    :func:`create_style` (on an empty element) and :func:`update_style` (on an
    existing one).
    """
    style_id = style.get(_W_STYLE_ID) or ""

    if "name" in spec:
        name = spec["name"]
        if not isinstance(name, str) or not name.strip():
            raise ValueError(f"'name' must be a non-empty style name, got {name!r}")
        _ensure_child(style, _W_NAME, _STYLE_RANK).set(_W_VAL, name)

    for field, tag in _REFERENCE_CHILDREN.items():
        if field not in spec:
            continue
        value = spec[field]
        if value is None:
            _drop_children(style, qn(tag))
            continue
        _ensure_child(style, qn(tag), _STYLE_RANK).set(
            _W_VAL, _reference_id(root, field, value, style_id)
        )

    if "q_format" in spec:
        if spec["q_format"]:
            _ensure_child(style, _W_Q_FORMAT, _STYLE_RANK)
        else:
            _drop_children(style, _W_Q_FORMAT)

    if "ui_priority" in spec:
        value = spec["ui_priority"]
        if value is None:
            _drop_children(style, _W_UI_PRIORITY)
        else:
            _ensure_child(style, _W_UI_PRIORITY, _STYLE_RANK).set(
                _W_VAL, _int_setting("ui_priority", value, minimum=0)
            )

    if "builtin" in spec:
        if spec["builtin"]:
            style.attrib.pop(_W_CUSTOM_STYLE, None)
        else:
            style.set(_W_CUSTOM_STYLE, "1")

    family = style.get(_W_TYPE) or "paragraph"
    if "paragraph_props" in spec and family != "paragraph":
        raise ValueError(
            f"style {style_id!r} is a {family} style; only a paragraph style has "
            "paragraph properties"
        )
    _apply_properties(style, _W_RPR, _run_directives(spec.get("run_props") or {}, pkg))
    _apply_properties(
        style, _W_PPR, _paragraph_directives(spec.get("paragraph_props") or {})
    )


def _insert_style(root: etree._Element, style: etree._Element) -> None:
    """Put a new ``w:style`` after the last one, where ``CT_Styles`` wants it.

    ``w:docDefaults`` and ``w:latentStyles`` come first in the sequence and the
    styles follow, so appending after the last style is right whether or not the
    sheet has either.
    """
    existing = [other for other in root.findall(_W_STYLE) if other is not style]
    if existing:
        existing[-1].addnext(style)
    else:
        root.append(style)


def create_style(pkg: DocxPackage | Document, spec: Mapping[str, Any]) -> StyleInfo:
    """Define a new paragraph or character style and return what it says.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
            Its style part is created if it has none.
        spec: what the style is, as a mapping of :data:`STYLE_SPEC_KEYS`:

            ``name`` (required)
                the ``w:name``, what Word's style pane shows.
            ``style_id``
                the ``w:styleId`` every ``w:pStyle`` will carry.  Derived from
                the name by :func:`derive_style_id` when omitted; give it to
                recreate one of Word's own styles, whose id is not the
                derivation of its name (``heading 1`` -> ``Heading1``).
            ``family``
                ``"paragraph"`` (the default) or ``"character"``.
            ``based_on``, ``next``, ``link``
                other styles, by id or by name; each is stored as the id.
            ``q_format``
                whether Word offers the style in its gallery.
            ``ui_priority``
                where it sorts in that gallery.
            ``builtin``
                ``True`` to declare the style as one of Word's own (no
                ``w:customStyle``).  A style created here is custom by default,
                which is what it is.
            ``run_props``, ``paragraph_props``
                the properties, in :func:`get_style`'s own vocabulary.  A
                character style has no paragraph properties.

    Returns:
        The :class:`StyleInfo` of the style that was written.

    Raises:
        TypeError: if `spec` is not a mapping or a value is of the wrong type.
        ValueError: on an unknown key, an unknown family, or a value the schema
            cannot store.
        LocatorError: ``already_exists`` if the id or the name is taken;
            ``not_found`` if ``based_on``/``next``/``link`` names no style.
        PackageError: if the style part cannot be edited.
    """
    spec = _check_spec(spec)
    name = spec.get("name")
    if not isinstance(name, str) or not name.strip():
        raise ValueError("a new style needs a 'name'")
    family = spec.get("family", "paragraph")
    if family not in WRITABLE_FAMILIES:
        raise ValueError(
            f"a style is created as one of {', '.join(WRITABLE_FAMILIES)}, got {family!r}"
        )
    given_id = spec.get("style_id")
    if given_id is None:
        style_id = derive_style_id(name)
    elif not isinstance(given_id, str) or not given_id.strip():
        raise ValueError(f"'style_id' must be a non-empty style id, got {given_id!r}")
    else:
        style_id = given_id

    root = _ensure_styles_root(pkg)
    _check_free(root, style_id, name)

    style = etree.SubElement(root, _W_STYLE)
    style.set(_W_TYPE, family)
    style.set(_W_STYLE_ID, style_id)
    if not spec.get("builtin", False):
        style.set(_W_CUSTOM_STYLE, "1")
    _insert_style(root, style)
    try:
        _apply_spec(root, style, {"q_format": False, **spec}, pkg=pkg)
    except BaseException:
        # A refused spec must leave the style sheet exactly as it was found,
        # never a half-written style Word would then offer to a human.
        root.remove(style)
        raise
    return _style_info(style)


def update_style(
    pkg: DocxPackage | Document, id_or_name: str, spec: Mapping[str, Any]
) -> StyleInfo:
    """Change an existing style, one property at a time.

    The spec is read as a set of *changes*, the way
    :func:`~word_document_server.engine.format.apply_rpr` reads a patch: a key it
    leaves out is left alone, and a key mapped to ``None`` removes what the style
    said -- so what it inherits applies again.  The change reaches every
    paragraph, run and table that names the style at once; call
    :func:`find_style_usage` first to know how many that is.

    ``style_id`` cannot be changed: every ``w:pStyle`` of the document carries it,
    so a new id is a new style -- :func:`clone_style`. ``family`` cannot be
    changed either, for the same reason.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
        id_or_name: the style to change, by ``w:styleId`` or by ``w:name``.
        spec: the changes, in the vocabulary of :func:`create_style`.

    Returns:
        The :class:`StyleInfo` of the style as it now stands.

    Raises:
        TypeError: if `spec` is not a mapping or a value is of the wrong type.
        ValueError: on an unknown key, on ``style_id``, on a ``family`` that
            differs from the style's own, or on a value the schema cannot store.
        LocatorError: ``not_found`` if no style answers to `id_or_name` or a
            reference names nothing; ``already_exists`` if the new name is taken.
        PackageError: if the package has no style part.
    """
    spec = _check_spec(spec)
    root = _require_styles_root(pkg)
    style = _find_style(root, id_or_name)
    style_id = style.get(_W_STYLE_ID) or ""

    if "style_id" in spec and spec["style_id"] != style_id:
        raise ValueError(
            f"the id of {style_id!r} cannot be changed: every w:pStyle and w:rStyle "
            "of the document carries it.  Clone the style under the new name instead"
        )
    family = style.get(_W_TYPE) or "paragraph"
    if "family" in spec and spec["family"] != family:
        raise ValueError(
            f"style {style_id!r} is a {family} style; changing its family would "
            f"orphan every place that uses it, so {spec['family']!r} is refused"
        )
    if "name" in spec and isinstance(spec["name"], str):
        _check_free(root, style_id, spec["name"], excluding=style)

    _apply_spec(root, style, spec, pkg=pkg)
    return _style_info(style)


def clone_style(
    pkg: DocxPackage | Document,
    id_or_name: str,
    new_name: str,
    spec: Mapping[str, Any] | None = None,
) -> StyleInfo:
    """Copy a style under a new name, then apply `spec` to the copy.

    The copy is made of the ``w:style`` element itself, not of the properties
    this module decodes: the borders, the shading, the east-asian typography and
    everything else the minimal model does not report come along untouched, which
    is the whole reason to clone a style rather than describe it again.

    Two things are deliberately not copied.  ``w:link`` is dropped, because a
    link is reciprocal and the original's twin still points at the original.  The
    copy is marked as a custom style, because it is one, whatever the original
    claimed.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
        id_or_name: the style to copy.
        new_name: the ``w:name`` of the copy; its id is derived from it unless
            `spec` gives ``style_id``.
        spec: changes to apply to the copy, in :func:`create_style`'s vocabulary.

    Returns:
        The :class:`StyleInfo` of the copy.

    Raises:
        TypeError: if `spec` is not a mapping or a value is of the wrong type.
        ValueError: on an unknown key, on a ``family`` that differs from the
            original's, or on a value the schema cannot store.
        LocatorError: ``not_found`` if no style answers to `id_or_name`;
            ``already_exists`` if the new name or id is taken.
        PackageError: if the package has no style part.
    """
    spec = _check_spec({} if spec is None else spec)
    if not isinstance(new_name, str) or not new_name.strip():
        raise ValueError(f"'new_name' must be a non-empty style name, got {new_name!r}")
    root = _require_styles_root(pkg)
    source = _find_style(root, id_or_name)
    family = source.get(_W_TYPE) or "paragraph"
    if "family" in spec and spec["family"] != family:
        raise ValueError(
            f"{id_or_name!r} is a {family} style; a copy of it is one too, so "
            f"{spec['family']!r} is refused"
        )
    given_id = spec.get("style_id")
    if given_id is None:
        style_id = derive_style_id(new_name)
    elif not isinstance(given_id, str) or not given_id.strip():
        raise ValueError(f"'style_id' must be a non-empty style id, got {given_id!r}")
    else:
        style_id = given_id
    _check_free(root, style_id, new_name)

    clone = copy.deepcopy(source)
    clone.set(_W_STYLE_ID, style_id)
    clone.set(_W_CUSTOM_STYLE, "1")
    _drop_children(clone, _W_LINK)
    _drop_children(clone, _W_NAME)
    _ensure_child(clone, _W_NAME, _STYLE_RANK).set(_W_VAL, new_name)
    source.addnext(clone)
    try:
        _apply_spec(root, clone, spec, pkg=pkg)
    except BaseException:
        root.remove(clone)
        raise
    return _style_info(clone)


# --------------------------------------------------------------------------------------
# Deleting
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class StyleDeletion:
    """What :func:`delete_style` did, in the two places a style is named.

    Attributes:
        style_id: the id that is gone.
        reassigned_to: the id the places that used it now name, and the one its
            ``basedOn``/``next``/``link`` references were repointed to; ``None``
            when the style was deleted without a reassignment.
        usages: where it was used, as :func:`find_style_usage` reported them
            *before* the reassignment -- the places whose style changed.
        references: the ids of the other styles whose ``basedOn``, ``next`` or
            ``link`` named the deleted style and was repointed or removed.  These
            are invisible from the body of the document and are the half of a
            deletion a caller cannot check for itself.
    """

    style_id: str
    reassigned_to: str | None
    usages: tuple[StyleUsage, ...]
    references: tuple[str, ...]


def _repoint_usages(pkg: DocxPackage, style_id: str, new_id: str) -> None:
    """Rewrite every ``w:pStyle``/``w:rStyle``/``w:tblStyle`` naming `style_id`."""
    for _, root in pkg.stories():
        for tag in (_W_PSTYLE, _W_RSTYLE, _W_TBL_STYLE):
            for element in root.iter(tag):
                if element.get(_W_VAL) == style_id:
                    element.set(_W_VAL, new_id)


def _repoint_references(
    root: etree._Element, style_id: str, new_id: str | None
) -> tuple[str, ...]:
    """Repoint -- or drop -- the style references naming `style_id`.

    A ``basedOn`` left pointing at a style that no longer exists is the very
    thing :func:`get_style` reports as a warning, so it is not left behind: it
    follows the reassignment when there is one, and is removed when there is
    none.
    """
    touched: list[str] = []
    for style in _styles_in(root):
        changed = False
        for tag in (_W_BASED_ON, _W_NEXT, _W_LINK):
            for child in style.findall(tag):
                if child.get(_W_VAL) != style_id:
                    continue
                if new_id is None:
                    style.remove(child)
                else:
                    child.set(_W_VAL, new_id)
                changed = True
        if changed:
            touched.append(style.get(_W_STYLE_ID) or "")
    return tuple(touched)


def delete_style(
    pkg: DocxPackage, id_or_name: str, reassign_to: str | None = None
) -> StyleDeletion:
    """Remove a style, after saying what happens to everything that used it.

    A style is not an isolated object: paragraphs, runs and tables name it, and
    other styles inherit from it.  Deleting it without dealing with both is how a
    document loses its formatting silently -- Word falls back to ``Normal`` for
    every orphaned paragraph and says nothing.  So a style anything uses is
    refused unless `reassign_to` says where those places go, and the references
    of the other styles are repointed in the same breath.

    Args:
        pkg: the package.  A :class:`DocxPackage` and not a python-docx
            ``Document``: finding the usages means walking every story, which is
            what this handle is for.
        id_or_name: the style to delete, by ``w:styleId`` or by ``w:name``.
        reassign_to: the style the places that used it should name instead, by id
            or by name.  It must be of the same family: a paragraph cannot point
            at a character style.

    Returns:
        A :class:`StyleDeletion`.

    Raises:
        TypeError: if `pkg` is not a :class:`DocxPackage`.
        ValueError: if `reassign_to` is the style being deleted, or if the style
            is the default of its family -- Word needs that one.
        LocatorError: ``not_found`` if either style is undefined;
            ``wrong_style_type`` if `reassign_to` is of another family;
            ``style_in_use`` if the style is used and `reassign_to` is ``None``.
        PackageError: if the package has no style part.
    """
    if not isinstance(pkg, DocxPackage):
        raise TypeError(
            "delete_style needs a DocxPackage: it walks every story to find the "
            f"usages, got {type(pkg).__name__}"
        )
    root = _require_styles_root(pkg)
    style = _find_style(root, id_or_name)
    style_id = style.get(_W_STYLE_ID) or ""
    family = style.get(_W_TYPE) or "paragraph"
    if style.get(_W_DEFAULT) not in (None, *_OFF):
        raise ValueError(
            f"style {style_id!r} is the default {family} style of this document; "
            "deleting it would leave every place that relies on the default "
            "without one"
        )

    new_id: str | None = None
    if reassign_to is not None:
        target = _find_style(root, reassign_to)
        new_id = target.get(_W_STYLE_ID) or ""
        if new_id == style_id:
            raise ValueError(
                f"'reassign_to' names {style_id!r}, the style being deleted"
            )
        target_family = target.get(_W_TYPE) or "paragraph"
        if target_family != family:
            raise LocatorError(
                "wrong_style_type",
                f"{style_id!r} is a {family} style and {new_id!r} is a "
                f"{target_family} style; the places that use the first cannot name "
                "the second",
            )

    usages = tuple(find_style_usage(pkg, style_id))
    if usages and new_id is None:
        raise LocatorError(
            "style_in_use",
            f"{len(usages)} place(s) use the style {style_id!r}; pass 'reassign_to' "
            "to say which style they should name instead.  doc_find_style_usage "
            "lists them",
        )
    if new_id is not None:
        _repoint_usages(pkg, style_id, new_id)
    references = _repoint_references(root, style_id, new_id)
    root.remove(style)
    return StyleDeletion(
        style_id=style_id,
        reassigned_to=new_id,
        usages=usages,
        references=references,
    )
