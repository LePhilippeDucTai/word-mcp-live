"""Reading the style sheet: what a style is, what it inherits, where it is used.

Direct formatting is only half of what a run looks like; the other half is the
style it names and everything that style inherits.  This module reads that half.
It is read-only by design: J05 gives agents the vocabulary to *see* the style
sheet -- :mod:`word_document_server.engine.format` is where writing lives, and
it writes ``w:rPr``, never ``w:style``.

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
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from lxml import etree

from word_document_server.engine.errors import PackageError
from word_document_server.engine.locators import _tables, indexed_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import segments, visible_text
from word_document_server.engine.xmlns import qn

__all__ = [
    "STYLES_PARTNAME",
    "USAGE_PREVIEW",
    "StyleDetail",
    "StyleInfo",
    "StyleLevel",
    "StyleUsage",
    "find_style_usage",
    "get_style",
    "list_styles",
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


def _style_elements(pkg: DocxPackage) -> list[etree._Element]:
    """Every ``w:style`` of the style sheet, in document order.

    A ``w:style`` without a ``w:styleId`` is skipped: nothing can reference it,
    so reporting it would only offer an id no caller can use.
    """
    root = _styles_root(pkg)
    if root is None:
        return []
    return [style for style in root.findall(_W_STYLE) if style.get(_W_STYLE_ID)]


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
