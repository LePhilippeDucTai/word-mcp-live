"""What a reader actually sees, and which layer decided it.

:mod:`word_document_server.engine.styles` answers "what does this style say".
:mod:`word_document_server.engine.format` answers "what does this run say".
Neither answers the question an agent asked to *change* something has to settle
first: this paragraph looks bold -- is the bold on the run, on its character
style, on the paragraph style, three levels up a ``basedOn`` chain, or in
``w:docDefaults``?  Clearing the wrong one either does nothing or changes eleven
other paragraphs.  :func:`effective_format` answers it: every property comes
back as ``{"value", "source"}``, the value a reader sees and the layer that put
it there.

The layers
----------
Bottom to top, the order Word applies (ECMA-376 §17.7.2):

``docDefaults``
    ``w:docDefaults`` of the style sheet.
``table_style:<id>``
    the ``w:tblStyle`` of the table the paragraph sits in, when it sits in one.
    **Not resolved in v1**: a table style's contribution depends on conditional
    formatting (``w:tblStylePr`` for the first row, the banded columns, ...)
    that this module does not read, so a property the table style declares is
    reported with the value ``"unresolved"`` rather than with a number that
    would be right on the body of the table and wrong on its header.  The
    reference is still named, which is the part an agent can act on.
``style:<id>`` (paragraph)
    the paragraph's ``w:pStyle`` -- or the style sheet's default paragraph style
    when it has none -- and its whole ``basedOn`` chain, root first.
``style:<id>`` (character)
    the run's ``w:rStyle`` and its chain.  Run properties only: a character
    style has no ``w:pPr``.
``direct``
    the ``w:rPr`` of the run and the ``w:pPr`` of the paragraph.

``theme``
    not a layer but an answer: when the layer that won expressed the property as
    a theme reference (``w:asciiTheme``, ``w:themeColor``) rather than as a
    literal, `source` is ``"theme"``, `theme` holds the reference and `from`
    holds the layer that carried it.  Resolved colours inherit the tolerance of
    :func:`~word_document_server.engine.theme.apply_tint_shade` (D-028): they
    are a rendering hint, and nothing here is ever written back.

Toggles are not overridden, they are XORed
------------------------------------------
``w:b``, ``w:i``, ``w:caps``, ``w:smallCaps``, ``w:strike``, ``w:dstrike``,
``w:outline``, ``w:shadow``, ``w:emboss``, ``w:imprint`` and ``w:vanish`` are
*toggle properties* (ECMA-376 §17.7.3): stated at several levels of the style
hierarchy, their effective value is the **exclusive or** of what each level
says, not the most specific one.  A bold style based on a bold style renders
*not* bold, which is the single most surprising thing about Word's formatting
model and the reason this module exists.  Direct formatting is outside that
rule: a ``w:rPr`` on the run wins outright.

When more than one style level stated a toggle, the entry carries a `sources`
list naming all of them -- otherwise ``{"value": false, "source":
"style:FixtureLeaf"}`` on a style whose XML plainly says ``<w:b/>`` reads as a
bug rather than as the XOR it is.

``w:dstrike`` is treated as a toggle here.  ECMA's own list in §17.7.3 leaves it
out; the surface this module serves groups it with the others, and the
difference only shows on a document that states a double strike at two style
levels at once.

Vocabulary
----------
Property names are :func:`~word_document_server.engine.styles.decode_rpr`'s and
:func:`~word_document_server.engine.styles.decode_ppr`'s, which are
:func:`~word_document_server.engine.format.apply_rpr`'s -- ``bold``,
``size_pt``, ``font``, ``color``, ``alignment`` -- so what is read here can be
handed back to the tool that writes it.  Six toggles the minimal model does not
decode are added for §17.7.3 (:data:`EXTRA_RUN_TOGGLES`), spelled the same way:
``double_strike``, ``outline``, ``shadow``, ``emboss``, ``imprint``, ``vanish``.

``indent``, ``spacing`` and ``numbering`` hold several independent settings and
are inherited one by one, so they come back as a dict of
``{"value", "source"}`` per key: ``indent["left"]`` may come from the style
while ``indent["hanging"]`` comes from the paragraph.

What it leaves out
------------------
* the properties of the numbering level (``w:lvl/w:pPr``), which sit between the
  paragraph style and the direct formatting.  The numbering *link* is reported
  -- ``numbering.num_id`` and ``numbering.level``, with its own provenance --
  and :mod:`word_document_server.engine.numbering` says what that level draws.
* table style conditional formatting, as described above.
* everything outside the minimal model: borders, east-asian typography, shading.
  An absent property means "this module does not report it", never "the document
  does not set it".
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from lxml import etree

from word_document_server.engine.errors import PackageError
from word_document_server.engine.locators import Target
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.styles import (
    STYLES_PARTNAME,
    decode_ppr,
    decode_rpr,
    get_style,
    list_styles,
)
from word_document_server.engine.textmodel import segments
from word_document_server.engine.theme import Theme, read_theme
from word_document_server.engine.xmlns import qn

__all__ = [
    "EXTRA_RUN_TOGGLES",
    "MERGED_PROPERTIES",
    "MIXED",
    "PARAGRAPH_PROPERTIES",
    "RUN_PROPERTIES",
    "TOGGLE_PROPERTIES",
    "UNRESOLVED",
    "effective_format",
]

#: The value of a property this module knows is set but cannot compute -- today,
#: only a property a table style declares.  A sentinel string rather than
#: ``None``: "I cannot say" and "nobody sets it" are different answers, and only
#: one of them should make an agent stop and look.
UNRESOLVED = "unresolved"

#: The value of a property that is not the same everywhere in the range asked
#: about.  A caller that wants the detail narrows the range.
MIXED = "mixed"

#: Toggles ``decode_rpr`` already decodes, in its vocabulary.
_MODEL_RUN_TOGGLES = ("bold", "italic", "caps", "small_caps", "strike")

#: The six toggles of ECMA-376 §17.7.3 the minimal model does not decode ->
#: their ``w:rPr`` child.  Read here so that the XOR rule covers the whole set.
EXTRA_RUN_TOGGLES: dict[str, str] = {
    "double_strike": qn("w:dstrike"),
    "outline": qn("w:outline"),
    "shadow": qn("w:shadow"),
    "emboss": qn("w:emboss"),
    "imprint": qn("w:imprint"),
    "vanish": qn("w:vanish"),
}

#: Every run property whose layers combine by exclusive or.
TOGGLE_PROPERTIES: tuple[str, ...] = (*_MODEL_RUN_TOGGLES, *EXTRA_RUN_TOGGLES)

#: Run properties where the most specific layer simply wins.
_PLAIN_RUN_PROPERTIES = (
    "underline",
    "size_pt",
    "highlight",
    "superscript",
    "subscript",
    "font",
    "font_east_asia",
    "font_cs",
    "color",
)

#: Every run property reported, in report order.
RUN_PROPERTIES: tuple[str, ...] = (*TOGGLE_PROPERTIES, *_PLAIN_RUN_PROPERTIES)

#: Paragraph properties where the most specific layer wins.  Word's toggle rule
#: is about run properties; ``w:keepNext`` and friends override normally.
_PLAIN_PARAGRAPH_PROPERTIES = (
    "alignment",
    "outline_level",
    "keep_next",
    "keep_lines",
    "page_break_before",
    "widow_control",
    "contextual_spacing",
)

#: Paragraph properties inherited key by key rather than as a block -- a level
#: that sets ``spacing.after`` keeps the ``spacing.line`` it inherits.
MERGED_PROPERTIES: tuple[str, ...] = ("indent", "spacing", "numbering")

#: Every paragraph property reported, in report order.
PARAGRAPH_PROPERTIES: tuple[str, ...] = (
    *_PLAIN_PARAGRAPH_PROPERTIES,
    *MERGED_PROPERTIES,
)

#: Properties whose decoded value is a ``{"value", "theme"}`` pair, and how to
#: resolve the reference.
_THEMED_FONTS = ("font", "font_east_asia", "font_cs")

#: What ``w:val`` spells when a toggle means "off".  Word writes all four.
_OFF = frozenset({"0", "false", "off"})

_W_STYLES = qn("w:styles")
_W_STYLE = qn("w:style")
_W_STYLE_ID = qn("w:styleId")
_W_VAL = qn("w:val")
_W_RPR = qn("w:rPr")
_W_PPR = qn("w:pPr")
_W_PSTYLE = qn("w:pStyle")
_W_RSTYLE = qn("w:rStyle")
_W_DOC_DEFAULTS = qn("w:docDefaults")
_W_RPR_DEFAULT = qn("w:rPrDefault")
_W_PPR_DEFAULT = qn("w:pPrDefault")
_W_TBL = qn("w:tbl")
_W_TBL_PR = qn("w:tblPr")
_W_TBL_STYLE = qn("w:tblStyle")

#: Source label of the document defaults layer.
_DOC_DEFAULTS = "docDefaults"

#: Source label of the direct formatting layer.
_DIRECT = "direct"


class _Unresolved:
    """The sentinel a table style's contribution carries inside the fold.

    A distinct object rather than the :data:`UNRESOLVED` string so that a
    document whose style really does set ``highlight="unresolved"`` cannot be
    mistaken for one this module gave up on.
    """

    __slots__ = ()

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "<unresolved>"


_UNRESOLVED = _Unresolved()


@dataclass(frozen=True)
class _Layer:
    """One level of the cascade, decoded.

    Attributes:
        source: what the report calls it -- ``"docDefaults"``,
            ``"table_style:Grid"``, ``"style:Heading1"``, ``"direct"``.
        run: the run properties this level alone states.
        paragraph: the paragraph properties it alone states.
    """

    source: str
    run: dict[str, Any]
    paragraph: dict[str, Any]


# --------------------------------------------------------------------------------------
# Reading the style sheet
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


def _style_elements(pkg: DocxPackage) -> dict[str, etree._Element]:
    """Style id -> its ``w:style`` element.

    Needed on top of :func:`~word_document_server.engine.styles.get_style`, which
    hands back decoded properties: :data:`EXTRA_RUN_TOGGLES` is outside the
    minimal model that decoder covers, so those six children are read from the
    element itself.  The *chain* still comes from ``get_style``; this is a
    supplement to it, not a second walk of ``basedOn``.
    """
    root = _styles_root(pkg)
    if root is None:
        return {}
    found: dict[str, etree._Element] = {}
    for style in root.findall(_W_STYLE):
        style_id = style.get(_W_STYLE_ID)
        if style_id:
            found.setdefault(style_id, style)
    return found


def _extra_toggles(properties: etree._Element | None) -> dict[str, bool]:
    """The six §17.7.3 toggles ``decode_rpr`` does not cover, from a ``w:rPr``."""
    if properties is None:
        return {}
    found: dict[str, bool] = {}
    for name, tag in EXTRA_RUN_TOGGLES.items():
        child = properties.find(tag)
        if child is not None:
            found[name] = child.get(_W_VAL) not in _OFF
    return found


def _decode_run(properties: etree._Element | None) -> dict[str, Any]:
    """A ``w:rPr`` in this module's vocabulary: the minimal model plus §17.7.3.

    ``char_style`` is dropped: the ``w:rStyle`` selects a layer, it is not a
    property that layer contributes, and it is reported on its own.
    """
    decoded = decode_rpr(properties)
    decoded.pop("char_style", None)
    decoded.update(_extra_toggles(properties))
    return decoded


def _level_run_props(
    level_props: dict[str, Any], element: etree._Element | None
) -> dict[str, Any]:
    """A :class:`~...styles.StyleLevel`'s run properties, plus its extra toggles."""
    decoded = dict(level_props)
    decoded.pop("char_style", None)
    decoded.update(_extra_toggles(None if element is None else element.find(_W_RPR)))
    return decoded


def _doc_defaults_layer(pkg: DocxPackage) -> _Layer:
    """The ``w:docDefaults`` level -- the floor every cascade stands on.

    Read here rather than taken from a ``basedOn`` chain because a document may
    define no style at all, and because every chain
    :func:`~word_document_server.engine.styles.get_style` returns ends with the
    same level: folding two chains would otherwise apply it twice, which for a
    toggle means XORing it with itself and losing it.
    """
    root = _styles_root(pkg)
    defaults = None if root is None else root.find(_W_DOC_DEFAULTS)
    run_properties = None
    paragraph_properties = None
    if defaults is not None:
        run_default = defaults.find(_W_RPR_DEFAULT)
        if run_default is not None:
            run_properties = run_default.find(_W_RPR)
        paragraph_default = defaults.find(_W_PPR_DEFAULT)
        if paragraph_default is not None:
            paragraph_properties = paragraph_default.find(_W_PPR)
    return _Layer(
        source=_DOC_DEFAULTS,
        run=_decode_run(run_properties),
        paragraph=decode_ppr(paragraph_properties),
    )


def _chain_layers(
    pkg: DocxPackage,
    style_id: str,
    elements: dict[str, etree._Element],
    warnings: list[str],
) -> list[_Layer]:
    """The layers of `style_id` and its ``basedOn`` ancestors, **root first**.

    The ``w:docDefaults`` level every chain ends with is dropped: it is added
    once, by :func:`_doc_defaults_layer`.  A style the document does not define
    contributes no layer and one warning -- the reader is then reading a
    document Word itself would render from its built-in defaults.
    """
    try:
        detail = get_style(pkg, style_id)
    except PackageError:
        warnings.append(
            f"style {style_id!r} is referenced but this document does not define it; "
            "its contribution is missing from the effective format"
        )
        return []
    warnings.extend(detail.warnings)
    layers: list[_Layer] = []
    for level in detail.chain:
        if not level.style_id:
            continue  # the docDefaults level, added once elsewhere
        layers.append(
            _Layer(
                source=f"style:{level.style_id}",
                run=_level_run_props(level.run_props, elements.get(level.style_id)),
                paragraph=dict(level.paragraph_props),
            )
        )
    layers.reverse()
    return layers


def _unresolved_like(properties: dict[str, Any]) -> dict[str, Any]:
    """`properties` with every leaf value replaced by the unresolved sentinel.

    The *keys* are the honest part -- they say which properties the table style
    has an opinion about -- and the values are the part v1 cannot compute.
    """
    blanked: dict[str, Any] = {}
    for name, value in properties.items():
        if name in MERGED_PROPERTIES and isinstance(value, dict):
            blanked[name] = dict.fromkeys(value, _UNRESOLVED)
        else:
            blanked[name] = _UNRESOLVED
    return blanked


def _table_style_layer(
    pkg: DocxPackage, style_id: str, warnings: list[str]
) -> list[_Layer]:
    """One layer standing for everything the table style of a cell contributes."""
    try:
        detail = get_style(pkg, style_id)
    except PackageError:
        warnings.append(
            f"table style {style_id!r} is applied but this document does not define it"
        )
        return []
    run: dict[str, Any] = {}
    paragraph: dict[str, Any] = {}
    for level in detail.chain:
        if not level.style_id:
            continue
        run.update(_unresolved_like(level.run_props))
        paragraph.update(_unresolved_like(level.paragraph_props))
    warnings.append(
        f"this paragraph is in a table styled {style_id!r}; a table style's "
        "contribution depends on conditional formatting this version does not "
        f"resolve, so the properties it declares are reported as {UNRESOLVED!r}"
    )
    return [_Layer(source=f"table_style:{style_id}", run=run, paragraph=paragraph)]


# --------------------------------------------------------------------------------------
# Reading the paragraph
# --------------------------------------------------------------------------------------


def _child_val(element: etree._Element | None, tag: str) -> str | None:
    """The ``w:val`` of the child `tag`, or ``None`` when there is no such child."""
    if element is None:
        return None
    child = element.find(tag)
    return None if child is None else child.get(_W_VAL)


def _paragraph_style_id(pkg: DocxPackage, paragraph: etree._Element) -> str | None:
    """The style that governs `paragraph`: its ``w:pStyle``, or the sheet default.

    A paragraph without a ``w:pStyle`` is not unstyled -- Word applies the style
    flagged ``w:default="1"`` in the paragraph family, which is where the body
    font and the paragraph spacing of most documents live.
    """
    named = _child_val(paragraph.find(_W_PPR), _W_PSTYLE)
    if named:
        return named
    for info in list_styles(pkg, "paragraph"):
        if info.default:
            return info.style_id
    return None


def _table_style_id(paragraph: etree._Element) -> str | None:
    """The ``w:tblStyle`` of the innermost table `paragraph` sits in, if any."""
    ancestor = paragraph.getparent()
    while ancestor is not None:
        if ancestor.tag == _W_TBL:
            return _child_val(ancestor.find(_W_TBL_PR), _W_TBL_STYLE)
        ancestor = ancestor.getparent()
    return None


def _runs_in_span(
    paragraph: etree._Element, start: int, end: int
) -> list[etree._Element]:
    """The ``w:r`` elements the span ``[start, end)`` touches, in document order.

    Read through :func:`~word_document_server.engine.textmodel.segments`, so a
    run inside a hyperlink or a tracked insertion counts like any other, and
    nothing is split: asking what a range looks like must not modify the
    document, which is why this does not go through
    :func:`~word_document_server.engine.ranges.resolve`.

    An empty span takes the runs it sits on the boundary of, which is what an
    insertion point would inherit.
    """
    found: list[etree._Element] = []
    seen: set[int] = set()
    for segment in segments(paragraph):
        if end > start:
            touched = segment.end > start and segment.start < end
        else:
            touched = segment.start <= start <= segment.end
        if not touched:
            continue
        run = segment.run
        if run is None or id(run) in seen:
            continue
        seen.add(id(run))
        found.append(run)
    return found


# --------------------------------------------------------------------------------------
# Folding the layers
# --------------------------------------------------------------------------------------


def _entry(value: Any, source: str) -> dict[str, Any]:
    """One reported property."""
    return {"value": UNRESOLVED if value is _UNRESOLVED else value, "source": source}


def _fold_toggle(name: str, style_layers: list[_Layer], direct: _Layer) -> dict[str, Any] | None:
    """Combine one toggle across the cascade -- XOR of the styles, direct on top."""
    if name in direct.run:
        return _entry(bool(direct.run[name]), direct.source)
    stated = [(layer.source, layer.run[name]) for layer in style_layers if name in layer.run]
    if not stated:
        return None
    value = False
    for source, raw in stated:
        if raw is _UNRESOLVED:
            return _entry(_UNRESOLVED, source)
        value ^= bool(raw)
    reported = _entry(value, stated[-1][0])
    if len(stated) > 1:
        # Without this, "value: false" under a style whose XML says <w:b/> reads
        # as a bug rather than as §17.7.3 doing its job.
        reported["sources"] = [source for source, _ in stated]
    return reported


def _fold_plain(name: str, layers: list[_Layer], attribute: str) -> dict[str, Any] | None:
    """The most specific layer that states `name` wins."""
    winner: tuple[str, Any] | None = None
    for layer in layers:
        properties: dict[str, Any] = getattr(layer, attribute)
        if name in properties:
            winner = (layer.source, properties[name])
    if winner is None:
        return None
    return _entry(winner[1], winner[0])


def _fold_merged(name: str, layers: list[_Layer]) -> dict[str, Any] | None:
    """A multi-part property, resolved key by key so each key keeps its own source."""
    found: dict[str, Any] = {}
    for layer in layers:
        value = layer.paragraph.get(name)
        if not isinstance(value, dict):
            continue
        for key, raw in value.items():
            found[key] = _entry(raw, layer.source)
    return found or None


def _fold_run(style_layers: list[_Layer], direct: _Layer) -> dict[str, Any]:
    """Every run property of one run, in report order."""
    folded: dict[str, Any] = {}
    for name in TOGGLE_PROPERTIES:
        reported = _fold_toggle(name, style_layers, direct)
        if reported is not None:
            folded[name] = reported
    for name in _PLAIN_RUN_PROPERTIES:
        reported = _fold_plain(name, [*style_layers, direct], "run")
        if reported is not None:
            folded[name] = reported
    return folded


def _fold_paragraph(layers: list[_Layer]) -> dict[str, Any]:
    """Every paragraph property, in report order."""
    folded: dict[str, Any] = {}
    for name in _PLAIN_PARAGRAPH_PROPERTIES:
        reported = _fold_plain(name, layers, "paragraph")
        if reported is not None:
            folded[name] = reported
    for name in MERGED_PROPERTIES:
        merged = _fold_merged(name, layers)
        if merged is not None:
            folded[name] = merged
    return folded


# --------------------------------------------------------------------------------------
# Resolving theme references
# --------------------------------------------------------------------------------------


def _resolve_font(
    entry: dict[str, Any], theme: Theme | None, warnings: list[str]
) -> dict[str, Any]:
    """Turn a ``{"value", "theme"}`` font into the typeface a reader sees."""
    raw = entry["value"]
    if not isinstance(raw, dict):
        return entry
    literal = raw.get("value")
    reference = raw.get("theme")
    if not reference:
        return {"value": literal, "source": entry["source"]}
    resolved = None if theme is None else theme.resolve_font(reference)
    if resolved is None:
        # Not a ``ST_Theme`` value, or no theme at all.  An *empty* slot is a
        # different answer: the theme defines no typeface for that script, which
        # is what ``<a:ea typeface=""/>`` says and is not a failure to resolve.
        warnings.append(
            f"font theme reference {reference!r} does not resolve against this "
            "document's theme; the value the document caches is reported instead"
        )
        return {"value": literal, "source": entry["source"], "theme": reference}
    return {
        "value": resolved,
        "source": "theme",
        "theme": reference,
        "from": entry["source"],
    }


def _resolve_color(
    entry: dict[str, Any], theme: Theme | None, warnings: list[str]
) -> dict[str, Any]:
    """Turn a ``{"value", "theme", "tint", "shade"}`` colour into ``RRGGBB``.

    The resolved colour inherits the tolerance of
    :func:`~word_document_server.engine.theme.apply_tint_shade` (D-028): a
    rendering hint, never something to write back.
    """
    raw = entry["value"]
    if not isinstance(raw, dict):
        return entry
    literal = raw.get("value")
    reference = raw.get("theme")
    if not reference:
        return {"value": literal, "source": entry["source"]}
    resolved = None
    explained = False
    if theme is not None:
        try:
            resolved = theme.resolve_color(reference, raw.get("tint"), raw.get("shade"))
        except (TypeError, ValueError) as exc:
            # A document carrying both a tint and a shade, or a malformed one:
            # the reason is worth quoting, so it replaces the generic message.
            warnings.append(f"theme colour {reference!r} could not be resolved: {exc}")
            explained = True
    if resolved is None:
        if not explained:
            warnings.append(
                f"colour theme reference {reference!r} does not resolve against this "
                "document's theme; the value the document caches is reported instead"
            )
        return {"value": literal, "source": entry["source"], "theme": reference}
    return {
        "value": resolved,
        "source": "theme",
        "theme": reference,
        "from": entry["source"],
    }


def _resolve_themed(
    folded: dict[str, Any], theme: Theme | None, warnings: list[str]
) -> dict[str, Any]:
    """Resolve every themed property of one run's folded properties."""
    for name in _THEMED_FONTS:
        if name in folded:
            folded[name] = _resolve_font(folded[name], theme, warnings)
    if "color" in folded:
        folded["color"] = _resolve_color(folded["color"], theme, warnings)
    return folded


# --------------------------------------------------------------------------------------
# Folding several runs into one answer
# --------------------------------------------------------------------------------------


def _merge_runs(per_run: list[dict[str, Any]]) -> dict[str, Any]:
    """One answer for the whole range: the common value, or :data:`MIXED`.

    A property stated on one run and absent from the next is mixed too -- the
    two runs do not look the same, which is the question being asked.
    """
    if not per_run:
        return {}
    if len(per_run) == 1:
        return dict(per_run[0])
    names: list[str] = []
    for folded in per_run:
        for name in folded:
            if name not in names:
                names.append(name)
    merged: dict[str, Any] = {}
    for name in names:
        entries = [folded.get(name) for folded in per_run]
        first = entries[0]
        if all(entry == first for entry in entries) and first is not None:
            merged[name] = first
            continue
        # The common source survives a disagreement about the value -- but only
        # when every run actually has one.  A run where no layer states the
        # property at all has no source to agree with, and reporting the other
        # runs' source would name a layer that is not acting on the whole range.
        sources = {None if entry is None else entry["source"] for entry in entries}
        merged[name] = {
            "value": MIXED,
            "source": sources.pop() if len(sources) == 1 else MIXED,
        }
    # Report order, not discovery order.
    return {name: merged[name] for name in RUN_PROPERTIES if name in merged}


# --------------------------------------------------------------------------------------
# The question
# --------------------------------------------------------------------------------------


def effective_format(pkg: DocxPackage, target: Target) -> dict[str, Any]:
    """What a reader sees over `target`, and which layer decided each property.

    Args:
        pkg: the package `target` was resolved against.
        target: a paragraph or a span inside one, as
            :func:`~word_document_server.engine.locators.resolve` or
            :func:`~word_document_server.engine.find.find` names it.  A whole
            paragraph is the span ``0..len(visible_text)``, which is what a
            ``{"paragraph": i}`` locator resolves to.

    Returns:
        A dict with:

        ``story``, ``index``, ``start``, ``end``
            where the answer applies.  `index` is the V2 paragraph index, or
            ``None`` in a table cell or a text box (D-016).
        ``paragraph_style``
            the style that governs the paragraph -- its ``w:pStyle``, or the
            style sheet's default paragraph style when it names none.
        ``char_styles``
            the ``w:rStyle`` ids the range's runs carry, in document order.
        ``table_style``
            the ``w:tblStyle`` of the table the paragraph sits in, or ``None``.
        ``run``, ``paragraph``
            property name -> ``{"value", "source"}``.  See the module docstring
            for the sources, for the XOR rule on toggles, and for ``indent``,
            ``spacing`` and ``numbering``, which nest one such entry per key.
        ``warnings``
            what made the reading incomplete: an undefined style, a ``basedOn``
            chain that loops, a table style, a theme reference that does not
            resolve.

        A property no layer states is absent, which means "nothing sets it";
        ``"unresolved"`` means "something sets it and this version cannot say
        what"; ``"mixed"`` means "not the same everywhere in this range".

    Raises:
        TypeError: if `target` is not a
            :class:`~word_document_server.engine.locators.Target`.
    """
    if not isinstance(target, Target):
        raise TypeError(
            f"expected a resolved locator Target, got {type(target).__name__}"
        )
    warnings: list[str] = []
    theme = read_theme(pkg)
    elements = _style_elements(pkg)
    paragraph = target.paragraph
    paragraph_properties = paragraph.find(_W_PPR)

    base: list[_Layer] = [_doc_defaults_layer(pkg)]
    table_style = _table_style_id(paragraph)
    if table_style:
        base.extend(_table_style_layer(pkg, table_style, warnings))
    paragraph_style = _paragraph_style_id(pkg, paragraph)
    if paragraph_style:
        base.extend(_chain_layers(pkg, paragraph_style, elements, warnings))

    direct_paragraph = _Layer(
        source=_DIRECT, run={}, paragraph=decode_ppr(paragraph_properties)
    )

    runs = _runs_in_span(paragraph, target.start, target.end)
    sources: list[etree._Element | None]
    if runs:
        sources = list(runs)
    else:
        # An empty paragraph, or a span that touches no run: the paragraph mark
        # carries the formatting a reader would see if they typed there.
        sources = [None]

    char_styles: list[str] = []
    per_run: list[dict[str, Any]] = []
    for run in sources:
        own = (
            paragraph_properties.find(_W_RPR)
            if run is None and paragraph_properties is not None
            else (None if run is None else run.find(_W_RPR))
        )
        layers = list(base)
        char_style = _child_val(own, _W_RSTYLE)
        if char_style:
            if char_style not in char_styles:
                char_styles.append(char_style)
            layers.extend(_chain_layers(pkg, char_style, elements, warnings))
        direct_run = _Layer(source=_DIRECT, run=_decode_run(own), paragraph={})
        per_run.append(_resolve_themed(_fold_run(layers, direct_run), theme, warnings))

    seen: set[str] = set()
    unique_warnings = [
        warning for warning in warnings if not (warning in seen or seen.add(warning))
    ]
    return {
        "story": target.story,
        "index": target.index,
        "start": target.start,
        "end": target.end,
        "paragraph_style": paragraph_style,
        "char_styles": char_styles,
        "table_style": table_style,
        "run": _merge_runs(per_run),
        "paragraph": _fold_paragraph([*base, direct_paragraph]),
        "warnings": unique_warnings,
    }
