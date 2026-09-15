"""Applying table styles: naming an existing gallery entry on a table.

A table does not carry its borders and shading the way a paragraph carries its
font: a ``w:tbl`` mostly *points* at a ``w:style`` of family ``table``
(``w:tblPr/w:tblStyle``), and the style itself defines up to twelve conditional
formats -- one whole-table default and eleven exceptions for the header row,
the banded rows, the corner cells and so on (``w:tblStylePr/@w:type``, ECMA-376
§17.4.65).  ``w:tblPr/w:tblLook`` (§17.4.56) is the second half of the picture:
it says *which* of those exceptions this particular table turns on, as a
four-hex-digit bitmask (``w:val``) mirrored by six explicit boolean attributes
that Word writes side by side with it for readers that do not decode the mask.

This module reads which table styles a document defines and what conditional
formats they carry (:func:`list_table_styles`), and writes the two elements
that apply one to a table (:func:`apply_table_style`) or remove one
(:func:`clear_table_style`).  It does not create or edit a table style: R-007
leaves that out of scope, so a `style_id` this module is asked to apply must
already be one :func:`~word_document_server.engine.styles.list_styles` reports
for the ``table`` family.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass

from lxml import etree

from word_document_server.engine.errors import PackageError
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.styles import STYLES_PARTNAME, list_styles
from word_document_server.engine.xmlns import qn

__all__ = [
    "LOOK_FLAGS",
    "TBL_PR_ORDER",
    "TableStyleInfo",
    "apply_table_style",
    "clear_table_style",
    "list_table_styles",
    "resolve_look",
]

_W_TBL = qn("w:tbl")
_W_STYLES = qn("w:styles")
_W_STYLE = qn("w:style")
_W_STYLE_ID = qn("w:styleId")
_W_TYPE = qn("w:type")
_W_TBL_STYLE_PR = qn("w:tblStylePr")

_W_TBL_PR = qn("w:tblPr")
_W_TBL_STYLE = qn("w:tblStyle")
_W_TBL_LOOK = qn("w:tblLook")
_W_VAL = qn("w:val")

#: Children of ``w:tblPr`` in schema order (ECMA-376 §17.4.59, ``CT_TblPrBase``
#: plus ``w:tblPrChange``).  Only ``w:tblStyle`` and ``w:tblLook`` are ever
#: created here; the rest of the sequence is kept so an element this module did
#: not write -- ``w:tblW``, ``w:tblBorders``, a caption -- is never displaced.
TBL_PR_ORDER: tuple[str, ...] = (
    "w:tblStyle",
    "w:tblpPr",
    "w:tblOverlap",
    "w:bidiVisual",
    "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize",
    "w:tblW",
    "w:jc",
    "w:tblCellSpacing",
    "w:tblInd",
    "w:tblBorders",
    "w:shd",
    "w:tblLayout",
    "w:tblCellMar",
    "w:tblLook",
    "w:tblCaption",
    "w:tblDescription",
    "w:tblPrChange",
)

_TBL_PR_RANK: dict[str, int] = {qn(tag): rank for rank, tag in enumerate(TBL_PR_ORDER)}

#: ``look`` key -> (``w:tblLook`` boolean attribute, its bit of the ``w:val``
#: mask).  Order matches the attribute order Word itself writes, which is also
#: the order :func:`apply_table_style` sets them in.  Bit values are ECMA-376
#: §17.18.92 (``ST_TblLook``), confirmed against a LibreOffice-written
#: ``w:val="04A0"`` paired with ``firstRow=1 firstColumn=1 noVBand=1``.
LOOK_FLAGS: dict[str, tuple[str, int]] = {
    "first_row": ("firstRow", 0x0020),
    "last_row": ("lastRow", 0x0040),
    "first_column": ("firstColumn", 0x0080),
    "last_column": ("lastColumn", 0x0100),
    "no_h_band": ("noHBand", 0x0200),
    "no_v_band": ("noVBand", 0x0400),
}


# --------------------------------------------------------------------------------------
# Placing children in schema order
# --------------------------------------------------------------------------------------


def _place(parent: etree._Element, child: etree._Element) -> None:
    """Move `child` to its :data:`TBL_PR_ORDER` position among `parent`'s children."""
    position = _TBL_PR_RANK[child.tag]
    for index, existing in enumerate(parent):
        if existing is child:
            continue
        other = _TBL_PR_RANK.get(existing.tag)
        if other is not None and other > position:
            parent.insert(index, child)
            return


def _ensure_tbl_pr(table: etree._Element) -> etree._Element:
    """Return `table`'s ``w:tblPr``, creating it as the first child if absent."""
    tbl_pr = table.find(_W_TBL_PR)
    if tbl_pr is not None:
        return tbl_pr
    tbl_pr = etree.Element(_W_TBL_PR)
    table.insert(0, tbl_pr)
    return tbl_pr


def _ensure_child(parent: etree._Element, tag: str) -> etree._Element:
    """Return the first `tag` child of `parent`, creating it in order if absent."""
    found = parent.find(tag)
    if found is not None:
        return found
    child = etree.SubElement(parent, tag)
    _place(parent, child)
    return child


# --------------------------------------------------------------------------------------
# Reading the table styles a document defines
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class TableStyleInfo:
    """One table style a document defines: identity plus what it varies.

    Attributes:
        style_id: the ``w:styleId``, what :func:`apply_table_style` takes.
        name: the ``w:name``, falling back to `style_id` like
            :class:`~word_document_server.engine.styles.StyleInfo` does.
        based_on: the ``w:basedOn`` this style inherits from, or ``None``.
        conditional_formats: the ``w:tblStylePr/@w:type`` this style defines,
            in the order they appear -- the exceptions ``w:tblLook`` can turn on
            or off for a table using this style.  ``"wholeTable"`` is the
            baseline every style is expected to carry; its absence means the
            style relies entirely on what it inherits through `based_on`.
    """

    style_id: str
    name: str
    based_on: str | None
    conditional_formats: tuple[str, ...]


def _styles_root(pkg: DocxPackage) -> etree._Element | None:
    """The live ``w:styles`` root, or ``None`` when the package has no style sheet.

    Reads through :class:`DocxPackage`'s own public API rather than through
    :mod:`word_document_server.engine.styles`: that module reads and decodes
    styles, it is not this module's to extend with a raw-XML accessor.
    """
    part = pkg.find_part(STYLES_PARTNAME)
    if part is None:
        return None
    root = pkg.root_of(part)
    return root if root.tag == _W_STYLES else None


def _table_style_elements(pkg: DocxPackage) -> dict[str, etree._Element]:
    """``w:styleId`` -> ``w:style`` element, for every style of family ``table``."""
    root = _styles_root(pkg)
    if root is None:
        return {}
    found: dict[str, etree._Element] = {}
    for style in root.findall(_W_STYLE):
        style_id = style.get(_W_STYLE_ID)
        if style_id and style.get(_W_TYPE) == "table":
            found[style_id] = style
    return found


def _conditional_formats(style: etree._Element) -> tuple[str, ...]:
    """The ``w:tblStylePr/@w:type`` values `style` defines, in document order."""
    return tuple(
        kind
        for child in style.findall(_W_TBL_STYLE_PR)
        if (kind := child.get(_W_TYPE)) is not None
    )


def list_table_styles(pkg: DocxPackage) -> list[TableStyleInfo]:
    """Every table style the document defines, in style-sheet order.

    Args:
        pkg: the package to read.

    Returns:
        One :class:`TableStyleInfo` per ``w:style`` of family ``table``.  Empty
        when the package has no style sheet or defines no table style.
    """
    elements = _table_style_elements(pkg)
    found: list[TableStyleInfo] = []
    for info in list_styles(pkg, family="table"):
        element = elements.get(info.style_id)
        conditional = () if element is None else _conditional_formats(element)
        found.append(
            TableStyleInfo(
                style_id=info.style_id,
                name=info.name,
                based_on=info.based_on,
                conditional_formats=conditional,
            )
        )
    return found


# --------------------------------------------------------------------------------------
# w:tblLook
# --------------------------------------------------------------------------------------


def resolve_look(look: Mapping[str, object] | None) -> dict[str, bool]:
    """The six :data:`LOOK_FLAGS` keys, each resolved to ``True``/``False``.

    A key `look` does not mention resolves to ``False`` -- the schema default,
    meaning that exception is off -- so ``None`` (or ``{}``) resolves to every
    flag off, the plain "whole table only" look.

    Args:
        look: a subset of :data:`LOOK_FLAGS`' keys, or ``None``.

    Returns:
        A dict with all six keys, values coerced to ``bool``.

    Raises:
        ValueError: if `look` names a key that is not one of :data:`LOOK_FLAGS`.
    """
    given = {} if look is None else look
    unknown = sorted(set(given) - set(LOOK_FLAGS))
    if unknown:
        raise ValueError(
            f"unknown look flag(s) {unknown}; known: {sorted(LOOK_FLAGS)}"
        )
    return {name: bool(given.get(name, False)) for name in LOOK_FLAGS}


def _set_tbl_look(tbl_pr: etree._Element, values: Mapping[str, bool]) -> None:
    """Write ``w:tblLook`` under `tbl_pr`: the mask and the six explicit flags.

    Both halves are written together, as Word itself does, so a reader that
    only decodes one of them still sees the right answer. `values` is already
    resolved (see :func:`resolve_look`): this helper never fails, so it is only
    called once every check that could reject the call has passed.
    """
    mask = 0
    for name, (_, bit) in LOOK_FLAGS.items():
        if values[name]:
            mask |= bit
    element = _ensure_child(tbl_pr, _W_TBL_LOOK)
    element.set(_W_VAL, f"{mask:04X}")
    for name, (attribute, _) in LOOK_FLAGS.items():
        element.set(qn(f"w:{attribute}"), "1" if values[name] else "0")


# --------------------------------------------------------------------------------------
# Applying and clearing
# --------------------------------------------------------------------------------------


def _require_table(table: etree._Element) -> None:
    if not isinstance(table, etree._Element) or table.tag != _W_TBL:
        got = getattr(table, "tag", type(table).__name__)
        raise TypeError(f"'table' must be a w:tbl element, got {got!r}")


def apply_table_style(
    pkg: DocxPackage,
    table: etree._Element,
    style_id: str,
    look: Mapping[str, object] | None = None,
) -> None:
    """Apply an existing table style to `table`, with its ``w:tblLook`` flags.

    Writes ``w:tblPr/w:tblStyle`` and ``w:tblPr/w:tblLook``, creating
    ``w:tblPr`` as the table's first child if it has none; both elements land
    at their :data:`TBL_PR_ORDER` position, leaving every other child of
    ``w:tblPr`` where it was.  No table style is created: `style_id` must
    already be one the document's style sheet defines for family ``table``.
    Every check runs before anything is written: a call this function refuses
    leaves `table` exactly as it was.

    Args:
        pkg: the package `table` belongs to, whose style sheet `style_id` is
            checked against.
        table: a ``w:tbl`` element.
        style_id: the ``w:styleId`` of an existing table style.
        look: which of :data:`LOOK_FLAGS` to turn on -- `first_row`,
            `last_row`, `first_column`, `last_column`, `no_h_band`,
            `no_v_band`. A key left out resolves to ``False``; ``None`` turns
            every flag off.

    Raises:
        TypeError: if `table` is not a ``w:tbl`` element.
        ValueError: if `style_id` is empty, or `look` names an unknown flag.
        PackageError: if the document defines no table style `style_id`.
    """
    _require_table(table)
    if not isinstance(style_id, str) or not style_id.strip():
        raise ValueError("a table style is applied by a non-empty style id")
    values = resolve_look(look)
    known = {info.style_id for info in list_styles(pkg, family="table")}
    if style_id not in known:
        raise PackageError(
            f"this document defines no table style {style_id!r}; it defines "
            f"{sorted(known)[:10]}"
        )

    tbl_pr = _ensure_tbl_pr(table)
    style_element = _ensure_child(tbl_pr, _W_TBL_STYLE)
    style_element.set(_W_VAL, style_id)
    _set_tbl_look(tbl_pr, values)


def clear_table_style(table: etree._Element) -> None:
    """Remove `table`'s style, leaving ``w:tblLook`` and everything else as is.

    A table with no ``w:tblPr`` or no ``w:tblStyle`` is left unchanged: clearing
    a style that is not there is not an error, it is already the answer asked
    for.

    Args:
        table: a ``w:tbl`` element.

    Raises:
        TypeError: if `table` is not a ``w:tbl`` element.
    """
    _require_table(table)
    tbl_pr = table.find(_W_TBL_PR)
    if tbl_pr is None:
        return
    style_element = tbl_pr.find(_W_TBL_STYLE)
    if style_element is not None:
        tbl_pr.remove(style_element)
