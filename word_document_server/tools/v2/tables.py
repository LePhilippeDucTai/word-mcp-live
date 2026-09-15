"""The V2 table-style tool: naming an existing gallery entry on a table.

A single call, :func:`doc_apply_table_style`, covers both directions
:mod:`word_document_server.engine.table_styles` offers: passing a `style`
applies it, passing ``None`` clears whatever style the table had -- the same
convention :func:`~word_document_server.tools.v2.text.doc_format_range` uses
for its `patch` ("a property set to null is removed").

Addressing
----------
`table_index` numbers the ``w:tbl`` elements of the main document body, in
document order, nested tables included -- the same space
:meth:`~word_document_server.engine.styles.StyleUsage.table` reports and
:func:`~word_document_server.engine.locators.resolve`'s ``table`` locator form
addresses.  There is no `story` parameter: a table in a header, footer or note
is out of reach of this tool.
"""

from __future__ import annotations

from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine.errors import LocatorError
from word_document_server.engine.locators import MAIN_STORY, _tables
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.table_styles import (
    apply_table_style,
    clear_table_style,
    resolve_look,
)
from word_document_server.tools.v2.registry import ToolSpec

__all__ = ["TOOLS", "doc_apply_table_style"]


def _table_at(pkg: DocxPackage, table_index: object) -> Any:
    """The ``w:tbl`` element of the main document body at `table_index`.

    Raises:
        ValueError: if `table_index` is not a non-negative integer.
        LocatorError: ``not_found`` if the body has no table there.
    """
    if isinstance(table_index, bool) or not isinstance(table_index, int):
        raise ValueError(
            f"'table_index' must be an integer, got {type(table_index).__name__}"
        )
    if table_index < 0:
        raise ValueError(f"'table_index' must be 0 or more, got {table_index}")
    root = dict(pkg.stories())[MAIN_STORY]
    tables = _tables(root)
    if table_index >= len(tables):
        raise LocatorError(
            "not_found",
            f"the document body has {len(tables)} table(s), so table "
            f"{table_index} does not exist",
        )
    return tables[table_index]


def doc_apply_table_style(
    filename: str,
    table_index: int,
    style: str | None,
    look: dict[str, bool] | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Apply an existing table style to a table, or clear the one it has.

    No table style is created here: `style` must already be one
    ``doc_list_styles(filename, family="table")`` reports, or the call is
    refused. Table style *creation* is out of scope for this tool.

    Args:
        filename: path to the .docx file.
        table_index: which table of the document body, 0-based, in document
            order (nested tables included). Headers, footers and notes are not
            addressed by this tool.
        style: the ``styleId`` of an existing table style to apply, or `None`
            to remove the table's current style, leaving its ``tblLook`` flags
            as they are.
        look: which conditional formats the style may draw for this table --
            `first_row`, `last_row`, `first_column`, `last_column`,
            `no_h_band`, `no_v_band`, each true or false. A key left out is
            false; ignored when `style` is `None`.
        dry_run: compute and report the change without writing the file.

    Returns:
        The usual report, plus `table_index`, `style` (echoed back, `None`
        when cleared) and `look` (the six flags actually written, or `None`
        when cleared).

    Raises:
        ValueError: if `table_index` is not a non-negative integer, if `style`
            is an empty string, or if `look` names an unknown flag.
    """
    pkg = DocxPackage.open(filename)
    table = _table_at(pkg, table_index)

    if style is None:
        clear_table_style(table)
        resolved_look = None
    else:
        apply_table_style(pkg, table, style, look)
        resolved_look = resolve_look(look)

    saved = not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "table_index": table_index,
        "style": style,
        "look": resolved_look,
        "changes": [],
        "warnings": [],
    }


TOOLS = [
    ToolSpec(
        fn=doc_apply_table_style,
        annotations=ToolAnnotations(title="Apply Table Style", destructiveHint=True),
        tags=frozenset({"v2", "write"}),
    ),
]
