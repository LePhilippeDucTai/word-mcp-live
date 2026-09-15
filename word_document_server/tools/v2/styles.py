"""The three V2 style tools: list the style sheet, read one style, find its uses.

They answer the half of a document's appearance that ``doc_inspect`` and
``doc_format_range`` cannot: what the styles *say*.  An agent asked to make a
document consistent needs to know that ``FixtureLeaf`` is based on
``FixtureBranch``, that the bold comes from there and not from the paragraph,
and which eleven paragraphs would change if the style did.  That is
:func:`doc_list_styles`, :func:`doc_get_style` and :func:`doc_find_style_usage`.

All three are read-only.  Writing styles is not part of this surface: a style
edit changes every paragraph that uses it at once, and the tool that does it
has to say which ones before it does -- which is what
:func:`doc_find_style_usage` is for.

Theme references
----------------
A style property that follows the theme is reported as
``{"value": "Aptos", "theme": "minorHAnsi"}``: the literal the document caches
*and* the reference.  The document's theme comes back in the same answer, under
`theme`, so the reference can be resolved without a second call -- and so an
agent can tell "this style pins Aptos" from "this style follows the theme, which
currently happens to be Aptos".
"""

from __future__ import annotations

import dataclasses
from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine.errors import LocatorError, PackageError
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.styles import (
    STYLE_FAMILIES,
    find_style_usage,
    get_style,
    list_styles,
)
from word_document_server.engine.theme import Theme, read_theme
from word_document_server.tools.v2.registry import ToolSpec

__all__ = [
    "TOOLS",
    "doc_find_style_usage",
    "doc_get_style",
    "doc_list_styles",
]


def _theme_report(theme: Theme | None) -> dict[str, Any] | None:
    """The document theme as plain data, or ``None`` when it defines none."""
    if theme is None:
        return None
    return {
        "name": theme.name,
        "major_fonts": dataclasses.asdict(theme.major),
        "minor_fonts": dataclasses.asdict(theme.minor),
        "colors": dict(theme.colors),
        "color_map": dict(theme.color_map),
    }


def doc_list_styles(filename: str, family: str | None = None) -> dict[str, Any]:
    """List the styles a Word document defines, with their inheritance wiring.

    Read this before applying a style: it is the only way to know which ids the
    document actually has, and `based_on` / `next_style` / `link` say how they
    relate. It does not say which styles are *used* -- `doc_find_style_usage`
    answers that for one style, and `doc_inspect` reports the styles the body
    references.

    Args:
        filename: path to the .docx file.
        family: keep only one family: `paragraph`, `character`, `table` or
            `numbering`. Omit it to list them all.

    Returns:
        The usual report, plus `styles` -- one entry per style with `style_id`
        (what `w:pStyle` references and what every other style tool takes),
        `name` (what Word's style pane shows), `family`, `builtin`, `default`,
        `based_on`, `next_style`, `link`, `ui_priority`, `q_format`, `hidden`
        and `semi_hidden` -- and `theme`, the document's theme fonts and colour
        scheme, or null when it defines none.
    """
    if family is not None and family not in STYLE_FAMILIES:
        raise ValueError(
            f"unknown style family {family!r}; known: {', '.join(STYLE_FAMILIES)}"
        )
    pkg = DocxPackage.open(filename)
    return {
        "styles": [dataclasses.asdict(info) for info in list_styles(pkg, family)],
        "theme": _theme_report(read_theme(pkg)),
    }


def doc_get_style(filename: str, style: str) -> dict[str, Any]:
    """Read one style: what it sets, what it inherits, and what the two resolve to.

    Args:
        filename: path to the .docx file.
        style: the style id (`Heading1`) or its name (`heading 1`). The id is
            tried first.

    Returns:
        The usual report, plus `style` and `theme`.

        `style` holds `info` (the same fields `doc_list_styles` reports),
        `run_props` and `paragraph_props` (what this style alone sets),
        `chain` (one level per step of the `based_on` chain, most specific
        first, ending with the document defaults) and `resolved` -- `run` and
        `paragraph`, the chain applied from the defaults up.

        `resolved` is what the style contributes, not what a reader sees: the
        direct formatting on a run and the numbering level's own properties sit
        on top of it and are not included.

        Properties use the same names `doc_format_range` takes (`bold`,
        `size_pt`, `font`, `color`, `alignment`, `indent`, `spacing`), so what
        is read here can be applied back without translation. A property that
        follows the theme is `{"value": ..., "theme": ...}`; resolve the
        reference against `theme`.
    """
    pkg = DocxPackage.open(filename)
    try:
        detail = get_style(pkg, style)
    except PackageError as exc:
        # D-022: an id the document does not define is `not_found`, the same
        # code a locator gives for a place that is not there.
        raise LocatorError("not_found", str(exc)) from exc
    return {
        "style": {
            "info": dataclasses.asdict(detail.info),
            "run_props": detail.run_props,
            "paragraph_props": detail.paragraph_props,
            "chain": [dataclasses.asdict(level) for level in detail.chain],
            "resolved": detail.resolved,
        },
        "theme": _theme_report(read_theme(pkg)),
        "warnings": list(detail.warnings),
    }


def doc_find_style_usage(
    filename: str, style_id: str, max_results: int | None = None
) -> dict[str, Any]:
    """Find every paragraph, run and table that applies a style.

    This is what to call before changing or removing a style: it says exactly
    what would move, and each usage carries the locator that reaches it.

    The match is on the style id exactly, and inheritance is not followed: a
    paragraph styled with something *based on* `style_id` is not reported, since
    it is a usage of its own style. Every story is searched -- body, headers,
    footers, footnotes, endnotes.

    Args:
        filename: path to the .docx file.
        style_id: the style id to look for, as `doc_list_styles` reports it. A
            style *name* finds nothing.
        max_results: stop after this many usages.

    Returns:
        The usual report, plus `usages` and `truncated`. Each usage has `kind`
        (`paragraph`, `run` or `table`), `story`, `locator`, `index` (the V2
        paragraph index, null in a table cell or a text box), `table`, `start`
        and `end` (a run's offsets in its paragraph's visible text, to paste
        into `doc_format_range`), and a `text` preview.

        `locator` is ready for any tool that takes one, with two exceptions: for
        a `table` it is `{"story", "table"}`, the table's address, and reaching
        a cell means adding `row` and `col`; and it is null for a paragraph in a
        text box, which has no locator of its own -- use a `find` locator there.
    """
    if max_results is not None and max_results < 0:
        raise ValueError(f"'max_results' must be 0 or more, got {max_results}")
    usages = find_style_usage(DocxPackage.open(filename), style_id)
    truncated = max_results is not None and len(usages) > max_results
    if truncated:
        usages = usages[:max_results]
    return {
        "usages": [dataclasses.asdict(usage) for usage in usages],
        "truncated": truncated,
        "warnings": (
            [f"stopped at 'max_results' ({max_results}); more usages exist"]
            if truncated
            else []
        ),
    }


TOOLS = [
    ToolSpec(
        fn=doc_list_styles,
        annotations=ToolAnnotations(title="List Styles", readOnlyHint=True),
        tags=frozenset({"v2", "read", "styles"}),
    ),
    ToolSpec(
        fn=doc_get_style,
        annotations=ToolAnnotations(title="Get Style", readOnlyHint=True),
        tags=frozenset({"v2", "read", "styles"}),
    ),
    ToolSpec(
        fn=doc_find_style_usage,
        annotations=ToolAnnotations(title="Find Style Usage", readOnlyHint=True),
        tags=frozenset({"v2", "read", "styles"}),
    ),
]
