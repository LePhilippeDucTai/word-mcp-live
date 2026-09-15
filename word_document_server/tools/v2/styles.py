"""The three V2 style tools: list the style sheet, read one style, find its uses.

They answer the half of a document's appearance that ``doc_inspect`` and
``doc_format_range`` cannot: what the styles *say*.  An agent asked to make a
document consistent needs to know that ``FixtureLeaf`` is based on
``FixtureBranch``, that the bold comes from there and not from the paragraph,
and which eleven paragraphs would change if the style did.  That is
:func:`doc_list_styles`, :func:`doc_get_style` and :func:`doc_find_style_usage`.

Three of them are read-only.  The three that write -- :func:`doc_create_style`,
:func:`doc_update_style`, :func:`doc_delete_style` -- work on the very same
vocabulary: what :func:`doc_get_style` reports under `run_props` and
`paragraph_props` is what they take, so "make this style look like that one" is
a read followed by a write and not a translation exercise.

A style edit is the widest edit this server offers: it changes every paragraph
that names the style, at once, and the report says how many that is.  Deleting
one is wider still, which is why :func:`doc_delete_style` refuses a style that
is in use until it is told what those places should name instead.

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
    STYLE_SPEC_KEYS,
    create_style,
    delete_style,
    find_style_usage,
    get_style,
    list_styles,
    update_style,
)
from word_document_server.engine.theme import Theme, read_theme
from word_document_server.tools.v2.registry import ToolSpec

__all__ = [
    "TOOLS",
    "doc_create_style",
    "doc_delete_style",
    "doc_find_style_usage",
    "doc_get_style",
    "doc_list_styles",
    "doc_update_style",
]

#: What every write tool here documents about the spec it takes.  Written once
#: because the three descriptions an agent reads must not drift apart.
_SPEC_DOC = f"keys: {', '.join(STYLE_SPEC_KEYS)}"


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


def _spec_keys(spec: Any) -> dict[str, Any]:
    """Check a spec is a JSON object before the engine sees it.

    The engine rejects an unknown key too; doing it here as well is what turns
    ``{"bold": true}`` -- a plausible mistake, since `run_props` does take
    `bold` -- into `invalid_argument` with the list of keys, rather than into a
    style created with nothing in it.
    """
    if not isinstance(spec, dict):
        raise ValueError(f"'spec' must be an object of {_SPEC_DOC}")
    unknown = sorted(set(spec) - set(STYLE_SPEC_KEYS))
    if unknown:
        raise ValueError(
            f"unknown spec keys: {', '.join(unknown)}; {_SPEC_DOC}"
        )
    return spec


def doc_create_style(
    filename: str, spec: dict[str, Any], dry_run: bool = False
) -> dict[str, Any]:
    """Define a new paragraph or character style in a Word document.

    Use this instead of formatting paragraphs one by one: a style is the only
    way a document stays consistent when it is edited later, and the only way
    one change reaches every place at once. `doc_list_styles` first -- a
    document usually already has a style for what you are about to define.

    Args:
        filename: path to the .docx file.
        spec: what the style is.
            `name` (required): what Word's style pane shows.
            `style_id`: the id `w:pStyle` will carry, derived from the name (its
                letters and digits, `"Fixture Body"` -> `"FixtureBody"`) when
                omitted.
            `family`: `paragraph` (default) or `character`.
            `based_on`, `next`, `link`: other styles, by id or by name; each
                must already exist, and is stored as the id. `based_on` is what
                a style inherits from, `next` the style Word gives the paragraph
                typed after one in this style, `link` the character style a
                paragraph style pairs with.
            `q_format`: offer the style in Word's gallery.
            `ui_priority`: where it sorts there.
            `builtin`: declare it as one of Word's own styles rather than a
                custom one. Leave it out unless you are recreating a style Word
                defines itself.
            `run_props`, `paragraph_props`: the properties, in the vocabulary
                `doc_get_style` reports and `doc_format_range` takes -- `bold`,
                `size_pt`, `font`, `color`, `alignment`, `indent`, `spacing`,
                `outline_level`. Measurements are in twips (1/1440 inch). A
                property may be given as `{"value": ..., "theme": ...}` to
                follow the document theme, which is what `doc_get_style`
                reports for one that does. A character style has no
                `paragraph_props`.
        dry_run: compute and report the change without writing the file.

    Returns:
        The usual report, plus `style` -- the new style's identity and wiring,
        as `doc_list_styles` reports it -- and `saved`.
    """
    spec = _spec_keys(spec)
    pkg = DocxPackage.open(filename)
    info = create_style(pkg, spec)
    saved = not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "style": dataclasses.asdict(info),
    }


def doc_update_style(
    filename: str, style: str, spec: dict[str, Any], dry_run: bool = False
) -> dict[str, Any]:
    """Change a style, and with it every place that uses it.

    This is the widest edit in this server: the change reaches every paragraph,
    run and table naming the style, at once. Call `doc_find_style_usage` first
    to see what that is; the report here says how many places it was.

    The spec is a set of *changes*: a key left out is left as it is, and a key
    set to null removes what the style said, so what it inherits from
    `based_on` applies again. `style_id` and `family` cannot be changed --
    every place that uses the style names the first and relies on the second,
    so a different one is a different style: clone it with a new name instead.

    Args:
        filename: path to the .docx file.
        style: the style to change, by id (`Heading1`) or by name
            (`heading 1`). The id is tried first.
        spec: the changes, in the vocabulary `doc_create_style` documents.
        dry_run: compute and report the change without writing the file.

    Returns:
        The usual report, plus `style` (the style as it now stands), `usages`
        (how many places it affects) and `saved`.
    """
    spec = _spec_keys(spec)
    pkg = DocxPackage.open(filename)
    info = update_style(pkg, style, spec)
    usages = find_style_usage(pkg, info.style_id)
    saved = not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "style": dataclasses.asdict(info),
        "usages": len(usages),
        "warnings": (
            [
                (
                    f"{len(usages)} place(s) use {info.style_id!r} and are affected "
                    "by this change"
                )
            ]
            if usages
            else []
        ),
    }


def doc_delete_style(
    filename: str,
    style: str,
    reassign_to: str | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Remove a style, reassigning everything that used it.

    A style nothing uses is removed outright. A style that *is* used is refused
    -- code `style_in_use` -- unless `reassign_to` says which style those
    places should name instead: deleting it silently would leave Word to fall
    back to `Normal` for every one of them, without saying so.

    The other styles are dealt with too: a `based_on`, `next` or `link` naming
    the deleted style follows the reassignment, or is dropped when there is
    none, so no style is left inheriting from something that no longer exists.

    Args:
        filename: path to the .docx file.
        style: the style to delete, by id or by name.
        reassign_to: the style the places that used it should name instead, by
            id or by name. It must be of the same family: a paragraph cannot
            name a character style.
        dry_run: compute and report the change without writing the file.

    Returns:
        The usual report, plus `style_id`, `reassigned_to`, `references` (the
        ids of the styles whose wiring was repointed) and `saved`. `changes`
        holds one entry per place that was reassigned; a style is not paragraph
        text, so `before` and `after` are equal -- they say which places moved.
    """
    pkg = DocxPackage.open(filename)
    deletion = delete_style(pkg, style, reassign_to)
    saved = not dry_run
    if saved:
        pkg.save(filename)
    warnings = []
    if deletion.references:
        warnings.append(
            "the basedOn/next/link of "
            f"{', '.join(deletion.references)} named {deletion.style_id!r} and "
            + (
                f"now name {deletion.reassigned_to!r}"
                if deletion.reassigned_to
                else "was removed"
            )
        )
    return {
        "dry_run": dry_run,
        "saved": saved,
        "style_id": deletion.style_id,
        "reassigned_to": deletion.reassigned_to,
        "references": list(deletion.references),
        "changes": [
            {
                "story": usage.story,
                "paragraph": usage.index,
                "before": usage.text,
                "after": usage.text,
            }
            for usage in deletion.usages
        ],
        "warnings": warnings,
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
    ToolSpec(
        fn=doc_create_style,
        annotations=ToolAnnotations(title="Create Style", destructiveHint=False),
        tags=frozenset({"v2", "write", "styles"}),
    ),
    ToolSpec(
        fn=doc_update_style,
        annotations=ToolAnnotations(title="Update Style", destructiveHint=True),
        tags=frozenset({"v2", "write", "styles"}),
    ),
    ToolSpec(
        fn=doc_delete_style,
        annotations=ToolAnnotations(title="Delete Style", destructiveHint=True),
        tags=frozenset({"v2", "write", "styles"}),
    ),
]
