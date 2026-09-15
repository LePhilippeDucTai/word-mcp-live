"""The one V2 batch tool: :func:`doc_apply_edits`.

An agent that has planned several edits should not have to save between them --
doing so exposes a half-edited document to autosave, and re-opens the package
once per edit for no reason.  :func:`doc_apply_edits` runs a whole plan on a
single :class:`~word_document_server.engine.package.DocxPackage`, in memory,
and writes it once, atomically, only if every edit in the plan succeeded.

Each entry of `edits` is a payload shaped like the arguments of
:func:`~word_document_server.tools.v2.text.doc_edit_text` or
:func:`~word_document_server.tools.v2.text.doc_format_range` (everything but
`filename` and its own `dry_run`, which this tool takes once for the whole
call): an entry carrying `patch` is a formatting edit, resolved and applied the
way `doc_format_range` does; every other entry is a text edit, resolved and
applied the way `doc_edit_text` does, `action` defaulting to `"replace"` just
as it does there.

Locators are resolved one edit at a time, against the document as the edits
before it left it -- never against a snapshot taken before the call -- so an
edit can target a paragraph a previous edit in the same call just wrote, and a
locator that only makes sense after an earlier edit works exactly as it would
across two separate calls.

All or nothing
---------------
The edits run in order.  The first one that fails stops the call: nothing
written by the edits before it survives, because none of them have been saved
yet -- the package is still only in memory -- and the failing edit's own report
is raised with the index of the edit prepended to its message, its `code`
(D-022) kept exactly as :mod:`~word_document_server.tools.v2.registry` would
report it standalone. A locator failure, a range that cannot be applied, bad
text: all of it leaves the file on disk untouched.

Key validation (D-024)
-----------------------
A batch edit is a plain `dict`: unlike a top-level call to `doc_edit_text` or
`doc_format_range`, it is never checked against that function's signature by
an MCP schema, so a misspelled key (`content` instead of `text`) would
otherwise be read by `.get()` as simply absent and silently applied as that
key's default -- for a `replace`, an empty `text`, erasing the targeted range
rather than refusing the call. Each entry is checked against the exact key
set of the tool it is shaped like (:data:`_TEXT_EDIT_KEYS`,
:data:`_FORMAT_EDIT_KEYS`) before any locator is resolved: an unknown key or a
missing `locator` raises before anything is read or written.
"""

from __future__ import annotations

from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.defaults import DEFAULT_AUTHOR
from word_document_server.engine import format as engine_format
from word_document_server.engine import ranges, revisions
from word_document_server.engine.errors import (
    EngineError,
    LocatorError,
    UnsupportedRange,
)
from word_document_server.engine.locators import resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.tools.v2.registry import ToolSpec
from word_document_server.tools.v2.text import _change, _span

__all__ = ["TOOLS", "doc_apply_edits"]

#: Exceptions one edit of the batch can fail with -- the same family
#: :mod:`~word_document_server.tools.v2.registry` reports for a standalone call.
_EDIT_FAILURES: tuple[type[Exception], ...] = (EngineError, OSError, ValueError, TypeError)

#: `action` values :func:`_apply_text_edit` knows, mirroring `doc_edit_text`'s
#: `Literal`. Checked explicitly here because a batch edit is a plain `dict`,
#: never validated against that `Literal` by an MCP schema the way a top-level
#: call to `doc_edit_text` is.
_ACTIONS = frozenset({"replace", "insert", "delete"})

#: The keys a text-edit entry may carry: `doc_edit_text`'s own parameters,
#: minus `filename` and its own `dry_run`, which this tool takes once for the
#: whole call. A batch edit is a plain `dict`, never validated against that
#: signature by an MCP schema the way a top-level call to `doc_edit_text` is
#: -- so an unrecognized key (a typo, a stray `content` instead of `text`)
#: must be caught here, before it is silently read as a default by `.get()`.
_TEXT_EDIT_KEYS = frozenset(
    {"locator", "action", "text", "start", "end", "track_changes", "author"}
)

#: The keys a format-edit entry may carry: `doc_format_range`'s own
#: parameters, minus `filename` and its own `dry_run`.
_FORMAT_EDIT_KEYS = frozenset({"locator", "patch", "start", "end"})


def _check_keys(edit: dict[str, Any], allowed: frozenset[str]) -> None:
    """Refuse an edit carrying a key outside `allowed`, or missing `locator`.

    Called before either helper reads the payload with `.get()`, which would
    otherwise treat an unknown key (a typo, `content` instead of `text`) as
    simply absent and silently apply its default -- see D-024.
    """
    unknown = set(edit) - allowed
    if unknown:
        raise ValueError(
            f"unknown key(s) {sorted(unknown)}; this edit accepts {sorted(allowed)}"
        )
    if "locator" not in edit:
        raise ValueError(f"'locator' is required; this edit accepts {sorted(allowed)}")


def _named(exc: Exception, index: int) -> Exception:
    """Return `exc` with its message naming the failing edit, its code kept.

    `error_code` (in :mod:`~word_document_server.tools.v2.registry`) reads
    `LocatorError.code` and `UnsupportedRange.reason` off the exception itself,
    and derives every other code from the exception's class, so both are
    preserved by re-raising the same class rather than a generic one.
    """
    message = f"edit {index}: {exc}"
    if isinstance(exc, LocatorError):
        return LocatorError(exc.code, message)
    if isinstance(exc, UnsupportedRange):
        return UnsupportedRange(exc.reason, message)
    return type(exc)(message)


def _apply_text_edit(pkg: DocxPackage, edit: dict[str, Any]) -> dict[str, Any]:
    """Apply one `doc_edit_text`-shaped entry of `edits`, in memory.

    Same rules as :func:`~word_document_server.tools.v2.text.doc_edit_text`,
    minus the parts that only make sense for a standalone call (`filename`,
    opening the package, saving).
    """
    _check_keys(edit, _TEXT_EDIT_KEYS)
    locator = edit.get("locator")
    action = edit.get("action", "replace")
    text = edit.get("text", "")
    start = edit.get("start")
    end = edit.get("end")
    track_changes = bool(edit.get("track_changes", False))
    author = edit.get("author", DEFAULT_AUTHOR)

    if action not in _ACTIONS:
        raise ValueError(
            f"'action' must be one of {sorted(_ACTIONS)}, got {action!r}"
        )
    if action == "delete" and text:
        raise ValueError("'delete' removes the range and takes no 'text'")
    if action == "insert" and end is not None:
        raise ValueError("'insert' writes at a single offset; give 'start', not 'end'")

    target = resolve(pkg, locator)
    span_start, span_end = _span(target, start, end)
    paragraph = target.paragraph
    before = visible_text(paragraph)
    warnings: list[str] = []
    mutated = True

    if action == "replace":
        if track_changes:
            revisions.tracked_replace(pkg, paragraph, span_start, span_end, text, author)
        else:
            ranges.replace_range(paragraph, span_start, span_end, text)
    elif action == "insert":
        if not text:
            mutated = False
            warnings.append("'text' is empty: nothing was inserted")
        elif track_changes:
            if not revisions.tracked_insert(pkg, paragraph, span_start, text, author):
                warnings.append(
                    f"the text was added to an insertion {author!r} already owns; no new "
                    "revision was created"
                )
        else:
            ranges.insert_text(paragraph, span_start, text)
    elif track_changes:
        pieces = ranges.resolve(paragraph, span_start, span_end, operation="delete")
        if not revisions.tracked_delete(pkg, pieces, author):
            warnings.append(
                f"the range was an insertion {author!r} had not yet had accepted; it was "
                "removed outright rather than marked as deleted"
            )
    else:
        ranges.delete_range(paragraph, span_start, span_end)

    return {"mutated": mutated, "change": _change(target, before), "warnings": warnings}


def _apply_format_edit(pkg: DocxPackage, edit: dict[str, Any]) -> dict[str, Any]:
    """Apply one `doc_format_range`-shaped entry of `edits`, in memory.

    Same rules as
    :func:`~word_document_server.tools.v2.text.doc_format_range`, minus the
    parts that only make sense for a standalone call.
    """
    _check_keys(edit, _FORMAT_EDIT_KEYS)
    locator = edit.get("locator")
    patch = edit.get("patch")
    start = edit.get("start")
    end = edit.get("end")

    if not patch:
        raise ValueError("'patch' must name at least one property to change")

    target = resolve(pkg, locator)
    span_start, span_end = _span(target, start, end)
    before = visible_text(target.paragraph)

    pieces = ranges.resolve(target.paragraph, span_start, span_end, operation="read")
    patched = engine_format.apply_rpr(pieces, patch, pkg=pkg)

    warnings: list[str] = []
    if not patched:
        warnings.append(
            f"range {span_start}..{span_end} covers no run; nothing was formatted"
        )
    return {"mutated": bool(patched), "change": _change(target, before), "warnings": warnings}


def doc_apply_edits(
    filename: str,
    edits: list[dict[str, Any]],
    dry_run: bool = False,
) -> dict[str, Any]:
    """Apply a sequence of text and formatting edits as one all-or-nothing unit.

    Each entry of `edits` is a payload shaped like the arguments of
    `doc_edit_text` or `doc_format_range`, minus `filename` and its own
    `dry_run`: an entry carrying `patch` is applied as `doc_format_range`
    would apply it; every other entry is applied as `doc_edit_text` would,
    `action` defaulting to `"replace"`. Locators are resolved one edit at a
    time, against the document as the edits before it in the same call left
    it, so a later entry can target text an earlier one just wrote.

    If any edit fails, none of the call's edits are written: the file is left
    exactly as it was, and the error names the index of the edit that failed.
    Otherwise the whole result is written in a single atomic save, unless
    `dry_run` is true.

    Args:
        filename: path to the .docx file.
        edits: the edits to apply, in order. Must not be empty; each entry
            must be a mapping.
        dry_run: compute and report every edit without writing the file.

    Returns:
        The usual report. `changes` holds one entry per edit that changed
        something, in the same shape `doc_edit_text` and `doc_format_range`
        use. `saved` says whether the file was written.
    """
    if not edits:
        raise ValueError("'edits' must hold at least one edit")

    pkg = DocxPackage.open(filename)
    changes: list[dict[str, Any]] = []
    warnings: list[str] = []
    mutated = False

    for index, edit in enumerate(edits):
        if not isinstance(edit, dict):
            raise TypeError(f"edit {index} must be a mapping, got {type(edit).__name__}")
        try:
            result = (
                _apply_format_edit(pkg, edit)
                if "patch" in edit
                else _apply_text_edit(pkg, edit)
            )
        except _EDIT_FAILURES as exc:
            raise _named(exc, index) from exc

        if result["mutated"]:
            mutated = True
            changes.append(result["change"])
        warnings.extend(f"edit {index}: {warning}" for warning in result["warnings"])

    saved = mutated and not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "changes": changes,
        "warnings": warnings,
    }


TOOLS = [
    ToolSpec(
        fn=doc_apply_edits,
        annotations=ToolAnnotations(title="Apply Edits", destructiveHint=True),
        tags=frozenset({"v2", "write"}),
    ),
]
