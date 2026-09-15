"""The four V2 text tools: look, search, edit, format.

They are the loop an agent runs on a document.  :func:`doc_inspect` says what is
in it and how to name the places in it; :func:`doc_find` says where a string is;
:func:`doc_edit_text` and :func:`doc_format_range` act on a place named by a
*locator* -- a plain dict, described in
:mod:`word_document_server.engine.locators`, that survives being written down
and sent back later, and that fails loudly rather than quietly hitting the wrong
paragraph when the document moved underneath it.

Offsets
-------
A locator resolves to a paragraph *and* a span inside it: the whole visible text
for ``{"paragraph": 3}``, the matched text for ``{"find": "Total"}``, the
bookmarked text for ``{"bookmark": "Summary"}``.  The `start` and `end`
arguments of the two editing tools override that span.  They are offsets in the
**visible text of the located paragraph**, counted from 0, exactly the numbers
:func:`doc_find` reports -- so a match can be edited by pasting its `start` and
`end` back, and a whole paragraph by passing neither.

Visible text is what a reader sees: text hidden under a tracked deletion is not
in it, text inside a hyperlink or a tracked insertion is.  It never spans two
paragraphs, so neither does any of these operations.

Reports
-------
Every tool answers the shape described in
:mod:`word_document_server.tools.v2.registry`.  Nothing is written when
``dry_run`` is true, and nothing is written when the call raises: the package is
edited in memory and saved -- atomically, once -- only at the end.
"""

from __future__ import annotations

from typing import Any, Literal

from mcp.types import ToolAnnotations

from word_document_server.defaults import DEFAULT_AUTHOR
from word_document_server.engine import format as engine_format
from word_document_server.engine import ranges, revisions
from word_document_server.engine.find import find as engine_find
from word_document_server.engine.inspect import inspect as engine_inspect
from word_document_server.engine.locators import Target, resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.tools.v2.registry import ToolSpec

__all__ = [
    "TOOLS",
    "doc_edit_text",
    "doc_find",
    "doc_format_range",
    "doc_inspect",
]

#: The default `stories` filter of :func:`doc_find`.  ``"body"`` is the friendly
#: alias of the main document story on the way in; it is never reported back,
#: which is D-006.
DEFAULT_STORIES = ("body",)

#: What :func:`doc_edit_text` knows how to do.
Action = Literal["replace", "insert", "delete"]


def _span(target: Target, start: int | None, end: int | None) -> tuple[int, int]:
    """The offsets to act on: the locator's span, unless the caller gave others."""
    return (
        target.start if start is None else start,
        target.end if end is None else end,
    )


def _change(target: Target, before: str) -> dict[str, Any]:
    """One ``changes`` entry for the paragraph `target` names.

    `paragraph` is the V2 index, ``None`` in a table cell or a text box (D-016);
    `after` is re-read from the live paragraph, so it reports what happened
    rather than what was asked for.
    """
    return {
        "story": target.story,
        "paragraph": target.index,
        "before": before,
        "after": visible_text(target.paragraph),
    }


def doc_inspect(filename: str) -> dict[str, Any]:
    """Describe a Word document and every place a locator can name in it.

    This is what to read before writing a locator: each block carries the
    paragraph index ``{"paragraph": i}`` takes, each table the
    ``{"table": t, "row": r, "col": c}`` coordinates, each bookmark its name,
    each heading its text.

    Args:
        filename: path to the .docx file.

    Returns:
        The usual report, plus `document`: `stories` (id and paragraph count),
        `counts`, `blocks` (index, kind, style, text preview, heading and list
        level, presence of fields, comments and revisions), `tables`,
        `sections`, `styles_used`, `bookmarks` and `fields`.  Paragraph text is
        truncated to a preview; use `doc_find` to locate exact strings.
    """
    return {"document": engine_inspect(DocxPackage.open(filename))}


def doc_find(
    filename: str,
    pattern: str,
    regex: bool = False,
    case: bool = True,
    whole_word: bool = False,
    stories: list[str] | None = None,
    max_results: int | None = None,
) -> dict[str, Any]:
    """Find where a string occurs in a Word document, paragraph by paragraph.

    The search runs on the visible text, so it neither turns up text hidden
    under a tracked deletion nor misses text that only exists inside a
    hyperlink, an insertion or a content control.  A match never spans two
    paragraphs.

    Args:
        filename: path to the .docx file.
        pattern: the text to look for, or a Python regular expression when
            `regex` is true.  Must not be empty, and must not be able to match
            zero characters.
        regex: treat `pattern` as a regular expression.
        case: match case-sensitively (the default).
        whole_word: only match on word boundaries.
        stories: which stories to search, defaulting to the body.  A story id is
            what `doc_inspect` reports (`document`, `header1`, `footer2`,
            `footnotes`, `endnotes`, ...); `body` is accepted as an alias of
            `document`.  A story the document does not have is skipped, not an
            error.
        max_results: stop after this many matches.

    Returns:
        The usual report, plus `matches` -- `story` (the real id, never `body`),
        `paragraph` (the index a locator takes, `null` in a table cell or a text
        box), `start`, `end`, `text` and `context` -- and `truncated`, true when
        `max_results` cut the search short.
    """
    if max_results is not None and max_results < 0:
        raise ValueError(f"'max_results' must be 0 or more, got {max_results}")
    wanted = list(DEFAULT_STORIES) if stories is None else list(stories)
    # One more than asked for, to tell "that is all of them" from "that is all
    # you asked for" without scanning the rest of the document.
    probe = None if max_results is None else max_results + 1
    matches = engine_find(
        DocxPackage.open(filename),
        pattern,
        regex=regex,
        case=case,
        whole_word=whole_word,
        stories=wanted,
        max_results=probe,
    )
    truncated = max_results is not None and len(matches) > max_results
    if truncated:
        matches = matches[:max_results]
    return {
        "matches": [
            {
                "story": match.story,
                "paragraph": match.index,
                "start": match.start,
                "end": match.end,
                "text": match.text,
                "context": match.context,
            }
            for match in matches
        ],
        "truncated": truncated,
        "warnings": (
            [f"stopped at 'max_results' ({max_results}); more matches exist"]
            if truncated
            else []
        ),
    }


def doc_edit_text(
    filename: str,
    locator: dict[str, Any],
    action: Action = "replace",
    text: str = "",
    start: int | None = None,
    end: int | None = None,
    track_changes: bool = False,
    author: str = DEFAULT_AUTHOR,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Replace, insert or delete text at a place named by a locator.

    Styles, fields, bookmarks, comments, hyperlinks and existing tracked changes
    around the edit are preserved; a range that would cut through a field or
    swallow an image is refused rather than approximated.

    Args:
        filename: path to the .docx file.
        locator: where to act. One of `{"paragraph": i}`, `{"find": s,
            "occurrence": n}`, `{"bookmark": name}`, `{"heading": s}`,
            `{"table": t, "row": r, "col": c, "paragraph": k}`, with the
            optional keys `story` and `expect_text`. `expect_text` is the text
            the caller believes is there: a mismatch fails as `stale_anchor`
            instead of editing the wrong paragraph.
        action: `replace` the span, `insert` at its start, or `delete` it.
        text: the text to write, for `replace` and `insert`. `\\n` becomes a line
            break and `\\t` a tab. Empty on a `replace` makes it a deletion.
        start: offset to act from, in the visible text of the located
            paragraph. Defaults to the start of the span the locator resolved
            to. For `insert`, this is the single offset the text lands at.
        end: offset to act up to, excluded. Defaults to the end of that span,
            and is not used by `insert`.
        track_changes: record the edit as a Word revision -- the old text is
            hidden rather than dropped, so it can be accepted or rejected in
            Word -- instead of applying it outright.
        author: the name recorded on the revision, when `track_changes` is true.
        dry_run: compute and report the edit without writing the file.

    Returns:
        The usual report. `changes` holds one entry for the paragraph touched,
        with its visible text `before` and `after`. `saved` says whether the
        file was written.
    """
    if action == "delete" and text:
        raise ValueError("'delete' removes the range and takes no 'text'")
    if action == "insert" and end is not None:
        raise ValueError("'insert' writes at a single offset; give 'start', not 'end'")

    pkg = DocxPackage.open(filename)
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

    saved = mutated and not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "changes": [_change(target, before)],
        "warnings": warnings,
    }


def doc_format_range(
    filename: str,
    locator: dict[str, Any],
    patch: dict[str, Any],
    start: int | None = None,
    end: int | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Change the character formatting of a range named by a locator.

    The patch is a set of *changes*, not a description of the final formatting:
    a property it does not mention is left alone, and a property set to `null`
    is removed, so what the style says applies again.

    Args:
        filename: path to the .docx file.
        locator: where to act; the same forms as `doc_edit_text`.
        patch: the properties to change. `bold`, `italic`, `caps`, `small_caps`,
            `strike`, `superscript`, `subscript`: true, false or null.
            `underline`: true, false, a style name (`single`, `double`,
            `wave`, ...) or null. `size_pt`: points, in halves. `font`: a font
            name. `color`: `auto` or `RRGGBB`. `highlight`: a Word highlight
            colour name. `char_style`: the id of a character style the document
            defines.
        start: offset to format from, in the visible text of the located
            paragraph. Defaults to the start of the span the locator resolved to.
        end: offset to format up to, excluded. Defaults to the end of that span.
        dry_run: compute and report the change without writing the file.

    Returns:
        The usual report, plus `runs`, the number of runs patched. Formatting
        never changes the visible text, so `before` and `after` in `changes` are
        equal; they are there to confirm which paragraph was reached.
    """
    if not patch:
        raise ValueError("'patch' must name at least one property to change")

    pkg = DocxPackage.open(filename)
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
    saved = bool(patched) and not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "runs": len(patched),
        "changes": [_change(target, before)],
        "warnings": warnings,
    }


TOOLS = [
    ToolSpec(
        fn=doc_inspect,
        annotations=ToolAnnotations(title="Inspect Document", readOnlyHint=True),
        tags=frozenset({"v2", "read"}),
    ),
    ToolSpec(
        fn=doc_find,
        annotations=ToolAnnotations(title="Find Text", readOnlyHint=True),
        tags=frozenset({"v2", "read"}),
    ),
    ToolSpec(
        fn=doc_edit_text,
        annotations=ToolAnnotations(title="Edit Text", destructiveHint=True),
        tags=frozenset({"v2", "write"}),
    ),
    ToolSpec(
        fn=doc_format_range,
        annotations=ToolAnnotations(title="Format Range", destructiveHint=True),
        tags=frozenset({"v2", "write"}),
    ),
]
