"""Structural comparison between two Word documents.

:func:`doc_compare` is the semantic layer over
:mod:`word_document_server.engine.compare`, the package's canonical
byte-for-byte comparison instrument (moved there in J06-P2 from the test
fidelity harness it started as).  The engine's own
:func:`~word_document_server.engine.compare.diff` answers in the *snapshot
space*: every ``w:p`` of a story, table cells and text boxes included, so that
no loss goes unseen (see that module's docstring).  An agent reading a
locator, though, thinks in the *V2* space
:func:`~word_document_server.engine.locators.resolve` understands, where a
paragraph inside a table cell or a text box has no such index (``None``,
D-016).  This module is where that translation happens; the engine module
itself is left untouched, since hundreds of tests pin its exact behaviour.

What is reported
-----------------
``parts``, ``content_types`` and ``relationships`` are reported close to the
engine's own vocabulary: they are package plumbing, not something a V2
locator names.

``paragraphs`` translates each snapshot-space key ``(story, index)`` into
``{"story", "paragraph"}`` -- the same two keys a ``changes`` entry of any
other V2 tool carries -- before reporting it.  A paragraph that was added is
located in `filename_b` (it does not exist in `filename_a`); one that was
removed is located in `filename_a`; one that changed keeps the same
snapshot-space position in both documents, so either side's translation names
the same place, and `filename_b`'s is the one reported.  Each entry also
carries the paragraph's visible text (`text` for an addition or a removal,
`text_before`/`text_after` for a change), and a changed entry's `fields` names
which parts of the paragraph's signature differ -- `text`, `style`, `runs`
(run-level formatting: bold, italic, images, breaks, ...), `markers`
(bookmarks and comment anchors gained or lost), `ppr` (paragraph-level
formatting), and the others
:mod:`~word_document_server.engine.compare` tracks.

``tables`` reports each table's own snapshot-space `index`, *not* translated:
it counts every ``w:tbl`` of a story, including nested ones, which is a wider
space than the ``table`` locator (main-story, top-level tables only)
resolves -- silently turning one into the other would misname a table a
caller then edits.

``counters_changed`` is the global delta of `comments`, `comment_references`,
`revisions`, `bookmarks`, `fields`, `hyperlinks` and the footnote/endnote
totals -- the same names :func:`~word_document_server.engine.compare.diff`
reports, unfiltered.  It is deliberately a *count*, not a list: the detail
behind a changed count -- which comment, which revision -- is what
`doc_inspect` and `doc_find` on the two files are for.
"""

from __future__ import annotations

from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine.compare import (
    ParagraphChange,
    Snapshot,
    TableChange,
    diff as engine_diff,
    snapshot as engine_snapshot,
)
from word_document_server.engine.find import _v2_index_map, iter_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.tools.v2.registry import ToolSpec

__all__ = ["TOOLS", "doc_compare"]


def _v2_positions(pkg: DocxPackage) -> dict[str, list[int | None]]:
    """Per story, the V2 index of the paragraph at each snapshot-space position.

    Position ``i`` of the returned list for a story is the V2 index (or
    ``None`` inside a table cell, a text box or an ``mc:AlternateContent``
    branch, per :func:`~word_document_server.engine.find._v2_index_map`) of
    the ``i``-th ``w:p`` of that story in document order -- the same walk
    :func:`~word_document_server.engine.compare.snapshot` does to number the
    snapshot space, so position and index line up without needing the live
    elements a second time.  A story `pkg` has no live view of (``comments``,
    which :meth:`~word_document_server.engine.package.DocxPackage.stories`
    never reports, unlike the engine's snapshot) is simply absent here, and
    :func:`_locator` reports ``None`` for it.
    """
    positions: dict[str, list[int | None]] = {}
    for story_id, root in pkg.stories():
        paragraphs = iter_paragraphs(root)
        index_map = _v2_index_map(paragraphs)
        positions[story_id] = [index_map.get(paragraph) for paragraph in paragraphs]
    return positions


def _locator(
    positions: dict[str, list[int | None]], story: str, position: int
) -> dict[str, Any]:
    """The ``{"story", "paragraph"}`` locator for a snapshot-space key."""
    story_positions = positions.get(story, [])
    v2_index = story_positions[position] if position < len(story_positions) else None
    return {"story": story, "paragraph": v2_index}


def _paragraph_entry(
    positions: dict[str, list[int | None]], key: tuple[str, int], text: str
) -> dict[str, Any]:
    story, position = key
    entry = _locator(positions, story, position)
    entry["text"] = text
    return entry


def _paragraph_change(
    positions: dict[str, list[int | None]],
    before: Snapshot,
    after: Snapshot,
    change: ParagraphChange,
) -> dict[str, Any]:
    story, position = change.key
    entry = _locator(positions, story, position)
    entry["fields"] = [item.field for item in change.changes]
    entry["text_before"] = before.paragraphs[change.key].text
    entry["text_after"] = after.paragraphs[change.key].text
    return entry


def _table_entry(key: tuple[str, int]) -> dict[str, Any]:
    story, index = key
    return {"story": story, "table": index}


def _table_change(change: TableChange) -> dict[str, Any]:
    story, index = change.key
    return {
        "story": story,
        "table": index,
        "fields": [item.field for item in change.changes],
    }


def doc_compare(filename_a: str, filename_b: str) -> dict[str, Any]:
    """Compare two Word documents structurally and report what changed.

    The comparison is positional, not a text diff: a paragraph inserted before
    others shifts every later one and is reported as that many paragraphs
    "changed" rather than one addition -- the same behaviour
    :func:`~word_document_server.engine.compare.diff` has always had.

    Args:
        filename_a: path to the first .docx file (the "before" side).
        filename_b: path to the second .docx file (the "after" side).

    Returns:
        The usual report, plus:

        `identical`: true when nothing differs.

        `parts`, `content_types`, `relationships`: package-level changes --
        `added`/`removed`/`changed` part names, `{"part", "content_type"}`
        entries, `{"part", "type", "target"}` relationship entries.

        `paragraphs`: `added`, `removed` (each `{"story", "paragraph", "text"}`)
        and `changed` (`{"story", "paragraph", "fields", "text_before",
        "text_after"}`), `paragraph` being the V2 locator index or `null` in a
        table cell or a text box. See the module docstring for `fields`.

        `tables`: `added`, `removed` (`{"story", "table"}`) and `changed`
        (adds `fields`); `table` is the engine's own table index, not a V2
        locator (see the module docstring).

        `sections_changed`: story ids whose body `w:sectPr` differs (page
        size, margins, header/footer references).

        `counters_changed`: `{"field", "before", "after"}` for every package
        total that differs -- comments, comment references, revisions,
        bookmarks, fields, hyperlinks, footnotes and endnotes.
    """
    pkg_a = DocxPackage.open(filename_a)
    pkg_b = DocxPackage.open(filename_b)

    before = engine_snapshot(filename_a)
    after = engine_snapshot(filename_b)
    delta = engine_diff(before, after)

    positions_a = _v2_positions(pkg_a)
    positions_b = _v2_positions(pkg_b)

    return {
        "identical": delta.is_empty(),
        "parts": {
            "added": list(delta.parts_added),
            "removed": list(delta.parts_removed),
            "changed": list(delta.parts_changed),
        },
        "content_types": {
            "added": [
                {"part": part, "content_type": ctype}
                for part, ctype in delta.content_types_added
            ],
            "removed": [
                {"part": part, "content_type": ctype}
                for part, ctype in delta.content_types_removed
            ],
        },
        "paragraphs": {
            "added": [
                _paragraph_entry(positions_b, key, after.paragraphs[key].text)
                for key in delta.paragraphs_added
            ],
            "removed": [
                _paragraph_entry(positions_a, key, before.paragraphs[key].text)
                for key in delta.paragraphs_removed
            ],
            "changed": [
                _paragraph_change(positions_b, before, after, change)
                for change in delta.paragraphs_changed
            ],
        },
        "tables": {
            "added": [_table_entry(key) for key in delta.tables_added],
            "removed": [_table_entry(key) for key in delta.tables_removed],
            "changed": [_table_change(change) for change in delta.tables_changed],
        },
        "relationships": {
            "added": [
                {"part": part, "type": rtype, "target": target}
                for part, rtype, target in delta.relationships_added
            ],
            "removed": [
                {"part": part, "type": rtype, "target": target}
                for part, rtype, target in delta.relationships_removed
            ],
        },
        "sections_changed": [item.field for item in delta.sect_pr_changed],
        "counters_changed": [
            {"field": item.field, "before": item.before, "after": item.after}
            for item in delta.counters_changed
        ],
    }


TOOLS = [
    ToolSpec(
        fn=doc_compare,
        annotations=ToolAnnotations(title="Compare Documents", readOnlyHint=True),
        tags=frozenset({"v2", "read"}),
    ),
]
