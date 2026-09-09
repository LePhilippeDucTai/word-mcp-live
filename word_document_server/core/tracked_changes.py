"""Tracked changes for the cross-platform tools, delegated to the engine.

This module is the thin adapter between ``tools/tracked_changes_tools.py`` --
whose function names, parameters and JSON shapes are fixed -- and
:mod:`word_document_server.engine.revisions`, which owns ``w:ins`` and
``w:del``.  It locates text, converts what it finds into engine calls and
formats the result; it never touches XML itself.

What the move to the engine fixes
---------------------------------
The previous implementation opened the ``.docx`` with :mod:`zipfile`, rewrote
``word/document.xml`` in place and carried its own copies of the run helpers.
Three defects came with that, and none of them can be written here any more:

* ``track_replace_in_doc`` re-searched the paragraph after every replacement,
  including the text it had just inserted, so ``"Risk" -> "Risk Risk"`` never
  returned.  Matches are now collected **once**, before any mutation, and
  applied right to left so that the offsets of the matches still pending are
  never shifted by the ones already applied.
* the splice reinserted the surviving fragments into the parent of the *first*
  matched run, which is the wrong element as soon as a match spans a
  ``w:hyperlink`` boundary, and rebuilt every fragment with the first run's
  ``w:rPr``.  :func:`~word_document_server.engine.ranges.resolve` splits runs in
  place instead, so each fragment keeps its own properties.
* the search ran over raw ``w:t`` text and could therefore match, and re-delete,
  text already hidden under a ``w:del``.
  :func:`~word_document_server.engine.find.find` reads the *visible* text of
  each paragraph, so deleted content is out of reach by construction.

Saving goes through :meth:`~word_document_server.engine.package.DocxPackage.save`,
which serialises to memory and writes atomically: an interrupted call leaves the
original file intact instead of a truncated archive.

Contract of the returned dictionaries
-------------------------------------
Unchanged, key for key.  A recording function returns ``success``/``error`` when
the text is not found and ``success`` plus its own counter (``replacements``,
``insertions``, ``deletions``) and ``message`` otherwise.  Only the free-text
``message`` gained detail: it now names what an apply left in place instead of
staying silent about it.

Failures the engine refuses on purpose -- a range that would cut a field, an
inserted paragraph mark with no paragraph to merge into -- are raised as
:class:`~word_document_server.engine.errors.EngineError`, not swallowed.  The
tool layer already turns any exception into ``{"success": false, "error": ...}``,
and nothing is written when one is raised: every function here saves once, at
the end, after all of its edits have succeeded.

Scope of each function
----------------------
Recording searches the body only, as it always did.  Listing and applying now
cover **every story** -- headers, footers, footnotes, endnotes -- because
:mod:`~word_document_server.engine.revisions` does; a revision in a header used
to be invisible to ``list_tracked_changes`` and untouched by ``accept``, which
reported a document fully applied while it was not.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import TYPE_CHECKING

from word_document_server.defaults import DEFAULT_AUTHOR

# The V2 paragraph index has one definition in the engine, and it is find's.
# Re-deriving it here to build the paragraph contexts of a listing would create a
# second index space that merely looks like the one `list_revisions` reports.
from word_document_server.engine.find import _v2_index_map, find, iter_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import resolve
from word_document_server.engine.revisions import (
    SUPPORTED_KINDS,
    accept,
    list_revisions,
    reject,
    tracked_delete,
    tracked_insert,
    tracked_replace,
)
from word_document_server.engine.textmodel import visible_text

if TYPE_CHECKING:
    from collections.abc import Iterable

    from word_document_server.engine.find import Match
    from word_document_server.engine.revisions import Revision

__all__ = [
    "accept_tracked_changes_in_doc",
    "list_tracked_changes_in_doc",
    "reject_tracked_changes_in_doc",
    "track_delete_in_doc",
    "track_insert_in_doc",
    "track_replace_in_doc",
]

#: Characters of paragraph text reported in ``paragraph_context``.
_CONTEXT_LIMIT = 100

#: Revision kinds reported as insertions and as deletions by
#: :func:`list_tracked_changes_in_doc`.  Exactly what the previous
#: ``root.iter(w:ins)`` / ``root.iter(w:del)`` scan reached: a paragraph-mark
#: revision lives in ``w:pPr/w:rPr`` and was matched by it too.  The four other
#: kinds the engine knows (``moveFrom``, ``moveTo``, ``rPrChange``,
#: ``pPrChange``) were never reported and still are not.
_INSERTION_KINDS = frozenset({"ins", "paragraph-mark-ins"})
_DELETION_KINDS = frozenset({"del", "paragraph-mark-del"})


# --------------------------------------------------------------------------
# Locating
# --------------------------------------------------------------------------


def _matches_last_first(pkg: DocxPackage, text: str) -> list[Match]:
    """Every occurrence of `text` in the body, latest in the document first.

    Applying a revision to ``[start, end)`` changes the visible text from
    `start` onwards and nothing before it, so working through the matches in
    reverse document order keeps the offsets of the ones still pending exactly
    as :func:`~word_document_server.engine.find.find` measured them.  This is
    what replaces the previous re-search loop, and with it the case where the
    replacement contained the text it replaced.
    """
    return list(reversed(find(pkg, text)))


def _not_found(text: str) -> dict:
    """The error dictionary a recording function returns for a missing text."""
    return {"success": False, "error": f"Text not found: '{text}'"}


# --------------------------------------------------------------------------
# Recording
# --------------------------------------------------------------------------


def track_replace_in_doc(
    filepath: str,
    old_text: str,
    new_text: str,
    author: str = DEFAULT_AUTHOR,
) -> dict:
    """Replace every occurrence of `old_text` with `new_text` as a revision.

    Each occurrence becomes the ``<w:del>old</w:del><w:ins>new</w:ins>`` pair
    Word writes, all of them stamped with one timestamp, as before.  Text hidden
    under an existing ``w:del`` is not searched and therefore never replaced.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place on success.
        old_text: text to mark as deleted; every occurrence in the body.
        new_text: replacement text, inserted right after the deleted run.  An
            empty string makes this a plain tracked deletion.
        author: ``w:author`` of both revisions.

    Returns:
        ``{"success": False, "error": ...}`` when `old_text` is nowhere in the
        body, otherwise ``{"success": True, "replacements": n, "message": ...}``.

    Raises:
        EngineError: if a range cannot be replaced -- it cuts a field, or
            swallows an image.  The file is left untouched.
    """
    pkg = DocxPackage.open(filepath)
    matches = _matches_last_first(pkg, old_text)
    if not matches:
        return _not_found(old_text)

    stamp = datetime.now(UTC)
    for match in matches:
        tracked_replace(
            pkg, match.paragraph, match.start, match.end, new_text, author, stamp
        )

    pkg.save(filepath)
    return {
        "success": True,
        "replacements": len(matches),
        "message": (
            f"Replaced {len(matches)} occurrence(s) of '{old_text}' with "
            f"'{new_text}' as tracked change by {author}"
        ),
    }


def track_insert_in_doc(
    filepath: str,
    after_text: str,
    insert_text: str,
    author: str = DEFAULT_AUTHOR,
) -> dict:
    """Insert `insert_text` right after the first occurrence of `after_text`.

    First occurrence only, as before.  The inserted run takes its formatting
    from the run on its left -- the end of the match -- and stays inside
    whatever container that run sits in: a hyperlink is never split by an
    insertion, the ``w:ins`` is pushed inside it instead.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place on success.
        after_text: text to search for in the body.
        insert_text: text to insert just past it.
        author: ``w:author`` of the insertion.

    Returns:
        ``{"success": False, "error": ...}`` when `after_text` is nowhere in the
        body, otherwise ``{"success": True, "insertions": 1, "message": ...}``.

    Raises:
        EngineError: if the insertion point is not writable -- inside a field
            instruction, inside deleted content.  The file is left untouched.
    """
    pkg = DocxPackage.open(filepath)
    matches = find(pkg, after_text, max_results=1)
    if not matches:
        return _not_found(after_text)

    match = matches[0]
    tracked_insert(pkg, match.paragraph, match.end, insert_text, author)

    pkg.save(filepath)
    return {
        "success": True,
        "insertions": 1,
        "message": (
            f"Inserted '{insert_text}' after '{after_text}' as tracked change "
            f"by {author}"
        ),
    }


def track_delete_in_doc(
    filepath: str,
    text: str,
    author: str = DEFAULT_AUTHOR,
) -> dict:
    """Mark every occurrence of `text` as deleted.

    The text is hidden, not removed: each covered run moves into a ``w:del``
    where it stood and its ``w:t`` becomes ``w:delText``, so rejecting the
    revision restores the document byte for byte.  Bookmarks and comment ranges
    caught inside a deleted stretch are kept.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place on success.
        text: text to mark as deleted; every occurrence in the body.
        author: ``w:author`` of the deletions.

    Returns:
        ``{"success": False, "error": ...}`` when `text` is nowhere in the body,
        otherwise ``{"success": True, "deletions": n, "message": ...}``.

    Raises:
        EngineError: if a range cannot be deleted -- it cuts a field, or
            swallows an image.  The file is left untouched.
    """
    pkg = DocxPackage.open(filepath)
    matches = _matches_last_first(pkg, text)
    if not matches:
        return _not_found(text)

    stamp = datetime.now(UTC)
    for match in matches:
        pieces = resolve(match.paragraph, match.start, match.end, operation="delete")
        tracked_delete(pkg, pieces, author, stamp)

    pkg.save(filepath)
    return {
        "success": True,
        "deletions": len(matches),
        "message": (
            f"Marked {len(matches)} occurrence(s) of '{text}' as deleted by {author}"
        ),
    }


# --------------------------------------------------------------------------
# Listing
# --------------------------------------------------------------------------


def _paragraph_texts(pkg: DocxPackage) -> dict[tuple[str, int], str]:
    """Visible text of every V2-indexed paragraph, keyed as a revision names it.

    :class:`~word_document_server.engine.revisions.Revision` is plain data and
    holds no element, so the paragraph context of a listing is looked up by
    ``(story, paragraph_index)``.  A paragraph outside the V2 index space -- in
    a table cell, in a text box -- has no key here, and a revision inside one
    therefore reports an empty context.
    """
    texts: dict[tuple[str, int], str] = {}
    for story, root in pkg.stories():
        for paragraph, index in _v2_index_map(iter_paragraphs(root)).items():
            texts[(story, index)] = visible_text(paragraph)
    return texts


def _entry(revision: Revision, contexts: dict[tuple[str, int], str]) -> dict:
    """One listing entry, in the shape the tool has always returned.

    ``id`` stays a string, as it was when it was read straight off the ``w:id``
    attribute; ``""`` when the revision carries none.  A paragraph-mark revision
    reports an empty ``text``: it revises the mark, not any run, and reporting
    the paragraph's own text there would read as inserted or deleted text.
    """
    mark = revision.kind.startswith("paragraph-mark-")
    key = (revision.story, revision.paragraph_index)
    context = "" if revision.paragraph_index is None else contexts.get(key, "")
    return {
        "id": "" if revision.id is None else str(revision.id),
        "author": revision.author or "Unknown",
        "date": revision.date,
        "text": "" if mark else revision.text,
        "paragraph_context": context[:_CONTEXT_LIMIT],
    }


def list_tracked_changes_in_doc(filepath: str) -> dict:
    """List the insertions and deletions of a document, in document order.

    Every story is scanned, not just the body: an insertion in a header is a
    tracked change like any other, and one that ``accept`` will apply.

    Returns:
        ``{"success": True, "insertions": [...], "deletions": [...],
        "total_insertions": n, "total_deletions": n, "total_changes": n}``,
        each entry carrying ``id``, ``author``, ``date``, ``text`` and
        ``paragraph_context``.
    """
    pkg = DocxPackage.open(filepath)
    contexts = _paragraph_texts(pkg)

    insertions: list[dict] = []
    deletions: list[dict] = []
    for revision in list_revisions(pkg):
        if revision.kind in _INSERTION_KINDS:
            insertions.append(_entry(revision, contexts))
        elif revision.kind in _DELETION_KINDS:
            deletions.append(_entry(revision, contexts))

    return {
        "success": True,
        "insertions": insertions,
        "deletions": deletions,
        "total_insertions": len(insertions),
        "total_deletions": len(deletions),
        "total_changes": len(insertions) + len(deletions),
    }


# --------------------------------------------------------------------------
# Applying
# --------------------------------------------------------------------------


def _wanted_ids(change_ids: Iterable[int] | None) -> set[int] | None:
    """The requested id filter as a set, or ``None`` for "every revision".

    Raises:
        ValueError: if an entry is not an integer.  Dropping it instead would
            apply a selection the caller never asked for, and report a success.
    """
    if change_ids is None:
        return None
    return {int(change_id) for change_id in change_ids}


def _selects(revision: Revision, ids: set[int] | None, author: str | None) -> bool:
    """Whether the ``change_ids`` / ``author`` filter covers `revision`."""
    if ids is not None and (revision.id is None or revision.id not in ids):
        return False
    return author is None or revision.author == author


def _message(
    head: str,
    unsupported: list[Revision],
    unaddressable: list[Revision],
    unknown: tuple[int, ...],
) -> str:
    """`head`, followed by everything the call deliberately did not apply.

    The head alone is the message this tool has always returned.  The clauses
    after it exist because the previous implementation reported ``"Accepted n
    tracked change(s)"`` while quietly leaving formatting revisions untouched
    and quietly ignoring an id the document did not carry, so a caller could
    not tell a fully applied document from a half applied one.
    """
    reasons: list[str] = []
    if unsupported:
        kinds = ", ".join(sorted({revision.kind for revision in unsupported}))
        reasons.append(
            f"{len(unsupported)} revision(s) left in place, of a kind this tool "
            f"does not apply ({kinds})"
        )
    if unaddressable:
        reasons.append(
            f"{len(unaddressable)} revision(s) left in place, carrying no usable w:id"
        )
    if unknown:
        listed = ", ".join(str(change_id) for change_id in unknown)
        reasons.append(
            f"{len(unknown)} requested id(s) matching no insertion or deletion: {listed}"
        )
    return head if not reasons else head + "; " + "; ".join(reasons)


def _apply(
    filepath: str,
    author: str | None,
    change_ids: Iterable[int] | None,
    *,
    accepting: bool,
) -> dict:
    """Shared body of :func:`accept_tracked_changes_in_doc` and its mirror.

    The engine applies a selection as a whole or refuses it as a whole, and it
    counts a revision it cannot apply as a reason to refuse -- including the
    row, cell and table-property revisions it does not even report.  This tool
    has never claimed to apply those, so it hands the engine an explicit list of
    the ids it does claim, rather than the open selection that would make an
    unrelated ``w:tblPrChange`` fail the whole call.  What is left out is named
    in ``message``, never dropped in silence.
    """
    verb, counter = ("Accepted", "accepted") if accepting else ("Rejected", "rejected")
    pkg = DocxPackage.open(filepath)
    wanted = _wanted_ids(change_ids)

    revisions = list_revisions(pkg)
    selected = [revision for revision in revisions if _selects(revision, wanted, author)]
    applicable = [
        revision
        for revision in selected
        if revision.kind in SUPPORTED_KINDS and revision.id is not None
    ]
    unsupported = [
        revision for revision in selected if revision.kind not in SUPPORTED_KINDS
    ]
    unaddressable = [
        revision
        for revision in selected
        if revision.kind in SUPPORTED_KINDS and revision.id is None
    ]
    unknown: tuple[int, ...] = (
        ()
        if wanted is None
        else tuple(
            sorted(
                wanted - {revision.id for revision in revisions if revision.id is not None}
            )
        )
    )

    if not applicable:
        head = f"No matching tracked changes found to {'accept' if accepting else 'reject'}"
        return {
            "success": True,
            "message": _message(head, unsupported, unaddressable, unknown),
            counter: 0,
        }

    apply_revisions = accept if accepting else reject
    applied = apply_revisions(pkg, [revision.id for revision in applicable], author)

    pkg.save(filepath)
    return {
        "success": True,
        counter: len(applied),
        "message": _message(
            f"{verb} {len(applied)} tracked change(s)",
            unsupported,
            unaddressable,
            unknown,
        ),
    }


def accept_tracked_changes_in_doc(
    filepath: str,
    author: str | None = None,
    change_ids: list[int] | None = None,
) -> dict:
    """Accept tracked changes: keep what was inserted, drop what was deleted.

    `author` and `change_ids` narrow the selection and combine, ``None`` for
    both accepting every revision of the package.  Accepting is all-or-nothing:
    the whole selection is checked before anything is written, so the file is
    either fully applied or untouched.

    A deleted paragraph mark now merges its paragraph into the following one,
    which is the effect Word applies and which the previous implementation
    reported without performing.  The formatting revisions this tool has never
    applied (``w:rPrChange``, ``w:pPrChange``) are still left in place, but the
    ``message`` now says so.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place when something
            was applied.
        author: only accept revisions carrying this ``w:author``.
        change_ids: only accept revisions carrying one of these ``w:id``.

    Returns:
        ``{"success": True, "accepted": n, "message": ...}``.  ``n`` is 0 and
        the file is left untouched when the selection matches nothing.

    Raises:
        EngineError: if the selection cannot be applied as a whole -- a deleted
            paragraph mark with no paragraph to merge into.  Nothing is written.
        ValueError: if `change_ids` holds a value that is not an integer.
    """
    return _apply(filepath, author, change_ids, accepting=True)


def reject_tracked_changes_in_doc(
    filepath: str,
    author: str | None = None,
    change_ids: list[int] | None = None,
) -> dict:
    """Reject tracked changes: drop what was inserted, restore what was deleted.

    Same selection rules and same all-or-nothing contract as
    :func:`accept_tracked_changes_in_doc`.  Rejecting an inserted paragraph mark
    merges its paragraph back into the following one, undoing the split the
    reviewer introduced.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place when something
            was applied.
        author: only reject revisions carrying this ``w:author``.
        change_ids: only reject revisions carrying one of these ``w:id``.

    Returns:
        ``{"success": True, "rejected": n, "message": ...}``.  ``n`` is 0 and
        the file is left untouched when the selection matches nothing.

    Raises:
        EngineError: if the selection cannot be applied as a whole -- an
            inserted paragraph mark with no paragraph to merge into.  Nothing is
            written.
        ValueError: if `change_ids` holds a value that is not an integer.
    """
    return _apply(filepath, author, change_ids, accepting=False)
