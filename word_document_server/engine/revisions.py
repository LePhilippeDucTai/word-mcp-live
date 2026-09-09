"""Tracked revisions: recording them, listing them, applying them.

This layer owns ``w:ins`` and ``w:del``.  Everything below it treats them as
structure to preserve -- :mod:`~word_document_server.engine.ranges` refuses a
range that reaches into deleted content and never rewrites a ``w:del`` -- so a
revision is created, accepted or rejected *here* or not at all.

It replaces ``core/tracked_changes.py``, whose two documented defects are the
reason this module exists:

* its replacement loop re-searched the paragraph for the old text after writing
  the new one, so a replacement whose new text contains the old text never
  terminated.  Nothing here loops on a re-search: a range is resolved once, on
  offsets computed before the edit, and every mutation works from element
  references rather than from a text scan.
* its accept/reject unwrapped ``w:ins`` and dropped ``w:del`` in the main
  document part only, silently ignored every other revision kind, and removed a
  deleted paragraph mark without merging the paragraphs -- reporting a success
  that had left the document half-applied.  :func:`accept` and :func:`reject`
  here work on every story, refuse as a whole rather than apply in part, and
  merge paragraphs when a deleted paragraph mark is accepted.

Recording a revision
--------------------
:func:`tracked_delete` hides text instead of removing it, :func:`tracked_insert`
marks new text as inserted, and :func:`tracked_replace` does both so that the
deletion and the insertion sit next to each other and carry the same timestamp.
All three take the package as their first argument, because the ``w:id`` they
write has to be unique across the *whole* package and
:func:`~word_document_server.engine.ids.next_annotation_id` is the only thing
that knows that.

Nesting is decided by the author of the insertion a run already sits in, which
is what Word does:

======================================  ==========================================
run to delete                           result
======================================  ==========================================
plain text                              wrapped in ``w:del``
inside a ``w:ins`` by another author     wrapped in ``w:del`` *inside* that
                                        ``w:ins`` -- both revisions stand
inside a ``w:ins`` by the same author    removed outright: an author deleting
                                        their own pending insertion undoes it
======================================  ==========================================

Symmetrically, :func:`tracked_insert` never writes a ``w:ins`` inside another
one: text typed inside another author's insertion splits it in two and stands
between the halves, and text typed inside the author's own insertion simply
joins it.

Splitting an insertion never takes a run out of the container it sits in.  When
another author's ``w:ins`` holds a ``w:hyperlink``, a ``w:sdt``, a
``w:smartTag``, a ``w:customXml``, a ``w:dir`` or a ``w:bdo``, the insertion is
split around that container and then *pushed inside* it: every stretch of
content that is not on the path to the new run is re-wrapped in a copy of the
insertion, with a fresh id, so that ``w:hyperlink/w:ins`` -- the form Word
itself writes -- is what comes out.  The container is never cloned, never split
and never left by a run, because a hyperlink cut in two is two hyperlinks and a
run moved out of one has silently lost its link.  The trade-off is that the
container stops being carried by the enclosing insertion: rejecting every
revision it now holds empties it instead of removing it.

Listing
-------
:func:`list_revisions` reports the revisions attached to a paragraph or to its
content: ``ins``, ``del``, ``moveFrom``, ``moveTo``, ``rPrChange``,
``pPrChange`` and the two paragraph-mark revisions.  Revisions that belong to a
table row, a cell, a section or the numbering (``w:trPr/w:ins``,
``w:tblPrChange``, ``w:sectPrChange``, ``w:cellIns``, ...) are *not* reported --
they are not paragraph-level and this part does not rewrite them -- but
:func:`accept` and :func:`reject` still refuse to run when the selection would
cover one, so a caller is never told the document was fully applied while such a
revision remains.  Revisions inside ``word/comments.xml`` are equally out of
reach: it is not a story.

``Revision.paragraph_index`` is the *V2 index* -- the same index space
:class:`~word_document_server.engine.find.Match` reports, and deliberately not a
fourth one: the rank of the paragraph among the top-level paragraphs of its
story, block ``w:sdt`` content included, table cells excluded, hence ``None``
for a paragraph in a cell and for a revision that is not in a paragraph at all.

Applying
--------
:func:`accept` and :func:`reject` are all-or-nothing.  They select revisions by
id and by author, and if anything selected is of a kind they cannot apply they
raise :class:`~word_document_server.engine.errors.UnsupportedRevision` listing
the ids *before* touching the document.  What they do apply:

``ins``
    accept unwraps it and keeps the text; reject removes it and its content.
``del``
    accept removes it and its content; reject unwraps it, turning ``w:delText``
    back into ``w:t``.
``paragraph-mark-del``
    accept removes the mark and merges the paragraph with the following one,
    which then governs the result -- a paragraph mark carries the paragraph's
    properties, so the surviving mark is the next paragraph's.  Reject only
    removes the revision, leaving the two paragraphs apart.
``paragraph-mark-ins``
    the mirror image: accept only removes the revision, leaving the paragraph
    the reviewer added where it stands; reject removes the mark *and* merges the
    paragraph into the following one, undoing the split.  The next paragraph
    governs the result there too, for the same reason.

A merge needs a paragraph to merge into, and the following sibling is not always
one -- a table, or the end of the story, stops it.  Accepting a deleted mark or
rejecting an inserted one in that position is refused with the ids, before any
mutation, rather than applied as a bare mark removal that would leave the split
in place while reporting it undone.

Removing a ``w:ins`` or a ``w:del`` never drops a marker: bookmarks, comment
ranges and permission ranges found inside are moved to the position the element
occupied, exactly as ``ranges`` does when it deletes a run.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
from typing import TYPE_CHECKING, Literal

from lxml import etree

from word_document_server.engine.errors import LocatorError, UnsupportedRevision

# The V2 paragraph index has exactly one definition in the engine, and it is
# find's.  Importing it -- private though it is -- is what keeps
# ``Revision.paragraph_index`` and ``Match.index`` from drifting into two index
# spaces that merely look alike.
from word_document_server.engine.find import _v2_index_map, iter_paragraphs
from word_document_server.engine.ids import next_annotation_id
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import Pieces, insert_text, resolve
from word_document_server.engine.textmodel import Segment, segments
from word_document_server.engine.xmlns import qn

if TYPE_CHECKING:
    from collections.abc import Collection, Iterator

__all__ = [
    "REVISION_KINDS",
    "SUPPORTED_KINDS",
    "Revision",
    "RevisionKind",
    "accept",
    "list_revisions",
    "reject",
    "tracked_delete",
    "tracked_insert",
    "tracked_replace",
]

#: What :func:`list_revisions` reports.  ``paragraph-mark-ins`` and
#: ``paragraph-mark-del`` are the ``w:ins`` / ``w:del`` that live in
#: ``w:pPr/w:rPr`` and revise the paragraph mark itself rather than any run.
RevisionKind = Literal[
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "rPrChange",
    "pPrChange",
    "paragraph-mark-ins",
    "paragraph-mark-del",
]

#: The runtime form of :data:`RevisionKind`.
REVISION_KINDS: frozenset[str] = frozenset(
    {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "rPrChange",
        "pPrChange",
        "paragraph-mark-ins",
        "paragraph-mark-del",
    }
)

#: The kinds :func:`accept` and :func:`reject` know how to apply.  Anything else
#: is refused rather than skipped; see the module docstring.
SUPPORTED_KINDS: frozenset[str] = frozenset(
    {"ins", "del", "paragraph-mark-del", "paragraph-mark-ins"}
)

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_TBL = qn("w:tbl")
_W_PPR = qn("w:pPr")
_W_RPR = qn("w:rPr")
_W_TRPR = qn("w:trPr")
_W_INS = qn("w:ins")
_W_DEL = qn("w:del")
_W_MOVE_FROM = qn("w:moveFrom")
_W_MOVE_TO = qn("w:moveTo")
_W_RPR_CHANGE = qn("w:rPrChange")
_W_PPR_CHANGE = qn("w:pPrChange")
_W_T = qn("w:t")
_W_DEL_TEXT = qn("w:delText")
_W_INSTR_TEXT = qn("w:instrText")
_W_DEL_INSTR_TEXT = qn("w:delInstrText")
_W_ID = qn("w:id")
_W_AUTHOR = qn("w:author")
_W_DATE = qn("w:date")
_W_SDT = qn("w:sdt")
_W_SDT_CONTENT = qn("w:sdtContent")

#: Property children of an inline container: they describe the container, they
#: are not content, and wrapping one in a ``w:ins`` would be nonsense.
_CONTAINER_PROPERTIES = frozenset(
    qn(tag) for tag in ("w:sdtPr", "w:sdtEndPr", "w:smartTagPr", "w:customXmlPr")
)

#: Elements whose textual content is stored but not shown.  A ``w:t`` under one
#: of them is already a ``w:delText``; the conversions below stop there so a
#: nested revision is never rewritten by its parent's.
_HIDING = frozenset({_W_DEL, _W_MOVE_FROM})

#: Tag -> kind, for the revision elements that are not ``w:ins`` / ``w:del``
#: (whose kind depends on where they sit; see :func:`_classify`).
_PLAIN_KINDS: dict[str, RevisionKind] = {
    _W_MOVE_FROM: "moveFrom",
    _W_MOVE_TO: "moveTo",
    _W_RPR_CHANGE: "rPrChange",
    _W_PPR_CHANGE: "pPrChange",
}

#: Every element the scan looks at, in one tuple so that ``root.iter`` visits
#: them in a single document-order pass.
_SCANNED_TAGS: tuple[str, ...] = (
    _W_INS,
    _W_DEL,
    _W_MOVE_FROM,
    _W_MOVE_TO,
    _W_RPR_CHANGE,
    _W_PPR_CHANGE,
    qn("w:tblPrChange"),
    qn("w:trPrChange"),
    qn("w:tcPrChange"),
    qn("w:sectPrChange"),
    qn("w:tblGridChange"),
    qn("w:numberingChange"),
    qn("w:cellIns"),
    qn("w:cellDel"),
    qn("w:cellMerge"),
)


# --------------------------------------------------------------------------
# Value types
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Revision:
    """One tracked revision, as :func:`list_revisions` reports it.

    Plain data: no element reference, so a listing stays meaningful after the
    document has been edited.  :func:`accept` and :func:`reject` address a
    revision by :attr:`id`, never by identity.

    Attributes:
        id: the ``w:id`` of the revision element, or ``None`` when it carries
            none or an unparsable one.  Such a revision cannot be selected by
            id; it is still applied by a call that selects everything.
        kind: see :data:`RevisionKind`.
        author: the ``w:author``, ``""`` when the element has none.
        date: the ``w:date``, verbatim, ``""`` when the element has none.
            Written as an ISO 8601 UTC stamp (``2026-09-09T10:11:12Z``) by the
            functions of this module; read as-is from an existing document.
        story: story id, as
            :meth:`~word_document_server.engine.package.DocxPackage.stories`
            names it (``"document"``, ``"header1"``, ``"footnotes"``, ...).
        paragraph_index: the paragraph's V2 index in its story (see the module
            docstring), ``None`` for a paragraph inside a table cell and for a
            revision that does not sit in a paragraph.
        text: the text the revision applies to -- what a ``w:ins`` or a
            ``w:moveTo`` shows, the suppressed text a ``w:del`` or a
            ``w:moveFrom`` holds, the text of the run a ``w:rPrChange``
            reformats, and the paragraph's visible text for a ``w:pPrChange``
            or a paragraph-mark revision.
    """

    id: int | None
    kind: RevisionKind
    author: str
    date: str
    story: str
    paragraph_index: int | None
    text: str


@dataclass(frozen=True)
class _Entry:
    """A reported revision and the live element it was read from."""

    revision: Revision
    element: etree._Element
    root: etree._Element
    paragraph: etree._Element | None


@dataclass(frozen=True)
class _Foreign:
    """A revision this part does not report and cannot apply."""

    id: int | None
    name: str
    author: str


class _Ids:
    """Hands out annotation ids that no element of the package uses.

    :func:`~word_document_server.engine.ids.next_annotation_id` is pure -- it
    reads the largest id in use and returns the next one without stamping it --
    so one call plus a local counter is enough for a whole operation, and cheaper
    than walking the package once per element written.
    """

    def __init__(self, pkg: DocxPackage) -> None:
        self._next = next_annotation_id(pkg)

    def take(self) -> int:
        value = self._next
        self._next += 1
        return value


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------


def _checked_author(author: str) -> str:
    """Return `author`, or refuse it.

    Raises:
        ValueError: if `author` is not a non-empty string.  ``w:author`` is
            required on every revision element, and an empty one makes the
            "same author" rules of this module meaningless.
    """
    if not isinstance(author, str) or not author:
        raise ValueError("author must be a non-empty string")
    return author


def _stamp(date: str | datetime | None) -> str:
    """Return the ``w:date`` value to write.

    ``None`` means now.  A naive :class:`~datetime.datetime` is read as UTC, an
    aware one is converted to it, and the result is second-precision ISO 8601
    with the ``Z`` suffix Word writes.  A string is trusted and returned as is,
    so a caller can reproduce an exact stamp.
    """
    if isinstance(date, str):
        return date
    if date is None:
        moment = datetime.now(UTC)
    elif date.tzinfo is None:
        moment = date.replace(tzinfo=UTC)
    else:
        moment = date.astimezone(UTC)
    return moment.replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _annotate(element: etree._Element, change_id: int, author: str, date: str) -> None:
    """Stamp a revision element with its id, author and date."""
    element.set(_W_ID, str(change_id))
    element.set(_W_AUTHOR, author)
    element.set(_W_DATE, date)


def _read_id(element: etree._Element) -> int | None:
    """The ``w:id`` of `element` as an int, or ``None`` if it has no usable one."""
    raw = element.get(_W_ID)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _enclosing(node: etree._Element, tag: str, stop: etree._Element) -> etree._Element | None:
    """The innermost ancestor of `node` tagged `tag`, searching below `stop`."""
    for ancestor in node.iterancestors():
        if ancestor is stop:
            return None
        if ancestor.tag == tag:
            return ancestor
    return None


def _hidden_by(node: etree._Element, stop: etree._Element) -> etree._Element | None:
    """The innermost ``w:del`` / ``w:moveFrom`` around `node`, up to `stop`."""
    for ancestor in node.iterancestors():
        if ancestor.tag in _HIDING:
            return ancestor
        if ancestor is stop:
            break
    return None


def _paragraph_of(node: etree._Element) -> etree._Element | None:
    """The innermost ``w:p`` `node` sits in, or ``None``."""
    return next(iter(node.iterancestors(_W_P)), None)


def _attached(element: etree._Element, root: etree._Element) -> bool:
    """Whether `element` is still part of the tree rooted at `root`.

    Applying a revision can carry another one away with it -- rejecting a
    ``w:ins`` removes the ``w:del`` nested in it -- so an element selected before
    the first mutation may be gone by the time its turn comes.
    """
    return element.getroottree().getroot() is root


def _markers_inside(element: etree._Element) -> list[etree._Element]:
    """The zero-width markers `element` holds, in document order.

    Read through :func:`~word_document_server.engine.textmodel.segments` rather
    than from a tag list of our own, so "what a marker is" keeps exactly one
    definition in the engine.
    """
    paragraph = _paragraph_of(element)
    if paragraph is None:
        return []
    return [
        segment.element
        for segment in segments(paragraph)
        if segment.kind == "marker" and element in segment.containers
    ]


def _detach(element: etree._Element) -> None:
    """Remove `element` and its content, leaving its markers where it was.

    A bookmark or a comment range that framed deleted text survives the text:
    the markers are moved to the position `element` occupied, in order.
    """
    parent = element.getparent()
    if parent is None:
        return
    markers = _markers_inside(element)
    # Read the index before moving anything: the markers come from *inside*
    # `element`, so pulling them out does not shift it, and each one inserted
    # before it pushes it one further right -- which `offset` follows.
    index = parent.index(element)
    for offset, marker in enumerate(markers):
        parent.insert(index + offset, marker)
    parent.remove(element)


def _unwrap(element: etree._Element) -> None:
    """Replace `element` by its children, in place and in order."""
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for offset, child in enumerate(list(element)):
        parent.insert(index + offset, child)
    parent.remove(element)


def _wrap(
    runs: list[etree._Element], tag: str, change_id: int, author: str, date: str
) -> etree._Element:
    """Build a revision element around `runs`, where the first one stood.

    `runs` must be consecutive siblings.  The element is created as a child of
    their parent and only then moved into place, the way
    :func:`~word_document_server.engine.ranges.split_run` does it: an index read
    before a mutation and reused after it is the defect this engine exists to
    avoid.
    """
    first = runs[0]
    parent = first.getparent()
    index = parent.index(first)
    element = etree.SubElement(parent, tag)
    _annotate(element, change_id, author, date)
    parent.insert(index, element)
    for run in runs:
        element.append(run)
    return element


def _hide_text(element: etree._Element) -> None:
    """Turn the text `element` now holds into deleted text.

    ``w:t`` becomes ``w:delText`` and ``w:instrText`` becomes
    ``w:delInstrText``; attributes, ``xml:space`` included, ride along with the
    tag.  Content already hidden by a nested ``w:del`` or ``w:moveFrom`` is left
    alone: it belongs to that revision, not to this one.
    """
    for node in list(element.iter(_W_T, _W_INSTR_TEXT)):
        if _hidden_by(node, element) is not element:
            continue
        node.tag = _W_DEL_TEXT if node.tag == _W_T else _W_DEL_INSTR_TEXT


def _show_text(element: etree._Element) -> None:
    """The inverse of :func:`_hide_text`, for a rejected deletion."""
    for node in list(element.iter(_W_DEL_TEXT, _W_DEL_INSTR_TEXT)):
        if _hidden_by(node, element) is not element:
            continue
        node.tag = _W_T if node.tag == _W_DEL_TEXT else _W_INSTR_TEXT


# --------------------------------------------------------------------------
# Recording a revision
# --------------------------------------------------------------------------


def _treatment(run: etree._Element, paragraph: etree._Element, author: str) -> str:
    """How a covered run must be deleted: ``"wrap"`` or ``"remove"``.

    A run sitting in an insertion the same author made has not been seen by
    anyone else yet, so deleting it removes it outright rather than recording a
    deletion of one's own insertion -- Word's rule.  Any other run is wrapped in
    a ``w:del`` *where it stands*, which is what puts the ``w:del`` inside
    another author's ``w:ins`` without any special case.
    """
    enclosing = _enclosing(run, _W_INS, paragraph)
    if enclosing is not None and enclosing.get(_W_AUTHOR, "") == author:
        return "remove"
    return "wrap"


def _groups(
    runs: tuple[etree._Element, ...], paragraph: etree._Element, author: str
) -> Iterator[tuple[str, list[etree._Element]]]:
    """Split `runs` into consecutive stretches that can share one ``w:del``.

    A stretch breaks on a change of treatment, of parent, or of adjacency: two
    runs separated by a bookmark must not be pulled into the same ``w:del``,
    because that would move the marker out of the text it sits in.
    """
    current: list[etree._Element] = []
    kind = ""
    for run in runs:
        treatment = _treatment(run, paragraph, author)
        joins = (
            current
            and treatment == kind
            and run.getparent() is current[-1].getparent()
            and current[-1].getnext() is run
        )
        if not joins:
            if current:
                yield kind, current
            current = []
            kind = treatment
        current.append(run)
    if current:
        yield kind, current


def tracked_delete(
    pkg: DocxPackage,
    pieces: Pieces,
    author: str,
    date: str | datetime | None = None,
) -> tuple[etree._Element, ...]:
    """Record the deletion of `pieces` as a revision, and return the ``w:del``.

    The text is hidden, never removed: every run the range covers whole moves
    into a ``w:del`` where it stood, its ``w:t`` becoming ``w:delText``, so a
    later :func:`reject` puts it back byte for byte.  Runs that already sit in
    an insertion by `author` are removed instead -- see the module docstring for
    the table of cases -- so the returned tuple can be shorter than the number
    of runs, and empty when the whole range was the author's own pending
    insertion.  Consecutive sibling runs treated the same way share one
    ``w:del``, and therefore one id.

    `pieces` comes from :func:`~word_document_server.engine.ranges.resolve` and
    must have been resolved for ``"delete"`` or ``"replace"``: those are the
    modes that already refused a range cutting a field, swallowing an image or
    reaching into content another revision hides, and this function relies on
    those guarantees rather than re-deriving them.

    Args:
        pkg: the package `pieces` belongs to, for id allocation.
        pieces: the resolved range, whose runs are deleted.
        author: value of ``w:author``; also decides the nesting.
        date: value of ``w:date``; ``None`` means now.  See :func:`_stamp`.

    Raises:
        ValueError: if `author` is empty.
        UnsupportedRevision: if `pieces` was resolved for reading or insertion,
            or covers no whole run.
    """
    author = _checked_author(author)
    stamp = _stamp(date)
    if pieces.operation not in ("delete", "replace"):
        raise UnsupportedRevision(
            f"pieces resolved for {pieces.operation!r} cannot be deleted; resolve the "
            "range with operation='delete' so that fields, images and hidden content "
            "are checked first"
        )
    if not pieces.runs:
        raise UnsupportedRevision(
            f"range {pieces.start}..{pieces.end} covers no whole run to delete"
        )

    allocator = _Ids(pkg)
    created: list[etree._Element] = []
    for treatment, group in _groups(pieces.runs, pieces.paragraph, author):
        if treatment == "remove":
            for run in group:
                parent = run.getparent()
                if parent is not None:
                    parent.remove(run)
            continue
        element = _wrap(group, _W_DEL, allocator.take(), author, stamp)
        _hide_text(element)
        created.append(element)
    return tuple(created)


def _copy_around(
    nodes: list[etree._Element], template: etree._Element, allocator: _Ids
) -> etree._Element:
    """Wrap `nodes` in a copy of `template` with a fresh id, where they stand.

    `nodes` must be consecutive siblings.  Every attribute of `template` rides
    along -- author and date above all -- except ``w:id``, since two revision
    elements must never share one.  `template` itself is only read, and may
    already be detached from the tree.
    """
    first = nodes[0]
    parent = first.getparent()
    index = parent.index(first)
    element = etree.SubElement(parent, template.tag)
    for key, value in template.attrib.items():
        element.set(key, value)
    element.set(_W_ID, str(allocator.take()))
    parent.insert(index, element)
    for node in nodes:
        element.append(node)
    return element


def _split_around(container: etree._Element, child: etree._Element, allocator: _Ids) -> None:
    """Move `child` out of `container`, keeping the document order intact.

    `child` must be a *direct* child of `container`; it ends up as a sibling of
    it: before it when nothing preceded it inside, after it when nothing
    followed, and between the two halves otherwise -- the second half being a
    copy of `container` with a fresh id.  A `container` the move emptied is
    removed rather than left behind as a revision element covering nothing.
    """
    parent = container.getparent()
    preceding = list(child.itersiblings(preceding=True))
    following = list(child.itersiblings())
    if not preceding:
        parent.insert(parent.index(container), child)
    elif not following:
        parent.insert(parent.index(container) + 1, child)
    else:
        parent.insert(parent.index(container) + 1, child)
        tail = _copy_around(following, container, allocator)
        parent.insert(parent.index(child) + 1, tail)
    if len(container) == 0:
        parent.remove(container)


def _child_holding(ancestor: etree._Element, node: etree._Element) -> etree._Element:
    """The child of `ancestor` that is `node` or holds it."""
    current = node
    while current.getparent() is not ancestor:
        current = current.getparent()
    return current


def _content_of(container: etree._Element) -> etree._Element:
    """Where `container` keeps its content: ``w:sdtContent`` for a ``w:sdt``."""
    if container.tag == _W_SDT:
        found = container.find(_W_SDT_CONTENT)
        if found is not None:
            return found
    return container


def _push_into(
    container: etree._Element,
    run: etree._Element,
    template: etree._Element,
    allocator: _Ids,
) -> None:
    """Re-record `container`'s content as `template`, sparing the path to `run`.

    Walks down from `container` to `run`.  At every step, the content that
    precedes the path and the content that follows it are each wrapped in their
    own copy of `template`, so that the revision `container` was carried by is
    now carried by everything inside it *except* the new run -- which is what
    lets :func:`tracked_insert` leave the container itself alone.  Property
    children (:data:`_CONTAINER_PROPERTIES`) are never wrapped.
    """
    node = container
    while node is not run:
        content = _content_of(node)
        step = _child_holding(content, run)
        siblings = [child for child in content if child.tag not in _CONTAINER_PROPERTIES]
        at = siblings.index(step)
        for group in (siblings[:at], siblings[at + 1 :]):
            if group:
                _copy_around(group, template, allocator)
        node = step


def _free_of_insertions(
    run: etree._Element, paragraph: etree._Element, author: str, allocator: _Ids
) -> bool:
    """Take `run` out of every insertion it must not be nested in.

    Returns ``True`` when the caller still has to wrap it in a ``w:ins`` of its
    own, and ``False`` when an enclosing insertion already attributes the run to
    `author` -- in which case adding one would record the same text as inserted
    twice.

    A run that is a direct child of the insertion simply splits it.  A run
    nested deeper sits in an inline container -- a hyperlink, a content control,
    a smart tag -- and the container is what the insertion is split around; the
    insertion is then pushed inside it by :func:`_push_into`, so the container
    is neither cloned nor left by the run.  See the module docstring.  Either
    way the loop terminates: each pass removes one ``w:ins`` from the run's
    ancestry and adds none.
    """
    while True:
        enclosing = _enclosing(run, _W_INS, paragraph)
        if enclosing is None:
            return True
        if enclosing.get(_W_AUTHOR, "") == author:
            return False
        child = _child_holding(enclosing, run)
        _split_around(enclosing, child, allocator)
        if child is not run:
            _push_into(child, run, enclosing, allocator)


def tracked_insert(
    pkg: DocxPackage,
    paragraph: object,
    offset: int,
    text: str,
    author: str,
    date: str | datetime | None = None,
    *,
    rpr_from: object = "left",
) -> tuple[etree._Element, ...]:
    """Insert `text` at `offset` as a revision, and return the ``w:ins``.

    The run is built by
    :func:`~word_document_server.engine.ranges.insert_text` -- same offsets, same
    ``rpr_from`` rule, same refusals -- and then wrapped in a ``w:ins``.  It is
    therefore a clone of its neighbour's formatting, ``\\n`` is a ``w:br`` and
    ``\\t`` a ``w:tab``, and an insertion point inside a field or inside deleted
    content is refused before anything is written.

    An empty `text` writes nothing and returns ``()``.  So does an insertion
    landing inside an insertion `author` already owns: the text is added to it
    rather than nested in a second one.  Inside *another* author's insertion,
    that insertion is split and the new ``w:ins`` stands between the halves.

    Raises:
        ValueError: if `author` is empty.
        InvalidText: if `text` holds a character a run cannot carry.
        LocatorError: ``out-of-range`` if `offset` is outside the visible text.
        UnsupportedRange: if `offset` has no writable position.
    """
    author = _checked_author(author)
    stamp = _stamp(date)
    runs = insert_text(paragraph, offset, text, rpr_from=rpr_from)
    if not runs:
        return ()
    allocator = _Ids(pkg)
    created: list[etree._Element] = []
    for run in runs:
        element = _paragraph_of(run)
        if element is None:  # pragma: no cover - insert_text always writes in a w:p
            continue
        if not _free_of_insertions(run, element, author, allocator):
            continue
        created.append(_wrap([run], _W_INS, allocator.take(), author, stamp))
    return tuple(created)


def tracked_replace(
    pkg: DocxPackage,
    paragraph: object,
    start: int,
    end: int,
    text: str,
    author: str,
    date: str | datetime | None = None,
) -> tuple[etree._Element, ...]:
    """Replace ``[start, end)`` with `text` as one revision pair.

    The deletion and the insertion are adjacent and carry the same stamp, and
    the new text lands right after the old one, which is the order Word writes:
    ``<w:del>old</w:del><w:ins>new</w:ins>``.  The replacement inherits the
    ``w:rPr`` of the first run of the range, like
    :func:`~word_document_server.engine.ranges.replace_range`.  An empty `text`
    makes this a plain tracked deletion.

    The insertion is written *before* the deletion is recorded, at the offset
    the range ends on, so the range is addressed on offsets that predate every
    mutation.  Nothing re-searches the paragraph for `text` afterwards: a
    replacement whose new text contains the old one -- the case that made
    ``core/tracked_changes.py`` loop forever -- terminates here like any other.

    Returns:
        The revision elements created, deletions first, in document order.

    Raises:
        ValueError: if `author` is empty.
        InvalidText: if `text` holds a character a run cannot carry.
        LocatorError: ``out-of-range`` for impossible offsets.
        UnsupportedRange: see
            :mod:`word_document_server.engine.ranges` for the reasons.
    """
    author = _checked_author(author)
    stamp = _stamp(date)
    pieces = resolve(paragraph, start, end, operation="replace")
    element = pieces.paragraph

    inserted: tuple[etree._Element, ...] = ()
    if text:
        anchor: object = pieces.runs[0] if pieces.runs else "left"
        inserted = tracked_insert(pkg, element, end, text, author, stamp, rpr_from=anchor)
        # The offsets still designate the same characters: the new text was
        # written at `end`, past the range, so it shifted nothing before it.
        pieces = resolve(element, start, end, operation="replace")
    deleted = tracked_delete(pkg, pieces, author, stamp)
    return (*deleted, *inserted)


# --------------------------------------------------------------------------
# Listing
# --------------------------------------------------------------------------


def _classify(element: etree._Element) -> RevisionKind | None:
    """The kind `element` reports as, or ``None`` when it is not paragraph-level.

    ``w:ins`` and ``w:del`` mean three different things depending on where they
    sit: a revision of the paragraph mark in ``w:pPr/w:rPr``, a revision of a
    table row in ``w:trPr``, and a revision of content anywhere else.
    """
    tag = element.tag
    plain = _PLAIN_KINDS.get(tag)
    if plain is not None:
        return plain
    if tag not in (_W_INS, _W_DEL):
        return None
    suffix = "ins" if tag == _W_INS else "del"
    parent = element.getparent()
    if parent is None:
        return None
    if parent.tag == _W_TRPR:
        return None
    if parent.tag == _W_RPR:
        grandparent = parent.getparent()
        if grandparent is not None and grandparent.tag == _W_PPR:
            return f"paragraph-mark-{suffix}"  # type: ignore[return-value]
        return None
    return suffix  # type: ignore[return-value]


def _revision_text(
    kind: RevisionKind, element: etree._Element, found: list[Segment]
) -> str:
    """The text :attr:`Revision.text` reports for one revision element."""
    if kind in ("pPrChange", "paragraph-mark-ins", "paragraph-mark-del"):
        return "".join(segment.text for segment in found)
    if kind == "rPrChange":
        holder = element.getparent()
        owner = None if holder is None else holder.getparent()
        if owner is None or owner.tag != _W_R:
            # The run properties of the paragraph mark, not of any run.
            return "".join(segment.text for segment in found)
        return "".join(segment.text for segment in found if segment.run is owner)
    hidden = kind in ("del", "moveFrom")
    return "".join(
        segment.source_text if hidden else segment.text
        for segment in found
        if element in segment.containers
    )


def _scan(pkg: DocxPackage) -> tuple[list[_Entry], list[_Foreign]]:
    """Walk every story once and sort what it finds into reported and foreign."""
    entries: list[_Entry] = []
    foreign: list[_Foreign] = []
    for story, root in pkg.stories():
        index_map = _v2_index_map(iter_paragraphs(root))
        # Keyed by the element itself, not by id(): an lxml proxy is unique per
        # node only as long as something holds it, and the dict does.
        cache: dict[etree._Element, list[Segment]] = {}
        for element in root.iter(*_SCANNED_TAGS):
            kind = _classify(element)
            if kind is None:
                foreign.append(
                    _Foreign(
                        id=_read_id(element),
                        name=etree.QName(element).localname,
                        author=element.get(_W_AUTHOR, ""),
                    )
                )
                continue
            paragraph = _paragraph_of(element)
            if paragraph is None:
                # A revision of the reported kinds that is not in a paragraph is
                # not something this part knows how to place, let alone apply.
                foreign.append(
                    _Foreign(
                        id=_read_id(element),
                        name=etree.QName(element).localname,
                        author=element.get(_W_AUTHOR, ""),
                    )
                )
                continue
            found = cache.get(paragraph)
            if found is None:
                found = segments(paragraph)
                cache[paragraph] = found
            entries.append(
                _Entry(
                    revision=Revision(
                        id=_read_id(element),
                        kind=kind,
                        author=element.get(_W_AUTHOR, ""),
                        date=element.get(_W_DATE, ""),
                        story=story,
                        paragraph_index=index_map.get(paragraph),
                        text=_revision_text(kind, element, found),
                    ),
                    element=element,
                    root=root,
                    paragraph=paragraph,
                )
            )
    return entries, foreign


def list_revisions(pkg: DocxPackage) -> list[Revision]:
    """Return every paragraph-level revision of `pkg`, in document order.

    Stories come in the order
    :meth:`~word_document_server.engine.package.DocxPackage.stories` gives them
    -- the main document first -- then the revisions of each story in document
    order.  See the module docstring for what is reported and what is not.
    """
    entries, _ = _scan(pkg)
    return [entry.revision for entry in entries]


# --------------------------------------------------------------------------
# Applying
# --------------------------------------------------------------------------


def _selects(
    change_id: int | None, author: str, ids: Collection[int] | None, by: str | None
) -> bool:
    """Whether the ``ids`` / ``author`` filter covers one revision."""
    if ids is not None and (change_id is None or change_id not in ids):
        return False
    return by is None or author == by


def _merge_target(paragraph: etree._Element) -> etree._Element | None:
    """The paragraph a deleted paragraph mark merges into, or ``None``.

    The next paragraph among the following siblings.  A table in between stops
    the search: paragraphs on either side of one cannot be joined.
    """
    for sibling in paragraph.itersiblings():
        if sibling.tag == _W_P:
            return sibling
        if sibling.tag == _W_TBL:
            return None
    return None


def _merge_into_next(paragraph: etree._Element) -> None:
    """Move `paragraph`'s content into the next paragraph and drop `paragraph`.

    The target keeps its own ``w:pPr``: a paragraph mark carries the paragraph's
    properties, and it is the target's mark that survives the merge.
    """
    target = _merge_target(paragraph)
    if target is None:  # pragma: no cover - checked before any mutation
        return
    at = 1 if len(target) and target[0].tag == _W_PPR else 0
    for offset, child in enumerate(c for c in paragraph if c.tag != _W_PPR):
        target.insert(at + offset, child)
    parent = paragraph.getparent()
    if parent is not None:
        parent.remove(paragraph)


def _joins_paragraphs(kind: RevisionKind, *, accepting: bool) -> bool:
    """Whether applying `kind` this way has to merge the paragraph into the next.

    A paragraph mark that a reviewer deleted really disappears when the deletion
    is accepted, and a mark they inserted really disappears when the insertion is
    rejected: both leave one paragraph where there were two.  The other two
    combinations only drop the revision element.
    """
    if kind == "paragraph-mark-del":
        return accepting
    if kind == "paragraph-mark-ins":
        return not accepting
    return False


def _unsupported(refused: list[tuple[int | None, str]]) -> UnsupportedRevision:
    """Build the refusal, naming every id that stands in the way."""
    listed = ", ".join(
        f"{'?' if change_id is None else change_id} ({name})" for change_id, name in refused
    )
    error = UnsupportedRevision(
        f"{len(refused)} selected revision(s) cannot be applied: {listed}. "
        "Nothing was changed: narrow the selection with ids= or author=."
    )
    # Handy for a tool layer that wants to report the ids rather than the text.
    error.ids = tuple(change_id for change_id, _ in refused)  # type: ignore[attr-defined]
    return error


def _apply(
    pkg: DocxPackage,
    ids: Collection[int] | None,
    author: str | None,
    *,
    accepting: bool,
) -> list[Revision]:
    """Shared body of :func:`accept` and :func:`reject`."""
    wanted = None if ids is None else set(ids)
    entries, foreign = _scan(pkg)
    selected = [
        entry
        for entry in entries
        if _selects(entry.revision.id, entry.revision.author, wanted, author)
    ]
    caught = [item for item in foreign if _selects(item.id, item.author, wanted, author)]
    refused: list[tuple[int | None, str]] = [(item.id, item.name) for item in caught]

    if wanted is not None:
        seen = {entry.revision.id for entry in selected} | {item.id for item in caught}
        missing = sorted(wanted - seen)
        if missing:
            raise LocatorError(
                "not-found",
                f"no revision with id {missing} in the package"
                + ("" if author is None else f" for author {author!r}"),
            )

    for entry in selected:
        kind = entry.revision.kind
        if kind not in SUPPORTED_KINDS:
            refused.append((entry.revision.id, kind))
        elif (
            _joins_paragraphs(kind, accepting=accepting)
            and entry.paragraph is not None
            and _merge_target(entry.paragraph) is None
        ):
            refused.append((entry.revision.id, f"{kind} without a paragraph to merge into"))
    if refused:
        raise _unsupported(refused)

    applied: list[Revision] = []
    for entry in selected:
        # A nested revision may already have gone with the one that held it.
        if not _attached(entry.element, entry.root):
            applied.append(entry.revision)
            continue
        _apply_one(entry, accepting=accepting)
        applied.append(entry.revision)
    return applied


def _apply_one(entry: _Entry, *, accepting: bool) -> None:
    """Apply one selected revision to the live tree."""
    kind = entry.revision.kind
    element = entry.element
    if kind == "ins":
        if accepting:
            _unwrap(element)
        else:
            _detach(element)
        return
    if kind == "del":
        if accepting:
            _detach(element)
        else:
            _show_text(element)
            _unwrap(element)
        return
    # A paragraph-mark revision: the revision element goes either way, and the
    # paragraphs merge only in the direction that removes the mark for good.
    paragraph = entry.paragraph
    holder = element.getparent()
    if holder is not None:
        holder.remove(element)
    if paragraph is not None and _joins_paragraphs(kind, accepting=accepting):
        _merge_into_next(paragraph)


def accept(
    pkg: DocxPackage,
    ids: Collection[int] | None = None,
    author: str | None = None,
) -> list[Revision]:
    """Accept the selected revisions of `pkg` and return them.

    `ids` and `author` narrow the selection and combine: ``None`` for both
    accepts every revision of the package.  Selection happens before any
    mutation, and so does the check that every selected revision can be applied,
    so this either accepts the whole selection or changes nothing at all.

    An insertion loses its wrapper and keeps its text; a deletion goes away with
    its content, its markers moved to where it stood; a deleted paragraph mark
    merges its paragraph into the following one, which governs the result; an
    inserted paragraph mark simply loses the revision and the split stands.

    Returns:
        The revisions accepted, in document order.  A revision nested in another
        one is listed even though its holder carried it away.

    Raises:
        LocatorError: ``not-found`` if `ids` names a revision the package does
            not have -- a silent no-op would look like a success.
        UnsupportedRevision: if the selection covers a kind this part does not
            apply, or a deleted paragraph mark with no paragraph to merge into,
            listing the ids.  Nothing has been changed when it is raised.
    """
    return _apply(pkg, ids, author, accepting=True)


def reject(
    pkg: DocxPackage,
    ids: Collection[int] | None = None,
    author: str | None = None,
) -> list[Revision]:
    """Reject the selected revisions of `pkg` and return them.

    Same selection rules and same all-or-nothing contract as :func:`accept`.
    An insertion goes away with its content, its markers moved to where it
    stood; a deletion is unwrapped and its ``w:delText`` becomes ``w:t`` again;
    a deleted paragraph mark simply loses the revision, and the two paragraphs
    stay apart; an inserted paragraph mark loses the revision *and* its paragraph
    is merged into the following one, which governs the result.

    Rejecting what :func:`tracked_delete`, :func:`tracked_insert` or
    :func:`tracked_replace` recorded restores the document they were given.

    Returns:
        The revisions rejected, in document order.

    Raises:
        LocatorError: ``not-found`` if `ids` names a revision the package does
            not have.
        UnsupportedRevision: if the selection covers a kind this part does not
            apply, or an inserted paragraph mark with no paragraph to merge into,
            listing the ids.  Nothing has been changed when it is raised.
    """
    return _apply(pkg, ids, author, accepting=False)
