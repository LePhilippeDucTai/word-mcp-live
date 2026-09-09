"""Editing a paragraph by character range: split, delete, insert, replace.

This is the only layer of the engine that writes text into a document.  Every
tool migrated onto the V2 engine goes through it, so a refusal here is worth
more than an approximate mutation: the functions below raise
:class:`~word_document_server.engine.errors.UnsupportedRange` rather than guess
what a caller meant when a range cuts through something indivisible.

Offsets
-------
Start and end are offsets in the paragraph's *visible* text, the one
:func:`~word_document_server.engine.textmodel.visible_text` reads.  A range is
half open: ``[start, end)`` covers the characters ``start`` up to but excluding
``end``.  ``ranges`` never re-derives that text policy -- widths come from the
segments ``textmodel`` hands out, so a carrier taught to ``textmodel`` (a new
kind of break, a new symbol) is measured correctly here without a second table
to keep in sync.

What a range covers
-------------------
Characters are covered when they sit in ``[start, end)``.  Zero-width content --
an image, a field character, a bookmark, deleted text -- carries no character,
so it sits *at* one offset and the half-open rule cannot decide whether a
boundary offset is inside or outside.  The rule taken here is that such an item
is covered only when it is **strictly** inside, ``start < offset < end``.  A
footnote reference sitting exactly where a deletion begins or ends is therefore
left alone instead of being swept away or blocking the operation, and only what
the deleted characters really surround is at stake.

Splitting follows the same rule, which is why :func:`split_run` takes a
``zero_width`` side: at the *start* of a range, zero-width content at the
boundary belongs to the left (outside) part; at the *end*, it belongs to the
right (outside) part.  Once both boundaries are split this way, every run of the
paragraph is either entirely covered or entirely untouched, which is what makes
"delete the covered runs" a safe implementation.

What is refused
---------------
:class:`UnsupportedRange` reasons, all stable:

``empty-range``
    ``start == end`` for ``delete`` or ``replace``.  Use :func:`insert_text`.
``crosses-field``
    the range overlaps a field reported by ``textmodel.fields`` without
    containing it.  A field's markers, instruction and cached result are
    consistent only as a whole.
``contains-field``
    ``delete`` or ``replace`` over a range that swallows a whole field.
``inside-field``
    an insertion point strictly inside a field.
``hidden-content``
    the range strictly contains deleted or moved-away content (``w:del``,
    ``w:moveFrom``), or an insertion would land inside it.  ``w:del`` is never
    rewritten here; :mod:`word_document_server.engine.revisions` owns it.
``opaque-content``
    ``delete`` or ``replace`` over a range that strictly contains an image, an
    embedded object, a note or comment reference, or a field character.
``unsplittable-range``
    a boundary could not be aligned on a run, which means the paragraph holds a
    construct this layer does not know how to cut.

What is never removed
---------------------
Only ``w:r`` elements are ever removed, and only when the range covers them
whole.  Containers stay: an emptied ``w:hyperlink`` keeps its ``r:id`` (so its
relationship still resolves), an emptied ``w:ins`` keeps the revision it stands
for, an emptied ``w:sdt`` keeps the content control.  Markers found inside a
deleted range are moved to the position the range started at, never dropped, so
bookmarks and comment ranges survive the deletion of the text they framed.

Stale references
----------------
Every function re-reads the paragraph from ``textmodel`` after each mutation and
recomputes indices from the live tree.  Nothing here captures a parent and an
index before mutating and reuses them afterwards -- the defect that made
``core/tracked_changes.py`` splice runs into the wrong element.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import Literal

from lxml import etree

from word_document_server.engine.errors import (
    InvalidText,
    LocatorError,
    UnsupportedRange,
)
from word_document_server.engine.textmodel import Segment, fields, segments
from word_document_server.engine.xmlns import qn

__all__ = [
    "Operation",
    "Pieces",
    "delete_range",
    "insert_text",
    "replace_range",
    "resolve",
    "split_run",
]

#: What the caller intends to do with a range.  ``read`` is the non-destructive
#: mode used by :mod:`word_document_server.engine.format`: it still refuses to
#: cut through a field or to straddle hidden content, but it tolerates images
#: and references inside the range, which ``delete`` and ``replace`` do not.
Operation = Literal["read", "insert", "delete", "replace"]

#: Operations that remove content, and therefore refuse opaque objects.
_DESTRUCTIVE: frozenset[str] = frozenset({"delete", "replace"})

#: One of the two sides of a boundary: which part zero-width content joins when
#: a run is split exactly at its position, and which run an insertion takes its
#: formatting and its context from.
_Side = Literal["left", "right"]

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_RPR = qn("w:rPr")
_W_T = qn("w:t")
_W_BR = qn("w:br")
_W_TAB = qn("w:tab")
_XML_SPACE = qn("xml:space")

_W_FLD_CHAR = qn("w:fldChar")
_W_FLD_CHAR_TYPE = qn("w:fldCharType")

#: Containers an insertion must never land inside, and the reason it is refused:
#: text put there would be invisible (``w:del``, ``w:moveFrom``) or would be
#: wiped by the next field update (``w:fldSimple``).
_CLOSED_CONTAINERS: dict[str, str] = {
    qn("w:del"): "hidden-content",
    qn("w:moveFrom"): "hidden-content",
    qn("w:fldSimple"): "inside-field",
}

#: Control characters Word rejects, with the same exemptions as
#: ``utils/text_safety.py``: tab, line feed and carriage return have a meaning
#: in Word's text model, every other C0 byte corrupts Find/Replace.
_ALLOWED_CONTROLS = frozenset({"\t", "\n", "\r"})


# --------------------------------------------------------------------------
# Value types
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Pieces:
    """What a range resolves to, once the paragraph has been split on it.

    Attributes:
        paragraph: the ``w:p`` the range belongs to.
        start: first covered offset, in the paragraph's visible text.
        end: offset just past the range.
        operation: the intent the range was resolved for.
        runs: the ``w:r`` elements the range covers *entirely*, in document
            order.  Formatting applies to these; deleting removes these.  They
            need not share a parent: a range may span a hyperlink boundary.
        markers: the zero-width markers strictly inside the range --
            ``w:bookmarkStart``, ``w:commentRangeEnd``, ``w:permStart``, ...
            A deletion moves them to :attr:`start`; it never drops one.
        segments: every covered segment, in document order, markers and
            zero-width objects included.
    """

    paragraph: etree._Element
    start: int
    end: int
    operation: Operation
    runs: tuple[etree._Element, ...]
    markers: tuple[etree._Element, ...]
    segments: tuple[Segment, ...]

    @property
    def is_empty(self) -> bool:
        """True when the range is an insertion point rather than a span."""
        return self.start == self.end

    @property
    def text(self) -> str:
        """The visible text the range covers."""
        return "".join(segment.text for segment in self.segments)


@dataclass(frozen=True)
class _RunView:
    """One ``w:r`` and the span of visible text it contributes."""

    run: etree._Element
    segments: tuple[Segment, ...]
    start: int
    end: int

    @property
    def width(self) -> int:
        return self.end - self.start


# --------------------------------------------------------------------------
# Reading the paragraph
# --------------------------------------------------------------------------


def _paragraph_element(paragraph: object) -> etree._Element:
    """Coerce `paragraph` to a live ``w:p`` element.

    Accepts the element itself or a python-docx ``Paragraph``, the same two
    forms :mod:`word_document_server.engine.textmodel` accepts, so a caller can
    pass what it already holds.

    Raises:
        TypeError: if `paragraph` is neither.
    """
    element = paragraph
    if not isinstance(element, etree._Element):
        for attribute in ("_p", "_element"):
            candidate = getattr(element, attribute, None)
            if isinstance(candidate, etree._Element):
                element = candidate
                break
    if not isinstance(element, etree._Element):
        raise TypeError(
            f"expected a w:p element or a python-docx paragraph, got {type(paragraph).__name__}"
        )
    if element.tag != _W_P:
        raise TypeError(f"expected a w:p element, got {element.tag!r}")
    return element


def _run_views(paragraph: etree._Element) -> list[_RunView]:
    """Group the paragraph's segments by the run that produces them.

    Runs that hold no content at all (a bare ``w:rPr``) produce no segment and
    are absent: nothing addresses them by offset, and nothing here removes them.
    """
    views: list[_RunView] = []
    current: etree._Element | None = None
    bucket: list[Segment] = []
    for segment in segments(paragraph):
        run = segment.run
        if run is None:
            continue
        if current is not None and run is not current:
            views.append(_view(current, bucket))
            bucket = []
        current = run
        bucket.append(segment)
    if current is not None:
        views.append(_view(current, bucket))
    return views


def _view(run: etree._Element, bucket: list[Segment]) -> _RunView:
    return _RunView(run, tuple(bucket), bucket[0].start, bucket[-1].end)


def _covered(start: int, end: int, segment: Segment) -> bool:
    """Whether the range ``[start, end)`` covers `segment`.

    Characters use the half-open rule; zero-width content must sit strictly
    inside, so a boundary image or bookmark belongs to neither side.
    """
    if segment.end > segment.start:
        return start <= segment.start and segment.end <= end
    return start < segment.start < end


# --------------------------------------------------------------------------
# Splitting
# --------------------------------------------------------------------------


def _plan_split(view: _RunView, offset: int, zero_width: _Side) -> tuple[int, int | None]:
    """Decide where a run is cut.

    Returns the index of the first segment that moves to the right part and,
    when the cut falls inside a segment, the character offset inside it.
    """
    position = 0
    for index, segment in enumerate(view.segments):
        width = segment.end - segment.start
        if width == 0:
            if position > offset or (position == offset and zero_width == "right"):
                return index, None
            continue
        if position >= offset:
            return index, None
        if position + width <= offset:
            position += width
            continue
        return index, offset - position
    return len(view.segments), None


def _split_view(
    view: _RunView, offset: int, zero_width: _Side
) -> tuple[etree._Element | None, etree._Element | None]:
    """Cut `view`'s run at a local `offset`, returning its two parts."""
    run = view.run
    parent = run.getparent()
    if parent is None:
        raise UnsupportedRange("unsplittable-range", "cannot split a run that has no parent")

    index, inside = _plan_split(view, offset, zero_width)
    if inside is None:
        moving = [segment.element for segment in view.segments[index:]]
        if not moving:
            return run, None
        if index == 0:
            return None, run
        carrier: etree._Element | None = None
        cut = 0
    else:
        carrier = view.segments[index].element
        if carrier.tag != _W_T:
            raise UnsupportedRange(
                "unsplittable-range",
                f"cannot cut <{etree.QName(carrier).localname}> in the middle",
            )
        cut = inside
        moving = [segment.element for segment in view.segments[index + 1 :]]

    for element in moving:
        if element.getparent() is not run:
            raise UnsupportedRange(
                "unsplittable-range",
                "a run holds content that is not one of its own children",
            )

    right = etree.SubElement(parent, run.tag)
    for key, value in run.attrib.items():
        right.set(key, value)
    properties = run.find(_W_RPR)
    if properties is not None:
        right.append(copy.deepcopy(properties))
    if carrier is not None:
        full = carrier.text or ""
        carrier.text = full[:cut]
        carrier.set(_XML_SPACE, "preserve")
        tail = etree.SubElement(right, _W_T)
        for key, value in carrier.attrib.items():
            tail.set(key, value)
        tail.text = full[cut:]
        tail.set(_XML_SPACE, "preserve")
    for element in moving:
        right.append(element)
    # The index is read now, after every mutation above, never before.
    parent.insert(parent.index(run) + 1, right)
    return run, right


def split_run(
    run: etree._Element, offset: int, *, zero_width: _Side = "right"
) -> tuple[etree._Element | None, etree._Element | None]:
    """Cut `run` at `offset` characters into it, and return ``(left, right)``.

    `offset` counts the characters `run` itself contributes to the paragraph's
    visible text, so ``0`` is the run's start and its width is its end.  The
    left part is the original element, mutated in place and left where it was;
    the right part is a new sibling inserted just after it, carrying a copy of
    ``w:rPr``, the same run attributes (``w:rsidR``, ...) and the rest of the
    children in their original order.  Text ends up with
    ``xml:space="preserve"`` on both sides, since a cut can expose a leading or
    trailing space.

    The function returns ``(None, run)`` or ``(run, None)`` and leaves the tree
    untouched when everything falls on one side, which makes it idempotent at
    the bounds: splitting twice at the same offset splits once.  Content that
    carries no character and sits exactly at `offset` -- an image, a bookmark, a
    field character -- goes to the part named by `zero_width`; that is the only
    ambiguity a character offset cannot settle on its own.

    Raises:
        TypeError: if `run` is not a ``w:r`` inside a ``w:p``.
        LocatorError: ``out-of-range`` if `offset` is negative or past the run.
        UnsupportedRange: ``unsplittable-range`` if the cut falls inside content
            that is not a ``w:t``.
    """
    if not isinstance(run, etree._Element) or run.tag != _W_R:
        raise TypeError(f"expected a w:r element, got {type(run).__name__}")
    paragraph = next(iter(run.iterancestors(_W_P)), None)
    if paragraph is None:
        raise TypeError("expected a w:r inside a w:p")
    view = next((item for item in _run_views(paragraph) if item.run is run), None)
    width = 0 if view is None else view.width
    if offset < 0 or offset > width:
        raise LocatorError(
            "out-of-range", f"offset {offset} is outside the run's 0..{width} characters"
        )
    if view is None:
        return None, run
    return _split_view(view, offset, zero_width)


def _split_paragraph_at(paragraph: etree._Element, offset: int, zero_width: _Side) -> None:
    """Cut every run that spans `offset` so that no run straddles it."""
    for view in _run_views(paragraph):
        if view.start <= offset <= view.end:
            _split_view(view, offset - view.start, zero_width)


# --------------------------------------------------------------------------
# Resolving a range
# --------------------------------------------------------------------------


def _check_fields(paragraph: etree._Element, start: int, end: int, operation: Operation) -> None:
    """Refuse a range that would cut a field, or swallow one destructively."""
    for entry in fields(paragraph):
        if start == end:
            if entry.start < start < entry.end:
                raise UnsupportedRange(
                    "inside-field",
                    f"offset {start} is inside the field '{entry.instr.strip()}'",
                )
            continue
        if end <= entry.start or entry.end <= start:
            continue
        if start <= entry.start and entry.end <= end:
            if operation in _DESTRUCTIVE:
                raise UnsupportedRange(
                    "contains-field",
                    f"{operation} would remove the field '{entry.instr.strip()}' "
                    f"spanning {entry.start}..{entry.end}",
                )
            continue
        raise UnsupportedRange(
            "crosses-field",
            f"range {start}..{end} cuts the field '{entry.instr.strip()}' "
            f"spanning {entry.start}..{entry.end}",
        )


def _check_content(
    found: list[Segment], start: int, end: int, operation: Operation
) -> None:
    """Refuse a range that strictly contains hidden or opaque content."""
    for segment in found:
        if segment.end > segment.start or not (start < segment.start < end):
            continue
        name = etree.QName(segment.element).localname
        if segment.kind == "hidden":
            raise UnsupportedRange(
                "hidden-content",
                f"range {start}..{end} contains deleted or moved-away content "
                f"(<{name}> at {segment.start})",
            )
        if segment.kind == "opaque" and operation in _DESTRUCTIVE:
            raise UnsupportedRange(
                "opaque-content",
                f"{operation} would remove <{name}> at offset {segment.start}, "
                "an object this layer cannot rewrite",
            )


def resolve(
    paragraph: object, start: int, end: int, *, operation: Operation = "read"
) -> Pieces:
    """Split `paragraph` so that ``[start, end)`` lines up on whole runs.

    Splitting never changes the visible text, the run properties or anything
    outside the paragraph: it only makes the boundaries addressable.  The
    paragraph is therefore left split even when a later check refuses -- a
    harmless, text-preserving trace.

    Args:
        paragraph: a ``w:p`` element or a python-docx paragraph.
        start: first covered offset in the visible text.
        end: offset just past the range; equal to `start` for an insertion
            point, which only ``read`` and ``insert`` accept.
        operation: what the caller intends to do; see :data:`Operation`.

    Returns:
        The :class:`Pieces` the range resolves to.

    Raises:
        LocatorError: ``out-of-range`` if the offsets are not ``0 <= start <=
            end <= len(visible_text(paragraph))``.
        UnsupportedRange: see the module docstring for the reasons.
    """
    element = _paragraph_element(paragraph)
    found = segments(element)
    length = sum(len(segment.text) for segment in found)
    if start < 0 or end < start or end > length:
        raise LocatorError(
            "out-of-range",
            f"range {start}..{end} is outside the paragraph's 0..{length} visible text",
        )
    if start == end and operation in _DESTRUCTIVE:
        raise UnsupportedRange("empty-range", f"{operation} needs a non-empty range")

    _check_fields(element, start, end, operation)
    _check_content(found, start, end, operation)

    if start == end:
        # One cut only: a second one with the opposite side policy would break
        # runs apart around the insertion point for nothing.
        _split_paragraph_at(element, start, "right")
    else:
        _split_paragraph_at(element, end, "right")
        _split_paragraph_at(element, start, "left")

    found = segments(element)
    covered = [segment for segment in found if _covered(start, end, segment)]
    _verify_alignment(found, start, end)

    runs: list[etree._Element] = []
    for view in _run_views(element):
        states = {_covered(start, end, segment) for segment in view.segments}
        if states == {True}:
            runs.append(view.run)
        elif True in states:
            raise UnsupportedRange(
                "unsplittable-range",
                f"range {start}..{end} covers part of a run that could not be split",
            )
    return Pieces(
        paragraph=element,
        start=start,
        end=end,
        operation=operation,
        runs=tuple(runs),
        markers=tuple(segment.element for segment in covered if segment.kind == "marker"),
        segments=tuple(covered),
    )


def _verify_alignment(found: list[Segment], start: int, end: int) -> None:
    """Fail loudly if the splits did not align the range on whole segments."""
    for segment in found:
        if segment.end == segment.start:
            continue
        overlaps = segment.start < end and start < segment.end
        if not overlaps:
            continue
        if not _covered(start, end, segment):
            raise UnsupportedRange(
                "unsplittable-range",
                f"range {start}..{end} cuts <{etree.QName(segment.element).localname}> "
                f"spanning {segment.start}..{segment.end}",
            )
        if segment.run is None:
            raise UnsupportedRange(
                "unsplittable-range",
                f"text at offset {segment.start} is not inside a run",
            )


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------


def _unsafe(character: str) -> bool:
    """Whether `character` cannot be stored in a run."""
    code = ord(character)
    if character in _ALLOWED_CONTROLS:
        return False
    if code < 0x20:  # C0 controls: Word's Find/Replace chokes on them
        return True
    if 0xD800 <= code <= 0xDFFF:  # lone surrogate, forbidden by XML 1.0
        return True
    return code in (0xFFFE, 0xFFFF)


def _checked_text(text: str) -> str:
    """Return `text` with its line endings normalised, or refuse it.

    ``\\r\\n`` and a bare ``\\r`` become ``\\n``: a run cannot hold a paragraph
    mark, and a line break is what Word shows for either.

    Raises:
        InvalidText: if `text` holds a control character Word rejects or a code
            point XML 1.0 forbids.
    """
    if not isinstance(text, str):
        raise TypeError(f"expected a string, got {type(text).__name__}")
    bad = sorted({f"U+{ord(character):04X}" for character in text if _unsafe(character)})
    if bad:
        raise InvalidText(
            f"text contains characters a run cannot hold: {', '.join(bad)}. "
            "Control bytes corrupt Word's Find/Replace and have caused full-document "
            "data loss; use \\t for a tab and \\n for a line break."
        )
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _chunks(text: str):
    """Split `text` into literal pieces and the single characters that are not."""
    buffer: list[str] = []
    for character in text:
        if character in ("\n", "\t"):
            if buffer:
                yield "".join(buffer)
                buffer = []
            yield character
        else:
            buffer.append(character)
    if buffer:
        yield "".join(buffer)


def _make_run(
    parent: etree._Element, index: int, properties: etree._Element | None, text: str
) -> etree._Element:
    """Build a run holding `text` and insert it at `index` in `parent`."""
    run = etree.SubElement(parent, _W_R)
    if properties is not None:
        run.append(copy.deepcopy(properties))
    for chunk in _chunks(text):
        if chunk == "\n":
            etree.SubElement(run, _W_BR)
        elif chunk == "\t":
            etree.SubElement(run, _W_TAB)
        else:
            carrier = etree.SubElement(run, _W_T)
            carrier.text = chunk
            carrier.set(_XML_SPACE, "preserve")
    parent.insert(index, run)
    return run


def _remove(pieces: Pieces) -> tuple[etree._Element, int]:
    """Remove the covered runs, keeping the markers at the range's start.

    Returns the parent and the index the removed content occupied, so that a
    replacement can be written exactly where it was.
    """
    if not pieces.runs:
        raise UnsupportedRange(
            "unsplittable-range",
            f"range {pieces.start}..{pieces.end} covers no whole run",
        )
    first = pieces.runs[0]
    parent = first.getparent()
    for marker in pieces.markers:
        holder = marker.getparent()
        if holder is not None:
            holder.remove(marker)
    # Read the index only now: detaching the markers may have shifted it.
    index = parent.index(first)
    for position, marker in enumerate(pieces.markers):
        parent.insert(index + position, marker)
    index += len(pieces.markers)
    for run in pieces.runs:
        holder = run.getparent()
        if holder is not None:
            holder.remove(run)
    return parent, index


def _escape_closed_containers(
    paragraph: etree._Element, parent: etree._Element, index: int
) -> tuple[etree._Element, int] | str:
    """Walk an insertion point out of a container it must not be written into.

    An insertion at the very start or the very end of a ``w:fldSimple``, a
    ``w:del`` or a ``w:moveFrom`` is at the same place in the text just outside
    it, so it is moved there.  Anywhere else inside one, the point is rejected:
    the refusal reason is returned instead of a position, and the caller tries
    its other candidate.
    """
    while parent is not paragraph and parent.tag in _CLOSED_CONTAINERS:
        grandparent = parent.getparent()
        if grandparent is None:
            break
        if index == 0:
            index = grandparent.index(parent)
        elif index >= len(parent):
            index = grandparent.index(parent) + 1
        else:
            return _CLOSED_CONTAINERS[parent.tag]
        parent = grandparent
    return parent, index


def _field_marks(
    paragraph: etree._Element, parent: etree._Element, index: int
) -> tuple[list[tuple[int, str | None, etree._Element]], int]:
    """The paragraph's ``w:fldChar`` marks, and where an insertion point sits.

    Positions come from a single document-order walk, so a mark and the point
    can be compared even when they live in different containers.
    """
    order = {element: position for position, element in enumerate(paragraph.iter())}
    if index < len(parent):
        boundary = order[parent[index]]
    else:
        boundary = order[list(parent.iter())[-1]] + 1
    marks = [
        (order[element], element.get(_W_FLD_CHAR_TYPE), element)
        for element in paragraph.iter(_W_FLD_CHAR)
    ]
    return marks, boundary


def _leave_open_field(
    paragraph: etree._Element, parent: etree._Element, index: int, forward: bool
) -> tuple[etree._Element, int] | None:
    """Move an insertion point out of the complex field it would fall into.

    ``w:fldChar`` marks are a flat begin/end sequence, so being "inside" a
    complex field is not something an ancestor chain can tell: it has to be
    counted along the document order of the paragraph.  A point with a non-zero
    depth sits in a field's instruction or in its cached result, both of which
    Word rewrites on the next update -- text written there is lost.

    The point is therefore pushed past the field's ``end`` (or back before its
    ``begin``).  Only zero-width content can separate a boundary offset from
    those marks -- characters in between would have put the offset *inside* the
    field span, which :func:`resolve` refuses -- so the move keeps the text
    position it was asked for.  ``None`` means the field never closes inside
    this paragraph and there is nowhere safe to write.
    """
    marks, boundary = _field_marks(paragraph, parent, index)
    depth = 0
    for position, kind, _ in marks:
        if position >= boundary:
            break
        if kind == "begin":
            depth += 1
        elif kind == "end" and depth > 0:
            depth -= 1
    if depth == 0:
        return parent, index

    if forward:
        for position, kind, element in marks:
            if position < boundary:
                continue
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth -= 1
                if depth == 0:
                    return _beside(element, after=True)
    else:
        for position, kind, element in reversed(marks):
            if position >= boundary:
                continue
            if kind == "end":
                depth += 1
            elif kind == "begin":
                depth -= 1
                if depth == 0:
                    return _beside(element, after=False)
    return None


def _beside(element: etree._Element, *, after: bool) -> tuple[etree._Element, int]:
    """The position just after (or before) `element`, run wrapper included."""
    holder = element
    parent = holder.getparent()
    if parent is not None and parent.tag == _W_R:
        holder = parent
        parent = holder.getparent()
    return parent, parent.index(holder) + (1 if after else 0)


def _insertion_point(
    paragraph: etree._Element, offset: int, prefer: _Side
) -> tuple[etree._Element, int, etree._Element | None]:
    """Locate where an insertion goes, and which run lends its properties.

    The new content sits next to the run named by `prefer`: after the run that
    ends at `offset`, or before the run that starts there.  That is also how it
    inherits its context -- inserting on the left of a hyperlink boundary
    extends the link, inserting on the right does not.  A candidate that would
    write inside a field or inside deleted content is skipped in favour of the
    other side; when neither side is writable, the caller is refused rather than
    handed a position where Word would silently drop the text.
    """
    views = _run_views(paragraph)
    left = None
    right = None
    for view in views:
        if view.end == offset and view.start < offset:
            left = view
        if right is None and view.start == offset and view.end > offset:
            right = view
    candidates: list[tuple[_RunView | None, etree._Element, int]] = []
    order = (left, right) if prefer == "left" else (right, left)
    for view in order:
        if view is None:
            continue
        parent = view.run.getparent()
        index = parent.index(view.run) + (1 if view is left else 0)
        candidates.append((view, parent, index))
    if not candidates:
        # No run carries a character next to `offset`: the paragraph is empty,
        # or holds nothing but zero-width content.
        for view in views:
            if view.start == view.end == offset:
                parent = view.run.getparent()
                candidates.append((None, parent, parent.index(view.run)))
                break
        else:
            candidates.append((None, paragraph, len(paragraph)))

    refusals: list[str] = []
    for view, parent, index in candidates:
        escaped = _escape_closed_containers(paragraph, parent, index)
        if isinstance(escaped, str):
            refusals.append(escaped)
            continue
        point = _leave_open_field(paragraph, *escaped, forward=view is not right)
        if point is None:
            refusals.append("inside-field")
            continue
        escaped = _escape_closed_containers(paragraph, *point)
        if isinstance(escaped, str):
            refusals.append(escaped)
            continue
        return (*escaped, None if view is None else view.run)
    raise UnsupportedRange(
        refusals[0],
        f"offset {offset} has no writable position: every candidate falls inside "
        "a field or inside deleted content",
    )


def _properties_for(
    rpr_from: _Side | etree._Element, anchor: etree._Element | None
) -> etree._Element | None:
    """Return the ``w:rPr`` the inserted run must carry, or ``None``."""
    if isinstance(rpr_from, etree._Element):
        if rpr_from.tag == _W_RPR:
            return rpr_from
        if rpr_from.tag == _W_R:
            return rpr_from.find(_W_RPR)
        raise TypeError(
            f"rpr_from must be 'left', 'right', a w:rPr or a w:r, got {rpr_from.tag!r}"
        )
    if rpr_from not in ("left", "right"):
        raise ValueError(f"rpr_from must be 'left', 'right' or an element, got {rpr_from!r}")
    return None if anchor is None else anchor.find(_W_RPR)


def delete_range(paragraph: object, start: int, end: int) -> str:
    """Delete ``[start, end)`` from `paragraph` and return the text removed.

    Runs the range covers whole are removed; runs it covers in part are split
    first.  Markers inside the range are moved to `start` rather than deleted,
    and containers left empty (a hyperlink, a tracked insertion, a content
    control) are kept, so no relationship and no revision disappears with the
    text.

    Raises:
        LocatorError: ``out-of-range`` for impossible offsets.
        UnsupportedRange: see the module docstring for the reasons.
    """
    pieces = resolve(paragraph, start, end, operation="delete")
    removed = pieces.text
    _remove(pieces)
    return removed


def insert_text(
    paragraph: object,
    offset: int,
    text: str,
    *,
    rpr_from: _Side | etree._Element = "left",
) -> tuple[etree._Element, ...]:
    """Insert `text` at `offset` and return the runs created.

    ``\\n`` becomes a ``w:br`` and ``\\t`` a ``w:tab``; ``\\r\\n`` and ``\\r``
    are normalised to ``\\n``.  Inserting an empty string is a no-op that
    changes nothing and returns ``()``.

    `rpr_from` picks both the formatting and the context of the new text:
    ``"left"`` copies the ``w:rPr`` of the run ending at `offset` and writes
    next to it, ``"right"`` does the same with the run starting there, and an
    explicit ``w:rPr`` (or a ``w:r`` to take it from) overrides the formatting
    while keeping the left-hand placement.  When the preferred side has no run,
    the other one is used.

    Raises:
        InvalidText: if `text` holds a control character Word rejects.
        LocatorError: ``out-of-range`` if `offset` is outside the visible text.
        UnsupportedRange: ``inside-field`` inside a field, ``hidden-content``
            inside deleted or moved-away content.
    """
    element = _paragraph_element(paragraph)
    checked = _checked_text(text)
    if not checked:
        return ()
    # For its checks and its split; an insertion point covers no piece.
    resolve(element, offset, offset, operation="insert")
    prefer: _Side = rpr_from if rpr_from in ("left", "right") else "left"
    parent, index, anchor = _insertion_point(element, offset, prefer)
    properties = _properties_for(rpr_from, anchor)
    return (_make_run(parent, index, properties, checked),)


def replace_range(
    paragraph: object, start: int, end: int, text: str
) -> tuple[etree._Element, ...]:
    """Replace ``[start, end)`` with `text` and return the runs created.

    The replacement inherits the ``w:rPr`` of the first run of the range and
    lands exactly where that run was, after any marker the deletion preserved --
    so a bookmark or a comment that framed the old text still frames the new
    one.  An empty `text` makes this a deletion, and returns ``()``.

    Raises:
        InvalidText: if `text` holds a control character Word rejects.
        LocatorError: ``out-of-range`` for impossible offsets.
        UnsupportedRange: see the module docstring for the reasons.
    """
    element = _paragraph_element(paragraph)
    checked = _checked_text(text)
    pieces = resolve(element, start, end, operation="replace")
    properties = pieces.runs[0].find(_W_RPR) if pieces.runs else None
    if properties is not None:
        properties = copy.deepcopy(properties)
    parent, index = _remove(pieces)
    if not checked:
        return ()
    return (_make_run(parent, index, properties, checked),)
