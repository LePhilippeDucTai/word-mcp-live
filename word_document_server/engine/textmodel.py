"""The text model: what a paragraph shows, and where each character lives.

Every layer above this one addresses text by character offset -- ``find`` returns
offsets, ``ranges`` splits runs at offsets, the V2 locators quote expected text.
This module is the single place that decides which characters exist and in which
order, so that one offset means the same thing everywhere.

The reading is *not* ``paragraph.text``.  python-docx only walks ``w:r`` children
of ``w:p``, so it drops every run wrapped in ``w:hyperlink``, ``w:ins``, ``w:sdt``
or ``w:fldSimple`` -- on a document with tracked changes it returns text that is
neither the original nor the revised version.  It is not
``utils.document_utils.get_effective_text`` either: iterating over every ``w:t``
in the subtree also picks up the tab *stops* declared in ``w:pPr/w:tabs`` (which
are not content) and cannot say where a character sits.

Policy
------
The walk classifies every element it meets into exactly one bucket:

skipped
    Property containers -- ``w:pPr``, ``w:rPr``, ``w:sdtPr``, ``w:pPrChange``,
    ``w:rPrChange``, ``w:fldData``, ... -- carry formatting, not content.  They
    are not descended into and produce no segment, so the ``w:tab`` of a tab stop
    and the ``w:del`` of a deleted paragraph mark never reach the text.
traversed
    ``w:r``, ``w:hyperlink``, ``w:ins``, ``w:moveTo``, ``w:sdt``/``w:sdtContent``,
    ``w:smartTag``, ``w:fldSimple``, ``w:customXml``, ``w:dir``, ``w:bdo`` are
    wrappers: their content belongs to the flow.  ``w:del`` and ``w:moveFrom``
    are traversed too, but everything textual under them becomes ``hidden``.
text
    ``w:t``; ``w:tab`` and ``w:ptab`` -> ``\\t``; ``w:br`` and ``w:cr`` -> ``\\n``;
    ``w:noBreakHyphen`` -> ``\\u2011``; ``w:softHyphen`` -> ``\\u00ad``; ``w:sym``
    -> the code point of its ``w:char`` attribute.
hidden
    ``w:delText`` and any textual element under ``w:del`` or ``w:moveFrom``:
    characters the document stores but does not show.  Zero width.  A marker or
    an opaque object under a deletion keeps its own kind -- a bookmark inside a
    deletion is still a bookmark the layers above must preserve.
marker
    ``bookmarkStart``/``End``, ``commentRangeStart``/``End``,
    ``permStart``/``End``, ``proofErr``, ``moveFrom``/``moveToRangeStart``/``End``.
    Zero width, and the layers above must never drop one.
opaque
    Everything else, starting with the objects the plan names: ``w:drawing``,
    ``w:pict``, ``w:object``, ``w:footnoteReference``, ``w:endnoteReference``,
    ``w:commentReference``, ``w:fldChar``, ``w:instrText``, ``w:delInstrText``.
    Zero width and never descended into.

Opaque is deliberately the *default*: an element this module has not been taught
about contributes nothing rather than silently leaking the text of a construct
whose semantics are unknown (``w:ruby`` shows one of its two texts,
``mc:AlternateContent`` shows one of its branches).  Teaching the engine a new
element means adding a row to one of the tables below, and a wrong row shows up
as a wrong expected text in ``tests/engine/test_textmodel.py``.

Fields
------
``w:fldSimple`` is atomic by construction.  A complex field is a flat sequence of
``w:fldChar`` markers -- ``begin``, optional ``separate``, ``end`` -- that can
nest, so :func:`fields` runs a stack machine over them.  The cached result of a
field is ordinary ``w:t`` and therefore *visible*: this module reports what the
package stores, it does not evaluate fields.  In a nested field the inner cached
result sits inside the outer instruction and is visible too; :func:`fields`
reports both spans so that ``ranges`` can refuse to cut through either.

Reading only
------------
Nothing here mutates the tree, and nothing is cached: the elements handed out are
live, and a caller that edits them simply calls back to get the new flow.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass
from typing import Literal

from lxml import etree

from word_document_server.engine.errors import LocatorError
from word_document_server.engine.xmlns import qn

__all__ = [
    "SEGMENT_KINDS",
    "Field",
    "Segment",
    "SegmentKind",
    "fields",
    "position",
    "segments",
    "visible_text",
]

#: What a :class:`Segment` can be.  ``text``, ``tab`` and ``break`` carry
#: characters; ``opaque``, ``marker`` and ``hidden`` are zero width.
SegmentKind = Literal["text", "tab", "break", "opaque", "marker", "hidden"]

#: The runtime form of :data:`SegmentKind`, for assertions and validation.
SEGMENT_KINDS: frozenset[str] = frozenset({"text", "tab", "break", "opaque", "marker", "hidden"})

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_T = qn("w:t")
_W_DEL_TEXT = qn("w:delText")
_W_SYM = qn("w:sym")
_W_CHAR = qn("w:char")
_W_FLD_CHAR = qn("w:fldChar")
_W_FLD_CHAR_TYPE = qn("w:fldCharType")
_W_FLD_SIMPLE = qn("w:fldSimple")
_W_INSTR = qn("w:instr")
_W_INSTR_TEXT = qn("w:instrText")
_W_DEL_INSTR_TEXT = qn("w:delInstrText")

# Property containers: formatting, never content.  ``w:pPr`` matters most -- it
# holds ``w:tabs/w:tab`` (tab *stops*) and the ``w:rPr`` of the paragraph mark,
# which may carry ``w:ins``/``w:del`` for a tracked paragraph-mark revision.
_SKIPPED = frozenset(
    qn(tag)
    for tag in (
        "w:pPr",
        "w:rPr",
        "w:sdtPr",
        "w:sdtEndPr",
        "w:smartTagPr",
        "w:customXmlPr",
        "w:tblPr",
        "w:tblPrEx",
        "w:trPr",
        "w:tcPr",
        "w:pPrChange",
        "w:rPrChange",
        "w:sectPr",
        "w:fldData",
    )
)

# Wrappers whose content belongs to the flow.  ``w:fldSimple`` is a container too
# but is handled on its own, because it also opens a field.
_CONTAINERS = frozenset(
    qn(tag)
    for tag in (
        "w:r",
        "w:hyperlink",
        "w:ins",
        "w:moveTo",
        "w:del",
        "w:moveFrom",
        "w:sdt",
        "w:sdtContent",
        "w:smartTag",
        "w:customXml",
        "w:dir",
        "w:bdo",
    )
)

#: Containers whose textual content exists but is not shown.
_HIDING = frozenset({qn("w:del"), qn("w:moveFrom")})

#: Zero-width anchors that the layers above must preserve verbatim.
_MARKERS = frozenset(
    qn(tag)
    for tag in (
        "w:bookmarkStart",
        "w:bookmarkEnd",
        "w:commentRangeStart",
        "w:commentRangeEnd",
        "w:permStart",
        "w:permEnd",
        "w:proofErr",
        "w:moveFromRangeStart",
        "w:moveFromRangeEnd",
        "w:moveToRangeStart",
        "w:moveToRangeEnd",
    )
)

#: Elements that carry characters, and the kind they produce when visible.
_CARRIERS: dict[str, SegmentKind] = {
    _W_T: "text",
    _W_DEL_TEXT: "hidden",
    qn("w:tab"): "tab",
    qn("w:ptab"): "tab",
    qn("w:br"): "break",
    qn("w:cr"): "break",
    _W_SYM: "text",
    qn("w:noBreakHyphen"): "text",
    qn("w:softHyphen"): "text",
}

#: Carriers whose character does not depend on the element's content.
_FIXED_CHARS = {
    qn("w:tab"): "\t",
    qn("w:ptab"): "\t",
    qn("w:br"): "\n",
    qn("w:cr"): "\n",
    # Spelled with chr(): a literal soft hyphen is invisible in a source file.
    qn("w:noBreakHyphen"): chr(0x2011),
    qn("w:softHyphen"): chr(0x00AD),
}

#: Stand-in for a ``w:sym`` whose ``w:char`` is missing or unusable.
_REPLACEMENT = chr(0xFFFD)


def _sym_char(element: etree._Element) -> str:
    """Return the single character a ``w:sym`` stands for.

    The stored code point is returned as is: Word writes symbol-font glyphs in
    the private use area (``F0B7`` for a Symbol bullet), and translating them
    would require the font's own encoding, which the package does not carry.
    Preserving the code point keeps offsets and round-trips exact.  A ``w:sym``
    is always exactly one character wide, whatever its ``w:char`` says, so an
    unreadable code cannot shift the offsets of the rest of the paragraph.
    """
    raw = element.get(_W_CHAR)
    if raw is None:
        return _REPLACEMENT
    try:
        code = int(raw, 16)
    except ValueError:
        return _REPLACEMENT
    # Surrogates, control characters and out-of-range values have no usable
    # single-character form.
    if code < 0x20 or code > 0x10FFFF or 0xD800 <= code <= 0xDFFF:
        return _REPLACEMENT
    return chr(code)


def _carried_text(element: etree._Element) -> str:
    """Return the characters `element` carries, visible or not."""
    tag = element.tag
    if tag in (_W_T, _W_DEL_TEXT):
        return element.text or ""
    if tag == _W_SYM:
        return _sym_char(element)
    return _FIXED_CHARS.get(tag, "")


@dataclass(frozen=True)
class Segment:
    """One indivisible piece of a paragraph's flow.

    Segments tile the paragraph: they are in document order, the first starts at
    0, each one starts where the previous ends, and the last ends at
    ``len(visible_text(paragraph))``.  Joining their :attr:`text` reproduces the
    visible text exactly, and ``len(text) == end - start`` always holds -- a
    segment that shows nothing (``opaque``, ``marker``, ``hidden``) is an empty
    span sitting at one offset, not a gap.

    Attributes:
        kind: see :data:`SegmentKind`.
        element: the live element that produces the segment (``w:t``, ``w:tab``,
            ``w:drawing``, ``w:bookmarkStart``, ...).
        containers: the ancestors between the paragraph (excluded) and
            `element` (excluded), outermost first -- typically
            ``(w:hyperlink, w:r)``.  Tells a caller what the segment is wrapped
            in without walking back up the tree.
        start: offset of the first character, in the paragraph's visible text.
        end: offset just past the last character.
        text: the visible characters, empty for a zero-width segment.
    """

    kind: SegmentKind
    element: etree._Element
    containers: tuple[etree._Element, ...]
    start: int
    end: int
    text: str

    @property
    def run(self) -> etree._Element | None:
        """The innermost ``w:r`` this segment sits in, or ``None``.

        A marker is usually a direct child of the paragraph and has no run.
        """
        for candidate in reversed(self.containers):
            if candidate.tag == _W_R:
                return candidate
        return None

    @property
    def source_text(self) -> str:
        """The characters the element carries, whether or not they are shown.

        Equal to :attr:`text` for visible content and empty for opaque objects
        and markers.  For a ``hidden`` segment it is the suppressed text -- what
        a ``w:delText`` holds, or the ``\\t`` a ``w:tab`` inside a ``w:del``
        stands for -- which is how a caller reads deleted text without walking
        the tree itself.
        """
        if self.kind != "hidden":
            return self.text
        return _carried_text(self.element)


@dataclass(frozen=True)
class Field:
    """A field, reported as one atomic span of the paragraph.

    Attributes:
        instr: the instruction.  The ``w:instr`` attribute for a ``w:fldSimple``,
            the concatenated ``w:instrText`` of the field's own level for a
            complex field -- the instruction of a nested field belongs to that
            nested field, not to its parent.  ``w:delInstrText`` is excluded: a
            deleted instruction is not part of the current one.
        start: offset where the field opens, in the paragraph's visible text.
        end: offset where it closes.  The span covers the cached result, since
            that is the only part of a field that shows.
        atomic: always ``True``.  A field's markers, instruction and result are
            consistent only as a whole, so the layers above edit a field by
            replacing it, never by cutting into it.
    """

    instr: str
    start: int
    end: int
    atomic: bool = True


@dataclass
class _OpenField:
    """A complex field whose ``begin`` has been seen but not yet its ``end``."""

    order: int
    start: int
    instr: list[str] = dataclasses.field(default_factory=list)
    in_result: bool = False


@dataclass(frozen=True)
class _Flow:
    """The whole reading of one paragraph, computed in a single walk."""

    text: str
    segments: tuple[Segment, ...]
    fields: tuple[Field, ...]


class _Scanner:
    """Single-pass walk producing the segments, the text and the fields at once.

    They share one traversal on purpose: three independent walks could disagree
    about an offset, and an offset that means two different things is exactly the
    bug this layer exists to prevent.
    """

    def __init__(self) -> None:
        self._segments: list[Segment] = []
        self._chunks: list[str] = []
        self._offset = 0
        self._stack: list[_OpenField] = []
        self._closed: list[tuple[int, Field]] = []
        self._order = 0
        # Instruction text seen outside any field of this paragraph: it belongs
        # to a field that opened in an earlier paragraph.
        self._orphan_instr: list[str] = []

    # -- walking --------------------------------------------------------------

    def scan(self, paragraph: etree._Element) -> _Flow:
        """Walk `paragraph` and return its flow."""
        self._walk(paragraph, (), hidden=False)
        # A field may legitimately span paragraphs (a TOC covers a whole page):
        # close what is still open at the paragraph's end rather than dropping it.
        while self._stack:
            self._close(self._stack.pop(), self._offset)
        self._closed.sort(key=lambda item: item[0])
        return _Flow(
            text="".join(self._chunks),
            segments=tuple(self._segments),
            fields=tuple(entry for _, entry in self._closed),
        )

    def _walk(
        self,
        parent: etree._Element,
        containers: tuple[etree._Element, ...],
        *,
        hidden: bool,
    ) -> None:
        for child in parent:
            tag = child.tag
            # Comments and processing instructions have a callable tag.
            if not isinstance(tag, str) or tag in _SKIPPED:
                continue
            if tag in _MARKERS:
                self._emit("marker", child, containers)
                continue
            if tag == _W_FLD_CHAR:
                self._emit("opaque", child, containers)
                self._on_fld_char(child)
                continue
            if tag == _W_INSTR_TEXT:
                self._emit("opaque", child, containers)
                self._on_instr_text(child)
                continue
            if tag == _W_DEL_INSTR_TEXT:
                self._emit("opaque", child, containers)
                continue
            if tag == _W_FLD_SIMPLE:
                self._on_fld_simple(child, containers, hidden=hidden)
                continue
            if tag in _CONTAINERS:
                self._walk(
                    child,
                    (*containers, child),
                    hidden=hidden or tag in _HIDING,
                )
                continue
            carried = _CARRIERS.get(tag)
            if carried is not None:
                kind: SegmentKind = "hidden" if hidden else carried
                text = "" if kind == "hidden" else _carried_text(child)
                self._emit(kind, child, containers, text)
                continue
            self._emit("opaque", child, containers)

    def _emit(
        self,
        kind: SegmentKind,
        element: etree._Element,
        containers: tuple[etree._Element, ...],
        text: str = "",
    ) -> None:
        start = self._offset
        if text:
            self._chunks.append(text)
            self._offset += len(text)
        self._segments.append(Segment(kind, element, containers, start, self._offset, text))

    # -- fields ---------------------------------------------------------------

    def _on_fld_char(self, element: etree._Element) -> None:
        char_type = element.get(_W_FLD_CHAR_TYPE)
        if char_type == "begin":
            self._order += 1
            self._stack.append(_OpenField(order=self._order, start=self._offset))
        elif char_type == "separate":
            if self._stack:
                self._stack[-1].in_result = True
        elif char_type == "end":
            if self._stack:
                self._close(self._stack.pop(), self._offset)
            else:
                # The matching ``begin`` is in an earlier paragraph.  Order 0
                # sorts such a field ahead of every field opened here, which is
                # where it belongs: its span starts at the paragraph's start.
                orphan = _OpenField(order=0, start=0, instr=list(self._orphan_instr))
                self._orphan_instr.clear()
                self._close(orphan, self._offset)

    def _on_instr_text(self, element: etree._Element) -> None:
        text = element.text or ""
        if not self._stack:
            self._orphan_instr.append(text)
            return
        top = self._stack[-1]
        # Instruction text after ``separate`` is malformed; the result region is
        # not part of the instruction, so it is dropped rather than concatenated.
        if not top.in_result:
            top.instr.append(text)

    def _on_fld_simple(
        self,
        element: etree._Element,
        containers: tuple[etree._Element, ...],
        *,
        hidden: bool,
    ) -> None:
        self._order += 1
        order = self._order
        start = self._offset
        self._walk(element, (*containers, element), hidden=hidden)
        self._closed.append(
            (order, Field(instr=element.get(_W_INSTR) or "", start=start, end=self._offset))
        )

    def _close(self, open_field: _OpenField, end: int) -> None:
        self._closed.append(
            (
                open_field.order,
                Field(instr="".join(open_field.instr), start=open_field.start, end=end),
            )
        )


def _paragraph_element(paragraph: object) -> etree._Element:
    """Coerce `paragraph` to a live ``w:p`` element.

    Accepts the element itself or a python-docx ``Paragraph``, so callers on
    either side of the engine boundary can pass what they already hold.

    Raises:
        TypeError: if `paragraph` is neither, or is an element that is not
            ``w:p`` -- a programming error, not a document condition.
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


def _flow(paragraph: object) -> _Flow:
    return _Scanner().scan(_paragraph_element(paragraph))


def segments(paragraph: object) -> list[Segment]:
    """Return the segments of `paragraph`, in document order.

    See :class:`Segment` for the tiling contract.  The result is computed on each
    call from the live tree, so it reflects edits made since the last call.
    """
    return list(_flow(paragraph).segments)


def visible_text(paragraph: object) -> str:
    """Return the text `paragraph` shows.

    Insertions and the content of hyperlinks, content controls, smart tags and
    field results are included; deletions, moved-away text and field
    instructions are not.  Equal to ``"".join(s.text for s in segments(p))``.
    """
    return _flow(paragraph).text


def fields(paragraph: object) -> list[Field]:
    """Return the fields of `paragraph`, outermost first at equal offsets.

    Covers both forms: ``w:fldSimple`` and the ``w:fldChar`` state machine,
    nesting included.  A field that opens in an earlier paragraph is reported
    from offset 0, and one that closes in a later paragraph is reported up to the
    end of this one, so a caller never sees a field it can safely cut through.
    """
    return list(_flow(paragraph).fields)


def position(paragraph: object, offset: int) -> tuple[Segment, int]:
    """Locate `offset` in `paragraph`: the segment holding it and where inside.

    The returned pair satisfies ``segment.start + local_offset == offset``.  The
    answer is the first segment that *opens* at `offset` or spans it -- the first
    one with ``start <= offset < end``.  When no segment does, which happens at
    ``len(visible_text(paragraph))`` (the insertion point at the very end) and in
    a paragraph whose visible text is empty, the answer is the last segment
    ending at `offset`, so a trailing marker is preferred to a marker further
    left.  A caller that wants the other side of a boundary reads it from the
    segment it gets: ``ranges.insert_text`` chooses through its own
    ``rpr_from`` argument rather than through a second flavour of this function.

    Raises:
        LocatorError: ``out-of-range`` if `offset` is negative or past the end of
            the visible text; ``empty-paragraph`` if the paragraph has no content
            at all, since then there is no segment to anchor to.
    """
    flow = _flow(paragraph)
    length = len(flow.text)
    if offset < 0 or offset > length:
        raise LocatorError(
            "out-of-range",
            f"offset {offset} is outside the paragraph's 0..{length} visible text",
        )
    if not flow.segments:
        raise LocatorError("empty-paragraph", "paragraph has no content to position into")

    found: Segment | None = None
    for segment in flow.segments:
        if segment.start > offset:
            break
        if segment.start <= offset < segment.end:
            return segment, offset - segment.start
        if segment.end == offset:
            found = segment
    if found is None:  # pragma: no cover - segments tile [0, length], so this cannot happen
        found = flow.segments[0]
    return found, offset - found.start
