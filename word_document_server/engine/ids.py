"""Allocation of the identifiers WordprocessingML requires to be unique.

Word silently misbehaves when two annotations share an id: a second bookmark
reusing an id makes the first one unreachable, a revision reusing an id makes
"accept all" drop the wrong run.  Every function here therefore looks at the
*whole* package, not at the part being edited, and returns an id no existing
element uses.

The functions are pure: they read the package and return a number, they never
stamp it anywhere.  Calling one twice without editing anything in between
returns the same value -- the caller owns the id as soon as it writes it.

Four id spaces coexist and must not be mixed:

``w:id`` on annotations
    Shared by revisions (``w:ins``, ``w:del``, ``w:moveFrom``, ``w:moveTo``),
    bookmarks and comment anchors.  :func:`next_annotation_id`.
``w:id`` on ``w:comment``
    The comment's own identity, referenced by ``w:commentReference``.
    :func:`next_comment_id`.
``w:id`` on ``w:footnote``
    Referenced by ``w:footnoteReference``; ids below 1 are reserved for the
    separator placeholders.  :func:`next_footnote_id`.
``w14:paraId``
    A 32-bit non-zero value, written as 8 hexadecimal digits, unique across the
    document.  :func:`next_para_id`.
"""

from __future__ import annotations

from collections.abc import Iterable, Sequence

from lxml import etree

from word_document_server.engine.errors import IdExhausted
from word_document_server.engine.package import COMMENTS_CONTENT_TYPE, DocxPackage
from word_document_server.engine.xmlns import qn

__all__ = [
    "ANNOTATION_ID_TAGS",
    "MAX_DECIMAL_ID",
    "PARA_ID_MAX",
    "PARA_ID_MIN",
    "next_annotation_id",
    "next_comment_id",
    "next_footnote_id",
    "next_para_id",
]

#: ``w:id`` is ``ST_DecimalNumber``, a signed 32-bit integer.
MAX_DECIMAL_ID = 2**31 - 1

#: Word rejects a ``w14:paraId`` of ``00000000`` or above ``7FFFFFFF``.
PARA_ID_MIN = 0x00000001
PARA_ID_MAX = 0x7FFFFFFF

#: Elements sharing the annotation id space.  The ``…RangeStart`` elements of a
#: move are included beyond the four revision elements and the three anchors:
#: Word allocates their ids from the same counter, so ignoring them would let
#: :func:`next_annotation_id` hand back an id a move range already uses.
ANNOTATION_ID_TAGS: tuple[str, ...] = (
    "w:ins",
    "w:del",
    "w:moveFrom",
    "w:moveTo",
    "w:moveFromRangeStart",
    "w:moveToRangeStart",
    "w:bookmarkStart",
    "w:commentRangeStart",
    "w:commentReference",
)

_COMMENT_ID_TAGS: tuple[str, ...] = (
    "w:comment",
    "w:commentRangeStart",
    "w:commentRangeEnd",
    "w:commentReference",
)

_FOOTNOTE_ID_TAGS: tuple[str, ...] = ("w:footnote", "w:footnoteReference")

_W_ID = qn("w:id")


def _annotated_roots(pkg: DocxPackage) -> list[etree._Element]:
    """Every root that may carry annotation ids: the stories plus the comments.

    ``word/comments.xml`` is not a story (the engine never walks it as content)
    but its paragraphs carry bookmarks, revisions and ``w14:paraId`` values that
    live in the same id spaces as the body's.
    """
    roots = [root for _, root in pkg.stories()]
    comments = pkg.find_part("/word/comments.xml")
    if comments is not None and comments.content_type == COMMENTS_CONTENT_TYPE:
        roots.append(pkg.root_of(comments))
    return roots


def _max_attribute_int(
    roots: Iterable[etree._Element], tags: Sequence[str], attribute: str
) -> int | None:
    """Largest integer value of `attribute` over `tags`, or ``None`` if unused.

    Values that are not integers are ignored rather than raising: a malformed id
    written by another producer must not make the engine unable to allocate.
    """
    wanted = tuple(qn(tag) for tag in tags)
    best: int | None = None
    for root in roots:
        for element in root.iter(*wanted):
            raw = element.get(attribute)
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                continue
            if best is None or value > best:
                best = value
    return best


def _next_after(max_seen: int | None, floor: int, space: str) -> int:
    """Return the first free id at or above `floor`, given the largest in use."""
    candidate = floor if max_seen is None else max(max_seen + 1, floor)
    if candidate > MAX_DECIMAL_ID:
        raise IdExhausted(f"no {space} id left at or below {MAX_DECIMAL_ID}")
    return candidate


def next_annotation_id(pkg: DocxPackage) -> int:
    """Return an unused id for a revision, a bookmark or a comment anchor.

    Ids start at 1: ``0`` is the id Word gives the implicit ``_GoBack``
    bookmark, so handing it out would collide on most real documents.

    Raises:
        IdExhausted: if the 32-bit space is full.
    """
    max_seen = _max_attribute_int(_annotated_roots(pkg), ANNOTATION_ID_TAGS, _W_ID)
    return _next_after(max_seen, floor=1, space="annotation")


def next_comment_id(pkg: DocxPackage) -> int:
    """Return an unused ``w:comment`` id.

    Anchors (``w:commentRangeStart`` and friends) are scanned as well, so an id
    already referenced from a story is never handed out even when the comment it
    points at has been deleted from ``word/comments.xml``.

    Raises:
        IdExhausted: if the 32-bit space is full.
    """
    max_seen = _max_attribute_int(_annotated_roots(pkg), _COMMENT_ID_TAGS, _W_ID)
    return _next_after(max_seen, floor=1, space="comment")


def next_footnote_id(pkg: DocxPackage) -> int:
    """Return an unused ``w:footnote`` id.

    Ids start at 1: the separator and continuation-separator placeholders own
    the values below it (Word uses ``-1`` and ``0``, LibreOffice ``0`` and
    ``1``), and they are never referenced from a story.

    Raises:
        IdExhausted: if the 32-bit space is full.
    """
    max_seen = _max_attribute_int(_annotated_roots(pkg), _FOOTNOTE_ID_TAGS, _W_ID)
    return _next_after(max_seen, floor=1, space="footnote")


def next_para_id(pkg: DocxPackage) -> str:
    """Return an unused ``w14:paraId``, as 8 uppercase hexadecimal digits.

    Allocation walks up from the largest value in use and wraps around at
    ``7FFFFFFF``, so a document whose ids sit near the top of the space keeps
    getting valid ids instead of overflowing into the range Word rejects.

    Raises:
        IdExhausted: if every value of the space is already in use.
    """
    attribute = qn("w14:paraId")
    used: set[int] = set()
    for root in _annotated_roots(pkg):
        # Word writes w14:paraId on w:p and w:tr, but the walk stays untyped:
        # an id another producer put elsewhere still has to be avoided.
        for element in root.iter("*"):
            raw = element.get(attribute)
            if raw is None:
                continue
            try:
                used.add(int(raw, 16))
            except ValueError:
                continue

    candidate = max(used) + 1 if used else PARA_ID_MIN
    if candidate > PARA_ID_MAX or candidate < PARA_ID_MIN:
        candidate = PARA_ID_MIN
    for _ in range(PARA_ID_MAX - PARA_ID_MIN + 1):
        if candidate not in used:
            return f"{candidate:08X}"
        candidate = PARA_ID_MIN if candidate >= PARA_ID_MAX else candidate + 1
    raise IdExhausted("no free w14:paraId left in the package")
