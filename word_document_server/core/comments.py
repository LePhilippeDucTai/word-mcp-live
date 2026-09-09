"""Reading the comments of a Word document, through the OOXML engine.

A comment is stored twice: its text and metadata live in ``word/comments.xml``,
its *anchor* lives in the story it annotates, as a ``w:commentRangeStart`` /
``w:commentRangeEnd`` pair plus a ``w:commentReference`` mark, all three keyed
on the comment's ``w:id``.  Reading a comment therefore means joining the two,
which is what :func:`extract_all_comments` does: every comment is reported with
the story it is anchored in, the paragraph it starts in, and the text it
actually covers.

Addressing
----------
``story`` is the id :meth:`~word_document_server.engine.package.DocxPackage.stories`
gives the part -- ``"document"``, ``"header1"``, ``"footnotes"``, ... -- never
the ``"body"`` alias, which only exists as an input filter of
:func:`~word_document_server.engine.find.find`.  ``paragraph_index`` is the V2
index of the paragraph the anchor starts in, and is ``None`` when that
paragraph is out of the V2 space: inside a table cell (``in_table`` then says
so) or inside a text box.

Failures
--------
Nothing here is swallowed.  A package that cannot be opened, a comments part
that cannot be parsed, an anchor that does not resolve: each raises, and the
tool layer turns the exception into an error response.  The previous
implementation fell back to scanning ``str(element)`` for the substring
``"commentRangeStart"`` and reported ``"Comment detected but content not
accessible"``, which reads as data and is not.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from typing import Any

from lxml import etree

from word_document_server.engine.find import iter_paragraphs
# The V2 index space has exactly one definition (D-009); it lives in find.py and
# is reused here rather than restated, so the two can never drift apart.
from word_document_server.engine.find import _v2_index_map as v2_index_map
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import segments, visible_text
from word_document_server.engine.xmlns import qn

__all__ = [
    "Anchor",
    "anchors",
    "extract_all_comments",
    "filter_comments_by_author",
    "get_comments_for_paragraph",
]

COMMENTS_PARTNAME = "/word/comments.xml"

_W_ID = qn("w:id")
_W_COMMENT = qn("w:comment")
_W_RANGE_START = qn("w:commentRangeStart")
_W_RANGE_END = qn("w:commentRangeEnd")
_W_REFERENCE = qn("w:commentReference")
_W_TC = qn("w:tc")
_ANCHOR_TAGS = (_W_RANGE_START, _W_RANGE_END, _W_REFERENCE)


@dataclass(frozen=True)
class _Point:
    """One anchor mark, placed in its story and in its paragraph's text."""

    story: str
    order: int
    paragraph: etree._Element
    index: int | None
    offset: int
    in_table: bool


@dataclass(frozen=True)
class Anchor:
    """Where a comment is attached, and what it covers.

    Attributes:
        story: id of the story the anchor lives in, as
            :meth:`~word_document_server.engine.package.DocxPackage.stories`
            names it.
        paragraph_index: V2 index of the paragraph the anchor starts in, or
            ``None`` when that paragraph has no V2 index (table cell, text box).
        in_table: whether that paragraph sits in a table cell.
        text: the visible text the anchor covers, paragraphs joined by a
            newline.  Empty for a comment anchored on an empty range, or on a
            reference mark with no range around it.
    """

    story: str
    paragraph_index: int | None
    in_table: bool
    text: str


def _story_points(
    story: str, root: etree._Element
) -> tuple[dict[str, _Point], dict[str, _Point], dict[str, _Point], list[etree._Element]]:
    """Locate every comment mark of one story.

    Returns the range starts, the range ends and the reference marks, each
    keyed by comment id, plus the story's paragraphs in document order (the
    order :attr:`_Point.order` indexes).
    """
    paragraphs = iter_paragraphs(root)
    index_map = v2_index_map(paragraphs)
    starts: dict[str, _Point] = {}
    ends: dict[str, _Point] = {}
    references: dict[str, _Point] = {}
    by_tag = {_W_RANGE_START: starts, _W_RANGE_END: ends, _W_REFERENCE: references}

    for order, paragraph in enumerate(paragraphs):
        if next(paragraph.iter(*_ANCHOR_TAGS), None) is None:
            continue
        in_table = any(ancestor.tag == _W_TC for ancestor in paragraph.iterancestors())
        # Reading the marks off the segments rather than off the tree gives each
        # one its offset in the visible text for free, and stops at a nested
        # ``w:p``: a mark inside a text box belongs to the text box's paragraph,
        # which this loop visits on its own.
        for segment in segments(paragraph):
            bucket = by_tag.get(segment.element.tag)
            if bucket is None:
                continue
            comment_id = segment.element.get(_W_ID)
            if comment_id is None:
                continue
            # A malformed document may repeat a mark; the first one wins, which
            # is the one that opens the range.
            bucket.setdefault(
                comment_id,
                _Point(
                    story=story,
                    order=order,
                    paragraph=paragraph,
                    index=index_map.get(paragraph),
                    offset=segment.start,
                    in_table=in_table,
                ),
            )
    return starts, ends, references, paragraphs


def _anchored_text(
    paragraphs: list[etree._Element], start: _Point, end: _Point | None
) -> str:
    """The visible text between two marks of the same story."""
    if end is None or end.order < start.order:
        return ""
    if end.order == start.order:
        return visible_text(start.paragraph)[start.offset : end.offset]
    pieces = [visible_text(start.paragraph)[start.offset :]]
    pieces += [visible_text(p) for p in paragraphs[start.order + 1 : end.order]]
    pieces.append(visible_text(end.paragraph)[: end.offset])
    return "\n".join(pieces)


def anchors(pkg: DocxPackage) -> dict[str, Anchor]:
    """Map every anchored comment id of `pkg` to where it is anchored."""
    found: dict[str, Anchor] = {}
    for story, root in pkg.stories():
        starts, ends, references, paragraphs = _story_points(story, root)
        for comment_id, point in starts.items():
            found[comment_id] = Anchor(
                story=point.story,
                paragraph_index=point.index,
                in_table=point.in_table,
                text=_anchored_text(paragraphs, point, ends.get(comment_id)),
            )
        # A reference mark without a range still tells which paragraph carries
        # the comment; it covers no text, so its anchored text stays empty.
        for comment_id, point in references.items():
            found.setdefault(
                comment_id,
                Anchor(
                    story=point.story,
                    paragraph_index=point.index,
                    in_table=point.in_table,
                    text="",
                ),
            )
    return found


def _parse_date(raw: str) -> str | None:
    """Normalise a ``w:date`` to ISO 8601, or return it unchanged if it is not."""
    if not raw:
        return None
    try:
        return datetime.datetime.fromisoformat(raw.replace("Z", "+00:00")).isoformat()
    except ValueError:
        return raw


def _comment_text(element: etree._Element) -> str:
    """The body of a ``w:comment``, as its paragraphs' visible text."""
    paragraphs = list(element.iter(qn("w:p")))
    if not paragraphs:
        return ""
    return "\n".join(visible_text(paragraph) for paragraph in paragraphs).strip()


def extract_all_comments(pkg: DocxPackage) -> list[dict[str, Any]]:
    """Extract every comment of `pkg`, with its metadata and its anchor.

    Args:
        pkg: the package to read.  Nothing is modified.

    Returns:
        One dict per ``w:comment``, in the order ``word/comments.xml`` stores
        them.  A comment whose anchor is missing from every story is still
        reported, with ``paragraph_index`` ``None`` and an empty
        ``reference_text``.

    Raises:
        PackageError: if the comments part exists but cannot be read.
    """
    part = pkg.find_part(COMMENTS_PARTNAME)
    if part is None:
        return []
    root = pkg.root_of(part)
    located = anchors(pkg)

    comments: list[dict[str, Any]] = []
    for index, element in enumerate(root.iter(_W_COMMENT)):
        comment_id = element.get(_W_ID, str(index))
        anchor = located.get(comment_id)
        comments.append(
            {
                "id": f"comment_{index + 1}",
                "comment_id": comment_id,
                "author": element.get(qn("w:author"), "Unknown"),
                "initials": element.get(qn("w:initials"), ""),
                "date": _parse_date(element.get(qn("w:date"), "")),
                "text": _comment_text(element),
                "story": None if anchor is None else anchor.story,
                "paragraph_index": None if anchor is None else anchor.paragraph_index,
                "in_table": False if anchor is None else anchor.in_table,
                "reference_text": "" if anchor is None else anchor.text,
            }
        )
    return comments


def filter_comments_by_author(
    comments: list[dict[str, Any]], author: str
) -> list[dict[str, Any]]:
    """Return the comments of `author`, matched case-insensitively."""
    wanted = author.lower()
    return [c for c in comments if (c.get("author") or "").lower() == wanted]


def get_comments_for_paragraph(
    comments: list[dict[str, Any]], paragraph_index: int, story: str = "document"
) -> list[dict[str, Any]]:
    """Return the comments anchored in one paragraph of one story.

    `paragraph_index` is a V2 index, the space
    :attr:`word_document_server.engine.find.Match.index` reports; a comment
    anchored in a table cell or a text box has no such index and is therefore
    never returned.
    """
    return [
        c
        for c in comments
        if c.get("paragraph_index") == paragraph_index and c.get("story") == story
    ]
