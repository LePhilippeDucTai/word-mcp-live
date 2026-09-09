"""Text search across a package's stories.

``find`` answers "where does this text appear" the same way a human would read
the document: on the *visible* text of each paragraph
(:func:`~word_document_server.engine.textmodel.visible_text`), never on the raw
``w:t`` content, so a search neither turns up text hidden under a tracked
deletion nor misses text that only exists inside a hyperlink, an insertion or a
content control.

Paragraph addressing
---------------------
Every :class:`Match` carries two, deliberately different, ways to name its
paragraph:

``paragraph``
    the live ``w:p`` element itself -- always present, usable straight away by
    a caller that wants to act on the paragraph without locating it again.
``index``
    the paragraph's *V2 index*: its rank, 0-based, among the top-level
    paragraphs of its story -- block ``w:sdt`` content counts as top-level,
    table cells do not.  This is the index space the stateless V2 locator
    ``{"paragraph": i, "expect_text": ...}`` addresses (see ``docs/plans``).
    A paragraph that sits inside a table cell has no such index -- the V2
    paragraph locator does not reach into tables, a ``{"table", "row", "col",
    "paragraph"}`` locator does instead -- so `index` is ``None`` for it.
    Three different "paragraph index" spaces coexist across this codebase
    (python-docx's own body list, a plain ``//w:p``, or ``body//w:p``); this
    one is none of them, on purpose.

Story names
-----------
:meth:`~word_document_server.engine.package.DocxPackage.stories` names the main
document story ``"document"``.  `find` additionally accepts the friendlier
alias ``"body"`` for it in the `stories` filter (it is the default), because a
caller of a *search* function is thinking "the body of the document", not
"the part named document.xml".  Every other story -- ``"header1"``,
``"footer2"``, ``"footnotes"``, ``"endnotes"``, ... -- is named exactly as
:meth:`DocxPackage.stories` names it; there is no plural alias for "every
header" or "every footer", because instances are not interchangeable (a
document can have distinct first-page, even-page and default headers).  A
requested story that does not exist in the package (no footnotes, no second
header) is silently skipped rather than treated as an error, so that a caller
can pass a fixed list of stories against documents that do not all have the
same parts.

Matching
--------
A literal search (`regex` false, the default) matches `pattern` as plain text,
case-sensitively unless `case` is false, and only on a whole-word boundary if
`whole_word` is true.  A `regex` search compiles `pattern` as a Python regular
expression instead, with `case` and `whole_word` layered on top the same way,
so whole-word matching or case-insensitivity does not have to be re-encoded
into every hand-written pattern.  A match never crosses a paragraph boundary:
each paragraph's visible text is searched on its own, so a phrase split across
two paragraphs is not found by design.

`pattern` must not be the empty string, and a `regex` pattern must not be able
to match zero characters -- both raise :class:`ValueError`, the second one as
soon as an actual empty match is produced while scanning, since an empty match
never advances and carries no useful position.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from word_document_server.engine.package import MAIN_STORY, DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import qn

if TYPE_CHECKING:
    from collections.abc import Sequence

__all__ = ["Match", "find", "iter_paragraphs"]

_W_P = qn("w:p")
_W_TC = qn("w:tc")

#: Friendly alias for :data:`~word_document_server.engine.package.MAIN_STORY`,
#: accepted (and defaulted to) by the `stories` filter of :func:`find`.
_BODY_ALIAS = "body"

#: Characters of visible text kept on each side of a match in ``Match.context``.
_CONTEXT_RADIUS = 40


@dataclass(frozen=True)
class Match:
    """One place `pattern` was found in a paragraph's visible text.

    Attributes:
        story: id of the story the match belongs to, as
            :meth:`~word_document_server.engine.package.DocxPackage.stories`
            names it (``"document"``, ``"header1"``, ``"footnotes"``, ...) --
            never the ``"body"`` alias, even when that alias selected it.
        paragraph: the live ``w:p`` element containing the match.
        index: the paragraph's V2 index within `story` (see the module
            docstring), or ``None`` if `paragraph` sits inside a table cell.
        start: offset of the first matched character in
            ``visible_text(paragraph)``.
        end: offset just past the last matched character.
        text: the matched text, equal to ``visible_text(paragraph)[start:end]``.
        context: up to :data:`_CONTEXT_RADIUS` characters of visible text on
            each side of the match, clipped to the paragraph's own text --
            never padded, never crossing into another paragraph.
    """

    story: str
    paragraph: etree._Element
    index: int | None
    start: int
    end: int
    text: str
    context: str


def iter_paragraphs(story_root: etree._Element) -> list[etree._Element]:
    """Return every ``w:p`` under `story_root`, in document order.

    Table cells and ``w:sdt`` content (block or inline) are descended into: this
    is every paragraph the story contains, not just the top-level ones -- see
    the module docstring for the narrower, V2 notion of paragraph index that
    `find` also reports. `story_root` is one of the roots
    :meth:`~word_document_server.engine.package.DocxPackage.stories` hands out,
    so it is never itself a ``w:p``.

    The walk is a plain depth-first traversal of the live XML tree, so it also
    surfaces a paragraph nested in a construct :mod:`textmodel` does not treat
    as running text (a text box's ``w:txbxContent``, for instance) -- callers
    that only want paragraphs contributing to the document's reading flow
    should filter on their own context, the same way `find` filters out table
    cells for V2 indexing.
    """
    return list(story_root.iter(_W_P))


def _v2_index_map(paragraphs: list[etree._Element]) -> dict[etree._Element, int]:
    """Map each top-level paragraph in `paragraphs` to its 0-based V2 index.

    A paragraph is top-level unless one of its ancestors, up to (and excluding)
    the story root, is a table cell (``w:tc``) -- a ``w:sdt`` ancestor, block or
    inline, does not disqualify it.  Paragraphs inside a table cell are simply
    absent from the returned map; they do not consume a slot in the count, so
    the numbering seen by a top-level paragraph is unaffected by tables that
    precede it.
    """
    index_map: dict[etree._Element, int] = {}
    counter = 0
    for paragraph in paragraphs:
        if any(ancestor.tag == _W_TC for ancestor in paragraph.iterancestors()):
            continue
        index_map[paragraph] = counter
        counter += 1
    return index_map


def _compile_pattern(pattern: str, *, regex: bool, case: bool, whole_word: bool) -> re.Pattern[str]:
    if pattern == "":
        raise ValueError("pattern must not be empty")
    core = pattern if regex else re.escape(pattern)
    if whole_word:
        core = rf"\b(?:{core})\b"
    flags = re.IGNORECASE if not case else 0
    try:
        return re.compile(core, flags)
    except re.error as exc:
        raise ValueError(f"invalid regular expression {pattern!r}: {exc}") from exc


def _resolve_story_ids(stories: Sequence[str]) -> list[str]:
    """Translate the `stories` filter to package story ids, in order, deduplicated."""
    resolved: list[str] = []
    seen: set[str] = set()
    for story in stories:
        story_id = MAIN_STORY if story == _BODY_ALIAS else story
        if story_id in seen:
            continue
        seen.add(story_id)
        resolved.append(story_id)
    return resolved


def find(
    pkg: DocxPackage,
    pattern: str,
    *,
    regex: bool = False,
    case: bool = True,
    whole_word: bool = False,
    stories: Sequence[str] = (_BODY_ALIAS,),
    max_results: int | None = None,
) -> list[Match]:
    """Search `pkg` for `pattern`, paragraph by paragraph, and return the matches.

    See the module docstring for what `regex`, `case`, `whole_word` and the
    `stories` filter mean, and for the two paragraph-addressing fields of
    :class:`Match`. Results are in document order: stories in the order given
    by `stories` (after resolving the ``"body"`` alias and deduplicating), then
    paragraphs of each story in document order, then matches of each paragraph
    in left-to-right order. `max_results` stops the search as soon as that many
    matches have been collected, so a caller after the first hit does not pay
    for scanning the rest of a large document.

    Raises:
        ValueError: `pattern` is the empty string, `regex` is true and
            `pattern` is not a valid regular expression, or a `regex` search
            produces a zero-length match.
    """
    compiled = _compile_pattern(pattern, regex=regex, case=case, whole_word=whole_word)
    roots_by_story = dict(pkg.stories())

    matches: list[Match] = []
    for story_id in _resolve_story_ids(stories):
        root = roots_by_story.get(story_id)
        if root is None:
            continue
        paragraphs = iter_paragraphs(root)
        index_map = _v2_index_map(paragraphs)
        for paragraph in paragraphs:
            text = visible_text(paragraph)
            for occurrence in compiled.finditer(text):
                if max_results is not None and len(matches) >= max_results:
                    return matches
                start, end = occurrence.start(), occurrence.end()
                if start == end:
                    raise ValueError(
                        f"pattern {pattern!r} matched an empty string in story "
                        f"{story_id!r} at offset {start}"
                    )
                context = text[max(0, start - _CONTEXT_RADIUS) : min(len(text), end + _CONTEXT_RADIUS)]
                matches.append(
                    Match(
                        story=story_id,
                        paragraph=paragraph,
                        index=index_map.get(paragraph),
                        start=start,
                        end=end,
                        text=occurrence.group(),
                        context=context,
                    )
                )
    return matches
