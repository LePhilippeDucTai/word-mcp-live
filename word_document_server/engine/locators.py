"""Stateless locators: naming a place in a document without a session.

A locator is a plain ``dict`` an agent can write, store, and send back later.
Nothing about it is registered anywhere: there is no session, no identifier
table, no handle to keep alive.  Two calls that pass the same locator against
the same bytes reach the same paragraph, and a locator that no longer describes
the document fails loudly instead of quietly hitting the wrong one.

Forms
-----
Exactly one of these keys selects the form:

``{"paragraph": i}``
    the paragraph at *V2 index* `i` -- see below.
``{"find": s, "occurrence": n, "within": locator}``
    the text `s`, searched the way
    :func:`~word_document_server.engine.find.find` searches (literally, on the
    visible text, case-sensitively).  ``occurrence`` is 1-based and optional:
    without it, `s` must occur exactly once or the locator is ``ambiguous``.
    ``within`` is another locator, and restricts the search to the span it
    resolves to.
``{"bookmark": name}``
    the paragraph holding the ``w:bookmarkStart`` named `name`; the span runs
    to the matching ``w:bookmarkEnd`` when it sits in the same paragraph.
``{"heading": s}``
    the heading paragraph whose text is `s`, or, failing an exact match, the
    one it is a prefix of.
``{"table": t, "row": r, "col": c, "paragraph": k}``
    the ``k``-th paragraph (default 0) of a table cell.  ``paragraph`` means
    "paragraph inside that cell" here, which is why the table form owns the key.

Two keys are accepted on every form:

``story``
    a story id as :meth:`~word_document_server.engine.package.DocxPackage.stories`
    names it (``"document"``, ``"header1"``, ``"footnotes"``, ...), defaulting
    to the main document.  ``"body"`` is accepted as an alias of ``"document"``
    on the way in, exactly as in :func:`find`, and is never reported back:
    :attr:`Target.story` always carries the real id.
``expect_text``
    what the caller believes the located paragraph reads: its exact visible
    text, or a prefix of it.  A mismatch is the ``stale_anchor`` failure -- the
    document moved under an index the caller recorded earlier -- and the error
    message names the closest paragraphs that *do* match, so the caller can
    re-aim without re-reading the whole document.

Index space
-----------
The ``paragraph`` form addresses the V2 index space, and
:attr:`Target.index` reports it: the rank of a ``w:p`` among the top-level
paragraphs of its story, block ``w:sdt`` content included, table cells and text
boxes excluded.  That space is computed by
:func:`~word_document_server.engine.find._v2_index_map`, the single filter
shared by :func:`find`, :func:`list_revisions` and, on the tool side, by
``utils.document_utils.indexed_paragraphs`` -- one filter, two wrappers, never a
third.  A paragraph outside it has no index: a cell paragraph is reached by the
``table`` form, a text-box paragraph by ``find`` alone, and both report
``index=None``.

Tables are numbered in their own space: every ``w:tbl`` of the story in document
order, nested tables included (a table in a cell follows its host), text boxes
and ``mc:AlternateContent`` branches excluded.  ``row`` counts the ``w:tr``
children of the table and ``col`` the ``w:tc`` children of that row -- a cell
merged across two grid columns occupies one position, not two.

Failures
--------
Every failure is a :class:`~word_document_server.engine.errors.LocatorError`
carrying one of four codes:

``invalid``
    the locator is not well formed -- no form key, two of them, an unknown key,
    a negative or non-integer index.  Nothing was looked up.
``not_found``
    the locator is well formed but the document has no such place.
``ambiguous``
    it describes more than one place, and the caller did not say which.
``stale_anchor``
    ``expect_text`` does not match what is actually there.

The message always describes what was found instead, and lists the closest
candidates when there are any; callers branch on `code`, humans read the rest.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from lxml import etree

from word_document_server.engine.errors import LocatorError, PackageError
from word_document_server.engine.find import (
    _BODY_ALIAS,
    _v2_index_map,
    find,
    iter_paragraphs,
)
from word_document_server.engine.package import MAIN_STORY, DocxPackage
from word_document_server.engine.textmodel import segments, visible_text
from word_document_server.engine.xmlns import qn

__all__ = ["Target", "indexed_paragraphs", "resolve"]

_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_TR = qn("w:tr")
_W_TC = qn("w:tc")
_W_PPR = qn("w:pPr")
_W_PSTYLE = qn("w:pStyle")
_W_OUTLINE_LVL = qn("w:outlineLvl")
_W_STYLE = qn("w:style")
_W_STYLE_ID = qn("w:styleId")
_W_NAME = qn("w:name")
_W_VAL = qn("w:val")
_W_BOOKMARK_START = qn("w:bookmarkStart")
_W_BOOKMARK_END = qn("w:bookmarkEnd")
_W_ID = qn("w:id")

#: Ancestors that take a ``w:tbl`` out of the table index space.  Narrower than
#: the paragraph one on purpose: a table nested in a cell *is* addressable, and
#: numbering it is the only way to reach its cells.
_UNINDEXED_TABLE_ANCESTORS = frozenset({qn("w:txbxContent"), qn("mc:AlternateContent")})

#: Style ids and style names that make a paragraph a heading for the ``heading``
#: form: ``Heading1``, ``heading 1``, ``Title``.
_HEADING_STYLE = re.compile(r"(?:heading\s*\d+|title)", re.IGNORECASE)

#: Keys that select a form, most specific first: ``table`` also carries a
#: ``paragraph`` key, so it has to be recognised before the ``paragraph`` form.
_FORM_KEYS = ("table", "paragraph", "find", "bookmark", "heading")

#: Keys accepted whatever the form.
_COMMON_KEYS = frozenset({"story", "expect_text"})

#: Keys each form accepts on top of :data:`_COMMON_KEYS`.
_FORM_EXTRA_KEYS: dict[str, frozenset[str]] = {
    "table": frozenset({"table", "row", "col", "paragraph"}),
    "paragraph": frozenset({"paragraph"}),
    "find": frozenset({"find", "occurrence", "within"}),
    "bookmark": frozenset({"bookmark"}),
    "heading": frozenset({"heading"}),
}

#: How many candidates an ``ambiguous``, ``not_found`` or ``stale_anchor``
#: message lists before giving up on being exhaustive.
_MAX_CANDIDATES = 5

#: How much of a paragraph's text a message quotes.
_PREVIEW = 40


@dataclass(frozen=True)
class Target:
    """The place a locator resolves to.

    Deliberately the same shape as
    :class:`~word_document_server.engine.find.Match`, so that a caller can treat
    a search result and a resolved locator alike.

    Attributes:
        story: id of the story, as :meth:`DocxPackage.stories` names it -- never
            the ``"body"`` alias, even when that alias selected it.
        paragraph: the live ``w:p`` element.
        index: the paragraph's V2 index within `story`, or ``None`` when it has
            none (a table cell, a text box).
        start: offset of the first targeted character in
            ``visible_text(paragraph)``.
        end: offset just past the last one.  Equal to `start` for an empty span,
            and to ``len(visible_text(paragraph))`` for a whole paragraph.
    """

    story: str
    paragraph: etree._Element
    index: int | None
    start: int
    end: int

    @property
    def text(self) -> str:
        """The visible text of the span, recomputed from the live paragraph."""
        return visible_text(self.paragraph)[self.start : self.end]


def indexed_paragraphs(story_root: etree._Element) -> list[etree._Element]:
    """Return the V2-indexed paragraphs of `story_root`, in document order.

    The list position *is* the V2 index.  Same filter as
    ``utils.document_utils.indexed_paragraphs`` -- both call
    :func:`~word_document_server.engine.find._v2_index_map`; this one is the
    engine-side wrapper, that one the tool-side one.
    """
    return list(_v2_index_map(iter_paragraphs(story_root)))


# --------------------------------------------------------------------------------------
# Locator validation
# --------------------------------------------------------------------------------------


def _form_of(locator: Any) -> str:
    """Return the form `locator` uses, after checking its keys.

    Raises:
        LocatorError: ``invalid`` if `locator` is not a dict, names no form or
            more than one, or carries a key the form does not accept.
    """
    if not isinstance(locator, dict):
        raise LocatorError("invalid", f"a locator must be a dict, got {type(locator).__name__}")
    named = [key for key in _FORM_KEYS if key in locator]
    # ``{"table": ..., "paragraph": k}`` is one form, not two: inside a table
    # locator, ``paragraph`` names the paragraph of the cell.
    if "table" in named:
        named = ["table"]
    if not named:
        raise LocatorError(
            "invalid",
            f"a locator must name one of {', '.join(sorted(_FORM_KEYS))}; got keys "
            f"{sorted(locator) or '[]'}",
        )
    if len(named) > 1:
        raise LocatorError(
            "invalid", f"a locator names exactly one form, got {sorted(named)}"
        )
    form = named[0]
    unknown = set(locator) - _COMMON_KEYS - _FORM_EXTRA_KEYS[form]
    if unknown:
        accepted = sorted(_COMMON_KEYS | _FORM_EXTRA_KEYS[form])
        raise LocatorError(
            "invalid",
            f"key(s) {sorted(unknown)} are not accepted by the {form!r} locator; "
            f"accepted: {accepted}",
        )
    return form


def _index_value(locator: dict[str, Any], key: str, *, default: int | None = None) -> int:
    """Return the non-negative integer stored under `key`.

    Raises:
        LocatorError: ``invalid`` if the key is missing without a `default`, or
            holds anything but a non-negative integer.  ``True`` is rejected:
            ``bool`` is an ``int`` in Python, but never an index here.
    """
    if key not in locator:
        if default is None:
            raise LocatorError("invalid", f"this locator requires an integer {key!r}")
        return default
    value = locator[key]
    if isinstance(value, bool) or not isinstance(value, int):
        raise LocatorError(
            "invalid", f"{key!r} must be an integer, got {type(value).__name__}"
        )
    if value < 0:
        raise LocatorError("invalid", f"{key!r} must be 0 or more, got {value}")
    return value


def _text_value(locator: dict[str, Any], key: str) -> str:
    """Return the non-empty string stored under `key`.

    Raises:
        LocatorError: ``invalid`` if it is missing, not a string, or empty.
    """
    value = locator.get(key)
    if not isinstance(value, str):
        raise LocatorError(
            "invalid",
            f"{key!r} must be a string, got {type(value).__name__}",
        )
    if not value:
        raise LocatorError("invalid", f"{key!r} must not be empty")
    return value


def _story_of(pkg: DocxPackage, locator: dict[str, Any]) -> tuple[str, etree._Element]:
    """Resolve the ``story`` key to ``(real story id, live root element)``.

    Raises:
        LocatorError: ``invalid`` if ``story`` is not a non-empty string,
            ``not_found`` if the package has no such story.
    """
    requested = locator.get("story", MAIN_STORY)
    if not isinstance(requested, str) or not requested:
        raise LocatorError("invalid", f"'story' must be a non-empty string, got {requested!r}")
    story_id = MAIN_STORY if requested == _BODY_ALIAS else requested
    roots = dict(pkg.stories())
    root = roots.get(story_id)
    if root is None:
        raise LocatorError(
            "not_found",
            f"this package has no story {requested!r}; it has {sorted(roots)}",
        )
    return story_id, root


# --------------------------------------------------------------------------------------
# Shared helpers
# --------------------------------------------------------------------------------------


def _preview(text: str) -> str:
    """Quote `text` for an error message, shortened to :data:`_PREVIEW`."""
    if len(text) <= _PREVIEW:
        return repr(text)
    return repr(text[:_PREVIEW] + "...")


def _paragraph_of(element: etree._Element) -> etree._Element | None:
    """The ``w:p`` `element` sits in, or ``None`` if it sits outside any."""
    if element.tag == _W_P:
        return element
    for ancestor in element.iterancestors(_W_P):
        return ancestor
    return None


def _whole(story_id: str, paragraph: etree._Element, index: int | None) -> Target:
    """A target covering the whole visible text of `paragraph`."""
    return Target(story_id, paragraph, index, 0, len(visible_text(paragraph)))


def _style_id(paragraph: etree._Element) -> str | None:
    """The ``w:pStyle`` of `paragraph`, or ``None`` when it has none."""
    properties = paragraph.find(_W_PPR)
    if properties is None:
        return None
    style = properties.find(_W_PSTYLE)
    return None if style is None else style.get(_W_VAL)


def _style_names(pkg: DocxPackage) -> dict[str, str]:
    """Map style id -> ``w:name`` for the package's styles part.

    Empty when the package has no styles part: a heading is then recognised by
    its style id alone, which is how python-docx writes them anyway.
    """
    try:
        part = pkg.find_part("/word/styles.xml")
    except PackageError:  # pragma: no cover - find_part only raises on an empty name
        return {}
    if part is None:
        return {}
    try:
        root = pkg.root_of(part)
    except PackageError:  # pragma: no cover - styles.xml is in LIVE_CONTENT_TYPES
        return {}
    names: dict[str, str] = {}
    for style in root.iter(_W_STYLE):
        style_id = style.get(_W_STYLE_ID)
        name = style.find(_W_NAME)
        if style_id is not None and name is not None:
            value = name.get(_W_VAL)
            if value is not None:
                names[style_id] = value
    return names


def _is_heading(paragraph: etree._Element, style_names: dict[str, str]) -> bool:
    """Whether `paragraph` reads as a heading.

    Three signals, any of which is enough: a style id that looks like a heading
    (``Heading2``), a style *name* that does (``heading 2``, for a document whose
    ids are localized), or an explicit ``w:outlineLvl`` in the paragraph's own
    properties -- the property Word itself uses to put a paragraph in the
    navigation pane.
    """
    properties = paragraph.find(_W_PPR)
    if properties is not None and properties.find(_W_OUTLINE_LVL) is not None:
        return True
    style_id = _style_id(paragraph)
    if style_id is None:
        return False
    if _HEADING_STYLE.fullmatch(style_id):
        return True
    name = style_names.get(style_id)
    return name is not None and bool(_HEADING_STYLE.fullmatch(name))


def _matches_expectation(text: str, expected: str) -> bool:
    """Whether `text` satisfies an ``expect_text`` -- exactly, or as a prefix."""
    return text == expected or text.startswith(expected)


def _check_expectation(
    locator: dict[str, Any], target: Target, root: etree._Element
) -> None:
    """Verify the ``expect_text`` key of `locator` against `target`.

    `root` is the story root the target was resolved in, used to list the
    paragraphs that *do* carry the expected text.

    Raises:
        LocatorError: ``invalid`` if ``expect_text`` is not a string,
            ``stale_anchor`` if it does not match the located paragraph.
    """
    if "expect_text" not in locator:
        return
    expected = locator["expect_text"]
    if not isinstance(expected, str):
        raise LocatorError(
            "invalid",
            f"'expect_text' must be a string, got {type(expected).__name__}",
        )
    found = visible_text(target.paragraph)
    if _matches_expectation(found, expected):
        return
    where = (
        "the located paragraph" if target.index is None else f"paragraph {target.index}"
    )
    message = (
        f"expected {_preview(expected)} at {where} of story {target.story!r}, "
        f"found {_preview(found)}"
    )
    candidates = _anchor_candidates(root, expected, target.index)
    if candidates:
        message += f"; closest paragraph(s) carrying that text: {candidates}"
    else:
        message += "; no paragraph of that story carries that text"
    raise LocatorError("stale_anchor", message)


def _anchor_candidates(root: etree._Element, expected: str, around: int | None) -> list[int]:
    """V2 indices whose paragraph matches `expected`, nearest to `around` first."""
    found = [
        index
        for index, paragraph in enumerate(indexed_paragraphs(root))
        if _matches_expectation(visible_text(paragraph), expected)
    ]
    if around is not None:
        found.sort(key=lambda index: (abs(index - around), index))
    return found[:_MAX_CANDIDATES]


# --------------------------------------------------------------------------------------
# The forms
# --------------------------------------------------------------------------------------


def _resolve_paragraph(
    pkg: DocxPackage, locator: dict[str, Any]
) -> tuple[Target, etree._Element]:
    index = _index_value(locator, "paragraph")
    story_id, root = _story_of(pkg, locator)
    paragraphs = indexed_paragraphs(root)
    if index >= len(paragraphs):
        raise LocatorError(
            "not_found",
            f"story {story_id!r} has {len(paragraphs)} indexed paragraph(s), "
            f"so index {index} does not exist",
        )
    return _whole(story_id, paragraphs[index], index), root


def _resolve_find(
    pkg: DocxPackage, locator: dict[str, Any]
) -> tuple[Target, etree._Element]:
    pattern = _text_value(locator, "find")
    occurrence: int | None = None
    if "occurrence" in locator:
        occurrence = _index_value(locator, "occurrence")
        if occurrence < 1:
            raise LocatorError(
                "invalid", "'occurrence' counts from 1, so it must be 1 or more"
            )

    within = locator.get("within")
    host: Target | None = None
    if within is not None:
        host = resolve(pkg, within)
        # The span decides the story: searching elsewhere than where 'within'
        # landed could only ever return nothing.
        story_id = host.story
        root = dict(pkg.stories())[story_id]
        if "story" in locator:
            requested, _ = _story_of(pkg, locator)
            if requested != story_id:
                raise LocatorError(
                    "invalid",
                    f"'within' resolves in story {story_id!r}, which contradicts the "
                    f"requested story {requested!r}",
                )
    else:
        story_id, root = _story_of(pkg, locator)

    matches = find(pkg, pattern, stories=(story_id,))
    if host is not None:
        matches = [
            match
            for match in matches
            if match.paragraph is host.paragraph
            and host.start <= match.start
            and match.end <= host.end
        ]

    where = f"story {story_id!r}" if host is None else "the 'within' span"
    if not matches:
        raise LocatorError("not_found", f"{_preview(pattern)} does not occur in {where}")
    if occurrence is None:
        if len(matches) > 1:
            places = ", ".join(
                f"paragraph {match.index} at offset {match.start}"
                for match in matches[:_MAX_CANDIDATES]
            )
            raise LocatorError(
                "ambiguous",
                f"{_preview(pattern)} occurs {len(matches)} times in {where} "
                f"({places}); add 'occurrence' to choose one",
            )
        match = matches[0]
    else:
        if occurrence > len(matches):
            raise LocatorError(
                "not_found",
                f"{_preview(pattern)} occurs {len(matches)} time(s) in {where}, "
                f"so occurrence {occurrence} does not exist",
            )
        match = matches[occurrence - 1]
    return Target(match.story, match.paragraph, match.index, match.start, match.end), root


def _resolve_bookmark(
    pkg: DocxPackage, locator: dict[str, Any]
) -> tuple[Target, etree._Element]:
    name = _text_value(locator, "bookmark")
    story_id, root = _story_of(pkg, locator)
    starts = [
        element for element in root.iter(_W_BOOKMARK_START) if element.get(_W_NAME) == name
    ]
    if not starts:
        known = sorted(
            {
                value
                for element in root.iter(_W_BOOKMARK_START)
                if (value := element.get(_W_NAME)) is not None
            }
        )
        raise LocatorError(
            "not_found",
            f"story {story_id!r} has no bookmark named {name!r}; it has "
            f"{known[:_MAX_CANDIDATES]}",
        )
    if len(starts) > 1:
        raise LocatorError(
            "ambiguous",
            f"story {story_id!r} carries {len(starts)} bookmarks named {name!r}; "
            "a locator cannot choose between them",
        )
    start_element = starts[0]
    paragraph = _paragraph_of(start_element)
    if paragraph is None:
        # Legal OOXML: a bookmark may wrap whole rows or the body itself, and
        # then it anchors no paragraph to act on.
        raise LocatorError(
            "not_found",
            f"bookmark {name!r} of story {story_id!r} is not inside a paragraph",
        )
    index = _v2_index_map(iter_paragraphs(root)).get(paragraph)
    start, end = _bookmark_span(paragraph, start_element)
    return Target(story_id, paragraph, index, start, end), root


def _bookmark_span(
    paragraph: etree._Element, start_element: etree._Element
) -> tuple[int, int]:
    """Offsets of a bookmark inside its paragraph.

    The span stops at the ``w:bookmarkEnd`` carrying the same ``w:id``, and at
    the end of the paragraph when that end lives in a later one -- a bookmark may
    legally span paragraphs, and a target never crosses one.
    """
    bookmark_id = start_element.get(_W_ID)
    start = 0
    end: int | None = None
    for segment in segments(paragraph):
        if segment.element is start_element:
            start = segment.start
            end = None
            continue
        if (
            end is None
            and segment.element.tag == _W_BOOKMARK_END
            and segment.element.get(_W_ID) == bookmark_id
        ):
            end = segment.start
    text_length = len(visible_text(paragraph))
    return start, text_length if end is None else end


def _resolve_heading(
    pkg: DocxPackage, locator: dict[str, Any]
) -> tuple[Target, etree._Element]:
    wanted = _text_value(locator, "heading")
    story_id, root = _story_of(pkg, locator)
    style_names = _style_names(pkg)
    headings = [
        (index, paragraph)
        for index, paragraph in enumerate(indexed_paragraphs(root))
        if _is_heading(paragraph, style_names)
    ]
    exact = [(index, p) for index, p in headings if visible_text(p) == wanted]
    # Exact wins outright: a document holding both "Results" and "Results and
    # discussion" must not be ambiguous for the caller who typed the short one.
    chosen = exact or [(index, p) for index, p in headings if visible_text(p).startswith(wanted)]
    if not chosen:
        available = [visible_text(p) for _, p in headings[:_MAX_CANDIDATES]]
        raise LocatorError(
            "not_found",
            f"story {story_id!r} has no heading reading {_preview(wanted)}; its "
            f"headings start with {available}",
        )
    if len(chosen) > 1:
        raise LocatorError(
            "ambiguous",
            f"{len(chosen)} headings of story {story_id!r} read {_preview(wanted)} "
            f"(paragraphs {[index for index, _ in chosen[:_MAX_CANDIDATES]]}); "
            "address one by its index instead",
        )
    index, paragraph = chosen[0]
    return _whole(story_id, paragraph, index), root


def _tables(story_root: etree._Element) -> list[etree._Element]:
    """Every addressable ``w:tbl`` of `story_root`, in document order."""
    return [
        table
        for table in story_root.iter(_W_TBL)
        if not any(
            ancestor.tag in _UNINDEXED_TABLE_ANCESTORS for ancestor in table.iterancestors()
        )
    ]


def _resolve_table(
    pkg: DocxPackage, locator: dict[str, Any]
) -> tuple[Target, etree._Element]:
    table_index = _index_value(locator, "table")
    row_index = _index_value(locator, "row")
    column_index = _index_value(locator, "col")
    paragraph_index = _index_value(locator, "paragraph", default=0)
    story_id, root = _story_of(pkg, locator)

    tables = _tables(root)
    if table_index >= len(tables):
        raise LocatorError(
            "not_found",
            f"story {story_id!r} has {len(tables)} table(s), so table {table_index} "
            "does not exist",
        )
    rows = tables[table_index].findall(_W_TR)
    if row_index >= len(rows):
        raise LocatorError(
            "not_found",
            f"table {table_index} of story {story_id!r} has {len(rows)} row(s), so "
            f"row {row_index} does not exist",
        )
    cells = rows[row_index].findall(_W_TC)
    if column_index >= len(cells):
        raise LocatorError(
            "not_found",
            f"row {row_index} of table {table_index} has {len(cells)} cell(s), so "
            f"column {column_index} does not exist",
        )
    paragraphs = cells[column_index].findall(_W_P)
    if paragraph_index >= len(paragraphs):
        raise LocatorError(
            "not_found",
            f"cell ({row_index}, {column_index}) of table {table_index} has "
            f"{len(paragraphs)} paragraph(s), so paragraph {paragraph_index} does "
            "not exist",
        )
    # A cell paragraph is out of the V2 index space by construction: index None
    # is the answer, not a number the ``paragraph`` form could reuse.
    return _whole(story_id, paragraphs[paragraph_index], None), root


_RESOLVERS = {
    "paragraph": _resolve_paragraph,
    "find": _resolve_find,
    "bookmark": _resolve_bookmark,
    "heading": _resolve_heading,
    "table": _resolve_table,
}


def resolve(pkg: DocxPackage, locator: dict[str, Any]) -> Target:
    """Resolve `locator` against `pkg` and return the place it names.

    See the module docstring for the five forms, the two keys every form
    accepts, and the four failure codes.  Resolution reads the live tree on
    every call and holds nothing afterwards: a caller that edits the document
    resolves again rather than reusing a :class:`Target` computed before the
    edit.

    Raises:
        LocatorError: ``invalid``, ``not_found``, ``ambiguous`` or
            ``stale_anchor``; see the module docstring.
    """
    form = _form_of(locator)
    target, root = _RESOLVERS[form](pkg, locator)
    _check_expectation(locator, target, root)
    return target
