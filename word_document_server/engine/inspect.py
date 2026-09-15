"""Document inspection: what an agent reads before writing a locator.

:func:`inspect` answers "what is in this document, and how do I name it" in one
plain-data call.  Everything it reports is addressable afterwards: a block
carries the V2 index the ``{"paragraph": i}`` locator takes, a table carries the
``{"table": t, ...}`` coordinates, a bookmark its name, a heading its text.  The
result is JSON-ready -- dicts, lists, strings, ints, bools and ``None``, no
lxml element, no live reference -- so the tool layer can hand it straight to a
model and so nothing in it goes stale in a way that could be mistaken for a
handle.

What is described, and from where
---------------------------------
``blocks`` describes the *main document story* only: it is the story the
locators number by default, and emitting a block per paragraph of every header
and footer would bury it.  The other stories are named in ``stories`` with their
paragraph counts, and are reachable with the ``story`` key of a locator;
``bookmarks`` and ``fields``, which an agent looks up by name rather than by
position, are reported across every story and carry the story they belong to.

Indices are the V2 ones throughout -- computed by the one filter,
:func:`~word_document_server.engine.find._v2_index_map` -- so a block index can
be pasted into a locator, and a bookmark or a field sitting in a table cell or a
text box reports ``"paragraph": null`` rather than a number that would address
another paragraph.

Text is truncated to :data:`TEXT_PREVIEW` characters, with ``truncated`` saying
so: inspection is a map, not a copy of the document.  Measurements are reported
in twips (1/1440 inch), the unit the package itself stores, never converted.
"""

from __future__ import annotations

from collections import Counter
from typing import Any

from lxml import etree

from word_document_server.engine.errors import PackageError
from word_document_server.engine.find import _v2_index_map, iter_paragraphs
from word_document_server.engine.locators import (
    _bookmark_span,
    _is_heading,
    _paragraph_of,
    _style_id,
    _style_names,
    _tables,
    indexed_paragraphs,
)
from word_document_server.engine.package import MAIN_STORY, DocxPackage, story_name
from word_document_server.engine.revisions import list_revisions
from word_document_server.engine.textmodel import fields, visible_text
from word_document_server.engine.xmlns import qn

__all__ = ["TEXT_PREVIEW", "inspect"]

#: Characters of text reported per block, per table cell and per heading.
TEXT_PREVIEW = 80

_W_P = qn("w:p")
_W_TR = qn("w:tr")
_W_TC = qn("w:tc")
_W_PPR = qn("w:pPr")
_W_NUMPR = qn("w:numPr")
_W_ILVL = qn("w:ilvl")
_W_OUTLINE_LVL = qn("w:outlineLvl")
_W_PSTYLE = qn("w:pStyle")
_W_RSTYLE = qn("w:rStyle")
_W_TBLSTYLE = qn("w:tblStyle")
_W_VAL = qn("w:val")
_W_NAME = qn("w:name")
_W_SECT_PR = qn("w:sectPr")
_W_SECT_PR_CHANGE = qn("w:sectPrChange")
_W_PG_SZ = qn("w:pgSz")
_W_PG_MAR = qn("w:pgMar")
_W_ORIENT = qn("w:orient")
_W_W = qn("w:w")
_W_H = qn("w:h")
_W_TYPE = qn("w:type")
_W_HEADER_REFERENCE = qn("w:headerReference")
_W_FOOTER_REFERENCE = qn("w:footerReference")
_W_BOOKMARK_START = qn("w:bookmarkStart")
_W_COMMENT = qn("w:comment")
_R_ID = qn("r:id")

#: Elements that make a paragraph carry a comment.
_COMMENT_TAGS = (
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:commentReference"),
)

#: Margin attributes of ``w:pgMar``, reported under these names in twips.
_MARGINS = ("top", "right", "bottom", "left", "header", "footer", "gutter")

#: Known page sizes, in twips, recognised within :data:`_PAGE_TOLERANCE`.
_PAGE_SIZES = {
    (11906, 16838): "A4",
    (16838, 23811): "A3",
    (8391, 11906): "A5",
    (12240, 15840): "Letter",
    (12240, 20160): "Legal",
    (12240, 18720): "Folio",
    (10440, 15840): "Executive",
}

#: How far a ``w:pgSz`` may sit from a known size and still be named by it --
#: Word rounds, and a document that went through a conversion drifts by a twip
#: or two.
_PAGE_TOLERANCE = 20


# --------------------------------------------------------------------------------------
# Blocks
# --------------------------------------------------------------------------------------


def _int_attr(element: etree._Element | None, name: str) -> int | None:
    """Read an integer attribute, returning ``None`` when it is absent or junk."""
    if element is None:
        return None
    raw = element.get(name)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _list_level(paragraph: etree._Element) -> int | None:
    """The numbering level of `paragraph`, or ``None`` when it carries none.

    Only *direct* numbering (``w:pPr/w:numPr``) counts.  Numbering inherited from
    a style is deliberately not resolved here: it would mean walking the style
    chain and the numbering part to produce a level the caller cannot act on any
    differently, and a wrong answer would read exactly like a right one.
    """
    properties = paragraph.find(_W_PPR)
    if properties is None:
        return None
    numbering = properties.find(_W_NUMPR)
    if numbering is None:
        return None
    level = _int_attr(numbering.find(_W_ILVL), _W_VAL)
    return 0 if level is None else level


def _heading_level(paragraph: etree._Element, style_names: dict[str, str]) -> int | None:
    """The outline level of a heading paragraph, 1-based, or ``None``.

    ``Title`` reads as level 0, above ``Heading1``, which is what Word means by
    it.  An explicit ``w:outlineLvl`` wins over the style, since that is the
    property Word itself honours.
    """
    properties = paragraph.find(_W_PPR)
    if properties is not None:
        outline = _int_attr(properties.find(_W_OUTLINE_LVL), _W_VAL)
        if outline is not None:
            return outline + 1
    style_id = _style_id(paragraph)
    if style_id is None:
        return None
    for candidate in (style_id, style_names.get(style_id, "")):
        cleaned = candidate.replace(" ", "").lower()
        if cleaned.startswith("heading") and cleaned[len("heading") :].isdigit():
            return int(cleaned[len("heading") :])
        if cleaned == "title":
            return 0
    return None


def _has_comment(paragraph: etree._Element) -> bool:
    """Whether a comment anchor or reference sits in `paragraph`."""
    return next(paragraph.iter(*_COMMENT_TAGS), None) is not None


def _block(
    index: int,
    paragraph: etree._Element,
    style_names: dict[str, str],
    revised: set[int],
) -> dict[str, Any]:
    text = visible_text(paragraph)
    heading = _is_heading(paragraph, style_names)
    level = _list_level(paragraph)
    if heading:
        kind = "heading"
    elif level is not None:
        kind = "list_item"
    else:
        kind = "paragraph"
    return {
        "index": index,
        "kind": kind,
        "style": _style_id(paragraph),
        "text": text[:TEXT_PREVIEW],
        "truncated": len(text) > TEXT_PREVIEW,
        "heading_level": _heading_level(paragraph, style_names) if heading else None,
        "list_level": level,
        "has_fields": bool(fields(paragraph)),
        "has_comments": _has_comment(paragraph),
        "has_revisions": index in revised,
    }


# --------------------------------------------------------------------------------------
# Tables
# --------------------------------------------------------------------------------------


def _table(index: int, table: etree._Element) -> dict[str, Any]:
    """Describe one table the way the ``table`` locator addresses it."""
    rows = table.findall(_W_TR)
    cells_per_row = [len(row.findall(_W_TC)) for row in rows]
    first_cell = ""
    if rows:
        first_cells = rows[0].findall(_W_TC)
        if first_cells:
            first_cell = " ".join(
                visible_text(paragraph) for paragraph in first_cells[0].findall(_W_P)
            ).strip()
    return {
        "index": index,
        "rows": len(rows),
        # The widest row, since ``col`` counts cells and a row may hold fewer of
        # them than another (a horizontal merge takes one position, not two).
        "columns": max(cells_per_row, default=0),
        "first_cell": first_cell[:TEXT_PREVIEW],
        "nested": any(ancestor.tag == _W_TC for ancestor in table.iterancestors()),
    }


# --------------------------------------------------------------------------------------
# Sections
# --------------------------------------------------------------------------------------


def _page_size_name(width: int | None, height: int | None) -> str | None:
    """Name the paper `width` x `height` twips stand for, either way round."""
    if width is None or height is None:
        return None
    short, long_side = sorted((width, height))
    for (known_short, known_long), name in _PAGE_SIZES.items():
        if (
            abs(short - known_short) <= _PAGE_TOLERANCE
            and abs(long_side - known_long) <= _PAGE_TOLERANCE
        ):
            return name
    return None


def _references(
    pkg: DocxPackage, root: etree._Element, section: etree._Element, tag: str
) -> dict[str, str]:
    """Map header/footer type -> story id for one ``w:sectPr``.

    A reference whose relationship is missing or external is skipped rather than
    reported as a story that cannot be opened.
    """
    found: dict[str, str] = {}
    for reference in section.findall(tag):
        rel_id = reference.get(_R_ID)
        if rel_id is None:
            continue
        try:
            target = pkg.rel_target(root, rel_id)
        except PackageError:
            continue
        if isinstance(target, str):  # pragma: no cover - a header is never external
            continue
        found[reference.get(_W_TYPE) or "default"] = story_name(str(target.partname))
    return found


def _sections(pkg: DocxPackage, root: etree._Element) -> list[dict[str, Any]]:
    """Describe every section of the main story, in document order.

    A ``w:sectPr`` held by a ``w:sectPrChange`` is a *previous* page setup
    recorded by a tracked revision, not a section of the document: it is left
    out, or every accepted layout change would add a phantom section.
    """
    live = [
        section
        for section in root.iter(_W_SECT_PR)
        if (parent := section.getparent()) is None or parent.tag != _W_SECT_PR_CHANGE
    ]
    described: list[dict[str, Any]] = []
    for index, section in enumerate(live):
        size = section.find(_W_PG_SZ)
        width = _int_attr(size, _W_W)
        height = _int_attr(size, _W_H)
        orient = None if size is None else size.get(_W_ORIENT)
        if orient is None and width is not None and height is not None:
            orient = "landscape" if width > height else "portrait"
        margins_element = section.find(_W_PG_MAR)
        described.append(
            {
                "index": index,
                "format": {
                    "width": width,
                    "height": height,
                    "name": _page_size_name(width, height),
                },
                "orientation": orient,
                "margins": {
                    name: _int_attr(margins_element, qn(f"w:{name}")) for name in _MARGINS
                },
                "headers": _references(pkg, root, section, _W_HEADER_REFERENCE),
                "footers": _references(pkg, root, section, _W_FOOTER_REFERENCE),
            }
        )
    return described


# --------------------------------------------------------------------------------------
# Cross-story inventories
# --------------------------------------------------------------------------------------


def _bookmarks(stories: list[tuple[str, etree._Element]]) -> list[dict[str, Any]]:
    """Every ``w:bookmarkStart`` of the package, in story then document order."""
    found: list[dict[str, Any]] = []
    for story_id, root in stories:
        index_map = _v2_index_map(iter_paragraphs(root))
        for element in root.iter(_W_BOOKMARK_START):
            name = element.get(_W_NAME)
            if name is None:
                continue
            paragraph = _paragraph_of(element)
            entry: dict[str, Any] = {
                "name": name,
                "story": story_id,
                "paragraph": None if paragraph is None else index_map.get(paragraph),
                "start": None,
                "end": None,
            }
            if paragraph is not None:
                entry["start"], entry["end"] = _bookmark_span(paragraph, element)
            found.append(entry)
    return found


def _fields(stories: list[tuple[str, etree._Element]]) -> list[dict[str, Any]]:
    """Every field of the package, instruction first, in document order."""
    found: list[dict[str, Any]] = []
    for story_id, root in stories:
        index_map = _v2_index_map(iter_paragraphs(root))
        for paragraph in iter_paragraphs(root):
            for field in fields(paragraph):
                found.append(
                    {
                        # Stripped: Word pads instructions with spaces
                        # (``" PAGE   \\* MERGEFORMAT "``) and the padding is
                        # noise to a reader.
                        "instr": field.instr.strip(),
                        "story": story_id,
                        "paragraph": index_map.get(paragraph),
                        "start": field.start,
                        "end": field.end,
                    }
                )
    return found


def _styles_used(stories: list[tuple[str, etree._Element]]) -> dict[str, dict[str, int]]:
    """Count the style ids actually referenced, by kind.

    Only references are counted, never definitions: a styles part lists dozens
    of styles a document does not use, and the question an agent asks before
    formatting is "what does this document already use".
    """
    counters: dict[str, Counter[str]] = {
        "paragraph": Counter(),
        "character": Counter(),
        "table": Counter(),
    }
    kinds = {_W_PSTYLE: "paragraph", _W_RSTYLE: "character", _W_TBLSTYLE: "table"}
    for _, root in stories:
        for element in root.iter(*kinds):
            value = element.get(_W_VAL)
            if value is not None:
                counters[kinds[element.tag]][value] += 1
    return {kind: dict(sorted(counter.items())) for kind, counter in counters.items()}


def _comment_count(pkg: DocxPackage) -> int:
    """Number of comments in ``word/comments.xml``, 0 when the part is absent."""
    part = pkg.find_part("/word/comments.xml")
    if part is None:
        return 0
    try:
        root = pkg.root_of(part)
    except PackageError:  # pragma: no cover - comments.xml is in LIVE_CONTENT_TYPES
        return 0
    return len(root.findall(_W_COMMENT))


# --------------------------------------------------------------------------------------
# The report
# --------------------------------------------------------------------------------------


def inspect(pkg: DocxPackage) -> dict[str, Any]:
    """Describe `pkg`: its blocks, tables, sections, stories and inventories.

    The returned mapping has eight keys:

    ``stories``
        one entry per story -- ``{"id", "paragraphs"}`` -- main document first,
        ``paragraphs`` being the size of that story's V2 index space.
    ``counts``
        the totals an agent checks before and after an edit: ``paragraphs``,
        ``tables``, ``sections``, ``bookmarks``, ``fields``, ``comments``,
        ``revisions``, ``stories``, ``words`` and ``characters``.  The first two
        and the last two are about the main document story; the rest cover the
        whole package.
    ``blocks``
        the paragraphs of the main document story: ``index`` (the V2 index a
        locator takes), ``kind`` (``heading``, ``list_item`` or ``paragraph``),
        ``style``, ``text`` (truncated, with ``truncated``), ``heading_level``,
        ``list_level``, and the three presence flags ``has_fields``,
        ``has_comments`` and ``has_revisions``.
    ``tables``
        ``index``, ``rows``, ``columns``, ``first_cell`` and ``nested`` for each
        table of the main story, numbered as the ``table`` locator numbers them.
    ``sections``
        ``format`` (width, height and paper name in twips), ``orientation``,
        ``margins`` and the ``headers``/``footers`` story ids of each section.
    ``styles_used``
        referenced style ids and their counts, split into ``paragraph``,
        ``character`` and ``table``.
    ``bookmarks``, ``fields``
        every one of them, across every story, each carrying its ``story`` and
        the V2 ``paragraph`` index it sits in (``None`` in a cell or a text box).

    Nothing is mutated and nothing is cached: the report describes the tree as
    it stands at the call.
    """
    stories = pkg.stories()
    document_root = pkg.document
    style_names = _style_names(pkg)
    revisions = list_revisions(pkg)
    revised = {
        revision.paragraph_index
        for revision in revisions
        if revision.story == MAIN_STORY and revision.paragraph_index is not None
    }

    paragraphs = indexed_paragraphs(document_root)
    blocks = [
        _block(index, paragraph, style_names, revised)
        for index, paragraph in enumerate(paragraphs)
    ]
    tables = [_table(index, table) for index, table in enumerate(_tables(document_root))]
    sections = _sections(pkg, document_root)
    bookmarks = _bookmarks(stories)
    document_fields = _fields(stories)
    texts = [visible_text(paragraph) for paragraph in paragraphs]

    return {
        "stories": [
            {"id": story_id, "paragraphs": len(indexed_paragraphs(root))}
            for story_id, root in stories
        ],
        "counts": {
            "stories": len(stories),
            "paragraphs": len(blocks),
            "tables": len(tables),
            "sections": len(sections),
            "bookmarks": len(bookmarks),
            "fields": len(document_fields),
            "comments": _comment_count(pkg),
            "revisions": len(revisions),
            "words": sum(len(text.split()) for text in texts),
            "characters": sum(len(text) for text in texts),
        },
        "blocks": blocks,
        "tables": tables,
        "sections": sections,
        "styles_used": _styles_used(stories),
        "bookmarks": bookmarks,
        "fields": document_fields,
    }
