"""Anchoring a comment on a document's text, through the OOXML engine.

A comment is not one part but five.  ``word/comments.xml`` holds the text,
``word/commentsExtended.xml`` holds the thread state Word's review pane reads,
``word/commentsIds.xml`` holds the durable identity that survives a copy/paste
between documents, ``word/people.xml`` declares the author, and the story itself
carries the three anchor marks (``w:commentRangeStart``,
``w:commentRangeEnd``, ``w:commentReference``).  A comment written without the
side-car parts shows up in the document but has no working reply thread, which
is why :func:`add_comment_to_doc` creates all four through
:meth:`~word_document_server.engine.package.DocxPackage.ensure_part` -- an
idempotent call that reuses an existing part and its relationship.

Nothing here parses the zip or allocates a relationship id by hand: the package
layer owns both, the anchor marks are placed by
:func:`~word_document_server.engine.format.wrap`, and the identifiers come from
:mod:`word_document_server.engine.ids`, so no id is ever reused.
"""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

from word_document_server.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS
from word_document_server.engine.errors import EngineError, IdExhausted
from word_document_server.engine.find import find
from word_document_server.engine.format import (
    COMMENT_REFERENCE_STYLE,
    character_style_exists,
)
from word_document_server.engine.format import comment as comment_envelope
from word_document_server.engine.format import wrap
from word_document_server.engine.ids import next_comment_id, next_para_id
from word_document_server.engine.package import COMMENTS_CONTENT_TYPE, DocxPackage
from word_document_server.engine.ranges import resolve
from word_document_server.engine.xmlns import MC, W14, W15, W, qn

__all__ = [
    "COMMENTS_EXTENDED_PARTNAME",
    "COMMENTS_IDS_PARTNAME",
    "COMMENTS_PARTNAME",
    "PEOPLE_PARTNAME",
    "add_comment_to_doc",
]

#: Word 2016 comment identity namespace.  It is the one namespace this file
#: needs that the engine never writes, so it is declared here rather than in
#: ``engine/xmlns.py``.
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"

COMMENTS_PARTNAME = "/word/comments.xml"
COMMENTS_EXTENDED_PARTNAME = "/word/commentsExtended.xml"
COMMENTS_IDS_PARTNAME = "/word/commentsIds.xml"
PEOPLE_PARTNAME = "/word/people.xml"

_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml."
CT_COMMENTS = COMMENTS_CONTENT_TYPE
CT_COMMENTS_EXTENDED = _WML + "commentsExtended+xml"
CT_COMMENTS_IDS = _WML + "commentsIds+xml"
CT_PEOPLE = _WML + "people+xml"

RT_COMMENTS = RT.COMMENTS
RT_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
RT_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
RT_PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people"

#: Empty roots for the four parts, used only when the package has none.  Each
#: declares exactly the prefixes its content uses, and marks the extension
#: namespace ignorable so a consumer that does not know it still opens the file.
_COMMENTS_XML = f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}" xmlns:mc="{MC}" mc:Ignorable="w14"/>'
_COMMENTS_EXTENDED_XML = (
    f'<w15:commentsEx xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15"/>'
)
_COMMENTS_IDS_XML = (
    f'<w16cid:commentsIds xmlns:w16cid="{W16CID}" xmlns:mc="{MC}" mc:Ignorable="w16cid"/>'
)
_PEOPLE_XML = f'<w15:people xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15"/>'

#: Half-point size Word gives comment body text.
_COMMENT_TEXT_HALF_POINTS = "20"

_W_ID = qn("w:id")
_W_T = qn("w:t")
_W_VAL = qn("w:val")
_XML_SPACE = qn("xml:space")


def _cid(local: str) -> str:
    """Clark-notation name in the Word 2016 comment identity namespace."""
    return f"{{{W16CID}}}{local}"


def _now_iso() -> str:
    """Current UTC instant, in the ``w:date`` form Word writes."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _text_run(parent: etree._Element, text: str) -> etree._Element:
    """Append a run carrying `text` verbatim, whitespace preserved."""
    run = etree.SubElement(parent, qn("w:r"))
    properties = etree.SubElement(run, qn("w:rPr"))
    etree.SubElement(properties, qn("w:sz")).set(_W_VAL, _COMMENT_TEXT_HALF_POINTS)
    etree.SubElement(properties, qn("w:szCs")).set(_W_VAL, _COMMENT_TEXT_HALF_POINTS)
    element = etree.SubElement(run, _W_T)
    element.text = text
    element.set(_XML_SPACE, "preserve")
    return run


def _append_comment(
    root: etree._Element,
    *,
    comment_id: int,
    para_id: str,
    text: str,
    author: str,
    initials: str,
    timestamp: str,
    styled: bool,
) -> etree._Element:
    """Append the ``w:comment`` holding the comment body and return it."""
    element = etree.SubElement(root, qn("w:comment"))
    element.set(_W_ID, str(comment_id))
    element.set(qn("w:author"), author)
    element.set(qn("w:date"), timestamp)
    element.set(qn("w:initials"), initials)

    paragraph = etree.SubElement(element, qn("w:p"))
    paragraph.set(qn("w14:paraId"), para_id)
    paragraph.set(qn("w14:textId"), para_id)

    # The reference mark of the comment's own first line, which is what Word
    # draws in the review pane; styled only when the package defines the style,
    # so the part never carries a dangling w:rStyle.
    marker = etree.SubElement(paragraph, qn("w:r"))
    if styled:
        properties = etree.SubElement(marker, qn("w:rPr"))
        etree.SubElement(properties, qn("w:rStyle")).set(_W_VAL, COMMENT_REFERENCE_STYLE)
    etree.SubElement(marker, qn("w:annotationRef"))

    _text_run(paragraph, text)
    return element


def _append_thread_entry(root: etree._Element, para_id: str) -> etree._Element:
    """Declare the comment as an open, unanswered thread root."""
    entry = etree.SubElement(root, qn("w15:commentEx"))
    entry.set(qn("w15:paraId"), para_id)
    entry.set(qn("w15:done"), "0")
    return entry


def _next_durable_id(root: etree._Element, preferred: str) -> str:
    """Return a durable id no entry of `root` uses, starting from `preferred`.

    `preferred` is the comment's ``w14:paraId``, already unique across the
    package; the walk only guards against a durable id another producer happened
    to write with the same value.
    """
    used = {
        entry.get(_cid("durableId"))
        for entry in root.iter(_cid("commentId"))
        if entry.get(_cid("durableId")) is not None
    }
    candidate = int(preferred, 16)
    for _ in range(len(used) + 1):
        value = f"{candidate:08X}"
        if value not in used:
            return value
        candidate = candidate + 1 if candidate < 0x7FFFFFFF else 1
    raise IdExhausted("no free durable comment id left in commentsIds.xml")


def _append_durable_id(root: etree._Element, para_id: str) -> etree._Element:
    """Give the comment the durable identity Word uses to track it across copies."""
    entry = etree.SubElement(root, _cid("commentId"))
    entry.set(_cid("paraId"), para_id)
    entry.set(_cid("durableId"), _next_durable_id(root, para_id))
    return entry


def _register_author(root: etree._Element, author: str) -> None:
    """Declare `author` in ``people.xml``, once."""
    for person in root.iter(qn("w15:person")):
        if person.get(qn("w15:author")) == author:
            return
    person = etree.SubElement(root, qn("w15:person"))
    person.set(qn("w15:author"), author)
    presence = etree.SubElement(person, qn("w15:presenceInfo"))
    presence.set(qn("w15:providerId"), "None")
    presence.set(qn("w15:userId"), author)


def add_comment_to_doc(
    filepath: str,
    target_text: str,
    comment_text: str,
    author: str = DEFAULT_AUTHOR,
    initials: str = DEFAULT_INITIALS,
) -> dict[str, Any]:
    """Add a comment to a Word document, anchored on the first `target_text`.

    The search runs on the *visible* text of the main story, so a target split
    across several runs, or sitting inside a hyperlink, an insertion or a
    content control, is found the way a reader sees it.  The anchored runs are
    left exactly where they are: a comment is three marks around them, never a
    container, so a target inside a hyperlink stays inside its hyperlink.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place and atomically.
        target_text: text of the document the comment is anchored on.
        comment_text: the comment body.
        author: comment author name.
        initials: author initials.

    Returns:
        A dict with ``success`` true and the comment's details, or ``success``
        false and an ``error`` message when the target cannot be commented.
    """
    path = Path(filepath)
    pkg = DocxPackage.open(path)

    matches = find(pkg, target_text, max_results=1)
    if not matches:
        return {"success": False, "error": f"Target text not found: '{target_text}'"}
    match = matches[0]

    try:
        pieces = resolve(match.paragraph, match.start, match.end)
    except EngineError as exc:
        return {
            "success": False,
            "error": f"Cannot anchor a comment on '{target_text}': {exc}",
        }

    comments_root = pkg.ensure_part(COMMENTS_PARTNAME, CT_COMMENTS, RT_COMMENTS, _COMMENTS_XML)
    extended_root = pkg.ensure_part(
        COMMENTS_EXTENDED_PARTNAME,
        CT_COMMENTS_EXTENDED,
        RT_COMMENTS_EXTENDED,
        _COMMENTS_EXTENDED_XML,
    )
    ids_root = pkg.ensure_part(
        COMMENTS_IDS_PARTNAME, CT_COMMENTS_IDS, RT_COMMENTS_IDS, _COMMENTS_IDS_XML
    )
    people_root = pkg.ensure_part(PEOPLE_PARTNAME, CT_PEOPLE, RT_PEOPLE, _PEOPLE_XML)

    comment_id = next_comment_id(pkg)
    para_id = next_para_id(pkg)

    try:
        wrap(pieces, comment_envelope(comment_id, pkg=pkg))
    except EngineError as exc:
        return {
            "success": False,
            "error": f"Cannot anchor a comment on '{target_text}': {exc}",
        }

    _append_comment(
        comments_root,
        comment_id=comment_id,
        para_id=para_id,
        text=comment_text,
        author=author,
        initials=initials,
        timestamp=_now_iso(),
        styled=character_style_exists(pkg, COMMENT_REFERENCE_STYLE),
    )
    _append_thread_entry(extended_root, para_id)
    _append_durable_id(ids_root, para_id)
    _register_author(people_root, author)

    pkg.save(path)

    shortened = target_text[:50] + ("..." if len(target_text) > 50 else "")
    return {
        "success": True,
        "comment_id": comment_id,
        "author": author,
        "target_text": target_text,
        "comment_text": comment_text,
        "message": f"Added comment #{comment_id} by {author} on text '{shortened}'",
    }
