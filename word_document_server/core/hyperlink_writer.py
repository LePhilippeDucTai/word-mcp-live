"""Adding, listing and removing hyperlinks, through the OOXML engine.

A hyperlink is a ``w:hyperlink`` element wrapping runs, plus -- for an external
target -- one relationship of the part the runs live in.  The two must stay in
step, which is the whole difficulty: a relationship allocated by hand collides
with an existing ``rId`` on the next save, and a link removed without its
relationship leaves a target nothing points at.

Both directions are therefore delegated: :meth:`DocxPackage.add_external_rel`
allocates (and deduplicates) the relationship,
:func:`~word_document_server.engine.format.wrap` builds the envelope around the
resolved range, and :func:`~word_document_server.engine.format.unwrap` takes it
apart again, leaving the very same runs -- text, properties and markers
untouched -- where the link used to be.  A relationship is dropped only once no
reference to it is left in the part.

Paragraph addressing
--------------------
`paragraph_index` is a V2 index: the rank of the paragraph among the top-level
paragraphs of the main story, table cells and text boxes excluded (see
:mod:`word_document_server.engine.find`).  A paragraph without a V2 index
cannot be addressed by index; its text can still be addressed by content.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from lxml import etree

from word_document_server.engine.errors import EngineError
from word_document_server.engine.find import Match, find, iter_paragraphs

# The V2 index space has exactly one definition (D-009); it lives in find.py and
# is reused here rather than restated, so the two can never drift apart.
from word_document_server.engine.find import _v2_index_map as v2_index_map
from word_document_server.engine.format import (
    apply_rpr,
    character_style_exists,
    unwrap,
)
from word_document_server.engine.format import hyperlink as hyperlink_envelope
from word_document_server.engine.format import wrap
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import resolve
from word_document_server.engine.textmodel import segments
from word_document_server.engine.xmlns import R, qn

__all__ = [
    "HYPERLINK_STYLE",
    "add_hyperlink_to_doc",
    "list_hyperlinks_in_doc",
    "remove_hyperlink_from_doc",
]

#: Style id Word gives the character style of a hyperlink.
HYPERLINK_STYLE = "Hyperlink"

#: Direct formatting used when the package does not define :data:`HYPERLINK_STYLE`
#: -- Word's own hyperlink colour, and the underline that goes with it.
_HYPERLINK_DIRECT_FORMAT = {"color": "0563C1", "underline": "single"}

_W_P = qn("w:p")
_W_HYPERLINK = qn("w:hyperlink")
_R_ID = qn("r:id")
_R_PREFIX = f"{{{R}}}"
_W_ANCHOR = qn("w:anchor")


def _v2_paragraph_count(root: etree._Element) -> int:
    """Number of paragraphs of `root` that have a V2 index."""
    return len(v2_index_map(iter_paragraphs(root)))


def _matches_in_paragraph(matches: list[Match], index: int) -> list[Match]:
    return [match for match in matches if match.index == index]


def _owning_paragraph(link: etree._Element) -> etree._Element | None:
    """The ``w:p`` `link` belongs to, however deeply it is nested."""
    return next(iter(link.iterancestors(_W_P)), None)


def _link_text(link: etree._Element) -> str:
    """The visible text of a hyperlink, read from its paragraph's flow.

    Reading it through :func:`~word_document_server.engine.textmodel.segments`
    rather than off the ``w:t`` children is what makes a link whose text sits
    inside a tracked insertion, a content control or a field result read the
    way it is displayed.
    """
    paragraph = _owning_paragraph(link)
    if paragraph is None:
        return ""
    return "".join(
        segment.text
        for segment in segments(paragraph)
        if any(container is link for container in segment.containers)
    )


def _link_target(pkg: DocxPackage, story_root: etree._Element, link: etree._Element) -> str:
    """The address `link` points at: an external URL, or ``#anchor``."""
    rId = link.get(_R_ID)
    if rId:
        target = pkg.rel_target(story_root, rId)
        return target if isinstance(target, str) else str(target.partname)
    anchor = link.get(_W_ANCHOR)
    return f"#{anchor}" if anchor else ""


def _reference_count(root: etree._Element, rId: str) -> int:
    """How many ``r:`` references of `root` still point at `rId`."""
    count = 0
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        count += sum(
            1
            for name, value in element.attrib.items()
            if value == rId and name.startswith(_R_PREFIX)
        )
    return count


def _drop_unused_rel(part: Part, story_root: etree._Element, rId: str) -> bool:
    """Remove the relationship `rId` from `part` if nothing references it any more.

    Returns:
        Whether the relationship was removed.
    """
    if _reference_count(story_root, rId) > 0 or rId not in part.rels:
        return False
    del part.rels[rId]
    return True


def add_hyperlink_to_doc(
    filepath: str,
    text: str,
    url: str,
    paragraph_index: int | None = None,
) -> dict[str, Any]:
    """Turn the first occurrence of `text` into a hyperlink pointing at `url`.

    The matched runs are moved into a ``w:hyperlink`` as they are: their own
    formatting survives, and only the hyperlink look is added on top -- the
    ``Hyperlink`` character style when the package defines it, its colour and
    underline written directly when it does not.  The external relationship is
    deduplicated, so linking ten paragraphs to one address adds one
    relationship.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place and atomically.
        text: the text to turn into a link.
        url: the address the link points at.
        paragraph_index: restrict the search to that V2 paragraph index.

    Returns:
        A dict with ``success`` true and the link's details, or ``success``
        false and an ``error`` message.
    """
    path = Path(filepath)
    pkg = DocxPackage.open(path)
    story_root = pkg.document

    if paragraph_index is not None:
        count = _v2_paragraph_count(story_root)
        if paragraph_index < 0 or paragraph_index >= count:
            return {
                "success": False,
                "error": f"Paragraph index {paragraph_index} out of range (0-{count - 1})",
            }
        matches = _matches_in_paragraph(find(pkg, text), paragraph_index)
    else:
        matches = find(pkg, text, max_results=1)

    if not matches:
        return {"success": False, "error": f"Text not found: '{text}'"}
    match = matches[0]

    try:
        pieces = resolve(match.paragraph, match.start, match.end)
        if character_style_exists(pkg, HYPERLINK_STYLE):
            apply_rpr(pieces, {"char_style": HYPERLINK_STYLE}, pkg=pkg)
        else:
            apply_rpr(pieces, _HYPERLINK_DIRECT_FORMAT, pkg=pkg)
        rId = pkg.add_external_rel(match.paragraph, url, RT.HYPERLINK)
        wrap(pieces, hyperlink_envelope(rId))
    except EngineError as exc:
        return {"success": False, "error": f"Cannot link '{text}': {exc}"}

    pkg.save(path)
    return {
        "success": True,
        "text": text,
        "url": url,
        "relationship_id": rId,
        "paragraph_index": match.index,
        "message": f"Added hyperlink to '{text}' pointing to {url}",
    }


def list_hyperlinks_in_doc(filepath: str) -> dict[str, Any]:
    """List every hyperlink of the document, in reading order.

    Every story is walked -- body, headers, footers, notes -- because a link in
    a footer is as real as one in the body.

    Args:
        filepath: path to the ``.docx`` file.  Nothing is modified.

    Returns:
        A dict with ``success`` true, ``hyperlinks`` and ``total_hyperlinks``.
    """
    pkg = DocxPackage.open(Path(filepath))
    links: list[dict[str, Any]] = []
    for story, root in pkg.stories():
        index_map = v2_index_map(iter_paragraphs(root))
        for link in root.iter(_W_HYPERLINK):
            paragraph = _owning_paragraph(link)
            links.append(
                {
                    "text": _link_text(link),
                    "url": _link_target(pkg, root, link),
                    "relationship_id": link.get(_R_ID),
                    "anchor": link.get(_W_ANCHOR),
                    "story": story,
                    "paragraph_index": None if paragraph is None else index_map.get(paragraph),
                }
            )
    return {"success": True, "hyperlinks": links, "total_hyperlinks": len(links)}


def remove_hyperlink_from_doc(
    filepath: str,
    text: str = "",
    paragraph_index: int | None = None,
) -> dict[str, Any]:
    """Remove the hyperlinks matching `text`, keeping their text in place.

    The runs the link held stay exactly where they were, with their formatting:
    only the ``w:hyperlink`` envelope disappears.  A relationship left pointing
    at nothing is removed too; one still used by another link is kept.

    Args:
        filepath: path to the ``.docx`` file, rewritten in place and atomically.
        text: the visible text of the link to remove; every link of the
            document is removed when it is empty and `paragraph_index` is given.
        paragraph_index: restrict the removal to that V2 paragraph index.

    Returns:
        A dict with ``success`` true, ``removed`` and ``relationships_removed``,
        or ``success`` false and an ``error`` message when nothing matched.
    """
    path = Path(filepath)
    pkg = DocxPackage.open(path)
    story_root = pkg.document
    index_map = v2_index_map(iter_paragraphs(story_root))

    if not text and paragraph_index is None:
        return {
            "success": False,
            "error": "text or paragraph_index is required to remove a hyperlink",
        }
    if paragraph_index is not None:
        count = len(index_map)
        if paragraph_index < 0 or paragraph_index >= count:
            return {
                "success": False,
                "error": f"Paragraph index {paragraph_index} out of range (0-{count - 1})",
            }

    targeted: list[etree._Element] = []
    for link in story_root.iter(_W_HYPERLINK):
        paragraph = _owning_paragraph(link)
        if paragraph is None:
            continue
        if paragraph_index is not None and index_map.get(paragraph) != paragraph_index:
            continue
        if text and text not in _link_text(link):
            continue
        targeted.append(link)

    if not targeted:
        return {"success": False, "error": f"No hyperlink found for text: '{text}'"}

    removed: list[dict[str, Any]] = []
    candidate_ids: list[str] = []
    for link in targeted:
        rId = link.get(_R_ID)
        removed.append(
            {
                "text": _link_text(link),
                "url": _link_target(pkg, story_root, link),
                "relationship_id": rId,
            }
        )
        unwrap(link)
        if rId:
            candidate_ids.append(rId)

    document_part = pkg.document_part
    dropped = sorted(
        {rId for rId in candidate_ids if _drop_unused_rel(document_part, story_root, rId)}
    )
    pkg.save(path)
    return {
        "success": True,
        "text": text,
        "removed": removed,
        "removed_count": len(removed),
        "relationships_removed": dropped,
        "message": f"Removed {len(removed)} hyperlink(s)",
    }
