"""Append the content of one package to another, by importing its elements.

The naive way to merge two documents is to read the source's text and write it
back into the target.  That is what this repository used to do, and it loses
everything text is not: run properties, images, fields, bookmarks, tracked
changes, table geometry, numbering.  This module does the opposite -- it moves
the *elements* themselves.  Each block of the source body is deep-copied into
the target body, and everything the copy points at outside itself is made to
resolve again in its new package:

relationship references
    every attribute in the relationship namespace (``r:id``, ``r:embed``,
    ``r:link``, ...) is resolved against the source's ``document.xml.rels``, the
    part it names is imported into the target (media are copied byte for byte,
    with their own relationships preserved under their original rIds), and the
    attribute is rewritten with the rId the target hands back.  An external
    relationship -- a hyperlink -- is recreated as an external relationship of
    the target.
styles
    a ``w:pStyle``, ``w:rStyle`` or ``w:tblStyle`` the target does not define is
    copied over from the source styles part, together with the closure of its
    ``w:basedOn``, ``w:link`` and ``w:next`` references, so the copied style is
    not left standing on a definition that does not exist.  A style id the
    target already defines is *never* overwritten: the target's own definition
    wins and a warning says so.
numbering
    a ``w:numPr`` points at a ``w:num`` of the source numbering part, whose id
    means something else -- or nothing -- in the target.  Every referenced
    ``w:num`` is therefore copied under a freshly allocated ``w:numId``, its
    ``w:abstractNum`` under a freshly allocated ``w:abstractNumId``, and the
    copied paragraphs are rewritten to the new ids.  Two references to the same
    list stay one list.
annotation ids
    a bookmark or a revision id that already exists in the target is renumbered
    (the pairs ``w:bookmarkStart``/``w:bookmarkEnd`` and the ``…RangeStart``/
    ``…RangeEnd`` of a move stay consistent), and so is a colliding
    ``w14:paraId``.  Ids that do not collide are left alone: an id is part of
    the document's identity and is not churned for nothing.

What this layer refuses, and what it drops
------------------------------------------
Comments and notes are **refused**, not dropped: importing them means
renumbering ``word/comments.xml``, ``commentsExtended.xml``, ``footnotes.xml``
and ``endnotes.xml`` and keeping their side-car parts in step, and doing half of
that silently is how a merge corrupts a document.  A source whose imported
content carries a ``w:commentReference``, a ``w:footnoteReference`` or a
``w:endnoteReference`` raises :class:`MergeRefused`, before the target is
touched.

Headers and footers are section-level, not block-level: the target keeps its
own, and the source's are ignored with a warning.  A section break imported from
the source loses its ``w:headerReference`` and ``w:footerReference`` children
for the same reason -- keeping them would point the target at parts that are not
there.  Picture bullets (``w:lvlPicBulletId``) are dropped from an imported
numbering definition, again with a warning.

The target's own final ``w:sectPr`` is never touched: imported blocks are
inserted *before* it, so the target keeps its page setup and its own
headers and footers.

Failure is not transactional.  :func:`append_document` mutates the target in
place, and a failure past the refusal check can leave it half-imported.  That is
survivable in practice because the caller works on a package opened in memory
and only writes it out once the merge has succeeded -- which is what
``tools.document_tools.merge_documents`` does.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree

from word_document_server.engine.errors import EngineError, IdExhausted, PackageError
from word_document_server.engine.ids import (
    ANNOTATION_ID_TAGS,
    MAX_DECIMAL_ID,
    PARA_ID_MAX,
)
from word_document_server.engine.package import COMMENTS_CONTENT_TYPE, DocxPackage
from word_document_server.engine.xmlns import R, W, qn

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

__all__ = ["MergeRefused", "Report", "append_document"]


_W_BODY = qn("w:body")
_W_SECT_PR = qn("w:sectPr")
_W_P = qn("w:p")
_W_R = qn("w:r")
_W_BR = qn("w:br")
_W_TYPE = qn("w:type")
_W_VAL = qn("w:val")
_W_ID = qn("w:id")
_W_NUM = qn("w:num")
_W_NUM_ID = qn("w:numId")
_W_NUM_PR = qn("w:numPr")
_W_ABSTRACT_NUM = qn("w:abstractNum")
_W_ABSTRACT_NUM_ID = qn("w:abstractNumId")
_W_STYLE = qn("w:style")
_W_STYLE_ID = qn("w:styleId")
_W14_PARA_ID = qn("w14:paraId")
_R_PREFIX = "{" + R + "}"

#: Empty parts created when the target has no styles or no numbering part yet.
_EMPTY_STYLES = f'<w:styles xmlns:w="{W}"/>'
_EMPTY_NUMBERING = f'<w:numbering xmlns:w="{W}"/>'

#: Elements whose ``w:val`` names a style id, wherever they appear -- in a
#: paragraph, in a run, in table properties, or inside an imported numbering
#: level (``w:pStyle``, ``w:styleLink``, ``w:numStyleLink``).
_STYLE_REFERENCE_TAGS: tuple[str, ...] = (
    "w:pStyle",
    "w:rStyle",
    "w:tblStyle",
    "w:numStyleLink",
    "w:styleLink",
)

#: Elements sharing the annotation ``w:id`` space, as
#: :data:`~word_document_server.engine.ids.ANNOTATION_ID_TAGS` lists them, plus
#: the range ends that must be renumbered with their start and the ``…Change``
#: family Word allocates from the same counter.  The list is explicit rather
#: than "every element carrying a ``w:id``" on purpose: ``w:sdtPr``,
#: ``w:permStart`` and the note references carry a ``w:id`` of their own space,
#: and renumbering one of those would break the reference it stands for.
_ANNOTATION_ID_TAGS: tuple[str, ...] = ANNOTATION_ID_TAGS + (
    "w:bookmarkEnd",
    "w:cellDel",
    "w:cellIns",
    "w:cellMerge",
    "w:commentRangeEnd",
    "w:customXmlDelRangeEnd",
    "w:customXmlDelRangeStart",
    "w:customXmlInsRangeEnd",
    "w:customXmlInsRangeStart",
    "w:customXmlMoveFromRangeEnd",
    "w:customXmlMoveFromRangeStart",
    "w:customXmlMoveToRangeEnd",
    "w:customXmlMoveToRangeStart",
    "w:moveFromRangeEnd",
    "w:moveToRangeEnd",
    "w:numberingChange",
    "w:pPrChange",
    "w:rPrChange",
    "w:sectPrChange",
    "w:tblGridChange",
    "w:tblPrChange",
    "w:tblPrExChange",
    "w:tcPrChange",
    "w:trPrChange",
)

#: What a source may not carry into a merge, and the word used to refuse it.
_REFUSED_REFERENCES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("comments", ("w:commentReference", "w:commentRangeStart", "w:commentRangeEnd")),
    ("footnotes", ("w:footnoteReference",)),
    ("endnotes", ("w:endnoteReference",)),
)


class MergeRefused(EngineError):
    """The source carries content this layer will not import rather than lose.

    Attributes:
        reason: stable machine-readable reason -- ``"comments"``,
            ``"footnotes"`` or ``"endnotes"``.  Callers branch on `reason`; the
            message is for humans and may change.
    """

    def __init__(self, reason: str, message: str | None = None) -> None:
        super().__init__(message if message is not None else reason)
        self.reason = reason


@dataclass(frozen=True)
class Report:
    """What :func:`append_document` did, and what it could not do faithfully.

    Attributes:
        blocks: number of body-level blocks imported from the source -- the
            page-break paragraph :func:`append_document` may insert of its own
            accord is not one of them.
        styles: style ids copied into the target styles part, in copy order.
        numbering: source ``w:numId`` -> target ``w:numId``, for every list
            definition the import had to renumber.
        parts: names of the parts created in the target (media and whatever they
            relate to), in creation order.
        warnings: human-readable lines describing what the import could not
            carry over -- headers and footers, picture bullets, a style the
            target already defines differently, a duplicated bookmark name.  An
            empty tuple means nothing was lost.
    """

    blocks: int = 0
    styles: tuple[str, ...] = ()
    numbering: dict[int, int] = field(default_factory=dict)
    parts: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()


# -- small helpers ----------------------------------------------------------------------


def _body_of(document: etree._Element) -> etree._Element:
    """Return the ``w:body`` of a ``w:document`` root."""
    body = document.find(_W_BODY)
    if body is None:
        raise PackageError("the main document part has no w:body")
    return body


def _final_sect_pr(body: etree._Element) -> etree._Element | None:
    """Return the body-level ``w:sectPr`` closing `body`, or ``None``."""
    last = body[-1] if len(body) else None
    return last if last is not None and last.tag == _W_SECT_PR else None


def _content_blocks(body: etree._Element) -> list[etree._Element]:
    """Return every child of `body` but the final ``w:sectPr``."""
    sect_pr = _final_sect_pr(body)
    return [child for child in body if child is not sect_pr and isinstance(child.tag, str)]


def _elements(nodes: Iterable[etree._Element]) -> Iterator[etree._Element]:
    """Yield `nodes` and all their descendants, skipping comments and PIs."""
    for node in nodes:
        for element in node.iter():
            if isinstance(element.tag, str):
                yield element


def _tagged(nodes: Iterable[etree._Element], tags: Iterable[str]) -> Iterator[etree._Element]:
    """Yield every descendant of `nodes` (themselves included) whose tag is in `tags`."""
    wanted = frozenset(qn(tag) for tag in tags)
    for element in _elements(nodes):
        if element.tag in wanted:
            yield element


def _canonical(element: etree._Element) -> bytes:
    """Exclusive canonical form of `element`, for comparing two definitions."""
    return etree.tostring(element, method="c14n", exclusive=True)


def _int_or_none(raw: str | None) -> int | None:
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _related_root(pkg: DocxPackage, reltype: str) -> etree._Element | None:
    """Root of the part the document part relates to under `reltype`, if any."""
    for relationship in pkg.document_part.rels.values():
        if relationship.reltype == reltype and not relationship.is_external:
            return pkg.root_of(relationship.target_part)
    return None


def _refuse_unsupported(blocks: Iterable[etree._Element]) -> None:
    """Raise :class:`MergeRefused` if `blocks` reference comments or notes."""
    for reason, tags in _REFUSED_REFERENCES:
        found = next(_tagged(blocks, tags), None)
        if found is None:
            continue
        raise MergeRefused(
            reason,
            f"the source cannot be appended: its content carries {reason}. "
            f"Importing them means renumbering their part and the side-car parts "
            f"Word keeps in step with it, which this layer does not do; merging "
            f"anyway would silently drop them.",
        )


def _annotation_ids(nodes: Iterable[etree._Element]) -> set[int]:
    """Every annotation ``w:id`` used under `nodes`."""
    found: set[int] = set()
    for element in _tagged(nodes, _ANNOTATION_ID_TAGS):
        value = _int_or_none(element.get(_W_ID))
        if value is not None:
            found.add(value)
    return found


def _para_ids(nodes: Iterable[etree._Element]) -> set[int]:
    """Every ``w14:paraId`` used under `nodes`, as integers."""
    found: set[int] = set()
    for element in _elements(nodes):
        raw = element.get(_W14_PARA_ID)
        if raw is None:
            continue
        try:
            found.add(int(raw, 16))
        except ValueError:
            continue
    return found


def _annotated_roots(pkg: DocxPackage) -> list[etree._Element]:
    """Roots of `pkg` that may carry annotation ids: its stories and its comments."""
    roots = [root for _, root in pkg.stories()]
    comments = pkg.find_part("/word/comments.xml")
    if comments is not None and comments.content_type == COMMENTS_CONTENT_TYPE:
        roots.append(pkg.root_of(comments))
    return roots


def _collision_map(used: set[int], incoming: set[int], *, ceiling: int, space: str) -> dict[int, int]:
    """Map each id of `incoming` that `used` already takes to a free one.

    Free means free of both sets: an id handed out must collide neither with the
    target nor with an id of the source that is being kept as it is.

    Raises:
        IdExhausted: if the space runs out before every collision is resolved.
    """
    taken = used | incoming
    candidate = (max(taken) + 1) if taken else 1
    remap: dict[int, int] = {}
    for old in sorted(used & incoming):
        if candidate > ceiling:
            raise IdExhausted(f"no {space} id left at or below {ceiling}")
        remap[old] = candidate
        candidate += 1
    return remap


# -- the import ------------------------------------------------------------------------


class _Import:
    """One ``append_document`` call: the state it threads through its steps."""

    def __init__(self, target: DocxPackage, source: DocxPackage) -> None:
        self.target = target
        self.source = source
        self.warnings: list[str] = []
        self.styles_copied: list[str] = []
        self.parts_created: list[str] = []
        self.num_map: dict[int, int] = {}
        self._unresolved_nums: set[int] = set()
        self._abstract_map: dict[int, int] = {}
        self._imported_parts: dict[Part, Part] = {}
        self._style_queue: list[str] = []
        self._num_queue: list[int] = []
        self._styles_seen: set[str] = set()
        #: ``w:style`` copies, appended to the target styles part once the
        #: worklists are drained; the numbering map is applied to them too,
        #: because a style may carry a ``w:numPr`` of its own.
        self.style_elements: list[etree._Element] = []

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    # -- relationships ------------------------------------------------------------------

    def rewrite_relationships(self, blocks: list[etree._Element]) -> None:
        """Point every relationship reference of `blocks` at the target's own rels."""
        owner = self.source.document_part
        for element in _elements(blocks):
            for name, value in list(element.attrib.items()):
                if not name.startswith(_R_PREFIX):
                    continue
                element.set(name, self._rebind(owner, value))

    def _rebind(self, owner: Part, rId: str) -> str:
        """Recreate the relationship `rId` of `owner` on the target document part."""
        relationship = owner.rels.get(rId)
        if relationship is None:
            raise PackageError(
                f"the source references the relationship {rId!r}, which "
                f"{owner.partname} does not declare"
            )
        if relationship.is_external:
            return self.target.add_external_rel(
                self.target.document_part, relationship.target_ref, relationship.reltype
            )
        imported = self._import_part(relationship.target_part)
        return self.target.document_part.relate_to(imported, relationship.reltype)

    def _import_part(self, source_part: Part) -> Part:
        """Copy `source_part` (and what it relates to) into the target package.

        The copy is made from ``blob``, so a part the engine holds unparsed comes
        out byte for byte identical.  Its own relationships are recreated *under
        their original rIds*, because the references that resolve them live
        inside the bytes we just copied and are not rewritten.
        """
        cached = self._imported_parts.get(source_part)
        if cached is not None:
            return cached

        blob = source_part.blob
        partname = str(source_part.partname)
        existing = self.target.find_part(partname)
        # A leaf part (an image) whose bytes and content type the target already
        # has under the same name *is* that part: merging the same document twice
        # must not leave two copies of the same picture behind.
        if (
            existing is not None
            and not source_part.rels
            and existing.content_type == source_part.content_type
            and existing.blob == blob
        ):
            self._imported_parts[source_part] = existing
            return existing

        created = Part(
            PackURI(self._free_partname(partname)),
            source_part.content_type,
            blob,
            self.target.package,
        )
        self._imported_parts[source_part] = created
        self.parts_created.append(str(created.partname))
        for relationship in source_part.rels.values():
            if relationship.is_external:
                created.rels.add_relationship(
                    relationship.reltype, relationship.target_ref, relationship.rId, True
                )
            else:
                child = self._import_part(relationship.target_part)
                created.rels.add_relationship(relationship.reltype, child, relationship.rId)
        return created

    def _free_partname(self, partname: str) -> str:
        """Return `partname`, or the first numbered variant the target is free of."""
        if self.target.find_part(partname) is None:
            return partname
        head, dot, extension = partname.rpartition(".")
        if not dot:
            head, extension = partname, ""
        suffix = f".{extension}" if dot else ""
        stem = head.rstrip("0123456789") or head
        index = 2
        while True:
            candidate = f"{stem}{index}{suffix}"
            if self.target.find_part(candidate) is None:
                return candidate
            index += 1

    # -- section-level content the target keeps its own of ------------------------------

    def drop_header_footer_references(self, blocks: list[etree._Element]) -> None:
        """Remove the header/footer references of any imported section break."""
        if any(
            relationship.reltype in (RT.HEADER, RT.FOOTER)
            for relationship in self.source.document_part.rels.values()
        ):
            self.warn(
                "the headers and footers of the source are not imported; the target "
                "keeps its own"
            )
        dropped = [
            element
            for element in _tagged(blocks, ("w:headerReference", "w:footerReference"))
        ]
        for element in dropped:
            element.getparent().remove(element)
        if dropped:
            self.warn(
                "a section break imported from the source referenced headers or "
                "footers; those references were dropped, so the imported section "
                "uses the headers and footers of the target"
            )

    # -- styles and numbering -----------------------------------------------------------

    def resolve_definitions(self, blocks: list[etree._Element]) -> None:
        """Copy the styles and the numbering definitions `blocks` depend on.

        Styles and numbering reference each other -- a style may carry a
        ``w:numPr``, an ``w:abstractNum`` may carry a ``w:pStyle`` -- so both
        worklists are drained until neither has anything left.
        """
        self._style_queue.extend(
            value
            for element in _tagged(blocks, _STYLE_REFERENCE_TAGS)
            if (value := element.get(_W_VAL))
        )
        self._num_queue.extend(
            value
            for element in _tagged(blocks, ("w:numId",))
            if (value := _int_or_none(element.get(_W_VAL))) is not None
        )
        while self._style_queue or self._num_queue:
            while self._style_queue:
                self._copy_style(self._style_queue.pop(0))
            while self._num_queue:
                self._copy_numbering(self._num_queue.pop(0))

        if self.style_elements:
            styles = _related_root(self.target, RT.STYLES)
            if styles is None:
                styles = self.target.ensure_part(
                    "/word/styles.xml", CT.WML_STYLES, RT.STYLES, _EMPTY_STYLES
                )
            for element in self.style_elements:
                styles.append(element)

    def _style_element(self, root: etree._Element | None, style_id: str) -> etree._Element | None:
        if root is None:
            return None
        for element in root.iter(_W_STYLE):
            if element.get(_W_STYLE_ID) == style_id:
                return element
        return None

    def _copy_style(self, style_id: str) -> None:
        if style_id in self._styles_seen:
            return
        self._styles_seen.add(style_id)

        in_source = self._style_element(_related_root(self.source, RT.STYLES), style_id)
        in_target = self._style_element(_related_root(self.target, RT.STYLES), style_id)
        if in_target is not None:
            if in_source is not None and _canonical(in_source) != _canonical(in_target):
                self.warn(
                    f"style {style_id!r} is defined in both documents with different "
                    f"content; the definition of the target is kept, so the imported "
                    f"text may not look exactly as it did in the source"
                )
            return
        if in_source is None:
            self.warn(
                f"style {style_id!r} is referenced by the source but defined in "
                f"neither document; the imported text falls back to the default style"
            )
            return

        element = copy.deepcopy(in_source)
        self.style_elements.append(element)
        self.styles_copied.append(style_id)
        for reference in _tagged([element], ("w:basedOn", "w:link", "w:next")):
            value = reference.get(_W_VAL)
            if value:
                self._style_queue.append(value)
        for reference in _tagged([element], ("w:numId",)):
            value = _int_or_none(reference.get(_W_VAL))
            if value is not None:
                self._num_queue.append(value)

    def _numbering_roots(self) -> tuple[etree._Element | None, etree._Element]:
        source = _related_root(self.source, RT.NUMBERING)
        target = _related_root(self.target, RT.NUMBERING)
        if target is None:
            target = self.target.ensure_part(
                "/word/numbering.xml", CT.WML_NUMBERING, RT.NUMBERING, _EMPTY_NUMBERING
            )
        return source, target

    def _copy_numbering(self, num_id: int) -> None:
        if num_id <= 0 or num_id in self.num_map or num_id in self._unresolved_nums:
            return
        source, target = self._numbering_roots()
        definition = None
        if source is not None:
            for element in source.iter(_W_NUM):
                if _int_or_none(element.get(_W_NUM_ID)) == num_id:
                    definition = element
                    break
        if definition is None:
            self._unresolved_nums.add(num_id)
            self.warn(
                f"the source references the list numId {num_id}, which its numbering "
                f"part does not define; the imported paragraphs lose their numbering "
                f"rather than joining an unrelated list of the target"
            )
            return

        element = copy.deepcopy(definition)
        new_num_id = self._free_number(target, _W_NUM, _W_NUM_ID)
        element.set(_W_NUM_ID, str(new_num_id))
        abstract = element.find(_W_ABSTRACT_NUM_ID)
        if abstract is not None:
            old_abstract = _int_or_none(abstract.get(_W_VAL))
            if old_abstract is not None:
                abstract.set(_W_VAL, str(self._copy_abstract(source, target, old_abstract)))
        self._insert_numbering_child(target, element, before_tags=("w:numIdMacAtCleanup",))
        self.num_map[num_id] = new_num_id

    def _copy_abstract(
        self, source: etree._Element | None, target: etree._Element, abstract_id: int
    ) -> int:
        """Copy the ``w:abstractNum`` `abstract_id` into `target`, once."""
        known = self._abstract_map.get(abstract_id)
        if known is not None:
            return known
        definition = None
        if source is not None:
            for element in source.iter(_W_ABSTRACT_NUM):
                if _int_or_none(element.get(_W_ABSTRACT_NUM_ID)) == abstract_id:
                    definition = element
                    break
        new_id = self._free_number(target, _W_ABSTRACT_NUM, _W_ABSTRACT_NUM_ID)
        self._abstract_map[abstract_id] = new_id
        if definition is None:
            self.warn(
                f"the source references the abstract list {abstract_id}, which its "
                f"numbering part does not define; the imported list has no levels"
            )
            element = etree.SubElement(target, _W_ABSTRACT_NUM)
            element.set(_W_ABSTRACT_NUM_ID, str(new_id))
        else:
            element = copy.deepcopy(definition)
            element.set(_W_ABSTRACT_NUM_ID, str(new_id))
            bullets = list(_tagged([element], ("w:lvlPicBulletId",)))
            for bullet in bullets:
                bullet.getparent().remove(bullet)
            if bullets:
                self.warn(
                    "an imported list used picture bullets; the picture is not "
                    "imported and those levels fall back to their text bullet"
                )
            for reference in _tagged([element], _STYLE_REFERENCE_TAGS):
                value = reference.get(_W_VAL)
                if value:
                    self._style_queue.append(value)
        self._insert_numbering_child(target, element, before_tags=("w:num", "w:numIdMacAtCleanup"))
        return new_id

    @staticmethod
    def _free_number(root: etree._Element, tag: str, attribute: str) -> int:
        """Return an id above every value of `attribute` on `tag` under `root`."""
        highest = 0
        for element in root.iter(tag):
            value = _int_or_none(element.get(attribute))
            if value is not None and value > highest:
                highest = value
        if highest >= MAX_DECIMAL_ID:
            raise IdExhausted(f"no {tag} id left at or below {MAX_DECIMAL_ID}")
        return highest + 1

    @staticmethod
    def _insert_numbering_child(
        root: etree._Element, element: etree._Element, *, before_tags: Iterable[str]
    ) -> None:
        """Place `element` under `root`, respecting the order ``w:numbering`` requires."""
        wanted = tuple(qn(tag) for tag in before_tags)
        for child in root:
            if child.tag in wanted:
                child.addprevious(element)
                return
        root.append(element)

    def apply_numbering_map(self, scopes: Iterable[etree._Element]) -> None:
        """Rewrite every ``w:numId`` of `scopes` to the id it got in the target."""
        doomed: list[etree._Element] = []
        for element in _tagged(scopes, ("w:numId",)):
            value = _int_or_none(element.get(_W_VAL))
            if value is None or value <= 0:
                continue
            mapped = self.num_map.get(value)
            if mapped is not None:
                element.set(_W_VAL, str(mapped))
            elif value in self._unresolved_nums:
                doomed.append(element)
        for element in doomed:
            properties = element.getparent()
            if properties is not None and properties.tag == _W_NUM_PR:
                properties.getparent().remove(properties)
            elif properties is not None:  # pragma: no cover - defensive
                properties.remove(element)

    # -- identifiers --------------------------------------------------------------------

    def renumber_ids(self, blocks: list[etree._Element]) -> None:
        """Give the copies fresh ids wherever they would collide with the target."""
        roots = _annotated_roots(self.target)

        annotations = _collision_map(
            _annotation_ids(roots),
            _annotation_ids(blocks),
            ceiling=MAX_DECIMAL_ID,
            space="annotation",
        )
        if annotations:
            for element in _tagged(blocks, _ANNOTATION_ID_TAGS):
                value = _int_or_none(element.get(_W_ID))
                if value in annotations:
                    element.set(_W_ID, str(annotations[value]))

        paragraphs = _collision_map(
            _para_ids(roots), _para_ids(blocks), ceiling=PARA_ID_MAX, space="w14:paraId"
        )
        if paragraphs:
            for element in _elements(blocks):
                raw = element.get(_W14_PARA_ID)
                if raw is None:
                    continue
                try:
                    value = int(raw, 16)
                except ValueError:
                    continue
                if value in paragraphs:
                    element.set(_W14_PARA_ID, f"{paragraphs[value]:08X}")

        self._warn_on_duplicate_bookmarks(roots, blocks)

    def _warn_on_duplicate_bookmarks(
        self, roots: list[etree._Element], blocks: list[etree._Element]
    ) -> None:
        """Warn about a bookmark name the target already uses.

        The name is *not* rewritten: it is what a cross-reference, a hyperlink
        anchor or a field points at, in documents this layer cannot see.  Word
        keeps the first of two same-named bookmarks, so the caller has to know.
        """
        existing = {
            element.get(qn("w:name"))
            for element in _tagged(roots, ("w:bookmarkStart",))
            if element.get(qn("w:name"))
        }
        clashing = sorted(
            {
                name
                for element in _tagged(blocks, ("w:bookmarkStart",))
                if (name := element.get(qn("w:name"))) in existing
            }
        )
        for name in clashing:
            self.warn(
                f"the bookmark name {name!r} exists in both documents; Word keeps "
                f"the one of the target, so a reference to the imported one may "
                f"point at the wrong place"
            )


def append_document(
    target: DocxPackage,
    source: DocxPackage,
    page_break: bool = False,
) -> Report:
    """Append the body of `source` to the body of `target`, in place.

    Every block of the source body -- paragraphs, tables, content controls --
    is deep-copied before the target's final ``w:sectPr``, and everything the
    copies point at is rebound to the target: relationships (media parts are
    copied, external links recreated), styles, numbering definitions, colliding
    annotation ids.  See the module docstring for the whole contract, including
    what is refused and what is dropped with a warning.

    `page_break` inserts a page-break paragraph between the target's existing
    content and the imported blocks.  It does nothing when the target body is
    empty: there is nothing to break away from.

    `target` is modified in place; nothing is written to disk.  A failure raised
    past the initial refusal check may leave `target` half-imported, so a caller
    that has to keep a document intact should save only once the call returned.

    Returns:
        A :class:`Report` describing what was imported and what was lost.

    Raises:
        MergeRefused: `source` carries comments, footnotes or endnotes.  Nothing
            has been modified when this is raised.
        PackageError: `source` references a relationship its document part does
            not declare, or either package has no ``w:body``.
        IdExhausted: an id space filled up while renumbering.
    """
    source_body = _body_of(source.document)
    target_body = _body_of(target.document)
    source_blocks = _content_blocks(source_body)
    _refuse_unsupported(source_blocks)

    job = _Import(target, source)
    blocks = [copy.deepcopy(block) for block in source_blocks]
    job.drop_header_footer_references(blocks)
    job.rewrite_relationships(blocks)
    job.resolve_definitions(blocks)
    job.apply_numbering_map([*blocks, *job.style_elements])
    job.renumber_ids(blocks)

    sect_pr = _final_sect_pr(target_body)
    if page_break and _content_blocks(target_body) and blocks:
        _insert(target_body, _new_page_break(target_body), sect_pr)
    for block in blocks:
        _insert(target_body, block, sect_pr)

    return Report(
        blocks=len(blocks),
        styles=tuple(job.styles_copied),
        numbering=dict(job.num_map),
        parts=tuple(job.parts_created),
        warnings=tuple(job.warnings),
    )


def _insert(
    body: etree._Element, element: etree._Element, sect_pr: etree._Element | None
) -> None:
    """Put `element` at the end of `body`, but before its final ``w:sectPr``."""
    if sect_pr is None:
        body.append(element)
    else:
        sect_pr.addprevious(element)


def _new_page_break(body: etree._Element) -> etree._Element:
    """Build a paragraph holding a single page break, declared in `body`'s tree."""
    paragraph = etree.SubElement(body, _W_P)
    run = etree.SubElement(paragraph, _W_R)
    etree.SubElement(run, _W_BR).set(_W_TYPE, "page")
    return paragraph
