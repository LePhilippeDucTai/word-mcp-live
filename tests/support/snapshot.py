"""Canonical snapshot of a WordprocessingML package, and structural comparison.

This module is the measuring instrument of the fidelity harness: it turns a
``.docx`` package into a comparable value so that a test can state "nothing
changed except this paragraph" and have that claim verified.

It uses ``zipfile`` and ``lxml`` only -- never ``python-docx`` -- so that it
stays usable to characterise ``python-docx`` itself.

Paragraph index space
---------------------
Three ``paragraph_index`` spaces coexist in this repository (python-docx body
paragraphs, ``//w:p``, ``body//w:p``).  This module deliberately uses a fourth,
the widest one, and calls it the *snapshot space*:

    every ``w:p`` element of a story part, in document order, base 0.

Concretely that is ``story_root.iter(w:p)``: paragraphs inside table cells,
inside block-level ``w:sdt``, and inside text boxes are all indexed, and so are
the paragraphs of headers, footers, footnotes, endnotes and comments (each in
its own story).  Nothing in a story escapes the index, which is what makes
"unchanged except" a meaningful assertion.  This space is *not* the V2 locator
space (``w:p`` of the body, cells excluded); no conversion is attempted here.

Text policy
-----------
Visible text follows the engine policy: ``w:t`` reached from a paragraph
through any element other than ``w:del``, ``w:instrText`` and ``w:delInstrText``
(so ``w:r``, ``w:hyperlink``, ``w:ins``, ``w:sdt``, ``w:smartTag``,
``w:fldSimple`` are all transparent), with ``w:tab`` -> ``\\t`` and
``w:br``/``w:cr`` -> ``\\n``.  ``w:pPr`` and ``w:sdtPr`` are property
containers and are never read as text.  Descent stops at a nested ``w:p``, so
text-box content belongs to its own paragraph and is not counted twice.

Deleted content is not lost: it is captured separately in
:attr:`ParagraphSignature.deleted_runs`, and field instructions in
:attr:`ParagraphSignature.field_instructions`, so that "``w:del`` is never
modified" and "fields stay atomic" remain checkable.

Canonicalisation
----------------
XML parts (``.xml`` and ``.rels``) are compared through C14N 1.0 of the whole
part; every other part is compared byte for byte.  ``[Content_Types].xml`` is
never compared as bytes: it is compared as the set of its ``(part, type)``
entries, since its entry order carries no meaning.  Element fragments (``rPr``)
use *exclusive* C14N so that unused namespace declarations inherited from the
part root do not leak into the signature.

C14N normalises attribute order, empty-element form, the XML declaration and
attribute-value escaping; it does *not* normalise inter-element whitespace nor
namespace prefixes, which therefore stay significant.  This is deliberate: a
serialiser that reindents a part it was not asked to touch has changed it.  A
caller that knowingly accepts such a rewrite lists the part in
``assert_unchanged_except(..., parts=...)``.

Element-by-element coverage
---------------------------
Allowing one paragraph lifts the C14N digest of its whole story part, so the
per-paragraph and per-table signatures are the *only* guard left over the rest
of that part.  They must therefore describe every carrier of meaning, not just
text: ``w:pPr`` and ``w:rPr`` (formatting, numbering, paragraph-mark revision),
the non-textual children of a run (``w:drawing``, ``w:object``, ``w:pict``,
``w:fldChar``, ``w:br``, ...), ``w:tblPr``/``w:trPr``/``w:tcPr`` (table style,
borders, widths) and ``w:sdtPr`` (identity of a content control, whether the
control wraps the paragraph or sits inside it).  Anything left out of these
signatures is a loss the harness cannot see.
"""

from __future__ import annotations

import hashlib
import io
import os
import posixpath
import zipfile
from collections.abc import Collection, Iterator, Mapping
from dataclasses import dataclass
from dataclasses import fields as dataclass_fields
from typing import Any

from lxml import etree

__all__ = [
    "MAIN_STORY",
    "PARAGRAPH_INDEX_SPACE",
    "Diff",
    "FieldChange",
    "ParagraphChange",
    "ParagraphSignature",
    "PartSignature",
    "RunSignature",
    "Snapshot",
    "TableChange",
    "TableSignature",
    "assert_unchanged_except",
    "diff",
    "snapshot",
]

PARAGRAPH_INDEX_SPACE = "story-root//w:p, document order, base 0"
"""Human-readable name of the index space documented in the module docstring."""

MAIN_STORY = "document"
"""Story id of the main document part, used when a bare int key is given."""

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

CONTENT_TYPES_PART = "[Content_Types].xml"

_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml."
STORY_CONTENT_TYPES = frozenset(
    {
        _WML + "document.main+xml",
        _WML + "template.main+xml",
        _WML + "header+xml",
        _WML + "footer+xml",
        _WML + "footnotes+xml",
        _WML + "endnotes+xml",
        _WML + "comments+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    }
)


def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


_P = _w("p")
_PPR = _w("pPr")
_RPR = _w("rPr")
_R = _w("r")
_T = _w("t")
_TAB = _w("tab")
_BR = _w("br")
_CR = _w("cr")
_DEL = _w("del")
_DEL_TEXT = _w("delText")
_INSTR_TEXT = _w("instrText")
_DEL_INSTR_TEXT = _w("delInstrText")
_SDT = _w("sdt")
_SDT_PR = _w("sdtPr")
_TBL = _w("tbl")
_TBL_PR = _w("tblPr")
_TR = _w("tr")
_TR_PR = _w("trPr")
_TC = _w("tc")
_TC_PR = _w("tcPr")
_TBL_GRID = _w("tblGrid")
_GRID_COL = _w("gridCol")
_GRID_SPAN = _w("gridSpan")
_V_MERGE = _w("vMerge")
_VAL = _w("val")
_P_STYLE = _w("pStyle")
_NUM_PR = _w("numPr")
_NUM_ID = _w("numId")
_ILVL = _w("ilvl")
_BOOKMARK_START = _w("bookmarkStart")
_NAME = _w("name")
_ID = _w("id")
_FLD_SIMPLE = _w("fldSimple")
_INSTR = _w("instr")
_FLD_CHAR = _w("fldChar")
_FLD_CHAR_TYPE = _w("fldCharType")

#: Property containers: they hold no document text and no document run.
_PROPERTY_TAGS = frozenset({_PPR, _RPR, _SDT_PR})
#: Subtrees hidden from the visible text of a paragraph.
_HIDDEN_TAGS = frozenset({_DEL, _INSTR_TEXT, _DEL_INSTR_TEXT})
_VISIBLE_SKIP = _PROPERTY_TAGS | _HIDDEN_TAGS
_VISIBLE_TEXT_TAGS = frozenset({_T})
_DELETED_TEXT_TAGS = frozenset({_T, _DEL_TEXT})

#: Run children already represented by :attr:`RunSignature.text` or by
#: :attr:`ParagraphSignature.field_instructions`; every other child is content
#: in its own right and is recorded in :attr:`RunSignature.children`.
_TEXT_CARRYING_TAGS = frozenset({_T, _DEL_TEXT, _INSTR_TEXT, _DEL_INSTR_TEXT})

#: Elements whose presence and identity a paragraph signature must preserve.
_MARKER_TAGS: dict[str, str] = {
    _w("bookmarkStart"): "name",
    _w("bookmarkEnd"): "bookmark-ref",
    _w("commentRangeStart"): "id",
    _w("commentRangeEnd"): "id",
    _w("commentReference"): "id",
    _w("footnoteReference"): "id",
    _w("endnoteReference"): "id",
    _w("permStart"): "id",
    _w("permEnd"): "id",
    _w("moveFromRangeStart"): "name",
    _w("moveFromRangeEnd"): "id",
    _w("moveToRangeStart"): "name",
    _w("moveToRangeEnd"): "id",
    _FLD_CHAR: "fldCharType",
}

#: Element local names counted as tracked revisions, anywhere in the package.
_REVISION_TAGS = frozenset(
    _w(name)
    for name in (
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "rPrChange",
        "pPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "tblGridChange",
        "sectPrChange",
        "numberingChange",
        "cellIns",
        "cellDel",
        "cellMerge",
    )
)

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)


# --------------------------------------------------------------------------
# Value types
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class PartSignature:
    """Fingerprint of one package part."""

    name: str
    kind: str  # "xml" (canonicalised) or "binary" (raw bytes)
    digest: str
    size: int


@dataclass(frozen=True)
class RunSignature:
    """One ``w:r``: its text, its run properties and its non-textual content."""

    text: str
    rpr: str  # exclusive C14N of w:rPr, "" when the run has none
    #: ``(local name, digest)`` of every child that is not text-carrying, in
    #: document order: images, embedded objects, note references, field
    #: characters, breaks, tabs, symbols...  The digest is of the child's
    #: exclusive C14N, so that losing a ``w:drawing`` *and* silently swapping
    #: the image it points at are both visible.
    children: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class ParagraphSignature:
    """Everything a paragraph must keep across a non-degrading edit."""

    story: str
    index: int
    style: str | None
    text: str
    runs: tuple[RunSignature, ...]
    deleted_runs: tuple[RunSignature, ...]
    markers: tuple[tuple[str, str], ...]
    field_instructions: tuple[str, ...]
    num_pr: tuple[str, str] | None
    #: Exclusive C14N of ``w:pPr``, "" when the paragraph has none.  Covers
    #: spacing, justification, borders, ``w:pPrChange`` and the paragraph-mark
    #: run properties -- including the ``w:rPr/w:del`` that marks the paragraph
    #: mark itself as deleted.
    ppr: str = ""
    #: Exclusive C14N of the ``w:sdtPr`` of each content control *inside* the
    #: paragraph, in document order.
    inline_controls: tuple[str, ...] = ()
    #: Same, for each ``w:sdt`` the paragraph is nested in, outermost first.
    #: A block-level control that loses its properties -- or its wrapping --
    #: changes the signature of every paragraph it contains.
    enclosing_controls: tuple[str, ...] = ()

    @property
    def key(self) -> tuple[str, int]:
        return (self.story, self.index)

    @property
    def deleted_text(self) -> str:
        return "".join(run.text for run in self.deleted_runs)


@dataclass(frozen=True)
class TableSignature:
    """Shape of a ``w:tbl`` plus the visible text of its cells."""

    story: str
    index: int
    rows: int
    grid_cols: int
    grid: tuple[tuple[tuple[int, str], ...], ...]  # per row, per cell: span, vMerge
    cells: tuple[str, ...]
    #: Exclusive C14N of ``w:tblPr``: table style, borders, width, layout.
    tbl_pr: str = ""
    #: Exclusive C14N of each ``w:trPr``, in row order ("" when a row has none).
    row_properties: tuple[str, ...] = ()
    #: Exclusive C14N of each ``w:tcPr``, in cell order ("" when a cell has
    #: none).  ``grid`` only extracts span and vertical merge from it; the
    #: digest keeps shading, borders and widths observable too.
    cell_properties: tuple[str, ...] = ()

    @property
    def key(self) -> tuple[str, int]:
        return (self.story, self.index)

    @property
    def structure(self) -> tuple[int, int, tuple[tuple[tuple[int, str], ...], ...]]:
        return (self.rows, self.grid_cols, self.grid)


@dataclass(frozen=True)
class Snapshot:
    """Canonical view of a package.  Compare two of them with :func:`diff`."""

    parts: Mapping[str, PartSignature]
    content_types: frozenset[tuple[str, str]]
    paragraphs: Mapping[tuple[str, int], ParagraphSignature]
    tables: Mapping[tuple[str, int], TableSignature]
    relationships: Mapping[str, frozenset[tuple[str, str]]]
    counters: Mapping[str, int]
    stories: tuple[str, ...]
    story_parts: Mapping[str, str]

    def story_paragraphs(self, story: str) -> tuple[ParagraphSignature, ...]:
        """Paragraphs of one story, in snapshot-space order."""
        return tuple(
            sig for key, sig in sorted(self.paragraphs.items()) if key[0] == story
        )

    def text(self, story: str = MAIN_STORY) -> tuple[str, ...]:
        """Visible text of every paragraph of a story, in order."""
        return tuple(sig.text for sig in self.story_paragraphs(story))


@dataclass(frozen=True)
class FieldChange:
    """One named value that differs between two signatures."""

    field: str
    before: Any
    after: Any


@dataclass(frozen=True)
class ParagraphChange:
    key: tuple[str, int]
    changes: tuple[FieldChange, ...]


@dataclass(frozen=True)
class TableChange:
    key: tuple[str, int]
    changes: tuple[FieldChange, ...]

    @property
    def is_structural(self) -> bool:
        return any(c.field != "cells" for c in self.changes)


@dataclass(frozen=True)
class Diff:
    """Difference between two snapshots.  Empty means "nothing changed"."""

    parts_added: tuple[str, ...] = ()
    parts_removed: tuple[str, ...] = ()
    parts_changed: tuple[str, ...] = ()
    content_types_added: tuple[tuple[str, str], ...] = ()
    content_types_removed: tuple[tuple[str, str], ...] = ()
    paragraphs_added: tuple[tuple[str, int], ...] = ()
    paragraphs_removed: tuple[tuple[str, int], ...] = ()
    paragraphs_changed: tuple[ParagraphChange, ...] = ()
    tables_added: tuple[tuple[str, int], ...] = ()
    tables_removed: tuple[tuple[str, int], ...] = ()
    tables_changed: tuple[TableChange, ...] = ()
    relationships_added: tuple[tuple[str, str, str], ...] = ()
    relationships_removed: tuple[tuple[str, str, str], ...] = ()
    counters_changed: tuple[FieldChange, ...] = ()

    def is_empty(self) -> bool:
        return not any(getattr(self, f.name) for f in _DIFF_FIELDS)

    def describe(self) -> str:
        """Readable, deterministic rendering of the non-empty sections."""
        lines: list[str] = []
        for name in ("parts_added", "parts_removed", "parts_changed"):
            for value in getattr(self, name):
                lines.append(f"{name}: {value}")
        for name in ("content_types_added", "content_types_removed"):
            for part, ctype in getattr(self, name):
                lines.append(f"{name}: {part} -> {ctype}")
        for name in ("paragraphs_added", "paragraphs_removed"):
            for story, index in getattr(self, name):
                lines.append(f"{name}: {story}[{index}]")
        for change in self.paragraphs_changed:
            story, index = change.key
            for item in change.changes:
                lines.append(
                    f"paragraph {story}[{index}].{item.field}: "
                    f"{item.before!r} -> {item.after!r}"
                )
        for name in ("tables_added", "tables_removed"):
            for story, index in getattr(self, name):
                lines.append(f"{name}: {story}[{index}]")
        for change in self.tables_changed:
            story, index = change.key
            for item in change.changes:
                lines.append(
                    f"table {story}[{index}].{item.field}: "
                    f"{item.before!r} -> {item.after!r}"
                )
        for name in ("relationships_added", "relationships_removed"):
            for part, rtype, target in getattr(self, name):
                lines.append(f"{name}: {part} {rtype} -> {target}")
        for item in self.counters_changed:
            lines.append(f"counter {item.field}: {item.before} -> {item.after}")
        return "\n".join(lines)


_DIFF_FIELDS = dataclass_fields(Diff)


# --------------------------------------------------------------------------
# Package reading
# --------------------------------------------------------------------------


def _read_package(
    source: str | os.PathLike[str] | bytes | bytearray,
) -> dict[str, bytes]:
    if isinstance(source, (bytes, bytearray)):
        handle: Any = io.BytesIO(bytes(source))
    else:
        handle = os.fspath(source)
    with zipfile.ZipFile(handle) as archive:
        return {
            info.filename: archive.read(info)
            for info in archive.infolist()
            if not info.is_dir()
        }


def _parse(data: bytes) -> etree._Element | None:
    try:
        return etree.fromstring(data, parser=_PARSER)
    except etree.XMLSyntaxError:
        return None


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _canonical_part(root: etree._Element) -> bytes:
    return etree.tostring(root.getroottree(), method="c14n", with_comments=True)


def _canonical_fragment(element: etree._Element | None) -> str:
    if element is None:
        return ""
    canonical = etree.tostring(
        element, method="c14n", exclusive=True, with_comments=False
    )
    return canonical.decode("utf-8")


def _is_xml_part(name: str) -> bool:
    return name.lower().endswith((".xml", ".rels"))


def _is_rels_part(name: str) -> bool:
    if not name.lower().endswith(".rels"):
        return False
    return posixpath.basename(posixpath.dirname(name)) == "_rels"


def _rels_owner(name: str) -> str:
    """Part described by a ``.rels`` part; ``"/"`` for the package itself."""
    directory, base = posixpath.split(name)
    owner_dir = posixpath.dirname(directory)
    owner = base[: -len(".rels")]
    if not owner:
        return "/"
    return posixpath.join(owner_dir, owner) if owner_dir else owner


def _story_id(part_name: str) -> str:
    if part_name.startswith("word/") and part_name.endswith(".xml"):
        return part_name[len("word/") : -len(".xml")]
    return part_name


def _parse_content_types(
    data: bytes,
) -> tuple[frozenset[tuple[str, str]], dict[str, str], dict[str, str]]:
    """Return the comparable entry set, the defaults and the overrides."""
    entries: set[tuple[str, str]] = set()
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    root = _parse(data)
    if root is None:
        return frozenset(), defaults, overrides
    for element in root:
        if not isinstance(element.tag, str):
            continue
        local = etree.QName(element).localname
        if local == "Default":
            extension = (element.get("Extension") or "").lower()
            ctype = element.get("ContentType") or ""
            defaults[extension] = ctype
            entries.add((f"*.{extension}", ctype))
        elif local == "Override":
            part = (element.get("PartName") or "").lstrip("/")
            ctype = element.get("ContentType") or ""
            overrides[part] = ctype
            entries.add((part, ctype))
    return frozenset(entries), defaults, overrides


def _content_type_of(
    name: str, defaults: Mapping[str, str], overrides: Mapping[str, str]
) -> str | None:
    if name in overrides:
        return overrides[name]
    extension = posixpath.splitext(name)[1].lstrip(".").lower()
    return defaults.get(extension)


def _parse_relationships(root: etree._Element | None) -> frozenset[tuple[str, str]]:
    if root is None:
        return frozenset()
    tag = f"{{{PACKAGE_RELS_NS}}}Relationship"
    return frozenset(
        (element.get("Type") or "", element.get("Target") or "")
        for element in root
        if element.tag == tag
    )


# --------------------------------------------------------------------------
# Story walking
# --------------------------------------------------------------------------


def _iter_elements(
    node: etree._Element, skip: Collection[str]
) -> Iterator[etree._Element]:
    """Pre-order descendants of ``node``, never crossing a nested ``w:p``."""
    for child in node:
        tag = child.tag
        if not isinstance(tag, str):  # comments and processing instructions
            continue
        if tag == _P or tag in skip:
            continue
        yield child
        yield from _iter_elements(child, skip)


def _text_of(
    node: etree._Element, text_tags: Collection[str], skip: Collection[str]
) -> str:
    pieces: list[str] = []
    for child in node:
        tag = child.tag
        if not isinstance(tag, str):
            continue
        if tag == _P or tag in skip:
            continue
        if tag in text_tags:
            pieces.append(child.text or "")
        elif tag == _TAB:
            pieces.append("\t")
        elif tag in (_BR, _CR):
            pieces.append("\n")
        else:
            pieces.append(_text_of(child, text_tags, skip))
    return "".join(pieces)


def _visible_text(node: etree._Element) -> str:
    return _text_of(node, _VISIBLE_TEXT_TAGS, _VISIBLE_SKIP)


def _fragment_digest(element: etree._Element) -> str:
    """Short, stable digest of an element's exclusive C14N."""
    return _sha(_canonical_fragment(element).encode("utf-8"))[:16]


def _run_children(run: etree._Element) -> tuple[tuple[str, str], ...]:
    """Non-textual children of a run: what ``text`` and ``rpr`` cannot show."""
    return tuple(
        (etree.QName(child).localname, _fragment_digest(child))
        for child in run
        if isinstance(child.tag, str)
        and child.tag != _RPR
        and child.tag not in _TEXT_CARRYING_TAGS
    )


def _run_signature(
    run: etree._Element, text_tags: Collection[str], skip: Collection[str]
) -> RunSignature:
    return RunSignature(
        text=_text_of(run, text_tags, skip),
        rpr=_canonical_fragment(run.find(_RPR)),
        children=_run_children(run),
    )


def _enclosing_controls(paragraph: etree._Element) -> tuple[str, ...]:
    """``w:sdtPr`` of every ``w:sdt`` the paragraph is nested in, outermost first."""
    controls = [
        _canonical_fragment(node.find(_SDT_PR))
        for node in paragraph.iterancestors(_SDT)
    ]
    controls.reverse()
    return tuple(controls)


def _bookmark_names(root: etree._Element) -> dict[str, str]:
    return {
        element.get(_ID) or "": element.get(_NAME) or ""
        for element in root.iter(_BOOKMARK_START)
    }


def _marker_key(
    element: etree._Element, kind: str, bookmark_names: Mapping[str, str]
) -> str:
    if kind == "name":
        return element.get(_NAME) or element.get(_ID) or ""
    if kind == "bookmark-ref":
        identifier = element.get(_ID) or ""
        return bookmark_names.get(identifier, identifier)
    if kind == "fldCharType":
        return element.get(_FLD_CHAR_TYPE) or ""
    return element.get(_ID) or ""


def _paragraph_signature(
    story: str, index: int, paragraph: etree._Element, bookmark_names: Mapping[str, str]
) -> ParagraphSignature:
    properties = paragraph.find(_PPR)
    style: str | None = None
    num_pr: tuple[str, str] | None = None
    if properties is not None:
        style_element = properties.find(_P_STYLE)
        if style_element is not None:
            style = style_element.get(_VAL)
        num_element = properties.find(_NUM_PR)
        if num_element is not None:
            num_id = num_element.find(_NUM_ID)
            ilvl = num_element.find(_ILVL)
            num_pr = (
                (num_id.get(_VAL) or "") if num_id is not None else "",
                (ilvl.get(_VAL) or "") if ilvl is not None else "",
            )

    runs: list[RunSignature] = []
    deleted_runs: list[RunSignature] = []
    markers: list[tuple[str, str]] = []
    instructions: list[str] = []
    inline_controls: list[str] = []

    for element in _iter_elements(paragraph, _PROPERTY_TAGS):
        tag = element.tag
        kind = _MARKER_TAGS.get(tag)
        if kind is not None:
            local = etree.QName(element).localname
            markers.append((local, _marker_key(element, kind, bookmark_names)))
        if tag in (_INSTR_TEXT, _DEL_INSTR_TEXT):
            instructions.append(element.text or "")
        elif tag == _FLD_SIMPLE:
            instructions.append(element.get(_INSTR) or "")
        elif tag == _SDT:
            inline_controls.append(_canonical_fragment(element.find(_SDT_PR)))
        elif tag == _R:
            if _is_hidden_run(element, paragraph):
                deleted_runs.append(
                    _run_signature(element, _DELETED_TEXT_TAGS, _PROPERTY_TAGS)
                )
            else:
                runs.append(_run_signature(element, _VISIBLE_TEXT_TAGS, _VISIBLE_SKIP))

    return ParagraphSignature(
        story=story,
        index=index,
        style=style,
        text=_visible_text(paragraph),
        runs=tuple(runs),
        deleted_runs=tuple(deleted_runs),
        markers=tuple(markers),
        field_instructions=tuple(instructions),
        num_pr=num_pr,
        ppr=_canonical_fragment(properties),
        inline_controls=tuple(inline_controls),
        enclosing_controls=_enclosing_controls(paragraph),
    )


def _is_hidden_run(run: etree._Element, paragraph: etree._Element) -> bool:
    """True when ``run`` sits under a ``w:del`` inside ``paragraph``."""
    node = run.getparent()
    while node is not None and node is not paragraph:
        if node.tag == _DEL:
            return True
        node = node.getparent()
    return False


def _table_signature(story: str, index: int, table: etree._Element) -> TableSignature:
    grid_element = table.find(_TBL_GRID)
    grid_cols = 0 if grid_element is None else len(grid_element.findall(_GRID_COL))
    grid: list[tuple[tuple[int, str], ...]] = []
    cells: list[str] = []
    row_properties: list[str] = []
    cell_properties: list[str] = []
    rows = 0
    for row in table:
        if row.tag != _TR:
            continue
        rows += 1
        row_properties.append(_canonical_fragment(row.find(_TR_PR)))
        row_grid: list[tuple[int, str]] = []
        for cell in row:
            if cell.tag != _TC:
                continue
            span = 1
            merge = ""
            properties = cell.find(_TC_PR)
            cell_properties.append(_canonical_fragment(properties))
            if properties is not None:
                span_element = properties.find(_GRID_SPAN)
                if span_element is not None:
                    try:
                        span = int(span_element.get(_VAL) or "1")
                    except ValueError:
                        span = 1
                merge_element = properties.find(_V_MERGE)
                if merge_element is not None:
                    merge = merge_element.get(_VAL) or "continue"
            row_grid.append((span, merge))
            cells.append(
                "\n".join(
                    _visible_text(child) for child in cell if child.tag == _P
                )
            )
        grid.append(tuple(row_grid))
    return TableSignature(
        story=story,
        index=index,
        rows=rows,
        grid_cols=grid_cols,
        grid=tuple(grid),
        cells=tuple(cells),
        tbl_pr=_canonical_fragment(table.find(_TBL_PR)),
        row_properties=tuple(row_properties),
        cell_properties=tuple(cell_properties),
    )


def _count_elements(roots: Mapping[str, etree._Element]) -> dict[str, int]:
    counters = dict.fromkeys(
        (
            "comments",
            "comment_references",
            "revisions",
            "bookmarks",
            "fields",
            "hyperlinks",
            "footnotes",
            "footnote_references",
            "endnotes",
            "endnote_references",
        ),
        0,
    )
    simple = {
        _w("comment"): "comments",
        _w("commentReference"): "comment_references",
        _BOOKMARK_START: "bookmarks",
        _w("hyperlink"): "hyperlinks",
        _w("footnote"): "footnotes",
        _w("footnoteReference"): "footnote_references",
        _w("endnote"): "endnotes",
        _w("endnoteReference"): "endnote_references",
        _FLD_SIMPLE: "fields",
    }
    for name, root in roots.items():
        if _is_rels_part(name):
            continue
        for element in root.iter():
            tag = element.tag
            if not isinstance(tag, str):
                continue
            key = simple.get(tag)
            if key is not None:
                counters[key] += 1
            if tag in _REVISION_TAGS:
                counters["revisions"] += 1
            elif tag == _FLD_CHAR and element.get(_FLD_CHAR_TYPE) == "begin":
                counters["fields"] += 1
    return counters


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------


def snapshot(path_or_bytes: str | os.PathLike[str] | bytes | bytearray) -> Snapshot:
    """Build the canonical snapshot of a ``.docx`` package."""
    raw = _read_package(path_or_bytes)
    content_types, defaults, overrides = _parse_content_types(
        raw.get(CONTENT_TYPES_PART, b"")
    )

    roots: dict[str, etree._Element] = {}
    for name, data in raw.items():
        if name == CONTENT_TYPES_PART or not _is_xml_part(name):
            continue
        root = _parse(data)
        if root is not None:
            roots[name] = root

    parts: dict[str, PartSignature] = {}
    for name in sorted(raw):
        if name == CONTENT_TYPES_PART:
            continue
        data = raw[name]
        root = roots.get(name)
        if root is None:
            parts[name] = PartSignature(name, "binary", _sha(data), len(data))
        else:
            parts[name] = PartSignature(
                name, "xml", _sha(_canonical_part(root)), len(data)
            )

    relationships = {
        _rels_owner(name): _parse_relationships(roots.get(name))
        for name in sorted(raw)
        if _is_rels_part(name)
    }

    paragraphs: dict[tuple[str, int], ParagraphSignature] = {}
    tables: dict[tuple[str, int], TableSignature] = {}
    stories: list[str] = []
    story_parts: dict[str, str] = {}
    for name in sorted(roots):
        if _content_type_of(name, defaults, overrides) not in STORY_CONTENT_TYPES:
            continue
        root = roots[name]
        story = _story_id(name)
        stories.append(story)
        story_parts[story] = name
        bookmark_names = _bookmark_names(root)
        for index, paragraph in enumerate(root.iter(_P)):
            paragraphs[(story, index)] = _paragraph_signature(
                story, index, paragraph, bookmark_names
            )
        for index, table in enumerate(root.iter(_TBL)):
            tables[(story, index)] = _table_signature(story, index, table)

    return Snapshot(
        parts=parts,
        content_types=content_types,
        paragraphs=paragraphs,
        tables=tables,
        relationships=relationships,
        counters=_count_elements(roots),
        stories=tuple(stories),
        story_parts=story_parts,
    )


_PARAGRAPH_FIELDS = (
    "style",
    "text",
    "runs",
    "deleted_runs",
    "markers",
    "field_instructions",
    "num_pr",
    "ppr",
    "inline_controls",
    "enclosing_controls",
)
_TABLE_FIELDS = (
    "rows",
    "grid_cols",
    "grid",
    "cells",
    "tbl_pr",
    "row_properties",
    "cell_properties",
)


def _compare_fields(
    before: Any, after: Any, names: Collection[str]
) -> tuple[FieldChange, ...]:
    return tuple(
        FieldChange(name, getattr(before, name), getattr(after, name))
        for name in names
        if getattr(before, name) != getattr(after, name)
    )


def diff(a: Snapshot, b: Snapshot) -> Diff:
    """Structural difference from snapshot ``a`` to snapshot ``b``."""
    parts_added = tuple(sorted(set(b.parts) - set(a.parts)))
    parts_removed = tuple(sorted(set(a.parts) - set(b.parts)))
    parts_changed = tuple(
        sorted(
            name
            for name in set(a.parts) & set(b.parts)
            if a.parts[name].digest != b.parts[name].digest
        )
    )

    paragraphs_added = tuple(sorted(set(b.paragraphs) - set(a.paragraphs)))
    paragraphs_removed = tuple(sorted(set(a.paragraphs) - set(b.paragraphs)))
    paragraphs_changed: list[ParagraphChange] = []
    for key in sorted(set(a.paragraphs) & set(b.paragraphs)):
        changes = _compare_fields(
            a.paragraphs[key], b.paragraphs[key], _PARAGRAPH_FIELDS
        )
        if changes:
            paragraphs_changed.append(ParagraphChange(key, changes))

    tables_added = tuple(sorted(set(b.tables) - set(a.tables)))
    tables_removed = tuple(sorted(set(a.tables) - set(b.tables)))
    tables_changed: list[TableChange] = []
    for key in sorted(set(a.tables) & set(b.tables)):
        changes = _compare_fields(a.tables[key], b.tables[key], _TABLE_FIELDS)
        if changes:
            tables_changed.append(TableChange(key, changes))

    rels_added: list[tuple[str, str, str]] = []
    rels_removed: list[tuple[str, str, str]] = []
    for owner in sorted(set(a.relationships) | set(b.relationships)):
        before = a.relationships.get(owner, frozenset())
        after = b.relationships.get(owner, frozenset())
        rels_added.extend((owner, t, target) for t, target in sorted(after - before))
        rels_removed.extend((owner, t, target) for t, target in sorted(before - after))

    counters_changed = tuple(
        FieldChange(name, a.counters.get(name, 0), b.counters.get(name, 0))
        for name in sorted(set(a.counters) | set(b.counters))
        if a.counters.get(name, 0) != b.counters.get(name, 0)
    )

    return Diff(
        parts_added=parts_added,
        parts_removed=parts_removed,
        parts_changed=parts_changed,
        content_types_added=tuple(sorted(b.content_types - a.content_types)),
        content_types_removed=tuple(sorted(a.content_types - b.content_types)),
        paragraphs_added=paragraphs_added,
        paragraphs_removed=paragraphs_removed,
        paragraphs_changed=tuple(paragraphs_changed),
        tables_added=tables_added,
        tables_removed=tables_removed,
        tables_changed=tuple(tables_changed),
        relationships_added=tuple(rels_added),
        relationships_removed=tuple(rels_removed),
        counters_changed=counters_changed,
    )


def _normalise_paragraph_keys(
    keys: Collection[tuple[str, int] | int],
) -> set[tuple[str, int]]:
    normalised: set[tuple[str, int]] = set()
    for key in keys:
        if isinstance(key, int):
            normalised.add((MAIN_STORY, key))
        else:
            normalised.add((key[0], key[1]))
    return normalised


def assert_unchanged_except(
    a: Snapshot,
    b: Snapshot,
    paragraphs: Collection[tuple[str, int] | int] = (),
    parts: Collection[str] = (),
    counters: Collection[str] = (),
) -> None:
    """Fail unless every difference between ``a`` and ``b`` was allowed.

    ``paragraphs`` lists the paragraph keys allowed to change, in the snapshot
    index space; a bare int means the main story.  Allowing a paragraph also
    allows the digest of its story part -- a paragraph cannot change without
    its part changing.  ``parts`` lists part names whose digest, relationships
    and added/removed state are allowed to change.

    ``counters`` relaxes a counter **globally**, for the whole package, not
    paragraph by paragraph: ``counters={"revisions"}`` accepts any revision
    count anywhere, including in paragraphs no caller declared.  It is the
    coarsest lever here and should be the last one reached for; what still
    guards those paragraphs is their own signature (``ppr`` carries
    ``w:pPrChange`` and the deleted paragraph mark, ``runs`` carries
    ``w:rPrChange``, ``deleted_runs`` carries ``w:del``), so a relaxed counter
    hides the *total*, never an individual paragraph.

    Table cell text is not re-checked here: cell paragraphs already live in the
    paragraph index space.  Everything else about a table -- row count, grid,
    spans, vertical merges, ``w:tblPr``, ``w:trPr``, ``w:tcPr`` -- must be
    identical unless its story part is listed in ``parts``.
    """
    allowed_paragraphs = _normalise_paragraph_keys(paragraphs)
    allowed_parts = set(parts)
    allowed_parts.update(
        a.story_parts.get(story) or b.story_parts.get(story, "")
        for story, _ in allowed_paragraphs
    )
    allowed_parts.discard("")
    allowed_stories = {
        story
        for story, part in list(a.story_parts.items()) + list(b.story_parts.items())
        if part in set(parts)
    }
    allowed_counters = set(counters)

    delta = diff(a, b)
    problems: list[str] = []

    def keep_part(name: str) -> bool:
        return name not in allowed_parts

    for name in delta.parts_added:
        if keep_part(name):
            problems.append(f"part added: {name}")
    for name in delta.parts_removed:
        if keep_part(name):
            problems.append(f"part removed: {name}")
    for name in delta.parts_changed:
        if keep_part(name):
            problems.append(f"part changed: {name}")

    for part, ctype in delta.content_types_added:
        if keep_part(part):
            problems.append(f"content type added: {part} -> {ctype}")
    for part, ctype in delta.content_types_removed:
        if keep_part(part):
            problems.append(f"content type removed: {part} -> {ctype}")

    for key in delta.paragraphs_added:
        if key not in allowed_paragraphs:
            problems.append(f"paragraph added: {key[0]}[{key[1]}]")
    for key in delta.paragraphs_removed:
        if key not in allowed_paragraphs:
            problems.append(f"paragraph removed: {key[0]}[{key[1]}]")
    for change in delta.paragraphs_changed:
        if change.key in allowed_paragraphs:
            continue
        for item in change.changes:
            problems.append(
                f"paragraph {change.key[0]}[{change.key[1]}].{item.field}: "
                f"{item.before!r} -> {item.after!r}"
            )

    for key in delta.tables_added:
        if key[0] not in allowed_stories:
            problems.append(f"table added: {key[0]}[{key[1]}]")
    for key in delta.tables_removed:
        if key[0] not in allowed_stories:
            problems.append(f"table removed: {key[0]}[{key[1]}]")
    for change in delta.tables_changed:
        if change.key[0] in allowed_stories:
            continue
        for item in change.changes:
            if item.field == "cells":
                continue
            problems.append(
                f"table {change.key[0]}[{change.key[1]}].{item.field}: "
                f"{item.before!r} -> {item.after!r}"
            )

    for owner, rtype, target in delta.relationships_added:
        if keep_part(owner):
            problems.append(f"relationship added: {owner} {rtype} -> {target}")
    for owner, rtype, target in delta.relationships_removed:
        if keep_part(owner):
            problems.append(f"relationship removed: {owner} {rtype} -> {target}")

    for item in delta.counters_changed:
        if item.field not in allowed_counters:
            problems.append(
                f"counter {item.field}: {item.before} -> {item.after}"
            )

    if problems:
        raise AssertionError(
            "document changed outside the allowed set:\n  "
            + "\n  ".join(problems)
        )
