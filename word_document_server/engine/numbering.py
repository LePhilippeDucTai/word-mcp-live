"""List numbering: the ``numbering.xml`` part, its definitions, and ``w:numPr``.

A bulleted or numbered paragraph carries no bullet and no number.  It carries a
``w:numPr`` naming a ``w:num`` of ``word/numbering.xml``, which names a
``w:abstractNum``, which holds the nine ``w:lvl`` definitions that say what the
label looks like at each depth (ECMA-376 §17.9).  Everything a reader sees about
a list therefore lives in a part the paragraph only points at.

Why this module exists
----------------------
The tools of this repository used to write ``w:numId`` 1 for a bullet and 2 for
a number, hard-coded.  Those two ids mean *whatever the document says they
mean*: on a package python-docx created from its own template they happen to
resolve (its template ships nine definitions), on a document written by Word
they point at an unrelated list the author already had -- so the new paragraphs
silently join it and continue its count -- and on a package with no numbering
part at all they resolve to nothing, which is a document Word repairs on open.

Every list this module hands out is therefore a *new* definition: a fresh
``w:abstractNum`` with its own ``w:nsid``, and a fresh ``w:num`` pointing at it.
Two calls never return the same id, so two lists created one after the other
count independently, and no existing list of the document is ever joined,
renumbered or rewritten.

Restarting
----------
Word restarts a list by pointing a *second* ``w:num`` at the same
``w:abstractNum`` and overriding each level's start value
(``w:lvlOverride``/``w:startOverride``).  :func:`restart_list` does exactly
that and returns the new ``w:numId``: the paragraphs that keep the old id keep
counting, the ones given the new id start again at the level's own start value.

What a caller works with
------------------------
The package argument is either a :class:`~word_document_server.engine.package.DocxPackage`
or a python-docx ``Document``.  Both are accepted on purpose: the V2 tools hold
the former and the older ``utils.document_utils`` tools hold the latter, and a
numbering id is only unique if *both* allocate from the same part of the same
package.  Paragraph arguments follow the rest of the engine: a ``w:p`` element
or a python-docx ``Paragraph``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from docx.oxml.parser import parse_xml
from lxml import etree

from word_document_server.engine.errors import IdExhausted, PackageError
from word_document_server.engine.format import set_ppr_child
from word_document_server.engine.ids import MAX_DECIMAL_ID
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.xmlns import W, qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph

__all__ = [
    "LIST_KINDS",
    "MAX_LEVELS",
    "NUMBERING_PART_NAME",
    "ListKind",
    "apply_list",
    "create_list_definition",
    "list_definitions",
    "numbering_root",
    "paragraph_list_info",
    "restart_list",
]

#: Part name of the numbering part, when this module has to create it.
NUMBERING_PART_NAME = "/word/numbering.xml"

#: Content of a numbering part that defines nothing yet.
_EMPTY_NUMBERING = f'<w:numbering xmlns:w="{W}"/>'

#: The list shapes :func:`create_list_definition` knows how to write.
LIST_KINDS: tuple[str, ...] = ("bullet", "decimal", "multilevel")

ListKind = Literal["bullet", "decimal", "multilevel"]

#: Word numbers nine levels, ``w:ilvl`` 0 to 8, and refuses a tenth.
MAX_LEVELS = 9

_W_ABSTRACT_NUM = qn("w:abstractNum")
_W_ABSTRACT_NUM_ID = qn("w:abstractNumId")
_W_ILVL = qn("w:ilvl")
_W_LVL = qn("w:lvl")
_W_LVL_JC = qn("w:lvlJc")
_W_LVL_OVERRIDE = qn("w:lvlOverride")
_W_LVL_TEXT = qn("w:lvlText")
_W_MULTI_LEVEL_TYPE = qn("w:multiLevelType")
_W_NSID = qn("w:nsid")
_W_NUM = qn("w:num")
_W_NUM_FMT = qn("w:numFmt")
_W_NUM_ID = qn("w:numId")
_W_NUM_PR = qn("w:numPr")
_W_PPR = qn("w:pPr")
_W_START = qn("w:start")
_W_START_OVERRIDE = qn("w:startOverride")
_W_VAL = qn("w:val")

#: Children of ``w:numbering``, in the order the schema requires.  A new
#: ``w:abstractNum`` goes before the first ``w:num``, a new ``w:num`` before
#: ``w:numIdMacAtCleanup``; both go last when neither is there.
_AFTER_ABSTRACT_NUM = ("w:num", "w:numIdMacAtCleanup")
_AFTER_NUM = ("w:numIdMacAtCleanup",)

#: Children of ``w:numPr``, in schema order.  ``w:numberingChange`` and ``w:ins``
#: belong to a tracked numbering change and are never written here, but they are
#: listed so that setting the ids on a paragraph that carries one keeps them
#: after ``w:numId`` instead of before it.
_NUMPR_ORDER = ("w:ilvl", "w:numId", "w:numberingChange", "w:ins")

#: The bullet glyphs Word cycles through, with the font each one needs.
_BULLETS: tuple[tuple[str, str], ...] = (
    ("", "Symbol"),
    ("o", "Courier New"),
    ("", "Wingdings"),
)

#: The number formats Word cycles through for a plain numbered list.
_NUMBER_FORMATS: tuple[str, ...] = ("decimal", "lowerLetter", "lowerRoman")

#: Indent of level 0, in twentieths of a point, and the step per level.  Word's
#: own defaults: each level is one tab further in, with a hanging indent that
#: puts the label in the margin it opens up.
_INDENT_STEP = 720
_HANGING = 360

#: An odd stride, so that the ``w:nsid`` values of two definitions created in a
#: row are far apart the way Word's random ones are.  ``w:nsid`` identifies a
#: list across documents; it only has to be unique within the part, and it has to
#: be reproducible, which rules out a random one.
_NSID_BASE = 0x10000000
_NSID_STRIDE = 0x01000193


# --------------------------------------------------------------------------------------
# The part
# --------------------------------------------------------------------------------------


def _document_part(pkg: DocxPackage | Document) -> Part:
    """The main document part of `pkg`, which may be either kind of handle.

    Raises:
        TypeError: if `pkg` is neither a :class:`DocxPackage` nor a python-docx
            ``Document``.
    """
    if isinstance(pkg, DocxPackage):
        return pkg.document_part
    part = getattr(pkg, "part", None)
    if isinstance(part, Part):
        return part
    raise TypeError(
        f"expected a DocxPackage or a python-docx Document, got {type(pkg).__name__}"
    )


def numbering_root(pkg: DocxPackage | Document, *, create: bool = False) -> etree._Element | None:
    """Return the root ``w:numbering`` element, creating the part on demand.

    The part is found through the document part's ``numbering`` relationship
    rather than by name, because that is what Word follows; a package that keeps
    its numbering under another part name is read all the same.  When `create`
    is true and no such relationship exists, an empty numbering part is added to
    the package and related to the document part -- and if a part already sits
    under :data:`NUMBERING_PART_NAME` without being related, it is related rather
    than duplicated.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
        create: add the part when the package has none.

    Returns:
        The live root element, or ``None`` when the package has no numbering part
        and `create` is false.

    Raises:
        PackageError: if the numbering part is not loaded as live XML, which
            would make every edit made here vanish at save time.
    """
    part = _document_part(pkg)
    for relationship in part.rels.values():
        if relationship.reltype == RT.NUMBERING and not relationship.is_external:
            return _live_root(relationship.target_part)
    if not create:
        return None

    package = part.package
    for existing in package.iter_parts():
        if str(existing.partname) == NUMBERING_PART_NAME:
            part.relate_to(existing, RT.NUMBERING)
            return _live_root(existing)

    element = parse_xml(_EMPTY_NUMBERING.encode("utf-8"))
    created = XmlPart(PackURI(NUMBERING_PART_NAME), CT.WML_NUMBERING, element, package)
    part.relate_to(created, RT.NUMBERING)
    return element


def _live_root(part: Part) -> etree._Element:
    """The live root of the numbering part, refusing a blob-backed one."""
    if not isinstance(part, XmlPart):
        raise PackageError(
            f"part {part.partname} holds the numbering but is not loaded as live XML "
            f"({part.content_type}); edits made on it would be dropped at save time"
        )
    return part.element


def _require_root(pkg: DocxPackage | Document) -> etree._Element:
    """The numbering root, or a :class:`PackageError` if the package has none."""
    root = numbering_root(pkg)
    if root is None:
        raise PackageError("the package has no numbering part, so it defines no list")
    return root


def _ensure_root(pkg: DocxPackage | Document) -> etree._Element:
    """The numbering root, created if the package has none."""
    root = numbering_root(pkg, create=True)
    if root is None:  # unreachable: with create=True the part is added or it raises
        raise PackageError("the numbering part could not be created")
    return root


# --------------------------------------------------------------------------------------
# Ids
# --------------------------------------------------------------------------------------


def _int_or_none(value: str | None) -> int | None:
    """`value` as an int, or ``None`` when it is absent or not a number."""
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _free_id(root: etree._Element, tag: str, attribute: str) -> int:
    """An id above every value of `attribute` on `tag`, so it collides with none.

    Raises:
        IdExhausted: if the space is full.
    """
    highest = 0
    for element in root.iter(tag):
        value = _int_or_none(element.get(attribute))
        if value is not None and value > highest:
            highest = value
    if highest >= MAX_DECIMAL_ID:
        raise IdExhausted(f"no {tag} id left at or below {MAX_DECIMAL_ID}")
    return highest + 1


def _free_nsid(root: etree._Element, abstract_id: int) -> str:
    """An unused ``w:nsid`` value, as the 8 hexadecimal digits Word writes."""
    used = {
        (element.get(_W_VAL) or "").upper()
        for element in root.iter(_W_NSID)
    }
    value = (_NSID_BASE + abstract_id * _NSID_STRIDE) & 0x7FFFFFFF
    # At most len(used) candidates can be taken, so one of the next len(used) + 1
    # is free; the loop is bounded rather than trusting that.
    for _ in range(len(used) + 1):
        candidate = f"{value:08X}"
        if candidate not in used:
            return candidate
        value = (value + 1) & 0x7FFFFFFF
    raise IdExhausted("no free w:nsid left in the numbering part")


# --------------------------------------------------------------------------------------
# Writing the definitions
# --------------------------------------------------------------------------------------


def _sub(parent: etree._Element, tag: str, **attributes: str) -> etree._Element:
    """Append a child to `parent`.  Built in place, so it inherits the prefixes."""
    child = etree.SubElement(parent, qn(tag))
    for name, value in attributes.items():
        child.set(qn(name.replace("_", ":", 1)), value)
    return child


def _insert_child(
    root: etree._Element, element: etree._Element, before_tags: tuple[str, ...]
) -> None:
    """Move `element` under `root`, before the first child of `before_tags`."""
    wanted = tuple(qn(tag) for tag in before_tags)
    for child in root:
        if child.tag in wanted:
            child.addprevious(element)
            return
    root.append(element)


def _level_label(kind: str, level: int) -> tuple[str, str]:
    """The ``w:numFmt`` and ``w:lvlText`` of one level of a `kind` list."""
    if kind == "bullet":
        return "bullet", _BULLETS[level % len(_BULLETS)][0]
    if kind == "decimal":
        return _NUMBER_FORMATS[level % len(_NUMBER_FORMATS)], f"%{level + 1}."
    # multilevel: the label spells the whole path, which is what makes it an
    # outline -- 1., 1.1., 1.1.1.
    return "decimal", "".join(f"%{i + 1}." for i in range(level + 1))


def _write_level(abstract: etree._Element, kind: str, level: int) -> None:
    """Append the ``w:lvl`` of `level` to `abstract`, in schema order."""
    number_format, text = _level_label(kind, level)
    lvl = _sub(abstract, "w:lvl", w_ilvl=str(level))
    _sub(lvl, "w:start", w_val="1")
    _sub(lvl, "w:numFmt", w_val=number_format)
    _sub(lvl, "w:lvlText", w_val=text)
    _sub(lvl, "w:lvlJc", w_val="left")
    properties = _sub(lvl, "w:pPr")
    _sub(
        properties,
        "w:ind",
        w_left=str(_INDENT_STEP * (level + 1)),
        w_hanging=str(_HANGING),
    )
    if kind == "bullet":
        font = _BULLETS[level % len(_BULLETS)][1]
        run_properties = _sub(lvl, "w:rPr")
        _sub(run_properties, "w:rFonts", w_ascii=font, w_hAnsi=font, w_hint="default")


def create_list_definition(
    pkg: DocxPackage | Document,
    kind: ListKind = "bullet",
    levels: int = MAX_LEVELS,
) -> int:
    """Define a new list in the package and return the ``w:numId`` to apply.

    The definition is new in both halves: a ``w:abstractNum`` carrying the level
    definitions under a freshly allocated ``w:abstractNumId`` and its own
    ``w:nsid``, and a ``w:num`` pointing at it under a freshly allocated
    ``w:numId``.  No existing definition is read, joined or modified, so a list
    created here counts on its own.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
            Its numbering part is created if it has none.
        kind: ``bullet`` (glyphs), ``decimal`` (one counter per level: 1., a., i.)
            or ``multilevel`` (an outline: 1., 1.1., 1.1.1.).
        levels: how many levels to define, from 1 to :data:`MAX_LEVELS`.  A level
            a paragraph uses but the definition does not describe has no label at
            all in Word, so the default defines the full ladder.

    Returns:
        The ``w:numId`` to pass to :func:`apply_list`.

    Raises:
        ValueError: if `kind` is not one of :data:`LIST_KINDS`, or if `levels` is
            outside 1..:data:`MAX_LEVELS`.
        IdExhausted: if the numbering part has no free id left.
        PackageError: if the numbering part cannot be edited.
    """
    if kind not in LIST_KINDS:
        raise ValueError(f"unknown list kind {kind!r}; known: {', '.join(LIST_KINDS)}")
    if not 1 <= levels <= MAX_LEVELS:
        raise ValueError(f"'levels' must be between 1 and {MAX_LEVELS}, got {levels}")

    root = _ensure_root(pkg)
    abstract_id = _free_id(root, _W_ABSTRACT_NUM, _W_ABSTRACT_NUM_ID)
    num_id = _free_id(root, _W_NUM, _W_NUM_ID)

    abstract = etree.SubElement(root, _W_ABSTRACT_NUM)
    abstract.set(_W_ABSTRACT_NUM_ID, str(abstract_id))
    _sub(abstract, "w:nsid", w_val=_free_nsid(root, abstract_id))
    if levels == 1:
        multi_level_type = "singleLevel"
    elif kind == "multilevel":
        multi_level_type = "multilevel"
    else:
        multi_level_type = "hybridMultilevel"
    _sub(abstract, "w:multiLevelType", w_val=multi_level_type)
    for level in range(levels):
        _write_level(abstract, kind, level)
    _insert_child(root, abstract, _AFTER_ABSTRACT_NUM)

    num = etree.SubElement(root, _W_NUM)
    num.set(_W_NUM_ID, str(num_id))
    _sub(num, "w:abstractNumId", w_val=str(abstract_id))
    _insert_child(root, num, _AFTER_NUM)
    return num_id


def restart_list(pkg: DocxPackage | Document, num_id: int) -> int:
    """Return a new ``w:numId`` that restarts the list `num_id` counts.

    The new ``w:num`` points at the same ``w:abstractNum`` -- the labels are the
    ones the list already had -- and overrides the start value of every level it
    defines, which is how Word writes "restart numbering".  The paragraphs left
    on `num_id` keep counting; the ones given the returned id start again.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.
        num_id: the list to restart.

    Returns:
        The new ``w:numId``.

    Raises:
        PackageError: if the package has no numbering part, if it defines no
            ``w:num`` under `num_id`, or if that one names no ``w:abstractNum``.
        IdExhausted: if the numbering part has no free id left.
    """
    root = _require_root(pkg)
    source = _num_element(root, num_id)
    abstract_id = _int_or_none(_attribute(source.find(_W_ABSTRACT_NUM_ID), _W_VAL))
    if abstract_id is None:
        raise PackageError(
            f"the list numId {num_id} names no abstract definition; it cannot be restarted"
        )

    new_id = _free_id(root, _W_NUM, _W_NUM_ID)
    num = etree.SubElement(root, _W_NUM)
    num.set(_W_NUM_ID, str(new_id))
    _sub(num, "w:abstractNumId", w_val=str(abstract_id))
    for level, start in _level_starts(root, abstract_id).items():
        override = _sub(num, "w:lvlOverride", w_ilvl=str(level))
        _sub(override, "w:startOverride", w_val=str(start))
    _insert_child(root, num, _AFTER_NUM)
    return new_id


def _num_element(root: etree._Element, num_id: int) -> etree._Element:
    """The ``w:num`` of `num_id`.

    Raises:
        PackageError: if the part defines no such list.
    """
    for element in root.iter(_W_NUM):
        if _int_or_none(element.get(_W_NUM_ID)) == num_id:
            return element
    raise PackageError(f"the numbering part defines no list numId {num_id}")


def _abstract_element(root: etree._Element, abstract_id: int) -> etree._Element | None:
    """The ``w:abstractNum`` of `abstract_id`, or ``None``."""
    for element in root.iter(_W_ABSTRACT_NUM):
        if _int_or_none(element.get(_W_ABSTRACT_NUM_ID)) == abstract_id:
            return element
    return None


def _level_starts(root: etree._Element, abstract_id: int) -> dict[int, int]:
    """The start value of every level `abstract_id` defines, by level."""
    abstract = _abstract_element(root, abstract_id)
    if abstract is None:
        return {}
    starts: dict[int, int] = {}
    for lvl in abstract.iter(_W_LVL):
        level = _int_or_none(lvl.get(_W_ILVL))
        if level is None:
            continue
        start = lvl.find(_W_START)
        starts[level] = 1 if start is None else (_int_or_none(start.get(_W_VAL)) or 1)
    return starts


# --------------------------------------------------------------------------------------
# Paragraphs
# --------------------------------------------------------------------------------------


def _ranked_child(parent: etree._Element, tag: str) -> etree._Element:
    """Find or create one child of ``w:numPr``, at its rank in the schema."""
    qualified = qn(tag)
    existing = parent.find(qualified)
    if existing is not None:
        return existing
    # Built under the parent, then moved: an element created in place inherits
    # the namespace prefixes already declared instead of carrying its own.
    child = etree.SubElement(parent, qualified)
    _insert_child(parent, child, _NUMPR_ORDER[_NUMPR_ORDER.index(tag) + 1 :])
    return child


def apply_list(paragraph: etree._Element | Paragraph, num_id: int, level: int = 0) -> None:
    """Make `paragraph` an item of the list `num_id`, at `level`.

    The ``w:numPr`` is written at its rank in ``w:pPr`` (before ``w:pBdr``, after
    ``w:pStyle``), which is where the schema puts it -- appending it at the end,
    as the older tools did, produces a ``w:pPr`` Word has to repair.  A paragraph
    that already belonged to a list is moved to this one rather than gaining a
    second ``w:numPr``.

    Args:
        paragraph: a ``w:p`` element or a python-docx ``Paragraph``.
        num_id: a ``w:numId`` the numbering part defines, as
            :func:`create_list_definition` returns.
        level: the ``w:ilvl``, 0 to :data:`MAX_LEVELS` - 1.

    Raises:
        TypeError: if `paragraph` is not a paragraph.
        ValueError: if `num_id` is below 1 or `level` is out of range.
    """
    if num_id < 1:
        raise ValueError(f"'num_id' must be 1 or more, got {num_id}")
    if not 0 <= level < MAX_LEVELS:
        raise ValueError(f"'level' must be between 0 and {MAX_LEVELS - 1}, got {level}")
    numbering = set_ppr_child(paragraph, "w:numPr", {})
    if numbering is None:  # unreachable: set_ppr_child only removes when attrs is None
        raise PackageError("w:numPr could not be created on the paragraph")
    _ranked_child(numbering, "w:ilvl").set(_W_VAL, str(level))
    _ranked_child(numbering, "w:numId").set(_W_VAL, str(num_id))


def paragraph_list_info(paragraph: etree._Element | Paragraph) -> dict[str, Any] | None:
    """What list `paragraph` belongs to, or ``None`` when it carries none.

    Only *direct* numbering (``w:pPr/w:numPr``) is reported, for the reason
    :func:`~word_document_server.engine.inspect._list_level` gives: numbering
    inherited from a paragraph style would have to be resolved through the style
    chain, and a wrong answer would read exactly like a right one.

    Args:
        paragraph: a ``w:p`` element or a python-docx ``Paragraph``.

    Returns:
        ``{"num_id": int | None, "level": int}`` -- `num_id` is ``None`` when the
        ``w:numPr`` carries no readable ``w:numId``, and `level` defaults to 0,
        which is what Word assumes for a missing ``w:ilvl``.

    Raises:
        TypeError: if `paragraph` is not a paragraph.
    """
    element = _paragraph_element(paragraph)
    properties = element.find(_W_PPR)
    if properties is None:
        return None
    numbering = properties.find(_W_NUM_PR)
    if numbering is None:
        return None
    level = _int_or_none(_attribute(numbering.find(_W_ILVL), _W_VAL))
    num_id = _int_or_none(_attribute(numbering.find(_W_NUM_ID), _W_VAL))
    return {"num_id": num_id, "level": 0 if level is None else level}


def _attribute(element: etree._Element | None, name: str) -> str | None:
    """Read an attribute off an element that may not be there."""
    return None if element is None else element.get(name)


def _paragraph_element(paragraph: etree._Element | Paragraph) -> etree._Element:
    """Coerce `paragraph` to a ``w:p`` element, as the rest of the engine does.

    Same contract as
    :func:`~word_document_server.engine.format.set_ppr_child`, which
    :func:`apply_list` goes through: the reading side of this module must accept
    exactly what its writing side accepts.
    """
    element = paragraph
    if not isinstance(element, etree._Element):
        for attribute in ("_p", "_element"):
            candidate = getattr(element, attribute, None)
            if isinstance(candidate, etree._Element):
                element = candidate
                break
    if not isinstance(element, etree._Element) or element.tag != qn("w:p"):
        raise TypeError(
            f"expected a w:p element or a python-docx paragraph, got {type(paragraph).__name__}"
        )
    return element


# --------------------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------------------


def list_definitions(pkg: DocxPackage | Document) -> list[dict[str, Any]]:
    """Describe every list the package defines, by ``w:numId``.

    Args:
        pkg: the package, as a :class:`DocxPackage` or a python-docx ``Document``.

    Returns:
        One entry per ``w:num``, sorted by `num_id`: `num_id`,
        `abstract_num_id` (``None`` when the ``w:num`` names none),
        `multi_level_type`, `levels` -- one ``{"level", "format", "text",
        "start"}`` per ``w:lvl`` of the abstract definition, in level order --
        and `overrides`, one ``{"level", "start"}`` per ``w:startOverride``, which
        is how a restarted list differs from the one it restarts.  The list is
        empty when the package has no numbering part.
    """
    root = numbering_root(pkg)
    if root is None:
        return []
    definitions: list[dict[str, Any]] = []
    for num in root.iter(_W_NUM):
        num_id = _int_or_none(num.get(_W_NUM_ID))
        if num_id is None:
            continue
        abstract_id = _int_or_none(_attribute(num.find(_W_ABSTRACT_NUM_ID), _W_VAL))
        abstract = None if abstract_id is None else _abstract_element(root, abstract_id)
        definitions.append(
            {
                "num_id": num_id,
                "abstract_num_id": abstract_id,
                "multi_level_type": (
                    None
                    if abstract is None
                    else _attribute(abstract.find(_W_MULTI_LEVEL_TYPE), _W_VAL)
                ),
                "levels": [] if abstract is None else _levels_of(abstract),
                "overrides": _overrides_of(num),
            }
        )
    definitions.sort(key=lambda entry: entry["num_id"])
    return definitions


def _levels_of(abstract: etree._Element) -> list[dict[str, Any]]:
    """The levels `abstract` defines, in level order."""
    levels: list[dict[str, Any]] = []
    for lvl in abstract.iter(_W_LVL):
        level = _int_or_none(lvl.get(_W_ILVL))
        if level is None:
            continue
        start = _int_or_none(_attribute(lvl.find(_W_START), _W_VAL))
        levels.append(
            {
                "level": level,
                "format": _attribute(lvl.find(_W_NUM_FMT), _W_VAL),
                "text": _attribute(lvl.find(_W_LVL_TEXT), _W_VAL),
                "start": 1 if start is None else start,
                "justification": _attribute(lvl.find(_W_LVL_JC), _W_VAL),
            }
        )
    levels.sort(key=lambda entry: entry["level"])
    return levels


def _overrides_of(num: etree._Element) -> list[dict[str, Any]]:
    """The start overrides `num` carries, in level order."""
    overrides: list[dict[str, Any]] = []
    for override in num.iter(_W_LVL_OVERRIDE):
        start = _int_or_none(_attribute(override.find(_W_START_OVERRIDE), _W_VAL))
        if start is None:
            continue
        level = _int_or_none(override.get(_W_ILVL))
        overrides.append({"level": 0 if level is None else level, "start": start})
    overrides.sort(key=lambda entry: entry["level"])
    return overrides
