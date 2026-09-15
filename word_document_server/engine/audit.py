"""The documentary audit: what is broken in this package, and what merely smells.

:func:`inspect <word_document_server.engine.inspect.inspect>` says what a
document *contains*; this module says what is *wrong* with it.  Both are read
only, and the difference is the judgement: an audit names a paragraph formatted
like a heading without being one, a style sheet with two entries a human cannot
tell apart, a ``w:numId`` pointing at a list the document does not define.

One shape for every answer
--------------------------
Every check produces the same :class:`Finding`: a `kind` (which check spoke), a
`severity`, a `locator`, a `message` for a human and a `data` mapping carrying
the numbers the message summarises.  A caller filters on `kind` and `severity`
and acts on `locator`; nothing has to be parsed out of a sentence.

Severity is about what the finding *is*, not about how urgent a human should
find it:

``error``
    the package contradicts itself -- a reference with no referent, a range with
    one end.  Word repairs or drops these silently, so they are lost work.
``warning``
    the document is valid and says something a reader will misread -- a heading
    that is not one, two styles named almost alike.
``info``
    an inventory a caller asked for by auditing: who revised what, which fields
    the document carries, where the direct formatting sits.

Locators, not positions
-----------------------
A finding about a paragraph carries the locator that reaches it, in the V2
vocabulary and in the one index space
:func:`~word_document_server.engine.find._v2_index_map` defines (D-006/D-016):
``{"story", "paragraph"}`` for an indexed paragraph, ``{"story", "table", "row",
"col", "paragraph"}`` for one in a table cell, and ``None`` for a paragraph in a
text box, which no locator reaches -- reported all the same, because a broken
bookmark in a text box is still broken.  A finding about the package as a whole
(a style, a numbering definition, an author) carries ``None`` too, and says what
it is about in `data`.

Reading only
------------
Nothing here writes, and nothing here normalises: an audit that repaired what it
found would make the next audit lie about the document a caller actually has.
The style sheet is read through :mod:`~word_document_server.engine.styles` and
the numbering through :mod:`~word_document_server.engine.numbering`, so what is
reported is what those two layers would report about the same document.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import Any

from lxml import etree

from word_document_server.engine.errors import PackageError
from word_document_server.engine.find import _v2_index_map, iter_paragraphs
from word_document_server.engine.locators import (
    _is_heading,
    _paragraph_of,
    _style_id,
    _style_names,
)
from word_document_server.engine.numbering import numbering_root, paragraph_list_info
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.revisions import list_revisions
from word_document_server.engine.styles import (
    _MERGED_PROPERTIES,
    StyleLevel,
    _cell_locators,
    _doc_defaults,
    _merge_into,
    decode_ppr,
    decode_rpr,
    get_style,
    list_styles,
)
from word_document_server.engine.textmodel import fields, visible_text
from word_document_server.engine.xmlns import qn

__all__ = [
    "DIRECT_FORMATTING_CROWD",
    "EMPTY_RUN_MIN",
    "FINDING_KINDS",
    "HEADING_LIKE_MAX_CHARS",
    "HEADING_LIKE_MIN_PT",
    "SEVERITIES",
    "TEXT_PREVIEW",
    "Finding",
    "audit",
]

#: The three severities, from the least to the most structural.  See the module
#: docstring for what each one claims.
SEVERITIES: tuple[str, ...] = ("info", "warning", "error")

#: Every `kind` :func:`audit` can produce, in the order the checks run.  A caller
#: switching on `kind` can be checked against this tuple; a kind absent from a
#: report simply means that check found nothing.
FINDING_KINDS: tuple[str, ...] = (
    "heading_like_paragraph",
    "direct_formatting_overrides_style",
    "mixed_fonts_in_paragraph",
    "consecutive_empty_paragraphs",
    "dangling_num_id",
    "dangling_abstract_num_id",
    "bookmark_without_end",
    "bookmark_without_start",
    "comment_range_without_end",
    "comment_range_without_start",
    "comment_without_anchor",
    "unused_custom_style",
    "similar_style_names",
    "revisions_by_author",
    "field_present",
)

#: Characters of text a finding quotes.  An audit is a list of places to look
#: at, not a copy of the document.
TEXT_PREVIEW = 80

#: A paragraph longer than this is prose, whatever it is formatted like, and is
#: never reported as heading-like.
HEADING_LIKE_MAX_CHARS = 80

#: Size, in points, from which direct formatting reads as "made bigger" on a
#: paragraph whose style states no size of its own.
HEADING_LIKE_MIN_PT = 14.0

#: Punctuation that ends a sentence rather than a title.  A line ending in one
#: is prose even when it is bold and short.
_SENTENCE_ENDINGS = ".;:"

#: Number of overridden properties from which direct formatting stops being a
#: touch-up and starts being a second style sheet.
DIRECT_FORMATTING_CROWD = 4

#: Shortest run of empty paragraphs worth reporting.  One empty paragraph is a
#: spacing habit; two in a row is layout done by hand.
EMPTY_RUN_MIN = 2

#: Field types whose cached result goes stale on its own -- they depend on
#: pagination or on another part of the document, so a package that does not ask
#: Word to update fields on open shows an old answer.
_STALE_PRONE_FIELDS = frozenset(
    {"TOC", "REF", "PAGEREF", "NUMPAGES", "SEQ", "INDEX", "STYLEREF", "TOA"}
)

#: What ``w:val`` spells when it means "off".  Word writes all three.
_OFF = frozenset({"0", "false", "off"})

#: Reported for a run whose font no layer states.
_UNSET_FONT = "(unset)"

#: Part name of the settings part, read for ``w:updateFields``.
_SETTINGS_PARTNAME = "/word/settings.xml"

#: Part name of the comments part.
_COMMENTS_PARTNAME = "/word/comments.xml"

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_T = qn("w:t")
_W_PPR = qn("w:pPr")
_W_RPR = qn("w:rPr")
_W_VAL = qn("w:val")
_W_ID = qn("w:id")
_W_NAME = qn("w:name")
_W_AUTHOR = qn("w:author")

_W_BOOKMARK_START = qn("w:bookmarkStart")
_W_BOOKMARK_END = qn("w:bookmarkEnd")
_W_COMMENT_RANGE_START = qn("w:commentRangeStart")
_W_COMMENT_RANGE_END = qn("w:commentRangeEnd")
_W_COMMENT_REFERENCE = qn("w:commentReference")
_W_COMMENT = qn("w:comment")

_W_NUM = qn("w:num")
_W_ABSTRACT_NUM = qn("w:abstractNum")
_W_ABSTRACT_NUM_ID = qn("w:abstractNumId")
_W_NUM_ID = qn("w:numId")

_W_PSTYLE = qn("w:pStyle")
_W_RSTYLE = qn("w:rStyle")
_W_TBL_STYLE = qn("w:tblStyle")
_W_NUM_STYLE_LINK = qn("w:numStyleLink")
_W_STYLE_LINK = qn("w:styleLink")

_W_UPDATE_FIELDS = qn("w:updateFields")


@dataclass(frozen=True)
class Finding:
    """One thing the audit has to say about a package.

    Attributes:
        kind: which check spoke -- one of :data:`FINDING_KINDS`.
        severity: one of :data:`SEVERITIES`.
        locator: the locator reaching the place the finding is about, ``None``
            for a finding about the package rather than about a paragraph (a
            style, a numbering definition, an author) and for a paragraph no
            locator reaches (a text box).
        message: one sentence, for a human.  It may change; `kind` and `data`
            are what a caller branches on.
        data: the detail the message summarises.  Plain JSON-ready data --
            strings, numbers, booleans, lists, dicts and ``None`` -- never an
            lxml element.
    """

    kind: str
    severity: str
    locator: dict[str, Any] | None
    message: str
    data: dict[str, Any] = field(default_factory=dict)


# --------------------------------------------------------------------------------------
# Reading helpers
# --------------------------------------------------------------------------------------


def _runs_of(paragraph: etree._Element) -> list[etree._Element]:
    """The runs that belong to `paragraph` itself, in document order.

    A ``w:r`` under a text box anchored in the paragraph is a run of *another*
    paragraph that happens to be nested in this one's subtree; counting it would
    let a caption's font decide whether the body paragraph has mixed fonts.
    """
    return [run for run in paragraph.iter(_W_R) if _paragraph_of(run) is paragraph]


def _run_text(run: etree._Element) -> str:
    """The visible text of one run.

    Only ``w:t`` counts: ``w:delText`` is text hidden under a tracked deletion
    and ``w:instrText`` is a field instruction, and neither is read by anyone.
    """
    return "".join(node.text or "" for node in run.findall(_W_T))


def _preview(text: str) -> str:
    """`text` cut to :data:`TEXT_PREVIEW`."""
    return text[:TEXT_PREVIEW]


def _flagged(element: etree._Element | None) -> bool:
    """Whether a bare toggle element is on -- absent means off."""
    return element is not None and element.get(_W_VAL) not in _OFF


def _update_fields(pkg: DocxPackage) -> bool:
    """Whether ``w:updateFields`` asks Word to refresh every field on open.

    The settings part is outside :data:`LIVE_CONTENT_TYPES
    <word_document_server.engine.package.LIVE_CONTENT_TYPES>` unless python-docx
    registers it, so the blob is parsed as a fallback (D-007): reading a
    detached copy is fine, nothing here writes.
    """
    part = pkg.find_part(_SETTINGS_PARTNAME)
    if part is None:
        return False
    try:
        root = pkg.root_of(part)
    except PackageError:  # pragma: no cover - python-docx registers settings.xml
        try:
            root = etree.fromstring(part.blob)
        except etree.XMLSyntaxError:
            return False
    return _flagged(root.find(_W_UPDATE_FIELDS))


def _comments_root(pkg: DocxPackage) -> etree._Element | None:
    """The live ``w:comments`` root, or ``None`` when the package has none."""
    part = pkg.find_part(_COMMENTS_PARTNAME)
    if part is None:
        return None
    try:
        return pkg.root_of(part)
    except PackageError:  # pragma: no cover - comments.xml is in LIVE_CONTENT_TYPES
        return None


# --------------------------------------------------------------------------------------
# Where a paragraph is
# --------------------------------------------------------------------------------------


class _Places:
    """The locators of one story's paragraphs, computed once.

    Both index spaces come from the modules that own them --
    :func:`~word_document_server.engine.find._v2_index_map` for the paragraph
    index and :func:`~word_document_server.engine.styles._cell_locators` for the
    table coordinates -- so an audit never invents a third numbering.
    """

    def __init__(self, story_id: str, root: etree._Element) -> None:
        self.story = story_id
        self.root = root
        self.paragraphs = iter_paragraphs(root)
        self._indices = _v2_index_map(self.paragraphs)
        self._cells = _cell_locators(story_id, root)

    def index(self, paragraph: etree._Element) -> int | None:
        """The V2 index of `paragraph`, or ``None`` outside that space."""
        return self._indices.get(paragraph)

    def indexed(self) -> list[tuple[int, etree._Element]]:
        """``(index, paragraph)`` for the V2-indexed paragraphs, in index order."""
        pairs = [(index, paragraph) for paragraph, index in self._indices.items()]
        pairs.sort(key=lambda pair: pair[0])
        return pairs

    def locator(self, paragraph: etree._Element) -> dict[str, Any] | None:
        """The locator reaching `paragraph`, or ``None`` when none does."""
        index = self._indices.get(paragraph)
        if index is not None:
            return {"story": self.story, "paragraph": index}
        return self._cells.get(paragraph)

    def locator_of(self, element: etree._Element) -> dict[str, Any] | None:
        """The locator of the paragraph `element` sits in, if it sits in one."""
        paragraph = _paragraph_of(element)
        return None if paragraph is None else self.locator(paragraph)


# --------------------------------------------------------------------------------------
# What a style says
# --------------------------------------------------------------------------------------


class _StyleSheet:
    """The style sheet, resolved once per style and cached.

    :func:`~word_document_server.engine.styles.get_style` resolves a chain
    *including* ``w:docDefaults``, which is the right answer for one style and
    the wrong one for a stack: layering a character style's resolution on top of
    a paragraph style's would let the defaults of the second beat the first.
    This class keeps the two apart -- :meth:`own` is a style's chain without the
    defaults, :attr:`defaults_run` and :attr:`defaults_paragraph` are the
    defaults on their own -- so they can be merged in Word's order.
    """

    def __init__(self, pkg: DocxPackage) -> None:
        self._pkg = pkg
        self.infos = list_styles(pkg)
        self.by_id = {info.style_id: info for info in self.infos}
        defaults = _doc_defaults(pkg)
        self.defaults_run = dict(defaults.run_props)
        self.defaults_paragraph = dict(defaults.paragraph_props)
        self.default_paragraph_style = next(
            (info.style_id for info in self.infos if info.family == "paragraph" and info.default),
            None,
        )
        self._own: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {}

    def own(self, style_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
        """What `style_id` and its ancestors say, document defaults excluded.

        A style the document does not define contributes nothing, which is what
        Word does with a dangling ``w:pStyle`` too.
        """
        cached = self._own.get(style_id)
        if cached is not None:
            return cached
        try:
            detail = get_style(self._pkg, style_id)
        except (PackageError, ValueError):
            resolved: tuple[dict[str, Any], dict[str, Any]] = ({}, {})
            self._own[style_id] = resolved
            return resolved
        levels: list[StyleLevel] = [level for level in detail.chain if level.style_id]
        run: dict[str, Any] = {}
        paragraph: dict[str, Any] = {}
        for level in reversed(levels):
            _merge_into(run, level.run_props)
            _merge_into(paragraph, level.paragraph_props)
        resolved = (run, paragraph)
        self._own[style_id] = resolved
        return resolved

    def paragraph_props(
        self, paragraph: etree._Element
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        """What the *style* of `paragraph` contributes: ``(run, paragraph)``.

        The paragraph's own ``w:pPr`` and the ``w:rPr`` of its runs are not in
        it -- that is the whole point: what is compared against this is the
        direct formatting.
        """
        style_id = _style_id(paragraph) or self.default_paragraph_style
        run = dict(self.defaults_run)
        properties = dict(self.defaults_paragraph)
        if style_id:
            own_run, own_paragraph = self.own(style_id)
            _merge_into(run, own_run)
            _merge_into(properties, own_paragraph)
        return run, properties

    def run_props(self, base: dict[str, Any], char_style: str | None) -> dict[str, Any]:
        """`base` with the character style of a run layered on top."""
        if not char_style:
            return base
        merged = dict(base)
        _merge_into(merged, self.own(char_style)[0])
        return merged


def _overridden(direct: dict[str, Any], inherited: dict[str, Any]) -> list[str]:
    """Properties `direct` sets that `inherited` also sets, to something else.

    A property the style says nothing about is *not* an override: it is the only
    thing saying what the text looks like.  The comparison follows the
    inheritance: ``indent``, ``spacing`` and ``numbering`` hold several
    independent settings and are compared one by one, so a paragraph that only
    moves the first line does not read as overriding the left indent it keeps;
    ``font`` and ``color`` are one property written over several attributes, so
    a level stating one states the whole thing.
    """
    names: list[str] = []
    for name, value in sorted(direct.items()):
        if name not in inherited:
            continue
        other = inherited[name]
        if name in _MERGED_PROPERTIES and isinstance(value, dict) and isinstance(other, dict):
            names.extend(
                f"{name}.{key}"
                for key, part in sorted(value.items())
                if key in other and other[key] != part
            )
        elif value != other:
            names.append(name)
    return names


def _font_label(properties: dict[str, Any]) -> str:
    """How a font reads in a report: the literal, or the theme it defers to."""
    font = properties.get("font")
    if font is None:
        return _UNSET_FONT
    if not isinstance(font, dict):  # pragma: no cover - decode_rpr always reports the pair
        return str(font)
    value = font.get("value")
    if value:
        return str(value)
    theme = font.get("theme")
    return f"theme:{theme}" if theme else _UNSET_FONT


# --------------------------------------------------------------------------------------
# The context every check is handed
# --------------------------------------------------------------------------------------


class _Audit:
    """Everything the checks share, read once."""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self.places = [_Places(story_id, root) for story_id, root in pkg.stories()]
        self.styles = _StyleSheet(pkg)
        self.style_names = _style_names(pkg)


# --------------------------------------------------------------------------------------
# Checks: formatting
# --------------------------------------------------------------------------------------


def _heading_like_paragraphs(state: _Audit) -> list[Finding]:
    """Short emphasised paragraphs that no heading style makes a heading.

    They read as titles, and nothing that reads the outline -- the navigation
    pane, a table of contents, this engine's ``heading`` locator -- sees them.
    """
    found: list[Finding] = []
    for place in state.places:
        for paragraph in place.paragraphs:
            text = visible_text(paragraph).strip()
            if not text or len(text) > HEADING_LIKE_MAX_CHARS:
                continue
            if text[-1] in _SENTENCE_ENDINGS:
                continue
            if _is_heading(paragraph, state.style_names):
                continue
            if paragraph_list_info(paragraph) is not None:
                continue
            cues = _emphasis_cues(paragraph, state.styles)
            if cues is None:
                continue
            found.append(
                Finding(
                    kind="heading_like_paragraph",
                    severity="warning",
                    locator=place.locator(paragraph),
                    message=(
                        f"{text[:40]!r} is formatted like a heading "
                        f"({', '.join(cues)}) but carries no heading style"
                    ),
                    data={
                        "story": place.story,
                        "text": _preview(text),
                        "cues": cues,
                        "style": _style_id(paragraph),
                    },
                )
            )
    return found


def _emphasis_cues(paragraph: etree._Element, styles: _StyleSheet) -> list[str] | None:
    """The direct emphasis every text-carrying run of `paragraph` shares.

    ``None`` when at least one such run carries none, which is what tells a
    title apart from a sentence with a bold word in it -- and when the paragraph
    carries no text-carrying run at all.
    """
    base_run, _ = styles.paragraph_props(paragraph)
    cues: set[str] = set()
    seen = False
    for run in _runs_of(paragraph):
        if not _run_text(run):
            continue
        seen = True
        direct = decode_rpr(run.find(_W_RPR))
        inherited = styles.run_props(base_run, direct.get("char_style"))
        own: set[str] = set()
        if direct.get("bold") is True:
            own.add("bold")
        if direct.get("caps") is True or direct.get("small_caps") is True:
            own.add("caps")
        size = direct.get("size_pt")
        if size is not None and _is_larger(size, inherited.get("size_pt")):
            own.add("larger")
        if not own:
            return None
        cues |= own
    return sorted(cues) if seen else None


def _is_larger(size: float, inherited: float | None) -> bool:
    """Whether a directly set size reads as "made bigger"."""
    if inherited is None:
        return size >= HEADING_LIKE_MIN_PT
    return size > inherited


def _direct_formatting_over_style(state: _Audit) -> list[Finding]:
    """Paragraphs whose direct formatting contradicts the style they name.

    Reported per paragraph with the properties named, because that is the
    decision a caller faces: either the style is wrong for this paragraph and
    another one should be applied, or the style is right and the direct
    formatting is what makes the document drift.
    """
    found: list[Finding] = []
    for place in state.places:
        for paragraph in place.paragraphs:
            base_run, base_paragraph = state.styles.paragraph_props(paragraph)
            paragraph_names = _overridden(
                decode_ppr(paragraph.find(_W_PPR)), base_paragraph
            )
            run_names: set[str] = set()
            runs = 0
            for run in _runs_of(paragraph):
                direct = decode_rpr(run.find(_W_RPR))
                char_style = direct.pop("char_style", None)
                names = _overridden(direct, state.styles.run_props(base_run, char_style))
                if names:
                    runs += 1
                    run_names |= set(names)
            properties = [*paragraph_names, *sorted(run_names)]
            if not properties:
                continue
            crowded = len(properties) >= DIRECT_FORMATTING_CROWD
            found.append(
                Finding(
                    kind="direct_formatting_overrides_style",
                    severity="warning" if crowded else "info",
                    locator=place.locator(paragraph),
                    message=(
                        "direct formatting overrides "
                        f"{len(properties)} propert{'ies' if len(properties) > 1 else 'y'} "
                        f"of style "
                        f"{_style_id(paragraph) or state.styles.default_paragraph_style!r}: "
                        f"{', '.join(properties)}"
                    ),
                    data={
                        "story": place.story,
                        "style": _style_id(paragraph),
                        "properties": properties,
                        "count": len(properties),
                        "paragraph_properties": paragraph_names,
                        "run_properties": sorted(run_names),
                        "runs": runs,
                    },
                )
            )
    return found


def _mixed_fonts(state: _Audit) -> list[Finding]:
    """Paragraphs whose runs do not all end up in the same font.

    The font each run actually gets is what is compared -- its own ``w:rFonts``,
    else its character style's, else the paragraph style's -- so a paragraph
    where one word was pasted in from elsewhere is reported even though nothing
    in it looks unusual on its own.
    """
    found: list[Finding] = []
    for place in state.places:
        for paragraph in place.paragraphs:
            base_run, _ = state.styles.paragraph_props(paragraph)
            labels: list[str] = []
            for run in _runs_of(paragraph):
                if not _run_text(run):
                    continue
                direct = decode_rpr(run.find(_W_RPR))
                inherited = state.styles.run_props(base_run, direct.get("char_style"))
                merged = dict(inherited)
                _merge_into(merged, direct)
                label = _font_label(merged)
                if label not in labels:
                    labels.append(label)
            if len(labels) < 2:
                continue
            found.append(
                Finding(
                    kind="mixed_fonts_in_paragraph",
                    severity="info",
                    locator=place.locator(paragraph),
                    message=f"this paragraph mixes {len(labels)} fonts: {', '.join(labels)}",
                    data={
                        "story": place.story,
                        "fonts": labels,
                        "text": _preview(visible_text(paragraph)),
                    },
                )
            )
    return found


def _empty_paragraph_runs(state: _Audit) -> list[Finding]:
    """Runs of consecutive empty paragraphs -- vertical spacing typed by hand.

    Only the V2 index space is walked: an empty paragraph closing a table cell
    is required by the schema, not a layout habit.
    """
    found: list[Finding] = []
    for place in state.places:
        streak: list[int] = []
        first: etree._Element | None = None
        for index, paragraph in place.indexed():
            if visible_text(paragraph) == "":
                if not streak:
                    first = paragraph
                streak.append(index)
                continue
            found.extend(_empty_run_finding(place, streak, first))
            streak, first = [], None
        found.extend(_empty_run_finding(place, streak, first))
    return found


def _empty_run_finding(
    place: _Places, streak: list[int], first: etree._Element | None
) -> list[Finding]:
    """The finding a finished streak of empty paragraphs deserves, if any."""
    if len(streak) < EMPTY_RUN_MIN or first is None:
        return []
    return [
        Finding(
            kind="consecutive_empty_paragraphs",
            severity="info",
            locator=place.locator(first),
            message=(
                f"{len(streak)} empty paragraphs in a row, from paragraph {streak[0]} "
                f"to {streak[-1]}"
            ),
            data={
                "story": place.story,
                "start": streak[0],
                "end": streak[-1],
                "count": len(streak),
            },
        )
    ]


# --------------------------------------------------------------------------------------
# Checks: references
# --------------------------------------------------------------------------------------


def _defined_numbering(pkg: DocxPackage) -> tuple[set[int], set[int], list[tuple[int, int]]]:
    """``(num ids, abstract ids, (num id, abstract id) pairs)`` of the package."""
    root = numbering_root(pkg)
    if root is None:
        return set(), set(), []
    abstracts = {
        value
        for element in root.findall(_W_ABSTRACT_NUM)
        if (value := _as_int(element.get(_W_ABSTRACT_NUM_ID))) is not None
    }
    nums: set[int] = set()
    pairs: list[tuple[int, int]] = []
    for element in root.findall(_W_NUM):
        num_id = _as_int(element.get(_W_NUM_ID))
        if num_id is None:
            continue
        nums.add(num_id)
        child = element.find(_W_ABSTRACT_NUM_ID)
        abstract_id = None if child is None else _as_int(child.get(_W_VAL))
        if abstract_id is not None:
            pairs.append((num_id, abstract_id))
    return nums, abstracts, pairs


def _as_int(value: str | None) -> int | None:
    """An attribute read as an int, or ``None`` when absent or not a number."""
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _dangling_numbering(state: _Audit) -> list[Finding]:
    """``w:numId`` and ``w:abstractNumId`` references with no definition.

    Word drops the numbering of a paragraph whose list it cannot find, and it
    does it silently: the paragraph keeps its indent and loses its number.
    """
    nums, abstracts, pairs = _defined_numbering(state.pkg)
    found: list[Finding] = []
    for place in state.places:
        for paragraph in place.paragraphs:
            info = paragraph_list_info(paragraph)
            if info is None:
                continue
            num_id = info["num_id"]
            # ``w:numId`` 0 is how Word *removes* numbering from a paragraph, so
            # it is a value, not a reference.
            if num_id is None or num_id == 0 or num_id in nums:
                continue
            found.append(
                Finding(
                    kind="dangling_num_id",
                    severity="error",
                    locator=place.locator(paragraph),
                    message=(
                        f"this paragraph is numbered with numId {num_id}, which the "
                        "numbering part does not define"
                    ),
                    data={
                        "story": place.story,
                        "num_id": num_id,
                        "level": info["level"],
                        "text": _preview(visible_text(paragraph)),
                    },
                )
            )
    for num_id, abstract_id in pairs:
        if abstract_id in abstracts:
            continue
        found.append(
            Finding(
                kind="dangling_abstract_num_id",
                severity="error",
                locator=None,
                message=(
                    f"list numId {num_id} points at abstractNumId {abstract_id}, "
                    "which the numbering part does not define"
                ),
                data={"num_id": num_id, "abstract_num_id": abstract_id},
            )
        )
    return found


def _paired_markers(
    place: _Places, start_tag: str, end_tag: str
) -> tuple[list[etree._Element], set[str], set[str]]:
    """``(start elements, start ids, end ids)`` of a marker pair in one story."""
    starts = [
        element for element in place.root.iter(start_tag) if element.get(_W_ID) is not None
    ]
    start_ids = {element.get(_W_ID) or "" for element in starts}
    end_ids = {
        element.get(_W_ID) or ""
        for element in place.root.iter(end_tag)
        if element.get(_W_ID) is not None
    }
    return starts, start_ids, end_ids


def _bookmark_ranges(state: _Audit) -> list[Finding]:
    """Bookmarks with one end only.

    A ``w:bookmarkStart`` whose ``w:bookmarkEnd`` is missing has no span: a
    cross-reference to it resolves to nothing, and Word discards it on the next
    save.
    """
    found: list[Finding] = []
    for place in state.places:
        starts, start_ids, end_ids = _paired_markers(
            place, _W_BOOKMARK_START, _W_BOOKMARK_END
        )
        for element in starts:
            marker = element.get(_W_ID) or ""
            if marker in end_ids:
                continue
            name = element.get(_W_NAME)
            found.append(
                Finding(
                    kind="bookmark_without_end",
                    severity="error",
                    locator=place.locator_of(element),
                    message=f"bookmark {name!r} opens and is never closed",
                    data={"story": place.story, "name": name, "id": marker},
                )
            )
        for marker in sorted(end_ids - start_ids):
            found.append(
                Finding(
                    kind="bookmark_without_start",
                    severity="error",
                    locator=None,
                    message=f"a bookmarkEnd with id {marker!r} closes a bookmark that never opens",
                    data={"story": place.story, "id": marker},
                )
            )
    return found


def _comment_ranges(state: _Audit) -> list[Finding]:
    """Comment ranges with one end only.

    The anchor of a comment is the pair; with one half missing Word falls back
    to anchoring on the reference mark alone, so the comment silently stops
    pointing at the text it was written about.
    """
    found: list[Finding] = []
    for place in state.places:
        starts, start_ids, end_ids = _paired_markers(
            place, _W_COMMENT_RANGE_START, _W_COMMENT_RANGE_END
        )
        for element in starts:
            marker = element.get(_W_ID) or ""
            if marker in end_ids:
                continue
            found.append(
                Finding(
                    kind="comment_range_without_end",
                    severity="error",
                    locator=place.locator_of(element),
                    message=f"the range of comment {marker!r} opens and is never closed",
                    data={"story": place.story, "comment_id": marker},
                )
            )
        for marker in sorted(end_ids - start_ids):
            found.append(
                Finding(
                    kind="comment_range_without_start",
                    severity="error",
                    locator=None,
                    message=f"the range of comment {marker!r} closes without opening",
                    data={"story": place.story, "comment_id": marker},
                )
            )
    return found


def _comments_without_anchor(state: _Audit) -> list[Finding]:
    """Comments no ``w:commentReference`` points at.

    They exist in ``comments.xml`` and appear nowhere: Word shows them in no
    margin, and the next round trip through a converter drops them.
    """
    root = _comments_root(state.pkg)
    if root is None:
        return []
    anchored = {
        element.get(_W_ID)
        for place in state.places
        for element in place.root.iter(_W_COMMENT_REFERENCE)
        if element.get(_W_ID) is not None
    }
    found: list[Finding] = []
    for comment in root.findall(_W_COMMENT):
        marker = comment.get(_W_ID)
        if marker is None or marker in anchored:
            continue
        found.append(
            Finding(
                kind="comment_without_anchor",
                severity="warning",
                locator=None,
                message=(
                    f"comment {marker!r} by {comment.get(_W_AUTHOR) or 'an unnamed author'} "
                    "is anchored nowhere in the document"
                ),
                data={
                    "comment_id": marker,
                    "author": comment.get(_W_AUTHOR),
                    "text": _preview(
                        " ".join(visible_text(p) for p in comment.findall(_W_P)).strip()
                    ),
                },
            )
        )
    return found


# --------------------------------------------------------------------------------------
# Checks: the style sheet
# --------------------------------------------------------------------------------------


def _referenced_styles(state: _Audit) -> set[str]:
    """Every style id something in the package points at.

    Three kinds of pointer count: a story that applies the style
    (``w:pStyle``/``w:rStyle``/``w:tblStyle``), another style that inherits from
    it or links to it, and a numbering level that imposes it.
    """
    referenced: set[str] = set()
    for place in state.places:
        for element in place.root.iter(_W_PSTYLE, _W_RSTYLE, _W_TBL_STYLE):
            value = element.get(_W_VAL)
            if value is not None:
                referenced.add(value)
    for info in state.styles.infos:
        referenced.update(
            value for value in (info.based_on, info.next_style, info.link) if value
        )
    root = numbering_root(state.pkg)
    if root is not None:
        for element in root.iter(_W_PSTYLE, _W_NUM_STYLE_LINK, _W_STYLE_LINK):
            value = element.get(_W_VAL)
            if value is not None:
                referenced.add(value)
    return referenced


def _unused_custom_styles(state: _Audit) -> list[Finding]:
    """Styles the document declares as its own and never uses.

    Only a style flagged ``w:customStyle="1"`` is reported: Word's own style
    sheet ships dozens of built-in definitions a document does not use, and
    listing those would bury the handful somebody added on purpose.
    """
    referenced = _referenced_styles(state)
    found: list[Finding] = []
    for info in state.styles.infos:
        if info.builtin or info.default or info.style_id in referenced:
            continue
        found.append(
            Finding(
                kind="unused_custom_style",
                severity="info",
                locator=None,
                message=(
                    f"custom style {info.name!r} (id {info.style_id!r}) is defined "
                    "and never used"
                ),
                data={
                    "style_id": info.style_id,
                    "name": info.name,
                    "family": info.family,
                },
            )
        )
    return found


def _normalized_name(name: str) -> str:
    """A style name reduced to what a reader distinguishes it by.

    Case, spaces and punctuation go: "Fixture Body", "fixture body" and
    "Fixture-Body" are one name in a style gallery, and a document carrying all
    three offers a human three entries they cannot tell apart.
    """
    return "".join(character for character in name.casefold() if character.isalnum())


def _similar_style_names(state: _Audit) -> list[Finding]:
    """Groups of styles whose names differ only by case, spacing or punctuation."""
    groups: dict[str, list[Any]] = {}
    for info in state.styles.infos:
        key = _normalized_name(info.name)
        if not key:
            continue
        groups.setdefault(key, []).append(info)
    found: list[Finding] = []
    for key, infos in sorted(groups.items()):
        if len(infos) < 2:
            continue
        names = [info.name for info in infos]
        found.append(
            Finding(
                kind="similar_style_names",
                severity="warning",
                locator=None,
                message=(
                    f"{len(infos)} styles are named almost alike and read as one in a "
                    f"style gallery: {', '.join(repr(name) for name in names)}"
                ),
                data={
                    "normalized": key,
                    "styles": [
                        {
                            "style_id": info.style_id,
                            "name": info.name,
                            "family": info.family,
                        }
                        for info in infos
                    ],
                },
            )
        )
    return found


# --------------------------------------------------------------------------------------
# Checks: inventories
# --------------------------------------------------------------------------------------


def _revisions_by_author(state: _Audit) -> list[Finding]:
    """Who left tracked changes in this document, and how many of each kind."""
    revisions = list_revisions(state.pkg)
    by_author: dict[str, Counter[str]] = {}
    for revision in revisions:
        by_author.setdefault(revision.author, Counter())[revision.kind] += 1
    found: list[Finding] = []
    for author, kinds in sorted(by_author.items()):
        total = sum(kinds.values())
        found.append(
            Finding(
                kind="revisions_by_author",
                severity="info",
                locator=None,
                message=(
                    f"{author!r} left {total} tracked "
                    f"{'revisions' if total > 1 else 'revision'}"
                ),
                data={
                    "author": author,
                    "count": total,
                    "kinds": dict(sorted(kinds.items())),
                },
            )
        )
    return found


def _field_name(instr: str) -> str:
    """The field type an instruction opens with, upper-cased.

    ``" TOC \\o \\"1-3\\" "`` -> ``"TOC"``.  An instruction with no first word at
    all -- a field whose instruction is a nested field and nothing else -- reads
    as ``"(unnamed)"`` rather than being dropped, since it is a field either way.
    """
    stripped = instr.strip()
    if not stripped:
        return "(unnamed)"
    return stripped.split(maxsplit=1)[0].upper()


def _field_inventory(state: _Audit) -> list[Finding]:
    """The fields the document carries, grouped by type.

    A field shows its *cached* result until something recomputes it, so the
    report says whether the package asks Word to refresh them on open
    (``w:updateFields`` in the settings part).  Nested fields are counted on
    their own, as :func:`~word_document_server.engine.textmodel.fields` reports
    them (D-005).
    """
    updates = _update_fields(state.pkg)
    counts: Counter[str] = Counter()
    locations: dict[str, list[dict[str, Any] | None]] = {}
    instructions: dict[str, list[str]] = {}
    for place in state.places:
        for paragraph in place.paragraphs:
            for one in fields(paragraph):
                name = _field_name(one.instr)
                counts[name] += 1
                locations.setdefault(name, []).append(place.locator(paragraph))
                known = instructions.setdefault(name, [])
                instruction = one.instr.strip()
                if instruction not in known:
                    known.append(instruction)
    found: list[Finding] = []
    for name, count in sorted(counts.items()):
        stale = name in _STALE_PRONE_FIELDS and not updates
        found.append(
            Finding(
                kind="field_present",
                severity="warning" if stale else "info",
                locator=None,
                message=(
                    f"{count} {name} field{'s' if count > 1 else ''}; "
                    + (
                        "the settings part does not ask Word to update fields on open, "
                        "so the cached result is what a reader sees"
                        if stale
                        else f"updateFields is {'on' if updates else 'off'}"
                    )
                ),
                data={
                    "field": name,
                    "count": count,
                    "update_fields": updates,
                    "instructions": instructions[name],
                    "locations": locations[name],
                },
            )
        )
    return found


#: The checks, in the order :func:`audit` runs them -- which is the order their
#: findings come back in.
_CHECKS = (
    _heading_like_paragraphs,
    _direct_formatting_over_style,
    _mixed_fonts,
    _empty_paragraph_runs,
    _dangling_numbering,
    _bookmark_ranges,
    _comment_ranges,
    _comments_without_anchor,
    _unused_custom_styles,
    _similar_style_names,
    _revisions_by_author,
    _field_inventory,
)


def audit(pkg: DocxPackage) -> list[Finding]:
    """Audit `pkg` and return everything the checks have to say about it.

    Twelve checks run, in the order of :data:`FINDING_KINDS`, and each walks the
    document in story then document order, so two audits of the same bytes
    return the same list in the same order.  An empty list means every check
    passed, not that the document was not read.

    What is checked, by family:

    *Formatting that says the wrong thing*
        ``heading_like_paragraph`` -- a short emphasised paragraph no heading
        style makes a heading; ``direct_formatting_overrides_style`` -- the
        properties a paragraph sets directly against what its style already
        says, counted; ``mixed_fonts_in_paragraph`` -- runs that do not end up
        in the same font; ``consecutive_empty_paragraphs`` -- vertical spacing
        typed by hand.

    *References with no referent*
        ``dangling_num_id`` and ``dangling_abstract_num_id`` -- numbering that
        points at a definition the package does not carry;
        ``bookmark_without_end``/``bookmark_without_start`` and
        ``comment_range_without_end``/``comment_range_without_start`` -- ranges
        with one end; ``comment_without_anchor`` -- a comment nothing points at.

    *A style sheet a human cannot read*
        ``unused_custom_style`` -- declared, never applied;
        ``similar_style_names`` -- names that differ only by case, spacing or
        punctuation.

    *Inventories*
        ``revisions_by_author`` -- who changed what, by kind;
        ``field_present`` -- which fields the document carries, with what
        ``w:updateFields`` says about their cached results.

    Args:
        pkg: the package to read.  It is not modified, and nothing in the
            returned findings refers to it: every value is plain data.

    Returns:
        One :class:`Finding` per thing found.
    """
    state = _Audit(pkg)
    found: list[Finding] = []
    for check in _CHECKS:
        found.extend(check(state))
    return found
