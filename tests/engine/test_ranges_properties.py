"""Random ranges over every fixture: what an edit must never disturb.

``tests/engine/test_ranges.py`` pins the behaviour of each function on hand-built
paragraphs.  This file asks a different question: run the same operations at
random offsets over the 19 fixtures -- tracked changes, comments, bookmarks,
fields, hyperlinks, drawings, footnotes, headers, content controls, tables --
and check, after every single one, that nothing outside the range moved.

The invariants, checked per operation:

* the visible text of the paragraph is exactly what the operation promised;
* the formatting of every character *outside* the range is untouched, compared
  character by character rather than run by run, since an edit is free to split
  and merge runs but not to restyle a character;
* the paragraph still holds the same number of markers -- a deletion moves a
  bookmark or a comment range, it never drops one;
* the text hidden under ``w:del`` and ``w:moveFrom`` is byte for byte the same,
  and so is ``w:pPr``;
* a refusal changes neither the text nor the formatting.

Then, once per fixture, at package scale:

* :func:`tests.support.package_check.validate_package` finds nothing, so no
  relationship, comment, note or numbering reference was orphaned;
* no insertion, deletion or move appeared or disappeared anywhere in the
  package (:func:`revision_tally`);
* :func:`tests.support.snapshot.assert_unchanged_except` finds no change outside
  the paragraphs the test edited -- neither in the other paragraphs, nor in the
  tables, the parts, the relationships, the body ``w:sectPr``, nor in the
  counters, which is what catches an image, a footnote reference or a hyperlink
  swept away by a deletion.
"""

from __future__ import annotations

import io
import random
import zipfile
from functools import cache
from pathlib import Path

import pytest
from lxml import etree

from tests.fixtures.builders import ALL_FIXTURES, build
from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.engine.errors import UnsupportedRange
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import (
    delete_range,
    insert_text,
    replace_range,
    resolve,
)
from word_document_server.engine.textmodel import segments, visible_text
from word_document_server.engine.xmlns import qn

W_P = qn("w:p")
W_RPR = qn("w:rPr")

#: Ranges drawn per paragraph.  Two is enough to cover a fixture without making
#: the suite crawl: the fixtures bring the variety, the seed brings the spread.
DRAWS = 2

#: What insertions write.  Plain text, a tab, a line break and non-ASCII, so the
#: ``w:tab`` / ``w:br`` translation and encoding are exercised everywhere.
PAYLOADS = ("X", "inserted ", "a\tb", "one\ntwo", "élan €")


@cache
def fixture_bytes(name: str) -> bytes:
    """Build a fixture once per session; the builders are deterministic."""
    return build(name)


@cache
def untouched_bytes(name: str) -> bytes:
    """The fixture as :class:`DocxPackage` rewrites it, with no edit at all.

    This, not the fixture itself, is the reference every comparison below is
    made against, so that what the tests measure is the work of ``ranges`` and
    nothing else.  Opening and saving is not byte-neutral today: the engine
    parses *every* XML part, including the ones python-docx leaves as opaque
    blobs (``theme1.xml``, ``webSettings.xml``, ``fontTable.xml``, ...), and
    python-docx's parser drops their indentation on the way through.  That is a
    property of ``engine/package.py``, not of a range operation, and it must not
    be able to hide behind an allowance here -- hence a baseline that already
    carries it.
    """
    return DocxPackage.open(fixture_bytes(name)).to_bytes()


def paragraphs_of(package: DocxPackage) -> list[tuple[str, int, etree._Element]]:
    """``(story, snapshot index, paragraph)`` for every paragraph of `package`.

    The index is the one :mod:`tests.support.snapshot` uses: every ``w:p`` of the
    story part in document order, table cells and content controls included.
    """
    return [
        (story, index, paragraph)
        for story, root in package.stories()
        for index, paragraph in enumerate(root.iter(W_P))
    ]


def char_formats(paragraph: etree._Element) -> list[str]:
    """The run properties in force for each visible character, one per character.

    Comparing this list rather than the runs themselves is what lets an edit
    split and merge runs freely while still proving that no character changed
    appearance.
    """
    formats: list[str] = []
    for segment in segments(paragraph):
        if not segment.text:
            continue
        run = segment.run
        properties = None if run is None else run.find(W_RPR)
        canonical = (
            ""
            if properties is None
            else etree.tostring(properties, method="c14n", exclusive=True).decode()
        )
        formats.extend([canonical] * len(segment.text))
    return formats


def marker_count(paragraph: etree._Element) -> int:
    """How many zero-width markers the paragraph holds."""
    return sum(1 for segment in segments(paragraph) if segment.kind == "marker")


def hidden_text(paragraph: etree._Element) -> list[str]:
    """The text the paragraph stores but does not show, piece by piece.

    Deleted and moved-away content is read through ``Segment.source_text``, so
    "``w:del`` is never modified" becomes an assertion instead of a hope.
    """
    return [
        segment.source_text for segment in segments(paragraph) if segment.kind == "hidden"
    ]


#: Every element ``tests.support.snapshot`` counts as a tracked revision.
REVISION_TAGS = frozenset(
    {
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
    }
)


def revision_tally(source: bytes | Path) -> dict[str, int]:
    """How many revision elements of each kind a package holds, part by part.

    The global ``revisions`` counter of a snapshot cannot be used as is: a split
    duplicates the ``w:rPr`` of the run it cuts, ``w:rPrChange`` included, so the
    total legitimately grows.  Counting per kind says the useful thing instead --
    no insertion, deletion or move may appear or disappear.
    """
    counts = dict.fromkeys(REVISION_TAGS, 0)
    handle = io.BytesIO(source) if isinstance(source, bytes) else source
    with zipfile.ZipFile(handle) as archive:
        for info in archive.infolist():
            if info.is_dir() or not info.filename.endswith(".xml"):
                continue
            root = etree.fromstring(archive.read(info))
            for element in root.iter():
                if not isinstance(element.tag, str):
                    continue
                local = etree.QName(element).localname
                if local in counts:
                    counts[local] += 1
    return counts


def canonical(element: etree._Element | None) -> str:
    if element is None:
        return ""
    return etree.tostring(element, method="c14n", exclusive=True).decode()


class Probe:
    """One paragraph, sampled before an operation so it can be judged after."""

    def __init__(self, paragraph: etree._Element) -> None:
        self.paragraph = paragraph
        self.xml = etree.tostring(paragraph)
        self.text = visible_text(paragraph)
        self.formats = char_formats(paragraph)
        self.markers = marker_count(paragraph)
        self.hidden = hidden_text(paragraph)
        self.ppr = canonical(paragraph.find(qn("w:pPr")))

    @property
    def mutated(self) -> bool:
        """Whether the paragraph's XML differs from what it was.

        A refused operation may leave the paragraph split -- splitting preserves
        the text and the formatting -- so "was it edited" is asked of the tree,
        not of the outcome.
        """
        return etree.tostring(self.paragraph) != self.xml

    def assert_kept(self, where: str) -> None:
        """The invariants every operation owes, whatever it did."""
        assert marker_count(self.paragraph) == self.markers, where
        assert hidden_text(self.paragraph) == self.hidden, where
        assert canonical(self.paragraph.find(qn("w:pPr"))) == self.ppr, where

    def assert_untouched(self, where: str) -> None:
        assert visible_text(self.paragraph) == self.text, where
        assert char_formats(self.paragraph) == self.formats, where
        self.assert_kept(where)

    def assert_became(self, where: str, text: str, kept: list[str]) -> None:
        assert visible_text(self.paragraph) == text, where
        assert char_formats(self.paragraph) == kept, where
        self.assert_kept(where)


def check_package(name: str, package: DocxPackage, touched, tmp_path: Path) -> None:
    """Save the edited package and compare it with the fixture it came from."""
    path = tmp_path / f"{name}.docx"
    package.save(path)
    assert validate_package(path) == [], name

    before = revision_tally(untouched_bytes(name))
    after = revision_tally(path)
    assert {kind: count for kind, count in after.items() if kind != "rPrChange"} == {
        kind: count for kind, count in before.items() if kind != "rPrChange"
    }, name
    assert after["rPrChange"] >= before["rPrChange"], name

    assert_unchanged_except(
        snapshot(untouched_bytes(name)),
        snapshot(path),
        paragraphs=sorted(touched),
        # Relaxed because a split copies the ``w:rPr`` it cuts, ``w:rPrChange``
        # included; the tally above is the precise version of the same check.
        counters={"revisions"},
    )


def open_fixture(name: str) -> tuple[DocxPackage, random.Random]:
    package = DocxPackage.open(fixture_bytes(name))
    return package, random.Random(f"J02-P3/{name}")


# --------------------------------------------------------------------------------------
# Deleting
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("name", sorted(ALL_FIXTURES))
def test_deleting_a_random_range_disturbs_nothing_else(name: str, tmp_path: Path) -> None:
    package, rng = open_fixture(name)
    touched: set[tuple[str, int]] = set()
    applied = 0

    for story, index, paragraph in paragraphs_of(package):
        for draw in range(DRAWS):
            probe = Probe(paragraph)
            if not probe.text:
                break
            start = rng.randrange(len(probe.text))
            end = rng.randrange(start + 1, len(probe.text) + 1)
            where = f"{name}/{story}[{index}] delete {start}..{end} #{draw}"
            try:
                removed = delete_range(paragraph, start, end)
            except UnsupportedRange:
                probe.assert_untouched(where)
            else:
                assert removed == probe.text[start:end], where
                probe.assert_became(
                    where,
                    probe.text[:start] + probe.text[end:],
                    probe.formats[:start] + probe.formats[end:],
                )
                applied += 1
            if probe.mutated:
                touched.add((story, index))

    assert applied, f"{name}: every deletion was refused, the test proves nothing"
    check_package(name, package, touched, tmp_path)


# --------------------------------------------------------------------------------------
# Inserting
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("name", sorted(ALL_FIXTURES))
def test_inserting_at_a_random_offset_disturbs_nothing_else(
    name: str, tmp_path: Path
) -> None:
    package, rng = open_fixture(name)
    touched: set[tuple[str, int]] = set()
    applied = 0

    for story, index, paragraph in paragraphs_of(package):
        for draw in range(DRAWS):
            probe = Probe(paragraph)
            offset = rng.randrange(len(probe.text) + 1)
            payload = rng.choice(PAYLOADS)
            side = rng.choice(("left", "right"))
            where = f"{name}/{story}[{index}] insert at {offset} {side} #{draw}"
            try:
                created = insert_text(paragraph, offset, payload, rpr_from=side)
            except UnsupportedRange:
                probe.assert_untouched(where)
            else:
                assert len(created) == 1, where
                text = probe.text[:offset] + payload + probe.text[offset:]
                assert visible_text(paragraph) == text, where
                formats = char_formats(paragraph)
                assert formats[:offset] == probe.formats[:offset], where
                assert formats[offset + len(payload) :] == probe.formats[offset:], where
                probe.assert_kept(where)
                applied += 1
            if probe.mutated:
                touched.add((story, index))

    assert applied, f"{name}: every insertion was refused, the test proves nothing"
    check_package(name, package, touched, tmp_path)


# --------------------------------------------------------------------------------------
# Replacing
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("name", sorted(ALL_FIXTURES))
def test_replacing_a_random_range_disturbs_nothing_else(
    name: str, tmp_path: Path
) -> None:
    package, rng = open_fixture(name)
    touched: set[tuple[str, int]] = set()
    applied = 0

    for story, index, paragraph in paragraphs_of(package):
        for draw in range(DRAWS):
            probe = Probe(paragraph)
            if not probe.text:
                break
            start = rng.randrange(len(probe.text))
            end = rng.randrange(start + 1, len(probe.text) + 1)
            payload = rng.choice(PAYLOADS)
            where = f"{name}/{story}[{index}] replace {start}..{end} #{draw}"
            try:
                replace_range(paragraph, start, end, payload)
            except UnsupportedRange:
                probe.assert_untouched(where)
            else:
                text = probe.text[:start] + payload + probe.text[end:]
                assert visible_text(paragraph) == text, where
                formats = char_formats(paragraph)
                assert formats[:start] == probe.formats[:start], where
                assert formats[start + len(payload) :] == probe.formats[end:], where
                probe.assert_kept(where)
                applied += 1
            if probe.mutated:
                touched.add((story, index))

    assert applied, f"{name}: every replacement was refused, the test proves nothing"
    check_package(name, package, touched, tmp_path)


# --------------------------------------------------------------------------------------
# The refusals themselves
# --------------------------------------------------------------------------------------


#: Upper bound: the only reasons a fixture may produce.  A fixture absent from
#: this map must accept *every* range it is asked about -- a layer that refuses
#: on a plain document would be unusable.
ALLOWED_REFUSALS: dict[str, set[str]] = {
    "comments": {"opaque-content"},
    "combined": {"contains-field", "crosses-field", "hidden-content", "opaque-content"},
    "drawings": {"opaque-content"},
    "fields": {"contains-field", "crosses-field"},
    "footnotes": {"opaque-content"},
    "headers_footers": {"contains-field", "crosses-field"},
    "tracked_changes": {"hidden-content"},
}

#: Lower bound: the reason each construct exists to trigger.  A refusal no
#: document can reach is dead code.
REQUIRED_REFUSALS: dict[str, set[str]] = {
    "comments": {"opaque-content"},
    "combined": {"crosses-field", "hidden-content", "opaque-content"},
    "drawings": {"opaque-content"},
    "fields": {"contains-field", "crosses-field"},
    "footnotes": {"opaque-content"},
    "headers_footers": {"contains-field"},
    "tracked_changes": {"hidden-content"},
}

#: How far a range is stretched from each start offset.  Every construct that
#: can refuse is local -- a field character, a deletion, an image -- so short
#: ranges from every offset find them all; the whole-paragraph range is added on
#: top so that containment is covered too.
SPAN = 8


@pytest.mark.parametrize("name", sorted(ALL_FIXTURES))
def test_the_refusals_a_fixture_produces_are_the_expected_ones(name: str) -> None:
    """Resolve range after range and map what each fixture refuses.

    ``resolve`` only splits, so the same paragraph can be asked about every one
    of its ranges in turn: the text, and therefore the offsets, never move.
    """
    package = DocxPackage.open(fixture_bytes(name))
    reasons: set[str] = set()
    for _, _, paragraph in paragraphs_of(package):
        text = visible_text(paragraph)
        for start in range(len(text)):
            ends = set(range(start + 1, min(start + SPAN, len(text)) + 1)) | {len(text)}
            for end in sorted(ends):
                try:
                    resolve(paragraph, start, end, operation="delete")
                except UnsupportedRange as refusal:
                    reasons.add(refusal.reason)
        assert visible_text(paragraph) == text, name

    assert reasons <= ALLOWED_REFUSALS.get(name, set()), f"{name}: got {sorted(reasons)}"
    assert reasons >= REQUIRED_REFUSALS.get(name, set()), f"{name}: got {sorted(reasons)}"
