"""Tests for :mod:`word_document_server.engine.revisions`.

Two scales, on purpose.  The unit cases below are written against hand-built
paragraphs whose XML is visible in the test and hosted in a real package (id
allocation reads the whole package, so a detached ``w:p`` would not do): a
revision is judged on the tree it leaves behind, because where a ``w:del`` lands
relative to a ``w:ins`` is the whole meaning of the operation.

The package-scale cases then run on the fixtures and compare snapshots.  Their
reference is never the fixture file itself but the fixture after one
``open``/``to_bytes`` round trip with no edit at all -- see
:func:`untouched_bytes` and ``tests/engine/test_ranges_properties.py``: opening
and saving is not byte-neutral yet, and that must not be able to hide behind an
allowance here.

The two defects this module replaces are pinned as regressions:
``test_tracked_replace_terminates_when_the_new_text_contains_the_old`` (the
infinite loop of ``core/tracked_changes.py``) and the several
``test_accept_refuses_...`` cases (its silent half-application).
"""

from __future__ import annotations

from copy import deepcopy
from datetime import UTC, datetime, timedelta, timezone
from functools import cache

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, diff, snapshot
from word_document_server.engine.errors import (
    LocatorError,
    UnsupportedRange,
    UnsupportedRevision,
)
from word_document_server.engine.find import find, iter_paragraphs
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import replace_range, resolve
from word_document_server.engine.revisions import (
    REVISION_KINDS,
    accept,
    list_revisions,
    reject,
    tracked_delete,
    tracked_insert,
    tracked_replace,
)
from word_document_server.engine.textmodel import visible_text
from word_document_server.engine.xmlns import NAMESPACES, qn

W_ID = qn("w:id")
W_AUTHOR = qn("w:author")
W_DATE = qn("w:date")
W_P = qn("w:p")
W_R = qn("w:r")
W_RPR = qn("w:rPr")
W_BODY = qn("w:body")
W_SECT_PR = qn("w:sectPr")
W_INS = qn("w:ins")
W_DEL = qn("w:del")

#: Every test that writes a revision uses these, so that nothing depends on the
#: clock and every stamp in an assertion is the one that was asked for.
AUTHOR = "Bob Reviewer"
OTHER = "Alice Author"
STAMP = "2026-01-02T03:04:05Z"


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------


@cache
def untouched_bytes(name: str) -> bytes:
    """The fixture as :class:`DocxPackage` rewrites it, with no edit at all.

    The reference for every snapshot comparison here; see the module docstring.
    """
    return DocxPackage.open(build(name)).to_bytes()


def paragraph_from(children: str) -> etree._Element:
    """Return a standalone ``w:p`` whose children are the given XML fragment."""
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in NAMESPACES.items())
    return etree.fromstring(f"<w:p {declarations}>{children}</w:p>")


def hosted(children: str, name: str = "simple") -> tuple[DocxPackage, etree._Element]:
    """A real package whose body ends with a paragraph built from `children`.

    The paragraph goes before the body's ``w:sectPr``, which has to stay last.
    """
    pkg = DocxPackage.open(build(name))
    body = dict(pkg.stories())["document"].find(W_BODY)
    paragraph = paragraph_from(children)
    section = body.find(W_SECT_PR)
    if section is None:
        body.append(paragraph)
    else:
        body.insert(body.index(section), paragraph)
    return pkg, paragraph


def paragraphs_of(pkg: DocxPackage, story: str = "document") -> list[etree._Element]:
    """Every paragraph of one story, in document order."""
    return iter_paragraphs(dict(pkg.stories())[story])


def names(element: etree._Element) -> list[str]:
    """Local names of the element's children, in order."""
    return [etree.QName(child).localname for child in element]


def texts(element: etree._Element) -> list[str]:
    """Local name and text of every text-carrying node under `element`."""
    return [
        f"{etree.QName(node).localname}:{node.text}"
        for node in element.iter(qn("w:t"), qn("w:delText"))
    ]


def properties(run: etree._Element) -> str:
    """Canonical form of a run's ``w:rPr``, "" when it has none."""
    found = run.find(W_RPR)
    return "" if found is None else etree.tostring(found, method="c14n", exclusive=True).decode()


def ids_of(elements) -> list[int]:
    """The ``w:id`` of each element, as ints."""
    return [int(element.get(W_ID)) for element in elements]


def all_annotation_ids(pkg: DocxPackage) -> list[str]:
    """Every ``w:id`` carried by a revision element, across every story."""
    return [
        element.get(W_ID)
        for _, root in pkg.stories()
        for element in root.iter(W_INS, W_DEL, qn("w:moveFrom"), qn("w:moveTo"))
        if element.get(W_ID) is not None
    ]


def split_by_an_inserted_mark(name: str = "mixed_runs") -> tuple[bytes, bytes, bytes]:
    """A fixture paragraph cut in two by an inserted paragraph mark.

    Returns three packages: the fixture untouched, the split recorded as a
    revision, and the very same split with no revision at all.  The second is
    what rejecting the mark must turn back into the first, and what accepting it
    must turn into the third -- so both directions are pinned against a document
    built without ``accept``/``reject`` having any part in it.

    The shape is Word's: cutting a paragraph puts the new mark on the *first*
    half, whose ``w:pPr`` is a copy of the mark being split, and leaves the
    original mark closing the second half.
    """
    source = untouched_bytes(name)
    pkg = DocxPackage.open(source)
    body = dict(pkg.stories())["document"].find(W_BODY)
    tail = next(
        candidate
        for candidate in body.iterchildren(W_P)
        if len(candidate.findall(W_R)) >= 2
        and candidate.find(f"{qn('w:pPr')}/{W_RPR}") is not None
    )
    head = etree.Element(W_P)
    body.insert(body.index(tail), head)
    head.append(deepcopy(tail.find(qn("w:pPr"))))
    mark = etree.Element(W_INS)
    mark.set(W_ID, "9001")
    mark.set(W_AUTHOR, AUTHOR)
    mark.set(W_DATE, STAMP)
    properties_of_mark = head.find(f"{qn('w:pPr')}/{W_RPR}")
    properties_of_mark.insert(0, mark)
    head.append(tail.find(W_R))
    tracked = pkg.to_bytes()
    properties_of_mark.remove(mark)
    return source, tracked, pkg.to_bytes()


# --------------------------------------------------------------------------------------
# tracked_delete
# --------------------------------------------------------------------------------------


def test_tracked_delete_wraps_the_runs_and_turns_their_text_into_deleted_text() -> None:
    pkg, paragraph = hosted(
        '<w:r><w:t xml:space="preserve">Hello </w:t></w:r>'
        "<w:r><w:rPr><w:b/></w:rPr><w:t>world</w:t></w:r>"
        "<w:r><w:t>!</w:t></w:r>"
    )
    pieces = resolve(paragraph, 0, 11, operation="delete")

    created = tracked_delete(pkg, pieces, AUTHOR, STAMP)

    assert len(created) == 1
    assert names(paragraph) == ["del", "r"]
    assert created[0].get(W_AUTHOR) == AUTHOR
    assert created[0].get(W_DATE) == STAMP
    # The text is hidden, not removed: it is still there, as w:delText.
    assert texts(created[0]) == ["delText:Hello ", "delText:world"]
    assert visible_text(paragraph) == "!"


def test_tracked_delete_keeps_the_rpr_of_every_piece() -> None:
    pkg, paragraph = hosted(
        "<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>"
        "<w:r><w:rPr><w:u w:val='single'/></w:rPr><w:t>under</w:t></w:r>"
    )
    before = [properties(run) for run in paragraph.iter(W_R)]

    tracked_delete(pkg, resolve(paragraph, 0, 15, operation="delete"), AUTHOR, STAMP)

    assert [properties(run) for run in paragraph.iter(W_R)] == before
    assert before == list(dict.fromkeys(before))  # the three really do differ


def test_tracked_delete_keeps_the_rpr_when_the_range_cuts_a_run_in_two() -> None:
    pkg, paragraph = hosted("<w:r><w:rPr><w:b/></w:rPr><w:t>abcdef</w:t></w:r>")

    tracked_delete(pkg, resolve(paragraph, 2, 4, operation="delete"), AUTHOR, STAMP)

    assert visible_text(paragraph) == "abef"
    # Three pieces now, and every one of them is still bold.
    assert len(list(paragraph.iter(W_R))) == 3
    assert {properties(run) for run in paragraph.iter(W_R)} == {
        etree.tostring(
            paragraph_from("<w:rPr><w:b/></w:rPr>")[0], method="c14n", exclusive=True
        ).decode()
    }


def test_tracked_delete_under_another_authors_insertion_nests_del_in_ins() -> None:
    pkg, paragraph = hosted(
        "<w:r><w:t>A</w:t></w:r>"
        f'<w:ins w:id="900" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:t>BCD</w:t></w:r></w:ins>"
        "<w:r><w:t>E</w:t></w:r>"
    )

    created = tracked_delete(pkg, resolve(paragraph, 1, 4, operation="delete"), AUTHOR, STAMP)

    assert names(paragraph) == ["r", "ins", "r"]
    insertion = paragraph[1]
    assert insertion.get(W_ID) == "900"
    assert names(insertion) == ["del"]
    assert created == (insertion[0],)
    assert created[0].get(W_AUTHOR) == AUTHOR
    # Both revisions stand: the text is inserted by one author and deleted by
    # the other, so it shows for neither.
    assert visible_text(paragraph) == "AE"


def test_tracked_delete_under_the_same_authors_insertion_removes_the_run() -> None:
    pkg, paragraph = hosted(
        "<w:r><w:t>A</w:t></w:r>"
        f'<w:ins w:id="900" w:author="{AUTHOR}" w:date="{STAMP}">'
        "<w:r><w:t>BCD</w:t></w:r></w:ins>"
        "<w:r><w:t>E</w:t></w:r>"
    )

    created = tracked_delete(pkg, resolve(paragraph, 1, 4, operation="delete"), AUTHOR, STAMP)

    # Deleting one's own pending insertion undoes it rather than recording a
    # deletion of it; the emptied w:ins is kept, like ranges keeps a container.
    assert created == ()
    assert names(paragraph) == ["r", "ins", "r"]
    assert names(paragraph[1]) == []
    assert visible_text(paragraph) == "AE"


def test_tracked_delete_gives_adjacent_runs_one_del_but_never_crosses_a_marker() -> None:
    pkg, paragraph = hosted(
        "<w:r><w:t>ab</w:t></w:r>"
        "<w:r><w:t>cd</w:t></w:r>"
        '<w:bookmarkStart w:id="700" w:name="Inside"/>'
        "<w:r><w:t>ef</w:t></w:r>"
        '<w:bookmarkEnd w:id="700"/>'
    )

    created = tracked_delete(pkg, resolve(paragraph, 0, 6, operation="delete"), AUTHOR, STAMP)

    # Two deletions, not one: pulling the runs on either side of the bookmark
    # into a single w:del would have moved the marker out of the text it frames.
    assert len(created) == 2
    assert names(paragraph) == ["del", "bookmarkStart", "del", "bookmarkEnd"]
    assert len(set(ids_of(created))) == 2


def test_tracked_delete_refuses_pieces_that_were_only_resolved_for_reading() -> None:
    pkg, paragraph = hosted("<w:r><w:t>abcdef</w:t></w:r>")
    pieces = resolve(paragraph, 0, 3, operation="read")

    with pytest.raises(UnsupportedRevision, match="operation='delete'"):
        tracked_delete(pkg, pieces, AUTHOR, STAMP)

    assert visible_text(paragraph) == "abcdef"


def test_tracked_delete_refuses_an_empty_author() -> None:
    pkg, paragraph = hosted("<w:r><w:t>abcdef</w:t></w:r>")
    pieces = resolve(paragraph, 0, 3, operation="delete")

    with pytest.raises(ValueError, match="author"):
        tracked_delete(pkg, pieces, "", STAMP)


def test_tracked_delete_inherits_the_refusals_of_the_range_layer() -> None:
    pkg, paragraph = hosted(
        '<w:r><w:t xml:space="preserve">a </w:t></w:r>'
        '<w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple>'
        '<w:r><w:t xml:space="preserve"> b</w:t></w:r>'
    )

    with pytest.raises(UnsupportedRange) as failure:
        tracked_delete(
            pkg, resolve(paragraph, 1, 4, operation="delete"), AUTHOR, STAMP
        )
    assert failure.value.reason in {"crosses-field", "contains-field"}


# --------------------------------------------------------------------------------------
# tracked_insert
# --------------------------------------------------------------------------------------


def test_tracked_insert_wraps_the_cloned_run_in_an_ins() -> None:
    pkg, paragraph = hosted("<w:r><w:rPr><w:b/></w:rPr><w:t>abcdef</w:t></w:r>")

    created = tracked_insert(pkg, paragraph, 3, "XY", AUTHOR, STAMP)

    assert len(created) == 1
    assert created[0].tag == W_INS
    assert created[0].get(W_AUTHOR) == AUTHOR
    assert created[0].get(W_DATE) == STAMP
    assert visible_text(paragraph) == "abcXYdef"
    # The run is a clone of its neighbour's formatting, not a bare run.
    inserted = created[0][0]
    assert properties(inserted) == properties(paragraph[0])


def test_tracked_insert_writes_nothing_for_empty_text() -> None:
    pkg, paragraph = hosted("<w:r><w:t>abcdef</w:t></w:r>")

    assert tracked_insert(pkg, paragraph, 3, "", AUTHOR, STAMP) == ()
    assert names(paragraph) == ["r"]


def test_tracked_insert_inside_another_authors_insertion_splits_it() -> None:
    pkg, paragraph = hosted(
        f'<w:ins w:id="900" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:t>abcdef</w:t></w:r></w:ins>"
    )

    created = tracked_insert(pkg, paragraph, 3, "XY", AUTHOR, STAMP)

    assert names(paragraph) == ["ins", "ins", "ins"]
    assert visible_text(paragraph) == "abcXYdef"
    assert [element.get(W_AUTHOR) for element in paragraph] == [OTHER, AUTHOR, OTHER]
    assert created == (paragraph[1],)
    # Never a w:ins inside a w:ins, and never two revisions sharing an id.
    assert paragraph[1].find(W_INS) is None
    assert len(set(all_annotation_ids(pkg))) == len(all_annotation_ids(pkg))


def test_tracked_insert_at_the_edge_of_another_authors_insertion_does_not_split_it() -> None:
    pkg, paragraph = hosted(
        f'<w:ins w:id="900" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:t>abcdef</w:t></w:r></w:ins>"
    )

    tracked_insert(pkg, paragraph, 6, "XY", AUTHOR, STAMP)

    assert names(paragraph) == ["ins", "ins"]
    assert [element.get(W_AUTHOR) for element in paragraph] == [OTHER, AUTHOR]
    assert visible_text(paragraph) == "abcdefXY"


def test_tracked_insert_inside_the_authors_own_insertion_does_not_nest() -> None:
    pkg, paragraph = hosted(
        f'<w:ins w:id="900" w:author="{AUTHOR}" w:date="{STAMP}">'
        "<w:r><w:t>abcdef</w:t></w:r></w:ins>"
    )

    created = tracked_insert(pkg, paragraph, 3, "XY", AUTHOR, STAMP)

    # The text joins the insertion that already attributes it to this author.
    assert created == ()
    assert names(paragraph) == ["ins"]
    assert paragraph[0].find(W_INS) is None
    assert visible_text(paragraph) == "abcXYdef"


def test_tracked_insert_refuses_to_write_inside_deleted_content() -> None:
    pkg, paragraph = hosted(
        "<w:r><w:t>ab</w:t></w:r>"
        f'<w:del w:id="900" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:delText>gone</w:delText></w:r></w:del>"
        "<w:r><w:t>cd</w:t></w:r>"
    )
    before = etree.tostring(paragraph)

    tracked_insert(pkg, paragraph, 2, "XY", AUTHOR, STAMP)

    # ranges walks the point out of the w:del rather than writing invisible text.
    assert visible_text(paragraph) == "abXYcd"
    assert b"<w:delText>gone</w:delText>" in before
    assert paragraph.find(W_DEL).find(W_INS) is None


# --------------------------------------------------------------------------------------
# tracked_replace
# --------------------------------------------------------------------------------------


def test_tracked_replace_puts_the_deletion_before_the_insertion() -> None:
    pkg, paragraph = hosted(
        '<w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>'
    )

    created = tracked_replace(pkg, paragraph, 6, 11, "everyone", AUTHOR, STAMP)

    assert names(paragraph) == ["r", "del", "ins"]
    assert [etree.QName(element).localname for element in created] == ["del", "ins"]
    assert visible_text(paragraph) == "Hello everyone"
    # One user action, one timestamp, two distinct ids.
    assert {element.get(W_DATE) for element in created} == {STAMP}
    assert len(set(ids_of(created))) == 2


def test_tracked_replace_takes_the_rpr_of_the_first_run_of_the_range() -> None:
    pkg, paragraph = hosted(
        "<w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>def</w:t></w:r>"
    )

    created = tracked_replace(pkg, paragraph, 0, 6, "XYZ", AUTHOR, STAMP)

    inserted = next(element for element in created if element.tag == W_INS)
    assert properties(inserted[0]) == properties(paragraph.find(W_DEL)[0])
    assert "<w:b" in properties(inserted[0])


def test_tracked_replace_terminates_when_the_new_text_contains_the_old() -> None:
    """Regression: ``core/tracked_changes.py:200-271`` never returned here.

    Its loop re-searched the paragraph for the old text after writing the new
    one, so any replacement whose result still contains the pattern ran forever.
    The acceptance command carries ``--timeout=60`` for exactly this test.
    """
    pkg, paragraph = hosted('<w:r><w:t xml:space="preserve">the text here</w:t></w:r>')

    created = tracked_replace(pkg, paragraph, 4, 8, "text and more text", AUTHOR, STAMP)

    assert visible_text(paragraph) == "the text and more text here"
    assert len(created) == 2


def test_tracked_replace_with_empty_text_is_a_tracked_deletion() -> None:
    pkg, paragraph = hosted("<w:r><w:t>abcdef</w:t></w:r>")

    created = tracked_replace(pkg, paragraph, 0, 3, "", AUTHOR, STAMP)

    assert [etree.QName(element).localname for element in created] == ["del"]
    assert visible_text(paragraph) == "def"


def test_tracked_replace_leaves_a_package_every_checker_still_accepts(tmp_path) -> None:
    pkg = DocxPackage.open(untouched_bytes("combined"))
    matches = find(pkg, "paragraph")
    assert matches, "the combined fixture must contain the word searched for"
    match = matches[0]

    tracked_replace(pkg, match.paragraph, match.start, match.end, "section", AUTHOR, STAMP)

    path = pkg.save(tmp_path / "replaced.docx")
    assert validate_package(path) == []


# --------------------------------------------------------------------------------------
# Stamps and ids
# --------------------------------------------------------------------------------------


def test_the_default_date_is_an_iso_utc_stamp() -> None:
    pkg, paragraph = hosted("<w:r><w:t>abcdef</w:t></w:r>")

    created = tracked_insert(pkg, paragraph, 0, "X", AUTHOR)

    stamp = created[0].get(W_DATE)
    assert stamp.endswith("Z")
    assert len(stamp) == len("2026-01-02T03:04:05Z")
    # Parsable back, and expressed in UTC.
    parsed = datetime.fromisoformat(stamp)
    assert parsed.tzinfo is not None
    assert parsed.utcoffset() == UTC.utcoffset(None)


def test_an_aware_datetime_is_converted_to_utc() -> None:
    pkg, paragraph = hosted("<w:r><w:t>abcdef</w:t></w:r>")
    moment = datetime(2026, 1, 2, 5, 4, 5, tzinfo=timezone(timedelta(hours=2)))

    created = tracked_insert(pkg, paragraph, 0, "X", AUTHOR, moment)

    assert created[0].get(W_DATE) == STAMP


def test_ids_never_collide_with_the_ones_the_document_already_uses() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = set(all_annotation_ids(pkg))
    paragraph = paragraphs_of(pkg)[1]

    tracked_replace(pkg, paragraph, 0, 11, "Replaced text, ", AUTHOR, STAMP)

    after = all_annotation_ids(pkg)
    assert len(after) == len(set(after))
    assert set(after) > before


# --------------------------------------------------------------------------------------
# list_revisions
# --------------------------------------------------------------------------------------


def test_list_revisions_reports_every_kind_of_the_fixture() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    found = list_revisions(pkg)

    assert [(item.id, item.kind) for item in found] == [
        (201, "ins"),
        (202, "del"),
        (203, "ins"),
        (204, "del"),
        (205, "rPrChange"),
        (206, "pPrChange"),
        (207, "paragraph-mark-del"),
        (208, "paragraph-mark-ins"),
    ]
    assert {item.kind for item in found} <= REVISION_KINDS
    assert {item.story for item in found} == {"document"}
    assert {item.date for item in found} == {"2024-01-01T00:00:00Z"}


def test_list_revisions_reads_shown_text_for_an_insertion_and_hidden_text_for_a_deletion() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    by_id = {item.id: item for item in list_revisions(pkg)}

    assert by_id[201].text == "inserted text, "
    assert by_id[202].text == "deleted text, "
    # 203 inserted text that 204 then deleted: nothing of it shows any more, so
    # the insertion carries no visible text while the deletion carries it all.
    assert by_id[203].text == ""
    assert by_id[204].text == "inserted then deleted"
    assert by_id[204].author == "Fixture Reviewer"


def test_list_revisions_uses_the_v2_paragraph_index_of_find() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    by_id = {item.id: item for item in list_revisions(pkg)}
    matches = {match.paragraph: match.index for match in find(pkg, "inserted text")}

    assert len(matches) == 1
    paragraph, index = next(iter(matches.items()))
    assert by_id[201].paragraph_index == index
    assert visible_text(paragraph).startswith("Kept text, ")


def test_list_revisions_reports_no_index_for_a_paragraph_in_a_table_cell() -> None:
    pkg = DocxPackage.open(untouched_bytes("tables"))
    cell_paragraph = next(
        paragraph
        for paragraph in paragraphs_of(pkg)
        if any(a.tag == qn("w:tc") for a in paragraph.iterancestors())
        and visible_text(paragraph)
    )
    tracked_insert(pkg, cell_paragraph, 0, "New ", AUTHOR, STAMP)

    found = list_revisions(pkg)

    assert len(found) == 1
    assert found[0].story == "document"
    assert found[0].paragraph_index is None


def test_list_revisions_covers_every_story() -> None:
    pkg = DocxPackage.open(untouched_bytes("headers_footers"))
    header = next(story for story, _ in pkg.stories() if story.startswith("header"))
    tracked_insert(pkg, paragraphs_of(pkg, header)[0], 0, "Draft ", AUTHOR, STAMP)
    tracked_insert(pkg, paragraphs_of(pkg)[0], 0, "Draft ", AUTHOR, STAMP)

    found = list_revisions(pkg)

    assert {item.story for item in found} == {"document", header}
    # The main document always comes first.
    assert found[0].story == "document"


def test_list_revisions_reports_a_move_and_reads_its_text_from_the_right_side() -> None:
    # No fixture carries a move, and the two halves read differently: the text
    # left behind is hidden, the text moved to is shown.
    pkg, paragraph = hosted(
        f'<w:moveFrom w:id="910" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:delText>moved away</w:delText></w:r></w:moveFrom>"
        f'<w:moveTo w:id="911" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:t>moved here</w:t></w:r></w:moveTo>"
    )

    found = [item for item in list_revisions(pkg) if item.id in {910, 911}]

    assert [(item.id, item.kind, item.text) for item in found] == [
        (910, "moveFrom", "moved away"),
        (911, "moveTo", "moved here"),
    ]
    assert visible_text(paragraph) == "moved here"


def test_accept_refuses_a_move_rather_than_applying_half_of_it() -> None:
    pkg, _ = hosted(
        f'<w:moveFrom w:id="910" w:author="{OTHER}" w:date="{STAMP}">'
        "<w:r><w:delText>moved away</w:delText></w:r></w:moveFrom>"
    )

    with pytest.raises(UnsupportedRevision, match="910 \\(moveFrom\\)"):
        accept(pkg, ids=[910])
    with pytest.raises(UnsupportedRevision, match="910 \\(moveFrom\\)"):
        reject(pkg, ids=[910])


def test_list_revisions_ignores_a_row_revision_but_accept_still_refuses_it() -> None:
    pkg = DocxPackage.open(untouched_bytes("tables"))
    row = next(iter(dict(pkg.stories())["document"].iter(qn("w:tr"))))
    row.insert(
        0,
        paragraph_from(
            f'<w:trPr><w:ins w:id="950" w:author="{OTHER}" w:date="{STAMP}"/></w:trPr>'
        )[0],
    )

    assert list_revisions(pkg) == []
    with pytest.raises(UnsupportedRevision, match="950"):
        accept(pkg)


# --------------------------------------------------------------------------------------
# accept / reject: the supported kinds
# --------------------------------------------------------------------------------------


def test_accept_unwraps_an_insertion_and_keeps_its_text() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[1]

    applied = accept(pkg, ids=[201])

    assert [item.id for item in applied] == [201]
    assert paragraph.find(W_INS) is None
    assert visible_text(paragraph) == "Kept text, inserted text, and kept tail."


def test_reject_removes_an_insertion_and_its_text() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[1]

    reject(pkg, ids=[201])

    assert paragraph.find(W_INS) is None
    assert visible_text(paragraph) == "Kept text, and kept tail."


def test_accept_removes_a_deletion_with_its_content() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[1]

    accept(pkg, ids=[202])

    assert paragraph.find(W_DEL) is None
    assert b"deleted text" not in etree.tostring(paragraph)
    assert visible_text(paragraph) == "Kept text, inserted text, and kept tail."


def test_reject_brings_deleted_text_back_as_visible_text() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[1]

    reject(pkg, ids=[202])

    assert paragraph.find(W_DEL) is None
    assert qn("w:delText") not in {node.tag for node in paragraph.iter()}
    assert visible_text(paragraph) == "Kept text, inserted text, deleted text, and kept tail."


def test_accepting_a_deletion_keeps_the_markers_it_framed() -> None:
    pkg, paragraph = hosted(
        '<w:r><w:t xml:space="preserve">Before </w:t></w:r>'
        f'<w:del w:id="900" w:author="{OTHER}" w:date="{STAMP}">'
        '<w:bookmarkStart w:id="700" w:name="Framed"/>'
        "<w:r><w:delText>gone</w:delText></w:r>"
        '<w:commentRangeEnd w:id="701"/>'
        "</w:del>"
        '<w:r><w:t xml:space="preserve"> after</w:t></w:r>'
    )

    accept(pkg, ids=[900])

    # The text goes, the anchors stay exactly where the deletion stood.
    assert names(paragraph) == ["r", "bookmarkStart", "commentRangeEnd", "r"]
    assert visible_text(paragraph) == "Before  after"


def test_rejecting_an_insertion_keeps_the_markers_it_framed() -> None:
    pkg, paragraph = hosted(
        '<w:r><w:t xml:space="preserve">Before </w:t></w:r>'
        f'<w:ins w:id="900" w:author="{OTHER}" w:date="{STAMP}">'
        '<w:bookmarkStart w:id="700" w:name="Framed"/>'
        "<w:r><w:t>new</w:t></w:r>"
        '<w:bookmarkEnd w:id="700"/>'
        "</w:ins>"
        '<w:r><w:t xml:space="preserve"> after</w:t></w:r>'
    )

    reject(pkg, ids=[900])

    assert names(paragraph) == ["r", "bookmarkStart", "bookmarkEnd", "r"]
    assert visible_text(paragraph) == "Before  after"


def test_accepting_a_deleted_paragraph_mark_merges_the_two_paragraphs() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = [visible_text(p) for p in paragraphs_of(pkg)]

    accept(pkg, ids=[207])

    after = [visible_text(p) for p in paragraphs_of(pkg)]
    assert len(after) == len(before) - 1
    assert after[-1] == before[-2] + before[-1]


def test_accepting_a_deleted_paragraph_mark_keeps_the_next_paragraphs_properties() -> None:
    pkg, first = hosted(
        f"<w:pPr><w:rPr><w:del w:id=\"900\" w:author=\"{OTHER}\" w:date=\"{STAMP}\"/></w:rPr>"
        '<w:jc w:val="center"/></w:pPr>'
        "<w:r><w:t>one</w:t></w:r>"
    )
    body = first.getparent()
    second = paragraph_from('<w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>two</w:t></w:r>')
    body.insert(body.index(first) + 1, second)

    accept(pkg, ids=[900])

    assert first.getparent() is None
    assert visible_text(second) == "onetwo"
    # The surviving mark is the second paragraph's, so its properties govern.
    assert second.find(qn("w:pPr")).find(qn("w:jc")).get(qn("w:val")) == "right"


def test_rejecting_a_deleted_paragraph_mark_leaves_the_paragraphs_apart() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = [visible_text(p) for p in paragraphs_of(pkg)]

    reject(pkg, ids=[207])

    assert [visible_text(p) for p in paragraphs_of(pkg)] == before
    assert list_revisions(pkg) == [
        item for item in list_revisions(DocxPackage.open(untouched_bytes("tracked_changes")))
        if item.id != 207
    ]


def test_accepting_an_inserted_paragraph_mark_keeps_the_paragraph_it_added() -> None:
    source, tracked, plain = split_by_an_inserted_mark()
    pkg = DocxPackage.open(tracked)

    assert [item.kind for item in accept(pkg)] == ["paragraph-mark-ins"]

    assert list_revisions(pkg) == []
    after = snapshot(pkg.to_bytes())
    assert diff(snapshot(plain), after).is_empty()
    # And really kept the split: the fixture before it is a different document.
    assert not diff(snapshot(source), after).is_empty()


def test_rejecting_an_inserted_paragraph_mark_restores_the_document_exactly() -> None:
    source, tracked, _ = split_by_an_inserted_mark()
    pkg = DocxPackage.open(tracked)

    assert [item.kind for item in reject(pkg)] == ["paragraph-mark-ins"]

    assert list_revisions(pkg) == []
    assert diff(snapshot(source), snapshot(pkg.to_bytes())).is_empty()


def test_rejecting_two_consecutive_inserted_marks_leaves_one_paragraph() -> None:
    pkg, first = hosted(
        f'<w:pPr><w:rPr><w:ins w:id="900" w:author="{OTHER}" w:date="{STAMP}"/></w:rPr>'
        '<w:jc w:val="left"/></w:pPr><w:r><w:t>one</w:t></w:r>'
    )
    body = first.getparent()
    second = paragraph_from(
        f'<w:pPr><w:rPr><w:ins w:id="901" w:author="{OTHER}" w:date="{STAMP}"/></w:rPr>'
        '<w:jc w:val="center"/></w:pPr><w:r><w:t>two</w:t></w:r>'
    )
    body.insert(body.index(first) + 1, second)
    third = paragraph_from('<w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>three</w:t></w:r>')
    body.insert(body.index(second) + 1, third)

    assert [item.id for item in reject(pkg, ids=[900, 901])] == [900, 901]

    assert first.getparent() is None
    assert second.getparent() is None
    assert visible_text(third) == "onetwothree"
    # The only surviving mark is the last one, so its properties govern.
    assert third.find(qn("w:pPr")).find(qn("w:jc")).get(qn("w:val")) == "right"


def test_accepting_the_fixtures_inserted_paragraph_mark_moves_no_text() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = [visible_text(p) for p in paragraphs_of(pkg)]

    assert [item.id for item in accept(pkg, ids=[208])] == [208]

    assert [visible_text(p) for p in paragraphs_of(pkg)] == before
    assert 208 not in {item.id for item in list_revisions(pkg)}


def test_reject_refuses_an_inserted_paragraph_mark_with_nothing_to_merge_into() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = snapshot(pkg.to_bytes())

    # 208 closes the last paragraph of the body: there is nothing after it to
    # merge into, so undoing the split is impossible and refused as a whole.
    with pytest.raises(UnsupportedRevision, match="merge into") as failure:
        reject(pkg, ids=[208])

    assert failure.value.ids == (208,)
    assert diff(before, snapshot(pkg.to_bytes())).is_empty()


def test_nested_revisions_are_applied_together() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[2]
    assert visible_text(paragraph) == "Before.  After."

    applied = reject(pkg, ids=[203, 204])

    # 204 sat inside 203 and went with it; it is still reported as applied.
    assert [item.id for item in applied] == [203, 204]
    assert {item.id for item in list_revisions(pkg)} == {201, 202, 205, 206, 207, 208}
    assert visible_text(paragraph) == "Before.  After."


def test_accepting_both_halves_of_an_insert_then_delete_removes_the_text() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[2]

    accept(pkg, ids=[203, 204])

    assert {item.id for item in list_revisions(pkg)} == {201, 202, 205, 206, 207, 208}
    assert visible_text(paragraph) == "Before.  After."
    assert b"inserted then deleted" not in etree.tostring(paragraph)


# --------------------------------------------------------------------------------------
# accept / reject: selection and refusals
# --------------------------------------------------------------------------------------


def test_accept_selects_by_author() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    applied = accept(pkg, author="Fixture Reviewer")

    assert [item.id for item in applied] == [204]
    assert {item.id for item in list_revisions(pkg)} == {201, 202, 203, 205, 206, 207, 208}


def test_accept_combines_the_id_and_author_filters() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    with pytest.raises(LocatorError) as failure:
        accept(pkg, ids=[201], author="Fixture Reviewer")

    assert failure.value.code == "not-found"


def test_accept_refuses_an_id_the_package_does_not_have() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    with pytest.raises(LocatorError, match="9999") as failure:
        accept(pkg, ids=[201, 9999])

    assert failure.value.code == "not-found"
    assert {item.id for item in list_revisions(pkg)} == {201, 202, 203, 204, 205, 206, 207, 208}


def test_accept_refuses_property_revisions_and_names_their_ids() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    with pytest.raises(UnsupportedRevision) as failure:
        accept(pkg, ids=[201, 205, 206])

    message = str(failure.value)
    assert "205 (rPrChange)" in message
    assert "206 (pPrChange)" in message
    assert failure.value.ids == (205, 206)


def test_a_refusal_changes_nothing_at_all() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = snapshot(pkg.to_bytes())

    with pytest.raises(UnsupportedRevision):
        accept(pkg)

    assert diff(before, snapshot(pkg.to_bytes())).is_empty()


def test_accepting_everything_is_refused_while_a_kind_is_unsupported() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))

    with pytest.raises(UnsupportedRevision) as failure:
        accept(pkg)

    # The half-application of core/tracked_changes.py: it unwrapped the two
    # insertions, dropped the deletions and reported success while every
    # property revision was still there.
    assert failure.value.ids == (205, 206)


def test_accept_refuses_a_deleted_paragraph_mark_with_nothing_to_merge_into() -> None:
    pkg, paragraph = hosted(
        f"<w:pPr><w:rPr><w:del w:id=\"900\" w:author=\"{OTHER}\" w:date=\"{STAMP}\"/>"
        "</w:rPr></w:pPr><w:r><w:t>last</w:t></w:r>"
    )
    assert paragraph.getnext().tag == W_SECT_PR

    with pytest.raises(UnsupportedRevision, match="merge into"):
        accept(pkg, ids=[900])

    # Rejecting it is always possible: nothing has to be joined.
    assert [item.id for item in reject(pkg, ids=[900])] == [900]
    assert visible_text(paragraph) == "last"


def test_an_empty_selection_applies_nothing() -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    before = snapshot(pkg.to_bytes())

    assert accept(pkg, ids=[]) == []

    assert diff(before, snapshot(pkg.to_bytes())).is_empty()


def test_accept_applies_every_story_in_one_call() -> None:
    pkg = DocxPackage.open(untouched_bytes("headers_footers"))
    header = next(story for story, _ in pkg.stories() if story.startswith("header"))
    tracked_insert(pkg, paragraphs_of(pkg, header)[0], 0, "Draft ", AUTHOR, STAMP)
    tracked_insert(pkg, paragraphs_of(pkg)[0], 0, "Draft ", AUTHOR, STAMP)

    applied = accept(pkg, author=AUTHOR)

    assert {item.story for item in applied} == {"document", header}
    assert list_revisions(pkg) == []
    assert visible_text(paragraphs_of(pkg, header)[0]).startswith("Draft ")
    assert visible_text(paragraphs_of(pkg)[0]).startswith("Draft ")


# --------------------------------------------------------------------------------------
# Package scale: what a tracked edit must leave untouched
# --------------------------------------------------------------------------------------


def test_rejecting_a_tracked_replace_restores_the_document_exactly() -> None:
    reference = snapshot(untouched_bytes("tracked_changes"))
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    paragraph = paragraphs_of(pkg)[1]

    created = tracked_replace(pkg, paragraph, 0, 11, "Replaced text, ", AUTHOR, STAMP)
    assert visible_text(paragraph) == "Replaced text, inserted text, and kept tail."

    reject(pkg, ids=ids_of(created))

    assert diff(reference, snapshot(pkg.to_bytes())).is_empty()


def test_rejecting_a_tracked_delete_restores_the_document_exactly() -> None:
    reference = snapshot(untouched_bytes("bookmarks"))
    pkg = DocxPackage.open(untouched_bytes("bookmarks"))
    match = find(pkg, "Bookmarked ")[0]

    created = tracked_delete(
        pkg,
        resolve(match.paragraph, match.start, match.end, operation="delete"),
        AUTHOR,
        STAMP,
    )
    assert "Bookmarked " not in visible_text(match.paragraph)

    reject(pkg, ids=ids_of(created))

    assert diff(reference, snapshot(pkg.to_bytes())).is_empty()


def test_rejecting_a_tracked_insert_restores_the_document_exactly() -> None:
    reference = snapshot(untouched_bytes("hyperlinks"))
    pkg = DocxPackage.open(untouched_bytes("hyperlinks"))
    paragraph = paragraphs_of(pkg)[1]

    created = tracked_insert(pkg, paragraph, 0, "Note: ", AUTHOR, STAMP)
    assert visible_text(paragraph).startswith("Note: ")

    reject(pkg, ids=ids_of(created))

    assert diff(reference, snapshot(pkg.to_bytes())).is_empty()


def test_accepting_a_tracked_replace_reads_like_a_plain_replace() -> None:
    tracked = DocxPackage.open(untouched_bytes("tracked_changes"))
    plain = DocxPackage.open(untouched_bytes("tracked_changes"))

    created = tracked_replace(
        tracked, paragraphs_of(tracked)[1], 0, 11, "Replaced text, ", AUTHOR, STAMP
    )
    accept(tracked, ids=ids_of(created))
    replace_range(paragraphs_of(plain)[1], 0, 11, "Replaced text, ")

    assert visible_text(paragraphs_of(tracked)[1]) == visible_text(paragraphs_of(plain)[1])
    # The accepted pair left no trace of itself; the fixture's own revisions,
    # which the range did not touch, are all still there.
    assert {item.id for item in list_revisions(tracked)} == {201, 202, 203, 204, 205, 206, 207, 208}


def test_a_tracked_edit_disturbs_nothing_outside_the_paragraph_it_touched() -> None:
    reference = snapshot(untouched_bytes("combined"))
    pkg = DocxPackage.open(untouched_bytes("combined"))
    match = find(pkg, "paragraph")[0]
    index = next(
        position
        for position, paragraph in enumerate(paragraphs_of(pkg))
        if paragraph is match.paragraph
    )

    tracked_replace(pkg, match.paragraph, match.start, match.end, "section", AUTHOR, STAMP)

    assert_unchanged_except(
        reference,
        snapshot(pkg.to_bytes()),
        paragraphs={index},
        counters={"revisions"},
    )


def test_accepting_a_deleted_paragraph_mark_disturbs_no_other_story(tmp_path) -> None:
    pkg = DocxPackage.open(untouched_bytes("tracked_changes"))
    reference = snapshot(pkg.to_bytes())

    accept(pkg, ids=[207])

    after = snapshot(pkg.to_bytes())
    # Only the body changed, and only by losing one paragraph.
    assert set(diff(reference, after).parts_changed) == {"word/document.xml"}
    assert diff(reference, after).paragraphs_removed
    assert validate_package(pkg.save(tmp_path / "merged.docx")) == []
