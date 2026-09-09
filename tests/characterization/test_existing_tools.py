"""Characterization of the existing cross-platform mutating tools.

For every tool listed in the plan, this module performs one minimal
operation on the ``combined`` fixture (built fresh per test via the
``combined_path`` fixture below) and compares the package before and after
with :mod:`tests.support.snapshot`. Tools that only touch what they were
asked to touch get a plain assertion via
:func:`tests.support.snapshot.assert_unchanged_except` plus
:func:`tests.support.package_check.validate_package`. Tools with a known
destructive or broken behaviour get a precise, targeted assertion of what
is lost or never happens, wrapped in ``xfail(strict=True)`` so a fix in J03
turns the failure into an unexpected pass that must be caught and this test
updated. See ``docs/audit/destructive-ops.md`` for the human-readable
summary generated from these xfail reasons.

Every tool under test is ``async``; :func:`_run` drives it with
``asyncio.run``. Paragraph identity is established by searching for the
literal, known text of the fixture rather than a hard-coded index: three
incompatible paragraph-index spaces coexist in this repository (python-docx
body paragraphs, ``//w:p``, ``body//w:p``), plus the even wider snapshot
space (``PARAGRAPH_INDEX_SPACE`` in ``tests/support/snapshot.py``), so an
index is only meaningful once its space is named. ``_index_of`` resolves a
paragraph in the *snapshot* space (what ``assert_unchanged_except`` wants);
``_docx_index_of`` resolves one in python-docx's own body-paragraph space
(what tools such as ``add_bookmark`` or ``delete_paragraph`` take as their
``paragraph_index`` argument). The two coincide only before the fixture's
first table/content-control, which is where every target used below is
picked from.
"""

from __future__ import annotations

import asyncio
import dataclasses
import json
import zipfile
from typing import Any

import pytest
from docx import Document as PDocument

from tests.support.package_check import validate_package
from tests.support.snapshot import (
    ParagraphSignature,
    Snapshot,
    assert_unchanged_except,
    snapshot,
)
from word_document_server.tools.comment_write_tools import add_comment
from word_document_server.tools.content_tools import (
    add_heading,
    add_paragraph,
    add_table_of_contents,
    delete_paragraph,
    insert_numbered_list_near_text_tool,
    replace_block_between_manual_anchors_tool,
    replace_paragraph_block_below_header_tool,
    search_and_replace,
)
from word_document_server.tools.document_tools import merge_documents
from word_document_server.tools.footnote_tools import add_footnote_robust_tool
from word_document_server.tools.format_tools import (
    create_custom_style,
    format_table_cell_text,
    format_text,
)
from word_document_server.tools.hyperlink_tools import manage_hyperlinks
from word_document_server.tools.layout_tools import (
    add_bookmark,
    add_header_footer,
    set_paragraph_spacing,
)
from word_document_server.tools.protection_tools import (
    protect_document,
    unprotect_document,
)
from word_document_server.tools.tracked_changes_tools import (
    accept_tracked_changes,
    reject_tracked_changes,
    track_delete,
    track_insert,
    track_replace,
)

pytestmark = pytest.mark.characterization


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def _run(coro: Any) -> Any:
    """Drive an ``async`` tool call to completion."""
    return asyncio.run(coro)


def _index_of(snap: Snapshot, text: str, story: str = "document") -> int:
    """Snapshot-space index of the paragraph whose text is exactly `text`."""
    matches = [
        i for i, sig in enumerate(snap.story_paragraphs(story)) if sig.text == text
    ]
    assert len(matches) == 1, (
        f"expected exactly one {story!r} paragraph with text {text!r} "
        f"(snapshot space), found at indices {matches}"
    )
    return matches[0]


def _docx_index_of(path, text: str) -> int:
    """python-docx body-paragraph-space index of the paragraph with `text`."""
    paragraphs = PDocument(str(path)).paragraphs
    matches = [i for i, p in enumerate(paragraphs) if p.text == text]
    assert len(matches) == 1, (
        f"expected exactly one python-docx body paragraph with text {text!r}, "
        f"found at indices {matches}"
    )
    return matches[0]


def _reindex_after_merge(
    after: Snapshot, absorbed_index: int, story: str = "document"
) -> Snapshot:
    """Shift ``story`` paragraph keys at or past ``absorbed_index`` by +1.

    Accepting a deleted paragraph mark or rejecting an inserted one merges
    two paragraphs and removes one ``w:p`` from ``story`` (D-008): every
    paragraph that followed the absorbed one now sits one key earlier in
    ``after``'s snapshot space than the same physical paragraph does in
    ``before``'s. Shifting those keys back by one re-aligns the two
    snapshots so :func:`~tests.support.snapshot.assert_unchanged_except` can
    compare them key by key; the absorbed key is then correctly reported as
    removed instead of every following paragraph being reported as changed.
    ``tables`` keys are numbered separately and never shift.
    """
    reindexed: dict[tuple[str, int], ParagraphSignature] = {}
    for (key_story, index), sig in after.paragraphs.items():
        if key_story == story and index >= absorbed_index:
            new_index = index + 1
            reindexed[(key_story, new_index)] = dataclasses.replace(
                sig, index=new_index
            )
        else:
            reindexed[(key_story, index)] = sig
    return dataclasses.replace(after, paragraphs=reindexed)


@pytest.fixture
def combined_path(fixture_docx):
    """A fresh ``combined.docx`` fixture, one per test."""
    return fixture_docx("combined")


# --------------------------------------------------------------------------
# Non-destructive tools: touch only what they were asked to touch
# --------------------------------------------------------------------------


def test_search_and_replace_touches_only_the_matched_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Plain then bold bold italic et la fin.")

    result = _run(search_and_replace(str(combined_path), "Plain then", "Changed then"))
    assert "Replaced 1" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[idx])
    assert validate_package(combined_path) == []


def test_delete_paragraph_removes_only_the_last_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    last_snapshot_idx = len(before.story_paragraphs("document")) - 1
    # delete_paragraph indexes python-docx's own body-paragraph space; it
    # coincides with the widest snapshot space at the very tail of the
    # document, which is not nested inside any table or content control.
    last_docx_idx = len(PDocument(str(combined_path)).paragraphs) - 1

    result = _run(delete_paragraph(str(combined_path), last_docx_idx))
    assert "deleted successfully" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[last_snapshot_idx])
    assert validate_package(combined_path) == []


def test_add_heading_only_appends_a_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    n = len(before.story_paragraphs("document"))

    result = _run(add_heading(str(combined_path), "Probe heading", level=2))
    assert "added to" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[n])
    assert validate_package(combined_path) == []


def test_add_paragraph_only_appends_a_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    n = len(before.story_paragraphs("document"))

    result = _run(add_paragraph(str(combined_path), "Probe paragraph text"))
    assert "added to" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[n])
    assert validate_package(combined_path) == []


def test_add_bookmark_only_adds_markers_and_bumps_the_counter(combined_path):
    target_text = "First paragraph of the simple fixture."
    before = snapshot(combined_path.read_bytes())
    snapshot_idx = _index_of(before, target_text)
    docx_idx = _docx_index_of(combined_path, target_text)

    result = _run(add_bookmark(str(combined_path), docx_idx, "ProbeBookmark"))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(
        before, after, paragraphs=[snapshot_idx], counters=["bookmarks"]
    )
    assert validate_package(combined_path) == []


def test_set_paragraph_spacing_touches_only_the_targeted_paragraph(combined_path):
    target_text = "First paragraph of the simple fixture."
    before = snapshot(combined_path.read_bytes())
    snapshot_idx = _index_of(before, target_text)
    docx_idx = _docx_index_of(combined_path, target_text)

    result = _run(
        set_paragraph_spacing(
            str(combined_path), paragraph_index=docx_idx, space_before_pt=12
        )
    )
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    # ppr is part of ParagraphSignature and does cover spacing, but this
    # tool writes spacing via w:pPr/w:spacing: listing the paragraph as
    # allowed is exactly what covers that change, nothing more.
    assert_unchanged_except(before, after, paragraphs=[snapshot_idx])
    assert validate_package(combined_path) == []


def test_add_comment_touches_only_the_target_paragraph_and_comments_part(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Body text under the second heading.")
    n_comments = len(before.story_paragraphs("comments"))

    result = _run(
        add_comment(
            str(combined_path),
            "Body text under the second heading",
            "A probe comment",
        )
    )
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(
        before,
        after,
        paragraphs=[idx, ("comments", n_comments)],
        # Since J03-P3 a comment is written as Word writes one: the text in
        # comments.xml, the thread state in commentsExtended.xml, the durable
        # identity in commentsIds.xml and the author in people.xml, each
        # related from word/document.xml.  The three side-car parts and the
        # rels part are therefore expected to change; the document itself is
        # still guarded paragraph by paragraph.
        parts=[
            "word/commentsExtended.xml",
            "word/commentsIds.xml",
            "word/people.xml",
            "word/_rels/document.xml.rels",
        ],
        counters=["comments", "comment_references"],
    )
    assert validate_package(combined_path) == []


def test_track_replace_touches_only_the_target_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Kept text, inserted text, and kept tail.")

    result = _run(track_replace(str(combined_path), "Kept text,", "Changed text,"))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[idx], counters=["revisions"])
    assert validate_package(combined_path) == []


def test_track_insert_touches_only_the_target_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Kept text, inserted text, and kept tail.")

    result = _run(track_insert(str(combined_path), "and kept tail.", " Inserted."))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[idx], counters=["revisions"])
    assert validate_package(combined_path) == []


def test_track_delete_touches_only_the_target_paragraph(combined_path):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Kept text, inserted text, and kept tail.")

    result = _run(track_delete(str(combined_path), "Kept text, "))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[idx], counters=["revisions"])
    assert validate_package(combined_path) == []


def test_accept_tracked_changes_touches_only_the_paragraphs_with_ins_or_del(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    idx_plain = _index_of(before, "Kept text, inserted text, and kept tail.")
    idx_nested = _index_of(before, "Before.  After.")
    # Paragraph-mark revisions: the deleted mark's <w:del> and the inserted
    # mark's <w:ins> both live in w:pPr/w:rPr, so accepting touches their
    # ppr too. Accepting the deleted mark also merges its paragraph with the
    # next one (D-008): idx_ins_mark's text joins idx_del_mark's paragraph
    # and idx_ins_mark's own w:p is removed, so it is the absorbed key --
    # every paragraph that followed it shifts down by one key in `after`.
    idx_del_mark = _index_of(
        before, "This paragraph mark is deleted, so this merges with the next one."
    )
    idx_ins_mark = _index_of(before, "This paragraph mark is inserted.")
    del_mark_text = before.paragraphs[("document", idx_del_mark)].text
    ins_mark_text = before.paragraphs[("document", idx_ins_mark)].text

    result = _run(accept_tracked_changes(str(combined_path)))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    merged = after.paragraphs[("document", idx_del_mark)]
    assert merged.text == del_mark_text + ins_mark_text

    after_reindexed = _reindex_after_merge(after, idx_ins_mark)
    assert_unchanged_except(
        before,
        after_reindexed,
        paragraphs=[idx_plain, idx_nested, idx_del_mark, idx_ins_mark],
        counters=["revisions"],
    )
    assert validate_package(combined_path) == []


def test_reject_tracked_changes_touches_only_the_paragraphs_with_ins_or_del(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    idx_plain = _index_of(before, "Kept text, inserted text, and kept tail.")
    idx_nested = _index_of(before, "Before.  After.")
    # Same paragraph-mark paragraphs as above: rejecting strips <w:del> and
    # <w:ins> from their ppr too. Rejecting the inserted mark also merges its
    # paragraph with the one that follows it (D-008): the following
    # paragraph's text joins idx_ins_mark's paragraph and the following
    # paragraph's own w:p is removed, so it is the absorbed key -- every
    # paragraph after it shifts down by one key in `after`.
    idx_del_mark = _index_of(
        before, "This paragraph mark is deleted, so this merges with the next one."
    )
    idx_ins_mark = _index_of(before, "This paragraph mark is inserted.")
    idx_absorbed = idx_ins_mark + 1
    ins_mark_text = before.paragraphs[("document", idx_ins_mark)].text
    absorbed_text = before.paragraphs[("document", idx_absorbed)].text

    result = _run(reject_tracked_changes(str(combined_path)))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    merged = after.paragraphs[("document", idx_ins_mark)]
    assert merged.text == ins_mark_text + absorbed_text

    after_reindexed = _reindex_after_merge(after, idx_absorbed)
    assert_unchanged_except(
        before,
        after_reindexed,
        paragraphs=[idx_plain, idx_nested, idx_del_mark, idx_ins_mark, idx_absorbed],
        counters=["revisions"],
    )
    assert validate_package(combined_path) == []


def test_manage_hyperlinks_touches_only_the_target_paragraph_and_rels(combined_path):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Second heading")

    result = _run(
        manage_hyperlinks(
            str(combined_path),
            action="add",
            text="Second heading",
            url="https://example.org/probe",
        )
    )
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(
        before,
        after,
        paragraphs=[idx],
        parts=["word/_rels/document.xml.rels"],
        counters=["hyperlinks"],
    )
    assert validate_package(combined_path) == []


def test_add_footnote_robust_touches_only_the_target_paragraph_and_footnotes_part(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Body text under the second heading.")
    n_footnotes = len(before.story_paragraphs("footnotes"))

    result = _run(
        add_footnote_robust_tool(
            str(combined_path),
            search_text="Body text under the second heading",
            footnote_text="Probe footnote.",
        )
    )
    assert result["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(
        before,
        after,
        paragraphs=[idx, ("footnotes", n_footnotes)],
        counters=["footnote_references", "footnotes"],
    )
    assert validate_package(combined_path) == []


def test_format_table_cell_text_touches_only_the_target_cell_paragraph(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Row two, column two")

    result = _run(format_table_cell_text(str(combined_path), 0, 1, 1, bold=True))
    assert "formatted successfully" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[idx])
    assert validate_package(combined_path) == []


def test_insert_numbered_list_near_text_only_appends_at_the_target(combined_path):
    before = snapshot(combined_path.read_bytes())
    n = len(before.story_paragraphs("document"))
    # Target the very last paragraph: inserting mid-document would shift the
    # snapshot-space index of every following paragraph, which is reindex
    # churn (nothing is lost), not something assert_unchanged_except's
    # exception list is meant to enumerate one by one.
    last_docx_idx = len(PDocument(str(combined_path)).paragraphs) - 1

    result = _run(
        insert_numbered_list_near_text_tool(
            str(combined_path),
            target_paragraph_index=last_docx_idx,
            list_items=["Item A", "Item B"],
        )
    )
    assert "inserted after paragraph" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[n, n + 1])
    assert validate_package(combined_path) == []


def test_replace_paragraph_block_below_header_touches_only_the_block(combined_path):
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, "Body text under the second heading.")

    result = _run(
        replace_paragraph_block_below_header_tool(
            str(combined_path), "Second heading", ["Replacement paragraph."]
        )
    )
    assert "Replaced content under" in result

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[idx])
    assert validate_package(combined_path) == []


def test_protect_then_unprotect_roundtrips_bytes(combined_path):
    before_bytes = combined_path.read_bytes()

    protect_result = _run(protect_document(str(combined_path), "secret123"))
    assert "encrypted successfully" in protect_result
    # While protected the file is an OLE/CFB container, not an OPC package:
    # this is the expected effect of encryption, not data loss.
    protected_issues = validate_package(combined_path)
    assert any(issue.code == "ZIP-UNREADABLE" for issue in protected_issues)

    unprotect_result = _run(unprotect_document(str(combined_path), "secret123"))
    assert "decrypted successfully" in unprotect_result

    assert combined_path.read_bytes() == before_bytes
    assert validate_package(combined_path) == []


# --------------------------------------------------------------------------
# Destructive or broken tools: pinned via a targeted, currently-failing
# assertion. See docs/audit/destructive-ops.md for the human-readable
# summary of these reasons.
# --------------------------------------------------------------------------


# Was xfail(strict=True) until J03-P2: accept_tracked_changes used to only
# strip the <w:del> marker from a deleted paragraph mark's w:pPr/w:rPr
# (root.iter(W("del")) matched it exactly like a run-level deletion); it
# never merged the paragraph with the following one the way Word does when a
# deleted paragraph mark is accepted. The revision was recorded as accepted
# (the marker gone, the counter dropped) but its actual effect -- one fewer
# paragraph break -- was never applied. engine/revisions.py now performs the
# merge via _merge_into_next (D-008).
def test_accept_tracked_changes_merges_a_deleted_paragraph_mark(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    n_before = len(before.story_paragraphs("document"))

    result = _run(accept_tracked_changes(str(combined_path)))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    n_after = len(after.story_paragraphs("document"))
    assert n_after == n_before - 1, (
        "accepting the deleted paragraph mark should merge "
        "'This paragraph mark is deleted...' into the paragraph that "
        "follows it, removing one paragraph from the document story"
    )


# Was xfail(strict=True) until J03-P2: reject_tracked_changes had the
# symmetric bug -- rejecting an inserted paragraph mark should undo the split
# it introduced and merge the paragraph back with the one that follows, the
# way Word does. The tool used to only strip the <w:ins> marker from
# w:pPr/w:rPr (root.iter(W("ins")) matched it exactly like a run-level
# insertion, and parent.remove(ins) was a no-op merge because the element had
# no children to reinsert), leaving the paragraph break in place.
# engine/revisions.py now performs the merge via _merge_into_next (D-008).
def test_reject_tracked_changes_merges_an_inserted_paragraph_mark(
    combined_path,
):
    before = snapshot(combined_path.read_bytes())
    n_before = len(before.story_paragraphs("document"))

    result = _run(reject_tracked_changes(str(combined_path)))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    n_after = len(after.story_paragraphs("document"))
    assert n_after == n_before - 1, (
        "rejecting the inserted paragraph mark should merge "
        "'This paragraph mark is inserted.' back into the paragraph that "
        "follows it, removing one paragraph from the document story"
    )


# Was xfail(strict=True) until J03-P4: format_text used to clear every run of
# the paragraph (run.clear()) and rebuild it as before/target/after runs, which
# lost the formatting of the text outside the requested range. It now patches
# the w:rPr of the runs the range resolves to, so the case below passes; the
# wider consequences of the rebuild are covered by tests/tools/test_format_text.py.
def test_format_text_keeps_untouched_run_formatting(combined_path):
    target_text = "Plain then bold bold italic et la fin."
    before = snapshot(combined_path.read_bytes())
    idx = _index_of(before, target_text)
    docx_idx = _docx_index_of(combined_path, target_text)

    # Format only "Plain" (positions 0:5); "bold " (positions 11:16) is
    # untouched and should keep its <w:b/> formatting.
    result = _run(format_text(str(combined_path), docx_idx, 0, 5, bold=True))
    assert "formatted successfully" in result

    after = snapshot(combined_path.read_bytes())
    bold_runs = [
        run
        for run in after.paragraphs[("document", idx)].runs
        if run.text == "bold " and "<w:b" in run.rpr
    ]
    assert bold_runs, "the untouched 'bold ' run should keep its <w:b/> formatting"


@pytest.mark.xfail(
    strict=True,
    reason=(
        "add_table_of_contents rebuilds the whole document in a blank "
        "Document(): only paragraph.text/cell.text and style names survive. "
        "comments.xml, commentsExtended.xml, footnotes.xml, endnotes.xml, "
        "headers/footers, images, hyperlinks, bookmarks, fields, tracked "
        "changes and table merges are all dropped. Already flagged "
        "DESTRUCTIVE in the tool docstring (J01-P1)."
    ),
)
def test_add_table_of_contents_drops_ancillary_parts(combined_path):
    before = snapshot(combined_path.read_bytes())
    assert "word/comments.xml" in before.parts  # sanity: the fixture has one

    result = _run(add_table_of_contents(str(combined_path)))
    assert "Table of contents" in result

    after = snapshot(combined_path.read_bytes())
    assert "word/comments.xml" in after.parts


def test_merge_documents_drops_ancillary_parts(tmp_path, combined_path):
    before = snapshot(combined_path.read_bytes())
    assert "word/comments.xml" in before.parts  # sanity: the source has one

    target = tmp_path / "merged.docx"
    result = _run(
        merge_documents(str(target), [str(combined_path)], add_page_breaks=False)
    )
    assert "Successfully merged" in result

    after = snapshot(target.read_bytes())
    assert "word/comments.xml" in after.parts


@pytest.mark.xfail(
    strict=True,
    reason=(
        "add_header_footer clears every existing paragraph of the header "
        "(header.paragraphs -> p.clear()) before writing the new text: the "
        "fixture's header run carrying an inline picture is discarded even "
        "though the picture has nothing to do with the new header text."
    ),
)
def test_add_header_footer_clears_existing_header_content(combined_path):
    before = snapshot(combined_path.read_bytes())
    before_runs = before.paragraphs[("header1", 0)].runs
    assert len(before_runs) >= 2  # sanity: text run + picture run

    result = _run(add_header_footer(str(combined_path), header_text="New header text"))
    assert json.loads(result)["success"] is True

    after = snapshot(combined_path.read_bytes())
    after_runs = after.paragraphs[("header1", 0)].runs
    assert len(after_runs) == len(
        before_runs
    ), "existing header content (including the picture run) should survive"


@pytest.mark.xfail(
    strict=True,
    reason=(
        "core.styles.create_style checks for an existing style with "
        "doc.styles.get_by_id(style_name, WD_STYLE_TYPE.PARAGRAPH), but "
        "get_by_id never raises: it returns the type's default style "
        "('Normal') when the id is not found. The except: branch that would "
        "call doc.styles.add_style is therefore never reached, so no style "
        "is ever written to styles.xml even though the tool always reports "
        "success."
    ),
)
def test_create_custom_style_never_creates_the_style(combined_path):
    result = _run(create_custom_style(str(combined_path), "ProbeStyle", bold=True))
    assert "created successfully" in result

    with zipfile.ZipFile(combined_path) as zf:
        styles_xml = zf.read("word/styles.xml")
    assert b'w:styleId="ProbeStyle"' in styles_xml


@pytest.mark.xfail(
    strict=True,
    reason=(
        "replace_block_between_manual_anchors compares el.tag == CT_P.tag "
        "(and CT_Tbl.tag), but CT_P/CT_Tbl are lxml element classes "
        "registered via element_class_lookup: CT_P.tag reads the unbound "
        "'tag' attribute descriptor of lxml.etree._Element, never a string, "
        "so the comparison is always False. start_idx stays None and the "
        "tool unconditionally reports the start anchor as not found."
    ),
)
def test_replace_block_between_manual_anchors_never_finds_the_anchor(combined_path):
    before = snapshot(combined_path.read_bytes())
    assert "Second paragraph, with a trailing sentence." in before.text()  # sanity

    result = _run(
        replace_block_between_manual_anchors_tool(
            str(combined_path),
            "Second paragraph, with a trailing sentence.",
            ["Replacement between anchors."],
            end_anchor_text="Second heading",
        )
    )
    assert "not found" not in result


@pytest.mark.timeout(10)
# Was xfail(strict=True) until J03-P2: track_replace_in_doc's while True loop
# used to re-scan the whole paragraph after every replacement, including runs
# it had just inserted inside the new <w:ins>. When new_text contained
# old_text as a substring (e.g. "Risk" -> "Risk Risk"), the freshly inserted
# text matched again on the next iteration and the paragraph grew forever;
# the call never returned. core/tracked_changes.py now applies replacements
# right-to-left over the matches found in a single upfront scan, so a
# freshly inserted run can never be rescanned.
def test_track_replace_terminates_when_replacement_contains_original(tmp_path):
    from tests.fixtures.builders import build

    path = tmp_path / "risk.docx"
    path.write_bytes(build("simple"))
    _run(add_paragraph(str(path), "This is a Risk statement."))

    result = _run(track_replace(str(path), "Risk", "Risk Risk"))
    payload = json.loads(result)
    assert payload["success"] is True
    assert payload["replacements"] == 1

    after = snapshot(path.read_bytes())
    idx = _index_of(after, "This is a Risk Risk statement.")
    assert after.story_paragraphs("document")[idx].text == (
        "This is a Risk Risk statement."
    )
