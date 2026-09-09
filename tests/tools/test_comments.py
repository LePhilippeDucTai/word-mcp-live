"""Tests for the comment tools, once they run on the OOXML engine.

Two things are checked here that the old implementation could not deliver.
Writing: a comment is the five parts Word writes, not just ``comments.xml``, its
ids come from the engine's allocators so none is ever reused, and the anchored
runs stay exactly where they were -- inside their hyperlink, inside their table
cell.  Reading: every comment is joined back to its anchor, so the story, the
paragraph and the *text the comment is about* are reported, instead of the
``"Comment detected but content not accessible"`` placeholder the previous
fallback produced.
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from typing import Any

import pytest

from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.core.comment_writer import (
    COMMENTS_EXTENDED_PARTNAME,
    COMMENTS_IDS_PARTNAME,
    COMMENTS_PARTNAME,
    PEOPLE_PARTNAME,
    W16CID,
    add_comment_to_doc,
)
from word_document_server.core.comments import extract_all_comments
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.xmlns import qn
from word_document_server.tools.comment_tools import (
    get_all_comments,
    get_comments_by_author,
    get_comments_for_paragraph,
)

W_ID = qn("w:id")
W_COMMENT = qn("w:comment")
W_RANGE_START = qn("w:commentRangeStart")
W_RANGE_END = qn("w:commentRangeEnd")
W_REFERENCE = qn("w:commentReference")
W_HYPERLINK = qn("w:hyperlink")
W_P = qn("w:p")

#: The three side-car parts and the rels part a comment necessarily touches.
COMMENT_FAMILY_PARTS = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/people.xml",
    "word/_rels/document.xml.rels",
]


def run(coro: Any) -> Any:
    """Drive an ``async`` tool call to completion."""
    return asyncio.run(coro)


def comment_ids(pkg: DocxPackage) -> list[str]:
    """Ids of the comments stored in ``word/comments.xml``, in document order."""
    part = pkg.find_part(COMMENTS_PARTNAME)
    if part is None:
        return []
    return [element.get(W_ID) for element in pkg.root_of(part).iter(W_COMMENT)]


def marks(pkg: DocxPackage, tag: str) -> list[str]:
    """Ids of the `tag` marks carried by the main story."""
    return [element.get(W_ID) for element in pkg.document.iter(tag)]


@pytest.fixture
def simple(fixture_docx) -> Path:
    return fixture_docx("simple")


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------


def test_add_comment_writes_the_whole_comment_family(simple: Path) -> None:
    result = add_comment_to_doc(
        str(simple), "First paragraph", "Look here", author="Alice", initials="AL"
    )
    assert result["success"] is True

    pkg = DocxPackage.open(simple)
    identifier = str(result["comment_id"])

    # The text, keyed on the id the anchor marks carry.
    assert comment_ids(pkg) == [identifier]
    assert marks(pkg, W_RANGE_START) == [identifier]
    assert marks(pkg, W_RANGE_END) == [identifier]
    assert marks(pkg, W_REFERENCE) == [identifier]

    comment = next(pkg.root_of(pkg.part(COMMENTS_PARTNAME)).iter(W_COMMENT))
    assert comment.get(qn("w:author")) == "Alice"
    assert comment.get(qn("w:initials")) == "AL"
    para_id = next(comment.iter(W_P)).get(qn("w14:paraId"))
    assert para_id is not None

    # The thread state, the durable identity and the author, each in its part.
    extended = pkg.root_of(pkg.part(COMMENTS_EXTENDED_PARTNAME))
    entry = next(extended.iter(qn("w15:commentEx")))
    assert entry.get(qn("w15:paraId")) == para_id
    assert entry.get(qn("w15:done")) == "0"

    ids_root = pkg.root_of(pkg.part(COMMENTS_IDS_PARTNAME))
    durable = next(ids_root.iter(f"{{{W16CID}}}commentId"))
    assert durable.get(f"{{{W16CID}}}paraId") == para_id
    assert durable.get(f"{{{W16CID}}}durableId") is not None

    people = pkg.root_of(pkg.part(PEOPLE_PARTNAME))
    assert [p.get(qn("w15:author")) for p in people.iter(qn("w15:person"))] == ["Alice"]

    assert validate_package(simple) == []


def test_add_comment_allocates_a_fresh_id_and_para_id_each_time(simple: Path) -> None:
    first = add_comment_to_doc(str(simple), "First paragraph", "One", author="Alice")
    second = add_comment_to_doc(str(simple), "Second paragraph", "Two", author="Bob")
    assert first["success"] is True and second["success"] is True
    assert first["comment_id"] != second["comment_id"]

    pkg = DocxPackage.open(simple)
    assert comment_ids(pkg) == [str(first["comment_id"]), str(second["comment_id"])]

    root = pkg.root_of(pkg.part(COMMENTS_PARTNAME))
    para_ids = [p.get(qn("w14:paraId")) for p in root.iter(W_P)]
    assert len(set(para_ids)) == len(para_ids), "a w14:paraId was reused"

    # The parts and their relationships are created once, not once per comment.
    assert len(pkg.root_of(pkg.part(PEOPLE_PARTNAME))) == 2, "one w15:person per author"
    targets = [rel.target_ref for rel in pkg.document_part.rels.values()]
    assert targets.count("comments.xml") == 1
    assert validate_package(simple) == []


def test_add_comment_keeps_the_anchored_run_inside_its_hyperlink(fixture_docx) -> None:
    path = fixture_docx("hyperlinks")
    assert add_comment_to_doc(str(path), "example.org", "About the link")["success"] is True

    pkg = DocxPackage.open(path)
    linked = [
        run_element
        for run_element in pkg.document.iter(qn("w:r"))
        if "".join(t.text or "" for t in run_element.iter(qn("w:t"))) == "example.org"
    ]
    assert len(linked) == 1
    parents = list(linked[0].iterancestors(W_HYPERLINK))
    assert parents, "the commented run was taken out of its w:hyperlink"
    assert validate_package(path) == []


def test_add_comment_reports_a_missing_target_and_leaves_the_file_alone(simple: Path) -> None:
    before = simple.read_bytes()
    result = add_comment_to_doc(str(simple), "no such text", "Look here")
    assert result["success"] is False
    assert "not found" in result["error"]
    assert simple.read_bytes() == before


def test_add_comment_touches_nothing_but_the_target_and_the_comment_family(
    fixture_docx,
) -> None:
    path = fixture_docx("combined")
    before = snapshot(path.read_bytes())
    n_comments = len(before.story_paragraphs("comments"))
    target_index = next(
        i
        for i, sig in enumerate(before.story_paragraphs("document"))
        if sig.text == "Body text under the second heading."
    )

    assert add_comment_to_doc(str(path), "Body text under", "Probe")["success"] is True

    after = snapshot(path.read_bytes())
    assert_unchanged_except(
        before,
        after,
        paragraphs=[target_index, ("comments", n_comments)],
        parts=COMMENT_FAMILY_PARTS,
        counters=["comments", "comment_references"],
    )
    assert validate_package(path) == []


def test_add_comment_leaves_no_temporary_file_behind(simple: Path) -> None:
    assert add_comment_to_doc(str(simple), "First paragraph", "Look here")["success"] is True
    assert [p.name for p in simple.parent.iterdir()] == [simple.name]


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------


def test_extract_all_comments_resolves_anchors_to_stories_and_v2_indexes(
    fixture_docx,
) -> None:
    pkg = DocxPackage.open(fixture_docx("comments"))
    comments = {c["comment_id"]: c for c in extract_all_comments(pkg)}
    assert set(comments) == {"1", "2", "3"}

    root = comments["1"]
    assert root["story"] == "document", "the real story id, never the 'body' alias"
    assert root["author"] == "Fixture Author"
    assert root["text"] == "Root comment anchored on the first sentence."
    assert root["reference_text"] == "This sentence carries a comment thread."
    assert root["in_table"] is False
    assert isinstance(root["paragraph_index"], int)

    # A reply is anchored on the very same range as the comment it answers.
    assert comments["2"]["paragraph_index"] == root["paragraph_index"]
    assert comments["2"]["reference_text"] == root["reference_text"]

    # A range spanning two paragraphs reports both, joined by a newline.
    assert comments["3"]["reference_text"] == (
        "The third comment starts here\nand ends in the next paragraph."
    )
    assert comments["3"]["paragraph_index"] == root["paragraph_index"] + 1


def test_extract_all_comments_reports_a_table_cell_anchor_without_a_v2_index(
    fixture_docx,
) -> None:
    path = fixture_docx("tables")
    assert add_comment_to_doc(str(path), "Row two, column two", "In a cell")["success"] is True

    (comment,) = extract_all_comments(DocxPackage.open(path))
    assert comment["story"] == "document"
    assert comment["in_table"] is True
    assert comment["paragraph_index"] is None, "a table cell is outside the V2 index space"
    assert comment["reference_text"] == "Row two, column two"


def test_extract_all_comments_returns_nothing_when_the_package_has_no_comments(
    simple: Path,
) -> None:
    assert extract_all_comments(DocxPackage.open(simple)) == []


def test_reading_never_falls_back_to_a_placeholder(fixture_docx) -> None:
    body = run(get_all_comments(str(fixture_docx("comments"))))
    assert "not accessible" not in body, "the str(element) fallback is back"
    assert "Unknown" not in body, "an author was lost on the way"


# --------------------------------------------------------------------------
# The MCP tools
# --------------------------------------------------------------------------


def test_get_all_comments_keeps_its_response_shape(fixture_docx) -> None:
    payload = json.loads(run(get_all_comments(str(fixture_docx("comments")))))
    assert payload["success"] is True
    assert payload["total_comments"] == len(payload["comments"]) == 3
    assert set(payload["comments"][0]) >= {
        "id",
        "comment_id",
        "author",
        "initials",
        "date",
        "text",
        "paragraph_index",
        "in_table",
        "reference_text",
    }


def test_get_comments_by_author_filters_case_insensitively(fixture_docx) -> None:
    payload = json.loads(
        run(get_comments_by_author(str(fixture_docx("comments")), "fixture reviewer"))
    )
    assert payload["success"] is True
    assert payload["author"] == "fixture reviewer"
    assert [c["comment_id"] for c in payload["comments"]] == ["2"]
    assert payload["total_comments"] == 1


def test_get_comments_for_paragraph_uses_the_v2_index(fixture_docx) -> None:
    path = fixture_docx("comments")
    anchored = extract_all_comments(DocxPackage.open(path))[0]
    index = anchored["paragraph_index"]

    payload = json.loads(run(get_comments_for_paragraph(str(path), index)))
    assert payload["success"] is True
    assert payload["paragraph_index"] == index
    assert payload["paragraph_text"].startswith("This sentence carries a comment thread.")
    assert {c["comment_id"] for c in payload["comments"]} == {"1", "2"}


def test_get_comments_for_paragraph_refuses_an_index_out_of_range(simple: Path) -> None:
    payload = json.loads(run(get_comments_for_paragraph(str(simple), 999)))
    assert payload["success"] is False
    assert "out of range" in payload["error"]


def test_a_tool_reports_the_error_instead_of_swallowing_it(tmp_path: Path) -> None:
    broken = tmp_path / "broken.docx"
    broken.write_bytes(b"not a package at all")
    payload = json.loads(run(get_all_comments(str(broken))))
    assert payload["success"] is False
    assert "Failed to extract comments" in payload["error"]
    assert payload["error"].strip() != "Failed to extract comments: "
