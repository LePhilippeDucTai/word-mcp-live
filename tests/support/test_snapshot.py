"""Tests for the canonical snapshot and the structural comparison.

Documents are built in memory: the focused tests write raw OOXML here, and the
end-to-end coverage tests reuse :mod:`tests.fixtures.builders`, which builds the
same packages every test of the suite uses.  No binary is committed, and the
module under test never imports python-docx.

The last two sections are the regression guard of the review finding on
``assert_unchanged_except``: allowing one paragraph lifts the C14N digest of its
whole story part, so anything the per-paragraph and per-table signatures fail to
describe becomes invisible.  Each test there deletes a real carrier of meaning
from a real fixture and requires the assertion to notice.
"""

from __future__ import annotations

import io
import zipfile
from functools import cache

import pytest
from lxml import etree

from tests.fixtures.builders import build
from tests.support.snapshot import (
    MAIN_STORY,
    assert_unchanged_except,
    diff,
    snapshot,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
NS = f'xmlns:w="{W}" xmlns:r="{R}" xmlns:v="{V}"'

PKG = "application/vnd.openxmlformats-package."
WML = "application/vnd.openxmlformats-officedocument.wordprocessingml."
DOCUMENT_CT = WML + "document.main+xml"
HEADER_CT = WML + "header+xml"
COMMENTS_CT = WML + "comments+xml"

OFFICE_DOCUMENT_REL = R + "/officeDocument"
HEADER_REL = R + "/header"
COMMENTS_REL = R + "/comments"
IMAGE_REL = R + "/image"

DEFAULT_DEFAULTS = (
    ("rels", PKG + "relationships+xml"),
    ("xml", "application/xml"),
)


def content_types_xml(defaults, overrides) -> str:
    entries = "".join(
        f'<Default Extension="{ext}" ContentType="{ctype}"/>'
        for ext, ctype in defaults
    ) + "".join(
        f'<Override PartName="{part}" ContentType="{ctype}"/>'
        for part, ctype in overrides
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f"{entries}</Types>"
    )


def rels_xml(relationships) -> str:
    entries = "".join(
        f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"'
        + (f' TargetMode="{mode}"' if mode else "")
        + "/>"
        for rid, rtype, target, mode in relationships
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        f'relationships">{entries}</Relationships>'
    )


def document_xml(body: str, declaration: bool = True) -> str:
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' if declaration else ""
    return f"{head}<w:document {NS}><w:body>{body}</w:body></w:document>"


def make_docx(
    body: str = "",
    *,
    document: str | None = None,
    extra_parts: dict[str, str | bytes] | None = None,
    defaults=DEFAULT_DEFAULTS,
    overrides=(("/word/document.xml", DOCUMENT_CT),),
    document_rels=(),
    package_rels=(("rId1", OFFICE_DOCUMENT_REL, "word/document.xml", None),),
    content_types: str | None = None,
) -> bytes:
    """Build a minimal but well-formed .docx package in memory."""
    parts: dict[str, str | bytes] = {
        "[Content_Types].xml": content_types
        or content_types_xml(defaults, overrides),
        "_rels/.rels": rels_xml(package_rels),
        "word/document.xml": document if document is not None else document_xml(body),
        "word/_rels/document.xml.rels": rels_xml(document_rels),
    }
    parts.update(extra_parts or {})
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            payload = parts[name]
            archive.writestr(
                name, payload.encode("utf-8") if isinstance(payload, str) else payload
            )
    return buffer.getvalue()


def para(*children: str, style: str | None = None, num: tuple[str, str] | None = None) -> str:
    properties = ""
    if style is not None:
        properties += f'<w:pStyle w:val="{style}"/>'
    if num is not None:
        properties += (
            f'<w:numPr><w:ilvl w:val="{num[1]}"/>'
            f'<w:numId w:val="{num[0]}"/></w:numPr>'
        )
    prefix = f"<w:pPr>{properties}</w:pPr>" if properties else ""
    return f"<w:p>{prefix}{''.join(children)}</w:p>"


def run(text: str, rpr: str = "") -> str:
    properties = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{properties}<w:t xml:space="preserve">{text}</w:t></w:r>'


# --------------------------------------------------------------------------
# Baseline
# --------------------------------------------------------------------------


def test_identical_packages_produce_an_empty_diff():
    document = make_docx(para(run("Hello")) + para(run("World")))
    assert diff(snapshot(document), snapshot(document)).is_empty()


def test_snapshot_accepts_bytes_and_paths(tmp_path):
    document = make_docx(para(run("Hello")))
    path = tmp_path / "sample.docx"
    path.write_bytes(document)
    assert diff(snapshot(document), snapshot(path)).is_empty()
    assert diff(snapshot(str(path)), snapshot(document)).is_empty()


# --------------------------------------------------------------------------
# Text policy
# --------------------------------------------------------------------------


def test_visible_text_follows_the_engine_policy():
    body = para(
        run("plain"),
        "<w:r><w:tab/><w:t>after-tab</w:t><w:br/><w:t>after-br</w:t>"
        "<w:cr/><w:t>after-cr</w:t></w:r>",
        f'<w:hyperlink r:id="rId9">{run("|link")}</w:hyperlink>',
        f'<w:ins w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z">{run("|ins")}</w:ins>',
        f"<w:sdt><w:sdtPr><w:alias w:val=\"ignored\"/></w:sdtPr>"
        f"<w:sdtContent>{run('|sdt')}</w:sdtContent></w:sdt>",
        f'<w:smartTag w:element="x">{run("|smart")}</w:smartTag>',
        f'<w:fldSimple w:instr=" PAGE ">{run("|fld")}</w:fldSimple>',
        '<w:del w:id="2" w:author="a" w:date="2020-01-01T00:00:00Z">'
        "<w:r><w:delText>HIDDEN</w:delText></w:r></w:del>",
        '<w:r><w:instrText xml:space="preserve"> REF _Ref1 </w:instrText></w:r>',
    )
    sig = snapshot(make_docx(body)).paragraphs[(MAIN_STORY, 0)]
    assert sig.text == (
        "plain\tafter-tab\nafter-br\nafter-cr|link|ins|sdt|smart|fld"
    )
    assert "HIDDEN" not in sig.text
    assert "REF" not in sig.text
    # Visible runs reconstruct the visible text exactly.
    assert "".join(item.text for item in sig.runs) == sig.text


def test_deleted_runs_and_field_instructions_are_kept_aside():
    body = para(
        run("kept"),
        '<w:del w:id="2" w:author="a" w:date="2020-01-01T00:00:00Z">'
        '<w:r><w:rPr><w:b/></w:rPr><w:delText>gone</w:delText></w:r></w:del>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        + run("1")
        + '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '<w:fldSimple w:instr=" TIME "/>',
    )
    sig = snapshot(make_docx(body)).paragraphs[(MAIN_STORY, 0)]
    assert sig.deleted_text == "gone"
    assert [item.rpr for item in sig.deleted_runs] != [""]
    assert sig.field_instructions == (" PAGE ", " TIME ")
    assert [key for name, key in sig.markers if name == "fldChar"] == [
        "begin",
        "separate",
        "end",
    ]
    assert sig.text == "kept1"


def test_deleted_text_change_is_reported_even_though_it_is_invisible():
    def build(deleted: str) -> bytes:
        return make_docx(
            para(
                run("kept"),
                '<w:del w:id="2" w:author="a" w:date="2020-01-01T00:00:00Z">'
                f"<w:r><w:delText>{deleted}</w:delText></w:r></w:del>",
            )
        )

    delta = diff(snapshot(build("gone")), snapshot(build("also gone")))
    changed = {item.field for item in delta.paragraphs_changed[0].changes}
    assert changed == {"deleted_runs"}


# --------------------------------------------------------------------------
# Index space
# --------------------------------------------------------------------------


def test_paragraph_index_space_covers_table_cells_in_document_order():
    body = (
        para(run("before"))
        + "<w:tbl><w:tblGrid><w:gridCol w:w=\"100\"/><w:gridCol w:w=\"100\"/>"
        "</w:tblGrid><w:tr>"
        f"<w:tc>{para(run('cell-a'))}</w:tc><w:tc>{para(run('cell-b'))}</w:tc>"
        "</w:tr></w:tbl>"
        + para(run("after"))
    )
    snap = snapshot(make_docx(body))
    assert snap.text() == ("before", "cell-a", "cell-b", "after")


def test_textbox_paragraph_is_indexed_apart_and_not_counted_twice():
    inner = para(run("inner"))
    body = para(
        run("outer"),
        f"<w:r><w:pict><v:shape><v:textbox><w:txbxContent>{inner}"
        "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>",
    )
    snap = snapshot(make_docx(body))
    assert snap.text() == ("outer", "inner")


def test_stories_are_indexed_separately():
    header = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:hdr {NS}>{para(run('header text'))}</w:hdr>"
    )
    document = make_docx(
        para(run("body text")),
        extra_parts={"word/header1.xml": header},
        overrides=(
            ("/word/document.xml", DOCUMENT_CT),
            ("/word/header1.xml", HEADER_CT),
        ),
        document_rels=(("rId4", HEADER_REL, "header1.xml", None),),
    )
    snap = snapshot(document)
    assert snap.stories == ("document", "header1")
    assert snap.text() == ("body text",)
    assert snap.text("header1") == ("header text",)


# --------------------------------------------------------------------------
# Paragraph signature detail
# --------------------------------------------------------------------------


def test_run_properties_change_is_detected_although_the_text_is_identical():
    before = snapshot(make_docx(para(run("same", "<w:b/>"))))
    after = snapshot(make_docx(para(run("same", "<w:i/>"))))
    delta = diff(before, after)
    assert delta.paragraphs_changed[0].key == (MAIN_STORY, 0)
    assert [item.field for item in delta.paragraphs_changed[0].changes] == ["runs"]
    assert delta.parts_changed == ("word/document.xml",)


def test_markers_record_bookmark_names_on_both_ends():
    body = para(
        '<w:bookmarkStart w:id="7" w:name="intro"/>',
        '<w:commentRangeStart w:id="3"/>',
        run("text"),
        '<w:commentRangeEnd w:id="3"/>',
        '<w:r><w:commentReference w:id="3"/></w:r>',
        '<w:bookmarkEnd w:id="7"/>',
    )
    sig = snapshot(make_docx(body)).paragraphs[(MAIN_STORY, 0)]
    assert sig.markers == (
        ("bookmarkStart", "intro"),
        ("commentRangeStart", "3"),
        ("commentRangeEnd", "3"),
        ("commentReference", "3"),
        ("bookmarkEnd", "intro"),
    )


def test_dropping_a_marker_is_reported():
    with_bookmark = make_docx(
        para('<w:bookmarkStart w:id="7" w:name="intro"/>', run("text"),
             '<w:bookmarkEnd w:id="7"/>')
    )
    without = make_docx(para(run("text")))
    delta = diff(snapshot(with_bookmark), snapshot(without))
    fields = {item.field for item in delta.paragraphs_changed[0].changes}
    assert fields == {"markers"}
    assert [item.field for item in delta.counters_changed] == ["bookmarks"]


def test_style_and_numbering_are_recorded():
    body = para(run("item"), style="ListParagraph", num=("3", "1"))
    sig = snapshot(make_docx(body)).paragraphs[(MAIN_STORY, 0)]
    assert sig.style == "ListParagraph"
    assert sig.num_pr == ("3", "1")

    plain = snapshot(make_docx(para(run("item")))).paragraphs[(MAIN_STORY, 0)]
    assert plain.style is None
    assert plain.num_pr is None


# --------------------------------------------------------------------------
# Parts, content types, relationships
# --------------------------------------------------------------------------


def test_canonicalisation_ignores_attribute_order_and_serialisation_form():
    first = document_xml(
        '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r></w:p>'
        "<w:p></w:p>"
    )
    second = document_xml(
        '<w:p><w:r><w:fldChar w:dirty="true" w:fldCharType="begin"/></w:r></w:p>'
        "<w:p/>",
        declaration=False,
    )
    before = snapshot(make_docx(document=first))
    after = snapshot(make_docx(document=second))
    assert before.parts["word/document.xml"].kind == "xml"
    assert (
        before.parts["word/document.xml"].digest
        == after.parts["word/document.xml"].digest
    )
    assert diff(before, after).is_empty()


def test_binary_parts_are_compared_byte_for_byte():
    def build(image: bytes) -> bytes:
        return make_docx(
            para(run("x")),
            extra_parts={"word/media/image1.png": image},
            defaults=DEFAULT_DEFAULTS + (("png", "image/png"),),
            document_rels=(("rId5", IMAGE_REL, "media/image1.png", None),),
        )

    before = snapshot(build(b"\x89PNG\r\n\x1a\n" + b"\x00" * 8))
    after = snapshot(build(b"\x89PNG\r\n\x1a\n" + b"\x00" * 7 + b"\x01"))
    assert before.parts["word/media/image1.png"].kind == "binary"
    assert diff(before, after).parts_changed == ("word/media/image1.png",)


def test_content_types_are_compared_as_a_set():
    ordered = content_types_xml(
        DEFAULT_DEFAULTS,
        (("/word/document.xml", DOCUMENT_CT), ("/word/header1.xml", HEADER_CT)),
    )
    shuffled = content_types_xml(
        tuple(reversed(DEFAULT_DEFAULTS)),
        (("/word/header1.xml", HEADER_CT), ("/word/document.xml", DOCUMENT_CT)),
    )
    header = f'<?xml version="1.0"?><w:hdr {NS}>{para(run("h"))}</w:hdr>'
    kwargs = {
        "extra_parts": {"word/header1.xml": header},
        "document_rels": (("rId4", HEADER_REL, "header1.xml", None),),
    }
    before = snapshot(make_docx(para(run("x")), content_types=ordered, **kwargs))
    after = snapshot(make_docx(para(run("x")), content_types=shuffled, **kwargs))
    assert ("word/document.xml", DOCUMENT_CT) in before.content_types
    assert ("*.rels", PKG + "relationships+xml") in before.content_types
    assert diff(before, after).is_empty()


def test_content_type_loss_is_reported():
    before = snapshot(make_docx(para(run("x"))))
    after = snapshot(
        make_docx(para(run("x")), overrides=(), defaults=DEFAULT_DEFAULTS)
    )
    delta = diff(before, after)
    assert delta.content_types_removed == (("word/document.xml", DOCUMENT_CT),)
    # No content type means no story: the paragraph disappears from the index.
    assert delta.paragraphs_removed == ((MAIN_STORY, 0),)


def test_relationships_are_compared_as_type_and_target_pairs():
    forward = (
        ("rId1", HEADER_REL, "header1.xml", None),
        ("rId2", COMMENTS_REL, "comments.xml", None),
    )
    renumbered = (
        ("rId9", COMMENTS_REL, "comments.xml", None),
        ("rId8", HEADER_REL, "header1.xml", None),
    )
    before = snapshot(make_docx(para(run("x")), document_rels=forward))
    after = snapshot(make_docx(para(run("x")), document_rels=renumbered))
    delta = diff(before, after)
    assert before.relationships["word/document.xml"] == {
        (HEADER_REL, "header1.xml"),
        (COMMENTS_REL, "comments.xml"),
    }
    assert delta.relationships_added == ()
    assert delta.relationships_removed == ()
    # The ids themselves still live in the canonical part digest.
    assert delta.parts_changed == ("word/_rels/document.xml.rels",)


def test_a_lost_relationship_is_reported():
    before = snapshot(
        make_docx(
            para(run("x")),
            document_rels=(("rId1", HEADER_REL, "header1.xml", None),),
        )
    )
    after = snapshot(make_docx(para(run("x"))))
    assert diff(before, after).relationships_removed == (
        ("word/document.xml", HEADER_REL, "header1.xml"),
    )


# --------------------------------------------------------------------------
# Counters
# --------------------------------------------------------------------------


def test_counters_cover_comments_revisions_bookmarks_fields_and_notes():
    comments = (
        f'<?xml version="1.0"?><w:comments {NS}>'
        f'<w:comment w:id="3" w:author="a">{para(run("note"))}</w:comment>'
        "</w:comments>"
    )
    body = (
        para(
            '<w:bookmarkStart w:id="7" w:name="intro"/>',
            '<w:commentRangeStart w:id="3"/>',
            f'<w:hyperlink r:id="rId9">{run("link")}</w:hyperlink>',
            '<w:commentRangeEnd w:id="3"/>',
            '<w:r><w:commentReference w:id="3"/></w:r>',
            '<w:bookmarkEnd w:id="7"/>',
        )
        + para(
            '<w:ins w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z">'
            + run("added")
            + "</w:ins>",
            '<w:del w:id="2" w:author="a" w:date="2020-01-01T00:00:00Z">'
            "<w:r><w:delText>removed</w:delText></w:r></w:del>",
            '<w:fldSimple w:instr=" PAGE "/>',
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText> TOC </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
            '<w:r><w:footnoteReference w:id="2"/></w:r>',
        )
    )
    counters = snapshot(
        make_docx(
            body,
            extra_parts={"word/comments.xml": comments},
            overrides=(
                ("/word/document.xml", DOCUMENT_CT),
                ("/word/comments.xml", COMMENTS_CT),
            ),
            document_rels=(("rId3", COMMENTS_REL, "comments.xml", None),),
        )
    ).counters
    assert counters["bookmarks"] == 1
    assert counters["comments"] == 1
    assert counters["comment_references"] == 1
    assert counters["hyperlinks"] == 1
    assert counters["revisions"] == 2
    assert counters["fields"] == 2
    assert counters["footnote_references"] == 1


# --------------------------------------------------------------------------
# Diff
# --------------------------------------------------------------------------


def test_diff_reports_added_removed_and_changed_paragraphs():
    before = snapshot(make_docx(para(run("one")) + para(run("two"))))
    after = snapshot(make_docx(para(run("ONE")) + para(run("two")) + para(run("three"))))
    delta = diff(before, after)
    assert delta.paragraphs_added == ((MAIN_STORY, 2),)
    assert delta.paragraphs_removed == ()
    assert [change.key for change in delta.paragraphs_changed] == [(MAIN_STORY, 0)]
    assert {item.field for item in delta.paragraphs_changed[0].changes} == {
        "text",
        "runs",
    }
    assert not delta.is_empty()
    assert "document[0].text" in delta.describe()
    assert "paragraphs_added: document[2]" in delta.describe()


# --------------------------------------------------------------------------
# assert_unchanged_except
# --------------------------------------------------------------------------


def two_paragraph_document(first: str, second: str) -> bytes:
    return make_docx(para(run(first)) + para(run(second)))


def test_assert_unchanged_except_accepts_the_declared_paragraph():
    before = snapshot(two_paragraph_document("one", "two"))
    after = snapshot(two_paragraph_document("ONE", "two"))
    assert_unchanged_except(before, after, paragraphs={0})
    assert_unchanged_except(before, after, paragraphs={(MAIN_STORY, 0)})


def test_assert_unchanged_except_rejects_an_undeclared_paragraph():
    before = snapshot(two_paragraph_document("one", "two"))
    after = snapshot(two_paragraph_document("ONE", "TWO"))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={0})
    assert "document[1].text" in str(caught.value)


def test_assert_unchanged_except_rejects_an_unallowed_part():
    before = snapshot(two_paragraph_document("one", "two"))
    after = snapshot(
        make_docx(
            para(run("one")) + para(run("two")),
            document_rels=(("rId4", HEADER_REL, "header1.xml", None),),
        )
    )
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after)
    message = str(caught.value)
    assert "part changed: word/_rels/document.xml.rels" in message
    assert "relationship added: word/document.xml" in message
    assert_unchanged_except(
        before, after, parts={"word/_rels/document.xml.rels", "word/document.xml"}
    )


def test_assert_unchanged_except_guards_the_counters():
    before = snapshot(make_docx(para(run("one"))))
    after = snapshot(
        make_docx(
            para(
                '<w:bookmarkStart w:id="7" w:name="intro"/>',
                run("one"),
                '<w:bookmarkEnd w:id="7"/>',
            )
        )
    )
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={0})
    assert "counter bookmarks: 0 -> 1" in str(caught.value)
    assert_unchanged_except(before, after, paragraphs={0}, counters={"bookmarks"})


def table_document(span: str, cell_text: str) -> bytes:
    body = (
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"100\"/><w:gridCol w:w=\"100\"/>"
        "</w:tblGrid><w:tr>"
        f'<w:tc><w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr>'
        f"{para(run(cell_text))}</w:tc>"
        f"<w:tc>{para(run('b'))}</w:tc>"
        "</w:tr></w:tbl>"
    )
    return make_docx(body)


def test_table_signature_records_shape_and_cell_text():
    snap = snapshot(table_document("1", "a"))
    table = snap.tables[(MAIN_STORY, 0)]
    assert table.rows == 1
    assert table.grid_cols == 2
    assert table.grid == (((1, ""), (1, "")),)
    assert table.cells == ("a", "b")


def test_table_structure_is_guarded_even_when_a_paragraph_is_allowed():
    before = snapshot(table_document("1", "a"))
    after = snapshot(table_document("2", "a"))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={0, 1})
    assert "table document[0].grid" in str(caught.value)


def test_table_cell_text_is_left_to_the_paragraph_check():
    before = snapshot(table_document("1", "a"))
    after = snapshot(table_document("1", "A"))
    delta = diff(before, after)
    assert [item.field for item in delta.tables_changed[0].changes] == ["cells"]
    assert_unchanged_except(before, after, paragraphs={0})
    with pytest.raises(AssertionError):
        assert_unchanged_except(before, after)


# --------------------------------------------------------------------------
# What a relaxed part still guards, element by element
# --------------------------------------------------------------------------

WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def drawing(name: str = "Picture 1") -> str:
    return (
        f'<w:drawing><wp:inline xmlns:wp="{WP}">'
        f'<wp:docPr id="1" name="{name}"/></wp:inline></w:drawing>'
    )


def inline_control(tag: str, text: str) -> str:
    return (
        f'<w:sdt><w:sdtPr><w:tag w:val="{tag}"/><w:id w:val="1"/></w:sdtPr>'
        f"<w:sdtContent>{run(text)}</w:sdtContent></w:sdt>"
    )


def block_control(tag: str, paragraph: str) -> str:
    return (
        f'<w:sdt><w:sdtPr><w:tag w:val="{tag}"/><w:id w:val="2"/></w:sdtPr>'
        f"<w:sdtContent>{paragraph}</w:sdtContent></w:sdt>"
    )


def test_run_children_record_non_textual_content():
    body = para(f"<w:r>{drawing()}</w:r>", run("caption"))
    sig = snapshot(make_docx(body)).paragraphs[(MAIN_STORY, 0)]
    assert [name for name, _ in sig.runs[0].children] == ["drawing"]
    assert sig.runs[1].children == ()


def test_losing_a_drawing_is_reported_although_the_text_is_identical():
    before = snapshot(
        make_docx(para(f"<w:r>{drawing()}</w:r>", run("caption")) + para(run("b")))
    )
    after = snapshot(make_docx(para("<w:r/>", run("caption")) + para(run("B"))))
    assert before.text() == ("caption", "b")
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={1})
    assert "document[0].runs" in str(caught.value)


def test_a_swapped_drawing_is_reported_although_the_run_shape_is_the_same():
    before = snapshot(make_docx(para(f"<w:r>{drawing('Picture 1')}</w:r>")))
    after = snapshot(make_docx(para(f"<w:r>{drawing('Other picture')}</w:r>")))
    delta = diff(before, after)
    assert [item.field for item in delta.paragraphs_changed[0].changes] == ["runs"]


def test_paragraph_properties_are_fingerprinted():
    def document(spacing: str) -> bytes:
        first = f'<w:p><w:pPr><w:spacing w:after="{spacing}"/></w:pPr>{run("a")}</w:p>'
        return make_docx(para(run("editable")) + first)

    before = snapshot(document("120"))
    after = snapshot(document("240"))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={0})
    assert "document[1].ppr" in str(caught.value)


def test_deleted_paragraph_mark_survives_a_relaxed_revision_counter():
    """``counters=`` relaxes a total; the paragraph signature keeps the detail."""
    mark = '<w:del w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z"/>'
    before = snapshot(
        make_docx(
            para(run("editable"))
            + f"<w:p><w:pPr><w:rPr>{mark}</w:rPr></w:pPr>{run('a')}</w:p>"
        )
    )
    after = snapshot(make_docx(para(run("EDITABLE")) + para(run("a"))))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(
            before, after, paragraphs={0}, counters={"revisions"}
        )
    assert "document[1].ppr" in str(caught.value)


def test_inline_content_control_properties_are_recorded():
    before = snapshot(
        make_docx(para(run("editable")) + para(inline_control("keep", "value")))
    )
    sig = before.paragraphs[(MAIN_STORY, 1)]
    assert len(sig.inline_controls) == 1
    assert 'w:val="keep"' in sig.inline_controls[0]

    after = snapshot(make_docx(para(run("EDITABLE")) + para(run("value"))))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={0})
    assert "document[1].inline_controls" in str(caught.value)


def test_block_content_control_is_visible_from_the_paragraph_it_wraps():
    wrapped = para(run("controlled"))
    before = snapshot(make_docx(block_control("keep", wrapped) + para(run("free"))))
    assert before.paragraphs[(MAIN_STORY, 0)].enclosing_controls != ()
    assert before.paragraphs[(MAIN_STORY, 1)].enclosing_controls == ()

    after = snapshot(make_docx(wrapped + para(run("FREE"))))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={1})
    assert "document[0].enclosing_controls" in str(caught.value)


def styled_table_document(tbl_pr: str, tc_pr: str) -> bytes:
    body = (
        f"<w:tbl>{tbl_pr}"
        '<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>'
        f"<w:tr><w:tc>{tc_pr}{para(run('cell'))}</w:tc></w:tr></w:tbl>"
    )
    return make_docx(para(run("editable")) + body)


def test_table_properties_are_guarded_even_when_a_paragraph_is_allowed():
    styled = '<w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>'
    shaded = '<w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr>'
    before = snapshot(styled_table_document(styled, shaded))
    assert 'w:val="Grid"' in before.tables[(MAIN_STORY, 0)].tbl_pr

    after = snapshot(styled_table_document("", ""))
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs={0, 1})
    message = str(caught.value)
    assert "table document[0].tbl_pr" in message
    assert "table document[0].cell_properties" in message


# --------------------------------------------------------------------------
# The same guarantee, measured on the shared fixtures
# --------------------------------------------------------------------------


@cache
def combined_docx() -> bytes:
    """The ``combined`` fixture, built once for this module."""
    return build("combined")


def rewrite_document(blob: bytes, mutate) -> bytes:
    """Apply `mutate` to the parsed ``word/document.xml`` and rezip the package."""
    parts: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(blob)) as archive:
        for info in archive.infolist():
            if not info.is_dir():
                parts[info.filename] = archive.read(info)
    root = etree.fromstring(parts["word/document.xml"])
    mutate(root)
    parts["word/document.xml"] = etree.tostring(
        root.getroottree(), xml_declaration=True, encoding="UTF-8", standalone=True
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            archive.writestr(name, parts[name])
    return buffer.getvalue()


def strip_elements(blob: bytes, localname: str) -> tuple[bytes, int]:
    """Delete every ``w:<localname>`` of the main part; return the bytes and count."""
    removed = 0

    def mutate(root) -> None:
        nonlocal removed
        for element in list(root.iter(f"{{{W}}}{localname}")):
            element.getparent().remove(element)
            removed += 1

    return rewrite_document(blob, mutate), removed


@pytest.mark.parametrize(
    ("localname", "expected_removals", "expected_field"),
    [
        ("drawing", 2, "runs"),  # inline images
        ("tblPr", 2, "tbl_pr"),  # table style, borders, width
        ("sdtPr", 3, "enclosing_controls"),  # content control identity
    ],
)
def test_deleting_a_carrier_of_meaning_is_reported_on_a_relaxed_part(
    localname: str, expected_removals: int, expected_field: str
):
    """Allowing one paragraph must not blind the assertion to the rest of the part."""
    original = combined_docx()
    mutated, removed = strip_elements(original, localname)
    assert removed == expected_removals

    before, after = snapshot(original), snapshot(mutated)
    assert "word/document.xml" in diff(before, after).parts_changed
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(before, after, paragraphs=[0])
    assert expected_field in str(caught.value)


def test_deleting_an_inline_content_control_is_reported_on_a_relaxed_part():
    mutated, _ = strip_elements(combined_docx(), "sdtPr")
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(
            snapshot(combined_docx()), snapshot(mutated), paragraphs=[0]
        )
    assert "inline_controls" in str(caught.value)


def test_paragraph_revisions_survive_a_relaxed_revision_counter_on_a_fixture():
    """The secondary finding: ``counters=`` must not hide a per-paragraph loss."""

    def mutate(root) -> None:
        for properties in list(root.iter(f"{{{W}}}pPr")):
            for change in list(properties.iter(f"{{{W}}}pPrChange")):
                change.getparent().remove(change)
            mark = properties.find(f"{{{W}}}rPr")
            if mark is not None:
                for deletion in list(mark.findall(f"{{{W}}}del")):
                    mark.remove(deletion)

    original = combined_docx()
    before = snapshot(original)
    after = snapshot(rewrite_document(original, mutate))
    assert before.counters["revisions"] > after.counters["revisions"]
    with pytest.raises(AssertionError) as caught:
        assert_unchanged_except(
            before, after, paragraphs=[0], counters={"revisions"}
        )
    assert ".ppr" in str(caught.value)
