"""Tests for the canonical snapshot and the structural comparison.

Every document here is built in memory from raw OOXML: no binary is committed,
and neither the module under test nor these tests import python-docx.
"""

from __future__ import annotations

import io
import zipfile

import pytest

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
