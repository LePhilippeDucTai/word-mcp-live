"""Tests for :mod:`word_document_server.engine.audit`.

An audit is a list of accusations, so what is pinned here is mostly the other
half of each one: the document that must *not* be reported.  A checker that
fires on everything is worth no more than one that fires on nothing, and only
the negative cases say which of the two this one is.

Three properties hold across every check:

*A finding names a place a locator reaches.*
    Every locator reported here is fed back through
    :func:`~word_document_server.engine.locators.resolve` and checked against
    the text the finding quotes.  An address that is one off would send an agent
    to edit the wrong paragraph.

*A finding is plain data.*
    The report crosses the MCP boundary, so no lxml element and no live
    reference may leak into it.

*Auditing changes nothing.*
    Read only, checked on the bytes: the package that comes out of an audit is
    the package that went in.

The documents that carry a defect -- a dangling ``w:numId``, a bookmark with one
end, a comment nobody points at -- are built by editing the live parts of an
opened fixture rather than by adding a fixture of their own: they are properties
of this reader, not documents any other test needs.
"""

from __future__ import annotations

import dataclasses
import json
import re
from collections import Counter
from typing import Any

from lxml import etree

from tests.fixtures.builders import build
from tests.support.snapshot import diff, snapshot
from word_document_server.engine.audit import (
    DIRECT_FORMATTING_CROWD,
    FINDING_KINDS,
    SEVERITIES,
    Finding,
    audit,
)
from word_document_server.engine.locators import resolve
from word_document_server.engine.numbering import numbering_root
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.revisions import list_revisions
from word_document_server.engine.styles import STYLES_PARTNAME
from word_document_server.engine.xmlns import NAMESPACES, qn
from word_document_server.tools.v2.audit import TOOLS, doc_audit

W_BODY = qn("w:body")
W_SECT_PR = qn("w:sectPr")
W_VAL = qn("w:val")


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------


def _fragment(xml: str) -> etree._Element:
    """Parse a prefixed fragment, declaring the namespaces it uses.

    The declarations go right after the element *name*, so that a self-closing
    tag does not get them appended after its slash.
    """
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in NAMESPACES.items())
    stripped = xml.strip()
    match = re.match(r"<[A-Za-z0-9_:.-]+", stripped)
    assert match is not None, f"not a tag: {stripped[:40]!r}"
    return etree.fromstring(f"{stripped[: match.end()]} {declarations}{stripped[match.end() :]}")


def _pkg(name: str) -> DocxPackage:
    return DocxPackage.open(build(name))


def _append_body(pkg: DocxPackage, *fragments: str) -> DocxPackage:
    """Append body-level fragments to the main story, before its ``w:sectPr``."""
    body = pkg.document.find(W_BODY)
    assert body is not None
    section = body.find(W_SECT_PR)
    for fragment in fragments:
        element = _fragment(fragment)
        if section is None:
            body.append(element)
        else:
            section.addprevious(element)
    return pkg


def _append_styles(pkg: DocxPackage, *fragments: str) -> DocxPackage:
    """Append ``w:style`` definitions to the style sheet."""
    root = pkg.root_of(pkg.part(STYLES_PARTNAME))
    for fragment in fragments:
        root.append(_fragment(fragment))
    return pkg


def _of_kind(findings: list[Finding], kind: str) -> list[Finding]:
    assert kind in FINDING_KINDS, f"undeclared kind {kind!r}"
    return [finding for finding in findings if finding.kind == kind]


def _one(findings: list[Finding], kind: str) -> Finding:
    matching = _of_kind(findings, kind)
    assert len(matching) == 1, f"expected one {kind!r} finding, got {len(matching)}"
    return matching[0]


def _kinds(findings: list[Finding]) -> dict[str, int]:
    return dict(sorted(Counter(finding.kind for finding in findings).items()))


# --------------------------------------------------------------------------------------
# Shape of the report
# --------------------------------------------------------------------------------------


def test_every_finding_uses_the_declared_vocabulary() -> None:
    for finding in audit(_pkg("combined")):
        assert finding.kind in FINDING_KINDS
        assert finding.severity in SEVERITIES
        assert finding.message
        assert isinstance(finding.data, dict)


def test_a_report_is_plain_json_ready_data() -> None:
    # No lxml element may leak out: the report crosses the MCP boundary, and a
    # live element there would read as a handle the caller could keep.
    report = [dataclasses.asdict(finding) for finding in audit(_pkg("combined"))]
    assert json.loads(json.dumps(report)) == report


def test_auditing_never_touches_the_package() -> None:
    source = build("combined")
    pkg = DocxPackage.open(source)
    audit(pkg)
    assert diff(snapshot(source), snapshot(pkg.to_bytes())).is_empty()


def test_two_audits_of_the_same_bytes_agree() -> None:
    source = build("combined")
    first = audit(DocxPackage.open(source))
    second = audit(DocxPackage.open(source))
    assert [dataclasses.asdict(f) for f in first] == [dataclasses.asdict(g) for g in second]


def test_every_locator_a_finding_carries_resolves() -> None:
    pkg = _pkg("combined")
    checked = 0
    for finding in audit(pkg):
        locator = finding.locator
        if locator is None:
            continue
        target = resolve(pkg, locator)
        assert target.story == locator["story"]
        if "table" not in locator:
            assert target.index == locator["paragraph"]
        checked += 1
    assert checked, "the combined fixture should produce located findings"


def test_a_document_with_nothing_wrong_reports_nothing() -> None:
    # The baseline fixture is a well-formed document, and an auditor that still
    # finds something in it is an auditor nobody will read.
    assert audit(_pkg("simple")) == []


# --------------------------------------------------------------------------------------
# heading_like_paragraph
# --------------------------------------------------------------------------------------


_BOLD_LINE = """
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Chapter one</w:t></w:r></w:p>
"""


def test_a_short_bold_line_reads_as_a_heading() -> None:
    finding = _one(audit(_append_body(_pkg("simple"), _BOLD_LINE)), "heading_like_paragraph")
    assert finding.severity == "warning"
    assert finding.data["text"] == "Chapter one"
    assert finding.data["cues"] == ["bold"]
    assert finding.data["style"] is None


def test_a_heading_style_makes_a_bold_line_a_heading() -> None:
    styled = """
    <w:p>
      <w:pPr><w:pStyle w:val="Heading3"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Chapter one</w:t></w:r>
    </w:p>
    """
    assert _of_kind(audit(_append_body(_pkg("simple"), styled)), "heading_like_paragraph") == []


def test_a_bold_sentence_is_not_a_heading() -> None:
    sentence = """
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>This whole sentence is bold.</w:t></w:r></w:p>
    """
    assert _of_kind(audit(_append_body(_pkg("simple"), sentence)), "heading_like_paragraph") == []


def test_a_line_with_one_emphasised_word_is_not_a_heading() -> None:
    # The cue has to hold for the whole line: a sentence with a bold word in it
    # is a sentence, and reporting it would drown the real findings.
    mixed = """
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Chapter </w:t></w:r>
      <w:r><w:t>one</w:t></w:r>
    </w:p>
    """
    assert _of_kind(audit(_append_body(_pkg("simple"), mixed)), "heading_like_paragraph") == []


def test_a_directly_enlarged_line_reads_as_a_heading() -> None:
    enlarged = """
    <w:p><w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>Chapter one</w:t></w:r></w:p>
    """
    finding = _one(audit(_append_body(_pkg("simple"), enlarged)), "heading_like_paragraph")
    assert finding.data["cues"] == ["larger"]


def test_a_numbered_bold_line_is_a_list_item_not_a_heading() -> None:
    item = """
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="900"/></w:numPr></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Bold list item</w:t></w:r>
    </w:p>
    """
    findings = audit(_append_body(_pkg("complex_numbering"), item))
    assert _of_kind(findings, "heading_like_paragraph") == []


# --------------------------------------------------------------------------------------
# direct_formatting_overrides_style
# --------------------------------------------------------------------------------------


def test_direct_formatting_that_contradicts_the_style_is_named() -> None:
    # The fixture's last paragraph is styled FixtureBody, which justifies, and
    # centers itself directly.
    finding = _one(
        audit(_pkg("paragraph_styles")), "direct_formatting_overrides_style"
    )
    assert finding.locator == {"story": "document", "paragraph": 5}
    assert finding.data["properties"] == ["alignment"]
    assert finding.data["count"] == 1
    assert finding.data["style"] == "FixtureBody"
    assert finding.severity == "info"


def test_a_property_the_style_says_nothing_about_is_not_an_override() -> None:
    # Bold over a style that never mentions bold is the only thing saying what
    # the text looks like; there is nothing to reconcile.
    findings = audit(_append_body(_pkg("simple"), _BOLD_LINE))
    assert _of_kind(findings, "direct_formatting_overrides_style") == []


def test_a_crowd_of_overrides_is_a_warning() -> None:
    crowded = """
    <w:p>
      <w:pPr><w:pStyle w:val="FixtureNote"/><w:ind w:left="0"/></w:pPr>
      <w:r>
        <w:rPr><w:i w:val="0"/><w:color w:val="000000"/><w:sz w:val="28"/></w:rPr>
        <w:t>Nothing of the style is left</w:t>
      </w:r>
    </w:p>
    """
    findings = _of_kind(
        audit(_append_body(_pkg("paragraph_styles"), crowded)),
        "direct_formatting_overrides_style",
    )
    crowd = [finding for finding in findings if finding.data["count"] >= DIRECT_FORMATTING_CROWD]
    assert len(crowd) == 1
    assert crowd[0].severity == "warning"
    assert crowd[0].data["paragraph_properties"] == ["indent.left"]
    assert set(crowd[0].data["run_properties"]) >= {"italic", "color", "size_pt"}


# --------------------------------------------------------------------------------------
# mixed_fonts_in_paragraph
# --------------------------------------------------------------------------------------


def test_a_paragraph_whose_runs_end_in_two_fonts_is_reported() -> None:
    # The fixture's first paragraph carries a FixtureCode run, which pins
    # Consolas while the rest follows the theme.
    finding = _one(audit(_pkg("character_styles")), "mixed_fonts_in_paragraph")
    assert "Consolas" in finding.data["fonts"]
    assert len(finding.data["fonts"]) == 2


def test_runs_that_differ_by_anything_but_the_font_are_not_mixed() -> None:
    assert _of_kind(audit(_pkg("mixed_runs")), "mixed_fonts_in_paragraph") == []


# --------------------------------------------------------------------------------------
# consecutive_empty_paragraphs
# --------------------------------------------------------------------------------------


def test_a_run_of_empty_paragraphs_is_reported_once_with_its_span() -> None:
    pkg = _append_body(_pkg("simple"), "<w:p/>", "<w:p/>", "<w:p/>")
    finding = _one(audit(pkg), "consecutive_empty_paragraphs")
    assert finding.data["count"] == 3
    assert finding.data["end"] - finding.data["start"] == 2
    assert finding.locator == {"story": "document", "paragraph": finding.data["start"]}


def test_a_single_empty_paragraph_is_a_habit_not_a_finding() -> None:
    pkg = _append_body(_pkg("simple"), "<w:p/>", "<w:p><w:r><w:t>Text.</w:t></w:r></w:p>", "<w:p/>")
    assert _of_kind(audit(pkg), "consecutive_empty_paragraphs") == []


# --------------------------------------------------------------------------------------
# dangling numbering
# --------------------------------------------------------------------------------------


_NUMBERED = """
<w:p>
  <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>
  <w:r><w:t>Numbered with {num_id}</w:t></w:r>
</w:p>
"""


def test_a_paragraph_numbered_with_an_undefined_list_is_an_error() -> None:
    pkg = _append_body(_pkg("complex_numbering"), _NUMBERED.format(num_id=999))
    finding = _one(audit(pkg), "dangling_num_id")
    assert finding.severity == "error"
    assert finding.data["num_id"] == 999
    assert finding.data["level"] == 0
    assert resolve(pkg, finding.locator).text == "Numbered with 999"


def test_num_id_zero_removes_numbering_and_is_not_a_reference() -> None:
    pkg = _append_body(_pkg("complex_numbering"), _NUMBERED.format(num_id=0))
    assert _of_kind(audit(pkg), "dangling_num_id") == []


def test_the_numbering_fixture_defines_everything_it_uses() -> None:
    findings = audit(_pkg("complex_numbering"))
    assert _of_kind(findings, "dangling_num_id") == []
    assert _of_kind(findings, "dangling_abstract_num_id") == []


def test_a_list_pointing_at_an_undefined_abstract_definition_is_an_error() -> None:
    pkg = _pkg("complex_numbering")
    root = numbering_root(pkg)
    assert root is not None
    root.append(_fragment('<w:num w:numId="950"><w:abstractNumId w:val="777"/></w:num>'))
    finding = _one(audit(pkg), "dangling_abstract_num_id")
    assert finding.locator is None
    assert finding.data == {"num_id": 950, "abstract_num_id": 777}


# --------------------------------------------------------------------------------------
# bookmark and comment ranges
# --------------------------------------------------------------------------------------


def test_a_bookmark_that_never_closes_is_an_error() -> None:
    orphan = """
    <w:p>
      <w:bookmarkStart w:id="800" w:name="NeverClosed"/>
      <w:r><w:t>An orphaned bookmark start.</w:t></w:r>
    </w:p>
    """
    finding = _one(audit(_append_body(_pkg("bookmarks"), orphan)), "bookmark_without_end")
    assert finding.severity == "error"
    assert finding.data["name"] == "NeverClosed"
    assert finding.data["id"] == "800"


def test_a_bookmark_end_without_a_start_is_an_error() -> None:
    orphan = '<w:p><w:bookmarkEnd w:id="801"/><w:r><w:t>Closing nothing.</w:t></w:r></w:p>'
    finding = _one(audit(_append_body(_pkg("bookmarks"), orphan)), "bookmark_without_start")
    assert finding.data["id"] == "801"


def test_a_bookmark_spanning_two_paragraphs_is_whole() -> None:
    # The fixture opens a bookmark in one paragraph and closes it in the next:
    # a pairing check that walked paragraphs instead of the story would call it
    # broken.
    findings = audit(_pkg("bookmarks"))
    assert _of_kind(findings, "bookmark_without_end") == []
    assert _of_kind(findings, "bookmark_without_start") == []


def test_a_comment_range_with_one_end_is_an_error() -> None:
    half = """
    <w:p>
      <w:commentRangeStart w:id="900"/>
      <w:r><w:t>Commented on, never closed.</w:t></w:r>
    </w:p>
    """
    pkg = _append_body(_pkg("comments"), half)
    finding = _one(audit(pkg), "comment_range_without_end")
    assert finding.severity == "error"
    assert finding.data["comment_id"] == "900"
    assert resolve(pkg, finding.locator).text == "Commented on, never closed."


def test_a_comment_range_closing_without_opening_is_an_error() -> None:
    half = '<w:p><w:commentRangeEnd w:id="901"/><w:r><w:t>Closing nothing.</w:t></w:r></w:p>'
    finding = _one(audit(_append_body(_pkg("comments"), half)), "comment_range_without_start")
    assert finding.data["comment_id"] == "901"


def test_the_comments_fixture_has_whole_ranges_and_anchored_comments() -> None:
    findings = audit(_pkg("comments"))
    for kind in (
        "comment_range_without_end",
        "comment_range_without_start",
        "comment_without_anchor",
    ):
        assert _of_kind(findings, kind) == []


def test_a_comment_nothing_points_at_is_reported() -> None:
    pkg = _pkg("comments")
    root = pkg.root_of(pkg.part("/word/comments.xml"))
    root.append(
        _fragment(
            """
            <w:comment w:id="99" w:author="Nobody" w:initials="N" w:date="2024-01-01T00:00:00Z">
              <w:p><w:r><w:t>Written and anchored nowhere.</w:t></w:r></w:p>
            </w:comment>
            """
        )
    )
    finding = _one(audit(pkg), "comment_without_anchor")
    assert finding.severity == "warning"
    assert finding.data["comment_id"] == "99"
    assert finding.data["author"] == "Nobody"
    assert finding.data["text"] == "Written and anchored nowhere."


# --------------------------------------------------------------------------------------
# the style sheet
# --------------------------------------------------------------------------------------


_UNUSED_STYLE = """
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FixtureUnused">
  <w:name w:val="Fixture Unused"/>
  <w:basedOn w:val="Normal"/>
  <w:rPr><w:b/></w:rPr>
</w:style>
"""


def test_a_custom_style_nobody_applies_is_reported() -> None:
    finding = _one(audit(_append_styles(_pkg("simple"), _UNUSED_STYLE)), "unused_custom_style")
    assert finding.locator is None
    assert finding.data == {
        "style_id": "FixtureUnused",
        "name": "Fixture Unused",
        "family": "paragraph",
    }


def test_applying_a_custom_style_makes_it_used() -> None:
    pkg = _append_styles(_pkg("simple"), _UNUSED_STYLE)
    _append_body(
        pkg,
        '<w:p><w:pPr><w:pStyle w:val="FixtureUnused"/></w:pPr><w:r><w:t>Used.</w:t></w:r></w:p>',
    )
    assert _of_kind(audit(pkg), "unused_custom_style") == []


def test_a_custom_style_another_style_inherits_from_is_used() -> None:
    heir = """
    <w:style w:type="paragraph" w:customStyle="1" w:styleId="FixtureHeir">
      <w:name w:val="Fixture Heir"/>
      <w:basedOn w:val="FixtureUnused"/>
    </w:style>
    """
    pkg = _append_styles(_pkg("simple"), _UNUSED_STYLE, heir)
    unused = {finding.data["style_id"] for finding in _of_kind(audit(pkg), "unused_custom_style")}
    assert unused == {"FixtureHeir"}


def test_an_unused_builtin_style_is_not_reported() -> None:
    # Word's own style sheet ships dozens of definitions a document never uses;
    # listing them would bury the one somebody added on purpose.
    assert _of_kind(audit(_pkg("simple")), "unused_custom_style") == []


def test_a_style_applied_only_in_a_comment_is_not_unused() -> None:
    # A comment is not a story (DocxPackage.stories() never reports
    # word/comments.xml), but a style applied only inside one is still applied.
    used_only_in_comment = """
    <w:style w:type="paragraph" w:customStyle="1" w:styleId="FixtureCommentStyle">
      <w:name w:val="Fixture Comment Style"/>
    </w:style>
    """
    pkg = _append_styles(_pkg("comments"), _UNUSED_STYLE, used_only_in_comment)
    comments_root = pkg.root_of(pkg.part("/word/comments.xml"))
    comment = comments_root.find(qn("w:comment"))
    assert comment is not None
    comment.append(
        _fragment(
            '<w:p><w:pPr><w:pStyle w:val="FixtureCommentStyle"/></w:pPr>'
            "<w:r><w:t>Styled inside a comment.</w:t></w:r></w:p>"
        )
    )
    unused = {finding.data["style_id"] for finding in _of_kind(audit(pkg), "unused_custom_style")}
    assert "FixtureUnused" in unused
    assert "FixtureCommentStyle" not in unused


def test_style_names_that_differ_only_by_case_and_punctuation_are_reported() -> None:
    twin = """
    <w:style w:type="paragraph" w:customStyle="1" w:styleId="FixtureTwin">
      <w:name w:val="fixture-body"/>
    </w:style>
    """
    other = """
    <w:style w:type="paragraph" w:customStyle="1" w:styleId="FixtureTriplet">
      <w:name w:val="Fixture  BODY"/>
    </w:style>
    """
    findings = audit(_append_styles(_pkg("paragraph_styles"), twin, other))
    finding = _one(findings, "similar_style_names")
    assert finding.severity == "warning"
    assert finding.data["normalized"] == "fixturebody"
    assert [entry["style_id"] for entry in finding.data["styles"]] == [
        "FixtureBody",
        "FixtureTwin",
        "FixtureTriplet",
    ]


def test_names_a_reader_tells_apart_are_not_similar() -> None:
    assert _of_kind(audit(_pkg("style_inheritance")), "similar_style_names") == []


# --------------------------------------------------------------------------------------
# inventories
# --------------------------------------------------------------------------------------


def test_revisions_are_grouped_by_author_and_kind() -> None:
    pkg = _pkg("tracked_changes")
    expected: dict[str, Counter[str]] = {}
    for revision in list_revisions(pkg):
        expected.setdefault(revision.author, Counter())[revision.kind] += 1
    findings = _of_kind(audit(pkg), "revisions_by_author")
    assert [finding.data["author"] for finding in findings] == sorted(expected)
    for finding in findings:
        author = finding.data["author"]
        assert finding.data["count"] == sum(expected[author].values())
        assert finding.data["kinds"] == dict(sorted(expected[author].items()))


def test_fields_are_grouped_by_type_with_their_places() -> None:
    pkg = _pkg("fields")
    findings = _of_kind(audit(pkg), "field_present")
    by_field = {finding.data["field"]: finding for finding in findings}
    assert set(by_field) == {"PAGE", "TOC", "IF", "REF"}
    toc = by_field["TOC"]
    assert toc.data["count"] == 1
    assert toc.data["instructions"] == ['TOC \\o "1-3" \\h \\z \\u']
    assert len(toc.data["locations"]) == 1
    for locator in toc.data["locations"]:
        assert resolve(pkg, locator).index == locator["paragraph"]


def test_a_stale_prone_field_is_a_warning_until_the_document_updates_them() -> None:
    assert _of_kind(audit(_pkg("fields")), "field_present")
    toc = [
        finding
        for finding in _of_kind(audit(_pkg("fields")), "field_present")
        if finding.data["field"] == "TOC"
    ]
    assert [finding.severity for finding in toc] == ["warning"]
    assert toc[0].data["update_fields"] is False

    pkg = _pkg("fields")
    settings = pkg.root_of(pkg.part("/word/settings.xml"))
    settings.append(_fragment('<w:updateFields w:val="true"/>'))
    refreshed = [
        finding
        for finding in _of_kind(audit(pkg), "field_present")
        if finding.data["field"] == "TOC"
    ]
    assert [finding.severity for finding in refreshed] == ["info"]
    assert refreshed[0].data["update_fields"] is True


# --------------------------------------------------------------------------------------
# The tool
# --------------------------------------------------------------------------------------


def test_the_tool_module_exports_its_specs() -> None:
    # D-024: a module of ``tools/v2`` with no ``TOOLS`` list is discovered as
    # "not a tool module", in silence.
    assert [spec.fn for spec in TOOLS] == [doc_audit]
    assert all(spec.annotations is not None and spec.annotations.read_only_hint for spec in TOOLS)


def test_doc_audit_reports_the_findings_and_counts_them(tmp_path: Any) -> None:
    path = tmp_path / "audited.docx"
    path.write_bytes(build("combined"))
    report = doc_audit(str(path))["audit"]

    findings = audit(DocxPackage.open(build("combined")))
    assert report["findings"] == [dataclasses.asdict(finding) for finding in findings]
    assert report["counts"]["total"] == len(findings)
    assert report["counts"]["by_kind"] == _kinds(findings)
    assert sum(report["counts"]["by_severity"].values()) == len(findings)
    assert json.loads(json.dumps(report)) == report
