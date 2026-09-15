"""Tests for :mod:`word_document_server.engine.compare` and the ``doc_compare`` tool.

Two things are pinned here, deliberately kept apart:

*the move itself* -- ``tests.support.snapshot`` is a pure re-export of
:mod:`word_document_server.engine.compare` (J06-P2).  The exhaustive coverage of
``snapshot``/``diff``/``assert_unchanged_except`` stays in
``tests/support/test_snapshot.py``, unedited, importing through the old path;
this module only checks that the shim really is the same objects, not a copy.

*the semantic layer* -- ``doc_compare`` (:mod:`word_document_server.tools.v2.compare`)
translates the engine's *snapshot-space* paragraph keys (every ``w:p`` of a
story, table cells and text boxes included) into the *V2* locator vocabulary
(``None`` inside a table cell or a text box, D-016) before reporting them.
That translation, not the structural diff underneath it, is what these tests
exercise: the diff engine itself is ``tests/support/test_snapshot.py``'s job.

The ``simple`` fixture (no table, no text box) is used whenever the snapshot
space and the V2 space coincide, so a paragraph can be named by a bare
integer and checked against both spaces at once; ``tables`` is used for the
one case where they diverge.
"""

from __future__ import annotations

import asyncio
import io
import json
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from lxml import etree

import tests.support.snapshot as snapshot_shim
import word_document_server.engine.compare as engine_compare
from tests.fixtures.builders import build
from word_document_server.tools.v2.compare import doc_compare
from word_document_server.tools.v2.registry import discover_tool_specs, v2_tools

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


_BODY = _w("body")
_P = _w("p")
_T = _w("t")
_PPR = _w("pPr")
_SECT_PR = _w("sectPr")
_TC = _w("tc")


def call(name: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Call a V2 tool the way an MCP client reaches it, and return its report."""
    return asyncio.run(v2_tools()[name](*args, **kwargs))


def _paths(tmp_path: Path, blob_a: bytes, blob_b: bytes) -> tuple[str, str]:
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    a.write_bytes(blob_a)
    b.write_bytes(blob_b)
    return str(a), str(b)


def _mutate_document(blob: bytes, mutate: Callable[[etree._Element], None]) -> bytes:
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


def _top_level_paragraphs(root: etree._Element) -> list[etree._Element]:
    body = root.find(_BODY)
    return [child for child in body if child.tag == _P]


# --------------------------------------------------------------------------
# The move: tests.support.snapshot is the engine module, not a copy of it
# --------------------------------------------------------------------------


def test_the_shim_reexports_the_engine_modules_own_objects():
    for name in engine_compare.__all__:
        assert getattr(snapshot_shim, name) is getattr(engine_compare, name)
    assert set(snapshot_shim.__all__) == set(engine_compare.__all__)


def test_snapshot_and_diff_work_directly_from_the_engine_module():
    document = build("simple")
    assert engine_compare.diff(
        engine_compare.snapshot(document), engine_compare.snapshot(document)
    ).is_empty()


# --------------------------------------------------------------------------
# doc_compare: registration
# --------------------------------------------------------------------------


def test_doc_compare_is_discovered_and_read_only():
    specs = {spec.name: spec for spec in discover_tool_specs()}
    assert "doc_compare" in specs
    spec = specs["doc_compare"]
    assert spec.annotations is not None
    assert spec.annotations.read_only_hint is True
    assert "v2" in spec.tags and "read" in spec.tags


# --------------------------------------------------------------------------
# doc_compare: identical documents
# --------------------------------------------------------------------------


def test_identical_documents_report_no_differences(tmp_path):
    blob = build("simple")
    a, b = _paths(tmp_path, blob, blob)

    report = doc_compare(a, b)

    assert report["identical"] is True
    assert report["parts"] == {"added": [], "removed": [], "changed": []}
    assert report["paragraphs"] == {"added": [], "removed": [], "changed": []}
    assert report["tables"] == {"added": [], "removed": [], "changed": []}
    assert report["relationships"] == {"added": [], "removed": []}
    assert report["sections_changed"] == []
    assert report["counters_changed"] == []


def test_the_report_is_plain_json_ready_data(tmp_path):
    a, b = _paths(tmp_path, build("simple"), build("tables"))
    report = doc_compare(a, b)
    assert json.loads(json.dumps(report)) == report


def test_doc_compare_is_called_through_the_registry_with_the_usual_envelope(tmp_path):
    blob = build("simple")
    a, b = _paths(tmp_path, blob, blob)

    report = call("doc_compare", a, b)

    assert report["status"] == "ok"
    assert report["dry_run"] is False
    assert report["changes"] == []
    assert report["warnings"] == []
    assert report["identical"] is True


# --------------------------------------------------------------------------
# doc_compare: a changed paragraph, in the space where V2 and snapshot coincide
# --------------------------------------------------------------------------


def test_changed_text_is_reported_with_the_v2_paragraph_index(tmp_path):
    """``simple`` has no table and no text box: V2 index == snapshot index."""
    before_blob = build("simple")

    def mutate(root: etree._Element) -> None:
        paragraph = _top_level_paragraphs(root)[1]  # "First paragraph..."
        next(paragraph.iter(_T)).text = "Rewritten paragraph text."

    after_blob = _mutate_document(before_blob, mutate)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)

    assert report["identical"] is False
    assert "word/document.xml" in report["parts"]["changed"]
    assert report["paragraphs"]["added"] == []
    assert report["paragraphs"]["removed"] == []
    assert len(report["paragraphs"]["changed"]) == 1
    entry = report["paragraphs"]["changed"][0]
    assert entry["story"] == "document"
    assert entry["paragraph"] == 1
    assert "text" in entry["fields"]
    assert "runs" in entry["fields"]
    assert entry["text_before"] == "First paragraph of the simple fixture."
    assert entry["text_after"] == "Rewritten paragraph text."


def test_changed_style_is_reported_among_the_changed_fields(tmp_path):
    before_blob = build("simple")

    def mutate(root: etree._Element) -> None:
        paragraph = _top_level_paragraphs(root)[2]  # "Second paragraph..."
        properties = etree.SubElement(paragraph, _PPR)
        style = etree.SubElement(properties, _w("pStyle"))
        style.set(_w("val"), "Quote")
        paragraph.insert(0, properties)

    after_blob = _mutate_document(before_blob, mutate)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)
    entry = report["paragraphs"]["changed"][0]
    assert entry["story"] == "document"
    assert entry["paragraph"] == 2
    assert "style" in entry["fields"]
    # The text itself did not move.
    assert entry["text_before"] == entry["text_after"]


# --------------------------------------------------------------------------
# doc_compare: added and removed paragraphs
# --------------------------------------------------------------------------


def test_an_added_paragraph_is_located_in_the_after_document(tmp_path):
    before_blob = build("simple")

    def add(root: etree._Element) -> None:
        body = root.find(_BODY)
        new_paragraph = etree.SubElement(body, _P)
        run = etree.SubElement(new_paragraph, _w("r"))
        text = etree.SubElement(run, _T)
        text.text = "A new, sixth paragraph."
        # Body children must end on w:sectPr: move it back to the end.
        sect_pr = body.find(_SECT_PR)
        if sect_pr is not None:
            body.remove(sect_pr)
            body.append(sect_pr)

    after_blob = _mutate_document(before_blob, add)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)

    assert report["paragraphs"]["removed"] == []
    assert len(report["paragraphs"]["added"]) == 1
    entry = report["paragraphs"]["added"][0]
    assert entry == {
        "story": "document",
        "paragraph": 5,
        "text": "A new, sixth paragraph.",
    }


def test_a_removed_paragraph_is_located_in_the_before_document(tmp_path):
    after_blob = build("simple")

    def add(root: etree._Element) -> None:
        body = root.find(_BODY)
        new_paragraph = etree.SubElement(body, _P)
        run = etree.SubElement(new_paragraph, _w("r"))
        text = etree.SubElement(run, _T)
        text.text = "A new, sixth paragraph."
        sect_pr = body.find(_SECT_PR)
        if sect_pr is not None:
            body.remove(sect_pr)
            body.append(sect_pr)

    before_blob = _mutate_document(after_blob, add)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)

    assert report["paragraphs"]["added"] == []
    assert len(report["paragraphs"]["removed"]) == 1
    entry = report["paragraphs"]["removed"][0]
    assert entry == {
        "story": "document",
        "paragraph": 5,
        "text": "A new, sixth paragraph.",
    }


# --------------------------------------------------------------------------
# doc_compare: the V2 and snapshot spaces diverge inside a table
# --------------------------------------------------------------------------


def test_a_changed_table_cell_paragraph_has_no_v2_index(tmp_path):
    """A paragraph inside a table cell is outside the V2 locator space (D-016)."""
    before_blob = build("tables")

    def mutate(root: etree._Element) -> None:
        cell_paragraph = next(
            p
            for p in root.iter(_P)
            if any(a.tag == _TC for a in p.iterancestors())
            and p.findtext(f".//{_T}") == "Row two, column two"
        )
        next(cell_paragraph.iter(_T)).text = "Row two, column two -- edited."

    after_blob = _mutate_document(before_blob, mutate)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)

    changed = report["paragraphs"]["changed"]
    assert len(changed) == 1
    entry = changed[0]
    assert entry["story"] == "document"
    assert entry["paragraph"] is None
    assert "text" in entry["fields"]
    assert entry["text_after"] == "Row two, column two -- edited."


def test_a_changed_table_reports_its_own_snapshot_space_index(tmp_path):
    """The ``table`` locator (main-story, top-level only) is a narrower space
    than the snapshot's table index (every ``w:tbl``, nested ones included);
    ``doc_compare`` reports the latter untranslated -- see the tool's docstring.
    """
    before_blob = build("tables")

    def widen(root: etree._Element) -> None:
        span = next(root.iter(_w("gridSpan")))
        span.set(_w("val"), "3")

    after_blob = _mutate_document(before_blob, widen)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)

    assert report["tables"]["added"] == []
    assert report["tables"]["removed"] == []
    assert len(report["tables"]["changed"]) == 1
    entry = report["tables"]["changed"][0]
    assert entry["story"] == "document"
    assert entry["table"] == 0
    assert "grid" in entry["fields"]


# --------------------------------------------------------------------------
# doc_compare: bookmarks and the global counters
# --------------------------------------------------------------------------


def test_a_new_bookmark_is_a_marker_change_and_a_counter_change(tmp_path):
    before_blob = build("simple")

    def add_bookmark(root: etree._Element) -> None:
        paragraph = _top_level_paragraphs(root)[1]
        start = etree.Element(_w("bookmarkStart"))
        start.set(_w("id"), "1")
        start.set(_w("name"), "intro")
        end = etree.Element(_w("bookmarkEnd"))
        end.set(_w("id"), "1")
        paragraph.insert(0, start)
        paragraph.append(end)

    after_blob = _mutate_document(before_blob, add_bookmark)
    a, b = _paths(tmp_path, before_blob, after_blob)

    report = doc_compare(a, b)

    entry = report["paragraphs"]["changed"][0]
    assert entry["paragraph"] == 1
    assert "markers" in entry["fields"]
    assert {"field": "bookmarks", "before": 0, "after": 1} in report["counters_changed"]
