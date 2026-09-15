"""Tests for :mod:`word_document_server.engine.table_styles`.

*Reading is delegated, never reimplemented.* :func:`list_table_styles` calls
:func:`~word_document_server.engine.styles.list_styles` for identity and
lineage; only the conditional-format inventory is read here, straight from the
style sheet through :class:`~word_document_server.engine.package.DocxPackage`'s
own public API -- not through a raw-XML accessor added to
:mod:`word_document_server.engine.styles`, which is a parallel part read-only
elsewhere in this milestone.

*Nothing half-applied.* :func:`apply_table_style` checks the style id and the
``look`` keys before writing a single element: a refused call leaves the table
exactly as it was, the same discipline :mod:`~word_document_server.engine.numbering`
holds for a list a caller cannot create.

*Schema order, not append order.* ``w:tblStyle`` and ``w:tblLook`` land at
their ECMA-376 §17.4.59 rank among ``w:tblPr``'s children, whether that
element already exists or is created for the call; a pre-existing sibling
(``w:tblW``, ``w:tblBorders``) never moves.

The python-docx template already ships about a hundred built-in table styles
(``TableGrid``, ``TableNormal``, the ``LightList``/``MediumGrid``/``ColorfulShading``
families and their accent variants), none of which carries a ``w:tblStylePr``.
A test that needs a style of its own -- to check lineage, conditional formats,
or that a given id is unknown -- appends a minimal ``w:style`` of family
``table``, named ``Fixture...`` so it can never collide with a built-in id,
straight to the live style sheet of an opened package: the same way
``tests/engine/test_styles_read.py`` builds its malformed style sheets, a
property of this reader/writer rather than a fixture other parts need.
"""

from __future__ import annotations

import asyncio
import io
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest
from docx import Document as PDocument
from lxml import etree

from tests.fixtures.builders import build
from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.engine.errors import PackageError
from word_document_server.engine.locators import _tables
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.styles import STYLES_PARTNAME, list_styles
from word_document_server.engine.table_styles import (
    LOOK_FLAGS,
    TBL_PR_ORDER,
    TableStyleInfo,
    apply_table_style,
    clear_table_style,
    list_table_styles,
    resolve_look,
)
from word_document_server.engine.xmlns import NAMESPACES, qn
from word_document_server.tools.v2.registry import v2_tools

W_TBL = qn("w:tbl")
W_TBL_PR = qn("w:tblPr")
W_TBL_STYLE = qn("w:tblStyle")
W_TBL_LOOK = qn("w:tblLook")
W_TBL_GRID = qn("w:tblGrid")
W_TR = qn("w:tr")
W_VAL = qn("w:val")

#: A minimal two-column grid plus one row, reused by every standalone table
#: built here -- the fixtures under test are ``w:tblPr``, not the grid.
_GRID = '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>'
_ROW = (
    "<w:tr>"
    '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    "</w:tr>"
)


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------


def _pkg(name: str) -> DocxPackage:
    return DocxPackage.open(build(name))


def _fragment(xml: str) -> etree._Element:
    """Parse a ``w:``-prefixed fragment, declaring the namespaces it uses."""
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in NAMESPACES.items())
    stripped = xml.strip()
    match = re.match(r"<[A-Za-z0-9_:.-]+", stripped)
    assert match is not None, f"not a tag: {stripped[:40]!r}"
    return etree.fromstring(
        f"{stripped[: match.end()]} {declarations}{stripped[match.end() :]}"
    )


def _table(children: str) -> etree._Element:
    """A standalone ``w:tbl`` (no ``w:tblPr``) with `children` after the grid."""
    return _fragment(f"<w:tbl>{children}</w:tbl>")


def _styles_element(pkg: DocxPackage) -> etree._Element:
    return pkg.root_of(pkg.part(STYLES_PARTNAME))


def _add_table_style(
    pkg: DocxPackage,
    style_id: str,
    *,
    name: str | None = None,
    based_on: str | None = None,
    conditional_formats: tuple[str, ...] = (),
) -> etree._Element:
    """Append a minimal ``w:style`` of family ``table`` to the live style sheet.

    Nothing here applies the style to any table -- it only makes `style_id` one
    :func:`~word_document_server.engine.table_styles.apply_table_style` accepts.
    """
    body = [f'<w:name w:val="{name or style_id}"/>']
    if based_on:
        body.append(f'<w:basedOn w:val="{based_on}"/>')
    for kind in conditional_formats:
        body.append(f'<w:tblStylePr w:type="{kind}"><w:tcPr/></w:tblStylePr>')
    element = _fragment(
        f'<w:style w:type="table" w:styleId="{style_id}">{"".join(body)}</w:style>'
    )
    _styles_element(pkg).append(element)
    return element


def names(element: etree._Element | None) -> list[str]:
    """Local names of the element's children, in order."""
    return [] if element is None else [etree.QName(child).localname for child in element]


def call(name: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Call a V2 tool the way an MCP client reaches it, and return its report."""
    return asyncio.run(v2_tools()[name](*args, **kwargs))


def _document_with_a_table(path: Path, style_id: str = "Grid") -> Path:
    """A one-table document on disk, with `style_id` defined but not applied."""
    document = PDocument()
    document.add_table(rows=2, cols=2)
    document.save(str(path))
    pkg = DocxPackage.open(str(path))
    _add_table_style(pkg, style_id)
    pkg.save(path)
    return path


# --------------------------------------------------------------------------------------
# TBL_PR_ORDER
# --------------------------------------------------------------------------------------


def test_tblStyle_precedes_tblW_and_tblLook_comes_late() -> None:
    assert TBL_PR_ORDER.index("w:tblStyle") < TBL_PR_ORDER.index("w:tblW")
    assert TBL_PR_ORDER.index("w:tblLook") > TBL_PR_ORDER.index("w:tblBorders")


# --------------------------------------------------------------------------------------
# list_table_styles
# --------------------------------------------------------------------------------------


def _without_styles(blob: bytes) -> bytes:
    """`blob` with the styles part, its content-type override and its rel removed."""
    source = zipfile.ZipFile(io.BytesIO(blob))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            if name == "word/styles.xml":
                continue
            data = source.read(name)
            if name == "[Content_Types].xml":
                data = re.sub(rb'<Override PartName="/word/styles\.xml"[^>]*/>', b"", data)
            if name == "word/_rels/document.xml.rels":
                data = re.sub(rb'<Relationship[^>]*styles\.xml"[^>]*/>', b"", data)
            target.writestr(name, data)
    return buffer.getvalue()


def test_a_package_with_no_style_sheet_defines_no_table_style() -> None:
    pkg = DocxPackage.open(_without_styles(build("simple")))
    assert list_table_styles(pkg) == []


def test_it_reports_exactly_the_ids_list_styles_reports_for_the_family() -> None:
    # The python-docx template ships about a hundred built-in table styles;
    # this pins delegation to list_styles rather than re-listing them here.
    pkg = _pkg("simple")
    expected = {info.style_id for info in list_styles(pkg, family="table")}

    assert {info.style_id for info in list_table_styles(pkg)} == expected


def test_a_table_style_is_reported_with_its_identity_and_lineage() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "FixtureGrid", name="Fixture Grid")
    _add_table_style(pkg, "FixtureGridFancy", name="Fancy Fixture Grid", based_on="FixtureGrid")

    found = {info.style_id: info for info in list_table_styles(pkg)}
    assert found["FixtureGrid"] == TableStyleInfo(
        style_id="FixtureGrid", name="Fixture Grid", based_on=None, conditional_formats=()
    )
    assert found["FixtureGridFancy"].based_on == "FixtureGrid"


def test_only_the_table_family_is_reported() -> None:
    # "style_inheritance" defines paragraph styles FixtureRoot/FixtureLeaf.
    pkg = _pkg("style_inheritance")
    _add_table_style(pkg, "FixtureTableStyle")

    ids = {info.style_id for info in list_table_styles(pkg)}
    assert "FixtureTableStyle" in ids
    assert "FixtureRoot" not in ids
    assert "FixtureLeaf" not in ids


def test_conditional_formats_are_read_in_document_order() -> None:
    pkg = _pkg("simple")
    _add_table_style(
        pkg,
        "FixtureFancy",
        conditional_formats=("wholeTable", "firstRow", "lastRow", "band1Horz"),
    )

    found = {info.style_id: info for info in list_table_styles(pkg)}
    assert found["FixtureFancy"].conditional_formats == (
        "wholeTable",
        "firstRow",
        "lastRow",
        "band1Horz",
    )


def test_a_style_with_no_conditional_format_reports_an_empty_tuple() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "FixturePlain")

    found = {info.style_id: info for info in list_table_styles(pkg)}
    assert found["FixturePlain"].conditional_formats == ()


# --------------------------------------------------------------------------------------
# resolve_look
# --------------------------------------------------------------------------------------


def test_resolve_look_of_none_turns_every_flag_off() -> None:
    assert resolve_look(None) == dict.fromkeys(LOOK_FLAGS, False)


def test_resolve_look_keeps_only_the_flags_given() -> None:
    resolved = resolve_look({"first_row": True, "no_v_band": True})
    assert resolved == {
        "first_row": True,
        "last_row": False,
        "first_column": False,
        "last_column": False,
        "no_h_band": False,
        "no_v_band": True,
    }


def test_resolve_look_rejects_an_unknown_flag() -> None:
    with pytest.raises(ValueError, match="header_row"):
        resolve_look({"header_row": True})


# --------------------------------------------------------------------------------------
# apply_table_style -- placement in w:tblPr
# --------------------------------------------------------------------------------------


def test_apply_creates_tblPr_as_the_first_child_when_absent() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _table(_GRID + _ROW)

    apply_table_style(pkg, table, "Grid")

    assert names(table)[0] == "tblPr"
    tbl_pr = table.find(W_TBL_PR)
    assert names(tbl_pr) == ["tblStyle", "tblLook"]
    assert tbl_pr.find(W_TBL_STYLE).get(W_VAL) == "Grid"


def test_apply_inserts_around_existing_children_in_schema_order() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _fragment(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders/></w:tblPr>'
        + _GRID
        + _ROW
        + "</w:tbl>"
    )

    apply_table_style(pkg, table, "Grid")

    assert names(table.find(W_TBL_PR)) == ["tblStyle", "tblW", "tblBorders", "tblLook"]


def test_apply_does_not_touch_other_tblPr_children() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _fragment(
        '<w:tbl><w:tblPr><w:tblW w:w="4500" w:type="dxa"/></w:tblPr>' + _GRID + _ROW + "</w:tbl>"
    )

    apply_table_style(pkg, table, "Grid")

    tbl_w = table.find(W_TBL_PR).find(qn("w:tblW"))
    assert tbl_w.get(qn("w:w")) == "4500"
    assert tbl_w.get(qn("w:type")) == "dxa"


def test_reapplying_updates_the_existing_elements_rather_than_duplicating() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    _add_table_style(pkg, "Banded")
    table = _table(_GRID + _ROW)

    apply_table_style(pkg, table, "Grid", {"first_row": True})
    apply_table_style(pkg, table, "Banded", {"no_v_band": True})

    tbl_pr = table.find(W_TBL_PR)
    assert len(tbl_pr.findall(W_TBL_STYLE)) == 1
    assert len(tbl_pr.findall(W_TBL_LOOK)) == 1
    assert tbl_pr.find(W_TBL_STYLE).get(W_VAL) == "Banded"
    assert tbl_pr.find(W_TBL_LOOK).get(W_VAL) == "0400"
    assert tbl_pr.find(W_TBL_LOOK).get(qn("w:firstRow")) == "0"


# --------------------------------------------------------------------------------------
# apply_table_style -- w:tblLook mask and explicit attributes
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("look", "mask", "attrs"),
    [
        (
            None,
            "0000",
            {
                "firstRow": "0",
                "lastRow": "0",
                "firstColumn": "0",
                "lastColumn": "0",
                "noHBand": "0",
                "noVBand": "0",
            },
        ),
        (
            # Matches python-docx's own default look for a plain w:tbl.
            {"first_row": True, "first_column": True, "no_v_band": True},
            "04A0",
            {
                "firstRow": "1",
                "lastRow": "0",
                "firstColumn": "1",
                "lastColumn": "0",
                "noHBand": "0",
                "noVBand": "1",
            },
        ),
        (
            {"last_row": True, "last_column": True, "no_h_band": True},
            "0340",
            {
                "firstRow": "0",
                "lastRow": "1",
                "firstColumn": "0",
                "lastColumn": "1",
                "noHBand": "1",
                "noVBand": "0",
            },
        ),
    ],
)
def test_look_flags_produce_the_documented_mask(
    look: dict[str, bool] | None, mask: str, attrs: dict[str, str]
) -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _table(_GRID + _ROW)

    apply_table_style(pkg, table, "Grid", look)

    element = table.find(W_TBL_PR).find(W_TBL_LOOK)
    assert element.get(W_VAL) == mask
    for attribute, expected in attrs.items():
        assert element.get(qn(f"w:{attribute}")) == expected


def test_look_attributes_are_written_in_words_own_order() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _table(_GRID + _ROW)

    apply_table_style(pkg, table, "Grid", {"first_row": True})

    element = table.find(W_TBL_PR).find(W_TBL_LOOK)
    assert list(element.attrib.keys()) == [
        W_VAL,
        qn("w:firstRow"),
        qn("w:lastRow"),
        qn("w:firstColumn"),
        qn("w:lastColumn"),
        qn("w:noHBand"),
        qn("w:noVBand"),
    ]


# --------------------------------------------------------------------------------------
# apply_table_style -- refusals write nothing
# --------------------------------------------------------------------------------------


def test_applying_an_undefined_style_is_refused_and_writes_nothing() -> None:
    pkg = _pkg("simple")
    table = _table(_GRID + _ROW)
    before = etree.tostring(table)

    with pytest.raises(PackageError, match="Ghost"):
        apply_table_style(pkg, table, "Ghost")

    assert etree.tostring(table) == before


@pytest.mark.parametrize("style_id", ["", "   "])
def test_an_empty_style_id_is_refused(style_id: str) -> None:
    pkg = _pkg("simple")
    table = _table(_GRID + _ROW)
    before = etree.tostring(table)

    with pytest.raises(ValueError):
        apply_table_style(pkg, table, style_id)

    assert etree.tostring(table) == before


def test_an_unknown_look_flag_is_refused_and_writes_nothing() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _table(_GRID + _ROW)
    before = etree.tostring(table)

    with pytest.raises(ValueError, match="header_row"):
        apply_table_style(pkg, table, "Grid", {"header_row": True})

    assert etree.tostring(table) == before


def test_apply_refuses_a_non_table_element() -> None:
    pkg = _pkg("simple")
    with pytest.raises(TypeError):
        apply_table_style(pkg, _fragment("<w:p/>"), "Grid")


def test_clear_refuses_a_non_table_element() -> None:
    with pytest.raises(TypeError):
        clear_table_style(_fragment("<w:p/>"))


# --------------------------------------------------------------------------------------
# clear_table_style
# --------------------------------------------------------------------------------------


def test_clear_removes_tblStyle_but_leaves_tblLook_and_siblings() -> None:
    pkg = _pkg("simple")
    _add_table_style(pkg, "Grid")
    table = _fragment(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' + _GRID + _ROW + "</w:tbl>"
    )
    apply_table_style(pkg, table, "Grid", {"first_row": True})

    clear_table_style(table)

    tbl_pr = table.find(W_TBL_PR)
    assert tbl_pr.find(W_TBL_STYLE) is None
    assert tbl_pr.find(qn("w:tblW")) is not None
    look = tbl_pr.find(W_TBL_LOOK)
    assert look is not None and look.get(qn("w:firstRow")) == "1"


def test_clear_on_a_table_with_no_tblPr_is_a_no_op() -> None:
    table = _table(_GRID + _ROW)
    before = etree.tostring(table)

    clear_table_style(table)

    assert table.find(W_TBL_PR) is None
    assert etree.tostring(table) == before


def test_clear_on_a_table_with_tblPr_but_no_tblStyle_is_a_no_op() -> None:
    table = _fragment(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' + _GRID + _ROW + "</w:tbl>"
    )
    before = etree.tostring(table)

    clear_table_style(table)

    assert etree.tostring(table) == before


# --------------------------------------------------------------------------------------
# doc_apply_table_style, the V2 tool over this module
# --------------------------------------------------------------------------------------


def test_doc_apply_table_style_applies_an_existing_style(tmp_path: Path) -> None:
    path = _document_with_a_table(tmp_path / "table.docx", style_id="Grid")
    before = snapshot(path.read_bytes())

    report = call(
        "doc_apply_table_style",
        str(path),
        0,
        "Grid",
        {"first_row": True, "no_v_band": True},
    )

    assert report["status"] == "ok"
    assert report["saved"] is True
    assert report["table_index"] == 0
    assert report["style"] == "Grid"
    assert report["look"] == {
        "first_row": True,
        "last_row": False,
        "first_column": False,
        "last_column": False,
        "no_h_band": False,
        "no_v_band": True,
    }
    assert report["changes"] == []

    pkg = DocxPackage.open(str(path))
    table = _tables(pkg.document)[0]
    tbl_pr = table.find(W_TBL_PR)
    assert names(tbl_pr) == ["tblStyle", "tblW", "tblLook"]
    assert tbl_pr.find(W_TBL_STYLE).get(W_VAL) == "Grid"
    assert tbl_pr.find(W_TBL_LOOK).get(W_VAL) == "0420"
    # The table's own shape never moved.
    assert len(table.findall(W_TR)) == 2
    assert len(table.find(W_TBL_GRID).findall(qn("w:gridCol"))) == 2

    assert_unchanged_except(
        before, snapshot(path.read_bytes()), parts=["word/document.xml"]
    )
    assert validate_package(path) == []


def test_doc_apply_table_style_style_none_clears_it(tmp_path: Path) -> None:
    path = _document_with_a_table(tmp_path / "table.docx", style_id="Grid")
    call("doc_apply_table_style", str(path), 0, "Grid", {"first_row": True})

    report = call("doc_apply_table_style", str(path), 0, None)

    assert report["status"] == "ok"
    assert report["style"] is None
    assert report["look"] is None

    pkg = DocxPackage.open(str(path))
    tbl_pr = _tables(pkg.document)[0].find(W_TBL_PR)
    assert tbl_pr.find(W_TBL_STYLE) is None
    # Clearing the style does not touch the look flags already written.
    assert tbl_pr.find(W_TBL_LOOK).get(qn("w:firstRow")) == "1"
    assert validate_package(path) == []


def test_doc_apply_table_style_addresses_a_nested_table_by_its_own_index(
    tmp_path: Path,
) -> None:
    path = tmp_path / "tables.docx"
    path.write_bytes(build("tables"))
    pkg = DocxPackage.open(str(path))
    _add_table_style(pkg, "Grid")
    pkg.save(path)

    report = call("doc_apply_table_style", str(path), 1, "Grid")

    assert report["status"] == "ok"
    pkg = DocxPackage.open(str(path))
    tables = _tables(pkg.document)
    assert tables[1].find(W_TBL_PR).find(W_TBL_STYLE).get(W_VAL) == "Grid"
    # The outer table (index 0) is untouched: it still names the fixture's own
    # unresolved reference.
    assert tables[0].find(W_TBL_PR).find(W_TBL_STYLE).get(W_VAL) == "TableGrid"
    assert validate_package(path) == []


@pytest.mark.parametrize(
    ("table_index", "style", "look", "code"),
    [
        (0, "Ghost", None, "package_error"),
        (5, "Grid", None, "not_found"),
        (-1, "Grid", None, "invalid_argument"),
        ("0", "Grid", None, "invalid_argument"),
        (True, "Grid", None, "invalid_argument"),
        (0, "", None, "invalid_argument"),
        (0, "Grid", {"header_row": True}, "invalid_argument"),
    ],
)
def test_a_refused_call_writes_nothing(
    tmp_path: Path,
    table_index: Any,
    style: str | None,
    look: dict[str, bool] | None,
    code: str,
) -> None:
    path = _document_with_a_table(tmp_path / "table.docx", style_id="Grid")
    before = path.read_bytes()

    report = call("doc_apply_table_style", str(path), table_index, style, look)

    assert report == {"status": "error", "code": code, "message": report["message"]}
    assert path.read_bytes() == before


def test_a_dry_run_reports_the_change_and_writes_nothing(tmp_path: Path) -> None:
    path = _document_with_a_table(tmp_path / "table.docx", style_id="Grid")
    before = path.read_bytes()

    report = call(
        "doc_apply_table_style", str(path), 0, "Grid", {"first_row": True}, dry_run=True
    )

    assert report["dry_run"] is True
    assert report["saved"] is False
    assert report["style"] == "Grid"
    assert path.read_bytes() == before
