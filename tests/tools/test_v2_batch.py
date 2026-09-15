"""``doc_apply_edits``: several edits, one in-memory package, one atomic save.

What is pinned here is not the individual edit operations -- those are already
covered by ``tests/tools/test_v2_text.py``, and this module reuses the exact
same locators and payload shapes -- but the three promises specific to the
batch: everything lands in the same report shape as `doc_edit_text` and
`doc_format_range`; a later edit can target what an earlier edit in the same
call just wrote, because locators are resolved against the live document, not
a snapshot taken before the call; and the first edit that fails leaves the
file exactly as it was, with the index of the offending edit in the report.

No helper is imported from another test module (see J04-P7's note on this):
``call`` is redefined here, the way ``test_v2_text.py`` does.
"""

from __future__ import annotations

import asyncio
from typing import Any

import pytest

from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, snapshot
from word_document_server.tools.v2.registry import (
    discover_tool_specs,
    register_v2_tools,
    v2_tools,
)

#: Keys the registry fills in on every successful call.
ENVELOPE = {"status", "dry_run", "changes", "warnings"}


def call(name: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Call a V2 tool the way an MCP client reaches it, and return its report."""
    return asyncio.run(v2_tools()[name](*args, **kwargs))


def apply_edits(path, edits, **kwargs) -> dict[str, Any]:
    """Shorthand for calling ``doc_apply_edits`` on `path`."""
    return call("doc_apply_edits", str(path), edits, **kwargs)


def texts(report: dict[str, Any]) -> list[str]:
    """The ``after`` text of every change in `report`."""
    return [change["after"] for change in report["changes"]]


# --------------------------------------------------------------------------
# Discovery and the envelope
# --------------------------------------------------------------------------


def test_apply_edits_is_discovered_and_registered():
    from fastmcp import FastMCP

    discovered = {spec.name for spec in discover_tool_specs()}
    assert "doc_apply_edits" in discovered

    server = FastMCP("test")
    registered = register_v2_tools(server)
    assert "doc_apply_edits" in registered


def test_the_report_carries_the_four_shared_keys(fixture_docx):
    report = apply_edits(
        fixture_docx("simple"),
        [{"locator": {"paragraph": 1}, "text": "x"}],
    )

    assert report["status"] == "ok"
    assert ENVELOPE <= set(report)
    assert isinstance(report["changes"], list)
    assert isinstance(report["warnings"], list)


# --------------------------------------------------------------------------
# Success: a mix of text and formatting edits, one atomic save
# --------------------------------------------------------------------------


def test_a_batch_of_edits_is_applied_and_saved_once(fixture_docx):
    path = fixture_docx("simple")
    before = snapshot(path)

    report = apply_edits(
        path,
        [
            {"locator": {"paragraph": 1}, "action": "replace", "text": "Rewritten first."},
            {"locator": {"find": "Rewritten first."}, "patch": {"bold": True}},
        ],
    )

    assert report["status"] == "ok"
    assert report["saved"] is True
    assert texts(report) == ["Rewritten first.", "Rewritten first."]
    assert validate_package(path) == []
    assert_unchanged_except(before, snapshot(path), paragraphs=[1])


def test_a_later_locator_resolves_against_what_an_earlier_edit_just_wrote(fixture_docx):
    """Locators are resolved one edit at a time, never off a snapshot."""
    path = fixture_docx("simple")

    report = apply_edits(
        path,
        [
            {"locator": {"paragraph": 1}, "action": "insert", "text": "Marker. ", "start": 0},
            {"locator": {"find": "Marker. "}, "action": "delete"},
        ],
    )

    assert report["status"] == "ok"
    assert texts(report) == [
        "Marker. First paragraph of the simple fixture.",
        "First paragraph of the simple fixture.",
    ]
    assert validate_package(path) == []


def test_an_edit_without_an_action_defaults_to_replace(fixture_docx):
    path = fixture_docx("simple")

    report = apply_edits(path, [{"locator": {"paragraph": 1}, "text": "Replaced."}])

    assert texts(report) == ["Replaced."]


def test_an_edit_that_changes_nothing_is_not_in_changes_but_warns(fixture_docx):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = apply_edits(path, [{"locator": {"paragraph": 1}, "action": "insert", "text": ""}])

    assert report["status"] == "ok"
    assert report["saved"] is False
    assert report["changes"] == []
    assert report["warnings"]
    assert path.read_bytes() == before


# --------------------------------------------------------------------------
# Failure in the middle of a batch: nothing written, index reported
# --------------------------------------------------------------------------


def test_a_failure_in_the_middle_writes_nothing_and_names_the_failing_edit(fixture_docx):
    path = fixture_docx("simple")
    before_bytes = path.read_bytes()
    before_snapshot = snapshot(path)

    report = apply_edits(
        path,
        [
            {"locator": {"paragraph": 1}, "action": "replace", "text": "This must not land."},
            {"locator": {"paragraph": 99}, "action": "replace", "text": "y"},
            {"locator": {"paragraph": 2}, "action": "replace", "text": "Nor this."},
        ],
    )

    assert report["status"] == "error"
    assert report["code"] == "not_found"
    assert "edit 1" in report["message"]
    assert path.read_bytes() == before_bytes
    assert_unchanged_except(before_snapshot, snapshot(path))


def test_a_locator_failure_keeps_the_engines_own_code(fixture_docx):
    """D-022: the failing edit's own code survives being wrapped with its index."""
    path = fixture_docx("simple")

    report = apply_edits(path, [{"locator": {"paragraph": "one"}, "text": "x"}])

    assert report["status"] == "error"
    assert report["code"] == "invalid"
    assert "edit 0" in report["message"]


def test_a_bad_patch_fails_at_its_own_index_and_writes_nothing(fixture_docx):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = apply_edits(
        path,
        [
            {"locator": {"paragraph": 1}, "text": "Landed only in memory."},
            {"locator": {"paragraph": 2}, "patch": {"glowing": True}},
        ],
    )

    assert report["status"] == "error"
    assert report["code"] == "invalid_argument"
    assert "edit 1" in report["message"]
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    ("edits", "code"),
    [
        ([], "invalid_argument"),
        (["not a mapping"], "invalid_argument"),
    ],
)
def test_a_malformed_batch_is_refused_before_anything_is_touched(fixture_docx, edits, code):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = apply_edits(path, edits)

    assert report["status"] == "error"
    assert report["code"] == code
    assert path.read_bytes() == before


# --------------------------------------------------------------------------
# dry_run
# --------------------------------------------------------------------------


def test_dry_run_reports_every_edit_and_writes_nothing(fixture_docx):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = apply_edits(
        path,
        [
            {"locator": {"paragraph": 1}, "text": "Would be rewritten."},
            {"locator": {"paragraph": 2}, "patch": {"italic": True}},
        ],
        dry_run=True,
    )

    assert report["dry_run"] is True
    assert report["saved"] is False
    assert texts(report) == ["Would be rewritten.", "Second paragraph, with a trailing sentence."]
    assert path.read_bytes() == before


def test_dry_run_still_fails_at_the_offending_edit(fixture_docx):
    path = fixture_docx("simple")
    before = path.read_bytes()

    report = apply_edits(
        path,
        [
            {"locator": {"paragraph": 1}, "text": "x"},
            {"locator": {"paragraph": 99}, "text": "y"},
        ],
        dry_run=True,
    )

    assert report["status"] == "error"
    assert report["code"] == "not_found"
    assert "edit 1" in report["message"]
    assert path.read_bytes() == before
