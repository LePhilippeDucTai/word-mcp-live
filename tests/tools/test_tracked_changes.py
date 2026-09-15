"""Tests for :mod:`word_document_server.core.tracked_changes`.

The module under test is an adapter, so these tests are written against what a
caller of the MCP tool can observe: the dictionary it gets back, and the
package on disk before and after the call.  The package is judged with
:mod:`tests.support.snapshot` -- ``assert_unchanged_except`` names the
paragraphs a call was allowed to touch and guards every other paragraph, part,
relationship and section of the document -- and with
:func:`tests.support.package_check.validate_package`.

The three defects the migration to
:mod:`word_document_server.engine.revisions` had to remove are pinned here as
regressions, since nothing else in the suite covers them at the tool level:

``test_replace_terminates_when_the_new_text_contains_the_old``
    the ``while True`` re-search that never returned for ``"Risk" ->
    "Risk Risk"``.  It carries its own short timeout so that a reintroduced loop
    fails the test instead of hanging the run.
``test_delete_inside_a_run_keeps_the_properties_of_each_fragment``
    the splice that rebuilt every surviving fragment with the *first* matched
    run's ``w:rPr``.
``test_text_hidden_under_a_deletion_is_never_matched`` and its siblings
    the scan over raw ``w:t`` that could match, and re-delete, text already
    hidden under a ``w:del``.

Paragraph indices are resolved by text through :func:`index_of`, never
hard-coded: the index space here is the snapshot one (every ``w:p`` of a story,
document order, base 0), which is neither the V2 locator space nor python-docx's
own -- see ``tests/support/snapshot.py``.
"""

from __future__ import annotations

import ast
import inspect

import pytest

from tests.support.package_check import validate_package
from tests.support.snapshot import Snapshot, assert_unchanged_except, snapshot
from word_document_server.core import tracked_changes as module
from word_document_server.core.tracked_changes import (
    accept_tracked_changes_in_doc,
    list_tracked_changes_in_doc,
    reject_tracked_changes_in_doc,
    track_delete_in_doc,
    track_insert_in_doc,
    track_replace_in_doc,
)
from word_document_server.engine.errors import UnsupportedRevision

AUTHOR = "Bob Reviewer"
OTHER = "Alice Author"

#: The author every fixture revision carries; see ``tests/fixtures/builders.py``.
FIXTURE_AUTHOR = "Fixture Author"
FIXTURE_REVIEWER = "Fixture Reviewer"


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def index_of(snap: Snapshot, text: str, story: str = "document") -> int:
    """Snapshot-space index of the one paragraph of `story` whose text is `text`."""
    found = [i for i, sig in enumerate(snap.story_paragraphs(story)) if sig.text == text]
    assert len(found) == 1, (
        f"expected exactly one {story!r} paragraph with text {text!r} "
        f"(snapshot space), found at {found}"
    )
    return found[0]


def runs_of(snap: Snapshot, index: int, story: str = "document"):
    """``(text, rPr)`` of every visible run of one paragraph, in document order."""
    signature = snap.paragraphs[(story, index)]
    return [(run.text, run.rpr) for run in signature.runs]


def deleted_runs_of(snap: Snapshot, index: int, story: str = "document"):
    """``(text, rPr)`` of every run hidden under a ``w:del``, in document order."""
    signature = snap.paragraphs[(story, index)]
    return [(run.text, run.rpr) for run in signature.deleted_runs]


def properties_by_text(snap: Snapshot, index: int) -> dict[str, str]:
    """Map each visible run's text to its canonical ``w:rPr`` in a paragraph."""
    return {run.text: run.rpr for run in snap.paragraphs[("document", index)].runs}


@pytest.fixture
def simple_path(fixture_docx):
    """A fresh ``simple.docx``: five paragraphs, no revision at all."""
    return fixture_docx("simple")


@pytest.fixture
def mixed_path(fixture_docx):
    """A fresh ``mixed_runs.docx``: one paragraph of four differently formatted runs."""
    return fixture_docx("mixed_runs")


@pytest.fixture
def tracked_path(fixture_docx):
    """A fresh ``tracked_changes.docx``: every revision kind the engine reports."""
    return fixture_docx("tracked_changes")


@pytest.fixture
def combined_path(fixture_docx):
    """A fresh ``combined.docx``: the revisions plus tables, headers, notes, fields."""
    return fixture_docx("combined")


# --------------------------------------------------------------------------
# The defects the migration removed
# --------------------------------------------------------------------------


@pytest.mark.timeout(20)
def test_replace_terminates_when_the_new_text_contains_the_old(simple_path):
    """Regression: ``"Risk" -> "Risk Risk"`` used to grow the paragraph forever."""
    before = snapshot(simple_path.read_bytes())
    index = index_of(before, "Second paragraph, with a trailing sentence.")

    result = track_replace_in_doc(
        str(simple_path), "trailing", "trailing trailing", AUTHOR
    )

    assert result["success"] is True
    assert result["replacements"] == 1
    after = snapshot(simple_path.read_bytes())
    assert after.paragraphs[("document", index)].text == (
        "Second paragraph, with a trailing trailing sentence."
    )
    assert after.paragraphs[("document", index)].deleted_text == "trailing"
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(simple_path) == []


@pytest.mark.timeout(20)
def test_replace_terminates_when_every_occurrence_grows(simple_path):
    """The same trap with several occurrences, and the count they must report.

    ``"heading"`` appears in ``"Second heading"`` and in ``"Body text under the
    second heading."``: each one is replaced exactly once, and the text the
    replacement adds is never searched again.
    """
    result = track_replace_in_doc(str(simple_path), "heading", "heading heading", AUTHOR)

    assert result["success"] is True
    assert result["replacements"] == 2
    listed = list_tracked_changes_in_doc(str(simple_path))
    assert [e["text"] for e in listed["insertions"]] == ["heading heading"] * 2
    assert [e["text"] for e in listed["deletions"]] == ["heading"] * 2
    assert validate_package(simple_path) == []


def test_delete_inside_a_run_keeps_the_properties_of_each_fragment(mixed_path):
    """Regression: the fragments used to be rebuilt with the first run's ``w:rPr``."""
    before = snapshot(mixed_path.read_bytes())
    index = index_of(before, "Plain then bold bold italic et la fin.")
    rpr = properties_by_text(before, index)
    assert rpr["bold "] != rpr["bold italic "], "fixture must give the runs distinct rPr"

    result = track_delete_in_doc(str(mixed_path), "d ita", AUTHOR)
    assert result["success"] is True

    after = snapshot(mixed_path.read_bytes())
    # "d ita" sits strictly inside the bold+italic run: both surviving halves
    # keep *its* properties, and the untouched runs keep theirs.
    assert runs_of(after, index) == [
        ("Plain then ", rpr["Plain then "]),
        ("bold ", rpr["bold "]),
        ("bol", rpr["bold italic "]),
        ("lic ", rpr["bold italic "]),
        ("et la fin.", rpr["et la fin."]),
    ]
    assert deleted_runs_of(after, index) == [("d ita", rpr["bold italic "])]
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(mixed_path) == []


def test_replace_across_runs_keeps_the_properties_of_each_fragment(mixed_path):
    """A range spanning three runs leaves each untouched fragment as it was."""
    before = snapshot(mixed_path.read_bytes())
    index = index_of(before, "Plain then bold bold italic et la fin.")
    rpr = properties_by_text(before, index)

    result = track_replace_in_doc(str(mixed_path), "then bold bold", "THEN", AUTHOR)
    assert result["success"] is True

    after = snapshot(mixed_path.read_bytes())
    visible = runs_of(after, index)
    assert [text for text, _ in visible] == ["Plain ", "THEN", " italic ", "et la fin."]
    assert dict(visible)["Plain "] == rpr["Plain then "]
    assert dict(visible)[" italic "] == rpr["bold italic "]
    assert dict(visible)["et la fin."] == rpr["et la fin."]
    # The inserted run inherits the first run of the replaced range.
    assert dict(visible)["THEN"] == rpr["Plain then "]
    assert deleted_runs_of(after, index) == [
        ("then ", rpr["Plain then "]),
        ("bold ", rpr["bold "]),
        ("bold", rpr["bold italic "]),
    ]
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(mixed_path) == []


def test_text_hidden_under_a_deletion_is_never_matched(tracked_path):
    """``"deleted text, "`` only exists inside a ``w:del``: no operation sees it."""
    original = tracked_path.read_bytes()

    for call in (
        lambda: track_replace_in_doc(str(tracked_path), "deleted text", "x", AUTHOR),
        lambda: track_delete_in_doc(str(tracked_path), "deleted text", AUTHOR),
        lambda: track_insert_in_doc(str(tracked_path), "deleted text", "x", AUTHOR),
    ):
        result = call()
        assert result["success"] is False
        assert result["error"] == "Text not found: 'deleted text'"

    assert tracked_path.read_bytes() == original


def test_replacing_around_hidden_text_leaves_the_deletion_untouched(tracked_path):
    """The ``w:del`` of the fixture survives an edit made in its own paragraph."""
    before = snapshot(tracked_path.read_bytes())
    index = index_of(before, "Kept text, inserted text, and kept tail.")
    hidden = before.paragraphs[("document", index)].deleted_runs

    result = track_replace_in_doc(str(tracked_path), "kept tail", "kept end", AUTHOR)
    assert result["success"] is True
    assert result["replacements"] == 1

    after = snapshot(tracked_path.read_bytes())
    signature = after.paragraphs[("document", index)]
    assert hidden[0] in signature.deleted_runs, "the fixture's w:del must be untouched"
    assert signature.text == "Kept text, inserted text, and kept end."
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(tracked_path) == []


# --------------------------------------------------------------------------
# Recording: how many occurrences, and where
# --------------------------------------------------------------------------


def test_replace_marks_every_occurrence(simple_path):
    before = snapshot(simple_path.read_bytes())
    first = index_of(before, "First paragraph of the simple fixture.")
    second = index_of(before, "Second paragraph, with a trailing sentence.")

    result = track_replace_in_doc(str(simple_path), "paragraph", "section", AUTHOR)

    assert result["success"] is True
    assert result["replacements"] == 2
    after = snapshot(simple_path.read_bytes())
    assert after.paragraphs[("document", first)].text == (
        "First section of the simple fixture."
    )
    assert after.paragraphs[("document", second)].text == (
        "Second section, with a trailing sentence."
    )
    assert_unchanged_except(
        before, after, paragraphs=[first, second], counters=["revisions"]
    )
    assert validate_package(simple_path) == []


def test_delete_marks_every_occurrence(simple_path):
    before = snapshot(simple_path.read_bytes())
    first = index_of(before, "First paragraph of the simple fixture.")
    second = index_of(before, "Second paragraph, with a trailing sentence.")

    result = track_delete_in_doc(str(simple_path), "paragraph", AUTHOR)

    assert result["success"] is True
    assert result["deletions"] == 2
    after = snapshot(simple_path.read_bytes())
    assert after.paragraphs[("document", first)].text == "First  of the simple fixture."
    assert after.paragraphs[("document", first)].deleted_text == "paragraph"
    assert after.paragraphs[("document", second)].deleted_text == "paragraph"
    assert_unchanged_except(
        before, after, paragraphs=[first, second], counters=["revisions"]
    )
    assert validate_package(simple_path) == []


def test_insert_marks_the_first_occurrence_only(simple_path):
    before = snapshot(simple_path.read_bytes())
    first = index_of(before, "First paragraph of the simple fixture.")

    result = track_insert_in_doc(str(simple_path), "paragraph", " (added)", AUTHOR)

    assert result["success"] is True
    assert result["insertions"] == 1
    after = snapshot(simple_path.read_bytes())
    assert after.paragraphs[("document", first)].text == (
        "First paragraph (added) of the simple fixture."
    )
    assert_unchanged_except(before, after, paragraphs=[first], counters=["revisions"])
    assert validate_package(simple_path) == []


def test_a_missing_text_reports_the_failure_and_writes_nothing(simple_path):
    original = simple_path.read_bytes()

    assert track_replace_in_doc(str(simple_path), "absent", "x", AUTHOR) == {
        "success": False,
        "error": "Text not found: 'absent'",
    }
    assert track_delete_in_doc(str(simple_path), "absent", AUTHOR) == {
        "success": False,
        "error": "Text not found: 'absent'",
    }
    assert track_insert_in_doc(str(simple_path), "absent", "x", AUTHOR) == {
        "success": False,
        "error": "Text not found: 'absent'",
    }
    assert simple_path.read_bytes() == original


def test_recording_touches_no_other_part_of_a_rich_document(combined_path):
    """The whole point of the engine: a table, a header and an image stay put."""
    before = snapshot(combined_path.read_bytes())
    index = index_of(before, "Body text under the second heading.")

    result = track_replace_in_doc(str(combined_path), "Body text under", "Text under", AUTHOR)
    assert result["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(combined_path) == []


# --------------------------------------------------------------------------
# Listing
# --------------------------------------------------------------------------


def test_listing_reports_the_documented_keys(tracked_path):
    listed = list_tracked_changes_in_doc(str(tracked_path))

    assert set(listed) == {
        "success",
        "insertions",
        "deletions",
        "total_insertions",
        "total_deletions",
        "total_changes",
    }
    assert listed["success"] is True
    assert listed["total_changes"] == (
        listed["total_insertions"] + listed["total_deletions"]
    )
    for entry in listed["insertions"] + listed["deletions"]:
        assert set(entry) == {"id", "author", "date", "text", "paragraph_context"}
        assert isinstance(entry["id"], str)


def test_listing_reports_the_fixture_revisions(tracked_path):
    listed = list_tracked_changes_in_doc(str(tracked_path))

    assert [(e["id"], e["text"]) for e in listed["insertions"]] == [
        ("201", "inserted text, "),
        ("203", ""),
        ("208", ""),
    ]
    assert [(e["id"], e["text"], e["author"]) for e in listed["deletions"]] == [
        ("202", "deleted text, ", FIXTURE_AUTHOR),
        ("204", "inserted then deleted", FIXTURE_REVIEWER),
        ("207", "", FIXTURE_AUTHOR),
    ]
    context = {e["id"]: e["paragraph_context"] for e in listed["deletions"]}
    assert context["202"] == "Kept text, inserted text, and kept tail."
    assert context["207"].startswith("This paragraph mark is deleted")


def test_listing_a_document_without_revisions_is_empty(simple_path):
    listed = list_tracked_changes_in_doc(str(simple_path))

    assert listed == {
        "success": True,
        "insertions": [],
        "deletions": [],
        "total_insertions": 0,
        "total_deletions": 0,
        "total_changes": 0,
    }


def test_listing_reads_a_recorded_change_back(simple_path):
    track_replace_in_doc(str(simple_path), "First paragraph", "Opening", AUTHOR)

    listed = list_tracked_changes_in_doc(str(simple_path))
    assert [e["text"] for e in listed["insertions"]] == ["Opening"]
    assert [e["text"] for e in listed["deletions"]] == ["First paragraph"]
    assert {e["author"] for e in listed["insertions"] + listed["deletions"]} == {AUTHOR}
    assert listed["insertions"][0]["date"].endswith("Z")


# --------------------------------------------------------------------------
# Applying: by author, by id, and all of them
# --------------------------------------------------------------------------


def test_accept_by_author_leaves_the_other_authors_alone(simple_path):
    mine = index_of(
        snapshot(simple_path.read_bytes()), "First paragraph of the simple fixture."
    )
    track_replace_in_doc(str(simple_path), "First paragraph", "Opening", AUTHOR)
    track_replace_in_doc(str(simple_path), "Second paragraph", "Closing", OTHER)
    before = snapshot(simple_path.read_bytes())

    result = accept_tracked_changes_in_doc(str(simple_path), author=AUTHOR)

    assert result["success"] is True
    assert result["accepted"] == 2  # the insertion and the deletion of one replace
    after = snapshot(simple_path.read_bytes())
    assert after.paragraphs[("document", mine)].text == "Opening of the simple fixture."
    assert after.paragraphs[("document", mine)].deleted_text == ""
    listed = list_tracked_changes_in_doc(str(simple_path))
    assert {e["author"] for e in listed["insertions"] + listed["deletions"]} == {OTHER}
    assert_unchanged_except(before, after, paragraphs=[mine], counters=["revisions"])
    assert validate_package(simple_path) == []


def test_reject_by_author_restores_only_that_author_s_paragraph(simple_path):
    original = snapshot(simple_path.read_bytes())
    mine = index_of(original, "First paragraph of the simple fixture.")
    track_replace_in_doc(str(simple_path), "First paragraph", "Opening", AUTHOR)
    track_replace_in_doc(str(simple_path), "Second paragraph", "Closing", OTHER)

    result = reject_tracked_changes_in_doc(str(simple_path), author=AUTHOR)

    assert result["success"] is True
    assert result["rejected"] == 2
    after = snapshot(simple_path.read_bytes())
    was = original.paragraphs[("document", mine)]
    restored = after.paragraphs[("document", mine)]
    assert restored.text == was.text
    assert restored.deleted_runs == ()
    assert (restored.style, restored.ppr, restored.markers) == (
        was.style,
        was.ppr,
        was.markers,
    )
    # Resolving the range split the run it started inside, and rejecting does
    # not glue it back: the boundary stays visible in the run structure. That
    # trace is text-preserving and carries the original properties, which is the
    # most that can be asserted here.
    assert "".join(run.text for run in restored.runs) == was.text
    assert {run.rpr for run in restored.runs} == {run.rpr for run in was.runs}
    assert validate_package(simple_path) == []


def test_accept_by_id_applies_only_that_revision(tracked_path):
    before = snapshot(tracked_path.read_bytes())
    index = index_of(before, "Kept text, inserted text, and kept tail.")

    result = accept_tracked_changes_in_doc(str(tracked_path), change_ids=[202])

    assert result["success"] is True
    assert result["accepted"] == 1
    after = snapshot(tracked_path.read_bytes())
    # 202 is the deletion of "deleted text, ": accepting it drops the text.
    assert after.paragraphs[("document", index)].deleted_text == ""
    assert after.paragraphs[("document", index)].text == (
        "Kept text, inserted text, and kept tail."
    )
    listed = list_tracked_changes_in_doc(str(tracked_path))
    assert "202" not in {e["id"] for e in listed["deletions"]}
    assert "201" in {e["id"] for e in listed["insertions"]}
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(tracked_path) == []


def test_reject_by_id_applies_only_that_revision(tracked_path):
    before = snapshot(tracked_path.read_bytes())
    index = index_of(before, "Kept text, inserted text, and kept tail.")

    result = reject_tracked_changes_in_doc(str(tracked_path), change_ids=[201])

    assert result["success"] is True
    assert result["rejected"] == 1
    after = snapshot(tracked_path.read_bytes())
    # 201 is the insertion of "inserted text, ": rejecting it removes the text.
    assert after.paragraphs[("document", index)].text == "Kept text, and kept tail."
    assert_unchanged_except(before, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(tracked_path) == []


def test_ids_and_author_combine(tracked_path):
    """A selection that names an id belonging to another author matches nothing."""
    original = tracked_path.read_bytes()

    result = accept_tracked_changes_in_doc(
        str(tracked_path), author=FIXTURE_REVIEWER, change_ids=[201, 202]
    )

    assert result == {
        "success": True,
        "message": "No matching tracked changes found to accept",
        "accepted": 0,
    }
    assert tracked_path.read_bytes() == original


def test_an_id_the_document_does_not_carry_is_reported_not_ignored(tracked_path):
    result = accept_tracked_changes_in_doc(str(tracked_path), change_ids=[201, 9999])

    assert result["success"] is True
    assert result["accepted"] == 1
    assert "9999" in result["message"]


def test_nothing_to_apply_leaves_the_file_untouched(simple_path):
    original = simple_path.read_bytes()

    accepted = accept_tracked_changes_in_doc(str(simple_path))
    rejected = reject_tracked_changes_in_doc(str(simple_path))

    assert accepted == {
        "success": True,
        "message": "No matching tracked changes found to accept",
        "accepted": 0,
    }
    assert rejected == {
        "success": True,
        "message": "No matching tracked changes found to reject",
        "rejected": 0,
    }
    assert simple_path.read_bytes() == original


def test_accepting_everything_applies_the_paragraph_mark_and_names_what_it_skipped(
    tracked_path,
):
    """A deleted paragraph mark really merges, and the skipped kinds are named."""
    before = snapshot(tracked_path.read_bytes())
    n_before = len(before.story_paragraphs("document"))

    result = accept_tracked_changes_in_doc(str(tracked_path))

    assert result["success"] is True
    after = snapshot(tracked_path.read_bytes())
    assert len(after.story_paragraphs("document")) == n_before - 1, (
        "accepting the deleted paragraph mark must merge its paragraph into the next"
    )
    # rPrChange and pPrChange are not applied by this tool, and the message says so
    # rather than reporting a document that is fully applied when it is not.
    assert "rPrChange" in result["message"]
    assert "pPrChange" in result["message"]
    listed = list_tracked_changes_in_doc(str(tracked_path))
    assert listed["total_changes"] == 0
    assert validate_package(tracked_path) == []


def test_a_refused_selection_writes_nothing(tracked_path):
    """All-or-nothing: an inserted mark with nothing to merge into aborts the call.

    The last paragraph of the fixture carries an inserted paragraph mark and has
    no following paragraph, so rejecting it cannot undo the split it recorded.
    The engine refuses the whole selection, and the file must still be the one
    that was read.
    """
    original = tracked_path.read_bytes()

    with pytest.raises(UnsupportedRevision):
        reject_tracked_changes_in_doc(str(tracked_path))

    assert tracked_path.read_bytes() == original


def test_rejecting_a_recorded_replace_restores_the_document(combined_path):
    """The round trip the engine promises, driven from the tool layer."""
    before = snapshot(combined_path.read_bytes())

    recorded = track_replace_in_doc(
        str(combined_path), "Second heading", "Third heading", AUTHOR
    )
    assert recorded["success"] is True
    rejected = reject_tracked_changes_in_doc(str(combined_path), author=AUTHOR)
    assert rejected["success"] is True

    after = snapshot(combined_path.read_bytes())
    assert_unchanged_except(before, after)
    assert validate_package(combined_path) == []


def test_applying_a_selection_touches_no_other_part(combined_path):
    """Accepting by author in a rich document leaves tables and headers alone."""
    before = snapshot(combined_path.read_bytes())
    index = index_of(before, "Body text under the second heading.")
    track_delete_in_doc(str(combined_path), "Body text ", AUTHOR)
    middle = snapshot(combined_path.read_bytes())

    result = accept_tracked_changes_in_doc(str(combined_path), author=AUTHOR)

    assert result["success"] is True
    assert result["accepted"] == 1
    after = snapshot(combined_path.read_bytes())
    assert after.paragraphs[("document", index)].text == "under the second heading."
    assert_unchanged_except(middle, after, paragraphs=[index], counters=["revisions"])
    assert validate_package(combined_path) == []


# --------------------------------------------------------------------------
# Structural guards
# --------------------------------------------------------------------------


def imported_roots(target) -> set[str]:
    """Top-level names `target` imports, read from its source rather than its prose.

    A plain substring scan would match the module docstring, which names the
    very things this guard is about.
    """
    tree = ast.parse(inspect.getsource(target))
    roots: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            roots.update(alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            roots.add(node.module.split(".")[0])
    return roots


def test_the_module_no_longer_rewrites_the_archive_itself():
    """``zipfile``, ``lxml`` and the duplicated run helpers are gone, for good.

    The non-atomic writer this module used to carry rebuilt the whole archive
    member by member on every call; a crash mid-write left a truncated package.
    Saving now goes through ``DocxPackage.save``, and no adapter has any reason
    to reach for a zip, for ``lxml`` or for a run helper of its own again.
    """
    roots = imported_roots(module)
    # Guard the guard: an ``isdisjoint`` over an empty set passes for the wrong
    # reason, and so does a ``vars()`` lookup on a module that failed to load.
    assert "word_document_server" in roots
    assert roots.isdisjoint({"zipfile", "lxml", "io", "copy", "re"})
    duplicated = {
        "_generate_id",
        "_now_iso",
        "_get_run_text",
        "_get_run_rpr",
        "_make_run",
        "_load_document_xml",
        "_save_document_xml",
        "_get_paragraphs",
        "_paragraph_text",
        "_find_text_in_paragraph",
    }
    assert duplicated.isdisjoint(vars(module))
