"""Re-export of the canonical snapshot and structural comparison instrument.

The implementation lives in :mod:`word_document_server.engine.compare` -- it
moved there so that the ``doc_compare`` V2 tool
(:mod:`word_document_server.tools.v2.compare`) can build on it without the
production code importing from ``tests/``.  This module stays as a thin
re-export so that the fidelity harness used by every test from J01 onward
(``assert_unchanged_except``, :class:`Diff`, :class:`TableSignature`, ...)
keeps working under its original import path,
``from tests.support.snapshot import ...``, unchanged.

See :mod:`word_document_server.engine.compare` for the full documentation.
"""

from __future__ import annotations

from word_document_server.engine.compare import (
    MAIN_STORY,
    PARAGRAPH_INDEX_SPACE,
    Diff,
    FieldChange,
    ParagraphChange,
    ParagraphSignature,
    PartSignature,
    RunSignature,
    Snapshot,
    TableChange,
    TableSignature,
    assert_unchanged_except,
    diff,
    snapshot,
)

__all__ = [
    "MAIN_STORY",
    "PARAGRAPH_INDEX_SPACE",
    "Diff",
    "FieldChange",
    "ParagraphChange",
    "ParagraphSignature",
    "PartSignature",
    "RunSignature",
    "Snapshot",
    "TableChange",
    "TableSignature",
    "assert_unchanged_except",
    "diff",
    "snapshot",
]
