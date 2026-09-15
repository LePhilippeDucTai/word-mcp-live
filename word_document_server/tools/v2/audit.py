"""The V2 tool that says what is wrong with a document.

``doc_inspect`` describes a document; :func:`doc_audit` judges it.  The two are
the same read, asked two different questions: "what can I address here" and
"what will a human, or Word, trip over here".

An audit is the first call to make on a document somebody else wrote.  Every
finding carries the locator of the place it is about, so what it reports can be
handed straight to ``doc_format_range``, ``doc_edit_text`` or ``doc_set_style``
without locating anything twice.
"""

from __future__ import annotations

import dataclasses
from collections import Counter
from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine.audit import audit
from word_document_server.engine.package import DocxPackage
from word_document_server.tools.v2.registry import ToolSpec

__all__ = ["TOOLS", "doc_audit"]


def doc_audit(filename: str) -> dict[str, Any]:
    """Audit a Word document: what is broken in it, and what merely reads wrong.

    Read this before editing a document somebody else wrote. Every finding is
    `{"kind", "severity", "locator", "message", "data"}`: `kind` says which
    check spoke, `locator` is the address to act on (`null` for a finding about
    the package rather than about a paragraph, and for a paragraph in a text
    box, which no locator reaches), and `data` carries the numbers the message
    summarises. Branch on `kind` and `severity`, never on the wording.

    `severity` is what the finding *is*, not how urgent it is. `error` means the
    package contradicts itself -- a reference with no referent, a range with one
    end -- and Word repairs or drops those silently. `warning` means the
    document is valid and says something a reader will misread. `info` is an
    inventory.

    The twelve checks:

    - `heading_like_paragraph`: short, emphasised, and no heading style, so the
      navigation pane and every table of contents ignore it.
    - `direct_formatting_overrides_style`: the properties a paragraph sets
      directly against what its style already says, named and counted.
    - `mixed_fonts_in_paragraph`: runs that do not end up in the same font.
    - `consecutive_empty_paragraphs`: vertical spacing typed by hand.
    - `dangling_num_id`, `dangling_abstract_num_id`: numbering pointing at a
      definition the package does not carry.
    - `bookmark_without_end`, `bookmark_without_start`,
      `comment_range_without_end`, `comment_range_without_start`: ranges with
      one end.
    - `comment_without_anchor`: a comment nothing in the text points at.
    - `unused_custom_style`: declared by this document, applied nowhere.
    - `similar_style_names`: names differing only by case, spacing or
      punctuation, which read as one entry in a style gallery.
    - `revisions_by_author`: who left tracked changes, by kind.
    - `field_present`: which fields the document carries, and whether
      `w:updateFields` asks Word to refresh their cached results on open.

    Nothing is written and nothing is normalised: the report describes the
    document as it stands.

    Args:
        filename: path to the .docx file.

    Returns:
        The usual report, plus `audit`: `findings` in check order, and `counts`
        (`total`, `by_severity`, `by_kind`).
    """
    findings = audit(DocxPackage.open(filename))
    return {
        "audit": {
            "findings": [dataclasses.asdict(finding) for finding in findings],
            "counts": {
                "total": len(findings),
                "by_severity": dict(
                    sorted(Counter(finding.severity for finding in findings).items())
                ),
                "by_kind": dict(
                    sorted(Counter(finding.kind for finding in findings).items())
                ),
            },
        }
    }


TOOLS = [
    ToolSpec(
        fn=doc_audit,
        annotations=ToolAnnotations(title="Audit Document", readOnlyHint=True),
        tags=frozenset({"v2", "read"}),
    ),
]
