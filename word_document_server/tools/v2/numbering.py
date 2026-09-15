"""The V2 list tool: turn located paragraphs into the items of one list.

A bulleted or numbered paragraph is not a paragraph whose text starts with a
bullet: it is a paragraph pointing at a definition in ``word/numbering.xml``,
which is what makes Word renumber the whole list when an item is inserted or
moved.  :func:`doc_apply_list` writes that pointer, and it writes a *new*
definition every call -- see
:mod:`word_document_server.engine.numbering` for why reusing an id the document
already has silently joins a list the author made somewhere else.

One call, one list: every paragraph named by `locators` joins the same
definition and they count together, in document order.  Two calls produce two
lists that count independently.
"""

from __future__ import annotations

from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine import numbering
from word_document_server.engine.locators import Target, resolve
from word_document_server.engine.numbering import ListKind
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.tools.v2.registry import ToolSpec

__all__ = ["TOOLS", "doc_apply_list"]


def _targets(pkg: DocxPackage, locators: list[dict[str, Any]]) -> list[Target]:
    """Resolve every locator, before anything is written.

    A locator that does not resolve fails the whole call with the engine's own
    code (``not_found``, ``ambiguous``, ``stale_anchor``): a list half applied is
    worse than a list not applied.

    Raises:
        ValueError: if `locators` is not a non-empty list of locators.
    """
    if isinstance(locators, dict) or not isinstance(locators, list):
        raise ValueError("'locators' must be a list of locators, even for a single paragraph")
    if not locators:
        raise ValueError("'locators' must name at least one paragraph")
    return [resolve(pkg, locator) for locator in locators]


def doc_apply_list(
    filename: str,
    locators: list[dict[str, Any]],
    kind: ListKind = "bullet",
    level: int = 0,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Make the paragraphs named by `locators` the items of one list.

    A new list definition is created for the call, so the items count on their
    own: they never continue a list the document already had, and a second call
    starts a second list. The paragraphs keep their style, their text and their
    formatting; only their list membership changes.

    Args:
        filename: path to the .docx file.
        locators: where to act, one locator per paragraph, in the order they
            should be numbered. Each is one of `{"paragraph": i}`, `{"find": s,
            "occurrence": n}`, `{"bookmark": name}`, `{"heading": s}`,
            `{"table": t, "row": r, "col": c, "paragraph": k}`, with the optional
            keys `story` and `expect_text` -- the same forms `doc_edit_text`
            takes.
        kind: `bullet` for glyphs, `decimal` for one counter per level (1., a.,
            i.), `multilevel` for an outline whose label spells the whole path
            (1., 1.1., 1.1.1.).
        level: the indent level of the items, 0 for the outermost. Every level
            up to 8 is defined, so a later call can nest under these ones by
            naming a deeper level of the same list.
        dry_run: compute and report the change without writing the file.

    Returns:
        The usual report, plus `num_id`, the id of the list that was created --
        pass it back nowhere, it is there to tell two lists apart in a report --
        and `saved`. A list label is not paragraph text, so `before` and `after`
        in `changes` are equal; they confirm which paragraphs were reached.
    """
    pkg = DocxPackage.open(filename)
    targets = _targets(pkg, locators)

    warnings: list[str] = []
    unique: list[Target] = []
    seen: set[int] = set()
    for target in targets:
        if id(target.paragraph) in seen:
            warnings.append(
                f"two locators named the same paragraph ({target.story}, "
                f"{target.index}); it joins the list once"
            )
            continue
        seen.add(id(target.paragraph))
        unique.append(target)

    for target in unique:
        current = numbering.paragraph_list_info(target.paragraph)
        if current is not None and current["num_id"] is not None:
            warnings.append(
                f"paragraph {target.index} of {target.story} already belonged to the "
                f"list numId {current['num_id']}; it leaves it for the new one"
            )

    num_id = numbering.create_list_definition(pkg, kind, levels=numbering.MAX_LEVELS)
    changes = []
    for target in unique:
        text = visible_text(target.paragraph)
        numbering.apply_list(target.paragraph, num_id, level)
        changes.append(
            {
                "story": target.story,
                "paragraph": target.index,
                "before": text,
                "after": visible_text(target.paragraph),
            }
        )

    saved = not dry_run
    if saved:
        pkg.save(filename)
    return {
        "dry_run": dry_run,
        "saved": saved,
        "num_id": num_id,
        "changes": changes,
        "warnings": warnings,
    }


TOOLS = [
    ToolSpec(
        fn=doc_apply_list,
        annotations=ToolAnnotations(title="Apply List Numbering", destructiveHint=True),
        tags=frozenset({"v2", "write"}),
    ),
]
