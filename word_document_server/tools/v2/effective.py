"""The V2 tool that answers "why does this look like this".

``doc_get_style`` says what a style declares and ``doc_format_range`` changes
what a run declares; neither tells an agent which of the two is responsible for
the bold it can see.  :func:`doc_get_effective_format` does: it resolves the
whole cascade -- document defaults, table style, paragraph style chain,
character style chain, direct formatting -- and reports, per property, the value
a reader sees *and* the layer that decided it.

That provenance is the point.  "Make this paragraph not bold" has three
different correct answers depending on where the bold lives, and two of them are
wrong in a way that only shows up eleven paragraphs later.
"""

from __future__ import annotations

import dataclasses
from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine.effective import effective_format
from word_document_server.engine.errors import LocatorError
from word_document_server.engine.locators import Target, resolve
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.tools.v2.registry import ToolSpec

__all__ = ["TOOLS", "doc_get_effective_format"]


def _respan(target: Target, start: int, end: int) -> Target:
    """`target` with the caller's offsets, checked against the paragraph.

    A :class:`~word_document_server.engine.locators.Target` is frozen, so the
    override is a replacement rather than a mutation -- and the paragraph it
    points at is the same live element either way.

    Raises:
        LocatorError: ``out-of-range`` if the offsets do not fit the paragraph.
            Reading past the end would silently answer about no run at all,
            which reads as "this text has no formatting".
    """
    if start == target.start and end == target.end:
        return target
    length = len(visible_text(target.paragraph))
    if start < 0 or end < start or end > length:
        raise LocatorError(
            "out-of-range",
            f"range {start}..{end} is outside the paragraph's 0..{length} visible text",
        )
    return dataclasses.replace(target, start=start, end=end)


def doc_get_effective_format(
    filename: str,
    locator: dict[str, Any],
    start: int | None = None,
    end: int | None = None,
) -> dict[str, Any]:
    """Report the formatting a reader sees on a range, and where each value comes from.

    Read this before changing formatting that a style might own. Every property
    comes back as `{"value", "source"}`, where `source` is the layer that
    decided it: `direct` (the run's own `w:rPr` or the paragraph's own `w:pPr`),
    `style:<id>` for a character or paragraph style anywhere in a `based_on`
    chain, `table_style:<id>`, `docDefaults`, or `theme` when the winning layer
    expressed the property as a theme reference rather than as a literal.

    Three values are answers, not data. `"unresolved"` means a layer sets the
    property and this version cannot compute what it sets -- today only a table
    style, whose contribution depends on conditional formatting that is not
    read. `"mixed"` means the range is not uniform; narrow it to see the detail.
    An absent property means nothing in the cascade sets it.

    Toggle properties (`bold`, `italic`, `caps`, `small_caps`, `strike`,
    `double_strike`, `outline`, `shadow`, `emboss`, `imprint`, `vanish`) do not
    override down the style chain, they XOR (ECMA-376 §17.7.3): a bold style
    based on a bold style renders *not* bold. When more than one style level
    stated a toggle, the entry carries a `sources` list naming all of them.
    Direct formatting is outside that rule and wins outright.

    A colour or a font resolved through the theme is a rendering hint: the
    arithmetic Word uses for `themeTint` and `themeShade` is fixed-point and
    this one is not, so a channel may land a unit or two away. Never write a
    resolved colour back into the document; change the reference instead.

    Args:
        filename: path to the .docx file.
        locator: where to look. The same forms as `doc_format_range`:
            `{"paragraph": i}`, `{"find": s, "occurrence": n}`,
            `{"bookmark": name}`, `{"heading": s}`,
            `{"table": t, "row": r, "col": c, "paragraph": k}`, with the
            optional keys `story` and `expect_text`.
        start: offset to read from, in the visible text of the located
            paragraph. Defaults to the start of the span the locator resolved
            to, which for `{"paragraph": i}` is the whole paragraph.
        end: offset to read up to, excluded. Defaults to the end of that span.

    Returns:
        The usual report, plus `effective`: `story`, `index`, `start` and `end`
        (where the answer applies), `paragraph_style` (the `w:pStyle`, or the
        style sheet's default paragraph style when the paragraph names none),
        `char_styles` (the `w:rStyle` ids the range's runs carry),
        `table_style`, and the two property maps `run` and `paragraph`.

        `indent`, `spacing` and `numbering` are inherited one setting at a time,
        so they nest one `{"value", "source"}` per key: `indent.left` may come
        from the style while `indent.hanging` comes from the paragraph.

        The properties of the numbering level itself are not merged in: the
        numbering *link* is reported, with its own provenance, and
        `doc_apply_list` is what acts on it.
    """
    pkg = DocxPackage.open(filename)
    target = resolve(pkg, locator)
    span_start = target.start if start is None else start
    span_end = target.end if end is None else end
    report = effective_format(pkg, _respan(target, span_start, span_end))
    return {"effective": report, "warnings": report["warnings"]}


TOOLS = [
    ToolSpec(
        fn=doc_get_effective_format,
        annotations=ToolAnnotations(title="Get Effective Format", readOnlyHint=True),
        tags=frozenset({"v2", "read", "styles"}),
    ),
]
