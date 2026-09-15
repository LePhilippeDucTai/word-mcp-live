"""Regression tests for the three macOS live-tool crashes found by reading
the code (audit report, decision D-001 / J04-P4):

1. word_live_set_page_layout's macOS branch referenced undefined names
   (page_width, page_height, ...) instead of the actual *_inches params
   converted to points -> NameError on every call.
2. core/word_mac.py used `re.sub` without importing `re` at module level
   -> NameError as soon as a heading number gets stripped.
3. word_live_format_text passed an integer highlight_color straight to
   mac_format_text, which expects a JXA color name string -> AttributeError
   ("int has no attribute 'replace'") inside _escape_js.

_MAC_AVAILABLE is forced to True and _run_jxa (the only OS boundary these
paths cross) is stubbed to capture the generated JXA script instead of
shelling out to osascript.
"""

import asyncio
import json

from word_document_server.core import word_mac
from word_document_server.tools import live_layout_tools
from word_document_server.tools import live_tools


class _JXACapture:
    """Stand-in for word_mac._run_jxa: records every script it is given."""

    def __init__(self, return_value: str = '{"ok": true}'):
        self.scripts = []
        self.return_value = return_value

    def __call__(self, script: str, timeout: int = 30) -> str:
        self.scripts.append(script)
        return self.return_value


def test_word_live_set_page_layout_mac_converts_inches_to_points(monkeypatch):
    monkeypatch.setattr(live_layout_tools, "_MAC_AVAILABLE", True)
    capture = _JXACapture()
    monkeypatch.setattr(word_mac, "_run_jxa", capture)

    result = asyncio.run(
        live_layout_tools.word_live_set_page_layout(
            filename="doc.docx",
            section_index=1,
            orientation="landscape",
            page_width_inches=11.0,
            page_height_inches=8.5,
            margin_top_inches=1.0,
            margin_bottom_inches=1.0,
            margin_left_inches=0.5,
            margin_right_inches=0.5,
        )
    )

    # No exception was raised and the script actually reached _run_jxa.
    assert result == capture.return_value
    assert len(capture.scripts) == 1
    script = capture.scripts[0]

    # inches -> points conversion (72 pts/inch), same as the COM branch.
    assert "ps.pageWidth = 792.0;" in script
    assert "ps.pageHeight = 612.0;" in script
    assert "ps.topMargin = 72.0;" in script
    assert "ps.bottomMargin = 72.0;" in script
    assert "ps.leftMargin = 36.0;" in script
    assert "ps.rightMargin = 36.0;" in script
    assert 'ps.orientation = "orient landscape";' in script


def test_mac_setup_heading_numbering_strips_numeric_prefix_without_nameerror(monkeypatch):
    monkeypatch.setattr(
        word_mac,
        "_run_jxa",
        lambda script, timeout=30: json.dumps({"h1_applied": 1, "h2_applied": 0}),
    )
    monkeypatch.setattr(
        word_mac,
        "mac_get_text",
        lambda filename=None: json.dumps(
            {"paragraphs": [{"index": 0, "text": "1. Introduction"}]}
        ),
    )
    replace_calls = []

    def fake_replace_text(filename=None, find_text=None, replace_text=None,
                           match_case=None, replace_all=None):
        replace_calls.append({"find_text": find_text, "replace_text": replace_text})
        return json.dumps({"replaced": 1})

    monkeypatch.setattr(word_mac, "mac_replace_text", fake_replace_text)

    # This raises NameError: name 're' is not defined without the module-level
    # `import re` fix, since strip_manual_numbers exercises `re.sub` directly.
    result = json.loads(
        word_mac.mac_setup_heading_numbering(
            filename="doc.docx",
            h1_paragraphs=[1],
            strip_manual_numbers=True,
        )
    )

    assert result["stripped"] == 1
    assert replace_calls == [{"find_text": "1. Introduction", "replace_text": "Introduction"}]


def test_word_live_format_text_mac_converts_highlight_color_int_to_name(monkeypatch):
    monkeypatch.setattr(live_tools, "_MAC_AVAILABLE", True)
    capture = _JXACapture()
    monkeypatch.setattr(word_mac, "_run_jxa", capture)

    # highlight_color=7 mirrors the WdColorIndex documented in the tool's own
    # docstring (7 = yellow). Passing the raw int used to raise AttributeError
    # inside _escape_js (int has no attribute 'replace').
    result = asyncio.run(
        live_tools.word_live_format_text(
            filename="doc.docx",
            start=0,
            end=10,
            highlight_color=7,
        )
    )

    assert result == capture.return_value
    assert len(capture.scripts) == 1
    assert 'r.highlightColorIndex = "yellow";' in capture.scripts[0]
