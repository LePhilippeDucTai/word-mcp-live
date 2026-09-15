"""Which platforms each MCP tool actually runs on.

:data:`PLATFORMS` is read by :mod:`scripts.gen_tools_md` (to label every row of
``TOOLS.md``) and by :func:`~word_document_server.tools.v2.capabilities.doc_capabilities`
(to count tool families). It is not read at tool-call time: nothing here stops
a call on the "wrong" platform, the tool itself does that (see the module
docstring of ``capabilities.py`` for why Word availability is only known to
live tools).

Three platform names are used throughout: ``"linux"``, ``"windows"``,
``"macos"``.

How the three tiers below were derived
---------------------------------------
Most tools (:data:`_DOCX_TOOLS`) read and write the ``.docx`` package directly
with python-docx / the engine layer: the file must be closed, but the code
that opens it runs identically on any OS, so these are cross-platform.

The ``word_live_*`` tools (plus ``word_screen_capture``) instead drive a
running Word application: COM on Windows, AppleScript/JXA on macOS. Each of
these wrappers branches on ``_MAC_AVAILABLE`` before falling back to the
Windows-only COM path (see ``tools/live_tools.py``, ``live_read_tools.py``,
``live_layout_tools.py``, ``screen_capture_tools.py``). A tool only counts as
macOS-capable here (:data:`_LIVE_MACOS_TOOLS`) when that branch calls a real
``core.word_mac.mac_*`` implementation; a tool whose macOS branch is a stub
error (``"... is not yet implemented on macOS"`` / "... does not expose this
feature") or that has no macOS branch at all is Windows-only
(:data:`_LIVE_WINDOWS_ONLY_TOOLS`), whatever its docstring's leading
``[Windows only]`` tag says -- that tag predates the incremental macOS
support and was never updated per tool, so it is not a reliable signal.

The ten ``doc_*`` V2 tools (:data:`_V2_TOOLS`) are python-docx-based like the
first tier, so they are cross-platform too.
"""

from __future__ import annotations

#: The three platform sets a tool can be marked with.
CROSS_PLATFORM: frozenset[str] = frozenset({"linux", "windows", "macos"})
WINDOWS_AND_MACOS: frozenset[str] = frozenset({"windows", "macos"})
WINDOWS_ONLY: frozenset[str] = frozenset({"windows"})

#: python-docx-based tools: the document must be closed, any OS.
_DOCX_TOOLS: tuple[str, ...] = (
    "create_document",
    "copy_document",
    "get_document_info",
    "get_document_text",
    "get_document_outline",
    "list_available_documents",
    "get_document_xml",
    "insert_header_near_text",
    "insert_line_or_paragraph_near_text",
    "insert_numbered_list_near_text",
    "add_paragraph",
    "add_heading",
    "add_picture",
    "add_table",
    "add_page_break",
    "delete_paragraph",
    "search_and_replace",
    "create_custom_style",
    "format_text",
    "format_table",
    "set_table_cell_shading",
    "apply_table_alternating_rows",
    "highlight_table_header",
    "merge_table_cells",
    "merge_table_cells_horizontal",
    "merge_table_cells_vertical",
    "set_table_cell_alignment",
    "set_table_alignment_all",
    "protect_document",
    "unprotect_document",
    "add_footnote_to_document",
    "add_footnote_after_text",
    "add_footnote_before_text",
    "add_footnote_enhanced",
    "add_endnote_to_document",
    "customize_footnote_style",
    "delete_footnote_from_document",
    "add_footnote_robust",
    "validate_document_footnotes",
    "delete_footnote_robust",
    "get_paragraph_text_from_document",
    "find_text_in_document",
    "get_highlighted_text",
    "convert_to_pdf",
    "replace_paragraph_block_below_header",
    "replace_block_between_manual_anchors",
    "get_all_comments",
    "get_comments_by_author",
    "get_comments_for_paragraph",
    "add_comment",
    "manage_hyperlinks",
    "set_table_column_width",
    "set_table_column_widths",
    "set_table_width",
    "auto_fit_table_columns",
    "format_table_cell_text",
    "set_table_cell_padding",
    "track_replace",
    "track_insert",
    "track_delete",
    "list_tracked_changes",
    "accept_tracked_changes",
    "reject_tracked_changes",
    "set_page_layout",
    "add_header_footer",
    "add_page_numbers",
    "add_section_break",
    "set_paragraph_spacing",
    "add_bookmark",
    "add_watermark",
    "add_table_of_contents",
    "merge_documents",
    "add_restricted_editing",
    "add_digital_signature",
    "verify_document",
)

#: Live tools (Word automation) whose macOS branch calls a real ``mac_*``
#: implementation -- see the module docstring for what "real" means here.
_LIVE_MACOS_TOOLS: tuple[str, ...] = (
    "word_screen_capture",
    "word_live_insert_text",
    "word_live_format_text",
    "word_live_replace_text",
    "word_live_add_table",
    "word_live_modify_table",
    "word_live_delete_text",
    "word_live_apply_list",
    "word_live_setup_heading_numbering",
    "word_live_get_text",
    "word_live_take_snapshot",
    "word_live_get_paragraph_format",
    "word_live_get_info",
    "word_live_list_open",
    "word_live_find_text",
    "word_live_get_comments",
    "word_live_add_comment",
    "word_live_delete_comment",
    "word_live_list_revisions",
    "word_live_accept_revisions",
    "word_live_reject_revisions",
    "word_live_get_page_text",
    "word_live_undo",
    "word_live_save",
    "word_live_toggle_track_changes",
    "word_live_diagnose_layout",
    "word_live_set_page_layout",
    "word_live_add_header_footer",
    "word_live_add_section_break",
    "word_live_set_paragraph_spacing",
    "word_live_add_bookmark",
)

#: Live tools with no working macOS branch: either a stub error (feature not
#: implemented / not exposed by the AppleScript dictionary) or no macOS branch
#: at all. Windows (COM) only.
_LIVE_WINDOWS_ONLY_TOOLS: tuple[str, ...] = (
    "word_live_insert_paragraphs",
    "word_live_format_table",
    "word_live_get_diff",
    "word_live_snapshot_status",
    "word_live_set_core_properties",
    "word_live_reply_to_comment",
    "word_live_resolve_comment",
    "word_live_get_undo_history",
    "word_live_insert_image",
    "word_live_insert_cross_reference",
    "word_live_list_cross_reference_items",
    "word_live_insert_equation",
    "word_live_add_page_numbers",
    "word_live_add_watermark",
)

#: The V2 semantic-engine tools (``doc_*``): python-docx-based, cross-platform.
_V2_TOOLS: tuple[str, ...] = (
    "doc_inspect",
    "doc_find",
    "doc_edit_text",
    "doc_format_range",
    "doc_apply_edits",
    "doc_capabilities",
    "doc_apply_list",
    "doc_list_styles",
    "doc_get_style",
    "doc_find_style_usage",
)

#: name -> platforms it runs on. See the module docstring for the tiers.
PLATFORMS: dict[str, frozenset[str]] = {
    **{name: CROSS_PLATFORM for name in _DOCX_TOOLS},
    **{name: WINDOWS_AND_MACOS for name in _LIVE_MACOS_TOOLS},
    **{name: WINDOWS_ONLY for name in _LIVE_WINDOWS_ONLY_TOOLS},
    **{name: CROSS_PLATFORM for name in _V2_TOOLS},
}
