"""
Hyperlink tools for Word Document Server.

These tools provide MCP interfaces for managing hyperlinks in Word documents.
"""

import json
import os
from typing import Optional

from word_document_server.core.hyperlink_writer import (
    add_hyperlink_to_doc,
    list_hyperlinks_in_doc,
    remove_hyperlink_from_doc,
)
from word_document_server.utils.file_utils import (
    check_file_writeable,
    ensure_docx_extension,
    get_file_lock,
)

#: Actions ``manage_hyperlinks`` understands.  ``add`` is the default and the
#: only one the tool had before; the other two were added without changing it.
ACTIONS = ("add", "remove", "list")


async def manage_hyperlinks(
    filename: str,
    action: str = "add",
    text: str = "",
    url: str = "",
    paragraph_index: Optional[int] = None,
) -> str:
    """Add, remove or list hyperlinks in a Word document.

    Args:
        filename: Path to Word document
        action: "add" to link text, "remove" to unlink it, "list" to report every link
        text: Text to convert to a hyperlink ("add"), or text of the links to
            remove ("remove")
        url: URL the hyperlink should point to (for "add")
        paragraph_index: If given, only that paragraph is searched (0-based,
            counting the top-level paragraphs of the body; a paragraph inside a
            table cell or a text box has no such index)

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    if action == "list":
        try:
            return json.dumps(
                list_hyperlinks_in_doc(filename), ensure_ascii=False, indent=2
            )
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to list hyperlinks: {str(e)}"})

    if action not in ACTIONS:
        return json.dumps(
            {
                "success": False,
                "error": f"Unknown action: {action}. Supported: "
                + ", ".join(repr(name) for name in ACTIONS),
            }
        )

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if action == "add":
        if not text:
            return json.dumps({"success": False, "error": "text cannot be empty"})
        if not url:
            return json.dumps({"success": False, "error": "url cannot be empty"})

        try:
            async with get_file_lock(filename):
                result = add_hyperlink_to_doc(filename, text, url, paragraph_index)
            return json.dumps(result, ensure_ascii=False, indent=2)
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to add hyperlink: {str(e)}"})

    # action == "remove"
    if not text and paragraph_index is None:
        return json.dumps(
            {"success": False, "error": "text or paragraph_index is required to remove a hyperlink"}
        )
    try:
        async with get_file_lock(filename):
            result = remove_hyperlink_from_doc(filename, text, paragraph_index)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to remove hyperlink: {str(e)}"})
