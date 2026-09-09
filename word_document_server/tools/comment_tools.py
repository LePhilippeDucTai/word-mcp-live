"""
Comment extraction tools for Word Document Server.

These tools provide high-level interfaces for extracting and analyzing
comments from Word documents through the MCP protocol.

Every comment is reported with the story it is anchored in and the V2 index of
its anchor paragraph -- the rank of that paragraph among the top-level
paragraphs of its story, table cells and text boxes excluded.  A comment
anchored inside a table cell has no such index (``paragraph_index`` is null and
``in_table`` is true).
"""
import json
import os
from typing import Any, Optional

from word_document_server.core.comments import (
    extract_all_comments,
    filter_comments_by_author,
)
from word_document_server.core.comments import (
    get_comments_for_paragraph as filter_comments_for_paragraph,
)
from word_document_server.engine.find import iter_paragraphs
from word_document_server.engine.find import _v2_index_map as v2_index_map
from word_document_server.engine.package import MAIN_STORY, DocxPackage
from word_document_server.engine.textmodel import visible_text
from word_document_server.utils.file_utils import ensure_docx_extension


def _missing_document(filename: str) -> Optional[str]:
    """JSON error response when `filename` is not a readable document, else None."""
    if not os.path.exists(filename):
        return json.dumps(
            {"success": False, "error": f"Document {filename} does not exist"}, indent=2
        )
    return None


def _v2_paragraphs(pkg: DocxPackage) -> list[Any]:
    """The main story's paragraphs that have a V2 index, in index order."""
    paragraphs = iter_paragraphs(pkg.document)
    index_map = v2_index_map(paragraphs)
    return [p for p in paragraphs if p in index_map]


async def get_all_comments(filename: str) -> str:
    """
    Extract all comments from a Word document.

    Args:
        filename: Path to the Word document

    Returns:
        JSON string containing all comments with metadata
    """
    filename = ensure_docx_extension(filename)

    missing = _missing_document(filename)
    if missing is not None:
        return missing

    try:
        comments = extract_all_comments(DocxPackage.open(filename))
        return json.dumps(
            {"success": True, "comments": comments, "total_comments": len(comments)},
            indent=2,
        )
    except Exception as e:
        return json.dumps(
            {"success": False, "error": f"Failed to extract comments: {str(e)}"}, indent=2
        )


async def get_comments_by_author(filename: str, author: str) -> str:
    """
    Extract comments from a specific author in a Word document.

    Args:
        filename: Path to the Word document
        author: Name of the comment author to filter by

    Returns:
        JSON string containing filtered comments
    """
    filename = ensure_docx_extension(filename)

    missing = _missing_document(filename)
    if missing is not None:
        return missing

    if not author or not author.strip():
        return json.dumps({"success": False, "error": "Author name cannot be empty"}, indent=2)

    try:
        comments = filter_comments_by_author(
            extract_all_comments(DocxPackage.open(filename)), author
        )
        return json.dumps(
            {
                "success": True,
                "author": author,
                "comments": comments,
                "total_comments": len(comments),
            },
            indent=2,
        )
    except Exception as e:
        return json.dumps(
            {"success": False, "error": f"Failed to extract comments: {str(e)}"}, indent=2
        )


async def get_comments_for_paragraph(filename: str, paragraph_index: int) -> str:
    """
    Extract comments for a specific paragraph in a Word document.

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph (0-based, V2 index space)

    Returns:
        JSON string containing comments for the specified paragraph
    """
    filename = ensure_docx_extension(filename)

    missing = _missing_document(filename)
    if missing is not None:
        return missing

    if paragraph_index < 0:
        return json.dumps(
            {"success": False, "error": "Paragraph index must be non-negative"}, indent=2
        )

    try:
        pkg = DocxPackage.open(filename)
        paragraphs = _v2_paragraphs(pkg)
        if paragraph_index >= len(paragraphs):
            return json.dumps(
                {
                    "success": False,
                    "error": f"Paragraph index {paragraph_index} is out of range. "
                    f"Document has {len(paragraphs)} paragraphs.",
                },
                indent=2,
            )

        comments = filter_comments_for_paragraph(
            extract_all_comments(pkg), paragraph_index, MAIN_STORY
        )
        return json.dumps(
            {
                "success": True,
                "paragraph_index": paragraph_index,
                "paragraph_text": visible_text(paragraphs[paragraph_index]),
                "comments": comments,
                "total_comments": len(comments),
            },
            indent=2,
        )
    except Exception as e:
        return json.dumps(
            {"success": False, "error": f"Failed to extract comments: {str(e)}"}, indent=2
        )
