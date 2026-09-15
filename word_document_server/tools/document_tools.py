"""
Document creation and manipulation tools for Word Document Server.
"""
import os
import json
from typing import Dict, List, Optional, Any
from docx import Document

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, create_document_copy
from word_document_server.utils.document_utils import get_document_properties, extract_document_text, get_document_structure, get_document_xml, insert_header_near_text, insert_line_or_paragraph_near_text
from word_document_server.core.styles import ensure_heading_style, ensure_table_style


async def create_document(filename: str, title: Optional[str] = None, author: Optional[str] = None) -> str:
    """Create a new Word document with optional metadata.
    
    Args:
        filename: Name of the document to create (with or without .docx extension)
        title: Optional title for the document metadata
        author: Optional author for the document metadata
    """
    filename = ensure_docx_extension(filename)
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot create document: {error_message}"
    
    try:
        doc = Document()
        
        # Set properties if provided
        if title:
            doc.core_properties.title = title
        if author:
            doc.core_properties.author = author
        
        # Ensure necessary styles exist
        ensure_heading_style(doc)
        ensure_table_style(doc)
        
        # Save the document
        doc.save(filename)
        
        return f"Document {filename} created successfully"
    except Exception as e:
        return f"Failed to create document: {str(e)}"


async def get_document_info(filename: str) -> str:
    """Get information about a Word document.
    
    Args:
        filename: Path to the Word document
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    try:
        properties = get_document_properties(filename)
        return json.dumps(properties, indent=2)
    except Exception as e:
        return f"Failed to get document info: {str(e)}"


async def get_document_text(filename: str, show_revisions: bool = False) -> str:
    """Extract all text from a Word document.

    By default returns the effective final text (insertions applied, deletions
    removed).  Set show_revisions=True to get inline redline markup where
    deletions appear as [-deleted-] and insertions as {+inserted+}.

    Args:
        filename: Path to the Word document
        show_revisions: If True, annotate tracked changes inline
    """
    filename = ensure_docx_extension(filename)

    return extract_document_text(filename, show_revisions=show_revisions)


async def get_document_outline(filename: str) -> str:
    """Get the structure of a Word document.
    
    Args:
        filename: Path to the Word document
    """
    filename = ensure_docx_extension(filename)
    
    structure = get_document_structure(filename)
    return json.dumps(structure, indent=2)


async def list_available_documents(directory: str = ".") -> str:
    """List all .docx files in the specified directory.
    
    Args:
        directory: Directory to search for Word documents
    """
    try:
        if not os.path.exists(directory):
            return f"Directory {directory} does not exist"
        
        docx_files = [f for f in os.listdir(directory) if f.endswith('.docx')]
        
        if not docx_files:
            return f"No Word documents found in {directory}"
        
        result = f"Found {len(docx_files)} Word documents in {directory}:\n"
        for file in docx_files:
            file_path = os.path.join(directory, file)
            size = os.path.getsize(file_path) / 1024  # KB
            result += f"- {file} ({size:.2f} KB)\n"
        
        return result
    except Exception as e:
        return f"Failed to list documents: {str(e)}"


async def copy_document(source_filename: str, destination_filename: Optional[str] = None) -> str:
    """Create a copy of a Word document.
    
    Args:
        source_filename: Path to the source document
        destination_filename: Optional path for the copy. If not provided, a default name will be generated.
    """
    source_filename = ensure_docx_extension(source_filename)
    
    if destination_filename:
        destination_filename = ensure_docx_extension(destination_filename)
    
    success, message, new_path = create_document_copy(source_filename, destination_filename)
    if success:
        return message
    else:
        return f"Failed to copy document: {message}"


async def merge_documents(target_filename: str, source_filenames: List[str], add_page_breaks: bool = True) -> str:
    """Merge multiple Word documents into a single document.

    The first source document becomes the merged document -- it keeps its own
    styles, numbering, sections, headers and footers -- and the body of every
    other source is appended to it, element by element, by
    :func:`~word_document_server.engine.merge.append_document`. Runs, direct
    formatting, images, hyperlinks, bookmarks, fields, tracked changes, content
    controls and table geometry are carried over as they are; styles and list
    definitions the merged document lacks are copied with them.

    A source whose content carries comments, footnotes or endnotes is refused
    rather than merged without them, and the target file is left untouched. The
    headers and footers of the appended sources are not imported; whatever the
    merge could not carry over faithfully is listed after the success message.

    Args:
        target_filename: Path to the target document (will be created or overwritten)
        source_filenames: List of paths to source documents to merge
        add_page_breaks: If True, add page breaks between documents
    """
    from word_document_server.engine.merge import append_document
    from word_document_server.engine.package import DocxPackage

    target_filename = ensure_docx_extension(target_filename)

    # Check if target file is writeable
    is_writeable, error_message = check_file_writeable(target_filename)
    if not is_writeable:
        return f"Cannot create target document: {error_message}"

    # Validate all source documents exist
    missing_files = []
    for filename in source_filenames:
        doc_filename = ensure_docx_extension(filename)
        if not os.path.exists(doc_filename):
            missing_files.append(doc_filename)

    if missing_files:
        return f"Cannot merge documents. The following source files do not exist: {', '.join(missing_files)}"

    try:
        if not source_filenames:
            Document().save(target_filename)
            return f"Successfully merged 0 documents into {target_filename}"

        paths = [ensure_docx_extension(filename) for filename in source_filenames]
        # The whole merge happens in memory and the file is written once, at the
        # end: a source that has to be refused leaves the target as it was.
        merged = DocxPackage.open(paths[0])
        warnings = []
        for path in paths[1:]:
            report = append_document(
                merged, DocxPackage.open(path), page_break=add_page_breaks
            )
            warnings.extend(f"{os.path.basename(path)}: {warning}" for warning in report.warnings)
        merged.save(target_filename)

        message = f"Successfully merged {len(source_filenames)} documents into {target_filename}"
        if warnings:
            message += "\nWarnings:\n" + "\n".join(f"- {warning}" for warning in warnings)
        return message
    except Exception as e:
        return f"Failed to merge documents: {str(e)}"


async def get_document_xml_tool(filename: str) -> str:
    """Get the raw XML structure of a Word document."""
    return get_document_xml(filename)
