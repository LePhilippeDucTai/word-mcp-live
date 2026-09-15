"""
Document utility functions for Word Document Server.
"""
import json
from dataclasses import dataclass
from typing import Dict, List, Any
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from lxml import etree

from word_document_server.engine.errors import EngineError, PackageError
from word_document_server.engine.find import _v2_index_map, iter_paragraphs
from word_document_server.engine.find import find as engine_find
from word_document_server.engine.format import PPR_ORDER
from word_document_server.engine.package import DocxPackage
from word_document_server.engine.ranges import insert_text, replace_range
from word_document_server.engine.textmodel import fields as paragraph_fields
from word_document_server.engine.textmodel import visible_text


def get_effective_text(paragraph) -> str:
    """Get the effective (final) text of a paragraph, correctly handling tracked changes.

    python-docx's paragraph.text only reads <w:r> elements that are direct
    children of <w:p>.  Tracked-change markup wraps runs in <w:ins> or <w:del>,
    so paragraph.text silently drops BOTH insertions AND deletions — producing
    broken text that is neither the original nor the accepted version.

    This function iterates all <w:t> elements in the paragraph XML.  Because
    inserted text uses <w:t> (inside <w:ins>/<w:r>) while deleted text uses
    <w:delText> (inside <w:del>/<w:r>), a simple iter('{…}t') naturally
    includes insertions and excludes deletions — giving the correct final text.

    For documents without tracked changes this returns the same result as
    paragraph.text (just slightly slower due to XML iteration).
    """
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W_T = f"{{{ns}}}t"
    W_MOVEFROM = f"{{{ns}}}moveFrom"

    texts = []
    for t_elem in paragraph._element.iter(W_T):
        # Skip text inside <w:moveFrom> (moved-away text should not appear)
        skip = False
        ancestor = t_elem.getparent()
        while ancestor is not None and ancestor is not paragraph._element:
            if ancestor.tag == W_MOVEFROM:
                skip = True
                break
            ancestor = ancestor.getparent()
        if not skip and t_elem.text:
            texts.append(t_elem.text)
    return "".join(texts)


def get_redline_text(paragraph) -> str:
    """Get paragraph text with tracked changes annotated inline.

    Returns text with deletions marked as [-deleted-] and insertions as {+inserted+}.
    Normal (untracked) text appears as-is.  For documents without tracked changes
    this returns the same as paragraph.text / get_effective_text().
    """
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W = lambda tag: f"{{{ns}}}{tag}"

    parts = []
    for child in paragraph._element:
        tag = child.tag
        if tag == W("r"):
            # Normal run
            for t in child.iter(W("t")):
                if t.text:
                    parts.append(t.text)
        elif tag == W("ins"):
            # Insertion
            ins_texts = []
            for t in child.iter(W("t")):
                if t.text:
                    ins_texts.append(t.text)
            if ins_texts:
                parts.append("{+" + "".join(ins_texts) + "+}")
        elif tag == W("del"):
            # Deletion
            del_texts = []
            for dt in child.iter(W("delText")):
                if dt.text:
                    del_texts.append(dt.text)
            if not del_texts:
                for t in child.iter(W("t")):
                    if t.text:
                        del_texts.append(t.text)
            if del_texts:
                parts.append("[-" + "".join(del_texts) + "-]")
        elif tag == W("moveFrom"):
            # Moved-away text (like deletion)
            mf_texts = []
            for t in child.iter(W("t")):
                if t.text:
                    mf_texts.append(t.text)
            if mf_texts:
                parts.append("[-" + "".join(mf_texts) + "-]")
        elif tag == W("moveTo"):
            # Moved-to text (like insertion)
            mt_texts = []
            for t in child.iter(W("t")):
                if t.text:
                    mt_texts.append(t.text)
            if mt_texts:
                parts.append("{+" + "".join(mt_texts) + "+}")
        # Other elements (bookmarks, comments, etc.) are skipped
    return "".join(parts)


def get_document_properties(doc_path: str) -> Dict[str, Any]:
    """Get properties of a Word document."""
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        core_props = doc.core_properties
        
        return {
            "title": core_props.title or "",
            "author": core_props.author or "",
            "subject": core_props.subject or "",
            "keywords": core_props.keywords or "",
            "created": str(core_props.created) if core_props.created else "",
            "modified": str(core_props.modified) if core_props.modified else "",
            "last_modified_by": core_props.last_modified_by or "",
            "revision": core_props.revision or 0,
            "page_count": len(doc.sections),
            "word_count": sum(len(get_effective_text(paragraph).split()) for paragraph in doc.paragraphs),
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables)
        }
    except Exception as e:
        return {"error": f"Failed to get document properties: {str(e)}"}


def extract_document_text(doc_path: str, show_revisions: bool = False) -> str:
    """Extract all text from a Word document.

    Args:
        doc_path: Path to the .docx file.
        show_revisions: If True, annotate tracked changes inline with
            [-deleted-] and {+inserted+} markers.  If False (default),
            return the effective final text with insertions applied and
            deletions removed.

    Note:
        Uses get_effective_text / get_redline_text instead of paragraph.text
        so that tracked-change documents are read correctly.
    """
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"

    text_fn = get_redline_text if show_revisions else get_effective_text

    try:
        doc = Document(doc_path)
        text = []

        for paragraph in doc.paragraphs:
            text.append(text_fn(paragraph))

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text.append(text_fn(paragraph))

        return "\n".join(text)
    except Exception as e:
        return f"Failed to extract text: {str(e)}"


def get_document_structure(doc_path: str) -> Dict[str, Any]:
    """Get the structure of a Word document."""
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        structure = {
            "paragraphs": [],
            "tables": []
        }
        
        # Get paragraphs
        for i, para in enumerate(doc.paragraphs):
            structure["paragraphs"].append({
                "index": i,
                "text": get_effective_text(para)[:100] + ("..." if len(get_effective_text(para)) > 100 else ""),
                "style": para.style.name if para.style else "Normal"
            })
        
        # Get tables
        for i, table in enumerate(doc.tables):
            table_data = {
                "index": i,
                "rows": len(table.rows),
                "columns": len(table.columns),
                "preview": []
            }
            
            # Get sample of table data
            max_rows = min(3, len(table.rows))
            for row_idx in range(max_rows):
                row_data = []
                max_cols = min(3, len(table.columns))
                for col_idx in range(max_cols):
                    try:
                        cell_text = table.cell(row_idx, col_idx).text
                        row_data.append(cell_text[:20] + ("..." if len(cell_text) > 20 else ""))
                    except IndexError:
                        row_data.append("N/A")
                table_data["preview"].append(row_data)
            
            structure["tables"].append(table_data)
        
        return structure
    except Exception as e:
        return {"error": f"Failed to get document structure: {str(e)}"}


def find_paragraph_by_text(doc, text, partial_match=False):
    """
    Find paragraphs containing specific text.
    
    Args:
        doc: Document object
        text: Text to search for
        partial_match: If True, matches paragraphs containing the text; if False, matches exact text
        
    Returns:
        List of paragraph indices that match the criteria
    """
    matching_paragraphs = []
    
    for i, para in enumerate(doc.paragraphs):
        para_text = get_effective_text(para)
        if partial_match and text in para_text:
            matching_paragraphs.append(i)
        elif not partial_match and para_text == text:
            matching_paragraphs.append(i)
            
    return matching_paragraphs


#: Prefix of the style names whose paragraphs a replacement never touches: a
#: table of contents is a generated view of the document, so rewriting it is
#: both pointless (Word regenerates it) and destructive (its entries carry the
#: field and hyperlink machinery that drives navigation).
TOC_STYLE_PREFIX = "TOC"


@dataclass(frozen=True)
class ReplacementReport:
    """Outcome of a search and replace over a whole package.

    Attributes:
        replaced: number of occurrences actually rewritten.
        skipped: number of occurrences left alone because they overlap a field.
            A field's markers, instruction and cached result are consistent only
            as a whole, so replacing text that reaches into one would corrupt it;
            such an occurrence is reported rather than silently dropped.
    """

    replaced: int
    skipped: int


def _style_names(pkg: DocxPackage) -> Dict[str, str]:
    """Map ``w:styleId`` to the style's ``w:name`` for every style of the package.

    Returns an empty map when the package has no live ``word/styles.xml``; a
    document without a styles part has no named style to skip.
    """
    part = pkg.find_part("word/styles.xml")
    if part is None:
        return {}
    try:
        root = pkg.root_of(part)
    except PackageError:
        return {}
    names: Dict[str, str] = {}
    for style in root.findall(qn('w:style')):
        style_id = style.get(qn('w:styleId'))
        if not style_id:
            continue
        name = style.find(qn('w:name'))
        value = None if name is None else name.get(qn('w:val'))
        names[style_id] = value or style_id
    return names


def _is_toc_paragraph(paragraph, style_names: Dict[str, str]) -> bool:
    """Whether `paragraph` carries a style whose name starts with ``TOC``.

    The name is the one written in ``styles.xml``, which is what the previous
    python-docx implementation compared too. A ``w:pStyle`` pointing at a style
    the package does not declare falls back to the style id, so a dangling
    ``TOC1`` reference is still recognised as a table-of-contents paragraph.
    """
    ppr = paragraph.find(qn('w:pPr'))
    if ppr is None:
        return False
    pstyle = ppr.find(qn('w:pStyle'))
    if pstyle is None:
        return False
    style_id = pstyle.get(qn('w:val'))
    if not style_id:
        return False
    return style_names.get(style_id, style_id).startswith(TOC_STYLE_PREFIX)


def replace_text_everywhere(pkg: DocxPackage, old_text: str, new_text: str) -> ReplacementReport:
    """Replace every occurrence of `old_text` with `new_text` across `pkg`.

    Every story the package holds is searched -- body, headers, footers,
    footnotes, endnotes -- through
    :func:`word_document_server.engine.find.find`, so an occurrence is found
    wherever it reads: split across runs, inside a hyperlink, a tracked
    insertion, a content control or a table cell. Rewriting goes through
    :func:`word_document_server.engine.ranges.replace_range`, which splits runs
    on the range boundaries and keeps every container it empties, so no
    relationship, revision or content control is lost.

    Occurrences of one paragraph are rewritten right to left, so the offsets of
    the occurrences still to process stay valid.

    Two kinds of occurrence are left alone: those in a paragraph styled ``TOC*``
    (see :data:`TOC_STYLE_PREFIX`), which are not counted at all because a table
    of contents is a generated view, and those overlapping a field, which are
    counted in :attr:`ReplacementReport.skipped` so the caller can say so.

    Args:
        pkg: the open package; the change is made in memory, saving is the
            caller's business.
        old_text: text to find; must not be empty.
        new_text: replacement text, possibly empty (which deletes the match).

    Returns:
        A :class:`ReplacementReport`.

    Raises:
        ValueError: if `old_text` is empty.
        UnsupportedRange: if an occurrence cannot be rewritten for a reason
            other than a field -- it strictly contains an image, a note
            reference or tracked-deleted content, say. Nothing is saved by this
            function, so the package on disk is untouched.
    """
    story_ids = [story_id for story_id, _ in pkg.stories()]
    matches = engine_find(pkg, old_text, stories=story_ids)

    # Group by paragraph while keeping document order; lxml elements are hashable
    # by identity, which is exactly the grouping wanted here.
    grouped: Dict[Any, list] = {}
    for match in matches:
        grouped.setdefault(match.paragraph, []).append(match)

    style_names = _style_names(pkg)
    replaced = 0
    skipped = 0
    for paragraph, occurrences in grouped.items():
        if _is_toc_paragraph(paragraph, style_names):
            continue
        # Read once, before any mutation: these spans share the coordinate
        # system of the match offsets.
        field_spans = [(field.start, field.end) for field in paragraph_fields(paragraph)]
        for match in reversed(occurrences):
            if any(match.start < end and start < match.end for start, end in field_spans):
                skipped += 1
                continue
            replace_range(paragraph, match.start, match.end, new_text)
            replaced += 1
    return ReplacementReport(replaced=replaced, skipped=skipped)


def find_and_replace_text(doc, old_text, new_text):
    """
    Find and replace text throughout the document, skipping Table of Contents (TOC) paragraphs.

    Args:
        doc: An open :class:`~word_document_server.engine.package.DocxPackage`
        old_text: Text to find
        new_text: Text to replace with

    Returns:
        Number of replacements made -- one per occurrence, not one per run.
        Occurrences skipped because they overlap a field are not counted; call
        :func:`replace_text_everywhere` directly to get that count too.
    """
    if not isinstance(doc, DocxPackage):
        raise TypeError(
            "find_and_replace_text now works on the OOXML engine and takes a "
            f"DocxPackage, not {type(doc).__name__}; open the file with "
            "DocxPackage.open(path)"
        )
    return replace_text_everywhere(doc, old_text, new_text).replaced


def get_document_xml(doc_path: str) -> str:
    """Extract and return the raw XML structure of the Word document (word/document.xml)."""
    import os
    import zipfile
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        with zipfile.ZipFile(doc_path) as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                return xml_file.read().decode('utf-8')
    except Exception as e:
        return f"Failed to extract XML: {str(e)}"


def insert_header_near_text(doc_path: str, target_text: str = None, header_title: str = "", position: str = 'after', header_style: str = 'Heading 1', target_paragraph_index: int = None) -> str:
    """Insert a header (with specified style) before or after the target paragraph. Specify by text or paragraph index. Skips TOC paragraphs in text search.

    `target_paragraph_index` is read in the V2 index space -- the one
    ``find_text`` reports, see :func:`indexed_paragraphs` -- and the text search
    walks that same space, so the ``(index N)`` of the reply is counted where
    `target_paragraph_index` would be read.
    """
    import os
    from docx import Document
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        doc = Document(doc_path)
        # The V2 index space, the one find_text reports: python-docx's
        # doc.paragraphs counts the direct body children only.
        paragraphs = [Paragraph(pp, doc) for pp in indexed_paragraphs(doc.element.body)]
        found = False
        para = None
        if target_paragraph_index is not None:
            if target_paragraph_index < 0 or target_paragraph_index >= len(paragraphs):
                return f"Invalid target_paragraph_index: {target_paragraph_index}. Document has {len(paragraphs)} paragraphs."
            para = paragraphs[target_paragraph_index]
            found = True
        else:
            for i, p in enumerate(paragraphs):
                # Skip TOC paragraphs
                if p.style and p.style.name.lower().startswith("toc"):
                    continue
                if target_text and target_text in get_effective_text(p):
                    para = p
                    found = True
                    break
        if not found or para is None:
            return f"Target paragraph not found (by index or text). (TOC paragraphs are skipped in text search)"
        # Save anchor index before insertion
        if target_paragraph_index is not None:
            anchor_index = target_paragraph_index
        else:
            anchor_index = None
            for i, p in enumerate(paragraphs):
                if p is para:
                    anchor_index = i
                    break
        new_para = doc.add_paragraph(header_title, style=header_style)
        if position == 'before':
            para._element.addprevious(new_para._element)
        else:
            para._element.addnext(new_para._element)
        doc.save(doc_path)
        if anchor_index is not None:
            return f"Header '{header_title}' (style: {header_style}) inserted {position} paragraph (index {anchor_index})."
        else:
            return f"Header '{header_title}' (style: {header_style}) inserted {position} the target paragraph."
    except Exception as e:
        return f"Failed to insert header: {str(e)}"


def insert_line_or_paragraph_near_text(doc_path: str, target_text: str = None, line_text: str = "", position: str = 'after', line_style: str = None, target_paragraph_index: int = None) -> str:
    """
    Insert a new line or paragraph (with specified or matched style) before or after the target paragraph.
    You can specify the target by text (first match) or by paragraph index.
    Skips paragraphs whose style name starts with 'TOC' if using text search.

    `target_paragraph_index` is read in the V2 index space -- the one
    ``find_text`` reports, see :func:`indexed_paragraphs` -- and the text search
    walks that same space, so the ``(index N)`` of the reply is counted where
    `target_paragraph_index` would be read.
    """
    import os
    from docx import Document
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        doc = Document(doc_path)
        # The V2 index space, the one find_text reports: python-docx's
        # doc.paragraphs counts the direct body children only.
        paragraphs = [Paragraph(pp, doc) for pp in indexed_paragraphs(doc.element.body)]
        found = False
        para = None
        if target_paragraph_index is not None:
            if target_paragraph_index < 0 or target_paragraph_index >= len(paragraphs):
                return f"Invalid target_paragraph_index: {target_paragraph_index}. Document has {len(paragraphs)} paragraphs."
            para = paragraphs[target_paragraph_index]
            found = True
        else:
            for i, p in enumerate(paragraphs):
                # Skip TOC paragraphs
                if p.style and p.style.name.lower().startswith("toc"):
                    continue
                if target_text and target_text in get_effective_text(p):
                    para = p
                    found = True
                    break
        if not found or para is None:
            return f"Target paragraph not found (by index or text). (TOC paragraphs are skipped in text search)"
        # Save anchor index before insertion
        if target_paragraph_index is not None:
            anchor_index = target_paragraph_index
        else:
            anchor_index = None
            for i, p in enumerate(paragraphs):
                if p is para:
                    anchor_index = i
                    break
        # Determine style: use provided or match target
        style = line_style if line_style else para.style
        new_para = doc.add_paragraph(line_text, style=style)
        if position == 'before':
            para._element.addprevious(new_para._element)
        else:
            para._element.addnext(new_para._element)
        doc.save(doc_path)
        if anchor_index is not None:
            return f"Line/paragraph inserted {position} paragraph (index {anchor_index}) with style '{style}'."
        else:
            return f"Line/paragraph inserted {position} the target paragraph with style '{style}'."
    except Exception as e:
        return f"Failed to insert line/paragraph: {str(e)}"


def add_bullet_numbering(paragraph, num_id=1, level=0):
    """
    Add bullet/numbering XML to a paragraph.

    Args:
        paragraph: python-docx Paragraph object
        num_id: Numbering definition ID (1=bullets, 2=numbers, etc.)
        level: Indentation level (0=first level, 1=second level, etc.)

    Returns:
        The modified paragraph
    """
    # Get or create paragraph properties
    pPr = paragraph._element.get_or_add_pPr()

    # Remove existing numPr if any (to avoid duplicates)
    existing_numPr = pPr.find(qn('w:numPr'))
    if existing_numPr is not None:
        pPr.remove(existing_numPr)

    # Create numbering properties element
    numPr = OxmlElement('w:numPr')

    # Set indentation level
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)

    # Set numbering definition ID
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)

    # Add to paragraph properties
    pPr.append(numPr)

    return paragraph


def insert_numbered_list_near_text(doc_path: str, target_text: str = None, list_items: list = None, position: str = 'after', target_paragraph_index: int = None, bullet_type: str = 'bullet') -> str:
    """
    Insert a bulleted or numbered list before or after the target paragraph. Specify by text or paragraph index. Skips TOC paragraphs in text search.
    Args:
        doc_path: Path to the Word document
        target_text: Text to search for in paragraphs (optional if using index)
        list_items: List of strings, each as a list item
        position: 'before' or 'after' (default: 'after')
        target_paragraph_index: Optional paragraph index to use as anchor
        bullet_type: 'bullet' for bullets (•), 'number' for numbers (1,2,3) (default: 'bullet')
    Returns:
        Status message

    `target_paragraph_index` is read in the V2 index space -- the one
    ``find_text`` reports, see :func:`indexed_paragraphs` -- and the text search
    walks that same space, so the ``(index N)`` of the reply is counted where
    `target_paragraph_index` would be read.
    """
    import os
    from docx import Document
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        doc = Document(doc_path)
        # The V2 index space, the one find_text reports: python-docx's
        # doc.paragraphs counts the direct body children only.
        paragraphs = [Paragraph(pp, doc) for pp in indexed_paragraphs(doc.element.body)]
        found = False
        para = None
        if target_paragraph_index is not None:
            if target_paragraph_index < 0 or target_paragraph_index >= len(paragraphs):
                return f"Invalid target_paragraph_index: {target_paragraph_index}. Document has {len(paragraphs)} paragraphs."
            para = paragraphs[target_paragraph_index]
            found = True
        else:
            for i, p in enumerate(paragraphs):
                # Skip TOC paragraphs
                if p.style and p.style.name.lower().startswith("toc"):
                    continue
                if target_text and target_text in get_effective_text(p):
                    para = p
                    found = True
                    break
        if not found or para is None:
            return f"Target paragraph not found (by index or text). (TOC paragraphs are skipped in text search)"
        # Save anchor index before insertion
        if target_paragraph_index is not None:
            anchor_index = target_paragraph_index
        else:
            anchor_index = None
            for i, p in enumerate(paragraphs):
                if p is para:
                    anchor_index = i
                    break
        # Determine numbering ID based on bullet_type
        num_id = 1 if bullet_type == 'bullet' else 2

        # Use ListParagraph style for proper list formatting
        style_name = None
        for candidate in ['List Paragraph', 'ListParagraph', 'Normal']:
            try:
                _ = doc.styles[candidate]
                style_name = candidate
                break
            except KeyError:
                continue
        if not style_name:
            style_name = None  # fallback to default

        new_paras = []
        for item in (list_items or []):
            p = doc.add_paragraph(item, style=style_name)
            # Add bullet numbering XML - this is the fix!
            add_bullet_numbering(p, num_id=num_id, level=0)
            new_paras.append(p)
        # Move the new paragraphs to the correct position
        for p in reversed(new_paras):
            if position == 'before':
                para._element.addprevious(p._element)
            else:
                para._element.addnext(p._element)
        doc.save(doc_path)
        list_type = "bulleted" if bullet_type == 'bullet' else "numbered"
        if anchor_index is not None:
            return f"{list_type.capitalize()} list with {len(new_paras)} items inserted {position} paragraph (index {anchor_index})."
        else:
            return f"{list_type.capitalize()} list with {len(new_paras)} items inserted {position} the target paragraph."
    except Exception as e:
        return f"Failed to insert numbered list: {str(e)}"


def is_toc_paragraph(para):
    """Devuelve True si el párrafo tiene un estilo de tabla de contenido (TOC)."""
    return para.style and para.style.name.upper().startswith("TOC")


def is_heading_paragraph(para):
    """Devuelve True si el párrafo tiene un estilo de encabezado (Heading 1, Heading 2, etc)."""
    return para.style and para.style.name.lower().startswith("heading")


# --------------------------------------------------------------------------
# Block operations on the body
# --------------------------------------------------------------------------
#
# Replacing "the block under a header" or "the block between two anchors" is a
# body-level operation: what sits between two anchors is not a list of
# paragraphs but a stretch of *block children* -- paragraphs, tables and
# block-level content controls alike. The previous implementation walked
# ``doc.paragraphs`` only, so a table inside the block survived a replacement
# that was supposed to remove it, and it compared ``el.tag == CT_P.tag``, which
# reads an unbound lxml attribute descriptor and is never equal to a string, so
# the start anchor was never found at all.
#
# Two things are never removed here. The body's own final ``w:sectPr`` is not a
# block child and is skipped by :data:`BLOCK_TAGS`, so a block that runs to the
# end of the document can no longer take the page setup with it. And the
# ``w:sectPr`` a *paragraph* of the block carries -- a section break -- is moved
# onto one of the replacement paragraphs rather than dropped.

_W_BODY = qn("w:body")
_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_SDT = qn("w:sdt")
_W_PPR = qn("w:pPr")
_W_SECT_PR = qn("w:sectPr")
_W_PSTYLE = qn("w:pStyle")
_W_STYLE = qn("w:style")
_W_NAME = qn("w:name")
_W_VAL = qn("w:val")
_W_TYPE = qn("w:type")
_W_STYLE_ID = qn("w:styleId")
_W_DEFAULT = qn("w:default")

#: Children of ``w:body`` that carry block content, and are therefore what a
#: block-level operation may remove. Everything else a body may hold -- first
#: and foremost its final ``w:sectPr`` -- is never part of a block.
BLOCK_TAGS = frozenset({_W_P, _W_TBL, _W_SDT})

#: Paragraph style names that close a block: a heading, in English or in
#: Spanish, or a table-of-contents entry. Same set the previous implementation
#: used, read from the style *name* rather than from ``python-docx``.
_BLOCK_END_STYLE_PREFIXES = ("heading", "título", "titulo", "toc")

#: Style a replacement paragraph gets when the caller names none.
DEFAULT_PARAGRAPH_STYLE = "Normal"

_PPR_RANK = {qn(tag): rank for rank, tag in enumerate(PPR_ORDER)}


def get_paragraph_style(el):
    """Return the ``w:styleId`` a ``w:p`` element references, or ``None``.

    Attribute names are stored in Clark notation on an lxml element, so the
    prefixed form has to be resolved before the lookup -- testing
    ``'w:val' in pStyle.attrib`` never matches and made this helper return
    ``None`` for every styled paragraph.
    """
    properties = el.find(_W_PPR)
    if properties is None:
        return None
    style = properties.find(_W_PSTYLE)
    if style is None:
        return None
    return style.get(_W_VAL)


def body_element(pkg):
    """Return the ``w:body`` of a package's main document part.

    Raises:
        PackageError: if the main document part holds no body.
    """
    body = pkg.document.find(_W_BODY)
    if body is None:
        raise PackageError("the main document part has no w:body")
    return body


def body_paragraphs(body):
    """Return the ``w:p`` children of `body`, in document order.

    This is python-docx's own ``Document.paragraphs`` index space -- direct
    children of the body, so neither table cells nor content-control content --
    which is the space the ``paragraph_index`` argument of the public tools has
    always addressed. It is *not* the V2 locator space of
    :mod:`word_document_server.engine.find`.
    """
    return [child for child in body if child.tag == _W_P]


def indexed_paragraphs(story_root):
    """Return the paragraphs of `story_root` in V2 index order.

    ``indexed_paragraphs(root)[i]`` is the paragraph that
    :func:`~word_document_server.utils.extended_document_utils.find_text`
    reports at ``paragraph_index`` ``i``: every ``w:p`` of the story in document
    order, the content of a block ``w:sdt`` included, table cells and text boxes
    excluded.  This is the space the ``paragraph_index`` argument of the public
    tools addresses, and it is wider than :func:`body_paragraphs` -- a paragraph
    sitting inside a content control has a body index of its own here but is not
    a child of the body, so a caller that removes one must go through its actual
    parent.

    The membership test is :func:`~word_document_server.engine.find._v2_index_map`
    itself, reused rather than restated so the reading tools and the writing
    tools cannot drift into two different index spaces.  It numbers in ascending
    order, so the insertion order of the map it returns is the index order.
    """
    return list(_v2_index_map(iter_paragraphs(story_root)))


def block_children(body):
    """Return the block children of `body`: paragraphs, tables and controls."""
    return [child for child in body if child.tag in BLOCK_TAGS]


def styles_root(pkg):
    """Root of ``word/styles.xml``, or ``None`` when the package has none."""
    part = pkg.find_part("/word/styles.xml")
    return None if part is None else pkg.root_of(part)


def paragraph_style_names(root):
    """Map every paragraph ``w:styleId`` of a styles part to its ``w:name``.

    A style with no ``w:type`` counts as a paragraph style, which is what the
    schema's default says.
    """
    names = {}
    if root is None:
        return names
    for style in root.iter(_W_STYLE):
        if style.get(_W_TYPE) not in (None, "paragraph"):
            continue
        style_id = style.get(_W_STYLE_ID)
        if not style_id:
            continue
        title = style.find(_W_NAME)
        names[style_id] = (title.get(_W_VAL) if title is not None else None) or style_id
    return names


def paragraph_style_label(element, style_names):
    """Name of the paragraph style of a ``w:p``, or ``None`` when it has none.

    Falls back to the raw ``w:styleId`` when the styles part does not define it,
    so a caller comparing against "Heading 1" still recognises a paragraph whose
    style is missing from ``styles.xml``.
    """
    style_id = get_paragraph_style(element)
    if style_id is None:
        return None
    return style_names.get(style_id, style_id)


def attach_ppr_child(paragraph, child):
    """Move `child` into the paragraph's ``w:pPr``, at its rank in the schema.

    ``w:pPr`` is an ``xsd:sequence``: appending to it is how a document ends up
    with ``w:sectPr`` before ``w:rPr`` and Word offering to repair it. The
    element is inserted before the first child that outranks it, and ``w:pPr``
    itself is created at the front of the paragraph when it is absent.
    """
    properties = paragraph.find(_W_PPR)
    if properties is None:
        properties = etree.SubElement(paragraph, _W_PPR)
        paragraph.insert(0, properties)
    rank = _PPR_RANK[child.tag]
    for index, existing in enumerate(properties):
        other = _PPR_RANK.get(existing.tag)
        if other is not None and other > rank:
            properties.insert(index, child)
            return
    properties.append(child)


def _ends_block(element, style_names):
    """Whether `element` is the heading or TOC entry that closes a block."""
    if element.tag != _W_P:
        return False
    label = paragraph_style_label(element, style_names)
    return label is not None and label.lower().startswith(_BLOCK_END_STYLE_PREFIXES)


def _is_toc_entry(element, style_names):
    """Whether `element` is a paragraph carrying a table-of-contents style."""
    label = paragraph_style_label(element, style_names)
    return label is not None and label.upper().startswith("TOC")


def _resolve_paragraph_style(root, name):
    """Resolve a paragraph style *name* to the ``w:styleId`` to write.

    Returns ``(style_id, True)`` when the document defines the style --
    ``style_id`` being ``None`` for the default paragraph style, which carries
    no ``w:pStyle`` at all, exactly as ``python-docx`` writes it -- and
    ``(None, False)`` when it defines no such style. Both the id and the name
    are accepted, case-insensitively: Word stores the built-in headings under
    the name ``"heading 1"`` while every caller spells them ``"Heading 1"``.
    """
    if root is None:
        # No styles part to resolve against: only the conventional default is
        # safe to assume; any other name is written as an id of its own.
        return (None, True) if name == DEFAULT_PARAGRAPH_STYLE else (name, True)
    wanted = name.strip().lower()
    for style in root.iter(_W_STYLE):
        if style.get(_W_TYPE) not in (None, "paragraph"):
            continue
        style_id = style.get(_W_STYLE_ID)
        title = style.find(_W_NAME)
        label = title.get(_W_VAL) if title is not None else None
        candidates = {value.strip().lower() for value in (style_id, label) if value}
        if wanted not in candidates:
            continue
        if style.get(_W_DEFAULT) in ("1", "true", "on"):
            return None, True
        return style_id, True
    return None, False


def _new_paragraph(body, text, style_id):
    """Build a paragraph holding `text`, appended at the end of `body`.

    The caller moves it where it belongs; building it in the tree rather than
    detached is what makes lxml reuse the ``w:`` prefix already declared on the
    document root instead of emitting a redundant namespace declaration.
    """
    paragraph = etree.SubElement(body, _W_P)
    if style_id:
        properties = etree.SubElement(paragraph, _W_PPR)
        etree.SubElement(properties, _W_PSTYLE).set(_W_VAL, style_id)
    insert_text(paragraph, 0, text)
    return paragraph


def _section_properties_in(elements):
    """Every ``w:sectPr`` a *block-level* paragraph of `elements` carries.

    A ``w:sectPr`` in a ``w:pPr`` is a section break, and the schema only allows
    one where a section can actually end: on a paragraph of the body, directly
    or inside a block-level content control. One found in a table cell is
    something Word ignores, and hoisting it into the body would *create* a
    section break the document never had -- so it goes with the table.
    """
    found = []
    for element in elements:
        if element.tag == _W_P:
            paragraphs = [element]
        elif element.tag == _W_SDT:
            paragraphs = list(element.iter(_W_P))
        else:
            paragraphs = []
        for paragraph in paragraphs:
            properties = paragraph.find(_W_PPR)
            if properties is not None:
                found.extend(properties.findall(_W_SECT_PR))
    return found


def _carries_section(paragraph):
    """Whether `paragraph` already ends a section of its own."""
    return paragraph.find(f"{_W_PPR}/{_W_SECT_PR}") is not None


def _anchor_is_free(anchor):
    """Whether the anchor can take a section break on top of its own content."""
    return anchor.tag == _W_P and not _carries_section(anchor)


def _section_capacity(anchor, replacement_count):
    """How many section breaks the replacement paragraphs can carry.

    A ``w:pPr`` holds at most one ``w:sectPr``, so each paragraph carries one
    break: the replacements, plus the anchor when it does not already end a
    section of its own.
    """
    return replacement_count + (1 if _anchor_is_free(anchor) else 0)


def _section_carriers(anchor, created, count):
    """The `count` paragraphs that will carry the block's section breaks.

    Sections keep both their order and their extent: the last break of the block
    must still fall on the last paragraph before whatever followed the block, so
    the breaks land on the *last* paragraphs available. The anchor is therefore
    only used when there are more breaks than replacement paragraphs.
    """
    free = list(created)
    if _anchor_is_free(anchor):
        free.insert(0, anchor)
    return free[len(free) - count :] if count else []


def _replace_block(body, anchor, removable, new_paragraphs, style_id):
    """Replace `removable` with paragraphs holding `new_paragraphs`.

    `anchor` is the block the replacement is written after; it is never removed.
    Returns the number of block elements removed, or a refusal message when the
    block holds more section breaks than there are paragraphs left to carry
    them -- dropping one would silently change the page setup of everything
    that follows.
    """
    carried = _section_properties_in(removable)
    if len(carried) > _section_capacity(anchor, len(new_paragraphs)):
        return (
            f"Refusing to replace the block: it holds {len(carried)} section "
            f"break(s) and only {len(new_paragraphs)} replacement paragraph(s) "
            "to carry them."
        )
    for section in carried:
        section.getparent().remove(section)
    for element in removable:
        body.remove(element)

    cursor = anchor
    created = []
    for text in new_paragraphs:
        paragraph = _new_paragraph(body, text, style_id)
        cursor.addnext(paragraph)
        cursor = paragraph
        created.append(paragraph)

    for section, target in zip(carried, _section_carriers(anchor, created, len(carried))):
        attach_ppr_child(target, section)
    return len(removable)


def _find_header(blocks, header_text, style_names):
    """Index of the block whose visible text is `header_text`, or ``None``.

    Case and surrounding whitespace are ignored, and a paragraph carrying a
    table-of-contents style is never a header: it is the entry *pointing* at
    one, and matching it would replace the table of contents instead of the section.
    """
    wanted = header_text.strip().lower()
    for index, element in enumerate(blocks):
        if element.tag != _W_P or _is_toc_entry(element, style_names):
            continue
        if visible_text(element).strip().lower() == wanted:
            return index
    return None


def _block_end(blocks, start, style_names):
    """Index of the first block after `start` closing it, or ``len(blocks)``."""
    return next(
        (
            index
            for index in range(start + 1, len(blocks))
            if _ends_block(blocks[index], style_names)
        ),
        len(blocks),
    )


def delete_block_under_header(doc, header_text):
    """Remove everything under a header, up to the next heading or TOC entry.

    Kept for callers holding an open ``python-docx`` document: it mutates `doc`
    in place and does not save. Like the rest of this section it now works on
    the block children of the body, so a table under the header is removed with
    the paragraphs instead of being left behind, and a section break inside the
    block moves onto the header rather than disappearing.

    Returns:
        ``(header element, blocks removed)``, or ``(None, 0)`` when no paragraph
        of the body carries that text.
    """
    body = doc.element.body
    style_names = paragraph_style_names(doc.styles.element)
    blocks = block_children(body)
    position = _find_header(blocks, header_text, style_names)
    if position is None:
        return None, 0
    end = _block_end(blocks, position, style_names)
    header = blocks[position]
    outcome = _replace_block(body, header, blocks[position + 1 : end], [], None)
    return header, 0 if isinstance(outcome, str) else outcome


def replace_paragraph_block_below_header(
    doc_path: str,
    header_text: str,
    new_paragraphs: list,
    detect_block_end_fn=None,
    new_paragraph_style: str = None,
) -> str:
    """Replace everything under a header, up to the next heading or TOC entry.

    The header is matched on the visible text of a body paragraph, ignoring case
    and surrounding whitespace, skipping paragraphs whose style is a
    table-of-contents entry. Everything between it and the next heading is
    removed -- tables and block-level content controls included, which the
    previous implementation left behind -- and replaced by one paragraph per
    entry of `new_paragraphs`. The document is written once, at the end.

    Args:
        doc_path: Path to the Word document.
        header_text: Visible text of the header paragraph to replace under.
        new_paragraphs: Text of each replacement paragraph.
        detect_block_end_fn: Accepted for backwards compatibility and ignored,
            as it has always been; the end of the block is the next heading.
        new_paragraph_style: Paragraph style name for the replacements.

    Returns:
        A status message; a message naming what went wrong is returned rather
        than raised.
    """
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} not found."

    new_paragraphs = list(new_paragraphs or [])
    style_to_use = new_paragraph_style or DEFAULT_PARAGRAPH_STYLE
    try:
        pkg = DocxPackage.open(doc_path)
        body = body_element(pkg)
        root = styles_root(pkg)
        style_names = paragraph_style_names(root)
        style_id, known = _resolve_paragraph_style(root, style_to_use)
        if not known:
            return f"Style '{style_to_use}' not found in document."

        blocks = block_children(body)
        position = _find_header(blocks, header_text, style_names)
        if position is None:
            return f"Header '{header_text}' not found in document."

        end = _block_end(blocks, position, style_names)
        outcome = _replace_block(
            body, blocks[position], blocks[position + 1 : end], new_paragraphs, style_id
        )
        if isinstance(outcome, str):
            return outcome
        pkg.save(doc_path)
    except (EngineError, OSError) as exc:
        return f"Failed to replace the block under '{header_text}': {exc}"

    return (
        f"Replaced content under '{header_text}' with {len(new_paragraphs)} "
        f"paragraph(s), style: {style_to_use}, removed {outcome} elements."
    )


def replace_block_between_manual_anchors(
    doc_path: str,
    start_anchor_text: str,
    new_paragraphs: list,
    end_anchor_text: str = None,
    match_fn=None,
    new_paragraph_style: str = None,
) -> str:
    """Replace every block between two anchors of the body.

    The start anchor is a body paragraph whose visible text equals
    `start_anchor_text`. The end of the block is `end_anchor_text` when one is
    given, the next heading otherwise -- not the "visually distinct paragraph"
    guess of the previous implementation, which stopped at the first run
    carrying a bold, caps or size property and therefore at almost any emphasis.
    Paragraphs, tables and block-level content controls between the two are
    removed and replaced by one paragraph per entry of `new_paragraphs`; the
    body's final ``w:sectPr`` never is. The document is written once, at the end.

    Args:
        doc_path: Path to the Word document.
        start_anchor_text: Visible text of the paragraph the block starts after.
        new_paragraphs: Text of each replacement paragraph.
        end_anchor_text: Visible text of the paragraph the block stops at.
        match_fn: Optional predicate overriding the text comparison, called as
            ``match_fn(text, element)`` for the start anchor and
            ``match_fn(text, element, is_end=True)`` for the end one.
        new_paragraph_style: Paragraph style name for the replacements.

    Returns:
        A status message; a message naming what went wrong is returned rather
        than raised.
    """
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} not found."

    new_paragraphs = list(new_paragraphs or [])
    style_to_use = new_paragraph_style or DEFAULT_PARAGRAPH_STYLE
    predicate = match_fn if callable(match_fn) else None
    try:
        pkg = DocxPackage.open(doc_path)
        body = body_element(pkg)
        root = styles_root(pkg)
        style_names = paragraph_style_names(root)
        style_id, known = _resolve_paragraph_style(root, style_to_use)
        if not known:
            return f"Style '{style_to_use}' not found in document."

        blocks = block_children(body)
        start = None
        for index, element in enumerate(blocks):
            if element.tag != _W_P:
                continue
            text = visible_text(element).strip()
            if predicate is not None:
                if predicate(text, element):
                    start = index
                    break
            elif text == start_anchor_text.strip():
                start = index
                break
        if start is None:
            return f"Start anchor '{start_anchor_text}' not found."

        end = None
        if predicate is not None or end_anchor_text:
            for index in range(start + 1, len(blocks)):
                element = blocks[index]
                if element.tag != _W_P:
                    continue
                text = visible_text(element).strip()
                if predicate is not None:
                    if predicate(text, element, is_end=True):
                        end = index
                        break
                elif text == end_anchor_text.strip():
                    end = index
                    break
            if end is None and end_anchor_text:
                return f"End anchor '{end_anchor_text}' not found."
        if end is None:
            end = _block_end(blocks, start, style_names)

        outcome = _replace_block(
            body, blocks[start], blocks[start + 1 : end], new_paragraphs, style_id
        )
        if isinstance(outcome, str):
            return outcome
        pkg.save(doc_path)
    except (EngineError, OSError) as exc:
        return f"Failed to replace the block after '{start_anchor_text}': {exc}"

    return (
        f"Replaced content between '{start_anchor_text}' and "
        f"'{end_anchor_text or 'next logical header'}' with "
        f"{len(new_paragraphs)} paragraph(s), style: {style_to_use}, "
        f"removed {outcome} elements."
    )
