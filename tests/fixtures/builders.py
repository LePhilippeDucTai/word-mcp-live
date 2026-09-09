"""Deterministic OOXML fixture builders.

Every ``build_<name>()`` function returns the bytes of a ``.docx`` package built from
the python-docx default template, enriched with hand-written XML fragments injected
through lxml. No binary fixture is ever committed: the packages are rebuilt on demand.

Determinism
-----------
The template ships fixed core properties, so the only source of variation left in a
``Document.save()`` is the ZIP member timestamp. :func:`_finalize` rewrites the archive
with a frozen timestamp, which makes ``build_x() == build_x()`` byte for byte. Every
identifier the fragments carry (revision ids, ``rsid``, ``w14:paraId``, bookmark ids,
annotation ids, dates) is a literal, never a generated value.

Naming contract
---------------
:data:`ALL_FIXTURES` maps a fixture name to its builder. Downstream parts (J01-P5, J02,
J03) address fixtures by that name, so names are part of the public contract of this
module and must not be renamed lightly.
"""

from __future__ import annotations

import io
import re
import struct
import zipfile
import zlib
from collections.abc import Callable

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.oxml import serialize_part_xml
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.shared import Inches, Pt

# --------------------------------------------------------------------------------------
# Frozen values
# --------------------------------------------------------------------------------------

#: ZIP member timestamp used for every part, so archives are byte-for-byte reproducible.
FIXED_ZIP_DATE = (2024, 1, 1, 0, 0, 0)

#: Timestamp stamped on every annotation (comments, revisions).
FIXED_DATE = "2024-01-01T00:00:00Z"

#: Author stamped on every annotation.
FIXED_AUTHOR = "Fixture Author"

#: Initials stamped on every comment.
FIXED_INITIALS = "FA"

# Content types and relationship types python-docx does not expose as constants.
CT_COMMENTS_EXTENDED = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
RT_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"

#: Namespaces the fragments below may use, declared only when actually referenced so the
#: serialized parts stay free of redundant ``xmlns`` noise.
_NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

_ROOT_TAG_RE = re.compile(r"^<([A-Za-z0-9_:.-]+)")

_PREFIX_RES = {
    prefix: re.compile(rf"(?<![\w-]){re.escape(prefix)}:") for prefix in _NAMESPACES
}


# --------------------------------------------------------------------------------------
# Low-level helpers
# --------------------------------------------------------------------------------------


def _element(xml: str):
    """Parse an XML fragment, declaring the namespace prefixes it references."""
    stripped = xml.strip()
    match = _ROOT_TAG_RE.match(stripped)
    if match is None:  # pragma: no cover - defensive, fragments are literals
        raise ValueError(f"fragment does not start with a tag: {stripped[:40]!r}")
    decls = " ".join(
        f'xmlns:{prefix}="{uri}"'
        for prefix, uri in _NAMESPACES.items()
        if _PREFIX_RES[prefix].search(stripped)
    )
    injected = f"{stripped[: match.end()]} {decls}{stripped[match.end():]}"
    return parse_xml(injected)


def _body(doc):
    """Return the ``w:body`` element of `doc`."""
    return doc.element.body


def _append_body(doc, *fragments: str) -> None:
    """Append body-level fragments, keeping the trailing ``w:sectPr`` last."""
    body = _body(doc)
    sect_pr = body.find(qn("w:sectPr"))
    for fragment in fragments:
        element = _element(fragment)
        if sect_pr is None:
            body.append(element)
        else:
            sect_pr.addprevious(element)


def _append_styles(doc, *fragments: str) -> None:
    """Append ``w:style`` definitions to the styles part."""
    styles = doc.styles.element
    for fragment in fragments:
        styles.append(_element(fragment))


def _add_xml_part(doc, partname: str, content_type: str, reltype: str, xml: str) -> Part:
    """Attach a new XML part to the document part and return it."""
    part = Part(
        PackURI(partname),
        content_type,
        serialize_part_xml(_element(xml)),
        doc.part.package,
    )
    doc.part.relate_to(part, reltype)
    return part


def _png() -> bytes:
    """Return a deterministic 1x1 red PNG (python-docx reads PNG headers natively)."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        payload = tag + data
        return struct.pack(">I", len(data)) + payload + struct.pack(">I", zlib.crc32(payload))

    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    idat = zlib.compress(b"\x00\xff\x00\x00", 9)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


def _normalize_zip(blob: bytes) -> bytes:
    """Rewrite `blob` with frozen ZIP metadata, preserving member order."""
    source = zipfile.ZipFile(io.BytesIO(blob))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            info = zipfile.ZipInfo(name, date_time=FIXED_ZIP_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, source.read(name))
    source.close()
    return buffer.getvalue()


def _finalize(doc) -> bytes:
    """Serialize `doc` to deterministic ``.docx`` bytes."""
    buffer = io.BytesIO()
    doc.save(buffer)
    return _normalize_zip(buffer.getvalue())


def _build(*populators: Callable[[object], None]) -> bytes:
    """Build a document from the default template by running `populators` in order."""
    doc = Document()
    for populate in populators:
        populate(doc)
    return _finalize(doc)


# --------------------------------------------------------------------------------------
# Populators
# --------------------------------------------------------------------------------------


def _populate_simple(doc) -> None:
    """Headings and plain paragraphs, the baseline every other fixture builds on."""
    doc.add_heading("Fixture: simple", level=1)
    doc.add_paragraph("First paragraph of the simple fixture.")
    doc.add_paragraph("Second paragraph, with a trailing sentence.")
    doc.add_heading("Second heading", level=2)
    doc.add_paragraph("Body text under the second heading.")


def _populate_mixed_runs(doc) -> None:
    """Runs split by ``rsid`` and by run properties inside a single paragraph."""
    doc.add_heading("Fixture: mixed runs", level=2)
    _append_body(
        doc,
        # Runs split by formatting: plain, bold, bold+italic, language override.
        """
        <w:p w:rsidR="00A10001" w:rsidRPr="00A10001" w:rsidRDefault="00A10001">
          <w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:pPr>
          <w:r w:rsidR="00A10001"><w:t xml:space="preserve">Plain then </w:t></w:r>
          <w:r w:rsidR="00A10002"><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold </w:t></w:r>
          <w:r w:rsidR="00A10002"><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">bold italic </w:t></w:r>
          <w:r w:rsidR="00A10003"><w:rPr><w:lang w:val="fr-FR"/></w:rPr><w:t>et la fin.</w:t></w:r>
        </w:p>
        """,
        # Runs split by rsid only: identical rPr, three fragments of one sentence.
        """
        <w:p w:rsidR="00A10004" w:rsidRDefault="00A10004">
          <w:r w:rsidR="00A10004"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">One senten</w:t></w:r>
          <w:r w:rsidR="00A10005"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">ce cut in thr</w:t></w:r>
          <w:r w:rsidR="00A10006"><w:rPr><w:i/></w:rPr><w:t>ee runs.</w:t></w:r>
        </w:p>
        """,
        # Non-text run content: tab, break, symbol.
        """
        <w:p w:rsidR="00A10007" w:rsidRDefault="00A10007">
          <w:r w:rsidR="00A10007"><w:t>Before</w:t><w:tab/><w:t>after tab</w:t><w:br/><w:t>after break</w:t></w:r>
          <w:r w:rsidR="00A10008"><w:rPr><w:noProof/></w:rPr><w:t>no proofing</w:t></w:r>
        </w:p>
        """,
    )


def _populate_paragraph_styles(doc) -> None:
    """Custom and built-in paragraph styles referenced through ``w:pStyle``."""
    _append_styles(
        doc,
        """
        <w:style w:type="paragraph" w:styleId="FixtureBody">
          <w:name w:val="Fixture Body"/>
          <w:qFormat/>
          <w:pPr><w:spacing w:before="120" w:after="120"/><w:jc w:val="both"/></w:pPr>
          <w:rPr><w:sz w:val="22"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="paragraph" w:styleId="FixtureNote">
          <w:name w:val="Fixture Note"/>
          <w:pPr><w:ind w:left="720"/></w:pPr>
          <w:rPr><w:i/><w:color w:val="595959"/></w:rPr>
        </w:style>
        """,
    )
    doc.add_heading("Fixture: paragraph styles", level=1)
    _append_body(
        doc,
        '<w:p><w:pPr><w:pStyle w:val="FixtureBody"/></w:pPr><w:r><w:t>Custom body style.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="FixtureNote"/></w:pPr><w:r><w:t>Custom note style.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>Built-in Quote style.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:ind w:left="720"/></w:pPr><w:r><w:t>Built-in ListParagraph style.</w:t></w:r></w:p>',
        # Direct formatting on top of a style: the engine must not confuse the two.
        '<w:p><w:pPr><w:pStyle w:val="FixtureBody"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Styled and directly centered.</w:t></w:r></w:p>',
    )


def _populate_character_styles(doc) -> None:
    """Character styles referenced through ``w:rStyle``."""
    _append_styles(
        doc,
        """
        <w:style w:type="character" w:styleId="FixtureEmphasis">
          <w:name w:val="Fixture Emphasis"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr><w:i/><w:color w:val="C00000"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="character" w:styleId="FixtureCode">
          <w:name w:val="Fixture Code"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:noProof/></w:rPr>
        </w:style>
        """,
    )
    doc.add_heading("Fixture: character styles", level=2)
    _append_body(
        doc,
        """
        <w:p>
          <w:r><w:t xml:space="preserve">Plain, then </w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="FixtureEmphasis"/></w:rPr><w:t xml:space="preserve">emphasised</w:t></w:r>
          <w:r><w:t xml:space="preserve">, then </w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="FixtureCode"/></w:rPr><w:t>code_span()</w:t></w:r>
          <w:r><w:t xml:space="preserve">, then </w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>built-in Strong</w:t></w:r>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
        """,
        # Character style plus direct formatting on the same run.
        """
        <w:p>
          <w:r><w:rPr><w:rStyle w:val="FixtureEmphasis"/><w:b/><w:u w:val="single"/></w:rPr><w:t>Style plus direct bold underline.</w:t></w:r>
        </w:p>
        """,
    )


def _populate_style_inheritance(doc) -> None:
    """A ``basedOn`` chain wired with ``w:next`` and ``w:link``."""
    _append_styles(
        doc,
        """
        <w:style w:type="paragraph" w:styleId="FixtureRoot">
          <w:name w:val="Fixture Root"/>
          <w:basedOn w:val="Normal"/>
          <w:next w:val="FixtureRoot"/>
          <w:qFormat/>
          <w:pPr><w:spacing w:after="200"/></w:pPr>
          <w:rPr><w:sz w:val="24"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="paragraph" w:styleId="FixtureBranch">
          <w:name w:val="Fixture Branch"/>
          <w:basedOn w:val="FixtureRoot"/>
          <w:next w:val="FixtureRoot"/>
          <w:link w:val="FixtureBranchChar"/>
          <w:pPr><w:ind w:left="360"/></w:pPr>
          <w:rPr><w:b/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="character" w:customStyle="1" w:styleId="FixtureBranchChar">
          <w:name w:val="Fixture Branch Char"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:link w:val="FixtureBranch"/>
          <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="paragraph" w:styleId="FixtureLeaf">
          <w:name w:val="Fixture Leaf"/>
          <w:basedOn w:val="FixtureBranch"/>
          <w:next w:val="FixtureBranch"/>
          <w:pPr><w:ind w:left="720"/></w:pPr>
          <w:rPr><w:i/></w:rPr>
        </w:style>
        """,
    )
    doc.add_heading("Fixture: style inheritance", level=1)
    _append_body(
        doc,
        '<w:p><w:pPr><w:pStyle w:val="FixtureRoot"/></w:pPr><w:r><w:t>Root of the chain.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="FixtureBranch"/></w:pPr><w:r><w:t>Branch, based on root, linked to a character style.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="FixtureLeaf"/></w:pPr><w:r><w:t>Leaf, based on branch.</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:rStyle w:val="FixtureBranchChar"/></w:rPr><w:t>Linked character style used on its own.</w:t></w:r></w:p>',
    )


def _populate_themes(doc) -> None:
    """Theme-relative fonts and colours (``asciiTheme``, ``themeColor`` + tint/shade)."""
    doc.add_heading("Fixture: themes", level=1)
    _append_body(
        doc,
        """
        <w:p>
          <w:pPr><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:pPr>
          <w:r>
            <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/></w:rPr>
            <w:t xml:space="preserve">Major theme font. </w:t>
          </w:r>
          <w:r>
            <w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr>
            <w:t xml:space="preserve">Accent 1. </w:t>
          </w:r>
          <w:r>
            <w:rPr><w:color w:val="B4C6E7" w:themeColor="accent1" w:themeTint="66"/></w:rPr>
            <w:t xml:space="preserve">Accent 1 tinted. </w:t>
          </w:r>
          <w:r>
            <w:rPr><w:color w:val="2F5597" w:themeColor="accent1" w:themeShade="BF"/></w:rPr>
            <w:t>Accent 1 shaded.</w:t>
          </w:r>
        </w:p>
        """,
        """
        <w:p>
          <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/></w:pPr>
          <w:r><w:t>Paragraph shaded with a theme fill.</w:t></w:r>
        </w:p>
        """,
    )


def _populate_complex_numbering(doc) -> None:
    """Multi-level ``abstractNum`` definitions, a restart, and a start override."""
    numbering = doc.part.numbering_part.element
    first_num = numbering.find(qn("w:num"))

    multilevel = _element(
        """
        <w:abstractNum w:abstractNumId="900">
          <w:nsid w:val="1A2B3C4D"/>
          <w:multiLevelType w:val="multilevel"/>
          <w:lvl w:ilvl="0">
            <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>
          </w:lvl>
          <w:lvl w:ilvl="1">
            <w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlRestart w:val="1"/>
            <w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
          </w:lvl>
          <w:lvl w:ilvl="2">
            <w:start w:val="3"/><w:numFmt w:val="lowerRoman"/><w:lvlRestart w:val="0"/>
            <w:lvlText w:val="%3)"/><w:lvlJc w:val="right"/>
            <w:pPr><w:ind w:left="1080" w:hanging="180"/></w:pPr>
          </w:lvl>
        </w:abstractNum>
        """
    )
    bullets = _element(
        """
        <w:abstractNum w:abstractNumId="901">
          <w:nsid w:val="5E6F7A8B"/>
          <w:multiLevelType w:val="hybridMultilevel"/>
          <w:lvl w:ilvl="0">
            <w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
            <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
          </w:lvl>
        </w:abstractNum>
        """
    )
    for abstract in (multilevel, bullets):
        if first_num is None:
            numbering.append(abstract)
        else:
            first_num.addprevious(abstract)

    numbering.append(_element('<w:num w:numId="900"><w:abstractNumId w:val="900"/></w:num>'))
    numbering.append(
        _element(
            """
            <w:num w:numId="901">
              <w:abstractNumId w:val="900"/>
              <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>
            </w:num>
            """
        )
    )
    numbering.append(_element('<w:num w:numId="902"><w:abstractNumId w:val="901"/></w:num>'))

    doc.add_heading("Fixture: complex numbering", level=1)
    _append_body(
        doc,
        *[
            f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'
            for ilvl, num_id, text in (
                (0, 900, "First top level item"),
                (1, 900, "Nested letter item"),
                (2, 900, "Nested roman item"),
                (0, 900, "Second top level item"),
                (0, 901, "Overridden list restarting at five"),
                (0, 902, "Bulleted item"),
            )
        ],
    )


def _populate_tables(doc) -> None:
    """A table with a horizontal merge, a vertical merge, and a nested table."""
    doc.add_heading("Fixture: tables", level=1)
    _append_body(
        doc,
        """
        <w:tbl>
          <w:tblPr>
            <w:tblStyle w:val="TableGrid"/>
            <w:tblW w:w="0" w:type="auto"/>
            <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
          </w:tblPr>
          <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
          <w:tr>
            <w:tc>
              <w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>
              <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Merged header across two columns</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
              <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Third column</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
              <w:p><w:r><w:t>Vertically merged cell</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
              <w:p><w:r><w:t>Row two, column two</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
              <w:tbl>
                <w:tblPr>
                  <w:tblStyle w:val="TableGrid"/>
                  <w:tblW w:w="0" w:type="auto"/>
                  <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
                </w:tblPr>
                <w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>
                <w:tr>
                  <w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Nested A</w:t></w:r></w:p></w:tc>
                  <w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Nested B</w:t></w:r></w:p></w:tc>
                </w:tr>
              </w:tbl>
              <w:p/>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr>
              <w:p/>
            </w:tc>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
              <w:p><w:r><w:t>Row three, column two</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
              <w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/></w:tcPr>
              <w:p><w:r><w:t>Row three, column three</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
        """,
        "<w:p><w:r><w:t>Paragraph after the table.</w:t></w:r></w:p>",
    )


def _populate_comments(doc) -> None:
    """A ``comments.xml`` part plus a ``commentsExtended.xml`` reply thread."""
    _append_styles(
        doc,
        """
        <w:style w:type="paragraph" w:styleId="CommentText">
          <w:name w:val="annotation text"/>
          <w:basedOn w:val="Normal"/>
          <w:link w:val="CommentTextChar"/>
          <w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
          <w:rPr><w:sz w:val="20"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="character" w:customStyle="1" w:styleId="CommentTextChar">
          <w:name w:val="Comment Text Char"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:link w:val="CommentText"/>
          <w:rPr><w:sz w:val="20"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="character" w:styleId="CommentReference">
          <w:name w:val="annotation reference"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
        </w:style>
        """,
    )
    _add_xml_part(
        doc,
        "/word/comments.xml",
        CT.WML_COMMENTS,
        RT.COMMENTS,
        f"""
        <w:comments>
          <w:comment w:id="1" w:author="{FIXED_AUTHOR}" w:initials="{FIXED_INITIALS}" w:date="{FIXED_DATE}">
            <w:p w14:paraId="10000001" w14:textId="10000001">
              <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
              <w:r><w:t>Root comment anchored on the first sentence.</w:t></w:r>
            </w:p>
          </w:comment>
          <w:comment w:id="2" w:author="Fixture Reviewer" w:initials="FR" w:date="{FIXED_DATE}">
            <w:p w14:paraId="10000002" w14:textId="10000002">
              <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
              <w:r><w:t>Reply to the root comment.</w:t></w:r>
            </w:p>
          </w:comment>
          <w:comment w:id="3" w:author="{FIXED_AUTHOR}" w:initials="{FIXED_INITIALS}" w:date="{FIXED_DATE}">
            <w:p w14:paraId="10000003" w14:textId="10000003">
              <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
              <w:r><w:t>Standalone comment spanning two paragraphs.</w:t></w:r>
            </w:p>
          </w:comment>
        </w:comments>
        """,
    )
    _add_xml_part(
        doc,
        "/word/commentsExtended.xml",
        CT_COMMENTS_EXTENDED,
        RT_COMMENTS_EXTENDED,
        """
        <w15:commentsEx>
          <w15:commentEx w15:paraId="10000001" w15:done="0"/>
          <w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="0"/>
          <w15:commentEx w15:paraId="10000003" w15:done="1"/>
        </w15:commentsEx>
        """,
    )

    doc.add_heading("Fixture: comments", level=1)
    _append_body(
        doc,
        # Range fully inside one paragraph, carrying the threaded comment 1 and its reply 2.
        """
        <w:p>
          <w:commentRangeStart w:id="1"/>
          <w:commentRangeStart w:id="2"/>
          <w:r><w:t xml:space="preserve">This sentence carries a comment thread.</w:t></w:r>
          <w:commentRangeEnd w:id="1"/>
          <w:commentRangeEnd w:id="2"/>
          <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
          <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="2"/></w:r>
          <w:r><w:t xml:space="preserve"> The rest of the paragraph is not commented.</w:t></w:r>
        </w:p>
        """,
        # Range spanning two paragraphs: start and end markers live in different w:p.
        """
        <w:p>
          <w:commentRangeStart w:id="3"/>
          <w:r><w:t>The third comment starts here</w:t></w:r>
        </w:p>
        """,
        """
        <w:p>
          <w:r><w:t>and ends in the next paragraph.</w:t></w:r>
          <w:commentRangeEnd w:id="3"/>
          <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="3"/></w:r>
        </w:p>
        """,
    )


def _populate_tracked_changes(doc) -> None:
    """Insertions, deletions, an insertion-then-deletion, and property revisions."""
    doc.add_heading("Fixture: tracked changes", level=1)
    stamp = f'w:author="{FIXED_AUTHOR}" w:date="{FIXED_DATE}"'
    _append_body(
        doc,
        # Plain insertion and plain deletion side by side.
        f"""
        <w:p>
          <w:r><w:t xml:space="preserve">Kept text, </w:t></w:r>
          <w:ins w:id="201" {stamp}><w:r><w:t xml:space="preserve">inserted text, </w:t></w:r></w:ins>
          <w:del w:id="202" {stamp}><w:r><w:delText xml:space="preserve">deleted text, </w:delText></w:r></w:del>
          <w:r><w:t>and kept tail.</w:t></w:r>
        </w:p>
        """,
        # Text inserted then deleted by a later revision: w:del nested inside w:ins.
        f"""
        <w:p>
          <w:r><w:t xml:space="preserve">Before. </w:t></w:r>
          <w:ins w:id="203" {stamp}>
            <w:del w:id="204" w:author="Fixture Reviewer" w:date="{FIXED_DATE}">
              <w:r><w:delText xml:space="preserve">inserted then deleted</w:delText></w:r>
            </w:del>
          </w:ins>
          <w:r><w:t xml:space="preserve"> After.</w:t></w:r>
        </w:p>
        """,
        # Run property revision: rPrChange records the previous rPr.
        f"""
        <w:p>
          <w:r>
            <w:rPr>
              <w:b/>
              <w:rPrChange w:id="205" {stamp}><w:rPr><w:i/></w:rPr></w:rPrChange>
            </w:rPr>
            <w:t>This run became bold, it used to be italic.</w:t>
          </w:r>
        </w:p>
        """,
        # Paragraph property revision: pPrChange records the previous pPr.
        f"""
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:pPrChange w:id="206" {stamp}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>
          </w:pPr>
          <w:r><w:t>This paragraph became centered.</w:t></w:r>
        </w:p>
        """,
        # Deleted paragraph mark: w:del inside the paragraph-mark rPr.
        f"""
        <w:p>
          <w:pPr><w:rPr><w:del w:id="207" {stamp}/></w:rPr></w:pPr>
          <w:r><w:t>This paragraph mark is deleted, so this merges with the next one.</w:t></w:r>
        </w:p>
        """,
        # Inserted paragraph mark, closing the pair.
        f"""
        <w:p>
          <w:pPr><w:rPr><w:ins w:id="208" {stamp}/></w:rPr></w:pPr>
          <w:r><w:t>This paragraph mark is inserted.</w:t></w:r>
        </w:p>
        """,
    )


def _populate_hyperlinks(doc) -> None:
    """An external hyperlink relationship and an internal anchor hyperlink."""
    _append_styles(
        doc,
        """
        <w:style w:type="character" w:styleId="Hyperlink">
          <w:name w:val="Hyperlink"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>
        </w:style>
        """,
    )
    r_id = doc.part.relate_to("https://example.org/fixture", RT.HYPERLINK, is_external=True)
    doc.add_heading("Fixture: hyperlinks", level=1)
    _append_body(
        doc,
        f"""
        <w:p>
          <w:r><w:t xml:space="preserve">An external link: </w:t></w:r>
          <w:hyperlink r:id="{r_id}" w:history="1">
            <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>example.org</w:t></w:r>
          </w:hyperlink>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
        """,
        """
        <w:p>
          <w:r><w:t xml:space="preserve">An internal link: </w:t></w:r>
          <w:hyperlink w:anchor="FixtureAnchor" w:history="1">
            <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>jump to the anchor</w:t></w:r>
          </w:hyperlink>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
        """,
        """
        <w:p>
          <w:bookmarkStart w:id="300" w:name="FixtureAnchor"/>
          <w:r><w:t>Anchor target paragraph.</w:t></w:r>
          <w:bookmarkEnd w:id="300"/>
        </w:p>
        """,
    )


def _populate_bookmarks(doc) -> None:
    """Bookmarks inside a paragraph, spanning paragraphs, and empty."""
    doc.add_heading("Fixture: bookmarks", level=1)
    _append_body(
        doc,
        # Bookmark wrapping a few runs inside one paragraph.
        """
        <w:p>
          <w:r><w:t xml:space="preserve">Before. </w:t></w:r>
          <w:bookmarkStart w:id="310" w:name="FixtureInline"/>
          <w:r><w:t xml:space="preserve">Bookmarked </w:t></w:r>
          <w:r><w:rPr><w:b/></w:rPr><w:t>span</w:t></w:r>
          <w:bookmarkEnd w:id="310"/>
          <w:r><w:t xml:space="preserve">. After.</w:t></w:r>
        </w:p>
        """,
        # Bookmark spanning two paragraphs.
        """
        <w:p>
          <w:bookmarkStart w:id="311" w:name="FixtureSpanning"/>
          <w:r><w:t>A bookmark opens here</w:t></w:r>
        </w:p>
        """,
        """
        <w:p>
          <w:r><w:t>and closes in this paragraph.</w:t></w:r>
          <w:bookmarkEnd w:id="311"/>
        </w:p>
        """,
        # Empty bookmark: start immediately followed by its end, a pure insertion point.
        """
        <w:p>
          <w:r><w:t xml:space="preserve">An empty bookmark sits here: </w:t></w:r>
          <w:bookmarkStart w:id="312" w:name="FixtureEmpty"/>
          <w:bookmarkEnd w:id="312"/>
          <w:r><w:t xml:space="preserve"> and text resumes.</w:t></w:r>
        </w:p>
        """,
        # Bookmark wrapping a whole table row, a range that is not run-level.
        """
        <w:p>
          <w:bookmarkStart w:id="313" w:name="_Hidden_Word_Bookmark"/>
          <w:r><w:t>Bookmark whose name follows the reserved underscore convention.</w:t></w:r>
          <w:bookmarkEnd w:id="313"/>
        </w:p>
        """,
    )


def _populate_fields(doc) -> None:
    """A simple field, a complex TOC field, and a nested field."""
    doc.add_heading("Fixture: fields", level=1)
    _append_body(
        doc,
        # fldSimple: atomic, instruction stored as an attribute.
        """
        <w:p>
          <w:r><w:t xml:space="preserve">Page </w:t></w:r>
          <w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT ">
            <w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>
          </w:fldSimple>
          <w:r><w:t xml:space="preserve"> of the document.</w:t></w:r>
        </w:p>
        """,
        # Complex field: begin / instrText / separate / cached result / end.
        """
        <w:p>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>Fixture: fields</w:t></w:r>
          <w:r><w:tab/><w:t>1</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
        """,
        # Nested field: an IF field whose instruction embeds a PAGE field.
        """
        <w:p>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>1</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:r><w:instrText xml:space="preserve"> = 1 "first page" "later page" </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>first page</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
        """,
        # Cross-reference field pointing at a bookmark, dirtied so Word refreshes it.
        """
        <w:p>
          <w:bookmarkStart w:id="320" w:name="FixtureRefTarget"/>
          <w:r><w:t>Referenced paragraph.</w:t></w:r>
          <w:bookmarkEnd w:id="320"/>
        </w:p>
        """,
        """
        <w:p>
          <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
          <w:r><w:instrText xml:space="preserve"> REF FixtureRefTarget \\h </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>Referenced paragraph.</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
        """,
    )


def _populate_footnotes(doc) -> None:
    """A ``footnotes.xml`` part and an ``endnotes.xml`` part, both referenced."""
    _append_styles(
        doc,
        """
        <w:style w:type="paragraph" w:styleId="FootnoteText">
          <w:name w:val="footnote text"/>
          <w:basedOn w:val="Normal"/>
          <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
          <w:rPr><w:sz w:val="20"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="character" w:styleId="FootnoteReference">
          <w:name w:val="footnote reference"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="paragraph" w:styleId="EndnoteText">
          <w:name w:val="endnote text"/>
          <w:basedOn w:val="Normal"/>
          <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
          <w:rPr><w:sz w:val="20"/></w:rPr>
        </w:style>
        """,
        """
        <w:style w:type="character" w:styleId="EndnoteReference">
          <w:name w:val="endnote reference"/>
          <w:basedOn w:val="DefaultParagraphFont"/>
          <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
        </w:style>
        """,
    )
    _add_xml_part(
        doc,
        "/word/footnotes.xml",
        CT.WML_FOOTNOTES,
        RT.FOOTNOTES,
        """
        <w:footnotes>
          <w:footnote w:type="separator" w:id="-1">
            <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>
          </w:footnote>
          <w:footnote w:type="continuationSeparator" w:id="0">
            <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>
          </w:footnote>
          <w:footnote w:id="2">
            <w:p>
              <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
              <w:r><w:t xml:space="preserve"> First footnote body.</w:t></w:r>
            </w:p>
          </w:footnote>
          <w:footnote w:id="3">
            <w:p>
              <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
              <w:r><w:t xml:space="preserve"> Second footnote body, with </w:t></w:r>
              <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
              <w:r><w:t xml:space="preserve"> text.</w:t></w:r>
            </w:p>
          </w:footnote>
        </w:footnotes>
        """,
    )
    _add_xml_part(
        doc,
        "/word/endnotes.xml",
        CT.WML_ENDNOTES,
        RT.ENDNOTES,
        """
        <w:endnotes>
          <w:endnote w:type="separator" w:id="-1">
            <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>
          </w:endnote>
          <w:endnote w:type="continuationSeparator" w:id="0">
            <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>
          </w:endnote>
          <w:endnote w:id="2">
            <w:p>
              <w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>
              <w:r><w:t xml:space="preserve"> Endnote body.</w:t></w:r>
            </w:p>
          </w:endnote>
        </w:endnotes>
        """,
    )

    doc.add_heading("Fixture: footnotes", level=1)
    _append_body(
        doc,
        """
        <w:p>
          <w:r><w:t>A sentence with a footnote</w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>
          <w:r><w:t xml:space="preserve"> and a second one</w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="3"/></w:r>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
        """,
        """
        <w:p>
          <w:r><w:t>A sentence with an endnote</w:t></w:r>
          <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
        """,
    )


def _populate_headers_footers(doc) -> None:
    """A header holding an image and a footer holding page-number fields."""
    doc.add_heading("Fixture: headers and footers", level=1)
    doc.add_paragraph("Body text; the header and footer carry the interesting parts.")

    section = doc.sections[0]
    section.different_first_page_header_footer = True

    header = section.header
    header.is_linked_to_previous = False
    header_paragraph = header.paragraphs[0]
    header_paragraph.text = "Fixture header "
    header_paragraph.add_run().add_picture(io.BytesIO(_png()), width=Pt(12))

    first_page_header = section.first_page_header
    first_page_header.is_linked_to_previous = False
    first_page_header.paragraphs[0].text = "Fixture first-page header"

    footer = section.footer
    footer.is_linked_to_previous = False
    footer_paragraph = footer.paragraphs[0]
    for fragment in (
        '<w:r><w:t xml:space="preserve">Page </w:t></w:r>',
        '<w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>',
        '<w:r><w:t xml:space="preserve"> of </w:t></w:r>',
        '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>',
    ):
        footer_paragraph._p.append(_element(fragment))


def _populate_sections(doc) -> None:
    """Two sections: portrait then landscape, with distinct page setup."""
    doc.add_heading("Fixture: sections", level=1)
    doc.add_paragraph("First section, portrait, default margins.")

    second = doc.add_section(WD_SECTION.NEW_PAGE)
    width, height = second.page_width, second.page_height
    second.orientation = WD_ORIENT.LANDSCAPE
    second.page_width, second.page_height = height, width
    second.left_margin = Inches(0.5)
    second.right_margin = Inches(0.5)

    doc.add_paragraph("Second section, landscape, narrow margins.")


def _populate_content_controls(doc) -> None:
    """A block-level ``w:sdt`` and an inline ``w:sdt``."""
    doc.add_heading("Fixture: content controls", level=1)
    _append_body(
        doc,
        # Block-level structured document tag wrapping a whole paragraph.
        """
        <w:sdt>
          <w:sdtPr>
            <w:alias w:val="Fixture block control"/>
            <w:tag w:val="fixture-block"/>
            <w:id w:val="770001"/>
            <w:text w:multiLine="1"/>
          </w:sdtPr>
          <w:sdtContent>
            <w:p><w:r><w:t>Paragraph inside a block content control.</w:t></w:r></w:p>
          </w:sdtContent>
        </w:sdt>
        """,
        # Inline structured document tag wrapping runs inside a paragraph.
        """
        <w:p>
          <w:r><w:t xml:space="preserve">Before the control, </w:t></w:r>
          <w:sdt>
            <w:sdtPr>
              <w:alias w:val="Fixture inline control"/>
              <w:tag w:val="fixture-inline"/>
              <w:id w:val="770002"/>
              <w:showingPlcHdr/>
              <w:text/>
            </w:sdtPr>
            <w:sdtContent>
              <w:r><w:t>inline control content</w:t></w:r>
            </w:sdtContent>
          </w:sdt>
          <w:r><w:t>, after the control.</w:t></w:r>
        </w:p>
        """,
        # A dropdown control, whose sdtPr carries list items.
        """
        <w:sdt>
          <w:sdtPr>
            <w:alias w:val="Fixture dropdown"/>
            <w:tag w:val="fixture-dropdown"/>
            <w:id w:val="770003"/>
            <w:dropDownList w:lastValue="Alpha">
              <w:listItem w:displayText="Alpha" w:value="alpha"/>
              <w:listItem w:displayText="Beta" w:value="beta"/>
            </w:dropDownList>
          </w:sdtPr>
          <w:sdtContent>
            <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
          </w:sdtContent>
        </w:sdt>
        """,
    )


def _populate_drawings(doc) -> None:
    """Inline images, each carrying its own drawing relationship."""
    doc.add_heading("Fixture: drawings", level=1)
    doc.add_paragraph("An inline image follows.")
    doc.add_picture(io.BytesIO(_png()), width=Inches(1))
    paragraph = doc.add_paragraph("A second inline image, inside a text paragraph: ")
    paragraph.add_run().add_picture(io.BytesIO(_png()), width=Pt(20))
    paragraph.add_run(" and text after it.")


def _populate_combined(doc) -> None:
    """Everything that can coexist in one document, sections last."""
    for populate in (
        _populate_simple,
        _populate_mixed_runs,
        _populate_paragraph_styles,
        _populate_character_styles,
        _populate_style_inheritance,
        _populate_themes,
        _populate_complex_numbering,
        _populate_tables,
        _populate_comments,
        _populate_tracked_changes,
        _populate_hyperlinks,
        _populate_bookmarks,
        _populate_fields,
        _populate_footnotes,
        _populate_content_controls,
        _populate_drawings,
        _populate_headers_footers,
        _populate_sections,
    ):
        populate(doc)


# --------------------------------------------------------------------------------------
# Builders
# --------------------------------------------------------------------------------------


def build_simple() -> bytes:
    """Headings and plain paragraphs only."""
    return _build(_populate_simple)


def build_mixed_runs() -> bytes:
    """Runs split by rsid and by run properties."""
    return _build(_populate_mixed_runs)


def build_paragraph_styles() -> bytes:
    """Custom and built-in paragraph styles."""
    return _build(_populate_paragraph_styles)


def build_character_styles() -> bytes:
    """Character styles applied through ``w:rStyle``."""
    return _build(_populate_character_styles)


def build_style_inheritance() -> bytes:
    """A ``basedOn`` chain with ``w:next`` and ``w:link``."""
    return _build(_populate_style_inheritance)


def build_themes() -> bytes:
    """Theme-relative fonts and colours."""
    return _build(_populate_themes)


def build_complex_numbering() -> bytes:
    """Multi-level numbering with restart and start override."""
    return _build(_populate_complex_numbering)


def build_tables() -> bytes:
    """Merged cells and a nested table."""
    return _build(_populate_tables)


def build_comments() -> bytes:
    """Comments with a reply thread."""
    return _build(_populate_comments)


def build_tracked_changes() -> bytes:
    """Insertions, deletions, and property revisions."""
    return _build(_populate_tracked_changes)


def build_hyperlinks() -> bytes:
    """External and anchor hyperlinks."""
    return _build(_populate_hyperlinks)


def build_bookmarks() -> bytes:
    """Inline, spanning, and empty bookmarks."""
    return _build(_populate_bookmarks)


def build_fields() -> bytes:
    """Simple, complex, and nested fields."""
    return _build(_populate_fields)


def build_footnotes() -> bytes:
    """Footnotes and endnotes."""
    return _build(_populate_footnotes)


def build_headers_footers() -> bytes:
    """Headers and footers holding fields and an image."""
    return _build(_populate_headers_footers)


def build_sections() -> bytes:
    """Two sections with distinct page setup."""
    return _build(_populate_sections)


def build_content_controls() -> bytes:
    """Block-level and inline content controls."""
    return _build(_populate_content_controls)


def build_drawings() -> bytes:
    """Inline images."""
    return _build(_populate_drawings)


def build_combined() -> bytes:
    """Every feature above in a single document."""
    return _build(_populate_combined)


#: Public naming contract: fixture name -> zero-argument builder returning ``.docx`` bytes.
ALL_FIXTURES: dict[str, Callable[[], bytes]] = {
    "simple": build_simple,
    "mixed_runs": build_mixed_runs,
    "paragraph_styles": build_paragraph_styles,
    "character_styles": build_character_styles,
    "style_inheritance": build_style_inheritance,
    "themes": build_themes,
    "complex_numbering": build_complex_numbering,
    "tables": build_tables,
    "comments": build_comments,
    "tracked_changes": build_tracked_changes,
    "hyperlinks": build_hyperlinks,
    "bookmarks": build_bookmarks,
    "fields": build_fields,
    "footnotes": build_footnotes,
    "headers_footers": build_headers_footers,
    "sections": build_sections,
    "content_controls": build_content_controls,
    "drawings": build_drawings,
    "combined": build_combined,
}


def build(name: str) -> bytes:
    """Build the fixture registered under `name`.

    Raises:
        KeyError: if `name` is not a registered fixture, listing the known names.
    """
    try:
        builder = ALL_FIXTURES[name]
    except KeyError:
        raise KeyError(f"unknown fixture {name!r}; known: {sorted(ALL_FIXTURES)}") from None
    return builder()
