"""OPC package validator for .docx files.

`validate_package` inspects an OOXML (.docx) package as an invariant of
non-degradation: it is used both to check fixtures produced by LibreOffice
and, later in this plan, to confirm that tools do not corrupt a document
while editing it. It never opens the file with Word or LibreOffice; it only
reasons about the zip package and its XML parts (same style as
`word_document_server/core/footnotes.py:validate_document_footnotes`).
"""

import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

# Namespace definitions (OOXML / OPC).
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_W = f"{{{W_NS}}}"

_CONTENT_TYPES_PART = "[Content_Types].xml"
_MAIN_RELATIONSHIP_TYPE_SUFFIX = "/officeDocument"

# Parts that may carry bookmarks and tracked changes (document body stories).
_BODY_PART_NAMES = {"word/document.xml", "word/footnotes.xml", "word/endnotes.xml"}

# A footnote/endnote carrying `w:type` (`separator`, `continuationSeparator`,
# `continuationNotice`) is a producer-owned placeholder, never referenced from the
# body. Word and LibreOffice disagree on which *id* values these placeholders get
# (LibreOffice: separator=0, continuation=1, real notes from 2; Word: separator=-1,
# continuation=0, real notes from 1), so the id value carries no meaning on its own
# -- only the presence of `w:type` marks a note as reserved.


def _is_body_part(name: str) -> bool:
    if name in _BODY_PART_NAMES:
        return True
    stem = name.rsplit("/", 1)[-1]
    return stem.startswith(("header", "footer"))


@dataclass(frozen=True)
class Issue:
    """A single package validation finding."""

    code: str
    part: str
    message: str


def _resolve_target(base_dir: str, target: str) -> str:
    """Resolve a relationship Target against the directory of its source part."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def _rels_source_dir(rels_path: str) -> str:
    """Return the directory of the part described by a `.rels` file.

    `_rels/.rels` describes the package root ("") ; `word/_rels/document.xml.rels`
    describes parts rooted at "word".
    """
    marker = "/_rels/"
    idx = rels_path.rfind(marker)
    if idx == -1:
        return ""
    return rels_path[:idx]


def _find_duplicates(occurrences: dict[str, list[str]]) -> dict[str, list[str]]:
    return {key: parts for key, parts in occurrences.items() if len(parts) > 1}


def validate_package(path: str | Path) -> list[Issue]:
    """Validate a .docx package and return a list of Issue findings (empty if clean)."""
    path = Path(path)
    issues: list[Issue] = []

    try:
        zf = zipfile.ZipFile(path, "r")
    except (zipfile.BadZipFile, OSError) as exc:
        return [Issue("ZIP-UNREADABLE", str(path), str(exc))]

    with zf:
        bad_entry = zf.testzip()
        if bad_entry is not None:
            issues.append(
                Issue("ZIP-CORRUPT", bad_entry, "zip entry failed its CRC/consistency check")
            )

        names = [n for n in zf.namelist() if not n.endswith("/")]
        name_set = set(names)

        # Parse every XML-like part once; reused by every check below.
        xml_roots: dict[str, etree._Element] = {}
        for name in names:
            if not name.endswith((".xml", ".rels")):
                continue
            try:
                data = zf.read(name)
            except (zipfile.BadZipFile, OSError, KeyError) as exc:
                issues.append(Issue("XML-UNREADABLE", name, str(exc)))
                continue
            try:
                xml_roots[name] = etree.fromstring(data)
            except etree.XMLSyntaxError as exc:
                issues.append(Issue("XML-MALFORMED", name, str(exc)))

        # Content type declared for every part (Override by exact name, else
        # Default by extension).
        ct_root = xml_roots.get(_CONTENT_TYPES_PART)
        if ct_root is None:
            issues.append(
                Issue("CT-PART-MISSING", _CONTENT_TYPES_PART, "package has no [Content_Types].xml")
            )
        else:
            defaults = {
                el.get("Extension", "").lower()
                for el in ct_root.findall(f"{{{CT_NS}}}Default")
            }
            overrides = {
                el.get("PartName", "")
                for el in ct_root.findall(f"{{{CT_NS}}}Override")
            }
            for name in names:
                if name == _CONTENT_TYPES_PART:
                    continue
                part_name = "/" + name
                if part_name in overrides:
                    continue
                ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                if ext in defaults:
                    continue
                issues.append(
                    Issue("CT-MISSING", name, f"no content type declared for part /{name}")
                )

        # Relationship targets exist, for the package root and every part's
        # own .rels file.
        main_part_target: str | None = None
        for rels_name in names:
            if not rels_name.endswith(".rels"):
                continue
            rels_root = xml_roots.get(rels_name)
            if rels_root is None:
                continue  # already reported as XML-MALFORMED / unreadable
            base_dir = _rels_source_dir(rels_name)
            for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
                target = rel.get("Target", "")
                if rel.get("TargetMode") == "External":
                    continue
                resolved = _resolve_target(base_dir, target)
                if resolved not in name_set:
                    issues.append(
                        Issue(
                            "REL-TARGET-MISSING",
                            rels_name,
                            f"relationship {rel.get('Id')} target '{resolved}' not found in package",
                        )
                    )
                if rels_name == "_rels/.rels" and (rel.get("Type") or "").endswith(
                    _MAIN_RELATIONSHIP_TYPE_SUFFIX
                ):
                    main_part_target = resolved

        # Main part present.
        if main_part_target is None:
            issues.append(
                Issue(
                    "MAIN-PART-MISSING",
                    "_rels/.rels",
                    "no officeDocument relationship in the package root",
                )
            )
        elif main_part_target not in name_set:
            issues.append(
                Issue("MAIN-PART-MISSING", main_part_target, "main document part is absent")
            )

        # Ids unique per family: bookmarks and ins/del (revisions) across body
        # parts; comment ids within comments.xml.
        body_parts = [n for n in names if _is_body_part(n) and n in xml_roots]

        bookmark_occurrences: dict[str, list[str]] = {}
        revision_occurrences: dict[str, list[str]] = {}
        for name in body_parts:
            root = xml_roots[name]
            for el in root.iter(f"{_W}bookmarkStart"):
                bm_id = el.get(f"{_W}id")
                if bm_id is not None:
                    bookmark_occurrences.setdefault(bm_id, []).append(name)
            for tag in ("ins", "del"):
                for el in root.iter(f"{_W}{tag}"):
                    rev_id = el.get(f"{_W}id")
                    if rev_id is not None:
                        revision_occurrences.setdefault(rev_id, []).append(name)

        for bm_id, parts in _find_duplicates(bookmark_occurrences).items():
            issues.append(
                Issue(
                    "ID-DUPLICATE",
                    ", ".join(sorted(set(parts))),
                    f"bookmark id '{bm_id}' used {len(parts)} times",
                )
            )
        for rev_id, parts in _find_duplicates(revision_occurrences).items():
            issues.append(
                Issue(
                    "ID-DUPLICATE",
                    ", ".join(sorted(set(parts))),
                    f"revision (ins/del) id '{rev_id}' used {len(parts)} times",
                )
            )

        comments_root = xml_roots.get("word/comments.xml")
        comment_content_ids: set[str] = set()
        if comments_root is not None:
            comment_occurrences: dict[str, list[str]] = {}
            for el in comments_root.iter(f"{_W}comment"):
                c_id = el.get(f"{_W}id")
                if c_id is not None:
                    comment_content_ids.add(c_id)
                    comment_occurrences.setdefault(c_id, []).append("word/comments.xml")
            for c_id, parts in _find_duplicates(comment_occurrences).items():
                issues.append(
                    Issue(
                        "ID-DUPLICATE",
                        "word/comments.xml",
                        f"comment id '{c_id}' used {len(parts)} times",
                    )
                )

        # Comment ids referenced from body parts must exist in comments.xml,
        # and comments.xml must not carry orphaned comments.
        comment_ref_ids: set[str] = set()
        for name in body_parts:
            for el in xml_roots[name].iter(f"{_W}commentReference"):
                ref_id = el.get(f"{_W}id")
                if ref_id is not None:
                    comment_ref_ids.add(ref_id)
        if comment_ref_ids and comments_root is None:
            issues.append(
                Issue("COMMENT-MISSING", "word/comments.xml", "comments referenced but part is absent")
            )
        else:
            for ref_id in sorted(comment_ref_ids - comment_content_ids):
                issues.append(
                    Issue(
                        "COMMENT-MISSING",
                        "word/comments.xml",
                        f"commentReference id '{ref_id}' has no matching comment",
                    )
                )
            for orphan_id in sorted(comment_content_ids - comment_ref_ids):
                issues.append(
                    Issue(
                        "COMMENT-ORPHAN",
                        "word/comments.xml",
                        f"comment id '{orphan_id}' is never referenced",
                    )
                )

        # Footnote / endnote ids referenced from body parts must exist in
        # their content part, and vice versa (mirrors validate_document_footnotes).
        for kind, ref_tag, content_part, content_tag in (
            ("footnote", "footnoteReference", "word/footnotes.xml", "footnote"),
            ("endnote", "endnoteReference", "word/endnotes.xml", "endnote"),
        ):
            ref_ids: set[str] = set()
            for name in body_parts:
                for el in xml_roots[name].iter(f"{_W}{ref_tag}"):
                    r_id = el.get(f"{_W}id")
                    if r_id is not None:
                        ref_ids.add(r_id)
            content_root = xml_roots.get(content_part)
            content_ids: set[str] = set()
            if content_root is not None:
                for el in content_root.iter(f"{_W}{content_tag}"):
                    c_id = el.get(f"{_W}id")
                    if c_id is None or el.get(f"{_W}type") is not None:
                        continue
                    content_ids.add(c_id)
            if ref_ids and content_root is None:
                issues.append(
                    Issue(
                        "NOTE-MISSING",
                        content_part,
                        f"{kind} references exist but {content_part} is absent",
                    )
                )
            else:
                for r_id in sorted(ref_ids - content_ids):
                    issues.append(
                        Issue(
                            "NOTE-MISSING",
                            content_part,
                            f"{ref_tag} id '{r_id}' has no matching {kind}",
                        )
                    )
                for c_id in sorted(content_ids - ref_ids):
                    issues.append(
                        Issue(
                            "NOTE-ORPHAN",
                            content_part,
                            f"{kind} id '{c_id}' is never referenced",
                        )
                    )

        # numId used by body parts resolve to a definition in numbering.xml.
        used_num_ids: set[str] = set()
        for name in body_parts:
            for el in xml_roots[name].iter(f"{_W}numId"):
                val = el.get(f"{_W}val")
                if val is not None:
                    used_num_ids.add(val)
        numbering_root = xml_roots.get("word/numbering.xml")
        if used_num_ids:
            if numbering_root is None:
                issues.append(
                    Issue(
                        "NUMID-UNRESOLVED",
                        "word/numbering.xml",
                        "numbering is used but word/numbering.xml is absent",
                    )
                )
            else:
                defined_num_ids = {
                    el.get(f"{_W}numId")
                    for el in numbering_root.iter(f"{_W}num")
                    if el.get(f"{_W}numId") is not None
                }
                for val in sorted(used_num_ids - defined_num_ids):
                    issues.append(
                        Issue(
                            "NUMID-UNRESOLVED",
                            "word/numbering.xml",
                            f"numId '{val}' has no matching <w:num> definition",
                        )
                    )

    # No stray temporary file next to the document on disk (outside the zip):
    # atomic saves must not leave a partially written sibling behind.
    parent = path.parent
    if parent.is_dir():
        for candidate in sorted(parent.iterdir()):
            if candidate == path:
                continue
            candidate_name = candidate.name
            if candidate_name.startswith("~$") or candidate_name.endswith(".tmp"):
                issues.append(
                    Issue(
                        "TEMP-FILE-PRESENT",
                        candidate_name,
                        f"stray temporary file next to the document: {candidate}",
                    )
                )

    return issues
