"""Namespace URIs and qualified-name helper for the OOXML engine.

The engine works directly on lxml elements, so it needs its own namespace table
rather than borrowing ``docx.oxml.ns``: the engine must stay usable on parts
python-docx knows nothing about (``commentsExtended``, ``people``, custom XML),
and its qualified names must not change when python-docx changes its own map.

Only the prefixes the engine actually writes are declared. :func:`qn` raises on
an unknown prefix rather than guessing, so a typo in a tag name fails at the
call site instead of silently producing an element in the null namespace.
"""

from __future__ import annotations

from functools import cache

#: WordprocessingML main namespace (``w:``).
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: Relationship references *inside* a part (``r:id``, ``r:embed``, ...).  Distinct
#: from the package relationship namespace used by ``.rels`` parts.
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: Word 2010 extensions (``w14:paraId``, ``w14:textId``, ...).
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

#: Word 2012 extensions (``w15:commentEx``, ``w15:person``, ...).
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

#: Markup compatibility (``mc:AlternateContent``, ``mc:Ignorable``).
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

#: The reserved XML namespace (``xml:space="preserve"``).
XML = "http://www.w3.org/XML/1998/namespace"

#: Prefix -> URI map accepted by :func:`qn` and usable as an lxml ``namespaces``
#: argument for XPath expressions.
NAMESPACES: dict[str, str] = {
    "w": W,
    "r": R,
    "w14": W14,
    "w15": W15,
    "mc": MC,
    "xml": XML,
}


@cache
def qn(tag: str) -> str:
    """Return the Clark-notation form of a prefixed name.

    ``qn("w:p")`` -> ``"{http://...wordprocessingml/2006/main}p"``.  Works for
    element tags and for attribute names alike.

    Raises:
        ValueError: if `tag` is not ``prefix:local``, or if the prefix is not one
            of :data:`NAMESPACES`.
    """
    prefix, sep, local = tag.partition(":")
    if not sep or not prefix or not local:
        raise ValueError(f"expected a 'prefix:local' name, got {tag!r}")
    try:
        uri = NAMESPACES[prefix]
    except KeyError:
        known = ", ".join(sorted(NAMESPACES))
        raise ValueError(f"unknown namespace prefix {prefix!r} in {tag!r}; known: {known}") from None
    return f"{{{uri}}}{local}"
