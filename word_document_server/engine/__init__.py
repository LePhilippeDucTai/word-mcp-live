"""Semantic OOXML engine.

A pure document-manipulation core: synchronous Python, lxml on top of the OPC
implementation of python-docx, no MCP dependency and no Word dependency, so it
runs on Linux and can be tested without any office suite.

Its contract is non-degradation: an operation touches what it was asked to
touch and leaves styles, relationships, fields, bookmarks, comments, tracked
changes, numbering and page layout byte-identical.  Layers are stacked bottom
up -- ``package`` (this milestone), then ``textmodel``, ``ranges``, ``format``,
and the feature modules above them.

Nothing here imports ``fastmcp`` or ``win32com``: the MCP tool layer depends on
the engine, never the other way round.
"""

from word_document_server.engine.errors import (
    EngineError,
    IdExhausted,
    InvalidText,
    LocatorError,
    PackageError,
    UnsupportedRange,
    UnsupportedRevision,
)
from word_document_server.engine.ids import (
    next_annotation_id,
    next_comment_id,
    next_footnote_id,
    next_para_id,
)
from word_document_server.engine.package import (
    MAIN_STORY,
    STORY_CONTENT_TYPES,
    DocxPackage,
    atomic_write_bytes,
    story_name,
)
from word_document_server.engine.xmlns import MC, NAMESPACES, W14, W15, XML, R, W, qn

__all__ = [
    "MAIN_STORY",
    "MC",
    "NAMESPACES",
    "STORY_CONTENT_TYPES",
    "W14",
    "W15",
    "XML",
    "DocxPackage",
    "EngineError",
    "IdExhausted",
    "InvalidText",
    "LocatorError",
    "PackageError",
    "R",
    "UnsupportedRange",
    "UnsupportedRevision",
    "W",
    "atomic_write_bytes",
    "next_annotation_id",
    "next_comment_id",
    "next_footnote_id",
    "next_para_id",
    "qn",
    "story_name",
]
