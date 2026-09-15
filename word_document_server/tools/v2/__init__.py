"""The V2 ``doc_*`` tool surface.

Every module of this package that exports ``TOOLS`` -- a list of
:class:`~word_document_server.tools.v2.registry.ToolSpec` -- is discovered and
registered by :func:`~word_document_server.tools.v2.registry.register_v2_tools`,
which ``main.py`` calls exactly once.  Adding a tool therefore means adding a
function and one entry to its module's ``TOOLS``; nothing in ``main.py``
changes, and no wrapper repeats the signature.

Read :mod:`word_document_server.tools.v2.registry` for the report shape every
one of these tools answers with, and for the error codes they map engine
failures to.
"""

from word_document_server.tools.v2.registry import (
    ToolSpec,
    discover_tool_specs,
    register_v2_tools,
    v2_tools,
)

__all__ = ["ToolSpec", "discover_tool_specs", "register_v2_tools", "v2_tools"]
