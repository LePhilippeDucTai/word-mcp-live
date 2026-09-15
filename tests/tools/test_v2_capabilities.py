"""``doc_capabilities``: what this process, on this machine, can actually reach.

What is pinned here is not the platform-detection logic itself -- that is a
thin wrapper over :func:`platform.system` and :func:`shutil.which` -- but that
the tool is discoverable and registered (D-024: it was not, until this part)
and that its counts are derived from
:data:`~word_document_server.tools.platforms.PLATFORMS`, the same source
``TOOLS.md`` and the registry consistency test read, rather than a second,
independently maintained tally.

``call`` is redefined here, the way ``test_v2_batch.py`` and
``test_v2_text.py`` do (see J04-P7's note on not importing test helpers across
modules).
"""

from __future__ import annotations

import asyncio
from typing import Any

from word_document_server.tools.platforms import PLATFORMS
from word_document_server.tools.v2.registry import (
    discover_tool_specs,
    register_v2_tools,
    v2_tools,
)


def call(name: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Call a V2 tool the way an MCP client reaches it, and return its report."""
    return asyncio.run(v2_tools()[name](*args, **kwargs))


def test_capabilities_is_discovered_and_registered():
    from fastmcp import FastMCP

    discovered = {spec.name for spec in discover_tool_specs()}
    assert "doc_capabilities" in discovered

    server = FastMCP("test")
    registered = register_v2_tools(server)
    assert "doc_capabilities" in registered


def test_the_report_counts_what_platforms_lists():
    report = call("doc_capabilities")

    assert report["status"] == "ok"
    assert report["tool_count"] == len(PLATFORMS)

    v2_count = sum(1 for name in PLATFORMS if name.startswith("doc_"))
    assert report["tool_families"]["v2_semantic"] == v2_count

    assert report["note"]
