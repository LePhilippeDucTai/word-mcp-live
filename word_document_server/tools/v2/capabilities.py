"""Runtime capability report for this server instance.

Where :data:`~word_document_server.tools.platforms.PLATFORMS` says what each
tool supports in the abstract, :func:`doc_capabilities` says what *this*
process, on *this* machine, can actually reach: which OS it runs on, whether
LibreOffice is on ``PATH`` (the ``convert_to_pdf`` fallback on Linux/macOS),
and how many tools fall in each family.

Word availability is only known to live tools: a live tool probes Word itself
at call time (COM on Windows, AppleScript/JXA on macOS) and reports failure
there. This function does not launch Word to check, so it cannot say whether
Word is actually installed and running on this machine -- only a live tool
call can.
"""

from __future__ import annotations

import platform
import shutil

from word_document_server.tools.platforms import PLATFORMS

#: Tools with no Linux entry in PLATFORMS are the Word-automation ("live")
#: ones; everything else is either a plain python-docx tool or a `doc_*` V2
#: tool, told apart by name below.
_LIVE_TOOL_NAMES = frozenset(
    name for name, platforms in PLATFORMS.items() if "linux" not in platforms
)

#: Prefix shared by every V2 semantic-engine tool.
_V2_TOOL_PREFIX = "doc_"


def _current_platform() -> str:
    """This process's OS, in the vocabulary :data:`PLATFORMS` uses."""
    system = platform.system()
    if system == "Windows":
        return "windows"
    if system == "Darwin":
        return "macos"
    return "linux"


def _libreoffice_available() -> bool:
    """Whether a LibreOffice CLI (``soffice`` or ``libreoffice``) is on PATH.

    This is the fallback ``convert_to_pdf`` uses on Linux/macOS (see
    ``tools/extended_document_tools.py``); Windows uses ``docx2pdf`` / Word
    instead and does not need it.
    """
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def _tool_family(name: str) -> str:
    """Which family `name` belongs to, for the counters below."""
    if name.startswith(_V2_TOOL_PREFIX):
        return "v2_semantic"
    if name in _LIVE_TOOL_NAMES:
        return "live"
    return "docx"


def doc_capabilities() -> dict[str, object]:
    """Platform, LibreOffice detection, and tool family counts for this process.

    Returns:
        A mapping with ``platform`` (``"linux"``, ``"windows"`` or
        ``"macos"``), ``libreoffice_available`` (bool), ``tool_families``
        (family name -> tool count), ``tool_count`` (total registered tools),
        and ``note`` -- the caveat in the module docstring, so a caller reading
        only this dict still sees it.
    """
    families: dict[str, int] = {}
    for name in PLATFORMS:
        family = _tool_family(name)
        families[family] = families.get(family, 0) + 1

    return {
        "platform": _current_platform(),
        "libreoffice_available": _libreoffice_available(),
        "tool_families": families,
        "tool_count": len(PLATFORMS),
        "note": "Word availability is only known to live tools",
    }
