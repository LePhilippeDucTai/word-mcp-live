"""Discovery, registration and the one report shape of the V2 ``doc_*`` tools.

A V2 tool is a plain Python function.  It takes the arguments an agent sends,
it returns a ``dict``, and it raises when it cannot do what it was asked.  It
does not open a lock, it does not catch anything, it does not build an
envelope: this module does all three, once, for every tool, so that the whole
surface answers in the same shape whatever happens.

Registration
------------
Each module of :mod:`word_document_server.tools.v2` exports ``TOOLS``, a list of
:class:`ToolSpec`.  :func:`register_v2_tools` imports every module of the
package, collects those lists and registers each function *as it stands*:
FastMCP derives the input schema from the signature and the description from the
docstring, so a parameter is added by adding a parameter.  ``main.py`` calls
:func:`register_v2_tools` once and names no tool.

The report
----------
Every call that returns answers a mapping with at least these four keys::

    {"status": "ok", "dry_run": false, "changes": [...], "warnings": [...]}

``changes``
    one entry per paragraph the call touched, ``{"story", "paragraph",
    "before", "after"}``: `story` is the real story id (never the ``"body"``
    alias, see D-006), `paragraph` is the **V2 paragraph index** -- the one
    space :func:`~word_document_server.engine.locators.resolve` and
    :func:`~word_document_server.engine.inspect.inspect` number, per D-016 --
    or ``null`` for a paragraph in a table cell or a text box, which has no
    such index.  `before` and `after` are the paragraph's visible text on
    either side of the operation.
``warnings``
    things the caller should know and that did not stop the call: an edit that
    changed nothing, a range that covered no run.
``dry_run``
    what the caller asked for, echoed back, so a report can never be mistaken
    for the record of a write that did not happen.

A tool adds its own keys on top (``document``, ``matches``, ``runs``,
``saved``, ...); the four above are filled in here when the tool leaves them
out, and ``status`` is not the tool's to set.

Failures
--------
A raised exception becomes::

    {"status": "error", "code": "...", "message": "..."}

`code` is machine-readable and stable; `message` is for a human and may change.
Locator codes -- ``invalid``, ``not_found``, ``ambiguous``, ``stale_anchor``
(D-022) -- and the ``reason`` of an
:class:`~word_document_server.engine.errors.UnsupportedRange` are passed through
untouched, because they are already the vocabulary the engine settled on.  The
remaining engine errors derive their code from their class name
(``InvalidText`` -> ``invalid_text``).  A missing file is ``file_not_found``,
any other OS failure is ``io_error``, and an argument the tool refuses is
``invalid_argument``.

Anything *else* -- an ``AttributeError``, a ``KeyError`` -- is left to
propagate: it is a bug in this server, not an answer to give an agent, and
burying it in an error dict would make it look like a document problem.

Paths
-----
A V2 tool opens the path it is given, with no extension fixing: the older tools
turn ``report`` into ``report.docx``, which is convenient until it turns a typo
into a second document.  Here a path that is not a package fails saying so, and
the file lock below is therefore taken on exactly the file that will be written.

Locking
-------
A tool that takes a ``filename`` parameter runs under the per-file
:func:`~word_document_server.utils.file_utils.get_file_lock`, so two concurrent
calls on the same document cannot interleave their read-modify-write; calls on
different documents still run in parallel.
"""

from __future__ import annotations

import functools
import importlib
import inspect
import pkgutil
import re
from collections.abc import Callable, Iterator, Mapping
from dataclasses import dataclass, field
from pathlib import Path
from types import ModuleType
from typing import Any

from mcp.types import ToolAnnotations

from word_document_server.engine.errors import (
    EngineError,
    LocatorError,
    UnsupportedRange,
)
from word_document_server.utils.file_utils import get_file_lock

__all__ = [
    "ToolSpec",
    "discover_tool_specs",
    "error_code",
    "error_report",
    "register_v2_tools",
    "v2_tools",
]

#: Name of the parameter a tool uses for the document it works on.  A tool that
#: has one is serialized on that file; a tool that has none is not.
FILENAME_PARAMETER = "filename"

#: Exceptions turned into an error report.  Everything else propagates -- see
#: the module docstring.
_REPORTED: tuple[type[Exception], ...] = (EngineError, OSError, ValueError, TypeError)

#: Splits a class name into words, to derive a code from it.
_CLASS_WORD = re.compile(r"[A-Z]+(?![a-z])|[A-Z][a-z0-9]*|[a-z0-9]+")


@dataclass(frozen=True)
class ToolSpec:
    """One V2 tool: the function, and what MCP should say about it.

    Attributes:
        fn: the implementation.  Its name is the tool name, its signature is the
            input schema and its docstring is the description, so none of the
            three is repeated here.
        annotations: the MCP behaviour hints (``readOnlyHint``,
            ``destructiveHint``, a human title).
        tags: free-form labels, used to group the surface into families.
    """

    fn: Callable[..., Any]
    annotations: ToolAnnotations | None = None
    tags: frozenset[str] = field(default_factory=frozenset)

    @property
    def name(self) -> str:
        """The tool name, which is the function's own name."""
        return self.fn.__name__


# --------------------------------------------------------------------------
# Reports
# --------------------------------------------------------------------------


def _derived_code(exc: Exception) -> str:
    """``InvalidText`` -> ``invalid_text``: a code from an exception class name."""
    name = type(exc).__name__
    words = _CLASS_WORD.findall(name)
    return "_".join(word.lower() for word in words) if words else name.lower()


def error_code(exc: Exception) -> str:
    """The stable code an agent branches on for `exc`.

    See the module docstring for the mapping and for why locator codes and
    :attr:`UnsupportedRange.reason` are passed through unchanged.
    """
    if isinstance(exc, LocatorError):
        return exc.code
    if isinstance(exc, UnsupportedRange):
        return exc.reason
    if isinstance(exc, FileNotFoundError):
        return "file_not_found"
    if isinstance(exc, OSError):
        return "io_error"
    if isinstance(exc, EngineError):
        return _derived_code(exc)
    return "invalid_argument"


def error_report(exc: Exception) -> dict[str, Any]:
    """The ``{"status": "error", "code", "message"}`` answer for `exc`."""
    return {"status": "error", "code": error_code(exc), "message": str(exc)}


def _success_report(result: object, tool: str) -> dict[str, Any]:
    """Fill the four mandatory keys around what `tool` returned.

    Raises:
        TypeError: if `tool` did not return a mapping.  That is a bug in the
            tool, not an answer to dress up as one.
    """
    if not isinstance(result, Mapping):
        raise TypeError(
            f"V2 tool {tool!r} must return a mapping, got {type(result).__name__}"
        )
    report: dict[str, Any] = {
        "status": "ok",
        "dry_run": bool(result.get("dry_run", False)),
        "changes": list(result.get("changes", ())),
        "warnings": list(result.get("warnings", ())),
    }
    for key, value in result.items():
        if key not in report and key != "status":
            report[key] = value
    return report


# --------------------------------------------------------------------------
# Wrapping
# --------------------------------------------------------------------------


def _locked_file(signature: inspect.Signature, args: tuple[Any, ...], kwargs: dict[str, Any]) -> str | None:
    """The path to serialize this call on, or ``None`` when there is none."""
    if FILENAME_PARAMETER not in signature.parameters:
        return None
    bound = signature.bind(*args, **kwargs)
    bound.apply_defaults()
    value = bound.arguments.get(FILENAME_PARAMETER)
    return None if value is None else str(value)


def _tool_callable(spec: ToolSpec) -> Callable[..., Any]:
    """Wrap `spec.fn` into the coroutine FastMCP registers.

    The wrapper keeps the wrapped signature and docstring -- that is what the
    schema and the description are derived from -- and adds the three things
    every V2 tool shares: the file lock, the report envelope and the error
    mapping.
    """
    fn = spec.fn
    name = spec.name
    signature = inspect.signature(fn)
    is_async = inspect.iscoroutinefunction(fn)

    async def _invoke(args: tuple[Any, ...], kwargs: dict[str, Any]) -> dict[str, Any]:
        result = fn(*args, **kwargs)
        if is_async:
            result = await result
        return _success_report(result, name)

    @functools.wraps(fn)
    async def call(*args: Any, **kwargs: Any) -> dict[str, Any]:
        try:
            filename = _locked_file(signature, args, kwargs)
            if filename is None:
                return await _invoke(args, kwargs)
            async with get_file_lock(filename):
                return await _invoke(args, kwargs)
        except _REPORTED as exc:
            return error_report(exc)

    # Set explicitly rather than relying on ``__wrapped__``: the schema must be
    # the tool's signature whatever introspection FastMCP happens to use.
    call.__signature__ = signature  # type: ignore[attr-defined]
    return call


# --------------------------------------------------------------------------
# Discovery
# --------------------------------------------------------------------------


def _modules() -> Iterator[ModuleType]:
    """Import every public module of this package, in a stable order."""
    here = Path(__file__)
    for info in sorted(pkgutil.iter_modules([str(here.parent)]), key=lambda item: item.name):
        if info.name.startswith("_") or info.name == here.stem:
            continue
        yield importlib.import_module(f"{__package__}.{info.name}")


def discover_tool_specs() -> list[ToolSpec]:
    """Every :class:`ToolSpec` the V2 package exports, in module order.

    A module without a ``TOOLS`` list is simply not a tool module.

    Raises:
        TypeError: if a ``TOOLS`` entry is not a :class:`ToolSpec`, or if a tool
            function has no docstring -- the docstring *is* the description an
            agent reads, so a missing one is a broken tool, not a style issue.
        ValueError: if two modules export a tool of the same name.
    """
    specs: list[ToolSpec] = []
    origin: dict[str, str] = {}
    for module in _modules():
        exported = getattr(module, "TOOLS", None)
        if exported is None:
            continue
        for spec in exported:
            if not isinstance(spec, ToolSpec):
                raise TypeError(
                    f"{module.__name__}.TOOLS must hold ToolSpec instances, got "
                    f"{type(spec).__name__}"
                )
            if not (spec.fn.__doc__ or "").strip():
                raise TypeError(
                    f"V2 tool {spec.name!r} has no docstring; its docstring is the "
                    "description MCP clients are given"
                )
            if spec.name in origin:
                raise ValueError(
                    f"two V2 modules export a tool named {spec.name!r}: "
                    f"{origin[spec.name]} and {module.__name__}"
                )
            origin[spec.name] = module.__name__
            specs.append(spec)
    return specs


def v2_tools() -> dict[str, Callable[..., Any]]:
    """The V2 tools by name, wrapped exactly as they are registered.

    The callables are coroutine functions returning the report described in the
    module docstring.  This is the seam tests and in-process callers use, so
    that what they exercise is what an MCP client reaches.
    """
    return {spec.name: _tool_callable(spec) for spec in discover_tool_specs()}


def register_v2_tools(mcp: Any) -> list[str]:
    """Register every discovered V2 tool on `mcp` and return their names.

    Args:
        mcp: the FastMCP server.

    Returns:
        The names registered, in registration order -- what a consistency check
        or a documentation generator needs without re-running the discovery.
    """
    registered: list[str] = []
    for spec in discover_tool_specs():
        mcp.tool(
            _tool_callable(spec),
            name=spec.name,
            annotations=spec.annotations,
            tags=set(spec.tags) or None,
        )
        registered.append(spec.name)
    return registered
