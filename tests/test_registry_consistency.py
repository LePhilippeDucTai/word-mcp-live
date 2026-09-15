"""Two invariants the tool registry must hold.

*every registered tool has a platform entry*
    :data:`~word_document_server.tools.platforms.PLATFORMS` is what
    ``scripts/gen_tools_md.py`` and
    :func:`~word_document_server.tools.v2.capabilities.doc_capabilities` read
    to describe a tool's availability. A tool missing from it would be
    silently dropped from ``TOOLS.md`` and the capability counts; an entry for
    a tool that no longer exists would silently inflate them.

*every wrapper in main.py forwards every parameter of its implementation*
    ``register_tools()`` in ``main.py`` defines one small wrapper function per
    legacy tool, whose body is ``return <implementation>(...)``. A parameter
    added to the implementation but not threaded through the wrapper's own
    signature and call is silently dropped at the MCP boundary -- exactly the
    bug this test reproduces for ``word_live_modify_table``'s
    ``scrub_orphans`` (J04-P3). The V2 ``doc_*`` tools are registered directly
    by :func:`~word_document_server.tools.v2.registry.register_v2_tools` and
    have no such wrapper in ``main.py``, so they are out of scope for the
    second check -- there is nothing to compare a wrapper against.
"""

from __future__ import annotations

import ast
import asyncio
import inspect
import textwrap
from typing import Any

import pytest

import word_document_server.main as main_module
from word_document_server.main import mcp, register_tools
from word_document_server.tools.platforms import PLATFORMS


def _registered_tools() -> list[Any]:
    """Every ``FunctionTool`` the server registers, in registration order.

    ``register_tools()`` is safe to call more than once (FastMCP overwrites
    same-named tools with a warning), so every test in this module can call
    this independently without coordinating with the others.
    """
    register_tools()
    return asyncio.run(mcp.list_tools())


def _main_py_wrappers() -> list[Any]:
    """The subset of registered tools whose ``fn`` is a wrapper defined in main.py.

    The V2 ``doc_*`` tools' ``fn`` keeps the ``__module__`` of the function
    defined in ``tools/v2/text.py`` (``functools.wraps`` in
    ``tools/v2/registry.py`` copies it across), so filtering on
    ``__module__`` cleanly separates the two kinds of tool.
    """
    return [t for t in _registered_tools() if t.fn.__module__ == "word_document_server.main"]


def test_every_registered_tool_has_a_platform_entry():
    names = {t.name for t in _registered_tools()}
    missing = sorted(names - set(PLATFORMS))
    assert not missing, f"tools registered but missing from PLATFORMS: {missing}"
    stale = sorted(set(PLATFORMS) - names)
    assert not stale, f"PLATFORMS entries for tools that are no longer registered: {stale}"


def _resolve_wrapper_call(fn: Any) -> tuple[ast.FunctionDef | ast.AsyncFunctionDef, ast.Call, Any]:
    """The wrapper's own AST def, its ``return`` call, and the callable it targets."""
    source = textwrap.dedent(inspect.getsource(fn))
    tree = ast.parse(source)
    funcdef = tree.body[0]
    assert isinstance(funcdef, (ast.FunctionDef, ast.AsyncFunctionDef)), (
        f"{fn.__qualname__}: expected a single function definition"
    )
    returns = [stmt for stmt in funcdef.body if isinstance(stmt, ast.Return)]
    assert len(returns) == 1, f"{fn.__qualname__}: expected exactly one return statement"
    call = returns[0].value
    assert isinstance(call, ast.Call), f"{fn.__qualname__}: return value is not a direct call"
    # Safe to eval: `call.func` is source *we* wrote in main.py (a dotted name
    # such as `live_tools.word_live_modify_table`), evaluated against main.py's
    # own module globals -- not attacker-controlled input.
    target = eval(compile(ast.Expression(call.func), "<ast>", "eval"), vars(main_module))
    return funcdef, call, target


def _forwarded_names(call: ast.Call) -> set[str]:
    """Every identifier passed as a positional or ``name=name`` keyword argument."""
    names = {arg.id for arg in call.args if isinstance(arg, ast.Name)}
    names |= {
        kw.value.id
        for kw in call.keywords
        if kw.arg is not None and isinstance(kw.value, ast.Name)
    }
    return names


@pytest.mark.parametrize("tool", _main_py_wrappers(), ids=lambda t: t.name)
def test_wrapper_forwards_every_implementation_parameter(tool: Any):
    funcdef, call, target = _resolve_wrapper_call(tool.fn)
    wrapper_params = {arg.arg for arg in funcdef.args.args}
    impl_params = set(inspect.signature(target).parameters)

    missing_in_signature = impl_params - wrapper_params
    assert not missing_in_signature, (
        f"{tool.name}: implementation parameter(s) {sorted(missing_in_signature)} "
        "are not in the wrapper's own signature in main.py"
    )

    missing_in_call = impl_params - _forwarded_names(call)
    assert not missing_in_call, (
        f"{tool.name}: implementation parameter(s) {sorted(missing_in_call)} "
        "are not forwarded in the wrapper's call in main.py"
    )
