"""Shared pytest fixtures for the whole test suite.

Fixture documents are never committed as binaries: they are rebuilt from
:mod:`tests.fixtures.builders` and written to a temporary directory on demand.
"""

from __future__ import annotations

from collections.abc import Callable
from functools import cache
from pathlib import Path

import pytest

from tests.fixtures.builders import build


@cache
def _cached_build(name: str) -> bytes:
    """Build `name` once per session; builders are deterministic, so caching is safe."""
    return build(name)


def write_fixture_docx(name: str, directory: Path) -> Path:
    """Write the fixture `name` into `directory` and return the resulting path."""
    path = Path(directory) / f"{name}.docx"
    path.write_bytes(_cached_build(name))
    return path


@pytest.fixture
def fixture_docx(tmp_path: Path) -> Callable[..., Path]:
    """Return a factory writing a named fixture document and returning its path.

    Usage::

        def test_something(fixture_docx):
            path = fixture_docx("tracked_changes")

    The file lands in the test's ``tmp_path`` unless another `directory` is given.
    """

    def _fixture_docx(name: str, directory: Path | None = None) -> Path:
        return write_fixture_docx(name, tmp_path if directory is None else directory)

    return _fixture_docx
