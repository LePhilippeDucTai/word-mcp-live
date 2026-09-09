"""Atomic-write invariant introduced when the save-hook monkey-patch
(``utils/save_utils.py``) was retired.

``save_utils.install_save_hook`` used to wrap ``Document.save`` globally to
re-inject parts python-docx strips; it never guaranteed anything about a
write interrupted mid-flight. It is gone now: writes to a live ``.docx`` go
either through the engine's own save path or, for the two tools that write
raw bytes outside python-docx (encrypting and decrypting a document in
``protection_tools``), through
:func:`word_document_server.engine.package.atomic_write_bytes`. What is
pinned here is the invariant that replaces the old best-effort
"catch the exception and try to rewrite the original bytes" pattern: a
temp-file-then-``os.replace`` write either fully succeeds or leaves the
destination byte-for-byte as it was, and never leaves a temporary file
behind.
"""

from __future__ import annotations

import asyncio
import inspect
import os
from pathlib import Path
from typing import Any

import pytest

from word_document_server.engine.package import atomic_write_bytes
from word_document_server.tools import protection_tools
from word_document_server.tools.protection_tools import (
    protect_document,
    unprotect_document,
)


def _run(coro: Any) -> Any:
    """Drive an ``async`` tool call to completion."""
    return asyncio.run(coro)


def _stray_files(directory: Path, keep: Path) -> list[Path]:
    """Every file in `directory` other than `keep` -- a leftover temp file
    from an interrupted atomic write would show up here."""
    return [p for p in directory.iterdir() if p != keep]


# --------------------------------------------------------------------------------------
# engine.package.atomic_write_bytes itself
# --------------------------------------------------------------------------------------


def test_atomic_write_bytes_replaces_content_on_success(tmp_path: Path) -> None:
    target = tmp_path / "file.bin"
    target.write_bytes(b"before")

    atomic_write_bytes(target, b"after")

    assert target.read_bytes() == b"after"
    assert _stray_files(tmp_path, target) == []


def test_atomic_write_bytes_leaves_original_intact_when_the_write_is_interrupted(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    target = tmp_path / "file.bin"
    target.write_bytes(b"original")

    def _boom(*args: object, **kwargs: object) -> None:
        raise OSError("simulated disk failure")

    monkeypatch.setattr(os, "fsync", _boom)

    with pytest.raises(OSError):
        atomic_write_bytes(target, b"new content")

    assert target.read_bytes() == b"original"
    assert _stray_files(tmp_path, target) == []


def test_atomic_write_bytes_leaves_no_temp_file_when_the_rename_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    target = tmp_path / "file.bin"
    target.write_bytes(b"original")

    def _boom(*args: object, **kwargs: object) -> None:
        raise OSError("simulated rename failure")

    monkeypatch.setattr(os, "replace", _boom)

    with pytest.raises(OSError):
        atomic_write_bytes(target, b"new content")

    assert target.read_bytes() == b"original"
    assert _stray_files(tmp_path, target) == []


# --------------------------------------------------------------------------------------
# protect_document / unprotect_document: encrypting and decrypting write raw
# bytes outside python-docx, through atomic_write_bytes.
# --------------------------------------------------------------------------------------


def test_protection_tools_no_longer_writes_through_a_plain_open() -> None:
    """The old ``open(filename, "wb")`` overwrite -- and its best-effort,
    try/except restore of the previous bytes -- is gone; both directions now
    delegate to ``atomic_write_bytes``."""
    source = inspect.getsource(protection_tools)
    assert 'open(filename, "wb")' not in source


def test_protect_document_leaves_original_intact_when_the_write_is_interrupted(
    tmp_path: Path, fixture_docx, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = fixture_docx("simple")
    before = source.read_bytes()

    def _boom(*args: object, **kwargs: object) -> None:
        raise OSError("simulated disk failure")

    monkeypatch.setattr(protection_tools, "atomic_write_bytes", _boom)

    result = _run(protect_document(str(source), "secret123"))

    assert "unchanged" in result.lower()
    assert source.read_bytes() == before
    assert _stray_files(tmp_path, source) == []


def test_unprotect_document_leaves_encrypted_file_intact_when_the_write_is_interrupted(
    tmp_path: Path, fixture_docx, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = fixture_docx("simple")
    _run(protect_document(str(source), "secret123"))
    encrypted = source.read_bytes()

    def _boom(*args: object, **kwargs: object) -> None:
        raise OSError("simulated disk failure")

    monkeypatch.setattr(protection_tools, "atomic_write_bytes", _boom)

    result = _run(unprotect_document(str(source), "secret123"))

    assert "unchanged" in result.lower()
    assert source.read_bytes() == encrypted
    assert _stray_files(tmp_path, source) == []


# --------------------------------------------------------------------------------------
# The save-hook monkey-patch itself is gone.
# --------------------------------------------------------------------------------------


def test_save_utils_module_is_removed() -> None:
    with pytest.raises(ModuleNotFoundError):
        import word_document_server.utils.save_utils  # noqa: F401
