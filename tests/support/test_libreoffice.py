"""Tests for tests/support/libreoffice.py.

LibreOffice is optional tooling (see PLAN.md Risks): every test that
actually shells out to `soffice` is guarded by `requires_libreoffice` so the
suite stays green in environments without it. Tests that only exercise error
handling (missing binary, missing source, missing fixture) do not need the
guard because they do not depend on soffice being present.
"""

import shutil

import pytest

from tests.support.libreoffice import (
    convert,
    fodt_to_docx,
    opens_in_libreoffice,
    requires_libreoffice,
    soffice_path,
)

_FIXTURE_NAMES = ["rich", "tables_lists", "notes_fields"]


@pytest.fixture(scope="session")
def odt_cache_dir(tmp_path_factory):
    """A session-scoped directory so .fodt -> .docx conversion runs at most once."""
    return tmp_path_factory.mktemp("odt_docx_cache")


def test_soffice_path_matches_shutil_which():
    expected = shutil.which("soffice") or shutil.which("libreoffice")
    assert soffice_path() == expected


def test_requires_libreoffice_marker_matches_availability():
    condition = requires_libreoffice.args[0]
    assert condition == (soffice_path() is None)


def test_convert_raises_when_binary_missing(monkeypatch, tmp_path):
    monkeypatch.setattr("tests.support.libreoffice.soffice_path", lambda: None)
    with pytest.raises(RuntimeError, match="not installed"):
        convert(tmp_path / "source.fodt", "docx", tmp_path)


def test_fodt_to_docx_unknown_fixture_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        fodt_to_docx("does-not-exist", tmp_path)


def test_opens_in_libreoffice_false_for_missing_file(tmp_path):
    # A nonexistent path can never open, with or without LibreOffice
    # installed: exercises both the "no soffice" short-circuit and the
    # "soffice ran but produced nothing" branch depending on the machine.
    assert opens_in_libreoffice(tmp_path / "missing.docx") is False


@requires_libreoffice
@pytest.mark.libreoffice
def test_convert_raises_on_missing_source(tmp_path):
    with pytest.raises(RuntimeError):
        convert(tmp_path / "missing.fodt", "docx", tmp_path)


@requires_libreoffice
@pytest.mark.libreoffice
@pytest.mark.parametrize("name", _FIXTURE_NAMES)
def test_fodt_to_docx_produces_nonempty_docx(name, odt_cache_dir):
    docx_path = fodt_to_docx(name, odt_cache_dir)
    assert docx_path.exists()
    assert docx_path.suffix == ".docx"
    assert docx_path.stat().st_size > 0


@requires_libreoffice
@pytest.mark.libreoffice
def test_fodt_to_docx_is_cached_across_calls(odt_cache_dir):
    first = fodt_to_docx("rich", odt_cache_dir)
    mtime_before = first.stat().st_mtime

    second = fodt_to_docx("rich", odt_cache_dir)

    assert second == first
    assert second.stat().st_mtime == mtime_before


@requires_libreoffice
@pytest.mark.libreoffice
def test_opens_in_libreoffice_true_for_valid_docx(odt_cache_dir):
    docx_path = fodt_to_docx("rich", odt_cache_dir)
    assert opens_in_libreoffice(docx_path) is True
