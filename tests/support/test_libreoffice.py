"""Tests for tests/support/libreoffice.py.

LibreOffice is optional tooling (see PLAN.md Risks): every test that
actually shells out to `soffice` is guarded by `requires_libreoffice` so the
suite stays green in environments without it. Tests that only exercise error
handling (missing binary, missing source, missing fixture) do not need the
guard because they do not depend on soffice being present.
"""

import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import url2pathname

import pytest

from tests.support.libreoffice import (
    convert,
    fodt_to_docx,
    opens_in_libreoffice,
    requires_libreoffice,
    soffice_path,
)

_FIXTURE_NAMES = ["rich", "tables_lists", "notes_fields"]


def _path_from_uri(uri: str) -> Path:
    """The path a `file://` URI names, on both path families.

    `convert` builds these with `Path.as_uri()`, which on Windows yields
    `file:///C:/...`. Stripping the `file://` prefix would leave `/C:/...`,
    which is not a path there -- and an `assert not ....exists()` on it would
    pass for the wrong reason, never able to fail.
    """
    return Path(url2pathname(urlparse(uri).path))


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


def _fake_run_that_writes_output(log):
    """Build a fake `subprocess.run` that records the profile URI it saw.

    Verifies the profile directory exists at call time (proving `convert`
    creates it before invoking soffice), appends the `-env:UserInstallation=`
    value to `log`, writes a non-empty output file, and reports success.
    """

    def _run(argv, capture_output, text, timeout, check):
        profile_args = [a for a in argv if a.startswith("-env:UserInstallation=")]
        assert len(profile_args) == 1
        profile_uri = profile_args[0].removeprefix("-env:UserInstallation=")
        profile_path = _path_from_uri(profile_uri)
        assert profile_path.is_dir()
        log.append(profile_uri)

        outdir = Path(argv[argv.index("--outdir") + 1])
        src = Path(argv[-1])
        fmt = argv[argv.index("--convert-to") + 1]
        (outdir / f"{src.stem}.{fmt}").write_bytes(b"fake docx content")
        return subprocess.CompletedProcess(argv, returncode=0)

    return _run


def test_convert_gives_each_call_its_own_profile(monkeypatch, tmp_path):
    monkeypatch.setattr("tests.support.libreoffice.soffice_path", lambda: "soffice")
    log: list[str] = []
    monkeypatch.setattr(
        "tests.support.libreoffice.subprocess.run", _fake_run_that_writes_output(log)
    )

    src = tmp_path / "source.fodt"
    src.write_text("fake source", encoding="utf-8")
    outdirs = [tmp_path / f"out{i}" for i in range(3)]

    with ThreadPoolExecutor(max_workers=3) as pool:
        results = list(pool.map(lambda outdir: convert(src, "docx", outdir), outdirs))

    assert len(results) == 3
    assert len(set(log)) == 3
    for profile_uri in log:
        profile_path = _path_from_uri(profile_uri)
        assert not profile_path.exists()


def test_convert_removes_its_profile_when_soffice_writes_nothing(monkeypatch, tmp_path):
    monkeypatch.setattr("tests.support.libreoffice.soffice_path", lambda: "soffice")
    seen_profiles: list[Path] = []

    def _run(argv, capture_output, text, timeout, check):
        profile_uri = next(
            a for a in argv if a.startswith("-env:UserInstallation=")
        ).removeprefix("-env:UserInstallation=")
        seen_profiles.append(_path_from_uri(profile_uri))
        # Writes nothing, mimicking soffice attaching to another instance.
        return subprocess.CompletedProcess(argv, returncode=0)

    monkeypatch.setattr("tests.support.libreoffice.subprocess.run", _run)

    src = tmp_path / "source.fodt"
    src.write_text("fake source", encoding="utf-8")

    with pytest.raises(RuntimeError, match="did not produce"):
        convert(src, "docx", tmp_path / "out")

    assert len(seen_profiles) == 1
    assert not seen_profiles[0].exists()


@requires_libreoffice
@pytest.mark.libreoffice
def test_concurrent_conversions_of_the_same_source_all_produce_output(tmp_path):
    """Regression test for the shared-profile bug (D-013).

    Before the fix, concurrent `soffice` invocations attached to a single
    shared instance keyed by the default profile, so all but one call
    reported success without producing output. Flaky/red before the fix,
    deterministically green after it.
    """
    outdirs = [tmp_path / f"out{i}" for i in range(3)]

    with ThreadPoolExecutor(max_workers=3) as pool:
        results = list(
            pool.map(lambda outdir: fodt_to_docx("rich", outdir), outdirs)
        )

    assert len(results) == 3
    for docx_path in results:
        assert docx_path.exists()
        assert docx_path.stat().st_size > 0


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
