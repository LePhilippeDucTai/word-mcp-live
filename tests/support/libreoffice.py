"""Helpers for driving LibreOffice as an optional fixture producer.

LibreOffice is never a source of truth for this project (see PLAN.md Risks):
it only produces .docx fixtures from committed .fodt sources, and can
optionally confirm that a document reopens without error. Every test that
depends on it must be skipped, never failed, when `soffice` is unavailable.
"""

import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

_SOFFICE_CANDIDATES = ("soffice", "libreoffice")

_FIXTURES_DIR = Path(__file__).resolve().parent.parent / "fixtures" / "odt"


def soffice_path() -> str | None:
    """Return the path to a usable soffice/libreoffice binary, or None."""
    for name in _SOFFICE_CANDIDATES:
        found = shutil.which(name)
        if found:
            return found
    return None


# Skip (never fail) any test that needs LibreOffice when it is not installed.
requires_libreoffice = pytest.mark.skipif(
    soffice_path() is None, reason="LibreOffice (soffice/libreoffice) is not installed"
)


def convert(src: Path, fmt: str, outdir: Path, timeout: int = 120) -> Path:
    """Convert `src` to `fmt` with LibreOffice headless, writing into `outdir`.

    Each call gets its own disposable user profile via
    `-env:UserInstallation=`. Without this, concurrent `soffice` invocations
    on the same machine share `~/.config/libreoffice`: soffice detects the
    existing instance for that profile, attaches to it, and exits 0 without
    converting anything, so a caller other than the one that "won" the
    instance gets a missing output file. A fresh, unique profile per call
    forces each conversion to start its own isolated instance, so concurrent
    conversions cannot collide. The profile directory is removed again once
    the call is done, on every exit path.

    Returns the path to the produced file. Raises RuntimeError if soffice is
    missing, the conversion process fails, or the output file is missing or
    empty.
    """
    soffice = soffice_path()
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice/libreoffice) is not installed")

    src = Path(src)
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    profile_dir = tempfile.mkdtemp(prefix="soffice-profile-")
    try:
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                f"-env:UserInstallation={Path(profile_dir).as_uri()}",
                "--convert-to", fmt,
                "--outdir", str(outdir),
                str(src),
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"soffice conversion of {src} to {fmt} failed "
                f"(exit {result.returncode}): {result.stderr or result.stdout}"
            )

        out_path = outdir / f"{src.stem}.{fmt}"
        if not out_path.exists():
            raise RuntimeError(
                f"soffice reported success but did not produce {out_path} "
                f"(stdout: {result.stdout})"
            )
        if out_path.stat().st_size == 0:
            raise RuntimeError(f"soffice produced an empty file: {out_path}")
        return out_path
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def fodt_to_docx(name: str, cache_dir: Path) -> Path:
    """Convert `tests/fixtures/odt/<name>.fodt` to .docx, cached per session.

    `cache_dir` is expected to be a directory created once per test session
    (typically via `tmp_path_factory`). Repeated calls with the same `name`
    and `cache_dir` reuse the cached file instead of reconverting.
    """
    cache_dir = Path(cache_dir)
    cached = cache_dir / f"{name}.docx"
    if cached.exists():
        return cached

    src = _FIXTURES_DIR / f"{name}.fodt"
    if not src.exists():
        raise FileNotFoundError(f"fodt fixture not found: {src}")

    produced = convert(src, "docx", cache_dir)
    if produced != cached:
        produced.replace(cached)
    return cached


def opens_in_libreoffice(path: Path) -> bool:
    """Return True if LibreOffice can open `path` and export it to a non-empty PDF.

    This is an optional round-trip check, not a validator of document
    correctness: a False result only means LibreOffice could not reopen the
    file, not that the file is invalid OOXML.
    """
    if soffice_path() is None:
        return False
    try:
        with tempfile.TemporaryDirectory() as tmp:
            pdf = convert(Path(path), "pdf", Path(tmp))
            return pdf.exists() and pdf.stat().st_size > 0
    except RuntimeError:
        return False
