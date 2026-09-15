"""The distribution's version and name are declared in four files; they must agree.

Nothing generates these four declarations from one another, so they drift silently:
before this guard existed ``pyproject.toml`` said 1.6.20 (the version the *upstream*
project had last published, inherited through the fork and never ours),
``manifest.json`` and ``server.json`` said 1.3.0, and
``word_document_server.__version__`` said 1.2.0 -- three different answers to "what
version is this?", none of them checkable by running anything.

The name has the same shape of problem and a sharper failure mode: ``server.json``'s
``identifier`` and the README's ``mcp-name`` trailer are what an MCP client and the MCP
registry resolve to install, so a rename that misses one of them points users at a
different distribution than the one this repository builds.
"""

from __future__ import annotations

import json
import tomllib
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

#: The name every declaration below must agree on.
DISTRIBUTION = "word-mcp-semantic"

#: The ``<!-- mcp-name: ... -->`` trailer README.md must end with.
MCP_NAME = f"io.github.LePhilippeDucTai/{DISTRIBUTION}"


def _pyproject() -> dict:
    return tomllib.loads((REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8"))


def _json(name: str) -> dict:
    return json.loads((REPO_ROOT / name).read_text(encoding="utf-8"))


def _declared_versions() -> dict[str, str]:
    """Every place the version is written down, keyed by a human-readable location."""
    server = _json("server.json")
    from word_document_server import __version__

    return {
        "pyproject.toml [project] version": _pyproject()["project"]["version"],
        "manifest.json version": _json("manifest.json")["version"],
        "server.json version": server["version"],
        "server.json packages[0] version": server["packages"][0]["version"],
        "word_document_server.__version__": __version__,
    }


def test_every_declared_version_agrees():
    """Bumping a release means bumping all of them; this says which one was missed."""
    declared = _declared_versions()
    distinct = set(declared.values())
    assert len(distinct) == 1, (
        "the declared versions have drifted apart: "
        + ", ".join(f"{where} = {version}" for where, version in sorted(declared.items()))
        + " -- set them all to the same value"
    )


def test_the_declared_version_is_a_release_number():
    """A three-part version, so the git tag and the PyPI release can be derived from it."""
    version = _pyproject()["project"]["version"]
    parts = version.split(".")
    assert len(parts) == 3 and all(p.isdigit() for p in parts), (
        f"expected a MAJOR.MINOR.PATCH version, got {version!r}"
    )


def test_every_declared_distribution_name_agrees():
    """A rename that misses one of these installs a different package than it builds."""
    server = _json("server.json")
    declared = {
        "pyproject.toml [project] name": _pyproject()["project"]["name"],
        "manifest.json name": _json("manifest.json")["name"],
        "manifest.json uvx args": _json("manifest.json")["server"]["mcp_config"]["args"][0],
        "server.json packages[0] identifier": server["packages"][0]["identifier"],
    }
    wrong = {where: name for where, name in declared.items() if name != DISTRIBUTION}
    assert not wrong, (
        f"these do not name {DISTRIBUTION!r}: "
        + ", ".join(f"{where} = {name!r}" for where, name in sorted(wrong.items()))
    )


def test_the_mcp_name_trailer_matches_server_json():
    """The MCP registry reads the trailer; it must be the last line and match ``server.json``."""
    readme = (REPO_ROOT / "README.md").read_text(encoding="utf-8")
    trailer = f"<!-- mcp-name: {MCP_NAME} -->"
    assert readme.rstrip().endswith(trailer), (
        f"README.md must end with {trailer!r}; it ends with "
        f"{readme.rstrip().splitlines()[-1]!r}"
    )
    assert _json("server.json")["name"] == MCP_NAME
