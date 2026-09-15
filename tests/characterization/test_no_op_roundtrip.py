"""No-op round-trip characterization: ``Document(path).save(out)``.

For every fixture (constructor-built and LibreOffice-produced), this pins
down what a *bare* python-docx round trip does to a ``.docx`` package when
none of this project's tools are involved and the ``save_utils`` monkey-patch
(:func:`word_document_server.utils.save_utils.install_save_hook`) is never
installed. ``word_document_server.main`` *is* imported during a test session
(transitively, via the root ``__init__.py``), but ``install_save_hook`` is
only ever imported and called inside ``main.py``'s ``run_server()``, which
this module never calls -- so the hook stays uninstalled here.

python-docx parts registered under a specialised ``Part`` subclass
(document, styles, numbering, header/footer...) get re-serialised; every
other part (comments, footnotes, endnotes, theme, fontTable...) is a generic
opaque ``Part`` that is read once and written back unchanged. Constructor
fixtures round-trip byte-for-byte identical (:func:`test_constructor_fixture_roundtrips_unchanged`).

LibreOffice-produced fixtures round-trip with two specific, harmless
divergences:

- ``_rels/.rels`` and ``word/_rels/document.xml.rels`` differ in their
  canonical XML digest, not because any relationship is added, removed, or
  retargeted (the relationship *sets* are identical), but because
  LibreOffice pretty-prints a trailing newline inside ``<Relationships>``
  that C14N treats as significant text content (see the "Canonicalisation"
  section of ``tests/support/snapshot.py``) and python-docx does not
  preserve when it rewrites the part.
- ``[Content_Types].xml`` drops a few now-redundant ``Default`` entries for
  extensions no part in the package actually uses (``rels``, ``png``,
  ``jpeg`` in these fixtures, which carry no image and whose ``.rels`` parts
  are already covered by the ``rels`` ``Default``). This is why
  ``[Content_Types].xml`` is compared as a *set* rather than byte-for-byte
  in the first place -- see
  :func:`test_libreoffice_fixture_roundtrips_with_known_rels_whitespace_only_diff`.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from tests.fixtures.builders import ALL_FIXTURES, build
from tests.support.libreoffice import fodt_to_docx, requires_libreoffice
from tests.support.package_check import validate_package
from tests.support.snapshot import assert_unchanged_except, diff, snapshot

pytestmark = pytest.mark.characterization

# The two relationship parts LibreOffice-sourced fixtures do not preserve
# byte-for-byte across a bare python-docx round trip (see module docstring).
_LIBREOFFICE_RELS_CHURN = ("_rels/.rels", "word/_rels/document.xml.rels")

# [Content_Types].xml Default entries for extensions unused by these
# fixtures, dropped by python-docx when it rewrites the part (see module
# docstring). Reusing them as "parts" is legitimate: assert_unchanged_except
# matches content-type entries against the same allowed-parts set.
_LIBREOFFICE_CONTENT_TYPE_CHURN = ("*.jpeg", "*.png")

_LIBREOFFICE_FIXTURE_NAMES = ("rich", "tables_lists", "notes_fields")


def _roundtrip(src: Path, out: Path) -> None:
    """``Document(src).save(out)`` -- the bare round trip under test."""
    doc = Document(str(src))
    doc.save(str(out))


@pytest.mark.parametrize("name", sorted(ALL_FIXTURES))
def test_constructor_fixture_roundtrips_unchanged(name: str, tmp_path: Path) -> None:
    """A python-docx round trip of a constructor fixture loses nothing.

    Every XML part is canonically identical, every binary part is identical
    byte for byte, no part is added or removed, and the package stays valid.
    """
    src = tmp_path / f"{name}.docx"
    src.write_bytes(build(name))
    out = tmp_path / f"{name}-out.docx"

    _roundtrip(src, out)

    before = snapshot(src.read_bytes())
    after = snapshot(out.read_bytes())
    delta = diff(before, after)
    assert delta.is_empty(), delta.describe()
    assert validate_package(out) == []


@requires_libreoffice
@pytest.mark.libreoffice
@pytest.mark.parametrize("name", _LIBREOFFICE_FIXTURE_NAMES)
def test_libreoffice_fixture_roundtrips_with_known_rels_whitespace_only_diff(
    name: str, tmp_path_factory
) -> None:
    """A python-docx round trip of a LibreOffice fixture loses nothing either.

    The only tolerated divergence is the C14N digest of the two relationship
    parts, caused by whitespace python-docx does not preserve when it
    rewrites them (see module docstring) -- the relationship *sets* stay
    identical, and no part is added or removed.
    """
    cache_dir = tmp_path_factory.mktemp(f"odt_docx_cache_{name}")
    src = fodt_to_docx(name, cache_dir)
    out = tmp_path_factory.mktemp(f"roundtrip_{name}") / f"{name}-out.docx"

    _roundtrip(src, out)

    before = snapshot(src.read_bytes())
    after = snapshot(out.read_bytes())

    # No part is lost or gained; only the two known rels parts may change.
    delta = diff(before, after)
    assert delta.parts_added == ()
    assert delta.parts_removed == ()
    assert set(delta.parts_changed) <= set(_LIBREOFFICE_RELS_CHURN)

    # The relationship *sets* are unaffected by the whitespace-only churn.
    assert before.relationships == after.relationships

    assert_unchanged_except(
        before, after, parts=_LIBREOFFICE_RELS_CHURN + _LIBREOFFICE_CONTENT_TYPE_CHURN
    )
    assert validate_package(out) == []


@requires_libreoffice
@pytest.mark.libreoffice
def test_libreoffice_fixture_preserves_comments_part(tmp_path_factory) -> None:
    """``comments.xml`` survives a bare round trip untouched (byte for byte)."""
    cache_dir = tmp_path_factory.mktemp("odt_docx_cache_comments")
    src = fodt_to_docx("rich", cache_dir)
    out = tmp_path_factory.mktemp("roundtrip_comments") / "rich-out.docx"

    _roundtrip(src, out)

    before = snapshot(src.read_bytes())
    after = snapshot(out.read_bytes())
    assert "word/comments.xml" in before.parts
    assert before.parts["word/comments.xml"] == after.parts["word/comments.xml"]


@requires_libreoffice
@pytest.mark.libreoffice
def test_libreoffice_fixture_preserves_footnotes_part(tmp_path_factory) -> None:
    """``footnotes.xml`` survives a bare round trip untouched (byte for byte)."""
    cache_dir = tmp_path_factory.mktemp("odt_docx_cache_footnotes")
    src = fodt_to_docx("notes_fields", cache_dir)
    out = tmp_path_factory.mktemp("roundtrip_footnotes") / "notes_fields-out.docx"

    _roundtrip(src, out)

    before = snapshot(src.read_bytes())
    after = snapshot(out.read_bytes())
    assert "word/footnotes.xml" in before.parts
    assert before.parts["word/footnotes.xml"] == after.parts["word/footnotes.xml"]
