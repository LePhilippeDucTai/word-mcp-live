"""The package layer: open a ``.docx``, reach its parts, save it atomically.

Everything above this module works on lxml elements; this module is the only
place that knows how those elements are stored in a zip container -- and it
never touches the zip itself.  Reading, writing, ``[Content_Types].xml`` and the
``.rels`` parts are delegated to the OPC implementation shipped with
python-docx (``docx.opc``), because hand-rolling them is exactly how the
existing tools corrupt documents: a part added without its content-type
override, a relationship id reused, a rewritten archive that drops the parts the
writer did not know about.

Live XML parts
--------------
python-docx registers a part class for only a handful of content types
(document, styles, numbering, settings, header, footer, image, ...).  Everything
else -- ``footnotes.xml``, ``endnotes.xml``, ``comments.xml``,
``commentsExtended.xml``, ``theme1.xml`` -- loads as a blob-backed ``Part``.
Parsing such a part yields a *detached copy*: edits made on it are silently
dropped at save time.  :class:`DocxPackage` therefore loads the package with a
part factory that falls back to ``XmlPart`` for any ``+xml`` content type, so
every root element this module hands out is the one that gets reserialized.

Saving
------
:meth:`DocxPackage.save` serializes to memory first, then writes through
:func:`atomic_write_bytes`: the destination is either the previous file or the
complete new one, never a truncated package, and no temporary file survives a
failure.
"""

from __future__ import annotations

import contextlib
import io
import os
import tempfile
from pathlib import Path
from typing import IO

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.package import PackageReader, PartFactory, Unmarshaller
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from docx.oxml.parser import parse_xml
from docx.package import Package
from lxml import etree

from word_document_server.engine.errors import PackageError

__all__ = [
    "COMMENTS_CONTENT_TYPE",
    "MAIN_STORY",
    "STORY_CONTENT_TYPES",
    "DocxPackage",
    "atomic_write_bytes",
    "story_name",
]

_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml."

#: Story id of the main document part.  Matches ``tests.support.snapshot``.
MAIN_STORY = "document"

#: Content types of the parts that carry a body of content -- the *stories* the
#: engine walks.  ``comments.xml`` is deliberately absent: comments are
#: annotations attached to a story, not a story of their own, and treating them
#: as one would make every "for each story" loop rewrite comment text too.
STORY_CONTENT_TYPES = frozenset(
    {
        _WML + "document.main+xml",
        _WML + "template.main+xml",
        _WML + "header+xml",
        _WML + "footer+xml",
        _WML + "footnotes+xml",
        _WML + "endnotes+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    }
)

#: Content type of ``word/comments.xml``.
COMMENTS_CONTENT_TYPE = _WML + "comments+xml"

_XML_CONTENT_TYPES = frozenset({"application/xml", "text/xml"})


def story_name(partname: str) -> str:
    """Return the story id of a part name.

    ``"/word/header2.xml"`` -> ``"header2"``, ``"word/document.xml"`` ->
    ``"document"``.  Same convention as ``tests.support.snapshot``, so a story
    id can be used to index a :class:`~tests.support.snapshot.Snapshot`.
    """
    name = partname.lstrip("/")
    if name.startswith("word/") and name.endswith(".xml"):
        return name[len("word/") : -len(".xml")]
    return name


def _normalize_partname(partname: str) -> str:
    """Return `partname` in OPC form (leading slash), accepting the zip form."""
    name = str(partname).strip()
    if not name:
        raise PackageError("part name must not be empty")
    return name if name.startswith("/") else "/" + name


def atomic_write_bytes(path: str | os.PathLike[str], data: bytes) -> None:
    """Write `data` to `path` so that `path` is never observed half-written.

    The bytes go to a temporary file in the *same directory* (so that the final
    rename stays within one filesystem and is therefore atomic), are flushed and
    ``fsync``-ed, and only then replace the destination.  If anything fails the
    temporary file is removed and the destination keeps its previous content.

    Raises:
        OSError: propagated unchanged from the filesystem.
    """
    target = Path(path)
    directory = target.parent if str(target.parent) else Path()
    fd, temp_name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=directory)
    temp_path = Path(temp_name)
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, target)
    except BaseException:
        with contextlib.suppress(OSError):
            temp_path.unlink()
        raise
    # Make the rename itself durable.  Not supported everywhere; a failure here
    # does not make the write any less correct.
    with contextlib.suppress(OSError):
        dir_fd = os.open(directory, os.O_RDONLY)
        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)


class _EnginePartFactory(PartFactory):
    """Part factory that keeps every XML part live rather than blob-backed."""

    @classmethod
    def _part_cls_for(cls, content_type: str):
        registered = cls.part_type_for.get(content_type)
        if registered is not None:
            return registered
        if content_type.endswith("+xml") or content_type in _XML_CONTENT_TYPES:
            return XmlPart
        return Part


class _EnginePackage(Package):
    """A WordprocessingML package loaded through :class:`_EnginePartFactory`."""

    @classmethod
    def open(cls, pkg_file: str | IO[bytes]) -> _EnginePackage:
        package = cls()
        Unmarshaller.unmarshal(PackageReader.from_file(pkg_file), package, _EnginePartFactory)
        return package


class DocxPackage:
    """A WordprocessingML package opened for element-level editing.

    Instances are mutable and *not* thread-safe: the roots handed out by
    :attr:`document`, :meth:`stories` and :meth:`ensure_part` are the live
    elements, and editing them is how the layers above this one work.
    """

    def __init__(self, package: _EnginePackage, source: Path | None = None) -> None:
        self._package = package
        self._source = source

    # -- construction ---------------------------------------------------------

    @classmethod
    def open(cls, source: str | os.PathLike[str] | bytes | bytearray | IO[bytes]) -> DocxPackage:
        """Open a package from a path, a ``bytes`` blob or a binary stream.

        Raises:
            FileNotFoundError: if `source` is a path that does not exist.
            PackageError: if the bytes are not a readable WordprocessingML
                package (bad zip, missing main document part, malformed XML).
        """
        path: Path | None = None
        if isinstance(source, (bytes, bytearray)):
            stream: IO[bytes] = io.BytesIO(bytes(source))
        elif isinstance(source, (str, os.PathLike)):
            path = Path(source)
            # Read once so that the file handle is closed before anything else
            # can touch it, and so a missing file fails as a plain OSError.
            stream = io.BytesIO(path.read_bytes())
        else:
            stream = source

        # Untrusted input boundary: a corrupt archive surfaces as anything from
        # BadZipFile to KeyError depending on where the reader gives up, and the
        # caller only needs to know the bytes are not a package.
        try:
            package = _EnginePackage.open(stream)
        except Exception as exc:
            where = f" ({path})" if path is not None else ""
            raise PackageError(f"not a readable OPC package{where}: {exc}") from exc

        try:
            document_part = package.main_document_part
        except (KeyError, ValueError) as exc:
            raise PackageError(f"package has no single main document part: {exc}") from exc
        if not isinstance(document_part, XmlPart):
            raise PackageError(f"main document part is not XML: {document_part.content_type}")
        return cls(package, path)

    @property
    def source_path(self) -> Path | None:
        """Path the package was opened from, or ``None`` for bytes/stream."""
        return self._source

    # -- parts ----------------------------------------------------------------

    @property
    def package(self) -> _EnginePackage:
        """The underlying OPC package.  Escape hatch; prefer the methods here."""
        return self._package

    @property
    def document_part(self) -> XmlPart:
        """The ``word/document.xml`` part."""
        return self._package.main_document_part  # type: ignore[return-value]

    @property
    def document(self) -> etree._Element:
        """Root ``w:document`` element of the main document part."""
        return self.document_part.element

    def _parts_by_name(self) -> dict[str, Part]:
        """Every part reachable from the package, keyed by OPC part name."""
        return {str(part.partname): part for part in self._package.iter_parts()}

    def find_part(self, partname: str) -> Part | None:
        """Return the part named `partname`, or ``None`` if the package has none.

        `partname` may be given in OPC form (``"/word/styles.xml"``) or in zip
        form (``"word/styles.xml"``).
        """
        return self._parts_by_name().get(_normalize_partname(partname))

    def part(self, partname: str) -> Part:
        """Return the part named `partname`.

        Raises:
            PackageError: if the package has no such part.
        """
        part = self.find_part(partname)
        if part is None:
            raise PackageError(f"no part named {partname!r} in the package")
        return part

    @staticmethod
    def root_of(part: Part) -> etree._Element:
        """Return the live root element of an XML part.

        Raises:
            PackageError: if `part` holds binary content (an image, a font).
        """
        if not isinstance(part, XmlPart):
            raise PackageError(f"part {part.partname} is not an XML part ({part.content_type})")
        return part.element

    def stories(self) -> list[tuple[str, etree._Element]]:
        """Return ``(story id, root element)`` for every story of the package.

        Stories are the body, then every header, footer, footnotes and endnotes
        part, sorted by story id -- the main document always comes first, so a
        caller that stops at the first entry gets the body.  Roots are live:
        editing them and calling :meth:`save` persists the change.
        """
        found: list[tuple[str, etree._Element]] = []
        for name, part in self._parts_by_name().items():
            if part.content_type not in STORY_CONTENT_TYPES:
                continue
            found.append((story_name(name), self.root_of(part)))
        found.sort(key=lambda item: (item[0] != MAIN_STORY, item[0]))
        return found

    def ensure_part(
        self,
        partname: str,
        content_type: str,
        reltype: str,
        initial_xml: str | bytes,
    ) -> etree._Element:
        """Return the root of `partname`, creating the part if it is absent.

        Idempotent in both directions: an existing part keeps its content (
        `initial_xml` is ignored), and the relationship from the document part is
        reused rather than duplicated -- ``relate_to`` returns the existing rId
        when one already points at the same target with the same type.

        Raises:
            PackageError: if a part already exists under `partname` with a
                different content type, or if `initial_xml` is not well formed.
        """
        key = _normalize_partname(partname)
        existing = self._parts_by_name().get(key)
        if existing is not None:
            if existing.content_type != content_type:
                raise PackageError(
                    f"part {key} already exists with content type "
                    f"{existing.content_type!r}, not {content_type!r}"
                )
            root = self.root_of(existing)
            self.document_part.relate_to(existing, reltype)
            return root

        payload = initial_xml.encode("utf-8") if isinstance(initial_xml, str) else initial_xml
        try:
            element = parse_xml(payload)
        except etree.XMLSyntaxError as exc:
            raise PackageError(f"initial XML for {key} is not well formed: {exc}") from exc
        part = XmlPart(PackURI(key), content_type, element, self._package)
        self.document_part.relate_to(part, reltype)
        return element

    # -- relationships --------------------------------------------------------

    def _as_part(self, part_or_root: Part | etree._Element) -> Part:
        """Accept either a part or the root element of one, return the part."""
        if isinstance(part_or_root, Part):
            return part_or_root
        root = part_or_root.getroottree().getroot()
        for part in self._package.iter_parts():
            if isinstance(part, XmlPart) and part.element is root:
                return part
        raise PackageError("element does not belong to any part of this package")

    def rel_target(self, root_part: Part | etree._Element, rId: str) -> Part | str:
        """Resolve `rId` against the relationships of one part.

        `root_part` is the part itself or the root element of an XML part -- what
        :meth:`stories` and :meth:`ensure_part` hand out -- so that resolving an
        ``r:id`` found in a header does not require looking the header part up
        again.  Returns the target :class:`Part` for an internal relationship, or
        the target URL as a ``str`` for an external one.

        Raises:
            PackageError: if the part carries no relationship under `rId`.
        """
        part = self._as_part(root_part)
        try:
            relationship = part.rels[rId]
        except KeyError:
            raise PackageError(f"part {part.partname} has no relationship {rId!r}") from None
        return relationship.target_ref if relationship.is_external else relationship.target_part

    def add_external_rel(
        self,
        part: Part | etree._Element,
        url: str,
        reltype: str = RT.HYPERLINK,
    ) -> str:
        """Relate `part` to an external `url` and return the rId to reference it.

        Deduplicated by target: relating the same part to the same URL with the
        same `reltype` twice returns the same rId, so a document that links the
        same address from ten paragraphs carries one relationship, not ten.

        Raises:
            PackageError: if `url` is empty.
        """
        if not url:
            raise PackageError("external relationship target must not be empty")
        return self._as_part(part).relate_to(url, reltype, is_external=True)

    # -- serialization --------------------------------------------------------

    def to_bytes(self) -> bytes:
        """Serialize the package to ``.docx`` bytes.

        Not byte-reproducible: the zip writer stamps every member with the
        current time, so two calls differ in archive metadata even when the
        document is untouched.  The fidelity contract of the engine is therefore
        stated on the canonical content of the parts, not on the archive bytes.
        """
        buffer = io.BytesIO()
        self._package.save(buffer)
        return buffer.getvalue()

    def save(self, path: str | os.PathLike[str]) -> Path:
        """Write the package to `path` atomically and return the path.

        Serialization happens first and in memory: if it fails, `path` has not
        been touched at all.
        """
        target = Path(path)
        atomic_write_bytes(target, self.to_bytes())
        return target
