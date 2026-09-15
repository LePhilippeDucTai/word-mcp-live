"""Typed exceptions raised by the OOXML engine.

Every failure the engine reports on purpose derives from :class:`EngineError`,
so a caller (the MCP tool layer, a test) can catch the whole family in one
clause and map it to a structured error response.

Operating-system failures are *not* wrapped: a missing file raises
``FileNotFoundError``, a full disk raises ``OSError``.  :class:`PackageError` is
about package semantics -- a file that is not an OPC package, a part that does
not exist, a relationship id that resolves to nothing -- not about I/O.
"""

from __future__ import annotations


class EngineError(Exception):
    """Base class of every error the engine raises deliberately."""


class PackageError(EngineError):
    """The package is not a usable WordprocessingML package, or lacks a part."""


class LocatorError(EngineError):
    """A locator does not resolve, or resolves ambiguously.

    Attributes:
        code: stable machine-readable reason (``"not-found"``, ``"ambiguous"``,
            ``"out-of-range"``, ...).  Callers branch on `code`; the message is
            for humans and may change.
    """

    def __init__(self, code: str, message: str | None = None) -> None:
        super().__init__(message if message is not None else code)
        self.code = code


class UnsupportedRange(EngineError):
    """A range exists but the requested operation cannot be applied to it.

    Attributes:
        reason: stable machine-readable reason (``"crosses-table-cell"``,
            ``"inside-field-instruction"``, ...).
    """

    def __init__(self, reason: str, message: str | None = None) -> None:
        super().__init__(message if message is not None else reason)
        self.reason = reason


class InvalidText(EngineError):
    """Text supplied by the caller cannot be stored in a WordprocessingML run.

    Covers characters XML 1.0 forbids and control characters Word rejects.
    """


class UnsupportedRevision(EngineError):
    """A tracked revision is of a kind the engine refuses to rewrite blindly."""


class IdExhausted(EngineError):
    """No free identifier is left in the space the caller asked for."""
