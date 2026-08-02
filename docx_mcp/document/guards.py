"""Centralized input validation for docx-mcp tool parameters."""

from __future__ import annotations

import re
from pathlib import Path


class InputGuard:
    """Validates tool parameters at the MCP boundary. All methods raise ValueError on bad input."""

    PARA_ID_RE = re.compile(r"^[0-9A-Fa-f]{1,8}$")
    HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB

    @staticmethod
    def para_id(value: str) -> str:
        """Validate that value is 1–8 hex digit para ID."""
        if not InputGuard.PARA_ID_RE.fullmatch(value):
            raise ValueError(f"Invalid para_id {value!r}: must be 1–8 hexadecimal characters")
        return value

    @staticmethod
    def output_path(value: str, *, suffix: str = ".docx") -> Path:
        """Validate an output path: expected suffix, no path traversal.

        Args:
            value: Caller-supplied destination path.
            suffix: Required file extension. Defaults to .docx so existing
                call sites keep their contract; export tools pass their own
                (e.g. ".pdf") rather than writing unguarded.
        """
        p = Path(value)
        if p.suffix.lower() != suffix.lower():
            raise ValueError(f"Invalid output path {value!r}: suffix must be {suffix}")
        # Resolve and check for traversal sequences
        try:
            resolved = p.resolve()
        except (OSError, ValueError) as exc:  # pragma: no cover
            raise ValueError(f"Invalid output path {value!r}: {exc}") from exc  # pragma: no cover
        # Reject raw traversal sequences in the original string
        if ".." in p.parts:
            raise ValueError(f"Invalid output path {value!r}: path traversal not allowed")
        # NOTE: absolute path confinement (allowlisting) is enforced at call
        # sites (save(), copy_document()) rather than here.
        return resolved

    @staticmethod
    def color_hex(value: str) -> str:
        """Validate 6-character hex color string (no leading #)."""
        if not InputGuard.HEX_COLOR_RE.fullmatch(value):
            raise ValueError(f"Invalid color {value!r}: must be exactly 6 hexadecimal characters")
        return value

    @staticmethod
    def bounded_int(value: int, lo: int, hi: int, name: str) -> int:
        """Validate that value is an integer in [lo, hi]."""
        if not isinstance(value, int) or isinstance(value, bool):
            raise ValueError(f"{name} must be an integer, got {type(value).__name__}")
        if not (lo <= value <= hi):
            raise ValueError(f"{name} must be between {lo} and {hi}, got {value}")
        return value

    @staticmethod
    def regex_pattern(value: str) -> re.Pattern:
        """Compile a user-supplied regex pattern. Raises ValueError on syntax error."""
        try:
            return re.compile(value)
        except re.error as exc:
            raise ValueError(f"Invalid regex {value!r}: {exc}") from exc
