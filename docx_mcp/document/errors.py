"""Stable error taxonomy for docx-mcp tools."""

from __future__ import annotations

from enum import Enum


class ErrCode(str, Enum):
    STYLE_NOT_FOUND = "STYLE_NOT_FOUND"
    PARA_NOT_FOUND = "PARA_NOT_FOUND"
    BOOKMARK_NOT_FOUND = "BOOKMARK_NOT_FOUND"
    BOOKMARK_DANGLING = "BOOKMARK_DANGLING"
    PART_NOT_FOUND = "PART_NOT_FOUND"
    INVALID_RELATIONSHIP = "INVALID_REL"
    NUMBERING_ORPHAN = "NUMBERING_ORPHAN"
    OOXML_INVALID = "OOXML_INVALID"
    ENCRYPTED_DOCUMENT = "ENCRYPTED_DOCUMENT"
    PII_DEPS_MISSING = "PII_DEPS_MISSING"
    NO_OPEN_DOCUMENT = "NO_OPEN_DOCUMENT"
    XPATH_ERROR = "XPATH_ERROR"
    MALFORMED_INPUT = "MALFORMED_INPUT"
    UNSAFE_PATH = "UNSAFE_PATH"


class DocxMcpError(Exception):
    def __init__(self, code: ErrCode, message: str, hint: str = ""):
        self.code = code
        self.hint = hint
        super().__init__(message)

    def to_dict(self) -> dict:
        return {
            "error": self.code.value,
            "message": str(self),
            "hint": self.hint,
        }
