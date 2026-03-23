"""Endnotes mixin: read endnotes."""

from __future__ import annotations

from .base import W


class EndnotesMixin:
    """Endnote operations."""

    def get_endnotes(self) -> list[dict]:
        """Get all endnotes (excluding separator endnotes id=0, -1)."""
        tree = self._tree("word/endnotes.xml")
        if tree is None:
            return []
        return [
            {"id": int(en.get(f"{W}id", "0")), "text": self._text(en)}
            for en in tree.findall(f"{W}endnote")
            if en.get(f"{W}id") not in ("0", "-1")
        ]
