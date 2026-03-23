"""Headers/footers mixin."""

from __future__ import annotations

from .base import W


class HeadersFootersMixin:
    """Header and footer operations."""

    def get_headers_footers(self) -> list[dict]:
        """Get all headers and footers with their text content."""
        results = []
        for rel_path, tree in self._trees.items():
            if not rel_path.startswith("word/header") and not rel_path.startswith("word/footer"):
                continue
            location = "header" if "header" in rel_path else "footer"
            text = self._text(tree)
            results.append(
                {
                    "part": rel_path,
                    "location": location,
                    "text": text,
                }
            )
        return sorted(results, key=lambda x: x["part"])
