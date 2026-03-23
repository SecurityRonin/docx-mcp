"""Styles mixin: enumerate document styles."""

from __future__ import annotations

from .base import W


class StylesMixin:
    """Style inspection."""

    def get_styles(self) -> list[dict]:
        """Get all defined styles."""
        tree = self._tree("word/styles.xml")
        if tree is None:
            return []
        result = []
        for s in tree.findall(f"{W}style"):
            name_el = s.find(f"{W}name")
            based_el = s.find(f"{W}basedOn")
            result.append(
                {
                    "id": s.get(f"{W}styleId", ""),
                    "name": name_el.get(f"{W}val", "") if name_el is not None else "",
                    "type": s.get(f"{W}type", ""),
                    "base_style": based_el.get(f"{W}val", "") if based_el is not None else "",
                }
            )
        return result
