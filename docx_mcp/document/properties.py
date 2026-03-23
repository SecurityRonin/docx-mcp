"""Properties mixin: read core document properties."""

from __future__ import annotations

from .base import CP, DC, DCTERMS


class PropertiesMixin:
    """Document property operations."""

    def get_properties(self) -> dict:
        """Get core document properties (title, creator, dates, etc.)."""
        tree = self._tree("docProps/core.xml")
        if tree is None:
            return {}

        def _val(tag: str) -> str:
            el = tree.find(tag)
            return el.text if el is not None and el.text else ""

        return {
            "title": _val(f"{DC}title"),
            "creator": _val(f"{DC}creator"),
            "subject": _val(f"{DC}subject"),
            "description": _val(f"{DC}description"),
            "last_modified_by": _val(f"{CP}lastModifiedBy"),
            "revision": _val(f"{CP}revision"),
            "created": _val(f"{DCTERMS}created"),
            "modified": _val(f"{DCTERMS}modified"),
        }
