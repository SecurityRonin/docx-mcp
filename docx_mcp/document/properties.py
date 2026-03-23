"""Properties mixin: read and write core document properties."""

from __future__ import annotations

from lxml import etree

from .base import CP, DC, DCTERMS

# Maps property name -> XML tag
_PROP_MAP = {
    "title": f"{DC}title",
    "creator": f"{DC}creator",
    "subject": f"{DC}subject",
    "description": f"{DC}description",
    "last_modified_by": f"{CP}lastModifiedBy",
}


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

    def set_properties(
        self,
        *,
        title: str | None = None,
        creator: str | None = None,
        subject: str | None = None,
        description: str | None = None,
    ) -> dict:
        """Set core document properties.

        Args:
            title: Document title.
            creator: Document author/creator.
            subject: Document subject.
            description: Document description/comments.
        """
        tree = self._require("docProps/core.xml")

        updates = {
            "title": title,
            "creator": creator,
            "subject": subject,
            "description": description,
        }

        for prop_name, value in updates.items():
            if value is None:
                continue
            tag = _PROP_MAP[prop_name]
            el = tree.find(tag)
            if el is None:
                el = etree.SubElement(tree, tag)
            el.text = value

        self._mark("docProps/core.xml")
        return self.get_properties()
