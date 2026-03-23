"""References mixin: cross-references via bookmarks and hyperlinks."""

from __future__ import annotations

from lxml import etree

from .base import W, _preserve


class ReferencesMixin:
    """Cross-reference operations."""

    def add_cross_reference(
        self,
        source_para_id: str,
        target_para_id: str,
        text: str,
    ) -> dict:
        """Add a cross-reference from one paragraph to another.

        If the target paragraph has an existing bookmark, reuses it.
        Otherwise, creates a new bookmark wrapping the target paragraph content.
        Appends a hyperlink with the reference text at the end of the source paragraph.

        Args:
            source_para_id: paraId of the paragraph where the reference link appears.
            target_para_id: paraId of the paragraph being referenced.
            text: Display text for the cross-reference link.
        """
        doc = self._require("word/document.xml")

        source = self._find_para(doc, source_para_id)
        if source is None:
            raise ValueError(f"Source paragraph '{source_para_id}' not found")

        target = self._find_para(doc, target_para_id)
        if target is None:
            raise ValueError(f"Target paragraph '{target_para_id}' not found")

        # Check for existing bookmark on target
        bookmark_name = None
        bm_start = target.find(f"{W}bookmarkStart")
        if bm_start is not None:
            bookmark_name = bm_start.get(f"{W}name")

        if bookmark_name is None:
            # Create bookmark wrapping target content
            bookmark_name = f"_Ref_{target_para_id}"
            bm_id = self._next_markup_id(doc)

            bm_start = etree.Element(f"{W}bookmarkStart")
            bm_start.set(f"{W}id", str(bm_id))
            bm_start.set(f"{W}name", bookmark_name)
            target.insert(0, bm_start)

            bm_end = etree.Element(f"{W}bookmarkEnd")
            bm_end.set(f"{W}id", str(bm_id))
            target.append(bm_end)

        # Add hyperlink at end of source paragraph
        hyperlink = etree.SubElement(source, f"{W}hyperlink")
        hyperlink.set(f"{W}anchor", bookmark_name)
        run = etree.SubElement(hyperlink, f"{W}r")
        rpr = etree.SubElement(run, f"{W}rPr")
        rstyle = etree.SubElement(rpr, f"{W}rStyle")
        rstyle.set(f"{W}val", "Hyperlink")
        t = etree.SubElement(run, f"{W}t")
        _preserve(t, text)

        self._mark("word/document.xml")

        return {
            "source_para_id": source_para_id,
            "target_para_id": target_para_id,
            "bookmark_name": bookmark_name,
            "text": text,
        }
