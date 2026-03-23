"""Endnotes mixin: read, add, and validate endnotes."""

from __future__ import annotations

from lxml import etree

from .base import NSMAP, W14, W, _preserve


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

    def add_endnote(self, para_id: str, text: str) -> dict:
        """Add an endnote to a paragraph.

        Creates the endnote definition in endnotes.xml and adds a superscript
        reference in the target paragraph.

        Args:
            para_id: paraId of the paragraph to attach the endnote to.
            text: Endnote text content.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        # Bootstrap endnotes.xml if missing
        en_tree = self._tree("word/endnotes.xml")
        if en_tree is None:
            en_tree = etree.Element(f"{W}endnotes", nsmap=NSMAP)
            # Add separator endnotes (id=-1 and 0)
            for sep_id in ("-1", "0"):
                sep = etree.SubElement(en_tree, f"{W}endnote")
                sep.set(f"{W}type", "separator" if sep_id == "0" else "continuationSeparator")
                sep.set(f"{W}id", sep_id)
                p = etree.SubElement(sep, f"{W}p")
                p.set(f"{W14}paraId", self._new_para_id())
                p.set(f"{W14}textId", "77777777")
            self._trees["word/endnotes.xml"] = en_tree

        # Determine next endnote ID
        existing_ids = [int(en.get(f"{W}id", "0")) for en in en_tree.findall(f"{W}endnote")]
        next_id = max(existing_ids, default=0) + 1

        # Create endnote definition
        endnote = etree.SubElement(en_tree, f"{W}endnote")
        endnote.set(f"{W}id", str(next_id))
        en_para = etree.SubElement(endnote, f"{W}p")
        en_para.set(f"{W14}paraId", self._new_para_id())
        en_para.set(f"{W14}textId", "77777777")
        en_run = etree.SubElement(en_para, f"{W}r")
        en_t = etree.SubElement(en_run, f"{W}t")
        _preserve(en_t, text)
        self._mark("word/endnotes.xml")

        # Add reference in paragraph
        ref_run = etree.SubElement(para, f"{W}r")
        rpr = etree.SubElement(ref_run, f"{W}rPr")
        etree.SubElement(rpr, f"{W}rStyle").set(f"{W}val", "EndnoteReference")
        etree.SubElement(ref_run, f"{W}endnoteReference").set(f"{W}id", str(next_id))
        self._mark("word/document.xml")

        return {"endnote_id": next_id, "para_id": para_id, "text": text}

    def validate_endnotes(self) -> dict:
        """Cross-reference endnote IDs between document and endnotes.xml."""
        doc = self._require("word/document.xml")
        en_tree = self._tree("word/endnotes.xml")

        # Collect references from document body
        ref_ids = set()
        for ref in doc.iter(f"{W}endnoteReference"):
            eid = ref.get(f"{W}id")
            if eid is not None:
                ref_ids.add(int(eid))

        # Collect definitions from endnotes.xml
        def_ids = set()
        if en_tree is not None:
            for en in en_tree.findall(f"{W}endnote"):
                eid = en.get(f"{W}id")
                if eid not in ("0", "-1") and eid is not None:
                    def_ids.add(int(eid))

        orphaned_refs = sorted(ref_ids - def_ids)
        orphaned_endnotes = sorted(def_ids - ref_ids)
        valid = len(orphaned_refs) == 0 and len(orphaned_endnotes) == 0

        return {
            "valid": valid,
            "total": len(def_ids),
            "orphaned_refs": orphaned_refs,
            "orphaned_endnotes": orphaned_endnotes,
        }
