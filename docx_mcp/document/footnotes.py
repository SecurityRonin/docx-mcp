"""Footnotes mixin: get, add, validate."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, _preserve


class FootnotesMixin:
    """Footnote operations."""

    def get_footnotes(self) -> list[dict]:
        fn_tree = self._tree("word/footnotes.xml")
        if fn_tree is None:
            return []
        result = []
        for fn in self._real_footnotes(fn_tree):
            result.append(
                {
                    "id": int(fn.get(f"{W}id", "0")),
                    "text": self._text(fn),
                }
            )
        return result

    def add_footnote(self, para_id: str, text: str) -> dict:
        """Add a footnote to a paragraph. Returns the new footnote ID."""
        doc = self._require("word/document.xml")
        fn_tree = self._require("word/footnotes.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        # Next ID
        existing = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        next_id = max(existing | {0}) + 1

        # Build footnote in footnotes.xml
        fn_el = etree.SubElement(fn_tree, f"{W}footnote")
        fn_el.set(f"{W}id", str(next_id))

        fn_para = etree.SubElement(fn_el, f"{W}p")
        fn_para.set(f"{W14}paraId", self._new_para_id())
        fn_para.set(f"{W14}textId", "77777777")

        ppr = etree.SubElement(fn_para, f"{W}pPr")
        ps = etree.SubElement(ppr, f"{W}pStyle")
        ps.set(f"{W}val", "FootnoteText")

        # Footnote ref mark
        ref_run = etree.SubElement(fn_para, f"{W}r")
        ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
        ref_style = etree.SubElement(ref_rpr, f"{W}rStyle")
        ref_style.set(f"{W}val", "FootnoteReference")
        etree.SubElement(ref_run, f"{W}footnoteRef")

        # Space
        sp_run = etree.SubElement(fn_para, f"{W}r")
        sp_t = etree.SubElement(sp_run, f"{W}t")
        _preserve(sp_t, " ")

        # Text
        txt_run = etree.SubElement(fn_para, f"{W}r")
        txt_t = etree.SubElement(txt_run, f"{W}t")
        _preserve(txt_t, text)

        self._mark("word/footnotes.xml")

        # Add reference in document paragraph
        r = etree.SubElement(para, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "FootnoteReference")
        fref = etree.SubElement(r, f"{W}footnoteReference")
        fref.set(f"{W}id", str(next_id))
        self._mark("word/document.xml")

        return {"footnote_id": next_id, "para_id": para_id}

    def validate_footnotes(self) -> dict:
        """Cross-reference footnote IDs between document.xml and footnotes.xml."""
        doc = self._tree("word/document.xml")
        fn_tree = self._tree("word/footnotes.xml")
        if doc is None:
            return {"error": "No document open"}
        if fn_tree is None:
            return {"valid": True, "references": 0, "definitions": 0}

        ref_ids = set()
        for ref in doc.iter(f"{W}footnoteReference"):
            fid = ref.get(f"{W}id")
            if fid:
                ref_ids.add(int(fid))

        def_ids = {int(f.get(f"{W}id", "0")) for f in self._real_footnotes(fn_tree)}

        missing = sorted(ref_ids - def_ids)
        orphans = sorted(def_ids - ref_ids)
        return {
            "valid": not missing and not orphans,
            "references": len(ref_ids),
            "definitions": len(def_ids),
            "missing_definitions": missing,
            "orphan_definitions": orphans,
        }
