"""Lists mixin: create bullet and numbered lists."""

from __future__ import annotations

from lxml import etree

from .base import NSMAP, W


class ListsMixin:
    """List operations."""

    def add_list(
        self,
        para_ids: list[str],
        *,
        style: str = "bullet",
    ) -> dict:
        """Apply list formatting to paragraphs.

        Creates numbering definitions in numbering.xml (bootstrapped if missing)
        and sets w:numPr on each target paragraph.

        Args:
            para_ids: List of paraIds to format as list items.
            style: "bullet" or "numbered".
        """
        doc = self._require("word/document.xml")

        # Bootstrap numbering.xml if missing
        num_tree = self._tree("word/numbering.xml")
        if num_tree is None:
            num_tree = etree.Element(f"{W}numbering", nsmap=NSMAP)
            self._trees["word/numbering.xml"] = num_tree
            self._mark("word/numbering.xml")

        # Determine next abstract num ID and num ID
        existing_abstract = num_tree.findall(f"{W}abstractNum")
        abs_id = (
            max(
                (int(a.get(f"{W}abstractNumId", "0")) for a in existing_abstract),
                default=-1,
            )
            + 1
        )
        existing_nums = num_tree.findall(f"{W}num")
        num_id = max((int(n.get(f"{W}numId", "0")) for n in existing_nums), default=0) + 1

        # Create abstract numbering definition
        abstract = etree.SubElement(num_tree, f"{W}abstractNum")
        abstract.set(f"{W}abstractNumId", str(abs_id))
        lvl = etree.SubElement(abstract, f"{W}lvl")
        lvl.set(f"{W}ilvl", "0")
        fmt = etree.SubElement(lvl, f"{W}numFmt")
        if style == "numbered":
            fmt.set(f"{W}val", "decimal")
            lvl_text = etree.SubElement(lvl, f"{W}lvlText")
            lvl_text.set(f"{W}val", "%1.")
        else:
            fmt.set(f"{W}val", "bullet")
            lvl_text = etree.SubElement(lvl, f"{W}lvlText")
            lvl_text.set(f"{W}val", "\u2022")

        # Create num entry referencing abstract
        num_el = etree.SubElement(num_tree, f"{W}num")
        num_el.set(f"{W}numId", str(num_id))
        ref = etree.SubElement(num_el, f"{W}abstractNumId")
        ref.set(f"{W}val", str(abs_id))

        self._mark("word/numbering.xml")

        # Apply numPr to each paragraph
        count = 0
        for pid in para_ids:
            para = self._find_para(doc, pid)
            if para is None:
                raise ValueError(f"Paragraph '{pid}' not found")

            ppr = para.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.SubElement(para, f"{W}pPr")
                para.remove(ppr)
                para.insert(0, ppr)

            num_pr = etree.SubElement(ppr, f"{W}numPr")
            ilvl = etree.SubElement(num_pr, f"{W}ilvl")
            ilvl.set(f"{W}val", "0")
            nid = etree.SubElement(num_pr, f"{W}numId")
            nid.set(f"{W}val", str(num_id))
            count += 1

        self._mark("word/document.xml")
        return {"list_id": num_id, "paragraphs_updated": count}
