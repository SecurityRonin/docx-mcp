"""Sections mixin: page breaks, section breaks, section properties."""

from __future__ import annotations

from lxml import etree

from .base import W14, W


class SectionsMixin:
    """Section and page break operations."""

    def add_page_break(self, para_id: str) -> dict:
        """Insert a page break after a paragraph.

        Creates a new paragraph containing a page break element.

        Args:
            para_id: paraId of the paragraph to insert after.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        new_para = etree.Element(f"{W}p")
        new_pid = self._new_para_id()
        new_para.set(f"{W14}paraId", new_pid)
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        br = etree.SubElement(run, f"{W}br")
        br.set(f"{W}type", "page")

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {"para_id": new_pid}

    def add_section_break(
        self,
        para_id: str,
        break_type: str = "nextPage",
    ) -> dict:
        """Add a section break at a paragraph.

        Inserts w:sectPr inside the paragraph's w:pPr to mark it as the last
        paragraph of its section.

        Args:
            para_id: paraId of the paragraph to place the section break on.
            break_type: "nextPage", "continuous", "evenPage", or "oddPage".
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(para, f"{W}pPr")
            # Move pPr to be first child
            para.remove(ppr)
            para.insert(0, ppr)

        sect_pr = etree.SubElement(ppr, f"{W}sectPr")
        type_el = etree.SubElement(sect_pr, f"{W}type")
        type_el.set(f"{W}val", break_type)

        self._mark("word/document.xml")

        return {"para_id": para_id, "break_type": break_type}

    def set_section_properties(
        self,
        *,
        para_id: str | None = None,
        width: int | None = None,
        height: int | None = None,
        orientation: str | None = None,
        margin_top: int | None = None,
        margin_bottom: int | None = None,
        margin_left: int | None = None,
        margin_right: int | None = None,
    ) -> dict:
        """Modify section properties (page size, orientation, margins).

        Args:
            para_id: paraId of paragraph with section break. None = body section.
            width: Page width in DXA (twips). 12240 = 8.5".
            height: Page height in DXA. 15840 = 11".
            orientation: "portrait" or "landscape".
            margin_top: Top margin in DXA.
            margin_bottom: Bottom margin in DXA.
            margin_left: Left margin in DXA.
            margin_right: Right margin in DXA.
        """
        doc = self._require("word/document.xml")

        if para_id is not None:
            # Paragraph-level section
            para = self._find_para(doc, para_id)
            if para is None:
                raise ValueError(f"Paragraph '{para_id}' not found")
            ppr = para.find(f"{W}pPr")
            sect_pr = ppr.find(f"{W}sectPr") if ppr is not None else None
            if sect_pr is None:
                raise ValueError(
                    f"No section break on paragraph '{para_id}'. Use add_section_break first."
                )
        else:
            # Body-level section
            body = doc.find(f"{W}body")
            sect_pr = body.find(f"{W}sectPr")
            if sect_pr is None:
                sect_pr = etree.SubElement(body, f"{W}sectPr")

        # Page size
        if width is not None or height is not None or orientation is not None:
            pg_sz = sect_pr.find(f"{W}pgSz")
            if pg_sz is None:
                pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")
            if width is not None:
                pg_sz.set(f"{W}w", str(width))
            if height is not None:
                pg_sz.set(f"{W}h", str(height))
            if orientation is not None:
                pg_sz.set(f"{W}orient", orientation)

        # Margins
        margin_vals = {
            "top": margin_top,
            "bottom": margin_bottom,
            "left": margin_left,
            "right": margin_right,
        }
        if any(v is not None for v in margin_vals.values()):
            pg_mar = sect_pr.find(f"{W}pgMar")
            if pg_mar is None:
                pg_mar = etree.SubElement(sect_pr, f"{W}pgMar")
            for attr, val in margin_vals.items():
                if val is not None:
                    pg_mar.set(f"{W}{attr}", str(val))

        self._mark("word/document.xml")

        # Build response from current state
        result: dict = {}
        pg_sz = sect_pr.find(f"{W}pgSz")
        if pg_sz is not None:
            w_val = pg_sz.get(f"{W}w")
            h_val = pg_sz.get(f"{W}h")
            if w_val:
                result["width"] = int(w_val)
            if h_val:
                result["height"] = int(h_val)
            orient = pg_sz.get(f"{W}orient")
            if orient:
                result["orientation"] = orient

        pg_mar = sect_pr.find(f"{W}pgMar")
        if pg_mar is not None:
            for attr in ("top", "bottom", "left", "right"):
                val = pg_mar.get(f"{W}{attr}")
                if val:
                    result[f"margin_{attr}"] = int(val)

        return result
