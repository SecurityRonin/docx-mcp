"""Formatting mixin: apply character formatting with tracked changes."""

from __future__ import annotations

from lxml import etree

from .base import W14, W, _now_iso, _preserve
from .errors import DocxMcpError, ErrCode


class FormattingMixin:
    """Apply formatting to text with tracked change markup."""

    def set_formatting(
        self,
        para_id: str,
        text: str,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: str | None = None,
        color: str | None = None,
        author: str = "Claude",
    ) -> dict:
        """Apply character formatting to a text substring with tracked changes.

        Finds the run containing `text`, splits if needed, clones the original
        rPr as rPrChange, and applies the new formatting properties.

        Args:
            para_id: Target paragraph paraId.
            text: Exact text to format.
            bold: Apply bold.
            italic: Apply italic.
            underline: Underline style (e.g., "single", "double").
            color: Font color as hex (e.g., "FF0000").
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        now = _now_iso()
        cid = self._next_markup_id(doc)

        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or t_el.text is None:
                continue
            full = t_el.text
            if text not in full:
                continue

            idx = full.index(text)
            old_rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(old_rpr) if old_rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)

            insert_at = pos

            # Text before match — plain run
            if idx > 0:
                before = self._make_run(full[:idx], rpr_bytes)
                parent.insert(insert_at, before)
                insert_at += 1

            # Formatted run with rPrChange for tracked change
            fmt_run = etree.Element(f"{W}r")
            new_rpr = etree.SubElement(fmt_run, f"{W}rPr")

            # Apply requested formatting
            if bold:
                etree.SubElement(new_rpr, f"{W}b")
            if italic:
                etree.SubElement(new_rpr, f"{W}i")
            if underline:
                u = etree.SubElement(new_rpr, f"{W}u")
                u.set(f"{W}val", underline)
            if color:
                c = etree.SubElement(new_rpr, f"{W}color")
                c.set(f"{W}val", color)

            # Record original formatting as rPrChange
            rpr_change = etree.SubElement(new_rpr, f"{W}rPrChange")
            rpr_change.set(f"{W}id", str(cid))
            rpr_change.set(f"{W}author", author)
            rpr_change.set(f"{W}date", now)
            if rpr_bytes:
                rpr_change.append(etree.fromstring(rpr_bytes))
            else:
                etree.SubElement(rpr_change, f"{W}rPr")

            fmt_t = etree.SubElement(fmt_run, f"{W}t")
            _preserve(fmt_t, text)
            parent.insert(insert_at, fmt_run)
            insert_at += 1

            # Text after match — plain run
            end = idx + len(text)
            if end < len(full):
                after = self._make_run(full[end:], rpr_bytes)
                parent.insert(insert_at, after)

            self._mark("word/document.xml")
            return {"formatted": True}

        raise ValueError(f"Text '{text}' not found in a single run of paragraph '{para_id}'.")

    def insert_paragraph(
        self,
        after_para_id: str,
        text: str,
        style: str | None = None,
    ) -> dict:
        doc = self._require("word/document.xml")
        target = self._find_para(doc, after_para_id)
        if target is None:
            raise ValueError(f"Paragraph '{after_para_id}' not found")

        new_pid = self._new_para_id()
        new_p = etree.Element(f"{W}p")
        new_p.set(f"{W14}paraId", new_pid)

        if style is not None:
            ppr = etree.SubElement(new_p, f"{W}pPr")
            ps = etree.SubElement(ppr, f"{W}pStyle")
            ps.set(f"{W}val", style)

        run = etree.SubElement(new_p, f"{W}r")
        t_el = etree.SubElement(run, f"{W}t")
        _preserve(t_el, text)

        target.addnext(new_p)
        self._mark("word/document.xml")
        return {"para_id": new_pid, "text": text}

    def update_paragraph(
        self,
        para_id: str,
        text: str | None = None,
        style: str | None = None,
        tracked: bool = False,
        author: str = "Claude",
    ) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        if text is not None:
            if tracked:
                # ── Tracked: wrap existing runs in w:del, append w:ins ──
                cid = self._next_markup_id(doc)
                now = _now_iso()

                # Wrap each existing w:r in w:del
                for run in list(para.findall(f"{W}r")):
                    para.remove(run)
                    # Convert w:t → w:delText inside runs
                    for t_el in run.iter(f"{W}t"):
                        t_el.tag = f"{W}delText"
                    del_el = etree.Element(f"{W}del")
                    del_el.set(f"{W}id", str(cid))
                    del_el.set(f"{W}author", author)
                    del_el.set(f"{W}date", now)
                    del_el.append(run)
                    para.append(del_el)

                # Append w:ins with new text
                ins_cid = self._next_markup_id(doc)
                ins_el = etree.Element(f"{W}ins")
                ins_el.set(f"{W}id", str(ins_cid))
                ins_el.set(f"{W}author", author)
                ins_el.set(f"{W}date", now)
                new_run = etree.SubElement(ins_el, f"{W}r")
                new_t = etree.SubElement(new_run, f"{W}t")
                _preserve(new_t, text)
                para.append(ins_el)
            else:
                # ── Silent: existing behavior ───────────────────────────
                for run in list(para.findall(f"{W}r")):
                    para.remove(run)
                run = etree.SubElement(para, f"{W}r")
                t_el = etree.SubElement(run, f"{W}t")
                _preserve(t_el, text)

        if style is not None:
            ppr = para.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                para.insert(0, ppr)
            pstyle = ppr.find(f"{W}pStyle")
            if pstyle is None:
                pstyle = etree.Element(f"{W}pStyle")
                ppr.insert(0, pstyle)
            pstyle.set(f"{W}val", style)

        if text is not None or style is not None:
            self._mark("word/document.xml")
        return {"para_id": para_id}

    def delete_paragraph(self, para_id: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        para.getparent().remove(para)
        self._mark("word/document.xml")
        return {"deleted": para_id}

    def set_paragraph_border(
        self,
        para_id: str,
        sides: list[str],
        color: str = "000000",
        size: int = 4,
    ) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        pbdr = ppr.find(f"{W}pBdr")
        if pbdr is None:
            pbdr = etree.Element(f"{W}pBdr")
            # Insert before w:shd to maintain CT_PPr schema order
            shd = ppr.find(f"{W}shd")
            if shd is not None:
                shd.addprevious(pbdr)
            else:
                ppr.append(pbdr)

        for side in sides:
            el = pbdr.find(f"{W}{side}")
            if el is None:
                el = etree.SubElement(pbdr, f"{W}{side}")
            el.set(f"{W}val", "single")
            el.set(f"{W}sz", str(size))
            el.set(f"{W}space", "0")
            el.set(f"{W}color", color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "sides": sides}

    def set_paragraph_shading(
        self,
        para_id: str,
        fill_color: str,
        pattern: str = "clear",
    ) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        shd = ppr.find(f"{W}shd")
        if shd is None:
            shd = etree.SubElement(ppr, f"{W}shd")
        shd.set(f"{W}val", pattern)
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", fill_color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "fill_color": fill_color}

    # ── Paragraph-level formatting ────────────────────────────────────────────

    _CM_TO_TWIPS = round(1440 / 2.54)  # 567

    def set_paragraph_indentation(
        self,
        para_id: str,
        *,
        left_cm: float | None = None,
        right_cm: float | None = None,
        first_line_cm: float | None = None,
        hanging_cm: float | None = None,
    ) -> dict:
        """Set indentation on a paragraph.

        Args:
            para_id: Target paragraph paraId.
            left_cm: Left indent in cm.
            right_cm: Right indent in cm.
            first_line_cm: First-line indent in cm (mutually exclusive with hanging_cm).
            hanging_cm: Hanging indent in cm (mutually exclusive with first_line_cm).
        """
        if first_line_cm is not None and hanging_cm is not None:
            raise ValueError("first_line_cm and hanging_cm are mutually exclusive")

        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        ind = ppr.find(f"{W}ind")
        if ind is None:
            ind = etree.SubElement(ppr, f"{W}ind")

        if left_cm is not None:
            ind.set(f"{W}left", str(round(left_cm * self._CM_TO_TWIPS)))
        if right_cm is not None:
            ind.set(f"{W}right", str(round(right_cm * self._CM_TO_TWIPS)))
        if first_line_cm is not None:
            ind.set(f"{W}firstLine", str(round(first_line_cm * self._CM_TO_TWIPS)))
            ind.attrib.pop(f"{W}hanging", None)
        elif hanging_cm is not None:
            ind.set(f"{W}hanging", str(round(hanging_cm * self._CM_TO_TWIPS)))
            ind.attrib.pop(f"{W}firstLine", None)

        self._mark("word/document.xml")
        return {
            "para_id": para_id,
            "left_cm": left_cm,
            "right_cm": right_cm,
            "first_line_cm": first_line_cm,
            "hanging_cm": hanging_cm,
        }

    def set_line_spacing(
        self,
        para_id: str,
        *,
        line_rule: str | None = None,
        line_value: int | None = None,
        space_before_pt: float | None = None,
        space_after_pt: float | None = None,
    ) -> dict:
        """Set line spacing and paragraph spacing.

        Args:
            para_id: Target paragraph paraId.
            line_rule: "auto" | "exact" | "atLeast" maps to w:lineRule.
            line_value: For "auto": 240ths of a line (240=single, 360=1.5x, 480=double).
                        For "exact"/"atLeast": value in twips.
            space_before_pt: Space before paragraph in points.
            space_after_pt: Space after paragraph in points.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        spacing = ppr.find(f"{W}spacing")
        if spacing is None:
            spacing = etree.SubElement(ppr, f"{W}spacing")

        if line_rule is not None:
            spacing.set(f"{W}lineRule", line_rule)
        if line_value is not None:
            spacing.set(f"{W}line", str(line_value))
        if space_before_pt is not None:
            spacing.set(f"{W}before", str(round(space_before_pt * 20)))
        if space_after_pt is not None:
            spacing.set(f"{W}after", str(round(space_after_pt * 20)))

        self._mark("word/document.xml")
        return {
            "para_id": para_id,
            "line_rule": line_rule,
            "line_value": line_value,
            "space_before_pt": space_before_pt,
            "space_after_pt": space_after_pt,
        }

    def get_paragraph_format(self, para_id: str) -> dict:
        """Read all formatting attributes of a paragraph.

        Returns a dict with style, alignment, indentation, spacing, border,
        shading, and numPr keys.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")

        # Style
        style = ""
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            if ps is not None:
                style = ps.get(f"{W}val") or ""

        # Alignment
        alignment = ""
        if ppr is not None:
            jc = ppr.find(f"{W}jc")
            if jc is not None:
                alignment = jc.get(f"{W}val") or ""

        # Indentation
        ind_dict = {"left_twips": 0, "right_twips": 0, "first_line_twips": 0, "hanging_twips": 0}
        if ppr is not None:
            ind = ppr.find(f"{W}ind")
            if ind is not None:

                def _int(attr: str) -> int:
                    v = ind.get(f"{W}{attr}")
                    return int(v) if v is not None else 0

                ind_dict["left_twips"] = _int("left")
                ind_dict["right_twips"] = _int("right")
                ind_dict["first_line_twips"] = _int("firstLine")
                ind_dict["hanging_twips"] = _int("hanging")

        # Spacing
        sp_dict = {"before_twips": 0, "after_twips": 0, "line_value": 0, "line_rule": ""}
        if ppr is not None:
            sp = ppr.find(f"{W}spacing")
            if sp is not None:

                def _int_sp(attr: str) -> int:
                    v = sp.get(f"{W}{attr}")
                    return int(v) if v is not None else 0

                sp_dict["before_twips"] = _int_sp("before")
                sp_dict["after_twips"] = _int_sp("after")
                sp_dict["line_value"] = _int_sp("line")
                sp_dict["line_rule"] = sp.get(f"{W}lineRule") or ""

        # Border
        border = False
        if ppr is not None:
            border = ppr.find(f"{W}pBdr") is not None

        # Shading
        shading = False
        if ppr is not None:
            shading = ppr.find(f"{W}shd") is not None

        # numPr
        num_pr = None
        if ppr is not None:
            np_el = ppr.find(f"{W}numPr")
            if np_el is not None:
                ilvl_el = np_el.find(f"{W}ilvl")
                numid_el = np_el.find(f"{W}numId")
                ilvl = int(ilvl_el.get(f"{W}val", "0")) if ilvl_el is not None else 0
                numid = int(numid_el.get(f"{W}val", "0")) if numid_el is not None else 0
                num_pr = {"numId": numid, "ilvl": ilvl}

        return {
            "style": style,
            "alignment": alignment,
            "indentation": ind_dict,
            "spacing": sp_dict,
            "border": border,
            "shading": shading,
            "numPr": num_pr,
        }

    def _get_run(self, para, run_idx: int):
        runs = para.findall(f"{W}r")
        if run_idx < 0 or run_idx >= len(runs):
            raise IndexError(f"Run index {run_idx} out of range (have {len(runs)})")
        return runs[run_idx]

    def _upsert_rpr(self, run):
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            run.insert(0, rpr)
        return rpr

    def get_runs(self, para_id: str) -> list:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        result = []
        for idx, run in enumerate(para.findall(f"{W}r")):
            text = "".join(t.text or "" for t in run.findall(f"{W}t"))
            rpr = run.find(f"{W}rPr")

            bold = False
            italic = False
            font = None
            size_pt = None
            color = None

            if rpr is not None:
                b_el = rpr.find(f"{W}b")
                bold = b_el is not None and b_el.get(f"{W}val") not in ("0", "false", "off")

                i_el = rpr.find(f"{W}i")
                italic = i_el is not None and i_el.get(f"{W}val") not in ("0", "false", "off")

                rfonts = rpr.find(f"{W}rFonts")
                if rfonts is not None:
                    font = rfonts.get(f"{W}ascii")

                sz_el = rpr.find(f"{W}sz")
                if sz_el is not None:
                    val = sz_el.get(f"{W}val")
                    if val is not None:
                        size_pt = int(val) / 2

                color_el = rpr.find(f"{W}color")
                if color_el is not None:
                    color = color_el.get(f"{W}val")

            result.append(
                {
                    "run_idx": idx,
                    "text": text,
                    "bold": bold,
                    "italic": italic,
                    "font": font,
                    "size_pt": size_pt,
                    "color": color,
                }
            )

        return result

    def set_run_font(self, para_id: str, run_idx: int, font_name: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        rfonts = rpr.find(f"{W}rFonts")
        if rfonts is None:
            rfonts = etree.SubElement(rpr, f"{W}rFonts")
        rfonts.set(f"{W}ascii", font_name)
        rfonts.set(f"{W}hAnsi", font_name)
        rfonts.set(f"{W}cs", font_name)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "font": font_name}

    def set_run_color(self, para_id: str, run_idx: int, color: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        color_el = rpr.find(f"{W}color")
        if color_el is None:
            color_el = etree.SubElement(rpr, f"{W}color")
        color_el.set(f"{W}val", color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "color": color}

    def set_run_size(self, para_id: str, run_idx: int, size_pt: float) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        half_pts = str(round(size_pt * 2))

        sz_el = rpr.find(f"{W}sz")
        if sz_el is None:
            sz_el = etree.SubElement(rpr, f"{W}sz")
        sz_el.set(f"{W}val", half_pts)

        sz_cs_el = rpr.find(f"{W}szCs")
        if sz_cs_el is None:
            sz_cs_el = etree.SubElement(rpr, f"{W}szCs")
        sz_cs_el.set(f"{W}val", half_pts)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "size_pt": size_pt}

    def set_character_spacing(self, para_id: str, run_idx: int, spacing_pt: float) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        spacing_el = rpr.find(f"{W}spacing")
        if spacing_el is None:
            spacing_el = etree.SubElement(rpr, f"{W}spacing")
        spacing_el.set(f"{W}val", str(round(spacing_pt * 20)))

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "spacing_pt": spacing_pt}

    def set_character_position(self, para_id: str, run_idx: int, position_pt: float) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        pos_el = rpr.find(f"{W}position")
        if pos_el is None:
            pos_el = etree.SubElement(rpr, f"{W}position")
        pos_el.set(f"{W}val", str(round(position_pt * 2)))

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "position_pt": position_pt}

    def set_run_highlight(self, para_id: str, run_idx: int, color: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        hl = rpr.find(f"{W}highlight")
        if hl is None:
            hl = etree.SubElement(rpr, f"{W}highlight")
        hl.set(f"{W}val", color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "color": color}

    def set_run_strikethrough(self, para_id: str, run_idx: int, double: bool = False) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        if double:
            strike = rpr.find(f"{W}strike")
            if strike is not None:
                rpr.remove(strike)
            dstrike = rpr.find(f"{W}dstrike")
            if dstrike is None:
                dstrike = etree.SubElement(rpr, f"{W}dstrike")
        else:
            dstrike = rpr.find(f"{W}dstrike")
            if dstrike is not None:
                rpr.remove(dstrike)
            strike = rpr.find(f"{W}strike")
            if strike is None:
                strike = etree.SubElement(rpr, f"{W}strike")

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "double": double}

    def set_run_superscript(self, para_id: str, run_idx: int) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        va = rpr.find(f"{W}vertAlign")
        if va is None:
            va = etree.SubElement(rpr, f"{W}vertAlign")
        va.set(f"{W}val", "superscript")

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "valign": "superscript"}

    def set_run_subscript(self, para_id: str, run_idx: int) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        va = rpr.find(f"{W}vertAlign")
        if va is None:
            va = etree.SubElement(rpr, f"{W}vertAlign")
        va.set(f"{W}val", "subscript")

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "valign": "subscript"}

    def set_run_underline(self, para_id: str, run_idx: int, style: str = "single") -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        u = rpr.find(f"{W}u")
        if u is None:
            u = etree.SubElement(rpr, f"{W}u")
        u.set(f"{W}val", style)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "style": style}

    def clear_run_formatting(self, para_id: str, run_idx: int) -> dict:
        """Remove all children from a run's rPr, causing it to inherit style formatting.

        Args:
            para_id: Target paragraph paraId.
            run_idx: Zero-based index of the run.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        try:
            run = self._get_run(para, run_idx)
        except IndexError as exc:
            raise ValueError(str(exc)) from exc

        rpr = run.find(f"{W}rPr")
        if rpr is not None:
            for child in list(rpr):
                rpr.remove(child)
            run.remove(rpr)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "cleared": True}

    def set_run_language(self, para_id: str, run_idx: int, language_code: str) -> dict:
        """Set the language tag on a run for spell-checking.

        Args:
            para_id: Target paragraph paraId.
            run_idx: Zero-based index of the run.
            language_code: BCP-47 language code (e.g., "en-US", "fr-FR").
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        try:
            run = self._get_run(para, run_idx)
        except IndexError as exc:
            raise ValueError(str(exc)) from exc
        rpr = self._upsert_rpr(run)

        lang = rpr.find(f"{W}lang")
        if lang is None:
            lang = etree.SubElement(rpr, f"{W}lang")
        lang.set(f"{W}val", language_code)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "language": language_code}

    def set_text_case(self, para_id: str, run_idx: int, case: str) -> dict:
        """Set text case transformation on a run.

        Args:
            para_id: Target paragraph paraId.
            run_idx: Zero-based index of the run.
            case: One of "upper" (all caps), "small" (small caps), "none" (remove case).
        """
        _valid = ("upper", "small", "none")
        if case not in _valid:
            raise ValueError(f"case must be one of {_valid!r}, got {case!r}")

        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        try:
            run = self._get_run(para, run_idx)
        except IndexError as exc:
            raise ValueError(str(exc)) from exc
        rpr = self._upsert_rpr(run)

        if case == "upper":
            sc = rpr.find(f"{W}smallCaps")
            if sc is not None:
                rpr.remove(sc)
            if rpr.find(f"{W}caps") is None:
                etree.SubElement(rpr, f"{W}caps")
        elif case == "small":
            caps = rpr.find(f"{W}caps")
            if caps is not None:
                rpr.remove(caps)
            if rpr.find(f"{W}smallCaps") is None:
                etree.SubElement(rpr, f"{W}smallCaps")
        else:  # "none"
            for tag in (f"{W}caps", f"{W}smallCaps"):
                el = rpr.find(tag)
                if el is not None:
                    rpr.remove(el)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "case": case}

    def find_replace_formatted(
        self,
        find: str,
        replace: str,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        color: str | None = None,
        size_pt: float | None = None,
    ) -> dict:
        if not find:
            raise ValueError("find must be non-empty")

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        if body is None:
            body = doc

        count = 0

        for para in body.iter(f"{W}p"):
            changed = True
            while changed:
                changed = False
                for run_el in list(para):
                    if run_el.tag != f"{W}r":
                        continue
                    t_el = run_el.find(f"{W}t")
                    if t_el is None or t_el.text is None:
                        continue
                    full = t_el.text
                    if find not in full:
                        continue

                    idx = full.index(find)
                    old_rpr = run_el.find(f"{W}rPr")
                    rpr_bytes = etree.tostring(old_rpr) if old_rpr is not None else None
                    parent = run_el.getparent()
                    pos = list(parent).index(run_el)
                    parent.remove(run_el)

                    insert_at = pos

                    before_text = full[:idx]
                    if before_text:
                        before_run = self._make_run(before_text, rpr_bytes)
                        parent.insert(insert_at, before_run)
                        insert_at += 1

                    fmt_run = etree.Element(f"{W}r")
                    has_fmt = (
                        bold is not None
                        or italic is not None
                        or color is not None
                        or size_pt is not None
                    )  # noqa: E501
                    new_rpr = etree.SubElement(fmt_run, f"{W}rPr") if has_fmt else None

                    if bold is True:
                        etree.SubElement(new_rpr, f"{W}b")
                    elif bold is False:
                        b_el = etree.SubElement(new_rpr, f"{W}b")
                        b_el.set(f"{W}val", "0")

                    if italic is True:
                        etree.SubElement(new_rpr, f"{W}i")
                    elif italic is False:
                        i_el = etree.SubElement(new_rpr, f"{W}i")
                        i_el.set(f"{W}val", "0")

                    if color is not None:
                        c_el = etree.SubElement(new_rpr, f"{W}color")
                        c_el.set(f"{W}val", color)

                    if size_pt is not None:
                        half = str(round(size_pt * 2))
                        sz_el = etree.SubElement(new_rpr, f"{W}sz")
                        sz_el.set(f"{W}val", half)
                        sz_cs_el = etree.SubElement(new_rpr, f"{W}szCs")
                        sz_cs_el.set(f"{W}val", half)

                    repl_t = etree.SubElement(fmt_run, f"{W}t")
                    _preserve(repl_t, replace)
                    parent.insert(insert_at, fmt_run)
                    insert_at += 1

                    after_text = full[idx + len(find) :]
                    if after_text:
                        after_run = self._make_run(after_text, rpr_bytes)
                        parent.insert(insert_at, after_run)

                    count += 1
                    changed = True
                    break

        if count > 0:
            self._mark("word/document.xml")

        return {"find": find, "replace": replace, "count": count}

    # ── Paragraph pagination controls ─────────────────────────────────────────

    def _set_ppr_bool_flag(self, para_id: str, tag: str, enabled: bool) -> dict:
        """Toggle a presence-only boolean element in w:pPr."""
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        existing = ppr.find(tag)
        changed = False
        if enabled:
            if existing is None:
                etree.SubElement(ppr, tag)
                changed = True
        else:
            if existing is not None:
                ppr.remove(existing)
                changed = True

        if changed:
            self._mark("word/document.xml")
        return {"para_id": para_id, "enabled": enabled}

    def set_keep_with_next(self, para_id: str, enabled: bool) -> dict:
        """Keep this paragraph on the same page as the next paragraph."""
        return self._set_ppr_bool_flag(para_id, f"{W}keepNext", enabled)

    def set_keep_lines_together(self, para_id: str, enabled: bool) -> dict:
        """Keep all lines of this paragraph on the same page."""
        return self._set_ppr_bool_flag(para_id, f"{W}keepLines", enabled)

    def set_page_break_before(self, para_id: str, enabled: bool) -> dict:
        """Force a page break before this paragraph."""
        return self._set_ppr_bool_flag(para_id, f"{W}pageBreakBefore", enabled)

    def set_widow_control(self, para_id: str, enabled: bool) -> dict:
        """Enable widow/orphan control for this paragraph."""
        return self._set_ppr_bool_flag(para_id, f"{W}widowControl", enabled)

    def insert_blockquote(self, para_id: str, text: str) -> dict:
        """Insert a blockquote paragraph after para_id: 720-twip left indent + italic.

        Returns:
            {"para_id": str (new paragraph), "text": str}
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        ref = self._find_para(body, para_id)
        if ref is None:
            raise DocxMcpError(ErrCode.PARA_NOT_FOUND, f"Paragraph '{para_id}' not found.")

        new_id = self._new_para_id()
        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", new_id)
        new_para.set(f"{W14}textId", "77777777")

        ppr = etree.SubElement(new_para, f"{W}pPr")
        ind = etree.SubElement(ppr, f"{W}ind")
        ind.set(f"{W}left", "720")

        r = etree.SubElement(new_para, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        etree.SubElement(rpr, f"{W}i")
        etree.SubElement(rpr, f"{W}iCs")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        ref.addnext(new_para)
        self._mark("word/document.xml")
        return {"para_id": new_id, "text": text}

    def insert_code_block(self, para_id: str, text: str, language: str = "") -> dict:
        """Insert a code-block paragraph after para_id: Courier New 10pt + gray shading.

        Returns:
            {"para_id": str (new paragraph), "text": str, "language": str}
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        ref = self._find_para(body, para_id)
        if ref is None:
            raise DocxMcpError(ErrCode.PARA_NOT_FOUND, f"Paragraph '{para_id}' not found.")

        new_id = self._new_para_id()
        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", new_id)
        new_para.set(f"{W14}textId", "77777777")

        ppr = etree.SubElement(new_para, f"{W}pPr")
        shd = etree.SubElement(ppr, f"{W}shd")
        shd.set(f"{W}val", "clear")
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", "F2F2F2")

        r = etree.SubElement(new_para, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        fonts = etree.SubElement(rpr, f"{W}rFonts")
        fonts.set(f"{W}ascii", "Courier New")
        fonts.set(f"{W}hAnsi", "Courier New")
        sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", "20")
        szcs = etree.SubElement(rpr, f"{W}szCs")
        szcs.set(f"{W}val", "20")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        ref.addnext(new_para)
        self._mark("word/document.xml")
        return {"para_id": new_id, "text": text, "language": language}
