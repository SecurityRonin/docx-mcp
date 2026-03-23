"""Formatting mixin: apply character formatting with tracked changes."""

from __future__ import annotations

from lxml import etree

from .base import W, _now_iso, _preserve


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
