"""Tables mixin: read and write table content."""

from __future__ import annotations

from lxml import etree

from .base import W14, W, _now_iso, _preserve


class TablesMixin:
    """Table operations."""

    def get_tables(self) -> list[dict]:
        """Get all tables with their content."""
        doc = self._require("word/document.xml")
        tables = []
        for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
            rows = []
            for tr in tbl.findall(f"{W}tr"):
                cells = []
                for tc in tr.findall(f"{W}tc"):
                    cells.append(self._text(tc))
                rows.append(cells)
            col_count = len(rows[0]) if rows else 0
            tables.append(
                {
                    "index": idx,
                    "row_count": len(rows),
                    "col_count": col_count,
                    "cells": rows,
                }
            )
        return tables

    def _get_table(self, table_idx: int) -> etree._Element:
        """Get table element by index, raising IndexError if not found."""
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if table_idx < 0 or table_idx >= len(tables):
            raise IndexError(f"Table index {table_idx} out of range (have {len(tables)})")
        return tables[table_idx]

    def add_table(
        self,
        para_id: str,
        rows: int,
        cols: int,
        *,
        author: str = "Claude",
    ) -> dict:
        """Insert a new table after a paragraph with tracked insertion.

        Args:
            para_id: paraId of the paragraph to insert after.
            rows: Number of rows.
            cols: Number of columns.
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        now = _now_iso()
        cid = self._next_markup_id(doc)

        tbl = etree.Element(f"{W}tbl")
        # Table properties
        tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
        tbl_style = etree.SubElement(tbl_pr, f"{W}tblStyle")
        tbl_style.set(f"{W}val", "TableGrid")
        tbl_w = etree.SubElement(tbl_pr, f"{W}tblW")
        tbl_w.set(f"{W}w", "0")
        tbl_w.set(f"{W}type", "auto")
        # Track change on table
        ins = etree.SubElement(tbl_pr, f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        # Grid columns
        grid = etree.SubElement(tbl, f"{W}tblGrid")
        for _ in range(cols):
            etree.SubElement(grid, f"{W}gridCol")

        # Rows and cells
        for _ in range(rows):
            tr = etree.SubElement(tbl, f"{W}tr")
            tr.set(f"{W14}paraId", self._new_para_id())
            tr.set(f"{W14}textId", "77777777")
            for _ in range(cols):
                tc = etree.SubElement(tr, f"{W}tc")
                p = etree.SubElement(tc, f"{W}p")
                p.set(f"{W14}paraId", self._new_para_id())
                p.set(f"{W14}textId", "77777777")

        para.addnext(tbl)
        self._mark("word/document.xml")

        # Calculate table index
        table_idx = list(doc.iter(f"{W}tbl")).index(tbl)
        return {"table_index": table_idx, "rows": rows, "cols": cols, "inserted": True}

    def modify_cell(
        self,
        table_idx: int,
        row: int,
        col: int,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Modify a table cell with tracked changes.

        Args:
            table_idx: Table index (0-based).
            row: Row index (0-based).
            col: Column index (0-based).
            text: New cell text.
            author: Author name for the revision.
        """
        tbl = self._get_table(table_idx)
        doc = self._require("word/document.xml")
        rows = tbl.findall(f"{W}tr")
        if row < 0 or row >= len(rows):
            raise IndexError(f"Row {row} out of range (have {len(rows)})")
        cells = rows[row].findall(f"{W}tc")
        if col < 0 or col >= len(cells):
            raise IndexError(f"Column {col} out of range (have {len(cells)})")

        tc = cells[col]
        now = _now_iso()
        cid = self._next_markup_id(doc)

        # Find first paragraph in cell
        para = tc.find(f"{W}p")
        if para is None:
            para = etree.SubElement(tc, f"{W}p")
            para.set(f"{W14}paraId", self._new_para_id())
            para.set(f"{W14}textId", "77777777")

        # Delete existing runs
        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or not t_el.text:
                continue
            rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(rpr) if rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)
            del_el = etree.Element(f"{W}del")
            del_el.set(f"{W}id", str(cid))
            del_el.set(f"{W}author", author)
            del_el.set(f"{W}date", now)
            del_run = etree.SubElement(del_el, f"{W}r")
            if rpr_bytes:
                del_run.append(etree.fromstring(rpr_bytes))
            dt = etree.SubElement(del_run, f"{W}delText")
            _preserve(dt, t_el.text)
            parent.insert(pos, del_el)
            cid = self._next_markup_id(doc)

        # Insert new text
        ins = etree.SubElement(para, f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)
        r = etree.SubElement(ins, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        self._mark("word/document.xml")
        return {"modified": True, "table_index": table_idx, "cell": [row, col], "text": text}

    def add_table_row(
        self,
        table_idx: int,
        row_idx: int | None = None,
        cells: list[str] | None = None,
        *,
        author: str = "Claude",
    ) -> dict:
        """Add a row to a table with tracked insertion.

        Args:
            table_idx: Table index (0-based).
            row_idx: Insert at this index. None = append at end.
            cells: Cell text content. None = empty cells.
            author: Author name for the revision.
        """
        tbl = self._get_table(table_idx)
        doc = self._require("word/document.xml")
        now = _now_iso()
        cid = self._next_markup_id(doc)

        # Determine column count from existing rows
        existing_rows = tbl.findall(f"{W}tr")
        col_count = len(existing_rows[0].findall(f"{W}tc")) if existing_rows else 1

        # Build new row
        tr = etree.Element(f"{W}tr")
        tr.set(f"{W14}paraId", self._new_para_id())
        tr.set(f"{W14}textId", "77777777")
        # Track change on row
        tr_pr = etree.SubElement(tr, f"{W}trPr")
        ins = etree.SubElement(tr_pr, f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        for i in range(col_count):
            tc = etree.SubElement(tr, f"{W}tc")
            p = etree.SubElement(tc, f"{W}p")
            p.set(f"{W14}paraId", self._new_para_id())
            p.set(f"{W14}textId", "77777777")
            if cells and i < len(cells):
                r = etree.SubElement(p, f"{W}r")
                t = etree.SubElement(r, f"{W}t")
                _preserve(t, cells[i])

        # Insert or append
        if row_idx is not None and row_idx < len(existing_rows):
            existing_rows[row_idx].addprevious(tr)
            final_idx = row_idx
        else:
            tbl.append(tr)
            final_idx = len(existing_rows)

        self._mark("word/document.xml")
        new_row_count = len(tbl.findall(f"{W}tr"))
        return {
            "table_index": table_idx,
            "row_index": final_idx,
            "row_count": new_row_count,
            "inserted": True,
        }

    def delete_table_row(
        self,
        table_idx: int,
        row_idx: int,
        *,
        author: str = "Claude",
    ) -> dict:
        """Delete a table row with tracked changes.

        Args:
            table_idx: Table index (0-based).
            row_idx: Row index to delete (0-based).
            author: Author name for the revision.
        """
        tbl = self._get_table(table_idx)
        doc = self._require("word/document.xml")
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range (have {len(rows)})")

        tr = rows[row_idx]
        now = _now_iso()
        cid = self._next_markup_id(doc)

        # Mark row itself as deleted via trPr
        tr_pr = tr.find(f"{W}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{W}trPr")
            tr.insert(0, tr_pr)
        del_el = etree.SubElement(tr_pr, f"{W}del")
        del_el.set(f"{W}id", str(cid))
        del_el.set(f"{W}author", author)
        del_el.set(f"{W}date", now)

        # Mark all runs in cells as deleted
        for tc in tr.findall(f"{W}tc"):
            for para in tc.findall(f"{W}p"):
                for run_el in list(para.findall(f"{W}r")):
                    t_el = run_el.find(f"{W}t")
                    if t_el is None or not t_el.text:
                        continue
                    cid = self._next_markup_id(doc)
                    rpr = run_el.find(f"{W}rPr")
                    rpr_bytes = etree.tostring(rpr) if rpr is not None else None
                    parent = run_el.getparent()
                    pos = list(parent).index(run_el)
                    parent.remove(run_el)
                    d = etree.Element(f"{W}del")
                    d.set(f"{W}id", str(cid))
                    d.set(f"{W}author", author)
                    d.set(f"{W}date", now)
                    dr = etree.SubElement(d, f"{W}r")
                    if rpr_bytes:
                        dr.append(etree.fromstring(rpr_bytes))
                    dt = etree.SubElement(dr, f"{W}delText")
                    _preserve(dt, t_el.text)
                    parent.insert(pos, d)

        self._mark("word/document.xml")
        return {"table_index": table_idx, "row_index": row_idx, "deleted": True}
