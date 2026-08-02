"""Tables mixin: read and write table content."""

from __future__ import annotations

import copy
import csv
import io

from lxml import etree

from .base import W14, W, _now_iso, _preserve
from .ooxml_order import (
    TBLBORDERS_ORDER,
    TBLPR_ORDER,
    TCMAR_ORDER,
    TCMAR_SIDES,
    TCPR_ORDER,
    TRPR_ORDER,
    ordered_set_child,
)

_CM_TO_TWIPS = 567

# One inch is 1440 twentieths of a point and 25.4 mm.
_MM_TO_TWIPS = 1440 / 25.4

# ST_TblWidth "pct" is expressed in fiftieths of a percent, so 100% is 5000.
_PERCENT_TO_PCT = 50

_TABLE_LAYOUT_MODES = ("fixed", "autofit")
_TABLE_WIDTH_UNITS = ("mm", "percent", "auto")


def _mm_to_twips(mm: float) -> int:
    return round(mm * _MM_TO_TWIPS)


def _validate_sides(sides: dict[str, float | None], what: str) -> None:
    """Reject a call that would write nothing, or a negative measurement."""
    if all(value is None for value in sides.values()):
        raise ValueError(f"{what}: give at least one of {', '.join(sides)}")
    for name, value in sides.items():
        if value is not None and value < 0:
            raise ValueError(f"{what}: {name} must be >= 0 mm, got {value}")


def _row_properties(tr: etree._Element) -> etree._Element:
    """Find or create w:trPr, which the schema requires as the first child."""
    tr_pr = tr.find(f"{W}trPr")
    if tr_pr is None:
        tr_pr = etree.Element(f"{W}trPr")
        tr.insert(0, tr_pr)
    return tr_pr


def _write_sides(container: etree._Element, sides: dict[str, float | None]) -> dict[str, float]:
    """Write the given sides of a CT_TcMar-shaped element, in schema order."""
    applied: dict[str, float] = {}
    for name in TCMAR_SIDES:
        value = sides.get(name)
        if value is None:
            continue
        el = ordered_set_child(container, name, TCMAR_ORDER)
        el.set(f"{W}w", str(_mm_to_twips(value)))
        el.set(f"{W}type", "dxa")
        applied[name] = value
    return applied


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
        """Get table element by index, raising IndexError if not found.

        Indexing is a pre-order walk, so a table nested inside a cell has its
        own index. Callers formatting documents with nested tables should read
        get_tables first rather than assuming top-level positions.
        """
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if table_idx < 0 or table_idx >= len(tables):
            raise IndexError(f"Table index {table_idx} out of range (have {len(tables)})")
        return tables[table_idx]

    def _get_cell(self, table_idx: int, row_idx: int, col_idx: int) -> etree._Element:
        """Resolve a cell by position, raising IndexError with the bounds."""
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range (have {len(rows)})")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range (have {len(cells)})")
        return cells[col_idx]

    @staticmethod
    def _cell_properties(tc: etree._Element) -> etree._Element:
        """Find or create w:tcPr, which the schema requires as the first child."""
        tc_pr = tc.find(f"{W}tcPr")
        if tc_pr is None:
            tc_pr = etree.Element(f"{W}tcPr")
            tc.insert(0, tc_pr)
        return tc_pr

    @staticmethod
    def _table_properties(tbl: etree._Element) -> etree._Element:
        """Find or create w:tblPr, which the schema requires as the first child."""
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        return tbl_pr

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
        tracked: bool = True,
    ) -> dict:
        """Modify a table cell, with tracked changes by default.

        Args:
            table_idx: Table index (0-based).
            row: Row index (0-based).
            col: Column index (0-based).
            text: New cell text.
            author: Author name for the revision.
            tracked: When False, overwrite cell text directly without w:del/w:ins markup.
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

        if tracked:
            # Delete existing runs as w:del
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

            # Insert new text as w:ins
            ins = etree.SubElement(para, f"{W}ins")
            ins.set(f"{W}id", str(cid))
            ins.set(f"{W}author", author)
            ins.set(f"{W}date", now)
            r = etree.SubElement(ins, f"{W}r")
            t = etree.SubElement(r, f"{W}t")
            _preserve(t, text)
        else:
            # Direct overwrite: remove existing runs, insert plain w:r
            for run_el in list(para.findall(f"{W}r")):
                para.remove(run_el)
            r = etree.SubElement(para, f"{W}r")
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

    # ── New table improvement methods ────────────────────────────────────────

    def merge_cells(
        self,
        table_index: int,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
    ) -> dict:
        """Merge a rectangular range of cells.

        Horizontal merge (same row): set w:gridSpan on start cell, remove
        intermediate cells.
        Vertical merge (same col): set w:vMerge val="restart" on top cell,
        w:vMerge (no val) on continuation cells.
        Rectangular: apply gridSpan on each row, then vMerge across rows.

        Returns {"merged_rows": int, "merged_cols": int, "table_index": int}.
        """
        tbl = self._get_table(table_index)
        rows_els = tbl.findall(f"{W}tr")

        merged_rows = end_row - start_row + 1
        merged_cols = end_col - start_col + 1

        if merged_rows < 1 or merged_cols < 1:
            raise ValueError("end indices must be >= start indices")

        if end_row >= len(rows_els):
            raise ValueError(f"end_row {end_row} out of range")

        for r in range(start_row, end_row + 1):
            row_el = rows_els[r]
            cells = row_el.findall(f"{W}tc")
            if end_col >= len(cells):
                raise ValueError(f"end_col {end_col} out of range in row {r}")

            if merged_cols > 1:
                # Set gridSpan on the first cell in the range
                first_tc = cells[start_col]
                grid_span = ordered_set_child(
                    self._cell_properties(first_tc), "gridSpan", TCPR_ORDER
                )
                grid_span.set(f"{W}val", str(merged_cols))

                # Remove intermediate cells
                for _c in range(start_col + 1, end_col + 1):
                    # Re-fetch cells after removals since list changed
                    current_cells = row_el.findall(f"{W}tc")
                    # The cell at start_col+1 becomes the next one to remove
                    if len(current_cells) > start_col + 1:
                        row_el.remove(current_cells[start_col + 1])

            if merged_rows > 1:
                # Apply vMerge to the start_col cell in this row
                tc = row_el.findall(f"{W}tc")[start_col]
                vmerge = ordered_set_child(self._cell_properties(tc), "vMerge", TCPR_ORDER)
                if r == start_row:
                    vmerge.set(f"{W}val", "restart")
                else:
                    # Continuation: no val attribute
                    if f"{W}val" in vmerge.attrib:
                        del vmerge.attrib[f"{W}val"]

        self._mark("word/document.xml")
        return {
            "table_index": table_index,
            "merged_rows": merged_rows,
            "merged_cols": merged_cols,
        }

    def set_header_row(self, table_index: int) -> dict:
        """Mark the first row as a header row that repeats across page breaks.

        Sets w:tblHeader on the first row's trPr.
        Returns {"table_index": int, "set": True}.
        """
        tbl = self._get_table(table_index)
        rows_els = tbl.findall(f"{W}tr")
        if not rows_els:
            raise ValueError("Table has no rows")

        first_row = rows_els[0]
        ordered_set_child(_row_properties(first_row), "tblHeader", TRPR_ORDER)

        self._mark("word/document.xml")
        return {"table_index": table_index, "set": True}

    def set_column_widths(self, table_index: int, widths_cm: list[float]) -> dict:
        """Set column widths in the table grid.

        Updates w:tblGrid/w:gridCol/@w:w (twips) and w:tcW on each cell.
        1 cm = 567 twips.
        Returns {"table_index": int, "column_count": int, "widths_cm": list[float]}.
        Raises ValueError if len(widths_cm) != column_count.
        """
        tbl = self._get_table(table_index)

        # Determine column count from tblGrid or first row
        grid = tbl.find(f"{W}tblGrid")
        if grid is not None:
            grid_cols = grid.findall(f"{W}gridCol")
            col_count = len(grid_cols)
        else:
            rows_els = tbl.findall(f"{W}tr")
            if rows_els:
                first_row_cells = rows_els[0].findall(f"{W}tc")
                col_count = 0
                for tc in first_row_cells:
                    tcPr = tc.find(f"{W}tcPr")
                    grid_span = tcPr.find(f"{W}gridSpan") if tcPr is not None else None
                    if grid_span is not None:
                        try:
                            col_count += int(grid_span.get(f"{W}val", "1"))
                        except ValueError:
                            col_count += 1
                    else:
                        col_count += 1
            else:
                col_count = 0

        if len(widths_cm) != col_count:
            raise ValueError(f"len(widths_cm)={len(widths_cm)} != column_count={col_count}")

        twips = [int(w * _CM_TO_TWIPS) for w in widths_cm]

        # Update tblGrid
        if grid is not None:
            for i, gc in enumerate(grid_cols):
                gc.set(f"{W}w", str(twips[i]))
        else:
            # Create tblGrid after tblPr (or at start)
            grid = etree.Element(f"{W}tblGrid")
            tbl_pr = tbl.find(f"{W}tblPr")
            if tbl_pr is not None:
                tbl_pr.addnext(grid)
            else:
                tbl.insert(0, grid)
            for tw in twips:
                gc = etree.SubElement(grid, f"{W}gridCol")
                gc.set(f"{W}w", str(tw))

        # Update tcW on each cell
        for row_el in tbl.findall(f"{W}tr"):
            cells = row_el.findall(f"{W}tc")
            for i, tc in enumerate(cells):
                if i >= col_count:
                    break
                tc_pr = tc.find(f"{W}tcPr")
                if tc_pr is None:
                    tc_pr = etree.Element(f"{W}tcPr")
                    tc.insert(0, tc_pr)
                tc_w = tc_pr.find(f"{W}tcW")
                if tc_w is None:
                    tc_w = etree.SubElement(tc_pr, f"{W}tcW")
                tc_w.set(f"{W}type", "dxa")
                tc_w.set(f"{W}w", str(twips[i]))

        self._mark("word/document.xml")
        return {
            "table_index": table_index,
            "column_count": col_count,
            "widths_cm": widths_cm,
        }

    def csv_to_table(
        self,
        para_id: str,
        csv_text: str,
        header_row: bool = True,
    ) -> dict:
        """Insert a table from CSV text after a paragraph.

        Returns {"table_index": int, "rows": int, "cols": int}.
        """
        reader = csv.reader(io.StringIO(csv_text))
        data = list(reader)
        if not data:
            raise ValueError("csv_text is empty")

        num_rows = len(data)
        num_cols = max(len(r) for r in data)

        result = self.add_table(para_id, num_rows, num_cols)
        table_index = result["table_index"]

        for r, row in enumerate(data):
            for c, cell_text in enumerate(row):
                self.modify_cell(table_index, r, c, cell_text)

        if header_row and num_rows > 0:
            self.set_header_row(table_index)

        return {"table_index": table_index, "rows": num_rows, "cols": num_cols}

    def table_to_csv(self, table_index: int) -> dict:
        """Export table content as CSV string.

        Returns {"csv": str, "rows": int, "cols": int}.
        """
        tbl = self._get_table(table_index)
        rows_data = []
        col_count = 0
        for tr in tbl.findall(f"{W}tr"):
            row_cells = [self._text(tc) for tc in tr.findall(f"{W}tc")]
            rows_data.append(row_cells)
            col_count = max(col_count, len(row_cells))

        buf = io.StringIO()
        writer = csv.writer(buf)
        for row in rows_data:
            writer.writerow(row)

        return {"csv": buf.getvalue(), "rows": len(rows_data), "cols": col_count}

    def delete_table(self, table_idx: int) -> dict:
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if table_idx < 0 or table_idx >= len(tables):
            raise IndexError(f"Table index {table_idx} out of range (have {len(tables)})")
        tbl = tables[table_idx]
        tbl.getparent().remove(tbl)
        self._mark("word/document.xml")
        return {"deleted": table_idx}

    def add_column_to_table(self, table_idx: int, header_text: str = "") -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        for i, tr in enumerate(rows):
            tc = etree.SubElement(tr, f"{W}tc")
            p = etree.SubElement(tc, f"{W}p")
            p.set(f"{W14}paraId", self._new_para_id())
            p.set(f"{W14}textId", "77777777")
            if i == 0 and header_text:
                r = etree.SubElement(p, f"{W}r")
                t = etree.SubElement(r, f"{W}t")
                t.text = header_text
        # Keep tblGrid in sync
        tbl_pr = tbl.find(f"{W}tblPr")
        tbl_grid = tbl.find(f"{W}tblGrid")
        if tbl_grid is None:
            tbl_grid = etree.Element(f"{W}tblGrid")
            if tbl_pr is not None:
                tbl_pr.addnext(tbl_grid)
            else:
                tbl.insert(0, tbl_grid)
        etree.SubElement(tbl_grid, f"{W}gridCol")
        col_count = len(tbl.findall(f"{W}tr")[0].findall(f"{W}tc")) if rows else 0
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "columns": col_count}

    def delete_column_from_table(self, table_idx: int, col_idx: int) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        # Validate all rows first to avoid partial mutation
        for tr in rows:
            cells = tr.findall(f"{W}tc")
            if col_idx < 0 or col_idx >= len(cells):
                raise IndexError(f"Column index {col_idx} out of range (have {len(cells)})")
        for tr in rows:
            cells = tr.findall(f"{W}tc")
            tr.remove(cells[col_idx])
        # Keep tblGrid in sync
        tbl_grid = tbl.find(f"{W}tblGrid")
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(f"{W}gridCol")
            if col_idx < len(grid_cols):
                tbl_grid.remove(grid_cols[col_idx])
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "col_idx": col_idx}

    def set_cell_width(self, table_idx: int, row_idx: int, col_idx: int, width_mm: float) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range")
        tc = cells[col_idx]
        tc_w = ordered_set_child(self._cell_properties(tc), "tcW", TCPR_ORDER)
        dxa = _mm_to_twips(width_mm)
        tc_w.set(f"{W}w", str(dxa))
        tc_w.set(f"{W}type", "dxa")
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "row_idx": row_idx, "col_idx": col_idx, "width_dxa": dxa}

    def set_cell_vertical_alignment(
        self, table_idx: int, row_idx: int, col_idx: int, alignment: str
    ) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range")
        tc = cells[col_idx]
        v_align = ordered_set_child(self._cell_properties(tc), "vAlign", TCPR_ORDER)
        v_align.set(f"{W}val", alignment)
        self._mark("word/document.xml")
        return {
            "table_idx": table_idx,
            "row_idx": row_idx,
            "col_idx": col_idx,
            "alignment": alignment,
        }  # noqa: E501

    def set_row_height(
        self, table_idx: int, row_idx: int, height_mm: float, rule: str = "exact"
    ) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range")
        tr = rows[row_idx]
        tr_height = ordered_set_child(_row_properties(tr), "trHeight", TRPR_ORDER)
        dxa = _mm_to_twips(height_mm)
        tr_height.set(f"{W}val", str(dxa))
        tr_height.set(f"{W}hRule", rule)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "row_idx": row_idx, "height_dxa": dxa}

    def set_table_alignment(self, table_idx: int, alignment: str) -> dict:
        tbl_pr = self._table_properties(self._get_table(table_idx))
        jc = ordered_set_child(tbl_pr, "jc", TBLPR_ORDER)
        jc.set(f"{W}val", alignment)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "alignment": alignment}

    def set_table_borders(
        self,
        table_idx: int,
        border_style: str = "single",
        color: str = "000000",
        size: int = 4,
    ) -> dict:
        tbl_pr = self._table_properties(self._get_table(table_idx))
        borders = ordered_set_child(tbl_pr, "tblBorders", TBLPR_ORDER)
        for stale in list(borders):
            borders.remove(stale)
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = ordered_set_child(borders, side, TBLBORDERS_ORDER)
            el.set(f"{W}val", border_style)
            el.set(f"{W}sz", str(size))
            el.set(f"{W}space", "0")
            el.set(f"{W}color", color)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "border_style": border_style, "color": color, "size": size}

    def set_cell_shading(
        self,
        table_idx: int,
        row_idx: int,
        col_idx: int,
        fill_color: str,
        pattern: str = "clear",
    ) -> dict:
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row {row_idx} out of range (have {len(rows)})")
        cells = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cells):
            raise IndexError(f"Column {col_idx} out of range (have {len(cells)})")
        tc = cells[col_idx]
        shd = ordered_set_child(self._cell_properties(tc), "shd", TCPR_ORDER)
        shd.set(f"{W}val", pattern)
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", fill_color)
        self._mark("word/document.xml")
        return {
            "table_idx": table_idx,
            "row_idx": row_idx,
            "col_idx": col_idx,
            "fill_color": fill_color,
        }  # noqa: E501

    # ── Geometry ─────────────────────────────────────────────────────────────

    def set_cell_padding(
        self,
        table_idx: int,
        row_idx: int,
        col_idx: int,
        *,
        top_mm: float | None = None,
        bottom_mm: float | None = None,
        left_mm: float | None = None,
        right_mm: float | None = None,
    ) -> dict:
        """Set the inner margins of one cell (w:tcMar).

        Sides left as None keep whatever the cell already has; they are not
        reset to zero. Measurements are millimetres, converted to the
        twentieths of a point OOXML stores.
        """
        sides = {"top": top_mm, "bottom": bottom_mm, "left": left_mm, "right": right_mm}
        _validate_sides(sides, "set_cell_padding")

        tc = self._get_cell(table_idx, row_idx, col_idx)
        tc_mar = ordered_set_child(self._cell_properties(tc), "tcMar", TCPR_ORDER)
        applied = _write_sides(tc_mar, sides)

        self._mark("word/document.xml")
        return {
            "table_idx": table_idx,
            "row_idx": row_idx,
            "col_idx": col_idx,
            "padding_mm": applied,
        }

    def set_table_cell_margins(
        self,
        table_idx: int,
        *,
        top_mm: float | None = None,
        bottom_mm: float | None = None,
        left_mm: float | None = None,
        right_mm: float | None = None,
    ) -> dict:
        """Set the table-wide default cell margins (w:tblCellMar).

        Individual cells override these via set_cell_padding.
        """
        sides = {"top": top_mm, "bottom": bottom_mm, "left": left_mm, "right": right_mm}
        _validate_sides(sides, "set_table_cell_margins")

        tbl_pr = self._table_properties(self._get_table(table_idx))
        cell_mar = ordered_set_child(tbl_pr, "tblCellMar", TBLPR_ORDER)
        applied = _write_sides(cell_mar, sides)

        self._mark("word/document.xml")
        return {"table_idx": table_idx, "margins_mm": applied}

    def set_table_layout(self, table_idx: int, mode: str) -> dict:
        """Set the table layout algorithm (w:tblLayout).

        "fixed" honours the declared column widths. "autofit" sizes columns to
        their content, which also requires clearing the explicit widths that
        would otherwise pin them — so this sets the table and every cell width
        back to auto.
        """
        if mode not in _TABLE_LAYOUT_MODES:
            raise ValueError(
                f"Unknown table layout mode {mode!r}: expected one of "
                f"{', '.join(_TABLE_LAYOUT_MODES)}"
            )

        tbl = self._get_table(table_idx)
        tbl_pr = self._table_properties(tbl)
        ordered_set_child(tbl_pr, "tblLayout", TBLPR_ORDER).set(f"{W}type", mode)

        if mode == "autofit":
            tbl_w = ordered_set_child(tbl_pr, "tblW", TBLPR_ORDER)
            tbl_w.set(f"{W}w", "0")
            tbl_w.set(f"{W}type", "auto")
            # Direct children only — a nested table keeps its own widths.
            for tr in tbl.findall(f"{W}tr"):
                for tc in tr.findall(f"{W}tc"):
                    tc_pr = tc.find(f"{W}tcPr")
                    tc_w = tc_pr.find(f"{W}tcW") if tc_pr is not None else None
                    if tc_w is not None:
                        tc_w.set(f"{W}w", "0")
                        tc_w.set(f"{W}type", "auto")

        self._mark("word/document.xml")
        return {"table_idx": table_idx, "mode": mode}

    def set_table_width(
        self,
        table_idx: int,
        width: float | None = None,
        unit: str = "mm",
    ) -> dict:
        """Set the preferred table width (w:tblW).

        Args:
            table_idx: Table index.
            width: Millimetres, or a percentage of the text column for
                unit="percent". Ignored (and may be None) for unit="auto".
            unit: "mm", "percent", or "auto".
        """
        if unit not in _TABLE_WIDTH_UNITS:
            raise ValueError(
                f"Unknown table width unit {unit!r}: expected one of "
                f"{', '.join(_TABLE_WIDTH_UNITS)}"
            )

        if unit == "auto":
            value, width_type = "0", "auto"
        else:
            if width is None:
                raise ValueError(f"width is required when unit is {unit!r}")
            if width < 0:
                raise ValueError(f"width must be >= 0, got {width}")
            if unit == "percent":
                if width > 100:
                    raise ValueError(f"percent width must be between 0 and 100, got {width}")
                value, width_type = str(round(width * _PERCENT_TO_PCT)), "pct"
            else:
                value, width_type = str(_mm_to_twips(width)), "dxa"

        tbl_pr = self._table_properties(self._get_table(table_idx))
        tbl_w = ordered_set_child(tbl_pr, "tblW", TBLPR_ORDER)
        tbl_w.set(f"{W}w", value)
        tbl_w.set(f"{W}type", width_type)

        self._mark("word/document.xml")
        return {"table_idx": table_idx, "width": width, "unit": unit}

    def set_table_style(self, table_idx: int, style_name: str) -> dict:
        tbl_pr = self._table_properties(self._get_table(table_idx))
        tbl_style = ordered_set_child(tbl_pr, "tblStyle", TBLPR_ORDER)
        tbl_style.set(f"{W}val", style_name)
        self._mark("word/document.xml")
        return {"table_idx": table_idx, "style_name": style_name}

    # ── Advanced table operations ────────────────────────────────────────────

    def split_table(self, table_idx: int, at_row_index: int) -> dict:
        """Split a table at at_row_index into two separate tables.

        Rows 0..at_row_index-1 stay in table 1; rows at_row_index..end go to
        the new table 2 inserted immediately after table 1 in the body.

        Args:
            table_idx: 0-based table index.
            at_row_index: 0-based row index; must be > 0 and < row_count.

        Returns:
            {"table1_rows": int, "table2_rows": int}
        """
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        row_count = len(rows)
        if at_row_index <= 0 or at_row_index >= row_count:
            raise ValueError(
                f"Split index out of range: at_row_index={at_row_index}, row_count={row_count}"
            )

        # Build table 2 as a deep copy of the original (preserves tblPr, tblGrid, etc.)
        tbl2 = copy.deepcopy(tbl)
        # Clear all w:tr children from tbl2, then repopulate with split rows
        for tr in tbl2.findall(f"{W}tr"):
            tbl2.remove(tr)
        for tr in rows[at_row_index:]:
            tbl2.append(copy.deepcopy(tr))

        # Remove split rows from original table 1
        for tr in rows[at_row_index:]:
            tbl.remove(tr)

        # Insert tbl2 immediately after tbl in its parent
        parent = tbl.getparent()
        tbl_pos = list(parent).index(tbl)
        parent.insert(tbl_pos + 1, tbl2)

        self._mark("word/document.xml")
        return {"table1_rows": at_row_index, "table2_rows": row_count - at_row_index}

    def duplicate_table_row(self, table_idx: int, row_index: int) -> dict:
        """Deep-copy a table row and insert the copy immediately after it.

        Clears w14:paraId / w14:textId attributes on copied paragraphs and
        assigns fresh unique IDs to avoid duplicates.

        Args:
            table_idx: 0-based table index.
            row_index: 0-based row to duplicate.

        Returns:
            {"row_index": row_index, "new_row_index": row_index + 1}
        """
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_index < 0 or row_index >= len(rows):
            raise ValueError(f"Row index {row_index} out of range (have {len(rows)})")

        original = rows[row_index]
        new_row = copy.deepcopy(original)

        # Replace row-level paraId / textId
        if new_row.get(f"{W14}paraId") is not None:
            new_row.set(f"{W14}paraId", self._new_para_id())
        if new_row.get(f"{W14}textId") is not None:
            new_row.set(f"{W14}textId", "77777777")

        # Replace paraId / textId attributes on all w:p in the copy
        for p in new_row.iter(f"{W}p"):
            if p.get(f"{W14}paraId") is not None:
                p.set(f"{W14}paraId", self._new_para_id())
            if p.get(f"{W14}textId") is not None:
                p.set(f"{W14}textId", self._new_para_id())

        # Insert after the original row
        original.addnext(new_row)

        self._mark("word/document.xml")
        return {"row_index": row_index, "new_row_index": row_index + 1}

    def sort_table(self, table_idx: int, column_index: int, ascending: bool = True) -> dict:
        """Sort the non-header rows of a table by the text content of a column.

        Header rows (rows with w:tr/w:trPr/w:tblHeader) are kept at the top
        in their original order.

        Args:
            table_idx: 0-based table index.
            column_index: 0-based column to sort by.
            ascending: True for A→Z, False for Z→A.

        Returns:
            {"sorted_rows": int, "column_index": int, "ascending": bool}
        """
        tbl = self._get_table(table_idx)
        all_rows = tbl.findall(f"{W}tr")

        header_rows = []
        data_rows = []
        for tr in all_rows:
            tr_pr = tr.find(f"{W}trPr")
            if tr_pr is not None and tr_pr.find(f"{W}tblHeader") is not None:
                header_rows.append(tr)
            else:
                data_rows.append(tr)

        def _row_sort_key(tr: etree._Element) -> str:
            cells = tr.findall(f"{W}tc")
            if column_index < len(cells):
                return "".join(t.text for t in cells[column_index].iter(f"{W}t") if t.text)
            return ""

        data_rows.sort(key=_row_sort_key, reverse=not ascending)

        # Remove all rows from table, then re-append in order
        for tr in all_rows:
            tbl.remove(tr)
        for tr in header_rows + data_rows:
            tbl.append(tr)

        self._mark("word/document.xml")
        return {"sorted_rows": len(data_rows), "column_index": column_index, "ascending": ascending}

    def get_table(self, table_idx: int) -> dict:
        """Get structured info for a single table by zero-based index.

        Args:
            table_idx: 0-based table index.

        Returns:
            {"index": int, "row_count": int, "col_count": int, "cells": list[list[str]]}

        Raises:
            IndexError: If table_idx is out of range.
        """
        tbl = self._get_table(table_idx)
        rows = []
        for tr in tbl.findall(f"{W}tr"):
            cells = []
            for tc in tr.findall(f"{W}tc"):
                cells.append(self._text(tc))
            rows.append(cells)
        col_count = len(rows[0]) if rows else 0
        return {
            "index": table_idx,
            "row_count": len(rows),
            "col_count": col_count,
            "cells": rows,
        }

    def get_cell_text(self, table_idx: int, row_idx: int, col_idx: int) -> dict:
        """Return text content of a specific cell.

        Args:
            table_idx: 0-based table index.
            row_idx: 0-based row index.
            col_idx: 0-based column index.

        Returns:
            {"table_index": int, "row_index": int, "col_index": int, "text": str}

        Raises:
            IndexError: If any index is out of range.
        """
        tbl = self._get_table(table_idx)
        rows = tbl.findall(f"{W}tr")
        if row_idx < 0 or row_idx >= len(rows):
            raise IndexError(f"Row index {row_idx} out of range")
        cols = rows[row_idx].findall(f"{W}tc")
        if col_idx < 0 or col_idx >= len(cols):
            raise IndexError(f"Column index {col_idx} out of range")
        return {
            "table_index": table_idx,
            "row_index": row_idx,
            "col_index": col_idx,
            "text": self._text(cols[col_idx]),
        }

    def copy_table(self, table_idx: int) -> dict:
        """Deep-copy a table and insert it immediately after the original.

        Args:
            table_idx: 0-based table index.

        Returns:
            {"source_index": int, "new_index": int}

        Raises:
            IndexError: If table_idx is out of range.
        """
        tbl = self._get_table(table_idx)
        new_tbl = copy.deepcopy(tbl)
        # Insert first so _new_para_id() can see already-assigned IDs in the copy
        tbl.addnext(new_tbl)
        # Regenerate paraIds — new_tbl is now in the document tree, so each
        # call to _new_para_id() sees the IDs set by previous iterations
        for p in new_tbl.iter(f"{W}p"):
            if p.get(f"{W14}paraId") is not None:
                p.set(f"{W14}paraId", self._new_para_id())
        self._mark("word/document.xml")
        return {"source_index": table_idx, "new_index": table_idx + 1}
