"""Table geometry: cell padding, default cell margins, layout mode, table width.

See docs/IMPLEMENTATION-PLAN-table-and-deploy.md §A.3. Units at the tool
boundary are mm (consistent with the existing set_cell_width) or percent,
never raw dxa.
"""

from __future__ import annotations

import pytest

from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.document.ooxml_order import TBLPR_ORDER, TCMAR_ORDER, TCPR_ORDER, find_out_of_order

# 2.54 mm is exactly 144 twentieths of a point, so the conversion is testable
# without rounding noise.
MM = 2.54
DXA = 144


def _doc_with_table(tmp_path, rows=3, cols=3):
    doc = DocxDocument.create(str(tmp_path / "t.docx"))
    tree = doc._tree("word/document.xml")
    para_id = tree.findall(f".//{W}p")[0].get(f"{W14}paraId")
    doc.add_table(para_id, rows, cols)
    return doc


def _cell(doc, row, col, table_idx=0):
    tbl = doc._get_table(table_idx)
    return tbl.findall(f"{W}tr")[row].findall(f"{W}tc")[col]


def _tcpr(doc, row, col):
    return _cell(doc, row, col).find(f"{W}tcPr")


def _tblpr(doc, table_idx=0):
    return doc._get_table(table_idx).find(f"{W}tblPr")


def _names(parent):
    from lxml import etree

    return [etree.QName(c).localname for c in parent if isinstance(c.tag, str)]


# ── set_cell_padding ────────────────────────────────────────────────────────


class TestSetCellPadding:
    def test_writes_tcmar_with_all_four_sides(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_padding(0, 0, 0, top_mm=MM, bottom_mm=MM, left_mm=MM, right_mm=MM)
        mar = _tcpr(doc, 0, 0).find(f"{W}tcMar")
        assert _names(mar) == list(TCMAR_ORDER)

    def test_converts_mm_to_dxa(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_padding(0, 0, 0, top_mm=MM)
        top = _tcpr(doc, 0, 0).find(f"{W}tcMar").find(f"{W}top")
        assert top.get(f"{W}w") == str(DXA)
        assert top.get(f"{W}type") == "dxa"

    def test_omitted_sides_are_not_written(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_padding(0, 0, 0, top_mm=MM)
        assert _names(_tcpr(doc, 0, 0).find(f"{W}tcMar")) == ["top"]

    def test_reapplication_replaces_rather_than_duplicates(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_padding(0, 0, 0, top_mm=MM)
        doc.set_cell_padding(0, 0, 0, top_mm=MM * 2)
        mar = _tcpr(doc, 0, 0).find(f"{W}tcMar")
        assert len(mar.findall(f"{W}top")) == 1
        assert mar.find(f"{W}top").get(f"{W}w") == str(DXA * 2)

    def test_sides_land_in_schema_order_regardless_of_call_order(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_padding(0, 0, 0, right_mm=MM)
        doc.set_cell_padding(0, 0, 0, top_mm=MM)
        doc.set_cell_padding(0, 0, 0, bottom_mm=MM)
        doc.set_cell_padding(0, 0, 0, left_mm=MM)
        assert _names(_tcpr(doc, 0, 0).find(f"{W}tcMar")) == list(TCMAR_ORDER)

    def test_returns_the_padding_it_applied(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        got = doc.set_cell_padding(0, 1, 2, top_mm=MM)
        assert got["table_idx"] == 0
        assert got["row_idx"] == 1
        assert got["col_idx"] == 2
        assert got["padding_mm"]["top"] == MM

    def test_marks_the_document_dirty(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc._modified.clear()
        doc.set_cell_padding(0, 0, 0, top_mm=MM)
        assert "word/document.xml" in doc._modified

    def test_row_out_of_range_raises(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_cell_padding(0, 99, 0, top_mm=MM)

    def test_column_out_of_range_raises(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_cell_padding(0, 0, 99, top_mm=MM)

    def test_negative_padding_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_cell_padding(0, 0, 0, top_mm=-1.0)

    def test_no_sides_given_is_rejected(self, tmp_path):
        """A call that would do nothing is a caller mistake, not a no-op."""
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_cell_padding(0, 0, 0)


class TestCellPaddingKeepsTcPrOrdered:
    def test_tcmar_lands_in_schema_position(self, tmp_path):
        """Shade, merge, then pad — tcPr must still validate."""
        doc = _doc_with_table(tmp_path)
        doc.set_cell_shading(0, 0, 0, "FF0000")
        doc.merge_cells(0, 0, 0, 0, 1)
        doc.set_cell_padding(0, 0, 0, top_mm=MM)
        assert find_out_of_order(_tcpr(doc, 0, 0), TCPR_ORDER) == []


# ── set_table_cell_margins ──────────────────────────────────────────────────


class TestSetTableCellMargins:
    def test_writes_tblcellmar(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_cell_margins(0, top_mm=MM, bottom_mm=MM, left_mm=MM, right_mm=MM)
        mar = _tblpr(doc).find(f"{W}tblCellMar")
        assert _names(mar) == list(TCMAR_ORDER)

    def test_converts_mm_to_dxa(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_cell_margins(0, left_mm=MM)
        left = _tblpr(doc).find(f"{W}tblCellMar").find(f"{W}left")
        assert left.get(f"{W}w") == str(DXA)
        assert left.get(f"{W}type") == "dxa"

    def test_reapplication_replaces_rather_than_duplicates(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_cell_margins(0, left_mm=MM)
        doc.set_table_cell_margins(0, left_mm=MM * 2)
        mar = _tblpr(doc).find(f"{W}tblCellMar")
        assert len(mar.findall(f"{W}left")) == 1

    def test_keeps_tblpr_in_schema_order(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_cell_margins(0, top_mm=MM)
        doc.set_table_borders(0)
        assert find_out_of_order(_tblpr(doc), TBLPR_ORDER) == []

    def test_returns_the_margins_it_applied(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        got = doc.set_table_cell_margins(0, top_mm=MM)
        assert got["margins_mm"]["top"] == MM

    def test_table_out_of_range_raises(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_table_cell_margins(9, top_mm=MM)


# ── set_table_layout ────────────────────────────────────────────────────────


class TestSetTableLayout:
    def test_fixed_writes_tbllayout(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_layout(0, "fixed")
        assert _tblpr(doc).find(f"{W}tblLayout").get(f"{W}type") == "fixed"

    def test_autofit_writes_tbllayout(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_layout(0, "autofit")
        assert _tblpr(doc).find(f"{W}tblLayout").get(f"{W}type") == "autofit"

    def test_autofit_sets_table_width_to_auto(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, 100, "mm")
        doc.set_table_layout(0, "autofit")
        tbl_w = _tblpr(doc).find(f"{W}tblW")
        assert tbl_w.get(f"{W}type") == "auto"
        assert tbl_w.get(f"{W}w") == "0"

    def test_autofit_clears_explicit_cell_widths(self, tmp_path):
        """Fixed tcW values would otherwise defeat auto-fit."""
        doc = _doc_with_table(tmp_path)
        doc.set_cell_width(0, 0, 0, 50.0)
        doc.set_table_layout(0, "autofit")
        tc_w = _tcpr(doc, 0, 0).find(f"{W}tcW")
        assert tc_w.get(f"{W}type") == "auto"
        assert tc_w.get(f"{W}w") == "0"

    def test_fixed_leaves_cell_widths_alone(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_width(0, 0, 0, 50.0)
        before = _tcpr(doc, 0, 0).find(f"{W}tcW").get(f"{W}w")
        doc.set_table_layout(0, "fixed")
        assert _tcpr(doc, 0, 0).find(f"{W}tcW").get(f"{W}w") == before

    def test_switching_modes_does_not_duplicate(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_layout(0, "fixed")
        doc.set_table_layout(0, "autofit")
        assert len(_tblpr(doc).findall(f"{W}tblLayout")) == 1

    def test_unknown_mode_names_the_offending_value(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError, match="sideways"):
            doc.set_table_layout(0, "sideways")

    def test_returns_the_mode(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        assert doc.set_table_layout(0, "fixed")["mode"] == "fixed"


# ── set_table_width ─────────────────────────────────────────────────────────


class TestSetTableWidth:
    def test_mm_writes_dxa(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, MM, "mm")
        tbl_w = _tblpr(doc).find(f"{W}tblW")
        assert tbl_w.get(f"{W}w") == str(DXA)
        assert tbl_w.get(f"{W}type") == "dxa"

    def test_percent_writes_fiftieths(self, tmp_path):
        """ST_TableWidth pct is expressed in fiftieths of a percent."""
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, 50.0, "percent")
        tbl_w = _tblpr(doc).find(f"{W}tblW")
        assert tbl_w.get(f"{W}w") == "2500"
        assert tbl_w.get(f"{W}type") == "pct"

    def test_full_width_percent(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, 100.0, "percent")
        assert _tblpr(doc).find(f"{W}tblW").get(f"{W}w") == "5000"

    def test_auto_ignores_the_width_value(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, None, "auto")
        tbl_w = _tblpr(doc).find(f"{W}tblW")
        assert tbl_w.get(f"{W}type") == "auto"
        assert tbl_w.get(f"{W}w") == "0"

    def test_reapplication_replaces_rather_than_duplicates(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, 10.0, "mm")
        doc.set_table_width(0, 20.0, "mm")
        assert len(_tblpr(doc).findall(f"{W}tblW")) == 1

    def test_percent_above_one_hundred_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_table_width(0, 150.0, "percent")

    def test_negative_width_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_table_width(0, -5.0, "mm")

    def test_missing_width_for_a_measured_unit_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError, match="width"):
            doc.set_table_width(0, None, "mm")

    def test_unknown_unit_names_the_offending_value(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError, match="furlongs"):
            doc.set_table_width(0, 1.0, "furlongs")

    def test_returns_what_it_applied(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        got = doc.set_table_width(0, 50.0, "percent")
        assert got == {"table_idx": 0, "width": 50.0, "unit": "percent"}

    def test_keeps_tblpr_in_schema_order(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_table_width(0, 50.0, "percent")
        doc.set_table_layout(0, "fixed")
        doc.set_table_cell_margins(0, top_mm=MM)
        assert find_out_of_order(_tblpr(doc), TBLPR_ORDER) == []
