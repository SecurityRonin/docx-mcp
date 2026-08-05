"""Row banding and header-row styling.

The two composite tools from docs/IMPLEMENTATION-PLAN-table-and-deploy.md
§A.3. Each replaces a loop of per-cell calls that an agent would otherwise
have to issue one cell at a time.
"""

from __future__ import annotations

import pytest

from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.document.ooxml_order import TCPR_ORDER, find_out_of_order


def _doc_with_table(tmp_path, rows=4, cols=2):
    doc = DocxDocument.create(str(tmp_path / "t.docx"))
    tree = doc._tree("word/document.xml")
    para_id = tree.findall(f".//{W}p")[0].get(f"{W14}paraId")
    doc.add_table(para_id, rows, cols)
    return doc


def _row(doc, row, table_idx=0):
    return doc._get_table(table_idx).findall(f"{W}tr")[row]


def _fills(doc, row):
    """The w:shd fill of every cell in a row, None where unshaded."""
    out = []
    for tc in _row(doc, row).findall(f"{W}tc"):
        tc_pr = tc.find(f"{W}tcPr")
        shd = tc_pr.find(f"{W}shd") if tc_pr is not None else None
        out.append(shd.get(f"{W}fill") if shd is not None else None)
    return out


# ── set_table_banding ───────────────────────────────────────────────────────


class TestSetTableBanding:
    def test_alternates_fills_down_the_table(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=4)
        doc.set_table_banding(0, odd_color="EEEEEE", even_color="FFFFFF")
        # Row 0 is the header and is skipped by default.
        assert _fills(doc, 1) == ["EEEEEE", "EEEEEE"]
        assert _fills(doc, 2) == ["FFFFFF", "FFFFFF"]
        assert _fills(doc, 3) == ["EEEEEE", "EEEEEE"]

    def test_header_row_is_skipped_by_default(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=3)
        doc.set_table_banding(0)
        assert _fills(doc, 0) == [None, None]

    def test_header_row_is_banded_when_not_skipped(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=3)
        doc.set_table_banding(0, odd_color="EEEEEE", skip_header=False)
        assert _fills(doc, 0) == ["EEEEEE", "EEEEEE"]

    def test_counts_the_rows_it_shaded(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=4)
        assert doc.set_table_banding(0)["rows_shaded"] == 3

    def test_counts_every_row_when_header_not_skipped(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=4)
        assert doc.set_table_banding(0, skip_header=False)["rows_shaded"] == 4

    def test_reapplication_does_not_duplicate_shading(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=3)
        doc.set_table_banding(0)
        doc.set_table_banding(0)
        tc_pr = _row(doc, 1).findall(f"{W}tc")[0].find(f"{W}tcPr")
        assert len(tc_pr.findall(f"{W}shd")) == 1

    def test_recolouring_replaces_the_previous_fill(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=3)
        doc.set_table_banding(0, odd_color="111111")
        doc.set_table_banding(0, odd_color="222222")
        assert _fills(doc, 1) == ["222222", "222222"]

    def test_keeps_tcpr_in_schema_order(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=3)
        doc.set_cell_padding(0, 1, 0, top_mm=1.0)
        doc.set_table_banding(0)
        tc_pr = _row(doc, 1).findall(f"{W}tc")[0].find(f"{W}tcPr")
        assert find_out_of_order(tc_pr, TCPR_ORDER) == []

    def test_single_row_table_with_skip_header_shades_nothing(self, tmp_path):
        doc = _doc_with_table(tmp_path, rows=1)
        assert doc.set_table_banding(0)["rows_shaded"] == 0

    def test_bad_color_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_table_banding(0, odd_color="zzzzzz")

    def test_table_out_of_range_raises(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.set_table_banding(9)

    def test_marks_the_document_dirty(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc._modified.clear()
        doc.set_table_banding(0)
        assert "word/document.xml" in doc._modified


# ── style_header_row ────────────────────────────────────────────────────────


class TestStyleHeaderRow:
    def test_shades_every_header_cell(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.style_header_row(0, fill_color="4472C4")
        assert _fills(doc, 0) == ["4472C4", "4472C4"]

    def test_leaves_body_rows_alone(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.style_header_row(0)
        assert _fills(doc, 1) == [None, None]

    def test_bolds_the_header_text(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "Name", tracked=False)
        doc.style_header_row(0)
        run = _row(doc, 0).findall(f"{W}tc")[0].find(f"{W}p").find(f"{W}r")
        assert run.find(f"{W}rPr").find(f"{W}b") is not None

    def test_applies_the_text_color(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "Name", tracked=False)
        doc.style_header_row(0, text_color="FFFFFF")
        rpr = _row(doc, 0).findall(f"{W}tc")[0].find(f"{W}p").find(f"{W}r").find(f"{W}rPr")
        assert rpr.find(f"{W}color").get(f"{W}val") == "FFFFFF"

    def test_marks_the_row_to_repeat_across_pages(self, tmp_path):
        """A styled header that vanishes on page 2 is the usual complaint."""
        doc = _doc_with_table(tmp_path)
        doc.style_header_row(0)
        assert _row(doc, 0).find(f"{W}trPr").find(f"{W}tblHeader") is not None

    def test_bold_can_be_turned_off(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "Name", tracked=False)
        doc.style_header_row(0, bold=False)
        b = (
            _row(doc, 0)
            .findall(f"{W}tc")[0]
            .find(f"{W}p")
            .find(f"{W}r")
            .find(f"{W}rPr")
            .find(f"{W}b")
        )
        assert b.get(f"{W}val") == "0"

    def test_counts_the_cells_it_styled(self, tmp_path):
        doc = _doc_with_table(tmp_path, cols=3)
        assert doc.style_header_row(0)["cells_styled"] == 3

    def test_reapplication_does_not_duplicate(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.style_header_row(0)
        doc.style_header_row(0)
        tc_pr = _row(doc, 0).findall(f"{W}tc")[0].find(f"{W}tcPr")
        assert len(tc_pr.findall(f"{W}shd")) == 1
        assert len(_row(doc, 0).find(f"{W}trPr").findall(f"{W}tblHeader")) == 1

    def test_keeps_tcpr_in_schema_order(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_padding(0, 0, 0, top_mm=1.0)
        doc.style_header_row(0)
        tc_pr = _row(doc, 0).findall(f"{W}tc")[0].find(f"{W}tcPr")
        assert find_out_of_order(tc_pr, TCPR_ORDER) == []

    def test_empty_table_is_rejected(self, tmp_path):
        from lxml import etree

        doc = _doc_with_table(tmp_path)
        tbl = doc._get_table(0)
        for tr in tbl.findall(f"{W}tr"):
            tbl.remove(tr)
        assert isinstance(tbl, etree._Element)
        with pytest.raises(ValueError):
            doc.style_header_row(0)

    def test_bad_fill_color_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.style_header_row(0, fill_color="nope")

    def test_bad_text_color_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.style_header_row(0, text_color="nope")
