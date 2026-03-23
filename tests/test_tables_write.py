"""Tests for Phase 3 table write tools."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  add_table
# ═══════════════════════════════════════════════════════════════════════════


class TestAddTable:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_add_table(self):
        result = _j(server.add_table("00000004", rows=2, cols=3))
        assert result["inserted"] is True
        assert result["rows"] == 2
        assert result["cols"] == 3
        # Now there should be 2 tables total
        tables = _j(server.get_tables())
        assert len(tables) == 2
        # New table is at index 0 (inserted before existing table in document order)
        new_tbl = tables[result["table_index"]]
        assert new_tbl["row_count"] == 2
        assert new_tbl["col_count"] == 3

    def test_add_table_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_table("DEADBEEF", rows=2, cols=2)

    def test_add_table_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_table("00000004", rows=2, cols=2)


# ═══════════════════════════════════════════════════════════════════════════
#  modify_cell
# ═══════════════════════════════════════════════════════════════════════════


class TestModifyCell:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_modify_cell(self):
        result = _j(server.modify_cell(0, 1, 0, "Updated"))
        assert result["modified"] is True
        assert result["text"] == "Updated"
        # Verify via get_tables that the cell now contains the new text
        tables = _j(server.get_tables())
        # New text appears in the cell (alongside tracked deletion of old)
        cell_text = tables[0]["cells"][1][0]
        assert "Updated" in cell_text

    def test_modify_cell_bad_table(self):
        with pytest.raises(IndexError):
            server.modify_cell(99, 0, 0, "text")

    def test_modify_cell_bad_row(self):
        with pytest.raises(IndexError):
            server.modify_cell(0, 99, 0, "text")

    def test_modify_cell_bad_col(self):
        with pytest.raises(IndexError):
            server.modify_cell(0, 0, 99, "text")

    def test_modify_cell_with_rpr(self):
        """Modify cell where run has rPr — preserved in deletion markup."""
        from lxml import etree

        from docx_mcp.document import W, W14

        # Add rPr to the first cell's run
        doc = server._doc._trees["word/document.xml"]
        for tbl in doc.iter(f"{W}tbl"):
            tr = tbl.find(f"{W}tr")
            tc = tr.find(f"{W}tc")
            p = tc.find(f"{W}p")
            r = p.find(f"{W}r")
            rpr = etree.SubElement(r, f"{W}rPr")
            etree.SubElement(rpr, f"{W}b")
            # Move rPr before t
            r.remove(rpr)
            r.insert(0, rpr)
            break
        result = _j(server.modify_cell(0, 0, 0, "Updated"))
        assert result["modified"] is True

    def test_modify_cell_empty_run(self):
        """Modify cell where paragraph has a run without w:t — skipped."""
        from lxml import etree

        from docx_mcp.document import W

        doc = server._doc._trees["word/document.xml"]
        for tbl in doc.iter(f"{W}tbl"):
            tr = tbl.find(f"{W}tr")
            tc = tr.find(f"{W}tc")
            p = tc.find(f"{W}p")
            # Insert empty run before existing content
            empty_r = etree.Element(f"{W}r")
            p.insert(0, empty_r)
            break
        result = _j(server.modify_cell(0, 0, 0, "Replaced"))
        assert result["modified"] is True

    def test_modify_cell_no_paragraph(self):
        """Modify cell where tc has no w:p — paragraph created."""
        from lxml import etree

        from docx_mcp.document import W

        doc = server._doc._trees["word/document.xml"]
        for tbl in doc.iter(f"{W}tbl"):
            tr = tbl.find(f"{W}tr")
            tc = tr.find(f"{W}tc")
            # Remove all paragraphs from cell
            for p in list(tc.findall(f"{W}p")):
                tc.remove(p)
            break
        result = _j(server.modify_cell(0, 0, 0, "New"))
        assert result["modified"] is True

    def test_modify_cell_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.modify_cell(0, 0, 0, "text")


# ═══════════════════════════════════════════════════════════════════════════
#  add_table_row
# ═══════════════════════════════════════════════════════════════════════════


class TestAddTableRow:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_append_row(self):
        result = _j(server.add_table_row(0))
        assert result["inserted"] is True
        assert result["row_index"] == 3  # appended after existing 3 rows
        tables = _j(server.get_tables())
        assert tables[0]["row_count"] == 4

    def test_append_row_with_cells(self):
        result = _j(server.add_table_row(0, cells=["A", "B"]))
        assert result["inserted"] is True
        tables = _j(server.get_tables())
        last_row = tables[0]["cells"][-1]
        assert last_row == ["A", "B"]

    def test_insert_row_at_index(self):
        result = _j(server.add_table_row(0, row_idx=1))
        assert result["inserted"] is True
        assert result["row_index"] == 1
        tables = _j(server.get_tables())
        assert tables[0]["row_count"] == 4

    def test_bad_table_index(self):
        with pytest.raises(IndexError):
            server.add_table_row(99)

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_table_row(0)


# ═══════════════════════════════════════════════════════════════════════════
#  delete_table_row
# ═══════════════════════════════════════════════════════════════════════════


class TestDeleteTableRow:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_delete_row(self):
        result = _j(server.delete_table_row(0, 1))
        assert result["deleted"] is True
        assert result["row_index"] == 1

    def test_bad_table_index(self):
        with pytest.raises(IndexError):
            server.delete_table_row(99, 0)

    def test_delete_row_with_rpr(self):
        """Delete row where cells have runs with rPr."""
        from lxml import etree

        from docx_mcp.document import W

        doc = server._doc._trees["word/document.xml"]
        for tbl in doc.iter(f"{W}tbl"):
            tr = list(tbl.findall(f"{W}tr"))[1]
            tc = tr.find(f"{W}tc")
            p = tc.find(f"{W}p")
            r = p.find(f"{W}r")
            rpr = etree.SubElement(r, f"{W}rPr")
            etree.SubElement(rpr, f"{W}i")
            r.remove(rpr)
            r.insert(0, rpr)
            break
        result = _j(server.delete_table_row(0, 1))
        assert result["deleted"] is True

    def test_delete_row_skips_empty_run(self):
        """Delete row where cell has a run without w:t."""
        from lxml import etree

        from docx_mcp.document import W

        doc = server._doc._trees["word/document.xml"]
        for tbl in doc.iter(f"{W}tbl"):
            tr = list(tbl.findall(f"{W}tr"))[2]
            tc = tr.find(f"{W}tc")
            p = tc.find(f"{W}p")
            empty_r = etree.Element(f"{W}r")
            p.insert(0, empty_r)
            break
        result = _j(server.delete_table_row(0, 2))
        assert result["deleted"] is True

    def test_bad_row_index(self):
        with pytest.raises(IndexError):
            server.delete_table_row(0, 99)

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.delete_table_row(0, 0)
