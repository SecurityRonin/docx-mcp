"""Tests for Phase 5 document structure tools."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import W


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  add_page_break
# ═══════════════════════════════════════════════════════════════════════════


class TestAddPageBreak:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_page_break(self):
        result = _j(server.add_page_break("00000004"))
        assert result["para_id"]  # new paragraph was created
        # Verify the break element exists in the XML
        doc = server._doc._trees["word/document.xml"]
        new_p = server._doc._find_para(doc, result["para_id"])
        assert new_p is not None
        br = new_p.find(f".//{W}br")
        assert br is not None
        assert br.get(f"{W}type") == "page"

    def test_page_break_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_page_break("DEADBEEF")

    def test_page_break_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_page_break("00000004")


# ═══════════════════════════════════════════════════════════════════════════
#  add_section_break
# ═══════════════════════════════════════════════════════════════════════════


class TestAddSectionBreak:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_section_break_next_page(self):
        result = _j(server.add_section_break("00000004", break_type="nextPage"))
        assert result["break_type"] == "nextPage"
        assert result["para_id"] == "00000004"
        # Verify sectPr in paragraph pPr
        doc = server._doc._trees["word/document.xml"]
        p = server._doc._find_para(doc, "00000004")
        sect_pr = p.find(f"{W}pPr/{W}sectPr")
        assert sect_pr is not None
        type_el = sect_pr.find(f"{W}type")
        assert type_el is not None
        assert type_el.get(f"{W}val") == "nextPage"

    def test_section_break_continuous(self):
        result = _j(server.add_section_break("00000004", break_type="continuous"))
        assert result["break_type"] == "continuous"
        doc = server._doc._trees["word/document.xml"]
        p = server._doc._find_para(doc, "00000004")
        sect_pr = p.find(f"{W}pPr/{W}sectPr")
        type_el = sect_pr.find(f"{W}type")
        assert type_el.get(f"{W}val") == "continuous"

    def test_section_break_creates_ppr(self):
        """Paragraph without existing pPr gets one created."""
        result = _j(server.add_section_break("00000006", break_type="nextPage"))
        assert result["para_id"] == "00000006"
        doc = server._doc._trees["word/document.xml"]
        p = server._doc._find_para(doc, "00000006")
        assert p.find(f"{W}pPr/{W}sectPr") is not None

    def test_section_break_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_section_break("DEADBEEF")

    def test_section_break_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_section_break("00000004")


# ═══════════════════════════════════════════════════════════════════════════
#  set_section_properties
# ═══════════════════════════════════════════════════════════════════════════


class TestSetSectionProperties:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_set_body_section_properties(self):
        """Modify the body-level (last) section properties."""
        # First, ensure there's a body sectPr by adding one
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = etree.SubElement(body, f"{W}sectPr")
        pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")
        pg_sz.set(f"{W}w", "12240")
        pg_sz.set(f"{W}h", "15840")

        result = _j(
            server.set_section_properties(
                width=15840,
                height=12240,
                orientation="landscape",
            )
        )
        assert result["width"] == 15840
        assert result["height"] == 12240
        assert result["orientation"] == "landscape"

    def test_set_section_properties_with_margins(self):
        """Set margins via set_section_properties."""
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = etree.SubElement(body, f"{W}sectPr")
        etree.SubElement(sect_pr, f"{W}pgSz")

        result = _j(
            server.set_section_properties(
                margin_top=1440,
                margin_bottom=1440,
                margin_left=1800,
                margin_right=1800,
            )
        )
        assert result["margin_top"] == 1440

    def test_set_section_properties_para_section(self):
        """Modify properties of a paragraph-level section break."""
        # Add a section break first
        server.add_section_break("00000004", break_type="nextPage")
        result = _j(
            server.set_section_properties(
                para_id="00000004",
                width=15840,
                height=12240,
            )
        )
        assert result["width"] == 15840

    def test_set_section_no_sectpr(self):
        """Body without sectPr — auto-creates one."""
        result = _j(server.set_section_properties(width=12240, height=15840))
        assert result["width"] == 12240

    def test_set_section_para_bad_para(self):
        """Non-existent paragraph raises error."""
        with pytest.raises(ValueError, match="not found"):
            server.set_section_properties(para_id="DEADBEEF", width=12240)

    def test_set_section_para_no_sectpr(self):
        """Paragraph without sectPr raises error."""
        with pytest.raises(ValueError, match="No section break"):
            server.set_section_properties(para_id="00000004", width=12240)

    def test_set_section_properties_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.set_section_properties(width=12240)


# ═══════════════════════════════════════════════════════════════════════════
#  add_cross_reference
# ═══════════════════════════════════════════════════════════════════════════


class TestAddCrossReference:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_cross_ref_to_heading(self):
        """Cross-reference to a heading creates bookmark and hyperlink."""
        result = _j(
            server.add_cross_reference(
                source_para_id="00000004",
                target_para_id="00000001",
                text="see Introduction",
            )
        )
        assert result["bookmark_name"]
        assert result["source_para_id"] == "00000004"
        assert result["target_para_id"] == "00000001"

    def test_cross_ref_existing_bookmark(self):
        """Cross-reference to paragraph with existing bookmark reuses it."""
        result = _j(
            server.add_cross_reference(
                source_para_id="00000002",
                target_para_id="00000004",  # has bookmark "section_bg"
                text="see Background",
            )
        )
        assert result["bookmark_name"] == "section_bg"

    def test_cross_ref_bad_source(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_cross_reference(
                source_para_id="DEADBEEF",
                target_para_id="00000001",
                text="link",
            )

    def test_cross_ref_bad_target(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_cross_reference(
                source_para_id="00000004",
                target_para_id="DEADBEEF",
                text="link",
            )

    def test_cross_ref_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.add_cross_reference(
                source_para_id="00000004",
                target_para_id="00000001",
                text="link",
            )
