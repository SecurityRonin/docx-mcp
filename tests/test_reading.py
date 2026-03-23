"""Tests for Phase 1 read-only tools."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  get_tables
# ═══════════════════════════════════════════════════════════════════════════


class TestGetTables:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_tables(self):
        tables = _j(server.get_tables())
        assert len(tables) == 1
        t = tables[0]
        assert t["index"] == 0
        assert t["row_count"] == 3
        assert t["col_count"] == 2
        assert t["cells"][0] == ["Header A", "Header B"]
        assert t["cells"][1] == ["Row 1 A", "Row 1 B"]
        assert t["cells"][2] == ["Row 2 A", "Row 2 B"]

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.get_tables()


# ═══════════════════════════════════════════════════════════════════════════
#  get_styles
# ═══════════════════════════════════════════════════════════════════════════


class TestGetStyles:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_styles(self):
        styles = _j(server.get_styles())
        assert len(styles) == 4  # Heading1, Heading2, FootnoteReference, TableGrid
        ids = {s["id"] for s in styles}
        assert "Heading1" in ids
        assert "TableGrid" in ids

    def test_style_fields(self):
        styles = _j(server.get_styles())
        h1 = next(s for s in styles if s["id"] == "Heading1")
        assert h1["name"] == "heading 1"
        assert h1["type"] == "paragraph"
        assert h1["base_style"] == "Normal"

    def test_character_style_no_base(self):
        """Character style without basedOn returns empty base_style."""
        styles = _j(server.get_styles())
        fnref = next(s for s in styles if s["id"] == "FootnoteReference")
        assert fnref["type"] == "character"
        assert fnref["base_style"] == ""

    def test_no_styles_xml(self, tmp_path: Path):
        """Document without styles.xml returns empty list."""
        path = tmp_path / "nostyles.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:t>Hello</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        assert _j(server.get_styles()) == []


# ═══════════════════════════════════════════════════════════════════════════
#  get_headers_footers
# ═══════════════════════════════════════════════════════════════════════════


class TestGetHeadersFooters:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_headers(self):
        hf = _j(server.get_headers_footers())
        assert len(hf) >= 1
        h = hf[0]
        assert h["part"] == "word/header1.xml"
        assert h["location"] == "header"


# ═══════════════════════════════════════════════════════════════════════════
#  get_properties
# ═══════════════════════════════════════════════════════════════════════════


class TestGetProperties:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_properties(self):
        props = _j(server.get_properties())
        assert props["title"] == "Test Document"
        assert props["creator"] == "Test Author"
        assert props["subject"] == "Test Subject"
        assert props["description"] == "Test Description"
        assert props["last_modified_by"] == "Test Editor"
        assert props["revision"] == "3"
        assert "2025-01-01" in props["created"]
        assert "2025-06-15" in props["modified"]

    def test_no_core_xml(self, tmp_path: Path):
        """Document without core.xml returns empty dict."""
        path = tmp_path / "nocore.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:t>Hello</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        assert _j(server.get_properties()) == {}


# ═══════════════════════════════════════════════════════════════════════════
#  get_images
# ═══════════════════════════════════════════════════════════════════════════


class TestGetImages:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_no_document_xml(self, tmp_path: Path):
        """get_images returns [] when document.xml is missing."""
        path = tmp_path / "noimages.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:t>Hello</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        # Remove document.xml tree to simulate missing part
        server._doc._trees.pop("word/document.xml", None)
        assert _j(server.get_images()) == []

    def test_blip_without_embed(self, tmp_path: Path):
        """Blip element without r:embed attribute is skipped."""
        path = tmp_path / "noembed.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
                ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:drawing><a:blip/></w:drawing></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        assert _j(server.get_images()) == []

    def test_returns_images(self):
        images = _j(server.get_images())
        assert len(images) == 1
        img = images[0]
        assert img["rId"] == "rId6"
        assert img["filename"] == "image1.png"
        assert img["content_type"] == "image/png"
        assert img["width_emu"] == 914400
        assert img["height_emu"] == 914400


# ═══════════════════════════════════════════════════════════════════════════
#  get_endnotes
# ═══════════════════════════════════════════════════════════════════════════


class TestGetEndnotes:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_endnotes(self):
        endnotes = _j(server.get_endnotes())
        assert len(endnotes) == 1
        assert endnotes[0]["id"] == 1
        assert "Endnote reference" in endnotes[0]["text"]

    def test_no_endnotes_xml(self, tmp_path: Path):
        """Document without endnotes.xml returns empty list."""
        path = tmp_path / "noendnotes.docx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="rels"'
                ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Override PartName="/word/document.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.document.main+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/officeDocument"'
                ' Target="word/document.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="00000001" w14:textId="77777777">'
                "<w:r><w:t>Hello</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )
        server.open_document(str(path))
        assert _j(server.get_endnotes()) == []
