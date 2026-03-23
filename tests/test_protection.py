"""Tests for Phase 6 protection, properties write, and merge tools."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  set_document_protection
# ═══════════════════════════════════════════════════════════════════════════


class TestSetDocumentProtection:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_protect_tracked_changes(self):
        result = _j(server.set_document_protection("trackedChanges"))
        assert result["edit"] == "trackedChanges"
        assert result["enforcement"] == "1"

    def test_protect_with_password(self):
        result = _j(server.set_document_protection("readOnly", password="secret"))
        assert result["edit"] == "readOnly"
        assert result["has_password"] is True

    def test_protect_comments(self):
        result = _j(server.set_document_protection("comments"))
        assert result["edit"] == "comments"

    def test_unprotect(self):
        server.set_document_protection("trackedChanges")
        result = _j(server.set_document_protection("none"))
        assert result["edit"] == "none"

    def test_protect_replaces_existing(self):
        """Second call replaces the first protection."""
        server.set_document_protection("trackedChanges")
        result = _j(server.set_document_protection("readOnly"))
        assert result["edit"] == "readOnly"

    def test_protect_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.set_document_protection("trackedChanges")


# ═══════════════════════════════════════════════════════════════════════════
#  set_properties
# ═══════════════════════════════════════════════════════════════════════════


class TestSetProperties:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_set_title(self):
        result = _j(server.set_properties(title="New Title"))
        assert result["title"] == "New Title"
        # Verify via get_properties
        props = _j(server.get_properties())
        assert props["title"] == "New Title"

    def test_set_multiple_properties(self):
        result = _j(
            server.set_properties(
                title="Updated",
                creator="New Author",
                subject="New Subject",
            )
        )
        assert result["title"] == "Updated"
        assert result["creator"] == "New Author"
        assert result["subject"] == "New Subject"

    def test_set_description(self):
        result = _j(server.set_properties(description="New Desc"))
        assert result["description"] == "New Desc"

    def test_create_missing_element(self):
        """Setting a property that has no XML element creates it."""
        # Remove dc:description to test creation
        tree = server._doc._trees["docProps/core.xml"]
        from docx_mcp.document import DC

        desc = tree.find(f"{DC}description")
        if desc is not None:
            tree.remove(desc)
        result = _j(server.set_properties(description="Created"))
        assert result["description"] == "Created"

    def test_set_properties_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.set_properties(title="test")


# ═══════════════════════════════════════════════════════════════════════════
#  merge_documents
# ═══════════════════════════════════════════════════════════════════════════


def _make_simple_docx(path: Path, text: str, para_id: str = "00000001") -> None:
    """Build a minimal docx with one paragraph."""
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
            "word/_rels/document.xml.rels",
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>",
        )
        zf.writestr(
            "word/document.xml",
            '<?xml version="1.0"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            "<w:body>"
            f'<w:p w14:paraId="{para_id}" w14:textId="77777777">'
            f"<w:r><w:t>{text}</w:t></w:r></w:p>"
            "</w:body></w:document>",
        )


class TestMergeDocuments:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_merge_basic(self, tmp_path: Path):
        """Merge a simple document — content appears in combined doc."""
        src = tmp_path / "source.docx"
        _make_simple_docx(src, "Merged content", para_id="10000001")
        result = _j(server.merge_documents(str(src)))
        assert result["paragraphs_added"] >= 1
        # Verify content is searchable
        search = _j(server.search_text("Merged content"))
        assert len(search) >= 1

    def test_merge_para_id_collision(self, tmp_path: Path):
        """Source with colliding paraId gets remapped."""
        src = tmp_path / "source.docx"
        _make_simple_docx(src, "Collision text", para_id="00000001")  # same as fixture
        result = _j(server.merge_documents(str(src)))
        assert result["paragraphs_added"] >= 1
        # Validate no duplicates
        validation = _j(server.validate_paraids())
        assert validation["valid"] is True

    def test_merge_no_document_xml(self, tmp_path: Path):
        """Source DOCX missing word/document.xml raises ValueError."""
        src = tmp_path / "bad.docx"
        with zipfile.ZipFile(src, "w") as zf:
            zf.writestr("dummy.txt", "nothing")
        with pytest.raises(ValueError, match="no word/document.xml"):
            server.merge_documents(str(src))

    def test_merge_no_body(self, tmp_path: Path):
        """Source with document.xml but no w:body returns 0."""
        src = tmp_path / "nobody.docx"
        with zipfile.ZipFile(src, "w") as zf:
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"/>',
            )
        result = _j(server.merge_documents(str(src)))
        assert result["paragraphs_added"] == 0

    def test_merge_skips_sectpr(self, tmp_path: Path):
        """Source sectPr is not appended to target."""
        src = tmp_path / "withsect.docx"
        with zipfile.ZipFile(src, "w") as zf:
            zf.writestr(
                "word/document.xml",
                '<?xml version="1.0"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"'
                ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                "<w:body>"
                '<w:p w14:paraId="20000001" w14:textId="77777777">'
                "<w:r><w:t>With sectPr</w:t></w:r></w:p>"
                '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
                "</w:body></w:document>",
            )
        result = _j(server.merge_documents(str(src)))
        assert result["paragraphs_added"] == 1  # paragraph yes, sectPr no

    def test_merge_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            server.merge_documents("/nonexistent/file.docx")

    def test_merge_no_document(self, tmp_path: Path):
        server.close_document()
        src = tmp_path / "source.docx"
        _make_simple_docx(src, "text")
        with pytest.raises(RuntimeError, match="No document"):
            server.merge_documents(str(src))
