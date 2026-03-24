"""Tests for document creation (blank skeleton and template mode)."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document import CT, W14, DocxDocument, W
from tests.conftest import _build_fixture


class TestCreateBlank:
    def test_creates_valid_docx_file(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        assert out.exists()
        assert zipfile.is_zipfile(out)
        doc.close()

    def test_contains_required_parts(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        required = [
            "[Content_Types].xml",
            "word/document.xml",
            "word/styles.xml",
            "word/settings.xml",
            "word/footnotes.xml",
            "word/endnotes.xml",
            "word/numbering.xml",
            "word/header1.xml",
            "docProps/core.xml",
        ]
        for part in required:
            assert part in doc._trees, f"Missing part: {part}"
        doc.close()

    def test_document_has_one_paragraph_with_para_id(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        body = doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 1
        pid = paras[0].get(f"{W14}paraId")
        assert pid is not None
        assert len(pid) == 8
        assert int(pid, 16) < 0x80000000
        # Also has textId
        tid = paras[0].get(f"{W14}textId")
        assert tid is not None
        doc.close()

    def test_styles_include_headings_and_custom(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        styles_root = doc._trees["word/styles.xml"]
        style_ids = {s.get(f"{W}styleId") for s in styles_root.findall(f"{W}style")}
        # Built-in headings
        for i in range(1, 7):
            assert f"Heading{i}" in style_ids, f"Missing Heading{i}"
        # Custom styles
        assert "CodeBlock" in style_ids
        assert "BlockQuote" in style_ids
        # Lists
        assert "ListBullet" in style_ids or "ListParagraph" in style_ids
        doc.close()

    def test_numbering_has_multilevel_definitions(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        num_root = doc._trees["word/numbering.xml"]
        abstracts = num_root.findall(f"{W}abstractNum")
        assert len(abstracts) >= 2  # bullet + numbered
        # Each should have 9 levels (ilvl 0-8)
        for abstract in abstracts:
            lvls = abstract.findall(f"{W}lvl")
            assert len(lvls) == 9, f"abstractNum {abstract.get(f'{W}abstractNumId')} has {len(lvls)} levels"
        doc.close()

    def test_returns_opened_document(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        assert doc.workdir is not None
        assert len(doc._trees) > 0
        doc.close()

    def test_footnotes_have_separators(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        fn = doc._trees["word/footnotes.xml"]
        seps = [f for f in fn.findall(f"{W}footnote") if f.get(f"{W}id") in ("-1", "0")]
        assert len(seps) == 2
        doc.close()

    def test_endnotes_have_separators(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        en = doc._trees["word/endnotes.xml"]
        seps = [e for e in en.findall(f"{W}endnote") if e.get(f"{W}id") in ("-1", "0")]
        assert len(seps) == 2
        doc.close()

    def test_content_types_has_all_overrides(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        doc = DocxDocument.create(str(out))
        from docx_mcp.document import CT
        ct = doc._trees["[Content_Types].xml"]
        part_names = {o.get("PartName") for o in ct.findall(f"{CT}Override")}
        required_parts = [
            "/word/document.xml",
            "/word/styles.xml",
            "/word/settings.xml",
            "/word/footnotes.xml",
            "/word/endnotes.xml",
            "/word/numbering.xml",
            "/word/header1.xml",
            "/docProps/core.xml",
        ]
        for part in required_parts:
            assert part in part_names, f"Missing override: {part}"
        doc.close()


class TestCreateFromTemplate:
    """Tests for DocxDocument.create() with template_path parameter."""

    @pytest.fixture()
    def template_dotx(self, tmp_path: Path) -> Path:
        """Create a minimal .dotx template (same ZIP structure as .docx)."""
        path = tmp_path / "template.dotx"
        _build_fixture(path)
        return path

    def test_template_creates_docx_from_dotx(self, tmp_path: Path, template_dotx: Path):
        """Creating from .dotx produces a valid .docx with document.xml."""
        out = tmp_path / "from_template.docx"
        doc = DocxDocument.create(str(out), template_path=str(template_dotx))
        assert out.exists()
        assert "word/document.xml" in doc._trees
        doc.close()

    def test_template_preserves_existing_styles(self, tmp_path: Path, template_dotx: Path):
        """Template's own styles (like Heading1) survive the create process."""
        out = tmp_path / "from_template.docx"
        doc = DocxDocument.create(str(out), template_path=str(template_dotx))
        styles = doc._trees["word/styles.xml"]
        style_ids = {s.get(f"{W}styleId") for s in styles.findall(f"{W}style")}
        assert "Heading1" in style_ids
        doc.close()

    def test_template_adds_custom_styles_if_missing(self, tmp_path: Path, template_dotx: Path):
        """CodeBlock and BlockQuote are injected when template lacks them."""
        out = tmp_path / "from_template.docx"
        doc = DocxDocument.create(str(out), template_path=str(template_dotx))
        styles = doc._trees["word/styles.xml"]
        style_ids = {s.get(f"{W}styleId") for s in styles.findall(f"{W}style")}
        assert "CodeBlock" in style_ids
        assert "BlockQuote" in style_ids
        doc.close()

    def test_template_missing_raises(self, tmp_path: Path):
        """FileNotFoundError when template_path points to a nonexistent file."""
        out = tmp_path / "from_template.docx"
        with pytest.raises(FileNotFoundError):
            DocxDocument.create(str(out), template_path=str(tmp_path / "no_such.dotx"))

    def test_template_bootstraps_numbering_if_missing(self, tmp_path: Path, template_dotx: Path):
        """If template lacks numbering.xml, create() bootstraps it."""
        # Build a cleaned copy without numbering.xml
        clean = tmp_path / "no_numbering.dotx"
        with zipfile.ZipFile(template_dotx, "r") as src, \
             zipfile.ZipFile(clean, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                if item.filename == "word/numbering.xml":
                    continue
                dst.writestr(item, src.read(item.filename))

        out = tmp_path / "from_clean.docx"
        doc = DocxDocument.create(str(out), template_path=str(clean))
        # numbering.xml should now exist in the tree
        assert "word/numbering.xml" in doc._trees
        num_root = doc._trees["word/numbering.xml"]
        abstracts = num_root.findall(f"{W}abstractNum")
        assert len(abstracts) >= 2  # bullet + numbered

        # Content type override should have been added
        ct = doc._trees["[Content_Types].xml"]
        part_names = {o.get("PartName") for o in ct.findall(f"{CT}Override")}
        assert "/word/numbering.xml" in part_names
        doc.close()


class TestCreateDocumentTool:
    def test_creates_and_opens_document(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        result = json.loads(server.create_document(str(out)))
        assert "paragraph_count" in result  # returns get_info() output
        assert server._doc is not None
        assert server._doc.workdir is not None

    def test_closes_previous_document(self, tmp_path: Path, test_docx: Path):
        # Open an existing doc first
        server.open_document(str(test_docx))
        assert server._doc is not None
        old_workdir = server._doc.workdir
        # Create new — should close old
        out = tmp_path / "new.docx"
        server.create_document(str(out))
        assert not old_workdir.exists()  # old workdir cleaned up

    def test_with_template(self, tmp_path: Path):
        template = tmp_path / "tmpl.dotx"
        _build_fixture(template)
        out = tmp_path / "from_tmpl.docx"
        result = json.loads(server.create_document(str(out), template_path=str(template)))
        assert "paragraph_count" in result
