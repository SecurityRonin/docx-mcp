"""Tests for paragraph CRUD + border/shading methods."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


class TestParagraphCRUD:
    def _open(self, path: Path) -> None:
        server._doc = None
        server.open_document(str(path))

    # ── insert_paragraph ───────────────────────────────────────────────────

    def test_insert_paragraph_adds_after_target(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.insert_paragraph("00000002", "Inserted text"))
        assert result["text"] == "Inserted text"
        new_pid = result["para_id"]

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
        root = doc._tree("word/document.xml")
        body = root.find(f"{W}body")
        paras = [p for p in body if p.tag == f"{W}p"]
        ids = [p.get(f"{W14}paraId") for p in paras]
        idx_target = ids.index("00000002")
        assert ids[idx_target + 1] == new_pid

    def test_insert_paragraph_with_style(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.insert_paragraph("00000001", "Styled para", style="Heading2"))
        new_pid = result["para_id"]

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, new_pid)
        ppr = para.find(f"{W}pPr")
        assert ppr is not None
        pstyle = ppr.find(f"{W}pStyle")
        assert pstyle is not None
        assert pstyle.get(f"{W}val") == "Heading2"

    def test_insert_paragraph_nonexistent_raises(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.insert_paragraph("DEADBEEF", "text")

    def test_insert_paragraph_text_in_run(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.insert_paragraph("00000003", "Run text"))
        new_pid = result["para_id"]

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, new_pid)
        texts = [t.text for t in para.iter(f"{W}t") if t.text]
        assert "Run text" in texts

    def test_insert_paragraph_fresh_para_id(self, test_docx: Path) -> None:
        self._open(test_docx)
        r1 = json.loads(server.insert_paragraph("00000002", "first"))
        r2 = json.loads(server.insert_paragraph("00000003", "second"))
        assert r1["para_id"] != r2["para_id"]

    # ── update_paragraph ───────────────────────────────────────────────────

    def test_update_paragraph_changes_text(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.update_paragraph("00000002", text="New content"))
        assert result["para_id"] == "00000002"

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000002")
        texts = [t.text for t in para.iter(f"{W}t") if t.text]
        assert texts == ["New content"]

    def test_update_paragraph_changes_style(self, test_docx: Path) -> None:
        self._open(test_docx)
        server.update_paragraph("00000002", style="Heading1")

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000002")
        ppr = para.find(f"{W}pPr")
        assert ppr is not None
        pstyle = ppr.find(f"{W}pStyle")
        assert pstyle is not None
        assert pstyle.get(f"{W}val") == "Heading1"

    def test_update_paragraph_nonexistent_raises(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.update_paragraph("DEADBEEF", text="x")

    # ── delete_paragraph ───────────────────────────────────────────────────

    def test_delete_paragraph_removes_it(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.delete_paragraph("00000004"))
        assert result["deleted"] == "00000004"

        doc = server._doc
        root = doc._tree("word/document.xml")
        found = doc._find_para(root, "00000004")
        assert found is None

    def test_delete_paragraph_nonexistent_raises(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.delete_paragraph("DEADBEEF")

    # ── set_paragraph_border ───────────────────────────────────────────────

    def test_set_paragraph_border_sets_elements(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(
            server.set_paragraph_border("00000002", ["top", "bottom"], color="FF0000", size=8)
        )
        assert result["para_id"] == "00000002"
        assert set(result["sides"]) == {"top", "bottom"}

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000002")
        ppr = para.find(f"{W}pPr")
        pbdr = ppr.find(f"{W}pBdr")
        assert pbdr is not None
        top = pbdr.find(f"{W}top")
        assert top is not None
        assert top.get(f"{W}color") == "FF0000"
        assert top.get(f"{W}sz") == "8"
        assert top.get(f"{W}val") == "single"
        bottom = pbdr.find(f"{W}bottom")
        assert bottom is not None
        left = pbdr.find(f"{W}left")
        assert left is None

    # ── set_paragraph_shading ──────────────────────────────────────────────

    def test_set_paragraph_shading_sets_fill(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_paragraph_shading("00000003", "FFFF00"))
        assert result["para_id"] == "00000003"
        assert result["fill_color"] == "FFFF00"

        doc = server._doc
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000003")
        ppr = para.find(f"{W}pPr")
        assert ppr is not None
        shd = ppr.find(f"{W}shd")
        assert shd is not None
        assert shd.get(f"{W}fill") == "FFFF00"
        assert shd.get(f"{W}val") == "clear"
        assert shd.get(f"{W}color") == "auto"

    # ── update_paragraph tracked (issue #9) ──────────────────────────────

    def test_update_paragraph_tracked_produces_del_and_ins(self, test_docx: Path) -> None:
        """tracked=True wraps old text in w:del and new text in w:ins."""
        self._open(test_docx)
        server.update_paragraph(
            "00000002", text="Tracked replacement", tracked=True, author="Tester"
        )
        changes = json.loads(server.get_tracked_changes())
        assert len(changes) >= 2, f"expected ≥2 changes, got {len(changes)}"
        assert any(c["type"] == "deletion" for c in changes), "no deletion found"
        assert any(c["type"] == "insertion" for c in changes), "no insertion found"

    def test_update_paragraph_tracked_false_unchanged(self, test_docx: Path) -> None:
        """tracked=False (default) produces no tracked changes — silent replace."""
        self._open(test_docx)
        server.update_paragraph("00000002", text="Silent replacement")
        changes = json.loads(server.get_tracked_changes())
        assert len(changes) == 0, "no tracked changes expected for silent mode"
        body = json.loads(server.get_body_text())["body"]
        assert "Silent replacement" in body, "text should still be replaced"
