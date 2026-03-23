"""Tests for Phase 2 track-change management tools."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  accept_changes
# ═══════════════════════════════════════════════════════════════════════════


class TestAcceptChanges:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_accept_all(self):
        """Accept all: insertions kept, deletions removed."""
        # Create an insertion and a deletion
        server.insert_text("00000004", "INSERTED", author="Tester")
        server.delete_text("00000004", "30 days", author="Tester")
        result = _j(server.accept_changes())
        assert result["accepted"] >= 2
        assert result["scope"] == "all"
        # Verify the insertion text is now normal (no w:ins wrapper)
        para4 = _j(server.get_paragraph("00000004"))
        assert "INSERTED" in para4["text"]
        # Verify deletion text is gone
        assert "30 days" not in para4["text"]

    def test_accept_by_author(self):
        """Accept only changes by a specific author; other author's changes kept."""
        server.insert_text("00000004", "BY_ALICE", author="Alice")
        server.insert_text("00000004", "BY_BOB", author="Bob")
        server.delete_text("00000004", "30 days", author="Bob")
        result = _j(server.accept_changes(author="Alice"))
        assert result["accepted"] >= 1
        assert result["scope"] == "by_author"
        # Alice's insertion accepted; Bob's insertion and deletion still tracked
        para4 = _j(server.get_paragraph("00000004"))
        assert "BY_ALICE" in para4["text"]

    def test_accept_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.accept_changes()


# ═══════════════════════════════════════════════════════════════════════════
#  reject_changes
# ═══════════════════════════════════════════════════════════════════════════


class TestRejectChanges:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_reject_all(self):
        """Reject all: insertions removed, deletions restored."""
        server.insert_text("00000004", "SHOULD_VANISH", author="Tester")
        server.delete_text("00000004", "30 days", author="Tester")
        result = _j(server.reject_changes())
        assert result["rejected"] >= 2
        assert result["scope"] == "all"
        # Insertion text removed
        para4 = _j(server.get_paragraph("00000004"))
        assert "SHOULD_VANISH" not in para4["text"]
        # Deletion text restored
        assert "30 days" in para4["text"]

    def test_reject_by_author(self):
        """Reject only changes by a specific author; other author's changes kept."""
        server.insert_text("00000004", "BY_ALICE", author="Alice")
        server.insert_text("00000004", "BY_BOB", author="Bob")
        server.delete_text("00000004", "30 days", author="Bob")
        result = _j(server.reject_changes(author="Alice"))
        assert result["rejected"] >= 1
        assert result["scope"] == "by_author"
        # Alice's insertion rejected; Bob's insertion and deletion still tracked
        para4 = _j(server.get_paragraph("00000004"))
        assert "BY_ALICE" not in para4["text"]

    def test_reject_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.reject_changes()


# ═══════════════════════════════════════════════════════════════════════════
#  set_formatting
# ═══════════════════════════════════════════════════════════════════════════


class TestSetFormatting:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_bold(self):
        result = _j(server.set_formatting("00000004", "contract", bold=True))
        assert result["formatted"] is True

    def test_italic(self):
        result = _j(server.set_formatting("00000004", "30 days", italic=True))
        assert result["formatted"] is True

    def test_bold_italic_color(self):
        result = _j(
            server.set_formatting("00000004", "effective", bold=True, italic=True, color="FF0000")
        )
        assert result["formatted"] is True

    def test_underline(self):
        result = _j(server.set_formatting("00000004", "contract", underline="single"))
        assert result["formatted"] is True

    def test_format_run_with_existing_rpr(self):
        """Formatting text in a run that already has rPr preserves old rPr in rPrChange."""
        # Paragraph 00000005 has <w:rPr><w:b/></w:rPr>
        result = _j(server.set_formatting("00000005", "Final", italic=True))
        assert result["formatted"] is True

    def test_format_skips_run_without_text(self):
        """Formatting skips runs that have no w:t element (e.g., endnote refs)."""
        # Paragraph 00000005 has a text run then an endnote reference run (no w:t).
        # Insert an empty run before the text run to force the skip-path.
        from lxml import etree

        from docx_mcp.document import W

        doc = server._doc._trees["word/document.xml"]
        para = doc.find(
            f'.//{W}p[@{W.replace("{", "").replace("}", "")}14:paraId="00000005"]'.replace(
                W.replace("{", "").replace("}", "") + "14:",
                "{http://schemas.microsoft.com/office/word/2010/wordml}",
            )
        )
        # Simpler: just find the paragraph by iterating
        for p in doc.iter(f"{W}p"):
            pid = p.get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId")
            if pid == "00000005":
                para = p
                break
        # Insert an empty run (no w:t) before existing children
        empty_run = etree.Element(f"{W}r")
        etree.SubElement(empty_run, f"{W}rPr")
        para.insert(0, empty_run)

        result = _j(server.set_formatting("00000005", "Final", italic=True))
        assert result["formatted"] is True

    def test_text_not_found(self):
        with pytest.raises(ValueError, match="not found"):
            server.set_formatting("00000004", "NONEXISTENT", bold=True)

    def test_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.set_formatting("DEADBEEF", "contract", bold=True)

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.set_formatting("00000004", "text", bold=True)
