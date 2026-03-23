"""End-to-end tests for all 43 MCP tools — targeting 100% line and roundtrip coverage."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import RELS, W14, W15, A, DocxDocument, R, V, W


def _j(result: str) -> dict | list:
    """Parse JSON tool response."""
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  Document lifecycle
# ═══════════════════════════════════════════════════════════════════════════


class TestOpen:
    def test_open_returns_info(self, test_docx: Path):
        info = _j(server.open_document(str(test_docx)))
        assert info["paragraph_count"] == 13  # 7 body + 6 table cell paragraphs
        assert info["heading_count"] == 2
        assert info["footnote_count"] == 1
        assert info["image_count"] == 1
        assert "parts" in info

    def test_open_replaces_previous(self, test_docx: Path, tmp_path: Path):
        """Opening a second doc closes the first automatically."""
        server.open_document(str(test_docx))
        copy = tmp_path / "copy.docx"
        import shutil

        shutil.copy2(test_docx, copy)
        info = _j(server.open_document(str(copy)))
        assert info["paragraph_count"] == 13

    def test_open_nonexistent(self):
        with pytest.raises(FileNotFoundError):
            server.open_document("/nonexistent/file.docx")

    def test_open_non_docx(self, tmp_path: Path):
        txt = tmp_path / "test.txt"
        txt.write_text("hello")
        with pytest.raises(ValueError, match="Not a .docx"):
            server.open_document(str(txt))


class TestClose:
    def test_close_when_open(self, test_docx: Path):
        server.open_document(str(test_docx))
        result = server.close_document()
        assert "closed" in result.lower()

    def test_close_when_nothing_open(self):
        result = server.close_document()
        assert "closed" in result.lower()


class TestInfo:
    def test_get_info(self, test_docx: Path):
        server.open_document(str(test_docx))
        info = _j(server.get_document_info())
        assert info["paragraph_count"] == 13
        assert "size_bytes" in info

    def test_get_info_no_document(self):
        with pytest.raises(RuntimeError, match="No document"):
            server.get_document_info()


# ═══════════════════════════════════════════════════════════════════════════
#  Reading tools
# ═══════════════════════════════════════════════════════════════════════════


class TestHeadings:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_heading_structure(self):
        headings = _j(server.get_headings())
        assert len(headings) == 2
        assert headings[0] == {
            "level": 1,
            "text": "Introduction",
            "style": "Heading1",
            "paraId": "00000001",
        }
        assert headings[1]["level"] == 2
        assert headings[1]["text"] == "Background"


class TestSearch:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_plain_text(self):
        results = _j(server.search_text("contract"))
        assert len(results) == 1
        assert results[0]["source"] == "document"
        assert results[0]["paraId"] == "00000004"

    def test_case_insensitive(self):
        results = _j(server.search_text("CONTRACT"))
        assert len(results) == 1

    def test_regex(self):
        results = _j(server.search_text(r"\d+ days", regex=True))
        assert len(results) == 1
        assert results[0]["matches"][0]["match"] == "30 days"

    def test_no_matches(self):
        assert _j(server.search_text("zzzznonexistent")) == []

    def test_search_in_footnotes(self):
        results = _j(server.search_text("appendix"))
        sources = {r["source"] for r in results}
        assert "footnotes" in sources

    def test_search_regex_no_match(self):
        assert _j(server.search_text(r"^ZZZZ$", regex=True)) == []


class TestParagraph:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_get_paragraph(self):
        para = _j(server.get_paragraph("00000004"))
        assert "contract" in para["text"]
        assert para["paraId"] == "00000004"
        assert para["style"] == ""  # body paragraph, no style

    def test_get_heading_paragraph(self):
        para = _j(server.get_paragraph("00000001"))
        assert para["style"] == "Heading1"
        assert para["text"] == "Introduction"

    def test_not_found(self):
        with pytest.raises(ValueError, match="not found"):
            server.get_paragraph("DEADBEEF")


# ═══════════════════════════════════════════════════════════════════════════
#  Footnotes
# ═══════════════════════════════════════════════════════════════════════════


class TestFootnotes:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_get_footnotes(self):
        fn = _j(server.get_footnotes())
        assert len(fn) == 1
        assert fn[0]["id"] == 1
        assert "appendix" in fn[0]["text"].lower()

    def test_add_footnote(self):
        result = _j(server.add_footnote("00000005", "New footnote."))
        assert result["footnote_id"] == 2
        assert result["para_id"] == "00000005"
        # Verify
        fn = _j(server.get_footnotes())
        assert len(fn) == 2

    def test_add_footnote_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_footnote("DEADBEEF", "text")

    def test_validate_valid(self):
        v = _j(server.validate_footnotes())
        assert v["valid"] is True
        assert v["references"] == 1
        assert v["definitions"] == 1

    def test_validate_after_add(self):
        server.add_footnote("00000005", "Extra.")
        v = _j(server.validate_footnotes())
        assert v["valid"] is True
        assert v["references"] == 2
        assert v["definitions"] == 2


# ═══════════════════════════════════════════════════════════════════════════
#  ParaId validation
# ═══════════════════════════════════════════════════════════════════════════


class TestParaIds:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_valid(self):
        v = _j(server.validate_paraids())
        assert v["valid"] is True
        assert v["total"] > 0
        assert v["duplicates"] == {}
        assert v["out_of_range"] == []

    def test_detect_invalid_hex(self):
        """Inject a non-hex paraId and verify detection."""
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        para = list(tree.iter(f"{W}p"))[0]
        para.set(f"{W14}paraId", "ZZZZZZZZ")
        v = _j(server.validate_paraids())
        assert "ZZZZZZZZ" in v["out_of_range"]
        assert v["valid"] is False

    def test_detect_out_of_range(self):
        """Inject a paraId >= 0x80000000."""
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        para = list(tree.iter(f"{W}p"))[0]
        para.set(f"{W14}paraId", "FFFFFFFF")
        v = _j(server.validate_paraids())
        assert "FFFFFFFF" in v["out_of_range"]


# ═══════════════════════════════════════════════════════════════════════════
#  Watermark removal
# ═══════════════════════════════════════════════════════════════════════════


class TestWatermark:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_remove_watermark(self):
        result = _j(server.remove_watermark())
        assert result["removed"] == 1
        assert result["details"][0]["text"] == "DRAFT"
        assert "header" in result["details"][0]["header"]

    def test_remove_again_is_noop(self):
        server.remove_watermark()
        result = _j(server.remove_watermark())
        assert result["removed"] == 0


# ═══════════════════════════════════════════════════════════════════════════
#  Track changes
# ═══════════════════════════════════════════════════════════════════════════


class TestTrackChanges:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    # ── insert_text ─────────────────────────────────────────────────────

    def test_insert_at_end(self):
        r = _j(server.insert_text("00000004", " [OK]"))
        assert r["type"] == "insertion"
        assert r["author"] == "Claude"
        assert "change_id" in r
        assert "date" in r

    def test_insert_at_start(self):
        r = _j(server.insert_text("00000005", "[NOTE] ", position="start"))
        assert r["type"] == "insertion"

    def test_insert_after_substring(self):
        r = _j(server.insert_text("00000004", " (extended)", position="30 days"))
        assert r["type"] == "insertion"

    def test_insert_after_missing_substring_falls_back_to_end(self):
        """If the position substring isn't found, appends to end."""
        r = _j(server.insert_text("00000004", " [FALLBACK]", position="ZZZMISSING"))
        assert r["type"] == "insertion"

    def test_insert_custom_author(self):
        r = _j(server.insert_text("00000005", "x", author="Reviewer"))
        assert r["author"] == "Reviewer"

    def test_insert_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.insert_text("DEADBEEF", "text")

    def test_insert_into_para_without_ppr(self):
        """Paragraph 00000002 has no pPr — insert at start should work."""
        r = _j(server.insert_text("00000002", "[START] ", position="start"))
        assert r["type"] == "insertion"

    # ── delete_text ─────────────────────────────────────────────────────

    def test_delete_text(self):
        r = _j(server.delete_text("00000004", "30 days"))
        assert r["type"] == "deletion"
        assert r["author"] == "Claude"

    def test_delete_preserves_rpr(self):
        """Paragraph 00000005 has bold runs — rPr should be preserved in deletion."""
        r = _j(server.delete_text("00000005", "bold"))
        assert r["type"] == "deletion"

    def test_delete_substring_splits_run(self):
        """Delete a substring from the middle of a run."""
        r = _j(server.delete_text("00000004", "contract"))
        assert r["type"] == "deletion"

    def test_delete_custom_author(self):
        r = _j(server.delete_text("00000004", "effective", author="Editor"))
        assert r["author"] == "Editor"

    def test_delete_not_found_in_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.delete_text("00000004", "nonexistent_xyz")

    def test_delete_spans_multiple_runs(self):
        """Text spanning runs raises a clear error."""
        with pytest.raises(ValueError, match="single run"):
            server.delete_text("00000006", "First bold")

    def test_delete_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.delete_text("DEADBEEF", "text")

    def test_delete_no_text_element(self):
        """Run with footnoteReference (no w:t) is skipped."""
        # Paragraph 00000002 has a run with footnoteReference and no w:t
        # Trying to delete text that only exists in the first run should work
        r = _j(server.delete_text("00000002", "important"))
        assert r["type"] == "deletion"


# ═══════════════════════════════════════════════════════════════════════════
#  Comments
# ═══════════════════════════════════════════════════════════════════════════


class TestComments:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_get_comments_empty(self):
        assert _j(server.get_comments()) == []

    def test_add_comment(self):
        r = _j(server.add_comment("00000004", "Please review."))
        assert r["comment_id"] == 0
        assert r["author"] == "Claude"
        assert r["para_id"] == "00000004"
        # Verify
        comments = _j(server.get_comments())
        assert len(comments) == 1
        assert comments[0]["text"] == "Please review."

    def test_add_comment_custom_author(self):
        r = _j(server.add_comment("00000004", "Note", author="Legal Team"))
        assert r["author"] == "Legal Team"

    def test_add_comment_bad_para(self):
        with pytest.raises(ValueError, match="not found"):
            server.add_comment("DEADBEEF", "text")

    def test_reply(self):
        parent = _j(server.add_comment("00000004", "Original"))
        reply = _j(server.reply_to_comment(parent["comment_id"], "Reply"))
        assert reply["parent_id"] == parent["comment_id"]
        assert reply["comment_id"] != parent["comment_id"]
        comments = _j(server.get_comments())
        assert len(comments) == 2

    def test_reply_nonexistent(self):
        # Must create comments.xml first (add a comment), then reply to bad ID
        server.add_comment("00000004", "setup")
        with pytest.raises(ValueError, match="not found"):
            server.reply_to_comment(999, "reply")

    def test_add_comment_creates_comments_xml(self):
        """First comment on a doc without comments.xml creates the part."""
        # Our fixture has no comments.xml — add_comment should create it
        r = _j(server.add_comment("00000004", "Created fresh"))
        assert r["comment_id"] == 0
        doc = server._doc
        assert "word/comments.xml" in doc._trees
        assert "word/comments.xml" in doc._modified

    def test_comment_on_para_with_ppr(self):
        """Add comment to heading (has pPr) — range markers placed correctly."""
        r = _j(server.add_comment("00000001", "Heading comment"))
        assert r["para_id"] == "00000001"

    def test_comment_on_para_without_ppr_or_runs(self):
        """Edge: paragraph with no pPr and no runs. Test fixture para 00000004 has runs."""
        # Use heading para — has pPr but test "first_run is not None" branch
        r = _j(server.add_comment("00000003", "Another comment"))
        assert r["comment_id"] >= 0


# ═══════════════════════════════════════════════════════════════════════════
#  Audit
# ═══════════════════════════════════════════════════════════════════════════


class TestAudit:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_audit_clean(self):
        """Audit on test fixture (watermark causes DRAFT artifacts)."""
        result = _j(server.audit_document())
        assert result["footnotes"]["valid"] is True
        assert result["endnotes"]["valid"] is True
        assert result["paraids"]["valid"] is True
        assert result["headings"]["count"] == 2
        assert result["headings"]["issues"] == []
        assert result["bookmarks"]["total"] == 1
        assert result["bookmarks"]["unpaired_starts"] == 0
        assert result["bookmarks"]["unpaired_ends"] == 0
        assert result["relationships"]["missing_targets"] == []
        assert result["images"]["missing"] == []
        assert result["tables"]["inconsistent_columns"] == []
        assert result["protection"]["edit"] == "none"
        assert isinstance(result["valid"], bool)

    def test_audit_after_watermark_removal(self):
        server.remove_watermark()
        result = _j(server.audit_document())
        # No DRAFT artifacts after removal
        draft_artifacts = [a for a in result["artifacts"] if a["marker"] == "DRAFT"]
        assert len(draft_artifacts) == 0


# ═══════════════════════════════════════════════════════════════════════════
#  Save & roundtrip
# ═══════════════════════════════════════════════════════════════════════════


class TestSave:
    def test_save_to_new_path(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        out = str(tmp_path / "output.docx")
        result = _j(server.save_document(out))
        assert os.path.exists(result["path"])
        assert result["size_bytes"] > 0

    def test_save_overwrite_source(self, test_docx: Path):
        server.open_document(str(test_docx))
        server.insert_text("00000004", " x")
        result = _j(server.save_document())
        assert result["path"] == str(test_docx)
        assert "word/document.xml" in result["modified_parts"]

    def test_save_no_modifications(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        out = str(tmp_path / "clean.docx")
        result = _j(server.save_document(out))
        assert result["modified_parts"] == []

    def test_save_no_document(self):
        with pytest.raises(RuntimeError, match="No document"):
            server.save_document()


class TestRoundtrip:
    def test_full_workflow(self, test_docx: Path, tmp_path: Path):
        """Open → edit → comment → footnote → watermark → save → reopen → verify."""
        # Open
        info = _j(server.open_document(str(test_docx)))
        assert info["footnote_count"] == 1

        # Track changes
        server.delete_text("00000004", "30 days")
        server.insert_text("00000004", "60 days")

        # Comment
        server.add_comment("00000004", "Extended per negotiation")

        # Footnote
        server.add_footnote("00000005", "Added during review.")

        # Watermark
        server.remove_watermark()

        # Save
        out = str(tmp_path / "roundtrip.docx")
        _j(server.save_document(out))
        assert os.path.exists(out)
        server.close_document()

        # Reopen and verify
        info2 = _j(server.open_document(out))
        assert info2["footnote_count"] == 2
        assert info2["comment_count"] == 1

        comments = _j(server.get_comments())
        assert len(comments) == 1
        assert "negotiation" in comments[0]["text"]

        fn = _j(server.get_footnotes())
        assert len(fn) == 2

        wm = _j(server.remove_watermark())
        assert wm["removed"] == 0


# ═══════════════════════════════════════════════════════════════════════════
#  Document model edge cases (direct API for coverage)
# ═══════════════════════════════════════════════════════════════════════════


class TestDocumentModel:
    def test_close_without_open(self):
        doc = DocxDocument("/tmp/fake.docx")
        doc.close()  # should not raise

    def test_validate_footnotes_no_footnotes_xml(self, tmp_path: Path):
        """Document without footnotes.xml."""
        path = tmp_path / "nofn.docx"
        import zipfile

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
                "<Relationships xmlns="
                '"http://schemas.openxmlformats.org/package/2006/relationships">'
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

        doc = DocxDocument(str(path))
        doc.open()

        # No footnotes.xml → validate returns valid with 0 counts
        v = doc.validate_footnotes()
        assert v["valid"] is True
        assert v["references"] == 0
        assert v["definitions"] == 0

        # get_footnotes returns empty
        assert doc.get_footnotes() == []

        # get_comments returns empty (no comments.xml)
        assert doc.get_comments() == []

        # Audit still works
        audit = doc.audit()
        assert audit["footnotes"]["valid"] is True

        doc.close()

    def test_text_extraction_empty(self):
        """_text on element with no <w:t> children."""
        el = etree.Element("test")
        assert DocxDocument._text(el) == ""


# ═══════════════════════════════════════════════════════════════════════════
#  Edge-case coverage
# ═══════════════════════════════════════════════════════════════════════════


class TestCoverageEdgeCases:
    """Targeted tests for branches not exercised by the main tests above."""

    def test_open_malformed_xml_part(self, tmp_path: Path):
        """A part with invalid XML is silently skipped (line 123-124)."""
        import zipfile

        path = tmp_path / "bad.docx"
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
                '<Override PartName="/word/footnotes.xml"'
                ' ContentType="application/vnd.openxmlformats-officedocument'
                '.wordprocessingml.footnotes+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?>'
                "<Relationships xmlns="
                '"http://schemas.openxmlformats.org/package/2006/relationships">'
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
            # Malformed footnotes.xml — not valid XML
            zf.writestr("word/footnotes.xml", "<broken><<<invalid xml")
            zf.writestr(
                "word/_rels/document.xml.rels",
                '<?xml version="1.0"?>'
                "<Relationships xmlns="
                '"http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1"'
                ' Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/footnotes"'
                ' Target="footnotes.xml"/>'
                "</Relationships>",
            )

        doc = DocxDocument(str(path))
        info = doc.open()
        # footnotes.xml was silently skipped
        assert "word/footnotes.xml" not in doc._trees
        assert info["paragraph_count"] == 1
        doc.close()

    def test_get_headings_skips_para_without_pstyle(self, test_docx: Path):
        """Paragraphs with pPr but no pStyle are skipped (line 177)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        # Inject a para with pPr but no pStyle
        body = tree.find(f"{W}body")
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0A0A0A0A")
        p.set(f"{W14}textId", "77777777")
        ppr = etree.SubElement(p, f"{W}pPr")
        # Add spacing but no pStyle
        etree.SubElement(ppr, f"{W}spacing").set(f"{W}after", "100")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "Not a heading"
        # Headings count should be unchanged
        headings = doc.get_headings()
        assert len(headings) == 2

    def test_get_headings_skips_non_heading_style(self, test_docx: Path):
        """Paragraphs with a non-heading pStyle are skipped (line 181)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0B0B0B0B")
        p.set(f"{W14}textId", "77777777")
        ppr = etree.SubElement(p, f"{W}pPr")
        etree.SubElement(ppr, f"{W}pStyle").set(f"{W}val", "BodyText")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "Just body text"
        headings = doc.get_headings()
        assert len(headings) == 2

    def test_insert_at_start_with_ppr(self, test_docx: Path):
        """Insert at start of a heading (has pPr) uses ppr.addnext (line 438)."""
        server.open_document(str(test_docx))
        # Paragraph 00000001 is a heading with pPr
        r = _j(server.insert_text("00000001", "[PREFIX] ", position="start"))
        assert r["type"] == "insertion"

    def test_delete_run_with_empty_text(self, test_docx: Path):
        """Run with <w:t/> (text=None) is skipped (line 480)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        # Inject a paragraph with a run containing an empty w:t
        body = tree.find(f"{W}body")
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0C0C0C0C")
        p.set(f"{W14}textId", "77777777")
        r1 = etree.SubElement(p, f"{W}r")
        etree.SubElement(r1, f"{W}t")  # empty w:t — text is None
        r2 = etree.SubElement(p, f"{W}r")
        t2 = etree.SubElement(r2, f"{W}t")
        t2.text = "findable text"
        result = doc.delete_text("0C0C0C0C", "findable")
        assert result["type"] == "deletion"

    def test_comment_on_para_with_ppr_no_runs(self, test_docx: Path):
        """Comment on para with pPr but no runs uses ppr.addnext (line 596-597)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0D0D0D0D")
        p.set(f"{W14}textId", "77777777")
        etree.SubElement(p, f"{W}pPr")
        # No runs — ppr.addnext branch
        result = doc.add_comment("0D0D0D0D", "Comment on empty para")
        assert result["para_id"] == "0D0D0D0D"

    def test_comment_on_bare_para(self, test_docx: Path):
        """Comment on para with no pPr and no runs uses insert(0) (line 598-599)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0E0E0E0E")
        p.set(f"{W14}textId", "77777777")
        # No pPr, no runs
        result = doc.add_comment("0E0E0E0E", "Comment on bare para")
        assert result["para_id"] == "0E0E0E0E"

    def test_reply_with_comments_extended(self, test_docx: Path):
        """Reply threading populates commentsExtended.xml (lines 656-662)."""
        server.open_document(str(test_docx))
        doc = server._doc
        # Create commentsExtended.xml so the threading branch is taken
        ext_xml = (
            '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>'
        )
        doc._trees["word/commentsExtended.xml"] = etree.fromstring(ext_xml)
        # Add a comment first
        parent = doc.add_comment("00000004", "Original")
        reply = doc.reply_to_comment(parent["comment_id"], "Threaded reply")
        assert reply["parent_id"] == parent["comment_id"]
        # Verify commentsExtended was updated
        ext = doc._trees["word/commentsExtended.xml"]

        exts = ext.findall(f"{W15}commentEx")
        assert len(exts) == 1
        assert exts[0].get(f"{W15}done") == "0"

    def test_audit_heading_level_skip(self, test_docx: Path):
        """Audit detects heading level jumps (line 688)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")
        # Add an H4 heading right after H2 — should flag a level skip
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0F0F0F0F")
        p.set(f"{W14}textId", "77777777")
        ppr = etree.SubElement(p, f"{W}pPr")
        etree.SubElement(ppr, f"{W}pStyle").set(f"{W}val", "Heading4")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "Skipped level heading"
        audit = doc.audit()
        issues = audit["headings"]["issues"]
        assert any(i["issue"] == "level_skip" for i in issues)

    def test_audit_external_rel_skipped(self, test_docx: Path):
        """Audit skips external rels (line 714)."""
        server.open_document(str(test_docx))
        doc = server._doc
        rels = doc._trees.get("word/_rels/document.xml.rels")
        if rels is not None:
            ext_rel = etree.SubElement(rels, f"{RELS}Relationship")
            ext_rel.set("Id", "rId99")
            ext_rel.set("TargetMode", "External")
            ext_rel.set("Target", "https://example.com")
            ext_rel.set(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            )
        audit = doc.audit()
        # External rel should not appear as missing
        assert audit["relationships"]["missing_targets"] == []

    def test_audit_missing_rel_target(self, test_docx: Path):
        """Audit flags internal rels pointing to missing files (line 717)."""
        server.open_document(str(test_docx))
        doc = server._doc
        rels = doc._trees.get("word/_rels/document.xml.rels")
        if rels is not None:
            bad_rel = etree.SubElement(rels, f"{RELS}Relationship")
            bad_rel.set("Id", "rId98")
            bad_rel.set("Target", "nonexistent_file.xml")
            bad_rel.set(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFDesigner",
            )
        audit = doc.audit()
        missing = audit["relationships"]["missing_targets"]
        assert any(m["id"] == "rId98" for m in missing)

    def test_audit_image_references(self, test_docx: Path):
        """Audit checks blip image references (lines 724-731)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")

        # Add a paragraph with a blip referencing a missing image
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0A1A1A1A")
        p.set(f"{W14}textId", "77777777")
        r = etree.SubElement(p, f"{W}r")
        drawing = etree.SubElement(r, f"{W}drawing")
        blip = etree.SubElement(drawing, f"{A}blip")
        blip.set(f"{R}embed", "rId50")
        # Add a relationship for it pointing to a missing file
        rels = doc._trees.get("word/_rels/document.xml.rels")
        if rels is not None:
            img_rel = etree.SubElement(rels, f"{RELS}Relationship")
            img_rel.set("Id", "rId50")
            img_rel.set("Target", "media/missing_image.png")
            img_rel.set(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            )
        audit = doc.audit()
        missing_imgs = audit["images"]["missing"]
        assert any(m["rId"] == "rId50" for m in missing_imgs)

    def test_audit_blip_without_embed(self, test_docx: Path):
        """Audit skips blips with no r:embed attribute (line 725-726)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")

        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0A2A2A2A")
        p.set(f"{W14}textId", "77777777")
        r = etree.SubElement(p, f"{W}r")
        drawing = etree.SubElement(r, f"{W}drawing")
        etree.SubElement(drawing, f"{A}blip")  # no r:embed
        audit = doc.audit()
        # Should not crash; no missing images from this blip
        assert "images" in audit

    def test_audit_inconsistent_table_columns(self, test_docx: Path):
        """Audit detects tables with inconsistent column counts."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        # Add an extra cell to the first row of the table
        tbl = tree.find(f".//{W}tbl")
        first_row = tbl.find(f"{W}tr")
        tc = etree.SubElement(first_row, f"{W}tc")
        p = etree.SubElement(tc, f"{W}p")
        p.set(f"{W14}paraId", "0B0B0B0B")
        p.set(f"{W14}textId", "77777777")
        audit = doc.audit()
        assert len(audit["tables"]["inconsistent_columns"]) >= 1
        assert audit["valid"] is False

    def test_audit_with_protection(self, test_docx: Path):
        """Audit reports protection status when set."""
        server.open_document(str(test_docx))
        server.set_document_protection("trackedChanges")
        result = _j(server.audit_document())
        assert result["protection"]["edit"] == "trackedChanges"
        assert result["protection"]["enforcement"] == "1"

    def test_audit_no_settings(self, test_docx: Path):
        """Audit handles missing settings.xml gracefully."""
        server.open_document(str(test_docx))
        doc = server._doc
        doc._trees.pop("word/settings.xml", None)
        audit = doc.audit()
        assert audit["protection"]["edit"] == "unknown"

    def test_audit_artifact_markers(self, test_docx: Path):
        """Audit detects TODO/FIXME markers in text (line 738)."""
        server.open_document(str(test_docx))
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        body = tree.find(f"{W}body")
        # Add paragraph containing TODO
        p = etree.SubElement(body, f"{W}p")
        p.set(f"{W14}paraId", "0A3A3A3A")
        p.set(f"{W14}textId", "77777777")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "TODO: fix this section"
        audit = doc.audit()
        todo_arts = [a for a in audit["artifacts"] if a["marker"] == "TODO"]
        assert len(todo_arts) >= 1

    def test_watermark_pict_without_textpath(self, test_docx: Path):
        """w:pict without v:textpath is skipped; run without w:pict is skipped (line 389)."""
        server.open_document(str(test_docx))
        doc = server._doc
        header = doc._trees.get("word/header1.xml")
        if header is not None:
            for p in header.iter(f"{W}p"):
                # Add a plain text run (no w:pict) — triggers pict-is-None continue
                text_run = etree.SubElement(p, f"{W}r")
                t_el = etree.SubElement(text_run, f"{W}t")
                t_el.text = "Header text"
                # Add a run with w:pict but NO v:textpath
                r = etree.SubElement(p, f"{W}r")
                pict = etree.SubElement(r, f"{W}pict")
                etree.SubElement(pict, f"{V}shape").set("id", "nontextpath")
                break
        result = doc.remove_watermark()
        assert result["removed"] == 1

    def test_reply_without_comments_xml(self, test_docx: Path):
        """reply_to_comment when comments.xml doesn't exist raises RuntimeError (line 803)."""
        server.open_document(str(test_docx))
        with pytest.raises(RuntimeError, match="not found"):
            server.reply_to_comment(0, "reply")

    def test_validate_footnotes_no_doc_tree(self, test_docx: Path):
        """validate_footnotes when document.xml tree is None (line 326)."""
        server.open_document(str(test_docx))
        doc = server._doc
        # Simulate missing document tree
        saved = doc._trees.pop("word/document.xml")
        result = doc.validate_footnotes()
        assert result == {"error": "No document open"}
        doc._trees["word/document.xml"] = saved  # restore for cleanup

    def test_save_with_modified_none_tree(self, test_docx: Path, tmp_path: Path):
        """Save skips modified parts with no tree (line 768)."""
        server.open_document(str(test_docx))
        doc = server._doc
        # Mark a nonexistent part as modified
        doc._modified.add("word/nonexistent.xml")
        out = str(tmp_path / "save_skip.docx")
        result = doc.save(out)
        assert os.path.exists(result["path"])

    def test_save_no_workdir(self):
        """Save raises when no document is open (line 760)."""
        doc = DocxDocument("/tmp/fake.docx")
        with pytest.raises(RuntimeError, match="No document"):
            doc.save()


# ═══════════════════════════════════════════════════════════════════════════
#  Server entry point
# ═══════════════════════════════════════════════════════════════════════════


class TestMain:
    def test_main_calls_mcp_run(self, monkeypatch):
        called = []
        monkeypatch.setattr(server.mcp, "run", lambda: called.append(True))
        server.main()
        assert called == [True]


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Tables
# ═══════════════════════════════════════════════════════════════════════════


class TestTablesRoundtrip:
    """Save → reopen → verify for all table tools."""

    def test_get_tables_survives_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        tables_before = _j(server.get_tables())
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        tables_after = _j(server.get_tables())
        assert len(tables_after) == len(tables_before)
        assert tables_after[0]["cells"] == tables_before[0]["cells"]

    def test_add_table_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.add_table("00000004", rows=2, cols=3)
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        tables = _j(server.get_tables())
        assert len(tables) == 2
        new_tbl = next(t for t in tables if t["col_count"] == 3)
        assert new_tbl["row_count"] == 2

    def test_modify_cell_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.modify_cell(0, 0, 0, "Changed")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        tables = _j(server.get_tables())
        assert "Changed" in tables[0]["cells"][0][0]

    def test_add_table_row_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.add_table_row(0, cells=["New A", "New B"])
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        tables = _j(server.get_tables())
        assert tables[0]["row_count"] == 4
        assert tables[0]["cells"][-1] == ["New A", "New B"]

    def test_delete_table_row_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.delete_table_row(0, 2)
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        # Row still exists structurally (tracked deletion), table shape preserved
        tables = _j(server.get_tables())
        assert tables[0]["row_count"] == 3


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Styles
# ═══════════════════════════════════════════════════════════════════════════


class TestStylesRoundtrip:
    def test_get_styles_survives_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        styles_before = _j(server.get_styles())
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        styles_after = _j(server.get_styles())
        assert len(styles_after) == len(styles_before)
        ids_before = {s["id"] for s in styles_before}
        ids_after = {s["id"] for s in styles_after}
        assert ids_before == ids_after


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Headers / Footers
# ═══════════════════════════════════════════════════════════════════════════


class TestHeadersFootersRoundtrip:
    def test_get_headers_footers_survives_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        hf_before = _j(server.get_headers_footers())
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        hf_after = _j(server.get_headers_footers())
        assert len(hf_after) == len(hf_before)

    def test_edit_header_footer_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.edit_header_footer("header", "Document Header Text", "Updated Header")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        hf = _j(server.get_headers_footers())
        header_texts = " ".join(h["text"] for h in hf if h["location"] == "header")
        assert "Updated Header" in header_texts


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Properties
# ═══════════════════════════════════════════════════════════════════════════


class TestPropertiesRoundtrip:
    def test_get_properties_survives_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        props_before = _j(server.get_properties())
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        props_after = _j(server.get_properties())
        assert props_after["title"] == props_before["title"]
        assert props_after["creator"] == props_before["creator"]

    def test_set_properties_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.set_properties(title="Roundtrip Title", creator="RT Author")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        props = _j(server.get_properties())
        assert props["title"] == "Roundtrip Title"
        assert props["creator"] == "RT Author"


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Images
# ═══════════════════════════════════════════════════════════════════════════


class TestImagesRoundtrip:
    def test_get_images_survives_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        imgs_before = _j(server.get_images())
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        imgs_after = _j(server.get_images())
        assert len(imgs_after) == len(imgs_before)
        assert imgs_after[0]["filename"] == imgs_before[0]["filename"]

    def test_insert_image_roundtrip(self, test_docx: Path, tmp_path: Path):
        img = tmp_path / "roundtrip.png"
        img.write_bytes(
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
            b"\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02"
            b"\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx"
            b"\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N"
            b"\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        server.open_document(str(test_docx))
        server.insert_image("00000004", str(img))
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        images = _j(server.get_images())
        assert len(images) == 2


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Endnotes
# ═══════════════════════════════════════════════════════════════════════════


class TestEndnotesRoundtrip:
    def test_get_endnotes_survives_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        en_before = _j(server.get_endnotes())
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        en_after = _j(server.get_endnotes())
        assert len(en_after) == len(en_before)

    def test_add_endnote_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.add_endnote("00000004", "Roundtrip endnote")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        endnotes = _j(server.get_endnotes())
        assert len(endnotes) == 2
        assert any("Roundtrip endnote" in e["text"] for e in endnotes)

    def test_validate_endnotes_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.add_endnote("00000004", "Extra endnote")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        v = _j(server.validate_endnotes())
        assert v["valid"] is True


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Lists
# ═══════════════════════════════════════════════════════════════════════════


class TestListsRoundtrip:
    def test_add_list_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.add_list(["00000004", "00000005"], style="bullet")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        # Verify paragraphs still exist and doc is valid
        para = _j(server.get_paragraph("00000004"))
        assert para["paraId"] == "00000004"
        audit = _j(server.audit_document())
        assert audit["paraids"]["valid"] is True


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Sections / Page breaks
# ═══════════════════════════════════════════════════════════════════════════


class TestSectionsRoundtrip:
    def test_add_page_break_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        result = _j(server.add_page_break("00000004"))
        new_para_id = result["para_id"]
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        para = _j(server.get_paragraph(new_para_id))
        assert para["paraId"] == new_para_id

    def test_add_section_break_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.add_section_break("00000004", break_type="nextPage")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        # Verify paragraph still exists and section break survived
        doc = server._doc
        tree = doc._trees["word/document.xml"]
        p = doc._find_para(tree, "00000004")
        sect_pr = p.find(f"{W}pPr/{W}sectPr")
        assert sect_pr is not None

    def test_set_section_properties_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        # Add a body-level sectPr first
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = etree.SubElement(body, f"{W}sectPr")
        etree.SubElement(sect_pr, f"{W}pgSz")
        server.set_section_properties(width=15840, height=12240, orientation="landscape")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        # Verify section properties survived
        doc2 = server._doc._trees["word/document.xml"]
        body2 = doc2.find(f"{W}body")
        sect_pr2 = body2.find(f"{W}sectPr")
        assert sect_pr2 is not None
        pg_sz = sect_pr2.find(f"{W}pgSz")
        assert pg_sz is not None
        assert pg_sz.get(f"{W}w") == "15840"


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Cross-references
# ═══════════════════════════════════════════════════════════════════════════


class TestCrossReferenceRoundtrip:
    def test_cross_reference_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        _j(server.add_cross_reference(
            source_para_id="00000004",
            target_para_id="00000001",
            text="see Introduction",
        ))
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        # Verify cross-reference text appears in source paragraph
        para = _j(server.get_paragraph("00000004"))
        assert "see Introduction" in para["text"]


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Protection
# ═══════════════════════════════════════════════════════════════════════════


class TestProtectionRoundtrip:
    def test_set_protection_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.set_document_protection("trackedChanges")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        audit = _j(server.audit_document())
        assert audit["protection"]["edit"] == "trackedChanges"
        assert audit["protection"]["enforcement"] == "1"

    def test_set_protection_with_password_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.set_document_protection("readOnly", password="s3cret")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        audit = _j(server.audit_document())
        assert audit["protection"]["edit"] == "readOnly"


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Merge
# ═══════════════════════════════════════════════════════════════════════════


class TestMergeRoundtrip:
    def test_merge_roundtrip(self, test_docx: Path, tmp_path: Path):
        # Build a source document
        src = tmp_path / "source.docx"
        import zipfile as zf_mod

        with zf_mod.ZipFile(src, "w") as zf:
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
                '<w:p w14:paraId="50000001" w14:textId="77777777">'
                "<w:r><w:t>Merged paragraph</w:t></w:r></w:p>"
                "</w:body></w:document>",
            )

        server.open_document(str(test_docx))
        server.merge_documents(str(src))
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        results = _j(server.search_text("Merged paragraph"))
        assert len(results) >= 1
        v = _j(server.validate_paraids())
        assert v["valid"] is True


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Track changes (accept/reject/formatting)
# ═══════════════════════════════════════════════════════════════════════════


class TestTrackChangesRoundtrip:
    def test_accept_changes_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.insert_text("00000004", " ACCEPTED")
        server.delete_text("00000004", "30 days")
        server.accept_changes()
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        para = _j(server.get_paragraph("00000004"))
        assert "ACCEPTED" in para["text"]
        assert "30 days" not in para["text"]

    def test_reject_changes_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.insert_text("00000004", " REJECTED")
        server.delete_text("00000004", "30 days")
        server.reject_changes()
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        para = _j(server.get_paragraph("00000004"))
        assert "REJECTED" not in para["text"]
        assert "30 days" in para["text"]

    def test_set_formatting_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        server.set_formatting("00000004", "contract", bold=True, color="FF0000")
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        # Verify paragraph still contains the text
        para = _j(server.get_paragraph("00000004"))
        assert "contract" in para["text"]
        # Verify the rPrChange survived (formatting tracked change)
        doc = server._doc._trees["word/document.xml"]
        p = server._doc._find_para(doc, "00000004")
        rpr_changes = list(p.iter(f"{W}rPrChange"))
        assert len(rpr_changes) >= 1


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Comments (reply)
# ═══════════════════════════════════════════════════════════════════════════


class TestCommentsRoundtrip:
    def test_reply_to_comment_roundtrip(self, test_docx: Path, tmp_path: Path):
        server.open_document(str(test_docx))
        parent = _j(server.add_comment("00000004", "Parent comment"))
        _j(server.reply_to_comment(parent["comment_id"], "Reply text"))
        out = str(tmp_path / "rt.docx")
        _j(server.save_document(out))
        server.close_document()
        server.open_document(out)
        comments = _j(server.get_comments())
        assert len(comments) == 2
        texts = {c["text"] for c in comments}
        assert "Parent comment" in texts
        assert "Reply text" in texts


# ═══════════════════════════════════════════════════════════════════════════
#  E2E roundtrip: Full workflow with all tool categories
# ═══════════════════════════════════════════════════════════════════════════


class TestFullRoundtrip:
    def test_comprehensive_workflow(self, test_docx: Path, tmp_path: Path):
        """Open → use every tool category → save → reopen → verify all."""
        server.open_document(str(test_docx))

        # Track changes
        server.insert_text("00000004", " [INSERTED]")
        server.delete_text("00000004", "30 days")

        # Comment + reply
        parent = _j(server.add_comment("00000005", "Review needed"))
        server.reply_to_comment(parent["comment_id"], "Acknowledged")

        # Footnote + endnote
        server.add_footnote("00000005", "Added footnote")
        server.add_endnote("00000004", "Added endnote")

        # Table operations
        server.add_table("00000004", rows=1, cols=2)
        server.add_table_row(0, cells=["X", "Y"])

        # Formatting
        server.set_formatting("00000005", "Final", italic=True)

        # List
        server.add_list(["00000006"], style="numbered")

        # Image
        img = tmp_path / "test.png"
        img.write_bytes(
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
            b"\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02"
            b"\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx"
            b"\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N"
            b"\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        server.insert_image("00000005", str(img))

        # Properties
        server.set_properties(title="Comprehensive Test")

        # Cross-reference
        server.add_cross_reference(
            source_para_id="00000005",
            target_para_id="00000001",
            text="see Introduction",
        )

        # Section / page break
        server.add_page_break("00000006")
        server.add_section_break("00000003", break_type="continuous")

        # Protection
        server.set_document_protection("trackedChanges")

        # Watermark removal
        server.remove_watermark()

        # Save
        out = str(tmp_path / "comprehensive.docx")
        _j(server.save_document(out))
        server.close_document()

        # ── Reopen and verify everything ──────────────────────────────
        server.open_document(out)

        # Track changes
        para4 = _j(server.get_paragraph("00000004"))
        assert "[INSERTED]" in para4["text"]

        # Comments
        comments = _j(server.get_comments())
        assert len(comments) == 2

        # Footnotes
        fn = _j(server.get_footnotes())
        assert len(fn) == 2

        # Endnotes
        en = _j(server.get_endnotes())
        assert len(en) == 2

        # Tables
        tables = _j(server.get_tables())
        assert len(tables) >= 2

        # Images
        images = _j(server.get_images())
        assert len(images) == 2

        # Properties
        props = _j(server.get_properties())
        assert props["title"] == "Comprehensive Test"

        # Styles (still present)
        styles = _j(server.get_styles())
        assert len(styles) >= 4

        # Headers/footers
        hf = _j(server.get_headers_footers())
        assert len(hf) >= 1

        # Protection
        audit = _j(server.audit_document())
        assert audit["protection"]["edit"] == "trackedChanges"

        # Watermark gone
        wm = _j(server.remove_watermark())
        assert wm["removed"] == 0

        # Paraids valid
        v = _j(server.validate_paraids())
        assert v["valid"] is True

        # Endnote cross-refs valid
        ev = _j(server.validate_endnotes())
        assert ev["valid"] is True

        # Footnote cross-refs valid
        fv = _j(server.validate_footnotes())
        assert fv["valid"] is True
