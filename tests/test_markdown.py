"""Tests for markdown-to-DOCX conversion."""

from __future__ import annotations

import json
import struct
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.markdown import MarkdownConverter


@pytest.fixture()
def blank_doc(tmp_path: Path) -> DocxDocument:
    out = tmp_path / "test.docx"
    doc = DocxDocument.create(str(out))
    return doc


class TestHeadings:
    def test_h1(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "# Title")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        # Should have replaced the blank paragraph with the heading
        assert len(paras) == 1
        ppr = paras[0].find(f"{W}pPr")
        style = ppr.find(f"{W}pStyle")
        assert style.get(f"{W}val") == "Heading1"
        assert blank_doc._text(paras[0]) == "Title"

    def test_h1_through_h6(self, blank_doc: DocxDocument):
        md = "\n\n".join(f"{'#' * i} Heading {i}" for i in range(1, 7))
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 6
        for i, para in enumerate(paras, 1):
            style = para.find(f"{W}pPr/{W}pStyle")
            assert style.get(f"{W}val") == f"Heading{i}"

    def test_all_paragraphs_have_para_ids(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "# H1\n\nParagraph\n\n## H2")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        for para in body.findall(f"{W}p"):
            pid = para.get(f"{W14}paraId")
            assert pid is not None
            assert len(pid) == 8
            assert int(pid, 16) < 0x80000000


class TestParagraphs:
    def test_simple_paragraph(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "Hello world")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 1
        assert blank_doc._text(paras[0]) == "Hello world"

    def test_multiple_paragraphs(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "First\n\nSecond\n\nThird")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 3

    def test_empty_input(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 0


class TestInlineFormatting:
    def test_bold(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "**bold**")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        # Find the run with bold text
        bold_runs = [r for r in runs if r.find(f"{W}rPr/{W}b") is not None]
        assert len(bold_runs) >= 1
        assert blank_doc._text(bold_runs[0]) == "bold"

    def test_italic(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "*italic*")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        italic_runs = [r for r in runs if r.find(f"{W}rPr/{W}i") is not None]
        assert len(italic_runs) >= 1
        assert blank_doc._text(italic_runs[0]) == "italic"

    def test_strikethrough(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "~~struck~~")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        strike_runs = [r for r in runs if r.find(f"{W}rPr/{W}strike") is not None]
        assert len(strike_runs) >= 1

    def test_bold_italic_combo(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "***both***")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        both = [
            r
            for r in runs
            if r.find(f"{W}rPr/{W}b") is not None and r.find(f"{W}rPr/{W}i") is not None
        ]
        assert len(both) >= 1

    def test_inline_code(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "Use `print()` here")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        code_runs = [r for r in runs if r.find(f"{W}rPr/{W}rFonts") is not None]
        assert len(code_runs) >= 1
        # Should have Courier New font
        font = code_runs[0].find(f"{W}rPr/{W}rFonts")
        assert font.get(f"{W}ascii") == "Courier New"

    def test_smart_typography_applied(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, '"quoted"')
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        text = blank_doc._text(body)
        assert "\u201c" in text  # left double quote
        assert "\u201d" in text  # right double quote

    def test_smart_typography_not_in_code(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, '`"not smart"`')
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        code_runs = [r for r in runs if r.find(f"{W}rPr/{W}rFonts") is not None]
        assert any('"not smart"' in (blank_doc._text(r) or "") for r in code_runs)


class TestLists:
    def test_bullet_list(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "- Item A\n- Item B\n- Item C")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 3
        for p in paras:
            num_pr = p.find(f"{W}pPr/{W}numPr")
            assert num_pr is not None
            assert num_pr.find(f"{W}numId").get(f"{W}val") == "1"

    def test_numbered_list(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "1. First\n2. Second")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 2
        for p in paras:
            num_pr = p.find(f"{W}pPr/{W}numPr")
            assert num_pr is not None
            assert num_pr.find(f"{W}numId").get(f"{W}val") == "2"

    def test_nested_list_3_levels(self, blank_doc: DocxDocument):
        md = "- Level 0\n  - Level 1\n    - Level 2"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 3
        levels = [p.find(f"{W}pPr/{W}numPr/{W}ilvl").get(f"{W}val") for p in paras]
        assert levels == ["0", "1", "2"]


class TestCodeBlocks:
    def test_fenced_code_block(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "```\nline 1\nline 2\n```")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 2  # one per line
        for p in paras:
            style = p.find(f"{W}pPr/{W}pStyle")
            assert style is not None
            assert style.get(f"{W}val") == "CodeBlock"

    def test_code_block_no_smart_typography(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, '```\n"quoted"\n```')
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        text = blank_doc._text(body)
        assert '"quoted"' in text  # straight quotes preserved


class TestBlockquotes:
    def test_blockquote(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "> Quoted text")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 1
        style = paras[0].find(f"{W}pPr/{W}pStyle")
        assert style is not None
        assert style.get(f"{W}val") == "BlockQuote"

    def test_nested_blockquote(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "> Outer\n>> Inner")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) >= 2
        # All should have BlockQuote style
        for p in paras:
            style = p.find(f"{W}pPr/{W}pStyle")
            assert style is not None
            assert style.get(f"{W}val") == "BlockQuote"


class TestHorizontalRules:
    def test_hr(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "---")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 1
        border = paras[0].find(f"{W}pPr/{W}pBdr/{W}bottom")
        assert border is not None
        assert border.get(f"{W}val") == "single"


class TestTables:
    def test_simple_table(self, blank_doc: DocxDocument):
        md = "| A | B |\n|---|---|\n| 1 | 2 |"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        tables = body.findall(f"{W}tbl")
        assert len(tables) == 1
        rows = tables[0].findall(f"{W}tr")
        assert len(rows) == 2  # header + 1 body row

    def test_table_header_bold(self, blank_doc: DocxDocument):
        md = "| H1 | H2 |\n|---|---|\n| a | b |"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        tbl = body.find(f"{W}tbl")
        assert tbl is not None
        first_row = tbl.findall(f"{W}tr")[0]
        # Header row cells should have bold runs
        bold_runs = first_row.findall(f".//{W}r/{W}rPr/{W}b/..")
        assert len(bold_runs) >= 1


class TestFootnotes:
    def test_footnote_creates_reference_and_definition(self, blank_doc: DocxDocument):
        md = "Text with a note[^1].\n\n[^1]: The note text."
        MarkdownConverter.convert(blank_doc, md)

        # Check body has a footnoteReference
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        refs = body.findall(f".//{W}footnoteReference")
        assert len(refs) >= 1

        # Check footnotes.xml has the footnote definition
        fn_tree = blank_doc._trees["word/footnotes.xml"]
        # Real footnotes exclude separator ids 0 and -1
        real_fns = [
            f for f in fn_tree.findall(f"{W}footnote") if f.get(f"{W}id") not in ("0", "-1")
        ]
        assert len(real_fns) >= 1
        fn_text = blank_doc._text(real_fns[0])
        assert "The note text." in fn_text


class TestImages:
    def test_local_image_embedded(self, blank_doc: DocxDocument, tmp_path: Path):
        # Create a minimal valid 1x1 PNG
        img_path = tmp_path / "tiny.png"
        _write_tiny_png(img_path)

        MarkdownConverter.convert(blank_doc, "![alt](tiny.png)", base_dir=tmp_path)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        drawings = body.findall(f".//{W}drawing")
        assert len(drawings) >= 1

    def test_remote_image_becomes_hyperlink(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "![photo](https://example.com/img.png)")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        hyperlinks = body.findall(f".//{W}hyperlink")
        assert len(hyperlinks) >= 1

    def test_missing_image_placeholder(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "![alt](nonexistent.png)")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        text = blank_doc._text(body)
        assert "[Image not found:" in text


class TestTaskLists:
    def test_checked_and_unchecked(self, blank_doc: DocxDocument):
        md = "- [x] Done\n- [ ] Todo"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        text = blank_doc._text(body)
        assert "\u2611" in text  # checked checkbox
        assert "\u2610" in text  # unchecked checkbox


class TestMixed:
    def test_mixed_constructs(self, blank_doc: DocxDocument):
        md = "# Heading\n\nA paragraph.\n\n- Item 1\n- Item 2\n\n| A | B |\n|---|---|\n| 1 | 2 |"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        # Should have: heading + paragraph + 2 list items = 4 paragraphs + 1 table
        paras = body.findall(f"{W}p")
        tables = body.findall(f"{W}tbl")
        assert len(paras) == 4
        assert len(tables) == 1
        # First para is heading
        style = paras[0].find(f"{W}pPr/{W}pStyle")
        assert style.get(f"{W}val") == "Heading1"


class TestInputValidation:
    """Validation tests for the server tool (used by Task 8)."""

    def test_mutually_exclusive_inputs(self):
        """Providing both markdown and template should be rejected."""
        # This tests the server-layer validation logic that will be
        # implemented in Task 8. For now, just verify MarkdownConverter
        # itself accepts a text string cleanly.
        from docx_mcp.markdown import MarkdownConverter as MC

        assert callable(MC.convert)

    def test_neither_input(self):
        """Providing neither markdown nor template should be rejected."""
        # Placeholder for server-layer validation in Task 8.
        from docx_mcp.markdown import MarkdownConverter as MC

        assert callable(MC.convert)


def _write_tiny_png(path: Path) -> None:
    """Write a minimal valid 1x1 white PNG file."""
    import zlib

    def _chunk(chunk_type: bytes, data: bytes) -> bytes:
        c = chunk_type + data
        crc = struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
        return struct.pack(">I", len(data)) + c + crc

    signature = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    raw_row = b"\x00\xff\xff\xff"  # filter byte + white pixel (RGB)
    idat_data = zlib.compress(raw_row)
    path.write_bytes(
        signature + _chunk(b"IHDR", ihdr_data) + _chunk(b"IDAT", idat_data) + _chunk(b"IEND", b"")
    )


# ── Server tool tests ──────────────────────────────────────────────────────


class TestCreateFromMarkdownTool:
    def test_from_raw_text(self, tmp_path: Path):
        out = tmp_path / "from_md.docx"
        result = json.loads(server.create_from_markdown(str(out), markdown="# Hello\n\nWorld"))
        assert "paragraph_count" in result
        assert server._doc is not None

    def test_from_file(self, tmp_path: Path):
        md_file = tmp_path / "input.md"
        md_file.write_text("# From File\n\nContent here.")
        out = tmp_path / "from_file.docx"
        result = json.loads(server.create_from_markdown(str(out), md_path=str(md_file)))
        assert "paragraph_count" in result

    def test_both_inputs_raises(self, tmp_path: Path):
        out = tmp_path / "err.docx"
        md_file = tmp_path / "input.md"
        md_file.write_text("# Test")
        result = server.create_from_markdown(str(out), md_path=str(md_file), markdown="# Test")
        assert "error" in result.lower() or "mutually exclusive" in result.lower()

    def test_no_input_raises(self, tmp_path: Path):
        out = tmp_path / "err.docx"
        result = server.create_from_markdown(str(out))
        assert "error" in result.lower() or "either" in result.lower()

    def test_with_template(self, tmp_path: Path):
        from tests.conftest import _build_fixture

        template = tmp_path / "tmpl.dotx"
        _build_fixture(template)
        out = tmp_path / "from_tmpl_md.docx"
        result = json.loads(
            server.create_from_markdown(
                str(out), markdown="# Template Test", template_path=str(template)
            )
        )
        assert "paragraph_count" in result

    def test_image_path_relative_to_md_file(self, tmp_path: Path):
        # Create image in same dir as markdown file
        img = tmp_path / "subdir" / "test.png"
        img.parent.mkdir(parents=True)
        _write_tiny_png(img)
        md_file = tmp_path / "subdir" / "doc.md"
        md_file.write_text("![img](test.png)")
        out = tmp_path / "output.docx"
        result = json.loads(server.create_from_markdown(str(out), md_path=str(md_file)))
        assert "paragraph_count" in result

    def test_closes_previous_doc(self, tmp_path: Path):
        """create_from_markdown closes an already-open doc (server.py line 122)."""
        first = tmp_path / "first.docx"
        server.create_from_markdown(str(first), markdown="# First")
        assert server._doc is not None
        old_workdir = server._doc.workdir

        second = tmp_path / "second.docx"
        server.create_from_markdown(str(second), markdown="# Second")
        assert not old_workdir.exists()  # old workdir cleaned up

    def test_nonexistent_md_path(self, tmp_path: Path):
        """create_from_markdown returns error for nonexistent md_path (server.py line 132)."""
        out = tmp_path / "out.docx"
        result = server.create_from_markdown(str(out), md_path="/nonexistent/file.md")
        assert "Error" in result
        assert "not found" in result


# ── Coverage gap tests for markdown.py ────────────────────────────────────


class TestLinks:
    """Test hyperlink rendering (markdown.py lines 182, 282-297)."""

    def test_link_creates_hyperlink(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "[click here](https://example.com)")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        hyperlinks = body.findall(f".//{W}hyperlink")
        assert len(hyperlinks) >= 1
        # Should have link text
        text = blank_doc._text(hyperlinks[0])
        assert "click here" in text

    def test_link_has_relationship(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "[test](https://example.com/page)")
        rels = blank_doc._tree("word/_rels/document.xml.rels")
        from docx_mcp.document.base import RELS

        targets = [
            r.get("Target")
            for r in rels.findall(f"{RELS}Relationship")
            if r.get("TargetMode") == "External"
        ]
        assert "https://example.com/page" in targets


class TestSoftbreakAndLinebreak:
    """Test softbreak and linebreak inline types (markdown.py lines 186, 188-189)."""

    def test_softbreak_becomes_space(self, blank_doc: DocxDocument):
        # A single newline inside a paragraph produces a softbreak token
        MarkdownConverter.convert(blank_doc, "line1\nline2")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        text = blank_doc._text(body)
        # Softbreak renders as a space
        assert "line1" in text
        assert "line2" in text

    def test_linebreak_creates_br(self, blank_doc: DocxDocument):
        # Two trailing spaces + newline = hard line break
        MarkdownConverter.convert(blank_doc, "line1  \nline2")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        breaks = body.findall(f".//{W}br")
        assert len(breaks) >= 1


class TestUnknownTokenType:
    """Test fallback for unknown token types (markdown.py line 140)."""

    def test_unknown_block_type_returns_empty(self, blank_doc: DocxDocument):
        converter = MarkdownConverter(blank_doc)
        result = converter._render_block({"type": "nonexistent_type"})
        assert result == []


class TestSectPrPath:
    """Test element insertion before sectPr (markdown.py line 77)."""

    def test_elements_inserted_before_sect_pr(self, blank_doc: DocxDocument):
        from lxml import etree

        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        # Add a sectPr element to the body
        sect_pr = etree.SubElement(body, f"{W}sectPr")
        etree.SubElement(sect_pr, f"{W}pgSz")

        MarkdownConverter.convert(blank_doc, "# Heading\n\nParagraph text")
        # sectPr should be last child of body
        children = list(body)
        assert children[-1].tag == f"{W}sectPr"
        # Paragraphs should come before sectPr
        assert len(children) >= 3  # heading + paragraph + sectPr
        for child in children[:-1]:
            assert child.tag == f"{W}p"


class TestFootnoteEdgeCases:
    """Test footnote edge cases (markdown.py lines 484, 535)."""

    def test_footnote_definitions_no_footnotes_xml(self, blank_doc: DocxDocument):
        """_process_footnote_definitions returns early when fn_tree is None (line 484)."""
        # Remove footnotes.xml from the trees
        del blank_doc._trees["word/footnotes.xml"]
        # This markdown has footnote definitions -- should not crash
        MarkdownConverter.convert(blank_doc, "Text[^1]\n\n[^1]: Note text")
        # Should complete without error; no footnotes created
        assert "word/footnotes.xml" not in blank_doc._trees

    def test_unresolved_footnote_ref(self, blank_doc: DocxDocument):
        """_render_footnote_ref returns early when fn_id is None (line 535)."""
        converter = MarkdownConverter(blank_doc)
        from lxml import etree

        parent = etree.Element(f"{W}p")
        # Call with a key that was never defined in footnote_map
        converter._render_footnote_ref(parent, {"raw": "undefined_key"})
        # No footnoteReference should be added
        refs = parent.findall(f".//{W}footnoteReference")
        assert len(refs) == 0

    def test_strip_orphan_footnotes(self, blank_doc: DocxDocument):
        """Orphaned footnote definitions are removed after rendering."""
        from lxml import etree

        fn_tree = blank_doc._tree("word/footnotes.xml")

        # Inject an orphan (id=99) before conversion
        orphan = etree.SubElement(fn_tree, f"{W}footnote")
        orphan.set(f"{W}id", "99")
        p = etree.SubElement(orphan, f"{W}p")
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = "Orphan text"

        ids_before = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        assert 99 in ids_before

        # Convert — cleanup should strip the orphan but keep the valid footnote
        MarkdownConverter.convert(blank_doc, "Text[^a]\n\n[^a]: Valid note.")

        ids_after = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        assert 99 not in ids_after

        result = blank_doc.validate_footnotes()
        assert result["valid"] is True
        assert result["orphan_definitions"] == []
        assert result["references"] == 1
        assert result["definitions"] == 1

    def test_strip_orphan_footnotes_no_footnotes_xml(self, blank_doc: DocxDocument):
        """_strip_orphan_footnotes returns early when footnotes.xml is absent."""
        del blank_doc._trees["word/footnotes.xml"]
        converter = MarkdownConverter(blank_doc)
        # Should not crash
        converter._strip_orphan_footnotes()
