"""Tests for copy_document, flatten_document, get_reading_time — RED phase."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

from docx_mcp import server

# ── copy_document ────────────────────────────────────────────────────────────


def test_copy_document_creates_file(test_docx, tmp_path):
    """copy_document saves a copy at the given path."""
    server._doc = None
    server.open_document(str(test_docx))
    out = str(tmp_path / "copy.docx")
    result = server._doc.copy_document(out)
    assert Path(out).exists()
    assert result["copied_to"] == out


def test_copy_document_is_valid_docx(test_docx, tmp_path):
    """The copied file is a valid ZIP containing word/document.xml."""
    server._doc = None
    server.open_document(str(test_docx))
    out = str(tmp_path / "copy2.docx")
    server._doc.copy_document(out)
    with zipfile.ZipFile(out, "r") as zf:
        assert "word/document.xml" in zf.namelist()


def test_copy_document_with_explicit_handle(test_docx, tmp_path):
    """copy_document reads from the document under the given handle."""
    server._doc = None
    h1 = "copy-handle-1"
    h2 = "copy-handle-2"
    try:
        server.open_document(str(test_docx), document_handle=h1)
        server.open_document(str(test_docx), document_handle=h2)
        out = str(tmp_path / "copy_h1.docx")
        result = json.loads(server.copy_document(out, document_handle=h1))
        assert result["copied_to"] == out
        assert Path(out).exists()
        with zipfile.ZipFile(out, "r") as zf:
            assert "word/document.xml" in zf.namelist()
    finally:
        server.close_document(document_handle=h1)
        server.close_document(document_handle=h2)


# ── flatten_document ─────────────────────────────────────────────────────────


def _docx_with_tracked_changes(tmp_path: Path) -> Path:
    """Build a minimal DOCX that has w:ins and w:del elements."""
    doc_xml = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001">
      <w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>inserted</w:t></w:r>
      </w:ins>
    </w:p>
    <w:p w14:paraId="00000002">
      <w:del w:id="2" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:delText>deleted</w:delText></w:r>
      </w:del>
    </w:p>
  </w:body>
</w:document>"""

    content_types = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    top_rels = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    path = tmp_path / "tracked.docx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", top_rels)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
    return path


def _docx_with_rpr_change(tmp_path: Path) -> Path:
    """Build a minimal DOCX that has w:rPrChange and w:pPrChange elements."""
    doc_xml = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000010">
      <w:pPr>
        <w:pPrChange w:id="3" w:author="Bob" w:date="2026-01-01T00:00:00Z">
          <w:pPr/>
        </w:pPrChange>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="4" w:author="Bob" w:date="2026-01-01T00:00:00Z">
            <w:rPr/>
          </w:rPrChange>
        </w:rPr>
        <w:t>hello</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""

    content_types = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    top_rels = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    path = tmp_path / "rprchange.docx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", top_rels)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
    return path


def test_flatten_document_removes_ins_del(tmp_path):
    """After flatten_document, no w:ins or w:del elements remain."""

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    path = _docx_with_tracked_changes(tmp_path)
    server._doc = None
    server.open_document(str(path))
    result = server._doc.flatten_document()
    assert result["changes_accepted"] >= 2
    doc = server._doc._tree("word/document.xml")
    assert list(doc.iter(f"{W}ins")) == []
    assert list(doc.iter(f"{W}del")) == []


def test_flatten_document_removes_rPrChange(tmp_path):
    """After flatten_document, no w:rPrChange or w:pPrChange elements remain."""

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    path = _docx_with_rpr_change(tmp_path)
    server._doc = None
    server.open_document(str(path))
    result = server._doc.flatten_document()
    assert result["formatting_changes_removed"] >= 2
    doc = server._doc._tree("word/document.xml")
    assert list(doc.iter(f"{W}rPrChange")) == []
    assert list(doc.iter(f"{W}pPrChange")) == []


# ── get_reading_time ─────────────────────────────────────────────────────────


def test_get_reading_time_basic(test_docx):
    """At 200 wpm, minutes = word_count / 200."""
    server._doc = None
    server.open_document(str(test_docx))
    wc = server._doc.get_word_count()
    result = server._doc.get_reading_time()
    assert result["words_per_minute"] == 200
    assert result["word_count"] == wc
    expected_minutes = round(wc / 200, 1)
    assert result["minutes"] == expected_minutes
    expected_seconds = round(wc / 200 * 60)
    assert result["seconds"] == expected_seconds


def test_get_reading_time_custom_wpm(test_docx):
    """Custom words_per_minute is used in calculation."""
    server._doc = None
    server.open_document(str(test_docx))
    wc = server._doc.get_word_count()
    result = server._doc.get_reading_time(words_per_minute=100)
    assert result["words_per_minute"] == 100
    assert result["minutes"] == round(wc / 100, 1)
    assert result["seconds"] == round(wc / 100 * 60)
