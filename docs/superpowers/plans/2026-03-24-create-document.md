# Create Document & Markdown-to-DOCX Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `create_document` and `create_from_markdown` tools so users can create new .docx files from scratch or from markdown content.

**Architecture:** New `CreationMixin` (classmethod factory that writes skeleton ZIP then opens it) + standalone `MarkdownConverter` (mistune parser → direct OOXML generation via lxml). Both build XML directly using the existing namespace constants and `_new_para_id()` helper.

**Tech Stack:** Python, lxml, mistune>=3.0, FastMCP

**Spec:** `docs/superpowers/specs/2026-03-23-create-document-design.md`

---

## Task 1: Add mistune dependency

**Files:**
- Modify: `pyproject.toml`

- [ ] **Step 1: Add mistune to dependencies**

In `pyproject.toml`, add `mistune>=3.0` to the `dependencies` list:

```toml
dependencies = [
    "mcp>=1.0.0",
    "lxml>=4.9.0",
    "mistune>=3.0",
]
```

Also bump version to `"0.3.0"`.

- [ ] **Step 2: Install and verify**

Run: `pip install -e ".[dev]"`
Expected: installs successfully including mistune

- [ ] **Step 3: Commit**

```bash
git add pyproject.toml
git commit -m "chore: add mistune dependency, bump version to 0.3.0"
```

---

## Task 2: CreationMixin — blank skeleton

**Files:**
- Create: `docx_mcp/document/creation.py`
- Create: `tests/test_creation.py`
- Modify: `docx_mcp/document/__init__.py`

### Context for implementer

The existing codebase builds DOCX files as ZIP archives of XML files. See `tests/conftest.py:_build_fixture()` for the exact pattern — it writes XML strings into a ZIP. The `BaseMixin.__init__` takes a file path and `open()` unpacks + parses it. Our `create()` classmethod writes the ZIP first, then instantiates and opens.

Key namespace constants (from `docx_mcp/document/base.py`):
```python
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    ...
}
```

The `_new_para_id()` method generates unique 8-hex-digit IDs < 0x80000000. The `_trees` dict caches parsed XML by relative path (e.g., `"word/document.xml"`). Call `_mark(rel_path)` to flag a tree as dirty for saving.

- [ ] **Step 1: Write failing tests for blank skeleton**

Create `tests/test_creation.py`:

```python
"""Tests for document creation (blank skeleton and template mode)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document import DocxDocument, W, W14


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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_creation.py -v`
Expected: FAIL — `DocxDocument` has no `create` classmethod

- [ ] **Step 3: Write CreationMixin**

Create `docx_mcp/document/creation.py`:

```python
"""Creation mixin: create blank DOCX documents and from templates."""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .base import BaseMixin, _now_iso


class CreationMixin:
    """Document creation operations."""

    @classmethod
    def create(cls, output_path: str, template_path: str | None = None) -> "CreationMixin":
        """Create a new DOCX and return an opened instance.

        Args:
            output_path: Path for the new .docx file.
            template_path: Optional .dotx template to copy from.
        """
        out = Path(output_path)

        if template_path:
            src = Path(template_path)
            if not src.exists():
                raise FileNotFoundError(f"Template not found: {src}")
            shutil.copy2(str(src), str(out))
        else:
            _write_blank_skeleton(out)

        instance = cls(str(out))
        instance.open()

        if template_path:
            _ensure_custom_styles(instance)
            _ensure_numbering(instance)

        return instance

    def get_info(self) -> dict:
        """Override point — BaseMixin.get_info will provide real impl."""
        ...


def _write_blank_skeleton(path: Path) -> None:
    """Write a minimal valid .docx ZIP archive."""
    import random

    def _pid() -> str:
        return f"{random.randint(1, 0x7FFFFFFF):08X}"

    # Generate unique paraIds for all paragraphs in the skeleton
    pids = {f"body1": _pid(), "fn_sep": _pid(), "fn_cont": _pid(),
            "en_sep": _pid(), "en_cont": _pid(), "hdr1": _pid()}

    files = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _TOP_RELS,
        "word/document.xml": _DOCUMENT_XML.format(**pids),
        "word/_rels/document.xml.rels": _DOC_RELS,
        "word/styles.xml": _STYLES_XML,
        "word/settings.xml": _SETTINGS_XML,
        "word/numbering.xml": _NUMBERING_XML,
        "word/footnotes.xml": _FOOTNOTES_XML.format(**pids),
        "word/endnotes.xml": _ENDNOTES_XML.format(**pids),
        "word/header1.xml": _HEADER_XML.format(**pids),
        "docProps/core.xml": _CORE_XML.format(now=_now_iso()),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content.strip())


def _ensure_custom_styles(doc: "CreationMixin") -> None:
    """Add CodeBlock and BlockQuote styles if missing in a template."""
    from lxml import etree
    from .base import W

    styles = doc._tree("word/styles.xml")
    if styles is None:
        return
    existing = {s.get(f"{W}styleId") for s in styles.findall(f"{W}style")}

    if "CodeBlock" not in existing:
        style = etree.SubElement(styles, f"{W}style")
        style.set(f"{W}type", "paragraph")
        style.set(f"{W}styleId", "CodeBlock")
        name = etree.SubElement(style, f"{W}name")
        name.set(f"{W}val", "Code Block")
        based = etree.SubElement(style, f"{W}basedOn")
        based.set(f"{W}val", "Normal")
        ppr = etree.SubElement(style, f"{W}pPr")
        shd = etree.SubElement(ppr, f"{W}shd")
        shd.set(f"{W}val", "clear")
        shd.set(f"{W}fill", "F2F2F2")
        spacing = etree.SubElement(ppr, f"{W}spacing")
        spacing.set(f"{W}before", "0")
        spacing.set(f"{W}after", "0")
        rpr = etree.SubElement(style, f"{W}rPr")
        font = etree.SubElement(rpr, f"{W}rFonts")
        font.set(f"{W}ascii", "Courier New")
        font.set(f"{W}hAnsi", "Courier New")
        sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", "18")  # 9pt = 18 half-points
        doc._mark("word/styles.xml")

    if "BlockQuote" not in existing:
        style = etree.SubElement(styles, f"{W}style")
        style.set(f"{W}type", "paragraph")
        style.set(f"{W}styleId", "BlockQuote")
        name = etree.SubElement(style, f"{W}name")
        name.set(f"{W}val", "Block Quote")
        based = etree.SubElement(style, f"{W}basedOn")
        based.set(f"{W}val", "Normal")
        ppr = etree.SubElement(style, f"{W}pPr")
        ind = etree.SubElement(ppr, f"{W}ind")
        ind.set(f"{W}left", "720")  # 0.5 inch = 720 twips
        pbdr = etree.SubElement(ppr, f"{W}pBdr")
        left_bdr = etree.SubElement(pbdr, f"{W}left")
        left_bdr.set(f"{W}val", "single")
        left_bdr.set(f"{W}sz", "24")  # 3pt
        left_bdr.set(f"{W}space", "4")
        left_bdr.set(f"{W}color", "AAAAAA")
        rpr = etree.SubElement(style, f"{W}rPr")
        etree.SubElement(rpr, f"{W}i")
        color = etree.SubElement(rpr, f"{W}color")
        color.set(f"{W}val", "555555")
        doc._mark("word/styles.xml")


def _ensure_numbering(doc: "CreationMixin") -> None:
    """Bootstrap numbering.xml if missing in a template."""
    from lxml import etree
    from .base import CT, NSMAP, RELS, W

    if doc._tree("word/numbering.xml") is not None:
        return

    # Parse the default numbering XML
    parser = etree.XMLParser(remove_blank_text=False)
    num_tree = etree.fromstring(_NUMBERING_XML.strip().encode(), parser)
    doc._trees["word/numbering.xml"] = num_tree
    doc._mark("word/numbering.xml")

    # Write file to workdir
    fp = doc.workdir / "word" / "numbering.xml"
    fp.parent.mkdir(parents=True, exist_ok=True)
    etree.ElementTree(num_tree).write(str(fp), xml_declaration=True, encoding="UTF-8")

    # Add content type override
    ct = doc._tree("[Content_Types].xml")
    if ct is not None:
        existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
        if "/word/numbering.xml" not in existing:
            ov = etree.SubElement(ct, f"{CT}Override")
            ov.set("PartName", "/word/numbering.xml")
            ov.set(
                "ContentType",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            )
            doc._mark("[Content_Types].xml")

    # Add relationship
    rels = doc._tree("word/_rels/document.xml.rels")
    if rels is not None:
        existing_targets = {r.get("Target") for r in rels.findall(f"{RELS}Relationship")}
        if "numbering.xml" not in existing_targets:
            import contextlib

            max_rid = 0
            for r in rels.findall(f"{RELS}Relationship"):
                rid = r.get("Id", "")
                if rid.startswith("rId"):
                    with contextlib.suppress(ValueError):
                        max_rid = max(max_rid, int(rid[3:]))
            rel = etree.SubElement(rels, f"{RELS}Relationship")
            rel.set("Id", f"rId{max_rid + 1}")
            rel.set(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
            )
            rel.set("Target", "numbering.xml")
            doc._mark("word/_rels/document.xml.rels")


# ── XML Templates ──────────────────────────────────────────────────────────
# These mirror the pattern in tests/conftest.py but include numbering.xml
# and all required Content-Type overrides.

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/footnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/header1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

_TOP_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    Target="docProps/core.xml"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    Target="footnotes.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    Target="header1.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    Target="endnotes.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  <Relationship Id="rId6"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Target="numbering.xml"/>
</Relationships>"""

_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="{body1}" w14:textId="77777777"/>
  </w:body>
</w:document>"""

_STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="3"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading5">
    <w:name w:val="heading 5"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="4"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading6">
    <w:name w:val="heading 6"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="5"/></w:pPr>
    <w:rPr><w:b/><w:i/><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListNumber">
    <w:name w:val="List Number"/><w:basedOn w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="EndnoteReference">
    <w:name w:val="endnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:sz w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText">
    <w:name w:val="endnote text"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:sz w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="Code Block"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:shd w:val="clear" w:fill="F2F2F2"/>
      <w:spacing w:before="0" w:after="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
      <w:sz w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BlockQuote">
    <w:name w:val="Block Quote"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="720"/>
      <w:pBdr>
        <w:left w:val="single" w:sz="24" w:space="4" w:color="AAAAAA"/>
      </w:pBdr>
    </w:pPr>
    <w:rPr><w:i/><w:color w:val="555555"/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
  </w:style>
</w:styles>"""

_SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:settings>"""

_NUMBERING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25E6"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25AA"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="3"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="4"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25E6"/><w:pPr><w:ind w:left="3600" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="5"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25AA"/><w:pPr><w:ind w:left="4320" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="6"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="5040" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="7"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25E6"/><w:pPr><w:ind w:left="5760" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="8"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25AA"/><w:pPr><w:ind w:left="6480" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="3"><w:numFmt w:val="decimal"/><w:lvlText w:val="%4."/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="4"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%5."/><w:pPr><w:ind w:left="3600" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="5"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%6."/><w:pPr><w:ind w:left="4320" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="6"><w:numFmt w:val="decimal"/><w:lvlText w:val="%7."/><w:pPr><w:ind w:left="5040" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="7"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%8."/><w:pPr><w:ind w:left="5760" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="8"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%9."/><w:pPr><w:ind w:left="6480" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>"""

_FOOTNOTES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:footnote w:type="separator" w:id="-1">
    <w:p w14:paraId="{fn_sep}" w14:textId="77777777"><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p w14:paraId="{fn_cont}" w14:textId="77777777"><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
</w:footnotes>"""

_ENDNOTES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:endnote w:type="separator" w:id="-1">
    <w:p w14:paraId="{en_sep}" w14:textId="77777777"><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p w14:paraId="{en_cont}" w14:textId="77777777"><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
</w:endnotes>"""

_HEADER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="{hdr1}" w14:textId="77777777"/>
</w:hdr>"""

_CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>docx-mcp</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>"""
```

- [ ] **Step 4: Add CreationMixin to DocxDocument**

In `docx_mcp/document/__init__.py`, add the import and mixin:

```python
from .creation import CreationMixin
```

Add `CreationMixin` to the `DocxDocument` class bases (after `BaseMixin`):

```python
class DocxDocument(
    BaseMixin,
    CreationMixin,
    ReadingMixin,
    ...
```

Add `"CreationMixin"` is NOT needed in `__all__` — only `DocxDocument` is public.

- [ ] **Step 5: Run tests to verify they pass**

Run: `pytest tests/test_creation.py -v`
Expected: all PASS

- [ ] **Step 6: Run full test suite to check no regressions**

Run: `pytest tests/ -x -q`
Expected: all pass (240 + new tests)

- [ ] **Step 7: Commit**

```bash
git add docx_mcp/document/creation.py docx_mcp/document/__init__.py tests/test_creation.py
git commit -m "feat: add CreationMixin for blank DOCX creation"
```

---

## Task 3: CreationMixin — template mode

**Files:**
- Modify: `tests/test_creation.py`
- Modify: `docx_mcp/document/creation.py` (already has template logic)

- [ ] **Step 1: Write failing tests for template mode**

Add to `tests/test_creation.py`:

```python
class TestCreateFromTemplate:
    @pytest.fixture()
    def template_docx(self, tmp_path: Path) -> Path:
        """Create a minimal .dotx template (same structure as .docx)."""
        from tests.conftest import _build_fixture
        path = tmp_path / "template.dotx"
        _build_fixture(path)
        return path

    def test_template_creates_docx_from_dotx(self, tmp_path: Path, template_docx: Path):
        out = tmp_path / "from_template.docx"
        doc = DocxDocument.create(str(out), template_path=str(template_docx))
        assert out.exists()
        assert "word/document.xml" in doc._trees
        doc.close()

    def test_template_preserves_existing_styles(self, tmp_path: Path, template_docx: Path):
        out = tmp_path / "from_template.docx"
        doc = DocxDocument.create(str(out), template_path=str(template_docx))
        styles = doc._trees["word/styles.xml"]
        style_ids = {s.get(f"{W}styleId") for s in styles.findall(f"{W}style")}
        assert "Heading1" in style_ids
        doc.close()

    def test_template_adds_custom_styles_if_missing(self, tmp_path: Path, template_docx: Path):
        out = tmp_path / "from_template.docx"
        doc = DocxDocument.create(str(out), template_path=str(template_docx))
        styles = doc._trees["word/styles.xml"]
        style_ids = {s.get(f"{W}styleId") for s in styles.findall(f"{W}style")}
        assert "CodeBlock" in style_ids
        assert "BlockQuote" in style_ids
        doc.close()

    def test_template_missing_raises(self, tmp_path: Path):
        out = tmp_path / "from_template.docx"
        with pytest.raises(FileNotFoundError):
            DocxDocument.create(str(out), template_path="/nonexistent.dotx")

    def test_template_bootstraps_numbering_if_missing(self, tmp_path: Path, template_docx: Path):
        """If template has no numbering.xml, it gets added."""
        # Remove numbering.xml from the template if present
        import zipfile
        import tempfile
        cleaned = tmp_path / "no_numbering.dotx"
        with zipfile.ZipFile(template_docx, "r") as zin:
            with zipfile.ZipFile(cleaned, "w") as zout:
                for item in zin.infolist():
                    if "numbering" not in item.filename:
                        zout.writestr(item, zin.read(item.filename))
        out = tmp_path / "output.docx"
        doc = DocxDocument.create(str(out), template_path=str(cleaned))
        assert "word/numbering.xml" in doc._trees
        doc.close()
```

- [ ] **Step 2: Run tests to verify they pass (or fail if template logic needs fixing)**

Run: `pytest tests/test_creation.py -v`
Expected: PASS (template logic is already in the create classmethod)

- [ ] **Step 3: Commit**

```bash
git add tests/test_creation.py
git commit -m "test: add template mode tests for CreationMixin"
```

---

## Task 4: Server tools — create_document

**Files:**
- Modify: `docx_mcp/server.py`
- Modify: `tests/test_creation.py`

- [ ] **Step 1: Write failing test for server tool**

Add to `tests/test_creation.py`:

```python
import json

class TestCreateDocumentTool:
    def test_creates_and_opens_document(self, tmp_path: Path):
        out = tmp_path / "new.docx"
        result = json.loads(server.create_document(str(out)))
        assert "paragraphs" in result  # returns get_info() output
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
        from tests.conftest import _build_fixture
        template = tmp_path / "tmpl.dotx"
        _build_fixture(template)
        out = tmp_path / "from_tmpl.docx"
        result = json.loads(server.create_document(str(out), template_path=str(template)))
        assert "paragraphs" in result
```

Note: `test_docx` fixture comes from `conftest.py`.

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_creation.py::TestCreateDocumentTool -v`
Expected: FAIL — `server.create_document` doesn't exist

- [ ] **Step 3: Add create_document tool to server.py**

Add after the `close_document` tool in `docx_mcp/server.py`:

```python
@mcp.tool()
def create_document(
    output_path: str,
    template_path: str | None = None,
) -> str:
    """Create a new blank .docx document (or from a .dotx template).

    The document is automatically opened for editing after creation.
    Use save_document to save changes, or start editing immediately
    with insert_text, add_table, etc.

    Args:
        output_path: Path for the new .docx file.
        template_path: Optional path to a .dotx template file.
    """
    global _doc
    if _doc is not None:
        _doc.close()
    _doc = DocxDocument.create(output_path, template_path=template_path)
    return _js(_doc.get_info())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_creation.py -v`
Expected: all PASS

- [ ] **Step 5: Run full suite**

Run: `pytest tests/ -x -q`
Expected: all pass

- [ ] **Step 6: Commit**

```bash
git add docx_mcp/server.py tests/test_creation.py
git commit -m "feat: add create_document MCP tool"
```

---

## Task 5: MarkdownConverter — smart typography

**Files:**
- Create: `docx_mcp/typography.py`
- Create: `tests/test_typography.py`

This is a pure function module — easiest to test in isolation first.

- [ ] **Step 1: Write failing tests**

Create `tests/test_typography.py`:

```python
"""Tests for smart typography conversion."""

from docx_mcp.typography import smartify


class TestSmartify:
    def test_double_quotes(self):
        assert smartify('"hello"') == '\u201Chello\u201D'

    def test_single_quotes(self):
        assert smartify("'hello'") == '\u2018hello\u2019'

    def test_apostrophe(self):
        assert smartify("it's") == 'it\u2019s'

    def test_twas_apostrophe(self):
        # Leading apostrophe after whitespace — still apostrophe (common case)
        assert smartify("'twas") == '\u2019twas' or smartify("'twas") == '\u2018twas'
        # The heuristic: preceded by start-of-string → left quote
        assert smartify("'twas") == '\u2018twas'

    def test_single_quotes_in_sentence(self):
        assert smartify("she said 'hello' today") == 'she said \u2018hello\u2019 today'

    def test_em_dash(self):
        assert smartify("word---word") == 'word\u2014word'

    def test_en_dash(self):
        assert smartify("word--word") == 'word\u2013word'

    def test_ellipsis(self):
        assert smartify("wait...") == 'wait\u2026'

    def test_no_change_for_plain_text(self):
        assert smartify("hello world") == "hello world"

    def test_mixed(self):
        result = smartify('"It\'s a test," she said---"really."')
        assert '\u201C' in result  # opening double quote
        assert '\u2019' in result  # apostrophe
        assert '\u2014' in result  # em dash
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_typography.py -v`
Expected: FAIL — module doesn't exist

- [ ] **Step 3: Implement smartify**

Create `docx_mcp/typography.py`:

```python
"""Smart typography: convert straight quotes and dashes to typographic equivalents."""

from __future__ import annotations

import re


def smartify(text: str) -> str:
    """Convert straight quotes, dashes, and ellipses to smart typography.

    Rules:
    - "text" → \u201Ctext\u201D (curly double quotes)
    - 'text' → \u2018text\u2019 (curly single quotes)
    - Apostrophe after letter → \u2019
    - --- → \u2014 (em dash)
    - -- → \u2013 (en dash)
    - ... → \u2026 (ellipsis)
    """
    # Order matters: longest patterns first
    # Em dash before en dash
    text = text.replace("---", "\u2014")
    text = text.replace("--", "\u2013")
    # Ellipsis
    text = text.replace("...", "\u2026")
    # Double quotes
    text = _convert_double_quotes(text)
    # Single quotes / apostrophes
    text = _convert_single_quotes(text)
    return text


def _convert_double_quotes(text: str) -> str:
    """Convert straight double quotes to curly."""
    result = []
    open_quote = True
    for char in text:
        if char == '"':
            result.append("\u201C" if open_quote else "\u201D")
            open_quote = not open_quote
        else:
            result.append(char)
    return "".join(result)


def _convert_single_quotes(text: str) -> str:
    """Convert straight single quotes to curly or apostrophe.

    Heuristic:
    - After a letter or digit → apostrophe (right single quote \u2019)
    - After whitespace, start-of-string, or opening punct → left single quote \u2018
    """
    result = []
    for i, char in enumerate(text):
        if char == "'":
            if i == 0 or (i > 0 and text[i - 1] in " \t\n\r([{"):
                result.append("\u2018")  # left single quote
            else:
                result.append("\u2019")  # apostrophe / right single quote
        else:
            result.append(char)
    return "".join(result)
```

- [ ] **Step 4: Run tests**

Run: `pytest tests/test_typography.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add docx_mcp/typography.py tests/test_typography.py
git commit -m "feat: add smart typography module"
```

---

## Task 6: MarkdownConverter — core rendering (headings, paragraphs, inline formatting)

**Files:**
- Create: `docx_mcp/markdown.py`
- Create: `tests/test_markdown.py`

### Context for implementer

The `MarkdownConverter` uses mistune 3.x to parse markdown into an AST, then walks the AST to build OOXML elements directly in the document's `_trees["word/document.xml"]`.

Mistune 3.x AST structure — each node is a dict with `"type"` key:
- `{"type": "heading", "attrs": {"level": 1}, "children": [{"type": "text", "raw": "Title"}]}`
- `{"type": "paragraph", "children": [{"type": "text", "raw": "Hello"}]}`
- `{"type": "strong", "children": [...]}`
- `{"type": "emphasis", "children": [...]}`
- `{"type": "strikethrough", "children": [...]}`
- `{"type": "codespan", "raw": "code"}`
- `{"type": "link", "attrs": {"url": "..."}, "children": [...]}`

- [ ] **Step 1: Write failing tests for headings and paragraphs**

Create `tests/test_markdown.py`:

```python
"""Tests for markdown-to-DOCX conversion."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument, W, W14
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
        both = [r for r in runs if r.find(f"{W}rPr/{W}b") is not None and r.find(f"{W}rPr/{W}i") is not None]
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
        assert "\u201C" in text  # left double quote
        assert "\u201D" in text  # right double quote

    def test_smart_typography_not_in_code(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, '`"not smart"`')
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        code_runs = [r for r in runs if r.find(f"{W}rPr/{W}rFonts") is not None]
        assert any('"not smart"' in (blank_doc._text(r) or "") for r in code_runs)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_markdown.py -v`
Expected: FAIL — module doesn't exist

- [ ] **Step 3: Implement MarkdownConverter core**

Create `docx_mcp/markdown.py`:

```python
"""Markdown to DOCX converter using mistune parser."""

from __future__ import annotations

import contextlib
from pathlib import Path

import mistune
from lxml import etree

from docx_mcp.document.base import RELS, W, W14, XML_SPACE, _preserve
from docx_mcp.typography import smartify


class MarkdownConverter:
    """Convert markdown to OOXML elements in a DocxDocument."""

    @classmethod
    def convert(
        cls,
        doc: object,
        text: str,
        *,
        base_dir: Path | None = None,
    ) -> None:
        """Parse markdown and populate the document body.

        Args:
            doc: A DocxDocument instance (opened via create()).
            text: Markdown text to convert.
            base_dir: Base directory for resolving relative image paths.
        """
        converter = cls(doc, base_dir=base_dir)
        converter._run(text)

    def __init__(self, doc: object, *, base_dir: Path | None = None):
        self._doc = doc
        self._base_dir = base_dir or Path.cwd()
        self._body = doc._trees["word/document.xml"].find(f"{W}body")
        self._footnote_map: dict[str, int] = {}  # markdown key → footnote id

    def _run(self, text: str) -> None:
        """Parse and render."""
        # Remove initial blank paragraph from skeleton
        for p in list(self._body.findall(f"{W}p")):
            self._body.remove(p)

        if not text.strip():
            return

        md = mistune.create_markdown(
            plugins=["table", "footnote", "strikethrough", "task_lists"],
        )
        # Get AST
        md_ast = mistune.create_markdown(
            renderer=None,
            plugins=["table", "footnote", "strikethrough", "task_lists"],
        )
        tokens = md_ast(text)

        # First pass: collect footnote definitions
        for token in tokens:
            if token["type"] == "footnote_list":
                self._process_footnote_definitions(token)

        # Second pass: render body tokens
        sect_pr = self._body.find(f"{W}sectPr")
        for token in tokens:
            if token["type"] == "footnote_list":
                continue  # Already processed
            elements = self._render_block(token)
            for el in elements:
                if sect_pr is not None:
                    sect_pr.addprevious(el)
                else:
                    self._body.append(el)

        self._doc._mark("word/document.xml")

    def _new_para(self, style: str | None = None) -> etree._Element:
        """Create a new <w:p> with paraId and optional style."""
        p = etree.Element(f"{W}p")
        p.set(f"{W14}paraId", self._doc._new_para_id())
        p.set(f"{W14}textId", "77777777")
        if style:
            ppr = etree.SubElement(p, f"{W}pPr")
            ps = etree.SubElement(ppr, f"{W}pStyle")
            ps.set(f"{W}val", style)
        return p

    def _make_run(self, text: str, *, bold: bool = False, italic: bool = False,
                   strike: bool = False, code: bool = False,
                   smart: bool = True) -> etree._Element:
        """Build a <w:r> with optional formatting."""
        r = etree.Element(f"{W}r")
        if bold or italic or strike or code:
            rpr = etree.SubElement(r, f"{W}rPr")
            if bold:
                etree.SubElement(rpr, f"{W}b")
            if italic:
                etree.SubElement(rpr, f"{W}i")
            if strike:
                etree.SubElement(rpr, f"{W}strike")
            if code:
                fonts = etree.SubElement(rpr, f"{W}rFonts")
                fonts.set(f"{W}ascii", "Courier New")
                fonts.set(f"{W}hAnsi", "Courier New")
        t = etree.SubElement(r, f"{W}t")
        final_text = smartify(text) if smart and not code else text
        _preserve(t, final_text)
        return r

    def _render_block(self, token: dict) -> list[etree._Element]:
        """Render a block-level token into OOXML elements."""
        t = token["type"]
        if t == "paragraph":
            return [self._render_paragraph(token)]
        elif t == "heading":
            return [self._render_heading(token)]
        elif t == "block_code":
            return self._render_code_block(token)
        elif t == "list":
            return self._render_list(token)
        elif t == "blockquote":
            return self._render_blockquote(token)
        elif t == "thematic_break":
            return [self._render_hr()]
        elif t == "table":
            return [self._render_table(token)]
        return []

    def _render_paragraph(self, token: dict) -> etree._Element:
        p = self._new_para()
        self._render_inline_children(p, token.get("children", []))
        return p

    def _render_heading(self, token: dict) -> etree._Element:
        level = token["attrs"]["level"]
        p = self._new_para(f"Heading{level}")
        self._render_inline_children(p, token.get("children", []))
        return p

    def _render_inline_children(self, parent: etree._Element, children: list[dict],
                                 *, bold: bool = False, italic: bool = False,
                                 strike: bool = False) -> None:
        """Recursively render inline tokens as runs."""
        for child in children:
            ct = child["type"]
            if ct == "text":
                parent.append(self._make_run(child["raw"], bold=bold, italic=italic, strike=strike))
            elif ct == "strong":
                self._render_inline_children(parent, child["children"], bold=True, italic=italic, strike=strike)
            elif ct == "emphasis":
                self._render_inline_children(parent, child["children"], bold=bold, italic=True, strike=strike)
            elif ct == "strikethrough":
                self._render_inline_children(parent, child["children"], bold=bold, italic=italic, strike=True)
            elif ct == "codespan":
                parent.append(self._make_run(child["raw"], code=True))
            elif ct == "link":
                self._render_link(parent, child)
            elif ct == "image":
                self._render_image(parent, child)
            elif ct == "softbreak":
                parent.append(self._make_run(" ", bold=bold, italic=italic, strike=strike))
            elif ct == "linebreak":
                r = etree.SubElement(parent, f"{W}r")
                etree.SubElement(r, f"{W}br")
            elif ct == "footnote_ref":
                self._render_footnote_ref(parent, child)

    def _render_code_block(self, token: dict) -> list[etree._Element]:
        """Render a fenced code block as CodeBlock-styled paragraphs."""
        code = token.get("raw", token.get("text", ""))
        lines = code.rstrip("\n").split("\n")
        result = []
        for line in lines:
            p = self._new_para("CodeBlock")
            p.append(self._make_run(line, code=True, smart=False))
            result.append(p)
        return result

    def _render_list(self, token: dict, depth: int = 0) -> list[etree._Element]:
        """Render bullet or numbered list."""
        ordered = token.get("attrs", {}).get("ordered", False)
        num_id = "2" if ordered else "1"  # from numbering.xml: 1=bullet, 2=numbered
        result = []
        for item in token.get("children", []):
            if item["type"] == "list_item":
                result.extend(self._render_list_item(item, num_id, depth))
        return result

    def _render_list_item(self, token: dict, num_id: str, depth: int) -> list[etree._Element]:
        """Render a single list item, handling nested lists."""
        result = []
        for child in token.get("children", []):
            if child["type"] == "paragraph":
                p = self._new_para()
                # Add numPr
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.SubElement(p, f"{W}pPr")
                    p.remove(ppr)
                    p.insert(0, ppr)
                num_pr = etree.SubElement(ppr, f"{W}numPr")
                ilvl = etree.SubElement(num_pr, f"{W}ilvl")
                ilvl.set(f"{W}val", str(depth))
                nid = etree.SubElement(num_pr, f"{W}numId")
                nid.set(f"{W}val", num_id)
                # Check for task list
                if "checked" in token.get("attrs", {}):
                    checked = token["attrs"]["checked"]
                    checkbox = "\u2611 " if checked else "\u2610 "
                    p.append(self._make_run(checkbox, smart=False))
                self._render_inline_children(p, child.get("children", []))
                result.append(p)
            elif child["type"] == "list":
                result.extend(self._render_list(child, depth=depth + 1))
        return result

    def _render_blockquote(self, token: dict, depth: int = 0) -> list[etree._Element]:
        """Render blockquote with increasing indent for nesting."""
        result = []
        for child in token.get("children", []):
            if child["type"] == "paragraph":
                p = self._new_para("BlockQuote")
                if depth > 0:
                    ppr = p.find(f"{W}pPr")
                    ind = ppr.find(f"{W}ind") if ppr is not None else None
                    if ind is None and ppr is not None:
                        ind = etree.SubElement(ppr, f"{W}ind")
                    if ind is not None:
                        ind.set(f"{W}left", str(720 * (depth + 1)))
                self._render_inline_children(p, child.get("children", []))
                result.append(p)
            elif child["type"] == "blockquote":
                result.extend(self._render_blockquote(child, depth=depth + 1))
        return result

    def _render_hr(self) -> etree._Element:
        """Render horizontal rule as paragraph with bottom border."""
        p = self._new_para()
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(p, f"{W}pPr")
            p.insert(0, ppr)
        pbdr = etree.SubElement(ppr, f"{W}pBdr")
        bottom = etree.SubElement(pbdr, f"{W}bottom")
        bottom.set(f"{W}val", "single")
        bottom.set(f"{W}sz", "6")
        bottom.set(f"{W}space", "1")
        bottom.set(f"{W}color", "auto")
        return p

    def _render_link(self, parent: etree._Element, token: dict) -> None:
        """Render a hyperlink."""
        url = token["attrs"]["url"]
        rid = self._add_hyperlink_rel(url)
        hyperlink = etree.SubElement(parent, f"{W}hyperlink")
        hyperlink.set(f"{R}id", rid)
        r = etree.SubElement(hyperlink, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "Hyperlink")
        color = etree.SubElement(rpr, f"{W}color")
        color.set(f"{W}val", "0563C1")
        u = etree.SubElement(rpr, f"{W}u")
        u.set(f"{W}val", "single")
        for child in token.get("children", []):
            if child["type"] == "text":
                t = etree.SubElement(r, f"{W}t")
                _preserve(t, smartify(child["raw"]))

    def _add_hyperlink_rel(self, url: str) -> str:
        """Add a hyperlink relationship and return the rId."""
        rels = self._doc._tree("word/_rels/document.xml.rels")
        max_rid = 0
        for r in rels.findall(f"{RELS}Relationship"):
            rid = r.get("Id", "")
            if rid.startswith("rId"):
                with contextlib.suppress(ValueError):
                    max_rid = max(max_rid, int(rid[3:]))
        rid = f"rId{max_rid + 1}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
        rel.set("Target", url)
        rel.set("TargetMode", "External")
        self._doc._mark("word/_rels/document.xml.rels")
        return rid

    def _render_image(self, parent: etree._Element, token: dict) -> None:
        """Render an image — embed local, hyperlink remote."""
        src = token["attrs"].get("url", token["attrs"].get("src", ""))
        alt = token.get("children", [{}])[0].get("raw", "") if token.get("children") else ""

        if src.startswith(("http://", "https://")):
            # Remote → hyperlink with alt text
            rid = self._add_hyperlink_rel(src)
            hyperlink = etree.SubElement(parent, f"{W}hyperlink")
            hyperlink.set(f"{R}id", rid)
            r = etree.SubElement(hyperlink, f"{W}r")
            rpr = etree.SubElement(r, f"{W}rPr")
            color = etree.SubElement(rpr, f"{W}color")
            color.set(f"{W}val", "0563C1")
            u = etree.SubElement(rpr, f"{W}u")
            u.set(f"{W}val", "single")
            t = etree.SubElement(r, f"{W}t")
            _preserve(t, alt or src)
        else:
            # Local → embed
            img_path = self._base_dir / src
            if not img_path.exists():
                parent.append(self._make_run(f"[Image not found: {src}]", smart=False))
                return
            # Use the document's insert_image internals
            self._embed_image(parent, str(img_path))

    def _embed_image(self, parent: etree._Element, image_path: str) -> None:
        """Embed a local image file."""
        import shutil

        src = Path(image_path)
        ext = src.suffix.lstrip(".")
        media_dir = self._doc.workdir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        existing = list(media_dir.glob("image*.*"))
        idx = len(existing) + 1
        filename = f"image{idx}.{ext}"
        shutil.copy2(str(src), str(media_dir / filename))

        # Relationship
        rels = self._doc._tree("word/_rels/document.xml.rels")
        max_rid = 0
        for r in rels.findall(f"{RELS}Relationship"):
            rid_str = r.get("Id", "")
            if rid_str.startswith("rId"):
                with contextlib.suppress(ValueError):
                    max_rid = max(max_rid, int(rid_str[3:]))
        rid = f"rId{max_rid + 1}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel.set("Target", f"media/{filename}")
        self._doc._mark("word/_rels/document.xml.rels")

        # Content type
        from docx_mcp.document.base import CT
        ct_tree = self._doc._tree("[Content_Types].xml")
        ct_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif"}
        content_type = ct_map.get(ext, f"image/{ext}")
        has_ext = any(d.get("Extension") == ext for d in ct_tree.findall(f"{CT}Default"))
        if not has_ext:
            default = etree.SubElement(ct_tree, f"{CT}Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)
            self._doc._mark("[Content_Types].xml")

        # Drawing XML (simplified inline image)
        ns_wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ns_pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        run = etree.SubElement(parent, f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        inline = etree.SubElement(drawing, f"{{{ns_wp}}}inline")
        extent = etree.SubElement(inline, f"{{{ns_wp}}}extent")
        extent.set("cx", "2000000")
        extent.set("cy", "2000000")
        graphic = etree.SubElement(inline, f"{{{ns_a}}}graphic")
        gdata = etree.SubElement(graphic, f"{{{ns_a}}}graphicData", uri=ns_pic)
        pic = etree.SubElement(gdata, f"{{{ns_pic}}}pic")
        blip_fill = etree.SubElement(pic, f"{{{ns_pic}}}blipFill")
        blip = etree.SubElement(blip_fill, f"{{{ns_a}}}blip")
        blip.set(f"{{{ns_r}}}embed", rid)

    def _render_table(self, token: dict) -> etree._Element:
        """Render a markdown table as w:tbl."""
        tbl = etree.Element(f"{W}tbl")
        tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
        style = etree.SubElement(tbl_pr, f"{W}tblStyle")
        style.set(f"{W}val", "TableGrid")
        tw = etree.SubElement(tbl_pr, f"{W}tblW")
        tw.set(f"{W}w", "0")
        tw.set(f"{W}type", "auto")
        # Add border
        tbl_borders = etree.SubElement(tbl_pr, f"{W}tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            bdr = etree.SubElement(tbl_borders, f"{W}{side}")
            bdr.set(f"{W}val", "single")
            bdr.set(f"{W}sz", "4")
            bdr.set(f"{W}space", "0")
            bdr.set(f"{W}color", "auto")

        # Head
        head = token.get("children", [])
        for i, child in enumerate(head):
            if child["type"] == "table_head":
                for row_token in child.get("children", []):
                    tr = self._render_table_row(row_token, bold=True)
                    tbl.append(tr)
            elif child["type"] == "table_body":
                for row_token in child.get("children", []):
                    tr = self._render_table_row(row_token, bold=False)
                    tbl.append(tr)
        return tbl

    def _render_table_row(self, token: dict, bold: bool = False) -> etree._Element:
        """Render a table row."""
        tr = etree.Element(f"{W}tr")
        tr.set(f"{W14}paraId", self._doc._new_para_id())
        tr.set(f"{W14}textId", "77777777")
        for cell_token in token.get("children", []):
            tc = etree.SubElement(tr, f"{W}tc")
            p = self._new_para()
            if cell_token.get("children"):
                self._render_inline_children(p, cell_token["children"], bold=bold)
            tc.append(p)
        return tr

    def _process_footnote_definitions(self, token: dict) -> None:
        """Process footnote definitions and add them to footnotes.xml."""
        fn_tree = self._doc._tree("word/footnotes.xml")
        if fn_tree is None:
            return

        existing = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        next_id = max(existing | {0}) + 1

        for item in token.get("children", []):
            if item["type"] == "footnote_item":
                key = item["attrs"].get("key", "")
                fn_el = etree.SubElement(fn_tree, f"{W}footnote")
                fn_el.set(f"{W}id", str(next_id))

                fn_para = etree.SubElement(fn_el, f"{W}p")
                fn_para.set(f"{W14}paraId", self._doc._new_para_id())
                fn_para.set(f"{W14}textId", "77777777")

                ppr = etree.SubElement(fn_para, f"{W}pPr")
                ps = etree.SubElement(ppr, f"{W}pStyle")
                ps.set(f"{W}val", "FootnoteText")

                # Footnote ref mark
                ref_run = etree.SubElement(fn_para, f"{W}r")
                ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
                ref_style = etree.SubElement(ref_rpr, f"{W}rStyle")
                ref_style.set(f"{W}val", "FootnoteReference")
                etree.SubElement(ref_run, f"{W}footnoteRef")

                # Space
                sp_run = etree.SubElement(fn_para, f"{W}r")
                sp_t = etree.SubElement(sp_run, f"{W}t")
                _preserve(sp_t, " ")

                # Text from children
                for child in item.get("children", []):
                    if child["type"] == "paragraph":
                        for inline in child.get("children", []):
                            if inline["type"] == "text":
                                txt_run = etree.SubElement(fn_para, f"{W}r")
                                txt_t = etree.SubElement(txt_run, f"{W}t")
                                _preserve(txt_t, smartify(inline["raw"]))

                self._footnote_map[key] = next_id
                next_id += 1

        self._doc._mark("word/footnotes.xml")

    def _render_footnote_ref(self, parent: etree._Element, token: dict) -> None:
        """Render a footnote reference in the body."""
        key = token.get("attrs", {}).get("key", token.get("key", ""))
        fn_id = self._footnote_map.get(key)
        if fn_id is None:
            return
        r = etree.SubElement(parent, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "FootnoteReference")
        fref = etree.SubElement(r, f"{W}footnoteReference")
        fref.set(f"{W}id", str(fn_id))
```

**Note to implementer:** The code above is a starting point. You will need to adjust based on the exact AST structure mistune 3.x produces. Install mistune and test interactively:
```python
import mistune
md = mistune.create_markdown(renderer=None, plugins=["table", "footnote", "strikethrough", "task_lists"])
print(md("# Hello\n\n**bold** *italic*"))
```

- [ ] **Step 4: Run tests**

Run: `pytest tests/test_markdown.py -v`
Expected: all PASS (may need adjustments based on mistune AST shape)

- [ ] **Step 5: Run full suite**

Run: `pytest tests/ -x -q`
Expected: all pass

- [ ] **Step 6: Commit**

```bash
git add docx_mcp/markdown.py tests/test_markdown.py
git commit -m "feat: add MarkdownConverter with headings, paragraphs, inline formatting"
```

---

## Task 7: MarkdownConverter — remaining constructs (lists, code blocks, blockquotes, tables, footnotes, images, task lists)

**Files:**
- Modify: `tests/test_markdown.py`

The implementation from Task 6 already includes rendering methods for all constructs. This task adds the remaining test coverage.

- [ ] **Step 1: Add tests for lists**

Add to `tests/test_markdown.py`:

```python
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
        assert style.get(f"{W}val") == "BlockQuote"

    def test_nested_blockquote(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "> Outer\n>> Inner")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) >= 2


class TestHorizontalRules:
    def test_hr(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "---")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 1
        border = paras[0].find(f"{W}pPr/{W}pBdr/{W}bottom")
        assert border is not None


class TestTables:
    def test_simple_table(self, blank_doc: DocxDocument):
        md = "| A | B |\n|---|---|\n| 1 | 2 |"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        tables = body.findall(f"{W}tbl")
        assert len(tables) == 1
        rows = tables[0].findall(f"{W}tr")
        assert len(rows) == 2  # header + 1 data row

    def test_table_header_bold(self, blank_doc: DocxDocument):
        md = "| H1 | H2 |\n|---|---|\n| d1 | d2 |"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        tbl = body.find(f"{W}tbl")
        first_row = tbl.findall(f"{W}tr")[0]
        # Header row cells should have bold runs
        first_cell = first_row.find(f"{W}tc")
        runs = first_cell.findall(f".//{W}r")
        assert any(r.find(f"{W}rPr/{W}b") is not None for r in runs)


class TestFootnotes:
    def test_footnote_creates_reference_and_definition(self, blank_doc: DocxDocument):
        md = "Text with a note[^1].\n\n[^1]: The note text."
        MarkdownConverter.convert(blank_doc, md)
        # Check body has footnoteReference
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        refs = list(body.iter(f"{W}footnoteReference"))
        assert len(refs) >= 1
        # Check footnotes.xml has the definition
        fn_tree = blank_doc._trees["word/footnotes.xml"]
        real_fns = [f for f in fn_tree.findall(f"{W}footnote") if f.get(f"{W}id") not in ("-1", "0")]
        assert len(real_fns) >= 1


class TestImages:
    def test_local_image_embedded(self, blank_doc: DocxDocument, tmp_path: Path):
        # Create a tiny PNG
        img = tmp_path / "test.png"
        img.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 50)
        MarkdownConverter.convert(blank_doc, f"![alt]({img})", base_dir=tmp_path)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        drawings = list(body.iter(f"{W}drawing"))
        assert len(drawings) >= 1

    def test_remote_image_becomes_hyperlink(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "![photo](https://example.com/img.png)")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        links = list(body.iter(f"{W}hyperlink"))
        assert len(links) >= 1

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
        assert "\u2611" in text  # checked
        assert "\u2610" in text  # unchecked


class TestMixed:
    def test_mixed_constructs(self, blank_doc: DocxDocument):
        md = "# Title\n\nParagraph with **bold**.\n\n- Item 1\n- Item 2\n\n| A | B |\n|---|---|\n| 1 | 2 |"
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        # Has heading, paragraphs, list items, and a table
        assert len(body.findall(f"{W}p")) >= 4
        assert len(body.findall(f"{W}tbl")) == 1
```

- [ ] **Step 2: Run all markdown tests**

Run: `pytest tests/test_markdown.py -v`
Expected: all PASS

- [ ] **Step 3: Run full suite**

Run: `pytest tests/ -x -q`
Expected: all pass

- [ ] **Step 4: Commit**

```bash
git add tests/test_markdown.py
git commit -m "test: add comprehensive markdown converter tests"
```

---

## Task 8: Server tool — create_from_markdown

**Files:**
- Modify: `docx_mcp/server.py`
- Modify: `tests/test_markdown.py`

- [ ] **Step 1: Write failing tests for server tool**

Add to `tests/test_markdown.py`:

```python
import json
from docx_mcp import server


class TestCreateFromMarkdownTool:
    def test_from_raw_text(self, tmp_path: Path):
        out = tmp_path / "from_md.docx"
        result = json.loads(server.create_from_markdown(str(out), markdown="# Hello\n\nWorld"))
        assert "paragraphs" in result
        assert server._doc is not None

    def test_from_file(self, tmp_path: Path):
        md_file = tmp_path / "input.md"
        md_file.write_text("# From File\n\nContent here.")
        out = tmp_path / "from_file.docx"
        result = json.loads(server.create_from_markdown(str(out), md_path=str(md_file)))
        assert "paragraphs" in result

    def test_both_inputs_raises(self, tmp_path: Path):
        out = tmp_path / "err.docx"
        md_file = tmp_path / "input.md"
        md_file.write_text("# Test")
        with pytest.raises(ValueError, match="mutually exclusive"):
            server.create_from_markdown(str(out), md_path=str(md_file), markdown="# Test")

    def test_no_input_raises(self, tmp_path: Path):
        out = tmp_path / "err.docx"
        with pytest.raises(ValueError, match="Either md_path or markdown"):
            server.create_from_markdown(str(out))

    def test_with_template(self, tmp_path: Path):
        from tests.conftest import _build_fixture
        template = tmp_path / "tmpl.dotx"
        _build_fixture(template)
        out = tmp_path / "from_md.docx"
        result = json.loads(
            server.create_from_markdown(str(out), markdown="# Hello", template_path=str(template))
        )
        assert "paragraphs" in result

    def test_image_paths_relative_to_md_file(self, tmp_path: Path):
        img = tmp_path / "photo.png"
        img.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 50)
        md_file = tmp_path / "doc.md"
        md_file.write_text("![pic](photo.png)")
        out = tmp_path / "output.docx"
        server.create_from_markdown(str(out), md_path=str(md_file))
        body = server._doc._trees["word/document.xml"].find(f"{W}body")
        drawings = list(body.iter(f"{W}drawing"))
        assert len(drawings) >= 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_markdown.py::TestCreateFromMarkdownTool -v`
Expected: FAIL — `server.create_from_markdown` doesn't exist

- [ ] **Step 3: Add create_from_markdown tool to server.py**

Add to `docx_mcp/server.py`:

```python
@mcp.tool()
def create_from_markdown(
    output_path: str,
    md_path: str | None = None,
    markdown: str | None = None,
    template_path: str | None = None,
) -> str:
    """Create a new .docx document from markdown content.

    Supports full GitHub-Flavored Markdown: headings, bold/italic/strikethrough,
    links, images, bullet/numbered/nested lists, code blocks, blockquotes,
    tables, footnotes, and task lists. Smart typography (curly quotes, em/en
    dashes, ellipses) is applied automatically.

    Args:
        output_path: Path for the new .docx file.
        md_path: Path to a .md file. Mutually exclusive with markdown.
        markdown: Raw markdown text. Mutually exclusive with md_path.
        template_path: Optional .dotx template for styles/page setup.
    """
    from pathlib import Path as P

    from docx_mcp.markdown import MarkdownConverter

    if md_path and markdown:
        raise ValueError("md_path and markdown are mutually exclusive")
    if not md_path and not markdown:
        raise ValueError("Either md_path or markdown must be provided")

    base_dir = None
    if md_path:
        md_file = P(md_path)
        if not md_file.exists():
            raise FileNotFoundError(f"Markdown file not found: {md_path}")
        markdown = md_file.read_text(encoding="utf-8")
        base_dir = md_file.parent

    global _doc
    if _doc is not None:
        _doc.close()
    _doc = DocxDocument.create(output_path, template_path=template_path)

    # Clear template body content if using a template
    if template_path:
        from docx_mcp.document import W as _W

        body = _doc._trees["word/document.xml"].find(f"{_W}body")
        sect_pr = body.find(f"{_W}sectPr")
        for child in list(body):
            if child.tag != f"{_W}sectPr":
                body.remove(child)

    MarkdownConverter.convert(_doc, markdown, base_dir=base_dir)
    return _js(_doc.get_info())
```

- [ ] **Step 4: Run tests**

Run: `pytest tests/test_markdown.py -v`
Expected: all PASS

- [ ] **Step 5: Run full suite**

Run: `pytest tests/ -x -q`
Expected: all pass

- [ ] **Step 6: Commit**

```bash
git add docx_mcp/server.py tests/test_markdown.py
git commit -m "feat: add create_from_markdown MCP tool"
```

---

## Task 9: E2E roundtrip tests

**Files:**
- Modify: `tests/test_e2e.py`

- [ ] **Step 1: Add roundtrip tests**

Add to `tests/test_e2e.py`:

```python
class TestCreateDocumentE2E:
    def test_create_save_reopen(self, tmp_path: Path):
        out = tmp_path / "created.docx"
        server.create_document(str(out))
        server.save_document(str(out))
        server.close_document()
        # Reopen
        result = json.loads(server.open_document(str(out)))
        assert "paragraphs" in result

    def test_create_edit_save_reopen(self, tmp_path: Path):
        out = tmp_path / "created.docx"
        info = json.loads(server.create_document(str(out)))
        # Insert text
        body = server._doc._trees["word/document.xml"].find(
            f"{server._doc.__class__.__mro__[0].__module__  and ''}{W}body"
        )
        # Simpler: just use insert_text on the first paragraph
        from docx_mcp.document import W, W14
        first_para = server._doc._trees["word/document.xml"].find(f"{W}body/{W}p")
        pid = first_para.get(f"{W14}paraId")
        server.insert_text(pid, "Hello from create", position="end")
        server.save_document(str(out))
        server.close_document()
        # Reopen and verify
        server.open_document(str(out))
        result = json.loads(server.search_text("Hello from create"))
        assert len(result) >= 1


class TestCreateFromMarkdownE2E:
    def test_markdown_roundtrip(self, tmp_path: Path):
        out = tmp_path / "from_md.docx"
        server.create_from_markdown(str(out), markdown="# Title\n\nParagraph with **bold**.")
        server.save_document(str(out))
        server.close_document()
        # Reopen
        server.open_document(str(out))
        headings = json.loads(server.get_headings())
        assert any(h["text"] == "Title" for h in headings)
        search = json.loads(server.search_text("bold"))
        assert len(search) >= 1

    def test_markdown_with_table_roundtrip(self, tmp_path: Path):
        md = "| A | B |\n|---|---|\n| 1 | 2 |"
        out = tmp_path / "table.docx"
        server.create_from_markdown(str(out), markdown=md)
        server.save_document(str(out))
        server.close_document()
        server.open_document(str(out))
        tables = json.loads(server.get_tables())
        assert len(tables) >= 1

    def test_create_then_track_changes(self, tmp_path: Path):
        out = tmp_path / "edited.docx"
        server.create_from_markdown(str(out), markdown="# Doc\n\nOriginal text here.")
        # Edit with tracked changes
        from docx_mcp.document import W, W14
        body = server._doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        text_para = [p for p in paras if "Original" in (server._doc._text(p) or "")]
        assert len(text_para) >= 1
        pid = text_para[0].get(f"{W14}paraId")
        server.delete_text(pid, "Original")
        server.insert_text(pid, "Updated")
        server.save_document(str(out))
        server.close_document()
        server.open_document(str(out))
        search = json.loads(server.search_text("Updated"))
        assert len(search) >= 1
```

- [ ] **Step 2: Run tests**

Run: `pytest tests/test_e2e.py -v -k "CreateDocument or CreateFromMarkdown"`
Expected: all PASS

- [ ] **Step 3: Commit**

```bash
git add tests/test_e2e.py
git commit -m "test: add e2e roundtrip tests for create_document and create_from_markdown"
```

---

## Task 10: Coverage + lint + docs

**Files:**
- Modify: `README.md`
- Modify: `docx_mcp/skill/SKILL.md`

- [ ] **Step 1: Run full suite with coverage**

Run: `pytest tests/ --cov=docx_mcp --cov-report=term-missing --cov-fail-under=100 -q`
Expected: 100% coverage. If not, add tests for uncovered lines.

- [ ] **Step 2: Run ruff**

Run: `ruff check . && ruff format --check .`
Expected: all pass. If not, fix with `ruff format .`

- [ ] **Step 3: Update README.md**

Update tool count from 43 to 45. Add new tools to the "Document Lifecycle" table:

```markdown
| `create_document` | Create a new blank .docx (or from a .dotx template) |
| `create_from_markdown` | Create a .docx from markdown (GitHub-Flavored Markdown) |
```

- [ ] **Step 4: Update SKILL.md**

Add new tools to the "Document Lifecycle" table in the skill file, and add a "Creating Documents" section to "Essential Patterns":

```markdown
### Creating Documents from Markdown

\```
1. create_from_markdown(output_path, markdown="# Title\n\nContent")
2. audit_document()  → verify integrity
3. save_document()   → save to disk
\```

Supports: headings, bold/italic/strikethrough, links, images, bullet/numbered/nested lists,
code blocks, blockquotes, tables, footnotes, task lists. Smart typography is applied automatically.
```

- [ ] **Step 5: Commit**

```bash
git add README.md docx_mcp/skill/SKILL.md
git commit -m "docs: update README and skill with new creation tools, 45 total"
```

- [ ] **Step 6: Final full test run**

Run: `pytest tests/ --cov=docx_mcp --cov-report=term-missing --cov-fail-under=100 -q`
Expected: all pass, 100% coverage

- [ ] **Step 7: Push and tag**

```bash
git push
git -c tag.gpgsign=false tag -a v0.3.0 -m "v0.3.0: add create_document and create_from_markdown"
git push origin v0.3.0
```
