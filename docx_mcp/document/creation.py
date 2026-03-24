"""Creation mixin: create blank DOCX documents and from templates."""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .base import _now_iso


class CreationMixin:
    """Document creation operations."""

    @classmethod
    def create(cls, output_path: str, template_path: str | None = None) -> CreationMixin:
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


def _write_blank_skeleton(path: Path) -> None:
    """Write a minimal valid .docx ZIP archive."""
    import random

    def _pid() -> str:
        return f"{random.randint(1, 0x7FFFFFFF):08X}"

    # Generate unique paraIds for all paragraphs in the skeleton
    pids = {"body1": _pid(), "fn_sep": _pid(), "fn_cont": _pid(),
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


def _ensure_custom_styles(doc: CreationMixin) -> None:
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


def _ensure_numbering(doc: CreationMixin) -> None:
    """Bootstrap numbering.xml if missing in a template."""
    from lxml import etree

    from .base import CT, RELS

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
