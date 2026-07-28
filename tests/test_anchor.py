"""
RED tests for three new track-changes features:

  1. Accepted-view flattener + cascading context anchor
     - delete_text / insert_text locate text using context_before / context_after
       rather than (or in addition to) exact paraId + substring
     - Anchor search falls back: full-context → context_before → context_after → unique-find
     - Accepted-view: w:ins content IS visible to the anchor; w:del content is NOT

  2. Normalisation / fuzzy matching
     - Smart quotes, NBSP, em-dash, en-dash normalised before matching
     - Whitespace collapsed for matching (bidirectional: match still wraps original chars)

  3. Multi-run spanning with rPr inheritance
     - delete_text succeeds when the target text spans multiple w:r elements
     - insert_text inherits rPr from the run at the insertion point
     - collapseDiff: only the changed portion becomes the tracked range
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import W14, W

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────


def _j(result: str) -> dict | list:
    return json.loads(result)


def _para_xml(para_id: str) -> etree._Element:
    """Return the raw paragraph element for a given paraId (requires open doc)."""
    doc = server._doc._trees["word/document.xml"]
    for p in doc.iter(f"{W}p"):
        pid = p.get(f"{W14}paraId")
        if pid == para_id:
            return p
    raise KeyError(f"Paragraph {para_id!r} not found")


def _para_text_visible(para_id: str) -> str:
    """
    Extract the 'accepted view' text from a paragraph:
    - include text from w:r and w:ins>w:r
    - exclude text from w:del
    """
    p = _para_xml(para_id)
    parts = []
    for child in p:
        tag = etree.QName(child.tag).localname
        if tag == "r":
            for t in child.iter(f"{W}t"):
                parts.append(t.text or "")
        elif tag == "ins":
            for r in child.iter(f"{W}r"):
                for t in r.iter(f"{W}t"):
                    parts.append(t.text or "")
        # w:del is skipped intentionally
    return "".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# Fixture: DOCX with a paragraph that already has tracked changes
# ─────────────────────────────────────────────────────────────────────────────


def _build_pretracked_docx(path: Path) -> None:
    """
    Build a DOCX where paragraph 00000010 already contains:
      - a w:ins wrapping "ALREADY_INSERTED"
      - a w:del wrapping "DELETED_WORD"
      - plain text "visible plain text around here"

    The accepted view reads: "visible plain text around here ALREADY_INSERTED"
    The raw XML also contains "DELETED_WORD" inside w:del (invisible to accepted view).
    """
    doc_xml = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000010" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">visible plain text around here </w:t></w:r>
      <w:del w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:delText>DELETED_WORD</w:delText></w:r>
      </w:del>
      <w:ins w:id="2" w:author="Alice" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t xml:space="preserve">ALREADY_INSERTED</w:t></w:r>
      </w:ins>
    </w:p>
    <w:p w14:paraId="00000011" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">Another paragraph with some text here.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000012" w14:textId="77777777">
      <w:r><w:t xml:space="preserve">Yet another paragraph entirely different.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

    content_types = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
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

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.strip())
        zf.writestr("_rels/.rels", top_rels.strip())
        zf.writestr("word/document.xml", doc_xml.strip())


@pytest.fixture()
def pretracked_docx(tmp_path: Path) -> Path:
    path = tmp_path / "pretracked.docx"
    _build_pretracked_docx(path)
    return path


# ─────────────────────────────────────────────────────────────────────────────
# Fixture: DOCX with smart-quote and NBSP content
# ─────────────────────────────────────────────────────────────────────────────


def _build_smartquote_docx(path: Path) -> None:
    """Paragraph 00000020 contains smart quotes, NBSP, em-dash."""
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<w:document\n"
        '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        "  <w:body>\n"
        '    <w:p w14:paraId="00000020" w14:textId="77777777">\n'
        '      <w:r><w:t xml:space="preserve">'
        "\u201cThe\u00a0contract\u201d\u2014effective\u00a0immediately."
        "</w:t></w:r>\n"
        "    </w:p>\n"
        "  </w:body>\n"
        "</w:document>"
    )

    content_types = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
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

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.strip())
        zf.writestr("_rels/.rels", top_rels.strip())
        zf.writestr("word/document.xml", doc_xml.strip())


@pytest.fixture()
def smartquote_docx(tmp_path: Path) -> Path:
    path = tmp_path / "smartquote.docx"
    _build_smartquote_docx(path)
    return path


# ─────────────────────────────────────────────────────────────────────────────
# 1. Accepted-view flattener + cascading context anchor
# ─────────────────────────────────────────────────────────────────────────────


class TestAcceptedViewFlattener:
    """
    The text visible in a paragraph (the 'accepted view') must:
      - include content inside w:ins wrappers
      - exclude content inside w:del wrappers
    The anchor search for insert_text / delete_text must operate on this
    accepted view, not raw XML text.
    """

    @pytest.fixture(autouse=True)
    def _open(self, pretracked_docx: Path):
        server.open_document(str(pretracked_docx))

    def test_accepted_view_includes_ins_content(self):
        """Text inside w:ins is part of the visible paragraph text."""
        visible = _para_text_visible("00000010")
        assert "ALREADY_INSERTED" in visible

    def test_accepted_view_excludes_del_content(self):
        """Text inside w:del is NOT part of the visible paragraph text."""
        visible = _para_text_visible("00000010")
        assert "DELETED_WORD" not in visible

    def test_delete_text_finds_text_after_existing_ins(self):
        """
        delete_text with context_before/context_after can target 'plain'
        even though the paragraph also has w:ins and w:del elements.
        The search must use the accepted view.
        """
        result = _j(
            server.delete_text(
                "00000010",
                "plain",
                context_before="visible ",
                context_after=" text",
            )
        )
        assert result["type"] == "deletion"
        # The word 'plain' should now be wrapped in w:del
        p = _para_xml("00000010")
        del_texts = [dt.text for dt in p.iter(f"{W}delText") if dt.text]
        assert any("plain" in t for t in del_texts)

    def test_insert_text_targets_text_adjacent_to_existing_ins(self):
        """
        insert_text with context anchoring places the insertion next to
        'ALREADY_INSERTED' (which is itself inside an existing w:ins).
        The accepted view must be aware of the w:ins content.
        """
        result = _j(
            server.insert_text(
                "00000010",
                "NEW_TEXT",
                context_before="ALREADY_INSERTED",
                context_after="",
            )
        )
        assert result["type"] == "insertion"
        visible = _para_text_visible("00000010")
        assert "NEW_TEXT" in visible


class TestCascadingContextAnchor:
    """
    Anchor search falls back through strategies:
      1. find + full context (context_before + context_after)
      2. find + context_before only
      3. find + context_after only
      4. find alone (only if unique across whole document)
    """

    @pytest.fixture(autouse=True)
    def _open(self, pretracked_docx: Path):
        server.open_document(str(pretracked_docx))

    def test_full_context_anchor(self):
        """Finds text using both context_before and context_after."""
        result = _j(
            server.delete_text(
                "00000011",
                "some",
                context_before="with ",
                context_after=" text",
            )
        )
        assert result["type"] == "deletion"

    def test_context_before_only_anchor(self):
        """Finds text using context_before when context_after is empty."""
        result = _j(
            server.delete_text(
                "00000011",
                "text",
                context_before="some ",
                context_after="",
            )
        )
        assert result["type"] == "deletion"

    def test_context_after_only_anchor(self):
        """Finds text using context_after when context_before is empty."""
        result = _j(
            server.delete_text(
                "00000011",
                "Another",
                context_before="",
                context_after=" paragraph",
            )
        )
        assert result["type"] == "deletion"

    def test_unique_find_anchor_no_context(self):
        """Finds text with no context if find is unique across the document."""
        # "entirely" only appears in paragraph 00000012
        result = _j(
            server.delete_text(
                "00000012",
                "entirely",
                context_before="",
                context_after="",
            )
        )
        assert result["type"] == "deletion"

    def test_find_in_multi_para_succeeds_with_explicit_para_id(self):
        """
        Find text in a specific paragraph succeeds even when the same text
        appears in other paragraphs — the caller already disambiguated by
        providing an explicit para_id (issue #6).
        """
        # "paragraph" appears in both 00000011 and 00000012,
        # but appears only once within 00000011 — should succeed
        result = _j(
            server.delete_text(
                "00000011",
                "paragraph",
                context_before="",
                context_after="",
            )
        )
        assert result["type"] == "deletion"


# ─────────────────────────────────────────────────────────────────────────────
# 2. Normalisation / fuzzy matching
# ─────────────────────────────────────────────────────────────────────────────


class TestNormalisation:
    """
    Smart quotes, NBSP, em-dash, en-dash are normalised to ASCII equivalents
    before anchor matching. The tracked change must wrap the ORIGINAL
    characters (not the normalised forms) in the output XML.
    """

    @pytest.fixture(autouse=True)
    def _open(self, smartquote_docx: Path):
        server.open_document(str(smartquote_docx))

    def test_delete_with_ascii_quotes_matches_smart_quotes(self):
        """
        LLM supplies ASCII double-quote; document contains U+201C/U+201D.
        delete_text should match via normalisation.
        """
        # LLM gives us plain ASCII "The contract" but doc has "The\u00a0contract"
        result = _j(
            server.delete_text(
                "00000020",
                '"The contract"',  # ASCII quotes
                context_before="",
                context_after="\u2014",
            )
        )
        assert result["type"] == "deletion"
        # The w:delText should contain the ORIGINAL smart-quote characters
        p = _para_xml("00000020")
        del_texts = "".join(dt.text for dt in p.iter(f"{W}delText") if dt.text)
        assert "\u201c" in del_texts or "\u201d" in del_texts

    def test_delete_with_regular_dash_matches_em_dash(self):
        """
        LLM supplies '--' or '-'; document contains U+2014 em-dash.
        """
        result = _j(
            server.delete_text(
                "00000020",
                "effective",
                context_before="-",  # normalised form of em-dash
                context_after=" immediately",
            )
        )
        assert result["type"] == "deletion"

    def test_delete_with_regular_space_matches_nbsp(self):
        """
        LLM supplies regular space; document contains U+00A0 non-breaking space.
        """
        result = _j(
            server.delete_text(
                "00000020",
                "contract",
                context_before="The ",  # regular space (doc has NBSP)
                context_after="\u201d",
            )
        )
        assert result["type"] == "deletion"

    def test_normalisation_does_not_alter_original_text(self):
        """
        After a normalised delete, rejecting the change restores the ORIGINAL
        text including smart quotes and NBSP — normalisation is for matching
        only, not for storage.
        """
        server.delete_text(
            "00000020",
            '"The contract"',
            context_before="",
            context_after="\u2014",
        )
        server.reject_changes()
        p = _para_xml("00000020")
        full_text = "".join(t.text or "" for t in p.iter(f"{W}t"))
        assert "\u201c" in full_text  # original smart quote restored


# ─────────────────────────────────────────────────────────────────────────────
# 3. Multi-run spanning with rPr inheritance
# ─────────────────────────────────────────────────────────────────────────────


class TestMultiRunSpanning:
    """
    The existing fixture paragraph 00000006 has:
      <w:r><w:t xml:space="preserve">First </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> last</w:t></w:r>

    The text "First bold" spans two runs with different formatting.
    Currently delete_text raises ValueError on this. The new implementation
    must handle it.
    """

    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_delete_text_spanning_two_runs(self):
        """delete_text succeeds when text spans a run boundary."""
        result = _j(server.delete_text("00000006", "First bold"))
        assert result["type"] == "deletion"
        # The deleted text should appear inside w:del
        p = _para_xml("00000006")
        del_texts = "".join(dt.text or "" for dt in p.iter(f"{W}delText"))
        assert "First" in del_texts
        assert "bold" in del_texts

    def test_delete_spanning_run_preserves_rpr_per_segment(self):
        """
        When deletion spans two runs with different rPr, the w:del element
        contains multiple w:r children — one per original run segment —
        each preserving its original rPr.
        """
        server.delete_text("00000006", "First bold")
        p = _para_xml("00000006")
        del_els = list(p.iter(f"{W}del"))
        assert len(del_els) >= 1
        del_runs = list(del_els[0].findall(f"{W}r"))
        # Should have at least 2 runs: one plain, one bold
        assert len(del_runs) >= 2
        # The bold run's rPr should include w:b
        bold_runs = [
            r
            for r in del_runs
            if r.find(f"{W}rPr") is not None and r.find(f"{W}rPr").find(f"{W}b") is not None
        ]
        assert len(bold_runs) >= 1

    def test_insert_inherits_rpr_from_adjacent_run(self):
        """
        insert_text placed after 'bold' (which has w:b) inherits the bold rPr
        in the inserted w:ins > w:r > w:rPr.
        """
        result = _j(
            server.insert_text(
                "00000006",
                "INSERTED",
                context_before="bold",
                context_after=" last",
            )
        )
        assert result["type"] == "insertion"
        p = _para_xml("00000006")
        ins_els = list(p.iter(f"{W}ins"))
        assert len(ins_els) >= 1
        # The inserted run should have w:rPr with w:b (inherited from 'bold' run)
        ins_run = ins_els[-1].find(f"{W}r")
        assert ins_run is not None
        rpr = ins_run.find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}b") is not None

    def test_collapse_diff_only_marks_changed_portion(self):
        """
        When find='First bold last' and replace='First red last',
        only 'bold' becomes the tracked deletion and 'red' the tracked insertion.
        The leading 'First ' and trailing ' last' are left as plain runs.
        """
        result = _j(
            server.replace_text(
                "00000006",
                find="First bold last",
                replace="First red last",
            )
        )
        assert result["type"] == "replacement"
        p = _para_xml("00000006")
        del_texts = "".join(dt.text or "" for dt in p.iter(f"{W}delText"))
        ins_texts = "".join(
            t.text or "" for t in p.iter(f"{W}ins") for t2 in t.iter(f"{W}t") for _ in [None]
        )  # noqa: E501
        # Simpler: collect ins w:t text
        ins_texts = ""
        for ins in p.iter(f"{W}ins"):
            for t in ins.iter(f"{W}t"):
                ins_texts += t.text or ""
        assert del_texts.strip() == "bold"
        assert ins_texts.strip() == "red"
        # 'First ' and ' last' should NOT appear in del or ins
        assert "First" not in del_texts
        assert "last" not in del_texts
