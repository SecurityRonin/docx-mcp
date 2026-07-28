"""Tests for mike integration: get_body_text, replace_text return, hyperlink, NBSP."""

from __future__ import annotations

import json

import pytest

from docx_mcp import server
from docx_mcp.document.tracks import _flatten_para

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


def test_flatten_para_includes_hyperlink_text(mike_corpus_docx):
    """_flatten_para must include text inside w:hyperlink runs."""
    server.open_document(str(mike_corpus_docx))
    doc = server._doc._require("word/document.xml")
    # Find paragraph 00000101 which contains a w:hyperlink with "thirty calendar days"
    para = None
    for p in doc.iter(f"{W}p"):
        if p.get(f"{W14}paraId") == "00000101":
            para = p
            break
    assert para is not None, "paragraph 00000101 not found"
    slots = _flatten_para(para)
    text = "".join(s.char for s in slots)
    assert "thirty calendar days" in text, (
        f"hyperlink text not found in _flatten_para output; got: {text!r}"
    )


def test_get_body_text_returns_body_string(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "body" in result
    assert isinstance(result["body"], str)
    assert len(result["body"]) > 0


def test_get_body_text_includes_hyperlink_text(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    # Hyperlink text must appear in body (not invisible)
    assert "thirty calendar days" in result["body"]


def test_get_body_text_accepted_view_excludes_deleted(mike_corpus_docx):
    """Accepted view: w:del text must NOT appear, w:ins text MUST appear."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "one hundred" not in result["body"]  # inside w:del → excluded
    assert "two hundred" in result["body"]  # inside w:ins → included


def test_get_body_text_returns_footnotes(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "footnotes" in result
    # Footnote id=1 in _FOOTNOTES_XML contains "See appendix A for supporting evidence."
    # Reserved ids -1 (separator) and 0 (continuation) must be excluded.
    assert "See appendix A" in result["footnotes"]


def test_get_body_text_real_contract(tmp_path):
    """Smoke test against real externally-sourced document."""
    import os

    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "real_contract.docx")
    if not os.path.exists(fixture):
        pytest.skip("real fixture not present")
    server.open_document(fixture)
    result = json.loads(server.get_body_text())
    assert len(result["body"]) > 100


def test_replace_text_returns_del_id_and_ins_id(mike_corpus_docx):
    """replace_text must return both del_id and ins_id for Accept/Reject card UI."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(
        server.replace_text(
            para_id="00000108",
            find="initial deposit",
            replace="upfront payment",
        )
    )
    assert "del_id" in result, "must include del_id"
    assert "ins_id" in result, "must include ins_id"
    assert result["del_id"] is not None
    assert result["ins_id"] is not None
    assert result["del_id"] != result["ins_id"]
    assert result["type"] == "replacement"


def test_replace_text_pure_deletion_has_no_ins_id(mike_corpus_docx):
    """Pure deletion (empty replace) must have ins_id=None."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(
        server.replace_text(
            para_id="00000109",
            find="ongoing ",
            replace="",
        )
    )
    assert result["del_id"] is not None
    assert result["ins_id"] is None


# ── Narrow no-break space normalization ──────────────────────────────────────


def test_replace_finds_text_with_narrow_nbsp(mike_corpus_docx):
    """replace_text must match '$5 000' when find uses regular space '$5 000'.
    Paragraph 00000102 contains U+202F narrow no-break space between '5' and '000'.
    """
    server.open_document(str(mike_corpus_docx))
    result = json.loads(
        server.replace_text(
            para_id="00000102",
            find="$5 000",  # normal space in find
            replace="$6 000",
        )
    )
    assert result["del_id"] is not None


# ── Smart quote normalization ─────────────────────────────────────────────────


def test_replace_finds_smart_quotes_with_ascii(mike_corpus_docx):
    """replace_text must match curly "Effective Date" when find uses ASCII double quotes."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(
        server.replace_text(
            para_id="00000103",
            find='"Effective Date"',  # ASCII double quotes
            replace='"Execution Date"',
        )
    )
    assert result["del_id"] is not None


# ── Context disambiguation ─────────────────────────────────────────────────────


def test_search_text_returns_both_ambiguous_matches(mike_corpus_docx):
    """'twelve months' appears in two paragraphs — search must return both."""
    server.open_document(str(mike_corpus_docx))
    # search_text returns a list of dicts with 'paraId' (camelCase) key
    results = json.loads(server.search_text(query="twelve months"))
    para_ids = [m["paraId"] for m in results]
    assert "00000104" in para_ids
    assert "00000105" in para_ids


def test_replace_with_context_disambiguates(mike_corpus_docx, tmp_path):
    """replace_text with context_before disambiguates between two identical phrases."""
    server.open_document(str(mike_corpus_docx))
    # Target only Section 1
    result = json.loads(
        server.replace_text(
            para_id="00000104",
            find="twelve months",
            replace="eighteen months",
            context_before="term shall be ",
        )
    )
    assert result["del_id"] is not None
    # Section 2 paragraph must be unchanged — save and re-read to verify
    out = tmp_path / "disambig_out.docx"
    server.save_document(str(out))
    server.open_document(str(out))
    body = json.loads(server.get_body_text())["body"]
    assert "twelve months" in body  # section 2 still has it


# ── Accepted-view with pre-existing tracked changes ───────────────────────────


def test_accepted_view_shows_ins_hides_del(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    body = json.loads(server.get_body_text())["body"]
    assert "two hundred" in body
    assert "one hundred" not in body


def test_replace_inside_existing_ins_raises_known_limitation(mike_corpus_docx):
    """KNOWN LIMITATION: replace_text on text inside a pre-existing w:ins raises ValueError.
    Paragraph 00000107 has w:del('one hundred') and w:ins('two hundred').
    The implementation blocks editing text inside an unresolved tracked insertion —
    accept/reject the existing change before issuing a new replacement.
    """
    server.open_document(str(mike_corpus_docx))
    with pytest.raises(ValueError, match="existing w:ins"):
        server.replace_text(
            para_id="00000107",
            find="two hundred",
            replace="three hundred",
        )


# ── Batch edits ───────────────────────────────────────────────────────────────


def test_batch_three_replacements_in_one_session(mike_corpus_docx, tmp_path):
    """Three replace_text calls in one open/save session all persist."""
    server.open_document(str(mike_corpus_docx))
    server.replace_text(para_id="00000108", find="initial deposit", replace="upfront payment")
    server.replace_text(para_id="00000109", find="maintenance fee", replace="service charge")
    server.replace_text(para_id="0000010A", find="termination penalty", replace="exit fee")
    out = tmp_path / "batch_out.docx"
    server.save_document(str(out))
    # Re-open and verify all three changes are tracked
    server.open_document(str(out))
    # get_tracked_changes returns a list of dicts with keys: type, change_id, author, date, para_id, text  # noqa: E501
    changes = json.loads(server.get_tracked_changes())
    by_para: dict[str, list[str]] = {}
    for c in changes:
        by_para.setdefault(c["para_id"], []).append(c["text"])
    # Each paragraph must have its own tracked change (catch silent per-replacement failure)
    para_108_texts = " ".join(by_para.get("00000108", []))
    para_109_texts = " ".join(by_para.get("00000109", []))
    para_10a_texts = " ".join(by_para.get("0000010A", []))
    assert "initial deposit" in para_108_texts or "upfront payment" in para_108_texts
    assert "maintenance fee" in para_109_texts or "service charge" in para_109_texts
    assert "termination penalty" in para_10a_texts or "exit fee" in para_10a_texts


# ── Accept/Reject roundtrip ───────────────────────────────────────────────────


def test_accept_change_from_replace_text(mike_corpus_docx, tmp_path):
    server.open_document(str(mike_corpus_docx))
    r = json.loads(server.replace_text(para_id="00000108", find="initial", replace="upfront"))
    del_id = r["del_id"]
    ins_id = r["ins_id"]
    # Accept del (accept the deletion — remove the old text markup)
    server.accept_change(del_id)
    # Accept ins (accept the insertion — make the new text permanent)
    server.accept_change(ins_id)
    out = tmp_path / "accepted.docx"
    server.save_document(str(out))
    server.open_document(str(out))
    body = json.loads(server.get_body_text())["body"]
    assert "upfront" in body
    assert "initial" not in body
    # Verify changes are PERSISTED (not just applied to in-memory accepted view)
    changes = json.loads(server.get_tracked_changes())
    remaining_ids = {c["change_id"] for c in changes}
    assert del_id not in remaining_ids, (
        f"del change {del_id} must be gone after accept; remaining: {remaining_ids}"
    )  # noqa: E501
    assert ins_id not in remaining_ids, (
        f"ins change {ins_id} must be gone after accept; remaining: {remaining_ids}"
    )  # noqa: E501


# ── Real fixture smoke tests ──────────────────────────────────────────────────


@pytest.mark.parametrize(
    "fixture_name",
    [
        "real_contract.docx",
        "real_tracked_changes.docx",
        "real_hyperlinks_footnotes.docx",
    ],
)
def test_get_body_text_real_fixtures(fixture_name, tmp_path):
    import os

    fixture = os.path.join(os.path.dirname(__file__), "fixtures", fixture_name)
    if not os.path.exists(fixture):
        pytest.skip(f"fixture {fixture_name} not present")
    server.open_document(fixture)
    result = json.loads(server.get_body_text())
    assert len(result["body"]) > 50, "real doc must have substantial body text"
    audit = json.loads(server.audit_document())
    # Check critical structural categories; ignore table column counts (merged cells
    # in real docs legitimately produce inconsistent counts but are valid OOXML).
    assert audit["footnotes"]["valid"], f"{fixture_name}: footnote audit failed"
    assert audit["paraids"]["valid"], f"{fixture_name}: paraId audit failed"
    assert not audit["relationships"]["missing_targets"], f"{fixture_name}: broken relationships"


# ── Issue #6: replace_text paragraph scope (no document-global uniqueness gate) ─


def test_replace_text_same_phrase_multiple_paras_with_para_id_succeeds(mike_corpus_docx, tmp_path):
    """replace_text with explicit para_id succeeds even when find text appears elsewhere.

    The mike_corpus fixture has "twelve months" in at least two paragraphs
    (para 00000104 in section 1 and another para in section 2 — verified by
    test_replace_with_context_disambiguates).  With the _doc_norm_count gate
    removed, replace_text should locate the text within the targeted paragraph
    and succeed without demanding context_before/context_after.
    """
    server.open_document(str(mike_corpus_docx))

    # No context — the call must NOT raise ValueError
    result = json.loads(
        server.replace_text(
            para_id="00000104",
            find="twelve months",
            replace="eighteen months",
        )
    )
    assert result["del_id"] is not None, "replace_text must produce a deletion"

    # Save and re-read to verify scoped replacement
    out = tmp_path / "scope_fix.docx"
    server.save_document(str(out))
    server.open_document(str(out))
    body = json.loads(server.get_body_text())["body"]

    # The other paragraph must still contain the original phrase
    assert "twelve months" in body, (
        "other paragraph must be untouched — 'twelve months' should still appear"
    )
    # The target paragraph should reflect the replacement
    assert "eighteen months" in body, "target paragraph must contain the replacement text"
