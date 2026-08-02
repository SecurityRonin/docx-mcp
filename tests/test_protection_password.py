"""set_document_protection writes a password hash Word can actually verify.

The pre-existing test suite asserted only `has_password is True`, which passed
against a hash no Word build would accept. See
docs/IMPLEMENTATION-PLAN-table-and-deploy.md §C.1 and §M0.1.
"""

from __future__ import annotations

import base64
from pathlib import Path

import pytest
from docx_mcp.document.passwordhash import ooxml_password_hash

from docx_mcp.document import DocxDocument
from docx_mcp.document.base import W

# Same Tier-1 vector as tests/test_password_hash.py — Word-authored artifact,
# POI-authored answer key (apache/poi test-data/document/bug56076.docx).
POI_PASSWORD = "Example"
POI_SALT_B64 = "2Z+i7o/0EZyUNakVeWzU/w=="
POI_EXPECTED_HASH = "MUHbcmpC9AnlLsd9v3lW0j30y6E="


def _make_doc(tmp_path: Path) -> DocxDocument:
    return DocxDocument.create(str(tmp_path / "prot.docx"))


def _protection(doc: DocxDocument):
    return doc._require("word/settings.xml").find(f"{W}documentProtection")


class TestEndToEndAgainstTheWordVector:
    """T1 — pin the whole write path, not just the hash function."""

    def test_written_hash_matches_the_word_authored_value(self, tmp_path: Path, monkeypatch):
        doc = _make_doc(tmp_path)
        monkeypatch.setattr(
            "docx_mcp.document.protection.os.urandom",
            lambda n: base64.b64decode(POI_SALT_B64),
        )
        doc.set_document_protection(
            "trackedChanges",
            password=POI_PASSWORD,
            algorithm="SHA-1",
            spin_count=100000,
        )
        assert _protection(doc).get(f"{W}hash") == POI_EXPECTED_HASH

    def test_written_salt_is_the_base64_of_the_raw_bytes(self, tmp_path: Path, monkeypatch):
        doc = _make_doc(tmp_path)
        monkeypatch.setattr(
            "docx_mcp.document.protection.os.urandom",
            lambda n: base64.b64decode(POI_SALT_B64),
        )
        doc.set_document_protection("trackedChanges", password=POI_PASSWORD, algorithm="SHA-1")
        assert _protection(doc).get(f"{W}salt") == POI_SALT_B64


class TestProtectionAttributes:
    def test_hash_is_recomputable_from_the_written_attributes(self, tmp_path: Path):
        """Whatever salt is generated, the stored hash must follow from it."""
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        prot = _protection(doc)

        recomputed = ooxml_password_hash(
            "s3cret",
            base64.b64decode(prot.get(f"{W}salt")),
            int(prot.get(f"{W}cryptSpinCount")),
            "SHA-512",
        )
        assert base64.b64encode(recomputed).decode() == prot.get(f"{W}hash")

    def test_provider_type_matches_the_algorithm(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        assert _protection(doc).get(f"{W}cryptProviderType") == "rsaAES"

    def test_provider_type_for_sha1_is_rsafull(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret", algorithm="SHA-1")
        assert _protection(doc).get(f"{W}cryptProviderType") == "rsaFull"

    def test_algorithm_sid_tracks_the_algorithm(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret", algorithm="SHA-1")
        assert _protection(doc).get(f"{W}cryptAlgorithmSid") == "4"

    def test_default_algorithm_is_sha512(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        assert _protection(doc).get(f"{W}cryptAlgorithmSid") == "14"

    def test_salt_is_sixteen_bytes(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        assert len(base64.b64decode(_protection(doc).get(f"{W}salt"))) == 16

    def test_each_call_generates_a_fresh_salt(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        first = _protection(doc).get(f"{W}salt")
        doc.set_document_protection("readOnly", password="s3cret")
        assert _protection(doc).get(f"{W}salt") != first

    def test_spin_count_is_reported_as_written(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret", spin_count=500)
        assert _protection(doc).get(f"{W}cryptSpinCount") == "500"


class TestUnchangedBehaviour:
    """Existing contract must survive the fix."""

    def test_no_password_writes_no_hash(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        result = doc.set_document_protection("comments")
        assert result["has_password"] is False
        assert _protection(doc).get(f"{W}hash") is None

    def test_none_removes_protection(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        doc.set_document_protection("none")
        assert _protection(doc) is None

    def test_settings_part_is_marked_dirty(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        assert "word/settings.xml" in doc._modified

    def test_result_reports_has_password(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        assert doc.set_document_protection("forms", password="s3cret")["has_password"] is True

    def test_survives_a_save_reopen_round_trip(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        doc.set_document_protection("readOnly", password="s3cret")
        written = _protection(doc).get(f"{W}hash")
        doc.save()
        doc.close()

        reopened = DocxDocument(str(tmp_path / "prot.docx"))
        reopened.open()
        assert _protection(reopened).get(f"{W}hash") == written
        reopened.close()


class TestRejectsBadInput:
    def test_unknown_algorithm_is_rejected_by_name(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="SHA-999"):
            doc.set_document_protection("readOnly", password="x", algorithm="SHA-999")

    def test_spin_count_above_the_word_ceiling_is_rejected(self, tmp_path: Path):
        """MS-OE376 §2.15.1.28 caps the legacy spin count at 5,000,000."""
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError, match="5000000"):
            doc.set_document_protection("readOnly", password="x", spin_count=5_000_001)
