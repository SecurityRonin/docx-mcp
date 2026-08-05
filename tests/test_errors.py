"""Tests for the stable error taxonomy (Phase 1a)."""

from __future__ import annotations

import pytest

from docx_mcp.document.errors import DocxMcpError, ErrCode


class TestErrCode:
    def test_all_members_are_strings(self):
        for member in ErrCode:
            assert isinstance(member, str), f"{member} is not a str"

    def test_style_not_found_is_str_instance(self):
        assert isinstance(ErrCode.STYLE_NOT_FOUND, str)

    def test_values_match_names(self):
        assert ErrCode.STYLE_NOT_FOUND.value == "STYLE_NOT_FOUND"
        assert ErrCode.PARA_NOT_FOUND.value == "PARA_NOT_FOUND"
        assert ErrCode.BOOKMARK_NOT_FOUND.value == "BOOKMARK_NOT_FOUND"
        assert ErrCode.BOOKMARK_DANGLING.value == "BOOKMARK_DANGLING"
        assert ErrCode.PART_NOT_FOUND.value == "PART_NOT_FOUND"
        assert ErrCode.INVALID_RELATIONSHIP.value == "INVALID_REL"
        assert ErrCode.NUMBERING_ORPHAN.value == "NUMBERING_ORPHAN"
        assert ErrCode.OOXML_INVALID.value == "OOXML_INVALID"
        assert ErrCode.PII_DEPS_MISSING.value == "PII_DEPS_MISSING"
        assert ErrCode.NO_OPEN_DOCUMENT.value == "NO_OPEN_DOCUMENT"
        assert ErrCode.XPATH_ERROR.value == "XPATH_ERROR"
        assert ErrCode.ENCRYPTED_DOCUMENT.value == "ENCRYPTED_DOCUMENT"

    def test_member_count(self):
        """The taxonomy is a published contract — growing it is a deliberate act."""
        assert len(ErrCode) == 14


class TestDocxMcpError:
    def test_is_exception_subclass(self):
        assert issubclass(DocxMcpError, Exception)

    def test_can_be_caught_as_exception(self):
        with pytest.raises(DocxMcpError):
            raise DocxMcpError(ErrCode.STYLE_NOT_FOUND, "style missing")

    def test_can_be_caught_as_docxmcperror(self):
        with pytest.raises(DocxMcpError):
            raise DocxMcpError(ErrCode.PARA_NOT_FOUND, "para missing")

    def test_code_attribute(self):
        err = DocxMcpError(ErrCode.OOXML_INVALID, "bad xml")
        assert err.code is ErrCode.OOXML_INVALID

    def test_message_via_str(self):
        err = DocxMcpError(ErrCode.XPATH_ERROR, "xpath failed")
        assert str(err) == "xpath failed"

    def test_hint_default_is_empty_string(self):
        err = DocxMcpError(ErrCode.NO_OPEN_DOCUMENT, "no doc")
        assert err.hint == ""

    def test_hint_explicit(self):
        err = DocxMcpError(ErrCode.PII_DEPS_MISSING, "missing deps", hint="pip install presidio")
        assert err.hint == "pip install presidio"

    def test_to_dict_keys(self):
        err = DocxMcpError(ErrCode.STYLE_NOT_FOUND, "missing style", hint="check styles")
        d = err.to_dict()
        assert set(d.keys()) == {"error", "message", "hint"}

    def test_to_dict_error_is_string_not_enum(self):
        err = DocxMcpError(ErrCode.STYLE_NOT_FOUND, "missing style")
        d = err.to_dict()
        assert d["error"] == "STYLE_NOT_FOUND"
        assert isinstance(d["error"], str)
        assert not isinstance(d["error"], ErrCode)

    def test_to_dict_message(self):
        err = DocxMcpError(ErrCode.PART_NOT_FOUND, "part gone")
        assert err.to_dict()["message"] == "part gone"

    def test_to_dict_hint_default_empty(self):
        err = DocxMcpError(ErrCode.BOOKMARK_NOT_FOUND, "no bookmark")
        assert err.to_dict()["hint"] == ""

    def test_to_dict_hint_populated(self):
        err = DocxMcpError(ErrCode.NUMBERING_ORPHAN, "orphan", hint="fix numbering")
        assert err.to_dict()["hint"] == "fix numbering"
