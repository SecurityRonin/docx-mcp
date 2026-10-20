"""Tests for where tracked replace_text places its insertion."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document.base import W


def _j(result: str) -> dict | list:
    return json.loads(result)


def _text(para_id: str) -> str:
    return _j(server.get_paragraph(para_id))["text"]


class TestReplaceTextPlacement:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_pure_insertion_mid_run(self):
        """An insertion with nothing deleted lands at its position, not at paragraph end."""
        server.replace_text(para_id="00000004", find="term is 30", replace="term is only 30")
        assert _text("00000004") == "The contract term is only 30 days from the effective date."

    def test_pure_insertion_at_run_boundary(self):
        """Inserting before a run that starts the change goes before that run."""
        server.replace_text(para_id="00000006", find="First bold", replace="First very bold")
        assert _text("00000006") == "First very bold last"

    def test_pure_insertion_takes_preceding_formatting(self):
        """Text inserted before a bold word is formatted like the text before it."""
        server.replace_text(para_id="00000006", find="First bold", replace="First very bold")
        doc = server._require_doc()
        para = doc._find_para(doc._require("word/document.xml"), "00000006")
        ins_runs = list(para.iter(f"{W}ins"))[0].findall(f"{W}r")
        assert ins_runs[0].find(f"{W}rPr/{W}b") is None

    def test_pure_insertion_at_paragraph_end(self):
        server.replace_text(
            para_id="00000004", find="effective date.", replace="effective date. Renewable."
        )
        assert (
            _text("00000004") == "The contract term is 30 days from the effective date. Renewable."
        )

    def test_pure_insertion_is_tracked_and_rejectable(self):
        result = _j(
            server.replace_text(para_id="00000004", find="term is 30", replace="term is only 30")
        )
        assert result["del_id"] is None
        assert result["ins_id"] is not None
        server.reject_changes()
        assert _text("00000004") == "The contract term is 30 days from the effective date."

    def test_second_replacement_earlier_in_paragraph(self):
        """A replacement before an existing tracked deletion anchors on its own deletion."""
        server.replace_text(para_id="00000004", find="effective date", replace="start date")
        server.replace_text(para_id="00000004", find="30 days", replace="45 days")
        assert _text("00000004") == "The contract term is 45 days from the start date."

    def test_second_replacement_earlier_accepts_cleanly(self):
        server.replace_text(para_id="00000004", find="effective date", replace="start date")
        server.replace_text(para_id="00000004", find="30 days", replace="45 days")
        server.accept_changes()
        assert _text("00000004") == "The contract term is 45 days from the start date."
