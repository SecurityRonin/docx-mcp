"""Tests for P9.5: PDF export — convert_to_pdf."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from docx_mcp.document import DocxDocument


def _make_doc(tmp_path: Path) -> DocxDocument:
    out = str(tmp_path / "test.docx")
    return DocxDocument.create(out)


def _fake_soffice(cmd, **kwargs):
    """subprocess.run stand-in that writes the PDF LibreOffice would produce."""
    src = Path(cmd[-1])
    outdir = Path(cmd[cmd.index("--outdir") + 1])
    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / (src.stem + ".pdf")).write_bytes(b"%PDF-1.4\n")
    return MagicMock(returncode=0, stderr="")


class TestConvertToPdf:
    def test_returns_pdf_path_key(self, tmp_path: Path):
        """convert_to_pdf returns dict with 'pdf_path' key."""
        doc = _make_doc(tmp_path)
        pdf_out = str(tmp_path / "out.pdf")

        with (
            patch("shutil.which", return_value="/usr/bin/libreoffice"),
            patch("subprocess.run", side_effect=_fake_soffice),
        ):
            result = doc.convert_to_pdf(pdf_out)

        assert "pdf_path" in result

    def test_raises_if_libreoffice_not_found(self, tmp_path: Path):
        """convert_to_pdf raises RuntimeError when LibreOffice is not installed."""
        doc = _make_doc(tmp_path)
        with (
            patch("shutil.which", return_value=None),
            pytest.raises(RuntimeError, match="LibreOffice"),
        ):  # noqa: E501
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

    def test_calls_libreoffice_subprocess(self, tmp_path: Path):
        """convert_to_pdf calls libreoffice --headless --convert-to pdf."""
        doc = _make_doc(tmp_path)
        pdf_out = str(tmp_path / "out.pdf")

        with (
            patch("shutil.which", return_value="/usr/bin/libreoffice"),
            patch("subprocess.run", side_effect=_fake_soffice) as mock_run,
        ):
            doc.convert_to_pdf(pdf_out)

        call_args = mock_run.call_args
        cmd = call_args[0][0]
        assert "--headless" in cmd
        assert "--convert-to" in cmd
        assert "pdf" in cmd

    def test_raises_if_subprocess_fails(self, tmp_path: Path):
        """convert_to_pdf raises RuntimeError when libreoffice exits non-zero."""
        doc = _make_doc(tmp_path)
        pdf_out = str(tmp_path / "out.pdf")

        with (
            patch("shutil.which", return_value="/usr/bin/libreoffice"),
            patch("subprocess.run") as mock_run,
        ):
            mock_run.return_value = MagicMock(returncode=1, stderr="error msg")
            with pytest.raises(RuntimeError, match="conversion failed"):
                doc.convert_to_pdf(pdf_out)

    def test_raises_if_no_document_open(self, tmp_path: Path):
        """convert_to_pdf raises RuntimeError when no document is open."""
        doc = DocxDocument.__new__(DocxDocument)
        doc.workdir = None
        with pytest.raises(RuntimeError, match="No document"):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))
