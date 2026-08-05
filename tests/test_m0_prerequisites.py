"""M0 prerequisites — bug fixes that block later milestones.

Covers:
  * gridSpan/gridBefore-aware table column validation (base.py + validation.py)
  * convert_to_pdf: no source mutation, cross-device move, raise-on-missing,
    output-path guard
  * OLE2 (encrypted / CFB) detection at open_document

See docs/IMPLEMENTATION-PLAN-table-and-deploy.md §M0.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.document.base import W
from docx_mcp.document.guards import InputGuard

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _make_doc(tmp_path: Path, name: str = "test.docx") -> DocxDocument:
    return DocxDocument.create(str(tmp_path / name))


def _append_table(doc: DocxDocument, table_xml: str) -> None:
    """Parse a <w:tbl> literal and append it to the document body."""
    root = doc._require("word/document.xml")
    body = root.find(f"{W}body")
    body.append(etree.fromstring(table_xml))
    doc._mark("word/document.xml")


def _cell(text: str = "x", span: int | None = None) -> str:
    span_xml = f'<w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr>' if span else ""
    return f"<w:tc>{span_xml}<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def _table(rows: list[str], cols: int = 3) -> str:
    grid = "".join('<w:gridCol w:w="2000"/>' for _ in range(cols))
    return f"<w:tbl {WNS}><w:tblPr/><w:tblGrid>{grid}</w:tblGrid>" + "".join(rows) + "</w:tbl>"


def _table_warnings(doc: DocxDocument) -> list[str]:
    return [w for w in doc._post_repair_warnings() if "column count" in w.lower()]


# ── M0.3: gridSpan-aware column validation ──────────────────────────────────


class TestMergedTableNotFlagged:
    """A horizontally merged table is valid; the validator must not warn."""

    def test_gridspan_row_is_not_inconsistent(self, tmp_path: Path):
        # Row 1: three plain cells. Row 2: one cell spanning 2 + one plain.
        # Both rows occupy 3 grid columns.
        doc = _make_doc(tmp_path)
        _append_table(
            doc,
            _table(
                [
                    f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>",
                    f"<w:tr>{_cell('ab', span=2)}{_cell('c')}</w:tr>",
                ]
            ),
        )
        assert _table_warnings(doc) == []

    def test_audit_agrees_with_save_time_warning(self, tmp_path: Path):
        """validation.audit() and _post_repair_warnings() must not disagree."""
        doc = _make_doc(tmp_path)
        _append_table(
            doc,
            _table(
                [
                    f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>",
                    f"<w:tr>{_cell('ab', span=2)}{_cell('c')}</w:tr>",
                ]
            ),
        )
        assert doc.audit()["tables"]["inconsistent_columns"] == []

    def test_gridbefore_and_gridafter_count_toward_width(self, tmp_path: Path):
        """w:gridBefore / w:gridAfter declare phantom columns (ragged-edge tables)."""
        doc = _make_doc(tmp_path)
        row2 = (
            '<w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>'
            f"{_cell('only')}</w:tr>"
        )
        _append_table(
            doc,
            _table([f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>", row2]),
        )
        assert _table_warnings(doc) == []

    def test_vmerge_continuation_rows_not_flagged(self, tmp_path: Path):
        """vMerge continuation cells are physically present — width is unchanged."""
        doc = _make_doc(tmp_path)
        cont = '<w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc>'
        _append_table(
            doc,
            _table(
                [
                    f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>",
                    f"<w:tr>{cont}{_cell('b2')}{_cell('c2')}</w:tr>",
                ]
            ),
        )
        assert _table_warnings(doc) == []


class TestGenuinelyRaggedTableStillFlagged:
    """The fix must not silence the real defect it was written to catch."""

    def test_short_row_without_span_still_warns(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        _append_table(
            doc,
            _table(
                [
                    f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>",
                    f"<w:tr>{_cell('a')}{_cell('b')}</w:tr>",
                ]
            ),
        )
        assert len(_table_warnings(doc)) == 1

    def test_audit_reports_the_effective_widths(self, tmp_path: Path):
        doc = _make_doc(tmp_path)
        _append_table(
            doc,
            _table(
                [
                    f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>",
                    f"<w:tr>{_cell('a')}{_cell('b')}</w:tr>",
                ]
            ),
        )
        issues = doc.audit()["tables"]["inconsistent_columns"]
        assert len(issues) == 1
        assert issues[0]["column_counts"] == [3, 2]

    def test_oversized_span_row_still_warns(self, tmp_path: Path):
        """A gridSpan that overshoots the grid is a genuine defect."""
        doc = _make_doc(tmp_path)
        _append_table(
            doc,
            _table(
                [
                    f"<w:tr>{_cell('a')}{_cell('b')}{_cell('c')}</w:tr>",
                    f"<w:tr>{_cell('wide', span=4)}</w:tr>",
                ]
            ),
        )
        assert len(_table_warnings(doc)) == 1


# ── M0.2: convert_to_pdf side effects and guards ────────────────────────────


def _fake_soffice(pdf_name: str):
    """subprocess.run stand-in that writes the PDF LibreOffice would produce."""

    def _run(cmd, **kwargs):
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        outdir.mkdir(parents=True, exist_ok=True)
        (outdir / pdf_name).write_bytes(b"%PDF-1.4\n%fake\n")
        return MagicMock(returncode=0, stderr="", stdout="")

    return _run


class TestConvertToPdfLeavesSourceAlone:
    def test_source_bytes_unchanged(self, tmp_path: Path):
        """convert_to_pdf must not rewrite the file it was opened from."""
        doc = _make_doc(tmp_path, "src.docx")
        doc.save()
        before = (tmp_path / "src.docx").read_bytes()

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", side_effect=_fake_soffice("src.pdf")),
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

        assert (tmp_path / "src.docx").read_bytes() == before

    def test_source_mtime_unchanged(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")
        doc.save()
        before = (tmp_path / "src.docx").stat().st_mtime_ns

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", side_effect=_fake_soffice("src.pdf")),
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

        assert (tmp_path / "src.docx").stat().st_mtime_ns == before

    def test_no_backup_file_created_next_to_source(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")
        doc.save()
        # save() itself backs up the file create() wrote, so snapshot first —
        # what matters is that convert_to_pdf adds nothing.
        before = sorted(tmp_path.glob("*.bak*"))

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", side_effect=_fake_soffice("src.pdf")),
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

        assert sorted(tmp_path.glob("*.bak*")) == before

    def test_pending_edits_are_included_in_the_pdf_input(self, tmp_path: Path):
        """The converted copy must reflect unsaved in-memory edits."""
        doc = _make_doc(tmp_path, "src.docx")
        _append_table(doc, _table([f"<w:tr>{_cell('marker')}{_cell('b')}{_cell('c')}</w:tr>"]))

        seen: dict[str, bytes] = {}

        def _capture(cmd, **kwargs):
            src = Path(cmd[-1])
            with zipfile.ZipFile(src) as zf:
                seen["document"] = zf.read("word/document.xml")
            outdir = Path(cmd[cmd.index("--outdir") + 1])
            (outdir / (src.stem + ".pdf")).write_bytes(b"%PDF-1.4\n")
            return MagicMock(returncode=0, stderr="")

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", side_effect=_capture),
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

        assert b"marker" in seen["document"]


class TestConvertToPdfFailsLoud:
    def test_raises_when_libreoffice_produces_nothing(self, tmp_path: Path):
        """Exit 0 with no output file is a failure, not a success."""
        doc = _make_doc(tmp_path, "src.docx")

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", return_value=MagicMock(returncode=0, stderr="")),
            pytest.raises(RuntimeError, match="no PDF"),
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

    def test_error_names_the_path_it_looked_for(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", return_value=MagicMock(returncode=0, stderr="")),
            pytest.raises(RuntimeError, match=r"src\.pdf"),
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

    def test_output_lands_at_the_requested_path(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")
        target = tmp_path / "nested" / "renamed.pdf"

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", side_effect=_fake_soffice("src.pdf")),
        ):
            result = doc.convert_to_pdf(str(target))

        assert target.exists()
        assert result["pdf_path"] == str(target)

    def test_uses_shutil_move_for_cross_device_safety(self, tmp_path: Path):
        """The temp-dir copy makes the move cross-device; Path.rename cannot."""
        doc = _make_doc(tmp_path, "src.docx")

        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", side_effect=_fake_soffice("src.pdf")),
            patch("shutil.move") as mv,
        ):
            doc.convert_to_pdf(str(tmp_path / "out.pdf"))

        assert mv.called


class TestConvertToPdfOutputGuard:
    def test_rejects_path_traversal(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")
        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            pytest.raises(ValueError, match="traversal"),
        ):
            doc.convert_to_pdf(str(tmp_path / ".." / "escaped.pdf"))

    def test_rejects_non_pdf_suffix(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")
        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            pytest.raises(ValueError, match="suffix"),
        ):
            doc.convert_to_pdf(str(tmp_path / "payload.sh"))

    def test_guard_runs_before_libreoffice_is_invoked(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "src.docx")
        with (
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run") as run,
            pytest.raises(ValueError),
        ):
            doc.convert_to_pdf(str(tmp_path / "payload.sh"))
        assert not run.called


class TestOutputPathGuardSuffix:
    def test_default_suffix_is_docx(self, tmp_path: Path):
        with pytest.raises(ValueError, match="suffix"):
            InputGuard.output_path(str(tmp_path / "a.pdf"))

    def test_suffix_is_parameterised(self, tmp_path: Path):
        got = InputGuard.output_path(str(tmp_path / "a.pdf"), suffix=".pdf")
        assert got.name == "a.pdf"

    def test_suffix_match_is_case_insensitive(self, tmp_path: Path):
        got = InputGuard.output_path(str(tmp_path / "a.PDF"), suffix=".pdf")
        assert got.name == "a.PDF"

    def test_traversal_rejected_for_any_suffix(self, tmp_path: Path):
        with pytest.raises(ValueError, match="traversal"):
            InputGuard.output_path(str(tmp_path / ".." / "a.pdf"), suffix=".pdf")


# ── M0.4: OLE2 / encrypted-container detection ──────────────────────────────

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class TestEncryptedDocumentDetection:
    def test_ole2_file_reports_encryption_not_corruption(self, tmp_path: Path):
        """An MS-OFFCRYPTO container is CFB, not ZIP — say so."""
        enc = tmp_path / "locked.docx"
        enc.write_bytes(OLE2_MAGIC + b"\x00" * 512)

        doc = DocxDocument(str(enc))
        with pytest.raises(Exception) as exc:
            doc.open()
        assert "encrypt" in str(exc.value).lower()

    def test_error_carries_a_dedicated_code(self, tmp_path: Path):
        from docx_mcp.document.errors import DocxMcpError, ErrCode

        enc = tmp_path / "locked.docx"
        enc.write_bytes(OLE2_MAGIC + b"\x00" * 512)

        doc = DocxDocument(str(enc))
        with pytest.raises(DocxMcpError) as exc:
            doc.open()
        assert exc.value.code is ErrCode.ENCRYPTED_DOCUMENT

    def test_error_shows_the_offending_magic_bytes(self, tmp_path: Path):
        """Fail-loud rule: name the actual value that was not recognised."""
        enc = tmp_path / "locked.docx"
        enc.write_bytes(OLE2_MAGIC + b"\x00" * 512)

        doc = DocxDocument(str(enc))
        with pytest.raises(Exception) as exc:
            doc.open()
        assert "d0cf11e0a1b11ae1" in str(exc.value).lower().replace(" ", "")

    def test_no_workdir_leaked_on_rejection(self, tmp_path: Path):
        enc = tmp_path / "locked.docx"
        enc.write_bytes(OLE2_MAGIC + b"\x00" * 512)

        from docx_mcp.document.errors import DocxMcpError

        doc = DocxDocument(str(enc))
        with pytest.raises(DocxMcpError):
            doc.open()
        assert doc.workdir is None

    def test_plain_garbage_still_reports_bad_zip(self, tmp_path: Path):
        """Non-OLE2 corruption keeps its existing, distinct diagnostic."""
        bad = tmp_path / "junk.docx"
        bad.write_bytes(b"not a zip at all" * 8)

        doc = DocxDocument(str(bad))
        with pytest.raises(Exception) as exc:
            doc.open()
        assert "zip" in str(exc.value).lower()

    def test_valid_docx_still_opens(self, tmp_path: Path):
        doc = _make_doc(tmp_path, "fine.docx")
        doc.close()
        reopened = DocxDocument(str(tmp_path / "fine.docx"))
        reopened.open()
        assert reopened.workdir is not None
        reopened.close()
