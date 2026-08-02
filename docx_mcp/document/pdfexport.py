"""PDF export mixin: convert the open document to PDF via LibreOffice headless."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from .guards import InputGuard


class PdfExportMixin:
    def convert_to_pdf(self, output_path: str) -> dict:
        """Convert the current document to PDF using LibreOffice headless.

        The document — including any unsaved in-memory edits — is written to a
        private temporary directory and *that copy* is converted, so the file
        this document was opened from is never rewritten and no .bak appears
        beside it. Export is therefore a read operation on the source.

        Args:
            output_path: Desired path for the output PDF file. Must end in
                .pdf and must not contain path-traversal segments.

        Returns:
            {"pdf_path": str}

        Raises:
            ValueError: If output_path fails the output guard.
            RuntimeError: If no document is open, LibreOffice is not found,
                          the conversion exits non-zero, or it exits zero
                          without producing a PDF.
        """
        if self.workdir is None:
            raise RuntimeError("No document is open.")

        # Guard before doing any work — a "readonly" server must not gain an
        # unguarded filesystem write through the export path.
        out = InputGuard.output_path(output_path, suffix=".pdf")

        lo = shutil.which("libreoffice") or shutil.which("soffice")
        if lo is None:
            raise RuntimeError(
                "LibreOffice not found. Install it and ensure 'libreoffice' or "
                "'soffice' is on PATH."
            )

        staging = Path(tempfile.mkdtemp(prefix="docx_mcp_pdf_"))
        try:
            staged = staging / (self.source_path.stem + ".docx")
            self.save(str(staged), backup=False)

            result = subprocess.run(
                [
                    lo,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(staging),
                    str(staged),
                ],
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                raise RuntimeError(
                    f"LibreOffice conversion failed (exit {result.returncode}): "
                    f"{result.stderr.strip()}"
                )

            # LibreOffice names the output after the input stem.
            generated = staging / (staged.stem + ".pdf")
            if not generated.exists():
                raise RuntimeError(
                    f"LibreOffice exited 0 but produced no PDF at {generated}. "
                    f"stderr: {result.stderr.strip() or '(empty)'}"
                )

            out.parent.mkdir(parents=True, exist_ok=True)
            # shutil.move, not Path.rename: staging is a temp dir that may sit
            # on a different device than the destination (EXDEV).
            shutil.move(str(generated), str(out))
        finally:
            shutil.rmtree(staging, ignore_errors=True)

        return {"pdf_path": str(out)}
