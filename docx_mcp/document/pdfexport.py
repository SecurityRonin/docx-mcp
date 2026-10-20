"""PDF export mixin: convert the open document to PDF via LibreOffice headless."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


class PdfExportMixin:
    def convert_to_pdf(self, output_path: str) -> dict:
        """Convert the current document to PDF using LibreOffice headless.

        Saves the document first, then invokes:
          libreoffice --headless --convert-to pdf --outdir <dir> <docx>

        Args:
            output_path: Desired path for the output PDF file.

        Returns:
            {"pdf_path": str}

        Raises:
            RuntimeError: If no document is open, LibreOffice is not found,
                          the conversion process exits with a non-zero code,
                          or it does not finish within the timeout.
        """
        if self.workdir is None:
            raise RuntimeError("No document is open.")

        lo = shutil.which("libreoffice") or shutil.which("soffice")
        if lo is None:
            raise RuntimeError(
                "LibreOffice not found. Install it and ensure 'libreoffice' or "
                "'soffice' is on PATH."
            )

        self.save(self.source_path, backup=False)

        out = Path(output_path)
        outdir = out.parent
        outdir.mkdir(parents=True, exist_ok=True)

        # stdin=DEVNULL: never let LibreOffice inherit the MCP server's stdio pipe;
        # on Windows soffice blocked on it until the client's request timed out.
        # timeout stays under the client's 60s tool-call limit so a hang
        # surfaces as a clear error instead of a dropped request.
        try:
            result = subprocess.run(
                [
                    lo,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(outdir),
                    str(self.source_path),
                ],
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                timeout=50,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(f"LibreOffice conversion timed out after {exc.timeout}s.") from exc
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}): {result.stderr.strip()}"
            )

        # LibreOffice names the output after the input stem
        generated = outdir / (Path(self.source_path).stem + ".pdf")
        if generated != out and generated.exists():
            generated.rename(out)

        return {"pdf_path": str(out)}
