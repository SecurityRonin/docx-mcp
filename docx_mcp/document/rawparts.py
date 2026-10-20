"""Raw XML part access — power-user escape hatch."""

from __future__ import annotations

import os
from pathlib import Path

from lxml import etree

from .errors import DocxMcpError, ErrCode


class RawPartsMixin:
    def list_parts(self) -> list[str]:
        """Return sorted list of all part paths in the DOCX zip."""
        parts: list[str] = []
        workdir: Path = self.workdir  # type: ignore[attr-defined]
        for root, _dirs, files in os.walk(workdir):
            for fname in files:
                fpath = Path(root) / fname
                # OPC part names always use "/", whatever the OS path separator is.
                arcname = fpath.relative_to(workdir).as_posix()
                parts.append(arcname)
        return sorted(parts)

    def read_part(self, part_path: str) -> dict:
        """Return raw pretty-printed XML of a part.

        Returns: {"part": str, "xml": str}
        Raises: DocxMcpError(ErrCode.PART_NOT_FOUND) if part_path not in zip.
        """
        tree = self._tree(part_path)  # type: ignore[attr-defined]
        if tree is None:
            raise DocxMcpError(
                ErrCode.PART_NOT_FOUND,
                f"Part not found: {part_path}",
                hint="Use list_parts() to see available parts.",
            )
        return {
            "part": part_path,
            "xml": etree.tostring(tree, pretty_print=True).decode(),
        }

    def write_part(self, part_path: str, xml: str) -> dict:
        """Replace a part's XML. Validates well-formedness first.

        Returns: {"part": str, "bytes_written": int}
        Raises: DocxMcpError(ErrCode.OOXML_INVALID) if xml is not well-formed.
        """
        try:
            new_tree = etree.fromstring(xml.encode())
        except etree.XMLSyntaxError as e:
            raise DocxMcpError(ErrCode.OOXML_INVALID, f"Invalid XML: {e}") from e

        self._trees[part_path] = new_tree  # type: ignore[attr-defined]
        self._mark(part_path)  # type: ignore[attr-defined]
        return {
            "part": part_path,
            "bytes_written": len(etree.tostring(new_tree)),
        }
