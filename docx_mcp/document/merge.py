"""Merge mixin: combine content from another DOCX into the current document."""

from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from .base import W14, W


class MergeMixin:
    """Document merge operations."""

    def merge_documents(self, source_path: str) -> dict:
        """Merge another DOCX document's body content into this document.

        Appends all body paragraphs and tables from the source document
        into the current document. ParaIds are remapped to avoid collisions.

        Args:
            source_path: Path to the DOCX file to merge in.
        """
        src = Path(source_path)
        if not src.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")

        # Collect existing paraIds
        existing_ids: set[str] = set()
        for tree in self._trees.values():
            for el in tree.iter():
                pid = el.get(f"{W14}paraId")
                if pid:
                    existing_ids.add(pid.upper())

        # Parse source document
        tmpdir = Path(tempfile.mkdtemp(prefix="docx_merge_"))
        try:
            with zipfile.ZipFile(src, "r") as zf:
                zf.extractall(tmpdir)

            src_doc_path = tmpdir / "word" / "document.xml"
            if not src_doc_path.exists():
                raise ValueError("Source document has no word/document.xml")

            parser = etree.XMLParser(remove_blank_text=False)
            src_tree = etree.parse(str(src_doc_path), parser).getroot()
            src_body = src_tree.find(f"{W}body")
            if src_body is None:
                return {"paragraphs_added": 0}

            # Remap paraIds in source elements
            for el in src_body.iter():
                pid = el.get(f"{W14}paraId")
                if pid and pid.upper() in existing_ids:
                    new_pid = self._new_para_id()
                    el.set(f"{W14}paraId", new_pid)
                    existing_ids.add(new_pid)

            # Append source body children (skip sectPr)
            added = 0
            for child in list(src_body):
                if child.tag == f"{W}sectPr":
                    continue
                body.append(child)
                added += 1

            if added > 0:
                self._mark("word/document.xml")

        finally:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)

        return {"paragraphs_added": added}
