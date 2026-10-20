"""Metadata sanitization mixin: strip identifying information from DOCX."""

from __future__ import annotations

import copy
import os
import zipfile
from pathlib import Path

from lxml import etree

from .base import W

# rsid attributes that expose revision-session fingerprints
_RSID_ATTRS = {
    f"{W}rsidR",
    f"{W}rsidRPr",
    f"{W}rsidDel",
    f"{W}rsidRDefault",
    f"{W}rsidSect",
    f"{W}rsidTr",
}

# Namespace URIs
_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_DC = "http://purl.org/dc/elements/1.1/"
_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class MetadataMixin:
    """Strip identifying metadata from DOCX files."""

    def sanitize_metadata(
        self,
        output_path: str,
        *,
        level: int = 1,
        redact_authors_as: str = "",
    ) -> dict:
        """Write a sanitized copy of the open document to *output_path*.

        Level 1: Remove all rsid attributes from document.xml.
        Level 2: + Replace w:author on w:ins / w:del with *redact_authors_as*
                   (defaults to "Anonymous" when empty string).
        Level 3: + Clear creator / lastModifiedBy / revision in core.xml
                 + Clear Company in app.xml
                 + Remove attachedTemplate / rsids from settings.xml

        Returns {"path": output_path, "level": level}.
        Raises ValueError if output_path is empty.
        """
        if not output_path:
            raise ValueError("output_path must be a non-empty string")
        if self.workdir is None:
            raise RuntimeError("No document is open")

        author_label = redact_authors_as or "Anonymous"

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for root_dir, _dirs, files in os.walk(self.workdir):
                for fname in sorted(files):
                    fpath = Path(root_dir) / fname
                    # "/" separators so the part-name checks below match on Windows too.
                    arcname = fpath.relative_to(self.workdir).as_posix()

                    # Use in-memory parsed tree when available (authoritative source)
                    if arcname in self._trees:
                        el = copy.deepcopy(self._trees[arcname])
                        raw = None
                    else:
                        el = None
                        raw = fpath.read_bytes()

                    if arcname == "word/document.xml" and el is not None:
                        self._sanitize_document_el(el, level, author_label)
                        data = etree.tostring(
                            el, xml_declaration=True, encoding="UTF-8", standalone=True
                        )
                    elif arcname == "word/settings.xml" and level >= 3 and el is not None:
                        self._sanitize_settings_el(el)
                        data = etree.tostring(
                            el, xml_declaration=True, encoding="UTF-8", standalone=True
                        )
                    elif arcname == "docProps/core.xml" and level >= 3 and el is not None:
                        self._sanitize_core_el(el)
                        data = etree.tostring(
                            el, xml_declaration=True, encoding="UTF-8", standalone=True
                        )
                    elif arcname == "docProps/app.xml" and level >= 3:
                        # app.xml is not cached in _trees — parse from disk bytes
                        app_el = etree.fromstring(raw)
                        self._sanitize_app_el(app_el)
                        data = etree.tostring(
                            app_el, xml_declaration=True, encoding="UTF-8", standalone=True
                        )
                    elif el is not None:
                        data = etree.tostring(
                            el, xml_declaration=True, encoding="UTF-8", standalone=True
                        )
                    else:
                        data = raw

                    out_zip.writestr(arcname, data)

        return {"path": output_path, "level": level}

    # ── per-part helpers (mutate element in-place) ───────────────────────────

    @staticmethod
    def _sanitize_document_el(root: etree._Element, level: int, author_label: str) -> None:
        # Level 1: strip rsid attributes
        for el in root.iter():
            for attr in list(el.attrib):
                if attr in _RSID_ATTRS:
                    del el.attrib[attr]

        # Level 2+: anonymize tracked-change authors
        if level >= 2:
            for el in root.iter(
                f"{{{_W}}}ins",
                f"{{{_W}}}del",
                f"{{{_W}}}rPrChange",
                f"{{{_W}}}pPrChange",
            ):
                author_attr = f"{{{_W}}}author"
                if author_attr in el.attrib:
                    el.attrib[author_attr] = author_label

    @staticmethod
    def _sanitize_settings_el(root: etree._Element) -> None:
        for el in list(root):
            if el.tag in (f"{{{_W}}}attachedTemplate", f"{{{_W}}}rsids"):
                root.remove(el)

    @staticmethod
    def _sanitize_core_el(root: etree._Element) -> None:
        _clear_tags = {
            f"{{{_DC}}}creator",
            f"{{{_DC}}}description",
            f"{{{_DC}}}subject",
            f"{{{_DC}}}title",
            f"{{{_CP}}}lastModifiedBy",
            f"{{{_CP}}}revision",
            f"{{{_CP}}}keywords",
            f"{{{_CP}}}category",
        }
        for el in root:
            if el.tag in _clear_tags:
                el.text = ""

    @staticmethod
    def _sanitize_app_el(root: etree._Element) -> None:
        _clear_tags = {
            f"{{{_EP}}}Company",
            f"{{{_EP}}}Manager",
        }
        for el in root:
            if el.tag in _clear_tags:
                el.text = ""
