"""Reading mixin: headings, search, paragraph access, document info."""

from __future__ import annotations

import contextlib
import re

from lxml import etree

from .base import W, W14, A, R, RELS


class ReadingMixin:
    """Read-only document inspection methods."""

    def get_info(self) -> dict:
        """Get document overview stats."""
        info: dict = {"path": str(self.source_path)}
        with contextlib.suppress(OSError):
            info["size_bytes"] = self.source_path.stat().st_size

        doc = self._tree("word/document.xml")
        if doc is not None:
            body = doc.find(f"{W}body")
            paras = list((body if body is not None else doc).iter(f"{W}p"))
            info["paragraph_count"] = len(paras)
            info["heading_count"] = len(self._find_headings(doc))
            info["image_count"] = len(list(doc.iter(f"{W}drawing")))

        fn = self._tree("word/footnotes.xml")
        if fn is not None:
            info["footnote_count"] = len(self._real_footnotes(fn))

        cm = self._tree("word/comments.xml")
        if cm is not None:
            info["comment_count"] = len(cm.findall(f"{W}comment"))

        info["parts"] = sorted(self._trees.keys())
        return info

    def get_headings(self) -> list[dict]:
        doc = self._require("word/document.xml")
        return self._find_headings(doc)

    def _find_headings(self, root: etree._Element) -> list[dict]:
        headings = []
        for para in root.iter(f"{W}p"):
            ppr = para.find(f"{W}pPr")
            if ppr is None:
                continue
            pstyle = ppr.find(f"{W}pStyle")
            if pstyle is None:
                continue
            style = pstyle.get(f"{W}val", "")
            m = re.match(r"^Heading(\d+)$", style)
            if not m:
                continue
            headings.append(
                {
                    "level": int(m.group(1)),
                    "text": self._text(para),
                    "style": style,
                    "paraId": para.get(f"{W14}paraId", ""),
                }
            )
        return headings

    def search_text(self, query: str, *, regex: bool = False) -> list[dict]:
        """Search for text across document body, footnotes, and comments."""
        results = []
        targets = [
            ("document", "word/document.xml"),
            ("footnotes", "word/footnotes.xml"),
            ("comments", "word/comments.xml"),
        ]
        for source, rel_path in targets:
            tree = self._tree(rel_path)
            if tree is None:
                continue
            for para in tree.iter(f"{W}p"):
                text = self._text(para)
                if not text:
                    continue
                if regex:
                    matches = list(re.finditer(query, text))
                    if not matches:
                        continue
                    match_info = [
                        {"start": m.start(), "end": m.end(), "match": m.group()} for m in matches
                    ]
                else:
                    if query.lower() not in text.lower():
                        continue
                    match_info = None
                results.append(
                    {
                        "source": source,
                        "paraId": para.get(f"{W14}paraId", ""),
                        "text": text[:300],
                        "matches": match_info,
                    }
                )
        return results

    def get_paragraph(self, para_id: str) -> dict:
        """Get full text and metadata for a paragraph by paraId."""
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        ppr = para.find(f"{W}pPr")
        style = ""
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            if ps is not None:
                style = ps.get(f"{W}val", "")
        return {
            "paraId": para_id,
            "style": style,
            "text": self._text(para),
        }
