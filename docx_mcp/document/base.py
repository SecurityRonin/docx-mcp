"""Base mixin: lifecycle, XML cache, namespace constants, shared helpers."""

from __future__ import annotations

import contextlib
import os
import random
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

# ── OOXML namespace constants ───────────────────────────────────────────────
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

REL_TYPES = {
    "comments": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "commentsExtended": "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "footnotes": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
}

CT_TYPES = {
    "comments": ("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
    "commentsExtended": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    ),
}


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _preserve(t_el: etree._Element, text: str) -> None:
    """Set text on a <w:t> or <w:delText> element with xml:space=preserve."""
    t_el.text = text
    t_el.set(XML_SPACE, "preserve")


class BaseMixin:
    """Lifecycle, XML cache, and shared helpers."""

    def __init__(self, path: str):
        self.source_path = Path(path).resolve()
        self.workdir: Path | None = None
        self._trees: dict[str, etree._Element] = {}
        self._modified: set[str] = set()

    # ── Open / Close ────────────────────────────────────────────────────────

    def open(self) -> dict:
        """Unpack DOCX and parse XML files. Returns document info."""
        if not self.source_path.exists():
            raise FileNotFoundError(f"File not found: {self.source_path}")
        if self.source_path.suffix.lower() != ".docx":
            raise ValueError(f"Not a .docx file: {self.source_path}")

        self.workdir = Path(tempfile.mkdtemp(prefix="docx_mcp_"))
        with zipfile.ZipFile(self.source_path, "r") as zf:
            zf.extractall(self.workdir)

        # Discover and parse XML files
        xml_files = ["[Content_Types].xml"]
        word_dir = self.workdir / "word"
        if word_dir.exists():
            for name in [
                "document.xml",
                "footnotes.xml",
                "endnotes.xml",
                "comments.xml",
                "commentsExtended.xml",
                "styles.xml",
                "numbering.xml",
                "settings.xml",
            ]:
                if (word_dir / name).exists():
                    xml_files.append(f"word/{name}")
            # Headers and footers
            for f in word_dir.iterdir():
                if f.name.startswith(("header", "footer")) and f.suffix == ".xml":
                    xml_files.append(f"word/{f.name}")

        # Relationship files
        rels_dir = word_dir / "_rels"
        if rels_dir.exists():
            for f in rels_dir.iterdir():
                if f.suffix == ".rels":
                    xml_files.append(f"word/_rels/{f.name}")

        for rel_path in xml_files:
            full_path = self.workdir / rel_path
            if full_path.exists():
                try:
                    parser = etree.XMLParser(remove_blank_text=False)
                    tree = etree.parse(str(full_path), parser)
                    self._trees[rel_path] = tree.getroot()
                except etree.XMLSyntaxError:
                    pass

        return self.get_info()

    def close(self) -> None:
        """Clean up temporary files."""
        if self.workdir and self.workdir.exists():
            shutil.rmtree(self.workdir, ignore_errors=True)
        self._trees.clear()
        self._modified.clear()
        self.workdir = None

    # ── Save ────────────────────────────────────────────────────────────────

    def save(self, output_path: str | None = None) -> dict:
        """Write modified XML back to files and repack into a .docx."""
        if self.workdir is None:
            raise RuntimeError("No document is open")

        output = Path(output_path) if output_path else self.source_path

        # Serialize modified trees
        for rel_path in self._modified:
            tree = self._trees.get(rel_path)
            if tree is None:
                continue
            fp = self.workdir / rel_path
            fp.parent.mkdir(parents=True, exist_ok=True)
            et = etree.ElementTree(tree)
            et.write(
                str(fp),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        # Repack
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(self.workdir):
                for fname in files:
                    fpath = Path(root) / fname
                    arcname = str(fpath.relative_to(self.workdir))
                    zf.write(fpath, arcname)

        modified = sorted(self._modified)
        self._modified.clear()
        return {
            "path": str(output),
            "size_bytes": output.stat().st_size,
            "modified_parts": modified,
        }

    # ── Private helpers ─────────────────────────────────────────────────────

    def _tree(self, rel_path: str) -> etree._Element | None:
        return self._trees.get(rel_path)

    def _require(self, rel_path: str) -> etree._Element:
        t = self._tree(rel_path)
        if t is None:
            raise RuntimeError(f"{rel_path} not found — is a document open?")
        return t

    def _mark(self, rel_path: str) -> None:
        self._modified.add(rel_path)

    @staticmethod
    def _text(element: etree._Element) -> str:
        """Concatenate all <w:t> text descendants."""
        return "".join(t.text for t in element.iter(f"{W}t") if t.text)

    @staticmethod
    def _real_footnotes(fn_root: etree._Element) -> list[etree._Element]:
        """Return footnote elements excluding separators (id 0 and -1)."""
        return [f for f in fn_root.findall(f"{W}footnote") if f.get(f"{W}id") not in ("0", "-1")]

    def _find_para(self, root: etree._Element, para_id: str) -> etree._Element | None:
        for p in root.iter(f"{W}p"):
            if p.get(f"{W14}paraId") == para_id:
                return p
        return None

    def _new_para_id(self) -> str:
        """Generate a unique paraId (8 hex digits, < 0x80000000)."""
        existing: set[str] = set()
        for tree in self._trees.values():
            for el in tree.iter():
                pid = el.get(f"{W14}paraId")
                if pid:
                    existing.add(pid.upper())
        while True:
            val = random.randint(1, 0x7FFFFFFF)
            pid = f"{val:08X}"
            if pid not in existing:
                return pid

    @staticmethod
    def _next_markup_id(doc: etree._Element) -> int:
        """Next available ID for ins/del/comment/bookmark markup."""
        max_id = 0
        for tag in (
            f"{W}ins",
            f"{W}del",
            f"{W}commentRangeStart",
            f"{W}commentRangeEnd",
            f"{W}bookmarkStart",
            f"{W}bookmarkEnd",
        ):
            for el in doc.iter(tag):
                eid = el.get(f"{W}id")
                if eid:
                    with contextlib.suppress(ValueError):
                        max_id = max(max_id, int(eid))
        return max_id + 1

    @staticmethod
    def _next_comment_id(cm_tree: etree._Element) -> int:
        max_id = -1
        for c in cm_tree.findall(f"{W}comment"):
            cid = c.get(f"{W}id")
            if cid:
                with contextlib.suppress(ValueError):
                    max_id = max(max_id, int(cid))
        return max_id + 1

    @staticmethod
    def _make_run(text: str, rpr_bytes: bytes | None) -> etree._Element:
        """Build a <w:r> element with optional copied rPr."""
        r = etree.Element(f"{W}r")
        if rpr_bytes:
            r.append(etree.fromstring(rpr_bytes))
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)
        return r

    def _create_comments_part(self) -> etree._Element:
        """Create word/comments.xml and register it in rels + content types."""
        root = etree.Element(
            f"{W}comments",
            nsmap={"w": NSMAP["w"], "w14": NSMAP["w14"], "r": NSMAP["r"]},
        )
        self._trees["word/comments.xml"] = root

        # Write file so it exists on disk
        fp = self.workdir / "word" / "comments.xml"
        fp.parent.mkdir(parents=True, exist_ok=True)
        etree.ElementTree(root).write(str(fp), xml_declaration=True, encoding="UTF-8")

        # Content type
        ct = self._tree("[Content_Types].xml")
        if ct is not None:
            existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
            if "/word/comments.xml" not in existing:
                ov = etree.SubElement(ct, f"{CT}Override")
                ov.set("PartName", "/word/comments.xml")
                ov.set("ContentType", CT_TYPES["comments"])
                self._mark("[Content_Types].xml")

        # Relationship
        rels = self._tree("word/_rels/document.xml.rels")
        if rels is not None:
            existing_targets = {r.get("Target") for r in rels.findall(f"{RELS}Relationship")}
            if "comments.xml" not in existing_targets:
                max_rid = 0
                for r in rels.findall(f"{RELS}Relationship"):
                    rid = r.get("Id", "")
                    if rid.startswith("rId"):
                        with contextlib.suppress(ValueError):
                            max_rid = max(max_rid, int(rid[3:]))
                rel = etree.SubElement(rels, f"{RELS}Relationship")
                rel.set("Id", f"rId{max_rid + 1}")
                rel.set("Type", REL_TYPES["comments"])
                rel.set("Target", "comments.xml")
                self._mark("word/_rels/document.xml.rels")

        self._mark("word/comments.xml")
        return root
