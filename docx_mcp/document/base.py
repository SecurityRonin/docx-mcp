"""Base mixin: lifecycle, XML cache, namespace constants, shared helpers."""

from __future__ import annotations

import contextlib
import os
import random
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from .errors import DocxMcpError, ErrCode
from .guards import InputGuard

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
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
DC = "{http://purl.org/dc/elements/1.1/}"
DCTERMS = "{http://purl.org/dc/terms/}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"

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


# ── Table grid geometry ─────────────────────────────────────────────────────

# Magic bytes of an OLE2/CFB compound file. MS-OFFCRYPTO wraps an encrypted
# OPC package in one of these, so a password-protected .docx is not a ZIP.
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _w_val_int(el: etree._Element | None, default: int) -> int:
    """Read a w:val attribute as an int, falling back to the schema default.

    An out-of-range or non-numeric w:val is ignored rather than fatal, which
    is how Word itself resolves one — see ECMA-376 Part 1 §17.4 on the
    ST_DecimalNumber attributes below.
    """
    if el is None:
        return default
    raw = el.get(f"{W}val")
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def row_grid_width(tr: etree._Element) -> int:
    """Width of a table row measured in grid columns, not in <w:tc> elements.

    A row spans `gridBefore + Σ gridSpan + gridAfter` columns of w:tblGrid
    (ECMA-376 Part 1 §17.4 — w:gridSpan, w:gridBefore, w:gridAfter). Counting
    cells instead reports a horizontally merged row as narrower than it is,
    and a ragged-edge row as narrower still.
    """
    width = 0
    tr_pr = tr.find(f"{W}trPr")
    if tr_pr is not None:
        width += _w_val_int(tr_pr.find(f"{W}gridBefore"), 0)
        width += _w_val_int(tr_pr.find(f"{W}gridAfter"), 0)
    for tc in tr.findall(f"{W}tc"):
        tc_pr = tc.find(f"{W}tcPr")
        span = tc_pr.find(f"{W}gridSpan") if tc_pr is not None else None
        width += _w_val_int(span, 1)
    return width


def table_row_widths(tbl: etree._Element) -> list[int]:
    """Grid width of every row in a table, in document order."""
    return [row_grid_width(tr) for tr in tbl.findall(f"{W}tr")]


class BaseMixin:
    """Lifecycle, XML cache, and shared helpers."""

    _MAX_ZIP_ENTRIES = 10_000
    _MAX_UNCOMPRESSED_BYTES = 500 * 1024 * 1024  # 500 MB

    def __init__(self, path: str):
        self.source_path = Path(path).resolve()
        self.workdir: Path | None = None
        self._trees: dict[str, etree._Element] = {}
        self._modified: set[str] = set()
        self._binaries: dict[str, bytes] = {}  # zip_path -> bytes for in-memory binary updates

    # ── Open / Close ────────────────────────────────────────────────────────

    def open(self) -> dict:
        """Unpack DOCX and parse XML files. Returns document info."""
        if not self.source_path.exists():
            raise FileNotFoundError(f"File not found: {self.source_path}")
        if self.source_path.suffix.lower() != ".docx":
            raise ValueError(f"Not a .docx file: {self.source_path}")

        # An encrypted .docx is an OLE2/CFB envelope, not a ZIP. Detect it here
        # so the diagnostic names the real condition instead of "bad ZIP".
        with self.source_path.open("rb") as fh:
            magic = fh.read(len(OLE2_MAGIC))
        if magic == OLE2_MAGIC:
            raise DocxMcpError(
                ErrCode.ENCRYPTED_DOCUMENT,
                f"{self.source_path.name} is an OLE2/CFB compound file "
                f"(leading bytes {magic.hex()}), not a ZIP-based .docx. This is "
                "how MS-OFFCRYPTO stores a password-encrypted document.",
                hint="Decrypt the file before opening it (e.g. with msoffcrypto-tool).",
            )

        self.workdir = Path(tempfile.mkdtemp(prefix="docx_mcp_"))
        try:
            zf_handle = zipfile.ZipFile(self.source_path, "r")
        except zipfile.BadZipFile as exc:
            shutil.rmtree(self.workdir, ignore_errors=True)
            self.workdir = None
            raise DocxMcpError(
                ErrCode.OOXML_INVALID,
                f"Corrupt or invalid DOCX (bad ZIP): {exc}",
                hint="The file may be truncated or not a valid .docx file.",
            ) from exc
        with zf_handle:
            infos = zf_handle.infolist()
            # V3: entry count bomb
            if len(infos) > self._MAX_ZIP_ENTRIES:
                shutil.rmtree(self.workdir, ignore_errors=True)
                self.workdir = None
                raise DocxMcpError(
                    ErrCode.OOXML_INVALID,
                    f"DOCX has {len(infos)} ZIP entries (limit {self._MAX_ZIP_ENTRIES})",
                    hint="This may be a ZIP bomb or a corrupted file.",
                )
            # V3: uncompressed size bomb
            total_uncompressed = sum(i.file_size for i in infos)
            if total_uncompressed > self._MAX_UNCOMPRESSED_BYTES:
                shutil.rmtree(self.workdir, ignore_errors=True)
                self.workdir = None
                raise DocxMcpError(
                    ErrCode.OOXML_INVALID,
                    f"DOCX uncompressed size {total_uncompressed} bytes exceeds "
                    f"limit of {self._MAX_UNCOMPRESSED_BYTES} bytes",
                    hint="This may be a ZIP bomb.",
                )
            # V1: ZipSlip — validate every entry before extraction
            for member in zf_handle.namelist():
                dest = (self.workdir / member).resolve()
                if not dest.is_relative_to(self.workdir.resolve()):
                    shutil.rmtree(self.workdir, ignore_errors=True)
                    self.workdir = None
                    raise DocxMcpError(
                        ErrCode.UNSAFE_PATH,
                        f"ZipSlip: entry escapes workdir: {member!r}",
                        hint="The DOCX contains a malicious ZIP entry name.",
                    )
            zf_handle.extractall(self.workdir)

        # Discover and parse XML files
        xml_files = ["[Content_Types].xml"]
        # Core properties
        core = self.workdir / "docProps" / "core.xml"
        if core.exists():
            xml_files.append("docProps/core.xml")
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
                    parser = etree.XMLParser(
                        remove_blank_text=False,
                        resolve_entities=False,
                        no_network=True,
                        huge_tree=False,
                    )
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
        self._binaries.clear()
        self.workdir = None

    # ── Save ────────────────────────────────────────────────────────────────

    def save(self, output_path: str | None = None, *, backup: bool = True) -> dict:
        """Write modified XML back to files and repack into a .docx."""
        if self.workdir is None:
            raise RuntimeError("No document is open")

        if output_path is not None:
            output_path = str(InputGuard.output_path(output_path))

        output = Path(output_path) if output_path else self.source_path

        # Backup existing file before overwriting
        backup_path = self._backup_if_exists(output) if backup else None

        # Auto-repair known corruption patterns
        repairs = self._pre_save_repair()

        # Lightweight validation — warn about issues that can't be auto-repaired
        warnings = self._post_repair_warnings()

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

        # Write in-memory binary overrides (e.g. replaced images)
        for zip_path, data in self._binaries.items():
            fp = self.workdir / zip_path
            fp.parent.mkdir(parents=True, exist_ok=True)
            fp.write_bytes(data)
        self._binaries.clear()

        # Repack
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(self.workdir):
                for fname in files:
                    fpath = Path(root) / fname
                    arcname = str(fpath.relative_to(self.workdir))
                    zf.write(fpath, arcname)

        modified = sorted(self._modified)
        self._modified.clear()
        result = {
            "path": str(output),
            "size_bytes": output.stat().st_size,
            "modified_parts": modified,
            "repairs": repairs,
            "warnings": warnings,
        }
        if backup_path:
            result["backup"] = str(backup_path)
        return result

    @staticmethod
    def _backup_if_exists(output: Path) -> Path | None:
        """Create a backup of output if it already exists.

        Uses .bak, .bak2, .bak3, ... to avoid overwriting previous backups.
        """
        if not output.exists():
            return None
        # Find next available backup name
        bak = output.with_suffix(output.suffix + ".bak")
        if not bak.exists():
            shutil.copy2(output, bak)
            return bak
        n = 2
        while True:
            bak = output.parent / (output.stem + output.suffix + f".bak{n}")
            if not bak.exists():
                break
            n += 1
        shutil.copy2(output, bak)
        return bak

    # ── Pre-save auto-repair ─────────────────────────────────────────────────

    def _pre_save_repair(self) -> dict:
        """Auto-repair known corruption patterns before writing.

        Returns a dict of repair counts (all zero if nothing was fixed).
        """
        repairs: dict[str, int] = {
            "orphan_footnotes_removed": 0,
            "orphan_endnotes_removed": 0,
            "paraids_deduplicated": 0,
            "broken_rels_removed": 0,
        }

        doc = self._tree("word/document.xml")
        if doc is None:
            return repairs

        body = doc.find(f"{W}body")

        # ── Orphan footnotes ──────────────────────────────────────────────
        fn_tree = self._tree("word/footnotes.xml")
        if fn_tree is not None and body is not None:
            ref_ids = {int(ref.get(f"{W}id", "0")) for ref in body.iter(f"{W}footnoteReference")}
            for fn in list(fn_tree.findall(f"{W}footnote")):
                fn_id = int(fn.get(f"{W}id", "0"))
                if fn_id >= 2 and fn_id not in ref_ids:
                    fn_tree.remove(fn)
                    repairs["orphan_footnotes_removed"] += 1
            if repairs["orphan_footnotes_removed"]:
                self._mark("word/footnotes.xml")

        # ── Orphan endnotes ───────────────────────────────────────────────
        en_tree = self._tree("word/endnotes.xml")
        if en_tree is not None and body is not None:
            ref_ids = {int(ref.get(f"{W}id", "0")) for ref in body.iter(f"{W}endnoteReference")}
            for en in list(en_tree.findall(f"{W}endnote")):
                en_id_str = en.get(f"{W}id", "0")
                if en_id_str in ("0", "-1"):
                    continue
                en_id = int(en_id_str)
                if en_id not in ref_ids:
                    en_tree.remove(en)
                    repairs["orphan_endnotes_removed"] += 1
            if repairs["orphan_endnotes_removed"]:
                self._mark("word/endnotes.xml")

        # ── Duplicate paraIds ─────────────────────────────────────────────
        seen: dict[str, bool] = {}
        for rel_path, tree in self._trees.items():
            if not rel_path.endswith(".xml"):
                continue
            for elem in tree.iter():
                pid = elem.get(f"{W14}paraId")
                if pid is None:
                    continue
                if pid in seen:
                    elem.set(f"{W14}paraId", self._new_para_id())
                    repairs["paraids_deduplicated"] += 1
                    self._mark(rel_path)
                else:
                    seen[pid] = True

        # ── Broken internal relationships ─────────────────────────────────
        rels_tree = self._tree("word/_rels/document.xml.rels")
        if rels_tree is not None:
            for rel in list(rels_tree.findall(f"{RELS}Relationship")):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                if target and not (self.workdir / "word" / target).exists():
                    rels_tree.remove(rel)
                    repairs["broken_rels_removed"] += 1
            if repairs["broken_rels_removed"]:
                self._mark("word/_rels/document.xml.rels")

        return repairs

    def _post_repair_warnings(self) -> list[str]:
        """Check for issues that can't be auto-repaired. Returns warning strings."""
        warnings: list[str] = []
        doc = self._tree("word/document.xml")
        if doc is None:
            return warnings

        # Heading level skips
        headings = self._find_headings(doc)
        prev = 0
        for h in headings:
            if h["level"] > prev + 1 and prev > 0:
                warnings.append(
                    f"Heading level skip: H{prev} -> H{h['level']} at '{h['text'][:50]}'"
                )
            prev = h["level"]

        # Unpaired bookmarks
        starts = {e.get(f"{W}id") for e in doc.iter(f"{W}bookmarkStart") if e.get(f"{W}id")}
        ends = {e.get(f"{W}id") for e in doc.iter(f"{W}bookmarkEnd") if e.get(f"{W}id")}
        unpaired = len(starts - ends) + len(ends - starts)
        if unpaired:
            warnings.append(f"{unpaired} unpaired bookmark(s)")

        # Inconsistent table columns (measured in grid columns — see row_grid_width)
        for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
            widths = table_row_widths(tbl)
            if widths and len(set(widths)) > 1:
                warnings.append(f"Table {idx + 1} has inconsistent column counts: {widths}")

        # Artifact markers (DRAFT, TODO, FIXME, XXX)
        for marker in ("DRAFT", "TODO", "FIXME", "XXX"):
            hits = self.search_text(marker)
            for hit in hits:
                warnings.append(f"{marker} marker in {hit['source']}: '{hit['text'][:60]}'")

        return warnings

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
