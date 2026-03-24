"""Markdown to DOCX converter using mistune parser."""

from __future__ import annotations

import contextlib
from pathlib import Path

import mistune
from lxml import etree
from mistune.plugins.footnotes import footnotes
from mistune.plugins.formatting import strikethrough
from mistune.plugins.table import table
from mistune.plugins.task_lists import task_lists

from docx_mcp.document.base import CT, RELS, W14, R, W, _preserve
from docx_mcp.typography import smartify


class MarkdownConverter:
    """Convert markdown to OOXML elements in a DocxDocument."""

    @classmethod
    def convert(
        cls,
        doc: object,
        text: str,
        *,
        base_dir: Path | None = None,
    ) -> None:
        """Parse markdown and populate the document body.

        Args:
            doc: A DocxDocument instance (opened via create()).
            text: Markdown text to convert.
            base_dir: Base directory for resolving relative image paths.
        """
        converter = cls(doc, base_dir=base_dir)
        converter._run(text)

    def __init__(self, doc: object, *, base_dir: Path | None = None):
        self._doc = doc
        self._base_dir = base_dir or Path.cwd()
        self._body = doc._trees["word/document.xml"].find(f"{W}body")
        self._footnote_map: dict[str, int] = {}  # markdown key -> footnote id

    def _run(self, text: str) -> None:
        """Parse and render."""
        # Remove existing body content (paragraphs + tables) from skeleton/template
        for p in list(self._body.findall(f"{W}p")):
            self._body.remove(p)
        for tbl in list(self._body.findall(f"{W}tbl")):
            self._body.remove(tbl)

        if not text.strip():
            return

        # Get AST from mistune
        md_ast = mistune.create_markdown(
            renderer=None,
            plugins=[table, footnotes, strikethrough, task_lists],
        )
        tokens = md_ast(text)

        # First pass: collect footnote definitions
        for token in tokens:
            if token["type"] == "footnotes":
                self._process_footnote_definitions(token)

        # Second pass: render body tokens
        sect_pr = self._body.find(f"{W}sectPr")
        for token in tokens:
            if token["type"] in ("footnotes", "blank_line"):
                continue  # Already processed or ignorable
            elements = self._render_block(token)
            for el in elements:
                if sect_pr is not None:
                    sect_pr.addprevious(el)
                else:
                    self._body.append(el)

        # Post-processing: remove orphaned footnote definitions
        self._strip_orphan_footnotes()

        self._doc._mark("word/document.xml")

    def _strip_orphan_footnotes(self) -> None:
        """Remove footnote definitions that have no reference in the body.

        Defends against orphaned footnotes that trigger Word's recovery
        dialog, regardless of parser behavior.
        """
        fn_tree = self._doc._tree("word/footnotes.xml")
        if fn_tree is None:
            return

        # Collect all footnoteReference IDs from the body
        ref_ids = {int(ref.get(f"{W}id", "0")) for ref in self._body.iter(f"{W}footnoteReference")}

        # Remove definitions (skip separators 0 and 1) with no body reference
        for fn in list(fn_tree.findall(f"{W}footnote")):
            fn_id = int(fn.get(f"{W}id", "0"))
            if fn_id >= 2 and fn_id not in ref_ids:
                fn_tree.remove(fn)
                self._doc._mark("word/footnotes.xml")

    def _new_para(self, style: str | None = None) -> etree._Element:
        """Create a new <w:p> with paraId and optional style."""
        p = etree.Element(f"{W}p")
        p.set(f"{W14}paraId", self._doc._new_para_id())
        p.set(f"{W14}textId", "77777777")
        if style:
            ppr = etree.SubElement(p, f"{W}pPr")
            ps = etree.SubElement(ppr, f"{W}pStyle")
            ps.set(f"{W}val", style)
        return p

    def _make_run(
        self,
        text: str,
        *,
        bold: bool = False,
        italic: bool = False,
        strike: bool = False,
        code: bool = False,
        smart: bool = True,
    ) -> etree._Element:
        """Build a <w:r> with optional formatting."""
        r = etree.Element(f"{W}r")
        if bold or italic or strike or code:
            rpr = etree.SubElement(r, f"{W}rPr")
            if bold:
                etree.SubElement(rpr, f"{W}b")
            if italic:
                etree.SubElement(rpr, f"{W}i")
            if strike:
                etree.SubElement(rpr, f"{W}strike")
            if code:
                fonts = etree.SubElement(rpr, f"{W}rFonts")
                fonts.set(f"{W}ascii", "Courier New")
                fonts.set(f"{W}hAnsi", "Courier New")
        t = etree.SubElement(r, f"{W}t")
        final_text = smartify(text) if smart and not code else text
        _preserve(t, final_text)
        return r

    def _render_block(self, token: dict) -> list[etree._Element]:
        """Render a block-level token into OOXML elements."""
        t = token["type"]
        if t == "paragraph":
            return [self._render_paragraph(token)]
        elif t == "heading":
            return [self._render_heading(token)]
        elif t == "block_code":
            return self._render_code_block(token)
        elif t == "list":
            return self._render_list(token)
        elif t == "block_quote":
            return self._render_blockquote(token)
        elif t == "thematic_break":
            return [self._render_hr()]
        elif t == "table":
            return [self._render_table(token)]
        return []

    def _render_paragraph(self, token: dict) -> etree._Element:
        p = self._new_para()
        self._render_inline_children(p, token.get("children", []))
        return p

    def _render_heading(self, token: dict) -> etree._Element:
        level = token["attrs"]["level"]
        p = self._new_para(f"Heading{level}")
        self._render_inline_children(p, token.get("children", []))
        return p

    def _render_inline_children(
        self,
        parent: etree._Element,
        children: list[dict],
        *,
        bold: bool = False,
        italic: bool = False,
        strike: bool = False,
    ) -> None:
        """Recursively render inline tokens as runs."""
        for child in children:
            ct = child["type"]
            if ct == "text":
                parent.append(self._make_run(child["raw"], bold=bold, italic=italic, strike=strike))
            elif ct == "strong":
                self._render_inline_children(
                    parent, child["children"], bold=True, italic=italic, strike=strike
                )
            elif ct == "emphasis":
                self._render_inline_children(
                    parent, child["children"], bold=bold, italic=True, strike=strike
                )
            elif ct == "strikethrough":
                self._render_inline_children(
                    parent, child["children"], bold=bold, italic=italic, strike=True
                )
            elif ct == "codespan":
                parent.append(self._make_run(child["raw"], code=True, smart=False))
            elif ct == "link":
                self._render_link(parent, child)
            elif ct == "image":
                self._render_image(parent, child)
            elif ct == "softbreak":
                parent.append(self._make_run(" ", bold=bold, italic=italic, strike=strike))
            elif ct == "linebreak":
                r = etree.SubElement(parent, f"{W}r")
                etree.SubElement(r, f"{W}br")
            elif ct == "footnote_ref":
                self._render_footnote_ref(parent, child)

    def _render_code_block(self, token: dict) -> list[etree._Element]:
        """Render a fenced code block as CodeBlock-styled paragraphs."""
        code = token.get("raw", "")
        lines = code.rstrip("\n").split("\n")
        result = []
        for line in lines:
            p = self._new_para("CodeBlock")
            p.append(self._make_run(line, code=True, smart=False))
            result.append(p)
        return result

    def _render_list(self, token: dict, depth: int = 0) -> list[etree._Element]:
        """Render bullet or numbered list."""
        ordered = token.get("attrs", {}).get("ordered", False)
        num_id = "2" if ordered else "1"  # from numbering.xml: 1=bullet, 2=numbered
        result = []
        for item in token.get("children", []):
            item_type = item["type"]
            if item_type in ("list_item", "task_list_item"):
                result.extend(self._render_list_item(item, num_id, depth))
        return result

    def _render_list_item(self, token: dict, num_id: str, depth: int) -> list[etree._Element]:
        """Render a single list item, handling nested lists and task items."""
        result = []
        is_task = token["type"] == "task_list_item"

        for child in token.get("children", []):
            child_type = child["type"]
            # In tight lists, mistune uses "block_text" instead of "paragraph"
            if child_type in ("paragraph", "block_text"):
                p = self._new_para()
                # Add numPr for list numbering
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.SubElement(p, f"{W}pPr")
                    p.insert(0, ppr)
                num_pr = etree.SubElement(ppr, f"{W}numPr")
                ilvl = etree.SubElement(num_pr, f"{W}ilvl")
                ilvl.set(f"{W}val", str(depth))
                nid = etree.SubElement(num_pr, f"{W}numId")
                nid.set(f"{W}val", num_id)
                # Check for task list checkbox
                if is_task:
                    checked = token.get("attrs", {}).get("checked", False)
                    checkbox = "\u2611 " if checked else "\u2610 "
                    p.append(self._make_run(checkbox, smart=False))
                self._render_inline_children(p, child.get("children", []))
                result.append(p)
            elif child_type == "list":
                result.extend(self._render_list(child, depth=depth + 1))
        return result

    def _render_blockquote(self, token: dict, depth: int = 0) -> list[etree._Element]:
        """Render blockquote with increasing indent for nesting."""
        result = []
        for child in token.get("children", []):
            if child["type"] == "paragraph":
                p = self._new_para("BlockQuote")
                if depth > 0:
                    ppr = p.find(f"{W}pPr")
                    if ppr is not None:
                        ind = ppr.find(f"{W}ind")
                        if ind is None:
                            ind = etree.SubElement(ppr, f"{W}ind")
                        ind.set(f"{W}left", str(720 * (depth + 1)))
                self._render_inline_children(p, child.get("children", []))
                result.append(p)
            elif child["type"] == "block_quote":
                result.extend(self._render_blockquote(child, depth=depth + 1))
        return result

    def _render_hr(self) -> etree._Element:
        """Render horizontal rule as paragraph with bottom border."""
        p = self._new_para()
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(p, f"{W}pPr")
            p.insert(0, ppr)
        pbdr = etree.SubElement(ppr, f"{W}pBdr")
        bottom = etree.SubElement(pbdr, f"{W}bottom")
        bottom.set(f"{W}val", "single")
        bottom.set(f"{W}sz", "6")
        bottom.set(f"{W}space", "1")
        bottom.set(f"{W}color", "auto")
        return p

    def _render_link(self, parent: etree._Element, token: dict) -> None:
        """Render a hyperlink."""
        url = token["attrs"]["url"]
        rid = self._add_hyperlink_rel(url)
        hyperlink = etree.SubElement(parent, f"{W}hyperlink")
        hyperlink.set(f"{R}id", rid)
        r = etree.SubElement(hyperlink, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "Hyperlink")
        color = etree.SubElement(rpr, f"{W}color")
        color.set(f"{W}val", "0563C1")
        u = etree.SubElement(rpr, f"{W}u")
        u.set(f"{W}val", "single")
        for child in token.get("children", []):
            if child["type"] == "text":
                t = etree.SubElement(r, f"{W}t")
                _preserve(t, smartify(child["raw"]))

    def _add_hyperlink_rel(self, url: str) -> str:
        """Add a hyperlink relationship and return the rId."""
        rels = self._doc._tree("word/_rels/document.xml.rels")
        max_rid = 0
        for rel in rels.findall(f"{RELS}Relationship"):
            rid = rel.get("Id", "")
            if rid.startswith("rId"):
                with contextlib.suppress(ValueError):
                    max_rid = max(max_rid, int(rid[3:]))
        rid = f"rId{max_rid + 1}"
        new_rel = etree.SubElement(rels, f"{RELS}Relationship")
        new_rel.set("Id", rid)
        new_rel.set(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        )
        new_rel.set("Target", url)
        new_rel.set("TargetMode", "External")
        self._doc._mark("word/_rels/document.xml.rels")
        return rid

    def _render_image(self, parent: etree._Element, token: dict) -> None:
        """Render an image -- embed local, hyperlink remote."""
        src = token["attrs"].get("url", "")
        alt = ""
        if token.get("children"):
            for child in token["children"]:
                if child["type"] == "text":
                    alt = child.get("raw", "")
                    break

        if src.startswith(("http://", "https://")):
            # Remote -> hyperlink with alt text
            rid = self._add_hyperlink_rel(src)
            hyperlink = etree.SubElement(parent, f"{W}hyperlink")
            hyperlink.set(f"{R}id", rid)
            r = etree.SubElement(hyperlink, f"{W}r")
            rpr = etree.SubElement(r, f"{W}rPr")
            color = etree.SubElement(rpr, f"{W}color")
            color.set(f"{W}val", "0563C1")
            u = etree.SubElement(rpr, f"{W}u")
            u.set(f"{W}val", "single")
            t = etree.SubElement(r, f"{W}t")
            _preserve(t, alt or src)
        else:
            # Local -> embed
            img_path = self._base_dir / src
            if not img_path.exists():
                parent.append(self._make_run(f"[Image not found: {src}]", smart=False))
                return
            self._embed_image(parent, str(img_path))

    def _embed_image(self, parent: etree._Element, image_path: str) -> None:
        """Embed a local image file."""
        import shutil

        src = Path(image_path)
        ext = src.suffix.lstrip(".")
        media_dir = self._doc.workdir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        existing = list(media_dir.glob("image*.*"))
        idx = len(existing) + 1
        filename = f"image{idx}.{ext}"
        shutil.copy2(str(src), str(media_dir / filename))

        # Relationship
        rels = self._doc._tree("word/_rels/document.xml.rels")
        max_rid = 0
        for rel in rels.findall(f"{RELS}Relationship"):
            rid_str = rel.get("Id", "")
            if rid_str.startswith("rId"):
                with contextlib.suppress(ValueError):
                    max_rid = max(max_rid, int(rid_str[3:]))
        rid = f"rId{max_rid + 1}"
        new_rel = etree.SubElement(rels, f"{RELS}Relationship")
        new_rel.set("Id", rid)
        new_rel.set(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        )
        new_rel.set("Target", f"media/{filename}")
        self._doc._mark("word/_rels/document.xml.rels")

        # Content type
        ct_tree = self._doc._tree("[Content_Types].xml")
        ct_map = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
        }
        content_type = ct_map.get(ext, f"image/{ext}")
        has_ext = any(d.get("Extension") == ext for d in ct_tree.findall(f"{CT}Default"))
        if not has_ext:
            default = etree.SubElement(ct_tree, f"{CT}Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)
            self._doc._mark("[Content_Types].xml")

        # Drawing XML (simplified inline image)
        ns_wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ns_pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        run = etree.SubElement(parent, f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        inline = etree.SubElement(drawing, f"{{{ns_wp}}}inline")
        extent = etree.SubElement(inline, f"{{{ns_wp}}}extent")
        extent.set("cx", "2000000")
        extent.set("cy", "2000000")
        graphic = etree.SubElement(inline, f"{{{ns_a}}}graphic")
        gdata = etree.SubElement(graphic, f"{{{ns_a}}}graphicData", uri=ns_pic)
        pic = etree.SubElement(gdata, f"{{{ns_pic}}}pic")
        blip_fill = etree.SubElement(pic, f"{{{ns_pic}}}blipFill")
        blip = etree.SubElement(blip_fill, f"{{{ns_a}}}blip")
        blip.set(f"{{{ns_r}}}embed", rid)

    def _render_table(self, token: dict) -> etree._Element:
        """Render a markdown table as w:tbl."""
        tbl = etree.Element(f"{W}tbl")
        tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
        style = etree.SubElement(tbl_pr, f"{W}tblStyle")
        style.set(f"{W}val", "TableGrid")
        tw = etree.SubElement(tbl_pr, f"{W}tblW")
        tw.set(f"{W}w", "0")
        tw.set(f"{W}type", "auto")
        # Add borders
        tbl_borders = etree.SubElement(tbl_pr, f"{W}tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            bdr = etree.SubElement(tbl_borders, f"{W}{side}")
            bdr.set(f"{W}val", "single")
            bdr.set(f"{W}sz", "4")
            bdr.set(f"{W}space", "0")
            bdr.set(f"{W}color", "auto")

        # Process children: table_head and table_body
        for child in token.get("children", []):
            if child["type"] == "table_head":
                # table_head contains table_cell elements directly (one header row)
                tr = self._render_table_head_row(child, bold=True)
                tbl.append(tr)
            elif child["type"] == "table_body":
                # table_body contains table_row elements
                for row_token in child.get("children", []):
                    tr = self._render_table_row(row_token, bold=False)
                    tbl.append(tr)
        return tbl

    def _render_table_head_row(self, token: dict, bold: bool = False) -> etree._Element:
        """Render the table head as a single row.

        In mistune 3.x, table_head contains table_cell children directly
        (not wrapped in a table_row).
        """
        tr = etree.Element(f"{W}tr")
        tr.set(f"{W14}paraId", self._doc._new_para_id())
        tr.set(f"{W14}textId", "77777777")
        for cell_token in token.get("children", []):
            if cell_token["type"] == "table_cell":
                tc = etree.SubElement(tr, f"{W}tc")
                p = self._new_para()
                if cell_token.get("children"):
                    self._render_inline_children(p, cell_token["children"], bold=bold)
                tc.append(p)
        return tr

    def _render_table_row(self, token: dict, bold: bool = False) -> etree._Element:
        """Render a table body row."""
        tr = etree.Element(f"{W}tr")
        tr.set(f"{W14}paraId", self._doc._new_para_id())
        tr.set(f"{W14}textId", "77777777")
        for cell_token in token.get("children", []):
            if cell_token["type"] == "table_cell":
                tc = etree.SubElement(tr, f"{W}tc")
                p = self._new_para()
                if cell_token.get("children"):
                    self._render_inline_children(p, cell_token["children"], bold=bold)
                tc.append(p)
        return tr

    def _process_footnote_definitions(self, token: dict) -> None:
        """Process footnote definitions and add them to footnotes.xml."""
        fn_tree = self._doc._tree("word/footnotes.xml")
        if fn_tree is None:
            return

        existing = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
        next_id = max(existing | {0}) + 1

        for item in token.get("children", []):
            if item["type"] == "footnote_item":
                key = item.get("attrs", {}).get("key", "")
                fn_el = etree.SubElement(fn_tree, f"{W}footnote")
                fn_el.set(f"{W}id", str(next_id))

                fn_para = etree.SubElement(fn_el, f"{W}p")
                fn_para.set(f"{W14}paraId", self._doc._new_para_id())
                fn_para.set(f"{W14}textId", "77777777")

                ppr = etree.SubElement(fn_para, f"{W}pPr")
                ps = etree.SubElement(ppr, f"{W}pStyle")
                ps.set(f"{W}val", "FootnoteText")

                # Footnote ref mark
                ref_run = etree.SubElement(fn_para, f"{W}r")
                ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
                ref_style = etree.SubElement(ref_rpr, f"{W}rStyle")
                ref_style.set(f"{W}val", "FootnoteReference")
                etree.SubElement(ref_run, f"{W}footnoteRef")

                # Space
                sp_run = etree.SubElement(fn_para, f"{W}r")
                sp_t = etree.SubElement(sp_run, f"{W}t")
                _preserve(sp_t, " ")

                # Text from children
                for child in item.get("children", []):
                    if child["type"] == "paragraph":
                        for inline in child.get("children", []):
                            if inline["type"] == "text":
                                txt_run = etree.SubElement(fn_para, f"{W}r")
                                txt_t = etree.SubElement(txt_run, f"{W}t")
                                _preserve(txt_t, smartify(inline["raw"]))

                self._footnote_map[key] = next_id
                next_id += 1

        self._doc._mark("word/footnotes.xml")

    def _render_footnote_ref(self, parent: etree._Element, token: dict) -> None:
        """Render a footnote reference in the body."""
        # In mistune 3.x, footnote_ref has 'raw' containing the key
        key = token.get("raw", token.get("attrs", {}).get("key", ""))
        fn_id = self._footnote_map.get(key)
        if fn_id is None:
            return
        r = etree.SubElement(parent, f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        rs = etree.SubElement(rpr, f"{W}rStyle")
        rs.set(f"{W}val", "FootnoteReference")
        fref = etree.SubElement(r, f"{W}footnoteReference")
        fref.set(f"{W}id", str(fn_id))
