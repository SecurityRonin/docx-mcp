"""Images mixin: list and insert embedded images."""

from __future__ import annotations

import shutil
from pathlib import Path

from lxml import etree

from .base import CT, RELS, W14, WP, A, R, W


class ImagesMixin:
    """Image operations."""

    def get_images(self) -> list[dict]:
        """Get all embedded images with metadata."""
        doc = self._tree("word/document.xml")
        rels = self._tree("word/_rels/document.xml.rels")
        if doc is None:
            return []
        images = []
        for blip in doc.iter(f"{A}blip"):
            embed = blip.get(f"{R}embed")
            if not embed:
                continue
            info: dict = {"rId": embed, "filename": "", "content_type": ""}
            if rels is not None:
                rel = rels.find(f'{RELS}Relationship[@Id="{embed}"]')
                if rel is not None:
                    info["filename"] = rel.get("Target", "").split("/")[-1]
            # Get dimensions from wp:extent
            drawing = blip.getparent()
            while drawing is not None and drawing.tag != f"{W}drawing":
                drawing = drawing.getparent()
            if drawing is not None:
                extent = drawing.find(f".//{WP}extent")
                if extent is not None:
                    info["width_emu"] = int(extent.get("cx", "0"))
                    info["height_emu"] = int(extent.get("cy", "0"))
            # Content type from [Content_Types].xml
            ct_tree = self._tree("[Content_Types].xml")
            if ct_tree is not None:
                ext = info["filename"].rsplit(".", 1)[-1] if "." in info["filename"] else ""
                for default in ct_tree.findall(f"{CT}Default"):
                    if default.get("Extension") == ext:
                        info["content_type"] = default.get("ContentType", "")
            images.append(info)
        return images

    def insert_image(
        self,
        para_id: str,
        image_path: str,
        *,
        width_emu: int = 2000000,
        height_emu: int = 2000000,
    ) -> dict:
        """Insert an image into the document after a paragraph.

        Args:
            para_id: paraId of the paragraph to insert after.
            image_path: Path to the image file on disk.
            width_emu: Image width in EMUs (914400 = 1 inch).
            height_emu: Image height in EMUs.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        src = Path(image_path)
        ext = src.suffix.lstrip(".")

        # Copy image to word/media/
        media_dir = self.workdir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        existing = list(media_dir.glob("image*.*"))
        idx = len(existing) + 1
        filename = f"image{idx}.{ext}"
        shutil.copy2(str(src), str(media_dir / filename))

        # Add relationship
        rels = self._tree("word/_rels/document.xml.rels")
        existing_rids = [r.get("Id") for r in rels.findall(f"{RELS}Relationship")]
        rid_num = (
            max(
                (int(r.replace("rId", "")) for r in existing_rids if r.startswith("rId")),
                default=0,
            )
            + 1
        )
        rid = f"rId{rid_num}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel.set("Target", f"media/{filename}")
        self._mark("word/_rels/document.xml.rels")

        # Ensure content type exists
        ct_tree = self._tree("[Content_Types].xml")
        ct_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif"}
        content_type = ct_map.get(ext, f"image/{ext}")
        has_ext = any(d.get("Extension") == ext for d in ct_tree.findall(f"{CT}Default"))
        if not has_ext:
            default = etree.SubElement(ct_tree, f"{CT}Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)
            self._mark("[Content_Types].xml")

        # Build drawing XML
        ns_wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ns_pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", self._new_para_id())
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        inline = etree.SubElement(drawing, f"{{{ns_wp}}}inline")
        extent = etree.SubElement(inline, f"{{{ns_wp}}}extent")
        extent.set("cx", str(width_emu))
        extent.set("cy", str(height_emu))
        graphic = etree.SubElement(inline, f"{{{ns_a}}}graphic")
        gdata = etree.SubElement(
            graphic,
            f"{{{ns_a}}}graphicData",
            uri=ns_pic,
        )
        pic = etree.SubElement(gdata, f"{{{ns_pic}}}pic")
        blip_fill = etree.SubElement(pic, f"{{{ns_pic}}}blipFill")
        blip = etree.SubElement(blip_fill, f"{{{ns_a}}}blip")
        blip.set(f"{{{ns_r}}}embed", rid)

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {
            "filename": filename,
            "rId": rid,
            "width_emu": width_emu,
            "height_emu": height_emu,
        }
