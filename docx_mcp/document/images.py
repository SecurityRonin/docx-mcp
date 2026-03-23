"""Images mixin: list embedded images."""

from __future__ import annotations

from .base import A, CT, R, RELS, W, WP


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
