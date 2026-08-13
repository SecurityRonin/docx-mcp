"""Tests for image management tools: delete_image, update_image, set_image_size, set_image_alt_text."""  # noqa: E501

from __future__ import annotations

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.document.base import RELS, W14, WP, A, R, W

_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"

_TINY_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
    b"\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02"
    b"\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx"
    b"\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N"
    b"\x00\x00\x00\x00IEND\xaeB`\x82"
)


@pytest.fixture
def sample_png(tmp_path):
    p = tmp_path / "sample.png"
    p.write_bytes(_TINY_PNG)
    return p


@pytest.fixture
def doc_with_image(tmp_path, sample_png):
    """DocxDocument with one inline image already inserted."""
    doc = DocxDocument.create(str(tmp_path / "test.docx"))
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    para_id = paras[0].get(f"{W14}paraId")
    result = doc.insert_image(para_id, str(sample_png))
    return doc, result["rId"]


class TestDeleteImage:
    def test_delete_removes_run_from_paragraph(self, doc_with_image):
        """delete_image removes the w:r containing the drawing from the document."""
        doc, rid = doc_with_image
        tree = doc._tree("word/document.xml")
        # Confirm blip exists before delete
        blips_before = [b for b in tree.iter(f"{A}blip") if b.get(f"{R}embed") == rid]
        assert len(blips_before) == 1

        result = doc.delete_image(rid)

        assert result == {"deleted": rid}
        tree = doc._tree("word/document.xml")
        blips_after = [b for b in tree.iter(f"{A}blip") if b.get(f"{R}embed") == rid]
        assert len(blips_after) == 0

    def test_delete_removes_relationship(self, doc_with_image):
        """delete_image removes the relationship entry from document.xml.rels."""
        doc, rid = doc_with_image
        rels = doc._tree("word/_rels/document.xml.rels")
        rel_before = rels.find(f'{RELS}Relationship[@Id="{rid}"]')
        assert rel_before is not None

        doc.delete_image(rid)

        rels = doc._tree("word/_rels/document.xml.rels")
        rel_after = rels.find(f'{RELS}Relationship[@Id="{rid}"]')
        assert rel_after is None

    def test_delete_marks_dirty(self, doc_with_image):
        """delete_image marks document.xml and rels as modified."""
        doc, rid = doc_with_image
        doc._modified.clear()
        doc.delete_image(rid)
        assert "word/document.xml" in doc._modified
        assert "word/_rels/document.xml.rels" in doc._modified

    def test_delete_nonexistent_rid_raises(self, doc_with_image):
        """delete_image raises ValueError for unknown rId."""
        doc, _ = doc_with_image
        with pytest.raises(ValueError, match="rId99"):
            doc.delete_image("rId99")


class TestUpdateImage:
    def test_update_stores_new_bytes_in_binaries(self, doc_with_image, tmp_path):
        """update_image stores new image bytes under the zip path in _binaries."""
        doc, rid = doc_with_image
        new_img = tmp_path / "new.png"
        new_img.write_bytes(_TINY_PNG + b"\x00")  # slightly different bytes

        result = doc.update_image(rid, str(new_img))

        assert result["rId"] == rid
        assert result["new_path"] == str(new_img)
        # At least one binary entry should contain the new bytes
        assert any(v == new_img.read_bytes() for v in doc._binaries.values())

    def test_update_nonexistent_rid_raises(self, doc_with_image, tmp_path):
        """update_image raises ValueError for unknown rId."""
        doc, _ = doc_with_image
        new_img = tmp_path / "new.png"
        new_img.write_bytes(_TINY_PNG)
        with pytest.raises(ValueError, match="rId99"):
            doc.update_image("rId99", str(new_img))


class TestSetImageSize:
    def test_set_size_updates_extent_cx_cy(self, doc_with_image):
        """set_image_size updates wp:extent cx and cy to correct EMU values."""
        doc, rid = doc_with_image
        width_cm, height_cm = 5.0, 3.0
        expected_cx = round(width_cm * 360000)
        expected_cy = round(height_cm * 360000)

        result = doc.set_image_size(rid, width_cm, height_cm)

        assert result["rId"] == rid
        assert result["width_cm"] == width_cm
        assert result["height_cm"] == height_cm
        assert result["width_emu"] == expected_cx
        assert result["height_emu"] == expected_cy

        tree = doc._tree("word/document.xml")
        blip = next(b for b in tree.iter(f"{A}blip") if b.get(f"{R}embed") == rid)
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        assert drawing is not None
        extent = drawing.find(f".//{WP}extent")
        assert extent is not None
        assert int(extent.get("cx")) == expected_cx
        assert int(extent.get("cy")) == expected_cy

    def test_set_size_marks_dirty(self, doc_with_image):
        """set_image_size marks document.xml as modified."""
        doc, rid = doc_with_image
        doc._modified.clear()
        doc.set_image_size(rid, 4.0, 2.0)
        assert "word/document.xml" in doc._modified

    def test_set_size_nonexistent_rid_raises(self, doc_with_image):
        """set_image_size raises ValueError for unknown rId."""
        doc, _ = doc_with_image
        with pytest.raises(ValueError, match="rId99"):
            doc.set_image_size("rId99", 5.0, 3.0)


class TestSetImageAltText:
    def test_set_alt_text_sets_descr_attribute(self, doc_with_image):
        """set_image_alt_text sets descr on wp:docPr."""
        doc, rid = doc_with_image

        # Insert a docPr so we have something to update (inline images may not have one)
        tree = doc._tree("word/document.xml")
        blip = next(b for b in tree.iter(f"{A}blip") if b.get(f"{R}embed") == rid)
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        # Ensure wp:docPr exists in the drawing
        doc_pr = drawing.find(f".//{WP}docPr")
        if doc_pr is None:
            inline = drawing.find(f"{WP}inline")
            if inline is None:
                inline = drawing.find(f"{WP}anchor")
            doc_pr = etree.SubElement(inline, f"{WP}docPr")
            doc_pr.set("id", "1")
            doc_pr.set("name", "Image 1")

        result = doc.set_image_alt_text(rid, "A red circle", title="Circle")

        assert result["rId"] == rid
        assert result["alt_text"] == "A red circle"

        tree = doc._tree("word/document.xml")
        blip = next(b for b in tree.iter(f"{A}blip") if b.get(f"{R}embed") == rid)
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        doc_pr = drawing.find(f".//{WP}docPr")
        assert doc_pr is not None
        assert doc_pr.get("descr") == "A red circle"

    def test_set_alt_text_marks_dirty(self, doc_with_image):
        """set_image_alt_text marks document.xml as modified."""
        doc, rid = doc_with_image
        # Ensure docPr exists
        tree = doc._tree("word/document.xml")
        blip = next(b for b in tree.iter(f"{A}blip") if b.get(f"{R}embed") == rid)
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        doc_pr = drawing.find(f".//{WP}docPr")
        if doc_pr is None:
            inline = drawing.find(f"{WP}inline") or drawing.find(f"{WP}anchor")
            doc_pr = etree.SubElement(inline, f"{WP}docPr")
            doc_pr.set("id", "1")
            doc_pr.set("name", "Image 1")

        doc._modified.clear()
        doc.set_image_alt_text(rid, "desc")
        assert "word/document.xml" in doc._modified

    def test_set_alt_text_nonexistent_rid_raises(self, doc_with_image):
        """set_image_alt_text raises ValueError for unknown rId."""
        doc, _ = doc_with_image
        with pytest.raises(ValueError, match="rId99"):
            doc.set_image_alt_text("rId99", "some text")


# ── server-level: document_handle support (issue #8) ─────────────────────────


def test_update_image_with_explicit_handle(tmp_path, sample_png):
    """update_image can target a document under a non-default handle."""
    import json

    from docx_mcp import server

    # Build a DOCX file with one image, saved to disk
    doc_path = tmp_path / "with_image.docx"
    doc = DocxDocument.create(str(doc_path))
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    para_id = paras[0].get(f"{W14}paraId")
    result = doc.insert_image(para_id, str(sample_png))
    rid = result["rId"]
    doc.save()
    doc.close()

    h1 = "image-handle-1"
    h2 = "image-handle-2"
    try:
        server.open_document(str(doc_path), document_handle=h1)
        server.open_document(str(doc_path), document_handle=h2)
        new_img = tmp_path / "new.png"
        new_img.write_bytes(_TINY_PNG + b"\x00")

        out = json.loads(server.update_image(rid, str(new_img), document_handle=h1))
        assert out["rId"] == rid
        # Handle h2's doc should NOT have the updated binary
        key2 = h2
        doc2 = server._docs[key2]
        assert not any(v == new_img.read_bytes() for v in doc2._binaries.values()), (
            "h2 document must not be affected by update on h1"
        )
    finally:
        server.close_document(document_handle=h1)
        server.close_document(document_handle=h2)
