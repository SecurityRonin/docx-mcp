"""Per-cell borders, cell alignment, and cell text formatting.

See docs/IMPLEMENTATION-PLAN-table-and-deploy.md §A.3. set_table_borders
already covers table-level borders; these are the per-cell equivalents plus
the two convenience tools that otherwise take three or more calls.
"""

from __future__ import annotations

import pytest

from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.document.ooxml_order import TCPR_ORDER, find_out_of_order


def _doc_with_table(tmp_path, rows=2, cols=2):
    doc = DocxDocument.create(str(tmp_path / "t.docx"))
    tree = doc._tree("word/document.xml")
    para_id = tree.findall(f".//{W}p")[0].get(f"{W14}paraId")
    doc.add_table(para_id, rows, cols)
    return doc


def _cell(doc, row=0, col=0, table_idx=0):
    tbl = doc._get_table(table_idx)
    return tbl.findall(f"{W}tr")[row].findall(f"{W}tc")[col]


def _tcpr(doc, row=0, col=0):
    return _cell(doc, row, col).find(f"{W}tcPr")


def _names(parent):
    from lxml import etree

    return [etree.QName(c).localname for c in parent if isinstance(c.tag, str)]


# ── set_cell_borders ────────────────────────────────────────────────────────


class TestSetCellBorders:
    def test_defaults_to_the_four_outer_sides(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0)
        borders = _tcpr(doc).find(f"{W}tcBorders")
        assert _names(borders) == ["top", "left", "bottom", "right"]

    def test_writes_style_size_and_color(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0, sides=["top"], style="double", color="FF0000", size=8)
        top = _tcpr(doc).find(f"{W}tcBorders").find(f"{W}top")
        assert top.get(f"{W}val") == "double"
        assert top.get(f"{W}sz") == "8"
        assert top.get(f"{W}color") == "FF0000"

    def test_named_sides_only(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0, sides=["top", "bottom"])
        assert _names(_tcpr(doc).find(f"{W}tcBorders")) == ["top", "bottom"]

    def test_successive_calls_compose_rather_than_replace(self, tmp_path):
        """Setting the left border must not wipe a top border set earlier."""
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0, sides=["top"])
        doc.set_cell_borders(0, 0, 0, sides=["left"])
        assert _names(_tcpr(doc).find(f"{W}tcBorders")) == ["top", "left"]

    def test_sides_land_in_schema_order(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0, sides=["right"])
        doc.set_cell_borders(0, 0, 0, sides=["top"])
        doc.set_cell_borders(0, 0, 0, sides=["insideV"])
        doc.set_cell_borders(0, 0, 0, sides=["bottom"])
        assert _names(_tcpr(doc).find(f"{W}tcBorders")) == [
            "top",
            "bottom",
            "right",
            "insideV",
        ]

    def test_reapplying_a_side_updates_it_in_place(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0, sides=["top"], size=4)
        doc.set_cell_borders(0, 0, 0, sides=["top"], size=12)
        borders = _tcpr(doc).find(f"{W}tcBorders")
        assert len(borders.findall(f"{W}top")) == 1
        assert borders.find(f"{W}top").get(f"{W}sz") == "12"

    def test_diagonals_are_available(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_borders(0, 0, 0, sides=["tl2br"])
        assert _names(_tcpr(doc).find(f"{W}tcBorders")) == ["tl2br"]

    def test_unknown_side_names_the_offending_value(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError, match="diagonalish"):
            doc.set_cell_borders(0, 0, 0, sides=["diagonalish"])

    def test_bad_color_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_cell_borders(0, 0, 0, color="not-hex")

    def test_empty_side_list_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_cell_borders(0, 0, 0, sides=[])

    def test_keeps_tcpr_in_schema_order(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_shading(0, 0, 0, "EEEEEE")
        doc.set_cell_borders(0, 0, 0)
        doc.set_cell_padding(0, 0, 0, top_mm=1.0)
        assert find_out_of_order(_tcpr(doc), TCPR_ORDER) == []

    def test_returns_the_sides_it_wrote(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        got = doc.set_cell_borders(0, 1, 1, sides=["top"])
        assert got["sides"] == ["top"]
        assert got["row_idx"] == 1


# ── set_cell_alignment ──────────────────────────────────────────────────────


class TestSetCellAlignment:
    def test_vertical_writes_valign(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_alignment(0, 0, 0, vertical="center")
        assert _tcpr(doc).find(f"{W}vAlign").get(f"{W}val") == "center"

    def test_horizontal_writes_jc_on_cell_paragraphs(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_alignment(0, 0, 0, horizontal="right")
        para = _cell(doc).find(f"{W}p")
        assert para.find(f"{W}pPr").find(f"{W}jc").get(f"{W}val") == "right"

    def test_both_axes_at_once(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_alignment(0, 0, 0, horizontal="center", vertical="bottom")
        assert _tcpr(doc).find(f"{W}vAlign").get(f"{W}val") == "bottom"
        assert _cell(doc).find(f"{W}p").find(f"{W}pPr").find(f"{W}jc") is not None

    def test_applies_to_every_paragraph_in_the_cell(self, tmp_path):
        from lxml import etree

        doc = _doc_with_table(tmp_path)
        tc = _cell(doc)
        etree.SubElement(tc, f"{W}p")
        doc.set_cell_alignment(0, 0, 0, horizontal="center")
        assert all(p.find(f"{W}pPr").find(f"{W}jc") is not None for p in tc.findall(f"{W}p"))

    def test_reapplication_does_not_duplicate_jc(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_alignment(0, 0, 0, horizontal="left")
        doc.set_cell_alignment(0, 0, 0, horizontal="center")
        ppr = _cell(doc).find(f"{W}p").find(f"{W}pPr")
        assert len(ppr.findall(f"{W}jc")) == 1
        assert ppr.find(f"{W}jc").get(f"{W}val") == "center"

    def test_vertical_only_leaves_paragraphs_alone(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.set_cell_alignment(0, 0, 0, vertical="center")
        assert _cell(doc).find(f"{W}p").find(f"{W}pPr") is None

    def test_neither_axis_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.set_cell_alignment(0, 0, 0)

    def test_unknown_vertical_names_the_offending_value(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError, match="sideways"):
            doc.set_cell_alignment(0, 0, 0, vertical="sideways")

    def test_unknown_horizontal_names_the_offending_value(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError, match="diagonal"):
            doc.set_cell_alignment(0, 0, 0, horizontal="diagonal")

    def test_returns_both_axes(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        got = doc.set_cell_alignment(0, 0, 0, horizontal="center", vertical="top")
        assert got["horizontal"] == "center"
        assert got["vertical"] == "top"


# ── format_cell ─────────────────────────────────────────────────────────────


class TestFormatCell:
    def test_bold_applies_to_the_cell_runs(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, bold=True)
        run = _cell(doc).find(f"{W}p").find(f"{W}r")
        assert run.find(f"{W}rPr").find(f"{W}b") is not None

    def test_bold_false_writes_an_explicit_off(self, tmp_path):
        """An absent w:b inherits from the style; val="0" overrides it."""
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, bold=False)
        b = _cell(doc).find(f"{W}p").find(f"{W}r").find(f"{W}rPr").find(f"{W}b")
        assert b.get(f"{W}val") == "0"

    def test_font_size_is_written_in_half_points(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, font_size_pt=11.0)
        rpr = _cell(doc).find(f"{W}p").find(f"{W}r").find(f"{W}rPr")
        assert rpr.find(f"{W}sz").get(f"{W}val") == "22"

    def test_font_name_sets_ascii_and_hansi(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, font_name="Calibri")
        fonts = _cell(doc).find(f"{W}p").find(f"{W}r").find(f"{W}rPr").find(f"{W}rFonts")
        assert fonts.get(f"{W}ascii") == "Calibri"
        assert fonts.get(f"{W}hAnsi") == "Calibri"

    def test_color_is_written(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, color="FF0000")
        rpr = _cell(doc).find(f"{W}p").find(f"{W}r").find(f"{W}rPr")
        assert rpr.find(f"{W}color").get(f"{W}val") == "FF0000"

    def test_underline_and_italic(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, italic=True, underline=True)
        rpr = _cell(doc).find(f"{W}p").find(f"{W}r").find(f"{W}rPr")
        assert rpr.find(f"{W}i") is not None
        assert rpr.find(f"{W}u").get(f"{W}val") == "single"

    def test_reports_how_many_runs_it_touched(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        assert doc.format_cell(0, 0, 0, bold=True)["runs_formatted"] == 1

    def test_empty_cell_formats_nothing_but_does_not_raise(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        assert doc.format_cell(0, 0, 0, bold=True)["runs_formatted"] == 0

    def test_reapplication_does_not_duplicate_properties(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc.format_cell(0, 0, 0, bold=True)
        doc.format_cell(0, 0, 0, bold=True)
        rpr = _cell(doc).find(f"{W}p").find(f"{W}r").find(f"{W}rPr")
        assert len(rpr.findall(f"{W}b")) == 1

    def test_no_properties_given_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.format_cell(0, 0, 0)

    def test_bad_color_is_rejected(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(ValueError):
            doc.format_cell(0, 0, 0, color="nope")

    def test_out_of_range_cell_raises(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.format_cell(0, 9, 0, bold=True)

    def test_marks_the_document_dirty(self, tmp_path):
        doc = _doc_with_table(tmp_path)
        doc.modify_cell(0, 0, 0, "hello", tracked=False)
        doc._modified.clear()
        doc.format_cell(0, 0, 0, bold=True)
        assert "word/document.xml" in doc._modified
