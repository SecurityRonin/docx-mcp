"""Schema-ordered child insertion for OOXML sequence types.

w:tcPr, w:tblPr and w:settings are xsd:sequence types: their children must
appear in schema order. etree.SubElement always appends, so writers that
find-or-create properties produce insertion order instead. See
docs/IMPLEMENTATION-PLAN-table-and-deploy.md §A.2.
"""

from __future__ import annotations

import pytest
from docx_mcp.document.ooxml_order import (
    TBLPR_ORDER,
    TCPR_ORDER,
    find_out_of_order,
    ordered_set_child,
)
from lxml import etree

from docx_mcp.document.base import W


def _tcpr(*localnames: str) -> etree._Element:
    el = etree.Element(f"{W}tcPr")
    for name in localnames:
        etree.SubElement(el, f"{W}{name}")
    return el


def _names(parent: etree._Element) -> list[str]:
    return [etree.QName(c).localname for c in parent if isinstance(c.tag, str)]


# ── ordered_set_child: placement ────────────────────────────────────────────


class TestInsertsAtSchemaPosition:
    def test_inserts_before_a_later_sibling(self):
        """gridSpan sorts before shd, so it must land in front of it."""
        pr = _tcpr("shd")
        ordered_set_child(pr, "gridSpan", TCPR_ORDER)
        assert _names(pr) == ["gridSpan", "shd"]

    def test_appends_when_it_sorts_last(self):
        pr = _tcpr("gridSpan")
        ordered_set_child(pr, "vAlign", TCPR_ORDER)
        assert _names(pr) == ["gridSpan", "vAlign"]

    def test_inserts_between_existing_children(self):
        pr = _tcpr("tcW", "shd", "vAlign")
        ordered_set_child(pr, "tcMar", TCPR_ORDER)
        assert _names(pr) == ["tcW", "shd", "tcMar", "vAlign"]

    def test_inserts_into_empty_parent(self):
        pr = etree.Element(f"{W}tcPr")
        ordered_set_child(pr, "shd", TCPR_ORDER)
        assert _names(pr) == ["shd"]

    def test_result_is_order_independent(self):
        """Applying the same properties in any sequence yields one canonical order."""
        forward = etree.Element(f"{W}tcPr")
        for name in ("gridSpan", "shd", "tcMar", "vAlign"):
            ordered_set_child(forward, name, TCPR_ORDER)

        backward = etree.Element(f"{W}tcPr")
        for name in ("vAlign", "tcMar", "shd", "gridSpan"):
            ordered_set_child(backward, name, TCPR_ORDER)

        assert _names(forward) == _names(backward)

    def test_works_for_tblpr_too(self):
        pr = etree.Element(f"{W}tblPr")
        ordered_set_child(pr, "tblLayout", TBLPR_ORDER)
        ordered_set_child(pr, "tblStyle", TBLPR_ORDER)
        ordered_set_child(pr, "tblW", TBLPR_ORDER)
        assert _names(pr) == ["tblStyle", "tblW", "tblLayout"]


class TestFindOrCreateIsIdempotent:
    def test_returns_the_existing_child(self):
        pr = _tcpr("shd")
        existing = pr.find(f"{W}shd")
        assert ordered_set_child(pr, "shd", TCPR_ORDER) is existing

    def test_does_not_duplicate_on_reapplication(self):
        pr = etree.Element(f"{W}tcPr")
        for _ in range(3):
            ordered_set_child(pr, "shd", TCPR_ORDER)
        assert _names(pr) == ["shd"]

    def test_preserves_attributes_of_the_existing_child(self):
        pr = _tcpr("shd")
        pr.find(f"{W}shd").set(f"{W}fill", "FF0000")
        got = ordered_set_child(pr, "shd", TCPR_ORDER)
        assert got.get(f"{W}fill") == "FF0000"

    def test_returns_an_element_in_the_tree(self):
        pr = etree.Element(f"{W}tcPr")
        child = ordered_set_child(pr, "tcMar", TCPR_ORDER)
        assert child.getparent() is pr


class TestRejectsUnknownNames:
    def test_raises_on_a_name_outside_the_order(self):
        pr = etree.Element(f"{W}tcPr")
        with pytest.raises(ValueError, match="notAnElement"):
            ordered_set_child(pr, "notAnElement", TCPR_ORDER)

    def test_nothing_is_inserted_when_rejected(self):
        pr = etree.Element(f"{W}tcPr")
        with pytest.raises(ValueError):
            ordered_set_child(pr, "bogus", TCPR_ORDER)
        assert len(pr) == 0


class TestToleratesForeignContent:
    def test_comments_do_not_break_placement(self):
        pr = _tcpr("shd")
        pr.insert(0, etree.Comment("authored by hand"))
        ordered_set_child(pr, "gridSpan", TCPR_ORDER)
        assert _names(pr) == ["gridSpan", "shd"]

    def test_unknown_elements_are_stepped_over(self):
        """An extension element we don't model must not misplace a known one."""
        pr = _tcpr("shd")
        etree.SubElement(pr, "{urn:vendor:ext}custom")
        ordered_set_child(pr, "gridSpan", TCPR_ORDER)
        assert _names(pr)[:2] == ["gridSpan", "shd"]


# ── find_out_of_order: the detector behind the M1.8 warning pass ────────────


class TestFindOutOfOrder:
    def test_clean_parent_reports_nothing(self):
        assert find_out_of_order(_tcpr("gridSpan", "shd", "vAlign"), TCPR_ORDER) == []

    def test_flags_a_swapped_pair(self):
        assert find_out_of_order(_tcpr("shd", "gridSpan"), TCPR_ORDER) == ["gridSpan"]

    def test_names_every_misplaced_child(self):
        got = find_out_of_order(_tcpr("vAlign", "shd", "gridSpan"), TCPR_ORDER)
        assert got == ["shd", "gridSpan"]

    def test_ignores_unknown_and_foreign_children(self):
        pr = _tcpr("gridSpan", "shd")
        etree.SubElement(pr, "{urn:vendor:ext}custom")
        pr.append(etree.Comment("x"))
        assert find_out_of_order(pr, TCPR_ORDER) == []

    def test_empty_parent_is_clean(self):
        assert find_out_of_order(etree.Element(f"{W}tcPr"), TCPR_ORDER) == []


# ── The order tables themselves ─────────────────────────────────────────────


class TestOrderTables:
    def test_tcpr_order_has_no_duplicates(self):
        assert len(TCPR_ORDER) == len(set(TCPR_ORDER))

    def test_tblpr_order_has_no_duplicates(self):
        assert len(TBLPR_ORDER) == len(set(TBLPR_ORDER))

    def test_tcpr_covers_the_properties_this_package_writes(self):
        for name in ("tcW", "gridSpan", "vMerge", "tcBorders", "shd", "tcMar", "vAlign"):
            assert name in TCPR_ORDER

    def test_tblpr_covers_the_properties_this_package_writes(self):
        for name in ("tblStyle", "tblW", "jc", "tblBorders", "tblLayout", "tblCellMar"):
            assert name in TBLPR_ORDER

    def test_tcpr_revision_marker_sorts_last(self):
        assert TCPR_ORDER[-1] == "tcPrChange"

    def test_tblpr_revision_marker_sorts_last(self):
        assert TBLPR_ORDER[-1] == "tblPrChange"
