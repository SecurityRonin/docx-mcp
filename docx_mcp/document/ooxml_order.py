"""Schema-ordered child insertion for OOXML sequence types.

`w:tcPr`, `w:tblPr` and friends are `xsd:sequence` complex types: their
children are valid only in schema order. `etree.SubElement` appends, so the
usual find-or-create idiom produces *insertion* order instead — shade a cell
and then merge it and you get `shd` before `gridSpan`, which the schema
forbids. Word tolerates it; stricter validators and some toolchains do not.

Every property writer goes through `ordered_set_child`, which finds or
creates a child at its schema position, so the resulting order is canonical
regardless of the order the tools were called in.

Order tables are the effective element sequence of the transitional
ECMA-376 Part 1 schema (§17.4 Tables). The package models only the
transitional namespace, so the strict-schema spellings (`start`/`end` in
place of `left`/`right`) are not represented here.
"""

from __future__ import annotations

from lxml import etree

from .base import W

# CT_TcPr = CT_TcPrBase (cnfStyle … headers) + the revision-tracking tail.
TCPR_ORDER: tuple[str, ...] = (
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
)

# CT_TblPr = CT_TblPrBase (tblStyle … tblDescription) + tblPrChange.
TBLPR_ORDER: tuple[str, ...] = (
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
)

# CT_TrPr — where gridBefore/gridAfter sit matters for row grid width.
TRPR_ORDER: tuple[str, ...] = (
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
    "ins",
    "del",
    "trPrChange",
)

# CT_TcMar / CT_TblCellMar, transitional spelling.
TCMAR_ORDER: tuple[str, ...] = ("top", "left", "bottom", "right")

# CT_TcBorders, transitional spelling.
TCBORDERS_ORDER: tuple[str, ...] = (
    "top",
    "left",
    "bottom",
    "right",
    "insideH",
    "insideV",
    "tl2br",
    "tr2bl",
)


def _local(el: etree._Element) -> str | None:
    """Local name of an element, or None for comments and processing instructions."""
    if not isinstance(el.tag, str):
        return None
    return etree.QName(el).localname


def ordered_set_child(
    parent: etree._Element,
    localname: str,
    order: tuple[str, ...],
    *,
    ns: str = W,
) -> etree._Element:
    """Find or create `parent`'s `localname` child at its schema position.

    Returns the existing child untouched if there is one, so callers can set
    attributes on it without duplicating the element. Children the order
    table does not know about — foreign-namespace extensions, comments — are
    stepped over rather than used as anchors.

    Raises:
        ValueError: If `localname` is not part of `order`. A typo here would
            otherwise place an element the schema has no slot for.
    """
    if localname not in order:
        raise ValueError(
            f"{localname!r} is not a child of this element's schema sequence. "
            f"Known children: {', '.join(order)}"
        )

    tag = f"{ns}{localname}"
    existing = parent.find(tag)
    if existing is not None:
        return existing

    position = order.index(localname)
    child = etree.Element(tag)
    for index, sibling in enumerate(parent):
        name = _local(sibling)
        if name in order and order.index(name) > position:
            parent.insert(index, child)
            return child
    parent.append(child)
    return child


def find_out_of_order(parent: etree._Element, order: tuple[str, ...]) -> list[str]:
    """Names of children that appear before a sibling they should follow.

    Reports the *later* member of each violating pair, i.e. the child that is
    sitting too early relative to what precedes it. Unknown and foreign
    children are ignored — they carry no ordering constraint we can check.
    """
    misplaced: list[str] = []
    highest = -1
    for child in parent:
        name = _local(child)
        if name is None or name not in order:
            continue
        position = order.index(name)
        if position < highest:
            misplaced.append(name)
        else:
            highest = position
    return misplaced
