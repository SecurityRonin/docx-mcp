"""DocxDocument: mixin composition and public API."""

from .base import (
    A,
    CP,
    CT,
    CT_TYPES,
    DC,
    DCTERMS,
    NSMAP,
    R,
    RELS,
    REL_TYPES,
    V,
    W,
    W14,
    W15,
    WP,
    XML_SPACE,
    BaseMixin,
    _now_iso,
    _preserve,
)
from .comments import CommentsMixin
from .endnotes import EndnotesMixin
from .footnotes import FootnotesMixin
from .headers_footers import HeadersFootersMixin
from .images import ImagesMixin
from .properties import PropertiesMixin
from .reading import ReadingMixin
from .styles import StylesMixin
from .tables import TablesMixin
from .tracks import TracksMixin
from .validation import ValidationMixin


class DocxDocument(
    BaseMixin,
    ReadingMixin,
    TracksMixin,
    CommentsMixin,
    FootnotesMixin,
    ValidationMixin,
    TablesMixin,
    StylesMixin,
    HeadersFootersMixin,
    PropertiesMixin,
    ImagesMixin,
    EndnotesMixin,
):
    """Word document editor with OOXML-level control."""

    pass


__all__ = [
    "DocxDocument",
    "W",
    "W14",
    "W15",
    "R",
    "V",
    "A",
    "CT",
    "RELS",
    "XML_SPACE",
    "WP",
    "DC",
    "DCTERMS",
    "CP",
    "NSMAP",
    "REL_TYPES",
    "CT_TYPES",
    "_now_iso",
    "_preserve",
]
