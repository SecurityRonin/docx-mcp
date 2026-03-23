"""DocxDocument: mixin composition and public API."""

from .base import (
    CP,
    CT,
    CT_TYPES,
    DC,
    DCTERMS,
    NSMAP,
    REL_TYPES,
    RELS,
    W14,
    W15,
    WP,
    XML_SPACE,
    A,
    BaseMixin,
    R,
    V,
    W,
    _now_iso,
    _preserve,
)
from .comments import CommentsMixin
from .endnotes import EndnotesMixin
from .footnotes import FootnotesMixin
from .formatting import FormattingMixin
from .headers_footers import HeadersFootersMixin
from .images import ImagesMixin
from .lists import ListsMixin
from .merge import MergeMixin
from .properties import PropertiesMixin
from .protection import ProtectionMixin
from .reading import ReadingMixin
from .references import ReferencesMixin
from .sections import SectionsMixin
from .styles import StylesMixin
from .tables import TablesMixin
from .tracks import TracksMixin
from .validation import ValidationMixin


class DocxDocument(
    BaseMixin,
    ReadingMixin,
    TracksMixin,
    FormattingMixin,
    CommentsMixin,
    FootnotesMixin,
    ValidationMixin,
    TablesMixin,
    StylesMixin,
    HeadersFootersMixin,
    ListsMixin,
    PropertiesMixin,
    ImagesMixin,
    EndnotesMixin,
    SectionsMixin,
    ReferencesMixin,
    ProtectionMixin,
    MergeMixin,
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
