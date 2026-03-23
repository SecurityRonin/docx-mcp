"""DocxDocument: mixin composition and public API."""

from .base import (
    A,
    CT,
    CT_TYPES,
    NSMAP,
    R,
    RELS,
    REL_TYPES,
    V,
    W,
    W14,
    W15,
    XML_SPACE,
    BaseMixin,
    _now_iso,
    _preserve,
)
from .comments import CommentsMixin
from .footnotes import FootnotesMixin
from .reading import ReadingMixin
from .tracks import TracksMixin
from .validation import ValidationMixin


class DocxDocument(
    BaseMixin,
    ReadingMixin,
    TracksMixin,
    CommentsMixin,
    FootnotesMixin,
    ValidationMixin,
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
    "NSMAP",
    "REL_TYPES",
    "CT_TYPES",
    "_now_iso",
    "_preserve",
]
