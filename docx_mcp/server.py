"""docx-mcp: MCP server for Word document editing.

Provides tools for reading and editing .docx files with full support for
track changes (w:ins/w:del), comments, footnotes, and structural validation.
"""

from __future__ import annotations

import json

from mcp.server.fastmcp import FastMCP

from docx_mcp.document import DocxDocument

mcp = FastMCP(
    "docx-mcp",
    instructions=(
        "This server edits Word (.docx) documents. Open a document first with "
        "open_document, then use other tools to read, edit, and save. Changes are "
        "made with proper Word track-changes markup so they appear as revisions "
        "in Microsoft Word / LibreOffice."
    ),
)

_doc: DocxDocument | None = None


def _js(obj: object) -> str:
    """Serialize to compact JSON for MCP responses."""
    return json.dumps(obj, indent=2, ensure_ascii=False)


def _require_doc() -> DocxDocument:
    if _doc is None:
        raise RuntimeError("No document is open. Call open_document first.")
    return _doc


# ── Document lifecycle ──────────────────────────────────────────────────────


@mcp.tool()
def open_document(path: str) -> str:
    """Open a .docx file for reading and editing.

    Unpacks the DOCX archive, parses all XML parts, and caches them in memory.
    Only one document can be open at a time; opening a new one closes the previous.

    Args:
        path: Absolute path to the .docx file.
    """
    global _doc
    if _doc is not None:
        _doc.close()
    _doc = DocxDocument(path)
    info = _doc.open()
    return _js(info)


@mcp.tool()
def close_document() -> str:
    """Close the current document and clean up temporary files."""
    global _doc
    if _doc is not None:
        _doc.close()
        _doc = None
    return "Document closed."


@mcp.tool()
def get_document_info() -> str:
    """Get overview stats: paragraph count, headings, footnotes, comments, images."""
    return _js(_require_doc().get_info())


# ── Reading ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_headings() -> str:
    """Get the document heading structure with levels, text, and paraIds.

    Returns a list of headings in document order, each with:
    - level (1-9)
    - text (heading content)
    - style (e.g., "Heading1")
    - paraId (unique paragraph identifier for targeting edits)
    """
    return _js(_require_doc().get_headings())


@mcp.tool()
def search_text(query: str, regex: bool = False) -> str:
    """Search for text across the document body, footnotes, and comments.

    Args:
        query: Text to search for (case-insensitive), or a regex pattern.
        regex: If true, treat query as a Python regular expression.

    Returns matching paragraphs with their paraId, source part, and context.
    """
    return _js(_require_doc().search_text(query, regex=regex))


@mcp.tool()
def get_paragraph(para_id: str) -> str:
    """Get the full text and style of a specific paragraph by its paraId.

    Args:
        para_id: The 8-character hex paraId (e.g., "1A2B3C4D").
    """
    return _js(_require_doc().get_paragraph(para_id))


# ── Tables ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_tables() -> str:
    """Get all tables with row/column counts and cell text content."""
    return _js(_require_doc().get_tables())


# ── Styles ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_styles() -> str:
    """Get all defined styles with ID, name, type, and base style."""
    return _js(_require_doc().get_styles())


# ── Headers / Footers ──────────────────────────────────────────────────────


@mcp.tool()
def get_headers_footers() -> str:
    """Get all headers and footers with their text content."""
    return _js(_require_doc().get_headers_footers())


# ── Properties ─────────────────────────────────────────────────────────────


@mcp.tool()
def get_properties() -> str:
    """Get core document properties (title, creator, subject, dates, revision)."""
    return _js(_require_doc().get_properties())


# ── Images ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_images() -> str:
    """Get all embedded images with rId, filename, content type, and dimensions."""
    return _js(_require_doc().get_images())


# ── Endnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_endnotes() -> str:
    """Get all endnotes with their ID and text content."""
    return _js(_require_doc().get_endnotes())


# ── Footnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_footnotes() -> str:
    """List all footnotes with their ID and text content."""
    return _js(_require_doc().get_footnotes())


@mcp.tool()
def add_footnote(para_id: str, text: str) -> str:
    """Add a footnote to a paragraph.

    Creates the footnote definition in footnotes.xml and adds a superscript
    reference number at the end of the target paragraph.

    Args:
        para_id: paraId of the paragraph to attach the footnote to.
        text: Footnote text content.
    """
    return _js(_require_doc().add_footnote(para_id, text))


@mcp.tool()
def validate_footnotes() -> str:
    """Cross-reference footnote IDs between document.xml and footnotes.xml.

    Checks that every footnote reference in the document body has a matching
    definition, and flags orphaned definitions with no reference.
    """
    return _js(_require_doc().validate_footnotes())


# ── Validation ──────────────────────────────────────────────────────────────


@mcp.tool()
def validate_paraids() -> str:
    """Check paraId uniqueness across all document parts.

    ParaIds must be unique across document.xml, footnotes.xml, headers, footers,
    and comments. They must also be valid 8-digit hex values < 0x80000000.
    """
    return _js(_require_doc().validate_paraids())


@mcp.tool()
def remove_watermark() -> str:
    """Remove VML watermarks (e.g., DRAFT) from all document headers.

    Detects and removes <v:shape> elements with <v:textpath> inside header XML
    files — the standard pattern for Word watermarks.
    """
    return _js(_require_doc().remove_watermark())


@mcp.tool()
def audit_document() -> str:
    """Run a comprehensive structural audit of the document.

    Checks:
    - Footnote cross-references (references vs definitions)
    - ParaId uniqueness and range validity
    - Heading level continuity (no skips like H2 -> H4)
    - Bookmark pairing (start/end matching)
    - Relationship targets (all referenced files exist)
    - Image references (all embedded images exist)
    - Residual artifacts (DRAFT, TODO, FIXME markers)

    Returns a detailed report with an overall valid/invalid status.
    """
    return _js(_require_doc().audit())


# ── Track changes ───────────────────────────────────────────────────────────


@mcp.tool()
def insert_text(
    para_id: str,
    text: str,
    position: str = "end",
    author: str = "Claude",
) -> str:
    """Insert text with Word track-changes markup (appears as a green underlined insertion in Word).

    Args:
        para_id: paraId of the target paragraph.
        text: Text to insert.
        position: Where to insert — "start", "end", or a substring to insert after.
        author: Author name for the revision (shown in Word's review pane).
    """
    return _js(_require_doc().insert_text(para_id, text, position=position, author=author))


@mcp.tool()
def delete_text(
    para_id: str,
    text: str,
    author: str = "Claude",
) -> str:
    """Mark text as deleted with Word track-changes markup (appears as red strikethrough in Word).

    Finds the exact text within the paragraph and wraps it in deletion markup.
    The text must exist within a single run (formatting span).

    Args:
        para_id: paraId of the target paragraph.
        text: Exact text to mark as deleted.
        author: Author name for the revision.
    """
    return _js(_require_doc().delete_text(para_id, text, author=author))


# ── Comments ────────────────────────────────────────────────────────────────


@mcp.tool()
def get_comments() -> str:
    """List all comments with their ID, author, date, and text."""
    return _js(_require_doc().get_comments())


@mcp.tool()
def add_comment(
    para_id: str,
    text: str,
    author: str = "Claude",
) -> str:
    """Add a comment anchored to a paragraph.

    Creates the comment in comments.xml and adds range markers
    (commentRangeStart/End) around the paragraph content.

    Args:
        para_id: paraId of the paragraph to comment on.
        text: Comment text.
        author: Author name (shown in Word's comment sidebar).
    """
    return _js(_require_doc().add_comment(para_id, text, author=author))


@mcp.tool()
def reply_to_comment(
    parent_id: int,
    text: str,
    author: str = "Claude",
) -> str:
    """Reply to an existing comment (creates a threaded reply).

    Args:
        parent_id: ID of the comment to reply to.
        text: Reply text.
        author: Author name.
    """
    return _js(_require_doc().reply_to_comment(parent_id, text, author=author))


# ── Save ────────────────────────────────────────────────────────────────────


@mcp.tool()
def save_document(output_path: str = "") -> str:
    """Save all changes back to a .docx file.

    Serializes modified XML parts and repacks into a DOCX archive.

    Args:
        output_path: Path for the output file. If empty, overwrites the original.
    """
    doc = _require_doc()
    path = output_path if output_path else None
    return _js(doc.save(path))


# ── Entry point ─────────────────────────────────────────────────────────────


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
