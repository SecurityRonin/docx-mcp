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
def create_document(
    output_path: str,
    template_path: str | None = None,
) -> str:
    """Create a new blank .docx document (or from a .dotx template).

    The document is automatically opened for editing after creation.
    Use save_document to save changes, or start editing immediately
    with insert_text, add_table, etc.

    Args:
        output_path: Path for the new .docx file.
        template_path: Optional path to a .dotx template file.
    """
    global _doc
    if _doc is not None:
        _doc.close()
    _doc = DocxDocument.create(output_path, template_path=template_path)
    return _js(_doc.get_info())


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


@mcp.tool()
def add_table(
    para_id: str,
    rows: int,
    cols: int,
    author: str = "Claude",
) -> str:
    """Insert a new table after a paragraph with tracked insertion.

    Args:
        para_id: paraId of the paragraph to insert after.
        rows: Number of rows.
        cols: Number of columns.
        author: Author name for the revision.
    """
    return _js(_require_doc().add_table(para_id, rows, cols, author=author))


@mcp.tool()
def modify_cell(
    table_idx: int,
    row: int,
    col: int,
    text: str,
    author: str = "Claude",
) -> str:
    """Modify a table cell with tracked changes (delete old, insert new).

    Args:
        table_idx: Table index (0-based).
        row: Row index (0-based).
        col: Column index (0-based).
        text: New cell text.
        author: Author name for the revision.
    """
    return _js(_require_doc().modify_cell(table_idx, row, col, text, author=author))


@mcp.tool()
def add_table_row(
    table_idx: int,
    row_idx: int = -1,
    cells: list[str] | None = None,
    author: str = "Claude",
) -> str:
    """Add a row to a table with tracked insertion.

    Args:
        table_idx: Table index (0-based).
        row_idx: Insert at this row index. -1 = append at end.
        cells: Cell text content. Empty = empty cells.
        author: Author name for the revision.
    """
    idx = row_idx if row_idx >= 0 else None
    return _js(_require_doc().add_table_row(table_idx, row_idx=idx, cells=cells, author=author))


@mcp.tool()
def delete_table_row(
    table_idx: int,
    row_idx: int,
    author: str = "Claude",
) -> str:
    """Delete a table row with tracked changes.

    Args:
        table_idx: Table index (0-based).
        row_idx: Row index to delete (0-based).
        author: Author name for the revision.
    """
    return _js(_require_doc().delete_table_row(table_idx, row_idx, author=author))


# ── Lists ──────────────────────────────────────────────────────────────────


@mcp.tool()
def add_list(
    para_ids: list[str],
    style: str = "bullet",
) -> str:
    """Apply list formatting to paragraphs (bullet or numbered).

    Creates numbering definitions and sets w:numPr on each target paragraph.

    Args:
        para_ids: List of paraIds to format as list items.
        style: "bullet" or "numbered".
    """
    return _js(_require_doc().add_list(para_ids, style=style))


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


@mcp.tool()
def edit_header_footer(
    location: str,
    old_text: str,
    new_text: str,
    author: str = "Claude",
) -> str:
    """Edit text in a header or footer with tracked changes.

    Args:
        location: "header" or "footer" (matches first found).
        old_text: Text to find and replace.
        new_text: Replacement text.
        author: Author name for the revision.
    """
    return _js(_require_doc().edit_header_footer(location, old_text, new_text, author=author))


# ── Properties ─────────────────────────────────────────────────────────────


@mcp.tool()
def get_properties() -> str:
    """Get core document properties (title, creator, subject, dates, revision)."""
    return _js(_require_doc().get_properties())


@mcp.tool()
def set_properties(
    title: str = "",
    creator: str = "",
    subject: str = "",
    description: str = "",
) -> str:
    """Set core document properties.

    Args:
        title: Document title. Empty = unchanged.
        creator: Document author/creator. Empty = unchanged.
        subject: Document subject. Empty = unchanged.
        description: Document description. Empty = unchanged.
    """
    return _js(
        _require_doc().set_properties(
            title=title or None,
            creator=creator or None,
            subject=subject or None,
            description=description or None,
        )
    )


# ── Images ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_images() -> str:
    """Get all embedded images with rId, filename, content type, and dimensions."""
    return _js(_require_doc().get_images())


@mcp.tool()
def insert_image(
    para_id: str,
    image_path: str,
    width_emu: int = 2000000,
    height_emu: int = 2000000,
) -> str:
    """Insert an image into the document after a paragraph.

    Args:
        para_id: paraId of the paragraph to insert after.
        image_path: Absolute path to the image file.
        width_emu: Image width in EMUs (914400 = 1 inch).
        height_emu: Image height in EMUs.
    """
    return _js(
        _require_doc().insert_image(para_id, image_path, width_emu=width_emu, height_emu=height_emu)
    )


# ── Endnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_endnotes() -> str:
    """Get all endnotes with their ID and text content."""
    return _js(_require_doc().get_endnotes())


@mcp.tool()
def add_endnote(para_id: str, text: str) -> str:
    """Add an endnote to a paragraph.

    Creates the endnote definition in endnotes.xml and adds a superscript
    reference in the target paragraph.

    Args:
        para_id: paraId of the paragraph to attach the endnote to.
        text: Endnote text content.
    """
    return _js(_require_doc().add_endnote(para_id, text))


@mcp.tool()
def validate_endnotes() -> str:
    """Cross-reference endnote IDs between document.xml and endnotes.xml.

    Checks that every endnote reference has a matching definition,
    and flags orphaned definitions with no reference.
    """
    return _js(_require_doc().validate_endnotes())


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


# ── Sections / Page breaks ─────────────────────────────────────────────


@mcp.tool()
def add_page_break(para_id: str) -> str:
    """Insert a page break after a paragraph.

    Creates a new paragraph containing a page break element.

    Args:
        para_id: paraId of the paragraph to insert after.
    """
    return _js(_require_doc().add_page_break(para_id))


@mcp.tool()
def add_section_break(
    para_id: str,
    break_type: str = "nextPage",
) -> str:
    """Add a section break at a paragraph.

    Inserts a section properties element in the paragraph, marking it as
    the last paragraph of its section.

    Args:
        para_id: paraId of the paragraph to place the section break on.
        break_type: "nextPage", "continuous", "evenPage", or "oddPage".
    """
    return _js(_require_doc().add_section_break(para_id, break_type=break_type))


@mcp.tool()
def set_section_properties(
    para_id: str = "",
    width: int = 0,
    height: int = 0,
    orientation: str = "",
    margin_top: int = 0,
    margin_bottom: int = 0,
    margin_left: int = 0,
    margin_right: int = 0,
) -> str:
    """Modify section properties (page size, orientation, margins).

    Args:
        para_id: paraId of paragraph with section break. Empty = body section.
        width: Page width in DXA (twips). 12240 = 8.5 inches.
        height: Page height in DXA. 15840 = 11 inches.
        orientation: "portrait" or "landscape". Empty = unchanged.
        margin_top: Top margin in DXA. 0 = unchanged.
        margin_bottom: Bottom margin in DXA. 0 = unchanged.
        margin_left: Left margin in DXA. 0 = unchanged.
        margin_right: Right margin in DXA. 0 = unchanged.
    """
    return _js(
        _require_doc().set_section_properties(
            para_id=para_id or None,
            width=width or None,
            height=height or None,
            orientation=orientation or None,
            margin_top=margin_top or None,
            margin_bottom=margin_bottom or None,
            margin_left=margin_left or None,
            margin_right=margin_right or None,
        )
    )


# ── Cross-references ──────────────────────────────────────────────────


@mcp.tool()
def add_cross_reference(
    source_para_id: str,
    target_para_id: str,
    text: str,
) -> str:
    """Add a cross-reference link from one paragraph to another.

    Creates a bookmark at the target (if none exists) and inserts a hyperlink
    in the source paragraph with the given display text.

    Args:
        source_para_id: paraId of the paragraph where the link appears.
        target_para_id: paraId of the paragraph being referenced.
        text: Display text for the cross-reference link.
    """
    return _js(_require_doc().add_cross_reference(source_para_id, target_para_id, text))


# ── Protection ─────────────────────────────────────────────────────────────


@mcp.tool()
def set_document_protection(
    edit: str,
    password: str = "",
) -> str:
    """Set document protection in settings.xml.

    Args:
        edit: Protection type — "trackedChanges", "comments", "readOnly",
              "forms", or "none" (removes protection).
        password: Optional password (hashed with SHA-512). Empty = no password.
    """
    return _js(_require_doc().set_document_protection(edit, password=password or None))


# ── Merge ──────────────────────────────────────────────────────────────────


@mcp.tool()
def merge_documents(source_path: str) -> str:
    """Merge another DOCX document's content into the current document.

    Appends body paragraphs and tables from the source. ParaIds are
    automatically remapped to avoid collisions.

    Args:
        source_path: Absolute path to the DOCX file to merge in.
    """
    return _js(_require_doc().merge_documents(source_path))


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


@mcp.tool()
def accept_changes(author: str = "") -> str:
    """Accept tracked changes — keep insertions, remove deletions.

    Args:
        author: If set, only accept changes by this author. Empty = all changes.
    """
    return _js(_require_doc().accept_changes(author=author or None))


@mcp.tool()
def reject_changes(author: str = "") -> str:
    """Reject tracked changes — remove insertions, restore deleted text.

    Args:
        author: If set, only reject changes by this author. Empty = all changes.
    """
    return _js(_require_doc().reject_changes(author=author or None))


@mcp.tool()
def set_formatting(
    para_id: str,
    text: str,
    bold: bool = False,
    italic: bool = False,
    underline: str = "",
    color: str = "",
    author: str = "Claude",
) -> str:
    """Apply character formatting to text with tracked-change markup.

    Finds the text within the paragraph, splits the run if needed, and applies
    formatting with rPrChange so it appears as a format revision in Word.

    Args:
        para_id: paraId of the target paragraph.
        text: Exact text to format.
        bold: Apply bold formatting.
        italic: Apply italic formatting.
        underline: Underline style (e.g., "single", "double"). Empty = no underline.
        color: Font color as hex (e.g., "FF0000"). Empty = no color change.
        author: Author name for the revision.
    """
    return _js(
        _require_doc().set_formatting(
            para_id,
            text,
            bold=bold,
            italic=italic,
            underline=underline or None,
            color=color or None,
            author=author,
        )
    )


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
