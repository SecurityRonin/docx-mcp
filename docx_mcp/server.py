"""docx-mcp: MCP server for Word document editing.

Provides tools for reading and editing .docx files with full support for
track changes (w:ins/w:del), comments, footnotes, and structural validation.

Documents are keyed by an opaque `document_handle`. Clients that run multiple
concurrent editing sessions against one server process (e.g. parallel agents)
MUST pass a unique handle per session. Legacy callers that omit the handle
transparently share a single `__default__` slot — same behavior as before.
"""

from __future__ import annotations

import json
import sys as _sys
import uuid

from mcp.server.fastmcp import FastMCP

from docx_mcp.document import DocxDocument

mcp = FastMCP(
    "docx-mcp",
    instructions=(
        "This server edits Word (.docx) documents. Open a document first with "
        "open_document, then use other tools to read, edit, and save. Changes are "
        "made with proper Word track-changes markup so they appear as revisions "
        "in Microsoft Word / LibreOffice.\n\n"
        "PARALLEL SESSIONS: every tool accepts an optional `document_handle` "
        "string. open_document/create_document/create_from_markdown return the "
        "handle under which the document is stored. If you run multiple "
        "concurrent editing flows, pass a unique handle (e.g. a UUID) to each "
        "tool call so sessions don't collide. Omitting the handle uses a shared "
        "`__default__` slot — fine for a single caller, unsafe for parallel ones."
    ),
)

_DEFAULT_HANDLE = "__default__"
_docs: dict[str, DocxDocument] = {}


def _js(obj: object) -> str:
    """Serialize to compact JSON for MCP responses."""
    return json.dumps(obj, indent=2, ensure_ascii=False)


def _key(handle: str) -> str:
    return handle or _DEFAULT_HANDLE


def _resolve(handle: str) -> tuple[str, DocxDocument]:
    """Return (key, doc) for a handle, or raise if no document is open there."""
    key = _key(handle)
    doc = _docs.get(key)
    if doc is None:
        raise RuntimeError(f"No document is open for handle {key!r}. Call open_document first.")
    return key, doc


def _store(handle: str, doc: DocxDocument) -> str:
    """Place a document under a handle, closing any previous occupant. Returns the key."""
    key = _key(handle) if handle else uuid.uuid4().hex
    existing = _docs.get(key)
    if existing is not None:
        existing.close()
    _docs[key] = doc
    return key


def _require_doc() -> DocxDocument:
    """Compat shim: return the __default__ slot, raising if nothing is open."""
    return _resolve(_DEFAULT_HANDLE)[1]


# ── Document lifecycle ──────────────────────────────────────────────────────


@mcp.tool()
def open_document(path: str, document_handle: str = "") -> str:
    """Open a .docx file for reading and editing.

    Unpacks the DOCX archive, parses all XML parts, and caches them in memory
    under a handle. Multiple documents may be open concurrently, each under
    its own handle.

    Args:
        path: Absolute path to the .docx file.
        document_handle: Optional handle to store this document under. Empty
            string uses the shared `__default__` slot (legacy behavior); pass
            a unique value (e.g. a UUID) per concurrent session for isolation.
    """
    doc = DocxDocument(path)
    info = doc.open()
    key = _store(document_handle or _DEFAULT_HANDLE, doc)
    out: dict[str, object] = {"handle": key}
    if isinstance(info, dict):
        out.update(info)
    else:
        out["info"] = info
    return _js(out)


@mcp.tool()
def close_document(document_handle: str = "") -> str:
    """Close a document and clean up temporary files.

    Args:
        document_handle: Handle of the document to close. Empty = `__default__` slot.
    """
    key = _key(document_handle)
    doc = _docs.pop(key, None)
    if doc is not None:
        doc.close()
        return f"Document {key!r} closed."
    return f"No document open for handle {key!r}."


@mcp.tool()
def create_document(
    output_path: str,
    template_path: str | None = None,
    document_handle: str = "",
) -> str:
    """Create a new blank .docx document (or from a .dotx template).

    The document is automatically opened for editing under the given handle.
    Use save_document to save changes, or start editing immediately with
    insert_text, add_table, etc.

    Args:
        output_path: Path for the new .docx file.
        template_path: Optional path to a .dotx template file.
        document_handle: Optional handle to store this document under.
    """
    doc = DocxDocument.create(output_path, template_path=template_path)
    key = _store(document_handle or _DEFAULT_HANDLE, doc)
    info = doc.get_info()
    out: dict[str, object] = {"handle": key}
    if isinstance(info, dict):
        out.update(info)
    else:
        out["info"] = info
    return _js(out)


@mcp.tool()
def create_from_markdown(
    output_path: str,
    md_path: str | None = None,
    markdown: str | None = None,
    template_path: str | None = None,
    document_handle: str = "",
) -> str:
    """Create a new .docx document from markdown content.

    Supports full GitHub-Flavored Markdown: headings, bold/italic/strikethrough,
    links, images, bullet/numbered/nested lists, code blocks, blockquotes,
    tables, footnotes, and task lists. Smart typography (curly quotes, em/en
    dashes, ellipses) is applied automatically.

    Provide exactly one of md_path or markdown. The document is automatically
    opened for editing under the given handle.

    Args:
        output_path: Path for the new .docx file.
        md_path: Path to a .md file. Mutually exclusive with markdown.
        markdown: Raw markdown text. Mutually exclusive with md_path.
        template_path: Optional path to a .dotx template file.
        document_handle: Optional handle to store this document under.
    """
    if md_path and markdown:
        return "Error: md_path and markdown are mutually exclusive — provide one, not both."
    if not md_path and not markdown:
        return "Error: Either md_path or markdown must be provided."

    doc = DocxDocument.create(output_path, template_path=template_path)

    base_dir = None
    if md_path:
        from pathlib import Path

        p = Path(md_path)
        if not p.exists():
            doc.close()
            return f"Error: Markdown file not found: {md_path}"
        markdown = p.read_text(encoding="utf-8")
        base_dir = p.parent

    from docx_mcp.markdown import MarkdownConverter

    MarkdownConverter.convert(doc, markdown, base_dir=base_dir)
    doc.save(backup=False)
    key = _store(document_handle or _DEFAULT_HANDLE, doc)
    info = doc.get_info()
    out: dict[str, object] = {"handle": key}
    if isinstance(info, dict):
        out.update(info)
    else:
        out["info"] = info
    return _js(out)


@mcp.tool()
def get_document_info(document_handle: str = "") -> str:
    """Get overview stats: paragraph count, headings, footnotes, comments, images."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_info())


# ── Reading ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_headings(document_handle: str = "") -> str:
    """Get the document heading structure with levels, text, and paraIds.

    Returns a list of headings in document order, each with:
    - level (1-9)
    - text (heading content)
    - style (e.g., "Heading1")
    - paraId (unique paragraph identifier for targeting edits)
    """
    _, doc = _resolve(document_handle)
    return _js(doc.get_headings())


@mcp.tool()
def search_text(query: str, regex: bool = False, document_handle: str = "") -> str:
    """Search for text across the document body, footnotes, and comments.

    Args:
        query: Text to search for (case-insensitive), or a regex pattern.
        regex: If true, treat query as a Python regular expression.

    Returns matching paragraphs with their paraId, source part, and context.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.search_text(query, regex=regex))


@mcp.tool()
def get_paragraph(para_id: str, document_handle: str = "") -> str:
    """Get the full text and style of a specific paragraph by its paraId.

    Args:
        para_id: The 8-character hex paraId (e.g., "1A2B3C4D").
    """
    _, doc = _resolve(document_handle)
    return _js(doc.get_paragraph(para_id))


# ── Tables ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_tables(document_handle: str = "") -> str:
    """Get all tables with row/column counts and cell text content."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_tables())


@mcp.tool()
def add_table(
    para_id: str,
    rows: int,
    cols: int,
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Insert a new table after a paragraph with tracked insertion."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_table(para_id, rows, cols, author=author))


@mcp.tool()
def modify_cell(
    table_idx: int,
    row: int,
    col: int,
    text: str,
    author: str = "Claude",
    tracked: bool = True,
    document_handle: str = "",
) -> str:
    """Modify a table cell.

    By default (tracked=True) the old content is marked as a deletion and the new
    content as an insertion — the human reviewer accepts/rejects in Word's Track
    Changes view. Pass tracked=False to overwrite the cell directly with no markup.

    Args:
        table_idx: Table index (0-based).
        row: Row index (0-based).
        col: Column index (0-based).
        text: New cell text.
        author: Author name shown in Word's review pane (tracked=True only).
        tracked: True (default) = tracked del+ins. False = direct overwrite, no markup.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.modify_cell(table_idx, row, col, text, author=author, tracked=tracked))


@mcp.tool()
def add_table_row(
    table_idx: int,
    row_idx: int = -1,
    cells: list[str] | None = None,
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Add a row to a table with tracked insertion. row_idx=-1 appends."""
    _, doc = _resolve(document_handle)
    idx = row_idx if row_idx >= 0 else None
    return _js(doc.add_table_row(table_idx, row_idx=idx, cells=cells, author=author))


@mcp.tool()
def delete_table_row(
    table_idx: int,
    row_idx: int,
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Delete a table row with tracked changes."""
    _, doc = _resolve(document_handle)
    return _js(doc.delete_table_row(table_idx, row_idx, author=author))


@mcp.tool()
def merge_cells(
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> str:
    """Merge a rectangular range of cells. Horizontal: gridSpan. Vertical: vMerge."""
    doc = _require_doc()
    return _js(doc.merge_cells(table_index, start_row, start_col, end_row, end_col))


@mcp.tool()
def set_header_row(table_index: int) -> str:
    """Mark the first row as a repeating header row."""
    doc = _require_doc()
    return _js(doc.set_header_row(table_index))


@mcp.tool()
def set_column_widths(table_index: int, widths_cm: list[float]) -> str:
    """Set column widths in cm. len(widths_cm) must match column count."""
    doc = _require_doc()
    return _js(doc.set_column_widths(table_index, widths_cm))


@mcp.tool()
def csv_to_table(para_id: str, csv_text: str, header_row: bool = True) -> str:
    """Insert a table from CSV text."""
    doc = _require_doc()
    return _js(doc.csv_to_table(para_id, csv_text, header_row))


@mcp.tool()
def table_to_csv(table_index: int) -> str:
    """Export a table as CSV string."""
    doc = _require_doc()
    return _js(doc.table_to_csv(table_index))


@mcp.tool()
def delete_table(table_idx: int) -> str:
    """Delete a table by index (0-based). Raises IndexError if out of range."""
    return _js(_require_doc().delete_table(table_idx))


@mcp.tool()
def add_column_to_table(table_idx: int, header_text: str = "") -> str:
    """Add a new column to every row of a table. First row gets header_text."""
    return _js(_require_doc().add_column_to_table(table_idx, header_text=header_text))


@mcp.tool()
def delete_column_from_table(table_idx: int, col_idx: int) -> str:
    """Delete a column (0-based) from every row of a table."""
    return _js(_require_doc().delete_column_from_table(table_idx, col_idx))


@mcp.tool()
def set_cell_width(table_idx: int, row_idx: int, col_idx: int, width_mm: float) -> str:
    """Set the width of a table cell in millimetres (stored as DXA)."""
    return _js(_require_doc().set_cell_width(table_idx, row_idx, col_idx, width_mm))


@mcp.tool()
def set_cell_vertical_alignment(table_idx: int, row_idx: int, col_idx: int, alignment: str) -> str:
    """Set vertical alignment of a table cell: top, center, or bottom."""
    return _js(_require_doc().set_cell_vertical_alignment(table_idx, row_idx, col_idx, alignment))


@mcp.tool()
def set_row_height(table_idx: int, row_idx: int, height_mm: float, rule: str = "exact") -> str:
    """Set row height in millimetres. rule: exact, atLeast, or auto."""
    return _js(_require_doc().set_row_height(table_idx, row_idx, height_mm, rule=rule))


@mcp.tool()
def set_table_alignment(table_idx: int, alignment: str) -> str:
    """Set table alignment: left, center, or right."""
    return _js(_require_doc().set_table_alignment(table_idx, alignment))


@mcp.tool()
def set_table_borders(
    table_idx: int,
    border_style: str = "single",
    color: str = "000000",
    size: int = 4,
) -> str:
    """Set borders on all six sides of a table (top, bottom, left, right, insideH, insideV)."""
    return _js(
        _require_doc().set_table_borders(
            table_idx, border_style=border_style, color=color, size=size
        )
    )  # noqa: E501


@mcp.tool()
def set_cell_shading(
    table_idx: int,
    row_idx: int,
    col_idx: int,
    fill_color: str,
    pattern: str = "clear",
) -> str:
    """Set background shading fill color on a table cell."""
    return _js(
        _require_doc().set_cell_shading(table_idx, row_idx, col_idx, fill_color, pattern=pattern)
    )  # noqa: E501


@mcp.tool()
def set_table_style(table_idx: int, style_name: str) -> str:
    """Apply a named table style (e.g. TableGrid, LightShading-Accent1) to a table."""
    return _js(_require_doc().set_table_style(table_idx, style_name))


# ── Lists ──────────────────────────────────────────────────────────────────


@mcp.tool()
def add_list(
    para_ids: list[str],
    style: str = "bullet",
    document_handle: str = "",
) -> str:
    """Apply list formatting to paragraphs (bullet or numbered)."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_list(para_ids, style=style))


# ── Styles ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_styles(document_handle: str = "") -> str:
    """Get all defined styles with ID, name, type, and base style."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_styles())


@mcp.tool()
def create_style(
    name: str,
    style_type: str,
    based_on: str | None = None,
    next_style: str | None = None,
) -> str:
    """Create a new style in the document.

    Args:
        name: Style name (used as styleId after removing spaces).
        style_type: "paragraph", "character", "table", or "numbering".
        based_on: Optional styleId this style inherits from.
        next_style: Optional styleId applied to the next paragraph.
    """
    return _js(
        _require_doc().create_style(name, style_type, based_on=based_on, next_style=next_style)
    )  # noqa: E501


@mcp.tool()
def update_style(
    name: str,
    based_on: str | None = None,
    next_style: str | None = None,
) -> str:
    """Update an existing style's basedOn and/or next properties.

    Args:
        name: Style name or styleId (case-insensitive).
        based_on: New basedOn styleId (replaces existing).
        next_style: New next styleId (replaces existing).
    """
    return _js(_require_doc().update_style(name, based_on=based_on, next_style=next_style))


@mcp.tool()
def delete_style(name: str) -> str:
    """Delete a style from the document.

    Args:
        name: Style name or styleId (case-insensitive).
    """
    return _js(_require_doc().delete_style(name))


@mcp.tool()
def get_style(name_or_id: str) -> str:
    """Get details of a single style by name or styleId (case-insensitive).

    Args:
        name_or_id: Style name or styleId to look up.

    Returns:
        {"style_id": str, "name": str, "type": str, "base_style": str, "next_style": str}
    """
    return _js(_require_doc().get_style(name_or_id))


@mcp.tool()
def copy_style(source_name_or_id: str, new_name: str) -> str:
    """Deep-copy an existing style under a new name.

    Args:
        source_name_or_id: Name or styleId of the style to copy.
        new_name: Name for the new style (spaces stripped for styleId).

    Returns:
        {"style_id": str, "name": str, "type": str}
    """
    return _js(_require_doc().copy_style(source_name_or_id, new_name))


@mcp.tool()
def apply_style_to_range(para_ids: list[str], style_name_or_id: str) -> str:
    """Apply a style to a list of paragraphs by their paraIds.

    Args:
        para_ids: List of paragraph paraIds to update.
        style_name_or_id: Style name or styleId to apply.

    Returns:
        {"applied": int, "style_id": str, "para_ids": list[str]}
    """
    return _js(_require_doc().apply_style_to_range(para_ids, style_name_or_id))


# ── Headers / Footers ──────────────────────────────────────────────────────


@mcp.tool()
def get_headers_footers(document_handle: str = "") -> str:
    """Get all headers and footers with their text content."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_headers_footers())


@mcp.tool()
def edit_header_footer(
    location: str,
    old_text: str,
    new_text: str,
    author: str = "Claude",
    tracked: bool = True,
    document_handle: str = "",
) -> str:
    """Edit text in a header or footer.

    By default (tracked=True) the change is recorded as a deletion of the old text
    and an insertion of the new text — the human reviewer accepts/rejects in Word.
    Pass tracked=False to replace the text directly with no revision markup.

    Args:
        location: "header" or "footer" (matches the first found of that type).
        old_text: Text to find and replace.
        new_text: Replacement text.
        author: Author name shown in Word's review pane (tracked=True only).
        tracked: True (default) = tracked del+ins. False = direct replacement, no markup.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.edit_header_footer(location, old_text, new_text, author=author, tracked=tracked))


@mcp.tool()
def delete_header(location: str = "default") -> str:
    """Delete a header by location: default, first, or even."""
    return _js(_require_doc().delete_header(location=location))


@mcp.tool()
def delete_footer(location: str = "default") -> str:
    """Delete a footer by location: default, first, or even."""
    return _js(_require_doc().delete_footer(location=location))


# ── Properties ─────────────────────────────────────────────────────────────


@mcp.tool()
def get_properties(document_handle: str = "") -> str:
    """Get core document properties (title, creator, subject, dates, revision)."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_properties())


@mcp.tool()
def set_properties(
    title: str = "",
    creator: str = "",
    subject: str = "",
    description: str = "",
    document_handle: str = "",
) -> str:
    """Set core document properties. Empty string = unchanged."""
    _, doc = _resolve(document_handle)
    return _js(
        doc.set_properties(
            title=title or None,
            creator=creator or None,
            subject=subject or None,
            description=description or None,
        )
    )


@mcp.tool()
def get_custom_properties() -> str:
    """Get custom document properties from docProps/custom.xml."""
    return _js(_require_doc().get_custom_properties())


@mcp.tool()
def set_custom_property(name: str, value: str, vt_type: str = "lpwstr") -> str:
    """Set (upsert) a custom document property.

    Args:
        name: Property name.
        value: Property value as a string.
        vt_type: VT type element name (lpwstr, i4, bool, etc.).
    """
    return _js(_require_doc().set_custom_property(name, value, vt_type=vt_type))


@mcp.tool()
def delete_custom_property(name: str) -> str:
    """Delete a custom document property by name.

    Args:
        name: Property name to delete.
    """
    return _js(_require_doc().delete_custom_property(name))


# ── Images ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_images(document_handle: str = "") -> str:
    """Get all embedded images with rId, filename, content type, and dimensions."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_images())


@mcp.tool()
def insert_image(
    para_id: str,
    image_path: str,
    width_emu: int = 2000000,
    height_emu: int = 2000000,
    document_handle: str = "",
) -> str:
    """Insert an image into the document after a paragraph."""
    _, doc = _resolve(document_handle)
    return _js(doc.insert_image(para_id, image_path, width_emu=width_emu, height_emu=height_emu))


@mcp.tool()
def insert_floating_image(
    para_id: str,
    image_path: str,
    width_cm: float,
    height_cm: float,
    h_pos: float = 0.0,
    v_pos: float = 0.0,
    wrap: str = "square",
) -> str:
    """Insert a floating (anchored) image. wrap: square|topbottom|none."""
    doc = _require_doc()
    return _js(
        doc.insert_floating_image(para_id, image_path, width_cm, height_cm, h_pos, v_pos, wrap)
    )  # noqa: E501


@mcp.tool()
def delete_image(rId: str) -> str:
    """Remove the drawing containing the image with the given rId from the document.

    Also removes the relationship entry from word/_rels/document.xml.rels.

    Args:
        rId: The relationship ID of the image to delete (e.g. 'rId5').
    """
    return _js(_require_doc().delete_image(rId))


@mcp.tool()
def update_image(rId: str, new_image_path: str) -> str:
    """Replace the binary for an existing image in-place.

    The image dimensions and position in the document are preserved.

    Args:
        rId: The relationship ID of the image to replace.
        new_image_path: Absolute path to the new image file on disk.
    """
    return _js(_require_doc().update_image(rId, new_image_path))


@mcp.tool()
def set_image_size(rId: str, width_cm: float, height_cm: float) -> str:
    """Resize an embedded image by updating its EMU extent attributes.

    Args:
        rId: The relationship ID of the image to resize.
        width_cm: New width in centimetres.
        height_cm: New height in centimetres.
    """
    return _js(_require_doc().set_image_size(rId, width_cm, height_cm))


@mcp.tool()
def set_image_alt_text(rId: str, alt_text: str, title: str = "") -> str:
    """Set accessibility alt text and title on an embedded image.

    Args:
        rId: The relationship ID of the image.
        alt_text: Alt text (descr attribute on wp:docPr).
        title: Optional title attribute on wp:docPr.
    """
    return _js(_require_doc().set_image_alt_text(rId, alt_text, title=title))


@mcp.tool()
def set_image_border(rId: str, border_pt: float, color: str = "000000") -> str:
    """Set or remove a border on an embedded image.

    Args:
        rId: The relationship ID of the image (e.g. 'rId6').
        border_pt: Border width in points. Use 0 to remove the border.
        color: RGB hex color string without '#' (default '000000' = black).

    Returns:
        JSON with rId, border_pt, and color fields.
    """
    return _js(_require_doc().set_image_border(rId, border_pt, color))


# ── Endnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_endnotes(document_handle: str = "") -> str:
    """Get all endnotes with their ID and text content."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_endnotes())


@mcp.tool()
def add_endnote(para_id: str, text: str, document_handle: str = "") -> str:
    """Add an endnote to a paragraph."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_endnote(para_id, text))


@mcp.tool()
def validate_endnotes(document_handle: str = "") -> str:
    """Cross-reference endnote IDs between document.xml and endnotes.xml."""
    _, doc = _resolve(document_handle)
    return _js(doc.validate_endnotes())


# ── Footnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_footnotes(document_handle: str = "") -> str:
    """List all footnotes with their ID and text content."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_footnotes())


@mcp.tool()
def add_footnote(para_id: str, text: str, url: str = "", document_handle: str = "") -> str:
    """Add a footnote to a paragraph. url, if provided, is rendered as a hotlink."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_footnote(para_id, text, url=url))


@mcp.tool()
def add_footnote_ref(para_id: str, footnote_id: int, document_handle: str = "") -> str:
    """Add a subsequent reference to an existing footnote without creating a new definition.

    Use when the same source must be cited again in a different paragraph. Inserts a
    hyperlink-wrapped footnoteReference that navigates to the same footnote as the
    original citation. Does not duplicate the footnote definition in footnotes.xml.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.add_footnote_ref(para_id, footnote_id))


@mcp.tool()
def validate_footnotes(document_handle: str = "") -> str:
    """Cross-reference footnote IDs between document.xml and footnotes.xml."""
    _, doc = _resolve(document_handle)
    return _js(doc.validate_footnotes())


@mcp.tool()
def update_footnote(footnote_id: int, text: str) -> str:
    """Update the text of an existing footnote.

    Args:
        footnote_id: The numeric ID of the footnote to update (must be >= 1).
        text: The new text content for the footnote.
    """
    return _js(_require_doc().update_footnote(footnote_id, text))


@mcp.tool()
def delete_footnote(footnote_id: int) -> str:
    """Delete a footnote and its in-body reference.

    Removes the footnote definition from footnotes.xml and removes the
    footnoteReference run from the document body.

    Args:
        footnote_id: The numeric ID of the footnote to delete.
    """
    return _js(_require_doc().delete_footnote(footnote_id))


@mcp.tool()
def update_endnote(endnote_id: int, text: str) -> str:
    """Update the text of an existing endnote.

    Args:
        endnote_id: The numeric ID of the endnote to update (must be >= 1).
        text: The new text content for the endnote.
    """
    return _js(_require_doc().update_endnote(endnote_id, text))


@mcp.tool()
def delete_endnote(endnote_id: int) -> str:
    """Delete an endnote and its in-body reference.

    Removes the endnote definition from endnotes.xml and removes the
    endnoteReference run from the document body.

    Args:
        endnote_id: The numeric ID of the endnote to delete.
    """
    return _js(_require_doc().delete_endnote(endnote_id))


# ── Sections / Page breaks ─────────────────────────────────────────────


@mcp.tool()
def add_page_break(para_id: str, document_handle: str = "") -> str:
    """Insert a page break after a paragraph."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_page_break(para_id))


@mcp.tool()
def add_section_break(
    para_id: str,
    break_type: str = "nextPage",
    document_handle: str = "",
) -> str:
    """Add a section break at a paragraph. break_type: nextPage/continuous/evenPage/oddPage."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_section_break(para_id, break_type=break_type))


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
    document_handle: str = "",
) -> str:
    """Modify section properties (page size, orientation, margins). 0/empty = unchanged."""
    _, doc = _resolve(document_handle)
    return _js(
        doc.set_section_properties(
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


@mcp.tool()
def set_page_size(width_mm: float, height_mm: float, para_id: str | None = None) -> str:
    """Set page size from millimetre values.

    Args:
        width_mm: Page width in mm (e.g. 210 for A4, 215.9 for Letter).
        height_mm: Page height in mm (e.g. 297 for A4, 279.4 for Letter).
        para_id: paraId of paragraph with section break. None = body section.
    """
    return _js(_require_doc().set_page_size(width_mm, height_mm, para_id=para_id))


@mcp.tool()
def set_page_margins(
    top_mm: float | None = None,
    bottom_mm: float | None = None,
    left_mm: float | None = None,
    right_mm: float | None = None,
    para_id: str | None = None,
) -> str:
    """Set page margins from millimetre values.

    Args:
        top_mm: Top margin in mm. None = unchanged.
        bottom_mm: Bottom margin in mm. None = unchanged.
        left_mm: Left margin in mm. None = unchanged.
        right_mm: Right margin in mm. None = unchanged.
        para_id: paraId of paragraph with section break. None = body section.
    """
    return _js(
        _require_doc().set_page_margins(
            top_mm=top_mm,
            bottom_mm=bottom_mm,
            left_mm=left_mm,
            right_mm=right_mm,
            para_id=para_id,
        )
    )


@mcp.tool()
def set_page_orientation(orientation: str, para_id: str | None = None) -> str:
    """Set page orientation, swapping width/height dimensions if needed.

    Args:
        orientation: "portrait" or "landscape".
        para_id: paraId of paragraph with section break. None = body section.
    """
    return _js(_require_doc().set_page_orientation(orientation, para_id=para_id))


@mcp.tool()
def get_sections() -> str:
    """List all sections in the document with their properties.

    Returns a JSON array of section objects, each containing:
    index, break_type, page_width, page_height, orientation, columns,
    margin_top, margin_bottom (all sizes in twips/DXA).
    The final section always has break_type="".
    """
    return _js(_require_doc().get_sections())


@mcp.tool()
def set_section_columns(
    section_index: int,
    num_columns: int,
    equal_width: bool = True,
) -> str:
    """Set the number of columns in a section.

    Args:
        section_index: Zero-based section index (use get_sections to find it).
        num_columns: Number of text columns (1 = single column).
        equal_width: If True, all columns are equal width. Default True.
    """
    return _js(_require_doc().set_section_columns(section_index, num_columns, equal_width))


@mcp.tool()
def delete_section_break(para_id: str) -> str:
    """Remove a section break from a paragraph.

    Removes the w:sectPr from the paragraph's w:pPr. After removal the
    paragraph's content flows into the next section rather than ending one.

    Args:
        para_id: paraId of the paragraph that holds the section break.

    Raises:
        ValueError: If the paragraph has no section break.
    """
    return _js(_require_doc().delete_section_break(para_id))


@mcp.tool()
def set_different_first_page(section_index: int, enabled: bool) -> str:
    """Enable or disable a different first-page header/footer for a section.

    When enabled, the section can have a unique header/footer on its first page,
    separate from the header/footer used on subsequent pages.

    Args:
        section_index: Zero-based section index (use get_sections to find it).
        enabled: True to enable different first page, False to disable.

    Returns:
        {"section_index": int, "different_first_page": bool}
    """
    return _js(_require_doc().set_different_first_page(section_index, enabled))


@mcp.tool()
def set_odd_even_headers(enabled: bool) -> str:
    """Enable or disable different odd/even page headers globally.

    This is a document-level setting stored in word/settings.xml. When enabled,
    even-numbered pages can use a different header/footer from odd-numbered pages.

    Args:
        enabled: True to enable different odd/even headers, False to disable.

    Returns:
        {"odd_even_headers": bool}
    """
    return _js(_require_doc().set_odd_even_headers(enabled))


# ── Cross-references ──────────────────────────────────────────────────


@mcp.tool()
def add_cross_reference(
    source_para_id: str,
    target_para_id: str,
    text: str,
    document_handle: str = "",
) -> str:
    """Add a cross-reference link from one paragraph to another."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_cross_reference(source_para_id, target_para_id, text))


# ── Protection ─────────────────────────────────────────────────────────────


@mcp.tool()
def set_document_protection(
    edit: str,
    password: str = "",
    algorithm: str = "SHA-512",
    spin_count: int = 100000,
    document_handle: str = "",
) -> str:
    """Set document protection. edit: trackedChanges/comments/readOnly/forms/none.

    The password is hashed the way Word verifies it (MS-OE376 §2.15.1.28).
    Note that algorithm truncates passwords to 15 characters. algorithm is one
    of MD5/SHA-1/SHA-256/SHA-384/SHA-512.
    """
    _, doc = _resolve(document_handle)
    return _js(
        doc.set_document_protection(
            edit,
            password=password or None,
            algorithm=algorithm,
            spin_count=spin_count,
        )
    )


# ── Merge ──────────────────────────────────────────────────────────────────


@mcp.tool()
def merge_documents(source_path: str, document_handle: str = "") -> str:
    """Merge another DOCX document's content into the current document."""
    _, doc = _resolve(document_handle)
    return _js(doc.merge_documents(source_path))


# ── Validation ──────────────────────────────────────────────────────────────


@mcp.tool()
def validate_paraids(document_handle: str = "") -> str:
    """Check paraId uniqueness across all document parts."""
    _, doc = _resolve(document_handle)
    return _js(doc.validate_paraids())


@mcp.tool()
def insert_watermark(text: str, diagonal: bool = True, document_handle: str = "") -> str:
    """Insert a VML watermark into the document's default header.

    Places a <v:shape> with a <v:textpath> inside the default header, which is
    the standard Word watermark pattern.

    Args:
        text: Watermark text (e.g. "DRAFT", "CONFIDENTIAL").
        diagonal: If True (default), diagonal orientation; if False, horizontal.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.insert_watermark(text, diagonal=diagonal))


@mcp.tool()
def remove_watermark(document_handle: str = "") -> str:
    """Remove VML watermarks (e.g., DRAFT) from all document headers."""
    _, doc = _resolve(document_handle)
    return _js(doc.remove_watermark())


@mcp.tool()
def audit_document(document_handle: str = "") -> str:
    """Run a comprehensive structural audit of the document."""
    _, doc = _resolve(document_handle)
    return _js(doc.audit())


# ── Track changes ───────────────────────────────────────────────────────────


@mcp.tool()
def insert_text(
    para_id: str,
    text: str,
    position: str = "end",
    author: str = "Claude",
    context_before: str = "",
    context_after: str = "",
    ignore_case: bool = False,
    tracked: bool = True,
    document_handle: str = "",
) -> str:
    """Insert text into a paragraph.

    By default (tracked=True) the insertion is marked as a proposed addition
    that appears underlined in Word's Track Changes view. The human reviewer
    must accept it in Word before it becomes permanent. Pass tracked=False to
    write the text directly with no revision markup.

    Args:
        para_id: paraId of the target paragraph.
        text: Text to insert.
        position: Where to insert — "start", "end", or a substring to insert after.
        author: Author name shown in Word's review pane (tracked=True only).
        context_before: Text immediately before the insertion point (for precise anchoring).
        context_after: Text immediately after the insertion point (for precise anchoring).
        ignore_case: If True, match context_before/context_after case-insensitively.
        tracked: True (default) = revision markup the human accepts/rejects in Word.
                 False = text written directly, no markup.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(
        doc.insert_text(
            para_id,
            text,
            position=position,
            author=author,
            context_before=context_before,
            context_after=context_after,
            ignore_case=ignore_case,
            tracked=tracked,
        )
    )


@mcp.tool()
def delete_text(
    para_id: str,
    text: str,
    author: str = "Claude",
    context_before: str = "",
    context_after: str = "",
    ignore_case: bool = False,
    tracked: bool = True,
    document_handle: str = "",
) -> str:
    """Delete text from a paragraph.

    Finds the text within the paragraph (across run boundaries if needed).
    With tracked=True (default) the deleted text stays visible as red strikethrough
    in Word's Track Changes view — the human reviewer must accept the deletion to
    remove it permanently. With tracked=False the text is removed immediately.

    Provide context_before/context_after to disambiguate when the same text
    appears multiple times, or when it contains smart quotes / special whitespace.

    Args:
        para_id: paraId of the target paragraph.
        text: Text to delete (ASCII quotes/dashes/spaces match their Unicode equivalents).
        author: Author name shown in Word's review pane (tracked=True only).
        context_before: Text immediately before the target (for precise anchoring).
        context_after: Text immediately after the target (for precise anchoring).
        ignore_case: If True, match text and context case-insensitively.
        tracked: True (default) = red strikethrough the human accepts/rejects.
                 False = text removed immediately, no markup.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(
        doc.delete_text(
            para_id,
            text,
            author=author,
            context_before=context_before,
            context_after=context_after,
            ignore_case=ignore_case,
            tracked=tracked,
        )
    )


@mcp.tool()
def replace_text(
    para_id: str,
    find: str,
    replace: str,
    author: str = "Claude",
    context_before: str = "",
    context_after: str = "",
    ignore_case: bool = False,
    tracked: bool = True,
    document_handle: str = "",
) -> str:
    """Replace text in a paragraph.

    With tracked=True (default) the old text is shown as red strikethrough and the
    new text as an underlined insertion — the human reviewer accepts/rejects in Word.
    Only the actually-changed portion is marked; common leading/trailing text is left
    as plain runs. With tracked=False the replacement is applied immediately without
    any visible markup.

    Args:
        para_id: paraId of the target paragraph.
        find: Text to find and replace (may span run boundaries).
        replace: Replacement text.
        author: Author name shown in Word's review pane (tracked=True only).
        context_before: Text immediately before the target (for precise anchoring).
        context_after: Text immediately after the target (for precise anchoring).
        tracked: True (default) = strikethrough + underline the human accepts/rejects.
                 False = replaced immediately, no markup.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(
        doc.replace_text(
            para_id,
            find=find,
            replace=replace,
            author=author,
            context_before=context_before,
            context_after=context_after,
            ignore_case=ignore_case,
            tracked=tracked,
        )
    )


@mcp.tool()
def get_body_text(document_handle: str = "") -> str:
    """Return the full accepted-view text of the document.

    Accepted view: w:ins text included, w:del text excluded.
    Includes text inside w:hyperlink runs.
    Paragraphs are joined by newline.
    Footnote text is returned separately.

    Returns JSON: {"body": str, "footnotes": str}
    """
    _, doc = _resolve(document_handle)
    return _js(doc.get_body_text())


@mcp.tool()
def get_tracked_changes(document_handle: str = "") -> str:
    """Return all pending tracked changes (insertions and deletions) as a JSON list.

    Each entry contains: type, change_id, author, date, para_id, text.
    Changes are returned in document order.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.get_tracked_changes())


@mcp.tool()
def accept_changes(author: str = "", document_handle: str = "") -> str:
    """Accept tracked changes — keep insertions, remove deletions. Empty author = all."""
    _, doc = _resolve(document_handle)
    return _js(doc.accept_changes(author=author or None))


@mcp.tool()
def reject_changes(author: str = "", document_handle: str = "") -> str:
    """Reject tracked changes — remove insertions, restore deleted text."""
    _, doc = _resolve(document_handle)
    return _js(doc.reject_changes(author=author or None))


@mcp.tool()
def accept_change(change_id: int) -> str:
    """Accept a single tracked change by its change_id.

    For insertions: keeps the inserted text (unwraps w:ins).
    For deletions: discards the deleted text (removes w:del).

    Args:
        change_id: The integer id attribute of the w:ins or w:del element.
    """
    return _js(_require_doc().accept_change(change_id))


@mcp.tool()
def reject_change(change_id: int) -> str:
    """Reject a single tracked change by its change_id.

    For insertions: discards the inserted text (removes w:ins).
    For deletions: keeps the deleted text (unwraps w:del, restoring text).

    Args:
        change_id: The integer id attribute of the w:ins or w:del element.
    """
    return _js(_require_doc().reject_change(change_id))


@mcp.tool()
def accept_all_changes() -> str:
    """Accept all tracked changes in document order.

    Returns a JSON object with the count of accepted changes: {"accepted": int}.
    """
    return _js(_require_doc().accept_all_changes())


@mcp.tool()
def reject_all_changes() -> str:
    """Reject all tracked changes in document order.

    Returns a JSON object with the count of rejected changes: {"rejected": int}.
    """
    return _js(_require_doc().reject_all_changes())


@mcp.tool()
def set_formatting(
    para_id: str,
    text: str,
    bold: bool = False,
    italic: bool = False,
    underline: str = "",
    color: str = "",
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Apply character formatting to text with tracked-change markup."""
    _, doc = _resolve(document_handle)
    return _js(
        doc.set_formatting(
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
def get_comments(document_handle: str = "") -> str:
    """List all comments with their ID, author, date, and text."""
    _, doc = _resolve(document_handle)
    return _js(doc.get_comments())


@mcp.tool()
def add_comment(
    para_id: str,
    text: str,
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Add a comment anchored to a paragraph."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_comment(para_id, text, author=author))


@mcp.tool()
def reply_to_comment(
    parent_id: int,
    text: str,
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Reply to an existing comment (creates a threaded reply)."""
    _, doc = _resolve(document_handle)
    return _js(doc.reply_to_comment(parent_id, text, author=author))


@mcp.tool()
def update_comment(comment_id: int, text: str) -> str:
    """Replace the text of an existing comment.

    Args:
        comment_id: ID of the comment to update.
        text: New comment text.
    """
    return _js(_require_doc().update_comment(comment_id, text))


@mcp.tool()
def delete_comment(comment_id: int) -> str:
    """Delete a comment and remove its range markers from the document.

    Args:
        comment_id: ID of the comment to delete.
    """
    return _js(_require_doc().delete_comment(comment_id))


@mcp.tool()
def resolve_comment(comment_id: int) -> str:
    """Mark a comment as resolved (sets w15:done='1' in commentsExtended.xml).

    Args:
        comment_id: ID of the comment to resolve.
    """
    return _js(_require_doc().resolve_comment(comment_id))


@mcp.tool()
def list_comment_threads() -> str:
    """List all comment threads (root comments with their replies).

    Returns a list of thread dicts: {root: {id, author, date, text}, replies: [...]}.
    """
    return _js(_require_doc().list_comment_threads())


# ── Save ────────────────────────────────────────────────────────────────────


@mcp.tool()
def save_document(output_path: str = "", document_handle: str = "") -> str:
    """Save all changes back to a .docx file.

    Args:
        output_path: Path for the output file. If empty, overwrites the original
            source path of the document under this handle.
        document_handle: Handle of the document to save. Empty = `__default__`.
    """
    _, doc = _resolve(document_handle)
    path = output_path if output_path else None
    return _js(doc.save(path))


@mcp.tool()
def scrub_pii(
    output_path: str = "",
    entities: list[str] | None = None,
    confidence_threshold: float = 0.35,
    dry_run: bool = False,
    also_sanitize_metadata: bool = True,
    redact_authors_as: str = "REDACTED",
) -> str:
    """[EXPERIMENTAL] Detect and redact PII from the open document using Presidio + spaCy NER.

    WARNING: This tool WILL miss PII. It is experimental and NOT suitable for production
    use or as the sole control for privileged, regulated, or legally sensitive documents.
    Always run with dry_run=True first and manually review every detected entity before
    committing a redacted file.

    Known limitations (statistical NER gaps):
      - Names in ALL-CAPS (ledger headers, table cells) are frequently missed.
      - Single-token names with no surrounding context are unreliable.
      - Non-English names (Arabic, CJK, African) have low recall on this English model.
      - Names embedded in legal boilerplate ("Borrower: Jane Doe") are often skipped.

    NER model (en_core_web_lg, ~560MB) downloads automatically on first use.

    Detects: PERSON, EMAIL_ADDRESS, PHONE_NUMBER, CREDIT_CARD, SSN, IP_ADDRESS,
             IBAN_CODE, US_BANK_NUMBER, US_PASSPORT, and more via Presidio.

    Redacted text is replaced with a solid black DrawingML rectangle — true XML
    redaction where the original text is deleted from the OOXML entirely, not
    merely hidden by formatting.

    Args:
        output_path: Destination path. Required when dry_run=False.
        entities: Presidio entity types to redact. None = all detected types.
        confidence_threshold: Presidio score floor (default 0.35).
        dry_run: If True, detect only — return entity list, write no file.
        also_sanitize_metadata: Apply level-3 metadata sanitization (default True).
        redact_authors_as: Replacement author string for metadata pass.
    """
    return _js(
        _require_doc().scrub_pii(
            output_path,
            entities=entities,
            confidence_threshold=confidence_threshold,
            dry_run=dry_run,
            also_sanitize_metadata=also_sanitize_metadata,
            redact_authors_as=redact_authors_as,
        )
    )


@mcp.tool()
def sanitize_metadata(
    output_path: str,
    level: int = 1,
    redact_authors_as: str = "",
) -> str:
    """Write a sanitized copy of the open document to output_path.

    Level 1: Remove rsid session-fingerprint attributes from document.xml.
    Level 2: + Replace tracked-change author names (w:author on w:ins/w:del).
    Level 3: + Clear creator/lastModifiedBy/revision in docProps/core.xml
             + Clear Company in docProps/app.xml
             + Remove attachedTemplate reference from word/settings.xml

    Args:
        output_path: Destination path for the sanitized DOCX. Must be non-empty.
        level: Sanitization depth (1, 2, or 3). Default 1.
        redact_authors_as: Replacement author string for level 2+. Default "Anonymous".
    """
    return _js(
        _require_doc().sanitize_metadata(
            output_path,
            level=level,
            redact_authors_as=redact_authors_as,
        )
    )


@mcp.tool()
def compare_documents(
    base_path: str,
    revised_path: str,
    output_path: str = "",
) -> str:
    """Diff two DOCX files and produce a tracked-change document.

    Paragraph-level LCS diff:
      - Unchanged paragraphs copied verbatim.
      - Deleted paragraphs (in base, absent in revised) wrapped in w:del.
      - Inserted paragraphs (in revised, absent in base) wrapped in w:ins.
      - Modified paragraphs (1:1 replacement) get word-level del+ins inline.

    The output is a valid DOCX readable in Word/LibreOffice showing the changes
    as tracked revisions.

    Args:
        base_path: Path to the original DOCX.
        revised_path: Path to the revised DOCX.
        output_path: Destination path. Auto-generated if empty.
    """
    return _js(DocxDocument.compare_documents(base_path, revised_path, output_path))


@mcp.tool()
def diff_to_text(
    base_path: str,
    revised_path: str,
    docx_output: str = "",
    text_output: str = "",
) -> str:
    """Compare two separate DOCX files and produce a tracked-change DOCX plus a plain-text summary.

    Use this when you have two distinct files (e.g. an original and a revised copy)
    and want to show what changed between them. Produces:
      1. A DOCX with tracked-change markup (deletions in red, insertions underlined).
      2. A .txt summary suitable for pasting into an email or pull-request description.

    If you have already made tracked edits to the currently open document and just
    want to summarise those changes, use generate_change_summary instead.

    Args:
        base_path: Path to the original DOCX.
        revised_path: Path to the revised DOCX.
        docx_output: Output path for the tracked-change DOCX (auto-generated if empty).
        text_output: Output path for the plain-text summary .txt (auto-generated if empty).
    """
    return _js(
        DocxDocument.diff_to_text(
            base_path,
            revised_path,
            docx_output=docx_output,
            text_output=text_output,
        )
    )


@mcp.tool()
def generate_change_summary(
    output_path: str = "",
    document_handle: str = "",
) -> str:
    """Summarise tracked changes already present in the open document as an email-ready .txt.

    Use this after making edits with tracked=True (the default) and saving, to produce
    a human-readable change log of what was modified. Reads the document's existing
    w:ins / w:del elements, groups adjacent deletion+insertion pairs as REPLACEMENT
    entries, and writes a numbered list with author, date, and text per change.

    Typical workflow:
      open_document → edit with tracked=True → save_document → generate_change_summary

    If you have two separate files to compare rather than an already-edited document,
    use diff_to_text instead.

    Args:
        output_path: Destination .txt path. Auto-generated from the document stem if empty.
        document_handle: Optional handle for concurrent session isolation.
    """
    _, doc = _resolve(document_handle)
    return _js(doc.generate_change_summary(output_path))


@mcp.tool()
def list_parts() -> str:
    """List all XML parts (files) in the open DOCX zip."""
    return _js(_require_doc().list_parts())


@mcp.tool()
def read_part(part_path: str) -> str:
    """Read raw XML of any DOCX part (e.g. 'word/document.xml').
    Use list_parts() to discover available parts.
    """
    return _js(_require_doc().read_part(part_path))


@mcp.tool()
def write_part(part_path: str, xml: str) -> str:
    """Replace a DOCX part with new XML. Validates well-formedness first.
    WARNING: Direct XML manipulation can corrupt the document if used incorrectly.
    """
    return _js(_require_doc().write_part(part_path, xml))


@mcp.tool()
def xpath_query(xpath: str, part: str = "word/document.xml") -> str:
    """Run XPath against any DOCX part. Pre-bound namespaces: w, w14, r, wp, a, mc.

    Examples:
      xpath="//w:p" — all paragraphs
      xpath="//w:t/text()" — all text content
      xpath="//w:p[w:pPr/w:pStyle/@w:val='Heading1']" — Heading 1 paragraphs
    """
    return _js(_require_doc().xpath_query(xpath, part))


# ── Bookmarks ───────────────────────────────────────────────────────────────


@mcp.tool()
def list_bookmarks() -> str:
    """List all bookmarks in the document."""
    return _js(_require_doc().list_bookmarks())


@mcp.tool()
def add_bookmark(para_id: str, name: str) -> str:
    """Add a named bookmark wrapping the specified paragraph."""
    return _js(_require_doc().add_bookmark(para_id, name))


@mcp.tool()
def remove_bookmark(name: str) -> str:
    """Remove a bookmark by name (keeps paragraph content)."""
    return _js(_require_doc().remove_bookmark(name))


@mcp.tool()
def get_bookmarked_text(name: str) -> str:
    """Get the text content within a named bookmark."""
    return _js(_require_doc().get_bookmarked_text(name))


# ── Hyperlinks ──────────────────────────────────────────────────────────────


@mcp.tool()
def list_hyperlinks() -> str:
    """List all hyperlinks in the document.

    Returns a list of dicts with keys: id, url_or_anchor, text, para_id, type.
    type is "external" (r:id based) or "internal" (w:anchor based).
    """
    return _js(_require_doc().list_hyperlinks())


@mcp.tool()
def add_hyperlink(para_id: str, text: str, url: str) -> str:
    """Append an external hyperlink at the end of a paragraph.

    Creates a relationship entry in word/_rels/document.xml.rels and a
    w:hyperlink element with a Hyperlink-styled run in word/document.xml.

    Args:
        para_id: w14:paraId of the target paragraph.
        text: Display text for the hyperlink.
        url: The URL the hyperlink points to.
    """
    return _js(_require_doc().add_hyperlink(para_id, text, url))


@mcp.tool()
def add_internal_link(para_id: str, text: str, bookmark: str) -> str:
    """Append an internal anchor hyperlink (w:anchor) at the end of a paragraph.

    Internal links reference bookmarks by name and do NOT add a relationship.

    Args:
        para_id: w14:paraId of the target paragraph.
        text: Display text for the hyperlink.
        bookmark: Name of the bookmark to link to.
    """
    return _js(_require_doc().add_internal_link(para_id, text, bookmark))


@mcp.tool()
def remove_hyperlink(para_id: str, url_or_anchor: str) -> str:
    """Remove a hyperlink wrapper, preserving the text runs inside.

    The paragraph text remains; only the w:hyperlink element is unwrapped.

    Args:
        para_id: w14:paraId of the paragraph containing the hyperlink.
        url_or_anchor: URL (for external) or bookmark name (for internal) to match.
    """
    return _js(_require_doc().remove_hyperlink(para_id, url_or_anchor))


@mcp.tool()
def update_hyperlink(r_id: str, new_url: str) -> str:
    """Update the target URL of an existing external hyperlink relationship.

    Args:
        r_id: The relationship ID (e.g. "rId7") to update.
        new_url: The new target URL.
    """
    return _js(_require_doc().update_hyperlink(r_id, new_url))


# ── Fields ──────────────────────────────────────────────────────────────────


@mcp.tool()
def add_field(para_id: str, field_code: str, cached_value: str = "") -> str:
    """Insert a Word field at end of paragraph.

    Common field codes: PAGE, NUMPAGES, DATE, SEQ Figure, REF MyBookmark,
    STYLEREF Heading.

    Args:
        para_id: w14:paraId of the target paragraph.
        field_code: The field instruction text (e.g. "PAGE").
        cached_value: Optional display text cached in the document.
    """
    return _js(_require_doc().add_field(para_id, field_code, cached_value))


@mcp.tool()
def update_fields() -> str:
    """Mark all fields as dirty so Word recalculates on open."""
    return _js(_require_doc().update_fields())


@mcp.tool()
def list_fields() -> str:
    """List all fields in the document with their codes and cached values."""
    return _js(_require_doc().list_fields())


@mcp.tool()
def get_field(field_id: str) -> str:
    """Return details of a single field by field_id.

    Args:
        field_id: Positional identifier returned by list_fields (e.g. "field_0").
    """
    return _js(_require_doc().get_field(field_id))


@mcp.tool()
def delete_field(field_id: str) -> str:
    """Remove a complete complex field (begin through end runs) from the document.

    Args:
        field_id: Positional identifier returned by list_fields (e.g. "field_0").
    """
    return _js(_require_doc().delete_field(field_id))


@mcp.tool()
def insert_date_field(para_id: str, date_format: str = r'\@ "MMMM d, yyyy"') -> str:
    """Insert a DATE field at the end of a paragraph.

    Args:
        para_id: w14:paraId of the target paragraph.
        date_format: Date picture switch (default: \\@ "MMMM d, yyyy").
    """
    return _js(_require_doc().insert_date_field(para_id, date_format))


@mcp.tool()
def insert_page_number_field(para_id: str) -> str:
    """Insert a PAGE field at the end of a paragraph.

    Args:
        para_id: w14:paraId of the target paragraph.
    """
    return _js(_require_doc().insert_page_number_field(para_id))


@mcp.tool()
def insert_if_field(
    para_id: str,
    condition: str,
    true_text: str,
    false_text: str,
) -> str:
    """Insert a Word IF conditional field at the end of a paragraph.

    Args:
        para_id: w14:paraId of the target paragraph.
        condition: The condition expression (e.g. "x > 0").
        true_text: Text to display when condition is true.
        false_text: Text to display when condition is false.
    """
    return _js(_require_doc().insert_if_field(para_id, condition, true_text, false_text))


@mcp.tool()
def insert_sequence_field(
    para_id: str,
    seq_name: str,
    reset: bool = False,
) -> str:
    """Insert a SEQ (sequence) field for figure/table numbering.

    Args:
        para_id: w14:paraId of the target paragraph.
        seq_name: The sequence identifier (e.g. "Figure", "Table").
        reset: If True, restarts the sequence counter at 1.
    """
    return _js(_require_doc().insert_sequence_field(para_id, seq_name, reset))


@mcp.tool()
def insert_merge_field(para_id: str, field_name: str) -> str:
    """Insert a MERGEFIELD (mail merge) field at the end of a paragraph.

    Args:
        para_id: w14:paraId of the target paragraph.
        field_name: The merge field name (e.g. "FirstName").
    """
    return _js(_require_doc().insert_merge_field(para_id, field_name))


# ── Table of Contents ────────────────────────────────────────────────────────


@mcp.tool()
def generate_toc(max_level: int = 3, title: str = "Table of Contents") -> str:
    """Generate a Table of Contents from document headings."""
    doc = _require_doc()
    return _js(doc.generate_toc(max_level, title))


@mcp.tool()
def update_toc() -> str:
    """Regenerate ToC entries from current headings."""
    doc = _require_doc()
    return _js(doc.update_toc())


@mcp.tool()
def generate_list_of_figures() -> str:
    """Insert a List of Figures field (requires SEQ Figure captions)."""
    doc = _require_doc()
    return _js(doc.generate_list_of_figures())


@mcp.tool()
def generate_list_of_tables() -> str:
    """Insert a List of Tables field (requires SEQ Table captions)."""
    doc = _require_doc()
    return _js(doc.generate_list_of_tables())


@mcp.tool()
def generate_tof(para_id: str, title: str = "List of Figures") -> str:
    """Insert a Table of Figures field block after the paragraph with para_id.

    Collects caption-styled paragraphs starting with 'Figure' as entries.
    """
    doc = _require_doc()
    return _js(doc.generate_tof(para_id, title))


@mcp.tool()
def generate_tot(para_id: str, title: str = "List of Tables") -> str:
    """Insert a Table of Tables field block after the paragraph with para_id.

    Collects caption-styled paragraphs starting with 'Table' as entries.
    """
    doc = _require_doc()
    return _js(doc.generate_tot(para_id, title))


# ── Content Controls ────────────────────────────────────────────────────────


@mcp.tool()
def add_content_control(
    para_id: str,
    tag: str,
    control_type: str,
    label: str = "",
    options: list[str] | None = None,
    default: str = "",
) -> str:
    """Wrap a paragraph in an SDT content control.

    Args:
        para_id: paraId of the paragraph to wrap.
        tag: Unique tag name for the control.
        control_type: One of text|checkbox|dropdown|date.
        label: Optional display label (w:alias).
        options: List of option strings for dropdown controls.
        default: Default display text (or date string for date controls).
    """
    doc = _require_doc()
    return _js(doc.add_content_control(para_id, tag, control_type, label, options, default))


@mcp.tool()
def get_content_controls() -> str:
    """List all SDT content controls in the document."""
    doc = _require_doc()
    return _js(doc.get_content_controls())


@mcp.tool()
def set_content_control_value(tag: str, value: str) -> str:
    """Update the value/text of a content control by its tag.

    Args:
        tag: Tag name of the control to update.
        value: New value. For checkbox: 'true'/'1' = checked, 'false'/'0' = unchecked.
    """
    doc = _require_doc()
    return _js(doc.set_content_control_value(tag, value))


@mcp.tool()
def lock_content_control(tag: str, lock: str = "sdtLocked") -> str:
    """Lock a content control to prevent editing.

    Args:
        tag: Tag name of the control to lock.
        lock: Lock type — sdtLocked|contentLocked|sdtContentLocked.
    """
    doc = _require_doc()
    return _js(doc.lock_content_control(tag, lock))


@mcp.tool()
def delete_content_control(control_id: str) -> str:
    """Remove an SDT content control wrapper, keeping its content in place.

    Args:
        control_id: The w:id value of the content control to unwrap.
    """
    doc = _require_doc()
    return _js(doc.delete_content_control(control_id))


@mcp.tool()
def get_content_control(control_id: str) -> str:
    """Return details of a single content control by its w:id.

    Args:
        control_id: The w:id value of the content control to retrieve.
    """
    doc = _require_doc()
    return _js(doc.get_content_control(control_id))


@mcp.tool()
def update_content_control(
    control_id: str,
    title: str | None = None,
    tag: str | None = None,
    placeholder_text: str | None = None,
) -> str:
    """Modify properties of an existing content control.

    Args:
        control_id: The w:id value of the content control to update.
        title: New title (w:alias/@w:val). Omit to leave unchanged.
        tag: New tag (w:tag/@w:val). Omit to leave unchanged.
        placeholder_text: New placeholder text. Omit to leave unchanged.
    """
    doc = _require_doc()
    return _js(
        doc.update_content_control(
            control_id, title=title, tag=tag, placeholder_text=placeholder_text
        )
    )  # noqa: E501


# ── Template Filling ─────────────────────────────────────────────────────────


@mcp.tool()
def fill_template(data: dict[str, str | list[str]], remove_empty: bool = False) -> str:
    """Fill SDT content controls from data dict. Keys match w:tag values.

    Args:
        data: Mapping of tag names to values. Use list[str] for repeating sections.
        remove_empty: If True, remove SDTs with no matching key in data.
    """
    doc = _require_doc()
    return _js(doc.fill_template(data, remove_empty))


@mcp.tool()
def list_template_fields() -> str:
    """List all SDT template fields (tag, label, type) in the document."""
    doc = _require_doc()
    return _js(doc.list_template_fields())


@mcp.tool()
def validate_template_data(data: dict) -> str:
    """Validate data dict covers all template fields. Returns missing and extra keys."""
    doc = _require_doc()
    return _js(doc.validate_template_data(data))


# ── Multilevel Lists ─────────────────────────────────────────────────────────


@mcp.tool()
def create_multilevel_list(name: str, levels: list[dict]) -> str:
    """Create a multilevel list in numbering.xml. Each level dict: {num_fmt, lvl_text, indent, hanging, style?}."""  # noqa: E501
    doc = _require_doc()
    return _js(doc.create_multilevel_list(name, levels))


@mcp.tool()
def restart_numbering(para_id: str, level: int = 0, start: int = 1) -> str:
    """Restart list numbering at a paragraph. Adds lvlOverride with startOverride."""
    doc = _require_doc()
    return _js(doc.restart_numbering(para_id, level, start))


@mcp.tool()
def suppress_numbering(para_id: str) -> str:
    """Remove list numbering from a paragraph by setting numId to 0."""
    doc = _require_doc()
    return _js(doc.suppress_numbering(para_id))


@mcp.tool()
def get_lists() -> str:
    """Return all list definitions from numbering.xml.

    Each entry: {abstract_num_id, num_format, levels}.
    Returns [] if no numbering.xml exists.
    """
    return _js(_require_doc().get_lists())


@mcp.tool()
def promote_list_item(para_id: str) -> str:
    """Decrease the list indentation level (ilvl) of a paragraph by 1, minimum 0.

    Returns {para_id, ilvl}. Raises ValueError if paragraph is not a list item.
    """
    return _js(_require_doc().promote_list_item(para_id))


@mcp.tool()
def demote_list_item(para_id: str) -> str:
    """Increase the list indentation level (ilvl) of a paragraph by 1, maximum 8.

    Returns {para_id, ilvl}. Raises ValueError if paragraph is not a list item.
    """
    return _js(_require_doc().demote_list_item(para_id))


# ── Litigation Tools ────────────────────────────────────────────────────────


@mcp.tool()
def bates_number(
    prefix: str, start: int = 1, digits: int = 6, position: str = "footer-right"
) -> str:  # noqa: E501
    """Add Bates numbering stamp to document footer.

    Args:
        prefix: Bates prefix string (e.g. "ACME-").
        start: Starting Bates number.
        digits: Zero-padding width for the number.
        position: Hint for stamp position (currently "footer-right").
    """
    doc = _require_doc()
    return _js(doc.bates_number(prefix, start, digits, position))


@mcp.tool()
def redact_text(
    pattern: str | None = None,
    para_ids: list[str] | None = None,
    exact_text: str | None = None,
    reason: str = "",
) -> str:
    """True redaction: remove text and replace with black rectangle. Use exact_text or pattern.

    Args:
        pattern: Regex pattern to match run text.
        para_ids: Optional list of paragraph paraId values to limit scope.
        exact_text: Exact string to match against run text.
        reason: Reason for redaction (stored in log).
    """
    doc = _require_doc()
    return _js(doc.redact_text(pattern, para_ids, exact_text, reason))


@mcp.tool()
def generate_redaction_log(output_path: str = "") -> str:
    """Write a DOCX table of all redactions made this session.

    Args:
        output_path: Destination path. If empty, writes to a temp file.
    """
    doc = _require_doc()
    return _js(doc.generate_redaction_log(output_path))


@mcp.tool()
def generate_privilege_log(output_path: str = "") -> str:
    """Generate a privilege log DOCX from document metadata.

    Args:
        output_path: Destination path. If empty, writes to a temp file.
    """
    doc = _require_doc()
    return _js(doc.generate_privilege_log(output_path))


# ── Equations ───────────────────────────────────────────────────────────────


@mcp.tool()
def add_equation(para_id: str, latex: str) -> str:
    """Insert a LaTeX equation as OMML. Requires: pip install latex2mathml.

    Args:
        para_id: paraId of the paragraph after which the equation is inserted.
        latex: LaTeX source string (e.g. r"\\frac{1}{2}").
    """
    doc = _require_doc()
    return _js(doc.add_equation(para_id, latex))


@mcp.tool()
def get_equations() -> str:
    """Return all equations in the document as OMML XML strings."""
    doc = _require_doc()
    return _js(doc.get_equations())


# ── Charts ───────────────────────────────────────────────────────────────────


@mcp.tool()
def insert_bar_chart(
    para_id: str,
    title: str,
    series: list[dict],
    categories: list[str],
    width_cm: float = 14.0,
    height_cm: float = 9.0,
) -> str:
    """Insert a native bar chart (no Excel required).

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.insert_bar_chart(para_id, title, series, categories, width_cm, height_cm))


@mcp.tool()
def insert_line_chart(
    para_id: str,
    title: str,
    series: list[dict],
    categories: list[str],
    width_cm: float = 14.0,
    height_cm: float = 9.0,
) -> str:
    """Insert a native line chart.

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.insert_line_chart(para_id, title, series, categories, width_cm, height_cm))


@mcp.tool()
def insert_pie_chart(
    para_id: str,
    title: str,
    series: list[dict],
    categories: list[str],
) -> str:
    """Insert a native pie chart (single series, fixed 14x9 cm).

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.insert_pie_chart(para_id, title, series, categories))


@mcp.tool()
def update_chart_data(chart_id: str, series: list[dict]) -> str:
    """Replace data series in an existing chart by chart_id.

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.update_chart_data(chart_id, series))


@mcp.tool()
def merge_review_rounds(reviewer_paths: list[str], base_path: str | None = None) -> str:
    """Merge tracked changes from N reviewer copies into the open document."""
    return _js(_require_doc().merge_review_rounds(reviewer_paths, base_path))


@mcp.tool()
def compare_contracts(other_path: str, output_path: str = "", align_by: str = "heading") -> str:
    """Clause-aware diff between the open contract and another .docx file."""
    return _js(_require_doc().compare_contracts(other_path, output_path, align_by))


# ── Session log ──────────────────────────────────────────────────────────────


@mcp.tool()
def get_session_log() -> str:
    """Return all operations performed this session as replayable JSON."""
    return _js(_require_doc().get_session_log())


@mcp.tool()
def export_session_script(output_path: str) -> str:
    """Write session operations as a Python replay script.

    Returns: {"output_path": str, "operations": int}
    """
    return _js(_require_doc().export_session_script(output_path))


# ── Paragraph CRUD + border/shading ────────────────────────────────────────


@mcp.tool()
def insert_paragraph(
    after_para_id: str,
    text: str,
    style: str = "",
) -> str:
    """Insert a new paragraph after the paragraph with the given paraId.

    Args:
        after_para_id: paraId of the paragraph to insert after.
        text: Text content for the new paragraph.
        style: Optional paragraph style name (e.g., "Heading1", "Normal").
    """
    return _js(_require_doc().insert_paragraph(after_para_id, text, style=style or None))


@mcp.tool()
def update_paragraph(
    para_id: str,
    text: str = "",
    style: str = "",
) -> str:
    """Update the text and/or style of an existing paragraph.

    Args:
        para_id: paraId of the paragraph to update.
        text: New text content. Empty string leaves text unchanged.
        style: New paragraph style name. Empty string leaves style unchanged.
    """
    return _js(
        _require_doc().update_paragraph(
            para_id,
            text=text or None,
            style=style or None,
        )
    )


@mcp.tool()
def delete_paragraph(para_id: str) -> str:
    """Delete the paragraph with the given paraId.

    Args:
        para_id: paraId of the paragraph to remove.
    """
    return _js(_require_doc().delete_paragraph(para_id))


@mcp.tool()
def set_paragraph_border(
    para_id: str,
    sides: list[str],
    color: str = "000000",
    size: int = 4,
) -> str:
    """Set borders on one or more sides of a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        sides: List of sides: "top", "bottom", "left", "right", "between".
        color: Border color as 6-digit hex (default "000000").
        size: Border width in eighths of a point (default 4 = 0.5pt).
    """
    return _js(_require_doc().set_paragraph_border(para_id, sides, color=color, size=size))


@mcp.tool()
def set_paragraph_shading(
    para_id: str,
    fill_color: str,
    pattern: str = "clear",
) -> str:
    """Set background shading on a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        fill_color: Fill color as 6-digit hex (e.g., "FFFF00").
        pattern: Shading pattern (default "clear").
    """
    return _js(_require_doc().set_paragraph_shading(para_id, fill_color, pattern=pattern))


@mcp.tool()
def set_paragraph_indentation(
    para_id: str,
    left_cm: float | None = None,
    right_cm: float | None = None,
    first_line_cm: float | None = None,
    hanging_cm: float | None = None,
) -> str:
    """Set indentation on a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        left_cm: Left indent in centimetres (None = unchanged).
        right_cm: Right indent in centimetres (None = unchanged).
        first_line_cm: First-line indent in cm (mutually exclusive with hanging_cm).
        hanging_cm: Hanging indent in cm (mutually exclusive with first_line_cm).
    """
    return _js(
        _require_doc().set_paragraph_indentation(
            para_id,
            left_cm=left_cm,
            right_cm=right_cm,
            first_line_cm=first_line_cm,
            hanging_cm=hanging_cm,
        )
    )


@mcp.tool()
def set_keep_with_next(para_id: str, enabled: bool) -> str:
    """Keep this paragraph on the same page as the next paragraph.

    Args:
        para_id: paraId of the target paragraph.
        enabled: True to enable keep-with-next, False to remove it.
    """
    return _js(_require_doc().set_keep_with_next(para_id, enabled))


@mcp.tool()
def set_keep_lines_together(para_id: str, enabled: bool) -> str:
    """Keep all lines of this paragraph on the same page.

    Args:
        para_id: paraId of the target paragraph.
        enabled: True to enable keep-lines-together, False to remove it.
    """
    return _js(_require_doc().set_keep_lines_together(para_id, enabled))


@mcp.tool()
def set_page_break_before(para_id: str, enabled: bool) -> str:
    """Force a page break before this paragraph.

    Args:
        para_id: paraId of the target paragraph.
        enabled: True to force page break before, False to remove it.
    """
    return _js(_require_doc().set_page_break_before(para_id, enabled))


@mcp.tool()
def set_widow_control(para_id: str, enabled: bool) -> str:
    """Enable widow/orphan control for this paragraph.

    Args:
        para_id: paraId of the target paragraph.
        enabled: True to enable widow control, False to remove it.
    """
    return _js(_require_doc().set_widow_control(para_id, enabled))


@mcp.tool()
def insert_blockquote(para_id: str, text: str) -> str:
    """Insert a blockquote paragraph after the given paragraph.

    The new paragraph has 720-twip left indent and italic formatting.

    Args:
        para_id: paraId of the reference paragraph (new para inserted after it).
        text: Text content of the blockquote.
    """
    return _js(_require_doc().insert_blockquote(para_id, text))


@mcp.tool()
def insert_code_block(para_id: str, text: str, language: str = "") -> str:
    """Insert a code-block paragraph after the given paragraph.

    The new paragraph uses Courier New 10pt with light-gray background shading.

    Args:
        para_id: paraId of the reference paragraph (new para inserted after it).
        text: Code text content.
        language: Optional language hint (stored in return value only).
    """
    return _js(_require_doc().insert_code_block(para_id, text, language=language))


@mcp.tool()
def set_line_spacing(
    para_id: str,
    line_rule: str | None = None,
    line_value: int | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
) -> str:
    """Set line spacing and paragraph spacing.

    Args:
        para_id: paraId of the target paragraph.
        line_rule: "auto" | "exact" | "atLeast". None leaves unchanged.
        line_value: For "auto": 240=single, 360=1.5x, 480=double. For "exact"/"atLeast": twips.
        space_before_pt: Space before paragraph in points (None = unchanged).
        space_after_pt: Space after paragraph in points (None = unchanged).
    """
    return _js(
        _require_doc().set_line_spacing(
            para_id,
            line_rule=line_rule,
            line_value=line_value,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
        )
    )


@mcp.tool()
def get_paragraph_format(para_id: str) -> str:
    """Read all formatting attributes of a paragraph.

    Returns style, alignment, indentation, line spacing, border, shading, and list info.

    Args:
        para_id: paraId of the target paragraph.
    """
    return _js(_require_doc().get_paragraph_format(para_id))


# ── Run-level formatting ───────────────────────────────────────────────────


@mcp.tool()
def get_runs(para_id: str) -> str:
    """Get all runs in a paragraph with their formatting properties.

    Args:
        para_id: paraId of the target paragraph.
    """
    return _js(_require_doc().get_runs(para_id))


@mcp.tool()
def set_run_font(para_id: str, run_idx: int, font_name: str) -> str:
    """Set the font of a specific run (zero-based index) in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        font_name: Font name (e.g., "Arial", "Times New Roman").
    """
    return _js(_require_doc().set_run_font(para_id, run_idx, font_name))


@mcp.tool()
def set_run_color(para_id: str, run_idx: int, color: str) -> str:
    """Set the font color of a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        color: Hex color without # (e.g., "FF0000").
    """
    return _js(_require_doc().set_run_color(para_id, run_idx, color))


@mcp.tool()
def set_run_size(para_id: str, run_idx: int, size_pt: float) -> str:
    """Set the font size of a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        size_pt: Font size in points (e.g., 12.0).
    """
    return _js(_require_doc().set_run_size(para_id, run_idx, size_pt))


@mcp.tool()
def set_character_spacing(para_id: str, run_idx: int, spacing_pt: float) -> str:
    """Set character spacing (tracking) for a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        spacing_pt: Spacing in points (positive = expanded, negative = condensed).
    """
    return _js(_require_doc().set_character_spacing(para_id, run_idx, spacing_pt))


@mcp.tool()
def set_character_position(para_id: str, run_idx: int, position_pt: float) -> str:
    """Set vertical character position (raised/lowered) for a specific run.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        position_pt: Offset in points (positive = raised, negative = lowered).
    """
    return _js(_require_doc().set_character_position(para_id, run_idx, position_pt))


@mcp.tool()
def set_run_highlight(para_id: str, run_idx: int, color: str) -> str:
    """Set highlight color of a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        color: Word highlight color name (e.g., "yellow", "green", "cyan", "none").
    """
    return _js(_require_doc().set_run_highlight(para_id, run_idx, color))


@mcp.tool()
def set_run_strikethrough(para_id: str, run_idx: int, double: bool = False) -> str:
    """Set strikethrough on a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        double: False for single strikethrough, True for double strikethrough.
    """
    return _js(_require_doc().set_run_strikethrough(para_id, run_idx, double))


@mcp.tool()
def set_run_superscript(para_id: str, run_idx: int) -> str:
    """Set superscript vertical alignment on a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
    """
    return _js(_require_doc().set_run_superscript(para_id, run_idx))


@mcp.tool()
def set_run_subscript(para_id: str, run_idx: int) -> str:
    """Set subscript vertical alignment on a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
    """
    return _js(_require_doc().set_run_subscript(para_id, run_idx))


@mcp.tool()
def set_run_underline(para_id: str, run_idx: int, style: str = "single") -> str:
    """Set underline style on a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        style: Underline style (e.g., "single", "double", "dotted", "none").
    """
    return _js(_require_doc().set_run_underline(para_id, run_idx, style))


@mcp.tool()
def clear_run_formatting(para_id: str, run_idx: int) -> str:
    """Remove all character formatting from a run, causing it to inherit paragraph/style defaults.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
    """
    return _js(_require_doc().clear_run_formatting(para_id, run_idx))


@mcp.tool()
def set_run_language(para_id: str, run_idx: int, language_code: str) -> str:
    """Set the language on a run for spell-checking purposes.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        language_code: BCP-47 language code (e.g., "en-US", "fr-FR", "de-DE").
    """
    return _js(_require_doc().set_run_language(para_id, run_idx, language_code))


@mcp.tool()
def set_text_case(para_id: str, run_idx: int, case: str) -> str:
    """Set text case transformation on a run.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        case: "upper" (all caps), "small" (small caps), or "none" (remove case transform).
    """
    return _js(_require_doc().set_text_case(para_id, run_idx, case))


@mcp.tool()
def export_markdown(output_path: str = "") -> str:
    """Export the open document as Markdown.

    Converts the document body (headings, bold/italic runs, lists, tables,
    plain paragraphs) to GitHub-Flavoured Markdown and writes it to a file.

    Args:
        output_path: Destination path for the .md file.
            Defaults to <workdir>/export.md when empty.

    Returns: {"output_path": str, "paragraphs": int, "tables": int}
    """
    return _js(_require_doc().export_markdown(output_path))


@mcp.tool()
def get_theme_colors() -> str:
    """Return the named color slots from word/theme/theme1.xml.

    Returns a dict mapping slot names (dk1, lt1, accent1, ...) to 6-digit hex strings.
    Returns an empty dict if the document has no theme file.
    """
    return _js(_require_doc().get_theme_colors())


@mcp.tool()
def set_theme_color(slot: str, hex_color: str) -> str:
    """Update a named color slot in the document theme.

    Args:
        slot: One of dk1, lt1, dk2, lt2, accent1-accent6, hlink, folHlink.
        hex_color: 6-character hex string without # (e.g. "FF0000").
    """
    return _js(_require_doc().set_theme_color(slot, hex_color))


@mcp.tool()
def insert_caption(after_para_id: str, text: str, label: str = "Figure") -> str:
    """Insert a caption paragraph after the specified paragraph.

    Args:
        after_para_id: paraId of the paragraph to insert after.
        text: Caption description text (without label/number prefix).
        label: Caption label, e.g. "Figure" or "Table" (default "Figure").
    """
    return _js(_require_doc().insert_caption(after_para_id, text, label=label))


@mcp.tool()
def find_replace_formatted(
    find: str,
    replace: str,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    size_pt: float | None = None,
) -> str:
    """Find all occurrences of a string and replace with formatted text.

    Replaces every occurrence of `find` across all paragraphs in the document
    with `replace`, applying the specified character formatting to the replacement
    run only.

    Args:
        find: Text to search for (must be non-empty).
        replace: Replacement text.
        bold: True to bold, False to explicitly un-bold, None to leave unchanged.
        italic: True to italicise, False to explicitly un-italicise, None to leave unchanged.
        color: Font color as 6-digit hex (e.g. "FF0000"). None leaves color unchanged.
        size_pt: Font size in points (e.g. 12.0). None leaves size unchanged.
    """
    return _js(
        _require_doc().find_replace_formatted(
            find,
            replace,
            bold=bold,
            italic=italic,
            color=color,
            size_pt=size_pt,
        )
    )


@mcp.tool()
def split_document(output_dir: str = "", at_heading_level: int = 1) -> str:
    """Split the open document into multiple DOCX files, one per heading section.

    Args:
        output_dir: Directory for output files. Defaults to <workdir>/split_output.
        at_heading_level: Heading level to split on (default 1).

    Returns: {"output_dir": str, "files": list[str], "parts": int}
    """
    return _js(
        _require_doc().split_document(output_dir=output_dir, at_heading_level=at_heading_level)
    )  # noqa: E501


@mcp.tool()
def get_word_count() -> str:
    """Return the word count of the open document body.

    Returns: {"word_count": int}
    """
    return _js({"word_count": _require_doc().get_word_count()})


@mcp.tool()
def get_statistics() -> str:
    """Return document statistics for the open document.

    Returns a dict with keys: word_count, character_count, paragraph_count,
    table_count, image_count, section_count.
    """
    return _js(_require_doc().get_statistics())


@mcp.tool()
def get_document_outline(max_level: int = 6) -> str:
    """Return a flat list of headings as a document outline.

    Walks the document body and returns every heading paragraph up to the
    specified level.

    Args:
        max_level: Maximum heading level to include (1–6, default 6).

    Returns:
        JSON list of dicts with keys: level (int), text (str), para_id (str).
    """
    return _js(_require_doc().get_document_outline(max_level=max_level))


@mcp.tool()
def set_document_language(language_code: str) -> str:
    """Set the default document language.

    Writes the BCP-47 language tag into the default paragraph style's run
    properties (``w:rPr/w:lang``) in ``word/styles.xml``.

    Args:
        language_code: BCP-47 language tag, e.g. "en-US", "fr-FR", "de-DE".

    Returns:
        {"language": language_code}
    """
    return _js(_require_doc().set_document_language(language_code))


@mcp.tool()
def set_track_changes(enabled: bool, author: str = "") -> str:
    """Enable or disable revision tracking in the document.

    When enabled, Word/LibreOffice will record future edits as tracked changes.

    Args:
        enabled: True to enable tracking, False to disable.
        author: Optional author name (stored in response for reference only).

    Returns:
        {"track_changes": bool, "author": str}
    """
    return _js(_require_doc().set_track_changes(enabled=enabled, author=author))


@mcp.tool()
def split_table(table_idx: int, at_row_index: int) -> str:
    """Split a table into two tables at the given row index.

    Rows 0..at_row_index-1 remain in the original table; rows
    at_row_index..end move to a new table inserted immediately after.

    Args:
        table_idx: 0-based table index.
        at_row_index: 0-based row index to split at (must be > 0 and < row_count).

    Returns:
        {"table1_rows": int, "table2_rows": int}
    """
    return _js(_require_doc().split_table(table_idx, at_row_index))


@mcp.tool()
def duplicate_table_row(table_idx: int, row_index: int) -> str:
    """Deep-copy a table row and insert the copy immediately after it.

    Args:
        table_idx: 0-based table index.
        row_index: 0-based index of the row to duplicate.

    Returns:
        {"row_index": int, "new_row_index": int}
    """
    return _js(_require_doc().duplicate_table_row(table_idx, row_index))


@mcp.tool()
def sort_table(table_idx: int, column_index: int, ascending: bool = True) -> str:
    """Sort the non-header rows of a table by the text content of a column.

    Header rows (marked with w:tblHeader) remain at the top. Rows with a
    missing cell at column_index sort as empty string.

    Args:
        table_idx: 0-based table index.
        column_index: 0-based column to sort by.
        ascending: True for A→Z (default), False for Z→A.

    Returns:
        {"sorted_rows": int, "column_index": int, "ascending": bool}
    """
    return _js(_require_doc().sort_table(table_idx, column_index, ascending))


@mcp.tool()
def get_table(table_idx: int) -> str:
    """Get structured info for a single table by zero-based index.

    Args:
        table_idx: 0-based table index.

    Returns:
        {"index": int, "row_count": int, "col_count": int, "cells": list[list[str]]}
    """
    return _js(_require_doc().get_table(table_idx))


@mcp.tool()
def get_cell_text(table_idx: int, row_idx: int, col_idx: int) -> str:
    """Return text content of a specific cell.

    Args:
        table_idx: 0-based table index.
        row_idx: 0-based row index.
        col_idx: 0-based column index.

    Returns:
        {"table_index": int, "row_index": int, "col_index": int, "text": str}
    """
    return _js(_require_doc().get_cell_text(table_idx, row_idx, col_idx))


@mcp.tool()
def copy_table(table_idx: int) -> str:
    """Deep-copy a table and insert the copy immediately after the original.

    Args:
        table_idx: 0-based table index.

    Returns:
        {"source_index": int, "new_index": int}
    """
    return _js(_require_doc().copy_table(table_idx))


@mcp.tool()
def copy_document(output_path: str) -> str:
    """Save a complete snapshot of the open document to a new path.

    The active session and source path are unchanged.

    Args:
        output_path: Destination file path (must end in .docx).

    Returns:
        {"copied_to": output_path}
    """
    return _js(_require_doc().copy_document(output_path))


@mcp.tool()
def convert_to_pdf(output_path: str) -> str:
    """Convert the open document to PDF using LibreOffice headless.

    Requires LibreOffice to be installed ('libreoffice' or 'soffice' on PATH).

    Args:
        output_path: Destination path for the output PDF file.

    Returns:
        {"pdf_path": str}
    """
    return _js(_require_doc().convert_to_pdf(output_path))


@mcp.tool()
def flatten_document() -> str:
    """Accept all tracked changes and remove all revision markup.

    Accepts every w:ins / w:del and strips all w:rPrChange and w:pPrChange
    elements, leaving plain text with no tracked-change metadata.

    Returns:
        {"changes_accepted": int, "formatting_changes_removed": int}
    """
    return _js(_require_doc().flatten_document())


@mcp.tool()
def get_reading_time(words_per_minute: int = 200) -> str:
    """Estimate reading time for the open document.

    Args:
        words_per_minute: Assumed reading speed (default 200 wpm).

    Returns:
        {"word_count": int, "words_per_minute": int, "minutes": float, "seconds": int}
    """
    return _js(_require_doc().get_reading_time(words_per_minute=words_per_minute))


# ── Accessibility (P8.8) ─────────────────────────────────────────────────────


@mcp.tool()
def set_alt_text(image_index: int, alt_text: str, title: str = "") -> str:
    """Set the alt text (and optionally title) on an image by 0-based index.

    Args:
        image_index: 0-based index across all wp:docPr elements in document.xml.
        alt_text: Accessibility description to set on the image.
        title: Optional title attribute; removed from XML if empty.

    Returns:
        {"image_index": int, "alt_text": str}
    """
    return _js(_require_doc().set_alt_text(image_index, alt_text, title=title))


@mcp.tool()
def get_alt_text(image_index: int) -> str:
    """Get the alt text and title for an image by 0-based index.

    Args:
        image_index: 0-based index across all wp:docPr elements in document.xml.

    Returns:
        {"image_index": int, "alt_text": str, "title": str}
    """
    return _js(_require_doc().get_alt_text(image_index))


@mcp.tool()
def check_accessibility() -> str:
    """Scan the document for accessibility issues.

    Checks images for missing alt text and tables for missing header rows.

    Returns:
        {"issue_count": int, "issues": list[dict]}
    """
    return _js(_require_doc().check_accessibility())


@mcp.tool()
def insert_text_box(
    para_id: str,
    text: str,
    width_cm: float = 5.0,
    height_cm: float = 2.0,
) -> str:
    """Insert an inline text box after the paragraph with para_id.

    Args:
        para_id: paraId of the reference paragraph (insert after this).
        text: Text content for the text box.
        width_cm: Width in centimetres (default 5.0).
        height_cm: Height in centimetres (default 2.0).

    Returns:
        JSON with para_id (new paragraph), text, width_cm, height_cm.
    """
    return _js(
        _require_doc().insert_text_box(para_id, text, width_cm=width_cm, height_cm=height_cm)
    )  # noqa: E501


# ── Backward-compat module property ─────────────────────────────────────────
# Tests and legacy callers that access `server._doc` directly continue to work:
# reads return _docs["__default__"], writes update the same slot.


class _Module(_sys.modules[__name__].__class__):
    @property
    def _doc(self) -> DocxDocument | None:
        return _docs.get(_DEFAULT_HANDLE)

    @_doc.setter
    def _doc(self, val: DocxDocument | None) -> None:
        if val is None:
            _docs.pop(_DEFAULT_HANDLE, None)
        else:
            _docs[_DEFAULT_HANDLE] = val


_sys.modules[__name__].__class__ = _Module


# ── Entry point ─────────────────────────────────────────────────────────────


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
