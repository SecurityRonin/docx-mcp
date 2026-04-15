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
        raise RuntimeError(
            f"No document is open for handle {key!r}. Call open_document first."
        )
    return key, doc


def _store(handle: str, doc: DocxDocument) -> str:
    """Place a document under a handle, closing any previous occupant. Returns the key."""
    key = _key(handle) if handle else uuid.uuid4().hex
    existing = _docs.get(key)
    if existing is not None:
        existing.close()
    _docs[key] = doc
    return key


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
    document_handle: str = "",
) -> str:
    """Modify a table cell with tracked changes (delete old, insert new)."""
    _, doc = _resolve(document_handle)
    return _js(doc.modify_cell(table_idx, row, col, text, author=author))


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
    document_handle: str = "",
) -> str:
    """Edit text in a header or footer with tracked changes."""
    _, doc = _resolve(document_handle)
    return _js(doc.edit_header_footer(location, old_text, new_text, author=author))


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
    return _js(
        doc.insert_image(para_id, image_path, width_emu=width_emu, height_emu=height_emu)
    )


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
def add_footnote(para_id: str, text: str, document_handle: str = "") -> str:
    """Add a footnote to a paragraph."""
    _, doc = _resolve(document_handle)
    return _js(doc.add_footnote(para_id, text))


@mcp.tool()
def validate_footnotes(document_handle: str = "") -> str:
    """Cross-reference footnote IDs between document.xml and footnotes.xml."""
    _, doc = _resolve(document_handle)
    return _js(doc.validate_footnotes())


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
    document_handle: str = "",
) -> str:
    """Set document protection. edit: trackedChanges/comments/readOnly/forms/none."""
    _, doc = _resolve(document_handle)
    return _js(doc.set_document_protection(edit, password=password or None))


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
    document_handle: str = "",
) -> str:
    """Insert text with Word track-changes markup.

    Args:
        para_id: paraId of the target paragraph.
        text: Text to insert.
        position: Where to insert — "start", "end", or a substring to insert after.
        author: Author name for the revision (shown in Word's review pane).
    """
    _, doc = _resolve(document_handle)
    return _js(doc.insert_text(para_id, text, position=position, author=author))


@mcp.tool()
def delete_text(
    para_id: str,
    text: str,
    author: str = "Claude",
    document_handle: str = "",
) -> str:
    """Mark text as deleted with Word track-changes markup (red strikethrough in Word)."""
    _, doc = _resolve(document_handle)
    return _js(doc.delete_text(para_id, text, author=author))


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


# ── Entry point ─────────────────────────────────────────────────────────────


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
