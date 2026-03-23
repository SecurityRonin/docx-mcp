---
name: docx-mcp
description: "Use when editing existing Word (.docx) documents with track changes, comments, footnotes, or structural validation. Triggers: reviewing contracts, marking up reports, adding revision comments, validating document structure, removing watermarks, auditing OOXML integrity, bulk text replacement with revisions, footnote management, paragraph-level edits. Requires the docx-mcp MCP server to be running."
---

# Editing Word Documents with docx-mcp

## Overview

The `docx-mcp` MCP server provides tools for reading and editing .docx files with proper OOXML markup. Edits appear as real revisions in Microsoft Word — red strikethrough for deletions, green underline for insertions, comments in the sidebar.

A .docx file is a ZIP archive of XML files. This server unpacks the archive, parses all XML parts with lxml, edits the cached DOM trees directly, and repacks modified XML back into a valid .docx archive. This gives full control over OOXML markup — essential for track changes, comments, and structural validation that higher-level libraries don't expose.

## When to Use

- Editing existing .docx files with tracked changes
- Adding comments or footnotes to documents
- Reviewing/auditing document structure
- Removing watermarks
- Bulk find-and-replace with revision marks
- Any task where changes must be visible as Word revisions

**Do NOT use for:** Creating new .docx from scratch (use docx-js instead), PDFs, spreadsheets, or `.doc` (legacy binary format — convert to .docx first).

## Workflow

```
1. open_document("/path/to/file.docx")
2. get_headings() or get_document_info()     → understand structure
3. search_text("clause text")                → find target paragraphs
4. get_paragraph(para_id)                    → verify exact text before editing
5. delete_text(para_id, "old text")          → tracked deletion
6. insert_text(para_id, "new text")          → tracked insertion
7. add_comment(para_id, "Reason for change") → explain the edit
8. audit_document()                          → verify integrity
9. save_document("/path/to/output.docx")     → save (or omit path to overwrite)
```

**Always `audit_document()` before saving** to catch structural issues (orphaned footnotes, duplicate paraIds, unpaired bookmarks, missing relationship targets).

## Tool Quick Reference

### Document Lifecycle

| Tool | Purpose | Key args |
|------|---------|----------|
| `open_document` | Open .docx for editing | `path` |
| `close_document` | Close and clean up | — |
| `get_document_info` | Stats overview | — |
| `save_document` | Save to .docx | `output_path` (optional) |

### Reading

| Tool | Purpose | Key args |
|------|---------|----------|
| `get_headings` | Heading tree with paraIds | — |
| `search_text` | Find text in body/footnotes/comments | `query`, `regex` |
| `get_paragraph` | Full text of one paragraph | `para_id` |

### Track Changes

| Tool | Purpose | Key args |
|------|---------|----------|
| `insert_text` | Tracked insertion (green underline) | `para_id`, `text`, `position` |
| `delete_text` | Tracked deletion (red strikethrough) | `para_id`, `text` |

### Comments

| Tool | Purpose | Key args |
|------|---------|----------|
| `get_comments` | List all comments | — |
| `add_comment` | Comment anchored to paragraph | `para_id`, `text` |
| `reply_to_comment` | Threaded reply | `parent_id`, `text` |

### Footnotes

| Tool | Purpose | Key args |
|------|---------|----------|
| `get_footnotes` | List all footnotes | — |
| `add_footnote` | Footnote with superscript ref | `para_id`, `text` |
| `validate_footnotes` | Cross-ref footnote IDs | — |

### Validation & Audit

| Tool | Purpose | Key args |
|------|---------|----------|
| `validate_paraids` | Check paraId uniqueness | — |
| `remove_watermark` | Remove DRAFT watermarks | — |
| `audit_document` | Full structural audit | — |

## Essential Patterns

### Replace Text (delete + insert on same paragraph)

```
1. search_text("30 days")              → find the paragraph
2. get_paragraph(para_id)              → verify exact text
3. delete_text(para_id, "30 days")     → tracked deletion
4. insert_text(para_id, "60 days")     → tracked insertion
```

Word shows both marks side by side: ~~30 days~~ **60 days**.

### Batch Edits Across Multiple Paragraphs

```
1. search_text("Net 30", regex=False)  → returns all matches with paraIds
2. For each match:
   a. get_paragraph(para_id)           → verify context
   b. delete_text(para_id, "Net 30")
   c. insert_text(para_id, "Net 60")
   d. add_comment(para_id, "Updated payment terms per Amendment 3")
3. audit_document()                    → verify no structural damage
```

### Add Explanatory Footnote

```
1. search_text("force majeure")        → find the clause
2. add_footnote(para_id, "Force majeure includes acts of God, war, pandemic, and other events beyond reasonable control.")
3. validate_footnotes()                → verify cross-references
```

### Full Document Review

```
1. open_document("/path/to/contract.docx")
2. get_document_info()                 → paragraph count, headings, footnotes
3. get_headings()                      → see structure at a glance
4. audit_document()                    → check for pre-existing issues
5. ... make edits ...
6. audit_document()                    → verify edits didn't break anything
7. save_document("/path/to/contract_reviewed.docx")
```

## Tips

- **paraId** is an 8-char hex string (e.g., `"1A2B3C4D"`). Get them from `get_headings()` or `search_text()`.
- **position** in `insert_text`: `"start"`, `"end"`, or a substring to insert after.
- **author** defaults to `"Claude"` for all tracked changes and comments.
- **Save to new file** to preserve the original: `save_document("/path/to/revised.docx")`.
- **Always verify before editing**: Use `get_paragraph()` to see the exact text before calling `delete_text()`. The text must match exactly within a single run.

## OOXML Pitfalls (Critical Knowledge)

These hard-won lessons prevent silent document corruption. Word may "repair" broken documents by silently rewriting your edits.

### ParaId Rules

Every `<w:p>` and `<w:tr>` element has a `w14:paraId` attribute.

| Rule | Detail |
|------|--------|
| Must be unique | Across ALL XML parts: document.xml, footnotes.xml, headers, footers, endnotes |
| Must be < 0x80000000 | Word reserves the high bit internally — values >= 0x80000000 cause validation failure |
| 8 hex digits | Always uppercase, zero-padded (e.g., `1A2B3C4D`) |
| Duplicates from copy-paste | Duplicating content creates duplicate paraIds — fix the second occurrence |

The `validate_paraids()` tool checks all of these. Run it after any structural edits.

### Footnote Rules

| Rule | Detail |
|------|--------|
| 1:1 mapping | Each footnote ID must be referenced **exactly once** in document.xml. Multiple references to the same ID corrupt footnotes 1 and 2 |
| IDs must be sequential | No gaps — Word may silently renumber on recovery |
| Reserved IDs | id="-1" (separator) and id="0" (continuation) are reserved — real footnotes start at id="1" |
| Each paragraph needs paraId | Every `<w:p>` inside a footnote needs its own unique paraId |

The `validate_footnotes()` tool checks reference/definition matching. Always run after adding footnotes.

**Word Recovery Warning:** When Word recovers/repairs a file, it renumbers ALL footnotes sequentially by document position, not by original ID. After any Word recovery, re-examine footnotes before further edits.

### Non-Breaking Spaces

Word uses non-breaking spaces (`\xa0`, U+00A0) and narrow no-break spaces (`\u202f`, U+202F) throughout. **Direct string matching fails silently.** The `search_text()` tool handles this internally, but if you're checking text returned by `get_paragraph()`, be aware that what looks like a space may be `\xa0`.

### Text Inside Hyperlinks

Text inside `w:hyperlink` elements may not appear in `paragraph.text` from some parsers. The docx-mcp tools handle this correctly, but be aware when working with documents containing many hyperlinks — the visible text in Word may differ from what a naive text extraction shows.

### Tracked Changes: Key Rules

| Rule | Detail |
|------|--------|
| Use `w:delText` inside `w:del` | Never `w:t` — causes validation errors |
| Preserve `w:rPr` formatting | Copy the original run's formatting into tracked change runs |
| Minimal edits only | Mark only what changes — don't wrap entire paragraphs |
| Deleting entire paragraphs | Must also mark the paragraph mark as deleted via `<w:del/>` inside `<w:pPr><w:rPr>`, otherwise accepting changes leaves empty paragraphs |

### Smart Quotes

When adding text, use smart quotes for professional typography:

| Entity | Character |
|--------|-----------|
| `\u2018` | ' (left single) |
| `\u2019` | ' (right single / apostrophe) |
| `\u201C` | " (left double) |
| `\u201D` | " (right double) |
| `\u2013` | -- (en dash) |

The docx-mcp tools accept Unicode text directly — pass the actual characters, not XML entities.

### Element Order in `<w:pPr>`

The OOXML schema requires a specific element order: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last. Out-of-order elements cause validation warnings and may trigger Word recovery.

### Heading Numbering

**NEVER embed literal section numbers in heading text** (e.g., "1.1 Background"). Heading numbers must come from Word's multilevel list numbering system. If you need to insert a heading via `insert_text()`, insert only the heading text — the numbering comes from the document's styles.

## Markdown-to-DOCX Conversion Pitfalls

When converting markdown content to DOCX edits via docx-mcp, watch for:

### Soft-Wrapped Lines

A long line in markdown may display across multiple lines in an editor. If you're extracting text from markdown to feed into `insert_text()` or `delete_text()`, ensure you're working with the logical line, not the display line. A `[bracketed construct]` that wraps across source lines must be treated as one unit.

### Fake Footnotes in Markdown

Markdown doesn't have real footnotes — `[^1]` syntax is a convention that some parsers support. When converting markdown with footnote-style references to DOCX:

1. Use `add_footnote()` to create real OOXML footnotes — don't insert `[1]` as plain text
2. Map each markdown `[^N]` reference to an `add_footnote()` call on the appropriate paragraph
3. Remove the footnote definition text from the body (it now lives in footnotes.xml)

### Superscript Numbers Concatenate

When extracting text from paragraphs that contain footnote references, the superscript reference numbers are invisible in the XML text but adjacent characters concatenate. Example: "File #8" followed by superscript footnote "73" extracts as "File #873". Account for this when using `search_text()` with regex patterns on footnoted text.

## Audit Checklist

Before delivering any edited document, run through this checklist:

```
1. audit_document()         → comprehensive structural check
2. validate_footnotes()     → if any footnotes were added/modified
3. validate_paraids()       → if any structural changes were made
4. save_document("new.docx") → save to new file (preserve original)
```

The `audit_document()` tool checks all of the following in one call:
- XML well-formedness of all parts
- Footnote cross-references (references vs definitions, duplicates, gaps)
- ParaId uniqueness and range validity (< 0x80000000)
- Heading level continuity (no skips like H2 -> H4)
- Bookmark pairing (start/end matching)
- Relationship targets (all referenced files exist)
- Image references (all embedded images exist in word/media/)
- Content type registration (all media extensions have entries)
- Residual artifacts (DRAFT, TODO, FIXME, PLACEHOLDER, TBD markers)
