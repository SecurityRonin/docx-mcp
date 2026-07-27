# docx-mcp

[![PyPI](https://img.shields.io/pypi/v/docx-mcp-server?color=blue)](https://pypi.org/project/docx-mcp-server/)
[![Python](https://img.shields.io/pypi/pyversions/docx-mcp-server)](https://pypi.org/project/docx-mcp-server/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/SecurityRonin/docx-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/SecurityRonin/docx-mcp/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/badge/coverage-100%25-brightgreen)](https://github.com/SecurityRonin/docx-mcp)
[![Sponsor](https://img.shields.io/github/sponsors/h4x0r?logo=githubsponsors&label=Sponsor&color=%23ea4aaa)](https://github.com/sponsors/h4x0r)
[![SafeSkill 91/100](https://img.shields.io/badge/SafeSkill-91%2F100_Verified%20Safe-brightgreen)](https://safeskill.dev/scan/securityronin-docx-mcp)

<a href="https://glama.ai/mcp/servers/SecurityRonin/docx-mcp">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/SecurityRonin/docx-mcp/badges/card.svg" alt="docx-mcp MCP server" />
</a>

Give your AI coding agent the ability to create, read, and edit Word documents. Edits appear as tracked changes in Microsoft Word — red strikethrough for deletions, green underline for insertions — and after any revision session your agent can produce an email-ready change log listing every insertion, deletion, and replacement. Or compare two separate files to get the same output automatically.

## Who This Is For

**Professionals who produce Word deliverables and want their AI agent to handle the document work directly:**

- **Legal** — contract review with tracked redlines, batch clause replacement across templates, comment annotations explaining each change, footnote management
- **Security & Penetration Testing** — generate pentest reports from markdown findings, merge appendices from multiple engagements, add executive-summary comments, remove DRAFT watermarks before delivery
- **Consulting** — build proposals and SOWs from templates, convert meeting notes to formatted deliverables, bulk-update payment terms across document sets
- **Compliance & Audit** — structural validation of document integrity, cross-reference checking, heading-level audits, protection enforcement

## What You Can Ask Your Agent To Do

**Review a contract:**
> "Open contract.docx, find every instance of 'Net 30', change it to 'Net 60', and add a comment on each explaining it was updated per Amendment 3. Save as contract_revised.docx."

**Generate a report from markdown:**
> "Convert my pentest-findings.md to a Word document using the client's report template. Add footnotes for each CVE reference."

**Batch-edit a template library:**
> "Open the MSA template, replace 'ACME Corp' with 'GlobalTech Inc' everywhere, update the effective date in the header, and set document protection to track-changes-only."

**Audit a document before sending:**
> "Open the final deliverable, run a structural audit, check for any DRAFT or TODO markers, validate all footnote cross-references, and remove any watermarks."

Your agent handles the entire workflow — opening the file, navigating the structure, making precise edits with full revision history, validating integrity, and saving — while you focus on the substance.

## Installation

**Claude Code** (recommended):

```bash
claude mcp add docx-mcp -- uvx docx-mcp-server
```

That's it. A companion [skill](docx_mcp/skill/SKILL.md) auto-installs the first time the server starts, teaching Claude the document editing workflow, OOXML pitfalls, and audit checklist. It auto-updates on every upgrade.

<details>
<summary>Other platforms and installation methods</summary>

**Claude Desktop:**

Add to your MCP settings:

```json
{
  "mcpServers": {
    "docx-mcp": {
      "command": "uvx",
      "args": ["docx-mcp-server"]
    }
  }
}
```

**Cursor / Windsurf / VS Code:**

Add to your MCP configuration file:

```json
{
  "mcpServers": {
    "docx-mcp": {
      "command": "uvx",
      "args": ["docx-mcp-server"]
    }
  }
}
```

**OpenClaw:**

```yaml
mcpServers:
  docx-mcp:
    command: uvx
    args:
      - docx-mcp-server
```

**With pip:**

```bash
pip install docx-mcp-server
```

</details>

## Capabilities

| Capability | What your agent can do |
|---|---|
| **Create documents** | Start blank, from a `.dotx` template, or from markdown — headings, tables, lists, images, footnotes, code blocks, smart typography |
| **Track changes** | Insert, delete, and replace text as tracked revisions — underlined insertions, red strikethrough deletions. Pass `tracked=False` to any editing tool to write the change directly with no revision markup. Accept or reject changes by author. |
| **Change log** | After editing with tracked changes, call `generate_change_summary` for a numbered .txt listing every insertion, deletion, and replacement. Or use `diff_to_text` to compare two separate files and get the same output automatically — useful for document versions you didn't edit yourself. |
| **Comments** | Add comments anchored to specific paragraphs, reply to comment threads |
| **Find and replace** | Search by text or regex across body, footnotes, and comments — then make targeted edits |
| **Tables** | Create tables, modify cells, add or delete rows — all with revision tracking |
| **Footnotes & endnotes** | Add, list, and validate cross-references |
| **Formatting** | Bold, italic, underline, color — with revision tracking so formatting changes are visible |
| **Headers & footers** | Read and edit header/footer content with tracked changes |
| **Images** | List embedded images, insert new ones with specified dimensions |
| **Sections & layout** | Page breaks, section breaks, page size, orientation, margins |
| **Cross-references** | Internal hyperlinks between paragraphs with bookmarks |
| **Document merge** | Combine content from multiple DOCX files |
| **Protection** | Lock documents for tracked-changes-only, read-only, or comments-only with passwords |
| **Structural audit** | Validate footnotes, headings, bookmarks, images, and internal consistency before delivery |
| **Watermark removal** | Detect and strip DRAFT watermarks from headers |
| **PII scrubbing** *(experimental)* | Detect names, emails, phone numbers, SSNs, and more using Presidio + spaCy NER. Produces true XML redaction — original text permanently deleted, replaced with a solid black rectangle. **Not for production use** — see [limitations](#pii-scrubbing-experimental). |

## Example: Contract Review with Redlines

```
1. open_document("services-agreement.docx")
2. get_headings()                              → see document structure
3. search_text("30 days")                      → find the payment clause
4. delete_text(para_id, "30 days")             → tracked deletion  (red strikethrough)
5. insert_text(para_id, "60 days")             → tracked insertion (green underline)
6. add_comment(para_id, "Extended per client request — see Amendment 3")
7. audit_document()                            → verify structural integrity
8. save_document("services-agreement_redlined.docx")
```

Open the output in Word and you see exactly what a human reviewer would produce — revision marks, comments in the margin, clean document structure.

## Example: Change Log After a Revision Session

```
1. open_document("contract.docx")
2. replace_text(para_id, "30 days", "60 days")           → tracked revision (del + ins)
3. add_comment(para_id, "Extended per Amendment 3")
4. save_document("contract_redlined.docx")
5. generate_change_summary("contract_changes.txt")
```

The `.txt` output reads like:

```
1. REPLACEMENT by Claude on 2026-05-25
   Deleted:  "30 days"
   Inserted: "60 days"
```

Paste it into an email, a PR description, or a handover note.

## Example: Compare Two Separate Document Versions

No editing session required — point it at any two files:

```
1. diff_to_text("proposal_v1.docx", "proposal_v2.docx")
```

Returns two files:
- `proposal_v1_diff.docx` — open in Word to see tracked changes inline
- `proposal_v1_diff.txt` — plain-text summary, ready to paste anywhere

## Example: Pentest Report from Markdown

```
1. create_from_markdown("pentest-report.docx",
       md_path="findings.md",
       template_path="client-template.dotx")
2. audit_document()                            → verify integrity
3. save_document()                             → ready for delivery
```

Your markdown findings — headings, tables of affected hosts, code blocks with proof-of-concept output, severity ratings — become a formatted Word document matching the client's template. Smart typography is applied automatically (curly quotes, em dashes, proper ellipses).

## PII Scrubbing (Experimental)

> **This feature is experimental and NOT suitable for production use or as the sole control for privileged, regulated, or legally sensitive documents.**

`scrub_pii` uses Microsoft Presidio with a spaCy NER model (`en_core_web_lg`, ~560 MB, downloaded on first call) to detect and permanently redact personal information. Redacted spans are replaced with solid black DrawingML rectangles — the original text is deleted from the XML, not merely hidden.

**What it detects reliably** (via regex/pattern matching, high confidence):
- Email addresses, phone numbers, credit card numbers, SSNs, IBANs, IP addresses, US passport numbers

**What it detects statistically** (via NER, will have misses):
- Person names, organisation names, locations, dates

**Known NER gaps — these will be missed:**
- Names in ALL-CAPS (common in ledgers, table headers, legal party designations)
- Single-token names with no surrounding context
- Non-English names (Arabic, CJK, African name patterns have low recall on this English model)
- Names embedded in legal boilerplate (`Lender: Jane Doe`, `Authorized Signatory: John Smith`)

**Recommended workflow:**

```
1. open_document("sensitive.docx")
2. scrub_pii(dry_run=True)             → review detected entities — check for misses
3. scrub_pii("sensitive_redacted.docx") → produce redacted copy
4. open the output in Word and verify manually before sharing
```

Always treat the output as a first pass requiring human review, not a finished redaction.

## How It Works

A `.docx` file is a ZIP archive of XML files. This server unpacks the archive, edits the XML directly, and repacks it. This is what gives it the ability to produce real tracked changes, comments, and footnotes — things that higher-level document libraries can't do.

Every edit is validated against the OOXML specification before saving, catching issues like orphaned footnotes, duplicate internal IDs, and broken cross-references that would otherwise cause Word to "repair" (and silently rewrite) your document.

<details>
<summary>Full tool inventory (200+ tools)</summary>

### Document Lifecycle

| Tool | Description |
|---|---|
| `open_document` | Open a .docx file for reading and editing |
| `create_document` | Create a new blank .docx (or from a .dotx template) |
| `create_from_markdown` | Create a .docx from GitHub-Flavored Markdown |
| `close_document` | Close the current document and clean up |
| `get_document_info` | Get overview stats (paragraphs, headings, footnotes, comments) |
| `save_document` | Save changes back to .docx (can overwrite or save to new path) |

### Reading

| Tool | Description |
|---|---|
| `get_headings` | Get heading structure with levels and text |
| `search_text` | Search across body, footnotes, and comments (text or regex) |
| `get_paragraph` | Get full text and style of a paragraph |

### Track Changes

| Tool | Description |
|---|---|
| `insert_text` | Insert text as tracked revision (underlined in Word) — or directly with `tracked=False` |
| `delete_text` | Mark text deleted (red strikethrough in Word) — or remove directly with `tracked=False` |
| `replace_text` | Replace text as a tracked del+ins pair — or directly with `tracked=False` |
| `modify_cell` | Modify a table cell — tracked by default, `tracked=False` for direct writes |
| `edit_header_footer` | Edit header/footer text — tracked by default, `tracked=False` for direct writes |
| `set_formatting` | Apply bold/italic/underline/color with tracked-change markup |
| `set_track_changes` | Enable or disable Word's built-in track-changes recording flag |
| `get_tracked_changes` | List all pending tracked changes (author, date, type, text) |
| `accept_changes` | Accept tracked changes, optionally filtered by author |
| `reject_changes` | Reject tracked changes, optionally filtered by author |

### Change Log

| Tool | Description |
|---|---|
| `generate_change_summary` | Export all tracked changes in the open document as a numbered .txt file — insertions, deletions, and adjacent del+ins pairs grouped as replacements |
| `compare_documents` | Diff two DOCX files and produce a tracked-change DOCX (Word-openable) |
| `diff_to_text` | Diff two DOCX files and produce both a tracked-change DOCX and a plain-text summary .txt |

### Tables

| Tool | Description |
|---|---|
| `get_tables` | Get all tables with row/column counts and cell content |
| `add_table` | Insert a new table after a paragraph |
| `modify_cell` | Modify a table cell with tracked changes |
| `add_table_row` | Add a row to a table |
| `delete_table_row` | Delete a table row with tracked changes |

### Lists

| Tool | Description |
|---|---|
| `add_list` | Apply bullet or numbered list formatting to paragraphs |

### Comments

| Tool | Description |
|---|---|
| `get_comments` | List all comments with ID, author, date, and text |
| `add_comment` | Add a comment anchored to a paragraph |
| `reply_to_comment` | Reply to an existing comment (threaded) |

### Footnotes & Endnotes

| Tool | Description |
|---|---|
| `get_footnotes` | List all footnotes with ID and text |
| `add_footnote` | Add a footnote with superscript reference |
| `validate_footnotes` | Validate footnote cross-references |
| `get_endnotes` | List all endnotes with ID and text |
| `add_endnote` | Add an endnote with superscript reference |
| `validate_endnotes` | Validate endnote cross-references |

### Headers, Footers & Styles

| Tool | Description |
|---|---|
| `get_headers_footers` | Get all headers and footers with text content |
| `edit_header_footer` | Edit header/footer text with tracked changes |
| `get_styles` | Get all defined styles |

### Properties & Images

| Tool | Description |
|---|---|
| `get_properties` | Get core document properties (title, creator, dates) |
| `set_properties` | Set core document properties |
| `get_images` | Get all embedded images with dimensions |
| `insert_image` | Insert an image after a paragraph |

### Sections & Cross-References

| Tool | Description |
|---|---|
| `add_page_break` | Insert a page break after a paragraph |
| `add_section_break` | Add a section break (nextPage, continuous, evenPage, oddPage) |
| `set_section_properties` | Set page size, orientation, and margins |
| `add_cross_reference` | Add a cross-reference link between paragraphs |

### Protection & Merge

| Tool | Description |
|---|---|
| `set_document_protection` | Set document protection with optional password |
| `merge_documents` | Merge content from another DOCX |

### Validation & Audit

| Tool | Description |
|---|---|
| `validate_paraids` | Check internal ID uniqueness across all document parts |
| `remove_watermark` | Remove VML watermarks from document headers |
| `audit_document` | Comprehensive structural audit |

</details>

## Requirements

- Python 3.10+
- Works on macOS, Linux, and Windows

---

[Privacy Policy](https://securityronin.github.io/docx-mcp/privacy/) · [Terms of Service](https://securityronin.github.io/docx-mcp/terms/) · © 2026 Security Ronin Ltd

<img referrerpolicy="no-referrer-when-downgrade" src="https://static.scarf.sh/a.png?x-pxid=95beebbb-0f2e-46cc-9a68-a8e66f613180" />
