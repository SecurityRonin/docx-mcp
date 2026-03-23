# docx-mcp

[![PyPI](https://img.shields.io/pypi/v/docx-mcp-server?color=blue)](https://pypi.org/project/docx-mcp-server/)
[![Python](https://img.shields.io/pypi/pyversions/docx-mcp-server)](https://pypi.org/project/docx-mcp-server/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/SecurityRonin/docx-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/SecurityRonin/docx-mcp/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/badge/coverage-100%25-brightgreen)](https://github.com/SecurityRonin/docx-mcp)
[![Sponsor](https://img.shields.io/github/sponsors/h4x0r?logo=githubsponsors&label=Sponsor&color=%23ea4aaa)](https://github.com/sponsors/h4x0r)

<a href="https://glama.ai/mcp/servers/SecurityRonin/docx-mcp">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/SecurityRonin/docx-mcp/badges/card.svg" alt="docx-mcp MCP server" />
</a>

MCP server for reading and editing Word (.docx) documents with track changes, comments, footnotes, and structural validation.

The only cross-platform MCP server that combines **track changes**, **comments**, and **footnotes** in a single package — with OOXML-level structural validation that no other server offers.

## Features

| Capability | Description |
|---|---|
| **Track changes** | Insert/delete text with proper `w:ins`/`w:del` markup — shows as revisions in Word |
| **Comments** | Add comments, reply to threads, read existing comments |
| **Footnotes** | Add footnotes, list all footnotes, validate cross-references |
| **ParaId validation** | Check uniqueness across all document parts (headers, footers, footnotes) |
| **Watermark removal** | Detect and remove VML watermarks (e.g., DRAFT) from headers |
| **Structural audit** | Validate footnotes, paraIds, heading levels, bookmarks, relationships, images |
| **Text search** | Search across body, footnotes, and comments — plain text or regex |
| **Heading extraction** | Get the full heading structure with levels and paragraph IDs |

## Installation

```bash
# Claude Code (recommended)
claude mcp add docx-mcp -- uvx docx-mcp-server

# With pip
pip install docx-mcp-server

# With uvx
uvx docx-mcp-server
```

> **Optional:** Install the companion [skill](skill/SKILL.md) for Claude Code — it teaches Claude when and how to use the tools automatically:
> ```bash
> curl -sSL https://raw.githubusercontent.com/SecurityRonin/docx-mcp/main/install.sh | bash
> ```

## Configuration

### Claude Desktop / Claude Code

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

### Cursor / Windsurf / VS Code

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

### OpenClaw

Add to your `openclaw.yaml`:

```yaml
mcpServers:
  docx-mcp:
    command: uvx
    args:
      - docx-mcp-server
```

Or via the CLI:

```bash
openclaw config set mcpServers.docx-mcp.command "uvx"
openclaw config set mcpServers.docx-mcp.args '["docx-mcp-server"]'
```

### With pip install

```json
{
  "mcpServers": {
    "docx-mcp": {
      "command": "docx-mcp"
    }
  }
}
```

## Available Tools

### Document Lifecycle

| Tool | Description |
|---|---|
| `open_document` | Open a .docx file for reading and editing |
| `close_document` | Close the current document and clean up |
| `get_document_info` | Get overview stats (paragraphs, headings, footnotes, comments) |
| `save_document` | Save changes back to .docx (can overwrite or save to new path) |

### Reading

| Tool | Description |
|---|---|
| `get_headings` | Get heading structure with levels, text, styles, and paraIds |
| `search_text` | Search across body, footnotes, and comments (text or regex) |
| `get_paragraph` | Get full text and style of a paragraph by paraId |

### Track Changes

| Tool | Description |
|---|---|
| `insert_text` | Insert text with tracked-change markup (`w:ins`) |
| `delete_text` | Mark text as deleted with tracked-change markup (`w:del`) |

### Comments

| Tool | Description |
|---|---|
| `get_comments` | List all comments with ID, author, date, and text |
| `add_comment` | Add a comment anchored to a paragraph |
| `reply_to_comment` | Reply to an existing comment (threaded) |

### Footnotes

| Tool | Description |
|---|---|
| `get_footnotes` | List all footnotes with ID and text |
| `add_footnote` | Add a footnote with superscript reference |
| `validate_footnotes` | Cross-reference IDs between document body and footnotes.xml |

### Validation & Audit

| Tool | Description |
|---|---|
| `validate_paraids` | Check paraId uniqueness and range validity across all parts |
| `remove_watermark` | Remove VML watermarks from document headers |
| `audit_document` | Comprehensive structural audit (footnotes, paraIds, headings, bookmarks, relationships, images, artifacts) |

## Example Workflow

```
1. open_document("/path/to/contract.docx")
2. get_headings()                          → see document structure
3. search_text("30 days")                  → find the clause
4. delete_text(para_id, "30 days")         → tracked deletion
5. insert_text(para_id, "60 days")         → tracked insertion
6. add_comment(para_id, "Extended per client request")
7. audit_document()                        → verify structural integrity
8. save_document("/path/to/contract_revised.docx")
```

The resulting document opens in Microsoft Word with proper revision marks — deletions shown as red strikethrough, insertions as green underline, comments in the sidebar.

## How It Works

A .docx file is a ZIP archive of XML files. This server:

1. **Unpacks** the archive to a temporary directory
2. **Parses** all XML parts with lxml and caches them in memory
3. **Edits** the cached DOM trees directly (no intermediate abstraction layer)
4. **Repacks** modified XML back into a valid .docx archive

This approach gives full control over OOXML markup — essential for track changes (`w:ins`/`w:del`), comments (`w:comment` + range markers), and structural validation that higher-level libraries like python-docx don't expose.

## Requirements

- Python 3.10+
- lxml

## License

MIT

<img referrerpolicy="no-referrer-when-downgrade" src="https://static.scarf.sh/a.png?x-pxid=95beebbb-0f2e-46cc-9a68-a8e66f613180" />
