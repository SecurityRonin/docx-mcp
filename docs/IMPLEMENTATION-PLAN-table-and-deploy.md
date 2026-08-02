# Implementation Plan — Table Formatting, Deployment & Safety, Document Protection

Incorporating features from two MIT-licensed competitors into docx-mcp's XML-direct architecture:

- **GongRzhe/Office-Word-MCP-Server** (Python, python-docx-based, archived Dec 2025) — table formatting surface, protection concepts
- **hongkongkiwi/docx-mcp** (Rust) — deployment architecture: readonly mode, tool filtering, resource limits, single-binary distribution, LibreOffice-free PDF

Source for every claim below: direct reading of all three codebases (docx-mcp at v0.7.4, both competitor repos cloned at HEAD, 2026-08-01). Adversarially reviewed against `b141be8` — corrections integrated inline (original review by Codex, 2026-08-02).

## Executive Summary

docx-mcp already covers most of GongRzhe's table surface (merge, per-cell shading, borders, widths, styles — 26 table tools exist today). The genuine gaps are **six formatting capabilities** (banding, header-row styling, cell padding, table-level cell margins, layout/auto-fit mode, table width) and **one structural fix** (schema-ordered `tcPr`/`tblPr` child insertion, which current code does not enforce). Part A adds 9 tools and one ordering helper, and fixes a false-positive in the save-time table validator (at **two** sites — `base.py` and `validation.py`).

Part B ports hongkongkiwi's safety model to the Python server: a config layer (env + CLI, resolved **before importing `server.py`** — decorators register at import time), read/write tool classification with filtering at registration *and* call time, configurable resource caps, and a PyInstaller `--onedir` distribution channel. PDF export gains a pure-Python fallback (ReportLab + bundled Liberation fonts) so the tool works without LibreOffice, at honestly-labeled reduced fidelity.

Part C **repairs** the existing `w:documentProtection` password hash (the spin loop concatenates salt where the spec requires an iteration counter — Word cannot verify the password docx-mcp writes) and extends it with editable ranges (`w:permStart`/`w:permEnd` — the real OOXML mechanism GongRzhe only simulated in a JSON sidecar), protection-state reading, and file-level password encryption via msoffcrypto. GongRzhe's "digital signatures" are **not** ported: they are a content-hash simulation, not XML-DSig — placeholder crypto that fabricates a security signal. We instead add a signature *inspection* tool and defer real signing.

Estimated scope: 16 new MCP tools (219 → 235), one new module (`config.py`), two extended modules (`tables.py`, `protection.py`), one rewritten fallback path (`pdfexport.py`), CI packaging matrix. Strict TDD throughout (RED commit, then GREEN commit, per repo discipline).

**M4 (distribution) is ON HOLD** pending two decisions: fpdf2's LGPL-3.0-only license conflicts with PyInstaller redistribution (§B.6), and macOS notarisation is unbudgeted (§Risks).

---

## 0. Current State (verified against source)

### docx-mcp architecture

- `docx_mcp/server.py` (3 087 lines): FastMCP server, **219 `@mcp.tool()` functions**, each a thin wrapper over a `DocxDocument` method. Documents keyed by `document_handle` in a module-level `_docs` dict; `_store()` inserts, `_resolve()` fetches.
- `docx_mcp/document/` — `DocxDocument` composed of ~45 mixins. Parts loaded as lxml trees via `_tree()`/`_require()`, dirty-tracked via `_mark()`, written back in `base.save()`.
- Save pipeline: `_pre_save_repair()` (orphan foot/endnotes, duplicate paraIds, broken rels) → `_post_repair_warnings()` (heading skips, unpaired bookmarks, **naive table column-count check**, DRAFT/TODO markers) → serialize → re-zip.
- `guards.py` `InputGuard`: para-ID regex, hex color, bounded ints, output-path traversal check (`.docx` suffix only, never applied to PDF paths — **⚠ arbitrary-file-write via `convert_to_pdf`**), `MAX_FILE_SIZE = 100 MB` (**⚠ dead code** — defined but referenced nowhere; no size gate exists at open or save).
- `cli.py`: no argparse — bare dispatch (`install-skill` | server). No flags, no env-var config.
- `protection.py`: writes `w:documentProtection` in settings.xml with SHA-512 + 100 000 spin-count. **⚠ BUG: the spin loop iterates `H(H_{n-1} + salt)` but MS-OFFCRYPTO §2.4.2.4 requires `H(H_{n-1} + iterator)` where iterator is a 32-bit LE counter. Word cannot verify the password. H0 and UTF-16-LE encoding are correct; the defect is one line. Existing test asserts only `has_password is True`, never round-trips through Word.** Types: trackedChanges, comments, readOnly, forms, none.
- `pdfexport.py`: LibreOffice-headless only; **hard-fails without LibreOffice**; calls `self.save(self.source_path)` — i.e. `convert_to_pdf` currently *writes the source file* as a side effect.
- Existing table tools (26): `get_tables`, `add_table`, `modify_cell`, `add_table_row`, `delete_table_row`, `merge_cells` (rectangular, gridSpan + vMerge), `set_header_row` (tblHeader), `set_column_widths`, `csv_to_table`, `table_to_csv`, `delete_table`, `add_column_to_table`, `delete_column_from_table`, `set_cell_width`, `set_cell_vertical_alignment`, `set_row_height`, `set_table_alignment`, `set_table_borders`, `set_cell_shading`, `set_table_style`, `split_table`, `duplicate_table_row`, `sort_table`, `get_table`, `get_cell_text`, `copy_table`. (Tracked changes are a `tracked=` parameter on CRUD tools, not separate registrations.)

### GongRzhe (what's actually there)

`word_document_server/core/tables.py` (866 lines, python-docx + raw OxmlElement): cell borders, table styling, per-cell shading, **alternating-row shading** (direct per-cell `w:shd`, not style-based banding), **header-row highlight**, merge (rect/horizontal/vertical), cell/table alignment, column widths (dxa/pct/auto), **table width** (`w:tblW`), **auto-fit** (`w:tblLayout type="autofit"` + all columns `auto`), cell text formatting, **cell padding** (`w:tcMar`, dxa or pct).

`core/protection.py` + `tools/protection_tools.py`: `protect_document` does **real file encryption** via `msoffcrypto.OfficeFile.encrypt(password=...)`. `add_restricted_editing` and `add_digital_signature` write a **JSON sidecar file** (`<name>.protection`) — no OOXML elements at all; the "signature" is a SHA-256 content hash plus signer name in that sidecar. `verify_document` reads the sidecar back.

### hongkongkiwi (what's actually there)

`src/security.rs` (339 lines): clap `Args` with env-var mirrors (`DOCX_MCP_READONLY`, `DOCX_MCP_WHITELIST`, `DOCX_MCP_BLACKLIST`, `DOCX_MCP_SANDBOX`, `DOCX_MCP_NO_EXTERNAL_TOOLS`, `DOCX_MCP_NO_NETWORK`, `DOCX_MCP_MAX_SIZE`, `DOCX_MCP_MAX_DOCS`). `SecurityConfig` defaults: 100 MB max doc, 50 max open docs. `is_command_allowed()` order: readonly-command check first (always allowed), then **whitelist takes precedence over blacklist**. Sandbox mode = disable external tools + network. A static `get_readonly_commands()` set classifies ~each tool; exports (`export_to_pdf`, `export_to_markdown`, preview) count as readonly because they never modify the source.

`src/converter.rs` + `src/pure_converter.rs` + `src/fonts.rs`: tiered PDF chain — LibreOffice → unoconv → `basic_docx_to_pdf` (pure Rust: printpdf + lopdf + rusttype). Five fonts embedded via `include_bytes!` (LiberationSans Regular/Bold/Italic, LiberationMono, NotoSans). Feature flags: `embedded-fonts`, `pure-rust-pdf` default; `external-tools` optional.

---

## Part A — Table Formatting (from GongRzhe)

### A.1 Gap analysis

| GongRzhe capability | docx-mcp today | Action |
|---|---|---|
| `merge_cells` / `_horizontal` / `_vertical` | `merge_cells` handles rectangular, horizontal, vertical | **None** (audit only, §A.5) |
| `set_cell_shading_by_position` | `set_cell_shading` | None |
| `set_cell_border` | `set_table_borders` (table-level only) | **Add** per-cell borders |
| `apply_alternating_row_shading` | — | **Add** `set_table_banding` |
| `highlight_header_row` | `set_header_row` sets repeat-header only, no styling | **Add** `style_header_row` |
| `set_cell_padding` (`w:tcMar`) | — | **Add** `set_cell_padding` |
| (table default margins, `w:tblCellMar`) | — | **Add** `set_table_cell_margins` |
| `auto_fit_table` (`w:tblLayout`) | — | **Add** `set_table_layout` |
| `set_table_width` (`w:tblW`) | — | **Add** `set_table_width` |
| `set_cell_alignment` (h + v combined) | vertical only (`set_cell_vertical_alignment`) | **Add** `set_cell_alignment` (unified; horizontal = `w:jc` on cell paragraphs) |
| `format_cell_text` | achievable via `modify_cell` + run tools on cell paraIds, but needs 3+ calls | **Add** `format_cell` convenience |
| `set_column_width(index)` single col | `set_column_widths` (all) + `set_cell_width` | None (composable) |

Not ported: GongRzhe's `copy_table` (exists), `apply_table_style` (exists as `set_table_style`), filename-per-call API shape (docx-mcp is handle-based; all new tools follow the existing `(table_idx, …, document_handle="")` convention).

### A.2 OOXML elements and the ordering seam

All references: ECMA-376 Part 1 (Fundamentals and Markup Language Reference), §17.4 "Tables". Element names below are the normative anchors; cross-check exact subsection numbers against the spec PDF at implementation time rather than trusting numbers here.

**The one structural change this work needs** (per the "one small honest seam" rule): `w:tcPr` and `w:tblPr` are `xsd:sequence` types — children must appear in schema order. Current code appends via `etree.SubElement` (verified in `merge_cells`, `set_cell_shading`), which produces out-of-order children when a cell already has properties (e.g. `shd` before `gridSpan` if shading was applied first). Word tolerates this; strict validators and some toolchains do not. Every new tool goes through one helper:

```python
# document/ooxml_order.py (new, ~60 lines)
TCPR_ORDER = ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
              "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
              "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange")
TBLPR_ORDER = ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
               "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
               "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
               "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange")
TCMAR_ORDER = ("top", "left", "bottom", "right")   # transitional schema

def ordered_set_child(parent, tag_localname, order, nsmap=W): ...
    # find-or-create child, inserting at the schema-correct position
```

Existing table tools are migrated to this helper opportunistically (when a test exposes an ordering violation), not in a bulk reformat — minimal diffs.

Elements per new tool:

- **Banding / header styling**: direct `w:shd` (`w:val` pattern, `w:fill` hex, `w:color`) per `w:tc`, same as today's `set_cell_shading`. The *style-based* alternative — `w:tblLook` attributes (`firstRow`, `lastRow`, `firstColumn`, `lastColumn`, `noHBand`, `noVBand`) activating a table style's `w:tblStylePr w:type="band1Horz"` etc. — is exposed as an option (`method="direct"|"style"`). Direct shading survives row insertion incorrectly (colors don't re-alternate); style-based banding re-computes but needs a style with band definitions. Document the trade-off in the tool docstring; default `direct` (works with any document, matches GongRzhe behavior).
- **Cell padding**: `w:tcMar` inside `w:tcPr` with child elements `top/left/bottom/right`, each `w:w` (twentieths of a point) + `w:type="dxa"` (units in the tool API are **mm**, converted, consistent with `set_cell_width(width_mm)` — not GongRzhe's raw dxa).
- **Table default margins**: `w:tblCellMar` inside `w:tblPr`, same child shape.
- **Layout mode**: `w:tblLayout w:type="fixed"|"autofit"` in `w:tblPr`; autofit also sets `w:tblW w:type="auto" w:w="0"` and clears explicit `w:tcW` widths (GongRzhe sets columns to `auto`; we clear `tcW` to `auto` type).
- **Table width**: `w:tblW` with `w:type` `dxa` (absolute; API takes mm) or `pct` (fiftieths of a percent; API takes 0–100 float) or `auto`.
- **Per-cell borders**: `w:tcBorders` in `w:tcPr`, children `top/left/bottom/right/insideH/insideV/tl2br/tr2bl`, each with `w:val` (border style), `w:sz` (eighths of a point), `w:color`.
- **Cell alignment**: vertical `w:vAlign` in `w:tcPr` (exists); horizontal `w:jc` on each `w:p` inside the `w:tc`.

### A.3 New MCP tools (Part A)

All follow existing conventions: `document_handle: str = ""` last parameter, return compact-JSON string via `_js()`, raise `ValueError`/`IndexError` through the FastMCP error path, colors validated by `InputGuard.color_hex`, indices by `InputGuard.bounded_int`.

| Tool | Parameters | Returns |
|---|---|---|
| `set_table_banding` | `table_idx:int, odd_color:str="FFFFFF", even_color:str="F2F2F2", skip_header:bool=True, method:str="direct"` | `{table_idx, rows_shaded, method}` |
| `style_header_row` | `table_idx:int, fill_color:str="4472C4", text_color:str="FFFFFF", bold:bool=True` | `{table_idx, cells_styled}` |
| `set_cell_padding` | `table_idx:int, row_idx:int, col_idx:int, top_mm:float\|None, bottom_mm:float\|None, left_mm:float\|None, right_mm:float\|None` | `{table_idx, row_idx, col_idx, padding_mm:{…}}` |
| `set_table_cell_margins` | `table_idx:int, top_mm, bottom_mm, left_mm, right_mm` (same optional shape) | `{table_idx, margins_mm:{…}}` |
| `set_table_layout` | `table_idx:int, mode:str` (`"autofit"`\|`"fixed"`) | `{table_idx, mode}` |
| `set_table_width` | `table_idx:int, width:float\|None, unit:str="mm"` (`"mm"`\|`"percent"`\|`"auto"`) | `{table_idx, width, unit}` |
| `set_cell_borders` | `table_idx:int, row_idx:int, col_idx:int, sides:list[str]\|None, style:str="single", color:str="000000", size:int=4` | `{table_idx, row_idx, col_idx, sides}` |
| `set_cell_alignment` | `table_idx:int, row_idx:int, col_idx:int, horizontal:str\|None, vertical:str\|None` | `{table_idx, row_idx, col_idx, horizontal, vertical}` |
| `format_cell` | `table_idx:int, row_idx:int, col_idx:int, bold, italic, underline, color, font_size_pt, font_name` (all optional) | `{table_idx, row_idx, col_idx, runs_formatted}` |

Implementation lands in `document/tables.py` (`TablesMixin`) — same file, same idiom as the 27 existing methods. `format_cell` delegates to existing run-formatting internals rather than duplicating `rPr` logic (DRY: search `formatting.py` for the run-property writer first).

**Tracked changes**: OOXML supports `w:tblPrChange`/`w:tcPrChange` for revision-tracked property edits. None of the new tools write them in phase 1 (GongRzhe doesn't either; the existing docx-mcp property tools — `set_cell_shading`, `set_table_borders` — don't either, so this is consistent). A follow-up "tracked table formatting" work item is recorded in §Remaining-work rather than speculatively built (YAGNI).

### A.4 Validation integration

1. **Fix the existing false positive** (at **two** sites): `_post_repair_warnings()` in `base.py:381-385` flags "inconsistent column counts" by counting `w:tc` per row. A legitimately merged table (gridSpan) trips it. Replace with effective-grid width: `sum(int(gridSpan.val or 1) for tc in row)` compared against `len(w:tblGrid/w:gridCol)`. **The identical bug exists at `validation.py:123-132`** in `validate_document` — fix both or the MCP tool and save-time warning disagree. Also handle `w:hMerge` (legacy horizontal-merge: continuation cells are physically present with no gridSpan) and `w:cellDel` (tracked-deleted cells).
2. **New warning**: `w:vMerge val="continue"` cell whose column has no `restart` above it (orphan continuation — renders unpredictably).
3. **New warning**: `w:tcPr`/`w:tblPr` children out of schema order (cheap check against the `ooxml_order` tables). **⚠ Sequence this AFTER migrating existing writers** — 101 `SubElement` calls across the package already violate ordering (including `set_document_protection` appending at the end of `settings.xml`). Shipping the detector before the fix reads as a regression.
4. All new writers call `_mark("word/document.xml")` so the save pipeline serializes them — no new pipeline hooks needed.

### A.5 Test strategy (Part A)

Per repo discipline: strict TDD, RED commit then GREEN commit per tool; tiering labeled per the project's evidence standard.

- **T3 unit tests** (`tests/test_table_formatting_extended.py`): per tool — element created, schema-ordered position, value conversion (mm→dxa, percent→pct fiftieths), idempotent re-application replaces rather than duplicates, out-of-range indices raise. Adversarial: apply shading then merge then padding on the same cell and assert `tcPr` child order.
- **T2 spec-schema validation**: add an env-gated test that validates saved `document.xml` against the ECMA-376 transitional `wml.xsd` (schemas are freely downloadable from ECMA; commit them under `tests/schemas/` if redistribution terms allow — verify — else download in CI and skip-when-absent locally). Ground truth is the published schema, independent of our code.
- **T2 round-trip oracle**: existing pattern in the repo (LibreOffice available in CI): `soffice --headless --convert-to docx` re-save of a document formatted by every new tool; assert LibreOffice preserves the elements (banding shd values, tcMar, tblLayout survive round-trip). Skip-when-absent.
- **T1 real-world fixture**: source one real .docx containing banded/merged/padded tables authored in Microsoft Word (provenance-documented per fleet standard in `tests/fixtures/README.md`), assert `get_table`/readers report it correctly and that a no-op open→save round-trip is byte-stable on the table parts.
- Coverage backstop: hypothesis fuzz on `set_table_width`/`set_cell_padding` numeric ranges (existing `hypothesis` dev-dep).

---

## Part B — Deployment & Safety (from hongkongkiwi)

### B.1 Config layer

New `docx_mcp/config.py`, resolved **once at server startup** (before tool registration), mirroring hkk's env-first design so container deployments need no wrapper scripts:

| Env var | CLI flag | Default | Meaning |
|---|---|---|---|
| `DOCX_MCP_READONLY` | `--readonly` | off | Only read-class tools exposed |
| `DOCX_MCP_TOOL_ALLOWLIST` | `--allow-tools a,b,c` | unset | If set, only these tools registered (precedence over blocklist, matching hkk) |
| `DOCX_MCP_TOOL_BLOCKLIST` | `--deny-tools a,b,c` | unset | These tools never registered |
| `DOCX_MCP_MAX_FILE_SIZE` | `--max-file-size BYTES` | 104857600 | Open/save size gate (replaces the `InputGuard.MAX_FILE_SIZE` constant) |
| `DOCX_MCP_MAX_OPEN_DOCS` | `--max-open-docs N` | 50 | Cap on `_docs` handles (hkk default) |
| `DOCX_MCP_NO_EXTERNAL_TOOLS` | `--no-external-tools` | off | Never spawn subprocesses (LibreOffice); PDF falls back to pure-Python |
| `DOCX_MCP_WORKDIR_ROOT` | `--workdir-root PATH` | unset | Confine open/save paths to this directory tree |
| `DOCX_MCP_SANDBOX` | `--sandbox` | off | Shorthand: no-external-tools + workdir-root required + readonly defaults off but caps enforced |

`cli.py` grows a real argparse layer (currently bare `sys.argv` dispatch); flags override env vars. **Config must resolve before `from docx_mcp.server import main`** — `@mcp.tool()` decorators execute at module import, so filtering at `mcp.run()` is too late. The seam exists: `cli.py:run_server()` does the import inside the function body.

**`DOCX_MCP_NO_NETWORK`** / `--no-network`: **required, not dead.** `pii.py:62-73` calls `spacy.cli.download()` on first `scrub_pii` invocation, pulling ~560 MB over the network. A grep gate for `requests/urllib/socket` cannot catch this (spaCy networks internally, several layers down). Enforce structurally: when `--no-network` is set, refuse to auto-download and require the model to be pre-installed.

### B.2 Readonly mode and tool filtering

Mechanism (Python/FastMCP equivalent of hkk's `is_command_allowed`):

1. **Classification at source**: replace bare `@mcp.tool()` with a local decorator `@doc_tool(write=True|False)` recorded in a registry. One mechanical sweep over server.py classifies all 219 tools (read: open/close/get_*/list_*/search/validate/audit/export/compare/statistics/convert_to_pdf-after-B.4-fix; write: everything that calls `_mark` transitively). The classification lives next to each tool — reviewable in one diff.
2. **Filtering at registration**: `doc_tool` consults the resolved config and simply *does not register* disallowed tools with FastMCP. The client never sees them in `tools/list` — this also cuts LLM context cost, a real benefit at 219+ tools (an allowlist deployment exposing 20 tools saves thousands of tokens per session).
3. **Defense in depth at call time**: the decorator also wraps the function to re-check the config before executing, so a stale client cache or registration bug cannot execute a write in readonly mode (secure-by-design: the wrong thing is structurally unreachable, not doc-discouraged).
4. Readonly-classified tools are always available regardless of blocklist, matching hkk semantics — except when an explicit allowlist is set, which is absolute (whitelist precedence, hkk-verified behavior).

Tests (T3): readonly server lists no write tools; direct dispatch of a write tool name errors; allowlist of 3 exposes exactly 3; allowlist+blocklist → allowlist wins. **Structural meta-test**: exercise each read-classified tool against a fixture and assert `doc._modified` is unchanged and no new file appeared on disk — this tests the *property* the classification claims, not just the presence of a decorator argument (a presence check catches missing annotations but not wrong-direction ones, which is the failure that breaks readonly). Split the 219-tool classification sweep by functional area, each pair of commits gated by this structural test.

### B.3 Resource limits

- **File size**: gate in `open()` *and* in `write_part`/save (config-driven; `InputGuard` gets the limit injected instead of a class constant). **⚠ This is new enforcement, not a refactor** — `MAX_FILE_SIZE` is dead code today (defined but never read). Documents over 100 MB that open today will start failing. Document this as a behaviour change.
- **Zip-bomb guard** (new, not in either competitor but required once size limits are advertised): on open, sum `ZipInfo.file_size` (uncompressed) and entry count; reject > 10× max-file-size uncompressed or > 10 000 parts. Loud diagnostic naming the offending totals per the fail-loud rule.
- **Open-document cap**: `_store()` refuses beyond `max_open_docs` with a message listing current handles.
- **Memory**: document option `--max-memory-mb` implemented via `resource.setrlimit(RLIMIT_AS)` on POSIX, warn-and-ignore on Windows. Off by default (a hard RLIMIT kills lxml mid-parse with MemoryError — acceptable for sandbox deployments, wrong as default).
- **Path confinement**: `--workdir-root` enforced in `InputGuard.output_path` *and* a new `InputGuard.input_path` used by `open_document` (the base.py comment notes confinement is currently enforced only at save/copy call sites; this closes the read side).

### B.4 `convert_to_pdf` side-effect fix (prerequisite)

Current implementation calls `self.save(self.source_path, backup=False)` — mutating the source file — before invoking LibreOffice. To classify PDF export as a read operation (hkk treats exports as readonly) it must save to a `tempfile.mkdtemp()` copy and convert that.

Additional fixes in the same pass:
- **`Path.rename` → `shutil.move`**: the temp-copy fix introduces cross-device moves (`EXDEV`), which `Path.rename` does not handle.
- **Raise on missing output**: when LibreOffice exits 0 but writes nothing, the function currently returns `{"pdf_path": <nonexistent path>}` — silent success on failure. Must raise, naming the path it looked for.
- **Output-path guard**: `InputGuard.output_path()` is only called inside `save()` and requires `.docx` suffix, so it never touches PDF destinations. `convert_to_pdf`'s `output_path` passes straight through, allowing arbitrary-file-write including via `..`. **A "readonly" server must not gain unguarded filesystem writes.** Extend `InputGuard` with a suffix-parameterised `output_path` and apply to all export tools before classifying them as read-class.

Ship with a test asserting the source file's bytes and mtime are untouched by `convert_to_pdf`.

### B.5 Single-binary distribution

Recommendation: **PyInstaller `--onedir`** per platform, uploaded as GitHub Release artifacts; keep `uvx docx-mcp-server` as the primary documented channel (it is already zero-install for MCP users).

- Why not Nuitka: 10–30× compile time in CI for no user-visible gain here. Why not shiv/pex: still requires a system Python — doesn't deliver hkk's "no runtime" property.
- `--onedir` over `--onefile`: an MCP stdio server is long-running; onefile's per-launch self-extraction adds startup latency and tempdir litter for zero benefit. Ship a tarball/zip with a `docx-mcp` entry binary. **Note: `--onedir` is not a "single binary"** — it ships a directory. hkk's "no runtime required" property survives; "single binary" does not. Use "standalone distribution" in user-facing docs.
- **Presidio/spaCy problem**: they are currently hard deps (pyproject) though `pii.py` imports them lazily inside methods (module top is clean — verified). Move them to an extra: `docx-mcp-server[pii]`, and have `scrub_pii` raise a clear install hint when absent. The default binary excludes them (spaCy + `en_core_web_lg` model add **~560 MB**, not 100 MB — per `pii.py`'s own docstring); optionally publish a second `-full` binary with `--collect-data` for the spaCy model. This is also the right dependency posture for uvx users who never touch PII scrubbing.
- CI matrix: macos-arm64, macos-x86_64, linux-x86_64, linux-aarch64, windows-x86_64. Smoke test each artifact in CI: launch, speak MCP initialize over stdio, `create_document` → `save_document`, assert exit 0 (verify-deploys rule: a green build is not a working binary).
- lxml/mistune bundle cleanly under PyInstaller (compiled wheels; add hidden-import entries as CI smoke tests reveal them).

### B.6 PDF generation without LibreOffice

Mirror hkk's tiered chain, honestly labeled:

1. **LibreOffice when present and allowed** (`allow_external_tools`) — full fidelity. Unchanged.
2. **Pure-Python fallback: ReportLab + bundled Liberation fonts** — new `document/pdf_basic.py`.
   - **ReportLab open-source edition**: BSD-3-Clause, pure Python, actively maintained. ~~fpdf2 was initially preferred for API simplicity but is LGPL-3.0-only — PyInstaller freezes it into a redistributed artifact, which is not dynamic linking; LGPL §4 requires relinking means that a NOTICES file cannot discharge.~~ WeasyPrint rejected: requires native Pango/Cairo — breaks the standalone distribution story.
   - Fonts: Liberation Sans/Mono (SIL OFL 1.1, redistributable with license file) bundled as package data — same set hkk embeds.
   - Fidelity scope (phase 1): paragraphs with bold/italic/size/color runs, headings, page size/margins/orientation from `sectPr`, basic tables (grid, widths, cell text, shading fills), page breaks, lists as indented text. **Not rendered**: floating images/textboxes, columns, footnote layout, fields, equations, tracked-change markup. The tool result carries `{"fidelity": "basic", "renderer": "fpdf2", "unrendered": [...counts...]}` so agents and users are never misled (no silent degradation).
   - `convert_to_pdf` gains `renderer: str = "auto"` (`auto` → LibreOffice else basic; `libreoffice` → hard-require; `basic` → force pure-Python for reproducible output).
- Tests: T3 unit (PDF parses via `pypdf`, page count, extracted text contains paragraph text); T2 oracle: `pdftotext`/pypdf extraction compared against `get_body_text` for a fixture corpus; visual fidelity is explicitly *not* asserted (would be self-graded).

---

## Part C — Document Protection (from GongRzhe)

### C.1 What exists vs what GongRzhe has

docx-mcp writes `w:documentProtection` with SHA-512 salted spin-count hashing, **but the hash is wrong** (see §0 bug note: salt instead of iterator in the spin loop). GongRzhe's restricted-editing and signature features write a `.protection` JSON sidecar next to the file — Word ignores it entirely. **Repair the hash (M0), then port the feature intent, not GongRzhe's implementation.**

### C.2 New capabilities

1. **`get_document_protection`** (read tool): reports `w:documentProtection` state — edit mode, enforcement, hash algorithm/spin count present, plus `w:writeProtection` in settings if present. Currently protection is write-only; agents can't audit it.
2. **Editable ranges — the real OOXML restricted-editing mechanism** (ECMA-376 §17.13 range permissions):
   - `add_editable_range(start_para_id, end_para_id, group="everyone")` → inserts paired `w:permStart w:id=N w:edGrp="everyone"` / `w:permEnd w:id=N` around the range. With `set_document_protection("readOnly", password=…)` this yields Word's "Editing Restrictions: read-only, with exceptions" — everything GongRzhe's sidecar pretended to do, in markup Word actually honors.
   - `list_editable_ranges()`, `remove_editable_range(range_id)`.
   - Validation: unpaired `permStart`/`permEnd` added to `_post_repair_warnings` (same shape as the existing bookmark pairing check).
3. **File password encryption** (`encrypt_document(path, password)` / `decrypt_document(path, password, output_path)`): via `msoffcrypto-tool`, whose `encrypt()` API GongRzhe's `protect_document` demonstrably uses. Operates on *closed files* (it re-writes the whole OPC container as a CFB envelope), so these are file-level tools that refuse to run on the currently-open handle. New optional extra `docx-mcp-server[crypto]`. Verify the encrypt API is non-experimental in the current msoffcrypto release before committing to it; if still experimental, ship decrypt-only and document why.
   - **⚠ Encrypted files are CFB/OLE2, not ZIP.** `base.py:91` opens with `zipfile.ZipFile`, which will raise `BadZipFile`. Add an OLE2 magic-byte check (`D0 CF 11 E0 A1 B1 1A E1`) at `open_document` with a diagnostic naming the actual condition — "this file is encrypted, use `decrypt_document` first" — not a generic corrupt-file error.
4. **`get_signatures`** (read tool): detect `_xmlsignatures/` parts in the package, report count, signer certificate subject/issuer/dates parsed via the `cryptography` package, and whether document bytes have been modified since signing is **not** claimed (digest verification across OPC canonicalization is out of scope — reported as `"verified": null` with an explanatory note, never a false positive).

### C.3 Explicitly not ported (and why)

- **GongRzhe's simulated digital signatures** (content hash + name in a sidecar): placeholder crypto. It produces an artifact that *looks like* a security control and verifies nothing Word or any other tool recognizes. Repo policy forbids shipping placeholder crypto; a real XML-DSig writer (OPC part canonicalization, relationship transforms, certificate handling) is a multi-week project on its own — recorded under Remaining Work with that scoping, not half-shipped.
- **The `.protection` sidecar pattern generally**: state about a document that lives outside the document is lost on copy/email — structurally the wrong design for a document format with native support.

### C.4 Validation integration

- `w:documentProtection` and range permissions live in settings.xml / document.xml, already dirty-tracked via `_mark`. The settings.xml element-order check (CT_Settings is also a sequence) joins the same `ooxml_order` warning pass from Part A — one mechanism, three property containers.
- Round-trip test: protect + permStart/permEnd doc through LibreOffice re-save (T2), assert protection and ranges survive; T1 fixture: a Word-authored document with editing restrictions, assert `get_document_protection` and `list_editable_ranges` read it correctly.

---

## Attribution & License Compliance

Create `THIRD_PARTY_NOTICES.md` at repo root (linked from README footer):

```
## Office-Word-MCP-Server
Copyright (c) 2025 GongRzhe — MIT License
https://github.com/GongRzhe/Office-Word-MCP-Server (archived Dec 2025)
Table-formatting tool design (banding, header styling, cell padding, auto-fit,
table width) and the msoffcrypto encryption approach are derived from this
project. Code was reimplemented against docx-mcp's lxml XML-direct
architecture; where fragments were adapted, the MIT license text below applies.

## docx-mcp (Rust)
Copyright (c) hongkongkiwi — MIT License
https://github.com/hongkongkiwi/docx-mcp
Deployment architecture concepts (readonly mode, tool allow/blocklists,
resource limits, tiered PDF fallback with embedded fonts) follow this
project's design.

[full MIT license texts]

## Bundled fonts
Liberation Sans / Liberation Mono — SIL Open Font License 1.1 (license text bundled
alongside the font files in docx_mcp/fonts/).
```

MIT obligations trigger on copied/substantial portions; most of this work is reimplementation-from-design against a different XML layer, but include the notices regardless — attribution is cheap, and some conversion tables (e.g. dxa/pct math, readonly-command classification) will be closely derived. Add ReportLab (BSD-3-Clause) and msoffcrypto-tool (MIT) entries when those deps land. Verify every license text against the repos' LICENSE files at implementation time, not from this plan.

---

## Sequencing & Milestones

TDD throughout: each numbered item is a RED commit (failing tests) followed by a GREEN commit.

**M0 — Prerequisites (bug fixes + safety, ship first)**
1. **Password hash fix**: `protection.py` spin loop — replace `salt` with 4-byte LE iteration counter per MS-OFFCRYPTO §2.4.2.4. Add test that verifies against an independent implementation or empirically in Word.
2. `convert_to_pdf` no longer mutates the source file, plus `shutil.move`, raise-on-missing, and output-path guard (§B.4)
3. gridSpan-aware column-count validation fix at **both** sites (`base.py` and `validation.py`), with `hMerge` tolerance (§A.4.1)
4. OLE2 magic-byte detection at `open_document` for encrypted files (§C.2.3)

**M1 — Part A tables** (largest user-visible win, no architectural risk)
5. `ooxml_order.py` helper
6. Migrate existing `SubElement` writers in `tcPr`/`tblPr`/`settings.xml` to ordered insertion
7. The 9 new table tools, one RED/GREEN pair each
8. Out-of-order warning pass (AFTER migration — shipping detector before fix reads as regression)
9. Schema-validation + round-trip test jobs

**M2 — Part B config & filtering** (enables safe deployment of everything else)
10. `config.py` + argparse CLI (config resolves before `import server`, not before `mcp.run()`)
11. `doc_tool` classification sweep (split by module, not one mega-diff) + registration/call-time filtering + `--no-network` enforcement
12. Resource limits (size gate — **new behaviour, not refactor** — zip-bomb guard, handle cap, path confinement)

**M3 — Part C protection** (M0 hash fix is prerequisite)
13. `get_document_protection`, editable ranges + validation
14. encrypt/decrypt via msoffcrypto (behind `[crypto]` extra), with OLE2 detection at open
15. `get_signatures` inspection

**M4 — Part B distribution — ⚠ ON HOLD** (pending fpdf2 license + macOS signing decisions)
16. Presidio/spaCy → `[pii]` extra (**ship this independently** — stands alone, cuts ~560 MB, no M4 risk)
17. ReportLab basic PDF renderer + fonts (~~fpdf2~~ LGPL-3.0-only conflicts with PyInstaller redistribution)
18. PyInstaller CI matrix + artifact smoke tests + **macOS notarisation** (unsigned bundles get Gatekeeper-blocked; needs paid Apple Developer account, `codesign`, `notarytool`, hardened-runtime entitlements)
19. `THIRD_PARTY_NOTICES.md` + README updates

Dependency notes: M2 blocks M4 (binary needs the CLI flags); M0.1 blocks M2's readonly classification of `convert_to_pdf`; M0 hash fix blocks M3; nothing blocks M1. **Ship M4.16 (`[pii]` extra) with M2 — it stands alone.**

## Remaining Work (recorded, deliberately not in scope)

- Tracked table-property changes (`w:tblPrChange`/`w:tcPrChange`) — extends docx-mcp's tracked-changes brand to formatting; medium effort.
- Real XML-DSig signing/verification — multi-week; requires OPC canonicalization and relationship-transform support; revisit if user demand materializes.
- Style-based banding authoring (creating `w:tblStylePr` band definitions in styles.xml) — `set_table_banding(method="style")` phase-1 only *activates* existing style bands via `tblLook`.

## Risks & Open Questions

- **msoffcrypto encryption maturity** — verify the encrypt API's status in the current release before M3.14; fall back to decrypt-only if experimental.
- **ECMA-376 schema redistribution** — confirm terms before committing `wml.xsd` to the repo; CI-download + local-skip otherwise (repo coverage-gate rule: committed tests must pass from committed bytes, so the schema job runs as a separate skip-when-absent T2 job, not in the coverage gate).
- **PyInstaller + lxml on linux-aarch64** — least-tested combination in the matrix; smoke test may surface hidden imports.
- **Tool-count context cost** — 235 tools post-M1 strengthens the case for shipping allowlist presets (e.g. `--profile review` exposing ~25 tools); consider as a follow-up UX item.
- **~~fpdf2~~ ReportLab licence under freezing** — fpdf2 (LGPL-3.0-only) was rejected because PyInstaller redistribution is not dynamic linking. ReportLab open-source (BSD-3-Clause) replaces it. Verify ReportLab's licence text at implementation time.
- **macOS notarisation** — unsigned PyInstaller bundles carry the quarantine attribute; Gatekeeper blocks them. Needs: paid Apple Developer account, Developer ID certificate, `codesign`, `notarytool` submission, stapling, hardened-runtime entitlements. This is routinely the single largest line item in "ship a binary" projects and was absent from the original plan.
- **macOS x86_64 runner** — GitHub's Intel macOS runners (`macos-13`) are at end of life. Apple-silicon runners cannot produce x86_64 PyInstaller output reliably. Either drop the target, accept Rosetta-based builds, or budget a self-hosted Intel Mac.
- **Coverage gate** — CI enforces `--cov-fail-under=88`. A config module, decorator sweep, PDF renderer, and packaging glue will all drag that number. Budget for holding the gate green.
- **Transitional vs strict namespace** — `TCMAR_ORDER` is hard-coded to transitional form (`top/left/bottom/right`). Strict-schema documents use `top/start/bottom/end`. Key the order table off the document's actual namespace rather than assuming transitional.
- **Table indexing includes nested tables** — `_get_table` uses `doc.iter(f"{W}tbl")` (pre-order traversal). All 9 new tools inherit this, which will surprise users formatting documents with nested tables. Document in tool docstrings.
- **Competitor claims are unverified from this workspace** — neither GongRzhe nor hongkongkiwi repos are cloned under `~/src`. Re-clone before M1 and re-verify licence texts from the repos' own LICENSE files.
