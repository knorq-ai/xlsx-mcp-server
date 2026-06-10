# xlsx-mcp-server

[![CI](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml)

A local [MCP](https://modelcontextprotocol.io/) server for reading and editing Excel (.xlsx) files. Works with Claude Code, Cursor, and any MCP-compatible client.

**47 tools** for cell data, formatting, formulas, range copy/sort/find-replace, sheet management, row/column operations, data validation, named ranges, cell merging, notes, sheet protection, and page setup — all running locally via stdio with no file uploads.

## Features

| Category | Tools |
|---|---|
| **Read** | `get_workbook_info`, `read_sheet`, `read_cell`, `search_cells`, `get_sheet_properties`, `list_named_ranges`, `list_data_validations`, `list_images` |
| **Write** | `write_cell`, `write_cells`, `write_row`, `write_rows`, `clear_cells`, `set_cell_note`, `create_workbook` |
| **Range ops** | `copy_range`, `find_replace`, `sort_range` |
| **Format** | `format_cells`, `format_cells_bulk` |
| **Rows & columns** | `set_column_width`, `set_column_widths`, `set_row_height`, `set_row_heights`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, `set_row_visibility`, `set_column_visibility` |
| **Sheet ops** | `add_sheet`, `rename_sheet`, `delete_sheet`, `copy_sheet`, `set_sheet_properties`, `protect_sheet`, `unprotect_sheet` |
| **View & layout** | `set_freeze_panes`, `set_auto_filter`, `remove_auto_filter`, `set_page_setup` |
| **Validation** | `add_data_validation`, `remove_data_validation` |
| **Structure** | `add_named_range`, `delete_named_range`, `merge_cells`, `unmerge_cells` |

### Bulk operations

The writing, formatting, and row/column tools have bulk variants (`write_cells`, `write_rows`, `format_cells_bulk`, `set_column_widths`, `set_row_heights`) that process multiple targets in a single file read/write cycle. Use these instead of calling the single-target versions in a loop.

### Formula support

Write formulas by prefixing the value with `=`:

```
write_cell  →  value: "=SUM(A1:A10)"
write_cells →  cells: [{cell: "B1", value: "=A1*2"}, {cell: "B2", value: "=VLOOKUP(...)"}]
```

`read_cell` returns both the formula and the cached result. To write a literal string that starts with `=`, prefix it with a single quote (`'=text` writes the string `=text` — Excel's escape rule). Formulas are not recalculated on edit, but recalc-on-open is enabled on every save, so Excel recomputes everything when the file is opened.

### Date and hyperlink values

All write tools accept object values for true Excel dates and hyperlinks:

```
write_cell →  value: {date: "2024-01-15"}                            // true Excel date cell
write_cell →  value: {hyperlink: "https://example.com", text: "Docs"} // hyperlink with display text
```

### read_sheet JSON format

`read_sheet` returns cell data as address-keyed maps inside a `<json>...</json>` block — an absent address means an empty cell:

```json
{
  "sheetName": "Sheet1",
  "range": "A1:C3",
  "cells": {"A1": "Product", "B1": "Price", "A2": "Widget", "B2": 9.99},
  "formulas": {"C2": {"f": "B2*2", "v": 19.98}},
  "dates": {"A3": "2024-01-15T00:00:00.000Z"},
  "mergedCells": ["A1:B1"]
}
```

Additional maps (`errors`, `hyperlinks`, `numFmts`, `notes`, and — with `include_styles: true` — `styles` in the `format_cells` vocabulary) appear only when present. Output is capped at 5,000 cells; truncated reads set `truncated: true`, so read large sheets in chunks via `range`.

## Quick start

### Option 1: Install from npm

```bash
npm install -g @knorq/xlsx-mcp-server
```

Then add to your MCP config (see [Configuration](#configuration) below).

### Option 2: Use npx (no install)

Just add the config — `npx` downloads and runs it automatically:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@knorq/xlsx-mcp-server"]
    }
  }
}
```

### Option 3: Build from source

```bash
git clone https://github.com/knorq-ai/xlsx-mcp-server.git
cd xlsx-mcp-server
npm install
npm run build
npm link        # makes `xlsx-mcp-server` available globally
```

## Configuration

### Claude Code

Add to your project's `.mcp.json` (per-project) or `~/.claude/settings.json` (global):

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@knorq/xlsx-mcp-server"]
    }
  }
}
```

### Cursor

Add to your MCP server configuration in Cursor settings:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@knorq/xlsx-mcp-server"]
    }
  }
}
```

### Using a local build (without npm)

If you built from source and ran `npm link`:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "xlsx-mcp-server"
    }
  }
}
```

Or reference the built file directly:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "node",
      "args": ["/absolute/path/to/xlsx-mcp-server/dist/index.js"]
    }
  }
}
```

## Distributing to others

### Via npm (recommended)

```bash
npm publish
```

Recipients install with:

```bash
npm install -g @knorq/xlsx-mcp-server
```

Or skip the install entirely — just share the `.mcp.json` config with the `npx` setup above and it works out of the box.

### Via zip / git

Share the repository. Recipients run:

```bash
git clone https://github.com/knorq-ai/xlsx-mcp-server.git
cd xlsx-mcp-server
npm install
npm run build
npm link
```

Then add the config above.

## Tool reference

### Reading

**`get_workbook_info`** — Sheet list, named range count, file properties.
```
file_path
```

**`read_sheet`** — Read cell data from a sheet as address-keyed JSON maps (see [read_sheet JSON format](#read_sheet-json-format)). Output capped at 5,000 cells.
```
file_path, sheet, range?, include_styles?
```

**`read_cell`** — Single cell's value, formula, type, and formatting. Formatting is returned in the same vocabulary `format_cells` accepts (`bold`, `fillColor`, …), so a read style can be passed straight back.
```
file_path, sheet, cell
```

**`search_cells`** — Search for text or numbers across cells.
```
file_path, query, sheet?, case_sensitive?, max_results?
```

**`get_sheet_properties`** — Sheet state, dimensions, freeze panes, auto filter, tab color.
```
file_path, sheet
```

**`list_named_ranges`** — List all named ranges with names and references.
```
file_path
```

**`list_data_validations`** — List data validation rules on a sheet.
```
file_path, sheet
```

**`list_images`** — List embedded images with names, extensions, and dimensions.
```
file_path, sheet
```

### Cell writing

**`write_cell`** — Set a cell's value or formula. Prefix with `=` for formulas; use `{date: "ISO"}` for dates and `{hyperlink, text}` for links.
```
file_path, sheet, cell, value
```

**`write_cells`** — Set multiple cells at once.
```
file_path, sheet, cells (array of {cell, value})
```

**`write_row`** — Write a row of values starting from a position.
```
file_path, sheet, row, values, start_column?
```

**`write_rows`** — Write multiple rows of data at once.
```
file_path, sheet, start_row, rows (2D array), start_column?
```

**`clear_cells`** — Clear cell values and/or formatting in a range. `mode: "values"` (default) keeps formatting, `"formats"` keeps values, `"all"` clears both.
```
file_path, sheet, range, mode?
```

**`set_cell_note`** — Set or remove a cell note (comment). Pass `null` to remove.
```
file_path, sheet, cell, note
```

**`create_workbook`** — Create a new empty .xlsx workbook.
```
file_path, sheet_name?
```

### Range operations

**`copy_range`** — Copy a range (values, formulas, formatting, merges) to another location, optionally on a different sheet. Relative formula references shift to the destination; `$`-anchored references stay fixed.
```
file_path, sheet, source_range, destination, dest_sheet?
```

**`find_replace`** — Find and replace text across plain string cells. Formulas, numbers, rich text, and hyperlinks are not modified. Searches all sheets unless one is specified.
```
file_path, query, replacement, sheet?, case_sensitive?, match_entire_cell?
```

**`sort_range`** — Sort the rows of a range by a key column. Values, formulas, and formatting move together; relative formula references are re-anchored. Fails if the range intersects merged cells.
```
file_path, sheet, range, key_column, ascending?, has_header?
```

### Formatting

**`format_cells`** — Apply formatting to a cell range: font (bold, italic, underline, strikethrough, name, size, color), fill (color, pattern), borders (style, color, sides), alignment (horizontal, vertical, wrap, rotation), number format.
```
file_path, sheet, range, format
```

**`format_cells_bulk`** — Apply different formatting to multiple ranges at once. Single file read/write cycle.
```
file_path, sheet, groups (array of {range, format})
```

### Rows and columns

**`set_column_width`** — Set the width of a column (in characters).
```
file_path, sheet, column, width
```

**`set_column_widths`** — Set widths for multiple columns at once.
```
file_path, sheet, columns (array of {column, width})
```

**`set_row_height`** — Set the height of a row (in points).
```
file_path, sheet, row, height
```

**`set_row_heights`** — Set heights for multiple rows at once.
```
file_path, sheet, rows (array of {row, height})
```

**`insert_rows`** — Insert empty rows at a position. Set `inherit_style: true` to copy formatting (and row height) from the row above.
```
file_path, sheet, row, count, inherit_style?
```

**`delete_rows`** — Delete rows at a position.
```
file_path, sheet, row, count
```

**`insert_columns`** — Insert empty columns at a position.
```
file_path, sheet, column, count
```

**`delete_columns`** — Delete columns at a position.
```
file_path, sheet, column, count
```

**`set_row_visibility`** — Hide or unhide a range of rows.
```
file_path, sheet, start_row, end_row, hidden
```

**`set_column_visibility`** — Hide or unhide a range of columns.
```
file_path, sheet, start_column, end_column, hidden
```

### Sheet operations

**`add_sheet`** — Add a new empty sheet.
```
file_path, name
```

**`rename_sheet`** — Rename an existing sheet.
```
file_path, sheet, new_name
```

**`delete_sheet`** — Delete a sheet from the workbook.
```
file_path, sheet
```

**`copy_sheet`** — Copy a sheet within the workbook.
```
file_path, source_sheet, new_name
```

**`set_sheet_properties`** — Set sheet visibility (`visible` / `hidden` / `veryHidden`) and/or tab color. A workbook must keep at least one visible sheet.
```
file_path, sheet, state?, tab_color?
```

**`protect_sheet`** — Protect a sheet against editing in Excel, optionally with a password. This is Excel UI-level protection, not encryption — this server and other tools can still modify the file.
```
file_path, sheet, password?
```

**`unprotect_sheet`** — Remove sheet protection.
```
file_path, sheet
```

### View & layout

**`set_freeze_panes`** — Freeze rows and/or columns. Pass 0 to both to unfreeze.
```
file_path, sheet, freeze_rows, freeze_columns
```

**`set_auto_filter`** — Enable auto filter on a range.
```
file_path, sheet, range
```

**`remove_auto_filter`** — Remove auto filter from a sheet.
```
file_path, sheet
```

**`set_page_setup`** — Configure print/PDF layout: orientation, print area, fit-to-page, paper size.
```
file_path, sheet, orientation?, print_area?, fit_to_width?, fit_to_height?, paper_size?
```

### Data validation

**`add_data_validation`** — Add a validation rule (list, whole, decimal, date, textLength, custom) with operator, messages, and prompts.
```
file_path, sheet, range, type, formulae, operator?, allow_blank?, show_error_message?, error_title?, error?, show_input_message?, prompt_title?, prompt?
```

**`remove_data_validation`** — Remove validation rules from a range.
```
file_path, sheet, range
```

### Named ranges

**`add_named_range`** — Add a named range (workbook-scoped or sheet-scoped).
```
file_path, name, range, sheet?
```

**`delete_named_range`** — Delete a named range.
```
file_path, name
```

### Cell merging

**`merge_cells`** — Merge a range of cells.
```
file_path, sheet, range
```

**`unmerge_cells`** — Unmerge a previously merged range.
```
file_path, sheet, range
```

## Known limitations

### Destructive or rejected

| Feature | Behavior |
|---------|----------|
| **Charts, pivot tables, slicers** | **Destroyed by any write operation.** Reading is safe, but any tool that saves the file removes them from the workbook. Do not edit workbooks whose charts or pivots must survive. |
| **VBA macros (.xlsm/.xltm)** | **Read-only.** Writes are rejected because saving would silently destroy the VBA project. |

### Not supported

| Feature | Detail |
|---------|--------|
| **Formula recalculation** | Formulas are NOT evaluated on edit. Cached results are read; formula cells without a cached result read as `(not calculated)`. Recalc-on-open is enabled on every save, so Excel recalculates when the file is opened. |
| **Conditional formatting** | Existing rules are preserved on save, but there are no tools to read or edit them. |
| **Formula ref auto-update** | Inserting/deleting rows or columns does NOT shift cell references inside formula text (e.g. `=SUM(A1:A10)` stays unchanged after a row insert). Merged cells and data validations DO shift correctly. Do structural changes before writing formulas. |

### Other limitations

- **copy_sheet is partial** — Copies cell values, styles, column widths, row heights, and merged cells. Does not copy data validation, conditional formatting, or view settings
- **Range size limit** — Write, format, and validation tools reject ranges exceeding 100,000 cells (lower it via `XLSX_MAX_CELLS_PER_CALL`)
- **File size limit** — Files larger than 100 MB cannot be opened

## Safety and reliability

- **Atomic saves** — files are written to a temp file and renamed, so a crash mid-save never truncates the workbook.
- **Cross-process write lock** — a `<file>.mcplock` advisory lock (with stale-owner detection) serializes writes across multiple server instances, on top of in-process serialization.

### Environment variables

Read once at server start — they cannot be overridden through tool parameters.

| Variable | Effect |
|----------|--------|
| `XLSX_MAX_CELLS_PER_CALL` | Hard cap on cells touched by a single write/format call. Defaults to 100,000; deployments can lower it. |
| `XLSX_TEMPLATE_MODE=1` + `XLSX_TEMPLATE_RANGES=Sheet1!A1:D10,Sheet1!F2:F100` | Template mode: all writes/formats/clears must fall inside the declared ranges; anything outside is rejected with `OUTSIDE_TEMPLATE_RANGE`. Also blocks structural operations (insert/delete rows/columns, delete/rename sheet) and `find_replace`. |
| `XLSX_BACKUP_ON_WRITE=1` | Copies the workbook to `<file>.bak` before each save. |

## Why MCP tools instead of raw Python?

AI agents can manipulate Excel via raw Python (openpyxl), but MCP tools are significantly more token-efficient:

| Metric | MCP tools | Raw Python |
|--------|-----------|------------|
| Output tokens per operation | **60–85% less** | Baseline (agent must generate full code) |
| Cost per operation | **50–80% less** | Baseline |
| Break-even | **2 operations** | — |
| Debug iterations | None (validated inputs) | ~1.5 retries/task on average |

The savings come primarily from **eliminating code generation** — output tokens cost 5× more than input tokens. MCP tool calls are small structured parameters (~30–50 tokens), while equivalent Python code requires ~80–200 output tokens per operation (imports, style objects, iteration, save).

Formatting operations see the largest savings (~75%) because openpyxl's styling API (`PatternFill`, `Border`, `Side`, `Font`) is particularly verbose. Simple cell read/write sees smaller but still meaningful savings (~60%).

See [docs/token-efficiency-analysis.md](docs/token-efficiency-analysis.md) for detailed scenario breakdowns.

## Requirements

- Node.js 18+

## License

MIT
