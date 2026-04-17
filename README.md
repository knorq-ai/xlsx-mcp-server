# xlsx-mcp-server

[![CI](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/knorq-ai/xlsx-mcp-server/actions/workflows/ci.yml)

A local [MCP](https://modelcontextprotocol.io/) server for reading and editing Excel (.xlsx) files. Works with Claude Code, Cursor, and any MCP-compatible client.

**37 tools** for cell data, formatting, formulas, sheet management, row/column operations, data validation, named ranges, and cell merging — all running locally via stdio with no file uploads.

## Features

| Category | Tools |
|---|---|
| **Read** | `get_workbook_info`, `read_sheet`, `read_cell`, `search_cells`, `get_sheet_properties`, `list_named_ranges`, `list_data_validations`, `list_images` |
| **Write** | `write_cell`, `write_cells`, `write_row`, `write_rows`, `clear_cells`, `create_workbook` |
| **Format** | `format_cells`, `format_cells_bulk` |
| **Rows & columns** | `set_column_width`, `set_column_widths`, `set_row_height`, `set_row_heights`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns` |
| **Sheet ops** | `add_sheet`, `rename_sheet`, `delete_sheet`, `copy_sheet` |
| **View** | `set_freeze_panes`, `set_auto_filter`, `remove_auto_filter` |
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

`read_cell` returns both the formula and the cached result.

## Quick start

### Option 1: Install from npm

```bash
npm install -g @llamadrive/xlsx-mcp-server
```

Then add to your MCP config (see [Configuration](#configuration) below).

### Option 2: Use npx (no install)

Just add the config — `npx` downloads and runs it automatically:

```json
{
  "mcpServers": {
    "xlsx-editor": {
      "command": "npx",
      "args": ["-y", "@llamadrive/xlsx-mcp-server"]
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
      "args": ["-y", "@llamadrive/xlsx-mcp-server"]
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
      "args": ["-y", "@llamadrive/xlsx-mcp-server"]
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
npm install -g @llamadrive/xlsx-mcp-server
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

**`read_sheet`** — Read cell data from a sheet with optional range.
```
file_path, sheet, range?
```

**`read_cell`** — Single cell's value, formula, type, and formatting.
```
file_path, sheet, cell
```

**`search_cells`** — Search for text or numbers across cells.
```
file_path, query, sheet?, case_sensitive?
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

**`write_cell`** — Set a cell's value or formula. Prefix with `=` for formulas.
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

**`clear_cells`** — Clear values in a range (keeps formatting).
```
file_path, sheet, range
```

**`create_workbook`** — Create a new empty .xlsx workbook.
```
file_path, sheet_name?
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

**`insert_rows`** — Insert empty rows at a position.
```
file_path, sheet, row, count
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

### View settings

**`set_freeze_panes`** — Freeze rows and/or columns. Pass 0 to unfreeze.
```
file_path, sheet, row, column
```

**`set_auto_filter`** — Enable auto filter on a range.
```
file_path, sheet, range
```

**`remove_auto_filter`** — Remove auto filter from a sheet.
```
file_path, sheet
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

### Not supported (use Python/openpyxl/xlwings instead)

| Feature | Detail |
|---------|--------|
| **Formula recalculation** | Cached results are read, but formulas are NOT recalculated when values change. Open the file in Excel to recalculate. |
| **Charts** | Cannot read, create, or modify charts. Existing charts are preserved on save. |
| **Pivot tables** | Cannot read or create pivot tables |
| **Conditional formatting** | Cannot read or create conditional formatting rules |
| **VBA/macros** | Macro-enabled workbooks (.xlsm) are not supported |
| **Formula ref auto-update** | Inserting/deleting rows or columns does NOT shift cell references in existing formulas (e.g. `=SUM(A1:A10)` stays unchanged after row insert) |

### Other limitations

- **copy_sheet is partial** — Copies cell values, styles, column widths, row heights, and merged cells. Does not copy data validation, conditional formatting, or view settings
- **Range size limit** — Write, format, and validation tools reject ranges exceeding 100,000 cells
- **File size limit** — Files larger than 100 MB cannot be opened

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
