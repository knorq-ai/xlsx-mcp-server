# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.0.0] — 2026-06-10

### Fixed — silent file corruption
- **Saves are now atomic** (temp file + rename). A crash or error mid-save no longer truncates the workbook to 0 bytes.
- **`insert_rows` / `delete_rows` / `insert_columns` / `delete_columns` no longer corrupt merged cells or data validations.** ExcelJS shifts cell content but not the serialized merge ranges, and never shifts data validations; both are now preserved at their shifted positions (including merges shrinking on partial deletes).
- **Structural ops and master-cell overwrites no longer break shared-formula groups.** Groups are materialized into per-cell translated formulas before splices and before overwriting a master, which previously made the next save throw.
- **`format_cells` no longer restyles unrelated cells.** ExcelJS shares style objects between same-styled cells after a file load; partial format updates leaked into them.
- **Writes to merged child cells are rejected** with guidance instead of silently overwriting the merge master's value. `clear_cells` skips merged children.
- `copy_sheet` deep-clones styles and no longer destroys per-cell formatting inside merged ranges.
- `fillPattern: "solid"` without `fillColor` preserves the existing fill color instead of overwriting it with white.
- Sheet names are validated (≤31 chars, no `* ? : \ / [ ]`); previously ExcelJS silently accepted 32+ chars, producing files Excel rejects.
- Cell addresses are bounded to the Excel grid (row ≤ 1,048,576, column ≤ XFD); previously out-of-grid rows were written silently, producing files Excel rejects.
- Writing to `.xlsm` / `.xltm` is rejected — saving would silently destroy all VBA macros.
- Shared-formula slave cells with a missing or non-formula master no longer crash reads, and the master address is never reported as the cell's formula (`sharedGroupMaster` field instead).
- Formula error results normalize to the error string (`#DIV/0!`) instead of `[object Object]`.
- Formula cells without cached results are no longer dropped from reads; they report `(not calculated)` and recalc-on-open is enabled on every save (`fullCalcOnLoad`).

### Changed (BREAKING)
- **`read_sheet` JSON payload redesigned** to address-keyed maps (`cells`, `formulas` `{f, v}`, `dates`, `errors`, `hyperlinks`, `numFmts`, `notes`, `styles`, `mergedCells`) — roughly 4× fewer tokens than the old per-cell objects. An absent address means an empty cell. The `compact` parameter is gone (the format is inherently compact), and the per-row text dump is replaced by a short summary (the JSON block is the single source of cell data).
- **`read_cell` returns formatting in the `format_cells` vocabulary** (`bold`, `fillColor`, …) instead of raw ExcelJS style objects, so a read style can be passed straight back to `format_cells`.
- `set_freeze_panes` parameters renamed `row`/`column` → `freeze_rows`/`freeze_columns`.
- `read_sheet` output is capped at 5,000 cells (truncation notice + chunked reads via `range`); `search_cells` gains `max_results` (default 100, max 1,000) and rejects empty queries.
- Server capability/limitations text moved to the MCP `instructions` field (it was previously in a field clients never read) and rewritten to be accurate: charts/pivot tables/slicers are **destroyed by any write**, not "unsupported".
- All tools are registered with MCP `ToolAnnotations` (`readOnlyHint`, `destructiveHint`, `idempotentHint`).
- `exceljs` is pinned to exactly 4.4.0 (workarounds depend on verified internals).

### Added
- **`copy_range`** — copy values, formulas, formatting, and merges within or across sheets; relative formula references shift to the destination.
- **`find_replace`** — bulk text replacement across plain string cells (formulas and numbers untouched).
- **`sort_range`** — row-wise sort by key column; styles and formulas move with their rows, relative references re-anchored.
- **Date and hyperlink write values** — `{date: "2024-01-15"}` writes a true Excel date; `{hyperlink, text}` writes a link (all write tools).
- **`set_cell_note`** — set/remove cell comments; reads now include notes.
- **`set_sheet_properties`** (visibility, tab color), **`set_row_visibility`** / **`set_column_visibility`**, **`protect_sheet`** / **`unprotect_sheet`**, **`set_page_setup`** (orientation, print area, fit-to-page, paper size).
- `clear_cells` gains `mode: values | formats | all`; `insert_rows` gains `inherit_style`.
- `read_sheet` gains `include_styles` (per-cell formatting in the `format_cells` shape).
- **Cross-process write lock** (`<file>.mcplock`, stale-owner detection) on top of the in-process serialization, so multiple MCP server instances can safely edit the same workbook.
- `XLSX_BACKUP_ON_WRITE=1` copies the workbook to `<file>.bak` before each save.
- Template mode now also blocks structural ops (insert/delete rows/columns, delete/rename sheet) and `find_replace`; `merge_cells`/`unmerge_cells` validate ranges and respect cell caps.
- Literal-`=` escape: `'=text` writes the string `=text` (Excel's escape rule).
- "Sheet not found" errors list the available sheet names.

## [2.0.0] — 2026-04-17

### Changed (BREAKING)
- Renamed package from `xlsx-mcp-server` to `@knorq/xlsx-mcp-server`. Update your `.mcp.json` / install commands to the scoped name.
- Pinned `engines.node` to `>=18.0.0`.

### Added
- `XLSX_MAX_CELLS_PER_CALL` environment variable: bounds bulk write/format operations. Defaults to the existing 100,000-cell range cap; deployments can lower it. Reject is enforced server-side, before the LLM-supplied range reaches the engine.
- Template mode (`XLSX_TEMPLATE_MODE=1` + `XLSX_TEMPLATE_RANGES=Sheet1!A1:D10,Sheet1!F2:F100`): when enabled, all writes/formats/clears must fall inside one of the declared ranges. Writes outside are rejected with `OUTSIDE_TEMPLATE_RANGE`. Enforces structural integrity at the server, not in the prompt.
- GitHub Actions workflow that publishes to npm with `--provenance --access public` on tag push, signed via OIDC.

## [1.1.0] and earlier

See git history.
