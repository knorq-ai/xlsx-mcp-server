#!/usr/bin/env node

/**
 * XLSX MCP Server — Local MCP server for reading, writing, formatting,
 * and managing Excel workbooks.
 *
 * Transport: stdio (runs locally, no file uploads)
 * Usage with Claude Code:  Add to ~/.claude/settings.json under mcpServers
 * Usage with Cursor:       Add to MCP server configuration
 */

import { createRequire } from "node:module";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  getWorkbookInfo,
  readSheet,
  readCell,
  searchCells,
  getSheetProperties,
  listWorkbookNamedRanges,
  listDataValidations,
  listImages,
  createWorkbook,
  writeCell,
  writeCells,
  writeRow,
  writeRows,
  clearCells,
  formatCells,
  formatCellsBulk,
  addSheet,
  renameSheet,
  deleteSheet,
  copySheet,
  setColumnWidth,
  setColumnWidths,
  setRowHeight,
  setRowHeights,
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  setFreeze,
  setSheetAutoFilter,
  removeSheetAutoFilter,
  addDataValidation,
  removeDataValidation,
  addNamedRange,
  deleteNamedRange,
  mergeCells,
  unmergeCells,
  EngineError,
} from "./xlsx-engine.js";
import {
  parseCellAddress,
  parseRange,
  columnLetterToNumber,
  type CellRange,
} from "./engine/cells.js";
import {
  loadSafetyConfig,
  assertCellCount,
  assertWithinTemplate,
} from "./engine/safety.js";

const require = createRequire(import.meta.url);
const { version: VERSION } = require("../package.json") as { version: string };

function formatError(e: unknown): string {
  if (e instanceof EngineError) {
    return `[${e.code}] ${e.message}`;
  }
  if (e instanceof Error) {
    return `[INTERNAL_ERROR] ${e.message}`;
  }
  return `[INTERNAL_ERROR] ${String(e)}`;
}

// F-002 enforcement: env-driven cell-count cap and template-mode whitelist.
// Loaded once at startup so the LLM cannot disable them via tool parameters.
const safetyConfig = loadSafetyConfig();

function cellAddrToRange(addr: string): CellRange {
  const { col, row } = parseCellAddress(addr);
  return { startCol: col, startRow: row, endCol: col, endRow: row };
}

function rowRange(row: number, startCol: number, endCol: number): CellRange {
  return { startCol, endCol, startRow: row, endRow: row };
}

// Shared schemas
const filePathSchema = z.string().describe("Absolute path to the .xlsx file");
const sheetSchema = z.union([z.string(), z.number().int().min(1)]).describe("Sheet name or 1-based index");
const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);
const cellAddressSchema = z.string().regex(/^[A-Za-z]+\d+$/, "Invalid cell address (expected A1 format)").describe("Cell address (e.g. 'A1', 'B5')");
const columnSchema = z.string().regex(/^[A-Za-z]+$/, "Invalid column letter").describe("Column letter (e.g. 'A', 'BC')");
const hexColorSchema = z.string().regex(/^[0-9A-Fa-f]{6}$/, "Invalid hex color (expected 6-char hex, e.g. 'FF0000')");
const rowSchema = z.number().int().min(1).describe("Row number (1-based)");
const countSchema = z.number().int().min(1).describe("Number of items");

// ---------------------------------------------------------------------------
// Server setup
// ---------------------------------------------------------------------------

const server = new McpServer({
  name: "xlsx-editor",
  version: VERSION,
  description: [
    "Read, write, format, and manage Excel (.xlsx) workbooks.",
    "",
    "Supported: cell read/write, formulas, formatting (font/fill/border/alignment/numFmt),",
    "merged cells, sheets, named ranges, data validation, row/column ops, freeze panes, auto filter.",
    "",
    "NOT supported (use Python/openpyxl/xlwings instead):",
    "- Formula recalculation — cached results are read, but formulas are NOT recalculated on edit",
    "- Charts — cannot read, create, or modify charts",
    "- Pivot tables — not supported",
    "- Conditional formatting — cannot read or create CF rules",
    "- VBA/macros — preserved but cannot be read or executed",
    "- Formula ref auto-update — inserting/deleting rows does NOT shift formula references",
  ].join("\n"),
});

// =========================================================================
// Reading tools (8)
// =========================================================================

server.tool(
  "get_workbook_info",
  "Get metadata and structure overview of an XLSX file — sheet list, named range count, and file properties.",
  {
    file_path: filePathSchema,
  },
  async ({ file_path }) => {
    try {
      const result = await getWorkbookInfo(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "read_sheet",
  "Read cell data from a sheet (values, formulas, types). Optionally specify a range like 'A1:C10'. Use compact=true to omit empty cells and merged-cell children for token-efficient output.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().optional().describe("Cell range to read (e.g. 'A1:C10'). Omit to read all data."),
    compact: z.boolean().optional().default(false).describe("Omit empty cells and merged-cell children. Reduces output for sheets with many merged cells."),
  },
  async ({ file_path, sheet, range, compact }) => {
    try {
      const result = await readSheet(file_path, sheet, range, compact);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "read_cell",
  "Read a single cell's value, formula, type, and formatting.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    cell: cellAddressSchema,
  },
  async ({ file_path, sheet, cell }) => {
    try {
      const result = await readCell(file_path, sheet, cell);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "search_cells",
  "Search for text or numbers in cells. Searches all sheets by default, or specify a sheet.",
  {
    file_path: filePathSchema,
    query: z.string().describe("Text to search for"),
    sheet: sheetSchema.optional().describe("Sheet to search in (omit for all sheets)"),
    case_sensitive: z.boolean().optional().default(false).describe("Case-sensitive search. Default false."),
  },
  async ({ file_path, query, sheet, case_sensitive }) => {
    try {
      const result = await searchCells(file_path, query, sheet, case_sensitive);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "list_named_ranges",
  "List all named ranges in the workbook.",
  {
    file_path: filePathSchema,
  },
  async ({ file_path }) => {
    try {
      const result = await listWorkbookNamedRanges(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "list_data_validations",
  "List data validation rules on a sheet.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
  },
  async ({ file_path, sheet }) => {
    try {
      const result = await listDataValidations(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "list_images",
  "List images embedded in a sheet.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
  },
  async ({ file_path, sheet }) => {
    try {
      const result = await listImages(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "get_sheet_properties",
  "Get sheet properties including freeze panes, auto filter, and tab color.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
  },
  async ({ file_path, sheet }) => {
    try {
      const result = await getSheetProperties(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// Cell Writing tools (5)
// =========================================================================

server.tool(
  "write_cell",
  "Set a single cell's value or formula. Start value with '=' for formulas (e.g. '=SUM(A1:A10)').",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    cell: cellAddressSchema,
    value: cellValueSchema.describe("Value to set. Start with '=' for formulas."),
  },
  async ({ file_path, sheet, cell, value }) => {
    try {
      assertCellCount(1, "write_cell", safetyConfig);
      assertWithinTemplate(sheet, cellAddrToRange(cell), safetyConfig);
      const result = await writeCell(file_path, sheet, cell, value);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "write_cells",
  "Set multiple cells at once (bulk). Each entry specifies a cell address and value.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    cells: z.array(z.object({
      cell: cellAddressSchema,
      value: cellValueSchema.describe("Value to set"),
    })).max(100000).describe("Array of cell edits (max 100,000)"),
  },
  async ({ file_path, sheet, cells }) => {
    try {
      assertCellCount(cells.length, "write_cells", safetyConfig);
      for (const c of cells) {
        assertWithinTemplate(sheet, cellAddrToRange(c.cell), safetyConfig);
      }
      const result = await writeCells(file_path, sheet, cells);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "write_row",
  "Write a row of values starting from a given row number and optional start column.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    row: rowSchema,
    values: z.array(cellValueSchema).max(16384).describe("Array of values to write (max 16,384 — Excel column limit)"),
    start_column: columnSchema.optional().describe("Start column letter (default 'A')"),
  },
  async ({ file_path, sheet, row, values, start_column }) => {
    try {
      assertCellCount(values.length, "write_row", safetyConfig);
      const startCol = columnLetterToNumber(start_column ?? "A");
      assertWithinTemplate(sheet, rowRange(row, startCol, startCol + values.length - 1), safetyConfig);
      const result = await writeRow(file_path, sheet, row, values, start_column);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "write_rows",
  "Write multiple rows of data at once (bulk). Ideal for inserting tabular data.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    start_row: rowSchema.describe("Starting row number (1-based)"),
    rows: z.array(z.array(cellValueSchema)).max(100000).describe("2D array of values: [[row1...], [row2...], ...] (max 100,000 rows)"),
    start_column: columnSchema.optional().describe("Start column letter (default 'A')"),
  },
  async ({ file_path, sheet, start_row, rows, start_column }) => {
    try {
      const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
      assertCellCount(rows.length * maxCols, "write_rows", safetyConfig);
      const startCol = columnLetterToNumber(start_column ?? "A");
      assertWithinTemplate(
        sheet,
        {
          startCol,
          endCol: startCol + Math.max(maxCols - 1, 0),
          startRow: start_row,
          endRow: start_row + Math.max(rows.length - 1, 0),
        },
        safetyConfig,
      );
      const result = await writeRows(file_path, sheet, start_row, rows, start_column);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "clear_cells",
  "Clear cell values in a range (keeps formatting).",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Range to clear (e.g. 'A1:C10')"),
  },
  async ({ file_path, sheet, range }) => {
    try {
      const r = parseRange(range);
      const cells = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
      assertCellCount(cells, "clear_cells", safetyConfig);
      assertWithinTemplate(sheet, r, safetyConfig);
      const result = await clearCells(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// Formatting tools (2)
// =========================================================================

const formatOptionsSchema = z.object({
  bold: z.boolean().optional().describe("Set bold"),
  italic: z.boolean().optional().describe("Set italic"),
  underline: z.boolean().optional().describe("Set underline"),
  strikethrough: z.boolean().optional().describe("Set strikethrough"),
  fontName: z.string().optional().describe("Font family name"),
  fontSize: z.number().min(1).max(409).optional().describe("Font size in points (1-409)"),
  fontColor: hexColorSchema.optional().describe("Font color as hex (e.g. 'FF0000')"),
  fillColor: hexColorSchema.optional().describe("Fill color as hex (e.g. 'FFFF00')"),
  fillPattern: z.enum(["solid", "none"]).optional().describe("Fill pattern"),
  borderStyle: z.enum(["thin", "medium", "thick", "double", "dotted", "dashed"]).optional().describe("Border style"),
  borderColor: hexColorSchema.optional().describe("Border color as hex"),
  borderTop: z.boolean().optional().describe("Apply border to top"),
  borderBottom: z.boolean().optional().describe("Apply border to bottom"),
  borderLeft: z.boolean().optional().describe("Apply border to left"),
  borderRight: z.boolean().optional().describe("Apply border to right"),
  horizontalAlignment: z.enum(["left", "center", "right", "justify"]).optional().describe("Horizontal alignment"),
  verticalAlignment: z.enum(["top", "middle", "bottom"]).optional().describe("Vertical alignment"),
  wrapText: z.boolean().optional().describe("Enable text wrapping"),
  textRotation: z.number().int().min(-90).max(90).optional().describe("Text rotation angle (-90 to 90)"),
  numFmt: z.string().optional().describe("Number format string (e.g. '#,##0.00', 'yyyy-mm-dd')"),
});

server.tool(
  "format_cells",
  "Apply formatting (font, fill, border, alignment, number format) to a cell range.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Cell range (e.g. 'A1:C10')"),
    format: formatOptionsSchema.describe("Format options to apply"),
  },
  async ({ file_path, sheet, range, format }) => {
    try {
      const r = parseRange(range);
      const cells = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
      assertCellCount(cells, "format_cells", safetyConfig);
      assertWithinTemplate(sheet, r, safetyConfig);
      const result = await formatCells(file_path, sheet, range, format);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "format_cells_bulk",
  "Apply different formatting to multiple ranges at once (bulk). One file I/O operation.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    groups: z.array(z.object({
      range: z.string().describe("Cell range"),
      format: formatOptionsSchema.describe("Format options"),
    })).max(1000).describe("Array of range-format groups (max 1,000)"),
  },
  async ({ file_path, sheet, groups }) => {
    try {
      let total = 0;
      for (const g of groups) {
        const r = parseRange(g.range);
        total += (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
        assertWithinTemplate(sheet, r, safetyConfig);
      }
      assertCellCount(total, "format_cells_bulk", safetyConfig);
      const result = await formatCellsBulk(file_path, sheet, groups);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// Row/Column tools (8)
// =========================================================================

server.tool(
  "set_column_width",
  "Set the width of a single column.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    column: columnSchema,
    width: z.number().min(0).max(255).describe("Column width in characters (0-255)"),
  },
  async ({ file_path, sheet, column, width }) => {
    try {
      const result = await setColumnWidth(file_path, sheet, column, width);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "set_column_widths",
  "Set widths for multiple columns at once (bulk).",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    columns: z.array(z.object({
      column: columnSchema,
      width: z.number().min(0).max(255).describe("Column width"),
    })).describe("Array of column-width pairs"),
  },
  async ({ file_path, sheet, columns }) => {
    try {
      const result = await setColumnWidths(file_path, sheet, columns);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "set_row_height",
  "Set the height of a single row.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    row: rowSchema,
    height: z.number().min(0).max(409).describe("Row height in points (0-409)"),
  },
  async ({ file_path, sheet, row, height }) => {
    try {
      const result = await setRowHeight(file_path, sheet, row, height);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "set_row_heights",
  "Set heights for multiple rows at once (bulk).",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    rows: z.array(z.object({
      row: z.number().int().min(1).describe("Row number"),
      height: z.number().min(0).max(409).describe("Row height"),
    })).describe("Array of row-height pairs"),
  },
  async ({ file_path, sheet, rows }) => {
    try {
      const result = await setRowHeights(file_path, sheet, rows);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "insert_rows",
  "Insert empty rows at the specified position. Existing rows shift down.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    row: rowSchema.describe("Row number to insert before (1-based)"),
    count: countSchema.describe("Number of rows to insert"),
  },
  async ({ file_path, sheet, row, count }) => {
    try {
      const result = await insertRows(file_path, sheet, row, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "delete_rows",
  "Delete rows at the specified position. Remaining rows shift up.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    row: rowSchema.describe("First row to delete (1-based)"),
    count: countSchema.describe("Number of rows to delete"),
  },
  async ({ file_path, sheet, row, count }) => {
    try {
      const result = await deleteRows(file_path, sheet, row, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "insert_columns",
  "Insert empty columns at the specified position. Existing columns shift right.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    column: columnSchema.describe("Column letter to insert before (e.g. 'B')"),
    count: countSchema.describe("Number of columns to insert"),
  },
  async ({ file_path, sheet, column, count }) => {
    try {
      const result = await insertColumns(file_path, sheet, column, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "delete_columns",
  "Delete columns at the specified position. Remaining columns shift left.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    column: columnSchema.describe("First column to delete (e.g. 'B')"),
    count: countSchema.describe("Number of columns to delete"),
  },
  async ({ file_path, sheet, column, count }) => {
    try {
      const result = await deleteColumns(file_path, sheet, column, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// Sheet Operation tools (4)
// =========================================================================

server.tool(
  "add_sheet",
  "Add a new empty sheet to the workbook.",
  {
    file_path: filePathSchema,
    name: z.string().describe("Name for the new sheet"),
  },
  async ({ file_path, name }) => {
    try {
      const result = await addSheet(file_path, name);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "rename_sheet",
  "Rename an existing sheet.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    new_name: z.string().describe("New name for the sheet"),
  },
  async ({ file_path, sheet, new_name }) => {
    try {
      const result = await renameSheet(file_path, sheet, new_name);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "delete_sheet",
  "Delete a sheet from the workbook.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
  },
  async ({ file_path, sheet }) => {
    try {
      const result = await deleteSheet(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "copy_sheet",
  "Copy a sheet within the workbook. Copies cell values, styles, column widths, row heights, and merged cells. Does not copy data validation, conditional formatting, or view settings.",
  {
    file_path: filePathSchema,
    source_sheet: sheetSchema.describe("Source sheet name or index"),
    new_name: z.string().describe("Name for the copied sheet"),
  },
  async ({ file_path, source_sheet, new_name }) => {
    try {
      const result = await copySheet(file_path, source_sheet, new_name);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// View Settings tools (3)
// =========================================================================

server.tool(
  "set_freeze_panes",
  "Freeze rows and/or columns. Set both to 0 to unfreeze.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    row: z.number().int().min(0).describe("Number of rows to freeze from top (0 to unfreeze)"),
    column: z.number().int().min(0).describe("Number of columns to freeze from left (0 to unfreeze)"),
  },
  async ({ file_path, sheet, row, column }) => {
    try {
      const result = await setFreeze(file_path, sheet, row, column);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "set_auto_filter",
  "Enable auto filter on a range.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Range for auto filter (e.g. 'A1:D1')"),
  },
  async ({ file_path, sheet, range }) => {
    try {
      const result = await setSheetAutoFilter(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "remove_auto_filter",
  "Remove auto filter from a sheet.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
  },
  async ({ file_path, sheet }) => {
    try {
      const result = await removeSheetAutoFilter(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// Data Validation / Named Ranges / Structure tools (7)
// =========================================================================

server.tool(
  "add_data_validation",
  "Add a data validation rule to a range (list, whole number, decimal, date, text length, custom).",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Range to apply validation (e.g. 'A1:A100')"),
    type: z.enum(["list", "whole", "decimal", "date", "textLength", "custom"]).describe("Validation type"),
    formulae: z.array(z.string()).describe("Validation formulae (e.g. ['\"Yes,No\"'] for list, ['1','100'] for range)"),
    operator: z.enum(["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"]).optional().describe("Comparison operator"),
    allow_blank: z.boolean().optional().default(true).describe("Allow blank cells"),
    show_error_message: z.boolean().optional().describe("Show error popup"),
    error_title: z.string().optional().describe("Error popup title"),
    error: z.string().optional().describe("Error popup message"),
    show_input_message: z.boolean().optional().describe("Show input hint"),
    prompt_title: z.string().optional().describe("Input hint title"),
    prompt: z.string().optional().describe("Input hint message"),
  },
  async ({ file_path, sheet, range, type, formulae, operator, allow_blank, show_error_message, error_title, error, show_input_message, prompt_title, prompt }) => {
    try {
      const result = await addDataValidation(file_path, sheet, range, {
        type,
        formulae,
        operator,
        allowBlank: allow_blank,
        showErrorMessage: show_error_message,
        errorTitle: error_title,
        error,
        showInputMessage: show_input_message,
        promptTitle: prompt_title,
        prompt,
      });
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "remove_data_validation",
  "Remove data validation rules from a range.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Range to remove validation from (e.g. 'A1:A100')"),
  },
  async ({ file_path, sheet, range }) => {
    try {
      const result = await removeDataValidation(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "add_named_range",
  "Add a named range to the workbook.",
  {
    file_path: filePathSchema,
    name: z.string().describe("Name for the range"),
    range: z.string().describe("Cell range (e.g. 'A1:C10')"),
    sheet: sheetSchema.optional().describe("Sheet the range belongs to (for scoped names)"),
  },
  async ({ file_path, name, range, sheet }) => {
    try {
      const result = await addNamedRange(file_path, name, range, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "delete_named_range",
  "Delete a named range from the workbook.",
  {
    file_path: filePathSchema,
    name: z.string().describe("Name of the range to delete"),
  },
  async ({ file_path, name }) => {
    try {
      const result = await deleteNamedRange(file_path, name);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "merge_cells",
  "Merge a range of cells.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Range to merge (e.g. 'A1:C1')"),
  },
  async ({ file_path, sheet, range }) => {
    try {
      const result = await mergeCells(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "unmerge_cells",
  "Unmerge a previously merged range of cells.",
  {
    file_path: filePathSchema,
    sheet: sheetSchema,
    range: z.string().describe("Range to unmerge (e.g. 'A1:C1')"),
  },
  async ({ file_path, sheet, range }) => {
    try {
      const result = await unmergeCells(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.tool(
  "create_workbook",
  "Create a new empty XLSX workbook. Fails if file already exists.",
  {
    file_path: filePathSchema,
    sheet_name: z.string().optional().describe("Name of the first sheet (default 'Sheet1')"),
  },
  async ({ file_path, sheet_name }) => {
    try {
      const result = await createWorkbook(file_path, sheet_name);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// ---------------------------------------------------------------------------
// Start server
// ---------------------------------------------------------------------------

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((e) => {
  console.error("Fatal error:", e);
  process.exit(1);
});
