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
  assertStructuralChangeAllowed,
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

const server = new McpServer(
  {
    name: "xlsx-editor",
    version: VERSION,
  },
  {
    instructions: [
      "Read, write, format, and manage Excel (.xlsx) workbooks.",
      "",
      "Supported: cell read/write, formulas, formatting (font/fill/border/alignment/numFmt),",
      "merged cells, sheets, named ranges, data validation, row/column ops, freeze panes,",
      "auto filter, cell notes, hyperlinks, date values, copy_range, find_replace, sort_range,",
      "sheet protection, page setup, row/column/sheet visibility.",
      "",
      "Recommended workflow: get_workbook_info → read_sheet (with range) → search_cells → edit tools.",
      "",
      "LIMITATIONS:",
      "- Formula recalculation: formulas are NOT evaluated on edit. Cells written as formulas",
      "  read back as '(not calculated)' until the file is opened in Excel (recalc-on-open is",
      "  enabled automatically on every save).",
      "- Charts, pivot tables, and slicers are NOT preserved: any write operation removes them",
      "  from the workbook. Do not edit workbooks whose charts/pivots must survive.",
      "- Macro-enabled workbooks (.xlsm/.xltm) are read-only: writes are rejected because VBA",
      "  projects cannot be preserved.",
      "- Conditional formatting rules are preserved on save, but there are no tools to read or",
      "  edit them yet.",
      "- Formula ref auto-update: inserting/deleting rows/columns does NOT shift references",
      "  inside formula text. Do structural changes BEFORE writing formulas.",
      "",
      "To write a literal string starting with '=', prefix it with a single quote ('=text).",
    ].join("\n"),
  },
);

// =========================================================================
// Reading tools (8)
// =========================================================================

server.registerTool(
  "get_workbook_info",
  {
    description: "Get metadata and structure overview of an XLSX file — sheet list, named range count, and file properties.",
    inputSchema: {
      file_path: filePathSchema,
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "read_sheet",
  {
    description: "Read cell data from a sheet (values, formulas, types). Optionally specify a range like 'A1:C10'. Use compact=true to omit empty cells and merged-cell children for token-efficient output. Output is capped at 5,000 cells — large sheets are truncated with a notice; read them in chunks via 'range'.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().optional().describe("Cell range to read (e.g. 'A1:C10'). Omit to read all data."),
      compact: z.boolean().optional().default(false).describe("Omit empty cells and merged-cell children. Reduces output for sheets with many merged cells."),
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "read_cell",
  {
    description: "Read a single cell's value, formula, type, and formatting.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      cell: cellAddressSchema,
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "search_cells",
  {
    description: "Search for text or numbers in cells. Searches all sheets by default, or specify a sheet.",
    inputSchema: {
      file_path: filePathSchema,
      query: z.string().min(1).describe("Text to search for (must not be empty)"),
      sheet: sheetSchema.optional().describe("Sheet to search in (omit for all sheets)"),
      case_sensitive: z.boolean().optional().default(false).describe("Case-sensitive search. Default false."),
      max_results: z.number().int().min(1).max(1000).optional().default(100).describe("Maximum matches to return (default 100, max 1000)"),
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
  },
  async ({ file_path, query, sheet, case_sensitive, max_results }) => {
    try {
      const result = await searchCells(file_path, query, sheet, case_sensitive, max_results);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "list_named_ranges",
  {
    description: "List all named ranges in the workbook.",
    inputSchema: {
      file_path: filePathSchema,
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "list_data_validations",
  {
    description: "List data validation rules on a sheet.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "list_images",
  {
    description: "List images embedded in a sheet.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "get_sheet_properties",
  {
    description: "Get sheet properties including freeze panes, auto filter, and tab color.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
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

server.registerTool(
  "write_cell",
  {
    description: "Set a single cell's value or formula. Start value with '=' for formulas (e.g. '=SUM(A1:A10)').",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      cell: cellAddressSchema,
      value: cellValueSchema.describe("Value to set. Start with '=' for formulas."),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "write_cells",
  {
    description: "Set multiple cells at once (bulk). Each entry specifies a cell address and value.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      cells: z.array(z.object({
        cell: cellAddressSchema,
        value: cellValueSchema.describe("Value to set"),
      })).max(100000).describe("Array of cell edits (max 100,000)"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "write_row",
  {
    description: "Write a row of values starting from a given row number and optional start column.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      row: rowSchema,
      values: z.array(cellValueSchema).max(16384).describe("Array of values to write (max 16,384 — Excel column limit)"),
      start_column: columnSchema.optional().describe("Start column letter (default 'A')"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "write_rows",
  {
    description: "Write multiple rows of data at once (bulk). Ideal for inserting tabular data.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      start_row: rowSchema.describe("Starting row number (1-based)"),
      rows: z.array(z.array(cellValueSchema)).max(100000).describe("2D array of values: [[row1...], [row2...], ...] (max 100,000 rows)"),
      start_column: columnSchema.optional().describe("Start column letter (default 'A')"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "clear_cells",
  {
    description: "Clear cell values in a range (keeps formatting).",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range to clear (e.g. 'A1:C10')"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "format_cells",
  {
    description: "Apply formatting (font, fill, border, alignment, number format) to a cell range.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Cell range (e.g. 'A1:C10')"),
      format: formatOptionsSchema.describe("Format options to apply"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "format_cells_bulk",
  {
    description: "Apply different formatting to multiple ranges at once (bulk). One file I/O operation.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      groups: z.array(z.object({
        range: z.string().describe("Cell range"),
        format: formatOptionsSchema.describe("Format options"),
      })).max(1000).describe("Array of range-format groups (max 1,000)"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "set_column_width",
  {
    description: "Set the width of a single column.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      column: columnSchema,
      width: z.number().min(0).max(255).describe("Column width in characters (0-255)"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "set_column_widths",
  {
    description: "Set widths for multiple columns at once (bulk).",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      columns: z.array(z.object({
        column: columnSchema,
        width: z.number().min(0).max(255).describe("Column width"),
      })).describe("Array of column-width pairs"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "set_row_height",
  {
    description: "Set the height of a single row.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      row: rowSchema,
      height: z.number().min(0).max(409).describe("Row height in points (0-409)"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "set_row_heights",
  {
    description: "Set heights for multiple rows at once (bulk).",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      rows: z.array(z.object({
        row: z.number().int().min(1).describe("Row number"),
        height: z.number().min(0).max(409).describe("Row height"),
      })).describe("Array of row-height pairs"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "insert_rows",
  {
    description: "Insert empty rows at the specified position. Existing rows shift down. WARNING: references inside existing formulas are NOT updated — do structural changes before writing formulas.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      row: rowSchema.describe("Row number to insert before (1-based)"),
      count: countSchema.describe("Number of rows to insert"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, row, count }) => {
    try {
      assertStructuralChangeAllowed("insert_rows", safetyConfig);
      const result = await insertRows(file_path, sheet, row, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "delete_rows",
  {
    description: "Delete rows at the specified position. Remaining rows shift up. WARNING: references inside remaining formulas are NOT updated and may break or point at wrong cells.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      row: rowSchema.describe("First row to delete (1-based)"),
      count: countSchema.describe("Number of rows to delete"),
    },
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, row, count }) => {
    try {
      assertStructuralChangeAllowed("delete_rows", safetyConfig);
      const result = await deleteRows(file_path, sheet, row, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "insert_columns",
  {
    description: "Insert empty columns at the specified position. Existing columns shift right. WARNING: references inside existing formulas are NOT updated — do structural changes before writing formulas.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      column: columnSchema.describe("Column letter to insert before (e.g. 'B')"),
      count: countSchema.describe("Number of columns to insert"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, column, count }) => {
    try {
      assertStructuralChangeAllowed("insert_columns", safetyConfig);
      const result = await insertColumns(file_path, sheet, column, count);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "delete_columns",
  {
    description: "Delete columns at the specified position. Remaining columns shift left. WARNING: references inside remaining formulas are NOT updated and may break or point at wrong cells.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      column: columnSchema.describe("First column to delete (e.g. 'B')"),
      count: countSchema.describe("Number of columns to delete"),
    },
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, column, count }) => {
    try {
      assertStructuralChangeAllowed("delete_columns", safetyConfig);
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

server.registerTool(
  "add_sheet",
  {
    description: "Add a new empty sheet to the workbook.",
    inputSchema: {
      file_path: filePathSchema,
      name: z.string().describe("Name for the new sheet"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
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

server.registerTool(
  "rename_sheet",
  {
    description: "Rename an existing sheet.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      new_name: z.string().describe("New name for the sheet"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, new_name }) => {
    try {
      assertStructuralChangeAllowed("rename_sheet", safetyConfig);
      const result = await renameSheet(file_path, sheet, new_name);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "delete_sheet",
  {
    description: "Delete a sheet from the workbook.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
    },
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet }) => {
    try {
      assertStructuralChangeAllowed("delete_sheet", safetyConfig);
      const result = await deleteSheet(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "copy_sheet",
  {
    description: "Copy a sheet within the workbook. Copies cell values, styles, column widths, row heights, and merged cells. Does not copy data validation, conditional formatting, or view settings.",
    inputSchema: {
      file_path: filePathSchema,
      source_sheet: sheetSchema.describe("Source sheet name or index"),
      new_name: z.string().describe("Name for the copied sheet"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
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

server.registerTool(
  "set_freeze_panes",
  {
    description: "Freeze rows and/or columns. Set both to 0 to unfreeze.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      freeze_rows: z.number().int().min(0).describe("Number of rows to freeze from top (0 to unfreeze)"),
      freeze_columns: z.number().int().min(0).describe("Number of columns to freeze from left (0 to unfreeze)"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, freeze_rows, freeze_columns }) => {
    try {
      const result = await setFreeze(file_path, sheet, freeze_rows, freeze_columns);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "set_auto_filter",
  {
    description: "Enable auto filter on a range.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range for auto filter (e.g. 'A1:D1')"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "remove_auto_filter",
  {
    description: "Remove auto filter from a sheet.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "add_data_validation",
  {
    description: "Add a data validation rule to a range (list, whole number, decimal, date, text length, custom).",
    inputSchema: {
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
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "remove_data_validation",
  {
    description: "Remove data validation rules from a range.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range to remove validation from (e.g. 'A1:A100')"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
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

server.registerTool(
  "add_named_range",
  {
    description: "Add a named range to the workbook.",
    inputSchema: {
      file_path: filePathSchema,
      name: z.string().describe("Name for the range"),
      range: z.string().describe("Cell range (e.g. 'A1:C10')"),
      sheet: sheetSchema.optional().describe("Sheet the range belongs to (for scoped names)"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
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

server.registerTool(
  "delete_named_range",
  {
    description: "Delete a named range from the workbook.",
    inputSchema: {
      file_path: filePathSchema,
      name: z.string().describe("Name of the range to delete"),
    },
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
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

server.registerTool(
  "merge_cells",
  {
    description: "Merge a range of cells.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range to merge (e.g. 'A1:C1')"),
    },
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, range }) => {
    try {
      const r = parseRange(range);
      const cells = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
      assertCellCount(cells, "merge_cells", safetyConfig);
      assertWithinTemplate(sheet, r, safetyConfig);
      const result = await mergeCells(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "unmerge_cells",
  {
    description: "Unmerge a previously merged range of cells.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range to unmerge (e.g. 'A1:C1')"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, range }) => {
    try {
      const r = parseRange(range);
      const cells = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
      assertCellCount(cells, "unmerge_cells", safetyConfig);
      assertWithinTemplate(sheet, r, safetyConfig);
      const result = await unmergeCells(file_path, sheet, range);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "create_workbook",
  {
    description: "Create a new empty XLSX workbook. Fails if file already exists.",
    inputSchema: {
      file_path: filePathSchema,
      sheet_name: z.string().optional().describe("Name of the first sheet (default 'Sheet1')"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
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
