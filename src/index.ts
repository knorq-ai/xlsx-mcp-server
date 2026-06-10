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
  copyRange,
  findReplace,
  sortRange,
  setCellNote,
  setSheetProperties,
  setRowVisibility,
  setColumnVisibility,
  protectSheet,
  unprotectSheet,
  setPageSetup,
  EngineError,
  ErrorCode,
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
const cellValueSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
  z.object({ date: z.string().describe("ISO date, e.g. '2024-01-15' or '2024-01-15T09:30:00Z'") }).strict(),
  z.object({
    hyperlink: z.string().describe("Target URL"),
    text: z.string().optional().describe("Display text (defaults to the URL)"),
  }).strict(),
]).describe("Cell value. Strings starting with '=' are written as formulas ('=text to escape a literal). Use {date: 'ISO'} for true Excel dates, {hyperlink, text} for links.");
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
    description: "Read cell data from a sheet as JSON maps keyed by cell address: 'cells' (plain values), 'formulas' ({f, v}), 'dates', 'errors', 'notes', 'mergedCells'. An absent address means an empty cell. Optionally specify a range like 'A1:C10'. Output is capped at 5,000 cells — large sheets are truncated with a notice; read them in chunks via 'range'. Set include_styles=true to also get per-cell formatting in the same shape format_cells accepts.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().optional().describe("Cell range to read (e.g. 'A1:C10'). Omit to read all data."),
      include_styles: z.boolean().optional().default(false).describe("Include per-cell formatting (format_cells vocabulary). Increases output size."),
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, range, include_styles }) => {
    try {
      const result = await readSheet(file_path, sheet, range, include_styles);
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
    description: "Clear cell values and/or formatting in a range. mode='values' (default) keeps formatting, 'formats' keeps values, 'all' clears both.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range to clear (e.g. 'A1:C10')"),
      mode: z.enum(["values", "formats", "all"]).optional().default("values").describe("What to clear (default 'values')"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, range, mode }) => {
    try {
      const r = parseRange(range);
      const cells = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
      assertCellCount(cells, "clear_cells", safetyConfig);
      assertWithinTemplate(sheet, r, safetyConfig);
      const result = await clearCells(file_path, sheet, range, mode);
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
    description: "Insert empty rows at the specified position. Existing rows shift down. Set inherit_style=true to copy formatting from the row above. WARNING: references inside existing formulas are NOT updated — do structural changes before writing formulas.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      row: rowSchema.describe("Row number to insert before (1-based)"),
      count: countSchema.describe("Number of rows to insert"),
      inherit_style: z.boolean().optional().default(false).describe("Copy formatting (and row height) from the row above the insertion point"),
    },
    annotations: { destructiveHint: false, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, row, count, inherit_style }) => {
    try {
      assertStructuralChangeAllowed("insert_rows", safetyConfig);
      const result = await insertRows(file_path, sheet, row, count, inherit_style);
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

// =========================================================================
// Range tools — copy, find/replace, sort
// =========================================================================

server.registerTool(
  "copy_range",
  {
    description: "Copy a cell range (values, formulas, formatting, merges) to another location, optionally on a different sheet. Relative references inside formulas are shifted to the destination ($-anchored references stay fixed). Existing content at the destination is overwritten.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema.describe("Source sheet"),
      source_range: z.string().describe("Range to copy (e.g. 'A1:C10')"),
      destination: cellAddressSchema.describe("Top-left cell of the destination (e.g. 'E1')"),
      dest_sheet: sheetSchema.optional().describe("Destination sheet (defaults to the source sheet)"),
    },
    // 重複領域コピーは再実行のたびに結果が変わるため idempotent ではない
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, sheet, source_range, destination, dest_sheet }) => {
    try {
      const src = parseRange(source_range);
      const cells = (src.endRow - src.startRow + 1) * (src.endCol - src.startCol + 1);
      assertCellCount(cells, "copy_range", safetyConfig);
      const dst = parseCellAddress(destination);
      const destRegion: CellRange = {
        startRow: dst.row,
        startCol: dst.col,
        endRow: dst.row + (src.endRow - src.startRow),
        endCol: dst.col + (src.endCol - src.startCol),
      };
      assertWithinTemplate(dest_sheet ?? sheet, destRegion, safetyConfig);
      const result = await copyRange(file_path, sheet, source_range, destination, dest_sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "find_replace",
  {
    description: "Find and replace text across plain string cells. Formulas, numbers, rich text, and hyperlinks are not modified. Searches all sheets unless one is specified.",
    inputSchema: {
      file_path: filePathSchema,
      query: z.string().min(1).describe("Text to find"),
      replacement: z.string().describe("Replacement text (may be empty to delete the match)"),
      sheet: sheetSchema.optional().describe("Sheet to operate on (omit for all sheets)"),
      case_sensitive: z.boolean().optional().default(false).describe("Case-sensitive matching. Default false."),
      match_entire_cell: z.boolean().optional().default(false).describe("Replace only cells whose entire content equals the query"),
    },
    annotations: { destructiveHint: true, idempotentHint: false, openWorldHint: false },
  },
  async ({ file_path, query, replacement, sheet, case_sensitive, match_entire_cell }) => {
    try {
      if (safetyConfig.templateMode) {
        throw new EngineError(
          ErrorCode.OUTSIDE_TEMPLATE_RANGE,
          "find_replace is disabled in template mode: it cannot be constrained to the declared ranges. Use search_cells + write_cells instead.",
        );
      }
      const result = await findReplace(file_path, query, replacement, sheet, case_sensitive, match_entire_cell);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "sort_range",
  {
    description: "Sort the rows of a range by a key column (values, formulas, and formatting move together; relative formula references are shifted). Numbers sort before text; empty cells always sort last. Fails if the range intersects merged cells.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      range: z.string().describe("Range to sort (e.g. 'A2:F50')"),
      key_column: columnSchema.describe("Column letter to sort by (must be inside the range)"),
      ascending: z.boolean().optional().default(true).describe("Sort direction. Default ascending."),
      has_header: z.boolean().optional().default(false).describe("Skip the first row of the range (header)"),
    },
    annotations: { destructiveHint: true, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, range, key_column, ascending, has_header }) => {
    try {
      const r = parseRange(range);
      const cells = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
      assertCellCount(cells, "sort_range", safetyConfig);
      assertWithinTemplate(sheet, r, safetyConfig);
      const result = await sortRange(file_path, sheet, range, key_column, ascending, has_header);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

// =========================================================================
// Cell note / sheet appearance / protection / page setup tools
// =========================================================================

server.registerTool(
  "set_cell_note",
  {
    description: "Set or remove a cell note (comment). Pass null to remove.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      cell: cellAddressSchema,
      note: z.union([z.string(), z.null()]).describe("Note text, or null to remove the note"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, cell, note }) => {
    try {
      // ノートも納品物に現れるコンテンツなので、テンプレートモードの
      // ホワイトリストを適用する
      assertCellCount(1, "set_cell_note", safetyConfig);
      assertWithinTemplate(sheet, cellAddrToRange(cell), safetyConfig);
      const result = await setCellNote(file_path, sheet, cell, note);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "set_sheet_properties",
  {
    description: "Set sheet visibility (visible/hidden/veryHidden) and/or tab color. A workbook must keep at least one visible sheet.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      state: z.enum(["visible", "hidden", "veryHidden"]).optional().describe("Sheet visibility. 'veryHidden' sheets can only be re-shown programmatically."),
      tab_color: z.union([hexColorSchema, z.null()]).optional().describe("Tab color as 6-char hex (e.g. 'FF0000'), or null to remove"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, state, tab_color }) => {
    try {
      const result = await setSheetProperties(file_path, sheet, { state, tabColor: tab_color });
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "set_row_visibility",
  {
    description: "Hide or unhide a range of rows.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      start_row: rowSchema.describe("First row"),
      end_row: rowSchema.describe("Last row (inclusive)"),
      hidden: z.boolean().describe("true to hide, false to unhide"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, start_row, end_row, hidden }) => {
    try {
      const result = await setRowVisibility(file_path, sheet, start_row, end_row, hidden);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "set_column_visibility",
  {
    description: "Hide or unhide a range of columns.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      start_column: columnSchema.describe("First column letter"),
      end_column: columnSchema.describe("Last column letter (inclusive)"),
      hidden: z.boolean().describe("true to hide, false to unhide"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, start_column, end_column, hidden }) => {
    try {
      const result = await setColumnVisibility(file_path, sheet, start_column, end_column, hidden);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "protect_sheet",
  {
    description: "Protect a sheet against editing in Excel, optionally with a password. NOTE: this is Excel UI-level protection, not encryption — this server and other tools can still modify the file.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      password: z.string().optional().describe("Optional protection password"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, password }) => {
    try {
      const result = await protectSheet(file_path, sheet, password);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "unprotect_sheet",
  {
    description: "Remove sheet protection.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet }) => {
    try {
      const result = await unprotectSheet(file_path, sheet);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return { content: [{ type: "text", text: formatError(e) }], isError: true };
    }
  },
);

server.registerTool(
  "set_page_setup",
  {
    description: "Configure print/PDF layout: orientation, print area, fit-to-page, paper size.",
    inputSchema: {
      file_path: filePathSchema,
      sheet: sheetSchema,
      orientation: z.enum(["portrait", "landscape"]).optional().describe("Page orientation"),
      print_area: z.string().optional().describe("Print area range (e.g. 'A1:G40')"),
      fit_to_width: z.number().int().min(1).optional().describe("Fit to N pages wide"),
      fit_to_height: z.number().int().min(1).optional().describe("Fit to N pages tall"),
      paper_size: z.enum(["A4", "A3", "letter", "legal"]).optional().describe("Paper size"),
    },
    annotations: { destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ file_path, sheet, orientation, print_area, fit_to_width, fit_to_height, paper_size }) => {
    try {
      const result = await setPageSetup(file_path, sheet, {
        orientation,
        printArea: print_area,
        fitToWidth: fit_to_width,
        fitToHeight: fit_to_height,
        paperSize: paper_size,
      });
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
