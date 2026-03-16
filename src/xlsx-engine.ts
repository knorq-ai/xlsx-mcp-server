/**
 * XLSX Engine — ExcelJS wrapper for the MCP server.
 *
 * バレルモジュール。engine/ サブモジュールを再エクスポートし、
 * index.ts が消費する公開 API 関数を定義する。
 */

import * as fs from "fs/promises";
import * as path from "path";
import { withFileLock } from "./engine/file-lock.js";
import ExcelJS from "exceljs";

// Re-export types and helpers
export { ErrorCode, EngineError } from "./engine/xlsx-io.js";
export type { ErrorCodeType } from "./engine/xlsx-io.js";
export type { CellData, SheetData, RowData, SearchMatch, CellRange } from "./engine/cells.js";
export type { CellFormatOptions, CellFormatBulkGroup } from "./engine/formatting.js";
export type { ImageInfo } from "./engine/images.js";

// Internal imports
import {
  ErrorCode,
  EngineError,
  openXlsx,
  saveXlsx,
  resolveSheet,
} from "./engine/xlsx-io.js";
import {
  parseCellAddress,
  parseRange,
  validateRangeSize,
  MAX_RANGE_CELLS,
  columnNumberToLetter,
  columnLetterToNumber,
  getCellData,
  setCellValue,
  readSheetData,
  searchInSheet,
  rangeToString,
} from "./engine/cells.js";
import type { CellData, SearchMatch } from "./engine/cells.js";
import {
  type CellFormatOptions,
  type CellFormatBulkGroup,
  applyCellFormat,
} from "./engine/formatting.js";
import {
  addWorksheet,
  renameWorksheet,
  deleteWorksheet,
  copyWorksheet,
} from "./engine/sheets.js";
import {
  insertRowsAt,
  deleteRowsAt,
  insertColumnsAt,
  deleteColumnsAt,
} from "./engine/rows-columns.js";
import {
  type DataValidationParams,
  addDataValidationRule,
  removeDataValidationRule,
} from "./engine/data-validation.js";
import { listSheetImages, type ImageInfo } from "./engine/images.js";
import {
  setFreezePanes,
  setAutoFilter,
  removeAutoFilter,
} from "./engine/view-settings.js";
import {
  addNamedRange as addNamedRangeImpl,
  deleteNamedRange as deleteNamedRangeImpl,
  listNamedRanges,
} from "./engine/named-ranges.js";

// =========================================================================
// Reading functions (no file lock needed)
// =========================================================================

export async function getWorkbookInfo(filePath: string): Promise<string> {
  const handle = await openXlsx(filePath);
  const wb = handle.workbook;

  const sheets = wb.worksheets.map((ws, i) => ({
    index: i + 1,
    name: ws.name,
    state: ws.state || "visible",
    rowCount: ws.rowCount,
    columnCount: ws.columnCount,
  }));

  const namedRanges = listNamedRanges(wb);

  const info = {
    fileName: path.basename(filePath),
    sheetCount: sheets.length,
    sheets,
    namedRangeCount: namedRanges.length,
    creator: wb.creator || undefined,
    lastModifiedBy: wb.lastModifiedBy || undefined,
    created: wb.created ? wb.created.toISOString() : undefined,
    modified: wb.modified ? wb.modified.toISOString() : undefined,
  };

  const lines: string[] = [];
  lines.push(`Workbook: ${info.fileName}`);
  lines.push(`Sheets: ${info.sheetCount}`);
  for (const s of sheets) {
    lines.push(`  [${s.index}] "${s.name}" (${s.state}) — ${s.rowCount} rows × ${s.columnCount} cols`);
  }
  if (namedRanges.length > 0) {
    lines.push(`Named ranges: ${namedRanges.length}`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify(info) + "</json>";
}

export async function readSheet(
  filePath: string,
  sheet: string | number,
  range?: string,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const ws = resolveSheet(handle.workbook, sheet);
  const data = readSheetData(ws, range);

  const lines: string[] = [];
  lines.push(`Sheet: "${data.sheetName}" | Range: ${data.range}`);
  lines.push(`Total: ${data.totalRows} rows × ${data.totalColumns} columns`);
  if (data.mergedCells && data.mergedCells.length > 0) {
    lines.push(`Merged cells: ${data.mergedCells.join(", ")}`);
  }
  lines.push("");

  for (const row of data.data) {
    const cells = row.cells
      .map((c) => {
        const val = c.formula ? `=${c.formula} → ${c.value}` : String(c.value ?? "");
        let label = `${c.address}: ${val}`;
        if (c.mergeRange) label += ` [merged: ${c.mergeRange}]`;
        else if (c.mergedWith) label += ` [→${c.mergedWith}]`;
        return label;
      })
      .join(" | ");
    lines.push(`Row ${row.row}: ${cells}`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify(data) + "</json>";
}

export async function readCell(
  filePath: string,
  sheet: string | number,
  cell: string,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const ws = resolveSheet(handle.workbook, sheet);
  const addr = parseCellAddress(cell);
  const c = ws.getRow(addr.row).getCell(addr.col);
  const data = getCellData(c);

  // Set mergeRange for master cells
  const merges = (ws as unknown as { _merges?: Record<string, { tl: string; br: string }> })._merges;
  if (merges) {
    const dim = merges[c.address];
    if (dim && dim.tl && dim.br) {
      data.mergeRange = `${dim.tl}:${dim.br}`;
    }
  }

  // Include style info
  const style: Record<string, unknown> = {};
  if (c.font) style.font = c.font;
  if (c.fill && (c.fill as ExcelJS.FillPattern).fgColor) style.fill = c.fill;
  if (c.border) style.border = c.border;
  if (c.alignment) style.alignment = c.alignment;
  if (c.numFmt) style.numFmt = c.numFmt;

  const result = { ...data, style };

  const lines: string[] = [];
  lines.push(`Cell ${data.address}: ${data.value ?? "(empty)"}`);
  if (data.formula) lines.push(`Formula: =${data.formula}`);
  lines.push(`Type: ${data.type}`);
  if (data.mergeRange) lines.push(`Merge: master of ${data.mergeRange}`);
  if (data.mergedWith) lines.push(`Merge: part of ${data.mergedWith}`);

  return lines.join("\n") + "\n\n<json>" + JSON.stringify(result) + "</json>";
}

export async function searchCells(
  filePath: string,
  query: string,
  sheet?: string | number,
  caseSensitive: boolean = false,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const matches: SearchMatch[] = [];

  if (sheet !== undefined) {
    const ws = resolveSheet(handle.workbook, sheet);
    matches.push(...searchInSheet(ws, query, caseSensitive));
  } else {
    for (const ws of handle.workbook.worksheets) {
      matches.push(...searchInSheet(ws, query, caseSensitive));
    }
  }

  const lines: string[] = [];
  lines.push(`Found ${matches.length} match(es) for "${query}"`);
  for (const m of matches) {
    const val = m.formula ? `=${m.formula} → ${m.value}` : String(m.value ?? "");
    lines.push(`  [${m.sheet}] ${m.address}: ${val}`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify({ matches }) + "</json>";
}

export async function getSheetProperties(
  filePath: string,
  sheet: string | number,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const ws = resolveSheet(handle.workbook, sheet);

  const props: Record<string, unknown> = {
    name: ws.name,
    state: ws.state || "visible",
    rowCount: ws.rowCount,
    columnCount: ws.columnCount,
  };

  // Freeze panes
  const views = ws.views;
  if (views && views.length > 0) {
    const v = views[0];
    if (v.state === "frozen") {
      props.freezePanes = {
        row: v.ySplit ?? 0,
        column: v.xSplit ?? 0,
      };
    }
  }

  // Auto filter
  if (ws.autoFilter) {
    props.autoFilter = ws.autoFilter;
  }

  // Tab color
  if (ws.properties?.tabColor) {
    props.tabColor = ws.properties.tabColor;
  }

  const lines: string[] = [];
  lines.push(`Sheet: "${ws.name}"`);
  lines.push(`State: ${props.state}`);
  lines.push(`Size: ${ws.rowCount} rows × ${ws.columnCount} columns`);
  if (props.freezePanes) {
    const fp = props.freezePanes as { row: number; column: number };
    lines.push(`Freeze panes: row ${fp.row}, col ${fp.column}`);
  }
  if (props.autoFilter) lines.push(`Auto filter: active`);

  return lines.join("\n") + "\n\n<json>" + JSON.stringify(props) + "</json>";
}

export async function listWorkbookNamedRanges(filePath: string): Promise<string> {
  const handle = await openXlsx(filePath);
  const ranges = listNamedRanges(handle.workbook);

  const lines: string[] = [];
  lines.push(`Named ranges: ${ranges.length}`);
  for (const r of ranges) {
    lines.push(`  ${r.name}: ${r.range}`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify({ namedRanges: ranges }) + "</json>";
}

export async function listDataValidations(
  filePath: string,
  sheet: string | number,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const ws = resolveSheet(handle.workbook, sheet);

  const validations: Array<{ address: string; type: string; formulae?: string[] }> = [];
  // ExcelJS stores data validations at model level after file reload
  const dvMap = (ws.model as unknown as Record<string, unknown>).dataValidations as
    Record<string, { type?: string; formulae?: string[] }> | undefined;
  if (dvMap) {
    for (const [address, dv] of Object.entries(dvMap)) {
      if (dv && dv.type) {
        validations.push({
          address,
          type: dv.type,
          formulae: dv.formulae,
        });
      }
    }
  }

  const lines: string[] = [];
  lines.push(`Data validations on "${ws.name}": ${validations.length}`);
  for (const v of validations) {
    lines.push(`  ${v.address}: ${v.type}${v.formulae ? ` [${v.formulae.join(", ")}]` : ""}`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify({ validations }) + "</json>";
}

export async function listImages(
  filePath: string,
  sheet: string | number,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const ws = resolveSheet(handle.workbook, sheet);
  const images = listSheetImages(handle.workbook, ws);

  const lines: string[] = [];
  lines.push(`Images on "${ws.name}": ${images.length}`);
  for (const img of images) {
    lines.push(`  ${img.name}: ${img.extension} (${img.width}×${img.height})`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify({ images }) + "</json>";
}

// =========================================================================
// Writing functions (file-locked)
// =========================================================================

export async function createWorkbook(
  filePath: string,
  sheetName?: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    // 既存ファイルの上書き防止
    try {
      await fs.access(filePath);
      throw new EngineError(
        ErrorCode.INVALID_PARAMETER,
        `File already exists: ${filePath}. Delete it first or use a different path.`,
      );
    } catch (e) {
      if (e instanceof EngineError) throw e;
      // ファイルが存在しない — 正常
    }
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet(sheetName ?? "Sheet1");
    await wb.xlsx.writeFile(filePath);
    return `Created workbook: ${filePath}`;
  });
}

export async function writeCell(
  filePath: string,
  sheet: string | number,
  cell: string,
  value: string | number | boolean | null,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const addr = parseCellAddress(cell);
    const c = ws.getRow(addr.row).getCell(addr.col);
    setCellValue(c, value);
    await saveXlsx(handle);
    return `Set ${cell} = ${value}`;
  });
}

export async function writeCells(
  filePath: string,
  sheet: string | number,
  cells: Array<{ cell: string; value: string | number | boolean | null }>,
): Promise<string> {
  if (cells.length > MAX_RANGE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Too many cells (${cells.length.toLocaleString()}). Maximum is ${MAX_RANGE_CELLS.toLocaleString()}.`,
    );
  }
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    for (const entry of cells) {
      const addr = parseCellAddress(entry.cell);
      const c = ws.getRow(addr.row).getCell(addr.col);
      setCellValue(c, entry.value);
    }
    await saveXlsx(handle);
    return `Updated ${cells.length} cell(s)`;
  });
}

export async function writeRow(
  filePath: string,
  sheet: string | number,
  row: number,
  values: Array<string | number | boolean | null>,
  startColumn?: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const startCol = startColumn ? columnLetterToNumber(startColumn) : 1;
    const r = ws.getRow(row);
    for (let i = 0; i < values.length; i++) {
      const c = r.getCell(startCol + i);
      setCellValue(c, values[i]);
    }
    r.commit();
    await saveXlsx(handle);
    return `Wrote ${values.length} value(s) to row ${row}`;
  });
}

export async function writeRows(
  filePath: string,
  sheet: string | number,
  startRow: number,
  rows: Array<Array<string | number | boolean | null>>,
  startColumn?: string,
): Promise<string> {
  const totalCells = rows.reduce((sum, row) => sum + row.length, 0);
  if (totalCells > MAX_RANGE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Too many cells (${totalCells.toLocaleString()}). Maximum is ${MAX_RANGE_CELLS.toLocaleString()}.`,
    );
  }
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const startCol = startColumn ? columnLetterToNumber(startColumn) : 1;
    for (let ri = 0; ri < rows.length; ri++) {
      const r = ws.getRow(startRow + ri);
      for (let ci = 0; ci < rows[ri].length; ci++) {
        const c = r.getCell(startCol + ci);
        setCellValue(c, rows[ri][ci]);
      }
      r.commit();
    }
    await saveXlsx(handle);
    return `Wrote ${rows.length} row(s) starting at row ${startRow}`;
  });
}

export async function clearCells(
  filePath: string,
  sheet: string | number,
  range: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const parsed = parseRange(range);
    validateRangeSize(parsed);
    let count = 0;
    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      const row = ws.getRow(r);
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        row.getCell(c).value = null;
        count++;
      }
    }
    await saveXlsx(handle);
    return `Cleared ${count} cell(s) in ${range}`;
  });
}

// =========================================================================
// Formatting
// =========================================================================

export async function formatCells(
  filePath: string,
  sheet: string | number,
  range: string,
  format: CellFormatOptions,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const parsed = parseRange(range);
    validateRangeSize(parsed);
    let count = 0;
    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      const row = ws.getRow(r);
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        applyCellFormat(row.getCell(c), format);
        count++;
      }
    }
    await saveXlsx(handle);
    return `Formatted ${count} cell(s) in ${range}`;
  });
}

export async function formatCellsBulk(
  filePath: string,
  sheet: string | number,
  groups: CellFormatBulkGroup[],
): Promise<string> {
  // 各グループの個別検証 + 累計セル数の検証をロック取得前に行う
  let cumulativeCells = 0;
  for (const group of groups) {
    const parsed = parseRange(group.range);
    validateRangeSize(parsed);
    cumulativeCells += (parsed.endRow - parsed.startRow + 1) * (parsed.endCol - parsed.startCol + 1);
  }
  if (cumulativeCells > MAX_RANGE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Total cells across all groups too large (${cumulativeCells.toLocaleString()}). Maximum is ${MAX_RANGE_CELLS.toLocaleString()}.`,
    );
  }
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    let totalCount = 0;
    for (const group of groups) {
      const parsed = parseRange(group.range);
      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        const row = ws.getRow(r);
        for (let c = parsed.startCol; c <= parsed.endCol; c++) {
          applyCellFormat(row.getCell(c), group.format);
          totalCount++;
        }
      }
    }
    await saveXlsx(handle);
    return `Formatted ${totalCount} cell(s) across ${groups.length} group(s)`;
  });
}

// =========================================================================
// Sheet operations
// =========================================================================

export async function addSheet(
  filePath: string,
  name: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    addWorksheet(handle.workbook, name);
    await saveXlsx(handle);
    return `Added sheet: "${name}"`;
  });
}

export async function renameSheet(
  filePath: string,
  sheet: string | number,
  newName: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const oldName = ws.name;
    renameWorksheet(handle.workbook, ws, newName);
    await saveXlsx(handle);
    return `Renamed sheet "${oldName}" → "${newName}"`;
  });
}

export async function deleteSheet(
  filePath: string,
  sheet: string | number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const name = ws.name;
    deleteWorksheet(handle.workbook, ws);
    await saveXlsx(handle);
    return `Deleted sheet: "${name}"`;
  });
}

export async function copySheet(
  filePath: string,
  sourceSheet: string | number,
  newName: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sourceSheet);
    copyWorksheet(handle.workbook, ws, newName);
    await saveXlsx(handle);
    return `Copied sheet "${ws.name}" → "${newName}"`;
  });
}

// =========================================================================
// Row / Column operations
// =========================================================================

export async function setColumnWidth(
  filePath: string,
  sheet: string | number,
  column: string,
  width: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const colNum = columnLetterToNumber(column);
    ws.getColumn(colNum).width = width;
    await saveXlsx(handle);
    return `Set column ${column} width = ${width}`;
  });
}

export async function setColumnWidths(
  filePath: string,
  sheet: string | number,
  columns: Array<{ column: string; width: number }>,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    for (const entry of columns) {
      const colNum = columnLetterToNumber(entry.column);
      ws.getColumn(colNum).width = entry.width;
    }
    await saveXlsx(handle);
    return `Set width for ${columns.length} column(s)`;
  });
}

export async function setRowHeight(
  filePath: string,
  sheet: string | number,
  row: number,
  height: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    ws.getRow(row).height = height;
    await saveXlsx(handle);
    return `Set row ${row} height = ${height}`;
  });
}

export async function setRowHeights(
  filePath: string,
  sheet: string | number,
  rows: Array<{ row: number; height: number }>,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    for (const entry of rows) {
      ws.getRow(entry.row).height = entry.height;
    }
    await saveXlsx(handle);
    return `Set height for ${rows.length} row(s)`;
  });
}

export async function insertRows(
  filePath: string,
  sheet: string | number,
  row: number,
  count: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    insertRowsAt(ws, row, count);
    await saveXlsx(handle);
    return `Inserted ${count} row(s) at row ${row}`;
  });
}

export async function deleteRows(
  filePath: string,
  sheet: string | number,
  row: number,
  count: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    deleteRowsAt(ws, row, count);
    await saveXlsx(handle);
    return `Deleted ${count} row(s) at row ${row}`;
  });
}

export async function insertColumns(
  filePath: string,
  sheet: string | number,
  column: string,
  count: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const colNum = columnLetterToNumber(column);
    insertColumnsAt(ws, colNum, count);
    await saveXlsx(handle);
    return `Inserted ${count} column(s) at column ${column}`;
  });
}

export async function deleteColumns(
  filePath: string,
  sheet: string | number,
  column: string,
  count: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const colNum = columnLetterToNumber(column);
    deleteColumnsAt(ws, colNum, count);
    await saveXlsx(handle);
    return `Deleted ${count} column(s) at column ${column}`;
  });
}

// =========================================================================
// View settings
// =========================================================================

export async function setFreeze(
  filePath: string,
  sheet: string | number,
  row: number,
  column: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    setFreezePanes(ws, row, column);
    await saveXlsx(handle);
    return `Set freeze panes: row ${row}, column ${column}`;
  });
}

export async function setSheetAutoFilter(
  filePath: string,
  sheet: string | number,
  range: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    setAutoFilter(ws, range);
    await saveXlsx(handle);
    return `Set auto filter: ${range}`;
  });
}

export async function removeSheetAutoFilter(
  filePath: string,
  sheet: string | number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    removeAutoFilter(ws);
    await saveXlsx(handle);
    return `Removed auto filter`;
  });
}

// =========================================================================
// Data validation
// =========================================================================

export async function addDataValidation(
  filePath: string,
  sheet: string | number,
  range: string,
  params: DataValidationParams,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    addDataValidationRule(ws, range, params);
    await saveXlsx(handle);
    return `Added data validation (${params.type}) to ${range}`;
  });
}

export async function removeDataValidation(
  filePath: string,
  sheet: string | number,
  range: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    removeDataValidationRule(ws, range);
    await saveXlsx(handle);
    return `Removed data validation from ${range}`;
  });
}

// =========================================================================
// Named ranges
// =========================================================================

export async function addNamedRange(
  filePath: string,
  name: string,
  range: string,
  sheet?: string | number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const sheetName = sheet !== undefined
      ? resolveSheet(handle.workbook, sheet).name
      : undefined;
    addNamedRangeImpl(handle.workbook, name, range, sheetName);
    await saveXlsx(handle);
    return `Added named range "${name}" = ${range}`;
  });
}

export async function deleteNamedRange(
  filePath: string,
  name: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    deleteNamedRangeImpl(handle.workbook, name);
    await saveXlsx(handle);
    return `Deleted named range "${name}"`;
  });
}

// =========================================================================
// Merge cells
// =========================================================================

export async function mergeCells(
  filePath: string,
  sheet: string | number,
  range: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    ws.mergeCells(range);
    await saveXlsx(handle);
    return `Merged cells: ${range}`;
  });
}

export async function unmergeCells(
  filePath: string,
  sheet: string | number,
  range: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    ws.unMergeCells(range);
    await saveXlsx(handle);
    return `Unmerged cells: ${range}`;
  });
}
