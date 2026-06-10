/**
 * XLSX Engine — ExcelJS wrapper for the MCP server.
 *
 * バレルモジュール。engine/ サブモジュールを再エクスポートし、
 * index.ts が消費する公開 API 関数を定義する。
 */

import * as fs from "fs/promises";
import * as path from "path";
import { createRequire } from "node:module";
import { withFileLock } from "./engine/file-lock.js";
import ExcelJS from "exceljs";

// ExcelJS 内部の数式スライド関数（相対参照の平行移動）。公開 API には無いが
// 共有数式の実装基盤で、コピー時の参照変換に使う。
const requireCjs = createRequire(import.meta.url);
const { slideFormula } = requireCjs("exceljs/lib/utils/shared-formula.js") as {
  slideFormula: (formula: string, fromAddr: string, toAddr: string) => string;
};

// Re-export types and helpers
export { ErrorCode, EngineError } from "./engine/xlsx-io.js";
export type { ErrorCodeType } from "./engine/xlsx-io.js";
export type { CellData, SheetData, RowData, SearchMatch, CellRange, ReadSheetOptions, CellWriteValue, SheetJson } from "./engine/cells.js";
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
  validateCellBounds,
  MAX_RANGE_CELLS,
  MAX_READ_CELLS,
  columnNumberToLetter,
  columnLetterToNumber,
  getCellData,
  setCellValue,
  readSheetData,
  toSheetJson,
  searchInSheet,
  rangeToString,
  materializeSharedFormulas,
  flattenNote,
} from "./engine/cells.js";
import type { CellData, SearchMatch, CellWriteValue } from "./engine/cells.js";
import {
  type CellFormatOptions,
  type CellFormatBulkGroup,
  applyCellFormat,
  summarizeCellStyle,
} from "./engine/formatting.js";
import {
  addWorksheet,
  renameWorksheet,
  deleteWorksheet,
  copyWorksheet,
  validateSheetName,
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
  includeStyles?: boolean,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const ws = resolveSheet(handle.workbook, sheet);
  // マップ形式（アドレスがキー）では空セル・結合子セルはキー不在で表現できる
  // ため、常に compact で走査する。
  const data = readSheetData(ws, { range, compact: true, includeStyles });
  const json = toSheetJson(data);

  // セルデータは <json> ブロックのみに載せる。テキスト部にも全セルを並べると
  // 出力トークンが約 2 倍になるため、テキスト部はサマリだけにする。
  const lines: string[] = [];
  lines.push(`Sheet: "${json.sheetName}" | Range: ${json.range}`);
  // dates / errors / sharedGroupMasters も値を持つ。アドレスの和集合で数える
  // （sharedGroupMasters のキーは cells にも現れ得るため単純加算は二重計上になる）
  const nonEmptyCount = new Set([
    ...Object.keys(json.cells),
    ...Object.keys(json.formulas ?? {}),
    ...Object.keys(json.dates ?? {}),
    ...Object.keys(json.errors ?? {}),
    ...Object.keys(json.sharedGroupMasters ?? {}),
  ]).size;
  lines.push(`Total: ${json.totalRows} rows × ${json.totalColumns} columns | ${nonEmptyCount} non-empty cell(s) returned (data in the JSON block below; absent address = empty cell)`);
  if (json.truncated) {
    lines.push(
      `⚠ Output truncated at row ${json.truncatedAtRow} (cell cap ${MAX_READ_CELLS.toLocaleString()}). ` +
      `Use the 'range' parameter (e.g. 'A${(json.truncatedAtRow ?? 0) + 1}:...') to read the remaining rows.`,
    );
  }
  if (json.mergedCells && json.mergedCells.length > 0) {
    lines.push(`Merged cells in range: ${json.mergedCells.join(", ")}`);
  }

  return lines.join("\n") + "\n\n<json>" + JSON.stringify(json) + "</json>";
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

  // 書式は format_cells が受け取る形式（CellFormatOptions）に要約して返す。
  // ExcelJS の生 style オブジェクトより大幅に小さく、そのまま format_cells に
  // 渡して書式を複製できる。
  const style = summarizeCellStyle(c);
  const result = style ? { ...data, style } : data;

  const lines: string[] = [];
  lines.push(`Cell ${data.address}: ${data.value ?? (data.uncalculated ? "(not calculated)" : "(empty)")}`);
  if (data.formula) lines.push(`Formula: =${data.formula}`);
  lines.push(`Type: ${data.type}`);
  if (data.hyperlink) lines.push(`Hyperlink: ${data.hyperlink}`);
  if (data.mergeRange) lines.push(`Merge: master of ${data.mergeRange}`);
  if (data.mergedWith) lines.push(`Merge: part of ${data.mergedWith}`);

  return lines.join("\n") + "\n\n<json>" + JSON.stringify(result) + "</json>";
}

export async function searchCells(
  filePath: string,
  query: string,
  sheet?: string | number,
  caseSensitive: boolean = false,
  maxResults: number = 100,
): Promise<string> {
  const handle = await openXlsx(filePath);
  const matches: SearchMatch[] = [];
  // maxResults+1 件目まで集めて打ち切りの有無を検出する
  const collectLimit = maxResults + 1;

  if (sheet !== undefined) {
    const ws = resolveSheet(handle.workbook, sheet);
    matches.push(...searchInSheet(ws, query, caseSensitive, collectLimit));
  } else {
    for (const ws of handle.workbook.worksheets) {
      if (matches.length >= collectLimit) break;
      matches.push(...searchInSheet(ws, query, caseSensitive, collectLimit - matches.length));
    }
  }

  const truncated = matches.length > maxResults;
  const shown = truncated ? matches.slice(0, maxResults) : matches;

  const lines: string[] = [];
  if (truncated) {
    lines.push(
      `Found ${maxResults}+ match(es) for "${query}" (output capped at ${maxResults}). ` +
      `Narrow the query, specify a sheet, or raise max_results.`,
    );
  } else {
    lines.push(`Found ${shown.length} match(es) for "${query}"`);
  }
  for (const m of shown) {
    const val = m.formula ? `=${m.formula} → ${m.value}` : String(m.value ?? "");
    lines.push(`  [${m.sheet}] ${m.address}: ${val}`);
  }

  const payload: { matches: SearchMatch[]; truncated?: boolean } = { matches: shown };
  if (truncated) payload.truncated = true;
  return lines.join("\n") + "\n\n<json>" + JSON.stringify(payload) + "</json>";
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
  if (sheetName !== undefined) {
    validateSheetName(sheetName);
  }
  return withFileLock(filePath, async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet(sheetName ?? "Sheet1");
    const buffer = await wb.xlsx.writeBuffer();
    try {
      // flag "wx": 既存ファイルがあれば EEXIST — 上書き防止を OS レベルで保証する
      await fs.writeFile(filePath, Buffer.from(buffer as ArrayBuffer), { flag: "wx" });
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === "EEXIST") {
        throw new EngineError(
          ErrorCode.INVALID_PARAMETER,
          `File already exists: ${filePath}. Delete it first or use a different path.`,
        );
      }
      throw e;
    }
    return `Created workbook: ${filePath}`;
  });
}

export async function writeCell(
  filePath: string,
  sheet: string | number,
  cell: string,
  value: CellWriteValue,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const addr = parseCellAddress(cell);
    const c = ws.getRow(addr.row).getCell(addr.col);
    setCellValue(c, value);
    await saveXlsx(handle);
    return `Set ${cell} = ${typeof value === "object" && value !== null ? JSON.stringify(value) : value}`;
  });
}

export async function writeCells(
  filePath: string,
  sheet: string | number,
  cells: Array<{ cell: string; value: CellWriteValue }>,
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
  values: CellWriteValue[],
  startColumn?: string,
): Promise<string> {
  const startColNum = startColumn ? columnLetterToNumber(startColumn) : 1;
  validateCellBounds(row, startColNum + Math.max(values.length - 1, 0));
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const r = ws.getRow(row);
    for (let i = 0; i < values.length; i++) {
      const c = r.getCell(startColNum + i);
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
  rows: CellWriteValue[][],
  startColumn?: string,
): Promise<string> {
  const totalCells = rows.reduce((sum, row) => sum + row.length, 0);
  if (totalCells > MAX_RANGE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Too many cells (${totalCells.toLocaleString()}). Maximum is ${MAX_RANGE_CELLS.toLocaleString()}.`,
    );
  }
  const startColNum = startColumn ? columnLetterToNumber(startColumn) : 1;
  const maxRowLen = rows.reduce((m, r) => Math.max(m, r.length), 0);
  validateCellBounds(
    startRow + Math.max(rows.length - 1, 0),
    startColNum + Math.max(maxRowLen - 1, 0),
  );
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    for (let ri = 0; ri < rows.length; ri++) {
      const r = ws.getRow(startRow + ri);
      for (let ci = 0; ci < rows[ri].length; ci++) {
        const c = r.getCell(startColNum + ci);
        setCellValue(c, rows[ri][ci]);
      }
      r.commit();
    }
    await saveXlsx(handle);
    return `Wrote ${rows.length} row(s) starting at row ${startRow}`;
  });
}

export type ClearMode = "values" | "formats" | "all";

export async function clearCells(
  filePath: string,
  sheet: string | number,
  range: string,
  mode: ClearMode = "values",
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
        const cell = row.getCell(c);
        // 結合セルの子はスキップ（ExcelJS は子への代入をマスターに委譲するため、
        // 範囲外のマスターを意図せずクリアしてしまう）
        if (cell.isMerged && cell.master.address !== cell.address) continue;
        if (mode === "values" || mode === "all") {
          // 共有数式マスターを直接 null にすると残されたスレーブが宙に浮き
          // 保存が失敗するため、グループを実体化してからクリアする
          const v = cell.value;
          if (
            typeof v === "object" && v !== null && "formula" in v &&
            (v as { shareType?: string }).shareType === "shared"
          ) {
            materializeSharedFormulas(ws, cell.address);
          }
          cell.value = null;
        }
        if (mode === "formats" || mode === "all") {
          cell.style = {};
        }
        count++;
      }
    }
    await saveXlsx(handle);
    const what = mode === "values" ? "values" : mode === "formats" ? "formatting" : "values and formatting";
    return `Cleared ${what} of ${count} cell(s) in ${range}`;
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
  validateCellBounds(row, 1);
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
  for (const entry of rows) {
    validateCellBounds(entry.row, 1);
  }
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
  inheritStyle: boolean = false,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    insertRowsAt(ws, row, count);
    if (inheritStyle) {
      // 挿入位置の直上の行（先頭挿入時は直下の行）から書式を引き継ぐ
      const refRowNum = row > 1 ? row - 1 : row + count;
      const refRow = ws.getRow(refRowNum);
      const colCount = Math.max(ws.columnCount, 1);
      for (let i = 0; i < count; i++) {
        const newRow = ws.getRow(row + i);
        if (refRow.height) newRow.height = refRow.height;
        for (let c = 1; c <= colCount; c++) {
          newRow.getCell(c).style = structuredClone(refRow.getCell(c).style);
        }
      }
    }
    await saveXlsx(handle);
    return `Inserted ${count} row(s) at row ${row}${inheritStyle ? " (styles inherited)" : ""}`;
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

// =========================================================================
// Range operations — copy, find/replace, sort
// =========================================================================

/**
 * 範囲をコピーする（値・数式・書式・結合）。
 * 数式の相対参照は移動量に合わせて変換される（絶対参照 $A$1 は不変）。
 */
export async function copyRange(
  filePath: string,
  sheet: string | number,
  sourceRange: string,
  destination: string,
  destSheet?: string | number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const wsDest = destSheet !== undefined ? resolveSheet(handle.workbook, destSheet) : ws;

    const src = parseRange(sourceRange);
    validateRangeSize(src);
    const dst = parseCellAddress(destination);
    const rowOffset = dst.row - src.startRow;
    const colOffset = dst.col - src.startCol;
    validateCellBounds(src.endRow + rowOffset, src.endCol + colOffset);

    // 共有数式を実体化してから扱う（スレーブのポインタはコピーできない）。
    // コピー先シートも実体化する — コピーで共有数式マスターを上書きすると
    // 残されたスレーブが宙に浮き、保存が失敗するため。
    materializeSharedFormulas(ws);
    if (wsDest !== ws) {
      materializeSharedFormulas(wsDest);
    }

    // スナップショット（コピー元とコピー先が重なる場合に備えて先に全部読む）
    interface CellSnap {
      r: number;
      c: number;
      value: ExcelJS.CellValue;
      style: Partial<ExcelJS.Style>;
    }
    const snaps: CellSnap[] = [];
    for (let r = src.startRow; r <= src.endRow; r++) {
      const row = ws.getRow(r);
      for (let c = src.startCol; c <= src.endCol; c++) {
        const cell = row.getCell(c);
        // 結合の子セルは null として扱う（結合自体は別途再現する）
        const isMergedChild = cell.isMerged && cell.master.address !== cell.address;
        snaps.push({
          r,
          c,
          value: isMergedChild ? null : cell.value,
          style: structuredClone(cell.style),
        });
      }
    }

    // コピー元の範囲に完全に含まれる結合をコピー先座標で再現する
    const srcMerges = (ws as unknown as { _merges?: Record<string, { tl: string; br: string }> })._merges ?? {};
    const mergesToCreate: string[] = [];
    for (const dim of Object.values(srcMerges)) {
      if (!dim?.tl || !dim?.br) continue;
      const m = parseRange(`${dim.tl}:${dim.br}`);
      const contained =
        m.startRow >= src.startRow && m.endRow <= src.endRow &&
        m.startCol >= src.startCol && m.endCol <= src.endCol;
      if (contained) {
        mergesToCreate.push(rangeToString({
          startRow: m.startRow + rowOffset,
          endRow: m.endRow + rowOffset,
          startCol: m.startCol + colOffset,
          endCol: m.endCol + colOffset,
        }));
      }
    }

    // コピー先に既存の結合があれば解除してから書き込む
    const destRegion: typeof src = {
      startRow: src.startRow + rowOffset,
      endRow: src.endRow + rowOffset,
      startCol: src.startCol + colOffset,
      endCol: src.endCol + colOffset,
    };
    const destMerges = (wsDest as unknown as { _merges?: Record<string, { tl: string; br: string }> })._merges ?? {};
    for (const dim of Object.values(destMerges)) {
      if (!dim?.tl || !dim?.br) continue;
      const m = parseRange(`${dim.tl}:${dim.br}`);
      const intersects =
        m.startRow <= destRegion.endRow && m.endRow >= destRegion.startRow &&
        m.startCol <= destRegion.endCol && m.endCol >= destRegion.startCol;
      if (intersects) {
        try {
          wsDest.unMergeCells(`${dim.tl}:${dim.br}`);
        } catch {
          // 解除済みなら続行
        }
      }
    }

    // 書き込み
    for (const snap of snaps) {
      const destCell = wsDest.getRow(snap.r + rowOffset).getCell(snap.c + colOffset);
      let value = snap.value;
      if (typeof value === "object" && value !== null && "formula" in value) {
        const fv = value as ExcelJS.CellFormulaValue;
        const fromAddr = `${columnNumberToLetter(snap.c)}${snap.r}`;
        const toAddr = `${columnNumberToLetter(snap.c + colOffset)}${snap.r + rowOffset}`;
        let translated = fv.formula;
        try {
          translated = slideFormula(fv.formula, fromAddr, toAddr);
        } catch {
          // 変換に失敗したら原文のままコピーする
        }
        // 移動後のキャッシュ結果は信頼できないので落とす（開いた時に再計算される）
        value = { formula: translated } as ExcelJS.CellFormulaValue;
      }
      destCell.value = value;
      destCell.style = structuredClone(snap.style);
    }

    // 結合を再現
    const wsDestInternal = wsDest as unknown as { mergeCellsWithoutStyle?: (range: string) => void };
    for (const m of mergesToCreate) {
      try {
        if (typeof wsDestInternal.mergeCellsWithoutStyle === "function") {
          wsDestInternal.mergeCellsWithoutStyle(m);
        } else {
          wsDest.mergeCells(m);
        }
      } catch {
        // 重複等で失敗した結合は諦める
      }
    }

    await saveXlsx(handle);
    const destLabel = destSheet !== undefined ? `${wsDest.name}!${rangeToString(destRegion)}` : rangeToString(destRegion);
    return `Copied ${sourceRange} → ${destLabel} (${snaps.length} cells)`;
  });
}

/**
 * 文字列セルの検索置換。プレーン文字列セルのみ対象
 * （数式・数値・リッチテキスト・ハイパーリンクは変更しない）。
 */
export async function findReplace(
  filePath: string,
  query: string,
  replacement: string,
  sheet?: string | number,
  caseSensitive: boolean = false,
  matchEntireCell: boolean = false,
): Promise<string> {
  if (query.length === 0) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, "query must not be empty");
  }
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const sheets = sheet !== undefined
      ? [resolveSheet(handle.workbook, sheet)]
      : handle.workbook.worksheets;

    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(escaped, caseSensitive ? "g" : "gi");

    let replacedCells = 0;
    const samples: string[] = [];
    for (const ws of sheets) {
      ws.eachRow((row) => {
        row.eachCell((cell) => {
          const v = cell.value;
          if (typeof v !== "string") return;
          if (cell.isMerged && cell.master.address !== cell.address) return;
          let next: string | undefined;
          if (matchEntireCell) {
            const equal = caseSensitive ? v === query : v.toLowerCase() === query.toLowerCase();
            if (equal) next = replacement;
          } else {
            regex.lastIndex = 0;
            if (regex.test(v)) {
              regex.lastIndex = 0;
              // 置換文字列は常にリテラル扱いにする（関数を渡さないと $& や
              // $' などの JS 置換パターンが展開されてデータが壊れる）
              next = v.replace(regex, () => replacement);
            }
          }
          if (next !== undefined && next !== v) {
            cell.value = next;
            replacedCells++;
            if (samples.length < 10) samples.push(`[${ws.name}] ${cell.address}`);
          }
        });
      });
    }

    if (replacedCells === 0) {
      return `No cells matched "${query}" — file unchanged`;
    }
    await saveXlsx(handle);
    const sampleText = samples.join(", ") + (replacedCells > samples.length ? ", …" : "");
    return `Replaced in ${replacedCells} cell(s): ${sampleText}`;
  });
}

/**
 * 範囲を行単位でソートする（値・数式・書式が行ごとに一緒に移動する）。
 * 範囲に結合セルが交差する場合はエラー。
 * 数式の相対参照は移動先に合わせて変換される。
 */
export async function sortRange(
  filePath: string,
  sheet: string | number,
  range: string,
  keyColumn: string,
  ascending: boolean = true,
  hasHeader: boolean = false,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const r = parseRange(range);
    validateRangeSize(r);
    const keyCol = columnLetterToNumber(keyColumn);
    if (keyCol < r.startCol || keyCol > r.endCol) {
      throw new EngineError(
        ErrorCode.INVALID_PARAMETER,
        `Key column ${keyColumn.toUpperCase()} is outside the range ${range}`,
      );
    }

    // 結合セルが範囲に交差していたら拒否（行の入れ替えで壊れるため）
    const merges = (ws as unknown as { _merges?: Record<string, { tl: string; br: string }> })._merges ?? {};
    for (const dim of Object.values(merges)) {
      if (!dim?.tl || !dim?.br) continue;
      const m = parseRange(`${dim.tl}:${dim.br}`);
      const intersects =
        m.startRow <= r.endRow && m.endRow >= r.startRow &&
        m.startCol <= r.endCol && m.endCol >= r.startCol;
      if (intersects) {
        throw new EngineError(
          ErrorCode.INVALID_PARAMETER,
          `Range ${range} intersects merged cells (${dim.tl}:${dim.br}). Unmerge before sorting.`,
        );
      }
    }

    materializeSharedFormulas(ws);

    const dataStartRow = r.startRow + (hasHeader ? 1 : 0);
    if (dataStartRow > r.endRow) {
      return `Nothing to sort in ${range}`;
    }

    interface SortRow {
      originalRow: number;
      cells: Array<{ value: ExcelJS.CellValue; style: Partial<ExcelJS.Style> }>;
      key: unknown;
    }
    const rows: SortRow[] = [];
    for (let rowNum = dataStartRow; rowNum <= r.endRow; rowNum++) {
      const row = ws.getRow(rowNum);
      const cells: SortRow["cells"] = [];
      for (let c = r.startCol; c <= r.endCol; c++) {
        const cell = row.getCell(c);
        cells.push({ value: cell.value, style: structuredClone(cell.style) });
      }
      const keyCell = row.getCell(keyCol);
      let key: unknown = keyCell.value;
      if (typeof key === "object" && key !== null) {
        if ("result" in key) key = (key as ExcelJS.CellFormulaValue).result ?? null;
        else if ("richText" in key) key = (key as ExcelJS.CellRichTextValue).richText.map((s) => s.text).join("");
        else if ("text" in key) key = (key as ExcelJS.CellHyperlinkValue).text;
        else if (key instanceof Date) key = key.getTime();
        else key = null;
      }
      if (key instanceof Date) key = key.getTime();
      rows.push({ originalRow: rowNum, cells, key });
    }

    // 数値 < 文字列 < 空（Excel の昇順と同じ並び）。空は方向に関わらず末尾。
    const rank = (k: unknown): number => {
      if (k === null || k === undefined || k === "") return 2;
      if (typeof k === "number" || typeof k === "boolean") return 0;
      return 1;
    };
    rows.sort((a, b) => {
      const ra = rank(a.key);
      const rb = rank(b.key);
      if (ra === 2 || rb === 2) return ra - rb;
      if (ra !== rb) return ascending ? ra - rb : rb - ra;
      let cmp = 0;
      if (ra === 0) cmp = Number(a.key) - Number(b.key);
      else cmp = String(a.key).localeCompare(String(b.key));
      return ascending ? cmp : -cmp;
    });

    // 書き戻し（数式の相対参照は行移動に合わせて変換）
    for (let i = 0; i < rows.length; i++) {
      const destRowNum = dataStartRow + i;
      const srcRowNum = rows[i].originalRow;
      const destRow = ws.getRow(destRowNum);
      for (let c = r.startCol; c <= r.endCol; c++) {
        const snap = rows[i].cells[c - r.startCol];
        const destCell = destRow.getCell(c);
        let value = snap.value;
        if (typeof value === "object" && value !== null && "formula" in value && srcRowNum !== destRowNum) {
          const fv = value as ExcelJS.CellFormulaValue;
          const fromAddr = `${columnNumberToLetter(c)}${srcRowNum}`;
          const toAddr = `${columnNumberToLetter(c)}${destRowNum}`;
          let translated = fv.formula;
          try {
            translated = slideFormula(fv.formula, fromAddr, toAddr);
          } catch {
            // 変換失敗時は原文のまま
          }
          value = { formula: translated } as ExcelJS.CellFormulaValue;
        }
        destCell.value = value;
        destCell.style = snap.style;
      }
    }

    await saveXlsx(handle);
    return `Sorted ${rows.length} row(s) in ${range} by column ${keyColumn.toUpperCase()} (${ascending ? "ascending" : "descending"})`;
  });
}

// =========================================================================
// Cell notes
// =========================================================================

export async function setCellNote(
  filePath: string,
  sheet: string | number,
  cell: string,
  note: string | null,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const addr = parseCellAddress(cell);
    const c = ws.getRow(addr.row).getCell(addr.col);
    if (note === null) {
      (c as unknown as { note: unknown }).note = undefined;
    } else {
      c.note = note;
      // ExcelJS は値も書式も無いセルをシート XML に書き出さないため、
      // 空セルのノートはファイル上で宙に浮き、次の保存で消える。
      // 既定値と同じ alignment を明示して（見た目は不変、styleId は非ゼロ）
      // セルをシリアライズ対象にする。
      const isBareCell =
        (c.value === null || c.value === undefined) &&
        !summarizeCellStyle(c);
      if (isBareCell) {
        c.alignment = { vertical: "bottom" };
      }
    }
    await saveXlsx(handle);
    return note === null ? `Removed note from ${cell}` : `Set note on ${cell}`;
  });
}

// =========================================================================
// Sheet visibility / tab color / protection / page setup
// =========================================================================

export async function setSheetProperties(
  filePath: string,
  sheet: string | number,
  options: { state?: "visible" | "hidden" | "veryHidden"; tabColor?: string | null },
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const changes: string[] = [];

    if (options.state !== undefined && options.state !== ws.state) {
      if (options.state !== "visible") {
        const otherVisible = handle.workbook.worksheets.some(
          (s) => s.id !== ws.id && (s.state === "visible" || !s.state),
        );
        if (!otherVisible) {
          throw new EngineError(
            ErrorCode.INVALID_PARAMETER,
            `Cannot hide "${ws.name}": a workbook must keep at least one visible sheet.`,
          );
        }
      }
      ws.state = options.state;
      changes.push(`state=${options.state}`);
    }

    if (options.tabColor !== undefined) {
      if (options.tabColor === null) {
        ws.properties.tabColor = undefined as unknown as ExcelJS.Color;
        changes.push("tabColor removed");
      } else {
        ws.properties.tabColor = { argb: `FF${options.tabColor}` };
        changes.push(`tabColor=#${options.tabColor}`);
      }
    }

    if (changes.length === 0) {
      return `No changes for sheet "${ws.name}"`;
    }
    await saveXlsx(handle);
    return `Updated sheet "${ws.name}": ${changes.join(", ")}`;
  });
}

export async function setRowVisibility(
  filePath: string,
  sheet: string | number,
  startRow: number,
  endRow: number,
  hidden: boolean,
): Promise<string> {
  if (endRow < startRow) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `end_row (${endRow}) must be >= start_row (${startRow})`);
  }
  validateCellBounds(endRow, 1);
  if (endRow - startRow + 1 > MAX_RANGE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Too many rows (${(endRow - startRow + 1).toLocaleString()}). Maximum is ${MAX_RANGE_CELLS.toLocaleString()}.`,
    );
  }
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    for (let r = startRow; r <= endRow; r++) {
      const row = ws.getRow(r);
      row.hidden = hidden;
      // ExcelJS はセルも明示的な高さも無い行をシリアライズしない
      // （hidden フラグだけでは保存時に消える）ため、空行を隠す場合は
      // 既定の行高を明示してシリアライズ対象にする。
      if (hidden && row.actualCellCount === 0 && !row.height) {
        row.height = ws.properties?.defaultRowHeight || 15;
      }
    }
    await saveXlsx(handle);
    return `${hidden ? "Hid" : "Unhid"} rows ${startRow}-${endRow}`;
  });
}

export async function setColumnVisibility(
  filePath: string,
  sheet: string | number,
  startColumn: string,
  endColumn: string,
  hidden: boolean,
): Promise<string> {
  const startCol = columnLetterToNumber(startColumn);
  const endCol = columnLetterToNumber(endColumn);
  if (endCol < startCol) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `end_column (${endColumn}) must be >= start_column (${startColumn})`);
  }
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    for (let c = startCol; c <= endCol; c++) {
      ws.getColumn(c).hidden = hidden;
    }
    await saveXlsx(handle);
    return `${hidden ? "Hid" : "Unhid"} columns ${startColumn.toUpperCase()}-${endColumn.toUpperCase()}`;
  });
}

export async function protectSheet(
  filePath: string,
  sheet: string | number,
  password?: string,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    await ws.protect(password ?? "", {});
    await saveXlsx(handle);
    return `Protected sheet "${ws.name}"${password ? " with password" : ""}`;
  });
}

export async function unprotectSheet(
  filePath: string,
  sheet: string | number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    ws.unprotect();
    await saveXlsx(handle);
    return `Unprotected sheet "${ws.name}"`;
  });
}

export interface PageSetupOptions {
  orientation?: "portrait" | "landscape";
  printArea?: string;
  fitToWidth?: number;
  fitToHeight?: number;
  paperSize?: "A4" | "A3" | "letter" | "legal";
}

const PAPER_SIZES: Record<string, number> = { letter: 1, legal: 5, A3: 8, A4: 9 };

export async function setPageSetup(
  filePath: string,
  sheet: string | number,
  options: PageSetupOptions,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openXlsx(filePath);
    const ws = resolveSheet(handle.workbook, sheet);
    const changes: string[] = [];
    ws.pageSetup = ws.pageSetup ?? {};

    if (options.orientation !== undefined) {
      ws.pageSetup.orientation = options.orientation;
      changes.push(`orientation=${options.orientation}`);
    }
    if (options.printArea !== undefined) {
      parseRange(options.printArea); // 形式検証
      ws.pageSetup.printArea = options.printArea;
      changes.push(`printArea=${options.printArea}`);
    }
    if (options.fitToWidth !== undefined || options.fitToHeight !== undefined) {
      ws.pageSetup.fitToPage = true;
      if (options.fitToWidth !== undefined) {
        ws.pageSetup.fitToWidth = options.fitToWidth;
        changes.push(`fitToWidth=${options.fitToWidth}`);
      }
      if (options.fitToHeight !== undefined) {
        ws.pageSetup.fitToHeight = options.fitToHeight;
        changes.push(`fitToHeight=${options.fitToHeight}`);
      }
    }
    if (options.paperSize !== undefined) {
      ws.pageSetup.paperSize = PAPER_SIZES[options.paperSize] as ExcelJS.PaperSize;
      changes.push(`paperSize=${options.paperSize}`);
    }

    if (changes.length === 0) {
      return `No page setup changes for sheet "${ws.name}"`;
    }
    await saveXlsx(handle);
    return `Updated page setup for "${ws.name}": ${changes.join(", ")}`;
  });
}
