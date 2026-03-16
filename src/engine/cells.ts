/**
 * Cell address parsing, value read/write, type conversion.
 *
 * A1 記法の解析と ExcelJS セルとのやり取りを行う。
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";

// ---------------------------------------------------------------------------
// A1 notation helpers
// ---------------------------------------------------------------------------

/** A1 アドレスから { col, row } (1-based) を返す */
export function parseCellAddress(addr: string): { col: number; row: number } {
  const m = addr.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) {
    throw new EngineError(ErrorCode.INVALID_RANGE, `Invalid cell address: ${addr}`);
  }
  const row = parseInt(m[2], 10);
  if (row < 1) {
    throw new EngineError(ErrorCode.INVALID_RANGE, `Invalid row number in address: ${addr}`);
  }
  return {
    col: columnLetterToNumber(m[1]),
    row,
  };
}

/** 列文字列 → 1-based 数値 (A=1, Z=26, AA=27, ...) */
export function columnLetterToNumber(letters: string): number {
  if (letters.length === 0) {
    throw new EngineError(ErrorCode.INVALID_RANGE, "Column letter must not be empty");
  }
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

/** 1-based 列番号 → 文字列 (1=A, 26=Z, 27=AA, ...) */
export function columnNumberToLetter(num: number): string {
  let s = "";
  let n = num;
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

// ---------------------------------------------------------------------------
// Range parsing
// ---------------------------------------------------------------------------

export interface CellRange {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

/** "A1:C5" or "A1" → CellRange (1-based) */
export function parseRange(range: string): CellRange {
  const parts = range.split(":");
  if (parts.length === 1) {
    const addr = parseCellAddress(parts[0]);
    return { startCol: addr.col, startRow: addr.row, endCol: addr.col, endRow: addr.row };
  }
  if (parts.length === 2) {
    const start = parseCellAddress(parts[0]);
    const end = parseCellAddress(parts[1]);
    return {
      startCol: Math.min(start.col, end.col),
      startRow: Math.min(start.row, end.row),
      endCol: Math.max(start.col, end.col),
      endRow: Math.max(start.row, end.row),
    };
  }
  throw new EngineError(ErrorCode.INVALID_RANGE, `Invalid range: ${range}`);
}

/** 範囲・バルク操作のセル数上限 (100,000 セル) */
export const MAX_RANGE_CELLS = 100_000;

/**
 * 範囲のセル数が上限を超えていないか検証する。
 * 書き込み・書式設定・データ検証など、セル単位でループする操作に使用。
 */
export function validateRangeSize(range: CellRange): void {
  const cells = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
  if (cells > MAX_RANGE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_RANGE,
      `Range too large (${cells.toLocaleString()} cells). Maximum is ${MAX_RANGE_CELLS.toLocaleString()} cells.`,
    );
  }
}

/** CellRange → "A1:C5" */
export function rangeToString(range: CellRange): string {
  const start = `${columnNumberToLetter(range.startCol)}${range.startRow}`;
  const end = `${columnNumberToLetter(range.endCol)}${range.endRow}`;
  return start === end ? start : `${start}:${end}`;
}

// ---------------------------------------------------------------------------
// Cell value helpers
// ---------------------------------------------------------------------------

export interface CellData {
  address: string;
  value: unknown;
  formula?: string;
  type: string;
  numFmt?: string;
  /** If this cell is the top-left of a merge, the full merge range (e.g. "A1:C1") */
  mergeRange?: string;
  /** If this cell is a non-master part of a merge, the master cell address */
  mergedWith?: string;
}

/** ExcelJS Cell → CellData */
export function getCellData(cell: ExcelJS.Cell): CellData {
  const result: CellData = {
    address: cell.address,
    value: null,
    type: "null",
  };

  // Merge info
  if (cell.isMerged) {
    const master = cell.master;
    if (master.address !== cell.address) {
      // This cell is a non-master part of a merge
      result.mergedWith = master.address;
    }
    // mergeRange for master cells is set in readSheetData (needs worksheet._merges)
  }

  const v = cell.value;
  if (v === null || v === undefined) {
    result.type = "null";
    return result;
  }

  // Formula
  if (typeof v === "object" && v !== null && "formula" in v) {
    const fv = v as ExcelJS.CellFormulaValue;
    result.formula = fv.formula;
    result.value = fv.result ?? null;
    result.type = "formula";
    if (cell.numFmt) result.numFmt = cell.numFmt;
    return result;
  }

  // SharedFormula
  if (typeof v === "object" && v !== null && "sharedFormula" in v) {
    const sv = v as ExcelJS.CellSharedFormulaValue;
    result.formula = sv.sharedFormula;
    result.value = sv.result ?? null;
    result.type = "formula";
    if (cell.numFmt) result.numFmt = cell.numFmt;
    return result;
  }

  // Rich text
  if (typeof v === "object" && v !== null && "richText" in v) {
    const rt = v as ExcelJS.CellRichTextValue;
    result.value = rt.richText.map((seg) => seg.text).join("");
    result.type = "richText";
    return result;
  }

  // Hyperlink
  if (typeof v === "object" && v !== null && "hyperlink" in v) {
    const hv = v as ExcelJS.CellHyperlinkValue;
    result.value = hv.text;
    result.type = "hyperlink";
    return result;
  }

  // Error
  if (typeof v === "object" && v !== null && "error" in v) {
    const ev = v as ExcelJS.CellErrorValue;
    result.value = ev.error;
    result.type = "error";
    return result;
  }

  // Date
  if (v instanceof Date) {
    result.value = v.toISOString();
    result.type = "date";
    if (cell.numFmt) result.numFmt = cell.numFmt;
    return result;
  }

  // Primitive types
  result.value = v;
  if (typeof v === "number") {
    result.type = "number";
    if (cell.numFmt) result.numFmt = cell.numFmt;
  } else if (typeof v === "boolean") {
    result.type = "boolean";
  } else {
    result.type = "string";
  }

  return result;
}

/**
 * セルに値を設定する。
 * value が "=" で始まる場合は数式として扱う。
 */
export function setCellValue(
  cell: ExcelJS.Cell,
  value: string | number | boolean | null,
): void {
  if (value === null) {
    cell.value = null;
    return;
  }

  if (typeof value === "string" && value.startsWith("=")) {
    cell.value = { formula: value.slice(1) } as ExcelJS.CellFormulaValue;
    return;
  }

  cell.value = value;
}

// ---------------------------------------------------------------------------
// Sheet data reading
// ---------------------------------------------------------------------------

export interface SheetData {
  sheetName: string;
  range: string;
  totalRows: number;
  totalColumns: number;
  data: RowData[];
  /** All merged cell ranges in the sheet (e.g. ["A1:C1", "D5:D10"]) */
  mergedCells?: string[];
}

export interface RowData {
  row: number;
  cells: CellData[];
}

/**
 * シートからデータを読み取る。range 指定可。
 */
export function readSheetData(
  ws: ExcelJS.Worksheet,
  range?: string,
): SheetData {
  const actualRowCount = ws.rowCount;
  const actualColCount = ws.columnCount;

  let startRow = 1;
  let endRow = actualRowCount;
  let startCol = 1;
  let endCol = actualColCount;

  if (range) {
    const parsed = parseRange(range);
    startRow = parsed.startRow;
    endRow = parsed.endRow;
    startCol = parsed.startCol;
    endCol = parsed.endCol;
  }

  // Collect merge ranges from worksheet internals.
  // ExcelJS stores merges as Range objects with .tl / .br getters.
  const merges = (ws as unknown as { _merges?: Record<string, { tl: string; br: string }> })._merges;
  const mergeMap = new Map<string, string>(); // master address → "A1:C1"
  const mergedCells: string[] = [];
  if (merges) {
    for (const [addr, dim] of Object.entries(merges)) {
      if (dim && dim.tl && dim.br) {
        const rangeLabel = `${dim.tl}:${dim.br}`;
        mergeMap.set(addr, rangeLabel);
        mergedCells.push(rangeLabel);
      }
    }
  }

  const data: RowData[] = [];

  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    const cells: CellData[] = [];
    let hasValue = false;

    for (let c = startCol; c <= endCol; c++) {
      const cell = row.getCell(c);
      if (cell.value !== null && cell.value !== undefined) {
        hasValue = true;
      }
      const cd = getCellData(cell);
      // Set mergeRange on master cells
      const mr = mergeMap.get(cell.address);
      if (mr) {
        cd.mergeRange = mr;
      }
      cells.push(cd);
    }

    // 空行をスキップ（range 指定時は含める）
    if (hasValue || range) {
      data.push({ row: r, cells });
    }
  }

  const rangeStr = range ?? (actualRowCount > 0
    ? `A1:${columnNumberToLetter(actualColCount)}${actualRowCount}`
    : "A1");

  const result: SheetData = {
    sheetName: ws.name,
    range: rangeStr,
    totalRows: actualRowCount,
    totalColumns: actualColCount,
    data,
  };
  if (mergedCells.length > 0) {
    result.mergedCells = mergedCells;
  }
  return result;
}

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

export interface SearchMatch {
  sheet: string;
  address: string;
  value: unknown;
  formula?: string;
}

/**
 * ワークシート内のセルを検索する。
 */
export function searchInSheet(
  ws: ExcelJS.Worksheet,
  query: string,
  caseSensitive: boolean,
): SearchMatch[] {
  const matches: SearchMatch[] = [];
  const q = caseSensitive ? query : query.toLowerCase();

  ws.eachRow((row) => {
    row.eachCell((cell) => {
      const data = getCellData(cell);
      const textValue = String(data.value ?? "");
      const target = caseSensitive ? textValue : textValue.toLowerCase();
      if (target.includes(q)) {
        const m: SearchMatch = {
          sheet: ws.name,
          address: cell.address,
          value: data.value,
        };
        if (data.formula) m.formula = data.formula;
        matches.push(m);
      }
    });
  });

  return matches;
}
