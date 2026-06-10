/**
 * Cell address parsing, value read/write, type conversion.
 *
 * A1 記法の解析と ExcelJS セルとのやり取りを行う。
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";
import { summarizeCellStyle, type CellFormatOptions } from "./formatting.js";

// ---------------------------------------------------------------------------
// A1 notation helpers
// ---------------------------------------------------------------------------

/** Excel の最大行数 (1,048,576) */
export const EXCEL_MAX_ROWS = 1_048_576;
/** Excel の最大列数 (XFD = 16,384) */
export const EXCEL_MAX_COLS = 16_384;

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
  if (row > EXCEL_MAX_ROWS) {
    throw new EngineError(
      ErrorCode.ROW_OUT_OF_RANGE,
      `Row ${row} exceeds Excel's maximum row ${EXCEL_MAX_ROWS.toLocaleString()} (address: ${addr})`,
    );
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
  if (n > EXCEL_MAX_COLS) {
    throw new EngineError(
      ErrorCode.COLUMN_OUT_OF_RANGE,
      `Column "${letters.toUpperCase()}" exceeds Excel's maximum column XFD (${EXCEL_MAX_COLS.toLocaleString()} columns)`,
    );
  }
  return n;
}

/**
 * 行番号と終端列が Excel の上限内に収まっているか検証する。
 * write_row / write_rows のように A1 アドレスを経由しない書き込みで使用。
 */
export function validateCellBounds(row: number, endCol: number): void {
  if (row > EXCEL_MAX_ROWS) {
    throw new EngineError(
      ErrorCode.ROW_OUT_OF_RANGE,
      `Row ${row} exceeds Excel's maximum row ${EXCEL_MAX_ROWS.toLocaleString()}`,
    );
  }
  if (endCol > EXCEL_MAX_COLS) {
    throw new EngineError(
      ErrorCode.COLUMN_OUT_OF_RANGE,
      `Column ${endCol} (${columnNumberToLetter(endCol)}) exceeds Excel's maximum column XFD (${EXCEL_MAX_COLS.toLocaleString()} columns)`,
    );
  }
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
  if (/^[A-Za-z]+$/.test(parts[0]) || /^\d+$/.test(parts[0])) {
    throw new EngineError(
      ErrorCode.INVALID_RANGE,
      `Whole-column/whole-row ranges like "A:A" or "1:3" are not supported. ` +
        `Specify explicit cells, e.g. "A1:A100" (range given: ${range}).`,
    );
  }
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
  /** Hyperlink target URL (for hyperlink cells) */
  hyperlink?: string;
  /**
   * True when a formula has no cached result (e.g. it was just written by
   * this server). The real value appears after Excel recalculates the file.
   */
  uncalculated?: boolean;
  /**
   * Degraded shared-formula group: the master cell address when the per-cell
   * formula could not be resolved (master missing or not a formula cell).
   * Never reported in `formula`.
   */
  sharedGroupMaster?: string;
  /** Cell formatting in format_cells vocabulary (only when includeStyles) */
  style?: CellFormatOptions;
  /** Cell note (comment) text */
  note?: string;
}

/** ExcelJS の note（string または richText オブジェクト）をプレーン文字列にする */
export function flattenNote(note: unknown): string | undefined {
  if (typeof note === "string") return note;
  if (typeof note === "object" && note !== null && "texts" in note) {
    const texts = (note as { texts: Array<{ text: string }> }).texts;
    if (Array.isArray(texts)) return texts.map((t) => t.text).join("");
  }
  return undefined;
}

/** 数式のキャッシュ済み結果を CellData.value 用に正規化する */
function normalizeFormulaResult(result: CellData, res: unknown): void {
  if (res === undefined || res === null) {
    result.value = null;
    result.uncalculated = true;
  } else if (res instanceof Date) {
    result.value = res.toISOString();
  } else if (typeof res === "object" && "error" in (res as Record<string, unknown>)) {
    result.value = (res as ExcelJS.CellErrorValue).error;
  } else {
    result.value = res;
  }
}

/** ExcelJS Cell → CellData */
export function getCellData(cell: ExcelJS.Cell): CellData {
  const result: CellData = {
    address: cell.address,
    value: null,
    type: "null",
  };

  const note = flattenNote(cell.note);
  if (note) result.note = note;

  // Merge info
  if (cell.isMerged) {
    const master = cell.master;
    if (master.address !== cell.address) {
      // Non-master part of a merge — return reference only, no duplicated value
      result.mergedWith = master.address;
      return result;
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
    normalizeFormulaResult(result, fv.result);
    result.type = "formula";
    if (cell.numFmt) result.numFmt = cell.numFmt;
    return result;
  }

  // SharedFormula (slave cell of a shared-formula group).
  // ExcelJS stores the *master cell's address* in `sharedFormula` (e.g. "G2"),
  // not this cell's actual formula. The cell-level `cell.formula` getter
  // resolves the per-cell formula by sliding the master's relative references
  // to this cell's position (so H2 in a G2:I2 group becomes `$C2*E2`, not `G2`).
  //
  // 他ソフトが生成した不整合ファイルでは、マスターが数式セルでない
  // （translation が throw する）/ マスターが存在しない（undefined が返る）
  // ことがある。その場合も formula にマスターアドレスを出さず、
  // sharedGroupMaster に分離して返す。
  if (typeof v === "object" && v !== null && "sharedFormula" in v) {
    const sv = v as ExcelJS.CellSharedFormulaValue;
    let translated: string | undefined;
    try {
      const f = cell.formula;
      translated = typeof f === "string" && f.length > 0 ? f : undefined;
    } catch {
      translated = undefined;
    }
    if (translated !== undefined) {
      result.formula = translated;
    } else if (typeof sv.sharedFormula === "string") {
      result.sharedGroupMaster = sv.sharedFormula;
    }
    normalizeFormulaResult(result, sv.result);
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
    result.value = hv.text ?? hv.hyperlink;
    result.hyperlink = hv.hyperlink;
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

/** 書き込み可能なセル値。オブジェクト形式は日付・ハイパーリンク用 */
export type CellWriteValue =
  | string
  | number
  | boolean
  | null
  | { date: string }
  | { hyperlink: string; text?: string };

/**
 * セルに値を設定する。
 * - "=" で始まる文字列は数式として扱う（リテラルにしたい場合は "'=" でエスケープ）
 * - { date: "2024-01-15" } は Excel の日付値として書き込む
 * - { hyperlink: "https://...", text?: "..." } はハイパーリンクとして書き込む
 * - 結合セルの子セルへの書き込みはマスター値を黙って上書きするため拒否する
 */
export function setCellValue(
  cell: ExcelJS.Cell,
  value: CellWriteValue,
): void {
  // 結合セルの子への書き込みガード：ExcelJS は子セルへの代入を
  // マスターセルに委譲するため、意図しないセルが上書きされる。
  if (cell.isMerged && cell.master.address !== cell.address) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Cell ${cell.address} is part of a merged range (master: ${cell.master.address}). ` +
        `Writing here would overwrite the master's value. Write to ${cell.master.address} instead, or unmerge_cells first.`,
    );
  }

  if (value === null) {
    cell.value = null;
    return;
  }

  if (typeof value === "object") {
    if ("date" in value) {
      const d = new Date(value.date);
      if (Number.isNaN(d.getTime())) {
        throw new EngineError(
          ErrorCode.INVALID_PARAMETER,
          `Invalid date: ${JSON.stringify(value.date)}. Use ISO format, e.g. "2024-01-15" or "2024-01-15T09:30:00Z".`,
        );
      }
      cell.value = d;
      return;
    }
    if ("hyperlink" in value) {
      cell.value = {
        text: value.text ?? value.hyperlink,
        hyperlink: value.hyperlink,
      } as ExcelJS.CellHyperlinkValue;
      return;
    }
  }

  if (typeof value === "string") {
    if (value.startsWith("=")) {
      detachSharedFormulaGroupIfMaster(cell);
      cell.value = { formula: value.slice(1) } as ExcelJS.CellFormulaValue;
      return;
    }
    // "'=" → リテラル文字列 "=..."（Excel と同じエスケープ規則）
    if (value.startsWith("'=")) {
      detachSharedFormulaGroupIfMaster(cell);
      cell.value = value.slice(1);
      return;
    }
  }

  detachSharedFormulaGroupIfMaster(cell);
  cell.value = value;
}

/**
 * 共有数式グループのマスターセルを上書きする前に、スレーブセルを
 * それぞれ独立した通常の数式セルに変換（実体化）する。
 * これを行わないとマスター上書き後の保存が
 * "Shared Formula master must exist above and or left of clone" で失敗する。
 */
function detachSharedFormulaGroupIfMaster(cell: ExcelJS.Cell): void {
  const v = cell.value;
  const isSharedMaster =
    typeof v === "object" &&
    v !== null &&
    "formula" in v &&
    (v as { shareType?: string }).shareType === "shared";
  if (!isSharedMaster) return;
  materializeSharedFormulas(cell.worksheet, cell.address);
}

/**
 * 共有数式グループを通常の数式セルに変換する。
 * masterAddress を指定するとそのマスターのグループのみ、省略すると
 * シート内の全グループを実体化する。
 * （splice 系操作の前にも呼ぶ — ExcelJS はスレーブの sharedFormula
 * ポインタをシフトしないため、放置すると保存時に throw する。）
 */
export function materializeSharedFormulas(
  ws: ExcelJS.Worksheet,
  masterAddress?: string,
): void {
  const slaves: ExcelJS.Cell[] = [];
  const masters: ExcelJS.Cell[] = [];

  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (typeof v !== "object" || v === null) return;
      if ("sharedFormula" in v) {
        const sv = v as ExcelJS.CellSharedFormulaValue;
        if (masterAddress === undefined || sv.sharedFormula === masterAddress) {
          slaves.push(cell);
        }
      } else if (
        "formula" in v &&
        (v as { shareType?: string }).shareType === "shared"
      ) {
        if (masterAddress === undefined || cell.address === masterAddress) {
          masters.push(cell);
        }
      }
    });
  });

  for (const cell of slaves) {
    const sv = cell.value as ExcelJS.CellSharedFormulaValue;
    let translated: string | undefined;
    try {
      const f = cell.formula;
      translated = typeof f === "string" && f.length > 0 ? f : undefined;
    } catch {
      translated = undefined;
    }
    if (translated !== undefined) {
      cell.value = { formula: translated, result: sv.result } as ExcelJS.CellFormulaValue;
    } else {
      // 数式を復元できない場合はキャッシュ済み結果を値として残す
      cell.value = (sv.result ?? null) as ExcelJS.CellValue;
    }
  }

  for (const cell of masters) {
    const fv = cell.value as ExcelJS.CellFormulaValue & { shareType?: string; ref?: string };
    cell.value = { formula: fv.formula, result: fv.result } as ExcelJS.CellFormulaValue;
  }
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
  /** True when compact mode was used (merged children and empty cells omitted) */
  compact?: boolean;
  /** True when output was cut off at the cell cap — use `range` to read the rest */
  truncated?: boolean;
  /** Last row included in the output when truncated */
  truncatedAtRow?: number;
}

export interface RowData {
  row: number;
  cells: CellData[];
}

export interface ReadSheetOptions {
  range?: string;
  /** Compact mode: omit merged children and empty cells to reduce output size */
  compact?: boolean;
  /** Include per-cell style (format_cells vocabulary) in the output */
  includeStyles?: boolean;
}

/**
 * read_sheet が 1 回の呼び出しで出力するセル数の上限。
 * 超過分は行単位で打ち切り、truncated フラグで通知する
 * （LLM コンテキストの溢れ防止。range 指定で続きを読める）。
 */
export const MAX_READ_CELLS = 5_000;

/**
 * シートからデータを読み取る。range / compact 指定可。
 */
export function readSheetData(
  ws: ExcelJS.Worksheet,
  options?: ReadSheetOptions,
): SheetData {
  const range = options?.range;
  const compact = options?.compact ?? false;

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
  // mergedCells には読み取り範囲と交差する結合のみ載せる（範囲読み取りで
  // シート全体の結合一覧を返すとトークンを浪費するため）。
  const merges = (ws as unknown as { _merges?: Record<string, { tl: string; br: string }> })._merges;
  const mergeMap = new Map<string, string>(); // master address → "A1:C1"
  const mergedCells: string[] = [];
  if (merges) {
    for (const [addr, dim] of Object.entries(merges)) {
      if (dim && dim.tl && dim.br) {
        const rangeLabel = `${dim.tl}:${dim.br}`;
        mergeMap.set(addr, rangeLabel);
        const mr = parseRange(rangeLabel);
        const intersects =
          mr.startRow <= endRow && mr.endRow >= startRow &&
          mr.startCol <= endCol && mr.endCol >= startCol;
        if (intersects) {
          mergedCells.push(rangeLabel);
        }
      }
    }
  }

  const data: RowData[] = [];
  let emittedCells = 0;
  let truncated = false;
  let truncatedAtRow = 0;

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
      if (options?.includeStyles && !cd.mergedWith) {
        const style = summarizeCellStyle(cell);
        if (style) cd.style = style;
      }

      // Compact mode: skip merged children and empty non-anchor cells.
      // 数式セルは結果未計算（value === null）でも省略しない — 省略すると
      // LLM が空セルと誤認して数式を上書きする。
      if (compact) {
        if (cd.mergedWith) continue;
        if (cd.value === null && !cd.formula && !cd.mergeRange) continue;
      }

      cells.push(cd);
    }

    // 空行をスキップ（range 指定時は含める。compact 時は空行を常にスキップ）
    if (compact) {
      if (cells.length > 0) {
        data.push({ row: r, cells });
        emittedCells += cells.length;
      }
    } else if (hasValue || range) {
      data.push({ row: r, cells });
      emittedCells += cells.length;
    }

    if (emittedCells >= MAX_READ_CELLS && r < endRow) {
      truncated = true;
      truncatedAtRow = r;
      break;
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
  if (compact) {
    result.compact = true;
  }
  if (truncated) {
    result.truncated = true;
    result.truncatedAtRow = truncatedAtRow;
  }
  return result;
}

// ---------------------------------------------------------------------------
// Compact JSON encoding for read_sheet
// ---------------------------------------------------------------------------

/**
 * read_sheet の <json> ペイロード。アドレスをキーにした密なマップ形式。
 * セルの型は JSON 値の型から自明（文字列/数値/真偽値）。空セルはキーが無い。
 * 数式・日付・エラーは別マップに分離して曖昧さを無くす。
 */
export interface SheetJson {
  sheetName: string;
  range: string;
  totalRows: number;
  totalColumns: number;
  /** プレーン値（文字列・数値・真偽値）。richText/hyperlink はテキストを載せる */
  cells: Record<string, string | number | boolean>;
  /** 数式セル。f = 数式（= なし）、v = キャッシュ済み結果（未計算なら省略） */
  formulas?: Record<string, { f: string; v?: unknown }>;
  /** 日付セル（ISO 8601） */
  dates?: Record<string, string>;
  /** エラーセル（"#DIV/0!" など） */
  errors?: Record<string, string>;
  /** ハイパーリンク（アドレス → URL。表示テキストは cells 側） */
  hyperlinks?: Record<string, string>;
  /** 数値書式（設定されているセルのみ） */
  numFmts?: Record<string, string>;
  /** セルノート（コメント） */
  notes?: Record<string, string>;
  /** includeStyles 時のみ。format_cells と同じ形式 */
  styles?: Record<string, CellFormatOptions>;
  /** 劣化した共有数式グループ（アドレス → マスターアドレス） */
  sharedGroupMasters?: Record<string, string>;
  /** 読み取り範囲と交差する結合セル範囲 */
  mergedCells?: string[];
  /** セル数上限で打ち切られた場合 true。range 指定で続きを読める */
  truncated?: boolean;
  truncatedAtRow?: number;
}

/** SheetData（行ベース内部表現） → SheetJson（マップ形式） */
export function toSheetJson(data: SheetData): SheetJson {
  const out: SheetJson = {
    sheetName: data.sheetName,
    range: data.range,
    totalRows: data.totalRows,
    totalColumns: data.totalColumns,
    cells: {},
  };
  const put = <T>(key: keyof SheetJson, addr: string, value: T): void => {
    const target = out as unknown as Record<string, Record<string, T>>;
    const map = target[key] ?? {};
    map[addr] = value;
    target[key] = map;
  };

  for (const row of data.data) {
    for (const c of row.cells) {
      if (c.mergedWith) continue;
      const addr = c.address;

      if (c.formula !== undefined) {
        const entry: { f: string; v?: unknown } = { f: c.formula };
        if (!c.uncalculated) entry.v = c.value;
        put("formulas", addr, entry);
      } else if (c.sharedGroupMaster !== undefined) {
        put("sharedGroupMasters", addr, c.sharedGroupMaster);
        if (c.value !== null) out.cells[addr] = c.value as string | number | boolean;
      } else if (c.type === "date") {
        put("dates", addr, c.value as string);
      } else if (c.type === "error") {
        put("errors", addr, String(c.value));
      } else if (c.value !== null && c.value !== undefined) {
        out.cells[addr] = c.value as string | number | boolean;
      }

      if (c.hyperlink) put("hyperlinks", addr, c.hyperlink);
      if (c.numFmt) put("numFmts", addr, c.numFmt);
      if (c.note) put("notes", addr, c.note);
      if (c.style) put("styles", addr, c.style);
    }
  }

  if (data.mergedCells && data.mergedCells.length > 0) out.mergedCells = data.mergedCells;
  if (data.truncated) {
    out.truncated = true;
    out.truncatedAtRow = data.truncatedAtRow;
  }
  return out;
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
 * limit を超えるマッチは収集しない（呼び出し側が打ち切りを検出できるよう、
 * limit+1 件目までは返す運用を想定）。
 */
export function searchInSheet(
  ws: ExcelJS.Worksheet,
  query: string,
  caseSensitive: boolean,
  limit: number = Number.POSITIVE_INFINITY,
): SearchMatch[] {
  const matches: SearchMatch[] = [];
  const q = caseSensitive ? query : query.toLowerCase();

  ws.eachRow((row) => {
    if (matches.length >= limit) return;
    row.eachCell((cell) => {
      if (matches.length >= limit) return;
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
