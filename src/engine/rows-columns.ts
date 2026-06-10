/**
 * Row and column operations — insert, delete.
 *
 * ExcelJS の spliceRows / spliceColumns はセル内容と定義名はシフトするが、
 * シリアライズに使われる結合セル情報（_merges）とデータ検証はシフトしない。
 * そのまま保存すると結合範囲が元の位置に残り、レイアウトが壊れる。
 * ここでは splice の前に結合を解除し、splice 後にシフト済みの座標で
 * 再結合・データ検証の再配置を行う。
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";
import {
  parseRange,
  rangeToString,
  materializeSharedFormulas,
  EXCEL_MAX_ROWS as MAX_ROWS,
  EXCEL_MAX_COLS as MAX_COLS,
  type CellRange,
} from "./cells.js";

function validateRowBounds(row: number, count: number): void {
  if (row < 1 || row > MAX_ROWS) {
    throw new EngineError(ErrorCode.ROW_OUT_OF_RANGE, `Row ${row} out of range (1-${MAX_ROWS})`);
  }
  if (count < 1) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `Count must be at least 1`);
  }
  if (count > MAX_ROWS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Count ${count.toLocaleString()} exceeds Excel's row limit (${MAX_ROWS.toLocaleString()})`,
    );
  }
}

function validateColumnBounds(col: number, count: number): void {
  if (col < 1 || col > MAX_COLS) {
    throw new EngineError(ErrorCode.COLUMN_OUT_OF_RANGE, `Column ${col} out of range (1-${MAX_COLS})`);
  }
  if (count < 1) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `Count must be at least 1`);
  }
  if (count > MAX_COLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Count ${count.toLocaleString()} exceeds Excel's column limit (${MAX_COLS.toLocaleString()})`,
    );
  }
}

// ---------------------------------------------------------------------------
// Range shifting helpers
// ---------------------------------------------------------------------------

/**
 * 挿入後の [start, end] 座標。
 * pos 以降に始まる範囲はシフト、pos を跨ぐ範囲は拡張（Excel と同じ挙動）。
 */
function mapInsert(
  start: number,
  end: number,
  pos: number,
  count: number,
): [number, number] {
  if (start >= pos) return [start + count, end + count];
  if (end >= pos) return [start, end + count];
  return [start, end];
}

/**
 * pos から count 個削除した後の [start, end] 座標。
 * 範囲全体が削除域に含まれる場合は null。
 */
function mapDelete(
  start: number,
  end: number,
  pos: number,
  count: number,
): [number, number] | null {
  const delEnd = pos + count - 1;
  const newStart = start < pos ? start : start > delEnd ? start - count : pos;
  const newEnd = end < pos ? end : end > delEnd ? end - count : pos - 1;
  if (newEnd < newStart) return null;
  return [newStart, newEnd];
}

type Axis = "row" | "col";
type RangeMapper = (range: CellRange) => CellRange | null;

/** 挿入・削除を CellRange 全体に適用するマッパーを作る */
function makeMapper(
  axis: Axis,
  op: "insert" | "delete",
  pos: number,
  count: number,
): RangeMapper {
  return (range: CellRange): CellRange | null => {
    const [start, end] = axis === "row"
      ? [range.startRow, range.endRow]
      : [range.startCol, range.endCol];
    const mapped = op === "insert"
      ? mapInsert(start, end, pos, count)
      : mapDelete(start, end, pos, count);
    if (mapped === null) return null;
    const [newStart, newEnd] = mapped;
    return axis === "row"
      ? { ...range, startRow: newStart, endRow: newEnd }
      : { ...range, startCol: newStart, endCol: newEnd };
  };
}

// ---------------------------------------------------------------------------
// Merge / data-validation preservation across splices
// ---------------------------------------------------------------------------

interface WsInternals {
  _merges?: Record<string, { tl: string; br: string }>;
  mergeCellsWithoutStyle?: (range: string) => void;
  dataValidations?: { model: Record<string, unknown> };
}

function getMergeRanges(ws: ExcelJS.Worksheet): CellRange[] {
  const merges = (ws as unknown as WsInternals)._merges;
  if (!merges) return [];
  const ranges: CellRange[] = [];
  for (const dim of Object.values(merges)) {
    if (dim && dim.tl && dim.br) {
      ranges.push(parseRange(`${dim.tl}:${dim.br}`));
    }
  }
  return ranges;
}

function remerge(ws: ExcelJS.Worksheet, range: CellRange): void {
  const rangeStr = rangeToString(range);
  const internal = ws as unknown as WsInternals;
  try {
    if (typeof internal.mergeCellsWithoutStyle === "function") {
      internal.mergeCellsWithoutStyle(rangeStr);
    } else {
      ws.mergeCells(rangeStr);
    }
  } catch {
    // 再結合に失敗した結合は諦める（splice 自体は成功している）
  }
}

/**
 * データ検証の辞書キー（"A1" / "A1:B5"、空白区切りの複合参照も可）を
 * mapper でシフトした新しい辞書を返す。
 */
function shiftDataValidations(
  model: Record<string, unknown>,
  mapper: RangeMapper,
): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const [key, dv] of Object.entries(model)) {
    if (dv === undefined || dv === null) continue;
    const parts: string[] = [];
    for (const ref of key.split(/\s+/).filter(Boolean)) {
      let mapped: CellRange | null;
      try {
        mapped = mapper(parseRange(ref));
      } catch {
        parts.push(ref); // 解析できないキーはそのまま維持
        continue;
      }
      if (mapped) parts.push(rangeToString(mapped));
    }
    if (parts.length > 0) {
      result[parts.join(" ")] = dv;
    }
  }
  return result;
}

/**
 * splice 前後で結合セルとデータ検証を保全しながら mutate を実行する。
 *
 * 1. 影響を受ける結合（axis 上で pos 以降に達するもの）を解除
 * 2. mutate（spliceRows / spliceColumns）
 * 3. シフト済み座標で再結合（1×1 に縮退した結合は破棄）
 * 4. データ検証キーをシフトした辞書に差し替え
 */
function spliceWithLayoutPreserved(
  ws: ExcelJS.Worksheet,
  axis: Axis,
  op: "insert" | "delete",
  pos: number,
  count: number,
  mutate: () => void,
): void {
  const mapper = makeMapper(axis, op, pos, count);

  // 共有数式グループを通常の数式に実体化する。ExcelJS の splice は
  // スレーブの sharedFormula ポインタをシフトしないため、放置すると
  // 保存時に "Shared Formula master must exist above..." で失敗する。
  materializeSharedFormulas(ws);

  const allMerges = getMergeRanges(ws);
  const affected = allMerges.filter((r) =>
    axis === "row" ? r.endRow >= pos : r.endCol >= pos,
  );
  for (const r of affected) {
    try {
      ws.unMergeCells(rangeToString(r));
    } catch {
      // 既に解除済みなど — 続行
    }
  }

  mutate();

  for (const r of affected) {
    const mapped = mapper(r);
    if (!mapped) continue;
    // 1 セルに縮退した結合は意味がないので破棄
    if (mapped.startRow === mapped.endRow && mapped.startCol === mapped.endCol) continue;
    remerge(ws, mapped);
  }

  const dvHolder = (ws as unknown as WsInternals).dataValidations;
  if (dvHolder && dvHolder.model && Object.keys(dvHolder.model).length > 0) {
    dvHolder.model = shiftDataValidations(dvHolder.model, mapper);
  }
}

// ---------------------------------------------------------------------------
// Public operations
// ---------------------------------------------------------------------------

/**
 * 指定位置に行を挿入する。
 */
export function insertRowsAt(
  ws: ExcelJS.Worksheet,
  row: number,
  count: number,
): void {
  validateRowBounds(row, count);
  spliceWithLayoutPreserved(ws, "row", "insert", row, count, () => {
    ws.spliceRows(row, 0, ...Array(count).fill([]));
  });
}

/**
 * 指定位置の行を削除する。
 */
export function deleteRowsAt(
  ws: ExcelJS.Worksheet,
  row: number,
  count: number,
): void {
  validateRowBounds(row, count);
  spliceWithLayoutPreserved(ws, "row", "delete", row, count, () => {
    ws.spliceRows(row, count);
  });
}

/**
 * 指定位置に列を挿入する。
 */
export function insertColumnsAt(
  ws: ExcelJS.Worksheet,
  col: number,
  count: number,
): void {
  validateColumnBounds(col, count);
  spliceWithLayoutPreserved(ws, "col", "insert", col, count, () => {
    ws.spliceColumns(col, 0, ...Array(count).fill([]));
  });
}

/**
 * 指定位置の列を削除する。
 */
export function deleteColumnsAt(
  ws: ExcelJS.Worksheet,
  col: number,
  count: number,
): void {
  validateColumnBounds(col, count);
  spliceWithLayoutPreserved(ws, "col", "delete", col, count, () => {
    ws.spliceColumns(col, count);
  });
}
