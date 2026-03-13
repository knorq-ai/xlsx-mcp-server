/**
 * Row and column operations — insert, delete.
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";

/** Excel の最大行数 */
const MAX_ROWS = 1_048_576;
/** Excel の最大列数 */
const MAX_COLS = 16_384;

function validateRowBounds(row: number, count: number): void {
  if (row < 1 || row > MAX_ROWS) {
    throw new EngineError(ErrorCode.ROW_OUT_OF_RANGE, `Row ${row} out of range (1-${MAX_ROWS})`);
  }
  if (count < 1) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `Count must be at least 1`);
  }
}

function validateColumnBounds(col: number, count: number): void {
  if (col < 1 || col > MAX_COLS) {
    throw new EngineError(ErrorCode.COLUMN_OUT_OF_RANGE, `Column ${col} out of range (1-${MAX_COLS})`);
  }
  if (count < 1) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `Count must be at least 1`);
  }
}

/**
 * 指定位置に行を挿入する。
 * ExcelJS の spliceRows を使用。
 */
export function insertRowsAt(
  ws: ExcelJS.Worksheet,
  row: number,
  count: number,
): void {
  validateRowBounds(row, count);
  ws.spliceRows(row, 0, ...Array(count).fill([]));
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
  ws.spliceRows(row, count);
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
  ws.spliceColumns(col, 0, ...Array(count).fill([]));
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
  ws.spliceColumns(col, count);
}
