/**
 * Sheet operations — add, rename, delete, copy.
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";

/** Excel のシート名の最大長 */
const MAX_SHEET_NAME_LENGTH = 31;
/** Excel のシート名で使用できない文字 */
const INVALID_SHEET_NAME_CHARS = /[*?:\\/[\]]/;

/**
 * シート名を検証する。
 * ExcelJS は 31 文字超を警告のみで受け入れ（Excel が開けないファイルになる）、
 * 不正文字では素の Error を投げるため、事前に EngineError で弾く。
 */
export function validateSheetName(name: string): void {
  if (name.length === 0) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, "Sheet name must not be empty");
  }
  if (name.length > MAX_SHEET_NAME_LENGTH) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Sheet name "${name}" is ${name.length} characters. Excel allows at most ${MAX_SHEET_NAME_LENGTH}.`,
    );
  }
  if (INVALID_SHEET_NAME_CHARS.test(name)) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Sheet name "${name}" contains invalid characters. Excel forbids: * ? : \\ / [ ]`,
    );
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Sheet name "${name}" must not start or end with an apostrophe.`,
    );
  }
}

/**
 * ワークシートを追加する。
 */
export function addWorksheet(
  workbook: ExcelJS.Workbook,
  name: string,
): ExcelJS.Worksheet {
  validateSheetName(name);
  // 同名チェック
  if (workbook.getWorksheet(name)) {
    throw new EngineError(ErrorCode.DUPLICATE_NAME, `Sheet already exists: "${name}"`);
  }
  return workbook.addWorksheet(name);
}

/**
 * ワークシートの名前を変更する。
 */
export function renameWorksheet(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  newName: string,
): void {
  validateSheetName(newName);
  if (workbook.getWorksheet(newName)) {
    throw new EngineError(ErrorCode.DUPLICATE_NAME, `Sheet already exists: "${newName}"`);
  }
  ws.name = newName;
}

/**
 * ワークシートを削除する。
 */
export function deleteWorksheet(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
): void {
  workbook.removeWorksheet(ws.id);
}

/**
 * ワークシートをコピーする。
 * ExcelJS には直接コピー API がないため、セル値と書式を手動でコピーする。
 */
export function copyWorksheet(
  workbook: ExcelJS.Workbook,
  source: ExcelJS.Worksheet,
  newName: string,
): ExcelJS.Worksheet {
  validateSheetName(newName);
  if (workbook.getWorksheet(newName)) {
    throw new EngineError(ErrorCode.DUPLICATE_NAME, `Sheet already exists: "${newName}"`);
  }

  const dest = workbook.addWorksheet(newName);

  // Copy column properties
  source.columns?.forEach((col, i) => {
    if (col.width) {
      dest.getColumn(i + 1).width = col.width;
    }
  });

  // Copy rows
  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const destRow = dest.getRow(rowNumber);
    destRow.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const destCell = destRow.getCell(colNumber);
      destCell.value = cell.value;
      // 深いコピー：シャローコピーだと font/border 等のネストオブジェクトを
      // コピー元シートと共有してしまい、後の書式変更が双方に波及する
      destCell.style = structuredClone(cell.style);
    });
    destRow.commit();
  });

  // Copy merged cells
  // Access through model since mergeCells is the only public API for merges.
  // mergeCells はマスターの style を結合範囲全体へ複製してセル個別の書式を
  // 壊すため、style に触れない内部 API を優先する。
  const merges = source.model?.merges;
  if (merges) {
    const destInternal = dest as unknown as {
      mergeCellsWithoutStyle?: (range: string) => void;
    };
    for (const merge of merges) {
      if (typeof destInternal.mergeCellsWithoutStyle === "function") {
        destInternal.mergeCellsWithoutStyle(merge);
      } else {
        dest.mergeCells(merge);
      }
    }
  }

  return dest;
}
