/**
 * Sheet operations — add, rename, delete, copy.
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";

/**
 * ワークシートを追加する。
 */
export function addWorksheet(
  workbook: ExcelJS.Workbook,
  name: string,
): ExcelJS.Worksheet {
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
      destCell.style = { ...cell.style };
    });
    destRow.commit();
  });

  // Copy merged cells
  // Access through model since mergeCells is the only public API for merges
  const merges = source.model?.merges;
  if (merges) {
    for (const merge of merges) {
      dest.mergeCells(merge);
    }
  }

  return dest;
}
