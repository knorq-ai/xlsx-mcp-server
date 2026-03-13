/**
 * Named ranges management.
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";

export interface NamedRangeInfo {
  name: string;
  range: string;
}

/**
 * ワークブックの名前付き範囲一覧を返す。
 */
export function listNamedRanges(workbook: ExcelJS.Workbook): NamedRangeInfo[] {
  const ranges: NamedRangeInfo[] = [];

  const definedNames = workbook.definedNames?.model;
  if (definedNames && Array.isArray(definedNames)) {
    for (const dn of definedNames) {
      ranges.push({
        name: dn.name,
        range: dn.ranges?.join(", ") ?? "",
      });
    }
  }

  return ranges;
}

/**
 * 名前付き範囲を追加する。
 */
export function addNamedRange(
  workbook: ExcelJS.Workbook,
  name: string,
  range: string,
  sheetName?: string,
): void {
  // Excel の規約に従い、シート名中の単一引用符はエスケープ（二重化）する
  const fullRange = sheetName
    ? `'${sheetName.replace(/'/g, "''")}'!${range}`
    : range;

  // Check for duplicates
  const existing = listNamedRanges(workbook);
  if (existing.some((r) => r.name === name)) {
    throw new EngineError(ErrorCode.DUPLICATE_NAME, `Named range already exists: "${name}"`);
  }

  // ExcelJS definedNames API: add(locStr, name)
  workbook.definedNames.add(fullRange, name);
}

/**
 * 名前付き範囲を削除する。
 *
 * ExcelJS の remove(locStr, name) は個別セル単位でしか動作しないため、
 * model を直接フィルタリングして削除する。
 */
export function deleteNamedRange(
  workbook: ExcelJS.Workbook,
  name: string,
): void {
  const existing = listNamedRanges(workbook);
  if (!existing.some((r) => r.name === name)) {
    throw new EngineError(ErrorCode.NAMED_RANGE_NOT_FOUND, `Named range not found: "${name}"`);
  }

  workbook.definedNames.model = workbook.definedNames.model.filter(
    (m) => m.name !== name,
  );
}
