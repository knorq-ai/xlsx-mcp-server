/**
 * Data validation rules.
 *
 * ExcelJS の TypeScript 型定義はワークシートレベルの dataValidations を公開しない。
 * そのため範囲内の各セルに個別に dataValidation を設定する。
 * 大きな範囲（数万セル以上）ではセル数に比例した処理時間がかかる。
 */

import ExcelJS from "exceljs";
import { ErrorCode, EngineError } from "./xlsx-io.js";
import { parseRange, validateRangeSize } from "./cells.js";

export interface DataValidationParams {
  type: "list" | "whole" | "decimal" | "date" | "textLength" | "custom";
  formulae: string[];
  operator?: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  allowBlank?: boolean;
}

/**
 * 指定範囲にデータ検証ルールを追加する。
 */
export function addDataValidationRule(
  ws: ExcelJS.Worksheet,
  range: string,
  params: DataValidationParams,
): void {
  if (!params.formulae || params.formulae.length === 0) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, "Data validation requires at least one formula");
  }

  const dv: ExcelJS.DataValidation = {
    type: params.type,
    formulae: params.formulae,
    allowBlank: params.allowBlank ?? true,
  };

  if (params.operator) dv.operator = params.operator;
  if (params.showErrorMessage !== undefined) dv.showErrorMessage = params.showErrorMessage;
  if (params.errorTitle) dv.errorTitle = params.errorTitle;
  if (params.error) dv.error = params.error;
  if (params.showInputMessage !== undefined) dv.showInputMessage = params.showInputMessage;
  if (params.promptTitle) dv.promptTitle = params.promptTitle;
  if (params.prompt) dv.prompt = params.prompt;

  const parsed = parseRange(range);
  validateRangeSize(parsed);
  for (let r = parsed.startRow; r <= parsed.endRow; r++) {
    for (let c = parsed.startCol; c <= parsed.endCol; c++) {
      ws.getRow(r).getCell(c).dataValidation = dv;
    }
  }
}

/**
 * 指定範囲のデータ検証ルールを削除する。
 */
export function removeDataValidationRule(
  ws: ExcelJS.Worksheet,
  range: string,
): void {
  const parsed = parseRange(range);
  validateRangeSize(parsed);
  for (let r = parsed.startRow; r <= parsed.endRow; r++) {
    for (let c = parsed.startCol; c <= parsed.endCol; c++) {
      const cell = ws.getRow(r).getCell(c);
      // ExcelJS の型定義は dataValidation に undefined を受け付けないが、
      // 実行時に undefined を代入すると検証ルールが解除される
      (cell as unknown as Record<string, unknown>).dataValidation = undefined;
    }
  }
}
