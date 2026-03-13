/**
 * View settings — freeze panes, auto filter.
 */

import ExcelJS from "exceljs";

/**
 * フリーズペインを設定する。
 * row: 固定する行数（0 で解除）
 * column: 固定する列数（0 で解除）
 */
export function setFreezePanes(
  ws: ExcelJS.Worksheet,
  row: number,
  column: number,
): void {
  if (row === 0 && column === 0) {
    // 解除: normal view
    ws.views = [{ state: "normal" }];
    return;
  }

  ws.views = [
    {
      state: "frozen",
      xSplit: column,
      ySplit: row,
      topLeftCell: undefined,
    },
  ];
}

/**
 * オートフィルタを設定する。
 */
export function setAutoFilter(
  ws: ExcelJS.Worksheet,
  range: string,
): void {
  ws.autoFilter = range;
}

/**
 * オートフィルタを解除する。
 *
 * ExcelJS の型定義は autoFilter を string として宣言しているが、
 * 実行時に undefined を代入するとフィルタが解除される。
 * TypeScript の型チェックを通すために二重キャストを使用する。
 */
export function removeAutoFilter(ws: ExcelJS.Worksheet): void {
  ws.autoFilter = undefined as unknown as string;
}
