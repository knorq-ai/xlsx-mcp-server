/**
 * 構造変更（行・列の挿入/削除、共有数式、上限・名前検証）の整合性テスト。
 *
 * - 結合セル・データ検証が splice 操作でシフトされることを確認する
 * - 共有数式グループが構造変更・マスター上書きで実体化されることを確認する
 * - Excel の行・列上限、シート名検証、全列範囲の拒否を確認する
 */

import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  tmpXlsxPath,
  trackTmpFile,
  EngineError,
  ErrorCode,
} from "./helpers.js";
import type { ErrorCodeType } from "./helpers.js";
import {
  writeCell,
  readSheet,
  mergeCells,
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  addDataValidation,
  listDataValidations,
  clearCells,
  addSheet,
  renameSheet,
} from "../xlsx-engine.js";
import ExcelJS from "exceljs";

afterEach(cleanupTmpFiles);

/** <json> ブロックをパースする */
function parseJson(result: string): Record<string, any> {
  return JSON.parse(result.split("<json>")[1].split("</json>")[0]);
}

/** Promise が指定コードの EngineError で reject することを検証する */
async function expectEngineError(
  promise: Promise<unknown>,
  code: ErrorCodeType,
  messagePart: string,
): Promise<void> {
  const err = await promise.then(
    () => null,
    (e: unknown) => e,
  );
  expect(err).toBeInstanceOf(EngineError);
  expect((err as EngineError).code).toBe(code);
  expect((err as EngineError).message).toContain(messagePart);
}

/**
 * 共有数式グループ G2:H2 を持つワークブックを raw ExcelJS で構築する。
 * マスター G2 = `$C2*D2`、スレーブ H2 は相対参照をずらした `$C2*E2` に相当する。
 */
async function createSharedFormulaWorkbook(): Promise<string> {
  const p = tmpXlsxPath();
  trackTmpFile(p);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Shared");
  ws.getCell("A1").value = "header";
  ws.getCell("C2").value = 770000;
  ws.getCell("D2").value = 0.1;
  ws.getCell("E2").value = 0.9;
  ws.getCell("G2").value = {
    formula: "$C2*D2",
    result: 77000,
    shareType: "shared",
    ref: "G2:H2",
  } as ExcelJS.CellFormulaValue;
  ws.getCell("H2").value = {
    sharedFormula: "G2",
    result: 693000,
  } as ExcelJS.CellSharedFormulaValue;
  await wb.xlsx.writeFile(p);
  return p;
}

// ---------------------------------------------------------------------------
// 1-2. 行の挿入・削除と結合セルのシフト
// ---------------------------------------------------------------------------

describe("merge shifting on insert_rows / delete_rows", () => {
  it("insert_rows above a merge shifts the merge down", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B5", "merged!");
    await mergeCells(p, 1, "B5:C6");

    await insertRows(p, 1, 2, 2);

    const json = parseJson(await readSheet(p, 1));
    expect(json.mergedCells).toEqual(["B7:C8"]);
    // マスター値も結合と一緒に移動する
    expect(json.cells["B7"]).toBe("merged!");
    expect(json.cells["B5"]).toBeUndefined();
  });

  it("delete_rows above a merge shifts the merge back up", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B5", "merged!");
    await mergeCells(p, 1, "B5:C6");

    await deleteRows(p, 1, 2, 2);

    const json = parseJson(await readSheet(p, 1));
    expect(json.mergedCells).toEqual(["B3:C4"]);
    expect(json.cells["B3"]).toBe("merged!");
  });

  it("deleting a row inside a merge shrinks the merge", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B5", "tall");
    await mergeCells(p, 1, "B5:C7");

    // 結合範囲の中間行（6 行目）を削除 → 結合は 1 行分縮む
    await deleteRows(p, 1, 6, 1);

    const json = parseJson(await readSheet(p, 1));
    expect(json.mergedCells).toEqual(["B5:C6"]);
    expect(json.cells["B5"]).toBe("tall");
  });
});

// ---------------------------------------------------------------------------
// 3. 列の挿入・削除と結合セルの水平シフト
// ---------------------------------------------------------------------------

describe("merge shifting on insert_columns / delete_columns", () => {
  it("insert_columns left of a merge shifts the merge right", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B5", "wide");
    await mergeCells(p, 1, "B5:C6");

    await insertColumns(p, 1, "A", 2);

    const json = parseJson(await readSheet(p, 1));
    expect(json.mergedCells).toEqual(["D5:E6"]);
    expect(json.cells["D5"]).toBe("wide");
  });

  it("delete_columns left of a merge shifts the merge left", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "C5", "wide");
    await mergeCells(p, 1, "C5:D6");

    await deleteColumns(p, 1, "A", 1);

    const json = parseJson(await readSheet(p, 1));
    expect(json.mergedCells).toEqual(["B5:C6"]);
    expect(json.cells["B5"]).toBe("wide");
  });
});

// ---------------------------------------------------------------------------
// 4. データ検証のシフト
// ---------------------------------------------------------------------------

describe("data validation shifting on insert_rows / delete_rows", () => {
  it("shifts a validation down on insert and back up on delete", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "D5", {
      type: "list",
      formulae: ['"Yes,No"'],
    });

    await insertRows(p, 1, 2, 2);

    let json = parseJson(await listDataValidations(p, 1));
    expect(json.validations.map((v: { address: string }) => v.address)).toEqual(["D7"]);
    expect(json.validations[0].type).toBe("list");
    expect(json.validations[0].formulae).toEqual(['"Yes,No"']);

    await deleteRows(p, 1, 2, 2);

    json = parseJson(await listDataValidations(p, 1));
    expect(json.validations.map((v: { address: string }) => v.address)).toEqual(["D5"]);
  });
});

// ---------------------------------------------------------------------------
// 5. 共有数式グループと行の挿入・削除
// ---------------------------------------------------------------------------

describe("shared formulas across insert_rows / delete_rows", () => {
  it("insert_rows on a sheet with a shared-formula group saves and keeps per-cell formulas", async () => {
    const p = await createSharedFormulaWorkbook();

    // ExcelJS の splice はスレーブの sharedFormula ポインタをシフトしない。
    // 実体化せずに保存すると "Shared Formula master must exist above..." で
    // 失敗するため、ここでは throw せず保存できることを確認する。
    await expect(insertRows(p, 1, 1, 1)).resolves.toContain("Inserted");

    const json = parseJson(await readSheet(p, 1));
    // セルは 1 行下にシフトする。数式参照は仕様上更新されない
    // （CLAUDE.md「数式参照の自動更新」の制限事項どおり）。
    expect(json.formulas["G3"].f).toBe("$C2*D2");
    expect(json.formulas["H3"].f).toBe("$C2*E2");
    // 劣化した共有数式グループ（マスター参照のみ）が残っていないこと
    expect(json.sharedGroupMasters).toBeUndefined();
  });

  it("delete_rows on a sheet with a shared-formula group saves and keeps per-cell formulas", async () => {
    const p = await createSharedFormulaWorkbook();

    await expect(deleteRows(p, 1, 1, 1)).resolves.toContain("Deleted");

    const json = parseJson(await readSheet(p, 1));
    expect(json.formulas["G1"].f).toBe("$C2*D2");
    expect(json.formulas["H1"].f).toBe("$C2*E2");
    expect(json.sharedGroupMasters).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// 6. 共有数式マスターの上書き
// ---------------------------------------------------------------------------

describe("write_cell over a shared-formula master", () => {
  it("succeeds and detaches slaves into independent formulas", async () => {
    const p = await createSharedFormulaWorkbook();

    await expect(writeCell(p, 1, "G2", "plain")).resolves.toContain("Set G2");

    const json = parseJson(await readSheet(p, 1));
    expect(json.cells["G2"]).toBe("plain");
    // スレーブはセル固有の数式に変換され、キャッシュ済み結果を保持する
    expect(json.formulas["H2"].f).toBe("$C2*E2");
    expect(json.formulas["H2"].v).toBe(693000);
    expect(json.sharedGroupMasters).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// 7. Excel の行・列上限
// ---------------------------------------------------------------------------

describe("Excel bounds validation", () => {
  it("rejects write_cell beyond the maximum row 1,048,576", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      writeCell(p, 1, "A1048577", 1),
      ErrorCode.ROW_OUT_OF_RANGE,
      "exceeds Excel's maximum row",
    );
  });

  it("rejects write_cell beyond the maximum column XFD", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      writeCell(p, 1, "XFE1", 1),
      ErrorCode.COLUMN_OUT_OF_RANGE,
      "exceeds Excel's maximum column XFD",
    );
  });

  it("rejects insert_rows with count exceeding the row limit", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      insertRows(p, 1, 1, 1_048_577),
      ErrorCode.INVALID_PARAMETER,
      "exceeds Excel's row limit",
    );
  });
});

// ---------------------------------------------------------------------------
// 8. シート名の検証
// ---------------------------------------------------------------------------

describe("sheet name validation", () => {
  it("add_sheet rejects a 32-character name with an actionable message", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      addSheet(p, "S".repeat(32)),
      ErrorCode.INVALID_PARAMETER,
      "Excel allows at most 31",
    );
  });

  it("add_sheet rejects a name with forbidden characters", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      addSheet(p, "bad:name"),
      ErrorCode.INVALID_PARAMETER,
      "invalid characters",
    );
  });

  it("add_sheet rejects a name with a leading apostrophe", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      addSheet(p, "'temp"),
      ErrorCode.INVALID_PARAMETER,
      "apostrophe",
    );
  });

  it("rename_sheet applies the same validation", async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      renameSheet(p, 1, "R".repeat(32)),
      ErrorCode.INVALID_PARAMETER,
      "Excel allows at most 31",
    );
    await expectEngineError(
      renameSheet(p, 1, "bad/name"),
      ErrorCode.INVALID_PARAMETER,
      "invalid characters",
    );
    await expectEngineError(
      renameSheet(p, 1, "'temp"),
      ErrorCode.INVALID_PARAMETER,
      "apostrophe",
    );
  });
});

// ---------------------------------------------------------------------------
// 9. 全列・全行範囲の拒否
// ---------------------------------------------------------------------------

describe("whole-column / whole-row range rejection", () => {
  it('clear_cells rejects "A:A" with an actionable message', async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      clearCells(p, 1, "A:A"),
      ErrorCode.INVALID_RANGE,
      "Whole-column/whole-row ranges",
    );
    // 代替の書き方（明示的なセル範囲）を案内していること
    const err = await clearCells(p, 1, "A:A").then(
      () => null,
      (e: unknown) => e,
    );
    expect((err as EngineError).message).toContain('"A1:A100"');
  });

  it('read_sheet rejects whole-row range "1:3"', async () => {
    const p = await createTmpWorkbook();
    await expectEngineError(
      readSheet(p, 1, "1:3"),
      ErrorCode.INVALID_RANGE,
      "Whole-column/whole-row ranges",
    );
  });
});
