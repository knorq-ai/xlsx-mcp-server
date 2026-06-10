/**
 * Safety-guard regression tests:
 * 上書き防止・出力上限・書式の共有切断・ファイルロック・エラー伝播など、
 * データを壊さないためのガードが効いていることを検証する。
 */

import { describe, it, expect, afterEach } from "vitest";
import * as fs from "fs/promises";
import ExcelJS from "exceljs";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  tmpXlsxPath,
  trackTmpFile,
} from "./helpers.js";
import {
  createWorkbook,
  writeCell,
  writeRows,
  readCell,
  readSheet,
  searchCells,
  formatCells,
  EngineError,
  ErrorCode,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

// ---------------------------------------------------------------------------
// 1. Atomic save — .xlsm 書き込み拒否（元ファイルは無傷）
// ---------------------------------------------------------------------------

describe("macro-enabled workbook write guard", () => {
  it("rejects writes to .xlsm with INVALID_PARAMETER and leaves the file byte-identical", async () => {
    // .xlsx を作って .xlsm にリネームし、疑似マクロブックを用意する
    const xlsxPath = await createTmpWorkbook();
    const xlsmPath = xlsxPath.replace(/\.xlsx$/, ".xlsm");
    trackTmpFile(xlsmPath);
    await fs.rename(xlsxPath, xlsmPath);
    const before = await fs.readFile(xlsmPath);

    await expect(writeCell(xlsmPath, 1, "A1", "boom")).rejects.toMatchObject({
      name: "EngineError",
      code: ErrorCode.INVALID_PARAMETER,
    });

    // 保存パスに入る前に拒否されるため、ファイルは 1 バイトも変わらない
    const after = await fs.readFile(xlsmPath);
    expect(after.equals(before)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 2. create_workbook — 既存ファイルの上書き防止
// ---------------------------------------------------------------------------

describe("create_workbook overwrite protection", () => {
  it("throws INVALID_PARAMETER when the path already exists", async () => {
    const p = await createTmpWorkbook();
    let err: unknown;
    try {
      await createWorkbook(p);
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
    expect((err as EngineError).message).toContain("already exists");
  });
});

// ---------------------------------------------------------------------------
// 3. read_sheet — 5,000 セル上限での行単位打ち切り
// ---------------------------------------------------------------------------

describe("read_sheet output truncation", () => {
  it("truncates 600×10 cells at row 500 and reports it", async () => {
    const p = await createTmpWorkbook();
    const rows = Array.from({ length: 600 }, (_, r) =>
      Array.from({ length: 10 }, (_, c) => r * 10 + c + 1),
    );
    await writeRows(p, 1, 1, rows);

    const result = await readSheet(p, 1);
    expect(result).toContain("⚠ Output truncated");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.truncated).toBe(true);
    expect(json.truncatedAtRow).toBe(500);
    // 500 行 × 10 列 = 上限ちょうどの 5,000 セルが返る
    expect(Object.keys(json.cells).length).toBe(5000);
  });
});

// ---------------------------------------------------------------------------
// 4. search_cells — max_results での打ち切り
// ---------------------------------------------------------------------------

describe("search_cells max_results cap", () => {
  it("returns exactly max_results matches and flags truncation", async () => {
    const p = await createTmpWorkbook();
    const rows = Array.from({ length: 60 }, (_, r) => [`needle-${r + 1}`]);
    await writeRows(p, 1, 1, rows);

    const result = await searchCells(p, "needle", 1, false, 50);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.matches.length).toBe(50);
    expect(json.truncated).toBe(true);
    expect(result).toContain("Found 50+ match(es)");
  });
});

// ---------------------------------------------------------------------------
// 5. compact 読み取り — 結果未計算の数式セルを省略しない
// ---------------------------------------------------------------------------

describe("uncached formula cells in compact read", () => {
  it("keeps a freshly written formula in the formulas map without v", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", 21);
    await writeCell(p, 1, "B1", "=A1*2");

    const result = await readSheet(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    // このサーバは数式を計算しないため結果キャッシュは無いが、
    // 省略すると LLM が空セルと誤認して上書きする — 必ず返す
    expect(json.formulas.B1.f).toBe("A1*2");
    expect("v" in json.formulas.B1).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// 6. リテラルエスケープ — "'=" は数式ではなく文字列
// ---------------------------------------------------------------------------

describe("literal '=' escape", () => {
  it(`stores "'=not a formula" as the string "=not a formula"`, async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "'=not a formula");

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("=not a formula");
    expect(json.type).toBe("string");
    expect(json.formula).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// 7. 書式の共有切断 — format_cells が隣のセルに波及しない
// ---------------------------------------------------------------------------

describe("style sharing isolation", () => {
  it("formatting one cell does not bleed into another cell loaded with the same style", async () => {
    // ExcelJS はファイル読み込み時に同一書式のセル間で style オブジェクトを
    // 共有する。同じ style オブジェクトを 2 セルに与えて保存・再読込し、
    // 共有状態を再現する。
    const p = tmpXlsxPath();
    trackTmpFile(p);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    const sharedStyle: Partial<ExcelJS.Style> = {
      font: { bold: true, size: 11, name: "Arial" },
    };
    const a1 = ws.getCell("A1");
    a1.value = "left";
    a1.style = sharedStyle;
    const b1 = ws.getCell("B1");
    b1.value = "right";
    b1.style = sharedStyle;
    await wb.xlsx.writeFile(p);

    // A1 だけに書式変更を適用する
    await formatCells(p, 1, "A1:A1", { fillColor: "FF0000", italic: true });

    // B1 は元の書式のまま（共有 style 経由で汚染されない）
    const bResult = await readCell(p, 1, "B1");
    const bJson = JSON.parse(bResult.split("<json>")[1].split("</json>")[0]);
    expect(bJson.style.bold).toBe(true);
    expect(bJson.style.fillColor).toBeUndefined();
    expect(bJson.style.italic).toBeUndefined();

    // A1 には適用されている
    const aResult = await readCell(p, 1, "A1");
    const aJson = JSON.parse(aResult.split("<json>")[1].split("</json>")[0]);
    expect(aJson.style.fillColor).toBe("FF0000");
    expect(aJson.style.italic).toBe(true);
    expect(aJson.style.bold).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 8. fillPattern: "solid" 単独指定 — 既存の塗り色を白で潰さない
// ---------------------------------------------------------------------------

describe("fillPattern solid without fillColor", () => {
  it("preserves the existing fill color", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "colored");
    await formatCells(p, 1, "A1:A1", { fillColor: "00FF00" });

    // fillPattern のみ指定 — 既存の緑が保持されること
    await formatCells(p, 1, "A1:A1", { fillPattern: "solid" });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.fillColor).toBe("00FF00");
  });
});

// ---------------------------------------------------------------------------
// 9. ファイルロック — 死んだ PID の stale ロックを奪取する
// ---------------------------------------------------------------------------

describe("stale cross-process lock", () => {
  it("steals a .mcplock owned by a dead PID and removes the lock afterwards", async () => {
    const p = await createTmpWorkbook();
    // ロックキーは realpath で正規化される（macOS の /var → /private/var 等）
    const lockPath = (await fs.realpath(p)) + ".mcplock";
    trackTmpFile(lockPath);
    // macOS の PID 上限 (99999) を超える番号 = 確実に存在しないプロセス
    await fs.writeFile(lockPath, "999999");

    const started = Date.now();
    const msg = await writeCell(p, 1, "A1", "stolen");
    expect(msg).toContain("Set A1");
    // タイムアウト (10s) を待たず、即座に奪取できること
    expect(Date.now() - started).toBeLessThan(5000);

    // 書き込み完了後はロックファイルが解放（削除）されている
    await expect(fs.stat(lockPath)).rejects.toThrow();
  }, 15000);
});

// ---------------------------------------------------------------------------
// 10. resolveSheet のエラーメッセージ — 利用可能シートの列挙とヒント
// ---------------------------------------------------------------------------

describe("sheet resolution error messages", () => {
  it("lists available sheet names when the sheet is not found", async () => {
    const p = await createTmpWorkbook("DataSheet");
    let err: unknown;
    try {
      await readSheet(p, "Bogus");
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.SHEET_NOT_FOUND);
    expect((err as EngineError).message).toContain('"DataSheet"');
  });

  it("hints to pass a JSON number when an absent digit-string sheet name is given", async () => {
    const p = await createTmpWorkbook(); // "Sheet1" のみ — シート "2" は存在しない
    let err: unknown;
    try {
      await readSheet(p, "2");
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).message).toContain(
      "pass the number 2 as a JSON number, not a string",
    );
    expect((err as EngineError).message).toContain('"Sheet1"');
  });
});

// ---------------------------------------------------------------------------
// 11. エラー値 — 数式のキャッシュ済みエラー結果とプレーンエラーセル
// ---------------------------------------------------------------------------

describe("error-value cells", () => {
  it("reads back a formula's cached #DIV/0! result and a plain error cell", async () => {
    const p = tmpXlsxPath();
    trackTmpFile(p);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "1/0",
      result: { error: "#DIV/0!" },
    } as ExcelJS.CellFormulaValue;
    ws.getCell("B1").value = { error: "#N/A" } as ExcelJS.CellErrorValue;
    await wb.xlsx.writeFile(p);

    const result = await readSheet(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    // 現在の仕様：数式セルのエラー結果は formulas マップの v に入る。
    // errors マップは数式を持たないプレーンなエラーセル専用である。
    expect(json.formulas.A1).toEqual({ f: "1/0", v: "#DIV/0!" });
    expect(json.errors.B1).toBe("#N/A");

    // read_cell でもエラー文字列が value として返る
    const cellResult = await readCell(p, 1, "A1");
    const cellJson = JSON.parse(cellResult.split("<json>")[1].split("</json>")[0]);
    expect(cellJson.value).toBe("#DIV/0!");
    expect(cellJson.type).toBe("formula");
  });
});
