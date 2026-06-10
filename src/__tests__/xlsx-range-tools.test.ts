/**
 * 範囲操作ツール（copy_range / find_replace / sort_range / clear_cells モード）の
 * エンジンレベルテスト。
 */

import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  EngineError,
  ErrorCode,
} from "./helpers.js";
import {
  writeCell,
  writeCells,
  writeRows,
  readSheet,
  readCell,
  formatCells,
  mergeCells,
  addSheet,
  copyRange,
  findReplace,
  sortRange,
  clearCells,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

/** read_sheet の <json> ブロックをパースする */
async function readJson(
  p: string,
  sheet: string | number,
  range?: string,
  includeStyles?: boolean,
) {
  const r = await readSheet(p, sheet, range, includeStyles);
  return JSON.parse(r.split("<json>")[1].split("</json>")[0]);
}

// ---------------------------------------------------------------------------
// copy_range
// ---------------------------------------------------------------------------

describe("copy_range", () => {
  it("copies values and styles to the destination", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Name", "Score"],
      ["alice", 90],
    ]);
    // ヘッダ行を太字にしてからコピーする
    await formatCells(p, 1, "A1:B1", { bold: true });

    const msg = await copyRange(p, 1, "A1:B2", "D5");
    expect(msg).toContain("Copied A1:B2 → D5:E6");
    expect(msg).toContain("4 cells");

    const json = await readJson(p, 1, undefined, true);
    // 値がコピーされている
    expect(json.cells.D5).toBe("Name");
    expect(json.cells.E5).toBe("Score");
    expect(json.cells.D6).toBe("alice");
    expect(json.cells.E6).toBe(90);
    // 書式（太字）も一緒にコピーされている
    expect(json.styles.D5.bold).toBe(true);
    expect(json.styles.E5.bold).toBe(true);
    expect(json.styles?.D6?.bold).toBeUndefined();
    // コピー元は変更されない
    expect(json.cells.A1).toBe("Name");
    expect(json.styles.A1.bold).toBe(true);
  });

  it("translates relative formula references; $-anchored refs stay fixed", async () => {
    const p = await createTmpWorkbook();
    await writeCells(p, 1, [
      { cell: "A4", value: 10 },
      { cell: "B4", value: 20 },
      { cell: "C2", value: "=A4+B4" },
      { cell: "D2", value: "=$A$4+B$4" },
    ]);

    // C2:D2 → G8 は 4 列右・6 行下への移動である
    await copyRange(p, 1, "C2:D2", "G8");

    const json = await readJson(p, 1);
    // 相対参照は移動量だけずれる: A4 → E10, B4 → F10
    expect(json.formulas.G8.f).toBe("E10+F10");
    // 絶対参照 $A$4 は不変、行アンカー B$4 は列のみずれて F$4 になる
    expect(json.formulas.H8.f).toBe("$A$4+F$4");
    // コピー元の数式は変更されない
    expect(json.formulas.C2.f).toBe("A4+B4");
    expect(json.formulas.D2.f).toBe("$A$4+B$4");
  });

  it("recreates merges contained in the source range at the destination", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Title");
    await writeRows(p, 1, 2, [["x", "y"]]);
    await mergeCells(p, 1, "A1:B1");

    await copyRange(p, 1, "A1:B2", "D5");

    const json = await readJson(p, 1);
    expect(json.mergedCells).toContain("A1:B1");
    expect(json.mergedCells).toContain("D5:E5");
    // マスターのみ値を持ち、結合の子はキー不在である
    expect(json.cells.D5).toBe("Title");
    expect(json.cells.E5).toBeUndefined();
    expect(json.cells.D6).toBe("x");
    expect(json.cells.E6).toBe("y");
  });

  it("copies to another sheet via dest_sheet", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [["a", 1]]);
    await addSheet(p, "Target");

    const msg = await copyRange(p, 1, "A1:B1", "B2", "Target");
    expect(msg).toContain("Target!B2:C2");

    const target = await readJson(p, "Target");
    expect(target.cells.B2).toBe("a");
    expect(target.cells.C2).toBe(1);
    // コピー元シートは変更されない
    const source = await readJson(p, 1);
    expect(source.cells.A1).toBe("a");
    expect(source.cells.B2).toBeUndefined();
  });

  it("overwrites existing destination content", async () => {
    const p = await createTmpWorkbook();
    await writeCells(p, 1, [
      { cell: "A1", value: "new" },
      // B1 は空のまま（空セルのコピーはコピー先を消す）
      { cell: "D1", value: "old1" },
      { cell: "E1", value: "old2" },
    ]);

    await copyRange(p, 1, "A1:B1", "D1");

    const json = await readJson(p, 1);
    expect(json.cells.D1).toBe("new");
    // 空セルのコピーで既存値が消える
    expect(json.cells.E1).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// find_replace
// ---------------------------------------------------------------------------

describe("find_replace", () => {
  it("replaces substrings in plain string cells only; formulas and numbers untouched", async () => {
    const p = await createTmpWorkbook();
    await writeCells(p, 1, [
      { cell: "A1", value: "hello world" },
      { cell: "A2", value: "say hello" },
      { cell: "B1", value: "=A1" },
      { cell: "C1", value: 123 },
    ]);

    const msg = await findReplace(p, "hello", "bye", 1);
    expect(msg).toContain("Replaced in 2 cell(s)");

    const json = await readJson(p, 1);
    expect(json.cells.A1).toBe("bye world");
    expect(json.cells.A2).toBe("say bye");
    // 数式セル・数値セルは変更されない
    expect(json.formulas.B1.f).toBe("A1");
    expect(json.cells.C1).toBe(123);
  });

  it("case_sensitive=true skips different-case text; default is case-insensitive", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Hello World");

    // 大文字小文字を区別すると "hello" は "Hello" にマッチしない
    const noMatch = await findReplace(p, "hello", "bye", 1, true);
    expect(noMatch).toContain("file unchanged");
    let json = await readJson(p, 1);
    expect(json.cells.A1).toBe("Hello World");

    // 区別しない（デフォルト）とマッチする
    const replaced = await findReplace(p, "hello", "bye", 1, false);
    expect(replaced).toContain("Replaced in 1 cell(s)");
    json = await readJson(p, 1);
    expect(json.cells.A1).toBe("bye World");
  });

  it("match_entire_cell=true replaces only exact whole-cell matches", async () => {
    const p = await createTmpWorkbook();
    await writeCells(p, 1, [
      { cell: "A1", value: "hello" },
      { cell: "A2", value: "hello world" },
    ]);

    const msg = await findReplace(p, "hello", "bye", 1, false, true);
    expect(msg).toContain("Replaced in 1 cell(s)");

    const json = await readJson(p, 1);
    expect(json.cells.A1).toBe("bye");
    // 部分一致のセルはセル全体一致モードでは変更されない
    expect(json.cells.A2).toBe("hello world");
  });

  it("returns 'file unchanged' when nothing matches", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "abc");

    const msg = await findReplace(p, "zzz", "x", 1);
    expect(msg).toContain("file unchanged");

    const json = await readJson(p, 1);
    expect(json.cells.A1).toBe("abc");
  });
});

// ---------------------------------------------------------------------------
// sort_range
// ---------------------------------------------------------------------------

describe("sort_range", () => {
  it("sorts ascending by key column with has_header (header stays)", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Name", "Score"],
      ["banana", 20],
      ["apple", 10],
      ["cherry", 30],
    ]);

    const msg = await sortRange(p, 1, "A1:B4", "A", true, true);
    expect(msg).toContain("Sorted 3 row(s)");
    expect(msg).toContain("ascending");

    const json = await readJson(p, 1);
    // ヘッダ行は動かない
    expect(json.cells.A1).toBe("Name");
    expect(json.cells.B1).toBe("Score");
    // データ行はキー列 A の昇順に並ぶ（B 列も行ごとに一緒に動く）
    expect(json.cells.A2).toBe("apple");
    expect(json.cells.B2).toBe(10);
    expect(json.cells.A3).toBe("banana");
    expect(json.cells.B3).toBe(20);
    expect(json.cells.A4).toBe("cherry");
    expect(json.cells.B4).toBe(30);
  });

  it("sorts descending by key column with has_header", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Name", "Score"],
      ["banana", 20],
      ["apple", 10],
      ["cherry", 30],
    ]);

    const msg = await sortRange(p, 1, "A1:B4", "B", false, true);
    expect(msg).toContain("descending");

    const json = await readJson(p, 1);
    expect(json.cells.A1).toBe("Name");
    expect(json.cells.A2).toBe("cherry");
    expect(json.cells.B2).toBe(30);
    expect(json.cells.A3).toBe("banana");
    expect(json.cells.B3).toBe(20);
    expect(json.cells.A4).toBe("apple");
    expect(json.cells.B4).toBe(10);
  });

  it("moves styles together with sorted rows", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Name", "Score"],
      ["banana", 20],
      ["apple", 10],
    ]);
    // "apple" の行（3 行目）だけ太字にする
    await formatCells(p, 1, "A3:B3", { bold: true });

    await sortRange(p, 1, "A1:B3", "A", true, true);

    const json = await readJson(p, 1, undefined, true);
    // "apple" は 2 行目に移動し、太字も一緒に移動する
    expect(json.cells.A2).toBe("apple");
    expect(json.styles.A2.bold).toBe(true);
    expect(json.styles.B2.bold).toBe(true);
    // 3 行目へ移動した "banana" の行は太字ではない
    expect(json.cells.A3).toBe("banana");
    expect(json.styles?.A3?.bold).toBeUndefined();
  });

  it("sorts numbers before text (ascending)", async () => {
    const p = await createTmpWorkbook();
    await writeCells(p, 1, [
      { cell: "A1", value: "apple" },
      { cell: "A2", value: 10 },
      { cell: "A3", value: 2 },
    ]);

    await sortRange(p, 1, "A1:A3", "A", true, false);

    const json = await readJson(p, 1);
    // 数値 < 文字列（Excel の昇順と同じ）
    expect(json.cells.A1).toBe(2);
    expect(json.cells.A2).toBe(10);
    expect(json.cells.A3).toBe("apple");
  });

  it("places empty key cells last regardless of direction", async () => {
    const p = await createTmpWorkbook();
    // A1 は空のまま
    await writeCells(p, 1, [
      { cell: "A2", value: "x" },
      { cell: "A3", value: 5 },
    ]);

    await sortRange(p, 1, "A1:A3", "A", true, false);
    let json = await readJson(p, 1);
    expect(json.cells.A1).toBe(5);
    expect(json.cells.A2).toBe("x");
    expect(json.cells.A3).toBeUndefined();

    // 降順でも空セルは末尾のままである
    await sortRange(p, 1, "A1:A3", "A", false, false);
    json = await readJson(p, 1);
    expect(json.cells.A1).toBe("x");
    expect(json.cells.A2).toBe(5);
    expect(json.cells.A3).toBeUndefined();
  });

  it("rejects a range that intersects merged cells with INVALID_PARAMETER", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["b", 2],
      ["a", 1],
    ]);
    await mergeCells(p, 1, "B1:B2");

    const err = await sortRange(p, 1, "A1:B2", "A", true, false).catch((e) => e);
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
    expect((err as EngineError).message).toContain("intersects merged cells");
  });

  it("rejects a key column outside the range with INVALID_PARAMETER", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["b", 2],
      ["a", 1],
    ]);

    const err = await sortRange(p, 1, "A1:B2", "D", true, false).catch((e) => e);
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
    expect((err as EngineError).message).toContain("outside the range");
  });
});

// ---------------------------------------------------------------------------
// clear_cells — モード別の動作
// ---------------------------------------------------------------------------

describe("clear_cells modes", () => {
  it("mode 'formats' keeps values and clears styles", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "keep me");
    await formatCells(p, 1, "A1:A1", { bold: true, fillColor: "FFFF00" });

    const msg = await clearCells(p, 1, "A1:A1", "formats");
    expect(msg).toContain("Cleared formatting");

    const json = await readJson(p, 1, undefined, true);
    expect(json.cells.A1).toBe("keep me");
    expect(json.styles?.A1).toBeUndefined();
  });

  it("mode 'all' clears both values and styles", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "wipe me");
    await formatCells(p, 1, "A1:A1", { bold: true });
    // 範囲が空になっても読めるよう、別セルに値を残す
    await writeCell(p, 1, "C3", "anchor");

    const msg = await clearCells(p, 1, "A1:A1", "all");
    expect(msg).toContain("values and formatting");

    const json = await readJson(p, 1, undefined, true);
    expect(json.cells.A1).toBeUndefined();
    expect(json.styles?.A1).toBeUndefined();
  });

  it("default mode 'values' clears values but keeps styles", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "clear me");
    await formatCells(p, 1, "A1:A1", { bold: true });

    const msg = await clearCells(p, 1, "A1:A1");
    expect(msg).toContain("Cleared values");

    const json = await readJson(p, 1, undefined, true);
    expect(json.cells.A1).toBeUndefined();
    // 注意: read_sheet は常に compact 走査のため、値が空のセルは書式付きでも
    // 出力から省略される（includeStyles でも見えない）。書式の保持確認は
    // read_cell で行う。
    expect(json.styles?.A1).toBeUndefined();
    const cellJson = JSON.parse(
      (await readCell(p, 1, "A1")).split("<json>")[1].split("</json>")[0],
    );
    expect(cellJson.value).toBeNull();
    expect(cellJson.style.bold).toBe(true);
  });
});
