/**
 * シート付随機能のテスト。
 * 日付・ハイパーリンク書き込み、セルノート、シート可視性・タブ色、
 * 行列の非表示、シート保護、ページ設定、スタイル継承挿入、結合セルのガード。
 */

import { describe, it, expect, afterEach } from "vitest";
import ExcelJS from "exceljs";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  ErrorCode,
} from "./helpers.js";
import {
  writeCell,
  readSheet,
  readCell,
  setCellNote,
  setSheetProperties,
  getSheetProperties,
  addSheet,
  setRowVisibility,
  setColumnVisibility,
  protectSheet,
  unprotectSheet,
  setPageSetup,
  insertRows,
  formatCells,
  mergeCells,
  clearCells,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

/** <json> ブロックをパースする */
function parseJson(result: string): any {
  return JSON.parse(result.split("<json>")[1].split("</json>")[0]);
}

/** 保存済みファイルを素の ExcelJS で開く（エンジンを介さない検証用） */
async function rawOpen(p: string): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(p);
  return wb;
}

describe("date values", () => {
  it("writes { date } and reads it back in the dates map as ISO", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", { date: "2024-06-15" });

    const result = await readSheet(p, 1, "A1:A1");
    const json = parseJson(result);
    // 日付セルは cells ではなく dates マップに ISO 8601 で載る
    expect(json.cells.A1).toBeUndefined();
    expect(json.dates.A1).toMatch(/^2024-06-15T00:00:00/);

    const cellJson = parseJson(await readCell(p, 1, "A1"));
    expect(cellJson.type).toBe("date");
    expect(cellJson.value).toMatch(/^2024-06-15/);
  });

  it("rejects an invalid date string with INVALID_PARAMETER", async () => {
    const p = await createTmpWorkbook();
    await expect(
      writeCell(p, 1, "A1", { date: "not-a-date" }),
    ).rejects.toMatchObject({
      code: ErrorCode.INVALID_PARAMETER,
      message: expect.stringContaining("Invalid date"),
    });
  });
});

describe("hyperlink values", () => {
  it("writes { hyperlink, text } and reads back URL and display text", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", {
      hyperlink: "https://example.com/page",
      text: "Example",
    });

    const json = parseJson(await readSheet(p, 1, "A1:A1"));
    // 表示テキストは cells 側、URL は hyperlinks 側に分離される
    expect(json.cells.A1).toBe("Example");
    expect(json.hyperlinks.A1).toBe("https://example.com/page");

    const cellJson = parseJson(await readCell(p, 1, "A1"));
    expect(cellJson.type).toBe("hyperlink");
    expect(cellJson.value).toBe("Example");
    expect(cellJson.hyperlink).toBe("https://example.com/page");
  });
});

describe("set_cell_note", () => {
  it("sets a note readable via readCell and the readSheet notes map", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B2", 42);
    await setCellNote(p, 1, "B2", "確認済みの値である");

    const cellJson = parseJson(await readCell(p, 1, "B2"));
    expect(cellJson.note).toBe("確認済みの値である");

    const sheetJson = parseJson(await readSheet(p, 1, "B2:B2"));
    expect(sheetJson.notes.B2).toBe("確認済みの値である");
  });

  it("removes a note with null", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B2", 42);
    await setCellNote(p, 1, "B2", "一時メモ");
    await setCellNote(p, 1, "B2", null);

    const cellJson = parseJson(await readCell(p, 1, "B2"));
    expect(cellJson.note).toBeUndefined();

    const sheetJson = parseJson(await readSheet(p, 1, "B2:B2"));
    expect(sheetJson.notes).toBeUndefined();
    // 値自体は残る
    expect(sheetJson.cells.B2).toBe(42);
  });
});

describe("set_sheet_properties", () => {
  it("hides a sheet when another visible sheet remains", async () => {
    const p = await createTmpWorkbook("Main");
    await addSheet(p, "Sub");
    await setSheetProperties(p, "Sub", { state: "hidden" });

    const props = parseJson(await getSheetProperties(p, "Sub"));
    expect(props.state).toBe("hidden");

    const mainProps = parseJson(await getSheetProperties(p, "Main"));
    expect(mainProps.state).toBe("visible");
  });

  it("rejects hiding the only visible sheet", async () => {
    const p = await createTmpWorkbook();
    await expect(
      setSheetProperties(p, 1, { state: "hidden" }),
    ).rejects.toMatchObject({
      code: ErrorCode.INVALID_PARAMETER,
      message: expect.stringContaining("at least one visible sheet"),
    });
  });

  it("sets and removes the tab color", async () => {
    const p = await createTmpWorkbook();
    await setSheetProperties(p, 1, { tabColor: "FF0000" });

    let props = parseJson(await getSheetProperties(p, 1));
    expect(props.tabColor).toEqual({ argb: "FFFF0000" });

    await setSheetProperties(p, 1, { tabColor: null });
    props = parseJson(await getSheetProperties(p, 1));
    expect(props.tabColor).toBeUndefined();
  });
});

describe("row / column visibility", () => {
  it("hides and unhides rows (raw ExcelJS verification)", async () => {
    const p = await createTmpWorkbook();
    // 注意: ExcelJS は「セルも height も無い行」をシリアライズしない
    // （row.model が null を返す）ため、空行の hidden フラグは保存時に
    // 失われる。ここでは値のある行で検証する。
    await writeCell(p, 1, "A2", "r2");
    await writeCell(p, 1, "A3", "r3");
    await writeCell(p, 1, "A4", "r4");
    await setRowVisibility(p, 1, 2, 3, true);

    let wb = await rawOpen(p);
    let ws = wb.worksheets[0];
    expect(ws.getRow(2).hidden).toBe(true);
    expect(ws.getRow(3).hidden).toBe(true);
    expect(ws.getRow(4).hidden).toBeFalsy();

    await setRowVisibility(p, 1, 2, 3, false);
    wb = await rawOpen(p);
    ws = wb.worksheets[0];
    expect(ws.getRow(2).hidden).toBeFalsy();
    expect(ws.getRow(3).hidden).toBeFalsy();
  });

  it("persists the hidden flag for completely empty rows", async () => {
    // ExcelJS はセルも明示的な高さも無い行をシリアライズしないため、
    // setRowVisibility は空行を隠すとき既定の行高を明示して
    // シリアライズ対象にする（これが無いと保存後にフラグが消える）。
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A5", "anchor"); // シート寸法の確保のみ
    await setRowVisibility(p, 1, 2, 3, true);

    const wb = await rawOpen(p);
    const ws = wb.worksheets[0];
    expect(ws.getRow(2).hidden).toBe(true);
    expect(ws.getRow(3).hidden).toBe(true);
  });

  it("hides columns (raw ExcelJS verification)", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "D1", "anchor");
    await setColumnVisibility(p, 1, "B", "C", true);

    const wb = await rawOpen(p);
    const ws = wb.worksheets[0];
    expect(ws.getColumn(2).hidden).toBe(true);
    expect(ws.getColumn(3).hidden).toBe(true);
    expect(ws.getColumn(4).hidden).toBeFalsy();
  });
});

describe("sheet protection", () => {
  it("protects and unprotects a sheet (round-trip)", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "locked content");

    const msg = await protectSheet(p, 1, "secret");
    expect(msg).toContain("Protected");

    // 保護フラグがファイルに永続化されている
    let wb = await rawOpen(p);
    let model = wb.worksheets[0].model as unknown as {
      sheetProtection?: { sheet?: boolean };
    };
    expect(model.sheetProtection?.sheet).toBe(true);

    // 保護後もエンジンで読み取り・解除ができる
    const cellJson = parseJson(await readCell(p, 1, "A1"));
    expect(cellJson.value).toBe("locked content");

    await unprotectSheet(p, 1);
    wb = await rawOpen(p);
    model = wb.worksheets[0].model as unknown as {
      sheetProtection?: { sheet?: boolean };
    };
    expect(model.sheetProtection?.sheet).toBeFalsy();

    // 解除後もファイルが正常に開ける
    const after = parseJson(await readCell(p, 1, "A1"));
    expect(after.value).toBe("locked content");
  });
});

describe("set_page_setup", () => {
  it("persists orientation, printArea and fitToWidth", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "report");
    await setPageSetup(p, 1, {
      orientation: "landscape",
      printArea: "A1:C10",
      fitToWidth: 1,
    });

    const wb = await rawOpen(p);
    const ps = wb.worksheets[0].pageSetup;
    expect(ps.orientation).toBe("landscape");
    expect(ps.printArea).toContain("A1:C10");
    expect(ps.fitToWidth).toBe(1);
    expect(ps.fitToPage).toBe(true);
  });
});

describe("insert_rows with inheritStyle", () => {
  it("copies the style from the row above into the inserted row", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Header");
    await writeCell(p, 1, "B1", "Header2");
    await writeCell(p, 1, "A2", "data");
    await formatCells(p, 1, "A1:B1", { bold: true, fillColor: "FFFF00" });

    await insertRows(p, 1, 2, 1, true);

    // 挿入された行 2 が行 1 の書式を引き継ぐ
    const newCell = parseJson(await readCell(p, 1, "A2"));
    expect(newCell.style.bold).toBe(true);
    expect(newCell.style.fillColor).toBe("FFFF00");

    // 元の行 1 はそのまま、旧行 2 のデータは行 3 へ移動している
    const header = parseJson(await readCell(p, 1, "A1"));
    expect(header.value).toBe("Header");
    const moved = parseJson(await readCell(p, 1, "A3"));
    expect(moved.value).toBe("data");
  });
});

describe("merged cell guards", () => {
  it("rejects writing to a merged child cell, mentioning the master address", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Master");
    await mergeCells(p, 1, "A1:B2");

    await expect(writeCell(p, 1, "B2", "oops")).rejects.toMatchObject({
      code: ErrorCode.INVALID_PARAMETER,
      message: expect.stringContaining("A1"),
    });

    // マスター値は無傷
    const master = parseJson(await readCell(p, 1, "A1"));
    expect(master.value).toBe("Master");
  });

  it("clearCells skips merged children whose master is outside the range", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Master");
    await writeCell(p, 1, "C1", "plain");
    await mergeCells(p, 1, "A1:B1");

    // 範囲に子セル B1 を含むがマスター A1 は含まない → マスター値は残る
    await clearCells(p, 1, "B1:C1");

    const master = parseJson(await readCell(p, 1, "A1"));
    expect(master.value).toBe("Master");
    const plain = await readCell(p, 1, "C1");
    expect(plain).toContain("(empty)");
  });

  it("clearCells clears the master value when the master is inside the range", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Master");
    await mergeCells(p, 1, "A1:B1");

    await clearCells(p, 1, "A1:B1");

    const master = await readCell(p, 1, "A1");
    expect(master).toContain("(empty)");
  });
});
