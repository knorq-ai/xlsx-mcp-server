/**
 * 最終レビューで確認された不具合の回帰テスト。
 * 共有数式マスター上書きの値型網羅、クロスシートコピー、置換パターンの
 * リテラル化、空セルのノート永続化、書式付き空セルの読み戻し、件数集計。
 */
import { describe, it, expect, afterEach } from "vitest";
import ExcelJS from "exceljs";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  tmpXlsxPath,
  trackTmpFile,
} from "./helpers.js";
import {
  writeCell,
  writeRows,
  copyRange,
  findReplace,
  clearCells,
  setCellNote,
  formatCells,
  readSheet,
  readCell,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

const json = (s: string) => JSON.parse(s.split("<json>")[1].split("</json>")[0]);

/** C1:C3 の共有数式グループ（マスター C1）を持つブックを作る */
async function createSharedFormulaFixture(): Promise<string> {
  const p = tmpXlsxPath();
  trackTmpFile(p);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("S");
  ws.getCell("A1").value = 1;
  ws.getCell("A2").value = 2;
  ws.getCell("A3").value = 3;
  ws.getCell("C1").value = {
    formula: "A1*10",
    result: 10,
    shareType: "shared",
    ref: "C1:C3",
  } as ExcelJS.CellFormulaValue;
  ws.getCell("C2").value = { sharedFormula: "C1", result: 20 } as ExcelJS.CellSharedFormulaValue;
  ws.getCell("C3").value = { sharedFormula: "C1", result: 30 } as ExcelJS.CellSharedFormulaValue;
  await wb.xlsx.writeFile(p);
  return p;
}

describe("shared-formula master overwrite — all value types", () => {
  it("null over a shared master saves and materializes slaves", async () => {
    const p = await createSharedFormulaFixture();
    await writeCell(p, "S", "C1", null);
    const j = json(await readSheet(p, "S"));
    expect(j.formulas.C2.f).toBe("A2*10");
    expect(j.formulas.C3.f).toBe("A3*10");
    expect(j.sharedGroupMasters).toBeUndefined();
  });

  it("{date} over a shared master saves and materializes slaves", async () => {
    const p = await createSharedFormulaFixture();
    await writeCell(p, "S", "C1", { date: "2024-06-15" });
    const j = json(await readSheet(p, "S"));
    expect(j.dates.C1).toContain("2024-06-15");
    expect(j.formulas.C2.f).toBe("A2*10");
  });

  it("{hyperlink} over a shared master saves and materializes slaves", async () => {
    const p = await createSharedFormulaFixture();
    await writeCell(p, "S", "C1", { hyperlink: "https://example.com", text: "link" });
    const j = json(await readSheet(p, "S"));
    expect(j.hyperlinks.C1).toBe("https://example.com");
    expect(j.formulas.C3.f).toBe("A3*10");
  });

  it("clear_cells over a shared master saves and materializes slaves", async () => {
    const p = await createSharedFormulaFixture();
    await clearCells(p, "S", "C1:C1");
    const j = json(await readSheet(p, "S"));
    expect(j.formulas.C1).toBeUndefined();
    expect(j.formulas.C2.f).toBe("A2*10");
  });
});

describe("copy_range onto another sheet with shared formulas", () => {
  it("materializes the destination sheet before overwriting a shared master", async () => {
    const p = await createSharedFormulaFixture();
    // 共有グループを持たない 2 枚目のシートからコピーする
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(p);
    const src = wb.addWorksheet("Src");
    src.getCell("A1").value = "x";
    src.getCell("B1").value = "y";
    await wb.xlsx.writeFile(p);

    // Src!A1:B1 → S!C1（共有マスターを上書き、スレーブ C2/C3 は残る）
    await copyRange(p, "Src", "A1:B1", "C1", "S");
    const j = json(await readSheet(p, "S"));
    expect(j.cells.C1).toBe("x");
    expect(j.formulas.C2.f).toBe("A2*10");
    expect(j.formulas.C3.f).toBe("A3*10");
  });
});

describe("find_replace literal replacement", () => {
  it("does not expand JS replacement patterns ($&, $', $$)", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [["price: 100 USD", "total cost"]]);

    await findReplace(p, "USD", "$& (dollars)", 1);
    await findReplace(p, "cost", "$' approx", 1);

    const j = json(await readSheet(p, 1));
    expect(j.cells.A1).toBe("price: 100 $& (dollars)");
    expect(j.cells.B1).toBe("total $' approx");
  });
});

describe("notes and styles on empty cells", () => {
  it("a note on an empty cell survives reopening and a subsequent unrelated write", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A5", "anchor");
    await setCellNote(p, 1, "C3", "note on empty cell");

    // 直後の読み戻し
    expect(json(await readCell(p, 1, "C3")).note).toBe("note on empty cell");

    // 別セルへの書き込み（open→save サイクル）を挟んでもノートが残る
    await writeCell(p, 1, "A6", "unrelated");
    expect(json(await readCell(p, 1, "C3")).note).toBe("note on empty cell");
    const j = json(await readSheet(p, 1));
    expect(j.notes.C3).toBe("note on empty cell");
  });

  it("styled-but-empty cells appear in read_sheet when include_styles=true", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "anchor");
    await formatCells(p, 1, "B2", { fillColor: "FFCC00" });

    const j = json(await readSheet(p, 1, undefined, true));
    expect(j.styles.B2.fillColor).toBe("FFCC00");
  });
});

describe("read_sheet summary count", () => {
  it("counts dates and errors as non-empty cells", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "text");
    await writeCell(p, 1, "A2", { date: "2024-01-15" });

    const result = await readSheet(p, 1);
    expect(result).toContain("2 non-empty cell(s) returned");
  });
});
