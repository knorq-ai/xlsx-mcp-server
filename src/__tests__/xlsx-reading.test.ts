import { describe, it, expect, afterEach } from "vitest";
import ExcelJS from "exceljs";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  tmpXlsxPath,
  trackTmpFile,
} from "./helpers.js";
import {
  getWorkbookInfo,
  readSheet,
  readCell,
  searchCells,
  getSheetProperties,
  writeCell,
  writeRows,
  mergeCells,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("get_workbook_info", () => {
  it("returns sheet list and metadata", async () => {
    const p = await createTmpWorkbook("TestSheet");
    const result = await getWorkbookInfo(p);

    expect(result).toContain("TestSheet");
    expect(result).toContain("Sheets: 1");
    expect(result).toContain("<json>");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.sheetCount).toBe(1);
    expect(json.sheets[0].name).toBe("TestSheet");
  });
});

describe("read_sheet", () => {
  it("reads all data from a sheet", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 25],
    ]);

    const result = await readSheet(p, 1);
    expect(result).toContain("Alice");
    expect(result).toContain("Bob");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.cells.A2).toBe("Alice");
    expect(json.cells.B2).toBe(30);
    expect(json.cells.A3).toBe("Bob");
  });

  it("reads a specific range", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["A", "B", "C"],
      [1, 2, 3],
      [4, 5, 6],
    ]);

    const result = await readSheet(p, 1, "A1:B2");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    // 範囲外のセル (C 列, 行 3) は含まれない
    expect(Object.keys(json.cells).sort()).toEqual(["A1", "A2", "B1", "B2"]);
    expect(json.cells.C1).toBeUndefined();
  });

  it("supports sheet name reference", async () => {
    const p = await createTmpWorkbook("Data");
    await writeCell(p, "Data", "A1", "hello");

    const result = await readSheet(p, "Data");
    expect(result).toContain("hello");
  });
});

describe("read_cell", () => {
  it("reads a single cell value and type", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B3", 42);

    const result = await readCell(p, 1, "B3");
    expect(result).toContain("42");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe(42);
    expect(json.type).toBe("number");
  });

  it("reads a formula cell", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", 10);
    await writeCell(p, 1, "A2", 20);
    await writeCell(p, 1, "A3", "=SUM(A1:A2)");

    const result = await readCell(p, 1, "A3");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.formula).toBe("SUM(A1:A2)");
    expect(json.type).toBe("formula");
  });

  it("reads an empty cell", async () => {
    const p = await createTmpWorkbook();
    const result = await readCell(p, 1, "Z99");
    expect(result).toContain("(empty)");
  });

  // Regression: shared-formula slave cells must report their own translated
  // formula, not the master cell's address. (GitHub issue #3)
  it("reads the translated formula for shared-formula slave cells", async () => {
    const p = tmpXlsxPath();
    trackTmpFile(p);

    // Build a fixture with a genuine shared formula group G2:I2.
    // Master G2 = `$C2*D2`; slaves H2/I2 reference the master via `sharedFormula`
    // and Excel slides the relative refs (D2 → E2 → F2).
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Split");
    ws.getCell("C2").value = 770000;
    ws.getCell("D2").value = 0.1;
    ws.getCell("E2").value = 0.9;
    ws.getCell("F2").value = 0;
    ws.getCell("G2").value = {
      formula: "$C2*D2",
      result: 77000,
      shareType: "shared",
      ref: "G2:I2",
    } as ExcelJS.CellFormulaValue;
    ws.getCell("H2").value = {
      sharedFormula: "G2",
      result: 693000,
    } as ExcelJS.CellSharedFormulaValue;
    ws.getCell("I2").value = {
      sharedFormula: "G2",
      result: 0,
    } as ExcelJS.CellSharedFormulaValue;
    await wb.xlsx.writeFile(p);

    const h2 = JSON.parse(
      (await readCell(p, 1, "H2")).split("<json>")[1].split("</json>")[0],
    );
    expect(h2.type).toBe("formula");
    expect(h2.value).toBe(693000);
    expect(h2.formula).toBe("$C2*E2"); // not "G2"

    const i2 = JSON.parse(
      (await readCell(p, 1, "I2")).split("<json>")[1].split("</json>")[0],
    );
    expect(i2.formula).toBe("$C2*F2"); // not "G2"

    // Master cell still reports its own formula.
    const g2 = JSON.parse(
      (await readCell(p, 1, "G2")).split("<json>")[1].split("</json>")[0],
    );
    expect(g2.formula).toBe("$C2*D2");
  });
});

describe("search_cells", () => {
  it("finds matching cells", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["apple", "banana"],
      ["cherry", "apple pie"],
    ]);

    const result = await searchCells(p, "apple");
    expect(result).toContain("2 match");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.matches.length).toBe(2);
  });

  it("supports case-sensitive search", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [["Apple", "apple"]]);

    const cs = await searchCells(p, "Apple", undefined, true);
    const jsonCs = JSON.parse(cs.split("<json>")[1].split("</json>")[0]);
    expect(jsonCs.matches.length).toBe(1);

    const ci = await searchCells(p, "Apple", undefined, false);
    const jsonCi = JSON.parse(ci.split("<json>")[1].split("</json>")[0]);
    expect(jsonCi.matches.length).toBe(2);
  });

  it("searches specific sheet", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "target");

    const result = await searchCells(p, "target", 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.matches.length).toBe(1);
  });
});

describe("get_sheet_properties", () => {
  it("returns basic properties", async () => {
    const p = await createTmpWorkbook("MySheet");
    const result = await getSheetProperties(p, "MySheet");

    expect(result).toContain("MySheet");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.name).toBe("MySheet");
  });
});

// ---------------------------------------------------------------------------
// Merged cell info in read_sheet / read_cell
// ---------------------------------------------------------------------------

describe("merged cell info", () => {
  it("read_sheet includes mergedCells list and per-cell merge info", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Header");
    await writeCell(p, 1, "D1", "Other");
    await mergeCells(p, 1, "A1:C1");

    const result = await readSheet(p, 1);
    expect(result).toContain("Merged cells in range: A1:C1");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.mergedCells).toEqual(["A1:C1"]);

    // Master cell (A1) carries the value; merged children are absent
    expect(json.cells.A1).toBe("Header");
    expect(json.cells.B1).toBeUndefined();
    expect(json.cells.C1).toBeUndefined();
    expect(json.cells.D1).toBe("Other");
  });

  it("read_cell shows merge info for master cell", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Merged");
    await mergeCells(p, 1, "A1:B2");

    const result = await readCell(p, 1, "A1");
    expect(result).toContain("Merge: master of A1:B2");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.mergeRange).toBe("A1:B2");
  });

  it("read_cell shows merge info for non-master cell", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Merged");
    await mergeCells(p, 1, "A1:B2");

    const result = await readCell(p, 1, "B2");
    expect(result).toContain("Merge: part of A1");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.mergedWith).toBe("A1");
  });

  it("merged children do not duplicate the master cell value", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Long repeated value");
    await mergeCells(p, 1, "A1:E1");

    const result = await readSheet(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    expect(json.cells.A1).toBe("Long repeated value");
    expect(json.mergedCells).toEqual(["A1:E1"]);

    for (const addr of ["B1", "C1", "D1", "E1"]) {
      expect(json.cells[addr]).toBeUndefined();
    }
  });
});

describe("map output format", () => {
  it("omits merged children", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Header");
    await writeCell(p, 1, "D1", "Other");
    await mergeCells(p, 1, "A1:C1");

    const result = await readSheet(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    expect(json.cells.A1).toBe("Header");
    expect(json.cells.D1).toBe("Other");
    expect(json.cells.B1).toBeUndefined();
    expect(json.cells.C1).toBeUndefined();
  });

  it("omits null/empty cells", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["A", null, "C", null],
      [null, null, null, null],
      ["D", null, null, "E"],
    ]);

    const result = await readSheet(p, 1, "A1:D3");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    expect(Object.keys(json.cells).sort()).toEqual(["A1", "A3", "C1", "D3"]);
  });

  it("separates formulas from plain values and flags uncached results", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", 10);
    await writeCell(p, 1, "A2", "=A1*2");

    const result = await readSheet(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    expect(json.cells.A1).toBe(10);
    expect(json.cells.A2).toBeUndefined();
    expect(json.formulas.A2.f).toBe("A1*2");
    // 書き込み直後はキャッシュ済み結果が無い → v は省略される
    expect(json.formulas.A2.v).toBeUndefined();
  });

  it("includes styles only when requested", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "styled");
    const { formatCells } = await import("../xlsx-engine.js");
    await formatCells(p, 1, "A1", { bold: true, fillColor: "FFFF00" });

    const plain = JSON.parse((await readSheet(p, 1)).split("<json>")[1].split("</json>")[0]);
    expect(plain.styles).toBeUndefined();

    const styled = JSON.parse((await readSheet(p, 1, undefined, true)).split("<json>")[1].split("</json>")[0]);
    expect(styled.styles.A1.bold).toBe(true);
    expect(styled.styles.A1.fillColor).toBe("FFFF00");
  });
});
