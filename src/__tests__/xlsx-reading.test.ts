import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
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
    expect(json.data.length).toBeGreaterThanOrEqual(3);
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
    expect(json.data.length).toBe(2);
    expect(json.data[0].cells.length).toBe(2);
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
    expect(result).toContain("Merged cells: A1:C1");

    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.mergedCells).toEqual(["A1:C1"]);

    // Master cell (A1) should have mergeRange
    const a1 = json.data[0].cells.find((c: { address: string }) => c.address === "A1");
    expect(a1.mergeRange).toBe("A1:C1");

    // Non-master cell (B1) should have mergedWith but no duplicated value
    const b1 = json.data[0].cells.find((c: { address: string }) => c.address === "B1");
    expect(b1.mergedWith).toBe("A1");
    expect(b1.value).toBeNull();
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

    const a1 = json.data[0].cells.find((c: { address: string }) => c.address === "A1");
    expect(a1.value).toBe("Long repeated value");
    expect(a1.mergeRange).toBe("A1:E1");

    for (const addr of ["B1", "C1", "D1", "E1"]) {
      const cell = json.data[0].cells.find((c: { address: string }) => c.address === addr);
      expect(cell.value).toBeNull();
      expect(cell.mergedWith).toBe("A1");
    }
  });
});

describe("compact output mode", () => {
  it("omits merged children in compact mode", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Header");
    await writeCell(p, 1, "D1", "Other");
    await mergeCells(p, 1, "A1:C1");

    const result = await readSheet(p, 1, undefined, true);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    const addresses = json.data[0].cells.map((c: { address: string }) => c.address);
    expect(addresses).toContain("A1");
    expect(addresses).toContain("D1");
    expect(addresses).not.toContain("B1");
    expect(addresses).not.toContain("C1");
    expect(json.compact).toBe(true);
  });

  it("omits null/empty cells in compact mode", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["A", null, "C", null],
      [null, null, null, null],
      ["D", null, null, "E"],
    ]);

    const result = await readSheet(p, 1, "A1:D3", true);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    // Row 2 (all null) should be omitted
    const rowNumbers = json.data.map((r: { row: number }) => r.row);
    expect(rowNumbers).not.toContain(2);

    // Row 1 should only have A1 and C1
    const row1 = json.data.find((r: { row: number }) => r.row === 1);
    const row1Addrs = row1.cells.map((c: { address: string }) => c.address);
    expect(row1Addrs).toEqual(["A1", "C1"]);
  });

  it("compact=false (default) includes all cells", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Header");
    await mergeCells(p, 1, "A1:C1");

    const result = await readSheet(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);

    const addresses = json.data[0].cells.map((c: { address: string }) => c.address);
    expect(addresses).toContain("B1");
    expect(addresses).toContain("C1");
    expect(json.compact).toBeUndefined();
  });
});
