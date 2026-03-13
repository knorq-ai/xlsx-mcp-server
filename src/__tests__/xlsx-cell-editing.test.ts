import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  writeCell,
  writeCells,
  writeRow,
  writeRows,
  clearCells,
  readSheet,
  readCell,
  createWorkbook,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("write_cell", () => {
  it("writes a string value", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "hello");

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("hello");
    expect(json.type).toBe("string");
  });

  it("writes a number value", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B2", 3.14);

    const result = await readCell(p, 1, "B2");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe(3.14);
    expect(json.type).toBe("number");
  });

  it("writes a boolean value", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "C3", true);

    const result = await readCell(p, 1, "C3");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe(true);
    expect(json.type).toBe("boolean");
  });

  it("writes a formula", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", 10);
    await writeCell(p, 1, "A2", "=A1*2");

    const result = await readCell(p, 1, "A2");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.formula).toBe("A1*2");
  });

  it("clears a cell with null", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "data");
    await writeCell(p, 1, "A1", null);

    const result = await readCell(p, 1, "A1");
    expect(result).toContain("(empty)");
  });
});

describe("write_cells (bulk)", () => {
  it("writes multiple cells at once", async () => {
    const p = await createTmpWorkbook();
    await writeCells(p, 1, [
      { cell: "A1", value: "Name" },
      { cell: "B1", value: "Score" },
      { cell: "A2", value: "Alice" },
      { cell: "B2", value: 95 },
    ]);

    const r1 = await readCell(p, 1, "A1");
    expect(JSON.parse(r1.split("<json>")[1].split("</json>")[0]).value).toBe("Name");

    const r2 = await readCell(p, 1, "B2");
    expect(JSON.parse(r2.split("<json>")[1].split("</json>")[0]).value).toBe(95);
  });
});

describe("write_row", () => {
  it("writes a row of values", async () => {
    const p = await createTmpWorkbook();
    await writeRow(p, 1, 1, ["X", "Y", "Z"]);

    const result = await readSheet(p, 1, "A1:C1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    const values = json.data[0].cells.map((c: { value: unknown }) => c.value);
    expect(values).toEqual(["X", "Y", "Z"]);
  });

  it("writes from a start column", async () => {
    const p = await createTmpWorkbook();
    await writeRow(p, 1, 1, [10, 20], "C");

    const c = await readCell(p, 1, "C1");
    expect(JSON.parse(c.split("<json>")[1].split("</json>")[0]).value).toBe(10);

    const d = await readCell(p, 1, "D1");
    expect(JSON.parse(d.split("<json>")[1].split("</json>")[0]).value).toBe(20);
  });
});

describe("write_rows (bulk)", () => {
  it("writes multiple rows", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Header1", "Header2"],
      [100, 200],
      [300, 400],
    ]);

    const result = await readSheet(p, 1, "A1:B3");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.data.length).toBe(3);
    expect(json.data[1].cells[0].value).toBe(100);
    expect(json.data[2].cells[1].value).toBe(400);
  });
});

describe("clear_cells", () => {
  it("clears values in a range", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      [1, 2, 3],
      [4, 5, 6],
    ]);
    await clearCells(p, 1, "A1:B2");

    const a1 = await readCell(p, 1, "A1");
    expect(a1).toContain("(empty)");

    // C1 should still have value
    const c1 = await readCell(p, 1, "C1");
    expect(JSON.parse(c1.split("<json>")[1].split("</json>")[0]).value).toBe(3);
  });
});

describe("create_workbook", () => {
  it("creates a new workbook", async () => {
    const { tmpXlsxPath, trackTmpFile } = await import("./helpers.js");
    const p = trackTmpFile(tmpXlsxPath());
    await createWorkbook(p, "MySheet");

    const result = await readSheet(p, "MySheet");
    expect(result).toContain("MySheet");
  });
});
