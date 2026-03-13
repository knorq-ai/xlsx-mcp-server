import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  writeRows,
  readSheet,
  readCell,
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  setColumnWidth,
  setColumnWidths,
  setRowHeight,
  setRowHeights,
  getSheetProperties,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("insert_rows", () => {
  it("inserts rows and shifts data down", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["A", "B"],
      ["C", "D"],
    ]);
    await insertRows(p, 1, 1, 2);

    // Original row 1 data should now be at row 3
    const cell = await readCell(p, 1, "A3");
    const json = JSON.parse(cell.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("A");
  });
});

describe("delete_rows", () => {
  it("deletes rows and shifts data up", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["row1"],
      ["row2"],
      ["row3"],
    ]);
    await deleteRows(p, 1, 1, 1);

    // Row 2 should now be at row 1
    const cell = await readCell(p, 1, "A1");
    const json = JSON.parse(cell.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("row2");
  });
});

describe("insert_columns", () => {
  it("inserts columns and shifts data right", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [["X", "Y"]]);
    await insertColumns(p, 1, "A", 1);

    // X should now be at B1
    const cell = await readCell(p, 1, "B1");
    const json = JSON.parse(cell.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("X");
  });
});

describe("delete_columns", () => {
  it("deletes columns and shifts data left", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [["A", "B", "C"]]);
    await deleteColumns(p, 1, "A", 1);

    // B should now be at A1
    const cell = await readCell(p, 1, "A1");
    const json = JSON.parse(cell.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("B");
  });
});

describe("set_column_width", () => {
  it("sets column width", async () => {
    const p = await createTmpWorkbook();
    const msg = await setColumnWidth(p, 1, "A", 20);
    expect(msg).toContain("width = 20");
  });
});

describe("set_column_widths", () => {
  it("sets multiple column widths", async () => {
    const p = await createTmpWorkbook();
    const msg = await setColumnWidths(p, 1, [
      { column: "A", width: 15 },
      { column: "B", width: 25 },
    ]);
    expect(msg).toContain("2 column(s)");
  });
});

describe("set_row_height", () => {
  it("sets row height", async () => {
    const p = await createTmpWorkbook();
    const msg = await setRowHeight(p, 1, 1, 30);
    expect(msg).toContain("height = 30");
  });
});

describe("set_row_heights", () => {
  it("sets multiple row heights", async () => {
    const p = await createTmpWorkbook();
    const msg = await setRowHeights(p, 1, [
      { row: 1, height: 20 },
      { row: 2, height: 40 },
    ]);
    expect(msg).toContain("2 row(s)");
  });
});
