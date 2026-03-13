import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  writeCells,
  writeCell,
  writeRows,
  formatCellsBulk,
  setColumnWidths,
  setRowHeights,
  readCell,
  readSheet,
  mergeCells,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("bulk operations", () => {
  it("write_cells handles many cells efficiently", async () => {
    const p = await createTmpWorkbook();
    const cells = [];
    for (let i = 1; i <= 50; i++) {
      cells.push({ cell: `A${i}`, value: `row${i}` });
    }
    const msg = await writeCells(p, 1, cells);
    expect(msg).toContain("50 cell(s)");

    const r = await readCell(p, 1, "A50");
    const json = JSON.parse(r.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("row50");
  });

  it("write_rows handles large datasets", async () => {
    const p = await createTmpWorkbook();
    const rows = [];
    for (let i = 0; i < 100; i++) {
      rows.push([`name${i}`, i, i * 1.5]);
    }
    const msg = await writeRows(p, 1, 1, rows);
    expect(msg).toContain("100 row(s)");

    const r = await readCell(p, 1, "A100");
    const json = JSON.parse(r.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("name99");
  });

  it("format_cells_bulk applies multiple formats in one operation", async () => {
    const p = await createTmpWorkbook();
    await writeRows(p, 1, 1, [
      ["Header1", "Header2"],
      [100, 200],
    ]);

    const msg = await formatCellsBulk(p, 1, [
      { range: "A1:B1", format: { bold: true, fillColor: "4472C4", fontColor: "FFFFFF" } },
      { range: "A2:B2", format: { numFmt: "#,##0" } },
    ]);
    expect(msg).toContain("4 cell(s)");
    expect(msg).toContain("2 group(s)");
  });
});

describe("merge_cells", () => {
  it("merges a range of cells", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Merged");
    const msg = await mergeCells(p, 1, "A1:C1");
    expect(msg).toContain("Merged cells");
  });
});
