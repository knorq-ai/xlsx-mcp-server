import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  writeCell,
  formatCells,
  formatCellsBulk,
  readCell,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("format_cells", () => {
  it("applies bold and font size", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Header");
    await formatCells(p, 1, "A1:A1", { bold: true, fontSize: 16 });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.bold).toBe(true);
    expect(json.style.fontSize).toBe(16);
  });

  it("applies fill color", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "B2", "colored");
    await formatCells(p, 1, "B2:B2", { fillColor: "FFFF00" });

    const result = await readCell(p, 1, "B2");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.fillColor).toBe("FFFF00");
  });

  it("applies borders", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "bordered");
    await formatCells(p, 1, "A1:A1", { borderStyle: "thin" });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.borderTop).toBe(true);
    expect(json.style.borderBottom).toBe(true);
  });

  it("applies alignment", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "centered");
    await formatCells(p, 1, "A1:A1", {
      horizontalAlignment: "center",
      verticalAlignment: "middle",
      wrapText: true,
    });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.horizontalAlignment).toBe("center");
    expect(json.style.verticalAlignment).toBe("middle");
    expect(json.style.wrapText).toBe(true);
  });

  it("applies number format", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", 1234.5);
    await formatCells(p, 1, "A1:A1", { numFmt: "#,##0.00" });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.numFmt).toBe("#,##0.00");
  });

  it("applies format to a range of cells", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "a");
    await writeCell(p, 1, "B1", "b");
    await writeCell(p, 1, "A2", "c");
    await writeCell(p, 1, "B2", "d");

    const msg = await formatCells(p, 1, "A1:B2", { bold: true });
    expect(msg).toContain("4 cell(s)");
  });

  it("applies font color", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "red text");
    await formatCells(p, 1, "A1:A1", { fontColor: "FF0000" });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.fontColor).toBe("FF0000");
  });

  it("applies font name", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "custom font");
    await formatCells(p, 1, "A1:A1", { fontName: "Courier New" });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.fontName).toBe("Courier New");
  });

  it("applies text rotation", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "angled");
    await formatCells(p, 1, "A1:A1", { textRotation: 45 });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.textRotation).toBe(45);
  });

  it("applies border to specific sides only", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "partial border");
    await formatCells(p, 1, "A1:A1", {
      borderStyle: "medium",
      borderTop: true,
      borderBottom: true,
      borderLeft: false,
      borderRight: false,
    });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.borderTop).toBe(true);
    expect(json.style.borderBottom).toBe(true);
    expect(json.style.borderStyle).toBe("medium");
    expect(json.style.borderLeft).toBeUndefined();
    expect(json.style.borderRight).toBeUndefined();
  });
});

describe("format_cells_bulk", () => {
  it("applies different formats to multiple ranges", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "header");
    await writeCell(p, 1, "A2", 100);

    await formatCellsBulk(p, 1, [
      { range: "A1:A1", format: { bold: true, fontSize: 14 } },
      { range: "A2:A2", format: { numFmt: "#,##0" } },
    ]);

    const h = await readCell(p, 1, "A1");
    const hJson = JSON.parse(h.split("<json>")[1].split("</json>")[0]);
    expect(hJson.style.bold).toBe(true);

    const d = await readCell(p, 1, "A2");
    const dJson = JSON.parse(d.split("<json>")[1].split("</json>")[0]);
    expect(dJson.style.numFmt).toBe("#,##0");
  });
});
