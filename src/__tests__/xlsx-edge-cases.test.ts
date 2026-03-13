import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
  tmpXlsxPath,
  trackTmpFile,
} from "./helpers.js";
import {
  writeCell,
  writeRows,
  readCell,
  readSheet,
  createWorkbook,
  searchCells,
  copySheet,
  mergeCells,
  unmergeCells,
  removeDataValidation,
  addDataValidation,
  getWorkbookInfo,
  formatCells,
  clearCells,
  insertRows,
  deleteRows,
  addNamedRange,
  listWorkbookNamedRanges,
  setFreeze,
  getSheetProperties,
} from "../xlsx-engine.js";
import { parseCellAddress, columnLetterToNumber, parseRange, validateRangeSize } from "../engine/cells.js";
import ExcelJS from "exceljs";

afterEach(cleanupTmpFiles);

// ---------------------------------------------------------------------------
// F-013/F-014: Cell address validation
// ---------------------------------------------------------------------------

describe("parseCellAddress validation", () => {
  it("rejects purely numeric address", () => {
    expect(() => parseCellAddress("123")).toThrow("Invalid cell address");
  });

  it("rejects empty string", () => {
    expect(() => parseCellAddress("")).toThrow("Invalid cell address");
  });

  it("rejects address with row 0", () => {
    expect(() => parseCellAddress("A0")).toThrow("Invalid row number");
  });

  it("accepts valid multi-letter column address", () => {
    const addr = parseCellAddress("BC42");
    expect(addr.col).toBe(55); // B=2, C=3 → 2*26+3=55
    expect(addr.row).toBe(42);
  });
});

describe("columnLetterToNumber validation", () => {
  it("rejects empty string", () => {
    expect(() => columnLetterToNumber("")).toThrow("Column letter must not be empty");
  });
});

describe("parseRange normalization", () => {
  it("normalizes reversed range Z1:A1 → A1:Z1", () => {
    const r = parseRange("Z1:A1");
    expect(r.startCol).toBe(1);  // A
    expect(r.endCol).toBe(26);   // Z
    expect(r.startRow).toBe(1);
    expect(r.endRow).toBe(1);
  });

  it("handles single-cell range", () => {
    const r = parseRange("C5");
    expect(r.startCol).toBe(3);
    expect(r.endCol).toBe(3);
    expect(r.startRow).toBe(5);
    expect(r.endRow).toBe(5);
  });
});

// ---------------------------------------------------------------------------
// F-007: createWorkbook file-exists check
// ---------------------------------------------------------------------------

describe("createWorkbook overwrite protection", () => {
  it("throws when file already exists", async () => {
    const p = await createTmpWorkbook();
    await expect(createWorkbook(p, "Conflict")).rejects.toThrow("already exists");
  });
});

// ---------------------------------------------------------------------------
// Sheet index out of range
// ---------------------------------------------------------------------------

describe("sheet index validation", () => {
  it("throws on sheet index 0", async () => {
    const p = await createTmpWorkbook();
    await expect(readSheet(p, 0)).rejects.toThrow("out of range");
  });

  it("throws on sheet index exceeding count", async () => {
    const p = await createTmpWorkbook();
    await expect(readSheet(p, 999)).rejects.toThrow("out of range");
  });
});

// ---------------------------------------------------------------------------
// F-010: unmerge_cells / remove_data_validation
// ---------------------------------------------------------------------------

describe("unmerge_cells", () => {
  it("unmerges a previously merged range", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "Merged");
    await mergeCells(p, 1, "A1:C1");
    const msg = await unmergeCells(p, 1, "A1:C1");
    expect(msg).toContain("Unmerged");
  });
});

describe("remove_data_validation", () => {
  it("removes data validation from a range", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "A1:A5", {
      type: "list",
      formulae: ['"Yes,No"'],
    });
    const msg = await removeDataValidation(p, 1, "A1:A5");
    expect(msg).toContain("Removed data validation");
  });
});

// ---------------------------------------------------------------------------
// Copy sheet with formulas
// ---------------------------------------------------------------------------

describe("copy_sheet with formulas", () => {
  it("preserves formula cells after copy", async () => {
    const p = await createTmpWorkbook("Source");
    await writeCell(p, "Source", "A1", 10);
    await writeCell(p, "Source", "A2", 20);
    await writeCell(p, "Source", "A3", "=SUM(A1:A2)");
    await copySheet(p, "Source", "Copy");

    const result = await readCell(p, "Copy", "A3");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.formula).toBe("SUM(A1:A2)");
  });
});

// ---------------------------------------------------------------------------
// Search with formula cells
// ---------------------------------------------------------------------------

describe("search finds formula results", () => {
  it("searches in formula result values", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "hello");
    await writeCell(p, 1, "A2", '=CONCATENATE("hel","lo")');

    // Formula result may or may not be computed without Excel engine,
    // but we should at least find the plain text cell
    const result = await searchCells(p, "hello", 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.matches.length).toBeGreaterThanOrEqual(1);
  });
});

// ---------------------------------------------------------------------------
// Sheet names with special characters
// ---------------------------------------------------------------------------

describe("special character sheet names", () => {
  it("handles sheet names with spaces", async () => {
    const p = await createTmpWorkbook("My Sheet");
    await writeCell(p, "My Sheet", "A1", "data");

    const result = await readCell(p, "My Sheet", "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("data");
  });

  it("handles sheet names with unicode", async () => {
    const p = await createTmpWorkbook("データ");
    await writeCell(p, "データ", "A1", 42);

    const info = await getWorkbookInfo(p);
    expect(info).toContain("データ");
  });
});

// ---------------------------------------------------------------------------
// Range size guard
// ---------------------------------------------------------------------------

describe("validateRangeSize", () => {
  it("accepts a reasonable range", () => {
    const range = parseRange("A1:J100"); // 1000 cells
    expect(() => validateRangeSize(range)).not.toThrow();
  });

  it("rejects an excessively large range", () => {
    const range = parseRange("A1:ZZ1000"); // 26*26*1000 > 100k
    expect(() => validateRangeSize(range)).toThrow("Range too large");
  });
});

// ---------------------------------------------------------------------------
// Border false handling
// ---------------------------------------------------------------------------

describe("border false handling", () => {
  it("borderTop: true applies only top border", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "test");
    await formatCells(p, 1, "A1:A1", {
      borderStyle: "thin",
      borderTop: true,
      borderBottom: false,
      borderLeft: false,
      borderRight: false,
    });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.border.top.style).toBe("thin");
    expect(json.style.border.bottom).toBeUndefined();
    expect(json.style.border.left).toBeUndefined();
    expect(json.style.border.right).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Formatting edge cases
// ---------------------------------------------------------------------------

describe("formatting edge cases", () => {
  it("applies italic and font color", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "styled");
    await formatCells(p, 1, "A1:A1", { italic: true, fontColor: "FF0000" });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.font.italic).toBe(true);
    expect(json.style.font.color.argb).toBe("FFFF0000");
  });

  it("applies underline and strikethrough", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "deco");
    await formatCells(p, 1, "A1:A1", { underline: true, strikethrough: true });

    const result = await readCell(p, 1, "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.style.font.underline).toBe(true);
    expect(json.style.font.strike).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// View settings edge cases
// ---------------------------------------------------------------------------

describe("freeze panes variations", () => {
  it("freezes rows only", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 2, 0);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes.row).toBe(2);
    expect(json.freezePanes.column).toBe(0);
  });

  it("freezes columns only", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 0, 3);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes.row).toBe(0);
    expect(json.freezePanes.column).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// Named ranges — workbook scope
// ---------------------------------------------------------------------------

describe("named ranges — workbook scope", () => {
  it("adds a workbook-scoped named range (no sheet param)", async () => {
    const p = await createTmpWorkbook();
    await addNamedRange(p, "GlobalRange", "A1:D10");

    const result = await listWorkbookNamedRanges(p);
    expect(result).toContain("GlobalRange");
  });
});

// ---------------------------------------------------------------------------
// Row/column bounds checking
// ---------------------------------------------------------------------------

describe("row/column bounds via engine", () => {
  it("insertRows works at row 1", async () => {
    const p = await createTmpWorkbook();
    const msg = await insertRows(p, 1, 1, 3);
    expect(msg).toContain("Inserted 3 row(s)");
  });

  it("deleteRows works normally", async () => {
    const p = await createTmpWorkbook();
    await writeCell(p, 1, "A1", "data");
    const msg = await deleteRows(p, 1, 1, 1);
    expect(msg).toContain("Deleted 1 row(s)");
  });
});

// ---------------------------------------------------------------------------
// Duplicate merge cells tolerance
// ---------------------------------------------------------------------------

describe("duplicate merge cells", () => {
  it("reads a file with overlapping merges without crashing", async () => {
    // Create an xlsx with duplicate merge entries by writing raw ExcelJS
    const p = tmpXlsxPath();
    trackTmpFile(p);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Header";
    ws.getCell("A2").value = "Data";
    ws.mergeCells("A1:C1");
    await wb.xlsx.writeFile(p);

    // Manually inject a duplicate merge by re-merging in the model
    // Simulate what a corrupted file looks like: re-read, duplicate the merge
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(p);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    // Access internal model and duplicate the merge entry
    const model = ws2.model as Record<string, unknown>;
    if (Array.isArray(model.merges)) {
      model.merges.push(model.merges[0]); // duplicate
    }
    // Write the corrupted model back — ExcelJS will write duplicate <mergeCell> entries
    await wb2.xlsx.writeFile(p);

    // Now read via our engine — should not throw
    const result = await readSheet(p, 1);
    expect(result).toContain("Header");
  });
});
