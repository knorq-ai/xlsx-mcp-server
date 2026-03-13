import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  addNamedRange,
  deleteNamedRange,
  listWorkbookNamedRanges,
  addSheet,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("add_named_range", () => {
  it("adds a named range scoped to a sheet", async () => {
    const p = await createTmpWorkbook("Data");
    await addNamedRange(p, "MyRange", "A1:C10", "Data");

    const result = await listWorkbookNamedRanges(p);
    expect(result).toContain("MyRange");
  });

  it("adds a workbook-scoped named range (no sheet)", async () => {
    const p = await createTmpWorkbook();
    await addNamedRange(p, "GlobalRange", "A1:D10");

    const result = await listWorkbookNamedRanges(p);
    expect(result).toContain("GlobalRange");
  });

  it("throws on duplicate name", async () => {
    const p = await createTmpWorkbook("Data");
    await addNamedRange(p, "Range1", "A1:A10", "Data");
    await expect(addNamedRange(p, "Range1", "B1:B10", "Data")).rejects.toThrow("already exists");
  });

  it("handles sheet names with special characters", async () => {
    const p = await createTmpWorkbook("Sheet's Data");
    await addNamedRange(p, "SpecialRange", "A1:B5", "Sheet's Data");

    const result = await listWorkbookNamedRanges(p);
    expect(result).toContain("SpecialRange");
  });
});

describe("delete_named_range", () => {
  it("deletes a named range", async () => {
    const p = await createTmpWorkbook("Data");
    await addNamedRange(p, "ToDelete", "A1:A5", "Data");
    await deleteNamedRange(p, "ToDelete");

    const result = await listWorkbookNamedRanges(p);
    expect(result).not.toContain("ToDelete");
  });

  it("throws on non-existent name", async () => {
    const p = await createTmpWorkbook();
    await expect(deleteNamedRange(p, "NoSuchRange")).rejects.toThrow("not found");
  });
});

describe("list_named_ranges", () => {
  it("returns empty list when no named ranges", async () => {
    const p = await createTmpWorkbook();
    const result = await listWorkbookNamedRanges(p);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.namedRanges).toEqual([]);
  });

  it("lists multiple named ranges", async () => {
    const p = await createTmpWorkbook("Data");
    await addNamedRange(p, "Range1", "A1:A10", "Data");
    await addNamedRange(p, "Range2", "B1:B10", "Data");

    const result = await listWorkbookNamedRanges(p);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.namedRanges.length).toBe(2);
  });
});
