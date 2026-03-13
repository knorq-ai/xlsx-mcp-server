import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  addSheet,
  renameSheet,
  deleteSheet,
  copySheet,
  writeCell,
  readCell,
  getWorkbookInfo,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("add_sheet", () => {
  it("adds a new sheet", async () => {
    const p = await createTmpWorkbook();
    await addSheet(p, "NewSheet");

    const info = await getWorkbookInfo(p);
    expect(info).toContain("NewSheet");
    const json = JSON.parse(info.split("<json>")[1].split("</json>")[0]);
    expect(json.sheetCount).toBe(2);
  });

  it("throws on duplicate name", async () => {
    const p = await createTmpWorkbook("Test");
    await expect(addSheet(p, "Test")).rejects.toThrow("already exists");
  });
});

describe("rename_sheet", () => {
  it("renames a sheet by name", async () => {
    const p = await createTmpWorkbook("Old");
    await renameSheet(p, "Old", "New");

    const info = await getWorkbookInfo(p);
    expect(info).toContain("New");
    expect(info).not.toContain('"Old"');
  });

  it("renames a sheet by index", async () => {
    const p = await createTmpWorkbook("First");
    await renameSheet(p, 1, "Renamed");

    const info = await getWorkbookInfo(p);
    expect(info).toContain("Renamed");
  });
});

describe("delete_sheet", () => {
  it("deletes a sheet", async () => {
    const p = await createTmpWorkbook();
    await addSheet(p, "ToDelete");
    await deleteSheet(p, "ToDelete");

    const info = await getWorkbookInfo(p);
    expect(info).not.toContain("ToDelete");
  });
});

describe("copy_sheet", () => {
  it("copies a sheet with data", async () => {
    const p = await createTmpWorkbook("Source");
    await writeCell(p, "Source", "A1", "copied data");
    await copySheet(p, "Source", "Destination");

    const result = await readCell(p, "Destination", "A1");
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.value).toBe("copied data");
  });

  it("throws on duplicate name", async () => {
    const p = await createTmpWorkbook("Exists");
    await expect(copySheet(p, "Exists", "Exists")).rejects.toThrow("already exists");
  });
});
