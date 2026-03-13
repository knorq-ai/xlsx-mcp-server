import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  setFreeze,
  setSheetAutoFilter,
  removeSheetAutoFilter,
  getSheetProperties,
  writeCell,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("set_freeze_panes", () => {
  it("freezes rows and columns", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 1, 2);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes.row).toBe(1);
    expect(json.freezePanes.column).toBe(2);
  });

  it("unfreezes with 0, 0", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 1, 1);
    await setFreeze(p, 1, 0, 0);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes).toBeUndefined();
  });

  it("freezes only rows (column=0)", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 3, 0);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes.row).toBe(3);
    expect(json.freezePanes.column).toBe(0);
  });

  it("freezes only columns (row=0)", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 0, 2);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes.row).toBe(0);
    expect(json.freezePanes.column).toBe(2);
  });

  it("overwrites previous freeze", async () => {
    const p = await createTmpWorkbook();
    await setFreeze(p, 1, 1, 1);
    await setFreeze(p, 1, 5, 3);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.freezePanes.row).toBe(5);
    expect(json.freezePanes.column).toBe(3);
  });
});

describe("set_auto_filter", () => {
  it("sets auto filter", async () => {
    const p = await createTmpWorkbook();
    await setSheetAutoFilter(p, 1, "A1:D1");

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.autoFilter).toBeDefined();
  });

  it("updates auto filter range", async () => {
    const p = await createTmpWorkbook();
    await setSheetAutoFilter(p, 1, "A1:C1");
    await setSheetAutoFilter(p, 1, "A1:F1");

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.autoFilter).toBeDefined();
  });
});

describe("remove_auto_filter", () => {
  it("removes auto filter", async () => {
    const p = await createTmpWorkbook();
    await setSheetAutoFilter(p, 1, "A1:D1");
    await removeSheetAutoFilter(p, 1);

    const result = await getSheetProperties(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.autoFilter).toBeUndefined();
  });

  it("no-op on sheet without filter", async () => {
    const p = await createTmpWorkbook();
    const msg = await removeSheetAutoFilter(p, 1);
    expect(msg).toContain("Removed");
  });
});
