import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  createTmpWorkbook,
} from "./helpers.js";
import {
  addDataValidation,
  removeDataValidation,
  listDataValidations,
  writeCell,
} from "../xlsx-engine.js";

afterEach(cleanupTmpFiles);

describe("add_data_validation", () => {
  it("adds a list validation", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "A1:A10", {
      type: "list",
      formulae: ['"Yes,No,Maybe"'],
    });

    const result = await listDataValidations(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.validations.length).toBeGreaterThan(0);
    expect(json.validations[0].type).toBe("list");
  });

  it("adds a whole number validation with range", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "B1:B5", {
      type: "whole",
      formulae: ["1", "100"],
      operator: "between",
      showErrorMessage: true,
      errorTitle: "Invalid",
      error: "Must be 1-100",
    });

    const result = await listDataValidations(p, 1);
    expect(result).toContain("whole");
  });

  it("adds a decimal validation", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "C1:C5", {
      type: "decimal",
      formulae: ["0.0", "99.9"],
      operator: "between",
    });

    const result = await listDataValidations(p, 1);
    expect(result).toContain("decimal");
  });

  it("adds a textLength validation", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "D1:D10", {
      type: "textLength",
      formulae: ["1", "50"],
      operator: "between",
    });

    const result = await listDataValidations(p, 1);
    expect(result).toContain("textLength");
  });

  it("adds a custom formula validation", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "E1:E5", {
      type: "custom",
      formulae: ["LEN(E1)>0"],
    });

    const result = await listDataValidations(p, 1);
    expect(result).toContain("custom");
  });

  it("adds validation with input message", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "A1:A5", {
      type: "list",
      formulae: ['"A,B,C"'],
      showInputMessage: true,
      promptTitle: "Select",
      prompt: "Choose a letter",
    });

    const result = await listDataValidations(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.validations.length).toBeGreaterThan(0);
  });

  it("adds validation with greaterThan operator", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "A1:A3", {
      type: "whole",
      formulae: ["10"],
      operator: "greaterThan",
    });

    const result = await listDataValidations(p, 1);
    expect(result).toContain("whole");
  });

  it("rejects empty formulae array", async () => {
    const p = await createTmpWorkbook();
    await expect(
      addDataValidation(p, 1, "A1:A5", {
        type: "list",
        formulae: [],
      }),
    ).rejects.toThrow("at least one formula");
  });
});

describe("remove_data_validation", () => {
  it("removes validation from a range", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "A1:A5", {
      type: "list",
      formulae: ['"Yes,No"'],
    });
    const msg = await removeDataValidation(p, 1, "A1:A5");
    expect(msg).toContain("Removed data validation");
  });
});

describe("list_data_validations", () => {
  it("returns empty list for sheet without validations", async () => {
    const p = await createTmpWorkbook();
    const result = await listDataValidations(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.validations).toEqual([]);
  });

  it("lists multiple validations on same sheet", async () => {
    const p = await createTmpWorkbook();
    await addDataValidation(p, 1, "A1:A5", {
      type: "list",
      formulae: ['"X,Y"'],
    });
    await addDataValidation(p, 1, "B1:B5", {
      type: "whole",
      formulae: ["0", "999"],
      operator: "between",
    });

    const result = await listDataValidations(p, 1);
    const json = JSON.parse(result.split("<json>")[1].split("</json>")[0]);
    expect(json.validations.length).toBeGreaterThanOrEqual(2);
  });
});
