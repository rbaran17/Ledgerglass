import { describe, expect, test } from "vitest";
import {
  addWorksheet,
  createBlankWorkbook,
} from "../../src/app/03-workbook-model.js";
import { calculateWorkbookValues } from "../../src/node/values.js";

describe("calculateWorkbookValues", () => {
  test("returns each sheet's calculated cells under the sheet's name", () => {
    const workbook = createBlankWorkbook("Budget");
    Object.assign(workbook.sheets[0].cells, {
      B2: { raw: 1200 },
      B3: { raw: 450.5 },
      B4: { raw: "=SUM(B2:B3)" },
      C4: { raw: "=1/0" },
    });
    const { report, sheets } = calculateWorkbookValues(workbook);
    expect(report.workbookId).toBe(workbook.id);
    // Plain objects, so the result compares equal to a literal and has the usual methods.
    expect(sheets).toStrictEqual({
      Sheet1: {
        B2: { type: "number", value: 1200 },
        B3: { type: "number", value: 450.5 },
        B4: { type: "number", value: 1650.5 },
        C4: { type: "error", value: "#DIV/0!" },
      },
    });
    expect(sheets.hasOwnProperty("Sheet1")).toBe(true);
  });

  test("keeps a sheet named __proto__ as a sheet", () => {
    let workbook = createBlankWorkbook("Budget");
    workbook.sheets[0].cells.A1 = { raw: 1 };
    workbook = addWorksheet(workbook, "__proto__");
    workbook.sheets[1].cells.B5 = { raw: "=1/0" };
    const { sheets } = calculateWorkbookValues(workbook);
    expect(Object.keys(sheets)).toEqual(["Sheet1", "__proto__"]);
    expect(Object.getPrototypeOf(sheets)).toBe(Object.prototype);
    expect(Object.hasOwn(sheets, "__proto__")).toBe(true);
    expect(sheets["__proto__"]).toEqual({
      B5: { type: "error", value: "#DIV/0!" },
    });
    expect(JSON.parse(JSON.stringify(sheets))).toEqual({
      Sheet1: { A1: { type: "number", value: 1 } },
      ["__proto__"]: { B5: { type: "error", value: "#DIV/0!" } },
    });
    // The other sheet's cells do not show through as inherited properties.
    expect(sheets.B5).toBeUndefined();
    expect({}.B5).toBeUndefined();
  });
});
