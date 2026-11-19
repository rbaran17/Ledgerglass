/** Reading calculated values in bulk, for scripts and servers. */
import { getWorkbookCalculation } from "../app/03-workbook-model.js";

/**
 * Calculates the workbook and returns every filled cell's result.
 *
 * @returns `{ report, sheets }` where `sheets[sheetName][address]` is `{ type, value }` and
 *   `report` is the calculation report. Errors appear as `{ type: "error", value: "#DIV/0!" }`.
 */
export function calculateWorkbookValues(workbook) {
  const { calculator, report } = getWorkbookCalculation(workbook);
  // Object.fromEntries makes every name the object's own property. Plain assignment does not
  // for a sheet named "__proto__": it would replace the object's prototype, and the sheet
  // would be missing from Object.keys and JSON.stringify.
  const sheets = Object.fromEntries(
    workbook.sheets.map((sheet) => [
      sheet.name,
      Object.fromEntries(
        Object.keys(sheet.cells).map((address) => {
          const result = calculator.getTypedValue(sheet.id, address);
          return [address, { type: result.type, value: result.value }];
        }),
      ),
    ]),
  );
  return { report, sheets };
}
