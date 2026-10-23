import * as XLSX from "xlsx";

export interface FixtureOptions {
  employeeCount?: number;
  missingItemCode?: boolean;
  missingClassification?: boolean;
  conflictingClassification?: boolean;
  invalidQuantity?: unknown;
  duplicateEmployee?: boolean;
  fallbackItemCode?: boolean;
  hiddenProductColumn?: boolean;
  missingProjectCode?: boolean;
  missingEmployeeName?: boolean;
}

export function makeSourceFixture(options: FixtureOptions = {}): ArrayBuffer {
  const count = options.employeeCount ?? 2;
  const rows: unknown[][] = Array.from({ length: 6 + count }, () => []);
  const set = (row: number, column: number, value: unknown) => { rows[row - 1][column - 1] = value; };

  set(3, 2, "EMPLOYEE CODE");
  set(3, 4, options.missingEmployeeName ? "" : "EMPLOYEE NAME");
  set(3, 5, "SITE_CODE__SHIP");
  set(3, 6, "SITE_CODE__DLV");
  set(3, 7, "REMARKS");
  set(3, 8, "SALES_PERS");
  set(3, 9, "PROJ_CODE");
  set(4, 4, "ITEM_CODE");
  set(5, 4, "ORDER_TYPE");
  set(6, 4, "TRAN_TYPE");

  set(3, 18, "PS Product");
  set(3, 19, "LI Product");
  set(3, 20, "Unused Product");
  set(options.fallbackItemCode ? 2 : 4, 18, options.missingItemCode ? "" : "PS-001");
  set(4, 19, "LI-001");
  set(5, 18, options.missingClassification ? "" : "PS");
  set(6, 18, options.conflictingClassification ? "LI" : "PS");
  set(5, 19, "LI");
  set(6, 19, "LI");

  for (let index = 0; index < count; index += 1) {
    const row = index + 7;
    set(row, 2, options.duplicateEmployee && index === 1 ? "E001" : `E${String(index + 1).padStart(3, "0")}`);
    set(row, 4, options.missingEmployeeName ? "" : `Employee ${index + 1}`);
    set(row, 5, `SHIP${index + 1}`);
    set(row, 6, `DLV${index + 1}`);
    set(row, 7, `Remark ${index + 1}`);
    set(row, 8, `SP${index + 1}`);
    set(row, 9, options.missingProjectCode ? "" : `PR${index + 1}`);
    set(row, 18, index === 0 && options.invalidQuantity !== undefined ? options.invalidQuantity : index + 1);
    set(row, 19, index + 2);
    set(row, 20, 0);
  }

  const workbook = XLSX.utils.book_new();
  const allocation = XLSX.utils.aoa_to_sheet(rows);
  if (options.hiddenProductColumn) allocation["!cols"] = Array.from({ length: 20 }, (_, index) => ({ hidden: index === 17 }));
  XLSX.utils.book_append_sheet(workbook, allocation, "allocation");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["ignored"]]), "summary");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["ignored"]]), "Pharmacist Kit allocation");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}
