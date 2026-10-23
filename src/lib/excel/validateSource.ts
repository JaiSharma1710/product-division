import * as XLSX from "xlsx";
import { detectLayout, type SourceLayout } from "@/lib/excel/detectLayout";
import { EMPLOYEE_HEADER_ALIASES, REQUIRED_EMPLOYEE_FIELDS, SOURCE, type EmployeeField } from "@/lib/excel/templateConfig";
import type { EmployeeAllocation, ParsedSource, ProductColumn, ValidationIssue } from "@/lib/types";

type Worksheet = XLSX.WorkSheet;

const normalize = (value: unknown) => String(value ?? "").trim().toUpperCase().replace(/[_\-]+/g, " ").replace(/\s+/g, " ");
const textValue = (value: unknown) => String(value ?? "").trim();

function cellAt(sheet: Worksheet, row: number, column: number): XLSX.CellObject | undefined {
  return sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })] as XLSX.CellObject | undefined;
}

function valueAt(sheet: Worksheet, row: number, column: number): unknown {
  return cellAt(sheet, row, column)?.v;
}

function isBlank(value: unknown): boolean {
  return value === undefined || value === null || (typeof value === "string" && value.trim() === "");
}

function issue(id: string, severity: "error" | "warning", en: string, hi: string, details: Partial<ValidationIssue> = {}): ValidationIssue {
  return { id, severity, en, hi, sheet: SOURCE.SHEET_NAME, ...details };
}

function isTotalColumn(name: string): boolean {
  const normalized = normalize(name);
  return /^(CALCULATED )?TOTAL\b/.test(normalized) || /\bTOTAL (SAMPLES?|PRODUCTS?|QTY|QUANTITY)\b/.test(normalized);
}

function isTotalRow(sheet: Worksheet, row: number, layout: SourceLayout): boolean {
  return [layout.columns.code, layout.columns.name]
    .some((column) => column !== undefined && /^(GRAND )?TOTAL\b/.test(normalize(valueAt(sheet, row, column))));
}

function hasAnyRowData(sheet: Worksheet, row: number, productColumns: readonly number[], layout: SourceLayout): boolean {
  return [...Object.values(layout.columns), ...productColumns].some((column) => !isBlank(valueAt(sheet, row, column)));
}

/** A column whose filled cells are all non-numeric text (such as e-mail addresses) holds information, not quantities. */
function isInformationColumn(sheet: Worksheet, rows: readonly number[], column: number): boolean {
  const values = rows.map((row) => valueAt(sheet, row, column)).filter((value) => !isBlank(value));
  return values.length > 0 && values.every((value) => typeof value === "string" && !Number.isFinite(Number(value.trim())));
}

const FIELD_LABELS: Record<EmployeeField, [string, string]> = {
  code: ["employee code", "कर्मचारी कोड"],
  name: ["employee name", "कर्मचारी नाम"],
  shipSite: ["shipping site", "भेजने की साइट"],
  deliverySite: ["delivery site", "डिलीवरी साइट"],
  remarks: ["remarks", "टिप्पणी"],
  salesPerson: ["salesperson code", "सेल्सपर्सन कोड"],
  projectCode: ["project code", "प्रोजेक्ट कोड"],
};

function validateQuantity(cell: XLSX.CellObject | undefined): { value?: number; error?: "formula" | "type" | "negative" | "fraction" | "infinite" } {
  if (!cell || isBlank(cell.v)) return {};
  if (cell.f && typeof cell.v !== "number") return { error: "formula" };
  if (typeof cell.v !== "number") return { error: "type" };
  if (!Number.isFinite(cell.v)) return { error: "infinite" };
  if (cell.v < 0) return { error: "negative" };
  if (!Number.isInteger(cell.v)) return { error: "fraction" };
  return { value: cell.v };
}

export function validateAllocationWorksheet(sheet: Worksheet): ParsedSource {
  const issues: ValidationIssue[] = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] ?? "A1:A1");
  const endColumn = range.e.c + 1;
  const endRow = range.e.r + 1;

  const detected = detectLayout(sheet);
  if (!detected.ok) {
    const labels = detected.missingLabels.map((field) => ({ itemCode: "ITEM_CODE", orderType: "ORDER_TYPE", tranType: "TRAN_TYPE" })[field]).join(", ");
    issues.push(issue(
      "metadata-labels", "error",
      `The product label rows were not found: ${labels}. Keep these labels in a column next to the product rows and upload again.`,
      `प्रोडक्ट label पंक्तियाँ नहीं मिलीं: ${labels}। इन labels को प्रोडक्ट पंक्तियों के पास किसी कॉलम में रखें और फ़ाइल दोबारा अपलोड करें।`,
    ));
    return { employees: [], products: [], excludedProducts: [], issues };
  }
  const layout = detected.layout;
  const letter = (column: number) => XLSX.utils.encode_col(column - 1);
  const columnOf = (field: EmployeeField) => layout.columns[field];
  const text = (row: number, field: EmployeeField) => {
    const column = columnOf(field);
    return column === undefined ? "" : textValue(valueAt(sheet, row, column));
  };

  for (const field of Object.keys(EMPLOYEE_HEADER_ALIASES) as EmployeeField[]) {
    if (field === "name" || columnOf(field) !== undefined) continue;
    const [enLabel, hiLabel] = FIELD_LABELS[field];
    const heading = EMPLOYEE_HEADER_ALIASES[field][0];
    if (REQUIRED_EMPLOYEE_FIELDS.includes(field)) {
      issues.push(issue(
        `header-${field}`, "error",
        `The ${enLabel} column (heading “${heading}”) was not found. Add a column with this heading anywhere on the allocation sheet and upload again.`,
        `${hiLabel} का कॉलम (शीर्षक “${heading}”) नहीं मिला। allocation शीट में इस शीर्षक वाला कॉलम जोड़ें और फ़ाइल दोबारा अपलोड करें।`,
      ));
    } else {
      issues.push(issue(
        `header-${field}`, "warning",
        `The optional ${enLabel} column (heading “${heading}”) was not found.${field === "code" ? " Duplicate employee rows cannot be checked." : " It will be left blank in the ERP files."}`,
        `वैकल्पिक ${hiLabel} कॉलम (शीर्षक “${heading}”) नहीं मिला।${field === "code" ? " दोहराई गई कर्मचारी पंक्तियों की जाँच नहीं हो सकेगी।" : " ERP फ़ाइलों में यह खाली रहेगा।"}`,
      ));
    }
  }

  if (layout.firstProductColumn === null) {
    issues.push(issue(
      "no-product-area", "error",
      `The product allocation area was not found. Product columns need values in the ITEM_CODE, ORDER_TYPE or TRAN_TYPE rows to the right of column ${letter(layout.labelColumn)}.`,
      `प्रोडक्ट allocation क्षेत्र नहीं मिला। कॉलम ${letter(layout.labelColumn)} के दाईं ओर प्रोडक्ट कॉलम की ITEM_CODE, ORDER_TYPE या TRAN_TYPE पंक्तियों में जानकारी होनी चाहिए।`,
      { cell: `${letter(layout.labelColumn)}${layout.itemCodeRow}` },
    ));
    return { employees: [], products: [], excludedProducts: [], issues, layout };
  }

  const candidateColumns: number[] = [];
  for (let column = layout.firstProductColumn; column <= endColumn; column += 1) {
    if (!layout.employeeColumns.has(column)) candidateColumns.push(column);
  }

  const employeeRows: number[] = [];
  for (let row = layout.firstDataRow; row <= endRow; row += 1) {
    if (hasAnyRowData(sheet, row, candidateColumns, layout) && !isTotalRow(sheet, row, layout)) employeeRows.push(row);
  }

  const products: ProductColumn[] = [];

  for (const column of candidateColumns) {
    const name = textValue(valueAt(sheet, layout.headerRow, column));
    const canonicalCode = textValue(valueAt(sheet, layout.itemCodeRow, column));
    const fallbackCode = layout.fallbackItemCodeRow === null ? "" : textValue(valueAt(sheet, layout.fallbackItemCodeRow, column));
    const orderType = normalize(valueAt(sheet, layout.orderTypeRow, column));
    const tranType = normalize(valueAt(sheet, layout.tranTypeRow, column));
    const hasAllocationCell = employeeRows.some((row) => !isBlank(valueAt(sheet, row, column)));
    const hasMetadata = Boolean(canonicalCode || fallbackCode || orderType || tranType);

    if (!name && !hasMetadata && !hasAllocationCell) continue;
    if (isTotalColumn(name)) continue;
    if (isInformationColumn(sheet, employeeRows, column)) {
      const columnLetter = letter(column);
      issues.push(issue(
        `information-column-${columnLetter}`, "warning",
        `Column ${columnLetter}${name || canonicalCode ? ` (${name || canonicalCode})` : ""} contains text, not quantities, so it was skipped.`,
        `कॉलम ${columnLetter}${name || canonicalCode ? ` (${name || canonicalCode})` : ""} में मात्रा नहीं बल्कि text है, इसलिए इसे छोड़ दिया गया।`,
        { cell: `${columnLetter}${layout.headerRow}`, productName: name || canonicalCode },
      ));
      continue;
    }

    const product: ProductColumn = {
      columnIndex: column,
      columnLetter: XLSX.utils.encode_col(column - 1),
      name,
      itemCode: canonicalCode || fallbackCode,
      orderType,
      tranType,
      itemCodeSourceRow: canonicalCode ? layout.itemCodeRow : fallbackCode ? layout.fallbackItemCodeRow : null,
    };
    products.push(product);
  }

  if (products.length === 0) {
    issues.push(issue(
      "no-products", "error",
      `No product columns were found from column ${letter(layout.firstProductColumn)} onward. Restore the product names and metadata in the original file.`,
      `कॉलम ${letter(layout.firstProductColumn)} के बाद कोई प्रोडक्ट कॉलम नहीं मिला। मूल फ़ाइल में प्रोडक्ट नाम और जानकारी ठीक करें।`,
      { cell: `${letter(layout.firstProductColumn)}${layout.headerRow}` },
    ));
  }

  const excludedProducts = products
    .filter((product) => !product.itemCode || !product.orderType || !product.tranType)
    .map((product) => ({
      columnLetter: product.columnLetter,
      name: product.name || product.itemCode || `Column ${product.columnLetter}`,
      hasItemCode: Boolean(product.itemCode),
      hasOrderType: Boolean(product.orderType),
      hasTranType: Boolean(product.tranType),
    }));
  const excludedColumns = new Set(excludedProducts.map((product) => product.columnLetter));
  const employees: EmployeeAllocation[] = [];
  const seenEmployeeCodes = new Map<string, number>();

  for (const row of employeeRows) {
    const allocations: EmployeeAllocation["allocations"] = [];
    const employeeCode = text(row, "code");
    const employeeName = text(row, "name");

    for (const product of products) {
      if (excludedColumns.has(product.columnLetter)) continue;
      const address = `${product.columnLetter}${row}`;
      const checked = validateQuantity(cellAt(sheet, row, product.columnIndex));
      if (checked.error) {
        const reason = checked.error === "negative" ? "negative" : checked.error === "fraction" ? "a fraction" : checked.error === "formula" ? "a formula without a saved numeric result" : "not a valid whole number";
        const hindiReason = checked.error === "negative" ? "ऋणात्मक" : checked.error === "fraction" ? "दशमलव" : checked.error === "formula" ? "ऐसा फ़ॉर्मूला जिसका संख्यात्मक परिणाम सेव नहीं है" : "मान्य पूर्ण संख्या नहीं";
        issues.push(issue(
          `quantity-${address}`, "error",
          `Quantity at ${address} for ${product.name || "an unnamed product"} is ${reason}. Enter a nonnegative whole number and upload again.`,
          `${product.name || "बिना नाम वाले प्रोडक्ट"} के लिए ${address} में मात्रा ${hindiReason} है। शून्य या धनात्मक पूर्ण संख्या भरें और फ़ाइल दोबारा अपलोड करें।`,
          { cell: address, sourceRow: row, employeeCode, employeeName, productName: product.name, quantity: cellAt(sheet, row, product.columnIndex)?.v },
        ));
      } else if (checked.value && checked.value > 0) {
        allocations.push({ product, quantity: checked.value });
      }
    }

    if (allocations.length === 0 && !products.some((product) => validateQuantity(cellAt(sheet, row, product.columnIndex)).error)) continue;

    const context = { sourceRow: row, employeeCode, employeeName };
    const requiredFields: EmployeeField[] = ["code", ...REQUIRED_EMPLOYEE_FIELDS];
    for (const field of requiredFields) {
      const column = columnOf(field);
      if (column === undefined || text(row, field)) continue;
      const [enLabel, hiLabel] = FIELD_LABELS[field];
      const cell = `${letter(column)}${row}`;
      issues.push(issue(
        `missing-${cell}`, "error",
        `${enLabel[0].toUpperCase()}${enLabel.slice(1)} is missing at ${cell}. Enter it in the original Excel file and upload again.`,
        `${cell} में ${hiLabel} नहीं मिला। मूल Excel फ़ाइल में सही जानकारी भरें और फ़ाइल दोबारा अपलोड करें।`,
        { ...context, cell },
      ));
    }

    const normalizedCode = normalize(employeeCode);
    if (normalizedCode) {
      const previousRow = seenEmployeeCodes.get(normalizedCode);
      if (previousRow) {
        issues.push(issue(
          `duplicate-${row}`, "error",
          `Employee code ${employeeCode} is repeated in rows ${previousRow} and ${row}. Keep one unambiguous row per employee and upload again.`,
          `कर्मचारी कोड ${employeeCode}, पंक्ति ${previousRow} और ${row} में दो बार है। प्रत्येक कर्मचारी के लिए केवल एक स्पष्ट पंक्ति रखें और फ़ाइल दोबारा अपलोड करें।`,
          { ...context, cell: `${letter(columnOf("code")!)}${row}` },
        ));
      } else seenEmployeeCodes.set(normalizedCode, row);
    }

    employees.push({
      sourceRow: row,
      employeeCode,
      employeeName,
      shipSite: text(row, "shipSite"),
      deliverySite: text(row, "deliverySite"),
      remarks: text(row, "remarks"),
      salesPerson: text(row, "salesPerson"),
      projectCode: text(row, "projectCode"),
      allocations,
    });
  }

  for (const product of products) {
    const label = product.name || product.itemCode;
    const metadataProblems: [boolean, string, string, number][] = [
      [!product.itemCode, `Product code is missing for ${label || "the product"}`, `${label || "प्रोडक्ट"} का प्रोडक्ट कोड नहीं मिला`, layout.itemCodeRow],
      [!product.orderType, `ORDER_TYPE is missing for ${label || "the product"}`, `${label || "प्रोडक्ट"} का ORDER_TYPE नहीं मिला`, layout.orderTypeRow],
      [!product.tranType, `TRAN_TYPE is missing for ${label || "the product"}`, `${label || "प्रोडक्ट"} का TRAN_TYPE नहीं मिला`, layout.tranTypeRow],
    ];

    for (const [missing, enText, hiText, row] of metadataProblems) {
      if (!missing) continue;
      const cell = `${product.columnLetter}${row}`;
      issues.push(issue(
        `metadata-${cell}`, "warning",
        `${enText}. This product will not be included in either ERP file.`,
        `${hiText}। यह प्रोडक्ट किसी भी ERP फ़ाइल में शामिल नहीं किया जाएगा।`,
        { cell, productName: product.name },
      ));
    }

    if (!excludedColumns.has(product.columnLetter)) {
      const validOrder = product.orderType === "PS" || product.orderType === "LI";
      const validTran = product.tranType === "PS" || product.tranType === "LI";
      if (!validOrder || !validTran || product.orderType !== product.tranType) {
        issues.push(issue(
          `classification-${product.columnLetter}`, "error",
          `${product.name || product.itemCode || "The product"} has conflicting or unsupported classification: ORDER_TYPE=${product.orderType || "blank"}, TRAN_TYPE=${product.tranType || "blank"}. Both must be PS or both must be LI.`,
          `${product.name || product.itemCode || "प्रोडक्ट"} की classification मेल नहीं खाती: ORDER_TYPE=${product.orderType || "खाली"}, TRAN_TYPE=${product.tranType || "खाली"}। दोनों PS या दोनों LI होने चाहिए।`,
          { cell: `${product.columnLetter}${layout.orderTypeRow}:${product.columnLetter}${layout.tranTypeRow}`, productName: product.name || product.itemCode },
        ));
      }
    }
  }

  if (employeeRows.length === 0 && !issues.some((item) => item.id === "no-products")) {
    issues.push(issue(
      "no-employees", "error",
      `No employee allocation rows with positive quantities were found from row ${layout.firstDataRow} onward.`,
      `पंक्ति ${layout.firstDataRow} के बाद धनात्मक मात्रा वाली कोई कर्मचारी allocation पंक्ति नहीं मिली।`,
      { cell: `${letter(layout.labelColumn)}${layout.firstDataRow}` },
    ));
  }

  return { employees, products, excludedProducts, issues, layout };
}
