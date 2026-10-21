import * as XLSX from "xlsx";
import { EMPLOYEE_HEADER_ALIASES, SOURCE } from "@/lib/excel/templateConfig";
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

function matchesHeader(value: unknown, aliases: readonly string[]): boolean {
  const normalized = normalize(value);
  return aliases.some((alias) => normalize(alias) === normalized);
}

function isTotalColumn(name: string): boolean {
  const normalized = normalize(name);
  return /^(CALCULATED )?TOTAL\b/.test(normalized) || /\bTOTAL (SAMPLES?|PRODUCTS?|QTY|QUANTITY)\b/.test(normalized);
}

function hasAnyRowData(sheet: Worksheet, row: number, endColumn: number): boolean {
  const employeeColumns = Object.values(SOURCE.EMPLOYEE);
  return employeeColumns.some((column) => !isBlank(valueAt(sheet, row, column))) ||
    Array.from({ length: Math.max(0, endColumn - SOURCE.FIRST_PRODUCT_COLUMN + 1) }, (_, index) => SOURCE.FIRST_PRODUCT_COLUMN + index)
      .some((column) => !isBlank(valueAt(sheet, row, column)));
}

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

  const headerChecks = [
    [SOURCE.EMPLOYEE.code, EMPLOYEE_HEADER_ALIASES.code, "employee code / कर्मचारी कोड"],
    [SOURCE.EMPLOYEE.shipSite, EMPLOYEE_HEADER_ALIASES.shipSite, "shipping site / भेजने की साइट"],
    [SOURCE.EMPLOYEE.deliverySite, EMPLOYEE_HEADER_ALIASES.deliverySite, "delivery site / डिलीवरी साइट"],
    [SOURCE.EMPLOYEE.remarks, EMPLOYEE_HEADER_ALIASES.remarks, "remarks / टिप्पणी"],
    [SOURCE.EMPLOYEE.salesPerson, EMPLOYEE_HEADER_ALIASES.salesPerson, "salesperson / सेल्सपर्सन"],
    [SOURCE.EMPLOYEE.projectCode, EMPLOYEE_HEADER_ALIASES.projectCode, "project code / प्रोजेक्ट कोड"],
  ] as const;

  for (const [column, aliases, label] of headerChecks) {
    const address = XLSX.utils.encode_cell({ r: SOURCE.HEADER_ROW - 1, c: column - 1 });
    if (!matchesHeader(valueAt(sheet, SOURCE.HEADER_ROW, column), aliases)) {
      issues.push(issue(
        `header-${address}`,
        "error",
        `The ${label.split(" / ")[0]} heading is missing or moved at cell ${address}. Restore the original allocation sheet layout and upload again.`,
        `${address} सेल में ${label.split(" / ")[1]} का शीर्षक नहीं मिला या उसकी जगह बदल गई है। allocation शीट का मूल प्रारूप ठीक करके फ़ाइल दोबारा अपलोड करें।`,
        { cell: address },
      ));
    }
  }

  if (endColumn < SOURCE.FIRST_PRODUCT_COLUMN) {
    issues.push(issue(
      "no-product-area", "error",
      "The product allocation area was not found. Products must begin from column R.",
      "प्रोडक्ट allocation क्षेत्र नहीं मिला। प्रोडक्ट कॉलम R से शुरू होने चाहिए।",
      { cell: "R3" },
    ));
    return { employees: [], products: [], excludedProducts: [], issues };
  }

  const employeeRows: number[] = [];
  for (let row = SOURCE.FIRST_EMPLOYEE_ROW; row <= endRow; row += 1) {
    if (hasAnyRowData(sheet, row, endColumn)) employeeRows.push(row);
  }

  const products: ProductColumn[] = [];

  for (let column = SOURCE.FIRST_PRODUCT_COLUMN; column <= endColumn; column += 1) {
    const name = textValue(valueAt(sheet, SOURCE.HEADER_ROW, column));
    const canonicalCode = textValue(valueAt(sheet, SOURCE.ITEM_CODE_ROW, column));
    const fallbackCode = textValue(valueAt(sheet, SOURCE.FALLBACK_ITEM_CODE_ROW, column));
    const orderType = normalize(valueAt(sheet, SOURCE.ORDER_TYPE_ROW, column));
    const tranType = normalize(valueAt(sheet, SOURCE.TRAN_TYPE_ROW, column));
    const hasAllocationCell = employeeRows.some((row) => !isBlank(valueAt(sheet, row, column)));
    const hasMetadata = Boolean(canonicalCode || fallbackCode || orderType || tranType);

    if (!name && !hasMetadata && !hasAllocationCell) continue;
    if (isTotalColumn(name)) continue;

    const product: ProductColumn = {
      columnIndex: column,
      columnLetter: XLSX.utils.encode_col(column - 1),
      name,
      itemCode: canonicalCode || fallbackCode,
      orderType,
      tranType,
      itemCodeSourceRow: canonicalCode ? 4 : fallbackCode ? 2 : null,
    };
    products.push(product);
  }

  if (products.length === 0) {
    issues.push(issue(
      "no-products", "error",
      "No product columns were found from column R onward. Restore the product names and metadata in the original file.",
      "कॉलम R के बाद कोई प्रोडक्ट कॉलम नहीं मिला। मूल फ़ाइल में प्रोडक्ट नाम और जानकारी ठीक करें।",
      { cell: "R3" },
    ));
  }

  const excludedProducts = products
    .filter((product) => !product.name || !product.itemCode || !product.orderType || !product.tranType)
    .map((product) => ({
      columnLetter: product.columnLetter,
      name: product.name || `Column ${product.columnLetter}`,
      hasItemCode: Boolean(product.itemCode),
      hasOrderType: Boolean(product.orderType),
      hasTranType: Boolean(product.tranType),
    }));
  const excludedColumns = new Set(excludedProducts.map((product) => product.columnLetter));
  const employees: EmployeeAllocation[] = [];
  const seenEmployeeCodes = new Map<string, number>();

  for (const row of employeeRows) {
    const allocations: EmployeeAllocation["allocations"] = [];
    const employeeCode = textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.code));
    const employeeName = textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.name));

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
    const requiredValues = [
      [employeeCode, "employee code", "कर्मचारी कोड", `B${row}`],
      [textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.shipSite)), "shipping site", "भेजने की साइट", `E${row}`],
      [textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.deliverySite)), "delivery site", "डिलीवरी साइट", `F${row}`],
      [textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.salesPerson)), "salesperson code", "सेल्सपर्सन कोड", `H${row}`],
    ] as const;

    for (const [value, enLabel, hiLabel, cell] of requiredValues) {
      if (!value) issues.push(issue(
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
          { ...context, cell: `B${row}` },
        ));
      } else seenEmployeeCodes.set(normalizedCode, row);
    }

    employees.push({
      sourceRow: row,
      employeeCode,
      employeeName,
      shipSite: textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.shipSite)),
      deliverySite: textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.deliverySite)),
      remarks: textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.remarks)),
      salesPerson: textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.salesPerson)),
      projectCode: textValue(valueAt(sheet, row, SOURCE.EMPLOYEE.projectCode)),
      allocations,
    });
  }

  for (const product of products) {
    const metadataProblems: [boolean, string, string, number][] = [
      [!product.name, "Product name is missing", "प्रोडक्ट का नाम नहीं मिला", SOURCE.HEADER_ROW],
      [!product.itemCode, `Product code is missing for ${product.name || "the product"}`, `${product.name || "प्रोडक्ट"} का प्रोडक्ट कोड नहीं मिला`, SOURCE.ITEM_CODE_ROW],
      [!product.orderType, `ORDER_TYPE is missing for ${product.name || "the product"}`, `${product.name || "प्रोडक्ट"} का ORDER_TYPE नहीं मिला`, SOURCE.ORDER_TYPE_ROW],
      [!product.tranType, `TRAN_TYPE is missing for ${product.name || "the product"}`, `${product.name || "प्रोडक्ट"} का TRAN_TYPE नहीं मिला`, SOURCE.TRAN_TYPE_ROW],
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
          `${product.name || "The product"} has conflicting or unsupported classification: ORDER_TYPE=${product.orderType || "blank"}, TRAN_TYPE=${product.tranType || "blank"}. Both must be PS or both must be LI.`,
          `${product.name || "प्रोडक्ट"} की classification मेल नहीं खाती: ORDER_TYPE=${product.orderType || "खाली"}, TRAN_TYPE=${product.tranType || "खाली"}। दोनों PS या दोनों LI होने चाहिए।`,
          { cell: `${product.columnLetter}${SOURCE.ORDER_TYPE_ROW}:${product.columnLetter}${SOURCE.TRAN_TYPE_ROW}`, productName: product.name },
        ));
      }
    }
  }

  if (employeeRows.length === 0 && !issues.some((item) => item.id === "no-products")) {
    issues.push(issue(
      "no-employees", "error",
      "No employee allocation rows with positive quantities were found from row 7 onward.",
      "पंक्ति 7 के बाद धनात्मक मात्रा वाली कोई कर्मचारी allocation पंक्ति नहीं मिली।",
      { cell: "B7" },
    ));
  }

  return { employees, products, excludedProducts, issues };
}
