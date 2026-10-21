import * as XLSX from "xlsx";
import { SOURCE } from "@/lib/excel/templateConfig";
import { validateAllocationWorksheet } from "@/lib/excel/validateSource";
import type { ParsedSource, ValidationIssue } from "@/lib/types";

function workbookIssue(en: string, hi: string): ParsedSource {
  const error: ValidationIssue = { id: "workbook", severity: "error", en, hi };
  return { employees: [], products: [], excludedProducts: [], issues: [error] };
}

export function readSourceWorkbook(data: ArrayBuffer | Uint8Array): ParsedSource {
  try {
    const workbook = XLSX.read(data, { type: "array", cellFormula: true, cellDates: false, dense: false });
    const matchingName = workbook.SheetNames.find((name) => name.trim().toLowerCase() === SOURCE.SHEET_NAME);
    if (!matchingName) {
      return workbookIssue(
        `The “${SOURCE.SHEET_NAME}” worksheet was not found. Use the main allocation workbook and keep the worksheet name unchanged.`,
        `“${SOURCE.SHEET_NAME}” worksheet नहीं मिला। मुख्य allocation workbook चुनें और worksheet का नाम न बदलें।`,
      );
    }
    const sheet = workbook.Sheets[matchingName];
    if (!sheet || !sheet["!ref"]) {
      return workbookIssue(
        "The allocation worksheet is empty or unreadable. Restore it from the original workbook and upload again.",
        "allocation worksheet खाली है या पढ़ा नहीं जा सका। इसे मूल workbook से ठीक करके फ़ाइल दोबारा अपलोड करें।",
      );
    }
    return validateAllocationWorksheet(sheet);
  } catch {
    return workbookIssue(
      "This workbook could not be read. It may be damaged, password-protected, or unsupported. Open it in Excel, remove any password, save as .xls or .xlsx, and upload again.",
      "यह workbook पढ़ा नहीं जा सका। फ़ाइल खराब, पासवर्ड से सुरक्षित या असमर्थित हो सकती है। इसे Excel में खोलें, पासवर्ड हटाएँ, .xls या .xlsx में सेव करें और दोबारा अपलोड करें।",
    );
  }
}
