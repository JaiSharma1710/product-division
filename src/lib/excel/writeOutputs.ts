import * as XLSX from "xlsx";
import JSZip from "jszip";
import { DATE_FORMAT, DETAIL_COLUMNS, HEADER_COLUMNS } from "@/lib/excel/templateConfig";
import type { GeneratedFiles, OrderWorkbookData, ValidationIssue } from "@/lib/types";

const XLS_MIME = "application/vnd.ms-excel";

function makeSheet<T extends object>(columns: readonly (keyof T)[], rows: T[], dateColumns: readonly number[]): XLSX.WorkSheet {
  const data: unknown[][] = [columns.map(String), ...rows.map((row) => columns.map((column) => row[column]))];
  const sheet = XLSX.utils.aoa_to_sheet(data, { cellDates: false });
  sheet["!cols"] = columns.map((column) => ({ wch: Math.max(12, Math.min(24, String(column).length + 2)) }));

  for (let row = 1; row <= rows.length; row += 1) {
    for (const column of dateColumns) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })] as XLSX.CellObject | undefined;
      if (cell) {
        cell.t = "n";
        cell.z = DATE_FORMAT;
      }
    }
  }
  return sheet;
}

export function writeOrderWorkbook(data: OrderWorkbookData): Uint8Array {
  if (data.headers.length > 65_535 || data.details.length > 65_535) {
    throw new Error("BIFF8_ROW_LIMIT");
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, makeSheet(HEADER_COLUMNS, data.headers, [1, 4, 5]), "Header");
  XLSX.utils.book_append_sheet(workbook, makeSheet(DETAIL_COLUMNS, data.details, [7, 8]), "Detail");
  const raw = XLSX.write(workbook, { type: "array", bookType: "biff8", cellStyles: true }) as ArrayBuffer;
  const bytes = new Uint8Array(raw);
  verifyOrderWorkbook(bytes, data);
  return bytes;
}

function sheetRows(sheet: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" }) as unknown[][];
}

export function verifyOrderWorkbook(bytes: Uint8Array, expected: OrderWorkbookData): void {
  const workbook = XLSX.read(bytes, { type: "array", cellDates: false });
  if (workbook.SheetNames.length !== 2 || workbook.SheetNames[0] !== "Header" || workbook.SheetNames[1] !== "Detail") {
    throw new Error("WORKSHEET_STRUCTURE");
  }
  const headerRows = sheetRows(workbook.Sheets.Header);
  const detailRows = sheetRows(workbook.Sheets.Detail);
  if (JSON.stringify(headerRows[0]) !== JSON.stringify(HEADER_COLUMNS) || JSON.stringify(detailRows[0]) !== JSON.stringify(DETAIL_COLUMNS)) {
    throw new Error("COLUMN_STRUCTURE");
  }
  if (headerRows.length - 1 !== expected.headers.length || detailRows.length - 1 !== expected.details.length) {
    throw new Error("ROW_COUNT");
  }

  const ids = new Set(headerRows.slice(1).map((row) => row[0]));
  for (let index = 1; index < detailRows.length; index += 1) {
    const row = detailRows[index];
    const expectedRow = expected.details[index - 1];
    if (!ids.has(row[0])) throw new Error("FOREIGN_KEY");
    if (row[2] !== expectedRow.ITEM_CODE || row[3] !== expectedRow.QTY_ORDER || row[4] !== expectedRow.QTY_CONFIRM) throw new Error("DETAIL_VALUE");
    if (typeof row[3] !== "number" || typeof row[4] !== "number" || typeof row[7] !== "number" || typeof row[8] !== "number") throw new Error("DETAIL_CELL_TYPE");
  }
  for (let index = 1; index < headerRows.length; index += 1) {
    const row = headerRows[index];
    if (typeof row[1] !== "number" || typeof row[4] !== "number" || typeof row[5] !== "number" || typeof row[12] !== "number") throw new Error("HEADER_CELL_TYPE");
  }
}

export async function writeAllOutputs(ps: OrderWorkbookData, li: OrderWorkbookData): Promise<GeneratedFiles> {
  const psBytes = writeOrderWorkbook(ps);
  const liBytes = writeOrderWorkbook(li);
  const zip = new JSZip();
  zip.file("DO_UPLOAD_PS.xls", psBytes);
  zip.file("DO_UPLOAD_LI.xls", liBytes);
  const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { ps: psBytes, li: liBytes, zip: zipBlob };
}

export function workbookBlob(bytes: Uint8Array): Blob {
  return new Blob([bytes.slice().buffer], { type: XLS_MIME });
}

export function writeValidationReport(issues: ValidationIssue[]): Uint8Array {
  const workbook = XLSX.utils.book_new();
  const rows = issues.map((item, index) => ({
    "Issue No. / क्रमांक": index + 1,
    "Severity / स्थिति": item.severity === "error" ? "Error / त्रुटि" : "Warning / चेतावनी",
    "Sheet / शीट": item.sheet ?? "",
    "Cell / सेल": item.cell ?? "",
    "Source Row / स्रोत पंक्ति": item.sourceRow ?? "",
    "Employee Code / कर्मचारी कोड": item.employeeCode ?? "",
    "Employee Name / कर्मचारी नाम": item.employeeName ?? "",
    "Product / प्रोडक्ट": item.productName ?? "",
    "Quantity / मात्रा": item.quantity ?? "",
    "English / अंग्रेज़ी": item.en,
    हिन्दी: item.hi,
  }));
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet["!cols"] = [10, 18, 14, 14, 12, 18, 24, 28, 12, 70, 70].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(workbook, sheet, "Validation Errors");
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer);
}
