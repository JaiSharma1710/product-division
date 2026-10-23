import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { convertWorkbook } from "@/lib/excel/convert";

const NOW = new Date("2026-10-20T02:00:00.000Z");

function workbook(rows: unknown[][], sheetName = "allocation"): ArrayBuffer {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([["notes"]]), "cover");
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return XLSX.write(book, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

// Headings on row 1, labels in column B, columns reordered, no employee code or product names.
const SHIFTED = [
  ["S.NO.", "", "SALES_PERS", "Site Code - Ship", "site_code__dlv", "REMARKS", "PS-A", "LI-A"],
  [1, "ITEM_CODE", "", "", "", "", "PSA01", "LIA01"],
  [2, "ORDER_TYPE", "", "", "", "", "PS", "LI"],
  [3, "TRAN_TYPE", "", "", "", "", "PS", "LI"],
  [4, "First Person", "SP1", "SHIP1", "DLV1", "Box 1", 3, 4],
  [5, "Second Person", "SP2", "SHIP2", "DLV2", "Box 2", 0, 6],
];

describe("layout detection by heading", () => {
  it("reads columns by heading when they are moved, reordered, or renamed in case and punctuation", async () => {
    const result = await convertWorkbook(workbook(SHIFTED), NOW);
    expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
    expect(result.ps!.details.map((row) => [row.ITEM_CODE, row.QTY_ORDER])).toEqual([["PSA01", 3]]);
    expect(result.li!.details.map((row) => [row.ITEM_CODE, row.QTY_ORDER])).toEqual([["LIA01", 4], ["LIA01", 6]]);
    expect(result.li!.headers[1]).toMatchObject({ SITE_CODE__SHIP: "SHIP2", SITE_CODE__DLV: "DLV2", SALES_PERS: "SP2", REMARKS: "Box 2", PROJ_CODE: "" });
  });

  it("warns instead of failing when optional columns are missing", async () => {
    const result = await convertWorkbook(workbook(SHIFTED), NOW);
    const warnings = result.issues.filter((issue) => issue.severity === "warning").map((issue) => issue.id);
    expect(warnings).toEqual(expect.arrayContaining(["header-code", "header-projectCode"]));
    expect(result.generated).toBeDefined();
  });

  it("blocks with a clear message when a required heading is missing", async () => {
    const rows = SHIFTED.map((row) => [...row]);
    rows[0][2] = "SOMETHING ELSE";
    const result = await convertWorkbook(workbook(rows), NOW);
    expect(result.generated).toBeUndefined();
    expect(result.issues.some((issue) => issue.id === "header-salesPerson" && issue.severity === "error")).toBe(true);
  });

  it("blocks when the metadata labels are missing", async () => {
    const rows = SHIFTED.map((row) => [...row]);
    rows[2][1] = "";
    const result = await convertWorkbook(workbook(rows), NOW);
    expect(result.generated).toBeUndefined();
    expect(result.issues.find((issue) => issue.id === "metadata-labels")?.en).toContain("ORDER_TYPE");
  });

  it("skips text-only columns such as e-mail addresses with a warning", async () => {
    const rows = SHIFTED.map((row) => [...row.slice(0, 6), "", ...row.slice(6)]);
    [["EMAIL-ID"], ["PS"], ["PS"], ["a@example.com"], ["b@example.com"]].forEach(([value], index) => { rows[index + 1][6] = value; });
    const result = await convertWorkbook(workbook(rows), NOW);
    expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
    expect(result.issues.some((issue) => issue.id === "information-column-G" && issue.severity === "warning")).toBe(true);
    expect(result.ps!.details).toHaveLength(1);
  });

  it("ignores a total row below the employees", async () => {
    const rows = [...SHIFTED, [6, "Total", "", "", "", "", 3, 10]];
    const result = await convertWorkbook(workbook(rows), NOW);
    expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
    expect(result.summary?.employees).toBe(2);
  });

  it("finds the allocation sheet by its layout when it has another name", async () => {
    const result = await convertWorkbook(workbook(SHIFTED, "PMD Allocation Oct"), NOW);
    expect(result.generated).toBeDefined();
    expect(result.summary?.employees).toBe(2);
  });
});
