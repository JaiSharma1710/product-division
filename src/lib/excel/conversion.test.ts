import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import { convertWorkbook } from "@/lib/excel/convert";
import { getIndiaCalendarDate, toExcelSerial } from "@/lib/date";
import { DETAIL_COLUMNS, ERP_CONSTANTS, HEADER_COLUMNS } from "@/lib/excel/templateConfig";
import { makeSourceFixture } from "@/lib/excel/testFixtures";

const NOW = new Date("2026-10-20T02:00:00.000Z");

async function valid(options = {}) {
  const result = await convertWorkbook(makeSourceFixture(options), NOW);
  expect(result.issues.filter((issue) => issue.severity === "error")).toHaveLength(0);
  expect(result.generated).toBeDefined();
  return result;
}

describe("Excel to ERP conversion", () => {
  it("produces readable PS and LI BIFF8 files with exact sheets and columns", async () => {
    const result = await valid();
    for (const bytes of [result.generated!.ps, result.generated!.li]) {
      expect(Array.from(bytes.slice(0, 8))).toEqual([208, 207, 17, 224, 161, 177, 26, 225]);
      const workbook = XLSX.read(bytes, { type: "array" });
      expect(workbook.SheetNames).toEqual(["Header", "Detail"]);
      expect(XLSX.utils.sheet_to_json(workbook.Sheets.Header, { header: 1 })[0]).toEqual(HEADER_COLUMNS);
      expect(XLSX.utils.sheet_to_json(workbook.Sheets.Detail, { header: 1 })[0]).toEqual(DETAIL_COLUMNS);
    }
  });

  it("keeps sequential DIST_ORDER values stable between PS and LI", async () => {
    const result = await valid({ employeeCount: 3 });
    const expected = ["DIST001", "DIST002", "DIST003"];
    expect(result.ps!.headers.map((row) => row.DIST_ORDER)).toEqual(expected);
    expect(result.li!.headers.map((row) => row.DIST_ORDER)).toEqual(expected);
  });

  it("resets LINE_NO per order and preserves quantities and confirmations", async () => {
    const result = await valid();
    for (const output of [result.ps!, result.li!]) {
      expect(output.details.filter((row) => row.DIST_ORDER === "DIST001").map((row) => row.LINE_NO)).toEqual([1]);
      for (const row of output.details) expect(row.QTY_CONFIRM).toBe(row.QTY_ORDER);
    }
    expect(result.ps!.details.map((row) => row.QTY_ORDER)).toEqual([1, 2]);
    expect(result.li!.details.map((row) => row.QTY_ORDER)).toEqual([2, 3]);
  });

  it("skips zero and blank products and separates PS from LI", async () => {
    const result = await valid();
    expect(result.ps!.details.every((row) => row.ITEM_CODE === "PS-001")).toBe(true);
    expect(result.li!.details.every((row) => row.ITEM_CODE === "LI-001")).toBe(true);
    expect(result.ps!.details.some((row) => row.ITEM_CODE === "LI-001")).toBe(false);
    expect(result.li!.details.some((row) => row.ITEM_CODE === "PS-001")).toBe(false);
  });

  it.each([
    ["conflicting classification", { conflictingClassification: true }, "classification-R"],
    ["negative quantity", { invalidQuantity: -1 }, "quantity-R7"],
    ["fractional quantity", { invalidQuantity: 1.5 }, "quantity-R7"],
    ["nonnumeric quantity", { invalidQuantity: "ten" }, "quantity-R7"],
    ["duplicate employee", { duplicateEmployee: true }, "duplicate-8"],
  ])("blocks all downloads for %s", async (_label, options, issueId) => {
    const result = await convertWorkbook(makeSourceFixture(options), NOW);
    expect(result.generated).toBeUndefined();
    expect(result.issues.some((issue) => issue.id === issueId && issue.severity === "error")).toBe(true);
  });

  it.each([
    ["item code", { missingItemCode: true }, "metadata-R4"],
    ["classification", { missingClassification: true }, "metadata-R5"],
  ])("excludes a product with missing %s without blocking downloads", async (_label, options, issueId) => {
    const result = await valid(options);
    expect(result.excludedProducts.some((product) => product.columnLetter === "R")).toBe(true);
    expect(result.issues.some((issue) => issue.id === issueId && issue.severity === "warning")).toBe(true);
    expect(result.ps?.details).toHaveLength(0);
    expect(result.li?.details).toHaveLength(2);
  });

  it("ignores an invalid allocation quantity when its product metadata is incomplete", async () => {
    const result = await valid({ missingItemCode: true, invalidQuantity: "not a number" });
    expect(result.generated).toBeDefined();
    expect(result.issues.some((issue) => issue.id === "quantity-R7")).toBe(false);
  });

  it("processes hidden product columns and a non-hardcoded employee count", async () => {
    const result = await valid({ employeeCount: 4, hiddenProductColumn: true });
    expect(result.summary?.employees).toBe(4);
    expect(result.ps?.details).toHaveLength(4);
  });

  it("uses the IST generation date in every real date cell", async () => {
    const result = await valid();
    const expected = toExcelSerial(getIndiaCalendarDate(NOW));
    for (const header of [...result.ps!.headers, ...result.li!.headers]) {
      expect([header.ORDER_DATE, header.SHIP_DATE, header.DUE_DATE]).toEqual([expected, expected, expected]);
    }
    for (const detail of [...result.ps!.details, ...result.li!.details]) {
      expect([detail.DUE_DATE, detail.SHIP_DATE]).toEqual([expected, expected]);
    }
  });

  it("keeps optional project codes blank and remarks in both outputs", async () => {
    const result = await valid({ missingProjectCode: true });
    expect(result.ps!.headers[0].PROJ_CODE).toBe("");
    expect(result.li!.headers[0].PROJ_CODE).toBe("");
    expect(result.ps!.headers[0].REMARKS).toBe("Remark 1");
    expect(result.li!.headers[0].REMARKS).toBe("Remark 1");
  });

  it("does not require the employee name heading or employee name values", async () => {
    const result = await valid({ missingEmployeeName: true });
    expect(result.summary?.employees).toBe(2);
    expect(result.issues.some((issue) => issue.cell === "D3" || issue.cell === "D7")).toBe(false);
  });

  it("uses the constrained row-2 item code fallback", async () => {
    const result = await valid({ fallbackItemCode: true });
    expect(result.ps!.details[0].ITEM_CODE).toBe("PS-001");
  });

  it("uses all required ERP constants without substituting the ship site", async () => {
    const result = await valid();
    const header = result.ps!.headers[0];
    expect(header.SITE_CODE__SHIP).toBe("SHIP1");
    expect(header.SITE_CODE).toBe(ERP_CONSTANTS.SITE_CODE);
    expect(header.ORDER_TYPE).toBe("PS");
    expect(header.TRAN_TYPE).toBe("PS");
    expect(header.PRICE_LIST).toBe("GTPSA");
  });

  it("places both correctly named real XLS files in the ZIP", async () => {
    const result = await valid();
    const zip = await JSZip.loadAsync(result.generated!.zip);
    expect(Object.keys(zip.files).sort()).toEqual(["DO_UPLOAD_LI.xls", "DO_UPLOAD_PS.xls"]);
    const ps = await zip.file("DO_UPLOAD_PS.xls")!.async("uint8array");
    expect(Array.from(ps.slice(0, 8))).toEqual([208, 207, 17, 224, 161, 177, 26, 225]);
  });

  it("provides English and Hindi for every validation issue", async () => {
    const result = await convertWorkbook(makeSourceFixture({ missingItemCode: true }), NOW);
    expect(result.issues.length).toBeGreaterThan(0);
    expect(result.issues.every((issue) => issue.en.length > 10 && issue.hi.length > 10)).toBe(true);
  });
});
