import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { convertWorkbook } from "@/lib/excel/convert";
import { DETAIL_COLUMNS, HEADER_COLUMNS } from "@/lib/excel/templateConfig";

const REAL_WORKBOOK = "Copy of 01 Input Allocation - September - 05-09 (1).xlsx";

describe("real allocation workbook product check", () => {
  it("converts the supplied workbook and preserves valid output relationships", async () => {
    const file = readFileSync(REAL_WORKBOOK);
    const data = Uint8Array.from(file).buffer;
    const result = await convertWorkbook(data, new Date("2026-10-20T06:00:00.000Z"));

    expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
    expect(result.generated).toBeDefined();
    expect(result.summary?.employees).toBe(32);
    expect(result.excludedProducts).toHaveLength(12);
    expect(result.ps?.headers).toHaveLength(32);
    expect(result.ps?.details).toHaveLength(408);
    expect(result.li?.headers).toHaveLength(32);
    expect(result.li?.details).toHaveLength(140);
    expect(result.summary?.totalLines).toBe(548);
    expect(result.ps?.headers.map((row) => row.DIST_ORDER)).toEqual(result.li?.headers.map((row) => row.DIST_ORDER));

    for (const workbook of [result.ps!, result.li!]) {
      const headerIds = new Set(workbook.headers.map((row) => row.DIST_ORDER));
      expect(workbook.details.every((row) => headerIds.has(row.DIST_ORDER))).toBe(true);
      expect(workbook.details.every((row) => row.QTY_ORDER === row.QTY_CONFIRM)).toBe(true);
      expect(Object.keys(workbook.headers[0])).toEqual(HEADER_COLUMNS);
      expect(Object.keys(workbook.details[0])).toEqual(DETAIL_COLUMNS);
    }
  });
});
