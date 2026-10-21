import { expect, test } from "@playwright/test";
import * as XLSX from "xlsx";

function sourceBuffer(): Buffer {
  const rows: unknown[][] = Array.from({ length: 7 }, () => []);
  const set = (row: number, column: number, value: unknown) => { rows[row - 1][column - 1] = value; };
  set(3, 2, "EMPLOYEE CODE"); set(3, 4, "EMPLOYEE NAME"); set(3, 5, "SITE_CODE__SHIP");
  set(3, 6, "SITE_CODE__DLV"); set(3, 7, "REMARKS"); set(3, 8, "SALES_PERS"); set(3, 9, "PROJ_CODE");
  set(3, 18, "PS Product"); set(4, 18, "PS-001"); set(5, 18, "PS"); set(6, 18, "PS");
  set(3, 19, "LI Product"); set(4, 19, "LI-001"); set(5, 19, "LI"); set(6, 19, "LI");
  set(3, 20, "Missing Metadata Product");
  set(7, 2, "E001"); set(7, 4, "Employee 1"); set(7, 5, "SHIP1"); set(7, 6, "DLV1");
  set(7, 7, "Remark"); set(7, 8, "SP1"); set(7, 18, 2); set(7, 19, 3); set(7, 20, 7);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "allocation");
  return Buffer.from(XLSX.write(workbook, { type: "array", bookType: "xlsx" }));
}

test("shows bilingual errors, accepts a corrected file, and downloads both outputs", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { level: 1 })).toContainText("Excel to ERP");
  const input = page.locator('input[type="file"]');
  await input.setInputFiles({ name: "broken.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from("wrong") });
  await expect(page.getByRole("heading", { name: /file needs correction/i })).toBeVisible();
  await expect(page.getByText(/“allocation” worksheet नहीं मिला/)).toBeVisible();

  await page.getByRole("button", { name: /Upload Another File/i }).click();
  await input.setInputFiles({
    name: "allocation.xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: sourceBuffer(),
  });
  await expect(page.getByRole("heading", { name: /Your files are ready/i })).toBeVisible();

  await page.getByRole("button", { name: /View missing details/i }).click();
  await expect(page.getByRole("dialog", { name: /Excluded product details/i })).toBeVisible();
  await expect(page.getByText("Missing Metadata Product")).toBeVisible();
  await page.getByRole("button", { name: /Close modal/i }).click();

  await page.getByRole("button", { name: /Preview PS file/i }).click();
  await expect(page.getByRole("region", { name: /DO_UPLOAD_PS.xls preview/i })).toBeVisible();
  await expect(page.getByRole("columnheader", { name: "DIST_ORDER" })).toBeVisible();
  await page.getByRole("tab", { name: /Detail/ }).click();
  await expect(page.getByRole("columnheader", { name: "ITEM_CODE" })).toBeVisible();
  await page.getByRole("button", { name: /Back to downloads/i }).click();
  await expect(page.getByRole("heading", { name: /Your files are ready/i })).toBeVisible();

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: /Download Both Files/i }).click();
  const downloaded = await downloadPromise;
  expect(downloaded.suggestedFilename()).toBe("DO_UPLOAD_FILES.zip");
});
