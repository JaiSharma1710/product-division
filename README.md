# Excel to ERP Upload Generator

A single-page, browser-only Next.js tool that converts the `allocation` worksheet from an allocation workbook into:

- `DO_UPLOAD_PS.xls`
- `DO_UPLOAD_LI.xls`

Both outputs are genuine BIFF8 Excel files with `Header` and `Detail` worksheets. A ZIP download contains both files.

## Privacy and architecture

All workbook reading, validation, transformation, writing, read-back verification, and ZIP creation happen in the browser. The app has:

- no API routes or server actions;
- no database or authentication;
- no upload or file storage service;
- no analytics or tracking;
- no localStorage or sessionStorage use.

The production build is a static export in `out/`.

## Local development

Requirements: Node.js 20 or newer and npm.

```bash
npm install
npm run dev
```

Open `http://localhost:3000` and choose a `.xls` or `.xlsx` allocation workbook.

SheetJS Community Edition is pinned to version `0.20.3` from the official SheetJS CDN tarball. It is intentionally not installed from the obsolete npm registry release.

## Checks

```bash
npm run lint
npm run typecheck
npm test
npx playwright install chromium   # first time only
npm run test:e2e
npm run build
```

Tests use small synthetic workbooks and contain no production employee data.

## Static deployment on Vercel

Import the repository into Vercel. The standard Next.js build command (`npm run build`) produces the static `out/` directory because `next.config.ts` sets `output: "export"`. No environment variables or runtime services are required.

The site may also be hosted on any static file host by publishing `out/`.

## Source workbook contract

The converter reads the worksheet named `allocation` (case-insensitive, surrounding spaces ignored). If no sheet has that name, it uses the first sheet that has the metadata labels below.

Nothing is read from a fixed row or column. The layout is found from headings and labels in the top 50 rows, so columns may be moved, reordered, or inserted:

- **Metadata rows**: the cells labelled `ITEM_CODE`, `ORDER_TYPE`, and `TRAN_TYPE` mark the product metadata rows. The column holding these labels is the *label column*.
- **Heading row**: the row with the most employee headings. Product names, when present, are read from this row.
- **Employee columns**, found by heading (matching ignores case, spaces, and punctuation; aliases are in `templateConfig.ts`):
  - required: `SITE_CODE__SHIP`, `SITE_CODE__DLV`, `SALES_PERS`;
  - optional: `REMARKS`, `PROJ_CODE`, employee code (`Emp Code`, `EMPLOYEE CODE`, …), and employee name. A missing optional column produces a warning and stays blank in the ERP files. Without an employee-code column, duplicate rows cannot be checked. Without a name heading, names are read from the label column.
- **Employee rows** start on the row below the lowest heading or metadata row. Rows whose name or code reads `Total` are skipped.
- **Product columns** start at the first column right of the label column that has a value in a metadata row, and continue to the end of the used range. Employee columns are skipped.
- Product names are optional; the item code identifies the product when there is no name.
- Calculated total columns are excluded only when their heading clearly identifies them as totals.
- Columns whose filled cells are all non-numeric text (for example `EMAIL-ID`) are skipped with a warning.
- Hidden columns are read normally.

### Constrained item-code fallback

The `ITEM_CODE` row is always the primary item-code location. The row directly above the heading row is used only when the same product's `ITEM_CODE` cell is blank and the fallback cell contains a saved scalar value. The fallback never infers a code from the product name, column, or prefix.

A product missing `ITEM_CODE`, `ORDER_TYPE`, or `TRAN_TYPE` is excluded from both ERP outputs instead of blocking the conversion. The results screen shows every excluded product in a bilingual table with Yes/No status for each required metadata field. Quantities in excluded columns are not processed. Populated but conflicting or unsupported classifications still block generation because the app cannot safely choose PS or LI.

Formula allocation cells are accepted only when the workbook contains a saved finite whole-number result. SheetJS Community Edition does not calculate Excel formulas in the browser; formulas without a saved numeric result are blocked with a correction message.

## Output safety checks

Before downloads are exposed, each generated `.xls` file is reopened and checked for:

- exact worksheet names and order;
- exact column names and order;
- expected header and detail row counts;
- valid `DIST_ORDER` relationships;
- unchanged item codes and quantities;
- numeric quantity and Excel date cell types.

Any blocking source or output verification error prevents all ERP downloads. The optional `Validation_Errors.xlsx` report is generated locally for correction.

Each PS and LI download has its own eye button. Preview opens in a focused full-page workbook view with a clear Back button, a Download button, and scrollable spreadsheet-style `Header` and `Detail` tables.

## Compatibility limitations

- BIFF8 `.xls` is limited to 65,536 rows per worksheet (65,535 data rows after the heading). Oversized output is blocked and must be split.
- The app does not copy macros, VBA, external links, formulas, or workbook-specific binary records into outputs. The stated ERP schemas do not require them.
- The three structural reference workbooks were not present in the initial project workspace. The implementation therefore follows the supplied field contract exactly, including `dd/mm/yyyy` numeric dates and explicit blank string cells. Exact style, blank-cell representation, or undocumented ERP binary-record parity should be regression-checked against those reference files when they are available.
- Password-protected or damaged files are not supported.

If an ERP installation requires macros, signed VBA, a particular workbook protection setting, or another proprietary binary record, that requirement is outside SheetJS Community Edition's plain BIFF8 generation and must be confirmed separately rather than assumed.
# product-division
