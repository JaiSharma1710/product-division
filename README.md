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

The converter reads only the worksheet named `allocation` (case-insensitive, surrounding spaces ignored):

- row 3: employee headings and product names;
- row 4: product `ITEM_CODE`;
- row 5: product `ORDER_TYPE`;
- row 6: product `TRAN_TYPE`;
- row 7 onward: employee allocations;
- employee fields: B, E, F, G, H, and I; employee name in D is optional context and its heading is not required;
- product search starts at column R and continues to the actual used range;
- calculated total columns are excluded only when their heading clearly identifies them as totals;
- hidden columns are read normally.

### Constrained item-code fallback

Row 4 is always the primary `ITEM_CODE` location. Row 2 is used only when the same product's row-4 cell is blank and the row-2 cell contains a saved scalar value. The fallback never infers a code from the product name, column, or prefix.

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
