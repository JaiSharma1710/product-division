import * as XLSX from "xlsx";
import { EMPLOYEE_HEADER_ALIASES, METADATA_LABEL_ALIASES, SOURCE, type EmployeeField } from "@/lib/excel/templateConfig";

type Worksheet = XLSX.WorkSheet;
type MetadataField = keyof typeof METADATA_LABEL_ALIASES;

/** Where each part of the allocation sheet was found. Rows and columns are 1-based. */
export interface SourceLayout {
  headerRow: number;
  itemCodeRow: number;
  orderTypeRow: number;
  tranTypeRow: number;
  fallbackItemCodeRow: number | null;
  firstDataRow: number;
  labelColumn: number;
  firstProductColumn: number | null;
  columns: Partial<Record<EmployeeField, number>>;
  /** Every column that holds employee details rather than product quantities. */
  employeeColumns: Set<number>;
}

export type LayoutResult =
  | { ok: true; layout: SourceLayout }
  | { ok: false; missingLabels: MetadataField[] };

/** Compares headings ignoring case, spaces, and punctuation, so "Site Code - Ship" matches "SITE_CODE__SHIP". */
export const compactHeading = (value: unknown) => String(value ?? "").toUpperCase().replace(/[^A-Z0-9]+/g, "");

function matchesAny(value: unknown, aliases: readonly string[]): boolean {
  const compact = compactHeading(value);
  return compact !== "" && aliases.some((alias) => compactHeading(alias) === compact);
}

function valueAt(sheet: Worksheet, row: number, column: number): unknown {
  return (sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })] as XLSX.CellObject | undefined)?.v;
}

function hasValue(sheet: Worksheet, row: number, column: number): boolean {
  const value = valueAt(sheet, row, column);
  return value !== undefined && value !== null && String(value).trim() !== "";
}

/**
 * Finds the allocation layout from its labels instead of fixed positions.
 *
 * The ITEM_CODE, ORDER_TYPE and TRAN_TYPE labels fix the metadata rows and the label column.
 * The row with the most employee headings is the heading row, which also carries product names.
 * Employee rows start below whichever of those rows is lowest.
 */
export function detectLayout(sheet: Worksheet): LayoutResult {
  const range = XLSX.utils.decode_range(sheet["!ref"] ?? "A1:A1");
  const endColumn = range.e.c + 1;
  const scanRows = Math.min(range.e.r + 1, SOURCE.HEADER_SCAN_ROWS);

  // Metadata labels: prefer a column that holds all three, as in every known layout.
  const labelHits: Record<MetadataField, { row: number; column: number }[]> = { itemCode: [], orderType: [], tranType: [] };
  for (let row = 1; row <= scanRows; row += 1) {
    for (let column = 1; column <= endColumn; column += 1) {
      const value = valueAt(sheet, row, column);
      for (const field of Object.keys(METADATA_LABEL_ALIASES) as MetadataField[]) {
        if (matchesAny(value, METADATA_LABEL_ALIASES[field])) labelHits[field].push({ row, column });
      }
    }
  }
  const missingLabels = (Object.keys(labelHits) as MetadataField[]).filter((field) => labelHits[field].length === 0);
  if (missingLabels.length > 0) return { ok: false, missingLabels };

  const sharedColumn = labelHits.itemCode
    .map((hit) => hit.column)
    .find((column) => labelHits.orderType.some((hit) => hit.column === column) && labelHits.tranType.some((hit) => hit.column === column));
  const pick = (field: MetadataField) => (sharedColumn ? labelHits[field].find((hit) => hit.column === sharedColumn)! : labelHits[field][0]);
  const itemCode = pick("itemCode");
  const orderType = pick("orderType");
  const tranType = pick("tranType");
  const labelColumn = itemCode.column;
  const metadataRows = new Set([itemCode.row, orderType.row, tranType.row]);
  const lastMetadataRow = Math.max(...metadataRows);

  // Employee headings may sit above, beside, or between the metadata rows.
  const headingHits = new Map<EmployeeField, { row: number; column: number }[]>();
  const hitsPerRow = new Map<number, number>();
  for (let row = 1; row <= lastMetadataRow; row += 1) {
    for (let column = 1; column <= endColumn; column += 1) {
      if (column === labelColumn && metadataRows.has(row)) continue;
      const value = valueAt(sheet, row, column);
      for (const field of Object.keys(EMPLOYEE_HEADER_ALIASES) as EmployeeField[]) {
        if (!matchesAny(value, EMPLOYEE_HEADER_ALIASES[field])) continue;
        headingHits.set(field, [...(headingHits.get(field) ?? []), { row, column }]);
        hitsPerRow.set(row, (hitsPerRow.get(row) ?? 0) + 1);
      }
    }
  }

  let headerRow = Math.max(1, Math.min(...metadataRows) - 1);
  let bestCount = 0;
  for (const [row, count] of [...hitsPerRow].sort(([a], [b]) => a - b)) {
    if (count > bestCount) [headerRow, bestCount] = [row, count];
  }

  const columns: SourceLayout["columns"] = {};
  const claimed = new Set<number>();
  for (const [field, hits] of headingHits) {
    const hit = hits.find((candidate) => candidate.row === headerRow && !claimed.has(candidate.column)) ??
      hits.find((candidate) => !claimed.has(candidate.column));
    if (!hit) continue;
    columns[field] = hit.column;
    claimed.add(hit.column);
  }
  // Without a name heading, names are kept in the label column under ITEM_CODE / ORDER_TYPE / TRAN_TYPE.
  if (columns.name === undefined && !claimed.has(labelColumn)) columns.name = labelColumn;

  const employeeColumns = new Set<number>([labelColumn, ...Object.values(columns)]);
  const firstDataRow = Math.max(headerRow, lastMetadataRow) + 1;
  const fallbackItemCodeRow = headerRow > 1 && !metadataRows.has(headerRow - 1) ? headerRow - 1 : null;

  // Products start at the first column right of the label column that carries product metadata.
  let firstProductColumn: number | null = null;
  for (let column = labelColumn + 1; column <= endColumn; column += 1) {
    if (employeeColumns.has(column)) continue;
    if ([...metadataRows].some((row) => hasValue(sheet, row, column))) {
      firstProductColumn = column;
      break;
    }
  }

  return {
    ok: true,
    layout: {
      headerRow,
      itemCodeRow: itemCode.row,
      orderTypeRow: orderType.row,
      tranTypeRow: tranType.row,
      fallbackItemCodeRow,
      firstDataRow,
      labelColumn,
      firstProductColumn,
      columns,
      employeeColumns,
    },
  };
}
