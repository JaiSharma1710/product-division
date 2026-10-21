"use client";

import { useState } from "react";
import { DETAIL_COLUMNS, HEADER_COLUMNS } from "@/lib/excel/templateConfig";
import type { DetailRow, HeaderRow, OrderWorkbookData } from "@/lib/types";

interface WorkbookPreviewProps {
  workbook: OrderWorkbookData;
  filename: string;
  onBack: () => void;
  onDownload: () => void;
}

function displayValue(column: string, value: string | number): string {
  if (column.includes("DATE") && typeof value === "number") {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000);
    return `${String(date.getUTCDate()).padStart(2, "0")}/${String(date.getUTCMonth() + 1).padStart(2, "0")}/${date.getUTCFullYear()}`;
  }
  return String(value);
}

function SheetTable<T extends HeaderRow | DetailRow>({ columns, rows }: { columns: readonly (keyof T)[]; rows: T[] }) {
  return (
    <div className="preview-table-scroll" tabIndex={0} aria-label="Excel sheet preview / Excel शीट पूर्वावलोकन">
      <table className="excel-table">
        <thead><tr><th className="row-number-cell" aria-label="Row number / पंक्ति संख्या">#</th>{columns.map((column) => <th scope="col" key={String(column)}>{String(column)}</th>)}</tr></thead>
        <tbody>
          {rows.length === 0 ? (
            <tr><th className="row-number-cell" scope="row">1</th><td colSpan={columns.length}>No rows / कोई पंक्ति नहीं</td></tr>
          ) : rows.map((row, rowIndex) => (
            <tr key={`${String(row.DIST_ORDER)}-${rowIndex}`}>
              <th className="row-number-cell" scope="row">{rowIndex + 2}</th>
              {columns.map((column) => <td key={String(column)}>{displayValue(String(column), row[column] as string | number)}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function WorkbookPreview({ workbook, filename, onBack, onDownload }: WorkbookPreviewProps) {
  const [sheet, setSheet] = useState<"Header" | "Detail">("Header");
  return (
    <main className="preview-page">
      <header className="preview-toolbar">
        <button className="preview-back" type="button" onClick={onBack}><span aria-hidden="true">←</span> Back to downloads / डाउनलोड पर वापस जाएँ</button>
        <div className="preview-title">
          <span className="preview-file-badge">XLS</span>
          <div><h1>{filename}</h1><p>Workbook preview / Workbook पूर्वावलोकन</p></div>
        </div>
        <button className="button button-primary preview-download" type="button" onClick={onDownload}>Download / डाउनलोड</button>
      </header>

      <section className="preview-workspace" aria-label={`${filename} preview`}>
        <div className="preview-sheet-bar">
          <div className="sheet-tabs" role="tablist" aria-label="Workbook sheets / Workbook शीट">
            {(["Header", "Detail"] as const).map((name) => (
              <button key={name} type="button" role="tab" aria-selected={sheet === name} className={sheet === name ? "is-selected" : ""} onClick={() => setSheet(name)}>
                {name} <span>{name === "Header" ? workbook.headers.length : workbook.details.length}</span>
              </button>
            ))}
          </div>
          <p>Use the scrollbars to see all columns and rows. / सभी कॉलम और पंक्तियाँ देखने के लिए स्क्रॉल करें।</p>
        </div>
        {sheet === "Header" ? <SheetTable<HeaderRow> columns={HEADER_COLUMNS} rows={workbook.headers} /> : <SheetTable<DetailRow> columns={DETAIL_COLUMNS} rows={workbook.details} />}
      </section>
    </main>
  );
}
