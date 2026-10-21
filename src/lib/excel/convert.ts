import { getIndiaCalendarDate } from "@/lib/date";
import { generateOrders } from "@/lib/excel/generateOrders";
import { readSourceWorkbook } from "@/lib/excel/readSource";
import { writeAllOutputs } from "@/lib/excel/writeOutputs";
import type { ConversionResult, ValidationIssue } from "@/lib/types";

export async function convertWorkbook(data: ArrayBuffer, now = new Date()): Promise<ConversionResult> {
  const source = readSourceWorkbook(data);
  if (source.issues.some((issue) => issue.severity === "error")) return { issues: source.issues, excludedProducts: source.excludedProducts };

  const { ps, li } = generateOrders(source, getIndiaCalendarDate(now));
  const activeEmployees = new Set([...ps.headers, ...li.headers].map((row) => row.DIST_ORDER)).size;

  if (ps.details.length > 65_535 || li.details.length > 65_535) {
    const limitIssue: ValidationIssue = {
      id: "biff8-limit",
      severity: "error",
      en: "The output has more than 65,535 detail rows in one file, which the legacy .xls format cannot store. Split the source workbook into smaller files.",
      hi: "एक आउटपुट फ़ाइल में 65,535 से अधिक detail पंक्तियाँ हैं, जिन्हें पुराना .xls प्रारूप सेव नहीं कर सकता। source workbook को छोटी फ़ाइलों में बाँटें।",
    };
    return { issues: [...source.issues, limitIssue], excludedProducts: source.excludedProducts };
  }

  try {
    const generated = await writeAllOutputs(ps, li);
    return {
      issues: source.issues,
      excludedProducts: source.excludedProducts,
      ps,
      li,
      generated,
      summary: {
        employees: activeEmployees,
        psOrders: ps.headers.length,
        liOrders: li.headers.length,
        totalLines: ps.details.length + li.details.length,
      },
    };
  } catch {
    const issue: ValidationIssue = {
      id: "output-verification",
      severity: "error",
      en: "The generated files did not pass the safety check, so no download was created. Please check the source workbook structure and try again.",
      hi: "बनाई गई फ़ाइलें सुरक्षा जाँच में पास नहीं हुईं, इसलिए डाउनलोड तैयार नहीं किया गया। source workbook का प्रारूप जाँचें और दोबारा प्रयास करें।",
    };
    return { issues: [...source.issues, issue], excludedProducts: source.excludedProducts };
  }
}
