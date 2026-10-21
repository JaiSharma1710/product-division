const IST_TIME_ZONE = "Asia/Kolkata";

export interface CalendarDate {
  year: number;
  month: number;
  day: number;
}

export function getIndiaCalendarDate(now = new Date()): CalendarDate {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: IST_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);

  const value = (type: Intl.DateTimeFormatPartTypes) =>
    Number(parts.find((part) => part.type === type)?.value);

  return { year: value("year"), month: value("month"), day: value("day") };
}

/** Excel's 1900 date system includes the historic fictitious 29-Feb-1900 day. */
export function toExcelSerial(date: CalendarDate): number {
  const excelEpoch = Date.UTC(1899, 11, 30);
  const current = Date.UTC(date.year, date.month - 1, date.day);
  return (current - excelEpoch) / 86_400_000;
}

export function formatIndiaDate(date: CalendarDate): string {
  return `${String(date.day).padStart(2, "0")}/${String(date.month).padStart(2, "0")}/${date.year}`;
}
