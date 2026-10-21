import { ERP_CONSTANTS } from "@/lib/excel/templateConfig";
import type { CalendarDate } from "@/lib/date";
import { toExcelSerial } from "@/lib/date";
import type { EmployeeAllocation, HeaderRow, OrderKind, OrderWorkbookData, ParsedSource } from "@/lib/types";

function orderId(index: number): string {
  return `DIST${String(index + 1).padStart(3, "0")}`;
}

function makeHeader(employee: EmployeeAllocation, id: string, kind: OrderKind, dateSerial: number): HeaderRow {
  return {
    DIST_ORDER: id,
    ORDER_DATE: dateSerial,
    SITE_CODE__SHIP: employee.shipSite,
    SITE_CODE__DLV: employee.deliverySite,
    SHIP_DATE: dateSerial,
    DUE_DATE: dateSerial,
    REMARKS: employee.remarks,
    DIST_ROUTE: "",
    PRICE_LIST: ERP_CONSTANTS.PRICE_LIST,
    LOC_CODE__GIT: ERP_CONSTANTS.LOC_CODE__GIT,
    SITE_CODE: ERP_CONSTANTS.SITE_CODE,
    CURR_CODE: ERP_CONSTANTS.CURR_CODE,
    EXCH_RATE: ERP_CONSTANTS.EXCH_RATE,
    STATUS: ERP_CONSTANTS.STATUS,
    MRP_RUN_ID: "",
    REMARKS1: "",
    ORDER_TYPE: kind,
    LOC_CODE__CONS: ERP_CONSTANTS.LOC_CODE__CONS,
    SUNDRY_TYPE: ERP_CONSTANTS.SUNDRY_TYPE,
    AUTO_RECEIPT: ERP_CONSTANTS.AUTO_RECEIPT,
    TRAN_TYPE: kind,
    LOC_CODE__GITBF: ERP_CONSTANTS.LOC_CODE__GITBF,
    AVALIABLE_YN: ERP_CONSTANTS.AVALIABLE_YN,
    TRAN_SER: ERP_CONSTANTS.TRAN_SER,
    PRICE_LIST__CLG: ERP_CONSTANTS.PRICE_LIST__CLG,
    TRANS_MODE: ERP_CONSTANTS.TRANS_MODE,
    SALES_PERS: employee.salesPerson,
    PROJ_CODE: employee.projectCode,
  };
}

export function generateOrders(source: ParsedSource, date: CalendarDate): { ps: OrderWorkbookData; li: OrderWorkbookData } {
  const serial = toExcelSerial(date);
  const ps: OrderWorkbookData = { kind: "PS", headers: [], details: [] };
  const li: OrderWorkbookData = { kind: "LI", headers: [], details: [] };
  let employeeOrderIndex = 0;

  for (const employee of source.employees) {
    const psAllocations = employee.allocations.filter(({ product }) => product.orderType === "PS" && product.tranType === "PS");
    const liAllocations = employee.allocations.filter(({ product }) => product.orderType === "LI" && product.tranType === "LI");
    if (psAllocations.length === 0 && liAllocations.length === 0) continue;

    const id = orderId(employeeOrderIndex);
    employeeOrderIndex += 1;

    for (const [target, kind, allocations] of [[ps, "PS", psAllocations], [li, "LI", liAllocations]] as const) {
      if (allocations.length === 0) continue;
      target.headers.push(makeHeader(employee, id, kind, serial));
      allocations.forEach(({ product, quantity }, index) => {
        target.details.push({
          DIST_ORDER: id,
          LINE_NO: index + 1,
          ITEM_CODE: product.itemCode,
          QTY_ORDER: quantity,
          QTY_CONFIRM: quantity,
          QTY_RECEIVED: 0,
          QTY_SHIPPED: 0,
          DUE_DATE: serial,
          SHIP_DATE: serial,
          PACK_INSTR: "",
          REAS_CODE: ERP_CONSTANTS.REAS_CODE,
        });
      });
    }
  }

  return { ps, li };
}
