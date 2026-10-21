import type { DetailRow, HeaderRow } from "@/lib/types";

export const HEADER_COLUMNS: (keyof HeaderRow)[] = [
  "DIST_ORDER", "ORDER_DATE", "SITE_CODE__SHIP", "SITE_CODE__DLV", "SHIP_DATE",
  "DUE_DATE", "REMARKS", "DIST_ROUTE", "PRICE_LIST", "LOC_CODE__GIT", "SITE_CODE",
  "CURR_CODE", "EXCH_RATE", "STATUS", "MRP_RUN_ID", "REMARKS1", "ORDER_TYPE",
  "LOC_CODE__CONS", "SUNDRY_TYPE", "AUTO_RECEIPT", "TRAN_TYPE", "LOC_CODE__GITBF",
  "AVALIABLE_YN", "TRAN_SER", "PRICE_LIST__CLG", "TRANS_MODE", "SALES_PERS", "PROJ_CODE",
];

export const DETAIL_COLUMNS: (keyof DetailRow)[] = [
  "DIST_ORDER", "LINE_NO", "ITEM_CODE", "QTY_ORDER", "QTY_CONFIRM", "QTY_RECEIVED",
  "QTY_SHIPPED", "DUE_DATE", "SHIP_DATE", "PACK_INSTR", "REAS_CODE",
];

export const DATE_FORMAT = "dd/mm/yyyy";

export const ERP_CONSTANTS = {
  PRICE_LIST: "GTPSA",
  LOC_CODE__GIT: "GIT",
  SITE_CODE: "MM003",
  CURR_CODE: "INR",
  EXCH_RATE: 1,
  STATUS: "P",
  LOC_CODE__CONS: "QUAR",
  SUNDRY_TYPE: "C",
  AUTO_RECEIPT: "Y",
  LOC_CODE__GITBF: "GIT",
  AVALIABLE_YN: "Y",
  TRAN_SER: "D-ORD",
  PRICE_LIST__CLG: "GTPSA",
  TRANS_MODE: "D",
  REAS_CODE: "TEST",
} as const;

export const SOURCE = {
  SHEET_NAME: "allocation",
  HEADER_ROW: 3,
  FALLBACK_ITEM_CODE_ROW: 2,
  ITEM_CODE_ROW: 4,
  ORDER_TYPE_ROW: 5,
  TRAN_TYPE_ROW: 6,
  FIRST_EMPLOYEE_ROW: 7,
  FIRST_PRODUCT_COLUMN: 18,
  EMPLOYEE: {
    code: 2,
    name: 4,
    shipSite: 5,
    deliverySite: 6,
    remarks: 7,
    salesPerson: 8,
    projectCode: 9,
  },
} as const;

export const EMPLOYEE_HEADER_ALIASES = {
  code: ["EMPLOYEE CODE", "EMP CODE", "EMP_CODE", "EMPLOYEE_CODE", "CODE"],
  name: ["EMPLOYEE NAME", "EMP NAME", "EMP_NAME", "EMPLOYEE_NAME", "NAME"],
  shipSite: ["SITE_CODE__SHIP", "SITE CODE SHIP", "SHIP SITE"],
  deliverySite: ["SITE_CODE__DLV", "SITE CODE DLV", "DELIVERY SITE"],
  remarks: ["REMARKS", "REMARK"],
  salesPerson: ["SALES_PERS", "SALES PERS", "SALESPERSON", "SALES PERSON"],
  projectCode: ["PROJ_CODE", "PROJECT CODE", "PROJ CODE"],
} as const;
