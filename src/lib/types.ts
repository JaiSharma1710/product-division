export type OrderKind = "PS" | "LI";

export interface BilingualText {
  en: string;
  hi: string;
}

export type ValidationSeverity = "error" | "warning";

export interface ValidationIssue extends BilingualText {
  id: string;
  severity: ValidationSeverity;
  sheet?: string;
  cell?: string;
  sourceRow?: number;
  employeeCode?: string;
  employeeName?: string;
  productName?: string;
  quantity?: unknown;
}

export interface ProductColumn {
  columnIndex: number;
  columnLetter: string;
  name: string;
  itemCode: string;
  orderType: string;
  tranType: string;
  itemCodeSourceRow: 2 | 4 | null;
}

export interface ExcludedProduct {
  columnLetter: string;
  name: string;
  hasItemCode: boolean;
  hasOrderType: boolean;
  hasTranType: boolean;
}

export interface Allocation {
  product: ProductColumn;
  quantity: number;
}

export interface EmployeeAllocation {
  sourceRow: number;
  employeeCode: string;
  employeeName: string;
  shipSite: string;
  deliverySite: string;
  remarks: string;
  salesPerson: string;
  projectCode: string;
  allocations: Allocation[];
}

export interface ParsedSource {
  employees: EmployeeAllocation[];
  products: ProductColumn[];
  excludedProducts: ExcludedProduct[];
  issues: ValidationIssue[];
}

export interface HeaderRow {
  DIST_ORDER: string;
  ORDER_DATE: number;
  SITE_CODE__SHIP: string;
  SITE_CODE__DLV: string;
  SHIP_DATE: number;
  DUE_DATE: number;
  REMARKS: string;
  DIST_ROUTE: string;
  PRICE_LIST: string;
  LOC_CODE__GIT: string;
  SITE_CODE: string;
  CURR_CODE: string;
  EXCH_RATE: number;
  STATUS: string;
  MRP_RUN_ID: string;
  REMARKS1: string;
  ORDER_TYPE: OrderKind;
  LOC_CODE__CONS: string;
  SUNDRY_TYPE: string;
  AUTO_RECEIPT: string;
  TRAN_TYPE: OrderKind;
  LOC_CODE__GITBF: string;
  AVALIABLE_YN: string;
  TRAN_SER: string;
  PRICE_LIST__CLG: string;
  TRANS_MODE: string;
  SALES_PERS: string;
  PROJ_CODE: string;
}

export interface DetailRow {
  DIST_ORDER: string;
  LINE_NO: number;
  ITEM_CODE: string;
  QTY_ORDER: number;
  QTY_CONFIRM: number;
  QTY_RECEIVED: number;
  QTY_SHIPPED: number;
  DUE_DATE: number;
  SHIP_DATE: number;
  PACK_INSTR: string;
  REAS_CODE: string;
}

export interface OrderWorkbookData {
  kind: OrderKind;
  headers: HeaderRow[];
  details: DetailRow[];
}

export interface GeneratedFiles {
  ps: Uint8Array;
  li: Uint8Array;
  zip: Blob;
}

export interface ConversionSummary {
  employees: number;
  psOrders: number;
  liOrders: number;
  totalLines: number;
}

export interface ConversionResult {
  issues: ValidationIssue[];
  excludedProducts: ExcludedProduct[];
  ps?: OrderWorkbookData;
  li?: OrderWorkbookData;
  summary?: ConversionSummary;
  generated?: GeneratedFiles;
}
