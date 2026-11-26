export type NumberFormat = "General" | "Number" | "Currency ₹" | "Currency $" | "Currency €" | "Currency £" | "Currency ¥" | "Currency AED" | "Accounting" | "Short Date" | "Long Date" | "Time" | "Percentage" | "Percentage %" | "Fraction" | "Scientific" | "Text" | "Date";
export type Align = "left" | "center" | "right";
export interface CellStyle {
  fontFamily?: string; fontSize?: number; bold?: boolean; italic?: boolean;
  underline?: boolean; doubleUnderline?: boolean; strike?: boolean; color?: string; backgroundColor?: string;
  align?: Align; verticalAlign?: "top" | "middle" | "bottom"; indent?: number; textOrientation?: "horizontal" | "vertical";
  wrap?: boolean; numberFormat?: NumberFormat; decimalPlaces?: number;
  border?: "all" | "outer" | "none"; borderColor?: string;
  borderTop?: string; borderRight?: string; borderBottom?: string; borderLeft?: string;
}
export interface CellData { raw: string; cachedValue?: string; style?: CellStyle; comment?: string; hyperlink?: string; image?: string; checkbox?: boolean; datePicker?: boolean; mergeRows?: number; mergeCols?: number; mergedInto?: string; mergeBackup?: Record<string, CellData>; }
export type CellMap = Record<string, CellData>;
export interface TableFilterState {
  selectedValues?: string[];
  condition?: { operator: "contains" | "notContains" | "equals" | "notEquals" | "beginsWith" | "endsWith" | "greaterThan" | "lessThan" | "between"; value: string; secondValue?: string };
}
export interface ExcelTableData {
  id: string;
  name: string;
  range: CellRange;
  hasHeaders: boolean;
  showTotalRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  bandedRows: boolean;
  bandedColumns: boolean;
  showFilterButtons: boolean;
  styleId: string;
  filters: Record<number, TableFilterState>;
}
export interface SheetData {
  id: string; name: string; cells: CellMap; rowCount: number; columnCount: number;
  columnWidths: Record<number, number>; rowHeights?: Record<number, number>; showGridlines: boolean; showHeadings: boolean;
  freezeTopRow: boolean; filters: boolean; hiddenRows?: number[]; hiddenColumns?: number[];
  tables?: ExcelTableData[];
}
export type ChartType = "column" | "line" | "scatter" | "pie" | "bar" | "statistical" | "combo" | "treemap";
export interface FloatingChartData {
  id: string; type: ChartType; title: string; range: string; variant?: string;
  x: number; y: number; width: number; height: number;
}
export interface FloatingElementData {
  id: string;
  type: "shape" | "image" | "text";
  shape?: "rectangle" | "ellipse" | "line" | "arrow" | "line-arrow" | "double-arrow" | "elbow-connector" | "curved-line" | "rounded-rectangle" | "snip-single-corner" | "triangle" | "right-triangle" | "diamond" | "pentagon" | "hexagon" | "star" | "heart" | "plus-cross" | "right-arrow" | "left-arrow" | "up-arrow" | "down-arrow" | "left-right-arrow" | "rect-callout" | "rounded-rect-callout" | "oval-callout" | "cloud-callout";
  x: number; y: number; width: number; height: number;
  text?: string; fill?: string; stroke?: string; src?: string;
  fontSize?: number; color?: string;
}
export type CalculationMode = "Automatic" | "Manual";
export type NamedRanges = Record<string,string>;
export interface WorkbookData {
  id: string; title: string; sheets: SheetData[]; charts: FloatingChartData[];
  elements?: FloatingElementData[];
  namedRanges?: NamedRanges;
  calculationMode?: CalculationMode;
  createdAt: number; modifiedAt: number; orientation: "portrait" | "landscape";
}
export type TemplateId = "blank" | "budget" | "sales" | "project";
export interface CellRange { startRow: number; startCol: number; endRow: number; endCol: number; }
export interface ClipboardCell { raw: string; style?: CellStyle; }
export interface ClipboardData { cells: ClipboardCell[][]; source?: string; }
export const DEFAULT_ROWS = 100;
export const DEFAULT_COLUMNS = 26;
export const DEFAULT_STYLE: CellStyle = {
  fontFamily: "Segoe UI", fontSize: 11, color: "#202124",
  backgroundColor: "#ffffff", align: "left", numberFormat: "General",
};
export const columnLabel = (n: number): string => {
  let value = n + 1, label = "";
  while (value > 0) {
    const r = (value - 1) % 26;
    label = String.fromCharCode(65 + r) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
};
export const addressOf = (row: number, col: number): string => `${columnLabel(col)}${row + 1}`;
export const parseAddress = (address: string): { row: number; col: number } | null => {
  const m = address.toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!m) return null;
  let col = 0;
  for (const c of m[1]) col = col * 26 + c.charCodeAt(0) - 64;
  return { row: Number(m[2]) - 1, col: col - 1 };
};
export const emptySheet = (name = "Sheet 1"): SheetData => ({
  id: crypto.randomUUID(), name, cells: {}, rowCount: DEFAULT_ROWS,
  columnCount: DEFAULT_COLUMNS, columnWidths: {}, showGridlines: true,
  showHeadings: true, freezeTopRow: false, filters: false, tables: [],
});
export const emptyWorkbook = (title = "Untitled Workbook"): WorkbookData => ({
  id: crypto.randomUUID(), title, sheets: [emptySheet()], charts: [], elements: [],
  createdAt: Date.now(), modifiedAt: Date.now(), orientation: "portrait",
});
