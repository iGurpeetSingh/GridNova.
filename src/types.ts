export type NumberFormat = "General" | "Number" | "Currency ₹" | "Currency $" | "Percentage %" | "Date";
export type Align = "left" | "center" | "right";
export interface CellStyle {
  fontFamily?: string; fontSize?: number; bold?: boolean; italic?: boolean;
  underline?: boolean; strike?: boolean; color?: string; backgroundColor?: string;
  align?: Align; wrap?: boolean; numberFormat?: NumberFormat;
  border?: "all" | "outer" | "none";
}
export interface CellData { raw: string; style?: CellStyle; comment?: string; hyperlink?: string; checkbox?: boolean; }
export type CellMap = Record<string, CellData>;
export interface SheetData {
  id: string; name: string; cells: CellMap; rowCount: number; columnCount: number;
  columnWidths: Record<number, number>; showGridlines: boolean; showHeadings: boolean;
  freezeTopRow: boolean; filters: boolean;
}
export interface FloatingChartData {
  id: string; type: "bar" | "line" | "pie"; title: string; range: string;
  x: number; y: number; width: number; height: number;
}
export interface WorkbookData {
  id: string; title: string; sheets: SheetData[]; charts: FloatingChartData[];
  createdAt: number; modifiedAt: number; orientation: "portrait" | "landscape";
}
export type TemplateId = "blank" | "budget" | "sales" | "project";
export interface CellRange { startRow: number; startCol: number; endRow: number; endCol: number; }
export interface ClipboardCell { raw: string; style?: CellStyle; }
export interface ClipboardData { cells: ClipboardCell[][]; }
export const DEFAULT_ROWS = 100;
export const DEFAULT_COLUMNS = 26;
export const DEFAULT_STYLE: CellStyle = {
  fontFamily: "Inter", fontSize: 11, color: "#202124",
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
  showHeadings: true, freezeTopRow: false, filters: false,
});
export const emptyWorkbook = (title = "Untitled Workbook"): WorkbookData => ({
  id: crypto.randomUUID(), title, sheets: [emptySheet()], charts: [],
  createdAt: Date.now(), modifiedAt: Date.now(), orientation: "portrait",
});