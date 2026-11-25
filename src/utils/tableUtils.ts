import type { CellMap, CellRange, ExcelTableData, TableFilterState } from "../types";
import { addressOf } from "../types";

export interface TablePalette { header: string; band: string; accent: string; text: string; }
const palettes: Record<string, TablePalette> = {
  TableStyleLight1: { header: "#F2F2F2", band: "#F7F7F7", accent: "#A6A6A6", text: "#202124" },
  TableStyleLight2: { header: "#D9EAD3", band: "#F3F8F1", accent: "#70AD47", text: "#24372A" },
  TableStyleLight9: { header: "#DDEBF7", band: "#F3F8FC", accent: "#5B9BD5", text: "#203040" },
  TableStyleMedium2: { header: "#4472C4", band: "#D9E2F3", accent: "#4472C4", text: "#FFFFFF" },
  TableStyleMedium4: { header: "#70AD47", band: "#E2F0D9", accent: "#70AD47", text: "#FFFFFF" },
  TableStyleMedium9: { header: "#5B9BD5", band: "#DDEBF7", accent: "#5B9BD5", text: "#FFFFFF" },
  TableStyleMedium16: { header: "#107C41", band: "#E6F2EC", accent: "#107C41", text: "#FFFFFF" },
  TableStyleMedium21: { header: "#8064A2", band: "#E4DFEC", accent: "#8064A2", text: "#FFFFFF" },
  TableStyleDark1: { header: "#404040", band: "#D9D9D9", accent: "#404040", text: "#FFFFFF" },
  TableStyleDark2: { header: "#1F4E78", band: "#D6E4F0", accent: "#1F4E78", text: "#FFFFFF" },
  TableStyleDark4: { header: "#375623", band: "#E2F0D9", accent: "#375623", text: "#FFFFFF" },
};
export const TABLE_STYLES = [
  "TableStyleLight1", "TableStyleLight2", "TableStyleLight9",
  "TableStyleMedium2", "TableStyleMedium4", "TableStyleMedium9", "TableStyleMedium16", "TableStyleMedium21",
  "TableStyleDark1", "TableStyleDark2", "TableStyleDark4",
] as const;
export const getTablePalette = (styleId: string): TablePalette => palettes[styleId] ?? palettes.TableStyleMedium16;

export function getTableAt(tables: ExcelTableData[] | undefined, row: number, col: number): ExcelTableData | undefined {
  return tables?.find(table => {
    const endRow = table.range.endRow + (table.showTotalRow ? 1 : 0);
    return row >= table.range.startRow && row <= endRow && col >= table.range.startCol && col <= table.range.endCol;
  });
}
export function getTableHeaderLabel(table: ExcelTableData, column: number, cells: CellMap): string {
  if (!table.hasHeaders) return `Column${column - table.range.startCol + 1}`;
  const value = cells[addressOf(table.range.startRow, column)]?.raw?.trim();
  return value || `Column${column - table.range.startCol + 1}`;
}
export function getTableColumnIndex(table: ExcelTableData, label: string, cells: CellMap): number {
  for (let col = table.range.startCol; col <= table.range.endCol; col++) {
    if (getTableHeaderLabel(table, col, cells).toLowerCase() === label.trim().toLowerCase()) return col;
  }
  return -1;
}
export function tableFilterPasses(table: ExcelTableData, row: number, cells: CellMap): boolean {
  if (row <= table.range.startRow || row > table.range.endRow) return true;
  for (const [columnKey, filter] of Object.entries(table.filters ?? {})) {
    const col = Number(columnKey), raw = cells[addressOf(row, col)]?.raw ?? "";
    if (filter.selectedValues && !filter.selectedValues.includes(raw)) return false;
    if (filter.condition && !matchesCondition(raw, filter.condition)) return false;
  }
  return true;
}
function matchesCondition(raw: string, condition: NonNullable<TableFilterState["condition"]>): boolean {
  const value = condition.value, source = raw.toLowerCase(), expected = value.toLowerCase();
  const left = Number(raw), right = Number(value), numeric = raw.trim() !== "" && value.trim() !== "" && Number.isFinite(left) && Number.isFinite(right);
  switch (condition.operator) {
    case "contains": return source.includes(expected);
    case "notContains": return !source.includes(expected);
    case "equals": return numeric ? left === right : source === expected;
    case "notEquals": return numeric ? left !== right : source !== expected;
    case "beginsWith": return source.startsWith(expected);
    case "endsWith": return source.endsWith(expected);
    case "greaterThan": return numeric ? left > right : raw.localeCompare(value, undefined, { numeric: true, sensitivity: "base" }) > 0;
    case "lessThan": return numeric ? left < right : raw.localeCompare(value, undefined, { numeric: true, sensitivity: "base" }) < 0;
    case "between": {
      const second = Number(condition.secondValue ?? "");
      return numeric && Number.isFinite(second) && left >= Math.min(right, second) && left <= Math.max(right, second);
    }
    default: return true;
  }
}
export function tableColumnValues(table: ExcelTableData, column: number, cells: CellMap): string[] {
  const values = new Set<string>();
  for (let row = table.range.startRow + 1; row <= table.range.endRow; row++) {
    values.add(cells[addressOf(row, column)]?.raw ?? "");
  }
  return [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
}
export function tableReferenceRange(table: ExcelTableData, selector: string | undefined): CellRange {
  const startRow = table.range.startRow, endRow = table.range.endRow;
  if (selector === "#Headers") return { startRow, endRow: startRow, startCol: table.range.startCol, endCol: table.range.endCol };
  if (selector === "#Totals") {
    const total = endRow + (table.showTotalRow ? 1 : 0);
    return { startRow: total, endRow: total, startCol: table.range.startCol, endCol: table.range.endCol };
  }
  if (selector === "#Data") return { startRow: startRow + 1, endRow, startCol: table.range.startCol, endCol: table.range.endCol };
  return { startRow: startRow, endRow: endRow + (table.showTotalRow ? 1 : 0), startCol: table.range.startCol, endCol: table.range.endCol };
}
export function normalizeTableName(value: string): string {
  const sanitized = value.trim().replace(/[^A-Za-z0-9_.]/g, "_").replace(/^[^A-Za-z_]+/, "");
  return sanitized || "Table1";
}
