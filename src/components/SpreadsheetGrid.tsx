import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { KeyboardEvent, MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent } from "react";
import { ExternalLink, Filter, Plus, Search } from "lucide-react";
import type { CellRange, CellStyle, SheetData, TableFilterState } from "../types";
import { addressOf, columnLabel, parseAddress } from "../types";
import { formatCellValue } from "../utils/formulaEngine";
import { getTableAt, getTablePalette, tableColumnValues, tableFilterPasses } from "../utils/tableUtils";
import { useWorkbookStore } from "../store/useWorkbookStore";

interface Props {
  sheet: SheetData; activeCell: string; selection: CellRange | null; showFormulas: boolean; zoom: number;
  namedRanges?:Record<string,string>; formulaBarFocused?: boolean;
  onSelect: (address: string, range?: CellRange) => void;
}
type ContextMenuState = { x: number; y: number; kind: "cell" | "row" | "column"; index: number; col: number } | null;
type InteractionState = "READY" | "SELECTING" | "FILLING" | "EDITING";
type FillAxis = "vertical" | "horizontal";
type OverlayRect = { left: number; top: number; width: number; height: number };
type FillDragState = { source: CellRange; target: CellRange; axis: FillAxis | null; startX: number; startY: number; clientX: number; clientY: number; previewValue: string };
type TableFilterMenuState = { tableId: string; col: number; x: number; y: number } | null;
type TableFilterOperator = NonNullable<TableFilterState["condition"]>["operator"];
const clampRange = (range: CellRange) => ({ startRow: Math.min(range.startRow, range.endRow), endRow: Math.max(range.startRow, range.endRow), startCol: Math.min(range.startCol, range.endCol), endCol: Math.max(range.startCol, range.endCol) });
const shiftFormulaReferences = (raw: string, rowDelta: number, colDelta: number) => raw.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (match, colLock: string, letters: string, rowLock: string, digits: string) => {
  const position = parseAddress(`${letters}${digits}`); if (!position) return match;
  return `${colLock ? "$" + letters : columnLabel(position.col + colDelta)}${rowLock ? "$" + digits : position.row + rowDelta + 1}`;
});
const adjustReferences = (raw: string, axis: "row" | "column", index: number, count: number, insert: boolean) => raw.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (match, colLock: string, letters: string, rowLock: string, digits: string, offset: number) => {
  const position = parseAddress(`${letters}${digits}`); if (!position) return match;
  const coordinate = axis === "row" ? position.row : position.col;
  const inRange = raw.slice(0, offset).trimEnd().endsWith(":") || raw.slice(offset + match.length).trimStart().startsWith(":");
  if (!insert && coordinate >= index && coordinate < index + count && !inRange) return "#REF!";
  const updated = insert && coordinate >= index ? coordinate + count : !insert && coordinate >= index + count ? coordinate - count : !insert && coordinate >= index ? index : coordinate;
  return axis === "row" ? `${colLock}${letters}${rowLock}${updated + 1}` : `${colLock}${columnLabel(updated)}${rowLock}${digits}`;
});

export default function SpreadsheetGrid({ sheet, activeCell, selection, showFormulas, zoom, namedRanges={}, formulaBarFocused = false, onSelect }: Props) {
  const setCell = useWorkbookStore(s => s.setCell);
  const updateSheet = useWorkbookStore(s => s.updateSheet);
  const setClipboard = useWorkbookStore(s => s.setClipboard);
  const clipboard = useWorkbookStore(s => s.clipboard);
  const undo = useWorkbookStore(s => s.undo);
  const redo = useWorkbookStore(s => s.redo);
  const [dragStart, setDragStart] = useState<{ row: number; col: number } | null>(null);
  const [columnFilters, setColumnFilters] = useState<Record<number, string>>({});
  const [editingAddress, setEditingAddress] = useState<string | null>(null);
  const [interactionState, setInteractionState] = useState<InteractionState>("READY");
  const [menu, setMenu] = useState<ContextMenuState>(null);
  const rowHeights = sheet.rowHeights ?? {};
  const [fillPreview, setFillPreview] = useState<CellRange | null>(null);
  const [fillDrag, setFillDrag] = useState<FillDragState | null>(null);
  const [copyRange, setCopyRange] = useState<CellRange | null>(null);
  const [overlayRects, setOverlayRects] = useState<{ selection: OverlayRect | null; fill: OverlayRect | null; copy: OverlayRect | null }>({ selection: null, fill: null, copy: null });
  const [resizePreview, setResizePreview] = useState<{ axis: "row" | "column"; index: number; value: number } | null>(null);
  const [formatDialogOpen, setFormatDialogOpen] = useState(false);
  const [formatDraft, setFormatDraft] = useState<CellStyle>({});
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [tableFilterMenu, setTableFilterMenu] = useState<TableFilterMenuState>(null);
  const [tableFilterSearch, setTableFilterSearch] = useState("");
  const [tableFilterValues, setTableFilterValues] = useState<string[]>([]);
  const [tableFilterOperator, setTableFilterOperator] = useState<TableFilterOperator>("contains");
  const [tableFilterConditionValue, setTableFilterConditionValue] = useState("");
  const [tableFilterConditionSecond, setTableFilterConditionSecond] = useState("");
  const resizeValue = useRef<{ axis: "row" | "column"; index: number; value: number } | null>(null);
  const fillSource = useRef<CellRange | null>(null);
  const fillModifier = useRef(false);
  const fillAxis = useRef<FillAxis | null>(null);
  const fillDragRef = useRef<FillDragState | null>(null);
  const pointerPosition = useRef({ x: 0, y: 0 });
  const lastDragAddress = useRef<string | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const editOriginal = useRef("");
  const interactionHandlersRef = useRef<{ update: (x: number, y: number) => void; finish: () => void }>({ update: () => undefined, finish: () => undefined });
  useEffect(() => {
    if (!menu) return;
    requestAnimationFrame(() => menuRef.current?.querySelector<HTMLButtonElement>("[role=menuitem]")?.focus());
    const dismissOutside = (event: PointerEvent) => { if (!menuRef.current?.contains(event.target as Node)) setMenu(null); };
    const dismissEscape = (event: globalThis.KeyboardEvent) => { if (event.key === "Escape") setMenu(null); };
    document.addEventListener("pointerdown", dismissOutside);
    document.addEventListener("keydown", dismissEscape);
    return () => { document.removeEventListener("pointerdown", dismissOutside); document.removeEventListener("keydown", dismissEscape); };
  }, [menu]);
  useLayoutEffect(() => {
    const panel = menuRef.current; if (!menu || !panel) return;
    const rect = panel.getBoundingClientRect(), margin = 8;
    const x = Math.max(margin, Math.min(menu.x, window.innerWidth - rect.width - margin));
    const y = Math.max(margin, Math.min(menu.y, window.innerHeight - rect.height - margin));
    if (x !== menu.x || y !== menu.y) setMenu(current => current ? { ...current, x, y } : null);
  }, [menu]);
  const active = parseAddress(activeCell) ?? { row: 0, col: 0 };
  const hiddenRows = sheet.hiddenRows ?? [];
  const hiddenColumns = sheet.hiddenColumns ?? [];
  const columns = useMemo(() => Array.from({ length: Math.max(sheet.columnCount, 26) }, (_, i) => i).filter(col => !hiddenColumns.includes(col)), [sheet.columnCount, hiddenColumns]);
  const rows = useMemo(() => Array.from({ length: Math.max(sheet.rowCount + 1, 100) }, (_, i) => i).filter(row =>
    !hiddenRows.includes(row) &&
    (row === 0 || Object.entries(columnFilters).every(([col, value]) => !value || (sheet.cells[addressOf(row, Number(col))]?.raw ?? "") === value)) &&
    (sheet.tables ?? []).every(table => tableFilterPasses(table, row, sheet.cells))
  ), [sheet.rowCount, sheet.cells, sheet.tables, columnFilters, hiddenRows]);
  const bounds = selection ? clampRange(selection) : { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
  const selected = (row: number, col: number) => row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol && col <= bounds.endCol;
  const setRange = (range: CellRange, address = addressOf(range.endRow, range.endCol)) => onSelect(address, range);
  const openTableFilter = (event: ReactMouseEvent<HTMLButtonElement>, tableId: string, col: number) => {
    event.preventDefault(); event.stopPropagation();
    const table = sheet.tables?.find(item => item.id === tableId);
    if (!table) return;
    const values = tableColumnValues(table, col, sheet.cells);
    const current = table.filters?.[col];
    const rect = event.currentTarget.getBoundingClientRect();
    setTableFilterSearch("");
    setTableFilterValues(current?.selectedValues ?? values);
    setTableFilterOperator(current?.condition?.operator ?? "contains");
    setTableFilterConditionValue(current?.condition?.value ?? "");
    setTableFilterConditionSecond(current?.condition?.secondValue ?? "");
    setTableFilterMenu({ tableId, col, x: Math.min(rect.left, window.innerWidth - 300), y: Math.min(rect.bottom + 4, window.innerHeight - 420) });
  };
  const applyTableFilter = () => {
    if (!tableFilterMenu) return;
    const table = sheet.tables?.find(item => item.id === tableFilterMenu.tableId);
    if (!table) return;
    const allValues = tableColumnValues(table, tableFilterMenu.col, sheet.cells);
    const filter: TableFilterState = {};
    if (tableFilterValues.length !== allValues.length) filter.selectedValues = tableFilterValues;
    if (tableFilterConditionValue.trim()) filter.condition = { operator: tableFilterOperator, value: tableFilterConditionValue.trim(), ...(tableFilterOperator === "between" ? { secondValue: tableFilterConditionSecond.trim() } : {}) };
    const filters = { ...(table.filters ?? {}) };
    if (filter.selectedValues || filter.condition) filters[tableFilterMenu.col] = filter;
    else delete filters[tableFilterMenu.col];
    updateSheet(sheet.id, { tables: (sheet.tables ?? []).map(item => item.id === table.id ? { ...item, filters } : item) });
    setTableFilterMenu(null);
  };
  const sortTableColumn = (tableId: string, col: number, descending: boolean) => {
    const table = sheet.tables?.find(item => item.id === tableId);
    if (!table) return;
    const orderedRows = Array.from({ length: Math.max(0, table.range.endRow - table.range.startRow) }, (_, i) => table.range.startRow + i + 1);
    orderedRows.sort((a, b) => {
      const left = sheet.cells[addressOf(a, col)]?.raw ?? "", right = sheet.cells[addressOf(b, col)]?.raw ?? "";
      const ln = left.trim() === "" ? null : Number(left), rn = right.trim() === "" ? null : Number(right);
      const compare = ln !== null && rn !== null && Number.isFinite(ln) && Number.isFinite(rn) ? ln - rn : left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
      return descending ? -compare : compare;
    });
    const source = orderedRows.map(row => Array.from({ length: table.range.endCol - table.range.startCol + 1 }, (_, offset) => sheet.cells[addressOf(row, table.range.startCol + offset)]));
    const next = { ...sheet.cells };
    orderedRows.forEach((row, index) => source[index].forEach((cell, offset) => {
      const key = addressOf(row, table.range.startCol + offset);
      if (cell) next[key] = cell; else delete next[key];
    }));
    updateSheet(sheet.id, { cells: next });
    setTableFilterMenu(null);
  };
  const finishFill = (preview?: CellRange, requestedAxis?: FillAxis) => {
    const source = fillSource.current, target = preview ?? fillPreview;
    const axis = requestedAxis ?? fillAxis.current ?? "vertical";
    const useSeriesModifier = fillModifier.current;
    fillSource.current = null; fillAxis.current = null; fillModifier.current = false; fillDragRef.current = null;
    setFillPreview(null); setFillDrag(null); setInteractionState(formulaBarFocused ? "EDITING" : "READY");
    if (!source || !target) return;
    const original = clampRange(source), destination = clampRange(target), next = { ...sheet.cells };
    const vertical = axis === "vertical";
    const sourceLength = vertical ? original.endRow - original.startRow + 1 : original.endCol - original.startCol + 1;
    const before = vertical ? destination.startRow < original.startRow : destination.startCol < original.startCol;
    const after = vertical ? destination.endRow > original.endRow : destination.endCol > original.endCol;
    const extendsRange = before || after;
    if (vertical) {
      for (let row = original.startRow; row <= original.endRow; row++) {
        if (row >= destination.startRow && row <= destination.endRow) continue;
        for (let col = original.startCol; col <= original.endCol; col++) {
          const key = addressOf(row, col); next[key] = { ...(next[key] ?? { raw: "" }), raw: "" };
        }
      }
    } else {
      for (let col = original.startCol; col <= original.endCol; col++) {
        if (col >= destination.startCol && col <= destination.endCol) continue;
        for (let row = original.startRow; row <= original.endRow; row++) {
          const key = addressOf(row, col); next[key] = { ...(next[key] ?? { raw: "" }), raw: "" };
        }
      }
    }
    if (extendsRange) {
      const monthNames = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
      const weekNames = ["sun","mon","tue","wed","thu","fri","sat"];
      for (let row = destination.startRow; row <= destination.endRow; row++) {
        for (let col = destination.startCol; col <= destination.endCol; col++) {
          if (row >= original.startRow && row <= original.endRow && col >= original.startCol && col <= original.endCol) continue;
          const coordinate = vertical ? row : col, orthogonal = vertical ? col : row, baseIndex = before ? 0 : sourceLength - 1;
          const sourceRow = vertical ? original.startRow + baseIndex : orthogonal;
          const sourceCol = vertical ? orthogonal : original.startCol + baseIndex;
          const sequence = Array.from({ length: sourceLength }, (_, i) => {
            const sr = vertical ? original.startRow + i : orthogonal, sc = vertical ? orthogonal : original.startCol + i;
            return sheet.cells[addressOf(sr, sc)]?.raw ?? "";
          });
          const template = sheet.cells[addressOf(sourceRow, sourceCol)] ?? { raw: "" };
          let raw = template.raw;
          const boundary = vertical ? (before ? original.startRow : original.endRow) : (before ? original.startCol : original.endCol);
          const offset = Math.max(0, before ? boundary - coordinate - 1 : coordinate - boundary - 1), sign = before ? -1 : 1;
          const numbers = sequence.map(Number), numeric = sequence.every(value => value.trim() !== "" && Number.isFinite(Number(value)));
          const step = sourceLength > 1 && numeric ? numbers[numbers.length - 1] - numbers[numbers.length - 2] : 1;
          const base = sequence[baseIndex] ?? raw, suffix = base.match(/^(.*?)(\d+)$/);
          const monthSeries = monthNames.includes(base.toLowerCase().slice(0, 3)) ? monthNames : null;
          const weekSeries = weekNames.includes(base.toLowerCase().slice(0, 3)) ? weekNames : null;
          if (raw.startsWith("=")) raw = shiftFormulaReferences(raw, row - sourceRow, col - sourceCol);
          else if (numeric && (sourceLength > 1 || useSeriesModifier)) raw = String(Number(base) + step * sign * (offset + 1));
          else if (monthSeries || weekSeries) {
            const series = monthSeries ?? weekSeries!, index = series.indexOf(base.toLowerCase().slice(0, 3));
            const item = series[((index + sign * (offset + 1)) % series.length + series.length) % series.length];
            raw = base.replace(/^([A-Za-z]{3})/i, item[0].toUpperCase() + item.slice(1));
          } else if (/^\d{4}-\d{2}-\d{2}$/.test(base)) {
            const date = new Date(`${base}T00:00:00Z`); date.setUTCDate(date.getUTCDate() + sign * (offset + 1)); raw = date.toISOString().slice(0, 10);
          } else if (suffix) raw = `${suffix[1]}${Number(suffix[2]) + sign * (offset + 1)}`;
          next[addressOf(row, col)] = { ...template, raw };
        }
      }
    }
    updateSheet(sheet.id, { cells: next, rowCount: Math.max(sheet.rowCount, destination.endRow + 1), columnCount: Math.max(sheet.columnCount, destination.endCol + 1) });
    const anchor = parseAddress(activeCell);
    const keepAnchor = anchor && anchor.row >= destination.startRow && anchor.row <= destination.endRow && anchor.col >= destination.startCol && anchor.col <= destination.endCol;
    setRange(destination, keepAnchor ? activeCell : addressOf(destination.startRow, destination.startCol));
  };
  const focusCell = (row: number, col: number) => requestAnimationFrame(() => gridRef.current?.querySelector<HTMLInputElement>(`input[data-cell="${addressOf(row, col)}"]`)?.focus());
  const previewFillValue = (source: CellRange, target: CellRange, axis: FillAxis) => {
    const original = clampRange(source), destination = clampRange(target), vertical = axis === "vertical";
    const before = vertical ? destination.startRow < original.startRow : destination.startCol < original.startCol;
    const sourceLength = vertical ? original.endRow - original.startRow + 1 : original.endCol - original.startCol + 1;
    const index = before ? 0 : sourceLength - 1, orthogonal = vertical ? original.startCol : original.startRow;
    const sequence = Array.from({ length: sourceLength }, (_, i) => sheet.cells[addressOf(vertical ? original.startRow + i : orthogonal, vertical ? orthogonal : original.startCol + i)]?.raw ?? "");
    const base = sequence[index] ?? "";
    const coordinate = vertical ? (before ? destination.startRow : destination.endRow) : (before ? destination.startCol : destination.endCol);
    const boundary = vertical ? (before ? original.startRow : original.endRow) : (before ? original.startCol : original.endCol);
    const offset = Math.max(0, Math.abs(coordinate - boundary) - 1), numbers = sequence.map(Number);
    const months = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    const weekdays = ["sun","mon","tue","wed","thu","fri","sat"];
    const monthSeries = months.includes(base.toLowerCase().slice(0, 3)) ? months : null;
    const weekSeries = weekdays.includes(base.toLowerCase().slice(0, 3)) ? weekdays : null;
    if (sequence.every(value => value.trim() !== "" && Number.isFinite(Number(value))) && (sourceLength > 1 || fillModifier.current)) {
      const step = sourceLength > 1 ? numbers.at(-1)! - numbers.at(-2)! : 1;
      return String(Number(base) + step * (before ? -1 : 1) * (offset + 1));
    }
    if (monthSeries || weekSeries) {
      const series = monthSeries ?? weekSeries!, index = series.indexOf(base.toLowerCase().slice(0, 3));
      const item = series[((index + (before ? -1 : 1) * (offset + 1)) % series.length + series.length) % series.length];
      return base.replace(/^([A-Za-z]{3})/i, item[0].toUpperCase() + item.slice(1));
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(base)) {
      const date = new Date(`${base}T00:00:00Z`);
      date.setUTCDate(date.getUTCDate() + (before ? -1 : 1) * (offset + 1));
      return date.toISOString().slice(0, 10);
    }
    const suffix = base.match(/^(.*?)(\d+)$/);
    return suffix ? `${suffix[1]}${Number(suffix[2]) + (before ? -1 : 1) * (offset + 1)}` : base;
  };
  const measureRange = (range: CellRange | null): OverlayRect | null => {
    const canvas = canvasRef.current; if (!canvas || !range) return null;
    const normalized = clampRange(range), visibleRows = rows.filter(row => row >= normalized.startRow && row <= normalized.endRow), visibleColumns = columns.filter(col => col >= normalized.startCol && col <= normalized.endCol);
    if (!visibleRows.length || !visibleColumns.length) return null;
    const firstAddress = addressOf(visibleRows[0], visibleColumns[0]), lastAddress = addressOf(visibleRows.at(-1)!, visibleColumns.at(-1)!);
    const firstCell = canvas.querySelector<HTMLElement>(`td[data-cell="${firstAddress}"]`) ?? canvas.querySelector<HTMLElement>(`input[data-cell="${firstAddress}"]`)?.closest("td");
    const lastCell = canvas.querySelector<HTMLElement>(`td[data-cell="${lastAddress}"]`) ?? canvas.querySelector<HTMLElement>(`input[data-cell="${lastAddress}"]`)?.closest("td");
    if (!firstCell || !lastCell) return null;
    const parentRect = canvas.getBoundingClientRect(), firstRect = firstCell.getBoundingClientRect(), lastRect = lastCell.getBoundingClientRect(), scale = Math.max(0.01, zoom / 100);
    return { left: (firstRect.left - parentRect.left) / scale, top: (firstRect.top - parentRect.top) / scale, width: (lastRect.right - firstRect.left) / scale, height: (lastRect.bottom - firstRect.top) / scale };
  };
  const updateOverlayRects = () => setOverlayRects({ selection: measureRange(bounds), fill: measureRange(fillPreview), copy: measureRange(copyRange) });
  useLayoutEffect(() => {
    const frame = requestAnimationFrame(updateOverlayRects);
    return () => cancelAnimationFrame(frame);
  }, [bounds.startRow, bounds.endRow, bounds.startCol, bounds.endCol, fillPreview, copyRange, rows, columns, sheet.columnWidths, rowHeights, zoom, sheet.showHeadings]);
  const updateDimensions = (row: number, col: number) => updateSheet(sheet.id, { rowCount: Math.max(sheet.rowCount, row + 1), columnCount: Math.max(sheet.columnCount, col + 1) });
  const updateRowHeight = (row: number, height: number) => updateSheet(sheet.id, { rowHeights: { ...(sheet.rowHeights ?? {}), [row]: Math.max(18, Math.min(120, height)) } });

  const beginResize = (event: ReactMouseEvent, axis: "row" | "column", index: number) => {
    event.preventDefault(); event.stopPropagation();
    const origin = axis === "row" ? event.clientY : event.clientX;
    const measured = (event.currentTarget as HTMLElement).parentElement?.getBoundingClientRect();
    const initial = axis === "row" ? rowHeights[index] ?? measured?.height ?? 27 : sheet.columnWidths[index] ?? measured?.width ?? 110;
    const move = (e: globalThis.MouseEvent) => {
      const delta = ((axis === "row" ? e.clientY : e.clientX) - origin) / (zoom / 100);
      const value = Math.max(axis === "row" ? 18 : 48, Math.min(axis === "row" ? 120 : 480, initial + delta));
      resizeValue.current = { axis, index, value }; setResizePreview(resizeValue.current);
    };
    const end = () => {
      window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", end);
      const final = resizeValue.current;
      if (final?.axis === "row") updateRowHeight(final.index, final.value);
      else if (final) updateSheet(sheet.id, { columnWidths: { ...sheet.columnWidths, [final.index]: final.value } });
      resizeValue.current = null; setResizePreview(null);
    };
    window.addEventListener("mousemove", move); window.addEventListener("mouseup", end);
  };
  const autoFitColumn = (col: number) => {
    const canvas = document.createElement("canvas"), context = canvas.getContext("2d");
    const width = Array.from({ length: Math.max(sheet.rowCount, 100) }, (_, row) => row).reduce((max, row) => {
      const cell = sheet.cells[addressOf(row, col)], style = cell?.style ?? {};
      const display = formatCellValue(cell?.raw ?? "", sheet.cells, false, style, sheet.tables ?? [], addressOf(row, col), namedRanges);
      if (context) context.font = `${style.italic ? "italic " : ""}${style.bold ? "700 " : ""}${style.fontSize ?? 11}px ${style.fontFamily ?? "Inter"}`;
      const measured = display.split("\n").reduce((lineMax, line) => Math.max(lineMax, context?.measureText(line).width ?? line.length * 7.2), 0) + 22;
      return Math.max(max, measured);
    }, 48);
    updateSheet(sheet.id, { columnWidths: { ...sheet.columnWidths, [col]: Math.min(480, Math.max(48, Math.ceil(width))) } });
  };
  const autoFitRow = (row: number) => {
    const height = columns.reduce((max, col) => {
      const cell = sheet.cells[addressOf(row, col)], style = cell?.style ?? {};
      const text = formatCellValue(cell?.raw ?? "", sheet.cells, false, style, sheet.tables ?? [], addressOf(row, col)), width = sheet.columnWidths[col] ?? 110;
      const canvas = document.createElement("canvas"), context = canvas.getContext("2d");
      if (context) context.font = `${style.italic ? "italic " : ""}${style.bold ? "700 " : ""}${style.fontSize ?? 11}px ${style.fontFamily ?? "Inter"}`;
      const lines = text.split("\n").reduce((sum, line) => sum + (style.wrap ? Math.max(1, Math.ceil((context?.measureText(line).width ?? line.length * 7.2) / Math.max(24, width - 12))) : 1), 0);
      return Math.max(max, lines * ((style.fontSize ?? 11) * 1.45) + 10);
    }, 27);
    updateRowHeight(row, Math.min(409, Math.ceil(height)));
  };
  const beginRange = (row: number, col: number, event: ReactPointerEvent<HTMLInputElement>) => {
    event.preventDefault();
    const anchor = event.shiftKey ? active : { row, col };
    setEditingAddress(null); setInteractionState("SELECTING");
    pointerPosition.current = { x: event.clientX, y: event.clientY };
    lastDragAddress.current = addressOf(row, col);
    setDragStart(anchor);
    setRange({ startRow: anchor.row, startCol: anchor.col, endRow: row, endCol: col }, addressOf(anchor.row, anchor.col));
  };
  const updatePointerTarget = (x: number, y: number) => {
    pointerPosition.current = { x, y };
    if (interactionState === "FILLING") {
      const current = fillDragRef.current; if (!current) return;
      let axis = current.axis;
      if (!axis) {
        const dx = Math.abs(x - current.startX), dy = Math.abs(y - current.startY);
        if (Math.max(dx, dy) < 4) return;
        axis = dy >= dx ? "vertical" : "horizontal"; fillAxis.current = axis;
      }
      const element = document.elementFromPoint(x, y)?.closest<HTMLElement>("[data-cell]");
      const position = element ? parseAddress(element.dataset.cell ?? "") : null;
      if (!position) return;
      const source = clampRange(current.source);
      const target: CellRange = axis === "vertical"
        ? { startRow: position.row < source.startRow ? position.row : source.startRow, endRow: position.row < source.startRow ? source.endRow : position.row, startCol: source.startCol, endCol: source.endCol }
        : { startRow: source.startRow, endRow: source.endRow, startCol: position.col < source.startCol ? position.col : source.startCol, endCol: position.col < source.startCol ? source.endCol : position.col };
      const next: FillDragState = { ...current, axis, target, clientX: x, clientY: y, previewValue: previewFillValue(current.source, target, axis) };
      fillDragRef.current = next; setFillDrag(next); setFillPreview(target); return;
    }
    if (interactionState === "SELECTING" && dragStart) {
      const element = document.elementFromPoint(x, y)?.closest<HTMLElement>("[data-cell]");
      const position = element ? parseAddress(element.dataset.cell ?? "") : null;
      if (!position) return;
      const address = addressOf(position.row, position.col);
      if (lastDragAddress.current === address) return;
      lastDragAddress.current = address;
      setRange({ startRow: dragStart.row, startCol: dragStart.col, endRow: position.row, endCol: position.col }, addressOf(dragStart.row, dragStart.col));
    }
  };
  const beginFillDrag = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault(); event.stopPropagation();
    setEditingAddress(null); setInteractionState("FILLING");
    const source = clampRange(bounds);
    fillSource.current = source; fillAxis.current = null; fillModifier.current = event.ctrlKey || event.altKey;
    pointerPosition.current = { x: event.clientX, y: event.clientY };
    const initial: FillDragState = { source, target: source, axis: null, startX: event.clientX, startY: event.clientY, clientX: event.clientX, clientY: event.clientY, previewValue: sheet.cells[addressOf(source.endRow, source.endCol)]?.raw ?? "" };
    fillDragRef.current = initial; setFillDrag(initial); setFillPreview(null);
    try { event.currentTarget.setPointerCapture(event.pointerId); } catch { /* Window pointer listeners also track the drag. */ }
  };
  interactionHandlersRef.current = {
    update: updatePointerTarget,
    finish: () => {
      if (interactionState === "FILLING") {
        const current = fillDragRef.current; finishFill(current?.target, current?.axis ?? undefined);
      } else if (interactionState === "SELECTING") {
        setDragStart(null); lastDragAddress.current = null;
        setInteractionState(formulaBarFocused ? "EDITING" : "READY");
        focusCell(active.row, active.col);
      }
    }
  };
  useEffect(() => {
    if (formulaBarFocused) {
      setEditingAddress(null); setInteractionState("EDITING");
    } else if (!editingAddress && interactionState === "EDITING") setInteractionState("READY");
  }, [formulaBarFocused, editingAddress]);
  const showContextMenu = (event: ReactMouseEvent, kind: NonNullable<ContextMenuState>["kind"], index: number, col = active.col) => {
    event.preventDefault(); event.stopPropagation();
    if (kind === "cell" && !selected(index, col)) onSelect(addressOf(index, col), { startRow: index, startCol: col, endRow: index, endCol: col });
    if (kind === "row" && !(index >= bounds.startRow && index <= bounds.endRow && bounds.startCol === 0 && bounds.endCol >= Math.max(25, sheet.columnCount - 1))) onSelect(addressOf(index, 0), { startRow: index, startCol: 0, endRow: index, endCol: Math.max(25, sheet.columnCount - 1) });
    if (kind === "column" && !(index >= bounds.startCol && index <= bounds.endCol && bounds.startRow === 0 && bounds.endRow >= Math.max(99, sheet.rowCount - 1))) onSelect(addressOf(0, index), { startRow: 0, startCol: index, endRow: Math.max(99, sheet.rowCount - 1), endCol: index });
    setMenu({ x: Math.max(4, Math.min(event.clientX, window.innerWidth - 260)), y: Math.max(4, Math.min(event.clientY, window.innerHeight - 520)), kind, index, col });
  };
  const openKeyboardContextMenu = (row: number, col: number) => {
    const address = addressOf(row, col), input = gridRef.current?.querySelector<HTMLInputElement>(`[data-cell="${address}"]`), rect = input?.getBoundingClientRect();
    if (!selected(row, col)) onSelect(address, { startRow: row, endRow: row, startCol: col, endCol: col });
    setMenu({ x: rect?.left ?? 180, y: rect?.bottom ?? 190, kind: "cell", index: row, col });
  };
  const applyRangeBorder = (mode: "all" | "outer" | "bottom" | "right" | "clear") => {
    const next = { ...sheet.cells }, color = sheet.cells[activeCell]?.style?.borderColor ?? "#64748b", border = `1px solid ${color}`;
    for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) {
      const key = addressOf(row, col), cell = next[key] ?? { raw: "" }, style = { ...(cell.style ?? {}) };
      if (mode === "clear") { delete style.border; delete style.borderTop; delete style.borderRight; delete style.borderBottom; delete style.borderLeft; }
      else if (mode === "all") { style.border = "all"; style.borderColor = color; }
      else if (mode === "outer") { delete style.border; style.borderTop = row === bounds.startRow ? border : style.borderTop; style.borderBottom = row === bounds.endRow ? border : style.borderBottom; style.borderLeft = col === bounds.startCol ? border : style.borderLeft; style.borderRight = col === bounds.endCol ? border : style.borderRight; }
      else if (mode === "bottom") style.borderBottom = border;
      else if (mode === "right") style.borderRight = border;
      next[key] = { ...cell, style };
    }
    updateSheet(sheet.id, { cells: next });
  };
  const copyToStore = (cut = false) => {
    setCopyRange(bounds);
    const cells = Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, row) => Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, col) => {
      const cell = sheet.cells[addressOf(bounds.startRow + row, bounds.startCol + col)];
      return { raw: cell?.raw ?? "", style: cell?.style ? { ...cell.style } : undefined };
    }));
    setClipboard({ cells, source: addressOf(bounds.startRow, bounds.startCol) });
    if (cut) {
      const next = { ...sheet.cells };
      for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) next[addressOf(row, col)] = { ...(next[addressOf(row, col)] ?? { raw: "" }), raw: "" };
      updateSheet(sheet.id, { cells: next });
    }
  };
  const pasteInternal = (mode: "all" | "values" | "formulas" | "formats" | "transpose" = "all") => {
    if (!clipboard?.cells?.length) { void pasteFromSystem(mode === "values"); return; }
    const source = parseAddress(clipboard.source ?? activeCell) ?? active;
    const matrix = mode === "transpose" ? clipboard.cells[0].map((_, col) => clipboard.cells.map(row => row[col] ?? { raw: "" })) : clipboard.cells;
    const next = { ...sheet.cells };
    matrix.forEach((line, row) => line.forEach((item, col) => {
      const address = addressOf(active.row + row, active.col + col), old = next[address] ?? { raw: "" };
      const rowDelta = active.row + row - source.row, colDelta = active.col + col - source.col;
      if (mode === "formats") next[address] = { ...old, style: item.style ? { ...item.style } : {} };
      else if (mode === "formulas") { if (item.raw.startsWith("=")) next[address] = { ...old, raw: shiftFormulaReferences(item.raw, rowDelta, colDelta) }; }
      else if (mode === "values") next[address] = { ...old, raw: item.raw.startsWith("=") ? String(formatCellValue(item.raw, sheet.cells, false, undefined, sheet.tables ?? [], addressOf(source.row + row, source.col + col), namedRanges)) : item.raw };
      else { const raw = item.raw.startsWith("=") ? shiftFormulaReferences(item.raw, rowDelta, colDelta) : item.raw; next[address] = { raw, style: item.style ? { ...item.style } : undefined }; }
    }));
    updateSheet(sheet.id, { cells: next, rowCount: Math.max(sheet.rowCount, active.row + matrix.length), columnCount: Math.max(sheet.columnCount, active.col + (matrix[0]?.length ?? 1)) });
  };
  const pasteFromSystem = async (valuesOnly = false) => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        const matrix = text.replace(/\r/g, "").replace(/\n$/, "").split("\n").map(row => row.split("\t"));
        const next = { ...sheet.cells };
        matrix.forEach((line, r) => line.forEach((raw, c) => { next[addressOf(active.row + r, active.col + c)] = { ...(next[addressOf(active.row + r, active.col + c)] ?? { raw: "" }), raw }; }));
        updateSheet(sheet.id, { cells: next, rowCount: Math.max(sheet.rowCount, active.row + matrix.length), columnCount: Math.max(sheet.columnCount, active.col + Math.max(...matrix.map(row => row.length))) });
        return;
      }
    } catch { /* Clipboard permission may be unavailable; use the in-app clipboard. */ }
    if (!clipboard) return;
    if (!valuesOnly) useWorkbookStore.getState().paste(activeCell);
    else {
      const next = { ...sheet.cells };
      clipboard.cells.forEach((line, row) => line.forEach((item, col) => { const key = addressOf(active.row + row, active.col + col); next[key] = { ...(next[key] ?? { raw: "" }), raw: item.raw }; }));
      updateSheet(sheet.id, { cells: next, rowCount: Math.max(sheet.rowCount, active.row + clipboard.cells.length), columnCount: Math.max(sheet.columnCount, active.col + (clipboard.cells[0]?.length ?? 1)) });
    }
  };
  const copySystem = async (cut = false) => {
    copyToStore(cut);
    try { await navigator.clipboard.writeText(Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, r) => Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, c) => sheet.cells[addressOf(bounds.startRow + r, bounds.startCol + c)]?.raw ?? "").join("\t")).join("\n")); } catch { /* The internal clipboard remains available. */ }
  };
  const shiftRows = (index: number, count: number, insert: boolean) => {
    const cells: SheetData["cells"] = {};
    Object.entries(sheet.cells).forEach(([key, cell]) => {
      const pos = parseAddress(key); if (!pos) return;
      if (!insert && pos.row >= index && pos.row < index + count) return;
      const row = insert && pos.row >= index ? pos.row + count : !insert && pos.row >= index + count ? pos.row - count : pos.row;
      cells[addressOf(row, pos.col)] = cell;
    });
    Object.entries(cells).forEach(([key, cell]) => { if (cell.raw.startsWith("=")) cells[key] = { ...cell, raw: adjustReferences(cell.raw, "row", index, count, insert) }; });
    const tables = (sheet.tables ?? []).flatMap(table => {
      let startRow = table.range.startRow, endRow = table.range.endRow;
      if (insert) {
        if (index <= startRow) { startRow += count; endRow += count; }
        else if (index <= endRow + 1) endRow += count;
      } else {
        const deletedEnd = index + count - 1;
        if (deletedEnd < startRow) { startRow -= count; endRow -= count; }
        else if (index <= endRow) {
          if (index <= startRow) return [];
          const removed = Math.max(0, Math.min(endRow, deletedEnd) - Math.max(startRow, index) + 1);
          endRow -= removed;
          if (endRow < startRow) return [];
        }
      }
      return [{ ...table, range: { ...table.range, startRow, endRow } }];
    });
    updateSheet(sheet.id, { cells, tables, rowCount: Math.max(100, sheet.rowCount + (insert ? count : -count)) });
  };
  const shiftColumns = (index: number, count: number, insert: boolean) => {
    const cells: SheetData["cells"] = {};
    Object.entries(sheet.cells).forEach(([key, cell]) => {
      const pos = parseAddress(key); if (!pos) return;
      if (!insert && pos.col >= index && pos.col < index + count) return;
      const col = insert && pos.col >= index ? pos.col + count : !insert && pos.col >= index + count ? pos.col - count : pos.col;
      cells[addressOf(pos.row, col)] = cell;
    });
    Object.entries(cells).forEach(([key, cell]) => { if (cell.raw.startsWith("=")) cells[key] = { ...cell, raw: adjustReferences(cell.raw, "column", index, count, insert) }; });
    const widths: Record<number, number> = {};
    Object.entries(sheet.columnWidths).forEach(([key, width]) => { const col = Number(key); if (!insert && col >= index && col < index + count) return; widths[insert && col >= index ? col + count : !insert && col >= index + count ? col - count : col] = width; });
    const tables = (sheet.tables ?? []).flatMap(table => {
      let startCol = table.range.startCol, endCol = table.range.endCol;
      if (insert) {
        if (index <= startCol) { startCol += count; endCol += count; }
        else if (index <= endCol + 1) endCol += count;
      } else {
        const deletedEnd = index + count - 1;
        if (deletedEnd < startCol) { startCol -= count; endCol -= count; }
        else if (index <= endCol) {
          if (index <= startCol) return [];
          const removed = Math.max(0, Math.min(endCol, deletedEnd) - Math.max(startCol, index) + 1);
          endCol -= removed;
          if (endCol < startCol) return [];
        }
      }
      return [{ ...table, range: { ...table.range, startCol, endCol } }];
    });
    updateSheet(sheet.id, { cells, columnWidths: widths, tables, columnCount: Math.max(26, sheet.columnCount + (insert ? count : -count)) });
  };
  const shiftSelectedCells = (direction: "down" | "right" | "up" | "left") => {
    const rowCount = bounds.endRow - bounds.startRow + 1, colCount = bounds.endCol - bounds.startCol + 1;
    const next: SheetData["cells"] = {};
    Object.entries(sheet.cells).forEach(([key, cell]) => {
      const pos = parseAddress(key); if (!pos) return;
      let row = pos.row, col = pos.col, remove = false;
      if (direction === "down" && col >= bounds.startCol && col <= bounds.endCol && row >= bounds.startRow) row += rowCount;
      if (direction === "up" && col >= bounds.startCol && col <= bounds.endCol && row >= bounds.startRow) { if (row <= bounds.endRow) remove = true; else row -= rowCount; }
      if (direction === "right" && row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol) col += colCount;
      if (direction === "left" && row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol) { if (col <= bounds.endCol) remove = true; else col -= colCount; }
      if (!remove) next[addressOf(row, col)] = cell;
    });
    updateSheet(sheet.id, { cells: next, rowCount: Math.max(100, sheet.rowCount + (direction === "down" ? rowCount : direction === "up" ? -rowCount : 0)), columnCount: Math.max(26, sheet.columnCount + (direction === "right" ? colCount : direction === "left" ? -colCount : 0)) });
  };
  const clearSelection = (format = false, all = false) => {
    const next = { ...sheet.cells };
    for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) {
      const key = addressOf(row, col);
      if (all) next[key] = { raw: "" };
      else if (format) next[key] = { ...(next[key] ?? { raw: "" }), style: {} };
      else next[key] = { ...(next[key] ?? { raw: "" }), raw: "" };
    }
    updateSheet(sheet.id, { cells: next });
  };
  const menuAction = (action: string) => {
    if (!menu) return;
    const countRows = bounds.endRow - bounds.startRow + 1;
    const countCols = bounds.endCol - bounds.startCol + 1;
    const contextTable = getTableAt(sheet.tables, menu.index, menu.col);
    if (action === "cut") void copySystem(true);
    else if (action === "copy") void copySystem();
    else if (action === "paste") clipboard ? pasteInternal("all") : void pasteFromSystem();
    else if (action === "paste-values") pasteInternal("values");
    else if (action === "paste-formulas") pasteInternal("formulas");
    else if (action === "paste-format") pasteInternal("formats");
    else if (action === "paste-transpose") pasteInternal("transpose");
    else if (action === "format-cells") { setFormatDraft({ ...(sheet.cells[activeCell]?.style ?? {}) }); setFormatDialogOpen(true); }
    else if (action === "show-shortcuts") setShortcutsOpen(true);
    else if (action === "clear-all") clearSelection(false, true);
    else if (action === "insert-cells-down") shiftSelectedCells("down");
    else if (action === "insert-cells-right") shiftSelectedCells("right");
    else if (action === "delete-cells-up") shiftSelectedCells("up");
    else if (action === "delete-cells-left") shiftSelectedCells("left");
    else if (action === "table-insert-row-above" && contextTable) shiftRows(menu.index, 1, true);
    else if (action === "table-insert-row-below" && contextTable) shiftRows(menu.index + 1, 1, true);
    else if (action === "table-delete-row" && contextTable) shiftRows(menu.index, 1, false);
    else if (action === "table-insert-column-left" && contextTable) shiftColumns(menu.col, 1, true);
    else if (action === "table-insert-column-right" && contextTable) shiftColumns(menu.col + 1, 1, true);
    else if (action === "table-delete-column" && contextTable) shiftColumns(menu.col, 1, false);
    else if (action === "table-clear-filters" && contextTable) updateSheet(sheet.id, { tables: (sheet.tables ?? []).map(table => table.id === contextTable.id ? { ...table, filters: {} } : table) });
    else if (action === "table-remove-duplicates" && contextTable) {
      const seen = new Set<string>(), keep: Array<Record<string, SheetData["cells"][string]>> = [];
      for (let row = contextTable.range.startRow + 1; row <= contextTable.range.endRow; row++) {
        const signature = Array.from({ length: contextTable.range.endCol - contextTable.range.startCol + 1 }, (_, offset) => sheet.cells[addressOf(row, contextTable.range.startCol + offset)]?.raw ?? "").join("\u001f");
        if (seen.has(signature)) continue;
        seen.add(signature);
        keep.push(Object.fromEntries(Array.from({ length: contextTable.range.endCol - contextTable.range.startCol + 1 }, (_, offset) => [addressOf(row, contextTable.range.startCol + offset), sheet.cells[addressOf(row, contextTable.range.startCol + offset)] ?? { raw: "" }])));
      }
      const next = { ...sheet.cells }, firstData = contextTable.range.startRow + 1;
      for (let row = firstData; row <= contextTable.range.endRow; row++) for (let col = contextTable.range.startCol; col <= contextTable.range.endCol; col++) delete next[addressOf(row, col)];
      keep.forEach((rowCells, offset) => Object.entries(rowCells).forEach(([source, cell]) => { const position = parseAddress(source)!; next[addressOf(firstData + offset, position.col)] = cell; }));
      const tables = (sheet.tables ?? []).map(table => table.id === contextTable.id ? { ...table, range: { ...table.range, endRow: firstData + keep.length - 1 } } : table);
      updateSheet(sheet.id, { cells: next, tables });
    }
    else if (action === "table-convert-range" && contextTable) {
      const palette = getTablePalette(contextTable.styleId), next = { ...sheet.cells };
      for (let row = contextTable.range.startRow; row <= contextTable.range.endRow + (contextTable.showTotalRow ? 1 : 0); row++) {
        for (let col = contextTable.range.startCol; col <= contextTable.range.endCol; col++) {
          const key = addressOf(row, col), cell = next[key] ?? { raw: "" }, style = { ...(cell.style ?? {}) };
          if (!style.backgroundColor) style.backgroundColor = row === contextTable.range.startRow ? palette.header : contextTable.bandedRows && (row - contextTable.range.startRow) % 2 === 0 ? palette.band : "#ffffff";
          if (row === contextTable.range.startRow && !style.color) style.color = palette.text;
          if (row === contextTable.range.startRow) style.bold = true;
          next[key] = { ...cell, style };
        }
      }
      updateSheet(sheet.id, { cells: next, tables: (sheet.tables ?? []).filter(table => table.id !== contextTable.id) });
    }
    else if (action === "border-all") applyRangeBorder("all");
    else if (action === "border-outer") applyRangeBorder("outer");
    else if (action === "border-bottom") applyRangeBorder("bottom");
    else if (action === "border-right") applyRangeBorder("right");
    else if (action === "border-clear") applyRangeBorder("clear");
    else if (action === "hide-row") updateSheet(sheet.id, { hiddenRows: [...new Set([...(sheet.hiddenRows ?? []), ...(menu.kind === "row" ? [menu.index] : Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, i) => bounds.startRow + i))])] });
    else if (action === "unhide-rows") updateSheet(sheet.id, { hiddenRows: [] });
    else if (action === "hide-column") updateSheet(sheet.id, { hiddenColumns: [...new Set([...(sheet.hiddenColumns ?? []), ...(menu.kind === "column" ? [menu.index] : Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, i) => bounds.startCol + i))])] });
    else if (action === "unhide-columns") updateSheet(sheet.id, { hiddenColumns: [] });
    else if (action === "insert-row-above") shiftRows(menu.kind === "row" ? menu.index : bounds.startRow, countRows, true);
    else if (action === "insert-row-below") shiftRows((menu.kind === "row" ? menu.index : bounds.endRow) + 1, countRows, true);
    else if (action === "delete-row") shiftRows(menu.kind === "row" ? menu.index : bounds.startRow, menu.kind === "row" ? countRows : 1, false);
    else if (action === "insert-col-left") shiftColumns(menu.kind === "column" ? menu.index : bounds.startCol, countCols, true);
    else if (action === "insert-col-right") shiftColumns((menu.kind === "column" ? menu.index : bounds.endCol) + 1, countCols, true);
    else if (action === "delete-col") shiftColumns(menu.kind === "column" ? menu.index : bounds.startCol, menu.kind === "column" ? countCols : 1, false);
    else if (action === "fill-down" || action === "fill-right") {
      const down = action === "fill-down", next = { ...sheet.cells };
      for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) {
        if (down ? row === bounds.startRow : col === bounds.startCol) continue;
        const sourceRow = down ? bounds.startRow : row, sourceCol = down ? col : bounds.startCol, source = sheet.cells[addressOf(sourceRow, sourceCol)] ?? { raw: "" };
        const raw = source.raw.startsWith("=") ? shiftFormulaReferences(source.raw, row - sourceRow, col - sourceCol) : source.raw;
        next[addressOf(row, col)] = { ...source, raw };
      }
      updateSheet(sheet.id, { cells: next });
    }
    else if (action === "clear") clearSelection();
    else if (action === "clear-format") clearSelection(true);
    else if (action === "resize-row") { const height = Number(window.prompt("Row height in pixels", String(rowHeights[menu.index] ?? 27))); if (Number.isFinite(height) && height >= 18) updateRowHeight(menu.index, height); }
    else if (action === "autofit-row") autoFitRow(menu.index);
    else if (action === "resize-column") { const width = Number(window.prompt("Column width in pixels", String(sheet.columnWidths[menu.index] ?? 110))); if (Number.isFinite(width) && width >= 48) updateSheet(sheet.id, { columnWidths: { ...sheet.columnWidths, [menu.index]: Math.min(480, width) } }); }
    else if (action === "autofit-column") autoFitColumn(menu.index);
    else if (action === "sort-asc" || action === "sort-desc") {
      const rangeRows = Array.from({ length: countRows }, (_, i) => bounds.startRow + i), startCol = bounds.startCol, endCol = bounds.endCol;
      const sorted = [...rangeRows].sort((a, b) => {
        const left = sheet.cells[addressOf(a, active.col)]?.raw ?? "", right = sheet.cells[addressOf(b, active.col)]?.raw ?? "";
        const ln = Number(left), rn = Number(right), compare = left && right && Number.isFinite(ln) && Number.isFinite(rn) ? ln - rn : left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
        return action === "sort-desc" ? -compare : compare;
      });
      const next = { ...sheet.cells };
      rangeRows.forEach(row => { for (let col = startCol; col <= endCol; col++) delete next[addressOf(row, col)]; });
      sorted.forEach((sourceRow, offset) => { for (let col = startCol; col <= endCol; col++) { const value = sheet.cells[addressOf(sourceRow, col)]; if (value) next[addressOf(bounds.startRow + offset, col)] = value; } });
      updateSheet(sheet.id, { cells: next });
    }
    else if (action === "merge") {
      if (countRows > 1 || countCols > 1) {
        const anchor = addressOf(bounds.startRow, bounds.startCol), next = { ...sheet.cells }, backup: Record<string, SheetData["cells"][string]> = {};
        for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) { const key = addressOf(row, col); backup[key] = next[key] ?? { raw: "" }; }
        next[anchor] = { ...(next[anchor] ?? { raw: "" }), mergeRows: countRows, mergeCols: countCols, mergeBackup: backup };
        for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) { const key = addressOf(row, col); if (key !== anchor) next[key] = { raw: "", mergedInto: anchor }; }
        updateSheet(sheet.id, { cells: next });
      }
    }
    else if (action === "unmerge") {
      const anchor = sheet.cells[activeCell]?.mergedInto ?? activeCell, merged = sheet.cells[anchor];
      if (merged?.mergeBackup) { const next = { ...sheet.cells, ...merged.mergeBackup }, restored = { ...next[anchor] }; delete restored.mergeRows; delete restored.mergeCols; delete restored.mergeBackup; next[anchor] = restored; updateSheet(sheet.id, { cells: next }); }
    }
    setMenu(null); if (action !== "format-cells" && action !== "show-shortcuts") focusCell(active.row, active.col);
  };
  const onKeyDown = (event: KeyboardEvent<HTMLInputElement>, row: number, col: number) => {
    const mod = event.ctrlKey || event.metaKey, address = addressOf(row, col), isEditing = editingAddress === address;
    const hasMultiCellSelection = bounds.startRow !== bounds.endRow || bounds.startCol !== bounds.endCol;
    if (hasMultiCellSelection && (event.key === "Enter" || event.key === "Tab")) {
      event.preventDefault();
      let nextRow = row, nextCol = col;
      if (event.key === "Enter") {
        nextRow += event.shiftKey ? -1 : 1;
        if (nextRow > bounds.endRow) nextRow = bounds.startRow;
        if (nextRow < bounds.startRow) nextRow = bounds.endRow;
      } else {
        nextCol += event.shiftKey ? -1 : 1;
        if (nextCol > bounds.endCol) { nextCol = bounds.startCol; nextRow = row + 1 > bounds.endRow ? bounds.startRow : row + 1; }
        if (nextCol < bounds.startCol) { nextCol = bounds.endCol; nextRow = row - 1 < bounds.startRow ? bounds.endRow : row - 1; }
      }
      if (isEditing) setEditingAddress(null);
      setInteractionState("READY");
      onSelect(addressOf(nextRow, nextCol), bounds);
      focusCell(nextRow, nextCol);
      return;
    }
    if (!isEditing && (event.key === "F10" && event.shiftKey || event.key === "ContextMenu")) { event.preventDefault(); openKeyboardContextMenu(row, col); return; }
    if (!isEditing && event.key === "F2") { event.preventDefault(); editOriginal.current = sheet.cells[address]?.raw ?? ""; setInteractionState("EDITING"); setEditingAddress(address); requestAnimationFrame(() => gridRef.current?.querySelector<HTMLInputElement>(`input[data-cell="${address}"]`)?.focus()); return; }
    if (!isEditing && !mod && !event.altKey && event.key.length === 1) { event.preventDefault(); editOriginal.current = sheet.cells[address]?.raw ?? ""; setCell(address, event.key); setInteractionState("EDITING"); setEditingAddress(address); return; }
    if (mod && event.key === "1") { event.preventDefault(); setFormatDraft({ ...(sheet.cells[address]?.style ?? {}) }); setFormatDialogOpen(true); return; }
    if (event.key === "F10" && event.shiftKey) { event.preventDefault(); openKeyboardContextMenu(row, col); return; }
    if (mod && event.key.toLowerCase() === "c") { event.preventDefault(); void copySystem(); return; }
    if (mod && event.key.toLowerCase() === "x") { event.preventDefault(); void copySystem(true); return; }
    if (mod && event.shiftKey && event.key.toLowerCase() === "v") { event.preventDefault(); void pasteFromSystem(true); return; }
    if (mod && event.key.toLowerCase() === "v") { event.preventDefault(); void pasteFromSystem(); return; }
    if (mod && event.key.toLowerCase() === "z") { event.preventDefault(); event.shiftKey ? redo() : undo(); return; }
    if (mod && event.key.toLowerCase() === "y") { event.preventDefault(); redo(); return; }
    if (mod && event.key.toLowerCase() === "a") { event.preventDefault(); setRange({ startRow: 0, startCol: 0, endRow: Math.max(99, sheet.rowCount - 1), endCol: Math.max(25, sheet.columnCount - 1) }, "A1"); return; }
    if (event.shiftKey && event.key === " ") { event.preventDefault(); setRange({ startRow: row, startCol: 0, endRow: row, endCol: Math.max(25, sheet.columnCount - 1) }, addressOf(row, 0)); return; }
    if (mod && event.key === " ") { event.preventDefault(); setRange({ startRow: 0, startCol: col, endRow: Math.max(99, sheet.rowCount - 1), endCol: col }, addressOf(0, col)); return; }
    if (mod && (event.key.toLowerCase() === "d" || event.key.toLowerCase() === "r")) {
      event.preventDefault(); const down = event.key.toLowerCase() === "d", next = { ...sheet.cells };
      for (let targetRow = bounds.startRow; targetRow <= bounds.endRow; targetRow++) for (let targetCol = bounds.startCol; targetCol <= bounds.endCol; targetCol++) {
        if (down && targetRow === bounds.startRow || !down && targetCol === bounds.startCol) continue;
        const sourceRow = down ? bounds.startRow : targetRow, sourceCol = down ? targetCol : bounds.startCol, source = sheet.cells[addressOf(sourceRow, sourceCol)];
        const raw = source?.raw ?? ""; next[addressOf(targetRow, targetCol)] = { ...(source ?? { raw: "" }), raw: raw.startsWith("=") ? shiftFormulaReferences(raw, targetRow - sourceRow, targetCol - sourceCol) : raw };
      }
      updateSheet(sheet.id, { cells: next }); return;
    }
    if (mod && event.key === ";") { event.preventDefault(); const now = new Date(), value = event.shiftKey ? now.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : now.toISOString().slice(0, 10); setCell(activeCell, value); return; }
    if (mod && event.key.toLowerCase() === "b") { event.preventDefault(); const cell = sheet.cells[activeCell]; useWorkbookStore.getState().setStyle({ bold: !cell?.style?.bold }); return; }
    if (mod && event.key.toLowerCase() === "i") { event.preventDefault(); const cell = sheet.cells[activeCell]; useWorkbookStore.getState().setStyle({ italic: !cell?.style?.italic }); return; }
    if (mod && event.key.toLowerCase() === "u") { event.preventDefault(); const cell = sheet.cells[activeCell]; useWorkbookStore.getState().setStyle({ underline: !cell?.style?.underline }); return; }
    if (mod && event.key === "\\") { event.preventDefault(); clearSelection(true); return; }
    if (mod && event.key === "End") {
      event.preventDefault();
      const used = Object.keys(sheet.cells).map(parseAddress).filter((position): position is { row: number; col: number } => Boolean(position));
      const nr = Math.max(row, ...used.map(position => position.row)), nc = Math.max(col, ...used.map(position => position.col)); updateDimensions(nr, nc); onSelect(addressOf(nr, nc), { startRow: nr, startCol: nc, endRow: nr, endCol: nc }); focusCell(nr, nc); return;
    }
    if (mod && (event.key === "ArrowRight" || event.key === "ArrowLeft" || event.key === "ArrowDown" || event.key === "ArrowUp")) {
      event.preventDefault();
      const vertical = event.key === "ArrowDown" || event.key === "ArrowUp", step = event.key === "ArrowDown" || event.key === "ArrowRight" ? 1 : -1;
      const limit = vertical ? Math.max(99, sheet.rowCount - 1) : Math.max(25, sheet.columnCount - 1);
      const position = vertical ? row : col; let nextPosition = position;
      const populated = (at: number) => sheet.cells[addressOf(vertical ? at : row, vertical ? col : at)]?.raw;
      if (populated(position)) { while (nextPosition + step >= 0 && nextPosition + step <= limit && populated(nextPosition + step)) nextPosition += step; }
      else { let candidate = nextPosition + step; while (candidate >= 0 && candidate <= limit && !populated(candidate)) candidate += step; nextPosition = candidate >= 0 && candidate <= limit ? candidate : step > 0 ? limit : 0; }
      const nr = vertical ? nextPosition : row, nc = vertical ? col : nextPosition; updateDimensions(nr, nc);
      if (event.shiftKey) setRange({ startRow: bounds.startRow, startCol: bounds.startCol, endRow: nr, endCol: nc }, activeCell);
      else onSelect(addressOf(nr, nc), { startRow: nr, startCol: nc, endRow: nr, endCol: nc });
      focusCell(nr, nc); return;
    }
    if ((event.key === "Delete" || event.key === "Backspace") && (row !== active.row || col !== active.col || selection)) { event.preventDefault(); clearSelection(); return; }
    if (event.key === "Escape") { setCopyRange(null); if (isEditing) setCell(address, editOriginal.current); setEditingAddress(null); setInteractionState("READY"); event.currentTarget.blur(); return; }
    let nextRow = row, nextCol = col;
    if (event.key === "ArrowUp") nextRow = Math.max(0, row - 1);
    else if (event.key === "ArrowDown" || (event.key === "Enter" && !event.shiftKey)) nextRow = row + 1;
    else if (event.key === "Enter" && event.shiftKey) nextRow = Math.max(0, row - 1);
    else if (event.key === "ArrowLeft" && event.currentTarget.selectionStart === 0) nextCol = Math.max(0, col - 1);
    else if (event.key === "ArrowRight" && event.currentTarget.selectionEnd === event.currentTarget.value.length) nextCol = col + 1;
    else if (event.key === "Tab") nextCol = Math.max(0, col + (event.shiftKey ? -1 : 1));
    else if (event.key === "Home" && mod) { nextRow = 0; nextCol = 0; }
    else if (event.key === "Home") nextCol = 0;
    else if (event.key === "PageDown") nextRow += 20;
    else if (event.key === "PageUp") nextRow = Math.max(0, nextRow - 20);
    else return;
    event.preventDefault(); updateDimensions(nextRow, nextCol);
    if (event.shiftKey) setRange({ startRow: bounds.startRow, startCol: bounds.startCol, endRow: nextRow, endCol: nextCol }, activeCell);
    else onSelect(addressOf(nextRow, nextCol), { startRow: nextRow, startCol: nextCol, endRow: nextRow, endCol: nextCol });
    focusCell(nextRow, nextCol);
  };
  const contextTable = menu ? getTableAt(sheet.tables, menu.index, menu.col) : undefined;
  const menuItem = (label: string, action: string, shortcut?: string) => <button key={action} type="button" role="menuitem" onClick={() => menuAction(action)} className="grid-context-item"><span>{label}</span>{shortcut && <span className="grid-context-shortcut">{shortcut}</span>}</button>;
  const filteringTable = tableFilterMenu ? sheet.tables?.find(item => item.id === tableFilterMenu.tableId) : undefined;
  const filterValues = filteringTable && tableFilterMenu ? tableColumnValues(filteringTable, tableFilterMenu.col, sheet.cells) : [];
  const numericFilter = filterValues.filter(value => value !== "").every(value => Number.isFinite(Number(value)));
  const filterOperators: { value: TableFilterOperator; label: string }[] = numericFilter
    ? [{ value: "equals", label: "Equals" }, { value: "notEquals", label: "Does not equal" }, { value: "greaterThan", label: "Greater than" }, { value: "lessThan", label: "Less than" }, { value: "between", label: "Between" }]
    : [{ value: "contains", label: "Contains" }, { value: "notContains", label: "Does not contain" }, { value: "equals", label: "Equals" }, { value: "notEquals", label: "Does not equal" }, { value: "beginsWith", label: "Begins with" }, { value: "endsWith", label: "Ends with" }];
  const applyFormatDialog = () => {
    const range = selection ? clampRange(selection) : { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
    const addresses = Array.from({ length: range.endRow - range.startRow + 1 }, (_, r) => Array.from({ length: range.endCol - range.startCol + 1 }, (_, c) => addressOf(range.startRow + r, range.startCol + c))).flat();
    useWorkbookStore.getState().setStyle(formatDraft, addresses);
    setFormatDialogOpen(false);
  };
  const formatDialog = formatDialogOpen && <div className="grid-dialog-backdrop" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) setFormatDialogOpen(false); }}><section className="grid-dialog" role="dialog" aria-modal="true" aria-labelledby="grid-format-title"><header><div><span>FORMAT CELLS</span><h2 id="grid-format-title">Format Cells</h2></div><button type="button" onClick={() => setFormatDialogOpen(false)} aria-label="Close Format Cells">×</button></header><div className="grid-format-fields"><label>Number format<select value={formatDraft.numberFormat ?? "General"} onChange={event => setFormatDraft(current => ({ ...current, numberFormat: event.target.value as CellStyle["numberFormat"] }))}><option>General</option><option>Number</option><option>Currency ₹</option><option>Currency $</option><option>Percentage %</option><option>Date</option></select></label><label>Font family<select value={formatDraft.fontFamily ?? "Inter"} onChange={event => setFormatDraft(current => ({ ...current, fontFamily: event.target.value }))}><option>Inter</option><option>Arial</option><option>Calibri</option><option>Georgia</option><option>Courier New</option></select></label><label>Font size<input type="number" min="8" max="72" value={formatDraft.fontSize ?? 11} onChange={event => setFormatDraft(current => ({ ...current, fontSize: Math.max(8, Math.min(72, Number(event.target.value) || 11)) }))}/></label><label>Alignment<select value={formatDraft.align ?? "left"} onChange={event => setFormatDraft(current => ({ ...current, align: event.target.value as CellStyle["align"] }))}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select></label><label>Font color<input type="color" value={formatDraft.color ?? "#26382d"} onChange={event => setFormatDraft(current => ({ ...current, color: event.target.value }))}/></label><label>Fill color<input type="color" value={formatDraft.backgroundColor ?? "#ffffff"} onChange={event => setFormatDraft(current => ({ ...current, backgroundColor: event.target.value }))}/></label><label>Decimal places<input type="number" min="0" max="8" value={formatDraft.decimalPlaces ?? 2} onChange={event => setFormatDraft(current => ({ ...current, decimalPlaces: Math.max(0, Math.min(8, Number(event.target.value) || 0)) }))}/></label><div className="grid-format-toggles"><label><input type="checkbox" checked={!!formatDraft.bold} onChange={event => setFormatDraft(current => ({ ...current, bold: event.target.checked }))}/> Bold</label><label><input type="checkbox" checked={!!formatDraft.italic} onChange={event => setFormatDraft(current => ({ ...current, italic: event.target.checked }))}/> Italic</label><label><input type="checkbox" checked={!!formatDraft.underline} onChange={event => setFormatDraft(current => ({ ...current, underline: event.target.checked }))}/> Underline</label><label><input type="checkbox" checked={!!formatDraft.strike} onChange={event => setFormatDraft(current => ({ ...current, strike: event.target.checked }))}/> Strikethrough</label><label><input type="checkbox" checked={!!formatDraft.wrap} onChange={event => setFormatDraft(current => ({ ...current, wrap: event.target.checked }))}/> Wrap text</label></div></div><footer><button type="button" className="grid-dialog-secondary" onClick={() => setFormatDialogOpen(false)}>Cancel</button><button type="button" onClick={applyFormatDialog}>Apply to selection</button></footer></section></div>;
  const shortcutsDialog = shortcutsOpen && <div className="grid-dialog-backdrop" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) setShortcutsOpen(false); }}><section className="grid-dialog grid-shortcuts-dialog" role="dialog" aria-modal="true" aria-labelledby="grid-shortcuts-title"><header><div><span>GRIDNOVA</span><h2 id="grid-shortcuts-title">Keyboard shortcuts</h2></div><button type="button" onClick={() => setShortcutsOpen(false)} aria-label="Close shortcuts">×</button></header><div className="grid-shortcut-grid">{[["F2","Edit selected cell"],["Shift + F10 / Menu","Open context menu"],["Ctrl + C / X / V","Copy / Cut / Paste"],["Ctrl + Shift + V","Paste values"],["Ctrl + Z / Y","Undo / Redo"],["Ctrl + 1","Format Cells"],["Ctrl + Arrow","Jump across data"],["Ctrl + Home / End","First / last used cell"],["Shift + Space","Select entire row"],["Ctrl + Space","Select entire column"],["Ctrl + D / R","Fill down / right"],["Ctrl + ;","Insert current date"],["Tab / Shift + Tab","Move right / left"],["Enter / Shift + Enter","Move down / up"]].map(([key,label]) => <div key={key}><kbd>{key}</kbd><span>{label}</span></div>)}</div><footer><button type="button" onClick={() => setShortcutsOpen(false)}>Close</button></footer></section></div>;

  useEffect(() => {
    if (interactionState !== "SELECTING" && interactionState !== "FILLING") return;
    const onMove = (event: globalThis.PointerEvent) => interactionHandlersRef.current.update(event.clientX, event.clientY);
    const onUp = () => interactionHandlersRef.current.finish();
    window.addEventListener("pointermove", onMove); window.addEventListener("pointerup", onUp);
    let frame = 0;
    const autoScroll = () => {
      const viewport = gridRef.current;
      if (viewport) {
        const rect = viewport.getBoundingClientRect(), point = pointerPosition.current, edge = 34;
        const velocity = (value: number, low: number, high: number) => value < low + edge ? -Math.min(22, (low + edge - value) * 0.55) : value > high - edge ? Math.min(22, (value - (high - edge)) * 0.55) : 0;
        const dx = velocity(point.x, rect.left, rect.right), dy = velocity(point.y, rect.top, rect.bottom);
        if (dx || dy) { viewport.scrollLeft += dx; viewport.scrollTop += dy; interactionHandlersRef.current.update(point.x, point.y); }
      }
      frame = requestAnimationFrame(autoScroll);
    };
    frame = requestAnimationFrame(autoScroll);
    return () => { window.removeEventListener("pointermove", onMove); window.removeEventListener("pointerup", onUp); cancelAnimationFrame(frame); };
  }, [interactionState]);

  return <div ref={gridRef} className={`sheet-scroll relative ${interactionState === "EDITING" ? "grid-editing" : "grid-not-editing select-none"}`} onContextMenu={event => showContextMenu(event, "cell", active.row, active.col)} onClick={() => menu && setMenu(null)}>
    <div ref={canvasRef} className="grid-canvas" style={{ zoom: zoom / 100, minWidth: "max-content" }}>
      <table className="border-collapse table-fixed" style={{ minWidth: "100%" }}>
        <colgroup>{sheet.showHeadings && <col style={{ width: 48 }} />}{columns.map(col => <col key={col} style={{ width: resizePreview?.axis === "column" && resizePreview.index === col ? resizePreview.value : sheet.columnWidths[col] ?? 110 }} />)}</colgroup>
        {sheet.showHeadings && <thead className="sticky top-0 z-20"><tr>
          <th className="grid-corner sticky left-0 z-20" title="Select all cells"><button aria-label="Select all cells" onClick={() => { setRange({ startRow: 0, startCol: 0, endRow: Math.max(99, sheet.rowCount - 1), endCol: Math.max(25, sheet.columnCount - 1) }, "A1"); focusCell(0, 0); }}>◢</button></th>
          {columns.map(col => <th key={col} className={`grid-column-heading ${col >= bounds.startCol && col <= bounds.endCol ? "is-selected" : ""}`} onClick={event => { const first = event.shiftKey ? active.col : col; setRange({ startRow: 0, endRow: Math.max(99, sheet.rowCount - 1), startCol: Math.min(first, col), endCol: Math.max(first, col) }, addressOf(0, col)); focusCell(0, col); }} onContextMenu={event => showContextMenu(event, "column", col)}>
            {sheet.filters ? <div className="grid-filter-heading"><span>{columnLabel(col)}</span><select aria-label={`Filter ${columnLabel(col)}`} value={columnFilters[col] ?? ""} onChange={event => setColumnFilters(value => ({ ...value, [col]: event.target.value }))}><option value="">All</option>{Array.from(new Set(rows.map(row => sheet.cells[addressOf(row, col)]?.raw ?? "").filter(Boolean))).sort().map(value => <option key={value} value={value}>{value}</option>)}</select><Filter size={10}/></div> : columnLabel(col)}
            <span className="grid-column-resize" role="separator" aria-label={`Resize column ${columnLabel(col)}`} onMouseDown={event => beginResize(event, "column", col)} onDoubleClick={() => autoFitColumn(col)} />
          </th>)}
        </tr></thead>}
        <tbody>{rows.map(row => <tr key={row}>
          {sheet.showHeadings && <th className={`grid-row-heading ${row >= bounds.startRow && row <= bounds.endRow ? "is-selected" : ""}`} style={{ height: resizePreview?.axis === "row" && resizePreview.index === row ? resizePreview.value : rowHeights[row] ?? 27 }} onClick={event => { const first = event.shiftKey ? active.row : row; setRange({ startRow: Math.min(first, row), endRow: Math.max(first, row), startCol: 0, endCol: Math.max(25, sheet.columnCount - 1) }, addressOf(row, 0)); focusCell(row, 0); }} onContextMenu={event => showContextMenu(event, "row", row)}>{row + 1}<span className="grid-row-resize" role="separator" aria-label={`Resize row ${row + 1}`} onMouseDown={event => beginResize(event, "row", row)} onDoubleClick={() => autoFitRow(row)} /></th>}
          {columns.map(col => {
            const address = addressOf(row, col), cell = sheet.cells[address], display = formatCellValue(cell?.raw ?? "", sheet.cells, showFormulas, cell?.style, sheet.tables ?? [], address, namedRanges), value = editingAddress === address ? (cell?.raw ?? "") : display;
            const isSelected = selected(row, col), isActive = address === activeCell, style = cell?.style ?? {};
            const table = getTableAt(sheet.tables, row, col), palette = table ? getTablePalette(table.styleId) : null;
            const tableHeader = !!table && table.hasHeaders && row === table.range.startRow;
            const tableTotal = !!table && table.showTotalRow && row === table.range.endRow + 1;
            const tableFirstColumn = !!table && table.firstColumn && col === table.range.startCol;
            const tableLastColumn = !!table && table.lastColumn && col === table.range.endCol;
            const tableBandRow = !!table && table.bandedRows && !tableHeader && !tableTotal && (row - table.range.startRow) % 2 === 0;
            const tableBandColumn = !!table && table.bandedColumns && (col - table.range.startCol) % 2 === 1;
            const tableFill = tableHeader ? palette?.header : tableTotal ? palette?.band : tableBandRow || tableBandColumn ? palette?.band : undefined;
            if (cell?.mergedInto) return null;
            const gridBorder = sheet.showGridlines ? "1px solid #E2E5E9" : "none";
            const edge = (side: "Top" | "Right" | "Bottom" | "Left") => {
              const explicit = style[`border${side}` as "borderTop" | "borderRight" | "borderBottom" | "borderLeft"];
              if (explicit) return explicit;
              if (style.border === "none") return "none";
              if (style.border === "all") return `1px solid ${style.borderColor ?? "#64748b"}`;
              if (style.border === "outer") {
                const isOuter = side === "Top" ? row === bounds.startRow : side === "Bottom" ? row === bounds.endRow : side === "Left" ? col === bounds.startCol : col === bounds.endCol;
                return isOuter ? `1px solid ${style.borderColor ?? "#64748b"}` : gridBorder;
              }
              return gridBorder;
            };
            return <td key={col} data-cell={address} colSpan={cell?.mergeCols ?? 1} rowSpan={cell?.mergeRows ?? 1} className={`grid-cell ${sheet.freezeTopRow && row === 0 ? "sticky top-7 z-10" : ""}`} style={{ height: resizePreview?.axis === "row" && resizePreview.index === row ? resizePreview.value : rowHeights[row] ?? 27, borderTop: edge("Top"), borderRight: edge("Right"), borderBottom: edge("Bottom"), borderLeft: edge("Left"), background: style.backgroundColor ?? (tableFill ? tableFill : isActive ? "#ffffff" : isSelected ? "rgba(16, 124, 65, 0.08)" : "#ffffff"), fontFamily: style.fontFamily ?? "Inter", fontSize: style.fontSize ?? 11, fontWeight: style.bold || tableHeader || tableTotal || tableFirstColumn || tableLastColumn ? 700 : 400, fontStyle: style.italic ? "italic" : "normal", color: style.color ?? (tableHeader ? palette?.text ?? "#ffffff" : "#26382d"), textAlign: style.align ?? "left", verticalAlign: style.verticalAlign ?? "middle", textDecoration: [style.doubleUnderline ? "underline double" : style.underline ? "underline" : "", style.strike ? "line-through" : ""].filter(Boolean).join(" "), whiteSpace: style.wrap ? "normal" : "nowrap", overflow: "visible" }} onContextMenu={event => showContextMenu(event, "cell", row, col)}>
              {cell?.checkbox ? <div className="grid h-full place-items-center"><input type="checkbox" checked={cell.raw === "TRUE"} onChange={event => setCell(address, event.target.checked ? "TRUE" : "FALSE")} aria-label={`Checkbox ${address}`} className="h-4 w-4 accent-[#107c41]" /></div> : cell?.hyperlink ? <a href={cell.hyperlink} target="_blank" rel="noreferrer" className="flex h-full items-center gap-1 px-2 text-[#2778cb] underline"><span className="truncate">{value || cell.hyperlink}</span><ExternalLink size={10} /></a> : <input data-cell={address} aria-label={`Cell ${address}`} readOnly={editingAddress !== address && !cell?.datePicker} onFocus={() => { editOriginal.current = cell?.raw ?? ""; }} onDoubleClick={(event) => { event.preventDefault(); editOriginal.current = cell?.raw ?? ""; setInteractionState("EDITING"); setEditingAddress(address); }} onBlur={() => { setEditingAddress(current => current === address ? null : current); if (!formulaBarFocused) setInteractionState("READY"); }} type={cell?.datePicker ? "date" : "text"} value={cell?.datePicker ? (cell.raw || "") : value} onClick={event => { if (event.shiftKey) { const anchor = parseAddress(activeCell) ?? { row, col }; setRange({ startRow: anchor.row, startCol: anchor.col, endRow: row, endCol: col }, activeCell); } else onSelect(address, { startRow: row, startCol: col, endRow: row, endCol: col }); focusCell(row, col); }} onPointerDown={event => beginRange(row, col, event)} onChange={event => setCell(address, event.target.value)} onKeyDown={event => onKeyDown(event, row, col)} className="cell-input" style={{ textAlign: style.align ?? "left", fontFamily: style.fontFamily ?? "Inter", fontSize: style.fontSize ?? 11, fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? "italic" : "normal", color: style.color ?? "#26382d", textDecoration: style.doubleUnderline ? "underline double" : style.underline ? "underline" : "none", whiteSpace: style.wrap ? "normal" : "nowrap", writingMode: style.textOrientation === "vertical" ? "vertical-rl" : "horizontal-tb", paddingLeft: 6 + Math.max(0,Math.min(12,style.indent??0))*10, paddingRight: tableHeader && table?.showFilterButtons ? 26 : 6, caretColor: editingAddress === address ? "auto" : "transparent" }} title={cell?.comment ?? ""} />}
              {tableHeader && table?.showFilterButtons && <button type="button" className={table.filters?.[col] ? "table-header-filter is-filtered" : "table-header-filter"} aria-label={`Filter ${table.name} ${columnLabel(col)}`} title={`Filter ${table.name} — ${columnLabel(col)}`} onMouseDown={event => event.preventDefault()} onClick={event => openTableFilter(event, table.id, col)}>▾</button>}
              {cell?.comment && <span className="comment-mark" title={cell.comment} />}
            </td>;
          })}
        </tr>)}</tbody>
      </table>
      <div className="grid-selection-overlay" style={overlayRects.selection ? { left: overlayRects.selection.left, top: overlayRects.selection.top, width: overlayRects.selection.width, height: overlayRects.selection.height } : { display: "none" }} aria-hidden="true" />
      {overlayRects.copy && <div className="grid-range-overlay grid-copy-overlay" style={overlayRects.copy} aria-hidden="true" />}
      {overlayRects.fill && <div className="grid-range-overlay grid-fill-preview-overlay" style={overlayRects.fill} aria-hidden="true" />}
      {overlayRects.selection && <button type="button" className="grid-fill-handle-overlay" aria-label="Drag to fill cells" title="Drag to fill" style={{ left: overlayRects.selection.left + overlayRects.selection.width - 3, top: overlayRects.selection.top + overlayRects.selection.height - 3 }} onPointerDown={beginFillDrag} onDoubleClick={() => {
        const occupied = rows.filter(candidate => candidate > bounds.endRow && sheet.cells[addressOf(candidate, bounds.endCol + 1)]?.raw);
        const target = { ...bounds, endRow: Math.max(bounds.endRow, occupied.at(-1) ?? bounds.endRow) };
        fillSource.current = bounds; fillAxis.current = "vertical"; fillModifier.current = false; finishFill(target, "vertical");
      }} />}
      {fillDrag && fillDrag.axis && <div className="grid-fill-tooltip" style={{ left: fillDrag.clientX + 12, top: fillDrag.clientY + 12 }}>{fillDrag.previewValue || " "}</div>}
      <button className="grid-add-rows" onClick={() => updateSheet(sheet.id, { rowCount: sheet.rowCount + 100 })}><Plus size={13} /> Add 100 more rows at bottom</button>
    </div>
    {tableFilterMenu && filteringTable && createPortal(<div className="table-filter-scrim" role="presentation" onMouseDown={() => setTableFilterMenu(null)}>
      <section className="table-filter-popover" role="dialog" aria-label={`Filter ${columnLabel(tableFilterMenu.col)}`} style={{ left: tableFilterMenu.x, top: tableFilterMenu.y }} onMouseDown={event => event.stopPropagation()}>
        <div className="table-filter-title"><strong>{columnLabel(tableFilterMenu.col)} — Filter</strong><button type="button" aria-label="Close filter" onClick={() => setTableFilterMenu(null)}>×</button></div>
        <div className="table-filter-sort"><button type="button" onClick={() => sortTableColumn(filteringTable.id, tableFilterMenu.col, false)}>Sort A to Z</button><button type="button" onClick={() => sortTableColumn(filteringTable.id, tableFilterMenu.col, true)}>Sort Z to A</button></div>
        <label className="table-filter-search"><Search size={13}/><input value={tableFilterSearch} onChange={event => setTableFilterSearch(event.target.value)} placeholder="Search values"/></label>
        <div className="table-filter-value-actions"><button type="button" onClick={() => setTableFilterValues(filterValues)}>Select all</button><button type="button" onClick={() => setTableFilterValues([])}>Clear</button></div>
        <div className="table-filter-values">
          {filterValues.filter(value => (value || "(Blanks)").toLowerCase().includes(tableFilterSearch.toLowerCase())).map(value => <label key={value || "__blank"}><input type="checkbox" checked={tableFilterValues.includes(value)} onChange={event => setTableFilterValues(current => event.target.checked ? [...current, value] : current.filter(item => item !== value))}/><span>{value || "(Blanks)"}</span></label>)}
        </div>
        <div className="table-filter-condition">
          <label>Filter condition<select value={tableFilterOperator} onChange={event => setTableFilterOperator(event.target.value as TableFilterOperator)}>{filterOperators.map(operator => <option key={operator.value} value={operator.value}>{operator.label}</option>)}</select></label>
          <input value={tableFilterConditionValue} onChange={event => setTableFilterConditionValue(event.target.value)} placeholder={numericFilter ? "Enter a number" : "Enter text"}/>
          {tableFilterOperator === "between" && <input value={tableFilterConditionSecond} onChange={event => setTableFilterConditionSecond(event.target.value)} placeholder="And"/>}
        </div>
        <div className="table-filter-footer"><button type="button" onClick={() => { const filters = { ...(filteringTable.filters ?? {}) }; delete filters[tableFilterMenu.col]; updateSheet(sheet.id, { tables: (sheet.tables ?? []).map(item => item.id === filteringTable.id ? { ...item, filters } : item) }); setTableFilterMenu(null); }}>Clear Filter</button><button type="button" className="table-filter-apply" onClick={applyTableFilter}>Apply</button></div>
      </section>
    </div>, document.body)}
    {menu && createPortal(<div ref={menuRef} role="menu" aria-label="Spreadsheet context menu" className="grid-context-menu" style={{ left: menu.x, top: menu.y }} onClick={event => event.stopPropagation()} onKeyDown={event => {
      const items = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>("[role=menuitem]")), index = items.indexOf(document.activeElement as HTMLButtonElement);
      if (event.key === "Escape") { event.preventDefault(); setMenu(null); focusCell(active.row, active.col); }
      else if (event.key === "ArrowDown" || event.key === "ArrowUp") { event.preventDefault(); items[(index + (event.key === "ArrowDown" ? 1 : -1) + items.length) % items.length]?.focus(); }
      else if (event.key === "Home" || event.key === "End") { event.preventDefault(); (event.key === "Home" ? items[0] : items.at(-1))?.focus(); }
    }}>
      {menuItem("Cut", "cut", "Ctrl+X")}{menuItem("Copy", "copy", "Ctrl+C")}{menuItem("Paste", "paste", "Ctrl+V")}
      <div className="grid-context-divider" />
      {menu.kind !== "column" && <>{menuItem(`Insert ${bounds.endRow - bounds.startRow + 1} row${bounds.endRow === bounds.startRow ? "" : "s"} above`, "insert-row-above")}{menuItem(`Insert ${bounds.endRow - bounds.startRow + 1} row${bounds.endRow === bounds.startRow ? "" : "s"} below`, "insert-row-below")}</>}
      {menu.kind !== "row" && <>{menuItem(`Insert ${bounds.endCol - bounds.startCol + 1} column${bounds.endCol === bounds.startCol ? "" : "s"} left`, "insert-col-left")}{menuItem(`Insert ${bounds.endCol - bounds.startCol + 1} column${bounds.endCol === bounds.startCol ? "" : "s"} right`, "insert-col-right")}</>}
      {menu.kind !== "column" && menuItem("Delete row", "delete-row")}{menu.kind !== "row" && menuItem("Delete column", "delete-col")}
      <div className="grid-context-divider" />{menuItem("Clear contents", "clear", "Delete")}{menuItem("Clear formatting", "clear-format", "Ctrl+\\")}
      {menu.kind === "cell" && <>
        {menuItem("Insert cells down", "insert-cells-down")}{menuItem("Insert cells right", "insert-cells-right")}{menuItem("Delete cells up", "delete-cells-up")}{menuItem("Delete cells left", "delete-cells-left")}
        {menuItem("Format Cells…", "format-cells", "Ctrl+1")}
        <div className="grid-context-divider" />
        {menuItem("Paste Values Only", "paste-values")}{menuItem("Paste Formulas", "paste-formulas")}{menuItem("Paste Formatting", "paste-format")}{menuItem("Paste Transposed", "paste-transpose")}
        <div className="grid-context-divider" />
        {menuItem("Sort range A → Z", "sort-asc")}{menuItem("Sort range Z → A", "sort-desc")}{menuItem("Fill Down", "fill-down", "Ctrl+D")}{menuItem("Fill Right", "fill-right", "Ctrl+R")}
        {menuItem("Merge cells", "merge")}{menuItem("Unmerge cells", "unmerge")}
        <div className="grid-context-divider" />
        {menuItem("All borders", "border-all")}{menuItem("Outside borders", "border-outer")}{menuItem("Bottom border", "border-bottom")}{menuItem("Right border", "border-right")}{menuItem("Remove borders", "border-clear")}
        {menuItem("Keyboard shortcuts…", "show-shortcuts")}{menuItem("Clear All", "clear-all")}
        {contextTable && <>
          <div className="grid-context-divider" />
          {menuItem("Insert Table Row Above", "table-insert-row-above")}
          {menuItem("Insert Table Row Below", "table-insert-row-below")}
          {menuItem("Delete Table Row", "table-delete-row")}
          {menuItem("Insert Table Column Left", "table-insert-column-left")}
          {menuItem("Insert Table Column Right", "table-insert-column-right")}
          {menuItem("Delete Table Column", "table-delete-column")}
          {menuItem("Clear Table Filters", "table-clear-filters")}
          {menuItem("Remove Duplicates from Table", "table-remove-duplicates")}
          {menuItem("Convert Table to Range", "table-convert-range")}
        </>}
      </>}
      {menu.kind === "row" && <>{menuItem("Hide row", "hide-row")}{menuItem("Unhide all rows", "unhide-rows")}{menuItem("Resize row…", "resize-row")}{menuItem("Fit row to data", "autofit-row")}</>}
      {menu.kind === "column" && <>{menuItem("Hide column", "hide-column")}{menuItem("Unhide all columns", "unhide-columns")}{menuItem("Resize column…", "resize-column")}{menuItem("Fit column to data", "autofit-column")}</>}
    </div>, document.body)}
    {formatDialog}
    {shortcutsDialog}
  </div>;
}
