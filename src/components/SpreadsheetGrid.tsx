import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { KeyboardEvent, MouseEvent as ReactMouseEvent } from "react";
import { ExternalLink, Filter, Plus } from "lucide-react";
import type { CellRange, SheetData } from "../types";
import { addressOf, columnLabel, parseAddress } from "../types";
import { formatCellValue } from "../utils/formulaEngine";
import { useWorkbookStore } from "../store/useWorkbookStore";

interface Props {
  sheet: SheetData; activeCell: string; selection: CellRange | null; showFormulas: boolean; zoom: number;
  onSelect: (address: string, range?: CellRange) => void;
}
type ContextMenuState = { x: number; y: number; kind: "cell" | "row" | "column"; index: number; col: number } | null;
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

export default function SpreadsheetGrid({ sheet, activeCell, selection, showFormulas, zoom, onSelect }: Props) {
  const setCell = useWorkbookStore(s => s.setCell);
  const updateSheet = useWorkbookStore(s => s.updateSheet);
  const setClipboard = useWorkbookStore(s => s.setClipboard);
  const clipboard = useWorkbookStore(s => s.clipboard);
  const undo = useWorkbookStore(s => s.undo);
  const redo = useWorkbookStore(s => s.redo);
  const [dragStart, setDragStart] = useState<{ row: number; col: number } | null>(null);
  const [dragging, setDragging] = useState(false);
  const [columnFilters, setColumnFilters] = useState<Record<number, string>>({});
  const [editingAddress, setEditingAddress] = useState<string | null>(null);
  const [menu, setMenu] = useState<ContextMenuState>(null);
  const rowHeights = sheet.rowHeights ?? {};
  const [fillPreview, setFillPreview] = useState<CellRange | null>(null);
  const [copyRange, setCopyRange] = useState<CellRange | null>(null);
  const [resizePreview, setResizePreview] = useState<{ axis: "row" | "column"; index: number; value: number } | null>(null);
  const resizeValue = useRef<{ axis: "row" | "column"; index: number; value: number } | null>(null);
  const fillSource = useRef<CellRange | null>(null);
  const fillModifier = useRef(false);
  const selectionGesture = useRef(false);
  const suppressHeaderClick = useRef(false);
  const finishFillRef = useRef<(preview?: CellRange) => void>(() => {});
  const gridRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const editOriginal = useRef("");
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
  const columns = useMemo(() => Array.from({ length: Math.max(sheet.columnCount, 26) }, (_, i) => i), [sheet.columnCount]);
  const rows = useMemo(() => Array.from({ length: Math.max(sheet.rowCount, 100) }, (_, i) => i).filter(row => row === 0 || Object.entries(columnFilters).every(([col, value]) => !value || (sheet.cells[addressOf(row, Number(col))]?.raw ?? "") === value)), [sheet.rowCount, sheet.cells, columnFilters]);
  const bounds = selection ? clampRange(selection) : { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
  // While filling, header indicators follow the live preview rather than flashing
  // between the original source range and the destination range.
  const headerBounds = fillSource.current && fillPreview ? clampRange(fillPreview) : bounds;
  const isFullRowSelection = headerBounds.startCol === 0 && headerBounds.endCol >= columns.length - 1;
  const isFullColumnSelection = headerBounds.startRow === 0 && headerBounds.endRow >= Math.max(99, sheet.rowCount - 1);
  const isEntireSheetSelection = isFullRowSelection && isFullColumnSelection;
  const showColumnHeaderSelection = !isFullRowSelection || isEntireSheetSelection;
  const showRowHeaderSelection = !isFullColumnSelection || isEntireSheetSelection;
  const selected = (row: number, col: number) => row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol && col <= bounds.endCol;
  const setRange = (range: CellRange, address = addressOf(range.endRow, range.endCol)) => onSelect(address, range);
  const finishFill = (preview?: CellRange) => {
    const source = fillSource.current, target = preview ?? fillPreview;
    fillSource.current = null; fillModifier.current = false; setFillPreview(null);
    if (!source || !target) return;
    const original = clampRange(source), destination = clampRange(target), next = { ...sheet.cells };
    const fillDown = destination.endRow > original.endRow, fillUp = destination.startRow < original.startRow;
    const fillRight = destination.endCol > original.endCol, fillLeft = destination.startCol < original.startCol;
    const axis = fillDown || fillUp ? "row" : "column";
    const sourceLength = axis === "row" ? original.endRow - original.startRow + 1 : original.endCol - original.startCol + 1;
    const sequence = Array.from({ length: sourceLength }, (_, i) => {
      const row = axis === "row" ? (fillUp ? original.startRow + i : original.endRow - sourceLength + 1 + i) : original.startRow;
      const col = axis === "column" ? (fillLeft ? original.startCol + i : original.endCol - sourceLength + 1 + i) : original.startCol;
      return sheet.cells[addressOf(row, col)]?.raw ?? "";
    });
    const numbers = sequence.map(Number);
    let step = sourceLength > 1 && numbers.every(Number.isFinite) ? numbers[numbers.length - 1] - numbers[numbers.length - 2] : 1;
    const monthNames = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    const weekNames = ["sun","mon","tue","wed","thu","fri","sat"];
    const series = sourceLength > 1 && monthNames.includes(sequence.at(-1)?.toLowerCase().slice(0, 3) ?? "") ? monthNames : sourceLength > 1 && weekNames.includes(sequence.at(-1)?.toLowerCase().slice(0, 3) ?? "") ? weekNames : null;
    for (let row = destination.startRow; row <= destination.endRow; row++) for (let col = destination.startCol; col <= destination.endCol; col++) {
      if (row >= original.startRow && row <= original.endRow && col >= original.startCol && col <= original.endCol) continue;
      const rowDelta = row - (fillDown ? original.endRow : original.startRow), colDelta = col - (fillRight ? original.endCol : original.startCol);
      const sourceRow = Math.max(original.startRow, Math.min(original.endRow, row)), sourceCol = Math.max(original.startCol, Math.min(original.endCol, col));
      const template = sheet.cells[addressOf(sourceRow, sourceCol)] ?? sheet.cells[addressOf(original.startRow, original.startCol)] ?? { raw: "" };
      let raw = template.raw;
      if (raw.startsWith("=")) raw = shiftFormulaReferences(raw, rowDelta, colDelta);
      else if ((sourceLength > 1 || /\D+\d+$/.test(raw) || fillModifier.current && numbers.every(Number.isFinite) || /^\d{4}-\d{2}-\d{2}$/.test(raw)) && (fillDown || fillUp || fillRight || fillLeft)) {
        const offset = fillDown ? row - original.endRow - 1 : fillUp ? original.startRow - row - 1 : fillRight ? col - original.endCol - 1 : original.startCol - col - 1;
        const backwards = fillUp || fillLeft;
        if (numbers.every(Number.isFinite) && (sourceLength > 1 || fillModifier.current)) raw = String(Number(backwards ? sequence[0] : sequence.at(-1)) + step * (backwards ? -1 : 1) * (offset + 1));
        else if (series) { const base = backwards ? sequence[0] : sequence.at(-1)!; const last = series.indexOf(base.toLowerCase().slice(0, 3)); const item = series[((last + (backwards ? -1 : 1) * (offset + 1)) % series.length + series.length) % series.length]; raw = base.replace(/^([A-Za-z]{3})/i, item[0].toUpperCase() + item.slice(1)); }
        else if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) { const date = new Date(`${backwards ? sequence[0] : raw}T00:00:00Z`); date.setUTCDate(date.getUTCDate() + (backwards ? -1 : 1) * (offset + 1)); raw = date.toISOString().slice(0, 10); }
        else { const base = backwards ? sequence[0] : raw; const match = base.match(/^(.*?)(\d+)$/); if (match) raw = `${match[1]}${Number(match[2]) + (backwards ? -1 : 1) * (offset + 1)}`; }
      }
      next[addressOf(row, col)] = { ...template, raw };
    }
    updateSheet(sheet.id, { cells: next, rowCount: Math.max(sheet.rowCount, destination.endRow + 1), columnCount: Math.max(sheet.columnCount, destination.endCol + 1) });
    setRange(destination, addressOf(destination.endRow, destination.endCol));
  };
  useLayoutEffect(() => {
    finishFillRef.current = finishFill;
  }, [finishFill]);

  useEffect(() => {
    const releasePointerGesture = () => {
      if (!selectionGesture.current) return;
      selectionGesture.current = false;
      setDragging(false);
      setDragStart(null);
      finishFillRef.current();
      // Keep the guard through the synthetic click that follows pointerup/mouseup.
      window.requestAnimationFrame(() => {
        suppressHeaderClick.current = false;
      });
    };

    window.addEventListener("pointerup", releasePointerGesture, true);
    window.addEventListener("pointercancel", releasePointerGesture, true);
    window.addEventListener("mouseup", releasePointerGesture, true);
    window.addEventListener("blur", releasePointerGesture);
    return () => {
      window.removeEventListener("pointerup", releasePointerGesture, true);
      window.removeEventListener("pointercancel", releasePointerGesture, true);
      window.removeEventListener("mouseup", releasePointerGesture, true);
      window.removeEventListener("blur", releasePointerGesture);
    };
  }, []);

  const focusCell = (row: number, col: number) => requestAnimationFrame(() => gridRef.current?.querySelector<HTMLInputElement>(`[data-cell="${addressOf(row, col)}"]`)?.focus());
  const updateDimensions = (row: number, col: number) => updateSheet(sheet.id, { rowCount: Math.max(sheet.rowCount, row + 1), columnCount: Math.max(sheet.columnCount, col + 1) });
  const updateRowHeight = (row: number, height: number) => updateSheet(sheet.id, { rowHeights: { ...(sheet.rowHeights ?? {}), [row]: Math.max(18, Math.min(120, height)) } });

  const beginResize = (event: ReactMouseEvent, axis: "row" | "column", index: number) => {
    event.preventDefault(); event.stopPropagation();
    const origin = axis === "row" ? event.clientY : event.clientX;
    const measured = (event.currentTarget as HTMLElement).parentElement?.getBoundingClientRect();
    const initial = axis === "row" ? rowHeights[index] ?? measured?.height ?? 27 : sheet.columnWidths[index] ?? measured?.width ?? 110;
    const move = (e: globalThis.MouseEvent) => {
      const delta = (axis === "row" ? e.clientY : e.clientX) - origin;
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
    const width = rows.reduce((max, row) => Math.max(max, (sheet.cells[addressOf(row, col)]?.raw ?? "").split("\n").reduce((lineMax, line) => Math.max(lineMax, line.length * 7.2 + 18), 48)), 48);
    updateSheet(sheet.id, { columnWidths: { ...sheet.columnWidths, [col]: Math.min(480, Math.max(48, width)) } });
  };
  const autoFitRow = (row: number) => {
    const height = columns.reduce((max, col) => {
      const text = sheet.cells[addressOf(row, col)]?.raw ?? "", width = sheet.columnWidths[col] ?? 110;
      const lines = text.split("\n").reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length * 7.2 / width)), 0);
      return Math.max(max, lines * 16 + 8);
    }, 27);
    updateRowHeight(row, height);
  };
  const beginRange = (row: number, col: number, event: ReactMouseEvent) => {
    const anchor = event.shiftKey ? active : { row, col };
    selectionGesture.current = true;
    suppressHeaderClick.current = true;
    setDragStart(anchor); setDragging(true);
    setRange({ startRow: anchor.row, startCol: anchor.col, endRow: row, endCol: col }, addressOf(row, col));
  };
  const showContextMenu = (event: ReactMouseEvent, kind: NonNullable<ContextMenuState>["kind"], index: number, col = active.col) => {
    event.preventDefault(); event.stopPropagation();
    if (kind === "cell" && !selected(index, col)) onSelect(addressOf(index, col), { startRow: index, startCol: col, endRow: index, endCol: col });
    if (kind === "row" && !(index >= bounds.startRow && index <= bounds.endRow && bounds.startCol === 0 && bounds.endCol >= columns.length - 1)) onSelect(addressOf(index, 0), { startRow: index, startCol: 0, endRow: index, endCol: Math.max(25, sheet.columnCount - 1) });
    if (kind === "column" && !(index >= bounds.startCol && index <= bounds.endCol && bounds.startRow === 0 && bounds.endRow >= rows.length - 1)) onSelect(addressOf(0, index), { startRow: 0, startCol: index, endRow: Math.max(99, sheet.rowCount - 1), endCol: index });
    setMenu({ x: Math.max(4, Math.min(event.clientX, window.innerWidth - 244)), y: Math.max(4, Math.min(event.clientY, window.innerHeight - 400)), kind, index, col });
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
    updateSheet(sheet.id, { cells, rowCount: Math.max(100, sheet.rowCount + (insert ? count : -count)) });
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
    updateSheet(sheet.id, { cells, columnWidths: widths, columnCount: Math.max(26, sheet.columnCount + (insert ? count : -count)) });
  };
  const clearSelection = (format = false) => {
    const next = { ...sheet.cells };
    for (let row = bounds.startRow; row <= bounds.endRow; row++) for (let col = bounds.startCol; col <= bounds.endCol; col++) {
      const key = addressOf(row, col);
      if (format) next[key] = { ...(next[key] ?? { raw: "" }), style: {} };
      else next[key] = { ...(next[key] ?? { raw: "" }), raw: "" };
    }
    updateSheet(sheet.id, { cells: next });
  };
  const menuAction = (action: string) => {
    if (!menu) return;
    const countRows = bounds.endRow - bounds.startRow + 1;
    const countCols = bounds.endCol - bounds.startCol + 1;
    if (action === "cut") void copySystem(true);
    else if (action === "copy") void copySystem();
    else if (action === "paste") void pasteFromSystem();
    else if (action === "insert-row-above") shiftRows(menu.kind === "row" ? menu.index : bounds.startRow, countRows, true);
    else if (action === "insert-row-below") shiftRows((menu.kind === "row" ? menu.index : bounds.endRow) + 1, countRows, true);
    else if (action === "delete-row") shiftRows(menu.kind === "row" ? menu.index : bounds.startRow, menu.kind === "row" ? countRows : 1, false);
    else if (action === "insert-col-left") shiftColumns(menu.kind === "column" ? menu.index : bounds.startCol, countCols, true);
    else if (action === "insert-col-right") shiftColumns((menu.kind === "column" ? menu.index : bounds.endCol) + 1, countCols, true);
    else if (action === "delete-col") shiftColumns(menu.kind === "column" ? menu.index : bounds.startCol, menu.kind === "column" ? countCols : 1, false);
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
    setMenu(null); focusCell(active.row, active.col);
  };
  const onKeyDown = (event: KeyboardEvent<HTMLInputElement>, row: number, col: number) => {
    const mod = event.ctrlKey || event.metaKey;
    if (event.key === "F10" && event.shiftKey) { event.preventDefault(); setMenu({ x: Math.min(window.innerWidth - 244, 180), y: Math.min(window.innerHeight - 400, 190), kind: "cell", index: row, col }); return; }
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
      if (event.shiftKey) setRange({ startRow: bounds.startRow, startCol: bounds.startCol, endRow: nr, endCol: nc }, addressOf(nr, nc));
      else onSelect(addressOf(nr, nc), { startRow: nr, startCol: nc, endRow: nr, endCol: nc });
      focusCell(nr, nc); return;
    }
    if ((event.key === "Delete" || event.key === "Backspace") && (row !== active.row || col !== active.col || selection)) { event.preventDefault(); clearSelection(); return; }
    if (event.key === "Escape") { setCopyRange(null); setCell(addressOf(row, col), editOriginal.current); setEditingAddress(null); event.currentTarget.blur(); return; }
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
    if (event.shiftKey) setRange({ startRow: bounds.startRow, startCol: bounds.startCol, endRow: nextRow, endCol: nextCol }, addressOf(nextRow, nextCol));
    else onSelect(addressOf(nextRow, nextCol), { startRow: nextRow, startCol: nextCol, endRow: nextRow, endCol: nextCol });
    focusCell(nextRow, nextCol);
  };
  const menuItem = (label: string, action: string, shortcut?: string) => <button key={action} role="menuitem" onClick={() => menuAction(action)} className="grid-context-item"><span>{label}</span>{shortcut && <span className="grid-context-shortcut">{shortcut}</span>}</button>;

  return <div ref={gridRef} className="sheet-scroll relative" onContextMenu={event => showContextMenu(event, "cell", active.row)} onClick={() => menu && setMenu(null)}>
    <div style={{ zoom: zoom / 100, minWidth: "max-content" }}>
      <table className="border-separate border-spacing-0 table-fixed" style={{ minWidth: "100%" }}>
        <colgroup>{sheet.showHeadings && <col style={{ width: 48 }} />}{columns.map(col => <col key={col} style={{ width: resizePreview?.axis === "column" && resizePreview.index === col ? resizePreview.value : sheet.columnWidths[col] ?? 110 }} />)}</colgroup>
        {sheet.showHeadings && <thead className="sticky top-0 z-20"><tr>
          <th className="grid-corner sticky left-0 z-20" title="Select all cells"><button aria-label="Select all cells" onClick={() => { setRange({ startRow: 0, startCol: 0, endRow: Math.max(99, sheet.rowCount - 1), endCol: Math.max(25, sheet.columnCount - 1) }, "A1"); focusCell(0, 0); }}>◢</button></th>
          {columns.map(col => { const headerActive = col >= headerBounds.startCol && col <= headerBounds.endCol && showColumnHeaderSelection; return <th key={col} className="grid-column-heading" style={{ backgroundColor: headerActive ? "#E6E6E6" : "#F8F9FA", color: headerActive ? "#107C41" : "#444746", fontWeight: headerActive ? 600 : 400, borderTop: headerActive ? "1px solid #D4D4D4" : 0, borderRight: "1px solid #D4D4D4", borderBottom: "1px solid #D4D4D4", borderLeft: 0, boxShadow: headerActive ? "inset 0 -2px 0 0 #107C41" : "none" }} onClick={event => { if (suppressHeaderClick.current) { event.preventDefault(); event.stopPropagation(); return; } const first = event.shiftKey ? active.col : col; setRange({ startRow: 0, endRow: Math.max(99, sheet.rowCount - 1), startCol: Math.min(first, col), endCol: Math.max(first, col) }, addressOf(0, col)); focusCell(0, col); }} onContextMenu={event => showContextMenu(event, "column", col)}>
            {sheet.filters ? <div className="grid-filter-heading"><span>{columnLabel(col)}</span><select aria-label={`Filter ${columnLabel(col)}`} value={columnFilters[col] ?? ""} onChange={event => setColumnFilters(value => ({ ...value, [col]: event.target.value }))}><option value="">All</option>{Array.from(new Set(rows.map(row => sheet.cells[addressOf(row, col)]?.raw ?? "").filter(Boolean))).sort().map(value => <option key={value} value={value}>{value}</option>)}</select><Filter size={10}/></div> : columnLabel(col)}
            <span className="grid-column-resize" role="separator" aria-label={`Resize column ${columnLabel(col)}`} onMouseDown={event => beginResize(event, "column", col)} onDoubleClick={() => autoFitColumn(col)} />
          </th>; })}
        </tr></thead>}
        <tbody>{rows.map(row => <tr key={row}>
          {sheet.showHeadings && <th className="grid-row-heading" style={{ height: resizePreview?.axis === "row" && resizePreview.index === row ? resizePreview.value : rowHeights[row] ?? 27, backgroundColor: row >= headerBounds.startRow && row <= headerBounds.endRow && showRowHeaderSelection ? "#E6E6E6" : "#F8F9FA", color: row >= headerBounds.startRow && row <= headerBounds.endRow && showRowHeaderSelection ? "#107C41" : "#444746", fontWeight: row >= headerBounds.startRow && row <= headerBounds.endRow && showRowHeaderSelection ? 600 : 400, borderTop: 0, borderRight: "1px solid #D4D4D4", borderBottom: "1px solid #D4D4D4", borderLeft: 0, boxShadow: row >= headerBounds.startRow && row <= headerBounds.endRow && showRowHeaderSelection ? "inset -2px 0 0 0 #107C41" : "none" }} onClick={event => { if (suppressHeaderClick.current) { event.preventDefault(); event.stopPropagation(); return; } const first = event.shiftKey ? active.row : row; setRange({ startRow: Math.min(first, row), endRow: Math.max(first, row), startCol: 0, endCol: Math.max(25, sheet.columnCount - 1) }, addressOf(row, 0)); focusCell(row, 0); }} onContextMenu={event => showContextMenu(event, "row", row)}>{row + 1}<span className="grid-row-resize" role="separator" aria-label={`Resize row ${row + 1}`} onMouseDown={event => beginResize(event, "row", row)} onDoubleClick={() => autoFitRow(row)} /></th>}
          {columns.map(col => {
            const address = addressOf(row, col), cell = sheet.cells[address], display = formatCellValue(cell?.raw ?? "", sheet.cells, showFormulas, cell?.style), value = editingAddress === address ? (cell?.raw ?? "") : display;
            const isSelected = selected(row, col), isActive = address === activeCell, style = cell?.style ?? {};
            if (cell?.mergedInto) return null;
            return <td key={col} colSpan={cell?.mergeCols ?? 1} rowSpan={cell?.mergeRows ?? 1} className={`grid-cell ${isSelected ? "selected" : ""} ${isActive ? "active-cell" : ""} ${isSelected && row === bounds.startRow ? "range-top" : ""} ${isSelected && row === bounds.endRow ? "range-bottom" : ""} ${isSelected && col === bounds.startCol ? "range-left" : ""} ${isSelected && col === bounds.endCol ? "range-right" : ""} ${copyRange && row >= copyRange.startRow && row <= copyRange.endRow && col >= copyRange.startCol && col <= copyRange.endCol ? "copy-outline" : ""} ${fillPreview && row >= fillPreview.startRow && row <= fillPreview.endRow && col >= fillPreview.startCol && col <= fillPreview.endCol ? "fill-preview" : ""} ${sheet.freezeTopRow && row === 0 ? "sticky top-7 z-10" : ""}`} style={{ height: resizePreview?.axis === "row" && resizePreview.index === row ? resizePreview.value : rowHeights[row] ?? 27, borderColor: sheet.showGridlines ? "#e8ece9" : "transparent", borderWidth: style.border === "none" ? 0 : 1, borderStyle: "solid", background: style.backgroundColor ?? (isActive ? "#fff" : isSelected ? "rgba(16,124,65,.08)" : "#fff"), fontFamily: style.fontFamily ?? "Inter", fontSize: style.fontSize ?? 11, fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? "italic" : "normal", color: style.color ?? "#26382d", textAlign: style.align ?? "left", textDecoration: [style.underline ? "underline" : "", style.strike ? "line-through" : ""].filter(Boolean).join(" "), whiteSpace: style.wrap ? "normal" : "nowrap", overflow: "visible" }} onContextMenu={event => showContextMenu(event, "cell", row, col)}>
              {cell?.checkbox ? <div className="grid h-full place-items-center"><input type="checkbox" checked={cell.raw === "TRUE"} onChange={event => setCell(address, event.target.checked ? "TRUE" : "FALSE")} aria-label={`Checkbox ${address}`} className="h-4 w-4 accent-[#107c41]" /></div> : cell?.hyperlink ? <a href={cell.hyperlink} target="_blank" rel="noreferrer" className="flex h-full items-center gap-1 px-2 text-[#2778cb] underline"><span className="truncate">{value || cell.hyperlink}</span><ExternalLink size={10} /></a> : <input data-cell={address} aria-label={`Cell ${address}`} onFocus={() => { editOriginal.current = cell?.raw ?? ""; setEditingAddress(address); }} onBlur={() => setEditingAddress(current => current === address ? null : current)} type={cell?.datePicker ? "date" : "text"} value={cell?.datePicker ? (cell.raw || "") : value} onClick={event => { if (event.shiftKey) { const anchor = parseAddress(activeCell) ?? { row, col }; setRange({ startRow: anchor.row, startCol: anchor.col, endRow: row, endCol: col }, address); } else onSelect(address, { startRow: row, startCol: col, endRow: row, endCol: col }); }} onMouseDown={event => beginRange(row, col, event)} onMouseEnter={() => { if (fillSource.current) { const source = clampRange(fillSource.current); setFillPreview({ startRow: Math.min(source.startRow, row), startCol: Math.min(source.startCol, col), endRow: Math.max(source.endRow, row), endCol: Math.max(source.endCol, col) }); } else if (dragging && dragStart) setRange({ startRow: dragStart.row, startCol: dragStart.col, endRow: row, endCol: col }, address); }} onChange={event => setCell(address, event.target.value)} onKeyDown={event => onKeyDown(event, row, col)} className="cell-input" style={{ textAlign: style.align ?? "left", fontFamily: style.fontFamily ?? "Inter", fontSize: style.fontSize ?? 11, fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? "italic" : "normal", color: style.color ?? "#26382d", textDecoration: style.underline ? "underline" : "none", whiteSpace: style.wrap ? "normal" : "nowrap" }} title={cell?.comment ?? ""} />}
              {row === bounds.endRow && col === bounds.endCol && <button type="button" className="grid-fill-handle" aria-label="Drag to fill cells" title="Drag to fill" onMouseDown={event => { event.preventDefault(); event.stopPropagation(); selectionGesture.current = true; suppressHeaderClick.current = true; fillSource.current = bounds; fillModifier.current = event.ctrlKey || event.altKey; setFillPreview(null); }} onDoubleClick={() => { const occupied = rows.filter(candidate => candidate > row && sheet.cells[addressOf(candidate, col + 1)]?.raw); const last = occupied.at(-1) ?? row; const target = { ...bounds, endRow: Math.max(row, last) }; fillSource.current = bounds; setFillPreview(target); requestAnimationFrame(() => finishFill(target)); }} />}
              {cell?.comment && <span className="comment-mark" title={cell.comment} />}
            </td>;
          })}
        </tr>)}</tbody>
      </table>
      <button className="grid-add-rows" onClick={() => updateSheet(sheet.id, { rowCount: sheet.rowCount + 100 })}><Plus size={13} /> Add 100 more rows at bottom</button>
    </div>
    {menu && createPortal(<div ref={menuRef} role="menu" aria-label="Spreadsheet context menu" className="grid-context-menu" style={{ left: menu.x, top: menu.y }} onClick={event => event.stopPropagation()} onKeyDown={event => {
      const items = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>("[role=menuitem]")), index = items.indexOf(document.activeElement as HTMLButtonElement);
      if (event.key === "Escape") { event.preventDefault(); setMenu(null); focusCell(active.row, active.col); }
      else if (event.key === "ArrowDown" || event.key === "ArrowUp") { event.preventDefault(); items[(index + (event.key === "ArrowDown" ? 1 : -1) + items.length) % items.length]?.focus(); }
    }}>
      {menuItem("Cut", "cut", "Ctrl+X")}{menuItem("Copy", "copy", "Ctrl+C")}{menuItem("Paste", "paste", "Ctrl+V")}
      <div className="grid-context-divider" />
      {menu.kind !== "column" && <>{menuItem(`Insert ${bounds.endRow - bounds.startRow + 1} row${bounds.endRow === bounds.startRow ? "" : "s"} above`, "insert-row-above")}{menuItem(`Insert ${bounds.endRow - bounds.startRow + 1} row${bounds.endRow === bounds.startRow ? "" : "s"} below`, "insert-row-below")}</>}
      {menu.kind !== "row" && <>{menuItem(`Insert ${bounds.endCol - bounds.startCol + 1} column${bounds.endCol === bounds.startCol ? "" : "s"} left`, "insert-col-left")}{menuItem(`Insert ${bounds.endCol - bounds.startCol + 1} column${bounds.endCol === bounds.startCol ? "" : "s"} right`, "insert-col-right")}</>}
      {menu.kind !== "column" && menuItem(menu.kind === "row" ? "Delete row" : "Delete row", "delete-row")}{menu.kind !== "row" && menuItem("Delete column", "delete-col")}
      <div className="grid-context-divider" />{menuItem("Clear contents", "clear", "Delete")}{menuItem("Clear formatting", "clear-format", "Ctrl+\\")}
      {menu.kind === "cell" && <>{menuItem("Sort range A → Z", "sort-asc")}{menuItem("Sort range Z → A", "sort-desc")}{menuItem("Merge cells", "merge")}{menuItem("Unmerge cells", "unmerge")}</>}
      {menu.kind === "row" && <>{menuItem("Resize row…", "resize-row")}{menuItem("Fit row to data", "autofit-row")}</>}{menu.kind === "column" && <>{menuItem("Resize column…", "resize-column")}{menuItem("Fit column to data", "autofit-column")}</>}
    </div>, document.body)}
  </div>;
}
