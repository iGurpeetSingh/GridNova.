import type { CellMap, CellRange, CellStyle, SheetData } from "../types";
import { addressOf, parseAddress } from "../types";
import { evaluate, type FormulaMatrix, type FormulaScalar, type FormulaValue, type FormulaNamedRange } from "./formulaEngine";
import { DependencyGraph, nodeKey, type DependencyPlan } from "./dependencyGraph";

export interface SpillGhostCell {
  origin: string;
  sheetId: string;
  value: FormulaScalar;
}

export interface SpillOrigin {
  address: string;
  range: CellRange;
  matrix: FormulaMatrix;
}

export interface SheetCalculation {
  sheetId: string;
  values: Record<string, FormulaScalar>;
  spillMatrices: Record<string, FormulaMatrix>;
  spillOrigins: Record<string, SpillOrigin>;
  ghostCells: Record<string, SpillGhostCell>;
  errors: Record<string, string>;
  circular: string[];
  blocked: string[];
}

export interface WorkbookCalculation {
  sheets: Record<string, SheetCalculation>;
  graph: DependencyGraph;
  plan: DependencyPlan;
}

export interface CalculationOptions {
  namedRanges?: Record<string, string | FormulaNamedRange>;
  maxArrayCells?: number;
  maxDependencyRangeCells?: number;
}

const MAX_SPILL_CELLS = 100000;

function normalizeSheetLookup(sheets: SheetData[]): Record<string, string> {
  const lookup: Record<string, string> = {};
  for (const sheet of sheets) {
    lookup[sheet.name] = sheet.id;
    lookup[sheet.name.toLowerCase()] = sheet.id;
    lookup[sheet.id] = sheet.id;
  }
  return lookup;
}

function scalar(value: FormulaValue): FormulaScalar {
  return Array.isArray(value) ? (value[0]?.[0] ?? "") : value;
}

function isFormulaCell(raw: string | undefined): boolean {
  return typeof raw === "string" && raw.trim().startsWith("=");
}

function emptySheetCalculation(sheetId: string): SheetCalculation {
  return {
    sheetId,
    values: {},
    spillMatrices: {},
    spillOrigins: {},
    ghostCells: {},
    errors: {},
    circular: [],
    blocked: [],
  };
}

function spillBounds(origin: string, matrix: FormulaMatrix): CellRange | null {
  const point = parseAddress(origin);
  if (!point || !matrix.length || !matrix[0]?.length) return null;
  return {
    startRow: point.row,
    startCol: point.col,
    endRow: point.row + matrix.length - 1,
    endCol: point.col + matrix[0].length - 1,
  };
}

/**
 * Calculates formula cells in dependency order and materializes dynamic-array
 * spill results as virtual grid cells. The source CellMaps are never mutated.
 *
 * The dependency graph is rebuilt from workbook formulas so that edits to
 * formulas, cross-sheet references and cleared source cells cannot leave stale
 * edges behind. The graph exposes an incremental plan for callers that retain
 * it between edits; this pure workbook calculation currently creates a complete
 * snapshot to guarantee that spill collisions and removed spills are cleared.
 */
export function calculateWorkbook(
  sourceSheets: SheetData[],
  options: CalculationOptions = {},
): WorkbookCalculation {
  const sheets = sourceSheets ?? [];
  const cellMaps: Record<string, CellMap> = {};
  for (const sheet of sheets) cellMaps[sheet.id] = sheet.cells ?? {};

  const sheetNameToId = normalizeSheetLookup(sheets);
  const graph = DependencyGraph.fromSheets(
    sheets.map(sheet => ({ id: sheet.id, name: sheet.name, cells: sheet.cells })),
    {
      namedRanges: options.namedRanges,
      maxExpandedRangeCells: options.maxDependencyRangeCells,
    },
  );

  const formulaNodes: string[] = [];
  for (const sheet of sheets) {
    for (const [address, cell] of Object.entries(sheet.cells ?? {})) {
      if (isFormulaCell(cell?.raw)) formulaNodes.push(nodeKey(sheet.id, address));
    }
  }

  const plan = graph.plan(formulaNodes, true);
  const result: WorkbookCalculation = {
    sheets: Object.fromEntries(sheets.map(sheet => [sheet.id, emptySheetCalculation(sheet.id)])),
    graph,
    plan,
  };

  for (const node of plan.circular) {
    const separator = node.lastIndexOf("!");
    const sheetId = node.slice(0, separator);
    const address = node.slice(separator + 1);
    const target = result.sheets[sheetId];
    if (!target) continue;
    target.values[address] = "#CIRC!";
    target.errors[address] = "#CIRC!";
    target.circular.push(address);
  }

  for (const node of plan.blocked) {
    const separator = node.lastIndexOf("!");
    const sheetId = node.slice(0, separator);
    const address = node.slice(separator + 1);
    const target = result.sheets[sheetId];
    if (!target) continue;
    target.values[address] = "#CIRC!";
    target.errors[address] = "#CIRC!";
    target.blocked.push(address);
  }

  const orderedNodes = plan.order;
  for (const node of orderedNodes) {
    const separator = node.lastIndexOf("!");
    const sheetId = node.slice(0, separator);
    const address = node.slice(separator + 1);
    const sheet = sheets.find(item => item.id === sheetId);
    const sheetResult = result.sheets[sheetId];
    const raw = sheet?.cells?.[address]?.raw ?? "";
    if (!sheet || !sheetResult || !isFormulaCell(raw)) continue;

    const calculated = evaluate(raw, sheet.cells, new Set([node]), {
      currentSheetId: sheetId,
      sheets: cellMaps,
      sheetNameToId,
      namedRanges: options.namedRanges,
      spills: collectSpillLookup(result.sheets),
      maxArrayCells: options.maxArrayCells ?? MAX_SPILL_CELLS,
    });

    if (!Array.isArray(calculated)) {
      sheetResult.values[address] = calculated;
      if (typeof calculated === "string" && calculated.startsWith("#")) sheetResult.errors[address] = calculated;
      continue;
    }

    const matrix = calculated as FormulaMatrix;
    const width = matrix[0]?.length ?? 0;
    if (!matrix.length || !width) {
      sheetResult.values[address] = "#CALC!";
      sheetResult.errors[address] = "#CALC!";
      continue;
    }

    // Normalize ragged outputs to a rectangle. Excel's array results are
    // rectangular, while this also prevents undefined grid values if a helper
    // function accidentally returns ragged rows.
    const normalized = matrix.map(row => Array.from({ length: width }, (_, column) => row[column] ?? "#N/A"));
    const bounds = spillBounds(address, normalized);
    if (!bounds) {
      sheetResult.values[address] = "#CALC!";
      sheetResult.errors[address] = "#CALC!";
      continue;
    }

    const cellCount = normalized.length * width;
    if (cellCount > (options.maxArrayCells ?? MAX_SPILL_CELLS)) {
      sheetResult.values[address] = "#NUM!";
      sheetResult.errors[address] = "#NUM!";
      continue;
    }

    if (cellCount === 1) {
      const value = normalized[0][0];
      sheetResult.values[address] = value;
      if (typeof value === "string" && value.startsWith("#")) sheetResult.errors[address] = value;
      continue;
    }

    let collision = false;
    for (let row = bounds.startRow; row <= bounds.endRow && !collision; row++) {
      for (let column = bounds.startCol; column <= bounds.endCol; column++) {
        const targetAddress = addressOf(row, column);
        if (targetAddress === address) continue;
        const existingRaw = sheet.cells[targetAddress]?.raw ?? "";
        if (String(existingRaw).trim() !== "" || sheetResult.ghostCells[targetAddress]) {
          collision = true;
          break;
        }
      }
    }

    if (collision) {
      sheetResult.values[address] = "#SPILL!";
      sheetResult.errors[address] = "#SPILL!";
      continue;
    }

    sheetResult.values[address] = normalized[0][0];
    sheetResult.spillMatrices[address] = normalized;
    sheetResult.spillOrigins[address] = { address, range: bounds, matrix: normalized };
    for (let row = 0; row < normalized.length; row++) {
      for (let column = 0; column < width; column++) {
        if (row === 0 && column === 0) continue;
        const ghostAddress = addressOf(bounds.startRow + row, bounds.startCol + column);
        sheetResult.ghostCells[ghostAddress] = {
          origin: address,
          sheetId,
          value: normalized[row][column],
        };
      }
    }
  }

  return result;
}

function collectSpillLookup(
  sheets: Record<string, SheetCalculation>,
): Record<string, FormulaMatrix> {
  const lookup: Record<string, FormulaMatrix> = {};
  for (const [sheetId, calculation] of Object.entries(sheets)) {
    for (const [address, matrix] of Object.entries(calculation.spillMatrices)) {
      lookup[nodeKey(sheetId, address)] = matrix;
      // Single-sheet legacy formulas may use the bare address spill operator.
      if (!(address in lookup)) lookup[address] = matrix;
    }
  }
  return lookup;
}

export function calculatedCellValue(
  calculation: SheetCalculation | undefined,
  address: string,
): FormulaScalar | undefined {
  const key = address.toUpperCase();
  const ghost = calculation?.ghostCells[key];
  if (ghost) return ghost.value;
  return calculation?.values[key];
}

export function spillOriginForCell(
  calculation: SheetCalculation | undefined,
  address: string,
): string | null {
  return calculation?.ghostCells[address.toUpperCase()]?.origin ?? null;
}

export function isSpillCell(
  calculation: SheetCalculation | undefined,
  address: string,
): boolean {
  return !!calculation?.ghostCells[address.toUpperCase()];
}

export function isCellInSpillRange(
  calculation: SheetCalculation | undefined,
  address: string,
): boolean {
  const point = parseAddress(address);
  if (!point || !calculation) return false;
  return Object.values(calculation.spillOrigins).some(({ range }) =>
    point.row >= range.startRow && point.row <= range.endRow &&
    point.col >= range.startCol && point.col <= range.endCol,
  );
}

export function formatCalculatedValue(
  raw: string,
  value: FormulaScalar | undefined,
  style?: CellStyle,
  showFormula = false,
): string {
  if (showFormula && raw.trim().startsWith("=")) return raw;
  const current = value ?? raw;
  if (typeof current === "string" && current.startsWith("#")) return current;
  const numeric = typeof current === "number"
    ? current
    : (String(current).trim() !== "" && Number.isFinite(Number(current)) ? Number(current) : null);

  switch (style?.numberFormat) {
    case "Percentage %":
      return numeric === null ? String(current) : new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: style.decimalPlaces ?? 2 }).format(numeric);
    case "Currency ₹":
      return numeric === null ? String(current) : new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: style.decimalPlaces ?? 2 }).format(numeric);
    case "Currency $":
      return numeric === null ? String(current) : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: style.decimalPlaces ?? 2 }).format(numeric);
    case "Number":
      return numeric === null ? String(current) : new Intl.NumberFormat("en-IN", { maximumFractionDigits: style.decimalPlaces ?? 2 }).format(numeric);
    case "Date": {
      const date = new Date(String(current));
      return Number.isNaN(date.getTime()) ? String(current) : date.toLocaleDateString();
    }
    default:
      return String(current ?? "");
  }
}

export function formatGhostValue(value: FormulaScalar, style?: CellStyle): string {
  return formatCalculatedValue("", value, style, false);
}
