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
  spillCandidates: Record<string, CellRange>;
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
    spillCandidates: {},
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
 * The graph can be retained between edits through calculateWorkbookIncremental
 * so only downstream formula cells are recalculated after a cell change.
 */
function buildGraph(sheets: SheetData[], options: CalculationOptions): DependencyGraph {
  return DependencyGraph.fromSheets(
    sheets.map(sheet => ({ id: sheet.id, name: sheet.name, cells: sheet.cells })),
    {
      namedRanges: options.namedRanges,
      maxExpandedRangeCells: options.maxDependencyRangeCells,
    },
  );
}

function formulaNodes(sheets: SheetData[]): string[] {
  const nodes: string[] = [];
  for (const sheet of sheets) {
    for (const [address, cell] of Object.entries(sheet.cells ?? {})) {
      if (isFormulaCell(cell?.raw)) nodes.push(nodeKey(sheet.id, address));
    }
  }
  return nodes;
}

function cloneSheetCalculation(previous: SheetCalculation, sheetId: string): SheetCalculation {
  return {
    sheetId,
    values: { ...previous.values },
    spillMatrices: { ...previous.spillMatrices },
    spillOrigins: { ...previous.spillOrigins },
    spillCandidates: { ...previous.spillCandidates },
    ghostCells: { ...previous.ghostCells },
    errors: { ...previous.errors },
    circular: [...previous.circular],
    blocked: [...previous.blocked],
  };
}

function clearCalculatedNode(calculation: SheetCalculation, address: string): void {
  delete calculation.values[address];
  delete calculation.spillMatrices[address];
  delete calculation.spillOrigins[address];
  delete calculation.spillCandidates[address];
  delete calculation.errors[address];
  calculation.circular = calculation.circular.filter(item => item !== address);
  calculation.blocked = calculation.blocked.filter(item => item !== address);
  for (const [ghostAddress, ghost] of Object.entries(calculation.ghostCells)) {
    if (ghost.origin === address) delete calculation.ghostCells[ghostAddress];
  }
}

function applyCalculationPlan(
  sheets: SheetData[],
  options: CalculationOptions,
  result: WorkbookCalculation,
  plan: DependencyPlan,
): void {
  const cellMaps: Record<string, CellMap> = {};
  for (const sheet of sheets) cellMaps[sheet.id] = sheet.cells ?? {};
  const sheetNameToId = normalizeSheetLookup(sheets);

  for (const node of plan.circular) {
    const splitAt = node.lastIndexOf("!");
    const sheetId = node.slice(0, splitAt);
    const address = node.slice(splitAt + 1);
    const target = result.sheets[sheetId];
    if (!target) continue;
    target.values[address] = "#CIRC!";
    target.errors[address] = "#CIRC!";
    if (!target.circular.includes(address)) target.circular.push(address);
  }

  for (const node of plan.blocked) {
    const splitAt = node.lastIndexOf("!");
    const sheetId = node.slice(0, splitAt);
    const address = node.slice(splitAt + 1);
    const target = result.sheets[sheetId];
    if (!target) continue;
    target.values[address] = "#CIRC!";
    target.errors[address] = "#CIRC!";
    if (!target.blocked.includes(address)) target.blocked.push(address);
  }

  for (const node of plan.order) {
    const splitAt = node.lastIndexOf("!");
    const sheetId = node.slice(0, splitAt);
    const address = node.slice(splitAt + 1);
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

    // Retain the intended footprint even when blocked, so clearing an
    // obstructing cell can invalidate this origin and immediately re-spill.
    sheetResult.spillCandidates[address] = bounds;
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
}

export function calculateWorkbook(
  sourceSheets: SheetData[],
  options: CalculationOptions = {},
): WorkbookCalculation {
  const sheets = sourceSheets ?? [];
  const graph = buildGraph(sheets, options);
  const plan = graph.plan(formulaNodes(sheets), true);
  const result: WorkbookCalculation = {
    sheets: Object.fromEntries(sheets.map(sheet => [sheet.id, emptySheetCalculation(sheet.id)])),
    graph,
    plan,
  };
  applyCalculationPlan(sheets, options, result, plan);
  return result;
}

export function changedCellNodes(previousSheets: SheetData[], nextSheets: SheetData[]): string[] {
  const previousById = new Map(previousSheets.map(sheet => [sheet.id, sheet]));
  const changed = new Set<string>();
  for (const sheet of nextSheets) {
    const previous = previousById.get(sheet.id);
    const before = previous?.cells ?? {};
    const after = sheet.cells ?? {};
    const addresses = new Set([...Object.keys(before), ...Object.keys(after)]);
    for (const address of addresses) {
      if ((before[address]?.raw ?? "") !== (after[address]?.raw ?? "")) {
        changed.add(nodeKey(sheet.id, address));
      }
    }
  }
  return [...changed];
}

export function hasSameSheetStructure(previousSheets: SheetData[], nextSheets: SheetData[]): boolean {
  return previousSheets.length === nextSheets.length &&
    previousSheets.every((sheet, index) =>
      sheet.id === nextSheets[index]?.id && sheet.name === nextSheets[index]?.name,
    );
}

/**
 * Rebuilds the dependency graph for the edited workbook, invalidates the
 * changed cells and their downstream formula dependents, then evaluates only
 * that topological subset. Previous spill footprints are watched as well so
 * edits inside a blocked spill range recalculate its origin.
 */
export function calculateWorkbookIncremental(
  sourceSheets: SheetData[],
  previous: WorkbookCalculation,
  changedNodes: Iterable<string>,
  options: CalculationOptions = {},
): WorkbookCalculation {
  const sheets = sourceSheets ?? [];
  const graph = buildGraph(sheets, options);
  const roots = new Set(changedNodes);

  // A spill obstruction is not a formula precedent, but editing it changes
  // whether the array origin can spill. Add every affected origin as a root.
  for (const [sheetId, oldCalculation] of Object.entries(previous.sheets)) {
    for (const [origin, range] of Object.entries(oldCalculation.spillCandidates)) {
      const obstructorChanged = [...roots].some(node => {
        const separator = node.lastIndexOf("!");
        if (node.slice(0, separator) !== sheetId) return false;
        const address = node.slice(separator + 1);
        const point = parseAddress(address);
        return !!point && address !== origin &&
          point.row >= range.startRow && point.row <= range.endRow &&
          point.col >= range.startCol && point.col <= range.endCol;
      });
      if (obstructorChanged) roots.add(nodeKey(sheetId, origin));
    }
  }

  const plan = graph.plan(roots, true);
  const nextResults: Record<string, SheetCalculation> = {};
  for (const sheet of sheets) {
    const oldCalculation = previous.sheets[sheet.id];
    nextResults[sheet.id] = oldCalculation
      ? cloneSheetCalculation(oldCalculation, sheet.id)
      : emptySheetCalculation(sheet.id);
  }
  const result: WorkbookCalculation = { sheets: nextResults, graph, plan };

  // Remove cached results from every changed formula/source and every
  // dependent formula before evaluating the affected subgraph.
  const clearNodes = new Set([...roots, ...plan.affected]);
  for (const node of clearNodes) {
    const separator = node.lastIndexOf("!");
    const sheetId = node.slice(0, separator);
    const address = node.slice(separator + 1);
    const calculation = result.sheets[sheetId];
    if (calculation) clearCalculatedNode(calculation, address);
  }

  applyCalculationPlan(sheets, options, result, plan);
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
