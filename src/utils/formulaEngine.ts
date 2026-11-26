import type { CellMap, CellStyle } from "../types";
import { addressOf, parseAddress } from "../types";

export type FormulaScalar = string | number | boolean;
export type FormulaMatrix = FormulaScalar[][];
export type FormulaValue = FormulaScalar | FormulaMatrix;

export interface FormulaNamedRange {
  sheetId?: string;
  reference: string;
}

export interface FormulaContext {
  cells: CellMap;
  currentSheetId?: string;
  sheets?: Record<string, CellMap>;
  sheetNameToId?: Record<string, string>;
  namedRanges?: Record<string, string | FormulaNamedRange>;
  spills?: Record<string, FormulaMatrix>;
  maxArrayCells?: number;
}

type TokenKind = "number" | "string" | "identifier" | "reference" | "error" | "operator" | "punctuation";
interface Token { kind: TokenKind; value: string; }
type ReferenceNode = { kind: "reference"; sheetName?: string; address: string };
type Ast =
  | { kind: "literal"; value: FormulaScalar }
  | ReferenceNode
  | { kind: "range"; start: ReferenceNode; end: ReferenceNode }
  | { kind: "spill"; reference: ReferenceNode }
  | { kind: "name"; name: string }
  | { kind: "unary"; operator: string; operand: Ast }
  | { kind: "binary"; operator: string; left: Ast; right: Ast }
  | { kind: "call"; name: string; args: Ast[] };

class FormulaError extends Error {
  constructor(public readonly code: string) { super(code); }
}

const CELL_REF = /^\$?[A-Z]{1,3}\$?[1-9]\d*$/i;
const ERROR_REF = /^#(?:DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|CIRC!|CIRCULAR!|FIELD!|BLOCKED!)/i;
const MAX_SAFE_ARRAY_CELLS = 1000000;

function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;
  while (index < source.length) {
    const char = source[index];
    if (/\s/.test(char)) { index++; continue; }

    if (char === '"') {
      let value = "";
      index++;
      let closed = false;
      while (index < source.length) {
        if (source[index] === '"' && source[index + 1] === '"') {
          value += '"';
          index += 2;
        } else if (source[index] === '"') {
          index++;
          closed = true;
          break;
        } else {
          value += source[index++];
        }
      }
      if (!closed) throw new FormulaError("#VALUE!");
      tokens.push({ kind: "string", value });
      continue;
    }

    if (char === "'") {
      let sheetName = "";
      index++;
      let closed = false;
      while (index < source.length) {
        if (source[index] === "'" && source[index + 1] === "'") {
          sheetName += "'";
          index += 2;
        } else if (source[index] === "'") {
          index++;
          closed = true;
          break;
        } else {
          sheetName += source[index++];
        }
      }
      if (!closed || source[index] !== "!") throw new FormulaError("#PARSE!");
      index++;
      const refMatch = source.slice(index).match(/^\$?[A-Z]{1,3}\$?[1-9]\d*/i);
      if (!refMatch) throw new FormulaError("#REF!");
      tokens.push({ kind: "reference", value: "'" + sheetName + "'!" + refMatch[0] });
      index += refMatch[0].length;
      continue;
    }

    const errorMatch = source.slice(index).match(ERROR_REF);
    if (errorMatch) {
      tokens.push({ kind: "error", value: errorMatch[0].toUpperCase() });
      index += errorMatch[0].length;
      continue;
    }

    const unquotedSheet = source.slice(index).match(/^([A-Z_][A-Z0-9_.]*)!(\$?[A-Z]{1,3}\$?[1-9]\d*)/i);
    if (unquotedSheet) {
      tokens.push({ kind: "reference", value: unquotedSheet[0] });
      index += unquotedSheet[0].length;
      continue;
    }

    const refMatch = source.slice(index).match(/^\$?[A-Z]{1,3}\$?[1-9]\d*/i);
    if (refMatch && CELL_REF.test(refMatch[0])) {
      tokens.push({ kind: "reference", value: refMatch[0] });
      index += refMatch[0].length;
      continue;
    }

    const numberMatch = source.slice(index).match(/^(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?/);
    if (numberMatch) {
      tokens.push({ kind: "number", value: numberMatch[0] });
      index += numberMatch[0].length;
      continue;
    }

    const identifierMatch = source.slice(index).match(/^[A-Z_][A-Z0-9_.]*/i);
    if (identifierMatch) {
      tokens.push({ kind: "identifier", value: identifierMatch[0] });
      index += identifierMatch[0].length;
      continue;
    }

    const operator = source.slice(index).match(/^(?:>=|<=|<>|[+\-*/^&=%<>])/);
    if (operator) {
      tokens.push({ kind: "operator", value: operator[0] });
      index += operator[0].length;
      continue;
    }

    if ("(),:;#".includes(char)) {
      tokens.push({ kind: "punctuation", value: char });
      index++;
      continue;
    }

    throw new FormulaError("#PARSE!");
  }
  return tokens;
}

class Parser {
  private index = 0;
  constructor(private readonly tokens: Token[]) {}
  private peek(): Token | undefined { return this.tokens[this.index]; }
  private take(): Token | undefined { return this.tokens[this.index++]; }
  private match(value: string): boolean {
    if (this.peek()?.value === value) { this.index++; return true; }
    return false;
  }

  parse(): Ast {
    const result = this.expression(0);
    if (this.index !== this.tokens.length) throw new FormulaError("#PARSE!");
    return result;
  }

  private expression(minimumPrecedence: number): Ast {
    let left = this.prefix();
    while (true) {
      const token = this.peek();
      if (!token || token.kind !== "operator") break;
      const precedence = this.precedence(token.value);
      if (precedence < minimumPrecedence) break;
      this.take();
      const right = this.expression(precedence + (token.value === "^" ? 0 : 1));
      left = { kind: "binary", operator: token.value, left, right };
    }
    return left;
  }

  private precedence(operator: string): number {
    if (["=", "<>", ">", "<", ">=", "<="].includes(operator)) return 1;
    if (operator === "&") return 2;
    if (operator === "+" || operator === "-") return 3;
    if (operator === "*" || operator === "/") return 4;
    if (operator === "^") return 5;
    return -1;
  }

  private prefix(): Ast {
    const token = this.take();
    if (!token) throw new FormulaError("#PARSE!");

    if (token.kind === "operator" && (token.value === "+" || token.value === "-")) {
      return { kind: "unary", operator: token.value, operand: this.expression(6) };
    }
    if (token.value === "(") {
      const value = this.expression(0);
      if (!this.match(")")) throw new FormulaError("#PARSE!");
      return value;
    }
    if (token.kind === "number") return { kind: "literal", value: Number(token.value) };
    if (token.kind === "string") return { kind: "literal", value: token.value };
    if (token.kind === "error") return { kind: "literal", value: token.value };

    if (token.kind === "reference") {
      const reference = this.parseReference(token.value);
      let node: Ast = reference;
      if (this.match(":")) {
        const endToken = this.take();
        if (!endToken || endToken.kind !== "reference") throw new FormulaError("#REF!");
        const end = this.parseReference(endToken.value);
        node = { kind: "range", start: reference, end: { ...end, sheetName: end.sheetName ?? reference.sheetName } };
      }
      if (this.match("#")) {
        if (node.kind !== "reference") throw new FormulaError("#REF!");
        node = { kind: "spill", reference: node };
      }
      if (this.match("%")) node = { kind: "unary", operator: "%", operand: node };
      return node;
    }

    if (token.kind === "identifier") {
      const upper = token.value.toUpperCase();
      if (upper === "TRUE") return { kind: "literal", value: true };
      if (upper === "FALSE") return { kind: "literal", value: false };
      if (this.match("(")) {
        const args: Ast[] = [];
        if (!this.match(")")) {
          do { args.push(this.expression(0)); } while (this.match(",") || this.match(";"));
          if (!this.match(")")) throw new FormulaError("#PARSE!");
        }
        return { kind: "call", name: upper, args };
      }
      if (this.match("%")) return { kind: "unary", operator: "%", operand: { kind: "name", name: token.value } };
      return { kind: "name", name: token.value };
    }

    throw new FormulaError("#PARSE!");
  }

  private parseReference(value: string): ReferenceNode {
    const bang = value.lastIndexOf("!");
    const sheetName = bang >= 0
      ? value.slice(0, bang).replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'")
      : undefined;
    const address = value.slice(bang + 1).replace(/\$/g, "").toUpperCase();
    if (!parseAddress(address)) throw new FormulaError("#REF!");
    return { kind: "reference", sheetName, address };
  }
}

function isMatrix(value: FormulaValue | undefined): value is FormulaMatrix {
  return Array.isArray(value);
}

function matrixOf(value: FormulaValue): FormulaMatrix {
  return isMatrix(value) ? value : [[value]];
}

function flatten(value: FormulaValue): FormulaScalar[] {
  return isMatrix(value) ? value.reduce<FormulaScalar[]>((all, row) => all.concat(row), []) : [value];
}

function first(value: FormulaValue | undefined, fallback: FormulaScalar = 0): FormulaScalar {
  if (value === undefined) return fallback;
  if (!isMatrix(value)) return value;
  return value[0]?.[0] ?? fallback;
}

function numberOf(value: FormulaScalar): number {
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value.trim() === "") return 0;
  const result = Number(value);
  if (!Number.isFinite(result)) throw new FormulaError("#VALUE!");
  return result;
}

function booleanOf(value: FormulaScalar): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const normalized = value.trim().toUpperCase();
  if (normalized === "TRUE") return true;
  if (normalized === "FALSE" || normalized === "") return false;
  return true;
}

function isError(value: FormulaValue): boolean {
  return flatten(value).some(item => typeof item === "string" && /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|CIRC!|CIRCULAR!)/i.test(item));
}

function sheetMap(context: FormulaContext, sheetName?: string): { map: CellMap; sheetId: string } {
  if (!sheetName) return { map: context.cells, sheetId: context.currentSheetId ?? "Sheet1" };
  const resolvedId = context.sheetNameToId?.[sheetName] ?? context.sheetNameToId?.[sheetName.toLowerCase()] ?? sheetName;
  const map = context.sheets?.[resolvedId] ?? context.sheets?.[sheetName] ?? context.cells;
  return { map, sheetId: resolvedId };
}

function parseNamedReference(name: string, context: FormulaContext): { sheetName?: string; start: string; end: string } | null {
  const definition = context.namedRanges?.[name] ?? context.namedRanges?.[name.toLowerCase()];
  if (!definition) return null;
  const reference = typeof definition === "string" ? definition : definition.reference;
  const explicitSheet = typeof definition === "string" ? undefined : definition.sheetId;
  const match = reference.match(/^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(\$?[A-Z]{1,3}\$?[1-9]\d*)(?:\s*:\s*(\$?[A-Z]{1,3}\$?[1-9]\d*))?$/i);
  if (!match) throw new FormulaError("#REF!");
  return {
    sheetName: explicitSheet ?? match[1] ?? match[2],
    start: match[3].replace(/\$/g, "").toUpperCase(),
    end: (match[4] ?? match[3]).replace(/\$/g, "").toUpperCase(),
  };
}

function cellValue(
  address: string,
  context: FormulaContext,
  sheetName: string | undefined,
  seen: Set<string>,
  scope: Map<string, FormulaValue>,
): FormulaValue {
  const selected = sheetMap(context, sheetName);
  const key = selected.sheetId + "!" + address;
  const cell = selected.map[address.toUpperCase()];
  if (!cell) return 0;
  const raw = String(cell.raw ?? "");
  if (!raw.trim().startsWith("=")) {
    if (raw.trim() === "") return 0;
    if (/^(TRUE|FALSE)$/i.test(raw.trim())) return raw.trim().toUpperCase() === "TRUE";
    const number = Number(raw);
    return Number.isFinite(number) ? number : raw;
  }
  if (seen.has(key)) return "#CIRC!";
  const nextSeen = new Set(seen);
  nextSeen.add(key);
  return evaluateInContext(raw, { ...context, cells: selected.map, currentSheetId: selected.sheetId }, nextSeen, scope);
}

function rangeValue(
  start: ReferenceNode,
  end: ReferenceNode,
  context: FormulaContext,
  seen: Set<string>,
  scope: Map<string, FormulaValue>,
): FormulaMatrix {
  const selected = sheetMap(context, start.sheetName ?? end.sheetName);
  const a = parseAddress(start.address), b = parseAddress(end.address);
  if (!a || !b) throw new FormulaError("#REF!");
  const rows = Math.abs(b.row - a.row) + 1;
  const columns = Math.abs(b.col - a.col) + 1;
  if (rows * columns > (context.maxArrayCells ?? MAX_SAFE_ARRAY_CELLS)) throw new FormulaError("#NUM!");
  const result: FormulaMatrix = [];
  for (let row = Math.min(a.row, b.row); row <= Math.max(a.row, b.row); row++) {
    const values: FormulaScalar[] = [];
    for (let column = Math.min(a.col, b.col); column <= Math.max(a.col, b.col); column++) {
      values.push(first(cellValue(addressOf(row, column), context, selected.sheetId, seen, scope)));
    }
    result.push(values);
  }
  return result;
}

function namedValue(name: string, context: FormulaContext, seen: Set<string>, scope: Map<string, FormulaValue>): FormulaValue {
  const local = scope.get(name.toLowerCase());
  if (local !== undefined) return local;
  const reference = parseNamedReference(name, context);
  if (!reference) throw new FormulaError("#NAME?");
  const start: ReferenceNode = { kind: "reference", address: reference.start, sheetName: reference.sheetName };
  const end: ReferenceNode = { kind: "reference", address: reference.end, sheetName: reference.sheetName };
  const result = rangeValue(start, end, context, seen, scope);
  return reference.start === reference.end ? (result[0]?.[0] ?? 0) : result;
}

function evaluateReference(reference: ReferenceNode, context: FormulaContext, seen: Set<string>, scope: Map<string, FormulaValue>): FormulaValue {
  return cellValue(reference.address, context, reference.sheetName, seen, scope);
}

function criteriaMatches(value: FormulaScalar, criterion: FormulaScalar): boolean {
  const text = String(criterion);
  const match = text.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  if (match) {
    const expected = match[2];
    const numericExpected = expected.trim() !== "" && Number.isFinite(Number(expected)) ? Number(expected) : null;
    const actualNumeric = typeof value === "number" ? value : (String(value).trim() !== "" && Number.isFinite(Number(value)) ? Number(value) : null);
    const left: string | number = numericExpected !== null && actualNumeric !== null ? actualNumeric : String(value).toLocaleLowerCase();
    const right: string | number = numericExpected !== null && actualNumeric !== null ? numericExpected : expected.toLocaleLowerCase();
    switch (match[1]) {
      case ">": return left > right;
      case "<": return left < right;
      case ">=": return left >= right;
      case "<=": return left <= right;
      case "<>": return left !== right;
      default: return left === right;
    }
  }
  if (text.includes("*") || text.includes("?")) {
    const escaped = text.replace(/[.*+?^$()|[\]\\]/g, "\\$&").replace(/[{}]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
    return new RegExp("^" + escaped + "$", "i").test(String(value));
  }
  if (typeof value === "number" && text.trim() !== "" && Number.isFinite(Number(text))) return value === Number(text);
  return String(value).toLocaleLowerCase() === text.toLocaleLowerCase();
}

function sortMatrix(matrix: FormulaMatrix, sortIndex: number, sortOrder: number): FormulaMatrix {
  const copy = matrix.map(row => [...row]);
  const column = Math.max(0, Math.min(copy[0]?.length ? copy[0].length - 1 : 0, Math.abs(sortIndex) - 1));
  copy.sort((a, b) => {
    const left = a[column], right = b[column];
    const compare = typeof left === "number" && typeof right === "number"
      ? left - right
      : String(left ?? "").localeCompare(String(right ?? ""), undefined, { numeric: true, sensitivity: "base" });
    return compare * (sortOrder < 0 ? -1 : 1);
  });
  return copy;
}

function wildcardMatch(value: FormulaScalar, pattern: string): boolean {
  const escaped = pattern.replace(/[.*+?^$()|[\]\\]/g, "\\$&").replace(/[{}]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
  return new RegExp("^" + escaped + "$", "i").test(String(value));
}

function lookupIndex(lookupValue: FormulaScalar, values: FormulaScalar[], matchMode: number, searchMode: number): number {
  const indices = Array.from({ length: values.length }, (_, index) => searchMode < 0 ? values.length - 1 - index : index);
  if (matchMode === 0 || matchMode === 2) {
    for (const index of indices) {
      const isMatch = matchMode === 2
        ? wildcardMatch(values[index], String(lookupValue))
        : String(values[index]).toLocaleLowerCase() === String(lookupValue).toLocaleLowerCase();
      if (isMatch) return index;
    }
    return -1;
  }
  let bestIndex = -1;
  let bestDistance = Number.POSITIVE_INFINITY;
  const target = numberOf(lookupValue);
  for (const index of indices) {
    const candidate = Number(values[index]);
    if (!Number.isFinite(candidate)) continue;
    if (matchMode === -1 && candidate <= target && target - candidate < bestDistance) {
      bestIndex = index;
      bestDistance = target - candidate;
    }
    if (matchMode === 1 && candidate >= target && candidate - target < bestDistance) {
      bestIndex = index;
      bestDistance = candidate - target;
    }
  }
  return bestIndex;
}

function matchIndex(lookupValue: FormulaScalar, values: FormulaScalar[], matchType: number): number {
  if (matchType === 0) return lookupIndex(lookupValue, values, 0, 1);
  if (matchType !== 1 && matchType !== -1) throw new FormulaError("#VALUE!");
  const target = numberOf(lookupValue);
  let bestIndex = -1;
  let bestValue = matchType === 1 ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY;
  values.forEach((value, index) => {
    const candidate = typeof value === "number" ? value : Number(value);
    if (!Number.isFinite(candidate)) return;
    if (matchType === 1 && candidate <= target && candidate > bestValue) {
      bestValue = candidate;
      bestIndex = index;
    }
    if (matchType === -1 && candidate >= target && candidate < bestValue) {
      bestValue = candidate;
      bestIndex = index;
    }
  });
  return bestIndex;
}

function evaluateCall(
  name: string,
  args: Ast[],
  context: FormulaContext,
  seen: Set<string>,
  scope: Map<string, FormulaValue>,
): FormulaValue {
  const get = (index: number, fallback: FormulaValue = 0): FormulaValue =>
    args[index] ? evaluateAst(args[index], context, seen, scope) : fallback;
  const scalar = (index: number, fallback: FormulaScalar = 0): FormulaScalar => first(get(index, fallback));
  const matrix = (index: number): FormulaMatrix => matrixOf(get(index));
  const flat = (index: number): FormulaScalar[] => flatten(get(index));
  const bool = (index: number, fallback = false): boolean => args[index] ? booleanOf(scalar(index)) : fallback;
  const num = (index: number, fallback = 0): number => args[index] ? numberOf(scalar(index)) : fallback;

  switch (name) {
    case "IF": {
      if (args.length < 2) throw new FormulaError("#VALUE!");
      const branch = booleanOf(scalar(0)) ? args[1] : (args[2] ?? { kind: "literal", value: false });
      return evaluateAst(branch, context, seen, scope);
    }
    case "IFERROR": {
      try { const value = get(0); return isError(value) ? get(1, "") : value; }
      catch { return get(1, ""); }
    }
    case "IFNA": {
      try {
        const value = get(0);
        return flatten(value).some(item => item === "#N/A") ? get(1, "") : value;
      } catch (error) {
        return error instanceof FormulaError && error.code === "#N/A" ? get(1, "") : "#VALUE!";
      }
    }
    case "LET": {
      if (args.length < 3 || args.length % 2 !== 1) throw new FormulaError("#VALUE!");
      const nextScope = new Map(scope);
      for (let index = 0; index < args.length - 1; index += 2) {
        const nameNode = args[index];
        if (nameNode.kind !== "name") throw new FormulaError("#NAME?");
        nextScope.set(nameNode.name.toLowerCase(), evaluateAst(args[index + 1], context, seen, nextScope));
      }
      return evaluateAst(args[args.length - 1], context, seen, nextScope);
    }
    case "IFS": {
      if (args.length < 2 || args.length % 2 !== 0) throw new FormulaError("#VALUE!");
      for (let index = 0; index < args.length; index += 2) {
        if (booleanOf(first(evaluateAst(args[index], context, seen, scope)))) return evaluateAst(args[index + 1], context, seen, scope);
      }
      throw new FormulaError("#N/A");
    }
    case "SWITCH": {
      if (args.length < 3) throw new FormulaError("#VALUE!");
      const expression = first(get(0));
      for (let index = 1; index + 1 < args.length; index += 2) {
        if (String(first(evaluateAst(args[index], context, seen, scope))).toLocaleLowerCase() === String(expression).toLocaleLowerCase()) {
          return evaluateAst(args[index + 1], context, seen, scope);
        }
      }
      return args.length % 2 === 0 ? get(args.length - 1, "#N/A") : "#N/A";
    }
    case "FILTER": {
      const source = matrix(0);
      const include = flatten(get(1)).map(booleanOf);
      const ifEmpty = get(2, "#CALC!");
      if (source.length === include.length) {
        const result = source.filter((_, index) => include[index]);
        return result.length ? result : ifEmpty;
      }
      if (source[0]?.length === include.length) {
        const result = source.map(row => row.filter((_, index) => include[index]));
        return result[0]?.length ? result : ifEmpty;
      }
      throw new FormulaError("#VALUE!");
    }
    case "SORT": {
      const source = matrix(0);
      const sortIndex = Math.trunc(num(1, 1));
      const sortOrder = num(2, 1);
      if (bool(3, false)) {
        const columns = source[0]?.map((_, column) => source.map(row => row[column])) ?? [];
        const sorted = sortMatrix(columns, sortIndex, sortOrder);
        return sorted[0]?.map((_, row) => sorted.map(column => column[row])) ?? [];
      }
      return sortMatrix(source, sortIndex, sortOrder);
    }
    case "SORTBY": {
      const source = matrix(0);
      const criteria: Array<{ values: FormulaScalar[]; order: number }> = [];
      for (let index = 1; index < args.length; index += 2) {
        const values = flat(index);
        if (values.length !== source.length) throw new FormulaError("#VALUE!");
        criteria.push({ values, order: args[index + 1] ? num(index + 1, 1) : 1 });
      }
      if (!criteria.length) throw new FormulaError("#VALUE!");
      return source.map((row, index) => ({ row, index }))
        .sort((left, right) => {
          for (const criterion of criteria) {
            const a = criterion.values[left.index], b = criterion.values[right.index];
            const compare = typeof a === "number" && typeof b === "number"
              ? a - b
              : String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" });
            if (compare !== 0) return compare * (criterion.order < 0 ? -1 : 1);
          }
          return left.index - right.index;
        }).map(item => item.row);
    }
    case "UNIQUE": {
      const source = matrix(0);
      const byColumn = bool(1, false), exactlyOnce = bool(2, false);
      const oriented = byColumn ? (source[0]?.map((_, column) => source.map(row => row[column])) ?? []) : source;
      const keys = oriented.map(row => JSON.stringify(row));
      const counts = new Map<string, number>();
      keys.forEach(key => counts.set(key, (counts.get(key) ?? 0) + 1));
      const seenKeys = new Set<string>();
      const unique = oriented.filter((_, index) => {
        const key = keys[index];
        if (exactlyOnce && counts.get(key) !== 1) return false;
        if (seenKeys.has(key)) return false;
        seenKeys.add(key);
        return true;
      });
      return byColumn ? (unique[0]?.map((_, row) => unique.map(column => column[row])) ?? []) : unique;
    }
    case "SEQUENCE": {
      const rows = Math.max(1, Math.trunc(num(0, 1)));
      const columns = Math.max(1, Math.trunc(num(1, 1)));
      const start = num(2, 1), step = num(3, 1);
      if (rows * columns > (context.maxArrayCells ?? MAX_SAFE_ARRAY_CELLS)) throw new FormulaError("#NUM!");
      return Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => start + (row * columns + column) * step));
    }
    case "TRANSPOSE": {
      const source = matrix(0);
      return source[0]?.map((_, column) => source.map(row => row[column])) ?? [];
    }
    case "TAKE":
    case "DROP": {
      const source = matrix(0);
      const rowsArg = Math.trunc(num(1, 1));
      const columnsArg = args[2] ? Math.trunc(num(2, source[0]?.length ?? 1)) : (source[0]?.length ?? 1);
      const height = source.length, width = source[0]?.length ?? 0;
      const rowStart = name === "TAKE" ? (rowsArg < 0 ? Math.max(0, height + rowsArg) : 0) : (rowsArg < 0 ? 0 : Math.min(height, rowsArg));
      const rowEnd = name === "TAKE" ? (rowsArg < 0 ? height : Math.min(height, rowsArg)) : (rowsArg < 0 ? Math.max(0, height + rowsArg) : height);
      const colStart = name === "TAKE" ? (columnsArg < 0 ? Math.max(0, width + columnsArg) : 0) : (columnsArg < 0 ? 0 : Math.min(width, columnsArg));
      const colEnd = name === "TAKE" ? (columnsArg < 0 ? width : Math.min(width, columnsArg)) : (columnsArg < 0 ? Math.max(0, width + columnsArg) : width);
      const result = source.slice(rowStart, rowEnd).map(row => row.slice(colStart, colEnd));
      return result.length && result[0].length ? result : "#CALC!";
    }
    case "HSTACK":
    case "VSTACK": {
      const inputs = args.map((_, index) => matrix(index));
      if (name === "VSTACK") {
        const width = Math.max(0, ...inputs.map(item => item[0]?.length ?? 0));
        return inputs.flatMap(item => item.map(row => Array.from({ length: width }, (_, index) => row[index] ?? "#N/A")));
      }
      const height = Math.max(0, ...inputs.map(item => item.length));
      return Array.from({ length: height }, (_, row) => inputs.flatMap(item => Array.from({ length: item[0]?.length ?? 0 }, (_, column) => item[row]?.[column] ?? "#N/A")));
    }
    case "XLOOKUP": {
      const lookup = scalar(0);
      const lookupValues = flat(1);
      const returnMatrix = matrix(2);
      const matchMode = num(4, 0), searchMode = num(5, 1);
      const index = lookupIndex(lookup, lookupValues, matchMode, searchMode);
      if (index < 0) return get(3, "#N/A");
      if (returnMatrix.length === lookupValues.length) {
        const matchingRow = returnMatrix[index] ?? ["#N/A"];
        return matchingRow.length === 1 ? matchingRow[0] : [matchingRow];
      }
      if (returnMatrix[0]?.length === lookupValues.length) {
        const matchingColumn = returnMatrix.map(row => [row[index] ?? "#N/A"]);
        return matchingColumn.length === 1 ? matchingColumn[0][0] : matchingColumn;
      }
      return flatten(returnMatrix)[index] ?? "#N/A";
    }
    case "XMATCH": {
      const index = lookupIndex(scalar(0), flat(1), num(2, 0), num(3, 1));
      return index < 0 ? "#N/A" : index + 1;
    }
    case "INDEX": {
      const source = matrix(0);
      const row = Math.trunc(num(1, 1)), column = Math.trunc(num(2, 1));
      if (row === 0 && column === 0) return source;
      if (row === 0) return source.map(line => [line[column - 1] ?? "#REF!"]);
      if (column === 0) return [source[row - 1] ?? ["#REF!"]];
      return source[row - 1]?.[column - 1] ?? "#REF!";
    }
    case "MATCH": {
      const values = flat(1);
      const mode = num(2, 1);
      const index = matchIndex(scalar(0), values, mode);
      return index < 0 ? "#N/A" : index + 1;
    }
    case "VLOOKUP":
    case "HLOOKUP": {
      const source = matrix(1);
      const index = Math.max(1, Math.trunc(num(2, 1)));
      const approximate = bool(3, true), horizontal = name === "HLOOKUP";
      const keys = horizontal ? (source[0] ?? []) : source.map(row => row[0] ?? "");
      const match = matchIndex(scalar(0), keys, approximate ? 1 : 0);
      if (match < 0) return "#N/A";
      return horizontal ? source[index - 1]?.[match] ?? "#REF!" : source[match]?.[index - 1] ?? "#REF!";
    }
    case "SUMIFS":
    case "COUNTIFS":
    case "AVERAGEIFS":
    case "MAXIFS":
    case "MINIFS": {
      const sumRange = name === "COUNTIFS" ? null : flatten(get(0));
      const firstCriteriaIndex = name === "COUNTIFS" ? 0 : 1;
      if ((args.length - firstCriteriaIndex) % 2 !== 0) throw new FormulaError("#VALUE!");
      const criteria: Array<{ values: FormulaScalar[]; expected: FormulaScalar }> = [];
      for (let index = firstCriteriaIndex; index + 1 < args.length; index += 2) {
        criteria.push({ values: flat(index), expected: scalar(index + 1) });
      }
      const count = criteria[0]?.values.length ?? sumRange?.length ?? 0;
      if (sumRange && sumRange.length !== count) throw new FormulaError("#VALUE!");
      if (criteria.some(item => item.values.length !== count)) throw new FormulaError("#VALUE!");
      const selected: number[] = [];
      for (let index = 0; index < count; index++) {
        if (!criteria.every(item => criteriaMatches(item.values[index], item.expected))) continue;
        if (name === "COUNTIFS") selected.push(1);
        else {
          const value = sumRange?.[index] ?? 0;
          const numeric = typeof value === "number" ? value : (String(value).trim() !== "" && Number.isFinite(Number(value)) ? Number(value) : null);
          if (numeric !== null) selected.push(numeric);
        }
      }
      if (name === "COUNTIFS") return selected.length;
      if (name === "SUMIFS") return selected.reduce((sum, value) => sum + value, 0);
      if (name === "AVERAGEIFS") return selected.length ? selected.reduce((sum, value) => sum + value, 0) / selected.length : "#DIV/0!";
      if (!selected.length) return 0;
      return name === "MAXIFS" ? Math.max(...selected) : Math.min(...selected);
    }
    case "TEXTSPLIT": {
      const text = String(scalar(0, ""));
      const columnDelimiter = String(scalar(1, ""));
      if (!columnDelimiter) throw new FormulaError("#VALUE!");
      const rowDelimiter = args[2] ? String(scalar(2)) : null;
      const ignoreEmpty = bool(3, false);
      const splitRows = rowDelimiter ? text.split(rowDelimiter) : [text];
      const rows = splitRows.map(line => {
        const parts = line.split(columnDelimiter);
        return ignoreEmpty ? parts.filter(part => part !== "") : parts;
      });
      const width = Math.max(0, ...rows.map(row => row.length));
      const pad = args[5] ? scalar(5) : "#N/A";
      return rows.map(row => Array.from({ length: width }, (_, index) => row[index] ?? pad));
    }
    case "TEXTJOIN": {
      const delimiter = String(scalar(0, ""));
      const ignoreEmpty = bool(1, false);
      const values = args.slice(2).flatMap((_, index) => flatten(get(index + 2))).filter(value => !(ignoreEmpty && String(value) === ""));
      return values.map(String).join(delimiter);
    }
    case "SUM":
    case "AVERAGE":
    case "COUNT":
    case "COUNTA":
    case "MIN":
    case "MAX": {
      const values = args.flatMap((_, index) => flatten(get(index)));
      if (name === "COUNTA") return values.filter(value => String(value) !== "").length;
      const numbers = values.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
      if (name === "COUNT") return numbers.length;
      if (name === "SUM") return numbers.reduce((sum, value) => sum + value, 0);
      if (name === "AVERAGE") return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
      if (!numbers.length) return 0;
      return name === "MIN" ? Math.min(...numbers) : Math.max(...numbers);
    }
    case "AND": return args.every((_, index) => flatten(get(index)).every(booleanOf));
    case "OR": return args.some((_, index) => flatten(get(index)).some(booleanOf));
    case "NOT": return !booleanOf(scalar(0));
    case "ABS": return Math.abs(numberOf(scalar(0)));
    case "MOD": return numberOf(scalar(0)) % numberOf(scalar(1));
    case "ROUND": {
      const digits = Math.trunc(num(1, 0)), factor = Math.pow(10, digits);
      return Math.round((numberOf(scalar(0)) + Number.EPSILON) * factor) / factor;
    }
    case "TODAY": return new Date().toISOString().slice(0, 10);
    case "NOW": return new Date().toISOString().replace("T", " ").slice(0, 19);
    default: throw new FormulaError("#NAME?");
  }
}

function evaluateAst(node: Ast, context: FormulaContext, seen: Set<string>, scope: Map<string, FormulaValue>): FormulaValue {
  switch (node.kind) {
    case "literal": return node.value;
    case "reference": return evaluateReference(node, context, seen, scope);
    case "range": return rangeValue(node.start, node.end, context, seen, scope);
    case "spill": {
      const selected = sheetMap(context, node.reference.sheetName);
      const key = selected.sheetId + "!" + node.reference.address;
      const value = context.spills?.[key] ?? context.spills?.[node.reference.address];
      if (value) return value;
      const resolved = evaluateReference(node.reference, context, seen, scope);
      return isMatrix(resolved) ? resolved : [[resolved]];
    }
    case "name": return namedValue(node.name, context, seen, scope);
    case "unary": {
      const value = first(evaluateAst(node.operand, context, seen, scope));
      if (node.operator === "%") return numberOf(value) / 100;
      return node.operator === "-" ? -numberOf(value) : numberOf(value);
    }
    case "binary": {
      const left = first(evaluateAst(node.left, context, seen, scope));
      const right = first(evaluateAst(node.right, context, seen, scope));
      const leftNumber = typeof left === "number" ? left : Number(left);
      const rightNumber = typeof right === "number" ? right : Number(right);
      switch (node.operator) {
        case "+": return numberOf(left) + numberOf(right);
        case "-": return numberOf(left) - numberOf(right);
        case "*": return numberOf(left) * numberOf(right);
        case "/": if (numberOf(right) === 0) throw new FormulaError("#DIV/0!"); return numberOf(left) / numberOf(right);
        case "^": return Math.pow(numberOf(left), numberOf(right));
        case "&": return String(left) + String(right);
        case "=": return String(left).toLocaleLowerCase() === String(right).toLocaleLowerCase();
        case "<>": return String(left).toLocaleLowerCase() !== String(right).toLocaleLowerCase();
        case ">": return Number.isFinite(leftNumber) && Number.isFinite(rightNumber) ? leftNumber > rightNumber : String(left) > String(right);
        case "<": return Number.isFinite(leftNumber) && Number.isFinite(rightNumber) ? leftNumber < rightNumber : String(left) < String(right);
        case ">=": return Number.isFinite(leftNumber) && Number.isFinite(rightNumber) ? leftNumber >= rightNumber : String(left) >= String(right);
        case "<=": return Number.isFinite(leftNumber) && Number.isFinite(rightNumber) ? leftNumber <= rightNumber : String(left) <= String(right);
        default: throw new FormulaError("#VALUE!");
      }
    }
    case "call": return evaluateCall(node.name, node.args, context, seen, scope);
  }
}

function evaluateInContext(formula: string, context: FormulaContext, seen: Set<string>, scope: Map<string, FormulaValue>): FormulaValue {
  const source = formula.trim().startsWith("=") ? formula.trim().slice(1) : formula.trim();
  return evaluateAst(new Parser(tokenize(source)).parse(), context, seen, scope);
}

export function evaluate(
  formula: string,
  cells: CellMap,
  seen = new Set<string>(),
  contextOverrides: Partial<Omit<FormulaContext, "cells">> = {},
): FormulaValue {
  try {
    return evaluateInContext(formula, { cells, ...contextOverrides }, seen, new Map());
  } catch (error) {
    return error instanceof FormulaError ? error.code : error instanceof Error ? error.message : "#ERROR!";
  }
}

export function displayValue(raw: string, cells: CellMap, showFormula = false): string {
  if (!raw.startsWith("=")) return raw;
  if (showFormula) return raw;
  return String(first(evaluate(raw, cells), ""));
}

export function numericValue(raw: string, cells: CellMap): number | null {
  const value = raw.startsWith("=") ? first(evaluate(raw, cells), "") : raw;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) return Number(value);
  if (typeof value === "boolean") return value ? 1 : 0;
  return null;
}

export function recalculate(cells: CellMap): CellMap {
  return { ...cells };
}

export function formatCellValue(raw: string, cells: CellMap, showFormula: boolean, style?: CellStyle): string {
  if (showFormula && raw.startsWith("=")) return raw;
  const value = first(raw.startsWith("=") ? evaluate(raw, cells) : raw, "");
  if (typeof value === "string" && value.startsWith("#")) return value;
  const n = typeof value === "number" ? value : (String(value).trim() !== "" && Number.isFinite(Number(value)) ? Number(value) : null);
  switch (style?.numberFormat) {
    case "Percentage %":
      return n === null ? String(value) : new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Currency ₹":
      return n === null ? String(value) : new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Currency $":
      return n === null ? String(value) : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 2 }).format(n);
    case "Number":
      return n === null ? String(value) : new Intl.NumberFormat("en-IN", { maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Date": {
      const date = new Date(String(value));
      return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleDateString();
    }
    default:
      return String(value ?? "");
  }
}
