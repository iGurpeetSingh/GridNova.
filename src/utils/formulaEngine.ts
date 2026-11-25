import type { CellMap, CellStyle, ExcelTableData } from "../types";
import { addressOf, parseAddress } from "../types";
import { getTableAt, getTableColumnIndex, tableReferenceRange } from "./tableUtils";

type Value = string | number | boolean;
const refPattern = /^[A-Z]{1,3}[1-9]\d*$/;
const tableNamePattern = /^[A-Za-z_][A-Za-z0-9_.]*/;

function resolveStructuredReference(name: string | null, inner: string, cells: CellMap, tables: ExcelTableData[], currentAddress?: string): string {
  const current = currentAddress ? parseAddress(currentAddress) : null;
  const table = name
    ? tables.find(item => item.name.toLowerCase() === name.toLowerCase())
    : current ? getTableAt(tables, current.row, current.col) : undefined;
  if (!table) throw new Error("#NAME?");
  const terms = [...inner.matchAll(/\[([^\]]+)\]/g)].map(match => match[1].trim());
  const selector = terms.find(term => term.startsWith("#")) ?? (inner.trim().startsWith("#") ? inner.trim() : undefined);
  const thisRow = inner.trim().startsWith("@") || selector?.toLowerCase() === "#this row";
  const columnName = terms.filter(term => !term.startsWith("#")).at(-1) ?? (terms.length || inner.trim().startsWith("#") ? "" : inner.replace(/^@/, "").trim());
  if (thisRow) {
    const firstDataRow = table.range.startRow + (table.hasHeaders ? 1 : 0);
    if (!current || current.row < firstDataRow || current.row > table.range.endRow) throw new Error("#VALUE!");
    if (!columnName) return addressOf(current.row, current.col);
    const column = getTableColumnIndex(table, columnName, cells);
    if (column < 0) throw new Error("#REF!");
    return addressOf(current.row, column);
  }
  if (!columnName) {
    const range = tableReferenceRange(table, selector);
    return `${addressOf(range.startRow, range.startCol)}:${addressOf(range.endRow, range.endCol)}`;
  }
  const column = getTableColumnIndex(table, columnName, cells);
  if (column < 0) throw new Error("#REF!");
  const range = tableReferenceRange(table, selector ?? "#Data");
  return `${addressOf(range.startRow, column)}:${addressOf(range.endRow, column)}`;
}

function expandStructuredReferences(source: string, cells: CellMap, tables: ExcelTableData[], currentAddress?: string): string {
  let output = "", i = 0;
  while (i < source.length) {
    // Current-row shorthand: [@Amount] or [@[Amount]]
    if (source[i] === "[" && source[i + 1] === "@") {
      let depth = 0, end = i;
      for (; end < source.length; end++) {
        if (source[end] === "[") depth++;
        else if (source[end] === "]" && --depth === 0) break;
      }
      if (end < source.length) {
        output += resolveStructuredReference(null, source.slice(i + 1, end), cells, tables, currentAddress);
        i = end + 1; continue;
      }
    }
    const match = source.slice(i).match(tableNamePattern);
    if (match) {
      const name = match[0], bracketIndex = i + name.length;
      if (source[bracketIndex] === "[") {
        let depth = 0, end = bracketIndex;
        for (; end < source.length; end++) {
          if (source[end] === "[") depth++;
          else if (source[end] === "]") { depth--; if (depth === 0) break; }
        }
        if (end < source.length) {
          output += resolveStructuredReference(name, source.slice(bracketIndex + 1, end), cells, tables, currentAddress);
          i = end + 1; continue;
        }
      }
      output += name; i += name.length; continue;
    }
    output += source[i]; i++;
  }
  return output;
}

function rangeValues(start: string, end: string, cells: CellMap, seen: Set<string>, tables: ExcelTableData[], currentAddress?: string): Value[] {
  const a = parseAddress(start), b = parseAddress(end);
  if (!a || !b) throw new Error("#REF!");
  const values: Value[] = [];
  for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
    for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
      values.push(resolve(addressOf(r, c), cells, seen, tables));
    }
  }
  return values;
}

function resolve(address: string, cells: CellMap, seen: Set<string>, tables: ExcelTableData[]): Value {
  const key = address.toUpperCase(), cell = cells[key];
  if (!cell) return "";
  if (seen.has(key)) throw new Error("#CIRCULAR!");
  if (cell.raw.startsWith("=")) {
    const next = new Set(seen); next.add(key);
    return evaluate(cell.raw, cells, next, tables, key);
  }
  const n = Number(cell.raw);
  return cell.raw.trim() !== "" && Number.isFinite(n) ? n : cell.raw;
}

function tokenize(source: string): string[] {
  const result: string[] = [];
  const re = /\s*(>=|<=|<>|[()+*/^=<>,:-]|"(?:[^"]|"")*"|(?:\d+\.?\d*|\.\d+)|[A-Za-z_][A-Za-z0-9_.]*)/gy;
  let pos = 0;
  while (pos < source.length) {
    re.lastIndex = pos;
    const match = re.exec(source);
    if (!match) throw new Error("#PARSE!");
    result.push(match[1]); pos = re.lastIndex;
  }
  return result;
}

const isError = (value: Value | undefined) => typeof value === "string" && value.startsWith("#");
const number = (value: Value | undefined) => value === undefined || value === "" ? 0 : Number(value);
const numeric = (values: Value[]) => values.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
const text = (value: Value | undefined) => value === undefined ? "" : String(value);
const criteriaMatches = (value: Value, criteria: Value): boolean => {
  const condition = String(criteria), operator = condition.match(/^(>=|<=|<>|>|<|=)/)?.[0] ?? "=";
  const expected = condition.slice(operator === "=" && !/^(>=|<=|<>|>|<|=)/.test(condition) ? 0 : operator.length);
  const leftNumber = Number(value), rightNumber = Number(expected);
  const bothNumeric = String(value).trim() !== "" && expected.trim() !== "" && Number.isFinite(leftNumber) && Number.isFinite(rightNumber);
  const left = bothNumeric ? leftNumber : String(value).toLowerCase(), right = bothNumeric ? rightNumber : expected.toLowerCase();
  if (operator === ">=") return left >= right;
  if (operator === "<=") return left <= right;
  if (operator === ">") return left > right;
  if (operator === "<") return left < right;
  if (operator === "<>") return left !== right;
  return left === right || String(value).toLowerCase() === expected.toLowerCase();
};

class Parser {
  private i = 0;
  constructor(private readonly tokens: string[], private readonly cells: CellMap, private readonly seen: Set<string>, private readonly tables: ExcelTableData[]) {}
  private peek(): string | undefined { return this.tokens[this.i]; }
  private take(): string | undefined { return this.tokens[this.i++]; }
  private match(v: string): boolean { if (this.peek() === v) { this.i++; return true; } return false; }
  parse(): Value { const value = this.compare(); if (this.i !== this.tokens.length) throw new Error("#PARSE!"); return value; }
  private compare(): Value {
    let a = this.add();
    while ([">", "<", ">=", "<=", "=", "<>"].includes(this.peek() ?? "")) {
      const op = this.take()!, b = this.add();
      const x = typeof a === "number" ? a : String(a), y = typeof b === "number" ? b : String(b);
      switch (op) { case ">": a = x > y; break; case "<": a = x < y; break; case ">=": a = x >= y; break; case "<=": a = x <= y; break; case "=": a = x === y; break; default: a = x !== y; }
    }
    return a;
  }
  private add(): Value {
    let a = this.multiply();
    while (this.peek() === "+" || this.peek() === "-") {
      const op = this.take(), b = this.multiply();
      a = op === "+" ? (typeof a === "string" || typeof b === "string" ? String(a) + String(b) : number(a) + number(b)) : number(a) - number(b);
    }
    return a;
  }
  private multiply(): Value {
    let a = this.power();
    while (this.peek() === "*" || this.peek() === "/") {
      const op = this.take(), b = this.power();
      if (op === "/" && number(b) === 0) throw new Error("#DIV/0!");
      a = op === "*" ? number(a) * number(b) : number(a) / number(b);
    }
    return a;
  }
  private power(): Value { let a = this.unary(); while (this.match("^")) a = Math.pow(number(a), number(this.unary())); return a; }
  private unary(): Value { if (this.match("-")) return -number(this.unary()); if (this.match("+")) return number(this.unary()); return this.primary(); }
  private argument(): Value[] {
    const first = this.peek(), second = this.tokens[this.i + 1];
    if (first && refPattern.test(first.toUpperCase()) && second === ":") {
      const start = this.take()!.toUpperCase(); this.take(); const end = this.take()?.toUpperCase();
      if (!end || !refPattern.test(end)) throw new Error("#REF!");
      return rangeValues(start, end, this.cells, this.seen, this.tables);
    }
    return [this.compare()];
  }
  private primary(): Value {
    const token = this.take();
    if (!token) throw new Error("#PARSE!");
    if (token === "(") { const value = this.compare(); if (!this.match(")")) throw new Error("#PARSE!"); return value; }
    if (/^"(?:[^"]|"")*"$/.test(token)) return token.slice(1, -1).replace(/""/g, '"');
    if (/^(?:\d+\.?\d*|\.\d+)$/.test(token)) return Number(token);
    if (token.toUpperCase() === "TRUE") return true;
    if (token.toUpperCase() === "FALSE") return false;
    if (this.match("(")) {
      const args: Value[][] = [];
      if (!this.match(")")) { do { args.push(this.argument()); } while (this.match(",")); if (!this.match(")")) throw new Error("#PARSE!"); }
      return this.call(token.toUpperCase(), args);
    }
    if (refPattern.test(token.toUpperCase())) return resolve(token.toUpperCase(), this.cells, this.seen, this.tables);
    throw new Error("#NAME?");
  }
  private call(fn: string, args: Value[][]): Value {
    const flat = args.flat(), nums = numeric(flat), a = args.map(item => item[0]);
    const sum = (values: number[]) => values.reduce((total, value) => total + value, 0);
    switch (fn) {
      case "SUM": return sum(nums);
      case "AVERAGE": return nums.length ? sum(nums) / nums.length : "#DIV/0!";
      case "COUNT": return nums.length;
      case "COUNTA": return flat.filter(value => value !== "").length;
      case "COUNTBLANK": return flat.filter(value => value === "").length;
      case "MIN": return nums.length ? Math.min(...nums) : 0;
      case "MAX": return nums.length ? Math.max(...nums) : 0;
      case "PRODUCT": return nums.reduce((total, value) => total * value, 1);
      case "ABS": return Math.abs(number(a[0]));
      case "ROUND": return roundTo(number(a[0]), number(a[1]));
      case "ROUNDUP": return roundDirected(number(a[0]), number(a[1]), "up");
      case "ROUNDDOWN": return roundDirected(number(a[0]), number(a[1]), "down");
      case "MOD": if (number(a[1]) === 0) return "#DIV/0!"; return number(a[0]) % number(a[1]);
      case "POWER": return Math.pow(number(a[0]), number(a[1]));
      case "SQRT": return number(a[0]) < 0 ? "#NUM!" : Math.sqrt(number(a[0]));
      case "IF": return a[0] ? a[1] ?? false : a[2] ?? false;
      case "IFERROR": return isError(a[0]) ? a[1] ?? "" : a[0] ?? "";
      case "AND": return flat.every(Boolean);
      case "OR": return flat.some(Boolean);
      case "NOT": return !Boolean(a[0]);
      case "ISBLANK": return a[0] === "";
      case "ISNUMBER": return typeof a[0] === "number" && Number.isFinite(a[0]);
      case "ISTEXT": return typeof a[0] === "string" && !isError(a[0]);
      case "ISERROR": return isError(a[0]);
      case "LEFT": return text(a[0]).slice(0, Math.max(0, number(a[1] ?? 1)));
      case "RIGHT": return text(a[0]).slice(-Math.max(0, number(a[1] ?? 1)));
      case "MID": return text(a[0]).slice(Math.max(0, number(a[1]) - 1), Math.max(0, number(a[1]) - 1) + Math.max(0, number(a[2])));
      case "LEN": return text(a[0]).length;
      case "TRIM": return text(a[0]).trim().replace(/\s+/g, " ");
      case "UPPER": return text(a[0]).toUpperCase();
      case "LOWER": return text(a[0]).toLowerCase();
      case "PROPER": return text(a[0]).toLowerCase().replace(/\b([a-z])/g, letter => letter.toUpperCase());
      case "CONCAT": case "CONCATENATE": return flat.map(value => String(value)).join("");
      case "VALUE": { const parsed = Number(text(a[0]).replace(/,/g, "")); return Number.isFinite(parsed) ? parsed : "#VALUE!"; }
      case "DATE": { const year = number(a[0]), month = number(a[1]), day = number(a[2]); const date = new Date(Date.UTC(year, month - 1, day)); return date.toISOString().slice(0, 10); }
      case "TODAY": return new Date().toISOString().slice(0, 10);
      case "NOW": return new Date().toISOString().replace("T", " ").slice(0, 19);
      case "YEAR": case "MONTH": case "DAY": {
        const date = new Date(text(a[0]));
        if (Number.isNaN(date.getTime())) return "#VALUE!";
        return fn === "YEAR" ? date.getUTCFullYear() : fn === "MONTH" ? date.getUTCMonth() + 1 : date.getUTCDate();
      }
      case "SUMIF": case "COUNTIF": case "AVERAGEIF": {
        const range = args[0] ?? [], criteria = a[1] ?? "", sumRange = args[2] ?? range;
        const matched = range.map((value, index) => ({ value, sum: sumRange[index] ?? value })).filter(item => criteriaMatches(item.value, criteria));
        if (fn === "COUNTIF") return matched.length;
        const values = numeric(matched.map(item => item.sum));
        return fn === "SUMIF" ? sum(values) : values.length ? sum(values) / values.length : "#DIV/0!";
      }
      case "SUMIFS": case "COUNTIFS": {
        const targetRange = args[0] ?? [], pairs: { range: Value[]; criteria: Value }[] = [];
        const offset = fn === "SUMIFS" ? 1 : 0;
        for (let i = offset; i + 1 < args.length; i += 2) pairs.push({ range: args[i], criteria: a[i + 1] ?? "" });
        const matches = targetRange.filter((_, index) => pairs.every(pair => criteriaMatches(pair.range[index] ?? "", pair.criteria)));
        return fn === "COUNTIFS" ? matches.length : sum(numeric(matches));
      }
      case "SUBTOTAL": {
        const code = number(a[0]), values = numeric(args.slice(1).flat());
        if ([1,101].includes(code)) return values.length ? sum(values) / values.length : 0;
        if ([2,102].includes(code)) return values.length;
        if ([3,103].includes(code)) return flat.length;
        if ([4,104].includes(code)) return values.length ? Math.max(...values) : 0;
        if ([5,105].includes(code)) return values.length ? Math.min(...values) : 0;
        if ([9,109].includes(code)) return sum(values);
        return "#VALUE!";
      }
      default: throw new Error("#NAME?");
    }
  }
}

function roundTo(value: number, digits: number): number { const factor = Math.pow(10, digits); return Math.round((value + Number.EPSILON) * factor) / factor; }
function roundDirected(value: number, digits: number, direction: "up" | "down"): number {
  const factor = Math.pow(10, digits), scaled = value * factor;
  return (direction === "up" ? (scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled)) : (scaled >= 0 ? Math.floor(scaled) : Math.ceil(scaled))) / factor;
}

export function evaluate(formula: string, cells: CellMap, seen = new Set<string>(), tables: ExcelTableData[] = [], currentAddress?: string): Value {
  try {
    const source = formula.startsWith("=") ? formula.slice(1) : formula;
    const expanded = expandStructuredReferences(source, cells, tables, currentAddress);
    return new Parser(tokenize(expanded), cells, seen, tables).parse();
  } catch (error) {
    return error instanceof Error ? error.message : "#ERROR!";
  }
}
export function displayValue(raw: string, cells: CellMap, showFormula = false, tables: ExcelTableData[] = [], currentAddress?: string): string {
  if (!raw.startsWith("=")) return raw;
  return showFormula ? raw : String(evaluate(raw, cells, new Set(), tables, currentAddress));
}
export function numericValue(raw: string, cells: CellMap, tables: ExcelTableData[] = [], currentAddress?: string): number | null {
  const value = raw.startsWith("=") ? evaluate(raw, cells, new Set(), tables, currentAddress) : raw;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) return Number(value);
  return null;
}
export function recalculate(cells: CellMap): CellMap { return { ...cells }; }
export function formatCellValue(raw: string, cells: CellMap, showFormula: boolean, style?: CellStyle, tables: ExcelTableData[] = [], currentAddress?: string): string {
  if (showFormula && raw.startsWith("=")) return raw;
  const value = raw.startsWith("=") ? evaluate(raw, cells, new Set(), tables, currentAddress) : raw;
  if (typeof value === "string" && value.startsWith("#")) return value;
  const n = typeof value === "number" ? value : (String(value).trim() !== "" && Number.isFinite(Number(value)) ? Number(value) : null);
  switch (style?.numberFormat) {
    case "Percentage %": return n === null ? String(value) : new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Currency ₹": return n === null ? String(value) : new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Currency $": return n === null ? String(value) : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Number": return n === null ? String(value) : new Intl.NumberFormat("en-IN", { maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Date": {
      const date = new Date(String(value));
      return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleDateString();
    }
    default: return String(value ?? "");
  }
}
