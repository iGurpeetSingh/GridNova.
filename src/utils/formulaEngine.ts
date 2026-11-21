import type { CellMap } from "../types";
import { addressOf, parseAddress } from "../types";

type Value = string | number | boolean;
const refPattern = /^[A-Z]{1,3}[1-9]\\d*$/;

function rangeValues(start: string, end: string, cells: CellMap, seen: Set<string>): Value[] {
  const a = parseAddress(start), b = parseAddress(end);
  if (!a || !b) throw new Error("#REF!");
  const values: Value[] = [];
  for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++)
    for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++)
      values.push(resolve(addressOf(r, c), cells, seen));
  return values;
}
function resolve(address: string, cells: CellMap, seen: Set<string>): Value {
  const key = address.toUpperCase(), cell = cells[key];
  if (!cell) return 0;
  if (seen.has(key)) throw new Error("#CIRCULAR!");
  if (cell.raw.startsWith("=")) {
    const next = new Set(seen); next.add(key);
    return evaluate(cell.raw, cells, next);
  }
  const n = Number(cell.raw);
  return cell.raw.trim() !== "" && Number.isFinite(n) ? n : cell.raw;
}

function tokenize(source: string): string[] {
  const result: string[] = [];
  const re = /\\s*(>=|<=|<>|[()+*/^=<>,-]|"(?:[^"]|"")*"|(?:\\d+\\.?\\d*|\\.\\d+)|[A-Za-z_][A-Za-z0-9_]*)/gy;
  let pos = 0;
  while (pos < source.length) {
    re.lastIndex = pos;
    const match = re.exec(source);
    if (!match) throw new Error("#PARSE!");
    result.push(match[1]); pos = re.lastIndex;
  }
  return result;
}
class Parser {
  private i = 0;
  constructor(private t: string[], private cells: CellMap, private seen: Set<string>) {}
  private peek() { return this.t[this.i]; }
  private take() { return this.t[this.i++]; }
  private match(v: string) { if (this.peek() === v) { this.i++; return true; } return false; }
  parse(): Value {
    const v = this.compare();
    if (this.i !== this.t.length) throw new Error("#PARSE!");
    return v;
  }
  private compare(): Value {
    let a = this.add();
    while ([">", "<", ">=", "<=", "=", "<>"].includes(this.peek())) {
      const op = this.take(), b = this.add();
      const x = typeof a === "number" ? a : String(a), y = typeof b === "number" ? b : String(b);
      switch (op) { case ">": a = x > y; break; case "<": a = x < y; break; case ">=": a = x >= y; break; case "<=": a = x <= y; break; case "=": a = x === y; break; default: a = x !== y; }
    }
    return a;
  }
  private add(): Value {
    let a = this.mul();
    while (this.peek() === "+" || this.peek() === "-") {
      const op = this.take(), b = this.mul();
      a = op === "+" ? (typeof a === "string" || typeof b === "string" ? String(a) + String(b) : Number(a) + Number(b)) : Number(a) - Number(b);
    }
    return a;
  }
  private mul(): Value {
    let a = this.power();
    while (this.peek() === "*" || this.peek() === "/") {
      const op = this.take(), b = this.power();
      if (op === "/" && Number(b) === 0) throw new Error("#DIV/0!");
      a = op === "*" ? Number(a) * Number(b) : Number(a) / Number(b);
    }
    return a;
  }
  private power(): Value {
    let a = this.unary();
    while (this.match("^")) a = Math.pow(Number(a), Number(this.unary()));
    return a;
  }
  private unary(): Value {
    if (this.match("-")) return -Number(this.unary());
    if (this.match("+")) return Number(this.unary());
    return this.primary();
  }
  private argument(): Value[] {
    const first = this.peek(), second = this.t[this.i + 1];
    if (first && refPattern.test(first.toUpperCase()) && second === ":") {
      const start = this.take().toUpperCase(); this.take();
      const end = this.take()?.toUpperCase();
      if (!end || !refPattern.test(end)) throw new Error("#REF!");
      return rangeValues(start, end, this.cells, this.seen);
    }
    return [this.compare()];
  }
  private primary(): Value {
    const token = this.take();
    if (!token) throw new Error("#PARSE!");
    if (token === "(") { const v = this.compare(); if (!this.match(")")) throw new Error("#PARSE!"); return v; }
    if (/^"(?:[^"]|"")*"$/.test(token)) return token.slice(1, -1).replace(/""/g, '"');
    if (/^(?:\\d+\\.?\\d*|\\.\\d+)$/.test(token)) return Number(token);
    if (token.toUpperCase() === "TRUE") return true;
    if (token.toUpperCase() === "FALSE") return false;
    if (this.match("(")) {
      const args: Value[][] = [];
      if (!this.match(")")) {
        do { args.push(this.argument()); } while (this.match(","));
        if (!this.match(")")) throw new Error("#PARSE!");
      }
      const all = args.flat(), nums = all.filter(v => typeof v === "number") as number[];
      switch (token.toUpperCase()) {
        case "SUM": return nums.reduce((s, n) => s + n, 0);
        case "AVERAGE": return nums.length ? nums.reduce((s, n) => s + n, 0) / nums.length : 0;
        case "COUNT": return nums.length;
        case "MIN": return nums.length ? Math.min(...nums) : 0;
        case "MAX": return nums.length ? Math.max(...nums) : 0;
        case "IF": return args[0]?.[0] ? args[1]?.[0] ?? false : args[2]?.[0] ?? false;
        default: throw new Error("#NAME?");
      }
    }
    if (refPattern.test(token.toUpperCase())) return resolve(token.toUpperCase(), this.cells, this.seen);
    throw new Error("#NAME?");
  }
}
export function evaluate(formula: string, cells: CellMap, seen = new Set<string>()): Value {
  try { return new Parser(tokenize(formula.startsWith("=") ? formula.slice(1) : formula), cells, seen).parse(); }
  catch (e) { return e instanceof Error ? e.message : "#ERROR!"; }
}
export function displayValue(raw: string, cells: CellMap, showFormula = false): string {
  if (!raw.startsWith("=")) return raw;
  return showFormula ? raw : String(evaluate(raw, cells));
}
export function numericValue(raw: string, cells: CellMap): number | null {
  const v = raw.startsWith("=") ? evaluate(raw, cells) : raw;
  if (typeof v === "number") return v;
  if (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) return Number(v);
  return null;
}
export function recalculate(cells: CellMap): CellMap { return { ...cells }; }