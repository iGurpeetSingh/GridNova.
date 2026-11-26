export type DependencyNode = string;

export interface NamedRangeDefinition {
  sheetId?: string;
  reference: string;
}

export interface DependencySourceSheet {
  id: string;
  name: string;
  cells: Record<string, { raw?: string } | undefined>;
}

export interface DependencyBuildOptions {
  sheetNameToId?: Record<string, string>;
  namedRanges?: Record<string, string | NamedRangeDefinition>;
  maxExpandedRangeCells?: number;
}

export interface DependencyPlan {
  affected: Set<DependencyNode>;
  order: DependencyNode[];
  circular: Set<DependencyNode>;
  blocked: Set<DependencyNode>;
}

const CELL_REFERENCE = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i;

function columnIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function columnName(index: number): string {
  let value = index + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function parseCellReference(value: string): { row: number; column: number; address: string } | null {
  const match = value.replace(/\$/g, '').match(/^([A-Z]{1,3})([1-9]\d*)$/i);
  if (!match) return null;
  const row = Number(match[2]) - 1;
  const column = columnIndex(match[1]);
  return { row, column, address: columnName(column) + String(row + 1) };
}

function splitSheetReference(token: string, currentSheetId: string): { sheetId: string; address: string } | null {
  const bang = token.lastIndexOf('!');
  const sheetPart = bang >= 0 ? token.slice(0, bang) : '';
  const addressPart = bang >= 0 ? token.slice(bang + 1) : token;
  const parsed = parseCellReference(addressPart);
  if (!parsed) return null;
  if (!sheetPart) return { sheetId: currentSheetId, address: parsed.address };
  const unquoted = sheetPart.startsWith("'") && sheetPart.endsWith("'")
    ? sheetPart.slice(1, -1).replace(/''/g, "'")
    : sheetPart;
  return { sheetId: unquoted, address: parsed.address };
}

function removeStringLiterals(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, match => ' '.repeat(match.length));
}

function normalizeSheetId(sheetName: string, options: DependencyBuildOptions): string {
  const cleanName = sheetName.startsWith("'") && sheetName.endsWith("'")
    ? sheetName.slice(1, -1).replace(/''/g, "'")
    : sheetName;
  const direct = options.sheetNameToId?.[cleanName] ?? options.sheetNameToId?.[cleanName.toLowerCase()];
  return direct ?? cleanName;
}

function expandRange(startToken: string, endToken: string, currentSheetId: string, options: DependencyBuildOptions): DependencyNode[] {
  const start = splitSheetReference(startToken, currentSheetId);
  const end = splitSheetReference(endToken, start?.sheetId ?? currentSheetId);
  if (!start || !end) return [];
  const startCell = parseCellReference(start.address);
  const endCell = parseCellReference(end.address);
  if (!startCell || !endCell) return [];
  const rowStart = Math.min(startCell.row, endCell.row);
  const rowEnd = Math.max(startCell.row, endCell.row);
  const colStart = Math.min(startCell.column, endCell.column);
  const colEnd = Math.max(startCell.column, endCell.column);
  const total = (rowEnd - rowStart + 1) * (colEnd - colStart + 1);
  const limit = options.maxExpandedRangeCells ?? 100000;
  if (total > limit) {
    throw new Error('Dependency range expands to ' + total + ' cells; limit is ' + limit + '.');
  }
  const result: DependencyNode[] = [];
  for (let row = rowStart; row <= rowEnd; row++) {
    for (let column = colStart; column <= colEnd; column++) {
      result.push(nodeKey(start.sheetId, columnName(column) + String(row + 1)));
    }
  }
  return result;
}

export function nodeKey(sheetId: string, address: string): DependencyNode {
  return sheetId + '!' + address.replace(/\$/g, '').toUpperCase();
}

export function splitNodeKey(node: DependencyNode): { sheetId: string; address: string } {
  const splitAt = node.lastIndexOf('!');
  return splitAt < 0
    ? { sheetId: '', address: node.toUpperCase() }
    : { sheetId: node.slice(0, splitAt), address: node.slice(splitAt + 1).toUpperCase() };
}

/**
 * Extracts direct cell dependencies from an Excel formula. String literals and
 * function names are ignored. Range references are expanded to their member
 * cells so changing any member invalidates the consuming formula.
 */
export function extractFormulaDependencies(
  formula: string,
  currentSheetId: string,
  options: DependencyBuildOptions = {},
): Set<DependencyNode> {
  const dependencies = new Set<DependencyNode>();
  if (!formula || !formula.trim().startsWith('=')) return dependencies;

  const source = removeStringLiterals(formula.slice(1));
  const referencePattern = /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?\$?[A-Z]{1,3}\$?[1-9]\d*/gi;
  const references: Array<{ token: string; start: number; end: number }> = [];
  let match: RegExpExecArray | null;

  while ((match = referencePattern.exec(source))) {
    references.push({ token: match[0], start: match.index, end: match.index + match[0].length });
  }

  const consumed = new Set<number>();
  for (let index = 0; index < references.length; index++) {
    if (consumed.has(index)) continue;
    const first = references[index];
    let nextIndex = index + 1;
    while (nextIndex < references.length && /\s/.test(source.slice(references[nextIndex - 1].end, references[nextIndex].start))) nextIndex++;

    if (
      nextIndex < references.length &&
      source.slice(first.end, references[nextIndex].start).trim() === ':' &&
      !consumed.has(nextIndex)
    ) {
      const last = references[nextIndex];
      const start = splitSheetReference(first.token, currentSheetId);
      const end = splitSheetReference(last.token, start?.sheetId ?? currentSheetId);
      if (start && end) {
        const normalizedStart = nodeKey(normalizeSheetId(start.sheetId, options), start.address);
        const normalizedEnd = nodeKey(normalizeSheetId(end.sheetId, options), end.address);
        const rangeNodes = expandRange(normalizedStart.slice(normalizedStart.lastIndexOf('!') + 1), normalizedEnd.slice(normalizedEnd.lastIndexOf('!') + 1), normalizeSheetId(start.sheetId, options), options);
        rangeNodes.forEach(node => dependencies.add(node));
      }
      consumed.add(index);
      consumed.add(nextIndex);
      index = nextIndex;
      continue;
    }

    const parsed = splitSheetReference(first.token, currentSheetId);
    if (parsed) dependencies.add(nodeKey(normalizeSheetId(parsed.sheetId, options), parsed.address));
  }

  // A spill reference such as A1# is a dependency on its origin cell. The
  // reference scan above already includes A1, so no separate edge is needed.
  const nameLookup = new Map<string, string | NamedRangeDefinition>();
  Object.entries(options.namedRanges ?? {}).forEach(([name, definition]) => nameLookup.set(name.toLowerCase(), definition));
  const namePattern = /[A-Za-z_][A-Za-z0-9_.]*/g;
  while ((match = namePattern.exec(source))) {
    const name = match[0];
    const lowerName = name.toLowerCase();
    if (!nameLookup.has(lowerName)) continue;
    const after = source.slice(namePattern.lastIndex).match(/^\s*/)?.[0].length ?? 0;
    if (source[namePattern.lastIndex + after] === '(' || source[namePattern.lastIndex + after] === '!') continue;
    const definition = nameLookup.get(lowerName)!;
    const reference = typeof definition === 'string' ? definition : definition.reference;
    const explicitSheet = typeof definition === 'string' ? '' : definition.sheetId ?? '';
    const rangeMatch = reference.match(/^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(\$?[A-Z]{1,3}\$?[1-9]\d*)(?:\s*:\s*(\$?[A-Z]{1,3}\$?[1-9]\d*))?$/i);
    if (!rangeMatch) continue;
    const sheetName = explicitSheet || rangeMatch[1] || rangeMatch[2] || currentSheetId;
    const sheetId = normalizeSheetId(sheetName, options);
    const startAddress = rangeMatch[3].replace(/\$/g, '').toUpperCase();
    const endAddress = (rangeMatch[4] ?? rangeMatch[3]).replace(/\$/g, '').toUpperCase();
    expandRange(startAddress, endAddress, sheetId, options).forEach(node => dependencies.add(node));
  }

  return dependencies;
}

export class DependencyGraph {
  private readonly dependencies = new Map<DependencyNode, Set<DependencyNode>>();
  private readonly dependents = new Map<DependencyNode, Set<DependencyNode>>();

  clear(): void {
    this.dependencies.clear();
    this.dependents.clear();
  }

  setDependencies(node: DependencyNode, sources: Iterable<DependencyNode>): void {
    this.removeDependencies(node);
    const next = new Set(sources);
    this.dependencies.set(node, next);
    for (const source of next) {
      let targets = this.dependents.get(source);
      if (!targets) {
        targets = new Set<DependencyNode>();
        this.dependents.set(source, targets);
      }
      targets.add(node);
    }
  }

  removeDependencies(node: DependencyNode): void {
    const previous = this.dependencies.get(node);
    if (previous) {
      for (const source of previous) {
        const targets = this.dependents.get(source);
        targets?.delete(node);
        if (targets?.size === 0) this.dependents.delete(source);
      }
    }
    this.dependencies.delete(node);
  }

  removeNode(node: DependencyNode): void {
    this.removeDependencies(node);
    const consumers = this.dependents.get(node);
    if (consumers) {
      for (const consumer of consumers) this.dependencies.get(consumer)?.delete(node);
    }
    this.dependents.delete(node);
  }

  getPrecedents(node: DependencyNode): Set<DependencyNode> {
    return new Set(this.dependencies.get(node) ?? []);
  }

  getDependents(node: DependencyNode): Set<DependencyNode> {
    return new Set(this.dependents.get(node) ?? []);
  }

  getDownstream(nodes: Iterable<DependencyNode>): Set<DependencyNode> {
    const affected = new Set<DependencyNode>();
    const queue = Array.from(nodes);
    while (queue.length) {
      const current = queue.shift()!;
      for (const dependent of this.dependents.get(current) ?? []) {
        if (affected.has(dependent)) continue;
        affected.add(dependent);
        queue.push(dependent);
      }
    }
    return affected;
  }

  /**
   * Produces a topological recalculation plan for changed cells. Formula cells
   * are ordered source-first. Strongly connected components identify actual
   * circular members; nodes blocked behind a cycle are returned separately.
   */
  plan(changedNodes: Iterable<DependencyNode>, includeChangedFormulas = true): DependencyPlan {
    const roots = new Set(changedNodes);
    const affected = this.getDownstream(roots);
    if (includeChangedFormulas) {
      for (const root of roots) if (this.dependencies.has(root)) affected.add(root);
    }

    const indegree = new Map<DependencyNode, number>();
    for (const node of affected) {
      let count = 0;
      for (const source of this.dependencies.get(node) ?? []) if (affected.has(source)) count++;
      indegree.set(node, count);
    }

    const queue = Array.from(affected).filter(node => (indegree.get(node) ?? 0) === 0).sort();
    const order: DependencyNode[] = [];
    while (queue.length) {
      const node = queue.shift()!;
      order.push(node);
      for (const target of this.dependents.get(node) ?? []) {
        if (!affected.has(target)) continue;
        const count = (indegree.get(target) ?? 0) - 1;
        indegree.set(target, count);
        if (count === 0) {
          queue.push(target);
          queue.sort();
        }
      }
    }

    const unresolved = new Set(Array.from(affected).filter(node => !order.includes(node)));
    const circular = this.findCircularMembers(unresolved);
    const blocked = new Set(Array.from(unresolved).filter(node => !circular.has(node)));
    return { affected, order, circular, blocked };
  }

  private findCircularMembers(candidates: Set<DependencyNode>): Set<DependencyNode> {
    let nextIndex = 0;
    const indexByNode = new Map<DependencyNode, number>();
    const lowLink = new Map<DependencyNode, number>();
    const stack: DependencyNode[] = [];
    const onStack = new Set<DependencyNode>();
    const circular = new Set<DependencyNode>();

    const visit = (node: DependencyNode) => {
      indexByNode.set(node, nextIndex);
      lowLink.set(node, nextIndex);
      nextIndex++;
      stack.push(node);
      onStack.add(node);

      for (const target of this.dependents.get(node) ?? []) {
        if (!candidates.has(target)) continue;
        if (!indexByNode.has(target)) {
          visit(target);
          lowLink.set(node, Math.min(lowLink.get(node)!, lowLink.get(target)!));
        } else if (onStack.has(target)) {
          lowLink.set(node, Math.min(lowLink.get(node)!, indexByNode.get(target)!));
        }
      }

      if (lowLink.get(node) !== indexByNode.get(node)) return;
      const component: DependencyNode[] = [];
      let member: DependencyNode;
      do {
        member = stack.pop()!;
        onStack.delete(member);
        component.push(member);
      } while (member !== node);

      const selfCycle = component.length === 1 && (this.dependencies.get(node)?.has(node) ?? false);
      if (component.length > 1 || selfCycle) component.forEach(item => circular.add(item));
    };

    for (const node of candidates) if (!indexByNode.has(node)) visit(node);
    return circular;
  }

  tracePrecedents(node: DependencyNode): Set<DependencyNode> {
    const result = new Set<DependencyNode>();
    const queue = [node];
    while (queue.length) {
      const current = queue.shift()!;
      for (const source of this.dependencies.get(current) ?? []) {
        if (result.has(source)) continue;
        result.add(source);
        queue.push(source);
      }
    }
    return result;
  }

  traceDependents(node: DependencyNode): Set<DependencyNode> {
    const result = new Set<DependencyNode>();
    const queue = [node];
    while (queue.length) {
      const current = queue.shift()!;
      for (const target of this.dependents.get(current) ?? []) {
        if (result.has(target)) continue;
        result.add(target);
        queue.push(target);
      }
    }
    return result;
  }

  static fromSheets(sheets: DependencySourceSheet[], options: DependencyBuildOptions = {}): DependencyGraph {
    const graph = new DependencyGraph();
    const sheetNameToId = { ...(options.sheetNameToId ?? {}) };
    for (const sheet of sheets) {
      sheetNameToId[sheet.name] = sheet.id;
      sheetNameToId[sheet.name.toLowerCase()] = sheet.id;
    }
    const resolvedOptions = { ...options, sheetNameToId };
    for (const sheet of sheets) {
      for (const [address, cell] of Object.entries(sheet.cells)) {
        if (!cell?.raw?.trim().startsWith('=')) continue;
        const node = nodeKey(sheet.id, address);
        graph.setDependencies(node, extractFormulaDependencies(cell.raw, sheet.id, resolvedOptions));
      }
    }
    return graph;
  }
}
