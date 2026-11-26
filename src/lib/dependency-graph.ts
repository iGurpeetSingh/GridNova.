import type { WorkbookData, SheetData, CellRange } from "../types";
import { addressOf, parseAddress } from "../types";
import { evaluate } from "../utils/formulaEngine";

export type CellNodeId = `${string}!${string}`;
export interface FormulaNode { id:CellNodeId; sheetId:string; sheetName:string; address:string; formula:string; }
export interface CellDependency { source:CellNodeId; target:CellNodeId; }
export interface RangeDependency { sheetId:string; range:CellRange; target:CellNodeId; }
export interface DependencyGraph {
  formulas:Map<CellNodeId,FormulaNode>;
  dependents:Map<CellNodeId,Set<CellNodeId>>;
  prerequisites:Map<CellNodeId,Set<CellNodeId>>;
  ranges:RangeDependency[];
}
export interface IncrementalResult { workbook:WorkbookData; graph:DependencyGraph; recalculated:CellNodeId[]; circularReferences:CellNodeId[]; }
const cleanAddress=(value:string)=>value.replace(/\$/g,"").toUpperCase();
export const cellNodeId=(sheetId:string,address:string):CellNodeId=>`${sheetId}!${cleanAddress(address)}`;
const nodeParts=(id:CellNodeId)=>{const i=id.lastIndexOf("!");return {sheetId:id.slice(0,i),address:id.slice(i+1)};};
const unquote=(value:string)=>value.replace(/^'/,"").replace(/'$/,"").replace(/''/g,"'");
const escapeRegExp=(value:string)=>value.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
function expandNames(formula:string,namedRanges:Record<string,string>={}):string{
  let value=formula.startsWith("=")?formula.slice(1):formula;
  value=value.replace(/"(?:[^"]|"")*"/g," ");
  for(const [name,reference] of Object.entries(namedRanges)){
    const match=new RegExp(`(^|[^A-Za-z0-9_.])${escapeRegExp(name)}(?=$|[^A-Za-z0-9_.])`,"gi");
    value=value.replace(match,(_whole,prefix:string)=>prefix+reference.replace(/^=/,""));
  }
  return value;
}
interface Reference { sheetName?:string; start:{row:number;col:number}; end:{row:number;col:number}; }
export function extractFormulaReferences(formula:string,namedRanges:Record<string,string>={}):Reference[]{
  const source=expandNames(formula,namedRanges);
  const pattern=/(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Z]{1,3}\$?[1-9]\d*))?/gi;
  const references:Reference[]=[];
  for(const match of source.matchAll(pattern)){
    const startIndex=match.index??0,endIndex=startIndex+match[0].length;
    if(source.slice(endIndex).trimStart().startsWith("("))continue;
    const start=parseAddress(cleanAddress(match[3]??"")),end=parseAddress(cleanAddress(match[4]??match[3]??""));
    if(!start||!end)continue;
    references.push({
      sheetName:match[1]?unquote(match[1]):match[2],
      start:{row:Math.min(start.row,end.row),col:Math.min(start.col,end.col)},
      end:{row:Math.max(start.row,end.row),col:Math.max(start.col,end.col)}
    });
  }
  return references;
}
const addEdge=(graph:DependencyGraph,source:CellNodeId,target:CellNodeId)=>{
  let outgoing=graph.dependents.get(source);if(!outgoing){outgoing=new Set();graph.dependents.set(source,outgoing);}outgoing.add(target);
  let incoming=graph.prerequisites.get(target);if(!incoming){incoming=new Set();graph.prerequisites.set(target,incoming);}incoming.add(source);
};
export function buildDependencyGraph(workbook:WorkbookData):DependencyGraph{
  const graph:DependencyGraph={formulas:new Map(),dependents:new Map(),prerequisites:new Map(),ranges:[]};
  const sheets=workbook.sheets;
  for(const sheet of sheets)for(const [address,cell] of Object.entries(sheet.cells)){
    if(!cell.raw?.startsWith("="))continue;
    const normalized=cleanAddress(address);
    const id=cellNodeId(sheet.id,normalized);
    graph.formulas.set(id,{id,sheetId:sheet.id,sheetName:sheet.name,address:normalized,formula:cell.raw});
  }
  for(const node of graph.formulas.values()){
    const currentSheet=sheets.find(sheet=>sheet.id===node.sheetId)!;
    for(const reference of extractFormulaReferences(node.formula,workbook.namedRanges??{})){
      const sourceSheet=reference.sheetName?sheets.find(sheet=>sheet.name.toLowerCase()===reference.sheetName!.toLowerCase()):currentSheet;
      if(!sourceSheet)continue;
      const range:CellRange={startRow:reference.start.row,startCol:reference.start.col,endRow:reference.end.row,endCol:reference.end.col};
      const isRange=range.startRow!==range.endRow||range.startCol!==range.endCol;
      if(isRange)graph.ranges.push({sheetId:sourceSheet.id,range,target:node.id});
      else addEdge(graph,cellNodeId(sourceSheet.id,addressOf(range.startRow,range.startCol)),node.id);
      // Formula-to-formula edges are needed for exact topological ordering and cycle detection.
      for(const sourceNode of graph.formulas.values()){
        if(sourceNode.sheetId!==sourceSheet.id)continue;
        const point=parseAddress(sourceNode.address);if(!point)continue;
        if(point.row>=range.startRow&&point.row<=range.endRow&&point.col>=range.startCol&&point.col<=range.endCol)addEdge(graph,sourceNode.id,node.id);
      }
    }
  }
  return graph;
}
export function getDependents(graph:DependencyGraph,id:CellNodeId):CellNodeId[]{
  const found=new Set<CellNodeId>(),queue=[id];
  while(queue.length){const current=queue.shift()!;for(const next of graph.dependents.get(current)??[]){if(found.has(next))continue;found.add(next);queue.push(next);}}
  return [...found];
}
export function getPrecedents(graph:DependencyGraph,id:CellNodeId):CellNodeId[]{
  return [...(graph.prerequisites.get(id)??[])];
}
export function findCircularReferences(graph:DependencyGraph):CellNodeId[]{
  let sequence=0;
  const indices=new Map<CellNodeId,number>(),low=new Map<CellNodeId,number>(),stack:CellNodeId[]=[],onStack=new Set<CellNodeId>(),cycles=new Set<CellNodeId>();
  const visit=(id:CellNodeId)=>{
    indices.set(id,sequence);low.set(id,sequence++);stack.push(id);onStack.add(id);
    for(const next of graph.dependents.get(id)??[]){
      if(!graph.formulas.has(next))continue;
      if(!indices.has(next)){visit(next);low.set(id,Math.min(low.get(id)!,low.get(next)!));}
      else if(onStack.has(next))low.set(id,Math.min(low.get(id)!,indices.get(next)!));
    }
    if(low.get(id)===indices.get(id)){
      const component:CellNodeId[]=[];let member:CellNodeId;
      do{member=stack.pop()!;onStack.delete(member);component.push(member);}while(member!==id);
      const selfCycle=component.length===1&&(graph.dependents.get(id)?.has(id)??false);
      if(component.length>1||selfCycle)component.forEach(item=>cycles.add(item));
    }
  };
  for(const id of graph.formulas.keys())if(!indices.has(id))visit(id);
  return [...cycles];
}
function rangeContains(range:CellRange,address:string):boolean{
  const point=parseAddress(address);return !!point&&point.row>=range.startRow&&point.row<=range.endRow&&point.col>=range.startCol&&point.col<=range.endCol;
}
function topologicalOrder(graph:DependencyGraph,affected:Set<CellNodeId>,cycles:Set<CellNodeId>):CellNodeId[]{
  const indegree=new Map<CellNodeId,number>();
  for(const id of affected)if(!cycles.has(id))indegree.set(id,0);
  for(const id of indegree.keys())for(const source of graph.prerequisites.get(id)??[])if(indegree.has(source))indegree.set(id,indegree.get(id)!+1);
  const queue=[...indegree].filter(([,degree])=>degree===0).map(([id])=>id),ordered:CellNodeId[]=[];
  while(queue.length){
    const id=queue.shift()!;ordered.push(id);
    for(const next of graph.dependents.get(id)??[]){
      if(!indegree.has(next))continue;
      const degree=indegree.get(next)!-1;indegree.set(next,degree);if(degree===0)queue.push(next);
    }
  }
  return ordered;
}
function formulaValue(value:unknown):string{
  if(value&&typeof value==="object"&&"__gridArray" in value){
    const matrix=(value as {values?:unknown[][]}).values??[];
    return String(matrix[0]?.[0]??"");
  }
  const result=String(value??"");
  return result==="#CIRCULAR!"?"#CIRC!":result;
}
function evaluateNode(workbook:WorkbookData,node:FormulaNode):string{
  const sheet=workbook.sheets.find(item=>item.id===node.sheetId);
  if(!sheet)return "#REF!";
  const value=evaluate(node.formula,sheet.cells,new Set(),sheet.tables??[],node.address,workbook.namedRanges??{},workbook.sheets,sheet.name);
  return formulaValue(value);
}
function applyCachedValue(workbook:WorkbookData,node:FormulaNode,value:string):WorkbookData{
  return {...workbook,sheets:workbook.sheets.map(sheet=>sheet.id!==node.sheetId?sheet:{...sheet,cells:{...sheet.cells,[node.address]:{...sheet.cells[node.address],cachedValue:value}}})};
}
function impactedFormulas(graph:DependencyGraph,workbook:WorkbookData,changes:{sheetId:string;address:string}[]):Set<CellNodeId>{
  const affected=new Set<CellNodeId>(),queue:CellNodeId[]=[];
  const enqueue=(id:CellNodeId)=>{if(!affected.has(id)){affected.add(id);queue.push(id);}};
  for(const change of changes){
    const id=cellNodeId(change.sheetId,change.address);
    if(graph.formulas.has(id))enqueue(id);
    for(const child of graph.dependents.get(id)??[])enqueue(child);
    for(const dependency of graph.ranges)if(dependency.sheetId===change.sheetId&&rangeContains(dependency.range,change.address))enqueue(dependency.target);
  }
  while(queue.length){
    const current=queue.shift()!;
    for(const child of graph.dependents.get(current)??[])enqueue(child);
  }
  return affected;
}
export function recalculateWorkbookIncremental(workbook:WorkbookData,changes:{sheetId:string;address:string}[]):IncrementalResult{
  const graph=buildDependencyGraph(workbook),affected=impactedFormulas(graph,workbook,changes);
  const cycles=new Set(findCircularReferences(graph));
  for(const id of cycles)if(affected.has(id))affected.add(id);
  let next=workbook;
  const recalculated:CellNodeId[]=[];
  for(const id of cycles){
    if(!affected.has(id))continue;
    const node=graph.formulas.get(id);if(node){next=applyCachedValue(next,node,"#CIRC!");recalculated.push(id);}
  }
  for(const id of topologicalOrder(graph,affected,cycles)){
    const node=graph.formulas.get(id);if(!node)continue;
    next=applyCachedValue(next,node,evaluateNode(next,node));recalculated.push(id);
  }
  const circularReferences=[...cycles].filter(id=>affected.has(id));
  return {workbook:next,graph,recalculated,circularReferences};
}
export function recalculateWorkbookFull(workbook:WorkbookData):IncrementalResult{
  const graph=buildDependencyGraph(workbook),changes=[...graph.formulas.values()].map(node=>({sheetId:node.sheetId,address:node.address}));
  return recalculateWorkbookIncremental(workbook,changes);
}
export function getSheetCircularReferences(workbook:WorkbookData,sheetId?:string):CellNodeId[]{
  const graph=buildDependencyGraph(workbook);
  return findCircularReferences(graph).filter(id=>!sheetId||nodeParts(id).sheetId===sheetId);
}
export function formatCircularReferenceLabel(id:CellNodeId,workbook:WorkbookData):string{
  const {sheetId,address}=nodeParts(id),sheet=workbook.sheets.find(item=>item.id===sheetId);
  return sheet&&workbook.sheets.length>1?`${sheet.name}!${address}`:address;
}
export function dependencyNodeParts(id:CellNodeId){return nodeParts(id);}
