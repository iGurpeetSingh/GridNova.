import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import type { WorkbookData, SheetData } from "../types";
import { parseAddress, addressOf } from "../types";
import { buildDependencyGraph, cellNodeId, extractFormulaReferences, type CellNodeId, type RangeDependency } from "../lib/dependency-graph";

type AuditMode="precedents"|"dependents";
interface Props { workbook:WorkbookData; sheet:SheetData; activeCell:string; mode:AuditMode|null; }
interface AuditEdge { from:string; to:string; }
interface DrawnEdge { d:string; start:{x:number;y:number}; end:{x:number;y:number}; }
function buildEdges(workbook:WorkbookData,sheet:SheetData,address:string,mode:AuditMode):AuditEdge[]{
  const current=cellNodeId(sheet.id,address),edges:AuditEdge[]=[];
  if(mode==="precedents"){
    const formula=sheet.cells[address]?.raw;
    if(!formula?.startsWith("="))return [];
    for(const reference of extractFormulaReferences(formula,workbook.namedRanges??{})){
      const targetSheet=reference.sheetName?workbook.sheets.find(item=>item.name.toLowerCase()===reference.sheetName!.toLowerCase()):sheet;
      if(!targetSheet||targetSheet.id!==sheet.id)continue;
      const height=reference.end.row-reference.start.row+1,width=reference.end.col-reference.start.col+1;
      const limit= Math.min(64,height*width);
      let count=0;
      for(let row=reference.start.row;row<=reference.end.row&&count<limit;row++)for(let col=reference.start.col;col<=reference.end.col&&count<limit;col++){
        const from=addressOf(row,col);if(from!==address){edges.push({from,to:address});count++;}
      }
    }
  }else{
    const graph=buildDependencyGraph(workbook);
    const sources=new Set<CellNodeId>(graph.dependents.get(current)??[]);
    graph.ranges.forEach((range:RangeDependency)=>{if(range.sheetId===sheet.id){const p=parseAddress(address);if(p&&p.row>=range.range.startRow&&p.row<=range.range.endRow&&p.col>=range.range.startCol&&p.col<=range.range.endCol)sources.add(range.target);}});
    sources.forEach(id=>{
      const separator=id.lastIndexOf("!"),sheetId=id.slice(0,separator),target=id.slice(separator+1);
      if(sheetId===sheet.id)edges.push({from:address,to:target});
    });
  }
  return edges;
}
export default function FormulaAuditOverlay({workbook,sheet,activeCell,mode}:Props){
  const [drawn,setDrawn]=useState<DrawnEdge[]>([]);
  const edges=useMemo(()=>mode?buildEdges(workbook,sheet,activeCell,mode):[],[workbook,sheet,activeCell,mode]);
  useEffect(()=>{
    if(!mode||!edges.length){setDrawn([]);return;}
    let raf=0;
    const update=()=>{
      const targetElement=document.querySelector<HTMLElement>(`[data-cell="${activeCell}"]`);
      if(!targetElement){setDrawn([]);return;}
      const targetRect=targetElement.getBoundingClientRect();
      const next:DrawnEdge[]=[];
      edges.forEach(edge=>{
        const fromElement=document.querySelector<HTMLElement>(`[data-cell="${edge.from}"]`);
        const toElement=document.querySelector<HTMLElement>(`[data-cell="${edge.to}"]`);
        if(!fromElement||!toElement)return;
        const fromRect=fromElement.getBoundingClientRect(),toRect=toElement.getBoundingClientRect();
        if(fromRect.width===0||toRect.width===0)return;
        const x1=fromRect.left+fromRect.width/2,y1=fromRect.top+fromRect.height/2;
        const x2=toRect.left+toRect.width/2,y2=toRect.top+toRect.height/2;
        const dx=x2-x1,dy=y2-y1,curve=Math.max(22,Math.min(90,Math.hypot(dx,dy)*.22));
        const direction=mode==="precedents"?1:-1;
        const c1x=x1+dx*.35,c1y=y1+dy*.35-direction*curve;
        const c2x=x1+dx*.65,c2y=y1+dy*.65-direction*curve;
        next.push({d:`M ${x1} ${y1} C ${c1x} ${c1y}, ${c2x} ${c2y}, ${x2} ${y2}`,start:{x:x1,y:y1},end:{x:x2,y:y2}});
      });
      setDrawn(next);
    };
    const schedule=()=>{cancelAnimationFrame(raf);raf=requestAnimationFrame(update);};
    schedule();window.addEventListener("resize",schedule);window.addEventListener("scroll",schedule,true);
    const observer=new ResizeObserver(schedule);document.querySelectorAll<HTMLElement>("[data-cell]").forEach(el=>observer.observe(el));
    return()=>{cancelAnimationFrame(raf);window.removeEventListener("resize",schedule);window.removeEventListener("scroll",schedule,true);observer.disconnect();};
  },[edges,mode,activeCell]);
  if(!mode||!drawn.length)return null;
  return createPortal(<svg aria-hidden="true" className="formula-audit-overlay" viewBox={`0 0 ${window.innerWidth} ${window.innerHeight}`} preserveAspectRatio="none">
    <defs><marker id="formula-audit-arrow" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L0,6 L6,3 z" fill="#0F6CBD"/></marker></defs>
    {drawn.map((edge,index)=><g key={index}><path d={edge.d} stroke="#0F6CBD" strokeWidth="2" fill="none" markerEnd="url(#formula-audit-arrow)"/><circle cx={edge.start.x} cy={edge.start.y} r="3.1" fill="#fff" stroke="#0F6CBD" strokeWidth="1.6"/></g>)}
  </svg>,document.body);
}
