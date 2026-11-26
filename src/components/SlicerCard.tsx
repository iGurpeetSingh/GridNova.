import { useEffect, useMemo, useRef, useState, type PointerEvent as ReactPointerEvent } from "react";
import { FilterX, Grip, Plus, X } from "lucide-react";
import type { SheetData, SlicerCardData } from "../types";
import { addressOf } from "../types";
import { useWorkbookStore } from "../store/useWorkbookStore";
import { tableColumnValues } from "../utils/tableUtils";
interface Props { sheet:SheetData; sourceSheet?:SheetData; slicer:SlicerCardData; }
type DragState={kind:"move"|"resize";pointerX:number;pointerY:number;left:number;top:number;width:number;height:number};
function rangeColumnValues(source:SheetData,slicer:SlicerCardData):string[]{
  const range=slicer.sourceRange;if(!range)return [];
  const result=new Set<string>();
  for(let row=range.startRow+1;row<=range.endRow;row++)result.add(source.cells[addressOf(row,slicer.columnIndex)]?.raw??"");
  return [...result].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"}));
}
export default function SlicerCard({sheet,sourceSheet,slicer}:Props){
  const updateSheet=useWorkbookStore(state=>state.updateSheet);
  const table=sheet.tables?.find(item=>item.id===slicer.tableId);
  const pivot=sheet.pivotTables?.find(item=>item.id===slicer.pivotTableId);
  const source=sourceSheet??sheet;
  const values=useMemo(()=>table?tableColumnValues(table,slicer.columnIndex,source.cells):rangeColumnValues(source,slicer),[table,slicer.columnIndex,slicer.sourceRange,source.cells]);
  const fieldIndex=slicer.fieldIndex??(slicer.sourceRange?Math.max(0,slicer.columnIndex-slicer.sourceRange.startCol):0);
  const active=table?.filters?.[slicer.columnIndex]?.selectedValues??pivot?.filters?.[fieldIndex]??values;
  const label=source.cells[addressOf(slicer.sourceRange?.startRow??table?.range.startRow??0,slicer.columnIndex)]?.raw??"Field";
  const [box,setBox]=useState({x:slicer.x,y:slicer.y,width:slicer.width,height:slicer.height});
  const drag=useRef<DragState|null>(null);
  const currentRef=useRef(box);currentRef.current=box;
  useEffect(()=>{setBox({x:slicer.x,y:slicer.y,width:slicer.width,height:slicer.height});},[slicer.x,slicer.y,slicer.width,slicer.height]);
  useEffect(()=>{
    const move=(event:PointerEvent)=>{
      if(!drag.current)return;
      const start=drag.current,dx=event.clientX-start.pointerX,dy=event.clientY-start.pointerY;
      const next=start.kind==="move"?{...currentRef.current,x:Math.max(0,start.left+dx),y:Math.max(0,start.top+dy)}:{...currentRef.current,width:Math.max(190,Math.min(420,start.width+dx)),height:Math.max(150,Math.min(620,start.height+dy))};
      currentRef.current=next;setBox(next);
    };
    const up=()=>{
      if(!drag.current)return;
      drag.current=null;
      const next=currentRef.current;
      updateSheet(sheet.id,{slicers:(sheet.slicers??[]).map(item=>item.id===slicer.id?{...item,x:next.x,y:next.y,width:next.width,height:next.height}:item)});
    };
    window.addEventListener("pointermove",move);window.addEventListener("pointerup",up);
    return()=>{window.removeEventListener("pointermove",move);window.removeEventListener("pointerup",up)};
  },[sheet.id,sheet.slicers,slicer.id,updateSheet]);
  if(!table&&!pivot)return null;
  const setSelected=(next:string[])=>{
    if(table){
      const filters={...(table.filters??{})};
      if(next.length===values.length)delete filters[slicer.columnIndex];
      else filters[slicer.columnIndex]={...(filters[slicer.columnIndex]??{}),selectedValues:next};
      updateSheet(sheet.id,{tables:(sheet.tables??[]).map(item=>item.id===table.id?{...item,filters}:item)});
    }else if(pivot){
      const filters={...(pivot.filters??{})};
      if(next.length===values.length)delete filters[fieldIndex];
      else filters[fieldIndex]=next;
      const filterFields=next.length===values.length?(pivot.filterFields??[]).filter(field=>field!==fieldIndex):Array.from(new Set([...(pivot.filterFields??[]),fieldIndex]));
      updateSheet(sheet.id,{pivotTables:(sheet.pivotTables??[]).map(item=>item.id===pivot.id?{...item,filters,filterFields}:item)});
    }
  };
  const begin=(kind:"move"|"resize")=>(event:ReactPointerEvent)=>{
    event.preventDefault();event.stopPropagation();
    drag.current={kind,pointerX:event.clientX,pointerY:event.clientY,left:box.x,top:box.y,width:box.width,height:box.height};
  };
  const toggle=(value:string)=>{
    if(!slicer.multiSelect){setSelected([value]);return;}
    const next=active.includes(value)?active.filter(item=>item!==value):[...active,value];
    setSelected(next);
  };
  return <section className="floating-slicer-card" style={{left:box.x,top:box.y,width:box.width,height:box.height}} aria-label={`Slicer ${table?.name??pivot?.name??"PivotTable"} ${label}`}>
    <header className="floating-slicer-header" onPointerDown={begin("move")}><div><Grip size={13}/><strong>{table?.name??pivot?.name??"PivotTable"} · {label}</strong></div><button type="button" title="Clear Filter" onPointerDown={event=>event.stopPropagation()} onClick={()=>setSelected(values)}><FilterX size={14}/></button><button type="button" title="Close slicer" onPointerDown={event=>event.stopPropagation()} onClick={()=>updateSheet(sheet.id,{slicers:(sheet.slicers??[]).filter(item=>item.id!==slicer.id)})}><X size={14}/></button></header>
    <div className="floating-slicer-toolbar"><button type="button" className={slicer.multiSelect?"is-active":""} onClick={()=>updateSheet(sheet.id,{slicers:(sheet.slicers??[]).map(item=>item.id===slicer.id?{...item,multiSelect:!item.multiSelect}:item)})}>Multi-select</button><span>{active.length===values.length?"All items":`${active.length} selected`}</span></div>
    <div className="floating-slicer-values">{values.map(value=><button key={value||"__blank"} type="button" aria-pressed={active.includes(value)} className={active.includes(value)?"is-selected":""} onClick={()=>toggle(value)}>{value||"(blank)"}</button>)}</div>
    <button type="button" className="floating-slicer-resize" aria-label="Resize slicer" onPointerDown={begin("resize")}><Plus size={11}/></button>
  </section>;
}
