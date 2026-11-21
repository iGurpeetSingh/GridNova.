import { useMemo, useRef, useState } from "react";
import { ExternalLink, MessageSquare } from "lucide-react";
import type { CellRange, SheetData } from "../types";
import { addressOf, columnLabel, parseAddress } from "../types";
import { displayValue } from "../utils/formulaEngine";
import { useWorkbookStore } from "../store/useWorkbookStore";

interface Props {
  sheet:SheetData; activeCell:string; selection:CellRange|null; showFormulas:boolean; zoom:number;
  onSelect:(address:string,range?:CellRange)=>void;
}
export default function SpreadsheetGrid({sheet,activeCell,selection,showFormulas,zoom,onSelect}:Props){
  const setCell=useWorkbookStore(s=>s.setCell);
  const updateSheet=useWorkbookStore(s=>s.updateSheet);
  const [dragStart,setDragStart]=useState<{row:number;col:number}|null>(null);
  const [dragging,setDragging]=useState(false);
  const [resize,setResize]=useState<{col:number;startX:number;startWidth:number}|null>(null);
  const gridRef=useRef<HTMLDivElement>(null);
  const coords=parseAddress(activeCell)??{row:0,col:0};
  const columns=useMemo(()=>Array.from({length:Math.max(sheet.columnCount,26)},(_,i)=>i),[sheet.columnCount]);
  const rows=useMemo(()=>Array.from({length:Math.max(sheet.rowCount,100)},(_,i)=>i),[sheet.rowCount]);
  const inside=(row:number,col:number)=>{
    if(!selection)return row===coords.row&&col===coords.col;
    return row>=Math.min(selection.startRow,selection.endRow)&&row<=Math.max(selection.startRow,selection.endRow)&&col>=Math.min(selection.startCol,selection.endCol)&&col<=Math.max(selection.startCol,selection.endCol);
  };
  const moveTo=(row:number,col:number)=>onSelect(addressOf(row,col),{startRow:dragStart?.row??row,startCol:dragStart?.col??col,endRow:row,endCol:col});
  const onKeyDown=(e:React.KeyboardEvent<HTMLInputElement>,row:number,col:number)=>{
    let nr=row,nc=col;
    if(e.key==="ArrowUp")nr=Math.max(0,row-1);
    else if(e.key==="ArrowDown")nr=row+1;
    else if(e.key==="ArrowLeft"&&e.currentTarget.selectionStart===0)nc=Math.max(0,col-1);
    else if(e.key==="ArrowRight"&&e.currentTarget.selectionEnd===e.currentTarget.value.length)nc=col+1;
    else if(e.key==="Enter")nr=row+1;
    else if(e.key==="Tab"){e.preventDefault();nc=col+(e.shiftKey?-1:1);if(nc<0){nc=0}onSelect(addressOf(nr,nc));return;}
    else if(e.key==="Escape"){e.currentTarget.blur();return;}
    else if(e.key==="Delete"||e.key==="Backspace"){if(e.currentTarget.value===""){setCell(addressOf(row,col),"");return;}return;}
    else return;
    e.preventDefault();updateSheet(sheet.id,{rowCount:Math.max(sheet.rowCount,nr+1),columnCount:Math.max(sheet.columnCount,nc+1)});
    onSelect(addressOf(nr,nc));
    requestAnimationFrame(()=>gridRef.current?.querySelector<HTMLInputElement>(`[data-cell="${addressOf(nr,nc)}"]`)?.focus());
  };
  const startResize=(e:React.MouseEvent,col:number)=>{
    e.preventDefault();e.stopPropagation();
    setResize({col,startX:e.clientX,startWidth:sheet.columnWidths[col]??110});
    const move=(event:MouseEvent)=>{if(!resize&&gridRef.current){const width=Math.max(55,Math.min(360,(sheet.columnWidths[col]??110)+event.clientX-e.clientX));updateSheet(sheet.id,{columnWidths:{...sheet.columnWidths,[col]:width}})}};
    const startX=e.clientX,startWidth=sheet.columnWidths[col]??110;
    const onMove=(event:MouseEvent)=>updateSheet(sheet.id,{columnWidths:{...sheet.columnWidths,[col]:Math.max(55,Math.min(360,startWidth+event.clientX-startX))}});
    const end=()=>{window.removeEventListener("mousemove",onMove);window.removeEventListener("mouseup",end);setResize(null)};
    window.addEventListener("mousemove",onMove);window.addEventListener("mouseup",end);
  };
  return <div ref={gridRef} className="sheet-scroll relative" onMouseUp={()=>setDragging(false)} onMouseLeave={()=>setDragging(false)}>
    <div style={{zoom:zoom/100, minWidth:"max-content"}}>
      <table className="border-collapse table-fixed" style={{minWidth:"100%"}}>
        <colgroup>{sheet.showHeadings&&<col style={{width:48}}/>}{columns.map(c=><col key={c} style={{width:sheet.columnWidths[c]??110}}/>)}</colgroup>
        {sheet.showHeadings&&<thead className="sticky top-0 z-20"><tr><th className="sticky left-0 z-30 h-7 border border-[#e5ebe7] bg-[#f2f6f3]"></th>{columns.map(c=><th key={c} className="relative h-7 border border-[#e5ebe7] bg-[#f2f6f3] text-[10px] font-semibold text-[#78877e]">{columnLabel(c)}<div onMouseDown={e=>startResize(e,c)} className="absolute right-[-3px] top-0 z-40 h-full w-2 cursor-col-resize"/></th>)}</tr></thead>}
        <tbody>{rows.map(row=><tr key={row}>
          {sheet.showHeadings&&<th className={`sticky left-0 z-10 h-[27px] border border-[#e5ebe7] bg-[#f4f7f5] text-[10px] font-medium text-[#86938b] ${sheet.freezeTopRow&&row===0?"sticky top-7 z-30":""}`}>{row+1}</th>}
          {columns.map(col=>{
            const address=addressOf(row,col),cell=sheet.cells[address],value=displayValue(cell?.raw??"",sheet.cells,showFormulas);
            const selected=inside(row,col),style=cell?.style??{};
            return <td key={col} className={`grid-cell ${selected?"selected":""}`} style={{borderColor:sheet.showGridlines?"#e8ece9":"transparent",background:selected?"#f0faf3":style.backgroundColor??"#fff",fontFamily:style.fontFamily??"Inter",fontSize:style.fontSize??11,fontWeight:style.bold?700:400,fontStyle:style.italic?"italic":"normal",color:style.color??"#26382d",textAlign:style.align??"left",textDecoration:[style.underline?"underline":"",style.strike?"line-through":""].filter(Boolean).join(" "),whiteSpace:style.wrap?"normal":"nowrap",overflow:"visible"}}>
              {cell?.checkbox?<div className="grid flex h-full items-center justify-center"><input type="checkbox" checked={cell.raw==="TRUE"} onChange={e=>setCell(address,e.target.checked?"TRUE":"FALSE")} aria-label={`Checkbox ${address}`} className="h-4 w-4 accent-[#107c41]"/></div>:cell?.hyperlink?<a href={cell.hyperlink} target="_blank" rel="noreferrer" className="flex h-full items-center gap-1 px-2 text-[#2778cb] underline"><span className="truncate">{value||cell.hyperlink}</span><ExternalLink size={10}/></a>:<input
                data-cell={address}
                aria-label={`Cell ${address}`}
                value={value}
                onFocus={()=>onSelect(address)}
                onClick={e=>{if(e.shiftKey)moveTo(row,col);else onSelect(address,{startRow:row,startCol:col,endRow:row,endCol:col})}}
                onMouseDown={()=>{setDragStart({row,col});setDragging(true);}}
                onMouseEnter={()=>{if(dragging&&dragStart)moveTo(row,col)}}
                onChange={e=>setCell(address,e.target.value)}
                onKeyDown={e=>onKeyDown(e,row,col)}
                className="cell-input"
                style={{textAlign:style.align??"left",fontFamily:style.fontFamily??"Inter",fontSize:style.fontSize??11,fontWeight:style.bold?700:400,fontStyle:style.italic?"italic":"normal",color:style.color??"#26382d",textDecoration:style.underline?"underline":"none",whiteSpace:style.wrap?"normal":"nowrap"}}
                title={cell?.comment??""}
              />}
              {cell?.comment&&<span className="comment-mark" title={cell.comment}/>}
            </td>;
          })}
        </tr>)}</tbody>
      </table>
    </div>
  </div>;
}