import { useMemo, useRef, useState } from "react";
import type { KeyboardEvent } from "react";
import { ExternalLink } from "lucide-react";
import type { CellRange, SheetData } from "../types";
import { addressOf, columnLabel, parseAddress } from "../types";
import { formatCellValue } from "../utils/formulaEngine";
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
  const [columnFilters,setColumnFilters]=useState<Record<number,string>>({});
  const gridRef=useRef<HTMLDivElement>(null);
  const editOriginal=useRef("");
  const coords=parseAddress(activeCell)??{row:0,col:0};
  const columns=useMemo(()=>Array.from({length:Math.max(sheet.columnCount,26)},(_,i)=>i),[sheet.columnCount]);
  const allRows=useMemo(()=>Array.from({length:Math.max(sheet.rowCount,100)},(_,i)=>i),[sheet.rowCount]);
  const rows=useMemo(()=>allRows.filter(row=>row===0||Object.entries(columnFilters).every(([col,value])=>!value||(sheet.cells[addressOf(row,Number(col))]?.raw??"")===value)),[allRows,sheet.cells,columnFilters]);
  const inside=(row:number,col:number)=>{
    if(!selection)return row===coords.row&&col===coords.col;
    return row>=Math.min(selection.startRow,selection.endRow)&&row<=Math.max(selection.startRow,selection.endRow)&&col>=Math.min(selection.startCol,selection.endCol)&&col<=Math.max(selection.startCol,selection.endCol);
  };
  const moveTo=(row:number,col:number)=>onSelect(addressOf(row,col),{startRow:dragStart?.row??row,startCol:dragStart?.col??col,endRow:row,endCol:col});
  const onKeyDown=(e:KeyboardEvent<HTMLInputElement>,row:number,col:number)=>{
    let nr=row,nc=col;
    if(e.key==="ArrowUp")nr=Math.max(0,row-1);
    else if(e.key==="ArrowDown")nr=row+1;
    else if(e.key==="ArrowLeft"&&e.currentTarget.selectionStart===0)nc=Math.max(0,col-1);
    else if(e.key==="ArrowRight"&&e.currentTarget.selectionEnd===e.currentTarget.value.length)nc=col+1;
    else if(e.key==="Enter")nr=row+1;
    else if(e.key==="Tab"){e.preventDefault();nc=Math.max(0,col+(e.shiftKey?-1:1));updateSheet(sheet.id,{rowCount:Math.max(sheet.rowCount,nr+1),columnCount:Math.max(sheet.columnCount,nc+1)});onSelect(addressOf(nr,nc));requestAnimationFrame(()=>gridRef.current?.querySelector<HTMLInputElement>(`[data-cell="${addressOf(nr,nc)}"]`)?.focus());return;}
    else if(e.key==="Escape"){setCell(addressOf(row,col),editOriginal.current);e.currentTarget.blur();return;}
    else if(e.key==="Delete"||e.key==="Backspace"){e.preventDefault();setCell(addressOf(row,col),"");return;}
    else return;
    e.preventDefault();updateSheet(sheet.id,{rowCount:Math.max(sheet.rowCount,nr+1),columnCount:Math.max(sheet.columnCount,nc+1)});
    onSelect(addressOf(nr,nc));
    requestAnimationFrame(()=>gridRef.current?.querySelector<HTMLInputElement>(`[data-cell="${addressOf(nr,nc)}"]`)?.focus());
  };
  const startResize=(e:React.MouseEvent,col:number)=>{
    e.preventDefault();e.stopPropagation();
    const startX=e.clientX,startWidth=sheet.columnWidths[col]??110;
    const onMove=(event:MouseEvent)=>updateSheet(sheet.id,{columnWidths:{...sheet.columnWidths,[col]:Math.max(55,Math.min(360,startWidth+event.clientX-startX))}});
    const end=()=>{window.removeEventListener("mousemove",onMove);window.removeEventListener("mouseup",end)};
    window.addEventListener("mousemove",onMove);window.addEventListener("mouseup",end);
  };
  return <div ref={gridRef} className="sheet-scroll relative" onMouseUp={()=>setDragging(false)} onMouseLeave={()=>setDragging(false)}>
    <div style={{zoom:zoom/100, minWidth:"max-content"}}>
      <table className="border-collapse table-fixed" style={{minWidth:"100%"}}>
        <colgroup>{sheet.showHeadings&&<col style={{width:48}}/>}{columns.map(c=><col key={c} style={{width:sheet.columnWidths[c]??110}}/>)}</colgroup>
        {sheet.showHeadings&&<thead className="sticky top-0 z-20"><tr><th className="sticky left-0 z-30 h-7 border border-[#e5ebe7] bg-[#f2f6f3]"></th>{columns.map(c=><th key={c} className="relative h-7 border border-[#e5ebe7] bg-[#f2f6f3] text-[10px] font-semibold text-[#78877e]">{columnLabel(c)}<div onMouseDown={e=>startResize(e,c)} className="absolute right-[-3px] top-0 z-40 h-full w-2 cursor-col-resize"/></th>)}</tr><tr><th className="sticky left-0 z-30 h-7 border border-[#e5ebe7] bg-[#f8faf8]"></th>{columns.map(c=><th key={c} className="h-7 border border-[#e5ebe7] bg-[#f8faf8] px-1">{sheet.filters?<select aria-label={`Filter ${columnLabel(c)}`} value={columnFilters[c]??""} onChange={e=>setColumnFilters(v=>({...v,[c]:e.target.value}))} className="max-w-full bg-transparent text-[9px] text-[#738178] outline-none"><option value="">All</option>{Array.from(new Set(allRows.slice(1).map(row=>sheet.cells[addressOf(row,c)]?.raw??"").filter(Boolean))).sort().map(v=><option key={v} value={v}>{v}</option>)}</select>:null}</th>)}</tr></thead>}
        <tbody>{rows.map(row=><tr key={row}>
          {sheet.showHeadings&&<th className={`sticky left-0 z-10 h-[27px] border border-[#e5ebe7] bg-[#f4f7f5] text-[10px] font-medium text-[#86938b] ${sheet.freezeTopRow&&row===0?"sticky top-7 z-30":""}`}>{row+1}</th>}
          {columns.map(col=>{
            const address=addressOf(row,col),cell=sheet.cells[address],value=formatCellValue(cell?.raw??"",sheet.cells,showFormulas,cell?.style);
            const selected=inside(row,col),style=cell?.style??{};
            if(cell?.mergedInto)return null;
            return <td key={col} colSpan={cell?.mergeCols??1} rowSpan={cell?.mergeRows??1} className={`grid-cell ${selected?"selected":""} ${sheet.freezeTopRow&&row===0?"sticky top-14 z-10":""}`} style={{borderColor:sheet.showGridlines?"#e8ece9":"transparent",borderWidth:style.border==="none"?0:1,borderStyle:"solid",background:style.backgroundColor??(selected?"#f0faf3":"#fff"),fontFamily:style.fontFamily??"Inter",fontSize:style.fontSize??11,fontWeight:style.bold?700:400,fontStyle:style.italic?"italic":"normal",color:style.color??"#26382d",textAlign:style.align??"left",textDecoration:[style.underline?"underline":"",style.strike?"line-through":""].filter(Boolean).join(" "),whiteSpace:style.wrap?"normal":"nowrap",overflow:"visible"}}>
              {cell?.checkbox?<div className="grid flex h-full items-center justify-center"><input type="checkbox" checked={cell.raw==="TRUE"} onChange={e=>setCell(address,e.target.checked?"TRUE":"FALSE")} aria-label={`Checkbox ${address}`} className="h-4 w-4 accent-[#107c41]"/></div>:cell?.hyperlink?<a href={cell.hyperlink} target="_blank" rel="noreferrer" className="flex h-full items-center gap-1 px-2 text-[#2778cb] underline"><span className="truncate">{value||cell.hyperlink}</span><ExternalLink size={10}/></a>:<input
                data-cell={address}
                aria-label={`Cell ${address}`}
                onFocus={()=>{editOriginal.current=cell?.raw??"";}}
                type={cell?.datePicker?"date":"text"}
                value={cell?.datePicker?(cell.raw||""):value}
                onClick={e=>{if(e.shiftKey){const anchor=parseAddress(activeCell)??{row,col};onSelect(address,{startRow:anchor.row,startCol:anchor.col,endRow:row,endCol:col})}else onSelect(address,{startRow:row,startCol:col,endRow:row,endCol:col})}}
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