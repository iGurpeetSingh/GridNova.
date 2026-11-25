import { useRef, useState } from "react";
import FloatingPortal from "./FloatingPortal";
import { Plus, ChevronDown } from "lucide-react";
import type { SheetData } from "../types";
interface Props { sheets:SheetData[]; activeId:string; onSelect:(id:string)=>void; onAdd:()=>void; onRename:(id:string,name:string)=>void; onDelete:(id:string)=>void; }
type AnchorRef = { current: HTMLButtonElement | null };
export default function SheetBar({sheets,activeId,onSelect,onAdd,onRename,onDelete}:Props){
  const [menu,setMenu]=useState<string|null>(null);
  const menuAnchors=useRef(new Map<string,AnchorRef>());
  const rename=(sheet:SheetData)=>{const name=window.prompt("Rename sheet",sheet.name);if(name?.trim())onRename(sheet.id,name);};
  return <div className="no-print flex h-11 shrink-0 items-center gap-1 border-t border-[#e2e9e4] bg-[#f7faf7] px-3">
    <button className="grid h-7 w-7 place-items-center rounded-lg text-[#107c41] hover:bg-[#e6f4eb]" onClick={onAdd} title="Add sheet"><Plus size={16}/></button>
    <div className="flex min-w-0 items-center gap-1 overflow-x-auto">{sheets.map(sh=>{
      let anchorRef=menuAnchors.current.get(sh.id);
      if(!anchorRef){anchorRef={current:null};menuAnchors.current.set(sh.id,anchorRef);}
      return <div key={sh.id} className={`relative flex shrink-0 items-center rounded-t-lg border-b-2 ${activeId===sh.id?"border-[#107c41] bg-white text-[#107c41]":"border-transparent text-[#849188] hover:bg-white"}`}>
        <button onClick={()=>onSelect(sh.id)} onDoubleClick={()=>rename(sh)} className="px-3 py-2 text-[11px] font-semibold">{sh.name}</button>
        <button ref={anchorRef} aria-label={`${sh.name} options`} aria-expanded={menu===sh.id} aria-haspopup="menu" onClick={()=>setMenu(menu===sh.id?null:sh.id)} className="mr-1 rounded p-1 hover:bg-[#edf4ef]"><ChevronDown size={12}/></button>
        <FloatingPortal open={menu===sh.id} anchorRef={anchorRef} onClose={()=>setMenu(null)} role="menu" ariaLabel={`${sh.name} sheet options`} className="ribbon-menu-panel"><button type="button" role="menuitem" onMouseDown={event=>event.preventDefault()} onClick={()=>{rename(sh);setMenu(null)}} className="ribbon-menu-option">Rename sheet</button><button type="button" role="menuitem" disabled={sheets.length<2} onMouseDown={event=>event.preventDefault()} onClick={()=>{onDelete(sh.id);setMenu(null)}} className="ribbon-menu-option disabled:opacity-40">Delete sheet</button></FloatingPortal>
      </div>;
    })}</div>
    <div className="ml-auto whitespace-nowrap text-[10px] text-[#91a096]">{sheets.length} {sheets.length===1?"sheet":"sheets"}</div>
  </div>;
}
