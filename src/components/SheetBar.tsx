import { useState } from "react";
import { Plus, X, ChevronDown } from "lucide-react";
import type { SheetData } from "../types";
interface Props { sheets:SheetData[]; activeId:string; onSelect:(id:string)=>void; onAdd:()=>void; onRename:(id:string,name:string)=>void; onDelete:(id:string)=>void; }
export default function SheetBar({sheets,activeId,onSelect,onAdd,onRename,onDelete}:Props){
  const [menu,setMenu]=useState<string|null>(null);
  const rename=(sheet:SheetData)=>{const name=window.prompt("Rename sheet",sheet.name);if(name?.trim())onRename(sheet.id,name);};
  return <div className="no-print flex h-11 shrink-0 items-center gap-1 border-t border-[#e2e9e4] bg-[#f7faf7] px-3">
    <button className="grid h-7 w-7 place-items-center rounded-lg text-[#107c41] hover:bg-[#e6f4eb]" onClick={onAdd} title="Add sheet"><Plus size={16}/></button>
    <div className="flex min-w-0 items-center gap-1 overflow-x-auto">{sheets.map(sh=><div key={sh.id} className={`relative flex shrink-0 items-center rounded-t-lg border-b-2 ${activeId===sh.id?"border-[#107c41] bg-white text-[#107c41]":"border-transparent text-[#849188] hover:bg-white"}`}>
      <button onClick={()=>onSelect(sh.id)} onDoubleClick={()=>rename(sh)} className="px-3 py-2 text-[11px] font-semibold">{sh.name}</button>
      <button onClick={()=>setMenu(menu===sh.id?null:sh.id)} className="mr-1 rounded p-1 hover:bg-[#edf4ef]"><ChevronDown size={12}/></button>
      {menu===sh.id&&<div className="absolute bottom-10 left-0 z-50 w-36 rounded-xl border border-[#e3eae5] bg-white p-1 shadow-xl"><button onClick={()=>{rename(sh);setMenu(null)}} className="w-full rounded-lg px-3 py-2 text-left text-xs hover:bg-[#f1f7f3]">Rename sheet</button><button disabled={sheets.length<2} onClick={()=>{onDelete(sh.id);setMenu(null)}} className="w-full rounded-lg px-3 py-2 text-left text-xs text-[#d64b46] hover:bg-[#fff1ef] disabled:opacity-40">Delete sheet</button></div>}
    </div>)}</div>
    <div className="ml-auto whitespace-nowrap text-[10px] text-[#91a096]">{sheets.length} {sheets.length===1?"sheet":"sheets"}</div>
  </div>;
}