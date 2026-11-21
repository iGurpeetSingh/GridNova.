import { FunctionSquare } from "lucide-react";
interface Props { address:string; value:string; onChange:(v:string)=>void; onCommit:()=>void; }
export default function FormulaBar({address,value,onChange,onCommit}:Props){
  return <div className="flex h-11 shrink-0 items-center gap-3 border-b border-[#e6ece8] bg-[#fbfcfb] px-3">
    <div className="flex h-7 w-[66px] shrink-0 items-center justify-center rounded-md border border-[#dfe7e1] bg-white text-xs font-semibold text-[#43564a]">{address}</div>
    <div className="flex items-center gap-1.5 text-[#799184]"><FunctionSquare size={17}/><span className="text-[11px] font-bold italic">fx</span></div>
    <input aria-label="Formula bar" value={value} onChange={e=>onChange(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){onCommit();}}} className="h-8 min-w-0 flex-1 rounded-md border border-transparent bg-transparent px-2 text-sm text-[#26382d] outline-none focus:border-[#a9d4b8] focus:bg-white"/>
  </div>;
}