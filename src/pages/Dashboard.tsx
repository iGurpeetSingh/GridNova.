import { useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { BarChart3, Clock3, Copy, FileSpreadsheet, FolderOpen, Grid2X2, LayoutList, MoreHorizontal, Plus, Search, Trash2, WalletCards, BriefcaseBusiness, CalendarDays, Pencil, Sparkles } from "lucide-react";
import { useWorkbookStore } from "../store/useWorkbookStore";
import type { TemplateId, WorkbookData } from "../types";

const templates: { id: TemplateId; name: string; description: string; icon: typeof FileSpreadsheet; tone: string; preview: string[] }[] = [
  { id:"blank", name:"Blank workbook", description:"Start with a clean spreadsheet", icon:FileSpreadsheet, tone:"green", preview:["","","","",""] },
  { id:"budget", name:"Monthly budget", description:"Track spending, plans and totals", icon:WalletCards, tone:"violet", preview:["Category","Planned","Actual","Housing","Food"] },
  { id:"sales", name:"Sales tracker", description:"Monitor revenue and product performance", icon:BarChart3, tone:"blue", preview:["Product","Units","Price","Revenue","Share"] },
  { id:"project", name:"Project schedule", description:"Organize tasks, owners and due dates", icon:CalendarDays, tone:"orange", preview:["Task","Owner","Start","Due","Status"] },
];
const relativeTime = (time:number) => {
  const diff=Math.max(0,Date.now()-time), mins=Math.floor(diff/60000);
  if(mins<1)return "Just now"; if(mins<60)return `${mins} min ago`;
  const hours=Math.floor(mins/60); if(hours<24)return `${hours} hr ago`;
  return `${Math.floor(hours/24)} days ago`;
};
function MiniPreview({workbook}:{workbook:WorkbookData}) {
  const sheet=workbook.sheets[0];
  return <div className="mt-4 overflow-hidden rounded-xl border border-[#e6ece7] bg-white">
    <div className="grid grid-cols-4 border-b bg-[#f8faf8]">{["A","B","C","D"].map(x=><div key={x} className="border-r px-2 py-1 text-[9px] text-[#9aa69e]">{x}</div>)}</div>
    {[0,1,2].map(row=><div key={row} className="grid grid-cols-4">{[0,1,2,3].map(col=>{const cell=sheet.cells[`${String.fromCharCode(65+col)}${row+1}`];return <div key={col} className="h-6 truncate border-b border-r px-1.5 py-1 text-[8px]" style={{background:cell?.style?.backgroundColor??"#fff",color:cell?.style?.color??"#8a968e",fontWeight:cell?.style?.bold?700:400}}>{cell?.raw?.startsWith("=")?"ƒx":cell?.raw??""}</div>})}</div>)}
  </div>;
}
export default function Dashboard() {
  const navigate=useNavigate();
  const workbooks=useWorkbookStore(s=>s.workbooks);
  const createWorkbook=useWorkbookStore(s=>s.createWorkbook);
  const deleteWorkbook=useWorkbookStore(s=>s.deleteWorkbook);
  const duplicateWorkbook=useWorkbookStore(s=>s.duplicateWorkbook);
  const renameWorkbook=useWorkbookStore(s=>s.renameWorkbook);
  const [query,setQuery]=useState(""); const [view,setView]=useState<"grid"|"list">("grid");
  const filtered=useMemo(()=>workbooks.filter(w=>w.title.toLowerCase().includes(query.toLowerCase())).sort((a,b)=>b.modifiedAt-a.modifiedAt),[workbooks,query]);
  const create=(id:TemplateId)=>navigate(`/workbook/${createWorkbook(id)}`);
  const rename=(w:WorkbookData)=>{const value=window.prompt("Rename workbook",w.title);if(value?.trim())renameWorkbook(w.id,value)};
  const remove=(w:WorkbookData)=>{if(window.confirm(`Delete “${w.title}”? This cannot be undone.`))deleteWorkbook(w.id)};
  return <main className="app-shell min-h-screen">
    <header className="no-print sticky top-0 z-30 border-b border-[#e7eee9] bg-white/85 backdrop-blur-xl">
      <div className="mx-auto flex max-w-[1440px] items-center gap-4 px-5 py-3.5">
        <div className="flex items-center gap-3">
          <div className="grid h-10 w-10 place-items-center rounded-2xl bg-[#e8f6ed] text-[#107c41]"><Grid2X2 size={21}/></div>
          <div><div className="text-[16px] font-extrabold tracking-tight text-[#1c3325]">GridNova <span className="font-medium text-[#84938a]">Excel</span></div><div className="text-[10px] font-semibold uppercase tracking-[.18em] text-[#9aa69e]">Your workspace</div></div>
        </div>
        <div className="ml-auto flex w-full max-w-[520px] items-center gap-3">
          <label className="flex min-w-0 flex-1 items-center gap-2 rounded-xl border border-[#e5ece7] bg-[#f8faf8] px-3 py-2.5 text-[#809087]"><Search size={16}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search your workbooks..." className="w-full bg-transparent text-sm text-[#26382d] outline-none placeholder:text-[#a4afa8]"/></label>
          <button className="btn btn-primary shrink-0 px-4 py-2.5" onClick={()=>create("blank")}><Plus size={16}/> New workbook</button>
        </div>
      </div>
    </header>
    <div className="dashboard-wrap mx-auto max-w-[1440px] px-8 pb-16 pt-9">
      <section className="mb-10 flex flex-wrap items-end justify-between gap-4">
        <div><div className="mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-[.18em] text-[#159052]"><Sparkles size={14}/> Welcome to your workspace</div><h1 className="text-3xl font-bold tracking-tight text-[#1c3325] sm:text-[34px]">Make room for great ideas.</h1><p className="mt-2 text-sm text-[#829087]">Create, organize and pick up where you left off.</p></div>
        <div className="rounded-2xl border border-[#e4ece6] bg-white/70 px-4 py-3 text-right"><div className="text-2xl font-bold text-[#1d3d2b]">{workbooks.length}</div><div className="text-[11px] font-semibold text-[#89968e]">Saved workbooks</div></div>
      </section>
      <section className="mb-11">
        <div className="mb-4 flex items-center justify-between"><div><h2 className="text-lg font-bold text-[#21382a]">Start with a template</h2><p className="mt-1 text-xs text-[#89968e]">A thoughtful starting point for your next project.</p></div></div>
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-4">
          {templates.map(t=>{const Icon=t.icon;return <button key={t.id} onClick={()=>create(t.id)} className="glass group rounded-2xl p-4 text-left transition duration-200 hover:-translate-y-1 hover:border-[#b9d8c3] hover:shadow-lg">
            <div className="flex items-center justify-between"><span className={`grid h-10 w-10 place-items-center rounded-xl ${t.tone==="green"?"bg-[#e4f5ea] text-[#149052]":t.tone==="violet"?"bg-[#f0eaff] text-[#8a5cf6]":t.tone==="blue"?"bg-[#e7f0ff] text-[#3978e8]":"bg-[#fff0e2] text-[#e58b36]"}`}><Icon size={20}/></span><span className="rounded-full bg-[#f1f6f2] px-2.5 py-1 text-[10px] font-bold text-[#839087]">Template</span></div>
            <div className="mt-4 text-sm font-bold text-[#263a2d]">{t.name}</div><div className="mt-1 text-xs text-[#8a978f]">{t.description}</div>
            <div className="mt-4 grid grid-cols-5 gap-1 overflow-hidden rounded-lg border border-[#e8eee9] bg-white p-1">{t.preview.map((p,i)=><div key={i} className={`h-5 truncate rounded-[3px] px-1 py-1 text-[7px] ${i===0&&t.id!=="blank"?"bg-[#e9f5ed] font-bold text-[#107c41]":"bg-[#f8faf8] text-[#9aa69e]"}`}>{p}</div>)}</div>
          </button>})}
        </div>
      </section>
      <section>
        <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
          <div><h2 className="text-lg font-bold text-[#21382a]">Recent workbooks</h2><p className="mt-1 text-xs text-[#89968e]">Your files, ready when you are.</p></div>
          <div className="flex items-center gap-2"><span className="mr-1 text-xs text-[#8b9990]">{filtered.length} files</span><div className="flex rounded-xl border border-[#e1e9e3] bg-white p-1"><button aria-label="Grid view" onClick={()=>setView("grid")} className={`rounded-lg p-2 ${view==="grid"?"bg-[#eaf5ed] text-[#107c41]":"text-[#9aa69e]"}`}><Grid2X2 size={15}/></button><button aria-label="List view" onClick={()=>setView("list")} className={`rounded-lg p-2 ${view==="list"?"bg-[#eaf5ed] text-[#107c41]":"text-[#9aa69e]"}`}><LayoutList size={15}/></button></div></div>
        </div>
        {filtered.length===0?<div className="glass rounded-2xl py-16 text-center"><FolderOpen className="mx-auto mb-3 text-[#a9b8ae]" size={34}/><div className="font-semibold text-[#405448]">{query?"No workbooks match your search":"Your workspace is ready"}</div><p className="mt-1 text-xs text-[#8b9990]">{query?"Try a different title.":"Create a workbook above to see it here."}</p></div>:
        <div className={view==="grid"?"grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-3":"flex flex-col gap-3"}>
          {filtered.map(w=><article key={w.id} className={`glass group relative rounded-2xl p-4 transition hover:border-[#c9ddce] hover:shadow-md ${view==="list"?"flex items-center gap-5":""}`}>
            <button onClick={()=>navigate(`/workbook/${w.id}`)} className={`min-w-0 text-left ${view==="list"?"flex flex-1 items-center gap-5": "w-full"}`}>
              <div className={`grid h-10 w-10 shrink-0 place-items-center rounded-xl bg-[#e9f5ed] text-[#16854b] ${view==="list"?"":"mb-3"}`}><FileSpreadsheet size={19}/></div>
              <div className="min-w-0 flex-1"><div className="truncate text-sm font-bold text-[#2a3d31]">{w.title}</div><div className="mt-1 flex flex-wrap items-center gap-2 text-[10px] text-[#8d9a91]"><span className="flex items-center gap-1"><Clock3 size={11}/>{relativeTime(w.modifiedAt)}</span><span>·</span><span>{w.sheets.length} {w.sheets.length===1?"sheet":"sheets"}</span></div></div>
            </button>
            {view==="grid"&&<MiniPreview workbook={w}/>}
            <div className={`flex items-center gap-1 ${view==="list"?"ml-auto":"mt-3 justify-end"}`}>
              <button title="Rename" onClick={()=>rename(w)} className="rounded-lg p-2 text-[#8b9a90] hover:bg-[#f0f6f1] hover:text-[#107c41]"><Pencil size={14}/></button>
              <button title="Duplicate" onClick={()=>duplicateWorkbook(w.id)} className="rounded-lg p-2 text-[#8b9a90] hover:bg-[#f0f6f1] hover:text-[#107c41]"><Copy size={14}/></button>
              <button title="Delete" onClick={()=>remove(w)} className="rounded-lg p-2 text-[#9b9290] hover:bg-[#fff0ef] hover:text-[#db5149]"><Trash2 size={14}/></button>
              <button title="Open workbook" onClick={()=>navigate(`/workbook/${w.id}`)} className="btn ml-1 !px-3 !py-1.5"><FolderOpen size={13}/> Open</button>
            </div>
          </article>)}
        </div>}
      </section>
      <footer className="mt-14 flex items-center justify-between border-t border-[#e4ece6] pt-5 text-[10px] text-[#98a59c]"><span>© {new Date().getFullYear()} GridNova Excel</span><span>Designed for focus. Built for momentum.</span></footer>
    </div>
  </main>;
}