import { useState } from "react";
import { Check, Copy, Link2, X, ShieldCheck } from "lucide-react";
interface Props { open:boolean; onClose:()=>void; workbookId:string; title:string; onExport:(type:"xlsx"|"csv"|"pdf")=>void; }
export default function ShareModal({open,onClose,workbookId,title,onExport}:Props){
  const [permission,setPermission]=useState<"Can View"|"Can Comment"|"Can Edit">("Can View");
  const [copied,setCopied]=useState(false);
  if(!open)return null;
  const shareUrl=`${window.location.origin}/workbook/${workbookId}?permission=${permission.toLowerCase().replaceAll(" ","-")}`;
  const copy=async()=>{try{await navigator.clipboard.writeText(shareUrl);setCopied(true);window.setTimeout(()=>setCopied(false),1800)}catch{window.prompt("Copy this link",shareUrl)}};
  return <div className="fixed inset-0 z-[9999] grid place-items-center bg-[#102217]/35 p-4 backdrop-blur-sm" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}>
    <section className="w-full max-w-[440px] rounded-3xl border border-white/70 bg-white p-6 shadow-2xl">
      <div className="flex items-start justify-between"><div className="flex items-center gap-3"><div className="grid h-11 w-11 place-items-center rounded-2xl bg-[#e8f5ed] text-[#107c41]"><Link2 size={20}/></div><div><h2 className="font-bold text-[#21382a]">Share workbook</h2><p className="mt-0.5 max-w-[270px] truncate text-xs text-[#8b9990]">{title}</p></div></div><button onClick={onClose} className="rounded-lg p-2 text-[#8c9991] hover:bg-[#f2f6f3]"><X size={18}/></button></div>
      <div className="mt-6"><label className="mb-2 block text-xs font-semibold text-[#5e7064]">People with the link can</label><div className="grid grid-cols-2 rounded-xl bg-[#f2f6f3] p-1">{(["Can View","Can Comment","Can Edit"] as const).map(v=><button key={v} onMouseDown={event=>event.preventDefault()} onClick={()=>setPermission(v)} className={`rounded-lg py-2 text-xs font-semibold ${permission===v?"bg-white text-[#107c41] shadow-sm":"text-[#89968e]"}`}>{v}</button>)}</div></div>
      <div className="mt-4 rounded-xl border border-[#e5ece7] bg-[#fafcfb] p-3"><div className="truncate text-xs text-[#738178]">{shareUrl}</div><button onClick={copy} className="btn btn-primary mt-3 w-full">{copied?<Check size={14}/>:<Copy size={14}/>} {copied?"Copied":"Copy link"}</button></div>
      <div className="mt-4 flex items-center gap-2 rounded-xl bg-[#f4f8f5] p-3 text-[11px] text-[#78867d]"><ShieldCheck size={15} className="shrink-0 text-[#169453]"/> Link permissions are informational in this local-only version. Shared editing requires a backend and user authentication.</div>
      <div className="mt-5 border-t border-[#edf1ee] pt-4"><div className="mb-3 text-xs font-semibold text-[#5e7064]">Export a copy</div><div className="grid grid-cols-3 gap-2">{(["xlsx","csv","pdf"] as const).map(t=><button key={t} onClick={()=>onExport(t)} className="btn justify-center uppercase">{t}</button>)}</div></div>
    </section>
  </div>;
}
