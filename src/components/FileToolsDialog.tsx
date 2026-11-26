import { createPortal } from "react-dom";
import { useEffect, useState } from "react";
import { Check, Clipboard, FileClock, Keyboard, MonitorCog, Settings2, X } from "lucide-react";
import type { WorkbookData } from "../types";

export interface FileVersionSnapshot { id:string; savedAt:number; workbook:WorkbookData; }
export type FileToolsMode="embed"|"history"|"regional"|"calculation"|"display"|"shortcuts";
interface Props {
  mode:FileToolsMode|null;
  workbook:WorkbookData;
  history:FileVersionSnapshot[];
  showFormulaBar:boolean;
  autoSaveEnabled:boolean;
  gridlines:boolean;
  onClose:()=>void;
  onRestore:(snapshot:FileVersionSnapshot)=>void;
  onRegionalChange:(value:NonNullable<WorkbookData["regionalSettings"]>)=>void;
  onCalculationChange:(value:"Automatic"|"Manual")=>void;
  onGridlinesChange:()=>void;
  onFormulaBarChange:()=>void;
  onAutoSaveChange:()=>void;
}
const defaultRegional={currency:"INR" as const,dateFormat:"DD/MM/YYYY" as const};
const shortcutRows=[["Ctrl + C","Copy selected cells"],["Ctrl + X","Cut selected cells"],["Ctrl + V","Paste cells"],["Ctrl + Z","Undo"],["Ctrl + Y","Redo"],["Ctrl + F","Find in sheet"],["Ctrl + 1","Format Cells"],["F2","Edit active cell"],["Enter","Commit cell value"],["Esc","Cancel cell edit"],["Arrow keys","Move active cell"],["Ctrl + D","Fill Down"],["Ctrl + R","Fill Right"]];
export default function FileToolsDialog({mode,workbook,history,showFormulaBar,autoSaveEnabled,gridlines,onClose,onRestore,onRegionalChange,onCalculationChange,onGridlinesChange,onFormulaBarChange,onAutoSaveChange}:Props){
  const [currency,setCurrency]=useState<"INR"|"USD"|"EUR">(workbook.regionalSettings?.currency??defaultRegional.currency);
  const [dateFormat,setDateFormat]=useState<"DD/MM/YYYY"|"MM/DD/YYYY"|"YYYY-MM-DD">(workbook.regionalSettings?.dateFormat??defaultRegional.dateFormat);
  const [copied,setCopied]=useState(false);
  useEffect(()=>{setCurrency(workbook.regionalSettings?.currency??"INR");setDateFormat(workbook.regionalSettings?.dateFormat??"DD/MM/YYYY");setCopied(false)},[workbook.id,mode]);
  if(!mode)return null;
  const embed=`<iframe src="${window.location.origin}/workbook/${workbook.id}?embed=1" title="${workbook.title.replace(/["<>]/g,"")}" width="100%" height="600" frameborder="0"></iframe>`;
  const copyEmbed=async()=>{try{await navigator.clipboard.writeText(embed);setCopied(true);window.setTimeout(()=>setCopied(false),1600)}catch{window.prompt("Copy embed code",embed)}};
  const title:Record<FileToolsMode,string>={embed:"Embed Workbook",history:"Version History",regional:"Regional Format Settings",calculation:"Calculation Options",display:"Grid & Display",shortcuts:"Keyboard Shortcuts"};
  const subtitle:Record<FileToolsMode,string>={embed:"Copy this snippet to embed the workbook in a website.",history:"Restore a previously saved local workbook snapshot.",regional:"Choose the default currency and date display format.",calculation:"Choose when formulas are recalculated.",display:"Control worksheet display and local saving.",shortcuts:"Keyboard shortcuts available in GridNova."};
  return createPortal(<div className="file-tools-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="file-tools-dialog" role="dialog" aria-modal="true" aria-labelledby="file-tools-title" onMouseDown={event=>event.stopPropagation()}>
      <header className="file-tools-header"><div className="file-tools-title-icon">{mode==="history"?<FileClock size={18}/>:mode==="display"?<MonitorCog size={18}/>:mode==="shortcuts"?<Keyboard size={18}/>:<Settings2 size={18}/>}</div><div className="min-w-0 flex-1"><h2 id="file-tools-title">{title[mode]}</h2><p>{subtitle[mode]}</p></div><button type="button" aria-label="Close" onMouseDown={event=>event.preventDefault()} onClick={onClose}><X size={17}/></button></header>
      {mode==="embed"&&<div className="file-tools-body"><label className="file-tools-label">Embed code</label><textarea className="file-tools-code" readOnly value={embed} rows={5}/><button className="file-tools-primary" type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>void copyEmbed()}>{copied?<Check size={15}/>:<Clipboard size={15}/>} {copied?"Copied":"Copy Embed Code"}</button><p className="file-tools-note">Embedding displays this workbook route. It does not grant public access or bypass workbook permissions.</p></div>}
      {mode==="history"&&<div className="file-tools-body file-tools-history">
        {history.length?history.map(item=><div className="file-history-entry" key={item.id}><div><strong>{item.workbook.title}</strong><span>{new Date(item.savedAt).toLocaleString()} · {item.workbook.sheets.length} worksheet(s)</span></div><button type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>onRestore(item)}>Restore</button></div>):<div className="file-tools-empty">No local snapshots have been captured yet. Changes are saved here as the workbook is edited.</div>}
      </div>}
      {mode==="regional"&&<div className="file-tools-body">
        <label className="file-tools-field"><span>Default Currency</span><select value={currency} onChange={event=>setCurrency(event.target.value as typeof currency)}><option value="INR">Indian Rupee (₹)</option><option value="USD">US Dollar ($)</option><option value="EUR">Euro (€)</option></select></label>
        <label className="file-tools-field"><span>Date Format</span><select value={dateFormat} onChange={event=>setDateFormat(event.target.value as typeof dateFormat)}><option value="DD/MM/YYYY">DD/MM/YYYY</option><option value="MM/DD/YYYY">MM/DD/YYYY</option><option value="YYYY-MM-DD">YYYY-MM-DD</option></select></label>
        <div className="file-tools-actions"><button type="button" className="file-tools-primary" onMouseDown={event=>event.preventDefault()} onClick={()=>{onRegionalChange({currency,dateFormat});onClose()}}>Save Settings</button></div>
      </div>}
      {mode==="calculation"&&<div className="file-tools-body">
        {(["Automatic","Manual"] as const).map(value=><label className="file-tools-radio" key={value}><input type="radio" name="calculation-mode" checked={(workbook.calculationMode??"Automatic")===value} onChange={()=>onCalculationChange(value)}/><span><strong>{value}</strong><small>{value==="Automatic"?"Recalculate formulas whenever workbook values change.":"Recalculate formulas only when requested."}</small></span></label>)}
        <p className="file-tools-note">This setting is stored with the current workbook.</p>
      </div>}
      {mode==="display"&&<div className="file-tools-body">
        <label className="file-tools-toggle"><span><strong>Show Gridlines</strong><small>Display cell grid borders in the active sheet.</small></span><input type="checkbox" checked={gridlines} onChange={onGridlinesChange}/></label>
        <label className="file-tools-toggle"><span><strong>Show Formula Bar</strong><small>Show the name box and formula input row.</small></span><input type="checkbox" checked={showFormulaBar} onChange={onFormulaBarChange}/></label>
        <label className="file-tools-toggle"><span><strong>Auto-Save</strong><small>Persist workbook changes in this browser's local storage.</small></span><input type="checkbox" checked={autoSaveEnabled} onChange={onAutoSaveChange}/></label>
        {!autoSaveEnabled&&<p className="file-tools-warning">Auto-Save is off. Changes remain in this browser session but will not be restored after a refresh.</p>}
      </div>}
      {mode==="shortcuts"&&<div className="file-tools-body file-tools-shortcuts">{shortcutRows.map(([keys,description])=><div key={keys}><kbd>{keys}</kbd><span>{description}</span></div>)}</div>}
      {mode!=="regional"&&<footer className="file-tools-footer"><button type="button" onMouseDown={event=>event.preventDefault()} onClick={onClose}>Close</button></footer>}
    </section>
  </div>,document.body);
}
