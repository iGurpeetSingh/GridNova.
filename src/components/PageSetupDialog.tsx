import { createPortal } from "react-dom";
import { useEffect, useState } from "react";
import { Check, FileText, LayoutGrid, X } from "lucide-react";
import type { PageMargins, SheetPageSetup } from "../types";

interface Props { open:boolean; settings:SheetPageSetup; onClose:()=>void; onSave:(settings:SheetPageSetup)=>void; onPrint:(settings:SheetPageSetup)=>void; }
const presets:Record<string,PageMargins>={
  Normal:{top:.75,bottom:.75,left:.7,right:.7},
  Wide:{top:1,bottom:1,left:1,right:1},
  Narrow:{top:.5,bottom:.5,left:.5,right:.5}
};
const units=(unit:"in"|"cm",value:number)=>unit==="cm"?value*2.54:value;
const inches=(unit:"in"|"cm",value:number)=>unit==="cm"?value/2.54:value;
export default function PageSetupDialog({open,settings,onClose,onSave,onPrint}:Props){
  const [draft,setDraft]=useState<SheetPageSetup>(settings);
  const [section,setSection]=useState<"page"|"sheet">("page");
  useEffect(()=>{if(open){setDraft({...settings,margins:{...settings.margins},rowBreaks:[...settings.rowBreaks],columnBreaks:[...settings.columnBreaks]});setSection("page")}},[open,settings]);
  if(!open)return null;
  const update=(changes:Partial<SheetPageSetup>)=>setDraft(current=>({...current,...changes}));
  const setMargin=(key:keyof PageMargins,value:string)=>{
    const parsed=Number(value);if(!Number.isFinite(parsed)||parsed<0)return;
    setDraft(current=>({...current,margins:{...current.margins,[key]:inches(current.marginUnit,parsed)}}));
  };
  return createPortal(<div className="page-setup-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="page-setup-dialog" role="dialog" aria-modal="true" aria-labelledby="page-setup-title" onMouseDown={event=>event.stopPropagation()}>
      <header className="page-setup-dialog-header">
        <div className="page-setup-title-icon"><FileText size={19}/></div>
        <div className="min-w-0 flex-1"><span>PAGE LAYOUT</span><h2 id="page-setup-title">Page Setup</h2><p>Configure how this worksheet appears when printed.</p></div>
        <button type="button" className="page-setup-close" aria-label="Close Page Setup" onMouseDown={event=>event.preventDefault()} onClick={onClose}><X size={18}/></button>
      </header>
      <nav className="page-setup-tabs" aria-label="Page setup sections">
        <button type="button" className={section==="page"?"active":""} onMouseDown={event=>event.preventDefault()} onClick={()=>setSection("page")}>Page</button>
        <button type="button" className={section==="sheet"?"active":""} onMouseDown={event=>event.preventDefault()} onClick={()=>setSection("sheet")}>Sheet</button>
      </nav>
      {section==="page"?<div className="page-setup-content">
        <section className="page-setup-section">
          <h3>Orientation</h3>
          <div className="page-setup-orientations">
            <button type="button" className={draft.orientation==="portrait"?"selected":""} onMouseDown={event=>event.preventDefault()} onClick={()=>update({orientation:"portrait"})}><span className="page-orientation-paper portrait"/><strong>Portrait</strong>{draft.orientation==="portrait"&&<Check size={14}/>}</button>
            <button type="button" className={draft.orientation==="landscape"?"selected":""} onMouseDown={event=>event.preventDefault()} onClick={()=>update({orientation:"landscape"})}><span className="page-orientation-paper landscape"/><strong>Landscape</strong>{draft.orientation==="landscape"&&<Check size={14}/>}</button>
          </div>
        </section>
        <div className="page-setup-two-columns">
          <section className="page-setup-section"><h3>Paper Size</h3><select value={draft.paperSize} onChange={event=>update({paperSize:event.target.value as SheetPageSetup["paperSize"]})}><option value="A4">A4 (210 × 297 mm)</option><option value="Letter">Letter (8.5 × 11 in)</option><option value="Legal">Legal (8.5 × 14 in)</option><option value="A3">A3 (297 × 420 mm)</option></select></section>
          <section className="page-setup-section"><h3>Margins</h3><div className="page-setup-margin-presets">{(["Normal","Wide","Narrow"] as const).map(name=><button key={name} type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>update({margins:{...presets[name]}})}>{name}</button>)}</div></section>
        </div>
        <section className="page-setup-section">
          <div className="page-setup-section-heading"><h3>Custom Margins</h3><select aria-label="Margin units" value={draft.marginUnit} onChange={event=>update({marginUnit:event.target.value as "in"|"cm"})}><option value="in">Inches</option><option value="cm">Centimeters</option></select></div>
          <div className="page-setup-margin-grid">
            {(["top","bottom","left","right"] as const).map(key=><label key={key}><span>{key[0].toUpperCase()+key.slice(1)}</span><div><input type="number" min="0" max={draft.marginUnit==="in"?2:5} step={draft.marginUnit==="in" ? .05 : .1} value={Number(units(draft.marginUnit,draft.margins[key]).toFixed(2))} onChange={event=>setMargin(key,event.target.value)}/><small>{draft.marginUnit}</small></div></label>)}
          </div>
        </section>
        <section className="page-setup-section">
          <h3>Scaling</h3>
          <select value={draft.scalingMode} onChange={event=>update({scalingMode:event.target.value as SheetPageSetup["scalingMode"]})}>
            <option value="none">No Scaling (100%)</option><option value="fit-sheet">Fit Sheet on One Page</option><option value="fit-columns">Fit All Columns on One Page</option><option value="fit-rows">Fit All Rows on One Page</option><option value="custom">Custom Percentage</option>
          </select>
          {draft.scalingMode==="custom"&&<label className="page-setup-scale-field"><span>Scale</span><input type="number" min="25" max="200" step="5" value={draft.scalingPercent} onChange={event=>update({scalingPercent:Math.max(25,Math.min(200,Number(event.target.value)||25))})}/><b>%</b></label>}
        </section>
      </div>:<div className="page-setup-content">
        <section className="page-setup-section"><h3>Print Options</h3>
          <label className="page-setup-check"><input type="checkbox" checked={draft.printGridlines} onChange={event=>update({printGridlines:event.target.checked})}/><span>Print Gridlines</span></label>
          <label className="page-setup-check"><input type="checkbox" checked={draft.printHeadings} onChange={event=>update({printHeadings:event.target.checked})}/><span>Print Row &amp; Column Headings</span></label>
        </section>
        <section className="page-setup-section"><h3>Print Titles</h3><p className="page-setup-hint">Optional text printed at the top and bottom of each page.</p>
          <label className="page-setup-text-field"><span>Header</span><input value={draft.header} maxLength={120} onChange={event=>update({header:event.target.value})} placeholder="Header text"/></label>
          <label className="page-setup-text-field"><span>Footer</span><input value={draft.footer} maxLength={120} onChange={event=>update({footer:event.target.value})} placeholder="Footer text"/></label>
        </section>
        <section className="page-setup-sheet-preview"><LayoutGrid size={20}/><span>These settings are saved to <strong>this worksheet</strong> and do not change other sheets.</span></section>
      </div>}
      <footer className="page-setup-dialog-footer">
        <button type="button" className="page-setup-cancel" onMouseDown={event=>event.preventDefault()} onClick={onClose}>Cancel</button>
        <button type="button" className="page-setup-print" onMouseDown={event=>event.preventDefault()} onClick={()=>onPrint(draft)}>Print Preview / Print</button>
        <button type="button" className="page-setup-save" onMouseDown={event=>event.preventDefault()} onClick={()=>onSave(draft)}>OK</button>
      </footer>
    </section>
  </div>,document.body);
}
