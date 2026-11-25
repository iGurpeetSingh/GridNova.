import { useEffect, useState } from "react";
import { Check, Pencil, Plus, Trash2, X } from "lucide-react";
type NamedRanges=Record<string,string>;
interface Props { open:boolean; namedRanges:NamedRanges; sheetName:string; onClose:()=>void; onSave:(ranges:NamedRanges)=>void; }
export default function NameManagerDialog({open,namedRanges,sheetName,onClose,onSave}:Props){
 const [draft,setDraft]=useState<NamedRanges>(namedRanges);
 const [name,setName]=useState(""),[refersTo,setRefersTo]=useState(""),[editing,setEditing]=useState<string|null>(null),[error,setError]=useState("");
 useEffect(()=>{if(open){setDraft(namedRanges);setName("");setRefersTo("");setEditing(null);setError("");}},[open,namedRanges]);
 if(!open)return null;
 const reset=()=>{setName("");setRefersTo("");setEditing(null);setError("");};
 const save=()=>{
  const normalized=name.trim();
  if(!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(normalized)){setError("Use a name starting with a letter or underscore; spaces are not allowed.");return;}
  if(!/^('([^']|'')+'|[A-Za-z_][A-Za-z0-9_ ]*)!?\s*!?\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(refersTo.trim())&&!/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(refersTo.trim())){
   setError("Enter a cell or range such as B2:B10, or Sheet1!B2:B10.");return;
  }
  const exists=Object.keys(draft).some(key=>key.toLowerCase()===normalized.toLowerCase()&&key.toLowerCase()!==(editing??"").toLowerCase());
  if(exists){setError("A named range with this name already exists.");return;}
  const next={...draft};if(editing&&editing!==normalized)delete next[editing];next[normalized]=refersTo.trim().replace(/^=/,"");
  setDraft(next);onSave(next);reset();
 };
 const beginEdit=(key:string)=>{setEditing(key);setName(key);setRefersTo(draft[key]);setError("");};
 return <div className="grid-dialog-backdrop name-manager-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose();}}>
  <section className="grid-dialog name-manager-dialog" role="dialog" aria-modal="true" aria-labelledby="name-manager-title">
   <header><div><span>FORMULA MANAGEMENT</span><h2 id="name-manager-title">Name Manager</h2><p>Create reusable names for cells and ranges in this workbook.</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
   <div className="name-manager-body">
    <div className="name-manager-editor">
     <label>Name<input value={name} onChange={event=>setName(event.target.value)} placeholder="e.g. SalesTotal"/></label>
     <label>Refers to<input value={refersTo} onChange={event=>setRefersTo(event.target.value)} placeholder={`${sheetName}!B2:B10`}/></label>
     {error&&<p className="name-manager-error">{error}</p>}
     <div className="name-manager-editor-actions">
      <button type="button" onMouseDown={event=>event.preventDefault()} onClick={save}><Check size={13}/>{editing?"Save Changes":"Add Name"}</button>
      {editing&&<button type="button" className="secondary" onClick={reset}>Cancel</button>}
     </div>
    </div>
    <div className="name-manager-list-heading"><strong>Workbook Names</strong><span>{Object.keys(draft).length} name(s)</span></div>
    <div className="name-manager-list">
     {Object.entries(draft).map(([key,value])=><div className="name-manager-row" key={key}>
      <div><strong>{key}</strong><small>{value}</small></div>
      <button type="button" title={`Edit ${key}`} aria-label={`Edit ${key}`} onMouseDown={event=>event.preventDefault()} onClick={()=>beginEdit(key)}><Pencil size={14}/></button>
      <button type="button" title={`Delete ${key}`} aria-label={`Delete ${key}`} onMouseDown={event=>event.preventDefault()} onClick={()=>{const next={...draft};delete next[key];setDraft(next);onSave(next);if(editing===key)reset();}}><Trash2 size={14}/></button>
     </div>)}
     {!Object.keys(draft).length&&<p className="name-manager-empty">No named ranges yet. Add a name above.</p>}
    </div>
   </div>
   <footer><span>Example usage: =SUM(SalesTotal)</span><button type="button" className="grid-dialog-secondary" onClick={onClose}>Close</button></footer>
  </section>
 </div>;
}
