import { useEffect, useState } from "react";
import { Link2, MessageSquare, X } from "lucide-react";
interface Props { open:boolean; kind:"link"|"comment"; initialValue?:string; initialText?:string; onClose:()=>void; onSubmit:(value:string,secondary?:string)=>void; }
export default function CellInsertDialog({open,kind,initialValue="",initialText="",onClose,onSubmit}:Props){
  const [value,setValue]=useState(initialValue),[secondary,setSecondary]=useState(initialText);
  useEffect(()=>{if(open){setValue(initialValue);setSecondary(initialText)}},[open,kind,initialValue,initialText]);
  if(!open)return null;
  const isLink=kind==="link";
  return <div className="grid-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="grid-dialog cell-insert-dialog" role="dialog" aria-modal="true" aria-labelledby="cell-insert-dialog-title">
      <header><div><span>GRIDNOVA INSERT</span><h2 id="cell-insert-dialog-title">{isLink?"Insert Link":"New Comment"}</h2><p>{isLink?"Attach a clickable link to the selected cell.":"Add a note to the active cell."}</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="cell-insert-fields">
        {isLink?<><label>Text to display<input value={secondary} onChange={event=>setSecondary(event.target.value)} placeholder="Link text"/></label><label>Address<input autoFocus value={value} onChange={event=>setValue(event.target.value)} placeholder="https://example.com"/></label></>
        :<label>Comment<textarea autoFocus value={value} onChange={event=>setValue(event.target.value)} placeholder="Write a comment…" rows={5}/></label>}
      </div>
      <footer><button type="button" className="grid-dialog-secondary" onClick={onClose}>Cancel</button><button type="button" disabled={!value.trim()||(isLink&&!/^https?:\/\//i.test(value.trim()))} onClick={()=>onSubmit(value.trim(),secondary.trim())}>{isLink?<Link2 size={14}/>:<MessageSquare size={14}/>} {isLink?"Insert Link":"Post Comment"}</button></footer>
    </section>
  </div>;
}
