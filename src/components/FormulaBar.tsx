import { useEffect, useRef, useState } from "react";
import { Check, ChevronDown, ChevronUp, FunctionSquare, Grid2X2, X } from "lucide-react";
import FloatingPortal from "./FloatingPortal";
interface Props {
 address:string;value:string;onChange:(v:string)=>void;onCommit:()=>void;
 onFocus?:()=>void;onBlur?:()=>void;onCancel?:()=>void;onInsertFunction?:()=>void;onNavigate?:(address:string)=>void;focusInput?:boolean;
}
export default function FormulaBar({address,value,onChange,onCommit,onFocus,onBlur,onCancel,onInsertFunction,onNavigate,focusInput=false}:Props){
 const [nameOpen,setNameOpen]=useState(false),[fxOpen,setFxOpen]=useState(false),[goTo,setGoTo]=useState(address),[expanded,setExpanded]=useState(false);
 const nameRef=useRef<HTMLButtonElement>(null),fxRef=useRef<HTMLButtonElement>(null),inputRef=useRef<HTMLInputElement>(null);
 useEffect(()=>{if(focusInput)inputRef.current?.focus();},[focusInput]);
 return <div className={`formula-bar no-print ${expanded?"is-expanded":""}`}>
  <div className="formula-name-box-wrap">
   <button ref={nameRef} type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>{setGoTo(address);setNameOpen(v=>!v)}} className="formula-name-box" aria-label="Name box" aria-expanded={nameOpen} title="Name box">{address}<ChevronDown size={11}/></button>
   <FloatingPortal open={nameOpen} anchorRef={nameRef} onClose={()=>setNameOpen(false)} role="dialog" ariaLabel="Go to cell" className="formula-name-menu">
    <label>Go to cell<input autoFocus value={goTo} onChange={event=>setGoTo(event.target.value.toUpperCase())} onKeyDown={event=>{if(event.key==="Enter"&&/^[A-Z]{1,3}[1-9][0-9]*$/.test(goTo)){onNavigate?.(goTo);setNameOpen(false)}if(event.key==="Escape")setNameOpen(false)}} placeholder="e.g. G9"/></label>
    <button type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>{if(/^[A-Z]{1,3}[1-9][0-9]*$/.test(goTo))onNavigate?.(goTo);setNameOpen(false)}}><Grid2X2 size={13}/>Go To</button>
   </FloatingPortal>
  </div>
  <div className="formula-action-box">
   <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-control formula-cancel" title="Cancel" aria-label="Cancel" onClick={()=>onCancel?.()}><X size={15}/></button>
   <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-control formula-enter" title="Enter" aria-label="Enter" onClick={onCommit}><Check size={15}/></button>
   <button ref={fxRef} type="button" onMouseDown={event=>event.preventDefault()} className="formula-fx" title="Insert Function" aria-label="Insert Function" aria-expanded={fxOpen} onClick={()=>setFxOpen(v=>!v)}><i>fx</i><ChevronDown size={9}/></button>
   <FloatingPortal open={fxOpen} anchorRef={fxRef} onClose={()=>setFxOpen(false)} role="menu" ariaLabel="Insert Function" className="formula-fx-menu">
    <button type="button" role="menuitem" onMouseDown={event=>event.preventDefault()} onClick={()=>{onInsertFunction?.();setFxOpen(false)}}><FunctionSquare size={16} color="#107C41"/><span><b>Insert Function</b><small>Browse all available formulas</small></span></button>
    <button type="button" role="menuitem" onMouseDown={event=>event.preventDefault()} onClick={()=>{onChange("=SUM(");setFxOpen(false)}}><span className="formula-fx-symbol">Σ</span><span><b>SUM</b><small>Add values in a range</small></span></button>
    <button type="button" role="menuitem" onMouseDown={event=>event.preventDefault()} onClick={()=>{onChange("=AVERAGE(");setFxOpen(false)}}><span className="formula-fx-symbol">x̄</span><span><b>AVERAGE</b><small>Return the arithmetic mean</small></span></button>
   </FloatingPortal>
  </div>
  <div className="formula-input-box">
   <input ref={inputRef} aria-label="Formula bar input" value={value} onFocus={onFocus} onBlur={onBlur} onChange={event=>onChange(event.target.value)} onKeyDown={event=>{if(event.key==="Enter"&&!event.nativeEvent.isComposing&&event.keyCode!==229)onCommit();}} className="formula-input"/>
   <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-expand" title={expanded?"Collapse formula bar":"Expand formula bar"} aria-label={expanded?"Collapse formula bar":"Expand formula bar"} aria-pressed={expanded} onClick={()=>setExpanded(v=>!v)}>{expanded?<ChevronUp size={12}/>:<ChevronDown size={12}/>}</button>
  </div>
 </div>;
}
