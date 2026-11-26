import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Check, ChevronDown } from "lucide-react";
interface Props {
  address:string; value:string; options:string[]; disabled?:boolean;
  inputTitle?:string; inputMessage?:string;
  onSelect:()=>void; onChange:(value:string)=>void;
}
interface Position {left:number;top:number;width:number;}
export default function ValidationCellDropdown({address,value,options,disabled=false,inputTitle,inputMessage,onSelect,onChange}:Props){
  const triggerRef=useRef<HTMLButtonElement>(null),menuRef=useRef<HTMLDivElement>(null);
  const [open,setOpen]=useState(false),[focused,setFocused]=useState(false),[position,setPosition]=useState<Position|null>(null),[activeIndex,setActiveIndex]=useState(0);
  const show=()=>{
    if(disabled)return;
    const rect=triggerRef.current?.getBoundingClientRect();if(!rect)return;
    const width=Math.max(rect.width,168),left=Math.max(8,Math.min(window.innerWidth-width-8,rect.left));
    const estimatedHeight=Math.min(248,options.length*32+12);
    const top=rect.bottom+4+estimatedHeight>window.innerHeight?Math.max(8,rect.top-estimatedHeight-4):rect.bottom+4;
    setPosition({left,top,width});setActiveIndex(Math.max(0,options.indexOf(value)));setOpen(true);onSelect();
  };
  useEffect(()=>{
    if(!open)return;
    const outside=(event:PointerEvent)=>{if(triggerRef.current?.contains(event.target as Node)||menuRef.current?.contains(event.target as Node))return;setOpen(false);};
    const reposition=()=>{
      const rect=triggerRef.current?.getBoundingClientRect();if(!rect)return;
      const width=Math.max(rect.width,168),left=Math.max(8,Math.min(window.innerWidth-width-8,rect.left)),estimatedHeight=Math.min(248,options.length*32+12);
      const top=rect.bottom+4+estimatedHeight>window.innerHeight?Math.max(8,rect.top-estimatedHeight-4):rect.bottom+4;
      setPosition({left,top,width});
    };
    document.addEventListener("pointerdown",outside);window.addEventListener("resize",reposition);window.addEventListener("scroll",reposition,true);
    requestAnimationFrame(()=>menuRef.current?.querySelector<HTMLButtonElement>("[role=option][aria-selected=true]")?.focus()??menuRef.current?.querySelector<HTMLButtonElement>("[role=option]")?.focus());
    return()=>{document.removeEventListener("pointerdown",outside);window.removeEventListener("resize",reposition);window.removeEventListener("scroll",reposition,true);};
  },[open,options.length,value]);
  const pick=(option:string)=>{onChange(option);setOpen(false);requestAnimationFrame(()=>triggerRef.current?.focus());};
  const menu=open&&position?createPortal(<div ref={menuRef} className="validation-portal-menu" role="listbox" aria-label={`Select value for ${address}`} style={{left:position.left,top:position.top,width:position.width}} onMouseDown={event=>event.preventDefault()}>
    {options.map((option,index)=><button key={`${option}-${index}`} type="button" role="option" aria-selected={value===option} className={value===option?"is-selected":""} onFocus={()=>setActiveIndex(index)} onClick={()=>pick(option)} onKeyDown={event=>{
      if(event.key==="ArrowDown"||event.key==="ArrowUp"){event.preventDefault();const next=(index+(event.key==="ArrowDown"?1:-1)+options.length)%options.length;setActiveIndex(next);menuRef.current?.querySelectorAll<HTMLButtonElement>("[role=option]")[next]?.focus();}
      if(event.key==="Escape"){event.preventDefault();setOpen(false);triggerRef.current?.focus();}
      if(event.key==="Home"||event.key==="End"){event.preventDefault();const next=event.key==="Home"?0:options.length-1;setActiveIndex(next);menuRef.current?.querySelectorAll<HTMLButtonElement>("[role=option]")[next]?.focus();}
    }}><span>{option||"(Blanks)"}</span>{value===option&&<Check size={14}/>}</button>)}
    {!options.length&&<span className="validation-portal-empty">No available values</span>}
  </div>,document.body):null;
  const tip=focused&&inputMessage?createPortal(<div className="validation-input-tip" style={{left:Math.max(8,Math.min(window.innerWidth-248,triggerRef.current?.getBoundingClientRect().left??8)),top:Math.max(8,(triggerRef.current?.getBoundingClientRect().bottom??0)+5)}}><strong>{inputTitle||"Input"}</strong><span>{inputMessage}</span></div>,document.body):null;
  return <>
    <button ref={triggerRef} type="button" disabled={disabled} className="validation-cell-trigger" aria-haspopup="listbox" aria-expanded={open} aria-label={`Choose a validated value for ${address}`} onMouseDown={event=>{event.preventDefault();event.stopPropagation();}} onFocus={()=>setFocused(true)} onBlur={()=>setFocused(false)} onClick={()=>open?setOpen(false):show()} onKeyDown={event=>{
      if(event.key==="ArrowDown"||event.key==="Enter"||event.key===" "){event.preventDefault();if(!open)show();}
      else if(event.key==="Escape"&&open){event.preventDefault();setOpen(false);}
    }}><span className="validation-cell-value">{value||"Select…"}</span><ChevronDown size={11}/></button>
    {menu}{tip}
  </>;
}
