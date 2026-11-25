import { useRef, type ReactNode } from "react";
import { Check, ChevronDown } from "lucide-react";
import FloatingPortal from "./FloatingPortal";

export interface HomeMenuItem {
  label:string;
  action?:string;
  icon:ReactNode;
  description?:string;
  dividerBefore?:boolean;
  disabled?:boolean;
  onSelect?:()=>void;
}
export function HomeMenu({label,children,items,open,onToggle,onClose,onAction,className="home-menu-panel",align="left"}:{
  label:string;children:ReactNode;items:HomeMenuItem[];open:boolean;onToggle:()=>void;onClose:()=>void;onAction:(action:string)=>void;className?:string;align?:"left"|"right";
}){
  const anchor=useRef<HTMLButtonElement>(null);
  return <div className="home-menu-wrap">
    <button ref={anchor} type="button" aria-label={label} title={label} aria-haspopup="menu" aria-expanded={open} className="home-menu-trigger" onMouseDown={event=>event.preventDefault()} onClick={onToggle}>{children}</button>
    <FloatingPortal open={open} anchorRef={anchor} onClose={onClose} role="menu" ariaLabel={label} align={align} className={className}>
      {items.map((item,index)=><div key={item.label+index} className={item.dividerBefore?"home-menu-item-divider":""}>
        <button type="button" role="menuitem" disabled={item.disabled} className="home-menu-item" onMouseDown={event=>event.preventDefault()} onClick={()=>{if(item.onSelect)item.onSelect();else if(item.action)onAction(item.action);onClose();}}>
          <span className="home-menu-item-icon">{item.icon}</span><span className="home-menu-item-copy"><strong>{item.label}</strong>{item.description&&<small>{item.description}</small>}</span>{item.disabled&&<small className="home-menu-disabled">Unavailable</small>}
        </button>
      </div>)}
    </FloatingPortal>
  </div>;
}
export function HomeColorMenu({kind,current,children,open,onToggle,onClose,onFormat}:{
  kind:"font"|"fill";current?:string;children:ReactNode;open:boolean;onToggle:()=>void;onClose:()=>void;onFormat:(style:{color?:string;backgroundColor?:string})=>void;
}){
  const anchor=useRef<HTMLButtonElement>(null);
  const theme=["#FFFFFF","#000000","#1F4E78","#2F75B5","#5B9BD5","#70AD47","#A5A5A5","#FFC000","#ED7D31","#C00000","#7030A0","#00B0F0","#4472C4","#548235","#BF9000","#A64D79"];
  const tints=["#F2F2F2","#7F7F7F","#D9E2F3","#DDEBF7","#DDEBF7","#E2F0D9","#E7E6E6","#FFF2CC","#FCE4D6","#F4CCCC","#E4DFEC","#DDEBF7","#D9EAD3","#E2F0D9","#FFF2CC","#FCE4D6"];
  return <div className="home-color-wrap"><button ref={anchor} type="button" aria-label={kind==="font"?"Font Color":"Fill Color"} title={kind==="font"?"Font Color":"Fill Color"} aria-haspopup="dialog" aria-expanded={open} className="home-color-trigger" onMouseDown={event=>event.preventDefault()} onClick={onToggle}>{children}</button>
    <FloatingPortal open={open} anchorRef={anchor} onClose={onClose} role="group" ariaLabel={kind==="font"?"Font colors":"Fill colors"} className="home-color-gallery">
      <div className="home-color-title">Theme Colors</div><div className="home-color-grid">{theme.map(color=><button key={color} type="button" title={color} aria-label={color} className={current===color?"is-current":""} style={{background:color}} onMouseDown={event=>event.preventDefault()} onClick={()=>{onFormat(kind==="font"?{color}:{backgroundColor:color});onClose();}}/>)}</div>
      <div className="home-color-title home-color-title--tints">Tint / Shade</div><div className="home-color-grid">{tints.map((color,index)=><button key={index} type="button" title={color} aria-label={color} style={{background:color}} onMouseDown={event=>event.preventDefault()} onClick={()=>{onFormat(kind==="font"?{color}:{backgroundColor:color});onClose();}}/>)}</div>
      <button type="button" className="home-color-more" onMouseDown={event=>event.preventDefault()} onClick={()=>{const color=window.prompt("Enter a color in hex format",current??"#000000");if(color&&/^#[0-9a-f]{6}$/i.test(color)){onFormat(kind==="font"?{color}:{backgroundColor:color});onClose();}}}>More Colors…</button>
    </FloatingPortal>
  </div>;
}
export function HomeSelectMenu({label,value,options,open,onToggle,onClose,onPick,icon}:{label:string;value:string;options:string[];open:boolean;onToggle:()=>void;onClose:()=>void;onPick:(value:string)=>void;icon:ReactNode}){
  const anchor=useRef<HTMLButtonElement>(null);
  return <div className="home-menu-wrap"><button ref={anchor} type="button" aria-label={label} title={label} aria-haspopup="menu" aria-expanded={open} className="home-select-trigger" onMouseDown={event=>event.preventDefault()} onClick={onToggle}><span>{value}</span><ChevronDown size={11}/></button>
    <FloatingPortal open={open} anchorRef={anchor} onClose={onClose} role="menu" ariaLabel={label} className="home-select-menu">
      {options.map(option=><button key={option} type="button" role="menuitemradio" aria-checked={value===option} className={`home-select-item ${value===option?"is-selected":""}`} onMouseDown={event=>event.preventDefault()} onClick={()=>{onPick(option);onClose();}}><span className="home-select-item-icon">{icon}</span><span>{option}</span>{value===option&&<Check size={13}/>}</button>)}
    </FloatingPortal>
  </div>;
}
export function HomeStackAction({label,icon,action,onAction}:{label:string;icon:ReactNode;action:string;onAction:(action:string)=>void}){
  return <button type="button" className="home-stack-action" title={label} onMouseDown={event=>event.preventDefault()} onClick={()=>onAction(action)}><span>{icon}</span><span>{label}</span></button>;
}
