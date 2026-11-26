import { useMemo, useRef, useState, type ReactNode } from "react";
import { ChevronDown, ChevronRight, Search } from "lucide-react";
import FloatingPortal from "./FloatingPortal";
import { AUTO_SUM_FUNCTIONS, FORMULA_GROUPS, type FormulaEntry } from "../lib/formula-catalog";

interface Props {
 label:string;category:string;icon:ReactNode;caption:ReactNode;onSelect:(name:string)=>void;
 onMore?:()=>void;moreMenu?:boolean;
}
export default function FormulaFunctionDropdown({label,category,icon,caption,onSelect,onMore,moreMenu=false}:Props){
 const anchor=useRef<HTMLButtonElement>(null),[open,setOpen]=useState(false),[query,setQuery]=useState(""),[activeCategory,setActiveCategory]=useState(""),[hovered,setHovered]=useState<FormulaEntry|null>(null);
 const autoSum=category==="AutoSum";
 const entries=autoSum?AUTO_SUM_FUNCTIONS.map(name=>FORMULA_GROUPS["Math & Trig"].find(x=>x.name===name)??FORMULA_GROUPS.Statistical.find(x=>x.name===name)).filter(Boolean) as FormulaEntry[]:(FORMULA_GROUPS[category]??[]);
 const categories=["Statistical","Engineering","Information","Web"];
 const activeEntries=moreMenu&&activeCategory?(FORMULA_GROUPS[activeCategory]??[]):entries;
 const filtered=useMemo(()=>activeEntries.filter(item=>`${item.name} ${item.syntax} ${item.description}`.toLowerCase().includes(query.toLowerCase())),[activeEntries,query]);
 const close=()=>{setOpen(false);setQuery("");setHovered(null);setActiveCategory("");};
 const choose=(name:string)=>{onSelect(name);close();};
 return <div className="formula-function-trigger-wrap">
  <button ref={anchor} type="button" className="formula-large-menu-trigger" title={label} aria-label={label} aria-expanded={open} onMouseDown={event=>event.preventDefault()} onClick={()=>setOpen(value=>!value)}>
   {icon}<span>{caption}</span><ChevronDown size={9}/>
  </button>
  <FloatingPortal open={open} anchorRef={anchor} onClose={close} role="menu" ariaLabel={`${label} functions`} className="formula-function-menu">
   {moreMenu&&!activeCategory?<div className="formula-function-category-list">
     <div className="formula-function-menu-title">More Functions</div>
     {categories.map(item=><button key={item} type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>{setActiveCategory(item);setQuery("");}}><span><b>{item}</b><small>{FORMULA_GROUPS[item]?.length??0} functions</small></span><ChevronRight size={14}/></button>)}
     <button type="button" className="formula-function-open-library" onMouseDown={event=>event.preventDefault()} onClick={()=>{close();onMore?.();}}>More Functions...</button>
   </div>:<>
    {moreMenu&&<div className="formula-function-submenu-heading"><button type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>{setActiveCategory("");setQuery("");}}>‹ All categories</button><strong>{activeCategory}</strong></div>}
    <label className="formula-function-search"><Search size={13}/><input autoFocus value={query} onChange={event=>setQuery(event.target.value)} placeholder="Filter functions..."/></label>
    <div className="formula-function-list">
     {filtered.map(item=><button type="button" role="menuitem" key={item.name} onMouseDown={event=>event.preventDefault()} onMouseEnter={()=>setHovered(item)} onFocus={()=>setHovered(item)} onClick={()=>choose(item.name)}>
      <span className="formula-menu-fx">fx</span><span className="formula-function-row-copy"><b>{item.name}</b><small>{item.syntax}</small></span>
     </button>)}
     {!filtered.length&&<div className="formula-function-empty">No matching functions.</div>}
    </div>
    <div className="formula-function-description">{hovered?.description??"Select a function to insert its name into the active cell."}</div>
    {autoSum&&<button type="button" className="formula-function-open-library" onMouseDown={event=>event.preventDefault()} onClick={()=>{close();onMore?.();}}>More Functions...</button>}
   </>}
  </FloatingPortal>
 </div>;
}
