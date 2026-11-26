import { useEffect, useMemo, useState, type RefObject } from "react";
import { Check, FunctionSquare } from "lucide-react";
import FloatingPortal from "../components/FloatingPortal";
import { FORMULA_CATALOG, getFormula } from "../lib/formula-catalog";

export interface ArgumentInfo { functionName:string; syntax:string; parameters:string[]; activeIndex:number; activeParameter:string; }
export function getFormulaArgumentInfo(raw:string):ArgumentInfo|null {
 const stack:{index:number;commas:number}[]=[];
 let inString=false;
 for(let i=0;i<raw.length;i++){
  const ch=raw[i];
  if(ch==='"'){if(inString&&raw[i+1]==='"'){i++;continue;}inString=!inString;continue;}
  if(inString)continue;
  if(ch==="(")stack.push({index:i,commas:0});
  else if(ch===")")stack.pop();
  else if(ch===","&&stack.length)stack[stack.length-1].commas++;
 }
 const current=stack.at(-1);if(!current)return null;
 const before=raw.slice(0,current.index).match(/([A-Za-z_][A-Za-z0-9_.]*)\s*$/);
 if(!before)return null;
 const entry=getFormula(before[1]);if(!entry)return null;
 const inside=entry.syntax.slice(entry.syntax.indexOf("(")+1,-1);
 const parameters=inside.split(",").map(item=>item.trim()).filter(Boolean);
 const activeParameter=parameters[Math.min(current.commas,parameters.length-1)]??"argument";
 return {functionName:entry.name,syntax:entry.syntax,parameters,activeIndex:current.commas,activeParameter:activeParameter.replace(/^\[|\]$/g,"")};
}
interface AutoProps {
 open:boolean;query:string;anchorRef:RefObject<HTMLElement|null>;activeIndex:number;
 onActiveIndexChange:(index:number)=>void;onSelect:(name:string)=>void;onClose:()=>void;
}
export function FormulaAutocomplete({open,query,anchorRef,activeIndex,onActiveIndexChange,onSelect,onClose}:AutoProps){
 const matches=useMemo(()=>FORMULA_CATALOG.filter(item=>item.name.toLowerCase().startsWith(query.toLowerCase())).slice(0,10),[query]);
 useEffect(()=>{if(activeIndex>=matches.length)onActiveIndexChange(0);},[activeIndex,matches.length,onActiveIndexChange]);
 if(!open||!matches.length)return null;
 return <FloatingPortal open={open} anchorRef={anchorRef} onClose={onClose} role="listbox" ariaLabel="Formula autocomplete" className="formula-autocomplete-menu">
  <div className="formula-autocomplete-heading"><FunctionSquare size={13}/> Functions</div>
  <div className="formula-autocomplete-results">
   {matches.map((item,index)=><button type="button" role="option" aria-selected={index===activeIndex} key={item.name} onMouseDown={event=>event.preventDefault()} onMouseEnter={()=>onActiveIndexChange(index)} onClick={()=>onSelect(item.name)}>
    <span className="formula-menu-fx">fx</span><span><b>{item.name}</b><small>{item.syntax}</small></span>{index===activeIndex&&<Check size={13}/>}
   </button>)}
  </div>
  <div className="formula-autocomplete-footer">Tab or Enter to insert · Esc to close</div>
 </FloatingPortal>;
}
interface TooltipProps { raw:string;anchorRef:RefObject<HTMLElement|null>; }
export function FormulaArgumentTooltip({raw,anchorRef}:TooltipProps){
 const info=getFormulaArgumentInfo(raw),[dismissed,setDismissed]=useState(false);
 useEffect(()=>setDismissed(false),[raw]);
 if(!info||dismissed)return null;
 return <FloatingPortal open={!!info&&!dismissed} anchorRef={anchorRef} onClose={()=>setDismissed(true)} role="status" ariaLabel="Formula argument help" className="formula-argument-tooltip">
  <div><b>{info.functionName}</b><code>{info.syntax}</code></div>
  <p>Argument {Math.min(info.activeIndex+1,info.parameters.length)}: <strong>{info.activeParameter}</strong></p>
 </FloatingPortal>;
}
export function formulaReferenceAddresses(raw:string):Set<string>{
 const out=new Set<string>();
 const source=raw.replace(/"(?:[^"]|"")*"/g,"");
 const refs=source.match(/\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?/gi)??[];
 for(const ref of refs){
  const [a,b=a]=ref.split(":"),parse=(v:string)=>{const m=v.replace(/\$/g,"").toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);if(!m)return null;let col=0;for(const ch of m[1])col=col*26+ch.charCodeAt(0)-64;return {row:Number(m[2])-1,col:col-1};};
  const start=parse(a),end=parse(b);if(!start||!end)continue;
  for(let row=Math.min(start.row,end.row);row<=Math.max(start.row,end.row)&&row<1000;row++)
   for(let col=Math.min(start.col,end.col);col<=Math.max(start.col,end.col)&&col<1000;col++){
    let n=col+1,label="";while(n){const rem=(n-1)%26;label=String.fromCharCode(65+rem)+label;n=Math.floor((n-1)/26);}out.add(label+(row+1));
   }
 }
 return out;
}
