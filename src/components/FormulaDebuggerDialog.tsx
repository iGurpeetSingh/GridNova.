import { useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { Check, ChevronRight, RotateCcw, X } from "lucide-react";
import type { CellMap, ExcelTableData, SheetData } from "../types";
import { evaluate } from "../utils/formulaEngine";
interface Props { open:boolean; formula:string; cells:CellMap; tables:ExcelTableData[]; namedRanges:Record<string,string>; sheets:SheetData[]; sheetName:string; address:string; onClose:()=>void; }
interface Step { expression:string; name:string; start:number; end:number; }
function findSteps(formula:string):Step[]{
  const source=formula.replace(/^=/,""),steps:Step[]=[];
  for(let index=0;index<source.length;index++){
    if(source[index]==='"'){index++;while(index<source.length){if(source[index]==='"'&&source[index+1]==='"'){index+=2;continue;}if(source[index++]==='"')break;}continue;}
    const match=source.slice(index).match(/^([A-Za-z][A-Za-z0-9_.]*)\s*\(/);
    if(!match)continue;
    const start=index+match[0].lastIndexOf("(");let depth=0,quoted=false,end=start;
    for(;end<source.length;end++){
      const ch=source[end];
      if(ch==='"'){if(quoted&&source[end+1]==='"'){end++;continue;}quoted=!quoted;continue;}
      if(quoted)continue;
      if(ch==="(")depth++;else if(ch===")"&&--depth===0)break;
    }
    if(depth===0)steps.push({expression:source.slice(index,end+1),name:match[1].toUpperCase(),start:index,end:end+1});
    index=start;
  }
  const unique=new Map<string,Step>();steps.sort((a,b)=>(a.end-a.start)-(b.end-b.start)).forEach(step=>unique.set(step.expression,step));
  return [...unique.values()];
}
export default function FormulaDebuggerDialog({open,formula,cells,tables,namedRanges,sheets,sheetName,address,onClose}:Props){
  const steps=useMemo(()=>findSteps(formula),[formula]);
  const [current,setCurrent]=useState(0);
  if(!open)return null;
  const evaluateStep=(step:Step)=>evaluate("="+step.expression,cells,new Set(),tables,address,namedRanges,sheets,sheetName);
  const completed=steps.slice(0,current).map(step=>({step,value:evaluateStep(step)}));
  const finalValue=evaluate(formula,cells,new Set(),tables,address,namedRanges,sheets,sheetName);
  return createPortal(<div className="formula-debug-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="formula-debug-dialog" role="dialog" aria-modal="true" aria-labelledby="formula-debug-title" onMouseDown={event=>event.stopPropagation()}>
      <header><div><span>FORMULA AUDITING</span><h2 id="formula-debug-title">Evaluate Formula</h2><p>{address} · Step through nested functions</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="formula-debug-expression"><label>Formula</label><code>{formula||"(No formula in active cell)"}</code></div>
      <div className="formula-debug-steps">
        {steps.length?steps.map((step,index)=><div key={step.expression} className={index===current&&current<steps.length?"is-current":index<current?"is-complete":""}><span>{index+1}</span><div><strong>{step.name}</strong><code>{step.expression}</code>{index<current&&<small>Result: {String(completed[index]?.value??"")}</small>}</div>{index<current&&<Check size={15}/>}</div>):<p>No nested function calls found. The full expression is evaluated as one calculation.</p>}
      </div>
      <div className="formula-debug-result"><span>Final result</span><strong>{String(finalValue)}</strong></div>
      <footer><button type="button" className="formula-debug-secondary" onClick={()=>setCurrent(0)}><RotateCcw size={14}/> Restart</button><button type="button" onClick={()=>setCurrent(value=>Math.min(steps.length,value+1))} disabled={current>=steps.length}><ChevronRight size={14}/> Evaluate Next</button><button type="button" className="formula-debug-secondary" onClick={onClose}>Close</button></footer>
    </section>
  </div>,document.body);
}
