import { useState } from "react";
import { createPortal } from "react-dom";
import { Calculator, Check, X } from "lucide-react";
import type { WorkbookData, SheetData } from "../types";
import { solveGoalSeek, type GoalSeekResult } from "../utils/goalSeek";
interface Props { open:boolean; workbook:WorkbookData; sheet:SheetData; activeCell:string; onClose:()=>void; onCommit:(address:string,value:number)=>void; }
export default function GoalSeekDialog({open,workbook,sheet,activeCell,onClose,onCommit}:Props){
  const [setCell,setSetCell]=useState(activeCell),[toValue,setToValue]=useState("0"),[changingCell,setChangingCell]=useState(""),[result,setResult]=useState<GoalSeekResult|null>(null),[error,setError]=useState("");
  if(!open)return null;
  const solve=()=>{
    setError("");setResult(null);
    const target=Number(toValue);
    if(!Number.isFinite(target)){setError("Enter a valid numeric target value.");return;}
    try{
      const solved=solveGoalSeek({workbook,sheet,setCell,targetValue:target,changingCell});
      if(!solved.success){setError(solved.message);setResult(solved);return;}
      setResult(solved);
    }catch(reason){setError(reason instanceof Error?reason.message:"Unable to solve this formula.");}
  };
  const reset=()=>{setResult(null);setError("");};
  return createPortal(<div className="advanced-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="advanced-dialog goal-seek-dialog" role="dialog" aria-modal="true" aria-labelledby="goal-seek-title" onMouseDown={event=>event.stopPropagation()}>
      <header><div className="advanced-dialog-icon"><Calculator size={19}/></div><div><h2 id="goal-seek-title">Goal Seek</h2><p>Find an input value that produces a target formula result.</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      {!result?<div className="advanced-dialog-body">
        <label className="advanced-field"><span>Set cell</span><input value={setCell} onChange={event=>setSetCell(event.target.value.toUpperCase())} placeholder="D10"/><small>Must contain a formula.</small></label>
        <label className="advanced-field"><span>To value</span><input value={toValue} onChange={event=>setToValue(event.target.value)} inputMode="decimal" placeholder="50000"/></label>
        <label className="advanced-field"><span>By changing cell</span><input value={changingCell} onChange={event=>setChangingCell(event.target.value.toUpperCase())} placeholder="B2"/><small>Must be a numeric input referenced by the Set cell formula.</small></label>
        {error&&<p className="advanced-error">{error}</p>}
      </div>:<div className="advanced-dialog-body">
        <div className={result.success?"goal-seek-success":"advanced-error"}>{result.success?"Solution found":result.message}</div>
        <div className="goal-seek-summary"><span>Target value</span><strong>{result.targetValue.toLocaleString()}</strong><span>Current value</span><strong>{result.currentValue.toLocaleString()}</strong><span>Found input</span><strong>{result.foundInput.toLocaleString(undefined,{maximumFractionDigits:8})}</strong><span>Iterations</span><strong>{result.iterations}</strong></div>
      </div>}
      <footer>{!result?<><button type="button" className="advanced-secondary" onClick={onClose}>Cancel</button><button type="button" onClick={solve}>Solve</button></>:<><button type="button" className="advanced-secondary" onClick={()=>{reset();}}>Restart</button><button type="button" className="advanced-secondary" onClick={onClose}>Cancel</button><button type="button" disabled={!result.success} onClick={()=>{onCommit(changingCell,result.foundInput);onClose();}}><Check size={14}/> OK</button></>}</footer>
    </section>
  </div>,document.body);
}
