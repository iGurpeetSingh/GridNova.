import { useState } from "react";
import { createPortal } from "react-dom";
import { LineChart, X, BarChart3 } from "lucide-react";
import type { CellRange, SparklineData } from "../types";
interface Props { open:boolean; type:SparklineData["type"]; activeCell:string; onClose:()=>void; onSave:(dataRange:CellRange,locationRange:CellRange,type:SparklineData["type"])=>void; }
export default function SparklineDialog({open,type,activeCell,onClose,onSave}:Props){
  const [dataRange,setDataRange]=useState("B2:G2"),[locationRange,setLocationRange]=useState(activeCell);
  if(!open)return null;
  const parse=(value:string):CellRange|null=>{
    const parts=value.trim().toUpperCase().split(":");const a=parts[0]?.match(/^([A-Z]+)([1-9]\d*)$/),b=(parts[1]??parts[0])?.match(/^([A-Z]+)([1-9]\d*)$/);
    if(!a||!b)return null;
    const col=(letters:string)=>{let result=0;for(const ch of letters)result=result*26+ch.charCodeAt(0)-64;return result-1;};
    const startRow=Number(a[2])-1,startCol=col(a[1]),endRow=Number(b[2])-1,endCol=col(b[1]);
    return {startRow:Math.min(startRow,endRow),startCol:Math.min(startCol,endCol),endRow:Math.max(startRow,endRow),endCol:Math.max(startCol,endCol)};
  };
  const save=()=>{const data=parse(dataRange),location=parse(locationRange);if(!data||!location)return;onSave(data,location,type);};
  return createPortal(<div className="advanced-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="advanced-dialog sparkline-dialog" role="dialog" aria-modal="true" aria-labelledby="spark-title" onMouseDown={event=>event.stopPropagation()}>
      <header><div className="advanced-dialog-icon"><LineChart size={19}/></div><div><h2 id="spark-title">Create {type==="line"?"Line":type==="column"?"Column":"Win/Loss"} Sparklines</h2><p>Insert compact charts directly into worksheet cells.</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="advanced-dialog-body">
        <label className="advanced-field"><span>Data Range</span><input value={dataRange} onChange={event=>setDataRange(event.target.value.toUpperCase())} placeholder="B2:G2"/></label>
        <label className="advanced-field"><span>Location Range</span><input value={locationRange} onChange={event=>setLocationRange(event.target.value.toUpperCase())} placeholder="H2 or H2:H10"/></label>
        <p className="advanced-muted">For multiple sparklines, use one source row per destination row (or one source column per destination row).</p>
      </div>
      <footer><button type="button" className="advanced-secondary" onClick={onClose}>Cancel</button><button type="button" onClick={save}>Insert Sparklines</button></footer>
    </section>
  </div>,document.body);
}
