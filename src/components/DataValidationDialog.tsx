import { useEffect, useState } from "react";
import { createPortal } from "react-dom";
import { X, ListChecks } from "lucide-react";
import type { CellRange, DataValidationRule, DataValidationType, DataValidationOperator } from "../types";
interface Props { open:boolean; range:CellRange; existing?:DataValidationRule; onClose:()=>void; onSave:(rule:DataValidationRule)=>void; }
const operators:DataValidationOperator[]=["between","not between","equal to","not equal to","greater than","less than","greater than or equal to","less than or equal to"];
export default function DataValidationDialog({open,range,existing,onClose,onSave}:Props){
  const [type,setType]=useState<DataValidationType>("List"),[operator,setOperator]=useState<DataValidationOperator>("between");
  const [formula1,setFormula1]=useState(""),[formula2,setFormula2]=useState(""),[listSource,setListSource]=useState("Paid, Pending, Cancelled");
  const [allowBlank,setAllowBlank]=useState(true),[showInputMessage,setShowInputMessage]=useState(false),[inputTitle,setInputTitle]=useState("Input"),[inputMessage,setInputMessage]=useState("");
  const [showErrorAlert,setShowErrorAlert]=useState(true),[errorStyle,setErrorStyle]=useState<"Stop"|"Warning"|"Information">("Stop"),[errorTitle,setErrorTitle]=useState("Invalid value"),[errorMessage,setErrorMessage]=useState("The value you entered is not valid for this cell.");
  useEffect(()=>{
    if(!open)return;
    setType(existing?.type??"List");setOperator(existing?.operator??"between");setFormula1(existing?.formula1??"");setFormula2(existing?.formula2??"");
    setListSource(existing?.listSource??"Paid, Pending, Cancelled");setAllowBlank(existing?.allowBlank??true);
    setShowInputMessage(existing?.showInputMessage??false);setInputTitle(existing?.inputTitle??"Input");setInputMessage(existing?.inputMessage??"");
    setShowErrorAlert(existing?.showErrorAlert??true);setErrorStyle(existing?.errorStyle??"Stop");setErrorTitle(existing?.errorTitle??"Invalid value");setErrorMessage(existing?.errorMessage??"The value you entered is not valid for this cell.");
  },[open,existing]);
  if(!open)return null;
  const needsLimit=type!=="List"&&type!=="Custom Formula";
  const save=()=>onSave({id:existing?.id??crypto.randomUUID(),range,type,operator:needsLimit?operator:undefined,formula1:type==="List"?undefined:formula1,formula2:operator==="between"||operator==="not between"?formula2:undefined,listSource:type==="List"?listSource:undefined,allowBlank,showInputMessage,inputTitle,inputMessage,showErrorAlert,errorStyle,errorTitle,errorMessage});
  return createPortal(<div className="advanced-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="advanced-dialog" role="dialog" aria-modal="true" aria-labelledby="validation-title" onMouseDown={event=>event.stopPropagation()}>
      <header><div className="advanced-dialog-icon"><ListChecks size={19}/></div><div><h2 id="validation-title">Data Validation</h2><p>Applies to {String.fromCharCode(65+range.startCol)}{range.startRow+1}:{String.fromCharCode(65+range.endCol)}{range.endRow+1}</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="advanced-dialog-body">
        <label className="advanced-field"><span>Allow</span><select value={type} onChange={event=>setType(event.target.value as DataValidationType)}><option>List</option><option>Whole Number</option><option>Decimal</option><option>Date</option><option>Text Length</option><option>Custom Formula</option></select></label>
        {type==="List"?<label className="advanced-field"><span>Source</span><input value={listSource} onChange={event=>setListSource(event.target.value)} placeholder="Paid, Pending, Cancelled or =F2:F6"/><small>Enter comma-separated values or a worksheet range such as =F2:F6.</small></label>:type==="Custom Formula"?<label className="advanced-field"><span>Formula</span><input value={formula1} onChange={event=>setFormula1(event.target.value)} placeholder="=A2&gt;B2"/></label>:<><label className="advanced-field"><span>Data</span><select value={operator} onChange={event=>setOperator(event.target.value as DataValidationOperator)}>{operators.map(value=><option key={value}>{value}</option>)}</select></label><label className="advanced-field"><span>{type==="Date"?"Start / minimum":"Minimum or value"}</span><input value={formula1} onChange={event=>setFormula1(event.target.value)} placeholder={type==="Date"?"2026-01-01":"0"} type={type==="Date"?"date":"text"}/></label>{(operator==="between"||operator==="not between")&&<label className="advanced-field"><span>End / maximum</span><input value={formula2} onChange={event=>setFormula2(event.target.value)} placeholder={type==="Date"?"2026-12-31":"100"} type={type==="Date"?"date":"text"}/></label>}</>}
        <label className="advanced-check"><input type="checkbox" checked={allowBlank} onChange={event=>setAllowBlank(event.target.checked)}/> Ignore blank cells</label>
        <div className="advanced-subheading">Input Message</div><label className="advanced-check"><input type="checkbox" checked={showInputMessage} onChange={event=>setShowInputMessage(event.target.checked)}/> Show input message when cell is selected</label>
        {showInputMessage&&<><label className="advanced-field"><span>Title</span><input value={inputTitle} onChange={event=>setInputTitle(event.target.value)}/></label><label className="advanced-field"><span>Message</span><textarea value={inputMessage} onChange={event=>setInputMessage(event.target.value)} rows={2}/></label></>}
        <div className="advanced-subheading">Error Alert</div><label className="advanced-check"><input type="checkbox" checked={showErrorAlert} onChange={event=>setShowErrorAlert(event.target.checked)}/> Show error alert after invalid data is entered</label>
        {showErrorAlert&&<><label className="advanced-field"><span>Style</span><select value={errorStyle} onChange={event=>setErrorStyle(event.target.value as typeof errorStyle)}><option>Stop</option><option>Warning</option><option>Information</option></select></label><label className="advanced-field"><span>Title</span><input value={errorTitle} onChange={event=>setErrorTitle(event.target.value)}/></label><label className="advanced-field"><span>Error message</span><textarea value={errorMessage} onChange={event=>setErrorMessage(event.target.value)} rows={2}/></label></>}
      </div>
      <footer><button type="button" className="advanced-secondary" onClick={onClose}>Cancel</button><button type="button" onClick={save}>Save Rule</button></footer>
    </section>
  </div>,document.body);
}
