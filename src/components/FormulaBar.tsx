import { ChevronDown, FunctionSquare } from "lucide-react";
interface Props { address:string; value:string; onChange:(v:string)=>void; onCommit:()=>void; onFocus?:()=>void; onBlur?:()=>void; }
export default function FormulaBar({address,value,onChange,onCommit,onFocus,onBlur}:Props){
  return <div className="formula-bar no-print">
    <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-name-box" aria-label="Name box" title="Name box">{address}<ChevronDown size={11}/></button>
    <span className="formula-divider" aria-hidden="true"/>
    <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-control formula-cancel" title="Cancel" aria-label="Cancel" onClick={()=>onChange("")}>×</button>
    <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-control formula-enter" title="Enter" aria-label="Enter" onClick={onCommit}>✓</button>
    <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-fx" title="Insert Function" aria-label="Insert Function"><FunctionSquare size={15}/><i>fx</i></button>
    <input aria-label="Formula bar input" value={value} onFocus={onFocus} onBlur={onBlur} onChange={e=>onChange(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.nativeEvent.isComposing&&e.keyCode!==229)onCommit();}} className="formula-input"/>
    <button type="button" onMouseDown={event=>event.preventDefault()} className="formula-expand" title="Expand formula bar" aria-label="Expand formula bar"><ChevronDown size={12}/></button>
  </div>;
}
