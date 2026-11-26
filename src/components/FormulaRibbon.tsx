import { useState,type ReactNode } from "react";
import {
 Calculator,CalendarClock,Check,ChevronDown,ChevronRight,Eye,FileSpreadsheet,Search,Sigma,Tag
} from "lucide-react";
import { HomeMenu } from "./HomeRibbonControls";
import FormulaFunctionDropdown from "./FormulaFunctionDropdown";
type CalculationMode="Automatic"|"Manual";
interface Props {
 onAction:(action:string)=>void;showFormulas:boolean;setShowFormulas:(value:boolean)=>void;
 calculationMode?:CalculationMode;onCalculationModeChange:(mode:CalculationMode)=>void;onOpenNameManager:()=>void;onOpenFunctionLibrary:()=>void;
}
function Group({label,children,className=""}:{label:string;children:ReactNode;className?:string}){
 return <section className={`formula-ribbon-group ${className}`} aria-label={label}><div className="formula-ribbon-group-content">{children}</div><div className="formula-ribbon-caption">{label}</div></section>;
}
function Trigger({label,children,onClick}:{label:string;children:ReactNode;onClick:()=>void}){
 return <button type="button" className="formula-large-trigger" title={label} aria-label={label} onMouseDown={event=>event.preventDefault()} onClick={onClick}>{children}</button>;
}
function BookIcon({glyph,color="#484644"}:{glyph:ReactNode;color?:string}){
 return <span className="formula-book-icon"><svg viewBox="0 0 28 30" width="27" height="29" aria-hidden="true"><path d="M5 2.5H21.5Q24 2.5 24 5V25.5H7Q4 25.5 4 22.5V4Q4 2.5 5 2.5Z" fill="#fff" stroke="#484644" strokeWidth="1.5"/><path d="M7 25.5H24V27H7Q4 27 4 24V22.5Q4 25.5 7 25.5Z" fill="#484644"/><path d="M7 4.8V23.2" stroke="#D2D0CE" strokeWidth="1"/></svg><span style={{color}}>{glyph}</span></span>;
}
export default function FormulaRibbon({onAction,showFormulas,setShowFormulas,calculationMode="Automatic",onCalculationModeChange,onOpenNameManager,onOpenFunctionLibrary}:Props){
 const [open,setOpen]=useState<string|null>(null),toggle=(key:string)=>setOpen(value=>value===key?null:key),close=()=>setOpen(null);
 const pick=(name:string)=>onAction(name==="SUM"?"autosum":`formula-insert:${name}`);
 const menu=(key:string,label:string,items:Parameters<typeof HomeMenu>[0]["items"],children:ReactNode)=>
  <HomeMenu label={label} open={open===key} onToggle={()=>toggle(key)} onClose={close} onAction={onAction} items={items} buttonClassName="formula-large-menu-trigger" className="formula-ribbon-menu">{children}</HomeMenu>;
 return <>
  <Group label="Function Library" className="formula-function-group">
   <Trigger label="Insert Function" onClick={onOpenFunctionLibrary}><span className="formula-fx-icon">fx</span><span>Insert<br/>Function</span></Trigger>
   <span className="formula-inner-divider"/>
   <FormulaFunctionDropdown label="AutoSum" category="AutoSum" icon={<Sigma size={24}/>} caption={<>AutoSum</>} onSelect={pick} onMore={onOpenFunctionLibrary}/>
   <FormulaFunctionDropdown label="Financial" category="Financial" icon={<BookIcon glyph="$" color="#107C41"/>} caption="Financial" onSelect={pick}/>
   <FormulaFunctionDropdown label="Logical" category="Logical" icon={<BookIcon glyph="?" color="#7719AA"/>} caption="Logical" onSelect={pick}/>
   <FormulaFunctionDropdown label="Text" category="Text" icon={<BookIcon glyph="A" color="#0F6CBD"/>} caption="Text" onSelect={pick}/>
   <FormulaFunctionDropdown label="Date & Time" category="Date & Time" icon={<BookIcon glyph={<CalendarClock size={16}/>} color="#D83B01"/>} caption={<>Date &<br/>Time</>} onSelect={pick}/>
   <FormulaFunctionDropdown label="Lookup & Reference" category="Lookup & Reference" icon={<BookIcon glyph={<Search size={16}/>} color="#0F6CBD"/>} caption={<>Lookup &<br/>Reference</>} onSelect={pick}/>
   <FormulaFunctionDropdown label="Math & Trig" category="Math & Trig" icon={<BookIcon glyph="θ" color="#107C41"/>} caption={<>Math &<br/>Trig</>} onSelect={pick}/>
   <FormulaFunctionDropdown label="More Functions" category="More Functions" moreMenu icon={<BookIcon glyph="…" color="#C43E1C"/>} caption={<>More<br/>Functions</>} onSelect={pick} onMore={onOpenFunctionLibrary}/>
  </Group>
  <Group label="Name Manager" className="formula-single-group">
   <Trigger label="Name Manager" onClick={onOpenNameManager}><span className="formula-tag-icon"><Tag size={24}/><i/></span><span>Name<br/>Manager</span></Trigger>
  </Group>
  <Group label="Formula Auditing" className="formula-single-group">
   <Trigger label="Show Formulas" onClick={()=>setShowFormulas(!showFormulas)}><span className={`formula-show-icon ${showFormulas?"is-active":""}`}><span>fx</span><Eye size={10}/></span><span>Show<br/>Formulas</span></Trigger>
  </Group>
  <Group label="Calculation" className="formula-calculation-group">
   {menu("calculation-options","Calculation Options",[
    {label:"Automatic",icon:calculationMode==="Automatic"?<Check size={16} color="#107C41"/>:<span className="formula-menu-check-empty"/>,description:"Recalculate formulas when values change",onSelect:()=>{onCalculationModeChange("Automatic");close();}},
    {label:"Manual",icon:calculationMode==="Manual"?<Check size={16} color="#107C41"/>:<span className="formula-menu-check-empty"/>,description:"Recalculate only when requested",onSelect:()=>{onCalculationModeChange("Manual");close();}}
   ],<><Calculator size={25}/><span>Calculation<br/>Options</span><ChevronDown size={9}/></>)}
   <div className="formula-calculate-actions">
    <button type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>onAction("calculate-workbook")}><Calculator size={15}/><span>Calculate Workbook</span></button>
    <button type="button" onMouseDown={event=>event.preventDefault()} onClick={()=>onAction("calculate-sheet")}><span className="formula-sheet-calc"><FileSpreadsheet size={15}/><Calculator size={8}/></span><span>Calculate Sheet</span></button>
   </div>
  </Group>
 </>;
}
