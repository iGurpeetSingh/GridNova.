import { useState,type ReactNode } from "react";
import {
  Calculator,CalendarClock,Check,ChevronDown,Ellipsis,Eye,FileSpreadsheet,
  Search,Sigma,Tag
} from "lucide-react";
import { HomeMenu } from "./HomeRibbonControls";
type CalculationMode="Automatic"|"Manual";
interface Props {
 onAction:(action:string)=>void;showFormulas:boolean;setShowFormulas:(value:boolean)=>void;
 calculationMode?:CalculationMode;onCalculationModeChange:(mode:CalculationMode)=>void;onOpenNameManager:()=>void;
}
type Func={name:string;syntax:string;hint:string};
const functions:Record<string,Func[]> = {
 AutoSum:[
  {name:"SUM",syntax:"SUM(number1, [number2], …)",hint:"Adds values."},
  {name:"AVERAGE",syntax:"AVERAGE(number1, [number2], …)",hint:"Returns the arithmetic mean."},
  {name:"COUNT",syntax:"COUNT(value1, [value2], …)",hint:"Counts numeric values."},
  {name:"MAX",syntax:"MAX(number1, [number2], …)",hint:"Returns the largest value."},
  {name:"MIN",syntax:"MIN(number1, [number2], …)",hint:"Returns the smallest value."}
 ],
 Financial:[
  {name:"PMT",syntax:"PMT(rate, nper, pv, [fv], [type])",hint:"Calculates a periodic loan payment."},
  {name:"FV",syntax:"FV(rate, nper, pmt, [pv], [type])",hint:"Returns the future value."},
  {name:"PV",syntax:"PV(rate, nper, pmt, [fv], [type])",hint:"Returns the present value."},
  {name:"NPV",syntax:"NPV(rate, value1, [value2], …)",hint:"Returns net present value."}
 ],
 Logical:[
  {name:"IF",syntax:"IF(logical_test, value_if_true, value_if_false)",hint:"Returns a value based on a condition."},
  {name:"AND",syntax:"AND(logical1, [logical2], …)",hint:"TRUE when all conditions are true."},
  {name:"OR",syntax:"OR(logical1, [logical2], …)",hint:"TRUE when any condition is true."},
  {name:"IFERROR",syntax:"IFERROR(value, value_if_error)",hint:"Returns a fallback when a formula errors."}
 ],
 Text:[
  {name:"CONCAT",syntax:"CONCAT(text1, [text2], …)",hint:"Combines text strings."},
  {name:"LEFT",syntax:"LEFT(text, [num_chars])",hint:"Returns characters from the left."},
  {name:"RIGHT",syntax:"RIGHT(text, [num_chars])",hint:"Returns characters from the right."},
  {name:"TRIM",syntax:"TRIM(text)",hint:"Removes extra spaces."},
  {name:"UPPER",syntax:"UPPER(text)",hint:"Converts text to uppercase."},
  {name:"LOWER",syntax:"LOWER(text)",hint:"Converts text to lowercase."}
 ],
 "Date & Time":[
  {name:"TODAY",syntax:"TODAY()",hint:"Returns today's date."},
  {name:"NOW",syntax:"NOW()",hint:"Returns current date and time."},
  {name:"DATE",syntax:"DATE(year, month, day)",hint:"Creates a date value."},
  {name:"DATEDIF",syntax:"DATEDIF(start_date, end_date, unit)",hint:"Calculates the difference between dates."}
 ],
 "Lookup & Reference":[
  {name:"VLOOKUP",syntax:"VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",hint:"Looks up a value vertically."},
  {name:"XLOOKUP",syntax:"XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",hint:"Looks up a value and returns a match."},
  {name:"INDEX",syntax:"INDEX(array, row_num, [column_num])",hint:"Returns a value by position."},
  {name:"MATCH",syntax:"MATCH(lookup_value, lookup_array, [match_type])",hint:"Returns a relative position."}
 ],
 "Math & Trig":[
  {name:"SUMIF",syntax:"SUMIF(range, criteria, [sum_range])",hint:"Adds values matching a condition."},
  {name:"ROUND",syntax:"ROUND(number, num_digits)",hint:"Rounds a number."},
  {name:"ABS",syntax:"ABS(number)",hint:"Returns the absolute value."},
  {name:"POWER",syntax:"POWER(number, power)",hint:"Raises a number to a power."},
  {name:"SQRT",syntax:"SQRT(number)",hint:"Returns the square root."}
 ],
 Statistical:[
  {name:"MEDIAN",syntax:"MEDIAN(number1, [number2], …)",hint:"Returns the median."},
  {name:"STDEV.S",syntax:"STDEV.S(number1, [number2], …)",hint:"Estimates standard deviation from a sample."},
  {name:"COUNTIF",syntax:"COUNTIF(range, criteria)",hint:"Counts cells matching a condition."}
 ],
 Engineering:[
  {name:"CONVERT",syntax:"CONVERT(number, from_unit, to_unit)",hint:"Converts measurement units."},
  {name:"BIN2DEC",syntax:"BIN2DEC(number)",hint:"Converts binary to decimal."}
 ],
 Information:[
  {name:"ISBLANK",syntax:"ISBLANK(value)",hint:"Checks whether a cell is blank."},
  {name:"ISNUMBER",syntax:"ISNUMBER(value)",hint:"Checks whether a value is numeric."},
  {name:"ISTEXT",syntax:"ISTEXT(value)",hint:"Checks whether a value is text."},
  {name:"ISERROR",syntax:"ISERROR(value)",hint:"Checks for an error value."}
 ]
};
const categoryOrder=["Financial","Logical","Text","Date & Time","Lookup & Reference","Math & Trig"];
const categoryIcon:Record<string,ReactNode>={
 Financial:<span className="formula-book-glyph formula-book-glyph--green">$</span>,
 Logical:<span className="formula-book-glyph formula-book-glyph--purple">?</span>,
 Text:<span className="formula-book-glyph formula-book-glyph--blue">A</span>,
 "Date & Time":<CalendarClock size={17} color="#D83B01"/>,
 "Lookup & Reference":<Search size={17} color="#0F6CBD"/>,
 "Math & Trig":<span className="formula-book-glyph formula-book-glyph--green">θ</span>,
 "More Functions":<span className="formula-book-glyph formula-book-glyph--orange">···</span>
};
function BookIcon({kind}:{kind:string}){
 return <span className="formula-book-icon"><svg viewBox="0 0 28 30" width="27" height="29" aria-hidden="true"><path d="M5 2.5H21.5Q24 2.5 24 5V25.5H7Q4 25.5 4 22.5V4Q4 2.5 5 2.5Z" fill="#fff" stroke="#484644" strokeWidth="1.5"/><path d="M7 25.5H24V27H7Q4 27 4 24V22.5Q4 25.5 7 25.5Z" fill="#484644"/><path d="M7 4.8V23.2" stroke="#D2D0CE" strokeWidth="1"/></svg><span className={`formula-book-center formula-book-center--${kind.toLowerCase().replaceAll(" ","-").replaceAll("&","")}`}>{categoryIcon[kind]??<Ellipsis size={17}/>}</span></span>;
}
function Group({label,children,className=""}:{label:string;children:ReactNode;className?:string}){
 return <section className={`formula-ribbon-group ${className}`} aria-label={label}><div className="formula-ribbon-group-content">{children}</div><div className="formula-ribbon-caption">{label}</div></section>;
}
function Trigger({label,children,onClick}:{label:string;children:ReactNode;onClick:()=>void}){
 return <button type="button" className="formula-large-trigger" title={label} aria-label={label} onMouseDown={event=>event.preventDefault()} onClick={onClick}>{children}</button>;
}
export default function FormulaRibbon({onAction,showFormulas,setShowFormulas,calculationMode="Automatic",onCalculationModeChange,onOpenNameManager}:Props){
 const [open,setOpen]=useState<string|null>(null),toggle=(key:string)=>setOpen(value=>value===key?null:key),close=()=>setOpen(null);
 const functionMenu=(category:string)=>{
  const items=functions[category]??[];
  return items.map(fn=>({label:fn.name,description:fn.syntax,icon:<span className="formula-menu-fx">fx</span>,action:category==="AutoSum"&&fn.name==="SUM"?"autosum":`formula-insert:${fn.name}`}));
 };
 const menu=(key:string,label:string,items:Parameters<typeof HomeMenu>[0]["items"],children:ReactNode)=>
  <HomeMenu label={label} open={open===key} onToggle={()=>toggle(key)} onClose={close} onAction={onAction} items={items} buttonClassName="formula-large-menu-trigger" className="formula-ribbon-menu">{children}</HomeMenu>;
 return <>
  <Group label="Function Library" className="formula-function-group">
   <Trigger label="Insert Function" onClick={()=>onAction("insert-function")}><span className="formula-fx-icon">fx</span><span>Insert<br/>Function</span></Trigger>
   <span className="formula-inner-divider"/>
   {menu("autosum","AutoSum",functionMenu("AutoSum"),<><Sigma size={24}/><span>AutoSum</span><ChevronDown size={9}/></>)}
   {categoryOrder.map(category=>menu(category,category,functionMenu(category),<><BookIcon kind={category}/><span>{category==="Date & Time"?<>Date &<br/>Time</>:category==="Lookup & Reference"?<>Lookup &<br/>Reference</>:category==="Math & Trig"?<>Math &<br/>Trig</>:category}</span><ChevronDown size={9}/></>))}
   {menu("more-functions","More Functions",[
    {label:"Statistical",icon:<span className="formula-menu-fx">fx</span>,description:"Statistical analysis functions",action:"formula-category:Statistical"},
    {label:"Engineering",icon:<span className="formula-menu-fx">fx</span>,description:"Engineering and conversion functions",action:"formula-category:Engineering"},
    {label:"Information",icon:<span className="formula-menu-fx">fx</span>,description:"Information and error-checking functions",action:"formula-category:Information"}
   ],<><BookIcon kind="More Functions"/><span>More<br/>Functions</span><ChevronDown size={9}/></>)}
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
