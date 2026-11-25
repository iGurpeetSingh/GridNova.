import { useState } from "react";
import type { ReactNode } from "react";
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownAZ, ArrowDownZA, ArrowDownToLine,
  ArrowUpToLine, Bold, Braces, CalendarDays, ChartColumn, ChartLine, Check,
  Clipboard, Copy, Eraser, FileDown, FilePlus2, FileSpreadsheet, FileText, Filter,
  FunctionSquare, Grid2X2, Highlighter, Italic, Link, ListFilter, Minus, PaintBucket,
  Paintbrush, Percent, Plus, Printer, Redo2, Scissors, Search, Sigma, SortAsc,
  Strikethrough, Table2, Text, Trash2, Underline, Undo2, Upload, WrapText,
  X, ZoomIn, ZoomOut, Eye, EyeOff, Rows3, Columns3, SquareCheck, MessageSquare,
  CalendarClock, AlignJustify, Palette, ArrowLeft, Save,
  ChevronDown, ChevronUp, ChevronRight, IndentDecrease, IndentIncrease,
  AlignVerticalJustifyStart, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd,
  Replace, Paintbrush2,
} from "lucide-react";
import type { CellStyle, SheetData } from "../types";

type Tab = "File"|"Home"|"Insert"|"Page Layout"|"Formulas"|"Data"|"View";
interface Props {
  activeTab:Tab; setActiveTab:(t:Tab)=>void; onAction:(action:string)=>void;
  onFormat:(style:Partial<CellStyle>)=>void; style:CellStyle; sheet:SheetData;
  showFormulas:boolean; setShowFormulas:(v:boolean)=>void; zoom:number; setZoom:(v:number)=>void;
  orientation:"portrait"|"landscape"; setOrientation:(v:"portrait"|"landscape")=>void;
}
const tabs:Tab[]=["File","Home","Insert","Page Layout","Formulas","Data","View"];
const tones:Record<string,string>={green:"#107c41",blue:"#3b82f6",indigo:"#6366f1",amber:"#e59a23",purple:"#8b5cf6",rose:"#e45a56",pink:"#e66aa8"};
function Tool({label,icon:Icon,action,onAction,tone="blue",active=false}:any){
  return <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction(action)} title={label} aria-label={label} className={`ribbon-action ${active?"active":""}`}>
    <Icon size={16} strokeWidth={1.7} style={{color:tone==="green"?"#107c41":tone==="amber"?"#c88100":tone==="rose"?"#c43b32":"#323130"}}/>
  </button>;
}
function Drop({label,value,options,onPick,tone="indigo"}:{label:string;value:string;options:string[];onPick:(v:string)=>void;tone?:string}){
  const [open,setOpen]=useState(false);
  return <div className={`relative ribbon-dropdown-wrap ribbon-dropdown-${label.toLowerCase().replaceAll(" ","-")}`}><button aria-label={label} title={label} onMouseDown={e=>e.preventDefault()} onClick={()=>setOpen(!open)} className="ribbon-dropdown"><span>{value||label}</span><ChevronDown size={10}/></button>
    {open&&<div className="absolute left-0 top-8 z-50 min-w-[140px] rounded-md border border-[#d2d0ce] bg-white p-1 shadow-lg">{options.map(option=><button key={option} onMouseDown={e=>e.preventDefault()} onClick={()=>{onPick(option);setOpen(false)}} className={`block w-full rounded-lg px-3 py-2 text-left text-xs hover:bg-[#eff7f1] ${value===option?"font-bold text-[#107c41]":"text-[#435448]"}`}>{option}</button>)}</div>}
  </div>;
}
function Group({label,children}:{label:string;children:ReactNode}){
  return <section aria-label={label} className="ribbon-group"><div className="ribbon-group-content">{children}</div><div className="ribbon-label">{label}</div></section>;
}
export default function Ribbon({activeTab,setActiveTab,onAction,onFormat,style,sheet,showFormulas,setShowFormulas,zoom,setZoom,orientation,setOrientation}:Props){
  const [colorMenu,setColorMenu]=useState<"text"|"fill"|null>(null);
  const swatches=["#202124","#107c41","#10b981","#3b82f6","#6366f1","#8b5cf6","#f59e0b","#ef4444","#ec4899","#ffffff","#dbeafe","#dcfce7","#fef3c7","#fee2e2","#f3e8ff"];
  const colorPicker=(kind:"text"|"fill")=><div className="relative"><button aria-label={kind==="text"?"Font Color":"Fill Color"} title={kind==="text"?"Font Color":"Fill Color"} onMouseDown={e=>e.preventDefault()} onClick={()=>setColorMenu(colorMenu===kind?null:kind)} className="ribbon-action ribbon-color-action"><span className="flex flex-col items-center"><span className="text-[14px] font-semibold" style={{color:kind==="text"?"#323130":"#323130"}}>{kind==="text"?"A":<PaintBucket size={15}/>}</span><span className="h-[3px] w-4" style={{background:kind==="text"?style.color??"#c43b32":style.backgroundColor??"#ffc000"}}/></span><ChevronDown size={9}/></button>{colorMenu===kind&&<div className="absolute left-0 top-8 z-50 grid w-[180px] grid-cols-5 gap-2 rounded-md border border-[#d2d0ce] bg-white p-3 shadow-lg">{swatches.map(c=><button key={c} aria-label={c} onMouseDown={e=>e.preventDefault()} onClick={()=>{onFormat(kind==="text"?{color:c}:{backgroundColor:c});setColorMenu(null)}} className="h-6 w-6 rounded-md border border-black/10" style={{background:c}}/>)}</div>}</div>;
  const button=(label:string,icon:any,id:string,tone="blue",active=false)=><Tool key={id} label={label} icon={icon} action={id} onAction={onAction} tone={tone} active={active}/>;
  const content=()=>{
    switch(activeTab){
      case "File":return <>
        <Group label="Workbook">{button("New workbook",FilePlus2,"new","green")}{button("Dashboard",ArrowLeft,"back","blue")}{button("Save status",Save,"save","green")}</Group>
        <Group label="Export">{button("Excel (.xlsx)",FileSpreadsheet,"export-xlsx","green")}{button("CSV",FileDown,"export-csv","blue")}{button("PDF",FileText,"export-pdf","rose")}{button("Print",Printer,"print","rose")}</Group>
      </>;
      case "Home":return <>
        <Group label="Clipboard"><div className="ribbon-quick-access">{button("Undo",Undo2,"undo")}{button("Redo",Redo2,"redo")}</div><button className="ribbon-paste" title="Paste" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("paste")}><Clipboard size={22} strokeWidth={1.7}/><span>Paste <ChevronDown size={10}/></span></button><div className="ribbon-stack">{button("Cut",Scissors,"cut")}{button("Copy",Copy,"copy")}{button("Format Painter",Paintbrush,"format-painter")}</div></Group>
        <Group label="Font"><div className="ribbon-font"><div className="ribbon-font-row"><Drop label="Font" value={style.fontFamily??"Inter"} options={["Inter","Arial","Calibri","Georgia","Courier New"]} onPick={v=>onFormat({fontFamily:v})}/><Drop label="Size" value={String(style.fontSize??11)} options={["10","11","12","14","16","18","20","24","28","32","36"]} onPick={v=>onFormat({fontSize:Number(v)})}/>{button("Increase Font",Plus,"font-up")}{button("Decrease Font",Minus,"font-down")}</div><div className="ribbon-font-row">{button("Bold",Bold,"bold","blue",!!style.bold)}{button("Italic",Italic,"italic","blue",!!style.italic)}{button("Underline",Underline,"underline","blue",!!style.underline)}{button("Strikethrough",Strikethrough,"strike","blue",!!style.strike)}<span className="ribbon-micro-divider"/><Drop label="Borders" value={style.border??"Borders"} options={["all","outer","none"]} onPick={v=>onFormat({border:v as CellStyle["border"]})}/>{colorPicker("fill")}{colorPicker("text")}</div></div></Group>
        <Group label="Alignment"><div className="ribbon-alignment-icons"><div className="ribbon-icon-row">{button("Top Align",AlignVerticalJustifyStart,"valign-top")}{button("Middle Align",AlignVerticalJustifyCenter,"valign-middle")}{button("Bottom Align",AlignVerticalJustifyEnd,"valign-bottom")}</div><div className="ribbon-icon-row">{button("Align Left",AlignLeft,"align-left","blue",style.align==="left"||!style.align)}{button("Align Center",AlignCenter,"align-center","blue",style.align==="center")}{button("Align Right",AlignRight,"align-right","blue",style.align==="right")}{button("Decrease Indent",IndentDecrease,"indent-down")}{button("Increase Indent",IndentIncrease,"indent-up")}</div></div><div className="ribbon-stack ribbon-align-stack"><button className={`ribbon-labeled ${style.wrap?"active":""}`} onClick={()=>onAction("wrap")} title="Wrap Text"><WrapText size={15}/>Wrap Text</button><button className="ribbon-labeled" onClick={()=>onAction("merge")} title="Merge & Center"><Grid2X2 size={15}/>Merge &amp; Center <ChevronDown size={9}/></button></div></Group>
        <Group label="Number"><div className="ribbon-number"><Drop label="Format" value={style.numberFormat??"General"} options={["General","Number","Currency ₹","Currency $","Percentage %","Date"]} onPick={v=>onFormat({numberFormat:v as CellStyle["numberFormat"]})}/><div className="ribbon-icon-row"><button type="button" className="ribbon-action" title="Currency (₹)" aria-label="Currency (₹)" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("currency-inr")}><span className="ribbon-symbol">₹</span><ChevronDown size={9}/></button>{button("Percent",Percent,"percent","amber")}{button("Comma Style",ListFilter,"comma","amber")}{button("Increase Decimal",Plus,"decimal-up")}{button("Decrease Decimal",Minus,"decimal-down")}</div></div></Group>
        <Group label="Editing"><div className="ribbon-stack">{button("AutoSum",Sigma,"autosum","green")}<button className="ribbon-labeled" onClick={()=>onAction("clear")} title="Clear"><Eraser size={15}/>Clear <ChevronDown size={9}/></button></div><button className="ribbon-large-action" title="Sort & Filter" onClick={()=>onAction("sort-az")}><ArrowDownAZ size={19}/><span>Sort &amp; Filter</span><ChevronDown size={9}/></button><button className="ribbon-large-action" title="Find & Select" onClick={()=>onAction("find")}><Search size={19}/><span>Find &amp; Select</span><ChevronDown size={9}/></button></Group>
      </>;
      case "Insert":return <>
        <Group label="Cells">{button("Row above",ArrowUpToLine,"row-above","green")}{button("Row below",ArrowDownToLine,"row-below","green")}{button("Delete row",Rows3,"delete-row","rose")}{button("Column left",Columns3,"col-left","green")}{button("Column right",Plus,"col-right","green")}{button("Delete column",Trash2,"delete-col","rose")}</Group>
        <Group label="Tables">{button("Format as table",Table2,"format-table","green")}</Group>
        <Group label="Charts">{button("Bar chart",ChartColumn,"chart-bar","purple")}{button("Line chart",ChartLine,"chart-line","purple")}{button("Pie chart",Braces,"chart-pie","purple")}</Group>
        <Group label="Elements">{button("Hyperlink",Link,"hyperlink","pink")}{button("Comment",MessageSquare,"comment","pink")}{button("Checkbox",SquareCheck,"checkbox","pink")}{button("Date",CalendarDays,"date","pink")}</Group>
      </>;
      case "Page Layout":return <>
        <Group label="Page setup"><Tool label="Portrait" icon={FileText} action="orientation-portrait" onAction={onAction} tone="blue" active={orientation==="portrait"}/><Tool label="Landscape" icon={FileText} action="orientation-landscape" onAction={onAction} tone="blue" active={orientation==="landscape"}/></Group>
        <Group label="Sheet options">{button(sheet.showGridlines?"Hide gridlines":"Show gridlines",Grid2X2,"gridlines","green",!sheet.showGridlines)}{button(sheet.showHeadings?"Hide headings":"Show headings",Rows3,"headings","blue",!sheet.showHeadings)}</Group>
      </>;
      case "Formulas":return <>
        <Group label="Function library">{button("AutoSum",Sigma,"autosum","green")}{button("SUM",Braces,"formula-sum","amber")}{button("AVERAGE",FunctionSquare,"formula-average","amber")}{button("COUNT",FunctionSquare,"formula-count","amber")}{button("MIN",Minus,"formula-min","amber")}{button("MAX",Plus,"formula-max","amber")}{button("IF",Braces,"formula-if","amber")}</Group>
        <Group label="Formula view"><Tool label={showFormulas?"Hide formulas":"Show formulas"} icon={showFormulas?EyeOff:Eye} action="toggle-formulas" onAction={()=>setShowFormulas(!showFormulas)} tone="amber" active={showFormulas}/></Group>
      </>;
      case "Data":return <>
        <Group label="Import">{button("Import CSV / Excel",Upload,"import","green")}</Group>
        <Group label="Sort & filter">{button("Sort A–Z",SortAsc,"sort-az","blue")}{button("Sort Z–A",ArrowDownZA,"sort-za","blue")}{button(sheet.filters?"Remove filters":"Add filters",Filter,"filters","blue",sheet.filters)}{button("Remove duplicates",Eraser,"duplicates","rose")}</Group>
      </>;
      case "View":return <>
        <Group label="Zoom"><Tool label="Zoom out" icon={ZoomOut} action="zoom-out" onAction={()=>setZoom(zoom-10)}/><span className="px-2 text-xs font-semibold text-[#526458]">{zoom}%</span><Tool label="Zoom in" icon={ZoomIn} action="zoom-in" onAction={()=>setZoom(zoom+10)}/><Drop label="Zoom" value={`${zoom}% options`} options={["50%","75%","100%","125%","150%","175%","200%"]} onPick={v=>setZoom(Number(v.replace("%","")))}/></Group>
        <Group label="Window">{button(sheet.freezeTopRow?"Unfreeze top row":"Freeze top row",Rows3,"freeze","blue",sheet.freezeTopRow)}</Group>
      </>;
    }
  };
  return <div className="no-print shrink-0 bg-[#f3f2f1] px-2 pt-1">
    <div className="overflow-visible rounded-t-lg border border-[#e1dfdd] bg-white shadow-[0_1px_3px_rgba(50,49,48,0.08)]">
      <div className="flex h-[32px] items-stretch border-b border-[#e1dfdd] px-3">{tabs.map(tab=><button key={tab} onMouseDown={e=>e.preventDefault()} onClick={()=>setActiveTab(tab)} className={`relative min-w-[54px] px-3 text-[13px] font-normal transition hover:bg-[#f3f2f1] ${activeTab===tab?"font-semibold text-[#107c41]":"text-[#605e5c]"}`}>{tab}{activeTab===tab&&<span className="absolute bottom-0 left-0 right-0 h-[3px] bg-[#107c41]"/>}</button>)}</div>
      <div className="ribbon-scroll flex items-stretch">{content()}</div>
    </div>
  </div>;
}
