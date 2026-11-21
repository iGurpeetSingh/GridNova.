import { useState } from "react";
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownAZ, ArrowDownZA, ArrowDownToLine,
  ArrowUpToLine, Bold, Braces, CalendarDays, ChartColumn, ChartLine, Check,
  Clipboard, Copy, Eraser, FileDown, FilePlus2, FileSpreadsheet, FileText, Filter,
  FunctionSquare, Grid2X2, Highlighter, Italic, Link, ListFilter, Minus, PaintBucket,
  Paintbrush, Percent, Plus, Printer, Redo2, Scissors, Search, Sigma, SortAsc,
  Strikethrough, Table2, Text, Trash2, Underline, Undo2, Upload, WrapText,
  X, ZoomIn, ZoomOut, Eye, EyeOff, Rows3, Columns3, SquareCheck, MessageSquare,
  CircleDollarSign, CalendarClock, AlignJustify, Palette, ArrowLeft, Save,
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
  return <button onMouseDown={e=>e.preventDefault()} onClick={()=>onAction(action)} title={label} className={`ribbon-action ${active?"active":""}`}>
    <Icon size={15} style={{color:tones[tone]??tones.blue}}/><span>{label}</span>
  </button>;
}
function Drop({label,value,options,onPick,tone="indigo"}:{label:string;value:string;options:string[];onPick:(v:string)=>void;tone?:string}){
  const [open,setOpen]=useState(false);
  return <div className="relative"><button onMouseDown={e=>e.preventDefault()} onClick={()=>setOpen(!open)} className="ribbon-action"><Text size={14} style={{color:tones[tone]}}/><span>{value||label}</span><span className="text-[9px] text-[#94a097]">▾</span></button>
    {open&&<div className="absolute left-0 top-9 z-50 min-w-[140px] rounded-xl border border-[#e2e9e4] bg-white p-1.5 shadow-xl">{options.map(option=><button key={option} onMouseDown={e=>e.preventDefault()} onClick={()=>{onPick(option);setOpen(false)}} className={`block w-full rounded-lg px-3 py-2 text-left text-xs hover:bg-[#eff7f1] ${value===option?"font-bold text-[#107c41]":"text-[#435448]"}`}>{option}</button>)}</div>}
  </div>;
}
function Group({label,children}:{label:string;children:React.ReactNode}){
  return <div className="ribbon-group"><div className="flex h-full flex-col justify-between"><div className="flex flex-1 flex-wrap content-center items-center gap-1">{children}</div><div className="ribbon-label">{label}</div></div></div>;
}
export default function Ribbon({activeTab,setActiveTab,onAction,onFormat,style,sheet,showFormulas,setShowFormulas,zoom,setZoom,orientation,setOrientation}:Props){
  const [colorMenu,setColorMenu]=useState<"text"|"fill"|null>(null);
  const swatches=["#202124","#107c41","#10b981","#3b82f6","#6366f1","#8b5cf6","#f59e0b","#ef4444","#ec4899","#ffffff","#dbeafe","#dcfce7","#fef3c7","#fee2e2","#f3e8ff"];
  const colorPicker=(kind:"text"|"fill")=><div className="relative"><button onMouseDown={e=>e.preventDefault()} onClick={()=>setColorMenu(colorMenu===kind?null:kind)} className="ribbon-action"><span className="flex flex-col items-center"><span className="text-[13px] font-bold" style={{color:kind==="text"?tones.indigo:tones.amber}}>A</span><span className="h-1 w-4 rounded" style={{background:kind==="text"?style.color??"#202124":style.backgroundColor??"#fff"}}/></span><span>{kind==="text"?"Text color":"Fill"}</span></button>{colorMenu===kind&&<div className="absolute left-0 top-9 z-50 grid w-[180px] grid-cols-5 gap-2 rounded-xl border border-[#e2e9e4] bg-white p-3 shadow-xl">{swatches.map(c=><button key={c} aria-label={c} onMouseDown={e=>e.preventDefault()} onClick={()=>{onFormat(kind==="text"?{color:c}:{backgroundColor:c});setColorMenu(null)}} className="h-6 w-6 rounded-md border border-black/10" style={{background:c}}/>)}</div>}</div>;
  const button=(label:string,icon:any,id:string,tone="blue",active=false)=><Tool key={id} label={label} icon={icon} action={id} onAction={onAction} tone={tone} active={active}/>;
  const content=()=>{
    switch(activeTab){
      case "File":return <>
        <Group label="Workbook">{button("New workbook",FilePlus2,"new","green")}{button("Dashboard",ArrowLeft,"back","blue")}{button("Save status",Save,"save","green")}</Group>
        <Group label="Export">{button("Excel (.xlsx)",FileSpreadsheet,"export-xlsx","green")}{button("CSV",FileDown,"export-csv","blue")}{button("PDF",FileText,"export-pdf","rose")}{button("Print",Printer,"print","rose")}</Group>
      </>;
      case "Home":return <>
        <Group label="Clipboard">{button("Undo",Undo2,"undo")}{button("Redo",Redo2,"redo")}{button("Cut",Scissors,"cut")}{button("Copy",Copy,"copy")}{button("Paste",Clipboard,"paste")}</Group>
        <Group label="Font"><Drop label="Font" value={style.fontFamily??"Inter"} options={["Inter","Arial","Calibri","Georgia","Courier New"]} onPick={v=>onFormat({fontFamily:v})}/><Drop label="Size" value={String(style.fontSize??11)} options={["10","11","12","14","16","18","20","24","28","32","36"]} onPick={v=>onFormat({fontSize:Number(v)})}/>{button("Bold",Bold,"bold","indigo",!!style.bold)}{button("Italic",Italic,"italic","indigo",!!style.italic)}{button("Underline",Underline,"underline","indigo",!!style.underline)}{button("Strike",Strikethrough,"strike","indigo",!!style.strike)}{colorPicker("text")}{colorPicker("fill")}{Drop label="Borders" value={style.border??"Borders"} options={["all","outer","none"]} onPick={v=>onFormat({border:v as CellStyle["border"]})}/></Group>
        <Group label="Alignment">{button("Left",AlignLeft,"align-left","blue",style.align==="left"||!style.align)}{button("Center",AlignCenter,"align-center","blue",style.align==="center")}{button("Right",AlignRight,"align-right","blue",style.align==="right")}{button("Wrap",WrapText,"wrap","blue",!!style.wrap)}{button("Merge",Grid2X2,"merge","blue")}</Group>
        <Group label="Number"><Drop label="Format" value={style.numberFormat??"General"} options={["General","Number","Currency ₹","Currency $","Percentage %","Date"]} onPick={v=>onFormat({numberFormat:v as CellStyle["numberFormat"]})}/>{button("%",Percent,"percent","amber")}{button("₹",CircleDollarSign,"currency-inr","amber")}{button("Comma",ListFilter,"comma","amber")}{button("− Dec",Minus,"decimal-down","amber")}{button("+ Dec",Plus,"decimal-up","amber")}</Group>
        <Group label="Editing">{button("AutoSum",Sigma,"autosum","green")}{button("Sort A–Z",ArrowDownAZ,"sort-az","blue")}{button("Sort Z–A",ArrowDownZA,"sort-za","blue")}{button("Find",Search,"find","blue")}{button("Clear",Eraser,"clear","rose")}</Group>
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
  return <div className="no-print shrink-0 border-b border-[#e1e9e3] bg-white">
    <div className="flex h-9 items-stretch border-b border-[#edf1ee] px-3">{tabs.map(tab=><button key={tab} onMouseDown={e=>e.preventDefault()} onClick={()=>setActiveTab(tab)} className={`relative min-w-[58px] px-3 text-xs font-semibold transition ${activeTab===tab?"bg-[#f6faf7] text-[#107c41]":"text-[#718078] hover:bg-[#f7faf8]"}`}>{tab}{activeTab===tab&&<span className="absolute bottom-0 left-2 right-2 h-[2px] rounded-full bg-[#107c41]"/>}</button>)}</div>
    <div className="ribbon-scroll flex min-h-[82px] items-stretch">{content()}</div>
  </div>;
}