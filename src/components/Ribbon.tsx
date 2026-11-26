import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import FloatingPortal from "./FloatingPortal";
import InsertRibbonIcon from "./InsertRibbonIcon";
import { ChartGalleryButton, EditingModeDropdown, InsertIconDropdown, ShapeGalleryButton } from "./InsertDropdownGallery";
import HomeRibbon from "./HomeRibbon";
import FormulaRibbon from "./FormulaRibbon";
import FileMenu from "./FileMenu";
import PageLayoutRibbon from "./PageLayoutRibbon";
import type { CalculationMode } from "../types";
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownAZ, ArrowDownZA, ArrowDownToLine,
  ArrowUpToLine, Bold, Braces, CalendarDays, ChartColumn, ChartLine, Check,
  Clipboard, Copy, Eraser, FileDown, FilePlus2, FileSpreadsheet, FileText, Filter,
  FunctionSquare, Calculator, Grid2X2, Highlighter, Italic, Link, ListFilter, Minus, PaintBucket,
  Paintbrush, Percent, Plus, Printer, Redo2, Scissors, Search, Sigma, SortAsc,
  Strikethrough, Table2, Text, Trash2, Underline, Undo2, Upload, WrapText,
  X, ZoomIn, ZoomOut, Eye, EyeOff, Rows3, Columns3, SquareCheck, MessageSquare,
  CalendarClock, AlignJustify, Palette, ArrowLeft, Save,
  ChevronDown, ChevronUp, ChevronRight, IndentDecrease, IndentIncrease, Pencil,
  AlignVerticalJustifyStart, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd,
  Replace, Paintbrush2, Share2, Link2, Users, ShieldCheck, TableProperties, FilterX, ListChecks,
} from "lucide-react";
import type { CellStyle, ExcelTableData, SheetData, SheetPageSetup, WorkbookData } from "../types";
import { parseAddress } from "../types";
import { TABLE_STYLES } from "../utils/tableUtils";

type Tab = "File"|"Home"|"Insert"|"Share"|"Page Layout"|"Formulas"|"Data"|"Review"|"View"|"Help"|"Draw"|"Table Design";
interface Props {
  activeTab:Tab; setActiveTab:(t:Tab)=>void; onAction:(action:string)=>void;
  onFormat:(style:Partial<CellStyle>)=>void; style:CellStyle; sheet:SheetData;
  showFormulas:boolean; setShowFormulas:(v:boolean)=>void; zoom:number; setZoom:(v:number)=>void;
  pageSetup:SheetPageSetup; activeCell:string; workbooks:WorkbookData[]; currentWorkbookId:string; currentSheetProtected:boolean; currentWorkbookProtected:boolean;
  activeTable?: ExcelTableData | null;
  calculationMode?:CalculationMode;onCalculationModeChange:(mode:CalculationMode)=>void;
  onOpenNameManager:()=>void;onOpenFunctionLibrary:()=>void;
}
const baseTabs:Tab[]=["File","Home","Insert","Share","Page Layout","Formulas","Data","Review","View","Help","Draw"];
const tones:Record<string,string>={green:"#107c41",blue:"#3b82f6",indigo:"#6366f1",amber:"#e59a23",purple:"#8b5cf6",rose:"#e45a56",pink:"#e66aa8"};
function Tool({label,icon:Icon,action,onAction,tone="blue",active=false}:any){
  return <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction(action)} title={label} aria-label={label} className={`ribbon-action ${active?"active":""}`}>
    <Icon size={16} strokeWidth={1.7} style={{color:tone==="green"?"#107c41":tone==="amber"?"#c88100":tone==="rose"?"#c43b32":"#323130"}}/>
  </button>;
}
function Drop({label,value,options,onPick,open,onToggle,onClose}:{label:string;value:string;options:string[];onPick:(v:string)=>void;open:boolean;onToggle:()=>void;onClose:()=>void}){
  const anchorRef=useRef<HTMLButtonElement>(null);
  return <div className={`ribbon-dropdown-wrap ribbon-dropdown-${label.toLowerCase().replaceAll(" ","-")}`}><button ref={anchorRef} type="button" aria-label={label} title={label} aria-haspopup="menu" aria-expanded={open} onMouseDown={event=>event.preventDefault()} onClick={onToggle} className="ribbon-dropdown"><span>{value||label}</span><ChevronDown size={10}/></button>
    <FloatingPortal open={open} anchorRef={anchorRef} onClose={onClose} role="menu" ariaLabel={label} className="ribbon-menu-panel">{options.map(option=><button type="button" key={option} onMouseDown={event=>event.preventDefault()} onClick={()=>{onPick(option);onClose()}} className={`ribbon-menu-option ${value===option?"selected":""}`}>{option}</button>)}</FloatingPortal>
  </div>;
}
function ColorPicker({kind,style,onFormat,open,onToggle,onClose}:{kind:"text"|"fill";style:CellStyle;onFormat:(style:Partial<CellStyle>)=>void;open:boolean;onToggle:()=>void;onClose:()=>void}){
  const anchorRef=useRef<HTMLButtonElement>(null);
  const swatches=["#202124","#107c41","#10b981","#3b82f6","#6366f1","#8b5cf6","#f59e0b","#ef4444","#ec4899","#ffffff","#dbeafe","#dcfce7","#fef3c7","#fee2e2","#f3e8ff"];
  return <div className="ribbon-color-picker"><button ref={anchorRef} type="button" aria-label={kind==="text"?"Font Color":"Fill Color"} title={kind==="text"?"Font Color":"Fill Color"} aria-haspopup="menu" aria-expanded={open} onMouseDown={event=>event.preventDefault()} onClick={onToggle} className="ribbon-action ribbon-color-action"><span className="flex flex-col items-center"><span className="text-[14px] font-semibold">{kind==="text"?"A":<PaintBucket size={15}/>}</span><span className="h-[3px] w-4" style={{background:kind==="text"?style.color??"#c43b32":style.backgroundColor??"#ffc000"}}/></span><ChevronDown size={9}/></button>
    <FloatingPortal open={open} anchorRef={anchorRef} onClose={onClose} role="group" ariaLabel={kind==="text"?"Font colors":"Fill colors"} className="ribbon-color-menu"><div className="ribbon-color-swatches">{swatches.map(color=><button type="button" key={color} aria-label={color} title={color} onMouseDown={event=>event.preventDefault()} onClick={()=>{onFormat(kind==="text"?{color}:{backgroundColor:color});onClose()}} className="ribbon-color-swatch" style={{background:color}}/>)}</div></FloatingPortal>
  </div>;
}
function ActionMenu({id,label,className,children,items,open,onToggle,onClose,onAction}:{id:string;label:string;className:string;children:ReactNode;items:{label:string;action?:string;onSelect?:()=>void}[];open:boolean;onToggle:()=>void;onClose:()=>void;onAction:(action:string)=>void}){
  const anchorRef=useRef<HTMLButtonElement>(null);
  return <div className="ribbon-action-menu"><button ref={anchorRef} type="button" title={label} aria-label={label} aria-haspopup="menu" aria-expanded={open} onMouseDown={event=>event.preventDefault()} onClick={onToggle} className={className}>{children}</button>
    <FloatingPortal open={open} anchorRef={anchorRef} onClose={onClose} role="menu" ariaLabel={label} className="ribbon-menu-panel">{items.map(item=><button type="button" role="menuitem" key={item.label} onMouseDown={event=>event.preventDefault()} onClick={()=>{if(item.onSelect)item.onSelect();else if(item.action)onAction(item.action);onClose()}} className="ribbon-menu-option">{item.label}</button>)}</FloatingPortal>
  </div>;
}
function Group({label,children}:{label:string;children:ReactNode}){
  return <section aria-label={label} className="ribbon-group"><div className="ribbon-group-content">{children}</div><div className="ribbon-label">{label}</div></section>;
}
function InsertGroup({label,children}:{label:string;children:ReactNode}){
  return <section aria-label={label} className="insert-ribbon-group"><div className="insert-ribbon-group-content">{children}</div><div className="insert-ribbon-caption">{label}</div></section>;
}
export default function Ribbon({activeTab,setActiveTab,onAction,onFormat,style,sheet,showFormulas,setShowFormulas,zoom,setZoom,pageSetup,activeCell,workbooks,currentWorkbookId,currentSheetProtected,currentWorkbookProtected,activeTable=null,calculationMode="Automatic",onCalculationModeChange,onOpenNameManager,onOpenFunctionLibrary}:Props){
  const tabs = activeTable ? [...baseTabs, "Table Design" as Tab] : baseTabs;
  const [openMenu,setOpenMenu]=useState<string|null>(null);
  const [collapsed,setCollapsed]=useState(false);
  const [editingMode,setEditingMode]=useState<"Editing"|"Reviewing"|"Viewing">("Editing");
  const toggleMenu=(key:string)=>setOpenMenu(current=>current===key?null:key);
  const closeMenu=()=>setOpenMenu(null);
  useEffect(()=>{setOpenMenu(null);},[activeTab]);
  const button=(label:string,icon:any,id:string,tone="blue",active=false)=><Tool key={id} label={label} icon={icon} action={id} onAction={onAction} tone={tone} active={active}/>;
  const colorPicker=(kind:"text"|"fill")=><ColorPicker key={kind} kind={kind} style={style} onFormat={onFormat} open={openMenu===kind} onToggle={()=>toggleMenu(kind)} onClose={closeMenu}/>;
  const drop=(id:string,props:{label:string;value:string;options:string[];onPick:(v:string)=>void})=><Drop key={id} {...props} open={openMenu===id} onToggle={()=>toggleMenu(id)} onClose={closeMenu}/>;
  const content=()=>{
    switch(activeTab){
      case "File":return null;
      case "Home":return <HomeRibbon onAction={onAction} onFormat={onFormat} style={style} sheet={sheet}/>;
      case "Insert":return <>
        <InsertGroup label="Tables">
          <InsertIconDropdown kind="pivot" open={openMenu==="insert-pivot"} onToggle={()=>toggleMenu("insert-pivot")} onClose={closeMenu} onAction={onAction}><InsertRibbonIcon name="pivot"/><span className="insert-command-label">PivotTable</span><ChevronDown className="insert-command-chevron" size={10}/></InsertIconDropdown>
          <button type="button" className="insert-command" title="Create a formatted Excel Table" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("format-table")}><InsertRibbonIcon name="table"/><span className="insert-command-label">Table</span></button>
          <InsertIconDropdown kind="forms" open={openMenu==="insert-forms"} onToggle={()=>toggleMenu("insert-forms")} onClose={closeMenu} onAction={onAction}><InsertRibbonIcon name="forms"/><span className="insert-command-label">Forms</span><ChevronDown className="insert-command-chevron" size={10}/></InsertIconDropdown>
        </InsertGroup>
        <InsertGroup label="Illustrations">
          <InsertIconDropdown kind="pictures" open={openMenu==="insert-pictures"} onToggle={()=>toggleMenu("insert-pictures")} onClose={closeMenu} onAction={onAction}><InsertRibbonIcon name="pictures"/><span className="insert-command-label">Pictures</span><ChevronDown className="insert-command-chevron" size={10}/></InsertIconDropdown>
          <ShapeGalleryButton onSelect={shape=>onAction(`shape-gallery:${shape}`)}><InsertRibbonIcon name="shapes"/><span className="insert-command-label">Shapes</span><ChevronDown className="insert-command-chevron" size={10}/></ShapeGalleryButton>
        </InsertGroup>
        <InsertGroup label="Controls">
          <button type="button" className="insert-command" title="Insert an interactive checkbox into the selected cells" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("checkbox")}><InsertRibbonIcon name="checkbox"/><span className="insert-command-label">Checkbox</span></button>
        </InsertGroup>
        <InsertGroup label="Charts">
          {([
            ["column","Column"],["line","Line"],["scatter","Scatter"],["pie","Pie"],
            ["bar","Bar"],["statistical","Statistical"],["combo","Combo"],["treemap","Other Charts"]
          ] as const).map(([kind,label])=><ChartGalleryButton key={kind} kind={kind} label={label} onSelect={onAction}><InsertRibbonIcon name={kind==="treemap"?"other":kind}/><span className={kind==="treemap"?"insert-command-label insert-command-label--two-word":"insert-command-label"}>{label}</span><ChevronDown className="insert-command-chevron" size={10}/></ChartGalleryButton>)}
        </InsertGroup>
        <InsertGroup label="Filters">
          <button type="button" className="insert-command" disabled={!activeTable} title={activeTable?"Insert a table slicer":"Select a Table or PivotTable to insert a Slicer"} onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("slicer")}><InsertRibbonIcon name="slicer"/><span className="insert-command-label">Slicer</span></button>
        </InsertGroup>
        <InsertGroup label="Links">
          <button type="button" className="insert-command" title="Insert a hyperlink into the active cell" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("hyperlink")}><InsertRibbonIcon name="link"/><span className="insert-command-label">Link</span></button>
        </InsertGroup>
        <InsertGroup label="Comments">
          <button type="button" className="insert-command" title="Add a comment to the active cell" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("comment")}><InsertRibbonIcon name="comment"/><span className="insert-command-label insert-command-label--two-word">New Comment</span></button>
        </InsertGroup>
        <InsertGroup label="Text">
          <button type="button" className="insert-command" title="Insert a text box over the grid" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("text-box")}><InsertRibbonIcon name="text"/><span className="insert-command-label insert-command-label--two-word">Text Box</span></button>
        </InsertGroup>
      </>;
      case "Share":return <>
        <Group label="Share Workbook">{button("Share workbook",Share2,"share-open","green")}{button("Copy workbook link",Link2,"share-copy-link","blue")}</Group>
        <Group label="Share Selection">{button("Create range link",Link2,"share-section-link","blue")}{button("Share selected table",Table2,"share-table-link","blue")}</Group>
        <Group label="Collaboration">{button("Manage access",ShieldCheck,"share-manage-access","blue")}{button("Invite people",Users,"share-open","green")}</Group>
      </>;
      case "Table Design":return activeTable ? <>
        <Group label="Properties">
          <button type="button" className="ribbon-labeled" title="Rename Table" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("table-rename")}><TableProperties size={16}/>{activeTable.name}<ChevronDown size={9}/></button>
          {button("Resize Table",Rows3,"table-resize","blue")}
        </Group>
        <Group label="Tools">{button("Remove Duplicates",Eraser,"table-remove-duplicates","rose")}{button("Convert to Range",Grid2X2,"table-convert-range","blue")}{button("Clear Filters",FilterX,"table-clear-filters","blue")}</Group>
        <Group label="Table Style Options">
          {button("Header Row",Table2,"table-toggle-headers","green",activeTable.hasHeaders)}
          {button("Total Row",Sigma,"table-toggle-totals","green",activeTable.showTotalRow)}
          {button("First Column",Columns3,"table-toggle-first-column","blue",activeTable.firstColumn)}
          {button("Last Column",Columns3,"table-toggle-last-column","blue",activeTable.lastColumn)}
          {button("Banded Rows",Rows3,"table-toggle-banded-rows","blue",activeTable.bandedRows)}
          {button("Banded Columns",Columns3,"table-toggle-banded-columns","blue",activeTable.bandedColumns)}
          {button("Filter Button",Filter,"table-toggle-filter-buttons","blue",activeTable.showFilterButtons)}
        </Group>
        <Group label="Table Styles">
          <ActionMenu id="table-styles" label="Table Styles Gallery" className="ribbon-labeled" items={TABLE_STYLES.map(id=>({label:id.replace("TableStyle","").replace(/([A-Z])/g," $1").trim(),action:`table-style-${id}`}))} open={openMenu==="table-styles"} onToggle={()=>toggleMenu("table-styles")} onClose={closeMenu} onAction={onAction}><Palette size={16}/>Table Styles <ChevronDown size={9}/></ActionMenu>
          {button("Clear Style",Eraser,"table-clear-style","rose")}
        </Group>
      </> : <Group label="Table Design">{button("Create Table",Table2,"format-table","green")}</Group>;
      case "Review":return <>
        <Group label="Comments">{button("New Comment",MessageSquare,"comment","blue")}{button("Previous Comment",ArrowUpToLine,"comment-previous","blue")}{button("Next Comment",ArrowDownToLine,"comment-next","blue")}{button("Delete Comment",Trash2,"comment-delete","rose")}{button("Show Comments",MessageSquare,"comments-panel","blue")}</Group>
        <Group label="Proofing & Stats">{button("Workbook Statistics",Calculator,"workbook-stats","blue")}{button("Protect Sheet",ShieldCheck,"protect-sheet","blue")}</Group>
      </>;
      case "Help":return <Group label="Help">{button("Keyboard shortcuts",FunctionSquare,"show-shortcuts","blue")}{button("Function help",Search,"insert-function","blue")}</Group>;
      case "Draw":return <Group label="Drawing">{button("Add comment",MessageSquare,"comment","blue")}</Group>;
      case "Page Layout":{
        const position=parseAddress(activeCell)??{row:0,col:0};
        const hasBreakAtActive=pageSetup.rowBreaks.includes(position.row)||pageSetup.columnBreaks.includes(position.col);
        const hasAnyBreaks=pageSetup.rowBreaks.length>0||pageSetup.columnBreaks.length>0;
        return <PageLayoutRibbon settings={pageSetup} hasPrintArea={!!pageSetup.printArea} hasBreakAtActive={hasBreakAtActive} hasAnyBreaks={hasAnyBreaks} onAction={onAction}/>;
      }
      case "Formulas":return <FormulaRibbon onAction={onAction} showFormulas={showFormulas} setShowFormulas={setShowFormulas} calculationMode={calculationMode} onCalculationModeChange={onCalculationModeChange} onOpenNameManager={onOpenNameManager} onOpenFunctionLibrary={onOpenFunctionLibrary}/>;
      case "Data":return <>
        <Group label="Import">{button("Import CSV / Excel",Upload,"import","green")}</Group>
        <Group label="Sort & filter">{button("Sort A–Z",SortAsc,"sort-az","blue")}{button("Sort Z–A",ArrowDownZA,"sort-za","blue")}{button(sheet.filters?"Remove filters":"Add filters",Filter,"filters","blue",sheet.filters)}{button("Remove duplicates",Eraser,"duplicates","rose")}</Group>
        <Group label="Data Tools">{button("Text to Columns",Columns3,"text-to-columns","blue")}{button("Data Validation",ListChecks,"data-validation","green")}{button("Clear Validation",Eraser,"clear-validation","rose")}</Group>
      </>;
      case "View":return <>
        <Group label="Zoom"><Tool label="Zoom out" icon={ZoomOut} action="zoom-out" onAction={()=>setZoom(zoom-10)}/><span className="px-2 text-xs font-semibold text-[#526458]">{zoom}%</span><Tool label="Zoom in" icon={ZoomIn} action="zoom-in" onAction={()=>setZoom(zoom+10)}/>{drop("zoom",{label:"Zoom",value:`${zoom}% options`,options:["50%","75%","100%","125%","150%","175%","200%"],onPick:v=>setZoom(Number(v.replace("%","")))})}</Group>
        <Group label="Window">{button(sheet.freezeTopRow?"Unfreeze top row":"Freeze top row",Rows3,"freeze","blue",sheet.freezeTopRow)}</Group>
      </>;
    }
  };
  return <div className="no-print ribbon-shell relative z-40 shrink-0 bg-[#F5F5F5] px-2 pt-1" onMouseDown={event=>{if((event.target as HTMLElement).closest("button"))event.preventDefault();}}>
    <div className="ribbon-card overflow-visible">
      <div className="ribbon-tabbar">
        <nav className="ribbon-tabs" aria-label="Workbook tabs">{tabs.map(tab=>tab==="File"?<FileMenu key="File" workbooks={workbooks} currentWorkbookId={currentWorkbookId} currentSheetProtected={currentSheetProtected} currentWorkbookProtected={currentWorkbookProtected} onAction={onAction}/>:<button key={tab} onMouseDown={e=>e.preventDefault()} onClick={()=>setActiveTab(tab)} className={`ribbon-tab ${activeTab===tab?"is-active":""}`}>{tab}{activeTab===tab&&<span className="ribbon-tab-indicator"/>}</button>)}</nav>
        <div className="ribbon-utilities">
          <button type="button" className="ribbon-utility-button" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("comments-panel")}><MessageSquare size={15}/>Comments</button>
          <EditingModeDropdown open={openMenu==="editing-mode"} onToggle={()=>toggleMenu("editing-mode")} onClose={closeMenu} mode={editingMode} onModeChange={setEditingMode}><Pencil size={14}/>{editingMode}<ChevronDown size={12}/></EditingModeDropdown>
        </div>
      </div>
      {!collapsed&&<div className="ribbon-scroll flex items-stretch">{content()}</div>}
      <button type="button" className="ribbon-collapse-button" aria-label={collapsed?"Expand ribbon":"Collapse ribbon"} title={collapsed?"Expand ribbon":"Collapse ribbon"} onMouseDown={e=>e.preventDefault()} onClick={()=>setCollapsed(value=>!value)}>{collapsed?<ChevronDown size={13}/>:<ChevronUp size={13}/>}</button>
    </div>
  </div>;
}
