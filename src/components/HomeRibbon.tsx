import { useState, type ReactNode } from "react";
import {
  AlignCenter,AlignLeft,AlignRight,AlignVerticalJustifyCenter,AlignVerticalJustifyEnd,AlignVerticalJustifyStart,
  AppWindow,ArrowDownAZ,ArrowDownToLine,ArrowDownZA,ArrowUpToLine,Bold,Braces,Calculator,ChevronDown,Clipboard,
  Columns3,Copy,Eraser,EyeOff,Filter,Grid2X2,IndentDecrease,IndentIncrease,Italic,ListChecks,Minus,Paintbrush2,
  PaintBucket,Palette,Percent,Plus,Redo2,Replace,Rows3,Scissors,Search,Sigma,SquareCheck,Strikethrough,
  Table2,Text,Trash2,Underline,Undo2,WrapText,X,Pencil,AlignJustify
} from "lucide-react";
import type { CellStyle, SheetData } from "../types";
import InsertRibbonIcon from "./InsertRibbonIcon";
import { HomeColorMenu,HomeMenu,HomeSelectMenu,HomeStackAction } from "./HomeRibbonControls";
import { TABLE_STYLES } from "../utils/tableUtils";
interface Props { onAction:(action:string)=>void; onFormat:(style:Partial<CellStyle>)=>void; style:CellStyle; sheet:SheetData; }
function Group({label,children}:{label:string;children:ReactNode}){return <section className="home-ribbon-group" aria-label={label}><div className="home-ribbon-group-content">{children}</div><div className="home-ribbon-caption">{label}</div></section>}
function IconAction({label,icon,action,onAction,active=false}:{label:string;icon:ReactNode;action:string;onAction:(a:string)=>void;active?:boolean}){return <button type="button" title={label} aria-label={label} className={`home-icon-action ${active?"is-active":""}`} onMouseDown={e=>e.preventDefault()} onClick={()=>onAction(action)}>{icon}</button>}
export default function HomeRibbon({onAction,onFormat,style}:Props){
 const [open,setOpen]=useState<string|null>(null),toggle=(k:string)=>setOpen(v=>v===k?null:k),close=()=>setOpen(null),fmt=(s:Partial<CellStyle>)=>onFormat(s);
 const menu=(key:string,label:string,items:Parameters<typeof HomeMenu>[0]["items"],children:ReactNode,buttonClassName:string)=> <HomeMenu label={label} open={open===key} onToggle={()=>toggle(key)} onClose={close} onAction={onAction} items={items} buttonClassName={buttonClassName}>{children}</HomeMenu>;
 return <>
 <Group label="Undo"><div className="home-undo-stack"><IconAction label="Undo" icon={<Undo2 size={17}/>} action="undo" onAction={onAction}/><IconAction label="Redo" icon={<Redo2 size={17}/>} action="redo" onAction={onAction}/></div></Group>
 <Group label="Clipboard">
  {menu("paste","Paste",[
   {label:"Paste",icon:<Clipboard size={18} color="#C88100"/>,action:"paste"},
   {label:"Paste Values",icon:<Text size={18} color="#0F6CBD"/>,action:"paste-values"},
   {label:"Paste Formulas",icon:<Braces size={18} color="#107C41"/>,action:"paste-formulas"},
   {label:"Paste Formatting",icon:<Paintbrush2 size={18} color="#C88100"/>,action:"paste-formats"}
  ],<><span className="home-paste-art"><Clipboard size={27} fill="#F8D36B" color="#C88100"/></span><span>Paste</span><ChevronDown size={10}/></>,"home-paste-trigger")}
  <div className="home-clipboard-stack"><HomeStackAction label="Cut" icon={<Scissors size={16}/>} action="cut" onAction={onAction}/><HomeStackAction label="Copy" icon={<Copy size={16}/>} action="copy" onAction={onAction}/><HomeStackAction label="Format Painter" icon={<Paintbrush2 size={16} color="#C88100"/>} action="format-painter" onAction={onAction}/></div>
 </Group>
 <Group label="Font">
  <div className="home-font-layout">
   <div className="home-font-row home-font-row--top">
    <HomeSelectMenu label="Font Family" value={style.fontFamily??"Calibri (Body)"} options={["Calibri (Body)","Arial","Inter","Georgia","Courier New","Times New Roman"]} icon={<Text size={14}/>} open={open==="font-family"} onToggle={()=>toggle("font-family")} onClose={close} onPick={v=>fmt({fontFamily:v==="Calibri (Body)"?"Calibri":v})} buttonClassName="home-font-family-select"/>
    <HomeSelectMenu label="Font Size" value={String(style.fontSize??11)} options={["8","9","10","11","12","14","16","18","20","22","24","26","28","36","48","72"]} icon={<Text size={14}/>} open={open==="font-size"} onToggle={()=>toggle("font-size")} onClose={close} onPick={v=>fmt({fontSize:Number(v)})} buttonClassName="home-font-size-select"/>
    <button type="button" title="Increase Font Size" className="home-font-size-step" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("font-up")}><b>A</b><i>⌃</i></button>
    <button type="button" title="Decrease Font Size" className="home-font-size-step home-font-size-step--small" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("font-down")}><b>A</b><i>⌄</i></button>
   </div>
   <div className="home-font-row home-font-row--bottom">
    <IconAction label="Bold" icon={<Bold size={16}/>} action="bold" onAction={onAction} active={!!style.bold}/>
    <IconAction label="Italic" icon={<Italic size={16}/>} action="italic" onAction={onAction} active={!!style.italic}/>
    <IconAction label="Underline" icon={<Underline size={16}/>} action="underline" onAction={onAction} active={!!style.underline}/>
    <IconAction label="Double Underline" icon={<span className="home-double-u">D</span>} action="double-underline" onAction={onAction} active={!!style.doubleUnderline}/>
    <IconAction label="Strikethrough" icon={<Strikethrough size={16}/>} action="strike" onAction={onAction} active={!!style.strike}/>
    {menu("borders","Borders",[
     {label:"Bottom Border",icon:<span className="home-border-icon home-border-icon--bottom"/>,onSelect:()=>fmt({borderBottom:"#323130"})},
     {label:"Top Border",icon:<span className="home-border-icon home-border-icon--top"/>,onSelect:()=>fmt({borderTop:"#323130"})},
     {label:"Left Border",icon:<span className="home-border-icon home-border-icon--left"/>,onSelect:()=>fmt({borderLeft:"#323130"})},
     {label:"Right Border",icon:<span className="home-border-icon home-border-icon--right"/>,onSelect:()=>fmt({borderRight:"#323130"})},
     {label:"All Borders",icon:<Grid2X2 size={18}/>,onSelect:()=>fmt({border:"all"})},
     {label:"Outside Borders",icon:<SquareCheck size={18}/>,onSelect:()=>fmt({border:"outer"})},
     {label:"No Border",icon:<Grid2X2 size={18}/>,onSelect:()=>fmt({border:"none",borderTop:"",borderRight:"",borderBottom:"",borderLeft:""})}
    ],<><Grid2X2 size={17}/><ChevronDown size={9}/></>,"home-small-menu-trigger")}
    <HomeColorMenu kind="fill" current={style.backgroundColor} open={open==="fill-color"} onToggle={()=>toggle("fill-color")} onClose={close} onFormat={fmt}><span className="home-color-glyph"><PaintBucket size={17}/><i style={{background:style.backgroundColor??"#FFC000"}}/></span><ChevronDown size={9}/></HomeColorMenu>
    <HomeColorMenu kind="font" current={style.color} open={open==="font-color"} onToggle={()=>toggle("font-color")} onClose={close} onFormat={fmt}><span className="home-font-color-glyph">A<i style={{background:style.color??"#C00000"}}/></span><ChevronDown size={9}/></HomeColorMenu>
   </div>
  </div>
 </Group>
 <Group label="Alignment">
  <div className="home-align-icons">
   <div className="home-align-row"><IconAction label="Top Align" icon={<AlignVerticalJustifyStart size={16}/>} action="valign-top" onAction={onAction} active={style.verticalAlign==="top"}/><IconAction label="Middle Align" icon={<AlignVerticalJustifyCenter size={16}/>} action="valign-middle" onAction={onAction} active={!style.verticalAlign||style.verticalAlign==="middle"}/><IconAction label="Bottom Align" icon={<AlignVerticalJustifyEnd size={16}/>} action="valign-bottom" onAction={onAction} active={style.verticalAlign==="bottom"}/></div>
   <div className="home-align-row"><IconAction label="Align Left" icon={<AlignLeft size={16}/>} action="align-left" onAction={onAction} active={!style.align||style.align==="left"}/><IconAction label="Align Center" icon={<AlignCenter size={16}/>} action="align-center" onAction={onAction} active={style.align==="center"}/><IconAction label="Align Right" icon={<AlignRight size={16}/>} action="align-right" onAction={onAction} active={style.align==="right"}/></div>
   <div className="home-align-row"><IconAction label="Decrease Indent" icon={<IndentDecrease size={16}/>} action="indent-down" onAction={onAction}/><IconAction label="Increase Indent" icon={<IndentIncrease size={16}/>} action="indent-up" onAction={onAction}/>
    {menu("orientation","Text Orientation",[
     {label:"Horizontal",icon:<Text size={17}/>,action:"orientation-horizontal"},
     {label:"Vertical Text",icon:<AlignJustify size={17}/>,action:"orientation-vertical"},
     {label:"Rotate Text Up",icon:<ArrowUpToLine size={17}/>,action:"orientation-rotate-up"}
    ],<><span className="home-orientation-mark">ab↗</span><ChevronDown size={9}/></>,"home-small-menu-trigger")}
   </div>
  </div>
  <span className="home-alignment-divider"/>
  <div className="home-align-actions">
   <button type="button" className={`home-labeled-action ${style.wrap?"active":""}`} title="Wrap Text" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("wrap")}><WrapText size={17} color="#0F6CBD"/>Wrap Text</button>
   {menu("merge","Merge & Center",[
    {label:"Merge & Center",icon:<Grid2X2 size={17} color="#0F6CBD"/>,action:"merge"},
    {label:"Merge Across",icon:<Table2 size={17}/>,action:"merge"},
    {label:"Merge Cells",icon:<Grid2X2 size={17}/>,action:"merge"},
    {label:"Unmerge Cells",icon:<Grid2X2 size={17}/>,action:"unmerge"}
   ],<><Grid2X2 size={16} color="#0F6CBD"/>Merge &amp; Center<ChevronDown size={9}/></>,"home-labeled-menu-trigger")}
  </div>
 </Group>
 <Group label="Number">
  <div className="home-number-layout">
   <HomeSelectMenu label="Number Format" value={style.numberFormat??"General"} options={["General","Number","Currency ₹","Currency $","Percentage %","Date"]} icon={<Calculator size={14}/>} open={open==="number-format"} onToggle={()=>toggle("number-format")} onClose={close} onPick={v=>fmt({numberFormat:v as CellStyle["numberFormat"]})} buttonClassName="home-number-select"/>
   <div className="home-number-row">
    {menu("currency","Accounting Number Format",[
     {label:"Indian Rupee (₹)",icon:<span className="home-account-icon">₹</span>,action:"currency-inr"},
     {label:"US Dollar ($)",icon:<span className="home-account-icon">$</span>,action:"currency-usd"},
     {label:"General",icon:<Calculator size={17}/>,action:"number-general"}
    ],<><span className="home-account-icon">₹</span><ChevronDown size={9}/></>,"home-small-menu-trigger")}
    <IconAction label="Percent Style" icon={<Percent size={16}/>} action="percent" onAction={onAction}/>
    <button type="button" title="Comma Style" className="home-icon-action home-comma-action" onMouseDown={e=>e.preventDefault()} onClick={()=>onAction("comma")}>,</button>
    <IconAction label="Increase Decimal" icon={<span className="home-decimal-icon">.0<span>←</span>.00</span>} action="decimal-up" onAction={onAction}/>
    <IconAction label="Decrease Decimal" icon={<span className="home-decimal-icon">.00<span>→</span>.0</span>} action="decimal-down" onAction={onAction}/>
   </div>
  </div>
 </Group>
 <Group label="Styles">
  <div className="home-large-action-row">
   {menu("conditional","Conditional Formatting",[
    {label:"Highlight Positive Values",icon:<span className="home-data-bars"><i/><i/><i/></span>,action:"conditional-positive"},
    {label:"Highlight Negative Values",icon:<span className="home-data-bars home-data-bars--red"><i/><i/><i/></span>,action:"conditional-negative"},
    {label:"Clear Conditional Formatting",icon:<Eraser size={17}/>,action:"conditional-clear"}
   ],<><span className="home-large-icon home-conditional-icon"><Table2 size={24}/><i/></span><span>Conditional<br/>Formatting</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
   {menu("format-table","Format As Table",[
    {label:"Format Selection as Table",icon:<Table2 size={18} color="#107C41"/>,action:"format-table"},
    ...TABLE_STYLES.slice(0,8).map(id=>({label:id.replace("TableStyle","").replace(/([A-Z])/g," $1").trim(),icon:<Table2 size={17} color="#0F6CBD"/>,action:`table-style-${id}`}))
   ],<><span className="home-large-icon"><InsertRibbonIcon name="table"/></span><span>Format As Table</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
   {menu("cell-styles","Cell Styles",[
    {label:"Normal",icon:<Table2 size={17}/>,action:"cell-style-normal"},
    {label:"Good",icon:<span className="home-style-chip home-style-chip--good">Good</span>,action:"cell-style-good"},
    {label:"Bad",icon:<span className="home-style-chip home-style-chip--bad">Bad</span>,action:"cell-style-bad"},
    {label:"Neutral",icon:<span className="home-style-chip home-style-chip--neutral">Neutral</span>,action:"cell-style-neutral"},
    {label:"Heading",icon:<span className="home-style-chip home-style-chip--heading">Aa</span>,action:"cell-style-heading"}
   ],<><span className="home-large-icon"><Palette size={24} color="#038387"/></span><span>Cell Styles</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
  </div>
 </Group>
 <Group label="Cells">
  <div className="home-large-action-row">
   {menu("insert-cells","Insert Cells",[
    {label:"Insert Sheet Rows",icon:<Rows3 size={18} color="#107C41"/>,action:"row-above"},
    {label:"Insert Sheet Columns",icon:<Columns3 size={18} color="#107C41"/>,action:"col-left"}
   ],<><span className="home-large-icon"><span className="home-grid-badge"><Grid2X2 size={23}/><Plus size={11}/></span></span><span>Insert</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
   {menu("delete-cells","Delete Cells",[
    {label:"Delete Sheet Rows",icon:<Rows3 size={18} color="#C4314B"/>,action:"delete-row"},
    {label:"Delete Sheet Columns",icon:<Columns3 size={18} color="#C4314B"/>,action:"delete-col"}
   ],<><span className="home-large-icon"><span className="home-grid-badge home-grid-badge--delete"><Grid2X2 size={23}/><X size={11}/></span></span><span>Delete</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
   {menu("format-cells","Format Cells",[
    {label:"AutoFit Column Width",icon:<Columns3 size={18}/>,action:"auto-fit-column"},
    {label:"AutoFit Row Height",icon:<Rows3 size={18}/>,action:"auto-fit-row"},
    {label:"Hide Rows",icon:<EyeOff size={18}/>,action:"hide-row"},
    {label:"Hide Columns",icon:<EyeOff size={18}/>,action:"hide-col"}
   ],<><span className="home-large-icon"><span className="home-grid-badge home-grid-badge--format"><Grid2X2 size={23}/><Pencil size={10}/></span></span><span>Format</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
  </div>
 </Group>
 <Group label="Editing">
  <div className="home-editing-stack">
   {menu("autosum","AutoSum",[
    {label:"Sum",icon:<Sigma size={18} color="#107C41"/>,action:"autosum"},
    {label:"Average",icon:<Calculator size={17}/>,action:"formula-average"},
    {label:"Count Numbers",icon:<ListChecks size={17}/>,action:"formula-count"},
    {label:"Maximum",icon:<ArrowUpToLine size={17}/>,action:"formula-max"},
    {label:"Minimum",icon:<ArrowDownToLine size={17}/>,action:"formula-min"}
   ],<><Sigma size={17} color="#107C41"/>AutoSum<ChevronDown size={9}/></>,"home-editing-menu-trigger")}
   {menu("clear","Clear",[
    {label:"Clear All",icon:<Eraser size={17} color="#8B5CF6"/>,action:"clear-all"},
    {label:"Clear Contents",icon:<Trash2 size={17}/>,action:"clear"},
    {label:"Clear Formats",icon:<Paintbrush2 size={17}/>,action:"clear-format"}
   ],<><Eraser size={16} color="#8B5CF6"/>Clear<ChevronDown size={9}/></>,"home-editing-menu-trigger")}
  </div>
  {menu("sort-filter","Sort & Filter",[
   {label:"Sort A to Z",icon:<ArrowDownAZ size={18}/>,action:"sort-az"},
   {label:"Sort Z to A",icon:<ArrowDownZA size={18}/>,action:"sort-za"},
   {label:"Filter",icon:<Filter size={18}/>,action:"filters"}
  ],<><span className="home-large-icon"><ArrowDownAZ size={22}/><Filter size={12}/></span><span>Sort &amp; Filter</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
  {menu("find-select","Find & Select",[
   {label:"Find",icon:<Search size={18}/>,action:"find"},
   {label:"Find & Replace",icon:<Replace size={18}/>,action:"find"},
   {label:"Go To Cell",icon:<Grid2X2 size={18}/>,action:"find"}
  ],<><span className="home-large-icon"><Search size={23}/></span><span>Find &amp; Select</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
 </Group>
 <Group label="Add-ins">
  {menu("addins","Add-ins",[
   {label:"Get Add-ins",icon:<AppWindow size={18} color="#E46A52"/>,action:"addin-store"},
   {label:"My Add-ins",icon:<Grid2X2 size={18} color="#E46A52"/>,action:"addin-store"}
  ],<><span className="home-large-icon home-addin-icon"><AppWindow size={24}/></span><span>Add-ins</span><ChevronDown size={9}/></>,"home-large-menu-trigger")}
 </Group>
 </>;
}
