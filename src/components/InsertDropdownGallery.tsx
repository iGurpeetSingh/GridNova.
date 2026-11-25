import { useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import {
  Check, ChevronDown, Database, Eye, FileSpreadsheet, FileText, Globe2, Image,
  ImagePlus, Images, Laptop, Link2, MessageSquare, Pencil, Table2, TextCursorInput,
} from "lucide-react";
import FloatingPortal from "./FloatingPortal";
import InsertRibbonIcon from "./InsertRibbonIcon";

type ChartKind="column"|"line"|"scatter"|"pie"|"bar"|"statistical"|"combo"|"treemap";
type GalleryItem={name:string;category:string;kind:ChartKind;variant:string;index:number};
const chartGroups:Record<ChartKind,{category:string;names:string[]}[]> = {
  column:[{category:"2-D Column",names:["Clustered Column","Stacked Column","100% Stacked Column"]},{category:"3-D Column",names:["3-D Clustered Column","3-D Stacked Column"]}],
  line:[{category:"2-D Line",names:["Line","Stacked Line","Line with Markers"]},{category:"2-D Area",names:["Area","Stacked Area"]}],
  scatter:[{category:"Scatter (X, Y)",names:["Scatter","Scatter with Smooth Lines and Markers","Scatter with Straight Lines","Bubble"]}],
  pie:[{category:"2-D Pie",names:["Pie","Pie of Pie"]},{category:"3-D & Doughnut",names:["3-D Pie","Doughnut"]}],
  bar:[{category:"2-D Bar",names:["Clustered Bar","Stacked Bar","100% Stacked Bar"]},{category:"3-D Bar",names:["3-D Clustered Bar"]}],
  statistical:[{category:"Histogram",names:["Histogram","Pareto"]},{category:"Box and Whisker",names:["Box & Whisker"]}],
  combo:[{category:"Combo",names:["Clustered Column - Line","Clustered Column - Line on Secondary Axis","Stacked Area - Clustered Column","Stacked Column - Line"]}],
  treemap:[{category:"Hierarchy & Waterfall",names:["Treemap","Sunburst","Waterfall","Funnel","Radar"]}],
};
const shapeGroups=[
  {category:"Lines",items:[["line","Line"],["line-arrow","Line Arrow"],["double-arrow","Double Arrow"],["elbow-connector","Elbow Connector"],["curved-line","Curved Line"]]},
  {category:"Rectangles",items:[["rectangle","Rectangle"],["rounded-rectangle","Rounded Rectangle"],["snip-single-corner","Snip Single Corner"]]},
  {category:"Basic Shapes",items:[["oval","Oval"],["triangle","Triangle"],["right-triangle","Right Triangle"],["diamond","Diamond"],["pentagon","Pentagon"],["hexagon","Hexagon"],["star","Star"],["heart","Heart"],["plus-cross","Plus/Cross"]]},
  {category:"Block Arrows",items:[["right-arrow","Right Arrow"],["left-arrow","Left Arrow"],["up-arrow","Up Arrow"],["down-arrow","Down Arrow"],["left-right-arrow","Left-Right Arrow"]]},
  {category:"Callouts",items:[["rect-callout","Rectangular Callout"],["rounded-rect-callout","Rounded Rectangular Callout"],["oval-callout","Oval Callout"],["cloud-callout","Cloud Callout"]]},
] as const;
const shapePaths:Record<string,string> = {
  line:"M4 24L26 6", "line-arrow":"M4 24L24 6M17 6H24V13", "double-arrow":"M5 24L24 6M5 24V17M5 24H12M24 6V13M24 6H17",
  "elbow-connector":"M5 5V22H24M5 5H13", "curved-line":"M3 24C8 2 18 2 25 22",
  rectangle:"M4 5H25V24H4Z","rounded-rectangle":"M8 4H21Q25 4 25 8V21Q25 25 21 25H8Q4 25 4 21V8Q4 4 8 4Z",
  "snip-single-corner":"M4 10L10 4H25V25H4Z", oval:"M3 14A11 10 0 1 0 25 14A11 10 0 1 0 3 14",
  triangle:"M14 3L26 25H2Z","right-triangle":"M4 4V25H25Z",diamond:"M14 2L26 14L14 26L2 14Z",
  pentagon:"M14 2L26 10L22 25H6L2 10Z",hexagon:"M8 3H20L27 14L20 25H8L1 14Z",
  star:"M14 2L17 10L26 10L19 16L22 25L14 20L6 25L9 16L2 10L11 10Z",
  heart:"M14 25C10 21 3 16 3 10A6 6 0 0 1 14 7A6 6 0 0 1 25 10C25 16 18 21 14 25Z",
  "plus-cross":"M11 3H17V11H25V17H17V25H11V17H3V11H11Z",
  "right-arrow":"M2 9H16V3L27 14L16 25V19H2Z","left-arrow":"M26 9H12V3L1 14L12 25V19H26Z",
  "up-arrow":"M9 26V12H3L14 1L25 12H19V26Z","down-arrow":"M9 2V16H3L14 27L25 16H19V2Z",
  "left-right-arrow":"M8 8H20V3L27 14L20 25V20H8V25L1 14L8 3Z",
  "rect-callout":"M3 4H26V19H15L9 26V19H3Z","rounded-rect-callout":"M8 4H21Q26 4 26 9V18Q26 22 21 22H14L8 27V22Q3 22 3 18V9Q3 4 8 4Z",
  "oval-callout":"M14 3C7 3 2 7 2 13S7 22 14 22H17L23 27V20C26 17 27 14 26 11C25 6 20 3 14 3Z",
  "cloud-callout":"M7 20Q1 19 2 13Q2 7 8 7Q10 1 17 4Q22 1 25 7Q29 8 26 14Q27 20 20 20H12L7 26Z",
};
function ShapeThumb({shape}:{shape:string}){
  const path=shapePaths[shape]??shapePaths.rectangle;
  return <svg viewBox="0 0 28 28" width="30" height="30" aria-hidden="true"><path d={path} fill={["line","line-arrow","double-arrow","elbow-connector","curved-line"].includes(shape)?"none":"#DCEEFF"} stroke="#0F6CBD" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>;
}
function ChartThumb({kind,index,name}:{kind:ChartKind;index:number;name:string}){
  const blue="#0F6CBD",light="#62ABF5",slate="#64748B";
  if(kind==="pie")return <svg viewBox="0 0 52 44" width="52" height="44"><circle cx="25" cy="22" r="15" fill="#DCEEFF" stroke={blue}/><path d={index===0?"M25 22V7A15 15 0 0 1 40 22Z":index===1?"M25 22L36 12A15 15 0 0 1 40 22Z":"M25 22L25 7A15 15 0 0 1 40 22Z"} fill={light} stroke={blue}/><path d="M25 22L14 32" stroke="#fff" strokeWidth="1.6"/></svg>;
  if(kind==="line"||kind==="scatter")return <svg viewBox="0 0 52 44" width="52" height="44"><path d="M5 5V37H48" fill="none" stroke="#CBD5E1"/>{kind==="scatter"?<g fill={blue}><circle cx="12" cy="29" r="2.5"/><circle cx="20" cy="13" r="2.5"/><circle cx="28" cy="25" r="2.5"/><circle cx="38" cy="9" r="2.5"/><circle cx="43" cy="20" r="2.5"/><circle cx="17" cy="32" r="2"/></g>:<><path d={index===1?"M7 31L18 23L28 27L39 16L47 20":"M7 28L18 20L28 24L39 11L47 14"} fill="none" stroke={blue} strokeWidth="2.2" strokeLinecap="round"/>{index!==0&&<g fill={light} stroke={blue}><circle cx="18" cy="20" r="2.2"/><circle cx="39" cy="11" r="2.2"/></g>}</>}</svg>;
  if(kind==="statistical")return <svg viewBox="0 0 52 44" width="52" height="44"><path d="M5 36H48" stroke="#CBD5E1"/><path d={index===2?"M9 21H43M9 21V31M43 21V31M18 10V32M34 10V32":"M8 35V27H16V35M16 35V18H25V35M25 35V9H34V35M34 35V20H43V35"} fill="none" stroke={blue} strokeWidth="2"/><path d="M22 7H30" stroke={slate} strokeWidth="1.5"/></svg>;
  if(kind==="treemap")return <svg viewBox="0 0 52 44" width="52" height="44"><rect x="5" y="5" width="42" height="34" rx="2" fill="#EFF6FC" stroke={blue}/><rect x="7" y="7" width="20" height="18" fill={light}/><rect x="29" y="7" width="16" height="11" fill={blue}/><rect x="29" y="20" width="16" height="17" fill="#A9CFF5"/><rect x="7" y="27" width="20" height="10" fill="#7AAFE7"/></svg>;
  return <svg viewBox="0 0 52 44" width="52" height="44"><path d="M5 37H48" stroke="#CBD5E1"/>{kind==="bar"?[0,1,2].map(i=><rect key={i} x="7" y={8+i*10} width={[34,26,39][(i+index)%3]} height="6" rx="1" fill={i%2?light:blue}/>):kind==="combo"?<><rect x="9" y="23" width="7" height="14" fill={light}/><rect x="23" y="14" width="7" height="23" fill={blue}/><rect x="37" y="8" width="7" height="29" fill={light}/><path d="M8 18L23 10L38 15L46 5" fill="none" stroke={slate} strokeWidth="2"/></>:<>{[0,1,2].map(i=><rect key={i} x={9+i*13} y={[20,11,6][(i+index)%3]} width="8" height={37-[20,11,6][(i+index)%3]} rx="1" fill={i%2?light:blue}/>)}</>}</svg>;
}
function GalleryFrame({open,anchorRef,onClose,label,children}:{open:boolean;anchorRef:React.RefObject<HTMLElement|null>;onClose:()=>void;label:string;children:ReactNode}){
  return <FloatingPortal open={open} anchorRef={anchorRef} onClose={onClose} role="dialog" ariaLabel={label} className="insert-gallery-portal">{children}</FloatingPortal>;
}
export function ChartGalleryButton({kind,label,children,onSelect}:{kind:ChartKind;label:string;children:ReactNode;onSelect:(action:string)=>void}){
  const [open,setOpen]=useState(false),[hovered,setHovered]=useState("");
  const anchor=useRef<HTMLButtonElement>(null);
  const items=useMemo(()=>{let index=0;return chartGroups[kind].map(group=>({...group,items:group.names.map(name=>({name,index:index++}))}));},[kind]);
  const first=items[0]?.items[0]?.name??"";
  const active=hovered||first;
  const choose=(index:number)=>{onSelect(`chart-gallery:${kind}:${index}`);setOpen(false);};
  return <div className="ribbon-action-menu"><button ref={anchor} type="button" title={`${label} chart types`} aria-label={`${label} chart types`} aria-haspopup="dialog" aria-expanded={open} className="insert-command" onMouseDown={e=>e.preventDefault()} onClick={()=>setOpen(v=>!v)}>{children}</button><GalleryFrame open={open} anchorRef={anchor} onClose={()=>setOpen(false)} label={`${label} chart gallery`}><div className="insert-chart-gallery">
    <header><strong>{label} Charts</strong><span>Choose a chart subtype</span></header>
    <div className="insert-chart-gallery-scroll">{items.map(group=><section key={group.category}><h3>{group.category}</h3><div className="insert-chart-gallery-grid">{group.items.map(item=><button key={item.name} type="button" title={item.name} aria-label={item.name} className="insert-chart-tile" onMouseDown={e=>e.preventDefault()} onMouseEnter={()=>setHovered(item.name)} onFocus={()=>setHovered(item.name)} onClick={()=>choose(item.index)}><ChartThumb kind={kind} index={item.index} name={item.name}/><span>{item.name}</span></button>)}</div></section>)}</div>
    <footer><span>{active}</span><button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>{onSelect("more-charts");setOpen(false)}}><FileSpreadsheet size={14}/>More Charts...</button></footer>
  </div></GalleryFrame></div>;
}
export function ShapeGalleryButton({children,onSelect}:{children:ReactNode;onSelect:(shape:string)=>void}){
  const [open,setOpen]=useState(false),anchor=useRef<HTMLButtonElement>(null),[hovered,setHovered]=useState("Rectangle");
  return <div className="ribbon-action-menu"><button ref={anchor} type="button" title="Shapes gallery" aria-label="Shapes gallery" aria-haspopup="dialog" aria-expanded={open} className="insert-command" onMouseDown={e=>e.preventDefault()} onClick={()=>setOpen(v=>!v)}>{children}</button><GalleryFrame open={open} anchorRef={anchor} onClose={()=>setOpen(false)} label="Shapes gallery"><div className="insert-shape-gallery">
    {shapeGroups.map(group=><section key={group.category}><h3>{group.category}</h3><div className="insert-shape-gallery-grid">{group.items.map(([shape,name])=><button key={shape} type="button" className="insert-shape-tile" title={name} aria-label={name} onMouseDown={e=>e.preventDefault()} onMouseEnter={()=>setHovered(name)} onFocus={()=>setHovered(name)} onClick={()=>{onSelect(shape);setOpen(false)}}><ShapeThumb shape={shape}/></button>)}</div></section>)}
    <footer><span>{hovered}</span><span>Choose a shape to insert</span></footer>
  </div></GalleryFrame></div>;
}
type IconItem={label:string;subtitle?:string;action:string;icon:ReactNode;disabled?:boolean};
function IconMenuButton({id,label,children,items,open,onToggle,onClose,onAction}:{id:string;label:string;children:ReactNode;items:IconItem[];open:boolean;onToggle:()=>void;onClose:()=>void;onAction:(action:string)=>void}){
  const anchor=useRef<HTMLButtonElement>(null);
  return <div className="ribbon-action-menu"><button ref={anchor} type="button" title={label} aria-label={label} aria-haspopup="menu" aria-expanded={open} className="insert-command" onMouseDown={e=>e.preventDefault()} onClick={onToggle}>{children}</button><FloatingPortal open={open} anchorRef={anchor} onClose={onClose} role="menu" ariaLabel={label} className="insert-icon-menu"><div className="insert-icon-menu-heading">{label}</div>{items.map(item=><button type="button" role="menuitem" key={item.action} disabled={item.disabled} onMouseDown={e=>e.preventDefault()} onClick={()=>{if(!item.disabled)onAction(item.action);onClose()}} className="insert-icon-menu-item"><span className="insert-icon-menu-art">{item.icon}</span><span className="insert-icon-menu-copy"><strong>{item.label}</strong>{item.subtitle&&<small>{item.subtitle}</small>}</span>{item.disabled&&<small className="insert-icon-menu-soon">Unavailable</small>}</button>)}</FloatingPortal></div>;
}
export function InsertIconDropdown({kind,children,open,onToggle,onClose,onAction}:{kind:"pivot"|"pictures"|"forms";children:ReactNode;open:boolean;onToggle:()=>void;onClose:()=>void;onAction:(action:string)=>void}){
  const anchor=useRef<HTMLButtonElement>(null);
  const title=kind==="pivot"?"PivotTable":kind==="pictures"?"Pictures":"Forms";
  const rows:IconItem[]=kind==="pivot"?[
    {label:"From Table/Range",subtitle:"Summarize data from selected range",action:"pivot-table",icon:<InsertRibbonIcon name="pivot"/>},
    {label:"From External Data Source",subtitle:"Connect to an external data source",action:"external-pivot",icon:<Database size={19} color="#0F6CBD"/>},
  ]:kind==="forms"?[
    {label:"New Form",subtitle:"Create a form synced to a new sheet",action:"forms-new",icon:<InsertRibbonIcon name="forms"/>},
  ]:[
    {label:"Place in Cell",subtitle:"This Device",action:"picture-cell-device",icon:<ImagePlus size={19} color="#0F6CBD"/>},
    {label:"Place in Cell",subtitle:"Stock Images",action:"picture-cell-stock",icon:<Images size={19} color="#038387"/>},
    {label:"Place over Cells",subtitle:"This Device",action:"picture-over-device",icon:<Laptop size={19} color="#0F6CBD"/>},
    {label:"Place over Cells",subtitle:"Online Pictures",action:"picture-over-online",icon:<Globe2 size={19} color="#038387"/>},
  ];
  return <div className="ribbon-action-menu"><button ref={anchor} type="button" title={title} aria-label={title} aria-haspopup="menu" aria-expanded={open} className="insert-command" onMouseDown={e=>e.preventDefault()} onClick={onToggle}>{children}</button><FloatingPortal open={open} anchorRef={anchor} onClose={onClose} role="menu" ariaLabel={title} className="insert-icon-menu"><div className="insert-icon-menu-heading">{title}</div>{kind==="pictures"?<><div className="insert-icon-menu-section">Place in Cell</div>{rows.slice(0,2).map(item=><MenuRow key={item.action} item={item} onAction={onAction} onClose={onClose}/ >)}<div className="insert-icon-menu-section">Place over Cells</div>{rows.slice(2).map(item=><MenuRow key={item.action} item={item} onAction={onAction} onClose={onClose}/>)}</>:rows.map(item=><MenuRow key={item.action} item={item} onAction={onAction} onClose={onClose}/>)}</FloatingPortal></div>;
}
function MenuRow({item,onAction,onClose}:{item:IconItem;onAction:(action:string)=>void;onClose:()=>void}){
  return <button type="button" role="menuitem" disabled={item.disabled} onMouseDown={e=>e.preventDefault()} onClick={()=>{if(!item.disabled)onAction(item.action);onClose()}} className="insert-icon-menu-item"><span className="insert-icon-menu-art">{item.icon}</span><span className="insert-icon-menu-copy"><strong>{item.subtitle&&item.label==="Place in Cell"||item.subtitle&&item.label==="Place over Cells"?item.subtitle:item.label}</strong>{item.label!=="Place in Cell"&&item.label!=="Place over Cells"&&item.subtitle&&<small>{item.subtitle}</small>}</span></button>;
}
export function EditingModeDropdown({children,open,onToggle,onClose,mode,onModeChange}:{children:ReactNode;open:boolean;onToggle:()=>void;onClose:()=>void;mode:"Editing"|"Reviewing"|"Viewing";onModeChange:(mode:"Editing"|"Reviewing"|"Viewing")=>void}){
  const anchor=useRef<HTMLButtonElement>(null);
  const items=[
    {label:"Editing" as const,subtitle:"Make changes directly",icon:<Pencil size={18} color="#107C41"/>},
    {label:"Reviewing" as const,subtitle:"Suggest changes and leave comments",icon:<MessageSquare size={18} color="#0F6CBD"/>},
    {label:"Viewing" as const,subtitle:"View only, no edits allowed",icon:<Eye size={18} color="#64748B"/>},
  ];
  return <div className="ribbon-action-menu"><button ref={anchor} type="button" title="Editing mode" aria-label="Editing mode" aria-haspopup="menu" aria-expanded={open} className="ribbon-utility-button" onMouseDown={e=>e.preventDefault()} onClick={onToggle}>{children}</button><FloatingPortal open={open} anchorRef={anchor} onClose={onClose} role="menu" ariaLabel="Editing mode" className="insert-icon-menu editing-mode-menu"><div className="insert-icon-menu-heading">Workbook mode</div>{items.map(item=><button type="button" role="menuitemradio" aria-checked={mode===item.label} key={item.label} className={`insert-icon-menu-item ${mode===item.label?"is-selected":""}`} onMouseDown={e=>e.preventDefault()} onClick={()=>{onModeChange(item.label);onClose()}}><span className="insert-icon-menu-art">{item.icon}</span><span className="insert-icon-menu-copy"><strong>{item.label}</strong><small>{item.subtitle}</small></span>{mode===item.label&&<Check size={16} className="insert-icon-menu-check"/>}</button>)}</FloatingPortal></div>;
}
