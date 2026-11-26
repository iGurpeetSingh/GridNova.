import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { ChevronDown, FilePlus2, RotateCcw, Scissors, Trash2 } from "lucide-react";
import type { SheetPageSetup } from "../types";

interface Props {
  settings:SheetPageSetup;
  activeCell:string;
  hasPrintArea:boolean;
  hasBreakAtActive:boolean;
  hasAnyBreaks:boolean;
  onOpenPageSetup:()=>void;
  onSetPrintArea:()=>void;
  onClearPrintArea:()=>void;
  onInsertBreak:()=>void;
  onRemoveBreak:()=>void;
  onResetBreaks:()=>void;
  onToggleRtl:()=>void;
}
type MenuKey="print"|"breaks";
function PageSetupGlyph(){return <svg viewBox="0 0 32 32" aria-hidden="true"><path d="M7 2.75h12l6 6V28H7z" fill="#fff" stroke="#323130" strokeWidth="1.45" strokeLinejoin="round"/><path d="M19 2.9V9h6" fill="none" stroke="#323130" strokeWidth="1.35"/><path d="M11 13h9M11 16h9" stroke="#8A8886" strokeWidth="1.15" strokeLinecap="round"/><g transform="translate(1 15)"><path d="M4 3.5h13a2 2 0 0 1 2 2V13H2V5.5a2 2 0 0 1 2-2Z" fill="#323130"/><path d="M5 1h11v5H5z" fill="#fff" stroke="#323130" strokeWidth="1.2"/><path d="M5 9h9v7H5z" fill="#fff" stroke="#323130" strokeWidth="1.1"/><path d="M15.5 6h1" stroke="#fff" strokeWidth="1.2" strokeLinecap="round"/></g></svg>}
function PrintAreaGlyph(){return <svg viewBox="0 0 32 32" aria-hidden="true"><path d="M6.5 2.75h12l6 6V28H6.5z" fill="#fff" stroke="#323130" strokeWidth="1.45" strokeLinejoin="round"/><path d="M18.5 2.9V9h6" fill="none" stroke="#323130" strokeWidth="1.35"/><rect x="9.5" y="11.5" width="11.5" height="12" rx=".5" fill="none" stroke="#0F6CBD" strokeWidth="1.65"/><path d="M11 14h8.5M11 16.5h8.5" stroke="#8A8886" strokeWidth="1" strokeLinecap="round"/><g transform="translate(15 19)"><path d="M2.5 3h10a1.6 1.6 0 0 1 1.6 1.6V10H1V4.6A1.6 1.6 0 0 1 2.5 3Z" fill="#323130"/><path d="M4 1h7v4H4z" fill="#fff" stroke="#323130" strokeWidth="1.1"/><path d="M4 7h7v5H4z" fill="#fff" stroke="#323130" strokeWidth="1"/></g></svg>}
function BreaksGlyph(){return <svg viewBox="0 0 32 32" aria-hidden="true"><path d="M7 2.75h12l6 6V28H7z" fill="#fff" stroke="#323130" strokeWidth="1.45" strokeLinejoin="round"/><path d="M19 2.9V9h6" fill="none" stroke="#323130" strokeWidth="1.35"/><path d="M10 12h12M10 17h12M10 22h12" stroke="#A19F9D" strokeWidth="1" strokeLinecap="round"/><path d="M7.8 16h16.5" stroke="#0F6CBD" strokeWidth="1.9" strokeDasharray="2.4 1.9" strokeLinecap="round"/></svg>}
function RtlGlyph(){return <svg viewBox="0 0 32 32" aria-hidden="true"><path d="M6.5 2.75h12l6 6V28h-18z" fill="#fff" stroke="#323130" strokeWidth="1.45" strokeLinejoin="round"/><path d="M18.5 2.9V9h6" fill="none" stroke="#323130" strokeWidth="1.35"/><path d="M21 14H10m0 0 4-4m-4 4 4 4" fill="none" stroke="#0F6CBD" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/><path d="M20 21h-9" stroke="#A19F9D" strokeWidth="1.15" strokeLinecap="round"/></svg>}
export default function PageLayoutRibbon({settings,activeCell,hasPrintArea,hasBreakAtActive,hasAnyBreaks,onOpenPageSetup,onSetPrintArea,onClearPrintArea,onInsertBreak,onRemoveBreak,onResetBreaks}:Props){
  const [openMenu,setOpenMenu]=useState<MenuKey|null>(null);
  const printAnchor=useRef<HTMLButtonElement>(null),breakAnchor=useRef<HTMLButtonElement>(null),panel=useRef<HTMLDivElement>(null);
  const [position,setPosition]=useState({left:8,top:8});
  const anchor=openMenu==="print"?printAnchor:breakAnchor;
  useLayoutEffect(()=>{
    if(!openMenu)return;
    const place=()=>{
      const trigger=anchor.current,menu=panel.current;if(!trigger||!menu)return;
      const rect=trigger.getBoundingClientRect(),box=menu.getBoundingClientRect(),margin=8;
      const left=Math.max(margin,Math.min(rect.left,window.innerWidth-box.width-margin));
      const below=rect.bottom+5,above=rect.top-box.height-5;
      const top=below+box.height<=window.innerHeight-margin?below:Math.max(margin,above);
      setPosition({left,top});
    };
    place();window.addEventListener("resize",place);window.addEventListener("scroll",place,true);
    return()=>{window.removeEventListener("resize",place);window.removeEventListener("scroll",place,true)};
  },[openMenu]);
  useEffect(()=>{
    if(!openMenu)return;
    const outside=(event:PointerEvent)=>{
      const target=event.target as Node;
      if(!panel.current?.contains(target)&&!printAnchor.current?.contains(target)&&!breakAnchor.current?.contains(target))setOpenMenu(null);
    };
    const escape=(event:KeyboardEvent)=>{if(event.key==="Escape")setOpenMenu(null)};
    document.addEventListener("pointerdown",outside);document.addEventListener("keydown",escape);
    return()=>{document.removeEventListener("pointerdown",outside);document.removeEventListener("keydown",escape)};
  },[openMenu]);
  const run=(fn:()=>void)=>{fn();setOpenMenu(null)};
  const toggle=(key:MenuKey)=>setOpenMenu(current=>current===key?null:key);
  const menuItems=openMenu==="print"?[
    {label:"Set Print Area",description:"Use the currently selected range",icon:<FilePlus2 size={17}/>,disabled:false,run:onSetPrintArea},
    {label:"Clear Print Area",description:hasPrintArea?"Remove the saved print range":"No print area is set",icon:<Trash2 size={17}/>,disabled:!hasPrintArea,run:onClearPrintArea}
  ]:[
    {label:"Insert Page Break",description:"Add a break above and to the left of the active cell",icon:<Scissors size={17}/>,disabled:false,run:onInsertBreak},
    {label:"Remove Page Break",description:"Remove the active row and column break",icon:<Trash2 size={17}/>,disabled:!hasBreakAtActive,run:onRemoveBreak},
    {label:"Reset All Page Breaks",description:"Remove every manual break on this sheet",icon:<RotateCcw size={17}/>,disabled:!hasAnyBreaks,run:onResetBreaks}
  ];
  return <div className="page-layout-ribbon">
    <section className="page-layout-group page-layout-group-setup" aria-label="Page Setup">
      <div className="page-layout-group-content">
        <button type="button" className="page-layout-command" title="Page Setup" aria-label="Page Setup" onMouseDown={event=>event.preventDefault()} onClick={onOpenPageSetup}>
          <span className="page-layout-glyph"><PageSetupGlyph/></span><span className="page-layout-command-label">Page<br/>Setup</span>
        </button>
        <button ref={printAnchor} type="button" className="page-layout-command" title="Print Area" aria-label="Print Area menu" aria-haspopup="menu" aria-expanded={openMenu==="print"} onMouseDown={event=>event.preventDefault()} onClick={()=>toggle("print")}>
          <span className="page-layout-glyph"><PrintAreaGlyph/></span><span className="page-layout-command-label">Print<br/>Area <ChevronDown size={10}/></span>
        </button>
        <button ref={breakAnchor} type="button" className="page-layout-command" title="Breaks" aria-label="Breaks menu" aria-haspopup="menu" aria-expanded={openMenu==="breaks"} onMouseDown={event=>event.preventDefault()} onClick={()=>toggle("breaks")}>
          <span className="page-layout-glyph"><BreaksGlyph/></span><span className="page-layout-command-label page-layout-command-label-break">Breaks <ChevronDown size={10}/></span>
        </button>
      </div>
      <div className="page-layout-group-caption">Page Setup</div>
    </section>
    <section className="page-layout-group page-layout-group-sheet" aria-label="Sheet Options">
      <div className="page-layout-group-content">
        <button type="button" className={`page-layout-command page-layout-rtl-command ${settings.isRtl?"is-active":""}`} title={settings.isRtl?"Switch to Left-to-Right":"Switch to Right-to-Left"} aria-label="Sheet Right-to-Left" aria-pressed={settings.isRtl} onMouseDown={event=>event.preventDefault()} onClick={onToggleRtl}>
          <span className="page-layout-glyph"><RtlGlyph/></span><span className="page-layout-command-label">Sheet<br/>Right-to-Left</span>
        </button>
      </div>
      <div className="page-layout-group-caption">Sheet Options</div>
    </section>
    {openMenu&&createPortal(<div ref={panel} role="menu" aria-label={openMenu==="print"?"Print Area":"Page Breaks"} className="page-layout-menu" style={{left:position.left,top:position.top}}>
      {menuItems.map(item=><button key={item.label} type="button" role="menuitem" disabled={item.disabled} className="page-layout-menu-item" onMouseDown={event=>event.preventDefault()} onClick={()=>run(item.run)}>
        <span className="page-layout-menu-icon">{item.icon}</span><span className="page-layout-menu-copy"><strong>{item.label}</strong><small>{item.description}</small></span>
      </button>)}
    </div>,document.body)}
  </div>;
}
