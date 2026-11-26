import { createPortal } from "react-dom";
import { useEffect, useLayoutEffect, useRef, useState, type KeyboardEvent, type ReactNode, type RefObject } from "react";
import {
  ArrowLeft, ArrowRightFromLine, CalendarClock, Calculator, ChevronLeft, ChevronRight,
  CircleHelp, Copy, Download, FileClock, FileDown, FileJson, FilePlus2, FileSpreadsheet,
  FileText, Folder, FolderInput, FolderOpen, Grid2X2, Info, Keyboard, Link2, LockKeyhole,
  MessageSquare, MonitorCog, Printer, Save, Settings2, Share2, SlidersHorizontal,
  Table2, Upload, Users, Wallet, X
} from "lucide-react";
import type { WorkbookData } from "../types";

type MenuNode = {
  id:string;
  separator?:boolean;
  label:string;
  action?:string;
  children?:MenuNode[];
  icon?:ReactNode;
  disabled?:boolean;
  title?:string;
  meta?:string;
};
type MenuGroup = {id:string;items:MenuNode[]};
interface Props {
  workbooks:WorkbookData[];
  currentWorkbookId:string;
  onAction:(action:string)=>void;
}
const icon=(Icon:typeof Folder,color:string,fill?:string)=><Icon size={17} strokeWidth={1.65} style={{color,fill:fill??"none"}}/>;
function makeGroups(workbooks:WorkbookData[],currentId:string):MenuGroup[]{
  const recent=[...workbooks].sort((a,b)=>b.modifiedAt-a.modifiedAt).slice(0,8);
  return [
    {id:"section-open",items:[
      {id:"open",label:"Open",icon:icon(FolderOpen,"#D83B01"),children:[
        {id:"new-blank",label:"New Blank Workbook",action:"file-new-blank",icon:icon(FilePlus2,"#107C41")},
        {id:"templates",label:"Templates",icon:icon(Table2,"#0F6CBD"),children:[
          {id:"monthly-budget",label:"Monthly Budget",action:"file-template:budget",icon:icon(Wallet,"#107C41")},
          {id:"sales-dashboard",label:"Sales Dashboard",action:"file-template:sales-dashboard",icon:icon(FileSpreadsheet,"#107C41")},
          {id:"gst-invoice",label:"GST Invoice",action:"file-template:gst-invoice",icon:icon(FileText,"#0F6CBD")},
          {id:"attendance-tracker",label:"Attendance Tracker",action:"file-template:attendance-tracker",icon:icon(CalendarClock,"#0F6CBD")}
        ]},
        {id:"open-device",label:"Open from This Device",action:"file-open-device",icon:icon(Upload,"#0F6CBD")},
        {id:"recent",label:"Recent Workbooks",icon:icon(FileClock,"#605E5C"),children:recent.length?recent.map(book=>({id:"recent-"+book.id,label:book.title,meta:new Date(book.modifiedAt).toLocaleString(),action:"file-open-workbook:"+book.id,icon:icon(FileSpreadsheet,book.id===currentId?"#107C41":"#605E5C")})):[{id:"recent-empty",label:"No recent workbooks",disabled:true}]}
      ]}
    ]},
    {id:"section-actions",items:[
      {id:"divider-before-actions",label:"",separator:true},
      {id:"share",label:"Share",icon:icon(Share2,"#0F6CBD"),children:[
        {id:"share-dialog",label:"Share & Permissions",action:"file-share-open",icon:icon(Users,"#0F6CBD")},
        {id:"share-copy-link",label:"Copy Link",action:"file-share-copy-link",icon:icon(Link2,"#0F6CBD")},
        {id:"share-embed",label:"Embed",action:"file-embed-open",icon:icon(Copy,"#605E5C")}
      ]},
      {id:"copy",label:"Create a Copy",icon:icon(Copy,"#323130"),children:[
        {id:"save-local-copy",label:"Save a Copy Locally",action:"file-save-copy",icon:icon(Save,"#107C41")},
        {id:"download-copy",label:"Download a Copy (.xlsx)",action:"file-copy-xlsx",icon:icon(Download,"#107C41")},
        {id:"duplicate-sheet",label:"Duplicate Active Sheet",action:"file-duplicate-sheet",icon:icon(Table2,"#0F6CBD")}
      ]},
      {id:"export",label:"Export",icon:icon(ArrowRightFromLine,"#323130"),children:[
        {id:"export-xlsx",label:"Download as Excel (.xlsx)",action:"file-export-xlsx",icon:icon(FileSpreadsheet,"#107C41")},
        {id:"export-csv",label:"Download as CSV (.csv)",action:"file-export-csv",icon:icon(FileDown,"#0F6CBD")},
        {id:"export-pdf",label:"Export as PDF",action:"file-export-pdf",icon:icon(FileText,"#C43B32")},
        {id:"export-json",label:"Download JSON Backup (.json)",action:"file-export-json",icon:icon(FileJson,"#605E5C")}
      ]},
      {id:"print",label:"Print",action:"file-print",icon:icon(Printer,"#323130")}
    ]},
    {id:"section-settings",items:[
      {id:"divider-before-settings",label:"",separator:true},
      {id:"move",label:"Move File",disabled:true,title:"Available for cloud-synced workbooks",icon:icon(FolderInput,"#A19F9D")},
      {id:"info",label:"Info",icon:icon(Info,"#323130"),children:[
        {id:"history",label:"Version History",action:"file-history-open",icon:icon(FileClock,"#605E5C")},
        {id:"statistics",label:"Workbook Statistics",action:"file-workbook-stats",icon:icon(Calculator,"#0F6CBD")},
        {id:"protect-sheet",label:"Protect / Unprotect Sheet",action:"file-protect-sheet",icon:icon(LockKeyhole,"#605E5C")},
        {id:"protect-workbook",label:"Protect / Unprotect Workbook",action:"file-protect-workbook",icon:icon(LockKeyhole,"#605E5C")},
        {id:"rename",label:"Rename Workbook",action:"file-rename",icon:icon(FileText,"#605E5C")}
      ]},
      {id:"options",label:"Options",icon:icon(SlidersHorizontal,"#323130"),children:[
        {id:"regional",label:"Regional Format Settings",action:"file-regional-open",icon:icon(Wallet,"#107C41")},
        {id:"calculation",label:"Calculation Options",action:"file-calculation-open",icon:icon(Calculator,"#0F6CBD")},
        {id:"display",label:"Grid & Display",action:"file-display-open",icon:icon(MonitorCog,"#605E5C")},
        {id:"shortcuts",label:"Keyboard Shortcuts",action:"file-shortcuts-open",icon:icon(Keyboard,"#605E5C")}
      ]}
    ]}
  ];
}
function findNode(groups:MenuGroup[],path:string[]):MenuNode|undefined{
  if(!path.length)return undefined;
  const root=groups.flatMap(group=>group.items).find(item=>item.id===path[0]);
  if(!root)return undefined;
  let current=root;
  for(const segment of path.slice(1)){
    const next=current.children?.find(item=>item.id===segment);
    if(!next)return undefined;
    current=next;
  }
  return current;
}
function PortalPanel({items,parentRef,root=false,mobile=false,depth,title,pathPrefix=[],onBack,onAction,onClose,onOpenPath,onEnter,onLeave,registerButton}:{items:MenuNode[];parentRef:RefObject<HTMLElement|null>;root?:boolean;mobile?:boolean;depth:number;title?:string;pathPrefix?:string[];onBack?:()=>void;onAction:(action:string)=>void;onClose:()=>void;onOpenPath:(path:string[])=>void;onEnter:()=>void;onLeave:()=>void;registerButton:(id:string,element:HTMLButtonElement|null)=>void}){
  const panelRef=useRef<HTMLDivElement>(null);
  const [position,setPosition]=useState({left:8,top:8});
  useLayoutEffect(()=>{
    const place=()=>{
      if(mobile){setPosition({left:8,top:Math.max(8,window.innerHeight- Math.min(window.innerHeight-24,Math.max(280,items.length*37+52))-8)});return;}
      const anchor=parentRef.current,panel=panelRef.current;if(!anchor||!panel)return;
      const a=anchor.getBoundingClientRect(),p=panel.getBoundingClientRect(),gap=root?0:4;
      let left=root?a.left: a.right+gap;
      if(!root&&left+p.width>window.innerWidth-8)left=a.left-p.width-gap;
      left=Math.max(8,Math.min(left,window.innerWidth-p.width-8));
      const top=root?a.bottom+5:Math.max(8,Math.min(a.top,window.innerHeight-p.height-8));
      setPosition({left,top});
    };
    place();window.addEventListener("resize",place);window.addEventListener("scroll",place,true);
    return()=>{window.removeEventListener("resize",place);window.removeEventListener("scroll",place,true)};
  },[parentRef,root,mobile,items.length]);
  const handleKeyDown=(event:KeyboardEvent<HTMLDivElement>)=>{
    const buttons=Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>(".file-menu-item:not(:disabled)"));
    const current=buttons.indexOf(document.activeElement as HTMLButtonElement);
    if(event.key==="ArrowDown"||event.key==="ArrowUp"){
      event.preventDefault();buttons[(current+(event.key==="ArrowDown"?1:-1)+buttons.length)%buttons.length]?.focus();
    }else if(event.key==="Home"||event.key==="End"){
      event.preventDefault();(event.key==="Home"?buttons[0]:buttons.at(-1))?.focus();
    }else if(event.key==="ArrowRight"){
      const currentButton=buttons[current],id=currentButton?.dataset.menuId;
      if(id&&currentButton.dataset.hasChildren==="true"){event.preventDefault();onOpenPath([...pathPrefix,id]);}
    }else if(event.key==="Escape"){event.preventDefault();onClose();}
    else if(event.key==="ArrowLeft"){event.preventDefault();if(onBack)onBack();else onClose();}
  };
  return createPortal(<div ref={panelRef} data-file-menu-surface="true" role="menu" aria-label={root?"File menu":"File submenu"} onMouseEnter={onEnter} onMouseLeave={onLeave} onKeyDown={handleKeyDown} onMouseDown={event=>event.preventDefault()} className={`file-menu-panel ${root?"file-menu-root":""} ${mobile?"file-menu-mobile":""}`} style={{left:position.left,top:position.top}}>
    {mobile&&<div className="file-menu-mobile-header">{depth>0?<button type="button" onMouseDown={event=>event.preventDefault()} onClick={onBack}><ChevronLeft size={17}/> Back</button>:<span/>}<strong>{title??"File"}</strong><button type="button" aria-label="Close File menu" onMouseDown={event=>event.preventDefault()} onClick={onClose}><X size={17}/></button></div>}
    {items.map(item=>item.separator?<div key={item.id} role="separator" className="file-menu-divider"/>:<button key={item.id} ref={element=>registerButton(item.id,element)} type="button" role="menuitem" disabled={item.disabled} data-has-children={!!item.children} data-menu-id={item.id} title={item.title} aria-haspopup={item.children?"menu":undefined} className="file-menu-item" onMouseEnter={()=>{if(item.children){onEnter();onOpenPath([...pathPrefix,item.id]);}else if(pathPrefix.length===0)onOpenPath([]);}} onClick={()=>{if(item.children){onOpenPath([...pathPrefix,item.id]);}else if(item.action){onAction(item.action);onClose();}}}>
      <span className="file-menu-item-icon">{item.icon??<span/>}</span><span className="file-menu-item-copy"><span>{item.label}</span>{item.meta&&<small>{item.meta}</small>}</span>{item.children&&<ChevronRight size={15} className="file-menu-chevron"/>}
    </button>)}
  </div>,document.body);
}
export default function FileMenu({workbooks,currentWorkbookId,onAction}:Props){
  const [open,setOpen]=useState(false);
  const [activePath,setActivePath]=useState<string[]>([]);
  const [mobile,setMobile]=useState(typeof window!=="undefined"&&window.innerWidth<768);
  const triggerRef=useRef<HTMLButtonElement>(null);
  const rootRef=useRef<HTMLDivElement>(null);
  const anchors=useRef(new Map<string,HTMLButtonElement>());
  const closeTimer=useRef<number|null>(null);
  const groups=makeGroups(workbooks,currentWorkbookId);
  const rootItems=groups.flatMap(group=>group.items);
  const register=(path:string,element:HTMLButtonElement|null)=>{if(element)anchors.current.set(path,element);else anchors.current.delete(path);};
  const registerRoot=(id:string,element:HTMLButtonElement|null)=>register(id,element);
  const registerSub=(prefix:string)=>(id:string,element:HTMLButtonElement|null)=>register(prefix+"/"+id,element);
  const cancelClose=()=>{if(closeTimer.current!==null){window.clearTimeout(closeTimer.current);closeTimer.current=null;}};
  const scheduleClose=()=>{cancelClose();closeTimer.current=window.setTimeout(()=>{setOpen(false);setActivePath([]);},240);};
  const close=()=>{cancelClose();setOpen(false);setActivePath([]);};
  const openPath=(path:string[])=>{cancelClose();setActivePath(path);};
  const runAction=(action:string)=>{close();onAction(action);};
  useEffect(()=>{
    const resize=()=>setMobile(window.innerWidth<768);
    window.addEventListener("resize",resize);
    return()=>window.removeEventListener("resize",resize);
  },[]);
  useEffect(()=>{
    if(!open)return;
    const outside=(event:PointerEvent)=>{
      const target=event.target as Node;
      if(rootRef.current?.contains(target)||(target instanceof Element&&target.closest("[data-file-menu-surface='true']")))return;
      close();
    };
    const escape=(event:KeyboardEvent)=>{if(event.key==="Escape")close();};
    document.addEventListener("pointerdown",outside);document.addEventListener("keydown",escape);
    requestAnimationFrame(()=>document.querySelector<HTMLButtonElement>(".file-menu-root .file-menu-item")?.focus());
    return()=>{document.removeEventListener("pointerdown",outside);document.removeEventListener("keydown",escape);};
  },[open]);
  const panels:ReactNode[]=[];
  if(open){
    if(!mobile||activePath.length===0)panels.push(<PortalPanel key="root" items={rootItems} parentRef={triggerRef} root mobile={mobile&&activePath.length===0} depth={0} pathPrefix={[]} onAction={runAction} onClose={close} onOpenPath={openPath} onEnter={cancelClose} onLeave={scheduleClose} registerButton={registerRoot}/>);
    if(activePath.length){
      const paths=mobile?[activePath]:activePath.map((_,index)=>activePath.slice(0,index+1));
      paths.forEach(path=>{
        const node=findNode(groups,path);
        if(!node?.children)return;
        const parentKey=path.join("/");
        const parentRef:RefObject<HTMLElement|null>={current:anchors.current.get(parentKey)??null};
        const title=node.label;
        panels.push(<PortalPanel key={"submenu-"+parentKey} items={node.children} parentRef={parentRef} mobile={mobile} depth={path.length} title={title} pathPrefix={path} onBack={()=>openPath(path.slice(0,-1))} onAction={runAction} onClose={close} onOpenPath={openPath} onEnter={cancelClose} onLeave={scheduleClose} registerButton={registerSub(parentKey)}/>);
      });
    }
  }
  return <div ref={rootRef} className="file-menu-trigger-wrap">
    <button ref={triggerRef} type="button" className={`ribbon-tab file-menu-trigger ${open?"is-open":""}`} aria-haspopup="menu" aria-expanded={open} aria-label="File" onMouseDown={event=>event.preventDefault()} onClick={()=>{if(open)close();else{setOpen(true);setActivePath([]);}}}>File</button>
    {open&&createPortal(<>{panels}</>,document.body)}
  </div>;
}
