import { useEffect, useState } from "react";
import { Grip, Maximize2, Trash2, X } from "lucide-react";
import type { FloatingElementData } from "../types";
import OfficeShape, { officeShapeName } from "./OfficeShape";
interface Props { element: FloatingElementData; onUpdate:(id:string,changes:Partial<FloatingElementData>)=>void; onDelete:(id:string)=>void; }
export default function FloatingElement({element,onUpdate,onDelete}:Props) {
  const [moving,setMoving]=useState(false),[resizing,setResizing]=useState(false);
  useEffect(()=>{
    const move=(event:MouseEvent)=>{
      if(moving)onUpdate(element.id,{x:Math.max(0,event.clientX-element.width/2),y:Math.max(70,event.clientY-18)});
      if(resizing)onUpdate(element.id,{width:Math.max(90,element.width+event.movementX),height:Math.max(56,element.height+event.movementY)});
    };
    const stop=()=>{setMoving(false);setResizing(false);};
    window.addEventListener("mousemove",move);window.addEventListener("mouseup",stop);
    return()=>{window.removeEventListener("mousemove",move);window.removeEventListener("mouseup",stop);};
  },[moving,resizing,element.id,element.width,element.height,onUpdate]);
  const content=element.type==="image"
    ? <img src={element.src} alt={element.text||"Inserted picture"} className="h-full w-full object-contain" draggable={false}/>
    : element.type==="text"
      ? <textarea aria-label="Text box" className="floating-element-text" value={element.text??"Text"} onChange={event=>onUpdate(element.id,{text:event.target.value})} style={{fontSize:element.fontSize??16,color:element.color??"#323130"}}/>
      : <OfficeShape shape={element.shape} fill={element.fill} stroke={element.stroke}/>;
  return <section className="floating-element-card" style={{left:element.x,top:element.y,width:element.width,height:element.height}}>
    <header className="floating-element-toolbar" onMouseDown={event=>{event.preventDefault();setMoving(true);}}><Grip size={12}/><span>{element.type==="image"?"Picture":element.type==="text"?"Text Box":officeShapeName(element.shape)}</span><button type="button" title="Delete element" aria-label="Delete element" onMouseDown={event=>event.preventDefault()} onClick={()=>onDelete(element.id)}><Trash2 size={12}/><X size={0}/></button></header>
    <div className="floating-element-content">{content}</div>
    <button type="button" aria-label="Resize element" className="floating-element-resize" onMouseDown={event=>{event.preventDefault();setResizing(true);}}><Maximize2 size={11}/></button>
  </section>;
}
