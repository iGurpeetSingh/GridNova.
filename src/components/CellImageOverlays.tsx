import { useLayoutEffect, useState } from "react";
import { createPortal } from "react-dom";
import type { SheetData } from "../types";
export default function CellImageOverlays({sheet}:{sheet:SheetData}){
 const [targets,setTargets]=useState<Map<string,HTMLElement>>(new Map());
 useLayoutEffect(()=>{
  const next=new Map<string,HTMLElement>();
  Object.entries(sheet.cells).forEach(([address,cell])=>{
   if(!cell.image)return;
   const node=document.querySelector<HTMLElement>(`[data-cell="${address}"]`);
   if(node)next.set(address,node);
  });
  setTargets(next);
 },[sheet.id,sheet.cells]);
 return <>{Object.entries(sheet.cells).map(([address,cell])=>{
  const node=targets.get(address);
  return cell.image&&node?createPortal(<img src={cell.image} alt="Image in cell" className="grid-cell-image" draggable={false}/>,node,address):null;
 })}</>;
}
