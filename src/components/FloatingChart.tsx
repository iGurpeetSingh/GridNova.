import { useEffect, useMemo, useState } from "react";
import { X, Grip, Maximize2 } from "lucide-react";
import { Bar, BarChart, CartesianGrid, Cell, Line, LineChart, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import type { FloatingChartData, SheetData } from "../types";
import { parseAddress } from "../types";
import { numericValue } from "../utils/formulaEngine";
interface Props { chart:FloatingChartData; sheet:SheetData; onUpdate:(id:string,changes:Partial<FloatingChartData>)=>void; onDelete:(id:string)=>void; }
const COLORS=["#16a05a","#4786ed","#9467df","#f39b3d","#e36b80","#16a7a1"];
function chartRows(chart:FloatingChartData,sheet:SheetData){
  const [start,end]=chart.range.split(":");const a=parseAddress(start??"A1"),b=parseAddress(end??start??"A1");
  if(!a||!b)return [];
  const rows=[];
  for(let r=a.row;r<=Math.min(b.row,sheet.rowCount-1);r++){
    const label=sheet.cells[`${String.fromCharCode(65+a.col)}${r+1}`]?.raw||`Row ${r+1}`;
    const item:Record<string,string|number>={name:label};
    for(let c=a.col+1;c<=b.col;c++){const raw=sheet.cells[`${String.fromCharCode(65+c)}${r+1}`]?.raw??"0";item[`series${c}`]=numericValue(raw,sheet.cells)??0;}
    if(b.col===a.col)item.value=numericValue(sheet.cells[`${String.fromCharCode(65+a.col)}${r+1}`]?.raw??"0",sheet.cells)??0;
    rows.push(item);
  }
  return rows;
}
export default function FloatingChart({chart,sheet,onUpdate,onDelete}:Props){
  const [moving,setMoving]=useState(false),[resizing,setResizing]=useState(false);
  const data=useMemo(()=>chartRows(chart,sheet),[chart,sheet]);
  useEffect(()=>{
    const move=(e:MouseEvent)=>{if(moving)onUpdate(chart.id,{x:Math.max(0,e.clientX-chart.width/2),y:Math.max(70,e.clientY-18)});if(resizing)onUpdate(chart.id,{width:Math.max(270,chart.width+e.movementX),height:Math.max(190,chart.height+e.movementY)});};
    const up=()=>{setMoving(false);setResizing(false)};
    window.addEventListener("mousemove",move);window.addEventListener("mouseup",up);
    return()=>{window.removeEventListener("mousemove",move);window.removeEventListener("mouseup",up)};
  },[moving,resizing,chart.id,chart.width,chart.height,onUpdate]);
  const seriesKeys=data.length?Object.keys(data[0]).filter(k=>k.startsWith("series")):[];
  return <section className="absolute z-20 overflow-hidden rounded-2xl border border-[#dfe8e1] bg-white shadow-[0_16px_48px_rgba(33,60,42,.18)]" style={{left:chart.x,top:chart.y,width:chart.width,height:chart.height}}>
    <header onMouseDown={e=>{setMoving(true);e.preventDefault()}} className="flex h-10 cursor-move items-center gap-2 border-b border-[#edf1ee] bg-[#f9fbf9] px-3"><Grip size={14} className="text-[#9aa79e]"/><span className="min-w-0 flex-1 truncate text-xs font-bold text-[#364a3b]">{chart.title}</span><button onMouseDown={e=>e.preventDefault()} onClick={()=>onDelete(chart.id)} className="rounded-lg p-1.5 text-[#9a9994] hover:bg-[#fff0ef] hover:text-[#d95650]"><X size={14}/></button></header>
    <div className="h-[calc(100%-40px)] w-full p-3 pb-5">
      {data.length===0?<div className="grid h-full place-items-center text-xs text-[#99a49d]">Select a data range to populate this chart.</div>:<ResponsiveContainer width="100%" height="100%">
        {chart.type==="pie"?<PieChart><Tooltip/><Pie data={data} dataKey="value" nameKey="name" outerRadius="75%" label>{data.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie></PieChart>:chart.type==="line"?<LineChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="#edf1ee"/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/>{seriesKeys.length?seriesKeys.map((k,i)=><Line key={k} dataKey={k} stroke={COLORS[i%COLORS.length]} strokeWidth={2} dot={false}/>):<Line dataKey="value" stroke={COLORS[0]} strokeWidth={2}/>}</LineChart>:<BarChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="#edf1ee"/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/>{seriesKeys.length?seriesKeys.map((k,i)=><Bar key={k} dataKey={k} fill={COLORS[i%COLORS.length]} radius={[4,4,0,0]}/>):<Bar dataKey="value" fill={COLORS[0]} radius={[4,4,0,0]}/>}</BarChart>}
      </ResponsiveContainer>}
    </div>
    <button aria-label="Resize chart" onMouseDown={e=>{e.preventDefault();setResizing(true)}} className="absolute bottom-0 right-0 grid h-6 w-6 cursor-se-resize place-items-center text-[#93a096]"><Maximize2 size={12}/></button>
  </section>;
}