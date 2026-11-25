import { useEffect, useMemo, useState } from "react";
import { X, Grip, Maximize2 } from "lucide-react";
import { Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, ComposedChart, Funnel, FunnelChart, LabelList, Line, LineChart, Pie, PieChart, PolarAngleAxis, PolarGrid, PolarRadiusAxis, Radar, RadarChart, ResponsiveContainer, Tooltip, XAxis, YAxis, Scatter, ScatterChart, Treemap } from "recharts";
import type { FloatingChartData, SheetData } from "../types";
import { addressOf, parseAddress } from "../types";
import { numericValue } from "../utils/formulaEngine";
interface Props { chart:FloatingChartData; sheet:SheetData; onUpdate:(id:string,changes:Partial<FloatingChartData>)=>void; onDelete:(id:string)=>void; }
const COLORS=["#0F6CBD","#62ABF5","#107C41","#038387","#8A70C4","#F4B544"];
function chartRows(chart:FloatingChartData,sheet:SheetData){
  const [start,end]=chart.range.split(":");const a=parseAddress(start??"A1"),b=parseAddress(end??start??"A1");
  if(!a||!b)return [];
  const rows:Record<string,string|number>[]=[];
  for(let r=a.row;r<=Math.min(b.row,sheet.rowCount);r++){
    const label=sheet.cells[addressOf(r,a.col)]?.raw||`Row ${r+1}`;
    const item:Record<string,string|number>={name:label,x:r-a.row+1};
    for(let c=a.col+1;c<=b.col;c++){
      const raw=sheet.cells[addressOf(r,c)]?.raw??"0";
      item[`series${c-a.col}`]=numericValue(raw,sheet.cells,sheet.tables??[],addressOf(r,c))??0;
    }
    if(b.col===a.col)item.value=numericValue(sheet.cells[addressOf(r,a.col)]?.raw??"0",sheet.cells,sheet.tables??[],addressOf(r,a.col))??0;
    else item.value=item.series1??0;
    item.size=Math.max(0,Number(item.value)||0);
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
  const label=chart.title||({column:"Column chart",line:"Line chart",scatter:"Scatter chart",pie:"Pie chart",bar:"Bar chart",statistical:"Statistical chart",combo:"Combo chart",treemap:"Treemap"}[chart.type]??"Chart");
  return <section className="absolute z-20 overflow-hidden rounded-xl border border-[#D2D0CE] bg-white shadow-[0_12px_36px_rgba(0,0,0,.17)]" style={{left:chart.x,top:chart.y,width:chart.width,height:chart.height}}>
    <header onMouseDown={e=>{setMoving(true);e.preventDefault()}} className="flex h-9 cursor-move items-center gap-2 border-b border-[#E5E5E5] bg-[#F8F8F8] px-3"><Grip size={14} className="text-[#8A8886]"/><span className="min-w-0 flex-1 truncate text-xs font-semibold text-[#323130]">{label}</span><button onMouseDown={e=>e.preventDefault()} onClick={()=>onDelete(chart.id)} className="rounded-md p-1.5 text-[#797775] hover:bg-[#FDE7E9] hover:text-[#C4314B]" aria-label="Delete chart"><X size={14}/></button></header>
    <div className="h-[calc(100%-36px)] w-full p-3 pb-5">
      {data.length===0?<div className="grid h-full place-items-center text-xs text-[#8A8886]">Select a data range to populate this chart.</div>:<ResponsiveContainer width="100%" height="100%">
        {chart.type==="pie"?<PieChart><Tooltip/><Pie data={data} dataKey="value" nameKey="name" outerRadius="72%" innerRadius={chart.variant?.toLowerCase().includes("doughnut")?"42%":0} paddingAngle={chart.variant?.toLowerCase().includes("doughnut")?2:0} label>{data.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie></PieChart>
        :chart.type==="combo"?<ComposedChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="#E5E5E5"/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/><Bar dataKey="value" fill={COLORS[0]} radius={[3,3,0,0]}/><Line dataKey="value" stroke="#323130" strokeWidth={2}/></ComposedChart>
        :chart.type==="line"&&chart.variant?.toLowerCase().includes("area")?<AreaChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="#E5E5E5"/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/>{seriesKeys.length?seriesKeys.map((k,i)=><Area key={k} type="monotone" dataKey={k} stroke={COLORS[i%COLORS.length]} fill={COLORS[i%COLORS.length]} fillOpacity={0.25} stackId={chart.variant?.toLowerCase().includes("stacked")?"area":undefined}/>):<Area type="monotone" dataKey="value" stroke={COLORS[0]} fill={COLORS[1]} fillOpacity={0.3}/>}</AreaChart>
        :chart.type==="line"?<LineChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="#E5E5E5"/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/>{seriesKeys.length?seriesKeys.map((k,i)=><Line key={k} dataKey={k} stroke={COLORS[i%COLORS.length]} strokeWidth={2} dot/>):<Line dataKey="value" stroke={COLORS[0]} strokeWidth={2} dot/>}</LineChart>
        :chart.type==="scatter"?<ScatterChart><CartesianGrid strokeDasharray="3 3" stroke="#E5E5E5"/><XAxis type="number" dataKey="x" name="Row" tick={{fontSize:10}}/><YAxis type="number" dataKey="value" name="Value" tick={{fontSize:10}}/><Tooltip cursor={{strokeDasharray:"3 3"}}/><Scatter data={data} fill={COLORS[0]}/></ScatterChart>
        :chart.type==="treemap"&&chart.variant?.toLowerCase().includes("funnel")?<FunnelChart><Tooltip/><Funnel data={data} dataKey="size" nameKey="name" fill={COLORS[0]}><LabelList dataKey="name" position="right"/></Funnel></FunnelChart>
        :chart.type==="treemap"&&chart.variant?.toLowerCase().includes("radar")?<RadarChart data={data}><PolarGrid stroke="#E5E5E5"/><PolarAngleAxis dataKey="name" tick={{fontSize:9}}/><PolarRadiusAxis tick={{fontSize:8}}/><Tooltip/><Radar dataKey="value" stroke={COLORS[0]} fill={COLORS[1]} fillOpacity={0.4}/></RadarChart>
        :chart.type==="treemap"&&chart.variant?.toLowerCase().includes("sunburst")?<PieChart><Tooltip/><Pie data={data} dataKey="size" nameKey="name" innerRadius="24%" outerRadius="78%" paddingAngle={1}>{data.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Pie data={data} dataKey="value" nameKey="name" innerRadius="0%" outerRadius="22%" stroke="#fff">{data.map((_,i)=><Cell key={i} fill={COLORS[(i+2)%COLORS.length]}/>)}</Pie></PieChart>
        :chart.type==="treemap"&&chart.variant?.toLowerCase().includes("waterfall")?<BarChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="#E5E5E5"/><XAxis dataKey="name" tick={{fontSize:9}}/><YAxis tick={{fontSize:9}}/><Tooltip/><Bar dataKey="value" fill={COLORS[0]} radius={[2,2,0,0]}>{data.map((item,i)=><Cell key={i} fill={Number(item.value)<0?"#D13438":COLORS[i%COLORS.length]}/>)}</Bar></BarChart>
        :chart.type==="treemap"?<Treemap data={data} dataKey="size" nameKey="name" stroke="#fff" fill={COLORS[0]}><Tooltip/></Treemap>
        :<BarChart data={data} layout={chart.type==="bar"?"vertical":"horizontal"}><CartesianGrid strokeDasharray="3 3" stroke="#E5E5E5"/>{chart.type==="bar"?<><XAxis type="number" tick={{fontSize:10}}/><YAxis type="category" dataKey="name" width={72} tick={{fontSize:10}}/></>:<><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/></>}<Tooltip/>{seriesKeys.length?seriesKeys.map((k,i)=><Bar key={k} dataKey={k} fill={COLORS[i%COLORS.length]} stackId={chart.variant?.toLowerCase().includes("stacked")?"stack":undefined}/>):<Bar dataKey="value" fill={COLORS[0]} radius={[3,3,0,0]}/>}</BarChart>}
      </ResponsiveContainer>}
    </div>
    <button aria-label="Resize chart" onMouseDown={e=>{e.preventDefault();setResizing(true)}} className="absolute bottom-0 right-0 grid h-6 w-6 cursor-se-resize place-items-center text-[#797775]"><Maximize2 size={12}/></button>
  </section>;
}
