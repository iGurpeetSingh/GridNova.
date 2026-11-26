import { useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { Check, Palette } from "lucide-react";
import type { CellData, SheetData, SparklineData } from "../types";
import { parseAddress, addressOf } from "../types";
import { useWorkbookStore } from "../store/useWorkbookStore";
interface Props { sheet:SheetData; sparkline:SparklineData; }
const colors=[{name:"Blue",value:"#0F6CBD"},{name:"Green",value:"#107C41"},{name:"Red",value:"#C4314B"},{name:"Purple",value:"#8764B8"},{name:"Orange",value:"#E36C09"}];
function sourceValues(sparkline:SparklineData,cells:Record<string,CellData>):number[]{
  const [first,lastRaw]=sparkline.dataRange.toUpperCase().split(":"),last=lastRaw??first;
  const a=parseAddress(first),b=parseAddress(last);if(!a||!b)return [];
  const values:number[]=[];
  for(let row=Math.min(a.row,b.row);row<=Math.max(a.row,b.row);row++)for(let col=Math.min(a.col,b.col);col<=Math.max(a.col,b.col);col++){
    const raw=cells[addressOf(row,col)]?.raw??"",number=Number(raw.startsWith("=")?cells[addressOf(row,col)]?.cachedValue??raw:raw);
    if(Number.isFinite(number))values.push(number);
  }
  return values;
}
export default function SparklineCell({sheet,sparkline}:Props){
  const [open,setOpen]=useState(false),[position,setPosition]=useState({x:8,y:8});
  const updateSheet=useWorkbookStore(state=>state.updateSheet);
  const values=useMemo(()=>sourceValues(sparkline,sheet.cells),[sparkline.dataRange,sheet.cells]);
  const width=104,height=24,pad=3,min=Math.min(0,...values),max=Math.max(0,...values),span=max-min||1;
  const points=values.map((value,index)=>({x:pad+(values.length<=1?width/2:index/(values.length-1)*(width-pad*2)),y:height-pad-(value-min)/span*(height-pad*2),value}));
  const path=points.map((point,index)=>`${index?"L":"M"}${point.x.toFixed(1)},${point.y.toFixed(1)}`).join(" ");
  const baseline=height-pad-(0-min)/span*(height-pad*2);
  const choose=(color:string)=>{updateSheet(sheet.id,{sparklines:{...(sheet.sparklines??{}),[sparkline.location]:{...sparkline,color}}});setOpen(false);};
  return <div className="sparkline-cell-wrap">
    <button type="button" className="sparkline-cell-button" aria-label={`Sparkline ${sparkline.type} for ${sparkline.location}`} onMouseDown={event=>event.preventDefault()} onClick={event=>{const rect=(event.currentTarget as HTMLElement).getBoundingClientRect();setPosition({x:Math.min(window.innerWidth-190,rect.left),y:Math.min(window.innerHeight-190,rect.bottom+4)});setOpen(v=>!v);}}>
      <svg viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none" aria-hidden="true">
        {sparkline.type==="line"?<><path d={path} fill="none" stroke={sparkline.color} strokeWidth="1.7" strokeLinejoin="round" strokeLinecap="round"/>{sparkline.showHighPoint&&points.length>0&&<circle cx={points.reduce((a,b)=>a.value>b.value?a:b).x} cy={points.reduce((a,b)=>a.value>b.value?a:b).y} r="2.5" fill="#107C41"/>}{sparkline.showLowPoint&&points.length>0&&<circle cx={points.reduce((a,b)=>a.value<b.value?a:b).x} cy={points.reduce((a,b)=>a.value<b.value?a:b).y} r="2.5" fill="#C4314B"/>}</>:sparkline.type==="column"?points.map((point,index)=><rect key={index} x={point.x-3} y={Math.min(point.y,baseline)} width="5.5" height={Math.max(1,Math.abs(baseline-point.y))} rx=".7" fill={sparkline.color}/>):<>{points.map((point,index)=><rect key={index} x={point.x-3} y={point.value>=0?baseline-5:baseline} width="5.5" height="5" rx=".5" fill={point.value>=0?sparkline.color:sparkline.negativeColor}/>)}</>}
      </svg><span><Palette size={10}/></span>
    </button>
    {open&&createPortal(<div className="sparkline-style-popover" style={{left:position.x,top:position.y}}><strong>Quick Style</strong>{colors.map(color=><button type="button" key={color.value} onMouseDown={event=>event.preventDefault()} onClick={()=>choose(color.value)}><i style={{background:color.value}}/>{color.name}{sparkline.color===color.value&&<Check size={12}/>}</button>)}<button type="button" className="sparkline-style-close" onClick={()=>setOpen(false)}>Close</button></div>,document.body)}
  </div>;
}
