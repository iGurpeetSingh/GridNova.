import type { CellMap, CellRange, ConditionalFormattingRule, SheetData } from "../types";
import { parseAddress, addressOf } from "../types";
import { numericValue } from "./formulaEngine";
function contains(range:CellRange,address:string){const point=parseAddress(address);return !!point&&point.row>=range.startRow&&point.row<=range.endRow&&point.col>=range.startCol&&point.col<=range.endCol;}
function colorMix(a:string,b:string,ratio:number){
  const parse=(value:string)=>{const v=value.replace("#","");return [0,2,4].map(offset=>parseInt(v.slice(offset,offset+2),16)||0);};
  const x=parse(a),y=parse(b),v=x.map((value,index)=>Math.round(value+(y[index]-value)*Math.max(0,Math.min(1,ratio))));
  return "#"+v.map(value=>value.toString(16).padStart(2,"0")).join("");
}
function percentile(values:number[],fraction:number):number{const sorted=[...values].sort((a,b)=>a-b);if(!sorted.length)return 0;const position=(sorted.length-1)*fraction,lower=Math.floor(position),upper=Math.ceil(position);return sorted[lower]+(sorted[upper]-sorted[lower])*(position-lower);}
function valuesFor(rule:ConditionalFormattingRule,sheet:SheetData):number[]{
  const values:number[]=[];
  for(let row=rule.range.startRow;row<=rule.range.endRow;row++)for(let col=rule.range.startCol;col<=rule.range.endCol;col++){
    const address=addressOf(row,col),value=numericValue(sheet.cells[address]?.raw??"",sheet.cells,sheet.tables??[],address);
    if(value!==null)values.push(value);
  }
  return values;
}
export interface ConditionalCellStyle { background?:string; color?:string; dataBar?:{color:string;percent:number;gradient:boolean}; icon?:{text:string;color:string}; }
export function conditionalStyleForCell(sheet:SheetData,address:string):ConditionalCellStyle{
  const result:ConditionalCellStyle={};
  const rules=[...(sheet.conditionalFormattingRules??[])].filter(rule=>rule.enabled&&contains(rule.range,address)).sort((a,b)=>a.priority-b.priority);
  const cell=sheet.cells[address],value=numericValue(cell?.raw??"",sheet.cells,sheet.tables??[],address),raw=cell?.raw??"";
  for(const rule of rules){
    const numbers=valuesFor(rule,sheet),min=numbers.length?Math.min(...numbers):0,max=numbers.length?Math.max(...numbers):0,avg=numbers.length?numbers.reduce((sum,n)=>sum+n,0)/numbers.length:0;
    if(rule.type==="dataBar"&&value!==null){
      const percent=max===min?100:Math.max(0,Math.min(100,(value-min)/(max-min)*100));
      result.dataBar={color:rule.color??"#63A5DB",percent,gradient:rule.gradient!==false};
    }else if(rule.type==="colorScale"&&value!==null){
      const ratio=max===min?0.5:(value-min)/(max-min);
      if(rule.midColor){result.background=ratio<=.5?colorMix(rule.minColor??"#F8696B",rule.midColor,ratio*2):colorMix(rule.midColor,rule.maxColor??"#63BE7B",(ratio-.5)*2);}
      else result.background=colorMix(rule.minColor??"#F8696B",rule.maxColor??"#63BE7B",ratio);
    }else if(rule.type==="iconSet"&&value!==null){
      const lower=percentile(numbers,1/3),upper=percentile(numbers,2/3),tier=value>=upper?2:value>=lower?1:0,style=rule.iconStyle??"arrows";
      result.icon=style==="trafficLights"?{text:"●",color:tier===2?"#107C41":tier===1?"#E3A008":"#C4314B"}:style==="checks"?{text:tier===2?"✓":tier===1?"−":"×",color:tier===2?"#107C41":tier===1?"#C28A00":"#C4314B"}:{text:tier===2?"↑":tier===1?"→":"↓",color:tier===2?"#107C41":tier===1?"#C28A00":"#C4314B"};
    }else if(rule.type==="greaterThan"&&value!==null&&value>Number(rule.formula1)){result.background=rule.backgroundColor??"#E2F0D9";result.color=rule.fontColor??"#256029";}
    else if(rule.type==="lessThan"&&value!==null&&value<Number(rule.formula1)){result.background=rule.backgroundColor??"#FCE4D6";result.color=rule.fontColor??"#9C0006";}
    else if(rule.type==="between"&&value!==null&&value>=Number(rule.formula1)&&value<=Number(rule.formula2)){result.background=rule.backgroundColor??"#FFF2CC";result.color=rule.fontColor??"#7F6000";}
    else if(rule.type==="equalTo"&&raw.toLowerCase()===String(rule.formula1??"").toLowerCase()){result.background=rule.backgroundColor??"#DDEBF7";result.color=rule.fontColor??"#1F4E79";}
    else if(rule.type==="textContains"&&raw.toLowerCase().includes((rule.text??rule.formula1??"").toLowerCase())){result.background=rule.backgroundColor??"#E4DFEC";result.color=rule.fontColor??"#553C77";}
    else if(rule.type==="duplicateValues"||rule.type==="uniqueValues"){
      const count=Object.values(sheet.cells).filter(item=>item.raw.toLowerCase()===raw.toLowerCase()&&raw!=="").length;
      if((rule.type==="duplicateValues"&&count>1)||(rule.type==="uniqueValues"&&count===1)){result.background=rule.backgroundColor??"#F4CCCC";result.color=rule.fontColor??"#990000";}
    }else if(rule.type==="topBottom"&&value!==null){
      const n=rule.count??10,sorted=[...numbers].sort((a,b)=>rule.top===false?a-b:b-a),threshold=sorted[Math.min(sorted.length-1,Math.max(0,n-1))];
      if(rule.top===false?value<=threshold:value>=threshold){result.background=rule.backgroundColor??"#E2F0D9";result.color=rule.fontColor??"#256029";}
    }else if(rule.type==="aboveAverage"&&value!==null&&value>avg){result.background=rule.backgroundColor??"#E2F0D9";result.color=rule.fontColor??"#256029";}
    else if(rule.type==="belowAverage"&&value!==null&&value<avg){result.background=rule.backgroundColor??"#FCE4D6";result.color=rule.fontColor??"#9C0006";}
  }
  return result;
}

export function buildConditionalFormattingMap(sheet:SheetData):Record<string,ConditionalCellStyle>{
  const result:Record<string,ConditionalCellStyle>={};
  const merge=(address:string,next:ConditionalCellStyle)=>{result[address]={...(result[address]??{}),...next};};
  const rules=[...(sheet.conditionalFormattingRules??[])].filter(rule=>rule.enabled).sort((a,b)=>a.priority-b.priority);
  for(const rule of rules){
    const rows:number[]=[];
    const addresses:string[]=[];
    const rawValues:string[]=[];
    for(let row=rule.range.startRow;row<=rule.range.endRow;row++)for(let col=rule.range.startCol;col<=rule.range.endCol;col++){
      const address=addressOf(row,col),raw=sheet.cells[address]?.raw??"";
      addresses.push(address);rawValues.push(raw);
      const numericValueRaw=numericValue(raw,sheet.cells,sheet.tables??[],address);
      if(numericValueRaw!==null)rows.push(numericValueRaw);
    }
    const min=rows.length?Math.min(...rows):0,max=rows.length?Math.max(...rows):0,avg=rows.length?rows.reduce((sum,value)=>sum+value,0)/rows.length:0;
    const rawCounts=new Map<string,number>();rawValues.forEach(value=>{if(value!=="")rawCounts.set(value.toLowerCase(),(rawCounts.get(value.toLowerCase())??0)+1);});
    addresses.forEach((address,index)=>{
      const raw=rawValues[index],value=numericValue(raw,sheet.cells,sheet.tables??[],address);
      const style:ConditionalCellStyle={};
      if(rule.type==="dataBar"&&value!==null)style.dataBar={color:rule.color??"#63A5DB",percent:max===min?100:Math.max(0,Math.min(100,(value-min)/(max-min)*100)),gradient:rule.gradient!==false};
      else if(rule.type==="colorScale"&&value!==null){const ratio=max===min?0.5:(value-min)/(max-min);style.background=rule.midColor?(ratio<=.5?colorMix(rule.minColor??"#F8696B",rule.midColor,ratio*2):colorMix(rule.midColor,rule.maxColor??"#63BE7B",(ratio-.5)*2)):colorMix(rule.minColor??"#F8696B",rule.maxColor??"#63BE7B",ratio);}
      else if(rule.type==="iconSet"&&value!==null){const lower=percentile(rows,1/3),upper=percentile(rows,2/3),tier=value>=upper?2:value>=lower?1:0,kind=rule.iconStyle??"arrows";style.icon=kind==="trafficLights"?{text:"●",color:tier===2?"#107C41":tier===1?"#E3A008":"#C4314B"}:kind==="checks"?{text:tier===2?"✓":tier===1?"−":"×",color:tier===2?"#107C41":tier===1?"#C28A00":"#C4314B"}:{text:tier===2?"↑":tier===1?"→":"↓",color:tier===2?"#107C41":tier===1?"#C28A00":"#C4314B"};}
      else if(rule.type==="greaterThan"&&value!==null&&value>Number(rule.formula1)){style.background=rule.backgroundColor??"#E2F0D9";style.color=rule.fontColor??"#256029";}
      else if(rule.type==="lessThan"&&value!==null&&value<Number(rule.formula1)){style.background=rule.backgroundColor??"#FCE4D6";style.color=rule.fontColor??"#9C0006";}
      else if(rule.type==="between"&&value!==null&&value>=Number(rule.formula1)&&value<=Number(rule.formula2)){style.background=rule.backgroundColor??"#FFF2CC";style.color=rule.fontColor??"#7F6000";}
      else if(rule.type==="equalTo"&&raw.toLowerCase()===String(rule.formula1??"").toLowerCase()){style.background=rule.backgroundColor??"#DDEBF7";style.color=rule.fontColor??"#1F4E79";}
      else if(rule.type==="textContains"&&raw.toLowerCase().includes((rule.text??rule.formula1??"").toLowerCase())){style.background=rule.backgroundColor??"#E4DFEC";style.color=rule.fontColor??"#553C77";}
      else if(rule.type==="duplicateValues"&&raw!==""&&(rawCounts.get(raw.toLowerCase())??0)>1){style.background=rule.backgroundColor??"#F4CCCC";style.color=rule.fontColor??"#990000";}
      else if(rule.type==="uniqueValues"&&raw!==""&&(rawCounts.get(raw.toLowerCase())??0)===1){style.background=rule.backgroundColor??"#E2F0D9";style.color=rule.fontColor??"#256029";}
      else if(rule.type==="aboveAverage"&&value!==null&&value>avg){style.background=rule.backgroundColor??"#E2F0D9";style.color=rule.fontColor??"#256029";}
      else if(rule.type==="belowAverage"&&value!==null&&value<avg){style.background=rule.backgroundColor??"#FCE4D6";style.color=rule.fontColor??"#9C0006";}
      else if(rule.type==="topBottom"&&value!==null){const top=rule.top!==false,sorted=[...rows].sort((a,b)=>top?b-a:a-b),threshold=sorted[Math.min(sorted.length-1,Math.max(0,(rule.count??10)-1))];if(top?value>=threshold:value<=threshold){style.background=rule.backgroundColor??"#E2F0D9";style.color=rule.fontColor??"#256029";}}
      if(style.background||style.color||style.dataBar||style.icon)merge(address,style);
    });
  }
  return result;
}
