import type { WorkbookData, SheetData } from "../types";
import { parseAddress } from "../types";
import { evaluate } from "./formulaEngine";
import { extractFormulaReferences } from "../lib/dependency-graph";
export interface GoalSeekInput { sheet:SheetData; workbook:WorkbookData; setCell:string; targetValue:number; changingCell:string; tolerance?:number; maxIterations?:number; }
export interface GoalSeekResult { success:boolean; foundInput:number; targetValue:number; currentValue:number; iterations:number; message:string; }
function evaluateAt(input:GoalSeekInput,address:string,value:number):number{
  const p=parseAddress(address);if(!p)throw new Error("The Set cell address is invalid.");
  const cells={...input.sheet.cells,[input.changingCell.toUpperCase()]:{...(input.sheet.cells[input.changingCell.toUpperCase()]??{raw:""}),raw:String(value),cachedValue:undefined}};
  const formula=input.sheet.cells[address.toUpperCase()]?.raw??"";
  const result=evaluate(formula,cells,new Set(),input.sheet.tables??[],address,input.workbook.namedRanges??{},input.workbook.sheets,input.sheet.name);
  const numeric=typeof result==="number"?result:Number(result);
  if(!Number.isFinite(numeric))throw new Error(`The Set cell evaluated to ${String(result)} instead of a number.`);
  return numeric;
}
export function solveGoalSeek(input:GoalSeekInput):GoalSeekResult{
  const target=input.targetValue,tolerance=input.tolerance??1e-6,maxIterations=input.maxIterations??200;
  const setAddress=input.setCell.toUpperCase(),changeAddress=input.changingCell.toUpperCase();
  const formula=input.sheet.cells[setAddress]?.raw??"";
  if(!formula.startsWith("="))throw new Error("Set cell must contain an Excel formula.");
  const originalRaw=input.sheet.cells[changeAddress]?.raw??"";
  const original=Number(originalRaw);
  if(originalRaw.trim()===""||!Number.isFinite(original))throw new Error("By changing cell must contain a numeric value.");
  const formulaReferences=extractFormulaReferences(formula,input.workbook.namedRanges??{});
  const changing=parseAddress(changeAddress);
  const referenced=formulaReferences.some(ref=>(!ref.sheetName||ref.sheetName.toLowerCase()===input.sheet.name.toLowerCase())&&!!changing&&changing.row>=ref.start.row&&changing.row<=ref.end.row&&changing.col>=ref.start.col&&changing.col<=ref.end.col);
  if(!referenced)throw new Error("The Set cell formula does not reference the By changing cell.");
  const objective=(value:number)=>evaluateAt(input,setAddress,value)-target;
  let currentValue=evaluateAt(input,setAddress,original),fCurrent=currentValue-target;
  if(Math.abs(fCurrent)<=tolerance)return {success:true,foundInput:original,targetValue:target,currentValue,iterations:0,message:"Target already reached."};
  const seed=Math.max(1,Math.abs(original)*0.1),samples:{x:number;f:number}[]=[{x:original,f:fCurrent}];
  let step=seed;
  for(let i=0;i<40;i++){
    for(const x of [original-step,original+step]){
      if(!Number.isFinite(x)||samples.some(item=>item.x===x))continue;
      const f=objective(x);samples.push({x,f});
      if(Math.abs(f)<=tolerance)return {success:true,foundInput:x,targetValue:target,currentValue:f+target,iterations:i+1,message:"Target reached."};
    }
    const sorted=[...samples].sort((a,b)=>a.x-b.x);
    const bracket=sorted.find((point,index)=>index>0&&((sorted[index-1].f<=0&&point.f>=0)||(sorted[index-1].f>=0&&point.f<=0)));
    if(bracket){
      const index=sorted.indexOf(bracket),left=sorted[index-1],right=bracket;
      let lo=left.x,hi=right.x,fLo=left.f,fHi=right.f,best= Math.abs(fLo)<Math.abs(fHi)?lo:hi;
      for(let iteration=1;iteration<=maxIterations;iteration++){
        const denominator=fHi-fLo;
        let candidate=denominator!==0?hi-fHi*(hi-lo)/denominator:(lo+hi)/2;
        const low=Math.min(lo,hi),high=Math.max(lo,hi);
        if(!Number.isFinite(candidate)||candidate<=low+0.05*(high-low)||candidate>=high-0.05*(high-low))candidate=(lo+hi)/2;
        const f=objective(candidate);
        if(Math.abs(f)<Math.abs(objective(best)))best=candidate;
        if(Math.abs(f)<=tolerance){currentValue=f+target;return {success:true,foundInput:candidate,targetValue:target,currentValue,iterations:i+iteration+1,message:"Target reached."};}
        if((fLo<=0&&f>=0)||(fLo>=0&&f<=0)){hi=candidate;fHi=f;}else{lo=candidate;fLo=f;}
        if(Math.abs(hi-lo)<=tolerance*Math.max(1,Math.abs(candidate)))break;
      }
      const finalValue=evaluateAt(input,setAddress,best);
      return {success:Math.abs(finalValue-target)<=Math.max(tolerance,1e-4),foundInput:best,targetValue:target,currentValue:finalValue,iterations:maxIterations,message:Math.abs(finalValue-target)<=Math.max(tolerance,1e-4)?"Best solution found.":"Unable to reach the target within tolerance."};
    }
    step*=2;
  }
  // Fallback secant iteration when the function is monotonic but a coarse bracket was not found.
  let x0=original,x1=original+(original===0?1:seed),f0=fCurrent,f1=objective(x1);
  for(let iteration=1;iteration<=maxIterations;iteration++){
    const denominator=f1-f0;if(Math.abs(denominator)<1e-15)break;
    const x2=x1-f1*(x1-x0)/denominator;if(!Number.isFinite(x2))break;
    const f2=objective(x2);
    if(Math.abs(f2)<=tolerance)return {success:true,foundInput:x2,targetValue:target,currentValue:f2+target,iterations:iteration,message:"Target reached."};
    x0=x1;f0=f1;x1=x2;f1=f2;
  }
  currentValue=evaluateAt(input,setAddress,x1);
  return {success:false,foundInput:x1,targetValue:target,currentValue,iterations:maxIterations,message:"No numeric solution was found. Check the formula and changing cell."};
}
