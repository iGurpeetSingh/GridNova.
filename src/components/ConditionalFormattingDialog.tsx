import { useEffect, useState } from "react";
import { createPortal } from "react-dom";
import { BarChart3, Palette, TrafficCone, X, Trash2 } from "lucide-react";
import type { CellRange, ConditionalFormattingRule, ConditionalRuleType } from "../types";
interface Props { open:boolean; range:CellRange; rules:ConditionalFormattingRule[]; onClose:()=>void; onSave:(rules:ConditionalFormattingRule[])=>void; }
const types:{value:ConditionalRuleType;label:string}[]=[
 {value:"dataBar",label:"Data Bars"},{value:"colorScale",label:"Color Scales"},{value:"iconSet",label:"Icon Sets"},
 {value:"greaterThan",label:"Greater Than"},{value:"lessThan",label:"Less Than"},{value:"between",label:"Between"},{value:"equalTo",label:"Equal To"},{value:"textContains",label:"Text Contains"},{value:"duplicateValues",label:"Duplicate Values"},{value:"uniqueValues",label:"Unique Values"},{value:"topBottom",label:"Top / Bottom N"},{value:"aboveAverage",label:"Above Average"},{value:"belowAverage",label:"Below Average"}
];
export default function ConditionalFormattingDialog({open,range,rules,onClose,onSave}:Props){
  const [nextRules,setNextRules]=useState<ConditionalFormattingRule[]>([]),[type,setType]=useState<ConditionalRuleType>("dataBar"),[formula1,setFormula1]=useState("0"),[formula2,setFormula2]=useState("100"),[text,setText]=useState(""),[color,setColor]=useState("#63A5DB"),[minColor,setMinColor]=useState("#F8696B"),[midColor,setMidColor]=useState("#FFEB84"),[maxColor,setMaxColor]=useState("#63BE7B"),[iconStyle,setIconStyle]=useState<"arrows"|"trafficLights"|"checks">("arrows"),[top,setTop]=useState(true),[count,setCount]=useState("10"),[gradient,setGradient]=useState(true),[editingId,setEditingId]=useState<string|null>(null);
  useEffect(()=>{if(open)setNextRules([...rules].sort((a,b)=>a.priority-b.priority));},[open,rules]);
  if(!open)return null;
  const edit=(rule:ConditionalFormattingRule)=>{
    setEditingId(rule.id);setType(rule.type);setFormula1(rule.formula1??"0");setFormula2(rule.formula2??"100");setText(rule.text??"");
    setColor(rule.color??"#63A5DB");setMinColor(rule.minColor??"#F8696B");setMidColor(rule.midColor??"#FFEB84");setMaxColor(rule.maxColor??"#63BE7B");
    setIconStyle(rule.iconStyle??"arrows");setTop(rule.top!==false);setCount(String(rule.count??10));setGradient(rule.gradient!==false);
  };
  const create=()=>{
    const existing=editingId?nextRules.find(item=>item.id===editingId):undefined;
    const rule:ConditionalFormattingRule={id:editingId??crypto.randomUUID(),range:existing?.range??range,type,enabled:existing?.enabled??true,priority:existing?.priority??Math.max(0,...nextRules.map(item=>item.priority))+1,formula1,formula2,text,color,minColor,midColor,maxColor,iconStyle,top,count:Number(count)||10,gradient};
    setNextRules(items=>editingId?items.map(item=>item.id===editingId?rule:item):[...items,rule]);setEditingId(null);
  };
  const update=(id:string,changes:Partial<ConditionalFormattingRule>)=>setNextRules(items=>items.map(item=>item.id===id?{...item,...changes}:item));
  const remove=(id:string)=>setNextRules(items=>items.filter(item=>item.id!==id).map((item,index)=>({...item,priority:index+1})));
  return createPortal(<div className="advanced-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="advanced-dialog" role="dialog" aria-modal="true" aria-labelledby="cf-title" onMouseDown={event=>event.stopPropagation()}>
      <header><div className="advanced-dialog-icon"><Palette size={19}/></div><div><h2 id="cf-title">Conditional Formatting Rules</h2><p>Live rules for selected range</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="advanced-dialog-body">
        <div className="advanced-subheading">New Rule</div><label className="advanced-field"><span>Rule type</span><select value={type} onChange={event=>setType(event.target.value as ConditionalRuleType)}>{types.map(item=><option value={item.value} key={item.value}>{item.label}</option>)}</select></label>
        {["greaterThan","lessThan","between","equalTo"].includes(type)&&<><label className="advanced-field"><span>Value</span><input value={formula1} onChange={event=>setFormula1(event.target.value)}/></label>{type==="between"&&<label className="advanced-field"><span>And</span><input value={formula2} onChange={event=>setFormula2(event.target.value)}/></label>}</>}
        {type==="textContains"&&<label className="advanced-field"><span>Text contains</span><input value={text} onChange={event=>setText(event.target.value)}/></label>}
        {type==="dataBar"&&<><div className="advanced-data-bar-themes"><span>Bar color themes</span>{[["Blue","#63A5DB"],["Green","#70AD47"],["Orange","#ED7D31"],["Teal","#038387"],["Purple","#8A70C4"],["Red","#C0504D"]].map(([label,value])=><button key={value} type="button" title={label} aria-label={label} className={color===value?"is-active":""} style={{background:value}} onClick={()=>setColor(value)}/>)}</div><label className="advanced-check"><input type="checkbox" checked={gradient} onChange={event=>setGradient(event.target.checked)}/> Gradient fill</label><label className="advanced-field"><span>Custom bar color</span><input type="color" value={color} onChange={event=>setColor(event.target.value)}/></label></>}
        {type==="colorScale"&&<div className="advanced-color-pair"><label><span>Minimum</span><input type="color" value={minColor} onChange={event=>setMinColor(event.target.value)}/></label><label><span>Midpoint</span><input type="color" value={midColor} onChange={event=>setMidColor(event.target.value)}/></label><label><span>Maximum</span><input type="color" value={maxColor} onChange={event=>setMaxColor(event.target.value)}/></label></div>}
        {type==="iconSet"&&<label className="advanced-field"><span>Icon style</span><select value={iconStyle} onChange={event=>setIconStyle(event.target.value as typeof iconStyle)}><option value="arrows">Directional arrows</option><option value="trafficLights">Traffic lights</option><option value="checks">Checks / crosses</option></select></label>}
        {type==="topBottom"&&<div className="advanced-color-pair"><label><span>Top or Bottom</span><select value={top?"top":"bottom"} onChange={event=>setTop(event.target.value==="top")}><option value="top">Top</option><option value="bottom">Bottom</option></select></label><label><span>Count</span><input value={count} onChange={event=>setCount(event.target.value)}/></label></div>}
        <div className="advanced-edit-actions"><button className="advanced-inline-add" type="button" onClick={create}>{editingId?"Update Rule":"Add Rule"}</button>{editingId&&<button className="advanced-inline-cancel" type="button" onClick={()=>setEditingId(null)}>Cancel Edit</button>}</div>
        <div className="advanced-subheading">Manage Rules ({nextRules.length})</div>
        {nextRules.length?nextRules.map(rule=><div className="advanced-rule-row" key={rule.id}><input type="checkbox" checked={rule.enabled} onChange={event=>update(rule.id,{enabled:event.target.checked})}/><div><strong>{types.find(item=>item.value===rule.type)?.label??rule.type}</strong><small>Priority {rule.priority}</small></div><button type="button" title="Move up" onClick={()=>{const i=nextRules.findIndex(item=>item.id===rule.id);if(i<=0)return;const copy=[...nextRules];[copy[i-1],copy[i]]=[copy[i],copy[i-1]];setNextRules(copy.map((item,index)=>({...item,priority:index+1})));}}>↑</button><button type="button" title="Move down" onClick={()=>{const i=nextRules.findIndex(item=>item.id===rule.id);if(i<0||i>=nextRules.length-1)return;const copy=[...nextRules];[copy[i+1],copy[i]]=[copy[i],copy[i+1]];setNextRules(copy.map((item,index)=>({...item,priority:index+1})));}} >↓</button><button type="button" title="Edit rule" onClick={()=>edit(rule)}>Edit</button><button type="button" aria-label="Delete rule" onClick={()=>remove(rule.id)}><Trash2 size={14}/></button></div>):<p className="advanced-muted">No rules defined for this sheet range.</p>}
      </div>
      <footer><button type="button" className="advanced-secondary" onClick={onClose}>Cancel</button><button type="button" onClick={()=>{onSave(nextRules);onClose();}}>Apply Rules</button></footer>
    </section>
  </div>,document.body);
}
