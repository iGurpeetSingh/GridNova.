import { useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { ChevronDown, X, Table2 } from "lucide-react";
import type { PivotAggregation, PivotTableDefinition, SheetData, CellRange } from "../types";
import { addressOf } from "../types";
interface Props { open:boolean; source:SheetData; sourceRange:CellRange; activeSheetId:string; onClose:()=>void; onCreate:(definition:PivotTableDefinition)=>void; }
type Zone="Rows"|"Columns"|"Values"|"Filters";
export default function PivotTableBuilderDialog({open,source,sourceRange,activeSheetId,onClose,onCreate}:Props){
  const headers=useMemo(()=>Array.from({length:sourceRange.endCol-sourceRange.startCol+1},(_,index)=>source.cells[addressOf(sourceRange.startRow,sourceRange.startCol+index)]?.raw||`Column ${index+1}`),[source,sourceRange]);
  const [assigned,setAssigned]=useState<Record<Zone,number[]>>({Rows:[],Columns:[],Values:[],Filters:[]});
  const [aggregation,setAggregation]=useState<PivotAggregation>("Sum"),[showRows,setShowRows]=useState(true),[showColumns,setShowColumns]=useState(true),[destination,setDestination]=useState<"new"|"existing">("new"),[destinationCell,setDestinationCell]=useState("A1"),[name,setName]=useState("PivotTable");
  const uniqueValues=useMemo(()=>headers.map((_,field)=>Array.from(new Set(Array.from({length:sourceRange.endRow-sourceRange.startRow},(_,offset)=>source.cells[addressOf(sourceRange.startRow+offset+1,sourceRange.startCol+field)]?.raw??"").filter(Boolean)))),[headers,source,sourceRange]);
  const [filterSelections,setFilterSelections]=useState<Record<number,string[]>>({});
  if(!open)return null;
  const setZone=(field:number,zone:Zone)=>setAssigned(current=>{
    const next={...current,Rows:current.Rows.filter(value=>value!==field),Columns:current.Columns.filter(value=>value!==field),Values:current.Values.filter(value=>value!==field),Filters:current.Filters.filter(value=>value!==field)};
    next[zone]=zone==="Values"?[field]:[...next[zone],field];
    if(zone==="Filters")setFilterSelections(current=>({...current,[field]:uniqueValues[field]??[]}));
    return next;
  });
  const selected=(field:number)=>Object.entries(assigned).find(([,fields])=>fields.includes(field))?.[0] as Zone|undefined;
  const create=()=>{
    if(assigned.Rows.length===0&&assigned.Columns.length===0&&assigned.Values.length===0)return;
    const valueField=assigned.Values[0]??0;
    const definition:PivotTableDefinition={id:crypto.randomUUID(),name,sourceSheetId:source.id,sourceRange,rowFields:assigned.Rows,columnFields:assigned.Columns,valueField,aggregation,filterFields:assigned.Filters,filters:Object.fromEntries(assigned.Filters.map(field=>[field,filterSelections[field]??uniqueValues[field]??[]])),showRowGrandTotals:showRows,showColumnGrandTotals:showColumns,destinationSheetId:destination==="new"?"__new__":activeSheetId,destinationCell:destinationCell.toUpperCase()};
    onCreate(definition);
  };
  return createPortal(<div className="advanced-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)onClose()}}>
    <section className="advanced-dialog pivot-builder-dialog" role="dialog" aria-modal="true" aria-labelledby="pivot-builder-title" onMouseDown={event=>event.stopPropagation()}>
      <header><div className="advanced-dialog-icon"><Table2 size={19}/></div><div><h2 id="pivot-builder-title">Create PivotTable</h2><p>Choose fields to summarize the selected data.</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="advanced-dialog-body pivot-builder-body">
        <div className="pivot-field-list"><strong>Available fields</strong>{headers.map((label,index)=><div key={index}><span>{label}</span><div>{(["Rows","Columns","Values","Filters"] as const).map(zone=><button type="button" key={zone} className={selected(index)===zone?"is-selected":""} onClick={()=>setZone(index,zone)}>{zone.slice(0,1)}</button>)}</div></div>)}<small>Tap R / C / V / F to assign each source column.</small></div>
        <div className="pivot-zones">{(["Rows","Columns","Values","Filters"] as const).map(zone=><div key={zone}><strong>{zone}</strong><div>{assigned[zone].length?assigned[zone].map(field=><span key={field}>{headers[field]}</span>):<small>Drop or assign fields</small>}</div></div>)}</div>
        {assigned.Filters.map(field=><label className="advanced-field pivot-filter-select" key={field}><span>Filter: {headers[field]}</span><select multiple value={filterSelections[field]??uniqueValues[field]??[]} onChange={event=>setFilterSelections(current=>({...current,[field]:Array.from(event.target.selectedOptions,option=>option.value)}))}>{(uniqueValues[field]??[]).map(value=><option key={value} value={value}>{value}</option>)}</select><small>Use Ctrl / Cmd to select multiple values.</small></label>)}
        <label className="advanced-field"><span>Aggregation</span><select value={aggregation} onChange={event=>setAggregation(event.target.value as PivotAggregation)}>{(["Sum","Count","Average","Max","Min"] as const).map(value=><option key={value}>{value}</option>)}</select></label>
        <div className="pivot-totals"><label className="advanced-check"><input type="checkbox" checked={showRows} onChange={event=>setShowRows(event.target.checked)}/> Show row grand totals</label><label className="advanced-check"><input type="checkbox" checked={showColumns} onChange={event=>setShowColumns(event.target.checked)}/> Show column grand totals</label></div>
        <div className="advanced-subheading">Destination</div><div className="pivot-destination"><label className="advanced-check"><input type="radio" name="pivot-destination" checked={destination==="new"} onChange={()=>setDestination("new")}/> New Sheet</label><label className="advanced-check"><input type="radio" name="pivot-destination" checked={destination==="existing"} onChange={()=>setDestination("existing")}/> Existing Sheet</label></div>
        {destination==="existing"&&<label className="advanced-field"><span>Destination cell</span><input value={destinationCell} onChange={event=>setDestinationCell(event.target.value.toUpperCase())}/></label>}
        <label className="advanced-field"><span>PivotTable name</span><input value={name} onChange={event=>setName(event.target.value)}/></label>
      </div>
      <footer><button type="button" className="advanced-secondary" onClick={onClose}>Cancel</button><button type="button" disabled={!assigned.Values.length} onClick={create}>Create PivotTable</button></footer>
    </section>
  </div>,document.body);
}
