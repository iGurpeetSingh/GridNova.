import type { CellData, CellMap, CellRange, PivotTableDefinition, SheetData, WorkbookData } from "../types";
import { addressOf, parseAddress } from "../types";
export interface PivotBuildResult { cells:CellMap; range:CellRange; rowCount:number; columnCount:number; }
interface RecordRow { values:string[]; rowKey:string; columnKey:string; value:number; validValue:boolean; }
function aggregate(values:number[],kind:PivotTableDefinition["aggregation"]):number{
  if(kind==="Count")return values.length;
  if(!values.length)return 0;
  if(kind==="Sum")return values.reduce((sum,value)=>sum+value,0);
  if(kind==="Average")return values.reduce((sum,value)=>sum+value,0)/values.length;
  if(kind==="Max")return Math.max(...values);
  return Math.min(...values);
}
export function buildPivotTable(definition:PivotTableDefinition,source:SheetData):PivotBuildResult{
  const range=definition.sourceRange,headers=Array.from({length:range.endCol-range.startCol+1},(_,index)=>source.cells[addressOf(range.startRow,range.startCol+index)]?.raw||`Column ${index+1}`);
  const fieldLabel=(field:number)=>headers[field]??`Column ${field+1}`;
  const records:RecordRow[]=[];
  const passesFilters=(values:string[])=>definition.filterFields.every(field=>{
    const selected=definition.filters?.[field];return !selected?.length||selected.includes(values[field]??"");
  });
  for(let row=range.startRow+1;row<=range.endRow;row++){
    const values=headers.map((_,index)=>source.cells[addressOf(row,range.startCol+index)]?.raw??"");
    if(!passesFilters(values))continue;
    const rowValues=definition.rowFields.map(field=>values[field]??"");
    const columnValues=definition.columnFields.map(field=>values[field]??"");
    const raw=values[definition.valueField]??"",numeric=Number(raw);
    records.push({values,rowKey:JSON.stringify(rowValues),columnKey:JSON.stringify(columnValues),value:numeric,validValue:raw.trim()!==""&&Number.isFinite(numeric)});
  }
  const rowGroups=new Map<string,string[]>(),columnGroups=new Map<string,string[]>();
  records.forEach(record=>{
    rowGroups.set(record.rowKey,definition.rowFields.map(field=>record.values[field]??""));
    columnGroups.set(record.columnKey,definition.columnFields.map(field=>record.values[field]??""));
  });
  if(!rowGroups.size)rowGroups.set("[]",[]);
  if(!columnGroups.size)columnGroups.set("[]",[]);
  const rowKeys=[...rowGroups.keys()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  const columnKeys=[...columnGroups.keys()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  const matrix:string[][]=[];
  const rowHeader=definition.rowFields.length?definition.rowFields.map(field=>fieldLabel(field)).join(" / "):"Values";
  const header=[rowHeader,...columnKeys.map(key=>columnGroups.get(key)?.join(" / ")||"Total")];
  if(definition.showRowGrandTotals)header.push("Grand Total");
  matrix.push(header);
  const rowTotals:number[]=[],columnTotals=columnKeys.map(()=>[] as number[]),allValues:number[]=[];
  rowKeys.forEach(rowKey=>{
    const rowValues=rowGroups.get(rowKey)??[],rowOutput=[rowValues.join(" / ")||"Total"],thisRowValues:number[]=[];
    columnKeys.forEach((columnKey,index)=>{
      const matched=records.filter(record=>record.rowKey===rowKey&&record.columnKey===columnKey);
      const vals=matched.filter(record=>definition.aggregation==="Count"?record.values[definition.valueField]!=="" : record.validValue).map(record=>definition.aggregation==="Count"?1:record.value);
      const value=aggregate(vals,definition.aggregation);
      rowOutput.push(vals.length?String(value):"");
      thisRowValues.push(...vals);columnTotals[index].push(...vals);allValues.push(...vals);
    });
    const rowTotal=aggregate(thisRowValues,definition.aggregation);rowTotals.push(...thisRowValues);
    if(definition.showRowGrandTotals)rowOutput.push(thisRowValues.length?String(rowTotal):"");
    matrix.push(rowOutput);
  });
  if(definition.showColumnGrandTotals){
    const totalRow=["Grand Total",...columnTotals.map(values=>values.length?String(aggregate(values,definition.aggregation)):"")];
    if(definition.showRowGrandTotals)totalRow.push(allValues.length?String(aggregate(allValues,definition.aggregation)):"");
    matrix.push(totalRow);
  }
  const start=parseAddress(definition.destinationCell)||{row:0,col:0},cells:CellMap={};
  matrix.forEach((row,rowIndex)=>row.forEach((raw,colIndex)=>{
    const address=addressOf(start.row+rowIndex,start.col+colIndex);
    const headerCell=rowIndex===0,totalCell=rowIndex===matrix.length-1&&definition.showColumnGrandTotals;
    const style=headerCell?{bold:true,color:"#ffffff",backgroundColor:"#107C41"}:totalCell?{bold:true,backgroundColor:"#E8F5EE"}:colIndex===0?{bold:true,backgroundColor:"#F3F6F4"}:undefined;
    cells[address]={raw,style};
  }));
  return {cells,range:{startRow:start.row,startCol:start.col,endRow:start.row+matrix.length-1,endCol:start.col+Math.max(0,...matrix.map(row=>row.length))-1},rowCount:matrix.length,columnCount:Math.max(0,...matrix.map(row=>row.length))};
}
export function refreshWorkbookPivots(workbook:WorkbookData):WorkbookData{
  const definitions=workbook.sheets.flatMap(destination=>(destination.pivotTables??[]).map(definition=>({destinationSheetId:destination.id,definition})));
  if(!definitions.length)return workbook;
  let sheets=workbook.sheets.map(sheet=>({...sheet,cells:{...sheet.cells}}));
  for(const item of definitions){
    const source=sheets.find(sheet=>sheet.id===item.definition.sourceSheetId),destination=sheets.find(sheet=>sheet.id===item.destinationSheetId);
    if(!source||!destination)continue;
    const definition=item.definition;
    const built=buildPivotTable(definition,source),updatedCells={...destination.cells};
    const previous=definition.outputRange??built.range;
    for(let row=previous.startRow;row<=previous.endRow;row++)for(let col=previous.startCol;col<=previous.endCol;col++)delete updatedCells[addressOf(row,col)];
    Object.assign(updatedCells,built.cells);
    const nextDefinition={...definition,outputRange:built.range};
    sheets=sheets.map(sheet=>sheet.id===destination.id?{...sheet,cells:updatedCells,pivotTables:(sheet.pivotTables??[]).map(pivot=>pivot.id===definition.id?nextDefinition:pivot),rowCount:Math.max(sheet.rowCount,built.range.endRow+1),columnCount:Math.max(sheet.columnCount,built.range.endCol+1)}:sheet);
  }
  return {...workbook,sheets,modifiedAt:Date.now()};
}
