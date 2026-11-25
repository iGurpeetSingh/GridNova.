import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";
import { useNavigate, useParams } from "react-router-dom";
import { ArrowLeft, Check, ChevronDown, Cloud, FileSpreadsheet } from "lucide-react";
import Ribbon from "../components/Ribbon";
import SpreadsheetGrid from "../components/SpreadsheetGrid";
import FormulaBar from "../components/FormulaBar";
import SheetBar from "../components/SheetBar";
const FloatingChart = lazy(() => import("../components/FloatingChart"));
import ShareModal from "../components/ShareModal";
import TableRangeDialog from "../components/TableRangeDialog";
import FormulaLibraryDialog from "../components/FormulaLibraryDialog";
import CellInsertDialog from "../components/CellInsertDialog";
import FloatingElement from "../components/FloatingElement";
import CellImageOverlays from "../components/CellImageOverlays";
import { useWorkbookStore } from "../store/useWorkbookStore";
import type { CellData, CellRange, CellStyle, ExcelTableData, FloatingChartData, FloatingElementData } from "../types";
import { addressOf, parseAddress } from "../types";
import { numericValue } from "../utils/formulaEngine";
import { getTableAt, getTableHeaderLabel, getTablePalette, normalizeTableName, tableColumnValues } from "../utils/tableUtils";


type Tab = "File"|"Home"|"Insert"|"Share"|"Page Layout"|"Formulas"|"Data"|"Review"|"View"|"Help"|"Draw"|"Table Design";
const parseRangeLabel = (value: string): CellRange | null => {
  const match = value.trim().toUpperCase().match(/^\$?([A-Z]+)\$?([1-9]\d*)\s*:\s*\$?([A-Z]+)\$?([1-9]\d*)$/);
  if (!match) return null;
  const start = parseAddress(`${match[1]}${match[2]}`), end = parseAddress(`${match[3]}${match[4]}`);
  if (!start || !end) return null;
  return { startRow: Math.min(start.row, end.row), startCol: Math.min(start.col, end.col), endRow: Math.max(start.row, end.row), endCol: Math.max(start.col, end.col) };
};
const rangeLabel = (range: CellRange) => `${addressOf(range.startRow, range.startCol)}:${addressOf(range.endRow, range.endCol)}`;
const shiftFormulaRowRefs = (raw: string, startRow: number) => raw.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (match, colLock: string, letters: string, rowLock: string, digits: string) => {
  const row = Number(digits) - 1;
  return row >= startRow ? `${colLock}${letters}${rowLock}${row + 2}` : match;
});
export default function WorkbookEditor(){
  const {id=""}=useParams();const navigate=useNavigate();
  const workbooks=useWorkbookStore(s=>s.workbooks);
  const workbook=workbooks.find(w=>w.id===id);
  const activeWorkbookId=useWorkbookStore(s=>s.activeWorkbookId);
  const activeSheetId=useWorkbookStore(s=>s.activeSheetId);
  const activeCell=useWorkbookStore(s=>s.activeCell);
  const selection=useWorkbookStore(s=>s.selection);
  const showFormulas=useWorkbookStore(s=>s.showFormulas);
  const zoom=useWorkbookStore(s=>s.zoom);
  const clipboard=useWorkbookStore(s=>s.clipboard);
  const setActiveWorkbook=useWorkbookStore(s=>s.setActiveWorkbook);
  const setActiveSheet=useWorkbookStore(s=>s.setActiveSheet);
  const setSelection=useWorkbookStore(s=>s.setSelection);
  const setCell=useWorkbookStore(s=>s.setCell);
  const setStyle=useWorkbookStore(s=>s.setStyle);
  const setClipboard=useWorkbookStore(s=>s.setClipboard);
  const paste=useWorkbookStore(s=>s.paste);
  const addSheet=useWorkbookStore(s=>s.addSheet);
  const renameSheet=useWorkbookStore(s=>s.renameSheet);
  const deleteSheet=useWorkbookStore(s=>s.deleteSheet);
  const updateSheet=useWorkbookStore(s=>s.updateSheet);
  const addChart=useWorkbookStore(s=>s.addChart);
  const updateChart=useWorkbookStore(s=>s.updateChart);
  const deleteChart=useWorkbookStore(s=>s.deleteChart);
  const addElement=useWorkbookStore(s=>s.addElement);
  const updateElement=useWorkbookStore(s=>s.updateElement);
  const deleteElement=useWorkbookStore(s=>s.deleteElement);
  const undo=useWorkbookStore(s=>s.undo);const redo=useWorkbookStore(s=>s.redo);
  const setShowFormulas=useWorkbookStore(s=>s.setShowFormulas);const setZoom=useWorkbookStore(s=>s.setZoom);
  const setOrientation=useWorkbookStore(s=>s.setOrientation);
  const [tab,setTab]=useState<Tab>("Home");const [formulaValue,setFormulaValue]=useState("");
  const [formulaBarFocused,setFormulaBarFocused]=useState(false);
  const [shareOpen,setShareOpen]=useState(false);const [renaming,setRenaming]=useState(false);const [titleValue,setTitleValue]=useState("");
  const [tableDialogOpen,setTableDialogOpen]=useState(false);
  const [tableDialogMode,setTableDialogMode]=useState<"create"|"resize">("create");
  const [formulaLibraryOpen,setFormulaLibraryOpen]=useState(false);
  const [cellDialogKind,setCellDialogKind]=useState<"link"|"comment"|null>(null);
  const [commentsPanelOpen,setCommentsPanelOpen]=useState(false);
  const [slicerOpen,setSlicerOpen]=useState(false);
  const [slicerColumn,setSlicerColumn]=useState("0");
  const [slicerValues,setSlicerValues]=useState<string[]>([]);
  const [picturePlacement,setPicturePlacement]=useState<"cell"|"over">("over");
  const [savePulse,setSavePulse]=useState(false);
  const fileRef=useRef<HTMLInputElement>(null);
  const pictureRef=useRef<HTMLInputElement>(null);
  const renameWorkbook=useWorkbookStore(s=>s.renameWorkbook);
  const setShowSaved=()=>{setSavePulse(true);window.setTimeout(()=>setSavePulse(false),1400)};
  useEffect(()=>{if(id&&activeWorkbookId!==id)setActiveWorkbook(id)},[id,activeWorkbookId,setActiveWorkbook]);
  useEffect(()=>{setTitleValue(workbook?.title??"")},[workbook?.title]);
  const sheet=workbook?.sheets.find(s=>s.id===activeSheetId)??workbook?.sheets[0];
  const activeRaw=sheet?.cells[activeCell]?.raw??"";
  useEffect(()=>{setFormulaValue(activeRaw)},[activeRaw,activeCell,activeSheetId]);
  const activeStyle=sheet?.cells[activeCell]?.style??{};
  const selectedAddresses=useMemo(()=>{
    if(!selection)return [activeCell];
    const out:string[]=[];
    for(let r=Math.min(selection.startRow,selection.endRow);r<=Math.max(selection.startRow,selection.endRow);r++)
      for(let c=Math.min(selection.startCol,selection.endCol);c<=Math.max(selection.startCol,selection.endCol);c++)out.push(addressOf(r,c));
    return out;
  },[selection,activeCell]);
  const stats=useMemo(()=>{
    if(!sheet)return {count:0,sum:0,average:0};
    const vals=selectedAddresses.map(a=>numericValue(sheet.cells[a]?.raw??"",sheet.cells,sheet.tables??[],a)).filter((v):v is number=>v!==null);
    return {count:vals.length,sum:vals.reduce((a,b)=>a+b,0),average:vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:0};
  },[sheet,selectedAddresses]);
  const activePosition = parseAddress(activeCell);
  const activeTable = activePosition && sheet ? getTableAt(sheet.tables, activePosition.row, activePosition.col) ?? null : null;
  const suggestedTableRange = useMemo(() => {
    if (!sheet) return "A1:D10";
    if (selection) return rangeLabel({ startRow: Math.min(selection.startRow, selection.endRow), startCol: Math.min(selection.startCol, selection.endCol), endRow: Math.max(selection.startRow, selection.endRow), endCol: Math.max(selection.startCol, selection.endCol) });
    const used = Object.entries(sheet.cells).map(([address, cell]) => ({ position: parseAddress(address), cell })).filter(item => item.position && (item.cell.raw !== "" || item.cell.checkbox || item.cell.datePicker));
    if (!used.length) return "A1:D10";
    const positions = used.map(item => item.position!).filter(Boolean);
    return rangeLabel({ startRow: Math.min(...positions.map(item => item.row)), startCol: Math.min(...positions.map(item => item.col)), endRow: Math.max(...positions.map(item => item.row)), endCol: Math.max(...positions.map(item => item.col)) });
  }, [sheet, selection]);
  useEffect(() => { if (tab === "Table Design" && !activeTable) setTab("Home"); }, [tab, activeTable]);
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const range = parseRangeLabel(params.get("range") ?? "");
    if (range && params.get("sheet") === sheet?.id) setSelection(range, addressOf(range.startRow, range.startCol));
  }, [workbook?.id, sheet?.id, setSelection]);
  if(!workbook||!sheet)return <main className="app-shell grid min-h-screen place-items-center p-6"><div className="glass max-w-md rounded-3xl p-8 text-center"><FileSpreadsheet className="mx-auto mb-3 text-[#149052]" size={36}/><h1 className="text-lg font-bold text-[#284331]">Workbook not found</h1><p className="mt-2 text-sm text-[#86948b]">This workbook may have been deleted or is not saved in this browser.</p><button onClick={()=>navigate("/")} className="btn btn-primary mt-5"><ArrowLeft size={14}/> Back to dashboard</button></div></main>;

  const selectCell=(address:string,range?:CellRange)=>setSelection(range??null,address);
  const commitTitle=()=>{renameWorkbook(workbook.id,titleValue);setRenaming(false);setShowSaved()};
  const exportFile=async(type:"xlsx"|"csv"|"pdf")=>{
    try {
      const tools=await import("../utils/exportUtils");
      if(type==="xlsx")await tools.exportXlsx(workbook);else if(type==="csv")await tools.exportCsv(sheet);else await tools.exportPdf(workbook);
      setShareOpen(false);setShowSaved();
    } catch(error) { window.alert(error instanceof Error?error.message:"Export failed."); }
  };
  const selectedRangeString=()=>{
    if(!selection)return activeCell;
    return `${addressOf(selection.startRow,selection.startCol)}:${addressOf(selection.endRow,selection.endCol)}`;
  };
  const updateCells=(cells:typeof sheet.cells)=>updateSheet(sheet.id,{cells});
  const applyCellFormat=(style:Partial<CellStyle>)=>{
    if(style.border==="outer"){
      const startRow=selection?Math.min(selection.startRow,selection.endRow):parseAddress(activeCell)?.row??0;
      const endRow=selection?Math.max(selection.startRow,selection.endRow):startRow;
      const startCol=selection?Math.min(selection.startCol,selection.endCol):parseAddress(activeCell)?.col??0;
      const endCol=selection?Math.max(selection.startCol,selection.endCol):startCol;
      const cells={...sheet.cells},border=`1px solid ${style.borderColor??"#64748b"}`;
      selectedAddresses.forEach(address=>{
        const pos=parseAddress(address);if(!pos)return;
        const cell=cells[address]??{raw:""},nextStyle={...(cell.style??{})};delete nextStyle.border;
        if(pos.row===startRow)nextStyle.borderTop=border;
        if(pos.row===endRow)nextStyle.borderBottom=border;
        if(pos.col===startCol)nextStyle.borderLeft=border;
        if(pos.col===endCol)nextStyle.borderRight=border;
        cells[address]={...cell,style:nextStyle};
      });
      updateCells(cells);return;
    }
    if(style.border==="all")setStyle({...style,borderColor:style.borderColor??"#64748b"},selectedAddresses);
    else if(style.border==="none"){
      const cells={...sheet.cells};
      selectedAddresses.forEach(address=>{const cell=cells[address]??{raw:""},nextStyle={...(cell.style??{})};delete nextStyle.borderTop;delete nextStyle.borderRight;delete nextStyle.borderBottom;delete nextStyle.borderLeft;nextStyle.border="none";cells[address]={...cell,style:nextStyle};});
      updateCells(cells);
    }else setStyle(style,selectedAddresses);
  };
  const copySelection=(remove=false)=>{
    const matrix:string[][]=[];const styles:(CellStyle|undefined)[][]=[];
    if(!selection){const cell=sheet.cells[activeCell];setClipboard({cells:[[{raw:cell?.raw??"",style:cell?.style}]],source:activeCell});if(remove)setCell(activeCell,"");return;}
    for(let r=Math.min(selection.startRow,selection.endRow);r<=Math.max(selection.startRow,selection.endRow);r++){
      const row:string[]=[],styleRow:(CellStyle|undefined)[]=[];
      for(let c=Math.min(selection.startCol,selection.endCol);c<=Math.max(selection.startCol,selection.endCol);c++){const cell=sheet.cells[addressOf(r,c)];row.push(cell?.raw??"");styleRow.push(cell?.style);}
      matrix.push(row);styles.push(styleRow);
    }
    setClipboard({cells:matrix.map((row,r)=>row.map((raw,c)=>({raw,style:styles[r][c]}))),source:addressOf(Math.min(selection.startRow,selection.endRow),Math.min(selection.startCol,selection.endCol))});
    if(remove){const cells={...sheet.cells};selectedAddresses.forEach(a=>{cells[a]={raw:""}});updateCells(cells);}
  };
  const shiftRowsCols=(axis:"row"|"column",index:number,mode:"insert"|"delete")=>{
    const cells:typeof sheet.cells={};
    Object.entries(sheet.cells).forEach(([address,cell])=>{
      const p=parseAddress(address);if(!p)return;
      const pos=axis==="row"?p.row:p.col;
      if(mode==="delete"&&pos===index)return;
      const next=mode==="insert"&&pos>=index?pos+1:mode==="delete"&&pos>index?pos-1:pos;
      cells[axis==="row"?addressOf(next,p.col):addressOf(p.row,next)]=cell;
    });
    const tables=(sheet.tables??[]).flatMap(table=>{
      let start=axis==="row"?table.range.startRow:table.range.startCol,end=axis==="row"?table.range.endRow:table.range.endCol;
      if(mode==="insert"){
        if(index<=start){start++;end++;}
        else if(index<=end+1)end++;
      }else{
        if(index<start){start--;end--;}
        else if(index<=end){if(index===start)return[];end--;if(end<start)return[];}
      }
      return [{...table,range:axis==="row"?{...table.range,startRow:start,endRow:end}:{...table.range,startCol:start,endCol:end}}];
    });
    updateSheet(sheet.id,{cells,tables,rowCount:axis==="row"?Math.max(1,sheet.rowCount+(mode==="insert"?1:-1)):sheet.rowCount,columnCount:axis==="column"?Math.max(1,sheet.columnCount+(mode==="insert"?1:-1)):sheet.columnCount});
  };
  const openTableDialog=(mode:"create"|"resize")=>{setTableDialogMode(mode);setTableDialogOpen(true)};
  const updateActiveTable=(changes:Partial<ExcelTableData>)=>{
    if(!activeTable)return;
    updateSheet(sheet.id,{tables:(sheet.tables??[]).map(table=>table.id===activeTable.id?{...table,...changes}:table)});
  };
  const submitTableRange=(rangeText:string,hasHeaders:boolean)=>{
    const parsed=parseRangeLabel(rangeText);
    if(!parsed){window.alert("Enter a valid range such as A1:D10.");return;}
    const resizing=tableDialogMode==="resize"&&activeTable;
    const overlaps=(sheet.tables??[]).some(table=>{
      if(resizing&&table.id===activeTable.id)return false;
      const end=table.range.endRow+(table.showTotalRow?1:0);
      return parsed.startRow<=end&&parsed.endRow>=table.range.startRow&&parsed.startCol<=table.range.endCol&&parsed.endCol>=table.range.startCol;
    });
    if(overlaps){window.alert("This range overlaps another Excel Table. Choose a separate range.");return;}
    let cells={...sheet.cells},range={...parsed};
    if(!resizing&&!hasHeaders){
      const shifted:typeof sheet.cells={};
      Object.entries(sheet.cells).forEach(([key,cell])=>{
        const position=parseAddress(key);if(!position)return;
        const row=position.row>=parsed.startRow?position.row+1:position.row;
        shifted[addressOf(row,position.col)]={...cell,raw:cell.raw.startsWith("=")?shiftFormulaRowRefs(cell.raw,parsed.startRow):cell.raw};
      });
      cells=shifted;range.endRow+=1;
    }
    const used=new Set<string>();
    for(let col=range.startCol;col<=range.endCol;col++){
      const key=addressOf(range.startRow,col);
      const initial=resizing?cells[key]?.raw?.trim():hasHeaders?cells[key]?.raw?.trim():"";
      const base=initial||("Column"+(col-range.startCol+1));
      let label=base,suffix=2;
      while(used.has(label.toLowerCase()))label=base+"_"+(suffix++);
      used.add(label.toLowerCase());
      cells[key]={...(cells[key]??{raw:""}),raw:label};
    }
    const tables=[...(sheet.tables??[])];
    if(resizing&&activeTable){
      updateSheet(sheet.id,{cells,tables:tables.map(table=>table.id===activeTable.id?{...table,range,hasHeaders}:table),rowCount:Math.max(sheet.rowCount,range.endRow+1),columnCount:Math.max(sheet.columnCount,range.endCol+1)});
    }else{
      const tableName=normalizeTableName("Table"+(tables.length+1));
      const table:ExcelTableData={id:crypto.randomUUID(),name:tableName,range,hasHeaders:true,showTotalRow:false,firstColumn:false,lastColumn:false,bandedRows:true,bandedColumns:false,showFilterButtons:true,styleId:"TableStyleMedium16",filters:{}};
      updateSheet(sheet.id,{cells,tables:[...tables,table],rowCount:Math.max(sheet.rowCount+(hasHeaders?0:1),range.endRow+1),columnCount:Math.max(sheet.columnCount,range.endCol+1)});
      selectCell(addressOf(range.startRow,range.startCol),range);
      setTab("Table Design");
    }
    setTableDialogOpen(false);setShowSaved();
  };
  const setTableTotalRow=(enable:boolean)=>{
    if(!activeTable||activeTable.showTotalRow===enable)return;
    const totalRow=activeTable.range.endRow+1;
    if(enable){
      shiftRowsCols("row",totalRow,"insert");
      const latest=useWorkbookStore.getState().workbooks.find(item=>item.id===workbook.id)?.sheets.find(item=>item.id===sheet.id);
      if(!latest)return;
      const cells={...latest.cells};
      for(let col=activeTable.range.startCol;col<=activeTable.range.endCol;col++){
        const header=latest.cells[addressOf(activeTable.range.startRow,col)]?.raw||("Column"+(col-activeTable.range.startCol+1));
        const values=Array.from({length:activeTable.range.endRow-activeTable.range.startRow},(_,i)=>latest.cells[addressOf(activeTable.range.startRow+i+1,col)]?.raw??"");
        const numericColumn=values.some(value=>value.trim()!==""&&Number.isFinite(Number(value)));
        cells[addressOf(totalRow,col)]={raw:"=SUBTOTAL("+(numericColumn?109:103)+","+activeTable.name+"["+header+"])",style:{bold:true}};
      }
      updateSheet(sheet.id,{cells,tables:(latest.tables??[]).map(table=>table.id===activeTable.id?{...table,range:activeTable.range,showTotalRow:true}:table)});
    }else{
      shiftRowsCols("row",totalRow,"delete");
      const latest=useWorkbookStore.getState().workbooks.find(item=>item.id===workbook.id)?.sheets.find(item=>item.id===sheet.id);
      if(latest)updateSheet(sheet.id,{tables:(latest.tables??[]).map(table=>table.id===activeTable.id?{...table,showTotalRow:false}:table)});
    }
  };
  const convertActiveTableToRange=()=>{
    if(!activeTable)return;
    const palette=getTablePalette(activeTable.styleId),cells={...sheet.cells};
    for(let row=activeTable.range.startRow;row<=activeTable.range.endRow+(activeTable.showTotalRow?1:0);row++)
      for(let col=activeTable.range.startCol;col<=activeTable.range.endCol;col++){
        const key=addressOf(row,col),cell=cells[key]??{raw:""},style={...(cell.style??{})};
        if(!style.backgroundColor)style.backgroundColor=row===activeTable.range.startRow?palette.header:activeTable.bandedRows&&(row-activeTable.range.startRow)%2===0?palette.band:"#ffffff";
        if(row===activeTable.range.startRow&&!style.color)style.color=palette.text;
        if(row===activeTable.range.startRow)style.bold=true;
        cells[key]={...cell,style};
      }
    updateSheet(sheet.id,{cells,tables:(sheet.tables??[]).filter(table=>table.id!==activeTable.id)});
    setTab("Home");
  };
  const removeActiveTableDuplicates=()=>{
    if(!activeTable)return;
    const seen=new Set<string>(),keep:Array<Array<typeof sheet.cells[string]|undefined>>=[];
    for(let row=activeTable.range.startRow+1;row<=activeTable.range.endRow;row++){
      const record=Array.from({length:activeTable.range.endCol-activeTable.range.startCol+1},(_,offset)=>sheet.cells[addressOf(row,activeTable.range.startCol+offset)]);
      const signature=record.map(cell=>cell?.raw??"").join("\u001f");
      if(!seen.has(signature)){seen.add(signature);keep.push(record);}
    }
    const cells={...sheet.cells},firstData=activeTable.range.startRow+1;
    for(let row=firstData;row<=activeTable.range.endRow+(activeTable.showTotalRow?1:0);row++)
      for(let col=activeTable.range.startCol;col<=activeTable.range.endCol;col++)delete cells[addressOf(row,col)];
    keep.forEach((record,offset)=>record.forEach((cell,col)=>{if(cell)cells[addressOf(firstData+offset,activeTable.range.startCol+col)]=cell;}));
    const endRow=activeTable.range.startRow+keep.length;
    if(activeTable.showTotalRow)for(let col=activeTable.range.startCol;col<=activeTable.range.endCol;col++){
      const old=sheet.cells[addressOf(activeTable.range.endRow+1,col)];
      if(old)cells[addressOf(endRow+1,col)]=old;
    }
    updateSheet(sheet.id,{cells,tables:(sheet.tables??[]).map(table=>table.id===activeTable.id?{...table,range:{...table.range,endRow}}:table)});
  };
  const insertFloatingElement=(element:Partial<FloatingElementData>)=>{
    addElement({id:crypto.randomUUID(),type:"shape",x:145,y:90,width:element.type==="text"?280:element.type==="image"?320:170,height:element.type==="text"?120:element.type==="image"?210:110,...element});
    setShowSaved();
  };
  const openSlicer=()=>{
    if(!activeTable)return;
    const first=activeTable.range.startCol;
    setSlicerColumn(String(first));
    setSlicerValues(tableColumnValues(activeTable,first,sheet.cells));
    setSlicerOpen(true);
  };
  const applySlicer=()=>{
    if(!activeTable)return;
    const column=Number(slicerColumn),allValues=tableColumnValues(activeTable,column,sheet.cells);
    const filters={...(activeTable.filters??{})};
    if(slicerValues.length===allValues.length)delete filters[column];
    else filters[column]={...(filters[column]??{}),selectedValues:slicerValues};
    updateSheet(sheet.id,{tables:(sheet.tables??[]).map(table=>table.id===activeTable.id?{...table,filters}:table)});
    setSlicerOpen(false);setShowSaved();
  };
  const importPicture=(file:File)=>{
    if(!file.type.startsWith("image/")){window.alert("Choose an image file.");return;}
    if(file.size>6*1024*1024){window.alert("Choose an image smaller than 6 MB.");return;}
    const reader=new FileReader();
    reader.onload=()=>{
      if(typeof reader.result!=="string")return;
      if(picturePlacement==="cell"){
        const cells={...sheet.cells},existing=cells[activeCell]??{raw:""};
        cells[activeCell]={...existing,image:reader.result};
        updateSheet(sheet.id,{cells});setShowSaved();
      }else insertFloatingElement({type:"image",src:reader.result,text:file.name,width:320,height:210,fill:"#ffffff",stroke:"#0F6CBD"});
    };
    reader.onerror=()=>window.alert("Unable to read this picture.");
    reader.readAsDataURL(file);
  };
  const createPivotTable=()=>{
    if(!selection||selectedAddresses.length<4){window.alert("Select a source range with a category column and at least one data row to insert a PivotTable.");return;}
    const source={startRow:Math.min(selection.startRow,selection.endRow),endRow:Math.max(selection.startRow,selection.endRow),startCol:Math.min(selection.startCol,selection.endCol),endCol:Math.max(selection.startCol,selection.endCol)};
    if(source.endCol===source.startCol){window.alert("A PivotTable needs a category column and at least one value column.");return;}
    const groups=new Map<string,{count:number;total:number}>();
    for(let row=source.startRow+1;row<=source.endRow;row++){
      const key=sheet.cells[addressOf(row,source.startCol)]?.raw?.trim()||"(blank)";
      const raw=sheet.cells[addressOf(row,source.startCol+1)]?.raw??"";
      const numeric=numericValue(raw,sheet.cells,sheet.tables??[],addressOf(row,source.startCol+1))??0;
      const existing=groups.get(key)??{count:0,total:0};
      existing.count+=1;existing.total+=numeric;groups.set(key,existing);
    }
    const startRow=source.startRow,startCol=source.endCol+2,endRow=startRow+groups.size+1;
    const cells={...sheet.cells};
    [["Row Labels","Count","Sum"],...Array.from(groups.entries()).map(([name,value])=>[name,String(value.count),String(value.total)])].forEach((row,index)=>row.forEach((raw,offset)=>{cells[addressOf(startRow+index,startCol+offset)]={raw,style:index===0?{bold:true,color:"#ffffff",backgroundColor:"#0F6CBD"}:undefined};}));
    const name=normalizeTableName("PivotTable"+((sheet.tables?.length??0)+1));
    const table:ExcelTableData={id:crypto.randomUUID(),name,range:{startRow,startCol,endRow,endCol:startCol+2},hasHeaders:true,showTotalRow:false,firstColumn:false,lastColumn:false,bandedRows:true,bandedColumns:false,showFilterButtons:true,styleId:"TableStyleMedium9",filters:{}};
    updateSheet(sheet.id,{cells,tables:[...(sheet.tables??[]),table],rowCount:Math.max(sheet.rowCount,endRow+1),columnCount:Math.max(sheet.columnCount,startCol+3)});
    selectCell(addressOf(startRow,startCol),table.range);setTab("Table Design");setShowSaved();
  };
  const copyShareLink=async(kind:"workbook"|"range"|"table")=>{
    let section="";
    if(kind!=="workbook"){
      const sharedRange=kind==="table"&&activeTable?rangeLabel(activeTable.range):selectedRangeString();
      section="?sheet="+encodeURIComponent(sheet.id)+"&range="+encodeURIComponent(sharedRange);
      if(kind==="table"&&activeTable)section+="&table="+encodeURIComponent(activeTable.name);
    }
    const url=window.location.origin+"/workbook/"+workbook.id+section;
    try{await navigator.clipboard.writeText(url);window.alert("Link copied. Actual shared access and permissions require an authenticated sharing backend.");}
    catch{window.prompt("Copy this link",url);}
  };
  const sortData=(descending:boolean)=>{
    const p=parseAddress(activeCell);if(!p)return;
    const header=sheet.cells[addressOf(0,p.col)]?.raw;
    const rows=Array.from({length:Math.max(0,sheet.rowCount-1)},(_,i)=>i+1);
    rows.sort((a,b)=>{
      const av=sheet.cells[addressOf(a,p.col)]?.raw??"",bv=sheet.cells[addressOf(b,p.col)]?.raw??"";
      const an=numericValue(av,sheet.cells,sheet.tables??[],addressOf(a,p.col)),bn=numericValue(bv,sheet.cells,sheet.tables??[],addressOf(b,p.col));
      const cmp=an!==null&&bn!==null?an-bn:av.localeCompare(bv,undefined,{numeric:true,sensitivity:"base"});
      return descending?-cmp:cmp;
    });
    const cells={...sheet.cells};
    const sourceRows=rows.map(r=>Array.from({length:sheet.columnCount},(_,c)=>sheet.cells[addressOf(r,c)]));
    rows.forEach((_,i)=>{for(let c=0;c<sheet.columnCount;c++){const target=addressOf(i+1,c),value=sourceRows[i][c];if(value)cells[target]=value;else delete cells[target];}});
    if(header!==undefined&&sheet.cells[addressOf(0,p.col)]?.raw!==header)return;
    updateCells(cells);
  };
  const action=(name:string)=>{
    switch(name){
      case "new":{const created=useWorkbookStore.getState().createWorkbook("blank");navigate(`/workbook/${created}`);break;}
      case "back":navigate("/");break;
      case "save":setShowSaved();break;
      case "export-xlsx":void exportFile("xlsx");break;case "export-csv":void exportFile("csv");break;case "export-pdf":void exportFile("pdf");break;
      case "print":{
        const printStyle=document.createElement("style");
        printStyle.dataset.gridnovaPrint="true";
        printStyle.textContent=`@media print{@page{size:${workbook.orientation};margin:12mm}}`;
        document.head.appendChild(printStyle);window.print();printStyle.remove();break;
      }
      case "undo":undo();break;case "redo":redo();break;
      case "cut":copySelection(true);break;case "copy":copySelection(false);break;case "paste":if(clipboard)paste(activeCell);break;
      case "bold":setStyle({bold:!activeStyle.bold},selectedAddresses);break;
      case "italic":setStyle({italic:!activeStyle.italic},selectedAddresses);break;
      case "underline":setStyle({underline:!activeStyle.underline},selectedAddresses);break;
      case "strike":setStyle({strike:!activeStyle.strike},selectedAddresses);break;
      case "align-left":setStyle({align:"left"},selectedAddresses);break;case "align-center":setStyle({align:"center"},selectedAddresses);break;case "align-right":setStyle({align:"right"},selectedAddresses);break;
      case "wrap":setStyle({wrap:!activeStyle.wrap},selectedAddresses);break;
      case "percent":setStyle({numberFormat:"Percentage %"},selectedAddresses);break;case "currency-inr":setStyle({numberFormat:"Currency ₹"},selectedAddresses);break;case "comma":setStyle({numberFormat:"Number"},selectedAddresses);break;
      case "font-up":{const v=Number(activeStyle.fontSize??11);setStyle({fontSize:Math.min(36,v+1)},selectedAddresses);break;}
      case "font-down":{const v=Number(activeStyle.fontSize??11);setStyle({fontSize:Math.max(10,v-1)},selectedAddresses);break;}
      case "decimal-up":{const v=Number(activeStyle.decimalPlaces??0);setStyle({numberFormat:activeStyle.numberFormat==="General"||!activeStyle.numberFormat?"Number":activeStyle.numberFormat,decimalPlaces:Math.min(8,v+1)},selectedAddresses);break;}
      case "decimal-down":{const v=Number(activeStyle.decimalPlaces??2);setStyle({numberFormat:activeStyle.numberFormat==="General"||!activeStyle.numberFormat?"Number":activeStyle.numberFormat,decimalPlaces:Math.max(0,v-1)},selectedAddresses);break;}
      case "autosum":case "formula-sum":{
        const p=parseAddress(activeCell);if(!p)break;
        const formula=p.row>0?`=SUM(${addressOf(0,p.col)}:${addressOf(p.row-1,p.col)})`:p.col>0?`=SUM(A1:${addressOf(p.row,p.col-1)})`:"=SUM(A1:A1)";
        setCell(activeCell,formula);break;
      }
      case "formula-average":case "formula-count":case "formula-min":case "formula-max":case "formula-if":{
        const fn:Record<string,string>={"formula-average":"AVERAGE","formula-count":"COUNT","formula-min":"MIN","formula-max":"MAX","formula-if":"IF"};
        setCell(activeCell,name==="formula-if"?'=IF(A1>50,"Pass","Fail")':`=${fn[name]}(${selectedRangeString()})`);break;
      }
      case "sort-az":sortData(false);break;case "sort-za":sortData(true);break;
      case "find":{
        const find=window.prompt("Find text");if(find===null||!find)break;
        const replace=window.prompt(`Replace “${find}” with (leave empty to find only)`);const cells={...sheet.cells};
        const found=Object.entries(cells).filter(([,c])=>c.raw.toLowerCase().includes(find.toLowerCase()));
        if(replace!==null&&replace!=="")found.forEach(([a,c])=>cells[a]={...c,raw:c.raw.split(find).join(replace)});
        updateCells(cells);if(found[0])selectCell(found[0][0]);window.alert(`${found.length} matching cell(s) found.`);break;
      }
      case "clear":{const cells={...sheet.cells};selectedAddresses.forEach(a=>cells[a]={...(cells[a]??{raw:""}),raw:""});updateCells(cells);break;}
      case "clear-format":{selectedAddresses.forEach(a=>useWorkbookStore.getState().clearFormat(a));break;}
      case "merge":{
        if(selectedAddresses.length<2){window.alert("Select two or more cells before merging.");break;}
        const startRow=Math.min(selection!.startRow,selection!.endRow),endRow=Math.max(selection!.startRow,selection!.endRow);
        const startCol=Math.min(selection!.startCol,selection!.endCol),endCol=Math.max(selection!.startCol,selection!.endCol);
        const anchor=addressOf(startRow,startCol),cells={...sheet.cells},backup:Record<string,CellData>={};
        selectedAddresses.forEach(a=>{backup[a]=cells[a]?structuredClone(cells[a]):{raw:""};});
        cells[anchor]={...(cells[anchor]??{raw:""}),mergeRows:endRow-startRow+1,mergeCols:endCol-startCol+1,mergeBackup:backup};
        selectedAddresses.filter(a=>a!==anchor).forEach(a=>{cells[a]={raw:"",mergedInto:anchor};});
        updateCells(cells);selectCell(anchor,{startRow,startCol,endRow,endCol});setStyle({align:"center"},[anchor]);break;
      }
      case "unmerge":{
        const anchor=sheet.cells[activeCell]?.mergedInto??activeCell;
        const merged=sheet.cells[anchor];
        if(!merged?.mergeBackup){window.alert("The active cell is not part of a merged range.");break;}
        const cells={...sheet.cells,...merged.mergeBackup};const restored={...cells[anchor]};
        delete restored.mergeRows;delete restored.mergeCols;delete restored.mergeBackup;cells[anchor]=restored;
        updateCells(cells);selectCell(anchor);break;
      }
      case "row-above":shiftRowsCols("row",parseAddress(activeCell)?.row??0,"insert");break;
      case "row-below":shiftRowsCols("row",(parseAddress(activeCell)?.row??0)+1,"insert");break;
      case "delete-row":shiftRowsCols("row",parseAddress(activeCell)?.row??0,"delete");break;
      case "col-left":shiftRowsCols("column",parseAddress(activeCell)?.col??0,"insert");break;
      case "col-right":shiftRowsCols("column",(parseAddress(activeCell)?.col??0)+1,"insert");break;
      case "delete-col":shiftRowsCols("column",parseAddress(activeCell)?.col??0,"delete");break;
      case "format-table":openTableDialog("create");break;
      case "table-resize":openTableDialog("resize");break;
      case "table-rename":{
        if(!activeTable)break;
        const entered=window.prompt("Table name",activeTable.name);
        if(entered===null)break;
        const next=normalizeTableName(entered);
        if((sheet.tables??[]).some(table=>table.id!==activeTable.id&&table.name.toLowerCase()===next.toLowerCase())){window.alert("A table with that name already exists.");break;}
        updateActiveTable({name:next});break;
      }
      case "table-toggle-headers":if(activeTable)updateActiveTable({hasHeaders:!activeTable.hasHeaders});break;
      case "table-toggle-totals":if(activeTable)setTableTotalRow(!activeTable.showTotalRow);break;
      case "table-toggle-first-column":if(activeTable)updateActiveTable({firstColumn:!activeTable.firstColumn});break;
      case "table-toggle-last-column":if(activeTable)updateActiveTable({lastColumn:!activeTable.lastColumn});break;
      case "table-toggle-banded-rows":if(activeTable)updateActiveTable({bandedRows:!activeTable.bandedRows});break;
      case "table-toggle-banded-columns":if(activeTable)updateActiveTable({bandedColumns:!activeTable.bandedColumns});break;
      case "table-toggle-filter-buttons":if(activeTable)updateActiveTable({showFilterButtons:!activeTable.showFilterButtons});break;
      case "table-clear-style":if(activeTable)updateActiveTable({styleId:"TableStyleLight1"});break;
      case "table-clear-filters":if(activeTable)updateActiveTable({filters:{}});break;
      case "table-remove-duplicates":removeActiveTableDuplicates();break;
      case "table-convert-range":convertActiveTableToRange();break;
      case "insert-function":setFormulaLibraryOpen(true);break;
      case "share-open":case "share-manage-access":setShareOpen(true);break;
      case "share-copy-link":copyShareLink("workbook");break;
      case "share-section-link":copyShareLink("range");break;
      case "share-table-link":copyShareLink("table");break;
      case "pivot-table":createPivotTable();break;
      case "external-pivot":window.alert("External data source connections are not configured in this local workbook. Select a range and use From Table/Range instead.");break;
      case "forms-new":{
        addSheet();
        const newSheetId=useWorkbookStore.getState().activeSheetId;
        if(newSheetId){
          renameSheet(newSheetId,"Form Responses");
          updateSheet(newSheetId,{cells:{A1:{raw:"Submission ID",style:{bold:true,color:"#ffffff",backgroundColor:"#038387"}},B1:{raw:"Submitted At",style:{bold:true,color:"#ffffff",backgroundColor:"#038387"}},C1:{raw:"Response",style:{bold:true,color:"#ffffff",backgroundColor:"#038387"}}}});
        }
        window.open("https://forms.office.com","_blank","noopener,noreferrer");
        window.alert("A Form Responses sheet has been created. Microsoft Forms is not connected for automatic response sync in this local workbook.");
        break;
      }
      case "forms-open":window.open("https://forms.office.com", "_blank", "noopener,noreferrer");break;
      case "forms-field":insertFloatingElement({type:"text",text:"Form response",fontSize:14,width:240,height:64,fill:"#ffffff",stroke:"#038387"});break;
      case "picture-cell-device":setPicturePlacement("cell");pictureRef.current?.click();break;
      case "picture-over-device":setPicturePlacement("over");pictureRef.current?.click();break;
      case "picture-cell-stock":case "picture-over-online":case "picture-stock":case "picture-online":window.alert("Stock and online image providers are not connected. Choose This Device to insert a local image.");break;
      case "shape-rectangle":case "shape-ellipse":case "shape-line":case "shape-arrow":insertFloatingElement({type:"shape",shape:name.slice("shape-".length) as FloatingElementData["shape"]});break;
      case "text-box":insertFloatingElement({type:"text",text:"Type text here",fontSize:16,color:"#323130",width:280,height:110,fill:"#ffffff",stroke:"#0F6CBD"});break;
      case "slicer":openSlicer();break;
      case "comments-panel":setCommentsPanelOpen(true);break;
      case "chart-bar":case "chart-line":case "chart-pie":{
        const type=name.replace("chart-","") as FloatingChartData["type"];
        const title=type==="bar"?"Sales by category":type==="line"?"Trend":"Distribution";
        addChart({id:crypto.randomUUID(),type,title,range:selectedRangeString(),x:Math.max(20,window.innerWidth-570),y:90,width:420,height:290});break;
      }
      case "hyperlink":setCellDialogKind("link");break;
      case "comment":setCellDialogKind("comment");break;
      case "checkbox":{
        const cells={...sheet.cells},checked=!selectedAddresses.every(address=>cells[address]?.checkbox&&cells[address]?.raw==="TRUE");
        selectedAddresses.forEach(address=>{const old=cells[address]??{raw:""};cells[address]={...old,raw:checked?"TRUE":"FALSE",checkbox:true};});
        updateSheet(sheet.id,{cells});break;
      }
      case "date":updateSheet(sheet.id,{cells:{...sheet.cells,[activeCell]:{...(sheet.cells[activeCell]??{raw:""}),raw:new Date().toISOString().slice(0,10),datePicker:true,style:{...(sheet.cells[activeCell]?.style??{}),numberFormat:"Date"}}}});break;
      case "orientation-portrait":setOrientation("portrait");break;case "orientation-landscape":setOrientation("landscape");break;
      case "gridlines":updateSheet(sheet.id,{showGridlines:!sheet.showGridlines});break;case "headings":updateSheet(sheet.id,{showHeadings:!sheet.showHeadings});break;
      case "toggle-formulas":setShowFormulas(!showFormulas);break;
      case "filters":updateSheet(sheet.id,{filters:!sheet.filters});break;
      case "duplicates":{
        const p=parseAddress(activeCell);if(!p)break;
        const seen=new Set<string>(),cells={...sheet.cells};let writeRow=1;
        for(let r=1;r<sheet.rowCount;r++){
          const value=cells[addressOf(r,p.col)]?.raw??"";
          if(value&&seen.has(value))continue;
          if(value)seen.add(value);
          if(writeRow!==r){for(let c=0;c<sheet.columnCount;c++){const source=cells[addressOf(r,c)],target=addressOf(writeRow,c);if(source)cells[target]=source;else delete cells[target];}}
          writeRow++;
        }
        for(let r=writeRow;r<sheet.rowCount;r++)for(let c=0;c<sheet.columnCount;c++)delete cells[addressOf(r,c)];
        updateCells(cells);break;
      }
      case "freeze":updateSheet(sheet.id,{freezeTopRow:!sheet.freezeTopRow});break;
      case "zoom-in":useWorkbookStore.getState().setZoom(zoom+10);break;case "zoom-out":useWorkbookStore.getState().setZoom(zoom-10);break;
      case "import":fileRef.current?.click();break;
      default:
        if(name.startsWith("table-style-")&&activeTable)updateActiveTable({styleId:name.slice("table-style-".length)});
        else if(name.startsWith("shape-gallery:")){
          const shape=name.slice("shape-gallery:".length) as FloatingElementData["shape"];
          insertFloatingElement({type:"shape",shape});
        }
        else if(name.startsWith("chart-gallery:")){
          const [,kindText,indexText]=name.split(":");
          const type=kindText as FloatingChartData["type"],index=Number(indexText);
          const variants:Record<string,string[]>={
            column:["Clustered Column","Stacked Column","100% Stacked Column","3-D Clustered Column","3-D Stacked Column"],
            line:["Line","Stacked Line","Line with Markers","Area","Stacked Area"],
            scatter:["Scatter","Scatter with Smooth Lines and Markers","Scatter with Straight Lines","Bubble"],
            pie:["Pie","Pie of Pie","3-D Pie","Doughnut"],
            bar:["Clustered Bar","Stacked Bar","100% Stacked Bar","3-D Clustered Bar"],
            statistical:["Histogram","Pareto","Box & Whisker"],
            combo:["Clustered Column - Line","Clustered Column - Line on Secondary Axis","Stacked Area - Clustered Column","Stacked Column - Line"],
            treemap:["Treemap","Sunburst","Waterfall","Funnel","Radar"]
          };
          const variant=variants[type]?.[index];
          if(!variant)break;
          const chart:FloatingChartData={id:crypto.randomUUID(),type,title:variant,variant,range:selectedRangeString(),x:Math.max(20,window.innerWidth-570),y:90,width:440,height:300};
          addChart(chart);
        }
        else if(name==="more-charts")window.alert("The chart gallery includes all available chart subtypes in this workbook build.");
        break;
    }
    setShowSaved();
  };
  const importFile=async(file:File)=>{
    try{
      const ext=file.name.split(".").pop()?.toLowerCase();
      const tools=await import("../utils/exportUtils");
      const rows=ext==="csv"?await tools.importCsvText(await file.text()):await tools.importExcelFile(file);
      updateSheet(sheet.id,{cells:tools.rowsToCells(rows),rowCount:Math.max(100,rows.length),columnCount:Math.max(26,...rows.map(r=>r.length))});
      selectCell("A1");setShowSaved();
    }catch(e){window.alert(e instanceof Error?e.message:"Unable to import this file.");}
    finally{if(fileRef.current)fileRef.current.value="";}
  };
  return <main className="app-shell workbook-shell flex h-screen h-[100dvh] w-screen min-h-0 flex-col overflow-hidden">
    <header className="no-print z-30 flex h-[58px] shrink-0 items-center gap-3 border-b border-[#e3ebe5] bg-white px-4">
      <button onMouseDown={e=>e.preventDefault()} onClick={()=>navigate("/")} className="grid h-9 w-9 shrink-0 place-items-center rounded-xl text-[#718178] hover:bg-[#f0f6f2]" title="Back to dashboard"><ArrowLeft size={18}/></button>
      <div className="grid h-9 w-9 shrink-0 place-items-center rounded-xl bg-[#e8f5ed] text-[#107c41]"><FileSpreadsheet size={19}/></div>
      <div className="min-w-0 flex-1">{renaming?<input autoFocus value={titleValue} onChange={e=>setTitleValue(e.target.value)} onBlur={commitTitle} onKeyDown={e=>{if(e.key==="Enter")commitTitle();if(e.key==="Escape")setRenaming(false)}} className="w-full max-w-[330px] rounded-md border border-[#abd5b8] px-2 py-1 text-sm font-bold text-[#284331] outline-none"/>:<button onClick={()=>setRenaming(true)} className="max-w-[360px] truncate text-left text-sm font-bold text-[#284331]">{workbook.title}<ChevronDown size={13} className="ml-1 inline text-[#93a197]"/></button>}<div className="text-[10px] text-[#95a198]">GridNova Excel · {sheet.name}</div></div>
      <div className="hidden items-center gap-1.5 rounded-full bg-[#edf8f0] px-3 py-1.5 text-[10px] font-semibold text-[#16864b] sm:flex">{savePulse?<Check size={12}/>:<Cloud size={12}/>} {savePulse?"Saved just now":"Saved locally"}</div>
      <div className="ml-auto flex items-center gap-1.5"><button onMouseDown={e=>e.preventDefault()} onClick={undo} className="grid h-8 w-8 place-items-center rounded-lg text-[#7b8c81] hover:bg-[#eff5f0]" title="Undo"><span className="text-sm">↶</span></button><button onMouseDown={e=>e.preventDefault()} onClick={redo} className="grid h-8 w-8 place-items-center rounded-lg text-[#7b8c81] hover:bg-[#eff5f0]" title="Redo"><span className="text-sm">↷</span></button><button onMouseDown={e=>e.preventDefault()} onClick={()=>setShareOpen(true)} className="rounded-full bg-[#087df0] px-4 py-2 text-xs font-bold text-white shadow-sm hover:bg-[#006ad2]">Share</button></div>
    </header>
    <Ribbon activeTab={tab} setActiveTab={setTab} onAction={action} onFormat={applyCellFormat} style={activeStyle} sheet={sheet} showFormulas={showFormulas} setShowFormulas={setShowFormulas} zoom={zoom} setZoom={setZoom} orientation={workbook.orientation} setOrientation={setOrientation} activeTable={activeTable}/>
    <FormulaBar address={activeCell} value={formulaValue} onChange={setFormulaValue} onCommit={()=>{setCell(activeCell,formulaValue);setShowSaved()}} onCancel={()=>setFormulaValue(activeRaw)} onInsertFunction={()=>setFormulaLibraryOpen(true)} onNavigate={address=>selectCell(address)} onFocus={()=>setFormulaBarFocused(true)} onBlur={()=>setFormulaBarFocused(false)}/>
    <div className="relative flex min-h-0 flex-1 overflow-hidden">
      <SpreadsheetGrid sheet={sheet} activeCell={activeCell} selection={selection} showFormulas={showFormulas} zoom={zoom} formulaBarFocused={formulaBarFocused} onSelect={selectCell}/>
      <CellImageOverlays sheet={sheet}/>
      {workbook.charts.map(chart=><Suspense key={chart.id} fallback={null}><FloatingChart chart={chart} sheet={sheet} onUpdate={updateChart} onDelete={deleteChart}/></Suspense>)}
      {(workbook.elements??[]).map(element=><FloatingElement key={element.id} element={element} onUpdate={updateElement} onDelete={deleteElement}/>)}
      <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)void importFile(f)}}/>
      <input ref={pictureRef} type="file" accept="image/*" className="hidden" onChange={event=>{const file=event.target.files?.[0];if(file)importPicture(file);if(pictureRef.current)pictureRef.current.value="";}}/>
    </div>
    <SheetBar sheets={workbook.sheets} activeId={sheet.id} onSelect={setActiveSheet} onAdd={addSheet} onRename={renameSheet} onDelete={id=>{if(window.confirm("Delete this sheet?"))deleteSheet(id)}}/>
    <footer className="no-print flex h-8 shrink-0 items-center justify-between border-t border-[#e2e9e4] bg-white px-4 text-[10px] text-[#829087]"><div className="flex items-center gap-4"><span>Ready</span><span>Count: <b className="text-[#52665a]">{stats.count}</b></span><span>Sum: <b className="text-[#52665a]">{stats.sum.toLocaleString(undefined,{maximumFractionDigits:2})}</b></span><span>Average: <b className="text-[#52665a]">{stats.average.toLocaleString(undefined,{maximumFractionDigits:2})}</b></span></div><div>{sheet.showGridlines?"Gridlines on":"Gridlines off"} · {zoom}%</div></footer>
    <ShareModal open={shareOpen} onClose={()=>setShareOpen(false)} workbookId={workbook.id} title={workbook.title} onExport={exportFile}/>
    <TableRangeDialog open={tableDialogOpen} title={tableDialogMode==="resize"?"Resize Table":"Create Table"} initialRange={tableDialogMode==="resize"&&activeTable?rangeLabel(activeTable.range):suggestedTableRange} initialHasHeaders={tableDialogMode==="resize"?(activeTable?.hasHeaders??true):true} submitLabel={tableDialogMode==="resize"?"Resize":"OK"} onClose={()=>setTableDialogOpen(false)} onSubmit={submitTableRange}/>
    <FormulaLibraryDialog open={formulaLibraryOpen} onClose={()=>setFormulaLibraryOpen(false)} onInsert={formula=>{setCell(activeCell,formula);setFormulaValue(formula);setFormulaBarFocused(true)}}/>
    <CellInsertDialog open={!!cellDialogKind} kind={cellDialogKind??"link"} initialValue={cellDialogKind==="comment"?(sheet.cells[activeCell]?.comment??""):(sheet.cells[activeCell]?.hyperlink??"")} initialText={cellDialogKind==="link"?(sheet.cells[activeCell]?.raw??""):""} onClose={()=>setCellDialogKind(null)} onSubmit={(value,secondary)=>{if(cellDialogKind==="link"){const url=value.trim(),label=secondary?.trim()||url;updateSheet(sheet.id,{cells:{...sheet.cells,[activeCell]:{...(sheet.cells[activeCell]??{raw:label}),raw:label,hyperlink:url}}});}else{updateSheet(sheet.id,{cells:{...sheet.cells,[activeCell]:{...(sheet.cells[activeCell]??{raw:""}),comment:value}}});}setCellDialogKind(null);setShowSaved();}}/>
    {commentsPanelOpen&&<div className="grid-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)setCommentsPanelOpen(false);}}><section className="grid-dialog comments-panel-dialog" role="dialog" aria-modal="true" aria-labelledby="comments-panel-title"><header><div><span>WORKBOOK REVIEW</span><h2 id="comments-panel-title">Comments</h2><p>{Object.values(sheet.cells).filter(cell=>cell.comment).length} cell comments in {sheet.name}</p></div><button type="button" onClick={()=>setCommentsPanelOpen(false)} aria-label="Close comments">×</button></header><div className="comments-panel-list">{Object.entries(sheet.cells).filter(([,cell])=>cell.comment).map(([address,cell])=><button type="button" key={address} className="comments-panel-item" onMouseDown={event=>event.preventDefault()} onClick={()=>{selectCell(address);setCommentsPanelOpen(false);}}><strong>{address}</strong><span>{cell.comment}</span></button>)}{!Object.values(sheet.cells).some(cell=>cell.comment)&&<p className="comments-panel-empty">No comments in this sheet yet. Use Insert → New Comment to add one.</p>}</div><footer><button type="button" onClick={()=>setCommentsPanelOpen(false)}>Close</button></footer></section></div>}
    {slicerOpen&&activeTable&&<div className="grid-dialog-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)setSlicerOpen(false);}}><section className="grid-dialog slicer-dialog" role="dialog" aria-modal="true" aria-labelledby="slicer-title"><header><div><span>TABLE FILTER</span><h2 id="slicer-title">Insert Slicer</h2><p>Choose a table column and the values to display.</p></div><button type="button" onClick={()=>setSlicerOpen(false)} aria-label="Close slicer">×</button></header><div className="slicer-dialog-fields"><label>Column<select value={slicerColumn} onChange={event=>{const col=Number(event.target.value);setSlicerColumn(event.target.value);setSlicerValues(tableColumnValues(activeTable,col,sheet.cells));}}>{Array.from({length:activeTable.range.endCol-activeTable.range.startCol+1},(_,offset)=>activeTable.range.startCol+offset).map(col=><option key={col} value={String(col)}>{getTableHeaderLabel(activeTable,col,sheet.cells)}</option>)}</select></label><div className="slicer-dialog-values">{tableColumnValues(activeTable,Number(slicerColumn),sheet.cells).map(value=><label key={value||"__blank"}><input type="checkbox" checked={slicerValues.includes(value)} onChange={event=>setSlicerValues(current=>event.target.checked?[...current,value]:current.filter(item=>item!==value))}/>{value||"(Blanks)"}</label>)}</div></div><footer><button type="button" className="grid-dialog-secondary" onClick={()=>setSlicerOpen(false)}>Cancel</button><button type="button" onClick={applySlicer}>Apply Slicer</button></footer></section></div>}
  </main>;
}
