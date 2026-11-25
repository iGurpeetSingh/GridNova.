import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";
import { useNavigate, useParams } from "react-router-dom";
import { ArrowLeft, Check, ChevronDown, Cloud, FileSpreadsheet } from "lucide-react";
import Ribbon from "../components/Ribbon";
import SpreadsheetGrid from "../components/SpreadsheetGrid";
import FormulaBar from "../components/FormulaBar";
import SheetBar from "../components/SheetBar";
const FloatingChart = lazy(() => import("../components/FloatingChart"));
import ShareModal from "../components/ShareModal";
import { useWorkbookStore } from "../store/useWorkbookStore";
import type { CellData, CellRange, CellStyle, FloatingChartData } from "../types";
import { addressOf, parseAddress } from "../types";
import { numericValue } from "../utils/formulaEngine";


type Tab = "File"|"Home"|"Insert"|"Page Layout"|"Formulas"|"Data"|"View";
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
  const undo=useWorkbookStore(s=>s.undo);const redo=useWorkbookStore(s=>s.redo);
  const setShowFormulas=useWorkbookStore(s=>s.setShowFormulas);const setZoom=useWorkbookStore(s=>s.setZoom);
  const setOrientation=useWorkbookStore(s=>s.setOrientation);
  const [tab,setTab]=useState<Tab>("Home");const [formulaValue,setFormulaValue]=useState("");
  const [shareOpen,setShareOpen]=useState(false);const [renaming,setRenaming]=useState(false);const [titleValue,setTitleValue]=useState("");
  const [savePulse,setSavePulse]=useState(false);
  const fileRef=useRef<HTMLInputElement>(null);
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
    const vals=selectedAddresses.map(a=>numericValue(sheet.cells[a]?.raw??"",sheet.cells)).filter((v):v is number=>v!==null);
    return {count:vals.length,sum:vals.reduce((a,b)=>a+b,0),average:vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:0};
  },[sheet,selectedAddresses]);
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
    updateSheet(sheet.id,{cells,rowCount:axis==="row"?Math.max(1,sheet.rowCount+(mode==="insert"?1:-1)):sheet.rowCount,columnCount:axis==="column"?Math.max(1,sheet.columnCount+(mode==="insert"?1:-1)):sheet.columnCount});
  };
  const sortData=(descending:boolean)=>{
    const p=parseAddress(activeCell);if(!p)return;
    const header=sheet.cells[addressOf(0,p.col)]?.raw;
    const rows=Array.from({length:Math.max(0,sheet.rowCount-1)},(_,i)=>i+1);
    rows.sort((a,b)=>{
      const av=sheet.cells[addressOf(a,p.col)]?.raw??"",bv=sheet.cells[addressOf(b,p.col)]?.raw??"";
      const an=numericValue(av,sheet.cells),bn=numericValue(bv,sheet.cells);
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
      case "format-table":{
        const range=selection??{startRow:0,startCol:0,endRow:9,endCol:3};const cells={...sheet.cells};
        for(let r=range.startRow;r<=range.endRow;r++)for(let c=range.startCol;c<=range.endCol;c++){const a=addressOf(r,c),old=cells[a]??{raw:""};cells[a]={...old,style:{...old.style,...(r===range.startRow?{bold:true,color:"#ffffff",backgroundColor:"#107c41"}:{backgroundColor:r%2===0?"#f0f7f2":"#ffffff"})}};}
        updateCells(cells);break;
      }
      case "chart-bar":case "chart-line":case "chart-pie":{
        const range=selectedRangeString();const chart:FloatingChartData={id:crypto.randomUUID(),type:name.replace("chart-","") as FloatingChartData["type"],title:name==="chart-bar"?"Sales by category":name==="chart-line"?"Trend":"Distribution",range,x:Math.max(20,window.innerWidth-570),y:175,width:420,height:290};
        addChart(chart);break;
      }
      case "hyperlink":{
        const url=window.prompt("Enter hyperlink URL");if(!url)break;
        updateSheet(sheet.id,{cells:{...sheet.cells,[activeCell]:{...(sheet.cells[activeCell]??{raw:url}),raw:sheet.cells[activeCell]?.raw||url,hyperlink:(url.startsWith("http://")||url.startsWith("https://"))?url:`https://${url}`}}});break;
      }
      case "comment":{
        const comment=window.prompt("Add a note to this cell",sheet.cells[activeCell]?.comment??"");if(comment===null)break;
        updateSheet(sheet.id,{cells:{...sheet.cells,[activeCell]:{...(sheet.cells[activeCell]??{raw:""}),comment}}});break;
      }
      case "checkbox":updateSheet(sheet.id,{cells:{...sheet.cells,[activeCell]:{...(sheet.cells[activeCell]??{raw:""}),raw:"FALSE",checkbox:true}}});break;
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
      default:break;
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
    <Ribbon activeTab={tab} setActiveTab={setTab} onAction={action} onFormat={style=>setStyle(style,selectedAddresses)} style={activeStyle} sheet={sheet} showFormulas={showFormulas} setShowFormulas={setShowFormulas} zoom={zoom} setZoom={setZoom} orientation={workbook.orientation} setOrientation={setOrientation}/>
    <FormulaBar address={activeCell} value={formulaValue} onChange={setFormulaValue} onCommit={()=>{setCell(activeCell,formulaValue);setShowSaved()}}/>
    <div className="relative flex min-h-0 flex-1 overflow-hidden">
      <SpreadsheetGrid sheet={sheet} activeCell={activeCell} selection={selection} showFormulas={showFormulas} zoom={zoom} onSelect={selectCell}/>
      {workbook.charts.map(chart=><Suspense key={chart.id} fallback={null}><FloatingChart chart={chart} sheet={sheet} onUpdate={updateChart} onDelete={deleteChart}/></Suspense>)}
      <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)void importFile(f)}}/>
    </div>
    <SheetBar sheets={workbook.sheets} activeId={sheet.id} onSelect={setActiveSheet} onAdd={addSheet} onRename={renameSheet} onDelete={id=>{if(window.confirm("Delete this sheet?"))deleteSheet(id)}}/>
    <footer className="no-print flex h-8 shrink-0 items-center justify-between border-t border-[#e2e9e4] bg-white px-4 text-[10px] text-[#829087]"><div className="flex items-center gap-4"><span>Ready</span><span>Count: <b className="text-[#52665a]">{stats.count}</b></span><span>Sum: <b className="text-[#52665a]">{stats.sum.toLocaleString(undefined,{maximumFractionDigits:2})}</b></span><span>Average: <b className="text-[#52665a]">{stats.average.toLocaleString(undefined,{maximumFractionDigits:2})}</b></span></div><div>{sheet.showGridlines?"Gridlines on":"Gridlines off"} · {zoom}%</div></footer>
    <ShareModal open={shareOpen} onClose={()=>setShareOpen(false)} workbookId={workbook.id} title={workbook.title} onExport={exportFile}/>
  </main>;
}
