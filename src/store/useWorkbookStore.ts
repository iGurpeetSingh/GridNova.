import { create } from "zustand";
import { persist } from "zustand/middleware";
import type { CellData, CellMap, CellRange, CellStyle, ClipboardData, FloatingChartData, FloatingElementData, SheetData, TemplateId, WorkbookData, PivotTableDefinition } from "../types";
import { addressOf, columnLabel, emptySheet, emptyWorkbook, parseAddress } from "../types";
import { recalculateWorkbookIncremental } from "../lib/dependency-graph";
import { refreshWorkbookPivots } from "../utils/pivotEngine";

interface Store {
  workbooks: WorkbookData[]; activeWorkbookId: string | null; activeSheetId: string | null; autoSaveEnabled:boolean;
  activeCell: string; selection: CellRange | null; clipboard: ClipboardData | null;
  showFormulas: boolean; zoom: number; undoStack: WorkbookData[][]; redoStack: WorkbookData[][];
  createWorkbook: (template?: TemplateId) => string; importWorkbook:(workbook:WorkbookData)=>string; deleteWorkbook: (id: string) => void;
  duplicateWorkbook: (id: string) => string | null; renameWorkbook: (id: string, title: string) => void;
  updateWorkbook: (id: string, changes: Partial<WorkbookData>) => void;
  setActiveWorkbook: (id: string | null) => void; setActiveSheet: (id: string) => void;
  addSheet: () => void; renameSheet: (id: string, name: string) => void; deleteSheet: (id: string) => void;
  setCell: (address: string, raw: string) => void; setStyle: (style: Partial<CellStyle>, addresses?: string[]) => void;
  clearCell: (address: string) => void; clearFormat: (address: string) => void;
  setSelection: (range: CellRange | null, activeCell?: string) => void;
  setClipboard: (data: ClipboardData | null) => void; paste: (target: string) => void;
  updateSheet: (id: string, changes: Partial<SheetData>) => void;
  createPivotTable:(definition:PivotTableDefinition,outputCells:CellMap,outputRange:CellRange)=>string|null;
  addChart: (chart: FloatingChartData) => void; updateChart: (id: string, changes: Partial<FloatingChartData>) => void; deleteChart: (id: string) => void;
  addElement: (element: FloatingElementData) => void; updateElement: (id: string, changes: Partial<FloatingElementData>) => void; deleteElement: (id: string) => void;
  setShowFormulas: (v: boolean) => void; setZoom: (v: number) => void; setOrientation: (v: "portrait" | "landscape") => void; setAutoSaveEnabled:(enabled:boolean)=>void;
  undo: () => void; redo: () => void;
}
const clone = <T,>(v: T): T => structuredClone(v);
const makeTemplate = (kind: TemplateId): WorkbookData => {
  const titles: Record<TemplateId, string> = { blank: "Untitled Workbook", budget: "Monthly Budget", sales: "Sales Tracker", project: "Project Schedule", "sales-dashboard":"Sales Dashboard", "gst-invoice":"GST Invoice", "attendance-tracker":"Attendance Tracker", "advanced-demo":"Advanced Excel Engines Demo" };
  const wb = emptyWorkbook(titles[kind]); const s = wb.sheets[0];
  const put = (address: string, raw: string, style?: CellStyle) => { s.cells[address] = { raw, style }; };
  const header: CellStyle = { bold: true, color: "#ffffff", backgroundColor: "#128451", fontSize: 11, align: "left" };
  if (kind === "budget") {
    [["Category","Planned","Actual","Difference"],["Housing","18000","17500","=B2-C2"],["Food","8000","7600","=B3-C3"],["Travel","4000","2500","=B4-C4"],["Utilities","3000","2900","=B5-C5"],["Total","=SUM(B2:B5)","=SUM(C2:C5)","=SUM(D2:D5)"]].forEach((r,ri)=>r.forEach((v,ci)=>put(addressOf(ri,ci),v,ri===0?header:ri===5?{bold:true,backgroundColor:"#e8f5ee"}:ci>0?{numberFormat:"Currency ₹"}:undefined)));
  } else if (kind === "sales" || kind === "sales-dashboard") {
    [["Product","Units","Unit Price","Revenue","Share"],["Nova Desk","12","4500","=B2*C2","=D2/SUM(D2:D5)"],["Orbit Lamp","18","1800","=B3*C3","=D3/SUM(D2:D5)"],["Cloud Chair","7","6200","=B4*C4","=D4/SUM(D2:D5)"],["Pulse Stand","21","950","=B5*C5","=D5/SUM(D2:D5)"]].forEach((r,ri)=>r.forEach((v,ci)=>put(addressOf(ri,ci),v,ri===0?header:ci===4?{numberFormat:"Percentage %"}:ci===2||ci===3?{numberFormat:"Currency ₹"}:undefined)));
  } else if (kind === "project") {
    [["Task","Owner","Start Date","Due Date","Status","Done"],["Discovery","Maya","2026-11-03","2026-11-07","Complete","TRUE"],["Wireframes","Arjun","2026-11-08","2026-11-14","In Progress","FALSE"],["Development","Sam","2026-11-15","2026-11-28","Not Started","FALSE"],["QA","Maya","2026-11-29","2026-12-04","Not Started","FALSE"]].forEach((r,ri)=>r.forEach((v,ci)=>put(addressOf(ri,ci),v,ri===0?header:undefined)));
    for(let row=2;row<=5;row++){
      s.cells[`C${row}`]={...s.cells[`C${row}`],datePicker:true};
      s.cells[`D${row}`]={...s.cells[`D${row}`],datePicker:true};
      s.cells[`F${row}`]={...s.cells[`F${row}`],checkbox:true};
    }
  } else if(kind==="gst-invoice"){
    const rows=[
      ["GST TAX INVOICE","","","","","",""],
      ["Invoice No","INV-2026-001","Invoice Date","2026-11-26","GSTIN","07AAAAA0000A1Z5",""],
      ["Bill To","Example Customer","","","","",""],
      ["Item","HSN/SAC","Qty","Rate (₹)","Taxable Value (₹)","GST 18% (₹)","Total (₹)"],
      ["Web Development","998314","1","25000","=C5*D5","=E5*18%","=E5+F5"],
      ["UI Design","998391","2","7500","=C6*D6","=E6*18%","=E6+F6"],
      ["Subtotal","","","","=SUM(E5:E6)","=SUM(F5:F6)","=SUM(G5:G6)"],
      ["Grand Total","","","","","","=G7"]
    ];
    rows.forEach((row,ri)=>row.forEach((value,ci)=>put(addressOf(ri,ci),value,ri===0?{bold:true,fontSize:18,color:"#ffffff",backgroundColor:"#0F6CBD"}:ri===3?header:ri>=6?{bold:true,backgroundColor:"#EAF2F8"}:undefined)));
    for(let col=3;col<=6;col++)for(let row=4;row<=7;row++){
      const key=addressOf(row,col);if(s.cells[key])s.cells[key].style={...(s.cells[key].style??{}),numberFormat:"Currency ₹"};
    }
    s.rowCount=30;s.columnCount=7;
  } else if(kind==="attendance-tracker"){
    const rows=[
      ["ATTENDANCE TRACKER","","","","","",""],
      ["Employee","Employee ID","Date","Status","Check In","Check Out","Hours"],
      ["Aarav","EMP-001","2026-11-26","Present","9","18","=F3-E3"],
      ["Diya","EMP-002","2026-11-26","Present","9.25","18.17","=F4-E4"],
      ["Kabir","EMP-003","2026-11-26","Leave","","","=0"],
      ["Meera","EMP-004","2026-11-26","Absent","","","=0"],
      ["Summary","Present","=COUNTIF(D3:D6,\"Present\")","Absent","=COUNTIF(D3:D6,\"Absent\")","Leave","=COUNTIF(D3:D6,\"Leave\")"]
    ];
    rows.forEach((row,ri)=>row.forEach((value,ci)=>put(addressOf(ri,ci),value,ri===0?{bold:true,fontSize:17,color:"#ffffff",backgroundColor:"#107C41"}:ri===1?header:ri===6?{bold:true,backgroundColor:"#E8F5EE"}:undefined)));
    for(let row=2;row<=5;row++){s.cells[addressOf(row,2)]={...s.cells[addressOf(row,2)],datePicker:true};s.cells[addressOf(row,3)]={...s.cells[addressOf(row,3)],validationOptions:["Present","Absent","Leave","Holiday"]};}
    s.rowCount=40;s.columnCount=7;
  } else if(kind==="advanced-demo"){
    const rows=[
      ["Order ID","Region","Product","Units","Unit Price","Revenue","Payment Status","Month"],
      ["ORD-1001","West","Nova Chair","12","750","=D2*E2","Paid","Jan"],
      ["ORD-1002","North","Orbit Lamp","18","420","=D3*E3","Pending","Jan"],
      ["ORD-1003","South","Cloud Desk","7","2400","=D4*E4","Paid","Feb"],
      ["ORD-1004","East","Nova Chair","21","750","=D5*E5","Cancelled","Feb"],
      ["ORD-1005","West","Pulse Stand","32","350","=D6*E6","Paid","Mar"],
      ["ORD-1006","North","Cloud Desk","9","2400","=D7*E7","Pending","Mar"],
      ["ORD-1007","South","Orbit Lamp","14","420","=D8*E8","Paid","Apr"],
      ["ORD-1008","East","Nova Chair","16","750","=D9*E9","Pending","Apr"],
      ["ORD-1009","West","Cloud Desk","5","2400","=D10*E10","Paid","May"],
      ["ORD-1010","North","Pulse Stand","27","350","=D11*E11","Cancelled","May"],
      ["ORD-1011","South","Nova Chair","11","750","=D12*E12","Paid","Jun"],
      ["ORD-1012","East","Orbit Lamp","23","420","=D13*E13","Pending","Jun"]
    ];
    rows.forEach((row,ri)=>row.forEach((value,ci)=>{
      put(addressOf(ri,ci),value,ri===0?header:ci===4||ci===5?{numberFormat:"Currency ₹",decimalPlaces:0}:undefined);
    }));
    s.tables=[{id:crypto.randomUUID(),name:"AdvancedSales",range:{startRow:0,startCol:0,endRow:12,endCol:7},hasHeaders:true,showTotalRow:false,firstColumn:false,lastColumn:false,bandedRows:true,bandedColumns:false,showFilterButtons:true,styleId:"TableStyleMedium2",filters:{}}];
    s.validationRules=[{id:crypto.randomUUID(),range:{startRow:1,startCol:6,endRow:12,endCol:6},type:"List",listSource:"Paid, Pending, Cancelled",allowBlank:false,showInputMessage:true,inputTitle:"Payment status",inputMessage:"Select Paid, Pending or Cancelled.",showErrorAlert:true,errorStyle:"Stop",errorTitle:"Invalid payment status",errorMessage:"Choose a value from the dropdown list."}];
    s.conditionalFormattingRules=[{id:crypto.randomUUID(),range:{startRow:1,startCol:5,endRow:12,endCol:5},type:"dataBar",enabled:true,priority:1,color:"#63A5DB",gradient:true}];
    put("J1","PROFIT CALCULATOR",{bold:true,color:"#ffffff",backgroundColor:"#0F6CBD",fontSize:13});
    put("J2","Units Sold",{bold:true});put("K2","120",{numberFormat:"Number",decimalPlaces:0});
    put("J3","Selling Price",{bold:true});put("K3","750",{numberFormat:"Currency ₹",decimalPlaces:0});
    put("J4","Unit Cost",{bold:true});put("K4","350",{numberFormat:"Currency ₹",decimalPlaces:0});
    put("J5","Fixed Costs",{bold:true});put("K5","25000",{numberFormat:"Currency ₹",decimalPlaces:0});
    put("J6","Net Profit",{bold:true,backgroundColor:"#E8F5EE"});put("K6","=K2*(K3-K4)-K5",{bold:true,backgroundColor:"#E8F5EE",numberFormat:"Currency ₹",decimalPlaces:0});
    put("J7","Goal Seek: Set K6 to 50000 by changing K2",{color:"#605E5C",fontSize:9});
    put("J9","Product Lookup",{bold:true});put("K9","Nova Chair");
    put("J10","Revenue",{bold:true});put("K10",'=XLOOKUP(K9,C2:C13,F2:F13,"Not found")',{numberFormat:"Currency ₹",decimalPlaces:0});
    put("J11","West · Paid Revenue",{bold:true});put("K11",'=SUMIFS(F2:F13,B2:B13,"West",G2:G13,"Paid")',{numberFormat:"Currency ₹",decimalPlaces:0});
    put("J12","Live Pending Orders",{bold:true,color:"#ffffff",backgroundColor:"#107C41"});
    put("K12",'=FILTER(A2:G13,G2:G13="Pending","No pending orders")');
    s.sparklines={I2:{id:crypto.randomUUID(),type:"line",dataRange:"F2:F13",location:"I2",color:"#0F6CBD",negativeColor:"#C4314B",showHighPoint:true,showLowPoint:true,showMarkers:true}};
    s.rowCount=40;s.columnCount=20;
  }
  wb.modifiedAt = Date.now(); return wb;
};
const activeWorkbook = (s: Store) => s.workbooks.find(w => w.id === s.activeWorkbookId);
const activeSheet = (s: Store) => activeWorkbook(s)?.sheets.find(x => x.id === s.activeSheetId);
const snapshot = (s: Store) => ({ undoStack: [...s.undoStack, clone(s.workbooks)].slice(-40), redoStack: [] as WorkbookData[][] });
const mutateActive = (s: Store, updater: (w: WorkbookData, sh: SheetData) => WorkbookData): Partial<Store> => {
  const w = activeWorkbook(s), sh = activeSheet(s); if (!w || !sh) return {};
  return { ...snapshot(s), workbooks: s.workbooks.map(item => item.id === w.id ? { ...updater(w, sh), modifiedAt: Date.now() } : item) };
};
const recalculateAfterMutation=(state:Store,result:Partial<Store>,sheetId:string,addresses:string[]):Partial<Store>=>{
  const workbookId=state.activeWorkbookId;
  if(!workbookId||!addresses.length)return result;
  const original=(result.workbooks??state.workbooks).find(item=>item.id===workbookId);
  if(!original)return result;
  const sourceHasPivots=original.sheets.some(destination=>(destination.pivotTables??[]).some(definition=>definition.sourceSheetId===sheetId));
  const base=sourceHasPivots?refreshWorkbookPivots(original):original;
  if(base.calculationMode==="Manual")return sourceHasPivots?{...result,workbooks:(result.workbooks??state.workbooks).map(item=>item.id===workbookId?base:item)}:result;
  const recalculated=recalculateWorkbookIncremental(base,addresses.map(address=>({sheetId,address}))).workbook;
  return {...result,workbooks:(result.workbooks??state.workbooks).map(item=>item.id===workbookId?recalculated:item)};
};
export const useWorkbookStore = create<Store>()(persist((set, get) => ({
  workbooks: [], activeWorkbookId: null, activeSheetId: null, activeCell: "A1", selection: null, autoSaveEnabled:true,
  clipboard: null, showFormulas: false, zoom: 100, undoStack: [], redoStack: [],
  createWorkbook: (template = "blank") => {
    const wb = makeTemplate(template);
    set(s => ({ workbooks: [wb, ...s.workbooks], activeWorkbookId: wb.id, activeSheetId: wb.sheets[0].id, activeCell: "A1", selection: null }));
    return wb.id;
  },
  importWorkbook: source => {
    const copy=clone(source);
    copy.id=crypto.randomUUID();copy.title=copy.title||"Imported Workbook";copy.createdAt=Date.now();copy.modifiedAt=Date.now();
    copy.sheets=(copy.sheets??[]).map(sh=>({...sh,id:crypto.randomUUID()}));
    if(!copy.sheets.length)copy.sheets=[emptySheet()];
    set(state=>({workbooks:[copy,...state.workbooks],activeWorkbookId:copy.id,activeSheetId:copy.sheets[0].id,activeCell:"A1",selection:null}));
    return copy.id;
  },
  deleteWorkbook: id => set(s => {
    const list = s.workbooks.filter(w => w.id !== id);
    const next = list[0];
    return { workbooks: list, activeWorkbookId: s.activeWorkbookId === id ? next?.id ?? null : s.activeWorkbookId,
      activeSheetId: s.activeWorkbookId === id ? next?.sheets[0]?.id ?? null : s.activeSheetId };
  }),
  duplicateWorkbook: id => {
    const source = get().workbooks.find(w => w.id === id); if (!source) return null;
    const copy = clone(source); copy.id = crypto.randomUUID(); copy.title = source.title + " copy"; copy.createdAt = Date.now(); copy.modifiedAt = Date.now();
    copy.sheets = copy.sheets.map(sh => ({ ...sh, id: crypto.randomUUID() }));
    set(s => ({ workbooks: [copy, ...s.workbooks] })); return copy.id;
  },
  renameWorkbook: (id,title) => set(s => ({ workbooks: s.workbooks.map(w => w.id === id ? { ...w,title: title.trim() || w.title,modifiedAt:Date.now() } : w) })),
  updateWorkbook: (id,changes) => set(s => ({ workbooks:s.workbooks.map(w=>w.id===id?{...w,...changes,modifiedAt:Date.now()}:w) })),
  setActiveWorkbook: id => set(s => { const w = s.workbooks.find(x => x.id === id); return { activeWorkbookId:id, activeSheetId:w?.sheets[0]?.id ?? null, activeCell:"A1",selection:null }; }),
  setActiveSheet: id => set({ activeSheetId:id,activeCell:"A1",selection:null }),
  addSheet: () => set(s => {
    const w = activeWorkbook(s); if (!w) return {};
    const sh = emptySheet(`Sheet ${w.sheets.length + 1}`);
    return { ...snapshot(s), workbooks:s.workbooks.map(x=>x.id===w.id?{...x,sheets:[...x.sheets,sh],modifiedAt:Date.now()}:x),activeSheetId:sh.id };
  }),
  renameSheet: (id,name) => set(s => {
    const w=activeWorkbook(s); if(!w) return {};
    return { workbooks:s.workbooks.map(x=>x.id===w.id?{...x,sheets:x.sheets.map(sh=>sh.id===id?{...sh,name:name.trim()||sh.name}:sh),modifiedAt:Date.now()}:x) };
  }),
  deleteSheet: id => set(s => {
    const w=activeWorkbook(s); if(!w || w.sheets.length<2) return {};
    const sheets=w.sheets.filter(sh=>sh.id!==id);
    return { ...snapshot(s),workbooks:s.workbooks.map(x=>x.id===w.id?{...x,sheets,modifiedAt:Date.now()}:x),activeSheetId:s.activeSheetId===id?sheets[0].id:s.activeSheetId };
  }),
  setCell: (address,raw) => set(s => {const sh=activeSheet(s);if(!sh)return {};const key=address.toUpperCase();const result=mutateActive(s,(w,ss)=>({...w,sheets:w.sheets.map(x=>x.id===ss.id?{...x,cells:{...x.cells,[key]:{...x.cells[key],raw}}}:x)}));return recalculateAfterMutation(s,result,sh.id,[key]);}),
  setStyle: (style,addresses) => set(s => {
    const sh=activeSheet(s); if(!sh) return {};
    const targets=addresses?.length?addresses:[s.activeCell];
    return mutateActive(s,(w,ss)=>({...w,sheets:w.sheets.map(x=>x.id===ss.id?{...x,cells:targets.reduce((acc,a)=>{const key=a.toUpperCase();acc[key]={...(acc[key]??{raw:""}),style:{...(acc[key]?.style??{}),...style}};return acc;},{...x.cells})}:x)}));
  }),
  clearCell: address => set(s=>{const sh=activeSheet(s);if(!sh)return {};const result=mutateActive(s,(w,active)=>({...w,sheets:w.sheets.map(x=>x.id===active.id?{...x,cells:{...x.cells,[address.toUpperCase()]:{raw:""}}}:x)}));return recalculateAfterMutation(s,result,sh.id,[address.toUpperCase()]);}),
  clearFormat: address => set(s=>mutateActive(s,(w,sh)=>({...w,sheets:w.sheets.map(x=>x.id===sh.id?{...x,cells:{...x.cells,[address]:{...(x.cells[address]??{raw:""}),style:{}}}}:x)}))),
  setSelection: (range,activeCell) => set(s=>({selection:range,activeCell:activeCell??s.activeCell})),
  setClipboard: clipboard => set({clipboard}),
  paste: target => set(s=>{
    const sh=activeSheet(s), clip=s.clipboard, start=parseAddress(target); if(!sh||!clip||!start)return {};
    const source=parseAddress(clip.source??target), rowDelta=start.row-(source?.row??start.row), colDelta=start.col-(source?.col??start.col);
    const shiftFormula=(raw:string)=>raw.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g,(match,colLock:string,letters:string,rowLock:string,digits:string)=>{
      const position=parseAddress(`${letters}${digits}`);if(!position)return match;
      return `${colLock?"$"+letters:columnLabel(position.col+colDelta)}${rowLock?"$"+digits:position.row+rowDelta+1}`;
    });
    const result=mutateActive(s,(w,ss)=>{
      const cells={...ss.cells};
      clip.cells.forEach((row,ri)=>row.forEach((item,ci)=>{const addr=addressOf(start.row+ri,start.col+ci),raw=item.raw.startsWith("=")?shiftFormula(item.raw):item.raw;cells[addr]={raw,style:item.style?{...item.style}:undefined};}));
      return {...w,sheets:w.sheets.map(x=>x.id===ss.id?{...x,cells,rowCount:Math.max(x.rowCount,start.row+clip.cells.length),columnCount:Math.max(x.columnCount,start.col+(clip.cells[0]?.length??1))}:x)};
    });
      const addresses=clip.cells.flatMap((row,ri)=>row.map((_,ci)=>addressOf(start.row+ri,start.col+ci)));
      return recalculateAfterMutation(s,result,sh.id,addresses);
  }),
  updateSheet: (id,changes) => set(s=>{
    const workbook=activeWorkbook(s),target=workbook?.sheets.find(item=>item.id===id);
    if(!workbook||!target)return {};
    const result=mutateActive(s,(w)=>({...w,sheets:w.sheets.map(x=>x.id===id?{...x,...changes}:x)}));
    const changedAddresses=changes.cells?Array.from(new Set([...Object.keys(target.cells),...Object.keys(changes.cells)])).filter(address=>target.cells[address]?.raw!==changes.cells?.[address]?.raw):[];
    if(changes.tables)changedAddresses.push(...Object.entries(target.cells).filter(([,cell])=>cell.raw.startsWith("=")).map(([address])=>address));
    const recalculated=recalculateAfterMutation(s,result,id,changedAddresses);
    if(changes.pivotTables){
      const workbookId=s.activeWorkbookId,updated=(recalculated.workbooks??s.workbooks).find(item=>item.id===workbookId);
      if(updated){const refreshed=refreshWorkbookPivots(updated);return {...recalculated,workbooks:(recalculated.workbooks??s.workbooks).map(item=>item.id===workbookId?refreshed:item)};}
    }
    return recalculated;
  }),
  createPivotTable:(definition,outputCells,outputRange)=>{
    const state=get(),workbook=activeWorkbook(state);
    if(!workbook||!workbook.sheets.some(item=>item.id===definition.sourceSheetId))return null;
    const newDestination=definition.destinationSheetId==="__new__";
    let destination=newDestination?emptySheet(definition.name.trim()||"PivotTable"):workbook.sheets.find(item=>item.id===definition.destinationSheetId);
    if(!destination)return null;
    if(newDestination){
      const baseName=destination.name;let name=baseName,index=2;
      while(workbook.sheets.some(item=>item.name.toLowerCase()===name.toLowerCase()))name=`${baseName} (${index++})`;
      destination={...destination,name};
    }
    const finalDefinition={...definition,destinationSheetId:destination.id,destinationCell:newDestination?"A1":definition.destinationCell,outputRange};
    const destinationNext={...destination,cells:{...destination.cells,...outputCells},pivotTables:[...(destination.pivotTables??[]),finalDefinition],rowCount:Math.max(destination.rowCount,outputRange.endRow+1),columnCount:Math.max(destination.columnCount,outputRange.endCol+1)};
    set(current=>({
      ...snapshot(current),
      workbooks:current.workbooks.map(item=>item.id===workbook.id?{...item,sheets:item.sheets.some(sheet=>sheet.id===destinationNext.id)?item.sheets.map(sheet=>sheet.id===destinationNext.id?destinationNext:sheet):[...item.sheets,destinationNext],modifiedAt:Date.now()}:item),
      activeSheetId:destinationNext.id,activeCell:newDestination?"A1":definition.destinationCell,selection:null
    }));
    return destinationNext.id;
  }),
  addChart: chart => set(s=>mutateActive(s,w=>({...w,charts:[...w.charts,chart]}))),
  updateChart: (id,changes) => set(s=>mutateActive(s,w=>({...w,charts:w.charts.map(c=>c.id===id?{...c,...changes}:c)}))),
  deleteChart: id => set(s=>mutateActive(s,w=>({...w,charts:w.charts.filter(c=>c.id!==id)}))),
  addElement: element => set(s=>mutateActive(s,w=>({...w,elements:[...(w.elements??[]),element]}))),
  updateElement: (id,changes) => set(s=>mutateActive(s,w=>({...w,elements:(w.elements??[]).map(element=>element.id===id?{...element,...changes}:element)}))),
  deleteElement: id => set(s=>mutateActive(s,w=>({...w,elements:(w.elements??[]).filter(element=>element.id!==id)}))),
  setShowFormulas: v=>set({showFormulas:v}), setZoom:v=>set({zoom:Math.max(50,Math.min(200,v))}),
  setOrientation: v=>set(s=>{const w=activeWorkbook(s);return w?{workbooks:s.workbooks.map(x=>x.id===w.id?{...x,orientation:v,modifiedAt:Date.now()}:x)}:{};}),
  setAutoSaveEnabled:enabled=>set({autoSaveEnabled:enabled}),
  undo: ()=>set(s=>{const prev=s.undoStack.at(-1);if(!prev)return {};return {workbooks:prev,undoStack:s.undoStack.slice(0,-1),redoStack:[...s.redoStack,clone(s.workbooks)]};}),
  redo: ()=>set(s=>{const next=s.redoStack.at(-1);if(!next)return {};return {workbooks:next,redoStack:s.redoStack.slice(0,-1),undoStack:[...s.undoStack,clone(s.workbooks)]};}),
}),{name:"gridnova-excel-storage",version:1,partialize:s=>({workbooks:s.autoSaveEnabled?s.workbooks:[],zoom:s.zoom,autoSaveEnabled:s.autoSaveEnabled})}));
