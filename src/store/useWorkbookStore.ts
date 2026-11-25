import { create } from "zustand";
import { persist } from "zustand/middleware";
import type { CellData, CellMap, CellRange, CellStyle, ClipboardData, FloatingChartData, SheetData, TemplateId, WorkbookData } from "../types";
import { addressOf, columnLabel, emptySheet, emptyWorkbook, parseAddress } from "../types";

interface Store {
  workbooks: WorkbookData[]; activeWorkbookId: string | null; activeSheetId: string | null;
  activeCell: string; selection: CellRange | null; clipboard: ClipboardData | null;
  showFormulas: boolean; zoom: number; undoStack: WorkbookData[][]; redoStack: WorkbookData[][];
  createWorkbook: (template?: TemplateId) => string; deleteWorkbook: (id: string) => void;
  duplicateWorkbook: (id: string) => string | null; renameWorkbook: (id: string, title: string) => void;
  setActiveWorkbook: (id: string | null) => void; setActiveSheet: (id: string) => void;
  addSheet: () => void; renameSheet: (id: string, name: string) => void; deleteSheet: (id: string) => void;
  setCell: (address: string, raw: string) => void; setStyle: (style: Partial<CellStyle>, addresses?: string[]) => void;
  clearCell: (address: string) => void; clearFormat: (address: string) => void;
  setSelection: (range: CellRange | null, activeCell?: string) => void;
  setClipboard: (data: ClipboardData | null) => void; paste: (target: string) => void;
  updateSheet: (id: string, changes: Partial<SheetData>) => void;
  addChart: (chart: FloatingChartData) => void; updateChart: (id: string, changes: Partial<FloatingChartData>) => void; deleteChart: (id: string) => void;
  setShowFormulas: (v: boolean) => void; setZoom: (v: number) => void; setOrientation: (v: "portrait" | "landscape") => void;
  undo: () => void; redo: () => void;
}
const clone = <T,>(v: T): T => structuredClone(v);
const makeTemplate = (kind: TemplateId): WorkbookData => {
  const titles: Record<TemplateId, string> = { blank: "Untitled Workbook", budget: "Monthly Budget", sales: "Sales Tracker", project: "Project Schedule" };
  const wb = emptyWorkbook(titles[kind]); const s = wb.sheets[0];
  const put = (address: string, raw: string, style?: CellStyle) => { s.cells[address] = { raw, style }; };
  const header: CellStyle = { bold: true, color: "#ffffff", backgroundColor: "#128451", fontSize: 11, align: "left" };
  if (kind === "budget") {
    [["Category","Planned","Actual","Difference"],["Housing","18000","17500","=B2-C2"],["Food","8000","7600","=B3-C3"],["Travel","4000","2500","=B4-C4"],["Utilities","3000","2900","=B5-C5"],["Total","=SUM(B2:B5)","=SUM(C2:C5)","=SUM(D2:D5)"]].forEach((r,ri)=>r.forEach((v,ci)=>put(addressOf(ri,ci),v,ri===0?header:ri===5?{bold:true,backgroundColor:"#e8f5ee"}:ci>0?{numberFormat:"Currency ₹"}:undefined)));
  } else if (kind === "sales") {
    [["Product","Units","Unit Price","Revenue","Share"],["Nova Desk","12","4500","=B2*C2","=D2/SUM(D2:D5)"],["Orbit Lamp","18","1800","=B3*C3","=D3/SUM(D2:D5)"],["Cloud Chair","7","6200","=B4*C4","=D4/SUM(D2:D5)"],["Pulse Stand","21","950","=B5*C5","=D5/SUM(D2:D5)"]].forEach((r,ri)=>r.forEach((v,ci)=>put(addressOf(ri,ci),v,ri===0?header:ci===4?{numberFormat:"Percentage %"}:ci===2||ci===3?{numberFormat:"Currency ₹"}:undefined)));
  } else if (kind === "project") {
    [["Task","Owner","Start Date","Due Date","Status","Done"],["Discovery","Maya","2026-11-03","2026-11-07","Complete","TRUE"],["Wireframes","Arjun","2026-11-08","2026-11-14","In Progress","FALSE"],["Development","Sam","2026-11-15","2026-11-28","Not Started","FALSE"],["QA","Maya","2026-11-29","2026-12-04","Not Started","FALSE"]].forEach((r,ri)=>r.forEach((v,ci)=>put(addressOf(ri,ci),v,ri===0?header:undefined)));
    for(let row=2;row<=5;row++){
      s.cells[`C${row}`]={...s.cells[`C${row}`],datePicker:true};
      s.cells[`D${row}`]={...s.cells[`D${row}`],datePicker:true};
      s.cells[`F${row}`]={...s.cells[`F${row}`],checkbox:true};
    }
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
export const useWorkbookStore = create<Store>()(persist((set, get) => ({
  workbooks: [], activeWorkbookId: null, activeSheetId: null, activeCell: "A1", selection: null,
  clipboard: null, showFormulas: false, zoom: 100, undoStack: [], redoStack: [],
  createWorkbook: (template = "blank") => {
    const wb = makeTemplate(template);
    set(s => ({ workbooks: [wb, ...s.workbooks], activeWorkbookId: wb.id, activeSheetId: wb.sheets[0].id, activeCell: "A1", selection: null }));
    return wb.id;
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
  setCell: (address,raw) => set(s => mutateActive(s,(w,sh)=>({...w,sheets:w.sheets.map(x=>x.id===sh.id?{...x,cells:{...x.cells,[address.toUpperCase()]:{...x.cells[address.toUpperCase()],raw}}}:x)}))),
  setStyle: (style,addresses) => set(s => {
    const sh=activeSheet(s); if(!sh) return {};
    const targets=addresses?.length?addresses:[s.activeCell];
    return mutateActive(s,(w,ss)=>({...w,sheets:w.sheets.map(x=>x.id===ss.id?{...x,cells:targets.reduce((acc,a)=>{const key=a.toUpperCase();acc[key]={...(acc[key]??{raw:""}),style:{...(acc[key]?.style??{}),...style}};return acc;},{...x.cells})}:x)}));
  }),
  clearCell: address => set(s=>mutateActive(s,(w,sh)=>({...w,sheets:w.sheets.map(x=>x.id===sh.id?{...x,cells:{...x.cells,[address]:{raw:""}}}:x)}))),
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
    return mutateActive(s,(w,ss)=>{
      const cells={...ss.cells};
      clip.cells.forEach((row,ri)=>row.forEach((item,ci)=>{const addr=addressOf(start.row+ri,start.col+ci),raw=item.raw.startsWith("=")?shiftFormula(item.raw):item.raw;cells[addr]={raw,style:item.style?{...item.style}:undefined};}));
      return {...w,sheets:w.sheets.map(x=>x.id===ss.id?{...x,cells,rowCount:Math.max(x.rowCount,start.row+clip.cells.length),columnCount:Math.max(x.columnCount,start.col+(clip.cells[0]?.length??1))}:x)};
    });
  }),
  updateSheet: (id,changes) => set(s=>mutateActive(s,(w,sh)=>({...w,sheets:w.sheets.map(x=>x.id===id?{...x,...changes}:x)}))),
  addChart: chart => set(s=>mutateActive(s,w=>({...w,charts:[...w.charts,chart]}))),
  updateChart: (id,changes) => set(s=>mutateActive(s,w=>({...w,charts:w.charts.map(c=>c.id===id?{...c,...changes}:c)}))),
  deleteChart: id => set(s=>mutateActive(s,w=>({...w,charts:w.charts.filter(c=>c.id!==id)}))),
  setShowFormulas: v=>set({showFormulas:v}), setZoom:v=>set({zoom:Math.max(50,Math.min(200,v))}),
  setOrientation: v=>set(s=>{const w=activeWorkbook(s);return w?{workbooks:s.workbooks.map(x=>x.id===w.id?{...x,orientation:v,modifiedAt:Date.now()}:x)}:{};}),
  undo: ()=>set(s=>{const prev=s.undoStack.at(-1);if(!prev)return {};return {workbooks:prev,undoStack:s.undoStack.slice(0,-1),redoStack:[...s.redoStack,clone(s.workbooks)]};}),
  redo: ()=>set(s=>{const next=s.redoStack.at(-1);if(!next)return {};return {workbooks:next,redoStack:s.redoStack.slice(0,-1),undoStack:[...s.undoStack,clone(s.workbooks)]};}),
}),{name:"gridnova-excel-storage",version:1,partialize:s=>({workbooks:s.workbooks,zoom:s.zoom})}));
