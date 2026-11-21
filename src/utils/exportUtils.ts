import * as XLSX from "xlsx";
import Papa from "papaparse";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import type { WorkbookData, SheetData } from "../types";
import { addressOf, parseAddress } from "../types";
import { evaluate } from "./formulaEngine";

export function sheetToRows(sheet:SheetData){
  const rows:string[][]=[];
  for(let r=0;r<sheet.rowCount;r++){
    const row:string[]=[];
    for(let c=0;c<sheet.columnCount;c++){
      const raw=sheet.cells[addressOf(r,c)]?.raw??"";
      const value=raw.startsWith("=")?evaluate(raw,sheet.cells):raw;
      row.push(String(value??""));
    }
    rows.push(row);
  }
  return rows;
}
export function exportXlsx(workbook:WorkbookData){
  const book=XLSX.utils.book_new();
  workbook.sheets.forEach(sheet=>{
    const ws=XLSX.utils.aoa_to_sheet(sheetToRows(sheet));
    Object.entries(sheet.cells).forEach(([address,cell])=>{
      const target=ws[address] as XLSX.CellObject|undefined;
      if(target&&cell.style){
        target.s={font:{bold:cell.style.bold,italic:cell.style.italic,underline:cell.style.underline,color:{rgb:(cell.style.color??"#202124").replace("#","")}},fill:{fgColor:{rgb:(cell.style.backgroundColor??"#ffffff").replace("#","")}},alignment:{horizontal:cell.style.align}};
      }
    });
    XLSX.utils.book_append_sheet(book,ws,sheet.name.slice(0,31)||"Sheet");
  });
  XLSX.writeFile(book,`${workbook.title.replace(/[\\/:*?"<>|]/g,"-")}.xlsx`);
}
export function exportCsv(sheet:SheetData){
  const csv=Papa.unparse(sheetToRows(sheet));
  const blob=new Blob([csv],{type:"text/csv;charset=utf-8"});
  const url=URL.createObjectURL(blob);const a=document.createElement("a");
  a.href=url;a.download="gridnova-sheet.csv";a.click();URL.revokeObjectURL(url);
}
export function exportPdf(workbook:WorkbookData){
  const doc=new jsPDF({orientation:workbook.orientation,unit:"pt",format:"a4"});
  workbook.sheets.forEach((sheet,index)=>{
    if(index>0)doc.addPage();
    doc.setFontSize(15);doc.text(`${workbook.title} — ${sheet.name}`,40,36);
    const rows=sheetToRows(sheet);
    const head=rows.length?[rows[0]]:[[]],body=rows.slice(1);
    autoTable(doc,{startY:52,head,body,styles:{fontSize:7,cellPadding:4,overflow:"linebreak"},headStyles:{fillColor:[16,124,65],textColor:255},alternateRowStyles:{fillColor:[245,249,246]},margin:{left:40,right:40},
      didParseCell:(data)=>{
        const rowIndex=data.section==="head"?0:data.row.index+1;
        const cell=sheet.cells[addressOf(rowIndex,data.column.index)];
        const style=cell?.style;
        if(!style)return;
        if(style.color) data.cell.styles.textColor=style.color;
        if(style.backgroundColor) data.cell.styles.fillColor=style.backgroundColor;
        if(style.bold||style.italic) data.cell.styles.fontStyle=style.bold&&style.italic?"bolditalic":style.bold?"bold":"italic";
        if(style.align) data.cell.styles.halign=style.align;
      }});
  });
  doc.save(`${workbook.title.replace(/[\\/:*?"<>|]/g,"-")}.pdf`);
}
export function importCsvText(text:string):string[][]{
  return Papa.parse<string[]>(text,{skipEmptyLines:false}).data;
}
export async function importExcelFile(file:File):Promise<string[][]>{
  const buffer=await file.arrayBuffer();const book=XLSX.read(buffer,{type:"array"});
  const first=book.Sheets[book.SheetNames[0]];
  return XLSX.utils.sheet_to_json<string[]>(first,{header:1,raw:false,defval:""}) as string[][];
}
export function rowsToCells(rows:string[][]){
  const cells:SheetData["cells"]={};
  rows.forEach((row,r)=>row.forEach((value,c)=>{cells[addressOf(r,c)]={raw:String(value??"")}}));
  return cells;
}
export function getRangeAddresses(range:string):string[]{
  const [start,end]=range.split(":");const a=parseAddress(start??"A1"),b=parseAddress(end??start??"A1");
  if(!a||!b)return [];
  const out:string[]=[];
  for(let r=Math.min(a.row,b.row);r<=Math.max(a.row,b.row);r++)
    for(let c=Math.min(a.col,b.col);c<=Math.max(a.col,b.col);c++)out.push(addressOf(r,c));
  return out;
}