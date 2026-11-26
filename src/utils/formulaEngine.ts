import type { CellMap, CellStyle, ExcelTableData, SheetData } from "../types";
import { addressOf, parseAddress } from "../types";
import { getTableAt, getTableColumnIndex, tableFilterPasses, tableReferenceRange } from "./tableUtils";

type Scalar = string | number | boolean;
interface ArrayValue { __gridArray:true; values:Scalar[][]; }
type Value = Scalar | ArrayValue;
const refPattern = /^\$?[A-Z]{1,3}\$?[1-9]\d*$/i;
const columnPattern = /^\$?[A-Z]{1,3}$/i;
const rowPattern = /^\$?[1-9]\d*$/;
const isArrayValue=(value:Value):value is ArrayValue=>typeof value==="object"&&value!==null&&"__gridArray" in value;
const scalar=(value:Value|undefined):Scalar=>isArrayValue(value)?(value.values[0]?.[0]??""):value??"";
const flatValues=(values:Value[])=>values.flatMap(value=>isArrayValue(value)?value.values.flat():[value]);

const tableNamePattern = /^[A-Za-z_][A-Za-z0-9_.]*/;

function resolveStructuredReference(name: string | null, inner: string, cells: CellMap, tables: ExcelTableData[], currentAddress?: string): string {
  const current = currentAddress ? parseAddress(currentAddress) : null;
  const table = name
    ? tables.find(item => item.name.toLowerCase() === name.toLowerCase())
    : current ? getTableAt(tables, current.row, current.col) : undefined;
  if (!table) throw new Error("#NAME?");
  const terms = [...inner.matchAll(/\[([^\]]+)\]/g)].map(match => match[1].trim());
  const selector = terms.find(term => term.startsWith("#")) ?? (inner.trim().startsWith("#") ? inner.trim() : undefined);
  const thisRow = inner.trim().startsWith("@") || selector?.toLowerCase() === "#this row";
  const columnName = terms.filter(term => !term.startsWith("#")).at(-1) ?? (terms.length || inner.trim().startsWith("#") ? "" : inner.replace(/^@/, "").trim());
  if (thisRow) {
    const firstDataRow = table.range.startRow + (table.hasHeaders ? 1 : 0);
    if (!current || current.row < firstDataRow || current.row > table.range.endRow) throw new Error("#VALUE!");
    if (!columnName) return addressOf(current.row, current.col);
    const column = getTableColumnIndex(table, columnName, cells);
    if (column < 0) throw new Error("#REF!");
    return addressOf(current.row, column);
  }
  if (!columnName) {
    const range = tableReferenceRange(table, selector);
    return `${addressOf(range.startRow, range.startCol)}:${addressOf(range.endRow, range.endCol)}`;
  }
  const column = getTableColumnIndex(table, columnName, cells);
  if (column < 0) throw new Error("#REF!");
  const range = tableReferenceRange(table, selector ?? "#Data");
  return `${addressOf(range.startRow, column)}:${addressOf(range.endRow, column)}`;
}

function expandStructuredReferences(source: string, cells: CellMap, tables: ExcelTableData[], currentAddress?: string): string {
  let output = "", i = 0;
  while (i < source.length) {
    // Current-row shorthand: [@Amount] or [@[Amount]]
    if (source[i] === "[" && source[i + 1] === "@") {
      let depth = 0, end = i;
      for (; end < source.length; end++) {
        if (source[end] === "[") depth++;
        else if (source[end] === "]" && --depth === 0) break;
      }
      if (end < source.length) {
        output += resolveStructuredReference(null, source.slice(i + 1, end), cells, tables, currentAddress);
        i = end + 1; continue;
      }
    }
    const match = source.slice(i).match(tableNamePattern);
    if (match) {
      const name = match[0], bracketIndex = i + name.length;
      if (source[bracketIndex] === "[") {
        let depth = 0, end = bracketIndex;
        for (; end < source.length; end++) {
          if (source[end] === "[") depth++;
          else if (source[end] === "]") { depth--; if (depth === 0) break; }
        }
        if (end < source.length) {
          output += resolveStructuredReference(name, source.slice(bracketIndex + 1, end), cells, tables, currentAddress);
          i = end + 1; continue;
        }
      }
      output += name; i += name.length; continue;
    }
    output += source[i]; i++;
  }
  return output;
}

function expandNamedRanges(source:string,namedRanges:Record<string,string>={}):string {
  const lookup=new Map(Object.entries(namedRanges).map(([name,ref])=>[name.toUpperCase(),ref]));
  if(!lookup.size)return source;
  let output="",i=0;
  while(i<source.length){
    if(source[i]==='"'){
      const start=i++;
      while(i<source.length){
        if(source[i]==='"'&&source[i+1]==='"'){i+=2;continue;}
        if(source[i++]==='"')break;
      }
      output+=source.slice(start,i);
      continue;
    }
    const match=source.slice(i).match(/^[A-Za-z_][A-Za-z0-9_.]*/);
    if(match){
      const replacement=lookup.get(match[0].toUpperCase());
      if(replacement){
        const localRef=replacement.trim().replace(/^=/,"").split("!").pop()?.replace(/\$/g,"")??"";
        output+=/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/i.test(localRef)?localRef.toUpperCase():match[0];
      }else output+=match[0];
      i+=match[0].length;
      continue;
    }
    output+=source[i++];
  }
  return output;
}

function rangeValues(start: string, end: string, cells: CellMap, seen: Set<string>, tables: ExcelTableData[], currentAddress?: string, ignoreFiltered = false, namedRanges:Record<string,string>={}, sheets:SheetData[]=[], currentSheetName?:string): ArrayValue {
  const clean=(value:string)=>value.replace(/\$/g,"").toUpperCase();
  const parseAxis=(value:string,other:string)=>{
    const p=parseAddress(clean(value));
    if(p)return p;
    if(columnPattern.test(value)){const col=clean(value).split("").reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)-1;return {row:0,col};}
    if(rowPattern.test(value))return {row:Number(clean(value))-1,col:0};
    return null;
  };
  const a=parseAxis(start,end),b=parseAxis(end,start);
  if(!a||!b)throw new Error("#REF!");
  const maxUsed=Object.keys(cells).map(parseAddress).filter((p):p is {row:number;col:number}=>!!p);
  const maxRow=Math.max(99,...maxUsed.map(p=>p.row)),maxCol=Math.max(25,...maxUsed.map(p=>p.col));
  const colOnly=columnPattern.test(start)&&columnPattern.test(end);
  const rowOnly=rowPattern.test(start)&&rowPattern.test(end);
  const startRow=colOnly?0:Math.min(a.row,b.row),endRow=colOnly?maxRow:Math.max(a.row,b.row);
  const startCol=rowOnly?0:Math.min(a.col,b.col),endCol=rowOnly?maxCol:Math.max(a.col,b.col);
  const matrix:Scalar[][]=[];
  for(let r=startRow;r<=endRow;r++){
    const row:Scalar[]=[];
    for(let col=startCol;col<=endCol;col++){
      const table=ignoreFiltered?getTableAt(tables,r,col):undefined;
      row.push(table&&!tableFilterPasses(table,r,cells)?"":scalar(resolve(addressOf(r,col),cells,seen,tables,namedRanges,sheets,currentSheetName)));
    }
    matrix.push(row);
  }
  return {__gridArray:true,values:matrix};
}

function resolve(address: string, cells: CellMap, seen: Set<string>, tables: ExcelTableData[], namedRanges:Record<string,string>={},sheets:SheetData[]=[],currentSheetName?:string): Value {
  const key = address.toUpperCase(), cell = cells[key];
  if (!cell) return "";
  if (seen.has(key)) throw new Error("#CIRCULAR!");
  if (cell.raw.startsWith("=")) {
    const next = new Set(seen); next.add(key);
    return evaluate(cell.raw, cells, next, tables, key, namedRanges,sheets,currentSheetName);
  }
  const n = Number(cell.raw);
  return cell.raw.trim() !== "" && Number.isFinite(n) ? n : cell.raw;
}

function tokenize(source: string): string[] {
  const result: string[] = [];
  const re = /\s*(>=|<=|<>|[()+*/^%=<>:&!,\-]|"(?:[^"]|"")*"|'(?:[^']|'')+'|(?:\d+\.?\d*|\.\d+)|\$?[A-Za-z_][A-Za-z0-9_.]*|\$?\d+)/gy;
  let pos = 0;
  while (pos < source.length) {
    re.lastIndex = pos;
    const match = re.exec(source);
    if (!match) throw new Error("#PARSE!");
    result.push(match[1]); pos = re.lastIndex;
  }
  return result;
}

const isError = (value: Value | undefined) => typeof scalar(value) === "string" && String(scalar(value)).startsWith("#");
const number = (value: Value | undefined) => value === undefined || scalar(value) === "" ? 0 : Number(scalar(value));
const numeric = (values: Value[]) => flatValues(values).filter((value): value is number => typeof value === "number" && Number.isFinite(value));
const text = (value: Value | undefined) => value === undefined ? "" : String(scalar(value));
const EXCEL_EPOCH_UTC=Date.UTC(1899,11,30);
const toDate=(value:Value|undefined):Date|null=>{
 const v=scalar(value);
 if(typeof v==="number"&&Number.isFinite(v))return new Date(EXCEL_EPOCH_UTC+Math.round(v*86400000));
 if(typeof v==="string"){
  if(!v.trim())return null;
  const parsed=new Date(v);
  return Number.isNaN(parsed.getTime())?null:parsed;
 }
 return null;
};
const toSerial=(value:Value|undefined):number=>{
 const v=scalar(value);
 if(typeof v==="number")return v;
 const date=toDate(v);return date?((Date.UTC(date.getUTCFullYear(),date.getUTCMonth(),date.getUTCDate())-EXCEL_EPOCH_UTC)/86400000):NaN;
};
const serialFromDate=(date:Date)=>((Date.UTC(date.getUTCFullYear(),date.getUTCMonth(),date.getUTCDate())-EXCEL_EPOCH_UTC)/86400000);
const dateFromSerial=(serial:number)=>new Date(EXCEL_EPOCH_UTC+Math.round(serial*86400000));
const dateOnly=(date:Date)=>new Date(Date.UTC(date.getUTCFullYear(),date.getUTCMonth(),date.getUTCDate()));
const excelWeekday=(date:Date,returnType=1)=>{
 const day=date.getUTCDay();
 if(returnType===2)return day===0?7:day;
 if(returnType===3)return day===0?6:day-1;
 if(returnType>=11&&returnType<=17){const first=(returnType-11+1)%7;return ((day-first+7)%7)+1;}
 return day+1;
};
const addMonths=(date:Date,months:number)=>{
 const day=date.getUTCDate(),out=new Date(Date.UTC(date.getUTCFullYear(),date.getUTCMonth()+months,1));
 const last=new Date(Date.UTC(out.getUTCFullYear(),out.getUTCMonth()+1,0)).getUTCDate();
 out.setUTCDate(Math.min(day,last));return out;
};
const average=(values:number[])=>values.length?values.reduce((s,v)=>s+v,0)/values.length:NaN;
const variance=(values:number[],sample:boolean)=>{
 if(values.length<(sample?2:1))return NaN;
 const mean=average(values),sum=values.reduce((acc,v)=>acc+(v-mean)**2,0);
 return sum/(values.length-(sample?1:0));
};
const percentile=(values:number[],k:number)=>{
 const sorted=[...values].sort((a,b)=>a-b);if(!sorted.length||k<0||k>1)return NaN;
 const pos=(sorted.length-1)*k,lo=Math.floor(pos),hi=Math.ceil(pos);
 return sorted[lo]+(sorted[hi]-sorted[lo])*(pos-lo);
};
const criteriaMatches=(value:Value,criteria:Value):boolean=>{
 const condition=String(scalar(criteria)),op=condition.match(/^(>=|<=|<>|>|<|=)/)?.[0]??"=";
 const expected=condition.slice(op==="="&&!/^(>=|<=|<>|>|<|=)/.test(condition)?0:op.length);
 const actual=scalar(value);
 if(/[*?]/.test(expected)){
  const pattern=expected.split("").map(char=>"\\\\^$.*+?()[]{}|".includes(char)?"\\\\"+char:char).join("").replace(/\\\\~/g,"~").replace(/\\*/g,".*").replace(/\\?/g,".");
  const match=new RegExp("^"+pattern+"$","i").test(String(actual));
  return op==="<>"?!match:match;
 }
 const leftNum=Number(actual),rightNum=Number(expected),both=String(actual).trim()!==""&&expected.trim()!==""&&Number.isFinite(leftNum)&&Number.isFinite(rightNum);
 const left: string|number=both?leftNum:String(actual).toLowerCase(),right: string|number=both?rightNum:expected.toLowerCase();
 switch(op){case ">=":return left>=right;case "<=":return left<=right;case ">":return left>right;case "<":return left<right;case "<>":return left!==right;default:return left===right;}
};


class Parser {
  private i=0;
  private variables=new Map<string,Scalar>();
  constructor(private readonly tokens:string[],private readonly cells:CellMap,private readonly seen:Set<string>,private readonly tables:ExcelTableData[],private readonly namedRanges:Record<string,string>={},private readonly currentAddress?:string,private readonly sheets:SheetData[]=[],private readonly currentSheetName?:string){}
  private peek(){return this.tokens[this.i];}
  private take(){return this.tokens[this.i++];}
  private match(v:string){if(this.peek()===v){this.i++;return true;}return false;}
  parse():Value{const value=this.compare();if(this.i!==this.tokens.length)throw new Error("#PARSE!");return value;}
  private compare():Value{
    let left=this.concat();
    while([">","<",">=","<=","=","<>"].includes(this.peek()??"")){
      const op=this.take()!,right=this.concat(),a=scalar(left),b=scalar(right);
      const an=Number(a),bn=Number(b),both=String(a).trim()!==""&&String(b).trim()!==""&&Number.isFinite(an)&&Number.isFinite(bn);
      const x:string|number=both?an:String(a),y:string|number=both?bn:String(b);
      switch(op){case ">":left=x>y;break;case "<":left=x<y;break;case ">=":left=x>=y;break;case "<=":left=x<=y;break;case "=":left=String(a).toLowerCase()===String(b).toLowerCase()||x===y;break;default:left=!(String(a).toLowerCase()===String(b).toLowerCase()||x===y);}
    }
    return left;
  }
  private concat():Value{
    let value=this.add();
    while(this.match("&"))value=String(scalar(value))+String(scalar(this.add()));
    return value;
  }
  private add():Value{
    let value=this.multiply();
    while(this.peek()==="+"||this.peek()==="-"){
      const op=this.take(),right=this.multiply();
      value=op==="+"?number(value)+number(right):number(value)-number(right);
    }
    return value;
  }
  private multiply():Value{
    let value=this.unary();
    while(this.peek()==="*"||this.peek()==="/"){
      const op=this.take(),right=this.unary();
      if(op==="/"&&number(right)===0)return "#DIV/0!";
      value=op==="*"?number(value)*number(right):number(value)/number(right);
    }
    return value;
  }
  private unary():Value{
    if(this.match("-"))return -number(this.unary());
    if(this.match("+"))return number(this.unary());
    return this.power();
  }
  private power():Value{
    const left=this.postfix();
    if(this.match("^"))return Math.pow(number(left),number(this.unary()));
    return left;
  }
  private postfix():Value{
    let value=this.primary();
    while(this.match("%"))value=number(value)/100;
    return value;
  }
  private readRangeArgument(forFunction?:string):Value[]|null{
    const save=this.i,first=this.peek();
    if(!first)return null;
    let sheetName:string|undefined;
    if(this.tokens[this.i+1]==="!"){
      sheetName=first.replace(/^'|'$/g,"").replace(/''/g,"'");
      this.i+=2;
    }
    const start=this.peek(),colon=this.tokens[this.i+1];
    if(start&&colon===":"&&validReference(start)){
      this.i+=2;const finish=this.take();
      if(!finish||!validReference(finish)){this.i=save;throw new Error("#REF!");}
      const targetSheet=sheetName?this.sheets.find(item=>item.name.toLowerCase()===sheetName!.toLowerCase()):undefined;
      if(sheetName&&!targetSheet){this.i=save;throw new Error("#REF!");}
      return [rangeValues(start,finish,targetSheet?.cells??this.cells,this.seen,targetSheet?.tables??this.tables,undefined,forFunction==="SUBTOTAL",this.namedRanges,this.sheets,targetSheet?.name??this.currentSheetName)];
    }
    if(sheetName&&start&&refPattern.test(start)){
      this.i++;return [this.resolveReference(start,sheetName)];
    }
    this.i=save;return null;
  }
  private resolveReference(ref:string,sheetName?:string):Value{
    const targetSheet=sheetName?this.sheets.find(item=>item.name.toLowerCase()===sheetName.toLowerCase()):undefined;
    if(sheetName&&!targetSheet)return "#REF!";
    const targetCells=targetSheet?.cells??this.cells,targetTables=targetSheet?.tables??this.tables;
    const normalized=ref.replace(/\$/g,"").toUpperCase(),source=targetCells[normalized];
    if(!source)return "";
    if(this.seen.has(normalized))return "#CIRCULAR!";
    if(source.raw.startsWith("=")){
      const seen=new Set(this.seen);seen.add(normalized);
      return evaluate(source.raw,targetCells,seen,targetTables,normalized,this.namedRanges,this.sheets,targetSheet?.name??this.currentSheetName);
    }
    const numericValue=Number(source.raw);
    return source.raw.trim()!==""&&Number.isFinite(numericValue)?numericValue:source.raw;
  }
  private argument(forFunction?:string):Value[]{
    const range=this.readRangeArgument(forFunction);
    return range??[this.compare()];
  }
  private parseLet():Value{
    const original=this.variables;this.variables=new Map(original);
    try{
      while(true){
        const name=this.take();
        if(!name||!/^[_A-Za-z][A-Za-z0-9_.]*$/.test(name)||name.includes("."))throw new Error("#VALUE!");
        if(!this.match(","))throw new Error("#PARSE!");
        const value=this.compare();
        if(this.match(",")){this.variables.set(name.toUpperCase(),scalar(value));continue;}
        if(!this.match(")"))throw new Error("#PARSE!");
        return value;
      }
    }finally{this.variables=original;}
  }
  private primary():Value{
    const token=this.take();if(!token)throw new Error("#PARSE!");
    if(token==="("){const value=this.compare();if(!this.match(")"))throw new Error("#PARSE!");return value;}
    if(/^"(?:[^"]|"")*"$/.test(token))return token.slice(1,-1).replace(/""/g,'"');
    if((/^'(?:[^']|'')+'$/.test(token)||/^[A-Za-z_][A-Za-z0-9_ ]*$/.test(token))&&this.match("!")){
      const sheetName=token.startsWith("'")?token.slice(1,-1).replace(/''/g,"'"):token,ref=this.take();
      if(!ref||!validReference(ref))throw new Error("#REF!");
      if(this.match(":")){const end=this.take();if(!end||!validReference(end))throw new Error("#REF!");const target=this.sheets.find(item=>item.name.toLowerCase()===sheetName.toLowerCase());if(!target)return "#REF!";return rangeValues(ref,end,target.cells,this.seen,target.tables??[],undefined,false,this.namedRanges,this.sheets,target.name);}
      return this.resolveReference(ref,sheetName);
    }
    if(/^(?:\d+\.?\d*|\.\d+)$/.test(token))return Number(token);
    if(token.toUpperCase()==="TRUE")return true;
    if(token.toUpperCase()==="FALSE")return false;
    const upper=token.toUpperCase();
    if(this.match("(")){
      if(upper==="LET")return this.parseLet();
      const args:Value[][]=[];
      if(!this.match(")")){do{args.push(this.argument(upper));}while(this.match(","));if(!this.match(")"))throw new Error("#PARSE!");}
      return this.call(upper,args);
    }
    if(refPattern.test(token))return this.resolveReference(token);
    if(this.variables.has(upper))return this.variables.get(upper)!;
    throw new Error("#NAME?");
  }
  private call(fn:string,args:Value[][]):Value{
    const flat=flatValues(args.flat()),nums=numeric(flat),a=args.map(item=>scalar(item[0]));
    const list=(index:number)=>flatValues(args[index]??[]);
    const matrix=(index:number):Scalar[][]=>{
      const item=args[index]?.[0];
      if(item&&isArrayValue(item))return item.values;
      return list(index).map(value=>[value]);
    };
    const sum=(values:number[])=>values.reduce((total,value)=>total+value,0);
    const pmt=(rate:number,nper:number,pv:number,fv=0,type=0)=>{
      if(nper===0)return NaN;if(rate===0)return -(pv+fv)/nper;
      const factor=Math.pow(1+rate,nper);
      return -(pv*factor+fv)*rate/((1+rate*type)*(factor-1));
    };
    const ipmt=(rate:number,period:number,nper:number,pv:number,fv=0,type=0)=>{
      if(period<1||period>nper)return NaN;
      const payment=pmt(rate,nper,pv,fv,type);
      if(rate===0)return 0;
      const balance=type===1?pv*(1+rate)**(period-1)+payment*((1+rate)**(period-1)-1)*(1+rate)/rate:pv*(1+rate)**(period-1)+payment*((1+rate)**(period-1)-1)/rate;
      return -balance*rate;
    };
    const irr=(values:number[],guess=0.1)=>{
      let rate=guess;
      for(let i=0;i<100;i++){
        let value=0,derivative=0;
        values.forEach((cash,index)=>{value+=cash/Math.pow(1+rate,index);if(index)derivative-=index*cash/Math.pow(1+rate,index+1);});
        if(Math.abs(value)<1e-8)return rate;
        if(!Number.isFinite(derivative)||derivative===0)break;
        const next=rate-value/derivative;if(!Number.isFinite(next)||next<=-0.999999)break;
        if(Math.abs(next-rate)<1e-10)return next;rate=next;
      }
      return NaN;
    };
    const xirr=(values:number[],dates:Value[],guess=0.1)=>{
      const serials=dates.map(toSerial);if(!serials.length||serials.some(v=>!Number.isFinite(v)))return NaN;
      const origin=serials[0];let rate=guess;
      for(let i=0;i<120;i++){
        let value=0,derivative=0;
        values.forEach((cash,index)=>{const years=(serials[index]-origin)/365,base=Math.pow(1+rate,years);value+=cash/base;derivative-=years*cash/(base*(1+rate));});
        if(Math.abs(value)<1e-8)return rate;if(!Number.isFinite(derivative)||derivative===0)break;
        const next=rate-value/derivative;if(!Number.isFinite(next)||next<=-0.999999)break;
        if(Math.abs(next-rate)<1e-10)return next;rate=next;
      }return NaN;
    };
    const holidaySerials=(index:number)=>new Set(list(index).map(toSerial).filter(Number.isFinite));
    const isWeekend=(date:Date,weekend:Scalar)=>{
      const code=typeof weekend==="number"?weekend:String(weekend||1);
      if(typeof code==="string"&&/^[01]{7}$/.test(code))return code[date.getUTCDay()]==="1";
      const numericCode=Number(code),weekendMap:Record<number,number[]>={1:[0,6],2:[0,1],3:[1,2],4:[2,3],5:[3,4],6:[4,5],7:[5,6],11:[0],12:[1],13:[2],14:[3],15:[4],16:[5],17:[6]};
      return (weekendMap[numericCode]??[0,6]).includes(date.getUTCDay());
    };
    const workdays=(start:Date,days:number,weekend:Scalar,holidays:Set<number>)=>{
      const date=dateOnly(start),step=days<0?-1:1;let remaining=Math.abs(Math.trunc(days));
      while(remaining>0){date.setUTCDate(date.getUTCDate()+step);if(!isWeekend(date,weekend)&&!holidays.has(serialFromDate(date)))remaining--;}
      return serialFromDate(date);
    };
    switch(fn){
      case "SUM":return sum(nums);
      case "AVERAGE":return nums.length?average(nums):"#DIV/0!";
      case "AVERAGEA":{const values=flat.map(v=>typeof v==="boolean"?(v?1:0):typeof v==="number"?v:0);return values.length?average(values):"#DIV/0!";}
      case "COUNT":return nums.length;
      case "COUNTA":return flat.filter(value=>value!=="").length;
      case "COUNTBLANK":return flat.filter(value=>value==="").length;
      case "MIN":return nums.length?Math.min(...nums):0;
      case "MAX":return nums.length?Math.max(...nums):0;
      case "MINA":{const values=flat.map(v=>typeof v==="boolean"?(v?1:0):typeof v==="number"?v:0);return values.length?Math.min(...values):0;}
      case "MAXA":{const values=flat.map(v=>typeof v==="boolean"?(v?1:0):typeof v==="number"?v:0);return values.length?Math.max(...values):0;}
      case "PRODUCT":return nums.reduce((total,value)=>total*value,1);
      case "SUMSQ":return sum(nums.map(value=>value*value));
      case "SUMPRODUCT":{const arrays=args.map((_,index)=>list(index)),length=Math.min(...arrays.map(v=>v.length));return Array.from({length},(_,i)=>arrays.reduce((product,values)=>product*number(values[i]),1)).reduce((s,v)=>s+v,0);}
      case "ABS":return Math.abs(number(a[0]));
      case "ROUND":return roundTo(number(a[0]),number(a[1]));
      case "ROUNDUP":return roundDirected(number(a[0]),number(a[1]),"up");
      case "ROUNDDOWN":return roundDirected(number(a[0]),number(a[1]),"down");
      case "MROUND":{const m=number(a[1]);return m===0?"#DIV/0!":Math.sign(number(a[0]))*Math.round(Math.abs(number(a[0])/m))*Math.abs(m);}
      case "CEILING":{const sig=number(a[1]??1);return sig===0?0:Math.ceil(number(a[0])/sig)*sig;}
      case "FLOOR":{const sig=number(a[1]??1);return sig===0?0:Math.floor(number(a[0])/sig)*sig;}
      case "INT":return Math.floor(number(a[0]));
      case "TRUNC":{const factor=10**number(a[1]??0);return Math.trunc(number(a[0])*factor)/factor;}
      case "MOD":if(number(a[1])===0)return "#DIV/0!";return number(a[0])-number(a[1])*Math.floor(number(a[0])/number(a[1]));
      case "QUOTIENT":if(number(a[1])===0)return "#DIV/0!";return Math.trunc(number(a[0])/number(a[1]));
      case "POWER":return Math.pow(number(a[0]),number(a[1]));
      case "SQRT":return number(a[0])<0?"#NUM!":Math.sqrt(number(a[0]));
      case "EXP":return Math.exp(number(a[0]));
      case "LN":return number(a[0])<=0?"#NUM!":Math.log(number(a[0]));
      case "LOG":{const base=number(a[1]??10);return number(a[0])<=0||base<=0||base===1?"#NUM!":Math.log(number(a[0]))/Math.log(base);}
      case "LOG10":return number(a[0])<=0?"#NUM!":Math.log10(number(a[0]));
      case "PI":return Math.PI;
      case "SIN":return Math.sin(number(a[0]));
      case "COS":return Math.cos(number(a[0]));
      case "TAN":return Math.tan(number(a[0]));
      case "ASIN":return Math.asin(number(a[0]));
      case "ACOS":return Math.acos(number(a[0]));
      case "ATAN":return Math.atan(number(a[0]));
      case "DEGREES":return number(a[0])*180/Math.PI;
      case "RADIANS":return number(a[0])*Math.PI/180;
      case "SIGN":return Math.sign(number(a[0]));
      case "RAND":return Math.random();
      case "RANDBETWEEN":return Math.floor(Math.random()*(number(a[1])-number(a[0])+1))+number(a[0]);
      case "FACT":{const n=Math.trunc(number(a[0]));if(n<0)return "#NUM!";let v=1;for(let i=2;i<=Math.min(n,170);i++)v*=i;return v;}
      case "GCD":{const gcd=(x:number,y:number):number=>y?gcd(y,x%y):Math.abs(x);return nums.reduce((acc,v)=>gcd(acc,Math.trunc(v)),0);}
      case "LCM":{const gcd=(x:number,y:number):number=>y?gcd(y,x%y):Math.abs(x);return nums.reduce((acc,v)=>v===0?0:Math.abs(acc*v)/Math.max(1,gcd(acc,Math.trunc(v))),1);}
      case "ROMAN":{let n=Math.trunc(number(a[0]));if(n<0||n>3999)return "#VALUE!";const roman:[number,string][]=[[1000,"M"],[900,"CM"],[500,"D"],[400,"CD"],[100,"C"],[90,"XC"],[50,"L"],[40,"XL"],[10,"X"],[9,"IX"],[5,"V"],[4,"IV"],[1,"I"]];let out="";for(const [v,s] of roman)while(n>=v){out+=s;n-=v;}return out;}
      case "ARABIC":{const s=text(a[0]).toUpperCase(),map:Record<string,number>={I:1,V:5,X:10,L:50,C:100,D:500,M:1000};let total=0;for(let i=0;i<s.length;i++)total+=(map[s[i]]??0)<(map[s[i+1]]??0)?-(map[s[i]]??0):(map[s[i]]??0);return total;}
      case "IF":return a[0]?a[1]??false:a[2]??false;
      case "IFS":for(let i=0;i+1<a.length;i+=2)if(a[i])return a[i+1];return "#N/A";
      case "AND":return flat.every(Boolean);
      case "OR":return flat.some(Boolean);
      case "NOT":return !Boolean(a[0]);
      case "XOR":return flat.filter(Boolean).length%2===1;
      case "TRUE":return true;
      case "FALSE":return false;
      case "IFERROR":return isError(a[0])?a[1]??"":a[0]??"";
      case "IFNA":return String(scalar(a[0]))==="#N/A"?a[1]??"":a[0]??"";
      case "SWITCH":{for(let i=1;i+1<a.length;i+=2)if(a[0]===a[i])return a[i+1];return a.length%2===0?a[a.length-1]??"#N/A":"#N/A";}
      case "ISBLANK":return a[0]==="";
      case "ISERR":return isError(a[0])&&String(a[0])!=="#N/A";
      case "ISERROR":return isError(a[0]);
      case "ISEVEN":return Math.abs(Math.trunc(number(a[0])))%2===0;
      case "ISODD":return Math.abs(Math.trunc(number(a[0]))%2)===1;
      case "ISFORMULA":return false;
      case "ISLOGICAL":return typeof a[0]==="boolean";
      case "ISNA":return String(a[0])==="#N/A";
      case "ISNONTEXT":return typeof a[0]!=="string";
      case "ISNUMBER":return typeof a[0]==="number"&&Number.isFinite(a[0]);
      case "ISREF":return false;
      case "ISTEXT":return typeof a[0]==="string"&&!isError(a[0]);
      case "N":return typeof a[0]==="number"?a[0]:typeof a[0]==="boolean"?(a[0]?1:0):0;
      case "NA":return "#N/A";
      case "TYPE":return typeof a[0]==="number"?1:typeof a[0]==="string"?2:typeof a[0]==="boolean"?4:isArrayValue(a[0])?64:16;
      case "ERROR.TYPE":{const codes:Record<string,number>={"#NULL!":1,"#DIV/0!":2,"#VALUE!":3,"#REF!":4,"#NAME?":5,"#NUM!":6,"#N/A":7,"#GETTING_DATA":8};return codes[String(a[0])]??"#N/A";}
      case "CELL":return a[0]??"";
      case "LEFT":return text(a[0]).slice(0,Math.max(0,number(a[1]??1)));
      case "RIGHT":return text(a[0]).slice(-Math.max(0,number(a[1]??1)));
      case "MID":return text(a[0]).slice(Math.max(0,number(a[1])-1),Math.max(0,number(a[1])-1)+Math.max(0,number(a[2])));
      case "LEN":return text(a[0]).length;
      case "TRIM":return text(a[0]).trim().replace(/\s+/g," ");
      case "CLEAN":return text(a[0]).replace(/[\x00-\x1F]/g,"");
      case "UPPER":return text(a[0]).toUpperCase();
      case "LOWER":return text(a[0]).toLowerCase();
      case "PROPER":return text(a[0]).toLowerCase().replace(/\b([a-z])/g,letter=>letter.toUpperCase());
      case "CONCAT":case "CONCATENATE":return flat.map(value=>String(value)).join("");
      case "TEXTJOIN":{const delimiter=text(a[0]),ignore=!!a[1],parts=flatValues(args.slice(2).flat()).map(String).filter(value=>!ignore||value!=="");return parts.join(delimiter);}
      case "FIND":case "SEARCH":{const needle=text(a[0]),haystack=text(a[1]),start=Math.max(1,number(a[2]??1))-1;const index=fn==="FIND"?haystack.indexOf(needle,start):haystack.toLowerCase().indexOf(needle.toLowerCase(),start);return index<0?"#VALUE!":index+1;}
      case "REPLACE":{const old=text(a[0]),start=Math.max(1,number(a[1]))-1,count=Math.max(0,number(a[2]));return old.slice(0,start)+text(a[3])+old.slice(start+count);}
      case "SUBSTITUTE":{const source=text(a[0]),old=text(a[1]),replacement=text(a[2]),instance=number(a[3]??0);if(!old)return source;let count=0;const escaped=old.split("").map(ch=>".*+?^()|[]\\".includes(ch)?"\\"+ch:ch).join("").replace(/\$/g,"\\$").replace(/{/g,"\\{").replace(/}/g,"\\}");return source.replace(new RegExp(escaped,"g"),match=>{count++;return !instance||count===instance?replacement:match;});}
      case "VALUE":case "NUMBERVALUE":{const value=text(a[0]).replace(/,/g,"");const parsed=Number(value);return Number.isFinite(parsed)?parsed:"#VALUE!";}
      case "EXACT":return text(a[0])===text(a[1]);
      case "REPT":return text(a[0]).repeat(Math.max(0,Math.min(32767,Math.trunc(number(a[1])))));
      case "CHAR":return number(a[0])<1||number(a[0])>255?"#VALUE!":String.fromCharCode(Math.trunc(number(a[0])));
      case "CODE":return text(a[0]).length?text(a[0]).charCodeAt(0):"#VALUE!";
      case "T":return typeof a[0]==="string"?a[0]:"";
      case "FIXED":case "DOLLAR":{const digits=Math.max(0,Math.min(30,Math.trunc(number(a[1]??2)))),fixed=new Intl.NumberFormat("en-US",{minimumFractionDigits:digits,maximumFractionDigits:digits,useGrouping:!a[2]}).format(number(a[0]));return fn==="DOLLAR"?"$"+fixed:fixed;}
      case "TEXTBEFORE":case "TEXTAFTER":{const source=text(a[0]),delimiter=text(a[1]),instance=number(a[2]??1),positions:number[]=[];let from=0;while((from=source.indexOf(delimiter,from))>=0){positions.push(from);from+=delimiter.length||1;}const idx=instance>0?positions[instance-1]:positions[positions.length+instance];if(idx===undefined)return a[5]??"#N/A";return fn==="TEXTBEFORE"?source.slice(0,idx):source.slice(idx+delimiter.length);}
      case "TEXTSPLIT":return {__gridArray:true,values:[text(a[0]).split(text(a[1]??","))]};
      case "TEXT":return formatText(number(a[0]),text(a[1]),a[0]);
      case "DATE":{let year=Math.trunc(number(a[0]));if(year<1900)year+=1900;return serialFromDate(new Date(Date.UTC(year,Math.trunc(number(a[1]))-1,Math.trunc(number(a[2]))));}
      case "DATEVALUE":{const date=toDate(a[0]);return date?serialFromDate(date):"#VALUE!";}
      case "TODAY":return serialFromDate(new Date());
      case "NOW":{const now=new Date();return serialFromDate(now)+(now.getUTCHours()*3600+now.getUTCMinutes()*60+now.getUTCSeconds())/86400;}
      case "TIME":return ((number(a[0])*3600+number(a[1])*60+number(a[2]))%86400)/86400;
      case "TIMEVALUE":{const match=text(a[0]).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);if(!match)return "#VALUE!";let hour=Number(match[1]),minute=Number(match[2]),second=Number(match[3]??0);if(match[4]?.toUpperCase()==="PM"&&hour<12)hour+=12;if(match[4]?.toUpperCase()==="AM"&&hour===12)hour=0;return (hour*3600+minute*60+second)/86400;}
      case "YEAR":case "MONTH":case "DAY":case "HOUR":case "MINUTE":case "SECOND":{const date=toDate(a[0]);if(!date)return "#VALUE!";if(fn==="YEAR")return date.getUTCFullYear();if(fn==="MONTH")return date.getUTCMonth()+1;if(fn==="DAY")return date.getUTCDate();if(fn==="HOUR")return date.getUTCHours();if(fn==="MINUTE")return date.getUTCMinutes();return date.getUTCSeconds();}
      case "DAYS":{const x=toSerial(a[0]),y=toSerial(a[1]);return Number.isFinite(x)&&Number.isFinite(y)?x-y:"#VALUE!";}
      case "DATEDIF":{const start=toDate(a[0]),end=toDate(a[1]),unit=text(a[2]).toUpperCase();if(!start||!end||end<start)return "#NUM!";let years=end.getUTCFullYear()-start.getUTCFullYear(),months=years*12+end.getUTCMonth()-start.getUTCMonth();if(end.getUTCDate()<start.getUTCDate())months--;years=Math.floor(months/12);if(unit==="Y")return years;if(unit==="M")return months;if(unit==="D")return toSerial(a[1])-toSerial(a[0]);if(unit==="YM")return months%12;if(unit==="MD")return end.getUTCDate()-start.getUTCDate();if(unit==="YD"){const startThis=new Date(Date.UTC(end.getUTCFullYear(),start.getUTCMonth(),start.getUTCDate()));return (end.getTime()-startThis.getTime())/86400000;}return "#NUM!";}
      case "EDATE":{const date=toDate(a[0]);return date?serialFromDate(addMonths(date,Math.trunc(number(a[1])))):"#VALUE!";}
      case "EOMONTH":{const date=toDate(a[0]);if(!date)return "#VALUE!";const moved=addMonths(date,Math.trunc(number(a[1]))+1);return serialFromDate(new Date(Date.UTC(moved.getUTCFullYear(),moved.getUTCMonth(),0)));}
      case "WEEKDAY":{const date=toDate(a[0]);return date?excelWeekday(date,number(a[1]??1)):"#VALUE!";}
      case "ISOWEEKNUM":{const date=toDate(a[0]);if(!date)return "#VALUE!";const d=dateOnly(date);d.setUTCDate(d.getUTCDate()+4-(d.getUTCDay()||7));const yearStart=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.ceil((((d.getTime()-yearStart.getTime())/86400000)+1)/7);}
      case "WEEKNUM":{const date=toDate(a[0]);if(!date)return "#VALUE!";const start=new Date(Date.UTC(date.getUTCFullYear(),0,1));return Math.floor((toSerial(a[0])-serialFromDate(start)+excelWeekday(start,number(a[1]??1))-1)/7)+1;}
      case "DAYS360":{const d1=toDate(a[0]),d2=toDate(a[1]);if(!d1||!d2)return "#VALUE!";let day1=d1.getUTCDate(),day2=d2.getUTCDate(),m1=d1.getUTCMonth()+1,m2=d2.getUTCMonth()+1,y1=d1.getUTCFullYear(),y2=d2.getUTCFullYear();if(a[2]){if(day1===31)day1=30;if(day2===31&&day1===30)day2=30;}else{if(m1===2&&day1===new Date(Date.UTC(y1,m1,0)).getUTCDate())day1=30;if(m2===2&&day2===new Date(Date.UTC(y2,m2,0)).getUTCDate()&&day1===30)day2=30;if(day1===31)day1=30;if(day2===31&&day1===30)day2=30;}return 360*(y2-y1)+30*(m2-m1)+(day2-day1);}
      case "YEARFRAC":{const x=toSerial(a[0]),y=toSerial(a[1]),basis=number(a[2]??0);if(!Number.isFinite(x)||!Number.isFinite(y))return "#VALUE!";const days=y-x;if(basis===2)return days/360;if(basis===3)return days/365;if(basis===1){const d1=toDate(a[0])!,d2=toDate(a[1])!;return days/((d1.getUTCFullYear()+d2.getUTCFullYear())/2%4===0?366:365);}return days/360;}
      case "NETWORKDAYS":case "NETWORKDAYS.INTL":case "WORKDAY":case "WORKDAY.INTL":{const start=toDate(a[0]);if(!start)return "#VALUE!";const intl=fn.endsWith(".INTL"),isWork=fn.startsWith("WORKDAY"),weekend=intl?(a[2]??1):1,holidayIdx=isWork?(intl?3:2):(intl?3:2),holidays=new Set(list(holidayIdx).map(toSerial).filter(Number.isFinite));if(isWork){const date=dateOnly(start),days=number(a[1]),step=days<0?-1:1;let remaining=Math.abs(Math.trunc(days));while(remaining>0){date.setUTCDate(date.getUTCDate()+step);if(!isWeekend(date,weekend)&&!holidays.has(serialFromDate(date)))remaining--;}return serialFromDate(date);}const end=toDate(a[1]);if(!end)return "#VALUE!";const step=end>=start?1:-1;let count=0,date=dateOnly(start),last=dateOnly(end);while((step>0&&date<=last)||(step<0&&date>=last)){if(!isWeekend(date,weekend)&&!holidays.has(serialFromDate(date)))count+=step;date.setUTCDate(date.getUTCDate()+step);}return count;}
      case "SUMIF":case "COUNTIF":case "AVERAGEIF":{const range=list(0),criteria=a[1]??"",target=list(2).length?list(2):range;const matched=range.map((value,index)=>({value,result:target[index]??""})).filter(item=>criteriaMatches(item.value,criteria));const values=numeric(matched.map(item=>item.result));return fn==="COUNTIF"?matched.length:fn==="SUMIF"?sum(values):values.length?average(values):"#DIV/0!";}
      case "SUMIFS":case "COUNTIFS":case "AVERAGEIFS":{const target=list(0),pairs:{range:Value[];criteria:Value}[]=[];const offset=fn==="COUNTIFS"?0:1;for(let i=offset;i+1<args.length;i+=2)pairs.push({range:list(i),criteria:a[i+1]??""});const matched=target.filter((_,index)=>pairs.every(pair=>criteriaMatches(pair.range[index]??"",pair.criteria)));const values=numeric(matched);return fn==="COUNTIFS"?matched.length:fn==="SUMIFS"?sum(values):values.length?average(values):"#DIV/0!";}
      case "MAXIFS":case "MINIFS":{const target=list(0),pairs:{range:Value[];criteria:Value}[]=[];for(let i=1;i+1<args.length;i+=2)pairs.push({range:list(i),criteria:a[i+1]??""});const values=target.filter((_,index)=>pairs.every(pair=>criteriaMatches(pair.range[index]??"",pair.criteria))).map(number);return values.length?(fn==="MAXIFS"?Math.max(...values):Math.min(...values)):"#N/A";}
      case "PMT":return pmt(number(a[0]),number(a[1]),number(a[2]),number(a[3]??0),number(a[4]??0));
      case "IPMT":return ipmt(number(a[0]),number(a[1]),number(a[2]),number(a[3]),number(a[4]??0),number(a[5]??0));
      case "PPMT":return pmt(number(a[0]),number(a[2]),number(a[3]),number(a[4]??0),number(a[5]??0))-ipmt(number(a[0]),number(a[1]),number(a[2]),number(a[3]),number(a[4]??0),number(a[5]??0));
      case "FV":{const r=number(a[0]),n=number(a[1]),payment=number(a[2]),pv=number(a[3]??0),type=number(a[4]??0);return r===0?-(pv+payment*n):-(pv*(1+r)**n+payment*(1+r*type)*(((1+r)**n-1)/r));}
      case "PV":{const r=number(a[0]),n=number(a[1]),payment=number(a[2]),fv=number(a[3]??0),type=number(a[4]??0);return r===0?-(fv+payment*n):-(fv+payment*(1+r*type)*(((1+r)**n-1)/r))/(1+r)**n;}
      case "NPER":{const r=number(a[0]),payment=number(a[1]),pv=number(a[2]),fv=number(a[3]??0),type=number(a[4]??0);if(r===0)return -(pv+fv)/payment;const x=(payment*(1+r*type)-fv*r)/(pv*r+payment*(1+r*type));return Math.log(x)/Math.log(1+r);}
      case "RATE":{const n=number(a[0]),payment=number(a[1]),pv=number(a[2]),fv=number(a[3]??0),type=number(a[4]??0);let r=number(a[5]??0.1);for(let i=0;i<100;i++){const f=r===0?pv+payment*n+fv:pv*(1+r)**n+payment*(1+r*type)*(((1+r)**n-1)/r)+fv;const h=1e-6,rr=r+h,ff=pv*(1+rr)**n+payment*(1+rr*type)*(((1+rr)**n-1)/rr)+fv;const next=r-f/((ff-f)/h);if(!Number.isFinite(next))break;if(Math.abs(next-r)<1e-9)return next;r=next;}return r;}
      case "EFFECT":{const nominal=number(a[0]),n=number(a[1]);return (1+nominal/n)**n-1;}
      case "NOMINAL":{const effect=number(a[0]),n=number(a[1]);return n*((1+effect)**(1/n)-1);}
      case "SLN":return (number(a[0])-number(a[1]))/number(a[2]);
      case "SYD":{const cost=number(a[0]),salvage=number(a[1]),life=number(a[2]),period=number(a[3]);return (cost-salvage)*(life-period+1)*2/(life*(life+1));}
      case "DB":{const cost=number(a[0]),salvage=number(a[1]),life=number(a[2]),period=number(a[3]),month=number(a[4]??12),rate=1-Math.pow(salvage/cost,1/life);let book=cost,value=0;for(let p=1;p<=period;p++){value=p===1?cost*rate*month/12:book*rate;book-=value;}return value;}
      case "DDB":{const cost=number(a[0]),salvage=number(a[1]),life=number(a[2]),period=number(a[3]),factor=number(a[4]??2);let book=cost,value=0;for(let p=1;p<=period;p++){value=Math.min(book-salvage,book*factor/life);book-=value;}return value;}
      case "CUMIPMT":case "CUMPRINC":{let total=0;const rate=number(a[0]),n=number(a[1]),pv=number(a[2]),start=number(a[3]),end=number(a[4]),type=number(a[5]);for(let p=start;p<=end;p++)total+=fn==="CUMIPMT"?ipmt(rate,p,n,pv,0,type):pmt(rate,n,pv)-ipmt(rate,p,n,pv,0,type);return total;}
      case "FVSCHEDULE":{let value=number(a[0]);for(const rate of list(1))value*=1+number(rate);return value;}
      case "PDURATION":return Math.log(number(a[2])/number(a[1]))/Math.log(1+number(a[0]));
      case "IRR":return irr(nums,number(a[1]??0.1));
      case "XIRR":return xirr(list(0).map(number),list(1),number(a[2]??0.1));
      case "MIRR":{const cash=list(0).map(number),finance=number(a[1]),reinvest=number(a[2]),n=cash.length;if(!n)return "#NUM!";let positive=0,negative=0;cash.forEach((v,i)=>{if(v>0)positive+=v*(1+reinvest)**(n-1-i);if(v<0)negative+=v/(1+finance)**i;});return Math.pow(-positive/negative,1/(n-1))-1;}
      case "XNPV":{const cash=list(1).map(number),dates=list(2).map(toSerial),rate=number(a[0]);if(!dates.length||dates.some(v=>!Number.isFinite(v)))return "#NUM!";return cash.reduce((acc,v,i)=>acc+v/(1+rate)**((dates[i]-dates[0])/365),0);}
      case "VLOOKUP":case "HLOOKUP":{const data=matrix(1),lookup=a[0],index=Math.trunc(number(a[2]))-1,approx=a[3]!==false;if(!data.length)return "#N/A";if(fn==="VLOOKUP"){const row=data.find(row=>String(row[0]).toLowerCase()===String(lookup).toLowerCase());if(row)return row[index]??"#REF!";if(approx){const sorted=[...data].sort((x,y)=>number(x[0])-number(y[0]));const row=[...sorted].reverse().find(x=>number(x[0])<=number(lookup));return row?.[index]??"#N/A";}}else{const header=data[0]??[],column=header.findIndex(v=>String(v).toLowerCase()===String(lookup).toLowerCase());if(column>=0)return data[index]?.[column]??"#REF!";}return "#N/A";}
      case "XLOOKUP":case "XMATCH":case "MATCH":{const look=a[0],source=list(1),idx=source.findIndex(v=>String(v).toLowerCase()===String(look).toLowerCase());if(fn==="XMATCH"||fn==="MATCH")return idx>=0?idx+1:"#N/A";const ret=list(2);if(idx>=0)return ret[idx]??"#N/A";return a[3]??"#N/A";}
      case "INDEX":{const data=matrix(0),row=Math.max(1,number(a[1]))-1,col=Math.max(1,number(a[2]??1))-1;return data[row]?.[col]??"#REF!";}
      case "LOOKUP":{const source=list(1),target=list(2),look=number(a[0]);let idx=-1;source.forEach((v,i)=>{if(number(v)<=look)idx=i;});return idx>=0?(target[idx]??source[idx]):"#N/A";}
      case "CHOOSE":{const idx=Math.trunc(number(a[0]));return a[idx]??"#VALUE!";}
      case "OFFSET":return args[0]?.[0]??"#REF!";
      case "INDIRECT":{const ref=text(a[0]).replace(/\$/g,"").toUpperCase();return refPattern.test(ref)?this.resolveReference(ref):"#REF!";}
      case "ROW":{const p=this.currentAddress?parseAddress(this.currentAddress):null;return p?p.row+1:1;}
      case "ROWS":return matrix(0).length;
      case "COLUMN":{const p=this.currentAddress?parseAddress(this.currentAddress):null;return p?p.col+1:1;}
      case "COLUMNS":return matrix(0)[0]?.length??0;
      case "ADDRESS":{const row=Math.max(1,number(a[0])),col=Math.max(1,number(a[1]));const abs=Math.trunc(number(a[2]??1)),a1=a[3]!==false;let n=col,label="";while(n){const rem=(n-1)%26;label=String.fromCharCode(65+rem)+label;n=Math.floor((n-1)/26);}return a1?(abs===1?"$":"")+label+(abs===1?"$":"")+row:(col===1?"":"R"+row)+"C"+col;}
      case "HYPERLINK":return a[1]??a[0]??"";
      case "TRANSPOSE":{const data=matrix(0),width=Math.max(0,...data.map(row=>row.length));return {__gridArray:true,values:Array.from({length:width},(_,col)=>data.map(row=>row[col]??""))};}
      case "UNIQUE":{const data=matrix(0),byCol=!!a[1],source=byCol?data[0]?.map((_,col)=>data.map(row=>row[col]))??[]:data,seen=new Set<string>(),out=source.filter(row=>{const key=JSON.stringify(row);if(seen.has(key))return false;seen.add(key);return true;});const result=byCol?Array.from({length:out[0]?.length??0},(_,col)=>out.map(row=>row[col])):out;return {__gridArray:true,values:result};}
      case "FILTER":{const data=matrix(0),include=list(1),filtered=data.filter((_,i)=>!!include[i]);return filtered.length?{__gridArray:true,values:filtered}:a[2]??"#CALC!";}
      case "SORT":{const data=matrix(0),col=Math.max(1,number(a[1]??1))-1,order=number(a[2]??1);return {__gridArray:true,values:[...data].sort((x,y)=>(x[col] as any)>(y[col] as any)?order:-order)};}
      case "MEDIAN":return nums.length?percentile(nums,0.5):"#NUM!";
      case "MODE":{const counts=new Map<number,number>();nums.forEach(v=>counts.set(v,(counts.get(v)??0)+1));const sorted=[...counts].sort((x,y)=>y[1]-x[1]);return sorted[0]?.[1]>1?sorted[0][0]:"#N/A";}
      case "LARGE":{const values=[...nums].sort((x,y)=>y-x);return values[number(a[1])-1]??"#NUM!";}
      case "SMALL":{const values=[...nums].sort((x,y)=>x-y);return values[number(a[1])-1]??"#NUM!";}
      case "RANK":{const ref=list(1).map(number);return ref.filter(v=>number(a[2]??0)===0?v>number(a[0]):v<number(a[0])).length+1;}
      case "PERCENTILE":return percentile(nums,number(a[1]));
      case "QUARTILE":return percentile(nums,number(a[1])/4);
      case "STDEV.S":return Math.sqrt(variance(nums,true));
      case "STDEV.P":return Math.sqrt(variance(nums,false));
      case "VAR.S":return variance(nums,true);
      case "VAR.P":return variance(nums,false);
      case "CORREL":{const x=list(0).map(number),y=list(1).map(number),n=Math.min(x.length,y.length),mx=average(x.slice(0,n)),my=average(y.slice(0,n));const top=x.slice(0,n).reduce((s,v,i)=>s+(v-mx)*(y[i]-my),0),bottom=Math.sqrt(x.slice(0,n).reduce((s,v)=>s+(v-mx)**2,0)*y.slice(0,n).reduce((s,v)=>s+(v-my)**2,0));return bottom?top/bottom:"#DIV/0!";}
      case "FORECAST":{const x=list(2).map(number),y=list(1).map(number),mx=average(x),my=average(y),slope=x.reduce((s,v,i)=>s+(v-mx)*(y[i]-my),0)/x.reduce((s,v)=>s+(v-mx)**2,0);return my+slope*(number(a[0])-mx);}
      case "GEOMEAN":return nums.length&&nums.every(v=>v>0)?Math.pow(nums.reduce((p,v)=>p*v,1),1/nums.length):"#NUM!";
      case "HARMEAN":return nums.length&&nums.every(v=>v>0)?nums.length/nums.reduce((s,v)=>s+1/v,0):"#NUM!";
      case "DEC2BIN":case "DEC2HEX":case "DEC2OCT":{const base=fn==="DEC2BIN"?2:fn==="DEC2HEX"?16:8,n=Math.trunc(number(a[0]));return n<0?"#NUM!":n.toString(base).toUpperCase().padStart(Math.trunc(number(a[1]??0)),"0");}
      case "BIN2DEC":case "HEX2DEC":case "OCT2DEC":{const base=fn==="BIN2DEC"?2:fn==="HEX2DEC"?16:8,parsed=parseInt(text(a[0]),base);return Number.isFinite(parsed)?parsed:"#NUM!";}
      case "CONVERT":{const unit=text(a[1]).toLowerCase(),target=text(a[2]).toLowerCase(),factors:Record<string,number>={m:1,km:1000,cm:0.01,mm:0.001,mi:1609.344,ft:0.3048,in:0.0254,yd:0.9144,kg:1,g:0.001,lb:0.45359237,oz:0.028349523125,l:1,ml:0.001,gal:3.785411784};return factors[unit]&&factors[target]?number(a[0])*factors[unit]/factors[target]:"#N/A";}
      case "DELTA":return number(a[0])===number(a[1]??0)?1:0;
      case "GESTEP":return number(a[0])>=number(a[1]??0)?1:0;
      case "COMPLEX":return String(number(a[0]))+(number(a[1])<0?"-":"+")+Math.abs(number(a[1]))+(a[2]??"i");
      case "IMABS":{const match=text(a[0]).match(/^([+-]?\d*\.?\d+)([+-]\d*\.?\d+)[ij]$/i);return match?Math.hypot(Number(match[1]),Number(match[2])):"#NUM!";}
      case "IMSUM":return list(0).map(text).concat(list(1).map(text)).join("");
      case "ENCODEURL":return encodeURIComponent(text(a[0]));
      case "SUBTOTAL":{const code=number(a[0]),values=numeric(args.slice(1).flat());if([1,101].includes(code))return values.length?average(values):0;if([2,102].includes(code))return values.length;if([3,103].includes(code))return flat.length;if([4,104].includes(code))return values.length?Math.max(...values):0;if([5,105].includes(code))return values.length?Math.min(...values):0;if([9,109].includes(code))return sum(values);return "#VALUE!";}
      case "AGGREGATE":{const code=number(a[0]),values=numeric(args.slice(2).flat());if(code===1)return average(values);if(code===2)return values.length;if(code===4)return Math.max(...values);if(code===5)return Math.min(...values);if(code===9)return sum(values);return "#VALUE!";}
      default:return "#NAME?";
    }
  }
}
function validReference(token:string){return refPattern.test(token)||columnPattern.test(token)||rowPattern.test(token);}

function formatText(value:number,format:string,original?:Value):string{
 const date=toDate(typeof original==="number"?original:value);
 if(/[dy]/i.test(format)&&date){
  const months=["January","February","March","April","May","June","July","August","September","October","November","December"];
  const month=months[date.getUTCMonth()],replacements:Record<string,string>={
   "yyyy":String(date.getUTCFullYear()),"yy":String(date.getUTCFullYear()).slice(-2),
   "mmmm":month,"mmm":month.slice(0,3),"mm":String(date.getUTCMonth()+1).padStart(2,"0"),
   "m":String(date.getUTCMonth()+1),"dd":String(date.getUTCDate()).padStart(2,"0"),"d":String(date.getUTCDate()),
   "hh":String(date.getUTCHours()).padStart(2,"0"),"h":String(date.getUTCHours()),
   "ss":String(date.getUTCSeconds()).padStart(2,"0")
  };
  return format.replace(/AM\/PM|yyyy|mmmm|mmm|yy|mm|dd|hh|ss|m|d|h/gi,token=>{
   if(token.toUpperCase()==="AM/PM")return date.getUTCHours()<12?"AM":"PM";
   return replacements[token.toLowerCase()]??token;
  });
 }
 if(format.includes("%"))return new Intl.NumberFormat("en-US",{style:"percent",maximumFractionDigits:(format.split(".")[1]??"").replace(/[^0]/g,"").length}).format(value);
 const decimals=format.match(/\\.([0#]+)/)?.[1].length??0;
 const grouped=format.includes(",");
 if(/^[\$₹€£]/.test(format)){const currency=format[0];return currency+new Intl.NumberFormat("en-US",{minimumFractionDigits:decimals,maximumFractionDigits:decimals,useGrouping:grouped}).format(value);}
 return new Intl.NumberFormat("en-US",{minimumFractionDigits:format.includes(".")?decimals:0,maximumFractionDigits:decimals,useGrouping:grouped}).format(value);
}

function roundTo(value: number, digits: number): number { const factor = Math.pow(10, digits); return Math.round((value + Number.EPSILON) * factor) / factor; }
function roundDirected(value: number, digits: number, direction: "up" | "down"): number {
  const factor = Math.pow(10, digits), scaled = value * factor;
  return (direction === "up" ? (scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled)) : (scaled >= 0 ? Math.floor(scaled) : Math.ceil(scaled))) / factor;
}

export function evaluate(formula: string, cells: CellMap, seen = new Set<string>(), tables: ExcelTableData[] = [], currentAddress?: string, namedRanges:Record<string,string>={},sheets:SheetData[]=[],currentSheetName?:string): Value {
  try {
    const source = formula.startsWith("=") ? formula.slice(1) : formula;
    const structured = expandStructuredReferences(source, cells, tables, currentAddress);
    const expanded = expandNamedRanges(structured,namedRanges);
    return new Parser(tokenize(expanded), cells, seen, tables, namedRanges,currentAddress,sheets,currentSheetName).parse();
  } catch (error) {
    return error instanceof Error ? error.message : "#ERROR!";
  }
}
export function displayValue(raw: string, cells: CellMap, showFormula = false, tables: ExcelTableData[] = [], currentAddress?: string, namedRanges:Record<string,string>={},sheets:SheetData[]=[],currentSheetName?:string): string {
  if (!raw.startsWith("=")) return raw;
  return showFormula ? raw : String(scalar(evaluate(raw, cells, new Set(), tables, currentAddress, namedRanges,sheets,currentSheetName)));
}
export function numericValue(raw: string, cells: CellMap, tables: ExcelTableData[] = [], currentAddress?: string, namedRanges:Record<string,string>={},sheets:SheetData[]=[],currentSheetName?:string): number | null {
  const value = raw.startsWith("=") ? evaluate(raw, cells, new Set(), tables, currentAddress, namedRanges,sheets,currentSheetName) : raw;
  const result=scalar(value);
  if (typeof result === "number") return Number.isFinite(result) ? result : null;
  if (typeof result === "string" && result.trim() !== "" && Number.isFinite(Number(result))) return Number(result);
  return null;
}
export function recalculate(cells: CellMap, tables:ExcelTableData[]=[], namedRanges:Record<string,string>={},sheets:SheetData[]=[],currentSheetName?:string): CellMap {
  const next={...cells};
  for(const [address,cell] of Object.entries(cells)){
    if(!cell.raw.startsWith("="))continue;
    const value=evaluate(cell.raw,cells,new Set(),tables,address,namedRanges,sheets,currentSheetName);
    next[address]={...cell,cachedValue:String(scalar(value))};
  }
  return next;
}
export function formatCellValue(raw: string, cells: CellMap, showFormula: boolean, style?: CellStyle, tables: ExcelTableData[] = [], currentAddress?: string, namedRanges:Record<string,string>={},sheets:SheetData[]=[],currentSheetName?:string): string {
  if (showFormula && raw.startsWith("=")) return raw;
  const evaluated = raw.startsWith("=") ? evaluate(raw, cells, new Set(), tables, currentAddress, namedRanges,sheets,currentSheetName) : raw;
  const value=scalar(evaluated);
  if (typeof value === "string" && value.startsWith("#")) return value;
  const n = typeof value === "number" ? value : (String(value).trim() !== "" && Number.isFinite(Number(value)) ? Number(value) : null);
  switch (style?.numberFormat) {
    case "Percentage %": return n === null ? String(value) : new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Currency ₹": return n === null ? String(value) : new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Currency $": return n === null ? String(value) : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Number": return n === null ? String(value) : new Intl.NumberFormat("en-IN", { maximumFractionDigits: style?.decimalPlaces ?? 2 }).format(n);
    case "Date": {
      const date = toDate(value);
      return date ? date.toLocaleDateString() : String(value);
    }
    default: return String(value ?? "");
  }
}
