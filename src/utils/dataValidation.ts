import type { DataValidationRule, SheetData } from "../types";
import { addressOf, parseAddress } from "../types";
import { evaluate } from "./formulaEngine";

export function rangeAddresses(range:DataValidationRule["range"]):string[]{
  const result:string[]=[];
  for(let row=range.startRow;row<=range.endRow;row++)for(let col=range.startCol;col<=range.endCol;col++)result.push(addressOf(row,col));
  return result;
}
export function getValidationRule(sheet:SheetData,address:string):DataValidationRule|undefined{
  const point=parseAddress(address);if(!point)return undefined;
  return (sheet.validationRules??[]).filter(rule=>point.row>=rule.range.startRow&&point.row<=rule.range.endRow&&point.col>=rule.range.startCol&&point.col<=rule.range.endCol).sort((a,b)=>b.id.localeCompare(a.id))[0];
}
export function getListValidationOptions(rule:DataValidationRule,sheet:SheetData):string[]{
  const source=(rule.listSource??rule.formula1??"").trim();
  if(source.startsWith("=")){
    const expression=source.slice(1);
    const separator=expression.lastIndexOf("!");
    const sheetName=separator>=0?expression.slice(0,separator).replace(/^'|'$/g,"").replace(/''/g,"'"):undefined;
    const reference=separator>=0?expression.slice(separator+1):expression;
    const match=reference.match(/^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i);
    if(!match)return [];
    const target=sheetName?sheet.name.toLowerCase()===sheetName.toLowerCase()?sheet:undefined:sheet;
    if(!target)return [];
    const start=parseAddress(match[1]+match[2]),end=parseAddress((match[3]??match[1])+(match[4]??match[2]));
    if(!start||!end)return [];
    const values:string[]=[];
    for(let row=Math.min(start.row,end.row);row<=Math.max(start.row,end.row);row++)for(let col=Math.min(start.col,end.col);col<=Math.max(start.col,end.col);col++){
      const value=target.cells[addressOf(row,col)]?.raw??"";
      if(value!==""&&!values.includes(value))values.push(value);
    }
    return values;
  }
  return source.split(",").map(value=>value.trim()).filter(Boolean);
}
function compare(left:string,right:string,operator:NonNullable<DataValidationRule["operator"]>):boolean{
  const a=Number(left),b=Number(right),numeric=left.trim()!==""&&right.trim()!==""&&Number.isFinite(a)&&Number.isFinite(b);
  const x=numeric?a:left.toLowerCase(),y=numeric?b:right.toLowerCase();
  switch(operator){
    case "between":return numeric?x>=y&&x<=y:left.localeCompare(right)>=0&&left.localeCompare(right)<=0;
    case "not between":return numeric?x<y||x>y:left.localeCompare(right)!==0;
    case "equal to":return left.toLowerCase()===right.toLowerCase()||x===y;
    case "not equal to":return !(left.toLowerCase()===right.toLowerCase()||x===y);
    case "greater than":return x>y;
    case "less than":return x<y;
    case "greater than or equal to":return x>=y;
    case "less than or equal to":return x<=y;
    default:return true;
  }
}
export interface ValidationResult { valid:boolean; message:string; }
export function validateCellValue(rule:DataValidationRule,value:string,sheet:SheetData,address:string,namedRanges:Record<string,string>={},sheets:SheetData[]=[sheet]):ValidationResult{
  if(value===""&&rule.allowBlank!==false)return {valid:true,message:""};
  let valid=true;
  const operator=rule.operator??"greater than or equal to";
  switch(rule.type){
    case "List":valid=getListValidationOptions(rule,sheet).some(option=>option.toLowerCase()===value.toLowerCase());break;
    case "Whole Number":case "Decimal":{
      const numeric=Number(value),min=Number(rule.formula1??0),max=Number(rule.formula2??rule.formula1??0);
      valid=value.trim()!==""&&Number.isFinite(numeric)&&(rule.type==="Decimal"||Number.isInteger(numeric));
      if(valid){
        switch(operator){
          case "between":valid=numeric>=min&&numeric<=max;break;
          case "not between":valid=numeric<min||numeric>max;break;
          case "equal to":valid=numeric===min;break;
          case "not equal to":valid=numeric!==min;break;
          case "greater than":valid=numeric>min;break;
          case "less than":valid=numeric<min;break;
          case "greater than or equal to":valid=numeric>=min;break;
          case "less than or equal to":valid=numeric<=min;break;
        }
      }
      break;
    }
    case "Date":{
      const date=Date.parse(value),min=Date.parse(rule.formula1??"1900-01-01"),max=Date.parse(rule.formula2??rule.formula1??"9999-12-31");
      valid=Number.isFinite(date);
      if(valid){
        switch(operator){
          case "between":valid=date>=min&&date<=max;break;
          case "not between":valid=date<min||date>max;break;
          case "equal to":valid=date===min;break;
          case "not equal to":valid=date!==min;break;
          case "greater than":valid=date>min;break;
          case "less than":valid=date<min;break;
          case "greater than or equal to":valid=date>=min;break;
          case "less than or equal to":valid=date<=min;break;
        }
      }
      break;
    }
    case "Text Length":{
      const length=value.length,min=Number(rule.formula1??0),max=Number(rule.formula2??rule.formula1??0);
      switch(operator){
        case "between":valid=length>=min&&length<=max;break;
        case "not between":valid=length<min||length>max;break;
        case "equal to":valid=length===min;break;
        case "not equal to":valid=length!==min;break;
        case "greater than":valid=length>min;break;
        case "less than":valid=length<min;break;
        case "greater than or equal to":valid=length>=min;break;
        case "less than or equal to":valid=length<=min;break;
      }
      break;
    }
    case "Custom Formula":{
      const base=parseAddress(address),origin={row:rule.range.startRow,col:rule.range.startCol};
      const dr=(base?.row??origin.row)-origin.row,dc=(base?.col??origin.col)-origin.col;
      const formula=(rule.formula1??"=TRUE").replace(/(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)/g,(match,colLock,letters,rowLock,digits)=>{
        const point=parseAddress(letters+digits);if(!point)return match;
        const column=colLock?letters:addressOf(0,point.col+dc).replace(/[0-9]/g,"");
        return (colLock?"$"+column:column)+(rowLock?"$"+digits:String(point.row+dr+1));
      });
      const result=evaluate(formula,sheet.cells,new Set(),sheet.tables??[],address,namedRanges,sheets,sheet.name);
      valid=result===true||result===1||String(result).toUpperCase()==="TRUE";break;
    }
  }
  return {valid,message:valid?"":rule.errorMessage||"The value you entered doesn't meet this cell's data validation rules."};
}
export interface CellValidationState { rule:DataValidationRule; options:string[]; }
export function buildValidationMap(sheet:SheetData):Record<string,CellValidationState>{
  const result:Record<string,CellValidationState>={};
  for(const rule of sheet.validationRules??[]){
    const options=rule.type==="List"?getListValidationOptions(rule,sheet):[];
    for(const address of rangeAddresses(rule.range))result[address]={rule,options};
  }
  return result;
}
