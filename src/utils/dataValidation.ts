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
  const source=rule.listSource??rule.formula1??"";
  if(source.trim().startsWith("=")){
    const match=source.trim().slice(1).match(/^(?:'([^']+)'|([^!]+))?!?\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i);
    if(!match)return [];
    const sheetName=match[1]??match[2],target=sheetName?sheetName===sheet.name?sheet:undefined:sheet;
    if(!target)return [];
    const start=parseAddress(match[3]+match[4]),end=parseAddress((match[5]??match[3])+(match[6]??match[4]));
    if(!start||!end)return [];
    const values:string[]=[];
    for(let row=start.row;row<=end.row;row++)for(let col=start.col;col<=end.col;col++){const value=target.cells[addressOf(row,col)]?.raw??"";if(value!==""&&!values.includes(value))values.push(value);}
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
export function validateCellValue(rule:DataValidationRule,value:string,sheet:SheetData,address:string):ValidationResult{
  if(value===""&&rule.allowBlank!==false)return {valid:true,message:""};
  let valid=true;
  switch(rule.type){
    case "List":valid=getListValidationOptions(rule,sheet).some(option=>option.toLowerCase()===value.toLowerCase());break;
    case "Whole Number":valid=/^-?\d+$/.test(value)&&compare(value,rule.formula1??"0",rule.operator??"greater than or equal to");if(valid&&["between","not between"].includes(rule.operator??""))valid=compare(value,rule.formula2??rule.formula1??"0",rule.operator==="between"?"less than or equal to":"greater than");break;
    case "Decimal":valid=value.trim()!==""&&Number.isFinite(Number(value))&&compare(value,rule.formula1??"0",rule.operator??"greater than or equal to");if(valid&&["between","not between"].includes(rule.operator??""))valid=compare(value,rule.formula2??rule.formula1??"0",rule.operator==="between"?"less than or equal to":"greater than");break;
    case "Date":{const parsed=Date.parse(value),min=Date.parse(rule.formula1??"1900-01-01"),max=Date.parse(rule.formula2??rule.formula1??"9999-12-31");valid=Number.isFinite(parsed)&&compare(String(parsed),String(min),rule.operator??"greater than or equal to");if(valid&&rule.operator==="between")valid=parsed<=max;break;}
    case "Text Length":valid=compare(String(value.length),rule.formula1??"0",rule.operator??"greater than or equal to");if(rule.operator==="between")valid=value.length<=Number(rule.formula2);break;
    case "Custom Formula":{
      const base=parseAddress(address),origin={row:rule.range.startRow,col:rule.range.startCol};
      const dr=(base?.row??origin.row)-origin.row,dc=(base?.col??origin.col)-origin.col;
      const formula=(rule.formula1??"=TRUE").replace(/(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)/g,(match,colLock,letters,rowLock,digits)=>{
        const p=parseAddress(letters+digits);if(!p)return match;
        return (colLock?"$"+letters:addressOf(0,p.col+dc).replace(/[0-9]/g,""))+(rowLock?"$"+digits:String(p.row+dr+1));
      });
      const result=evaluate(formula,sheet.cells,new Set(),sheet.tables??[],address);
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
