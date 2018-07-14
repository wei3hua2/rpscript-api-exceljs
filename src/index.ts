/**
 * @module exceljs
 */

import ExcelJS,{Workbook,Worksheet, AutoFilter,Row,Column, ColumnExtension, Cell} from 'exceljs';
import {RpsContext,RpsModule,rpsAction} from 'rpscript-interface';
import R from 'ramda';

let MOD_ID = "exceljs"

export interface ExcelContext {
  workbook?:Workbook
}

@RpsModule(MOD_ID)
export default class RPSModule {

  constructor(ctx:RpsContext){
    ctx.addModuleContext(MOD_ID,{});
  }

  @rpsAction({verbName:'new-excel'})
  async  newExcel(ctx:RpsContext,opts:Object) : Promise<Workbook>{
    let workbook = new Workbook();
    workbook.creator = opts['creator'];
    workbook.lastModifiedBy = opts['lastModifiedBy'];
    workbook.properties.date1904 = opts['date1904'];

    ctx.addModuleContext(MOD_ID,{workbook:workbook});

    return workbook;
  }

  @rpsAction({verbName:'read-excel'})
  async  readExcel(ctx:RpsContext,opts:Object, filename:string) : Promise<Workbook>{
    var workbook = new Workbook();
    await workbook.xlsx.readFile(filename);

    ctx.addModuleContext(MOD_ID,{workbook:workbook});

    return workbook;
  }

  @rpsAction({verbName:'write-excel'})
  async  writeExcel(ctx:RpsContext,opts:Object, filename:string) : Promise<Workbook>{
    let workbook = this.getCurrentWorkbook(ctx);
    await workbook.xlsx.writeFile(filename);

    ctx.addModuleContext(MOD_ID,{workbook:null});

    return workbook;
  }

  @rpsAction({verbName:'add-worksheet'})
  async  addWorksheet(ctx:RpsContext,opts:Object,sheetname:string) : Promise<Worksheet>{
    let workbook:Workbook = this.getCurrentWorkbook(ctx);
    
    return workbook.addWorksheet(sheetname);
  }
  @rpsAction({verbName:'remove-worksheet'})
  async  removeWorksheet(ctx:RpsContext,opts:Object,sheetname:string|number) : Promise<void>{
    let workbook:Workbook = this.getCurrentWorkbook(ctx);
    return workbook.removeWorksheet(sheetname);
  }
  @rpsAction({verbName:'duplicate-worksheet'})
  async  dupWorksheet(ctx:RpsContext,opts:Object,sheetname:string, newName:string) : Promise<Worksheet>{
    let workbook:Workbook = this.getCurrentWorkbook(ctx);
    
    let ws = workbook.addWorksheet(newName);
    let duWs = this.getWorksheet(ctx,sheetname);

    ws.columns = duWs.columns;
    ws.addRows( duWs.getSheetValues() );

    return ws;
  }


  @rpsAction({verbName:'excel-auto-filter'})
  async  autofilter(ctx:RpsContext,opts:Object, sheetname:string|number,
    param1:any, param2?:any, param3?:any, param4?:any) : Promise<AutoFilter>{
    let sheet = this.getWorksheet(ctx,sheetname);
    
    if(!param2) sheet.autoFilter = param1;
    else if (!param3) sheet.autoFilter = {from:param1,to:param2};
    else sheet.autoFilter = { from :{row:param1,column:param2} , to:{row:param3,column:param4} };
    
    return sheet.autoFilter;
  }

  @rpsAction({verbName:'worksheet-add-headers'})
  async addHeaders(ctx:RpsContext,opts:Object, sheetname:string|number, 
    headers:string[]) : Promise<any>{
      let sheet = this.getWorksheet(ctx,sheetname);
      sheet.columns = R.map(h => {
        return {header:h,key:h}
      },headers);

      return Promise.resolve(true);
  }

  @rpsAction({verbName:'worksheet-add-header'})
  async addHeader(ctx:RpsContext,opts:Object, sheetname:string|number, 
    header:string, key?:string, width?:number,outlineLevel?:number) : Promise<any>{
      let sheet = this.getWorksheet(ctx,sheetname);
      sheet.columns = sheet.columns || [];
      
      // let h:any = {header:header};
      // h.width = width ? width : 20;
      // h.key = key ? key : header;
      // h.outlineLevel = outlineLevel ? outlineLevel : 0;
      
      sheet.columns.push({
        header:header, 
        key:key ? key : header,
        outlineLevel : outlineLevel ? outlineLevel : 0});

      return Promise.resolve(true);
  }
  
  @rpsAction({verbName:'for-each-column-cell'})
  async  eachColumnCell(ctx:RpsContext,opts:Object, 
    sheetname:string|number, col:string|number, 
    perform:(cell,rowNo)=>void) : Promise<void>{
      
      let sheet = this.getWorksheet(ctx,sheetname);
      let column = sheet.getColumn(col);

      column.eachCell(perform);
  }

  @rpsAction({verbName:'worksheet-get-column'})
  async  getColumn(ctx:RpsContext,opts:Object, sheetname:string|number,col:string|number) : Promise<Column|ColumnExtension>{
    let sheet = this.getWorksheet(ctx,sheetname);
    return sheet.getColumn(col);
  }
  @rpsAction({verbName:'worksheet-get-row'})
  async  getRow(ctx:RpsContext,opts:Object, sheetname:string|number,row:number) : Promise<Row>{
    let sheet = this.getWorksheet(ctx,sheetname);
    return sheet.getRow(row);
  }
  @rpsAction({verbName:'worksheet-get-cell'})
  async  getCell(ctx:RpsContext,opts:Object, sheetname:string|number,cell:string) : Promise<Cell>{
    let sheet = this.getWorksheet(ctx,sheetname);
    
    if(opts['verbose']) return sheet.getCell(cell);
    else return sheet.getCell(cell);
  }

  @rpsAction({verbName:'worksheet-add-row'})
  async  addRow(ctx:RpsContext,opts:Object, sheetname:string|number,data:Object|Array<any>) : Promise<Row>{
    let sheet = this.getWorksheet(ctx,sheetname);
    return sheet.addRow(data);
  }
  @rpsAction({verbName:'worksheet-add-rows'})
  async  addRows(ctx:RpsContext,opts:Object, sheetname:string|number,datas:any) : Promise<void>{
    let sheet = this.getWorksheet(ctx,sheetname);
    return sheet.addRows(datas);
  }

  @rpsAction({verbName:'worksheet-set-cell'})
  async  setCell(ctx:RpsContext,opts:Object, sheetname:string|number,cellPosition:string, value?:any) : Promise<void>{
    let cell:Cell = this.getWorksheet(ctx,sheetname).getCell(cellPosition);
    let formula = opts['formula'], numFmt = opts['numberFormat'];
    

    if(value) cell.value = value;
    if(formula) cell.value = cell.value = { formula: formula, result: value || '?'};

    if(numFmt) cell.numFmt = numFmt;
    
    let font = this.parseFont(opts);
    let alignment = this.parseAlignment(opts);
    let border = this.parseBorder(opts);
    let fill = this.parseFill(opts);

    if(font) cell.font = font;
    if(alignment) cell.alignment = alignment;
    if(border) cell.border = border;
    if(fill) cell.fill = fill;
  }

  private parseFont (opts:Object) {

    let font = {
      name : opts['fontName'],
      family : opts['fontFamily'],
      size : opts['fontSize'],
      underline : opts['fontUnderline'],
      bold : opts['fontBold'],
      italic : opts['fontItalic'],
      strike : opts['fontStrike'],
      outline : opts['fontOutline'],
      color : opts['fontColor'] ? {argb: opts['fontColor']} : undefined
    };

    font = R.reject(R.isNil,font);

    if(R.keys(font).length === 0) return undefined;
    else return font;
  }

  private parseAlignment (opts:Object) {

    let alignment = {
      horizontal : opts['horizontal'],
      vertical : opts['vertical'],
      wrapText : opts['wrapText'],
      indent : opts['indent'],
      readingOrder : opts['readingOrder'],
      textRotation : opts['textRotation']
    };

    alignment = R.reject(R.isNil,alignment);

    if(R.keys(alignment).length === 0) return undefined;
    else return alignment;
  }

  private parseBorder (opts:Object) {
    let border = {
      top : {style:opts['topStyle'],color:opts['topColor']},
      bottom : {style:opts['bottomStyle'],color:opts['bottomColor']},
      left : {style:opts['leftStyle'],color:opts['leftColor']},
      right : {style:opts['rightStyle'],color:opts['rightColor']}
    };

    border = R.reject(R.isNil,border);

    if(R.keys(border).length === 0) return undefined;
    else return border;
  }
  private parseFill (opts:Object) {
    let fill = {
      type: opts['fillType'],
      pattern:opts['fillPattern'],
      fgColor : opts['fgColor'] ? {argb: opts['fgColor']} : undefined,
      bgColor : opts['bgColor'] ? {argb: opts['bgColor']} : undefined
    }

    fill = R.reject(R.isNil,fill);

    if(R.keys(fill).length === 0) return undefined;
    else return fill;
  }

  @rpsAction({verbName:'worksheet-append-column'})
  async appendColumn(ctx:RpsContext,opts:Object, 
    sheetname:string|number, columnName:string, 
    data?:string|number|Function|Array<any> ) : Promise<void>{

    let formula = opts['formula'], numFmt = opts['numberFormat'], width = opts['width'];
    let sheet:Worksheet = this.getWorksheet(ctx,sheetname);
    let colPosition = sheet.actualColumnCount+1;

    let result = [columnName];
    result = result.concat(await this.setColumnData(sheet,data));
    
    sheet.spliceColumns(colPosition,0,result);

    //update formula
    this.setFormula(sheet,formula);
    
    if(numFmt) sheet.getColumn(colPosition).numFmt = numFmt;
    
    sheet.getColumn(colPosition).width = width || 20;

    
    let font = this.parseFont(opts);
    let alignment = this.parseAlignment(opts);
    let border = this.parseBorder(opts);
    let fill = this.parseFill(opts);

    if(font) sheet.getColumn(colPosition).font = font;
    if(alignment) sheet.getColumn(colPosition).alignment = alignment;
    if(border) sheet.getColumn(colPosition).border = border;
    if(fill) sheet.getColumn(colPosition).fill = fill;
  }

  private async setColumnData (sheet:Worksheet,data?:string|number|Function|Array<any>) :Promise<Array<any>>{
    let result = [];
    if(typeof data==='string' || typeof data==='number')
      result = result.concat(R.repeat(data,sheet.actualRowCount-1));

    else if(typeof data==='function'){
      for(var i =1;i<sheet.actualRowCount;i++){
        let row = sheet.getRow(i);
        let output = await data(row,i);
        result.push(output);
      }
    }
    return result;
  }

  private setFormula (sheet:Worksheet,formula?:string) :void {
    if(formula){
      for(var i =2;i<sheet.actualRowCount;i++){
        let row = sheet.getRow(i);
        let cell:Cell = row.getCell(sheet.actualColumnCount);
        let rowIndex = cell.row, form ="";

        if(formula.indexOf('$row') >= 0)
          form = formula.replace(new RegExp('[$]row', 'g'), rowIndex);
        else form = formula

        cell.value = { formula: form, result: '?'};
        row.commit();
      }
    }
  }

  @rpsAction({verbName:'export-excel-to-csv'})
  exportToCsv(ctx:RpsContext,opts:Object, filename:string) : Promise<void>{
    let workbook:Workbook = this.getCurrentWorkbook(ctx);

    return workbook.csv.writeFile(filename)
  }



  private getCurrentWorkbook (ctx:RpsContext) :Workbook {
    let workbook:Workbook = ctx.getModuleContext(MOD_ID)['workbook'];
    return workbook;
  }

  @rpsAction({verbName:'get-worksheet'})
  async getWS (ctx:RpsContext,opts:{},sheetname:string|number) : Promise<Worksheet>{
    return this.getWorksheet(ctx,sheetname);
  }
  
  private getWorksheet(ctx:RpsContext,sheetname:string|number) : Worksheet{
    let workbook:Workbook = ctx.getModuleContext(MOD_ID)['workbook'];
    return workbook.getWorksheet(sheetname);
  }

}
