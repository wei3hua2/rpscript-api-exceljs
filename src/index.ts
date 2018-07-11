/**
 * @module exceljs
 */

import ExcelJS,{Workbook,Worksheet, AutoFilter,Row,Column, ColumnExtension, Cell} from 'exceljs';
import {RpsContext,RpsModule,rpsAction} from 'rpscript-interface';

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

  // @rpsAction({verbName:'excel-commit'})
  // async commit(ctx:RpsContext,opts:Object) : Promise<void>{
  //   let sheet = this.getWorksheet(ctx,sheetname);
  //   return sheet.addRows(datas);
  // }

  @rpsAction({verbName:'excel-auto-filter'})
  async  autofilter(ctx:RpsContext,opts:Object, sheetname:string|number,
    param1:any, param2?:any, param3?:any, param4?:any) : Promise<AutoFilter>{
    let sheet = this.getWorksheet(ctx,sheetname);
    
    if(!param2) sheet.autoFilter = param1;
    else if (!param3) sheet.autoFilter = {from:param1,to:param2};
    else sheet.autoFilter = { from :{row:param1,column:param2} , to:{row:param3,column:param4} };
    
    return sheet.autoFilter;
  }

  @rpsAction({verbName:'worksheet-get-headers'})
  async addHeader(ctx:RpsContext,opts:Object, sheetname:string|number, 
    header:string, key?:string, width?:number,outlineLevel?:number) : Promise<any>{
      let sheet = this.getWorksheet(ctx,sheetname);

      sheet.columns.push({header:header,key:key,width:width,outlineLevel:outlineLevel});
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
  @rpsAction({verbName:'worksheet-append-column'})
  async appendColumn(ctx:RpsContext,opts:Object, 
    sheetname:string|number, columnName:string, data:Array<any>) : Promise<void>{

    let sheet = this.getWorksheet(ctx,sheetname);
    
    sheet.spliceColumns(sheet.actualColumnCount+1,0,data);

    sheet.commit();
  }

  @rpsAction({verbName:'export-excel-to-csv'})
  exportToCsv(ctx:RpsContext,opts:Object, filename:string) : Promise<void>{
    let workbook:Workbook = this.getCurrentWorkbook(ctx);

    return workbook.csv.writeFile(filename)
  }

  //get column
  //each cell
  // splice column
  //get cell

  //add row
  //get row
  //last row
  //splice row
  
  //merge cells
  //unmerge cells


  //commit
  //add page break
  //add image


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