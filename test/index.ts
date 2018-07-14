import {expect} from 'chai';
import m from 'mocha';
import fs from 'fs';
import RPSModule from '../src/index';
import { RpsContext } from 'rpscript-interface';

m.describe('exceljs', () => {

  m.xit('should able to perform excel IO', async function () {
    let TEST_XLSX = 'test/sampledata.xlsx';
    let WRITE_XLSX = 'test/sampledata_result.xlsx';

    if(fs.existsSync(WRITE_XLSX)) fs.unlinkSync(WRITE_XLSX);
    fs.copyFileSync('test/sampledata_original.xlsx',TEST_XLSX);
    
    let ctx =new RpsContext;
    let md = new RPSModule(ctx);

    //LOAD WORKSHEET
    let workbook = await md.readExcel(ctx,{},TEST_XLSX);

    //remove worksheets
    await md.removeWorksheet(ctx,{},'SampleNumbers');
    await md.removeWorksheet(ctx,{},'MyLinks');
    await md.removeWorksheet(ctx,{},'Instructions');
    
    //check if sheet removed
    let instruction = await md.getWS(ctx,{},'Instructions');
    let salesorders = await md.getWS(ctx,{},'SalesOrders');
    expect(instruction).to.be.undefined;
    expect(salesorders).not.to.be.undefined;

    //copy sheets
    let orders = await md.dupWorksheet(ctx,{},'SalesOrders','Orders');
    expect(orders).not.to.be.undefined;

    let row = await md.getRow(ctx,{},'SalesOrders',1);
    expect(row.values).to.be.deep.equals(
      [undefined, 'OrderDate','Region','Rep','Item','Units','Unit Cost','Total' ]);

    let cols = await md.getColumn(ctx,{},'SalesOrders',7);
    //@ts-ignore
    await md.appendColumn(ctx,{},'Orders','DupTotal',cols.values);
    
    //write to result
    await workbook.xlsx.writeFile(WRITE_XLSX);

    //LOAD WORKSHEET
    let resultWB = await md.readExcel(ctx,{},WRITE_XLSX);
    // let resultorders = await md.getWS(ctx,{},'Orders');

    // expect(salesorders.getSheetValues()).to.be.deep.equals(resultorders.getSheetValues());

    let cell = await md.getCell(ctx,{},'Orders','H3');
    
    // cell.value = { formula: 'H5+H4', result: 1 };
    cell.value = 1;

    await workbook.xlsx.writeFile(WRITE_XLSX);

  }).timeout(0);

  m.it('should append excel', async function () {
    let TEST_XLSX = 'test/sampledata.xlsx';
    let WRITE_XLSX = 'test/sampledata_result.xlsx';

    if(fs.existsSync(WRITE_XLSX)) fs.unlinkSync(WRITE_XLSX);
    
    let ctx =new RpsContext;
    let md = new RPSModule(ctx);

    //LOAD WORKSHEET
    let workbook = await md.readExcel(ctx,{},TEST_XLSX);

    // let salesorders = await md.getWS(ctx,{},'SalesOrders');
    // expect(salesorders).not.to.be.undefined;


    // await md.appendColumn(ctx,{},'SalesOrders','WonderFn',async function (row,count) {
    //   return count;
    // });

    // await md.appendColumn(ctx,{formula:'A1'},'SalesOrders','WonderFormula');
    await md.appendColumn(ctx,{
      fontColor:'FF112266',bgColor:'FF11DD22',fillType:'solid',
      formula:'IF(B$row="East","East Side","No Side")'},'SalesOrders','Which Side');

    await md.setCell(ctx,{fontColor:'FF11AA11'},'SalesOrders','C1');
    await md.setCell(ctx,{
      bgColor:'99111122',fillType:'pattern',fillPattern:'lightGray'
    },'SalesOrders','D1');
    // let ws = await md.addWorksheet(ctx,{},'NewSheet');
    
    // await md.addHeader(ctx,{},'NewSheet','ID','id');
    // await md.addHeader(ctx,{},'NewSheet','Unit No');
    // await md.addHeader(ctx,{},'NewSheet','Desc','description');
    // await md.addHeader(ctx,{},'NewSheet','Weird','weird');
    // await md.addHeaders(ctx,{},'NewSheet',['id','name','description']);

    // await md.addRow(ctx,{},'NewSheet',[1,'hello','world']);
    // await md.addRow(ctx,{},'NewSheet',{id:1,name:'hello',weird:'world'});
    // await md.addRows(ctx,{},'NewSheet',[
    //   {id:2,'Unit No':'hello2',description:'world2'},
    //   {id:3,name:'hello3',description:'world3'}]);

    //write to result
    await workbook.xlsx.writeFile(WRITE_XLSX);


  }).timeout(0);

})
