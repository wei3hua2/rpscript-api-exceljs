import {expect} from 'chai';
import m from 'mocha';
import fs from 'fs';
import RPSModule from '../src/index';
import { RpsContext } from 'rpscript-interface';

m.describe('exceljs', () => {

  m.it('should able to perform excel IO', async function () {
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

})
