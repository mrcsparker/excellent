var assert = require('node:assert');
var fs = require('node:fs');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');

var Workbook = excellentPackage.Workbook;
var WorkbookLoader = excellentPackage.WorkbookLoader;
var XlsxReader = excellentPackage.XlsxReader;
var XLSX_READER_MODE = excellentPackage.XLSX_READER_MODE;

describe('ExcellentJavaScriptConsumer', function() {
  'use strict';

  it('supports the documented CommonJS compatibility workflow', async function() {
    var workbook = new Workbook();
    var sheet = workbook.createSheet('Sheet1');
    var loader = new WorkbookLoader();
    var xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx');
    var parsed;
    var serialized;

    sheet.setCellValue('A1', 2);
    sheet.setCellFormula('A2', 'this.A1+1');

    serialized = loader.serialize(workbook);
    parsed = await new XlsxReader({
      mode: XLSX_READER_MODE.FORMULAS
    }).load(xlsxFile);

    assert.deepEqual(serialized, {
      Sheet1: [
        [2],
        ['[function]this.A1+1']
      ]
    });
    assert.equal(sheet.getCellValue('A2'), 3);
    assert.equal(workbook.getCellValue('Sheet1', 'A2'), 3);
    assert.equal(parsed.getCellValue('Sheet1', 'A1'), 3);
  });
});
