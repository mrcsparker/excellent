var assert = require('node:assert');
var fs = require('node:fs');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');

describe('ExcellentPackageEntrypoint', function() {
  'use strict';

  it('loads the public package entrypoint', function() {
    assert.equal(typeof excellentPackage.Cell, 'function');
    assert.ok(excellentPackage.ExcelError);
    assert.equal(typeof excellentPackage.XlsxReader, 'function');
    assert.equal(typeof excellentPackage.WorkbookLoader, 'function');
    assert.equal(typeof excellentPackage.ExcelErrorValue, 'function');
    assert.equal(typeof excellentPackage.FormulaEvaluator, 'function');
    assert.equal(typeof excellentPackage.FormulaFunctionRegistry, 'function');
    assert.equal(typeof excellentPackage.isExcelError, 'function');
    assert.equal(typeof excellentPackage.Workbook, 'function');
    assert.ok(excellentPackage.XLSX_READER_MODE);
    assert.equal(typeof excellentPackage.Excellent, 'undefined');
    assert.equal(typeof excellentPackage.ExcellentLoader, 'undefined');
    assert.equal(typeof excellentPackage.Xlsx, 'undefined');
    assert.equal(typeof excellentPackage.Loader, 'undefined');
    assert.equal(typeof excellentPackage.XlsxSimple, 'undefined');
  });

  it('parses an xlsx file through the package entrypoint', async function() {
    var xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx');
    var reader = new excellentPackage.XlsxReader();
    var parsed = await reader.load(xlsxFile);

    assert.equal(parsed.workbook.Sheet1.A1, 3);
    assert.equal(parsed.workbook.Sheet1.getFormulaSource('A1'), 'Formula.SUM(1,2)');
    assert.equal('_A1' in parsed.workbook.Sheet1, false);
  });
});
