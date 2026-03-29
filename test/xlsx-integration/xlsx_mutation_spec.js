var assert = require('node:assert');
var fs = require('node:fs');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');
var formulaAssertions = require('../helpers/formula_assertions.js');
var assertFormulaSource = formulaAssertions.assertFormulaSource;

describe('ExcellentXlsxMutation', function() {
  'use strict';

  it('recomputes loaded shared-formula chains after repeated input updates', async function() {
    var workbook = await new excellentPackage.XlsxReader().load(fs.readFileSync('./test/data/sharedFormulas.xlsx'));
    var sheet = workbook.getSheet('Shared');

    assert.equal(workbook.getCellValue('Shared', 'B1'), 2);
    assert.equal(workbook.getCellValue('Shared', 'B2'), 6);
    assert.equal(workbook.getCellValue('Shared', 'B3'), 10);
    assert.equal(workbook.getCellValue('Shared', 'C1'), 18);

    workbook.setCellValue('Shared', 'A1', 100);

    assert.equal(workbook.getCellValue('Shared', 'B1'), 101);
    assert.equal(workbook.getCellValue('Shared', 'B2'), 6);
    assert.equal(workbook.getCellValue('Shared', 'B3'), 10);
    assert.equal(workbook.getCellValue('Shared', 'C1'), 117);

    sheet.setCellValue('A3', -1);

    assert.equal(workbook.getCellValue('Shared', 'B1'), 101);
    assert.equal(workbook.getCellValue('Shared', 'B2'), 6);
    assert.equal(workbook.getCellValue('Shared', 'B3'), 0);
    assert.equal(workbook.getCellValue('Shared', 'C1'), 107);
    assert.deepEqual(sheet.functions, ['B1', 'B2', 'B3', 'C1']);
    assert.deepEqual(sheet.variables, ['A1', 'A2', 'A3']);
    assertFormulaSource(sheet, 'B1', 'this.A1+1');
    assertFormulaSource(sheet, 'B2', 'this.A2+1');
    assertFormulaSource(sheet, 'B3', 'this.A3+1');
    assertFormulaSource(sheet, 'C1', 'Formula.SUM([this.B1,this.B2,this.B3])');
  });

  it('recomputes loaded cross-sheet dependency chains across multiple input updates', async function() {
    var workbook = await new excellentPackage.XlsxReader().load(fs.readFileSync('./test/data/crossSheetWorkbook.xlsx'));
    var inputs = workbook.getSheet('Inputs');
    var outputs = workbook.getSheet('Outputs');
    var trace;

    assert.equal(workbook.getCellValue('Outputs', 'A1'), 5);
    assert.equal(workbook.getCellValue('Outputs', 'A2'), 6);
    assert.equal(workbook.getCellValue('Outputs', 'B1'), 10);

    workbook.setCellValue('Inputs', 'A2', 20);

    assert.equal(workbook.getCellValue('Outputs', 'A1'), 5);
    assert.equal(workbook.getCellValue('Outputs', 'A2'), 21);
    assert.equal(workbook.getCellValue('Outputs', 'B1'), 25);

    inputs.setCellValue('A1', 1);

    assert.equal(workbook.getCellValue('Outputs', 'A1'), 2);
    assert.equal(workbook.getCellValue('Outputs', 'A2'), 21);
    assert.equal(workbook.getCellValue('Outputs', 'B1'), 22);
    assert.deepEqual(inputs.variables, ['A1', 'A2']);
    assert.deepEqual(outputs.functions, ['A1', 'B1', 'A2']);
    assertFormulaSource(outputs, 'A1', "self.workbook['Inputs'].A1+1");
    assertFormulaSource(outputs, 'A2', "self.workbook['Inputs'].A2+1");
    assertFormulaSource(outputs, 'B1', "Formula.SUM(self.workbook['Inputs'].A1,this.A2)");

    trace = workbook.traceCell('Outputs', 'B1');
    assert.equal(trace.value, 22);
    assert.deepEqual(trace.precedents, [
      {
        cellName: 'A1',
        key: 'Inputs!A1',
        sheetName: 'Inputs'
      },
      {
        cellName: 'A2',
        key: 'Outputs!A2',
        sheetName: 'Outputs'
      }
    ]);
  });

  it('recomputes loaded range formulas when source cells change', async function() {
    var workbook = await new excellentPackage.XlsxReader().load(fs.readFileSync('./test/data/simpleRange.xlsx'));
    var sheet = workbook.getSheet('Sheet1');

    assert.equal(workbook.getCellValue('Sheet1', 'A20'), 36);
    assert.equal(workbook.getCellValue('Sheet1', 'A21'), 273);
    assert.equal(workbook.getCellValue('Sheet1', 'A23'), 300);
    assert.equal(workbook.getCellValue('Sheet1', 'A24'), 12.5);
    assert.equal(workbook.getCellValue('Sheet1', 'A25'), 134);

    workbook.setCellValue('Sheet1', 'A1', 100);

    assert.equal(workbook.getCellValue('Sheet1', 'A20'), 135);
    assert.equal(workbook.getCellValue('Sheet1', 'A21'), 273);
    assert.equal(workbook.getCellValue('Sheet1', 'A23'), 399);
    assert.equal(workbook.getCellValue('Sheet1', 'A24'), 16.625);
    assert.equal(workbook.getCellValue('Sheet1', 'A25'), 134);

    sheet.setCellValue('B4', 20);

    assert.equal(workbook.getCellValue('Sheet1', 'A20'), 135);
    assert.equal(workbook.getCellValue('Sheet1', 'A21'), 273);
    assert.equal(workbook.getCellValue('Sheet1', 'A23'), 399);
    assert.equal(workbook.getCellValue('Sheet1', 'A24'), 16.625);
    assert.equal(workbook.getCellValue('Sheet1', 'A25'), 152);
    assertFormulaSource(sheet, 'A20', 'Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1])');
    assertFormulaSource(sheet, 'A21', 'Formula.SUM([this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])');
    assertFormulaSource(sheet, 'A23', 'Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])');
    assertFormulaSource(sheet, 'A24', 'Formula.AVERAGE([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])');
    assertFormulaSource(sheet, 'A25', 'Formula.SUM([this.B1,this.B2,this.B3,this.B4,this.B5,this.B6,this.B7,this.B8,this.B9,this.B10,this.B11,this.B12,this.B13,this.B14,this.B15,this.B16])');
  });

  it('keeps quoted-sheet references and error formulas stable after loaded input updates', async function() {
    var workbook = await new excellentPackage.XlsxReader().load(fs.readFileSync('./test/data/quotedSheetAndErrors.xlsx'));
    var budget = workbook.getSheet('Budget 2026');
    var summary = workbook.getSheet('Summary');

    assert.equal(workbook.getCellValue('Summary', 'A1'), 8);
    assert.equal(workbook.getCellValue('Summary', 'A3'), 10);
    assert.equal(workbook.getCellValue('Summary', 'A5'), 99);
    assert.equal(workbook.getCellValue('Summary', 'A6'), 77);

    summary.setCellValue('A2', 10);

    assert.equal(workbook.getCellValue('Summary', 'A1'), 8);
    assert.equal(workbook.getCellValue('Summary', 'A3'), 18);
    assert.equal(workbook.getCellValue('Summary', 'A5'), 99);
    assert.equal(workbook.getCellValue('Summary', 'A6'), 77);

    workbook.setCellValue('Budget 2026', 'A1', 50);
    budget.setCellValue('A2', 1);

    assert.equal(workbook.getCellValue('Summary', 'A1'), 51);
    assert.equal(workbook.getCellValue('Summary', 'A3'), 11);
    assert.equal(workbook.getCellValue('Summary', 'A5'), 99);
    assert.equal(workbook.getCellValue('Summary', 'A6'), 77);
    assert.deepEqual(summary.functions, ['A1', 'A3', 'A4', 'A5', 'A6']);
    assert.deepEqual(summary.variables, ['A2']);
    assertFormulaSource(summary, 'A1', "self.workbook['Budget 2026'].A1+1");
    assertFormulaSource(summary, 'A3', "Formula.SUM(this.A2,self.workbook['Budget 2026'].A2)");
    assertFormulaSource(summary, 'A5', 'Formula.IFERROR(#DIV/0!,99)');
    assertFormulaSource(summary, 'A6', 'Formula.IFNA(#N/A,77)');
  });
});
