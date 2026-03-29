var assert = require('node:assert');
var fs = require('node:fs');
var test = require('node:test');
var before = test.before;
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');
var XlsxReader = excellentPackage.XlsxReader;
var formulaAssertions = require('../helpers/formula_assertions.js');
var assertFormulaSource = formulaAssertions.assertFormulaSource;

describe('ExcellentXlsxFormulaWorkbooks', function() {
  'use strict';

  it('loads shared formulas from xlsx files and invalidates them after mutation', async function() {
    var excellent = new XlsxReader();
    var parsed = await excellent.load(fs.readFileSync('./test/data/sharedFormulas.xlsx'));
    var sheet = parsed.workbook.Shared;

    assertFormulaSource(sheet, 'B1', 'this.A1+1');
    assertFormulaSource(sheet, 'B2', 'this.A2+1');
    assertFormulaSource(sheet, 'B3', 'this.A3+1');
    assertFormulaSource(sheet, 'C1', 'Formula.SUM([this.B1,this.B2,this.B3])');
    assert.equal(sheet.B1, 2);
    assert.equal(sheet.B2, 6);
    assert.equal(sheet.B3, 10);
    assert.equal(sheet.C1, 18);

    sheet.A2 = 8;
    assert.equal(sheet.B2, 9);
    assert.equal(sheet.C1, 21);
  });

  it('loads cross-sheet formulas from xlsx files and recomputes after mutation', async function() {
    var excellent = new XlsxReader();
    var parsed = await excellent.load(fs.readFileSync('./test/data/crossSheetWorkbook.xlsx'));

    assert.equal(parsed.workbook.Outputs.A1, 5);
    assert.equal(parsed.workbook.Outputs.A2, 6);
    assert.equal(parsed.workbook.Outputs.B1, 10);
    assertFormulaSource(parsed.workbook.Outputs, 'A1', "self.workbook['Inputs'].A1+1");
    assertFormulaSource(parsed.workbook.Outputs, 'B1', "Formula.SUM(self.workbook['Inputs'].A1,this.A2)");

    parsed.workbook.Inputs.A1 = 10;
    assert.equal(parsed.workbook.Outputs.A1, 11);
    assert.equal(parsed.workbook.Outputs.B1, 16);
  });

  it('loads quoted sheet names, absolute references, escaped strings, and error literals from xlsx files', async function() {
    var excellent = new XlsxReader();
    var parsed = await excellent.load(fs.readFileSync('./test/data/quotedSheetAndErrors.xlsx'));
    var summary = parsed.workbook.Summary;

    assert.equal(summary.A1, 8);
    assert.equal(summary.A3, 10);
    assert.equal(summary.A4, 1);
    assert.equal(summary.A5, 99);
    assert.equal(summary.A6, 77);
    assertFormulaSource(summary, 'A1', "self.workbook['Budget 2026'].A1+1");
    assertFormulaSource(summary, 'A3', "Formula.SUM(this.A2,self.workbook['Budget 2026'].A2)");
    assertFormulaSource(summary, 'A4', 'Formula.IF("He said \\"hi\\""=="He said \\"hi\\"",1,0)');
    assertFormulaSource(summary, 'A5', 'Formula.IFERROR(#DIV/0!,99)');
    assertFormulaSource(summary, 'A6', 'Formula.IFNA(#N/A,77)');
  });

  describe('Simple Formulas', function() {
    var content;

    before(async function() {
      var xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx');
      var excellent = new XlsxReader();
      var parsed = await excellent.load(xlsxFile);

      content = parsed.workbook.Sheet1;
    });

    it('SUM(1,2)', function() {
      assertFormulaSource(content, 'A1', 'Formula.SUM(1,2)');
      assert.equal(content.A1, 3);
    });

    it('AVERAGE(A1)', function() {
      assertFormulaSource(content, 'A2', 'Formula.AVERAGE(this.A1)');
      assert.equal(content.A2, 3);
    });

    it('COUNT(A2)', function() {
      assertFormulaSource(content, 'A3', 'Formula.COUNT(this.A2)');
      assert.equal(content.A3, 1);
    });

    it('POWER(10,2)', function() {
      assertFormulaSource(content, 'A4', 'Formula.POWER(10,2)');
      assert.equal(content.A4, 100);
    });

    it('SQRT(100)', function() {
      assertFormulaSource(content, 'A5', 'Formula.SQRT(100)');
      assert.equal(content.A5, 10);
    });

    it('FACT(5)', function() {
      assertFormulaSource(content, 'A6', 'Formula.FACT(5)');
      assert.equal(content.A6, 120);
    });

    it('FLOOR(100,2)', function() {
      assertFormulaSource(content, 'A7', 'Formula.FLOOR(100,2)');
      assert.equal(content.A7, 100);
    });

    it('ABS(-1)', function() {
      assertFormulaSource(content, 'A8', 'Formula.ABS(-1)');
      assert.equal(content.A8, 1);
    });

    it('MOD(55,3)', function() {
      assertFormulaSource(content, 'A9', 'Formula.MOD(55,3)');
      assert.equal(content.A9, 1);
    });
  });

  describe('Simple Ranges', function() {
    var content;

    before(async function() {
      var xlsxFile = fs.readFileSync('./test/data/simpleRange.xlsx');
      var excellent = new XlsxReader();
      var parsed = await excellent.load(xlsxFile);

      content = parsed.workbook.Sheet1;
    });

    it('SUM(A1:H1)', function() {
      assertFormulaSource(content, 'A20', "Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1])");
      assert.equal(content.A20, 36);
    });

    it('SUM(B1:H3)', function() {
      assertFormulaSource(content, 'A21', "Formula.SUM([this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])");
      assert.equal(content.A21, 273);
    });

    it('SUM(I3:L3)', function() {
      assertFormulaSource(content, 'A22', "Formula.SUM([this.I3,this.J3,this.K3,this.L3])");
      assert.equal(content.A22, 106);
    });

    it('SUM(A1:H3)', function() {
      assertFormulaSource(content, 'A23', "Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])");
      assert.equal(content.A23, 300);
    });

    it('AVERAGE(A1:H3)', function() {
      assertFormulaSource(content, 'A24', "Formula.AVERAGE([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])");
      assert.equal(content.A24, 12.5);
    });

    it('SUM(B1:B16)', function() {
      assertFormulaSource(content, 'A25', "Formula.SUM([this.B1,this.B2,this.B3,this.B4,this.B5,this.B6,this.B7,this.B8,this.B9,this.B10,this.B11,this.B12,this.B13,this.B14,this.B15,this.B16])");
      assert.equal(content.A25, 134);
    });
  });
});
