var assert = require('node:assert');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');

describe('ExcellentDomainModel', function() {
  'use strict';

  it('stores worksheet cells as Cell instances with explicit accessors', function() {
    var workbook = new excellentPackage.Workbook();
    var sheet = workbook.createSheet('Sheet1');
    var valueCell;
    var formulaCell;

    sheet.setCellValue('A1', 2);
    sheet.setCellFormula('A2', 'this.A1+1');

    valueCell = sheet.getCell('A1');
    formulaCell = sheet.getCell('A2');

    assert.ok(valueCell instanceof excellentPackage.Cell);
    assert.ok(formulaCell instanceof excellentPackage.Cell);
    assert.equal(valueCell.address, 'A1');
    assert.equal(valueCell.key, 'Sheet1!A1');
    assert.equal(formulaCell.kind, 'formula');
    assert.equal(valueCell.getRawValue(), 2);
    assert.equal(formulaCell.getRawValue(), undefined);
    assert.equal(formulaCell.getFormulaSource(), 'this.A1+1');
    assert.equal(formulaCell.getComputedValue(), 3);
    assert.equal(sheet.rows[0][0], valueCell);
    assert.equal(sheet.rows[1][0], formulaCell);
    assert.equal(workbook.getCell('Sheet1', 'A2'), formulaCell);
    assert.equal(sheet.getCellValue('A1'), 2);
    assert.equal(sheet.getCellValue('A2'), 3);
    assert.equal(sheet.A2, 3);
    assert.equal(sheet.getFormulaSource('A2'), 'this.A1+1');
    assert.equal(workbook.getFormulaSource('Sheet1', 'A2'), 'this.A1+1');
    assert.equal('_A2' in sheet, false);
  });

  it('keeps property access as convenience sugar while explicit methods stay primary', function() {
    var workbook = new excellentPackage.Workbook();
    var sheet = workbook.createSheet('Sheet1');

    sheet.setCellValue('A1', 2);
    sheet.setCellFormula('A2', 'this.A1+1');

    assert.equal(sheet.getCellValue('A2'), 3);
    assert.equal(sheet.A2, 3);

    sheet.A2 = 9;

    assert.equal(sheet.getCellValue('A2'), 9);
    assert.equal(sheet.getCell('A2').isFormula(), false);
    assert.equal(sheet.getFormulaSource('A2'), undefined);
  });

  it('keeps workbook mutation explicit instead of tracking an implicit current sheet', function() {
    var workbook = new excellentPackage.Workbook();
    var sheet = workbook.createSheet('Sheet1');

    sheet.setCellValue('A1', 5);
    workbook.setCellFormula('Sheet1', 'A2', 'this.A1+1');

    assert.equal('currentSheet' in workbook, false);
    assert.equal(typeof workbook.addCellVal, 'undefined');
    assert.equal(typeof workbook.addCellFunc, 'undefined');
    assert.equal(workbook.getSheet('Sheet1').A2, 6);
  });

  it('round-trips workbooks through WorkbookLoader modern methods', function() {
    var workbook = new excellentPackage.Workbook();
    var loader = new excellentPackage.WorkbookLoader();
    var sheet = workbook.createSheet('Sheet1');
    var serialized;
    var restored;

    sheet.setCellValue('A1', 4);
    sheet.setCellFormula('A2', 'this.A1+5');

    serialized = loader.serialize(workbook);
    restored = loader.deserialize(serialized);

    assert.deepEqual(serialized, {
      Sheet1: [
        [4],
        ['[function]this.A1+5']
      ]
    });
    assert.equal(restored.getSheet('Sheet1').A2, 9);
    assert.ok(restored.getSheet('Sheet1').getCell('A2') instanceof excellentPackage.Cell);
  });

  it('flushes formula caches once after batched workbook mutations', function() {
    var workbook = new excellentPackage.Workbook();
    var sheet = workbook.createSheet('Sheet1');

    sheet.setCellValue('A1', 2);
    sheet.setCellFormula('A2', 'this.A1+1');
    assert.equal(sheet.getCellValue('A2'), 3);

    workbook.beginMutationBatch();
    sheet.setCellValue('A1', 10);
    sheet.setCellFormula('A3', 'this.A2+5');
    workbook.endMutationBatch();

    assert.equal(sheet.getCellValue('A2'), 11);
    assert.equal(sheet.getCellValue('A3'), 16);
  });

  it('exposes a FormulaEvaluator for compile and reference collection', function() {
    var evaluator = new excellentPackage.FormulaEvaluator();
    var compiledFormula = evaluator.compile("Formula.SUM(this.A1,self.workbook['Other'].B2)");
    var references = evaluator.collectReferences(compiledFormula);

    assert.equal(compiledFormula.expression, "Formula.SUM(this.A1,self.workbook['Other'].B2)");
    assert.deepEqual(references, [
      { ref: 'A1', sheet: null },
      { ref: 'B2', sheet: 'Other' }
    ]);
    assert.equal(
      evaluator.serialize(compiledFormula),
      "Formula.SUM(this.A1,self.workbook['Other'].B2)"
    );
  });
});
