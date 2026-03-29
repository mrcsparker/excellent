var assert = require('node:assert');
var fs = require('node:fs');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');

describe('ExcellentXlsxReader', function() {
  'use strict';

  function createProfileCollector() {
    return {
      counts: Object.create(null),
      durationsMs: Object.create(null),
      incrementCount(label, amount) {
        var incrementBy = amount === undefined ? 1 : amount;

        this.counts[label] = (this.counts[label] || 0) + incrementBy;
      },
      async measureAsync(label, callback) {
        var startTime = Date.now();

        try {
          return await callback();
        } finally {
          this.durationsMs[label] = (this.durationsMs[label] || 0) + (Date.now() - startTime);
        }
      },
      measureSync(label, callback) {
        var startTime = Date.now();

        try {
          return callback();
        } finally {
          this.durationsMs[label] = (this.durationsMs[label] || 0) + (Date.now() - startTime);
        }
      }
    };
  }

  it('loads formula workbooks through the unified reader in formula mode by default', async function() {
    var xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx');
    var reader = new excellentPackage.XlsxReader();
    var parsed = await reader.load(xlsxFile);
    var sheet = parsed.getSheet('Sheet1');

    assert.equal(sheet.A1, 3);
    assert.equal(sheet.getFormulaSource('A1'), 'Formula.SUM(1,2)');
    assert.equal('_A1' in sheet, false);
    assert.equal(sheet.functions.length, 9);
    assert.equal(sheet.variables.length, 0);
  });

  it('loads cached worksheet values through the same reader in values-only mode', async function() {
    var xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx');
    var reader = new excellentPackage.XlsxReader({
      mode: excellentPackage.XLSX_READER_MODE.VALUES_ONLY
    });
    var parsed = await reader.load(xlsxFile);
    var sheet = parsed.getSheet('Sheet1');

    assert.equal(sheet.A1, 3);
    assert.equal(sheet.getFormulaSource('A1'), undefined);
    assert.equal(sheet.functions.length, 0);
    assert.equal(sheet.variables.length, 9);
    assert.equal(sheet.getCell('A1').kind, 'value');
  });

  it('supports values-only mode for cross-sheet workbooks through XlsxReader options', async function() {
    var xlsxFile = fs.readFileSync('./test/data/crossSheetWorkbook.xlsx');
    var reader = new excellentPackage.XlsxReader({
      mode: excellentPackage.XLSX_READER_MODE.VALUES_ONLY
    });
    var workbook = await reader.load(xlsxFile);
    var inputs = workbook.getSheet('Inputs');
    var outputs = workbook.getSheet('Outputs');

    assert.deepEqual(workbook.getSheetNames(), ['Inputs', 'Outputs']);
    assert.equal(inputs.A1, 4);
    assert.equal(inputs.A2, 5);
    assert.equal(outputs.A1, 5);
    assert.equal(outputs.functions.length, 0);
    assert.equal(outputs.getFormulaSource('A1'), undefined);
  });

  it('loads sheets incrementally and exposes the evolving workbook after each sheet', async function() {
    var callbackWorkbook;
    var snapshots = [];
    var xlsxFile = fs.readFileSync('./test/data/crossSheetWorkbook.xlsx');
    var reader = new excellentPackage.XlsxReader();
    var workbook = await reader.loadIncremental(xlsxFile, async function(event) {
      var maybeOutput = event.workbook.getCellValue('Outputs', 'A1');

      snapshots.push({
        outputValue: excellentPackage.isExcelError(maybeOutput) ? maybeOutput.code : maybeOutput,
        sheetIndex: event.sheetIndex,
        sheetName: event.sheetName,
        sheetNames: event.workbook.getSheetNames().slice(),
        worksheetA1: event.worksheet.getCellValue('A1')
      });
      callbackWorkbook = event.workbook;
      await Promise.resolve();
    });

    assert.equal(workbook, callbackWorkbook);
    assert.deepEqual(snapshots, [
      {
        outputValue: '#REF!',
        sheetIndex: 0,
        sheetName: 'Inputs',
        sheetNames: ['Inputs'],
        worksheetA1: 4
      },
      {
        outputValue: 5,
        sheetIndex: 1,
        sheetName: 'Outputs',
        sheetNames: ['Inputs', 'Outputs'],
        worksheetA1: 5
      }
    ]);
  });

  it('reports load timing phases and counts through an optional profiler', async function() {
    var profile = createProfileCollector();
    var xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx');
    var reader = new excellentPackage.XlsxReader({ profile: profile });
    var workbook = await reader.load(xlsxFile);

    assert.equal(workbook.getCellValue('Sheet1', 'A1'), 3);
    assert.equal(profile.counts['xlsx.loads'], 1);
    assert.equal(profile.counts['workbook.sheetsLoaded'], 1);
    assert.ok(profile.counts['worksheet.rows'] >= 1);
    assert.ok(profile.counts['worksheet.cells.formula'] >= 1);
    assert.ok(profile.durationsMs['xlsx.loadZip'] >= 0);
    assert.ok(profile.durationsMs['workbook.xmlToJson'] >= 0);
    assert.ok(profile.durationsMs['worksheet.xmlToJson'] >= 0);
    assert.ok(profile.durationsMs['worksheet.compileFormula'] >= 0);
  });

  it('rejects unknown reader modes', function() {
    assert.throws(function() {
      return new excellentPackage.XlsxReader({ mode: 'mystery-mode' });
    }, {
      message: 'Unsupported XLSX reader mode: mystery-mode',
      name: 'TypeError'
    });
  });
});
