var assert = require('node:assert');
var childProcess = require('node:child_process');
var test = require('node:test');
var describe = test.describe;
var it = test.it;

describe('ExcellentEsmConsumer', function() {
  'use strict';

  it('supports the documented native ESM import workflow', function() {
    var output = childProcess.execFileSync(process.execPath, [
      '--input-type=module',
      '-e',
      [
        "import { readFileSync } from 'node:fs';",
        "import { Workbook, XlsxReader } from 'excellent';",
        "import { Workbook as BrowserWorkbook, XlsxReader as BrowserXlsxReader } from 'excellent/browser';",
        "const workbook = new Workbook();",
        "const sheet = workbook.createSheet('Sheet1');",
        "const xlsxFile = readFileSync('./test/data/simpleFormula.xlsx');",
        "sheet.setCellValue('A1', 5);",
        "sheet.setCellFormula('A2', 'this.A1+2');",
        "const parsed = await new XlsxReader().load(xlsxFile);",
        "console.log(JSON.stringify({",
        "  computed: sheet.getCellValue('A2'),",
        "  parsedValue: parsed.getCellValue('Sheet1', 'A1'),",
        "  browserWorkbookMatches: BrowserWorkbook === Workbook,",
        "  browserReaderMatches: BrowserXlsxReader === XlsxReader",
        "}));"
      ].join('\n')
    ], {
      cwd: process.cwd(),
      encoding: 'utf8'
    });
    var result = JSON.parse(output.trim());

    assert.equal(result.computed, 7);
    assert.equal(result.parsedValue, 3);
    assert.equal(result.browserWorkbookMatches, true);
    assert.equal(result.browserReaderMatches, true);
  });
});
