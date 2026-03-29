'use strict';

const process = require('node:process');
const { performance } = require('node:perf_hooks');
const JSZip = require('jszip');
const { Workbook, XlsxReader } = require('../dist');

const DEFAULT_OPTIONS = Object.freeze({
  iterations: 5,
  rows: 2000,
  sharedRows: 4000,
  warmups: 1
});

function writeStdout(message) {
  process.stdout.write(String(message) + '\n');
}

function writeStderr(message) {
  process.stderr.write(String(message) + '\n');
}

function printUsage() {
  writeStdout('Usage: npm run bench:evaluation -- [--rows <count>] [--shared-rows <count>] [--iterations <count>] [--warmups <count>]');
  writeStdout('');
  writeStdout('Defaults:');
  writeStdout('  --rows ' + String(DEFAULT_OPTIONS.rows));
  writeStdout('  --shared-rows ' + String(DEFAULT_OPTIONS.sharedRows));
  writeStdout('  --iterations ' + String(DEFAULT_OPTIONS.iterations));
  writeStdout('  --warmups ' + String(DEFAULT_OPTIONS.warmups));
}

function parsePositiveInteger(rawValue, flagName) {
  const parsedValue = Number.parseInt(rawValue, 10);

  if (!Number.isInteger(parsedValue) || parsedValue <= 0) {
    throw new Error('Expected ' + flagName + ' to be a positive integer, received "' + rawValue + '".');
  }

  return parsedValue;
}

function readFlagValue(argv, index, flagName) {
  const nextValue = argv[index + 1];

  if (nextValue === undefined) {
    throw new Error('Missing value for ' + flagName + '.');
  }

  return nextValue;
}

function parseArgs(argv) {
  const options = {
    iterations: DEFAULT_OPTIONS.iterations,
    rows: DEFAULT_OPTIONS.rows,
    sharedRows: DEFAULT_OPTIONS.sharedRows,
    warmups: DEFAULT_OPTIONS.warmups
  };

  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];

    if (argument === '--help' || argument === '-h') {
      printUsage();
      process.exit(0);
    }

    if (argument === '--rows') {
      options.rows = parsePositiveInteger(readFlagValue(argv, index, '--rows'), '--rows');
      index += 1;
      continue;
    }

    if (argument.startsWith('--rows=')) {
      options.rows = parsePositiveInteger(argument.slice('--rows='.length), '--rows');
      continue;
    }

    if (argument === '--shared-rows') {
      options.sharedRows = parsePositiveInteger(readFlagValue(argv, index, '--shared-rows'), '--shared-rows');
      index += 1;
      continue;
    }

    if (argument.startsWith('--shared-rows=')) {
      options.sharedRows = parsePositiveInteger(argument.slice('--shared-rows='.length), '--shared-rows');
      continue;
    }

    if (argument === '--iterations') {
      options.iterations = parsePositiveInteger(readFlagValue(argv, index, '--iterations'), '--iterations');
      index += 1;
      continue;
    }

    if (argument.startsWith('--iterations=')) {
      options.iterations = parsePositiveInteger(argument.slice('--iterations='.length), '--iterations');
      continue;
    }

    if (argument === '--warmups') {
      options.warmups = parsePositiveInteger(readFlagValue(argv, index, '--warmups'), '--warmups');
      index += 1;
      continue;
    }

    if (argument.startsWith('--warmups=')) {
      options.warmups = parsePositiveInteger(argument.slice('--warmups='.length), '--warmups');
      continue;
    }

    throw new Error('Unknown argument: ' + argument);
  }

  return options;
}

function createContentTypesXml(sheetCount) {
  const worksheetOverrides = Array.from({ length: sheetCount }, function(_value, index) {
    return '<Override PartName="/xl/worksheets/sheet' + String(index + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
  }).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    worksheetOverrides,
    '</Types>'
  ].join('');
}

function createRootRelationshipsXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>'
  ].join('');
}

function createWorkbookRelationshipsXml(sheetCount) {
  const sheetRelationships = Array.from({ length: sheetCount }, function(_value, index) {
    return '<Relationship Id="rId' + String(index + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + String(index + 1) + '.xml"/>';
  }).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    sheetRelationships,
    '</Relationships>'
  ].join('');
}

function createWorkbookXml(sheetNames) {
  const sheetXml = sheetNames.map(function(sheetName, index) {
    return '<sheet name="' + sheetName + '" sheetId="' + String(index + 1) + '" r:id="rId' + String(index + 1) + '"/>';
  }).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="24026"/>',
    '<sheets>',
    sheetXml,
    '</sheets>',
    '</workbook>'
  ].join('');
}

function createWorksheetXml(rows) {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<sheetData>',
    rows.join(''),
    '</sheetData>',
    '</worksheet>'
  ].join('');
}

function createSharedFormulaWorksheetXml(rowCount) {
  const rows = [];

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    const valueCell = '<c r="A' + String(rowNumber) + '"><v>' + String(rowNumber) + '</v></c>';
    let formulaCell = '<c r="B' + String(rowNumber) + '"><f t="shared" si="0"/><v>' + String(rowNumber * 2) + '</v></c>';

    if (rowNumber === 1) {
      formulaCell = '<c r="B1"><f t="shared" ref="B1:B' + String(rowCount) + '" si="0">A1*2</f><v>2</v></c>';
    }

    rows.push('<row r="' + String(rowNumber) + '" spans="1:2">' + valueCell + formulaCell + '</row>');
  }

  return createWorksheetXml(rows);
}

async function createWorkbookBuffer(sheetName, sheetXml) {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', createContentTypesXml(1));
  zip.file('_rels/.rels', createRootRelationshipsXml());
  zip.file('xl/workbook.xml', createWorkbookXml([sheetName]));
  zip.file('xl/_rels/workbook.xml.rels', createWorkbookRelationshipsXml(1));
  zip.file('xl/worksheets/sheet1.xml', sheetXml);

  return await zip.generateAsync({
    compression: 'DEFLATE',
    type: 'nodebuffer'
  });
}

async function createSharedFormulaWorkbookBuffer(rowCount) {
  return await createWorkbookBuffer('Shared', createSharedFormulaWorksheetXml(rowCount));
}

function createLargeSheetWorkbook(rowCount) {
  const workbook = new Workbook();
  const sheet = workbook.createSheet('Model');

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    const addressA = 'A' + String(rowNumber);
    const addressB = 'B' + String(rowNumber);
    const addressC = 'C' + String(rowNumber);
    const addressD = 'D' + String(rowNumber);

    sheet.setCellValue(addressA, rowNumber);
    sheet.setCellFormula(addressB, 'this.' + addressA + '*2');
    sheet.setCellFormula(addressC, 'this.' + addressB + '+1');
    sheet.setCellFormula(addressD, 'this.' + addressC + '+this.' + addressA);
  }

  return workbook;
}

function expectedLargeSheetTotal(rowCount, inputOffset) {
  const sumOfInputs = (rowCount * (rowCount + 1)) / 2;

  return (3 * sumOfInputs) + (3 * rowCount * inputOffset) + rowCount;
}

function expectedSharedFormulaTotal(rowCount, inputOffset) {
  const sumOfInputs = (rowCount * (rowCount + 1)) / 2;

  return (2 * sumOfInputs) + (2 * rowCount * inputOffset);
}

function evaluateLargeSheet(workbook, rowCount) {
  const sheet = workbook.requireSheet('Model');
  let total = 0;

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    total += Number(sheet.getCellValue('D' + String(rowNumber)));
  }

  return total;
}

function mutateLargeSheetInputs(workbook, rowCount, inputOffset) {
  const sheet = workbook.requireSheet('Model');

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    sheet.setCellValue('A' + String(rowNumber), rowNumber + inputOffset);
  }
}

function evaluateSharedFormulaSheet(workbook, rowCount) {
  const sheet = workbook.requireSheet('Shared');
  let total = 0;

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    total += Number(sheet.getCellValue('B' + String(rowNumber)));
  }

  return total;
}

function mutateSharedFormulaInputs(workbook, rowCount, inputOffset) {
  const sheet = workbook.requireSheet('Shared');

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    sheet.setCellValue('A' + String(rowNumber), rowNumber + inputOffset);
  }
}

function assertExpected(label, actualValue, expectedValue) {
  if (actualValue !== expectedValue) {
    throw new Error(label + ' produced ' + String(actualValue) + ', expected ' + String(expectedValue) + '.');
  }
}

function summarizeDurations(durations) {
  const sortedDurations = durations.slice().sort(function compareDurations(left, right) {
    return left - right;
  });
  const totalDuration = durations.reduce(function sumDurations(total, value) {
    return total + value;
  }, 0);

  return {
    averageMs: totalDuration / durations.length,
    maxMs: sortedDurations[sortedDurations.length - 1],
    minMs: sortedDurations[0]
  };
}

function formatDuration(value) {
  return value.toFixed(2);
}

function printResults(results, options) {
  writeStdout('Excellent evaluation benchmarks');
  writeStdout('Node ' + process.version + ' | ' + process.platform + ' ' + process.arch);
  writeStdout('rows=' + String(options.rows) + ' sharedRows=' + String(options.sharedRows) + ' iterations=' + String(options.iterations) + ' warmups=' + String(options.warmups));
  writeStdout('');

  const labelWidth = Math.max('Benchmark'.length, ...results.map(function(result) {
    return result.label.length;
  })) + 2;
  const sizeWidth = Math.max('Shape'.length, ...results.map(function(result) {
    return result.shape.length;
  })) + 2;
  const durationWidth = 10;
  const checksumWidth = 16;
  const header = [
    'Benchmark'.padEnd(labelWidth),
    'Shape'.padEnd(sizeWidth),
    'avg ms'.padStart(durationWidth),
    'min ms'.padStart(durationWidth),
    'max ms'.padStart(durationWidth),
    'checksum'.padStart(checksumWidth)
  ].join(' ');

  writeStdout(header);
  writeStdout('-'.repeat(header.length));

  for (const result of results) {
    writeStdout([
      result.label.padEnd(labelWidth),
      result.shape.padEnd(sizeWidth),
      formatDuration(result.averageMs).padStart(durationWidth),
      formatDuration(result.minMs).padStart(durationWidth),
      formatDuration(result.maxMs).padStart(durationWidth),
      String(result.checksum).padStart(checksumWidth)
    ].join(' '));
  }
}

async function measureBenchmarkCase(benchmarkCase, options, suiteContext) {
  const durations = [];
  let checksum = 0;

  for (let iteration = 0; iteration < options.warmups + options.iterations; iteration += 1) {
    const context = await benchmarkCase.prepare(options, suiteContext);
    const startTime = performance.now();
    const actualValue = await benchmarkCase.run(context, options, suiteContext);
    const elapsedMs = performance.now() - startTime;

    assertExpected(benchmarkCase.label, actualValue, benchmarkCase.expected(options));

    if (iteration >= options.warmups) {
      durations.push(elapsedMs);
      checksum = actualValue;
    }
  }

  return Object.assign({
    checksum: checksum,
    label: benchmarkCase.label,
    shape: benchmarkCase.shape(options)
  }, summarizeDurations(durations));
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const sharedFormulaBuffer = await createSharedFormulaWorkbookBuffer(options.sharedRows);
  const suiteContext = {
    sharedFormulaBuffer: sharedFormulaBuffer
  };
  const benchmarkCases = [
    {
      expected(currentOptions) {
        return expectedLargeSheetTotal(currentOptions.rows, 0);
      },
      label: 'large-sheet cold evaluation',
      async prepare(currentOptions) {
        return createLargeSheetWorkbook(currentOptions.rows);
      },
      async run(workbook, currentOptions) {
        return evaluateLargeSheet(workbook, currentOptions.rows);
      },
      shape(currentOptions) {
        return String(currentOptions.rows) + ' rows / ' + String(currentOptions.rows * 3) + ' formulas';
      }
    },
    {
      expected(currentOptions) {
        return expectedLargeSheetTotal(currentOptions.rows, 1);
      },
      label: 'large-sheet recalculation',
      async prepare(currentOptions) {
        const workbook = createLargeSheetWorkbook(currentOptions.rows);

        assertExpected('large-sheet cache priming', evaluateLargeSheet(workbook, currentOptions.rows), expectedLargeSheetTotal(currentOptions.rows, 0));
        return workbook;
      },
      async run(workbook, currentOptions) {
        mutateLargeSheetInputs(workbook, currentOptions.rows, 1);
        return evaluateLargeSheet(workbook, currentOptions.rows);
      },
      shape(currentOptions) {
        return String(currentOptions.rows) + ' rows / ' + String(currentOptions.rows * 3) + ' formulas';
      }
    },
    {
      expected(currentOptions) {
        return expectedSharedFormulaTotal(currentOptions.sharedRows, 0);
      },
      label: 'shared-formula cold evaluation',
      async prepare(currentOptions, currentSuiteContext) {
        return await new XlsxReader().load(currentSuiteContext.sharedFormulaBuffer);
      },
      async run(workbook, currentOptions) {
        return evaluateSharedFormulaSheet(workbook, currentOptions.sharedRows);
      },
      shape(currentOptions) {
        return String(currentOptions.sharedRows) + ' rows / ' + String(currentOptions.sharedRows) + ' shared formulas';
      }
    },
    {
      expected(currentOptions) {
        return expectedSharedFormulaTotal(currentOptions.sharedRows, 1);
      },
      label: 'shared-formula recalculation',
      async prepare(currentOptions, currentSuiteContext) {
        const workbook = await new XlsxReader().load(currentSuiteContext.sharedFormulaBuffer);

        assertExpected('shared-formula cache priming', evaluateSharedFormulaSheet(workbook, currentOptions.sharedRows), expectedSharedFormulaTotal(currentOptions.sharedRows, 0));
        return workbook;
      },
      async run(workbook, currentOptions) {
        mutateSharedFormulaInputs(workbook, currentOptions.sharedRows, 1);
        return evaluateSharedFormulaSheet(workbook, currentOptions.sharedRows);
      },
      shape(currentOptions) {
        return String(currentOptions.sharedRows) + ' rows / ' + String(currentOptions.sharedRows) + ' shared formulas';
      }
    }
  ];
  const results = [];

  for (const benchmarkCase of benchmarkCases) {
    results.push(await measureBenchmarkCase(benchmarkCase, options, suiteContext));
  }

  printResults(results, options);
}

main().catch(function(error) {
  writeStderr(String(error && (error.stack || error)));
  process.exitCode = 1;
});
