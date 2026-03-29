'use strict';

const process = require('node:process');
const { performance } = require('node:perf_hooks');
const JSZip = require('jszip');
const { XlsxReader } = require('../dist');

const DEFAULT_OPTIONS = Object.freeze({
  iterations: 5,
  mixedRows: 4000,
  sharedRows: 8000,
  warmups: 1
});

function writeStdout(message) {
  process.stdout.write(String(message) + '\n');
}

function writeStderr(message) {
  process.stderr.write(String(message) + '\n');
}

function printUsage() {
  writeStdout('Usage: npm run profile:xlsx-load -- [--mixed-rows <count>] [--shared-rows <count>] [--iterations <count>] [--warmups <count>]');
  writeStdout('');
  writeStdout('Defaults:');
  writeStdout('  --mixed-rows ' + String(DEFAULT_OPTIONS.mixedRows));
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
    mixedRows: DEFAULT_OPTIONS.mixedRows,
    sharedRows: DEFAULT_OPTIONS.sharedRows,
    warmups: DEFAULT_OPTIONS.warmups
  };

  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];

    if (argument === '--help' || argument === '-h') {
      printUsage();
      process.exit(0);
    }

    if (argument === '--mixed-rows') {
      options.mixedRows = parsePositiveInteger(readFlagValue(argv, index, '--mixed-rows'), '--mixed-rows');
      index += 1;
      continue;
    }

    if (argument.startsWith('--mixed-rows=')) {
      options.mixedRows = parsePositiveInteger(argument.slice('--mixed-rows='.length), '--mixed-rows');
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

function createContentTypesXml(sheetCount, hasSharedStrings) {
  const worksheetOverrides = Array.from({ length: sheetCount }, function(_value, index) {
    return '<Override PartName="/xl/worksheets/sheet' + String(index + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
  }).join('');
  const sharedStringsOverride = hasSharedStrings
    ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    : '';

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    worksheetOverrides,
    sharedStringsOverride,
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

function createWorkbookRelationshipsXml(sheetCount, hasSharedStrings) {
  const sheetRelationships = Array.from({ length: sheetCount }, function(_value, index) {
    return '<Relationship Id="rId' + String(index + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + String(index + 1) + '.xml"/>';
  }).join('');
  const sharedStringsRelationship = hasSharedStrings
    ? '<Relationship Id="rId' + String(sheetCount + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    : '';

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    sheetRelationships,
    sharedStringsRelationship,
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

function createSharedStringsXml(values) {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + String(values.length) + '" uniqueCount="' + String(values.length) + '">',
    values.map(function(value) {
      return '<si><t>' + value + '</t></si>';
    }).join(''),
    '</sst>'
  ].join('');
}

function createMixedWorkbookArtifacts(rowCount) {
  const rows = [];
  const sharedStrings = [];

  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    const label = 'Label ' + String(rowNumber);
    const sharedStringIndex = sharedStrings.push(label) - 1;
    const sharedCell = '<c r="A' + String(rowNumber) + '" t="s"><v>' + String(sharedStringIndex) + '</v></c>';
    const numericCell = '<c r="B' + String(rowNumber) + '"><v>' + String(rowNumber) + '</v></c>';
    const doubledFormula = '<c r="C' + String(rowNumber) + '"><f>B' + String(rowNumber) + '*2</f><v>' + String(rowNumber * 2) + '</v></c>';
    const incrementedFormula = '<c r="D' + String(rowNumber) + '"><f>C' + String(rowNumber) + '+1</f><v>' + String((rowNumber * 2) + 1) + '</v></c>';

    rows.push('<row r="' + String(rowNumber) + '" spans="1:4">' + sharedCell + numericCell + doubledFormula + incrementedFormula + '</row>');
  }

  return {
    sharedStringsXml: createSharedStringsXml(sharedStrings),
    worksheetXml: createWorksheetXml(rows)
  };
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

async function createWorkbookBuffer(sheetName, sheetXml, sharedStringsXml) {
  const zip = new JSZip();
  const hasSharedStrings = typeof sharedStringsXml === 'string';

  zip.file('[Content_Types].xml', createContentTypesXml(1, hasSharedStrings));
  zip.file('_rels/.rels', createRootRelationshipsXml());
  zip.file('xl/workbook.xml', createWorkbookXml([sheetName]));
  zip.file('xl/_rels/workbook.xml.rels', createWorkbookRelationshipsXml(1, hasSharedStrings));
  zip.file('xl/worksheets/sheet1.xml', sheetXml);

  if (hasSharedStrings) {
    zip.file('xl/sharedStrings.xml', sharedStringsXml);
  }

  return await zip.generateAsync({
    compression: 'DEFLATE',
    type: 'nodebuffer'
  });
}

async function createMixedWorkbookBuffer(rowCount) {
  const artifacts = createMixedWorkbookArtifacts(rowCount);

  return await createWorkbookBuffer('Profiled', artifacts.worksheetXml, artifacts.sharedStringsXml);
}

async function createSharedFormulaWorkbookBuffer(rowCount) {
  return await createWorkbookBuffer('Shared', createSharedFormulaWorksheetXml(rowCount));
}

function createProfileCollector() {
  return {
    counts: Object.create(null),
    durationsMs: Object.create(null),
    incrementCount(label, amount = 1) {
      this.counts[label] = (this.counts[label] ?? 0) + amount;
    },
    async measureAsync(label, callback) {
      const startTime = performance.now();

      try {
        return await callback();
      } finally {
        this.durationsMs[label] = (this.durationsMs[label] ?? 0) + (performance.now() - startTime);
      }
    },
    measureSync(label, callback) {
      const startTime = performance.now();

      try {
        return callback();
      } finally {
        this.durationsMs[label] = (this.durationsMs[label] ?? 0) + (performance.now() - startTime);
      }
    }
  };
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

function summarizeMetrics(metricSamples) {
  const totals = Object.create(null);

  for (const sample of metricSamples) {
    for (const [label, value] of Object.entries(sample)) {
      totals[label] = (totals[label] ?? 0) + value;
    }
  }

  return Object.entries(totals).map(function(entry) {
    const label = entry[0];
    const totalValue = entry[1];

    return {
      average: totalValue / metricSamples.length,
      label: label
    };
  }).sort(function compareMetrics(left, right) {
    return right.average - left.average;
  });
}

function formatDuration(value) {
  return value.toFixed(2);
}

function formatCount(value) {
  return Number.isInteger(value) ? String(value) : value.toFixed(2);
}

async function measureProfileCase(profileCase, options) {
  const totalDurations = [];
  const durationSnapshots = [];
  const countSnapshots = [];
  let verificationValue;

  for (let iteration = 0; iteration < options.warmups + options.iterations; iteration += 1) {
    const profile = createProfileCollector();
    const buffer = await profileCase.createBuffer(options);
    const startTime = performance.now();
    const workbook = await new XlsxReader({ profile: profile }).load(buffer);
    const totalElapsedMs = performance.now() - startTime;
    const currentVerification = profileCase.verify(workbook, options);

    if (iteration >= options.warmups) {
      totalDurations.push(totalElapsedMs);
      durationSnapshots.push(profile.durationsMs);
      countSnapshots.push(profile.counts);
      verificationValue = currentVerification;
    }
  }

  return {
    counts: summarizeMetrics(countSnapshots),
    durations: summarizeMetrics(durationSnapshots),
    shape: profileCase.shape(options),
    summary: summarizeDurations(totalDurations),
    verification: verificationValue
  };
}

function printPhaseTable(totalAverageMs, durations) {
  const topDurations = durations.slice(0, 8);
  const phaseWidth = Math.max('Phase'.length, ...topDurations.map(function(duration) {
    return duration.label.length;
  })) + 2;
  const durationWidth = 10;
  const shareWidth = 9;
  const header = [
    'Phase'.padEnd(phaseWidth),
    'avg ms'.padStart(durationWidth),
    '% total'.padStart(shareWidth)
  ].join(' ');

  writeStdout(header);
  writeStdout('-'.repeat(header.length));

  for (const duration of topDurations) {
    const share = totalAverageMs === 0 ? 0 : (duration.average / totalAverageMs) * 100;

    writeStdout([
      duration.label.padEnd(phaseWidth),
      formatDuration(duration.average).padStart(durationWidth),
      formatDuration(share).padStart(shareWidth)
    ].join(' '));
  }
}

function printCountTable(counts) {
  const countWidth = Math.max('Count'.length, ...counts.map(function(entry) {
    return entry.label.length;
  })) + 2;
  const valueWidth = 12;
  const header = [
    'Count'.padEnd(countWidth),
    'avg'.padStart(valueWidth)
  ].join(' ');

  writeStdout(header);
  writeStdout('-'.repeat(header.length));

  for (const entry of counts) {
    writeStdout([
      entry.label.padEnd(countWidth),
      formatCount(entry.average).padStart(valueWidth)
    ].join(' '));
  }
}

function printCaseResult(title, result) {
  writeStdout(title);
  writeStdout('shape: ' + result.shape);
  writeStdout('load avg ' + formatDuration(result.summary.averageMs) + ' ms | min ' + formatDuration(result.summary.minMs) + ' ms | max ' + formatDuration(result.summary.maxMs) + ' ms');
  writeStdout('verification: ' + result.verification);
  writeStdout('');
  printPhaseTable(result.summary.averageMs, result.durations);
  writeStdout('');
  printCountTable(result.counts);
  writeStdout('');
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const profileCases = [
    {
      title: 'Mixed workbook load',
      async createBuffer(currentOptions) {
        return await createMixedWorkbookBuffer(currentOptions.mixedRows);
      },
      shape(currentOptions) {
        return String(currentOptions.mixedRows) + ' rows / 1 shared string + 2 formulas per row';
      },
      verify(workbook, currentOptions) {
        const sheet = workbook.requireSheet('Profiled');
        const lastRow = currentOptions.mixedRows;

        return String(sheet.getCellValue('A1')) + ' | ' +
          String(sheet.getCellValue('C' + String(lastRow))) + ' | ' +
          String(sheet.getCellValue('D' + String(lastRow)));
      }
    },
    {
      title: 'Shared-formula workbook load',
      async createBuffer(currentOptions) {
        return await createSharedFormulaWorkbookBuffer(currentOptions.sharedRows);
      },
      shape(currentOptions) {
        return String(currentOptions.sharedRows) + ' rows / 1 shared formula per row';
      },
      verify(workbook, currentOptions) {
        const sheet = workbook.requireSheet('Shared');
        return String(sheet.getCellValue('B1')) + ' | ' + String(sheet.getCellValue('B' + String(currentOptions.sharedRows)));
      }
    }
  ];

  writeStdout('Excellent XLSX load profile');
  writeStdout('Node ' + process.version + ' | ' + process.platform + ' ' + process.arch);
  writeStdout('mixedRows=' + String(options.mixedRows) + ' sharedRows=' + String(options.sharedRows) + ' iterations=' + String(options.iterations) + ' warmups=' + String(options.warmups));
  writeStdout('');

  for (const profileCase of profileCases) {
    const result = await measureProfileCase(profileCase, options);

    printCaseResult(profileCase.title, result);
  }
}

main().catch(function handleError(error) {
  writeStderr(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
