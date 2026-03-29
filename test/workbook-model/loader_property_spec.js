var assert = require('node:assert');
var test = require('node:test');
var fc = require('fast-check');
var excellentPackage = require('../..');

var describe = test.describe;
var it = test.it;

var PROPERTY_RUNS = 150;
var PROPERTY_SEED = 20260330;
var SCALAR_VALUE_ARBITRARY = fc.oneof(
  fc.boolean(),
  fc.integer({ min: -50, max: 50 }),
  fc.string({ maxLength: 8 }),
  fc.constant(null)
);

function propertyAssert(property) {
  fc.assert(property, {
    numRuns: PROPERTY_RUNS,
    seed: PROPERTY_SEED
  });
}

function buildWorkbookPayload(sheetNames, sheetSpecs) {
  var workbookPayload = {};

  sheetNames.forEach(function(sheetName, sheetIndex) {
    var previousSheetName = sheetIndex > 0 ? sheetNames[sheetIndex - 1] : null;
    var spec = sheetSpecs[sheetIndex];

    if (spec === undefined) {
      throw new Error('Missing sheet spec for ' + sheetName);
    }

    workbookPayload[sheetName] = [
      [spec.a1, spec.b1, spec.c1],
      [
        '[function]this.A1+' + String(spec.a2Delta),
        '[function]this.B1*' + String(spec.b2Multiplier),
        spec.c2
      ],
      [
        previousSheetName === null
          ? '[function]this.A2+this.B2'
          : '[function]self.workbook[\'' + previousSheetName.replace(/'/g, '\\\'') + '\'].A1+' + String(spec.crossSheetDelta),
        '[function]this.A2+this.B2',
        spec.c3
      ]
    ];
  });

  return workbookPayload;
}

function getSnapshot(workbook) {
  var snapshot = {};

  workbook.getSheetNames().forEach(function(sheetName) {
    snapshot[sheetName] = {
      A1: workbook.getCellValue(sheetName, 'A1'),
      A2: workbook.getCellValue(sheetName, 'A2'),
      A3: workbook.getCellValue(sheetName, 'A3'),
      B1: workbook.getCellValue(sheetName, 'B1'),
      B2: workbook.getCellValue(sheetName, 'B2'),
      B3: workbook.getCellValue(sheetName, 'B3'),
      C1: workbook.getCellValue(sheetName, 'C1'),
      C2: workbook.getCellValue(sheetName, 'C2'),
      C3: workbook.getCellValue(sheetName, 'C3'),
      formulaSourceA2: workbook.getFormulaSource(sheetName, 'A2'),
      formulaSourceA3: workbook.getFormulaSource(sheetName, 'A3'),
      formulaSourceB2: workbook.getFormulaSource(sheetName, 'B2'),
      formulaSourceB3: workbook.getFormulaSource(sheetName, 'B3')
    };
  });

  return snapshot;
}

var workbookPayloadArbitrary = fc.uniqueArray(
  fc.constantFrom('Inputs', 'Outputs', 'Budget 2026'),
  { minLength: 1, maxLength: 3 }
).chain(function(sheetNames) {
  return fc.array(fc.record({
    a1: fc.integer({ min: -50, max: 50 }),
    a2Delta: fc.integer({ min: -20, max: 20 }),
    b1: fc.integer({ min: -50, max: 50 }),
    b2Multiplier: fc.integer({ min: -5, max: 5 }),
    c1: SCALAR_VALUE_ARBITRARY,
    c2: SCALAR_VALUE_ARBITRARY,
    c3: SCALAR_VALUE_ARBITRARY,
    crossSheetDelta: fc.integer({ min: -20, max: 20 })
  }), { minLength: sheetNames.length, maxLength: sheetNames.length }).map(function(sheetSpecs) {
    return buildWorkbookPayload(sheetNames, sheetSpecs);
  });
});

describe('ExcellentWorkbookLoaderProperties', function() {
  'use strict';

  it('round-trips generated serialized workbooks without changing payloads or computed values', function() {
    var loader = new excellentPackage.WorkbookLoader();

    propertyAssert(fc.property(workbookPayloadArbitrary, function(workbookPayload) {
      var workbook = loader.deserialize(workbookPayload);
      var serialized = loader.serialize(workbook);
      var restoredWorkbook = loader.deserialize(serialized);
      var restoredSerialized = loader.serialize(restoredWorkbook);

      assert.deepEqual(serialized, workbookPayload);
      assert.deepEqual(restoredSerialized, workbookPayload);
      assert.deepEqual(getSnapshot(restoredWorkbook), getSnapshot(workbook));
    }));
  });
});
