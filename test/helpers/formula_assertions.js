'use strict';

var assert = require('node:assert');

function assertExcelError(FormulaRuntime, value, code) {
  assert.equal(FormulaRuntime.isExcelError(value), true);
  assert.equal(value.code, code);
  assert.equal(String(value), code);
}

function assertFormulaSource(worksheet, cellName, expectedSource) {
  assert.equal(worksheet.getFormulaSource(cellName), expectedSource);
}

module.exports = {
  assertExcelError: assertExcelError,
  assertFormulaSource: assertFormulaSource
};
