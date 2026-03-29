var assert = require('node:assert');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var FormulaParser = require('../../dist/excellent.parser.js').FormulaParser;
var FormulaRuntime = require('../../dist/formula/index.js');
var regressionCases = require('../data/formula_parser_regressions.json');

describe('ExcellentFormulaParserRegression', function() {
  'use strict';

  for (const testCase of regressionCases.valid) {
    it('matches regression case: ' + testCase.label, function() {
      var ast = FormulaParser.parse(testCase.formula);

      assert.deepEqual(ast, testCase.ast);
      assert.equal(FormulaRuntime.serializeFormulaAst(ast), testCase.debug);
    });
  }

  for (const testCase of regressionCases.invalid) {
    it('matches syntax error regression case: ' + testCase.label, function() {
      assert.throws(function() {
        FormulaParser.parse(testCase.formula);
      }, function(error) {
        return error.message === testCase.message &&
          JSON.stringify(error.location) === JSON.stringify(testCase.location);
      });
    });
  }
});
