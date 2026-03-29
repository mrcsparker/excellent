var assert = require('node:assert');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');
var Workbook = excellentPackage.Workbook;
var FormulaParser = require('../../dist/excellent.parser.js').FormulaParser;
var FormulaRuntime = require('../../dist/formula/index.js');
var FormulaPackage = require('@formulajs/formulajs');
var Formula = FormulaPackage.default || FormulaPackage;
var formulaAssertions = require('../helpers/formula_assertions.js');
var assertExcelError = formulaAssertions.assertExcelError.bind(null, FormulaRuntime);
var assertFormulaSource = formulaAssertions.assertFormulaSource;

function createDirectRuntime() {
  var workbook = new Workbook();
  var worksheet = workbook.addSheet('Sheet1');

  return {
    functionRegistry: workbook.functionRegistry,
    workbook: workbook,
    worksheet: worksheet
  };
}

function createTraceRuntime(baseRuntime) {
  return {
    evaluationState: {
      active: new Set(),
      stack: []
    },
    functionRegistry: baseRuntime.functionRegistry,
    traceState: {
      active: new Set(),
      stack: []
    },
    workbook: baseRuntime.workbook,
    worksheet: baseRuntime.worksheet
  };
}

function compileExpression(expression) {
  return FormulaRuntime.compileFormula(expression);
}

describe('ExcellentEvaluator', function() {
  'use strict';

  it('should be able to evaluate simple formulas from formula.js', function() {
    assert.equal(Formula.SUM(1, 2), 3);
    assert.equal(Formula.DELTA(42, 42), 1);
  });

  it('evaluates compiled worksheet formulas without eval', function() {
    var workbook = new Workbook();
    var sheet1 = workbook.addSheet('Sheet1');
    var sheet2;

    sheet1.addCellVal('A1', 2);
    sheet1.addCellFunc('A2', 'this.A1^3');

    sheet2 = workbook.addSheet('Sheet2');
    sheet2.addCellFunc('B1', "self.workbook['Sheet1'].A2+1");

    assertFormulaSource(sheet1, 'A2', 'this.A1^3');
    assert.equal(sheet1.A2, 8);
    assert.equal(sheet2.B1, 9);
  });

  it('invalidates dependent formula caches when an input cell changes', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellVal('A1', 2);
    sheet.addCellFunc('A2', 'this.A1+1');
    sheet.addCellFunc('A3', 'this.A2+1');

    assert.equal(sheet.A3, 4);

    sheet.A1 = 10;

    assert.equal(sheet.A2, 11);
    assert.equal(sheet.A3, 12);
  });

  it('detects circular worksheet dependencies', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellFunc('A1', 'this.A2');
    sheet.addCellFunc('A2', 'this.A1');

    assert.throws(function() {
      return sheet.A1;
    }, function(error) {
      return error instanceof FormulaRuntime.FormulaCycleError &&
        /Sheet1!A1 -> Sheet1!A2 -> Sheet1!A1/.test(error.message);
    });
  });

  it('exposes direct precedents and dependents across worksheets', function() {
    var workbook = new Workbook();
    var sheet1 = workbook.addSheet('Sheet1');
    var sheet2;

    sheet1.addCellVal('A1', 3);
    sheet1.addCellFunc('A2', 'this.A1+1');

    sheet2 = workbook.addSheet('Sheet2');
    sheet2.addCellFunc('B1', "self.workbook['Sheet1'].A2+5");

    assert.deepEqual(workbook.getPrecedents('Sheet2', 'B1'), [{
      cellName: 'A2',
      key: 'Sheet1!A2',
      sheetName: 'Sheet1'
    }]);

    assert.deepEqual(workbook.getDependents('Sheet1', 'A2'), [{
      cellName: 'B1',
      key: 'Sheet2!B1',
      sheetName: 'Sheet2'
    }]);
  });

  it('traverses dependency graphs transitively', function() {
    var workbook = new Workbook();
    var sheet1 = workbook.addSheet('Sheet1');
    var sheet2;

    sheet1.addCellVal('A1', 1);
    sheet1.addCellFunc('A2', 'this.A1+1');

    sheet2 = workbook.addSheet('Sheet2');
    sheet2.addCellFunc('B1', "self.workbook['Sheet1'].A2+1");
    sheet2.addCellFunc('B2', 'this.B1+1');

    assert.deepEqual(workbook.traversePrecedents('Sheet2', 'B2'), [
      {
        cellName: 'B1',
        key: 'Sheet2!B1',
        sheetName: 'Sheet2'
      },
      {
        cellName: 'A2',
        key: 'Sheet1!A2',
        sheetName: 'Sheet1'
      },
      {
        cellName: 'A1',
        key: 'Sheet1!A1',
        sheetName: 'Sheet1'
      }
    ]);

    assert.deepEqual(workbook.traverseDependents('Sheet1', 'A1'), [
      {
        cellName: 'A2',
        key: 'Sheet1!A2',
        sheetName: 'Sheet1'
      },
      {
        cellName: 'B1',
        key: 'Sheet2!B1',
        sheetName: 'Sheet2'
      },
      {
        cellName: 'B2',
        key: 'Sheet2!B2',
        sheetName: 'Sheet2'
      }
    ]);
  });

  it('traces nested formula evaluation across worksheets', function() {
    var workbook = new Workbook();
    var sheet1 = workbook.addSheet('Sheet1');
    var sheet2;
    var trace;

    sheet1.addCellVal('A1', 2);
    sheet1.addCellFunc('A2', 'this.A1+1');

    sheet2 = workbook.addSheet('Sheet2');
    sheet2.addCellFunc('B1', "self.workbook['Sheet1'].A2+5");
    sheet2.addCellFunc('B2', 'this.B1*2');

    trace = workbook.traceCell('Sheet2', 'B2');

    assert.equal(trace.kind, 'formula');
    assert.equal(trace.expression, 'this.B1*2');
    assert.equal(trace.value, 16);
    assert.deepEqual(trace.precedents, [{
      cellName: 'B1',
      key: 'Sheet2!B1',
      sheetName: 'Sheet2'
    }]);
    assert.equal(trace.evaluation.type, 'binary-expression');
    assert.equal(trace.evaluation.left.type, 'cell-reference');
    assert.equal(trace.evaluation.left.cell.key, 'Sheet2!B1');
    assert.equal(trace.evaluation.left.cell.value, 8);
    assert.equal(trace.evaluation.left.cell.evaluation.type, 'binary-expression');
    assert.equal(trace.evaluation.left.cell.evaluation.left.type, 'cell-reference');
    assert.equal(trace.evaluation.left.cell.evaluation.left.cell.key, 'Sheet1!A2');
    assert.equal(trace.evaluation.left.cell.evaluation.left.cell.value, 3);
  });

  it('traces raw value cells without pretending they are formulas', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');
    var trace;

    sheet.addCellVal('A1', 7);
    trace = workbook.traceCell('Sheet1', 'A1');

    assert.deepEqual(trace, {
      cellName: 'A1',
      key: 'Sheet1!A1',
      kind: 'value',
      precedents: [],
      rawValue: 7,
      sheetName: 'Sheet1',
      value: 7
    });
  });

  it('returns first-class Excel error values and propagates them through formulas', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellVal('A1', 1);
    sheet.addCellFunc('A2', 'this.A1/0');
    sheet.addCellFunc('A3', FormulaParser.parse('SUM(A2,1)'));

    assertExcelError(sheet.A2, '#DIV/0!');
    assertExcelError(sheet.A3, '#DIV/0!');
  });

  it('treats missing sheet references as #REF! instead of undefined', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellFunc('A1', "self.workbook['Missing'].B2+1");

    assertExcelError(sheet.A1, '#REF!');
    assertExcelError(workbook.traceCell('Missing', 'B2').value, '#REF!');
  });

  it('returns #NAME? for unknown formula functions instead of throwing', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellVal('A1', 4);
    sheet.addCellFunc('A2', FormulaParser.parse('DOUBLE(A1)'));

    assertExcelError(sheet.A2, '#NAME?');
  });

  it('short-circuits IF and only evaluates the selected branch', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellFunc('A1', FormulaParser.parse('IF(0,1/0,42)'));
    sheet.addCellFunc('A2', FormulaParser.parse('IF(1,42,1/0)'));

    assert.equal(sheet.A1, 42);
    assert.equal(sheet.A2, 42);
  });

  it('handles IFERROR and IFNA using Excel error values', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellFunc('A1', FormulaParser.parse('IFERROR(1/0,99)'));
    sheet.addCellFunc('A2', FormulaParser.parse('IFNA(MATCH("x",A3:A4,0),77)'));
    sheet.addCellVal('A3', 'a');
    sheet.addCellVal('A4', 'b');
    sheet.addCellFunc('A5', FormulaParser.parse('IFNA(1/0,77)'));

    assert.equal(sheet.A1, 99);
    assert.equal(sheet.A2, 77);
    assertExcelError(sheet.A5, '#DIV/0!');
  });

  it('uses formulajs semantics for INDEX and MATCH instead of legacy overrides', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellVal('A1', 10);
    sheet.addCellVal('B1', 20);
    sheet.addCellVal('A2', 'A');
    sheet.addCellVal('B2', 'B');
    sheet.addCellFunc('C1', FormulaParser.parse('INDEX(A1:B1,MATCH("B",A2:B2,0))'));

    assert.equal(sheet.C1, 20);
  });

  it('evaluates absolute references against the normalized cell address', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellVal('A1', 2);
    sheet.addCellVal('A2', 3);
    sheet.addCellFunc('A3', FormulaParser.parse('SUM($A$1,A2)'));

    assert.equal(sheet.A3, 5);
    assert.deepEqual(workbook.getPrecedents('Sheet1', 'A3'), [
      {
        cellName: 'A1',
        key: 'Sheet1!A1',
        sheetName: 'Sheet1'
      },
      {
        cellName: 'A2',
        key: 'Sheet1!A2',
        sheetName: 'Sheet1'
      }
    ]);
  });

  it('parses and evaluates escaped string literals', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');
    var parsedFormula = FormulaParser.parse('IF("He said ""hi"""="He said ""hi""",1,0)');

    assert.equal(
      FormulaRuntime.serializeFormulaAst(parsedFormula),
      'Formula.IF("He said \\"hi\\""=="He said \\"hi\\"",1,0)'
    );

    sheet.addCellFunc('A1', parsedFormula);
    assert.equal(sheet.A1, 1);
  });

  it('coerces empty cells, nulls, and numeric strings consistently', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellVal('A1', null);
    sheet.addCellVal('A2', '2');
    sheet.addCellFunc('A3', 'this.A1+1');
    sheet.addCellFunc('A4', 'this.A2+1');
    sheet.addCellFunc('A5', FormulaParser.parse('SUM(B1,1)'));

    assert.equal(sheet.A3, 1);
    assert.equal(sheet.A4, 3);
    assert.equal(sheet.A5, 1);
  });

  it('supports explicit Excel error literals in formulas', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    sheet.addCellFunc('A1', FormulaParser.parse('IFERROR(#DIV/0!,99)'));
    sheet.addCellFunc('A2', FormulaParser.parse('SUM(#VALUE!,1)'));

    assert.equal(sheet.A1, 99);
    assertExcelError(sheet.A2, '#VALUE!');
  });

  it('detects circular dependencies across worksheets', function() {
    var workbook = new Workbook();
    var sheet1 = workbook.addSheet('Sheet1');
    var sheet2 = workbook.addSheet('Sheet2');

    sheet1.addCellFunc('A1', FormulaParser.parse('Sheet2!B1'));
    sheet2.addCellFunc('B1', FormulaParser.parse('Sheet1!A1'));

    assert.throws(function() {
      return sheet1.A1;
    }, function(error) {
      return error instanceof FormulaRuntime.FormulaCycleError &&
        /Sheet1!A1 -> Sheet2!B1 -> Sheet1!A1/.test(error.message);
    });
  });

  it('supports workbook-scoped custom formula functions', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    workbook.registerFunction('DOUBLE', function(value) {
      return Number(value) * 2;
    });

    sheet.addCellVal('A1', 4);
    sheet.addCellFunc('A2', FormulaParser.parse('DOUBLE(A1)'));
    sheet.addCellFunc('A3', 'Formula.DOUBLE(this.A2)');

    assert.equal(workbook.hasFunction('DOUBLE'), true);
    assert.equal(sheet.A2, 8);
    assert.equal(sheet.A3, 16);
    assert.equal(FormulaRuntime.Formula.DOUBLE, undefined);
  });

  it('keeps custom functions scoped to each workbook', function() {
    var workbook = new Workbook();
    var otherWorkbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');
    var otherSheet = otherWorkbook.addSheet('Sheet1');

    workbook.registerFunction('DOUBLE', function(value) {
      return Number(value) * 2;
    });

    sheet.addCellVal('A1', 4);
    sheet.addCellFunc('A2', FormulaParser.parse('DOUBLE(A1)'));

    otherSheet.addCellVal('A1', 4);
    otherSheet.addCellFunc('A2', FormulaParser.parse('DOUBLE(A1)'));

    assert.equal(sheet.A2, 8);
    assertExcelError(otherSheet.A2, '#NAME?');
  });

  it('supports overriding built-in functions without mutating the defaults', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    assert.throws(function() {
      workbook.registerFunction('SUM', function() {
        return 99;
      });
    }, /Formula function already exists: SUM/);

    workbook.registerFunction('SUM', function() {
      return 99;
    }, { override: true });

    sheet.addCellFunc('A1', FormulaParser.parse('SUM(1,2)'));
    assert.equal(sheet.A1, 99);

    workbook.unregisterFunction('SUM');
    assert.equal(sheet.A1, 3);
    assert.equal(FormulaRuntime.Formula.SUM(1, 2), 3);
  });

  it('rejects async custom functions', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');

    workbook.registerFunction('ASYNC_DOUBLE', async function(value) {
      return Number(value) * 2;
    });

    sheet.addCellVal('A1', 4);
    sheet.addCellFunc('A2', FormulaParser.parse('ASYNC_DOUBLE(A1)'));

    assertExcelError(sheet.A2, '#VALUE!');
  });

  it('covers omitted IF-style fallbacks and traces them through the evaluator', function() {
    var workbook = new Workbook();
    var sheet = workbook.addSheet('Sheet1');
    var missingFalseTrace;
    var missingTrueTrace;
    var ifErrorTrace;
    var ifNaTrace;
    var ifConditionErrorTrace;

    sheet.addCellFunc('A1', 'Formula.IF()');
    sheet.addCellFunc('A2', 'Formula.IF(1)');
    sheet.addCellFunc('A3', FormulaParser.parse('IFERROR(1/0)'));
    sheet.addCellFunc('A4', FormulaParser.parse('IFNA(MATCH("x",A5:A6,0))'));
    sheet.addCellVal('A5', 'a');
    sheet.addCellVal('A6', 'b');
    sheet.addCellFunc('A7', FormulaParser.parse('IF(#VALUE!,1,2)'));

    assert.equal(sheet.A1, false);
    assert.equal(sheet.A2, true);
    assert.equal(sheet.A3, '');
    assert.equal(sheet.A4, '');
    assertExcelError(sheet.A7, '#VALUE!');

    missingFalseTrace = workbook.traceCell('Sheet1', 'A1');
    missingTrueTrace = workbook.traceCell('Sheet1', 'A2');
    ifErrorTrace = workbook.traceCell('Sheet1', 'A3');
    ifNaTrace = workbook.traceCell('Sheet1', 'A4');
    ifConditionErrorTrace = workbook.traceCell('Sheet1', 'A7');

    assert.equal(missingFalseTrace.value, false);
    assert.equal(missingFalseTrace.evaluation.arguments[0].value, false);
    assert.equal(missingFalseTrace.evaluation.arguments[1].value, false);

    assert.equal(missingTrueTrace.value, true);
    assert.equal(missingTrueTrace.evaluation.arguments[0].value, 1);
    assert.equal(missingTrueTrace.evaluation.arguments[1].value, true);

    assert.equal(ifErrorTrace.value, '');
    assert.equal(ifErrorTrace.evaluation.arguments[1].value, '');

    assert.equal(ifNaTrace.value, '');
    assert.equal(ifNaTrace.evaluation.arguments[1].value, '');

    assertExcelError(ifConditionErrorTrace.value, '#VALUE!');
    assertExcelError(ifConditionErrorTrace.evaluation.value, '#VALUE!');
  });

  it('evaluates and traces direct identifiers and member access in the compatibility runtime', function() {
    var runtime = createDirectRuntime();
    var traceRuntime = createTraceRuntime(runtime);
    var namespaceValue;
    var workbookValue;
    var unknownIdentifier;
    var nestedMemberValue;
    var computedMemberValue;
    var nestedMemberTrace;
    var computedMemberTrace;
    var propertyErrorFormula = {
      ast: {
        computed: true,
        object: {
          type: 'ThisExpression'
        },
        property: {
          code: '#VALUE!',
          type: 'ErrorLiteral'
        },
        type: 'MemberExpression'
      },
      expression: 'this[#VALUE!]'
    };
    var propertyErrorTrace;

    runtime.workbook.extra = {
      answer: 7,
      values: ['zero', 9]
    };

    namespaceValue = FormulaRuntime.evaluateCompiledFormula(compileExpression('Formula'), runtime);
    workbookValue = FormulaRuntime.evaluateCompiledFormula(compileExpression('self'), runtime);
    unknownIdentifier = FormulaRuntime.evaluateCompiledFormula(compileExpression('UnknownIdentifier'), runtime);
    nestedMemberValue = FormulaRuntime.evaluateCompiledFormula(compileExpression('self.extra.answer'), runtime);
    computedMemberValue = FormulaRuntime.evaluateCompiledFormula(compileExpression('self.extra.values[1]'), runtime);
    nestedMemberTrace = FormulaRuntime.traceCompiledFormula(compileExpression('self.extra.answer'), traceRuntime);
    computedMemberTrace = FormulaRuntime.traceCompiledFormula(compileExpression('self.extra.values[1]'), traceRuntime);
    propertyErrorTrace = FormulaRuntime.traceCompiledFormula(propertyErrorFormula, createTraceRuntime(runtime));

    assert.equal(typeof namespaceValue.SUM, 'function');
    assert.strictEqual(workbookValue, runtime.workbook);
    assertExcelError(unknownIdentifier, '#NAME?');
    assert.equal(nestedMemberValue, 7);
    assert.equal(computedMemberValue, 9);

    assert.equal(nestedMemberTrace.type, 'member-expression');
    assert.equal(nestedMemberTrace.property, 'answer');
    assert.equal(nestedMemberTrace.value, 7);

    assert.equal(computedMemberTrace.type, 'member-expression');
    assert.equal(computedMemberTrace.property, 1);
    assert.equal(computedMemberTrace.value, 9);

    assertExcelError(
      FormulaRuntime.evaluateCompiledFormula(propertyErrorFormula, runtime),
      '#VALUE!'
    );
    assert.equal(propertyErrorTrace.type, 'member-expression');
    assertExcelError(propertyErrorTrace.value, '#VALUE!');
  });

  it('covers direct compatibility call-expression branches and guard rails', function() {
    var runtime = createDirectRuntime();
    var memberCallTrace;
    var nestedCallTrace;
    var invalidCallFormula = {
      ast: {
        arguments: [],
        callee: {
          type: 'Literal',
          value: 1
        },
        type: 'CallExpression'
      },
      expression: '1()'
    };

    runtime.workbook.extra = {
      answer: 7,
      invoke: function(value) {
        return this.answer + Number(value);
      },
      makeAdder: function() {
        return function(value) {
          return Number(value) + 5;
        };
      }
    };

    assert.equal(
      FormulaRuntime.evaluateCompiledFormula(compileExpression('self.extra.invoke(2)'), runtime),
      9
    );
    assert.equal(
      FormulaRuntime.evaluateCompiledFormula(compileExpression('self.extra.makeAdder()(3)'), runtime),
      8
    );

    memberCallTrace = FormulaRuntime.traceCompiledFormula(
      compileExpression('self.extra.invoke(2)'),
      createTraceRuntime(runtime)
    );
    nestedCallTrace = FormulaRuntime.traceCompiledFormula(
      compileExpression('self.extra.makeAdder()(3)'),
      createTraceRuntime(runtime)
    );

    assert.equal(memberCallTrace.type, 'call-expression');
    assert.equal(memberCallTrace.callee, 'invoke');
    assert.equal(memberCallTrace.value, 9);

    assert.equal(nestedCallTrace.type, 'call-expression');
    assert.equal(nestedCallTrace.callee, 'call');
    assert.equal(nestedCallTrace.value, 8);

    assert.throws(function() {
      return FormulaRuntime.evaluateCompiledFormula(compileExpression('self.extra.answer(1)'), runtime);
    }, /Expected callable member: answer/);

    assert.throws(function() {
      return FormulaRuntime.traceCompiledFormula(
        compileExpression('self.extra.answer(1)'),
        createTraceRuntime(runtime)
      );
    }, /Expected callable member: answer/);

    assert.throws(function() {
      return FormulaRuntime.evaluateCompiledFormula(invalidCallFormula, runtime);
    }, /Expected callable expression/);

    assert.throws(function() {
      return FormulaRuntime.traceCompiledFormula(
        invalidCallFormula,
        createTraceRuntime(runtime)
      );
    }, /Expected callable expression/);
  });

  it('throws explicit evaluator errors for unsupported AST nodes', function() {
    var runtime = createDirectRuntime();
    var unsupportedFormula = {
      ast: {
        type: 'UnsupportedNode'
      },
      expression: '[unsupported]'
    };

    assert.throws(function() {
      return FormulaRuntime.evaluateCompiledFormula(unsupportedFormula, runtime);
    }, /Unsupported formula node type/);

    assert.throws(function() {
      return FormulaRuntime.traceCompiledFormula(
        unsupportedFormula,
        createTraceRuntime(runtime)
      );
    }, /Unsupported formula node type/);
  });
});
