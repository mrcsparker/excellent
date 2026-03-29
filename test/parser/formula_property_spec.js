var assert = require('node:assert');
var test = require('node:test');
var fc = require('fast-check');
var excellentPackage = require('../..');
var FormulaParser = require('../../dist/excellent.parser.js').FormulaParser;
var FormulaRuntime = require('../../dist/formula/index.js');

var describe = test.describe;
var it = test.it;

var PROPERTY_RUNS = 150;
var PROPERTY_SEED = 20260329;
var TARGET_CELL = 'Z99';
var SHEET_NAMES = [null, 'Inputs', 'Sheet 2', 'Budget 2026'];

function columnIndexToName(columnIndex) {
  var currentIndex = columnIndex;
  var label = '';

  do {
    label = String.fromCharCode(65 + (currentIndex % 26)) + label;
    currentIndex = Math.floor(currentIndex / 26) - 1;
  } while (currentIndex >= 0);

  return label;
}

function quoteSheetName(sheetName) {
  if (/^[A-Za-z0-9_]+$/.test(sheetName)) {
    return sheetName;
  }

  return '\'' + sheetName.replace(/'/g, '\'\'') + '\'';
}

function getCellName(reference) {
  return columnIndexToName(reference.columnIndex) + String(reference.row);
}

function formatCellReference(reference) {
  var cellName = getCellName(reference);
  var columnName = cellName.replace(/[0-9]+$/, '');
  var rowNumber = String(reference.row);
  var prefix = reference.sheet === null ? '' : quoteSheetName(reference.sheet) + '!';

  return prefix +
    (reference.absoluteColumn ? '$' : '') + columnName +
    (reference.absoluteRow ? '$' : '') + rowNumber;
}

function normalizeCellReference(reference) {
  return {
    ref: getCellName(reference),
    sheet: reference.sheet
  };
}

function ensureSheet(workbook, sheetName) {
  var existingSheet = workbook.getSheet(sheetName);

  if (existingSheet !== undefined) {
    return existingSheet;
  }

  return workbook.createSheet(sheetName);
}

function createWorkbookForReferences(referenceEntries) {
  var workbook = new excellentPackage.Workbook();
  var calcSheet = ensureSheet(workbook, 'Calc');

  referenceEntries.forEach(function(entry) {
    var targetSheet = ensureSheet(workbook, entry.reference.sheet || 'Calc');

    targetSheet.setCellValue(getCellName(entry.reference), entry.value);
  });

  return {
    workbook: workbook,
    worksheet: calcSheet
  };
}

function createWorkbookForRange(rangeSpec) {
  var workbook = new excellentPackage.Workbook();
  var calcSheet = ensureSheet(workbook, 'Calc');
  var targetSheet = ensureSheet(workbook, rangeSpec.sheet || 'Calc');
  var references = [];
  var values = [];

  for (var rowOffset = 0; rowOffset < rangeSpec.height; rowOffset += 1) {
    for (var columnOffset = 0; columnOffset < rangeSpec.width; columnOffset += 1) {
      var reference = {
        absoluteColumn: false,
        absoluteRow: false,
        columnIndex: rangeSpec.startColumnIndex + columnOffset,
        row: rangeSpec.startRow + rowOffset,
        sheet: rangeSpec.sheet
      };
      var valueIndex = rowOffset * rangeSpec.width + columnOffset;
      var nextValue = rangeSpec.values[valueIndex];

      if (nextValue === undefined) {
        throw new Error('Range value generation mismatch.');
      }

      targetSheet.setCellValue(getCellName(reference), nextValue);
      references.push(reference);
      values.push(nextValue);
    }
  }

  return {
    references: references,
    values: values,
    workbook: workbook,
    worksheet: calcSheet
  };
}

function formatRange(rangeSpec) {
  var startReference = formatCellReference({
    absoluteColumn: rangeSpec.startAbsoluteColumn,
    absoluteRow: rangeSpec.startAbsoluteRow,
    columnIndex: rangeSpec.startColumnIndex,
    row: rangeSpec.startRow,
    sheet: rangeSpec.sheet
  });
  var endReference = formatCellReference({
    absoluteColumn: rangeSpec.endAbsoluteColumn,
    absoluteRow: rangeSpec.endAbsoluteRow,
    columnIndex: rangeSpec.startColumnIndex + rangeSpec.width - 1,
    row: rangeSpec.startRow + rangeSpec.height - 1,
    sheet: null
  });

  return startReference + ':' + endReference;
}

function createNumberLiteral(value) {
  return {
    type: 'number',
    value: value
  };
}

function createStringLiteral(value) {
  return {
    type: 'string',
    value: value
  };
}

function createReferenceExpression(reference) {
  return {
    reference: reference,
    type: 'reference'
  };
}

function createUnaryExpression(operator, argument) {
  return {
    argument: argument,
    operator: operator,
    type: 'unary'
  };
}

function createBinaryExpression(left, operator, right) {
  return {
    left: left,
    operator: operator,
    right: right,
    type: 'binary'
  };
}

function createFunctionCall(name, argumentsList) {
  return {
    arguments: argumentsList,
    name: name,
    type: 'call'
  };
}

function isAtomicExpression(expression) {
  return expression.type === 'number' ||
    expression.type === 'reference' ||
    expression.type === 'string';
}

function formatGeneratedFormula(expression) {
  switch (expression.type) {
    case 'binary':
      return '(' +
        formatGeneratedFormula(expression.left) +
        ' ' + expression.operator + ' ' +
        formatGeneratedFormula(expression.right) +
        ')';
    case 'call':
      return expression.name + '(' + expression.arguments.map(formatGeneratedFormula).join(',') + ')';
    case 'number':
      return String(expression.value);
    case 'reference':
      return formatCellReference(expression.reference);
    case 'string':
      return '"' + expression.value.replace(/"/g, '""') + '"';
    case 'unary':
      if (isAtomicExpression(expression.argument)) {
        return expression.operator + formatGeneratedFormula(expression.argument);
      }

      return expression.operator + '(' + formatGeneratedFormula(expression.argument) + ')';
    default:
      throw new Error('Unsupported generated formula expression type: ' + expression.type);
  }
}

function propertyAssert(property) {
  fc.assert(property, {
    numRuns: PROPERTY_RUNS,
    seed: PROPERTY_SEED
  });
}

var cellReferenceArbitrary = fc.record({
  absoluteColumn: fc.boolean(),
  absoluteRow: fc.boolean(),
  columnIndex: fc.integer({ min: 0, max: 7 }),
  row: fc.integer({ min: 1, max: 12 }),
  sheet: fc.constantFrom.apply(fc, SHEET_NAMES)
});

var cellReferenceEntriesArbitrary = fc.uniqueArray(fc.record({
  reference: cellReferenceArbitrary,
  value: fc.integer({ min: -50, max: 50 })
}), {
  maxLength: 5,
  minLength: 1,
  selector: function(entry) {
    var normalizedReference = normalizeCellReference(entry.reference);

    return String(normalizedReference.sheet || '') + '!' + normalizedReference.ref;
  }
});

var rangeSpecArbitrary = fc.record({
  endAbsoluteColumn: fc.boolean(),
  endAbsoluteRow: fc.boolean(),
  height: fc.integer({ min: 1, max: 4 }),
  sheet: fc.constant(null),
  startAbsoluteColumn: fc.boolean(),
  startAbsoluteRow: fc.boolean(),
  startColumnIndex: fc.integer({ min: 0, max: 4 }),
  startRow: fc.integer({ min: 1, max: 4 }),
  width: fc.integer({ min: 1, max: 4 })
}).chain(function(rangeSpec) {
  return fc.array(fc.integer({ min: -25, max: 25 }), {
    maxLength: rangeSpec.width * rangeSpec.height,
    minLength: rangeSpec.width * rangeSpec.height
  }).map(function(values) {
    return Object.assign({}, rangeSpec, {
      values: values
    });
  });
});

var generatedExpressionArbitrary = fc.letrec(function(tie) {
  var atomArbitrary = fc.oneof(
    fc.integer({ min: -20, max: 20 }).map(createNumberLiteral),
    fc.string({ maxLength: 6 }).map(createStringLiteral),
    cellReferenceArbitrary.map(createReferenceExpression)
  );

  return {
    expression: fc.oneof(
      atomArbitrary,
      fc.constantFrom('+', '-').chain(function(operator) {
        return tie('expression').map(function(argument) {
          return createUnaryExpression(operator, argument);
        });
      }),
      fc.tuple(
        tie('expression'),
        fc.constantFrom('+', '-', '*', '/', '^', '&', '=', '<', '>'),
        tie('expression')
      ).map(function(parts) {
        return createBinaryExpression(parts[0], parts[1], parts[2]);
      }),
      fc.array(tie('expression'), { minLength: 1, maxLength: 3 }).map(function(argumentsList) {
        return createFunctionCall('SUM', argumentsList);
      }),
      tie('expression').map(function(argument) {
        return createFunctionCall('ABS', [argument]);
      }),
      fc.tuple(tie('expression'), tie('expression'), tie('expression')).map(function(parts) {
        return createFunctionCall('IF', [parts[0], parts[1], parts[2]]);
      })
    )
  };
}).expression;

describe('ExcellentFormulaProperties', function() {
  'use strict';

  it('parses generated cell-reference formulas and preserves normalized references', function() {
    var evaluator = new excellentPackage.FormulaEvaluator();

    propertyAssert(fc.property(cellReferenceEntriesArbitrary, function(referenceEntries) {
      var formulaText = 'SUM(' + referenceEntries.map(function(entry) {
        return formatCellReference(entry.reference);
      }).join(',') + ')';
      var parsedFormula = FormulaParser.parse(formulaText);
      var workbookContext = createWorkbookForReferences(referenceEntries);
      var expectedReferences = referenceEntries.map(function(entry) {
        return normalizeCellReference(entry.reference);
      });
      var expectedValue = referenceEntries.reduce(function(sum, entry) {
        return sum + entry.value;
      }, 0);

      workbookContext.worksheet.setCellFormula(TARGET_CELL, parsedFormula);

      assert.deepEqual(evaluator.collectReferences(parsedFormula), expectedReferences);
      assert.equal(
        workbookContext.workbook.getCellValue('Calc', TARGET_CELL),
        expectedValue
      );
      assert.equal(
        workbookContext.workbook.getFormulaSource('Calc', TARGET_CELL),
        FormulaRuntime.serializeFormulaAst(parsedFormula)
      );
    }));
  });

  it('parses generated range formulas and expands references in row-major order', function() {
    var evaluator = new excellentPackage.FormulaEvaluator();

    propertyAssert(fc.property(
      fc.constantFrom('SUM', 'AVERAGE'),
      rangeSpecArbitrary,
      function(functionName, rangeSpec) {
        var workbookContext = createWorkbookForRange(rangeSpec);
        var formulaText = functionName + '(' + formatRange(rangeSpec) + ')';
        var parsedFormula = FormulaParser.parse(formulaText);
        var expectedReferences = workbookContext.references.map(normalizeCellReference);
        var sum = workbookContext.values.reduce(function(total, value) {
          return total + value;
        }, 0);
        var expectedValue = functionName === 'SUM'
          ? sum
          : sum / workbookContext.values.length;

        workbookContext.worksheet.setCellFormula(TARGET_CELL, parsedFormula);

        assert.deepEqual(evaluator.collectReferences(parsedFormula), expectedReferences);
        assert.equal(
          workbookContext.workbook.getCellValue('Calc', TARGET_CELL),
          expectedValue
        );
      }
    ));
  });

  it('keeps generated parser canonicalization stable through formula serialization', function() {
    var evaluator = new excellentPackage.FormulaEvaluator();

    propertyAssert(fc.property(generatedExpressionArbitrary, function(expression) {
      var formulaText = formatGeneratedFormula(expression);
      var parsedFormula = FormulaParser.parse(formulaText);
      var canonical = FormulaRuntime.serializeFormulaAst(parsedFormula);
      var normalizedCanonical = evaluator.serialize(canonical);

      assert.equal(evaluator.serialize(parsedFormula), canonical);
      assert.equal(evaluator.serialize(normalizedCanonical), normalizedCanonical);
      assert.deepEqual(
        evaluator.collectReferences(parsedFormula),
        evaluator.collectReferences(normalizedCanonical)
      );
    }));
  });
});
