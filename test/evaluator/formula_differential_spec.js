var assert = require('node:assert');
var test = require('node:test');
var fc = require('fast-check');
var hyperformula = require('hyperformula');
var excellentPackage = require('../..');
var FormulaParser = require('../../dist/excellent.parser.js').FormulaParser;

var describe = test.describe;
var it = test.it;
var HyperFormula = hyperformula.HyperFormula;

var PROPERTY_RUNS = 100;
var PROPERTY_SEED = 20260331;
var NUMERIC_TOLERANCE = 1e-9;

function formula(expression) {
  return {
    formula: expression
  };
}

function isFormulaCell(value) {
  return value !== null &&
    typeof value === 'object' &&
    typeof value.formula === 'string';
}

function getCellAddress(cellName) {
  return {
    col: excellentPackage.Util.getColFromCell(cellName),
    row: excellentPackage.Util.getRowFromCell(cellName)
  };
}

function normalizeValue(value) {
  if (excellentPackage.isExcelError(value)) {
    return String(value);
  }

  if (value !== null && typeof value === 'object') {
    var maybeString = String(value);

    if (maybeString.startsWith('#')) {
      return maybeString;
    }
  }

  if (typeof value === 'number' && Object.is(value, -0)) {
    return 0;
  }

  return value;
}

function buildExcellentWorkbook(sheetData) {
  var workbook = new excellentPackage.Workbook();

  Object.keys(sheetData).forEach(function(sheetName) {
    workbook.createSheet(sheetName);
  });

  Object.entries(sheetData).forEach(function(entry) {
    var sheetName = entry[0];
    var rows = entry[1];
    var sheet = workbook.getSheet(sheetName);

    rows.forEach(function(row, rowIndex) {
      if (row === undefined || row === null) {
        return;
      }

      row.forEach(function(cellValue, columnIndex) {
        var cellName;

        if (cellValue === undefined) {
          return;
        }

        cellName = String(excellentPackage.Util.toBase26(columnIndex) + (rowIndex + 1));

        if (isFormulaCell(cellValue)) {
          sheet.setCellFormula(cellName, FormulaParser.parse(cellValue.formula));
        } else {
          sheet.setCellValue(cellName, cellValue);
        }
      });
    });
  });

  return workbook;
}

function buildHyperFormulaWorkbook(sheetData) {
  var normalizedSheets = {};

  Object.entries(sheetData).forEach(function(entry) {
    var sheetName = entry[0];
    var rows = entry[1];

    normalizedSheets[sheetName] = rows.map(function(row) {
      if (row === undefined || row === null) {
        return [];
      }

      return row.map(function(cellValue) {
        if (cellValue === undefined) {
          return null;
        }

        if (isFormulaCell(cellValue)) {
          return '=' + cellValue.formula;
        }

        return cellValue;
      });
    });
  });

  return HyperFormula.buildFromSheets(normalizedSheets, {
    licenseKey: 'gpl-v3'
  });
}

function compareAddressValues(excellentWorkbook, hyperFormulaWorkbook, addresses) {
  addresses.forEach(function(address) {
    var cellAddress = getCellAddress(address.cellName);
    var excellentValue = excellentWorkbook.getCellValue(address.sheetName, address.cellName);
    var hyperFormulaValue = hyperFormulaWorkbook.getCellValue({
      col: cellAddress.col,
      row: cellAddress.row,
      sheet: hyperFormulaWorkbook.getSheetId(address.sheetName)
    });

    assertComparableValues(
      normalizeValue(excellentValue),
      normalizeValue(hyperFormulaValue),
      address.sheetName + '!' + address.cellName
    );
  });
}

function assertComparableValues(leftValue, rightValue, label) {
  if (typeof leftValue === 'number' && typeof rightValue === 'number') {
    assert.ok(
      Math.abs(leftValue - rightValue) <= NUMERIC_TOLERANCE,
      label + ': expected ' + String(leftValue) + ' to be within ' + String(NUMERIC_TOLERANCE) + ' of ' + String(rightValue)
    );
    return;
  }

  assert.deepEqual(leftValue, rightValue, label);
}

function setExcellentCell(workbook, mutation) {
  workbook.setCellValue(mutation.sheetName, mutation.cellName, mutation.value);
}

function setHyperFormulaCell(hyperFormulaWorkbook, mutation) {
  var cellAddress = getCellAddress(mutation.cellName);

  hyperFormulaWorkbook.setCellContents({
    col: cellAddress.col,
    row: cellAddress.row,
    sheet: hyperFormulaWorkbook.getSheetId(mutation.sheetName)
  }, [[mutation.value]]);
}

function propertyAssert(property) {
  fc.assert(property, {
    numRuns: PROPERTY_RUNS,
    seed: PROPERTY_SEED
  });
}

function createReference(sheetName, cellName) {
  return {
    cellName: cellName,
    sheetName: sheetName
  };
}

function formatReference(reference) {
  if (reference.sheetName === null) {
    return reference.cellName;
  }

  return reference.sheetName + '!' + reference.cellName;
}

function isAtomicFormulaExpression(expression) {
  return expression.type === 'literal' || expression.type === 'reference';
}

function formatFormulaExpression(expression) {
  switch (expression.type) {
    case 'binary':
      return '(' + formatFormulaExpression(expression.left) + expression.operator + formatFormulaExpression(expression.right) + ')';
    case 'call':
      return expression.name + '(' + expression.arguments.map(formatFormulaExpression).join(',') + ')';
    case 'compare':
      return '(' + formatFormulaExpression(expression.left) + expression.operator + formatFormulaExpression(expression.right) + ')';
    case 'if':
      return 'IF(' +
        formatFormulaExpression(expression.condition) + ',' +
        formatFormulaExpression(expression.whenTrue) + ',' +
        formatFormulaExpression(expression.whenFalse) + ')';
    case 'iferror':
      return 'IFERROR(' +
        formatFormulaExpression(expression.candidate) + ',' +
        formatFormulaExpression(expression.fallback) + ')';
    case 'literal':
      return String(expression.value);
    case 'reference':
      return formatReference(expression.reference);
    case 'unary':
      if (isAtomicFormulaExpression(expression.argument)) {
        return expression.operator + formatFormulaExpression(expression.argument);
      }

      return expression.operator + '(' + formatFormulaExpression(expression.argument) + ')';
    default:
      throw new Error('Unsupported differential expression type: ' + expression.type);
  }
}

var numericReferenceArbitrary = fc.constantFrom(
  createReference(null, 'A1'),
  createReference(null, 'B1'),
  createReference(null, 'A2'),
  createReference(null, 'B2'),
  createReference('Inputs', 'A1'),
  createReference('Inputs', 'B1'),
  createReference('Inputs', 'A2'),
  createReference('Inputs', 'B2')
);

var numericExpressionArbitrary = fc.letrec(function(tie) {
  var atomArbitrary = fc.oneof(
    fc.integer({ min: -8, max: 8 }).map(function(value) {
      return {
        type: 'literal',
        value: value
      };
    }),
    numericReferenceArbitrary.map(function(reference) {
      return {
        reference: reference,
        type: 'reference'
      };
    })
  );

  var comparisonArbitrary = fc.tuple(
    tie('numeric'),
    fc.constantFrom('=', '<', '>'),
    tie('numeric')
  ).map(function(parts) {
    return {
      left: parts[0],
      operator: parts[1],
      right: parts[2],
      type: 'compare'
    };
  });

  return {
    numeric: fc.oneof(
      atomArbitrary,
      fc.tuple(fc.constantFrom('+', '-'), tie('numeric')).map(function(parts) {
        return {
          argument: parts[1],
          operator: parts[0],
          type: 'unary'
        };
      }),
      fc.tuple(tie('numeric'), fc.constantFrom('+', '-', '*', '/'), tie('numeric')).map(function(parts) {
        return {
          left: parts[0],
          operator: parts[1],
          right: parts[2],
          type: 'binary'
        };
      }),
      fc.array(tie('numeric'), { minLength: 1, maxLength: 3 }).map(function(argumentsList) {
        return {
          arguments: argumentsList,
          name: 'SUM',
          type: 'call'
        };
      }),
      tie('numeric').map(function(argument) {
        return {
          arguments: [argument],
          name: 'ABS',
          type: 'call'
        };
      }),
      fc.tuple(comparisonArbitrary, tie('numeric'), tie('numeric')).map(function(parts) {
        return {
          condition: parts[0],
          type: 'if',
          whenFalse: parts[2],
          whenTrue: parts[1]
        };
      }),
      fc.tuple(tie('numeric'), tie('numeric')).map(function(parts) {
        return {
          candidate: parts[0],
          fallback: parts[1],
          type: 'iferror'
        };
      })
    )
  };
}).numeric;

describe('ExcellentDifferentialAgainstHyperFormula', function() {
  'use strict';

  it('matches HyperFormula for overlapping formulas across sheets before and after mutations', function() {
    var sheetData = {
      Inputs: [
        [4, 5, 2, 0, 'a', 'b'],
        [10, 3]
      ],
      'Budget 2026': [
        [7, 8]
      ],
      Calc: [
        [
          formula('Inputs!A1+1'),
          formula('SUM(Inputs!A1,Inputs!B1)'),
          formula('AVERAGE(Inputs!A1,Inputs!B1,Inputs!C1)'),
          formula('IF(Inputs!A1<Inputs!B1,Inputs!B1,Inputs!A1)')
        ],
        [
          formula('IFERROR(Inputs!A1/Inputs!D1,99)'),
          formula('ABS(-Inputs!B1)'),
          formula('MOD(Inputs!B2*Inputs!C1,Inputs!C1+1)'),
          formula("'Budget 2026'!A1+'Budget 2026'!A2")
        ],
        [
          formula('Inputs!E1&Inputs!F1'),
          formula('SUM(A1:B1)'),
          formula('A1>B1'),
          formula('A1=B1')
        ]
      ]
    };
    var comparedAddresses = [
      { cellName: 'A1', sheetName: 'Calc' },
      { cellName: 'B1', sheetName: 'Calc' },
      { cellName: 'C1', sheetName: 'Calc' },
      { cellName: 'D1', sheetName: 'Calc' },
      { cellName: 'A2', sheetName: 'Calc' },
      { cellName: 'B2', sheetName: 'Calc' },
      { cellName: 'C2', sheetName: 'Calc' },
      { cellName: 'D2', sheetName: 'Calc' },
      { cellName: 'A3', sheetName: 'Calc' },
      { cellName: 'B3', sheetName: 'Calc' },
      { cellName: 'C3', sheetName: 'Calc' },
      { cellName: 'D3', sheetName: 'Calc' }
    ];
    var mutations = [
      { cellName: 'A1', sheetName: 'Inputs', value: 10 },
      { cellName: 'D1', sheetName: 'Inputs', value: 2 },
      { cellName: 'A2', sheetName: 'Budget 2026', value: 1 },
      { cellName: 'E1', sheetName: 'Inputs', value: 'x' }
    ];
    var excellentWorkbook = buildExcellentWorkbook(sheetData);
    var hyperFormulaWorkbook = buildHyperFormulaWorkbook(sheetData);

    compareAddressValues(excellentWorkbook, hyperFormulaWorkbook, comparedAddresses);

    mutations.forEach(function(mutation) {
      setExcellentCell(excellentWorkbook, mutation);
      setHyperFormulaCell(hyperFormulaWorkbook, mutation);
      compareAddressValues(excellentWorkbook, hyperFormulaWorkbook, comparedAddresses);
    });
  });

  it('matches HyperFormula for overlapping Excel error formulas', function() {
    var sheetData = {
      Errors: [
        [
          formula('1/0'),
          formula('SUM(A1,1)'),
          formula('UNKNOWNFUNC(1)'),
          formula('IFERROR(A1,99)'),
          formula('NA()'),
          formula('IFNA(E1,77)')
        ]
      ]
    };
    var comparedAddresses = [
      { cellName: 'A1', sheetName: 'Errors' },
      { cellName: 'B1', sheetName: 'Errors' },
      { cellName: 'C1', sheetName: 'Errors' },
      { cellName: 'D1', sheetName: 'Errors' },
      { cellName: 'E1', sheetName: 'Errors' },
      { cellName: 'F1', sheetName: 'Errors' }
    ];

    compareAddressValues(
      buildExcellentWorkbook(sheetData),
      buildHyperFormulaWorkbook(sheetData),
      comparedAddresses
    );
  });

  it('matches HyperFormula for generated numeric formulas in the shared subset', function() {
    propertyAssert(fc.property(
      fc.record({
        calcA1: fc.integer({ min: -8, max: 8 }),
        calcA2: fc.integer({ min: -8, max: 8 }),
        calcB1: fc.integer({ min: -8, max: 8 }),
        calcB2: fc.integer({ min: -8, max: 8 }),
        inputA1: fc.integer({ min: -8, max: 8 }),
        inputA2: fc.integer({ min: -8, max: 8 }),
        inputB1: fc.integer({ min: -8, max: 8 }),
        inputB2: fc.integer({ min: -8, max: 8 })
      }),
      numericExpressionArbitrary,
      function(inputValues, expression) {
        var formulaText = formatFormulaExpression(expression);
        var sheetData = {
          Calc: [
            [inputValues.calcA1, inputValues.calcB1],
            [inputValues.calcA2, inputValues.calcB2],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [],
            [undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, formula(formulaText)]
          ],
          Inputs: [
            [inputValues.inputA1, inputValues.inputB1],
            [inputValues.inputA2, inputValues.inputB2]
          ]
        };
        var excellentWorkbook = buildExcellentWorkbook(sheetData);
        var hyperFormulaWorkbook = buildHyperFormulaWorkbook(sheetData);
        var excellentValue = excellentWorkbook.getCellValue('Calc', 'Z26');
        var hyperFormulaValue = hyperFormulaWorkbook.getCellValue({
          col: getCellAddress('Z26').col,
          row: getCellAddress('Z26').row,
          sheet: hyperFormulaWorkbook.getSheetId('Calc')
        });

        assertComparableValues(
          normalizeValue(excellentValue),
          normalizeValue(hyperFormulaValue),
          formulaText
        );
      }
    ));
  });
});
