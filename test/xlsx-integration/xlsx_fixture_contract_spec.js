var assert = require('node:assert');
var fs = require('node:fs');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var excellentPackage = require('../..');

var FIXTURE_CONTRACTS = [
  {
    fileName: 'simpleFormula.xlsx',
    workbook: {
      fileVersion: '5.5.23515',
      sheetNames: ['Sheet1'],
      type: 'xl'
    },
    sheets: {
      Sheet1: {
        cellNames: ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9'],
        functions: ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9'],
        variables: [],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: 'Formula.SUM(1,2)',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 0,
            value: 3
          },
          A4: {
            address: 'A4',
            columnIndex: 0,
            formulaSource: 'Formula.POWER(10,2)',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 3,
            value: 100
          },
          A9: {
            address: 'A9',
            columnIndex: 0,
            formulaSource: 'Formula.MOD(55,3)',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 8,
            value: 1
          }
        }
      }
    }
  },
  {
    fileName: 'simpleRange.xlsx',
    workbook: {
      fileVersion: '5.5.23515',
      sheetNames: ['Sheet1'],
      type: 'xl'
    },
    sheets: {
      Sheet1: {
        cellNames: [
          'A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1',
          'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2',
          'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3',
          'I3', 'J3', 'K3', 'L3', 'B4', 'E4', 'F4', 'G4',
          'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12',
          'B13', 'B14', 'B15', 'B16', 'A20', 'A21', 'A22', 'A23',
          'A24', 'A25'
        ],
        functions: ['A20', 'A21', 'A22', 'A23', 'A24', 'A25'],
        variables: [
          'A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1',
          'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2',
          'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3',
          'I3', 'J3', 'K3', 'L3', 'B4', 'E4', 'F4', 'G4',
          'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12',
          'B13', 'B14', 'B15', 'B16'
        ],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 1,
            rowIndex: 0,
            value: 1
          },
          L3: {
            address: 'L3',
            columnIndex: 11,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 28,
            rowIndex: 2,
            value: 28
          },
          B4: {
            address: 'B4',
            columnIndex: 1,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 2,
            rowIndex: 3,
            value: 2
          },
          A20: {
            address: 'A20',
            columnIndex: 0,
            formulaSource: 'Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1])',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 19,
            value: 36
          },
          A24: {
            address: 'A24',
            columnIndex: 0,
            formulaSource: 'Formula.AVERAGE([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 23,
            value: 12.5
          },
          A25: {
            address: 'A25',
            columnIndex: 0,
            formulaSource: 'Formula.SUM([this.B1,this.B2,this.B3,this.B4,this.B5,this.B6,this.B7,this.B8,this.B9,this.B10,this.B11,this.B12,this.B13,this.B14,this.B15,this.B16])',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 24,
            value: 134
          }
        }
      }
    }
  },
  {
    fileName: 'sharedFormulas.xlsx',
    workbook: {
      fileVersion: '7.7.24026',
      sheetNames: ['Shared'],
      type: 'xl'
    },
    sheets: {
      Shared: {
        cellNames: ['A1', 'B1', 'B2', 'B3', 'C1', 'A2', 'A3'],
        functions: ['B1', 'B2', 'B3', 'C1'],
        variables: ['A1', 'A2', 'A3'],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 1,
            rowIndex: 0,
            value: 1
          },
          B1: {
            address: 'B1',
            columnIndex: 1,
            formulaSource: 'this.A1+1',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 0,
            value: 2
          },
          B2: {
            address: 'B2',
            columnIndex: 1,
            formulaSource: 'this.A2+1',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 1,
            value: 6
          },
          B3: {
            address: 'B3',
            columnIndex: 1,
            formulaSource: 'this.A3+1',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 2,
            value: 10
          },
          C1: {
            address: 'C1',
            columnIndex: 2,
            formulaSource: 'Formula.SUM([this.B1,this.B2,this.B3])',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 0,
            value: 18
          }
        }
      }
    }
  },
  {
    fileName: 'crossSheetWorkbook.xlsx',
    workbook: {
      fileVersion: '7.7.24026',
      sheetNames: ['Inputs', 'Outputs'],
      type: 'xl'
    },
    sheets: {
      Inputs: {
        cellNames: ['A1', 'A2'],
        functions: [],
        variables: ['A1', 'A2'],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 4,
            rowIndex: 0,
            value: 4
          },
          A2: {
            address: 'A2',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 5,
            rowIndex: 1,
            value: 5
          }
        }
      },
      Outputs: {
        cellNames: ['A1', 'B1', 'A2'],
        functions: ['A1', 'B1', 'A2'],
        variables: [],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: "self.workbook['Inputs'].A1+1",
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 0,
            value: 5
          },
          B1: {
            address: 'B1',
            columnIndex: 1,
            formulaSource: "Formula.SUM(self.workbook['Inputs'].A1,this.A2)",
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 0,
            value: 10
          },
          A2: {
            address: 'A2',
            columnIndex: 0,
            formulaSource: "self.workbook['Inputs'].A2+1",
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 1,
            value: 6
          }
        }
      }
    }
  },
  {
    fileName: 'quotedSheetAndErrors.xlsx',
    workbook: {
      fileVersion: '7.7.24026',
      sheetNames: ['Budget 2026', 'Summary'],
      type: 'xl'
    },
    sheets: {
      'Budget 2026': {
        cellNames: ['A1', 'A2'],
        functions: [],
        variables: ['A1', 'A2'],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 7,
            rowIndex: 0,
            value: 7
          },
          A2: {
            address: 'A2',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 8,
            rowIndex: 1,
            value: 8
          }
        }
      },
      Summary: {
        cellNames: ['A1', 'A2', 'A3', 'A4', 'A5', 'A6'],
        functions: ['A1', 'A3', 'A4', 'A5', 'A6'],
        variables: ['A2'],
        cells: {
          A1: {
            address: 'A1',
            columnIndex: 0,
            formulaSource: "self.workbook['Budget 2026'].A1+1",
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 0,
            value: 8
          },
          A2: {
            address: 'A2',
            columnIndex: 0,
            formulaSource: undefined,
            kind: 'value',
            rawValue: 2,
            rowIndex: 1,
            value: 2
          },
          A3: {
            address: 'A3',
            columnIndex: 0,
            formulaSource: "Formula.SUM(this.A2,self.workbook['Budget 2026'].A2)",
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 2,
            value: 10
          },
          A4: {
            address: 'A4',
            columnIndex: 0,
            formulaSource: 'Formula.IF("He said \\"hi\\""=="He said \\"hi\\"",1,0)',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 3,
            value: 1
          },
          A5: {
            address: 'A5',
            columnIndex: 0,
            formulaSource: 'Formula.IFERROR(#DIV/0!,99)',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 4,
            value: 99
          },
          A6: {
            address: 'A6',
            columnIndex: 0,
            formulaSource: 'Formula.IFNA(#N/A,77)',
            kind: 'formula',
            rawValue: undefined,
            rowIndex: 5,
            value: 77
          }
        }
      }
    }
  }
];

function normalizeValue(value) {
  if (excellentPackage.isExcelError(value)) {
    return {
      excelError: value.code
    };
  }

  if (Array.isArray(value)) {
    return value.map(normalizeValue);
  }

  return value;
}

function snapshotCell(cell) {
  return {
    address: cell.address,
    columnIndex: cell.columnIndex,
    formulaSource: cell.getFormulaSource(),
    kind: cell.kind,
    rawValue: normalizeValue(cell.getRawValue()),
    rowIndex: cell.rowIndex,
    value: normalizeValue(cell.getComputedValue())
  };
}

function assertSheetContract(workbook, sheetName, expectedSheet) {
  var sheet = workbook.getSheet(sheetName);

  assert.ok(sheet, 'missing sheet ' + sheetName);
  assert.deepEqual(sheet.getCellNames(), expectedSheet.cellNames);
  assert.deepEqual(sheet.getCells().map(function(cell) {
    return cell.address;
  }), expectedSheet.cellNames);
  assert.deepEqual(sheet.functions, expectedSheet.functions);
  assert.deepEqual(sheet.variables, expectedSheet.variables);

  Object.entries(expectedSheet.cells).forEach(function(entry) {
    var cellName = entry[0];
    var expectedCell = entry[1];
    var cell = sheet.getCell(cellName);

    assert.ok(cell, 'missing cell ' + sheetName + '!' + cellName);
    assert.deepEqual(snapshotCell(cell), expectedCell);
    assert.equal(workbook.getCell(sheetName, cellName), cell);
    assert.equal(workbook.getFormulaSource(sheetName, cellName), expectedCell.formulaSource);
  });
}

describe('ExcellentXlsxFixtureContracts', function() {
  'use strict';

  FIXTURE_CONTRACTS.forEach(function(contract) {
    it('loads ' + contract.fileName + ' with the expected values and workbook metadata', async function() {
      var reader = new excellentPackage.XlsxReader();
      var workbook = await reader.load(fs.readFileSync('./test/data/' + contract.fileName));

      assert.equal(workbook.type, contract.workbook.type);
      assert.equal(workbook.fileVersion, contract.workbook.fileVersion);
      assert.deepEqual(workbook.getSheetNames(), contract.workbook.sheetNames);

      Object.entries(contract.sheets).forEach(function(entry) {
        assertSheetContract(workbook, entry[0], entry[1]);
      });
    });
  });
});
