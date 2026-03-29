'use strict';

import type { CompiledFormula, WorkbookCellValue } from '../formula';
import { Cell } from './cell';
import type { EvaluationState, WorkbookRows } from './types';
import type { Workbook } from './workbook';

function removeCellName(cellNames: string[], targetCellName: string) {
  return cellNames.filter(function filterCellName(cellName) {
    return cellName !== targetCellName;
  });
}


class Worksheet {
  [key: string]: unknown;
  cells!: Map<string, Cell>;
  functionSet!: Set<string>;
  functions: string[];
  name: string;
  ownerWorkbook!: Workbook;
  rows: WorkbookRows;
  variableSet!: Set<string>;
  variables: string[];

  constructor(name: string, ownerWorkbook: Workbook) {
    this.name = name;
    this.variables = [];
    this.functions = [];
    this.rows = [];

    Object.defineProperty(this, 'cells', {
      configurable: true,
      enumerable: false,
      value: new Map(),
      writable: true
    });

    Object.defineProperty(this, 'functionSet', {
      configurable: true,
      enumerable: false,
      value: new Set<string>(),
      writable: true
    });

    Object.defineProperty(this, 'ownerWorkbook', {
      configurable: true,
      enumerable: false,
      value: ownerWorkbook,
      writable: false
    });

    Object.defineProperty(this, 'variableSet', {
      configurable: true,
      enumerable: false,
      value: new Set<string>(),
      writable: true
    });
  }

  addCellFunc(cellName: string, formulaExpression: string | CompiledFormula) {
    this.setCellFormula(cellName, formulaExpression);
    return this;
  }

  addCellVal(cellName: string, cellValue: WorkbookCellValue) {
    this.setCellValue(cellName, cellValue);
    return this;
  }

  defineCellAccessors(cellName: string) {
    if (Object.prototype.hasOwnProperty.call(this, cellName)) {
      return;
    }

    Object.defineProperty(this, cellName, {
      configurable: true,
      enumerable: false,
      get: () => {
        return this.getCellValue(cellName);
      },
      set: (nextValue) => {
        this.setCellValue(cellName, nextValue as WorkbookCellValue);
      }
    });
  }

  ensureCell(cellName: string) {
    const existingCell = this.getCell(cellName);

    if (existingCell !== undefined) {
      return existingCell;
    }

    const cell = new Cell(this, cellName);

    this.cells.set(cellName, cell);
    this.defineCellAccessors(cellName);
    return cell;
  }

  getCell(cellName: string) {
    return this.cells.get(cellName);
  }

  getCellValue(cellName: string, evaluationState?: EvaluationState) {
    const cell = this.getCell(cellName);

    if (cell === undefined) {
      return undefined;
    }

    return cell.getComputedValue(evaluationState);
  }

  getCellNames() {
    return Array.from(this.cells.keys());
  }

  getCells() {
    return Array.from(this.cells.values());
  }

  getCompiledFormula(cellName: string) {
    const cell = this.getCell(cellName);

    if (cell === undefined) {
      return undefined;
    }

    return cell.getCompiledFormula();
  }

  getFormulaSource(cellName: string) {
    const cell = this.getCell(cellName);

    if (cell === undefined) {
      return undefined;
    }

    return cell.getFormulaSource();
  }

  getRawCellValue(cellName: string) {
    const cell = this.getCell(cellName);

    if (cell === undefined) {
      return undefined;
    }

    return cell.getRawValue();
  }

  hasCell(cellName: string) {
    return this.getCell(cellName) !== undefined;
  }

  hasFormula(cellName: string) {
    const cell = this.getCell(cellName);

    return cell !== undefined && cell.isFormula();
  }

  markFormula(cellName: string) {
    if (this.variableSet.delete(cellName)) {
      this.variables = removeCellName(this.variables, cellName);
    }

    if (!this.functionSet.has(cellName)) {
      this.functionSet.add(cellName);
      this.functions.push(cellName);
    }
  }

  markVariable(cellName: string) {
    if (this.functionSet.delete(cellName)) {
      this.functions = removeCellName(this.functions, cellName);
    }

    if (!this.variableSet.has(cellName)) {
      this.variableSet.add(cellName);
      this.variables.push(cellName);
    }
  }

  setCellFormula(cellName: string, formulaExpression: string | CompiledFormula) {
    const cell = this.ensureCell(cellName);
    const compiledFormula = this.ownerWorkbook.formulaEvaluator.compile(formulaExpression);

    this.markFormula(cellName);
    cell.setFormula(compiledFormula);
    this.syncRowEntry(cell);
    this.ownerWorkbook.registerFormulaCell(this.name, cellName, compiledFormula);
    return cell;
  }

  setCellValue(cellName: string, cellValue: WorkbookCellValue) {
    const cell = this.ensureCell(cellName);

    this.markVariable(cellName);
    cell.setValue(cellValue);
    this.syncRowEntry(cell);
    this.ownerWorkbook.registerValueCell(this.name, cellName);
    return cell;
  }

  setCompiledFormula(cellName: string, compiledFormula: CompiledFormula) {
    const cell = this.ensureCell(cellName);

    cell.setCompiledFormula(compiledFormula);
    this.syncRowEntry(cell);
    return cell;
  }

  setRawCellValue(cellName: string, value: WorkbookCellValue) {
    const cell = this.ensureCell(cellName);

    cell.setRawValue(value);
    this.syncRowEntry(cell);
    return cell;
  }

  syncRowEntry(cell: Cell) {
    const row = this.rows[cell.rowIndex] ?? [];

    this.rows[cell.rowIndex] = row;
    row[cell.columnIndex] = cell;
  }
}

export { Worksheet };
