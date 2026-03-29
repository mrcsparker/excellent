'use strict';

import type { CompiledFormula, WorkbookCellValue } from '../formula';
import { Util } from '../excellent.util';
import type { EvaluationState } from './types';
import type { Worksheet } from './worksheet';

class Cell {
  address: string;
  columnIndex: number;
  compiledFormula: CompiledFormula | undefined;
  index: string;
  kind: 'formula' | 'value';
  rawValue: WorkbookCellValue;
  rowIndex: number;
  worksheet!: Worksheet;

  constructor(worksheet: Worksheet, address: string) {
    this.address = address;
    this.columnIndex = Util.getColFromCell(address);
    this.index = address;
    this.kind = 'value';
    this.rawValue = undefined;
    this.rowIndex = Util.getRowFromCell(address);

    Object.defineProperty(this, 'compiledFormula', {
      configurable: true,
      enumerable: false,
      value: undefined,
      writable: true
    });

    Object.defineProperty(this, 'worksheet', {
      configurable: true,
      enumerable: false,
      value: worksheet,
      writable: false
    });
  }

  get key() {
    return this.ownerWorkbook.getCellKey(this.sheetName, this.address);
  }

  get ownerWorkbook() {
    return this.worksheet.ownerWorkbook;
  }

  get sheetName() {
    return this.worksheet.name;
  }

  get value() {
    return this.rawValue;
  }

  set value(nextValue) {
    this.rawValue = nextValue;
  }

  getCompiledFormula() {
    return this.compiledFormula;
  }

  getComputedValue(evaluationState?: EvaluationState) {
    if (!this.isFormula()) {
      return this.getRawValue();
    }

    return this.ownerWorkbook.getCellValue(this.sheetName, this.address, evaluationState);
  }

  getFormulaSource() {
    if (!this.isFormula()) {
      return undefined;
    }

    return this.compiledFormula?.expression;
  }

  getRawValue() {
    if (this.isFormula()) {
      return undefined;
    }

    return this.rawValue;
  }

  isFormula() {
    return this.kind === 'formula';
  }

  setCompiledFormula(compiledFormula: CompiledFormula) {
    this.compiledFormula = compiledFormula;
    return this;
  }

  setFormula(compiledFormula: CompiledFormula) {
    this.kind = 'formula';
    this.compiledFormula = compiledFormula;
    this.rawValue = compiledFormula.expression;
    return this;
  }

  setRawValue(value: WorkbookCellValue) {
    this.rawValue = value;
    return this;
  }

  setValue(value: WorkbookCellValue) {
    this.kind = 'value';
    this.compiledFormula = undefined;
    this.rawValue = value;
    return this;
  }
}

export { Cell };
