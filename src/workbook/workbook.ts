'use strict';

import {
  ExcelError,
  FormulaEvaluator,
  type FormulaFunction,
  type FormulaFunctionOptions,
  FormulaFunctionRegistry,
  FormulaCycleError,
  FormulaEvaluationError,
  type CompiledFormula,
  type WorkbookCellValue
} from '../formula';
import type { Cell } from './cell';
import type {
  CellReferenceInfo,
  EvaluationState,
  WorkbookOptions,
  WorkbookSheets,
  WorkbookTraceCell,
  WorkbookTraceEvaluation
} from './types';
import { Worksheet } from './worksheet';

function toError(error: unknown): Error {
  if (error instanceof Error) {
    return error;
  }

  return new Error(String(error));
}

class Workbook {
  cellCache!: Map<string, WorkbookCellValue>;
  cellDependencies!: Map<string, Set<string>>;
  cellDependents!: Map<string, Set<string>>;
  fileVersion: string;
  formulaEvaluator!: FormulaEvaluator;
  functionRegistry!: FormulaFunctionRegistry;
  mutationBatchDepth!: number;
  pendingCacheReset!: boolean;
  type: string;
  workbook: WorkbookSheets;

  constructor(options?: WorkbookOptions) {
    const resolvedOptions = options || {};
    const formulaEvaluator = resolvedOptions.formulaEvaluator instanceof FormulaEvaluator
      ? resolvedOptions.formulaEvaluator
      : new FormulaEvaluator();
    const functionRegistry = resolvedOptions.functionRegistry instanceof FormulaFunctionRegistry
      ? resolvedOptions.functionRegistry.clone()
      : new FormulaFunctionRegistry();

    if (resolvedOptions.functions !== undefined) {
      functionRegistry.registerMany(resolvedOptions.functions);
    }

    this.fileVersion = '';
    this.type = '';
    this.workbook = {};

    Object.defineProperty(this, 'cellCache', {
      configurable: true,
      enumerable: false,
      value: new Map(),
      writable: true
    });

    Object.defineProperty(this, 'cellDependencies', {
      configurable: true,
      enumerable: false,
      value: new Map(),
      writable: true
    });

    Object.defineProperty(this, 'cellDependents', {
      configurable: true,
      enumerable: false,
      value: new Map(),
      writable: true
    });

    Object.defineProperty(this, 'formulaEvaluator', {
      configurable: true,
      enumerable: false,
      value: formulaEvaluator,
      writable: false
    });

    Object.defineProperty(this, 'functionRegistry', {
      configurable: true,
      enumerable: false,
      value: functionRegistry,
      writable: true
    });

    Object.defineProperty(this, 'mutationBatchDepth', {
      configurable: true,
      enumerable: false,
      value: 0,
      writable: true
    });

    Object.defineProperty(this, 'pendingCacheReset', {
      configurable: true,
      enumerable: false,
      value: false,
      writable: true
    });
  }

  addSheet(sheetName: string): Worksheet {
    return this.createSheet(sheetName);
  }

  beginMutationBatch() {
    this.mutationBatchDepth += 1;
    return this;
  }

  clearDependenciesForCell(cellKey: string) {
    const currentDependencies = this.cellDependencies.get(cellKey);

    if (currentDependencies === undefined) {
      return;
    }

    for (const dependencyKey of currentDependencies) {
      const dependents = this.cellDependents.get(dependencyKey);

      if (dependents === undefined) {
        continue;
      }

      dependents.delete(cellKey);

      if (dependents.size === 0) {
        this.cellDependents.delete(dependencyKey);
      }
    }

    this.cellDependencies.delete(cellKey);
  }

  createSheet(sheetName: string): Worksheet {
    const existingSheet = this.workbook[sheetName];

    if (existingSheet !== undefined) {
      return existingSheet;
    }

    const createdSheet = new Worksheet(sheetName, this);

    this.workbook[sheetName] = createdSheet;
    return createdSheet;
  }

  endMutationBatch() {
    if (this.mutationBatchDepth === 0) {
      throw new Error('Cannot end a workbook mutation batch that was not started.');
    }

    this.mutationBatchDepth -= 1;

    if (this.mutationBatchDepth === 0 && this.pendingCacheReset) {
      this.invalidateAllFormulaCaches();
      this.pendingCacheReset = false;
    }

    return this;
  }

  evaluateFormulaCell(cell: Cell, evaluationState?: EvaluationState) {
    const cellKey = cell.key;

    if (this.cellCache.has(cellKey)) {
      return this.cellCache.get(cellKey);
    }

    const state = evaluationState || {
      active: new Set(),
      stack: []
    };

    if (state.active.has(cellKey)) {
      throw new FormulaCycleError(state.stack.concat(cellKey));
    }

    state.active.add(cellKey);
    state.stack.push(cellKey);

    try {
      const compiledFormula = cell.getCompiledFormula();

      if (compiledFormula === undefined) {
        throw new Error('Formula cell is missing a compiled formula: ' + cellKey);
      }

      const value = this.formulaEvaluator.evaluate(compiledFormula, {
        evaluationState: state,
        functionRegistry: this.functionRegistry,
        workbook: this,
        worksheet: cell.worksheet
      });

      this.cellCache.set(cellKey, value);
      return value;
    } catch (error) {
      if (error instanceof FormulaCycleError || error instanceof FormulaEvaluationError) {
        throw error;
      }

      throw new FormulaEvaluationError(cellKey, toError(error));
    } finally {
      state.stack.pop();
      state.active.delete(cellKey);
    }
  }

  getCell(sheetName: string, cellName: string) {
    const sheet = this.getSheet(sheetName);

    if (sheet === undefined) {
      return undefined;
    }

    return sheet.getCell(cellName);
  }

  getCellKey(sheetName: string, cellName: string): string {
    return sheetName + '!' + cellName;
  }

  getCellValue(sheetName: string, cellName: string, evaluationState?: EvaluationState) {
    const sheet = this.getSheet(sheetName);

    if (sheet === undefined) {
      return ExcelError.REF;
    }

    const cell = sheet.getCell(cellName);

    if (cell === undefined || !cell.isFormula()) {
      return cell === undefined ? undefined : cell.getRawValue();
    }

    return this.evaluateFormulaCell(cell, evaluationState);
  }

  getFormulaSource(sheetName: string, cellName: string) {
    const sheet = this.getSheet(sheetName);

    if (sheet === undefined) {
      return undefined;
    }

    return sheet.getFormulaSource(cellName);
  }

  getDependents(sheetName: string, cellName: string): CellReferenceInfo[] {
    const cellKey = this.getCellKey(sheetName, cellName);
    const dependents = this.cellDependents.get(cellKey);

    if (dependents === undefined) {
      return [];
    }

    return Array.from(dependents).map(this.parseCellKey.bind(this));
  }

  getFunctionNames() {
    return this.functionRegistry.list();
  }

  getPrecedents(sheetName: string, cellName: string): CellReferenceInfo[] {
    const cellKey = this.getCellKey(sheetName, cellName);
    const precedents = this.cellDependencies.get(cellKey);

    if (precedents === undefined) {
      return [];
    }

    return Array.from(precedents).map(this.parseCellKey.bind(this));
  }

  getRawWorkbook() {
    return this.workbook;
  }

  getSheet(sheetName: string): Worksheet | undefined {
    return this.workbook[sheetName];
  }

  getSheetNames() {
    return Object.keys(this.workbook);
  }

  hasFunction(name: string): boolean {
    return this.functionRegistry.has(name);
  }

  invalidateAllFormulaCaches() {
    this.cellCache.clear();
  }

  isMutationBatchActive() {
    return this.mutationBatchDepth > 0;
  }

  invalidateCellAndDependents(sheetName: string, cellName: string) {
    const queue: string[] = [this.getCellKey(sheetName, cellName)];
    const visited = new Set<string>();

    while (queue.length > 0) {
      const currentCellKey = queue.pop();

      if (currentCellKey === undefined) {
        continue;
      }

      if (visited.has(currentCellKey)) {
        continue;
      }

      visited.add(currentCellKey);
      this.cellCache.delete(currentCellKey);

      const dependents = this.cellDependents.get(currentCellKey);

      if (dependents === undefined) {
        continue;
      }

      for (const dependentKey of dependents) {
        queue.push(dependentKey);
      }
    }
  }

  parseCellKey(cellKey: string): CellReferenceInfo {
    const separatorIndex = cellKey.indexOf('!');

    if (separatorIndex === -1) {
      throw new Error('Invalid cell key: ' + cellKey);
    }

    return {
      cellName: cellKey.slice(separatorIndex + 1),
      key: cellKey,
      sheetName: cellKey.slice(0, separatorIndex)
    };
  }

  registerFormulaCell(sheetName: string, cellName: string, compiledFormula: CompiledFormula) {
    const cellKey = this.getCellKey(sheetName, cellName);
    const dependencies = this.formulaEvaluator.collectReferences(compiledFormula);

    this.clearDependenciesForCell(cellKey);
    this.cellDependencies.set(cellKey, new Set());
    const dependencySet = this.cellDependencies.get(cellKey);

    if (dependencySet === undefined) {
      throw new Error('Missing dependency set for formula cell: ' + cellKey);
    }

    for (const dependency of dependencies) {
      const dependencyKey = this.getCellKey(dependency.sheet ?? sheetName, dependency.ref);

      dependencySet.add(dependencyKey);

      if (!this.cellDependents.has(dependencyKey)) {
        this.cellDependents.set(dependencyKey, new Set());
      }

      const dependentSet = this.cellDependents.get(dependencyKey);

      if (dependentSet === undefined) {
        throw new Error('Missing dependent set for cell: ' + dependencyKey);
      }

      dependentSet.add(cellKey);
    }

    if (this.isMutationBatchActive()) {
      this.pendingCacheReset = true;
      return;
    }

    this.invalidateCellAndDependents(sheetName, cellName);
  }

  registerFunction(name: string, implementation: FormulaFunction, options?: FormulaFunctionOptions) {
    this.functionRegistry.register(name, implementation, options);
    this.invalidateAllFormulaCaches();
    return this;
  }

  registerFunctions(functionMap: Record<string, FormulaFunction>, options?: FormulaFunctionOptions) {
    this.functionRegistry.registerMany(functionMap, options);
    this.invalidateAllFormulaCaches();
    return this;
  }

  registerValueCell(sheetName: string, cellName: string) {
    const cellKey = this.getCellKey(sheetName, cellName);

    this.clearDependenciesForCell(cellKey);

    if (this.isMutationBatchActive()) {
      this.pendingCacheReset = true;
      return;
    }

    this.invalidateCellAndDependents(sheetName, cellName);
  }

  requireSheet(sheetName: string): Worksheet {
    const sheet = this.getSheet(sheetName);

    if (sheet === undefined) {
      throw new Error('Unknown worksheet: ' + sheetName);
    }

    return sheet;
  }

  setCellFormula(sheetName: string, cellName: string, formulaExpression: string | CompiledFormula) {
    const sheet = this.requireSheet(sheetName);

    sheet.setCellFormula(cellName, formulaExpression);
    return this;
  }

  setCellValue(sheetName: string, cellName: string, cellValue: WorkbookCellValue) {
    const sheet = this.requireSheet(sheetName);

    sheet.setCellValue(cellName, cellValue);
    return this;
  }

  setFileVersion(fileVersion: string) {
    this.fileVersion = fileVersion;
    return this;
  }

  setType(type: string) {
    this.type = type;
    return this;
  }

  traceCell(sheetName: string, cellName: string, traceState?: EvaluationState): WorkbookTraceCell {
    const cellKey = this.getCellKey(sheetName, cellName);
    const sheet = this.getSheet(sheetName);

    if (sheet === undefined) {
      return {
        cellName: cellName,
        key: cellKey,
        kind: 'missing',
        precedents: [],
        sheetName: sheetName,
        value: ExcelError.REF
      };
    }

    const cell = sheet.getCell(cellName);

    if (cell === undefined || !cell.isFormula()) {
      return {
        cellName: cellName,
        key: cellKey,
        kind: 'value',
        precedents: [],
        rawValue: cell === undefined ? undefined : cell.getRawValue(),
        sheetName: sheetName,
        value: cell === undefined ? undefined : cell.getRawValue()
      };
    }

    const state = traceState || {
      active: new Set(),
      stack: []
    };

    if (state.active.has(cellKey)) {
      throw new FormulaCycleError(state.stack.concat(cellKey));
    }

    state.active.add(cellKey);
    state.stack.push(cellKey);

    try {
      const compiledFormula = cell.getCompiledFormula();

      if (compiledFormula === undefined) {
        throw new Error('Formula cell is missing a compiled formula: ' + cellKey);
      }

      const evaluation = this.formulaEvaluator.trace(compiledFormula, {
        evaluationState: state,
        functionRegistry: this.functionRegistry,
        traceState: state,
        workbook: this,
        worksheet: sheet
      }) as WorkbookTraceEvaluation;
      const traceValue = evaluation.value as WorkbookCellValue;

      this.cellCache.set(cellKey, traceValue);

      return {
        cellName: cellName,
        evaluation: evaluation,
        expression: compiledFormula.expression,
        key: cellKey,
        kind: 'formula',
        precedents: this.getPrecedents(sheetName, cellName),
        sheetName: sheetName,
        value: traceValue
      };
    } catch (error) {
      if (error instanceof FormulaCycleError || error instanceof FormulaEvaluationError) {
        throw error;
      }

      throw new FormulaEvaluationError(cellKey, toError(error));
    } finally {
      state.stack.pop();
      state.active.delete(cellKey);
    }
  }

  traverseDependents(sheetName: string, cellName: string): CellReferenceInfo[] {
    return this.traverseGraph(sheetName, cellName, this.cellDependents);
  }

  traverseGraph(sheetName: string, cellName: string, adjacencyMap: Map<string, Set<string>>): CellReferenceInfo[] {
    const originKey = this.getCellKey(sheetName, cellName);
    const visited = new Set<string>();
    const orderedKeys: string[] = [];
    const stack: string[] = Array.from(adjacencyMap.get(originKey) ?? []);

    while (stack.length > 0) {
      const currentKey = stack.pop();

      if (currentKey === undefined) {
        continue;
      }

      if (visited.has(currentKey)) {
        continue;
      }

      visited.add(currentKey);
      orderedKeys.push(currentKey);

      const nextKeys = adjacencyMap.get(currentKey);

      if (nextKeys === undefined) {
        continue;
      }

      for (const nextKey of nextKeys) {
        if (!visited.has(nextKey)) {
          stack.push(nextKey);
        }
      }
    }

    return orderedKeys.map(this.parseCellKey.bind(this));
  }

  traversePrecedents(sheetName: string, cellName: string): CellReferenceInfo[] {
    return this.traverseGraph(sheetName, cellName, this.cellDependencies);
  }

  unregisterFunction(name: string) {
    this.functionRegistry.unregister(name);
    this.invalidateAllFormulaCaches();
    return this;
  }

  zeroOutNullRows() {
    for (const sheet of Object.values(this.workbook)) {
      for (const row of sheet.rows) {
        if (row === null || row === undefined) {
          continue;
        }

        for (const cell of row) {
          if (cell === null || cell === undefined) {
            continue;
          }

          if (cell.getRawValue() === null) {
            sheet.setCellValue(cell.address, 0);
          }
        }
      }
    }
  }
}

export { Workbook };
