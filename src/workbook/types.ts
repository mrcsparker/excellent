'use strict';

import type {
  FormulaEvaluator,
  FormulaFunction,
  FormulaFunctionRegistry,
  WorkbookCellValue
} from '../formula';
import type { Cell } from './cell';
import type { Worksheet } from './worksheet';

type EvaluationState = {
  active: Set<string>;
  stack: string[];
};
type WorkbookOptions = {
  formulaEvaluator?: FormulaEvaluator;
  functionRegistry?: FormulaFunctionRegistry;
  functions?: Record<string, FormulaFunction>;
};
type CellReferenceInfo = {
  cellName: string;
  key: string;
  sheetName: string;
};
type WorkbookRow = Array<Cell | undefined>;
type WorkbookRows = Array<WorkbookRow | undefined>;
type WorkbookSheets = Record<string, Worksheet>;
type WorkbookHandle = {
  workbook: WorkbookSheets;
};
type WorkbookTraceEvaluation = {
  [key: string]: unknown;
  type: string;
  value: unknown;
};
type WorkbookTraceCell = {
  cellName: string;
  expression?: string;
  evaluation?: WorkbookTraceEvaluation;
  key: string;
  kind: 'formula' | 'missing' | 'value';
  precedents: CellReferenceInfo[];
  rawValue?: WorkbookCellValue;
  sheetName: string;
  value: WorkbookCellValue;
};

export type {
  CellReferenceInfo,
  EvaluationState,
  WorkbookHandle,
  WorkbookOptions,
  WorkbookRow,
  WorkbookRows,
  WorkbookSheets,
  WorkbookTraceCell,
  WorkbookTraceEvaluation
};
