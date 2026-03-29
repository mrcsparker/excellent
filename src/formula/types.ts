'use strict';

import type { ExcelErrorCodeValue, ExcelErrorValue } from './errors';
import type { FormulaFunctionRegistry } from './registry';

type WorkbookScalarValue = Date | boolean | number | string | null | undefined | ExcelErrorValue;
type WorkbookCellValue = WorkbookScalarValue | WorkbookCellValue[];
type FormulaLiteralValue = boolean | number | string | null;
type FormulaBinaryOperator = '&' | '*' | '+' | '-' | '/' | '<' | '==' | '>' | '^';
type FormulaUnaryOperator = '+' | '-';
type FormulaFunctionImplementation = (...args: WorkbookCellValue[]) => unknown;
type FormulaRegistryOverrideOptions = {
  override?: boolean;
};
type FormulaIdentifierNode = {
  name: string;
  type: 'Identifier';
};
type FormulaLiteralNode = {
  type: 'Literal';
  value: FormulaLiteralValue;
};
type FormulaThisExpressionNode = {
  type: 'ThisExpression';
};
type FormulaArrayExpressionNode = {
  elements: FormulaAstNode[];
  type: 'ArrayExpression';
};
type FormulaBinaryExpressionNode = {
  left: FormulaAstNode;
  operator: FormulaBinaryOperator;
  right: FormulaAstNode;
  type: 'BinaryExpression';
};
type FormulaCallExpressionNode = {
  arguments: FormulaAstNode[];
  callee: FormulaAstNode;
  type: 'CallExpression';
};
type FormulaCellReferenceNode = {
  ref: string;
  sheet: string | null;
  type: 'CellReference';
};
type FormulaErrorLiteralNode = {
  code: ExcelErrorCodeValue;
  type: 'ErrorLiteral';
};
type FormulaFunctionCallExpressionNode = {
  arguments: FormulaAstNode[];
  name: string;
  type: 'FormulaCallExpression';
};
type FormulaMemberExpressionNode = {
  computed: boolean;
  object: FormulaAstNode;
  property: FormulaAstNode;
  type: 'MemberExpression';
};
type FormulaUnaryExpressionNode = {
  argument: FormulaAstNode;
  operator: FormulaUnaryOperator;
  type: 'UnaryExpression';
};
type FormulaAstNode =
  | FormulaArrayExpressionNode
  | FormulaBinaryExpressionNode
  | FormulaCallExpressionNode
  | FormulaCellReferenceNode
  | FormulaErrorLiteralNode
  | FormulaFunctionCallExpressionNode
  | FormulaIdentifierNode
  | FormulaLiteralNode
  | FormulaMemberExpressionNode
  | FormulaThisExpressionNode
  | FormulaUnaryExpressionNode;
type CompiledFormula = {
  ast: FormulaAstNode;
  expression: string;
};
type FormulaCellReference = {
  ref: string;
  sheet: string | null;
};
type FormulaFunction = FormulaFunctionImplementation;
type FormulaFunctionOptions = FormulaRegistryOverrideOptions;
type FormulaRegistryFunctions = Record<string, FormulaFunction>;
type FormulaFunctionMap = Record<string, FormulaFunctionImplementation>;
type FormulaFunctionNamespace = Readonly<FormulaFunctionMap>;
type FormulaRegistryOptions = {
  baseFunctions?: Record<string, FormulaFunctionImplementation>;
  functions?: FormulaRegistryFunctions;
};
type FormulaExpressionInput = CompiledFormula | FormulaAstNode | FormulaLiteralValue | string;
type FormulaExecutionState = {
  active: Set<string>;
  stack: string[];
};
type FormulaRuntimeWorkbook = {
  getCellValue(sheetName: string, cellName: string, evaluationState?: FormulaExecutionState): WorkbookCellValue;
  traceCell(sheetName: string, cellName: string, traceState?: FormulaExecutionState): {
    key: string;
    value: FormulaResolvedValue | undefined;
  };
};
type FormulaRuntimeWorksheet = {
  name: string;
};
type FormulaResolvedValue =
  | WorkbookCellValue
  | FormulaFunctionImplementation
  | FormulaFunctionNamespace
  | FormulaRuntimeWorkbook
  | FormulaRuntimeWorksheet
  | FormulaResolvedValue[]
  | Record<PropertyKey, unknown>;
type FormulaEvaluationRuntime = {
  evaluationState?: FormulaExecutionState;
  functionRegistry: FormulaFunctionRegistry;
  workbook: FormulaRuntimeWorkbook;
  worksheet: FormulaRuntimeWorksheet;
};
type FormulaTraceRuntime = {
  evaluationState?: FormulaExecutionState;
  functionRegistry: FormulaFunctionRegistry;
  traceState: FormulaExecutionState;
  workbook: FormulaRuntimeWorkbook;
  worksheet: FormulaRuntimeWorksheet;
};
type FormulaRuntime = FormulaEvaluationRuntime | FormulaTraceRuntime;
type FormulaTraceResult = {
  type: string;
  value: FormulaResolvedValue | undefined;
} & Record<string, unknown>;
type FormulaMemberTarget = {
  objectValue: FormulaResolvedValue | undefined;
  propertyValue: FormulaResolvedValue | undefined;
  value: FormulaResolvedValue | undefined;
};

export type {
  CompiledFormula,
  FormulaArrayExpressionNode,
  FormulaAstNode,
  FormulaBinaryExpressionNode,
  FormulaBinaryOperator,
  FormulaCallExpressionNode,
  FormulaCellReference,
  FormulaCellReferenceNode,
  FormulaErrorLiteralNode,
  FormulaEvaluationRuntime,
  FormulaExpressionInput,
  FormulaExecutionState,
  FormulaFunction,
  FormulaFunctionCallExpressionNode,
  FormulaFunctionImplementation,
  FormulaFunctionMap,
  FormulaFunctionNamespace,
  FormulaFunctionOptions,
  FormulaIdentifierNode,
  FormulaLiteralNode,
  FormulaLiteralValue,
  FormulaMemberExpressionNode,
  FormulaMemberTarget,
  FormulaRegistryFunctions,
  FormulaRegistryOptions,
  FormulaRegistryOverrideOptions,
  FormulaResolvedValue,
  FormulaRuntime,
  FormulaRuntimeWorkbook,
  FormulaRuntimeWorksheet,
  FormulaThisExpressionNode,
  FormulaTraceResult,
  FormulaTraceRuntime,
  FormulaUnaryExpressionNode,
  FormulaUnaryOperator,
  WorkbookCellValue,
  WorkbookScalarValue
};
