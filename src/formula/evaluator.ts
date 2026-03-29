'use strict';

import {
  collectCellReferences,
  compileFormula,
  extractReferenceFromJsMemberExpression,
  normalizeCellReference,
  serializeFormulaAst
} from './compiler';
import {
  ExcelError,
  ExcelErrorCode,
  type ExcelErrorValue,
  findExcelError,
  getExcelError,
  toExcelError
} from './errors';
import {
  assertSynchronousFunctionResult,
  normalizeFunctionName
} from './registry';
import type { FormulaFunctionRegistry } from './registry';
import type {
  CompiledFormula,
  FormulaAstNode,
  FormulaBinaryOperator,
  FormulaCallExpressionNode,
  FormulaCellReference,
  FormulaEvaluationRuntime,
  FormulaExpressionInput,
  FormulaFunctionImplementation,
  FormulaMemberExpressionNode,
  FormulaMemberTarget,
  FormulaResolvedValue,
  FormulaRuntime,
  FormulaTraceResult,
  FormulaTraceRuntime,
  FormulaUnaryOperator,
  WorkbookCellValue
} from './types';

function normalizeFormulaResult(value: unknown): FormulaResolvedValue | undefined {
  const excelError = toExcelError(value);

  if (excelError !== null) {
    return excelError;
  }

  if (!Array.isArray(value)) {
    return value as FormulaResolvedValue | undefined;
  }

  return value.map(normalizeFormulaResult);
}

function coerceToNumber(value: unknown): number | ExcelErrorValue {
  const excelError = findExcelError(value);

  if (excelError !== null) {
    return excelError;
  }

  if (Array.isArray(value)) {
    return ExcelError.VALUE;
  }

  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'boolean') {
    return value ? 1 : 0;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : ExcelError.VALUE;
  }

  if (typeof value === 'string') {
    const numericValue = Number(value.trim());

    if (!Number.isNaN(numericValue)) {
      return numericValue;
    }

    return ExcelError.VALUE;
  }

  return ExcelError.VALUE;
}

function coerceToComparableValue(value: unknown): Date | boolean | number | string | ExcelErrorValue {
  const excelError = findExcelError(value);

  if (excelError !== null) {
    return excelError;
  }

  if (Array.isArray(value)) {
    return ExcelError.VALUE;
  }

  if (value === null || value === undefined) {
    return 0;
  }

  if (
    value instanceof Date ||
    typeof value === 'boolean' ||
    typeof value === 'number' ||
    typeof value === 'string'
  ) {
    return value;
  }

  return ExcelError.VALUE;
}

function coerceToConcatenatedValue(value: unknown): string | ExcelErrorValue {
  const excelError = findExcelError(value);

  if (excelError !== null) {
    return excelError;
  }

  if (Array.isArray(value)) {
    return ExcelError.VALUE;
  }

  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }

  if (value instanceof Date || typeof value === 'number' || typeof value === 'string') {
    return String(value);
  }

  return ExcelError.VALUE;
}

function coerceToConditionValue(value: unknown): boolean | ExcelErrorValue {
  const excelError = findExcelError(value);

  if (excelError !== null) {
    return excelError;
  }

  if (Array.isArray(value)) {
    return ExcelError.VALUE;
  }

  if (value === null || value === undefined || value === '') {
    return false;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  if (typeof value === 'string') {
    return value.length > 0;
  }

  return Boolean(value);
}

function invokeFormulaFunction(
  registry: FormulaFunctionRegistry,
  name: string,
  argumentsValues: WorkbookCellValue[]
): FormulaResolvedValue | undefined {
  try {
    return normalizeFormulaResult(registry.invoke(name, argumentsValues));
  } catch (error) {
    const excelError = toExcelError(error);

    if (excelError !== null) {
      return excelError;
    }

    throw error;
  }
}

function evaluateBinaryExpression(
  operator: FormulaBinaryOperator,
  leftValue: unknown,
  rightValue: unknown
): WorkbookCellValue {
  const excelError = findExcelError([leftValue, rightValue]);

  if (excelError !== null) {
    return excelError;
  }

  switch (operator) {
    case '&': {
      const coercedLeft = coerceToConcatenatedValue(leftValue);
      const coercedRight = coerceToConcatenatedValue(rightValue);
      const textError = findExcelError([coercedLeft, coercedRight]);

      if (textError !== null) {
        return textError;
      }

      return String(coercedLeft) + String(coercedRight);
    }
    case '+':
    case '-':
    case '*':
    case '/':
    case '^': {
      const coercedLeft = coerceToNumber(leftValue);
      const coercedRight = coerceToNumber(rightValue);
      const numericError = findExcelError([coercedLeft, coercedRight]);

      if (numericError !== null) {
        return numericError;
      }

      const leftNumber = coercedLeft as number;
      const rightNumber = coercedRight as number;

      if (operator === '+') {
        return leftNumber + rightNumber;
      }

      if (operator === '-') {
        return leftNumber - rightNumber;
      }

      if (operator === '*') {
        return leftNumber * rightNumber;
      }

      if (operator === '/') {
        if (rightNumber === 0) {
          return ExcelError.DIV0;
        }

        return leftNumber / rightNumber;
      }

      return Math.pow(leftNumber, rightNumber);
    }
    case '<':
    case '>':
    case '==':
    {
      const comparableLeft = coerceToComparableValue(leftValue);
      const comparableRight = coerceToComparableValue(rightValue);
      const comparisonError = findExcelError([comparableLeft, comparableRight]);

      if (comparisonError !== null) {
        return comparisonError;
      }

      const normalizedLeft = comparableLeft as Date | boolean | number | string;
      const normalizedRight = comparableRight as Date | boolean | number | string;

      if (operator === '<') {
        return normalizedLeft < normalizedRight;
      }

      if (operator === '>') {
        return normalizedLeft > normalizedRight;
      }

      // Intentionally mirror the historical generated-expression semantics.
      return normalizedLeft == normalizedRight;
    }
    default:
      throw new Error('Unsupported binary operator: ' + operator);
  }
}

function evaluateUnaryExpression(operator: FormulaUnaryOperator, value: unknown): WorkbookCellValue {
  const coercedValue = coerceToNumber(value);
  const excelError = findExcelError(coercedValue);

  if (excelError !== null) {
    return excelError;
  }

  const numericValue = coercedValue as number;

  switch (operator) {
    case '-':
      return -numericValue;
    case '+':
      return +numericValue;
    default:
      throw new Error('Unsupported unary operator: ' + operator);
  }
}

function createLiteralTrace(value: FormulaResolvedValue | undefined): FormulaTraceResult {
  return {
    type: 'literal',
    value: value
  };
}

function extractFormulaNamespaceMemberCall(node: FormulaAstNode): string | null {
  if (
    node.type === 'CallExpression' &&
    node.callee.type === 'MemberExpression' &&
    !node.callee.computed &&
    node.callee.object.type === 'Identifier' &&
    node.callee.object.name === 'Formula' &&
    node.callee.property.type === 'Identifier'
  ) {
    return node.callee.property.name;
  }

  return null;
}

function evaluateFormulaFunctionCall(
  functionName: string,
  argumentNodes: FormulaAstNode[],
  runtime: FormulaEvaluationRuntime
): FormulaResolvedValue | undefined {
  const normalizedFunctionName = normalizeFunctionName(functionName);

  if (normalizedFunctionName === 'IF') {
    const conditionValue = argumentNodes[0] === undefined ? false : evaluateFormulaNode(argumentNodes[0], runtime);
    const coercedCondition = coerceToConditionValue(conditionValue);
    const conditionError = findExcelError(coercedCondition);

    if (conditionError !== null) {
      return conditionError;
    }

    if (coercedCondition) {
      return argumentNodes[1] === undefined ? true : evaluateFormulaNode(argumentNodes[1], runtime);
    }

    return argumentNodes[2] === undefined ? false : evaluateFormulaNode(argumentNodes[2], runtime);
  }

  if (normalizedFunctionName === 'IFERROR') {
    const candidateValue = argumentNodes[0] === undefined ? undefined : evaluateFormulaNode(argumentNodes[0], runtime);

    if (findExcelError(candidateValue) !== null) {
      return argumentNodes[1] === undefined ? '' : evaluateFormulaNode(argumentNodes[1], runtime);
    }

    return candidateValue;
  }

  if (normalizedFunctionName === 'IFNA') {
    const candidateValue = argumentNodes[0] === undefined ? undefined : evaluateFormulaNode(argumentNodes[0], runtime);
    const candidateError = findExcelError(candidateValue);

    if (candidateError !== null && candidateError.code === ExcelErrorCode.NA) {
      return argumentNodes[1] === undefined ? '' : evaluateFormulaNode(argumentNodes[1], runtime);
    }

    return candidateValue;
  }

  const argumentsValues = argumentNodes.map(function(argument) {
    return evaluateFormulaNode(argument, runtime);
  });
  const argumentError = findExcelError(argumentsValues);

  if (argumentError !== null) {
    return argumentError;
  }

  return invokeFormulaFunction(runtime.functionRegistry, normalizedFunctionName, argumentsValues as WorkbookCellValue[]);
}

function traceFormulaFunctionCall(
  functionName: string,
  argumentNodes: FormulaAstNode[],
  runtime: FormulaTraceRuntime
): FormulaTraceResult {
  const normalizedFunctionName = normalizeFunctionName(functionName);

  if (normalizedFunctionName === 'IF') {
    const conditionTrace = argumentNodes[0] === undefined ? createLiteralTrace(false) : traceFormulaNode(argumentNodes[0], runtime);
    const coercedCondition = coerceToConditionValue(conditionTrace.value);
    const conditionError = findExcelError(coercedCondition);

    if (conditionError !== null) {
      return {
        arguments: [conditionTrace],
        name: normalizedFunctionName,
        type: 'formula-call',
        value: conditionError
      };
    }

    if (coercedCondition) {
      const trueTrace = argumentNodes[1] === undefined ? createLiteralTrace(true) : traceFormulaNode(argumentNodes[1], runtime);

      return {
        arguments: [conditionTrace, trueTrace],
        name: normalizedFunctionName,
        type: 'formula-call',
        value: trueTrace.value
      };
    }

    const falseTrace = argumentNodes[2] === undefined ? createLiteralTrace(false) : traceFormulaNode(argumentNodes[2], runtime);

    return {
      arguments: [conditionTrace, falseTrace],
      name: normalizedFunctionName,
      type: 'formula-call',
      value: falseTrace.value
    };
  }

  if (normalizedFunctionName === 'IFERROR') {
    const candidateTrace = argumentNodes[0] === undefined ? createLiteralTrace(undefined) : traceFormulaNode(argumentNodes[0], runtime);

    if (findExcelError(candidateTrace.value) !== null) {
      const fallbackTrace = argumentNodes[1] === undefined ? createLiteralTrace('') : traceFormulaNode(argumentNodes[1], runtime);

      return {
        arguments: [candidateTrace, fallbackTrace],
        name: normalizedFunctionName,
        type: 'formula-call',
        value: fallbackTrace.value
      };
    }

    return {
      arguments: [candidateTrace],
      name: normalizedFunctionName,
      type: 'formula-call',
      value: candidateTrace.value
    };
  }

  if (normalizedFunctionName === 'IFNA') {
    const candidateTrace = argumentNodes[0] === undefined ? createLiteralTrace(undefined) : traceFormulaNode(argumentNodes[0], runtime);
    const candidateError = findExcelError(candidateTrace.value);

    if (candidateError !== null && candidateError.code === ExcelErrorCode.NA) {
      const fallbackTrace = argumentNodes[1] === undefined ? createLiteralTrace('') : traceFormulaNode(argumentNodes[1], runtime);

      return {
        arguments: [candidateTrace, fallbackTrace],
        name: normalizedFunctionName,
        type: 'formula-call',
        value: fallbackTrace.value
      };
    }

    return {
      arguments: [candidateTrace],
      name: normalizedFunctionName,
      type: 'formula-call',
      value: candidateTrace.value
    };
  }

  const argumentTraces = argumentNodes.map(function(argument) {
    return traceFormulaNode(argument, runtime);
  });
  const argumentError = findExcelError(argumentTraces.map(function(argumentTrace) {
    return argumentTrace.value;
  }));

  if (argumentError !== null) {
    return {
      arguments: argumentTraces,
      name: normalizedFunctionName,
      type: 'formula-call',
      value: argumentError
    };
  }

  return {
    arguments: argumentTraces,
    name: normalizedFunctionName,
    type: 'formula-call',
    value: invokeFormulaFunction(
      runtime.functionRegistry,
      normalizedFunctionName,
      argumentTraces.map(function(argumentTrace) {
        return argumentTrace.value as WorkbookCellValue;
      })
    )
  };
}

function toPropertyKey(value: FormulaResolvedValue | undefined): PropertyKey | null {
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'symbol') {
    return value;
  }

  return null;
}

function getIndexedValue(
  objectValue: FormulaResolvedValue | undefined,
  propertyValue: FormulaResolvedValue | undefined
): FormulaResolvedValue | undefined {
  if (objectValue === null || objectValue === undefined) {
    return undefined;
  }

  const propertyKey = toPropertyKey(propertyValue);

  if (propertyKey === null) {
    return undefined;
  }

  return (objectValue as Record<PropertyKey, FormulaResolvedValue | undefined>)[propertyKey];
}

function formatPropertyValue(propertyValue: FormulaResolvedValue | undefined): string {
  if (typeof propertyValue === 'string' || typeof propertyValue === 'number') {
    return String(propertyValue);
  }

  return '[computed]';
}

function getMemberTarget(node: FormulaMemberExpressionNode, runtime: FormulaRuntime): FormulaMemberTarget {
  const objectValue = evaluateFormulaNode(node.object, runtime);
  const propertyValue = node.computed
    ? evaluateFormulaNode(node.property, runtime)
    : node.property.type === 'Identifier'
      ? node.property.name
      : undefined;

  return {
    objectValue,
    propertyValue,
    value: getIndexedValue(objectValue, propertyValue)
  };
}

function evaluateCallExpression(node: FormulaCallExpressionNode, runtime: FormulaEvaluationRuntime): FormulaResolvedValue | undefined {
  const formulaFunctionName = extractFormulaNamespaceMemberCall(node);

  if (formulaFunctionName !== null) {
    return evaluateFormulaFunctionCall(formulaFunctionName, node.arguments, runtime);
  }

  const argumentsValues = node.arguments.map(function(argument) {
    return evaluateFormulaNode(argument, runtime);
  });
  const argumentError = findExcelError(argumentsValues);

  if (argumentError !== null) {
    return argumentError;
  }

  if (node.callee.type === 'MemberExpression') {
    const target = getMemberTarget(node.callee, runtime);
    const targetError = findExcelError([target.objectValue, target.propertyValue]);

    if (targetError !== null) {
      return targetError;
    }

    if (typeof target.value !== 'function') {
      throw new Error('Expected callable member: ' + formatPropertyValue(target.propertyValue));
    }

    const memberFunction = target.value as FormulaFunctionImplementation;

    return normalizeFormulaResult(assertSynchronousFunctionResult(
      getTraceCallName(node.callee),
      memberFunction.apply(target.objectValue, argumentsValues as WorkbookCellValue[])
    ));
  }

  const callee = evaluateFormulaNode(node.callee, runtime);
  const calleeError = findExcelError(callee);

  if (calleeError !== null) {
    return calleeError;
  }

  if (typeof callee !== 'function') {
    throw new Error('Expected callable expression.');
  }

  return normalizeFormulaResult(assertSynchronousFunctionResult(
    getTraceCallName(node.callee),
    callee(...(argumentsValues as WorkbookCellValue[]))
  ));
}

function getTraceCallName(node: FormulaAstNode): string {
  if (node.type === 'MemberExpression' && !node.computed && node.property.type === 'Identifier') {
    return node.property.name;
  }

  if (node.type === 'Identifier') {
    return node.name;
  }

  return 'call';
}

function traceCellReference(sheetName: string, cellName: string, runtime: FormulaTraceRuntime): FormulaTraceResult {
  const normalizedCellName = normalizeCellReference(cellName);
  const cellTrace = runtime.workbook.traceCell(sheetName, normalizedCellName, runtime.traceState);

  return {
    cell: cellTrace,
    cellName: normalizedCellName,
    key: cellTrace.key,
    sheetName: sheetName,
    type: 'cell-reference',
    value: cellTrace.value
  };
}

function traceCallExpression(node: FormulaCallExpressionNode, runtime: FormulaTraceRuntime): FormulaTraceResult {
  const formulaFunctionName = extractFormulaNamespaceMemberCall(node);

  if (formulaFunctionName !== null) {
    return traceFormulaFunctionCall(formulaFunctionName, node.arguments, runtime);
  }

  const argumentTraces = node.arguments.map(function(argument) {
    return traceFormulaNode(argument, runtime);
  });
  const argumentError = findExcelError(argumentTraces.map(function(argumentTrace) {
    return argumentTrace.value;
  }));

  if (argumentError !== null) {
    return {
      arguments: argumentTraces,
      callee: getTraceCallName(node.callee),
      type: 'call-expression',
      value: argumentError
    };
  }

  if (node.callee.type === 'MemberExpression') {
    const target = getMemberTarget(node.callee, runtime);
    const targetError = findExcelError([target.objectValue, target.propertyValue]);

    if (targetError !== null) {
      return {
        arguments: argumentTraces,
        callee: getTraceCallName(node.callee),
        type: 'call-expression',
        value: targetError
      };
    }

    if (typeof target.value !== 'function') {
      throw new Error('Expected callable member: ' + formatPropertyValue(target.propertyValue));
    }

    const memberFunction = target.value as FormulaFunctionImplementation;

    return {
      arguments: argumentTraces,
      callee: getTraceCallName(node.callee),
      type: 'call-expression',
      value: normalizeFormulaResult(assertSynchronousFunctionResult(
        getTraceCallName(node.callee),
        memberFunction.apply(target.objectValue, argumentTraces.map(function(argumentTrace) {
          return argumentTrace.value as WorkbookCellValue;
        }))
      ))
      };
  }

  const callee = evaluateFormulaNode(node.callee, runtime);
  const calleeError = findExcelError(callee);

  if (calleeError !== null) {
    return {
      arguments: argumentTraces,
      callee: getTraceCallName(node.callee),
      type: 'call-expression',
      value: calleeError
    };
  }

  if (typeof callee !== 'function') {
    throw new Error('Expected callable expression.');
  }

  return {
    arguments: argumentTraces,
    callee: getTraceCallName(node.callee),
    type: 'call-expression',
    value: normalizeFormulaResult(assertSynchronousFunctionResult(
      getTraceCallName(node.callee),
      callee(...argumentTraces.map(function(argumentTrace) {
        return argumentTrace.value as WorkbookCellValue;
      }))
    ))
  };
}

function traceFormulaNode(node: FormulaAstNode, runtime: FormulaTraceRuntime): FormulaTraceResult {
  switch (node.type) {
    case 'ArrayExpression': {
      const elementTraces = node.elements.map(function(element) {
        return traceFormulaNode(element, runtime);
      });

      return {
        elements: elementTraces,
        type: 'array-expression',
        value: elementTraces.map(function(elementTrace) {
          return elementTrace.value;
        })
      };
    }
    case 'BinaryExpression': {
      const leftTrace = traceFormulaNode(node.left, runtime);
      const rightTrace = traceFormulaNode(node.right, runtime);

      return {
        left: leftTrace,
        operator: node.operator,
        right: rightTrace,
        type: 'binary-expression',
        value: evaluateBinaryExpression(node.operator, leftTrace.value, rightTrace.value)
      };
    }
    case 'CellReference':
      return traceCellReference(node.sheet || runtime.worksheet.name, node.ref, runtime);
    case 'CallExpression':
      return traceCallExpression(node, runtime);
    case 'ErrorLiteral':
      return {
        code: node.code,
        type: 'error-literal',
        value: getExcelError(node.code)
      };
    case 'FormulaCallExpression': {
      return traceFormulaFunctionCall(node.name, node.arguments, runtime);
    }
    case 'Identifier':
      return {
        name: node.name,
        type: 'identifier',
        value: evaluateFormulaNode(node, runtime)
      };
    case 'Literal':
      return {
        type: 'literal',
        value: node.value
      };
    case 'MemberExpression': {
      const extractedReference = extractReferenceFromJsMemberExpression(node);

      if (extractedReference !== null) {
        return traceCellReference(extractedReference.sheet || runtime.worksheet.name, extractedReference.ref, runtime);
      }

      const target = getMemberTarget(node, runtime);
      const targetError = findExcelError([target.objectValue, target.propertyValue]);

      if (targetError !== null) {
        return {
          property: target.propertyValue,
          type: 'member-expression',
          value: targetError
        };
      }

      return {
        property: target.propertyValue,
        type: 'member-expression',
        value: normalizeFormulaResult(target.value)
      };
    }
    case 'ThisExpression':
      return {
        type: 'this-expression',
        value: runtime.worksheet
      };
    case 'UnaryExpression': {
      const argumentTrace = traceFormulaNode(node.argument, runtime);

      return {
        argument: argumentTrace,
        operator: node.operator,
        type: 'unary-expression',
        value: evaluateUnaryExpression(node.operator, argumentTrace.value)
      };
    }
    default:
      throw new Error('Unsupported formula node type.');
  }
}

function evaluateFormulaNode(node: FormulaAstNode, runtime: FormulaRuntime): FormulaResolvedValue | undefined {
  switch (node.type) {
    case 'ArrayExpression':
      return node.elements.map(function(element) {
        return evaluateFormulaNode(element, runtime);
      });
    case 'BinaryExpression':
      return evaluateBinaryExpression(
        node.operator,
        evaluateFormulaNode(node.left, runtime),
        evaluateFormulaNode(node.right, runtime)
      );
    case 'CellReference':
      return runtime.workbook.getCellValue(
        node.sheet || runtime.worksheet.name,
        normalizeCellReference(node.ref),
        runtime.evaluationState
      );
    case 'CallExpression':
      return evaluateCallExpression(node, runtime);
    case 'ErrorLiteral':
      return getExcelError(node.code);
    case 'FormulaCallExpression':
      return evaluateFormulaFunctionCall(node.name, node.arguments, runtime);
    case 'Identifier':
      if (node.name === 'Formula') {
        return runtime.functionRegistry.getNamespace();
      }

      if (node.name === 'self') {
        return runtime.workbook;
      }

      return ExcelError.NAME;
    case 'Literal':
      return node.value;
    case 'MemberExpression': {
      const reference = extractReferenceFromJsMemberExpression(node);

      if (reference !== null) {
        return runtime.workbook.getCellValue(
          reference.sheet ?? runtime.worksheet.name,
          normalizeCellReference(reference.ref),
          runtime.evaluationState
        );
      }

      const target = getMemberTarget(node, runtime);
      const targetError = findExcelError([target.objectValue, target.propertyValue]);

      if (targetError !== null) {
        return targetError;
      }

      return normalizeFormulaResult(target.value);
    }
    case 'ThisExpression':
      return runtime.worksheet;
    case 'UnaryExpression':
      return evaluateUnaryExpression(node.operator, evaluateFormulaNode(node.argument, runtime));
    default:
      throw new Error('Unsupported formula node type.');
  }
}

function evaluateCompiledFormula(compiledFormula: CompiledFormula, runtime: FormulaEvaluationRuntime): WorkbookCellValue {
  return evaluateFormulaNode(compiledFormula.ast, runtime) as WorkbookCellValue;
}

function traceCompiledFormula(compiledFormula: CompiledFormula, runtime: FormulaTraceRuntime): FormulaTraceResult {
  return traceFormulaNode(compiledFormula.ast, runtime);
}

class FormulaEvaluator {
  compile(expression: FormulaExpressionInput): CompiledFormula {
    return compileFormula(expression);
  }

  collectReferences(expression: FormulaExpressionInput): FormulaCellReference[] {
    return collectCellReferences(this.compile(expression).ast);
  }

  evaluate(expression: FormulaExpressionInput, runtime: FormulaEvaluationRuntime): WorkbookCellValue {
    return evaluateCompiledFormula(this.compile(expression), runtime);
  }

  trace(expression: FormulaExpressionInput, runtime: FormulaTraceRuntime): FormulaTraceResult {
    return traceCompiledFormula(this.compile(expression), runtime);
  }

  serialize(expression: FormulaExpressionInput): string {
    const compiledFormula = this.compile(expression);

    try {
      return serializeFormulaAst(compiledFormula.ast);
    } catch {
      return compiledFormula.expression;
    }
  }
}

export {
  FormulaEvaluator,
  evaluateCompiledFormula,
  traceCompiledFormula
};
