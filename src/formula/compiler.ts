'use strict';

import * as acorn from 'acorn';
import type {
  CompiledFormula,
  FormulaAstNode,
  FormulaBinaryExpressionNode,
  FormulaCellReference,
  FormulaExpressionInput
} from './types';

function isFormulaAstNode(value: unknown): value is FormulaAstNode {
  return value !== null &&
    typeof value === 'object' &&
    typeof (value as { type?: unknown }).type === 'string';
}

function isCompiledFormula(value: unknown): value is CompiledFormula {
  return value !== null &&
    typeof value === 'object' &&
    'ast' in value &&
    isFormulaAstNode((value as { ast?: unknown }).ast) &&
    typeof (value as { expression?: unknown }).expression === 'string';
}

function normalizeCellReference(ref: string): string;
function normalizeCellReference<T>(ref: T): T;
function normalizeCellReference(ref: unknown) {
  if (typeof ref !== 'string') {
    return ref;
  }

  return ref.replace(/\$/g, '');
}

function compileFormula(expression: FormulaExpressionInput): CompiledFormula {
  if (isCompiledFormula(expression)) {
    return {
      ast: expression.ast,
      expression: expression.expression || serializeFormulaAst(expression.ast)
    };
  }

  if (isFormulaAstNode(expression)) {
    return {
      ast: expression,
      expression: serializeFormulaAst(expression)
    };
  }

  if (typeof expression !== 'string') {
    return {
      ast: {
        type: 'Literal',
        value: expression
      },
      expression: String(expression)
    };
  }

  return {
    ast: acorn.parseExpressionAt(expression, 0, {
      ecmaVersion: 'latest'
    }) as FormulaAstNode,
    expression: expression
  };
}

function getFormulaAstPrecedence(node: FormulaAstNode): number {
  switch (node.type) {
    case 'BinaryExpression':
      switch (node.operator) {
        case '==':
        case '<':
        case '>':
          return 1;
        case '&':
          return 2;
        case '+':
        case '-':
          return 3;
        case '*':
        case '/':
          return 4;
        case '^':
          return 5;
        default:
          return 0;
      }
    case 'UnaryExpression':
      return 6;
    case 'FormulaCallExpression':
    case 'ArrayExpression':
    case 'CellReference':
    case 'ErrorLiteral':
    case 'Literal':
      return 7;
    default:
      return 7;
  }
}

function needsParentheses(childNode: FormulaAstNode, parentNode: FormulaBinaryExpressionNode, side: 'left' | 'right'): boolean {
  const childPrecedence = getFormulaAstPrecedence(childNode);
  const parentPrecedence = getFormulaAstPrecedence(parentNode);

  if (childPrecedence < parentPrecedence) {
    return true;
  }

  if (childNode.type !== 'BinaryExpression' || childPrecedence !== parentPrecedence || side !== 'right') {
    return false;
  }

  return parentNode.operator === '-' ||
    parentNode.operator === '/' ||
    parentNode.operator === '^' ||
    parentNode.operator === '==' ||
    parentNode.operator === '<' ||
    parentNode.operator === '>' ||
    parentNode.operator === '&';
}

function needsUnaryDisambiguation(childNode: FormulaAstNode, parentNode: FormulaBinaryExpressionNode): boolean {
  if (childNode.type !== 'UnaryExpression') {
    return false;
  }

  if (
    (parentNode.operator === '+' || parentNode.operator === '-') &&
    childNode.operator === parentNode.operator
  ) {
    return true;
  }

  return parentNode.operator === '&' && childNode.operator === '+';
}

function serializeBinaryExpression(node: FormulaBinaryExpressionNode): string {
  const leftExpression = serializeFormulaAst(node.left);
  const rightExpression = serializeFormulaAst(node.right);
  const wrappedLeftExpression = needsParentheses(node.left, node, 'left') ? '(' + leftExpression + ')' : leftExpression;
  const wrappedRightExpression = needsParentheses(node.right, node, 'right') || needsUnaryDisambiguation(node.right, node)
    ? '(' + rightExpression + ')'
    : rightExpression;

  if (node.operator === '&') {
    return wrappedLeftExpression + '+""+' + wrappedRightExpression;
  }

  return wrappedLeftExpression + node.operator + wrappedRightExpression;
}

function serializeFormulaAst(node: FormulaAstNode): string {
  switch (node.type) {
    case 'ArrayExpression':
      return '[' + node.elements.map(serializeFormulaAst).join(',') + ']';
    case 'BinaryExpression':
      return serializeBinaryExpression(node);
    case 'CellReference':
      if (node.sheet !== null && node.sheet !== undefined) {
        return "self.workbook['" + node.sheet + "']." + node.ref;
      }

      return 'this.' + node.ref;
    case 'ErrorLiteral':
      return node.code;
    case 'FormulaCallExpression':
      return 'Formula.' + node.name + '(' + node.arguments.map(serializeFormulaAst).join(',') + ')';
    case 'Literal':
      if (typeof node.value === 'string') {
        return JSON.stringify(node.value);
      }

      return String(node.value);
    case 'UnaryExpression':
      if (node.argument.type === 'BinaryExpression' || node.argument.type === 'UnaryExpression') {
        return node.operator + '(' + serializeFormulaAst(node.argument) + ')';
      }

      return node.operator + serializeFormulaAst(node.argument);
    default:
      throw new Error('Unsupported formula AST node type: ' + node.type);
  }
}

function extractReferenceFromJsMemberExpression(node: FormulaAstNode): FormulaCellReference | null {
  if (node.type !== 'MemberExpression' || node.computed) {
    return null;
  }

  const memberNode = node;
  const objectNode = memberNode.object;
  const thisPropertyNode = memberNode.property;

  if (objectNode.type === 'ThisExpression' && thisPropertyNode.type === 'Identifier') {
    return {
      ref: thisPropertyNode.name,
      sheet: null
    };
  }

  if (objectNode.type !== 'MemberExpression') {
    return null;
  }

  const workbookMemberNode = objectNode;
  const workbookRootNode = workbookMemberNode.object;
  const sheetLiteralNode = workbookMemberNode.property;
  const workbookCellPropertyNode = memberNode.property;

  if (workbookRootNode.type !== 'MemberExpression') {
    return null;
  }

  const workbookIdentifierNode = workbookRootNode.object;
  const workbookPropertyNode = workbookRootNode.property;

  if (
    workbookMemberNode.computed &&
    !workbookRootNode.computed &&
    workbookIdentifierNode.type === 'Identifier' &&
    workbookIdentifierNode.name === 'self' &&
    workbookPropertyNode.type === 'Identifier' &&
    workbookPropertyNode.name === 'workbook' &&
    sheetLiteralNode.type === 'Literal' &&
    workbookCellPropertyNode.type === 'Identifier' &&
    typeof sheetLiteralNode.value === 'string'
  ) {
    return {
      ref: workbookCellPropertyNode.name,
      sheet: sheetLiteralNode.value
    };
  }

  return null;
}

function collectCellReferences(node: FormulaAstNode, references: FormulaCellReference[] = []): FormulaCellReference[] {
  const collectedReferences = references;
  const extractedReference = extractReferenceFromJsMemberExpression(node);

  if (extractedReference !== null) {
    collectedReferences.push({
      ref: normalizeCellReference(extractedReference.ref),
      sheet: extractedReference.sheet
    });
    return collectedReferences;
  }

  switch (node.type) {
    case 'ArrayExpression':
      for (const element of node.elements) {
        collectCellReferences(element, collectedReferences);
      }
      break;
    case 'BinaryExpression':
      collectCellReferences(node.left, collectedReferences);
      collectCellReferences(node.right, collectedReferences);
      break;
    case 'CallExpression':
      collectCellReferences(node.callee, collectedReferences);
      for (const argument of node.arguments) {
        collectCellReferences(argument, collectedReferences);
      }
      break;
    case 'CellReference':
      collectedReferences.push({
        ref: normalizeCellReference(node.ref),
        sheet: node.sheet ?? null
      });
      break;
    case 'FormulaCallExpression':
      for (const argument of node.arguments) {
        collectCellReferences(argument, collectedReferences);
      }
      break;
    case 'MemberExpression':
      collectCellReferences(node.object, collectedReferences);
      if (node.computed) {
        collectCellReferences(node.property, collectedReferences);
      }
      break;
    case 'UnaryExpression':
      collectCellReferences(node.argument, collectedReferences);
      break;
    default:
      break;
  }

  return collectedReferences;
}

export {
  collectCellReferences,
  compileFormula,
  extractReferenceFromJsMemberExpression,
  normalizeCellReference,
  serializeFormulaAst
};
