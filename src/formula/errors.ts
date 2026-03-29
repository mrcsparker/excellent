'use strict';

const ExcelErrorCode = Object.freeze({
  DIV0: '#DIV/0!',
  NAME: '#NAME?',
  NA: '#N/A',
  REF: '#REF!',
  VALUE: '#VALUE!'
});

type ExcelErrorCodeValue = typeof ExcelErrorCode[keyof typeof ExcelErrorCode];

class ExcelErrorValue extends Error {
  code: ExcelErrorCodeValue;
  isExcelError: true;

  constructor(code: ExcelErrorCodeValue) {
    super(code);
    this.name = 'ExcelErrorValue';
    this.code = code;
    this.isExcelError = true;
  }

  toJSON(): ExcelErrorCodeValue {
    return this.code;
  }

  override toString(): ExcelErrorCodeValue {
    return this.code;
  }
}

const excelErrorValues = new Map<ExcelErrorCodeValue, ExcelErrorValue>(Object.values(ExcelErrorCode).map(function createExcelErrorValue(code) {
  return [code, Object.freeze(new ExcelErrorValue(code))];
}));

const ExcelError = Object.freeze({
  DIV0: getExcelError(ExcelErrorCode.DIV0),
  NAME: getExcelError(ExcelErrorCode.NAME),
  NA: getExcelError(ExcelErrorCode.NA),
  REF: getExcelError(ExcelErrorCode.REF),
  VALUE: getExcelError(ExcelErrorCode.VALUE)
});

class FormulaFunctionCollisionError extends Error {
  functionName: string;

  constructor(name: string) {
    super('Formula function already exists: ' + name);
    this.name = 'FormulaFunctionCollisionError';
    this.functionName = name;
  }
}

class FormulaFunctionNotFoundError extends Error {
  functionName: string;

  constructor(name: string) {
    super('Unknown formula function: ' + name);
    this.name = 'FormulaFunctionNotFoundError';
    this.functionName = name;
  }
}

class AsyncFormulaFunctionError extends Error {
  functionName: string;

  constructor(name: string) {
    super('Formula function must return a synchronous value: ' + name);
    this.name = 'AsyncFormulaFunctionError';
    this.functionName = name;
  }
}

class FormulaCycleError extends Error {
  cyclePath: string[];

  constructor(cyclePath: string[]) {
    super('Circular formula dependency detected: ' + cyclePath.join(' -> '));
    this.name = 'FormulaCycleError';
    this.cyclePath = cyclePath;
  }
}

class FormulaEvaluationError extends Error {
  cause: Error;
  cellKey: string;

  constructor(cellKey: string, cause: Error) {
    super('Failed to evaluate formula for ' + cellKey + ': ' + cause.message);
    this.name = 'FormulaEvaluationError';
    this.cause = cause;
    this.cellKey = cellKey;
  }
}

function getExcelError(code: string): ExcelErrorValue {
  const excelError = excelErrorValues.get(code as ExcelErrorCodeValue);

  if (excelError === undefined) {
    throw new Error('Unsupported Excel error code: ' + code);
  }

  return excelError;
}

function isExcelError(value: unknown): value is ExcelErrorValue {
  return value instanceof ExcelErrorValue;
}

function normalizeExcelErrorCode(code: unknown): ExcelErrorCodeValue | null {
  if (typeof code !== 'string') {
    return null;
  }

  const normalizedCode = code.trim() as ExcelErrorCodeValue;

  return excelErrorValues.has(normalizedCode) ? normalizedCode : null;
}

function toExcelError(value: unknown, fallbackCode?: ExcelErrorCodeValue): ExcelErrorValue | null {
  if (value instanceof ExcelErrorValue) {
    return value;
  }

  if (value instanceof FormulaFunctionNotFoundError) {
    return ExcelError.NAME;
  }

  if (value instanceof AsyncFormulaFunctionError) {
    return ExcelError.VALUE;
  }

  if (value instanceof Error) {
    const normalizedErrorCode = normalizeExcelErrorCode((value as Error & { code?: unknown }).code);

    if (normalizedErrorCode !== null) {
      return getExcelError(normalizedErrorCode);
    }

    const normalizedMessageCode = normalizeExcelErrorCode(value.message);

    if (normalizedMessageCode !== null) {
      return getExcelError(normalizedMessageCode);
    }
  }

  if (fallbackCode !== undefined) {
    return getExcelError(fallbackCode);
  }

  return null;
}

function findExcelError(value: unknown): ExcelErrorValue | null {
  const directError = toExcelError(value);

  if (directError !== null) {
    return directError;
  }

  if (!Array.isArray(value)) {
    return null;
  }

  for (const item of value) {
    const nestedError = findExcelError(item);

    if (nestedError !== null) {
      return nestedError;
    }
  }

  return null;
}

export {
  AsyncFormulaFunctionError,
  ExcelError,
  type ExcelErrorCodeValue,
  ExcelErrorCode,
  ExcelErrorValue,
  FormulaCycleError,
  FormulaEvaluationError,
  FormulaFunctionCollisionError,
  FormulaFunctionNotFoundError,
  findExcelError,
  getExcelError,
  isExcelError,
  normalizeExcelErrorCode,
  toExcelError
};
