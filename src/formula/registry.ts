'use strict';

import * as FormulaPackage from '@formulajs/formulajs';
import {
  AsyncFormulaFunctionError,
  FormulaFunctionCollisionError,
  FormulaFunctionNotFoundError
} from './errors';
import type {
  FormulaFunctionImplementation,
  FormulaFunctionMap,
  FormulaFunctionNamespace,
  FormulaRegistryOptions,
  FormulaRegistryOverrideOptions,
  WorkbookCellValue
} from './types';

const FormulaSource = FormulaPackage as Record<string, unknown>;

function createDefaultFormulaFunctions(): Readonly<FormulaFunctionMap> {
  const formulaFunctions: FormulaFunctionMap = Object.create(null) as FormulaFunctionMap;

  for (const [name, value] of Object.entries(FormulaSource)) {
    if (typeof value === 'function') {
      formulaFunctions[name.toUpperCase()] = value as FormulaFunctionImplementation;
    }
  }

  return Object.freeze(formulaFunctions);
}

const Formula = createDefaultFormulaFunctions();

function normalizeFunctionName(name: unknown): string {
  if (typeof name !== 'string' || name.trim() === '') {
    throw new TypeError('Formula function names must be non-empty strings.');
  }

  return name.trim().toUpperCase();
}

function collectFunctionEntries(functionMap: Record<string, FormulaFunctionImplementation> | undefined): Array<[string, FormulaFunctionImplementation]> {
  const entries: Array<[string, FormulaFunctionImplementation]> = [];

  for (const [name, value] of Object.entries(functionMap || {})) {
    entries.push([normalizeFunctionName(name), value]);
  }

  return entries;
}

function isPromiseLike<TResult = unknown>(value: unknown): value is PromiseLike<TResult> {
  return value !== null &&
    typeof value === 'object' &&
    typeof (value as { then?: unknown }).then === 'function';
}

function assertSynchronousFunctionResult<TResult>(functionName: string, value: TResult | PromiseLike<TResult>): TResult {
  if (isPromiseLike(value)) {
    throw new AsyncFormulaFunctionError(functionName);
  }

  return value;
}

class FormulaFunctionRegistry {
  baseFunctions: Map<string, FormulaFunctionImplementation>;
  localFunctions: Map<string, FormulaFunctionImplementation>;
  namespaceCache: FormulaFunctionNamespace | null;

  constructor(options?: FormulaRegistryOptions) {
    const resolvedOptions = options || {};
    const baseFunctions = resolvedOptions.baseFunctions || Formula;

    this.baseFunctions = new Map(collectFunctionEntries(baseFunctions));
    this.localFunctions = new Map();
    this.namespaceCache = null;

    if (resolvedOptions.functions !== undefined) {
      this.registerMany(resolvedOptions.functions);
    }
  }

  clone(): FormulaFunctionRegistry {
    const clone = new FormulaFunctionRegistry({
      baseFunctions: Object.fromEntries(this.baseFunctions)
    });

    for (const [name, implementation] of this.localFunctions) {
      clone.localFunctions.set(name, implementation);
    }

    return clone;
  }

  register(name: string, implementation: FormulaFunctionImplementation, options?: FormulaRegistryOverrideOptions): this {
    const normalizedName = normalizeFunctionName(name);
    const resolvedOptions = options || {};
    const override = resolvedOptions.override === true;

    if (typeof implementation !== 'function') {
      throw new TypeError('Formula functions must be callable: ' + normalizedName);
    }

    if (!override && this.has(normalizedName)) {
      throw new FormulaFunctionCollisionError(normalizedName);
    }

    this.localFunctions.set(normalizedName, implementation);
    this.namespaceCache = null;
    return this;
  }

  registerMany(functionMap: Record<string, FormulaFunctionImplementation>, options?: FormulaRegistryOverrideOptions): this {
    for (const [name, implementation] of Object.entries(functionMap)) {
      this.register(name, implementation, options);
    }

    return this;
  }

  unregister(name: string): this {
    const normalizedName = normalizeFunctionName(name);

    this.localFunctions.delete(normalizedName);
    this.namespaceCache = null;
    return this;
  }

  has(name: string): boolean {
    return this.get(name) !== undefined;
  }

  get(name: string): FormulaFunctionImplementation | undefined {
    const normalizedName = normalizeFunctionName(name);

    if (this.localFunctions.has(normalizedName)) {
      return this.localFunctions.get(normalizedName);
    }

    return this.baseFunctions.get(normalizedName);
  }

  list(): string[] {
    return Array.from(new Set([
      ...this.baseFunctions.keys(),
      ...this.localFunctions.keys()
    ])).sort();
  }

  getNamespace(): FormulaFunctionNamespace {
    if (this.namespaceCache === null) {
      this.namespaceCache = Object.freeze(Object.assign(
        Object.create(null),
        Object.fromEntries(this.baseFunctions),
        Object.fromEntries(this.localFunctions)
      ));
    }

    const namespaceCache = this.namespaceCache;

    if (namespaceCache === null) {
      throw new Error('Formula function namespace cache was not initialized.');
    }

    return namespaceCache;
  }

  invoke(name: string, argumentsValues: WorkbookCellValue[]): unknown {
    const normalizedName = normalizeFunctionName(name);
    const implementation = this.get(normalizedName);

    if (implementation === undefined) {
      throw new FormulaFunctionNotFoundError(normalizedName);
    }

    return assertSynchronousFunctionResult(
      normalizedName,
      implementation.apply(this.getNamespace(), argumentsValues)
    );
  }
}

export {
  Formula,
  FormulaFunctionRegistry,
  assertSynchronousFunctionResult,
  normalizeFunctionName
};
