'use strict';

import type { FormulaAstNode } from './formula';

declare function require(moduleName: string): unknown;

type FormulaParserModule = {
  parse(input: string): FormulaAstNode;
};

const FormulaParserRuntime = require('../generated/excellent.parser.js') as FormulaParserModule;
const FormulaParser = FormulaParserRuntime;

export {
  FormulaParser,
  type FormulaParserModule
};
