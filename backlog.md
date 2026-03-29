# Backlog

Status:

- Completed and archived as a modernization record.
- Some file references below point to paths that existed earlier in the refactor and were later renamed or removed.

## Assumptions

- Breaking changes are allowed.
- The core product idea stays the same: parse XLSX files, translate formulas, and evaluate workbooks in JavaScript.
- We should keep both Node and browser support, but we do not need to preserve the current global-heavy API or old tooling.
- The end state should be modern, typed, testable, and maintainable.

## Target State

- Runtime source is written in TypeScript, not legacy CommonJS object factories.
- Public API is explicit and documented, with modern module exports.
- Formula handling is based on a real AST/compiler or evaluator, not JavaScript strings plus `eval`.
- Generated artifacts live in build output, not mixed into hand-authored source.
- Browser support comes from a modern build pipeline, not legacy demo/vendor scaffolding.
- Tests cover parser correctness, workbook evaluation, serialization, error handling, and regressions.

## P0: Architecture Reset

- [x] Replace the current global mutation/module pattern in [src/index.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/index.ts), [src/excellent.loader.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.loader.ts), [src/excellent.util.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.util.ts), [src/excellent.workbook.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.workbook.ts), [src/excellent.xlsx.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.xlsx.ts), and [src/excellent.xlsx-simple.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.xlsx-simple.ts) with real modules and classes.
- [x] Define a clean domain model:
  - `Workbook`
  - `Worksheet`
  - `Cell`
  - `WorkbookLoader`
  - `XlsxReader`
  - `FormulaCompiler` or `FormulaEvaluator`
- [x] Collapse duplicated XLSX parsing logic in [src/excellent.xlsx.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.xlsx.ts) and [src/excellent.xlsx-simple.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.xlsx-simple.ts) into one implementation with clear feature flags or internal modes.
- [x] Remove the implicit mutation-heavy API around `currentSheet` and replace it with explicit methods and internal state boundaries.
- [x] Stop treating generated parser output as source-of-truth. The grammar in [grammar/formula_parser.pegjs](/Users/mrcsparker/Documents/GitHub/excellent/grammar/formula_parser.pegjs) should be authoritative, and generated files should move out of `src/`.

Done when:

- No runtime code depends on `root.Excellent`, `exports.Excellent`, or browser globals for core behavior.
- The codebase has a clear separation between domain logic, parsing, IO, and packaging.

## P0: Remove `eval` and String-Based Execution

- [x] Replace formula execution in [src/excellent.workbook.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.workbook.ts) with an AST-driven evaluator or a controlled compiler that produces callable functions.
- [x] Change the parser contract so formulas are not represented primarily as strings like `Formula.SUM(this.A1,this.A2)`.
- [x] Add explicit handling for:
  - dependency tracking
  - lazy evaluation
  - memoization
  - cache invalidation after cell updates
  - cycle detection
  - formula/runtime errors
- [x] Introduce internal error types instead of leaking raw parser/runtime failures.

Done when:

- Workbook evaluation no longer relies on `eval`.
- Formula evaluation is deterministic, testable, and isolated from object property side effects.

## P0: Convert Runtime Source to TypeScript

- [x] Move runtime code from JavaScript to TypeScript.
- [x] Replace the former hand-maintained root declarations with types emitted from source into [dist/index.d.ts](/Users/mrcsparker/Documents/GitHub/excellent/dist/index.d.ts).
- [x] Turn on a stricter TypeScript baseline:
  - `noUncheckedIndexedAccess`
  - `exactOptionalPropertyTypes`
  - `useUnknownInCatchVariables`
  - `noImplicitOverride`
  - `noPropertyAccessFromIndexSignature`
- [x] Introduce narrow domain types for workbook cells, worksheets, parsed XML nodes, parser AST nodes, and serialized workbook payloads.

Done when:

- `index.d.ts` is generated from source, not hand-authored.
- TypeScript describes the real runtime model instead of broad `unknown`/record-shaped approximations.

## P1: Public API Redesign

- [x] Decide on the modern public API and document it before implementation.
- [x] Replace the current namespace-style surface with explicit exports.
- [x] Design for both JavaScript and TypeScript consumers.

Possible target shape:

```ts
import { XlsxReader, WorkbookLoader } from 'excellent';

const workbook = await new XlsxReader().load(bytes);
const json = new WorkbookLoader().serialize(workbook);
```

- [x] Decide whether raw translated formulas should stay public, move behind a debug API, or be removed entirely.
- [x] Decide how cell access should work:
  - property access
  - `getCell('A1')`
  - both, with one clearly primary

Done when:

- The API is small, explicit, and documented.
- Deprecated alias exports like `ExcellentLoader` are removed unless they still serve a clean purpose.

## P1: Packaging and Build Cleanup

- [x] Add an `exports` map in [package.json](/Users/mrcsparker/Documents/GitHub/excellent/package.json).
- [x] Decide on package format:
  - ESM-first with a CJS compatibility build
  - or ESM-only if we want a hard clean break
- [x] Move build outputs to `dist/`.
- [x] Stop checking generated runtime artifacts into source locations.
- [x] Rework [scripts/build-parser.js](/Users/mrcsparker/Documents/GitHub/excellent/scripts/build-parser.js) and [scripts/build-browser.js](/Users/mrcsparker/Documents/GitHub/excellent/scripts/build-browser.js) to build from typed source.
- [x] Delete legacy package/tooling artifacts that no longer belong:
  - [bower.json](/Users/mrcsparker/Documents/GitHub/excellent/bower.json)
  - [gulpfile.js](/Users/mrcsparker/Documents/GitHub/excellent/gulpfile.js)
  - [Makefile](/Users/mrcsparker/Documents/GitHub/excellent/Makefile)
  - [src/excellent.compat.js](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.compat.js)
- [x] Decide whether checked-in bundles like [excellent.js](/Users/mrcsparker/Documents/GitHub/excellent/excellent.js) and [excellent.min.js](/Users/mrcsparker/Documents/GitHub/excellent/excellent.min.js) should be published artifacts only, not repository-maintained source files.

Done when:

- The package has one clean source tree and one clean output tree.
- There is no dead Bower/Gulp-era scaffolding left in the repo.

## P1: Parser and Formula Coverage

- [x] Redesign the grammar output from [grammar/formula_parser.pegjs](/Users/mrcsparker/Documents/GitHub/excellent/grammar/formula_parser.pegjs) to produce an AST or IR instead of JS code strings.
- [x] Implement first-class Excel error semantics instead of falling through to generic JavaScript behavior:
  - `#DIV/0!`
  - `#VALUE!`
  - `#REF!`
  - `#NAME?`
  - `#N/A`
  - propagation rules across formulas and ranges
- [x] Add tests for currently under-specified behavior:
  - cross-sheet references
  - shared formulas
  - absolute vs relative references
  - string escaping
  - empty cells and nulls
  - numeric coercion
  - workbook mutation after load
  - formula errors
  - circular dependencies
- [x] Add golden fixtures for real XLSX structures beyond the two current test files in [test/data](/Users/mrcsparker/Documents/GitHub/excellent/test/data).
- [x] Add parser regression tests so grammar changes are safer.

Done when:

- Parser behavior is specified by tests, not tribal knowledge.
- Formula semantics can be changed confidently without fear of silent breakage.
- Excel-style error behavior is deliberate and test-backed.

## P1: Observability and Debugging

- [x] Add a formula trace/debug API that can explain workbook evaluation.
- [x] Expose workbook introspection primitives:
  - precedents for a cell
  - dependents for a cell
  - dependency graph traversal
  - evaluation trace for a computed value
- [x] Decide whether the debug surface should be public API or an advanced/internal module with a stable contract.
  - Public API: `Workbook.traceCell(sheetName, cellName)`

Done when:

- A user can answer “why does `A1` equal this?” without reading internal code or stepping through runtime internals.

## P1: Extensibility

- [x] Add a custom function registry so consumers can inject domain-specific spreadsheet functions without mutating global formula state.
- [x] Define function registration semantics:
  - global vs per-workbook scope
  - sync vs async support
  - name collision rules
  - error mapping rules
- [x] Document how custom functions interact with formula compilation, caching, and browser builds.

Done when:

- Consumers can extend the engine intentionally instead of patching internals or monkey-patching `Formula`.

## P2: Performance and Correctness Work

- [x] Add evaluation benchmarks for large sheets and shared formula cases.
- [x] Profile XML parsing and workbook construction to identify hotspots before optimizing.
- [x] Add caching only where measurement justifies it.
- [x] Consider a streaming or incremental load path later if the architecture supports it cleanly.

Done when:

- Performance work is driven by measurements, not guesses.

## P2: Test Strategy Upgrade

- [x] Keep the current unit tests, but reorganize them by layer:
  - parser
  - workbook model
  - evaluator
  - XLSX integration
  - public API
- [x] Add coverage reporting.
- [x] Add browser smoke tests for the bundled/browser build.
- [x] Add fixture-based integration tests that load actual XLSX workbooks and assert both values and metadata.
- [x] Add mutation tests around updating inputs after load so dependency invalidation is trustworthy.
- [x] Add property-based tests around formula parsing, cell references, ranges, and serialization invariants.
- [x] Add differential tests against a reference spreadsheet engine where practical to catch semantic drift in formula behavior.

Done when:

- A refactor of the core evaluator or parser can be validated quickly and with confidence.
- Correctness is checked against both invariants and an external reference, not only handwritten examples.

## P2: Demo and Documentation

- [x] Replace the current Angular/jQuery demo in [demo/index.html](/Users/mrcsparker/Documents/GitHub/excellent/demo/index.html) and [demo/scripts/demo.js](/Users/mrcsparker/Documents/GitHub/excellent/demo/scripts/demo.js) with a small modern demo or remove it until a modern replacement exists.
- [x] Rewrite [README.md](/Users/mrcsparker/Documents/GitHub/excellent/README.md) around the new API, supported feature set, architecture, and limitations.
- [x] Document what is intentionally unsupported.
- [x] Document the release/build workflow so the repo is maintainable by someone who did not originally write it.

Done when:

- The README reflects the actual product.
- The demo represents the current architecture instead of anchoring the repo to legacy browser code.

## Suggested Execution Order

1. Freeze the target API and package strategy.
2. Move runtime source to TypeScript modules.
3. Replace the formula string/`eval` pipeline with AST-driven evaluation.
4. Rebuild workbook, worksheet, and loader around explicit classes.
5. Clean up packaging and delete dead legacy artifacts.
6. Expand tests and fixtures aggressively.
7. Rewrite docs and demo last, once the architecture is stable.

## Non-Negotiable Quality Bar

- No `eval` in runtime code.
- No global mutation as the core composition mechanism.
- No generated files treated as hand-maintained source.
- No dead packaging/tooling history left in the critical path.
- No hand-wavy TypeScript types that drift away from runtime reality.

## P3: Remaining Gaps

- [x] Make the published package actually ESM-first at runtime, with native `import` support and a tested CJS compatibility path.
- [x] Remove the runtime TypeScript carve-outs in [tsconfig.runtime.json](/Users/mrcsparker/Documents/GitHub/excellent/tsconfig.runtime.json) so the shipped runtime passes the same strict baseline as the repo.
- [x] Reorganize the core runtime into [src/formula/index.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/formula/index.ts), [src/formula/errors.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/formula/errors.ts), [src/formula/compiler.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/formula/compiler.ts), [src/formula/evaluator.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/formula/evaluator.ts), [src/workbook/index.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/workbook/index.ts), [src/workbook/cell.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/workbook/cell.ts), [src/workbook/worksheet.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/workbook/worksheet.ts), and [src/workbook/workbook.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/workbook/workbook.ts) with no flat compatibility layer.
- [x] Raise direct coverage for [src/formula/evaluator.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/formula/evaluator.ts) and close the largest remaining evaluator branch gaps.
- [x] Remove the authored-source dependency on build output by replacing the `src` to `dist/generated` parser coupling in [src/excellent.parser.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.parser.ts).
- [x] Continue large-workbook scalability work with another profiling pass and a concrete decision on whether true streaming XLSX support is worth the added complexity.

Done when:

- The package format, source layout, and strictness level match the repo’s stated quality bar instead of a transitional compromise.
- The hardest parts of the evaluator are both better covered and easier to change confidently.
