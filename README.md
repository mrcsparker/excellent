# Excellent

Excellent is a typed workbook engine for JavaScript and TypeScript. It reads `.xlsx` files, evaluates spreadsheet formulas, and exposes the workbook as an explicit `Workbook` / `Worksheet` / `Cell` model that you can inspect, mutate, serialize, and trace.

This repo is no longer the original prototype. The runtime is now TypeScript-based, `eval` is gone, formulas are parsed into an AST, and the public API is centered on named exports and explicit workbook methods.

## What Excellent Is For

Use Excellent when you want to:

- load an Excel workbook into application code
- evaluate formulas in Node.js or in the browser
- inspect dependencies and trace computed values
- mutate workbook inputs and recalculate downstream cells
- serialize workbook state into a JSON shape that Excellent can reload later

Use it as a workbook runtime, not as a spreadsheet UI toolkit or a full Excel clone.

## Highlights

- Explicit runtime model: `Workbook`, `Worksheet`, `Cell`, `XlsxReader`, `WorkbookLoader`
- AST-based formula evaluation with no runtime `eval`
- Excel-style error values such as `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, and `#N/A`
- Shared formulas, ranges, cross-sheet references, quoted sheet names, and absolute references
- Dependency tracking, lazy evaluation, memoization, invalidation, and cycle detection
- Debug/introspection APIs: precedents, dependents, graph traversal, and `traceCell()`
- Workbook-scoped custom functions without mutating global formula state
- Browser bundle plus a modern demo surface in [demo/index.html](demo/index.html)
- Type declarations emitted from source and shipped with the package

## Installation

```bash
npm install excellent
```

Requirements:

- Node.js `>=20.13.0`

The package ships generated declarations from source. You do not need a separate `@types` package.

## Quick Start

### Node.js JavaScript (ESM)

```js
import { XlsxReader } from 'excellent';
import { readFile } from 'node:fs/promises';

const bytes = await readFile('./model.xlsx');
const workbook = await new XlsxReader().load(bytes);

console.log(workbook.getCellValue('Sheet1', 'A1'));
```

### TypeScript

```ts
import { XlsxReader } from 'excellent';
import { readFile } from 'node:fs/promises';

const bytes = await readFile('./model.xlsx');
const workbook = await new XlsxReader().load(bytes);

console.log(workbook.getCellValue('Sheet1', 'A1'));
```

### Browser

Build the browser bundle, then load `dist/excellent.js`:

```html
<script src="./dist/excellent.js"></script>
<script>
  (async function () {
    const response = await fetch('./test/data/simpleFormula.xlsx');
    const bytes = await response.arrayBuffer();
    const workbook = await new Excellent.XlsxReader().load(bytes);

    console.log(workbook.getCellValue('Sheet1', 'A1'));
  }());
</script>
```

The browser build exposes the same runtime concepts under the global `Excellent` object.

### CommonJS compatibility

```js
const { XlsxReader } = require('excellent');
const fs = require('node:fs/promises');

(async function () {
  const bytes = await fs.readFile('./model.xlsx');
  const workbook = await new XlsxReader().load(bytes);

  console.log(workbook.getCellValue('Sheet1', 'A1'));
}());
```

## Core API

### Read an XLSX workbook

```ts
import { XlsxReader } from 'excellent';

const workbook = await new XlsxReader().load(bytes);
const total = workbook.getCellValue('Inputs', 'B4');
```

### Load incrementally by sheet

```ts
import { XlsxReader } from 'excellent';

const reader = new XlsxReader();

const workbook = await reader.loadIncremental(bytes, async ({ sheetName, worksheet, workbook }) => {
  console.log(sheetName);
  console.log(worksheet.getCellValue('A1'));
  console.log(workbook.getSheetNames());
});
```

`loadIncremental()` is incremental by worksheet, not true zip-level streaming. The XLSX container is still opened first and sheets are materialized one at a time. After the current profiling pass, true streaming is intentionally not planned until measurements show container materialization or retained memory is the real bottleneck.

### Build a workbook in memory

```ts
import { Workbook } from 'excellent';

const workbook = new Workbook();
const sheet = workbook.createSheet('Sheet1');

sheet.setCellValue('A1', 4);
sheet.setCellFormula('A2', 'this.A1+1');

console.log(sheet.getCellValue('A2')); // 5
```

### Serialize and reload workbook state

```ts
import { WorkbookLoader } from 'excellent';

const loader = new WorkbookLoader();
const json = loader.serialize(workbook);
const restoredWorkbook = loader.deserialize(json);
```

### Register custom functions

```ts
import { FormulaFunctionRegistry, Workbook } from 'excellent';

const functionRegistry = new FormulaFunctionRegistry()
  .register('DOUBLE', (value) => Number(value) * 2);

const workbook = new Workbook({ functionRegistry });
const sheet = workbook.createSheet('Sheet1');

sheet.setCellValue('A1', 4);
sheet.setCellFormula('A2', 'Formula.DOUBLE(this.A1)');

console.log(sheet.getCellValue('A2')); // 8
```

Custom function rules today:

- registration is workbook-scoped by default
- async custom functions are not supported
- name collisions throw unless you pass `{ override: true }`
- changing the function registry invalidates formula caches

### Trace a computed cell

```ts
import { XlsxReader } from 'excellent';

const workbook = await new XlsxReader().load(bytes);
const trace = workbook.traceCell('Sheet1', 'A1');

console.log(trace.value);
console.log(trace.precedents);
console.log(trace.evaluation);
```

Related debug APIs:

- `workbook.getPrecedents(sheetName, cellName)`
- `workbook.getDependents(sheetName, cellName)`
- `workbook.traversePrecedents(sheetName, cellName)`
- `workbook.traverseDependents(sheetName, cellName)`
- `workbook.getFormulaSource(sheetName, cellName)`
- `worksheet.getFormulaSource(cellName)`
- `cell.getCompiledFormula()`

## API Conventions

- Named exports are the primary package surface.
- `Workbook` is the central runtime object regardless of whether it came from XLSX or JSON.
- Explicit methods such as `getCellValue()` and `setCellValue()` are the primary contract.
- Property sugar like `sheet.A1` still exists for convenience, but it is not the main documented API.
- Creating formula cells is explicit: use `setCellFormula(...)`, not property assignment.

## Supported Feature Set

Excellent currently covers these core behaviors:

- XLSX workbook loading in Node.js and the browser
- shared strings and shared formulas
- ranges such as `SUM(A1:B9)`
- cross-sheet references, including quoted sheet names
- relative and absolute references
- workbook mutation and recalculation after load
- dependency graph tracking and cycle detection
- Excel-style error propagation for supported formulas
- `IF`, `IFERROR`, `IFNA`, `INDEX`, `MATCH`, and a broader formula subset backed by `@formulajs/formulajs`
- JSON serialization and deserialization through `WorkbookLoader`
- browser bundle smoke coverage and fixture-backed integration coverage

The test suite also includes:

- parser regression tests
- property-based tests
- differential tests against HyperFormula where the semantics overlap

## Architecture

At a high level, the runtime is organized like this:

1. The XLSX reader opens the workbook zip, reads workbook metadata, shared strings, and worksheet XML.
2. The Peggy grammar parses formulas into an AST instead of generating JavaScript source strings.
3. The formula evaluator compiles, evaluates, serializes, and traces those AST nodes.
4. The workbook model stores explicit `Cell` objects inside `Worksheet` and `Workbook` classes.
5. The workbook runtime owns dependency tracking, memoization, invalidation, cycle detection, and trace APIs.
6. `WorkbookLoader` provides the JSON serialization boundary for saving and reloading workbook state.

Relevant docs:

- [docs/public-api.md](docs/public-api.md)
- [docs/package-format.md](docs/package-format.md)
- [docs/performance-profile.md](docs/performance-profile.md)
- [docs/release-workflow.md](docs/release-workflow.md)
- [docs/unsupported.md](docs/unsupported.md)

## Current Limitations

Excellent is intentionally narrower than Excel itself.

- It is a workbook runtime, not a full spreadsheet application.
- It does not write `.xlsx` files back out today.
- It does not render formatting, styling, charts, images, pivot tables, or macros/VBA.
- `loadIncremental()` is not true streaming; it is sheet-by-sheet materialization after the workbook container is opened, and true streaming is intentionally out of scope for now.
- Custom functions are sync-only.
- The debug formula source is engine-oriented source used for inspection, not a promise of original Excel formula text round-tripping.
- The published package is ESM-first and also ships a CJS compatibility entrypoint for `require('excellent')`.

For the explicit out-of-scope contract, see [docs/unsupported.md](docs/unsupported.md).

## Development

Common repo commands:

```bash
npm test
npm run test:browser
npm run build
npm run coverage
```

You can also profile or benchmark the runtime:

```bash
npm run bench:evaluation
npm run profile:xlsx-load
```

Maintainer-facing build and release flow is documented in [docs/release-workflow.md](docs/release-workflow.md).

## Demo

The repo includes a browser-local demo surface in [demo/index.html](demo/index.html). Build the browser bundle, serve the repo over HTTP, and open the demo page to load fixture workbooks or your own `.xlsx` files.
