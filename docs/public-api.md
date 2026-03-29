# Public API

This document describes the current public API for Excellent.

## Design Goals

- Prefer explicit named exports over namespace objects.
- Treat `Workbook` as the core runtime model, regardless of whether it came from XLSX or JSON.
- Keep I/O boundaries async where the underlying dependencies require it.
- Make the JavaScript and TypeScript story the same, not two different APIs.
- Keep convenience accessors possible, but do not make them the main contract.

## Primary Exports

These are the exports the package should document and optimize for:

- `XlsxReader`
- `WorkbookLoader`
- `Workbook`
- `Worksheet`
- `Cell`
- `FormulaFunctionRegistry`
- `FormulaEvaluator`
- `ExcelError`
- `isExcelError`
- `XLSX_READER_MODE`

## Primary Workflows

### Read an XLSX workbook

```ts
import { XlsxReader } from 'excellent';

const workbook = await new XlsxReader().load(bytes);
const total = workbook.getCellValue('Sheet1', 'A1');
```

### Load an XLSX workbook incrementally by sheet

```ts
import { XlsxReader } from 'excellent';

const reader = new XlsxReader();

const workbook = await reader.loadIncremental(bytes, async ({ sheetName, worksheet, workbook }) => {
  console.log(sheetName);
  console.log(worksheet.getCellValue('A1'));
  console.log(workbook.getSheetNames());
});
```

### Serialize and reload workbook state

```ts
import { WorkbookLoader } from 'excellent';

const loader = new WorkbookLoader();
const json = loader.serialize(workbook);
const restoredWorkbook = loader.deserialize(json);
```

### Build a workbook in memory

```ts
import { Workbook } from 'excellent';

const workbook = new Workbook();
const sheet = workbook.createSheet('Sheet1');

sheet.setCellValue('A1', 4);
sheet.setCellFormula('A2', 'this.A1+1');
const total = sheet.getCellValue('A2');
```

### Register custom functions

```ts
import { FormulaFunctionRegistry, Workbook } from 'excellent';

const functionRegistry = new FormulaFunctionRegistry().register('DOUBLE', (value) => Number(value) * 2);
const workbook = new Workbook({ functionRegistry });
```

## Chosen API Rules

### 1. Named exports are the primary surface

The documented API is:

```ts
import {
  Workbook,
  WorkbookLoader,
  XlsxReader
} from 'excellent';
```

The package entrypoint uses named exports directly. The browser IIFE still publishes those exports under the global `Excellent` object.

### 2. `Workbook` is the main runtime object

Readers and loaders should return a `Workbook`.

That means the package should revolve around these paths:

- `await new XlsxReader().load(bytes) -> Workbook`
- `await new XlsxReader().loadIncremental(bytes, onSheetLoaded) -> Workbook`
- `new WorkbookLoader().deserialize(json) -> Workbook`
- `new WorkbookLoader().serialize(workbook) -> LoaderWorkbook`

### 3. Explicit methods are the primary cell API

Primary read/write APIs:

- `workbook.getCell(sheetName, cellName)`
- `workbook.getCellValue(sheetName, cellName)`
- `workbook.setCellValue(sheetName, cellName, value)`
- `workbook.setCellFormula(sheetName, cellName, expression)`
- `worksheet.getCell(cellName)`
- `worksheet.getCellValue(cellName)`
- `worksheet.setCellValue(cellName, value)`
- `worksheet.setCellFormula(cellName, expression)`

Choice:

- Keep both explicit methods and property sugar.
- `worksheet.getCellValue('A1')` is the primary read API.
- `worksheet.getCell('A1')` is for the cell model and metadata, not the main value-read path.
- Dynamic property access such as `sheet.A1` remains convenience sugar for reads and literal writes, but it is not the primary documented contract.

### 4. Debug data is debug data

Raw and translated formula representations should not be treated as the main end-user API.

The intended debug/introspection surface is:

- `Workbook.traceCell(sheetName, cellName)`
- `Workbook.getFormulaSource(sheetName, cellName)`
- `Worksheet.getFormulaSource(cellName)`
- `Cell.getFormulaSource()`
- `Cell.getCompiledFormula()`
- `Cell.getRawValue()` for literal value cells
- dependency graph APIs such as `getPrecedents()` and `getDependents()`

Underscore accessors like `_A1` have been removed. Explicit debug methods are the supported contract.

### 5. JavaScript and TypeScript use the same conceptual API

JavaScript ESM example:

```js
import { Workbook, XlsxReader } from 'excellent';

const workbook = await new XlsxReader().load(bytes);
const sheet = workbook.requireSheet('Sheet1');
sheet.setCellValue('A1', 2);
```

Current JavaScript contract:

- In modern Node.js JavaScript, prefer native `import { ... } from 'excellent'`.
- The package also ships a stable CJS compatibility path for `require('excellent')`.
- Both module syntaxes expose the same named-export runtime concepts.
- `loadIncremental()` is an incremental worksheet API, not a promise of true zip-streaming XLSX reads.

CommonJS compatibility example:

```js
const { Workbook, XlsxReader } = require('excellent');

(async function() {
  const workbook = await new XlsxReader().load(bytes);
  const sheet = workbook.requireSheet('Sheet1');

  sheet.setCellValue('A1', 2);
}());
```

TypeScript example:

```ts
import { Workbook, XlsxReader } from 'excellent';

const workbook = await new XlsxReader().load(bytes);
const sheet = workbook.requireSheet('Sheet1');
sheet.setCellValue('A1', 2);
```

Current TypeScript contract:

- Use named imports from `'excellent'`.
- Rely on the emitted package declarations for the public type surface.
- Do not learn a separate namespace or compatibility layer.

The package should not require JavaScript and TypeScript consumers to learn different runtime concepts, only different module syntax appropriate to the current package format.

## Removed Legacy Surface

These compatibility aliases are no longer exported:

- `Xlsx`
- `Loader`
- `XlsxSimple`
- `WorkbookLoader.load(...)`
- `WorkbookLoader.unload(...)`

Use the explicit surface instead:

- `XlsxReader`
- `WorkbookLoader.deserialize(...)`
- `WorkbookLoader.serialize(...)`
