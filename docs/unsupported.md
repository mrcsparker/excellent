# Intentionally Unsupported

This document defines what Excellent does not try to be.

That distinction matters because the project is a workbook runtime, not a full Excel replacement. These are deliberate scope boundaries, not bugs filed under "maybe later" by default.

## Product Boundary

Excellent is designed to:

- read `.xlsx` workbooks
- evaluate formulas
- expose workbook data and dependencies to application code
- support recalculation, serialization, and debugging

Excellent is not designed to replicate the whole Excel product surface.

## Intentionally Out Of Scope

### 1. Writing `.xlsx` files

Excellent does not currently emit a new `.xlsx` workbook file.

What it does support instead:

- loading `.xlsx`
- mutating the in-memory workbook model
- serializing that workbook model through `WorkbookLoader`

If you need round-tripping back into a real Excel file, that is outside the current scope.

### 2. Formatting and presentation features

Excellent does not treat spreadsheet presentation features as part of the runtime contract.

That includes things like:

- styles and themes
- fonts, fills, borders, and number-format presentation
- layout/presentation metadata as a first-class API surface
- charts and images
- pivot tables as rendered spreadsheet features

The engine is concerned with workbook data and formula behavior, not reproducing Excel's visual layer.

### 3. Macros, VBA, and Excel automation features

Excellent does not execute or model:

- VBA
- macros
- Excel application automation behavior

The runtime is intentionally pure JavaScript/TypeScript workbook evaluation.

### 4. Full Excel feature parity

Excellent does not claim complete compatibility with every Excel feature, workbook shape, or formula edge case.

What the repo does claim:

- a tested and expanding subset of formula behavior
- explicit coverage for the scenarios in the test suite
- differential checks against a reference engine where practical

What it does not claim:

- "drop-in Excel clone"
- total function parity
- identical behavior for every obscure workbook construct

If behavior is not covered by the documented API and tests, it should not be assumed.

### 5. Spreadsheet UI responsibilities

Excellent is not a spreadsheet grid component or end-user spreadsheet application.

The demo in [demo/index.html](../demo/index.html) is just a repo demo surface for the engine. It is not the main product.

Out of scope here:

- collaborative spreadsheet editing
- full worksheet UX parity with Excel or Google Sheets
- a reusable spreadsheet UI framework maintained by this package

### 6. True streaming XLSX reads

`XlsxReader.loadIncremental()` is incremental by worksheet, but it is not a true streaming zip reader.

The current contract is:

- the workbook container is opened first
- worksheets are materialized one at a time
- callbacks can observe the workbook as each worksheet finishes loading

The current contract is not:

- constant-memory streaming over the zip container
- byte-level progressive parsing of arbitrary workbook parts

Current decision:

- after the March 29, 2026 profiling pass, true streaming remains intentionally out of scope
- the measured bottleneck is still XML parse and formula compilation, not zip container reads
- revisit only if future profiling shows container materialization or memory retention is the dominant real-world problem

For the measurement behind that decision, see [docs/performance-profile.md](./performance-profile.md).

### 7. Async custom functions

Custom functions are intentionally synchronous today.

If a custom function returns a `Promise`, the engine does not treat that as a supported async recalculation model. The current runtime stays synchronous and maps that case to an Excel-style error result instead.

### 8. Exact original formula text round-tripping

Excellent exposes formula source for debugging through methods like `getFormulaSource()` and `traceCell()`, but that source is engine-oriented debug information.

It is not a promise that Excellent preserves:

- the exact original Excel text formatting
- original whitespace
- original source text as a stable round-trip artifact

If you need precise source preservation, that is a different requirement than the current runtime is designed for.

## How To Read This Document

If a feature appears here, treat it as outside the intended scope unless the repo explicitly changes that contract later.

If a feature is missing from this document, that still does not imply support by default. The positive support contract is defined by:

- the public API
- the README
- the tests
- the implementation itself
