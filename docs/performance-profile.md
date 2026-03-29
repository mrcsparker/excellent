# XLSX Load Profile

This document records the XLSX load hotspots that drove the current optimization pass, plus the resulting post-change profile.

Command:

```bash
npm run profile:xlsx-load
```

Environment used for the measurements below:

- Node `v22.13.1`
- `darwin arm64`
- script defaults: `mixedRows=4000`, `sharedRows=8000`, `iterations=5`, `warmups=1`

## Optimization Summary

- The original hotspot was workbook construction, not XML parsing.
- The optimized loader now batches cache invalidation during workbook construction and avoids repeated worksheet bookkeeping in the hottest mutation path.
- After that change, the dominant cost moved from cell storage to XML parsing and formula compilation.
- That shift is exactly what we wanted before considering more aggressive architecture changes.

The optimization that landed was targeted to the measured hot path:

- batch workbook mutations during `XlsxReader.load()` and `WorkbookLoader.deserialize()`
- avoid invalidating formula caches per cell while building a fresh workbook
- avoid synchronizing new cells into worksheet rows twice
- add worksheet membership caches so `variables` / `functions` list maintenance stays O(1) for inserts

## Baseline Before Optimization

The initial profile showed:

- mixed workbook load: `790.91 ms`
- shared-formula workbook load: `659.92 ms`

Dominant phases at baseline:

- `worksheet.storeFormulaCell`: `359.41 ms` in the mixed case
- `worksheet.storeValueCell`: `357.61 ms` in the mixed case
- `worksheet.storeValueCell`: `532.05 ms` in the shared-formula case

That is why the optimization focused on workbook mutation overhead instead of `xmlToJson`.

## Current Profile: Mixed Workbook Load

Shape:

- `4000` rows
- `1` shared string per row
- `2` formulas per row

Average total load time: `114.22 ms`

Top phases:

- `worksheet.parseXml`: `37.10 ms` (`32.49%`)
- `worksheet.compileFormula`: `29.41 ms` (`25.75%`)
- `worksheet.storeFormulaCell`: `16.48 ms` (`14.43%`)
- `worksheet.storeValueCell`: `12.86 ms` (`11.26%`)
- `worksheet.xmlToJson`: `5.24 ms` (`4.59%`)

## Current Profile: Shared-Formula Workbook Load

Shape:

- `8000` rows
- `1` shared formula per row

Average total load time: `124.69 ms`

Top phases:

- `worksheet.parseXml`: `48.82 ms` (`39.15%`)
- `worksheet.compileFormula`: `23.46 ms` (`18.81%`)
- `worksheet.storeFormulaCell`: `16.53 ms` (`13.26%`)
- `worksheet.storeValueCell`: `12.07 ms` (`9.68%`)
- `worksheet.xmlToJson`: `7.07 ms` (`5.67%`)
- `worksheet.expandSharedFormulas`: `5.68 ms` (`4.56%`)

## Implication For The Next Item

The next performance item should not add random caching. The measured hotspot has moved, so the next decision should be architectural:

- if load performance needs to go further, the next likely frontier is XML parsing / workbook ingestion shape
- that makes the remaining backlog item about streaming or incremental load the right next place to look

## Second Profiling Pass: Larger Workloads

Command:

```bash
npm run profile:xlsx-load -- --mixed-rows 12000 --shared-rows 24000 --iterations 5 --warmups 1
```

Environment used for the measurements below:

- Node `v22.13.1`
- `darwin arm64`
- script flags: `mixedRows=12000`, `sharedRows=24000`, `iterations=5`, `warmups=1`

### Mixed Workbook Load

Shape:

- `12000` rows
- `1` shared string per row
- `2` formulas per row

Average total load time: `2240.63 ms`

Top phases:

- `worksheet.parseXml`: `751.40 ms` (`33.54%`)
- `worksheet.compileFormula`: `524.86 ms` (`23.42%`)
- `worksheet.storeFormulaCell`: `335.28 ms` (`14.96%`)
- `worksheet.storeValueCell`: `239.19 ms` (`10.68%`)
- `sharedStrings.parseXml`: `118.17 ms` (`5.27%`)
- `worksheet.xmlToJson`: `111.84 ms` (`4.99%`)
- `zip.read.worksheetXml`: `56.27 ms` (`2.51%`)

### Shared-Formula Workbook Load

Shape:

- `24000` rows
- `1` shared formula per row

Average total load time: `1684.13 ms`

Top phases:

- `worksheet.parseXml`: `649.84 ms` (`38.59%`)
- `worksheet.compileFormula`: `339.69 ms` (`20.17%`)
- `worksheet.storeFormulaCell`: `231.05 ms` (`13.72%`)
- `worksheet.storeValueCell`: `185.18 ms` (`11.00%`)
- `worksheet.xmlToJson`: `76.76 ms` (`4.56%`)
- `worksheet.expandSharedFormulas`: `53.62 ms` (`3.18%`)
- `zip.read.worksheetXml`: `37.51 ms` (`2.23%`)

### Evaluation Benchmark Cross-Check

Command:

```bash
npm run bench:evaluation -- --rows 6000 --shared-rows 12000 --iterations 5 --warmups 1
```

Results:

- large-sheet cold evaluation: `49.38 ms`
- large-sheet recalculation: `66.06 ms`
- shared-formula cold evaluation: `80.26 ms`
- shared-formula recalculation: `136.49 ms`

That cross-check matters because it shows the engine is not spending seconds in recalculation after load. The load bottleneck is still ingestion work, especially XML parse plus formula compilation.

## Streaming Decision

Decision date: `2026-03-29`

We should not build true streaming XLSX support right now.

Why:

- zip reads are a small slice of total load time at larger workbook sizes
  - mixed case: `zip.read.worksheetXml` is only `2.51%`
  - shared-formula case: `zip.read.worksheetXml` is only `2.23%`
- the dominant costs are still CPU-heavy parse and compile phases
  - `worksheet.parseXml`
  - `worksheet.compileFormula`
  - `worksheet.xmlToJson`
- true streaming would require replacing core assumptions in the current stack
  - `jszip` container handling
  - DOM-style XML parsing
  - shared-string handling
  - shared-formula expansion
  - browser parity for the same API
- the current [`loadIncremental()`](../README.md#load-incrementally-by-sheet) API already provides the practical benefit most consumers need: worksheet-at-a-time observation and control over when sheet-level work happens

Revisit this only if future measurements on real workloads show that container materialization or retained memory, not XML parse plus formula compilation, is the actual blocker.
