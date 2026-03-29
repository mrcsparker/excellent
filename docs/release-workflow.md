# Release And Build Workflow

This document is the maintainer-facing workflow for building, validating, packaging, and publishing Excellent.

The goal is simple: one source tree, one generated output tree, and a release process that can be repeated by someone who did not originally write the repo.

## Repo Shape

Authoritative source lives here:

- runtime source: `src/**/*.ts`
- formula grammar: `grammar/formula_parser.pegjs`
- build scripts: `scripts/*.js`
- tests: `test/**`

Generated output lives here:

- Node runtime output: `dist/*.js`
- native ESM runtime entry: `dist/index.mjs`
- emitted declarations: `dist/*.d.ts`
- source maps: `dist/*.js.map`
- ESM source map: `dist/index.mjs.map`
- generated parser runtime: `generated/excellent.parser.js`
- browser bundles: `dist/excellent.js` and `dist/excellent.min.js`

Important rule:

- `dist/` is generated output and is ignored in git
- `generated/` is generated output and is ignored in git
- source changes belong in `src/`, `grammar/`, `scripts/`, `test/`, or `docs/`, not in `dist/` or `generated/`

## Build Graph

The build is intentionally ordered because the TypeScript runtime depends on the generated parser.

### `npm run build:parser`

Runs:

```bash
node ./scripts/build-parser.js
```

What it does:

- reads [grammar/formula_parser.pegjs](/Users/mrcsparker/Documents/GitHub/excellent/grammar/formula_parser.pegjs)
- generates the Peggy parser runtime
- writes it to [generated/excellent.parser.js](/Users/mrcsparker/Documents/GitHub/excellent/generated/excellent.parser.js)

Why it must run first:

- [src/excellent.parser.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/excellent.parser.ts) imports that generated parser from `../generated/excellent.parser.js`
- runtime compilation is therefore not valid until the parser has already been generated

### `npm run build:runtime`

Runs:

```bash
npm run build:clean
npm run build:runtime:core
```

What it does:

- clears stale build output from previous layouts
- runs the runtime compilation pipeline
- keeps deleted modules from lingering in `dist/`

### `npm run build:runtime:core`

Runs:

```bash
npm run build:parser
tsc -p ./tsconfig.runtime.json
node ./scripts/build-module.js
```

What it does:

- generates the parser
- compiles `src/**/*.ts` into CJS runtime files under `dist/`
- bundles a native ESM package entry at `dist/index.mjs`
- emits `.d.ts` files from source

This produces the Node/runtime package surface used by:

- [index.js](/Users/mrcsparker/Documents/GitHub/excellent/index.js)
- [browser.mjs](/Users/mrcsparker/Documents/GitHub/excellent/browser.mjs)
- [browser.js](/Users/mrcsparker/Documents/GitHub/excellent/browser.js)
- the CLI in [bin/excellent.js](/Users/mrcsparker/Documents/GitHub/excellent/bin/excellent.js)

### `npm run build:browser`

Runs:

```bash
npm run build:parser
node ./scripts/build-browser.js
```

What it does:

- regenerates the parser first
- bundles [src/index.ts](/Users/mrcsparker/Documents/GitHub/excellent/src/index.ts) with esbuild
- writes browser IIFE bundles to:
  - [dist/excellent.js](/Users/mrcsparker/Documents/GitHub/excellent/dist/excellent.js)
  - [dist/excellent.min.js](/Users/mrcsparker/Documents/GitHub/excellent/dist/excellent.min.js)

### `npm run build`

Runs the full distributable build:

```bash
npm run build:runtime
npm run build:browser
```

Use this when you want the complete package output, including the browser bundle.

## Quality Gates

### `npm test`

This is the main gate.

It runs:

- `npm run lint`
- the Node test suites under `test/evaluator`, `test/parser`, `test/public-api`, `test/workbook-model`, and `test/xlsx-integration`

### `npm run test:browser`

Builds the browser bundle and runs the Playwright/browser smoke suites.

Use this when browser packaging or the demo changes.

### `npm run coverage`

Runs lint and then emits remapped coverage data with `c8`.

Useful when changing parser/evaluator/workbook behavior and wanting to see what code still lacks direct test coverage.

## Fixture And Benchmark Utilities

### `npm run build:test-fixtures`

Regenerates the committed XLSX fixtures in `test/data/`.

Use this only when you intentionally change the fixture generator or fixture expectations.

### `npm run bench:evaluation`

Runs the evaluation benchmark harness.

Use this when changing evaluator or workbook performance characteristics.

### `npm run profile:xlsx-load`

Profiles XLSX loading phases.

Use this when changing workbook ingestion, worksheet mutation behavior, XML parsing, or load-path performance.

## Packaging Rules

The published package is controlled by the `files` whitelist in [package.json](/Users/mrcsparker/Documents/GitHub/excellent/package.json).

Only these categories are meant to ship:

- `dist/`
- `generated/`
- `browser.mjs`
- `index.js`
- `browser.js`
- `bin/`
- `README.md`
- `LICENSE.txt`

That means the tarball should not include:

- `src/`
- `test/`
- `docs/`
- local Playwright artifacts
- old demo/vendor scaffolding

## `prepare` vs `prepack`

The repo uses both on purpose.

### `prepare`

`prepare` runs:

```bash
npm run build:runtime:core
```

Reason:

- local installs from the repo or git need the Node/runtime output
- that keeps the package usable when installed directly from source
- it intentionally does not clean `dist/`, so it does not delete browser bundles that `prepack` already generated for packaging

### `prepack`

`prepack` runs:

```bash
npm run build
```

Reason:

- an npm tarball needs the full published output, including the browser bundles
- that makes `npm pack` and `npm publish` build the actual distributable package right before packaging

## Release Checklist

Use this exact flow.

### 1. Start from a clean working tree

Make sure your intended changes are committed and you are not accidentally packaging local junk.

### 2. Run the quality gates

```bash
npm test
npm run test:browser
npm run build
```

If you changed evaluator or loader behavior materially, also consider:

```bash
npm run coverage
npm run bench:evaluation
npm run profile:xlsx-load
```

### 3. Inspect the package contents

Always run:

```bash
npm pack --dry-run
```

Confirm that:

- `dist/` artifacts are present
- `README.md`, `LICENSE.txt`, `browser.mjs`, `index.js`, `browser.js`, and `bin/` are present
- `src/`, `test/`, `docs/`, and local artifact directories are absent

Do not publish without checking this. The tarball is the actual release, not the working tree.

### 4. Bump the version

Update the package version in [package.json](/Users/mrcsparker/Documents/GitHub/excellent/package.json) using your normal release/versioning process.

### 5. Pack locally if you want a final sanity check

```bash
npm pack
```

Optionally install that tarball into a scratch project and verify:

- `import('excellent')` works
- `require('excellent')` works
- TypeScript resolves declarations
- the browser bundle exists under `dist/` in the package

### 6. Publish

After the dry-run/package check is clean:

```bash
npm publish
```

## When To Touch Which Files

### Formula grammar changes

Edit:

- [grammar/formula_parser.pegjs](/Users/mrcsparker/Documents/GitHub/excellent/grammar/formula_parser.pegjs)

Then run:

```bash
npm run build:parser
npm test
```

### Runtime/API changes

Edit:

- `src/**/*.ts`

Then run:

```bash
npm test
npm run test:browser
```

### Browser/demo changes

Edit:

- [demo/index.html](/Users/mrcsparker/Documents/GitHub/excellent/demo/index.html)
- [demo/demo.css](/Users/mrcsparker/Documents/GitHub/excellent/demo/demo.css)
- [demo/scripts/demo.js](/Users/mrcsparker/Documents/GitHub/excellent/demo/scripts/demo.js)
- possibly [scripts/build-browser.js](/Users/mrcsparker/Documents/GitHub/excellent/scripts/build-browser.js)

Then run:

```bash
npm run test:browser
npm run build
```

### Fixture changes

Edit:

- [scripts/generate-test-fixtures.js](/Users/mrcsparker/Documents/GitHub/excellent/scripts/generate-test-fixtures.js)

Then run:

```bash
npm run build:test-fixtures
npm test
```

## Maintainer Notes

- Do not hand-edit files in `dist/`.
- Do not treat generated parser output as authored source.
- Do not publish based on assumptions about what npm will include; always check `npm pack --dry-run`.
- If release packaging changes, update this document and [README.md](/Users/mrcsparker/Documents/GitHub/excellent/README.md) together.
