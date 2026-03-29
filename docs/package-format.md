# Package Format Decision

## Decision

Excellent ships as an ESM-first package with a CJS compatibility build.

That means the published package shape is:

- ESM as the primary runtime entry for modern JavaScript and TypeScript consumers
- CJS as a compatibility entry for `require('excellent')`
- one shared declaration surface
- a separate browser bundle path that stays outside the Node package-format decision

## Why This Direction

- The authored runtime is already modern TypeScript and the public API is designed around named exports.
- TypeScript consumers already use `import { Workbook, XlsxReader } from 'excellent'`.
- ESM-first is the cleaner long-term direction for modern Node.js, bundlers, and tooling.
- ESM-only would force a hard break while the repo still has a CLI, tests, and transitional JavaScript consumer coverage built around CJS.
- A compatibility CJS build keeps the migration professional instead of forcing unnecessary churn during the packaging cleanup.

## Published Layout

The package now publishes a shape like:

```json
{
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "import": "./dist/index.mjs",
      "require": "./dist/index.js"
    },
    "./browser": {
      "types": "./dist/index.d.ts",
      "import": "./browser.mjs",
      "require": "./browser.js"
    },
    "./package.json": "./package.json"
  }
}
```

Implementation details:

- `dist/index.mjs` is the native ESM entrypoint.
- `dist/index.js` remains the CJS compatibility entrypoint.
- `browser.mjs` and `browser.js` mirror that split for the `excellent/browser` subpath.
- browser bundles such as `dist/excellent.js` and `dist/excellent.min.js` remain generated outputs, not repository-maintained source files.

## Why `.mjs` Instead Of A Package-Wide `type: module`

The repo intentionally keeps the package default in CommonJS mode for now.

That avoids a broad repo-wide rewrite of maintainer scripts, the CLI, and test files while still publishing a stable native ESM entrypoint for consumers. It also keeps the dual-package behavior explicit instead of relying on package-wide mode switches.

## Non-Goals

This decision does not require:

- changing the browser IIFE bundle format
- converting repo-internal scripts or tests to ESM
