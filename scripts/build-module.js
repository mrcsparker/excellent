'use strict';

var path = require('path');
var esbuild = require('esbuild');

var rootDir = path.resolve(__dirname, '..');
var entryPoint = path.join(rootDir, 'src', 'index.ts');
var outputPath = path.join(rootDir, 'dist', 'index.mjs');
var externalDependencies = [
  '@formulajs/formulajs',
  '@xmldom/xmldom',
  'acorn',
  'jszip'
];

async function build() {
  await esbuild.build({
    bundle: true,
    entryPoints: [entryPoint],
    external: externalDependencies,
    format: 'esm',
    outfile: outputPath,
    platform: 'node',
    sourcemap: true,
    target: ['node20']
  });
}

build().catch(function(err) {
  process.stderr.write(String(err && (err.stack || err)) + '\n');
  process.exit(1);
});
