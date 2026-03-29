'use strict';

var path = require('path');
var esbuild = require('esbuild');

var rootDir = path.resolve(__dirname, '..');
var entryPoint = path.join(rootDir, 'src', 'index.ts');

async function build() {
  var baseOptions = {
    bundle: true,
    entryPoints: [entryPoint],
    format: 'iife',
    globalName: 'Excellent',
    platform: 'browser',
    target: ['es2020']
  };

  await esbuild.build(Object.assign({}, baseOptions, {
    minify: false,
    outfile: path.join(rootDir, 'dist', 'excellent.js')
  }));

  await esbuild.build(Object.assign({}, baseOptions, {
    minify: true,
    outfile: path.join(rootDir, 'dist', 'excellent.min.js')
  }));
}

build().catch(function(err) {
  process.stderr.write(String(err && (err.stack || err)) + '\n');
  process.exit(1);
});
