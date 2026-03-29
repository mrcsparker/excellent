#!/usr/bin/env node

'use strict';

var fs = require('fs');
var path = require('path');
var excellentPackage = require('..');

function writeStdout(message) {
  process.stdout.write(message + '\n');
}

function writeStderr(message) {
  process.stderr.write(message + '\n');
}

function printUsage() {
  writeStdout('Usage: excellent <path-to.xlsx>');
}

async function main() {
  var inputPath = process.argv[2];

  if (!inputPath || inputPath === '--help' || inputPath === '-h') {
    printUsage();
    process.exit(inputPath ? 0 : 1);
  }

  var resolvedPath = path.resolve(process.cwd(), inputPath);
  var xlsxFile = fs.readFileSync(resolvedPath);
  var reader = new excellentPackage.XlsxReader();
  var parsed = await reader.load(xlsxFile);

  writeStdout(JSON.stringify(parsed.workbook, null, 2));
}

main().catch(function(err) {
  writeStderr(String(err && (err.stack || err)));
  process.exit(1);
});
