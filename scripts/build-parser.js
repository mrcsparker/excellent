'use strict';

var fs = require('fs');
var path = require('path');
var peggy = require('peggy');

var rootDir = path.resolve(__dirname, '..');
var grammarPath = path.join(rootDir, 'grammar', 'formula_parser.pegjs');
var generatedDir = path.join(rootDir, 'generated');
var distGeneratedDir = path.join(rootDir, 'dist', 'generated');
var nodeOutputPath = path.join(generatedDir, 'excellent.parser.js');
var grammar = fs.readFileSync(grammarPath, 'utf8');

function writeParser(outputPath, options) {
  var source = peggy.generate(grammar, Object.assign({
    grammarSource: grammarPath,
    output: 'source'
  }, options));

  fs.writeFileSync(outputPath, source);
}

fs.rmSync(distGeneratedDir, { force: true, recursive: true });
fs.rmSync(generatedDir, { force: true, recursive: true });
fs.mkdirSync(generatedDir, { recursive: true });
writeParser(nodeOutputPath, {
  format: 'commonjs'
});
