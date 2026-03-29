'use strict';

var fs = require('fs');
var path = require('path');

var rootDir = path.resolve(__dirname, '..');
var buildArtifacts = [
  'dist',
  'generated',
  'lib'
];

for (var index = 0; index < buildArtifacts.length; index += 1) {
  fs.rmSync(path.join(rootDir, buildArtifacts[index]), {
    force: true,
    recursive: true
  });
}
