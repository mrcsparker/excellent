var assert = require('node:assert');
var test = require('node:test');
var describe = test.describe;
var it = test.it;
var packageName = 'excellent';

var packageByName = require(packageName);
var packageByPath = require('../..');
var browserEntry = require(packageName + '/browser');
var packageManifest = require(packageName + '/package.json');

describe('ExcellentPackageExports', function() {
  'use strict';

  it('resolves the package root through the exports map', function() {
    assert.equal(packageByName.Workbook, packageByPath.Workbook);
    assert.equal(packageByName.XlsxReader, packageByPath.XlsxReader);
  });

  it('resolves the browser subpath through the exports map', function() {
    assert.equal(browserEntry.Workbook, packageByPath.Workbook);
    assert.equal(browserEntry.XlsxReader, packageByPath.XlsxReader);
  });

  it('allows explicit package metadata access through the exports map', function() {
    assert.equal(packageManifest.name, 'excellent');
    assert.equal(packageManifest.exports['.'].import, './dist/index.mjs');
    assert.equal(packageManifest.exports['./browser'].import, './browser.mjs');
    assert.equal(packageManifest.exports['./browser'].require, './browser.js');
  });
});
