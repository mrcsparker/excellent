var assert = require('node:assert');
var fs = require('node:fs');
var http = require('node:http');
var path = require('node:path');
var test = require('node:test');
var after = test.after;
var before = test.before;
var describe = test.describe;
var it = test.it;
var playwright = require('playwright');

var REPO_ROOT = path.resolve(__dirname, '..', '..');
var HOST = '127.0.0.1';

function getContentType(filePath) {
  if (filePath.endsWith('.html')) {
    return 'text/html; charset=utf-8';
  }

  if (filePath.endsWith('.js')) {
    return 'application/javascript; charset=utf-8';
  }

  if (filePath.endsWith('.xlsx')) {
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }

  return 'application/octet-stream';
}

function serveFile(response, filePath) {
  var content = fs.readFileSync(filePath);

  response.writeHead(200, {
    'content-length': String(content.length),
    'content-type': getContentType(filePath)
  });
  response.end(content);
}

function createSmokeHtml(bundleName) {
  return [
    '<!doctype html>',
    '<html lang="en">',
    '<head><meta charset="utf-8"><title>Excellent Browser Smoke</title></head>',
    '<body data-status="booting">',
    '<pre id="status">booting</pre>',
    '<script src="/dist/' + bundleName + '"></script>',
    '<script>',
    '(async function runSmoke(){',
    '  var statusNode = document.getElementById("status");',
    '  try {',
    '    if (typeof window.Excellent !== "object" || window.Excellent === null) {',
    '      throw new Error("Missing Excellent browser global.");',
    '    }',
    '    if (typeof window.Excellent.Workbook !== "function") {',
    '      throw new Error("Missing Workbook export on browser bundle.");',
    '    }',
    '    if (typeof window.Excellent.XlsxReader !== "function") {',
    '      throw new Error("Missing XlsxReader export on browser bundle.");',
    '    }',
    '    if (typeof window.Excellent.XLSX_READER_MODE !== "object") {',
    '      throw new Error("Missing XLSX_READER_MODE export on browser bundle.");',
    '    }',
    '    var workbook = new window.Excellent.Workbook();',
    '    var sheet = workbook.createSheet("Inline");',
    '    sheet.setCellValue("A1", 2);',
    '    sheet.setCellFormula("A2", "this.A1+1");',
    '    var response = await fetch("/test/data/simpleFormula.xlsx");',
    '    var bytes = await response.arrayBuffer();',
    '    var parsed = await new window.Excellent.XlsxReader().load(bytes);',
    '    window.__EXCELLENT_SMOKE__ = {',
    '      bundleName: "' + bundleName + '",',
    '      inlineFormulaValue: workbook.getCellValue("Inline", "A2"),',
    '      mode: window.Excellent.XLSX_READER_MODE.FORMULAS,',
    '      parsedFormulaSource: parsed.getFormulaSource("Sheet1", "A1"),',
    '      parsedSheetNames: parsed.getSheetNames(),',
    '      parsedValue: parsed.getCellValue("Sheet1", "A1")',
    '    };',
    '    document.body.dataset.status = "ready";',
    '    statusNode.textContent = JSON.stringify(window.__EXCELLENT_SMOKE__);',
    '  } catch (error) {',
    '    document.body.dataset.status = "error";',
    '    statusNode.textContent = String(error && (error.stack || error.message || error));',
    '  }',
    '}());',
    '</script>',
    '</body>',
    '</html>'
  ].join('');
}

function createSmokeServer() {
  return http.createServer(function handleRequest(request, response) {
    var url = new URL(request.url, 'http://' + HOST);

    if (url.pathname === '/smoke') {
      var bundleName = url.searchParams.get('bundle') || 'excellent.js';

      response.writeHead(200, {
        'content-type': 'text/html; charset=utf-8'
      });
      response.end(createSmokeHtml(bundleName));
      return;
    }

    if (url.pathname === '/dist/excellent.js' || url.pathname === '/dist/excellent.min.js') {
      serveFile(response, path.join(REPO_ROOT, url.pathname));
      return;
    }

    if (url.pathname === '/test/data/simpleFormula.xlsx') {
      serveFile(response, path.join(REPO_ROOT, url.pathname));
      return;
    }

    response.writeHead(404, {
      'content-type': 'text/plain; charset=utf-8'
    });
    response.end('Not found: ' + url.pathname);
  });
}

async function waitForListen(server) {
  return await new Promise(function(resolve) {
    server.listen(0, HOST, function() {
      var address = server.address();

      if (address === null || typeof address === 'string') {
        throw new Error('Unable to determine smoke server address.');
      }

      resolve(address.port);
    });
  });
}

describe('ExcellentBrowserBundle', function() {
  'use strict';

  var baseUrl;
  var browser;
  var server;

  before(async function() {
    server = createSmokeServer();
    baseUrl = 'http://' + HOST + ':' + String(await waitForListen(server));
    browser = await playwright.chromium.launch({
      headless: true
    });
  });

  after(async function() {
    await browser.close();
    await new Promise(function(resolve, reject) {
      server.close(function(error) {
        if (error) {
          reject(error);
          return;
        }

        resolve();
      });
    });
  });

  ['excellent.js', 'excellent.min.js'].forEach(function(bundleName) {
    it('loads ' + bundleName + ' in a real browser and parses an xlsx fixture', async function() {
      var context = await browser.newContext();
      var page = await context.newPage();

      try {
        await page.goto(baseUrl + '/smoke?bundle=' + bundleName, {
          waitUntil: 'load'
        });
        await page.waitForFunction(function() {
          return globalThis.document.body.dataset.status === 'ready' ||
            globalThis.document.body.dataset.status === 'error';
        });

        var status = await page.getAttribute('body', 'data-status');
        var result = await page.evaluate(function() {
          return globalThis.window.__EXCELLENT_SMOKE__ || null;
        });
        var errorText;

        if (status !== 'ready') {
          errorText = await page.textContent('#status');
          throw new Error(errorText || 'Browser smoke page failed without an error message.');
        }

        assert.deepEqual(result, {
          bundleName: bundleName,
          inlineFormulaValue: 3,
          mode: 'formulas',
          parsedFormulaSource: 'Formula.SUM(1,2)',
          parsedSheetNames: ['Sheet1'],
          parsedValue: 3
        });
      } finally {
        await context.close();
      }
    });
  });
});
