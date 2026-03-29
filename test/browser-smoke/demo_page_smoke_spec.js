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

var HOST = '127.0.0.1';
var REPO_ROOT = path.resolve(__dirname, '..', '..');

function getContentType(filePath) {
  if (filePath.endsWith('.css')) {
    return 'text/css; charset=utf-8';
  }

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

function createDemoServer() {
  return http.createServer(function handleRequest(request, response) {
    var requestUrl = new URL(request.url, 'http://' + HOST);

    if (requestUrl.pathname === '/' || requestUrl.pathname === '/demo' || requestUrl.pathname === '/demo/index.html') {
      serveFile(response, path.join(REPO_ROOT, 'demo', 'index.html'));
      return;
    }

    if (requestUrl.pathname === '/demo/demo.css') {
      serveFile(response, path.join(REPO_ROOT, 'demo', 'demo.css'));
      return;
    }

    if (requestUrl.pathname === '/demo/scripts/demo.js') {
      serveFile(response, path.join(REPO_ROOT, 'demo', 'scripts', 'demo.js'));
      return;
    }

    if (requestUrl.pathname === '/dist/excellent.js') {
      serveFile(response, path.join(REPO_ROOT, 'dist', 'excellent.js'));
      return;
    }

    if (requestUrl.pathname.startsWith('/test/data/') && requestUrl.pathname.endsWith('.xlsx')) {
      serveFile(response, path.join(REPO_ROOT, requestUrl.pathname));
      return;
    }

    response.writeHead(404, {
      'content-type': 'text/plain; charset=utf-8'
    });
    response.end('Not found: ' + requestUrl.pathname);
  });
}

async function waitForListen(server) {
  return await new Promise(function(resolve) {
    server.listen(0, HOST, function() {
      var address = server.address();

      if (address === null || typeof address === 'string') {
        throw new Error('Unable to determine demo smoke server address.');
      }

      resolve(address.port);
    });
  });
}

describe('ExcellentDemoPage', function() {
  'use strict';

  var baseUrl;
  var browser;
  var server;

  before(async function() {
    server = createDemoServer();
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

  it('loads the modern demo shell and opens a fixture workbook', async function() {
    var context = await browser.newContext();
    var page = await context.newPage();
    var state;

    try {
      await page.goto(baseUrl + '/demo/index.html', {
        waitUntil: 'load'
      });
      await page.waitForFunction(function() {
        return globalThis.document.body.dataset.demoState === 'idle' ||
          globalThis.document.body.dataset.demoState === 'error';
      });

      assert.equal(await page.textContent('#workspace-title'), 'Ready for workbook');
      await page.click('[data-sample-id="simpleFormula"]');
      await page.waitForFunction(function() {
        return globalThis.document.body.dataset.demoState === 'ready' ||
          globalThis.document.body.dataset.demoState === 'error';
      });

      state = await page.evaluate(function() {
        return globalThis.window.__EXCELLENT_DEMO__ || null;
      });

      assert.deepEqual(state, {
        activeSheet: 'Sheet1',
        demoState: 'ready',
        selectedCell: 'A1',
        selectedDisplayValue: '3',
        selectedFormulaSource: 'Formula.SUM(1,2)',
        showFormulas: false,
        sourceLabel: 'Fixture sample',
        workbookName: 'simpleFormula.xlsx'
      });

      assert.equal(await page.textContent('#workspace-title'), 'Sheet1');
      assert.equal(await page.textContent('#selected-cell-label'), 'Sheet1!A1');
    } finally {
      await context.close();
    }
  });
});
