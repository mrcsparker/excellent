(function bootstrapDemo() {
  'use strict';

  var SNAPSHOT_KEY = 'excellent.demo.snapshot';
  var SAMPLE_WORKBOOKS = Object.freeze([
    {
      detail: 'Single-sheet formula sanity check',
      fileName: 'simpleFormula.xlsx',
      id: 'simpleFormula',
      label: 'Starter Formula',
      source: '../test/data/simpleFormula.xlsx'
    },
    {
      detail: 'Cross-sheet references and recalculation',
      fileName: 'crossSheetWorkbook.xlsx',
      id: 'crossSheetWorkbook',
      label: 'Cross-Sheet Model',
      source: '../test/data/crossSheetWorkbook.xlsx'
    },
    {
      detail: 'Quoted sheet names and Excel errors',
      fileName: 'quotedSheetAndErrors.xlsx',
      id: 'quotedSheetAndErrors',
      label: 'Errors + Quotes',
      source: '../test/data/quotedSheetAndErrors.xlsx'
    },
    {
      detail: 'Shared formula expansion across a sheet',
      fileName: 'sharedFormulas.xlsx',
      id: 'sharedFormulas',
      label: 'Shared Formulas',
      source: '../test/data/sharedFormulas.xlsx'
    }
  ]);
  var STATE_LABELS = Object.freeze({
    booting: 'Booting studio',
    error: 'Load failed',
    idle: 'Ready for workbook',
    loading: 'Parsing workbook',
    ready: 'Workbook loaded'
  });

  document.addEventListener('DOMContentLoaded', function onContentLoaded() {
    var demoApp = new DemoApp(document);

    demoApp.init();
  });

  function DemoApp(documentObject) {
    this.document = documentObject;
    this.window = documentObject.defaultView;
    this.Excel = this.window.Excellent;
    this.state = {
      activeSampleId: '',
      activeSheet: '',
      demoState: 'booting',
      lastError: '',
      lastParseDurationMs: null,
      selectedCell: '',
      showFormulas: false,
      sourceLabel: 'No workbook loaded yet',
      workbook: null,
      workbookName: ''
    };
    this.elements = getElements(documentObject);
  }

  DemoApp.prototype.init = function init() {
    if (!hasBrowserSupport(this.window)) {
      this.setError('This demo requires a browser with FileReader and ArrayBuffer support.');
      return;
    }

    if (!this.Excel || typeof this.Excel.XlsxReader !== 'function') {
      this.setError('The Excellent browser bundle did not load correctly.');
      return;
    }

    this.bindEvents();
    this.renderSampleButtons();
    this.setDemoState('idle');
    this.render();
  };

  DemoApp.prototype.bindEvents = function bindEvents() {
    var self = this;

    this.elements.fileInput.addEventListener('change', function onFileChange(event) {
      void self.handleFileUpload(event);
    });

    this.elements.loadStarterButton.addEventListener('click', function onStarterClick() {
      void self.loadSampleById(SAMPLE_WORKBOOKS[0].id);
    });

    this.elements.sampleList.addEventListener('click', function onSampleListClick(event) {
      var sampleButton = event.target.closest('[data-sample-id]');

      if (sampleButton === null || sampleButton.dataset.sampleId === undefined) {
        return;
      }

      void self.loadSampleById(sampleButton.dataset.sampleId);
    });

    this.elements.saveSnapshotButton.addEventListener('click', function onSaveSnapshot() {
      self.saveSnapshot();
    });

    this.elements.restoreSnapshotButton.addEventListener('click', function onRestoreSnapshot() {
      self.restoreSnapshot();
    });

    this.elements.showFormulasToggle.addEventListener('change', function onToggleChange() {
      self.state.showFormulas = self.elements.showFormulasToggle.checked;
      self.renderGrid();
      self.renderInspector();
      self.renderSelectedSummary();
      self.syncPublicState();
    });

    this.elements.sheetTabs.addEventListener('click', function onSheetTabClick(event) {
      var sheetButton = event.target.closest('[data-sheet-name]');

      if (sheetButton === null || sheetButton.dataset.sheetName === undefined) {
        return;
      }

      self.setActiveSheet(sheetButton.dataset.sheetName);
    });

    this.elements.sheetGridBody.addEventListener('click', function onGridClick(event) {
      var cellTarget = event.target.closest('[data-cell-address]');

      if (cellTarget === null || cellTarget.dataset.cellAddress === undefined) {
        return;
      }

      self.selectCell(cellTarget.dataset.cellAddress);
    });

    this.elements.sheetGridBody.addEventListener('focusin', function onGridFocus(event) {
      var cellTarget = event.target.closest('[data-cell-address]');

      if (cellTarget === null || cellTarget.dataset.cellAddress === undefined) {
        return;
      }

      self.selectCell(cellTarget.dataset.cellAddress);
    });

    this.elements.sheetGridBody.addEventListener('keydown', function onGridKeydown(event) {
      if (event.key !== 'Enter') {
        return;
      }

      if (!(event.target instanceof HTMLInputElement)) {
        return;
      }

      if (event.target.dataset.cellEditor === undefined) {
        return;
      }

      event.preventDefault();
      event.target.blur();
    });

    this.elements.sheetGridBody.addEventListener('change', function onGridChange(event) {
      if (!(event.target instanceof HTMLInputElement)) {
        return;
      }

      if (event.target.dataset.cellEditor === undefined) {
        return;
      }

      self.updateEditableCell(event.target.dataset.cellEditor, event.target.value);
    });

    this.elements.dependencyPanel.addEventListener('click', function onDependencyClick(event) {
      var jumpButton = event.target.closest('[data-jump-sheet]');

      if (jumpButton === null || jumpButton.dataset.jumpSheet === undefined || jumpButton.dataset.jumpCell === undefined) {
        return;
      }

      self.jumpToCell(jumpButton.dataset.jumpSheet, jumpButton.dataset.jumpCell);
    });
  };

  DemoApp.prototype.handleFileUpload = async function handleFileUpload(event) {
    var input = event.target;
    var file;

    if (!(input instanceof HTMLInputElement) || input.files === null) {
      return;
    }

    file = input.files[0];

    if (file === undefined) {
      return;
    }

    try {
      await this.loadWorkbookFromFile(file);
    } finally {
      input.value = '';
    }
  };

  DemoApp.prototype.loadWorkbookFromFile = async function loadWorkbookFromFile(file) {
    var bytes = await readFileAsArrayBuffer(file);

    await this.loadWorkbookFromBytes(bytes, file.name, 'Uploaded workbook');
  };

  DemoApp.prototype.loadSampleById = async function loadSampleById(sampleId) {
    var response;
    var sample = getSampleById(sampleId);

    if (sample === null) {
      return;
    }

    this.state.activeSampleId = sample.id;
    this.setDemoState('loading');
    this.renderSampleButtons();

    try {
      response = await this.window.fetch(sample.source);

      if (!response.ok) {
        throw new Error('Unable to load sample workbook: ' + sample.fileName);
      }

      await this.loadWorkbookFromBytes(
        await response.arrayBuffer(),
        sample.fileName,
        'Fixture sample'
      );
    } catch (error) {
      this.setError(error);
    }
  };

  DemoApp.prototype.loadWorkbookFromBytes = async function loadWorkbookFromBytes(bytes, fileName, sourceLabel) {
    var durationMs;
    var startedAt = this.window.performance.now();
    var workbookReader = new this.Excel.XlsxReader();

    this.setDemoState('loading');

    try {
      this.applyWorkbook(await workbookReader.load(bytes), {
        durationMs: this.window.performance.now() - startedAt,
        fileName: fileName,
        sourceLabel: sourceLabel
      });
    } catch (error) {
      durationMs = this.window.performance.now() - startedAt;
      this.state.lastParseDurationMs = durationMs;
      this.setError(error);
    }
  };

  DemoApp.prototype.applyWorkbook = function applyWorkbook(workbook, options) {
    var sheetNames = workbook.getSheetNames();
    var firstSheet = sheetNames[0] || '';

    this.state.workbook = workbook;
    this.state.workbookName = options.fileName;
    this.state.sourceLabel = options.sourceLabel;
    this.state.lastParseDurationMs = options.durationMs;
    this.state.lastError = '';
    this.state.activeSheet = firstSheet;
    this.state.selectedCell = this.getDefaultCellForSheet(firstSheet);
    this.setDemoState('ready');
    this.render();
  };

  DemoApp.prototype.saveSnapshot = function saveSnapshot() {
    var loader;
    var serialized;

    if (this.state.workbook === null) {
      return;
    }

    loader = new this.Excel.WorkbookLoader();
    serialized = loader.serialize(this.state.workbook);
    this.window.localStorage.setItem(SNAPSHOT_KEY, JSON.stringify(serialized));
    this.renderStatus();
    this.syncPublicState();
  };

  DemoApp.prototype.restoreSnapshot = function restoreSnapshot() {
    var loader;
    var snapshot = this.window.localStorage.getItem(SNAPSHOT_KEY);

    if (snapshot === null) {
      return;
    }

    try {
      loader = new this.Excel.WorkbookLoader();
      this.applyWorkbook(loader.deserialize(JSON.parse(snapshot)), {
        durationMs: null,
        fileName: 'Snapshot workbook',
        sourceLabel: 'Browser snapshot'
      });
      this.state.activeSampleId = '';
      this.renderSampleButtons();
    } catch (error) {
      this.setError(error);
    }
  };

  DemoApp.prototype.setActiveSheet = function setActiveSheet(sheetName) {
    if (this.state.workbook === null || !this.state.workbook.getSheetNames().includes(sheetName)) {
      return;
    }

    this.state.activeSheet = sheetName;
    this.state.selectedCell = this.getDefaultCellForSheet(sheetName);
    this.render();
  };

  DemoApp.prototype.selectCell = function selectCell(cellName) {
    this.state.selectedCell = cellName;
    this.renderGrid();
    this.renderInspector();
    this.renderSelectedSummary();
    this.syncPublicState();
  };

  DemoApp.prototype.jumpToCell = function jumpToCell(sheetName, cellName) {
    if (this.state.workbook === null) {
      return;
    }

    if (!this.state.workbook.getSheetNames().includes(sheetName)) {
      return;
    }

    this.state.activeSheet = sheetName;
    this.state.selectedCell = cellName;
    this.render();
  };

  DemoApp.prototype.updateEditableCell = function updateEditableCell(cellName, inputValue) {
    var worksheet = this.getActiveWorksheet();

    if (worksheet === null) {
      return;
    }

    worksheet.setCellValue(cellName, parseEditableValue(inputValue));
    this.state.selectedCell = cellName;
    this.render();
  };

  DemoApp.prototype.setDemoState = function setDemoState(nextState) {
    this.state.demoState = nextState;
    this.elements.body.dataset.demoState = nextState;
  };

  DemoApp.prototype.setError = function setError(error) {
    this.state.lastError = toErrorMessage(error);
    this.setDemoState('error');
    this.render();
  };

  DemoApp.prototype.hasSnapshot = function hasSnapshot() {
    return this.window.localStorage.getItem(SNAPSHOT_KEY) !== null;
  };

  DemoApp.prototype.getActiveWorksheet = function getActiveWorksheet() {
    if (this.state.workbook === null || this.state.activeSheet === '') {
      return null;
    }

    try {
      return this.state.workbook.requireSheet(this.state.activeSheet);
    } catch {
      return null;
    }
  };

  DemoApp.prototype.getSelectedCell = function getSelectedCell() {
    var worksheet = this.getActiveWorksheet();

    if (worksheet === null || this.state.selectedCell === '') {
      return null;
    }

    return worksheet.getCell(this.state.selectedCell) || null;
  };

  DemoApp.prototype.getDefaultCellForSheet = function getDefaultCellForSheet(sheetName) {
    var preferredCell;
    var worksheet = this.state.workbook === null ? null : this.state.workbook.getSheet(sheetName) || null;
    var cells;

    if (worksheet === null) {
      return '';
    }

    cells = worksheet.getCells().slice().sort(compareCells);
    preferredCell = cells.find(function findLiteralCell(cell) {
      return !cell.isFormula();
    }) || cells[0];

    return preferredCell === undefined ? '' : preferredCell.address;
  };

  DemoApp.prototype.getWorkbookStats = function getWorkbookStats() {
    var workbook = this.state.workbook;

    if (workbook === null) {
      return {
        formulaCount: 0,
        inputCount: 0,
        rowCount: 0,
        sheetCount: 0
      };
    }

    return workbook.getSheetNames().reduce(function reduceStats(stats, sheetName) {
      var worksheet = workbook.requireSheet(sheetName);

      stats.formulaCount += worksheet.functions.length;
      stats.inputCount += worksheet.variables.length;
      stats.rowCount += getRowCount(worksheet);
      stats.sheetCount += 1;
      return stats;
    }, {
      formulaCount: 0,
      inputCount: 0,
      rowCount: 0,
      sheetCount: 0
    });
  };

  DemoApp.prototype.render = function render() {
    this.renderStatus();
    this.renderSampleButtons();
    this.renderWorkspaceHeader();
    this.renderSheetTabs();
    this.renderGrid();
    this.renderInspector();
    this.renderSelectedSummary();
    this.syncPublicState();
  };

  DemoApp.prototype.renderStatus = function renderStatus() {
    var stateLabel = STATE_LABELS[this.state.demoState];

    if (stateLabel === undefined) {
      stateLabel = STATE_LABELS.idle;
    }

    this.elements.statusPill.textContent = stateLabel;
    this.elements.workbookLabel.textContent = this.state.workbookName || 'Awaiting workbook';
    this.elements.sourceLabel.textContent = this.state.lastError !== '' ?
      this.state.lastError :
      this.state.sourceLabel;
    this.elements.durationLabel.textContent = formatDuration(this.state.lastParseDurationMs);
    this.elements.snapshotLabel.textContent = this.hasSnapshot() ?
      'Snapshot available' :
      'No snapshot saved';
    this.elements.saveSnapshotButton.disabled = this.state.workbook === null;
    this.elements.restoreSnapshotButton.disabled = !this.hasSnapshot();
    this.elements.showFormulasToggle.checked = this.state.showFormulas;
  };

  DemoApp.prototype.renderSampleButtons = function renderSampleButtons() {
    var activeSampleId = this.state.activeSampleId;

    this.elements.sampleList.innerHTML = SAMPLE_WORKBOOKS.map(function mapSample(sample) {
      var activeClass = sample.id === activeSampleId ? ' sample-button--active' : '';

      return [
        '<button class="sample-button',
        activeClass,
        '" type="button" data-sample-id="',
        escapeAttribute(sample.id),
        '">',
        '<strong>',
        escapeHtml(sample.label),
        '</strong>',
        '<span>',
        escapeHtml(sample.detail),
        '</span>',
        '</button>'
      ].join('');
    }).join('');
  };

  DemoApp.prototype.renderWorkspaceHeader = function renderWorkspaceHeader() {
    var activeWorksheet = this.getActiveWorksheet();
    var stats = this.getWorkbookStats();
    var workbook = this.state.workbook;

    if (workbook === null || activeWorksheet === null) {
      this.elements.workspaceTitle.textContent = 'Ready for workbook';
      this.elements.workspaceDescription.textContent = 'Load a fixture or upload a workbook to render sheets, adjust inputs, and inspect formula behavior.';
      this.elements.metricStrip.innerHTML = createMetricStrip([
        createMetric('Sheets', '0'),
        createMetric('Inputs', '0'),
        createMetric('Formulas', '0'),
        createMetric('Engine', 'Browser')
      ]);
      return;
    }

    this.elements.workspaceTitle.textContent = activeWorksheet.name;
    this.elements.workspaceDescription.textContent = [
      stats.sheetCount,
      'sheets,',
      stats.inputCount,
      'editable inputs,',
      stats.formulaCount,
      'formula cells, and',
      stats.rowCount,
      'materialized rows in this workbook.'
    ].join(' ');
    this.elements.metricStrip.innerHTML = createMetricStrip([
      createMetric('Sheets', String(stats.sheetCount)),
      createMetric('Inputs', String(stats.inputCount)),
      createMetric('Formulas', String(stats.formulaCount)),
      createMetric('Engine', workbook.type || 'XLSX')
    ]);
  };

  DemoApp.prototype.renderSheetTabs = function renderSheetTabs() {
    var self = this;
    var workbook = this.state.workbook;

    if (workbook === null) {
      this.elements.sheetTabs.hidden = true;
      this.elements.sheetTabs.innerHTML = '';
      return;
    }

    this.elements.sheetTabs.hidden = false;
    this.elements.sheetTabs.innerHTML = workbook.getSheetNames().map(function mapSheet(sheetName) {
      var activeClass = sheetName === self.state.activeSheet ? ' sheet-tab--active' : '';

      return [
        '<button class="sheet-tab',
        activeClass,
        '" type="button" data-sheet-name="',
        escapeAttribute(sheetName),
        '">',
        escapeHtml(sheetName),
        '</button>'
      ].join('');
    }).join('');
  };

  DemoApp.prototype.renderGrid = function renderGrid() {
    var bounds;
    var worksheet = this.getActiveWorksheet();

    if (worksheet === null) {
      this.elements.emptyState.hidden = false;
      this.elements.sheetFrame.hidden = true;
      this.elements.sheetGridHead.innerHTML = '';
      this.elements.sheetGridBody.innerHTML = '';
      return;
    }

    bounds = getWorksheetBounds(worksheet);
    this.elements.emptyState.hidden = bounds.hasCells;
    this.elements.sheetFrame.hidden = !bounds.hasCells;

    if (!bounds.hasCells) {
      this.elements.sheetGridHead.innerHTML = '';
      this.elements.sheetGridBody.innerHTML = '';
      return;
    }

    this.elements.sheetGridHead.innerHTML = this.createGridHead(bounds.columnCount);
    this.elements.sheetGridBody.innerHTML = this.createGridBody(worksheet, bounds.rowCount, bounds.columnCount);
  };

  DemoApp.prototype.createGridHead = function createGridHead(columnCount) {
    var headerCells = ['<tr><th scope="col">#</th>'];
    var columnIndex;

    for (columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      headerCells.push('<th scope="col">' + escapeHtml(toColumnName(columnIndex + 1)) + '</th>');
    }

    headerCells.push('</tr>');
    return headerCells.join('');
  };

  DemoApp.prototype.createGridBody = function createGridBody(worksheet, rowCount, columnCount) {
    var rowIndex;
    var rows = [];

    for (rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      rows.push(this.createGridRow(worksheet, rowIndex, columnCount));
    }

    return rows.join('');
  };

  DemoApp.prototype.createGridRow = function createGridRow(worksheet, rowIndex, columnCount) {
    var columnIndex;
    var columns = ['<tr><th scope="row">' + String(rowIndex + 1) + '</th>'];

    for (columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      columns.push(this.createGridCell(worksheet, rowIndex, columnIndex));
    }

    columns.push('</tr>');
    return columns.join('');
  };

  DemoApp.prototype.createGridCell = function createGridCell(worksheet, rowIndex, columnIndex) {
    var address = toColumnName(columnIndex + 1) + String(rowIndex + 1);
    var cell = worksheet.getCell(address) || null;
    var isSelected = address === this.state.selectedCell;
    var selectedClass = isSelected ? ' is-selected' : '';

    if (cell === null) {
      return '<td class="sheet-cell sheet-cell--empty' + selectedClass + '"><div class="sheet-cell__ghost"></div></td>';
    }

    if (cell.isFormula()) {
      return [
        '<td class="sheet-cell sheet-cell--formula',
        selectedClass,
        '" data-cell-address="',
        escapeAttribute(address),
        '">',
        '<button class="sheet-cell__button" type="button" data-cell-address="',
        escapeAttribute(address),
        '">',
        '<span class="sheet-cell__value">',
        escapeHtml(formatDisplayValue(worksheet.getCellValue(address))),
        '</span>',
        this.state.showFormulas && cell.getFormulaSource() !== undefined ?
          '<span class="sheet-cell__formula">' + escapeHtml(cell.getFormulaSource()) + '</span>' :
          '',
        '<span class="sheet-cell__meta">',
        escapeHtml(address),
        '</span>',
        '</button>',
        '</td>'
      ].join('');
    }

    return [
      '<td class="sheet-cell sheet-cell--input',
      selectedClass,
      '" data-cell-address="',
      escapeAttribute(address),
      '">',
      '<label class="sheet-cell__editor" data-cell-address="',
      escapeAttribute(address),
      '">',
      '<input class="sheet-cell__input" type="text" data-cell-editor="',
      escapeAttribute(address),
      '" value="',
      escapeAttribute(formatInputValue(cell.getRawValue())),
      '">',
      '<span class="sheet-cell__meta">',
      escapeHtml(address),
      '</span>',
      '</label>',
      '</td>'
    ].join('');
  };

  DemoApp.prototype.renderSelectedSummary = function renderSelectedSummary() {
    var cell = this.getSelectedCell();
    var worksheet = this.getActiveWorksheet();

    if (worksheet === null || cell === null) {
      this.elements.selectedCellLabel.textContent = 'No cell selected';
      this.elements.selectedCellSubtitle.textContent = 'Choose a populated cell to inspect its inputs, outputs, and trace.';
      this.elements.gridToolbarNote.textContent = 'Computed cells update immediately when editable inputs change.';
      return;
    }

    this.elements.selectedCellLabel.textContent = worksheet.name + '!' + cell.address;
    this.elements.selectedCellSubtitle.textContent = cell.isFormula() ?
      'Formula cell with traceable dependencies and computed output.' :
      'Literal input cell. Update the value directly in the grid to recalculate downstream formulas.';
    this.elements.gridToolbarNote.textContent = cell.isFormula() ?
      'Use the dependency pills on the right to jump through the formula graph.' :
      'Literal changes invalidate dependents immediately across the workbook.';
  };

  DemoApp.prototype.renderInspector = function renderInspector() {
    var cell = this.getSelectedCell();
    var dependencyMarkup;
    var summaryMarkup;
    var traceMarkup;
    var workbook = this.state.workbook;
    var worksheet = this.getActiveWorksheet();

    if (workbook === null || worksheet === null || cell === null) {
      this.elements.inspectorSummary.innerHTML = [
        '<div class="summary-card">',
        '<div class="panel-label">Inspector</div>',
        '<span class="inspector-copy">Load a workbook and select a populated cell to see its current value, formula source, and trace output.</span>',
        '</div>'
      ].join('');
      this.elements.dependencyPanel.innerHTML = [
        '<div class="summary-card">',
        '<span class="dependency-copy">Precedents and dependents will appear here once a cell is selected.</span>',
        '</div>'
      ].join('');
      this.elements.tracePanel.innerHTML = [
        '<div class="summary-card">',
        '<span class="trace-note">Formula trace output will render here for computed cells.</span>',
        '</div>'
      ].join('');
      return;
    }

    summaryMarkup = this.createInspectorSummary(workbook, worksheet, cell);
    dependencyMarkup = this.createDependencyMarkup(workbook, worksheet.name, cell.address);
    traceMarkup = this.createTraceMarkup(workbook, worksheet, cell);

    this.elements.inspectorSummary.innerHTML = summaryMarkup;
    this.elements.dependencyPanel.innerHTML = dependencyMarkup;
    this.elements.tracePanel.innerHTML = traceMarkup;
  };

  DemoApp.prototype.createInspectorSummary = function createInspectorSummary(workbook, worksheet, cell) {
    var formulaSource = cell.getFormulaSource();
    var value = worksheet.getCellValue(cell.address);

    return [
      '<div class="summary-grid">',
      createSummaryCard('Cell', worksheet.name + '!' + cell.address, cell.isFormula() ? 'Computed formula cell' : 'Editable literal cell'),
      createSummaryCard('Current Value', formatDisplayValue(value), cell.isFormula() ? 'Resolved workbook output' : 'Raw literal stored in the worksheet'),
      createSummaryCard('Raw Input', formatDisplayValue(cell.getRawValue()), cell.isFormula() ? 'Compiled source is stored as the raw cell value' : 'Literal value used during recalculation'),
      createSummaryCard('Workbook Meta', workbook.type || 'XLSX', workbook.fileVersion === '' ? 'No file version metadata present' : workbook.fileVersion),
      '</div>',
      formulaSource !== undefined ?
        '<div class="formula-box"><div class="panel-label">Formula Source</div><code>' + escapeHtml(formulaSource) + '</code></div>' :
        '<div class="formula-box"><div class="panel-label">Formula Source</div><span class="inspector-copy">This cell is a literal input, so there is no formula source to display.</span></div>'
    ].join('');
  };

  DemoApp.prototype.createDependencyMarkup = function createDependencyMarkup(workbook, sheetName, cellName) {
    return [
      '<div class="dependency-group">',
      '<div class="panel-label">Precedents</div>',
      createDependencyList(workbook.getPrecedents(sheetName, cellName)),
      '</div>',
      '<div class="dependency-group">',
      '<div class="panel-label">Dependents</div>',
      createDependencyList(workbook.getDependents(sheetName, cellName)),
      '</div>'
    ].join('');
  };

  DemoApp.prototype.createTraceMarkup = function createTraceMarkup(workbook, worksheet, cell) {
    var trace;

    if (!cell.isFormula()) {
      return [
        '<div class="trace-box">',
        '<div class="panel-label">Literal Cell</div>',
        '<span class="trace-note">Literal cells do not produce a formula trace. Select a computed cell to inspect the evaluation tree.</span>',
        '</div>'
      ].join('');
    }

    trace = workbook.traceCell(worksheet.name, cell.address);

    return [
      '<div class="trace-box">',
      '<div class="panel-label">Expression</div>',
      '<code>',
      escapeHtml(trace.expression),
      '</code>',
      '</div>',
      '<div class="trace-box">',
      '<div class="panel-label">Trace Output</div>',
      '<pre>',
      escapeHtml(stringifyTrace(trace.evaluation)),
      '</pre>',
      '</div>'
    ].join('');
  };

  DemoApp.prototype.syncPublicState = function syncPublicState() {
    var cell = this.getSelectedCell();
    var worksheet = this.getActiveWorksheet();

    this.window.__EXCELLENT_DEMO__ = {
      activeSheet: this.state.activeSheet,
      demoState: this.state.demoState,
      selectedCell: this.state.selectedCell,
      selectedDisplayValue: worksheet === null || cell === null ?
        '' :
        formatDisplayValue(worksheet.getCellValue(cell.address)),
      selectedFormulaSource: cell === null ? '' : cell.getFormulaSource() || '',
      showFormulas: this.state.showFormulas,
      sourceLabel: this.state.sourceLabel,
      workbookName: this.state.workbookName
    };
  };

  function getElements(documentObject) {
    return {
      body: documentObject.body,
      dependencyPanel: getElement(documentObject, 'dependency-panel'),
      durationLabel: getElement(documentObject, 'duration-label'),
      emptyState: getElement(documentObject, 'empty-state'),
      fileInput: getElement(documentObject, 'file-input'),
      gridToolbarNote: getElement(documentObject, 'grid-toolbar-note'),
      inspectorSummary: getElement(documentObject, 'inspector-summary'),
      loadStarterButton: getElement(documentObject, 'load-starter-workbook'),
      metricStrip: getElement(documentObject, 'metric-strip'),
      restoreSnapshotButton: getElement(documentObject, 'restore-snapshot'),
      sampleList: getElement(documentObject, 'sample-list'),
      saveSnapshotButton: getElement(documentObject, 'save-snapshot'),
      selectedCellLabel: getElement(documentObject, 'selected-cell-label'),
      selectedCellSubtitle: getElement(documentObject, 'selected-cell-subtitle'),
      sheetFrame: getElement(documentObject, 'sheet-frame'),
      sheetGridBody: getElement(documentObject, 'sheet-grid-body'),
      sheetGridHead: getElement(documentObject, 'sheet-grid-head'),
      sheetTabs: getElement(documentObject, 'sheet-tabs'),
      showFormulasToggle: getElement(documentObject, 'show-formulas'),
      snapshotLabel: getElement(documentObject, 'snapshot-label'),
      sourceLabel: getElement(documentObject, 'source-label'),
      statusPill: getElement(documentObject, 'status-pill'),
      tracePanel: getElement(documentObject, 'trace-panel'),
      workbookLabel: getElement(documentObject, 'workbook-label'),
      workspaceDescription: getElement(documentObject, 'workspace-description'),
      workspaceTitle: getElement(documentObject, 'workspace-title')
    };
  }

  function getElement(documentObject, id) {
    var element = documentObject.getElementById(id);

    if (element === null) {
      throw new Error('Missing demo element: #' + id);
    }

    return element;
  }

  function hasBrowserSupport(windowObject) {
    return Boolean(windowObject.FileReader && windowObject.ArrayBuffer && windowObject.localStorage);
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function(resolve, reject) {
      var reader = new FileReader();

      reader.onerror = function onError() {
        reject(reader.error || new Error('Unable to read workbook.'));
      };

      reader.onload = function onLoad(event) {
        if (event.target === null || event.target.result === null) {
          reject(new Error('Workbook read returned no data.'));
          return;
        }

        resolve(event.target.result);
      };

      reader.readAsArrayBuffer(file);
    });
  }

  function getSampleById(sampleId) {
    return SAMPLE_WORKBOOKS.find(function findSample(sample) {
      return sample.id === sampleId;
    }) || null;
  }

  function createMetric(label, value) {
    return {
      label: label,
      value: value
    };
  }

  function createMetricStrip(metrics) {
    return metrics.map(function mapMetric(metric) {
      return [
        '<div class="metric">',
        '<span class="metric-label">',
        escapeHtml(metric.label),
        '</span>',
        '<span class="metric-value">',
        escapeHtml(metric.value),
        '</span>',
        '</div>'
      ].join('');
    }).join('');
  }

  function createSummaryCard(label, value, detail) {
    return [
      '<div class="summary-card">',
      '<div class="panel-label">',
      escapeHtml(label),
      '</div>',
      '<strong>',
      escapeHtml(value),
      '</strong>',
      '<span>',
      escapeHtml(detail),
      '</span>',
      '</div>'
    ].join('');
  }

  function createDependencyList(references) {
    if (references.length === 0) {
      return [
        '<div class="dependency-list">',
        '<span class="dependency-pill dependency-pill--empty">None</span>',
        '</div>'
      ].join('');
    }

    return [
      '<div class="dependency-list">',
      references.map(function mapReference(reference) {
        return [
          '<button class="dependency-pill" type="button" data-jump-sheet="',
          escapeAttribute(reference.sheetName),
          '" data-jump-cell="',
          escapeAttribute(reference.cellName),
          '">',
          escapeHtml(reference.sheetName + '!' + reference.cellName),
          '</button>'
        ].join('');
      }).join(''),
      '</div>'
    ].join('');
  }

  function getWorksheetBounds(worksheet) {
    var cells = worksheet.getCells();
    var maxColumnIndex;
    var maxRowIndex;

    if (cells.length === 0) {
      return {
        columnCount: 0,
        hasCells: false,
        rowCount: 0
      };
    }

    maxColumnIndex = Math.max.apply(null, cells.map(function mapColumn(cell) {
      return cell.columnIndex;
    }));
    maxRowIndex = Math.max.apply(null, cells.map(function mapRow(cell) {
      return cell.rowIndex;
    }));

    return {
      columnCount: maxColumnIndex + 1,
      hasCells: true,
      rowCount: maxRowIndex + 1
    };
  }

  function getRowCount(worksheet) {
    return worksheet.rows.reduce(function countRows(total, row) {
      return Array.isArray(row) ? total + 1 : total;
    }, 0);
  }

  function formatDuration(durationMs) {
    if (durationMs === null || durationMs === undefined) {
      return 'No parse run yet';
    }

    return Math.round(durationMs) + ' ms';
  }

  function formatDisplayValue(value) {
    if (value === undefined) {
      return '';
    }

    if (value === null) {
      return 'null';
    }

    if (typeof value === 'number' && Object.is(value, -0)) {
      return '0';
    }

    if (typeof value === 'string') {
      return value;
    }

    if (typeof value === 'boolean') {
      return String(value);
    }

    if (Array.isArray(value)) {
      return stringifyTrace(value);
    }

    if (window.Excellent && typeof window.Excellent.isExcelError === 'function' && window.Excellent.isExcelError(value)) {
      return String(value);
    }

    return stringifyTrace(value);
  }

  function formatInputValue(value) {
    if (value === undefined || value === null) {
      return '';
    }

    return String(value);
  }

  function parseEditableValue(value) {
    var trimmed = value.trim();

    if (trimmed === '') {
      return '';
    }

    if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(trimmed)) {
      return Number.parseFloat(trimmed);
    }

    return value;
  }

  function toColumnName(columnNumber) {
    var letter = String.fromCodePoint(65 + ((columnNumber - 1) % 26));
    var prefix = Math.floor((columnNumber - 1) / 26);

    if (prefix > 0) {
      return toColumnName(prefix) + letter;
    }

    return letter;
  }

  function compareCells(leftCell, rightCell) {
    if (leftCell.rowIndex !== rightCell.rowIndex) {
      return leftCell.rowIndex - rightCell.rowIndex;
    }

    return leftCell.columnIndex - rightCell.columnIndex;
  }

  function stringifyTrace(value) {
    var seen = new WeakSet();

    try {
      return JSON.stringify(value, function replaceTrace(_key, currentValue) {
        if (typeof currentValue === 'function') {
          return '[Function]';
        }

        if (typeof currentValue === 'object' && currentValue !== null) {
          if (window.Excellent && typeof window.Excellent.isExcelError === 'function' && window.Excellent.isExcelError(currentValue)) {
            return String(currentValue);
          }

          if (seen.has(currentValue)) {
            return '[Circular]';
          }

          seen.add(currentValue);
        }

        return currentValue;
      }, 2);
    } catch {
      return String(value);
    }
  }

  function toErrorMessage(error) {
    if (error instanceof Error) {
      return error.message;
    }

    return String(error);
  }

  function escapeHtml(value) {
    return value
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  function escapeAttribute(value) {
    return escapeHtml(value);
  }
}());
