'use strict';

import JSZip from 'jszip';
import { compileFormula } from './formula';
import type {
  CompiledFormula,
  FormulaEvaluator,
  FormulaFunction,
  FormulaFunctionRegistry,
  WorkbookCellValue
} from './formula';
import { FormulaParser } from './excellent.parser';
import { Util } from './excellent.util';
import { Workbook } from './workbook';
import type { Worksheet } from './workbook';
import {
  SharedStrings,
  getDomParserConstructor,
  incrementProfileCount,
  loadWorkbookFromZip,
  measureProfileAsync,
  measureProfileSync,
  normalizeCollection,
  readZipText,
  type XlsxLoadProfiler,
  type XlsxSheetLoadEvent,
  type XlsxSheetLoadHandler,
  type XmlTextNode
} from './excellent.xlsx.shared';

const XLSX_READER_MODE = Object.freeze({
  FORMULAS: 'formulas',
  VALUES_ONLY: 'values-only'
});

type ReaderMode = typeof XLSX_READER_MODE[keyof typeof XLSX_READER_MODE];
type DomParserConstructor = new () => DOMParser;
type WorksheetXmlLoaderOptions = {
  domParserCtor: DomParserConstructor;
  mode: ReaderMode;
  profile: XlsxLoadProfiler | undefined;
  sharedStrings: SharedStrings;
  worksheet: Worksheet;
  xmlString: string;
};
type XlsxReaderOptions = {
  DOMParser?: DomParserConstructor;
  formulaEvaluator?: FormulaEvaluator;
  functionRegistry?: FormulaFunctionRegistry;
  functions?: Record<string, FormulaFunction>;
  mode?: ReaderMode;
  profile?: XlsxLoadProfiler;
};
type WorksheetColumnAttributes = {
  r?: string;
  t?: string;
};
type WorksheetFormulaAttributes = {
  ref?: string;
  si?: string;
  t?: string;
};
type WorksheetFormulaNode = XmlTextNode & {
  '@'?: WorksheetFormulaAttributes;
};
type WorksheetColumnNode = {
  '@'?: WorksheetColumnAttributes;
  f?: WorksheetFormulaNode;
  v?: XmlTextNode;
};
type WorksheetRowNode = {
  c?: WorksheetColumnNode | WorksheetColumnNode[];
};
type WorksheetSheetDataNode = {
  row?: WorksheetRowNode | WorksheetRowNode[];
};
type WorksheetXmlDocument = {
  worksheet: {
    sheetData: WorksheetSheetDataNode;
  };
};
type WorksheetColumnDefinition = {
  cellName: string;
  formulaText: string | null;
  value: WorkbookCellValue | undefined;
};

type XlsxInput = Parameters<typeof JSZip.loadAsync>[0];

function createWorkbookOptions(options?: XlsxReaderOptions): XlsxReader['workbookOptions'] {
  const workbookOptions: XlsxReader['workbookOptions'] = {};

  if (options?.formulaEvaluator !== undefined) {
    workbookOptions.formulaEvaluator = options.formulaEvaluator;
  }

  if (options?.functionRegistry !== undefined) {
    workbookOptions.functionRegistry = options.functionRegistry;
  }

  if (options?.functions !== undefined) {
    workbookOptions.functions = options.functions;
  }

  return workbookOptions;
}

function getRequiredMatch(value: string, pattern: RegExp, label: string): string {
  const match = value.match(pattern);

  if (match === null || match[0] === undefined) {
    throw new Error('Unable to parse ' + label + ' from "' + value + '".');
  }

  return match[0];
}

function normalizeReaderMode(mode: ReaderMode | undefined) {
  if (mode === undefined) {
    return XLSX_READER_MODE.FORMULAS;
  }

  if (mode === XLSX_READER_MODE.FORMULAS || mode === XLSX_READER_MODE.VALUES_ONLY) {
    return mode;
  }

  throw new TypeError('Unsupported XLSX reader mode: ' + String(mode));
}

function uniqueValues(values: string[]) {
  return Array.from(new Set(values));
}

function buildFormulaExpression(formula: string, profile: XlsxLoadProfiler | undefined): CompiledFormula {
  let normalizedFormula = formula.trim();

  if (normalizedFormula.startsWith('+')) {
    normalizedFormula = normalizedFormula.slice(1);
  }

  return measureProfileSync(profile, 'worksheet.compileFormula', function compileWorksheetFormula() {
    return compileFormula(FormulaParser.parse(normalizedFormula));
  });
}

class WorksheetXmlLoader {
  domParserCtor: DomParserConstructor;
  mode: ReaderMode;
  profile: XlsxLoadProfiler | undefined;
  sharedStrings: SharedStrings;
  worksheet: Worksheet;
  xmlString: string;

  constructor(options: WorksheetXmlLoaderOptions) {
    this.domParserCtor = options.domParserCtor;
    this.mode = options.mode;
    this.profile = options.profile;
    this.sharedStrings = options.sharedStrings;
    this.worksheet = options.worksheet;
    this.xmlString = options.xmlString;
  }

  addValueCell(cellName: string, cellValue: WorkbookCellValue) {
    incrementProfileCount(this.profile, 'worksheet.cells.value');
    measureProfileSync(this.profile, 'worksheet.storeValueCell', () => {
      this.worksheet.setCellValue(cellName, cellValue);
    });
  }

  addFormulaCell(cellName: string, formulaText: string) {
    incrementProfileCount(this.profile, 'worksheet.cells.formula');
    const compiledFormula = buildFormulaExpression(formulaText.replace(/\$/g, ''), this.profile);

    measureProfileSync(this.profile, 'worksheet.storeFormulaCell', () => {
      this.worksheet.setCellFormula(cellName, compiledFormula);
    });
  }

  buildColumn(column: WorksheetColumnNode | undefined) {
    const columnDefinition = this.getColumnDefinition(column);

    if (columnDefinition === null) {
      incrementProfileCount(this.profile, 'worksheet.cells.skipped');
      return;
    }

    if (column !== undefined && this.shouldPopulateSharedFormulas(column)) {
      this.populateSharedFormulas(column);
      return;
    }

    if (this.mode === XLSX_READER_MODE.FORMULAS && columnDefinition.formulaText !== null) {
      this.addFormulaCell(columnDefinition.cellName, columnDefinition.formulaText);
      return;
    }

    if (columnDefinition.value === undefined) {
      incrementProfileCount(this.profile, 'worksheet.cells.skipped');
      return;
    }

    this.addValueCell(columnDefinition.cellName, columnDefinition.value);
  }

  getColumnDefinition(column: WorksheetColumnNode | undefined): WorksheetColumnDefinition | null {
    if (column === undefined) {
      return null;
    }

    const columnId = column['@'];

    if (columnId === undefined) {
      return null;
    }

    const cellName = columnId.r;
    const cellType = columnId.t;

    if (cellName === undefined) {
      return null;
    }

    if (cellType === 's') {
      const sharedStringIndex = column.v?.['#text'];

      if (sharedStringIndex === undefined) {
        return null;
      }

      return {
        cellName,
        formulaText: null,
        value: this.sharedStrings.get(parseInt(sharedStringIndex, 10))
      };
    }

    const formulaText = column.f?.['#text'];

    if (formulaText !== undefined) {
      const valueText = column.v?.['#text'];

      return {
        cellName,
        formulaText: formulaText,
        value: valueText === undefined ? undefined : this.parseValueNode(valueText)
      };
    }

    if (column.v === undefined) {
      return {
        cellName,
        formulaText: null,
        value: undefined
      };
    }

    const valueText = column.v['#text'];

    if (valueText === undefined) {
      return null;
    }

    return {
      cellName,
      formulaText: null,
      value: this.parseValueNode(valueText)
    };
  }

  load() {
    const xml = measureProfileSync<Node>(this.profile, 'worksheet.parseXml', () => {
      return new this.domParserCtor().parseFromString(this.xmlString, 'text/xml');
    });
    const json = measureProfileSync(this.profile, 'worksheet.xmlToJson', function worksheetXmlToJson() {
      return Util.xmlToJson(xml) as WorksheetXmlDocument;
    });
    const rows = normalizeCollection(json.worksheet.sheetData.row);

    incrementProfileCount(this.profile, 'worksheet.rows', rows.length);
    for (const row of rows) {
      const columns = normalizeCollection(row.c);

      incrementProfileCount(this.profile, 'worksheet.cellNodes', columns.length);
      for (const column of columns) {
        this.buildColumn(column);
      }
    }
  }

  parseValueNode(value: string): WorkbookCellValue {
    if (Util.isNumber(value)) {
      return parseFloat(value);
    }

    return value;
  }

  populateSharedFormulas(column: WorksheetColumnNode) {
    const formulaNode = column.f;

    if (formulaNode === undefined) {
      return;
    }

    const shared = formulaNode['@'];
    const formulaText = formulaNode['#text'];

    if (shared === undefined || formulaText === undefined || shared.t !== 'shared') {
      return;
    }

    let reference = shared.ref;
    const sharedIndex = shared.si;

    if (reference === undefined || sharedIndex === undefined) {
      return;
    }

    incrementProfileCount(this.profile, 'worksheet.cells.sharedFormulaAnchors');

    if (reference.split(':').length !== 2) {
      reference = reference + ':' + reference;
    }

    const matches = uniqueValues(formulaText.match(/[A-Z]+[1-9][0-9]*/g) || []);
    const [start, end] = reference.split(':');

    if (start === undefined || end === undefined) {
      return;
    }

    const startRow = parseInt(getRequiredMatch(start, /[0-9]+/gi, 'start row'), 10);
    const startCol = getRequiredMatch(start, /[A-Z]+/gi, 'start column');
    const startColDec = Util.fromBase26(startCol);
    const endRow = parseInt(getRequiredMatch(end, /[0-9]+/gi, 'end row'), 10);
    const endCol = getRequiredMatch(end, /[A-Z]+/gi, 'end column');
    const totalRows = endRow - startRow + 1;
    const totalCols = Util.fromBase26(endCol) - Util.fromBase26(startCol) + 1;
    const sharedFormulas: Record<string, string> = {};

    measureProfileSync(this.profile, 'worksheet.expandSharedFormulas', () => {
      for (let rowOffset = 1; rowOffset <= totalRows; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < totalCols; columnOffset += 1) {
          const currentCell = String(Util.toBase26(startColDec + columnOffset) + (startRow + rowOffset - 1));

          for (const match of matches) {
            const matchRow = parseInt(getRequiredMatch(match, /[0-9]+/gi, 'matched row'), 10);
            const matchCol = getRequiredMatch(match, /[A-Z]+/gi, 'matched column');
            const matchColDec = Util.fromBase26(matchCol);
            const matchFullCol = Util.toBase26(matchColDec + columnOffset);
            const matchFullRow = matchRow + rowOffset - 1;
            const matchCell = String(matchFullCol + matchFullRow);
            const matchCellStart = '$' + matchFullCol + matchFullRow;
            const matchCellStartReplacement = '$' + matchCol + matchFullRow;
            const matchCellEnd = matchFullCol + '$' + matchFullRow;
            const matchCellEndReplacement = matchFullCol + '$' + matchRow;
            const matchSkipStart = '$' + matchCol + matchRow;
            const matchSkipEnd = matchCol + '$' + matchRow;

            const baseFormula = sharedFormulas[currentCell] ?? formulaText;
            sharedFormulas[currentCell] = baseFormula;

            if (baseFormula.includes(matchSkipStart) || baseFormula.includes(matchSkipEnd)) {
              continue;
            }

            let replacementCell = matchCell;

            if (baseFormula.includes(matchCellStart)) {
              replacementCell = matchCellStartReplacement;
            } else if (baseFormula.includes(matchCellEnd)) {
              replacementCell = matchCellEndReplacement;
            }

            sharedFormulas[currentCell] = baseFormula.replace(match, replacementCell);
          }
        }
      }
    });

    incrementProfileCount(this.profile, 'worksheet.cells.sharedFormulaExpanded', Object.keys(sharedFormulas).length);
    for (const [cellName, sharedFormula] of Object.entries(sharedFormulas)) {
      this.addFormulaCell(cellName, sharedFormula);
    }
  }

  shouldPopulateSharedFormulas(column: WorksheetColumnNode) {
    return this.mode === XLSX_READER_MODE.FORMULAS &&
      column.f !== undefined &&
      column.f['@'] !== undefined &&
      column.f['@'].t === 'shared';
  }
}

class XlsxReader {
  domParserCtor: DomParserConstructor;
  mode: ReaderMode;
  profile: XlsxLoadProfiler | undefined;
  sharedStrings: SharedStrings;
  workbookOptions: {
    formulaEvaluator?: FormulaEvaluator;
    functionRegistry?: FormulaFunctionRegistry;
    functions?: Record<string, FormulaFunction>;
  };

  constructor(options?: XlsxReaderOptions) {
    const resolvedOptions: XlsxReaderOptions = options ?? {};

    this.domParserCtor = resolvedOptions.DOMParser ?? getDomParserConstructor();
    this.mode = normalizeReaderMode(resolvedOptions.mode);
    this.profile = resolvedOptions.profile;
    this.sharedStrings = new SharedStrings(this.domParserCtor);
    this.workbookOptions = createWorkbookOptions(resolvedOptions);
  }

  async load(xlsxFile: XlsxInput) {
    return await this.loadInternal(xlsxFile);
  }

  async loadIncremental(xlsxFile: XlsxInput, onSheetLoaded?: XlsxSheetLoadHandler) {
    return await this.loadInternal(xlsxFile, onSheetLoaded);
  }

  private async loadInternal(xlsxFile: XlsxInput, onSheetLoaded?: XlsxSheetLoadHandler) {
    incrementProfileCount(this.profile, 'xlsx.loads');
    const zip = await measureProfileAsync(this.profile, 'xlsx.loadZip', async function loadZipArchive() {
      return await JSZip.loadAsync(xlsxFile);
    });
    const workbook = new Workbook(this.workbookOptions);
    const sharedStringXml = await measureProfileAsync(this.profile, 'zip.read.sharedStringsXml', async function readSharedStringsXml() {
      return await readZipText(zip, 'xl/sharedStrings.xml');
    });

    if (sharedStringXml !== null) {
      this.sharedStrings.set(sharedStringXml, this.profile);
    }

    workbook.beginMutationBatch();

    try {
      const incrementalSheetHandler = onSheetLoaded === undefined
        ? undefined
        : async function handleLoadedSheet(event: XlsxSheetLoadEvent) {
          const hasMoreSheets = event.sheetIndex + 1 < event.sheetCount;

          workbook.endMutationBatch();

          try {
            await onSheetLoaded(event);
          } finally {
            if (hasMoreSheets) {
              workbook.beginMutationBatch();
            }
          }
        };

      return await loadWorkbookFromZip({
        createSheetLoader: (options) => new WorksheetXmlLoader(Object.assign({}, options, {
          mode: this.mode
        })),
        domParserCtor: this.domParserCtor,
        onSheetLoaded: incrementalSheetHandler,
        profile: this.profile,
        sharedStrings: this.sharedStrings,
        workbook: workbook,
        zip: zip
      });
    } finally {
      if (workbook.isMutationBatchActive()) {
        workbook.endMutationBatch();
      }
    }
  }
}

export {
  XLSX_READER_MODE,
  type ReaderMode as XlsxReaderMode,
  type XlsxReaderOptions,
  type XlsxSheetLoadEvent,
  type XlsxSheetLoadHandler,
  XlsxReader
};
