'use strict';

import { DOMParser as NodeDomParser } from '@xmldom/xmldom';
import { Util } from './excellent.util';
import type { Workbook, Worksheet } from './workbook';

type DomParserConstructor = new () => DOMParser;
type ZipFile = {
  async(type: 'string'): Promise<string>;
};
type ZipArchive = {
  file(fileName: string): ZipFile | null | undefined;
};
type XmlTextNode = {
  '#text'?: string;
};
type SharedStringItemNode = {
  t?: XmlTextNode;
};
type SharedStringsXmlDocument = {
  sst: {
    si: SharedStringItemNode | SharedStringItemNode[];
  };
};
type WorkbookFileVersionAttributes = {
  appName: string;
  lastEdited: string;
  lowestEdited: string;
  rupBuild: string;
};
type WorkbookFileVersionNode = {
  '@': WorkbookFileVersionAttributes;
};
type WorkbookSheetAttributes = {
  'r:id': string;
  name: string;
  sheetId?: string;
};
type WorkbookSheetNode = {
  '@'?: WorkbookSheetAttributes;
};
type WorkbookXmlDocument = {
  workbook: {
    fileVersion: WorkbookFileVersionNode;
    sheets: {
      sheet: WorkbookSheetNode | WorkbookSheetNode[];
    };
  };
};
type SheetLoader = {
  load(): Promise<void> | void;
};
type WorkbookLoadOptions = {
  createSheetLoader: (options: {
    domParserCtor: DomParserConstructor;
    profile: XlsxLoadProfiler | undefined;
    sharedStrings: SharedStrings;
    worksheet: Worksheet;
    xmlString: string;
  }) => SheetLoader;
  domParserCtor: DomParserConstructor;
  onSheetLoaded: XlsxSheetLoadHandler | undefined;
  profile: XlsxLoadProfiler | undefined;
  sharedStrings: SharedStrings;
  workbook: Workbook;
  zip: ZipArchive;
};
type XlsxSheetLoadEvent = {
  sheetCount: number;
  sheetIndex: number;
  sheetName: string;
  workbook: Workbook;
  worksheet: Worksheet;
};
type XlsxSheetLoadHandler = (event: XlsxSheetLoadEvent) => Promise<void> | void;
type XlsxLoadProfiler = {
  incrementCount(label: string, amount?: number): void;
  measureAsync<TResult>(label: string, callback: () => Promise<TResult>): Promise<TResult>;
  measureSync<TResult>(label: string, callback: () => TResult): TResult;
};

function getDomParserConstructor(): DomParserConstructor {
  if (typeof globalThis.DOMParser === 'function') {
    return globalThis.DOMParser;
  }

  return NodeDomParser;
}

function readZipText(zip: ZipArchive, fileName: string): Promise<string | null> {
  const file = zip.file(fileName);

  if (file === null || file === undefined) {
    return Promise.resolve(null);
  }

  return file.async('string');
}

function incrementProfileCount(profile: XlsxLoadProfiler | undefined, label: string, amount = 1) {
  if (profile === undefined) {
    return;
  }

  profile.incrementCount(label, amount);
}

function measureProfileAsync<TResult>(
  profile: XlsxLoadProfiler | undefined,
  label: string,
  callback: () => Promise<TResult>
) {
  if (profile === undefined) {
    return callback();
  }

  return profile.measureAsync(label, callback);
}

function measureProfileSync<TResult>(
  profile: XlsxLoadProfiler | undefined,
  label: string,
  callback: () => TResult
) {
  if (profile === undefined) {
    return callback();
  }

  return profile.measureSync(label, callback);
}

function normalizeCollection<T>(value: T | T[] | null | undefined): T[] {
  if (value === null || value === undefined) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function unwrapAttributesNode<TAttributes extends Record<string, string | undefined>>(
  value: { '@'?: TAttributes } | null | undefined
): TAttributes | undefined {
  return value?.['@'];
}

class SharedStrings {
  domParserCtor: DomParserConstructor;
  stringList: string[];

  constructor(domParserCtor: DomParserConstructor = getDomParserConstructor()) {
    this.domParserCtor = domParserCtor;
    this.stringList = [];
  }

  set(xmlData: string, profile?: XlsxLoadProfiler) {
    incrementProfileCount(profile, 'xlsx.sharedStringTables');
    const xml = measureProfileSync<Node>(profile, 'sharedStrings.parseXml', () => {
      return new this.domParserCtor().parseFromString(xmlData, 'text/xml');
    });
    const json = measureProfileSync(profile, 'sharedStrings.xmlToJson', function sharedStringsXmlToJson() {
      return Util.xmlToJson(xml) as SharedStringsXmlDocument;
    });
    const sharedItems = normalizeCollection(json.sst.si);

    incrementProfileCount(profile, 'sharedStrings.items', sharedItems.length);
    this.stringList = measureProfileSync(profile, 'sharedStrings.materialize', function materializeSharedStrings() {
      return sharedItems.map(function mapSharedString(data) {
        if (data.t !== undefined) {
          return data.t['#text'] ?? '';
        }

        return '';
      });
    });
  }

  get(index: number): string | undefined {
    return this.stringList[index];
  }
}

async function loadWorkbookFromZip(options: WorkbookLoadOptions) {
  const {
    createSheetLoader,
    domParserCtor,
    onSheetLoaded,
    profile,
    sharedStrings,
    workbook,
    zip
  } = options;

  incrementProfileCount(profile, 'xlsx.workbooks');
  const workbookXml = await measureProfileAsync(profile, 'zip.read.workbookXml', async function readWorkbookXml() {
    return await readZipText(zip, 'xl/workbook.xml');
  });
  if (workbookXml === null) {
    throw new Error('Missing xl/workbook.xml in XLSX file.');
  }

  const xml = measureProfileSync<Node>(profile, 'workbook.parseXml', function parseWorkbookXml() {
    return new domParserCtor().parseFromString(workbookXml, 'text/xml');
  });
  const json = measureProfileSync(profile, 'workbook.xmlToJson', function workbookXmlToJson() {
    return Util.xmlToJson(xml) as WorkbookXmlDocument;
  });
  const workbookJson = json.workbook;
  const attrs = workbookJson.fileVersion['@'];
  const sheetList = normalizeCollection(workbookJson.sheets.sheet);

  incrementProfileCount(profile, 'workbook.sheetsDeclared', sheetList.length);
  measureProfileSync(profile, 'workbook.applyMetadata', function applyWorkbookMetadata() {
    workbook
      .setType(attrs.appName)
      .setFileVersion(attrs.lastEdited + '.' + attrs.lowestEdited + '.' + attrs.rupBuild);
  });

  let loadedSheetCount = 0;
  for (const rawSheet of sheetList) {
    const sheet = unwrapAttributesNode(rawSheet as Record<string, unknown>) as {
      name: string;
      'r:id': string;
    } | undefined;

    if (sheet === undefined) {
      continue;
    }

    const sheetId = sheet['r:id'].replace('rId', '');
    const sheetXml = await measureProfileAsync(profile, 'zip.read.worksheetXml', async function readWorksheetXml() {
      return await readZipText(zip, 'xl/worksheets/sheet' + sheetId + '.xml');
    });

    if (sheetXml === null) {
      continue;
    }

    incrementProfileCount(profile, 'workbook.sheetsLoaded');
    const worksheet = measureProfileSync(profile, 'workbook.createSheet', function createWorksheet() {
      return workbook.addSheet(sheet.name);
    });
    const sheetLoader = createSheetLoader({
      domParserCtor,
      profile,
      sharedStrings,
      worksheet: worksheet,
      xmlString: sheetXml
    });

    await sheetLoader.load();

    if (onSheetLoaded !== undefined) {
      await onSheetLoaded({
        sheetCount: sheetList.length,
        sheetIndex: loadedSheetCount,
        sheetName: sheet.name,
        workbook: workbook,
        worksheet: worksheet
      });
    }

    loadedSheetCount += 1;
  }

  return workbook;
}

  export {
  SharedStrings,
  getDomParserConstructor,
  incrementProfileCount,
  loadWorkbookFromZip,
  measureProfileAsync,
  measureProfileSync,
  normalizeCollection,
  readZipText,
  type SharedStringsXmlDocument,
  type WorkbookFileVersionAttributes,
  type WorkbookSheetAttributes,
  type WorkbookSheetNode,
  type WorkbookXmlDocument,
  type XlsxLoadProfiler,
  type XlsxSheetLoadEvent,
  type XlsxSheetLoadHandler,
  type XmlTextNode,
  unwrapAttributesNode
};
