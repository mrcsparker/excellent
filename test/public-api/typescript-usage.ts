import {
  ExcelError,
  isExcelError,
  FormulaFunctionRegistry,
  FormulaEvaluator,
  Workbook,
  WorkbookLoader,
  XLSX_READER_MODE,
  XlsxReader,
  type CompiledFormula,
  type FormulaAstNode,
  type LoaderWorkbook,
  type SerializedSheet,
  type SerializedWorkbookCell,
  type XlsxSheetLoadEvent,
  type WorkbookHandle,
  type Cell
} from '../..';

async function parseWorkbook(bytes: Uint8Array): Promise<WorkbookHandle> {
  const xlsx = new XlsxReader({
    mode: XLSX_READER_MODE.FORMULAS
  });
  const incrementalSnapshots: string[] = [];
  const onSheetLoaded = async function(event: XlsxSheetLoadEvent): Promise<void> {
    incrementalSnapshots.push(event.sheetName + ':' + String(event.sheetIndex));
    await Promise.resolve();
  };
  const incrementallyParsed = await xlsx.loadIncremental(bytes, onSheetLoaded);
  const parsed = await xlsx.load(bytes);
  const sheet = parsed.workbook['Sheet1'];

  void incrementalSnapshots;
  void incrementallyParsed;

  if (sheet === undefined) {
    return parsed;
  }

  const firstCell = sheet['A1'];
  const firstCellModel = sheet.getCell('A1');

  void firstCell;
  void sheet.getCellValue('A1');
  void firstCellModel;

  return parsed;
}

function roundTripLoader(workbook: WorkbookHandle): LoaderWorkbook {
  const loader = new WorkbookLoader();
  const serialized = loader.serialize(workbook);
  const restored = loader.deserialize(serialized);

  void restored;

  return serialized;
}

function buildTypedFormulaArtifacts(): CompiledFormula {
  const ast: FormulaAstNode = {
    arguments: [
      { type: 'Literal', value: 1 },
      { type: 'Literal', value: 2 }
    ],
    name: 'SUM',
    type: 'FormulaCallExpression'
  };
  const serializedCell: SerializedWorkbookCell = '[function]SUM(1,2)';
  const serializedSheet: SerializedSheet = [[serializedCell, 3]];

  void serializedSheet;

  return {
    ast,
    expression: 'Formula.SUM(1,2)'
  };
}

function configureWorkbook(): Workbook {
  const registry = new FormulaFunctionRegistry().register('DOUBLE', (value: unknown) => Number(value) * 2);
  const workbook = new Workbook({
    formulaEvaluator: new FormulaEvaluator(),
    functionRegistry: registry
  });

  workbook.registerFunction('TRIPLE', (value: unknown) => Number(value) * 3);
  workbook.createSheet('Sheet1');
  workbook.setCellValue('Sheet1', 'A1', 3);
  workbook.setCellFormula('Sheet1', 'A2', 'this.A1+1');
  void workbook.hasFunction('DOUBLE');
  void workbook.getFunctionNames();

  return workbook;
}

function inspectCell(cell: Cell | undefined): string | null {
  if (cell === undefined) {
    return null;
  }

  return cell.address;
}

function readExcelError(value: unknown): string | null {
  if (isExcelError(value)) {
    return value.code;
  }

  return ExcelError.REF.code;
}

async function main(): Promise<void> {
  const workbook = await parseWorkbook(new Uint8Array());
  const serialized = roundTripLoader(workbook);
  const compiledFormula = buildTypedFormulaArtifacts();
  const configuredWorkbook = configureWorkbook();
  const valuesOnlyReader = new XlsxReader({
    mode: XLSX_READER_MODE.VALUES_ONLY
  });
  const inspectedCell = inspectCell(configuredWorkbook.getCell('Sheet1', 'A2'));
  const maybeError = readExcelError(configuredWorkbook.traceCell('Missing', 'A1').value);

  void serialized;
  void configuredWorkbook;
  void compiledFormula;
  void valuesOnlyReader;
  void inspectedCell;
  void maybeError;
}

void main();
