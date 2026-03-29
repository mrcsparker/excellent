'use strict';

import type { WorkbookCellValue } from './formula';
import { Util } from './excellent.util';
import { Workbook } from './workbook';
import type { WorkbookHandle, WorkbookOptions, WorkbookSheets } from './workbook';

type SerializedFormulaCell = `[function]${string}`;
type SerializedWorkbookCell = WorkbookCellValue;
type SerializedRow = Array<SerializedWorkbookCell | undefined>;
type SerializedSheet = Array<SerializedRow | undefined>;
type LoaderWorkbook = Record<string, SerializedSheet | null | undefined>;
type WorkbookLike = WorkbookHandle | WorkbookSheets;

function createWorkbookOptions(options?: WorkbookOptions): WorkbookOptions {
  const workbookOptions: WorkbookOptions = {};

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

function isFormulaCell(value: SerializedWorkbookCell | undefined): value is SerializedFormulaCell {
  return value !== null &&
    value !== undefined &&
    !Util.isNumber(value) &&
    typeof value === 'string' &&
    value.lastIndexOf('[function]', 0) === 0;
}

function getSerializedCellValue(value: SerializedWorkbookCell | undefined): WorkbookCellValue | string | undefined {
  if (!isFormulaCell(value)) {
    return value;
  }

  return value.replace('[function]', '');
}

function createSerializedFormulaCell(formulaExpression: string): SerializedFormulaCell {
  return ('[function]' + formulaExpression) as SerializedFormulaCell;
}

function isWorkbookHandle(workbookLike: WorkbookLike): workbookLike is WorkbookHandle {
  return Object.prototype.hasOwnProperty.call(workbookLike, 'workbook');
}

function getWorkbookHandle(workbookLike: WorkbookLike): WorkbookSheets {
  if (isWorkbookHandle(workbookLike)) {
    return workbookLike.workbook;
  }

  return workbookLike;
}

class WorkbookLoader {
  workbookOptions: WorkbookOptions;

  constructor(options?: WorkbookOptions) {
    this.workbookOptions = createWorkbookOptions(options);
  }

  deserialize(json: LoaderWorkbook): Workbook {
    const workbook = new Workbook(this.workbookOptions);

    workbook.beginMutationBatch();

    try {
      for (const [sheetName, sheetRows] of Object.entries(json)) {
        const sheet = workbook.createSheet(sheetName);

        if (sheetRows === null || sheetRows === undefined) {
          continue;
        }

        for (let rowIndex = 0; rowIndex < sheetRows.length; rowIndex += 1) {
          const row = sheetRows[rowIndex];

          if (row === null || row === undefined) {
            continue;
          }

          for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
            const cellName = String(Util.toBase26(columnIndex) + (rowIndex + 1));
            const cellValue = getSerializedCellValue(row[columnIndex]);

            if (isFormulaCell(row[columnIndex])) {
              sheet.setCellFormula(cellName, cellValue as string);
            } else {
              sheet.setCellValue(cellName, cellValue);
            }
          }
        }
      }
    } finally {
      workbook.endMutationBatch();
    }

    return workbook;
  }

  load(json: LoaderWorkbook): Workbook {
    return this.deserialize(json);
  }

  serialize(workbookLike: WorkbookLike): LoaderWorkbook {
    const json: LoaderWorkbook = {};
    const workbook = getWorkbookHandle(workbookLike);

    for (const [sheetName, sheet] of Object.entries(workbook)) {
      json[sheetName] = [];

      for (const variable of sheet.variables) {
        const rowIndex = Util.getRowFromCell(variable);
        const columnIndex = Util.getColFromCell(variable);

        if (json[sheetName][rowIndex] === undefined) {
          json[sheetName][rowIndex] = [];
        }

        json[sheetName][rowIndex][columnIndex] = sheet.getRawCellValue(variable);
      }

      for (const formulaCell of sheet.functions) {
        const rowIndex = Util.getRowFromCell(formulaCell);
        const columnIndex = Util.getColFromCell(formulaCell);
        const formulaSource = sheet.getFormulaSource(formulaCell);

        if (json[sheetName][rowIndex] === undefined) {
          json[sheetName][rowIndex] = [];
        }

        if (formulaSource === undefined) {
          throw new Error('Formula cell is missing source during serialization: ' + sheetName + '!' + formulaCell);
        }

        json[sheetName][rowIndex][columnIndex] = createSerializedFormulaCell(formulaSource);
      }
    }

    for (const sheetRows of Object.values(json)) {
      if (sheetRows === null || sheetRows === undefined) {
        continue;
      }

      for (let rowIndex = 0; rowIndex < sheetRows.length; rowIndex += 1) {
        if (sheetRows[rowIndex] === undefined) {
          sheetRows[rowIndex] = [];
        }
      }
    }

    return json;
  }
}

export {
  type SerializedFormulaCell,
  type SerializedRow,
  type SerializedSheet,
  type SerializedWorkbookCell,
  type LoaderWorkbook,
  WorkbookLoader
};
