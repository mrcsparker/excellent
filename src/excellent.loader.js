// var workbook = { "Sheet1":
//    [ [ "Name", "Age", "Days", "Holidays", "Left" ],
//        [ "John", 45, 134, 7, "[function] this.B2 - this.C2;" ],
//        [ "Abe", 81, 23, 1, "[function] this.B3 - this.C3;" ],
//        [ "Mike", 92, 12, 2, "[function] this.B4 - this.C4;" ],
//        [ "Tonto", 54, 132, 3, "[function] this.B5 - this.C5;" ] ] };

var Excellent = Excellent || {};

if (typeof exports !== 'undefined') {
  var Excellent = require('./excellent.workbook.js').Excellent;
}

Excellent.Loader = function() {
  'use strict';

  var self = {};

  self.load = function(json) {

    var workbook = new Excellent.Workbook(),
      currentSheet = {};

    function isFunction(n) {
      return n !== null && n !== undefined && !Excellent.Util.isNumber(n) && (n.lastIndexOf('[function]', 0) === 0);
    }

    function getVal(valData) {
      var retVal = valData;

      if (isFunction(valData)) {
        retVal = valData.replace('[function]', '');
      }
      return retVal;
    }

    function addRow(rowData, rowId) {
      var cellName, cellVal, i;

      if (rowData === null || rowData === undefined) {
        return;
      }

      for (i = 0; i < rowData.length; i += 1) {

        cellName = String(Excellent.Util.toBase26(i) + (rowId + 1));
        cellVal = getVal(rowData[i]);

        if (isFunction(rowData[i])) {
          currentSheet.addCellFunc(cellName, cellVal);
        } else {
          currentSheet.addCellVal(cellName, cellVal);
        }
      }
    }

    function addCells(sheetData) {
      var i;

      for (i = 0; i < sheetData.length; i += 1) {
        addRow(sheetData[i], i);
      }
    }

    function addSheets() {
      var sheetName;

      for (sheetName in json) {
        if (json.hasOwnProperty(sheetName)) {
          currentSheet = workbook.addSheet(sheetName);
          addCells(json[sheetName]);
        }
      }
    }

    addSheets();

    return workbook;

  };

  self.unload = function(excellentObject) {
    var json = {},
      workbook = excellentObject.workbook;

    function loadJson() {
      var sheet, currentSheet, variable, row, col, i;

      for (sheet in workbook) {
        if (workbook.hasOwnProperty(sheet)) {
          json[sheet] = [];

          currentSheet = workbook[sheet];

          for (i = 0; i < currentSheet.variables.length; i += 1) {
            variable = currentSheet.variables[i];
            row = Excellent.Util.getRowFromCell(variable);
            col = Excellent.Util.getColFromCell(variable);

            if (json[sheet][row] === undefined) {
              json[sheet][row] = [];
            }

            json[sheet][row][col] = currentSheet['_' + variable];
          }

          for (i = 0; i < currentSheet.functions.length; i += 1) {
            variable = currentSheet.functions[i];
            row = Excellent.Util.getRowFromCell(variable);
            col = Excellent.Util.getColFromCell(variable);

            if (json[sheet][row] === undefined) {
              json[sheet][row] = [];
            }
            json[sheet][row][col] = '[function]' + currentSheet['_' + variable];
          }
        }
      }

      for (sheet in json) {
        if (json.hasOwnProperty(sheet)) {
          for (i = 0; i < json[sheet].length; i += 1) {
            if (json[sheet][i] === undefined) {
              json[sheet][i] = [];
            }
          }
        }
      }
    }

    loadJson();

    return json;
  };

  return self;
};

if (typeof exports !== 'undefined') {
  exports.Excellent = Excellent;
}
