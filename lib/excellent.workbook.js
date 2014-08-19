var Excellent = Excellent || {};

if (typeof exports !== 'undefined') {
  var Excellent = require('./excellent.util.js').Excellent,
    Formula = require('formulajs');
}

var Formula = Formula || {};

Formula.VLOOKUP = function(needle, index, exactmatch) {
  'use strict';

  var i, row;

  index = index || 0;
  exactmatch = exactmatch || false;
  for (i = 0; i < this.length; i += 1) {
    row = this[i];

    if ((exactmatch && row[0] === needle) || row[0].toLowerCase().indexOf(needle.toLowerCase()) !== -1) {
      return (index < row.length ? row[index] : row);
    }
  }
  return null;
};

Formula.MATCH = function(needle, searchArray, exactmatch) {
  'use strict';

  var i, item;

  exactmatch = exactmatch || false;
  for (i = 0; i < searchArray.length; i += 1) {
    item = searchArray[i];

    if ((exactmatch && item === needle) || String(item).toLowerCase() === String(needle).toLowerCase()) {
      return item;
    }
  }
  return null;
};

Formula.INDEX = function(searchArray, index) {
  'use strict';

  return searchArray[index];
};

Excellent.Workbook = function() {
  'use strict';

  var self = {};

  self.workbook = {};
  self.currentSheet = {};
  self.type = "";
  self.fileVersion = "";

  self.addSheet = function(sheetName) {
    if (!self.workbook.hasOwnProperty(sheetName)) {
      self.workbook[sheetName] = {};
    }

    self.currentSheet = self.workbook[sheetName];
    self.currentSheet.variables = [];
    self.currentSheet.functions = [];
    self.currentSheet.rows = [];
    return this;
  };

  self.setType = function(type) {
    self.type = type;
    return this;
  };

  self.setFileVersion = function(fileVersion) {
    self.fileVersion = fileVersion;
    return this;
  };

  self.addCellVal = function(cellName, cellVal) {
    var sheet, row, col, val;

    self.currentSheet['_' + cellName] = cellVal;
    self.currentSheet.variables.push(cellName);

    sheet = self.currentSheet;

    if (sheet[cellName]) {
      return;
    }

    Object.defineProperty(sheet, cellName, {
      get: function() {
        return sheet['_' + cellName];
      },
      set: function(val) {
        sheet['_' + cellName] = val;
      }
    });

    row = Excellent.Util.getRowFromCell(cellName);
    col = Excellent.Util.getColFromCell(cellName);

    if (self.currentSheet.rows[row] === undefined) {
      self.currentSheet.rows[row] = [];
    }

    self.currentSheet.rows[row][col] = {
      index: cellName
    };

    val = sheet['_' + cellName];

    self.currentSheet.rows[row][col] = {
      index: cellName,
      value: val
    };

    return this;
  };

  self.addCellFunc = function(cellName, cellVal) {
    var sheet, row, col, func;

    self.currentSheet['_' + cellName] = cellVal;
    self.currentSheet.functions.push(cellName);

    sheet = self.currentSheet;

    if (sheet[cellName]) {
      return;
    }

    Object.defineProperty(sheet, cellName, {
      get: function() {
        return eval(sheet['_' + cellName]);
      },
      set: function(val) {
        sheet['_' + cellName] = val;
      }
    });

    row = Excellent.Util.getRowFromCell(cellName);
    col = Excellent.Util.getColFromCell(cellName);

    if (self.currentSheet.rows[row] === undefined) {
      self.currentSheet.rows[row] = [];
    }

    func = sheet['_' + cellName];

    self.currentSheet.rows[row][col] = {
      index: cellName,
      value: func
    };

    return this;
  };

  self.getRawWorkbook = function() {
    return self.workbook;
  };

  self.zeroOutNullRows = function() {
    var sheet, cell;

    for (sheet in self.workbook) {
      if (self.workbook.hasOwnProperty(sheet)) {
        self.workbook[sheet].rows.forEach(function(line, lineIndex) {
          line.forEach(function(c, rowIndex) {
            if (c === null) {
              self.workbook[sheet].rows[lineIndex][rowIndex] = 0;
            }
          });
        });
      }
    }

    for (sheet in self.workbook) {
      if (self.workbook.hasOwnProperty(sheet)) {
        for (cell in self.workbook[sheet]) {
          if (self.workbook[sheet].hasOwnProperty(cell)) {
            if (self.workbook[sheet][cell] === null) {
              self.workbook[sheet][cell] = 0;
            }
          }
        }
      }
    }
  };

  return self;
};

if (typeof exports !== 'undefined') {
  exports.Excellent = Excellent;
}
