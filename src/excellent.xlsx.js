var Excellent = Excellent || {};

if (typeof exports !== 'undefined') {
  var Excellent = require('./excellent.workbook.js').Excellent;
  var JSZip = require('jszip');
  var FormulaParser = require('./excellent.parser.js');
}

if (typeof DOMParser === 'undefined') {
  var DOMParser = require('xmldom').DOMParser;
}

Array.prototype.contains = function(v) {
  'use strict';

  var i;

  for (i = 0; i < this.length; i += 1) {
    if (this[i] === v) {
      return true;
    }
  }
  return false;
};

Array.prototype.unique = function() {
  'use strict';

  var i;
  var arr = [];

  for (i = 0; i < this.length; i += 1) {
    if (!arr.contains(this[i])) {
      arr.push(this[i]);
    }
  }
  return arr;
};

Excellent.Xlsx = function() {
  'use strict';

  var self = {};
  var workbook = new Excellent.Workbook();

  var XlsxStrings = (function() {
    var stringList = [];

    function getText(data) {
      var retVal = '';

      if (data.t !== undefined) {
        retVal = data.t['#text'];
      } else {
        // TODO: Handle nested data
        retVal = '';
      }

      return retVal;
    }

    function populateList(xmlString) {
      var xml;
      var json;
      var tmpList;
      var si;

      xml = new DOMParser().parseFromString(xmlString, 'text/xml');
      json = Excellent.Util.xmlToJson(xml);

      tmpList = [];

      si = json.sst.si;
      Excellent.Util.each(si, function(data) {
        tmpList.push(getText(data));
      });

      stringList = tmpList;
    }

    function getListItem(id) {
      return stringList[id];
    }

    return {
      set: function(xmlData) {
        populateList(xmlData);
      },
      get: function(id) {
        return getListItem(id);
      }
    };
  }());

  function XlsxSheet(xmlString) {
    var self = {};
    var sheetData;
    var xml;
    var json;

    xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    json = Excellent.Util.xmlToJson(xml);
    sheetData = json.worksheet.sheetData;

    function matchesCell(str, cell) {
      return str.indexOf(cell) > -1;
    }

    function buildFunction(formula) {
      var output = '';

      // We are going to cheat a bit and remove
      // some simple excelisms

      // remove excess spaces
      formula = formula.trim();

      // old excel formula sometimes are written as
      // +formula
      // rather than
      // =formula
      // remove it!
      if (formula && formula.substring(0, 1) === '+') {
        formula = formula.substring(1);
      }

      output = FormulaParser.parse(formula);
      output = output.trim();
      return output;
    }

    // Very, very ugly.  Still working this out.
    // Almost there, then code cleanup.
    // Wow. This is ugly.  Really, almost there on this one, then
    // I will clean up the code.
    function populateSharedFormulas(c) {
      var shared,
        ref,
        si,

        matches,
        func;

      shared = c.f['@'];

      func = c.f['#text'];

      if (shared === undefined || func === undefined || shared.t !== 'shared') {
        return;
      }

      ref = shared.ref;

      si = shared.si;

      if (ref === undefined || si === undefined) {
        return;
      }

      // We are going to treat all of the refs as ranges
      if (ref.split(':').length !== 2) {
        ref = ref + ':' + ref;
      }

      // Pull out all of the variables.
      matches = func.match(/[A-Z]+[1-9][0-9]*/g).unique();

      var refData = ref.split(':'),
        start = refData[0],
        end = refData[1],

        startRow = parseInt(start.match(/[0-9]+/gi)[0], 10),
        startCol = start.match(/[A-Z]+/gi)[0],
        startColDec = Excellent.Util.fromBase26(startCol),

        endRow = parseInt(end.match(/[0-9]+/gi)[0], 10),
        endCol = end.match(/[A-Z]+/gi)[0],

        // Total rows and cols
        totalRows = endRow - startRow + 1,
        totalCols = Excellent.Util.fromBase26(endCol) - Excellent.Util.fromBase26(startCol) + 1,

        // Loop vars
        matchIter,
        curRow,
        curCol,
        curCell = '',
        match,
        matchRow,
        matchCol,
        matchColDec,
        matchCell = '',
        sharedFormulas = {},

        matchFullCol,
        matchFullRow,

        matchCellStart,
        matchCellStartReplace,

        matchCellEnd,
        matchCellEndReplace,

        matchSkipStart,
        matchSkipEnd;

      for (curRow = 1; curRow <= totalRows; curRow += 1) {

        for (curCol = 0; curCol < totalCols; curCol += 1) {

          // Get the current cell id
          curCell = String(Excellent.Util.toBase26(startColDec + curCol) + (startRow + curRow - 1));

          for (matchIter = 0; matchIter < matches.length; matchIter += 1) {

            match = matches[matchIter];
            matchRow = parseInt(match.match(/[0-9]+/gi)[0], 10);
            matchCol = match.match(/[A-Z]+/gi)[0];
            matchColDec = Excellent.Util.fromBase26(matchCol);

            matchFullCol = Excellent.Util.toBase26(matchColDec + curCol);
            matchFullRow = (matchRow + curRow - 1);

            matchCell = String(matchFullCol + matchFullRow);
            matchCellStart = String('$' + matchFullCol + matchFullRow);
            matchCellStartReplace = String('$' + matchCol + matchFullRow);

            matchCellEnd = matchFullCol + '$' + matchFullRow;
            matchCellEndReplace = matchFullCol + '$' + matchRow;

            matchSkipStart = String('$' + matchCol + matchRow);
            matchSkipEnd = matchCol + '$' + matchRow;

            if (sharedFormulas[curCell] === undefined) {
              sharedFormulas[curCell] = func;
            }

            if (matchesCell(sharedFormulas[curCell], matchSkipStart)) {
              continue;
            }

            if (matchesCell(sharedFormulas[curCell], matchSkipEnd)) {
              continue;
            }

            if (matchesCell(sharedFormulas[curCell], matchCellStart)) {
              matchCell = matchCellStartReplace;
            } else if (matchesCell(sharedFormulas[curCell], matchCellEnd)) {
              matchCell = matchCellEndReplace;
            }

            sharedFormulas[curCell] = sharedFormulas[curCell].replace(match, matchCell);
          }

        }
        curCol = 0;
      }

      Excellent.Util.each(sharedFormulas, function(columnText, columnValue) {
        workbook.addCellFunc(columnValue, buildFunction(columnText.replace(/\$/g, '')));
      });

      return sharedFormulas;
    }

    // This code is pretty ugly.  Going to clean it up as soon as
    // I get the various Excel types worked out.
    function buildColumn(c) {
      var columnId;
      var columnValue;
      var columnType;

      if (c === undefined) {
        return;
      }

      columnId = c['@'];
      if (columnId === undefined) {
        return;
      }

      columnValue = columnId.r; // Column ID: A2, A3, etc
      columnType = columnId.t; // shared text?

      if (c.f !== undefined && c.f['@'] !== undefined && c.f['@'].t === 'shared') {
        populateSharedFormulas(c);
      } else if (columnType === 's') { // shared string
        workbook.addCellVal(columnValue, XlsxStrings.get(parseInt(c.v['#text'], 10)));
      } else if (c.f !== undefined && c.f['#text'] !== undefined) { // function
        workbook.addCellFunc(columnValue, buildFunction(c.f['#text'].replace(/\$/g, '')));
      } else if (c.v !== undefined) { // value
        if (Excellent.Util.isNumber(c.v['#text'])) {
          workbook.addCellVal(columnValue, parseFloat(c.v['#text']));
        } else {
          workbook.addCellVal(columnValue, c.v['#text']);
        }
      }
    }

    function buildColumns(columns) {
      Excellent.Util.each(columns, function(c) {
        buildColumn(c);
      });
    }

    self.load = function() {
      Excellent.Util.each(sheetData, function(data) {
        Excellent.Util.each(data, function(column) {
          if (column.c === undefined || column.c.length === undefined) {
            buildColumns([column.c]);
          } else {
            buildColumns(column.c);
          }
        });
      });
    };

    return self;
  }

  function XlsxWorkbook(xmlString, zipLib) {
    var self = {};
    var xml;
    var zip;
    var json;

    xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    zip = zipLib;
    json = Excellent.Util.xmlToJson(xml).workbook;

    function buildFileVersion() {
      var attrs;
      var fileVersion;

      attrs = json.fileVersion['@'];
      fileVersion =
        attrs.lastEdited + '.' +
        attrs.lowestEdited + '.' +
        attrs.rupBuild;

      workbook.setType(attrs.appName).setFileVersion(fileVersion);
    }

    function buildSheet(s) {
      var sheetId;
      var sheetName;
      var xlsxSheet;

      sheetId = s['r:id'].replace('rId', '');

      sheetName = 'xl/worksheets/sheet' + sheetId + '.xml';
      xlsxSheet = new XlsxSheet(zip.file(sheetName).asText());
      xlsxSheet.load();
    }

    function buildSheets() {
      var sheets = json.sheets.sheet;

      Excellent.Util.each(sheets, function(s) {

        if (s === undefined) {
          return;
        }

        if (s['@'] !== undefined) {
          s = s['@'];
        }

        workbook.addSheet(s.name);
        buildSheet(s);
      });
    }

    self.load = function() {
      buildFileVersion();
      buildSheets();
    };

    return self;
  }

  self.load = function(xlsxFile) {

    function extractFiles(zipFile) {
      var zip;
      var stringZip;
      var zipWorkbook;
      var xlsxWorkbook;

      zip = new JSZip(zipFile);

      stringZip = zip.file('xl/sharedStrings.xml');
      if (stringZip !== null) {
        XlsxStrings.set(stringZip.asText());
      }

      zipWorkbook = zip.file('xl/workbook.xml').asText();
      xlsxWorkbook = new XlsxWorkbook(zipWorkbook, zip);
      xlsxWorkbook.load();

      return workbook;
    }
    return extractFiles(xlsxFile);
  };

  return self;
};

if (typeof exports !== 'undefined') {
  exports.Excellent = Excellent;
}
