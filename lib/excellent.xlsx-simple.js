var Excellent = Excellent || {};

if (typeof exports !== 'undefined') {
  var Excellent = require('./excellent.workbook.js').Excellent,
    JSZip = require('jszip');
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

  var i, arr = [];
  for (i = 0; i < this.length; i += 1) {
    if (!arr.contains(this[i])) {
      arr.push(this[i]);
    }
  }
  return arr;
};

Excellent.XlsxSimple = function() {
  'use strict';

  var self = {},
    workbook = new Excellent.Workbook();

  var XlsxStrings = (function() {
    var stringList = [];

    function getText(data) {
      var retVal = '';

      if (data.t !== undefined) {
        retVal = data.t['#text'];
      } else {
        // TODO: Handle nested data
      }

      return retVal;
    }

    function populateList(xmlString) {
      var xml, json, tmpList, si;

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

    var sheetData, xml, json;

    xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    json = Excellent.Util.xmlToJson(xml);
    sheetData = json.worksheet.sheetData;

    // This code is pretty ugly.  Going to clean it up as soon as
    // I get the various Excel types worked out.
    function buildColumn(c) {
      var columnId,
        columnValue,
        columnType;

      if (c === undefined) {
        return;
      }

      columnId = c['@'];
      if (columnId === undefined) {
        return;
      }

      columnValue = columnId.r; // Column ID: A2, A3, etc
      columnType = columnId.t; // shared text?

      if (columnType === 's') { // shared string
        workbook.addCellVal(columnValue, XlsxStrings.get(parseInt(c.v['#text'], 10)));
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
    var self = {},

      xml,
      zip,
      json;

    xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    zip = zipLib;
    json = Excellent.Util.xmlToJson(xml).workbook;

    function buildFileVersion() {
      var attrs, fileVersion;

      attrs = json.fileVersion['@'];
      fileVersion = attrs.lastEdited + '.' + attrs.lowestEdited + '.' + attrs.rupBuild;

      workbook.setType(attrs.appName).setFileVersion(fileVersion);
    }

    function buildSheet(s) {
      var sheetId, sheetName, xlsxSheet;

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
      var zip,
        stringZip,
        xlsxWorkbook;

      zip = new JSZip(zipFile);

      stringZip = zip.file('xl/sharedStrings.xml');
      if (stringZip !== null) {
        XlsxStrings.set(stringZip.asText());
      }

      xlsxWorkbook = new XlsxWorkbook(zip.file('xl/workbook.xml').asText(), zip);
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
