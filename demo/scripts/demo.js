function DemoCtrl($scope) {

  'use strict';

  $scope.workbook = {};

  $scope.Object = window.Object;

  $scope.hasSavedFullWorkbook = function() {
    return localStorage.getItem('workbook');
  };

  $scope.saveFullWorkbook = function() {

    var excellentLoader = new Excellent.Loader(),
      json = excellentLoader.unload($scope.fullWorkbook);

    localStorage.setItem('workbook', JSON.stringify(json));
  };

  $scope.loadFullWorkbook = function() {

    var excellentLoader = new Excellent.Loader(),
      workbook = excellentLoader.load(JSON.parse(localStorage.getItem('workbook')));

    console.log(workbook);

    $scope.fullWorkbook = workbook;
    $scope.workbook = workbook.workbook;

  };

  $scope.isNumber = function(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
  };

  $scope.updateCell = function(workbook, sheet, cell) {
    if ($scope.isNumber(cell.value)) {
      workbook[sheet][cell.index] = parseFloat(cell.value, 10);
    } else {
      workbook[sheet][cell.index] = cell.value;
    }
  };

  function numToChar(number) {
    var numeric = (number - 1) % 26;
    var letter = chr(65 + numeric);
    var number2 = parseInt((number - 1) / 26);
    if (number2 > 0) {
      return numToChar(number2) + letter;
    } else {
      return letter;
    }
  }

  function chr(codePt) {
    if (codePt > 0xFFFF) {
      codePt -= 0x10000;
      return String.fromCharCode(0xD800 + (codePt >> 10), 0xDC00 + (codePt & 0x3FF));
    }
    return String.fromCharCode(codePt);
  }

  $scope.sheetColumns = function(worksheet) {
    var highest = worksheet.rows[0],
      columns = [],
      i;

    worksheet.rows.forEach(function(row) {
      if (!row || !highest) {
        return;
      }
      if (highest.length < row.length) {
        highest = row;
      }
    });

    if (!highest) {
      return;
    }

    for (i = 0; i < highest.length; i += 1) {
      columns.push(numToChar(i + 1));
    }

    return columns;

  };

  $scope.upload = function(uploadFiles) {
    var uploadFile = uploadFiles.files[0],
      reader = new FileReader();

    reader.onload = (function(file) {
      return function(e) {
        try {
          var excellent = new Excellent.Xlsx(),
            parsed = excellent.load(e.target.result);
          $scope.$apply(function() {
            $scope.fullWorkbook = parsed;
            $scope.workbook = parsed.workbook;
            console.log(parsed);
          });
        } catch (e) {
          alert('There was a problem parsing this file.  Excellent is still under development. \n\nMessage:\n\n' + e.message);
        }
      };
    })(uploadFile);

    reader.readAsArrayBuffer(uploadFile);
  };
}
