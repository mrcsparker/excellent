'use strict';

var workbook;

$(function() {
  if (!window.FileReader || !window.ArrayBuffer) {
    alert("You will need a recent browser to use this demo :(");
    return;
  }

  var $result = $("#result");
  $("#file").on("change", function(evt) {
    // remove content
    $result.html("");

    // see http://www.html5rocks.com/en/tutorials/file/dndfiles/

    var files = evt.target.files;
    for (var i = 0, f; f = files[i]; i++) {

      if (f.type !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
        $result.append("<div class='warning'>" + f.name + " isn't an XLSX file</div>");
      }
      var reader = new FileReader();

      // Closure to capture the file information.
      reader.onload = (function(theFile) {
        return function(e) {
          var $title = $("<h3>", {
            text: theFile.name
          });
          $result.append($title);
          var $ul = $("<ul>");
          try {

            var dateBefore = new Date();
            // read the content of the file with Excellent
            var excellent = new Excellent.Xlsx();
            parsed = excellent.load(e.target.result);
            var dateAfter = new Date();

            $title.append($("<span>", {
              text: " (parsed in " + (dateAfter - dateBefore) + "ms)"
            }));

            $.each(parsed.workbook, function(index, sheet) {
              $ul.append("<li><strong>" + index + "</strong></li>");
              $.each(sheet.rows, function(index, row) {
                if (row === null || row === undefined) {
                  return;
                }
                $ul.append("<li><hr /></li>");
                $ul.append("<li>");
                $.each(row, function(index, cell) {
                  if (cell === undefined) {
                    return;
                  }
                  $ul.append("&nbsp; " + cell + " : " + sheet[cell] + ", &nbsp; &nbsp; ");
                });
                $ul.append("</li>");
              });
            });

            //$.each(zip.files, function (index, zipEntry) {
            //    $ul.append("<li>" + zipEntry.name + "</li>");
            //});

          } catch (e) {
            $ul.append("<li class='error'>Error reading " + theFile.name + " : " + e.message + "</li>");
          }
          $result.append($ul);
        }
      })(f);

      // read the file !
      // readAsArrayBuffer and readAsBinaryString both produce valid content for JSZip.
      reader.readAsArrayBuffer(f);
      // reader.readAsBinaryString(f);
    }
  });
});
