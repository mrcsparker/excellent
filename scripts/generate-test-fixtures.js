'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const JSZip = require('jszip');

function createContentTypesXml(sheetCount) {
  const worksheetOverrides = Array.from({ length: sheetCount }, function(_value, index) {
    return '<Override PartName="/xl/worksheets/sheet' + String(index + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
  }).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    worksheetOverrides,
    '</Types>'
  ].join('');
}

function createRootRelationshipsXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>'
  ].join('');
}

function createWorkbookRelationshipsXml(sheetCount) {
  const sheetRelationships = Array.from({ length: sheetCount }, function(_value, index) {
    return '<Relationship Id="rId' + String(index + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + String(index + 1) + '.xml"/>';
  }).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    sheetRelationships,
    '</Relationships>'
  ].join('');
}

function createWorkbookXml(sheetNames) {
  const sheetXml = sheetNames.map(function(sheetName, index) {
    return '<sheet name="' + sheetName + '" sheetId="' + String(index + 1) + '" r:id="rId' + String(index + 1) + '"/>';
  }).join('');

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="24026"/>',
    '<sheets>',
    sheetXml,
    '</sheets>',
    '</workbook>'
  ].join('');
}

function createWorksheetXml(rows) {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<sheetData>',
    rows.join(''),
    '</sheetData>',
    '</worksheet>'
  ].join('');
}

async function writeFixture(outputDirectory, fixture) {
  const zip = new JSZip();
  const sheetNames = fixture.sheets.map(function(sheet) {
    return sheet.name;
  });

  zip.file('[Content_Types].xml', createContentTypesXml(fixture.sheets.length));
  zip.file('_rels/.rels', createRootRelationshipsXml());
  zip.file('xl/workbook.xml', createWorkbookXml(sheetNames));
  zip.file('xl/_rels/workbook.xml.rels', createWorkbookRelationshipsXml(fixture.sheets.length));

  fixture.sheets.forEach(function(sheet, index) {
    zip.file('xl/worksheets/sheet' + String(index + 1) + '.xml', sheet.xml);
  });

  await fs.writeFile(
    path.join(outputDirectory, fixture.fileName),
    await zip.generateAsync({
      compression: 'DEFLATE',
      type: 'nodebuffer'
    })
  );
}

async function main() {
  const outputDirectory = path.join(__dirname, '..', 'test', 'data');
  const fixtures = [
    {
      fileName: 'sharedFormulas.xlsx',
      sheets: [
        {
          name: 'Shared',
          xml: createWorksheetXml([
            '<row r="1" spans="1:3"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1+1</f><v>2</v></c><c r="C1"><f>SUM(B1:B3)</f><v>18</v></c></row>',
            '<row r="2" spans="1:2"><c r="A2"><v>5</v></c><c r="B2"><f t="shared" si="0"/><v>6</v></c></row>',
            '<row r="3" spans="1:2"><c r="A3"><v>9</v></c><c r="B3"><f t="shared" si="0"/><v>10</v></c></row>'
          ])
        }
      ]
    },
    {
      fileName: 'crossSheetWorkbook.xlsx',
      sheets: [
        {
          name: 'Inputs',
          xml: createWorksheetXml([
            '<row r="1" spans="1:1"><c r="A1"><v>4</v></c></row>',
            '<row r="2" spans="1:1"><c r="A2"><v>5</v></c></row>'
          ])
        },
        {
          name: 'Outputs',
          xml: createWorksheetXml([
            '<row r="1" spans="1:2"><c r="A1"><f>Inputs!A1+1</f><v>5</v></c><c r="B1"><f>SUM(Inputs!A1,A2)</f><v>10</v></c></row>',
            '<row r="2" spans="1:1"><c r="A2"><f>Inputs!A2+1</f><v>6</v></c></row>'
          ])
        }
      ]
    },
    {
      fileName: 'quotedSheetAndErrors.xlsx',
      sheets: [
        {
          name: 'Budget 2026',
          xml: createWorksheetXml([
            '<row r="1" spans="1:1"><c r="A1"><v>7</v></c></row>',
            '<row r="2" spans="1:1"><c r="A2"><v>8</v></c></row>'
          ])
        },
        {
          name: 'Summary',
          xml: createWorksheetXml([
            '<row r="1" spans="1:1"><c r="A1"><f>\'Budget 2026\'!A1+1</f><v>8</v></c></row>',
            '<row r="2" spans="1:1"><c r="A2"><v>2</v></c></row>',
            '<row r="3" spans="1:1"><c r="A3"><f>SUM($A$2,\'Budget 2026\'!A2)</f><v>10</v></c></row>',
            '<row r="4" spans="1:1"><c r="A4"><f>IF("He said ""hi"""="He said ""hi""",1,0)</f><v>1</v></c></row>',
            '<row r="5" spans="1:1"><c r="A5"><f>IFERROR(#DIV/0!,99)</f><v>99</v></c></row>',
            '<row r="6" spans="1:1"><c r="A6"><f>IFNA(#N/A,77)</f><v>77</v></c></row>'
          ])
        }
      ]
    }
  ];

  await fs.mkdir(outputDirectory, { recursive: true });

  for (const fixture of fixtures) {
    await writeFixture(outputDirectory, fixture);
  }
}

main().catch(function(error) {
  process.stderr.write(String(error) + '\n');
  process.exitCode = 1;
});
