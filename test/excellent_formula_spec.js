var assert = require('assert');
var fs = require('fs-extra');
var Excellent = require('../src').Excellent;
var Formula = require('formulajs');

describe('ExcellentFormula', function() {
  'use strict';

  it('should be able to evaluate simple formulas from formula.js', function() {
    assert.equal(Formula.SUM(1, 2), 3);
    assert.equal(Formula.DELTA(42, 42), 1);
  });

  describe('Simple Formulas: basic tests of Formula / Excellent integration', function() {

    var xlsxFile;
    var excellent;
    var parsed;
    var content;

    xlsxFile = fs.readFileSync('./test/data/simpleFormula.xlsx', 'binary');
    excellent = new Excellent.Xlsx();
    parsed = excellent.load(xlsxFile);
    content = parsed.workbook.Sheet1;

    it('SUM(1,2)', function() {
      assert.equal(content._A1, 'Formula.SUM(1,2)');
      assert.equal(content.A1, 3);
    });

    it('AVERAGE(A1)', function() {
      assert.equal(content._A2, 'Formula.AVERAGE(this.A1)');
      assert.equal(content.A2, 3);
    });

    it('COUNT(A2)', function() {
      assert.equal(content._A3, 'Formula.COUNT(this.A2)');
      assert.equal(content.A3, 1);
    });

    it('POWER(10,2)', function() {
      assert.equal(content._A4, 'Formula.POWER(10,2)');
      assert.equal(content.A4, 100);
    });

    it('SQRT(100)', function() {
      assert.equal(content._A5, 'Formula.SQRT(100)');
      assert.equal(content.A5, 10);
    });

    it('FACT(5)', function() {
      assert.equal(content._A6, 'Formula.FACT(5)');
      assert.equal(content.A6, 120);
    });

    it('FLOOR(100,2)', function() {
      assert.equal(content._A7, 'Formula.FLOOR(100,2)');
      assert.equal(content.A7, 100);
    });

    it('ABS(-1)', function() {
      assert.equal(content._A8, 'Formula.ABS(-1)');
      assert.equal(content.A8, 1);
    });

    it('MOD(55,3)', function() {
      assert.equal(content._A9, 'Formula.MOD(55,3)');
      assert.equal(content.A9, 1);
    });
  });

  describe('Simple ranges', function() {
    var xlsxFile;
    var excellent;
    var parsed;
    var content;

    xlsxFile = fs.readFileSync('./test/data/simpleRange.xlsx', 'binary');
    excellent = new Excellent.Xlsx();
    parsed = excellent.load(xlsxFile);
    content = parsed.workbook.Sheet1;

    it('SUM(A1:H1)', function() {
      assert.equal(content._A20, "Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1])");
      assert.equal(content.A20, 36);
    });

    it('SUM(B1:H3)', function() {
      assert.equal(content._A21, "Formula.SUM([this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])");
      assert.equal(content.A21, 273);
    });

    it('SUM(I3:L3)', function() {
      assert.equal(content._A22, "Formula.SUM([this.I3,this.J3,this.K3,this.L3])");
      assert.equal(content.A22, 106);
    });

    it('SUM(A1:H3)', function() {
      assert.equal(content._A23, "Formula.SUM([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])");
      assert.equal(content.A23, 300);
    });

    it('AVERAGE(A1:H3)', function() {
      assert.equal(content._A24, "Formula.AVERAGE([this.A1,this.B1,this.C1,this.D1,this.E1,this.F1,this.G1,this.H1,this.A2,this.B2,this.C2,this.D2,this.E2,this.F2,this.G2,this.H2,this.A3,this.B3,this.C3,this.D3,this.E3,this.F3,this.G3,this.H3])");
      assert.equal(content.A24, 12.5);
    });

    it('SUM(B1:B16)', function() {
      assert.equal(content._A25, "Formula.SUM([this.B1,this.B2,this.B3,this.B4,this.B5,this.B6,this.B7,this.B8,this.B9,this.B10,this.B11,this.B12,this.B13,this.B14,this.B15,this.B16])");
      assert.equal(content.A25, 134);
    });
  });
});
