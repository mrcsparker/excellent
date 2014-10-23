var assert = require('assert'),
  FormulaParser = require('../src/excellent.parser.js');

describe('ExcellentFormulaParser', function() {
  'use strict';

  describe('Atom', function() {
    it('should be able to parse a number', function() {
      assert.equal(FormulaParser.parse('1'), 1);
      assert.equal(FormulaParser.parse('100'), 100);
      assert.equal(FormulaParser.parse('101'), 101);
      assert.equal(FormulaParser.parse('201'), 201);
      assert.equal(FormulaParser.parse('234'), 234);
    });

    it('should be able to parse a floating point number', function() {
      assert.equal(FormulaParser.parse('1.01'), 1.01);
      assert.equal(FormulaParser.parse('12.01'), 12.01);
      assert.equal(FormulaParser.parse('99.99'), 99.99);
    });

    it('should be able to convert a percentage into a number', function() {
      assert.equal(FormulaParser.parse('100%'), 1.00);
      assert.equal(FormulaParser.parse('99%'), 0.99);
      assert.equal(FormulaParser.parse('75%'), 0.75);
      assert.equal(FormulaParser.parse('75%+25%'), '0.75+0.25');
      assert.equal(FormulaParser.parse('100%+D4'), '1+this.D4');
    });

    it('should be able to parse a simple paren expression', function() {
      assert.equal(FormulaParser.parse('(1)'), '(1)');
      assert.equal(FormulaParser.parse('(12)'), '(12)');
      assert.equal(FormulaParser.parse('(123)'), '(123)');
      assert.equal(FormulaParser.parse('(999)'), '(999)');
    });

    it('should be able to parse a double quoted string', function() {
      assert.equal(FormulaParser.parse('"a"'), '"a"');
      assert.equal(FormulaParser.parse('"Hello world"'), '"Hello world"');
      assert.equal(FormulaParser.parse('"asdfasdfasdf"'), '"asdfasdfasdf"');
    });

    it('should be able to parse a single quoted string', function() {
      assert.equal(FormulaParser.parse("'a'"), '"a"');
      assert.equal(FormulaParser.parse("'Hello world'"), '"Hello world"');
      assert.equal(FormulaParser.parse("'asdfasdfasdf'"), '"asdfasdfasdf"');
    });

  });

  describe('Arithmetic', function() {
    it('should be able to parse simple arithmetic', function() {
      assert.equal(FormulaParser.parse("1 + 1"), "1+1");
      assert.equal(FormulaParser.parse("1 + 1 + 1"), "1+1+1");
      assert.equal(FormulaParser.parse("1 + 1 + 2 + 3"), "1+1+2+3");
      assert.equal(FormulaParser.parse("999 + 998 + 997"), "999+998+997");
      assert.equal(FormulaParser.parse("9 + 8 - 7 + 6 + 3 - 2 * 1 / 2"), "9+8-7+6+3-2*1/2");
      assert.equal(FormulaParser.parse("(1+B5)^(B5-B4)"), "(1+this.B5)^(this.B5-this.B4)");
    });

    it('should be able to parse arithmetic with parens', function() {
      assert.equal(FormulaParser.parse('1 + (1 + 1)'), '1+(1+1)');
      assert.equal(FormulaParser.parse('1 + (1 + 1) + 1'), '1+(1+1)+1');
      assert.equal(FormulaParser.parse('1 + (1 + 1) / (2 * 2)'), '1+(1+1)/(2*2)');
      assert.equal(FormulaParser.parse('1 + ((1 + 1) + 2)'), '1+((1+1)+2)');
    });
  });

  describe('Identifiers', function() {
    it('should be able to return an identifier', function() {
      assert.equal(FormulaParser.parse('A1'), 'this.A1');
      assert.equal(FormulaParser.parse('B123'), 'this.B123');
      assert.equal(FormulaParser.parse('Z9'), 'this.Z9');
      assert.equal(FormulaParser.parse('A23'), 'this.A23');
      assert.equal(FormulaParser.parse('X10'), 'this.X10');
    });

    it('should be able to have identifiers mixed with arithmetic', function() {
      assert.equal(FormulaParser.parse('A1 + 1'), 'this.A1+1');
      assert.equal(FormulaParser.parse('A1 + 1 + B1'), 'this.A1+1+this.B1');
      assert.equal(FormulaParser.parse('A1 + 1 + C1'), 'this.A1+1+this.C1');
      assert.equal(FormulaParser.parse('A1 + 1    + D2 - 1'), 'this.A1+1+this.D2-1');
      assert.equal(FormulaParser.parse('A1 + A2'), 'this.A1+this.A2');
    });

    it('should be able to have arithmetic with parens', function() {
      assert.equal(FormulaParser.parse('(A1 + 1)'), '(this.A1+1)');
      assert.equal(FormulaParser.parse('(A1 + 1) + B1'), '(this.A1+1)+this.B1');
      assert.equal(FormulaParser.parse('((A1 + 1) + C1)'), '((this.A1+1)+this.C1)');
      assert.equal(FormulaParser.parse('A1 + 1    + (D2 - 1)'), 'this.A1+1+(this.D2-1)');
      assert.equal(FormulaParser.parse('A1 + (A2)'), 'this.A1+(this.A2)');
    });

    it('should be able to handle $ bill yall', function() {
      var output = FormulaParser.parse('SUM($C3*D3,$C4*D4,$C5*D5,$C6*D6,$C7*D7,$C8*D8)');
      assert.equal(output, 'Formula.SUM(this.$C3*this.D3,this.$C4*this.D4,this.$C5*this.D5,this.$C6*this.D6,this.$C7*this.D7,this.$C8*this.D8)');
    });

    it('should be able to handle a negative variable', function() {
      assert.equal(FormulaParser.parse('-D231'), '-this.D231');
    });
  });

  describe('Functions', function() {
    it('should be able to return a simple function call', function() {
      assert.equal(FormulaParser.parse('SUM(A1,A2)'), 'Formula.SUM(this.A1,this.A2)');
      assert.equal(FormulaParser.parse('ADD(A1,A2)'), 'Formula.ADD(this.A1,this.A2)');
      assert.equal(FormulaParser.parse('IF(A1,A2,A3)'), 'Formula.IF(this.A1,this.A2,this.A3)');
      assert.equal(FormulaParser.parse('TODAY()'), 'Formula.TODAY()');
    });

    it('should be able to mix arithmetic with function calls', function() {
      assert.equal(FormulaParser.parse('SUM(A1) + SUM(A2)'), 'Formula.SUM(this.A1)+Formula.SUM(this.A2)');
      assert.equal(FormulaParser.parse('ADD(A1) + SUM(A2)'), 'Formula.ADD(this.A1)+Formula.SUM(this.A2)');
      assert.equal(FormulaParser.parse('IF(A1,A2, A3) + SUM(A2)'), 'Formula.IF(this.A1,this.A2,this.A3)+Formula.SUM(this.A2)');
    });

    it('should be able to mix range with arithmetic with function calls', function() {
      assert.equal(FormulaParser.parse('INDEX(AC9:AL9,1)*L9'), 'Formula.INDEX([this.AC9,this.AD9,this.AE9,this.AF9,this.AG9,this.AH9,this.AI9,this.AJ9,this.AK9,this.AL9],1)*this.L9');
      assert.equal(FormulaParser.parse('INDEX(AC9:AL9,MATCH(I9,AC8:AL8,0))*L9'), 'Formula.INDEX([this.AC9,this.AD9,this.AE9,this.AF9,this.AG9,this.AH9,this.AI9,this.AJ9,this.AK9,this.AL9],Formula.MATCH(this.I9,[this.AC8,this.AD8,this.AE8,this.AF8,this.AG8,this.AH8,this.AI8,this.AJ8,this.AK8,this.AL8],0))*this.L9');
    });
  });

  describe('IF', function() {
    it('should handle simple IF', function() {
      assert.equal(FormulaParser.parse('IF(1, 2, 3)'), 'Formula.IF(1,2,3)');
      assert.equal(FormulaParser.parse('IF(1=1, 2, 3)'), 'Formula.IF(1==1,2,3)');
      assert.equal(FormulaParser.parse('IF(A1=1, 2, 3)'), 'Formula.IF(this.A1==1,2,3)');
      assert.equal(FormulaParser.parse("IF(A1='a', 2, 3)"), "Formula.IF(this.A1==\"a\",2,3)");
      assert.equal(FormulaParser.parse("IF(A1='', 2, 3)"), "Formula.IF(this.A1==\"\",2,3)");
    });
  });

  describe('Special characters', function() {
    it('should be able to handle excel-specific characters', function() {
      assert.equal(FormulaParser.parse('A&B'), 'this.A+""+this.B');
      assert.equal(FormulaParser.parse("'Hello'&'World'"), '"Hello"+""+"World"');
    });
  });

  describe('Range', function() {
    it('should be able to handle a simple range', function() {
      assert.equal(FormulaParser.parse('SUM(A1:A25)'),
        "Formula.SUM([this.A1,this.A2,this.A3,this.A4,this.A5,this.A6,this.A7,this.A8,this.A9,this.A10,this.A11,this.A12,this.A13,this.A14,this.A15,this.A16,this.A17,this.A18,this.A19,this.A20,this.A21,this.A22,this.A23,this.A24,this.A25])");
      assert.equal(FormulaParser.parse('SUM(A1:C1)'), 'Formula.SUM([this.A1,this.B1,this.C1])');
    });
  });

});
