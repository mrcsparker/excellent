/**
 * Excel formula parser for Excellent
 *
 * Parses Excel formulas into AST structures
 * that Excellent can evaluate without using eval.
 */

{
function literal(value) {
  return {
    type: 'Literal',
    value: value
  };
}

function errorLiteral(code) {
  return {
    type: 'ErrorLiteral',
    code: code
  };
}

function cellReference(ref, sheet) {
  return {
    type: 'CellReference',
    ref: ref,
    sheet: sheet || null
  };
}

function binaryExpression(left, operator, right) {
  return {
    type: 'BinaryExpression',
    left: left,
    operator: operator,
    right: right
  };
}

function unaryExpression(operator, argument) {
  return {
    type: 'UnaryExpression',
    operator: operator,
    argument: argument
  };
}

function arrayExpression(elements) {
  return {
    type: 'ArrayExpression',
    elements: elements
  };
}

function formulaCallExpression(name, args) {
  return {
    type: 'FormulaCallExpression',
    name: name,
    arguments: args
  };
}

function foldBinary(head, tail) {
  var result = head;
  var i;

  for (i = 0; i < tail.length; i += 1) {
    result = binaryExpression(result, tail[i][1], tail[i][3]);
  }

  return result;
}

function fromBase26(number) {
  "use strict";

  number = number.toUpperCase();

  var s = 0;
  var i;
  var dec = 0;

  if (number !== null && typeof number !== "undefined" && number.length > 0) {
    for (i = 0; i < number.length; i += 1) {
      s = number.charCodeAt(number.length - i - 1) - "A".charCodeAt(0);
      dec += (Math.pow(26, i)) * (s + 1);
    }
  }

  return dec - 1;
}

function toBase26(value) {
  "use strict";

  value = Math.abs(value);

  var converted = "";
  var iteration = false;
  var remainder;

  do {
    remainder = value % 26;

    if (iteration && value < 25) {
      remainder -= 1;
    }

    converted = String.fromCharCode((remainder + "A".charCodeAt(0))) + converted;
    value = Math.floor((value - remainder) / 26);
    iteration = true;
  } while (value > 0);

  return converted;
}

function buildRange(start, end) {
  var startRow = parseInt(start.match(/[0-9]+/gi)[0], 10);
  var startCol = start.match(/[A-Z]+/gi)[0];
  var startColDec = fromBase26(startCol);
  var endRow = parseInt(end.match(/[0-9]+/gi)[0], 10);
  var endCol = end.match(/[A-Z]+/gi)[0];
  var totalRows = endRow - startRow + 1;
  var totalCols = fromBase26(endCol) - fromBase26(startCol) + 1;
  var curCol = 0;
  var curRow = 1;
  var cells = [];
  var curCell = "";

  for (; curRow <= totalRows; curRow += 1) {
    for (; curCol < totalCols; curCol += 1) {
      curCell = toBase26(startColDec + curCol) + "" + (startRow + curRow - 1);
      cells.push(cellReference(curCell));
    }
    curCol = 0;
  }

  return arrayExpression(cells);
}
}

start
  = Expression

Expression
  = ComparisonExpression

ComparisonExpression
  = head:ConcatenationExpression
    tail:(__ operator:("=" / "<" / ">") __ right:ConcatenationExpression)* {
      var normalizedTail = tail.map(function(entry) {
        return [entry[0], entry[1] === "=" ? "==" : entry[1], entry[2], entry[3]];
      });

      return foldBinary(head, normalizedTail);
    }

ConcatenationExpression
  = head:AdditiveExpression
    tail:(__ operator:"&" __ right:AdditiveExpression)* {
      return foldBinary(head, tail);
    }

AdditiveExpression
  = head:MultiplicativeExpression
    tail:(__ operator:("+" / "-") __ right:MultiplicativeExpression)* {
      return foldBinary(head, tail);
    }

MultiplicativeExpression
  = head:PowerExpression
    tail:(__ operator:("*" / "/") __ right:PowerExpression)* {
      return foldBinary(head, tail);
    }

PowerExpression
  = head:UnaryExpression
    tail:(__ operator:"^" __ right:UnaryExpression)* {
      return foldBinary(head, tail);
    }

UnaryExpression
  = operator:("+" / "-") __ argument:UnaryExpression {
      return unaryExpression(operator, argument);
    }
  / Primary

Primary
  = ParenExpression
  / Number
  / ErrorLiteral
  / Range
  / FunctionCall
  / Variable
  / String

// (A1 + B1) + C1
ParenExpression
  = "(" __ expression:Expression __ ")" {
      return expression;
    }

Number
  // Percentages
  = digits:([0-9]+ "%") {
      return literal(parseFloat(digits.join("").replace(/,/g, "").replace(/%/g, ""), 10) / 100.0);
    }
    // Regular floating-point numbers
    /
    digits:([0-9]+ ("." [0-9]+)? (("e" / "E") ("+" / "-")? [0-9]+)?) {
      return literal(parseFloat(digits.join("").replace(/,/g,""), 10));
    }

String
  = '"' str:DoubleQuotedCharacter* '"' {
      return literal(str.join(""));
    }
  / "'" str:SingleQuotedCharacter* "'" {
      return literal(str.join(""));
    }

DoubleQuotedCharacter
  = '""' {
      return '"';
    }
  / char:[^"\n] {
      return char;
    }

SingleQuotedCharacter
  = "''" {
      return "'";
    }
  / char:[^'\n] {
      return char;
    }

ErrorLiteral
  = "#DIV/0!" {
      return errorLiteral('#DIV/0!');
    }
  / "#VALUE!" {
      return errorLiteral('#VALUE!');
    }
  / "#REF!" {
      return errorLiteral('#REF!');
    }
  / "#NAME?" {
      return errorLiteral('#NAME?');
    }
  / "#N/A" {
      return errorLiteral('#N/A');
    }

FunctionCall
  = __ name:Identifier __ args:(Arguments) {
    return formulaCallExpression(name, args);
  }

Range
  = start:Identifier __ ":" __ end:Identifier {
    return buildRange(start, end);
}

Arguments
  = "(" __ args:ArgumentList? __ ")" {
    return args || [];
  }

ArgumentList
  = head:Expression tail:(__ "," __ Expression)* {
    var result = [head];
    for (var i = 0; i < tail.length; i += 1) {
      result.push(tail[i][3]);
    }
    return result;
  }

Variable
  = "'" workbook:AlphaNumeric "'!" name:IdentifierVariable {
    return cellReference(name, workbook.join(""));
  }
  / workbook:Identifier "!" name:IdentifierVariable {
      return cellReference(name, workbook);
    }
  / name:IdentifierVariable {
      return cellReference(name);
    }

IdentifierVariable
  = name:Identifier {
      return name;
  }

Identifier
  = name:IdentifierName {
      return name
    }

IdentifierName
  = start:IdentifierStart parts:IdentifierPart* {
      return start + parts.join("");
    }

IdentifierStart
  = Alpha
  / "$"
  / "_"

IdentifierPart
  = IdentifierStart
  / Digit

AlphaNumeric
  = ('&' / '-' / '_' / ' ' / [a-z] / [A-Z] / [0-9])*

Alpha
  = [a-zA-Z]

Digits
  = (Digit)*

Digit
  = [0-9]
__
  = WhiteSpace*

WhiteSpace "whitespace"
  = [ \t\r\n]
