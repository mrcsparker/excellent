/**
 * Excel formula parser for Excellent
 *
 * Translates Excel formulas into javascript
 * data structures that Excellent can evaluate
 */

{
// From https://github.com/joshatjben/excelFormulaUtilitiesJS
// Modified for Excellent

// This was Modified from a function at http://en.wikipedia.org/wiki/Hexavigesimal
// Pass in the base 26 string, get back integer
var fromBase26 = function (number) {

    "use strict";

    number = number.toUpperCase();

    var s = 0, i, dec = 0;

    if (number !== null && typeof number !== "undefined" && number.length > 0) {
        for (i = 0; i < number.length; i += 1) {
            s = number.charCodeAt(number.length - i - 1) - "A".charCodeAt(0);
            dec += (Math.pow(26, i)) * (s + 1);
        }
    }

    return dec - 1;
};

var toBase26 = function (value) {

    "use strict";

    value = Math.abs(value);

    var converted = "", iteration = false, remainder;

    // Repeatedly divide the number by 26 and convert the
    // remainder into the appropriate letter.
    do {
        remainder = value % 26;

        // Compensate for the last letter of the series being corrected on 2 or more iterations.
        if (iteration && value < 25) {
            remainder -= 1;
        }

        converted = String.fromCharCode((remainder + 'A'.charCodeAt(0))) + converted;
        value = Math.floor((value - remainder) / 26);

        iteration = true;
    } while (value > 0);

    return converted;
};
}

start
  = Expression

Expression
  = ArithmeticExpression

ArithmeticOperator
  = "*"
  / "/"
  / "+"
  / "-"
  / "^"
  / "="
  / "&"
  / "<"
  / ">"

ArithmeticExpression
  = head:Atom
    tail:(__ ArithmeticOperator __ Atom)* {
      var result = [head],
        oper;

      for (var i = 0; i < tail.length; i += 1) {
        oper = tail[i][1];

        oper = oper.replace('=', '==');
        oper = oper.replace('&', '+""+');

        result.push(oper);
        result.push(tail[i][3]);
      }

      return result.join("");
  }

// Atoms self-evalute.
// This is the simplest expression form
Atom
  = ParenExpression
  / Number
  / Range
  / FunctionCall
  / Variable
  / String

// (A1 + B1) + C1
ParenExpression
  = "(" expression:Expression ")" {
      return "(" + expression + ")";
    }

Number
  // Percentages
  = digits:([0-9]+ "%") {
      console.log(parseFloat(digits.join("").replace(/,/g, "").replace(/%/g, ""), 10));
      return parseFloat(digits.join("").replace(/,/g, "").replace(/%/g, ""), 10) / 100.0;
    }
    // Regular floating-point numbers
    /
    digits:("-"? [0-9]+ ("." [0-9]+)? ("e"[0-9]+)?) {
      return parseFloat(digits.join("").replace(/,/g,""), 10);
    }

String
  = '"' str:[^'"\n]+ '"' {
      return '"' + str.join("") + '"';
    }
  / "'" str:[^'"\n]+ "'" {
      return '"' + str.join("") + '"';
    }
  / "''" {
    return '""';
  }
  / '""' {
    return '""';
  }

FunctionCall
  = __ name:Identifier __ arguments:(Arguments) {
    return "Formula." + name + arguments;
  }

Range
  = start:Identifier __ ":" __ end:Identifier {

    // Taken from https://github.com/joshatjben/excelFormulaUtilitiesJS
    // and modified for Excellent

    var startRow = parseInt(start.match(/[0-9]+/gi)[0]),
        startCol = start.match(/[A-Z]+/gi)[0],
        startColDec = fromBase26(startCol),

        endRow = parseInt(end.match(/[0-9]+/gi)[0]),
        endCol = end.match(/[A-Z]+/gi)[0],
        endColDec = fromBase26(endCol),

        // Total rows and cols
        totalRows = endRow - startRow + 1,
        totalCols = fromBase26(endCol) - fromBase26(startCol) + 1,

        // Loop vars
        curCol = 0,
        curRow = 1,
        curCell = "",
        retVal = [];

        for (; curRow <= totalRows; curRow += 1) {
            for (; curCol < totalCols; curCol += 1) {
                // Get the current cell id
                curCell = toBase26(startColDec + curCol) + "" + (startRow+curRow-1) ;
                retVal.push("this." + curCell);
            }
            curCol = 0;
        }

    return "[" + retVal.join(",") + "]";
}

Arguments
  = "(" __ arguments:ArgumentList? __ ")" {
    if (arguments.length > 0) {
      return "(" + arguments.join(",") + ")";
    } else {
      return "()";
    }
  }

ArgumentList
  = head:Expression tail:(__ "," __ Expression)* {
    var result = [head];
    for (var i = 0; i < tail.length; i++) {
      result.push(tail[i][3]);
    }
    return result;
  }

Variable
  = "'" workbook:AlphaNumeric "'!" name:IdentifierVariable {
    return "self.workbook['" + workbook.join("") + "']." + name;
  }
  / workbook:Identifier "!" name:IdentifierVariable {
      return "self.workbook['" + workbook + "']." + name;
    }
  / "-" name:IdentifierVariable {
      return "-this." + name;
    }
  / name:IdentifierVariable {
      return "this." + name;
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
  = ('&' / '-' / '_' / ' ' / [a-z] / [A-Z])*

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

