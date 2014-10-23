var Excellent = Excellent || {};

Excellent.Util = (function() {
  'use strict';

  var self = {};

  // Variation of code I found on
  // http://davidwalsh.name/convert-xml-json
  self.xmlToJson = function(xml) {

    var obj = {},
      i, old, attribute, item;

    if (xml.nodeType === 1 && xml.attributes.length > 0) { // element

      obj['@'] = {};

      for (i = 0; i < xml.attributes.length; i += 1) {
        attribute = xml.attributes.item(i);
        obj['@'][attribute.nodeName] = attribute.nodeValue;
      }
    } else if (xml.nodeType === 3) { // text
      obj = xml.nodeValue;
    }

    if (xml.hasChildNodes()) {
      for (i = 0; i < xml.childNodes.length; i += 1) {

        item = xml.childNodes.item(i);

        if (obj[item.nodeName] === undefined) {
          obj[item.nodeName] = self.xmlToJson(item);
        } else {
          if (obj[item.nodeName].push === undefined) {
            old = obj[item.nodeName];
            obj[item.nodeName] = [];
            obj[item.nodeName].push(old);
          }
          obj[item.nodeName].push(self.xmlToJson(item));
        }

      }
    }
    return obj;
  };

  self.fromBase26 = function(number) {
    number = number.toUpperCase();

    var s = 0,
      i, dec = 0;

    if (number !== null && number !== undefined && number.length > 0) {
      for (i = 0; i < number.length; i += 1) {
        s = number.charCodeAt(number.length - i - 1) - 'A'.charCodeAt(0);
        dec += (Math.pow(26, i)) * (s + 1);
      }
    }

    return dec - 1;
  };

  self.toBase26 = function(value) {
    value = Math.abs(value);

    var converted = '',
      iteration = false,
      remainder;

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

  self.isNumber = function(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
  };

  self.getRowFromCell = function(val) {
    return parseInt(val.match(/[0-9]+/gi)[0], 10) - 1;
  };

  self.getColFromCell = function(val) {
    val = val.match(/[A-Z]+/gi)[0];
    return Excellent.Util.fromBase26(val);
  };

  self.each = function(obj, iterator, context) {
    if (obj === null || obj === undefined) {
      return;
    }

    var key, breaker = {};

    if (obj.forEach === Array.prototype.forEach) {
      obj.forEach(iterator, context);
    } else {

      for (key in obj) {
        if (obj.hasOwnProperty(key)) {
          if (iterator.call(context, obj[key], key, obj) === breaker) {
            return;
          }
        }
      }
    }
  };

  return self;
}());

if (typeof exports !== 'undefined') {
  exports.Excellent = Excellent;
}
