/**
 * To use
 * var Excellent = require('excellent').Excellent;
 */

require('traceur/bin/traceur-runtime');

exports.Excellent = require('./excellent.xlsx').Excellent;
exports.ExcellentLoader = require('./excellent.loader').Excellent;
