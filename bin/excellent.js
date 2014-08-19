#!/usr/bin/env node

var optimist = require('optimist'),
  Excellent = require('../lib').Excellent;

// optimist is a great library for taking the hassle
// out of parsing CLI options.
var argv = optimist
  .options('h', {
    alias: 'hello',
    describe: 'print hello world message.'
  }).argv;

if (argv.hello) {
  var hello = new Hello();
  console.log(hello.sayHello());
} else {
  console.log(optimist.help());
}
