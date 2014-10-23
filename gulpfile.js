/**
 * Copyright 2014 Chris Parker
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

'use strict';

var gulp = require('gulp');
var gutil = require('gulp-util');
var jscs = require('gulp-jscs');
var jshint = require('gulp-jshint');
var mocha  = require('gulp-mocha');

gulp.task('build-lib', function() {
  gulp.src([
    'src/excellent.parser.browser.js',
    'src/excellent.util.js',
    'src/excellent.workbook.js',
    'src/excellent.loader.js',
    'src/excellent.xlsx.js',
    'src/excellent.xlsx-simple.js'
  ]);
});

gulp.task('jscs', function() {
  gulp.src('src/*.js')
    .pipe(jscs());
});

gulp.task('jshint', function() {
  return gulp
    .src(['gulpfile.js', 'src/*.js', 'test/*.js'])
    .pipe(jshint())
    .pipe(jshint.reporter('default'));
});

gulp.task('test', function() {
  return gulp
    .src('test/*.js')
    .pipe(mocha());
});

gulp.task('default', function() {
  // Default task code
});
