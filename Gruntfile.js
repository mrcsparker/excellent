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

module.exports = function(grunt) {

  'use strict';

  var copyright = '/*\n' +
    'Copyright 2014 Chris Parker. Licensed under the Apache 2.0 License.\n' +
    'http://www.apache.org/licenses/LICENSE-2.0.html\n' +
    '*/\n\n';


  var excellentOpen = '(function() {\n"use strict";\n\n';
  var excellentClose = '\n})();';

  var project = {
    files: ['bower_components/jstat/index.js',
      'bower_components/lodash/dist/lodash.compat.js',
      'bower_components/momentjs/moment.js',
      'bower_components/numeral/index.js',
      'bower_components/numeric/index.js',
      'bower_components/underscore.string/index.js',
      'src/excellent.compat.js',
      'bower_components/formulajs/lib/formula.js',
      'bower_components/jszip/jszip.js',
      'bower_components/jszip/jszip-load.js',
      'bower_components/jszip/jszip-inflate.js',
      'src/excellent.parser.browser.js',
      'src/excellent.util.js',
      'src/excellent.workbook.js',
      'src/excellent.loader.js',
      'src/excellent.xlsx.js',
      'src/excellent.xlsx-simple.js'
    ]
  };

  /*
  var projectSimple = {
    files: ['bower_components/jszip/jszip.js',
      'bower_components/jszip/jszip-load.js',
      'bower_components/jszip/jszip-inflate.js',
      'src/excellent.util.js',
			'src/excellent.workbook.js',
			'src/excellent.xlsx-simple.js'
    ]
  };
  */

  grunt.initConfig({
    pkg: grunt.file.readJSON('package.json'),

    header: copyright + excellentOpen,
    copyright: copyright,
    footer: excellentClose,

    concat: {
      options: {
        stripBanners: 'true',
        banner: '<%= header %>',
        footer: '<%= footer %>'
      },
      dist: {
        src: project.files,
        dest: '<%= pkg.name %>.js'
      }
    },

    uglify: {
      options: {
        preserveComments: 'false',
        banner: '<%= copyright %>'
      },
      dist: {
        src: ['<%= concat.dist.dest %>'],
        dest: '<%= pkg.name %>.min.js'
      }
    },

    mochaTest: {
      test: {
        options: {
          reporter: 'spec',
          captureFile: 'results.txt', // Optionally capture the reporter output to a file
          quiet: false // Optionally suppress output to standard out (defaults to false)
        },
        src: ['test/**/*.js']
      }
    }

  });

  grunt.loadNpmTasks('grunt-contrib-uglify');
  grunt.loadNpmTasks('grunt-contrib-concat');
  grunt.loadNpmTasks('grunt-mocha-test');

  grunt.registerTask('default', ['concat', 'uglify']);
  grunt.registerTask('build', ['concat', 'uglify']);
  grunt.registerTask('test', ['mochaTest']);
};
