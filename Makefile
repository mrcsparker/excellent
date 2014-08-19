# Add binaries from NPM modules to PATH
PATH:=node_modules/.bin:${PATH}

BOWER_SCRIPTS = \
	bower_components/jstat/index.js \
	bower_components/lodash/dist/lodash.compat.js \
	bower_components/momentjs/moment.js \
	bower_components/numeral/index.js \
	bower_components/numeric/index.js \
	bower_components/underscore.string/index.js \
	lib/excellent.compat.js \
	bower_components/formulajs/lib/formula.js \
	bower_components/jszip/jszip.js \
	bower_components/jszip/jszip-load.js \
	bower_components/jszip/jszip-inflate.js


SULLA_SCRIPTS = \
	lib/excellent.parser.browser.js \
	lib/excellent.util.js \
	lib/excellent.workbook.js \
	lib/excellent.loader.js \
	lib/excellent.xlsx.js \
	lib/excellent.xlsx-simple.js

BOWER_SIMPLE_SCRIPTS = \
	bower_components/jszip/jszip.js \
	bower_components/jszip/jszip-load.js \
	bower_components/jszip/jszip-inflate.js

SULLA_SIMPLE_SCRIPTS = \
	lib/excellent.util.js \
	lib/excellent.workbook.js \
	lib/excellent.xlsx-simple.js

test:
	mocha -u bdd

peg:
	pegjs --track-line-and-column ./grammar/formula_parser.pegjs ./lib/excellent.parser.js
	pegjs -e FormulaParser --track-line-and-column ./grammar/formula_parser.pegjs ./lib/excellent.parser.browser.js

build:
	rm -rf dist
	mkdir -p dist
	cat $(BOWER_SCRIPTS) > dist/bower.combine.js
	uglifyjs dist/bower.combine.js > dist/bower.uglify.js
	cat $(SULLA_SCRIPTS) > dist/excellent.combine.js
	uglifyjs dist/excellent.combine.js > dist/excellent.uglify.js
	cat dist/bower.uglify.js dist/excellent.uglify.js > dist/excellent.min.js
	cat dist/bower.combine.js dist/excellent.combine.js > dist/excellent.js
	cp dist/excellent.js excellent.js
	cp dist/excellent.min.js excellent.min.js

build-simple:
	rm -rf dist
	mkdir -p dist
	cat $(BOWER_SIMPLE_SCRIPTS) > dist/bower.combine.js
	uglifyjs dist/bower.combine.js > dist/bower.uglify.js
	cat $(SULLA_SIMPLE_SCRIPTS) > dist/excellent.combine.js
	uglifyjs dist/excellent.combine.js > dist/excellent.uglify.js
	cat dist/bower.uglify.js dist/excellent.uglify.js > dist/excellent.min.js
	cat dist/bower.combine.js dist/excellent.combine.js > dist/excellent.js
	cp dist/excellent.js excellent-simple.js
	cp dist/excellent.min.js excellent-simple.min.js

.PHONY: test
