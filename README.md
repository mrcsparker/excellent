## Excellent

https://github.com/mrcsparker/excellent

### What is it?

Excellent is an XLSX parser and interpreter written in Javascript that runs on both the client (browser) and the server (node.js).

It is able to parse and evaluate XLSX files, giving you access to an XLSX file via a Javascript data structure.  It does this by translating the XLSX file (variables, formulas, strings, sheets) into Javascript-compatible functions and an easy-to-use data structure.

### Why?

So that people can still write their business logic in Excel and we can load it into our rich client apps.

### Installation

__Node__

    npm install excellent
    
__Browser__

Copy the `excellent.js` file to the appropriate directory.  Bower support is on the way.
### API

	// Show basic data structure
	var excellent = new Excellent();
	console.log(excellent.parseFile(YourXLSXFile);
	
	// Let's get some info about Sheet1/A1
	var excellent = new Excellent();
	var data = excellent.parseFile(xlxFile);	
	console.log(data.sheets.Sheet1.A1);
	
	// If you have a worksheet with a long name, you can access it like:
	console.log(data.sheets['Long-named worksheet'].A1);
		

### API Output

Excellent produces a simple data structure from an XLSX file.  It includes:

* Column names
* Formulas
* Variables
* Sheet information

Pivot tables, styles, and VB are all ignored.

#### Formula translation into Javascript

If you want to see how Excellent translates XLSX functions, there are some helpers available.  For example, if you have a formula which looks like:

	=SUM(A1,A2)

You can see the translation via:

	var excellent = new Excellent();
	var data = excellent.parseFile(xlsxFile);
	console.log(data.sheets.Sheet1._A1);
	
Which will produce something like:

	Formula.SUM(this.A1,this.A2)
	
All cell values have an underscore (`_A1`, `_A2`, `_A3`, etc) version with the raw associated data.


### TODO / Status

The code is pretty ugly right now.  Excel is a hairy format and I am still working my way through it.

Currently the library supports most Excel functionality, including:

* Most Excel formulas
* Shared strings : `sharedStrings.xml`
* Shared formulas
* Ranges : `SUM(A1:B9)`
* Cross-sheet identifiers: `=Sheet1!A8+Sheet2!A8`

I also have a set of unit tests to check in.  Right now, they are using sensitive data, so I am in the process of converting them over.

There are also a lot of things that Excellent can do that are not covered in this README including:

* Save Excel output to a serialized format that can be reloaded by Excellent later.
* Evaluate Excel data structures on-the-fly, including multi-tabbed Excel documents with cross-tab formulas and variables.

The API will be updated to include how to handle this, and other, features.

