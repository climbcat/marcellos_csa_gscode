

250824:

Initialized the backup-able dev project for the CSA-MF order list.

The main issues are:
1) backup/version control
2) the awful performance of running the damn .gs script in the free cloud service - [prohibitive] as it times out after 6 minutes

The google sheet itself can be backed up in .xlsx format.
The code can be backed up by copy-paste (until I find the download button).

We can also push the xlsx, but I don't see any reason why we would do this.
For a performant solution, we could build the output sheet as a xlsx using some JS api.
Such xlsx can be imported/opened/printed/uploaded/app-accessible - whatever works for users.

The plan:

Working on the .xlsx, outputting printable .xlsx would be preferable, since performance would not be an issue ever again.

This process could be wrapped in a script or web app. Maybe a dev could make a simple web app with a button and a list of 
historical printable sheets or something.

In order to make the current .gs script translate into something that works on the .xlsx, we need to do some tricks.
[Ultimately the .gs script should be ported to just be some node-JS stuff working on the .xlsx.]

Tricks from gptee below; First, practices to speed up the .gs script performance, Second, the plan to port into .xlsx.


gptee tips:


1) values: 

<pre>

let range = sheet.getRange("A1:D100");
let values = range.getValues(); // 2D array

...

range.setValues(values);
</pre>


2) formatting (batched): 

<pre>
range.setBackgrounds([
  ["#ffff00", "#ff0000", "#00ff00", "#0000ff"],
  ["#ffffff", "#cccccc", "#999999", "#000000"]
]);

setFontWeights()
setFontColors()
setNumberFormats()
</pre>

So instead of looping cell by cell, build a 2D array of formatting and apply once.



3) 

Let’s build a mini fake SpreadsheetApp in Node that works with a local spreadsheet file.

- The following code example mimics:

SpreadsheetApp.getActiveSpreadsheet()

getSheetByName()

getRange().getValues()

getRange().setValues()

- It’s only a tiny subset of the real API (no formatting, no formulas, etc.), but enough to test most data logic locally.

<pre>
// fakeSpreadsheetApp.js
const XLSX = require("xlsx");

class FakeRange {
  constructor(sheet, range, data) {
    this.sheet = sheet;
    this.range = range;
    this.data = data; // 2D array
  }

  getValues() {
    return this.data;
  }

  setValues(values) {
    this.data = values;
    // write back to worksheet
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cellAddr = XLSX.utils.encode_cell({
          r: this.range.s.r + r,
          c: this.range.s.c + c,
        });
        this.sheet[cellAddr] = { t: "s", v: values[r][c] };
      }
    }
  }
}

class FakeSheet {
  constructor(workbook, sheetName) {
    this.workbook = workbook;
    this.sheet = workbook.Sheets[sheetName];
  }

  getRange(r, c, numRows = 1, numCols = 1) {
    const values = [];
    for (let i = 0; i < numRows; i++) {
      const row = [];
      for (let j = 0; j < numCols; j++) {
        const cellAddr = XLSX.utils.encode_cell({ r: r - 1 + i, c: c - 1 + j });
        const cell = this.sheet[cellAddr];
        row.push(cell ? cell.v : "");
      }
      values.push(row);
    }

    return new FakeRange(this.sheet, { s: { r: r - 1, c: c - 1 } }, values);
  }
}

class FakeSpreadsheetApp {
  constructor(filename) {
    this.workbook = XLSX.readFile(filename);
    this.filename = filename;
  }

  getActiveSpreadsheet() {
    return this;
  }

  getSheetByName(name) {
    return new FakeSheet(this.workbook, name);
  }

  save(filename = this.filename) {
    XLSX.writeFile(this.workbook, filename);
  }
}

module.exports = FakeSpreadsheetApp;
</pre>

Usage example: 

<pre>
const FakeSpreadsheetApp = require("./fakeSpreadsheetApp");

// load local Excel file
const app = new FakeSpreadsheetApp("test.xlsx");
const sheet = app.getSheetByName("Sheet1");

// get range A1:B2
const range = sheet.getRange(1, 1, 2, 2);
console.log("Before:", range.getValues());

// modify values
range.setValues([
  ["Hello", "World"],
  ["Foo", "Bar"],
]);

// save back
app.save("test-modified.xlsx");
console.log("After:", range.getValues());
</pre>



