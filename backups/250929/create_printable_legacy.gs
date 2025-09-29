function createPrintable() {
  console.log("createPrintable()");
  var originalFile = SpreadsheetApp.getActiveSpreadsheet(); //this sheetfile that should be copied
  var membersPerPage = 2;
  var filename = "Printable sheets " + Date();

  var sheetsToPrint=originalFile.getSheets();

  var newDocument=SpreadsheetApp.create(filename);

  var noOfSheets=sheetsToPrint.length;

  for(var i=0; i<noOfSheets; i++) {
    console.log("Behandlar ark: " + sheetsToPrint[i].getName());
    appendPrintableSheets(newDocument,sheetsToPrint[i], membersPerPage); 
  }
}

function appendPrintableSheets(newDocument, sheetToPrint, membersPerPage) {
  
  // Create group Total page by copying original sheet and remove everyting beyond summation column
  currentSheetName="Total order " + sheetToPrint.getName();

  //var currentSheet=sheetToPrint.copyTo(SpreadsheetApp.openById(newDocument.getId())).setName(currentSheetName);
  var currentSheet=sheetToPrint.copyTo(newDocument).setName(currentSheetName);

  // Add group info
  currentSheet.getRange("A1").setValue(sheetToPrint.getRange("A1").getDisplayValue()); // fixes issue with copied formulas, otherwise cell A1 will get #REF in some sheets

  // Add values of total orders
  currentSheet.getRange("B:B").setValues(sheetToPrint.getRange("B:B").getDisplayValues()); // copies values instead of formulas NB there is an s on the end!!

  // Adds row with sheet name
  currentSheet.insertRowBefore(1); 
  currentSheet.getRange("A1").setValue(currentSheetName);
  currentSheet.getRange("A1").setFontSize(15).setFontWeight("bold");

  // Add values to grand  total first sheet
  var totalSheet= newDocument.getSheets()[0];

  totalSheet.getRange(1,totalSheet.getLastColumn()+1,sheetToPrint.getMaxRows(),1).setValues(sheetToPrint.getRange("B:B").getDisplayValues()); // copies values instead of formulas NB there is an s on the end!!

  // and of added code
  setAlternatingColours(currentSheet); //set alternativ colours
  currentSheet.setFrozenColumns(0);
  currentSheet.deleteColumns(3,currentSheet.getMaxColumns()-2); //delete all other orders, just keep total
  ///////////////

  // number of sheets needed given members per page
  var numberOfSheets=Math.ceil((sheetToPrint.getMaxColumns()-2)/2/membersPerPage);
  
  for (var i=0; i < numberOfSheets;i++) { // create a few copies
    var currentSheetName=sheetToPrint.getName() + " " + (i+1)+ " of " + numberOfSheets;
    Logger.log(currentSheetName);
    
    //var currentSheet=sheetToPrint.copyTo(SpreadsheetApp.openById(newDocument.getId())).setName(currentSheetName); //duplicate sheet to new Document and changes name
    var currentSheet=sheetToPrint.copyTo(newDocument).setName(currentSheetName); //duplicate sheet to new Document and changes name
    currentSheet.getRange("A1").setValue(sheetToPrint.getRange("A1").getDisplayValue()); // fixes issue with copied formulas, otherwise cell A1 will get #REF in some sheets
    currentSheet.getRange("B:B").setValue(sheetToPrint.getRange("B:B").getDisplayValue()); // copies values instead of formulas
    // Adds row with sheet name
    currentSheet.insertRowBefore(1); 
    currentSheet.getRange("A1").setValue(currentSheetName);
    currentSheet.getRange("A1").setFontSize(15).setFontWeight("bold");

    setAlternatingColours(currentSheet); //set alternativ colours

    //delete the unwanted columns, starting at column 3
    currentSheetColumns=currentSheet.getMaxColumns();

    // delete all columns after current actual
    for (var j = 3 + membersPerPage*2*(1 + i); j <= currentSheetColumns; j++) {  
      currentSheet.deleteColumn(currentSheet.getMaxColumns());
    }

    if (i > 0) { //for all remaining columns that should not be there
      currentSheet.deleteColumns(3,i*membersPerPage*2);
    }

    currentSheet.deleteColumn(2); //Delete total column
  }
}

