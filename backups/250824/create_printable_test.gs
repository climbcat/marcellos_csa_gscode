function createPrintable_Test() {
    console.log("createPrintable_Test()");

    var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var destDoc = SpreadsheetApp.create("Printable_sheets_TEST_" + Date());

    createTemplate(destDoc, srcSheets[0]);
    templateSheet = destDoc.getSheets()[1];

    for(var i = 0; i < srcSheets.length; i++) {
        console.log("Behandlar ark: " + srcSheets[i].getName());
        updateSharedTotalsPage(destDoc, srcSheets[i]); 
        createGroupTotalsPage(destDoc, srcSheets[i]);
        createMemberPages(destDoc, srcSheets[i], templateSheet); 
    }

    destDoc.deleteSheet(templateSheet);
}

function updateSharedTotalsPage(destDoc, srcSheet) {
  // Add values to grand  total first sheet
  var totalSheet = destDoc.getSheets()[0];
  totalSheet
      .getRange( 1, totalSheet.getLastColumn() + 1, srcSheet.getMaxRows(), 1 )
      .setValues( srcSheet.getRange("B:B").getDisplayValues() );
}

function createGroupTotalsPage(destDoc, srcSheet) {
    currentSheetName = "Total order " + srcSheet.getName();

    var destSheet = srcSheet.copyTo(destDoc)
                            .setName(currentSheetName);

    // Add group info
    destSheet
        .getRange("A1")
        .setValue( srcSheet.getRange("A1").getDisplayValue() );

    // Add values of total orders
    destSheet
        .getRange("B:B")
        .setValues( srcSheet.getRange("B:B").getDisplayValues() );

    // Adds row with sheet name
    destSheet.insertRowBefore(1); 
    destSheet
        .getRange("A1")
        .setValue(currentSheetName);
    destSheet
        .getRange("A1")
        .setFontSize(15).setFontWeight("bold");

    setAlternatingColours(destSheet); // set alternativ colours
    destSheet.setFrozenColumns(0);
    destSheet.deleteColumns( 3,destSheet.getMaxColumns() - 2 );
}


function createTemplate(destDoc, srcSheet) {
    srcSheet
        .copyTo(destDoc)
        .setName("template");
    templateSheet = destDoc.getSheets()[1];

    var srcRowCount = srcSheet.getMaxRows();

    templateSheet.deleteColumns(5, templateSheet.getMaxColumns() - 5);
    templateSheet.setName("template");

    // set labels/descriptions column
    var data = srcSheet
        .getRange(1, 1, srcRowCount, 1)
        .getValues();
    templateSheet
        .getRange(2, 1, srcRowCount, 1)
        .setValues(data);
}


function createMemberPages(destDoc, srcSheet, templateSheet) {
    var srcRowCount = srcSheet.getMaxRows();
    
    // NOTE: this field 1,1 is different for each group, with contact info, etc.
    var note = srcSheet
        .getRange(1, 1, 1, 1)
        .getValues();
    templateSheet
        .getRange(2, 1, 1, 1)
        .setValues(note);

    var numberOfSheets = Math.ceil((srcSheet.getMaxColumns()-2)/2/2);
    for (var i = 0; i < numberOfSheets; i++) {
        var currentSheetName = srcSheet.getName() + " " + (i+1) + " of " + numberOfSheets;
        Logger.log(currentSheetName);
      
        var destSheet = templateSheet
            .copyTo(destDoc)
            .setName(currentSheetName); // duplicate template sheet and set name
        var data = srcSheet
            .getRange(1, 2 + 4*i, srcRowCount, 4)
            .getValues();

        destSheet
            .getRange(2, 2, srcRowCount, 4)
            .setValues(data);
        destSheet.getRange(1,1)
            .setValue(currentSheetName)
            .setFontSize(15).setFontWeight("bold");
    }
}
