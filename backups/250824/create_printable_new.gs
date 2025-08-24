function createPrintableSheetsNew() {
    console.log("createPrintableSheetsNew()");

    var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var destDoc = SpreadsheetApp.create("Printable_sheets_NEW_" + Date());

    createTemplate(destDoc, srcSheets[0]);
    templateSheet = destDoc.getSheets()[1];

    // prepare totals sheet
    totalSheet = destDoc.getSheets()[0];
    totalSheet
        .getRange( 4, totalSheet.getLastColumn() + 1, srcSheets[0].getMaxRows() - 3, 1 )
        .setValues( srcSheets[0].getRange("A4:A").getDisplayValues() )
        .setFontSize(15).setFontWeight("bold");

    for(var i = 0; i < srcSheets.length; i++) {
        console.log("Behandlar ark: " + srcSheets[i].getName());
        updateSharedTotalsPage(totalSheet, srcSheets[i]); 
        createGroupTotalsPage(destDoc, srcSheets[i]);
        createMemberPages(destDoc, srcSheets[i], templateSheet); 
    }

    destDoc.deleteSheet(templateSheet);
}

function createPickingListNew() {
    console.log("createPickingListNew()");

    var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var destDoc = SpreadsheetApp.create("Picking_list_NEW_" + Date());

    createTemplate(destDoc, srcSheets[0]);
    templateSheet = destDoc.getSheets()[1];

    // prepare totals sheet
    totalSheet = destDoc.getSheets()[0];
    totalSheet
        .getRange( 4, totalSheet.getLastColumn() + 1, srcSheets[0].getMaxRows() - 3, 1 )
        .setValues( srcSheets[0].getRange("A4:A").getDisplayValues() )
        .setFontSize(15).setFontWeight("bold");

    for(var i = 0; i < srcSheets.length; i++) {
        console.log("Behandlar ark: " + srcSheets[i].getName());
        updateSharedTotalsPage(totalSheet, srcSheets[i]);

        // no group pages, only the shared totals page == picking list
        //createGroupTotalsPage(destDoc, srcSheets[i]);
        //createMemberPages(destDoc, srcSheets[i], templateSheet); 
    }

    destDoc.deleteSheet(templateSheet);
}

function updateSharedTotalsPage(totalSheet, srcSheet) {

    // make space for the "total-totals" at column 2
    total_cols = totalSheet.getLastColumn();
    col = total_cols + 2;
    if (total_cols > 2) {
        col = total_cols + 1;
    }

    totalSheet
      .getRange( 1, col, srcSheet.getMaxRows(), 1 )
      .setValues( srcSheet.getRange("B:B").getDisplayValues() )
      .setFontSize(15).setFontWeight("bold");

    totalSheet
      .getRange( 2, col, 1, 1)
      .setValue( srcSheet.getName() )
      .setFontSize(15).setFontWeight("bold");
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
        .setValues(data)
        .setFontSize(15).setFontWeight("bold");
}

function createMemberPages(destDoc, srcSheet, templateSheet) {
    var srcRowCount = srcSheet.getMaxRows();
    
    // NOTE: this field 1,1 is different for each group, with contact info, etc.
    var note = srcSheet
        .getRange(1, 1, 1, 1)
        .getValues();
    templateSheet
        .getRange(2, 1, 1, 1)
        .setValues(note)
        .setFontSize(15).setFontWeight("bold");

    var numberOfSheets = Math.ceil((srcSheet.getMaxColumns()-2)/2/2);
    for (var i = 0; i < numberOfSheets; i++) {
        var currentSheetName = srcSheet.getName() + " " + (i+1) + " of " + numberOfSheets;
        Logger.log(currentSheetName);
      
        var destSheet = templateSheet
            .copyTo(destDoc)
            .setName(currentSheetName); // duplicate template sheet and set name
        var data = srcSheet
            .getRange(1, 3 + 4*i, srcRowCount, 4)
            .getValues();

        destSheet
            .getRange(2, 2, srcRowCount, 4)
            .setValues(data)
            .setFontSize(15).setFontWeight("bold");
        destSheet.getRange(1,1)
            .setValue(currentSheetName)
            .setFontSize(15).setFontWeight("bold");
    }
}
