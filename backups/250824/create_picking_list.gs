function createPickingList() {
  console.log("createPickingList()");

  var originalFile = SpreadsheetApp.getActiveSpreadsheet(); //this sheetfile that should be used
  var filename = "Picking List " + Date();
  var newDocument=SpreadsheetApp.create(filename);
  var sheetsToRead=originalFile.getSheets();
  var noOfSheets=sheetsToRead.length;

  var totalSheet =newDocument.getSheets()[0];
  totalSheet.setName("Total orders");  

  totalSheet.insertColumnBefore(1);
  var targetRange=sheetsToRead[0].getRange("A:A").getDisplayValues();

  totalSheet.getRange(1,1,sheetsToRead[0].getMaxRows(),1).setValues(targetRange);   

  totalSheet.insertRowBefore(1); 
  totalSheet.insertRowBefore(1); 

  totalSheet.getRange("A3").setValue("");

  totalSheet.getRange("A1").setValue("Total orders");
  totalSheet.getRange("A1").setFontSize(15).setFontWeight("bold");
  setAlternatingColours(totalSheet); 

  totalSheet.insertColumnBefore(2); 
  
  
  for (var i =3;i<sheetsToRead[0].getMaxRows()+3;i++) {
    var cell = totalSheet.getRange(i,2,1,1);

    // This one must be updated when adding or removing groups! C[1] to C11] are groups 1 to eleven!, 
    //this can  be fixed with a string concatenation using number of sheets, but I dont have time now :(
    cell.setFormulaR1C1("=SUM(R[0]C[1]:R[0]C[11])");
  }

  totalSheet.deleteRows(sheetsToRead[0].getMaxRows()+7, totalSheet.getMaxRows()-(sheetsToRead[0].getMaxRows()+7)); 
  totalSheet.deleteColumns(noOfSheets+3,totalSheet.getMaxColumns()-(noOfSheets+3));

  totalSheet.getRange("B4").setValue("v-TOTAL-v");
  totalSheet.getRange("B4").setFontSize(15).setFontWeight("bold");
  
  for(var i=0;i<noOfSheets;i++) {
    sheetToRead=sheetsToRead[i];

    console.log("Behandlar ark: " + sheetsToRead[i].getName());

    totalSheet.getRange(1,totalSheet.getLastColumn()+1,sheetToRead.getMaxRows(),1).setValues(sheetToRead.getRange("B:B").getDisplayValues());
  }
}
