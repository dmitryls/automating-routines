function Report_Pushups() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var lastSheet = sheet.getSheets()[sheet.getNumSheets()-1];  
  var CurrentRow = 2;
  var PushupsRow = 9;
  var LastRow = SpreadsheetApp.getActiveSheet().getMaxRows();
  var CurrentColumn = 2;
  var LastColumn = SpreadsheetApp.getActiveSheet().getMaxColumns();
  var CurrentCell = "";
 
  //look for pushups row if it is not 9th
  if (String(lastSheet.getRange(PushupsRow,1).getValues()) != "Pushups") {
    while (CurrentRow <= LastRow) {
      CurrentRowName = String(lastSheet.getRange(CurrentRow,1).getValues());
      ++CurrentRow; 
      if (CurrentRowName == "Pushups") {
        PushupsRow = CurrentRow;
        break;
      }
    }  
  }
  
  //look for non-reported Done's, report them to Beeminder and change to Done*
  while (CurrentColumn <= LastColumn) {
    CurrentCell = String(lastSheet.getRange(PushupsRow,CurrentColumn).getValues());
    if (CurrentCell == "Done") {
      //report
      MailApp.sendEmail("bot@beeminder.com", "xxxxxx/pushups", "^ 1");
      //add * to mark it as already reported
      lastSheet.getRange(PushupsRow,CurrentColumn).setValue("Done*");   
    }
    ++CurrentColumn; 
  }
}
