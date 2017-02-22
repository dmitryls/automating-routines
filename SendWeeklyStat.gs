function SendWeeklyStat() { 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var subject = "Weekly summary";
  var recipients = "xxxxx@gmail.com"
  var LastRow = SpreadsheetApp.getActiveSheet().getMaxRows();
  var LastRowDate = String(SpreadsheetApp.getActiveSheet().getRange(LastRow,2).getValues());
  
  //remove time after "at..." for the rows created by Trello
   if (LastRowDate.indexOf('at') !== -1) 
       {
         LastRowDate = LastRowDate.substr(0,LastRowDate.indexOf("at")-1);
       };    
  LastRowDate = new Date(LastRowDate);
  
  var CurrentRow = LastRow;
  var CurrentRowDate = LastRowDate;
  var weekago = new Date(new Date().getTime()-(7*24*60*60*1000));

 //look for the earliest row created a week ago  
  if (LastRowDate >= weekago)
  {
  while (CurrentRowDate >= weekago)
  {
    --CurrentRow;
    CurrentRowDate = sheet.getRange(CurrentRow,2).getValues();    
    
    //remove time (after "at..." for the rows created by Trello
    if (CurrentRowDate.indexOf("at") !== -1) 
       {CurrentRowDate = CurrentRowDate.substr(0,CurrentRowDate.indexOf("at")-1);
        //Logger.log("LastRowDate string cut after 'at' from Trello - "+LastRowDate)
       };   
    CurrentRowDate = new Date(CurrentRowDate);
  };

//create html and send it    
    var data = sheet.getRange('A'+CurrentRow+':E'+LastRow).getValues(); // define range here
    var message = '<HTML><BODY><table style=";border-collapse:collapse;" border = 1 cellpadding = 5><tr>';
    message += '<td>Due Date</td> <td>Completed at</td> <td>Task</td> <td>Note</td> <td>Folder</td> </tr><tr>';
    for (var row=0; row < data.length; ++row)
     {
      for(var col = 0;col < data[0].length; ++col)
      {
        message += '<td>'+data[row][col]+'</td>';
      }
      message += '</tr><tr>';
      }
      message += '</tr></table></body></HTML>';
      MailApp.sendEmail(recipients, subject, "", {htmlBody: message});
  } 
  else
  {
    MailApp.sendEmail(recipients, subject, 'Nothing was done this week!');
  }    
}