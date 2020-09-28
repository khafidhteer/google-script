/*
Pseudo Code
1. Assign the active spreadsheet into a variable namely 'ss'
2. Assign the desired sheet of the active spreadsheet into a variable namely 'sheet1'
3. Copy sheet1 and assign to sheet2
4. Copy sheet2 content and paste value only into the same sheet (sheet2)
5. Make a new variable which its value is the name of desired attachment
6. If there is a sheet which has the name of desired attachment (old data from previous action) then delete it
7. Change the sheet2 name with the name of desired attachment
8. Set sheet2 as the active sheet
9. Get the spreadsheet and sheet2 ID to generate the sheet url
*/

function attachExcel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Sheet Name');
  //Logger.log(sheet1)
  sheet1.copyTo(ss);
  var sheet2 = ss.getSheetByName('Copy of Sheet Name');
  sheet2.getDataRange().copyTo(sheet2.getDataRange(), {contentsOnly:true});
  var name = 'Sheet Name (excel)';
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  sheet2.setName(name);
  ss.setActiveSheet(sheet2);
  var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();
  //Logger.log(ssID)
  var sh = SpreadsheetApp.getActive().getSheetByName(name)
  var shID = sh.getSheetId().toString()
  //Logger.log(shID)
  var date = Utilities.formatDate(new Date(new Date().getTime()-1*(24*3600*1000)), "GMT+07", "dd MMM yyyy");
  var recipients = "address@email.com";
  var subject = "Email Subject - "+date;
  var body = HtmlService.createHtmlOutputFromFile('index').getContent();
  var requestData = {"method": "GET", "headers":{"Authorization":"Bearer "+ScriptApp.getOAuthToken()}};
  var url = "https://docs.google.com/spreadsheets/d/"+ ssID + "/export?format=xlsx&id="+ssID+"&gid="+shID;
  var result = UrlFetchApp.fetch(url, requestData);
  var contents = result.getContent();
  var aname = name+date;
 
  MailApp.sendEmail(recipients, subject, body, {cc:"address2@email.com",name: "Sender Name", attachments:[{fileName:aname+".xlsx", content:contents, mimeType:"application//xlsx"}]});

  }
