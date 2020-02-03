// Khafidh Tri Ramdhani
// February 03 2020
// This script is made to copy the spreadsheet to the backup folder

function backupSpreadsheet() {

// generates the date when the spreadsheet is copied in variable DateOfCopying as dd MMM yyyy
var DateOfCopying = Utilities.formatDate(new Date(), "GMT+07", "dd MMM yyyy");

// gets the name of the original file follow by "copy" and when the spreadsheet was copied
// the name will be like this: Spreadsheet Name Copy 03 Feb 2020
var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " Copy " + DateOfCopying;

// gets the destination folder by their ID, change the *** with your folder ID (characters *** on this --
// -- link "https://drive.google.com/drive/u/0/folders/***"
  var destination = DriveApp.getFolderById("***");

// gets the current Google Sheet file
var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())

// makes copy of "file" with "name" at the "destination"
file.makeCopy(name, destination);
}
