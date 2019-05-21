// adds a dropdown to the toolbar labelled 'function' to run specific functions
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Functions')
      //.addItem('Sort CRM', 'sortCRM')
      .addItem('Add Local Business', 'openInputDialogLB')
      .addItem('Add Sponsor', 'openInputDialogS')
      .addToUi();
}

// function that adds a spreadsheet for storing the clients
function addSpreadsheet() {
  var name = new Date();
  
  // get the CRM spreadsheet and copy it
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0');
  var newSS = ss.copy("CRM-D Clients: " + name);
  
  // get the file of the new spreadsheet and remove it from its current location (drive)
  var file = DriveApp.getFileById(newSS.getId());
  file.getParents().next().removeFile(file);
  
  // get the folder that contains the active spreadsheet and copy the file into it
  var folderID = DriveApp.getFileById(ss.getId()).getParents().next().getId();
  DriveApp.getFolderById(folderID).addFile(file);
  
  // delete the CRM sheet and the Storage sheets leaving only the template behind
  newSS.deleteSheet(newSS.getSheets()[0]);
  newSS.deleteSheet(newSS.getSheets()[0]);
  newSS.deleteSheet(newSS.getSheets()[0]);
  newSS.deleteSheet(newSS.getSheets()[0]);
  newSS.deleteSheet(newSS.getSheets()[0]);
  newSS.deleteSheet(newSS.getSheets()[0]);
  newSS.deleteSheet(newSS.getSheets()[0]);
  
  newSS.deleteSheet(newSS.getSheets()[1]);
  newSS.deleteSheet(newSS.getSheets()[1]);
  
  // copy the url of the new spreadsheet and store the storage sheet
  var link = newSS.getUrl();
  var sheet = ss.getSheetByName('Storage');
  
  // insert information into the top of the storage sheet list
  sheet.insertRowBefore(2);
  sheet.getRange("A2").setValue('=HYPERLINK("'+link+'","'+name+'")');
  sheet.getRange("B2").setValue(link);
  sheet.getRange("C2").setValue(0);
  sheet.getRange("D2").setValue(200);
}

// function that sorts the CRM
function sortCRM() {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0').getSheets()[0];
  var currentDate = new Date();
  
  // show all rows
  sheet.showRows(2, sheet.getLastRow() - 1);
  
  // check each row of the CRM
  var row;
  for (row = 2; row < sheet.getLastRow() + 1; row++) {
    var callBackDate = new Date(sheet.getRange(row, 8).getValue());
    var lastCall = new Date(sheet.getRange(row, 7).getValue());
    var pp = 0;
    
    // if the call back date has passed, add 10 to the priority point
    if(callBackDate.valueOf() < currentDate){
      sheet.getRange(row, 9).setValue(pp + 10);
      
    // else if the date is yet to pass or the client has already been called, hide the row
    } else if(callBackDate.valueOf() > currentDate || lastCall.valueOf()){
      sheet.hideRows(row);
    }  
  }
  
  // sort by the priority first, then by call back date, then by the priority points
  sheet.sort(5, false);
  sheet.sort(8, true);
  sheet.sort(9, false);
}

// function to add a folder to store client files
function addFolder() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // if already given a folder location, print error message and quit the function
  if (sheet.getRange("M3").getValue() != '<folder not created yet>') {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error: Folder already created for this client');
    return;
  }
  
  // get the folder that contains the active spreadsheet and make a folder with the name given in B2
  var folderName = sheet.getRange("B2").getValue();
  //var folderID = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next().getId();
  var folderID = '1nG5fcunNVR4ZyZO8Khloqydl-7NbELTL';
  var newFolder = DriveApp.getFolderById(folderID).createFolder(folderName);
  
  // print the url of the new folder into cell M3
  sheet.getRange("M3").setValue(newFolder.getUrl());
}

//function to remove a folder created to store client files
function removeFolder() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // if there is no folder to delete, print error and exit function
  if (sheet.getRange("M3").getValue() == '<folder not created yet>') {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error: No Folder to delete');
    return;
  }
  
  // get the folder ID by taking the url and removing the non ID part, then move it into the trash
  var folderID = sheet.getRange("M3").getValue().replace('https://drive.google.com/drive/folders/', '');
  DriveApp.getFolderById(folderID).setTrashed(true);
  
  // when done, set cell M3 with the standard value
  sheet.getRange("M3").setValue('<folder not created yet>');
}

// Function that searches through the Local Business sheet, finds duplicate businesses and removes them, populating the information into the other dupe
function removeDuplicateLB() {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0').getSheets()[0];
  
  sheet.sort(1)
  
}
