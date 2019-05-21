// These are scripts associated with each individual client call sheet
// Alexandre Singer
// May 2019

// function that runs when call notes are submitted 
function submitForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var date = new Date();
  
  // store content
  var callNotes = sheet.getRange("B3:F3").getValue();
  var resolutionNotes = sheet.getRange("B7:F7").getValue();
  var caller = sheet.getRange("B11").getValue();
  var duration = sheet.getRange("B12").getValue();
  var callBack = sheet.getRange("B13").getValue();
  var ccID = sheet.getRange("I9").getValue();
  
  if (callBack == 'yes') {
    var callBackDate = new Date(sheet.getRange("B14").getValue());
  } else {
    var callBackDate = '';
  }
  
  // log content
  logData(sheet, date, caller, duration, callNotes, resolutionNotes, callBack, callBackDate);
  
  // update CRM information
  updateCRM(ccID, date, callBackDate);
  
  // reset the form for the next call
  resetForm(sheet);
}

function updateCRM(ccID, date, callBackDate) {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DOH9TvaNkQw-RtCw8vHqs1lDK3Wsrt8ZcoB2_5lj7EA/edit#gid=0').getSheets()[0];
  
  // search for the client call id to find the row that contains the client's information
  var row;
  for (row = 2; row < sheet.getLastRow() + 1; row++) {
    if (ccID == sheet.getRange(row, 1).getValue()){
      break;
    }
  }
  
  // using this row, update the information
  sheet.getRange(row, 7).setValue(date);
  sheet.getRange(row, 8).setValue(callBackDate);
}

function resetForm(sheet) {
  // clear content
  sheet.getRange("B3:B13").clearContent();
  sheet.getRange("B15:D15").clearContent();
  
  // reset the calendar button
  onClickNo();
}

// function that logs the data into the call log in the spreadsheet
function logData(sheet, date, caller, duration, callNotes, resolutionNotes, callBack, callBackDate) {
  sheet.insertRowBefore(20);
  sheet.getRange("A20").setValue(date);
  sheet.getRange("B20").setValue(caller);
  sheet.getRange("C20").setValue(duration);
  sheet.getRange("D20:H20").setValue(callNotes);
  sheet.getRange("D20:H20").merge();
  sheet.getRange("I20:M20").setValue(resolutionNotes);
  sheet.getRange("I20:M20").merge();
  sheet.getRange("N20").setValue(callBack);
  sheet.getRange("O20").setValue(callBackDate);
}

// function that populates cell B13 with the value 'yes' when run
function onClickYes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // update information in cell B13 to yes
  sheet.getRange("B13").setValue('yes');
  
  // allow the use of the calendar
  sheet.getRange("B14").setBackground(null);
  sheet.getRange("B14").setValue("select date");
}

// function that populates cell B13 with the value 'no' when run
function onClickNo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // update information in cell B13 to no
  sheet.getRange("B13").setValue('no');
  
  // dissallow the use of the calendar
  sheet.getRange("B14").setBackground('Grey');
  sheet.getRange("B14").setValue("call back not selected");
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
  var folderID = '1SzMDzIkKNr4OXTqoVnjCH9yf38_rZBps';
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
