// function that runs when call notes are submitted 
function submitForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var date = new Date();
  
  // store content
  var formID = sheet.getRange(1,1).getValue();
  var callNotes = sheet.getRange("B3:F3").getValue();
  var resolutionNotes = sheet.getRange("B7:F7").getValue();
  var caller = sheet.getRange("B11").getValue();
  var duration = sheet.getRange("B12").getValue();
  var callBack = sheet.getRange("B13").getValue();
  var ccID = sheet.getRange("I8").getValue();
  
  if (callBack == 'yes') {
    var callBackDate = new Date(sheet.getRange("B14").getValue());
  } else {
    var callBackDate = '';
  }
  
  // log content
  logData(sheet, date, caller, duration, callNotes, resolutionNotes, callBack, callBackDate);
  
  // update CRM information
  if (formID == 1) {
    var ccID = sheet.getRange("I12").getValue();
    updateLB(ccID, date, callBackDate);
  } else if (formID == 2) {
    var ccID = sheet.getRange("I12").getValue();
    updateS(ccID, date, callBackDate);
  } else if (formID == 3) {
    var ccID = sheet.getRange("I11").getValue();
    updateNA(ccID, date, callBackDate);
  }
  
  // reset the form for the next call
  resetForm(sheet);
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

function updateLB(ccID, date, callBackDate) {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0').getSheets()[0];
  
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

function updateS(ccID, date, callBackDate) {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0').getSheets()[1];
  
  // search for the client call id to find the row that contains the client's information
  var row;
  for (row = 2; row < sheet.getLastRow() + 1; row++) {
    if (ccID == sheet.getRange(row, 1).getValue()){
      break;
    }
  }
  
  // using this row, update the information
  sheet.getRange(row, 6).setValue(date);
  sheet.getRange(row, 7).setValue(callBackDate);
}

function updateNA(ccID, date, callBackDate) {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0').getSheets()[2];
  
  // search for the client call id to find the row that contains the client's information
  var row;
  for (row = 2; row < sheet.getLastRow() + 1; row++) {
    if (ccID == sheet.getRange(row, 1).getValue()){
      break;
    }
  }
  
  // using this row, update the information
  sheet.getRange(row, 4).setValue(date);
  sheet.getRange(row, 5).setValue(callBackDate);
}

function resetForm(sheet) {
  // clear content
  sheet.getRange("B3:B13").clearContent();
  sheet.getRange("B15:D15").clearContent();
  
  // reset the calendar button
  onClickNo();
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
