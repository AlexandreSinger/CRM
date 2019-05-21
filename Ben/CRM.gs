// This is the main code for the CRM tab
// Alexandre Singer
// May 2019

// adds a dropdown to the toolbar labelled 'function' to run specific functions
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Functions')
      .addItem('Sort CRM', 'sortCRM')
      .addItem('Add client', 'openInputDialog')
      .addToUi();
}

// opens a popup of the html 'page.html'
function openInputDialog() {
  var html = HtmlService
    .createHtmlOutputFromFile('PromptClientInfo')
    .setWidth(400)
    .setHeight(585);
  SpreadsheetApp.getUi()
       .showModalDialog(html, 'Add Client');
}

// when run from the html, the inputs from the form is used to add a client
function itemAdd(form) {
  addClient(form.companyName, form.sponsorID, form.contactName, form.phone, form.email, form.website, form.prospectLeadSold, form.priority);
}

// function that creates a client based on parameters given
function addClient(companyName, sponsorID, contactName, phone, email, website, prospectLeadSold, priority) {
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DOH9TvaNkQw-RtCw8vHqs1lDK3Wsrt8ZcoB2_5lj7EA/edit#gid=0');
  
  // if the spreadsheet the client will be inputed is full, create a new one
  if(ss.getSheetByName('Storage').getRange("C2").getValue() == 200){
    addSpreadsheet();
  }
  
  // clone the Template and put it in spreadhseet, storing the url of the new sheet
  var clientCallID = ss.getSheetByName('Storage').getRange("G1").getValue() + 1;
  var link = cloneSheet(companyName, sponsorID, clientCallID, contactName, phone, email, website);
  
  // in the first page of CRM sheet, add the new client's information
  var sheet = ss.getSheets()[0];
  sheet.insertRowBefore(2);
  sheet.getRange("A2").setValue(clientCallID);
  sheet.getRange("B2").setValue('=HYPERLINK("'+link+'","'+companyName+'")');
  sheet.getRange("C2").setValue(sponsorID);
  sheet.getRange("D2").setValue(contactName);
  sheet.getRange("E2").setValue(prospectLeadSold);
  sheet.getRange("F2").setValue(priority);
  ss.getSheetByName('Storage').getRange("G1").setValue(clientCallID);
}

function cloneSheet(companyName, sponsorID, clientCallID, contactName, phone, email, website) {
  // copy the template sheet from CRM and paste into new spreadsheet
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DOH9TvaNkQw-RtCw8vHqs1lDK3Wsrt8ZcoB2_5lj7EA/edit#gid=0');
  var ssDest = SpreadsheetApp.openByUrl(ss.getSheetByName('Storage').getRange("B2").getValue());
  var sheet = ss.getSheetByName('Template').copyTo(ssDest);
  
  // name the copied sheet. if the client already exists in the spreadsheet, change the name of the new sheet
  var sheetName = companyName;
  var old = ssDest.getSheetByName(companyName);
  var i = 0;
  while (old) {
    i = i + 1;
    sheetName = companyName + ' ' + i;
    old = ssDest.getSheetByName(sheetName);
  }
  sheet.setName(sheetName);
  
  // fill in the template with the client's information
  sheet.getRange("B2").setValue(companyName);
  
  sheet.getRange("I3").setValue(contactName);
  sheet.getRange("I4").setValue(phone);
  sheet.getRange("I5").setValue('=HYPERLINK("mailto:' + email + '","' + email + '")');
  sheet.getRange("I6").setValue(website);
  
  sheet.getRange("I8").setValue(sponsorID);
  sheet.getRange("I9").setValue(clientCallID);
  
  // create a snapshot url from the sponsor ID if the sponsor ID exists
  if(sponsorID){
    var snapshot = 'http://admin.applica-solutions.com/Admin/SponsorInfo?SId=' + sponsorID + '&cHdkPWdyb3d0aDEyMw==';
  } else {
    var snapshot = sponsorID;
  }
  
  sheet.getRange("I11").setValue(snapshot);
  
  // store the location of the amount of sheets in the spreadsheet from the storage tab of the CRM
  var amountLocation = ss.getSheetByName('Storage').getRange("C2");
  
  // if there was nothing in the sheet originally, because the template had to be left there when the spreadsheet was created it is then deleted to save space
  if (amountLocation.getValue() == 0){
    ssDest.deleteSheet(ssDest.getSheets()[0]);
  }
  
  // increment the amount of sheets shown in the CRM
  amountLocation.setValue(amountLocation.getValue()+1);
  
  // return a url to that sheet
  return ssDest.getUrl() +'#gid='+ sheet.getSheetId();
}

// function that adds a spreadsheet for storing the clients
function addSpreadsheet() {
  var name = new Date();
  
  // get the CRM spreadsheet and copy it
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DOH9TvaNkQw-RtCw8vHqs1lDK3Wsrt8ZcoB2_5lj7EA/edit#gid=0');
  var newSS = ss.copy("CRM Clients: " + name);
  
  // get the file of the new spreadsheet and remove it from its current location (drive)
  var file = DriveApp.getFileById(newSS.getId());
  file.getParents().next().removeFile(file);
  
  // get the folder that contains the active spreadsheet and copy the file into it
  var folderID = DriveApp.getFileById(ss.getId()).getParents().next().getId();
  DriveApp.getFolderById(folderID).addFile(file);
  
  // delete the CRM sheet and the Storage sheets leaving only the template behind
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
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DOH9TvaNkQw-RtCw8vHqs1lDK3Wsrt8ZcoB2_5lj7EA/edit#gid=0').getSheets()[0];
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
      pp += 10;
      
    // else if the date is yet to pass, hide the row
    } else if(callBackDate.valueOf() > currentDate){
      sheet.hideRows(row);
      
    // else if the cleint has already been called, move to bottom of the list
    } else if(lastCall.valueOf()){
      pp -= 10;
    }
    
    // set the priority point of the client in the CRM
    sheet.getRange(row, 10).setValue(pp);
  }
  
  // sort by the priority first, then by call back date, then by the priority points
  sheet.sort(6, false);
  sheet.sort(8, true);
  sheet.sort(9, false);
}

// function that allows bulk numbers of clients to be added
function bulkAdd() {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DOH9TvaNkQw-RtCw8vHqs1lDK3Wsrt8ZcoB2_5lj7EA/edit#gid=0').getSheetByName('BulkAdd');
  
  // for each row in the BulkAdd tab, add a client with the given specifications
  var row;
  var count = 0;
  for (row = 2; row < sheet.getLastRow() + 1; row++) {
    // if the client name is blank or is painted yellow (signify already added), then skip it
    if(!(sheet.getRange(row, 1).isBlank()) && sheet.getRange(row, 1).getBackground() != '#ffff00') {
      var companyName = sheet.getRange(row, 1).getValue();
      var sponsorID = sheet.getRange(row, 2).getValue();
      var contactName = sheet.getRange(row, 3).getValue();
      var phone = sheet.getRange(row, 4).getValue();
      var email = sheet.getRange(row, 5).getValue();
      var website = sheet.getRange(row, 6).getValue();
      var prospectLeadSold = sheet.getRange(row, 7).getValue();
      var priority = sheet.getRange(row, 8).getValue();
      
      addClient(companyName, sponsorID, contactName, phone, email, website, prospectLeadSold, priority)
      
      // once a client has been added, paint the background yellow to signify that they have been added
      sheet.getRange(row, 1, 1, 8).setBackground('yellow');
      count++;
    }
    // to prevent a bug that comes from a timeout error, the bulk add must stop after 100 iterations
    if(count >= 100){
      return;
    }
  }
}
