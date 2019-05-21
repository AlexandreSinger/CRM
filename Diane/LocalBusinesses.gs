// opens a popup of the html 'page.html'
function openInputDialogLB() {
  var html = HtmlService
    .createHtmlOutputFromFile('PromptLBInfo')
    .setWidth(400)
    .setHeight(700);
  SpreadsheetApp.getUi()
       .showModalDialog(html, 'Add Local Business');
}

// when run from the html, the inputs from the form is used to add a client
function itemAddLB(form) {
  addLocalBusiness(form.sponsorName, form.firstName, form.lastName, form.phone, form.email, form.companyName, form.sponsorID, form.businessType, form.website, form.wasInvitationSent);
}

// function that creates a client based on parameters given
function addLocalBusiness(sponsorName, firstName, lastName, phone, email, companyName, sponsorID, businessType, website, wasInvitationSent) {
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0');
  
  var clientName = firstName + ' ' + lastName;
  
  // if the spreadsheet the client will be inputed is full, create a new one
  if(ss.getSheetByName('Storage').getRange("C2").getValue() == 200){
    addSpreadsheet();
  }
  
  // clone the Template and put it in spreadhseet, storing the url of the new sheet
  var clientCallID = ss.getSheetByName('Storage').getRange("G1").getValue() + 1;
  var link = cloneSheetLB(sponsorName, companyName, clientName, phone, email, sponsorID, clientCallID, businessType, website);
  
  // in the first page of CRM sheet, add the new client's information
  var sheet = ss.getSheets()[0];
  sheet.insertRowBefore(2);
  sheet.getRange("A2").setValue(clientCallID);
  sheet.getRange("B2").setValue('=HYPERLINK("'+link+'","'+companyName+'")');
  sheet.getRange("C2").setValue(clientName);
  sheet.getRange("D2").setValue(sponsorName);
  sheet.getRange("E2").setValue(sponsorID);
  sheet.getRange("F2").setValue(wasInvitationSent);
  ss.getSheetByName('Storage').getRange("G1").setValue(clientCallID);
}

function cloneSheetLB(sponsorName, companyName, clientName, phone, email, sponsorID, clientCallID, businessType, website) {
  // copy the template sheet from CRM and paste into new spreadsheet
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0');
  var ssDest = SpreadsheetApp.openByUrl(ss.getSheetByName('Storage').getRange("B2").getValue());
  var sheet = ss.getSheetByName('LB Template').copyTo(ssDest);
  
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
  sheet.getRange("I3").setValue(clientName);
  sheet.getRange("I4").setValue(phone);
  sheet.getRange("I5").setValue('=HYPERLINK("mailto:' + email + '","' + email + '")');
  sheet.getRange("I6").setValue(website);
  
  sheet.getRange("I8").setValue(businessType);
  
  sheet.getRange("I10").setValue(sponsorName);
  sheet.getRange("I11").setValue(sponsorID);
  sheet.getRange("I12").setValue(clientCallID);
  
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

// function that allows bulk numbers of clients to be added
function bulkAddLB() {
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1rrHvRSSwjj7RbT8M9OAAtdrw9gbcdantzom5uSYuOh0/edit#gid=0').getSheetByName('BulkAddLB');
  
  // for each row in the BulkAdd tab, add a client with the given specifications
  var row;
  var count = 0;
  for (row = 2; row < sheet.getLastRow() + 1; row++) {
    // if the company name is blank or is painted yellow (signify already added), then skip it
    if(!(sheet.getRange(row, 1).isBlank()) && sheet.getRange(row, 6).getBackground() != '#ffff00') {
      var companyName = sheet.getRange(row, 1).getValue();
      var firstName = sheet.getRange(row, 2).getValue();
      var lastName = sheet.getRange(row, 3).getValue();
      var phone = sheet.getRange(row, 4).getValue();
      var email = sheet.getRange(row, 5).getValue();
      var website = sheet.getRange(row, 6).getValue();
      var businessType = sheet.getRange(row, 7).getValue();
      var sponsorName = sheet.getRange(row, 8).getValue();
      var sponsorID = sheet.getRange(row, 9).getValue();
      var wasInvitationSent = sheet.getRange(row, 10).getValue();
      
      addLocalBusiness(sponsorName, firstName, lastName, phone, email, companyName, sponsorID, businessType, website, wasInvitationSent);
      
      // once a client has been added, paint the background yellow to signify that they have been added
      sheet.getRange(row, 1, 1, 10).setBackground('yellow');
    }
    
    // to prevent a bug that comes from a timeout error, the bulk add must stop after 100 iterations
    if(count >= 100){
      return;
    }
  }
}
