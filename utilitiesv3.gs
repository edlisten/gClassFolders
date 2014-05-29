//Because we can't know whether a user is installing gClassFolders in a spreadsheet with other sheets
//we need a way to access the roster from the sheet's immutable ID
//This function is to be used whenever we need to get the roster sheet
// (note: Sheet ID and sheet index are different.  Index is mutable, id is not)
function getRosterSheet() { 
  var sheetId = PropertiesService.getDocumentProperties().getProperty('sheetID');
  if (this.SSKEY) {
    var ss = SpreadsheetApp.openById(this.SSKEY);
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheets = ss.getSheets();
  var tempSheetId = '';
  for (var i=0; i<sheets.length; i++) {
    tempSheetId = sheets[i].getSheetId();
    if (tempSheetId==sheetId) {
      return sheets[i];
    } 
  }
  return;
}

function getLogSheet() { 
  var sheetId = PropertiesService.getDocumentProperties().getProperty('logSheetID');
  if (this.SSKEY) {
    var ss = SpreadsheetApp.openById(this.SSKEY);
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheets = ss.getSheets();
  var tempSheetId = '';
  for (var i=0; i<sheets.length; i++) {
    tempSheetId = sheets[i].getSheetId();
    if (tempSheetId==sheetId) {
      return sheets[i];
    } 
  }
  return;
}

function getSendFileSheet() { 
  var sheetId = PropertiesService.getDocumentProperties().getProperty('sendFileSheetId');
  if (this.SSKEY) {
    var ss = SpreadsheetApp.openById(this.SSKEY);
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheets = ss.getSheets();
  var tempSheetId = '';
  for (var i=0; i<sheets.length; i++) {
    tempSheetId = sheets[i].getSheetId();
    if (tempSheetId==sheetId) {
      return sheets[i];
    } 
  }
  return;
}

//This will check to see if gCF is set up yet, if the person 
// put this code into the top of a function that should not run if the page has not been set up yet:
//if (setupGcfCheck() == false){ return;}
function setupGcfCheck(){
  var labels = PropertiesService.getDocumentProperties().getProperty('labels');
  var ui = SpreadsheetApp.getUi(); 
  var alreadyRan = false;
  if (!labels) {
    var result = ui.alert(
      'You have not yet set up the page for gClassFolders',
      'Would you like to do this now?',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    setupGCF();
    alreadyRan = false;
  } else { alreadyRan = false;}
  } else { alreadyRan = true;}
  
  return alreadyRan;
}


//Closes current panel
function closePanel(){
  var app = UiApp.getActiveApplication();
  app.close();
//  var sheet = getRosterSheet();
//  sheet.activate();
  return app;
}
////Closes current panel
//function closePanel2(){
//  var app = UiApp.getActiveApplication();
//  app.close();
////  var sheet = getRosterSheet();
////  sheet.activate();
//  return app;
//}



function returnIndices() {
  var sheet = getRosterSheet();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  
//  var sFnameIndex = headers.indexOf(defaultHeadings.sFname);
//  if (sFnameIndex==-1) {
//    badHeaders();
//    return;
//  }
//  var sLnameIndex = headers.indexOf(defaultHeadings.sLname);
//  if (sLnameIndex==-1) {
//    badHeaders();
//    return;
//  }
//  var sEmailIndex = headers.indexOf(defaultHeadings.sEmail);
//  if (sEmailIndex==-1) {
//    badHeaders();
//    return;
//  }
//  var clsNameIndex  = headers.indexOf(defaultHeadings.clsName);
//  if (clsNameIndex==-1) {
//    badHeaders();
//    return;
//  }
//  var clsPerIndex = headers.indexOf(defaultHeadings.clsPer);
//  if (clsNameIndex==-1) {
//    badHeaders();
//    return;
//  }
//  var tEmailIndex = headers.indexOf(defaultHeadings.tEmail);
//  if (tEmailIndex==-1) {
//    badHeaders();
//    return;
//  }  
  
  
  //refresh headers to pull in new folder id columns;
 // headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  //Get indices of folder Id columns
 
  try{
  var indices = new Object();
  indices.sFnameIndex = headers.indexOf(defaultHeadingsI.sFname);
  indices.sLnameIndex = headers.indexOf(defaultHeadingsI.sLname);
  indices.sEmailIndex = headers.indexOf(defaultHeadingsI.sEmail);
  indices.clsNameIndex = headers.indexOf(defaultHeadingsI.clsName);
  indices.clsPerIndex = headers.indexOf(defaultHeadingsI.clsPer);
  indices.tEmailIndex = headers.indexOf(defaultHeadingsI.tEmail);
  
  indices.sDropStatusIndex = headers.indexOf(defaultIDsI.status);
  indices.dbfIdIndex = headers.indexOf(defaultIDsI.assignmentFID);
  indices.crfIdIndex = headers.indexOf(defaultIDsI.classRootFID);
  indices.cvfIdIndex = headers.indexOf(defaultIDsI.classViewFID);
  indices.cefIdIndex = headers.indexOf(defaultIDsI.classEditFID);
  indices.rsfIdIndex = headers.indexOf(defaultIDsI.rootStudentFID);
  indices.tfIdIndex = headers.indexOf(defaultIDsI.teacherFID);
  indices.tShareStatusIndex = headers.indexOf(defaultIDsI.teacherShareStatus);
    
  var indicesString = JSON.stringify(indices);
  PropertiesService.getDocumentProperties().setProperty('indices', indicesString);
  
  }catch(err){
   Browser.msgBox("Error with indices " + err);
  }

 
  return indices;
} //end indices


function sendFileIndices() {
  var sheet = getSendFileSheet();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  
 
  try{
  var indices = new Object();
  indices.dateSent = headers.indexOf('dateSent');
  indices.template = headers.indexOf('template');
  indices.templateId = headers.indexOf('templateId'); 
  indices.permission = headers.indexOf('permission');
  indices.status = headers.indexOf('status');
  indices.copiedFileName = headers.indexOf('copiedFileName');  
  indices.copiedFileIds = headers.indexOf('copiedFileIds');
  indices.assigmentFolders = headers.indexOf('assigmentFolders');
    
    
//  var indicesString = JSON.stringify(indices);
//  PropertiesService.getDocumentProperties().setProperty('indices', indicesString);
  
  }catch(err){
   Browser.msgBox("Error with indices " + err);
  }

  return indices;
} //end indices







function getImage(imageID){
  //var imageID = "0Bwyqwd2fAHMMdXVzTURyU21QNXM"
  var theImage = "https://drive.google.com/uc?export=view&id=" + imageID;
  return theImage;
}




// used for picking a file
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
