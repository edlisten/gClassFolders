//remove all instances of defaultLabels.class


var GCLASSICONURL = 'https://sites.google.com/site/gclassfolders/_/rsrc/1360538205205/config/customLogo.gif?revision=1';
var GCLASSLAUNCHERICONURL = 'https://sites.google.com/site/gclassfolders/_/rsrc/1360538205205/config/customLogo.gif?revision=1';
var userEmail = Session.getEffectiveUser().getEmail();
var SSKEY = PropertiesService.getDocumentProperties().getProperty('ssKey');

// This list was taken from the list of available languages in Google Translate service, responsible for our UI internationalization.
var googleLangList = ['English: en'];

// Default hedings and labels
var defaultLabels = {dropBoxes: "Assignment Folders", dropBox: "Assignment Folder", period: "Period", edit: "Edit", view: "View", teacher: "Teacher", saved: "false"};
var defaultHeadings ={sFname: "Student First Name", sLname: "Student Last Name", sEmail: "Student Email", clsName: "Class Name", clsPer: "Period ~Optional~", tEmail: "Teacher Email(s)"}; 
var defaultHeadingsI ={sFname: "sFname", sLname: "sLname", sEmail: "sEmail", clsName: "clsName", clsPer: "clsPer", tEmail: "tEmail"};
var defaultIDs ={status: "Status", assignmentFID: "Assignment Folder ID",classRootFID: "Class Root Folder ID", classViewFID: "Class View Folder ID", classEditFID: "Class Edit Folder ID", rootStudentFID: "Root Student Folder ID",teacherFID: "Teacher Folder ID", teacherShareStatus: "Teacher Status" };
var defaultIDsI ={status: "status", assignmentFID: "assignmentFID",classRootFID: "classRootFID", classViewFID: "classViewFID", classEditFID: "classEditFID", rootStudentFID: "rootStudentFID",teacherFID: "teacherFID", teacherShareStatus: "tStatus" };





var labels = function(){
  var properties2 = PropertiesService.getDocumentProperties().getProperty('labels');
  var labels = JSON.parse(properties2);
 // CacheService.getPrivateCache().put('labels', labels, 660);
  return labels;
  }



function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var addOn = ui.createAddonMenu();
  addOn.addItem("Setup Sheet", "setupGCF");
  addOn.addItem("Create Folders", "createClassFolders");
  addOn.addSeparator();
  addOn.addItem("Send File", "sendFile")
  addOn.addSubMenu(ui.createMenu("Tools")
          .addItem("Manage Students", "actionsSidebar")         
          .addItem('Rename Labels', 'renameFolderLabels')
          .addItem("Start Over", "startOver"))
  addOn.addSeparator();
  addOn.addItem("Donate", "donate");
  addOn.addItem("Get Help", "gethelp");
  addOn.addToUi();

  
} //end onOpen





//Creates the RosterSheet
//Sets the default labels into Properties & Cache
//Sets property 'alreadyRan' to false
function setupGCF(){
 
 var documentProperties = PropertiesService.getDocumentProperties();
 var properties = documentProperties.getProperties(); 
 var ui = SpreadsheetApp.getUi();

 // Only run if never ran before 
 if (!properties.alreadyRan){
//   if (properties.alreadyRan){

 //Create Folders & Shares has not yet been ran
 documentProperties.setProperty('alreadyRan', 'false');
  
   
 // push Default Labels to Properties and Private Cache 
 var labels = JSON.stringify(defaultLabels);  
 documentProperties.setProperty('labels', labels);
// CacheService.getPrivateCache().put('labels', labels, 660);
 
 // Set Default headings into Properties
 properties.headings = JSON.stringify(defaultHeadings);
  
// Creates the Roster Sheet
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var activeSheet = ss.getActiveSheet().getName(); 

  
 if (activeSheet == 'Sheet1'){
    ss.renameActiveSheet('gClassRoster');
//    ss.insertSheet('Log');
  }
  
   var sheet = ss.getSheetByName('gClassRoster');
  if ((!sheet)&&(!properties.sheetId)) {
    sheet = ss.insertSheet('gClassRoster');  
  }
//   var logSheet = ss.getSheetByName('Log');
//   if (!logSheet) {ss.insertSheet('Log');}
   
  
  var sheetId = sheet.getSheetId();   
  documentProperties.setProperty('sheetID', sheetId); 
//  var logSheetId = logSheet.getSheetId();
//  documentProperties.setProperty('logSheetID', logSheetId); 
  
//  logSheet.getRange('A1').setValue('Identifier');
//  logSheet.getRange('B1').setValue('Information');
//  logSheet.setFrozenRows(1);
//  SpreadsheetApp.flush();
//  sheet.activate();
//  SpreadsheetApp.flush();

 // This inserts five rows before the first row
 sheet.insertRowsBefore(1, 2);
 SpreadsheetApp.flush();
  sheet.getRange("A1").setValue(defaultHeadings.sFname);
  sheet.getRange("B1").setValue(defaultHeadings.sLname);
  sheet.getRange("C1").setValue(defaultHeadings.sEmail); 
  sheet.getRange("D1").setValue(defaultHeadings.clsName);
  sheet.getRange("E1").setValue(defaultHeadings.clsPer);
  sheet.getRange("F1").setValue(defaultHeadings.tEmail);
   
  sheet.getRange("A2").setValue(defaultHeadingsI.sFname);
  sheet.getRange("B2").setValue(defaultHeadingsI.sLname);
  sheet.getRange("C2").setValue(defaultHeadingsI.sEmail); 
  sheet.getRange("D2").setValue(defaultHeadingsI.clsName);
  sheet.getRange("E2").setValue(defaultHeadingsI.clsPer);
  sheet.getRange("F2").setValue(defaultHeadingsI.tEmail);
   
  sheet.setColumnWidth(1,130); 
  sheet.setColumnWidth(2,130); 
  sheet.setColumnWidth(3,160); 
  sheet.setColumnWidth(4,130);
  sheet.setColumnWidth(5,130); 
  sheet.setColumnWidth(6,160); 
  sheet.setColumnWidth(7,220);   
  sheet.setFrozenRows(2);

  sheet.hideRows(2);
   SpreadsheetApp.flush();
 } else { // end if ran before
  Browser.msgBox("gClassFolders has already been setup for this sheet");
 }
 
  initialDialoge();

  
}// End setupGCF()



//resets the spreadsheet to original. This will include onOpen() once that is finished testing.
function startOver(){ 

//Check if the person really wants to do this
    var ui = SpreadsheetApp.getUi(); 
    var result = ui.alert(
      'WARNING!! This will delete all your data in the spreadsheet!!!',
      'Running this will not effect the folders that have been created previously, you will need to delete those manually. \n Running this will allow you to create a new folder set with the same class name. If you renamed any labels, you will need to reset them. \n\n Would you like to do this now?',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.NO) {
      return;
  } 
  
  PropertiesService.getDocumentProperties().deleteAllProperties();
  CacheService.getPrivateCache().remove('labels');
  var ss = SpreadsheetApp.getActive();
  var roster = ss.getSheetByName('gClassRoster');
  var prop = ss.getSheetByName('Properties');
  var log = ss.getSheetByName('Log');
  var sendFile = ss.getSheetByName('Send File Log');
  
  Logger.getLog();
  if (roster != null){
    ss.insertSheet('Sheet1');
    ss.deleteSheet(roster);}
  if (prop != null){ss.deleteSheet(prop);}
  if (log != null){ss.deleteSheet(log);}
  if (sendFile != null){ss.deleteSheet(sendFile);}
  onOpen();
  
  
   var result = ui.alert(
      'Would you like to run \"Setup Sheet\" for gClassFolders at this time?',
      'Clicking Yes will create the headers used by the script',
  ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
      setupGCF();
  } 
  
  
}





function initialDialoge() {
  var title = "Welcome to gClassFolders";
  var linkTutorial = "https://docs.google.com/presentation/d/1ebm9rbySV0ukevQhENin41ioXpUO-JzXTcvltXnjE_A/embed?start=false&loop=false&delayms=3000";
  var linkCommunity = "https://plus.google.com/communities/115718335045383669895";
  var linkAuthor = "http://www.edlisten.com";
  var linkWebsite = "http://www.gclassfolders.com";
  var linkDonate = "http://www.gclassfolders.com/donate"
  var descriptionText = "Thank you for your interest in gClassFolders. You may click the tutorial link below the image or follow the buttons to either start entering in data or change the default labels.";
  
  var doc = SpreadsheetApp.getActive();
  
  var app = UiApp.createApplication().setHeight(225).setWidth(400).setTitle(title);
  var panel = app.createVerticalPanel();
  var bodyGrid = app.createGrid(1,2);
  var buttonGrid = app.createGrid(1,2).setStyleAttribute('verticalAlign', 'top');
  var vPanel2 = app.createVerticalPanel().setStyleAttribute('verticalAlign', 'top');
  
  // Create Links
  var dText = app.createLabel(descriptionText).setStyleAttribute('paddingBottom', 10);
  var image = app.createImage(getImage('0Bwyqwd2fAHMMWGhOaThDYXgwNjQ')).setStyleAttribute('paddingRight', 10);
  var anchorTutorial = app.createAnchor("Tutorial", linkTutorial).setId("anchorTutorial");
  var anchorCommunity = app.createAnchor("G+ Community", linkCommunity).setId("linkCommunity");
  var anchorAuthor = app.createAnchor("Author's Blog", linkAuthor).setId("linkAuthor");
  var anchorWebsite = app.createAnchor("gClassFolder's Website", linkWebsite).setId("linkWebsite");
  var anchorDonate = app.createAnchor("Donate", linkDonate).setId("linkDonate");
  
  //Add Links to panel
  vPanel2.add(anchorTutorial);
  vPanel2.add(anchorCommunity);
  vPanel2.add(anchorAuthor);
  vPanel2.add(anchorWebsite);
  vPanel2.add(anchorDonate);
 
  //Select Label Names
  var labelHandler = app.createServerClickHandler('renameFolderLabels');
  var labelButton = app.createButton('Change Label Names').setStyleAttribute('marginTop', 10).addClickHandler(labelHandler);
  
  //Close Button
  var closeHandler = app.createServerClickHandler('closePanel');
  var closeButton = app.createButton('Go To Spreadsheet').setStyleAttribute('marginTop', 10).addClickHandler(closeHandler);
  
  bodyGrid
  .setWidget(0,0, image)
  .setWidget(0,1, vPanel2);
  
  buttonGrid
    .setWidget(0,0, closeButton)
    .setWidget(0,1, labelButton);

 //Build panel layout 
  panel.add(dText);
  panel.add(bodyGrid);
  panel.add(buttonGrid);
  app.add(panel); 
  doc.show(app);
}



function gethelp() {
  var title = "Get Help Using gClassFolders";
  var linkTutorial = "https://docs.google.com/presentation/d/1ebm9rbySV0ukevQhENin41ioXpUO-JzXTcvltXnjE_A/embed?start=false&loop=false&delayms=3000";
  var linkCommunity = "https://plus.google.com/communities/115718335045383669895";
  var linkAuthor = "http://www.edlisten.com";
  var linkWebsite = "http://www.gclassfolders.com";
  var studentInstructions = "https://docs.google.com/document/d/1RxvGrBhaLMIdE6T0VgTei6HUTExrt-_Ycy-cCDuG53A/edit?usp=sharing";
  var descriptionText = "Help is only a click away.";
  
  var doc = SpreadsheetApp.getActive();
  
  var app = UiApp.createApplication().setHeight(225).setWidth(400).setTitle(title);
  var panel = app.createVerticalPanel();
  var bodyGrid = app.createGrid(1,2);
  var buttonGrid = app.createGrid(1,2).setStyleAttribute('verticalAlign', 'top');
  var vPanel2 = app.createVerticalPanel().setStyleAttribute('verticalAlign', 'top');
  
  var dText = app.createLabel(descriptionText).setStyleAttribute('paddingBottom', 10);
  var image = app.createImage(getImage('0B4GLYStYeHYkS1RHUEVmTG1uMHM')).setStyleAttribute('paddingRight', 10);
  
  // Create Links
  var anchorStudent = app.createAnchor("Student Instructions", studentInstructions).setId("studentInstructions");
  var anchorTutorial = app.createAnchor("Tutorial", linkTutorial).setId("anchorTutorial");
  var anchorCommunity = app.createAnchor("G+ Community", linkCommunity).setId("linkCommunity");
  var anchorAuthor = app.createAnchor("Author's Blog", linkAuthor).setId("linkAuthor");
  var anchorWebsite = app.createAnchor("gClassFolder's Website", linkWebsite).setId("linkWebsite");
  
  //Add Links to panel
  vPanel2.add(anchorTutorial);
  vPanel2.add(anchorCommunity);
  vPanel2.add(anchorStudent);
  vPanel2.add(anchorAuthor);
  vPanel2.add(anchorWebsite);
 

  //Close Button
  var closeHandler = app.createServerClickHandler('closePanel');
  var closeButton = app.createButton('Go To Spreadsheet').setStyleAttribute('marginTop', 10).addClickHandler(closeHandler);
  
  bodyGrid
  .setWidget(0,0, image)
  .setWidget(0,1, vPanel2);
  
  buttonGrid
    .setWidget(0,0, closeButton);
   // .setWidget(0,1, labelButton);

 //Build panel layout 
  panel.add(dText);
  panel.add(bodyGrid);
  panel.add(buttonGrid);
  app.add(panel); 
  doc.show(app);
}

function donate() {
  var description = 'If you have found gClassFolders helpful please consider donating to the project.  This will help me continue to support it in the future.'; 
  
  var app = UiApp.createApplication().setHeight(100).setWidth(400).setTitle('Donate');
  var panel = app.createVerticalPanel();
  var descriptionLabel = app.createLabel(description).setStyleAttribute('marginBottom', '15px');

  var link = app.createAnchor("Go To The Donate Page", "http://www.gclassfolders.com/donate");
  
  panel.add(descriptionLabel);

  panel.add(link);
  app.add(panel);
  SpreadsheetApp.getActive().show(app);
  
}


//*********************  Used for debugging and testing only *****************************


function documentPropertiesToLog(){
var properties = PropertiesService.getDocumentProperties().getProperties();
Logger.log(properties);
} //end documentPropertiesToLog

function privateCacheToLog(){
 // var theCache = this.labels().class;
  var theCache = this.labels();
  //var dropBoxLabel = this.labels().dropBox;
  //var dropBoxLabels = labels().dropBoxes;
  //var classLabel = labels().class;
  //var classesLabel = labels().classes;
  //var periodLabel = labels().period;
  Logger.log(theCache);
}

