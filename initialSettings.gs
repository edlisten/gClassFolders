


//Builds the UI for initial settings
function gClassFolders_folderLabels() {
  var app = UiApp.createApplication().setHeight(300);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var title = app.createLabel("Choose your folder labels").setId('title').setStyleAttribute('fontSize', '18px').setStyleAttribute('marginBottom', '8px');
  var documentProperties = PropertiesService.getDocumentProperties();
  var properties = documentProperties.getProperties();
  var initialLock = properties.alreadyRan;
  var defaultLabels = {dropBox: "Assignment Folder", dropBoxes: "Assignment Folders", class: "Class", classes: "Classes", period: "Period"};
  
  if (!properties.labels){
  properties.labels = Utilities.jsonStringify(defaultLabels);
  var labels = properties.labels;
  CacheService.getPrivateCache().put('labels', labels, 660);
  }
  
//  var labelProperties = JSON.parse(properties.labels);
//    var dropBoxLabel = labelProperties.dropBox;
//    var dropBoxLabels = labelProperties.dropBoxes;
//    var classLabel = labelProperties.class;
//    var classesLabel = labelProperties.classes;
//    var periodLabel = labelProperties.period;


//Get labes from Cache  
    var dropBoxLabel = this.labels().dropBox;
    var dropBoxLabels = this.labels().dropBoxes;
    var classLabel = this.labels().class;
    var classesLabel = this.labels().classes;
    var periodLabel = this.labels().period;
  
 
  
  if (initialLock == "true"){
    
   var namingLabel = app.createLabel("These settings have been locked becuase folders have already been created.").setId('namingLabel').setStyleAttribute('marginTop', '5px');
   var namingGrid = app.createGrid(5, 2).setId('namingGrid').setCellPadding(3);
   namingGrid
    .setWidget(0, 0, app.createLabel(defaultLabels.dropBox))
    .setWidget(0, 1, app.createTextBox().setName('dropBox').setValue(dropBoxLabel).setEnabled(false))
    .setWidget(1, 0, app.createLabel(defaultLabels.dropBoxes))
    .setWidget(1, 1, app.createTextBox().setName('dropBoxes').setValue(dropBoxLabels).setEnabled(false))
    .setWidget(2, 0, app.createLabel(defaultLabels.class))
    .setWidget(2, 1, app.createTextBox().setName('class').setValue(classLabel).setEnabled(false))
    .setWidget(3, 0, app.createLabel(defaultLabels.classes))
    .setWidget(3, 1, app.createTextBox().setName('classes').setValue(classesLabel).setEnabled(false))
    .setWidget(4, 0, app.createLabel(defaultLabels.period)) 
    .setWidget(4, 1, app.createTextBox().setName('period').setValue(periodLabel).setEnabled(false));
   
  var closeButton = app.createButton('close');
  var closeHandler = app.createServerClickHandler('closePanel');
  closeButton.addClickHandler(closeHandler);
    
   panel.add(title);
   panel.add(namingLabel);
   panel.add(namingGrid);
   app.add(panel);
   app.add(closeButton);                                    
  } else {
  var saveHandler = app.createServerHandler('saveLabelSettings').addCallbackElement(panel);
  var namingLabel = app.createLabel("The terms below can be renamed. These labels determine how important folders and columns will be named by gClassFolders. Once you create the folders you will not be able to change these settings.").setId('namingLabel').setStyleAttribute('marginTop', '5px');
  var namingGrid = app.createGrid(5, 2).setId('namingGrid').setCellPadding(3);
    
  namingGrid
   .setWidget(0, 0, app.createLabel(defaultLabels.dropBox))
   .setWidget(0, 1, app.createTextBox().setName('dropBox').setValue(dropBoxLabel))
   .setWidget(1, 0, app.createLabel(defaultLabels.dropBoxes))
   .setWidget(1, 1, app.createTextBox().setName('dropBoxes').setValue(dropBoxLabels))
   .setWidget(2, 0, app.createLabel(defaultLabels.class))
   .setWidget(2, 1, app.createTextBox().setName('class').setValue(classLabel))
   .setWidget(3, 0, app.createLabel(defaultLabels.classes))
   .setWidget(3, 1, app.createTextBox().setName('classes').setValue(classesLabel))
   .setWidget(4, 0, app.createLabel(defaultLabels.period))
   .setWidget(4, 1, app.createTextBox().setName('period').setValue(periodLabel));  
     
  panel.add(title);
  panel.add(namingLabel);
  panel.add(namingGrid);
  app.add(panel);
  app.add(app.createButton("Save", saveHandler).setId('button'));
  }
  ss.show(app);
  return app;
}


function closePanel(){
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


// Saves folder label settings
function saveLabelSettings(e) {
  
  var documentProperties = PropertiesService.getDocumentProperties();
  var properties = documentProperties.getProperties();
  var app = UiApp.getActiveApplication();
  var dropBox= e.parameter.dropBox;
  var dropBoxes = e.parameter.dropBoxes;
  var class = e.parameter.class;
  var classes = e.parameter.classes;
  var period = e.parameter.period;
  properties.labels = Utilities.jsonStringify({dropBox: dropBox, dropBoxes: dropBoxes, class: class, classes: classes, period: period});
  properties.ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  documentProperties.setProperties(properties);
  var labels = properties.labels;
  CacheService.getPrivateCache().put('labels', labels, 660);
  
  //create or update RosterSheet
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var sheetId = properties.sheetId;
  if (!sheetId) {
    createRosterSheet();
    }else {
      fixHeaders();
    }
  
  app.close();
  return app;
  
}

//Creates a NEW sheet, inserts required headings, and stores the 
//sheet Id for use elsewhere (eliminates trusting that "Active" sheet contains the roster)
//This only ever runs once in most instances.  Runs from menu on first use or if the user has deleted the roster sheet for some reason.
function createRosterSheet(properties){
  var documentProperties = PropertiesService.getDocumentProperties();
  if (!properties) {
    var properties = PropertiesService.getDocumentProperties().getProperties();
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('gClassRoster');
  if ((!sheet)&&(!properties.sheetId)) {
    sheet = ss.insertSheet('gClassRoster');  
  }
  properties.sheetId = sheet.getSheetId();
  setgClassFoldersTeacherUid();
  setgClassFoldersTeacherSid();
  sheet.getRange("A1").setValue("Student First Name").setComment("Don't change the name of this header!");
  sheet.getRange("B1").setValue("Student Last Name").setComment("Don't change the name of this header!");
  sheet.getRange("C1").setValue("Student Email").setComment("Don't change the name of this header!");
  sheet.getRange("D1").setValue(this.labels().class + " Name").setComment("Don't change the name of this header!");
  sheet.getRange("E1").setValue(this.labels().period + " ~Optional~").setComment("Don't change the name of this header!");
  sheet.getRange("F1").setValue("Teacher Email(s)").setComment("Don't change the name of this header!");
  SpreadsheetApp.flush();
  sheet.setFrozenRows(1);
  properties.sheetId = sheet.getSheetId();
  documentProperties.setProperties(properties);
  var hideRange = sheet.getRange("H1:N1");
  sheet.hideColumn(hideRange);
  ss.setActiveSheet(sheet);

}

//Function used to create folder Id Headings when the user runs the folder creation process
//If the user is re-running folder creation, this checks to see if the headings exist
function createFolderIdHeadings(){
  var sheet = getRosterSheet();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('Student ' + this.labels().dropbox + ' Folder Id')==-1) {
    sheet.getRange(1, lastCol+1, 1, 6).setValues([['Student ' + this.labels().dropBox + ' Id','Class Root Folder' + ' Id','Class View Folder' + ' Id','Class Edit Folder' + ' Id','Root Student Folder' + ' Id','Teacher Folder Id']]).setComment("Don't manually change or delete any of these column headers or row values");
    SpreadsheetApp.flush();
  }
}


//function prompts user to fix messed up headers in the sheet
function badHeaders() {
  var button = Browser.Buttons.YES_NO;
  if(Browser.msgBox("Required headers are are missing or impropertly labeled. Do you want the script to try fixing your headers?", button))
  {
    fixHeaders();
    Browser.msgBox("gClassFolders has attempted to fix your headers.  Please check that everything in your roster sheet is as expected.");
  }
}




//function assigns headers to the sheet.  Headers are translated according to language and custom header settings.
function fixHeaders() {
  var sheet = getRosterSheet();
  sheet.getRange("A1").setValue("Student First Name").setComment("Don't change the name of this header!");
  sheet.getRange("B1").setValue("Student Last Name").setComment("Don't change the name of this header!");
  sheet.getRange("C1").setValue("Student Email").setComment("Don't change the name of this header!");
  sheet.getRange("D1").setValue(this.labels().class + " Name").setComment("Class folders are created only for unique class names. Don\'t change the name of this header!");
  sheet.getRange("E1").setValue(this.labels().period + " ~Optional~").setComment("Don't change the name of this header!");
  sheet.getRange("F1").setValue("Teacher Email(s)").setComment("Don't change the name of this header!");
}
