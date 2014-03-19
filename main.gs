var GCLASSICONURL = 'https://sites.google.com/site/gclassfolders/_/rsrc/1360538205205/config/customLogo.gif?revision=1';
var GCLASSLAUNCHERICONURL = 'https://sites.google.com/site/gclassfolders/_/rsrc/1360538205205/config/customLogo.gif?revision=1';
var userEmail = Session.getEffectiveUser().getEmail();
var SSKEY = PropertiesService.getDocumentProperties().getProperty('ssKey');

// This list was taken from the list of available languages in Google Translate service, responsible for our UI internationalization.
var googleLangList = ['English: en'];

// This object is responsible for returning the values for custom labels for "Assignment Folder", "Class", and "Period" throughout the script
var labels = function() { var labels = CacheService.getPrivateCache().get('labels');
                         if (!labels) {
                           labels = PropertiesService.getDocumentProperties().getProperty('labels');
                           CacheService.getPrivateCache().put('labels', labels, 660);
                         }
                         if (labels) {
                           labels = Utilities.jsonParse(labels);
                         } else {
                           labels =  {dropBox: "Assignment Folder", dropBoxes: "Assignment Folders", class: "Class", classes: "Classes", period: "Period"};
                         }
                         return labels;
                        }




//This function runs automatically when the spreadsheet opens, and provides the initial menu to the script.
//Defined separately from myOnOpen() to avoid issues with PropertiesService.getDocumentProperties() not being able to be called from a built in trigger
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var addOn = SpreadsheetApp.getUi().createAddonMenu();
  addOn.addItem("Initial settings", "gClassFolders_folderLabels");
  addOn.addItem("Create new folders and shares", "createClassFoldersCheck");
  addOn.addItem("Sort sheet by " + this.labels().class + ", " + this.labels().period + ", last name", "sortsheetCheck");
  addOn.addItem("Perform bulk operations on selected student(s)", "bulkOperationsUiCheck");
  //addOn.addItem("Quick Actions", "quickActions");
  addOn.addToUi();
  

//  var menu = SpreadsheetApp.getUi().createMenu('gClassFolders');
//  menu.addItem("Initial settings", "gClassFolders_folderLabels");
//  menu.addItem("Create new folders and shares", "createClassFolders");
//  menu.addItem("Sort sheet by " + this.labels().class + ", " + this.labels().period + ", last name", "sortsheet");
//  menu.addItem("Perform bulk operations on selected student(s)", "bulkOperationsUi");
//  //menu.addItem("Quick Actions", "quickActions");
//  menu.addToUi();
  
  
} //end onOpen


function createClassFoldersCheck(){
  var labels = PropertiesService.getDocumentProperties().getProperty('labels');
  var ui = SpreadsheetApp.getUi(); // Same variations.
  if (!labels) {
    var result = ui.alert(
      'You have not yet set up the page for gCF',
      'Would you like to initialize now',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    gClassFolders_folderLabels();
  } 
    
  } else { createClassFolders();}
} //end createClassFoldersCheck



function sortsheetCheck(){
    var labels = PropertiesService.getDocumentProperties().getProperty('labels');
  var ui = SpreadsheetApp.getUi(); // Same variations.
  if (!labels) {
    var result = ui.alert(
      'You have not yet set up the page for gCF',
      'Would you like to initialize now',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    gClassFolders_folderLabels();
  } 
    
  } else { sortsheet();}

}

function bulkOperationsUiCheck(){
      var labels = PropertiesService.getDocumentProperties().getProperty('labels');
  var ui = SpreadsheetApp.getUi(); // Same variations.
  if (!labels) {
    var result = ui.alert(
      'You have not yet set up the page for gCF',
      'Would you like to initialize now',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    gClassFolders_folderLabels();
  } 
    
  } else { bulkOperationsUi();}
  
}

//*********************  Used for debugging and testing only *****************************

//resets the spreadsheet to original. This will include onOpen() once that is finished testing.
function startOver(){ 
  PropertiesService.getDocumentProperties().deleteAllProperties();
  CacheService.getPrivateCache().remove('labels');
  var ss = SpreadsheetApp.getActive();
  var roster = ss.getSheetByName('gClassRoster');
  var prop = ss.getSheetByName('Properties');
  
  if (roster != null){ss.deleteSheet(roster);}
  if (prop != null){ss.deleteSheet(prop);}
  onOpen();
}

function documentPropertiesToLog(){
var properties = PropertiesService.getDocumentProperties().getProperties();
Logger.log(properties);
} //end documentPropertiesToLog

function privateCacheToLog(){
 // var theCache = this.labels().class;
  var theCache = this.labels();
  //var theCache2 = labels().class;
  var dropBoxLabel = this.labels().dropBox;
  var dropBoxLabels = labels().dropBoxes;
  var classLabel = labels().class;
  var classesLabel = labels().classes;
  var periodLabel = labels().period;
  Logger.log(theCache);
}


function lockInitialSettings(){
PropertiesService.getDocumentProperties().setProperty('alreadyRan', 'true');
}

