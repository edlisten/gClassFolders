function setResumeTrigger(lock) {
  lock.releaseLock();
  ScriptApp.newTrigger('createClassFolders').timeBased().after(30000).create();
  Browser.msgBox("Folder creation process will resume in 30 sec to avoid timeout");
  return;
}

function removeResumeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=='createClassFolders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  return;
}


//Because we can't know whether a user is installing gClassFolders in a spreadsheet with other sheets
//we need a way to access the roster from the sheet's immutable ID
//This function is to be used whenever we need to get the roster sheet
// (note: Sheet ID and sheet index are different.  Index is mutable, id is not)
function getRosterSheet() {
  var sheetId = parseInt(PropertiesService.getDocumentProperties().getProperty('sheetId'));
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


//used to sort roster sheet by classname, period, and last name
function sortsheet(classIndex, perIndex, lNameIndex) {
  var sheet = getRosterSheet();
  if ((!classIndex)||(!perIndex)||(lNameIndex)) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    classIndex = headers.indexOf(this.labels().class + " Name");
    perIndex = headers.indexOf(this.labels().period + " ~Optional~");
    lNameIndex = headers.indexOf("Student Last Name");
  }
  try {
    sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).sort([classIndex+1, perIndex+1, lNameIndex+1]);
  } catch(err) {
    Browser.msgBox("You cannot sort until you have entered student class enrollments");
  }
  //sort by cls then by Per
  //Logger.log(sheet);
}


function getClassRowNumsFromRSF(dataRange, indices, rsfId) {
  var classRowNums = [];
    for (var i=1; i<dataRange.length; i++) {
      if (dataRange[i][indices.rsfIdIndex]==rsfId) {
        classRowNums.push(i+1);
      }
    }
  return classRowNums;
}


function getClassRowNumsFromCRF(dataRange, indices, crfId) {
  var classRowNums = [];
    for (var i=1; i<dataRange.length; i++) {
      if (dataRange[i][indices.crfIdIndex]==crfId) {
        classRowNums.push(i+1);
      }
    }
  return classRowNums;
}

function getSectionsNotSelected(dataRange, indices, rsfIds, crfId) {
  var rootsNotSelected = [];
  for (var i=1; i<dataRange.length; i++ ) {
    if ((rsfIds.indexOf(dataRange[i][indices.rsfIdIndex])==-1)&&(dataRange[i][indices.crfIdIndex]==crfId)&&(rootsNotSelected.indexOf(dataRange[i][indices.rsfIdIndex])==-1)) {
      rootsNotSelected.push(dataRange[i][indices.rsfIdIndex]);
    }
  }
  return rootsNotSelected;
}


function getClassRoster(dataRange, indices, className, per) {
  var crfId = ''; // class root folder id
  var rsfId = ''; // root student folder id
  var cefId = ''; // class edit folder id
  var cvfId = ''; // class view folder id
  var classRows = [];
  classRows.push(dataRange[0])
  for (var i=1; i<dataRange.length; i++) {
    if (per) {
      if((dataRange[i][indices.clsNameIndex]==className)&&(dataRange[i][indices.clsPerIndex]==per)) {
        classRows.push(dataRange[i]);
      }
    }
  if ((!per)||(per=='')) {
      if(dataRange[i][indices.clsNameIndex]==className) {
        classRows.push(dataRange[i]);
      }
    }
  }
  return classRows;
}


function getClassRosterAsObjects(dataRange, indices, className, per) {
  var crfId = ''; // class root folder id
  var rsfId = ''; // root student folder id
  var cefId = ''; // class edit folder id
  var cvfId = ''; // class view folder id
  var classRows = [];
  var rowNums = [];
  for (var i=1; i<dataRange.length; i++) {
    if (per) {
      if((dataRange[i][indices.clsNameIndex]==className)&&(dataRange[i][indices.clsPerIndex]==per)) {
        classRows.push(dataRange[i]);
        rowNums.push(i+1);
      }
    }
    if ((!per)||(per=='')) {
      if(dataRange[i][indices.clsNameIndex]==className) {
        classRows.push(dataRange[i]);
        rowNums.push(i+1);
      }
    }
  }
  var studentObjects = [];
  for (var i=0; i<classRows.length; i++) {
    studentObjects[i] = new Object();
    studentObjects[i]['sFName'] = classRows[i][indices.sFnameIndex];
    studentObjects[i]['sLName'] = classRows[i][indices.sLnameIndex];
    studentObjects[i]['sEmail'] = classRows[i][indices.sEmailIndex];
    studentObjects[i]['dbfId'] = classRows[i][indices.dbfIdIndex]; 
    studentObjects[i]['cvfId'] = classRows[i][indices.cvfIdIndex]; 
    studentObjects[i]['cefId'] = classRows[i][indices.cefIdIndex];
    studentObjects[i]['rsfId'] = classRows[i][indices.rsfIdIndex];
    studentObjects[i]['crfId'] = classRows[i][indices.crfIdIndex];
    studentObjects[i]['tfId'] = classRows[i][indices.tfIdIndex];
    studentObjects[i]['clsName'] = classRows[i][indices.clsNameIndex];
    studentObjects[i]['clsPer'] = classRows[i][indices.clsPerIndex];
    studentObjects[i]['tEmail'] = classRows[i][indices.tEmailIndex];
    studentObjects[i]['row'] = rowNums[i];
  }
  return studentObjects;
}




function getClassFolderId(classRoster, folderIndex) {
  var folderId="";
  for (var i=1; i<classRoster.length; i++) {
    if (classRoster[i][folderIndex]!="") {
      folderId = classRoster[i][folderIndex];
      return folderId;
    }
  }
  return folderId;
}


function getUniqueClassNames(dataRange, clsNameIndex, crfIdIndex) {
  var classNames = [];
  for (var i=1; i<dataRange.length; i++) {
    var thisClassName = dataRange[i][clsNameIndex];
    var thisClassRoot = dataRange[i][crfIdIndex];
    if ((classNames.indexOf(thisClassName)==-1)&&(thisClassName!='')&&(thisClassRoot!='')) {
      classNames.push(thisClassName);
    }
  }
  classNames.sort();
  return classNames;
}


function returnEmailAsArray(emailValue) {
  emailValue = emailValue.replace(/\s+/g, '');
  var emailArray = emailValue.split(",");
  return emailArray;  
}


function getUniqueClassPeriods(dataRange, clsNameIndex, clsPerIndex, rsfIdIndex, labelObject) {
  var classPers = [];
  for (var i=0; i<dataRange.length; i++) {
    var thisClassPer = dataRange[i][clsNameIndex] + " " + labelObject.period + " " + dataRange[i][clsPerIndex];
    var thisStudentRoot = dataRange[i][rsfIdIndex];
    if ((classPers.indexOf(thisClassPer)==-1)&&(thisClassPer!='')&&(thisStudentRoot!='')) {
      classPers.push(thisClassPer);
    }
  }
  classPers.sort();
  return classPers;
}

function getRootClassFoldersByRSF(dataRange, rsfId, rsfIdIndex, crfIdIndex, cefIdIndex, cvfIdIndex) {
  for (var i=1; i<dataRange.length; i++) {
    var thisRsfId = dataRange[i][rsfIdIndex];
    if (thisRsfId == rsfId) {
      var obj = new Object();
      obj.crfId = dataRange[i][crfIdIndex];
      obj.cefId = dataRange[i][cefIdIndex];
      obj.cvfId = dataRange[i][cvfIdIndex];
      return obj;
    }
  }
  return;
}

function getTeacherEmailsByRSF(dataRange, rsfId, rsfIdIndex, tEmailIndex) {
  for (var i=1; i<dataRange.length; i++) {
    var thisRsfId = dataRange[i][rsfIdIndex];
    if (thisRsfId == rsfId) {
      var obj = new Object();
      obj.tEmails = dataRange[i][tEmailIndex];
      return obj;
    }
  }
  return;
}



function getUniqueClassPeriodObjects(dataRange, clsNameIndex, clsPerIndex, rsfIdIndex, labelObj) {
  var classPers = [];
  var processed = [];
  var k = 0;
  for (var i=1; i<dataRange.length; i++) {
    var thisClassPer = dataRange[i][clsNameIndex];
    if (dataRange[i][clsPerIndex]!='') {
      thisClassPer += " " + labelObj.period + " " + dataRange[i][clsPerIndex];
    }
    var thisStudentRoot = dataRange[i][rsfIdIndex];
    if ((processed.indexOf(thisClassPer)==-1)&&(thisClassPer!='')&&(thisStudentRoot!='')) {
      classPers[k] = new Object();
      classPers[k].classPer = thisClassPer;
      classPers[k].rsfId = thisStudentRoot;
      processed.push(thisClassPer);
      k++;
    }
  }
  classPers.sort(
    function compareNames(a, b) {
      var nameA = a.classPer.toLowerCase( );
      var nameB = b.classPer.toLowerCase( );
      if (nameA < nameB) {return -1}
      if (nameA > nameB) {return 1}
      return 0;
    })
    return classPers;
}

function saveIndices(indices) {
  if (indices) {
    var indicesString = Utilities.jsonStringify(indices);
    PropertiesService.getDocumentProperties().setProperty('indices', indicesString);
  }
}


function writeProperties() {
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var propertyArray = [];
  var i = 0;
  for (var key in properties) {
    propertyArray[i] = [];
    propertyArray[i][0] = key;
    propertyArray[i][1] = properties[key]
    i++;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName('Properties');
  if (!sheet) {
    sheet = ss.insertSheet('Properties');
  }
  sheet.getRange(1, 1, propertyArray.length, 2).setValues(propertyArray);
  sheet.getRange("A1").setComment("This sheet is used by gClassHub to understand how your roster is organized.");
  //set focus back to 'gClassFolders'
  var gClassRoster = ss.getSheetByName('gClassRoster');
  SpreadsheetApp.setActiveSheet(gClassRoster);
  
}


function returnIndices(dataRange, labelObject) {
  var sheet = getRosterSheet();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var sFnameIndex = headers.indexOf('Student First Name');
  if (sFnameIndex==-1) {
    badHeaders();
    return;
  }
  var sLnameIndex = headers.indexOf('Student Last Name');
  if (sLnameIndex==-1) {
    badHeaders();
    return;
  }
  var sEmailIndex = headers.indexOf('Student Email');
  if (sEmailIndex==-1) {
    badHeaders();
    return;
  }
  var clsNameIndex  = headers.indexOf(labelObject.class +' Name');
  if (clsNameIndex==-1) {
    badHeaders();
    return;
  }
  var clsPerIndex = headers.indexOf(labelObject.period + " ~Optional~");
  if (clsNameIndex==-1) {
    badHeaders();
    return;
  }
  var tEmailIndex = headers.indexOf('Teacher Email(s)');
  if (tEmailIndex==-1) {
    badHeaders();
    return;
  }  
  
  //Add columns for tracking status of folder creation and share if they don't already exist
  //retrieve their indices
  
  var sDropStatusIndex = headers.indexOf("Status: Student " + labelObject.dropBox);
  if (sDropStatusIndex==-1) {
    sheet.getRange(1,lastCol+1,1,2).setValues([["Status: Student " + labelObject.dropBox,"Status: Teacher Share"]]).setComment("Don't change this header. When gClassFolders is run, class and dropbox folders will be created or updated for any students without a value in this column. To update a student's email address or name, just clear their status value and run again. To move students between classes, use the menu.");
    headers.push("Status: Student " + labelObject.dropbox);
    headers.push("Status: Teacher Share");
    SpreadsheetApp.flush()
  }
  sDropStatusIndex = headers.indexOf("Status: Student " + labelObject.dropBox);
  var tShareStatusIndex = headers.indexOf("Status: Teacher Share");
  var dbfIdIndex = headers.indexOf('Student ' + labelObject.dropBox + ' Id');
  if (dbfIdIndex==-1) {
    createFolderIdHeadings(); //create Folder ID headings if they don't exist
  }
  
  //refresh headers to pull in new folder id columns;
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  //Get indices of folder Id columns
  var indices = new Object();
  indices.dbfIdIndex = headers.indexOf('Student ' + labelObject.dropBox + ' Id');
  indices.crfIdIndex = headers.indexOf('Class Root Folder Id');
  indices.cvfIdIndex = headers.indexOf('Class View Folder Id');
  indices.cefIdIndex = headers.indexOf('Class Edit Folder Id');
  indices.rsfIdIndex = headers.indexOf('Root Student Folder Id');
  indices.tfIdIndex = headers.indexOf('Teacher Folder Id');
  indices.sFnameIndex = sFnameIndex;
  indices.sLnameIndex = sLnameIndex;
  indices.sEmailIndex = sEmailIndex;
  indices.clsNameIndex = clsNameIndex;
  indices.clsPerIndex = clsPerIndex;
  indices.tEmailIndex = tEmailIndex;
  indices.sDropStatusIndex = sDropStatusIndex;
  indices.tShareStatusIndex = tShareStatusIndex;
  indices.dbfIdIndex = dbfIdIndex;
  return indices;
}



function gClassFolders_createGATrackingUrl(encoded_page_name)
{
  var utmcc = gClassFolders_createGACookie();
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (eduSetting=="true") {
    encoded_page_name = "edu/" + encoded_page_name;
  }
  if (utmcc == null)
    {
      return null;
    }
 
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.gClassFolders-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=UA-38070753-1&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  return ga_url_full;
}


function gClassFolders_createGACookie()
{
  var a = "";
  var b = "100000000";
  var c = "200000000";
  var d = "";

  var dt = new Date();
  var ms = dt.getTime();
  var ms_str = ms.toString();
 
  var gClassFolders_school_uid = UserProperties.getProperty("gClassFolders_school_uid");
  var gClassFolders_teacher_uid = UserProperties.getProperty("gClassFolders_teacher_uid");
  if ((gClassFolders_teacher_uid == null) && (gClassFolders_school_uid == ""))
    {
      // shouldn't happen unless user explicitly removed flubaroo_uid from properties.
      return null;
    }
  
  if (gClassFolders_teacher_uid) {
    a = gClassFolders_teacher_uid.substring(0,9);
    d = gClassFolders_teacher_uid.substring(9);
  }
  
  if (gClassFolders_school_uid) {
    a = gClassFolders_school_uid.substring(0,9);
    d = gClassFolders_school_uid.substring(9);
  }
  
  utmcc = "__utma%3D451096098." + a + "." + b + "." + c + "." + d 
          + ".1%3B%2B__utmz%3D451096098." + d + ".1.1.utmcsr%3D(direct)%7Cutmccn%3D(direct)%7Cutmcmd%3D(none)%3B";
 
  return utmcc;
}



function gClassFolders_logStudentFolderCreation()
{
  var ga_url = gClassFolders_createGATrackingUrl("Student%20Class%20Folder%20Created");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}


function gClassFolders_logTeacherClassFolderCreated()
{
  var ga_url = gClassFolders_createGATrackingUrl("Teacher%20Class%20Folder%20Created");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}


function gClassFolders_logStudentClassFolderArchived()
{
  var ga_url = gClassFolders_createGATrackingUrl("Student%20Class%20Folder%20Archived");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}


function gClassFolders_getInstitutionalTrackerObject() {
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  if ((institutionalTrackingString)&&(institutionalTrackingString != "not participating")) {
    var institutionTrackingObject = Utilities.jsonParse(institutionalTrackingString);
    return institutionTrackingObject;
  }
}


function gClassFolders_logRepeatTeacherInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("Repeat%20Teacher%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}



function gClassFolders_logRepeatSchoolInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("Repeat%20School%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}


function gClassFolders_logFirstTeacherInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("First%20Teacher%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}


function gClassFolders_logFirstSchoolInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("First%20School%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
}



function setgClassFoldersTeacherUid()
{ 
  var gClassFolders_teacher_uid = UserProperties.getProperty("gClassFolders_teacher_uid");
  if (gClassFolders_teacher_uid == null || gClassFolders_teacher_uid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
 
      UserProperties.setProperty("gClassFolders_teacher_uid", ms_str);
      gClassFolders_logFirstTeacherInstall();
    }
}


function setgClassFoldersSchoolUid()
{ 
  var gClassFolders_school_uid = UserProperties.getProperty("gClassFolders_school_uid");
  if (gClassFolders_school_uid == null || gClassFolders_school_uid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
 
      UserProperties.setProperty("gClassFolders_school_uid", ms_str);
      gClassFolders_logFirstSchoolInstall();
    }
}


function setgClassFoldersTeacherSid()
{ 
  var gClassFolders_teacher_sid = PropertiesService.getDocumentProperties().getProperty("gClassFolders_teacher_sid");
  if (gClassFolders_teacher_sid == null || gClassFolders_teacher_sid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
      PropertiesService.getDocumentProperties().setProperty("gClassFolders_teacher_sid", ms_str);
      var gClassFolders_teacher_uid = UserProperties.getProperty("gClassFolders_teacher_uid");
      if (gClassFolders_teacher_uid != null || gClassFolders_teacher_uid != "") {
        gClassFolders_logRepeatTeacherInstall();
      }
    }
}


function setgClassFoldersSchoolSid()
{ 
  var gClassFolders_teacher_sid = PropertiesService.getDocumentProperties().getProperty("gClassFolders_school_sid");
  if (gClassFolders_teacher_sid == null || gClassFolders_teacher_sid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
      PropertiesService.getDocumentProperties().setProperty("gClassFolders_school_sid", ms_str);
      var gClassFolders_school_uid = UserProperties.getProperty("gClassFolders_school_uid");
      if (gClassFolders_school_uid != null || gClassFolders_school_uid != "") {
        gClassFolders_logRepeatSchoolInstall();
      }
    }
}


/**
* Invokes a function, performing up to 5 retries with exponential backoff.
* Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of 
* about 32 seconds before it gives up and rethrows the last error. 
* See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff 
* <br>Author: peter.herrmann@gmail.com (Peter Herrmann)
<h3>Examples:</h3>
<pre>//Calls an anonymous function that concatenates a greeting with the current Apps user's email
var example1 = GASRetry.call(function(){return "Hello, " + Session.getActiveUser().getEmail();});
</pre><pre>//Calls an existing function
var example2 = GASRetry.call(myFunction);
</pre><pre>//Calls an anonymous function that calls an existing function with an argument
var example3 = GASRetry.call(function(){myFunction("something")});
</pre><pre>//Calls an anonymous function that invokes DocsList.setTrashed on myFile and logs retries with the Logger.log function.
var example4 = GASRetry.call(function(){myFile.setTrashed(true)}, Logger.log);
</pre>
*
* @param {Function} func The anonymous or named function to call.
* @param {Function} optLoggerFunction Optionally, you can pass a function that will be used to log 
to in the case of a retry. For example, Logger.log (no parentheses) will work.
* @return {*} The value returned by the called function.
*/
function call(func, optLoggerFunction) {
  for (var n=0; n<6; n++) {
    try {
      return func();
    } catch(e) {
      if (optLoggerFunction) {optLoggerFunction("GASRetry " + n + ": " + e)}
      if (n == 5) {
        throw e;
      } 
      Utilities.sleep((Math.pow(2,n)*500) + (Math.round(Math.random() * 500)));
    }    
  }
}
