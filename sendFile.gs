


function sendFile() {
    var html = HtmlService.createHtmlOutputFromFile('PickerSendFile.html')
      .setWidth(600).setHeight(425);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a template file to copy to assignment folders.');
}



// runs after the picker and creates the next dialogue box.
function fileToSendResults(templateTitle,templateId, templateURL, templateDoc){
  var docInfoArray = [templateTitle,templateId, templateURL];
  
  if (!templateTitle){ 
    templateTitle = "No File Selected";
    templateId = "No File Selected";
    templateURL = "No File Selected";
    templateDoc = "No File Selected";
  }
  var app = UiApp.createApplication().setTitle("Send File To:").setHeight(410);
  var panel = app.createVerticalPanel();
  
   //Question Info
   var html = "Template Info";
   html += "<ul><li>Name: <b>"+templateTitle+ "</b></li>";
//  html += "<li>ID: " + templateId + "</li>";
//  html += "<li>url: " + templateURL + "</li>";
//  html += "<li>Doc: " + templateDoc + "</li>";
   html += "<hr>";
   
   var questionInfo = app.createHTML(html).setStyleAttribute('border','1px').setStyleAttribute('margin', '8px');
  
  //handlers
  var studentListH = app.createServerHandler("studentListHF").addCallbackElement(panel);
  var populateStudentListH = app.createServerHandler('populateStudentListH').addCallbackElement(panel);
  
  // Copied File New Name
   var cfNewName = app.createTextBox().setName('cfNewName').setValue(templateTitle);
   var cfStudentTF = app.createCheckBox("Append Student Name").setName('cfStudentTF').setValue(true);
   var cfGrid = app.createGrid(2,2);
  cfGrid
  .setWidget(0,0, app.createLabel("Set copied file name:"))
  .setWidget(1,0, cfNewName)
  .setWidget(1,1, cfStudentTF);
  
  
  // Class List to choose from.
  var chooseLabel = app.createLabel("Select destination").setStyleAttribute('marginRight', '8px').setStyleAttribute('marginTop', '8px');
  var classList = app.createListBox().setId('classList').setName('classList').setStyleAttribute('margin', '8px').addChangeHandler(studentListH);
  returnClassesToSendFile();
  
  // student choice
  var studentList = app.createListBox().setId('studentList').setName('studentList').setStyleAttribute('margin', '8px').addChangeHandler(studentListH);
  studentList.addItem("All Students", "allStudents");
  studentList.addItem("Select Students", "selectStudents");
 
  // Close
  var closeHandler = app.createServerClickHandler('closePanel');
  var closeButton = app.createButton('Close').setId('closeButton').addClickHandler(closeHandler);
  
  // Send
  var sendButtonH = app.createServerHandler("sendButtonH").addCallbackElement(panel);
  var sendButton = app.createButton('Send Template To Selected').setId("sendButton").addClickHandler(sendButtonH).addClickHandler(closeHandler);
  
  // button grid
  var buttonGrid = app.createGrid(1,2);
  buttonGrid
    .setWidget(0, 0, sendButton)
    .setWidget(0, 1, closeButton);
  
  //hidden elements
  var docInfo = app.createHidden().setName("docInfo").setID("docInfo").setValue(JSON.stringify(docInfoArray));

  //Selected Students List
  var selectStudentsLabel = app.createLabel("Hold CTRL key to select multiple").setId("selectStudentsLabel").setVisible(false);
  var selectStudentsList = app.createListBox(true).setId('selectStudentsList').setPixelSize(300, 100).setName('selectStudentsList').setStyleAttribute('margin', '8px').setEnabled(false);
  var selectedAFIds = getAssignmentFolderIds('allClasses');
  for (var i=0; i<selectedAFIds[0].length; i++){
      selectStudentsList.addItem(selectedAFIds[0][i],selectedAFIds[1][i]).setItemSelected(i, true); 
//      selectStudentsList.addItem(selectedAFIds[0][i],[selectedAFIds[0][i],selectedAFIds[1][i]]).setItemSelected(i, true); 
  }


  //hidden
  panel.add(docInfo);
  
  //build ui
  panel.add(questionInfo);
  panel.add(cfGrid);
  panel.add(chooseLabel);
  panel.add(classList);
  panel.add(studentList);
  panel.add(selectStudentsLabel);
  panel.add(selectStudentsList);
  panel.add(buttonGrid);
  app.add(panel);
  SpreadsheetApp.getUi().showModalDialog(app, "Choose a class and send");
  
}




function sendButtonH(e){
  try {
  var templateId = JSON.parse(e.parameter.docInfo); //[templateTitle,templateId, templateURL]
  var selectStudentsList = e.parameter.selectStudentsList;
//  Logger.log(selectStudentsList);
  } catch(err){}
  
  
//  if  ((!templateId)||(!selectStudentsList)) { 
//    var templateId = ['Doctopus How-to', '1m4Lq6x66Th3O5kCBDlB8KuzN0zMATbqH3fgIPVqIs4U', 'https://docs.google.com/document/d/1m4Lq6x66Th3O5kCBDlB8KuzN0zMATbqH3fgIPVqIs4U/edit?usp=drive_web'];
//    var selectStudentsList = "0Bwyqwd2fAHMMNXZLRGQ1THpDaHM,0Bwyqwd2fAHMMTTNiNm5tVGt1WEE,0Bwyqwd2fAHMMMHFUNXVGQlBVX0U,0Bwyqwd2fAHMMblhjS0EwazdDS2c";
//  } 
  
  // Creates the Send File Sheet
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var documentProperties = PropertiesService.getDocumentProperties();
 var sendFileSheet = getSendFileSheet(); 
 var status = "Error"; 
 
 //var sendFileSheet = ss.getSheetByName('Send File Log');  //just for testing
  
 // Create Send File Sheet
  if (!sendFileSheet){ 
    sendFileSheet = ss.insertSheet('Send File Log');
    sendFileSheet.getRange('A1').setValue('Date Sent');
    sendFileSheet.getRange('A2').setValue('dateSent'); 
    sendFileSheet.getRange('B1').setValue('Template');
    sendFileSheet.getRange('B2').setValue('template');
    sendFileSheet.getRange('C1').setValue('Status');
    sendFileSheet.getRange('C2').setValue('status');
    sendFileSheet.getRange('D1').setValue('Permission');
    sendFileSheet.getRange('D2').setValue('permission');
    sendFileSheet.getRange('E1').setValue('Template ID');
    sendFileSheet.getRange('E2').setValue('templateId');
    sendFileSheet.getRange('F1').setValue('Copied File Name');
    sendFileSheet.getRange('F2').setValue('copiedFileName');
    sendFileSheet.getRange('G1').setValue('Copied File ID\'s');
    sendFileSheet.getRange('G2').setValue('copiedFileIds');
    sendFileSheet.getRange('H1').setValue('Assigment Folders');
    sendFileSheet.getRange('H2').setValue('assigmentFolders');
    sendFileSheet.setFrozenRows(2);
    sendFileSheet.hideRows(2);
    sendFileSheet.setColumnWidth(1,110);   
    sendFileSheet.setColumnWidth(2,200); 
    sendFileSheet.setColumnWidth(3,110);
    sendFileSheet.setColumnWidth(4,110);
    sendFileSheet.setColumnWidth(5,150);
    sendFileSheet.setColumnWidth(6,150);
    sendFileSheet.setColumnWidth(7,200);
    sendFileSheet.setColumnWidth(8,200);
    sendFileSheet.getRange('H:H').setWrap(true);
    sendFileSheet.getRange('A:A').setNumberFormat("dd/MM/yy HH:mm");
    
    var sendFileSheetId = sendFileSheet.getSheetId();   
    documentProperties.setProperty('sendFileSheetId', sendFileSheetId); 
    SpreadsheetApp.flush();
  }
  
  
  
  //Process
  var destFolderName = e.parameter.cfNewName;
  var tfStudentName = e.parameter.cfStudentTF;

  
  try{
    var newDestFolderName = destFolderName;
    if (tfStudentName == 'true'){newDestFolderName += " - Student Name";}   
    var newFileArray = copyFileToFolder(templateId[1],selectStudentsList,destFolderName,tfStudentName);
    status = "Copied";
  }catch(err){
    status = "Error sending File: " + err;
  }
  
  
  // Add to Sheet
  var indices = sendFileIndices();
  var lastRow =  sendFileSheet.getLastRow()+1;
  var now = new Date();
  var templateLink = '=hyperlink("'+ templateId[2] +'";"'+templateId[0] + '")'
  
  sendFileSheet.getRange(lastRow,indices.dateSent+1).setValue(now);
  sendFileSheet.getRange(lastRow,indices.template+1).setValue(templateLink);
  sendFileSheet.getRange(lastRow,indices.status+1).setValue(status);
  sendFileSheet.getRange(lastRow,indices.permission+1).setValue("Edit");
  sendFileSheet.getRange(lastRow,indices.templateId+1).setValue(templateId[1]);
  sendFileSheet.getRange(lastRow,indices.copiedFileName+1).setValue(newDestFolderName);
  sendFileSheet.getRange(lastRow,indices.copiedFileIds+1).setValue(newFileArray);
  sendFileSheet.getRange(lastRow,indices.assigmentFolders+1).setValue(selectStudentsList);
  SpreadsheetApp.flush();
  
}


function copyFileToFolder(sourceFileId,destFolderId,destFolderName,tfStudentName){
  
//  if  ((!sourceFileId)||(!destFolderId)) { 
//    var sourceFileId = "1m4Lq6x66Th3O5kCBDlB8KuzN0zMATbqH3fgIPVqIs4U";
//    var tfStudentName = true;
//    var destFolderName = "new Name";
//    var destFolderId = "0Bwyqwd2fAHMMVmNsdl9QNllnNDA,0Bwyqwd2fAHMMNXZLRGQ1THpDaHM,0Bwyqwd2fAHMMTTNiNm5tVGt1WEE,0Bwyqwd2fAHMMMHFUNXVGQlBVX0U,0Bwyqwd2fAHMMblhjS0EwazdDS2c";
//  }

  var destFolderIdArray = destFolderId.split(',');
  
  var rosterSheet = getRosterSheet();
  var rosterIndicies = returnIndices();
  var sourceFile = DriveApp.getFileById(sourceFileId);
  var newFileArray = [];

  for (var j = 0; j<destFolderIdArray.length; j++){
    var testDest = destFolderIdArray[j];
    var destFolder = DriveApp.getFolderById(destFolderIdArray[j]);
    var studentName = "";
    var studentNameReturn = "";
    if (tfStudentName == 'true'){
      for (var i=3; i < rosterSheet.getLastRow(); i++){
        var currentRowAFId = rosterSheet.getRange(i,rosterIndicies.dbfIdIndex+1).getValue();
        if (destFolderIdArray[j] == currentRowAFId){
          var sfName = rosterSheet.getRange(i,rosterIndicies.sFnameIndex+1).getValue();
          var slName = rosterSheet.getRange(i,rosterIndicies.sLnameIndex+1).getValue();
          studentName = " - " + sfName+" "+slName;
        } 
      }
    }
    var newFile = sourceFile.makeCopy(destFolder).setName(destFolderName + studentName);
    newFileArray.push(newFile.getId());
    
  }
  return newFileArray;
}



function returnClassesToSendFile(e){
  var app = UiApp.getActiveApplication();
  var ui = SpreadsheetApp.getUi();
  var sheet = getRosterSheet();
  var dataRange = sheet.getDataRange().getValues();
  var indices = returnIndices();
  var selectedAFIds = getAssignmentFolderIds();

  
  var classList = app.getElementById('classList');
  classList.clear();
  classList.addItem("All Classes", "allClasses");
  var classNames = getUniqueClassNamesWithId(dataRange, indices.clsNameIndex, indices.crfIdIndex, indices.clsPerIndex, indices.rsfIdIndex)
  for (var i = 0; i < classNames[0].length; i++) { 
    classList.addItem(classNames[0][i],classNames[1][i]);
  }

  return app;
}


function studentListHF(e){
  var app = UiApp.getActiveApplication();
  var classListValue = e.parameter.classList;
  var studentListValue = e.parameter.studentList;
  var selectStudentsList = app.getElementById('selectStudentsList');
  var selectedAFIds = getAssignmentFolderIds(classListValue);
  
  
  selectStudentsList.clear().setEnabled(false);
           for (var i=0; i<selectedAFIds[0].length; i++){
              selectStudentsList.addItem(selectedAFIds[0][i],selectedAFIds[1][i]).setItemSelected(i, true); 
           }
  
  switch (studentListValue){
    case "allStudents":
      app.getElementById('selectStudentsLabel').setVisible(false);
      selectStudentsList.setEnabled(false);
           for (var i=0; i<selectedAFIds[0].length; i++){
              selectStudentsList.setItemSelected(i, true);            
           }
    break; 
      
    case "selectStudents":
      app.getElementById('selectStudentsLabel').setVisible(true);
      selectStudentsList.setEnabled(true);
  
           for (var i=0; i<selectedAFIds[0].length; i++){
              selectStudentsList.setItemSelected(i, false);            
           }
    break; 

    default:
      Browser.msgBox("Aww Snap, Something went wrong.");
  }
   
  return app;
}


/////////////////////////////////////////
function getAssignmentFolderIds(parentFolderId){ 
  var sheet = getRosterSheet();
  var dataRange = sheet.getDataRange().getValues();
  var indices = returnIndices();
  var clsNameIndex = indices.clsNameIndex;
  var crfIdIndex = indices.crfIdIndex;
  var rsfIdIndex = indices.rsfIdIndex;
  var dbfIdIndex = indices.dbfIdIndex; 
  var sFnameIndex = indices.sFnameIndex;
  var sFnameIndex = indices.sLnameIndex;
  var perNameIndex = indices.clsPerIndex;
  var assignmentFolderIds =[[],[]];
  var suffix = "";
  
  
  
//  if (!parentFolderId){
//    parentFolderId = dataRange[10][rsfIdIndex];
////    parentFolderId = dataRange[7][crfIdIndex];
//  }
  
 for (var i=2; i<dataRange.length; i++) {
   var aFolderId = dataRange[i][dbfIdIndex];
   var clsName = dataRange[i][clsNameIndex];
   var rootId = dataRange[i][rsfIdIndex];
   var crfId = dataRange[i][crfIdIndex];
   var sFname = dataRange[i][sFnameIndex];
   var sLname = dataRange[i][sFnameIndex];
   var perName = dataRange[i][perNameIndex];
   
   if (parentFolderId == "allClasses"){
     suffix = " - "+clsName +" "+ perName;
     assignmentFolderIds[0].push(sFname +" "+sLname + suffix);
     assignmentFolderIds[1].push(aFolderId);
   } else {
   
   if (crfId != rootId){
     if (parentFolderId == crfId){
     rootId = dataRange[i][crfIdIndex];
       if (perName !="") {suffix = " - "+perName; } 
     }
   }
   
   if (rootId ==parentFolderId){
     assignmentFolderIds[0].push(sFname +" "+sLname + suffix);
     assignmentFolderIds[1].push(aFolderId);
   }
   }
 }

  return assignmentFolderIds;
  
}
  

                                           
  function getUniqueClassNamesWithId(dataRange, clsNameIndex, crfIdIndex, clsPerIndex, rsfIdIndex) {
    if (!dataRange){
      try{
  var sheet = getRosterSheet();
  var dataRange = sheet.getDataRange().getValues();
  var indices = returnIndices();
  var clsNameIndex = indices.clsNameIndex;
  var crfIdIndex = indices.crfIdIndex;
  var clsPerIndex = indices.clsPerIndex; 
  var rsfIdIndex = indices.rsfIdIndex;      
      } catch(err){ Logger.log(err);}
    }
      
  var classNames = [[],[]];
  for (var i=2; i<dataRange.length; i++) {
    var thisClassName = dataRange[i][clsNameIndex];
    var thisClassPer = dataRange[i][clsPerIndex];
    var thisClassRoot = dataRange[i][crfIdIndex];
    var thisRootId = dataRange[i][rsfIdIndex];
    
    if ((classNames[0].indexOf(thisClassName)==-1)&&(thisClassName!='')&&(thisClassRoot!='')) {
      classNames[0].push(thisClassName);
      classNames[1].push(thisClassRoot);
    }
    
    if ((classNames[0].indexOf(thisClassName)>-1)&&(classNames[0].indexOf(thisClassName+"-"+thisClassPer)==-1)&&(thisClassPer!='')){  
      classNames[0].push(thisClassName+"-"+thisClassPer);
      classNames[1].push(thisRootId);
    }
    
  }
  return classNames;
}                                         
                                           
 
