//http://stackoverflow.com/questions/19045064/google-apps-script-active-cell-changes-in-handler-from-the-onedit-code-that-call
// http://stackoverflow.com/questions/12549085/using-and-modifying-global-variables-within-handler-functions
// https://sites.google.com/site/appsscriptforbusiness/example---how-to/twolistboxes

// create hidden field that holds the row number.  That way I am not relying on the current active cell, but rather the one that was selected when the button was hit

function createLogSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var documentProperties = PropertiesService.getDocumentProperties();
  var logSheetId = getLogSheet();
  if (!logSheetId) {
    var logSheet = ss.insertSheet('Log');
    logSheetId = logSheet.getSheetId();
    documentProperties.setProperty('logSheetID', logSheetId);
    logSheet.getRange('A1').setValue('Identifier');
    logSheet.getRange('B1').setValue('Information');
    logSheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    
  }
}





function actionsSidebar() {
  if (setupGcfCheck() == false){ return;}
 
  
  var app = UiApp.createApplication().setTitle("Student Management");
  var panel = app.createVerticalPanel().setWidth("250px");
  
  //handlers
  var updateForSelected = app.createServerHandler("updateForSelected").addCallbackElement(panel);
  var getCurrentHandler = app.createServerHandler("getSelectedHandler").addCallbackElement(panel);
  var runActionHandler = app.createServerHandler("runActionHandler").addCallbackElement(panel);
  var closeHandler = app.createServerClickHandler('closePanel');
  
  //panels
  var quesInfoPanel = app.createVerticalPanel().setId("quesInfoPanel").setStyleAttribute('margin', '8px').setWidth("100%");
  var actionLabel = app.createLabel("Select an action").setStyleAttribute('margin', '8px').setStyleAttribute('marginBottom', '0px');
  var actionOnSelected = app.createListBox().setId("actionOnSelected").setName('actionOnSelected').addChangeHandler(updateForSelected).setStyleAttribute('margin', '8px').setStyleAttribute('marginTop', '0px').setWidth("100%");
  var forSelected = app.createListBox().setId("forSelected").setName('forSelected').setStyleAttribute('margin', '8px').setWidth("100%").setVisible(false) 
  var actionDescription = app.createLabel("").setHeight("140px").setWidth("100%").setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('margin', '8px');
  var actionSettingsPanel = app.createVerticalPanel().setId('actionSettingsPanel').setStyleAttribute('margin', '8px').setWidth("100%").setVisible(false) 
  
  //buttons
  var getCurrentButton = app.createButton('Update Selected').setStyleAttribute('margin', '8px').addClickHandler(getCurrentHandler).setStyleAttribute('background', 'FFFFFF').setStyleAttribute('border', '0').setStyleAttribute('color', 'blue').setStyleAttribute('textDecoration', 'underline');
  var runActionButton = app.createButton('Run Action').setId("actionButton").setStyleAttribute('margin', '8px').addClickHandler(runActionHandler);
  var closeButton = app.createButton('Close').setId('closeButton').setStyleAttribute('margin', '8px').addClickHandler(closeHandler);
  
  //hidden
  var selectedStudentsH = app.createHidden().setValue("").setId("selectedStudentsH").setName("selectedStudentsH");
  
  

  //build the interface  
  getSelectedHandler();
  panel.add(getCurrentButton);
  panel.add(quesInfoPanel);
  panel.add(actionLabel);
  panel.add(actionOnSelected);
  panel.add(forSelected);
  panel.add(actionSettingsPanel);
  panel.add(closeButton);
  panel.add(runActionButton);
  panel.add(selectedStudentsH);
  app.add(panel);
  SpreadsheetApp.getUi().showSidebar(app);
  
}




/////////////////////////////////////////////////////////////////////////////////////////////////////////
function updateForSelected(e){
  var ss = SpreadsheetApp.getActiveSheet();
  var app = UiApp.getActiveApplication();
  var ssAR = SpreadsheetApp.getActiveRange();
  var numbOfSel = ssAR.getNumRows();
  var indices = returnIndices();
  var dataRange = ss.getDataRange().getValues();
  var selectedAction = e.parameter.actionOnSelected;
  var forSelectedList = app.getElementById("forSelected");
  var actionSettingsPanel = app.getElementById('actionSettingsPanel');
  var studentRowIndex =JSON.parse(e.parameter.selectedStudentsH); //array
  forSelectedList.setVisible(false).clear();
  actionSettingsPanel.setVisible(false).clear();
  
  //get selected
  if (numbOfSel == 1){
   var sFname = dataRange[studentRowIndex][indices.sFnameIndex];
   var sLname = dataRange[studentRowIndex][indices.sLnameIndex]; 
   var sEmail = dataRange[studentRowIndex][indices.sEmailIndex];  
   var clsName = dataRange[studentRowIndex][indices.clsNameIndex]; 
   var clsPer = dataRange[studentRowIndex][indices.clsPerIndex]; 
  

    switch(selectedAction){
//    case 'sendFile': 
//       forSelectedList.setVisible(true);  
//       forSelectedList.addItem(sFname +" " +sLname +" in " + clsName, 'forStuInClass');
//       forSelectedList.addItem(sFname +" " +sLname +" in all classes", 'forAllClasses');
//       forSelectedList.addItem(clsName, 'forClass');
//       if (clsPer != ''){
//         forSelectedList.addItem(clsPer+" in "+clsName, 'forPeriod');
//       }
//       
//       var openHtmlPicker = app.createServerHandler('htmlPicker');
//        
//       actionSettingsPanel.setVisible(true); 
//       actionSettingsPanel.add(app.createLabel("Choose a template to be copied."));     
//       actionSettingsPanel.add(app.createButton("Browse Drive").addClickHandler(openHtmlPicker));
//       actionSettingsPanel.add(app.createLabel("No File Chosen Yet!").setId("chosenFile")); 
//        
//    break;
    case 'remove':
  forSelectedList.setVisible(true);  
  forSelectedList.addItem(sFname +" " +sLname +" in " + clsName, 'forStuInClass');
  forSelectedList.addItem(sFname +" " +sLname +" in all classes", 'forAllClasses');
  forSelectedList.addItem(clsName, 'forClass');
    if (clsPer != ''){
    forSelectedList.addItem(clsPer+" in "+clsName, 'forPeriod');
    }
    break;
    case 'fixEmail':
        //for selected
        forSelectedList.setVisible(true);  
        forSelectedList.addItem(sFname +" " +sLname +" in " + clsName, 'forStuInClass');
        forSelectedList.addItem(sFname +" " +sLname +" in all classes", 'forAllClasses');
        
        actionSettingsPanel.setVisible(true); 
        actionSettingsPanel.add(app.createLabel("New Email"));
        actionSettingsPanel.add(app.createTextBox().setName("newEmail"));
        
        
    break;
    case 'rename':
        forSelectedList.setVisible(true);  
        forSelectedList.addItem(sFname +" " +sLname +" in " + clsName, 'forStuInClass');
        forSelectedList.addItem(sFname +" " +sLname +" in all classes", 'forAllClasses');
        
        actionSettingsPanel.setVisible(true); 
        actionSettingsPanel.add(app.createLabel("First Name"));
        actionSettingsPanel.add(app.createTextBox().setName("newFname"));
        actionSettingsPanel.add(app.createLabel("Last Name"));
        actionSettingsPanel.add(app.createTextBox().setName("newLname"));
        
    break;
//    default:
//      Browser.msgBox("You have selected a feature that is not yet available");
  }
    
  } // end if numbOfSel =1
  
   if (numbOfSel > 1){
    switch(selectedAction){
//    case 'sendFile':  
//        actionSettingsPanel.setVisible(true); 
//        actionSettingsPanel.add(app.createLabel("Choose a template to be copied."));
//        actionSettingsPanel.add(app.createTextBox().setName("newEmail"));       
//    break;
    }
   }
  
  return app;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
function runActionHandler(e){
  var ss = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var action = e.parameter.actionOnSelected;
  var forSelected = e.parameter.forSelected;
  var studentRowIndex =JSON.parse(e.parameter.selectedStudentsH); //array
  var numberSelected = studentRowIndex.length;
  var indices = returnIndices();
  var dataRange = ss.getDataRange().getValues();
  var sheet = getRosterSheet();
  var lastRow = ss.getLastRow();
  createLogSheet();

  
  switch(forSelected){
//    case 'forStuInClass': 
//      //No Change
//    var studentRowIndex =JSON.parse(e.parameter.selectedStudentsH); //array
//    var numberSelected = studentRowIndex.length;
//    break;
    case 'forAllClasses':
    studentRowIndex = returnStudentRows(studentRowIndex);
    numberSelected = studentRowIndex.length;

    break;
    case 'forClass':
      var clsFolderId = dataRange[studentRowIndex][indices.crfIdIndex];
      var clsPerFolderId = dataRange[studentRowIndex][indices.rsfIdIndex];
    break;
    case 'forPeriod':
      var clsFolderId = dataRange[studentRowIndex][indices.crfIdIndex];
      var clsPerFolderId = dataRange[studentRowIndex][indices.rsfIdIndex];
    break;
//    default:
//      Browser.msgBox("You have selected a feature that is not yet available");
  }
  
  
  switch(action) {
//    case 'sendFile':
//      
//      
//      
//      
//      ui.alert(action+ " to " + studentRowIndex + " " + forSelected);
//      break;
      
    case 'remove': 
      
      if (forSelected =='forStuInClass' || forSelected =='forAllClasses' || forSelected == ""){
      for (var i=0; i<numberSelected; i++){
      var studentRow = studentRowIndex[i];
      var sEmail = dataRange[studentRow][indices.sEmailIndex];
      var editFolder = dataRange[studentRow][indices.cefIdIndex];  
      var viewFolder = dataRange[studentRow][indices.cvfIdIndex]; 
      var AssignFolder = dataRange[studentRow][indices.dbfIdIndex]; 
 
        try{
          DriveApp.getFolderById(editFolder).removeEditor(sEmail);
          DriveApp.getFolderById(viewFolder).removeViewer(sEmail);
          DriveApp.getFolderById(AssignFolder).setTrashed(true);
        } catch(err) {}//Browser.msgBox(err.message);}
      sheet.getRange(studentRow+1, indices.sDropStatusIndex+1).setValue("Removed").setFontColor("red");
        moveToLog(studentRow+1,"action Remove: student", true); 
      SpreadsheetApp.flush();
      }
  } 
      if (forSelected == 'forClass'){
        DriveApp.getFolderById(clsFolderId).setTrashed(true);
        for (i=lastRow-1; i > 1;i--){
          var assignFolder = dataRange[i][indices.crfIdIndex]; 
          if (assignFolder == clsFolderId){
            moveToLog(i+1,"action Remove: Class", true); 
          } 
        }
      }
      if (forSelected == 'forPeriod'){
        if (clsFolderId != clsPerFolderId){
        DriveApp.getFolderById(clsPerFolderId).setTrashed(true);
        }
        for (i=lastRow-1; i > 1;i--){
          var assignFolder = dataRange[i][indices.rsfIdIndex]; 
          if (assignFolder == clsPerFolderId && clsFolderId != clsPerFolderId) {
            moveToLog(i+1,"action Remove: Period", true); 
          } 
        }
      } // end for per

      ui.alert("Remove operation completed");
      break;
    case "notSelected":
      Browser.msgBox("Please select an action");
      break;
      
    case 'fixEmail': 
      var newEmail = e.parameter.newEmail; 
      var validEmail = true;
      
      if (newEmail == ""){
        ui.alert("Both the first and last name fields must be filled out");
        break;
      }
      
      if (forSelected =='forStuInClass'|| forSelected =='forAllClasses'){
        for (var i=0; i<numberSelected; i++){
         var studentRow = studentRowIndex[i];
         var oldEmail = dataRange[studentRow][indices.sEmailIndex];
         var editFolder = dataRange[studentRow][indices.cefIdIndex];  
         var viewFolder = dataRange[studentRow][indices.cvfIdIndex]; 
         var AssignFolder = dataRange[studentRow][indices.dbfIdIndex];
//          Browser.msgBox(newEmail+" "+oldEmail+" "+editFolder+" "+viewFolder+" "+AssignFolder);
         
          // Add new student
          try{
          DriveApp.getFolderById(editFolder).addEditor(newEmail);
          DriveApp.getFolderById(viewFolder).addViewer(newEmail);
          DriveApp.getFolderById(AssignFolder).addEditor(newEmail);
          } catch(err) { 
            validEmail = false;
            break;}
            
          // Remove old student
          DriveApp.getFolderById(editFolder).removeEditor(oldEmail);
          DriveApp.getFolderById(viewFolder).removeViewer(oldEmail);
          DriveApp.getFolderById(AssignFolder).removeEditor(oldEmail);
          
          // Update spreadsheet 
          moveToLog(studentRow+1, "Action Fix Email", false); 
          sheet.getRange(studentRow+1,indices.sEmailIndex+1).setValue(newEmail);
          sheet.getRange(studentRow+1, indices.sDropStatusIndex+1).setValue("Email Changed").setFontColor("red");
          
        }  
      }
      
      if ( validEmail == false){           
            ui.alert("Not a valid email address");
      } else {
            ui.alert("Fixed Email operation completed");}
      
      break;
      
    case 'rename':
      var newFirstName = e.parameter.newFname; 
      var newLastName = e.parameter.newLname;
      
      if (newFirstName == "" || newLastName == ""){
        ui.alert("Both the first and last name fields must be filled out");
        break;
      }
      
      if (forSelected =='forStuInClass'|| forSelected =='forAllClasses'){
        for (var i=0; i<numberSelected; i++){
         var studentRow = studentRowIndex[i];
         var AssignFolder = dataRange[studentRow][indices.dbfIdIndex]; 
         var className = dataRange[studentRow][indices.clsNameIndex];
         var dropBoxLabel = this.labels().dropBox;
          DriveApp.getFolderById(AssignFolder).setName(newLastName+ ', '+newFirstName+' - '+className+' - '+ dropBoxLabel);
//          copy row to log, and update sheet
        moveToLog(studentRow+1, "Action Rename", false);
        sheet.getRange(studentRow+1,indices.sFnameIndex+1).setValue(newFirstName);
        sheet.getRange(studentRow+1,indices.sLnameIndex+1).setValue(newLastName);
        SpreadsheetApp.flush();
       }
      }

      ui.alert("Student renamed to: "+ newFirstName + " " +newLastName);
      break;
      
    default:
      Browser.msgBox("You have selected a feature that is not yet available");
  }
  
} // end runActionHandler()



//////////////////////////////////////////////////////////////////////////////////////////////////
function getSelectedHandler(e) {
  var app = UiApp.getActiveApplication();
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSheet();
  var ssAR = SpreadsheetApp.getActiveRange();
  var numbOfSel = ssAR.getNumRows();
  var indices = returnIndices();
  var dataRange = ss.getDataRange().getValues();
  var selectedStart = ssAR.getRow();
  var selectedLast = ssAR.getLastRow();
  
  var quesInfoPanel = app.getElementById("quesInfoPanel");
  quesInfoPanel.clear();
  var actionOnSelectedPanel = app.getElementById("actionOnSelected");
  actionOnSelectedPanel.setVisible(true).clear();
  var selectedStudentsH = app.getElementById("selectedStudentsH");
  var selectedStudents =[];
  var forSelectedList = app.getElementById("forSelected");
  forSelectedList.setVisible(false).clear();
  var actionSettingsPanel = app.getElementById("actionSettingsPanel");
  actionSettingsPanel.setVisible(false).clear();
  var actionButton = app.getElementById("actionButton");
  actionButton.setVisible(true);
  var closeButton = app.getElementById("closeButton");
  closeButton.setVisible(false);
  
  
// Check if selected is availible to be ran  
  if (indices.sDropStatusIndex == -1) { 
   var questionInfo = app.createHTML("Classes have not yet been created").setStyleAttribute('backgroundColor', '#CACACA').setStyleAttribute('padding', '5px').setId('description');
    closeButton.setVisible(true);
    actionOnSelectedPanel.setVisible(false);
    actionButton.setVisible(false);
    quesInfoPanel.add(questionInfo);
   return app;
  }
  
  
//  var firstSelectedValue = ssAR.getValues()[0];
  if (ssAR.getRowIndex() == 1 || ssAR.getRowIndex() == 2) { 
   var questionInfo = app.createHTML("You have selected a header row.").setStyleAttribute('backgroundColor', '#CACACA').setStyleAttribute('padding', '5px').setId('description');
    closeButton.setVisible(true);
    actionOnSelectedPanel.setVisible(false);
    actionButton.setVisible(false);
    quesInfoPanel.add(questionInfo);
   return app;
  }
  

  
  
  
  //get selected
  if (numbOfSel == 1){
    try {
   var rIndex = ssAR.getRowIndex()-1;
   var sFname = dataRange[rIndex][indices.sFnameIndex];
   var sLname = dataRange[rIndex][indices.sLnameIndex]; 
   var sEmail = dataRange[rIndex][indices.sEmailIndex];  
   var clsName = dataRange[rIndex][indices.clsNameIndex]; 
   var clsPer = dataRange[rIndex][indices.clsPerIndex]; 
   selectedStudents.push(rIndex);
          } catch(err){ 
    var questionInfo = app.createHTML("You have selected an invalid row. Selected row must contain data.").setStyleAttribute('backgroundColor', '#CACACA').setStyleAttribute('padding', '5px').setId('description');
      closeButton.setVisible(true);
      actionOnSelectedPanel.setVisible(false);
      actionButton.setVisible(false);
      quesInfoPanel.add(questionInfo);
    return app;
    }
      
      
      
   var html = "Question Info";
   html += "<ul><li>Student: " +sFname +" " +sLname + "</li>";
   html += "<li>Class: " + clsName + "</li>";
   html += "<li>Period: "+ clsPer + "</li>";
   html += "<li>Row: "+ (rIndex+1) +"</li>";
    
   var questionInfo = app.createHTML(html).setStyleAttribute('backgroundColor', '#CACACA').setStyleAttribute('padding', '5px').setId('description');
 
    
      //With Selected
   actionOnSelectedPanel.addItem('', 'notSelected');
//   actionOnSelectedPanel.addItem('Send File', 'sendFile');
   actionOnSelectedPanel.addItem('Remove', 'remove');
   actionOnSelectedPanel.addItem('Fix Student Email', 'fixEmail');
   actionOnSelectedPanel.addItem('Rename Student', 'rename');  
    

  }
  if (numbOfSel > 1){
      var questionInfo = app.createScrollPanel().setStyleAttribute('border', '1px solid grey').setHeight("100px");
      var selectedStart = ssAR.getRow()-1;
      var selectedLast = ssAR.getLastRow();
      var vertical = app.createVerticalPanel();

    try{
    for (var i=selectedStart; i < selectedLast; i++) {
         var sFname = dataRange[i][indices.sFnameIndex];
         var sLname = dataRange[i][indices.sLnameIndex];
         vertical.add(app.createLabel(sFname+" "+sLname));
         selectedStudents.push(i);
        }
    } catch(err){ 
    var questionInfo = app.createHTML("You have selected an invalid row. All selected rows must contain data.").setStyleAttribute('backgroundColor', '#CACACA').setStyleAttribute('padding', '5px').setId('description');
      closeButton.setVisible(true);
      actionOnSelectedPanel.setVisible(false);
      actionButton.setVisible(false);
      quesInfoPanel.add(questionInfo);
    return app;
    }
    questionInfo.add(vertical);
    
      actionOnSelectedPanel.addItem('', 'notSelected');
//      actionOnSelectedPanel.addItem('Send File', 'sendFile');
      actionOnSelectedPanel.addItem('Remove', 'remove');
  } // End if multiple selected


  selectedStudentsH.setValue(JSON.stringify(selectedStudents));
  quesInfoPanel.add(questionInfo);

  return app;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////
// function fixEmailClick(e) {
//   var app = UiApp.getActiveApplication();
//   app.getElementById('button').setText('Clicked!');
//   return app;
// }


////////////////////////////////////////////////////////////////////////////////////////////////////
function returnStudentRows(studentRow){
  var ss = SpreadsheetApp.getActiveSheet();
  var dataRange = ss.getDataRange().getValues();
  var sheet = getRosterSheet();
  var indices = returnIndices();
  var sEmail = dataRange[studentRow][indices.sEmailIndex];
  var studentRows =[];
  
  for (var i = 3; i < dataRange.length; i++) { 
    var rowValue = sheet.getRange(i, indices.sEmailIndex+1).getValue();
    if(rowValue == sEmail){
       studentRows.push(i-1);
    }
  }
  return studentRows;
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////
function moveToLog(row, identifier, removeFromSheet){
// row = Row to move, itendifier = The text that goes into column A, removeFromSheet = true or false
// if (!row){ row = 3;}  //Test Row
// if (!identifier){identifier = 'test';}
// if (!removeFromSheet){removeFromSheet = false;}
 var sheet = getRosterSheet();
 var logSheet = getLogSheet();
 var sheetLastC = sheet.getLastColumn();
 var sourceToCopy = sheet.getRange(row,1,1,sheetLastC);
 var lastRow = logSheet.getLastRow();
 var destinationRange =  logSheet.getRange(lastRow+1,2,1,sheetLastC)
 logSheet.getRange(lastRow+1,1).setValue(identifier);
 sourceToCopy.copyTo(destinationRange);
  if (removeFromSheet == true){
  sheet.deleteRow(row);
  }
  SpreadsheetApp.flush();
  return;
 }
  



//  indices.sFnameIndex = headers.indexOf(defaultHeadingsI.sFname);
//  indices.sLnameIndex = headers.indexOf(defaultHeadingsI.sLname);
//  indices.sEmailIndex = headers.indexOf(defaultHeadingsI.sEmail);
//  indices.clsNameIndex = headers.indexOf(defaultHeadingsI.clsName);
//  indices.clsPerIndex = headers.indexOf(defaultHeadingsI.clsPer);
//  indices.tEmailIndex = headers.indexOf(defaultHeadingsI.tEmail);
//  
//  indices.sDropStatusIndex = headers.indexOf(defaultIDsI.status);
//  indices.dbfIdIndex = headers.indexOf(defaultIDsI.assignmentFID);
//  indices.crfIdIndex = headers.indexOf(defaultIDsI.classRootFID);
//  indices.cvfIdIndex = headers.indexOf(defaultIDsI.classViewFID);
//  indices.cefIdIndex = headers.indexOf(defaultIDsI.classEditFID);
//  indices.rsfIdIndex = headers.indexOf(defaultIDsI.rootStudentFID);
//  indices.tfIdIndex = headers.indexOf(defaultIDsI.teacherFID);
//  indices.tShareStatusIndex = headers.indexOf(defaultIDsI.teacherShareStatus);

