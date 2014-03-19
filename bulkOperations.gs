function bulkOperationsUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var activeSheet = ss.getActiveSheet();
  var mode = PropertiesService.getDocumentProperties().getProperty('mode');
  var activeSheetId = activeSheet.getSheetId();
  var rosterSheet = getRosterSheet();
  var rosterSheetId = rosterSheet.getSheetId();
  var labelObject = this.labels();
  var lang = properties.lang;
  if (activeSheetId==rosterSheetId) {
    var app = UiApp.createApplication().setTitle("Perform bulk student operations").setHeight(450);
    var waitingPanel = app.createVerticalPanel().setId('waitingImage');
    var waitingImageUrl = "https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/goldballs.gif?attachauth=ANoY7coUFQKLFJRBrV-yRwgZ3p6jVsn_UbJIlzFstZAAyF1r6Xj8wCNG6yjkbeOxVf80Oo_55TUl-VvXL0OtztWjaN9_wF7pclOhemgkGWvYYSJSWLhJzp1tqMdJDoDYVaK4cpOHO1jCJDTRUmt3jNpZMo0xboBIi9W_yTbZW-8kY8nDJ3nWDrkHbmZfSPy1fh7qitwMR3kANmtQRq2EfYTJbzx56bMcFCEc4Eq3zvrirBGHllBdPTeFspBZfj5ew3e2Ffmx0phu&attredirects=0";
    var waitingImage = app.createImage(waitingImageUrl);
    waitingPanel.setStyleAttribute('position', 'absolute')
    .setWidth('200px')
    .setStyleAttribute('backgroundColor', 'white')
    .setStyleAttribute('top', '75px')
    .setStyleAttribute('left', '150px');
    waitingPanel.add(app.createLabel('Please do not edit the roster sheet until script is finished operating on student rows.'));
    waitingPanel.add(waitingImage).setVisible(false);
    var panel = app.createVerticalPanel();
    var dataRange = rosterSheet.getDataRange();
    var indices = returnIndices(dataRange, labelObject);
    var activeRange = activeSheet.getActiveRange();
    var topRow = activeRange.getRow();
    var numRows = activeRange.getNumRows();
    var values = [];
    if (topRow!=1) {
      var values = rosterSheet.getRange(topRow, 1, numRows, rosterSheet.getLastColumn()).getValues();
    } else {
      var noneSelected = app.createLabel("You have not highlighted any rows in the spreadsheet. Please return to the roster sheet and highlight students, and then try the bulk operations menu item again.");
    }
    if (activeSheet.getSheetId()!=PropertiesService.getDocumentProperties().getProperty('sheetId')) {
      var noneSelected = app.createLabel("You were not in the roster sheet when you selected the bulk operations menu item.  Please return to the roster sheet and highlight students, and then try the bulk operations menu item again.");
    }
    var topGrid = app.createGrid(1, 4);
    topGrid.setWidget(0, 0, app.createLabel('First Name')).setStyleAttribute(0, 0, 'width','150px').setStyleAttribute('backgroundColor', '#e5e5e5')
    .setWidget(0, 1, app.createLabel('Last Name')).setStyleAttribute(0, 1, 'width','150px')
    .setWidget(0, 2, app.createLabel(labelObject.class)).setStyleAttribute(0, 2, 'width','150px')
    .setWidget(0, 3, app.createLabel(labelObject.period)).setStyleAttribute(0, 3, 'width','150px');
    var grid = app.createGrid(values.length, 5);
    var scrollPanel = app.createScrollPanel().setHeight("200px").setStyleAttribute('border', '1px solid grey');
    var studentObjects = [];
    for (var i=0; i<values.length; i++ ) {
      studentObjects[i] = new Object();
      studentObjects[i]['sFName'] = values[i][indices.sFnameIndex];
      studentObjects[i]['sLName'] = values[i][indices.sLnameIndex];
      studentObjects[i]['sEmail'] = values[i][indices.sEmailIndex];
      studentObjects[i]['dbfId'] = values[i][indices.dbfIdIndex]; 
      studentObjects[i]['cvfId'] = values[i][indices.cvfIdIndex]; 
      studentObjects[i]['cefId'] = values[i][indices.cefIdIndex];
      studentObjects[i]['rsfId'] = values[i][indices.rsfIdIndex];
      studentObjects[i]['crfId'] = values[i][indices.crfIdIndex];
      studentObjects[i]['tfId'] = values[i][indices.tfIdIndex];
      studentObjects[i]['clsName'] = values[i][indices.clsNameIndex];
      studentObjects[i]['clsPer'] = values[i][indices.clsPerIndex];
      studentObjects[i]['tEmail'] = values[i][indices.tEmailIndex];
      studentObjects[i]['row'] = topRow + i;
      var studentObjectString = Utilities.jsonStringify(studentObjects[i]);
      var bgColor = 'whiteSmoke';
      if (i % 2 === 0) {
        bgColor = 'white';
      }
      grid.setWidget(i, 0, app.createLabel(values[i][indices.sFnameIndex])).setStyleAttribute(i, 0, 'width','150px').setStyleAttribute(i, 0, 'backgroundColor',bgColor).setStyleAttribute(i, 0, 'borderTop', '1px solid #e5e5e5')
      .setWidget(i, 1, app.createLabel(values[i][indices.sLnameIndex])).setStyleAttribute(i, 1, 'width','150px').setStyleAttribute(i, 1, 'backgroundColor',bgColor).setStyleAttribute(i, 1, 'borderTop', '1px solid #e5e5e5')
      .setWidget(i, 2, app.createLabel(values[i][indices.clsNameIndex])).setStyleAttribute(i, 2, 'width','150px').setStyleAttribute(i, 2, 'backgroundColor',bgColor).setStyleAttribute(i, 2, 'borderTop', '1px solid #e5e5e5');
      if (values[i][indices.clsPerIndex]!='') {
        grid.setWidget(i, 3, app.createLabel(values[i][indices.clsPerIndex])).setStyleAttribute(i, 3, 'width','150px').setStyleAttribute(i, 3, 'backgroundColor',bgColor).setStyleAttribute(i, 3, 'borderTop', '1px solid #e5e5e5');
      } else {
        grid.setWidget(i, 3, app.createLabel("")).setStyleAttribute(i, 3, 'width','150px').setStyleAttribute(i, 3, 'backgroundColor',bgColor).setStyleAttribute(i, 3, 'borderTop', '1px solid #e5e5e5');
      }
      grid.setWidget(i, 4, app.createHidden('student-'+i).setValue(studentObjectString));
    }
    panel.add(app.createHidden('numStudents').setValue(numRows))
    panel.add(topGrid);
    scrollPanel.add(grid);
    if (values.length==0) {
      grid.resize(1, 1);
      grid.setWidget(0, 0, noneSelected);
    }
    panel.add(scrollPanel);
    
    var operationSelectGrid = app.createGrid(2, 1).setId('operationSelectGrid');
    var operationSelectList = app.createListBox().setName('operation');
    var changeHandler = app.createServerHandler('refreshDescriptor').addCallbackElement(operationSelectList);
    operationSelectList.addItem('Remove from ' + labelObject.class, 'remove')
    .addItem('Add teacher to ' + labelObject.class, 'add teacher')
    .addItem('Add student aide','add aide')
    .addItem('Move ' + labelObject.dropBox, 'move');
    operationSelectList.addChangeHandler(changeHandler);
    var operationDescriptor = app.createLabel("Removing students will archive their " + labelObject.class + " " + labelObject.dropBox + " and remove them from teacher " + labelObject.dropBox + ", " + labelObject.class + " view, and " + labelObject.class + " edit folders.").setId('operationDescriptor');
    var operationSettingsPanel = app.createVerticalPanel().setId('operationSettingsPanel');
    operationSettingsPanel.add(operationDescriptor)
    var operationScroll = app.createScrollPanel(operationSettingsPanel).setHeight("140px").setWidth("100%").setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('margin', '8px');
    operationSelectGrid.setWidget(0, 0, operationSelectList)
    .setWidget(1, 0, operationScroll);
    panel.add(operationSelectGrid);
    
    var button = app.createButton('Run operation');
    var buttonServerHandler = app.createServerHandler('bulkOperateOnStudents').addCallbackElement(panel);
    var buttonClientHandler = app.createClientHandler().forTargets(waitingPanel).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.2').forTargets(button).setEnabled(false);
    button.addClickHandler(buttonServerHandler).addClickHandler(buttonClientHandler);
    panel.add(button);
    app.add(panel);
    app.add(waitingPanel);
    ss.show(app);
    return app;
  } else {
    Browser.msgBox("You are not currently in the roster sheet. Please return to the roster sheet and try again.");
  }
}

function refreshDescriptor(e) {
  var app = UiApp.getActiveApplication();
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var lang = properties.lang;
  var operationSettingsPanel = app.getElementById('operationSettingsPanel');
  var descriptorLabel = app.getElementById('operationDescriptor');
  var operation = e.parameter.operation;
  var labelObject = this.labels();
  switch(operation)
  {
    case 'remove':
      operationSettingsPanel.clear();
      descriptorLabel.setText("Removing students will archive their " + labelObject.dropBox + " and remove them from teacher " + labelObject.dropBox + " folder, " + labelObject.class + " view, and " + labelObject.class + " edit folders.").setStyleAttribute("margin","5px");
      operationSettingsPanel.add(descriptorLabel);
      break;
    case 'add teacher':
      operationSettingsPanel.clear();
      descriptorLabel.setText("Teacher will be added to all relevant " + labelObject.class + " folders and to all " + labelObject.dropBoxes + " in any of the " + labelObject.classes + " selected.").setStyleAttribute("margin","5px");
      operationSettingsPanel.add(descriptorLabel);
      operationSettingsPanel.add(app.createLabel("Teacher email address").setStyleAttribute("margin","5px"));
      operationSettingsPanel.add(app.createTextBox().setName('tEmail').setStyleAttribute("margin","5px"));
      break;
    case 'add aide':
      operationSettingsPanel.clear();
      descriptorLabel.setText("School aide will be added only to the relevant " + labelObject.dropBox + " as editor and to class edit and class view folders with the same privileges as student.").setStyleAttribute("margin","5px");
      operationSettingsPanel.add(descriptorLabel).setStyleAttribute("marginTop","5px");
      operationSettingsPanel.add(app.createLabel("Student aide email address").setStyleAttribute("margin","5px"));
      operationSettingsPanel.add(app.createTextBox().setName('tEmail').setStyleAttribute("margin","5px"));
      break;
    case 'move':
      operationSettingsPanel.clear();
      var sheet = getRosterSheet();
      var dataRange = sheet.getDataRange().getValues();
      var indices = returnIndices(dataRange, labelObject);
      var uniqueClasses = getUniqueClassPeriodObjects(dataRange, indices.clsNameIndex, indices.clsPerIndex, indices.rsfIdIndex, labelObject);
      descriptorLabel.setText("Moving " + labelObject.dropBoxes + " will preserve all work and place them in a new " + labelObject.class + " and " + labelObject.period + " " + labelObject.dropBox + " root, changing teacher and student access rights as necessary.").setStyleAttribute("margin","5px");
      operationSettingsPanel.add(descriptorLabel);
      operationSettingsPanel.add(app.createLabel("Destination " + labelObject.class + " / " + labelObject.period).setStyleAttribute("margin","5px"));
      var sectionSelector = app.createListBox().setName('destinationRsfId').setStyleAttribute("margin","5px");
      for (var i=0; i<uniqueClasses.length; i++) {
        sectionSelector.addItem(uniqueClasses[i].classPer, uniqueClasses[i].classPer+"||"+uniqueClasses[i].rsfId);
      }
      operationSettingsPanel.add(sectionSelector);
      break;
  }
  return app;
}


function bulkOperateOnStudents(e) {
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var lang = properties.lang;
  var app = UiApp.getActiveApplication();
  var operation = e.parameter.operation;
  var sheet = getRosterSheet();
  var dataRange = sheet.getDataRange();
  var labelObject = this.labels();
  var indices = returnIndices(dataRange, labelObject);
  var numStudents = parseInt(e.parameter.numStudents);
  var driveRoot = call(function() {return DriveApp.getRootFolder();});
  var studentObjects = [];
  for (var i=0; i<numStudents; i++) {
    studentObjects[i] = Utilities.jsonParse(e.parameter['student-'+i]);
  }
  
  //begin switch case
  switch(operation) {
    case 'remove':
      if (!properties.topDBArchiveFolderId) { //creates an archive folder if it doesn't yet exist.
        properties.topDBArchiveFolderId = call(function() { return DriveApp.createFolder('gClassFolders Archived Student ' + labelObject.dropBoxes).getId();});
        PropertiesService.getDocumentProperties().setProperties(properties);
      }
      if (DriveApp.getFolderById(properties.topDBArchiveFolderId).isTrashed()==true) { //Creates a new folder if old folder is trashed.
        properties.topDBArchiveFolderId = call(function() { return DriveApp.createFolder('gClassFolders Archived Student ' + labelObject.dropBoxes).getId();});
        PropertiesService.getDocumentProperties().setProperties(properties);
      }
      topDBArchiveFolderId = properties.topDBArchiveFolderId;
      var date = Utilities.formatDate(new Date(), timeZone, "M/d/yy");
      for (var i=0; i<studentObjects.length; i++) {
        var status = "You may delete this row. ";
        var sFName = studentObjects[i]['sFName'];
        var sLName = studentObjects[i]['sLName'];
        var sEmail = studentObjects[i]['sEmail'].toLowerCase().replace(/\s/g, "");
        var dbfId = studentObjects[i]['dbfId'];
        var cvfId = studentObjects[i]['cvfId'];
        var cefId = studentObjects[i]['cefId']; 
        var rsfId = studentObjects[i]['rsfId'];
        var tfId = studentObjects[i]['tfId'];
        var row =  studentObjects[i]['row']; 
        //remove rights from class edit, class view
        try {
          call(function(){DriveApp.getFolderById(cvfId).removeViewer(sEmail);});
          call(function(){DriveApp.getFolderById(cefId).removeEditor(sEmail);});
          status += sEmail + " removed from class view and edit folders. ";
        } catch(err) {
          status += "Error removing " + sEmail + " from class view and class edit folders. ";
        }
        try {
          var topDBArchiveFolder = call(function(){return DriveApp.getFolderById(topDBArchiveFolderId);});
          var dropboxFolder = call(function(){return DriveApp.getFolderById(dbfId);});
          call(function(){ topDBArchiveFolder.addFolder(dropboxFolder);});
          var currentDbName = dropboxFolder.getName();
          var dropboxRoot = call(function(){return DriveApp.getFolderById(rsfId);});
          call(function(){dropboxRoot.removeFolder(dropboxFolder);});
          call(function(){dropboxFolder.setName(currentDbName + " - Removed from class by " + this.userEmail + ", " + date);});
          status += sEmail + " dropbox folder moved to student archive folder. ";
        } catch(err) {
          Browser.msgBox(err.message);
          return;
          status += "Error moving " + sEmail + " dropbox folder to \"gClassFolders - Removed Students\" folder. ";
        }
        sheet.getRange(row, indices.sDropStatusIndex+1).setValue(status).setFontColor("red");
        SpreadsheetApp.flush();
      }
      app.close();
      return app;
      break;
    case 'add teacher':
      var sheet = getRosterSheet();
      var dataRange = sheet.getDataRange().getValues();
      var indices = returnIndices(dataRange, labelObject);
      var tEmail = e.parameter.tEmail;
      tEmail = tEmail.replace(/\s/g, "").toLowerCase();
      for (var i=0; i<studentObjects.length; i++) {
        var idsProcessed = [];
        var rsfsProcessed = [];
        var sEmail = studentObjects[i]['sEmail'].replace(/\s/g, "").toLowerCase();
        var dbfId = studentObjects[i]['dbfId'];
        var cvfId = studentObjects[i]['cvfId'];
        var cefId = studentObjects[i]['cefId']; 
        var rsfId = studentObjects[i]['rsfId'];
        var crfId = studentObjects[i]['crfId'];
        var tfId = studentObjects[i]['tfId'];
        var clsName = studentObjects[i]['clsName'];
        var clsPer = studentObjects[i]['clsPer'];
        var sEmail = studentObjects[i]['sEmail'];
        var row =  studentObjects[i]['row']; 
        var status = tEmail + "\n";
        var newTEmails = studentObjects[i]['tEmail'].replace(/\s/g, "").split(",");
        if (idsProcessed.indexOf(crfId)==-1) {
          try {
            call(function(){DriveApp.getFolderById(crfId).addEditor(tEmail);});
            idsProcessed.push(crfId);
            status += "added to " + clsName + " root folder,\n";
          } catch(err) {
            status += "error adding as editor on " + clsName + " root folder: " + err.message + "\n";
          }
          if (idsProcessed.indexOf(cefId)==-1) {
            try {
              call(function(){DriveApp.getFolderById(cefId).addEditor(tEmail);});
              idsProcessed.push(cefId);
              status += "added to " + clsName + " edit folder,\n";
            } catch(err) {
              status += "error adding as editor on " + clsName + " edit folder: " + err.message + "\n";
            }
          }
          newTEmails.push(tEmail);
          newTEmails = newTEmails.join(",");
          var comment = tEmail + " added as teacher to " + clsName;
          if (clsPer!='') {
            comment += labelObject.period + " " + clsPer;
          } 
          comment += " by " + this.userEmail + " on " + Utilities.formatDate(new Date(), timeZone, 'M/d/yy');
          var classRowNums = getClassRowNumsFromCRF(dataRange, indices, crfId);
          for (var k=0; k<classRowNums.length; k++) {
            sheet.getRange(classRowNums[k], indices.tEmailIndex+1).setValue(newTEmails).setFontColor("blue").setComment(comment);
          }
          SpreadsheetApp.flush();
        }
        if (idsProcessed.indexOf(cvfId)==-1) {
          try {
            call(function(){DriveApp.getFolderById(cvfId).addEditor(tEmail);});
            idsProcessed.push(cvfId);
            status += "added to " + clsName + " view folder,\n";
          } catch(err) {
            
            status += "error adding as editor on " + clsName + " view folder: " + err.message + ",\n";
          }
        }
        if (idsProcessed.indexOf(rsfId)==-1) {
          try{
            call(function(){DriveApp.getFolderById(rsfId).addEditor(tEmail);});
            idsProcessed.push(rsfId);
            rsfsProcessed.push(rsfId);
            status += "added to " + clsName + " ";
            if (clsPer!='') {
              status += labelObject.period + clsPer; 
            }
            status += labelObject.dropBox + ", \n";
          } catch(err) {
            status += "error adding to " + clsName + " ";
            if (clsPer!='') {
              status += labelObject.period + " " + clsPer; 
            }
            status += labelObject.dropBox + " folder,\n";
          }
        }
        if (idsProcessed.indexOf(tfId)==-1) {
          try{
            call(function(){DriveApp.getFolderById(tfId).addEditor(tEmail);});
            idsProcessed.push(tfId);
            status += "added to " + clsName + " teacher folder,\n";
          } catch(err) {
            status += "error adding as editor on " + clsName + " teacher folder: " + err.message + "\n";
          }
        }
      }
      app.close();
      Browser.msgBox(status);
      return app;
      break;
    case 'add aide':
      var sheet = getRosterSheet();
      var dataRange = sheet.getDataRange().getValues();
      var indices = returnIndices(dataRange, labelObject);
      var tEmail = e.parameter.tEmail;
      tEmail = tEmail.replace(/\s/g, "");
      for (var i=0; i<studentObjects.length; i++) {
        var idsProcessed = [];
        var sEmail = studentObjects[i]['sEmail'].replace(/\s/g, "").toLowerCase();
        var dbfId = studentObjects[i]['dbfId'];
        var cvfId = studentObjects[i]['cvfId'];
        var cefId = studentObjects[i]['cefId']; 
        var rsfId = studentObjects[i]['rsfId'];
        var crfId = studentObjects[i]['crfId'];
        var clsName = studentObjects[i]['clsName'];
        var clsPer = studentObjects[i]['clsPer'];
        var sEmail = studentObjects[i]['sEmail'];
        var row =  studentObjects[i]['row']; 
        var status = tEmail + "\n";
        var newTEmails = studentObjects[i]['tEmail'].replace(/\s/g, "").split(",");
        if (idsProcessed.indexOf(cvfId)==-1) {
          try {
            call(function(){DriveApp.getFolderById(cvfId).addViewer(tEmail);});
            idsProcessed.push(cvfId);
            status += "added to " + clsName + " root folder,\n";
          } catch(err) {
            status += "error adding as editor on " + clsName + " root folder:" + err.message + "\n";
          }
        }
        if (idsProcessed.indexOf(cefId)==-1) {
          try {
            call(function(){DriveApp.getFolderById(cefId).addEditor(tEmail);});
            idsProcessed.push(cefId);
            status += "added to " + clsName + " edit folder,\n";
          } catch(err) {
            status += "error adding as editor on " + clsName + " edit folder: " + err.message + "\n";
          }
        }
        if (idsProcessed.indexOf(cvfId)==-1) {
          try {
            call(function(){DriveApp.getFolderById(cvfId).addEditor(tEmail);});
            idsProcessed.push(cvfId);
            status += "added to " + clsName + " view folder,\n";
          } catch(err) {
            status += "error adding as editor on " + clsName + "view folder: " + err.message + "\n";
          }
        }
        try {
          call(function(){DriveApp.getFolderById(dbfId).addEditor(tEmail);});
          newTEmails.push(tEmail);
        } catch(err) {
          status += "error adding as editor on " + sEmail + " student dropbox folder: " + err.message + "\n";
        }
        newTEmails = newTEmails.join(",");
        var comment = tEmail + " added as student aide by " + this.userEmail + " on " + Utilities.formatDate(new Date(), timeZone, 'M/d/yy');
        sheet.getRange(row, indices.tEmailIndex+1).setValue(newTEmails).setFontColor("green").setComment(comment);
        SpreadsheetApp.flush();
      }
      app.close();
      Browser.msgBox(status);
      return app;
      break;
    case 'move':
      var sheet = getRosterSheet();
      var dataRange = sheet.getDataRange().getValues();
      var indices = returnIndices(dataRange, labelObject);
      
      //load the existing root folder ID info for students and teachers 
      var destinationRsfId = e.parameter.destinationRsfId.split("||")[1];
      var destinationClass = e.parameter.destinationRsfId.split("||")[0].split(" " + labelObject.period + " ")[0];
      var destinationPer = e.parameter.destinationRsfId.split("||")[0].split(" " + labelObject.period + " ")[1];
      var destinationCrfObject = getRootClassFoldersByRSF(dataRange, destinationRsfId, indices.rsfIdIndex, indices.crfIdIndex, indices.cefIdIndex, indices.cvfIdIndex);
      var destinationCrfId = destinationCrfObject.crfId;
      for (var i=0; i<studentObjects.length; i++) {
        var idsProcessed = [];
        var sEmail = studentObjects[i]['sEmail'];
        var dbfId = studentObjects[i]['dbfId'];
        var cvfId = studentObjects[i]['cvfId'];
        var cefId = studentObjects[i]['cefId']; 
        var rsfId = studentObjects[i]['rsfId'];
        var crfId = studentObjects[i]['crfId'];
        var clsName = studentObjects[i]['clsName'];
        var clsPer = studentObjects[i]['clsPer'];
        var sEmail = studentObjects[i]['sEmail'];
        var row =  studentObjects[i]['row']; 
        var status = "";
        //fix to include try catch, etc.
        var rootStuFolder = call(function() {return DriveApp.getFolderById(rsfId);});
        var dropBoxFolder = call(function() {return DriveApp.getFolderById(dbfId);});
        var destRootStuFolder = call(function() { return DriveApp.getFolderById(destinationRsfId);});
        var destTeachers = getTeacherEmailsByRSF(dataRange, destinationRsfId, indices.rsfIdIndex, indices.tEmailIndex);
        var destTeachers = destTeachers.tEmails.replace(/\s/g, "").split(",");
        var destRsfUrl = destRootStuFolder.getUrl();
        call(function() {destRootStuFolder.addFolder(dropBoxFolder);});
        call(function() {rootStuFolder.removeFolder(dropBoxFolder);});
      
        for (var k=0; k<destTeachers.length; k++) {
          call(function(){dropBoxFolder.addEditor(destTeachers[k]);});
        }
        var comment = "Moved from " + clsName ;
        if ((clsPer)&&(clsPer!='')) {
          comment += " " + labelObject.period + " " + clsPer;
        }
        comment += " to " + destinationClass;
        if ((destinationPer)&&(destinationPer!='')) {
          comment += " " + labelObject.period + " " + destinationPer;
        }
        comment += " by " + this.userEmail + " on " + Utilities.formatDate(new Date(), timeZone, 'M/d/yy');          
        if ((destinationPer)&&(destinationPer!='')) {
          sheet.getRange(row,indices.clsPerIndex+1).setValue(destinationPer).setFontColor("blue").setComment(comment);
        }
        sheet.getRange(row,indices.tEmailIndex+1).setValue(destTeachers.join(",")).setFontColor("blue");
        sheet.getRange(row,indices.rsfIdIndex+1).setFormula('=hyperlink("' + destRsfUrl + '";"' + destinationRsfId + '")');
        if (destinationCrfId!=crfId) { //Need to remove student rights on old cef and add to new cvf
          call(function(){DriveApp.getFolderById(cefId).removeEditor(sEmail);});
          call(function(){DriveApp.getFolderById(cvfId).removeEditor(sEmail);});
          var destCef = call(function(){ return DriveApp.getFolderById(destinationCrfObject.cefId);});
          call(function() {destCef.addEditor(sEmail);});
          var destCvf = call(function(){ return DriveApp.getFolderById(destinationCrfObject.cvfId);});
          call(function() {destCvf.addViewer(sEmail);});
          sheet.getRange(row,indices.clsNameIndex+1).setValue(destinationClass).setFontColor("blue").setComment(comment);
          sheet.getRange(row,indices.cefIdIndex+1).setFormula('=hyperlink("' + destCef.getUrl() + '";"' + destinationCrfObject.cefId + '")');
          sheet.getRange(row,indices.cvfIdIndex+1).setFormula('=hyperlink("' + destCvf.getUrl() + '";"' + destinationCrfObject.cvfId + '")');
        } 
      }
      app.close();
      return app;
      break;
      
      
    default:
      Browser.msgBox("You have selected a feature that is not yet available");
  }
}
