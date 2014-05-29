


//Builds the UI for initial settings
function renameFolderLabels() {
  //This will check to see if gCF is set up yet, if the person 
  if (setupGcfCheck() == false){ return;}

  
  var app = UiApp.createApplication().setHeight(380).setWidth(700);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var panel1 = app.createVerticalPanel(); //title
  var panel2 = app.createHorizontalPanel(); //grid and help image
  var title = app.createLabel("Choose your folder labels").setId('title').setStyleAttribute('fontSize', '18px').setStyleAttribute('marginBottom', '8px');
  var documentProperties = PropertiesService.getDocumentProperties();
 //documentProperties.setProperty('runCreate', 'true');
  var properties = documentProperties.getProperties();
  var saveButtonText = "Save";
  
  
//  if (properties.runCreate == "true"){
//    saveButtonText = "Save and create folders";
//  }
  
  
  // Example Image
  var helpImageTitle = app.createLabel("Example").setStyleAttribute('font-weight','bold')
  var helpImage = app.createImage(getImage("0Bwyqwd2fAHMMWms4NDVDSWZ2Vlk"));
  var helpGrid = app.createGrid(2,1).setId('helpGrid').setCellPadding(3);
  helpGrid
  .setWidget(0,0, helpImageTitle)
  .setWidget(1,0, helpImage);
  

  
 
//Get labes from Cache  
  var dropBoxLabels = this.labels().dropBoxes;  
  var dropBoxLabel = this.labels().dropBox;
  var periodLabel = this.labels().period;
  var editLabel = this.labels().edit;
  var viewLabel = this.labels().view;
  var teacherLabel = this.labels().teacher;
    
    
  
  if (properties.alreadyRan == "true"){
    
   var namingLabel = app.createLabel("These settings have been locked becuase folders have already been created.").setId('namingLabel').setStyleAttributes({'marginTop': '5px', 'marginBottom': '20px'});
   var namingGrid = app.createGrid(7, 2).setId('namingGrid').setCellPadding(3);
   namingGrid
      .setWidget(0, 0, app.createLabel("Default Label").setStyleAttribute('font-weight','bold'))
   .setWidget(0, 1, app.createLabel("Current Label").setStyleAttribute('font-weight','bold'))
   .setWidget(1, 0, app.createLabel(defaultLabels.dropBoxes))
   .setWidget(1, 1, app.createTextBox().setName('dropBoxes').setValue(dropBoxLabels).setEnabled(false)) 
   .setWidget(2, 0, app.createLabel(defaultLabels.dropBox))
   .setWidget(2, 1, app.createTextBox().setName('dropBox').setValue(dropBoxLabel).setEnabled(false))
   .setWidget(3, 0, app.createLabel(defaultLabels.period))
   .setWidget(3, 1, app.createTextBox().setName('period').setValue(periodLabel).setEnabled(false))
   .setWidget(4, 0, app.createLabel(defaultLabels.edit))
   .setWidget(4, 1, app.createTextBox().setName('edit').setValue(editLabel).setEnabled(false))
   .setWidget(5, 0, app.createLabel(defaultLabels.view))
   .setWidget(5, 1, app.createTextBox().setName('view').setValue(viewLabel).setEnabled(false))
   .setWidget(6, 0, app.createLabel(defaultLabels.teacher))
   .setWidget(6, 1, app.createTextBox().setName('teacher').setValue(teacherLabel).setEnabled(false));
 
   

  var closeButton = app.createButton('close');
  var closeHandler = app.createServerClickHandler('closePanel');
  var endButton = closeButton.addClickHandler(closeHandler);
    
                                  
  } else {
  var saveHandler = app.createServerHandler('saveLabelSettings').addCallbackElement(panel);
  var namingLabel = app.createLabel("The terms below can be renamed. These labels determine how important folders and columns will be named by gClassFolders. Once you create the folders you will not be able to change these settings.").setId('namingLabel').setStyleAttributes({'marginTop': '5px', 'marginBottom': '20px'});
  var namingGrid = app.createGrid(7, 2).setId('namingGrid').setCellPadding(3);
  var closeHandler = app.createServerClickHandler('closePanel');
 
    
  namingGrid
   //.setStyleAttribute('top', '100px')
   .setWidget(0, 0, app.createLabel("Default Label").setStyleAttribute('font-weight','bold'))
   .setWidget(0, 1, app.createLabel("Current Label").setStyleAttribute('font-weight','bold'))
   .setWidget(1, 0, app.createLabel(defaultLabels.dropBoxes))
   .setWidget(1, 1, app.createTextBox().setName('dropBoxes').setValue(dropBoxLabels)) 
   .setWidget(2, 0, app.createLabel(defaultLabels.dropBox))
   .setWidget(2, 1, app.createTextBox().setName('dropBox').setValue(dropBoxLabel))
   .setWidget(3, 0, app.createLabel(defaultLabels.period))
   .setWidget(3, 1, app.createTextBox().setName('period').setValue(periodLabel))
   .setWidget(4, 0, app.createLabel(defaultLabels.edit))
   .setWidget(4, 1, app.createTextBox().setName('edit').setValue(editLabel))
   .setWidget(5, 0, app.createLabel(defaultLabels.view))
   .setWidget(5, 1, app.createTextBox().setName('view').setValue(viewLabel))
   .setWidget(6, 0, app.createLabel(defaultLabels.teacher))
   .setWidget(6, 1, app.createTextBox().setName('teacher').setValue(teacherLabel));

  
 // var endButton = app.createButton(saveButtonText, saveHandler).setId('button');
  var endButton = app.createButton(saveButtonText).setId('button');
  endButton.addClickHandler(saveHandler);
  endButton.addClickHandler(closeHandler);  
     
  }

  
  //build the visual
  panel1.add(title);
  panel1.add(namingLabel);
  panel2.add(namingGrid);
  panel2.add(helpGrid);
  panel.add(panel1);
  panel.add(panel2);
  app.add(panel);
  app.add(endButton);  
  
  
  ss.show(app);  
  return app;
}  // End Set Labels



// Saves folder label settings
function saveLabelSettings(e) {
  
  var documentProperties = PropertiesService.getDocumentProperties();
  var properties = documentProperties.getProperties();
  var app = UiApp.getActiveApplication();
  
  var dropBoxes = e.parameter.dropBoxes;
  var dropBox= e.parameter.dropBox;
  var period = e.parameter.period;
  var edit = e.parameter.edit;
  var view = e.parameter.view;
  var teacher = e.parameter.teacher;
  var saved = "true";
  var runCreate = properties.runCreate;
  
  properties.labels = JSON.stringify({dropBoxes: dropBoxes, dropBox: dropBox, period: period, edit: edit, view: view, teacher: teacher, saved: saved});
  properties.ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  documentProperties.setProperties(properties);
  
  
  app.close();
//  if (runCreate == "true"){
//    documentProperties.setProperty('runCreate', 'false');
//    //startCF()
//    createClassFolders2();
//  }
    
  
  //var labels = properties.labels;
  // CacheService.getPrivateCache().put('labels', labels, 660);
  
  //create or update RosterSheet
  //var properties = PropertiesService.getDocumentProperties().getProperties();
  //var sheetId = properties.sheetId;
  //fixHeaders();
  
  return app;
  
}



