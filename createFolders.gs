


function createClassFolders(){ //Create student folders
  //this step looks to see if the currently logged in user is looking to 
  //transfer ownership of folders to another teacher
  //sortsheet();
  
  //if initialSettings False open up Create Labels
  //PropertiesService.getDocumentProperties().setProperty('initialSettings', true);
  
  removeResumeTrigger();
  var lock = LockService.getPublicLock();
  lock.releaseLock();
  lock = LockService.getPublicLock();
  if (lock.tryLock(500)) {
    var startTime = new Date().getTime();
    var properties = PropertiesService.getDocumentProperties().getProperties();
    var dropBoxLabel = this.labels().dropBox;
    var dropBoxLabels = this.labels().dropBoxes;
    var periodLabel = this.labels().period;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssOwner = ss.getOwner().getEmail().toLowerCase();
    var currUser = Session.getEffectiveUser().getEmail().toLowerCase();
    var sheet = getRosterSheet(); 
    var dataRange = sheet.getDataRange().getValues();
    var labelObject = this.labels();
    var indices = returnIndices(dataRange, labelObject);
    var driveRoot = call(function() {return DriveApp.getRootFolder();});
    
    
    
    //This function adds robustness to the script by ensuring that we are looking in the correct
    //array indices for each of the elements.  If essential headers are missing, user is prompted to allow the script to auto-repair them.
    var indices = returnIndices(dataRange, labelObject);
    saveIndices(indices);
    writeProperties();
    SpreadsheetApp.setActiveSheet(sheet);
    //Sort by class, period, and last name to help consolidate rosters.  
    //note that this step is no longer technically necessary to ensure folder uniqueness.
    //left this in to provide ease of completeness check on class rosters
  //  sortsheet(indices.clsNameIndex,indices.clsPerIndex, indices.sLnameIndex);
    
    
    //now that all headings have been checked and indices identified
    //reload sheet and get 2D array of sheet data in case anything has changed.
    var sheet = getRosterSheet();
    var dataRange = sheet.getDataRange();
    dataRange = dataRange.getValues();
    
    //Remove wrap from non-header rows to economize on space.
    var wrapRange = sheet.getRange(1, 2, sheet.getLastRow(), sheet.getLastColumn()).setWrap(false);
    
    //Initialize counters
    var studentFoldersCreated = 0;
    
    var userEmail = Session.getEffectiveUser().getEmail(); //used later to check if script running user is the teacher whose email is listed, for ownership purposes 
    var clsFoldersCreated = []; //array to store all new folders created
    var editors = ss.getEditors();
    var editorEmails = [];
    for (var j=0; j<editors.length; j++) {
      editorEmails.push(editors[j].getEmail());
    }
    
    var interrupted = false;
    for (var i = 1; i < dataRange.length; i++) { //commence loop through all student class/period entries
      var loopStart = new Date().getTime();
      if ((loopStart - startTime)>310000) {
        setResumeTrigger(lock);
        interrupted = true;
        break;
      }
      var statusTagStudent = ""; //string used to concatenate student status messages
      var sFname = dataRange[i][indices.sFnameIndex]; // note that all sheet values are now 
      var sLname = dataRange[i][indices.sLnameIndex]; // addressed by variable index.  This is a safer
      var sEmail = dataRange[i][indices.sEmailIndex];  // way to roll than fixed indices, 
      var clsName = dataRange[i][indices.clsNameIndex]; // given how easy it is to drag a column in Google Spreadsheets
      var clsPer = dataRange[i][indices.clsPerIndex];    
      var tEmails = returnEmailAsArray(dataRange[i][indices.tEmailIndex]); //converts value of email column to an array of emails
      if (tEmails[0]=='') {
        tEmails[0]=userEmail;
        sheet.getRange(i+1, indices.tEmailIndex+1).setValue(userEmail); //if email is blank, assume the person running the script is the teacher
      }
      
      var sDropStatus = dataRange[i][indices.sDropStatusIndex];
      var tShareStatus = dataRange[i][indices.tShareStatusIndex];
      var rootStuFolderId = dataRange[i][indices.rsfIdIndex];
      var dropboxRootId = dataRange[i][indices.rsfIdIndex];
      var dropboxLabelId;
      if (i>0) {
        if ((dataRange[i][indices.clsNameIndex]==dataRange[i-1][indices.clsNameIndex])&&(dataRange[i][indices.clsPerIndex]!=dataRange[i-1][indices.clsPerIndex])&&(dataRange[i][indices.dbfIdIndex]=="")) {
          var dropbox = call(function() { return DriveApp.getFolderById(dataRange[i-1][indices.rsfIdIndex]);}); //if student dropbox already exists in sheet 
          var dropboxParents = dropbox.getParents();
          while (dropboxParents.hasNext) {
            var dropBoxParent = dropboxParents.next();
            break;
           //find its parent folder id in case a new period folder is needed
          }
          dropboxLabelId = dropBoxParent.getId(); 
        }
      }
      var clsFolderId = dataRange[i][indices.crfIdIndex];
      var classViewId = dataRange[i][indices.cvfIdIndex];
      var classEditId = dataRange[i][indices.cefIdIndex];
      var teacherFId = dataRange[i][indices.tfIdIndex];
      
      if ((sDropStatus=="")&&(clsName!='')) { //only create folders in rows where students have blank status and a class assigned.
        var uniqueClasses = getUniqueClassNames(dataRange, indices.clsNameIndex, indices.crfIdIndex); //returns array of all classes that already have class root folders listed in the sheet
        if (uniqueClasses.indexOf(clsName)==-1) { //only create new class folder if this class folder doesn't already exist
          try {
            var clsFolder = DriveApp.createFolder(clsName);
          } catch(err)  {  
            if (err.message.search("too many times")>0) {
              Browser.msgBox("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard");
              return;
            }
          }
          try {
            gClassFolders_logTeacherClassFolderCreated();
          } catch(err) {
          }
          var clsFolderId = clsFolder.getId();
          dataRange[i][indices.crfIdIndex] = clsFolderId;
          clsFoldersCreated.push(clsName);
          var classEdit = clsName +" - Edit";  
          var classView = clsName +" - View"; 
          var teacherFolderLabel = clsName + " - Teacher";
          var tMessage = "Folders created for " + clsName;
          //treat the first listed teacher email as primary...allow secondary teachers to be added
          for (var j=0; j<tEmails.length; j++) {//Transfer ownership of rootFolder to teacher if teacher email is designated.  Check that designated email is not the user running the script.
            try {
              if ((tEmails[j] != "")&&(tEmails[j] != userEmail)){
                DriveApp.getFolderById(clsFolderId).addEditor(tEmails[j]); 
                tMessage += ", " + tEmails[j] + " " + t("added as editor.");
              } else { //do this if teacher email is the same as that of the script user, or if tEmail is blank. This can only happen in teacher mode.
                tMessage += ", you're the teacher.";
                sheet.getRange(i+1, indices.tShareStatusIndex+1).setValue(tMessage);  
              }
            } catch(err) {
              DriveApp.getFolderById(clsFolderId).addEditor(tEmails[j]);
              tMessage += ", Error sharing folder for: " + tEmails[j] + "Error: " + err;
            }
            sheet.getRange(i+1, indices.tShareStatusIndex+1).setValue(tMessage);
            if ((tEmails[j]!=userEmail)&&(editors.indexOf(tEmails[j])==-1)) {
              ss.addViewer(tEmails[j]);
            }
          }
          //Create class edit, class view, and dropbox sub-folders
          try {
            var classEditId = call(function(){ return DriveApp.getFolderById(clsFolderId).createFolder(classEdit).getId();});
            var classViewId = call(function(){ return DriveApp.getFolderById(clsFolderId).createFolder(classView).getId();});
            var teacherFId =  call(function() { return DriveApp.getFolderById(clsFolderId).createFolder(teacherFolderLabel).getId();});
          } catch(err) {  
            if (err.message.search("too many times")>0) {
              Browser.msgBox("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard");
              return;
            }
          }
          dataRange[i][indices.tfIdIndex] = teacherFId;
          rootStuFolderId = call(function() { return DriveApp.getFolderById(clsFolderId).createFolder(dropBoxLabels).getId();}); //assign rootStuFolderId for now, pending a check whether period exists
          for (var j=0; j<tEmails.length; j++) {
            if ((tEmails[j]!="")&&(tEmails[j] != userEmail)) {//execute only if teacher email field is neither blank nor the same as the user running the script
              try {
                call(function(){DriveApp.getFolderById(classEditId).addEditor(tEmails[j]);});
                call(function(){DriveApp.getFolderById(classViewId).addEditor(tEmails[j]);});
                call(function(){DriveApp.getFolderById(rootStuFolderId).addEditor(tEmails[j]);});
                call(function(){DriveApp.getFolderById(teacherFId).addEditor(tEmails[j]);});
              } catch (err) {
                tMessage += ", Error sharing folder for: " + tEmails[j] + "Error: " + err;
              }        
            }
          }
          var dropboxLabelId = rootStuFolderId;
          //move to next username in class
        } // End of create class Folders
        var classRoster = null;
        var perRoster = null;
        if(rootStuFolderId=="") {
          perRoster = getClassRoster(dataRange, indices, clsName, clsPer);
          rootStuFolderId =  getClassFolderId(perRoster, indices.rsfIdIndex);
        }  
        if ((!dropboxLabelId)||(dropboxLabelId=="")) {
          classRoster = getClassRoster(dataRange, indices, clsName);
          if (!rootStuFolderId) {  
            rootStuFolderId =  getClassFolderId(classRoster, indices.rsfIdIndex);
          }      
          var dropboxRoot = call(function() { return DriveApp.getFolderById(rootStuFolderId);});
          dropboxLabelId = dropboxRoot.getId();
        }
        if (clsFolderId=="") {
          if (!classRoster) {
            classRoster = getClassRoster(dataRange, indices, clsName);
          }
          clsFolderId =  getClassFolderId(classRoster, indices.crfIdIndex);
        }
        if (classViewId=="") {
          classRoster = getClassRoster(dataRange, indices, clsName);
          classViewId = getClassFolderId(classRoster, indices.cvfIdIndex);
        }
        if (classEditId=="") {
          classRoster = getClassRoster(dataRange, indices, clsName);
          classEditId = getClassFolderId(classRoster, indices.cefIdIndex);
        }
        if (teacherFId=="") {
          classRoster = getClassRoster(dataRange, indices, clsName);
          teacherFId = getClassFolderId(classRoster, indices.tfIdIndex);
        }
        //If a class period is chosen, look to see if it is new or already existing
        if (clsPer != "") {
          var uniqueClasses = getUniqueClassPeriods(dataRange, indices.clsNameIndex, indices.clsPerIndex, indices.rsfIdIndex, labelObject); //get unique ClassPer as array
          if (uniqueClasses.indexOf(clsName + " " + periodLabel + " " + clsPer)==-1) { //look to see if this row's ClassPer exists in the array.  If not make a new student dropbox folder for the period
            rootStuFolderId = call(function() { return DriveApp.getFolderById(dropboxLabelId).createFolder(clsName + " " + periodLabel + " " + clsPer + " " + dropBoxLabels).getId();});
            clsFoldersCreated.push(clsName + " " + periodLabel + " " + clsPer);
            for (var j=0; j<tEmails.length; j++) {
              if ((tEmails[j] != "")&&(tEmails[j] != userEmail)) {
                call(function() {DriveApp.getFolderById(rootStuFolderId).addEditor(tEmails[j]);});
              }
            }
          }
        } // End if Per
        
        //Create students
        var dbfId = dataRange[i][indices.dbfIdIndex];
        var studentFolderObj = createDropbox(sLname,sFname,sEmail,clsName,classEditId,classViewId,rootStuFolderId,tEmails, userEmail, properties, dropBoxLabel);
        studentFoldersCreated++;
        var values = [];
        values[0] = [];
        dataRange[i][indices.dbfIdIndex] = studentFolderObj.studentDropboxId;
        dataRange[i][indices.crfIdIndex] = clsFolderId;
        dataRange[i][indices.cvfIdIndex] = classViewId; 
        dataRange[i][indices.cefIdIndex] = classEditId;
        dataRange[i][indices.rsfIdIndex] = rootStuFolderId;
        dataRange[i][indices.tfIdIndex] = teacherFId;
        
        values[0].push('=hyperlink("'+ studentFolderObj.studentDropbox.getUrl() +'";"'+studentFolderObj.studentDropboxId + '")');
        values[0].push('=hyperlink("'+ DriveApp.getFolderById(clsFolderId).getUrl() + '";"' + clsFolderId + '")');
        values[0].push('=hyperlink("' + studentFolderObj.classView.getUrl() + '";"' + classViewId + '")');
        values[0].push('=hyperlink("' + studentFolderObj.classEdit.getUrl() + '";"' + classEditId + '")');
        values[0].push('=hyperlink("' + studentFolderObj.rootStudentFolder.getUrl() + '";"' + rootStuFolderId + '")');
        values[0].push('=hyperlink("' + DriveApp.getFolderById(teacherFId).getUrl() + '";"' + teacherFId + '")');
        sheet.getRange(i+1, indices.dbfIdIndex + 1, 1, 6).setFormulas(values).setFontColor('black');
 
        //add Status 
        sheet.getRange(i+1, indices.sDropStatusIndex+1).setValue(studentFolderObj.statusTagF).setFontColor('black');
        SpreadsheetApp.flush();
      }
    }//end loop through all student class/period entries
    
    var msg = '';
    if (interrupted) {
      msg += 'The folder creation process was interrupted and will restart automatically to avoid script timeout. Please allow the script at least 1 minute to resume before attempting to resume manually. ';
    }
    if (clsFoldersCreated.length>0) {
      msg = "Class folders were created for:\n" + clsFoldersCreated.join(", \n") + "\n \n";
    } 
    if (studentFoldersCreated>0) {
      PropertiesService.getDocumentProperties().setProperty('alreadyRan', 'true');
      
      //onOpen();
      msg += " " + studentFoldersCreated + " new " + dropBoxLabels + " were created.\n \n";
    } else {
      msg += "No new folders were created.  Folders are only created for rows with a blank \"Status: Student Dropbox \" value\n \n";
    }
    lock.releaseLock();
    Browser.msgBox(msg);
  } else {
    Browser.msgBox("It appears the folder creation process is already underway. Please don't interrupt!");
  }
}






function createDropbox(sLnameF,sFnameF,sEmailF,clsNameF,classEditIdF,classViewIdF,rootStuFolderId,tEmails,userEmail, properties, dropboxLabel) {
  var returnObject = new Object();
  var dropboxNameF = sLnameF + ", " + sFnameF + " - " + clsNameF + " - " + dropboxLabel;
  var rootStudentFolder = call(function() {return DriveApp.getFolderById(rootStuFolderId);});
  var studentDropbox = call(function() {return rootStudentFolder.createFolder(dropboxNameF);});
  returnObject.statusTagF = dropboxLabel + " created";
  try {
    var classEdit = call(function() {return DriveApp.getFolderById(classEditIdF);});
    var classView = call(function() {return DriveApp.getFolderById(classViewIdF);});
    call(function() {classEdit.addEditor(sEmailF);});
    call(function() {classView.addViewer(sEmailF);});
    call(function() {studentDropbox.addEditor(sEmailF);});
    returnObject.statusTagF += ", and shared with " + sEmailF;
  } catch(e) {
    Logger.log("Error with email (" + sEmailF + "). " + e.message);
    returnObject.statusTagF += ", Error with Student email: folder created but not shared"; 
  }
  var studentDropboxId = studentDropbox.getId()
  returnObject.studentDropboxId = studentDropboxId;
  returnObject.studentDropbox = studentDropbox;
  returnObject.classEdit = classEdit;
  returnObject.classView = classView;
  returnObject.rootStudentFolder = rootStudentFolder;
  for (var j=0; j<tEmails.length;j++) {
    if ((tEmails[j] != "")&&(tEmails[j] != userEmail)) {  
      try {
        call(function(){ DriveApp.getFolderById(studentDropboxId).addEditor(tEmails[j]); });
        returnObject.statusTagF += ", editing rights added for " + tEmails[j];
      } catch(err) {
        returnObject.statusTagF += ", error giving editing rights to " + tEmails[j] + "." + err;
      }
    }
  } 
  return returnObject;
}
