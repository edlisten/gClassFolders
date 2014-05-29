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



function getClassRowNumsFromCRF(dataRange, indices, crfId) {
  var classRowNums = [];
    for (var i=1; i<dataRange.length; i++) {
      if (dataRange[i][indices.crfIdIndex]==crfId) {
        classRowNums.push(i+1);
      }
    }
  return classRowNums;
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
  for (var i=2; i<dataRange.length; i++) {
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
