// gClassFolders
// Original resource by EdListen.com 
// Original Author: Bjorn Behrendt bj@edlisten.com
// Version 2.1.2-dev (9/9/2013) a collaboration with Andrew Stillman and YouPD.org, a project of New Visions for Public Schools.
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Version 3.0 (3/15/14) ported as a script to the New Google Sheets by Andrew Stillman.  Please do not contact re: maintenance issues.
// Version 3.1 maintained by Bjorn Behrendt.  Post all support issues to https://plus.google.com/communities/115718335045383669895


//To do: All strings in the UI can be abstracted as properties of language objects
//This is currently the best practice in internationalizing Apps Script
//Don't have time to do this, but this should get you started...
//Just replace any string in the UI with LANG.stringName  where stringName is the property name you've assigned that particular string.

var LANG_EN = {
   appTitle: "gClassFolders",
   saveLowerCase: "save",
 //etc.

}


var LANG_SV = {
  appTitle: "gClassFolderSchteinhimmler",
  saveLowerCase: "swedish word for save",
 //etc. 
};


var LANG = LANG_EN;
var locale = Session.getActiveUserLocale();
if(locale == 'sv') LANG = LANG_SV;
