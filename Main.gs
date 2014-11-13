/*
 import MbfrA4t4piz3urmVxsyjgRCayG_O8PAw4 library as GSUtils
*/


function sortByTotal() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet != null) {
    
    var NUMFROZENCOLUMNS = sheet.getFrozenColumns()
    var NUMFROZENROWS = sheet.getFrozenRows()
    var FIRSTROW = NUMFROZENROWS + 1
    var LASTROW = sheet.getLastRow() 
    var LASTCOL = sheet.getLastColumn()
    
    //Sort by column C descending
    sheet.getRange(FIRSTROW, 1, LASTROW - FIRSTROW, LASTCOL).sort({column: LINEAL_HEADER[1], ascending: false});
    
  }
}



function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Create new MDMD sheet",
    functionName : "createMDMD"
  }, {
    name : "Update current MDMD",
    functionName : "updateMDMD"
  }, {
    name : "Sort current sheet by total",
    functionName : "sortByTotal"
  }];
  sheet.addMenu("Scripts", entries);
};
