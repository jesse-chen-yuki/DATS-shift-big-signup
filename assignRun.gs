/*
assign run number based on operator sign up slip

@param {number} badge number of the operator
@return {number} run number assigned to operator
*/

function AssignRun(badge) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  SpreadsheetApp.getActiveSpreadsheet();

  console.log('hello');
  //badge = 6004
  var taken = findTaken(badge);
  console.log('taken list ' + taken);
  var selectionList = findSelection(badge);
  console.log('selection list '+selectionList);
  var shiftChoice = makeSelection(badge, taken, selectionList)
  console.log('final choice '+shiftChoice);

  return shiftChoice;
}

/* select the first shift available and add to main list
@PARAM {number} badge number of the operator
@PARAM {list} list of shift already taken
@PARAM {list} list of operator choice slip
@custum function
*/

function makeSelection(badge,takenList,selectionList){
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  //ss = SpreadsheetApp.getActiveSpreadsheet();
  // console.log('selection list ' + selectionList);
  for (i = 0; i<selectionList.length;i++){
    //console.log('selection '+selectionList[i]);
    if (takenList.indexOf(selectionList[i])==-1){
      return selectionList[i]
    }
  }
}

/* find out the shift choices operaters chose

@PARAM {number} badge number of the operator
@RETURN {list} of operator selection
@costume function

*/

function findSelection(badge){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[2]);
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var selected = []

  for (var i = 1; i < values.length; i++) {
    //console.log(values[i]);

    if (values[i][2] === badge){
      for (j = 3;j<values[i].length;j++){
        selected.push(values[i][j]);
      }
      
      SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
      SpreadsheetApp.getActiveSpreadsheet();
      return selected;
    }
  }  
}

/* find out the shift choices already taken

@PARAM {number} badge number of the current operator
@RETURN {list} of shift choices already taken
@costume function

*/

function findTaken(badge) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[2]);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var taken = []

  for (var i = 3; i < values.length; i++) {
    //console.log(values[i]);
    //var row = "";
    
    if (values[i][1] == badge){
        return taken
      }  else {
        taken.push(values[i][4])
      }
    
  }  
}
