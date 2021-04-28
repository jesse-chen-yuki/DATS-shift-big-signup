/**
This program iterates through all operator and calls AssignRun function to determine a shift to be assigned to an operator for a big signup.

Version 1.3
Written by Jesse Xi Chen

@ custom function

Pre-conditions:
1st sheet of the file contain main information about the signing times for each operator
  Program assumes the list is sorted in ascending seniority order
  2nd column list the badge numbers of the operator
  4th column may contain 'not signing'
Digital sign up choices form is imported as the 3rd sheet. 

Post-conditions:
Choices assigned to operator is updated in the 5th column of the main sheet.

Update:
Ver 1.3 2021-04-28
-Allow continuation of sign up process in the middle of the list. 

TODO:
For Ver 1.4+
-Optimize functions: find selection
-Event trigger upon new response from the google form doucment (shift choice form).
  -no reset, find the last row with valid entry,
  -find list of operator updates we can make this time.
  -wait on future input.
  -may get sepecial flag as final choice
-Create valid input list and compare input against it.
-Testing for robustness, bad input.

*/

DEBUG = 0;

/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 * Add execute main program to the drop down menu
 * 
 * IMPORTANT: Can Menu be seen by operators? operators should not be able to trigger menu items.
 * 
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Fill Choices')
    .addItem('Mode 1, All slip Received', 'assignMode1')
    .addItem('Mode 2，Continue mode Not implemented', 'assignMode2')
    .addToUi();
}

/**  assignMode1 is used to fill in the column of regular shift choice assuming all the input from operator has been received.
 * called by the drop down menu
 * assumes all choice slips are submitted.
 **/

function assignMode1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  ss.setActiveSheet(sheet);
  //var sheet = SpreadsheetApp.getActiveSpreadsheet();

  //console.log('hello');
  reset(sheet);
  
  var badgeColumn = sheet.getRange("B:D").getValues();
  if (DEBUG){
    //console.log(badgeColumn)
  }
  var currentRow;
  var badge, finalChoice;
  var reliefNumber = 1;
  var takenList = [];

  // go through the column of all badge numbers
  for (var rowIndex=3; rowIndex<badgeColumn.length; rowIndex++){
    currentRow = rowIndex+1;
    badge = badgeColumn[rowIndex][0];

    //check for valid badge number skip rows that are empty or have non-badge info, or not signing
    if (badge<6001 || badge >6900 || badgeColumn[rowIndex][2].toString().trim()==='Not Signing'){
      continue;
    }

    if (DEBUG){
      console.log('processing: '+badge);
      console.log('takenList: '+takenList);
    }
    finalChoice = assignRun(badge,takenList);
    if (DEBUG){
      console.log('finalChoice: '+finalChoice);
    }
  
    if (finalChoice.toString().toLowerCase()[0] == 'r'){
      finalChoice = 'Relief ';
      finalChoice+=reliefNumber++;
    }
    
    takenList.push(finalChoice);   
    sheet.getRange("E"+currentRow).setValue(finalChoice);
  }
}

/** assignMode2 is the second mode of updating the choice column
 * it will check what the current progress is, finding the last operator with a choice decision
 * check out how many more operator can be processed before waiting on more information
 * make the update required
 */
function assignMode2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheets()[0];
  ss.setActiveSheet(sheet1);

  var badgeColumns = sheet1.getRange("B:E").getValues();

  // find current row that have existing info
  var currentRow = findNextToUpdate(badgeColumns);
  console.log('row to update: '+ currentRow);

  // find taken list
  var takenList = findTaken(sheet1);

  // find index for relief
  var reliefNumber = findReliefIndex(takenList)+1;

  var badge, finalChoice;
  

  // go through the column of all badge numbers
  for (var rowIndex=currentRow; rowIndex<badgeColumns.length; rowIndex++){
    currentRow = rowIndex+1;
    badge = badgeColumns[rowIndex][0];

    //check for valid badge number skip rows that are empty or have non-badge info, or not signing
    if (badge<6001 || badge >6900 || badgeColumns[rowIndex][2].toString().trim()==='Not Signing'){
      continue;
    }

    if (DEBUG){
      console.log('processing: '+badge);
      console.log('takenList: '+takenList);
    }
    finalChoice = assignRun(badge,takenList);
    if (DEBUG){
      console.log('finalChoice: '+finalChoice);
    }
  
    if (finalChoice.toString().toLowerCase()[0] == 'r'){
      finalChoice = 'Relief ';
      finalChoice+=reliefNumber++;
    }
    
    takenList.push(finalChoice);
    sheet1.getRange("E"+currentRow).setValue(finalChoice);
  }

}

/** findNextToUpdate takes in a range list that contain info about operator badge and signup result up till now.
 * it will return the row number within the range that need to be updated next.
 */

function findNextToUpdate(aRange){
  for (var i=3;i<aRange.length;i++){
    console.log(aRange[i]);
    if (aRange[i][0]>6000&&aRange[i][0]<6900&&aRange[i][3]==''){
      if (aRange[i][2].toString().trim()==='Not Signing'){
        continue;
      } else {return i;
      }
    }
  }
}

/** findTaken returns a list of shift already taken in the choice column 
find out the shift choices already taken

@PARAM {sheet} current sheet containing operator info
@RETURN {list} of shift choices already taken
@costume function

*/
function findTaken(sheet) {
  var taken=[];
  var takenColumns = sheet.getRange("E4:E").getValues();
  for (var i = 0;i<takenColumns.length;i++){
    taken.push(takenColumns[i][0]);
  }

  return taken;
}

/** findReliefIndex checks existing choices for relief position and pass on the index for the next relief */
function findReliefIndex(aList){
  //console.log('aList '+ aList)
  
  index = 0;
  for (var i =0; i<aList.length;i++){
    //console.log(aList[i]);
    if (aList[i].toString().match(/Relief/g)){
            index++;
    }
  }
  if (DEBUG) console.log('last index ' + index);
  return index;
 
}


/** reset function clears all content on the choice column */
function reset(sheet){
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rangeList = sheet.getRangeList(['E4:E160']);
  rangeList.clear({contentsOnly: true});
  //Utilities.sleep(5000);
}

/**
assign run number based on operator sign up slip

@PARAM {number} badge number of the operator
@return {number} run number assigned to operator

@custom function
*/

function assignRun(badge,taken) {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  //SpreadsheetApp.getActiveSpreadsheet();

  //console.info('In assignRun');
  //badge = 6004
  //var taken = findTaken(badge);
  //console.log('taken list ' + taken);
  var selectionList = findSelection(badge);
  //console.log('selection list '+selectionList);
  var shiftChoice = makeSelection(taken, selectionList)
  //console.log('final choice '+shiftChoice);
  if (DEBUG){
    console.log('taken list ' + taken);
    console.log('selection list '+selectionList);
    console.log('final choice '+shiftChoice);
  }
  return shiftChoice;
  
  /*return new Promise((resolve,reject)=>{
    setTimeout(() => {
    resolve(shiftChoice);
    }, 1000)}); */
}



/** select the first shift available and add to main list
@PARAM {number} badge number of the operator
@PARAM {list} list of shift already taken
@PARAM {list} list of operator choice slip
@custum function
*/

function makeSelection(takenList,selectionList){
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  //ss = SpreadsheetApp.getActiveSpreadsheet();
  // console.log('selection list ' + selectionList);
  for (var i = 0; i<selectionList.length;i++){
    //console.log('selection '+selectionList[i]);
    if (takenList.indexOf(selectionList[i])==-1){
      return selectionList[i]
    }
  }
}

/** find out the shift choices operaters chose
@PARAM {number} badge number of the operator
@RETURN {list} of operator selection
@costume function

*/

function findSelection(badge){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2];
  //SpreadsheetApp.setActiveSheet(ss.getSheets()[2]);
  //var sheet = ss.getSheetByName("sheetName");
  //var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var selected = []

  //var badgeRange = sheet.

  for (var i = 1; i < values.length; i++) {
    /*
    console.log(values[i]);
    console.log(values[i][2]);
    console.log(badge); 
    */

    if (values[i][2] === badge){
      for (j = 3;j<values[i].length;j++){
        selected.push(values[i][j]);
      }
      
      //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
      //SpreadsheetApp.getActiveSpreadsheet();
      return selected;
    }
  }  
}






