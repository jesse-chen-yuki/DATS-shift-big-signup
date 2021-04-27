/* 
This program iterates through all operator and calls AssignRun function to determine a shift to be assigned to an operator for a big signup.

Version 1.2
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
Ver 1.2
-optimized taken list, delete uselsee Appsheet access.
-create custom menu function for easy access

TODO:
For Ver 1.3+
-Create valid input list and compare input against it.
-Testing for robustness, bad input.
-Optimize functions: find selection

-Allow continuation of sign up process in the middle of the list. Event trigger upon new response from the google form doucment (shift choice form).
-Use Triggering event from Google Form to automate the function call. 

*/

DEBUG = 1;

/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 * Add execute main program to the drop down menu
 * 
 * IMPORTANT: Can Menu be seen by operators, operators should not be able to trigger menu items.
 * 
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Fill Choices')
    .addItem('Mode 1, All slip Received', 'assignMode1')
    .addItem('Mode 2', 'assignMode2')
    .addSeparator()
    .addItem(
      'Separate title/author at first comma', 'splitAtFirstComma')
    .addItem(
      'Separate title/author at last "by"', 'splitAtLastBy')
    .addSeparator()
    .addItem(
      'Fill in blank titles and author cells', 'fillInTheBlanks')
    .addToUi();
  //assignMode1();
}

/**  assignMode1 is used to fill in the column of regular shift choice assuming all the input from operator has been received.
 * called by the drop down menu
 * 
 **/

function assignMode1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  ss.setActiveSheet(sheet);
  //var sheet = SpreadsheetApp.getActiveSpreadsheet();

  //console.log('hello');
  reset();
  
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
    //ss = SpreadsheetApp.getActiveSpreadsheet();
    //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
    //sheet = SpreadsheetApp.getActiveSpreadsheet();
   
    sheet.getRange("E"+currentRow).setValue(finalChoice);
  }
}


function reset(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rangeList = sheet.getRangeList(['E4:E160']);
  rangeList.clear({contentsOnly: true});
  //Utilities.sleep(5000);
}
/*
let promise = new Promise(function(resolve, reject) {
  // the function is executed automatically when the promise is constructed
  let run = 
  // after 1 second signal that the job is done with the result "done"
  setTimeout(() => {resolve("done"), 5000});
});*/

