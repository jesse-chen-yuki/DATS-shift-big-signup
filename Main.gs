/* 
This program iterates through all operator and calls AssignRun function to determine a shift to be assigned to an operator for a big signup.

Version 1.1
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

TODO:
For Ver 1.2+
-Create valid input list and compare input against it.
-Testing for robustness, bad input.
-Optimize functions.

-Allow continuation of sign up process in the middle of the list. Event trigger upon new response from the google form doucment (shift choice form).

*/

function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  //console.log('hello');
  reset();
  
  var badgeColumn = sheet.getRange("B:D").getValues();
  //console.log(badgeColumn)
  var currentRow;
  var badge, finalChoice;
  var reliefNumber = 1;

  // go through the column of all badge numbers
  for (var rowIndex=3; rowIndex<badgeColumn.length; rowIndex++){
    currentRow = rowIndex+1;
    badge = badgeColumn[rowIndex][0];

    //check for valid badge number skip rows that are empty or have non-badge info, or not signing
    if (badge<6001 || badge >6900 || badgeColumn[rowIndex][2].toString().trim()==='Not Signing'){
      continue;
    }

    console.log('processing: '+badge);
    finalChoice = assignRun(badge);
    if (finalChoice.toString().toLowerCase() == 'relief'){
      finalChoice+=reliefNumber++;
    }
    
    //ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
    sheet = SpreadsheetApp.getActiveSpreadsheet();
   
    sheet.getRange("E"+currentRow).setValue(finalChoice);
      
    /*
    sheet.getRange("E"+currentRow).setValue('=AssignRun(B'+currentRow+')');
    */
  }
}


function reset(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeList = sheet.getRangeList(['E4:E147']);
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

