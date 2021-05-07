/**
This program iterates through all operator and calls AssignRun function to determine a shift to be assigned to an operator for a big signup.

Version 1.4
Written by Jesse Xi Chen

Pre-conditions:
1st sheet of the file contain main information about the signing times for each operator
  Program assumes the list is sorted in ascending seniority order
  2nd column list the badge numbers of the operator
  4th column may contain 'not signing'
Digital sign up choices form are all sumbitted and is imported as the sheet 'Slips3'. 

Post-conditions:
Choices assigned to operator is updated in the 5th column of the main sheet upon form submission.

Update:
Ver 1.3 2021-04-28
-Allow continuation of sign up process in the middle of the list. 

Ver 1.3.1 2021-05-05
-removeForm function implemented
-Create form within the excel and add onto as the last choice
  -setup form creation within the menu option

Ver 1.3.2 2021-05-06
-Add function to find choice slip submitted, check for the first non-submit badge number
-Hult execution when choice slip can not be found
  -create toast message with the last updated badge number 
-Create google form for choice slip, add response as 'Slips3'

Ver 1.4 2021-05-07
-Event trigger upon new submit response from the google form doucment (shift choice form, sheet 'Slips3').
  -run assignMode2 on submission
  -wait on future input.
-Add description and instruction section to the Form.
-Add setting of verified login
-update createChoiceForm, resetForm

TODO:
For Ver 1.4+
-trigger on Edit of form response, call assignMode2 too. (design choice)
-setup data input validation? range for fulltime, parttime
-Create valid input list and compare input against it.
-Add section of operator badge next in line to submit, email notification (require email addr).
-Testing for robustness, bad input.

*/

DEBUG = 1;

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
  // ui may choose to hide items once update is automated with user submit
  ui.createMenu('Signup Tasks')
    .addItem('Create Choice Form', 'createChoiceForm' )
    .addItem('Reset Form', 'resetForm')
    .addSeparator()
    .addItem('Reset all choices', 'resetChoices')
    .addItem('Mode 1, All slips Received, Outdated, do not use ', 'assignMode1')
    .addItem('Mode 2，Continue mode, stop on no slip', 'assignMode2')
    .addToUi();
  assignMode2();
}

/** removes the form associated with the sign up
 *  assumes the name of the form response is Slip3
 * 
 */

function resetForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  if (DEBUG){
    showSheetList_();
  }
  
  var formSheet = ss.getSheetByName('Slips3')
  if (formSheet){
    var formUrl = formSheet.getFormUrl();
    if (formUrl)
      var form = FormApp.openByUrl(formUrl);
      var formId = form.getId();
      form.removeDestination();
      //form.setTrashed(true);
      try {
        var file = DriveApp.getFileById(formId);
      }
      catch (fileE) {
        try {
          file = DriveApp.getFolderById(fileId);
        }
        catch (folderE) {
          throw folderE;
        }
      }
      file.setTrashed(true);
    ss.deleteSheet(formSheet);
  }
  ss.setActiveSheet(ss.getSheets()[0]);
  createChoiceForm();
  resetChoices();

  //ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
  //    .create();
  ss.toast('form reset complete')
}


/**
 * Creates a Google Form that allows operators to fill in their choice slips
 * 
 * put the form response as the 5th sheet for now
 */

function createChoiceForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  if (DEBUG){
    showSheetList_();
  }
  ss.setActiveSheet(sheets[0]);

  // assume response sheet is the 5th sheet with name Slips3
  var responseSheet = ss.getSheetByName('Slips3');
  if (responseSheet){
    // response sheet exist
    ss.setActiveSheet(responseSheet);
    // check the existence of the form point to current ss
    var formUrl = SpreadsheetApp.getActive().getActiveSheet().getFormUrl();
    if (formUrl != null) {
      ss.toast('form response already exist.');
      return;
    }
  } else {
    // response sheet does not exist yet
    var sheetName = sheets[0].getName();
    // insert new sheet and set index to 0
    //var newSheet = ss.insertSheet('Slips3',0);
    var maxChoice = 20;
    var responseLocation = 5;
    
    // description field of the form
    var desc = 'This is the digital equivalent of the sign up choice form \n\n' +
      'Instruction:\n'+
      'Please enter your name, badge number and the run choices in order of preference.';

    // set up the choice slip form
    var form = FormApp.create('Big Signup Choice Slips '+ sheetName);
    form.setDescription(desc)
        .setAllowResponseEdits(true)
        .setCollectEmail(true)
        .setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())

    if (!DEBUG){
      form.setLimitOneResponsePerUser(true)
        .setRequireLogin(true)
    } 
    
    //form.setDestination(FormApp.DestinationType.SPREADSHEET, newSheet.getSheetId());

    if (DEBUG){
      showSheetList_();
    }
    var tempSheet = ss.insertSheet('temp');
    if (DEBUG){
      showSheetList_();
    }
    sheets = ss.getSheets();
    ss.setActiveSheet(sheets[0]).setName('Slips3');
    ss.moveActiveSheet(6);
    ss.deleteSheet(tempSheet);
    
    if (DEBUG){
      showSheetList_();
    }
    ss.setActiveSheet(sheets[0]);
    
    //var formSheet = SpreadsheetApp.getSheets()[0];
    //form.addTextItem().setTitle('Email').setRequired(true);
    form.addTextItem().setTitle('Name').setRequired(true);
    form.addTextItem().setTitle('Badge').setRequired(true);
    for (var i = 1;i<= maxChoice;i++){
      //form.addTextItem().setTitle('Choice '+i);
      if(i ==1){
        form.addTextItem().setTitle('Choice '+i).setRequired(true);
      } else {
        form.addTextItem().setTitle('Choice '+i);
      }
    }
    
    ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
  }
  ss.setActiveSheet(ss.getSheets()[0]);
  resetChoices();
  ss.toast('form creation complete')
}

/**
 * 
 * run assignMode2 after each user submission
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {

  assignMode2();

  
  if (DEBUG)
    console.log(e.namedValues)
  /*
  var response = [];
  var badge = e.namedValues['Badge'][0];
  var slipChoices = [];
  for (var i = 5; i<25;i++ ){
    choice = e.namedValues[badge][i]
    if (!choice)
      break;
    slipChoices.push(choice);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Slips3');
  sheet.insertRows(e.namedValues);

  /*
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};

  // Grab the session data again so that we can match it to the user's choices.
  var response = [];
  var values = SpreadsheetApp.getActive().getSheetByName('Conference Setup')
     .getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var day = session[1].toLocaleDateString();
    var time = session[2].toLocaleTimeString();
    var timeslot = time + ' ' + day;

    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[timeslot] && e.namedValues[timeslot] == title) {
      response.push(session);
    }
  }
  sendInvites_(user, response);
  sendDoc_(user, response);
  */
  
}

/**  assignMode1 is used to fill in the column of regular shift choice assuming all inputs from operator has been received.
 * called by the drop down menu
 * assumes all choice slips are submitted.
 **/

function assignMode1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  ss.setActiveSheet(sheet);
  //var sheet = SpreadsheetApp.getActiveSpreadsheet();

  //console.log('hello');
  resetChoices();
  
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
    // use assign mode 1
    finalChoice = assignRun_(badge,takenList,1);
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
  ss.toast('all run assignment complete')
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
  var currentRow = findNextToUpdate_(badgeColumns);
  if (DEBUG){
    console.log('row to update: '+ currentRow);
  }

  // find taken list
  var takenList = findTaken_(sheet1);

  // find index for relief
  var reliefNumber = findReliefIndex_(takenList)+1;

  var badge, finalChoice;
  
  // find how many updates this iteration can achieve
  var opList = getOperList_(badgeColumns);


  /** get the list of operators who need to enter a choice slip */
  function getOperList_(badgeList){
    var opList = [];
    var badge;
    for (var i=0; i<badgeList.length;i++){
      badge = badgeList[i][0]
      if (badge>=6001 && badge <6900 && badgeList[i][2].toString().trim()!=='Not Signing'){
        opList.push(badge);
      } else {
        continue;
      }
    }
    return opList;
  }
  
  var opSubmitList = getOperSubmitList_();

  // get the list of operator who have already submitted
  function getOperSubmitList_(){
    var slipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      'Slips3');
    var data = slipSheet.getDataRange().getValues();
    var col = data[0].indexOf('Badge');
    var badges = [];
    if (col == -1) {
      return -1;
    } else {
      //var badges = sheet.getDataRange(col+':'+col).getValues();
      for (var i=1;i<data.length;i++){
        badges.push(data[i][col]);
      }
      badges.sort();
      if (DEBUG){
        console.log('submitted badges list: '+ badges);
      }
      return badges;
    }
  }

  var firstNonSubmit = findFirstNonSubmit_(opList,opSubmitList);

  /** using the operator list and submitted list to find the last operator we can process 
   * 
  */
  function findFirstNonSubmit_(opList,opSubmitList){
    if (opSubmitList.length == 0){
      return opList[0];
    }
    
    for (var i=0;i<opList.length;i++){
      if (opSubmitList.indexOf(opList[i])>=0){
        continue;
      } else {
        return opList[i];
      }
    }
  }

  if (DEBUG){
    console.log('first non submit ' + firstNonSubmit);
  }


  // go through the column of all badge numbers
  for (var rowIndex=currentRow; rowIndex<badgeColumns.length; rowIndex++){
    currentRow = rowIndex+1;
    badge = badgeColumns[rowIndex][0];

    // current badge does not have corresponding choice slip
    if (badge ==firstNonSubmit){

      
      ss.toast(badge + ' does not have slip entered. Terminating');
      return;
    }

    //check for valid badge number skip rows that are empty or have non-badge info, or not signing
    if (badge<6001 || badge >6900 || badgeColumns[rowIndex][2].toString().trim()==='Not Signing'){
      continue;
    }

    if (DEBUG){
      console.log('processing: '+badge);
      console.log('takenList: '+takenList);
    }
    // use assign mode 2
    finalChoice = assignRun_(badge,takenList,2);
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
  ss.toast('all run assignment complete')
}




/** findNextToUpdate_ takes in a range list that contain info about operator badge and signup result up till now.
 * it will return the row number within the range that need to be updated next.
 */

function findNextToUpdate_(aRange){
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

/** findTaken_ returns a list of shift already taken in the choice column 
find out the shift choices already taken

@PARAM {sheet} current sheet containing operator info
@RETURN {list} of shift choices already taken
@costume function

*/
function findTaken_(sheet) {
  var taken=[];
  var takenColumns = sheet.getRange("E4:E").getValues();
  for (var i = 0;i<takenColumns.length;i++){
    taken.push(takenColumns[i][0]);
  }

  return taken;
}

/** findReliefIndex checks existing choices for relief position and pass on the index for the next relief 
 * 
*/
function findReliefIndex_(aList){
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


/** reset function clears all content on the choice column 
 * 
*/
function resetChoices(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var rangeList = sheet.getRangeList(['E4:E160']);
  rangeList.clear({contentsOnly: true});
  //Utilities.sleep(5000);
  ss.toast('all choice reset complete');
}

/**
assign run number based on operator sign up slip

@PARAM {number} badge number of the operator
@PARAM {list} taken: list of runs already taken
@PARAM {number} assignMode: legacy mode
@return {number} run number assigned to operator


@custom function
*/

function assignRun_(badge,taken,assignMode) {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  //SpreadsheetApp.getActiveSpreadsheet();

  //console.info('In assignRun');
  //badge = 6004
  //var taken = findTaken(badge);
  //console.log('taken list ' + taken);
  var selectionList = findSelection_(badge,assignMode);
  //console.log('selection list '+selectionList);
  var shiftChoice = makeSelection_(taken, selectionList)
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

function makeSelection_(takenList,selectionList){
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
* 
* @PARAM {number} badge number of the operator
* @RETURN {list} of operator selection
* @custom function
*/

function findSelection_(badge,assignMode){
  if (assignMode == 1){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Slips');
  } else {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Slips3');
  }
  //SpreadsheetApp.setActiveSheet(ss.getSheets()[2]);
  //var sheet = ss.getSheetByName("sheetName");
  //var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var badgeCol = values[0].indexOf('Badge');
  var choice1Col = values[0].indexOf('Choice 1');

  var selected = []

  //var badgeRange = sheet.

  for (var i = 1; i < values.length; i++) {

    if (values[i][badgeCol] === badge){
      for (j = choice1Col;j<values[i].length;j++){
        selected.push(values[i][j]);
      }
      
      //SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
      //SpreadsheetApp.getActiveSpreadsheet();
      return selected;
    }
  }  
}


/**
 * shows the list of all sheets in the curent SS
 */
function showSheetList_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i<sheets.length; i++){
    console.log(sheets[i].getName());
  }
}


/**
 *  gets the sheet that corresponds to the response
 */
function getFormResponseSheet_(wkbkId, formUrl) {
  const matches = SpreadsheetApp.openById(wkbkId).getSheets().filter(
    function (sheet) {
      return sheet.getFormUrl() === formUrl;
    });
  return matches[0]; // a `Sheet` or `undefined`
}



