//Global variable that contains information about all spreadsheets and can be used with both functions
var SHEETINFO = SchoolInfo.SheetInfo;

function moveRow(e) {
  //@NotOnlyCurrentDoc
  //
  /* Do not change items or variables in this script. These items are imported from a library for all the 
     spreadsheets for school reassignment. Go to the library in the school reassignment folder called "SchoolInfo" 
     and make changes there. You should not have to change anything in this script, and the script should work 
     for all spreadsheets.
     To set up this script in each sheet make sure you: 
     - Import the library by clicking on "Resources" >> "Libraries.." and then adding the script id from "SchoolInfo" >> 
       "File" >> "Project properties" >> "Info" >> "Script ID" to the empty box to import the library. 
     - Set up triggers for both moveRow() and reValidate() functions
     - Insure that "File" >> "Project properties" >> "Scopes" lists both: 
       2 OAuth Scopes required by the script:
       https://www.googleapis.com/auth/script.send_mail
       https://www.googleapis.com/auth/spreadsheets
       If not, go to "View" >> "Show manifest file" and add: 
       "oauthScopes": ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/script.send_mail"],
       to appscript.json
  */
  // Create an array with the required values by column (for each sheet): 
  var required = [];
  var requiredTo = SchoolInfo.RequiredTo; 
  var requiredFrom = SchoolInfo.RequiredFrom; 
  var requiredFinal = SchoolInfo.RequiredFinal;
  // Set the column number to activate the script
  var column;
  var colTo = SchoolInfo.ColTo; 
  var colFrom = SchoolInfo.ColFrom; 
  var colFinal = SchoolInfo.ColFinal; 
  // Declare the move messages
  var moveMsg;
  var toMoveMsg = SchoolInfo.ToMoveMsg;
  var fromMoveMsg = SchoolInfo.FromMoveMsg;
  var finalMoveMsg = SchoolInfo.FinalMoveMsg;
  // Custom Content: 
  var uiConfirmAlert = SchoolInfo.UiConfirmAlert;
  var uiConfirmMsg = SchoolInfo.UiConfirmMsg;
  var uiCancelAlert = SchoolInfo.UiCancelAlert;
  // Num of emails to be sent
  var ssNum;
  var ssNumSchool = SchoolInfo.SsNumSchool;
  var ssNumFinal = SchoolInfo.SsNumFinal;
  // Get the spreadsheet bound to this script
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = ss.getName();
  // Choose the correct array based on the name of the spreadsheet
  var info = [];
  if (spreadsheetName == SHEETINFO[0][8]) { info = SHEETINFO[0]; }
  else if (spreadsheetName == SHEETINFO[1][8]) { info = SHEETINFO[1]; }
  else if (spreadsheetName == SHEETINFO[2][8]) { info = SHEETINFO[2]; }
  else if (spreadsheetName == SHEETINFO[3][8]) { info = SHEETINFO[3]; }
  else if (spreadsheetName == SHEETINFO[4][8]) { info = SHEETINFO[4]; }
  else if (spreadsheetName == SHEETINFO[5][8]) { info = SHEETINFO[5]; }
  else if (spreadsheetName == SHEETINFO[6][8]) { info = SHEETINFO[6]; }
  else if (spreadsheetName == SHEETINFO[7][8]) { info = SHEETINFO[7]; }
  else if (spreadsheetName == SHEETINFO[8][8]) { info = SHEETINFO[8]; }
  else if (spreadsheetName == SHEETINFO[9][8]) { info = SHEETINFO[9]; }
  // Get the location of the edit on the spreadsheet
  var editRange = e.range;
  // Get the name of the sheet where the cell edit took place
  var sheetName = editRange.getSheet().getSheetName();
  // Get the active sheet by name
  var fromSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // Get the column number (A = 1, B = 2, and so on...) where the edit took place
  var columnOfCellEdited = editRange.getColumn();
  // 
  if (sheetName === info[1] && sheetName != SHEETINFO[6][1]) {
    column = colTo;
    moveMsg = toMoveMsg;
    required = requiredTo;
    ssNum = ssNumSchool;
  } else if (sheetName === info[2] && sheetName != SHEETINFO[6][2]) {
    column = colFrom;
    moveMsg = fromMoveMsg;
    required = requiredFrom;
    ssNum = ssNumSchool;
  } else if (sheetName === SHEETINFO[6][3]) {
    column = colFinal;
    moveMsg = finalMoveMsg;
    required = requiredFinal;
    ssNum = ssNumFinal;
  }
  /////////////////////////////////////////////////////////////////////////////////////////////////////////
  // If the column number equals curCol and the sheet is the first sheet...
  if (columnOfCellEdited === column) { 
    // Get the range: getRange(row, column, numRows, numColumns)
    var range = fromSheet.getRange(3, 1, fromSheet.getLastRow()-2, fromSheet.getLastColumn());
    // Get all the rows with data (values)
    var values = range.getValues();
    // Get browser interface
    var ui = SpreadsheetApp.getUi();
    /////////////////////////////////////////////////////////////////////////////////////////////////////////
    // This will loop through each row in the spreadsheet (each row that isn't frozen at the top)
    for (i in values) {
      // Create an error message
      var error = [];
      // Create a counter
      var count = 0;
      // The number of the current row
      var curRow = parseInt(i) + 3; // Add three to account for starting with 0 in an array and the 2 top frozen rows
      // Get the from school sheet from the spreadsheet to which the data will be copied
      var fromRange = fromSheet.getRange(curRow, 1, 1, column - 1); 
      // Range from which the data will be moved
      var move = fromSheet.getRange(curRow, column, 1, 1);
      // Position of the "move" cell
      var moveCell = parseInt(column)-1;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////
      // The below part of the script only runs if the user selects "Move" in column
      // Creates an error message if there are cells empty that are supposed to be filled
      if (values[i][moveCell] == moveMsg) { // i = row position in array, moveCell = the column position in the array
        // j cycles through each column on row i, which is the row with "Move" in it
        for (var j = 0; j < values[i].length - 1; j++) { // The - 1: since the row should always be the same length, we can cut off the "Move"
          // k cycles through all the values in the "required" array
          for (var k = 0; k < required.length; k++) {
            var curCol = parseInt(required[k]); 
            var curColArr = curCol - 1; // This variable is the position in the preValues array for required fields, only chk this
            if (parseInt(j) + 1 == curCol && values[i][curColArr] == '') {
              var cell = fromSheet.getRange(curRow, curCol, 1, 1).getA1Notation();
              error.push('Cell '+cell+' may not be left blank.'); // Throws an error if a required cell blank
              count++;
            }
          }
        }
        // Below creates the actual error message
        if (count > 0) { // If the count of errors (empty cells) is zero then...
          var errorMsg = error.join(' ');
          ui.alert(
            errorMsg,
            ui.ButtonSet.OK
          );
          move.clearContent();
          break; // The loop stops as soon as there is an cell empty that is supposed to be filled
        }
        var sender = [];
        // Add sender info to emailInfo
        sender.push(info[5]); //fname
        sender.push(info[6]); //lname
        sender.push(info[7]); //email
        //
        // This variable holds the school the request is being sent to (destination sheet)
        var recSchool = [];
        // If the name of the current sheet is "To" then the destination sheet will be a "From", else to County Office
        if (sheetName === info[1] && sheetName != SHEETINFO[6][1]) { // This is the origin sheet "To [School]"
          if (values[i][13] === "Employee"|| values[i][21] === "NC"){//EAF added - checking for Employee or No Concern in Discipline column. 
          recSchool.push(SHEETINFO[6][1]);//EAF added - send Employee records straight to Final Approval Sheet
        }else {recSchool.push(values[i][9]);}//EAF changed values[i][3] to values[i][9] (School of Residence)
        } 
        else if (sheetName === info[2] && sheetName != SHEETINFO[6][2]) { // This is the origin sheet "From [School]"
          recSchool.push(SHEETINFO[6][1]);
        } else if (sheetName === SHEETINFO[6][3]) { // This is the origin sheet "School Reassignment - Final Approval"
          for (var y = 0; y < SHEETINFO.length; y++) {
            if (values[i][11] === SHEETINFO[y][4] || values[i][9] == SHEETINFO[y][2]) { //EAF changed (values[i][2](School of ) === SHEETINFO[y][4] || values[i][3] == SHEETINFO[y][2]) to (values[i][11] === SHEETINFO[y][4] || values[i][9] == SHEETINFO[y][2])
              recSchool.push(SHEETINFO[y][3]);
            } 
          }
        }
        //Logger.log(recSchool);
        ///EAF added send an email of the log to the active user
//var recipient = Session.getActiveUser().getEmail();
//var subject = 'Log from Student Reassignment Script - CES';
//var body = Logger.getLog();
//MailApp.sendEmail(recipient, subject, body);
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Create an array with the spreadsheet ID
        var spreadsheetId = [];
        // Create an array to store recipient info
        var recipient = [];
        // Loop through the sheetInfo array above 
        for (var l = 0; l < SHEETINFO.length; l++) {
          // if the a value from the sheetInfo array in the school the data will be
          // copied to matches the indicated school in the school column...
          // has to match the 3rd element in array, the "From" school...
          // 2 = "From [School]" & 4 = just school initials & 3 = the Final Decision sheet (#3)
          for (var w = 0; w < recSchool.length; w++) {
            if (SHEETINFO[l][2] === recSchool[w]) {
              // Save one value to the array so it can be used outside the loop 
              spreadsheetId.push(SHEETINFO[l][0]);
              // Add recipient info to emailInfo
              recipient.push([SHEETINFO[l][5],SHEETINFO[l][6],SHEETINFO[l][7]]); //fname, lname, email, spreadsheet url
            } else if (SHEETINFO[l][3] === recSchool[w]) {
              spreadsheetId.push(SHEETINFO[l][0]);
              // Add recipient info to emailInfo
              recipient.push([SHEETINFO[l][5],SHEETINFO[l][6],SHEETINFO[l][7]]); //fname, lname, email, spreadsheet url
            }
          }
        }
        // Create a pop up window
        var result = ui.alert(
           uiConfirmAlert,
           uiConfirmMsg,
        ui.ButtonSet.YES_NO);
        // If you click Yes then...
        if (result == ui.Button.YES) {
          // Clear "move" from the cell in the origin
          move.clearContent();
          // Clear the data validation from the cell
          move.clearDataValidations();
          // Set the value of the "move" cell to "Moved." 
          move.setValue("SENT");
          // Hide the row
          fromSheet.hideRow(fromRange); 
          // Protect the row in the "SENT" row
          var protection1 = fromRange.protect().setDescription('Do not edit. Already completed by school of choice.');
          protection1.setWarningOnly(true); // It's a warning only
          // Create a loop to append rows and send emails
          for (var x = 0; x < spreadsheetId.length; x++) {
            // Access the spreadsheet to which the data will be copied
            var toSS = SpreadsheetApp.openById(spreadsheetId[x]);
            // Record the url of the spreadsheet to send in the email
            var toUrl = toSS.getUrl();
            // The range of data on the origin sheet
            var toSheet = toSS.getSheetByName(recSchool[x]);
            // Get the integer that represents the id number of the sheet in the spreadsheet. This is used
            // so that when the email opens the spreadsheet, it will open to the specific sheet for the 
            // request. 
            var toSheetID = toSheet.getSheetId();
            // Add the data from the origin sheet to the destination sheet
            toSheet.appendRow(values[i]);
            // Clear "move" from the cell in the origin
            toSheet.getRange(toSheet.getLastRow(), column, 1, 1).clearContent();
            // The range to which the values will be moved
            var toRange = toSheet.getRange(toSheet.getLastRow(), 1, 1, parseInt(column) - 1);
            // Protect the row on the destination sheet with a warning so that data will not be changed by the wrong person          
            var protection2 = toRange.protect().setDescription('Do not edit. Already completed by school of choice.');
            protection2.setWarningOnly(true); // It's a warning only
            // Clear move from the destination sheet cell 11
            // Then send email
            var subject =  'Re: School Reassignment';
            var greeting = '<p>'+recipient[x][0]+',</p><p>'+sender[0]+' '+sender[1];
            var mainBody = ' has completed a school reassignment request. '+
                           'To view this request, <a href="'+toUrl+'#gid='+toSheetID+'">please click here.</a></p>';
            var body = greeting + mainBody;
            MailApp.sendEmail({
              to: recipient[x][2],
              subject: subject, 
              htmlBody: body, 
            });    
          }
        } else {
          // This cancels and clears the "move" box
          ui.alert(uiCancelAlert);
          move.clearContent();
        }
      }
    }
  }
}

// These two functions can be run manually to update the formatting of the spreadsheet
// This function will clear the conditional formatting
function clearFormatting () {
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var ss = s.getActiveSheet();
  ss.clearConditionalFormatRules();
}
// This spreadsheet sets all conditional formatting and validation for all three sheets
function formatAndValidate() {
  // Get the spreadsheet bound to this script
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = ss.getName();
  var info = [];
  if (spreadsheetName == "Cooleemee Student Reassignment") { info = ["To CES", "From CES", "CES", "CES Final Decision"]; }
  else if (spreadsheetName == "Cornatzer Student Reassignment") { info = ["To CZES", "From CZES", "CZES", "CZES Final Decision"]; }
  else if (spreadsheetName == "Mocksville Student Reassignment") { info = ["To MES", "From MES", "MES", "MES Final Decision"]; }
  else if (spreadsheetName == "Pinebrook Student Reassignment") { info = ["To PES", "From PES", "PES", "PES Final Decision"]; }
  else if (spreadsheetName == "Shady Grove Student Reassignment") { info = ["To SGES", "From SGES", "SGES", "SGES Final Decision"]; }
  else if (spreadsheetName == "William R. Davie Student Reassignment") { info = ["To WRD", "From WRD", "WRD", "WRD Final Decision"]; }
  var toSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info[0]);
  var fromSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info[1]);
  var finalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info[3]);
  // *****--------*****--------*****--------*****--------*****--------*****--------*****--------*****--------*****--------*****--------
  // Add validation rules for the to sheet
  //Deleting "Have they filled out SR1?", so commenting out validation for column A
  //var colA = SpreadsheetApp.newDataValidation()
  //  .setAllowInvalid(false)
  //  .requireValueInList(['Yes', 'No'])
  //  .build();
  //toSheet.getRange(3, 1, toSheet.getMaxRows()-2, 1).setDataValidation(colA);
  var colE = SpreadsheetApp.newDataValidation()//Only allow certain request types
    .setAllowInvalid(false)
    .requireValueInList(['New','Renewal'])
    .build();
  toSheet.getRange(3, 5, toSheet.getMaxRows()-2, 1).setDataValidation(colE);
  var colJ = SpreadsheetApp.newDataValidation()//give choices of From schools (School of Residence)
    .setAllowInvalid(false)
    .requireValueInList(['From CES','From CZES','From MES','From PES','From SGES','From WRD'])
    .build();
  toSheet.getRange(3, 10, toSheet.getMaxRows()-2, 1).setDataValidation(colJ); 
  var colL = SpreadsheetApp.newDataValidation()//set only option to the school of choice for this sheet
    .setAllowInvalid(false)
    .requireValueInList([info[2]])
    .build();
  toSheet.getRange(3, 12, toSheet.getMaxRows()-2, 1).setDataValidation(colL);
  var colN = SpreadsheetApp.newDataValidation()//Only allow certain applicant types
    .setAllowInvalid(false)
    .requireValueInList(['Employee','Non Employee'])
    .build();
  toSheet.getRange(3, 14, toSheet.getMaxRows()-2, 1).setDataValidation(colN);
  var colY = SpreadsheetApp.newDataValidation()//School of choice Recommend
    .setAllowInvalid(false)
    .requireValueInList(['Approve', 'Disapprove'])
    .build();
  toSheet.getRange(3, 25, toSheet.getMaxRows()-2, 1).setDataValidation(colY);
  var colZ = SpreadsheetApp.newDataValidation()//Require valid date
   .setAllowInvalid(false)
   .requireDate()
   .build();
  toSheet.getRange(3, 26, toSheet.getMaxRows()-2, 1).setDataValidation(colZ);
  var colAB = SpreadsheetApp.newDataValidation()//Only allow correct text for commands
   .setAllowInvalid(false)
   .requireValueInList(['Move'])
   .build();
  toSheet.getRange(3, 28, toSheet.getMaxRows()-2, 1).setDataValidation(colAB);
  toSheet.getRange(3, 1, fromSheet.getMaxRows()-2, 22).setBackground("#ececec");
  //
  var disapprove = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Disapprove")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#FF0000")
    .setRanges([toSheet.getRange(1,1,toSheet.getMaxRows(),toSheet.getMaxColumns())])
    .build();
  var disapproves = toSheet.getConditionalFormatRules();
  disapproves.push(disapprove);
  toSheet.setConditionalFormatRules(disapproves);
  //
  var approve = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Approve")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#055719")
    .setRanges([toSheet.getRange(1,1,toSheet.getMaxRows(),toSheet.getMaxColumns())])
    .build();
  var approves = toSheet.getConditionalFormatRules();
  approves.push(approve);
  toSheet.setConditionalFormatRules(approves);
    //
  var no = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#FF0000")
    .setRanges([toSheet.getRange(1,1,toSheet.getMaxRows(),toSheet.getMaxColumns())])
    .build();
  var nos = toSheet.getConditionalFormatRules();
  nos.push(no);
  toSheet.setConditionalFormatRules(nos);
  //
  var yes = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Yes")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#055719")
    .setRanges([toSheet.getRange(1,1,toSheet.getMaxRows(),toSheet.getMaxColumns())])
    .build();
  var yess = toSheet.getConditionalFormatRules();
  yess.push(yes);
  toSheet.setConditionalFormatRules(yess);
  // *****--------*****--------*****--------*****--------*****--------*****--------*****--------*****--------*****--------*****--------
  // Add validation rules for the from sheet
  var colAC = SpreadsheetApp.newDataValidation()//School of Residence Recommend column
   .setAllowInvalid(false)
   .requireValueInList(['Approve', 'Disapprove'])
   .build();
  fromSheet.getRange(3, 29, fromSheet.getMaxRows()-2, 1).setDataValidation(colAC);
  //var colM = SpreadsheetApp.newDataValidation()//Remove Eligible question
  // .setAllowInvalid(false)
  // .requireValueInList(['Yes', 'No'])
  // .build();
  //fromSheet.getRange(3, 13, fromSheet.getMaxRows()-2, 1).setDataValidation(colM);
  var colAD = SpreadsheetApp.newDataValidation()//Require valid date from residence school
   .setAllowInvalid(false)
   .requireDate()
   .build();
  fromSheet.getRange(3, 30, fromSheet.getMaxRows()-2, 1).setDataValidation(colAD);
  var colAF = SpreadsheetApp.newDataValidation()//only allow correct command to move
   .setAllowInvalid(false)
   .requireValueInList(['Move Decision'])
   .build();
  fromSheet.getRange(3, 32, fromSheet.getMaxRows()-2, 1).setDataValidation(colAF);
  fromSheet.getRange(3, 1, fromSheet.getMaxRows()-2, 27).setBackground("#ececec");
    //
  var disapprove = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Disapprove")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#FF0000")
    .setRanges([fromSheet.getRange(1,1,fromSheet.getMaxRows(),fromSheet.getMaxColumns())])
    .build();
  var disapproves = fromSheet.getConditionalFormatRules();
  disapproves.push(disapprove);
  fromSheet.setConditionalFormatRules(disapproves);
  //
  var approve = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Approve")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#055719")
    .setRanges([fromSheet.getRange(1,1,fromSheet.getMaxRows(),fromSheet.getMaxColumns())])
    .build();
  var approves = fromSheet.getConditionalFormatRules();
  approves.push(approve);
  fromSheet.setConditionalFormatRules(approves);
    //
  var no = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#FF0000")
    .setRanges([fromSheet.getRange(1,1,fromSheet.getMaxRows(),fromSheet.getMaxColumns())])
    .build();
  var nos = fromSheet.getConditionalFormatRules();
  nos.push(no);
  fromSheet.setConditionalFormatRules(nos);
  //
  var yes = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Yes")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#055719")
    .setRanges([fromSheet.getRange(1,1,fromSheet.getMaxRows(),fromSheet.getMaxColumns())])
    .build();
  var yess = fromSheet.getConditionalFormatRules();
  yess.push(yes);
  fromSheet.setConditionalFormatRules(yess);
  // Add validation rules for the storage sheet
  finalSheet.getRange(3, 1, fromSheet.getMaxRows()-2, 33).setBackground("#ececec");
  //
  var disapprove = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Disapprove")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#FF0000")
    .setRanges([finalSheet.getRange(1,1,finalSheet.getMaxRows(),finalSheet.getMaxColumns())])
    .build();
  var disapproves = finalSheet.getConditionalFormatRules();
  disapproves.push(disapprove);
  finalSheet.setConditionalFormatRules(disapproves);
  //
  var approve = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Approve")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#055719")
    .setRanges([finalSheet.getRange(1,1,finalSheet.getMaxRows(),finalSheet.getMaxColumns())])
    .build();
  var approves = finalSheet.getConditionalFormatRules();
  approves.push(approve);
  finalSheet.setConditionalFormatRules(approves);
    //
  var no = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#FF0000")
    .setRanges([finalSheet.getRange(1,1,finalSheet.getMaxRows(),finalSheet.getMaxColumns())])
    .build();
  var nos = finalSheet.getConditionalFormatRules();
  nos.push(no);
  finalSheet.setConditionalFormatRules(nos);
  //
  var yes = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Yes")
    .setBold(true)
    .setFontColor("#FFFFFF")
    .setBackground("#055719")
    .setRanges([finalSheet.getRange(1,1,finalSheet.getMaxRows(),finalSheet.getMaxColumns())])
    .build();
  var yess = finalSheet.getConditionalFormatRules();
  yess.push(yes);
  finalSheet.setConditionalFormatRules(yess);
}
