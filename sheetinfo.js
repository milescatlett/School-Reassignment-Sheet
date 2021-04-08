//@NotOnlyCurrentDoc
// This is a library of constants used for all spreadsheets
// Create an array with:
// 0 = spreadsheet ID; 1 = sheet #1 name; 2 = sheet #2 name; 3 = sheet #3 name; 4 = school initials; 5 = principal first name;
// 6 = principal last name; 7 = principal email (using mine for testing, uncomment below for actual emails); 8 = spreadsheet name;
// Entries 0 - 9 (updated on 5/26)
var SheetInfo = [
    ["1WtxfvtpxvhNJCABgHGLwf78hdspbYuA6MGGw9BCA-AA","To School1","From School1","School1 Final Decision","School1","FirstName","LastName","Email","School1 Student Reassignment"],
    ["1S-wJSIMywDp7A5wVv6FapOfEDMAxe2n7t2eZiXAyDAo","To School2","From School2","School2 Final Decision","School2","FirstName","LastName","Email","School2 Student Reassignment"],
    ["1knE8FBUhi3PWjTe3WxS2Clxy0K1TJouS8EhKu2oRHEU","To School3","From School3","School3 Final Decision","School3","FirstName","LastName","Email","School3 Student Reassignment"],
    ["1b-0gT0PhxvVH2zNk63o9U4No_7Zg-cJ_4tvfs9kGH2E","To School4","From School4","School4 Final Decision","School4","FirstName ","LastName","Email","School4 Student Reassignment"],
    ["1rtZtkT0mY3SN22A49VNtmcSAvmmrPnqARsRnNmaa0qs","To School5","From School5","School5 Final Decision","School5","FirstName","LastName","Email","School5 Student Reassignment"],
    ["1NGfMjE6AnwCJWXidJx2_6Yi3Gw8_lThGWxYxDqOkt4o","To School6","From School6","School6 Final Decision","School6","FirstName","LastName","Email","William R. Davie Student Reassignment"],
    ["1nr6KqgxbqokMc-hq0gIFNsfSlGcBOwAanpKjI6bACZ0","DCS Student Reassignments","DCS Student Reassignments","DCS Student Reassignments","SI","County","Office","Email","Student Reassignment"],
    ["1hmj1SR0CBbsnw000K1tiXpOZVY8qC_cwrI7uEyoT3pQ","To School7","From School7","School7 Final Decision","School7","FirstName","LastName","Email","School7 Student Reassignment"],  
    ["1rqbfRfmkoOAd4LFbQTpnsJ6YZ7gvoSBswoh5z5Qfsbg","To School8","From School8","School8 Final Decision","School8","FirstName","LastName","Email","School8 Student Reassignment"],
    ["1042WHbPlewyZoP7pg3gsgCYMq8Kw0VVj3BSVeYmKmZg","To School9","From School9","School9 Final Decision","School9","FirstName","LastName","Email","School9 Student Reassignment"]
  ];
// Create an array with the required values by column (for each sheet): 
var RequiredTo = [1,2,3,4,5,6,7,9,10,11,12,14,23,25,26,27]; // These are the required columns in "To [School]"
var RequiredFrom = [1,2,3,4,5,6,7,9,10,11,12,14,23,25,26,27,29,30,31]; // These are the required columns in "From [School]"
var RequiredFinal = [1,2,3,4,5,6,7,9,10,11,12,14,25,26,27,32,33]; // These are the required columns in "Final Decision"
// Set the column numbers to activate the script
var ColTo = 28; // for the "To" sheet
var ColFrom = 32; // for the "From" sheet
var ColFinal = 34; // for the School Reassignment - Final Approval spreadsheet
// These are the confirmation alerts for the pop up windows
var UiConfirmAlert = 'Please confirm';
var UiConfirmMsg = 'The current row will be moved. This action cannot be undone.';
var UiCancelAlert = 'Operation canceled.';
// Declare the number of emails to be sent
var SsNumSchool = 1;
var SsNumFinal = 2;
// Declare the "move message" in the cell intended to move the row to the next sheet
var ToMoveMsg = "Move";
var FromMoveMsg = "Move Decision";
var FinalMoveMsg = "Move FINAL";
