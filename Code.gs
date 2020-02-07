// THIS SCRIPT WILL READ DATA FROM A FROM ONE SHEET AND THEN SCAN A 
// SECOND SHEET TO FIND THE PARENT ACCESS CODE AND EMAIL THE PARENT THE CODE

// GLOBAL VARIABLES
var target; // Student ID
var code; // Access code
var rows = 35000; // Rows in Parent Access Code sheet from Schoology
var fName; // Students first name
var lName; // Students last name
var school; 
var parentEmail; // email to send the access code to
var completion; // true = Access code sent false = access code not found
var DATE = "01/24/2020"  // This is the last updated date sent to parents in the code not found email

function main() // RUNS all the functions
{
  
  // Grab sheet 
  var sheet = SpreadsheetApp.getActiveSheet();
  var ss = SpreadsheetApp.openById("15N38dApSNGwEglQKE9d1ORIjGu4CsE4cdvWpbqreTcA"); //Request Parent Access Code (AUTO) (Responses)
  var activeSheet = ss.getSheets()[1]; // "access data on different tabs"
  ss.setActiveSheet(sheet); // First Sheet
  
  // Grab info to look up
  GrabStudentID();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  
  // Returns the active range
  var range = sheet.getActiveRange();
  
  var data1 = sheet.getRange(2,4).getValue(); // Grabs Parent Access code of row 2
 // Logger.log(data1);
  
  
  // Grabs the 35000 Parent Access codes in 2d array
  var numRows = 35000;
  // startRow, startColumn, numRows, NumColumns
  var values = sheet.getSheetValues(1, 1, numRows, 4);
  
 // Jump start the loop - To shorten run time
 var jumpTo =0;
 if(target>=63898343) // Find this row in the CODES sheet, and minus 1 from that row to assign jumpTo value
    jumpTo=16212;
  
 // Loop Throught the Data
 completion = false;
 for(var i=0+jumpTo; i<numRows; i++)
 {
      // Do this if matching StudentID if found
      if(values[i][0]==target)
      {    
        Logger.log("TARGET FOUND");
        Logger.log(values[i][3]); // Access Code
        code = values[i][3];
        fName = values[i][2];
        lName = values[i][1];
        emailParent();
        completion = true; // Code was found, end loop
      }
 }
 if(completion==false) // NO studentID found
 {
    emailParentNotFound(); // Tell the parent the code wasn't found
 }
  
 markEmailed(); // Mark on the spreadsheet 1 that the parent was emailed
  
}

// This function grabs the last form entries Student ID & Email 
function GrabStudentID()
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0]; // Changes the Sheet Tab to look at
                                // [0] reads Form Response Sheet
  

 // This logs the value in the very last cell of this sheet
 var lastRow = sheet.getLastRow();
 var lastID = sheet.getRange(lastRow, 3);// Student ID column
 var lastEmail = sheet.getRange(lastRow, 2); // Email Address column
 var lastSchool = sheet.getRange(lastRow,7); // School column

 // Assign values to global variables 
 target = lastID.getValue();
 parentEmail = lastEmail.getValue();
 school = lastSchool.getValue();
 
 // DEBUGGING  
 Logger.log(target);
 Logger.log(parentEmail);

}

// Craft a message and email the access code to the parent
function emailParent()
{
    // format email message
    var message = "Parent Access code for " + fName +", " + lName + " ID: " + target + " is" + "\n\n"+ code + 
    "\n\nCopy and paste the code above and paste it into the sign up link below to create your parent account \n" +
    "\nhttps://app.schoology.com/register.php?type=parent\n\nAdditional Schoology information can found at Parent Portal\n\nhttps://sites.google.com/apps.district279.org/parentportal/schoology"; 

    // format subject of the email
    var subject = "Schoology - Parent Access Code - " + lName + ", " + fName + " - " + school;
    
    // send the email
    MailApp.sendEmail(parentEmail, subject, message);
}

// Craft a message saying the code was not found and email the parent
function emailParentNotFound()
{
    // format email message
    var message = "Parent Access code NOT FOUND\n\n" 
    
    +"Possible Reasons Why...\n\n"
    
    + "Only students enrolled before " + DATE
    + " will have codes available for email"
    + "\n\nAn incorrect student ID was entered, double check that " + target + " is correct"
    + "\n\n*This is an automatic response email*";
   
    // format subject of the email
    var subject = "Schoology - Parent Access Code - NOT FOUND";

    // send the email
    MailApp.sendEmail(parentEmail, subject, message);
  
}

function markEmailed()
{
  // Set EMAIL
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // Changes the Sheet Tab to look at
                                 // [0] reads Form Response Sheet
  var lastRow = sheet.getLastRow();
  var lastEmail = sheet.getRange(lastRow,8);
  
  // Mark Sheet with status
  if(completion==true)
    lastEmail.setValue("EMAIL_SENT_COMPLETE");
  else 
    lastEmail.setValue("CODE_NOT_FOUND");
}