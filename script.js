function docFunction() {
// Get document template, copy it as a new temp doc, and save the Doc’s id
// Change the values below
  //--------------------------------------------------------------------------------------------------------------------------------
  var startRow = 1; // If the script fails, find the last email sent on the "Record" sheet, find row num for that student on "Ids" sheet, restart
  var docTemplate = "ID"; // *** replace with your template ID ***
  var docName = "Grades Spring 2020"; // *** specify the name of the document ***
  var sheetName = "Sheet1"; // ***specify the name of the sheet you want to pull data from***
  var idSheetName = "IDs";
  var recordSheetName = "Record";
  var sharefolder = DriveApp.getFolderById('ID'); // *** Change each year from newly created folders
  //--------------------------------------------------------------------------------------------------------------------------------
  // Get the sheets
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var ids = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(idSheetName);
  var record = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(recordSheetName);
  // Fetch the range of cells for the sheets
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow()-1, 12); //startRow, startCol, numRows, numCols
  var idRange = ids.getRange(startRow, 1, ids.getLastRow()-1, 5); //startRow, startCol, numRows, numCols
  // Fetch values for each row in the Range of all the sheets
  var data = dataRange.getValues();
  var dataId = idRange.getValues();
  // Make a cell to record the email quota each time the script is run
  var quotaCount = record.getRange("E2");
  // Get the date and time of access of quota
  var today = new Date();
  // create the actual count in numerical form
  var quota = parseInt(MailApp.getRemainingDailyQuota());
  // Clear the cell of the previous quota
  quotaCount.clearContent();
  // Place new quota with date/time
  quotaCount.setValue(quota+' ('+today+')');
  //
  for (var i = 0; i < ids.getLastRow()-1; i++) { 
    if (quota > 10) {
      var row = dataId[i];
      var stuNum = row[0];   // A column, Student Number
      var stuLast = row[1];  // B column, Student Last Name
      var stuFirst = row[2]; // C column, Student First Name
      var email = row[4];    // E column, Student Email
      //var email = 'foile@davie.k12.nc.us'; // FOR TESTING, uncomment above line & recomment this one*****
      var name = stuFirst+' '+stuLast;
      var documentName = name+': '+docName;
      var copyId = DriveApp.getFileById(docTemplate).makeCopy(documentName, sharefolder).getId();
      var link = DriveApp.getFileById(copyId).getUrl();
      // Open the temporary document
      var copyDoc = DocumentApp.openById(copyId);
      // Get the document’s body section
      var copyBody = copyDoc.getActiveSection();
      //Add in courses with if then and copybody commands
      copyBody.replaceText('<<Student First>>', stuFirst);
      copyBody.replaceText('<<Student Last>>', stuLast);
      var grades = []; // Make sure transcript grade is the one showing
      for (var j = 0; j < sheet.getLastRow()-1; j++) {
        if (stuNum == data[j][0]) {
          grades.push([data[j][6],data[j][11]]);
        }
      }
      for (var q = 0; q < 7; q++) {
        var crsName;
        var crsGrade;
        if (typeof grades[q] != 'undefined') {
          crsName = grades[q][0];
          crsGrade = grades[q][1];
        } else if (typeof grades[q] == 'undefined') {
          crsName = ' ';
          crsGrade = ' ';
        }
        copyBody.replaceText('<<Crs Name '+(q+1)+'>>', crsName);
        copyBody.replaceText('<<Transcript Grade '+(q+1)+'>>', crsGrade);
      }
      // Save and close the document
      copyDoc.saveAndClose();
      copyDoc.addViewer(email);
      //This email goes to the student
      // *** specify subject of email ***
      var subject = "CONFIDENTIAL: "+docName;
      // *** specify body text of email ***
      var body = "<p>" + name + ",</p><p> Attached is a file that explains how your grades from this semester will show on your transcript. " +
        "Please review it carefully and respond through the google form if there is something you would like to change. </p> " +
          "<p><a href='"+link+"'>Click here for a copy of your Grades Spring 2020 Document "+
        "</a>. "; 
      MailApp.sendEmail({
        to: email,
        bcc: "Email", // Also need to uncomment this line to not get all those emails...****
        subject: subject, 
        htmlBody: body, 
        noReply: true
      });
      record.appendRow([name+'-'+stuNum,'Email SENT',link]);
    }
  }
}
