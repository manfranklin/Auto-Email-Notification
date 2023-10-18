/**
 * Author: Manuel F. Pedro
 * Date: 2023-10-18
 * Description: This function creates a new Google Docs document using data from Google Sheets and a predefined template. It replaces placeholders with the data and saves the new document to a specified folder. This function is useful for generating customized documents from a google form response sheet.
 * 
 * Assuming that the google sheets file contains the following place holders 
 * {{TIMESTAMP}}, {{NAME}}, {{POSITION}},{{REMARK}}
 *  
 * and that the google doc file contains the following header in the first row:
 * Timestamp, Name, Position, Remarks
 */

function createDocFromTemplate() {
  // Define the maximum number of retries and initialize retry count.
  var maxRetries = 2;
  var retryCount = 0;

  // Retry the operation in case of errors.
  while (retryCount < maxRetries) {
    try {
      // Get the active spreadsheet and the specific sheet.
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('Responses');
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var lastRow = sheet.getLastRow();
      var incrementLink = lastRow-1;

      // Helper functions to get column index and data.
      var getColumnIndex = (columnName) => headers.indexOf(columnName);
      var getData = (columnName) => data[lastRow - 1][getColumnIndex(columnName)];

      // Template and folder IDs.
      var templateDocId = 'REPLACE WITH YOUR GOOGLE DOC TEMPLATE ID';
      var folderId = 'REPLACE WITH YOUR GDRIVE FOLDER ID';

      // Get the timestamp, format it, and store other data.
      var timestamp = new Date(getData('Timestamp'));
      var formattedDate = Utilities.formatDate(timestamp, 'GMT', 'yyyy-MM-dd');

      // Placeholders for document replacement.
      var placeholders = {
        '{{TIMESTAMP}}': formattedDate,
        '{{NAME}}': getData('Name'),
        '{{POSITION}}': getData('Position'),
        '{{REMARK}}': getData('Remarks')
      };
      // Open the template document and create a new copy.
      var templateDoc = DriveApp.getFileById(templateDocId);
      var newDoc = templateDoc.makeCopy('PDP-' + incrementLink +'-'+placeholders['{{NAME}}'], DriveApp.getFolderById(folderId));
      var doc = DocumentApp.openById(newDoc.getId());
      var body = doc.getBody();

      // Replace placeholders in the document.
      for (var placeholder in placeholders) {
        body.replaceText(placeholder, placeholders[placeholder]);
      }

      // Save and close the document.
      doc.saveAndClose();

      // Exit the loop if successful.
      break;
    } catch (e) {
      // Handle errors, log the error, increment retry count, and wait before retrying.
      console.log('Error: ' + e.toString());
      retryCount++;
      Utilities.sleep(5000);
    }
  }
}
