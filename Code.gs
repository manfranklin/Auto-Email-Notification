/**
 * Author: Manuel F. Pedro
 * Date: 2023-09-27
 * Description: This function reads a Google Sheets file and sends renewal notifications from to users when the renewal date is within 45 days.
*/

function sendNotification() {
  // Define the sheet where your data is located
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the data range
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0]; // Assuming the first row contains headers

  // Get the current date 
  const currentDate = new Date();

// Set how many days in advance you want to receive the notification
const days_in_advance = 45;

  // Loop through each row in the data (starting from the 2nd row)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    // Get the supplier's name (first column)
    const supplierName = row[0];

    // Loop through each column (starting from the 2nd column)
    for (let j = 1; j < row.length; j++) {

      // Check if the cell contains a date
      if (row[j] instanceof Date) {
        const renewalDate = row[j];

        // Calculate the number of days until renewal
        const daysUntilRenewal = calculateDaysDifference(renewalDate, currentDate);

        // Check if it's 45 days until renewal
        if (daysUntilRenewal === days_in_advance) {

          // Get the column header as the subject
          const columnHeader = headers[j];

          // Compose the email subject and message
          const subject = `Renewal Notification - ${supplierName}`;
          const message = `The ${columnHeader} for ${supplierName} will occur in 45 days from today (${renewalDate.toDateString()}).`;

          // Log email notification details
          console.log(`Email sent: Subject - ${subject}, Message - ${message}`);

          // Send an email notification, uncomment the line below  when ready to send the message
          //MailApp.sendEmail('recipient_example@email.com', subject, message);
        }
      }
    }
  }
}
// Log when the script completes
console.log("Script execution completed.");

// Function to calculate days between two dates
function calculateDaysDifference(finalDate, initialDate) {
  const millisecondsPerDay = 1000 * 60 * 60 * 24;
  const timeDifference = finalDate - initialDate; // it yields the time interval between these dates in milliseconds.
  return Math.floor(timeDifference / millisecondsPerDay); // it rounds it down to the nearest whole number since we are interested in the number of days, not the precise milliseconds.
}
