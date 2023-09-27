# Google Sheets & App Scripts: Automatic Email Notifications Setup Guide

Google Apps Script is a powerful tool that allows you to automate tasks and integrate various Google Workspace applications. It uses modern JavaScript and offers built-in libraries for popular Google Workspace apps like Gmail, Calendar, Drive, and more. The best part? You don't need to install anything—it comes with a built-in code editor in your browser, and your scripts run on Google's servers. Learn more at [Google Apps Script Overview](https://developers.google.com/apps-script/overview).

In this guide, we will walk you through the process of setting up automatic email notifications in Google Sheets using Apps Script. Specifically, we will create a function that sends renewal email notifications to users when the renewal date is 30 days away from the current day. Follow these steps to get started:

### Step 1: Create a Google Sheet

Begin by creating a new Google Sheet or open an existing one where you want to manage contract expiration.

![Create-Google-Sheet](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/1-Create-Google-Sheet.png)

### Step 2: Set Up Your Data

In your Google Sheet, create a table with columns for relevant contract information, such as Contract Name, Start Date, End Date, and Client Contact. Input your contract details into this table.

![Setup-Data](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/2-Setup-Data.png)

### Step 3: Open Script Editor

From the Google Sheet, go to "Extensions" > "Apps Script" to open the Google Apps Script editor.

![Open-Script-Editor](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/3-Open-Script-Editor.png)

A new tab will pop out with a empty project as shown bellow

![Empty Script](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/4-Empty-Project.png)

### Step 4: Write the Script

In the script editor, replace the existing function with the provided code below. Be sure to rename the script project by clicking on "Untitled Project" in the upper left corner.

```jsx
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
        if (daysUntilRenewal === 45) {

          // Get the column header as the subject
          const columnHeader = headers[j];

          // Compose the email subject and message
          const subject = `Renewal Notification - ${supplierName}`;
          const message = `The ${columnHeader} for ${supplierName} will occur in 45 days from today (${renewalDate.toDateString()}).`;

          // Log email notification details
          console.log(`Email sent: Subject - ${subject}, Message - ${message}`);

          // Send an email notification, uncomment the line below  when ready to send the message, and replace recipient_example@email.com with the email you wantg to send notification to.
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
```

### Step 5: Save and Run the Script

Save the code by clicking the archive icon at the top of the code editor. Then, click "Run" to execute the script. You may need to authorize the script to access your Google Sheets and send emails.

![Save Script](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/5-Save-Script.png)

If executed correctly, and provided that your sample data file contains expiration dates set 45 days ahead of the current date, your output should resemble the following:

![Executed Script](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/6-Executed-Script.png)

### Step 6: Test the Script

Before relying on the script for important reminders, test it with sample data. With line 45 uncommented, run the script and check your email inbox for a notification with the subject line like "Renewal Notification - Supplier 3."

```jsx
 // Send an email notification
 MailApp.sendEmail('recipient_example@email.com', subject, message);
```

### Step 7: Set Triggers

To automate the script, set up triggers to run it at specific intervals. Go to the left panel and select "Triggers." Click the "Add Trigger" button and configure a time-driven event source trigger to run `sendNotification`function daily. 

![Set Triggers](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/7-Set-Trigger.png)

![Add Trigger](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/8-Add-Trigger.png)

![Config Trigger](https://github.com/manfranklin/Auto-Email-Notification/blob/main/images/9-Config-Trigger.png)

### Step 8: Monitor and Refine

Keep an eye on the script's performance and refine it as needed. You can add more conditions or customize the email message to suit your specific requirements.

Automating tasks with Google Sheets and Apps Script can save you time and ensure important deadlines are met. Consider other scenarios where automated email notifications can be beneficia

- **Inventory Management:** Send notifications when stock levels are low or out of stock.
- **Project Management:** Send reminders for project deadlines and milestones.
- **Payment Reminders:** Automate reminders for overdue invoices.
- **Employee Contracts:** Alert managers when employee contracts are about to expire.
- **Performance Metrics:** Share weekly or monthly performance reports with stakeholders.
- **Customer Birthdays/Anniversaries:** Send personalized greetings to loyal customers.

 Enjoy experimenting and automating your tasks!

## **Sources**

"Google Apps Script: Overview." Google Developers, Google, developers.google.com/apps-script/overview.
