# Auto-Email-Notification
This repository contains Google Apps Script files utilized for sending automated notification emails from Google Sheets.
Google Apps Script is a powerful tool that allows you to automate tasks and integrate various Google Workspace applications. It uses modern JavaScript and offers built-in libraries for popular Google Workspace apps like Gmail, Calendar, Drive, and more. The best part? You don't need to install anything—it comes with a built-in code editor in your browser, and your scripts run on Google's servers. Learn more at [Google Apps Script Overview](https://developers.google.com/apps-script/overview).

In this guide, we will walk you through the process of setting up automatic email notifications in Google Sheets using Apps Script. Specifically, we will create a function that sends renewal email notifications to users when the renewal date is 30 days away from the current day. Follow these steps to get started:

### Step 1: Create a Google Sheet

Begin by creating a new Google Sheet or open an existing one where you want to manage contract expiration.

![Untitled](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/eca062e4-c4a5-401d-afc4-9d40ceef3386/Untitled.png)

### Step 2: Set Up Your Data

In your Google Sheet, create a table with columns for relevant contract information, such as Contract Name, Start Date, End Date, and Client Contact. Input your contract details into this table.

![Untitled](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/f4e4a346-cf78-4661-99bc-81a34f736305/Untitled.png)

### Step 3: Open Script Editor

From the Google Sheet, go to "Extensions" > "Apps Script" to open the Google Apps Script editor.

![Screenshot 2023-09-26 154125.png](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/43ea5e96-d090-41f6-986a-e19506762e18/Screenshot_2023-09-26_154125.png)

A new tab will pop out with a empty project as shown bellow

![Untitled](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/1709bd5d-8f3d-4601-ba74-cdda10d16d33/Untitled.png)

### Step 4: Write the Script

In the script editor, copy and paste the provided code into the `Code.gs` file. Be sure to rename the script project by clicking on "Untitled Project" in the upper left corner.

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

![2.png](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/a996fa61-e19f-4ae4-bc2b-253193c5590a/2.png)

If executed correctly, and provided that your sample data file contains expiration dates set 45 days ahead of the current date, your output should resemble the following:

![Untitled](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/aab2fe81-e2bc-4f54-8b73-dde3deb52cbf/Untitled.png)

### Step 6: Test the Script

Before relying on the script for important reminders, test it with sample data. With line 45 uncommented, run the script and check your email inbox for a notification with the subject line like "Renewal Notification - Supplier 3."

```jsx
 // Send an email notification
 MailApp.sendEmail('recipient_example@email.com', subject, message);
```

### Step 7: Set Triggers

To automate the script, set up triggers to run it at specific intervals. Go to the left panel and select "Triggers." Click the "Add Trigger" button and configure a time-driven event source trigger to run `sendNotification`function daily. 

![Screenshot 2023-09-27 105526.png](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/6f301bc5-ce70-46f2-bc42-8aba32941f3d/Screenshot_2023-09-27_105526.png)

![Screenshot 2023-09-27 105804.png](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/5a85c3a0-f822-45e1-8155-b9fdc17714c5/Screenshot_2023-09-27_105804.png)

![Untitled](https://prod-files-secure.s3.us-west-2.amazonaws.com/89abda67-c01d-42f4-9f5f-55e551c18fac/cdc200b1-cd0d-456b-aec8-11db53aa5e02/Untitled.png)

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
