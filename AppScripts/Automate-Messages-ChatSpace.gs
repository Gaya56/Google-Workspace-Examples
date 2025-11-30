* Google Apps Script for Automating Google Chat Messages from Google Sheets
 *
 * Description: This script automates sending messages to Google Chat spaces
 * based on data within a Google Sheet. It processes rows, sends messages,
 * and updates the status back to the sheet.
 *
 * Features:
 * -   Reads message data from a sheet named 'Messages'.
 * -   Handles space names, message content, and scheduled dates.
 * -   Skips rows where messages have already been sent (based on Status).
 * -   Writes message sending status back to the sheet.
 * -   Adds a custom menu to the Google Sheet for easy execution.
 *
* https://developers.google.com/workspace/chat/quickstart/webhooks#apps-script_1

function onOpen() {
  // Get the UI service
  const ui = SpreadsheetApp.getUi();

  // Create a menu with two options
  ui.createMenu('Schedule Chat Messages')
      .addItem('Run Send Messages', 'sendMessages')
      .addItem('Setup', 'createTrigger')
      .addToUi();
}

function sendMessages() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Messages');
  const [head, ...data] = sheet.getDataRange().getValues();
  const statusColumn = head.indexOf('Status') + 1; 

  
  let j = 0
  data.forEach((row, i) => {
    const space = row[head.indexOf('Space')];
    const message = row[head.indexOf('Message')];
    const date = row[head.indexOf('Date')]; // Assuming this is a Date object
    const status = row[head.indexOf('Status')];
    const statusRange = sheet.getRange(i + 2, statusColumn);
    
    if (status != 'Message sent' && date < new Date()) { 
      j++
      SpreadsheetApp.getActiveSpreadsheet().toast(`Sending message ${j}`, 'Sending messages', 10);
      const spaceData = getSpace(space); // Get the space data
      if (spaceData.length > 0) { // Check if a space was found
        const spaceUrl = spaceData[0].url; // Extract the URL from the first element
        try {
          webhook(message, spaceUrl); 
        } catch (e) {
          console.log(e);
          statusRange.setValue('Error: ' + e); 
        } finally {
          statusRange.setValue('Message sent'); 
        }
      } else {
        // Handle the case where no space is found
        console.error("Space not found:", space);
        statusRange.setValue('Error: Space not found');
      }
    }
  });
}

function webhook(message,url) {
  const options = {
    "method": "post",
    "headers": {"Content-Type": "application/json; charset=UTF-8"},
    "payload": JSON.stringify({"text": message})
  };

  // Exponential backoff parameters
  const maxRetries = 5;
  let retryCount = 0;
  let initialDelay = 1000; // 1 second

  while (retryCount < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      console.log(response);
      break; // Success, exit the loop
    } catch (e) {
      console.error(`Attempt ${retryCount + 1} failed: ${e}`);
      retryCount++;
      // Exponential backoff with jitter
      let delay = initialDelay * Math.pow(2, retryCount) + Math.random() * 1000;
      Utilities.sleep(delay);
    }
  }

  if (retryCount === maxRetries) {
    console.error(`Max retries reached. Giving up.`);
    // You might want to handle this failure, e.g., send an email alert
  }
}

function getSpace(spaceName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Spaces');
  const [head, ...data] = sheet.getDataRange().getValues();
  let space = []
  data.forEach((row,i) => {
    const name = row[head.indexOf('Name')] // Changed name to rowName to avoid variable shadowing
    const url = row[head.indexOf('URL')]
    if (name == spaceName) { // Added condition to filter results
      space.push({
        name, // Use rowName here as well
        url
      })
    }
  })
  return space
}