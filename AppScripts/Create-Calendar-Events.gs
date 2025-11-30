/**
 * Google Apps Script for Automating Google Calendar Events from Google Sheets
 *
 * Description: This script automates the creation of Google Calendar events: "Sheet: `https://docs.google.com/spreadsheets/d/1Hi9eUDnuLpvsRqZK5vKDJe4LWpPPpplhqDDoix0c23c/edit?gid=0#gid=0`"
 * from data within a Google Sheet.  It processes rows, creates
 * events, and writes the event status and ID back to the sheet.
 *
 * Features:
 * -   Reads event data from a sheet named 'Calendar Events Data'.
 * -   Handles event titles, descriptions, start/end times, and attendees.
 * -   Skips rows where events have already been created (based on Event ID).
 * -   Writes event creation status and Event IDs back to the sheet.
 * -   Adds a custom menu to the Google Sheet for easy execution.
 *
 * Usage:
 * 1.  Ensure your Google Sheet is named 'Calendar Events Data' and has the
 * following headers in the first row: 'Event Title', 'Event Description',
 * 'Start Time', 'End Time', 'Attendees' (optional), 'Status', 'Event ID'.
 * 2.  Copy and paste this script into the Google Apps Script editor.
 * 3.  Run the 'createEventsFromSheet' function from the custom menu
 * 'Calendar Actions' in your Google Sheet.
 *
 * Author: Peter Horner
 * YouTube Channel: https://www.youtube.com/@PeterHornerGoogleTech
 *
 * License: Apache License, Version 2.0
 * (http://www.apache.org/licenses/LICENSE-2.0)
 *
 * Disclaimer:
 * This script is provided as-is, without any warranty.  Use at your own risk.
 */

const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Actions')
    .addItem('Create Calendar Events', 'createEventsFromSheet')
    .addToUi();
};

const createEventsFromSheet = () => {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Calendar Events Data');
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet named 'Calendar Events Data' not found.");
    return;
  }

  const dataRange = sheet.getDataRange(); // Process all rows in the sheet
  const [head, ...data] = dataRange.getValues();

  const getColumnIndex = (columnName) => {
    const index = head.indexOf(columnName);
    if (index === -1) {
      SpreadsheetApp.getUi().alert(`Column '${columnName}' not found in the header row.`);
      throw new Error(`Column '${columnName}' not found.`);
    }
    return index;
  };

  try {
    const titleCol = getColumnIndex('Event Title');
    const descriptionCol = getColumnIndex('Event Description');
    const startTimeCol = getColumnIndex('Start Time');
    const endTimeCol = getColumnIndex('End Time');
    const attendeesCol = head.indexOf('Attendees'); // Optional
    const statusCol = getColumnIndex('Status');
    const idCol = getColumnIndex('Event ID');

    const calendar = CalendarApp.getDefaultCalendar();
    const output = [];

    data.forEach((row, i) => {
      const title = row[titleCol];
      const description = row[descriptionCol];
      const startTime = row[startTimeCol];
      const endTime = row[endTimeCol];
      const attendees = attendeesCol !== -1 ? row[attendeesCol] : '';
      const existingEventId = row[idCol]; // Read the event ID from the sheet for this row

      let statusToSet = "";
      // Initialize idToSet with the existing ID.
      // This ensures if skipped, the existing ID is maintained.
      let idToSet = existingEventId;

      if (existingEventId) {
        statusToSet = "Skipped: Event already created.";
        // idToSet already holds existingEventId, so no change needed for the ID here.
        console.log(`Skipping row ${i + 2}: Event already created (ID: ${existingEventId})`);
      } else if (!title || !startTime || !endTime) {
        statusToSet = "Error: Missing title, start time, or end time.";
        idToSet = ""; // No ID if error and no previous ID
        console.error(`Error creating event in row ${i + 2}: Missing required fields.`);
      } else if (!(startTime instanceof Date) || !(endTime instanceof Date)) {
        statusToSet = "Error: Start and/or end time is not a valid date/time.";
        idToSet = ""; // No ID if error and no previous ID
        console.error(`Error creating event in row ${i + 2}: Invalid date/time format.`);
      } else {
        const eventOptions = {
          description: description || '',
          sendInvites: true,
        };
        if (attendees) {
          eventOptions.guests = attendees;
        }

        try {
          const event = calendar.createEvent(title, startTime, endTime, eventOptions);
          idToSet = event.getId(); // Successfully created, so use new ID
          statusToSet = "Created";
          console.log('Created event with id:', idToSet);
        } catch (error) {
          statusToSet = `Error: ${error.message}`; // Use error.message for a cleaner status
          idToSet = ""; // Error during creation, so no ID
          console.error(`Error creating event in row ${i + 2}:`, error);
        }
      }
      output.push([statusToSet, idToSet]);
    });

    // Write the status and event ID back to the sheet
    if (output.length > 0) { // Only write if there's output data
      const startRow = 2; // Start from the second row (after the header)
      // Prepare ranges for batch setting values for efficiency
      const statusValues = output.map(o => [o[0]]);
      const idValues = output.map(o => [o[1]]);

      sheet.getRange(startRow, statusCol + 1, output.length, 1).setValues(statusValues);
      sheet.getRange(startRow, idCol + 1, output.length, 1).setValues(idValues);
    }

    SpreadsheetApp.getUi().alert('Calendar events creation process completed.');

  } catch (error) {
    console.error("Error during script execution:", error);
    SpreadsheetApp.getUi().alert(`An error occurred: ${error.message}`);
  }
};