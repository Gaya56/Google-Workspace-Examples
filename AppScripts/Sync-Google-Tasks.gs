/**
 * @OnlyCurrentDoc
 * This script syncs a Google Sheet with a Google Tasks list,
 * supporting parent-child relationships.
 */

// --- MENU ---

/**
 * Creates a custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Google Tasks Sync')
    .addItem('Sync Tasks to Task List', 'syncTasks')
    .addToUi();
}

// --- CORE SYNC LOGIC ---

/**
 * The main function to perform a two-way sync between the sheet and Google Tasks.
 */
function syncTasks() {
  const ui = SpreadsheetApp.getUi();

  // 1. Get the target task list from the user
  const taskListId = getTaskListIdFromUser();
  if (!taskListId) {
    return; // User cancelled the prompt
  }

  // 2. Get data from both the Sheet and Google Tasks
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetHeaders = getSheetHeaders(sheet);
  const sheetData = getSheetData(sheet, sheetHeaders);
  const googleTasksMap = getGoogleTasks(taskListId);

  const writeData = []; // To store data for writing back to the sheet
  const newlyCreatedParentIds = new Map(); // Tracks new parent tasks created in this run

  // 3. Sync from Sheet to Google Tasks (Create & Update)
  for (const rowData of sheetData) {
    const taskId = rowData.data['Task ID'];

    if (taskId && googleTasksMap.has(taskId)) {
      // --- UPDATE EXISTING TASK ---
      const updatedTask = updateTaskInGoogle(taskListId, rowData, googleTasksMap.get(taskId), sheetHeaders);
      writeData.push(formatRowForSheet(updatedTask, sheetHeaders, rowData.parentTitle));
      googleTasksMap.delete(taskId); // Remove from map to track it's been processed

    } else {
      // --- CREATE NEW TASK ---
      const newTask = createTaskInGoogle(taskListId, rowData, sheetData, newlyCreatedParentIds);
      if (newTask) {
         writeData.push(formatRowForSheet(newTask, sheetHeaders, rowData.parentTitle));
         // If this new task is a parent, add its ID to our tracker for this session
         if (!rowData.parentTitle) {
            newlyCreatedParentIds.set(newTask.title, newTask.id);
         }
      } else {
        // If creation failed (e.g., parent not found or title empty), push original data back
        writeData.push(formatRowForSheet(rowData.data, sheetHeaders, rowData.parentTitle, true));
      }
    }
  }

  // 4. Sync from Google Tasks to Sheet (Import New)
  for (const task of googleTasksMap.values()) {
     // Any tasks left in the map are new in Google Tasks and need to be imported
     let parentTitle = '';
     if (task.parent) {
         // This is inefficient but necessary if the parent isn't in the sheet yet.
         try {
            const parentTask = Tasks.Tasks.get(taskListId, task.parent);
            parentTitle = parentTask.title;
         } catch(e) {
            console.warn(`Could not find parent task with ID ${task.parent} for task "${task.title}"`);
         }
     }
     writeData.push(formatRowForSheet(task, sheetHeaders, parentTitle));
  }


  // 5. Write all updated data back to the sheet in one go
  if (writeData.length > 0) {
    // Clear existing data to handle deletions correctly
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheetHeaders.length).clearContent();
    }
    sheet.getRange(2, 1, writeData.length, sheetHeaders.length).setValues(writeData);
  } else if (sheet.getLastRow() > 1) {
    // Handle case where all tasks are deleted
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheetHeaders.length).clearContent();
  }


  ui.alert('Sync Complete!', 'Your Google Sheet and Google Tasks are now in sync.', ui.ButtonSet.OK);
}


// --- HELPER FUNCTIONS ---

/**
 * Prompts the user to select a task list.
 * @return {string|null} The ID of the selected task list, or null if cancelled.
 */
function getTaskListIdFromUser() {
  const ui = SpreadsheetApp.getUi();
  const lists = Tasks.Tasklists.list();
  if (!lists.items || lists.items.length === 0) {
    ui.alert('No task lists found in your Google Account.');
    return null;
  }

  const listNames = lists.items.map(list => list.title);
  const response = ui.prompt(
    'Select a Task List',
    'Enter the name of the task list to sync:',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const selectedTitle = response.getResponseText().trim();
  const selectedList = lists.items.find(list => list.title === selectedTitle);

  if (selectedList) {
    return selectedList.id;
  } else {
    ui.alert(`Task list "${selectedTitle}" not found. Please try again.`);
    return null;
  }
}

/**
 * Retrieves all tasks from a given task list and returns them as a Map.
 * @param {string} taskListId The ID of the task list.
 * @return {Map<string, GoogleAppsScript.Tasks.Schema.Task>} A map of task IDs to task objects.
 */
function getGoogleTasks(taskListId) {
  const tasksMap = new Map();
  let pageToken;
  do {
    const result = Tasks.Tasks.list(taskListId, {
      showCompleted: true,
      showHidden: true,
      pageToken: pageToken
    });
    if (result.items) {
      for (const task of result.items) {
        tasksMap.set(task.id, task);
      }
    }
    pageToken = result.nextPageToken;
  } while (pageToken);
  return tasksMap;
}

/**
 * Gets the headers from the first row of the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to read from.
 * @return {string[]} An array of header strings.
 */
function getSheetHeaders(sheet) {
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}


/**
 * Reads the sheet data and organizes it into an array of objects.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {string[]} headers The array of column headers.
 * @return {Array<{row: number, parentTitle: string, data: Object}>} The structured sheet data.
 */
function getSheetData(sheet, headers) {
    if (sheet.getLastRow() < 2) {
      return []; // Return empty array if there's no data
    }
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length);
    const values = dataRange.getValues();
    const parentIndex = headers.indexOf('Parent');

    return values.map((row, index) => {
        const rowData = {};
        headers.forEach((header, i) => {
            rowData[header] = row[i];
        });
        return {
            row: index + 2, // 1-based index for sheet rows
            parentTitle: parentIndex !== -1 ? row[parentIndex] : '',
            data: rowData
        };
    });
}

/**
 * Creates a new task in Google Tasks.
 * @param {string} taskListId The task list ID.
 * @param {Object} rowData The row data object from the sheet.
 * @param {Array} allSheetData All data from the sheet for parent lookups.
 * @param {Map} newlyCreatedParentIds A map of newly created parent tasks in this run.
 * @return {GoogleAppsScript.Tasks.Schema.Task|null} The created task object or null.
 */
function createTaskInGoogle(taskListId, rowData, allSheetData, newlyCreatedParentIds) {
    // If the task title is empty or just whitespace, do not create the task.
    const taskTitle = rowData.data['Task Title'];
    if (!taskTitle || taskTitle.trim() === '') {
        console.log(`Skipping task creation because title is empty on row ${rowData.row}.`);
        return null; // Return null to indicate no task was created.
    }
  
    const taskDetails = {
        title: taskTitle,
        notes: rowData.data['Notes'],
        status: rowData.data['Completed?'] === true ? 'completed' : 'needsAction'
    };
    if (rowData.data['Deadline'] instanceof Date) {
        taskDetails.due = rowData.data['Deadline'].toISOString();
    }

    const options = {};
    if (rowData.parentTitle) {
        // Find the parent's ID
        let parentId = newlyCreatedParentIds.get(rowData.parentTitle);
        if (!parentId) {
            const parentRow = allSheetData.find(p => p.data['Task Title'] === rowData.parentTitle && p.data['Task ID']);
            if (parentRow) {
                parentId = parentRow.data['Task ID'];
            }
        }

        if (parentId) {
            options.parent = parentId;
        } else {
            console.warn(`Parent task "${rowData.parentTitle}" not found for sub-task "${taskDetails.title}". Creating as a top-level task.`);
        }
    }

    try {
        const newTask = Tasks.Tasks.insert(taskDetails, taskListId, options);
        console.log(`Created task: ${newTask.title} (ID: ${newTask.id})`);
        // Update the Task ID in the original data object for later writing
        rowData.data['Task ID'] = newTask.id;
        return newTask;
    } catch (e) {
        console.error(`Failed to create task "${taskDetails.title}": ${e.message}`);
        return null;
    }
}

/**
 * Updates an existing task in Google Tasks if changes are detected.
 * @param {string} taskListId The task list ID.
 * @param {Object} rowData The row data object from the sheet.
 * @param {GoogleAppsScript.Tasks.Schema.Task} googleTask The existing task object from the API.
 * @param {string[]} headers The array of column headers.
 * @return {GoogleAppsScript.Tasks.Schema.Task} The (potentially updated) task object.
 */
function updateTaskInGoogle(taskListId, rowData, googleTask, headers) {
    const payload = {};
    const sheetTitle = rowData.data['Task Title'] || '';
    const sheetNotes = rowData.data['Notes'] || '';
    const sheetStatus = (rowData.data['Completed?'] === true || rowData.data['Status'] === 'completed') ? 'completed' : 'needsAction';
    const sheetDeadline = rowData.data['Deadline'];

    if (sheetTitle && sheetTitle !== googleTask.title) payload.title = sheetTitle;
    if (sheetNotes !== (googleTask.notes || '')) payload.notes = sheetNotes;
    if (sheetStatus !== googleTask.status) payload.status = sheetStatus;

    // Date comparison
    const googleTaskDueDate = googleTask.due ? new Date(googleTask.due) : null;
    if (sheetDeadline instanceof Date && (!googleTaskDueDate || sheetDeadline.getTime() !== googleTaskDueDate.getTime())) {
        payload.due = sheetDeadline.toISOString();
    } else if (!sheetDeadline && googleTaskDueDate) {
        payload.due = null; // Clear the due date
    }

    if (Object.keys(payload).length > 0) {
        try {
            console.log(`Updating task: ${googleTask.title}`);
            return Tasks.Tasks.patch(payload, taskListId, googleTask.id);
        } catch (e) {
            console.error(`Failed to update task "${googleTask.title}": ${e.message}`);
        }
    }
    // If no changes, return the original task object from Google
    return googleTask;
}

/**
 * Formats a task object into a 2D array row for writing to the sheet.
 * @param {GoogleAppsScript.Tasks.Schema.Task | Object} task The task object or a plain data object.
 * @param {string[]} headers An array of the sheet headers in order.
 * @param {string} parentTitle The title of the parent task.
 * @param {boolean} isRawData If true, task is a plain object from the sheet, not a Google Task resource.
 * @return {Array} An array representing a row in the spreadsheet.
 */
function formatRowForSheet(task, headers, parentTitle = '', isRawData = false) {
    const taskStatus = isRawData ? task['Status'] : task.status;
    const isCompleted = isRawData ? task['Completed?'] : taskStatus === 'completed';

    const taskData = {
        'Parent': parentTitle,
        'Task Title': isRawData ? task['Task Title'] : task.title,
        'Notes': isRawData ? task['Notes'] : task.notes,
        'Deadline': isRawData ? task['Deadline'] : (task.due ? new Date(task.due) : null),
        'Task ID': isRawData ? task['Task ID'] : task.id,
        'Status': taskStatus,
        'Completed?': isCompleted
    };

    return headers.map(header => taskData[header] !== undefined ? taskData[header] : '');
}