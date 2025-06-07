//People.gs 

/**
 * Handles edits specifically for the People sheet
 * Called from the consolidated onEdit function in Core.gs
 * @param {Object} e The edit event object
 */
function handlePeopleSheetEdit(e) {
  try {
    // Get the edited sheet
    const sheet = e.range.getSheet();
    const spreadsheet = sheet.getParent();
    
    // Get the full row and column of the edit
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const newValue = e.range.getValue();
    
    // Skip if editing the header row
    if (row === 1) {
      Logger.log('Skipping header row edit');
      return;
    }
    
    // Get all headers to find the indexes of required columns
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Headers: ' + headers.join(', '));
    
    // Get the indices of the Status and Category columns
    const statusColIndex = findColumnIndex(headers, 'status');
    const categoryColIndex = findColumnIndex(headers, 'category');
    const nameColIndex = findColumnIndex(headers, 'name');
    
    Logger.log('Column indices - Status: ' + statusColIndex + 
                ', Category: ' + categoryColIndex + 
                ', Name: ' + nameColIndex);
    
    // Convert header index (0-based) to column number (1-based)
    const statusCol = statusColIndex + 1;
    const categoryCol = categoryColIndex + 1;
    const nameCol = nameColIndex + 1;
    
    // Skip if we couldn't find the required columns
    if (statusColIndex === -1 || categoryColIndex === -1 || nameColIndex === -1) {
      Logger.log('Required columns not found in People sheet');
      return;
    }
    
    // If the edit was in the status column and the value is "Accepted"
    if (col === statusCol && newValue === 'Accepted') {
      Logger.log('Status changed to Accepted');
      
      // Get the category for this row
      const categoryValue = sheet.getRange(row, categoryCol).getValue();
      Logger.log('Category is: ' + categoryValue);
      
      // Check if category is "Speaker"
      if (categoryValue === 'Speaker') {
        const speakerName = sheet.getRange(row, nameCol).getValue();
        Logger.log('Found accepted speaker: ' + speakerName);
        
        // Create a task for collecting bio and headshot
        createSpeakerTask(spreadsheet, speakerName);
        
        // Show a notification to the user
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Task for collecting bio and headshot from "${speakerName}" has been created.`,
          "Speaker Task Created",
          5  // Show for 5 seconds
        );
      }
    } 
    // If the edit was in the category column and the value is "Speaker"
    else if (col === categoryCol && newValue === 'Speaker') {
      Logger.log('Category changed to Speaker');
      
      // Get the status for this row
      const statusValue = sheet.getRange(row, statusCol).getValue();
      Logger.log('Status is: ' + statusValue);
      
      // Check if status is "Accepted"
      if (statusValue === 'Accepted') {
        const speakerName = sheet.getRange(row, nameCol).getValue();
        Logger.log('Found accepted speaker: ' + speakerName);
        
        // Create a task for collecting bio and headshot
        createSpeakerTask(spreadsheet, speakerName);
        
        // Show a notification to the user
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Task for collecting bio and headshot from "${speakerName}" has been created.`,
          "Speaker Task Created",
          5  // Show for 5 seconds
        );
      }
    }
  } catch (error) {
    Logger.log('Error in handlePeopleSheetEdit: ' + error.toString());
  }
}

/**
 * Helper function to find column index by case-insensitive header name
 * 
 * @param {Array} headers Array of header values
 * @param {string} columnName Column name to find (case-insensitive)
 * @return {number} 0-based index of the column, or -1 if not found
 */
function findColumnIndex(headers, columnName) {
  const lowerColumnName = columnName.toLowerCase();
  return headers.findIndex(header => 
    header.toString().toLowerCase().trim() === lowerColumnName);
}

/**
 * Creates a task for collecting bio and headshot from a speaker
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet
 * @param {string} speakerName The name of the speaker
 */
function createSpeakerTask(spreadsheet, speakerName) {
  try {
    // Get the Task Management sheet
    const taskMgmtSheet = spreadsheet.getSheetByName('Task Management');
    if (!taskMgmtSheet) {
      Logger.log('Task Management sheet not found');
      return;
    }
    
    // Get the Event Description sheet to find the event start date
    const eventDescSheet = spreadsheet.getSheetByName('Event Description');
    if (!eventDescSheet) {
      Logger.log('Event Description sheet not found');
      return;
    }
    
    // Try to find the Start Date row (accounting for different possible labels)
    const startDateRow = findRowByLabel(eventDescSheet, 'Start Date (And Time)') || 
                         findRowByLabel(eventDescSheet, 'Start Date');
    
    if (!startDateRow) {
      Logger.log('Start Date row not found in Event Description sheet');
      return;
    }
    
    // Get the start date
    const startDateValue = eventDescSheet.getRange(startDateRow, 2).getValue();
    if (!startDateValue || !(startDateValue instanceof Date)) {
      Logger.log('Invalid or missing start date');
      return;
    }
    
    // Calculate due date (2 days before event)
    const dueDate = new Date(startDateValue);
    dueDate.setDate(dueDate.getDate() - 2);
    
    // Get the last row in the Task Management sheet
    const lastRow = taskMgmtSheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Generate a unique task ID
    const taskId = generateTaskId();
    
    // Prepare the task data
    const taskData = [
      taskId,                                                    // Task ID
      `Collect Bio & Headshot from ${speakerName}`,              // Task Name
      `Request and collect bio and headshot from ${speakerName} for event profile`, // Description
      'Staffing',                                                // Category
      '',                                                        // Owner (left blank)
      dueDate,                                                   // Due Date
      'Not Started',                                             // Status
      'Medium',                                                  // Priority
      '',                                                        // Related Session
      'No'                                                       // Reminder Sent?
    ];
    
    // Insert the task row
    taskMgmtSheet.getRange(newRow, 1, 1, taskData.length).setValues([taskData]);
    
    // Format the due date
    taskMgmtSheet.getRange(newRow, 6).setNumberFormat('yyyy-mm-dd');
    
    // Apply alternating row colors
    if (newRow % 2 === 0) {
      taskMgmtSheet.getRange(newRow, 1, 1, taskData.length).setBackground('#f3f3f3');
    }
    
    Logger.log(`Created task for collecting bio from ${speakerName}`);
    
  } catch (error) {
    Logger.log('Error creating speaker task: ' + error.toString());
  }
}

/**
 * Helper function to find a row by label in column A
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search
 * @param {string} label The label to find in column A
 * @return {number|null} The row number (1-based) or null if not found
 */
function findRowByLabel(sheet, label) {
  const range = sheet.getRange('A:A');
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === label) {
      return i + 1; // Convert to 1-based row index
    }
  }
  
  return null;
}

/**
 * Generates a unique task ID
 * @return {string} A unique task ID
 */
function generateTaskId() {
  const timestamp = new Date().getTime();
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `T-${timestamp.toString().slice(-6)}-${random}`;
}

/**
 * Sets up the People sheet with headers, formatting, and sample data.
 * Updated to ensure dropdowns are applied only to data rows (starting from row 2).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object
 * @param {boolean} addSampleData Whether to add sample data to the sheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The configured People sheet
 */
function setupPeopleSheet(ss, addSampleData = true) {
  // Allow calling with or without spreadsheet parameter
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName('People');
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('People');
    sheet.setTabColor('#b45f06'); // Brown color
  } else {
    // Clear existing content if sheet already exists
    sheet.clear();
  }
  
  // Ensure we have at least 900 rows
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < 900) {
    sheet.insertRowsAfter(currentMaxRows, 900 - currentMaxRows);
  }
  
  // Define headers - REMOVED Google Form Responses and Form Submission Date columns
  const headers = ['Name', 'Category', 'Role/Position', 'Status', 'Email', 'Phone', 'Assigned Tasks'];
  
  // Set header values
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Set column widths
  const widths = [150, 120, 150, 120, 200, 120, 200];
  for (let i = 0; i < headers.length; i++) {
    if (i < widths.length) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Apply sample data if requested
  if (addSampleData) {
    const sampleData = [
      ['Jane Doe', 'Staff', 'Event Manager', 'Active', 'jane@example.com', '555-1234', ''],
      ['John Smith', 'Volunteer', 'Setup Crew', 'Active', 'john@example.com', '555-5678', '']
    ];
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  }
  
  // Apply dropdown validations to ALL data rows ONLY (rows 2-900, not header)
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Staff', 'Volunteer', 'Speaker', 'Participant'], true)
    .build();
  
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Pending', 'Confirmed', 'Accepted'], true)
    .build();
  
  // Apply to data rows ONLY (rows 2-900, not header row)
  sheet.getRange(2, 2, 899, 1).setDataValidation(categoryRule); // Category (column B)
  sheet.getRange(2, 4, 899, 1).setDataValidation(statusRule); // Status (column D)
  
  // Format headers with blue background and white text
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8') // Blue background
    .setFontColor('#ffffff') // White text
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set normal white background for the first data row and all other rows
  sheet.getRange(2, 1, 899, headers.length).setBackground(null);
  
  // Apply custom alternating row colors - manually setting alternating colors for data rows only
  for (let i = 2; i <= 900; i += 2) {
    // For even numbered rows, apply light gray background
    sheet.getRange(i, 1, 1, headers.length).setBackground('#f3f3f3');
  }
  
  // Create filter view for all rows
  sheet.getRange(1, 1, 900, headers.length).createFilter();
  
  return sheet;
}

/**
 * Sets dropdowns for Category, Status, and Assigned Tasks in People sheet.
 * Updated to pull task names from Task Management for Assigned Tasks dropdown.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} peopleSheet People sheet
 * @param {Number} numRows Number of rows to apply validation to
 * @param {Object} lists Configuration lists
 * @param {GoogleAppsScript.Spreadsheet.Sheet} taskMgmtSheet Task Management sheet for task names
 * @return {Array} List of updated dropdown fields
 */
function setPeopleDropdowns(peopleSheet, numRows, lists, taskMgmtSheet) {
  if (!peopleSheet) return [];
  
  // Always apply to data rows only (starting from row 2)
  const updated = [];
  
  // Ensure we have enough rows
  const currentMaxRows = peopleSheet.getMaxRows();
  if (currentMaxRows < numRows + 1) { // +1 for header
    peopleSheet.insertRowsAfter(currentMaxRows, (numRows + 1) - currentMaxRows);
  }
  
  // Apply Category dropdown from Config if available
  if (lists && lists['People Categories'] && lists['People Categories'].length) {
    const catRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lists['People Categories'], true)
      .build();
    // Apply to data rows only (rows 2-to-end)
    peopleSheet.getRange(2, 2, numRows).setDataValidation(catRule);
    updated.push("Category");
  }
  
  // Apply Status dropdown from Config if available
  if (lists && lists['People Statuses'] && lists['People Statuses'].length) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lists['People Statuses'], true)
      .build();
    // Apply to data rows only (rows 2-to-end)
    peopleSheet.getRange(2, 4, numRows).setDataValidation(statusRule);
    updated.push("Status");
  }
  
  // Add Assigned Tasks dropdown from Task Management sheet
  if (taskMgmtSheet) {
    try {
      // Get task names from Task Management sheet
      const taskLastRow = taskMgmtSheet.getLastRow();
      
      if (taskLastRow > 1) {
        // Get headers to find Task Name column
        const taskHeaders = taskMgmtSheet.getRange(1, 1, 1, taskMgmtSheet.getLastColumn()).getValues()[0];
        const taskNameColIndex = taskHeaders.findIndex(header => 
          header.toString().toLowerCase().trim() === 'task name');
        
        if (taskNameColIndex !== -1) {
          // Get all task names, skipping header row
          const taskNameRange = taskMgmtSheet.getRange(2, taskNameColIndex + 1, taskLastRow - 1, 1);
          const taskNameValues = taskNameRange.getValues();
          
          // Filter non-empty task names
          const allTaskNames = taskNameValues
            .filter(row => row[0])
            .map(row => row[0]);
          
          if (allTaskNames.length > 0) {
            // Find Assigned Tasks column in People sheet
            const peopleHeaders = peopleSheet.getRange(1, 1, 1, peopleSheet.getLastColumn()).getValues()[0];
            const assignedTasksColIndex = peopleHeaders.findIndex(header => 
              header.toString().toLowerCase().trim() === 'assigned tasks');
            
            if (assignedTasksColIndex !== -1) {
              // Create validation rule with all task names
              const taskRule = SpreadsheetApp.newDataValidation()
                .requireValueInList(allTaskNames, true)
                .build();
              
              // Apply to Assigned Tasks column - data rows only
              peopleSheet.getRange(2, assignedTasksColIndex + 1, numRows).setDataValidation(taskRule);
              updated.push("Assigned Tasks");
            }
          }
        }
      }
    } catch (error) {
      Logger.log('Error setting up Assigned Tasks dropdown: ' + error.toString());
    }
  }
  
  return updated;
}