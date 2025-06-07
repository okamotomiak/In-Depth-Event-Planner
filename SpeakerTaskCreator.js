//SpeakerTaskCreator.gs
/**
 * 
 * This file contains the speaker task creation functionality with an installable trigger.
 * This solution avoids the limitations of simple triggers and works with form submissions.
 */

/**
 * Create an installable trigger for the speaker task creation feature.
 * RUN THIS FUNCTION ONCE to set up the installable trigger.
 */
function createSpeakerTaskInstallableTrigger() {
  // First, delete any existing triggers with this name to avoid duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'checkForAcceptedSpeakers') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create a new trigger to run on edit
  ScriptApp.newTrigger('checkForAcceptedSpeakers')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  // Also create a trigger that runs every hour to catch form submissions
  ScriptApp.newTrigger('checkForAcceptedSpeakers')
    .timeBased()
    .everyHours(1)
    .create();
  
  // Display a message to confirm the trigger was created
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Speaker task creation trigger has been set up successfully!',
    'Setup Complete',
    5
  );
}

/**
 * Main function to check for accepted speakers and create tasks.
 * This runs on edit (via installable trigger) and periodically (every hour).
 * 
 * @param {Object} e The event object from the trigger (optional)
 */
function checkForAcceptedSpeakers(e) {
  try {
    Logger.log('checkForAcceptedSpeakers triggered');
    
    // Get the spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // For performance reasons, exit early if this is an edit event in a non-People sheet
    if (e && e.range && e.range.getSheet().getName() !== 'People') {
      Logger.log('Edit was not in People sheet, exiting');
      return;
    }
    
    // If it's an edit, check if it's relevant (status column in people sheet)
    if (e && e.range) {
      const editedSheet = e.range.getSheet();
      if (editedSheet.getName() !== 'People') return;
      
      // Get the edited row and column
      const row = e.range.getRow();
      if (row === 1) return; // Skip header row
      
      // Get the headers to find column indices
      const headers = editedSheet.getRange(1, 1, 1, editedSheet.getLastColumn()).getValues()[0];
      const statusColIndex = findColumnIndexByHeader(headers, 'Status');
      const categoryColIndex = findColumnIndexByHeader(headers, 'Category');
      const col = e.range.getColumn();
      
      Logger.log(`Edit event - row: ${row}, column: ${col}, status column: ${statusColIndex + 1}, category column: ${categoryColIndex + 1}`);
      
      // Check if we edited the Status column or Category column
      const editedStatus = (statusColIndex + 1 === col);
      const editedCategory = (categoryColIndex + 1 === col);
      
      if (!editedStatus && !editedCategory) {
        Logger.log('Edit was not in Status or Category column, exiting');
        return;
      }
      
      // Get the current values for status and category
      const statusValue = editedSheet.getRange(row, statusColIndex + 1).getValue();
      const categoryValue = editedSheet.getRange(row, categoryColIndex + 1).getValue();
      
      Logger.log(`Current values - Status: "${statusValue}", Category: "${categoryValue}"`);
      
      // Only proceed if Status is "Accepted" AND Category is "Speaker"
      if (statusValue !== 'Accepted' || categoryValue !== 'Speaker') {
        Logger.log('Person is not an Accepted Speaker, exiting');
        return;
      }
      
      // Get the person's name
      const nameColIndex = findColumnIndexByHeader(headers, 'Name');
      if (nameColIndex === -1) {
        Logger.log('Name column not found');
        return;
      }
      
      const speakerName = editedSheet.getRange(row, nameColIndex + 1).getValue();
      Logger.log(`Found accepted speaker: ${speakerName}`);
      
      // Create a task for this specific speaker
      createSpeakerTask(ss, speakerName);
      
      // Show a notification
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Task created for collecting bio and headshot from "${speakerName}"`,
        'Speaker Task Created',
        5
      );
      
      return; // We processed the specific edit, no need for full sheet scan
    }
    
    // If we reached here, either:
    // 1. This is a time-based trigger (no edit event), or
    // 2. Edit was in People sheet but not specifically in status/category columns
    
    // Process the entire People sheet
    processPeopleSheet(ss);
    
  } catch (error) {
    logError('Error in checkForAcceptedSpeakers', error);
  }
}

/**
 * Process the entire People sheet looking for accepted speakers without tasks
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 */
function processPeopleSheet(ss) {
  try {
    Logger.log('Processing entire People sheet');
    
    // Get the People sheet
    const peopleSheet = ss.getSheetByName('People');
    if (!peopleSheet) {
      logError('People sheet not found');
      return;
    }
    
    // Get all data from the People sheet
    const peopleData = peopleSheet.getDataRange().getValues();
    if (peopleData.length <= 1) {
      // Sheet is empty or only has headers
      Logger.log('People sheet is empty or only has headers');
      return;
    }
    
    // Get headers to find column indices
    const headers = peopleData[0];
    const nameColIndex = findColumnIndexByHeader(headers, 'Name');
    const categoryColIndex = findColumnIndexByHeader(headers, 'Category');
    const statusColIndex = findColumnIndexByHeader(headers, 'Status');
    
    Logger.log(`Column indices - Name: ${nameColIndex}, Category: ${categoryColIndex}, Status: ${statusColIndex}`);
    
    // Make sure required columns exist
    if (nameColIndex === -1 || categoryColIndex === -1 || statusColIndex === -1) {
      logError(`Required column not found. Name: ${nameColIndex}, Category: ${categoryColIndex}, Status: ${statusColIndex}`);
      return;
    }
    
    // Get existing tasks to avoid duplicates
    const existingTasks = getExistingSpeakerTasks(ss);
    Logger.log(`Found ${existingTasks.length} existing speaker tasks: ${existingTasks.join(', ')}`);
    
    // Process each row (starting from row 2)
    let tasksCreated = 0;
    for (let i = 1; i < peopleData.length; i++) {
      const row = peopleData[i];
      const name = row[nameColIndex];
      const category = row[categoryColIndex];
      const status = row[statusColIndex];
      
      // Skip empty rows
      if (!name) continue;
      
      // Check if this is an accepted speaker
      if (category === 'Speaker' && status === 'Accepted') {
        Logger.log(`Found accepted speaker: ${name}`);
        
        // Check if a task already exists for this speaker
        if (!existingTasks.includes(name)) {
          // Create a task for this speaker
          createSpeakerTask(ss, name);
          tasksCreated++;
          
          // Show a notification
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Task created for collecting bio and headshot from "${name}"`,
            'Speaker Task Created',
            5
          );
        } else {
          Logger.log(`Speaker ${name} already has a task, skipping`);
        }
      }
    }
    
    Logger.log(`Created ${tasksCreated} new speaker tasks`);
  } catch (error) {
    logError('Error in processPeopleSheet', error);
  }
}

/**
 * Get a list of speaker names that already have bio/headshot collection tasks
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @return {Array} Array of speaker names with existing tasks
 */
function getExistingSpeakerTasks(ss) {
  try {
    // Get the Task Management sheet
    const taskSheet = ss.getSheetByName('Task Management');
    if (!taskSheet) {
      Logger.log('Task Management sheet not found');
      return [];
    }
    
    // Get all task data
    const taskData = taskSheet.getDataRange().getValues();
    if (taskData.length <= 1) {
      Logger.log('Task Management sheet has no tasks');
      return []; // No tasks
    }
    
    // Find the task name column index
    const taskNameColIndex = findColumnIndexByHeader(taskData[0], 'Task Name');
    if (taskNameColIndex === -1) {
      Logger.log('Task Name column not found in Task Management sheet');
      return [];
    }
    
    // Find all tasks related to collecting bio/headshot
    const speakerNames = [];
    for (let i = 1; i < taskData.length; i++) {
      const taskName = taskData[i][taskNameColIndex];
      if (taskName && taskName.includes('Collect Bio & Headshot from')) {
        // Extract the speaker name
        const match = taskName.match(/Collect Bio & Headshot from (.+)/);
        if (match && match[1]) {
          speakerNames.push(match[1]);
        }
      }
    }
    
    return speakerNames;
  } catch (error) {
    logError('Error in getExistingSpeakerTasks', error);
    return [];
  }
}

/**
 * Create a task for collecting bio and headshot from a speaker
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @param {string} speakerName The name of the speaker
 */
function createSpeakerTask(ss, speakerName) {
  try {
    Logger.log(`Creating task for speaker: ${speakerName}`);
    
    // Get the event start date from Event Description sheet
    const eventDescSheet = ss.getSheetByName('Event Description');
    if (!eventDescSheet) {
      logError('Event Description sheet not found');
      return;
    }
    
    // Find the start date row
    const startDateRow = findRowByValue(eventDescSheet, 'Start Date (And Time)') || 
                        findRowByValue(eventDescSheet, 'Start Date');
    
    if (!startDateRow) {
      logError('Start Date row not found in Event Description sheet');
      return;
    }
    
    // Get the start date value
    const startDateValue = eventDescSheet.getRange(startDateRow, 2).getValue();
    Logger.log(`Event start date value: ${startDateValue}, type: ${typeof startDateValue}`);
    
    if (!startDateValue) {
      logError('Start Date value is empty');
      return;
    }
    
    let startDate;
    if (startDateValue instanceof Date) {
      startDate = new Date(startDateValue);
    } else {
      // Try to parse the date string
      startDate = new Date(startDateValue);
      if (isNaN(startDate.getTime())) {
        logError(`Invalid date: ${startDateValue}`);
        return;
      }
    }
    
    // Calculate due date (2 days before event start)
    const dueDate = new Date(startDate);
    dueDate.setDate(dueDate.getDate() - 2);
    Logger.log(`Due date calculated: ${dueDate}`);
    
    // Create the task in Task Management sheet
    const taskMgmtSheet = ss.getSheetByName('Task Management');
    if (!taskMgmtSheet) {
      logError('Task Management sheet not found');
      return;
    }
    
    // Find the last row with data
    const lastRow = taskMgmtSheet.getLastRow();
    const insertRow = lastRow + 1;
    
    // Generate task ID
    const taskId = generateUniqueTaskId();
    
    // Set task values
    const taskData = [
      taskId,  // Task ID
      `Collect Bio & Headshot from ${speakerName}`,  // Task Name
      `Request and collect bio and headshot from ${speakerName} for event profile`,  // Description
      'Staffing',  // Category
      '',  // Owner (leave blank as requested)
      dueDate,  // Due Date
      'Not Started',  // Status
      'Medium',  // Priority
      '',  // Related Session
      'No'  // Reminder Sent
    ];
    
    Logger.log(`Inserting task at row ${insertRow}: ${taskData.join(', ')}`);
    
    // Insert the task
    taskMgmtSheet.getRange(insertRow, 1, 1, taskData.length).setValues([taskData]);
    
    // Format the date cell
    taskMgmtSheet.getRange(insertRow, 6).setNumberFormat('yyyy-mm-dd');
    
    Logger.log(`Task created successfully for speaker: ${speakerName}`);
  } catch (error) {
    logError(`Error creating task for ${speakerName}`, error);
  }
}

/**
 * Helper function to find column index by header name (case-insensitive)
 * 
 * @param {Array} headers Array of header values
 * @param {string} headerName Header name to find (case-insensitive)
 * @return {number} 0-based index of the column, or -1 if not found
 */
function findColumnIndexByHeader(headers, headerName) {
  const lowerName = headerName.toLowerCase();
  return headers.findIndex(header => 
    String(header).toLowerCase().trim() === lowerName);
}

/**
 * Helper function to find a row by value in column A
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search
 * @param {string} value The value to find in column A
 * @return {number|null} The row number (1-based) or null if not found
 */
function findRowByValue(sheet, value) {
  const dataRange = sheet.getRange('A:A');
  const values = dataRange.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) {
      return i + 1; // Convert to 1-based row number
    }
  }
  
  return null;
}

/**
 * Generates a unique task ID
 * 
 * @return {string} A unique task ID
 */
function generateUniqueTaskId() {
  // Use timestamp plus random characters
  const timestamp = new Date().getTime();
  const random = Math.random().toString(36).substring(2, 6); // Get 4 random alphanumeric chars
  
  return 'T-' + timestamp.toString().slice(-6) + '-' + random.toUpperCase();
}

/**
 * Helper function to log errors consistently
 * 
 * @param {string} message Error message
 * @param {Error} error Error object (optional)
 */
function logError(message, error) {
  if (error) {
    Logger.log(`ERROR: ${message} - ${error.toString()}`);
    console.error(`ERROR: ${message} - ${error.toString()}`);
  } else {
    Logger.log(`ERROR: ${message}`);
    console.error(`ERROR: ${message}`);
  }
}