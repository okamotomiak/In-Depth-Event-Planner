//Core.gs - Central functionality and utilities
/**
 * The ONLY onOpen function in the entire project.
 * Consolidates menu creation to prevent conflicts and improve performance.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create the Event Planner Setup menu with updated order
    ui.createMenu('Event Planner Setup')
      .addItem('Update Dashboard', 'setupDashboard')
      .addSeparator()
      .addItem('Generate AI Tasks', 'generateAITasksWithSchedule')
      .addItem('Generate Preliminary Schedule', 'generatePreliminarySchedule')
    .addSeparator()
    .addItem('Create/Reset Cue Builder Sheet', 'setupCueBuilderSheet') // <-- New
    .addItem('Generate Professional Cue Sheet', 'generateProfessionalCueSheet') // <-- Updated
    .addSeparator()
      .addItem('Generate Google Forms', 'showFormGeneratorDialog')
      .addItem('Update All Dropdowns', 'updateAllDropdowns')
          .addSubMenu(SpreadsheetApp.getUi().createMenu('Tutorial System')
        .addItem('Create Tutorial Overlays', 'createFullTutorialSystem')
        .addItem('Remove Tutorial Overlays', 'removeTutorialSystem'))
      .addToUi();

  // Show a welcome message if the spreadsheet appears blank
  if (isSpreadsheetBlank()) {
    ui.alert(
      'Welcome to Event Planner Pro',
      'This sheet looks empty. Explore the "Event Planner Setup" menu and start with the Quick Event Setup to begin planning.',
      ui.ButtonSet.OK
    );
  }

}

/**
 * The consolidated onEdit function for the entire project.
 * Handles all edit events in the spreadsheet.
 * @param {Object} e The edit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Get the row and column
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  Logger.log(`Edit detected in sheet: ${sheetName}, row: ${row}, column: ${col}`);
  
  // Handle Schedule sheet edits for:
  // 1. Duration calculation
  // 2. Session status changes to "Confirmed"
  if (sheetName === 'Schedule') {
    // Handle duration calculation for time changes
    handleScheduleEdit(e);
    
    // Handle session status changes to "Confirmed"
    handleSessionStatusChange(e);
  }
  
  // Handle People sheet edits for speaker task creation
  if (sheetName === 'People') {
    handlePeopleSheetEdit(e);
  }
  
  // Handle Task Management sheet edits
  if (sheetName === 'Task Management') {
    // Get column number of the edit
    const col = e.range.getColumn();
    
    // Find the Status column (usually column 7)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.findIndex(header => 
      header.toString().toLowerCase() === 'status');
    
    // If Status column was edited, update dashboard metrics
    if (statusColIndex !== -1 && col === statusColIndex + 1) {
      Logger.log('Task status changed, updating dashboard metrics...');
      updateDashboardTaskMetrics();
    }
  }
  
  // Handle Dashboard refresh button click
  if (sheetName === 'Dashboard') {
    handleDashboardEdit(e);
  }
}

/**
 * Finds the row number where a field label appears in column A.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {string} label The label to find in column A.
 * @return {number|null} The 1-based row number, or null if not found.
 */
function _findRow(sheet, label) {
  const vals = sheet.getRange('A:A').getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === label) return i + 1;
  }
  return null;
}

/**
 * Returns today's date in yyyy-mm-dd format
 * @return {string} Formatted date string
 */
function _getTodayString() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Utility function to get sheet by name, creating it if it doesn't exist.
 * Consolidates repeated sheet access pattern for performance.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @param {string} sheetName The name of the sheet
 * @param {string} tabColor Optional tab color
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The sheet
 */
function _getOrCreateSheet(ss, sheetName, tabColor) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (tabColor) sheet.setTabColor(tabColor);
  }
  
  return sheet;
}

/**
 * Utility function to apply formats to a range in batch
 * instead of individual method calls.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to format
 * @param {Object} formats The formatting to apply
 */
function _batchFormatRange(range, formats) {
  // Apply all formats in one go - reduces API calls
  if (formats.background) range.setBackground(formats.background);
  if (formats.fontColor) range.setFontColor(formats.fontColor);
  if (formats.fontWeight) range.setFontWeight(formats.fontWeight);
  if (formats.horizontalAlignment) range.setHorizontalAlignment(formats.horizontalAlignment);
  if (formats.verticalAlignment) range.setVerticalAlignment(formats.verticalAlignment);
  if (formats.fontSize) range.setFontSize(formats.fontSize);
  if (formats.numberFormat) range.setNumberFormat(formats.numberFormat);
}

/**
 * Reads the Config sheet and returns an object of lists.
 * Optimized to read in a single operation.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @return {Object} Map of list names to array of options
 */
function _getConfigLists(ss) {
  const config = ss.getSheetByName('Config');
  if (!config) return {};
  
  const lastRow = config.getLastRow();
  if (lastRow < 2) return {};
  
  // Read all data at once
  const rows = config.getRange(2, 1, lastRow - 1, 2).getValues();
  const lists = {};
  
  // Process all lists in one pass
  rows.forEach(r => {
    const key = r[0] ? r[0].toString().trim() : null;
    const val = r[1] ? r[1].toString() : '';
    if (key) lists[key] = val.length ? val.split(',').map(s => s.trim()) : [];
  });
  
  return lists;
}

/**
 * Gets a sheet by name, logs error if not found.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string} sheetName The name of the sheet to get.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet object or null if not found.
 */
function _getSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Error: Sheet "${sheetName}" not found.`);
  }
  return sheet;
}

/**
 * Reads all data from a sheet, returns headers and data rows separately.
 * Headers are converted to lowercase. Logs error if sheet is empty.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @return {{headers: string[], dataRows: Array<Array<any>>}|null} Object with headers/data or null on error.
 */
function _getSheetData(sheet) {
  if (!sheet) return null; // Should be caught by _getSheet already

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    Logger.log(`Error: Sheet "${sheet.getName()}" is empty.`);
    return null;
  }
  const headers = data[0].map(h => h.toString().toLowerCase()); // Lowercase headers
  const dataRows = data.slice(1); // Exclude header row
   if (dataRows.length === 0) {
    Logger.log(`Warning: Sheet "${sheet.getName()}" has headers but no data rows.`);
    // Return headers but empty dataRows
  }
  return { headers, dataRows };
}

/**
 * Finds the 0-based indices of required columns from a list of headers. Case-insensitive.
 * Logs an error if any required column is not found.
 * @param {string[]} headers Array of header strings (expected lowercase).
 * @param {string[]} requiredCols Array of required column names (case doesn't matter).
 * @return {Object|null} A map of { 'lowercase_col_name': index } or null if any column is missing.
 */
function _getColumnsIndices(headers, requiredCols) {
  if (!headers || headers.length === 0) {
      Logger.log('Error: Cannot find columns in empty headers.');
      return null;
  }
  const indices = {};
  let allFound = true;

  requiredCols.forEach(colName => {
    const lowerColName = colName.toLowerCase();
    const index = headers.indexOf(lowerColName);
    if (index === -1) {
      Logger.log(`Error: Required column "${colName}" not found in headers: [${headers.join(', ')}]`);
      allFound = false;
    } else {
      indices[lowerColName] = index;
    }
  });

  return allFound ? indices : null;
}

/**
 * Handles a session status change to "Confirmed" in the Schedule sheet.
 * @param {Object} e The edit event object
 */
function handleSessionStatusChange(e) {
  try {
    const sheet = e.range.getSheet();
    
    // Only proceed if edit is in column 8 (Status column) 
    if (e.range.getColumn() !== 8) return;
    
    // Skip header row
    if (e.range.getRow() === 1) return;
    
    // Check if the new value is "Confirmed"
    const newValue = e.range.getValue();
    if (newValue !== 'Confirmed') return;
    
    // Get information about the confirmed session
    const row = e.range.getRow();
    const sessionData = sheet.getRange(row, 1, 1, 7).getValues()[0];
    
    // Extract session details
    const sessionDate = sessionData[0];
    const sessionTime = sessionData[1];
    const sessionTitle = sessionData[4];
    const sessionLead = sessionData[5];
    const sessionLocation = sessionData[6];
    
    // Format the date and time for display
    let dateStr = '';
    if (sessionDate instanceof Date) {
      const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
      dateStr = sessionDate.toLocaleDateString(undefined, options);
    }
    
    // Show a confirmation message
    const message = `Session "${sessionTitle}" is now confirmed:\n` + 
                    `Date: ${dateStr}\n` + 
                    `Time: ${sessionTime}\n` + 
                    `Location: ${sessionLocation}\n` + 
                    `Lead: ${sessionLead || 'Not assigned'}`;
    
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Session Confirmed', 5);
    
    // Log the confirmation
    Logger.log(`Session confirmed - Title: ${sessionTitle}, Date: ${dateStr}, Time: ${sessionTime}, Location: ${sessionLocation}, Lead: ${sessionLead}`);
    
  } catch (error) {
    Logger.log(`Error in handleSessionStatusChange: ${error}`);
  }
}

/**
 * Handles edit events in the Dashboard sheet
 * Specifically looks for clicks on the refresh button
 * @param {Object} e The edit event object
 */
function handleDashboardEdit(e) {
  // Only proceed if edit is in Dashboard sheet
  if (!e || !e.range || e.range.getSheet().getName() !== 'Dashboard') return;
  
  // Check if the edit is in the refresh button cell (H1)
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  if (row === 1 && col === 8) {
    // User clicked the refresh button
    setupDashboard();
    
    // Show a confirmation toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Dashboard refreshed with latest data', 
      'Refresh Complete', 
      3);
  }
}

/**
 * Updates dashboard task metrics when task status changes
 */
function updateDashboardTaskMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('Dashboard');
    
    if (!dashboardSheet) return;
    
    // Only refresh if the Dashboard sheet exists
    setupDashboard();
    
  } catch (error) {
    Logger.log(`Error updating dashboard metrics: ${error}`);
  }
}

/**
 * Removes all dashboard triggers for auto-update
 */
function removeAllDashboardTriggers() {
const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    const trigger = allTriggers[i];
    if (trigger.getHandlerFunction() === 'handleDashboardEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Creates a session reminder for upcoming events
 * Can be run as a scheduled job (e.g., daily)
 */
function createSessionReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName('Schedule');
  
  if (!scheduleSheet) return;
  
  // Get all sessions
  const lastRow = scheduleSheet.getLastRow();
  if (lastRow <= 1) return; // Only header row
  
  const scheduleData = scheduleSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  
  // Get today's date
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set to midnight for date comparison
  
  // Get configured look-ahead days from Config sheet (default to 2)
  const lookAheadDays = getLookAheadDays() || 2;
  
  // Calculate the future date to check for
  const futureDate = new Date(today);
  futureDate.setDate(today.getDate() + lookAheadDays);
  
  // Filter for sessions happening on the future date
  const upcomingSessions = scheduleData.filter(row => {
    if (!row[0] || !(row[0] instanceof Date)) return false;
    
    const sessionDate = new Date(row[0]);
    sessionDate.setHours(0, 0, 0, 0); // Set to midnight for comparison
    
    return sessionDate.getTime() === futureDate.getTime();
  });
  
  // If there are upcoming sessions, show a notification
  if (upcomingSessions.length > 0) {
    let message = `Reminder: ${upcomingSessions.length} sessions scheduled for ${formatDate(futureDate)}:\n\n`;
    
    upcomingSessions.forEach((session, index) => {
      if (index < 5) { // Limit to showing 5 sessions in the notification
        message += `• ${session[4]} at ${session[1]}\n`;
      } else if (index === 5) {
        message += `• ... and ${upcomingSessions.length - 5} more\n`;
      }
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      `Upcoming Sessions - ${formatDate(futureDate)}`,
      10  // Show for 10 seconds
    );
  }
}

/**
 * Gets the look-ahead days from Config sheet
 * @return {number} Number of days to look ahead (default 1)
 */
function getLookAheadDays() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    
    if (!configSheet) return 1;
    
    // Find the "Look-Ahead Days" row
    const lookAheadRow = findRowByValue(configSheet, 'Look-Ahead Days');
    if (!lookAheadRow) return 1;
    
    // Get the value
    const lookAheadValue = configSheet.getRange(lookAheadRow, 2).getValue();
    if (!lookAheadValue) return 1;
    
    // Convert to number, default to 1 if not a number
    const days = parseInt(lookAheadValue);
    return isNaN(days) ? 1 : days;
    
  } catch (error) {
    Logger.log(`Error getting look-ahead days: ${error}`);
    return 1; // Default to 1 day
  }
}

/**
 * Helper function to find a row by value in column A
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet
 * @param {string} value The value to find
 * @return {number|null} Row number or null if not found
 */
function findRowByValue(sheet, value) {
  const range = sheet.getRange('A:A');
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) {
      return i + 1; // Convert to 1-based row number
    }
  }
  
  return null;
}

/**
 * Format a date as a user-friendly string
 * @param {Date} date The date to format
 * @return {string} Formatted date string
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) return '';
  
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString(undefined, options);
}

/**
 * Trigger function to run automatically daily
 * Set this up using "Project triggers" in the Apps Script editor
 */
function dailyTrigger() {
  // Run session reminders
  createSessionReminders();
  
  // Check for upcoming tasks
  createTaskReminders();
}

/**
 * Creates reminders for upcoming tasks
 */
function createTaskReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName('Task Management');
  
  if (!taskSheet) return;
  
  // Get all tasks
  const lastRow = taskSheet.getLastRow();
  if (lastRow <= 1) return; // Only header row
  
  // Get the task data
  const taskData = taskSheet.getDataRange().getValues();
  const headers = taskData[0];
  
  // Find column indices
  const dueDateColIndex = headers.findIndex(header => 
    header.toString().toLowerCase().trim() === 'due date');
  const statusColIndex = headers.findIndex(header => 
    header.toString().toLowerCase().trim() === 'status');
  const taskNameColIndex = headers.findIndex(header => 
    header.toString().toLowerCase().trim() === 'task name');
  const reminderColIndex = headers.findIndex(header => 
    header.toString().toLowerCase().trim() === 'reminder sent?');
  
  if (dueDateColIndex === -1 || statusColIndex === -1 || 
      taskNameColIndex === -1 || reminderColIndex === -1) {
    Logger.log('Required columns not found in Task Management sheet');
    return;
  }
  
  // Get reminder lead time from Config (default to 2 days)
  const reminderLeadTime = getReminderLeadTime() || 2;
  
  // Get today's date
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set to midnight for date comparison
  
  // Calculate the future date to check for
  const reminderDate = new Date(today);
  reminderDate.setDate(today.getDate() + reminderLeadTime);
  
  // Find tasks due on the reminder date that haven't had reminders sent yet
  const tasksDue = [];
  const tasksToUpdate = [];
  
  for (let i = 1; i < taskData.length; i++) {
    const row = taskData[i];
    const dueDate = row[dueDateColIndex];
    const status = row[statusColIndex];
    const taskName = row[taskNameColIndex];
    const reminderSent = row[reminderColIndex];
    
    // Skip tasks that are already done, cancelled, or have had reminders sent
    if (status === 'Done' || status === 'Cancelled' || reminderSent === 'Yes') continue;
    
    // Check if the task is due on the reminder date
    if (dueDate instanceof Date) {
      const taskDueDate = new Date(dueDate);
      taskDueDate.setHours(0, 0, 0, 0); // Set to midnight for comparison
      
      if (taskDueDate.getTime() === reminderDate.getTime()) {
        tasksDue.push(taskName);
        tasksToUpdate.push(i + 1); // Convert to 1-based row index
      }
    }
  }
  
  // If there are tasks due, show a notification and update the Reminder Sent column
  if (tasksDue.length > 0) {
    let message = `Tasks due in ${reminderLeadTime} days (${formatDate(reminderDate)}):\n\n`;
    
    tasksDue.forEach((task, index) => {
      if (index < 5) { // Limit to showing 5 tasks in the notification
        message += `• ${task}\n`;
      } else if (index === 5) {
        message += `• ... and ${tasksDue.length - 5} more\n`;
      }
    });
    
    // Show notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      `Task Reminders`,
      10  // Show for 10 seconds
    );
    
    // Update the Reminder Sent column for all tasks
    tasksToUpdate.forEach(rowIndex => {
      taskSheet.getRange(rowIndex, reminderColIndex + 1).setValue('Yes');
    });
  }
}

/**
 * Gets the reminder lead time from Config sheet
 * @return {number} Number of days lead time for reminders (default 2)
 */
function getReminderLeadTime() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    
    if (!configSheet) return 2;
    
    // Find the "Reminder Lead Time (days)" row
    const reminderRow = findRowByValue(configSheet, 'Reminder Lead Time (days)');
    if (!reminderRow) return 2;
    
    // Get the value
    const reminderValue = configSheet.getRange(reminderRow, 2).getValue();
    if (!reminderValue) return 2;
    
    // Convert to number, default to 2 if not a number
    const days = parseInt(reminderValue);
    return isNaN(days) ? 2 : days;
    
  } catch (error) {
    Logger.log(`Error getting reminder lead time: ${error}`);
    return 2; // Default to 2 days
  }
}

/**
 * Determines if the spreadsheet is completely blank
 * @return {boolean} True if blank, false otherwise
 */
function isSpreadsheetBlank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  if (sheets.length !== 1) return false;

  const sheet = sheets[0];
  const dataRange = sheet.getDataRange();
  const hasOneCell =
    dataRange.getNumRows() === 1 && dataRange.getNumColumns() === 1;
  const isEmpty = dataRange.getValue() === '';

  return hasOneCell && isEmpty;
}