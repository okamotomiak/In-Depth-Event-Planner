//Config.gs 
/**
 * Creates and populates the Config sheet with default configuration values.
 * This function can be run once to initialize the Config sheet or to reset it to defaults.
 */
function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  
  // Create the sheet if it doesn't exist
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.setTabColor('#6aa84f'); // Green color for configuration
  } else {
    // Clear existing content if sheet already exists
    configSheet.clear();
  }
  
  // Set up column headers
  const headers = ["Key / Template Name", "Value / Subject", "Body"];
  configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  configSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set column widths
  configSheet.setColumnWidth(1, 200); // Key / Template Name
  configSheet.setColumnWidth(2, 250); // Value / Subject
  configSheet.setColumnWidth(3, 450); // Body
  
  // Prepare the configuration data
  // Reorganized to put form entries right after OpenAI API Key
  const configData = [
    ["Event Type", "Single,Multi", ""],
    ["Status Options", "Tentative,Confirmed,Cancelled", ""],
    ["People Categories", "Staff,Volunteer,Speaker,Participant", ""],
    ["People Statuses", "Potential,Invited,Accepted,Registered,Unavailable", ""],
    ["Budget Cost Basis", "Flat,Per Person", ""],
    ["Location List", "Main Hall,Room 101,Room 102,Outdoor Area", ""],
    ["Task Status Options", "Not Started,In Progress,Blocked,Done,Cancelled", ""],
    ["Task Priority Options", "High,Medium,Low", ""],
    ["Owners", "John Doe,Jane Smith,Alex Johnson,Maria Garcia,Sam Wilson,Eduardo,Mika,Junpei", ""],
    ["Look-Ahead Days", "1", ""],
    ["Reminder Lead Time (days)", "2", ""],
    ["New-Event Template ID", "", ""],
    ["Default Food Rate ($/person)", "10", ""],
    ["OpenAI API Key", "", "[I will enter the key manually]"],
    // Form entries right after OpenAI API Key
    ["registration form", "", ""],
    ["volunteer sign up", "", ""],
    ["speaker form", "", ""],
    // Email templates with no special formatting
    ["InviteTemplate", "Invitation: {{name}} for [EVENT NAME]", "Hi {{name}},\n\nYou are invited to [EVENT NAME]!\n\n[Add event details like date, time, location.]\n\nPlease RSVP by [RSVP Date].\n\nMore info here: [Link]\n\nBest regards,\n[Your Name/Org]"],
    ["ReminderTemplate", "Reminder: [EVENT NAME] is coming up!", "Hi {{name}},\n\nJust a friendly reminder about the upcoming event: [EVENT NAME] on [Date] at [Time].\n\nLocation: [Location]\n\nWe look forward to seeing you!\n\nBest regards,\n[Your Name/Org]"],
    ["ThankYouTemplate", "Thank You for Attending [EVENT NAME]!", "Hi {{name}},\n\nThank you for attending [EVENT NAME]!\n\nWe hope you enjoyed it. [Optional: Add link to slides, photos, feedback survey, etc.]\n\nBest regards,\n[Your Name/Org]"]
  ];
  
  // Insert the configuration data
  configSheet.getRange(2, 1, configData.length, configData[0].length).setValues(configData);
  
  // Format the rows
  // Set alternating row colors for better readability
  const dataRows = configSheet.getRange(2, 1, configData.length, 3);
  dataRows.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  
  // Highlight specific rows
  // OpenAI API Key row (row 14)
  const apiKeyRow = configSheet.getRange(15, 1, 1, 3);
  apiKeyRow.setBackground('#ffe6cc'); // Light orange
  
  // Form-related entries (rows 15-17)
  const formRows = configSheet.getRange(16, 1, 3, 3);
  formRows.setBackground('#cfe2f3'); // Light blue
  
  // REMOVE background color from email template rows (rows 18-20)
  const emailTemplateRows = configSheet.getRange(19, 1, 3, 3);
  emailTemplateRows.setBackground(null); // No background color
  
  // Format line breaks in the Body column
  const bodyColumn = configSheet.getRange(2, 3, configData.length, 1);
  bodyColumn.setWrap(true);
  
  // Freeze the header row
  configSheet.setFrozenRows(1);
  
  // Add note to OpenAI API Key cell
  configSheet.getRange(15, 2).setNote('Enter your OpenAI API key here. It will be used for AI functionality.');
  
  // Add notes to form URL cells
  const formNoteText = 'This cell will store the Google Form URL after it is created. Do not modify manually.';
  configSheet.getRange(16, 2, 3, 1).setNote(formNoteText);
  
  // Format text alignment
  configSheet.getRange(2, 1, configData.length, 1).setHorizontalAlignment('left'); // Keys left-aligned
  configSheet.getRange(2, 2, configData.length, 1).setHorizontalAlignment('left'); // Values left-aligned
  configSheet.getRange(2, 3, configData.length, 1).setHorizontalAlignment('left'); // Body left-aligned
  
  // Format units for numerical values
  configSheet.getRange(14, 2).setNumberFormat('0.00'); // Default Food Rate
  
  // Alert the user
  SpreadsheetApp.getUi().alert('Config sheet has been set up with default values.');
  
  return configSheet;
}

/**
 * Updates all dropdowns across all sheets in a single operation.
 * This function centralizes dropdown management for all sheets.
 */
function updateAllDropdowns() {
  // Get spreadsheet once and reuse
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lists = _getConfigLists(ss);
  
  // Cache sheets to avoid repeated getSheetByName calls
  const sheets = {
    eventDesc: ss.getSheetByName('Event Description'),
    schedule: ss.getSheetByName('Schedule'),
    people: ss.getSheetByName('People'),
    taskMgmt: ss.getSheetByName('Task Management')
  };
  
  // Cache actual data extents rather than using fixed sizes
  const rowCounts = {};
  Object.keys(sheets).forEach(key => {
    if (sheets[key]) {
      // Use max of lastRow or at least 100 rows to ensure dropdown coverage
      rowCounts[key] = Math.max(100, sheets[key].getLastRow() - 1);
    }
  });
  
  let updatedElements = [];
  
  // Update Event Description sheet
  if (setEventTypeDropdown(sheets.eventDesc, lists)) {
    updatedElements.push("Event Type");
  }
  
  // Update Schedule sheet dropdowns
  const scheduleUpdates = setScheduleDropdowns(ss, sheets, rowCounts, lists);
  if (scheduleUpdates.length > 0) {
    updatedElements.push("Schedule (" + scheduleUpdates.join(", ") + ")");
  }
  
  // Update People sheet dropdowns - now passing Task Management sheet for Assigned Tasks dropdown
  const peopleUpdates = setPeopleDropdowns(sheets.people, rowCounts.people, lists, sheets.taskMgmt);
  if (peopleUpdates.length > 0) {
    updatedElements.push("People (" + peopleUpdates.join(", ") + ")");
  }
  
  // Update Task Management sheet dropdowns including Owner from People sheet
  const taskUpdates = updateTaskManagementDropdowns(ss, sheets, rowCounts);
  if (taskUpdates.length > 0) {
    updatedElements.push("Task Management (" + taskUpdates.join(", ") + ")");
  }
  
  // Provide feedback on what was updated
  const message = updatedElements.length > 0 
    ? "Updated dropdowns for: " + updatedElements.join(", ")
    : "No dropdowns needed updating";
    
  SpreadsheetApp.getUi().alert('All Dropdowns Updated', message, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(message);
}

/**
 * Updates all dropdowns in the Task Management sheet including Owner from People sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @param {Object} sheets Cached sheet references
 * @param {Object} rowCounts Cached row counts
 * @return {Array} List of updated dropdown fields
 */
function updateTaskManagementDropdowns(ss, sheets, rowCounts) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!sheets) {
    sheets = {
      taskMgmt: ss.getSheetByName('Task Management'),
      people: ss.getSheetByName('People'),
      schedule: ss.getSheetByName('Schedule')
    };
  }
  
  const taskSheet = sheets.taskMgmt;
  if (!taskSheet) return [];
  
  // Calculate how many rows to apply validations to
  // Use the actual last row of data, or a minimum of 100 rows for future data
  const taskLastRow = taskSheet.getLastRow();
  const numRows = Math.max(100, taskLastRow > 1 ? taskLastRow - 1 : 10);
  
  const updated = [];

  // Get configuration lists
  const lists = _getConfigLists(ss);
  
  // Define dropdown options from Config sheet when possible
  const statusOptions = lists['Task Status Options'] || ['Not Started', 'In Progress', 'Done', 'Overdue', 'Blocked', 'Cancelled'];
  const priorityOptions = lists['Task Priority Options'] || ['Low', 'Medium', 'High', 'Critical'];
  const reminderOptions = ['Yes', 'No'];
  const categoryOptions = ['Venue', 'Marketing', 'Logistics', 'Program', 'Budget', 'Staffing', 'Technology', 'Communications', 'Other'];
  const ownerOptions = lists['Owners'] || [];

  // Create rules for batch application
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build();
  
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(priorityOptions, true)
    .build();
  
  const reminderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(reminderOptions, true)
    .build();
    
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryOptions, true)
    .build();

  // Apply rules in batch to data rows only (starting from row 2)
  taskSheet.getRange(2, 4, numRows).setDataValidation(categoryRule); // Category (Column 4)
  taskSheet.getRange(2, 7, numRows).setDataValidation(statusRule); // Status (Column 7)
  taskSheet.getRange(2, 8, numRows).setDataValidation(priorityRule); // Priority (Column 8)
  taskSheet.getRange(2, 10, numRows).setDataValidation(reminderRule); // Reminder Sent? (Column 10)
  
  updated.push("Category", "Status", "Priority", "Reminder");

  // Use Owners from Config sheet if available
  if (ownerOptions.length > 0) {
    Logger.log("Using Owners from Config sheet: " + ownerOptions.join(", "));
    
    const ownerRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(ownerOptions, true)
      .build();
    
    // Owner is column 5 in Task Management sheet
    taskSheet.getRange(2, 5, numRows).setDataValidation(ownerRule);
    updated.push("Owner");
  }
  // If no Owners in Config, fall back to People sheet names
  else {
    // FALLBACK: Fetch ALL people from People sheet for Owner dropdown
    const peopleSheet = sheets.people;
    if (peopleSheet) {
      const lastPeopleRow = peopleSheet.getLastRow();
      if (lastPeopleRow > 1) {
        // Get all names from row 2 to last row - properly excludes header
        const dataRowCount = lastPeopleRow - 1; // Exclude header row
        const nameRange = peopleSheet.getRange(2, 1, dataRowCount, 1);
        const nameValues = nameRange.getValues();
        
        // Filter non-empty names
        const allPeople = nameValues
          .filter(row => row[0])
          .map(row => row[0]);
        
        if (allPeople.length > 0) {
          Logger.log("Fallback to People sheet for Owner dropdown: " + allPeople.join(", "));
          
          const ownerRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(allPeople, true)
            .build();
          
          // Owner is column 5 in Task Management sheet - apply to data rows only
          taskSheet.getRange(2, 5, numRows).setDataValidation(ownerRule);
          updated.push("Owner");
        } else {
          Logger.log("No people found in People sheet for Owner dropdown");
        }
      } else {
        Logger.log("People sheet has only a header row, no people to populate Owner dropdown");
      }
    } else {
      Logger.log("People sheet not found, cannot create Owner dropdown");
    }
  }
  
  // Fetch session titles from Schedule sheet for Related Session dropdown
  const scheduleSheet = sheets.schedule;
  if (scheduleSheet) {
    const lastScheduleRow = scheduleSheet.getLastRow();
    if (lastScheduleRow > 1) {
      // Get all session titles from row 2 to last row - properly excludes header
      const dataRowCount = lastScheduleRow - 1; // Exclude header row
      const sessionRange = scheduleSheet.getRange(2, 5, dataRowCount, 1);
      const sessionValues = sessionRange.getValues();
      
      // Filter non-empty session titles
      const allSessions = sessionValues
        .filter(row => row[0])
        .map(row => row[0]);
      
      if (allSessions.length > 0) {
        const sessionRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(allSessions, true)
          .build();
        // Related Session is column 9 in Task Management sheet - apply to data rows only
        taskSheet.getRange(2, 9, numRows).setDataValidation(sessionRule);
        updated.push("Related Session");
      }
    }
  }
  
  return updated;
}

/**
 * Sets the Event Type dropdown in the Event Description sheet
 * Updated to pull options from Config sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Event Description sheet
 * @param {Object} lists Configuration lists
 * @return {boolean} True if dropdown was set, false otherwise
 */
function setEventTypeDropdown(sheet, lists) {
  if (!sheet) return false;
  
  // Find the row containing "Single- or Multi-Day?"
  const eventTypeRow = _findRow(sheet, 'Single- or Multi-Day?');
  if (!eventTypeRow) return false;
  
  // Only proceed if we have options from Config
  if (!lists || !lists['Event Type'] || !lists['Event Type'].length) return false;
  
  // Create validation rule
  const eventTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(lists['Event Type'], true)
    .build();
  
  // Apply the validation to the cell next to "Single- or Multi-Day?"
  sheet.getRange(eventTypeRow, 2).setDataValidation(eventTypeRule);
  
  return true;
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
 * Legacy function for backward compatibility - now redirects to updateAllDropdowns
 */
function wireUpDropdowns() {
  updateAllDropdowns();
}

/**
 * Legacy function for backward compatibility - now redirects to updateAllDropdowns
 */
function wireUpTaskManagementDropdowns() {
  updateAllDropdowns();
}

/**
 * Legacy function for backward compatibility - now redirects to updateAllDropdowns
 */
function updateScheduleLeadDropdown() {
  updateAllDropdowns();
}