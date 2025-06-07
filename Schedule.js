//Schedule.gs

/**
 * Sets up the Schedule sheet with headers, formatting, and sample data.
 * Fixed to ensure proper alternating colors for all data rows.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object
 * @param {boolean} addSampleData Whether to add sample data to the sheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The configured Schedule sheet
 */
/**
 * Sets up the Schedule sheet with headers, formatting, and sample data.
 * Fixed to ensure proper alternating colors for all data rows.
 * MODIFIED: Added an "Add to Cue" checkbox column.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object
 * @param {boolean} addSampleData Whether to add sample data to the sheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The configured Schedule sheet
 */
function setupScheduleSheet(ss, addSampleData = true) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName('Schedule');
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Schedule');
    sheet.setTabColor('#f1c232'); // Yellow/gold color
  } else {
    // Clear existing content if sheet already exists
    sheet.clear();
  }
  
  // Ensure we have at least 900 rows
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < 900) {
    sheet.insertRowsAfter(currentMaxRows, 900 - currentMaxRows);
  }
  
  // Define headers with the new "Add to Cue" column
  const headers = ['Date', 'Start Time', 'End Time', 'Duration', 'Session Title', 'Lead', 'Location', 'Status', 'Add to Cue', 'Notes'];
  
  // Set header values
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Set column widths
  const widths = [100, 100, 100, 100, 200, 150, 150, 120, 80, 300];
  for (let i = 0; i < headers.length; i++) {
    if (i < widths.length) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Add sample data if requested
  if (addSampleData) {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const dayAfter = new Date();
    dayAfter.setDate(dayAfter.getDate() + 2);
    
    const sampleData = [
      [tomorrow, '9:00 AM', '10:00 AM', '1 hour', 'Opening Session', 'Jane Doe', 'Main Hall', 'Confirmed', true, 'Welcome address and introduction'],
      [dayAfter, '1:00 PM', '3:00 PM', '2 hours', 'Workshop', 'John Smith', 'Room 101', 'Tentative', false, 'Interactive workshop session']
    ];
    
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  }
  
  // Apply data validations to data rows ONLY (rows 2-900, not header)
  const locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Main Hall', 'Room 101', 'Room 102', 'Outdoor Area'], true)
    .build();
    
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Tentative', 'Confirmed', 'Cancelled'], true)
    .build();
  
  // Apply validation rules to all data rows (rows 2-900, not header)
  sheet.getRange(2, 7, 899, 1).setDataValidation(locationRule); // Location (column G)
  sheet.getRange(2, 8, 899, 1).setDataValidation(statusRule); // Status (column H)
  
  // Insert checkboxes for the "Add to Cue" column
  sheet.getRange(2, 9, 899, 1).insertCheckboxes();
  
  // Set number formats in batch for all rows
  sheet.getRange(2, 1, 899, 1).setNumberFormat('yyyy-mm-dd'); // Date (column A)
  sheet.getRange(2, 2, 899, 2).setNumberFormat('hh:mm am/pm'); // Start/End Time (columns B-C)
  
  // Format headers with blue background and white text
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8') // Blue background
    .setFontColor('#ffffff') // White text
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set normal white background for the first data row and all other rows
  sheet.getRange(2, 1, 899, headers.length).setBackground(null);
  
  // Apply custom alternating row colors - manually setting alternating colors
  for (let i = 2; i <= 900; i += 2) {
    // For even numbered rows, apply light gray background
    sheet.getRange(i, 1, 1, headers.length).setBackground('#f3f3f3');
  }
  
  // Create filter view for all rows
  sheet.getRange(1, 1, 900, headers.length).createFilter();
  
  // Setup duration calculation
  setupDurationCalculation(ss);
  
  return sheet;
}

/**
 * Sets up duration calculation in the Schedule sheet
 * Fixed to properly display 60m as 1h with a tolerance range
 */
function setupDurationCalculation(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheet = ss.getSheetByName('Schedule');
  if (!sheet) return;
  
  // Get the last row with data
  const lastRow = Math.max(2, sheet.getLastRow());
  
  // Skip if only header row exists
  if (lastRow <= 1) return;
  
  // Format Start Time and End Time columns correctly first
  // Use a simple time format that works reliably with calculations
  sheet.getRange(2, 2, lastRow-1, 2).setNumberFormat('h:mm AM/PM');
  
  // Set formula for duration calculation in column 4 (Duration)
  // Using a refined formula with tolerance range for 60 minutes
  for (let row = 2; row <= lastRow; row++) {
    const formula = `=IF(AND(B${row}<>"",C${row}<>""),
      IF(
        AND(((C${row}-B${row})*24*60)>=59.98, ((C${row}-B${row})*24*60)<=60.02),
        "1h",
        IF(
          MOD((C${row}-B${row})*24*60,60)=0,
          TEXT(INT((C${row}-B${row})*24),"0") & "h",
          IF(
            INT((C${row}-B${row})*24)=0,
            TEXT(MOD((C${row}-B${row})*24*60,60),"0") & "m",
            TEXT(INT((C${row}-B${row})*24),"0") & "h " & TEXT(MOD((C${row}-B${row})*24*60,60),"0") & "m"
          )
        )
      ),
    "")`;
    sheet.getRange(row, 4).setFormula(formula);
  }
}

/**
 * Handles edits specifically for the Schedule sheet
 * Called from the consolidated onEdit function in Core.gs
 * @param {Object} e The edit event object
 */
function handleScheduleEdit(e) {
  // Only proceed if edit is in Schedule sheet
  if (!e || !e.range || e.range.getSheet().getName() !== 'Schedule') return;
  
  // Only proceed if edit is in column B (Start Time) or C (End Time)
  const col = e.range.getColumn();
  if (col !== 2 && col !== 3) return;
  
  // Get the row being edited
  const row = e.range.getRow();
  if (row <= 1) return; // Skip header row
  
  // Get the sheet and values
  const sheet = e.range.getSheet();
  const startTimeCell = sheet.getRange(row, 2);
  const endTimeCell = sheet.getRange(row, 3);
  const durationCell = sheet.getRange(row, 4);
  
  const startTimeValue = startTimeCell.getValue();
  const endTimeValue = endTimeCell.getValue();
  
  // Skip if either time is missing
  if (!startTimeValue || !endTimeValue) {
    durationCell.setValue("");
    return;
  }
  
  // Parse the time strings into Date objects
  try {
    // Create Date objects for today with the specified times
    const startTime = parseTimeString(startTimeValue);
    const endTime = parseTimeString(endTimeValue);
    
    if (!startTime || !endTime) {
      durationCell.setValue("Format Error");
      return;
    }
    
    // Calculate duration in milliseconds
    let durationMs = endTime.getTime() - startTime.getTime();
    
    // Handle case where end time is on the next day
    if (durationMs < 0) {
      durationMs += 24 * 60 * 60 * 1000; // Add a day in milliseconds
    }
    
    // Convert to total minutes first
    const totalMinutes = Math.floor(durationMs / (60 * 1000));
    
    // Format the duration string in simplified format with 60m = 1h fix
    let durationStr = "";
    
    // Add a small tolerance for 60 minutes to handle floating point issues
    if (totalMinutes >= 59 && totalMinutes <= 61) {
      // Approximately 60 minutes = 1 hour
      durationStr = "1h";
    } else if (totalMinutes > 60) {
      // More than an hour
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      
      if (minutes === 0) {
        // Full hours, no minutes
        durationStr = hours + "h";
      } else {
        // Hours and minutes
        durationStr = hours + "h " + minutes + "m";
      }
    } else {
      // Less than an hour, just minutes
      durationStr = totalMinutes + "m";
    }
    
    durationCell.setValue(durationStr);
    
  } catch (error) {
    durationCell.setValue("Error");
    Logger.log("Error calculating duration: " + error.toString());
  }
}

/**
 * Helper function to parse a time string in the format "10:00 AM" or "2:30 PM"
 * @param {string|Date} timeStr - The time string or Date object to parse
 * @return {Date|null} - A Date object set to today with the specified time, or null if parsing fails
 */
function parseTimeString(timeStr) {
  try {
    // Handle Date objects directly
    if (timeStr instanceof Date) {
      return timeStr;
    }
    
    // Make sure we have a string
    timeStr = String(timeStr).trim();
    
    // Regular expression to match time formats with more flexibility
    // This will match "10:00 AM", "2:30PM", "14:00", etc.
    const timeRegex = /^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?$/;
    const match = timeStr.match(timeRegex);
    
    if (!match) {
      // Try an alternative format with period separator like "2.30 PM"
      const altTimeRegex = /^(\d{1,2})[\.\s](\d{2})\s*(AM|PM|am|pm)?$/;
      const altMatch = timeStr.match(altTimeRegex);
      
      if (!altMatch) return null;
      
      // Use the alternative match
      let hours = parseInt(altMatch[1], 10);
      const minutes = parseInt(altMatch[2], 10);
      const period = altMatch[3] ? altMatch[3].toUpperCase() : null;
      
      // Convert to 24-hour format if period is specified
      if (period === 'PM' && hours < 12) {
        hours += 12;
      } else if (period === 'AM' && hours === 12) {
        hours = 0;
      }
      
      // Create a Date object for today with the parsed time
      const date = new Date();
      date.setHours(hours, minutes, 0, 0);
      return date;
    }
    
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const period = match[3] ? match[3].toUpperCase() : null;
    
    // If period is specified, convert to 24-hour format
    if (period === 'PM' && hours < 12) {
      hours += 12;
    } else if (period === 'AM' && hours === 12) {
      hours = 0;
    }
    
    // Create a Date object for today with the parsed time
    const date = new Date();
    date.setHours(hours, minutes, 0, 0);
    return date;
    
  } catch (error) {
    Logger.log("Error parsing time string: " + error.toString());
    return null;
  }
}

/**
 * Standalone function to update duration calculation in Schedule sheet
 * This can be called from the menu to apply duration calculation to existing data
 */
function updateScheduleDurations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDurationCalculation(ss);
  SpreadsheetApp.getUi().alert('Duration calculation has been updated in the Schedule sheet.');
}

/**
 * Sets dropdowns for Lead, Location & Status in Schedule.
 * Updated to get Lead dropdown from People sheet names.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @param {Object} sheets Cached sheet references
 * @param {Object} rowCounts Cached row counts
 * @param {Object} lists Configuration lists
 * @return {Array} List of updated dropdown fields
 */
function setScheduleDropdowns(ss, sheets, rowCounts, lists) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!sheets) {
    sheets = {
      schedule: ss.getSheetByName('Schedule'),
      people: ss.getSheetByName('People')
    };
  }
  
  const scheduleSheet = sheets.schedule;
  if (!scheduleSheet) return [];
  
  const numRows = rowCounts && rowCounts.schedule 
    ? rowCounts.schedule 
    : Math.max(10, scheduleSheet.getLastRow() - 1);
    
  const updated = [];
  
  // Set Location dropdown
  if (lists && lists['Location List'] && lists['Location List'].length) {
    const locationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lists['Location List'], true)
      .build();
    // Starting from row 2 (first data row)
    scheduleSheet.getRange(2, 7, numRows).setDataValidation(locationRule);
    updated.push("Location");
  }
  
  // Set Status dropdown
  if (lists && lists['Status Options'] && lists['Status Options'].length) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lists['Status Options'], true)
      .build();
    // Starting from row 2 (first data row)
    scheduleSheet.getRange(2, 8, numRows).setDataValidation(statusRule);
    updated.push("Status");
  }
  
  // Update Lead dropdown with people from People sheet
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
        // Create validation rule with all people names
        const leadRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(allPeople, true)
          .build();
        
        // Apply to Lead column (column 6) - data rows only
        scheduleSheet.getRange(2, 6, numRows).setDataValidation(leadRule);
        updated.push("Lead");
      }
    }
  }
  
  return updated;
}