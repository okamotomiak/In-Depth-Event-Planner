// CueBuilder.js - Contains functions to set up and manage the Cue Builder sheet.

/**
 * Creates and sets up the "Cue Builder" sheet.
 * This sheet is used for creating the detailed, step-by-step program flow.
 */
function setupCueBuilderSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  let sheet = ss.getSheetByName('Cue Builder');
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Cue Builder');
    sheet.setTabColor('#ff6d01'); // A distinct orange color
    ui.alert('"Cue Builder" sheet created!', 'Please use this sheet to build your detailed program flow.', ui.ButtonSet.OK);
  } else {
    const response = ui.alert(
      'Reset Cue Builder?',
      'The "Cue Builder" sheet already exists. Do you want to clear it and reset it with default formatting?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
    sheet.clear();
  }
  
  // Define headers for the Cue Builder
  const headers = [
    'Schedule Item',
    'Cue Title',
    'Lead / Talent',
    'Est. Duration (Mins)',
    'MC Script / Notes',
    'Lighting Cue',
    'Audio / Sound Cue',
    'Visuals / Screen Cue'
  ];
  
  // Set header values and formatting
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#434343').setFontColor('#ffffff').setFontWeight('bold');
  
  // Set column widths
  const widths = [180, 200, 150, 100, 400, 200, 200, 200];
  widths.forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });
  
  // Freeze header row and apply text wrapping
  sheet.setFrozenRows(1);
  sheet.getRange('A2:H').setWrap(true).setVerticalAlignment('top');
  
  // Apply alternating row colors
  sheet.getRange('A2:H').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  
  // Populate the dropdowns
  updateCueBuilderDropdowns();
}

/**
 * Populates the "Schedule Item" dropdown in the "Cue Builder" sheet
 * with titles from the "Schedule" sheet.
 */
function updateCueBuilderDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName('Schedule');
  const cueBuilderSheet = ss.getSheetByName('Cue Builder');
  
  if (!scheduleSheet || !cueBuilderSheet) {
    Logger.log('Required sheets (Schedule or Cue Builder) not found for dropdown creation.');
    return;
  }
  
  // Get session titles from the Schedule sheet
  const lastRow = scheduleSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Schedule sheet to create dropdowns from.');
    return;
  }
  
  const sessionTitlesRange = scheduleSheet.getRange(2, 5, lastRow - 1, 1);
  const sessionTitles = sessionTitlesRange.getValues()
    .flat()
    .filter(title => title.toString().trim() !== ''); // Filter out empty titles
    
  if (sessionTitles.length === 0) {
    Logger.log('No session titles found in Schedule sheet.');
    return;
  }
  
  // Create a data validation rule (dropdown)
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(sessionTitles, true)
    .setAllowInvalid(false)
    .setHelpText('Select a schedule block for this cue.')
    .build();
    
  // Apply the rule to the "Schedule Item" column in the Cue Builder sheet
  const dropdownColumn = cueBuilderSheet.getRange('A2:A');
  dropdownColumn.setDataValidation(rule);
  
  Logger.log('Successfully updated "Schedule Item" dropdown in Cue Builder sheet.');
}