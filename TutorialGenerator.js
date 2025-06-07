//TutorialGenerator.gs - Creates guided tutorial overlays for the Event Planner

/**
 * Main function to set up the complete tutorial system
 * Run this once to add tutorials to all sheets
 */
function createFullTutorialSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Show confirmation dialog
  const response = ui.alert(
    'Create Tutorial System',
    'This will add tutorial columns to your sheets to help users learn the system. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    // Initialize tutorial progress tracking
    initializeTutorialTracking(ss);
    
    // Add tutorials to each sheet
    addEventDescriptionTutorial(ss);
    addPeopleTutorial(ss);
    addTaskManagementTutorial(ss);
    addScheduleTutorial(ss);
    addBudgetTutorial(ss);
    addDashboardTutorial(ss);
    
    // Show completion message
    ui.alert(
      'Tutorial System Created!',
      'Tutorial guidance has been added to all sheets. Look for the blue tutorial columns on the right side of each sheet.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error creating tutorial system: ' + error.toString());
    ui.alert('Error', 'Failed to create tutorial system: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Initialize tutorial progress tracking in Config sheet
 */
function initializeTutorialTracking(ss) {
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;
  
  // Find a good place to add tutorial tracking (after existing config)
  const lastRow = configSheet.getLastRow();
  const startRow = lastRow + 2;
  
  // Add tutorial tracking section
  const tutorialData = [
    ['=== TUTORIAL PROGRESS ===', ''],
    ['Tutorial Step 1 - Event Setup', 'false'],
    ['Tutorial Step 2 - Add People', 'false'],
    ['Tutorial Step 3 - Generate Tasks', 'false'],
    ['Tutorial Step 4 - Create Schedule', 'false'],
    ['Tutorial Step 5 - Setup Budget', 'false'],
    ['Tutorial Step 6 - Generate Forms', 'false'],
    ['Tutorial Step 7 - View Dashboard', 'false'],
    ['Tutorial Step 8 - All Complete', 'false']
  ];
  
  configSheet.getRange(startRow, 1, tutorialData.length, 2).setValues(tutorialData);
  
  // Format the header
  configSheet.getRange(startRow, 1, 1, 2)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

/**
 * Add tutorial to Event Description sheet
 */
function addEventDescriptionTutorial(ss) {
  const sheet = ss.getSheetByName('Event Description');
  if (!sheet) return;
  
  // Find the tutorial column (start after existing data)
  const lastCol = Math.max(sheet.getLastColumn(), 2);
  const tutorialCol = lastCol + 2; // Leave one empty column
  
  // Tutorial header
  sheet.getRange(1, tutorialCol, 1, 2).merge();
  sheet.getRange(1, tutorialCol).setValue('📚 Getting Started: Step 1 / 8');
  sheet.getRange(1, tutorialCol)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  
  // Progress indicators
  const progressRow = 2;
  const checkboxes = ['☑️', '⬜', '⬜', '⬜', '⬜', '⬜', '⬜', '⬜'];
  sheet.getRange(progressRow, tutorialCol, 1, 2).merge();
  sheet.getRange(progressRow, tutorialCol).setValue(checkboxes.join(' '));
  sheet.getRange(progressRow, tutorialCol).setHorizontalAlignment('center');
  
  // Tutorial content
  const tutorialContent = [
    ['', ''],
    ['Step 1: Set up your event basics', ''],
    ['', ''],
    ['✨ Welcome to your Event Planner!', ''],
    ['', ''],
    ['Let\'s start by filling out your event details.', ''],
    ['This information helps our AI create better', ''],
    ['tasks and schedules for you.', ''],
    ['', ''],
    ['👆 Fill out these key fields:', ''],
    ['• Event Name', ''],
    ['• Start Date (And Time)', ''],
    ['• End Date (And Time)', ''],
    ['• Theme or Focus', ''],
    ['• Description & Messaging', ''],
    ['', ''],
    ['💡 Pro tip: The more detail you provide', ''],
    ['in the description fields, the better our', ''],
    ['AI will understand your event and create', ''],
    ['relevant tasks for you!', ''],
    ['', ''],
    ['When done, go to the People sheet →', ''],
    ['', ''],
    ['Need help? Click here for examples:', ''],
    ['=HYPERLINK("https://help.example.com", "View Sample Events")', '']
  ];
  
  // Add tutorial content
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2).setValues(tutorialContent);
  
  // Format tutorial area
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2)
    .setBackground('#f8f9fa')
    .setFontSize(10)
    .setWrap(true);
  
  // Make tutorial columns narrower
  sheet.setColumnWidth(tutorialCol, 200);
  sheet.setColumnWidth(tutorialCol + 1, 200);
  
  // Add border around tutorial
  sheet.getRange(1, tutorialCol, tutorialContent.length + 3, 2)
    .setBorder(true, true, true, true, true, true, '#4285f4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Add tutorial to People sheet
 */
function addPeopleTutorial(ss) {
  const sheet = ss.getSheetByName('People');
  if (!sheet) return;
  
  const lastCol = Math.max(sheet.getLastColumn(), 7);
  const tutorialCol = lastCol + 2;
  
  // Tutorial header
  sheet.getRange(1, tutorialCol, 1, 2).merge();
  sheet.getRange(1, tutorialCol).setValue('📚 Getting Started: Step 2 / 8');
  sheet.getRange(1, tutorialCol)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  
  // Progress indicators  
  const checkboxes = ['☑️', '☑️', '⬜', '⬜', '⬜', '⬜', '⬜', '⬜'];
  sheet.getRange(2, tutorialCol, 1, 2).merge();
  sheet.getRange(2, tutorialCol).setValue(checkboxes.join(' '));
  sheet.getRange(2, tutorialCol).setHorizontalAlignment('center');
  
  const tutorialContent = [
    ['', ''],
    ['Step 2: Add your team members', ''],
    ['', ''],
    ['👥 Add people who will help with your event:', ''],
    ['', ''],
    ['• Staff - Your core team members', ''],
    ['• Volunteers - People helping out', ''],
    ['• Speakers - Presenters at your event', ''],
    ['• Participants - General attendees', ''],
    ['', ''],
    ['🎯 Try adding one person of each type:', ''],
    ['', ''],
    ['1. Click on row 2 and add a Staff member', ''],
    ['2. Add a Speaker (we\'ll use this later!)', ''],
    ['3. Add a Volunteer', ''],
    ['', ''],
    ['💡 When you set a Speaker\'s status to', ''],
    ['"Accepted", the system automatically', ''],
    ['creates a task to collect their bio!', ''],
    ['', ''],
    ['Next: Go to Task Management sheet →', ''],
    ['', ''],
    ['🔧 You can also auto-generate forms', ''],
    ['later from the Event Planner Setup menu!', '']
  ];
  
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2).setValues(tutorialContent);
  
  // Format tutorial area
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2)
    .setBackground('#f8f9fa')
    .setFontSize(10)
    .setWrap(true);
  
  sheet.setColumnWidth(tutorialCol, 200);
  sheet.setColumnWidth(tutorialCol + 1, 200);
  
  sheet.getRange(1, tutorialCol, tutorialContent.length + 3, 2)
    .setBorder(true, true, true, true, true, true, '#4285f4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Add tutorial to Task Management sheet
 */
function addTaskManagementTutorial(ss) {
  const sheet = ss.getSheetByName('Task Management');
  if (!sheet) return;
  
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const tutorialCol = lastCol + 2;
  
  // Tutorial header
  sheet.getRange(1, tutorialCol, 1, 2).merge();
  sheet.getRange(1, tutorialCol).setValue('📚 Getting Started: Step 3 / 8');
  sheet.getRange(1, tutorialCol)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  
  // Progress indicators
  const checkboxes = ['☑️', '☑️', '☑️', '⬜', '⬜', '⬜', '⬜', '⬜'];
  sheet.getRange(2, tutorialCol, 1, 2).merge();
  sheet.getRange(2, tutorialCol).setValue(checkboxes.join(' '));
  sheet.getRange(2, tutorialCol).setHorizontalAlignment('center');
  
  const tutorialContent = [
    ['', ''],
    ['Step 3: Generate AI-powered tasks', ''],
    ['', ''],
    ['🤖 This is where the magic happens!', ''],
    ['', ''],
    ['Our AI will analyze your event details', ''],
    ['and create a comprehensive task list', ''],
    ['with proper due dates and priorities.', ''],
    ['', ''],
    ['🚀 To generate AI tasks:', ''],
    ['', ''],
    ['1. Go to Event Planner Setup menu', ''],
    ['2. Click "Generate AI Tasks"', ''],
    ['3. Wait for the AI to work (30-60 seconds)', ''],
    ['4. Review the generated tasks below', ''],
    ['', ''],
    ['✨ The AI considers:', ''],
    ['• Your event type and theme', ''],
    ['• Event duration and timing', ''],
    ['• Event description and objectives', ''],
    ['• Industry best practices', ''],
    ['', ''],
    ['💡 You can always add, edit, or delete', ''],
    ['tasks after they\'re generated.', ''],
    ['', ''],
    ['Next: Go to Schedule sheet →', '']
  ];
  
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2).setValues(tutorialContent);
  
  // Format tutorial area
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2)
    .setBackground('#f8f9fa')
    .setFontSize(10)
    .setWrap(true);
  
  sheet.setColumnWidth(tutorialCol, 200);
  sheet.setColumnWidth(tutorialCol + 1, 200);
  
  sheet.getRange(1, tutorialCol, tutorialContent.length + 3, 2)
    .setBorder(true, true, true, true, true, true, '#4285f4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Add tutorial to Schedule sheet
 */
function addScheduleTutorial(ss) {
  const sheet = ss.getSheetByName('Schedule');
  if (!sheet) return;
  
  const lastCol = Math.max(sheet.getLastColumn(), 9);
  const tutorialCol = lastCol + 2;
  
  // Tutorial header
  sheet.getRange(1, tutorialCol, 1, 2).merge();
  sheet.getRange(1, tutorialCol).setValue('📚 Getting Started: Step 4 / 8');
  sheet.getRange(1, tutorialCol)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  
  // Progress indicators
  const checkboxes = ['☑️', '☑️', '☑️', '☑️', '⬜', '⬜', '⬜', '⬜'];
  sheet.getRange(2, tutorialCol, 1, 2).merge();
  sheet.getRange(2, tutorialCol).setValue(checkboxes.join(' '));
  sheet.getRange(2, tutorialCol).setHorizontalAlignment('center');
  
  const tutorialContent = [
    ['', ''],
    ['Step 4: Create your event schedule', ''],
    ['', ''],
    ['📅 Generate a preliminary schedule:', ''],
    ['', ''],
    ['🎯 Option 1: AI-Generated Schedule', ''],
    ['• Go to Event Planner Setup menu', ''],
    ['• Click "Generate Preliminary Schedule"', ''],
    ['• AI creates sessions based on your event', ''],
    ['• Uses approved locations from Config', ''],
    ['• Respects your event time constraints', ''],
    ['', ''],
    ['✋ Option 2: Manual Entry', ''],
    ['• Add sessions manually row by row', ''],
    ['• Duration calculates automatically', ''],
    ['• Use dropdowns for Location and Status', ''],
    ['', ''],
    ['⚡ Pro features:', ''],
    ['• When you change Status to "Confirmed",', ''],
    ['  you\'ll get a confirmation notification', ''],
    ['• Lead dropdown pulls from People sheet', ''],
    ['• Time validation prevents conflicts', ''],
    ['', ''],
    ['Next: Go to Budget sheet →', '']
  ];
  
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2).setValues(tutorialContent);
  
  // Format tutorial area
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2)
    .setBackground('#f8f9fa')
    .setFontSize(10)
    .setWrap(true);
  
  sheet.setColumnWidth(tutorialCol, 200);
  sheet.setColumnWidth(tutorialCol + 1, 200);
  
  sheet.getRange(1, tutorialCol, tutorialContent.length + 3, 2)
    .setBorder(true, true, true, true, true, true, '#4285f4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Add tutorial to Budget sheet
 */
function addBudgetTutorial(ss) {
  const sheet = ss.getSheetByName('Budget');
  if (!sheet) return;
  
  const lastCol = Math.max(sheet.getLastColumn(), 6);
  const tutorialCol = lastCol + 2;
  
  // Tutorial header
  sheet.getRange(1, tutorialCol, 1, 2).merge();
  sheet.getRange(1, tutorialCol).setValue('📚 Getting Started: Step 5 / 8');
  sheet.getRange(1, tutorialCol)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  
  // Progress indicators
  const checkboxes = ['☑️', '☑️', '☑️', '☑️', '☑️', '⬜', '⬜', '⬜'];
  sheet.getRange(2, tutorialCol, 1, 2).merge();
  sheet.getRange(2, tutorialCol).setValue(checkboxes.join(' '));
  sheet.getRange(2, tutorialCol).setHorizontalAlignment('center');
  
  const tutorialContent = [
    ['', ''],
    ['Step 5: Plan your event budget', ''],
    ['', ''],
    ['💰 This budget template includes:', ''],
    ['', ''],
    ['📈 Revenue sections:', ''],
    ['• Registration fees (regular & early bird)', ''],
    ['• Donations (org & individual)', ''],
    ['• Sales & other income', ''],
    ['', ''],
    ['📊 Expense categories:', ''],
    ['• Venue & production costs', ''],
    ['• Program & speaker fees', ''],
    ['• Food & refreshments', ''],
    ['• Lodging & transportation', ''],
    ['• Staff & miscellaneous', ''],
    ['', ''],
    ['🎯 Try entering some sample numbers:', ''],
    ['• Registration fee: $50 per person', ''],
    ['• Expected attendees: 100', ''],
    ['• Venue cost: $2000', ''],
    ['• Catering: $30 per person', ''],
    ['', ''],
    ['📱 Watch the totals update automatically!', ''],
    ['', ''],
    ['Next: Generate Forms →', '']
  ];
  
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2).setValues(tutorialContent);
  
  // Format tutorial area
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2)
    .setBackground('#f8f9fa')
    .setFontSize(10)
    .setWrap(true);
  
  sheet.setColumnWidth(tutorialCol, 200);
  sheet.setColumnWidth(tutorialCol + 1, 200);
  
  sheet.getRange(1, tutorialCol, tutorialContent.length + 3, 2)
    .setBorder(true, true, true, true, true, true, '#4285f4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Add tutorial to Dashboard sheet
 */
function addDashboardTutorial(ss) {
  const sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return;
  
  const lastCol = Math.max(sheet.getLastColumn(), 12);
  const tutorialCol = lastCol + 2;
  
  // Tutorial header  
  sheet.getRange(1, tutorialCol, 1, 2).merge();
  sheet.getRange(1, tutorialCol).setValue('🎉 You did it!');
  sheet.getRange(1, tutorialCol)
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  
  // All completed checkboxes
  const checkboxes = ['☑️', '☑️', '☑️', '☑️', '☑️', '☑️', '☑️', '☑️'];
  sheet.getRange(2, tutorialCol, 1, 2).merge();
  sheet.getRange(2, tutorialCol).setValue(checkboxes.join(' '));
  sheet.getRange(2, tutorialCol).setHorizontalAlignment('center');
  
  const tutorialContent = [
    ['', ''],
    ['You\'re all set up, so what\'s next?', ''],
    ['', ''],
    ['🎯 Your event planner is now ready!', ''],
    ['', ''],
    ['✨ What you\'ve accomplished:', ''],
    ['• Set up your event basics', ''],
    ['• Added team members', ''],
    ['• Generated AI-powered tasks', ''],
    ['• Created a preliminary schedule', ''],
    ['• Planned your budget', ''],
    ['• Ready to generate forms', ''],
    ['', ''],
    ['🚀 Next steps:', ''],
    ['• Generate Google Forms from the menu', ''],
    ['• Assign tasks to team members', ''],
    ['• Confirm speakers and sessions', ''],
    ['• Monitor progress on this dashboard', ''],
    ['', ''],
    ['💡 Remember: Click the 🔄 Refresh button', ''],
    ['to update dashboard metrics anytime!', ''],
    ['', ''],
    ['📚 Need help? Check the Config sheet', ''],
    ['for email templates and settings.', ''],
    ['', ''],
    ['HOW TO HIDE THESE TUTORIAL COLUMNS', ''],
    ['', ''],
    ['Click the ➖ above column ' + getColumnLetter(tutorialCol), ''],
    ['to collapse this tutorial section', ''],
    ['since you\'re all set with your setup.', '']
  ];
  
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2).setValues(tutorialContent);
  
  // Format tutorial area
  sheet.getRange(4, tutorialCol, tutorialContent.length, 2)
    .setBackground('#f8f9fa')
    .setFontSize(10)
    .setWrap(true);
  
  // Special formatting for completion
  sheet.getRange(4, tutorialCol, 5, 2).setBackground('#e8f5e8');
  
  sheet.setColumnWidth(tutorialCol, 200);
  sheet.setColumnWidth(tutorialCol + 1, 200);
  
  sheet.getRange(1, tutorialCol, tutorialContent.length + 3, 2)
    .setBorder(true, true, true, true, true, true, '#34a853', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * Helper function to convert column number to letter
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    columnNumber--;
    letter = String.fromCharCode(65 + (columnNumber % 26)) + letter;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return letter;
}

/**
 * Remove all tutorial columns from all sheets
 */
function removeTutorialSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Remove Tutorial System',
    'This will remove all tutorial columns from your sheets. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const sheetNames = ['Event Description', 'People', 'Task Management', 'Schedule', 'Budget', 'Dashboard'];
  
  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    // Look for tutorial columns (they have blue headers)
    const lastCol = sheet.getLastColumn();
    for (let col = lastCol; col >= 1; col--) {
      const cellValue = sheet.getRange(1, col).getValue();
      if (cellValue.toString().includes('📚 Getting Started') || 
          cellValue.toString().includes('🎉 You did it!')) {
        // Found tutorial column, delete it and the next one
        sheet.deleteColumns(col, 2);
        break;
      }
    }
  });
  
  ui.alert('Tutorial System Removed', 'All tutorial columns have been removed.', ui.ButtonSet.OK);
}

/**
 * Add tutorial functions to the menu
 * Call this from your onOpen function in Core.gs
 */
function addTutorialMenuItems(menu) {
  menu.addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Tutorial System')
        .addItem('Create Tutorial Overlays', 'createFullTutorialSystem')
        .addItem('Remove Tutorial Overlays', 'removeTutorialSystem'));
}