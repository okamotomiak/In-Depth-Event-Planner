//FormGenerator.gs 

/**
 * Adds the "Generate Google Forms" menu item to the Event Planner menu
 * This function should be called from onOpen() in Core.gs
 */
function addFormGeneratorMenuItem(menu) {
  return menu.addItem('Generate Google Forms', 'showFormGeneratorDialog');
}

/**
 * Shows a dialog to select which forms to generate
 * Now modified to automatically generate all forms without asking
 */
function showFormGeneratorDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the event name to display in the dialog
  const eventName = getEventName();
  if (!eventName) {
    ui.alert('Error', 'Event name not found in Event Description sheet. Please complete the Event Description sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  // Check if Config sheet has the necessary form entries
  if (!checkConfigEntries()) {
    // If form entries don't exist, run the setupConfigSheet function
    setupConfigSheet();
  }
  
  // Delete ALL existing form submission triggers first to prevent trigger limit errors
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    const trigger = allTriggers[i];
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Automatically generate all three forms without asking
  const regFormUrl = createRegistrationForm();
  const volFormUrl = createVolunteerSignupForm();
  const speakerFormUrl = createSpeakerBioForm();
  
  // Show summary message to the user
  let message = 'Form Generation Complete:\n\n';
  
  if (regFormUrl) message += '✓ Registration Form created\n';
  else message += '✗ Registration Form failed\n';
  
  if (volFormUrl) message += '✓ Volunteer Signup Form created\n';
  else message += '✗ Volunteer Signup Form failed\n';
  
  if (speakerFormUrl) message += '✓ Speaker Bio Form created\n';
  else message += '✗ Speaker Bio Form failed\n';
  
  message += '\nAll form links have been saved to the Config sheet.';
  ui.alert('Forms Generated', message, ui.ButtonSet.OK);
}

/**
 * Checks if Config sheet has the necessary form entries
 * @return {boolean} True if all form entries exist, false otherwise
 */
function checkConfigEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) return false;
  
  // Get all data from Config sheet
  const data = configSheet.getDataRange().getValues();
  
  // Check for the form keys - now they should be after the OpenAI API Key
  const formKeys = ['registration form', 'volunteer sign up', 'speaker form'];
  let allFound = true;
  
  // Check if each key exists
  formKeys.forEach(key => {
    const found = data.some(row => 
      row[0] && row[0].toString().toLowerCase() === key.toLowerCase()
    );
    
    if (!found) {
      Logger.log(`Form key not found in Config sheet: ${key}`);
      allFound = false;
    }
  });
  
  return allFound;
}

/**
 * Creates a registration form for event participants
 * @return {string|null} The URL of the created form, or null if creation failed
 */
function createRegistrationForm() {
  try {
    const eventName = getEventName();
    if (!eventName) {
      throw new Error('Event name not found in Event Description sheet');
    }
    
    // Get the current spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create a new form
    const formTitle = `${eventName} - Participant Registration`;
    const form = FormApp.create(formTitle)
      .setTitle(formTitle)
      .setDescription(`Registration form for ${eventName}. Please fill out all required fields.`)
      .setConfirmationMessage(`Thank you for registering for ${eventName}. We'll be in touch with more details soon!`)
      .setAllowResponseEdits(true)
      .setCollectEmail(true);
    
    // Add standard questions
    form.addTextItem()
      .setTitle('Full Name')
      .setRequired(true);
    
    // NOTE: Email is already collected by setCollectEmail(true) above
    // No need to add a separate email field
      
    form.addTextItem()
      .setTitle('Phone Number')
      .setRequired(false);
      
    form.addParagraphTextItem()
      .setTitle('Dietary Restrictions or Preferences')
      .setRequired(false);
      
    form.addParagraphTextItem()
      .setTitle('Accessibility Needs')
      .setRequired(false);
      
    // Add a multiple choice question about how they heard about the event
    form.addMultipleChoiceItem()
      .setTitle('How did you hear about this event?')
      .setChoiceValues(['Email', 'Social Media', 'Website', 'Word of Mouth', 'Other'])
      .setRequired(false);
    
    // Set up response destination
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Set up a trigger for the form, if possible
    try {
      ScriptApp.newTrigger('processRegistrationForm')
        .forForm(form.getId())
        .onFormSubmit()
        .create();
    } catch (triggerError) {
      // Log but continue if trigger creation fails
      Logger.log(`Warning: Could not create trigger for registration form: ${triggerError}`);
    }
    
    // Save the form URL to the Config sheet
    saveFormUrl('registration form', form.getPublishedUrl(), form.getEditUrl());
    
    // Return the URL
    return form.getPublishedUrl();
    
  } catch (error) {
    Logger.log(`Error creating registration form: ${error}`);
    return null;
  }
}

/**
 * Creates a signup form for event volunteers
 * @return {string|null} The URL of the created form, or null if creation failed
 */
function createVolunteerSignupForm() {
  try {
    const eventName = getEventName();
    if (!eventName) {
      throw new Error('Event name not found in Event Description sheet');
    }
    
    // Get the current spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create a new form
    const formTitle = `${eventName} - Volunteer Signup`;
    const form = FormApp.create(formTitle)
      .setTitle(formTitle)
      .setDescription(`Volunteer signup form for ${eventName}. Join our team to help make this event a success!`)
      .setConfirmationMessage(`Thank you for volunteering for ${eventName}. We'll contact you soon with more details!`)
      .setAllowResponseEdits(true)
      .setCollectEmail(true);
    
    // Add standard questions
    form.addTextItem()
      .setTitle('Full Name')
      .setRequired(true);
    
    // NOTE: Email is already collected by setCollectEmail(true) above
    // No need to add a separate email field
      
    form.addTextItem()
      .setTitle('Phone Number')
      .setRequired(true);
      
    // Add availability question
    form.addCheckboxItem()
      .setTitle('Availability')
      .setChoiceValues(['Setup Day', 'Event Day - Morning', 'Event Day - Afternoon', 'Event Day - Evening', 'Cleanup Day'])
      .setRequired(true);
      
    // Add preferred roles
    form.addCheckboxItem()
      .setTitle('Preferred Role/Area')
      .setChoiceValues(['Registration', 'Setup/Teardown', 'Technical Support', 'Food & Beverage', 'Logistics', 'Communications', 'General Support'])
      .setRequired(true);
      
    // Add T-shirt size
    form.addMultipleChoiceItem()
      .setTitle('T-Shirt Size')
      .setChoiceValues(['XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL'])
      .setRequired(true);
      
    // Add experience question
    form.addParagraphTextItem()
      .setTitle('Previous Volunteer Experience')
      .setRequired(false);
    
    // Set up response destination
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Set up a trigger for the form, if possible
    try {
      ScriptApp.newTrigger('processVolunteerForm')
        .forForm(form.getId())
        .onFormSubmit()
        .create();
    } catch (triggerError) {
      // Log but continue if trigger creation fails
      Logger.log(`Warning: Could not create trigger for volunteer form: ${triggerError}`);
    }
    
    // Save the form URL to the Config sheet
    saveFormUrl('volunteer sign up', form.getPublishedUrl(), form.getEditUrl());
    
    // Return the URL
    return form.getPublishedUrl();
    
  } catch (error) {
    Logger.log(`Error creating volunteer signup form: ${error}`);
    return null;
  }
}

/**
 * Creates a bio form for event speakers
 * @return {string|null} The URL of the created form, or null if creation failed
 */
function createSpeakerBioForm() {
  try {
    const eventName = getEventName();
    if (!eventName) {
      throw new Error('Event name not found in Event Description sheet');
    }
    
    // Get the current spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create a new form
    const formTitle = `${eventName} - Speaker Information`;
    const form = FormApp.create(formTitle)
      .setTitle(formTitle)
      .setDescription(`Speaker information form for ${eventName}. Please provide your details for our program.`)
      .setConfirmationMessage(`Thank you for submitting your speaker information for ${eventName}.`)
      .setAllowResponseEdits(true)
      .setCollectEmail(true);
    
    // Add standard questions
    form.addTextItem()
      .setTitle('Full Name')
      .setRequired(true);
    
    // NOTE: Email is already collected by setCollectEmail(true) above
    // No need to add a separate email field
      
    form.addTextItem()
      .setTitle('Phone Number')
      .setRequired(true);
      
    form.addTextItem()
      .setTitle('Session Title')
      .setRequired(true);
      
    form.addParagraphTextItem()
      .setTitle('Session Description')
      .setRequired(true)
      .setHelpText('Please provide a brief description (50-100 words) of your session for the event program.');
      
    form.addParagraphTextItem()
      .setTitle('Speaker Bio')
      .setRequired(true)
      .setHelpText('Please provide a brief professional bio (50-100 words) for the event program.');
      
    // Headshot photo link
    form.addParagraphTextItem()
      .setTitle('Headshot Photo Link')
      .setRequired(false)
      .setHelpText('If you have a professional headshot available online, please provide the URL link here. Otherwise, we will contact you separately to request one.');
      
    // Add AV and setup needs
    form.addCheckboxItem()
      .setTitle('AV/Technical Requirements')
      .setChoiceValues(['Projector', 'Audio Connection', 'Microphone', 'Internet Connection', 'Whiteboard/Flip Chart', 'Other (specify in notes)'])
      .setRequired(false);
      
    form.addParagraphTextItem()
      .setTitle('Additional Notes or Requirements')
      .setRequired(false);
    
    // Set up response destination
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Set up a trigger for the form, if possible
    try {
      ScriptApp.newTrigger('processSpeakerForm')
        .forForm(form.getId())
        .onFormSubmit()
        .create();
    } catch (triggerError) {
      // Log but continue if trigger creation fails
      Logger.log(`Warning: Could not create trigger for speaker form: ${triggerError}`);
    }
    
    // Save the form URL to the Config sheet
    saveFormUrl('speaker form', form.getPublishedUrl(), form.getEditUrl());
    
    // Return the URL
    return form.getPublishedUrl();
    
  } catch (error) {
    Logger.log(`Error creating speaker bio form: ${error}`);
    return null;
  }
}

/**
 * Processes registration form submissions
 * @param {Object} e - The form submit event
 */
function processRegistrationForm(e) {
  try {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    const email = formResponse.getRespondentEmail();
    
    // Extract the form responses
    let data = {
      name: '',
      email: email,
      phone: '',
      category: 'Participant',
      status: 'Registered',
      role: '',
      notes: ''
    };
    
    // Process each item response
    itemResponses.forEach(itemResponse => {
      const question = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      
      // Map the answers to the data object
      if (question === 'Full Name') {
        data.name = answer;
      } else if (question === 'Phone Number') {
        data.phone = answer;
      } else if (question === 'Dietary Restrictions or Preferences' || 
                question === 'Accessibility Needs' ||
                question === 'How did you hear about this event?') {
        // Add key info to notes if needed
        if (answer) {
          data.notes += `${question}: ${answer}\n`;
        }
      }
    });
    
    // Add or update the person in the People sheet
    addOrUpdatePersonInPeopleSheet(data);
    
  } catch (error) {
    Logger.log(`Error processing registration form submission: ${error}`);
  }
}

/**
 * Processes volunteer form submissions
 * @param {Object} e - The form submit event
 */
function processVolunteerForm(e) {
  try {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    const email = formResponse.getRespondentEmail();
    
    // Extract the form responses
    let data = {
      name: '',
      email: email,
      phone: '',
      category: 'Volunteer',
      status: 'Potential',
      role: '',
      notes: ''
    };
    
    // Process each item response
    itemResponses.forEach(itemResponse => {
      const question = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      
      // Map the answers to the data object
      if (question === 'Full Name') {
        data.name = answer;
      } else if (question === 'Phone Number') {
        data.phone = answer;
      } else if (question === 'Preferred Role/Area') {
        // Store the first selected role as the primary role
        if (Array.isArray(answer) && answer.length > 0) {
          data.role = answer[0];
          // Add all roles to notes
          data.notes += `Preferred Roles: ${answer.join(', ')}\n`;
        }
      } else if (question === 'Availability' || 
                question === 'T-Shirt Size' ||
                question === 'Previous Volunteer Experience') {
        // Add to notes
        if (answer) {
          if (Array.isArray(answer)) {
            data.notes += `${question}: ${answer.join(', ')}\n`;
          } else {
            data.notes += `${question}: ${answer}\n`;
          }
        }
      }
    });
    
    // Add or update the person in the People sheet
    addOrUpdatePersonInPeopleSheet(data);
    
  } catch (error) {
    Logger.log(`Error processing volunteer form submission: ${error}`);
  }
}

/**
 * Processes speaker form submissions
 * @param {Object} e - The form submit event
 */
function processSpeakerForm(e) {
  try {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    const email = formResponse.getRespondentEmail();
    
    // Extract the form responses
    let data = {
      name: '',
      email: email,
      phone: '',
      category: 'Speaker',
      status: 'Potential',
      role: '',
      notes: ''
    };
    
    // Process each item response
    itemResponses.forEach(itemResponse => {
      const question = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      
      // Map the answers to the data object
      if (question === 'Full Name') {
        data.name = answer;
      } else if (question === 'Phone Number') {
        data.phone = answer;
      } else if (question === 'Session Title') {
        data.role = answer;
      } else if (question === 'Session Description' || 
                question === 'Speaker Bio' ||
                question === 'Headshot Photo Link' ||
                question === 'AV/Technical Requirements' ||
                question === 'Additional Notes or Requirements') {
        // Add to notes
        if (answer) {
          if (Array.isArray(answer)) {
            data.notes += `${question}: ${answer.join(', ')}\n`;
          } else {
            data.notes += `${question}: ${answer}\n`;
          }
        }
      }
    });
    
    // Add or update the person in the People sheet
    addOrUpdatePersonInPeopleSheet(data);
    
  } catch (error) {
    Logger.log(`Error processing speaker form submission: ${error}`);
  }
}

/**
 * Adds a new person to the People sheet or updates an existing entry if the email already exists
 * @param {Object} data - The person data to add or update
 */
function addOrUpdatePersonInPeopleSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const peopleSheet = ss.getSheetByName('People');
  
  if (!peopleSheet) {
    Logger.log('People sheet not found');
    return;
  }
  
  // Get the headers to ensure we're adding data to the right columns
  const headers = peopleSheet.getRange(1, 1, 1, peopleSheet.getLastColumn()).getValues()[0];
  
  // Find the column indices
  const nameColIndex = headers.findIndex(header => header === 'Name') + 1;
  const categoryColIndex = headers.findIndex(header => header === 'Category') + 1;
  const roleColIndex = headers.findIndex(header => header === 'Role/Position') + 1;
  const statusColIndex = headers.findIndex(header => header === 'Status') + 1;
  const emailColIndex = headers.findIndex(header => header === 'Email') + 1;
  const phoneColIndex = headers.findIndex(header => header === 'Phone') + 1;
  const assignedTasksColIndex = headers.findIndex(header => header === 'Assigned Tasks') + 1;
  
  // Check if all required columns exist
  if (!nameColIndex || !categoryColIndex || !emailColIndex) {
    Logger.log('Required columns not found in People sheet');
    return;
  }
  
  // Check if this person already exists in the sheet (by email)
  const allData = peopleSheet.getDataRange().getValues();
  let existingRowIndex = -1;
  
  // Skip header row, start from row 1 (index 0)
  for (let i = 1; i < allData.length; i++) {
    // Check if this row has the same email
    if (allData[i][emailColIndex - 1] === data.email) {
      existingRowIndex = i;
      break;
    }
  }
  
  // Format the row data
  const rowData = [];
  if (nameColIndex) rowData[nameColIndex - 1] = data.name || '';
  if (categoryColIndex) rowData[categoryColIndex - 1] = data.category || '';
  if (roleColIndex) rowData[roleColIndex - 1] = data.role || '';
  if (statusColIndex) rowData[statusColIndex - 1] = data.status || '';
  if (emailColIndex) rowData[emailColIndex - 1] = data.email || '';
  if (phoneColIndex) rowData[phoneColIndex - 1] = data.phone || '';
  
  // Update or add the row
  if (existingRowIndex !== -1) {
    // Existing record found - update it
    const currentRow = existingRowIndex + 1; // Convert to 1-based index
    
    // Get the current assigned tasks (preserve this value)
    if (assignedTasksColIndex) {
      const currentAssignedTasks = peopleSheet.getRange(currentRow, assignedTasksColIndex).getValue();
      rowData[assignedTasksColIndex - 1] = currentAssignedTasks;
    } else {
      rowData[assignedTasksColIndex - 1] = '';
    }
    
    // Fill any undefined values with empty strings
    for (let i = 0; i < headers.length; i++) {
      if (rowData[i] === undefined) {
        rowData[i] = '';
      }
    }
    
    // Update the existing row
    peopleSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
    Logger.log(`Updated existing person: ${data.name} (${data.email}) at row ${currentRow}`);
  } else {
    // No existing record - add a new one
    
    // Fill in values for the new row
    if (assignedTasksColIndex) rowData[assignedTasksColIndex - 1] = '';
    
    // Fill any undefined values with empty strings
    for (let i = 0; i < headers.length; i++) {
      if (rowData[i] === undefined) {
        rowData[i] = '';
      }
    }
    
    // Add the new row at the end
    const lastRow = peopleSheet.getLastRow();
    const newRow = lastRow + 1;
    peopleSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    Logger.log(`Added new person: ${data.name} (${data.email}) to row ${newRow}`);
  }
}

/**
 * Gets the event name from the Event Description sheet
 * @return {string|null} The event name or null if not found
 */
function getEventName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventDescSheet = ss.getSheetByName('Event Description');
  
  if (!eventDescSheet) {
    Logger.log('Event Description sheet not found');
    return null;
  }
  
  // Use the _findRow helper to find the Event Name row
  const eventNameRow = _findRow(eventDescSheet, 'Event Name');
  if (!eventNameRow) {
    Logger.log('Event Name row not found in Event Description sheet');
    return null;
  }
  
  // Get the event name from column B
  const eventName = eventDescSheet.getRange(eventNameRow, 2).getValue();
  return eventName ? eventName.toString() : null;
}

/**
 * Saves a form URL to the Config sheet with hyperlink
 * @param {string} key - The key in the Config sheet (e.g., 'registration form')
 * @param {string} url - The form URL to save (published URL)
 * @param {string} editUrl - The edit URL for the form
 * @return {boolean} True if successful, false otherwise
 */
function saveFormUrl(key, url, editUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    Logger.log('Config sheet not found');
    return false;
  }
  
  // Find the row with the specified key
  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === key.toLowerCase()) {
      // Save the URL in the Value/Subject column (column B) as a hyperlink
      const range = configSheet.getRange(i + 1, 2);
      
      // Set the hyperlink and use the form type as the link text
      const formType = key.charAt(0).toUpperCase() + key.slice(1); // Capitalize first letter
      range.setFormula(`=HYPERLINK("${url}","${formType}")`);
      
      Logger.log(`Form URL saved for key: ${key}`);
      return true;
    }
  }
  
  // Key not found
  Logger.log(`Key not found in Config sheet: ${key}`);
  return false;
}

/**
 * Helper function to find a row by field name in column A
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object
 * @param {string} label The label to find in column A
 * @return {number|null} The 1-based row number, or null if not found
 */
function _findRow(sheet, label) {
  const vals = sheet.getRange('A:A').getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === label) return i + 1;
  }
  return null;
}