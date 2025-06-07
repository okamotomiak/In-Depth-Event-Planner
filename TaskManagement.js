//TaskManagement.gs 

/**
 * Configuration constants
 */
const TASK_SHEET_NAME = 'Task Management';
const EVENT_DESC_SHEET_NAME = 'Event Description';
const CONFIG_SHEET_NAME = 'Config';
const DEFAULT_PRE_EVENT_DAYS = 4;
const DEFAULT_POST_EVENT_DAYS = 2;

/**
 * Main function to generate AI tasks for the event
 * Can be triggered from a menu item
 */
function generateAITasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  Logger.log('Starting AI Task Generation...');
  
  try {
    // Step 1: Get event information
    const eventInfo = getEventInformation();
    if (!eventInfo) {
      ui.alert('Error', 'Could not retrieve event information. Please ensure the Event Description sheet is properly filled out.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 2: Get OpenAI API key
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      ui.alert('Error', 'OpenAI API key not found. Please add your API key to Script Properties or cell B15 in Config sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 3: Show loading message
    ui.alert('Processing', 'Generating tasks using AI. This may take a few moments...', ui.ButtonSet.OK);
    
    // Step 4: Generate tasks using OpenAI
    const tasks = generateTasksWithAI(eventInfo, apiKey);
    if (!tasks || tasks.length === 0) {
      ui.alert('Error', 'Failed to generate tasks. Please try again later.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 5: Clear existing tasks
    clearExistingTasks();
    
    // Step 6: Add generated tasks to the Task Management sheet
    const taskCount = addTasksToSheet(tasks, eventInfo);
    
    // Step 7: Show success message
    ui.alert('Success', `Generated and added ${taskCount} tasks to the Task Management sheet.`, ui.ButtonSet.OK);
  
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    ui.alert('Error', 'An error occurred: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Sets up the Task Management sheet with headers, formatting, and dropdowns.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object (optional)
 * @param {boolean} addSampleData Whether to add sample data to the sheet (default: false)
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The configured Task Management sheet
 */
function setupTaskManagementSheet(ss, addSampleData = false) {
  // Allow calling with or without spreadsheet parameter
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if the sheet already exists, if so remove it
  let sheet = ss.getSheetByName('Task Management');
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // Create the new sheet
  sheet = ss.insertSheet('Task Management');
  sheet.setTabColor('#9fc5e8'); // Light blue color
  
  // Ensure we have at least 900 rows for future data
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < 900) {
    sheet.insertRowsAfter(currentMaxRows, 900 - currentMaxRows);
  }
  
  // Define headers
  const headers = [
    'Task ID', 
    'Task Name', 
    'Description', 
    'Category', 
    'Owner', 
    'Due Date', 
    'Status', 
    'Priority', 
    'Related Session', 
    'Reminder Sent?'
  ];
  
  // Set header values
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Set column widths
  const widths = [120, 200, 300, 120, 150, 100, 100, 100, 150, 120];
  for (let i = 0; i < headers.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Format headers with blue background and white text
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8') // Blue background
    .setFontColor('#ffffff') // White text
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Get data from Config sheet for dropdown options
  const configData = getConfigDropdownOptions(ss);
  
  // Create dropdown validations
  // Category dropdown (Column 4)
  const categoryOptions = ['Venue', 'Marketing', 'Logistics', 'Program', 'Budget', 'Staffing', 'Technology', 'Communications', 'Other'];
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryOptions, true)
    .build();
  sheet.getRange(2, 4, 899, 1).setDataValidation(categoryRule);
  
  // Owner dropdown from Config sheet (Column 5)
  if (configData.owners && configData.owners.length > 0) {
    const ownerRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(configData.owners, true)
      .build();
    sheet.getRange(2, 5, 899, 1).setDataValidation(ownerRule);
  }
  
  // Format Due Date column (Column 6)
  sheet.getRange(2, 6, 899, 1).setNumberFormat('yyyy-mm-dd');
  
  // Status dropdown from Config sheet (Column 7)
  if (configData.taskStatus && configData.taskStatus.length > 0) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(configData.taskStatus, true)
      .build();
    sheet.getRange(2, 7, 899, 1).setDataValidation(statusRule);
  } else {
    // Default status options if not found in Config
    const defaultStatusOptions = ['Not Started', 'In Progress', 'Done', 'Blocked', 'Cancelled'];
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(defaultStatusOptions, true)
      .build();
    sheet.getRange(2, 7, 899, 1).setDataValidation(statusRule);
  }
  
  // Priority dropdown from Config sheet (Column 8)
  if (configData.taskPriority && configData.taskPriority.length > 0) {
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(configData.taskPriority, true)
      .build();
    sheet.getRange(2, 8, 899, 1).setDataValidation(priorityRule);
  } else {
    // Default priority options if not found in Config
    const defaultPriorityOptions = ['High', 'Medium', 'Low'];
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(defaultPriorityOptions, true)
      .build();
    sheet.getRange(2, 8, 899, 1).setDataValidation(priorityRule);
  }
  
  // Related Session dropdown - will be populated in updateAllDropdowns
  
  // Reminder Sent dropdown (Column 10)
  const reminderOptions = ['Yes', 'No'];
  const reminderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(reminderOptions, true)
    .build();
  sheet.getRange(2, 10, 899, 1).setDataValidation(reminderRule);
  
  // Apply alternating row colors to all data rows
  for (let i = 2; i <= 900; i += 2) {
    sheet.getRange(i, 1, 1, headers.length).setBackground('#f3f3f3');
  }
  
  // Add sample data if requested
  if (addSampleData) {
    const today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate() + 1);
    
    const sampleTasks = [
      [
        'T-001',
        'Create Event Budget',
        'Prepare detailed budget including venue costs, catering, staff, and materials',
        'Budget',
        'John Doe',
        tomorrow,
        'Not Started',
        'High',
        '',
        'No'
      ],
      [
        'T-002',
        'Book Venue',
        'Contact venue options, get quotes, and secure reservation',
        'Venue',
        'Jane Smith',
        tomorrow,
        'In Progress',
        'High',
        '',
        'No'
      ]
    ];
    
    sheet.getRange(2, 1, sampleTasks.length, sampleTasks[0].length).setValues(sampleTasks);
  }
  
  // Create filter view for all rows
  sheet.getRange(1, 1, 900, headers.length).createFilter();
  
  return sheet;
}

/**
 * Helper function to get dropdown options from Config sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @return {Object} Object containing various dropdown options
 */
function getConfigDropdownOptions(ss) {
  const result = {
    owners: [],
    taskStatus: [],
    taskPriority: []
  };
  
  try {
    const configSheet = ss.getSheetByName('Config');
    if (!configSheet) return result;
    
    const data = configSheet.getDataRange().getValues();
    
    // Find specific entries in Config sheet
    for (let i = 0; i < data.length; i++) {
      const key = data[i][0];
      if (!key) continue;
      
      const keyLower = key.toString().toLowerCase();
      const value = data[i][1];
      
      if (keyLower === 'owners' && value) {
        // Owners list
        result.owners = value.toString().split(',').map(item => item.trim());
      } else if (keyLower === 'task status options' && value) {
        // Task Status options
        result.taskStatus = value.toString().split(',').map(item => item.trim());
      } else if (keyLower === 'task priority options' && value) {
        // Task Priority options
        result.taskPriority = value.toString().split(',').map(item => item.trim());
      }
    }
  } catch (error) {
    Logger.log(`Error getting Config options: ${error}`);
  }
  
  return result;
}



/**
 * Gets comprehensive event information from the Event Description sheet
 * @return {Object|null} Object with event details or null if not found
 */
function getEventInformation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName(EVENT_DESC_SHEET_NAME);
  
  if (!eventSheet) {
    Logger.log('Event Description sheet not found');
    return null;
  }
  
  const data = eventSheet.getDataRange().getValues();
  
  // Initialize event info object
  const eventInfo = {
    eventId: '',
    eventName: '',
    startDate: null,
    endDate: null,
    isMultiDay: false,
    location: '',
    theme: '',
    objectives: '',
    description: '',
    detailedDescription: '',
    keyMessages: ''
  };
  
  // Map field names to property names
  const fieldMap = {
    'Event ID': 'eventId',
    'Event Name': 'eventName',
    'Start Date (And Time)': 'startDate',
    'Start Date': 'startDate',
    'End Date (And Time)': 'endDate',
    'End Date': 'endDate',
    'Single- or Multi-Day?': 'isMultiDay',
    'Location': 'location',
    'Theme': 'theme',
    'Short Objectives': 'objectives',
    'Description & Messaging': 'description',
    'Detailed Description': 'detailedDescription',
    'Key Messages': 'keyMessages'
  };
  
  // Extract data from sheet
  for (let i = 0; i < data.length; i++) {
    const fieldName = data[i][0];
    const fieldValue = data[i][1];
    
    const propertyName = fieldMap[fieldName];
    if (propertyName) {
      // Parse dates for date fields
      if (propertyName === 'startDate' || propertyName === 'endDate') {
        if (fieldValue instanceof Date) {
          eventInfo[propertyName] = fieldValue;
        } else if (typeof fieldValue === 'string') {
          // Try to parse string as date
          try {
            const parsedDate = new Date(fieldValue);
            if (!isNaN(parsedDate.getTime())) {
              eventInfo[propertyName] = parsedDate;
            }
          } catch (e) {
            Logger.log(`Error parsing date: ${fieldValue}`);
          }
        }
      } 
      // Parse multi-day flag
      else if (propertyName === 'isMultiDay') {
        eventInfo[propertyName] = fieldValue === 'Multi';
      }
      // Set other properties directly
      else {
        eventInfo[propertyName] = fieldValue;
      }
    }
  }
  
  // Validate required fields
  if (!eventInfo.eventName || !eventInfo.startDate) {
    Logger.log('Missing required event information');
    return null;
  }
  
  // Set end date equal to start date if not specified (single-day event)
  if (!eventInfo.endDate) {
    eventInfo.endDate = new Date(eventInfo.startDate);
  }
  
  // Calculate event duration in days
  const eventDuration = Math.round(
    (eventInfo.endDate.getTime() - eventInfo.startDate.getTime()) / (1000 * 60 * 60 * 24)
  ) + 1; // +1 to include both start and end days
  
  eventInfo.durationDays = eventDuration;
  
  // Log the event info for debugging
  Logger.log(`Event: ${eventInfo.eventName}`);
  Logger.log(`Dates: ${eventInfo.startDate} to ${eventInfo.endDate} (${eventInfo.durationDays} days)`);
  
  return eventInfo;
}

/**
 * Gets the OpenAI API key from script properties or Config sheet
 * @return {string|null} API key or null if not found
 */
function getOpenAIApiKey() {
  try {
    // Try script properties first
    const scriptProperties = PropertiesService.getScriptProperties();
    let apiKey = scriptProperties.getProperty('OPENAI_API_KEY');
    
    if (apiKey) {
      Logger.log('Retrieved API key from script properties');
      return apiKey;
    }
    
    // Try Config sheet next
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (configSheet) {
      // Try direct access to B15
      apiKey = configSheet.getRange("B15").getValue();
      
      if (!apiKey) {
        // Try searching for the row with "OpenAI API Key"
        const keyRange = configSheet.getRange("A:A");
        const keyValues = keyRange.getValues();
        
        for (let i = 0; i < keyValues.length; i++) {
          if (keyValues[i][0] === "OpenAI API Key") {
            apiKey = configSheet.getRange(i + 1, 2).getValue();
            break;
          }
        }
      }
    }
    
    return apiKey || null;
  } catch (e) {
    Logger.log('Error getting API key: ' + e.toString());
    return null;
  }
}

/**
 * Saves API key to script properties for secure storage
 */
function saveApiKeyToScriptProperties() {
  const ui = SpreadsheetApp.getUi();
  
  // Option 1: Get key from Config sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  let apiKey = null;
  
  if (configSheet) {
    apiKey = configSheet.getRange("B15").getValue();
  }
  
  // If not found in Config sheet, ask user directly
  if (!apiKey) {
    const response = ui.prompt(
      'OpenAI API Key',
      'Please enter your OpenAI API key:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      apiKey = response.getResponseText().trim();
    } else {
      return; // User cancelled
    }
  }
  
  // Save to script properties
  if (apiKey) {
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
    ui.alert('Success', 'API key saved to script properties successfully!', ui.ButtonSet.OK);
  } else {
    ui.alert('Error', 'No API key provided.', ui.ButtonSet.OK);
  }
}

/**
 * Tests if the OpenAI API key is working
 */
function testApiKey() {
  const ui = SpreadsheetApp.getUi();
  const apiKey = getOpenAIApiKey();
  
  if (!apiKey) {
    ui.alert('Error', 'No API key found. Please add your API key first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Make a simple API call
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: "Hello" }],
      max_tokens: 10
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      ui.alert('Success', 'API key is working correctly!', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `API returned error code: ${responseCode}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Failed to test API key: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Generate tasks using OpenAI API based on event information
 * @param {Object} eventInfo Event information
 * @param {string} apiKey OpenAI API key
 * @return {Array} Array of task objects
 */
function generateTasksWithAI(eventInfo, apiKey) {
  try {
    // Format dates for the prompt
    const startDate = formatDate(eventInfo.startDate);
    const endDate = formatDate(eventInfo.endDate);
    
    // Create a detailed prompt with all available event information
    const prompt = `
Generate a comprehensive list of tasks for planning and executing the following event:

EVENT DETAILS:
- Name: ${eventInfo.eventName}
- Date(s): ${startDate}${eventInfo.durationDays > 1 ? ` to ${endDate}` : ''}
- Duration: ${eventInfo.durationDays} day(s)
- Location: ${eventInfo.location || 'TBD'}
- Theme: ${eventInfo.theme || 'N/A'}
- Objectives: ${eventInfo.objectives || 'N/A'}
- Description: ${eventInfo.description || 'N/A'}
- Detailed Description: ${eventInfo.detailedDescription || 'N/A'}
- Key Messages: ${eventInfo.keyMessages || 'N/A'}

Please create a comprehensive task list divided into three phases:
1. PRE-EVENT TASKS (tasks to complete before ${startDate}, with most due ${DEFAULT_PRE_EVENT_DAYS} days before)
2. DURING-EVENT TASKS (tasks to complete during the event, between ${startDate} and ${endDate})
3. POST-EVENT TASKS (tasks to complete after ${endDate}, with most due ${DEFAULT_POST_EVENT_DAYS} days after)

For each task, include:
- Task Name: A concise, clear description of the task
- Description: A detailed explanation of what needs to be done
- Category: One of the following: Venue, Marketing, Logistics, Program, Budget, Staffing, Technology, Communications, Other
- Priority: Critical, High, Medium, or Low
- Timeline: When this task should be completed (e.g., "4 days before event," "During event, day 1," "2 days after event")
- Status: Should always be "Not Started"

Return your response in this exact JSON format:
{
  "tasks": [
    {
      "name": "Task Name",
      "description": "Detailed task description",
      "category": "Category",
      "priority": "Priority",
      "timeline": "Timeline",
      "status": "Not Started"
    },
    ...more tasks...
  ]
}

IMPORTANT: Include a diverse range of tasks covering all necessary aspects of event planning and execution based on the provided event details.
`;

    // Call OpenAI API
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: "You are an expert event planning assistant that creates detailed task lists for events."
        },
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: 0.7,
      max_tokens: 4000
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    Logger.log('Calling OpenAI API...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`API Error (${responseCode}): ${response.getContentText()}`);
      throw new Error(`OpenAI API returned error code: ${responseCode}`);
    }
    
    const responseJson = JSON.parse(response.getContentText());
    const content = responseJson.choices[0].message.content;
    
    Logger.log(`Received response (${content.length} chars)`);
    
    // Parse the JSON response
    const tasks = parseTasksFromResponse(content);
    
    Logger.log(`Successfully extracted ${tasks.length} tasks`);
    return tasks;
    
  } catch (error) {
    Logger.log('Error generating tasks: ' + error.toString());
    throw error;
  }
}

/**
 * Parse tasks from the OpenAI response
 * @param {string} response The raw text response from OpenAI
 * @return {Array} Array of task objects
 */
function parseTasksFromResponse(response) {
  try {
    // Try to find and parse JSON in the response
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    
    if (jsonMatch) {
      const jsonString = jsonMatch[0];
      
      // Try to repair potentially truncated JSON
      let repairedJson = jsonString;
      
      // Check if JSON is truncated
      if (repairedJson.includes('"tasks": [') && 
          (!repairedJson.endsWith('}') || !repairedJson.includes(']}')) || 
          repairedJson.endsWith(',')) {
        
        // Remove trailing commas
        repairedJson = repairedJson.replace(/,\s*$/, '');
        
        // Add missing brackets/braces
        const openBrackets = (repairedJson.match(/\[/g) || []).length;
        const closeBrackets = (repairedJson.match(/\]/g) || []).length;
        for (let i = 0; i < openBrackets - closeBrackets; i++) {
          repairedJson += ']';
        }
        
        const openBraces = (repairedJson.match(/\{/g) || []).length;
        const closeBraces = (repairedJson.match(/\}/g) || []).length;
        for (let i = 0; i < openBraces - closeBraces; i++) {
          repairedJson += '}';
        }
        
        Logger.log('Repaired potentially truncated JSON');
      }
      
      // Parse the JSON
      const parsedData = JSON.parse(repairedJson);
      
      if (parsedData && parsedData.tasks && Array.isArray(parsedData.tasks)) {
        return parsedData.tasks;
      }
    }
    
    // If JSON parsing fails, try extracting tasks directly
    const tasks = extractTasksDirectly(response);
    if (tasks.length > 0) {
      return tasks;
    }
    
    // If all else fails, return empty array
    Logger.log('Failed to parse tasks from response');
    return [];
    
  } catch (error) {
    Logger.log('Error parsing tasks: ' + error.toString());
    // Try direct extraction as fallback
    return extractTasksDirectly(response);
  }
}

/**
 * Extract tasks directly from response text using regex as a fallback
 * @param {string} response The OpenAI response text
 * @return {Array} Array of task objects
 */
function extractTasksDirectly(response) {
  try {
    const tasks = [];
    
    // Extract individual task components
    const namePattern = /"name":\s*"([^"]*)"/g;
    const descPattern = /"description":\s*"([^"]*)"/g;
    const catPattern = /"category":\s*"([^"]*)"/g;
    const priorityPattern = /"priority":\s*"([^"]*)"/g;
    const timelinePattern = /"timeline":\s*"([^"]*)"/g;
    
    const names = extractMatches(response, namePattern);
    const descriptions = extractMatches(response, descPattern);
    const categories = extractMatches(response, catPattern);
    const priorities = extractMatches(response, priorityPattern);
    const timelines = extractMatches(response, timelinePattern);
    
    // Find how many complete tasks we can create
    const count = Math.min(
      names.length,
      descriptions.length,
      categories.length,
      priorities.length,
      timelines.length
    );
    
    // Create task objects
    for (let i = 0; i < count; i++) {
      tasks.push({
        name: names[i],
        description: descriptions[i],
        category: categories[i],
        priority: priorities[i],
        timeline: timelines[i],
        status: "Not Started"
      });
    }
    
    Logger.log(`Extracted ${tasks.length} tasks directly`);
    return tasks;
    
  } catch (error) {
    Logger.log('Error in direct extraction: ' + error.toString());
    return [];
  }
}

/**
 * Helper function to extract all regex matches
 * @param {string} text The text to search
 * @param {RegExp} pattern The regex pattern with capture group
 * @return {Array} Array of matched strings
 */
function extractMatches(text, pattern) {
  const results = [];
  let match;
  
  // Create a new regex to reset the lastIndex
  const regex = new RegExp(pattern.source, pattern.flags);
  
  while ((match = regex.exec(text)) !== null) {
    if (match[1]) {
      results.push(match[1]);
    }
  }
  
  return results;
}

/**
 * Clears existing tasks from the Task sheet
 */
function clearExistingTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  
  if (!taskSheet) {
    Logger.log('Task sheet not found');
    return;
  }
  
  const lastRow = taskSheet.getLastRow();
  
  // If sheet is empty or only has header row, nothing to clear
  if (lastRow <= 1) {
    return;
  }
  
  // Clear all data rows (preserving header row)
  taskSheet.getRange(2, 1, lastRow - 1, taskSheet.getLastColumn()).clear();
  
  // Reapply alternating row colors
  for (let i = 2; i <= 900; i += 2) {
    taskSheet.getRange(i, 1, 1, taskSheet.getLastColumn()).setBackground('#f3f3f3');
  }
  
  Logger.log('Cleared existing tasks');
}

/**
 * Adds tasks to the Task Management sheet
 * @param {Array} tasks Array of task objects
 * @param {Object} eventInfo Event information for calculating due dates
 * @return {number} Number of tasks added
 */
function addTasksToSheet(tasks, eventInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  
  if (!taskSheet) {
    Logger.log('Task sheet not found');
    throw new Error('Task Management sheet not found');
  }
  
  // Get the starting row
  const startRow = Math.max(2, taskSheet.getLastRow() + 1);
  
  // Create header row if needed
  if (startRow === 2) {
    const headers = [
      'Task ID', 'Task Name', 'Description', 'Category', 
      'Owner', 'Due Date', 'Status', 'Priority',
      'Related Session', 'Reminder Sent?'
    ];
    
    taskSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    taskSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4a86e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  // Sort tasks by category and priority
  const sortedTasks = sortTasks(tasks);
  
  // Create task data for batch insertion
  const taskData = [];
  
  for (const task of sortedTasks) {
    // Generate a unique task ID
    const taskId = generateTaskId();
    
    // Calculate due date based on timeline and event info
    const dueDate = calculateDueDate(task.timeline, eventInfo);
    
    // Create a row for this task
    taskData.push([
      taskId,                // Task ID
      task.name,             // Task Name
      task.description,      // Description
      task.category,         // Category
      '',                    // Owner (left blank)
      dueDate,               // Due Date
      task.status || 'Not Started', // Status
      task.priority,         // Priority
      '',                    // Related Session (left blank)
      'No'                   // Reminder Sent?
    ]);
  }
  
  // Insert all tasks in one batch operation
  if (taskData.length > 0) {
    taskSheet.getRange(startRow, 1, taskData.length, taskData[0].length).setValues(taskData);
    
    // Format the date column
    taskSheet.getRange(startRow, 6, taskData.length).setNumberFormat('yyyy-mm-dd');
    
    // Apply alternating row colors for readability
    for (let i = 0; i < taskData.length; i++) {
      const row = startRow + i;
      if (row % 2 === 0) {
        taskSheet.getRange(row, 1, 1, taskData[0].length).setBackground('#f3f3f3');
      }
    }
  }
  
  return taskData.length;
}

/**
 * Sorts tasks by category and priority for a better organization
 * @param {Array} tasks Array of task objects
 * @return {Array} Sorted array of tasks
 */
function sortTasks(tasks) {
  // Define category and priority order
  const categoryOrder = {
    'Venue': 1,
    'Marketing': 2,
    'Logistics': 3,
    'Program': 4,
    'Budget': 5,
    'Staffing': 6,
    'Technology': 7,
    'Communications': 8,
    'Other': 9
  };
  
  const priorityOrder = {
    'Critical': 1,
    'High': 2,
    'Medium': 3,
    'Low': 4
  };
  
  // Sort based on category first, then priority
  return [...tasks].sort((a, b) => {
    // Sort by category
    const catOrderA = categoryOrder[a.category] || 999;
    const catOrderB = categoryOrder[b.category] || 999;
    
    if (catOrderA !== catOrderB) {
      return catOrderA - catOrderB;
    }
    
    // If same category, sort by priority
    const priorityOrderA = priorityOrder[a.priority] || 999;
    const priorityOrderB = priorityOrder[b.priority] || 999;
    
    return priorityOrderA - priorityOrderB;
  });
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
 * Calculates a due date based on a timeline description and event info
 * @param {string} timeline Timeline description (e.g., "4 days before event")
 * @param {Object} eventInfo Event information with dates
 * @return {Date} Calculated due date
 */
function calculateDueDate(timeline, eventInfo) {
  if (!timeline || !eventInfo || !eventInfo.startDate) {
    // Default to 4 days before event
    return addDays(eventInfo.startDate, -DEFAULT_PRE_EVENT_DAYS);
  }
  
  const timelineStr = timeline.toLowerCase();
  
  // Pre-event tasks
  if (timelineStr.includes('before') || timelineStr.includes('prior')) {
    // Check for specific number of days
    const dayMatch = timelineStr.match(/(\d+)\s+days?/);
    if (dayMatch) {
      const days = parseInt(dayMatch[1]);
      return addDays(eventInfo.startDate, -days);
    }
    
    // Check for "week before" patterns
    if (timelineStr.includes('week before') || timelineStr.includes('1 week before')) {
      return addDays(eventInfo.startDate, -7);
    } else if (timelineStr.includes('2 weeks before')) {
      return addDays(eventInfo.startDate, -14);
    } else if (timelineStr.includes('month before')) {
      return addDays(eventInfo.startDate, -30);
    }
    
    // Default pre-event due date
    return addDays(eventInfo.startDate, -DEFAULT_PRE_EVENT_DAYS);
  }
  
  // During-event tasks
  if (timelineStr.includes('during') || timelineStr.includes('day of')) {
    // Check for specific day number
    const dayMatch = timelineStr.match(/day\s*(\d+)/i);
    if (dayMatch && eventInfo.durationDays > 1) {
      const dayNum = parseInt(dayMatch[1]) - 1; // 0-based day index
      const maxOffset = eventInfo.durationDays - 1; // Don't go beyond end date
      return addDays(eventInfo.startDate, Math.min(dayNum, maxOffset));
    }
    
    // Default to first day of event
    return new Date(eventInfo.startDate);
  }
  
  // Post-event tasks
  if (timelineStr.includes('after') || timelineStr.includes('following')) {
    // Check for specific number of days
    const dayMatch = timelineStr.match(/(\d+)\s+days?/);
    if (dayMatch) {
      const days = parseInt(dayMatch[1]);
      return addDays(eventInfo.endDate, days);
    }
    
    // Default post-event due date
    return addDays(eventInfo.endDate, DEFAULT_POST_EVENT_DAYS);
  }
  
  // Default fallback
  return addDays(eventInfo.startDate, -DEFAULT_PRE_EVENT_DAYS);
}

/**
 * Adds days to a date (negative values subtract days)
 * @param {Date} date The base date
 * @param {number} days Number of days to add/subtract
 * @return {Date} New date
 */
function addDays(date, days) {
  if (!date || !(date instanceof Date)) {
    return new Date(); // Default to today if invalid date
  }
  
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * Formats a date as YYYY-MM-DD
 * @param {Date} date The date to format
 * @return {string} Formatted date string
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) {
    return '';
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Add menu items to the spreadsheet menu
 * This should be called from your onOpen function
 * @param {Object} menu The menu object to add items to
 */
function addTaskMenuItems(menu) {
  menu.addItem('Generate AI Tasks', 'generateAITasks')
      .addItem('Save API Key to Script Properties', 'saveApiKeyToScriptProperties')
      .addItem('Test API Key', 'testApiKey');
  // Diagnostic button removed as requested
}