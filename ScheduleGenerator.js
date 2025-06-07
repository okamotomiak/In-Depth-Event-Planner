// ScheduleGenerator.gs - Automatically generate a preliminary schedule

/**
 * Adds the "Generate Preliminary Schedule" menu item to the Event Planner menu
 * This function should be called from onOpen() in Core.gs
 */
function addScheduleGeneratorMenuItem(menu) {
  return menu.addItem('Generate Preliminary Schedule', 'generatePreliminarySchedule');
}

/**
 * Main function to generate a preliminary schedule using OpenAI
 * Modified to enforce approved locations
 */
function generatePreliminarySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Check if required sheets exist
  const eventDescSheet = ss.getSheetByName('Event Description');
  const peopleSheet = ss.getSheetByName('People');
  const scheduleSheet = ss.getSheetByName('Schedule');
  
  if (!eventDescSheet || !peopleSheet || !scheduleSheet) {
    ui.alert('Error', 'Required sheets (Event Description, People, Schedule) not found. Please make sure all sheets are set up.', ui.ButtonSet.OK);
    return;
  }
  
  // Step 1: Gather event details
  const eventDetails = getEventDetails(eventDescSheet);
  if (!eventDetails || !eventDetails.eventName || !eventDetails.startDate || !eventDetails.endDate) {
    ui.alert('Error', 'Required event details (Event Name, Start Date, End Date) are missing. Please complete the Event Description sheet.', ui.ButtonSet.OK);
    return;
  }
  
  // Step 2: Get speakers
  const speakers = getSpeakersFromPeopleSheet(peopleSheet);
  
  // Step 3: Get approved locations from Config sheet
  const approvedLocations = getApprovedLocationList(ss);
  
  // Step 4: Get OpenAI API key
  const apiKey = getOpenAIApiKey();
  if (!apiKey) {
    ui.alert('Error', 'OpenAI API key not found. Please add your API key to the Config sheet.', ui.ButtonSet.OK);
    return;
  }
  
  // Show loading message
  ui.alert('Processing', 'Generating preliminary schedule. This may take a few moments...', ui.ButtonSet.OK);
  
  try {
    // Step 5: Generate prompt and call OpenAI
    const prompt = generatePrompt(eventDetails, speakers, approvedLocations);
    const scheduleData = callOpenAIForSchedule(prompt, apiKey, eventDetails, approvedLocations);
    
    if (!scheduleData || scheduleData.length === 0) {
      ui.alert('Error', 'Failed to generate schedule or no schedule was returned from OpenAI.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 6: Populate Schedule sheet
    const scheduleCount = populateScheduleSheet(scheduleData, eventDetails, scheduleSheet, approvedLocations);
    
    // Show success message
    ui.alert('Success', `${scheduleCount} schedule items have been added to the Schedule sheet.`, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log('Error generating preliminary schedule: ' + error.toString());
    ui.alert('Error', 'An error occurred while generating the schedule: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Gets the approved location list from the Config sheet
 * These are the only locations that should be used for sessions
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @return {Array} Array of approved location strings
 */
function getApprovedLocationList(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    Logger.log('Config sheet not found, using default locations');
    return ['Main Hall']; // Default fallback
  }
  
  // Find the Location List row in the Config sheet
  const data = configSheet.getDataRange().getValues();
  let locationList = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'Location List') {
      // Location list is in column 2 (B) as a comma-separated string
      const locationString = data[i][1];
      if (locationString) {
        locationList = locationString.split(',').map(loc => loc.trim());
      }
      break;
    }
  }
  
  if (locationList.length === 0) {
    Logger.log('No locations found in Config sheet, using default locations');
    return ['Main Hall']; // Default fallback
  }
  
  Logger.log('Found approved locations: ' + locationList.join(', '));
  return locationList;
}

/**
 * Validates and corrects a location to ensure it's from the approved list
 * @param {string} location The location to validate
 * @param {Array} approvedLocations Array of approved location strings
 * @return {string} A valid location from the approved list
 */
function validateLocation(location, approvedLocations) {
  if (!location || !approvedLocations || approvedLocations.length === 0) {
    return "Main Hall"; // Default fallback
  }
  
  // Check if the location exactly matches an approved location
  const exactMatch = approvedLocations.find(loc => 
    loc.toLowerCase() === location.toLowerCase()
  );
  
  if (exactMatch) {
    return exactMatch; // Return the exact casing from the approved list
  }
  
  // If no exact match, look for a partial match
  for (const approved of approvedLocations) {
    if (location.toLowerCase().includes(approved.toLowerCase()) || 
        approved.toLowerCase().includes(location.toLowerCase())) {
      return approved; // Return the approved location with correct casing
    }
  }
  
  // If no match found, return the first approved location as default
  return approvedLocations[0];
}

/**
 * Extracts event details from the Event Description sheet
 * Enhanced to improve date/time parsing and ensure Description fields are properly captured
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Event Description sheet
 * @return {Object} Event details including name, dates, theme, audience, objectives, descriptions, etc.
 */
function getEventDetails(sheet) {
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Initialize object to store event details
  const eventDetails = {
    eventName: null,
    startDate: null,
    endDate: null,
    startTime: null,
    endTime: null,
    theme: null,
    targetAudience: null,
    objectives: null,
    description: null,
    detailedDescription: null,
    keyMessages: null
  };
  
  // Find the values using _findRow helper from Core.gs where possible
  const eventNameRow = _findRow(sheet, 'Event Name');
  if (eventNameRow) {
    eventDetails.eventName = sheet.getRange(eventNameRow, 2).getValue();
  }
  
  // Look for Start Date with different possible labels
  const startDateRow = _findRow(sheet, 'Start Date (And Time)') || _findRow(sheet, 'Start Date');
  if (startDateRow) {
    const startDateCell = sheet.getRange(startDateRow, 2);
    const startDateValue = startDateCell.getValue();
    const startDateFormat = startDateCell.getNumberFormat();
    
    // If the cell contains a date value
    if (startDateValue instanceof Date) {
      eventDetails.startDate = startDateValue;
      
      // Extract time component if it exists
      const hours = startDateValue.getHours();
      const minutes = startDateValue.getMinutes();
      if (hours !== 0 || minutes !== 0) {
        // Format as 24-hour time
        eventDetails.startTime = hours + ':' + (minutes < 10 ? '0' + minutes : minutes);
        
        // Also store as formatted 12-hour time for prompt readability
        const ampm = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours % 12 || 12; // Convert to 12-hour format
        eventDetails.startTimeFormatted = displayHours + ':' + 
                                         (minutes < 10 ? '0' + minutes : minutes) + 
                                         ' ' + ampm;
      }
    } else if (typeof startDateValue === 'string') {
      // Try to parse date from string format
      try {
        // Handle strings like "May 24, 2023, 9:00 AM"
        const parsedDate = new Date(startDateValue);
        if (!isNaN(parsedDate.getTime())) {
          eventDetails.startDate = parsedDate;
          
          // Extract time if present in the string
          const timeMatch = startDateValue.match(/(\d{1,2}):(\d{2})\s*(am|pm|AM|PM)/i);
          if (timeMatch) {
            let hours = parseInt(timeMatch[1]);
            const minutes = parseInt(timeMatch[2]);
            const ampm = timeMatch[3].toUpperCase();
            
            // Convert to 24-hour format for internal use
            if (ampm === 'PM' && hours < 12) hours += 12;
            if (ampm === 'AM' && hours === 12) hours = 0;
            
            eventDetails.startTime = hours + ':' + (minutes < 10 ? '0' + minutes : minutes);
            eventDetails.startTimeFormatted = timeMatch[1] + ':' + timeMatch[2] + ' ' + ampm;
          }
        }
      } catch (e) {
        Logger.log('Error parsing start date string: ' + e.toString());
      }
    }
    
    // Log for debugging
    Logger.log('Start Date Cell Value: ' + startDateValue);
    Logger.log('Start Date Cell Format: ' + startDateFormat);
    Logger.log('Parsed Start Date: ' + eventDetails.startDate);
    Logger.log('Parsed Start Time: ' + eventDetails.startTimeFormatted);
  }
  
  // Look for End Date with different possible labels
  const endDateRow = _findRow(sheet, 'End Date (And Time)') || _findRow(sheet, 'End Date');
  if (endDateRow) {
    const endDateCell = sheet.getRange(endDateRow, 2);
    const endDateValue = endDateCell.getValue();
    const endDateFormat = endDateCell.getNumberFormat();
    
    // If the cell contains a date value
    if (endDateValue instanceof Date) {
      eventDetails.endDate = endDateValue;
      
      // Extract time component if it exists
      const hours = endDateValue.getHours();
      const minutes = endDateValue.getMinutes();
      if (hours !== 0 || minutes !== 0) {
        // Format as 24-hour time
        eventDetails.endTime = hours + ':' + (minutes < 10 ? '0' + minutes : minutes);
        
        // Also store as formatted 12-hour time for prompt readability
        const ampm = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours % 12 || 12; // Convert to 12-hour format
        eventDetails.endTimeFormatted = displayHours + ':' + 
                                       (minutes < 10 ? '0' + minutes : minutes) + 
                                       ' ' + ampm;
      }
    } else if (typeof endDateValue === 'string') {
      // Try to parse date from string format
      try {
        // Handle strings like "May 24, 2023, 3:00 PM"
        const parsedDate = new Date(endDateValue);
        if (!isNaN(parsedDate.getTime())) {
          eventDetails.endDate = parsedDate;
          
          // Extract time if present in the string
          const timeMatch = endDateValue.match(/(\d{1,2}):(\d{2})\s*(am|pm|AM|PM)/i);
          if (timeMatch) {
            let hours = parseInt(timeMatch[1]);
            const minutes = parseInt(timeMatch[2]);
            const ampm = timeMatch[3].toUpperCase();
            
            // Convert to 24-hour format for internal use
            if (ampm === 'PM' && hours < 12) hours += 12;
            if (ampm === 'AM' && hours === 12) hours = 0;
            
            eventDetails.endTime = hours + ':' + (minutes < 10 ? '0' + minutes : minutes);
            eventDetails.endTimeFormatted = timeMatch[1] + ':' + timeMatch[2] + ' ' + ampm;
          }
        }
      } catch (e) {
        Logger.log('Error parsing end date string: ' + e.toString());
      }
    }
    
    // Log for debugging
    Logger.log('End Date Cell Value: ' + endDateValue);
    Logger.log('End Date Cell Format: ' + endDateFormat);
    Logger.log('Parsed End Date: ' + eventDetails.endDate);
    Logger.log('Parsed End Time: ' + eventDetails.endTimeFormatted);
  }
  
  // If end date is not found, use start date (for single-day events)
  if (!eventDetails.endDate && eventDetails.startDate) {
    eventDetails.endDate = new Date(eventDetails.startDate);
    
    // If there's an end time but no end date, use start date with end time
    if (eventDetails.endTime && !eventDetails.endDate) {
      const [hours, minutes] = eventDetails.endTime.split(':').map(num => parseInt(num, 10));
      eventDetails.endDate = new Date(eventDetails.startDate);
      eventDetails.endDate.setHours(hours, minutes, 0, 0);
    }
  }
  
  // Look for Theme, which might be called different things
  const themeRow = _findRow(sheet, 'Theme') || _findRow(sheet, 'Event Theme') || _findRow(sheet, 'Topics') || _findRow(sheet, 'Focus');
  if (themeRow) {
    eventDetails.theme = sheet.getRange(themeRow, 2).getValue();
  }
  
  // Look for Target Audience with different possible labels
  const audienceRow = _findRow(sheet, 'Target Audience') || _findRow(sheet, 'Audience') || _findRow(sheet, 'Attendees');
  if (audienceRow) {
    eventDetails.targetAudience = sheet.getRange(audienceRow, 2).getValue();
  }
  
  // Look for Objectives
  const objectivesRow = _findRow(sheet, 'Objectives');
  if (objectivesRow) {
    eventDetails.objectives = sheet.getRange(objectivesRow, 2).getValue();
  }
  
  // Look for various description fields - ensuring we capture all of them
  const descriptionRow = _findRow(sheet, 'Description & Messaging') || _findRow(sheet, 'Description');
  if (descriptionRow) {
    eventDetails.description = sheet.getRange(descriptionRow, 2).getValue();
    Logger.log('Found Description & Messaging: ' + (eventDetails.description ? 'Yes' : 'No'));
  }
  
  const detailedDescRow = _findRow(sheet, 'Detailed Description');
  if (detailedDescRow) {
    eventDetails.detailedDescription = sheet.getRange(detailedDescRow, 2).getValue();
    Logger.log('Found Detailed Description: ' + (eventDetails.detailedDescription ? 'Yes' : 'No'));
  }
  
  const keyMessagesRow = _findRow(sheet, 'Key Messages');
  if (keyMessagesRow) {
    eventDetails.keyMessages = sheet.getRange(keyMessagesRow, 2).getValue();
    Logger.log('Found Key Messages: ' + (eventDetails.keyMessages ? 'Yes' : 'No'));
  }
  
  // Calculate event duration in days
  if (eventDetails.startDate instanceof Date && eventDetails.endDate instanceof Date) {
    // Set both dates to midnight for accurate day calculation
    const startDay = new Date(eventDetails.startDate);
    startDay.setHours(0, 0, 0, 0);
    const endDay = new Date(eventDetails.endDate);
    endDay.setHours(0, 0, 0, 0);
    
    const oneDay = 24 * 60 * 60 * 1000; // milliseconds in a day
    const diffDays = Math.round(Math.abs((endDay - startDay) / oneDay)) + 1;
    eventDetails.durationDays = diffDays;
  } else {
    eventDetails.durationDays = 1; // Default to 1 day if dates are not available
  }
  
  // Default start and end times if not specified
  if (!eventDetails.startTime) {
    eventDetails.startTime = '9:00';
    eventDetails.startTimeFormatted = '9:00 AM';
  }
  
  if (!eventDetails.endTime) {
    eventDetails.endTime = '17:00';
    eventDetails.endTimeFormatted = '5:00 PM';
  }
  
  // Log what we found for debugging
  Logger.log('Event Details Found:');
  Logger.log(`- Event Name: ${eventDetails.eventName}`);
  Logger.log(`- Start Date: ${eventDetails.startDate}`);
  Logger.log(`- Start Time: ${eventDetails.startTimeFormatted || 'Not specified'}`);
  Logger.log(`- End Date: ${eventDetails.endDate}`);
  Logger.log(`- End Time: ${eventDetails.endTimeFormatted || 'Not specified'}`);
  Logger.log(`- Duration: ${eventDetails.durationDays} day(s)`);
  Logger.log(`- Theme: ${eventDetails.theme}`);
  Logger.log(`- Target Audience: ${eventDetails.targetAudience}`);
  Logger.log(`- Objectives: ${eventDetails.objectives}`);
  Logger.log(`- Description: ${eventDetails.description ? 'Found' : 'Not found'}`);
  Logger.log(`- Detailed Description: ${eventDetails.detailedDescription ? 'Found' : 'Not found'}`);
  Logger.log(`- Key Messages: ${eventDetails.keyMessages ? 'Found' : 'Not found'}`);
  
  return eventDetails;
}

/**
 * Gets speakers from the People sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The People sheet
 * @return {Array} List of speaker names
 */
function getSpeakersFromPeopleSheet(sheet) {
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  const speakers = [];
  
  // Get the column indices from the header row
  const headers = data[0].map(header => header.toString().toLowerCase());
  const nameIndex = headers.indexOf('name');
  const categoryIndex = headers.indexOf('category');
  
  // If the required columns are not found, return empty array
  if (nameIndex === -1 || categoryIndex === -1) {
    Logger.log('Required columns not found in People sheet.');
    return speakers;
  }
  
  // Look for people with category 'Speaker'
  for (let i = 1; i < data.length; i++) {
    const category = data[i][categoryIndex];
    if (category && category.toString().toLowerCase() === 'speaker') {
      const name = data[i][nameIndex];
      if (name) {
        speakers.push(name.toString());
      }
    }
  }
  
  Logger.log(`Found ${speakers.length} speakers: ${speakers.join(', ')}`);
  return speakers;
}

/**
 * Generates a prompt for OpenAI based on event details and speakers
 * Now includes specifically approved locations and instructs not to include lead/speaker
 * @param {Object} eventDetails - Event details
 * @param {Array} speakers - List of speaker names
 * @param {Array} approvedLocations - List of approved locations from Config
 * @return {string} Prompt for OpenAI
 */
function generatePrompt(eventDetails, speakers, approvedLocations) {
  // Determine the event type description based on duration
  let eventTypeDesc = "one-day";
  if (eventDetails.durationDays > 1) {
    eventTypeDesc = `${eventDetails.durationDays}-day`;
  }
  
  // Format dates for clarity
  const formattedStartDate = formatDate(eventDetails.startDate);
  const formattedEndDate = formatDate(eventDetails.endDate);
  
  // Format approved locations for inclusion in prompt
  const locationsText = approvedLocations.join(', ');
  
  // Base prompt with explicit event details including exact dates
  let prompt = `
Draft a detailed sample schedule outline for a ${eventTypeDesc} event titled '${eventDetails.eventName}'`;
  
  // Add theme if available
  if (eventDetails.theme) {
    prompt += ` focused on '${eventDetails.theme}'`;
  }
  
  // Add target audience if available
  if (eventDetails.targetAudience) {
    prompt += ` for ${eventDetails.targetAudience}`;
  }
  
  // Add date specificity
  prompt += ` taking place on ${formattedStartDate}`;
  if (eventDetails.durationDays > 1) {
    prompt += ` through ${formattedEndDate}`;
  }
  prompt += `.

IMPORTANT TIME CONSTRAINTS:
- The event will START at exactly ${eventDetails.startTimeFormatted} each day
- The event will END at exactly ${eventDetails.endTimeFormatted} each day
- NO sessions should be scheduled before ${eventDetails.startTimeFormatted} or after ${eventDetails.endTimeFormatted}

For this ${eventTypeDesc} event, please include:
1. Realistic timing for each session (start and end times) with appropriate breaks
2. Appropriate session titles that reflect the event context
3. Location for each session MUST be one of these exact options: ${locationsText}
4. IMPORTANT: Leave the Lead/Speaker field EMPTY for all sessions - this will be assigned later`;
  
  // Add objectives if available
  if (eventDetails.objectives) {
    prompt += `
5. Sessions that align with these objectives: ${eventDetails.objectives}`;
  }
  
  // Add description context - ensure all description fields are included
  // and prominently featured in the prompt
  let descriptionAdded = false;
  let descriptionText = "\nEVENT CONTEXT (IMPORTANT - USE THIS TO CREATE RELEVANT SESSIONS):\n";
  
  if (eventDetails.description) {
    descriptionText += `- Description & Messaging: ${eventDetails.description}\n\n`;
    descriptionAdded = true;
  }
  
  if (eventDetails.detailedDescription) {
    descriptionText += `- Detailed Description: ${eventDetails.detailedDescription}\n\n`;
    descriptionAdded = true;
  }
  
  if (eventDetails.keyMessages) {
    descriptionText += `- Key Messages: ${eventDetails.keyMessages}\n\n`;
    descriptionAdded = true;
  }
  
  if (descriptionAdded) {
    prompt += `${descriptionText}CREATE SESSION TITLES AND CONTENT THAT DIRECTLY ALIGN WITH THIS EVENT CONTEXT.`;
  }
  
  // Add speakers if available, but clarify they should not be assigned yet
  if (speakers && speakers.length > 0) {
    prompt += `
6. Note: These are the available speakers, but DO NOT assign them to sessions in your response. Leave the speaker/lead field empty: ${speakers.join(', ')}`;
  }
  
  // Add specific format requirements for easier parsing
  prompt += `

Format your response as JSON with the following structure for each session:
{
  "schedule": [
    {
      "date": "YYYY-MM-DD",
      "startTime": "HH:MM AM/PM",
      "endTime": "HH:MM AM/PM",
      "title": "Session Title",
      "speaker": "",  // LEAVE THIS EMPTY
      "location": "Session Location (MUST be one of: ${locationsText})",
      "status": "Tentative"
    },
    ...more sessions...
  ]
}

IMPORTANT SCHEDULING GUIDELINES:
1. Ensure each day's schedule is logical and follows a typical event flow
2. Include appropriate breaks (coffee, lunch, etc.) in the schedule
3. Start times and end times should be in sequence without overlaps
4. Make sure sessions align with the event theme, objectives, and descriptions provided
5. Distribute session topics evenly throughout the day to maintain engagement
6. CRITICALLY IMPORTANT: The dates must be EXACTLY between ${formattedStartDate} and ${formattedEndDate}, inclusive
7. CRITICALLY IMPORTANT: Start each day no earlier than ${eventDetails.startTimeFormatted} and end no later than ${eventDetails.endTimeFormatted}
8. CRITICALLY IMPORTANT: ALL locations MUST be chosen from this exact list: ${locationsText}
9. CRITICALLY IMPORTANT: LEAVE THE SPEAKER/LEAD FIELD EMPTY FOR ALL SESSIONS
`;
  
  Logger.log('Generated OpenAI Prompt:');
  Logger.log(prompt);
  
  return prompt;
}

/**
 * Calls OpenAI API to generate a schedule
 * Modified to enforce approved locations in the response
 * @param {string} prompt - The prompt for OpenAI
 * @param {string} apiKey - OpenAI API key
 * @param {Object} eventDetails - Event details for fallback parsing
 * @param {Array} approvedLocations - List of approved locations
 * @return {Array} Array of schedule items
 */
function callOpenAIForSchedule(prompt, apiKey, eventDetails, approvedLocations) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: "gpt-3.5-turbo-16k",
    messages: [
      {
        role: "system", 
        content: "You are an experienced event planner that specializes in creating detailed, realistic event schedules. You always follow time constraints exactly as provided."
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
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log('Error from OpenAI API: ' + response.getContentText());
      throw new Error('OpenAI API returned error code: ' + responseCode);
    }
    
    const responseJson = JSON.parse(response.getContentText());
    const responseText = responseJson.choices[0].message.content;
    
    // Log a portion of the response for debugging
    Logger.log('OpenAI Response Preview (first 500 chars):');
    Logger.log(responseText.substring(0, 500) + '...');
    
    // Parse the response to extract schedule items
    try {
      // Extract JSON from the response
      let jsonMatch = responseText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        Logger.log('No JSON found in response: ' + responseText);
        // Attempt to parse unstructured response as fallback
        return parseUnstructuredScheduleResponse(responseText, eventDetails, approvedLocations);
      }
      
      const jsonData = JSON.parse(jsonMatch[0]);
      const scheduleItems = jsonData.schedule || [];
      
      // Validate each location against the approved list
      return scheduleItems.map(item => {
        // Create a new object to avoid modifying the original
        const validatedItem = {...item};
        validatedItem.location = validateLocation(item.location, approvedLocations);
        // Ensure the Lead/Speaker field is empty
        validatedItem.speaker = "";
        return validatedItem;
      });
    } catch (e) {
      Logger.log('Error parsing OpenAI response: ' + e.toString());
      Logger.log('Response was: ' + responseText);
      
      // Fallback to unstructured parsing
      return parseUnstructuredScheduleResponse(responseText, eventDetails, approvedLocations);
    }
  } catch (e) {
    Logger.log('Error calling OpenAI API: ' + e.toString());
    throw e;
  }
}

/**
 * Fallback parser for unstructured schedule responses
 * Modified to enforce approved locations
 * @param {string} response The text response from OpenAI
 * @param {Object} eventDetails Event details for date and time context
 * @param {Array} approvedLocations List of approved locations
 * @return {Array} Array of schedule objects
 */
function parseUnstructuredScheduleResponse(response, eventDetails, approvedLocations) {
  // Try to parse the response as a list of schedule items
  const scheduleItems = [];
  const lines = response.split('\n');
  
  let currentDate = eventDetails.startDate; // Start with first event day
  let currentDay = 0;
  
  // Default start time from event details or reasonable default
  let defaultStartTime = eventDetails.startTimeFormatted || '9:00 AM';
  defaultStartTime = defaultStartTime.trim();
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    
    // Check if this line denotes a new day
    if (line.match(/day\s*\d+|day\s*one|day\s*two|day\s*three|first\s*day|second\s*day|third\s*day/i)) {
      // Update the current date based on day number
      const dayMatch = line.match(/\d+|one|two|three|first|second|third/i);
      if (dayMatch) {
        let dayNum = dayMatch[0].toLowerCase();
        // Convert text numbers to numeric
        if (dayNum === 'one' || dayNum === 'first') dayNum = 1;
        else if (dayNum === 'two' || dayNum === 'second') dayNum = 2;
        else if (dayNum === 'three' || dayNum === 'third') dayNum = 3;
        else dayNum = parseInt(dayNum);
        
        // Update current day counter
        currentDay = dayNum - 1; // 0-indexed
        
        // Create new date for this day
        if (eventDetails.startDate instanceof Date) {
          currentDate = new Date(eventDetails.startDate);
          currentDate.setDate(currentDate.getDate() + currentDay);
        }
      }
      continue;
    }
    
    // Try to match time pattern: HH:MM AM/PM - HH:MM AM/PM
    const timePattern = /(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm))\s*(?:-|–|to)\s*(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm))/i;
    const timeMatch = line.match(timePattern);
    
    if (timeMatch) {
      let startTime = timeMatch[1];
      let endTime = timeMatch[2];
      
      // Apply event start time constraints if this is the first session of the day
      // Logic: If this is the first session and it starts earlier than event start time
      const isFirstSessionOfDay = scheduleItems.length === 0 || 
                                  (scheduleItems.length > 0 && 
                                   !areSameDates(scheduleItems[scheduleItems.length-1].date, currentDate));
                                   
if (isFirstSessionOfDay) {
        // Check if proposed start time is earlier than event start time
        const eventStartTimeParts = defaultStartTime.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
        const sessionStartTimeParts = startTime.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
        
        if (eventStartTimeParts && sessionStartTimeParts) {
          let eventHour = parseInt(eventStartTimeParts[1]);
          const eventMinutes = parseInt(eventStartTimeParts[2]);
          const eventPeriod = eventStartTimeParts[3].toUpperCase();
          
          if (eventPeriod === 'PM' && eventHour < 12) eventHour += 12;
          if (eventPeriod === 'AM' && eventHour === 12) eventHour = 0;
          
          let sessionHour = parseInt(sessionStartTimeParts[1]);
          const sessionMinutes = parseInt(sessionStartTimeParts[2]);
          const sessionPeriod = sessionStartTimeParts[3].toUpperCase();
          
          if (sessionPeriod === 'PM' && sessionHour < 12) sessionHour += 12;
          if (sessionPeriod === 'AM' && sessionHour === 12) sessionHour = 0;
          
          // If session starts earlier than event start time, use event start time
          if ((sessionHour < eventHour) || 
              (sessionHour === eventHour && sessionMinutes < eventMinutes)) {
            startTime = defaultStartTime;
            
            // Adjust end time to maintain same duration
            const sessionDurationMs = calculateTimeDifferenceMs(startTime, endTime);
            if (sessionDurationMs > 0) {
              const startTimeDate = new Date();
              eventHour = eventHour % 12; // Convert back to 12-hour format if needed
              if (eventPeriod === 'PM' && eventHour < 12) eventHour += 12;
              if (eventPeriod === 'AM' && eventHour === 12) eventHour = 0;
              
              startTimeDate.setHours(eventHour, eventMinutes, 0, 0);
              const endTimeDate = new Date(startTimeDate.getTime() + sessionDurationMs);
              
              // Format end time as HH:MM AM/PM
              let endHour = endTimeDate.getHours();
              const endMinutes = endTimeDate.getMinutes();
              const endPeriod = endHour >= 12 ? 'PM' : 'AM';
              endHour = endHour % 12 || 12; // Convert to 12-hour format
              
              endTime = `${endHour}:${endMinutes < 10 ? '0' + endMinutes : endMinutes} ${endPeriod}`;
            }
          }
        }
      }
      
      // Apply event end time constraints
      // Ensure no sessions end after the event end time
      const eventEndTimeParts = eventDetails.endTimeFormatted.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
      const sessionEndTimeParts = endTime.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
      
      if (eventEndTimeParts && sessionEndTimeParts) {
        let eventHour = parseInt(eventEndTimeParts[1]);
        const eventMinutes = parseInt(eventEndTimeParts[2]);
        const eventPeriod = eventEndTimeParts[3].toUpperCase();
        
        if (eventPeriod === 'PM' && eventHour < 12) eventHour += 12;
        if (eventPeriod === 'AM' && eventHour === 12) eventHour = 0;
        
        let sessionHour = parseInt(sessionEndTimeParts[1]);
        const sessionMinutes = parseInt(sessionEndTimeParts[2]);
        const sessionPeriod = sessionEndTimeParts[3].toUpperCase();
        
        if (sessionPeriod === 'PM' && sessionHour < 12) sessionHour += 12;
        if (sessionPeriod === 'AM' && sessionHour === 12) sessionHour = 0;
        
        // If session ends later than event end time, adjust it
        if ((sessionHour > eventHour) || 
            (sessionHour === eventHour && sessionMinutes > eventMinutes)) {
          
          // Use the event end time as the session end time
          const endHour12 = eventHour % 12 || 12; // Convert to 12-hour format
          endTime = `${endHour12}:${eventMinutes < 10 ? '0' + eventMinutes : eventMinutes} ${eventPeriod}`;
          
          // Also adjust the start time if necessary to ensure the session has a reasonable duration
          // (e.g., at least 15 minutes)
          const minSessionDurationMs = 15 * 60 * 1000; // 15 minutes in milliseconds
          
          const adjustedEndTime = new Date();
          adjustedEndTime.setHours(eventHour, eventMinutes, 0, 0);
          
          const startTimeObj = parseTimeString(startTime);
          if (startTimeObj) {
            const sessionDurationMs = adjustedEndTime.getTime() - startTimeObj.getTime();
            
            // If the adjusted session would be too short, adjust the start time
            if (sessionDurationMs < minSessionDurationMs) {
              const newStartTimeObj = new Date(adjustedEndTime.getTime() - minSessionDurationMs);
              const startHour12 = newStartTimeObj.getHours() % 12 || 12;
              const startMinutes = newStartTimeObj.getMinutes();
              const startPeriod = newStartTimeObj.getHours() >= 12 ? 'PM' : 'AM';
              
              startTime = `${startHour12}:${startMinutes < 10 ? '0' + startMinutes : startMinutes} ${startPeriod}`;
            }
          }
        }
      }
      
      // Get the remainder of the line after the time
      let remainingText = line.substring(timeMatch[0].length).trim();
      
      // Look for location in brackets or parentheses
      let location = approvedLocations[0]; // Default to first approved location
      const locationPattern = /[\(\[](.*?)[\)\]]/;
      const locationMatch = remainingText.match(locationPattern);
      
      if (locationMatch) {
        // Validate the extracted location against approved list
        const extractedLocation = locationMatch[1].trim();
        location = validateLocation(extractedLocation, approvedLocations);
        remainingText = remainingText.replace(locationMatch[0], '').trim();
      }
      
      // Extract session title and speaker
      let title = remainingText;
      
      // Add the schedule item - ALWAYS with empty speaker/lead
      scheduleItems.push({
        date: sessionDate,
        startTime: startTime,
        endTime: endTime,
        title: title || "Untitled Session",
        speaker: "", // Always empty for preliminary schedule
        location: location,
        status: "Tentative"
      });
    }
  }
  
  // Ensure we have at least some sessions
  if (scheduleItems.length === 0) {
    // If no sessions were parsed, create a basic template
    const startDate = new Date(eventDetails.startDate);
    const formattedStartTime = eventDetails.startTimeFormatted;
    
    // Calculate an end time 1 hour after start
    const startTimeObj = parseTimeString(formattedStartTime);
    let endTimeObj = new Date(startTimeObj.getTime() + 60 * 60 * 1000); // 1 hour later
    let endHour = endTimeObj.getHours() % 12 || 12;
    let endMinutes = endTimeObj.getMinutes();
    let endPeriod = endTimeObj.getHours() >= 12 ? 'PM' : 'AM';
    let formattedEndTime = `${endHour}:${endMinutes < 10 ? '0' + endMinutes : endMinutes} ${endPeriod}`;
    
    // Add a basic opening session
    scheduleItems.push({
      date: startDate,
      startTime: formattedStartTime,
      endTime: formattedEndTime,
      title: "Opening Session",
      speaker: "", // Always empty for preliminary schedule
      location: approvedLocations[0], // Use first approved location
      status: "Tentative"
    });
  }
  
  return scheduleItems;
}

/**
 * Helper function to check if two dates represent the same calendar day
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @return {boolean} True if dates are the same day, false otherwise
 */
function areSameDates(date1, date2) {
  if (!(date1 instanceof Date) || !(date2 instanceof Date)) {
    return false;
  }
  
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * Calculate time difference in milliseconds between two time strings
 * @param {string} startTimeStr - Start time string in format "HH:MM AM/PM"
 * @param {string} endTimeStr - End time string in format "HH:MM AM/PM"
 * @return {number} Difference in milliseconds
 */
function calculateTimeDifferenceMs(startTimeStr, endTimeStr) {
  try {
    const startTime = parseTimeString(startTimeStr);
    const endTime = parseTimeString(endTimeStr);
    
    if (!startTime || !endTime) {
      return 0;
    }
    
    let diffMs = endTime.getTime() - startTime.getTime();
    if (diffMs < 0) {
      diffMs += 24 * 60 * 60 * 1000; // Add a day if end time is on next day
    }
    
    return diffMs;
  } catch (e) {
    Logger.log('Error calculating time difference: ' + e);
    return 0;
  }
}

/**
 * Populates the Schedule sheet with generated schedule items
 * Enhanced to ensure all locations come from the approved list
 * and Lead column remains blank
 * @param {Array} scheduleData - Array of schedule items
 * @param {Object} eventDetails - Event details
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Schedule sheet
 * @param {Array} approvedLocations - List of approved locations
 * @return {number} The number of schedule items added
 */
function populateScheduleSheet(scheduleData, eventDetails, sheet, approvedLocations) {
  if (!scheduleData || scheduleData.length === 0) {
    return 0;
  }
  
  // Clear the existing schedule data (preserving the header row)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Prepare the data for batch insertion
  const scheduleRows = [];
  
  // Get event constraints for validation
  const eventStartTime = parseTimeString(eventDetails.startTimeFormatted);
  const eventEndTime = parseTimeString(eventDetails.endTimeFormatted);
  
  // Time constraints in milliseconds for validation
  const earliestAllowedTimeMs = eventStartTime ? eventStartTime.getTime() : null;
  const latestAllowedTimeMs = eventEndTime ? eventEndTime.getTime() : null;
  
  // Process each schedule item with validation
  for (const item of scheduleData) {
    let itemDate;
    // Handle date which could be a string or Date object
    if (item.date instanceof Date) {
      itemDate = item.date;
    } else if (typeof item.date === 'string') {
      // Try to parse the date string
      try {
        itemDate = new Date(item.date);
        if (isNaN(itemDate.getTime())) {
          // If invalid, fallback to event start date
          itemDate = new Date(eventDetails.startDate);
        }
      } catch (e) {
        // Fallback to event start date on error
        itemDate = new Date(eventDetails.startDate);
      }
    } else {
      // Default to event start date if no date provided
      itemDate = new Date(eventDetails.startDate);
    }
    
    // Validate the date is within event range
    if (itemDate < eventDetails.startDate) {
      itemDate = new Date(eventDetails.startDate);
    } else if (itemDate > eventDetails.endDate) {
      itemDate = new Date(eventDetails.endDate);
    }
    
    // Ensure start time is not earlier than event start time for first session of the day
    let startTime = item.startTime || eventDetails.startTimeFormatted || '9:00 AM';
    let endTime = item.endTime || '10:00 AM';
    
    // Validate session times against event constraints
    const sessionStartTime = parseTimeString(startTime);
    const sessionEndTime = parseTimeString(endTime);
    
    if (sessionStartTime && sessionEndTime && earliestAllowedTimeMs && latestAllowedTimeMs) {
      const sessionStartMs = sessionStartTime.getTime();
      const sessionEndMs = sessionEndTime.getTime();
      
      // Adjust session if it starts too early
      if (sessionStartMs < earliestAllowedTimeMs) {
        // Use event start time
        startTime = eventDetails.startTimeFormatted;
        
        // Maintain original session duration if possible
        const originalDuration = sessionEndMs - sessionStartMs;
        if (originalDuration > 0) {
          const newEndTime = new Date(earliestAllowedTimeMs + originalDuration);
          const newEndHour = newEndTime.getHours();
          const newEndMinute = newEndTime.getMinutes();
          const period = newEndHour >= 12 ? 'PM' : 'AM';
          const hour12 = newEndHour % 12 || 12;
          
          endTime = `${hour12}:${newEndMinute < 10 ? '0' + newEndMinute : newEndMinute} ${period}`;
        }
      }
      
      // Adjust session if it ends too late
      const adjustedSessionEndTime = parseTimeString(endTime);
      if (adjustedSessionEndTime && adjustedSessionEndTime.getTime() > latestAllowedTimeMs) {
        // Use event end time
        endTime = eventDetails.endTimeFormatted;
        
        // If this makes the session too short or start after end, adjust the start time too
        const adjustedSessionStartTime = parseTimeString(startTime);
        const minSessionMs = 15 * 60 * 1000; // Minimum 15 minutes
        
        if (adjustedSessionStartTime && 
            (latestAllowedTimeMs - adjustedSessionStartTime.getTime() < minSessionMs || 
             adjustedSessionStartTime.getTime() >= latestAllowedTimeMs)) {
          // Start 30 min before end time
          const newStartTime = new Date(latestAllowedTimeMs - 30 * 60 * 1000);
          const newStartHour = newStartTime.getHours();
          const newStartMinute = newStartTime.getMinutes();
          const period = newStartHour >= 12 ? 'PM' : 'AM';
          const hour12 = newStartHour % 12 || 12;
          
          startTime = `${hour12}:${newStartMinute < 10 ? '0' + newStartMinute : newStartMinute} ${period}`;
        }
      }
    }
    
    // Validate location against approved locations
    const location = validateLocation(item.location, approvedLocations);
    
    // Create a row for this schedule item
    // NOTE: The key change - setting the Lead field (column 6) to empty string regardless of what's returned from OpenAI
    scheduleRows.push([
      itemDate,                    // Date
      startTime,                   // Start Time
      endTime,                     // End Time
      '',                          // Duration (will be calculated)
      item.title || 'Untitled Session', // Session Title
      '',                          // Lead/Speaker - ALWAYS EMPTY for preliminary schedule
      location,                    // Location - validated against approved list
      item.status || 'Tentative',  // Status
      ''                           // Notes
    ]);
  }
  
  // Sort rows by date and time
  scheduleRows.sort((a, b) => {
    // First sort by date
    const dateCompare = a[0] - b[0];
    if (dateCompare !== 0) return dateCompare;
    
    // Then by start time
    const startTimeA = parseTimeString(a[1]);
    const startTimeB = parseTimeString(b[1]);
    if (startTimeA && startTimeB) {
      return startTimeA.getTime() - startTimeB.getTime();
    }
    return 0;
  });
  
  // Batch insert all schedule items
  if (scheduleRows.length > 0) {
    sheet.getRange(2, 1, scheduleRows.length, scheduleRows[0].length).setValues(scheduleRows);
    
    // Format date and time columns
    sheet.getRange(2, 1, scheduleRows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 2, scheduleRows.length, 2).setNumberFormat('hh:mm am/pm');
    
    // Re-apply the duration calculation
    setupDurationCalculation(sheet.getParent());
  }
  
  return scheduleRows.length;
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
    
    // Regular expression to match time format: HH:MM AM/PM
    const timeRegex = /^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)$/;
    const match = timeStr.match(timeRegex);
    
    if (!match) return null;
    
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const period = match[3].toUpperCase();
    
    // Convert to 24-hour format
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
 * Format a date as YYYY-MM-DD for API calls
 * @param {Date} date The date to format
 * @return {string} Formatted date string
 */
function formatDate(date) {
  if (!date) return '';
  
  // Ensure date is a Date object 
  if (!(date instanceof Date)) {
    try {
      date = new Date(date);
    } catch (e) {
      return '';
    }
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Gets the OpenAI API key from the Config sheet with improved debugging
 * @return {string|null} The API key or null if not found
 */
function getOpenAIApiKey() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    Logger.log('Config sheet not found');
    return null;
  }
  
  try {
    // First method: Try to find the row labeled "OpenAI API Key"
    const keyRange = configSheet.getRange("A:A");
    const keyValues = keyRange.getValues();
    let apiKeyRow = -1;
    
    Logger.log('Searching for OpenAI API Key row in Config sheet...');
    
    for (let i = 0; i < keyValues.length; i++) {
      if (keyValues[i][0] === "OpenAI API Key") {
        apiKeyRow = i + 1; // Convert to 1-based row index
        Logger.log(`Found OpenAI API Key row at position ${apiKeyRow}`);
        break;
      }
    }
    
    let apiKey = null;
    
    if (apiKeyRow !== -1) {
      // If we found the label, get the key from that row
      apiKey = configSheet.getRange(apiKeyRow, 2).getValue();
      Logger.log(`Found API key via label search at row ${apiKeyRow}`);
    } else {
      // Fallback to directly accessing B15 (most reliable)
      apiKey = configSheet.getRange("B15").getValue();
      Logger.log('Using direct cell reference to B15 for API key');
    }
    
    const apiKeyStr = apiKey ? apiKey.toString().trim() : '';
    
    if (!apiKeyStr) {
      Logger.log('OpenAI API Key is empty');
    } else {
      Logger.log(`Found API key (length: ${apiKeyStr.length}, first 4 chars: ${apiKeyStr.substring(0, 4)}...)`);
    }
    
    return apiKeyStr || null;
  } catch (e) {
    Logger.log('Error getting API key: ' + e.toString());
    return null;
  }
}