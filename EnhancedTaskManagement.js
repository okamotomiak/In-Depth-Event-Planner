//EnhancedTaskManagement.gs - Upgraded task generator that analyzes the schedule

/**
 * Enhanced version of generateAITasks that includes schedule analysis
 * This replaces the existing generateAITasks function
 */
function generateAITasksWithSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  Logger.log('Starting Enhanced AI Task Generation with Schedule Analysis...');
  
  try {
    // Step 1: Get event information
    const eventInfo = getEventInformation();
    if (!eventInfo) {
      ui.alert('Error', 'Could not retrieve event information. Please ensure the Event Description sheet is properly filled out.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 2: Get schedule information
    const scheduleInfo = getScheduleInformation(ss);
    
    // Step 3: Get people information (speakers, staff, etc.)
    const peopleInfo = getPeopleInformation(ss);
    
    // Step 4: Get OpenAI API key
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      ui.alert('Error', 'OpenAI API key not found. Please add your API key to Script Properties or cell B15 in Config sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 5: Show loading message
    ui.alert('Processing', 'Generating enhanced tasks using AI with schedule analysis. This may take a few moments...', ui.ButtonSet.OK);
    
    // Step 6: Generate tasks using enhanced prompt
    const tasks = generateEnhancedTasksWithAI(eventInfo, scheduleInfo, peopleInfo, apiKey);
    if (!tasks || tasks.length === 0) {
      ui.alert('Error', 'Failed to generate tasks. Please try again later.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 7: Clear existing tasks
    clearExistingTasks();
    
    // Step 8: Add generated tasks to the Task Management sheet
    const taskCount = addTasksToSheet(tasks, eventInfo);
    
    // Step 9: Show success message
    ui.alert('Success', `Generated and added ${taskCount} enhanced tasks to the Task Management sheet, including ${scheduleInfo.sessions.length} session-specific tasks.`, ui.ButtonSet.OK);
  
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    ui.alert('Error', 'An error occurred: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Gets comprehensive schedule information from the Schedule sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @return {Object} Object with schedule details
 */
function getScheduleInformation(ss) {
  const scheduleSheet = ss.getSheetByName('Schedule');
  
  const scheduleInfo = {
    sessions: [],
    sessionCount: 0,
    hasMusic: false,
    hasWorkshops: false,
    hasSpeakers: false,
    requiresAV: false,
    locations: [],
    sessionTypes: []
  };
  
  if (!scheduleSheet) {
    Logger.log('Schedule sheet not found, generating basic tasks only');
    return scheduleInfo;
  }
  
  const data = scheduleSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('Schedule sheet is empty, generating basic tasks only');
    return scheduleInfo;
  }
  
  // Get headers to find column indices
  const headers = data[0];
  const dateIndex = headers.findIndex(h => h.toString().toLowerCase().includes('date'));
  const startTimeIndex = headers.findIndex(h => h.toString().toLowerCase().includes('start'));
  const endTimeIndex = headers.findIndex(h => h.toString().toLowerCase().includes('end'));
  const titleIndex = headers.findIndex(h => h.toString().toLowerCase().includes('title'));
  const leadIndex = headers.findIndex(h => h.toString().toLowerCase().includes('lead'));
  const locationIndex = headers.findIndex(h => h.toString().toLowerCase().includes('location'));
  const statusIndex = headers.findIndex(h => h.toString().toLowerCase().includes('status'));
  
  Logger.log(`Schedule column indices - Date: ${dateIndex}, Title: ${titleIndex}, Location: ${locationIndex}`);
  
  // Process each session
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[titleIndex] || row[titleIndex].toString().trim() === '') continue;
    
    const session = {
      date: row[dateIndex] || '',
      startTime: row[startTimeIndex] || '',
      endTime: row[endTimeIndex] || '',
      title: row[titleIndex] || '',
      lead: row[leadIndex] || '',
      location: row[locationIndex] || '',
      status: row[statusIndex] || 'Tentative'
    };
    
    scheduleInfo.sessions.push(session);
    
    // Analyze session characteristics
    const title = session.title.toLowerCase();
    const location = session.location.toLowerCase();
    
    // Detect session types
    if (title.includes('music') || title.includes('concert') || title.includes('band') || title.includes('performance')) {
      scheduleInfo.hasMusic = true;
      scheduleInfo.sessionTypes.push('music');
    }
    
    if (title.includes('workshop') || title.includes('training') || title.includes('hands-on') || title.includes('demo')) {
      scheduleInfo.hasWorkshops = true;
      scheduleInfo.sessionTypes.push('workshop');
    }
    
    if (title.includes('presentation') || title.includes('keynote') || title.includes('talk') || title.includes('speech')) {
      scheduleInfo.hasSpeakers = true;
      scheduleInfo.sessionTypes.push('presentation');
    }
    
    if (title.includes('av') || title.includes('audio') || title.includes('video') || title.includes('projection')) {
      scheduleInfo.requiresAV = true;
    }
    
    // Collect unique locations
    if (session.location && !scheduleInfo.locations.includes(session.location)) {
      scheduleInfo.locations.push(session.location);
    }
  }
  
  scheduleInfo.sessionCount = scheduleInfo.sessions.length;
  scheduleInfo.sessionTypes = [...new Set(scheduleInfo.sessionTypes)]; // Remove duplicates
  
  Logger.log(`Schedule analysis: ${scheduleInfo.sessionCount} sessions, types: ${scheduleInfo.sessionTypes.join(', ')}`);
  
  return scheduleInfo;
}

/**
 * Gets people information from the People sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet
 * @return {Object} Object with people details
 */
function getPeopleInformation(ss) {
  const peopleSheet = ss.getSheetByName('People');
  
  const peopleInfo = {
    speakers: [],
    staff: [],
    volunteers: [],
    totalPeople: 0
  };
  
  if (!peopleSheet) {
    Logger.log('People sheet not found');
    return peopleInfo;
  }
  
  const data = peopleSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('People sheet is empty');
    return peopleInfo;
  }
  
  // Get headers
  const headers = data[0];
  const nameIndex = headers.findIndex(h => h.toString().toLowerCase().includes('name'));
  const categoryIndex = headers.findIndex(h => h.toString().toLowerCase().includes('category'));
  const statusIndex = headers.findIndex(h => h.toString().toLowerCase().includes('status'));
  
  // Process each person
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (!row[nameIndex]) continue;
    
    const person = {
      name: row[nameIndex],
      category: row[categoryIndex] || '',
      status: row[statusIndex] || ''
    };
    
    // Categorize people
    const category = person.category.toLowerCase();
    if (category.includes('speaker')) {
      peopleInfo.speakers.push(person);
    } else if (category.includes('staff')) {
      peopleInfo.staff.push(person);
    } else if (category.includes('volunteer')) {
      peopleInfo.volunteers.push(person);
    }
    
    peopleInfo.totalPeople++;
  }
  
  Logger.log(`People analysis: ${peopleInfo.speakers.length} speakers, ${peopleInfo.staff.length} staff, ${peopleInfo.volunteers.length} volunteers`);
  
  return peopleInfo;
}

/**
 * Generate enhanced tasks using OpenAI API with schedule and people analysis
 * @param {Object} eventInfo Event information
 * @param {Object} scheduleInfo Schedule analysis
 * @param {Object} peopleInfo People information
 * @param {string} apiKey OpenAI API key
 * @return {Array} Array of task objects
 */
function generateEnhancedTasksWithAI(eventInfo, scheduleInfo, peopleInfo, apiKey) {
  try {
    // Format dates for the prompt
    const startDate = formatDate(eventInfo.startDate);
    const endDate = formatDate(eventInfo.endDate);
    
    // Create session-specific context
    let sessionContext = '';
    if (scheduleInfo.sessions.length > 0) {
      sessionContext = `\n\nSCHEDULE ANALYSIS (VERY IMPORTANT - CREATE SPECIFIC TASKS FOR THESE):
The event includes ${scheduleInfo.sessionCount} scheduled sessions:

`;
      
      scheduleInfo.sessions.forEach((session, index) => {
        sessionContext += `${index + 1}. "${session.title}" at ${session.location} (${session.startTime} - ${session.endTime})`;
        if (session.lead) sessionContext += ` - Led by ${session.lead}`;
        sessionContext += `\n`;
      });
      
      sessionContext += `\nSession Types Detected: ${scheduleInfo.sessionTypes.join(', ')}
Locations Used: ${scheduleInfo.locations.join(', ')}
Music/Performances: ${scheduleInfo.hasMusic ? 'Yes' : 'No'}
Workshops/Training: ${scheduleInfo.hasWorkshops ? 'Yes' : 'No'}
Speaker Presentations: ${scheduleInfo.hasSpeakers ? 'Yes' : 'No'}

CRITICAL: Create specific preparation tasks for EACH session listed above, considering:
- Setup requirements for each location
- AV/technical needs for each session type
- Coordination with session leaders
- Materials and equipment needed
- Pre-event testing and rehearsals`;
    }
    
    // Create people context
    let peopleContext = '';
    if (peopleInfo.totalPeople > 0) {
      peopleContext = `\n\nTEAM ANALYSIS:
- ${peopleInfo.speakers.length} Speakers: ${peopleInfo.speakers.map(s => s.name).join(', ')}
- ${peopleInfo.staff.length} Staff Members: ${peopleInfo.staff.map(s => s.name).join(', ')}
- ${peopleInfo.volunteers.length} Volunteers: ${peopleInfo.volunteers.map(v => v.name).join(', ')}

Create coordination and communication tasks for team members.`;
    }
    
    // Create the enhanced prompt
    const prompt = `Generate a comprehensive task list for planning and executing the following event:

EVENT DETAILS:
- Name: ${eventInfo.eventName}
- Date(s): ${startDate}${eventInfo.durationDays > 1 ? ` to ${endDate}` : ''}
- Duration: ${eventInfo.durationDays} day(s)
- Location: ${eventInfo.location || 'TBD'}
- Theme: ${eventInfo.theme || 'N/A'}
- Objectives: ${eventInfo.objectives || 'N/A'}
- Description: ${eventInfo.description || 'N/A'}${sessionContext}${peopleContext}

Create tasks in three phases:
1. PRE-EVENT TASKS (due before ${startDate})
2. DURING-EVENT TASKS (during the event)
3. POST-EVENT TASKS (after ${endDate})

IMPORTANT REQUIREMENTS:
- Create specific tasks for EACH scheduled session
- Include session-specific setup, coordination, and breakdown tasks
- Consider the unique requirements of each session type
- Include tasks for coordinating with named speakers/leaders
- Add location-specific preparation tasks
- Include testing and rehearsal tasks for technical sessions

For each task, include:
- Task Name: Specific and actionable
- Description: Detailed explanation
- Category: Venue, Marketing, Logistics, Program, Budget, Staffing, Technology, Communications, Other
- Priority: Critical, High, Medium, or Low
- Timeline: Specific timing relative to event or session
- Status: "Not Started"

Return response in this JSON format:
{
  "tasks": [
    {
      "name": "Task Name",
      "description": "Detailed task description",
      "category": "Category",
      "priority": "Priority",
      "timeline": "Timeline",
      "status": "Not Started"
    }
  ]
}

Focus on creating actionable, specific tasks rather than generic ones. Each session should have 2-4 related tasks.`;

    // Call OpenAI API with enhanced prompt
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: "You are an expert event planning assistant that creates detailed, session-specific task lists for events. You analyze schedules and create specific preparation tasks for each session."
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
    
    Logger.log('Calling OpenAI API with enhanced prompt...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`API Error (${responseCode}): ${response.getContentText()}`);
      throw new Error(`OpenAI API returned error code: ${responseCode}`);
    }
    
    const responseJson = JSON.parse(response.getContentText());
    const content = responseJson.choices[0].message.content;
    
    Logger.log(`Received enhanced response (${content.length} chars)`);
    
    // Parse the JSON response
    const tasks = parseTasksFromResponse(content);
    
    Logger.log(`Successfully extracted ${tasks.length} enhanced tasks`);
    return tasks;
    
  } catch (error) {
    Logger.log('Error generating enhanced tasks: ' + error.toString());
    throw error;
  }
}

/**
 * Add enhanced menu item to replace the basic task generator
 * Add this to your onOpen function
 */
function addEnhancedTaskMenuItem(menu) {
  return menu.addItem('Generate Enhanced AI Tasks (with Schedule)', 'generateAITasksWithSchedule');
}