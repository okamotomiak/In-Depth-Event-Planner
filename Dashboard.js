// Dashboard.gs - Add chart functionality

/**
 * Creates a task category bar chart for the dashboard
 * This pulls live data from the Task Management sheet
 * Modified to remove the chart title text
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} taskSheet The task management sheet
 */
function createTaskCategoryChart(dashboardSheet, taskSheet) {
  try {
    // Find task category column
    const taskHeaders = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
    const categoryColIndex = taskHeaders.findIndex(header => 
      header.toString().toLowerCase().trim() === 'category');
    
    if (categoryColIndex === -1) {
      Logger.log('Category column not found in Task Management sheet');
      return;
    }
    
    // Get category values (skip header row)
    const lastRow = taskSheet.getLastRow();
    
    // Remove any existing chart first
    const existingCharts = dashboardSheet.getCharts();
    for (let i = 0; i < existingCharts.length; i++) {
      if (existingCharts[i].getOptions().get('title') === 'Tasks by Category') {
        dashboardSheet.removeChart(existingCharts[i]);
      }
    }
    
    // If there's no data (only header row or empty sheet), don't create a new chart
    if (lastRow <= 1) {
      Logger.log('No data in Task Management sheet, no chart created');
      return;
    }
    
    const categoryRange = taskSheet.getRange(2, categoryColIndex + 1, lastRow - 1, 1);
    const categoryValues = categoryRange.getValues();
    
    // Count occurrences of each category
    const categoryCounts = {};
    for (let i = 0; i < categoryValues.length; i++) {
      const category = categoryValues[i][0];
      if (!category) continue; // Skip empty cells
      
      if (categoryCounts[category]) {
        categoryCounts[category]++;
      } else {
        categoryCounts[category] = 1;
      }
    }
    
    // If no valid categories found, don't create a chart
    if (Object.keys(categoryCounts).length === 0) {
      Logger.log('No valid categories found in Task Management sheet, no chart created');
      return;
    }
    
    // Create a hidden data range for the chart
    // First, find or create a hidden sheet for chart data
    let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_ChartData2');
    if (!dataSheet) {
      dataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('_ChartData2');
      dataSheet.hideSheet(); // Hide this utility sheet
    } else {
      dataSheet.clear(); // Clear any existing data
    }
    
    // Populate data for chart
    const chartData = [['Category', 'Count']];
    Object.keys(categoryCounts).forEach(category => {
      chartData.push([category, categoryCounts[category]]);
    });
    
    // Write data to the hidden sheet
    dataSheet.getRange(1, 1, chartData.length, 2).setValues(chartData);
    
    // Create the chart - REMOVED the row 11 header text
    const chart = dashboardSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 1, chartData.length, 2))
      .setPosition(12, 7, 0, 0) // Row 12, Column 7 (adjusted position)
      .setOption('title', 'Tasks by Category')
      .setOption('legend', {position: 'none'})
      .setOption('width', 500)
      .setOption('height', 300)
      .build();
    
    // Add the new chart
    dashboardSheet.insertChart(chart);
    
    // No longer adding a header above the chart
    // REMOVED: dashboardSheet.getRange('G11').setValue('Tasks by Category Visualization')
    
    Logger.log('Task category chart created successfully');
  } catch (error) {
    Logger.log(`Error creating task category chart: ${error}`);
  }
}

/**
 * Builds a visual dashboard layout with live formulas, metrics, and tables.
 * Modified to only include the task category bar chart.
 */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dashboard');
  
  // Store references to other sheets for direct references
  const taskSheet = ss.getSheetByName('Task Management');
  const eventDescSheet = ss.getSheetByName('Event Description');
  
  if (!sheet) sheet = ss.insertSheet('Dashboard');
  sheet.clear();
  sheet.setTabColor('#e06666');
  sheet.setFrozenRows(1);

  // Title
  const titleRange = sheet.getRange('A1:H1');
  titleRange.merge();
  titleRange.setValue('Event Planner Dashboard');
  
  // Batch format the title instead of individual calls
  titleRange.setFontWeight('bold')
           .setFontSize(16)
           .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 30);

  // Rather than using INDIRECT, calculate values directly in script
  
  // Get task count directly
  let totalTasks = 0;
  let completedTasks = 0;
  
  // Calculate status counts for all status types
  const statusCounts = {
    'Not Started': 0,
    'In Progress': 0,
    'Done': 0,
    'Overdue': 0,
    'Blocked': 0,
    'Cancelled': 0
  };
  
  if (taskSheet) {
    const lastRow = Math.max(2, taskSheet.getLastRow());
    if (lastRow > 1) {
      // Get all task data at once
      const dataRange = taskSheet.getRange(1, 1, lastRow, taskSheet.getLastColumn());
      const taskData = dataRange.getValues();
      
      // Find the status column index
      const taskHeaders = taskData[0];
      const statusColIndex = taskHeaders.findIndex(header => 
        header.toString().toLowerCase() === 'status');
      
      if (statusColIndex !== -1) {
        // Count tasks by status - start from row 1 (after header)
        for (let i = 1; i < taskData.length; i++) {
          const row = taskData[i];
          
          // Skip empty rows (check if first cell/Task ID is empty)
          if (!row[0]) continue;
          
          totalTasks++;
          
          const status = row[statusColIndex];
          
          // Update status counts
          if (status in statusCounts) {
            statusCounts[status]++;
          }
          
          // Count completed tasks
          if (status === 'Done') {
            completedTasks++;
          }
        }
      }
    }
  }
  
  // Calculate percent completed
  const percentCompleted = totalTasks > 0 ? completedTasks / totalTasks : 0;
  
  // Find attendance and profit goals
  let attendanceGoal = 0;
  let profitGoal = 0;
  
  if (eventDescSheet) {
    const lastRow = Math.max(2, eventDescSheet.getLastRow());
    if (lastRow > 1) {
      const labelRange = eventDescSheet.getRange(1, 1, lastRow, 1);
      const valueRange = eventDescSheet.getRange(1, 2, lastRow, 1);
      
      const labels = labelRange.getValues();
      const values = valueRange.getValues();
      
      for (let i = 0; i < labels.length; i++) {
        const label = labels[i][0];
        if (label === 'Attendance Goal (#)') {
          attendanceGoal = values[i][0] || 0;
        } else if (label === 'Profit Goal ($)') {
          profitGoal = values[i][0] || 0;
        }
      }
    }
  }
  
  // Section: KPI Summary Boxes - now using direct values instead of expensive INDIRECT formulas
  const kpiLabels = ['Total Tasks', 'Completed Tasks', '% Completed', 'Expected Attendance', 'Profit Goal'];
  const kpiValues = [totalTasks, completedTasks, percentCompleted, attendanceGoal, profitGoal];
  
  // Batch set values and formatting for KPIs
  const kpiLabelRanges = [];
  const kpiValueRanges = [];
  
  for (let i = 0; i < kpiLabels.length; i++) {
    const col = i * 2 + 1;
    // Store ranges for batch operations
    kpiLabelRanges.push(sheet.getRange(2, col, 1, 2));
    kpiValueRanges.push(sheet.getRange(3, col, 1, 2));
  }
  
  // Merge cells in batch operations when possible
  kpiLabelRanges.forEach(range => range.merge());
  kpiValueRanges.forEach(range => range.merge());
  
  // Set KPI row heights
  sheet.setRowHeights(2, 2, 28);
  
  // Set values in batch where possible
  for (let i = 0; i < kpiLabels.length; i++) {
    const labelRange = kpiLabelRanges[i];
    const valueRange = kpiValueRanges[i];
    
    labelRange.setValue(kpiLabels[i])
              .setFontWeight('bold')
              .setHorizontalAlignment('center')
              .setBackground('#cccccc');
    
    // Set value with appropriate formatting
    valueRange.setValue(kpiValues[i])
              .setFontSize(14)
              .setHorizontalAlignment('center');
              
    // Format percentage cell specially
    if (i === 2) { // % Completed
      valueRange.setNumberFormat('0.0%');
    } else if (i === 4) { // Profit Goal
      valueRange.setNumberFormat('$#,##0.00');
    }
  }

  // Section: Upcoming Sessions Table - use direct queries instead of INDIRECT
  sheet.getRange('A5').setValue('Upcoming Sessions (next 10)')
    .setFontWeight('bold')
    .setFontSize(12);
    
  const headerRange = sheet.getRange('A6:I6');
  headerRange.setValues([
    ['Date','Start Time','End Time','Duration','Session Title','Lead','Location','Status','Notes']
  ]);
  headerRange.setFontWeight('bold').setBackground('#eeeeee');
  
  // Get upcoming sessions directly (replacing QUERY with direct script)
  const scheduleSheet = ss.getSheetByName('Schedule');
  let upcomingSessions = [['No upcoming sessions']]; // Default message
  
  if (scheduleSheet) {
    const lastRow = Math.max(2, scheduleSheet.getLastRow());
    
    if (lastRow > 1) {
      // Get all schedule data at once
      const scheduleRange = scheduleSheet.getRange(2, 1, lastRow-1, 9);
      const scheduleData = scheduleRange.getValues();
      
      // Filter for future sessions
      const today = new Date();
      today.setHours(0, 0, 0, 0); // Set to midnight for date comparison
      
      const futureSessions = scheduleData.filter(row => {
        // Check if date is valid and in the future
        return row[0] instanceof Date && row[0] >= today;
      });
      
      // Sort by date (ascending)
      futureSessions.sort((a, b) => a[0] - b[0]);
      
      // Take only the first 10
      upcomingSessions = futureSessions.slice(0, 10);
      
      // If no future sessions were found
      if (upcomingSessions.length === 0) {
        upcomingSessions = [['No upcoming sessions']];
      }
    }
  }
  
  // Set the upcoming sessions data
  if (upcomingSessions.length > 0) {
    const sessionRange = sheet.getRange(7, 1, upcomingSessions.length, upcomingSessions[0].length);
    sessionRange.setValues(upcomingSessions);
    
    // Format dates if we have actual session data (not the "No upcoming sessions" message)
    if (upcomingSessions[0].length > 1) {
      sheet.getRange(7, 1, upcomingSessions.length, 1).setNumberFormat('yyyy-mm-dd');
      sheet.getRange(7, 2, upcomingSessions.length, 2).setNumberFormat('hh:mm am/pm');
    }
  }

  // Add Task Status Summary Table - direct calculation instead of COUNTIF with INDIRECT
  sheet.getRange('K5').setValue('Task Status Summary').setFontWeight('bold');
  
  const statusHeaderRange = sheet.getRange('K6:L6');
  statusHeaderRange.setValues([['Status', 'Count']])
                   .setFontWeight('bold')
                   .setBackground('#eeeeee');
  
  // Create the status summary table with ALL status types
  const statusLabels = ['Not Started', 'In Progress', 'Done', 'Overdue', 'Blocked', 'Cancelled'];
  
  // Create data array for the status summary
  const statusData = statusLabels.map(label => [label, statusCounts[label] || 0]);
  
  // Set status labels and counts
  sheet.getRange(7, 11, statusLabels.length, 2).setValues(statusData);
  
  // Format the dashboard cleanly
  sheet.setRowHeight(6, 24);
  sheet.setColumnWidths(1, 9, 110);

  // Create only the category bar chart
  createTaskCategoryChart(sheet, taskSheet);

  // Add Button to Refresh Dashboard
  const buttonCell = sheet.getRange('H1');
  buttonCell.setValue('🔄 Refresh')
           .setFontWeight('bold')
           .setFontColor('#ffffff')
           .setBackground('#4CAF50')
           .setHorizontalAlignment('center');
  sheet.setColumnWidth(8, 100);
  
  const refreshButton = sheet.getRange('H1');
  refreshButton.setNote('Click to refresh dashboard');
  
  // Remove ALL automatic triggers
  removeAllDashboardTriggers();
  
  // Add a simple click-to-refresh trigger for the 'Refresh' button only
  // This doesn't auto-update the dashboard but makes the refresh button work
  const dashboardEditTrigger = ScriptApp.newTrigger('handleDashboardEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * Removes all dashboard automatic update triggers
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
 * Handles edit events in the Dashboard sheet
 * Specifically for the refresh button
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
 * Updates dashboard task metrics
 * Used when task statuses change
 */
function updateDashboardTaskMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('Dashboard');
    
    if (!dashboardSheet) return;
    
    // Only re-run the dashboard setup if the dashboard exists
    setupDashboard();
    
  } catch (error) {
    Logger.log(`Error updating dashboard metrics: ${error}`);
  }
}