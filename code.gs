// The ID of the blank Spreadsheet Sandbox (ensure sharing is "Anyone with the link can view")
const TEMPLATE_SHEET_ID = "1pgvPk3DF8ef4R2N_4hfP4LBynq8yoZa8rwqaS5T-EoA";

// 1. Serve the HTML file when the user opens the web app URL
function doGet(e) {
  // --- DEV MODE ENABLED ---
  // We bypass the setupNewUser logic and hardcode the Master Sheet ID
  // This ensures you and Brooke are looking at the exact same testing database.
  const userDbId = TEMPLATE_SHEET_ID; 
  
  // Save it to properties so the rest of the functions can find it
  PropertiesService.getUserProperties().setProperty('USER_DB_ID', userDbId);

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MPP - Daily Check-In')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Function to copy the template to the user's personal drive
function setupNewUser() {
  const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
  
  // Create a copy in the root of their Google Drive
  const newFile = templateFile.makeCopy("My Mood Pattern Planner Data");
  const newFileId = newFile.getId();
  
  // Clear the sample data (A2:H) so the user gets a fresh start, leaving formulas intact
  const newDb = SpreadsheetApp.openById(newFileId);
  const logsSheet = newDb.getSheetByName("Logs");
  if (logsSheet) {
    const lastRow = logsSheet.getLastRow();
    if (lastRow > 1) {
      // Clears Row 2 to the bottom, strictly for Columns A through H (8 columns)
      logsSheet.getRange(2, 1, lastRow - 1, 8).clearContent(); 
    }
  }
  
  // Save this ID to their Google account so the app remembers it next time
  PropertiesService.getUserProperties().setProperty('USER_DB_ID', newFileId);
  
  return newFileId;
}

// 2. Process the data sent from the HTML form
function processForm(formData) {
  try {
    const userDbId = PropertiesService.getUserProperties().getProperty('USER_DB_ID');
    const sheet = SpreadsheetApp.openById(userDbId).getSheetByName("Logs");
    const timestamp = new Date();
    
    // Determine the Applicable Event Time
    let eventTime = timestamp;
    // If frontend passed a customTime (HH:MM), inject it into today's date
    if (formData.customTime) {
      const timeParts = formData.customTime.split(':');
      if (timeParts.length === 2) {
        eventTime.setHours(parseInt(timeParts[0], 10));
        eventTime.setMinutes(parseInt(timeParts[1], 10));
        eventTime.setSeconds(0);
      }
    }
    
    // Auto-calculate Time Segment dynamically from the Engine sheet
    const engineSheet = SpreadsheetApp.openById(userDbId).getSheetByName("Engine");
    const segmentData = engineSheet.getRange("E2:G6").getValues();
    const segmentsConfig = segmentData.map(row => ({
      name: row[0],
      startHour: parseInt(row[1])
    })).sort((a, b) => a.startHour - b.startHour); // Sort chronologically

    const hour = eventTime.getHours();
    let calculatedSegment = segmentsConfig[segmentsConfig.length - 1].name; // Default to the latest segment (Night)
    for (let i = 0; i < segmentsConfig.length; i++) {
      if (hour >= segmentsConfig[i].startHour) {
        calculatedSegment = segmentsConfig[i].name;
      } else {
        break; // If the hour is less than the boundary, the previous segment was correct
      }
    }

    // Structure the row exactly as our sheet headers dictate
    const newRow = [
      eventTime,                // Column A: Applicable Event Time
      calculatedSegment,        // Column B: Auto-calculated Time Segment
      parseInt(formData.mood),  // Column C: Mood
      parseInt(formData.energy),// Column D: Energy
      parseInt(formData.stress),// Column E: Stress
      formData.activity,        // Column F: Activity
      formData.tags,            // Column G: Tags (Optional)
      formData.notes            // Column H: Brief Note (Optional)
    ];
    
    // Find the first completely empty row in Column A to avoid formula overlap
    const columnA = sheet.getRange("A:A").getValues();
    let emptyRowIndex = 1; 
    
    for (let i = 0; i < columnA.length; i++) {
      if (columnA[i][0] === "" || columnA[i][0] === null) {
        emptyRowIndex = i + 1; // +1 because JS arrays are 0-indexed, Sheets are 1-indexed
        break;
      }
    }

    // Push the 8-item array explicitly into A:H of the found empty row
    sheet.getRange(emptyRowIndex, 1, 1, 8).setValues([newRow]);
    
    // Save the Submission Timestamp directly next to the inputs (Column I / 9)
    sheet.getRange(emptyRowIndex, 9).setValue(timestamp);
    
    // Force Google Sheets to calculate the dynamic formulas in Sorted_Logs immediately
    SpreadsheetApp.flush();

    // --- OVERLAY DATA RETRIEVAL FIX ---
    // The math is no longer on the raw 'Logs' sheet. We must fetch it from 'Sorted_Logs'
    const sortedSheet = SpreadsheetApp.openById(userDbId).getSheetByName("Sorted_Logs");
    const sortedData = sortedSheet.getDataRange().getValues();
    
    let targetRow = null;
    // Loop backwards to find the exact entry we just submitted by matching the Submission Time (Now Column I / Index 8)
    for (let i = sortedData.length - 1; i >= 0; i--) {
      if (sortedData[i][8] && new Date(sortedData[i][8]).getTime() === timestamp.getTime()) {
        targetRow = sortedData[i];
        break;
      }
    }
    
    if (targetRow) {
      // Return the specific calculations back to the HTML interface
      return { 
        status: "Success", 
        forceIndex: targetRow[18], // Column S
        truthScore: targetRow[20], // Column U
        alignment: targetRow[22],  // Column W
        risk: targetRow[23]        // Column X
      };
    } else {
      // Failsafe if the row isn't instantly found
      return { status: "Success", alignment: "Calculating...", risk: "Calculating...", forceIndex: 0, truthScore: 10 };
    }

  } catch (error) {
    Logger.log(error);
    return { status: "Error", error: error.message };
  }
}

// Fetch unique activities and tags directly from your historical sheet data
function getUniqueOptions() {
  const userDbId = PropertiesService.getUserProperties().getProperty('USER_DB_ID');
  const sheet = SpreadsheetApp.openById(userDbId).getSheetByName("Logs");
  const data = sheet.getDataRange().getValues();
  
  let tags = new Set();
  let engineActivities = [];
  let segmentsConfig = [];
  
  // Skip row 0 (headers), scan history for previously used inputs (Tags ONLY)
  for (let i = 1; i < data.length; i++) {
    const tagString = data[i][6]; // Column G (Tags)
    if (tagString) {
      let rowTags = tagString.toString().split(',');
      rowTags.forEach(t => {
        if(t.trim() !== "") tags.add(t.trim());
      });
    }
  }
  
  const engineSheet = SpreadsheetApp.openById(userDbId).getSheetByName("Engine");
  
  if (engineSheet) {
    // 1. Fetch Dynamic Activities strictly from Engine K2:K5
    const activityData = engineSheet.getRange("K2:K5").getValues();
    engineActivities = activityData.map(row => row[0].toString().trim()).filter(val => val !== "");
    
    // 2. Fetch Time Segment Boundaries from Engine E2:G6
    const segmentData = engineSheet.getRange("E2:G6").getValues();
    segmentsConfig = segmentData.map(row => ({
      name: row[0],
      startHour: parseInt(row[1]),
      label: row[2]
    })).sort((a, b) => a.startHour - b.startHour); 
  }

  return {
    activities: engineActivities, // Passes exact Engine sequence, no alphabetic sorting!
    tags: Array.from(tags).sort(),
    segments: segmentsConfig
  };
}

// Fetch the most recent check-in logs for the Dashboard Timeline
function getRecentLogs() {
  try {
    const userDbId = PropertiesService.getUserProperties().getProperty('USER_DB_ID');
    // --- CHRONOLOGICAL FIX ---
    // We fetch the data from the dynamically sorted view so that the 
    // OFFSET and SLOPE formulas calculated on the sheet are perfectly accurate.
    const sheet = SpreadsheetApp.openById(userDbId).getSheetByName("Sorted_Logs");
    const data = sheet.getDataRange().getValues();
    
    const validLogs = [];
    
    // Loop through all data starting from row 1 (skipping headers)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // FIX 1: Check Column B (Time Segment) instead of A. 
      // If there is no Time Segment AND no Submission Time (Column X), it is a ghost row.
      if (!row[1] && !row[23]) continue; 
      
      validLogs.push({
        segment: row[1],        // Column B: Time Segment
        mood: row[2],           // Column C: Mood
        energy: row[3],         // Column D: Energy
        stress: row[4],         // Column E: Stress
        activity: row[5],       // Column F: Activity
        tags: row[6],           // Column G: Tags
        notes: row[7],          // Column H: Notes
        
        // --- SHIFTED KINEMATICS & PHYSICS DATA ---
        moodVelocity: row[10],  // Column K: Rate of change
        stressAnomaly: row[11], // Column L: Anomaly Flag (STDEV)
        energyTrend: row[16],   // Column Q: 3-Day Slope Trajectory
        forceIndex: row[18],    // Column S: Weighted Task/Energy Load
        totalEffect: row[19],   // Column T: Force + Stress Overhang
        truthScore: row[20],    // Column U: Raw Mathematical Truth (1-10)
        
        riskIndex: row[21],     // Column V: System Risk Index
        alignment: row[22],     // Column W: Truth Insight
        riskInsight: row[23],   // Column X: Risk Insight
        
        // Submission Time is now sitting in Column I (Index 8)
        timestamp: row[8] ? new Date(row[8]).toISOString() : "" 
      });
    }
    
    // FIX 2: Isolate the last 5 valid logs (representing one full daily cycle), then reverse them
    const recentLogs = validLogs.slice(-5).reverse();
    return recentLogs;
    
  } catch (error) {
    Logger.log("Error fetching logs: " + error);
    return [];
  }
}
