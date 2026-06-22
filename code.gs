// 1. Serve the HTML file when the user opens the web app URL
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MPP - Daily Check-In')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Process the data sent from the HTML form
function processForm(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
    const timestamp = new Date();
    
    // Convert string inputs to numbers for math calculations
    const mood = parseInt(formData.mood);
    const energy = parseInt(formData.energy);
    const stress = parseInt(formData.stress);
    
    // --- PLACEHOLDER MATH LOGIC ---
    // Here we calculate the abstract concepts Brooke asked for.
    // Example Risk Index: High stress + low energy + low mood = high risk.
    let riskIndex = stress - energy + (10 - mood);
    if (riskIndex < 0) riskIndex = 0; // Prevent negative risk
    
    // Placeholder for Truth Score and Alignment
    let truthScore = "Pending Array Match"; 
    let alignmentFlag = (stress > 7 && energy < 4) ? "Distorted" : "Aligned";
    // ------------------------------
    
    // Structure the row exactly as our sheet headers dictate
    const newRow = [
      timestamp,
      formData.timeSegment,
      mood,
      energy,
      stress,
      formData.activity,
      formData.tags,
      formData.notes,
      riskIndex,        // Calculated backend
      truthScore,       // Calculated backend
      alignmentFlag     // Calculated backend
    ];
    
    // Append to the sheet
    sheet.appendRow(newRow);
    
    return "Success";
  } catch (error) {
    return "Error: " + error.toString();
  }
}

// Fetch unique activities and tags directly from your historical sheet data
function getUniqueOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  const data = sheet.getDataRange().getValues();
  
  // Default base options
  let activities = new Set(["Work / Focus", "Admin", "Social", "Rest / Break"]);
  let tags = new Set();
  
  // Skip row 0 (headers), scan history for previously used inputs
  for (let i = 1; i < data.length; i++) {
    const act = data[i][5]; // Column F (Activity)
    const tagString = data[i][6]; // Column G (Tags)
    
    if (act && act.toString().trim() !== "") {
      activities.add(act.toString().trim());
    }
    
    if (tagString) {
      // Split tags by comma in case user enters multiple
      let rowTags = tagString.toString().split(',');
      rowTags.forEach(t => {
        if(t.trim() !== "") tags.add(t.trim());
      });
    }
  }
  
  return {
    activities: Array.from(activities).sort(),
    tags: Array.from(tags).sort()
  };
}
