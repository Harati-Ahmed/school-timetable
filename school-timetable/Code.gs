/**
 * @OnlyCurrentDoc
 * @NotOnlyCurrentDoc
 */

// Store required OAuth scopes
const REQUIRED_SCOPES = [
  'https://www.googleapis.com/auth/script.external_request',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/script.scriptapp'
];

function doGet(e) {
  return HtmlService.createHtmlOutput("Timetable Sync Service is running.");
}

function checkAndSetupPermissions() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const hasPermissions = scriptProperties.getProperty('hasRequiredPermissions');
  
  if (!hasPermissions) {
    // Store the required scopes
    scriptProperties.setProperty('hasRequiredPermissions', 'true');
    scriptProperties.setProperty('oauthScopes', JSON.stringify(REQUIRED_SCOPES));
    
    // Force re-authorization
    ScriptApp.invalidateAuth();
  }
}

function onEdit(e) {
  // Check permissions first
  checkAndSetupPermissions();
  
  // Handle case where function is called without event object
  if (!e) {
    console.log("No event object provided");
    return;
  }

  // Get the active sheet and edited range
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var sheetName = sheet.getName();
  
  console.log("Sheet edited:", sheetName);
  console.log("Range edited:", range.getA1Notation());
  
  // Only process edits in Teachers or Classes sheets
  if (sheetName !== 'Teachers' && sheetName !== 'Classes') {
    console.log("Skipping - not a target sheet");
    return;
  }

  // Skip if not editing a period cell
  var col = range.getColumn();
  console.log("Column edited:", col);
  
  if (sheetName === 'Teachers' && (col < 4 || col > 14 || col === 7 || col === 11)) {
    console.log("Skipping - not a period column in Teachers sheet");
    return;
  }
  if (sheetName === 'Classes' && (col < 3 || col > 13 || col === 6 || col === 10)) {
    console.log("Skipping - not a period column in Classes sheet");
    return;
  }

  // Get the spreadsheet ID from the URL
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // Prepare the payload
  var payload = {
    range: sheetName + '!' + range.getA1Notation(),
    values: [[range.getValue()]]
  };
  
  console.log("Sending payload:", JSON.stringify(payload));

  // Call the backend API
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    console.log("Calling API...");
    var response = UrlFetchApp.fetch('http://localhost:8000/api/sheets/' + spreadsheetId, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    console.log('API Response Code:', responseCode);
    console.log('API Response:', responseText);
  } catch(error) {
    console.error('Error calling API:', error.toString());
  }
} 