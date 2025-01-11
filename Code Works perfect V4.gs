// Constants
const TEACHERS_SHEET_NAME = 'Teachers';
const CLASSES_SHEET_NAME = 'Classes';
const SUMMARY_SHEET_NAME = 'Summary';
const EDIT_IN_PROGRESS = 'EDIT_IN_PROGRESS';
const EDIT_TIMEOUT = 30000; // 30 seconds
const CACHE_KEY = 'TIMETABLE_CACHE_';
const VALIDATION_DELAY = 500; // 500ms delay for validation
const MAX_RETRIES = 3;  // Maximum number of retries for stuck operations

const FIRST_PERIOD = 3;    // Column C (1st period)
const LAST_PERIOD = 13;    // Column M (9th period)
const BREAK_COL = 6;       // Column F (Break)
const LUNCH_COL = 10;      // Column J (Lunch)

// Utility Functions
function showStatus(message, title = 'Status', duration = 5) {
  SpreadsheetApp.getActive().toast(message, title, duration);
  Logger.log(`${title}: ${message}`);
}

function isEditInProgress() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const editStartTime = scriptProperties.getProperty('EDIT_START_TIME');
  if (!editStartTime) return false;
  
  const now = Date.now();
  if (now - parseInt(editStartTime) > EDIT_TIMEOUT) {
    scriptProperties.deleteProperty('EDIT_IN_PROGRESS');
    scriptProperties.deleteProperty('EDIT_START_TIME');
    return false;
  }
  return scriptProperties.getProperty('EDIT_IN_PROGRESS') === 'true';
}

function setEditInProgress() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('EDIT_IN_PROGRESS', 'true');
  scriptProperties.setProperty('EDIT_START_TIME', Date.now().toString());
}

function clearEditInProgress() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('EDIT_IN_PROGRESS');
  scriptProperties.deleteProperty('EDIT_START_TIME');
}

function clearCache() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const cacheKeys = ['Classes', 'Teachers', 'Timestamp'].map(key => CACHE_KEY + key);
  cacheKeys.forEach(key => scriptProperties.deleteProperty(key));
}

function setCacheValue(key, value) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(CACHE_KEY + key, JSON.stringify({
    value: value,
    timestamp: Date.now()
  }));
}

function getCacheValue(key) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const data = scriptProperties.getProperty(CACHE_KEY + key);
  if (!data) return null;
  
  const parsed = JSON.parse(data);
  if (Date.now() - parsed.timestamp > CACHE_DURATION) {
    scriptProperties.deleteProperty(CACHE_KEY + key);
    return null;
  }
  return parsed.value;
}

function batchUpdate(sheet, updates) {
  if (!updates || updates.length === 0) return;
  
  const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  const values = range.getValues();
  
  updates.forEach(update => {
    if (update.row > 0 && update.col > 0 && update.row <= values.length && update.col <= values[0].length) {
      values[update.row - 1][update.col - 1] = update.value;
    }
  });
  
  range.setValues(values);
}

// Create menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Timetable System')
    .addItem('Setup System', 'setupEntireSystem')
    .addItem('Refresh Summary', 'updateSummary')
    .addItem('Validate Timetable', 'validateTimetable')
    .addItem('Clear All Assignments', 'clearAllAssignments')
    .addSeparator()
    .addItem('Backup Data', 'backupData')
    .addItem('Restore Data', 'restoreData')
    .addToUi();
}

function setupEntireSystem() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('No active spreadsheet found. Please open a spreadsheet before running this script.');
    }
    
    setupSheets(ss);
    setupHeaders(ss);
    setupClassesSheet(ss);
    setupDataValidation(ss);
    setupBreakLunchColumns(ss);
    setupSummarySheet(ss);
    setupTriggers();
    
    SpreadsheetApp.getUi().alert('Setup Complete! The timetable system is ready to use.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

function setupSheets(ss) {
  try {
    // Keep track of existing sheets
    const existingSheets = ss.getSheets();
    
    // Create new sheets if they don't exist
    if (!ss.getSheetByName(TEACHERS_SHEET_NAME)) {
      ss.insertSheet(TEACHERS_SHEET_NAME);
    }
    if (!ss.getSheetByName(CLASSES_SHEET_NAME)) {
      ss.insertSheet(CLASSES_SHEET_NAME);
    }
    if (!ss.getSheetByName(SUMMARY_SHEET_NAME)) {
      ss.insertSheet(SUMMARY_SHEET_NAME);
    }
    
    // Delete other sheets
    existingSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (![TEACHERS_SHEET_NAME, CLASSES_SHEET_NAME, SUMMARY_SHEET_NAME].includes(sheetName)) {
        ss.deleteSheet(sheet);
      }
    });
    
    // Reorder sheets
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    
    ss.setActiveSheet(teachersSheet);
    ss.moveActiveSheet(1);
    ss.setActiveSheet(classesSheet);
    ss.moveActiveSheet(2);
    ss.setActiveSheet(summarySheet);
    ss.moveActiveSheet(3);
  } catch (error) {
    throw new Error('Error setting up sheets: ' + error.message);
  }
}

function setupHeaders(ss) {
  // Teachers sheet - Grid layout with timing
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  teachersSheet.clear();
  
  // Set up header rows
  const mainHeader = [
    'SI',
    'Periods ->',
    '1\n08:00-08:50',
    '2\n08:50-09:30',
    '3\n09:30-10:10',
    'Break\n10:10-10:30',
    '4\n10:30-11:10',
    '5\n11:10-11:50',
    '6\n11:50-12:30',
    'Lunch\n12:30-01:00',
    '7\n01:00-01:40',
    '8\n01:40-02:20',
    '9\n02:20-03:00'
  ];
  
  const subHeader = ['', 'Teachers / Subject'];
  
  // Set headers
  teachersSheet.getRange(1, 1, 1, mainHeader.length).setValues([mainHeader]);
  teachersSheet.getRange(2, 1, 1, 2).setValues([subHeader]);
  
  // Format headers
  teachersSheet.getRange(1, 1, 2, mainHeader.length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // Set column widths
  teachersSheet.setColumnWidth(1, 30);  // SI column
  teachersSheet.setColumnWidth(2, 200); // Teachers/Subject column
  for (let i = 3; i <= mainHeader.length; i++) {
    teachersSheet.setColumnWidth(i, 100);
  }
  
  // Set row height for header
  teachersSheet.setRowHeight(1, 60);
  
  // Add initial teacher rows with SI numbers and subjects
  const teachers = [
    ['1', 'Raju Bumb / English'],
    ['2', 'Prabhat Karan / Maths'],
    ['3', 'Shobha Hans / Science'],
    ['4', 'Krishna Naidu / Hindi'],
    ['5', 'Faraz Mangal / SST'],
    ['6', 'Rimi Loke / Sanskrit'],
    ['7', 'Amir Kar / English'],
    ['8', 'Suraj Narayanan / Maths'],
    ['9', 'Alaknanda Chaudry / Science'],
    ['10', 'Preet Mittal / English'],
    ['11', 'John Lalla / English'],
    ['12', 'Ujwal Mohan / Maths'],
    ['13', 'Aadish Mathur / Science'],
    ['14', 'Iqbal Beharry / Hindi'],
    ['15', 'Manjari Shenoy / SST'],
    ['16', 'Aayushi Suri / Sanskrit'],
    ['17', 'Parvez Mathur / SST'],
    ['18', 'Qabool Malhotra / Hindi'],
    ['19', 'Nagma Andra / Sanskrit'],
    ['20', 'Krishna Arora / Hindi'],
    ['21', 'John Lalla / SST'],
    ['22', 'Nitin Banu / Sanskrit'],
    ['23', 'Ananda Debnath / Hindi'],
    ['24', 'Balaram Bhandari / Hindi'],
    ['25', 'Ajay Chaudhri / SST'],
    ['26', 'Niranjan Varma / English'],
    ['27', 'Nur Patel / Maths'],
    ['28', 'Aadish Mathur / English'],
    ['29', 'Nur Patel / Hindi'],
    ['30', 'John Lalla / English'],
    ['31', 'Aadish Mathur / SST']
  ];
  
  teachersSheet.getRange(3, 1, teachers.length, 2).setValues(teachers);
  
  // Format teacher rows
  teachersSheet.getRange(3, 1, teachers.length, mainHeader.length)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // Center align SI column
  teachersSheet.getRange(1, 1, teachers.length + 2, 1).setHorizontalAlignment('center');
  
  // Color the subjects in red
  teachers.forEach((teacher, index) => {
    const cell = teachersSheet.getRange(index + 3, 2);
    const parts = teacher[1].split(' / ');
    cell.setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(teacher[1])
        .setTextStyle(0, parts[0].length + 3, SpreadsheetApp.newTextStyle().build())
        .setTextStyle(parts[0].length + 3, teacher[1].length, SpreadsheetApp.newTextStyle().setForegroundColor('#ff0000').build())
        .build()
    );
  });
  
  // Format headers with thick colored border
  teachersSheet.getRange(1, 1, teachers.length + 2, mainHeader.length)
    .setBorder(
      true, true, true, true, null, null,
      '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
    );
}

function setupClassesSheet(ss) {
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  classesSheet.clear();
  
  // Set up header rows
  const mainHeader = [
    'SI',
    'Periods ->\nClasses',
    '1\n08:00-08:50',
    '2\n08:50-09:30',
    '3\n09:30-10:10',
    'Break\n10:10-10:30',
    '4\n10:30-11:10',
    '5\n11:10-11:50',
    '6\n11:50-12:30',
    'Lunch\n12:30-01:00',
    '7\n01:00-01:40',
    '8\n01:40-02:20',
    '9\n02:20-03:00'
  ];
  
  // Set headers
  classesSheet.getRange(1, 1, 1, mainHeader.length).setValues([mainHeader]);
  
  // Format headers
  classesSheet.getRange(1, 1, 1, mainHeader.length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // Set column widths
  classesSheet.setColumnWidth(1, 30);  // SI column
  classesSheet.setColumnWidth(2, 200); // Classes column
  for (let i = 3; i <= mainHeader.length; i++) {
    classesSheet.setColumnWidth(i, 100);
  }
  
  // Set row height for header
  classesSheet.setRowHeight(1, 60);
  
  // Add classes with SI numbers
  const classes = [
    ['1', 'Nursery'],
    ['2', 'LKG - A'],
    ['3', 'LKG - B'],
    ['4', 'UKG - A'],
    ['5', 'UKG - B'],
    ['6', 'Grade - 1A'],
    ['7', 'Grade - 1B'],
    ['8', 'Grade - 2A'],
    ['9', 'Grade - 2B'],
    ['10', 'Grade - 3A'],
    ['11', 'Grade - 3B'],
    ['12', 'Grade - 4A'],
    ['13', 'Grade - 4B'],
    ['14', 'Grade - 5'],
    ['15', 'Grade - 6'],
    ['16', 'Grade - 7'],
    ['17', 'Grade - 8'],
    ['18', 'Grade - 9'],
    ['19', 'Grade - 10'],
    ['20', 'Grade - 11'],
    ['21', 'Grade - 12']
  ];
  
  // Add classes
  classesSheet.getRange(2, 1, classes.length, 2).setValues(classes);
  
  // Format class rows - add both internal and external borders
  classesSheet.getRange(1, 1, classes.length + 1, mainHeader.length)
    .setBorder(true, true, true, true, true, true) // Add internal borders first
    .setVerticalAlignment('middle');
  
  // Then add the thick outer border
  classesSheet.getRange(1, 1, classes.length + 1, mainHeader.length)
    .setBorder(
      true, true, true, true, null, null, // top, left, bottom, right, vertical, horizontal
      '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  
  // Center align SI column
  classesSheet.getRange(1, 1, classes.length + 1, 1).setHorizontalAlignment('center');
  
  // Update the getUniqueClasses function to match these classes
  return classes.map(c => c[1]);
}

// Add cache management
function clearCache() {
  cachedClasses = null;
  cachedTeachers = null;
  cacheTimestamp = 0;
}

// Update setupDataValidation function
function setupDataValidation(ss) {
  try {
    if (!ss) throw new Error('Spreadsheet not found');
    
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    
    if (!teachersSheet || !classesSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Get the last row for both sheets
    const teachersLastRow = teachersSheet.getLastRow();
    const classesLastRow = classesSheet.getLastRow();
    
    if (teachersLastRow < 3 || classesLastRow < 2) {
      throw new Error('Invalid sheet structure');
    }
    
    // Get all data in one batch operation
    const teachersRange = teachersSheet.getRange(3, FIRST_PERIOD, teachersLastRow - 2, LAST_PERIOD - FIRST_PERIOD + 1);
    const classesRange = classesSheet.getRange(2, FIRST_PERIOD, classesLastRow - 1, LAST_PERIOD - FIRST_PERIOD + 1);
    const teachersAllData = teachersRange.getValues();
    const classesAllData = classesRange.getValues();
    
    // Get names separately
    const teachersNameData = teachersSheet.getRange(3, 2, teachersLastRow - 2, 1).getValues();
    const classesNameData = classesSheet.getRange(2, 2, classesLastRow - 1, 1).getValues();
    
    // Get all available classes and teachers
    const availableClasses = classesNameData.map(row => row[0].trim()).filter(Boolean);
    const availableTeachers = teachersNameData.map(row => row[0].trim()).filter(Boolean);
    
    // Batch process validation rules
    const teacherValidations = [];
    const classValidations = [];
    
    // Process each column
    for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
      if (col === BREAK_COL || col === LUNCH_COL) continue;
      
      const colIndex = col - FIRST_PERIOD;
      
      // Get assignments for this period
      const assignedTeachers = new Map();
      const assignedClasses = new Set();
      
      // Get assignments from both sheets
      teachersAllData.forEach((row, rowIndex) => {
        const classValue = row[colIndex];
        if (classValue) {
          assignedClasses.add(classValue);
          const teacherInfo = teachersNameData[rowIndex][0];
          assignedTeachers.set(classValue, teacherInfo);
        }
      });
      
      // Create validation rules for teachers sheet
      for (let rowIndex = 0; rowIndex < teachersAllData.length; rowIndex++) {
        const currentClass = teachersAllData[rowIndex][colIndex];
        const validClasses = availableClasses.filter(c => 
          !assignedClasses.has(c) || c === currentClass
        );
        
        if (validClasses.length > 0) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(validClasses)
            .setAllowInvalid(false)
            .build();
          teacherValidations.push({
            range: teachersSheet.getRange(rowIndex + 3, col),
            rule: rule
          });
        }
      }
      
      // Create validation rules for classes sheet
      for (let rowIndex = 0; rowIndex < classesAllData.length; rowIndex++) {
        const currentTeacher = classesAllData[rowIndex][colIndex];
        const currentTeacherName = currentTeacher ? currentTeacher.split(' / ')[0] : '';
        
        // Get all assigned teacher names for this period
        const assignedTeacherNames = new Set();
        classesAllData.forEach((row, i) => {
          if (i !== rowIndex && row[colIndex]) {
            const [teacherName] = row[colIndex].split(' / ');
            assignedTeacherNames.add(teacherName);
          }
        });
        
        // Show all available teachers for the current class
        const validTeachers = availableTeachers.filter(t => {
          const [teacherName] = t.split(' / ');
          return !assignedTeacherNames.has(teacherName) || teacherName === currentTeacherName;
        });
        
        if (validTeachers.length > 0) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(validTeachers)
            .setAllowInvalid(false)
            .build();
          classValidations.push({
            range: classesSheet.getRange(rowIndex + 2, col),
            rule: rule
          });
        }
      }
    }
    
    // Apply all validation rules in batch
    teacherValidations.forEach(validation => {
      validation.range.setDataValidation(validation.rule);
    });
    
    classValidations.forEach(validation => {
      validation.range.setDataValidation(validation.rule);
    });
    
    showStatus('Data validation rules have been updated', 'Validation Update');
  } catch (error) {
    showStatus('Error updating validation rules: ' + error.message, 'Validation Error');
    Logger.log('Error in setupDataValidation: ' + error);
  }
}

function setupBreakLunchColumns(ss) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Get the last row of each sheet
  const teachersLastRow = teachersSheet.getLastRow();
  const classesLastRow = classesSheet.getLastRow();
  
  // Color Break column
  teachersSheet.getRange(1, BREAK_COL, teachersLastRow).setBackground('#ffcdd2');
  classesSheet.getRange(1, BREAK_COL, classesLastRow).setBackground('#ffcdd2');
  
  // Color Lunch column
  teachersSheet.getRange(1, LUNCH_COL, teachersLastRow).setBackground('#ffcdd2');
  classesSheet.getRange(1, LUNCH_COL, classesLastRow).setBackground('#ffcdd2');
}

function setupTriggers() {
  // Delete existing triggers with error handling
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    try {
      ScriptApp.deleteTrigger(trigger);
    } catch (error) {
      Logger.log('Error deleting trigger: ' + error);
    }
  });
  
  // Create new edit trigger with error handling
  try {
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
    showStatus('Triggers set up successfully', 'Setup Complete');
  } catch (error) {
    showStatus('Error setting up triggers: ' + error.message, 'Setup Error');
    Logger.log('Error in setupTriggers: ' + error);
  }
}

// Cache for unique classes and teachers
let cachedClasses = null;
let cachedTeachers = null;
let cacheTimestamp = 0;
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

// Optimize getUniqueClasses with caching
function getUniqueClasses() {
  const now = Date.now();
  if (cachedClasses && now - cacheTimestamp < CACHE_DURATION) {
    return cachedClasses;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const lastRow = classesSheet.getLastRow();
  
  const classData = classesSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  cachedClasses = Array.from(new Set(classData.map(row => row[0]).filter(Boolean)));
  cacheTimestamp = now;
  return cachedClasses;
}

// Optimize getUniqueTeachers with caching
function getUniqueTeachers(ss) {
  const now = Date.now();
  if (cachedTeachers && now - cacheTimestamp < CACHE_DURATION) {
    return cachedTeachers;
  }
  
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const lastRow = teachersSheet.getLastRow();
  if (lastRow < 3) return [];
  
  const teacherData = teachersSheet.getRange(3, 2, lastRow - 2, 1).getValues();
  cachedTeachers = Array.from(new Set(teacherData.map(row => row[0]).filter(Boolean)));
  cacheTimestamp = now;
  return cachedTeachers;
}

// Add this function to set up the Summary sheet
function setupSummarySheet(ss) {
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  summarySheet.clear();
  
  // Set up headers for Teachers table (left side)
  const teachersHeaders = [
    ['Teachers Summary'],
    ['SI', 'Teacher Name', 'Total\nPeriods']
  ];
  
  // Set up headers for Classwise Subject Allotment table (right side)
  const classHeaders = [
    ['Classwise Subject Allotment'],
    ['SI', 'Classes', 'English', 'Maths', 'Science', 'Hindi', 'SST', 'Sanskrit']
  ];
  
  // Write headers
  summarySheet.getRange(1, 1, 1, 3).merge().setValue(teachersHeaders[0][0]);
  summarySheet.getRange(2, 1, 1, 3).setValues([teachersHeaders[1]]);
  
  const rightTableWidth = classHeaders[1].length;
  summarySheet.getRange(1, 5, 1, rightTableWidth).merge().setValue(classHeaders[0][0]);
  summarySheet.getRange(2, 5, 1, rightTableWidth).setValues([classHeaders[1]]);
  
  // Format headers
  const headerRanges = [
    summarySheet.getRange(1, 1, 1, 3),  // Teachers main header
    summarySheet.getRange(2, 1, 1, 3),  // Teachers subheader
    summarySheet.getRange(1, 5, 1, rightTableWidth),  // Classes main header
    summarySheet.getRange(2, 5, 1, rightTableWidth)   // Classes subheader
  ];
  
  headerRanges.forEach(range => {
    range.setBackground('#f3f3f3')
         .setFontWeight('bold')
         .setBorder(true, true, true, true, true, true)
         .setWrap(true)
         .setVerticalAlignment('middle')
         .setHorizontalAlignment('center')
         .setBorder(
           true, true, true, true, null, null,
           '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
         );
  });
  
  // Set column widths
  summarySheet.setColumnWidth(1, 30);   // SI
  summarySheet.setColumnWidth(2, 150);  // Teacher Name
  summarySheet.setColumnWidth(3, 80);   // Total Periods
  summarySheet.setColumnWidth(4, 30);   // Gap between tables
  summarySheet.setColumnWidth(5, 30);   // SI
  summarySheet.setColumnWidth(6, 150);  // Classes
  for (let i = 7; i < 7 + 6; i++) {    // Subject columns
    summarySheet.setColumnWidth(i, 100);
  }
  
  // Set row heights
  summarySheet.setRowHeight(1, 30);
  summarySheet.setRowHeight(2, 40);
  
  // Initialize with specific row counts
  const TEACHER_ROWS = 26;  // Fixed number of rows for Teachers table
  const classes = getUniqueClasses();
  const CLASS_ROWS = 21;    // Fixed number of rows for Classes table
  
  // Add SI numbers and classes to the right table (only 21 rows)
  const classData = Array.from({length: CLASS_ROWS}, (_, i) => [i + 1, i < classes.length ? classes[i] : '']);
  summarySheet.getRange(3, 5, CLASS_ROWS, 2).setValues(classData);
  
  // Add SI numbers to the left table (only 26 rows)
  const siData = Array.from({length: TEACHER_ROWS}, (_, i) => [i + 1]);
  summarySheet.getRange(3, 1, TEACHER_ROWS, 1).setValues(siData);
  
  // Format data areas
  const leftTable = summarySheet.getRange(3, 1, TEACHER_ROWS, 3);
  const rightTable = summarySheet.getRange(3, 5, CLASS_ROWS, rightTableWidth);
  
  [leftTable, rightTable].forEach(range => {
    range.setBorder(true, true, true, true, true, true)
         .setVerticalAlignment('middle')
         .setBackground('white')
         .setFontSize(10);
  });
  
  // Set alignments for data areas
  summarySheet.getRange(3, 1, TEACHER_ROWS, 1).setHorizontalAlignment('center'); // Left SI
  summarySheet.getRange(3, 2, TEACHER_ROWS, 1).setHorizontalAlignment('left');   // Teacher names
  summarySheet.getRange(3, 3, TEACHER_ROWS, 1).setHorizontalAlignment('center'); // Total periods
  summarySheet.getRange(3, 5, CLASS_ROWS, 1).setHorizontalAlignment('center');   // Right SI
  summarySheet.getRange(3, 6, CLASS_ROWS, 1).setHorizontalAlignment('left');     // Class names
  summarySheet.getRange(3, 7, CLASS_ROWS, 6).setHorizontalAlignment('center');   // Subject columns
  
  // Add alternating row colors
  for (let i = 0; i < TEACHER_ROWS; i++) {
    const rowNumber = i + 3;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    summarySheet.getRange(rowNumber, 1, 1, 3).setBackground(color);
  }
  
  for (let i = 0; i < CLASS_ROWS; i++) {
    const rowNumber = i + 3;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    summarySheet.getRange(rowNumber, 5, 1, rightTableWidth).setBackground(color);
  }
  
  // Update the summary data immediately
  updateSummary();
}

// Add the onEdit trigger function
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    // Only process edits in Teachers or Classes sheets
    if (sheet.getName() !== TEACHERS_SHEET_NAME && sheet.getName() !== CLASSES_SHEET_NAME) {
      return;
    }
    
    // Only process edits in the valid period columns
    if (col < FIRST_PERIOD || col > LAST_PERIOD || col === BREAK_COL || col === LUNCH_COL) {
      return;
    }
    
    // Force clear any stuck edit flags that are older than the timeout
    const scriptProperties = PropertiesService.getScriptProperties();
    const editStartTime = scriptProperties.getProperty('EDIT_START_TIME');
    if (editStartTime && Date.now() - parseInt(editStartTime) > EDIT_TIMEOUT) {
      clearEditInProgress();
    }
    
    // Retry logic for stuck edits
    let retryCount = 0;
    while (isEditInProgress() && retryCount < MAX_RETRIES) {
      Utilities.sleep(VALIDATION_DELAY);
      retryCount++;
      if (retryCount === MAX_RETRIES) {
        clearEditInProgress(); // Force clear if still stuck after max retries
        showStatus('Cleared stuck edit lock after ' + MAX_RETRIES + ' retries', 'Edit Recovery');
      }
    }
    
    // Set edit in progress flag with timestamp
    setEditInProgress();
    
    try {
      // Handle multiple cell operations
      if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
        const oldValues = range.getValues();
        const updates = [];
        
        // Process each cell in the range
        for (let i = 0; i < range.getNumRows(); i++) {
          for (let j = 0; j < range.getNumColumns(); j++) {
            const currentRow = row + i;
            const currentCol = col + j;
            
            // Skip break and lunch columns
            if (currentCol === BREAK_COL || currentCol === LUNCH_COL) continue;
            
            // Create a single-cell event object
            const cellEvent = {
              source: e.source,
              range: sheet.getRange(currentRow, currentCol),
              value: '',  // For deletion
              oldValue: oldValues[i][j],
              row: currentRow
            };
            
            // Only process if there was a value in the cell
            if (cellEvent.oldValue && cellEvent.oldValue.toString().trim() !== '') {
              if (sheet.getName() === TEACHERS_SHEET_NAME && currentRow > 2) {
                updates.push({event: cellEvent, type: 'teachers'});
              } else if (sheet.getName() === CLASSES_SHEET_NAME && currentRow >= 2) {
                updates.push({event: cellEvent, type: 'classes'});
              }
            }
          }
        }
        
        // Process all updates in batch
        if (updates.length > 0) {
          // Sort updates to process teachers first, then classes
          updates.sort((a, b) => a.type === 'teachers' ? -1 : 1);
          
          updates.forEach(update => {
            if (update.type === 'teachers') {
              updateClassesFromTeachers(update.event);
            } else {
              updateTeachersFromClasses(update.event);
            }
          });
        }
      }
      // Handle single cell edit
      else {
        if (sheet.getName() === TEACHERS_SHEET_NAME && row > 2) {
          updateClassesFromTeachers({
            ...e,
            row: row
          });
        } else if (sheet.getName() === CLASSES_SHEET_NAME && row >= 2) {
          updateTeachersFromClasses({
            ...e,
            row: row
          });
        }
      }
      
      // Clear cache after successful edit
      clearCache();
      
      // Add a small delay before updating validation to prevent rapid consecutive updates
      Utilities.sleep(VALIDATION_DELAY);
      
      // Update validation and summary
      setupDataValidation(e.source);
      updateSummary();
    } finally {
      // Ensure we always clear the edit flag
      clearEditInProgress();
    }
    
  } catch (error) {
    clearEditInProgress();
    showStatus('Error processing edit: ' + error.message, 'Edit Error');
    Logger.log('Error in onEdit: ' + error);
  }
}

// Improve the synchronization functions
function updateClassesFromTeachers(e) {
  if (!e || !e.source) {
    showStatus('Invalid event object in updateClassesFromTeachers', 'Sync Error');
    return;
  }
  
  const ss = e.source;
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  if (!teachersSheet || !classesSheet) {
    showStatus('Required sheets not found', 'Sync Error');
    return;
  }
  
  const range = e.range;
  const row = e.row;
  const col = range.getColumn();
  const value = e.value || '';
  const oldValue = e.oldValue || '';
  
  try {
    // Get teacher name and subject
    const teacherInfo = teachersSheet.getRange(row, 2).getValue().toString().trim();
    if (!teacherInfo.includes(' / ')) {
      showStatus('Invalid teacher format: ' + teacherInfo, 'Sync Error');
      return;
    }
    
    const [teacherName, subject] = teacherInfo.split(' / ');
    
    // Find matching class row
    const classesRange = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1);
    const classesData = classesRange.getValues();
    let classRow = -1;
    
    // If this is a deletion
    if (oldValue && !value) {
      for (let i = 0; i < classesData.length; i++) {
        if (classesData[i][0].trim() === oldValue.trim()) {
          classRow = i + 2;
          break;
        }
      }
      if (classRow !== -1) {
        const currentValue = classesSheet.getRange(classRow, col).getValue();
        if (currentValue.includes(teacherName)) {
          classesSheet.getRange(classRow, col).clearContent();
        }
      }
    }
    // If this is a new value
    else if (value) {
      for (let i = 0; i < classesData.length; i++) {
        if (classesData[i][0].trim() === value.trim()) {
          classRow = i + 2;
          break;
        }
      }
      if (classRow !== -1) {
        classesSheet.getRange(classRow, col).setValue(`${teacherName} / ${subject}`);
      }
    }
  } catch (error) {
    showStatus('Error in updateClassesFromTeachers: ' + error.message, 'Sync Error');
    Logger.log('Error in updateClassesFromTeachers: ' + error);
  }
}

function updateTeachersFromClasses(e) {
  if (!e || !e.source) {
    showStatus('Invalid event object in updateTeachersFromClasses', 'Sync Error');
    return;
  }
  
  const ss = e.source;
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  if (!teachersSheet || !classesSheet) {
    showStatus('Required sheets not found', 'Sync Error');
    return;
  }
  
  const range = e.range;
  const row = e.row;
  const col = range.getColumn();
  const value = e.value || '';
  const oldValue = e.oldValue || '';
  
  try {
    // Get class name
    const className = classesSheet.getRange(row, 2).getValue().toString().trim();
    if (!className) {
      showStatus('Invalid class name', 'Sync Error');
      return;
    }
    
    // Find matching teacher row
    const teachersRange = teachersSheet.getRange(3, 2, teachersSheet.getLastRow() - 2, 1);
    const teachersData = teachersRange.getValues();
    let teacherRow = -1;
    
    // If this is a deletion
    if (oldValue) {
      const [oldTeacherName] = oldValue.split(' / ');
      for (let i = 0; i < teachersData.length; i++) {
        const [currentTeacher] = teachersData[i][0].split(' / ');
        if (currentTeacher.trim() === oldTeacherName.trim()) {
          teacherRow = i + 3;
          break;
        }
      }
      if (teacherRow !== -1) {
        const currentValue = teachersSheet.getRange(teacherRow, col).getValue();
        if (currentValue === className) {
          teachersSheet.getRange(teacherRow, col).clearContent();
        }
      }
    }
    // If this is a new value
    else if (value) {
      const [newTeacherName] = value.split(' / ');
      for (let i = 0; i < teachersData.length; i++) {
        const [currentTeacher] = teachersData[i][0].split(' / ');
        if (currentTeacher.trim() === newTeacherName.trim()) {
          teacherRow = i + 3;
          break;
        }
      }
      if (teacherRow !== -1) {
        teachersSheet.getRange(teacherRow, col).setValue(className);
      }
    }
  } catch (error) {
    showStatus('Error in updateTeachersFromClasses: ' + error.message, 'Sync Error');
    Logger.log('Error in updateTeachersFromClasses: ' + error);
  }
}

// Add summary update function
function updateSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet not found');
    
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    
    if (!teachersSheet || !summarySheet || !classesSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Get all teachers and their data
    const lastRow = teachersSheet.getLastRow();
    if (lastRow < 3) {
      setupSummarySheet(ss);
      return;
    }
    
    const teacherRows = teachersSheet.getRange(3, 2, lastRow - 2, LAST_PERIOD - 1).getValues();
    
    // Create a map to store unique teachers and their total periods
    const teacherMap = new Map();
    
    // Process teacher data and combine duplicates
    teacherRows.forEach(row => {
      if (!row[0]) return; // Skip empty rows
      
      const teacherInfo = row[0].split(' / ');
      const teacher = teacherInfo[0];
      const subject = teacherInfo[1];
      let totalPeriods = 0;
      
      // Count periods
      for (let i = 1; i < row.length; i++) {
        if (row[i] && i !== BREAK_COL - 2 && i !== LUNCH_COL - 2) {
          totalPeriods++;
        }
      }
      
      // Add periods to existing teacher or create new entry
      if (teacherMap.has(teacher)) {
        const existing = teacherMap.get(teacher);
        existing.periods += totalPeriods;
        existing.subjects.add(subject);
      } else {
        teacherMap.set(teacher, {
          periods: totalPeriods,
          subjects: new Set([subject])
        });
      }
    });
    
    // Convert map to array and sort by total periods (descending), then by name
    const sortedTeachers = Array.from(teacherMap.entries())
      .sort((a, b) => {
        const periodDiff = b[1].periods - a[1].periods;
        return periodDiff !== 0 ? periodDiff : a[0].localeCompare(b[0]);
      });
    
    // Create the summary data array
    const teacherSummaryData = sortedTeachers.map((entry, index) => {
      const teacher = entry[0];
      const subjects = Array.from(entry[1].subjects).join(', ');
      return [
        index + 1,                          // SI
        `${teacher} (${subjects})`,         // Teacher name with subjects
        entry[1].periods                    // Total periods
      ];
    });
    
    // Write teacher summary data (left side)
    if (teacherSummaryData.length > 0) {
      summarySheet.getRange(3, 2, 26, 2).clearContent();
      const rowsToWrite = Math.min(teacherSummaryData.length, 26);
      for (let i = 0; i < rowsToWrite; i++) {
        summarySheet.getRange(3 + i, 2, 1, 2).setValues([[teacherSummaryData[i][1], teacherSummaryData[i][2]]]);
      }
    } else {
      summarySheet.getRange(3, 2, 26, 2).clearContent();
    }
    
    // Process class-subject assignments (right side)
    const subjects = ['English', 'Maths', 'Science', 'Hindi', 'SST', 'Sanskrit'];
    const classes = getUniqueClasses();
    
    // Create a map to store subject counts for each class
    const classSubjectCountMap = new Map();
    
    // Initialize the map for all classes
    classes.forEach(className => {
      classSubjectCountMap.set(className, new Map());
      subjects.forEach(subject => {
        classSubjectCountMap.get(className).set(subject, 0);
      });
    });
    
    // Count periods for each subject in each class
    const classesData = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, LAST_PERIOD - 1).getValues();
    
    classesData.forEach((row, classIndex) => {
      const className = row[0];
      if (!className) return;
      
      // Start from column 1 (skipping class name in column 0)
      for (let i = 1; i < row.length; i++) {
        if (i !== BREAK_COL - 2 && i !== LUNCH_COL - 2) {
          const cellValue = row[i];
          if (cellValue) {
            const subject = cellValue.split(' / ')[1];
            if (subjects.includes(subject)) {
              const currentCount = classSubjectCountMap.get(className).get(subject);
              classSubjectCountMap.get(className).set(subject, currentCount + 1);
            }
          }
        }
      }
    });
    
    // Create the class-subject matrix with counts
    const subjectCountsData = classes.map(className => {
      const subjectCounts = classSubjectCountMap.get(className);
      return subjects.map(subject => subjectCounts.get(subject));
    });
    
    // Write class-subject data
    if (subjectCountsData.length > 0) {
      summarySheet.getRange(3, 7, 21, subjects.length).clearContent();
      const rowsToWrite = Math.min(subjectCountsData.length, 21);
      summarySheet.getRange(3, 7, rowsToWrite, subjects.length)
        .setValues(subjectCountsData.slice(0, rowsToWrite));
    }
    
    // Format the summary sheet
    formatSummarySheet(summarySheet);
    
    showStatus('Summary updated successfully', 'Summary Update');
  } catch (error) {
    showStatus('Error updating summary: ' + error.message, 'Summary Error');
    Logger.log('Error in updateSummary: ' + error);
  }
}

// Helper function to format the summary sheet
function formatSummarySheet(summarySheet) {
  // Format left table (26 rows)
  const leftTable = summarySheet.getRange(3, 1, 26, 3);
  leftTable.setBorder(true, true, true, true, true, true);
  leftTable.setVerticalAlignment('middle');
  
  // Format right table (21 rows)
  const rightTable = summarySheet.getRange(3, 5, 21, 8);
  rightTable.setBorder(true, true, true, true, true, true);
  rightTable.setVerticalAlignment('middle');
  
  // Add thick colored borders to both tables
  // Left table thick border
  summarySheet.getRange(2, 1, 27, 3).setBorder(
    true, true, true, true, null, null, // top, left, bottom, right, vertical, horizontal
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK // color, style
  );
  
  // Right table thick border
  summarySheet.getRange(2, 5, 22, 8).setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Set alignments
  summarySheet.getRange(3, 1, 26, 1).setHorizontalAlignment('center'); // Left SI
  summarySheet.getRange(3, 2, 26, 1).setHorizontalAlignment('left');   // Teacher names
  summarySheet.getRange(3, 3, 26, 1).setHorizontalAlignment('center'); // Total periods
  summarySheet.getRange(3, 5, 21, 1).setHorizontalAlignment('center'); // Right SI
  summarySheet.getRange(3, 6, 21, 1).setHorizontalAlignment('left');   // Class names
  summarySheet.getRange(3, 7, 21, 6).setHorizontalAlignment('center'); // Subject columns
  
  // Add alternating row colors
  for (let i = 0; i < 26; i++) {
    const rowNumber = i + 3;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    summarySheet.getRange(rowNumber, 1, 1, 3).setBackground(color);
  }
  
  for (let i = 0; i < 21; i++) {
    const rowNumber = i + 3;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    summarySheet.getRange(rowNumber, 5, 1, 8).setBackground(color);
  }
}

// Add this function to your existing code
function validateTimetable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  const errors = [];
  const warnings = [];
  
  try {
    // Check sheet structure
    if (!teachersSheet || !classesSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Check teacher format and subject validity
    const validSubjects = ['English', 'Maths', 'Science', 'Hindi', 'SST', 'Sanskrit'];
    const teachersRange = teachersSheet.getRange(3, 2, teachersSheet.getLastRow() - 2, 1);
    const teachersData = teachersRange.getValues();
    teachersData.forEach((row, index) => {
      const value = row[0].toString().trim();
      if (value) {
        if (!value.includes(' / ')) {
          errors.push(`Invalid teacher format in row ${index + 3}: ${value}`);
        } else {
          const [, subject] = value.split(' / ');
          if (!validSubjects.includes(subject)) {
            errors.push(`Invalid subject "${subject}" in row ${index + 3}`);
          }
        }
      }
    });
    
    // Check for duplicate teacher names
    const teacherNames = new Set();
    const duplicateTeachers = new Set();
    teachersData.forEach((row, index) => {
      const value = row[0].toString().trim();
      if (value) {
        const [teacherName] = value.split(' / ');
        if (teacherNames.has(teacherName)) {
          duplicateTeachers.add(teacherName);
        }
        teacherNames.add(teacherName);
      }
    });
    
    if (duplicateTeachers.size > 0) {
      warnings.push('Duplicate teacher names found: ' + Array.from(duplicateTeachers).join(', '));
    }
    
    // Check for assignments in break/lunch periods
    for (const col of [BREAK_COL, LUNCH_COL]) {
      // Check teachers sheet
      const teacherBreakData = teachersSheet.getRange(3, col, teachersSheet.getLastRow() - 2, 1).getValues();
      teacherBreakData.forEach((cell, index) => {
        if (cell[0]) {
          errors.push(`Assignment found in ${col === BREAK_COL ? 'break' : 'lunch'} period for teacher in row ${index + 3}`);
        }
      });
      
      // Check classes sheet
      const classBreakData = classesSheet.getRange(2, col, classesSheet.getLastRow() - 1, 1).getValues();
      classBreakData.forEach((cell, index) => {
        if (cell[0]) {
          errors.push(`Assignment found in ${col === BREAK_COL ? 'break' : 'lunch'} period for class in row ${index + 2}`);
        }
      });
    }
    
    // Check for duplicate assignments in each period
    for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
      if (col === BREAK_COL || col === LUNCH_COL) continue;
      
      const teacherAssignments = new Map();
      const classAssignments = new Map();
      
      // Check teachers sheet
      const teacherCol = teachersSheet.getRange(3, col, teachersSheet.getLastRow() - 2, 1).getValues();
      teacherCol.forEach((cell, index) => {
        const value = cell[0];
        if (value) {
          if (classAssignments.has(value)) {
            errors.push(`Class "${value}" assigned multiple times in period ${col - 2}`);
          }
          classAssignments.set(value, index + 3);
        }
      });
      
      // Check classes sheet
      const classCol = classesSheet.getRange(2, col, classesSheet.getLastRow() - 1, 1).getValues();
      classCol.forEach((cell, index) => {
        const value = cell[0];
        if (value) {
          const [teacher] = value.split(' / ');
          if (teacherAssignments.has(teacher)) {
            errors.push(`Teacher "${teacher}" assigned multiple times in period ${col - 2}`);
          }
          teacherAssignments.set(teacher, index + 2);
        }
      });
    }
    
    // Check for inconsistencies between teachers and classes sheets
    for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
      if (col === BREAK_COL || col === LUNCH_COL) continue;
      
      const teacherAssignments = new Map();
      const classAssignments = new Map();
      
      // Get teacher assignments
      const teacherCol = teachersSheet.getRange(3, col, teachersSheet.getLastRow() - 2, 1).getValues();
      const teacherInfoCol = teachersSheet.getRange(3, 2, teachersSheet.getLastRow() - 2, 1).getValues();
      teacherCol.forEach((cell, index) => {
        const className = cell[0];
        if (className) {
          const teacherInfo = teacherInfoCol[index][0];
          teacherAssignments.set(className, teacherInfo);
        }
      });
      
      // Get class assignments
      const classCol = classesSheet.getRange(2, col, classesSheet.getLastRow() - 1, 1).getValues();
      const classNameCol = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1).getValues();
      classCol.forEach((cell, index) => {
        const teacherInfo = cell[0];
        if (teacherInfo) {
          const className = classNameCol[index][0];
          classAssignments.set(className, teacherInfo);
        }
      });
      
      // Check for inconsistencies
      for (const [className, teacherInfo] of teacherAssignments) {
        const classAssignment = classAssignments.get(className);
        if (!classAssignment) {
          errors.push(`Inconsistency in period ${col - 2}: Class "${className}" is assigned in teachers sheet but not in classes sheet`);
        } else {
          const [teacher1] = teacherInfo.split(' / ');
          const [teacher2] = classAssignment.split(' / ');
          if (teacher1 !== teacher2) {
            errors.push(`Inconsistency in period ${col - 2}: Class "${className}" has different teacher assignments`);
          }
        }
      }
    }
    
    // Show results
    const ui = SpreadsheetApp.getUi();
    if (errors.length > 0 || warnings.length > 0) {
      let message = '';
      if (errors.length > 0) {
        message += 'ERRORS:\n' + errors.join('\n') + '\n\n';
      }
      if (warnings.length > 0) {
        message += 'WARNINGS:\n' + warnings.join('\n');
      }
      ui.alert('Validation Results', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Validation Successful', 'No errors or warnings found in the timetable.', ui.ButtonSet.OK);
    }
  } catch (error) {
    showStatus('Error during validation: ' + error.message, 'Validation Error');
    Logger.log('Error in validateTimetable: ' + error);
  }
}

// Add clear function
function clearAllAssignments() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Assignments',
    'Are you sure you want to clear all assignments? This cannot be undone.',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Clear teachers assignments
  teachersSheet.getRange(3, FIRST_PERIOD, teachersSheet.getLastRow() - 2, LAST_PERIOD - FIRST_PERIOD + 1).clearContent();
  
  // Clear classes assignments
  classesSheet.getRange(2, FIRST_PERIOD, classesSheet.getLastRow() - 1, LAST_PERIOD - FIRST_PERIOD + 1).clearContent();
  
  // Update validation and summary
  setupDataValidation(ss);
  updateSummary();
  
  ui.alert('Clear Complete', 'All assignments have been cleared.', ui.ButtonSet.OK);
}

// Improve backup and restore functions
function backupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  try {
    // Validate sheets exist
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    
    if (!teachersSheet || !classesSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Get sheet data with error handling
    let teachersData, classesData;
    try {
      teachersData = teachersSheet.getRange(1, 1, teachersSheet.getLastRow(), teachersSheet.getLastColumn()).getValues();
    } catch (error) {
      throw new Error('Failed to read teachers data: ' + error.message);
    }
    
    try {
      classesData = classesSheet.getRange(1, 1, classesSheet.getLastRow(), classesSheet.getLastColumn()).getValues();
    } catch (error) {
      throw new Error('Failed to read classes data: ' + error.message);
    }
    
    // Validate data
    if (!teachersData || !classesData || teachersData.length === 0 || classesData.length === 0) {
      throw new Error('Invalid sheet data');
    }
    
    // Create backup object with metadata
    const backup = {
      version: '1.0',
      timestamp: new Date().toISOString(),
      teachers: teachersData,
      classes: classesData,
      metadata: {
        teachersRows: teachersData.length,
        teachersCols: teachersData[0].length,
        classesRows: classesData.length,
        classesCols: classesData[0].length
      }
    };
    
    // Convert to JSON and split into chunks
    const backupJson = JSON.stringify(backup);
    const chunkSize = 8000; // Leave room for property name
    const chunks = [];
    
    for (let i = 0; i < backupJson.length; i += chunkSize) {
      chunks.push(backupJson.slice(i, i + chunkSize));
    }
    
    // Clear old backup data
    clearBackupData();
    
    // Store backup metadata
    scriptProperties.setProperty('BACKUP_VERSION', backup.version);
    scriptProperties.setProperty('BACKUP_TIMESTAMP', backup.timestamp);
    scriptProperties.setProperty('BACKUP_CHUNKS', chunks.length.toString());
    
    // Store chunks with error handling
    let storedChunks = 0;
    for (let i = 0; i < chunks.length; i++) {
      try {
        scriptProperties.setProperty(`BACKUP_CHUNK_${i}`, chunks[i]);
        storedChunks++;
      } catch (error) {
        clearBackupData();
        throw new Error(`Failed to store chunk ${i + 1}/${chunks.length}: ${error.message}`);
      }
    }
    
    showStatus(`Backup completed successfully (${storedChunks} chunks)`, 'Backup Complete');
  } catch (error) {
    showStatus('Backup failed: ' + error.message, 'Backup Error');
    Logger.log('Error in backupData: ' + error);
  }
}

function clearBackupData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  try {
    // Get all properties
    const props = scriptProperties.getProperties();
    
    // Delete backup-related properties
    Object.keys(props).forEach(key => {
      if (key.startsWith('BACKUP_')) {
        scriptProperties.deleteProperty(key);
      }
    });
  } catch (error) {
    Logger.log('Error in clearBackupData: ' + error);
  }
}

function restoreData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // Check if backup exists
    const backupVersion = scriptProperties.getProperty('BACKUP_VERSION');
    const backupTimestamp = scriptProperties.getProperty('BACKUP_TIMESTAMP');
    const chunkCount = parseInt(scriptProperties.getProperty('BACKUP_CHUNKS') || '0');
    
    if (!backupVersion || !backupTimestamp || chunkCount === 0) {
      throw new Error('No valid backup found');
    }
    
    // Show confirmation with backup details
    const response = ui.alert(
      'Restore Data',
      `Are you sure you want to restore from the backup created on ${new Date(backupTimestamp).toLocaleString()}?\n` +
      'Current data will be lost.',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    // Reconstruct backup data from chunks
    let backupJson = '';
    for (let i = 0; i < chunkCount; i++) {
      const chunk = scriptProperties.getProperty(`BACKUP_CHUNK_${i}`);
      if (!chunk) {
        throw new Error(`Backup data is corrupted (missing chunk ${i + 1}/${chunkCount})`);
      }
      backupJson += chunk;
    }
    
    const backup = JSON.parse(backupJson);
    
    // Validate backup data
    if (!backup.version || !backup.timestamp || !backup.teachers || !backup.classes || !backup.metadata) {
      throw new Error('Invalid backup data format');
    }
    
    // Validate data dimensions
    if (backup.teachers.length !== backup.metadata.teachersRows ||
        backup.teachers[0].length !== backup.metadata.teachersCols ||
        backup.classes.length !== backup.metadata.classesRows ||
        backup.classes[0].length !== backup.metadata.classesCols) {
      throw new Error('Backup data dimensions mismatch');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Restore teachers data
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    if (!teachersSheet) throw new Error('Teachers sheet not found');
    
    // Clear and resize sheet before restoring
    teachersSheet.clear();
    if (teachersSheet.getMaxRows() < backup.teachers.length) {
      teachersSheet.insertRows(1, backup.teachers.length - teachersSheet.getMaxRows());
    }
    if (teachersSheet.getMaxColumns() < backup.teachers[0].length) {
      teachersSheet.insertColumns(1, backup.teachers[0].length - teachersSheet.getMaxColumns());
    }
    
    // Restore data
    teachersSheet.getRange(1, 1, backup.teachers.length, backup.teachers[0].length).setValues(backup.teachers);
    
    // Restore classes data
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    if (!classesSheet) throw new Error('Classes sheet not found');
    
    // Clear and resize sheet before restoring
    classesSheet.clear();
    if (classesSheet.getMaxRows() < backup.classes.length) {
      classesSheet.insertRows(1, backup.classes.length - classesSheet.getMaxRows());
    }
    if (classesSheet.getMaxColumns() < backup.classes[0].length) {
      classesSheet.insertColumns(1, backup.classes[0].length - classesSheet.getMaxColumns());
    }
    
    // Restore data
    classesSheet.getRange(1, 1, backup.classes.length, backup.classes[0].length).setValues(backup.classes);
    
    // Update validation and summary
    setupDataValidation(ss);
    updateSummary();
    
    showStatus(`Data restored from backup (${new Date(backup.timestamp).toLocaleString()})`, 'Restore Complete');
  } catch (error) {
    showStatus('Restore failed: ' + error.message, 'Restore Error');
    Logger.log('Error in restoreData: ' + error);
  }
} 