// Constants
const TEACHERS_SHEET_NAME = 'Teachers';
const CLASSES_SHEET_NAME = 'Classes';
const SUMMARY_SHEET_NAME = 'Summary';

const FIRST_PERIOD = 3;    // Column C (1st period)
const LAST_PERIOD = 13;    // Column M (9th period) - Changed from 11 to 13
const BREAK_COL = 6;       // Column F (Break)
const LUNCH_COL = 10;      // Column J (Lunch)

// Create menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Timetable System')
    .addItem('Setup System', 'setupEntireSystem')
    .addItem('Refresh Summary', 'updateSummary')
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
    protectSheets();
    
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

function setupDataValidation(ss) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Get the last row for both sheets
  const teachersLastRow = teachersSheet.getLastRow();
  const classesLastRow = classesSheet.getLastRow();
  
  // Create validation rules
  const classRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getUniqueClasses(), true)
    .setAllowInvalid(false)
    .build();
    
  const teacherRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getUniqueTeachers(ss), true)
    .setAllowInvalid(false)
    .build();
  
  // Clear existing validation first
  teachersSheet.getRange(3, FIRST_PERIOD, teachersLastRow - 2, LAST_PERIOD - FIRST_PERIOD + 1).clearDataValidations();
  classesSheet.getRange(2, FIRST_PERIOD, classesLastRow - 1, LAST_PERIOD - FIRST_PERIOD + 1).clearDataValidations();
  
  // Apply validation column by column to ensure no columns are missed
  for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
    // Skip break and lunch columns
    if (col === BREAK_COL || col === LUNCH_COL) continue;
    
    // Apply to Teachers sheet
    const teachersRange = teachersSheet.getRange(3, col, teachersLastRow - 2);
    teachersRange.setDataValidation(classRule);
    
    // Apply to Classes sheet
    const classesRange = classesSheet.getRange(2, col, classesLastRow - 1);
    classesRange.setDataValidation(teacherRule);
    
    // Log for debugging
    Logger.log(`Applied validation to column ${col}`);
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
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new edit trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// Helper function to get unique classes
function getUniqueClasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const lastRow = classesSheet.getLastRow();
  
  // Get all class names from column 2 (skip header)
  const classData = classesSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const uniqueClasses = new Set(classData.map(row => row[0]).filter(Boolean));
  return Array.from(uniqueClasses);
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
    const value = e.value || '';
    const oldValue = e.oldValue || '';
    
    // Check if this might be a paste operation (multiple cells or unexpected value format)
    if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
      range.clearContent();
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Pasting multiple cells is not allowed. Please enter values individually.',
        'Operation Not Allowed',
        5
      );
      return;
    }
    
    // For Teachers sheet
    if (sheet.getName() === TEACHERS_SHEET_NAME && row >= 3 && col >= FIRST_PERIOD) {
      // Validate the value format for class names
      if (value && !getUniqueClasses().includes(value)) {
        range.clearContent();
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Please use the dropdown to select a valid class.',
          'Invalid Entry',
          5
        );
        return;
      }
    }
    
    // For Classes sheet
    if (sheet.getName() === CLASSES_SHEET_NAME && row >= 2 && col >= FIRST_PERIOD) {
      // Validate the value format for teacher names (should contain " / ")
      if (value && (!value.includes(' / ') || !getUniqueTeachers(e.source).includes(value))) {
        range.clearContent();
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Please use the dropdown to select a valid teacher.',
          'Invalid Entry',
          5
        );
        return;
      }
    }
    
    // Handle name changes first
    if (col === 2 && ((sheet.getName() === TEACHERS_SHEET_NAME && row >= 3) || 
        (sheet.getName() === CLASSES_SHEET_NAME && row >= 2))) {
      handleNameChange(e);
      return;
    }
    
    let needsUpdate = false;
  
    // Modify the condition to include row 2 for Classes sheet
    if ((sheet.getName() === TEACHERS_SHEET_NAME && row > 2) || 
        (sheet.getName() === CLASSES_SHEET_NAME && row >= 2)) {
      if (col >= FIRST_PERIOD && col <= LAST_PERIOD && col !== BREAK_COL && col !== LUNCH_COL) {
        const teachersSheet = e.source.getSheetByName(TEACHERS_SHEET_NAME);
        const classesSheet = e.source.getSheetByName(CLASSES_SHEET_NAME);
        
        if (sheet.getName() === TEACHERS_SHEET_NAME) {
          // Check if the class is already assigned to another teacher at this time
          const allTeacherRows = teachersSheet.getRange(3, 1, teachersSheet.getLastRow() - 2, col + 1).getValues();
          const currentTeacher = teachersSheet.getRange(row, 2).getValue().split(' / ')[0];
          
          if (value) {
            // Check for conflicts
            let hasConflict = false;
            let conflictMessage = '';
            
            // Check if class is already assigned
            for (let i = 0; i < allTeacherRows.length; i++) {
              if (i !== row - 3 && allTeacherRows[i][col - 1] === value) {
                hasConflict = true;
                conflictMessage = `This class is already assigned to ${allTeacherRows[i][1].split(' / ')[0]} at this time slot.`;
                break;
              }
            }
            
            // Check if teacher is already teaching
            for (let i = 0; i < allTeacherRows.length; i++) {
              const teacherName = allTeacherRows[i][1].split(' / ')[0];
              if (teacherName === currentTeacher && i !== row - 3 && allTeacherRows[i][col - 1]) {
                hasConflict = true;
                conflictMessage = `${teacherName} is already teaching ${allTeacherRows[i][col - 1]} at this time slot.`;
                break;
              }
            }
            
            if (hasConflict) {
              range.clearContent();
              SpreadsheetApp.getActiveSpreadsheet().toast(conflictMessage, 'Schedule Conflict', 5);
              return;
            }
          }
          
          updateClassesFromTeachers(e);
          needsUpdate = true;
          
        } else if (sheet.getName() === CLASSES_SHEET_NAME) {
          if (value) {
            // Check for conflicts
            const teacherName = value.split(' / ')[0];
            const className = classesSheet.getRange(row, 2).getValue();
            let hasConflict = false;
            let conflictMessage = '';
            
            // Check if this teacher is already assigned in the same time slot
            const allClassRows = classesSheet.getRange(2, col, classesSheet.getLastRow() - 1, 1).getValues();
            for (let i = 0; i < allClassRows.length; i++) {
              const currentRow = i + 2;
              if (currentRow !== row && allClassRows[i][0]) {
                const existingTeacher = allClassRows[i][0].split(' / ')[0];
                if (existingTeacher === teacherName) {
                  hasConflict = true;
                  const conflictClass = classesSheet.getRange(currentRow, 2).getValue();
                  conflictMessage = `${teacherName} is already teaching ${conflictClass} at this time slot.`;
                  break;
                }
              }
            }
            
            if (!hasConflict) {
              const teacherRows = teachersSheet.getRange(3, 2, teachersSheet.getLastRow() - 2).getValues();
              let teacherFound = false;
              
              for (let i = 0; i < teacherRows.length; i++) {
                if (teacherRows[i][0] === value) {
                  teacherFound = true;
                  const existingClass = teachersSheet.getRange(i + 3, col).getValue();
                  if (existingClass && existingClass !== className) {
                    hasConflict = true;
                    conflictMessage = `${teacherName} is already assigned to ${existingClass} at this time slot.`;
                  }
                  break;
                }
              }
              
              if (!teacherFound) {
                hasConflict = true;
                conflictMessage = `Teacher ${teacherName} not found in the Teachers sheet.`;
              }
            }
            
            if (hasConflict) {
              range.clearContent();
              SpreadsheetApp.getActiveSpreadsheet().toast(conflictMessage, 'Schedule Conflict', 5);
              return;
            }
          }
          
          updateTeachersFromClasses(e);
          needsUpdate = true;
        }
        
        if (needsUpdate) {
          SpreadsheetApp.getActiveSpreadsheet().toast('Updating summary...', 'Status', 3);
          updateSummary();
          SpreadsheetApp.getActiveSpreadsheet().toast('Summary updated', 'Status', 3);
        }
      }
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'An error occurred: ' + error.message,
      'Error',
      5
    );
  }
}

// Add synchronization functions
function updateClassesFromTeachers(e) {
  const ss = e.source;
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value || '';
  
  if (value) {
    const teacherInfo = teachersSheet.getRange(row, 2).getValue();
    const className = value;
    
    // Find the class row and update it
    const classData = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1).getValues();
    let classFound = false;
    
    for (let i = 0; i < classData.length; i++) {
      if (classData[i][0].trim() === className.trim()) {
        classesSheet.getRange(i + 2, col).setValue(teacherInfo);
        classFound = true;
        break;
      }
    }
    
    if (!classFound) {
      Logger.log(`Class not found: ${className}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Unable to sync: Class "${className}" not found in Classes sheet`,
        'Sync Error',
        5
      );
    }
  } else {
    // Clear the corresponding cell in Classes sheet
    const className = e.oldValue;
    if (className) {
      const classData = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < classData.length; i++) {
        if (classData[i][0].trim() === className.trim()) {
          classesSheet.getRange(i + 2, col).clearContent();
          break;
        }
      }
    }
  }
}

function updateTeachersFromClasses(e) {
  const ss = e.source;
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value || '';
  
  // Get the class name for this row
  const className = classesSheet.getRange(row, 2).getValue().trim();
  
  if (value) {
    // Find all instances of this teacher in the Teachers sheet
    const teacherRows = teachersSheet.getRange(3, 2, teachersSheet.getLastRow() - 2).getValues();
    let teacherFound = false;
    let teacherUpdated = false;
    
    for (let i = 0; i < teacherRows.length; i++) {
      if (teacherRows[i][0].trim() === value.trim()) {
        teacherFound = true;
        // Check if this teacher slot is available
        const existingClass = teachersSheet.getRange(i + 3, col).getValue();
        if (!existingClass) {
          teachersSheet.getRange(i + 3, col).setValue(className);
          teacherUpdated = true;
          break;
        }
      }
    }
    
    if (!teacherFound) {
      Logger.log(`Teacher not found: ${value}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Unable to sync: Teacher "${value}" not found in Teachers sheet`,
        'Sync Error',
        5
      );
    } else if (!teacherUpdated) {
      Logger.log(`Teacher found but no available slot: ${value}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Unable to sync: No available slot for teacher "${value.split(' / ')[0]}" in Teachers sheet`,
        'Sync Error',
        5
      );
    }
  } else {
    // Clear the corresponding cell in Teachers sheet
    const oldValue = e.oldValue;
    if (oldValue) {
      const teacherRows = teachersSheet.getRange(3, 2, teachersSheet.getLastRow() - 2).getValues();
      for (let i = 0; i < teacherRows.length; i++) {
        if (teacherRows[i][0].trim() === oldValue.trim()) {
          const existingClass = teachersSheet.getRange(i + 3, col).getValue();
          if (existingClass && existingClass.trim() === className.trim()) {
            teachersSheet.getRange(i + 3, col).clearContent();
          }
        }
      }
    }
  }
}

// Add summary update function
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
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

// First, add a helper function to get unique teachers with subjects
function getUniqueTeachers(ss) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const lastRow = teachersSheet.getLastRow();
  if (lastRow < 3) return [];
  
  const teacherData = teachersSheet.getRange(3, 2, lastRow - 2, 1).getValues();
  const uniqueTeachers = new Set(teacherData.map(row => row[0]).filter(Boolean));
  return Array.from(uniqueTeachers);
}

// Add a function to reset validation if needed
function resetValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDataValidation(ss);
  SpreadsheetApp.getActiveSpreadsheet().toast('Data validation has been reset', 'Complete');
}

// Add this new function to handle name changes
function handleNameChange(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const newValue = e.value || '';
  const oldValue = e.oldValue || '';
  
  // Only proceed if we're editing the name columns (column 2) in Teachers or Classes sheets
  if (col === 2) {
    if (sheet.getName() === TEACHERS_SHEET_NAME && row >= 3) {
      updateTeacherNameChanges(e.source, oldValue, newValue, row);
    } else if (sheet.getName() === CLASSES_SHEET_NAME && row >= 2) {
      updateClassNameChanges(e.source, oldValue, newValue, row);
    }
  }
}

// Function to update teacher name changes across sheets
function updateTeacherNameChanges(ss, oldTeacher, newTeacher, teacherRow) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Get the subject from the new teacher name
  const teacherParts = newTeacher.split(' / ');
  if (teacherParts.length !== 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Teacher name must be in format "Name / Subject"',
      'Invalid Format',
      5
    );
    return;
  }
  
  // Update Classes sheet
  const classesLastRow = classesSheet.getLastRow();
  const classesRange = classesSheet.getRange(2, FIRST_PERIOD, classesLastRow - 1, LAST_PERIOD - FIRST_PERIOD + 1);
  const classesValues = classesRange.getValues();
  let hasChanges = false;
  
  // Get the old teacher name and subject
  const oldTeacherParts = oldTeacher.split(' / ');
  const oldTeacherName = oldTeacherParts[0];
  
  // First, clear the validation to prevent errors during update
  for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
    if (col !== BREAK_COL && col !== LUNCH_COL) {
      classesSheet.getRange(2, col, classesLastRow - 1).clearDataValidations();
    }
  }
  
  // Update each cell where the old teacher name appears
  for (let i = 0; i < classesValues.length; i++) {
    for (let j = 0; j < classesValues[i].length; j++) {
      const cellValue = classesValues[i][j];
      if (cellValue && cellValue.split(' / ')[0] === oldTeacherName) {
        classesValues[i][j] = newTeacher;
        hasChanges = true;
      }
    }
  }
  
  // Apply changes if any were made
  if (hasChanges) {
    classesRange.setValues(classesValues);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Updated teacher name from "${oldTeacherName}" to "${teacherParts[0]}"`,
      'Update Complete',
      5
    );
  }
  
  // Update data validation for Classes sheet
  const teacherRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getUniqueTeachers(ss), true)
    .setAllowInvalid(false)
    .build();
    
  // Apply updated validation to Classes sheet
  for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
    if (col !== BREAK_COL && col !== LUNCH_COL) {
      classesSheet.getRange(2, col, classesLastRow - 1).setDataValidation(teacherRule);
    }
  }
  
  // Update Teachers sheet validation as well
  const teachersLastRow = teachersSheet.getLastRow();
  const classRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getUniqueClasses(), true)
    .setAllowInvalid(false)
    .build();
    
  // Apply updated validation to Teachers sheet
  for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
    if (col !== BREAK_COL && col !== LUNCH_COL) {
      teachersSheet.getRange(3, col, teachersLastRow - 2).setDataValidation(classRule);
    }
  }
  
  // Update summary sheet
  updateSummary();
}

// Function to update class name changes across sheets
function updateClassNameChanges(ss, oldClass, newClass, classRow) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  
  // Update Teachers sheet
  const teachersLastRow = teachersSheet.getLastRow();
  const teachersRange = teachersSheet.getRange(3, FIRST_PERIOD, teachersLastRow - 2, LAST_PERIOD - FIRST_PERIOD + 1);
  const teachersValues = teachersRange.getValues();
  let hasChanges = false;
  
  // Update each cell where the old class name appears
  for (let i = 0; i < teachersValues.length; i++) {
    for (let j = 0; j < teachersValues[i].length; j++) {
      if (teachersValues[i][j] === oldClass) {
        teachersValues[i][j] = newClass;
        hasChanges = true;
      }
    }
  }
  
  // Apply changes if any were made
  if (hasChanges) {
    teachersRange.setValues(teachersValues);
  }
  
  // Update Summary sheet class names
  const summaryClassRange = summarySheet.getRange(3, 6, 21, 1); // Column F, starting from row 3
  const summaryClassValues = summaryClassRange.getValues();
  let summaryChanged = false;
  
  // Update class name in summary
  for (let i = 0; i < summaryClassValues.length; i++) {
    if (summaryClassValues[i][0] === oldClass) {
      summaryClassValues[i][0] = newClass;
      summaryChanged = true;
    }
  }
  
  // Apply changes to summary if any were made
  if (summaryChanged) {
    summaryClassRange.setValues(summaryClassValues);
  }
  
  // Update data validation for Teachers sheet
  const classRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getUniqueClasses(), true)
    .setAllowInvalid(false)
    .build();
    
  // Apply updated validation to Teachers sheet
  for (let col = FIRST_PERIOD; col <= LAST_PERIOD; col++) {
    if (col !== BREAK_COL && col !== LUNCH_COL) {
      teachersSheet.getRange(3, col, teachersLastRow - 2).setDataValidation(classRule);
    }
  }
  
  if (hasChanges || summaryChanged) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Updated class name from "${oldClass}" to "${newClass}"`,
      'Update Complete',
      5
    );
  }
  
  // Update summary sheet data
  updateSummary();
}

// Add this function to your existing code
function protectSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    ss.getSheetByName(TEACHERS_SHEET_NAME),
    ss.getSheetByName(CLASSES_SHEET_NAME),
    ss.getSheetByName(SUMMARY_SHEET_NAME)
  ];
  
  sheets.forEach(sheet => {
    if (!sheet) return;
    
    // Remove any existing protections
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(protection => protection.remove());
    
    // Protect the sheet
    const protection = sheet.protect();
    
    // Get the current user as editor
    const me = Session.getEffectiveUser();
    
    // Allow the current user to edit
    protection.addEditor(me);
    
    // Remove all other editors to prevent them from copying/pasting
    protection.removeEditors(protection.getEditors());
    
    // But allow them to edit certain ranges
    if (sheet.getName() === TEACHERS_SHEET_NAME) {
      protection.setUnprotectedRanges([
        sheet.getRange(3, 2, sheet.getLastRow() - 2, LAST_PERIOD - 1) // Teacher data area
      ]);
    } else if (sheet.getName() === CLASSES_SHEET_NAME) {
      protection.setUnprotectedRanges([
        sheet.getRange(2, 2, sheet.getLastRow() - 1, LAST_PERIOD - 1) // Class data area
      ]);
    }
    // Summary sheet remains fully protected
  });
} 