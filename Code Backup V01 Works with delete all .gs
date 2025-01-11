// Constants for sheet names
const TEACHERS_SHEET_NAME = 'Teachers';
const CLASSES_SHEET_NAME = 'Classes';
const SUMMARY_SHEET_NAME = 'Summary';
const CONFIG_TEACHERS_NAME = 'Config_Teachers';
const CONFIG_CLASSES_NAME = 'Config_Classes';
const CONFIG_SUBJECTS_NAME = 'Config_Subjects';

// Column constants remain the same
const TEACHERS_FIRST_PERIOD = 4;    // Column D (1st period)
const TEACHERS_LAST_PERIOD = 14;    // Column N (9th period)
const TEACHERS_BREAK_COL = 7;       // Column G (Break)
const TEACHERS_LUNCH_COL = 11;      // Column K (Lunch)

const CLASSES_FIRST_PERIOD = 3;     // Column C (1st period)
const CLASSES_LAST_PERIOD = 13;     // Column M (9th period)
const CLASSES_BREAK_COL = 6;        // Column F (Break)
const CLASSES_LUNCH_COL = 10;       // Column J (Lunch)

// Create menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Timetable System')
    .addItem('Setup Structure', 'setupStructure')
    .addSubMenu(ui.createMenu('Manage Database')
      .addItem('Show Config Sheets', 'showConfigSheets')
      .addItem('Hide Config Sheets', 'hideConfigSheets')
      .addItem('Deploy Data from Config', 'deployFromConfig')
      .addItem('Update Dropdowns', 'refreshDropdowns')
      .addSeparator()
      .addItem('Add New Teacher', 'addNewTeacher')
      .addItem('Add New Class', 'addNewClass')
      .addItem('Add New Subject', 'addNewSubject'))
    .addItem('Refresh Summary', 'updateSummary')
    .addItem('Clear All Data', 'clearAllData')
    .addToUi();
}

function setupStructure() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('No active spreadsheet found. Please open a spreadsheet before running this script.');
    }
    
    // Create sheets and set up basic structure
    setupSheets(ss);
    setupEmptyConfigSheets(ss);
    setupEmptyMainSheets(ss);
    setupEmptySummarySheet(ss);
    
    // Show config sheets for editing
    showConfigSheets();
    
    SpreadsheetApp.getUi().alert(
      'Structure Setup Complete!\n\n' +
      '1. Config sheets are now visible\n' +
      '2. Edit the config sheets with your data\n' +
      '3. Use "Deploy Data from Config" when ready to populate the sheets'
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

function setupEmptyConfigSheets(ss) {
  // Setup Teachers Config
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  teachersConfig.clear();
  teachersConfig.getRange('A1:C1').setValues([['Teacher ID', 'Teacher Name', 'Subject']]);
  teachersConfig.getRange('A1:C1')
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  teachersConfig.setColumnWidth(1, 100);
  teachersConfig.setColumnWidth(2, 200);
  teachersConfig.setColumnWidth(3, 150);
  
  // Add teachers data with their subjects
  const teachersData = [
    ['T001', 'Raju Bumb', 'English'],
    ['T002', 'Prabhat Karan', 'Maths'],
    ['T003', 'Shobha Hans', 'Science'],
    ['T004', 'Krishna Naidu', 'Hindi'],
    ['T005', 'Faraz Mangal', 'SST'],
    ['T006', 'Rimi Loke', 'Sanskrit'],
    ['T007', 'Amir Kar', 'English'],
    ['T008', 'Suraj Narayanan', 'Maths'],
    ['T009', 'Alaknanda Chaudry', 'Science'],
    ['T010', 'Preet Mittal', 'Hindi'],
    ['T011', 'John Lalla', 'SST'],
    ['T012', 'Ujwal Mohan', 'Sanskrit'],
    ['T013', 'Aadish Mathur', 'English'],
    ['T014', 'Iqbal Beharry', 'Maths'],
    ['T015', 'Manjari Shenoy', 'Science'],
    ['T016', 'Aayushi Suri', 'Hindi'],
    ['T017', 'Parvez Mathur', 'SST'],
    ['T018', 'Qabool Malhotra', 'Sanskrit'],
    ['T019', 'Nagma Andra', 'English'],
    ['T020', 'Krishna Arora', 'Maths'],
    ['T021', 'Nitin Banu', 'Science'],
    ['T022', 'Ananda Debnath', 'Hindi'],
    ['T023', 'Balaram Bhandari', 'SST'],
    ['T024', 'Ajay Chaudhri', 'Sanskrit'],
    ['T025', 'Niranjan Varma', 'English'],
    ['T026', 'Nur Patel', 'Maths']
  ];
  
  // Setup Classes Config
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  classesConfig.clear();
  classesConfig.getRange('A1:C1').setValues([['Class ID', 'Class Name', 'Section']]);
  classesConfig.getRange('A1:C1')
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  classesConfig.setColumnWidth(1, 100);
  classesConfig.setColumnWidth(2, 150);
  classesConfig.setColumnWidth(3, 100);
  
  // Add classes data
  const classesData = [
    ['C001', 'Nursery', ''],
    ['C002', 'LKG', 'A'],
    ['C003', 'LKG', 'B'],
    ['C004', 'UKG', 'A'],
    ['C005', 'UKG', 'B'],
    ['C006', 'Grade 1', 'A'],
    ['C007', 'Grade 1', 'B'],
    ['C008', 'Grade 2', 'A'],
    ['C009', 'Grade 2', 'B'],
    ['C010', 'Grade 3', 'A'],
    ['C011', 'Grade 3', 'B'],
    ['C012', 'Grade 4', 'A'],
    ['C013', 'Grade 4', 'B'],
    ['C014', 'Grade 5', ''],
    ['C015', 'Grade 6', ''],
    ['C016', 'Grade 7', ''],
    ['C017', 'Grade 8', ''],
    ['C018', 'Grade 9', ''],
    ['C019', 'Grade 10', ''],
    ['C020', 'Grade 11', ''],
    ['C021', 'Grade 12', '']
  ];
  
  // Setup Subjects Config
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  subjectsConfig.clear();
  subjectsConfig.getRange('A1:B1').setValues([['Subject ID', 'Subject Name']]);
  subjectsConfig.getRange('A1:B1')
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  subjectsConfig.setColumnWidth(1, 100);
  subjectsConfig.setColumnWidth(2, 200);
  
  // Add subjects data
  const subjectsData = [
    ['S001', 'English'],
    ['S002', 'Maths'],
    ['S003', 'Science'],
    ['S004', 'Hindi'],
    ['S005', 'SST'],
    ['S006', 'Sanskrit']
  ];
  
  // Write data to sheets
  teachersConfig.getRange(2, 1, teachersData.length, 3).setValues(teachersData);
  classesConfig.getRange(2, 1, classesData.length, 3).setValues(classesData);
  subjectsConfig.getRange(2, 1, subjectsData.length, 2).setValues(subjectsData);
  
  // Add alternating colors and borders to all config sheets
  const configSheets = [
    { sheet: teachersConfig, rows: teachersData.length, cols: 3 },
    { sheet: classesConfig, rows: classesData.length, cols: 3 },
    { sheet: subjectsConfig, rows: subjectsData.length, cols: 2 }
  ];
  
  configSheets.forEach(({sheet, rows, cols}) => {
    // Add alternating row colors
    for (let i = 0; i < rows; i++) {
      const rowNumber = i + 2;
      const color = i % 2 === 0 ? 'white' : '#f8f9fa';
      sheet.getRange(rowNumber, 1, 1, cols).setBackground(color);
    }
    
    // Add thick colored border
    sheet.getRange(1, 1, rows + 1, cols)
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
                
    // Center align ID columns
    sheet.getRange(2, 1, rows, 1).setHorizontalAlignment('center');
  });
  
  // Add subject validation to Teachers Config
  const subjectNames = subjectsData.map(row => row[1]);
  const subjectValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(subjectNames)
    .setAllowInvalid(false)
    .build();
  teachersConfig.getRange(2, 3, teachersData.length, 1).setDataValidation(subjectValidation);
  
  // Add ID format validation using simpler rules
  const teacherIdRule = SpreadsheetApp.newDataValidation()
    .requireTextContains('T')
    .setHelpText('Teacher ID must start with T followed by numbers (e.g., T001)')
    .build();
  
  const classIdRule = SpreadsheetApp.newDataValidation()
    .requireTextContains('C')
    .setHelpText('Class ID must start with C followed by numbers (e.g., C001)')
    .build();
  
  const subjectIdRule = SpreadsheetApp.newDataValidation()
    .requireTextContains('S')
    .setHelpText('Subject ID must start with S followed by numbers (e.g., S001)')
    .build();
  
  // Apply validation rules to the ID columns
  teachersConfig.getRange(2, 1, teachersData.length, 1).setDataValidation(teacherIdRule);
  classesConfig.getRange(2, 1, classesData.length, 1).setDataValidation(classIdRule);
  subjectsConfig.getRange(2, 1, subjectsData.length, 1).setDataValidation(subjectIdRule);
  
  // Protect headers in config sheets
  protectHeaders(teachersConfig, 1);
  protectHeaders(classesConfig, 1);
  protectHeaders(subjectsConfig, 1);
}

function deployFromConfig() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Deploy Data',
    'This will populate the Teachers and Classes sheets with data from your config sheets. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get config sheets
    const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
    const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
    
    // Verify config data exists
    const teachersLastRow = teachersConfig.getLastRow();
    const classesLastRow = classesConfig.getLastRow();
    
    if (teachersLastRow < 2 || classesLastRow < 2) {
      throw new Error('Config sheets are empty. Please add data to the config sheets first.');
    }
    
    // Get the data from config sheets
    const teacherData = teachersConfig.getRange(2, 2, teachersLastRow - 1, 2).getValues();
    const classData = classesConfig.getRange(2, 2, classesLastRow - 1, 2).getValues();
    
    // Clear existing sheets
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    
    // Clear existing data and validations
    if (teachersSheet.getLastRow() > 1) {
      teachersSheet.getRange(2, 1, teachersSheet.getLastRow() - 1, teachersSheet.getLastColumn()).clear();
    }
    if (classesSheet.getLastRow() > 1) {
      classesSheet.getRange(2, 1, classesSheet.getLastRow() - 1, classesSheet.getLastColumn()).clear();
    }
    
    // Set up headers and structure
    setupHeaders(ss);
    setupClassesSheet(ss);
    setupBreakLunchColumns(ss);
    
    // Set up dropdowns for all period columns at once
    setupDropdowns(ss);
    
    // Update summary
    updateSummary();
    
    ui.alert('Success', 'Data has been deployed from config sheets.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to deploy data: ' + error.message, ui.ButtonSet.OK);
  }
}

function showConfigSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheets = [
      { sheet: ss.getSheetByName(CONFIG_TEACHERS_NAME), name: 'Teachers Config' },
      { sheet: ss.getSheetByName(CONFIG_CLASSES_NAME), name: 'Classes Config' },
      { sheet: ss.getSheetByName(CONFIG_SUBJECTS_NAME), name: 'Subjects Config' }
    ];
    
    const missingSheets = configSheets
      .filter(({sheet}) => !sheet)
      .map(({name}) => name);
    
    if (missingSheets.length > 0) {
      throw new Error('Missing configuration sheets: ' + missingSheets.join(', ') + 
                     '. Please run Setup System first.');
    }
    
    configSheets.forEach(({sheet}) => sheet.showSheet());
    SpreadsheetApp.getUi().alert('Configuration sheets are now visible. You can edit them and then use "Update Dropdowns" to apply changes.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error showing config sheets: ' + error.message);
  }
}

function hideConfigSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheets = [
      { sheet: ss.getSheetByName(CONFIG_TEACHERS_NAME), name: 'Teachers Config' },
      { sheet: ss.getSheetByName(CONFIG_CLASSES_NAME), name: 'Classes Config' },
      { sheet: ss.getSheetByName(CONFIG_SUBJECTS_NAME), name: 'Subjects Config' }
    ];
    
    configSheets.forEach(({sheet, name}) => {
      if (sheet) sheet.hideSheet();
    });
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error hiding config sheets: ' + error.message);
  }
}

function refreshDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDropdowns(ss);
}

function setupSheets(ss) {
  try {
    // List of all required sheets
    const requiredSheets = [
      { name: TEACHERS_SHEET_NAME, isConfig: false },
      { name: CLASSES_SHEET_NAME, isConfig: false },
      { name: SUMMARY_SHEET_NAME, isConfig: false },
      { name: CONFIG_TEACHERS_NAME, isConfig: true },
      { name: CONFIG_CLASSES_NAME, isConfig: true },
      { name: CONFIG_SUBJECTS_NAME, isConfig: true }
    ];
    
    // Create or get each required sheet
    requiredSheets.forEach(sheetInfo => {
      let sheet = ss.getSheetByName(sheetInfo.name);
      if (!sheet) {
        sheet = ss.insertSheet(sheetInfo.name);
      }
      if (sheetInfo.isConfig) {
        sheet.hideSheet();
      }
    });
    
    // Delete any other sheets
    ss.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (!requiredSheets.some(s => s.name === sheetName)) {
        ss.deleteSheet(sheet);
      }
    });
    
    // Reorder visible sheets
    const visibleSheets = [TEACHERS_SHEET_NAME, CLASSES_SHEET_NAME, SUMMARY_SHEET_NAME];
    visibleSheets.forEach((sheetName, index) => {
      const sheet = ss.getSheetByName(sheetName);
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    });
    
  } catch (error) {
    throw new Error('Error setting up sheets: ' + error.message);
  }
}

function setupConfigSheets(ss) {
  // Setup Teachers Config
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  teachersConfig.clear();
  teachersConfig.getRange('A1:C1').setValues([['Teacher ID', 'Teacher Name', 'Subject']]);
  teachersConfig.getRange('A1:C1')
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  teachersConfig.setColumnWidth(1, 100);
  teachersConfig.setColumnWidth(2, 200);
  teachersConfig.setColumnWidth(3, 150);
  
  // Setup Classes Config
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  classesConfig.clear();
  classesConfig.getRange('A1:C1').setValues([['Class ID', 'Class Name', 'Section']]);
  classesConfig.getRange('A1:C1')
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  classesConfig.setColumnWidth(1, 100);
  classesConfig.setColumnWidth(2, 150);
  classesConfig.setColumnWidth(3, 100);
  
  // Setup Subjects Config
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  subjectsConfig.clear();
  subjectsConfig.getRange('A1:B1').setValues([['Subject ID', 'Subject Name']]);
  subjectsConfig.getRange('A1:B1')
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  subjectsConfig.setColumnWidth(1, 100);
  subjectsConfig.setColumnWidth(2, 200);
  
  // Add teachers data with their subjects
  const teachersData = [
    ['T001', 'Raju Bumb', 'English'],
    ['T002', 'Prabhat Karan', 'Maths'],
    ['T003', 'Shobha Hans', 'Science'],
    ['T004', 'Krishna Naidu', 'Hindi'],
    ['T005', 'Faraz Mangal', 'SST'],
    ['T006', 'Rimi Loke', 'Sanskrit'],
    ['T007', 'Amir Kar', 'English'],
    ['T008', 'Suraj Narayanan', 'Maths'],
    ['T009', 'Alaknanda Chaudry', 'Science'],
    ['T010', 'Preet Mittal', 'Hindi'],
    ['T011', 'John Lalla', 'SST'],
    ['T012', 'Ujwal Mohan', 'Sanskrit'],
    ['T013', 'Aadish Mathur', 'English'],
    ['T014', 'Iqbal Beharry', 'Maths'],
    ['T015', 'Manjari Shenoy', 'Science'],
    ['T016', 'Aayushi Suri', 'Hindi'],
    ['T017', 'Parvez Mathur', 'SST'],
    ['T018', 'Qabool Malhotra', 'Sanskrit'],
    ['T019', 'Nagma Andra', 'English'],
    ['T020', 'Krishna Arora', 'Maths'],
    ['T021', 'Nitin Banu', 'Science'],
    ['T022', 'Ananda Debnath', 'Hindi'],
    ['T023', 'Balaram Bhandari', 'SST'],
    ['T024', 'Ajay Chaudhri', 'Sanskrit'],
    ['T025', 'Niranjan Varma', 'English'],
    ['T026', 'Nur Patel', 'Maths']
  ];
  
  // Add classes data
  const classesData = [
    ['C001', 'Nursery', ''],
    ['C002', 'LKG', 'A'],
    ['C003', 'LKG', 'B'],
    ['C004', 'UKG', 'A'],
    ['C005', 'UKG', 'B'],
    ['C006', 'Grade 1', 'A'],
    ['C007', 'Grade 1', 'B'],
    ['C008', 'Grade 2', 'A'],
    ['C009', 'Grade 2', 'B'],
    ['C010', 'Grade 3', 'A'],
    ['C011', 'Grade 3', 'B'],
    ['C012', 'Grade 4', 'A'],
    ['C013', 'Grade 4', 'B'],
    ['C014', 'Grade 5', ''],
    ['C015', 'Grade 6', ''],
    ['C016', 'Grade 7', ''],
    ['C017', 'Grade 8', ''],
    ['C018', 'Grade 9', ''],
    ['C019', 'Grade 10', ''],
    ['C020', 'Grade 11', ''],
    ['C021', 'Grade 12', '']
  ];
  
  // Add subjects data
  const subjectsData = [
    ['S001', 'English'],
    ['S002', 'Maths'],
    ['S003', 'Science'],
    ['S004', 'Hindi'],
    ['S005', 'SST'],
    ['S006', 'Sanskrit']
  ];
  
  // Write data
  teachersConfig.getRange(2, 1, teachersData.length, 3).setValues(teachersData);
  classesConfig.getRange(2, 1, classesData.length, 3).setValues(classesData);
  subjectsConfig.getRange(2, 1, subjectsData.length, 2).setValues(subjectsData);
  
  // Add alternating colors to all config sheets
  const configSheets = [
    { sheet: teachersConfig, rows: teachersData.length, cols: 3 },
    { sheet: classesConfig, rows: classesData.length, cols: 3 },
    { sheet: subjectsConfig, rows: subjectsData.length, cols: 2 }
  ];
  
  configSheets.forEach(({sheet, rows, cols}) => {
    // Add alternating row colors
    for (let i = 0; i < rows; i++) {
      const rowNumber = i + 2;
      const color = i % 2 === 0 ? 'white' : '#f8f9fa';
      sheet.getRange(rowNumber, 1, 1, cols).setBackground(color);
    }
    
    // Add thick colored border
    sheet.getRange(1, 1, rows + 1, cols)
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
                
    // Center align ID columns
    sheet.getRange(2, 1, rows, 1).setHorizontalAlignment('center');
  });
  
  // Add subject validation to Teachers Config
  const subjectNames = subjectsData.map(row => row[1]);
  const subjectValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(subjectNames)
    .setAllowInvalid(false)
    .build();
  teachersConfig.getRange(2, 3, teachersData.length, 1).setDataValidation(subjectValidation);
  
  // Add ID format validation using simpler rules
  const teacherIdRule = SpreadsheetApp.newDataValidation()
    .requireTextContains('T')
    .setHelpText('Teacher ID must start with T followed by numbers (e.g., T001)')
    .build();
  
  const classIdRule = SpreadsheetApp.newDataValidation()
    .requireTextContains('C')
    .setHelpText('Class ID must start with C followed by numbers (e.g., C001)')
    .build();
  
  const subjectIdRule = SpreadsheetApp.newDataValidation()
    .requireTextContains('S')
    .setHelpText('Subject ID must start with S followed by numbers (e.g., S001)')
    .build();
  
  // Apply validation rules to the ID columns
  teachersConfig.getRange(2, 1, teachersData.length, 1).setDataValidation(teacherIdRule);
  classesConfig.getRange(2, 1, classesData.length, 1).setDataValidation(classIdRule);
  subjectsConfig.getRange(2, 1, subjectsData.length, 1).setDataValidation(subjectIdRule);
  
  // Protect headers in config sheets
  protectHeaders(teachersConfig, 1);
  protectHeaders(classesConfig, 1);
  protectHeaders(subjectsConfig, 1);
}

function setupHeaders(ss) {
  // Teachers sheet - Grid layout with timing
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  teachersSheet.clear();
  
  // Get teacher data from config to determine row count
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  const teacherData = teachersConfig.getRange(2, 2, teachersConfig.getLastRow() - 1, 2).getValues();
  
  const initialRows = teacherData.length; // Dynamic row count based on actual teachers
  
  // Set up header rows
  const mainHeader = [
    'SI',
    'Teacher Name',
    'Subject',
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
  teachersSheet.getRange(1, 1, 1, mainHeader.length).setValues([mainHeader]);
  
  // Format headers
  teachersSheet.getRange(1, 1, 1, mainHeader.length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // Set column widths
  teachersSheet.setColumnWidth(1, 30);   // SI column
  teachersSheet.setColumnWidth(2, 150);  // Teacher Name column
  teachersSheet.setColumnWidth(3, 100);  // Subject column
  for (let i = 4; i <= mainHeader.length; i++) {
    teachersSheet.setColumnWidth(i, 100);
  }
  
  // Set row height for header
  teachersSheet.setRowHeight(1, 60);
  
  // Create teacher rows with subjects from config
  const teacherRows = teacherData.map((row, index) => [
    index + 1,      // SI
    row[0],         // Teacher Name
    row[1]          // Subject (from config)
  ]);
  
  // Write teacher data
  teachersSheet.getRange(2, 1, initialRows, 3).setValues(teacherRows);
  
  // Clear any existing data validations from Teacher Name and Subject columns
  teachersSheet.getRange(2, 2, initialRows, 2).clearDataValidations();
  
  // Format the table area
  const tableRange = teachersSheet.getRange(1, 1, initialRows + 1, mainHeader.length);
  tableRange.setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // Add thick border around the entire table
  tableRange.setBorder(
      true, true, true, true, null, null,
      '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  
  // Center align SI column and Subject column
  teachersSheet.getRange(1, 1, initialRows + 1, 1).setHorizontalAlignment('center');
  teachersSheet.getRange(1, 3, initialRows + 1, 1).setHorizontalAlignment('center');
  
  // Left align Teacher Name column
  teachersSheet.getRange(2, 2, initialRows, 1).setHorizontalAlignment('left');
  
  // Add alternating row colors
  for (let i = 0; i < initialRows; i++) {
    const rowNumber = i + 2;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    teachersSheet.getRange(rowNumber, 1, 1, mainHeader.length).setBackground(color);
  }
  
  // Clear any content and formatting below the table
  const totalRows = teachersSheet.getMaxRows();
  if (totalRows > initialRows + 1) {
    teachersSheet.deleteRows(initialRows + 2, totalRows - (initialRows + 1));
  }
  
  // Protect header row
  protectHeaders(teachersSheet, 1);
  
  // Make Teacher Name and Subject columns read-only
  const protection = teachersSheet.getRange(2, 2, initialRows, 2).protect();
  protection.setDescription('Teacher and Subject columns - Protected');
  protection.setWarningOnly(true);
}

function setupClassesSheet(ss) {
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  classesSheet.clear();
  
  // Get class data from config to determine row count
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  const classData = classesConfig.getRange(2, 2, classesConfig.getLastRow() - 1, 2).getValues();
  
  const initialRows = classData.length; // Dynamic row count based on actual classes
  
  // Set up header rows
  const mainHeader = [
    'SI',
    'Classes',
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
  
  // Create class rows with proper formatting
  const classRows = classData.map((row, index) => {
    const className = row[0];
    const section = row[1];
    let formattedName;
    
    if (!section) {
      formattedName = className; // e.g., "Nursery"
    } else if (className.startsWith('Grade')) {
      formattedName = `${className.replace(' ', ' - ')}${section}`; // e.g., "Grade - 1A"
    } else {
      formattedName = `${className} - ${section}`; // e.g., "LKG - A"
    }
    
    return [index + 1, formattedName];
  });
  
  // Write class data
  classesSheet.getRange(2, 1, initialRows, 2).setValues(classRows);
  
  // Format the table area
  const tableRange = classesSheet.getRange(1, 1, initialRows + 1, mainHeader.length);
  tableRange.setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // Add thick border around the entire table
  tableRange.setBorder(
    true, true, true, true, null, null,
      '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  
  // Center align SI column
  classesSheet.getRange(1, 1, initialRows + 1, 1).setHorizontalAlignment('center');
  
  // Add alternating row colors
  for (let i = 0; i < initialRows; i++) {
    const rowNumber = i + 2;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    classesSheet.getRange(rowNumber, 1, 1, mainHeader.length).setBackground(color);
  }
  
  // Clear any content and formatting below the table
  const totalRows = classesSheet.getMaxRows();
  if (totalRows > initialRows + 1) {
    classesSheet.deleteRows(initialRows + 2, totalRows - (initialRows + 1));
  }
  
  // Protect header row
  protectHeaders(classesSheet, 1);
  
  // Make Classes column read-only
  const protection = classesSheet.getRange(2, 2, initialRows, 1).protect();
  protection.setDescription('Classes column - Protected');
  protection.setWarningOnly(true);
}

function setupBreakLunchColumns(ss) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Get the actual number of rows from the sheets
  const teachersLastRow = teachersSheet.getLastRow();
  const classesLastRow = classesSheet.getLastRow();
  
  // Color Break columns
  teachersSheet.getRange(1, TEACHERS_BREAK_COL, teachersLastRow).setBackground('#ffcdd2');
  classesSheet.getRange(1, CLASSES_BREAK_COL, classesLastRow).setBackground('#ffcdd2');
  
  // Color Lunch columns
  teachersSheet.getRange(1, TEACHERS_LUNCH_COL, teachersLastRow).setBackground('#ffcdd2');
  classesSheet.getRange(1, CLASSES_LUNCH_COL, classesLastRow).setBackground('#ffcdd2');
  
  // Protect Break and Lunch columns
  protectBreakLunchColumns(teachersSheet, TEACHERS_BREAK_COL, TEACHERS_LUNCH_COL, teachersLastRow);
  protectBreakLunchColumns(classesSheet, CLASSES_BREAK_COL, CLASSES_LUNCH_COL, classesLastRow);
}

function protectBreakLunchColumns(sheet, breakCol, lunchCol, lastRow) {
  try {
    // Protect Break column
    const breakProtection = sheet.getRange(2, breakCol, lastRow - 1, 1).protect();
    breakProtection.setDescription('Break Column - Protected');
    breakProtection.setWarningOnly(true);
    
    // Protect Lunch column
    const lunchProtection = sheet.getRange(2, lunchCol, lastRow - 1, 1).protect();
    lunchProtection.setDescription('Lunch Column - Protected');
    lunchProtection.setWarningOnly(true);
  } catch (error) {
    console.log('Error protecting Break/Lunch columns: ' + error.message);
  }
}

function setupSummarySheet(ss) {
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  summarySheet.clear();
  
  // Get the number of teachers and classes from config sheets
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  
  const numTeachers = Math.max(teachersConfig.getLastRow() - 1, 1);
  const numClasses = Math.max(classesConfig.getLastRow() - 1, 1);
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  
  // Set up headers for Teachers table (left side)
  const teachersHeaders = [
    ['Teachers Summary'],
    ['SI', 'Teacher Name', 'Total\nPeriods']
  ];
  
  // Set up headers for Classwise Subject Allotment table (right side)
  const classHeaders = [
    ['Classwise Subject Allotment'],
    ['SI', 'Classes', ...subjects]
  ];
  
  const rightTableWidth = classHeaders[1].length;
  
  // Write headers
  summarySheet.getRange(1, 1, 1, 3).merge().setValue(teachersHeaders[0][0]);
  summarySheet.getRange(2, 1, 1, 3).setValues([teachersHeaders[1]]);
  
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
         .setHorizontalAlignment('center');
  });
  
  // Set column widths
  summarySheet.setColumnWidth(1, 30);   // SI
  summarySheet.setColumnWidth(2, 150);  // Teacher Name
  summarySheet.setColumnWidth(3, 80);   // Total Periods
  summarySheet.setColumnWidth(4, 30);   // Gap between tables
  summarySheet.setColumnWidth(5, 30);   // SI
  summarySheet.setColumnWidth(6, 150);  // Classes
  for (let i = 7; i < 7 + subjects.length; i++) {    // Subject columns
    summarySheet.setColumnWidth(i, 100);
  }
  
  // Set row heights
  summarySheet.setRowHeight(1, 30);
  summarySheet.setRowHeight(2, 40);
  
  // Add empty rows based on actual data
  const maxRows = Math.max(numTeachers, numClasses);
  const emptyTeacherRows = Array(maxRows).fill(['', '', '']);
  const emptyClassRows = Array(maxRows).fill(['', '', ...Array(subjects.length).fill('')]);
  
  // Add empty rows to both tables
  summarySheet.getRange(3, 1, maxRows, 3).setValues(emptyTeacherRows);
  summarySheet.getRange(3, 5, maxRows, rightTableWidth).setValues(emptyClassRows);
  
  // Format data areas
  const leftTable = summarySheet.getRange(1, 1, maxRows + 2, 3);
  const rightTable = summarySheet.getRange(1, 5, maxRows + 2, rightTableWidth);
  
  [leftTable, rightTable].forEach(range => {
    range.setBorder(true, true, true, true, true, true)
         .setVerticalAlignment('middle')
         .setFontSize(10);
  });
  
  // Set alignments for data areas
  summarySheet.getRange(3, 1, maxRows, 1).setHorizontalAlignment('center'); // Left SI
  summarySheet.getRange(3, 2, maxRows, 1).setHorizontalAlignment('left');   // Teacher names
  summarySheet.getRange(3, 3, maxRows, 1).setHorizontalAlignment('center'); // Total periods
  summarySheet.getRange(3, 5, maxRows, 1).setHorizontalAlignment('center'); // Right SI
  summarySheet.getRange(3, 6, maxRows, 1).setHorizontalAlignment('left');   // Class names
  summarySheet.getRange(3, 7, maxRows, subjects.length).setHorizontalAlignment('center'); // Subject columns
  
  // Add alternating row colors
  for (let i = 0; i < maxRows; i++) {
    const rowNumber = i + 3;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    summarySheet.getRange(rowNumber, 1, 1, 3).setBackground(color);
    summarySheet.getRange(rowNumber, 5, 1, rightTableWidth).setBackground(color);
  }
  
  // Add thick colored borders to both tables
  leftTable.setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  rightTable.setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Protect header rows
  protectHeaders(summarySheet, 2);
  
  // Clear any content and formatting below the tables
  const totalRows = summarySheet.getMaxRows();
  if (totalRows > maxRows + 2) {
    summarySheet.deleteRows(maxRows + 3, totalRows - (maxRows + 2));
  }
}

function setupDropdowns(ss) {
  try {
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
    const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
    
    // Get the actual number of rows from the sheets
    const teachersLastRow = teachersSheet.getLastRow();
    const classesLastRow = classesSheet.getLastRow();
    
    // First, clear ALL existing validations
    if (teachersLastRow > 1) {
      teachersSheet.getRange(2, 1, teachersLastRow - 1, teachersSheet.getLastColumn()).clearDataValidations();
    }
    if (classesLastRow > 1) {
      classesSheet.getRange(2, 1, classesLastRow - 1, classesSheet.getLastColumn()).clearDataValidations();
    }
    
    // Define period columns for Teachers sheet (D through N, excluding Break and Lunch)
    const teacherPeriodCols = [4, 5, 6, 8, 9, 10, 12, 13, 14]; // D, E, F, H, I, J, L, M, N
    
    // Define period columns for Classes sheet (C through M, excluding Break and Lunch)
    const classPeriodCols = [3, 4, 5, 7, 8, 9, 11, 12, 13]; // C, D, E, G, H, I, K, L, M
    
    // Get class data for Teachers sheet dropdowns
    const classData = classesConfig.getRange(2, 2, classesConfig.getLastRow() - 1, 2).getValues();
    const formattedClasses = classData.map(row => {
      const className = row[0];
      const section = row[1];
      if (!section) return className;
      if (className.startsWith('Grade')) return `${className.replace(' ', ' - ')}${section}`;
      return `${className} - ${section}`;
    });
    
    // Get teacher-subject combinations for Classes sheet dropdowns
    const teacherData = teachersConfig.getRange(2, 2, teachersConfig.getLastRow() - 1, 2).getValues();
    const teacherSubjectCombos = teacherData.map(row => `${row[0]} / ${row[1]}`);
    
    // Set up initial dropdowns for Teachers sheet
    teacherPeriodCols.forEach(col => {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(formattedClasses)
        .setAllowInvalid(false)
        .build();
      teachersSheet.getRange(2, col, teachersLastRow - 1, 1).setDataValidation(rule);
    });
    
    // Set up initial dropdowns for Classes sheet
    classPeriodCols.forEach(col => {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(teacherSubjectCombos)
        .setAllowInvalid(false)
        .build();
      classesSheet.getRange(2, col, classesLastRow - 1, 1).setDataValidation(rule);
    });
    
  } catch (error) {
    console.log('Error setting up dropdowns: ' + error.message);
  }
}

// Cache constants
const CACHE_KEYS = {
  TEACHERS_DATA: 'teachersData',
  CLASSES_DATA: 'classesData',
  SUBJECTS_DATA: 'subjectsData',
  PERIOD_VALUES: 'periodValues'
};

const CACHE_DURATION = 21600; // 6 hours in seconds

// Cache management functions
function getFromCache(key) {
  const cache = CacheService.getScriptCache();
  const data = cache.get(key);
  return data ? JSON.parse(data) : null;
}

function setInCache(key, data) {
  const cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(data), CACHE_DURATION);
}

function clearCache() {
  const cache = CacheService.getScriptCache();
  Object.values(CACHE_KEYS).forEach(key => cache.remove(key));
}

function updatePeriodDropdowns(ss, sheetName, periodCol) {
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  
  // Get cached data or fetch from sheets
  let formattedClasses = getFromCache(CACHE_KEYS.CLASSES_DATA);
  let teacherSubjectCombos = getFromCache(CACHE_KEYS.TEACHERS_DATA);
  
  if (!formattedClasses || !teacherSubjectCombos) {
    const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
    const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
    
    if (sheetName === TEACHERS_SHEET_NAME) {
      const classData = classesConfig.getRange(2, 2, classesConfig.getLastRow() - 1, 2).getValues();
      formattedClasses = classData.map(row => {
        const className = row[0];
        const section = row[1];
        if (!section) return className;
        if (className.startsWith('Grade')) return `${className.replace(' ', ' - ')}${section}`;
        return `${className} - ${section}`;
      });
      setInCache(CACHE_KEYS.CLASSES_DATA, formattedClasses);
    } else {
      const teacherData = teachersConfig.getRange(2, 2, teachersConfig.getLastRow() - 1, 2).getValues();
      teacherSubjectCombos = teacherData.map(row => `${row[0]} / ${row[1]}`);
      setInCache(CACHE_KEYS.TEACHERS_DATA, teacherSubjectCombos);
    }
  }
  
  // Get used values for this period
  const periodValues = sheet.getRange(2, periodCol, lastRow - 1, 1).getValues();
  const usedValues = new Set();
  
  periodValues.forEach(([value]) => {
    if (value) {
      if (sheetName === CLASSES_SHEET_NAME) {
        const [teacherName] = value.split(' / ');
        usedValues.add(teacherName);
      } else {
        usedValues.add(value);
      }
    }
  });
  
  // Update dropdowns for each row
  const rules = [];
  for (let row = 2; row <= lastRow; row++) {
    const currentValue = sheet.getRange(row, periodCol).getValue();
    let availableOptions;
    
    if (sheetName === TEACHERS_SHEET_NAME) {
      availableOptions = formattedClasses.filter(className => 
        !usedValues.has(className) || className === currentValue
      );
    } else {
      let currentTeacher = '';
      if (currentValue) {
        [currentTeacher] = currentValue.split(' / ');
      }
      availableOptions = teacherSubjectCombos.filter(combo => {
        const [teacherName] = combo.split(' / ');
        return !usedValues.has(teacherName) || teacherName === currentTeacher;
      });
    }
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(availableOptions)
      .setAllowInvalid(false)
      .build();
    rules.push({ range: sheet.getRange(row, periodCol), rule: rule });
  }
  
  // Apply all rules at once
  rules.forEach(({range, rule}) => range.setDataValidation(rule));
}

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    
    // Skip if editing headers
    if (row === 1) return;
    
    // Handle different sheets
    switch(sheetName) {
      case TEACHERS_SHEET_NAME:
        if (col >= TEACHERS_FIRST_PERIOD && col <= TEACHERS_LAST_PERIOD) {
          // Skip Break and Lunch columns
          if (col === TEACHERS_BREAK_COL || col === TEACHERS_LUNCH_COL) return;
          
          const teachersSheet = e.source.getSheetByName(TEACHERS_SHEET_NAME);
          const classesSheet = e.source.getSheetByName(CLASSES_SHEET_NAME);
          
          // Handle multiple cell selection
          const affectedCols = new Set();
          for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
              const currentRow = row + r;
              const currentCol = col + c;
              
              if (currentCol === TEACHERS_BREAK_COL || currentCol === TEACHERS_LUNCH_COL) continue;
              
              const teacherName = teachersSheet.getRange(currentRow, 2).getValue();
              const teacherSubject = teachersSheet.getRange(currentRow, 3).getValue();
              const selectedClass = teachersSheet.getRange(currentRow, currentCol).getValue();
              
              if (teacherName && teacherSubject) {
                // Update corresponding class cells
                const classesData = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1).getValues();
                classesData.forEach((classRow, index) => {
                  const className = classRow[0];
                  if (className) {
                    const classRowNum = index + 2;
                    const classCol = currentCol - 1;
                    const classCell = classesSheet.getRange(classRowNum, classCol);
                    const classCellValue = classCell.getValue();
                    
                    if (selectedClass) {
                      if (className === selectedClass || className.replace(' - ', ' ') === selectedClass) {
                        classCell.setValue(`${teacherName} / ${teacherSubject}`);
                      } else if (classCellValue && classCellValue.startsWith(`${teacherName} /`)) {
                        classCell.clearContent();
                      }
                    } else {
                      if (classCellValue && classCellValue.startsWith(`${teacherName} /`)) {
                        classCell.clearContent();
                      }
                    }
                  }
                });
                affectedCols.add(currentCol);
              }
            }
          }
          
          // Update dropdowns for all affected columns
          affectedCols.forEach(col => {
            updatePeriodDropdowns(e.source, TEACHERS_SHEET_NAME, col);
            updatePeriodDropdowns(e.source, CLASSES_SHEET_NAME, col - 1);
          });
          
          updateSummary();
        }
        break;
        
      case CLASSES_SHEET_NAME:
        if (col >= CLASSES_FIRST_PERIOD && col <= CLASSES_LAST_PERIOD) {
          // Skip Break and Lunch columns
          if (col === CLASSES_BREAK_COL || col === CLASSES_LUNCH_COL) return;
          
          const teachersSheet = e.source.getSheetByName(TEACHERS_SHEET_NAME);
          const classesSheet = e.source.getSheetByName(CLASSES_SHEET_NAME);
          
          // Handle multiple cell selection
          const affectedCols = new Set();
          for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
              const currentRow = row + r;
              const currentCol = col + c;
              
              if (currentCol === CLASSES_BREAK_COL || currentCol === CLASSES_LUNCH_COL) continue;
              
              const className = classesSheet.getRange(currentRow, 2).getValue();
              const selectedTeacher = classesSheet.getRange(currentRow, currentCol).getValue();
              
              if (className) {
                if (selectedTeacher) {
                  // If a teacher is selected in class's sheet
                  const [teacherName] = selectedTeacher.split(' / ');
                  
                  // Find the teacher row
                  const teachersData = teachersSheet.getRange(2, 2, teachersSheet.getLastRow() - 1, 1).getValues();
                  teachersData.forEach((teacherRow, index) => {
                    if (teacherRow[0] === teacherName) {
                      const teacherRowNum = index + 2;
                      const teacherCol = currentCol + 1;
                      teachersSheet.getRange(teacherRowNum, teacherCol).setValue(className);
                    } else if (teacherRow[0]) {
                      // Clear this period for other teachers who had this class
                      const teacherRowNum = index + 2;
                      const teacherCol = currentCol + 1;
                      const teacherCell = teachersSheet.getRange(teacherRowNum, teacherCol);
                      const teacherCellValue = teacherCell.getValue();
                      if (teacherCellValue === className || 
                          teacherCellValue.replace(' - ', ' ') === className) {
                        teacherCell.clearContent();
                      }
                    }
                  });
                  affectedCols.add(currentCol);
                } else {
                  // If cell is cleared in class's sheet
                  const teachersData = teachersSheet.getRange(2, 2, teachersSheet.getLastRow() - 1, 1).getValues();
                  teachersData.forEach((teacherRow, index) => {
                    if (teacherRow[0]) {
                      const teacherRowNum = index + 2;
                      const teacherCol = currentCol + 1;
                      const teacherCell = teachersSheet.getRange(teacherRowNum, teacherCol);
                      const teacherCellValue = teacherCell.getValue();
                      if (teacherCellValue === className || 
                          teacherCellValue.replace(' - ', ' ') === className) {
                        teacherCell.clearContent();
                      }
                    }
                  });
                  affectedCols.add(currentCol);
                }
              }
            }
          }
          
          // Update dropdowns for all affected columns
          affectedCols.forEach(col => {
            updatePeriodDropdowns(e.source, CLASSES_SHEET_NAME, col);
            updatePeriodDropdowns(e.source, TEACHERS_SHEET_NAME, col + 1);
          });
          
          updateSummary();
        }
        break;
    }
  } catch (error) {
    console.log('Error in onEdit: ' + error.message);
  }
}

function updateSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  
    if (!teachersSheet || !classesSheet || !summarySheet) {
      console.log('Required sheets not found. Please run Setup System first.');
    return;
  }
  
    updateTeachersSummary(teachersSheet, summarySheet);
    updateClasswiseSummary(classesSheet, summarySheet);
  } catch (error) {
    console.log('Error updating summary: ' + error.message);
  }
}

function updateTeachersSummary(teachersSheet, summarySheet) {
  // Get teachers data
  const teachersData = teachersSheet.getRange(2, 1, teachersSheet.getLastRow() - 1, 3).getValues();
  const periodCols = [];
  for (let col = TEACHERS_FIRST_PERIOD; col <= TEACHERS_LAST_PERIOD; col++) {
    if (col !== TEACHERS_BREAK_COL && col !== TEACHERS_LUNCH_COL) {
      periodCols.push(col);
    }
  }
  
  // Process each teacher
  const summaryData = teachersData.map((row, index) => {
    if (!row[1]) return null; // Skip empty rows
    
    // Count periods for this teacher
    let periodCount = 0;
    periodCols.forEach(col => {
      const cellValue = teachersSheet.getRange(index + 2, col).getValue();
      if (cellValue) periodCount++;
    });
    
    return [index + 1, row[1], periodCount];
  }).filter(row => row !== null);
  
  // Write to summary sheet
  if (summaryData.length > 0) {
    summarySheet.getRange(3, 1, summaryData.length, 3).setValues(summaryData);
  }
}

function updateClasswiseSummary(classesSheet, summarySheet) {
  // Get classes data
  const classesData = classesSheet.getRange(2, 1, classesSheet.getLastRow() - 1, 2).getValues();
  const subjectsConfig = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SUBJECTS_NAME);
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  
  // Process each class
  const summaryData = classesData.map((row, index) => {
    if (!row[1]) return null; // Skip empty rows
    
    const className = row[1];
    const subjectCounts = subjects.map(subject => {
      let count = 0;
      for (let col = CLASSES_FIRST_PERIOD; col <= CLASSES_LAST_PERIOD; col++) {
        if (col !== CLASSES_BREAK_COL && col !== CLASSES_LUNCH_COL) {
          const cellValue = classesSheet.getRange(index + 2, col).getValue();
          if (cellValue && cellValue.includes(subject)) count++;
        }
      }
      return count;
    });
    
    return [index + 1, className, ...subjectCounts];
  }).filter(row => row !== null);
  
  // Write to summary sheet with dynamic column count
  if (summaryData.length > 0) {
    const totalColumns = 2 + subjects.length; // SI + Class Name + Subject columns
    summarySheet.getRange(3, 5, summaryData.length, totalColumns).setValues(summaryData);
  }
}

// Add SI numbers automatically
function addSINumbers(sheet, startRow, column, count) {
  const numbers = Array(count).fill().map((_, i) => [i + 1]);
  sheet.getRange(startRow, column, count, 1).setValues(numbers);
}

// Add function to update summary headers based on subjects
function updateSummaryHeaders(ss) {
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  
  // Update right table headers
  const rightTableHeaders = ['SI', 'Classes', ...subjects];
  summarySheet.getRange(2, 5, 1, rightTableHeaders.length).setValues([rightTableHeaders]);
  
  // Update column widths for subject columns
  for (let i = 0; i < subjects.length; i++) {
    summarySheet.setColumnWidth(7 + i, 100);
  }
}

// Add function to clear all data while keeping structure
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Data',
    'This will clear all data while keeping the structure. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    
    // Clear Teachers sheet data
    const teachersLastRow = Math.max(2, teachersSheet.getLastRow());
    teachersSheet.getRange(2, 2, teachersLastRow - 1, teachersSheet.getLastColumn() - 1).clearContent();
    
    // Clear Classes sheet data
    const classesLastRow = Math.max(2, classesSheet.getLastRow());
    classesSheet.getRange(2, 2, classesLastRow - 1, classesSheet.getLastColumn() - 1).clearContent();
    
    // Clear Summary sheet data
    const summaryLastRow = Math.max(3, summarySheet.getLastRow());
    summarySheet.getRange(3, 1, summaryLastRow - 2, 3).clearContent(); // Left table
    summarySheet.getRange(3, 5, summaryLastRow - 2, summarySheet.getLastColumn() - 4).clearContent(); // Right table
    
    // Restore SI numbers
    addSINumbers(teachersSheet, 2, 1, 20);
    addSINumbers(classesSheet, 2, 1, 20);
    
    // Refresh dropdowns
    setupDropdowns(ss);
    
    ui.alert('Success', 'All data has been cleared while keeping the structure.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to clear data: ' + error.message, ui.ButtonSet.OK);
  }
}

// Add this function to protect headers
function protectHeaders(sheet, numHeaderRows) {
  try {
    const protection = sheet.getRange(1, 1, numHeaderRows, sheet.getLastColumn()).protect();
    protection.setDescription('Header - Protected');
    protection.setWarningOnly(true);
  } catch (error) {
    console.log('Error protecting headers: ' + error.message);
  }
}

// Update sync functions to handle empty values
function syncTeacherToClass(ss, teacherRow, teacherCol) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Skip if Break or Lunch column
  if (teacherCol === TEACHERS_BREAK_COL || teacherCol === TEACHERS_LUNCH_COL) return;
  
  // Get the value from teacher's cell
  const teacherCell = teachersSheet.getRange(teacherRow, teacherCol);
  const teacherValue = teacherCell.getValue();
  
  // Get teacher info
  const teacherName = teachersSheet.getRange(teacherRow, 2).getValue();
  const teacherSubject = teachersSheet.getRange(teacherRow, 3).getValue();
  
  // If teacher has name and subject
  if (teacherName && teacherSubject) {
    // If value exists, update corresponding class cell
    if (teacherValue) {
      // Find the class row
      const classesData = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1).getValues();
      const classRow = classesData.findIndex(row => {
        const className = row[0];
        return className === teacherValue || className.replace(' - ', ' ') === teacherValue;
      }) + 2;
      
      if (classRow >= 2) {
        const classCol = teacherCol - 1; // Classes sheet columns are offset by 1
        const classCell = classesSheet.getRange(classRow, classCol);
        classCell.setValue(`${teacherName} / ${teacherSubject}`);
      }
    } else {
      // If value is empty (deletion), clear corresponding class cells
      const classesData = classesSheet.getRange(2, 2, classesSheet.getLastRow() - 1, 1).getValues();
      classesData.forEach((row, index) => {
        if (row[0]) { // If class exists
          const classRow = index + 2;
          const classCol = teacherCol - 1;
          const classCell = classesSheet.getRange(classRow, classCol);
          const classCellValue = classCell.getValue();
          // Only clear if this cell contains this teacher
          if (classCellValue && classCellValue.startsWith(`${teacherName} /`)) {
            classCell.clearContent();
          }
        }
      });
    }
  }
}

function syncClassToTeacher(ss, classRow, classCol) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Skip if Break or Lunch column
  if (classCol === CLASSES_BREAK_COL || classCol === CLASSES_LUNCH_COL) return;
  
  // Get the value from class's cell
  const classCell = classesSheet.getRange(classRow, classCol);
  const classValue = classCell.getValue();
  const className = classesSheet.getRange(classRow, 2).getValue();
  
        if (className) {
    if (classValue) {
      // If value exists, update corresponding teacher cell
      const [teacherName] = classValue.split(' / ');
      
      // Find the teacher row with matching name and subject
      const teachersData = teachersSheet.getRange(2, 2, teachersSheet.getLastRow() - 1, 2).getValues();
      const teacherRow = teachersData.findIndex(row => row[0] === teacherName) + 2;
      
      if (teacherRow >= 2) {
        const teacherCol = classCol + 1; // Teachers sheet columns are offset by 1
        const teacherCell = teachersSheet.getRange(teacherRow, teacherCol);
        teacherCell.setValue(className);
      }
        } else {
      // If value is empty (deletion), clear corresponding teacher cells
      const teachersData = teachersSheet.getRange(2, 2, teachersSheet.getLastRow() - 1, 1).getValues();
      teachersData.forEach((row, index) => {
        if (row[0]) { // If teacher exists
          const teacherRow = index + 2;
          const teacherCol = classCol + 1;
          const teacherCell = teachersSheet.getRange(teacherRow, teacherCol);
          const teacherCellValue = teacherCell.getValue();
          // Only clear if this cell contains this class
          if (teacherCellValue === className || 
              teacherCellValue.replace(' - ', ' ') === className) {
            teacherCell.clearContent();
          }
        }
      });
    }
  }
}

// Add function to update all class period dropdowns
function updateAllClassPeriodDropdowns(ss) {
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const lastRow = Math.max(2, classesSheet.getLastRow());
  
  for (let row = 2; row <= lastRow; row++) {
    setupClassPeriodDropdowns(ss, row);
  }
}

// Add function to update all teacher period dropdowns
function updateAllTeacherPeriodDropdowns(ss) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const lastRow = Math.max(2, teachersSheet.getLastRow());
  
  for (let row = 2; row <= lastRow; row++) {
    setupTeacherPeriodDropdowns(ss, row);
  }
}

// Add function to handle config sheet formatting
function formatConfigSheet(sheet, numColumns, newRow) {
  // Get the current last row
  const lastRow = sheet.getLastRow();
  
  // Format the new row
  const range = sheet.getRange(newRow, 1, 1, numColumns);
  const color = (newRow % 2 === 0) ? 'white' : '#f8f9fa';
  range.setBackground(color)
       .setBorder(true, true, true, true, true, true)
       .setVerticalAlignment('middle');
  
  // Center align ID column
  sheet.getRange(newRow, 1, 1, 1).setHorizontalAlignment('center');
  
  // Add ID based on sheet type
  const prefix = sheet.getName().includes('Teachers') ? 'T' : 
                sheet.getName().includes('Classes') ? 'C' : 'S';
  sheet.getRange(newRow, 1).setValue(prefix + String(newRow - 1).padStart(3, '0'));
  
  // Update the thick border for the entire table
  sheet.getRange(1, 1, newRow, numColumns)
    .setBorder(true, true, true, true, null, null,
              '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Add subject validation if it's the Teachers config sheet
  if (sheet.getName() === CONFIG_TEACHERS_NAME) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
    const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
    const subjectValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(subjects)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(newRow, 3).setDataValidation(subjectValidation);
  }
  
  // Add ID format validation
  const idRule = SpreadsheetApp.newDataValidation()
    .requireTextContains(prefix)
    .setHelpText(`ID must start with ${prefix} followed by numbers (e.g., ${prefix}001)`)
    .build();
  sheet.getRange(newRow, 1).setDataValidation(idRule);
  
  // Protect headers
  protectHeaders(sheet, 1);
  
  // Expand the main sheets and summary sheet when config sheets are expanded
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheet.getName() === CONFIG_TEACHERS_NAME || sheet.getName() === CONFIG_CLASSES_NAME) {
    setupEmptyMainSheets(ss);
    setupEmptySummarySheet(ss);
  } else if (sheet.getName() === CONFIG_SUBJECTS_NAME) {
    setupEmptySummarySheet(ss);
    setupDropdowns(ss);
  }
}

// Add function to handle main sheet formatting
function formatMainSheet(sheet, startRow, endRow) {
  const lastColumn = sheet.getLastColumn();
  
  // Format all data rows
  const dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, lastColumn);
  dataRange.setBorder(true, true, true, true, true, true)
          .setVerticalAlignment('middle');
  
  // Add alternating colors
  for (let row = startRow; row <= endRow; row++) {
    const color = (row % 2 === 0) ? 'white' : '#f8f9fa';
    sheet.getRange(row, 1, 1, lastColumn).setBackground(color);
  }
  
  // Add thick border around the entire table
  sheet.getRange(1, 1, endRow, lastColumn)
    .setBorder(true, true, true, true, null, null,
              '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Update SI numbers
  for (let row = startRow; row <= endRow; row++) {
    sheet.getRange(row, 1).setValue(row - 1);
  }
  
  // Clear any content and formatting below the table
  const totalRows = sheet.getMaxRows();
  if (totalRows > endRow) {
    sheet.deleteRows(endRow + 1, totalRows - endRow);
  }
}

function setupEmptyMainSheets(ss) {
  // Get the number of teachers and classes from config sheets
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  
  const numTeachers = Math.max(teachersConfig.getLastRow() - 1, 1);
  const numClasses = Math.max(classesConfig.getLastRow() - 1, 1);
  
  // Set up empty Teachers sheet
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  teachersSheet.clear();
  
  // Set up header rows
  const teachersHeader = [
    'SI',
    'Teacher Name',
    'Subject',
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
  teachersSheet.getRange(1, 1, 1, teachersHeader.length).setValues([teachersHeader]);
  
  // Format headers
  teachersSheet.getRange(1, 1, 1, teachersHeader.length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // Set column widths
  teachersSheet.setColumnWidth(1, 30);   // SI column
  teachersSheet.setColumnWidth(2, 150);  // Teacher Name column
  teachersSheet.setColumnWidth(3, 100);  // Subject column
  for (let i = 4; i <= teachersHeader.length; i++) {
    teachersSheet.setColumnWidth(i, 100);
  }
  
  // Set row height for header
  teachersSheet.setRowHeight(1, 60);
  
  // Create empty rows based on number of teachers
  const emptyTeacherRows = Array(numTeachers).fill().map(() => 
    Array(teachersHeader.length).fill('')
  );
  teachersSheet.getRange(2, 1, numTeachers, teachersHeader.length).setValues(emptyTeacherRows);
  
  // Format the table area
  const teachersTableRange = teachersSheet.getRange(1, 1, numTeachers + 1, teachersHeader.length);
  teachersTableRange.setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // Add thick border around the entire table
  teachersTableRange.setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Add alternating row colors
  for (let i = 0; i < numTeachers; i++) {
    const rowNumber = i + 2;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    teachersSheet.getRange(rowNumber, 1, 1, teachersHeader.length).setBackground(color);
  }
  
  // Set up empty Classes sheet
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  classesSheet.clear();
  
  // Set up header rows
  const classesHeader = [
    'SI',
    'Classes',
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
  classesSheet.getRange(1, 1, 1, classesHeader.length).setValues([classesHeader]);
  
  // Format headers
  classesSheet.getRange(1, 1, 1, classesHeader.length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // Set column widths
  classesSheet.setColumnWidth(1, 30);  // SI column
  classesSheet.setColumnWidth(2, 200); // Classes column
  for (let i = 3; i <= classesHeader.length; i++) {
    classesSheet.setColumnWidth(i, 100);
  }
  
  // Set row height for header
  classesSheet.setRowHeight(1, 60);
  
  // Create empty rows based on number of classes
  const emptyClassRows = Array(numClasses).fill().map(() => 
    Array(classesHeader.length).fill('')
  );
  classesSheet.getRange(2, 1, numClasses, classesHeader.length).setValues(emptyClassRows);
  
  // Format the table area
  const classesTableRange = classesSheet.getRange(1, 1, numClasses + 1, classesHeader.length);
  classesTableRange.setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // Add thick border around the entire table
  classesTableRange.setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Add alternating row colors
  for (let i = 0; i < numClasses; i++) {
    const rowNumber = i + 2;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    classesSheet.getRange(rowNumber, 1, 1, classesHeader.length).setBackground(color);
  }
  
  // Protect headers
  protectHeaders(teachersSheet, 1);
  protectHeaders(classesSheet, 1);
  
  // Color Break and Lunch columns
  teachersSheet.getRange(1, TEACHERS_BREAK_COL, numTeachers + 1).setBackground('#ffcdd2');
  teachersSheet.getRange(1, TEACHERS_LUNCH_COL, numTeachers + 1).setBackground('#ffcdd2');
  classesSheet.getRange(1, CLASSES_BREAK_COL, numClasses + 1).setBackground('#ffcdd2');
  classesSheet.getRange(1, CLASSES_LUNCH_COL, numClasses + 1).setBackground('#ffcdd2');
  
  // Clear any existing data validations
  if (teachersSheet.getMaxRows() > 1) {
    teachersSheet.getRange(2, 1, teachersSheet.getMaxRows() - 1, teachersSheet.getLastColumn()).clearDataValidations();
  }
  if (classesSheet.getMaxRows() > 1) {
    classesSheet.getRange(2, 1, classesSheet.getMaxRows() - 1, classesSheet.getLastColumn()).clearDataValidations();
  }
  
  // Clear any content and formatting below the tables
  const teachersTotalRows = teachersSheet.getMaxRows();
  if (teachersTotalRows > numTeachers + 1) {
    teachersSheet.deleteRows(numTeachers + 2, teachersTotalRows - (numTeachers + 1));
  }
  
  const classesTotalRows = classesSheet.getMaxRows();
  if (classesTotalRows > numClasses + 1) {
    classesSheet.deleteRows(numClasses + 2, classesTotalRows - (numClasses + 1));
  }
}

function setupEmptySummarySheet(ss) {
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  summarySheet.clear();
  
  // Get the number of teachers and classes from config sheets
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  
  const numTeachers = Math.max(teachersConfig.getLastRow() - 1, 1);
  const numClasses = Math.max(classesConfig.getLastRow() - 1, 1);
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  
  // Set up headers for Teachers table (left side)
  const teachersHeaders = [
    ['Teachers Summary'],
    ['SI', 'Teacher Name', 'Total\nPeriods']
  ];
  
  // Set up headers for Classwise Subject Allotment table (right side)
  const classHeaders = [
    ['Classwise Subject Allotment'],
    ['SI', 'Classes', ...subjects]
  ];
  
  const rightTableWidth = classHeaders[1].length;
  
  // Write headers
  summarySheet.getRange(1, 1, 1, 3).merge().setValue(teachersHeaders[0][0]);
  summarySheet.getRange(2, 1, 1, 3).setValues([teachersHeaders[1]]);
  
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
         .setHorizontalAlignment('center');
  });
  
  // Set column widths
  summarySheet.setColumnWidth(1, 30);   // SI
  summarySheet.setColumnWidth(2, 150);  // Teacher Name
  summarySheet.setColumnWidth(3, 80);   // Total Periods
  summarySheet.setColumnWidth(4, 30);   // Gap between tables
  summarySheet.setColumnWidth(5, 30);   // SI
  summarySheet.setColumnWidth(6, 150);  // Classes
  for (let i = 7; i < 7 + subjects.length; i++) {    // Subject columns
    summarySheet.setColumnWidth(i, 100);
  }
  
  // Set row heights
  summarySheet.setRowHeight(1, 30);
  summarySheet.setRowHeight(2, 40);
  
  // Add empty rows based on actual data
  const maxRows = Math.max(numTeachers, numClasses);
  const emptyTeacherRows = Array(maxRows).fill(['', '', '']);
  const emptyClassRows = Array(maxRows).fill(['', '', ...Array(subjects.length).fill('')]);
  
  // Add empty rows to both tables
  summarySheet.getRange(3, 1, maxRows, 3).setValues(emptyTeacherRows);
  summarySheet.getRange(3, 5, maxRows, rightTableWidth).setValues(emptyClassRows);
  
  // Format data areas
  const leftTable = summarySheet.getRange(1, 1, maxRows + 2, 3);
  const rightTable = summarySheet.getRange(1, 5, maxRows + 2, rightTableWidth);
  
  [leftTable, rightTable].forEach(range => {
    range.setBorder(true, true, true, true, true, true)
         .setVerticalAlignment('middle')
         .setFontSize(10);
  });
  
  // Set alignments for data areas
  summarySheet.getRange(3, 1, maxRows, 1).setHorizontalAlignment('center'); // Left SI
  summarySheet.getRange(3, 2, maxRows, 1).setHorizontalAlignment('left');   // Teacher names
  summarySheet.getRange(3, 3, maxRows, 1).setHorizontalAlignment('center'); // Total periods
  summarySheet.getRange(3, 5, maxRows, 1).setHorizontalAlignment('center'); // Right SI
  summarySheet.getRange(3, 6, maxRows, 1).setHorizontalAlignment('left');   // Class names
  summarySheet.getRange(3, 7, maxRows, subjects.length).setHorizontalAlignment('center'); // Subject columns
  
  // Add alternating row colors
  for (let i = 0; i < maxRows; i++) {
    const rowNumber = i + 3;
    const color = i % 2 === 0 ? 'white' : '#f8f9fa';
    summarySheet.getRange(rowNumber, 1, 1, 3).setBackground(color);
    summarySheet.getRange(rowNumber, 5, 1, rightTableWidth).setBackground(color);
  }
  
  // Add thick colored borders to both tables
  leftTable.setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  rightTable.setBorder(
    true, true, true, true, null, null,
    '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK
  );
  
  // Protect header rows
  protectHeaders(summarySheet, 2);
  
  // Clear any content and formatting below the tables
  const totalRows = summarySheet.getMaxRows();
  if (totalRows > maxRows + 2) {
    summarySheet.deleteRows(maxRows + 3, totalRows - (maxRows + 2));
  }
}

function addNewTeacher() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  
  // Show config sheet
  teachersConfig.showSheet();
  
  // Get last row and generate next ID
  const lastRow = teachersConfig.getLastRow();
  const nextId = 'T' + String(lastRow).padStart(3, '0');
  
  // Add new row
  const newRow = lastRow + 1;
  teachersConfig.insertRowAfter(lastRow);
  
  // Format the new row
  formatConfigSheet(teachersConfig, 3, newRow);
  
  // Add subject validation to the new row
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  const subjectValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(subjects)
    .setAllowInvalid(false)
    .build();
  teachersConfig.getRange(newRow, 3).setDataValidation(subjectValidation);
  
  // Expand Teachers sheet
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const teachersLastRow = teachersSheet.getLastRow();
  teachersSheet.insertRowAfter(teachersLastRow);
  
  // Format new row in Teachers sheet
  const newTeacherRow = teachersLastRow + 1;
  const color = (newTeacherRow % 2 === 0) ? 'white' : '#f8f9fa';
  const range = teachersSheet.getRange(newTeacherRow, 1, 1, teachersSheet.getLastColumn());
  range.setBackground(color)
       .setBorder(true, true, true, true, true, true)
       .setVerticalAlignment('middle');
  
  // Update SI number
  teachersSheet.getRange(newTeacherRow, 1).setValue(newTeacherRow - 1);
  
  // Update Summary sheet
  setupEmptySummarySheet(ss);
  
  ui.alert('New teacher row added. Please enter the teacher details in row ' + newRow);
}

function addNewClass() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  
  // Show config sheet
  classesConfig.showSheet();
  
  // Get last row and generate next ID
  const lastRow = classesConfig.getLastRow();
  const nextId = 'C' + String(lastRow).padStart(3, '0');
  
  // Add new row
  const newRow = lastRow + 1;
  classesConfig.insertRowAfter(lastRow);
  
  // Format the new row
  formatConfigSheet(classesConfig, 3, newRow);
  
  // Expand Classes sheet
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const classesLastRow = classesSheet.getLastRow();
  classesSheet.insertRowAfter(classesLastRow);
  
  // Format new row in Classes sheet
  const newClassRow = classesLastRow + 1;
  const color = (newClassRow % 2 === 0) ? 'white' : '#f8f9fa';
  const range = classesSheet.getRange(newClassRow, 1, 1, classesSheet.getLastColumn());
  range.setBackground(color)
       .setBorder(true, true, true, true, true, true)
       .setVerticalAlignment('middle');
  
  // Update SI number
  classesSheet.getRange(newClassRow, 1).setValue(newClassRow - 1);
  
  // Update Summary sheet
  setupEmptySummarySheet(ss);
  
  ui.alert('New class row added. Please enter the class details in row ' + newRow);
}

function addNewSubject() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  
  // Show config sheet
  subjectsConfig.showSheet();
  
  // Get last row and generate next ID
  const lastRow = subjectsConfig.getLastRow();
  const nextId = 'S' + String(lastRow).padStart(3, '0');
  
  // Add new row
  const newRow = lastRow + 1;
  subjectsConfig.insertRowAfter(lastRow);
  
  // Format the new row
  formatConfigSheet(subjectsConfig, 2, newRow);
  
  // Update subject validation in Teachers Config sheet
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  const subjectValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(subjects)
    .setAllowInvalid(false)
    .build();
  
  // Apply validation to all rows in Teachers Config
  const teachersLastRow = teachersConfig.getLastRow();
  if (teachersLastRow > 1) {
    teachersConfig.getRange(2, 3, teachersLastRow - 1, 1).setDataValidation(subjectValidation);
  }
  
  // Update Summary sheet to include new subject column
  setupEmptySummarySheet(ss);
  
  // Update dropdowns in main sheets
  setupDropdowns(ss);
  
  ui.alert('New subject row added. Please enter the subject details in row ' + newRow);
} 