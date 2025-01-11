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

// Cache management
const Cache = {
  TIMEOUT: 21600, // 6 hours in seconds
  
  getCache() {
    return CacheService.getScriptCache();
  },
  
  get(key) {
    try {
      const cache = this.getCache();
      const data = cache.get(key);
      return data ? JSON.parse(data) : null;
    } catch (error) {
      console.error(`Cache get error for key ${key}:`, error);
      return null;
    }
  },
  
  set(key, data, timeout = this.TIMEOUT) {
    try {
      const cache = this.getCache();
      cache.put(key, JSON.stringify(data), timeout);
      return true;
    } catch (error) {
      console.error(`Cache set error for key ${key}:`, error);
      return false;
    }
  },
  
  remove(key) {
    try {
      const cache = this.getCache();
      cache.remove(key);
      return true;
    } catch (error) {
      console.error(`Cache remove error for key ${key}:`, error);
      return false;
    }
  },
  
  clear() {
    try {
      const cache = this.getCache();
      cache.removeAll(['configData', 'teachersData', 'classesData']);
      return true;
    } catch (error) {
      console.error('Cache clear error:', error);
      return false;
    }
  },
  
  withCache(key, operation, timeout = this.TIMEOUT) {
    let data = this.get(key);
    if (!data) {
      data = operation();
      this.set(key, data, timeout);
    }
    return data;
  }
};

// Replace existing cache functions with new Cache object
function getConfigData(forceRefresh = false) {
  if (!forceRefresh) {
    const cachedData = Cache.get('configData');
    if (cachedData) return cachedData;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    teachers: SheetManager.getRequiredSheet(ss, CONFIG_TEACHERS_NAME),
    classes: SheetManager.getRequiredSheet(ss, CONFIG_CLASSES_NAME),
    subjects: SheetManager.getRequiredSheet(ss, CONFIG_SUBJECTS_NAME)
  };
  
  const configData = {
    teachers: sheets.teachers.getDataRange().getValues().slice(1),
    classes: sheets.classes.getDataRange().getValues().slice(1),
    subjects: sheets.subjects.getDataRange().getValues().slice(1)
  };
  
  Cache.set('configData', configData);
  return configData;
}

// Add function to invalidate config cache
function invalidateConfigCache() {
  const cache = getCache();
  cache.remove('configData');
}

// Optimized setup function
function setupStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    Transaction.start('SETUP_IN_PROGRESS');
    
    // List of all required sheets
    const requiredSheets = [
      {
        name: CONFIG_TEACHERS_NAME,
        headers: ['ID', 'Teacher Name', 'Subject'],
        columnWidths: [100, 200, 150],
        isConfig: true
      },
      {
        name: CONFIG_CLASSES_NAME,
        headers: ['ID', 'Class Name', 'Section'],
        columnWidths: [100, 200, 100],
        isConfig: true
      },
      {
        name: CONFIG_SUBJECTS_NAME,
        headers: ['ID', 'Subject Name'],
        columnWidths: [100, 200],
        isConfig: true
      },
      {
        name: TEACHERS_SHEET_NAME,
        headers: ['SI', 'Teacher Name', 'Subject', '1', '2', '3', 'Break', '4', '5', '6', 'Lunch', '7', '8', '9'],
        columnWidths: [50, 200, 150, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100],
        isConfig: false
      },
      {
        name: CLASSES_SHEET_NAME,
        headers: ['SI', 'Class', '1', '2', '3', 'Break', '4', '5', '6', 'Lunch', '7', '8', '9'],
        columnWidths: [50, 200, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100],
        isConfig: false
      },
      {
        name: SUMMARY_SHEET_NAME,
        isConfig: false
      }
    ];
    
    // Create or get each required sheet
    requiredSheets.forEach(config => {
      let sheet = ss.getSheetByName(config.name);
      
      if (!sheet) {
        sheet = ss.insertSheet(config.name);
      }
      
      if (config.headers) {
        const headerRange = sheet.getRange(1, 1, 1, config.headers.length);
        headerRange.setValues([config.headers])
                  .setBackground('#f3f3f3')
                  .setFontWeight('bold')
                  .setBorder(true, true, true, true, true, true);
        
        // Set column widths
        config.columnWidths.forEach((width, index) => {
          sheet.setColumnWidth(index + 1, width);
        });
        
        // Special formatting for period columns
        if (config.name === TEACHERS_SHEET_NAME || config.name === CLASSES_SHEET_NAME) {
          // Format break and lunch columns
          const breakCol = config.name === TEACHERS_SHEET_NAME ? TEACHERS_BREAK_COL : CLASSES_BREAK_COL;
          const lunchCol = config.name === TEACHERS_SHEET_NAME ? TEACHERS_LUNCH_COL : CLASSES_LUNCH_COL;
          
          sheet.getRange(1, breakCol, sheet.getMaxRows(), 1).setBackground('#ffe0b2');
          sheet.getRange(1, lunchCol, sheet.getMaxRows(), 1).setBackground('#ffe0b2');
        }
      }
      
      if (config.isConfig) {
        sheet.hideSheet();
      }
    });
    
    // Setup Summary sheet structure
    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    setupSummaryStructure(summarySheet);
    
    // Reorder visible sheets
    const visibleSheets = [TEACHERS_SHEET_NAME, CLASSES_SHEET_NAME, SUMMARY_SHEET_NAME];
    visibleSheets.forEach((sheetName, index) => {
      const sheet = ss.getSheetByName(sheetName);
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    });
    
    // Clear cache after setup
    Cache.clear();
    
    // End setup transaction
    Transaction.end('SETUP_IN_PROGRESS');
    
    ui.alert('Setup Complete', 'The system has been set up successfully.', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Setup failed:', error);
    
    // Rollback on error
    try {
      if (Transaction.isInProgress('SETUP_IN_PROGRESS')) {
        // Delete partially created sheets
        const sheetsToCheck = [
          CONFIG_TEACHERS_NAME,
          CONFIG_CLASSES_NAME,
          CONFIG_SUBJECTS_NAME,
          TEACHERS_SHEET_NAME,
          CLASSES_SHEET_NAME,
          SUMMARY_SHEET_NAME
        ];
        
        sheetsToCheck.forEach(sheetName => {
          const sheet = ss.getSheetByName(sheetName);
          if (sheet) {
            ss.deleteSheet(sheet);
          }
        });
      }
    } catch (rollbackError) {
      console.error('Rollback failed:', rollbackError);
    }
    
    // Clear setup flag
    Transaction.end('SETUP_IN_PROGRESS');
    
    // Show error to user
    ui.alert(
      'Setup Failed',
      'An error occurred during setup. The system has attempted to rollback changes. Please try again.',
      ui.ButtonSet.OK
    );
  }
}

function setupSummaryStructure(summarySheet) {
  // Set up basic structure for summary sheet
  const headers = [
    ['Teacher Summary', '', ''],
    ['SI', 'Teacher Name', 'Total Periods']
  ];
  
  const headerRange = summarySheet.getRange(1, 1, 2, 3);
  headerRange.setValues(headers);
  
  // Format headers
  summarySheet.getRange(1, 1, 1, 3).merge()
              .setBackground('#f3f3f3')
              .setFontWeight('bold')
              .setHorizontalAlignment('center');
  
  summarySheet.getRange(2, 1, 1, 3)
              .setBackground('#f3f3f3')
              .setFontWeight('bold')
              .setBorder(true, true, true, true, true, true);
  
  // Set column widths
  summarySheet.setColumnWidth(1, 50);   // SI
  summarySheet.setColumnWidth(2, 200);  // Teacher Name
  summarySheet.setColumnWidth(3, 100);  // Total Periods
  
  // Add thick border
  summarySheet.getRange(1, 1, 2, 3)
              .setBorder(true, true, true, true, null, null,
                        '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function handleSetupError(error, ss) {
  console.error('Setup failed:', error);
  
  // Attempt to rollback changes
  try {
    const setupInProgress = PropertiesService.getScriptProperties().getProperty('SETUP_IN_PROGRESS');
    if (setupInProgress) {
      // Delete partially created sheets
      const sheetsToCheck = [
        CONFIG_TEACHERS_NAME,
        CONFIG_CLASSES_NAME,
        CONFIG_SUBJECTS_NAME,
        TEACHERS_SHEET_NAME,
        CLASSES_SHEET_NAME,
        SUMMARY_SHEET_NAME
      ];
      
      sheetsToCheck.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          ss.deleteSheet(sheet);
        }
      });
    }
  } catch (rollbackError) {
    console.error('Rollback failed:', rollbackError);
  }
  
  // Clear setup flag
  PropertiesService.getScriptProperties().deleteProperty('SETUP_IN_PROGRESS');
  
  // Show error to user
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Failed',
    'An error occurred during setup. The system has attempted to rollback changes. Please try again.',
    ui.ButtonSet.OK
  );
}

function applyBatchUpdates(updates) {
  updates.forEach(update => {
    const { sheet, range, values, formatting } = update;
    
    if (values) {
      range.setValues(values);
    }
    
    if (formatting) {
      if (formatting.background) range.setBackground(formatting.background);
      if (formatting.fontWeight) range.setFontWeight(formatting.fontWeight);
      if (formatting.borders) {
        range.setBorder(true, true, true, true, true, true);
      }
    }
  });
  
  SpreadsheetApp.flush();
}

// Optimized update function
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  
  // Get all data at once
  const teachersData = teachersSheet.getDataRange().getValues();
  const updates = [];
  
  // Calculate total periods for each teacher
  const teacherSummary = teachersData.slice(1)
    .map((row, index) => {
      if (!row[1]) return null; // Skip empty rows
      
      const totalPeriods = row.slice(3)
        .filter((cell, i) => {
          const col = i + 4;
          return col !== TEACHERS_BREAK_COL && 
                 col !== TEACHERS_LUNCH_COL && 
                 cell !== '';
        }).length;
      
      return [index + 1, row[1], totalPeriods];
    })
    .filter(row => row !== null)
    .sort((a, b) => b[2] - a[2]) // Sort by total periods
    .map((row, index) => [index + 1, row[1], row[2]]); // Update SI numbers
  
  // Batch update summary
  if (teacherSummary.length > 0) {
    updates.push({
      range: summarySheet.getRange(3, 1, teacherSummary.length, 3),
      values: teacherSummary
    });
  }
  
  // Apply all updates
  batchUpdate(summarySheet, updates);
  
  // Update classwise summary
  updateClasswiseSummary(ss);
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

// Optimized data validation
function validateData(data, type) {
  const errors = [];
  
  switch(type) {
    case 'teacher':
      if (!data.name || !data.subject) {
        errors.push('Teacher name and subject are required');
      }
      if (data.id && !/^T\d{3}$/.test(data.id)) {
        errors.push('Invalid teacher ID format (should be T followed by 3 digits)');
      }
      break;
      
    case 'class':
      if (!data.name) {
        errors.push('Class name is required');
      }
      if (data.id && !/^C\d{3}$/.test(data.id)) {
        errors.push('Invalid class ID format (should be C followed by 3 digits)');
      }
      break;
      
    case 'subject':
      if (!data.name) {
        errors.push('Subject name is required');
      }
      if (data.id && !/^S\d{3}$/.test(data.id)) {
        errors.push('Invalid subject ID format (should be S followed by 3 digits)');
      }
      break;
  }
  
  return errors;
}

// Optimized data management with error handling
function deployFromConfig() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Start deployment transaction
    PropertiesService.getScriptProperties().setProperty('DEPLOY_IN_PROGRESS', 'true');
    
    // Get and validate config data
    const configData = getConfigData(true); // force refresh
    const validationErrors = validateConfigData(configData);
    
    if (validationErrors.length > 0) {
      throw new Error('Configuration validation failed:\n' + validationErrors.join('\n'));
    }
    
    // Get required sheets
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    
    if (!teachersSheet || !classesSheet || !summarySheet) {
      const missing = [
        !teachersSheet && TEACHERS_SHEET_NAME,
        !classesSheet && CLASSES_SHEET_NAME,
        !summarySheet && SUMMARY_SHEET_NAME
      ].filter(Boolean);
      throw new Error('Required sheets not found: ' + missing.join(', '));
    }
    
    // Clear existing data while preserving headers
    clearSheetData(teachersSheet);
    clearSheetData(classesSheet);
    clearSheetData(summarySheet);
    
    // Set up main sheets with config data
    setupEmptyMainSheets(ss);
    
    // Set up summary sheet
    setupEmptySummarySheet(ss);
    
    // Set up dropdowns
    setupDropdowns(ss);
    
    // Update summary data
    updateSummary();
    
    // End deployment transaction
    PropertiesService.getScriptProperties().deleteProperty('DEPLOY_IN_PROGRESS');
    
    ui.alert('Success', 'Data has been successfully deployed from configuration.', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Deployment failed:', error);
    
    // Attempt rollback
    try {
      if (PropertiesService.getScriptProperties().getProperty('DEPLOY_IN_PROGRESS')) {
        const sheets = [
          ss.getSheetByName(TEACHERS_SHEET_NAME),
          ss.getSheetByName(CLASSES_SHEET_NAME),
          ss.getSheetByName(SUMMARY_SHEET_NAME)
        ];
        
        // Reset sheets to empty state
        sheets.forEach(sheet => {
          if (sheet) {
            clearSheetData(sheet);
          }
        });
        
        // Reapply formatting
        setupEmptyMainSheets(ss);
        setupEmptySummarySheet(ss);
      }
    } catch (rollbackError) {
      console.error('Rollback failed:', rollbackError);
    }
    
    // Clear deployment flag
    PropertiesService.getScriptProperties().deleteProperty('DEPLOY_IN_PROGRESS');
    
    // Show error to user
    ui.alert('Error', 'Failed to deploy data: ' + error.message, ui.ButtonSet.OK);
  }
}

function validateConfigData(configData) {
  const errors = [];
  
  // Check for duplicate IDs
  const teacherIds = new Set();
  const classIds = new Set();
  const subjectIds = new Set();
  
  // Validate teachers
  configData.teachers.forEach((teacher, index) => {
    const [id, name, subject] = teacher;
    const rowNum = index + 2; // Add 2 to account for 1-based index and header row
    
    // Check for duplicate ID
    if (teacherIds.has(id)) {
      errors.push(`Duplicate Teacher ID '${id}' found at row ${rowNum}`);
    } else {
      teacherIds.add(id);
    }
    
    // Validate ID format
    if (!id || !/^T\d{3}$/.test(id)) {
      errors.push(`Invalid Teacher ID format at row ${rowNum}. Must be 'T' followed by 3 digits`);
    }
    
    // Validate name and subject
    if (!name || name.trim() === '') {
      errors.push(`Missing Teacher Name at row ${rowNum}`);
    }
    if (!subject || subject.trim() === '') {
      errors.push(`Missing Subject at row ${rowNum}`);
    }
  });
  
  // Validate classes
  configData.classes.forEach((classData, index) => {
    const [id, name, section] = classData;
    const rowNum = index + 2;
    
    // Check for duplicate ID
    if (classIds.has(id)) {
      errors.push(`Duplicate Class ID '${id}' found at row ${rowNum}`);
    } else {
      classIds.add(id);
    }
    
    // Validate ID format
    if (!id || !/^C\d{3}$/.test(id)) {
      errors.push(`Invalid Class ID format at row ${rowNum}. Must be 'C' followed by 3 digits`);
    }
    
    // Validate name
    if (!name || name.trim() === '') {
      errors.push(`Missing Class Name at row ${rowNum}`);
    }
  });
  
  // Validate subjects
  configData.subjects.forEach((subject, index) => {
    const [id, name] = subject;
    const rowNum = index + 2;
    
    // Check for duplicate ID
    if (subjectIds.has(id)) {
      errors.push(`Duplicate Subject ID '${id}' found at row ${rowNum}`);
    } else {
      subjectIds.add(id);
    }
    
    // Validate ID format
    if (!id || !/^S\d{3}$/.test(id)) {
      errors.push(`Invalid Subject ID format at row ${rowNum}. Must be 'S' followed by 3 digits`);
    }
    
    // Validate name
    if (!name || name.trim() === '') {
      errors.push(`Missing Subject Name at row ${rowNum}`);
    }
  });
  
  // Validate subject references
  const validSubjects = new Set(configData.subjects.map(s => s[1]));
  configData.teachers.forEach((teacher, index) => {
    const subject = teacher[2];
    if (!validSubjects.has(subject)) {
      errors.push(`Invalid subject '${subject}' for teacher at row ${index + 2}`);
    }
  });
  
  return errors;
}

function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) { // Preserve header row
    sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
  }
}

function createDeploymentUpdates(configData, teachersSheet, classesSheet) {
  const updates = [];
  
  // Prepare teachers data
  const teachersData = configData.teachers.map((teacher, index) => {
    return [index + 1, teacher[1], teacher[2]]; // SI, Name, Subject
  });
  
  if (teachersData.length > 0) {
    updates.push({
      sheet: teachersSheet,
      range: teachersSheet.getRange(2, 1, teachersData.length, 3),
      values: teachersData
    });
  }
  
  // Prepare classes data
  const classesData = configData.classes.map((classData, index) => {
    return [index + 1, classData[1]]; // SI, Class Name
  });
  
  if (classesData.length > 0) {
    updates.push({
      sheet: classesSheet,
      range: classesSheet.getRange(2, 1, classesData.length, 2),
      values: classesData
    });
  }
  
  return updates;
}

function handleDeploymentError(error, ss) {
  console.error('Deployment failed:', error);
  
  // Attempt to rollback changes
  try {
    const deployInProgress = PropertiesService.getScriptProperties().getProperty('DEPLOY_IN_PROGRESS');
    if (deployInProgress) {
      // Reset sheets to empty state
      const sheets = [
        ss.getSheetByName(TEACHERS_SHEET_NAME),
        ss.getSheetByName(CLASSES_SHEET_NAME),
        ss.getSheetByName(SUMMARY_SHEET_NAME)
      ];
      
      sheets.forEach(sheet => {
        if (sheet) {
          clearSheetData(sheet);
        }
      });
      
      // Reapply formatting
      setupEmptyMainSheets(ss);
      setupEmptySummarySheet(ss);
    }
  } catch (rollbackError) {
    console.error('Rollback failed:', rollbackError);
  }
  
  // Clear deployment flag
  PropertiesService.getScriptProperties().deleteProperty('DEPLOY_IN_PROGRESS');
  
  // Show error to user
  handleError(error, 'deployment');
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
  setupEmptyMainSheets(ss);
  setupEmptyConfigSheets(ss);
  setupEmptySummarySheet(ss);
  setupDropdowns(ss);
  setupBreakLunchColumns(ss);
  protectHeaders(ss);
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

// Optimized dropdown management
function setupDropdowns(ss) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  // Get all config data at once
  const configData = getConfigData();
  
  // Create validation rules once
  const teacherPeriodRules = createTeacherPeriodRules(configData);
  const classPeriodRules = createClassPeriodRules(configData);
  
  // Apply rules in batches
  applyDropdownRules(teachersSheet, teacherPeriodRules);
  applyDropdownRules(classesSheet, classPeriodRules);
}

function createTeacherPeriodRules(configData) {
  const rules = [];
  const periodColumns = [
    [4, 5, 6],    // Period 1-3
    [8, 9, 10],   // Period 4-6
    [12, 13, 14]  // Period 7-9
  ];
  
  // Format class names
  const classOptions = configData.classes.map(row => {
    const className = row[1];
    const section = row[2];
    return section ? `${className} - ${section}` : className;
  });
  
  // Create validation rule for each period group
  periodColumns.forEach(columns => {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(classOptions)
      .setAllowInvalid(false)
      .build();
    
    columns.forEach(col => {
      rules.push({ col, validation });
    });
  });
  
  return rules;
}

function createClassPeriodRules(configData) {
  const rules = [];
  const periodColumns = [
    [3, 4, 5],    // Period 1-3
    [7, 8, 9],    // Period 4-6
    [11, 12, 13]  // Period 7-9
  ];
  
  // Create teacher-subject combinations
  const teacherOptions = configData.teachers.map(row => 
    `${row[1]} / ${row[2]}`
  );
  
  // Create validation rule for each period group
  periodColumns.forEach(columns => {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(teacherOptions)
      .setAllowInvalid(false)
      .build();
    
    columns.forEach(col => {
      rules.push({ col, validation });
    });
  });
  
  return rules;
}

function applyDropdownRules(sheet, rules) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const updates = [];
  rules.forEach(({ col, validation }) => {
    updates.push({
      range: sheet.getRange(2, col, lastRow - 1, 1),
      validation: validation
    });
  });
  
  batchUpdate(sheet, updates);
}

// Improved sheet synchronization with concurrency handling
function syncSheets(e) {
  const lock = LockService.getScriptLock();
  try {
    // Try to get lock for synchronization
    if (!lock.tryLock(10000)) {
      throw new TimetableError(
        'Unable to sync sheets due to concurrent edits',
        'SYNC_ERROR',
        { suggestion: 'Please wait a moment and try again' }
      );
    }
    
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    if (row === 1) return; // Skip header row
    
    // Get all changes at once
    const changes = collectChanges(sheet, range);
    if (changes.length === 0) return;
    
    // Create batch updates
    const updates = createSyncUpdates(sheet.getName(), changes);
    
    // Apply updates in transaction
    const ss = e.source;
    PropertiesService.getScriptProperties().setProperty('SYNC_IN_PROGRESS', 'true');
    
    try {
      applySyncUpdates(ss, updates);
      
      // Update dropdowns only for affected columns
      const affectedColumns = new Set(changes.map(change => change.col));
      updateAffectedDropdowns(ss, sheet.getName(), Array.from(affectedColumns));
      
      // Update summary if needed
      if (changes.some(change => affectsTeacherLoad(change))) {
        updateSummary();
      }
      
    } catch (error) {
      // Attempt recovery
      handleSyncError(error, ss, sheet.getName());
      throw error;
    } finally {
      PropertiesService.getScriptProperties().deleteProperty('SYNC_IN_PROGRESS');
    }
    
  } finally {
    lock.releaseLock();
  }
}

function handleSyncError(error, ss, sourceSheetName) {
  console.error('Sync error:', error);
  
  try {
    // Get the paired sheet name
    const pairedSheetName = sourceSheetName === TEACHERS_SHEET_NAME ? 
      CLASSES_SHEET_NAME : TEACHERS_SHEET_NAME;
    
    // Refresh both sheets
    [sourceSheetName, pairedSheetName].forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        refreshSheet(sheet);
      }
    });
    
    // Refresh dropdowns
    setupDropdowns(ss);
    
  } catch (recoveryError) {
    console.error('Recovery failed:', recoveryError);
    throw new TimetableError(
      'Failed to recover from sync error',
      'RECOVERY_ERROR',
      { originalError: error, recoveryError }
    );
  }
}

function affectsTeacherLoad(change) {
  // Check if the change affects teacher workload
  return change.col >= TEACHERS_FIRST_PERIOD && 
         change.col <= TEACHERS_LAST_PERIOD &&
         change.col !== TEACHERS_BREAK_COL &&
         change.col !== TEACHERS_LUNCH_COL;
}

function collectChanges(sheet, range) {
  const changes = [];
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  const values = range.getValues();
  
  for (let i = 0; i < numRows; i++) {
    for (let j = 0; j < numCols; j++) {
      const currentRow = range.getRow() + i;
      const currentCol = range.getColumn() + j;
      
      // Skip break and lunch columns
      if (isBreakOrLunchColumn(sheet.getName(), currentCol)) continue;
      
      changes.push({
        row: currentRow,
        col: currentCol,
        value: values[i][j]
      });
    }
  }
  
  return changes;
}

function createSyncUpdates(sourceSheetName, changes) {
  const updates = [];
  
  changes.forEach(change => {
    if (sourceSheetName === TEACHERS_SHEET_NAME) {
      updates.push(createTeacherToClassUpdate(change));
    } else if (sourceSheetName === CLASSES_SHEET_NAME) {
      updates.push(createClassToTeacherUpdate(change));
    }
  });
  
  return updates.filter(update => update !== null);
}

function createTeacherToClassUpdate(change) {
  return {
    targetSheet: CLASSES_SHEET_NAME,
    targetCol: change.col - 1,
    sourceRow: change.row,
    value: change.value,
    type: 'teacherToClass'
  };
}

function createClassToTeacherUpdate(change) {
  return {
    targetSheet: TEACHERS_SHEET_NAME,
    targetCol: change.col + 1,
    sourceRow: change.row,
    value: change.value,
    type: 'classToTeacher'
  };
}

function applySyncUpdates(ss, updates) {
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  
  const teacherUpdates = [];
  const classUpdates = [];
  
  updates.forEach(update => {
    const targetSheet = update.targetSheet === TEACHERS_SHEET_NAME ? teachersSheet : classesSheet;
    const range = targetSheet.getRange(update.sourceRow, update.targetCol);
    
    if (update.type === 'teacherToClass') {
      if (update.value) {
        const teacherName = teachersSheet.getRange(update.sourceRow, 2).getValue();
        const subject = teachersSheet.getRange(update.sourceRow, 3).getValue();
        classUpdates.push({
          range,
          values: [[`${teacherName} / ${subject}`]]
        });
      } else {
        classUpdates.push({
          range,
          values: [['']]
        });
      }
    } else {
      if (update.value) {
        const [teacherName] = update.value.split(' / ');
        teacherUpdates.push({
          range,
          values: [[teacherName]]
        });
      } else {
        teacherUpdates.push({
          range,
          values: [['']]
        });
      }
    }
  });
  
  if (teacherUpdates.length > 0) batchUpdate(teachersSheet, teacherUpdates);
  if (classUpdates.length > 0) batchUpdate(classesSheet, classUpdates);
}

function updateAffectedDropdowns(ss, sourceSheetName, affectedColumns) {
  if (sourceSheetName === TEACHERS_SHEET_NAME) {
    affectedColumns.forEach(col => {
      updatePeriodDropdowns(ss, TEACHERS_SHEET_NAME, col);
      updatePeriodDropdowns(ss, CLASSES_SHEET_NAME, col - 1);
    });
  } else {
    affectedColumns.forEach(col => {
      updatePeriodDropdowns(ss, CLASSES_SHEET_NAME, col);
      updatePeriodDropdowns(ss, TEACHERS_SHEET_NAME, col + 1);
    });
  }
}

function isBreakOrLunchColumn(sheetName, col) {
  if (sheetName === TEACHERS_SHEET_NAME) {
    return col === TEACHERS_BREAK_COL || col === TEACHERS_LUNCH_COL;
  } else {
    return col === CLASSES_BREAK_COL || col === CLASSES_LUNCH_COL;
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
  // Set up Teachers sheet
  const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
  if (!teachersSheet) {
    throw new Error('Teachers sheet not found');
  }
  
  // Set up Classes sheet
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  if (!classesSheet) {
    throw new Error('Classes sheet not found');
  }
  
  // Get data from config sheets
  const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
  const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
  
  if (!teachersConfig || !classesConfig) {
    throw new Error('Config sheets not found');
  }
  
  const teacherData = teachersConfig.getRange(2, 1, teachersConfig.getLastRow() - 1, 3).getValues();
  const classData = classesConfig.getRange(2, 1, classesConfig.getLastRow() - 1, 3).getValues();
  
  // Clear existing data
  teachersSheet.clear();
  classesSheet.clear();
  
  // Set up Teachers sheet headers
  const teachersHeaders = [
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
  
  // Set up Classes sheet headers
  const classesHeaders = [
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
  
  // Write headers
  teachersSheet.getRange(1, 1, 1, teachersHeaders.length).setValues([teachersHeaders]);
  classesSheet.getRange(1, 1, 1, classesHeaders.length).setValues([classesHeaders]);
  
  // Format headers
  [teachersSheet, classesSheet].forEach(sheet => {
    const headers = sheet === teachersSheet ? teachersHeaders : classesHeaders;
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#f3f3f3')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 60);
  });
  
  // Set column widths for Teachers sheet
  teachersSheet.setColumnWidth(1, 30);   // SI
  teachersSheet.setColumnWidth(2, 150);  // Teacher Name
  teachersSheet.setColumnWidth(3, 100);  // Subject
  for (let i = 4; i <= teachersHeaders.length; i++) {
    teachersSheet.setColumnWidth(i, 100);
  }
  
  // Set column widths for Classes sheet
  classesSheet.setColumnWidth(1, 30);   // SI
  classesSheet.setColumnWidth(2, 200);  // Classes
  for (let i = 3; i <= classesHeaders.length; i++) {
    classesSheet.setColumnWidth(i, 100);
  }
  
  // Write teacher data
  const teacherRows = teacherData.map((row, index) => {
    return [index + 1, row[1], row[2]]; // SI, Name, Subject
  });
  teachersSheet.getRange(2, 1, teacherRows.length, 3).setValues(teacherRows);
  
  // Write class data
  const classRows = classData.map((row, index) => {
    const className = row[1];
    const section = row[2];
    const formattedName = section ? `${className} - ${section}` : className;
    return [index + 1, formattedName];
  });
  classesSheet.getRange(2, 1, classRows.length, 2).setValues(classRows);
  
  // Format data rows
  [
    { sheet: teachersSheet, rows: teacherRows.length, cols: teachersHeaders.length },
    { sheet: classesSheet, rows: classRows.length, cols: classesHeaders.length }
  ].forEach(({ sheet, rows, cols }) => {
    // Add alternating colors
    for (let i = 0; i < rows; i++) {
      const rowNum = i + 2;
      const color = i % 2 === 0 ? 'white' : '#f8f9fa';
      sheet.getRange(rowNum, 1, 1, cols).setBackground(color);
    }
    
    // Add borders
    sheet.getRange(1, 1, rows + 1, cols)
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle');
    
    // Add thick border
    sheet.getRange(1, 1, rows + 1, cols)
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  
  // Color Break and Lunch columns
  teachersSheet.getRange(1, TEACHERS_BREAK_COL, teacherRows.length + 1).setBackground('#ffcdd2');
  teachersSheet.getRange(1, TEACHERS_LUNCH_COL, teacherRows.length + 1).setBackground('#ffcdd2');
  classesSheet.getRange(1, CLASSES_BREAK_COL, classRows.length + 1).setBackground('#ffcdd2');
  classesSheet.getRange(1, CLASSES_LUNCH_COL, classRows.length + 1).setBackground('#ffcdd2');
  
  // Protect headers
  protectHeaders(teachersSheet, 1);
  protectHeaders(classesSheet, 1);
  
  // Protect Break and Lunch columns
  protectBreakLunchColumns(teachersSheet, TEACHERS_BREAK_COL, TEACHERS_LUNCH_COL, teacherRows.length + 1);
  protectBreakLunchColumns(classesSheet, CLASSES_BREAK_COL, CLASSES_LUNCH_COL, classRows.length + 1);
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
  
  try {
    // Start transaction
    PropertiesService.getScriptProperties().setProperty('ADD_TEACHER_IN_PROGRESS', 'true');
    
    const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
    if (!teachersConfig) {
      throw new Error('Teachers Config sheet not found');
    }
    
    // Show config sheet
    teachersConfig.showSheet();
    
    // Get last row and generate next ID
    const lastRow = teachersConfig.getLastRow();
    const nextId = 'T' + String(lastRow).padStart(3, '0');
    
    // Add new row
    const newRow = lastRow + 1;
    teachersConfig.insertRowAfter(lastRow);
    
    // Format the new row
    const range = teachersConfig.getRange(newRow, 1, 1, 3);
    const color = (newRow % 2 === 0) ? 'white' : '#f8f9fa';
    
    // Basic formatting
    range.setBackground(color)
         .setBorder(true, true, true, true, true, true)
         .setVerticalAlignment('middle');
    
    // Set ID column
    teachersConfig.getRange(newRow, 1)
                 .setValue(nextId)
                 .setHorizontalAlignment('center');
    
    // Add subject validation
    const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
    if (!subjectsConfig) {
      throw new Error('Subjects Config sheet not found');
    }
    
    const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
    const subjectValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(subjects)
      .setAllowInvalid(false)
      .build();
    teachersConfig.getRange(newRow, 3).setDataValidation(subjectValidation);
    
    // Add ID validation
    const teacherIdRule = SpreadsheetApp.newDataValidation()
      .requireTextContains('T')
      .setHelpText('Teacher ID must start with T followed by numbers (e.g., T001)')
      .build();
    teachersConfig.getRange(newRow, 1).setDataValidation(teacherIdRule);
    
    // Update main Teachers sheet
    const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
    if (!teachersSheet) {
      throw new Error('Teachers sheet not found');
    }
    
    const teachersLastRow = teachersSheet.getLastRow();
    teachersSheet.insertRowAfter(teachersLastRow);
    
    // Format new row in Teachers sheet
    const newTeacherRow = teachersLastRow + 1;
    const mainSheetColor = (newTeacherRow % 2 === 0) ? 'white' : '#f8f9fa';
    const mainRange = teachersSheet.getRange(newTeacherRow, 1, 1, teachersSheet.getLastColumn());
    
    mainRange.setBackground(mainSheetColor)
             .setBorder(true, true, true, true, true, true)
             .setVerticalAlignment('middle');
    
    // Set row height
    teachersSheet.setRowHeight(newTeacherRow, 21);
    
    // Update SI number
    teachersSheet.getRange(newTeacherRow, 1).setValue(newTeacherRow - 1);
    
    // Update thick borders for both sheets
    teachersConfig.getRange(1, 1, newRow, 3)
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
                
    teachersSheet.getRange(1, 1, newTeacherRow, teachersSheet.getLastColumn())
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    
    // Update Summary sheet
    setupEmptySummarySheet(ss);
    
    // Refresh dropdowns
    setupDropdowns(ss);
    
    // End transaction
    PropertiesService.getScriptProperties().deleteProperty('ADD_TEACHER_IN_PROGRESS');
    
    ui.alert('Success', 'New teacher row added. Please enter the teacher details in row ' + newRow, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Error adding new teacher:', error);
    
    // Attempt rollback
    try {
      if (PropertiesService.getScriptProperties().getProperty('ADD_TEACHER_IN_PROGRESS')) {
        // Rollback changes if needed
        const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
        const teachersSheet = ss.getSheetByName(TEACHERS_SHEET_NAME);
        
        if (teachersConfig && teachersConfig.getLastRow() > 1) {
          teachersConfig.deleteRow(teachersConfig.getLastRow());
        }
        
        if (teachersSheet && teachersSheet.getLastRow() > 1) {
          teachersSheet.deleteRow(teachersSheet.getLastRow());
        }
      }
    } catch (rollbackError) {
      console.error('Rollback failed:', rollbackError);
    }
    
    // Clear transaction flag
    PropertiesService.getScriptProperties().deleteProperty('ADD_TEACHER_IN_PROGRESS');
    
    ui.alert('Error', 'Failed to add new teacher: ' + error.message, ui.ButtonSet.OK);
  }
}

function addNewClass() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Start transaction
    PropertiesService.getScriptProperties().setProperty('ADD_CLASS_IN_PROGRESS', 'true');
    
    const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
    if (!classesConfig) {
      throw new Error('Classes Config sheet not found');
    }
    
    // Show config sheet
    classesConfig.showSheet();
    
    // Get last row and generate next ID
    const lastRow = classesConfig.getLastRow();
    const nextId = 'C' + String(lastRow).padStart(3, '0');
    
    // Add new row
    const newRow = lastRow + 1;
    classesConfig.insertRowAfter(lastRow);
    
    // Format the new row
    const range = classesConfig.getRange(newRow, 1, 1, 3);
    const color = (newRow % 2 === 0) ? 'white' : '#f8f9fa';
    
    // Basic formatting
    range.setBackground(color)
         .setBorder(true, true, true, true, true, true)
         .setVerticalAlignment('middle');
    
    // Set ID column
    classesConfig.getRange(newRow, 1)
                .setValue(nextId)
                .setHorizontalAlignment('center');
    
    // Add ID validation
    const classIdRule = SpreadsheetApp.newDataValidation()
      .requireTextContains('C')
      .setHelpText('Class ID must start with C followed by numbers (e.g., C001)')
      .build();
    classesConfig.getRange(newRow, 1).setDataValidation(classIdRule);
    
    // Update main Classes sheet
    const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
    if (!classesSheet) {
      throw new Error('Classes sheet not found');
    }
    
    const classesLastRow = classesSheet.getLastRow();
    classesSheet.insertRowAfter(classesLastRow);
    
    // Format new row in Classes sheet
    const newClassRow = classesLastRow + 1;
    const mainSheetColor = (newClassRow % 2 === 0) ? 'white' : '#f8f9fa';
    const mainRange = classesSheet.getRange(newClassRow, 1, 1, classesSheet.getLastColumn());
    
    mainRange.setBackground(mainSheetColor)
             .setBorder(true, true, true, true, true, true)
             .setVerticalAlignment('middle');
    
    // Set row height
    classesSheet.setRowHeight(newClassRow, 21);
    
    // Update SI number
    classesSheet.getRange(newClassRow, 1).setValue(newClassRow - 1);
    
    // Update thick borders for both sheets
    classesConfig.getRange(1, 1, newRow, 3)
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
                
    classesSheet.getRange(1, 1, newClassRow, classesSheet.getLastColumn())
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    
    // Update Summary sheet
    setupEmptySummarySheet(ss);
    
    // Refresh dropdowns
    setupDropdowns(ss);
    
    // End transaction
    PropertiesService.getScriptProperties().deleteProperty('ADD_CLASS_IN_PROGRESS');
    
    ui.alert('Success', 'New class row added. Please enter the class details in row ' + newRow, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Error adding new class:', error);
    
    // Attempt rollback
    try {
      if (PropertiesService.getScriptProperties().getProperty('ADD_CLASS_IN_PROGRESS')) {
        // Rollback changes if needed
        const classesConfig = ss.getSheetByName(CONFIG_CLASSES_NAME);
        const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
        
        if (classesConfig && classesConfig.getLastRow() > 1) {
          classesConfig.deleteRow(classesConfig.getLastRow());
        }
        
        if (classesSheet && classesSheet.getLastRow() > 1) {
          classesSheet.deleteRow(classesSheet.getLastRow());
        }
      }
    } catch (rollbackError) {
      console.error('Rollback failed:', rollbackError);
    }
    
    // Clear transaction flag
    PropertiesService.getScriptProperties().deleteProperty('ADD_CLASS_IN_PROGRESS');
    
    ui.alert('Error', 'Failed to add new class: ' + error.message, ui.ButtonSet.OK);
  }
}

function addNewSubject() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Start transaction
    PropertiesService.getScriptProperties().setProperty('ADD_SUBJECT_IN_PROGRESS', 'true');
    
    const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
    if (!subjectsConfig) {
      throw new Error('Subjects Config sheet not found');
    }
    
    const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
    if (!teachersConfig) {
      throw new Error('Teachers Config sheet not found');
    }
    
    // Show config sheet
    subjectsConfig.showSheet();
    
    // Get last row and generate next ID
    const lastRow = subjectsConfig.getLastRow();
    const nextId = 'S' + String(lastRow).padStart(3, '0');
    
    // Add new row
    const newRow = lastRow + 1;
    subjectsConfig.insertRowAfter(lastRow);
    
    // Format the new row
    const range = subjectsConfig.getRange(newRow, 1, 1, 2);
    const color = (newRow % 2 === 0) ? 'white' : '#f8f9fa';
    
    // Basic formatting
    range.setBackground(color)
         .setBorder(true, true, true, true, true, true)
         .setVerticalAlignment('middle');
    
    // Set ID column
    subjectsConfig.getRange(newRow, 1)
                 .setValue(nextId)
                 .setHorizontalAlignment('center');
    
    // Add ID validation
    const subjectIdRule = SpreadsheetApp.newDataValidation()
      .requireTextContains('S')
      .setHelpText('Subject ID must start with S followed by numbers (e.g., S001)')
      .build();
    subjectsConfig.getRange(newRow, 1).setDataValidation(subjectIdRule);
    
    // Update thick borders
    subjectsConfig.getRange(1, 1, newRow, 2)
      .setBorder(true, true, true, true, null, null,
                '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    
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
    
    // End transaction
    PropertiesService.getScriptProperties().deleteProperty('ADD_SUBJECT_IN_PROGRESS');
    
    ui.alert('Success', 'New subject row added. Please enter the subject details in row ' + newRow, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Error adding new subject:', error);
    
    // Attempt rollback
    try {
      if (PropertiesService.getScriptProperties().getProperty('ADD_SUBJECT_IN_PROGRESS')) {
        // Rollback changes if needed
        const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
        
        if (subjectsConfig && subjectsConfig.getLastRow() > 1) {
          subjectsConfig.deleteRow(subjectsConfig.getLastRow());
        }
        
        // Restore original subject validation
        const teachersConfig = ss.getSheetByName(CONFIG_TEACHERS_NAME);
        if (teachersConfig) {
          const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
          const subjectValidation = SpreadsheetApp.newDataValidation()
            .requireValueInList(subjects)
            .setAllowInvalid(false)
            .build();
          
          const teachersLastRow = teachersConfig.getLastRow();
          if (teachersLastRow > 1) {
            teachersConfig.getRange(2, 3, teachersLastRow - 1, 1).setDataValidation(subjectValidation);
          }
        }
      }
    } catch (rollbackError) {
      console.error('Rollback failed:', rollbackError);
    }
    
    // Clear transaction flag
    PropertiesService.getScriptProperties().deleteProperty('ADD_SUBJECT_IN_PROGRESS');
    
    ui.alert('Error', 'Failed to add new subject: ' + error.message, ui.ButtonSet.OK);
  }
}

// Error handling utilities
const ErrorTypes = {
  CONFIG: 'CONFIG_ERROR',
  SYNC: 'SYNC_ERROR',
  VALIDATION: 'VALIDATION_ERROR',
  TRANSACTION: 'TRANSACTION_ERROR',
  PERMISSION: 'PERMISSION_ERROR',
  SYSTEM: 'SYSTEM_ERROR'
};

class TimetableError extends Error {
  constructor(message, type = ErrorTypes.SYSTEM, details = {}) {
    super(message);
    this.name = 'TimetableError';
    this.type = type;
    this.details = details;
    this.timestamp = new Date();
  }
  
  static fromError(error, type = ErrorTypes.SYSTEM, additionalDetails = {}) {
    return new TimetableError(
      error.message,
      type,
      { ...additionalDetails, originalError: error }
    );
  }
}

function handleError(error, operation) {
  const ui = SpreadsheetApp.getUi();
  console.error(`Error in ${operation}:`, error);
  
  let userMessage;
  switch (error.type) {
    case 'CONFIG_ERROR':
      userMessage = 'Configuration error: Please check your config sheets and try again.';
      break;
    case 'SYNC_ERROR':
      userMessage = 'Synchronization error: Changes could not be applied. The system will attempt to recover.';
      attemptRecovery(error);
      break;
    case 'VALIDATION_ERROR':
      userMessage = 'Validation error: Please check your input data and try again.';
      break;
    default:
      userMessage = 'An unexpected error occurred. Please try again or contact support.';
  }
  
  ui.alert('Error', userMessage, ui.ButtonSet.OK);
}

function attemptRecovery(error) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Clear cache to ensure fresh data
    clearCache();
    
    // Reload config data
    getConfigData(true); // force refresh
    
    // Resync affected sheets
    if (error.details.affectedSheets) {
      error.details.affectedSheets.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          refreshSheet(sheet);
        }
      });
    }
    
    // Update dropdowns
    setupDropdowns(ss);
    
    // Update summary
    updateSummary();
    
  } catch (recoveryError) {
    console.error('Recovery failed:', recoveryError);
  }
}

function refreshSheet(sheet) {
  const sheetName = sheet.getName();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Store current data
  const currentData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  // Clear and reformat sheet
  sheet.clear();
  sheet.getRange(1, 1, lastRow, lastCol).setValues(currentData);
  
  // Reapply formatting
  if (sheetName === TEACHERS_SHEET_NAME || sheetName === CLASSES_SHEET_NAME) {
    formatMainSheet(sheet, 2, lastRow);
  } else if (sheetName.startsWith('Config_')) {
    formatConfigSheet(sheet, lastCol, lastRow);
  }
}

function updateClasswiseSummary(ss) {
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  const classesSheet = ss.getSheetByName(CLASSES_SHEET_NAME);
  const subjectsConfig = ss.getSheetByName(CONFIG_SUBJECTS_NAME);
  
  // Get all data at once
  const classesData = classesSheet.getDataRange().getValues();
  const subjects = subjectsConfig.getRange(2, 2, subjectsConfig.getLastRow() - 1, 1).getValues().flat();
  const updates = [];
  
  // Process each class
  const classSummary = classesData.slice(1)
    .map((row, index) => {
      if (!row[1]) return null; // Skip empty rows
      
      // Count subjects for each class
      const subjectCounts = subjects.map(() => 0);
      
      // Check each period (skip SI, Class Name, Break, and Lunch)
      row.slice(2).forEach((cell, periodIndex) => {
        if (!cell || 
            periodIndex + 3 === CLASSES_BREAK_COL || 
            periodIndex + 3 === CLASSES_LUNCH_COL) return;
        
        // Extract subject from "Teacher / Subject" format
        const subject = cell.split(' / ')[1];
        const subjectIndex = subjects.indexOf(subject);
        if (subjectIndex !== -1) {
          subjectCounts[subjectIndex]++;
        }
      });
      
      return [index + 1, row[1], ...subjectCounts];
    })
    .filter(row => row !== null);
  
  // Update summary sheet with class data
  if (classSummary.length > 0) {
    const startCol = 5; // Starting column for class summary
    updates.push({
      range: summarySheet.getRange(3, startCol, classSummary.length, classSummary[0].length),
      values: classSummary
    });
  }
  
  // Apply all updates
  batchUpdate(summarySheet, updates);
}

function updatePeriodDropdowns(ss, sheetName, periodCol) {
  // Skip if Break or Lunch column
  if ((sheetName === TEACHERS_SHEET_NAME && 
       (periodCol === TEACHERS_BREAK_COL || periodCol === TEACHERS_LUNCH_COL)) ||
      (sheetName === CLASSES_SHEET_NAME && 
       (periodCol === CLASSES_BREAK_COL || periodCol === CLASSES_LUNCH_COL))) {
    return;
  }
  
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // Get config data
  const configData = getConfigData();
  
  // Create validation rule based on sheet type
  let validation;
  if (sheetName === TEACHERS_SHEET_NAME) {
    // For Teachers sheet, options are class names
    const classOptions = configData.classes.map(row => {
      const className = row[1];
      const section = row[2];
      return section ? `${className} - ${section}` : className;
    });
    
    validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(classOptions)
      .setAllowInvalid(false)
      .build();
  } else {
    // For Classes sheet, options are "Teacher / Subject" combinations
    const teacherOptions = configData.teachers.map(row => 
      `${row[1]} / ${row[2]}`
    );
    
    validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(teacherOptions)
      .setAllowInvalid(false)
      .build();
  }
  
  // Apply validation to the column
  sheet.getRange(2, periodCol, lastRow - 1, 1).setDataValidation(validation);
}

// Utility functions for common operations
const Utils = {
  formatRange: (range, options = {}) => {
    const {
      background = '#f3f3f3',
      fontWeight = 'normal',
      borders = true,
      wrap = true,
      vAlign = 'middle',
      hAlign = 'center'
    } = options;
    
    let rangeFormat = range.setBackground(background)
                          .setVerticalAlignment(vAlign)
                          .setHorizontalAlignment(hAlign);
    
    if (borders) {
      rangeFormat.setBorder(true, true, true, true, true, true);
    }
    if (wrap) {
      rangeFormat.setWrap(true);
    }
    if (fontWeight !== 'normal') {
      rangeFormat.setFontWeight(fontWeight);
    }
    
    return rangeFormat;
  },
  
  addThickBorder: (range) => {
    range.setBorder(true, true, true, true, null, null,
                   '#980000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  },
  
  generateId: (prefix, number) => {
    return `${prefix}${String(number).padStart(3, '0')}`;
  },
  
  createValidation: (options) => {
    const { type, values, helpText } = options;
    let rule;
    
    if (type === 'list') {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values)
        .setAllowInvalid(false);
    } else if (type === 'text') {
      rule = SpreadsheetApp.newDataValidation()
        .requireTextContains(values);
    }
    
    if (helpText) {
      rule.setHelpText(helpText);
    }
    
    return rule.build();
  },
  
  protectRange: (range, description) => {
    const protection = range.protect();
    protection.setDescription(description);
    protection.setWarningOnly(true);
    return protection;
  },
  
  alternateRowColors: (sheet, startRow, numRows, numCols) => {
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      const color = i % 2 === 0 ? 'white' : '#f8f9fa';
      sheet.getRange(row, 1, 1, numCols).setBackground(color);
    }
  }
};

// Transaction management
const Transaction = {
  start: (key) => {
    PropertiesService.getScriptProperties().setProperty(key, 'true');
  },
  
  end: (key) => {
    PropertiesService.getScriptProperties().deleteProperty(key);
  },
  
  isInProgress: (key) => {
    return PropertiesService.getScriptProperties().getProperty(key) === 'true';
  },
  
  withTransaction: async (key, operation) => {
    Transaction.start(key);
    try {
      await operation();
      Transaction.end(key);
    } catch (error) {
      Transaction.end(key);
      throw error;
    }
  }
};

// Sheet management
const SheetManager = {
  getRequiredSheet: (ss, sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    return sheet;
  },
  
  ensureSheetExists: (ss, sheetName) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    return sheet;
  },
  
  clearSheetData: (sheet, preserveHeaders = true) => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > (preserveHeaders ? 1 : 0)) {
      sheet.getRange(preserveHeaders ? 2 : 1, 1, 
                    lastRow - (preserveHeaders ? 1 : 0), lastCol).clear();
    }
  }
};  

// Batch operations
const BatchOperations = {
  BATCH_SIZE: 50,
  
  update(sheet, updates) {
    const chunks = this.chunkUpdates(updates, this.BATCH_SIZE);
    
    chunks.forEach(chunk => {
      const operations = [];
      
      chunk.forEach(update => {
        const { range, values, background, borders, validation, formula, note } = update;
        
        if (values) operations.push(() => range.setValues(values));
        if (background) operations.push(() => range.setBackground(background));
        if (borders) operations.push(() => range.setBorder(...borders));
        if (validation) operations.push(() => range.setDataValidation(validation));
        if (formula) operations.push(() => range.setFormula(formula));
        if (note) operations.push(() => range.setNote(note));
      });
      
      this.executeOperations(operations);
      SpreadsheetApp.flush();
    });
  },
  
  chunkUpdates(updates, size) {
    const chunks = [];
    for (let i = 0; i < updates.length; i += size) {
      chunks.push(updates.slice(i, i + size));
    }
    return chunks;
  },
  
  executeOperations(operations) {
    operations.forEach(operation => {
      try {
        operation();
      } catch (error) {
        console.error('Batch operation failed:', error);
        throw error;
      }
    });
  },
  
  withBatch(sheet, operation) {
    const updates = [];
    operation(updates);
    this.update(sheet, updates);
  }
};

// Replace existing batch update function
function batchUpdate(sheet, updates) {
  return BatchOperations.update(sheet, updates);
}

const ErrorHandler = {
  handle(error, operation) {
    console.error(`Error in ${operation}:`, error);
    
    let userMessage;
    let recoveryAction;
    
    switch (error.type) {
      case ErrorTypes.CONFIG:
        userMessage = 'Configuration error: Please check your config sheets and try again.';
        recoveryAction = () => this.recoverConfig();
        break;
        
      case ErrorTypes.SYNC:
        userMessage = 'Synchronization error: Changes could not be applied. The system will attempt to recover.';
        recoveryAction = () => this.recoverSync(error.details);
        break;
        
      case ErrorTypes.VALIDATION:
        userMessage = 'Validation error: ' + error.message;
        break;
        
      case ErrorTypes.TRANSACTION:
        userMessage = 'Transaction error: The operation could not be completed. Please try again.';
        recoveryAction = () => this.recoverTransaction(error.details);
        break;
        
      case ErrorTypes.PERMISSION:
        userMessage = 'Permission error: You do not have the required permissions for this operation.';
        break;
        
      default:
        userMessage = 'An unexpected error occurred. Please try again or contact support.';
        recoveryAction = () => this.recoverSystem();
    }
    
    if (recoveryAction) {
      try {
        recoveryAction();
      } catch (recoveryError) {
        console.error('Recovery failed:', recoveryError);
        userMessage += ' Recovery failed.';
      }
    }
    
    SpreadsheetApp.getUi().alert('Error', userMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  },
  
  recoverConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Cache.clear();
    setupConfigSheets(ss);
    setupDropdowns(ss);
  },
  
  recoverSync(details) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { sourceSheet, targetSheet } = details;
    
    if (sourceSheet) {
      const sheet = SheetManager.getRequiredSheet(ss, sourceSheet);
      this.refreshSheet(sheet);
    }
    
    if (targetSheet) {
      const sheet = SheetManager.getRequiredSheet(ss, targetSheet);
      this.refreshSheet(sheet);
    }
    
    setupDropdowns(ss);
    updateSummary();
  },
  
  recoverTransaction(details) {
    const { transactionKey } = details;
    if (transactionKey) {
      Transaction.end(transactionKey);
    }
  },
  
  recoverSystem() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Cache.clear();
    setupDropdowns(ss);
    updateSummary();
  },
  
  refreshSheet(sheet) {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    const currentData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    sheet.clear();
    sheet.getRange(1, 1, lastRow, lastCol).setValues(currentData);
    
    if (sheetName === TEACHERS_SHEET_NAME || sheetName === CLASSES_SHEET_NAME) {
      formatMainSheet(sheet, 2, lastRow);
    } else if (sheetName.startsWith('Config_')) {
      formatConfigSheet(sheet, lastCol, lastRow);
    }
  }
};