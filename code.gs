// Google Apps Script Code (Code.gs)
// This file should be created in Google Apps Script (script.google.com)

// Configuration - Update these with your actual Google Sheets IDs
const CONFIG = {
  SPREADSHEET_ID: '1OzQrcODKRHOshteILxowgCXAInVypWE9VoIAolK5AQA', // Replace with your Google Sheets ID
  SHEETS: {
    USERS: 'Users',
    PROFILES: 'Profiles', 
    EXAMS: 'Exams',
    LICENSES: 'Licenses',
    DONATIONS: 'Donations',
    DASHBOARD_STATS: 'Dashboard_Stats'
  }
};

/**
 * Web App Entry Point - Required for Google Apps Script Web App
 */
function doGet(e) {
  try {
    // Return the HTML file
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Alumni & Student Tracking System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('doGet error: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

/**
 * Handle POST requests
 */
function doPost(e) {
  try {
    const action = e.parameter.action;
    const data = JSON.parse(e.parameter.data || '{}');
    
    switch (action) {
      case 'authenticate':
        return ContentService.createTextOutput(JSON.stringify(authenticate(data.email, data.password)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'register':
        return ContentService.createTextOutput(JSON.stringify(register(data)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'saveData':
        return ContentService.createTextOutput(JSON.stringify(saveData(data.sheetName, data.userData)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'loadData':
        return ContentService.createTextOutput(JSON.stringify(loadData(data.sheetName, data.email)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getDashboardStats':
        return ContentService.createTextOutput(JSON.stringify(getDashboardStats()))
          .setMimeType(ContentService.MimeType.JSON);
      
      default:
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    Logger.log('doPost error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Include HTML files (for HtmlService.createHtmlOutputFromFile)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize the spreadsheet with required sheets and headers
 */
function initializeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Create Users sheet
    createSheetIfNotExists(ss, CONFIG.SHEETS.USERS, [
      'Email', 'Password', 'Role', 'UserType', 'FirstName', 'LastName', 
      'RegistrationDate', 'LastLogin', 'IsActive'
    ]);
    
    // Create Profiles sheet
    createSheetIfNotExists(ss, CONFIG.SHEETS.PROFILES, [
      'Email', 'Title', 'FirstNameTh', 'LastNameTh', 'FirstNameEn', 'LastNameEn',
      'StudentId', 'GraduationClass', 'Advisor', 'BirthDate', 'Gender', 
      'BirthCountry', 'Nationality', 'Ethnicity', 'Phone', 'GPAX', 'Address',
      'EmergencyName', 'EmergencyRelation', 'EmergencyPhone', 'Awards',
      'CareerPlan', 'StudyPlan', 'InternationalPlan', 'WillTakeThaiLicense',
      'LastUpdated'
    ]);
    
    // Create Exams sheet
    createSheetIfNotExists(ss, CONFIG.SHEETS.EXAMS, [
      'Email', 'ExamRound', 'ExamSession', 'ExamYear', 'Subject1', 'Subject2',
      'Subject3', 'Subject4', 'Subject5', 'Subject6', 'Subject7', 'Subject8',
      'PassedSubjects', 'AllPassed', 'ExamDate'
    ]);
    
    // Create Licenses sheet
    createSheetIfNotExists(ss, CONFIG.SHEETS.LICENSES, [
      'Email', 'MemberNumber', 'LicenseNumber', 'IssueDate', 'PermitDate',
      'ExpiryDate', 'LastUpdated'
    ]);
    
    // Create Donations sheet
    createSheetIfNotExists(ss, CONFIG.SHEETS.DONATIONS, [
      'Email', 'Amount', 'Purpose', 'OtherPurpose', 'Message', 'TaxName',
      'TaxId', 'TaxAddress', 'Status', 'DonationDate', 'ProcessedDate'
    ]);
    
    // Create Dashboard Stats sheet
    createSheetIfNotExists(ss, CONFIG.SHEETS.DASHBOARD_STATS, [
      'StatType', 'Category', 'Value', 'LastUpdated'
    ]);
    
    Logger.log('Spreadsheet initialized successfully');
    return { success: true, message: 'Spreadsheet initialized' };
    
  } catch (error) {
    Logger.log('Error initializing spreadsheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Create a sheet if it doesn't exist
 */
function createSheetIfNotExists(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  
  return sheet;
}

/**
 * User Authentication
 */
function authenticate(email, password) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password');
    const roleCol = headers.indexOf('Role');
    const nameCol = headers.indexOf('FirstName');
    const lastNameCol = headers.indexOf('LastName');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email && data[i][passwordCol] === password) {
        // Update last login
        sheet.getRange(i + 1, headers.indexOf('LastLogin') + 1).setValue(new Date());
        
        return {
          success: true,
          user: {
            email: data[i][emailCol],
            role: data[i][roleCol],
            name: `${data[i][nameCol]} ${data[i][lastNameCol]}`,
            userType: data[i][roleCol]
          }
        };
      }
    }
    
    return { success: false, error: 'Invalid credentials' };
    
  } catch (error) {
    Logger.log('Authentication error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * User Registration
 */
function register(userData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    // Check if email already exists
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === userData.email) {
        return { success: false, error: 'Email already exists' };
      }
    }
    
    // Add new user
    const newRow = [
      userData.email,
      userData.password,
      userData.userType,
      userData.userType,
      userData.firstName,
      userData.lastName,
      new Date(),
      '',
      true
    ];
    
    sheet.appendRow(newRow);
    
    return { success: true, message: 'User registered successfully' };
    
  } catch (error) {
    Logger.log('Registration error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save data to specified sheet
 */
function saveData(sheetName, data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, error: `Sheet ${sheetName} not found` };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Check if record exists (based on email)
    const emailCol = headers.indexOf('Email');
    if (emailCol === -1) {
      return { success: false, error: 'Email column not found' };
    }
    
    const existingData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][emailCol] === data.email) {
        rowIndex = i + 1;
        break;
      }
    }
    
    // Prepare row data
    const rowData = headers.map(header => {
      const key = header.charAt(0).toLowerCase() + header.slice(1).replace(/([A-Z])/g, '$1');
      return data[key] || data[header.toLowerCase()] || '';
    });
    
    if (rowIndex > 0) {
      // Update existing record
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // Add new record
      sheet.appendRow(rowData);
    }
    
    // Update dashboard stats if needed
    if (sheetName === CONFIG.SHEETS.PROFILES || sheetName === CONFIG.SHEETS.EXAMS) {
      updateDashboardStats();
    }
    
    return { success: true, message: 'Data saved successfully' };
    
  } catch (error) {
    Logger.log(`Save data error for ${sheetName}: ` + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Load data from specified sheet
 */
function loadData(sheetName, email = null) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, error: `Sheet ${sheetName} not found` };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: null };
    }
    
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    
    if (email && emailCol !== -1) {
      // Find specific user data
      for (let i = 1; i < data.length; i++) {
        if (data[i][emailCol] === email) {
          const userData = {};
          headers.forEach((header, index) => {
            const key = header.charAt(0).toLowerCase() + header.slice(1);
            userData[key] = data[i][index];
          });
          return { success: true, data: userData };
        }
      }
      return { success: true, data: null };
    } else {
      // Return all data
      const allData = [];
      for (let i = 1; i < data.length; i++) {
        const rowData = {};
        headers.forEach((header, index) => {
          const key = header.charAt(0).toLowerCase() + header.slice(1);
          rowData[key] = data[i][index];
        });
        allData.push(rowData);
      }
      return { success: true, data: allData };
    }
    
  } catch (error) {
    Logger.log(`Load data error for ${sheetName}: ` + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get dashboard statistics
 */
function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Get user statistics
    const userStats = getUserStats(ss);
    const examStats = getExamStats(ss);
    const donationStats = getDonationStats(ss);
    const demographicStats = getDemographicStats(ss);
    
    const stats = {
      users: userStats,
      exams: examStats,
      donations: donationStats,
      demographics: demographicStats,
      lastUpdated: new Date()
    };
    
    return { success: true, data: stats };
    
  } catch (error) {
    Logger.log('Dashboard stats error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get user statistics
 */
function getUserStats(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROFILES);
    if (!sheet) return { students: 0, alumni: 0, total: 0 };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { students: 0, alumni: 0, total: 0 };
    
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    
    let students = 0;
    let alumni = 0;
    
    // Check user types from Users sheet
    const usersSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.USERS);
    if (usersSheet) {
      const usersData = usersSheet.getDataRange().getValues();
      const usersHeaders = usersData[0];
      const userTypeCol = usersHeaders.indexOf('UserType');
      
      for (let i = 1; i < usersData.length; i++) {
        if (usersData[i][userTypeCol] === 'student') {
          students++;
        } else if (usersData[i][userTypeCol] === 'alumni') {
          alumni++;
        }
      }
    }
    
    return {
      students: students,
      alumni: alumni,
      total: students + alumni
    };
    
  } catch (error) {
    Logger.log('User stats error: ' + error.toString());
    return { students: 0, alumni: 0, total: 0 };
  }
}

/**
 * Get exam statistics
 */
function getExamStats(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.EXAMS);
    if (!sheet) return { passed: 0, failed: 0, passRate: 0, subjectStats: {} };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { passed: 0, failed: 0, passRate: 0, subjectStats: {} };
    
    const headers = data[0];
    const allPassedCol = headers.indexOf('AllPassed');
    
    let passed = 0;
    let total = data.length - 1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][allPassedCol] === true || data[i][allPassedCol] === 'TRUE') {
        passed++;
      }
    }
    
    const failed = total - passed;
    const passRate = total > 0 ? Math.round((passed / total) * 100) : 0;
    
    return {
      passed: passed,
      failed: failed,
      passRate: passRate,
      subjectStats: getSubjectStats(data, headers)
    };
    
  } catch (error) {
    Logger.log('Exam stats error: ' + error.toString());
    return { passed: 0, failed: 0, passRate: 0, subjectStats: {} };
  }
}

/**
 * Get subject-wise statistics
 */
function getSubjectStats(data, headers) {
  const subjects = [
    'การผดุงครรภ์',
    'การพยาบาลมารดาและทารก', 
    'การพยาบาลเด็กและวัยรุ่น',
    'การพยาบาลผู้ใหญ่',
    'การพยาบาลผู้สูงอายุ',
    'การพยาบาลสุขภาพจิตและจิตเวชศาสตร์',
    'การพยาบาลอนามัยชุมชนและการรักษาพยาบาลขั้นต้น',
    'กฎหมายว่าด้วยวิชาชีพการพยาบาลฯ'
  ];
  
  const subjectStats = {};
  
  subjects.forEach((subject, index) => {
    const subjectCol = headers.indexOf(`Subject${index + 1}`);
    if (subjectCol !== -1) {
      let passed = 0;
      let total = 0;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][subjectCol] && data[i][subjectCol] !== '') {
          total++;
          if (data[i][subjectCol] === 'pass') {
            passed++;
          }
        }
      }
      
      subjectStats[subject] = {
        passed: passed,
        total: total,
        passRate: total > 0 ? Math.round((passed / total) * 100) : 0
      };
    }
  });
  
  return subjectStats;
}

/**
 * Get donation statistics
 */
function getDonationStats(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DONATIONS);
    if (!sheet) return { amount: 0, total: 0, purposes: {} };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { amount: 0, total: 0, purposes: {} };
    
    const headers = data[0];
    const amountCol = headers.indexOf('Amount');
    const purposeCol = headers.indexOf('Purpose');
    
    let totalAmount = 0;
    let totalDonations = data.length - 1;
    const purposes = {};
    
    for (let i = 1; i < data.length; i++) {
      const amount = parseFloat(data[i][amountCol]) || 0;
      totalAmount += amount;
      
      const purpose = data[i][purposeCol] || 'อื่นๆ';
      purposes[purpose] = (purposes[purpose] || 0) + 1;
    }
    
    return {
      amount: totalAmount,
      total: totalDonations,
      purposes: purposes
    };
    
  } catch (error) {
    Logger.log('Donation stats error: ' + error.toString());
    return { amount: 0, total: 0, purposes: {} };
  }
}

/**
 * Get demographic statistics
 */
function getDemographicStats(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROFILES);
    if (!sheet) return { gender: {}, age: {}, nationality: {}, graduationClass: {} };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { gender: {}, age: {}, nationality: {}, graduationClass: {} };
    
    const headers = data[0];
    const genderCol = headers.indexOf('Gender');
    const birthDateCol = headers.indexOf('BirthDate');
    const nationalityCol = headers.indexOf('Nationality');
    const graduationClassCol = headers.indexOf('GraduationClass');
    
    const stats = {
      gender: {},
      age: {},
      nationality: {},
      graduationClass: {}
    };
    
    for (let i = 1; i < data.length; i++) {
      // Gender stats
      const gender = data[i][genderCol] || 'ไม่ระบุ';
      stats.gender[gender] = (stats.gender[gender] || 0) + 1;
      
      // Age stats
      if (data[i][birthDateCol]) {
        const birthDate = new Date(data[i][birthDateCol]);
        const age = new Date().getFullYear() - birthDate.getFullYear();
        let ageGroup = '41+';
        if (age <= 25) ageGroup = '20-25';
        else if (age <= 30) ageGroup = '26-30';
        else if (age <= 35) ageGroup = '31-35';
        else if (age <= 40) ageGroup = '36-40';
        
        stats.age[ageGroup] = (stats.age[ageGroup] || 0) + 1;
      }
      
      // Nationality stats
      const nationality = data[i][nationalityCol] || 'ไม่ระบุ';
      stats.nationality[nationality] = (stats.nationality[nationality] || 0) + 1;
      
      // Graduation class stats
      const graduationClass = data[i][graduationClassCol] || 'ไม่ระบุ';
      stats.graduationClass[graduationClass] = (stats.graduationClass[graduationClass] || 0) + 1;
    }
    
    return stats;
    
  } catch (error) {
    Logger.log('Demographic stats error: ' + error.toString());
    return { gender: {}, age: {}, nationality: {}, graduationClass: {} };
  }
}

/**
 * Update dashboard statistics cache
 */
function updateDashboardStats() {
  try {
    const stats = getDashboardStats();
    if (stats.success) {
      // You can cache these stats in the Dashboard_Stats sheet if needed
      Logger.log('Dashboard stats updated successfully');
    }
  } catch (error) {
    Logger.log('Update dashboard stats error: ' + error.toString());
  }
}

/**
 * Reset password (for admin use)
 */
function resetPassword(email, newPassword) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        sheet.getRange(i + 1, passwordCol + 1).setValue(newPassword);
        return { success: true, message: 'Password reset successfully' };
      }
    }
    
    return { success: false, error: 'User not found' };
    
  } catch (error) {
    Logger.log('Reset password error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get all users (for admin)
 */
function getAllUsers() {
  try {
    const result = loadData(CONFIG.SHEETS.USERS);
    return result;
  } catch (error) {
    Logger.log('Get all users error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete user (for admin)
 */
function deleteUser(email) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = [CONFIG.SHEETS.USERS, CONFIG.SHEETS.PROFILES, CONFIG.SHEETS.EXAMS, CONFIG.SHEETS.LICENSES, CONFIG.SHEETS.DONATIONS];
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const emailCol = headers.indexOf('Email');
        
        if (emailCol !== -1) {
          for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][emailCol] === email) {
              sheet.deleteRow(i + 1);
            }
          }
        }
      }
    });
    
    return { success: true, message: 'User deleted successfully' };
    
  } catch (error) {
    Logger.log('Delete user error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Test function to verify setup
 */
function testSetup() {
  try {
    Logger.log('Testing Google Apps Script setup...');
    
    // Test spreadsheet access
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('Spreadsheet access: OK');
    
    // Test sheet creation
    const result = initializeSpreadsheet();
    Logger.log('Sheet initialization: ' + JSON.stringify(result));
    
    // Test data operations
    const testData = {
      email: 'test@example.com',
      firstName: 'Test',
      lastName: 'User',
      userType: 'student',
      password: 'test123',
      registrationDate: new Date()
    };
    
    const saveResult = saveData(CONFIG.SHEETS.USERS, testData);
    Logger.log('Save test: ' + JSON.stringify(saveResult));
    
    const loadResult = loadData(CONFIG.SHEETS.USERS, 'test@example.com');
    Logger.log('Load test: ' + JSON.stringify(loadResult));
    
    Logger.log('All tests completed successfully!');
    return { success: true, message: 'Setup test completed' };
    
  } catch (error) {
    Logger.log('Test setup error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// รันฟังก์ชันนี้ครั้งแรก
function setup() {
  initializeSpreadsheet();
  testSetup();
}

