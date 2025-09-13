// Google Apps Script Code (Code.gs)
// This file should be created in Google Apps Script (script.google.com)

// Configuration - Update these with your actual Google Sheets IDs
const CONFIG = {
  SPREADSHEET_ID: '1w0OQk2NMBZp64CQOhhGuHXHYk8-87cBztWSGxsNT7iE', // Replace with your Google Sheets ID
  SHEETS: {
    USERS: 'Users',
    PROFILES: 'Profiles', 
    EXAMS: 'Exams',
    LICENSES: 'Licenses',
    DONATIONS: 'Donations'
  }
};

/**
 * Web App Entry Point - Required for Google Apps Script Web App
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Alumni & Student Tracking System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('doGet error: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

function doPost(e) {
  try {
    const action = e.parameter.action;
    const data = JSON.parse(e.parameter.data || '{}');
    let response;

    switch (action) {
      case 'authenticate':
        response = authenticate(data.email, data.password);
        break;
      case 'register':
        response = register(data);
        break;
      case 'saveData':
        response = saveData(data.sheetName, data.payload);
        break;
      case 'loadData':
        response = loadData(data.sheetName, data.key);
        break;
      case 'getUsersForAdmin':
        response = getUsersForAdmin();
        break;
      case 'getDashboardStats':
         response = getDashboardStats();
         break;
      default:
        response = { success: false, error: 'Unknown action' };
    }
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('doPost error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Initialize the spreadsheet with required sheets and headers
 */
function initializeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // [อัปเดต] เพิ่ม 'UserID' เป็นคอลัมน์แรกในทุกชีตที่เกี่ยวข้อง
    createSheetIfNotExists(ss, CONFIG.SHEETS.USERS, ['UserID', 'Email', 'Password', 'Role', 'UserType', 'FirstName', 'LastName', 'RegistrationDate', 'LastLogin', 'IsActive']);
    createSheetIfNotExists(ss, CONFIG.SHEETS.PROFILES, ['UserID', 'Email', 'Title', 'FirstNameTh', 'LastNameTh', 'FirstNameEn', 'LastNameEn', 'StudentId', 'GraduationClass', 'Advisor', 'BirthDate', 'Gender', 'BirthCountry', 'Nationality', 'Ethnicity', 'Phone', 'GPAX', 'Address', 'EmergencyName', 'EmergencyRelation', 'EmergencyPhone', 'Awards', 'CareerPlan', 'StudyPlan', 'InternationalPlan', 'WillTakeThaiLicense', 'LastUpdated']);
    createSheetIfNotExists(ss, CONFIG.SHEETS.EXAMS, ['UserID', 'Email', 'ExamRound', 'ExamSession', 'ExamYear', 'Results', 'PassedSubjects', 'AllPassed', 'ExamDate']);
    createSheetIfNotExists(ss, CONFIG.SHEETS.LICENSES, ['UserID', 'Email', 'MemberNumber', 'LicenseNumber', 'IssueDate', 'PermitDate', 'ExpiryDate', 'LastUpdated']);
    createSheetIfNotExists(ss, CONFIG.SHEETS.DONATIONS, ['UserID', 'Email', 'Amount', 'Purpose', 'OtherPurpose', 'Message', 'TaxName', 'TaxId', 'TaxAddress', 'Status', 'DonationDate', 'ProcessedDate']);

    Logger.log('Spreadsheet initialized successfully');
    return { success: true, message: 'Spreadsheet initialized' };
  } catch (error) {
    Logger.log('Error initializing spreadsheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * User Authentication
 */
function authenticate(email, password) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) return { success: false, error: 'Users sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password');
    const userIDCol = headers.indexOf('UserID'); // [อัปเดต]
    const roleCol = headers.indexOf('Role');
    const nameCol = headers.indexOf('FirstName');
    const lastNameCol = headers.indexOf('LastName');

    for (let i = 0; i < data.length; i++) {
      if (data[i][emailCol] === email && data[i][passwordCol] === password) {
        sheet.getRange(i + 2, headers.indexOf('LastLogin') + 1).setValue(new Date());
        
        // [อัปเดต] ส่ง UserID กลับไปด้วย
        return {
          success: true,
          user: {
            userID: data[i][userIDCol],
            email: data[i][emailCol],
            role: data[i][roleCol],
            name: `${data[i][nameCol]} ${data[i][lastNameCol]}`
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
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) return { success: false, error: 'Users sheet not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const emailCol = headers.indexOf('Email');

    if (data.some(row => row[emailCol] === userData.email)) {
      return { success: false, error: 'Email already exists' };
    }
    
    const newUserId = 'U' + new Date().getTime(); 
    const newRow = [
      newUserId, userData.email, userData.password, userData.userType, userData.userType,
      userData.firstName, userData.lastName, new Date(), '', true
    ];
    sheet.appendRow(newRow);
    
    // [อัปเดต] สร้าง Profile ว่างๆ ให้ทันทีเพื่อสร้าง Foreign Key
    const profilesSheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.PROFILES);
    const profileHeaders = profilesSheet.getRange(1, 1, 1, profilesSheet.getLastColumn()).getValues()[0];
    const newProfileRow = profileHeaders.map(header => {
        if (header === 'UserID') return newUserId;
        if (header === 'Email') return userData.email;
        if (header === 'FirstNameTh') return userData.firstName;
        if (header === 'LastNameTh') return userData.lastName;
        return '';
    });
    profilesSheet.appendRow(newProfileRow);

    return { success: true, message: 'User registered successfully' };
  } catch (error) {
    Logger.log('Registration error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save data to specified sheet (now uses UserID) - [แก้ไขบั๊ก]
 */
function saveData(sheetName, data) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) return { success: false, error: `Sheet ${sheetName} not found` };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const userIDCol = headers.indexOf('UserID');
    if (userIDCol === -1) return { success: false, error: 'UserID column not found in sheet: ' + sheetName };

    const existingData = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // Start loop from 1 to skip header row in the array
    for (let i = 1; i < existingData.length; i++) {
      // Robust comparison to avoid type issues (e.g., number vs string)
      if (String(existingData[i][userIDCol]) === String(data.userID)) {
        rowIndex = i + 1; // Sheet rows are 1-indexed
        break;
      }
    }
    
    // This mapping logic correctly converts Sheet Headers (e.g., FirstNameTh)
    // to the camelCase keys used by the frontend data object (e.g., firstNameTh)
    const rowData = headers.map(header => {
      const key = header.charAt(0).toLowerCase() + header.slice(1);
      // Handle the specific case of UserID vs userID
      if (header === 'UserID') {
          return data.userID || '';
      }
      return data[key] !== undefined ? data[key] : '';
    });

    if (rowIndex > 0) {
      // Update existing record
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // This part should ideally not be reached for updates, but acts as a fallback.
      sheet.appendRow(rowData);
    }
    
    return { success: true, message: 'Data saved successfully' };
  } catch (error) {
    Logger.log(`Save data error for ${sheetName}: ` + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Load data from specified sheet (now uses UserID)
 */
function loadData(sheetName, userID) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) return { success: false, data: null, error: `Sheet ${sheetName} not found` };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: null };
    
    const headers = data.shift();
    const userIDCol = headers.indexOf('UserID');
    
    if (userID && userIDCol !== -1) {
      for (let i = 0; i < data.length; i++) {
        if (data[i][userIDCol] === userID) {
          const record = {};
          headers.forEach((header, index) => {
             const key = header.charAt(0).toLowerCase() + header.slice(1);
             record[key] = data[i][index];
          });
          return { success: true, data: record };
        }
      }
      return { success: true, data: null };
    }
    return { success: false, data: null, error: 'UserID not provided or UserID column not found.'};
  } catch (error) {
    Logger.log(`Load data error for ${sheetName}: ` + error.toString());
    return { success: false, data: null, error: error.toString() };
  }
}


function setup() {
  initializeSpreadsheet();
}

function createSheetIfNotExists(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getUsersForAdmin() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const usersHeaders = usersData.shift();
    
    const profilesSheet = ss.getSheetByName(CONFIG.SHEETS.PROFILES);
    const profilesData = profilesSheet.getDataRange().getValues();
    const profilesHeaders = profilesData.shift();
    
    const profilesMap = profilesData.reduce((acc, row) => {
        const profile = {};
        profilesHeaders.forEach((header, i) => profile[header] = row[i]);
        if(profile.UserID) acc[profile.UserID] = profile;
        return acc;
    }, {});

    const combinedUsers = usersData.map(userRow => {
      const user = {};
      usersHeaders.forEach((header, i) => user[header] = userRow[i]);
      const userProfile = profilesMap[user.UserID] || {};
      
      return {
        userID: user.UserID,
        email: user.Email,
        userType: user.UserType,
        type: user.UserType,
        titleTh: userProfile.Title || '',
        firstNameTh: user.FirstName || '',
        lastNameTh: user.LastName || '',
        studentId: userProfile.StudentId || '',
        graduationClass: userProfile.GraduationClass || '',
        phone: userProfile.Phone || ''
      };
    });
    return { success: true, data: combinedUsers };
  } catch (error) {
    Logger.log('getUsersForAdmin error: ' + error.toString());
    return { success: false, error: error.toString(), data: [] };
  }
}

// =============================================================
// ==========[ Dashboard Functions Re-integrated ]==========
// =============================================================

function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
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

function getUserStats(spreadsheet) {
  const usersSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.USERS);
  if (!usersSheet) return { students: 0, alumni: 0, total: 0 };
  
  const usersData = usersSheet.getDataRange().getValues();
  usersData.shift(); // remove headers

  let students = 0;
  let alumni = 0;
  
  const userTypeCol = 3; // Index of 'UserType' column
  
  usersData.forEach(row => {
    if (row[userTypeCol] === 'student') {
      students++;
    } else if (row[userTypeCol] === 'alumni') {
      alumni++;
    }
  });
  
  return {
    students: students,
    alumni: alumni,
    total: students + alumni
  };
}

function getExamStats(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.EXAMS);
  if (!sheet) return { passed: 0, failed: 0, passRate: 0, subjectStats: {} };
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { passed: 0, failed: 0, passRate: 0, subjectStats: {} };
  
  const headers = data.shift();
  const allPassedCol = headers.indexOf('AllPassed');
  
  let passedCount = 0;
  
  data.forEach(row => {
    if (row[allPassedCol] === true || String(row[allPassedCol]).toUpperCase() === 'TRUE') {
      passedCount++;
    }
  });
  
  const totalExams = data.length;
  const failedCount = totalExams - passedCount;
  const passRate = totalExams > 0 ? Math.round((passedCount / totalExams) * 100) : 0;
  
  return {
    passed: passedCount,
    failed: failedCount,
    passRate: passRate
  };
}


function getDonationStats(spreadsheet) {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DONATIONS);
    if (!sheet) return { amount: 0, total: 0 };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { amount: 0, total: 0 };
    
    const headers = data.shift();
    const amountCol = headers.indexOf('Amount');
    
    let totalAmount = 0;
    
    data.forEach(row => {
        const amount = parseFloat(row[amountCol]) || 0;
        totalAmount += amount;
    });

    return {
        amount: totalAmount,
        total: data.length
    };
}


function getDemographicStats(spreadsheet) {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROFILES);
    if (!sheet) return { gender: {}, graduationClass: {} };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { gender: {}, graduationClass: {} };
    
    const headers = data.shift();
    const genderCol = headers.indexOf('Gender');
    const classCol = headers.indexOf('GraduationClass');

    const stats = {
        gender: {},
        graduationClass: {}
    };

    data.forEach(row => {
        const gender = row[genderCol] || 'ไม่ระบุ';
        stats.gender[gender] = (stats.gender[gender] || 0) + 1;
        
        const gradClass = row[classCol] || 'ไม่ระบุ';
        stats.graduationClass[gradClass] = (stats.graduationClass[gradClass] || 0) + 1;
    });

    return stats;
}

