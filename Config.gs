/**
 * Config.gs
 * การตั้งค่าระบบและค่าคงที่
 */

// Spreadsheet ID
const SPREADSHEET_ID = '1qlpxCNo7kefla-dZCdx9Jr70e-x-9rVNkpau6QmPeUk';

// ชื่อชีตต่างๆ
const SHEET_NAMES = {
  USERS: 'Users',
  EMPLOYEES: 'Employees',
  BACKLOG_R: 'Backlog_R',
  BACKLOG_EMS: 'Backlog_EMS',
  BACKLOG_COD: 'Backlog_COD',
  RETURN_R: 'Return_R',
  RETURN_EMS: 'Return_EMS',
  RETURN_COD: 'Return_COD'
};

// ระดับผู้ใช้งาน
const ACCESS_LEVELS = {
  USER: 'User',
  ADMIN: 'Admin',
  POWER_USER: 'PowerUser'
};

// คอลัมน์ในชีต Users (ตามลำดับ A, B, C, ...)
const USER_COLUMNS = {
  USERNAME: 0,
  PASSWORD: 1,
  NAME: 2,
  ACCESS_LEVEL: 3
};

// คอลัมน์ในชีต Employees (ตามลำดับ A, B, C, ...)
const EMPLOYEE_COLUMNS = {
  USERNAME: 0,
  NAME: 1,
  PAYMENT_SIDE: 2  // โซน (เช่น ซ้าย, ขวา, กลาง)
};

// คอลัมน์ในชีตงาน (Backlog/Return)
const WORK_COLUMNS = {
  BARCODE: 0,
  OPERATOR: 1,
  DATE: 2,
  STATUS: 3,
  PAYMENT_SIDE: 4
};

/**
 * ฟังก์ชันเพื่อดึง Spreadsheet object
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw new Error('ไม่สามารถเข้าถึง Google Sheets ได้: ' + error.message);
  }
}

/**
 * ฟังก์ชันเพื่อดึงชีตตามชื่อ
 */
function getSheet(sheetName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('ไม่พบชีต: ' + sheetName);
    }

    return sheet;
  } catch (error) {
    throw new Error('เกิดข้อผิดพลาดในการดึงชีต ' + sheetName + ': ' + error.message);
  }
}

/**
 * ฟังก์ชันเพื่อตรวจสอบว่าชีตทั้งหมดพร้อมใช้งานหรือไม่
 */
function checkAllSheets() {
  const ss = getSpreadsheet();
  const missingSheets = [];

  for (const key in SHEET_NAMES) {
    const sheetName = SHEET_NAMES[key];
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      missingSheets.push(sheetName);
    }
  }

  if (missingSheets.length > 0) {
    return {
      success: false,
      message: 'ไม่พบชีตต่อไปนี้: ' + missingSheets.join(', ')
    };
  }

  return {
    success: true,
    message: 'ชีตทั้งหมดพร้อมใช้งาน'
  };
}
