/**
 * Config.gs
 * ไฟล์การตั้งค่าระบบ SDIP - Smart Delivery Insight Platform
 *
 * เก็บ Spreadsheet ID, ชื่อชีต และ Column mapping
 */

// ===================================
// Spreadsheet Configuration
// ===================================

const SPREADSHEET_ID = '1qlpxCNo7kefla-dZCdx9Jr70e-x-9rVNkpau6QmPeUk';

// ===================================
// Sheet Names
// ===================================

const SHEET_NAMES = {
  // ชีตพนักงาน (รวม Login)
  EMPLOYEES: 'SDIPEmployee',

  // ชีตงานค้าง (Backlog)
  BACKLOG_R: 'SDIPBacklogR',
  BACKLOG_EMS: 'SDIPBacklogEMS',
  BACKLOG_COD: 'SDIPBacklogCOD',

  // ชีตงานคืน (Return)
  RETURN_R: 'SDIPReturnedR',
  RETURN_EMS: 'SDIPReturnedEMS',
  RETURN_COD: 'SDIPReturnedCOD',

  // ชีตเสริม
  RETURN_NOTE: 'SDIPReturnNote',
  RETURNED_WORK: 'SDIPReturnedWork',
  DATA_ANALYSIS: 'SDIPDataAnalysis'
};

// ===================================
// Column Mappings
// ===================================

/**
 * SDIPEmployee - ข้อมูลพนักงาน (15 คอลัมน์)
 */
const EMPLOYEE_COLUMNS = {
  ORDER: 0,           // A - ลำดับ
  USER_STATUS: 1,     // B - UserStatus
  USERNAME: 2,        // C - Username
  PASSWORD: 3,        // D - Password
  ACCESS_LEVEL: 4,    // E - Useraccesslevel
  NAME: 5,            // F - Name
  POSITION: 6,        // G - Position
  PAYMENT_SIDE: 7,    // H - Paymentside (โซน)
  DISTRICT: 8,        // I - district
  VILLAGE_NO: 9,      // J - VillageNo.
  WORK_PHONE: 10,     // K - WorkPhone
  MOBILE_PHONE: 11,   // L - MobilePhone
  LINE_ID: 12,        // M - LINEID
  GMAIL: 13           // N - GMAIL.COM
};

/**
 * งานค้าง/งานคืน (15 คอลัมน์)
 * ใช้กับ: SDIPBacklogR, SDIPReturnedR, SDIPBacklogEMS,
 *         SDIPReturnedEMS, SDIPBacklogCOD, SDIPReturnedCOD
 */
const WORK_COLUMNS = {
  ORDER: 0,                    // A - ลำดับ
  BARCODE: 1,                  // B - หมายเลขสิ่งของ
  RECIPIENT_INFO: 2,           // C - ชือ - ที่อยู่ผู้รับ
  OPERATOR: 3,                 // D - ผู้ดำเนินการ
  SCAN_DATETIME: 4,            // E - วันที่ เวลา (สแกนผลการนำจ่ายล่าสุด)
  REASON: 5,                   // F - สาเหตุ
  COD_AMOUNT: 6,               // G - COD
  IS_LAZADA: 7,                // H - Lazada?
  DAYS_FROM_DEPOSIT: 8,        // I - จำนวนวันที่ถือครอง (นับจากวันที่รับฝาก)
  DAYS_FROM_DESTINATION: 9,    // J - จำนวนวันที่ถือครอง (นับจากวันที่ปลายทาง)
  ATTEMPT_COUNT: 10,           // K - จำนวนครั้งที่พยายามนำจ่าย
  DELIVERED: 11,               // L - ส่ง มอบ
  APPOINTMENT: 12,             // M - นัดรับ/แจ้ง
  CONTACT_RECIPIENT: 13,       // N - ติดต่อผู้รับ
  PREPARE_RETURN: 14           // O - เตรียมคืน
};

// ===================================
// Access Levels
// ===================================

const ACCESS_LEVELS = {
  USER: 'User',           // พนักงานทั่วไป - ดูข้อมูลตัวเองเท่านั้น
  ADMIN: 'Admin',         // ผู้ดูแลระบบ - ดูข้อมูลทั้งหมด + จัดการผู้ใช้
  POWER_USER: 'PowerUser' // ผู้ควบคุม - ดูข้อมูลทั้งหมด
};

// ===================================
// Helper Functions
// ===================================

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
 * ฟังก์ชันตรวจสอบว่าชีตทั้งหมดพร้อมใช้งานหรือไม่
 */
function checkAllSheets() {
  const ss = getSpreadsheet();
  const missingSheets = [];
  const foundSheets = [];

  // ตรวจสอบเฉพาะชีตที่จำเป็น
  const requiredSheets = [
    SHEET_NAMES.EMPLOYEES,
    SHEET_NAMES.BACKLOG_R,
    SHEET_NAMES.BACKLOG_EMS,
    SHEET_NAMES.BACKLOG_COD,
    SHEET_NAMES.RETURN_R,
    SHEET_NAMES.RETURN_EMS,
    SHEET_NAMES.RETURN_COD
  ];

  requiredSheets.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      missingSheets.push(sheetName);
    } else {
      foundSheets.push(sheetName);
    }
  });

  if (missingSheets.length > 0) {
    return {
      success: false,
      message: 'ไม่พบชีตต่อไปนี้: ' + missingSheets.join(', ')
    };
  }

  return {
    success: true,
    message: 'ชีตทั้งหมดพร้อมใช้งาน (' + foundSheets.length + ' ชีต)'
  };
}

/**
 * ฟังก์ชันแปลง Column index เป็น Column letter (A, B, C, ...)
 */
function columnToLetter(column) {
  let temp;
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

/**
 * ฟังก์ชันแปลง Column letter เป็น index
 */
function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;

  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }

  return column;
}
