/**
 * Code.gs
 * ไฟล์หลักของ SDIP - Smart Delivery Insight Platform
 *
 * ระบบตรวจสอบงานค้างสำหรับไปรษณีย์ไทย (ฉบับย่อ)
 * - งานค้าง (Backlog): R / EMS / COD
 * - งานคืน (Return): R / EMS / COD
 *
 * @version 1.0.0
 * @author Claude AI
 */

/**
 * ฟังก์ชันทดสอบการเชื่อมต่อ Spreadsheet
 */
function testConnection() {
  try {
    const result = checkAllSheets();
    Logger.log(result.message);
    return result;
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ฟังก์ชันสำหรับสร้างชีตที่ขาดหายไป (ถ้ามี)
 */
function createMissingSheets() {
  try {
    const ss = getSpreadsheet();

    // สร้างชีต Users ถ้ายังไม่มี
    if (!ss.getSheetByName(SHEET_NAMES.USERS)) {
      const usersSheet = ss.insertSheet(SHEET_NAMES.USERS);
      usersSheet.getRange('A1:D1').setValues([['Username', 'Password', 'Name', 'AccessLevel']]);
      usersSheet.getRange('A2:D2').setValues([['admin', 'admin123', 'ผู้ดูแลระบบ', 'Admin']]);
      Logger.log('สร้างชีต Users เรียบร้อย');
    }

    // สร้างชีต Employees ถ้ายังไม่มี
    if (!ss.getSheetByName(SHEET_NAMES.EMPLOYEES)) {
      const empSheet = ss.insertSheet(SHEET_NAMES.EMPLOYEES);
      empSheet.getRange('A1:C1').setValues([['Username', 'Name', 'PaymentSide']]);
      Logger.log('สร้างชีต Employees เรียบร้อย');
    }

    // สร้างชีตงานค้าง
    const backlogSheets = [SHEET_NAMES.BACKLOG_R, SHEET_NAMES.BACKLOG_EMS, SHEET_NAMES.BACKLOG_COD];
    backlogSheets.forEach(function(sheetName) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        sheet.getRange('A1:E1').setValues([['Barcode', 'Operator', 'Date', 'Status', 'PaymentSide']]);
        Logger.log('สร้างชีต ' + sheetName + ' เรียบร้อย');
      }
    });

    // สร้างชีตงานคืน
    const returnSheets = [SHEET_NAMES.RETURN_R, SHEET_NAMES.RETURN_EMS, SHEET_NAMES.RETURN_COD];
    returnSheets.forEach(function(sheetName) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        sheet.getRange('A1:E1').setValues([['Barcode', 'Operator', 'Date', 'Status', 'PaymentSide']]);
        Logger.log('สร้างชีต ' + sheetName + ' เรียบร้อย');
      }
    });

    return {
      success: true,
      message: 'ตรวจสอบและสร้างชีตที่จำเป็นเรียบร้อยแล้ว'
    };

  } catch (error) {
    Logger.log('Error in createMissingSheets: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ฟังก์ชันสำหรับทดสอบดึงข้อมูลพนักงาน
 */
function testGetEmployees() {
  try {
    const employees = getEmployees();
    Logger.log('Found ' + employees.length + ' employees');
    return employees;
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return null;
  }
}

/**
 * ฟังก์ชันสำหรับทดสอบ Login
 */
function testLogin() {
  try {
    const result = authenticateUser('admin', 'admin123');
    Logger.log(JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return null;
  }
}
