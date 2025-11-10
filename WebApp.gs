/**
 * WebApp.gs
 * Web App handlers สำหรับ Google Apps Script
 */

/**
 * ฟังก์ชันหลักสำหรับแสดงหน้า Web App
 * @param {Object} e - Event object
 * @return {HtmlOutput} - HTML output
 */
function doGet(e) {
  try {
    // ถ้ามี parameter page=dashboard ให้แสดงหน้า dashboard
    if (e.parameter.page === 'dashboard') {
      return HtmlService.createHtmlOutputFromFile('Dashboard')
        .setTitle('SDIP Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // ถ้าไม่มี parameter หรือไม่ใช่ dashboard ให้แสดงหน้า login
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('SDIP Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log('Error in doGet: ' + error.message);
    return HtmlService.createHtmlOutput('<h1>เกิดข้อผิดพลาด</h1><p>' + error.message + '</p>');
  }
}

/**
 * ฟังก์ชันสำหรับจัดการ POST requests
 * @param {Object} e - Event object
 * @return {ContentOutput} - JSON output
 */
function doPost(e) {
  try {
    const action = e.parameter.action;

    switch (action) {
      case 'login':
        return ContentService
          .createTextOutput(JSON.stringify(handleLogin(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getData':
        return ContentService
          .createTextOutput(JSON.stringify(handleGetData(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      default:
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            message: 'Invalid action'
          }))
          .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    Logger.log('Error in doPost: ' + error.message);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันสำหรับจัดการ Login request
 * @param {Object} params - Parameters
 * @return {Object} - Result
 */
function handleLogin(params) {
  const username = params.username;
  const password = params.password;

  return authenticateUser(username, password);
}

/**
 * ฟังก์ชันสำหรับจัดการ Get Data request
 * @param {Object} params - Parameters
 * @return {Object} - Result
 */
function handleGetData(params) {
  const username = params.username;
  const accessLevel = params.accessLevel;

  // ถ้าเป็น User ธรรมดา ให้ดึงเฉพาะข้อมูลของตัวเอง
  if (accessLevel === ACCESS_LEVELS.USER) {
    return getEmployeeData(username);
  }

  // ถ้าเป็น Admin หรือ PowerUser ให้ดึงข้อมูลทั้งหมด
  return getAllDashboardData();
}

/**
 * ฟังก์ชัน Server-side function สำหรับ google.script.run
 * เรียกใช้จาก client-side
 */

/**
 * ฟังก์ชันสำหรับ Login (เรียกจาก client-side)
 */
function serverLogin(username, password) {
  return authenticateUser(username, password);
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลทั้งหมด (เรียกจาก client-side)
 */
function serverGetAllData() {
  return getAllDashboardData();
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลพนักงานคนเดียว (เรียกจาก client-side)
 */
function serverGetEmployeeData(username) {
  return getEmployeeData(username);
}

/**
 * ฟังก์ชันสำหรับค้นหา Barcode (เรียกจาก client-side)
 */
function serverSearchBarcode(barcode) {
  return searchByBarcode(barcode);
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลผู้ใช้ทั้งหมด (Admin only)
 */
function serverGetAllUsers() {
  return getAllUsers();
}

/**
 * ฟังก์ชันสำหรับเพิ่มผู้ใช้ (Admin only)
 */
function serverAddUser(userData) {
  return addUser(userData);
}

/**
 * ฟังก์ชันสำหรับลบผู้ใช้ (Admin only)
 */
function serverDeleteUser(username) {
  return deleteUser(username);
}

/**
 * ฟังก์ชันสำหรับเปลี่ยน Password
 */
function serverChangePassword(username, oldPassword, newPassword) {
  return changePassword(username, oldPassword, newPassword);
}

/**
 * ฟังก์ชันสำหรับตรวจสอบสถานะของ Spreadsheet
 */
function serverCheckSheets() {
  return checkAllSheets();
}

/**
 * ฟังก์ชันสำหรับสร้างชีตที่ขาดหาย
 */
function serverCreateMissingSheets() {
  return createMissingSheets();
}
