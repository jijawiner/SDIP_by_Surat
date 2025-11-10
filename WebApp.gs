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
    // Check for test mode (เพิ่ม ?test=true ใน URL เพื่อเข้าหน้าทดสอบ)
    if (e && e.parameter && e.parameter.test === 'true') {
      Logger.log('Serving Test.html for debugging');
      return HtmlService.createHtmlOutputFromFile('Test')
        .setTitle('SDIP Test Page')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Single Page Application - แสดง Index.html เสมอ (รวม Login + Dashboard)
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('SDIP - Smart Delivery Insight Platform')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log('Error in doGet: ' + error.message);
    return HtmlService.createHtmlOutput('<h1>เกิดข้อผิดพลาด</h1><p>' + error.message + '</p>');
  }
}

/**
 * Server-side functions เรียกจาก client-side ผ่าน google.script.run
 */

/**
 * ฟังก์ชันสำหรับ Login
 */
function serverLogin(username, password) {
  return authenticateUser(username, password);
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลทั้งหมด
 * หมายเหตุ: getAllDashboardData() ถูกกำหนดไว้แล้วใน DataService.gs
 * และสามารถเรียกใช้จาก client ได้โดยตรง (ไม่ต้องมี wrapper function)
 */
function serverGetAllData() {
  return getAllDashboardData();
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลพนักงานคนเดียว
 */
function serverGetEmployeeData(username) {
  return getEmployeeData(username);
}

/**
 * ฟังก์ชันสำหรับค้นหา Barcode
 */
function serverSearchBarcode(barcode) {
  return searchByBarcode(barcode);
}

/**
 * ฟังก์ชันสำหรับดึงสถิติพนักงาน
 */
function serverGetEmployeeStats(username) {
  return getEmployeeWorkStats(username);
}

/**
 * ฟังก์ชันสำหรับดึงสถิติรวม
 */
function serverGetTotalStats() {
  return getTotalStats();
}

/**
 * ฟังก์ชันสำหรับเปลี่ยน Password
 */
function serverChangePassword(username, oldPassword, newPassword) {
  return changePassword(username, oldPassword, newPassword);
}

/**
 * ฟังก์ชันสำหรับตรวจสอบสถานะชีต
 */
function serverCheckSheets() {
  return checkAllSheets();
}
