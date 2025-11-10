/**
 * Setup.gs
 * สคริปต์สำหรับติดตั้งและสร้างชีตเริ่มต้น
 */

/**
 * ฟังก์ชันหลักสำหรับติดตั้งระบบครั้งแรก
 * รันฟังก์ชันนี้หลังจาก copy โค้ดมาใหม่
 */
function setupSDIP() {
  try {
    Logger.log('🚀 เริ่มติดตั้ง SDIP...');

    // 1. ตรวจสอบ Spreadsheet
    const ss = getSpreadsheet();
    Logger.log('✅ เชื่อมต่อ Spreadsheet สำเร็จ: ' + ss.getName());

    // 2. สร้างชีตที่จำเป็น
    Logger.log('📋 กำลังสร้างชีตที่จำเป็น...');
    const result = createAllSheets();

    if (result.success) {
      Logger.log('✅ สร้างชีตสำเร็จ');
      Logger.log(result.message);

      // 3. เพิ่มข้อมูลตัวอย่าง
      Logger.log('📝 กำลังเพิ่มข้อมูลตัวอย่าง...');
      addSampleData();

      Logger.log('🎉 ติดตั้ง SDIP สำเร็จ!');
      Logger.log('ℹ️ สามารถ Login ด้วย:');
      Logger.log('   Username: admin');
      Logger.log('   Password: admin123');

      // แสดงข้อความสำเร็จ
      SpreadsheetApp.getUi().alert(
        '✅ ติดตั้งสำเร็จ!\n\n' +
        'สร้างชีตทั้งหมดเรียบร้อยแล้ว\n\n' +
        'ข้อมูล Login:\n' +
        'Username: admin\n' +
        'Password: admin123\n\n' +
        'คุณสามารถ Deploy Web App ได้แล้ว'
      );

      return {
        success: true,
        message: 'ติดตั้ง SDIP สำเร็จ'
      };
    } else {
      throw new Error(result.message);
    }

  } catch (error) {
    Logger.log('❌ เกิดข้อผิดพลาด: ' + error.message);
    SpreadsheetApp.getUi().alert('❌ เกิดข้อผิดพลาด\n\n' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * สร้างชีตทั้งหมดที่จำเป็น
 */
function createAllSheets() {
  try {
    const ss = getSpreadsheet();
    const sheetsCreated = [];

    // 1. สร้างชีต Users
    if (!ss.getSheetByName(SHEET_NAMES.USERS)) {
      const usersSheet = ss.insertSheet(SHEET_NAMES.USERS);

      // Set headers
      usersSheet.getRange('A1:D1').setValues([['Username', 'Password', 'Name', 'AccessLevel']]);

      // Format headers
      usersSheet.getRange('A1:D1')
        .setBackground('#667eea')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

      // Set column widths
      usersSheet.setColumnWidth(1, 120); // Username
      usersSheet.setColumnWidth(2, 120); // Password
      usersSheet.setColumnWidth(3, 200); // Name
      usersSheet.setColumnWidth(4, 100); // AccessLevel

      // Add default admin user
      usersSheet.getRange('A2:D2').setValues([['admin', 'admin123', 'ผู้ดูแลระบบ', 'Admin']]);

      // Freeze header row
      usersSheet.setFrozenRows(1);

      sheetsCreated.push('Users');
      Logger.log('✅ สร้างชีต Users');
    }

    // 2. สร้างชีต Employees
    if (!ss.getSheetByName(SHEET_NAMES.EMPLOYEES)) {
      const empSheet = ss.insertSheet(SHEET_NAMES.EMPLOYEES);

      // Set headers
      empSheet.getRange('A1:C1').setValues([['Username', 'Name', 'PaymentSide']]);

      // Format headers
      empSheet.getRange('A1:C1')
        .setBackground('#9c27b0')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

      // Set column widths
      empSheet.setColumnWidth(1, 120); // Username
      empSheet.setColumnWidth(2, 200); // Name
      empSheet.setColumnWidth(3, 100); // PaymentSide

      // Freeze header row
      empSheet.setFrozenRows(1);

      sheetsCreated.push('Employees');
      Logger.log('✅ สร้างชีต Employees');
    }

    // 3. สร้างชีตงานค้าง (Backlog)
    const backlogSheets = [
      { name: SHEET_NAMES.BACKLOG_R, label: 'R', color: '#d32f2f' },
      { name: SHEET_NAMES.BACKLOG_EMS, label: 'EMS', color: '#1976d2' },
      { name: SHEET_NAMES.BACKLOG_COD, label: 'COD', color: '#f57c00' }
    ];

    backlogSheets.forEach(function(sheetInfo) {
      if (!ss.getSheetByName(sheetInfo.name)) {
        const sheet = ss.insertSheet(sheetInfo.name);

        // Set headers
        sheet.getRange('A1:E1').setValues([['Barcode', 'Operator', 'Date', 'Status', 'PaymentSide']]);

        // Format headers
        sheet.getRange('A1:E1')
          .setBackground(sheetInfo.color)
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');

        // Set column widths
        sheet.setColumnWidth(1, 150); // Barcode
        sheet.setColumnWidth(2, 100); // Operator
        sheet.setColumnWidth(3, 120); // Date
        sheet.setColumnWidth(4, 100); // Status
        sheet.setColumnWidth(5, 100); // PaymentSide

        // Freeze header row
        sheet.setFrozenRows(1);

        sheetsCreated.push(sheetInfo.name);
        Logger.log('✅ สร้างชีต ' + sheetInfo.name);
      }
    });

    // 4. สร้างชีตงานคืน (Return)
    const returnSheets = [
      { name: SHEET_NAMES.RETURN_R, label: 'R', color: '#c62828' },
      { name: SHEET_NAMES.RETURN_EMS, label: 'EMS', color: '#1565c0' },
      { name: SHEET_NAMES.RETURN_COD, label: 'COD', color: '#e65100' }
    ];

    returnSheets.forEach(function(sheetInfo) {
      if (!ss.getSheetByName(sheetInfo.name)) {
        const sheet = ss.insertSheet(sheetInfo.name);

        // Set headers
        sheet.getRange('A1:E1').setValues([['Barcode', 'Operator', 'Date', 'Status', 'PaymentSide']]);

        // Format headers
        sheet.getRange('A1:E1')
          .setBackground(sheetInfo.color)
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');

        // Set column widths
        sheet.setColumnWidth(1, 150); // Barcode
        sheet.setColumnWidth(2, 100); // Operator
        sheet.setColumnWidth(3, 120); // Date
        sheet.setColumnWidth(4, 100); // Status
        sheet.setColumnWidth(5, 100); // PaymentSide

        // Freeze header row
        sheet.setFrozenRows(1);

        sheetsCreated.push(sheetInfo.name);
        Logger.log('✅ สร้างชีต ' + sheetInfo.name);
      }
    });

    if (sheetsCreated.length > 0) {
      return {
        success: true,
        message: 'สร้างชีตสำเร็จ: ' + sheetsCreated.join(', ')
      };
    } else {
      return {
        success: true,
        message: 'ชีตทั้งหมดมีอยู่แล้ว ไม่ต้องสร้างใหม่'
      };
    }

  } catch (error) {
    Logger.log('❌ Error in createAllSheets: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการสร้างชีต: ' + error.message
    };
  }
}

/**
 * เพิ่มข้อมูลตัวอย่างเพื่อทดสอบ
 */
function addSampleData() {
  try {
    const ss = getSpreadsheet();

    // เพิ่มพนักงานตัวอย่าง
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    if (empSheet && empSheet.getLastRow() === 1) {
      const sampleEmployees = [
        ['emp001', 'สมชาย ใจดี', 'ซ้าย'],
        ['emp002', 'สมหญิง รักงาน', 'ขวา'],
        ['emp003', 'สมศักดิ์ ขยัน', 'กลาง']
      ];

      empSheet.getRange(2, 1, sampleEmployees.length, 3).setValues(sampleEmployees);
      Logger.log('✅ เพิ่มข้อมูลพนักงานตัวอย่าง');
    }

    // เพิ่มงานค้างตัวอย่าง
    const backlogRSheet = ss.getSheetByName(SHEET_NAMES.BACKLOG_R);
    if (backlogRSheet && backlogRSheet.getLastRow() === 1) {
      const sampleBacklog = [
        ['RR123456789TH', 'emp001', new Date(), 'Pending', 'ซ้าย'],
        ['RR987654321TH', 'emp002', new Date(), 'Pending', 'ขวา']
      ];

      backlogRSheet.getRange(2, 1, sampleBacklog.length, 5).setValues(sampleBacklog);
      Logger.log('✅ เพิ่มข้อมูลงานค้างตัวอย่าง');
    }

    Logger.log('✅ เพิ่มข้อมูลตัวอย่างสำเร็จ');

  } catch (error) {
    Logger.log('⚠️ Warning in addSampleData: ' + error.message);
    // ไม่ throw error เพราะเป็นแค่ข้อมูลตัวอย่าง
  }
}

/**
 * ฟังก์ชันสำหรับลบชีตทั้งหมด (ใช้ตอนต้องการเริ่มใหม่)
 * ⚠️ ระวัง! ฟังก์ชันนี้จะลบชีตทั้งหมด
 */
function resetAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ คำเตือน',
    'คุณต้องการลบชีตทั้งหมดและเริ่มใหม่หรือไม่?\n\nข้อมูลทั้งหมดจะถูกลบ!',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    Logger.log('ยกเลิกการลบชีต');
    return;
  }

  try {
    const ss = getSpreadsheet();

    // ลบชีตทั้งหมดที่เกี่ยวข้อง
    for (const key in SHEET_NAMES) {
      const sheetName = SHEET_NAMES[key];
      const sheet = ss.getSheetByName(sheetName);

      if (sheet) {
        ss.deleteSheet(sheet);
        Logger.log('🗑️ ลบชีต: ' + sheetName);
      }
    }

    Logger.log('✅ ลบชีตทั้งหมดเรียบร้อย');
    ui.alert('✅ ลบชีตสำเร็จ!\n\nกรุณารันฟังก์ชัน setupSDIP() เพื่อสร้างชีตใหม่');

  } catch (error) {
    Logger.log('❌ Error in resetAllSheets: ' + error.message);
    ui.alert('❌ เกิดข้อผิดพลาด\n\n' + error.message);
  }
}

/**
 * ฟังก์ชันตรวจสอบสถานะชีตทั้งหมด
 */
function checkSheetsStatus() {
  try {
    const ss = getSpreadsheet();
    let status = '📊 สถานะชีตใน SDIP\n\n';

    for (const key in SHEET_NAMES) {
      const sheetName = SHEET_NAMES[key];
      const sheet = ss.getSheetByName(sheetName);

      if (sheet) {
        const rowCount = sheet.getLastRow();
        status += '✅ ' + sheetName + ' (มี ' + rowCount + ' แถว)\n';
      } else {
        status += '❌ ' + sheetName + ' (ไม่พบชีต)\n';
      }
    }

    Logger.log(status);
    SpreadsheetApp.getUi().alert(status);

    return status;

  } catch (error) {
    Logger.log('❌ Error in checkSheetsStatus: ' + error.message);
    return 'เกิดข้อผิดพลาด: ' + error.message;
  }
}
