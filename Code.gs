/**
 * Code.gs
 * ไฟล์หลักของ SDIP - Smart Delivery Insight Platform
 *
 * ระบบตรวจสอบงานค้างสำหรับไปรษณีย์ไทย
 * - งานค้าง (Backlog): R / EMS / COD
 * - งานคืน (Return): R / EMS / COD
 *
 * @version 2.0.0
 * @author Claude AI
 * @date 2025-11-10
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
 * ฟังก์ชันทดสอบการ Login
 */
function testLogin() {
  try {
    // ทดสอบ login ด้วย username/password ตัวอย่าง
    // เปลี่ยนเป็น username/password จริงในชีต SDIPEmployee
    const result = authenticateUser('testuser', 'testpass');
    Logger.log(JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return null;
  }
}

/**
 * ฟังก์ชันทดสอบการดึงข้อมูล
 */
function testGetData() {
  try {
    Logger.log('=== ทดสอบดึงข้อมูล ===');

    // 1. ดึงข้อมูลพนักงาน
    const employees = getAllEmployees();
    Logger.log('พนักงานทั้งหมด: ' + employees.length + ' คน');

    // 2. ดึงข้อมูลงานค้าง
    const backlogR = getBacklogData('R');
    const backlogEMS = getBacklogData('EMS');
    const backlogCOD = getBacklogData('COD');
    Logger.log('งานค้าง R: ' + backlogR.length);
    Logger.log('งานค้าง EMS: ' + backlogEMS.length);
    Logger.log('งานค้าง COD: ' + backlogCOD.length);

    // 3. ดึงข้อมูลงานคืน
    const returnR = getReturnData('R');
    const returnEMS = getReturnData('EMS');
    const returnCOD = getReturnData('COD');
    Logger.log('งานคืน R: ' + returnR.length);
    Logger.log('งานคืน EMS: ' + returnEMS.length);
    Logger.log('งานคืน COD: ' + returnCOD.length);

    Logger.log('=== ทดสอบสำเร็จ ===');

    return {
      success: true,
      employees: employees.length,
      backlog: {
        r: backlogR.length,
        ems: backlogEMS.length,
        cod: backlogCOD.length
      },
      returned: {
        r: returnR.length,
        ems: returnEMS.length,
        cod: returnCOD.length
      }
    };

  } catch (error) {
    Logger.log('Error: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ฟังก์ชันทดสอบการค้นหา Barcode
 */
function testSearchBarcode() {
  try {
    // ทดสอบค้นหา barcode (เปลี่ยนเป็น barcode จริงในข้อมูล)
    const result = searchByBarcode('BC069609595TH');
    Logger.log(JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return null;
  }
}

/**
 * ฟังก์ชันแสดงข้อมูลสถิติรวม
 */
function showTotalStats() {
  try {
    const stats = getTotalStats();

    if (stats) {
      Logger.log('=== สถิติรวมทั้งหมด ===');
      Logger.log('พนักงาน: ' + stats.employees + ' คน');
      Logger.log('');
      Logger.log('งานค้าง:');
      Logger.log('  R: ' + stats.backlog.r);
      Logger.log('  EMS: ' + stats.backlog.ems);
      Logger.log('  COD: ' + stats.backlog.cod);
      Logger.log('  รวม: ' + stats.backlog.total);
      Logger.log('');
      Logger.log('งานคืน:');
      Logger.log('  R: ' + stats.returned.r);
      Logger.log('  EMS: ' + stats.returned.ems);
      Logger.log('  COD: ' + stats.returned.cod);
      Logger.log('  รวม: ' + stats.returned.total);
      Logger.log('======================');
    }

    return stats;

  } catch (error) {
    Logger.log('Error: ' + error.message);
    return null;
  }
}

/**
 * ฟังก์ชันสำหรับเมนู custom ใน Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 SDIP')
    .addItem('📊 แสดงสถิติรวม', 'showTotalStats')
    .addSeparator()
    .addItem('🧪 ทดสอบการเชื่อมต่อ', 'testConnection')
    .addItem('🧪 ทดสอบดึงข้อมูล', 'testGetData')
    .addSeparator()
    .addItem('ℹ️ เกี่ยวกับ SDIP', 'showAbout')
    .addToUi();
}

/**
 * ฟังก์ชันแสดงข้อมูลเกี่ยวกับ SDIP
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '📦 SDIP - Smart Delivery Insight Platform',
    'เวอร์ชัน: 2.0.0\n' +
    'สร้างโดย: Claude AI\n' +
    'วันที่: 10 พฤศจิกายน 2568\n\n' +
    'ระบบตรวจสอบงานค้างสำหรับไปรษณีย์ไทย\n\n' +
    'ฟีเจอร์:\n' +
    '• ตรวจสอบงานค้าง (R/EMS/COD)\n' +
    '• ตรวจสอบงานคืน (R/EMS/COD)\n' +
    '• รายงานสถิติพนักงาน\n' +
    '• ค้นหา Barcode\n' +
    '• ระบบ Login 3 ระดับ',
    ui.ButtonSet.OK
  );
}
