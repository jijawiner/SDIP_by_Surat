/**
 * DataService.gs
 * ระบบจัดการข้อมูลจาก Google Sheets
 *
 * ดึงข้อมูลจาก 6 ชีตหลัก:
 * - SDIPBacklogR, SDIPBacklogEMS, SDIPBacklogCOD (งานค้าง)
 * - SDIPReturnedR, SDIPReturnedEMS, SDIPReturnedCOD (งานคืน)
 */

/**
 * ฟังก์ชันอ่านข้อมูลจากชีตงานค้าง/งานคืน
 * @param {string} sheetName - ชื่อชีต
 * @return {Array} - รายการงาน
 */
function getWorkData(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();

    const workItems = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // ข้าม row ที่ไม่มี barcode
      if (!row[WORK_COLUMNS.BARCODE]) continue;

      workItems.push({
        order: row[WORK_COLUMNS.ORDER] || i,
        barcode: row[WORK_COLUMNS.BARCODE].toString().trim(),
        recipientInfo: row[WORK_COLUMNS.RECIPIENT_INFO] || '',
        operator: row[WORK_COLUMNS.OPERATOR] || '',
        scanDatetime: row[WORK_COLUMNS.SCAN_DATETIME] || '',
        reason: row[WORK_COLUMNS.REASON] || '',
        codAmount: row[WORK_COLUMNS.COD_AMOUNT] || 0,
        isLazada: row[WORK_COLUMNS.IS_LAZADA] || '',
        daysFromDeposit: row[WORK_COLUMNS.DAYS_FROM_DEPOSIT] || 0,
        daysFromDestination: row[WORK_COLUMNS.DAYS_FROM_DESTINATION] || 0,
        attemptCount: row[WORK_COLUMNS.ATTEMPT_COUNT] || 0,
        delivered: row[WORK_COLUMNS.DELIVERED] || '',
        appointment: row[WORK_COLUMNS.APPOINTMENT] || '',
        contactRecipient: row[WORK_COLUMNS.CONTACT_RECIPIENT] || '',
        prepareReturn: row[WORK_COLUMNS.PREPARE_RETURN] || ''
      });
    }

    return workItems;

  } catch (error) {
    Logger.log('Error in getWorkData (' + sheetName + '): ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูล ' + sheetName + ': ' + error.message);
  }
}

/**
 * ฟังก์ชันดึงข้อมูลงานค้าง (Backlog)
 * @param {string} type - ประเภทงาน (R, EMS, COD)
 * @return {Array} - รายการงานค้าง
 */
function getBacklogData(type) {
  let sheetName;

  switch (type.toUpperCase()) {
    case 'R':
      sheetName = SHEET_NAMES.BACKLOG_R;
      break;
    case 'EMS':
      sheetName = SHEET_NAMES.BACKLOG_EMS;
      break;
    case 'COD':
      sheetName = SHEET_NAMES.BACKLOG_COD;
      break;
    default:
      throw new Error('ประเภทงานไม่ถูกต้อง: ' + type);
  }

  return getWorkData(sheetName);
}

/**
 * ฟังก์ชันดึงข้อมูลงานคืน (Return)
 * @param {string} type - ประเภทงาน (R, EMS, COD)
 * @return {Array} - รายการงานคืน
 */
function getReturnData(type) {
  let sheetName;

  switch (type.toUpperCase()) {
    case 'R':
      sheetName = SHEET_NAMES.RETURN_R;
      break;
    case 'EMS':
      sheetName = SHEET_NAMES.RETURN_EMS;
      break;
    case 'COD':
      sheetName = SHEET_NAMES.RETURN_COD;
      break;
    default:
      throw new Error('ประเภทงานไม่ถูกต้อง: ' + type);
  }

  return getWorkData(sheetName);
}

/**
 * ฟังก์ชันดึงข้อมูลทั้งหมดสำหรับ Dashboard
 * @return {Object} - ข้อมูลทั้งหมด
 */
function getAllDashboardData() {
  try {
    Logger.log('=== getAllDashboardData called ===');

    Logger.log('Fetching employees...');
    const employees = getAllEmployees();
    Logger.log('Employees fetched: ' + employees.length);

    Logger.log('Fetching backlog data...');
    const backlogR = getBacklogData('R');
    Logger.log('Backlog R: ' + backlogR.length);

    const backlogEMS = getBacklogData('EMS');
    Logger.log('Backlog EMS: ' + backlogEMS.length);

    const backlogCOD = getBacklogData('COD');
    Logger.log('Backlog COD: ' + backlogCOD.length);

    Logger.log('Fetching return data...');
    const returnR = getReturnData('R');
    Logger.log('Return R: ' + returnR.length);

    const returnEMS = getReturnData('EMS');
    Logger.log('Return EMS: ' + returnEMS.length);

    const returnCOD = getReturnData('COD');
    Logger.log('Return COD: ' + returnCOD.length);

    const data = {
      employees: employees,
      backlog: {
        r: backlogR,
        ems: backlogEMS,
        cod: backlogCOD
      },
      returned: {
        r: returnR,
        ems: returnEMS,
        cod: returnCOD
      },
      // Field names สำหรับ client-side ใช้งาน
      fields: {
        EMPLOYEE: {
          USERNAME: 'username',
          NAME: 'name',
          PAYMENT_SIDE: 'paymentSide',
          ACCESS_LEVEL: 'accessLevel',
          POSITION: 'position'
        },
        WORK_ITEM: {
          BARCODE: 'barcode',
          OPERATOR: 'operator',
          RECIPIENT_INFO: 'recipientInfo',
          SCAN_DATETIME: 'scanDatetime',
          REASON: 'reason',
          COD_AMOUNT: 'codAmount',
          IS_LAZADA: 'isLazada',
          DAYS_FROM_DEPOSIT: 'daysFromDeposit',
          DAYS_FROM_DESTINATION: 'daysFromDestination',
          ATTEMPT_COUNT: 'attemptCount'
        }
      }
    };

    Logger.log('Data structure created successfully');
    Logger.log('=== getAllDashboardData completed ===');

    return data;

  } catch (error) {
    Logger.log('ERROR in getAllDashboardData: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    // Return empty structure instead of throwing
    return {
      employees: [],
      backlog: { r: [], ems: [], cod: [] },
      returned: { r: [], ems: [], cod: [] },
      fields: {
        EMPLOYEE: {
          USERNAME: 'username',
          NAME: 'name',
          PAYMENT_SIDE: 'paymentSide',
          ACCESS_LEVEL: 'accessLevel',
          POSITION: 'position'
        },
        WORK_ITEM: {
          BARCODE: 'barcode',
          OPERATOR: 'operator',
          RECIPIENT_INFO: 'recipientInfo',
          SCAN_DATETIME: 'scanDatetime',
          REASON: 'reason',
          COD_AMOUNT: 'codAmount',
          IS_LAZADA: 'isLazada',
          DAYS_FROM_DEPOSIT: 'daysFromDeposit',
          DAYS_FROM_DESTINATION: 'daysFromDestination',
          ATTEMPT_COUNT: 'attemptCount'
        }
      },
      error: error.message
    };
  }
}

/**
 * ฟังก์ชันดึงข้อมูลสำหรับพนักงานคนเดียว
 * @param {string} username - Username ของพนักงาน
 * @return {Object} - ข้อมูลของพนักงาน
 */
function getEmployeeData(username) {
  try {
    const allData = getAllDashboardData();

    // ฟังก์ชันกรองข้อมูลเฉพาะพนักงานคนนี้
    const filterByUsername = function(items) {
      return items.filter(function(item) {
        return item.operator && item.operator.toString().trim() === username.trim();
      });
    };

    return {
      employees: allData.employees.filter(function(emp) {
        return emp.username === username.trim();
      }),
      backlog: {
        r: filterByUsername(allData.backlog.r),
        ems: filterByUsername(allData.backlog.ems),
        cod: filterByUsername(allData.backlog.cod)
      },
      returned: {
        r: filterByUsername(allData.returned.r),
        ems: filterByUsername(allData.returned.ems),
        cod: filterByUsername(allData.returned.cod)
      },
      fields: allData.fields
    };

  } catch (error) {
    Logger.log('Error in getEmployeeData: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลพนักงาน: ' + error.message);
  }
}

/**
 * ฟังก์ชันค้นหางานจาก Barcode
 * @param {string} barcode - Barcode ที่ต้องการค้นหา
 * @return {Object} - ข้อมูลงานที่พบ
 */
function searchByBarcode(barcode) {
  try {
    const searchBarcode = barcode.toString().trim().toUpperCase();
    const results = [];

    // ค้นหาในงานค้าง
    ['R', 'EMS', 'COD'].forEach(function(type) {
      const backlogData = getBacklogData(type);
      backlogData.forEach(function(item) {
        if (item.barcode.toUpperCase() === searchBarcode) {
          results.push({
            type: 'backlog',
            category: type,
            data: item
          });
        }
      });
    });

    // ค้นหาในงานคืน
    ['R', 'EMS', 'COD'].forEach(function(type) {
      const returnData = getReturnData(type);
      returnData.forEach(function(item) {
        if (item.barcode.toUpperCase() === searchBarcode) {
          results.push({
            type: 'return',
            category: type,
            data: item
          });
        }
      });
    });

    return {
      success: true,
      results: results,
      count: results.length,
      barcode: searchBarcode
    };

  } catch (error) {
    Logger.log('Error in searchByBarcode: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการค้นหา: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันนับจำนวนงานตามพนักงาน
 * @param {string} username - Username ของพนักงาน
 * @return {Object} - สถิติงาน
 */
function getEmployeeWorkStats(username) {
  try {
    const data = getEmployeeData(username);

    return {
      username: username,
      backlog: {
        r: data.backlog.r.length,
        ems: data.backlog.ems.length,
        cod: data.backlog.cod.length,
        total: data.backlog.r.length + data.backlog.ems.length + data.backlog.cod.length
      },
      returned: {
        r: data.returned.r.length,
        ems: data.returned.ems.length,
        cod: data.returned.cod.length,
        total: data.returned.r.length + data.returned.ems.length + data.returned.cod.length
      },
      totalWork: data.backlog.r.length + data.backlog.ems.length + data.backlog.cod.length +
                 data.returned.r.length + data.returned.ems.length + data.returned.cod.length
    };

  } catch (error) {
    Logger.log('Error in getEmployeeWorkStats: ' + error.message);
    return null;
  }
}

/**
 * ฟังก์ชันดึงสถิติรวมทั้งหมด
 * @return {Object} - สถิติรวม
 */
function getTotalStats() {
  try {
    const data = getAllDashboardData();

    return {
      employees: data.employees.length,
      backlog: {
        r: data.backlog.r.length,
        ems: data.backlog.ems.length,
        cod: data.backlog.cod.length,
        total: data.backlog.r.length + data.backlog.ems.length + data.backlog.cod.length
      },
      returned: {
        r: data.returned.r.length,
        ems: data.returned.ems.length,
        cod: data.returned.cod.length,
        total: data.returned.r.length + data.returned.ems.length + data.returned.cod.length
      }
    };

  } catch (error) {
    Logger.log('Error in getTotalStats: ' + error.message);
    return null;
  }
}
