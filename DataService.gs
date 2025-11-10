/**
 * DataService.gs
 * ระบบจัดการข้อมูลจาก Google Sheets
 */

/**
 * ฟังก์ชันดึงข้อมูลพนักงานทั้งหมด
 * @return {Array} - รายการพนักงาน
 */
function getEmployees() {
  try {
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();

    const employees = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // ข้าม row ที่ไม่มี username
      if (!row[EMPLOYEE_COLUMNS.USERNAME]) continue;

      employees.push({
        username: row[EMPLOYEE_COLUMNS.USERNAME],
        name: row[EMPLOYEE_COLUMNS.NAME],
        paymentSide: row[EMPLOYEE_COLUMNS.PAYMENT_SIDE] || 'ไม่ระบุ'
      });
    }

    return employees;

  } catch (error) {
    Logger.log('Error in getEmployees: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลพนักงาน: ' + error.message);
  }
}

/**
 * ฟังก์ชันดึงข้อมูลงานค้าง (Backlog)
 * @param {string} type - ประเภทงาน (R, EMS, COD)
 * @return {Array} - รายการงานค้าง
 */
function getBacklogData(type) {
  try {
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

    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();

    const backlog = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // ข้าม row ที่ไม่มี barcode
      if (!row[WORK_COLUMNS.BARCODE]) continue;

      backlog.push({
        barcode: row[WORK_COLUMNS.BARCODE],
        operator: row[WORK_COLUMNS.OPERATOR],
        date: row[WORK_COLUMNS.DATE],
        status: row[WORK_COLUMNS.STATUS],
        paymentSide: row[WORK_COLUMNS.PAYMENT_SIDE]
      });
    }

    return backlog;

  } catch (error) {
    Logger.log('Error in getBacklogData: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลงานค้าง: ' + error.message);
  }
}

/**
 * ฟังก์ชันดึงข้อมูลงานคืน (Return)
 * @param {string} type - ประเภทงาน (R, EMS, COD)
 * @return {Array} - รายการงานคืน
 */
function getReturnData(type) {
  try {
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

    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();

    const returnData = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // ข้าม row ที่ไม่มี barcode
      if (!row[WORK_COLUMNS.BARCODE]) continue;

      returnData.push({
        barcode: row[WORK_COLUMNS.BARCODE],
        operator: row[WORK_COLUMNS.OPERATOR],
        date: row[WORK_COLUMNS.DATE],
        status: row[WORK_COLUMNS.STATUS],
        paymentSide: row[WORK_COLUMNS.PAYMENT_SIDE]
      });
    }

    return returnData;

  } catch (error) {
    Logger.log('Error in getReturnData: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลงานคืน: ' + error.message);
  }
}

/**
 * ฟังก์ชันดึงข้อมูลทั้งหมดสำหรับ Dashboard
 * @return {Object} - ข้อมูลทั้งหมด
 */
function getAllDashboardData() {
  try {
    const data = {
      employees: getEmployees(),
      backlog: {
        r: getBacklogData('R'),
        ems: getBacklogData('EMS'),
        cod: getBacklogData('COD')
      },
      returned: {
        r: getReturnData('R'),
        ems: getReturnData('EMS'),
        cod: getReturnData('COD')
      },
      fields: {
        EMPLOYEE: {
          USERNAME: 'username',
          NAME: 'name',
          PAYMENT_SIDE: 'paymentSide'
        },
        WORK_ITEM: {
          BARCODE: 'barcode',
          OPERATOR: 'operator',
          DATE: 'date',
          STATUS: 'status',
          PAYMENT_SIDE: 'paymentSide'
        }
      }
    };

    return data;

  } catch (error) {
    Logger.log('Error in getAllDashboardData: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูล: ' + error.message);
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

    // กรองข้อมูลเฉพาะพนักงานคนนี้
    const filterByUsername = function(items) {
      return items.filter(function(item) {
        return item.operator === username;
      });
    };

    return {
      employees: allData.employees.filter(function(emp) {
        return emp.username === username;
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
 * ฟังก์ชันสำหรับค้นหางานจาก Barcode
 * @param {string} barcode - Barcode ที่ต้องการค้นหา
 * @return {Object} - ข้อมูลงานที่พบ
 */
function searchByBarcode(barcode) {
  try {
    const allData = getAllDashboardData();
    const results = [];

    // ค้นหาในงานค้าง
    ['r', 'ems', 'cod'].forEach(function(type) {
      allData.backlog[type].forEach(function(item) {
        if (item.barcode === barcode) {
          results.push({
            type: 'backlog',
            category: type.toUpperCase(),
            data: item
          });
        }
      });
    });

    // ค้นหาในงานคืน
    ['r', 'ems', 'cod'].forEach(function(type) {
      allData.returned[type].forEach(function(item) {
        if (item.barcode === barcode) {
          results.push({
            type: 'return',
            category: type.toUpperCase(),
            data: item
          });
        }
      });
    });

    return {
      success: true,
      results: results,
      count: results.length
    };

  } catch (error) {
    Logger.log('Error in searchByBarcode: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการค้นหา: ' + error.message
    };
  }
}
