// ================================================================
// 🏢 SDIP-V2 - Google Apps Script Version (Fixed)
// ระบบตรวจสอบงานค้าง - ไปรษณีย์ไทย
// ================================================================

/**
 * ⚙️ Configuration
 */
const CONFIG = {
  // Firebase Realtime Database URL
  FIREBASE_URL: 'https://x-85fc7-default-rtdb.asia-southeast1.firebasedatabase.app',

  // ⚠️ IMPORTANT: ใส่ Firebase Database Secret ที่นี่
  // หา Secret ได้ที่: Firebase Console > Project Settings > Service Accounts > Database secrets
  FIREBASE_SECRET: '-FgpZbl8mveCB7YxzRVo9pkLMuc5T33AmmaN7u4WF', // Firebase Database Secret

  // Base Path - ข้อมูลอยู่ภายใต้ path นี้
  BASE_PATH: 'สำเนาของ 004xProgram SDIP 84180',

  // Collections
  COLLECTIONS: {
    EMPLOYEE: 'สำเนาของ 004xProgram SDIP 84180/Employee',
    BACKLOG_R: 'สำเนาของ 004xProgram SDIP 84180/SDIPBacklogR',
    BACKLOG_EMS: 'สำเนาของ 004xProgram SDIP 84180/SDIPBacklogEMS',
    BACKLOG_COD: 'สำเนาของ 004xProgram SDIP 84180/SDIPBacklogCOD',
    RETURNED_R: 'สำเนาของ 004xProgram SDIP 84180/SDIPReturnedR',
    RETURNED_EMS: 'สำเนาของ 004xProgram SDIP 84180/SDIPReturnedEMS',
    RETURNED_COD: 'สำเนาของ 004xProgram SDIP 84180/SDIPReturnedCOD',
    SDIP_WMS: 'สำเนาของ 004xProgram SDIP 84180/SDIPWMS',
    SDIP_WRP: 'สำเนาของ 004xProgram SDIP 84180/SDIPWRP'
  },

  // External URLs
  URLS: {
    FORM_301: 'https://script.google.com/macros/s/AKfycbx6a6jmUzx1MFVRWYRgHxr_5SAE8bkkaPeNZ3oFSJgrLcgXY-anxuAb1hiDVsvfeYY0-Q/exec',
    REPORT_301: 'https://script.google.com/macros/s/AKfycbz7aD_2SSu0VePe1L6JBg6RbNw8phOyOyxg3xViHQjDhRIoD4kavAHktEIANf7alCer9Q/exec'
  },

  // Field Names
  FIELDS: {
    EMPLOYEE: {
      ORDER: '01_ลำดับ',
      STATUS: '02_UserStatus',
      USERNAME: '03_Username',
      PASSWORD: '04_Password',
      ACCESS_LEVEL: '05_Useraccesslevel',
      NAME: '06_Name',
      POSITION: '07_Position',
      PAYMENT_SIDE: '08_Paymentside',
      DISTRICT: '09_district',
      VILLAGE_NO: '10_VillageNo_',
      WORK_PHONE: '11_WorkPhone'
    },
    WORK_ITEM: {
      ORDER: '01_ลำดับ',
      TRACKING_NUMBER: '02_หมายเลขสิ่งของ',
      RECIPIENT: '03_ชือ - ที่อยู่ผู้รับ',
      OPERATOR: '04_ผู้ดำเนินการ',
      SCAN_DATE: '05_วันที่ เวลา (สแกนผลการนำจ่ายล่าสุด)',
      REASON: '06_สาเหตุ',
      COD: '07_COD',
      LAZADA: '08_Lazada?',
      DAYS_HELD_1: '09_จำนวนวันที่ถือครอง (นับจากวันที่รับฝาก)',
      DAYS_HELD_2: '10_จำนวนวันที่ถือครอง (นับจากวันที่ปลายทาง)',
      ATTEMPTS: '11_จำนวนครั้งที่พยายามนำจ่าย'
    },
    WMS: {
      TRACKING_NUMBER: '02_หมายเลขสิ่งของ',
      RECIPIENT: '03_ชือ - ที่อยู่ผู้รับ',
      OPERATOR: '05_ผู้ดำเนินการ',
      AMOUNT: '08_จำนวนเงิน (บาท)',
      DAYS_HELD: '10_จำนวนวันที่ถือครองเงิน (นับจากวันที่เก็บเงิน)',
      STATUS: '11_รายการ'
    },
    WRP: {
      TRACKING_NUMBER: '02_หมายเลขสิ่งของ',
      RECIPIENT: '03_ชือ - ที่อยู่ผู้รับ',
      OPERATOR: '04_ผู้ดำเนินการ',
      SCAN_DATE: '05_วันที่ เวลา (สแกนสิ่งของ)',
      DELIVERY_STATUS: '06_สถานะบันทึกนำจ่าย',
      COD: '07_COD  '  // มีช่องว่าง 2 ตัวต่อท้าย (ตรงกับข้อมูลจริงใน Firebase)
    }
  }
};

// ================================================================
// 🌐 Web App Entry Point
// ================================================================

/**
 * doGet - Handle GET requests
 */
function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'login';

  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('📊 Dashboard - SDIP-V2')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Default: Login page
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .setTitle('🔐 เข้าสู่ระบบ - SDIP-V2')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get Script URL
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Get 301 Form/Report URLs
 */
function get301URLs() {
  return {
    form301: CONFIG.URLS.FORM_301,
    report301: CONFIG.URLS.REPORT_301
  };
}

// ================================================================
// 🔥 Firebase Helper Functions (FIXED)
// ================================================================

/**
 * ดึงข้อมูลจาก Firebase Realtime Database
 * วิธีแก้: เพิ่ม auth parameter และจัดการ error ดีขึ้น
 */
function getFirebaseData(path) {
  try {
    var url = CONFIG.FIREBASE_URL + '/' + path + '.json';

    // เพิ่ม auth parameter (ถ้ามี Secret)
    if (CONFIG.FIREBASE_SECRET && CONFIG.FIREBASE_SECRET.length > 0) {
      url += '?auth=' + CONFIG.FIREBASE_SECRET;
    }

    Logger.log('Fetching from: ' + url);

    var options = {
      method: 'get',
      contentType: 'application/json',
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log('Response Code: ' + responseCode);

    if (responseCode === 200) {
      if (!responseText || responseText === 'null') {
        Logger.log('Warning: Empty data for path: ' + path);
        return null;
      }
      var data = JSON.parse(responseText);
      return data;
    } else {
      Logger.log('Firebase GET Error: ' + responseCode + ' - ' + responseText);

      // แสดงข้อความ error ที่เข้าใจง่าย
      if (responseCode === 401) {
        throw new Error('Firebase Authentication Error: กรุณาตรวจสอบ FIREBASE_SECRET');
      } else if (responseCode === 403) {
        throw new Error('Firebase Permission Error: Database Rules ไม่อนุญาตให้เข้าถึง');
      } else if (responseCode === 404) {
        throw new Error('Firebase Error: ไม่พบข้อมูลที่ path: ' + path);
      } else {
        throw new Error('Firebase Error: ' + responseCode + ' - ' + responseText);
      }
    }
  } catch (error) {
    Logger.log('getFirebaseData Error: ' + error.toString());
    throw error;
  }
}

/**
 * เขียนข้อมูลไป Firebase
 */
function setFirebaseData(path, data) {
  try {
    var url = CONFIG.FIREBASE_URL + '/' + path + '.json';

    if (CONFIG.FIREBASE_SECRET && CONFIG.FIREBASE_SECRET.length > 0) {
      url += '?auth=' + CONFIG.FIREBASE_SECRET;
    }

    var options = {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode === 200) {
      return true;
    } else {
      Logger.log('Firebase PUT Error: ' + responseCode + ' - ' + response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log('setFirebaseData Error: ' + error.toString());
    return false;
  }
}

/**
 * แปลง Firebase Object/Array เป็น Array
 */
function toArray(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;

  var result = [];
  for (var key in data) {
    if (data.hasOwnProperty(key)) {
      var item = data[key];
      if (typeof item === 'object' && item !== null) {
        item._key = key;
        result.push(item);
      }
    }
  }
  return result;
}

// ================================================================
// 🔐 Authentication Functions
// ================================================================

/**
 * ตรวจสอบ Login
 */
function checkLogin(username, password) {
  try {
    if (!username || !password) {
      return { success: false, message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
    }

    // Get employee data from Firebase
    var employees = getFirebaseData(CONFIG.COLLECTIONS.EMPLOYEE);

    if (!employees) {
      return { success: false, message: 'ไม่สามารถเชื่อมต่อฐานข้อมูลได้ กรุณาตรวจสอบการตั้งค่า' };
    }

    var employeeArray = toArray(employees);

    if (employeeArray.length === 0) {
      return { success: false, message: 'ไม่พบข้อมูลพนักงานในระบบ' };
    }

    // Find matching user
    for (var i = 0; i < employeeArray.length; i++) {
      var emp = employeeArray[i];
      var empUsername = emp[CONFIG.FIELDS.EMPLOYEE.USERNAME];
      var empPassword = emp[CONFIG.FIELDS.EMPLOYEE.PASSWORD];
      var empStatus = emp[CONFIG.FIELDS.EMPLOYEE.STATUS];
      var empAccessLevel = emp[CONFIG.FIELDS.EMPLOYEE.ACCESS_LEVEL] || 'User';

      if (empUsername &&
          empUsername.toLowerCase() === username.toLowerCase() &&
          empPassword === password &&
          empStatus === 'Active') {

        return {
          success: true,
          username: empUsername,
          name: emp[CONFIG.FIELDS.EMPLOYEE.NAME],
          accessLevel: empAccessLevel,
          paymentSide: emp[CONFIG.FIELDS.EMPLOYEE.PAYMENT_SIDE],
          position: emp[CONFIG.FIELDS.EMPLOYEE.POSITION]
        };
      }
    }

    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };

  } catch (error) {
    Logger.log('checkLogin Error: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.message,
      details: error.toString()
    };
  }
}

// ================================================================
// 📊 Data Retrieval Functions
// ================================================================

/**
 * ดึงข้อมูลพนักงานทั้งหมด
 */
function getAllEmployees() {
  try {
    var employees = getFirebaseData(CONFIG.COLLECTIONS.EMPLOYEE);
    if (!employees) return [];

    return toArray(employees).filter(function(emp) {
      return emp && emp[CONFIG.FIELDS.EMPLOYEE.STATUS] === 'Active';
    });
  } catch (error) {
    Logger.log('getAllEmployees Error: ' + error.toString());
    throw error;
  }
}

/**
 * ดึงข้อมูลงานค้าง
 */
function getBacklogData() {
  try {
    var backlogR = toArray(getFirebaseData(CONFIG.COLLECTIONS.BACKLOG_R)) || [];
    var backlogEMS = toArray(getFirebaseData(CONFIG.COLLECTIONS.BACKLOG_EMS)) || [];
    var backlogCOD = toArray(getFirebaseData(CONFIG.COLLECTIONS.BACKLOG_COD)) || [];

    return {
      r: backlogR,
      ems: backlogEMS,
      cod: backlogCOD
    };
  } catch (error) {
    Logger.log('getBacklogData Error: ' + error.toString());
    return { r: [], ems: [], cod: [] };
  }
}

/**
 * ดึงข้อมูลงานคืน
 */
function getReturnedData() {
  try {
    var returnedR = toArray(getFirebaseData(CONFIG.COLLECTIONS.RETURNED_R)) || [];
    var returnedEMS = toArray(getFirebaseData(CONFIG.COLLECTIONS.RETURNED_EMS)) || [];
    var returnedCOD = toArray(getFirebaseData(CONFIG.COLLECTIONS.RETURNED_COD)) || [];

    return {
      r: returnedR,
      ems: returnedEMS,
      cod: returnedCOD
    };
  } catch (error) {
    Logger.log('getReturnedData Error: ' + error.toString());
    return { r: [], ems: [], cod: [] };
  }
}

/**
 * ดึงข้อมูลเงิน COD
 */
function getMoneyData() {
  try {
    var wmsData = getFirebaseData(CONFIG.COLLECTIONS.SDIP_WMS);
    return toArray(wmsData) || [];
  } catch (error) {
    Logger.log('getMoneyData Error: ' + error.toString());
    return [];
  }
}

/**
 * ดึงข้อมูลรอบันทึก (แยก R/EMS/COD)
 * แยกเป็น 2 ชุด: prepare (ทั้งหมด) และ record (กรองแล้ว)
 */
function getRecordData() {
  try {
    var wrpData = getFirebaseData(CONFIG.COLLECTIONS.SDIP_WRP);
    var allItems = toArray(wrpData) || [];

    // เตรียม = ทั้งหมด
    var prepareR = [];
    var prepareEMS = [];
    var prepareCOD = [];

    // บันทึก = กรองแล้ว (ไม่มี "สถานะ" ในคอลัม F)
    var recordR = [];
    var recordEMS = [];
    var recordCOD = [];

    Logger.log('=== getRecordData Debug ===');
    Logger.log('Total items: ' + allItems.length);

    allItems.forEach(function(item, index) {
      // ใช้ WRP fields
      var codField = item[CONFIG.FIELDS.WRP.COD]; // คอลัม G (07_COD  )
      var trackingNumber = item[CONFIG.FIELDS.WRP.TRACKING_NUMBER]; // คอลัม B
      var deliveryStatus = item[CONFIG.FIELDS.WRP.DELIVERY_STATUS]; // คอลัม F

      // Normalize ข้อมูล
      var codValue = codField ? String(codField).trim() : '';
      var trackingValue = trackingNumber ? String(trackingNumber).trim() : '';
      var statusValue = deliveryStatus ? String(deliveryStatus) : '';

      // ตรวจสอบว่ามี "สถานะ" ในคอลัม F หรือไม่
      var hasStatus = statusValue.includes('สถานะ');

      // Log รายละเอียด 5 รายการแรก
      if (index < 5) {
        Logger.log('---');
        Logger.log('Item #' + index);
        Logger.log('  Tracking: "' + trackingValue + '"');
        Logger.log('  COD: "' + codValue + '"');
        Logger.log('  Status: "' + statusValue + '"');
        Logger.log('  Has "สถานะ": ' + hasStatus);
      }

      // แยกประเภทตามกฎ (R/EMS/COD)
      var categoryArrayPrepare, categoryArrayRecord;

      // 1. COD: คอลัม G = "yes"
      if (codValue.toLowerCase() === 'yes') {
        categoryArrayPrepare = prepareCOD;
        categoryArrayRecord = recordCOD;
      }
      // 2. EMS: คอลัม G = "NO" และคอลัม B ขึ้นต้นด้วย J, W, หรือ E
      else if (codValue.toUpperCase() === 'NO') {
        if (trackingValue && /^[JWE]/i.test(trackingValue)) {
          categoryArrayPrepare = prepareEMS;
          categoryArrayRecord = recordEMS;
        }
        // 3. R: คอลัม G = "NO" และคอลัม B ไม่ได้ขึ้นต้นด้วย J, W, หรือ E
        else {
          categoryArrayPrepare = prepareR;
          categoryArrayRecord = recordR;
        }
      }
      // ถ้าไม่ตรงเงื่อนไข ให้ใส่ใน R
      else {
        categoryArrayPrepare = prepareR;
        categoryArrayRecord = recordR;
      }

      // เพิ่มเข้า prepare ทุกรายการ
      categoryArrayPrepare.push(item);

      // เพิ่มเข้า record เฉพาะที่ไม่มี "สถานะ"
      if (!hasStatus) {
        categoryArrayRecord.push(item);
      }

      if (index < 5) {
        Logger.log('  → Added to prepare, record: ' + (hasStatus ? 'NO (has status)' : 'YES'));
      }
    });

    Logger.log('=== Results ===');
    Logger.log('Prepare - R: ' + prepareR.length + ', EMS: ' + prepareEMS.length + ', COD: ' + prepareCOD.length);
    Logger.log('Record - R: ' + recordR.length + ', EMS: ' + recordEMS.length + ', COD: ' + recordCOD.length);
    Logger.log('================');

    return {
      prepare: {
        r: prepareR,
        ems: prepareEMS,
        cod: prepareCOD
      },
      record: {
        r: recordR,
        ems: recordEMS,
        cod: recordCOD
      }
    };
  } catch (error) {
    Logger.log('getRecordData Error: ' + error.toString());
    return {
      prepare: { r: [], ems: [], cod: [] },
      record: { r: [], ems: [], cod: [] }
    };
  }
}

/**
 * ดึงข้อมูลทั้งหมดพร้อมกัน (พร้อมคำนวณสถิติ)
 */
function getAllDashboardData() {
  try {
    var employees = getAllEmployees();
    var backlogData = getBacklogData();
    var returnedData = getReturnedData();
    var moneyData = getMoneyData();
    var recordData = getRecordData();

    // คำนวณสถิติ byEmployee
    var backlogStats = calculateStatsByEmployee(backlogData);
    var returnedStats = calculateStatsByEmployee(returnedData);
    var prepareStats = calculateStatsByEmployee(recordData.prepare);
    var recordStats = calculateStatsByEmployee(recordData.record);

    return {
      employees: employees,
      backlog: backlogData,
      returned: returnedData,
      money: moneyData,
      prepare: recordData.prepare,
      record: recordData.record,
      fields: CONFIG.FIELDS,
      timestamp: new Date().toISOString(),
      // เพิ่มสถิติรวม
      backlogStats: {
        total: backlogData.r.length + backlogData.ems.length + backlogData.cod.length,
        r: backlogData.r.length,
        ems: backlogData.ems.length,
        cod: backlogData.cod.length,
        byEmployee: backlogStats
      },
      returnStats: {
        total: returnedData.r.length + returnedData.ems.length + returnedData.cod.length,
        r: returnedData.r.length,
        ems: returnedData.ems.length,
        cod: returnedData.cod.length,
        byEmployee: returnedStats
      },
      prepareStats: {
        total: recordData.prepare.r.length + recordData.prepare.ems.length + recordData.prepare.cod.length,
        r: recordData.prepare.r.length,
        ems: recordData.prepare.ems.length,
        cod: recordData.prepare.cod.length,
        byEmployee: prepareStats
      },
      recordStats: {
        total: recordData.record.r.length + recordData.record.ems.length + recordData.record.cod.length,
        r: recordData.record.r.length,
        ems: recordData.record.ems.length,
        cod: recordData.record.cod.length,
        byEmployee: recordStats
      }
    };
  } catch (error) {
    Logger.log('getAllDashboardData Error: ' + error.toString());
    return {
      employees: [],
      backlog: { r: [], ems: [], cod: [] },
      returned: { r: [], ems: [], cod: [] },
      money: [],
      prepare: { r: [], ems: [], cod: [] },
      record: { r: [], ems: [], cod: [] },
      fields: CONFIG.FIELDS,
      timestamp: new Date().toISOString(),
      error: error.toString()
    };
  }
}

/**
 * คำนวณสถิติแยกตาม Employee
 */
function calculateStatsByEmployee(dataObj) {
  var stats = {};

  // นับจาก R, EMS, COD
  ['r', 'ems', 'cod'].forEach(function(type) {
    var items = dataObj[type] || [];
    items.forEach(function(item) {
      var operator = item[CONFIG.FIELDS.WORK_ITEM.OPERATOR];
      if (!operator) return;

      if (!stats[operator]) {
        stats[operator] = { r: 0, ems: 0, cod: 0, total: 0 };
      }

      stats[operator][type]++;
      stats[operator].total++;
    });
  });

  return stats;
}

/**
 * ดึงรายละเอียดงานสำหรับแสดงในตาราง (เมื่อคลิกที่ตัวเลข)
 * @param {string} dataType - 'backlog', 'return', 'money', 'prepare', 'record'
 * @param {string} workType - 'R', 'EMS', 'COD', หรือ null (สำหรับ money)
 * @param {string} employeeFilter - username ของพนักงาน (optional)
 * @return {object} { headers: [], data: [[...], [...]] }
 */
function getDetailWorkData(dataType, workType, employeeFilter) {
  try {
    var headers = [];
    var items = [];
    var rows = [];

    // ===== MONEY DATA =====
    if (dataType === 'money') {
      Logger.log('=== getDetailWorkData: MONEY ===');
      Logger.log('Employee Filter: ' + employeeFilter);

      headers = [
        'ลำดับ',
        'หมายเลขติดตาม',
        'รายละเอียด',
        'ผู้ดำเนินการ',
        'จำนวนเงิน (บาท)',
        'วันถือครองเงิน',
        'รายการ'
      ];

      // ดึงข้อมูล WMS
      var wmsData = getFirebaseData(CONFIG.COLLECTIONS.SDIP_WMS);
      items = toArray(wmsData);

      Logger.log('Total WMS items: ' + items.length);

      // Log first item fields to check structure
      if (items.length > 0) {
        Logger.log('First item keys: ' + Object.keys(items[0]).join(', '));
        Logger.log('Sample operator field: ' + items[0][CONFIG.FIELDS.WMS.OPERATOR]);
      }

      var matchCount = 0;
      items.forEach(function(item, index) {
        var operator = item[CONFIG.FIELDS.WMS.OPERATOR];

        // Filter by employee ถ้ามี
        if (employeeFilter) {
          // Log comparison for debugging
          if (index < 10) {
            Logger.log('Item ' + index + ' - Operator: "' + operator + '" | Filter: "' + employeeFilter + '"');
            Logger.log('  Match (===): ' + (operator === employeeFilter));
            Logger.log('  Match (trimmed): ' + ((operator || '').trim() === (employeeFilter || '').trim()));
          }

          // Use trimmed comparison to avoid whitespace issues
          if (!operator || (operator.trim() !== employeeFilter.trim())) {
            return;
          }
          matchCount++;
        }

        rows.push([
          index + 1,
          item[CONFIG.FIELDS.WMS.TRACKING_NUMBER] || '',
          item[CONFIG.FIELDS.WMS.RECIPIENT] || '',
          operator || '',
          item[CONFIG.FIELDS.WMS.AMOUNT] || '',
          item[CONFIG.FIELDS.WMS.DAYS_HELD] || '',
          item[CONFIG.FIELDS.WMS.STATUS] || ''
        ]);
      });

      if (employeeFilter) {
        Logger.log('Filter "' + employeeFilter + '" matched ' + matchCount + ' out of ' + items.length + ' items');
      }

      Logger.log('Total rows after filtering: ' + rows.length);
      return { headers: headers, data: rows };
    }

    // ===== PREPARE DATA =====
    if (dataType === 'prepare') {
      headers = [
        'ลำดับ',
        'หมายเลขติดตาม',
        'รายละเอียด',
        'ผู้ดำเนินการ',
        'วันที่สแกน',
        'สถานะบันทึกนำจ่าย',
        'COD'
      ];

      // ดึงข้อมูล prepare (จาก getRecordData)
      var recordData = getRecordData();

      if (workType === 'R') items = recordData.prepare.r;
      else if (workType === 'EMS') items = recordData.prepare.ems;
      else if (workType === 'COD') items = recordData.prepare.cod;

      items.forEach(function(item, index) {
        var operator = item[CONFIG.FIELDS.WRP.OPERATOR];

        // Filter by employee ถ้ามี
        if (employeeFilter && operator !== employeeFilter) {
          return;
        }

        rows.push([
          index + 1,
          item[CONFIG.FIELDS.WRP.TRACKING_NUMBER] || '',
          item[CONFIG.FIELDS.WRP.RECIPIENT] || '',
          operator || '',
          item[CONFIG.FIELDS.WRP.SCAN_DATE] || '',
          item[CONFIG.FIELDS.WRP.DELIVERY_STATUS] || '',
          item[CONFIG.FIELDS.WRP.COD] || ''
        ]);
      });

      return { headers: headers, data: rows };
    }

    // ===== RECORD DATA =====
    if (dataType === 'record') {
      headers = [
        'ลำดับ',
        'หมายเลขติดตาม',
        'รายละเอียด',
        'ผู้ดำเนินการ',
        'วันที่สแกน',
        'สถานะบันทึกนำจ่าย',
        'COD'
      ];

      // ดึงข้อมูล record (จาก getRecordData)
      var recordData = getRecordData();

      if (workType === 'R') items = recordData.record.r;
      else if (workType === 'EMS') items = recordData.record.ems;
      else if (workType === 'COD') items = recordData.record.cod;

      items.forEach(function(item, index) {
        var operator = item[CONFIG.FIELDS.WRP.OPERATOR];

        // Filter by employee ถ้ามี
        if (employeeFilter && operator !== employeeFilter) {
          return;
        }

        rows.push([
          index + 1,
          item[CONFIG.FIELDS.WRP.TRACKING_NUMBER] || '',
          item[CONFIG.FIELDS.WRP.RECIPIENT] || '',
          operator || '',
          item[CONFIG.FIELDS.WRP.SCAN_DATE] || '',
          item[CONFIG.FIELDS.WRP.DELIVERY_STATUS] || '',
          item[CONFIG.FIELDS.WRP.COD] || ''
        ]);
      });

      return { headers: headers, data: rows };
    }

    // ===== BACKLOG/RETURN DATA (existing logic) =====
    headers = [
      'ลำดับ',
      'หมายเลขติดตาม',
      'รายละเอียด',
      'ผู้ดำเนินการ',
      'วันที่สแกน',
      'เหตุผล',
      'COD',
      'Lazada',
      'วันถือครอง1',
      'วันถือครอง2',
      'ครั้งจัดส่ง'
    ];

    // เลือก collection ที่ถูกต้อง
    var collectionPath = '';
    if (dataType === 'backlog') {
      if (workType === 'R') collectionPath = CONFIG.COLLECTIONS.BACKLOG_R;
      else if (workType === 'EMS') collectionPath = CONFIG.COLLECTIONS.BACKLOG_EMS;
      else if (workType === 'COD') collectionPath = CONFIG.COLLECTIONS.BACKLOG_COD;
    } else if (dataType === 'return') {
      if (workType === 'R') collectionPath = CONFIG.COLLECTIONS.RETURNED_R;
      else if (workType === 'EMS') collectionPath = CONFIG.COLLECTIONS.RETURNED_EMS;
      else if (workType === 'COD') collectionPath = CONFIG.COLLECTIONS.RETURNED_COD;
    }

    if (!collectionPath) {
      return { headers: headers, data: [] };
    }

    // ดึงข้อมูล
    var rawData = getFirebaseData(collectionPath);
    items = toArray(rawData);

    // แปลงเป็น rows
    items.forEach(function(item, index) {
      var operator = item[CONFIG.FIELDS.WORK_ITEM.OPERATOR];

      // Filter by employee ถ้ามี
      if (employeeFilter && operator !== employeeFilter) {
        return;
      }

      rows.push([
        item[CONFIG.FIELDS.WORK_ITEM.ORDER] || (index + 1),
        item[CONFIG.FIELDS.WORK_ITEM.TRACKING_NUMBER] || '',
        item[CONFIG.FIELDS.WORK_ITEM.RECIPIENT] || '',
        operator || '',
        item[CONFIG.FIELDS.WORK_ITEM.SCAN_DATE] || '',
        item[CONFIG.FIELDS.WORK_ITEM.REASON] || '',
        item[CONFIG.FIELDS.WORK_ITEM.COD] || '',
        item[CONFIG.FIELDS.WORK_ITEM.LAZADA] || '',
        item[CONFIG.FIELDS.WORK_ITEM.DAYS_HELD_1] || '',
        item[CONFIG.FIELDS.WORK_ITEM.DAYS_HELD_2] || '',
        item[CONFIG.FIELDS.WORK_ITEM.ATTEMPTS] || ''
      ]);
    });

    return { headers: headers, data: rows };

  } catch (error) {
    Logger.log('getDetailWorkData Error: ' + error.toString());
    return {
      headers: [],
      data: [],
      error: error.toString()
    };
  }
}

// ================================================================
// 🧪 Test Functions
// ================================================================

/**
 * ทดสอบการเชื่อมต่อ Firebase
 */
function testFirebaseConnection() {
  Logger.log('=== Testing Firebase Connection ===');
  Logger.log('Firebase URL: ' + CONFIG.FIREBASE_URL);
  Logger.log('Has Secret: ' + (CONFIG.FIREBASE_SECRET.length > 0 ? 'Yes' : 'No'));

  try {
    var employees = getFirebaseData(CONFIG.COLLECTIONS.EMPLOYEE);

    if (employees) {
      var count = Object.keys(employees).length;
      Logger.log('✅ Firebase connection successful!');
      Logger.log('Number of employees: ' + count);

      // Show first employee as sample
      var firstKey = Object.keys(employees)[0];
      if (firstKey) {
        Logger.log('Sample employee: ' + JSON.stringify(employees[firstKey], null, 2));
      }

      return {
        success: true,
        count: count,
        message: 'เชื่อมต่อสำเร็จ'
      };
    } else {
      Logger.log('❌ No data returned from Firebase');
      return {
        success: false,
        message: 'ไม่มีข้อมูล'
      };
    }
  } catch (error) {
    Logger.log('❌ Firebase connection failed!');
    Logger.log('Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ทดสอบข้อมูล WRP Record (สำหรับ debug การแยก R/EMS/COD)
 */
function testWRPData() {
  Logger.log('=== Testing WRP Data ===');

  try {
    var wrpData = getFirebaseData(CONFIG.COLLECTIONS.SDIP_WRP);
    var allItems = toArray(wrpData) || [];

    Logger.log('Total WRP items: ' + allItems.length);

    if (allItems.length === 0) {
      Logger.log('❌ No WRP data found!');
      return;
    }

    // แสดง keys ทั้งหมดของ item แรก
    var firstItem = allItems[0];
    Logger.log('\n📋 Available fields in first item:');
    Logger.log(Object.keys(firstItem).join('\n'));

    // แสดงตัวอย่างข้อมูล 5 รายการแรก
    Logger.log('\n📊 Sample data (first 5 items):');
    for (var i = 0; i < Math.min(5, allItems.length); i++) {
      var item = allItems[i];
      Logger.log('\n--- Item #' + i + ' ---');

      // ลองดึงข้อมูลด้วย field names ต่างๆ
      var codField_v1 = item['07_COD'];
      var codField_v2 = item['COD'];
      var codField_v3 = item[CONFIG.FIELDS.WRP.COD];

      var trackingField_v1 = item['02_หมายเลขสิ่งของ'];
      var trackingField_v2 = item[CONFIG.FIELDS.WRP.TRACKING_NUMBER];

      Logger.log('COD attempts:');
      Logger.log('  item["07_COD"] = ' + codField_v1 + ' (type: ' + typeof codField_v1 + ')');
      Logger.log('  item["COD"] = ' + codField_v2 + ' (type: ' + typeof codField_v2 + ')');
      Logger.log('  item[CONFIG.FIELDS.WRP.COD] = ' + codField_v3 + ' (type: ' + typeof codField_v3 + ')');

      Logger.log('Tracking attempts:');
      Logger.log('  item["02_หมายเลขสิ่งของ"] = ' + trackingField_v1);
      Logger.log('  item[CONFIG.FIELDS.WRP.TRACKING_NUMBER] = ' + trackingField_v2);

      // แสดงข้อมูลเต็มของ item (ตัดให้สั้น)
      var itemStr = JSON.stringify(item);
      if (itemStr.length > 500) itemStr = itemStr.substring(0, 500) + '...';
      Logger.log('Full item: ' + itemStr);
    }

    Logger.log('\n✅ Test complete - check logs above');

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
  }
}

/**
 * ทดสอบข้อมูล WMS (Money/COD) - สำหรับ debug
 */
function testWMSData() {
  Logger.log('=== Testing WMS Data (Money/COD) ===');

  try {
    var wmsData = getFirebaseData(CONFIG.COLLECTIONS.SDIP_WMS);
    var allItems = toArray(wmsData) || [];

    Logger.log('Total WMS items: ' + allItems.length);

    if (allItems.length === 0) {
      Logger.log('❌ No WMS data found!');
      return;
    }

    // แสดง keys ทั้งหมดของ item แรก
    var firstItem = allItems[0];
    Logger.log('\n📋 Available fields in first item:');
    Logger.log(Object.keys(firstItem).join('\n'));

    // แสดงตัวอย่างข้อมูล 5 รายการแรก
    Logger.log('\n📊 Sample data (first 5 items):');
    for (var i = 0; i < Math.min(5, allItems.length); i++) {
      var item = allItems[i];
      Logger.log('\n--- Item #' + i + ' ---');

      // ลองดึงข้อมูลด้วย field names ต่างๆ
      Logger.log('Operator field attempts:');
      Logger.log('  item["05_ผู้ดำเนินการ"] = "' + item['05_ผู้ดำเนินการ'] + '"');
      Logger.log('  item["ผู้ดำเนินการ"] = "' + item['ผู้ดำเนินการ'] + '"');
      Logger.log('  item[CONFIG.FIELDS.WMS.OPERATOR] = "' + item[CONFIG.FIELDS.WMS.OPERATOR] + '"');

      Logger.log('Tracking:');
      Logger.log('  item[CONFIG.FIELDS.WMS.TRACKING_NUMBER] = "' + item[CONFIG.FIELDS.WMS.TRACKING_NUMBER] + '"');

      Logger.log('Amount:');
      Logger.log('  item[CONFIG.FIELDS.WMS.AMOUNT] = "' + item[CONFIG.FIELDS.WMS.AMOUNT] + '"');

      // แสดงข้อมูลเต็มของ item (ตัดให้สั้น)
      var itemStr = JSON.stringify(item);
      if (itemStr.length > 500) itemStr = itemStr.substring(0, 500) + '...';
      Logger.log('Full item: ' + itemStr);
    }

    Logger.log('\n✅ Test complete - check logs above');

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
  }
}

/**
 * ทดสอบ getDetailWorkData สำหรับ Money (ไม่ filter)
 */
function testMoneyDetailData() {
  Logger.log('=== Testing getDetailWorkData for Money (no filter) ===');

  try {
    // ทดสอบแบบไม่ filter
    var result = getDetailWorkData('money', null, null);

    Logger.log('Headers: ' + JSON.stringify(result.headers));
    Logger.log('Total rows: ' + result.data.length);

    if (result.error) {
      Logger.log('❌ Error in result: ' + result.error);
    }

    // แสดง 3 rows แรก
    Logger.log('\n📊 First 3 rows:');
    for (var i = 0; i < Math.min(3, result.data.length); i++) {
      Logger.log('Row ' + i + ': ' + JSON.stringify(result.data[i]));
    }

    Logger.log('\n✅ Test complete');
    return result;

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return { error: error.toString(), stack: error.stack };
  }
}

/**
 * ทดสอบ getDetailWorkData สำหรับ Money (มี filter)
 */
function testMoneyDetailDataWithFilter() {
  Logger.log('=== Testing getDetailWorkData for Money (with filter) ===');

  try {
    // ทดสอบกับ username ที่เราเห็นใน log
    var testUsername = 'visit.ko'; // จาก log ก่อนหน้า
    var result = getDetailWorkData('money', null, testUsername);

    Logger.log('Filter: ' + testUsername);
    Logger.log('Headers: ' + JSON.stringify(result.headers));
    Logger.log('Total rows: ' + result.data.length);

    if (result.error) {
      Logger.log('❌ Error in result: ' + result.error);
    }

    // แสดงทุก rows
    Logger.log('\n📊 All filtered rows:');
    for (var i = 0; i < result.data.length; i++) {
      Logger.log('Row ' + i + ': ' + JSON.stringify(result.data[i]));
    }

    Logger.log('\n✅ Test complete');
    return result;

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return { error: error.toString(), stack: error.stack };
  }
}

/**
 * ทดสอบ Login
 */
function testLogin() {
  Logger.log('=== Testing Login ===');
  var result = checkLogin('titikarn.se', '1234');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * ทดสอบดึงข้อมูลทั้งหมด
 */
function testGetAllData() {
  Logger.log('=== Testing Get All Data ===');
  try {
    var data = getAllDashboardData();
    Logger.log('Employees: ' + data.employees.length);
    Logger.log('Backlog R: ' + data.backlog.r.length);
    Logger.log('Backlog EMS: ' + data.backlog.ems.length);
    Logger.log('Backlog COD: ' + data.backlog.cod.length);
    Logger.log('Returned R: ' + data.returned.r.length);
    Logger.log('Money (WMS): ' + data.money.length);
    Logger.log('Record (WRP): ' + data.record.length);

    if (data.error) {
      Logger.log('Error: ' + data.error);
    }

    return data;
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return { error: error.toString() };
  }
}

/**
 * ตรวจสอบ Configuration
 */
function checkConfig() {
  Logger.log('=== Configuration Check ===');
  Logger.log('Firebase URL: ' + CONFIG.FIREBASE_URL);
  Logger.log('Firebase Secret Length: ' + CONFIG.FIREBASE_SECRET.length);
  Logger.log('Has Secret: ' + (CONFIG.FIREBASE_SECRET.length > 0 ? 'YES ✅' : 'NO ❌'));

  if (CONFIG.FIREBASE_SECRET.length === 0) {
    Logger.log('⚠️ WARNING: No Firebase Secret set!');
    Logger.log('Please add your Firebase Database Secret to CONFIG.FIREBASE_SECRET');
    Logger.log('Find it at: Firebase Console > Project Settings > Service Accounts > Database secrets');
  }

  return {
    hasSecret: CONFIG.FIREBASE_SECRET.length > 0,
    url: CONFIG.FIREBASE_URL
  };
}

/**
 * ดูโครงสร้างข้อมูล Firebase (Debug)
 */
function debugFirebaseStructure() {
  Logger.log('=== Firebase Database Structure ===');

  try {
    // ดูข้อมูลทั้งหมดที่ root
    var url = CONFIG.FIREBASE_URL + '/.json?shallow=true';
    if (CONFIG.FIREBASE_SECRET && CONFIG.FIREBASE_SECRET.length > 0) {
      url += '&auth=' + CONFIG.FIREBASE_SECRET;
    }

    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(response.getContentText());

    Logger.log('Collections in Firebase:');
    Logger.log(JSON.stringify(data, null, 2));

    // ดูข้อมูล Employee โดยเฉพาะ
    Logger.log('\n=== Employee Collection Details ===');
    var empData = getFirebaseData('Employee');

    if (!empData) {
      Logger.log('❌ Employee collection is NULL or does not exist');
    } else {
      Logger.log('Employee data type: ' + typeof empData);
      Logger.log('Is Array: ' + Array.isArray(empData));

      if (Array.isArray(empData)) {
        Logger.log('Employee count (Array): ' + empData.length);
        Logger.log('First employee: ' + JSON.stringify(empData[0], null, 2));
      } else if (typeof empData === 'object') {
        var keys = Object.keys(empData);
        Logger.log('Employee count (Object): ' + keys.length);
        Logger.log('Keys: ' + keys.join(', '));
        if (keys.length > 0) {
          Logger.log('First employee: ' + JSON.stringify(empData[keys[0]], null, 2));
        }
      } else {
        Logger.log('Unexpected data type: ' + typeof empData);
        Logger.log('Data: ' + JSON.stringify(empData));
      }
    }

    return {
      success: true,
      collections: data,
      employeeData: empData
    };

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดูข้อมูล Raw จาก Firebase (ไม่มีการแปลง)
 */
function debugRawData() {
  Logger.log('=== Raw Firebase Data ===');

  try {
    var url = CONFIG.FIREBASE_URL + '/Employee.json';
    if (CONFIG.FIREBASE_SECRET && CONFIG.FIREBASE_SECRET.length > 0) {
      url += '?auth=' + CONFIG.FIREBASE_SECRET;
    }

    Logger.log('URL: ' + url);

    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();
    var text = response.getContentText();

    Logger.log('Response Code: ' + code);
    Logger.log('Response Length: ' + text.length);
    Logger.log('Response (first 1000 chars): ' + text.substring(0, 1000));

    if (text && text !== 'null') {
      var data = JSON.parse(text);
      Logger.log('Parsed data type: ' + typeof data);
      Logger.log('Is Array: ' + Array.isArray(data));

      if (Array.isArray(data)) {
        Logger.log('Array length: ' + data.length);
      } else if (typeof data === 'object' && data !== null) {
        Logger.log('Object keys: ' + Object.keys(data).length);
      }
    } else {
      Logger.log('⚠️ Data is null or empty');
    }

    return {
      code: code,
      length: text.length,
      data: text
    };

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return {
      error: error.toString()
    };
  }
}
