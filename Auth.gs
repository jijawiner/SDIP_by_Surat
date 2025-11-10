/**
 * Auth.gs
 * ระบบ Authentication และการจัดการผู้ใช้งาน
 *
 * อ่านข้อมูล Login จากชีต SDIPEmployee
 */

/**
 * ฟังก์ชันสำหรับตรวจสอบ Username และ Password
 * @param {string} username - Username ที่ต้องการตรวจสอบ
 * @param {string} password - Password ที่ต้องการตรวจสอบ
 * @return {Object} - ข้อมูลผู้ใช้ หรือ error message
 */
function authenticateUser(username, password) {
  try {
    if (!username || !password) {
      return {
        success: false,
        message: 'กรุณากรอก Username และ Password'
      };
    }

    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const dbUsername = row[EMPLOYEE_COLUMNS.USERNAME];
      const dbPassword = row[EMPLOYEE_COLUMNS.PASSWORD];
      const name = row[EMPLOYEE_COLUMNS.NAME];
      const accessLevel = row[EMPLOYEE_COLUMNS.ACCESS_LEVEL];
      const paymentSide = row[EMPLOYEE_COLUMNS.PAYMENT_SIDE];
      const userStatus = row[EMPLOYEE_COLUMNS.USER_STATUS];

      // ข้าม row ที่ไม่มี username
      if (!dbUsername) continue;

      // ตรวจสอบ username และ password
      if (dbUsername.toString().trim() === username.trim() &&
          dbPassword.toString() === password) {

        // ตรวจสอบสถานะ user (ถ้ามี)
        if (userStatus && userStatus.toString().toLowerCase() === 'inactive') {
          return {
            success: false,
            message: 'บัญชีผู้ใช้นี้ถูกระงับการใช้งาน'
          };
        }

        return {
          success: true,
          user: {
            username: dbUsername.toString().trim(),
            name: name || 'ไม่ระบุชื่อ',
            accessLevel: accessLevel || ACCESS_LEVELS.USER,
            paymentSide: paymentSide || 'ไม่ระบุโซน'
          }
        };
      }
    }

    // ไม่พบผู้ใช้หรือ password ไม่ถูกต้อง
    return {
      success: false,
      message: 'Username หรือ Password ไม่ถูกต้อง'
    };

  } catch (error) {
    Logger.log('Error in authenticateUser: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการตรวจสอบผู้ใช้: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลผู้ใช้ทั้งหมด (สำหรับ Admin เท่านั้น)
 * @return {Array} - รายการผู้ใช้ทั้งหมด
 */
function getAllUsers() {
  try {
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();

    const users = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const username = row[EMPLOYEE_COLUMNS.USERNAME];

      // ข้าม row ที่ไม่มี username
      if (!username) continue;

      users.push({
        username: username.toString().trim(),
        name: row[EMPLOYEE_COLUMNS.NAME] || 'ไม่ระบุชื่อ',
        accessLevel: row[EMPLOYEE_COLUMNS.ACCESS_LEVEL] || ACCESS_LEVELS.USER,
        paymentSide: row[EMPLOYEE_COLUMNS.PAYMENT_SIDE] || 'ไม่ระบุโซน',
        position: row[EMPLOYEE_COLUMNS.POSITION] || '',
        userStatus: row[EMPLOYEE_COLUMNS.USER_STATUS] || 'Active'
        // ไม่ส่ง password กลับไปด้วยเหตุผลด้านความปลอดภัย
      });
    }

    return users;

  } catch (error) {
    Logger.log('Error in getAllUsers: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลผู้ใช้: ' + error.message);
  }
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลพนักงานทั้งหมด (รวม field เพิ่มเติม)
 * @return {Array} - รายการพนักงานทั้งหมด
 */
function getAllEmployees() {
  try {
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();

    const employees = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const username = row[EMPLOYEE_COLUMNS.USERNAME];

      // ข้าม row ที่ไม่มี username
      if (!username) continue;

      employees.push({
        order: row[EMPLOYEE_COLUMNS.ORDER] || i,
        userStatus: row[EMPLOYEE_COLUMNS.USER_STATUS] || 'Active',
        username: username.toString().trim(),
        name: row[EMPLOYEE_COLUMNS.NAME] || 'ไม่ระบุชื่อ',
        accessLevel: row[EMPLOYEE_COLUMNS.ACCESS_LEVEL] || ACCESS_LEVELS.USER,
        position: row[EMPLOYEE_COLUMNS.POSITION] || '',
        paymentSide: row[EMPLOYEE_COLUMNS.PAYMENT_SIDE] || 'ไม่ระบุโซน',
        district: row[EMPLOYEE_COLUMNS.DISTRICT] || '',
        villageNo: row[EMPLOYEE_COLUMNS.VILLAGE_NO] || '',
        workPhone: row[EMPLOYEE_COLUMNS.WORK_PHONE] || '',
        mobilePhone: row[EMPLOYEE_COLUMNS.MOBILE_PHONE] || '',
        lineId: row[EMPLOYEE_COLUMNS.LINE_ID] || '',
        gmail: row[EMPLOYEE_COLUMNS.GMAIL] || ''
      });
    }

    return employees;

  } catch (error) {
    Logger.log('Error in getAllEmployees: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลพนักงาน: ' + error.message);
  }
}

/**
 * ฟังก์ชันสำหรับเปลี่ยน Password (สำหรับผู้ใช้ทั่วไป)
 * @param {string} username - Username
 * @param {string} oldPassword - Password เดิม
 * @param {string} newPassword - Password ใหม่
 * @return {Object} - ผลลัพธ์การเปลี่ยน password
 */
function changePassword(username, oldPassword, newPassword) {
  try {
    // ตรวจสอบ password เดิม
    const authResult = authenticateUser(username, oldPassword);

    if (!authResult.success) {
      return {
        success: false,
        message: 'Password เดิมไม่ถูกต้อง'
      };
    }

    // เปลี่ยน password
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[EMPLOYEE_COLUMNS.USERNAME].toString().trim() === username.trim()) {
        // อัพเดท password ในเซลล์
        const passwordCell = sheet.getRange(i + 1, EMPLOYEE_COLUMNS.PASSWORD + 1);
        passwordCell.setValue(newPassword);

        return {
          success: true,
          message: 'เปลี่ยน Password เรียบร้อยแล้ว'
        };
      }
    }

    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ไม่พบผู้ใช้'
    };

  } catch (error) {
    Logger.log('Error in changePassword: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการเปลี่ยน Password: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันตรวจสอบสิทธิ์การเข้าถึง
 * @param {string} username - Username
 * @param {string} requiredLevel - ระดับสิทธิ์ที่ต้องการ
 * @return {boolean} - true ถ้ามีสิทธิ์
 */
function checkAccess(username, requiredLevel) {
  try {
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[EMPLOYEE_COLUMNS.USERNAME].toString().trim() === username.trim()) {
        const userLevel = row[EMPLOYEE_COLUMNS.ACCESS_LEVEL] || ACCESS_LEVELS.USER;

        // Admin มีสิทธิ์ทุกอย่าง
        if (userLevel === ACCESS_LEVELS.ADMIN) {
          return true;
        }

        // PowerUser มีสิทธิ์รองลงมา
        if (userLevel === ACCESS_LEVELS.POWER_USER &&
            requiredLevel !== ACCESS_LEVELS.ADMIN) {
          return true;
        }

        // User ธรรมดาเข้าถึงได้เฉพาะข้อมูลตัวเอง
        if (userLevel === ACCESS_LEVELS.USER &&
            requiredLevel === ACCESS_LEVELS.USER) {
          return true;
        }

        return false;
      }
    }

    return false;

  } catch (error) {
    Logger.log('Error in checkAccess: ' + error.message);
    return false;
  }
}
