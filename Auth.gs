/**
 * Auth.gs
 * ระบบ Authentication และการจัดการผู้ใช้งาน
 */

/**
 * ฟังก์ชันสำหรับตรวจสอบ Username และ Password
 * @param {string} username - Username ที่ต้องการตรวจสอบ
 * @param {string} password - Password ที่ต้องการตรวจสอบ
 * @return {Object} - ข้อมูลผู้ใช้ หรือ null ถ้าไม่ถูกต้อง
 */
function authenticateUser(username, password) {
  try {
    if (!username || !password) {
      return {
        success: false,
        message: 'กรุณากรอก Username และ Password'
      };
    }

    const sheet = getSheet(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dbUsername = row[USER_COLUMNS.USERNAME];
      const dbPassword = row[USER_COLUMNS.PASSWORD];
      const name = row[USER_COLUMNS.NAME];
      const accessLevel = row[USER_COLUMNS.ACCESS_LEVEL];

      // ตรวจสอบ username และ password
      if (dbUsername === username && dbPassword === password) {
        return {
          success: true,
          user: {
            username: dbUsername,
            name: name,
            accessLevel: accessLevel
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
    const sheet = getSheet(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    const users = [];

    // ข้าม header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // ข้าม row ที่ไม่มี username
      if (!row[USER_COLUMNS.USERNAME]) continue;

      users.push({
        username: row[USER_COLUMNS.USERNAME],
        name: row[USER_COLUMNS.NAME],
        accessLevel: row[USER_COLUMNS.ACCESS_LEVEL]
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
 * ฟังก์ชันสำหรับเพิ่มผู้ใช้ใหม่ (สำหรับ Admin เท่านั้น)
 * @param {Object} userData - ข้อมูลผู้ใช้ใหม่
 * @return {Object} - ผลลัพธ์การเพิ่มผู้ใช้
 */
function addUser(userData) {
  try {
    // ตรวจสอบข้อมูลที่จำเป็น
    if (!userData.username || !userData.password || !userData.name || !userData.accessLevel) {
      return {
        success: false,
        message: 'กรุณากรอกข้อมูลให้ครบถ้วน'
      };
    }

    // ตรวจสอบว่า username ซ้ำหรือไม่
    const sheet = getSheet(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLUMNS.USERNAME] === userData.username) {
        return {
          success: false,
          message: 'Username นี้มีอยู่ในระบบแล้ว'
        };
      }
    }

    // เพิ่มผู้ใช้ใหม่
    sheet.appendRow([
      userData.username,
      userData.password,
      userData.name,
      userData.accessLevel
    ]);

    return {
      success: true,
      message: 'เพิ่มผู้ใช้เรียบร้อยแล้ว'
    };

  } catch (error) {
    Logger.log('Error in addUser: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการเพิ่มผู้ใช้: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันสำหรับลบผู้ใช้ (สำหรับ Admin เท่านั้น)
 * @param {string} username - Username ที่ต้องการลบ
 * @return {Object} - ผลลัพธ์การลบผู้ใช้
 */
function deleteUser(username) {
  try {
    // ป้องกันการลบ admin
    if (username === 'admin') {
      return {
        success: false,
        message: 'ไม่สามารถลบผู้ใช้ admin ได้'
      };
    }

    const sheet = getSheet(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    // หาแถวที่ต้องการลบ
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLUMNS.USERNAME] === username) {
        sheet.deleteRow(i + 1); // +1 เพราะ sheet row เริ่มที่ 1
        return {
          success: true,
          message: 'ลบผู้ใช้เรียบร้อยแล้ว'
        };
      }
    }

    return {
      success: false,
      message: 'ไม่พบผู้ใช้ที่ต้องการลบ'
    };

  } catch (error) {
    Logger.log('Error in deleteUser: ' + error.message);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการลบผู้ใช้: ' + error.message
    };
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
    const sheet = getSheet(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLUMNS.USERNAME] === username) {
        sheet.getRange(i + 1, USER_COLUMNS.PASSWORD + 1).setValue(newPassword);
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
