/**
 * ExportData.gs
 * สคริปต์สำหรับ Export ข้อมูลจาก Google Sheets เพื่อให้ AI ทำความเข้าใจ
 *
 * วิธีใช้:
 * 1. Copy สคริปต์นี้ไปใส่ใน Apps Script ของ Google Sheets
 * 2. รันฟังก์ชัน exportAllSheetsData()
 * 3. Copy ผลลัพธ์จาก Logger ไปให้ AI
 */

function exportAllSheetsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let output = '=' .repeat(80) + '\n';
  output += 'GOOGLE SHEETS DATA EXPORT\n';
  output += 'Spreadsheet: ' + ss.getName() + '\n';
  output += 'Total Sheets: ' + sheets.length + '\n';
  output += '='.repeat(80) + '\n\n';

  // Loop through each sheet
  sheets.forEach(function(sheet, index) {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    output += '\n' + '='.repeat(80) + '\n';
    output += 'SHEET #' + (index + 1) + ': ' + sheetName + '\n';
    output += '='.repeat(80) + '\n';
    output += 'Rows: ' + lastRow + ' | Columns: ' + lastCol + '\n';
    output += '-'.repeat(80) + '\n\n';

    if (lastRow > 0 && lastCol > 0) {
      // Get first 20 rows or all rows if less than 20
      const rowsToRead = Math.min(lastRow, 20);
      const data = sheet.getRange(1, 1, rowsToRead, lastCol).getValues();

      // Print header
      output += 'HEADER ROW:\n';
      output += JSON.stringify(data[0]) + '\n\n';

      // Print data
      output += 'DATA (First ' + (rowsToRead - 1) + ' rows):\n';
      for (let i = 1; i < data.length; i++) {
        output += 'Row ' + (i + 1) + ': ' + JSON.stringify(data[i]) + '\n';
      }

      // Column info
      output += '\n' + '-'.repeat(80) + '\n';
      output += 'COLUMN DETAILS:\n';
      for (let i = 0; i < data[0].length; i++) {
        output += 'Column ' + String.fromCharCode(65 + i) + ' (' + (i + 1) + '): ' + data[0][i] + '\n';
      }
    } else {
      output += '(Empty sheet)\n';
    }

    output += '\n';
  });

  output += '\n' + '='.repeat(80) + '\n';
  output += 'END OF EXPORT\n';
  output += '='.repeat(80) + '\n';

  // Log output
  Logger.log(output);

  // Also show in UI
  SpreadsheetApp.getUi().alert(
    'Export Complete!\n\n' +
    'ข้อมูลถูก export แล้ว\n\n' +
    'กรุณาไปที่:\n' +
    'View > Logs (Ctrl+Enter)\n\n' +
    'แล้ว Copy ข้อมูลทั้งหมดไปให้ AI'
  );

  return output;
}

/**
 * ฟังก์ชันสำหรับ export เฉพาะชีตที่ระบุ
 */
function exportSpecificSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('ERROR: Sheet "' + sheetName + '" not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let output = '='.repeat(80) + '\n';
  output += 'SHEET: ' + sheetName + '\n';
  output += 'Rows: ' + lastRow + ' | Columns: ' + lastCol + '\n';
  output += '='.repeat(80) + '\n\n';

  if (lastRow > 0 && lastCol > 0) {
    const rowsToRead = Math.min(lastRow, 20);
    const data = sheet.getRange(1, 1, rowsToRead, lastCol).getValues();

    // Print in table format
    data.forEach(function(row, index) {
      output += 'Row ' + (index + 1) + ':\t' + row.join('\t|\t') + '\n';
    });
  }

  Logger.log(output);
  return output;
}

/**
 * ฟังก์ชันสำหรับแสดงรายชื่อชีตทั้งหมด
 */
function listAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let output = 'Spreadsheet: ' + ss.getName() + '\n';
  output += 'Total Sheets: ' + sheets.length + '\n\n';

  sheets.forEach(function(sheet, index) {
    output += (index + 1) + '. ' + sheet.getName() +
              ' (' + sheet.getLastRow() + ' rows, ' +
              sheet.getLastColumn() + ' columns)\n';
  });

  Logger.log(output);
  SpreadsheetApp.getUi().alert(output);

  return output;
}
