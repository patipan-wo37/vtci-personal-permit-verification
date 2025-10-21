function doGet(e) {
  try {
    const employeeId = e.parameter.id;
    
    if (!employeeId) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'ไม่พบรหัสบัตรเจ้าหน้าที่'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // เปิด Google Sheets ด้วย ID ที่กำหนด
    const spreadsheet = SpreadsheetApp.openById('1yUDgJ-4aPbMdntYsopCNjNn12FBSUDw3ZQl1shPyQYk');
    const sheet = spreadsheet.getSheetByName('บัตรเจ้าหน้าที่');
    
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'ไม่พบชีต "บัตรเจ้าหน้าที่"'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // ดึงข้อมูลทั้งหมด
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'ไม่มีข้อมูลในชีต'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // หาข้อมูลเจ้าหน้าที่ (ตรวจสอบในคอลัมน์ B = index 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentId = row[1] ? row[1].toString().trim() : ''; // คอลัมน์ B
      
      if (currentId === employeeId.trim()) {
        const employee = {
          id: row[1] || '', // คอลัมน์ B - รหัสบัตรเจ้าหน้าที่
          name: row[2] || '', // คอลัมน์ C - ชื่อ-นามสกุล
          agency: row[3] || '', // คอลัมน์ D - หน่วยงาน
          department: row[4] || '', // คอลัมน์ E - แผนก
          position: row[5] || '', // คอลัมน์ F - ตำแหน่ง
          areaNumber: row[9] || '', // คอลัมน์ J - หมายเลขพื้นที่
          expiryDate: row[10] || '', // คอลัมน์ K - วันที่บัตรหมดอายุ
          cardType: row[11] || '', // คอลัมน์ L - ประเภทบัตร
          issuingAgency: row[12] || '', // คอลัมน์ M - หน่วยงานที่ออกบัตร
          additionalInfo: row[13] || '', // คอลัมน์ N - ข้อมูลบัตรเพิ่มเติม
          remarks: row[14] || '', // คอลัมน์ O - หมายเหตุ
          photo: row[15] || '', // คอลัมน์ P - รูปภาพ
          pdfRequest: row[16] || '' // คอลัมน์ Q - คำขอทำบัตรอนุญาตฯ
        };
        
        return ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            employee: employee
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ไม่พบข้อมูล
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'ไม่พบข้อมูลเจ้าหน้าที่ที่มีรหัสบัตร: ' + employeeId
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'เกิดข้อผิดพลาดในระบบ: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}