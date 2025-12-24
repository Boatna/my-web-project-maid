// Google Apps Script สำหรับระบบงานแม่บ้าน
// คัดลอกโค้ดนี้ไปวางใน Google Apps Script

// Google Sheet ID - แก้ไขเป็น ID ของ Google Sheet ของคุณ
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

// ชื่อ sheet ใน Google Sheet
const SHEET_NAMES = {
  SUBMISSIONS: 'งานแม่บ้าน',
  TASKS: 'รายการงาน',
  EMPLOYEES: 'พนักงาน'
};

// ฟังก์ชันหลักสำหรับรับข้อมูลจากเว็บแอป
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'submit';
    
    let result;
    
    switch(action) {
      case 'submit':
        result = saveSubmission(data);
        break;
      case 'get_tasks':
        result = getTasks();
        break;
      case 'get_employees':
        result = getEmployees();
        break;
      case 'get_history':
        result = getHistory(data.employeeId);
        break;
      default:
        result = { success: false, message: 'Invalid action' };
    }
    
    return createResponse(result);
  } catch (error) {
    return createResponse({ 
      success: false, 
      message: 'Error: ' + error.toString() 
    });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  let result;
  
  switch(action) {
    case 'get_tasks':
      result = getTasks();
      break;
    case 'get_employees':
      result = getEmployees();
      break;
    default:
      result = { 
        success: true, 
        message: 'Housekeeping System API is running',
        version: '1.0.0'
      };
  }
  
  return createResponse(result);
}

// บันทึกการส่งงาน
function saveSubmission(data) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.SUBMISSIONS);
    
    // ตรวจสอบว่ามีหัวคอลัมน์หรือไม่
    if (sheet.getLastRow() === 0) {
      const headers = [
        'วันที่ส่ง', 'เวลา', 'รหัสพนักงาน', 'ชื่อพนักงาน', 
        'รหัสงาน', 'ชื่องาน', 'พื้นที่', 'หมายเหตุ', 'มีรูปภาพ', 
        'สถานะ', 'ผู้ตรวจสอบ', 'หมายเหตุการตรวจสอบ'
      ];
      sheet.appendRow(headers);
    }
    
    const now = new Date();
    const dateString = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
    const timeString = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');
    
    // บันทึกข้อมูล
    const rowData = [
      dateString,
      timeString,
      data.employeeId,
      data.employeeName,
      data.taskId,
      data.taskName,
      data.area,
      data.notes || '',
      data.hasImage ? 'มี' : 'ไม่มี',
      'รอตรวจสอบ',
      '',
      ''
    ];
    
    sheet.appendRow(rowData);
    
    // หากมีรูปภาพ ให้บันทึกลง Google Drive
    let imageUrl = '';
    if (data.hasImage && data.imageData) {
      imageUrl = saveImageToDrive(data);
    }
    
    // ส่งอีเมลแจ้งเตือน (ถ้าต้องการ)
    sendNotificationEmail(data);
    
    return {
      success: true,
      message: 'บันทึกข้อมูลสำเร็จ',
      submissionId: sheet.getLastRow(),
      imageUrl: imageUrl
    };
  } catch (error) {
    return {
      success: false,
      message: 'บันทึกข้อมูลไม่สำเร็จ: ' + error.toString()
    };
  }
}

// บันทึกรูปลง Google Drive
function saveImageToDrive(data) {
  try {
    const folderName = 'Housekeeping_System_Images';
    const dateString = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
    
    // สร้างหรือหาโฟลเดอร์
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    // สร้างโฟลเดอร์ย่อยตามวันที่
    const subFolders = folder.getFoldersByName(dateString);
    let subFolder;
    
    if (subFolders.hasNext()) {
      subFolder = subFolders.next();
    } else {
      subFolder = folder.createFolder(dateString);
    }
    
    // สร้างชื่อไฟล์
    const fileName = `submission_${data.employeeId}_${data.taskId}_${Date.now()}.jpg`;
    
    // แปลง base64 เป็น blob
    const imageData = data.imageData.replace(/^data:image\/\w+;base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(imageData), 'image/jpeg', fileName);
    
    // บันทึกลง Drive
    const file = subFolder.createFile(blob);
    
    // ตั้งค่าให้ใครๆ ก็เข้าถึงได้ (หรือตั้งค่าสิทธิ์ตามต้องการ)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (error) {
    console.error('Error saving image:', error);
    return '';
  }
}

// ส่งอีเมลแจ้งเตือน
function sendNotificationEmail(data) {
  try {
    // ตั้งค่าผู้รับอีเมล
    const adminEmail = 'admin@example.com'; // แก้ไขเป็นอีเมลของผู้ดูแลระบบ
    
    const subject = `มีการส่งงานใหม่จาก ${data.employeeName}`;
    const body = `
มีการส่งงานใหม่ในระบบงานแม่บ้าน

รายละเอียด:
- พนักงาน: ${data.employeeName} (รหัส: ${data.employeeId})
- งาน: ${data.taskName}
- พื้นที่: ${data.area}
- วันที่ส่ง: ${new Date().toLocaleDateString('th-TH')}
- เวลา: ${new Date().toLocaleTimeString('th-TH')}
- หมายเหตุ: ${data.notes || 'ไม่มี'}

สามารถตรวจสอบข้อมูลเพิ่มเติมได้ที่ Google Sheet
    `;
    
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: body
    });
    
    return true;
  } catch (error) {
    console.error('Error sending email:', error);
    return false;
  }
}

// ดึงข้อมูลงาน
function getTasks() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.TASKS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: false, message: 'ไม่พบข้อมูลงาน' };
    }
    
    const headers = data[0];
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const task = {};
      
      for (let j = 0; j < headers.length; j++) {
        task[headers[j]] = row[j];
      }
      
      tasks.push(task);
    }
    
    return {
      success: true,
      tasks: tasks
    };
  } catch (error) {
    return {
      success: false,
      message: 'ดึงข้อมูลไม่สำเร็จ: ' + error.toString()
    };
  }
}

// ดึงข้อมูลพนักงาน
function getEmployees() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: false, message: 'ไม่พบข้อมูลพนักงาน' };
    }
    
    const headers = data[0];
    const employees = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employee = {};
      
      for (let j = 0; j < headers.length; j++) {
        employee[headers[j]] = row[j];
      }
      
      employees.push(employee);
    }
    
    return {
      success: true,
      employees: employees
    };
  } catch (error) {
    return {
      success: false,
      message: 'ดึงข้อมูลไม่สำเร็จ: ' + error.toString()
    };
  }
}

// ดึงประวัติการส่งงาน
function getHistory(employeeId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.SUBMISSIONS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: false, message: 'ไม่พบประวัติการส่งงาน' };
    }
    
    const headers = data[0];
    const history = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // กรองตามรหัสพนักงาน (ถ้ามีการระบุ)
      if (employeeId && row[2] != employeeId) {
        continue;
      }
      
      const record = {};
      
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = row[j];
      }
      
      history.push(record);
    }
    
    return {
      success: true,
      history: history
    };
  } catch (error) {
    return {
      success: false,
      message: 'ดึงข้อมูลไม่สำเร็จ: ' + error.toString()
    };
  }
}

// สร้างการตอบกลับ
function createResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}