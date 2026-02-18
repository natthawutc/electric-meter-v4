/**
 * ระบบจดมิเตอร์ไฟฟ้า (USER VERSION) - v5.2 EXACT MATCH VALIDATION
 * Update: Compare Record Col C/D vs Master Col B/C
 */
const APP_VERSION = "อัพเดทล่าสุด 16/02/2569 (v5.2 Check Move)";

const SPREADSHEET_ID = '1RsfsZlCJykyIOpW5aHbb5dmnjKovgw8org3_h3EcAQA';
const SHEET_METERS = 'รายชื่อมิเตอร์';
const SHEET_EMPLOYEES = 'DB_พนักงาน';
const SHEET_RECORDS = 'บันทึกเลขมิเตอร์';
const PHOTO_FOLDER_ID = '1AcSMtEpvDvNNpO-OSM1xnPvdai2GgBun'; 

function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('มิเตอร์ไฟฟ้า.USER v5.2')
    .setFaviconUrl('https://i.postimg.cc/FR0bC5RT/Logo-App-(1).png')    
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  return { version: APP_VERSION };
}

// Helper: ดึงประวัติการจดล่าสุด
function getLastReadingsMap(ss) {
  const recordSheet = ss.getSheetByName(SHEET_RECORDS);
  let lastRecords = {}; 
  if (recordSheet) {
    const recordData = recordSheet.getDataRange().getValues();
    // วนลูปจากล่างขึ้นบนเพื่อหา record ล่าสุดของแต่ละมิเตอร์
    for (let i = recordData.length - 1; i >= 1; i--) {
      let mId = String(recordData[i][1]).trim(); // Col B: รหัสมิเตอร์
      if (mId && lastRecords[mId] === undefined) {
        let rawDate = recordData[i][0];
        let dateStr = (rawDate instanceof Date) ? rawDate.toISOString() : String(rawDate);
        
        // เก็บข้อมูลประวัติเพื่อใช้ตรวจสอบ (Record Col C, Col D)
        lastRecords[mId] = {
          dateStr: dateStr,
          reading: recordData[i][4], // Col E: เลขอ่าน
          lastZone: recordData[i][3], // Col D: โซนที่จดครั้งล่าสุด
          lastName: recordData[i][2]  // Col C: ชื่อมิเตอร์ครั้งล่าสุด
        };
      }
    }
  }
  return lastRecords;
}

// ดึงมิเตอร์ตามรหัสผู้รับผิดชอบ (Col H = Index 7)
function getMetersByOwner(ownerId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const meterSheet = ss.getSheetByName(SHEET_METERS);
  if (!meterSheet) return [];
  
  const meterData = meterSheet.getDataRange().getValues();
  const lastRecords = getLastReadingsMap(ss);
  const targetOwner = String(ownerId).trim();

  // กรองเฉพาะมิเตอร์ที่เป็นของ Owner คนนี้
  return meterData.slice(1).filter(row => {
    // Col H (Owner) ต้องตรง และ Col G (Status) ต้องเป็น "ใช้งานอยู่"
    return String(row[7]).trim() === targetOwner && String(row[6]).trim() === "ใช้งานอยู่";
  }).map(row => {
    const mId = String(row[0]).trim();
    const history = lastRecords[mId] || {};
    return {
      id: mId,
      name: row[1], // Master Col B: ชื่อมิเตอร์ปัจจุบัน
      zone: row[2], // Master Col C: โซนปัจจุบัน
      lastReading: history.reading !== undefined ? history.reading : 0,
      lastRecordedDate: history.dateStr || "",
      lastZone: history.lastZone || "", // Record Col D: โซนครั้งก่อน
      lastName: history.lastName || ""  // Record Col C: ชื่อครั้งก่อน
    };
  });
}

function verifyEmployeeIdInternal(sheet, id) {
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const employeeIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
  return employeeIds.includes(String(id).trim());
}

function verifyEmployeeId(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMPLOYEES);
  return verifyEmployeeIdInternal(sheet, id);
}

function submitReadingData(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Validation
    const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
    if (!empSheet) throw new Error("ไม่พบชีทฐานข้อมูลพนักงาน");

    if (!verifyEmployeeIdInternal(empSheet, payload.ownerId)) {
      throw new Error(`ไม่พบรหัสผู้รับผิดชอบ "${payload.ownerId}" ในระบบ`);
    }

    let finalRecorderId = payload.ownerId;
    if (payload.substituteId && payload.substituteId.trim() !== "") {
      if (!verifyEmployeeIdInternal(empSheet, payload.substituteId)) {
        throw new Error(`ไม่พบรหัสผู้บันทึกแทน "${payload.substituteId}" ในระบบ`);
      }
      finalRecorderId = payload.substituteId; 
    }

    let recordSheet = ss.getSheetByName(SHEET_RECORDS);
    if (!recordSheet) {
      recordSheet = ss.insertSheet(SHEET_RECORDS);
      recordSheet.appendRow(["วันที่-เวลา", "รหัสมิเตอร์", "ชื่อมิเตอร์", "โซนที่เลือก", "เลขหน้าปัดที่อ่านได้", "รหัสผู้จด", "หมายเหตุ", "แนบรูป : URL"]);
    }

    const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy, HH:mm:ss");
    
    payload.readings.forEach(reading => {
      let photoUrl = "";
      if (reading.imageBody && reading.imageBody.includes("base64,")) {
        try {
          const splitData = reading.imageBody.split(',');
          const contentType = splitData[0].match(/:(.*?);/)[1];
          const bytes = Utilities.base64Decode(splitData[1]);
          const blob = Utilities.newBlob(bytes, contentType, `IMG_${reading.id}_${now.getTime()}.jpg`);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoUrl = file.getUrl();
        } catch (imgErr) {
          photoUrl = "Error: " + imgErr.message;
        }
      }

      recordSheet.appendRow([
        formattedDate,
        reading.id,
        reading.name,
        reading.zone || "-", 
        reading.val,
        finalRecorderId, 
        reading.note || "",
        photoUrl
      ]);
    });

    SpreadsheetApp.flush();
    return "บันทึกสำเร็จ " + payload.readings.length + " รายการ";
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}