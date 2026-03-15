/**
 * 🏫 PSSMS - Phuphrabat Smart School Management System
 * ระบบบริหารจัดการสถานศึกษา 4 ฝ่าย (Single Page Application)
 * พัฒนาโดย: ครูน๊อต ศิกษก เดินรีบรัมย์
 * Updated: 2026-02-26 | ปรับปรุงโครงสร้างฐานข้อมูล (Level, Room, Location)
 */

// ==========================================
// 1. CORE FUNCTIONS (ระบบหลักของเว็บ)
// ==========================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('PSSMS - โรงเรียนภูพระบาทวิทยา')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPage(pageName) {
  try {
    return HtmlService.createTemplateFromFile(pageName).evaluate().getContent();
  } catch (e) {
    return "ไม่พบหน้าเว็บ: " + pageName;
  }
}

// 🛡️ ระบบรักษาความปลอดภัย: ตรวจสอบสิทธิ์ครูผู้สอน
function verifyTeacherPermission(teacherId, subjectCode, className, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. สิทธิพิเศษ: ถ้าเป็น Admin ให้ผ่านได้เลยทุกกรณี
  const userSheet = ss.getSheetByName("User_Database");
  if (userSheet) {
    const users = userSheet.getDataRange().getValues();
    const userRow = users.find(r => String(r[0]).trim() === String(teacherId).trim());
    if (userRow && String(userRow[3]).toUpperCase() === 'ADMIN') return true;
  }

  // 2. เช็คจากตารางสอน (Timetable)
  const timeSheet = ss.getSheetByName("Timetable_Database");
  if (!timeSheet) return false;

  const timeData = timeSheet.getDataRange().getDisplayValues();
  const sSub = String(subjectCode).trim().toLowerCase();
  const sClass = String(className).trim().replace(/\s/g, '').toLowerCase();
  const sTeacher = String(teacherId).trim().toLowerCase();
  const sTerm = String(term).trim();
  const sYear = String(year).trim();

  for (let i = 1; i < timeData.length; i++) {
    const row = timeData[i];
    const tCode = String(row[0]).trim().toLowerCase();
    const tName = String(row[1]).trim().toLowerCase();
    const tClassID = String(`${row[2]}/${row[3]}`).trim().replace(/\s/g, '').toLowerCase();
    const tTeacher = String(row[5]).trim().toLowerCase();
    
    // อนุโลมให้วิชาโฮมรูม (HR)
    const isHR = (tCode === 'hr' || tName.includes('โฮมรูม'));
    const isTargetSub = (tCode === sSub) || (sSub === 'hr' && isHR);

    // ถ้าวิชาตรง ห้องตรง เทอมตรง ปีตรง
    if (isTargetSub && tClassID === sClass && String(row[8]).trim() === sTerm && String(row[9]).trim() === sYear) {
      // เช็คว่ารหัสครูตรงกันไหม
      if (tTeacher === sTeacher || tTeacher.includes(sTeacher) || sTeacher.includes(tTeacher)) {
        return true; 
      }
    }
  }
  
  // ถ้าหาจนจบแล้วไม่เจอชื่อครูคนนี้สอนวิชานี้ = แอบอ้าง!
  return false; 
}

// ==========================================
// 2. AUTHENTICATION & CONFIG (ยืนยันตัวตน + ตั้งค่าระบบ)
// ==========================================

function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูลผู้ใช้" };
  
  const config = getSystemConfig();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(username) && String(data[i][1]) === String(password)) {
      return {
        status: "success",
        role: data[i][3], 
        name: data[i][2], 
        id: data[i][0],   
        dept: data[i][4], 
        currentTerm: config.term,
        currentYear: config.year 
      };
    }
  }
  return { status: "fail", message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

function getSystemConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("System_Settings");
  let config = { term: "1", year: "2568", termStart: "", termEnd: "", termHistory: {} };
  
  if (!sheet) return config;
  
  const data = sheet.getDataRange().getValues();
  
  // ตรวจสอบว่าเป็นโครงสร้างเก่าหรือไม่
  const isOldFormat = data.some(r => r[0] === "Current_Term");
  
  if (isOldFormat) {
    // ใช้วิธีอ่านแบบเก่าไปก่อน
    data.forEach(row => {
      if(row[0] === "Current_Term") config.term = String(row[1]);
      if(row[0] === "Current_Year") config.year = String(row[1]);
      if(row[0] === "Term_Start") config.termStart = String(row[1]);
      if(row[0] === "Term_End") config.termEnd = String(row[1]);
    });
    // แอบจำข้อมูลเก่าไว้ใน History ด้วย
    config.termHistory[`${config.term}_${config.year}`] = { start: config.termStart, end: config.termEnd };
  } else {
    // วิธีอ่านแบบใหม่ (ระบบ Pro)
    data.forEach(row => {
      if (row[0] === "Active" && row[1] === "Term") {
        config.term = String(row[2]);
        config.year = String(row[3]);
      } else if (row[0] === "TermData") {
        const termKey = String(row[1]); // เช่น "1_2568"
        config.termHistory[termKey] = {
          start: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          end: row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "yyyy-MM-dd") : ""
        };
      }
    });
    // ดึงวันที่ของเทอมปัจจุบันมาโชว์
    const currentKey = `${config.term}_${config.year}`;
    if (config.termHistory[currentKey]) {
      config.termStart = config.termHistory[currentKey].start;
      config.termEnd = config.termHistory[currentKey].end;
    }
  }
  
  return config;
}

function saveSystemConfig(term, year, startDate, endDate) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("System_Settings");
    if (!sheet) sheet = ss.insertSheet("System_Settings");

    const data = sheet.getDataRange().getValues();
    const targetKey = `${term}_${year}`;

    // 1. ล้างข้อมูลโครงสร้างแบบเก่าทิ้ง (Migration)
    for (let i = data.length - 1; i >= 0; i--) {
       if (["Current_Term", "Current_Year", "Term_Start", "Term_End", "Key"].includes(data[i][0])) {
           sheet.deleteRow(i + 1);
       }
    }

    // 2. ดึงข้อมูลใหม่หลังจากล้างของเก่า
    const newData = sheet.getDataRange().getValues();
    let activeUpdated = false;
    let termDataUpdated = false;

    // 3. อัปเดตข้อมูลแบบแยกหมวดหมู่
    for (let i = 0; i < newData.length; i++) {
       if (newData[i][0] === "Active" && newData[i][1] === "Term") {
           sheet.getRange(i + 1, 3, 1, 2).setValues([[term, year]]);
           activeUpdated = true;
       }
       if (newData[i][0] === "TermData" && newData[i][1] === targetKey) {
           sheet.getRange(i + 1, 3, 1, 2).setValues([[startDate, endDate]]);
           termDataUpdated = true;
       }
    }

    // 4. ถ้าไม่มีข้อมูลให้เพิ่มแถวใหม่
    if (!activeUpdated) sheet.appendRow(["Active", "Term", term, year]);
    if (!termDataUpdated) sheet.appendRow(["TermData", targetKey, startDate, endDate]);

    return { status: 'success', message: `✅ บันทึกและตั้งเป็นภาคเรียนปัจจุบัน (${term}/${year}) เรียบร้อย` };
  } catch(e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. DASHBOARD & STATS
// ==========================================

// ==========================================
// ⚙️ การตั้งค่าระบบ (Global Configuration)
// ==========================================
const NOTION_TOKEN = 'ntn_K30250483172wMxPDJaiHUmHF5DRmU3aNj7y5RuglMk6iq'; 
const DATABASE_ID = '1b4a3504c04d48c182068d064c38d1e1'; 
const PROJECT_ID = '1920b44e-92fd-8013-828a-c06028c1c231'; // โรงเรียนภูพระบาทวิทยา

// ==========================================
// 🌐 1. ฟังก์ชันรับคำสั่งจาก Web App (PSSMS)
// ==========================================
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  
  if (data.action === "create") {
    const responseText = sendTaskToNotion(data.taskName);
    const result = JSON.parse(responseText); // แปลงเป็น Object เพื่อดึง ID
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "id": result.id // ส่ง ID กลับไปให้หน้าเว็บจำไว้
    })).setMimeType(ContentService.MimeType.JSON);
  } 
  
  else if (data.action === "update") {
    // รับค่า isDone มาด้วยเพื่อให้ติ๊กเข้า-ออกได้
    updateTaskStatus(data.pageId, data.isDone);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
           .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 🚀 2. ฟังก์ชันสร้างรายการใหม่ (เรียกจาก Scripts.html)
// ==========================================
function sendTaskToNotion(taskName) {
  const url = 'https://api.notion.com/v1/pages';
  const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

  const payload = {
    "parent": { "database_id": DATABASE_ID },
    "icon": { "type": "emoji", "emoji": "✏️" },
    "properties": {
      "Name": { "title": [{ "text": { "content": taskName } }] },
      "Date": { "date": { "start": today } },
      "Status": { "status": { "name": "Not started" } },
      "Projects": { "relation": [{ "id": PROJECT_ID }] }
    }
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + NOTION_TOKEN,
      "Notion-Version": "2022-06-28", 
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  return response.getContentText(); // ส่ง JSON กลับไปให้ successHandler
}

function testNotionDirectly() {
  Logger.log("กำลังทดสอบส่งข้อมูลเข้า Notion...");
  const responseText = sendTaskToNotion("เทสการผูกโปรเจกต์ 🚀");
  Logger.log("คำตอบจาก Notion: " + responseText);
}

// ==========================================
// 🔄 3. ฟังก์ชันอัปเดตสถานะ (เรียกจาก Scripts.html)
// ==========================================
function updateTaskStatus(pageId, isDone) {
  const url = 'https://api.notion.com/v1/pages/' + pageId;
  const statusName = (isDone === true || isDone === undefined) ? "Done" : "Not started";

  const payload = {
    "properties": {
      "Status": { "status": { "name": statusName } },
      "Archive": { "checkbox": (isDone === true || isDone === undefined) }
    }
  };

  const options = {
    "method": "patch",
    "headers": {
      "Authorization": "Bearer " + NOTION_TOKEN,
      "Notion-Version": "2022-06-28", 
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  UrlFetchApp.fetch(url, options);
}

function findRealProjectID() {
  const NOTION_TOKEN = 'ntn_K30250483172wMxPDJaiHUmHF5DRmU3aNj7y5RuglMk6iq'; 
  // ID ของงาน "เทสการผูกโปรเจกต์ 🚀" ที่เพิ่งสร้าง
  const pageId = '30d0b44e-92fd-814a-8d01-cfe39c57da45'; 
  
  const url = 'https://api.notion.com/v1/pages/' + pageId;
  const options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + NOTION_TOKEN,
      "Notion-Version": "2022-06-28"
    },
    "muteHttpExceptions": true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  // ให้ระบบเจาะจงปริ้นท์แค่รหัส Project ออกมา
  if(data.properties && data.properties.Projects && data.properties.Projects.relation.length > 0) {
    Logger.log("เจอตัวการแล้ว! รหัสที่แท้จริงคือ: " + data.properties.Projects.relation[0].id);
  } else {
    Logger.log("อ้าว.. ใน Notion ยังไม่ได้เลือกโปรเจกต์ครับ ลองกลับไปเลือกด้วยมือก่อนนะ");
  }
}

function finalTestNotion() {
  const NOTION_TOKEN = 'ntn_K30250483172wMxPDJaiHUmHF5DRmU3aNj7y5RuglMk6iq'; 
  const DATABASE_ID = '1b4a3504c04d48c182068d064c38d1e1'; // ตาราง To-Do
  
  // รหัสที่มีขีดกลาง (ที่ถูกต้อง 100% จากการล้วงความลับ)
  const PROJECT_ID = '1920b44e-92fd-8013-828a-c06028c1c231'; 
  
  const url = 'https://api.notion.com/v1/pages';
  const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

  const payload = {
    "parent": { "database_id": DATABASE_ID },
    "icon": { "type": "emoji", "emoji": "✅" },
    "properties": {
      "Name": { "title": [{ "text": { "content": "เทสผูกโปรเจกต์ ดาบสุดท้าย! ⚔️" } }] },
      "Date": { "date": { "start": today } },
      "Status": { "status": { "name": "Not started" } },
      "Projects": { "relation": [{ "id": PROJECT_ID }] }
    }
  };
  
  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + NOTION_TOKEN,
      "Notion-Version": "2022-06-28", 
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  Logger.log("คำตอบจาก Notion: " + response.getContentText());
}

// ฟังก์ชันจัดการข้อมูล To-Do

function getTodoList(userId) {
  if (!userId) return "[]";
  try {
    const rawData = PropertiesService.getScriptProperties().getProperty('TODO_' + userId);
    return rawData ? rawData : "[]";
  } catch(e) { 
    return "[]"; 
  }
}

function saveTodoList(userId, todosJSON) {
  if (!userId || !todosJSON) return false;
  try { 
    PropertiesService.getScriptProperties().setProperty('TODO_' + userId, todosJSON); 
    return true; 
  } catch(e) { return false; }
}

function getAdminStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSystemConfig(); 
  const currentYear = String(config.year);
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const attSheet = ss.getSheetByName("Attendance_Database");
  let presenceCount = 0;
  if (attSheet && attSheet.getLastRow() > 1) {
    const attData = attSheet.getDataRange().getValues();
    presenceCount = attData.filter(r => {
      // r[1] คือวันที่, r[2] คือเทอม, r[3] คือปี
      const isToday = Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), "yyyy-MM-dd") === todayStr;
      const isTermMatch = String(r[2]) === String(config.term);
      const isYearMatch = String(r[3]) === String(config.year);
      return isToday && isTermMatch && isYearMatch && String(r[10]) === "มา";
    }).length;
  }
  
  const budgetSheet = ss.getSheetByName("Budgets");
  let budgetPercent = 0;
  if (budgetSheet && budgetSheet.getLastRow() > 1) {
    const bData = budgetSheet.getDataRange().getValues().slice(1);
    let total = 0, used = 0;
    bData.filter(r => String(r[6]) === currentYear).forEach(r => { total += Number(r[2]); used += Number(r[3]); });
    budgetPercent = total > 0 ? Math.round((used / total) * 100) : 0;
  }

  return { academic: presenceCount > 0 ? 100 : 0, budget: budgetPercent, personnel: 0, general: 0 };
}

function getStudentSummaryStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("User_Database");
    if (!sheet) return [];
    const config = getSystemConfig();
    const data = sheet.getDataRange().getValues().slice(1);
    let summary = {}; 
    data.forEach(row => {
      if (String(row[3]) === 'Student' && String(row[6]) === config.year) {
        let grade = String(row[4]).split('/')[0] || "ไม่ระบุ";
        if (!summary[grade]) summary[grade] = { male: 0, female: 0, total: 0 };
        if (/^(นาย|ด\.ช\.|ดช\.|เด็กชาย)/.test(row[2])) summary[grade].male++;
        else summary[grade].female++;
        summary[grade].total++;
      }
    });
    return Object.keys(summary).map(g => ({ grade: g, ...summary[g] })).sort((a,b) => a.grade.localeCompare(b.grade, 'th'));
  } catch (e) { return []; }
}

// ==========================================
// 4. USER MANAGEMENT (CRUD + CSV)
// ==========================================

function getAllUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  if (!sheet) return [];
  const config = getSystemConfig();
  const data = sheet.getDataRange().getValues().slice(1);
  return data.filter(r => r[3] !== 'Student' || String(r[6]) === config.year);
}

function addUser(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const config = getSystemConfig();
  // บันทึกสถานะลงคอลัมน์ที่ 8 (H)
  sheet.appendRow([form.username, form.password, form.fullname, form.role, form.dept, form.email, config.year, form.status || "ปกติ"]);
  return {status: 'success', message: 'เพิ่มผู้ใช้งานสำเร็จ'};
}

function editUser(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(form.username)){
      // อัปเดตข้อมูลรวมถึงสถานะ (คอลัมน์ 2-8)
      sheet.getRange(i+1, 2, 1, 7).setValues([[form.password, form.fullname, form.role, form.dept, form.email, data[i][6], form.status]]);
      return {status: 'success', message: 'แก้ไขสำเร็จ'};
    }
  }
  return {status: 'fail', message: 'ไม่พบผู้ใช้'};
}

function deleteUser(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const data = sheet.getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(username)){ sheet.deleteRow(i+1); return {status: 'success', message: 'ลบสำเร็จ'}; }
  }
  return {status: 'fail', message: 'ไม่พบข้อมูล'};
}

function importStudentCSV(base64Data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const config = getSystemConfig();
  const decoded = Utilities.base64Decode(base64Data);
  const csv = Utilities.parseCsv(Utilities.newBlob(decoded).getDataAsString('UTF-8'));
  const exist = sheet.getDataRange().getValues().map(r => String(r[0]));
  let news = [];
  for (let i = 2; i < csv.length; i++) {
    let id = String(csv[i][5]).trim();
    if (!id || exist.includes(id)) continue;
    news.push(["'" + id, "'" + csv[i][2], `${csv[i][6]}${csv[i][7]} ${csv[i][8]}`, "Student", `ม.${csv[i][3]}/${csv[i][4]}`, "-", config.year]);
  }
  if (news.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, news.length, 7).setValues(news);
  return { status: 'success', message: `นำเข้าสำเร็จ ${news.length} รายการ` };
}

function importTeacherCSV(base64Data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const decoded = Utilities.base64Decode(base64Data);
  const csv = Utilities.parseCsv(Utilities.newBlob(decoded).getDataAsString('UTF-8'));
  const exist = sheet.getDataRange().getValues().map(r => String(r[0]));
  let news = [];
  for (let i = 1; i < csv.length; i++) {
    if (!csv[i][0] || exist.includes(csv[i][0])) continue;
    news.push(["'" + csv[i][0], "teacher1234", csv[i][1], "Teacher", csv[i][2], "-", ""]);
  }
  if (news.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, news.length, 7).setValues(news);
  return { status: 'success', message: `นำเข้าสำเร็จ ${news.length} ท่าน` };
}

// ==========================================
// 5. ACADEMIC & ATTENDANCE (งานวิชาการ + เช็คชื่อ)
// ==========================================

/**
 * 🚨 ดึงข้อมูลนักเรียนกลุ่มเสี่ยง "รวมทุกวิชา" สำหรับ Teacher Dashboard
 */
function getTeacherAtRiskDashboard(teacherId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName("Attendance_Database");
  const timeSheet = ss.getSheetByName("Timetable_Database");
  
  const attData = attSheet ? attSheet.getDataRange().getDisplayValues() : [];
  const timeData = timeSheet ? timeSheet.getDataRange().getDisplayValues() : [];

  const normalize = (str) => String(str || "").replace(/[^a-zA-Z0-9ก-๙]/g, '');
  const searchTeacherId = String(teacherId).trim().toLowerCase();
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();

  // --- 1. หาก่อนว่าครูคนนี้สอนวิชาอะไร ห้องไหนบ้าง ---
  const teacherClasses = {}; 
  for (let i = 1; i < timeData.length; i++) {
    const row = timeData[i];
    const tTeacherID = String(row[5]).trim().toLowerCase();
    const tTerm = String(row[8]).trim();
    const tYear = String(row[9]).trim();

    if (tTeacherID === searchTeacherId && tTerm === targetTerm && tYear === targetYear) {
      const tCode = normalize(row[0]);
      const tLevel = String(row[2]).trim();
      const tRoom = String(row[3]).trim();
      const tClassID = normalize(`${tLevel}/${tRoom}`);
      const key = `${tCode}_${tClassID}`; // สร้าง Key เฉพาะวิชา+ห้อง

      if (!teacherClasses[key]) {
        teacherClasses[key] = {
          rawCode: row[0],
          rawName: row[1],
          rawClassID: `${tLevel}/${tRoom}`,
          periodsPerWeek: 0,
          sessions: new Set(),
          students: {}
        };
      }
      teacherClasses[key].periodsPerWeek++;
    }
  }

  // --- 2. กวาดข้อมูลเช็คชื่อเฉพาะวิชาที่ครูคนนี้สอน ---
  for (let i = 1; i < attData.length; i++) {
    const row = attData[i];
    if (!row[1]) continue;
    
    const rowTerm = String(row[2]).trim();
    const rowYear = String(row[3]).trim();
    if(rowTerm !== targetTerm || rowYear !== targetYear) continue;

    const rowSub = normalize(row[4]);
    const rowClass = normalize(row[6]);
    const key = `${rowSub}_${rowClass}`;

    // ถ้าเจอข้อมูลตรงกับวิชาที่ครูสอน ให้เก็บสถานะ
    if (teacherClasses[key]) {
      const stdID = String(row[8]).trim();
      const stdName = row[9];
      const status = row[10];
      const sessionID = String(row[12]).trim() || (row[1] + "_" + row[7]);

      teacherClasses[key].sessions.add(sessionID);

      if (!teacherClasses[key].students[stdID]) {
        teacherClasses[key].students[stdID] = { name: stdName, records: {} };
      }
      // บันทึกสถานะล่าสุดของคาบนั้น
      teacherClasses[key].students[stdID].records[sessionID] = status;
    }
  }

  // --- 3. คัดกรองและแบ่งกลุ่มนักเรียน ---
  const weeksPerTerm = 20;
  let critical = []; // < 60%
  let ms = [];       // 60-79.99%
  let risk = [];     // 80-85%

  for (const key in teacherClasses) {
    const cls = teacherClasses[key];
    const currentTotalTaught = cls.sessions.size;
    
    if (currentTotalTaught === 0) continue; // ข้ามวิชาที่ยังไม่ได้สอน

    const actualPeriodsPerWeek = cls.periodsPerWeek > 0 ? cls.periodsPerWeek : 3;
    const totalCoursePeriods = actualPeriodsPerWeek * weeksPerTerm;

    for (const stdID in cls.students) {
      const student = cls.students[stdID];
      let present = 0, late = 0, leave = 0, absent = 0;

      for (const sess in student.records) {
        const s = student.records[sess];
        if (s === 'มา') present++;
        else if (s === 'สาย') late++;
        else if (s === 'ลา') leave++;
        else if (s === 'ขาด') absent++;
      }

      // ใช้สูตร 100% ลดหลั่นลงมา
      const totalMissed = absent + leave;
      const percent = ((totalCoursePeriods - totalMissed) / totalCoursePeriods) * 100;

      if (percent <= 85) {
        const studentData = {
          id: stdID,
          name: student.name,
          subjectCode: cls.rawCode,
          subjectName: cls.rawName,
          className: cls.rawClassID,
          present, late, leave, absent,
          percent: percent.toFixed(2),
          taught: currentTotalTaught
        };

        if (percent < 60) critical.push(studentData);
        else if (percent < 80) ms.push(studentData);
        else risk.push(studentData);
      }
    }
  }

  // เรียงลำดับจาก % น้อยไปมาก (วิกฤตสุดขึ้นก่อน)
  critical.sort((a, b) => parseFloat(a.percent) - parseFloat(b.percent));
  ms.sort((a, b) => parseFloat(a.percent) - parseFloat(b.percent));
  risk.sort((a, b) => parseFloat(a.percent) - parseFloat(b.percent));

  return { critical, ms, risk };
}

/**
 * 🏫 ดึงตารางสอนครูตาม "วันที่เลือก" (รองรับการเช็คชื่อย้อนหลัง)
 */
function getTeacherTimetableByDate(teacherId, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  const config = getSystemConfig();
  const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  
  // แปลงวันที่ที่ส่งมา ให้กลายเป็นชื่อวัน (จันทร์ - ศุกร์)
  let targetDateObj = dateStr ? new Date(dateStr) : new Date();
  const targetDayName = days[targetDateObj.getDay()]; 
  
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  
  const searchTeacherId = String(teacherId).trim().toLowerCase();
  const searchTerm = String(config.term).trim();
  const searchYear = String(config.year).trim();

  return data.slice(1).map(r => {
      const tTeacherID = String(r[5]).trim().toLowerCase();
      const tDay = String(r[6]).trim(); 
      const tTerm = String(r[8]).trim();
      const tYear = String(r[9]).trim();
      
      // เทียบกับ targetDayName
      if (tTeacherID === searchTeacherId && tDay === targetDayName && tTerm === searchTerm && tYear === searchYear) {
         const tLevel = String(r[2]).trim();
         const tRoom = String(r[3]).trim();
         const tLoc = String(r[4]).trim();
         const tClassID = `${tLevel}/${tRoom}`; 
         return [r[0], r[1], tClassID, tRoom, tLoc, r[7], r[6]]; 
      }
      return null;
  }).filter(item => item !== null);
}


/**
 * 🏫 ดึงตารางสอนครูรายวัน (Updated for New Schema)
 * Returns: [Code, Name, Class(ประกอบร่าง), Room, Location, Period, Day]
 */
function getTeacherTimetable(teacherId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  const config = getSystemConfig();
  const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  const today = days[new Date().getDay()];
  
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  
  const searchTeacherId = String(teacherId).trim().toLowerCase();
  const searchTerm = String(config.term).trim();
  const searchYear = String(config.year).trim();

  // New Schema: 
  // 0:Code, 1:Name, 2:Level, 3:Room, 4:Location, 5:Teacher, 6:Day, 7:Period, 8:Term, 9:Year
  
  return data.slice(1).map(r => {
      // Map ให้เป็นรูปแบบที่เข้าใจง่าย
      const tTeacherID = String(r[5]).trim().toLowerCase();
      const tDay = String(r[6]).trim();
      const tTerm = String(r[8]).trim();
      const tYear = String(r[9]).trim();
      
      if (tTeacherID === searchTeacherId && tDay === today && tTerm === searchTerm && tYear === searchYear) {
         const tLevel = String(r[2]).trim();
         const tRoom = String(r[3]).trim();
         const tLoc = String(r[4]).trim();
         const tClassID = `${tLevel}/${tRoom}`; // ประกอบร่าง: ม.5/1

         // Return format ที่หน้าบ้านคาดหวัง: [Code, Name, ClassID, Room, Location, Period]
         // Index 2 เดิมคือ ClassString ตอนนี้ส่ง ClassID แทน
         // ส่งข้อมูลย่อยไปด้วยเผื่อใช้
         return [r[0], r[1], tClassID, tRoom, tLoc, r[7], r[6]]; 
      }
      return null;
  }).filter(item => item !== null);
}


function getStudentsByClass(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const config = getSystemConfig();
  if (!sheet) return [];

  // 🌟 ใช้ getDisplayValues เพื่อให้ดึงข้อมูลออกมาเป็นตัวอักษร 100% กันปัญหาประเภทข้อมูลไม่ตรง
  const data = sheet.getDataRange().getDisplayValues();
  
  // 🌟 ลบช่องว่าง (Space) ออกทั้งหมด เช่น "ม. 1 / 2" จะกลายเป็น "ม.1/2" ทันที
  const targetClass = String(className).replace(/\s+/g, ''); 
  const targetYear = String(config.year).trim(); 

  const filtered = data.slice(1).filter(r => {
    const rowRole = String(r[3]).trim().toLowerCase();
    const rowClass = String(r[4]).replace(/\s+/g, ''); // ลบช่องว่างในฐานข้อมูลก่อนเทียบ
    const rowYear = String(r[6]).trim();
    
    // เช็คสถานะแขวนลอย (คอลัมน์ H) ถ้าว่างให้ถือว่า 'ปกติ'
    let rowStatus = "ปกติ";
    if (r.length > 7 && String(r[7]).trim() !== "") {
      rowStatus = String(r[7]).trim();
    }

    // 🛡️ เช็คเงื่อนไข (ยืดหยุ่นขึ้น)
    const isStudent = (rowRole === 'student' || rowRole === 'นักเรียน');
    const isClassMatch = (rowClass === targetClass);
    // ปีการศึกษาต้องตรง หรือถ้าครูไม่ได้ใส่ปีการศึกษาให้เด็ก (เป็นช่องว่าง) ก็ให้ดึงมาโชว์ด้วย
    const isYearMatch = (rowYear === targetYear || rowYear === ""); 
    const isStatusNormal = (rowStatus === 'ปกติ');

    // ถ้าผ่านด่านทั้งหมด ถึงจะเอารายชื่อมาแสดง
    return isStudent && isClassMatch && isYearMatch && isStatusNormal;
  });

  return filtered;
}

function updateAttendanceStatus(studentId, sessionID, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };

  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8]).trim() === String(studentId).trim() && 
        String(data[i][12]).trim() === String(sessionID).trim()) {
      
      sheet.getRange(i + 1, 11).setValue(newStatus);
      return { status: "success", message: "อัปเดตสถานะเรียบร้อย" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลที่ต้องการแก้ไข" };
}

function getStudentAttendanceHistory(studentId, subjectCode, className) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance_Database");
  const data = sheet.getDataRange().getValues();
  const history = [];

  data.slice(1).forEach(row => {
    if (String(row[8]) === studentId && String(row[4]) === subjectCode && String(row[6]) === className) {
      history.push({
        date: row[1], 
        period: row[7], 
        status: row[10], 
        sessionId: row[12] 
      });
    }
  });
  return history.reverse(); 
}

function saveAttendanceBatch(list) {
  if (!list || list.length === 0) return { status: "error", message: "ไม่มีข้อมูล" };
  const first = list[0];

  // 🚨 ด่านตรวจ: เช็คสิทธิ์ผู้สอนรายวิชา
  if (!verifyTeacherPermission(first.teacherId, first.subjectCode, first.className, first.term, first.year)) {
     return { status: "error", message: "❌ ความปลอดภัย: คุณไม่มีสิทธิ์เช็คชื่อในวิชาและห้องนี้!" };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // เข้าคิว
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance_Database");
    const ts = new Date();
    const rows = list.map(item => [
      ts, item.date, item.term, item.year, item.subjectCode, item.subjectName,
      item.className, item.period, item.studentId, item.studentName, item.status, 
      item.teacherId, `${item.date}|${item.subjectCode}|${item.className}|${item.period}`
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    SpreadsheetApp.flush(); // ดันข้อมูลลงชีต
    return { status: "success", message: "✅ เช็คชื่อเรียบร้อย" };
    
  } catch (e) { 
    return { status: "error", message: "คิวบันทึกเต็ม กรุณากดบันทึกอีกครั้งครับ" }; 
  } finally {
    lock.releaseLock(); // คืนคิว
  }
}

/**
 * 📊 ดึงรายงานสถิติ (ฉบับแก้ไขสูตร: เริ่ม 100% แล้วค่อยๆ ลด + แก้บั๊กโควตา)
 */
function getSemesterReport(subjectCode, className, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName("Attendance_Database");
  const timeSheet = ss.getSheetByName("Timetable_Database");
  
  const attData = attSheet ? attSheet.getDataRange().getDisplayValues() : [];
  const timeData = timeSheet ? timeSheet.getDataRange().getDisplayValues() : [];

  const normalize = (str) => String(str || "").replace(/[^a-zA-Z0-9ก-๙]/g, '');

  const cleanSub = normalize(subjectCode);
  const cleanClass = normalize(className); 
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();
  
  // --- 1. คำนวณคาบทั้งหมดตามโครงสร้าง (Total Expected) ---
  let periodsPerWeek = 0;
  
  // พยายามค้นหาแบบเข้มข้น (หาจาก New Schema Level/Room)
  for (let i = 1; i < timeData.length; i++) {
    const row = timeData[i];
    // New Schema: 0:Code, 2:Level, 3:Room, 8:Term, 9:Year
    const tLevel = String(row[2]).trim();
    const tRoom = String(row[3]).trim();
    const tClassID = normalize(`${tLevel}/${tRoom}`); // ประกอบร่าง "ม51"

    if (normalize(row[0]) === cleanSub && 
        tClassID === cleanClass &&
        String(row[8]).trim() === targetTerm && 
        String(row[9]).trim() === targetYear) {
      periodsPerWeek++;
    }
  }

  // 🛡️ ระบบกันพลาด: ถ้าหาไม่เจอจริงๆ ให้สมมติว่ามี 3 คาบ/สัปดาห์ (กันโควตาเป็น 0)
  if (periodsPerWeek === 0) {
    console.log("⚠️ หาคาบเรียนไม่เจอ ใช้ค่า Default 3 คาบ/สัปดาห์");
    periodsPerWeek = 3;
  }

  const weeksPerTerm = 20; 
  const totalCoursePeriods = periodsPerWeek * weeksPerTerm; // คะแนนเต็ม (เช่น 60 คาบ)
  const maxAbsenceQuota = Math.floor(totalCoursePeriods * 0.2); // ขาดได้สูงสุด (20%)

  // --- 2. ดึงข้อมูลนักเรียน (Group By Session) ---
  const studentDataMap = {}; 
  const allSessions = new Set();
  const studentInfo = {}; 

  for (let i = 1; i < attData.length; i++) {
    const row = attData[i];
    if (!row[1]) continue;

    const rowSub = normalize(row[4]);
    const rowClass = normalize(row[6]); 

    if (rowSub === cleanSub && rowClass === cleanClass) {
      
      const stdID = String(row[8]).trim();
      const stdName = row[9];
      const status = row[10];
      const sessionID = String(row[12]).trim() || (row[1] + "_" + row[7]);
      
      allSessions.add(sessionID);
      
      if (!studentInfo[stdID]) studentInfo[stdID] = stdName;
      if (!studentDataMap[stdID]) studentDataMap[stdID] = {};
      
      studentDataMap[stdID][sessionID] = status;
    }
  }
  
  const currentTotalTaught = allSessions.size;

  // --- 3. สรุปผลรายคน (สูตรใหม่: เริ่ม 100% แล้วลดลง) ---
  const reportData = Object.keys(studentDataMap).map(stdID => {
    let present = 0, late = 0, leave = 0, absent = 0;
    
    const records = studentDataMap[stdID];
    for (const sessKey in records) {
      const s = records[sessKey];
      if (s === 'มา') present++;
      else if (s === 'สาย') late++;
      else if (s === 'ลา') leave++;
      else if (s === 'ขาด') absent++;
    }

    // สูตรใหม่: (คาบทั้งหมด - ขาด - ลา) / คาบทั้งหมด * 100
    // *หมายเหตุ: ถ้าครูอยากให้ "ลา" ไม่เสียคะแนน ให้ลบ leave ออกจากสูตรลบ
    const totalMissed = absent + leave; // นับทั้งขาดและลา เป็นตัวหักคะแนน
    
    // คำนวณ % จาก "ทั้งเทอม" (Start at 100%)
    let percent = ((totalCoursePeriods - totalMissed) / totalCoursePeriods) * 100;
    
    // แต่ถ้า % ปัจจุบันจริงๆ (Present/Taught) มันดีกว่า ก็ให้โชว์อันที่ดีกว่า (Optional)
    // หรือเอาแบบตรงไปตรงมาตามที่ครูขอคือสูตรนี้เลย:
    
    return { 
      id: stdID, 
      name: studentInfo[stdID], 
      present, late, leave, absent, 
      total: present + late + leave + absent,
      percent: percent.toFixed(2), 
      currentTotalTaught 
    };
  });

  return {
    students: reportData.sort((a, b) => a.id.localeCompare(b.id)),
    meta: { 
      periodsPerWeek, 
      weeksPerTerm, 
      totalCoursePeriods, 
      maxAbsenceQuota,
      currentTotalTaught 
    }
  };
}

/**
 * 📊 ดึงรายงานสถิติ "รวมทุกรายวิชา" (All Subjects Report)
 */
function getAllSubjectsReport(teacherId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName("Attendance_Database");
  const timeSheet = ss.getSheetByName("Timetable_Database");

  const attData = attSheet ? attSheet.getDataRange().getDisplayValues() : [];
  const timeData = timeSheet ? timeSheet.getDataRange().getDisplayValues() : [];

  const normalize = (str) => String(str || "").replace(/[^a-zA-Z0-9ก-๙]/g, '');
  const searchTeacherId = String(teacherId).trim().toLowerCase();
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();

  // 1. กวาดรายวิชาที่ครูสอนทั้งหมดในเทอมนี้
  const teacherClasses = {};
  for (let i = 1; i < timeData.length; i++) {
    const row = timeData[i];
    const tTeacherID = String(row[5]).trim().toLowerCase();
    
    if (tTeacherID === searchTeacherId && String(row[8]).trim() === targetTerm && String(row[9]).trim() === targetYear) {
      const tCode = normalize(row[0]);
      const tClassID = normalize(`${String(row[2]).trim()}/${String(row[3]).trim()}`);
      const key = `${tCode}_${tClassID}`;

      if (!teacherClasses[key]) {
        teacherClasses[key] = {
          rawCode: row[0],
          rawName: row[1],
          rawClassID: `${String(row[2]).trim()}/${String(row[3]).trim()}`,
          periodsPerWeek: 0,
          sessions: new Set(),
          students: {}
        };
      }
      teacherClasses[key].periodsPerWeek++;
    }
  }

  // 2. ดึงประวัติการเช็คชื่อทั้งหมด แล้วจับคู่กับวิชาที่สอน
  for (let i = 1; i < attData.length; i++) {
    const row = attData[i];
    if (!row[1]) continue;
    if(String(row[2]).trim() !== targetTerm || String(row[3]).trim() !== targetYear) continue;

    const key = `${normalize(row[4])}_${normalize(row[6])}`;
    if (teacherClasses[key]) {
      const stdID = String(row[8]).trim();
      const stdName = row[9];
      const sessionID = String(row[12]).trim() || (row[1] + "_" + row[7]);

      teacherClasses[key].sessions.add(sessionID);
      if (!teacherClasses[key].students[stdID]) teacherClasses[key].students[stdID] = { name: stdName, records: {} };
      
      teacherClasses[key].students[stdID].records[sessionID] = row[10]; // เก็บสถานะ
    }
  }

  // 3. ประมวลผลเป็น % ของแต่ละคนในแต่ละวิชา
  const weeksPerTerm = 20;
  let allStudents = [];

  for (const key in teacherClasses) {
    const cls = teacherClasses[key];
    const currentTotalTaught = cls.sessions.size;
    
    if (currentTotalTaught === 0) continue; // ข้ามวิชาที่ยังไม่เคยสอนเลย

    const actualPeriodsPerWeek = cls.periodsPerWeek > 0 ? cls.periodsPerWeek : 3;
    const totalCoursePeriods = actualPeriodsPerWeek * weeksPerTerm;
    const maxAbsenceQuota = Math.floor(totalCoursePeriods * 0.2);

    for (const stdID in cls.students) {
      const student = cls.students[stdID];
      let present = 0, late = 0, leave = 0, absent = 0;
      
      for (const sess in student.records) {
        const s = student.records[sess];
        if (s === 'มา') present++; else if (s === 'สาย') late++; else if (s === 'ลา') leave++; else if (s === 'ขาด') absent++;
      }

      const totalMissed = absent + leave;
      const percent = ((totalCoursePeriods - totalMissed) / totalCoursePeriods) * 100;
      const remainingQuota = maxAbsenceQuota - totalMissed;

      allStudents.push({
        id: stdID,
        name: student.name,
        subjectCode: cls.rawCode,
        subjectName: cls.rawName,
        className: cls.rawClassID,
        present, late, leave, absent,
        percent: percent.toFixed(2),
        taught: currentTotalTaught,
        totalCoursePeriods: totalCoursePeriods,
        remainingQuota: remainingQuota
      });
    }
  }

  // 4. เรียงลำดับตาม: วิชา -> ห้อง -> รหัสนักเรียน
  allStudents.sort((a, b) => {
    if (a.subjectCode !== b.subjectCode) return a.subjectCode.localeCompare(b.subjectCode);
    if (a.className !== b.className) return a.className.localeCompare(b.className);
    return a.id.localeCompare(b.id);
  });

  return allStudents;
}

/**
 * 🏫 ดึงรายวิชาสำหรับ Dropdown (Updated for New Schema)
 */
function getTeacherSubjects(userId, userRole, targetTerm, targetYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues(); 
  const subjects = [];
  const uniqueKeys = new Set();
  
  const searchUserId = String(userId).trim().toLowerCase(); 
  const searchTerm = String(targetTerm).trim();
  const searchYear = String(targetYear).trim();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // New Schema Mapping
    const tCode = row[0].trim();
    const tName = row[1].trim();
    const tLevel = String(row[2]).trim();
    const tRoom = String(row[3]).trim();
    const tLoc = String(row[4]).trim();
    
    const tClassID = `${tLevel}/${tRoom}`; // "ม.5/1"
    const tDisplay = `${tClassID} (${tLoc})`; // "ม.5/1 (114)"
    
    const tTeacherID = String(row[5]).trim().toLowerCase(); // Index 5
    const tTerm = String(row[8]).trim(); // Index 8
    const tYear = String(row[9]).trim(); // Index 9

    let isOwner = false;
    
    if (userRole && userRole.toUpperCase() === 'ADMIN') {
      isOwner = true; 
    } else {
      // เทียบรหัสครู
      if (tTeacherID === searchUserId) isOwner = true;
      else if (tTeacherID === "teacher" + searchUserId) isOwner = true;
      else if (searchUserId === "teacher" + tTeacherID) isOwner = true;
      else if (tTeacherID.replace(/\D/g,'') !== "" && tTeacherID.replace(/\D/g,'') === searchUserId.replace(/\D/g,'')) isOwner = true;
    }

    if (isOwner && tTerm === searchTerm && tYear === searchYear) {
      const key = `${tCode}-${tClassID}`;
      if (!uniqueKeys.has(key)) {
        uniqueKeys.add(key);
        // ส่ง tClassID ไปเป็น Index 2 (เพื่อให้ตรงกับ Attendance)
        subjects.push([tCode, tName, tClassID, tDisplay]); 
      }
    }
  }
  
  return subjects;
}


function saveLessonRecord(record) {
  // 🚨 ด่านตรวจ: เช็คสิทธิ์ผู้สอน
  if (!verifyTeacherPermission(record.teacherId, record.subjectCode, record.className, record.term, record.year)) {
     return { status: "error", message: "❌ ความปลอดภัย: คุณไม่มีสิทธิ์บันทึกข้อมูลวิชานี้!" };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // เข้าคิว
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Academic_Records");
    const config = getSystemConfig();
    sheet.appendRow([
      new Date(), record.date, config.term, config.year, record.subjectCode, record.subjectName, 
      record.className, record.period, record.topic, record.totalPresent, record.totalAbsent, 
      record.totalLeave, record.teacherId, record.signature, 
      `${record.date}|${record.subjectCode}|${record.className}|${record.period}`
    ]);
    
    SpreadsheetApp.flush(); // ดันข้อมูลลงชีต
    return { status: "success", message: "✅ บันทึกข้อมูลการสอนเรียบร้อยแล้ว" };
    
  } catch (e) {
    return { status: "error", message: "คิวบันทึกเต็ม กรุณากดบันทึกอีกครั้งครับ" };
  } finally {
    lock.releaseLock(); // คืนคิว
  }
}

function getTodayAttendanceHistory(targetDateStr, subjectCode, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const cleanTargetDate = String(targetDateStr).trim(); 
  const cleanSub = String(subjectCode).trim().replace(/\s/g, ''); 
  const cleanClass = String(className).trim().replace(/\s/g, '');
  
  const uniqueHistory = {}; 

  // 🌟 เริ่มหาจากแถวล่างสุด (ล่าสุด) ถอยหลังขึ้นไป
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (!row[1]) continue; 

    let rowDateStr = "";
    if (row[1] instanceof Date) {
      rowDateStr = Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd");
    } else {
      rowDateStr = String(row[1]).substring(0, 10);
    }

    // 🚀 ท่าไม้ตาย: ถ้าวันที่ในชีต "เก่ากว่า" วันที่เราค้นหา แปลว่าทะลุไปวันอื่นแล้ว ให้หยุดหาทันที!
    if (rowDateStr < cleanTargetDate) {
       break;
    }

    const rowSub = String(row[4]).trim().replace(/\s/g, '');
    const rowClass = String(row[6]).trim().replace(/\s/g, '');

    if (rowDateStr === cleanTargetDate && rowSub === cleanSub && rowClass === cleanClass) {
      const rawID = String(row[8]).trim();
      const idNoZero = String(parseInt(rawID, 10)); 
      
      if (!uniqueHistory[idNoZero]) {
        uniqueHistory[idNoZero] = {
          studentId: rawID,
          cleanId: idNoZero,
          status: row[10],
          period: row[7],
          sessionId: row[12],
          studentName: row[9] 
        };
      }
    }
  }
  return Object.values(uniqueHistory);
}

function getCourseSessionList(subjectCode, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const cleanSub = String(subjectCode).trim().replace(/\s/g, '');
  const cleanClass = String(className).trim().replace(/\s/g, '');
  
  const sessionMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;

    const rowSub = String(row[4]).trim().replace(/\s/g, '');
    const rowClass = String(row[6]).trim().replace(/\s/g, '');

    if (rowSub === cleanSub && rowClass === cleanClass) {
      let dateKey = "";
      try {
         dateKey = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      } catch (e) { continue; }

      if (!sessionMap[dateKey]) {
        sessionMap[dateKey] = {
          date: dateKey,
          displayDate: Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
          period: row[7],
          students: new Set() 
        };
      }
      sessionMap[dateKey].students.add(String(row[8]).trim());
    }
  }

  return Object.values(sessionMap).map(s => ({
    date: s.date,
    displayDate: s.displayDate,
    period: s.period,
    count: s.students.size 
  })).sort((a, b) => b.date.localeCompare(a.date));
}

function updateAttendanceBatch(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };

  const data = sheet.getDataRange().getValues();
  
  const updateMap = {};
  list.forEach(item => {
    updateMap[String(item.studentId).trim()] = item.status;
  });

  const targetSessionID = String(list[0].sessionId).trim();
  let updateCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const rowSessionID = String(data[i][12]).trim();
    const rowStudentID = String(data[i][8]).trim();

    if (rowSessionID === targetSessionID && updateMap[rowStudentID]) {
      sheet.getRange(i + 1, 11).setValue(updateMap[rowStudentID]);
      updateCount++;
    }
  }

  return { status: "success", message: `อัปเดตข้อมูล ${updateCount} รายการเรียบร้อย` };
}

// ==========================================
// 6. TIMETABLE SYSTEM (ระบบจัดการตารางสอน)
// ==========================================

/**
 * 📅 ดึงข้อมูลตารางสอนแบบกรอง (Updated for New Schema)
 */
function getFilteredTimetables(tid, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();
  
  console.log("🔍 ค้นหาตารางสอนสำหรับ -> เทอม:", targetTerm, "ปี:", targetYear);

  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // New Schema Mapping: 
    // TeacherID=5, Term=8, Year=9
    const rowTid  = String(row[5]).trim(); 
    const rowTerm = String(row[8]).trim();
    const rowYear = String(row[9]).trim();

    const isTermMatch = (rowTerm === targetTerm);
    const isYearMatch = (rowYear === targetYear);
    const isTeacherMatch = (tid === "" || rowTid === tid);

    if (isTermMatch && isYearMatch && isTeacherMatch) {
      results.push({
        rowIndex: i + 1,
        data: row // ส่ง row ทั้งหมดกลับไป (หน้าบ้านต้องรู้ index ใหม่)
      });
    }
  }

  return results;
}

function importTimetableCSV(base64Data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName("Timetable_Database");
    
    // ... (ส่วน Import CSV อาจต้องปรับ Logic การ Split ให้เข้าช่องใหม่ ถ้า CSV ยังเป็น Format เก่า) ...
    // แต่ถ้าครูใช้ Manual Input หรือแก้ไขผ่านหน้าเว็บ ส่วนนี้อาจยังไม่กระทบมาก
    // เพื่อความชัวร์ ผมคงโค้ดเดิมไว้ก่อน หรือถ้าครูมี CSV Format ใหม่ ค่อยปรับครับ
    
    return { status: 'error', message: 'ฟังก์ชัน Import CSV อยู่ระหว่างปรับปรุงให้รองรับโครงสร้างใหม่' };
  } catch (e) { return { status: 'error', message: e.message }; }
}

function updateTimetableRow(idx, data) {
  // data ที่ส่งมาจะเป็น [Code, Name, Level, Room, Location, Teacher, Day, Period, Term, Year]
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timetable_Database").getRange(idx, 1, 1, 10).setValues([data]);
  return { status: "success", message: "อัปเดตเรียบร้อย" };
}

function deleteTimetableRow(idx) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timetable_Database").deleteRow(idx);
  return { status: "success", message: "ลบเรียบร้อย" };
}

// ==========================================
// 7. DATABASE SETUP & FIX
// ==========================================

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { name: "User_Database", headers: ["Username", "Password", "FullName", "Role", "Department", "Email", "Year"] },
    { name: "Attendance_Database", headers: ["Timestamp", "Date", "Term", "Year", "SubjectCode", "SubjectName", "Class", "Period", "StudentID", "StudentName", "Status", "TeacherID", "SessionID"] },
    { name: "Academic_Records", headers: ["Timestamp", "Date", "Term", "Year", "SubjectCode", "SubjectName", "Class", "Period", "Topic", "Present", "Absent", "Leave", "TeacherID", "Signature", "SessionID"] },
    { name: "Budgets", headers: ["ProjectID", "ProjectName", "BudgetAmount", "UsedAmount", "Balance", "Status", "Year"] },
    { name: "Leave_Records", headers: ["Timestamp", "StaffName", "Type", "StartDate", "EndDate", "Reason", "Status", "Year"] },
    { name: "Maintenance", headers: ["ID", "Timestamp", "Location", "Issue", "Reporter", "Status", "Technician"] },
    { name: "System_Settings", headers: ["Key", "Value"] },
    { name: "Timetable_Database", headers: ["SubjectCode", "SubjectName", "Level", "Room", "Location", "TeacherID", "Day", "Period", "Term", "Year"] },
    { name: "Morning_Activity", headers: ["Timestamp", "Date", "Term", "Year", "Class", "StudentID", "StudentName", "Area_Status", "Duty_Status", "Flag_Status", "TeacherID", "SessionID"] },
    { name: "Sarabun_Database", headers: ["Timestamp", "DocType", "DocNumber", "Subject", "Requester", "TargetDate", "Status", "FileURL", "Year"] }
  ];

  sheets.forEach(sh => {
    let s = ss.getSheetByName(sh.name) || ss.insertSheet(sh.name);
    // เช็คว่า Header เปลี่ยนไหม ถ้าเปลี่ยนให้เขียนทับ
    s.getRange(1, 1, 1, sh.headers.length).setValues([sh.headers]).setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
    
    if(sh.name === "System_Settings" && s.getLastRow() === 1) { s.appendRow(["Current_Term", "1"]); s.appendRow(["Current_Year", "2568"]); }
  });
  
  const uSheet = ss.getSheetByName("User_Database");
  if (uSheet.getLastRow() === 1) uSheet.appendRow(["admin", "1234", "ครูน๊อต ศิกษก", "Admin", "บริหาร", "not@school.ac.th", "2568"]);
  return "✅ ฐานข้อมูลพร้อมใช้งาน!";
}

/**
 * 🔄 ฟังก์ชันอัปเกรดโครงสร้างตารางสอน (Migrate Data) - ฉบับ No UI (ผ่านฉลุย)
 */
function migrateTimetableStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Timetable_Database";
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.error("❌ ไม่พบ Sheet ชื่อ: " + sheetName);
    return;
  }

  // 1. อ่านข้อมูลดิบทั้งหมดออกมา
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();

  // ถ้ามีแค่หัวตาราง หรือไม่มีข้อมูล ก็จบการทำงาน
  if (values.length <= 1) {
    console.log("⚠️ ไม่มีข้อมูลให้แปลง");
    return;
  }

  const newRows = [];

  // 2. วนลูปแปลงข้อมูล (เริ่มแถวที่ 1 ข้าม Header เดิม)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    // --- อ่านค่าจากโครงสร้างเก่า ---
    const subjectCode = row[0];
    const subjectName = row[1];
    const oldClassString = String(row[2]).trim(); 
    const teacherID = row[3];
    const day = row[4];
    const period = row[5];
    const term = row[6];
    const year = row[7];

    // --- ✂️ ผ่าตัดแยก Level / Location ---
    let level = oldClassString;
    let location = ""; 
    let room = "1";    // ✅ บังคับเป็นเลข 1

    const parts = oldClassString.split(/\s+/); 
    
    if (parts.length >= 2) {
      level = parts[0];      
      location = parts[1];   
    } else {
      level = parts[0];      
      location = "-";        
    }

    // --- 📝 จัดเรียงเข้าคอลัมน์ใหม่ ---
    newRows.push([
      subjectCode, // Col A
      subjectName, // Col B
      level,       // Col C (New)
      room,        // Col D (New -> 1)
      location,    // Col E (New)
      teacherID,   // Col F (Moved)
      day,         // Col G
      period,      // Col H
      term,        // Col I (Moved)
      year         // Col J (Moved)
    ]);
  }

  // 3. เขียนข้อมูลใหม่ทับลงไป
  sheet.clearContents(); 

  const newHeader = [
    "SubjectCode", "SubjectName", "Level", "Room", "Location", 
    "TeacherID", "Day", "Period", "Term", "Year"
  ];
  
  sheet.appendRow(newHeader);

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#fff2cc");
  sheet.setFrozenRows(1);
  
  console.log("✅ แปลงโครงสร้างเสร็จสมบูรณ์!");
  console.log(`📊 แปลงไปทั้งหมด ${newRows.length} รายการ`);
}

// ==========================================
// 🚀 ฟังก์ชันสร้างฐานข้อมูลสำหรับระบบ ปพ.5 แบบอัตโนมัติ
// ==========================================
function setupPorPor5Database() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. สร้าง Sheet: Subject_Config (ตั้งค่ารายวิชาและตัวชี้วัด)
  let sheetConfig = ss.getSheetByName("Subject_Config");
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet("Subject_Config");
    sheetConfig.appendRow(["subject_id", "subject_code", "class_name", "term", "year", "score_ratio", "indicators_json", "teacher_id"]);
    sheetConfig.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
    sheetConfig.setFrozenRows(1);
  }

  // 2. สร้าง Sheet: Score_Database (เก็บคะแนนดิบ)
  let sheetScore = ss.getSheetByName("Score_Database");
  if (!sheetScore) {
    sheetScore = ss.insertSheet("Score_Database");
    sheetScore.appendRow(["uid", "student_id", "subject_code", "indicator_id", "score", "term", "year"]);
    sheetScore.getRange("A1:G1").setFontWeight("bold").setBackground("#fff2cc");
    sheetScore.setFrozenRows(1);
  }

  // 3. สร้าง Sheet: Qualitative_Assess (ประเมินคุณลักษณะ อ่านเขียน สมรรถนะ)
  let sheetQual = ss.getSheetByName("Qualitative_Assess");
  if (!sheetQual) {
    sheetQual = ss.insertSheet("Qualitative_Assess");
    sheetQual.appendRow(["student_id", "subject_code", "term", "year", "reading_writing", "char_json", "comp_json"]);
    sheetQual.getRange("A1:G1").setFontWeight("bold").setBackground("#c9daf8");
    sheetQual.setFrozenRows(1);
  }

  // 4. สร้าง Sheet: Grade_Summary (สรุปผลการเรียนและตัดเกรด)
  let sheetGrade = ss.getSheetByName("Grade_Summary");
  if (!sheetGrade) {
    sheetGrade = ss.insertSheet("Grade_Summary");
    sheetGrade.appendRow(["student_id", "subject_code", "total_score", "grade", "remedial_status", "attendance_percent"]);
    sheetGrade.getRange("A1:F1").setFontWeight("bold").setBackground("#f4cccc");
    sheetGrade.setFrozenRows(1);
  }

  return "✅ สร้างฐานข้อมูล ปพ.5 ทั้ง 4 แผ่นเรียบร้อยแล้วครับ!";
}

// ==========================================
// ระบบตั้งค่าการพิมพ์ ปพ.5 (ผู้ลงนาม & ครูที่ปรึกษา)
// ==========================================

// 1. ฟังก์ชันดึงข้อมูลการตั้งค่า (ถ้าชีตไม่มี ระบบจะสร้างให้ทันที)
function getPrintConfigData(term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Print_Config');
  
  // ถ้ายังไม่มีชีต ให้สร้างใหม่
  if (!sheet) {
     sheet = ss.insertSheet('Print_Config');
     sheet.appendRow(['term', 'year', 'sys_data_json', 'homeroom_data_json']);
  }
  
  const data = sheet.getDataRange().getValues();
  
  // ค้นหาข้อมูลของเทอมและปีที่ระบุ
  for (let i = 1; i < data.length; i++) {
     if (String(data[i][0]) === String(term) && String(data[i][1]) === String(year)) {
         return {
             status: 'success',
             sys: JSON.parse(data[i][2] || '{}'),
             hr: JSON.parse(data[i][3] || '[]')
         };
     }
  }
  
  // ถ้าไม่เจอของเทอมนี้ ให้ส่งค่าว่างๆ กลับไป (หรือค่าเริ่มต้น)
  return { 
     status: 'success', 
     sys: { school_name: 'โรงเรียนภูพระบาทวิทยา', principal_name: '', measure_head: '', academic_head: '' }, 
     hr: [] 
  };
}

// 2. ฟังก์ชันบันทึกข้อมูลการตั้งค่า
function savePrintConfigData(payload) {
  const lock = LockService.getScriptLock();
  try {
     lock.waitLock(10000);
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     let sheet = ss.getSheetByName('Print_Config');
     
     if (!sheet) {
         sheet = ss.insertSheet('Print_Config');
         sheet.appendRow(['term', 'year', 'sys_data_json', 'homeroom_data_json']);
     }
     
     const data = sheet.getDataRange().getValues();
     let found = false;
     
     // วนหาว่าเคยมีข้อมูลเทอม/ปีนี้หรือยัง ถ้ามีให้เซฟทับบรรทัดเดิม
     for (let i = 1; i < data.length; i++) {
         if (String(data[i][0]) === String(payload.term) && String(data[i][1]) === String(payload.year)) {
             sheet.getRange(i + 1, 3).setValue(JSON.stringify(payload.sys));
             sheet.getRange(i + 1, 4).setValue(JSON.stringify(payload.hr));
             found = true; 
             break;
         }
     }
     
     // ถ้าไม่เคยมี ให้ขึ้นบรรทัดใหม่
     if (!found) {
         sheet.appendRow([payload.term, payload.year, JSON.stringify(payload.sys), JSON.stringify(payload.hr)]);
     }
     
     return { status: 'success', message: `✅ บันทึกตั้งค่า ปพ.5 ของภาคเรียนที่ ${payload.term}/${payload.year} เรียบร้อย!` };
  } catch(e) {
     return { status: 'error', message: e.message };
  } finally {
     lock.releaseLock();
  }
}

// ==========================================
// 🧑‍🏫 ดึงรายชื่อครูทั้งหมดไปทำ Dropdown (สำหรับตั้งค่า ปพ.5)
// ==========================================
function getTeacherListForDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const teachers = [];
  
  // วนลูปดึงเฉพาะชื่อ-สกุล (คอลัมน์ Index 2) ของคนที่เป็น Teacher หรือ Admin
  for (let i = 1; i < data.length; i++) {
    const role = String(data[i][3]).trim().toUpperCase();
    if (role === 'TEACHER') {
      teachers.push(String(data[i][2]).trim()); 
    }
  }
  
  // เรียงลำดับชื่อ ก-ฮ ให้หาใน Dropdown ง่ายๆ
  return teachers.sort((a, b) => a.localeCompare(b, 'th')); 
}

// ==========================================
// 🖨️ รวมร่าง HTML Template ปพ.5
// ==========================================
function generatePP5Template(payload) {
  const template = HtmlService.createTemplateFromFile('Template_PP5');
  template.data = payload;
  return template.evaluate().getContent();
}

// ==========================================
// 12. LESSON RECORD & FILE UPLOAD (บันทึกหลังสอนแบบละเอียด)
// ==========================================

/**
 * 🛠️ สร้างโฟลเดอร์สำหรับเก็บรูปภาพและไฟล์งานนักเรียน (ถ้ายังไม่มี)
 */
function getOrCreateUploadFolder() {
  const folderName = "PSSMS_Uploads";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

/**
 * 🛠️ อัปโหลดไฟล์จาก Base64 ไปยัง Google Drive
 */
function uploadFileToDrive(base64Data, filename) {
  // ถ้าไม่ได้เลือกไฟล์ หรือข้อมูลว่างเปล่า ให้คืนค่าว่างกลับไป
  if (!base64Data || base64Data === "" || base64Data === "null") return ""; 
  
  try {
    const folder = getOrCreateUploadFolder();
    
    // แยก data type ออกจาก base64 string
    const splitBase = base64Data.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    
    // สร้างไฟล์ลงโฟลเดอร์
    const blob = Utilities.newBlob(byteCharacters, type, filename);
    const file = folder.createFile(blob);
    
    // เปิดแชร์ลิงก์ให้ทุกคนสามารถเปิดดูรูปได้
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl(); // คืนค่าเป็น Link ของรูปภาพ
    
  } catch (e) {
    // 🚨 ถ้าระบบพัง (เช่น ไม่ได้รับอนุญาต หรือไฟล์ใหญ่ไป) ให้เขียน Error ลง Sheet แทนช่องว่าง!
    console.error("Upload Error: " + e.message);
    return "Error: " + e.message; 
  }
}

/**
 * 💾 บันทึกข้อมูลการสอนแบบละเอียด (DPA & 3R8C)
 */
function saveDetailedLessonRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Detailed_Lesson_Records");
  
  // ถ้ายังไม่มี Sheet ให้สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet("Detailed_Lesson_Records");
    const headers = [
      "Timestamp", "Date", "Term", "Year", "SubjectCode", "SubjectName", "Class", "Period", 
      "Topic", "Outcomes", "Problems", "Solutions", "DPA_Indicators", "Skills_3R8C", 
      "Student_Results", "WorkFileURL", "AtmosphereImageURL", "TeacherID", "SessionID"
    ];
    sheet.appendRow(headers);
    sheet.getRange("A1:S1").setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
  }

  // จัดการไฟล์อัปโหลด
  let workUrl = "";
  let imageUrl = "";
  const timeStampStr = new Date().getTime();
  
  if (record.workFileBase64) {
    workUrl = uploadFileToDrive(record.workFileBase64, `Work_${record.subjectCode}_${timeStampStr}`);
  }
  if (record.imageFileBase64) {
    imageUrl = uploadFileToDrive(record.imageFileBase64, `Atmosphere_${record.subjectCode}_${timeStampStr}`);
  }

  const config = getSystemConfig();
  const sessionID = `${record.date}|${record.subjectCode}|${record.className}|${record.period}`;

  sheet.appendRow([
    new Date(), record.date, config.term, config.year, record.subjectCode, record.subjectName,
    record.className, record.period, record.topic, record.outcomes, record.problems, record.solutions,
    JSON.stringify(record.dpa), JSON.stringify(record.skills), record.studentResults,
    workUrl, imageUrl, record.teacherId, sessionID
  ]);

  return { status: "success", message: "✅ บันทึกข้อมูลการสอนแบบละเอียดเรียบร้อยแล้ว!" };
}

/**
 * 📚 ดึงข้อมูลบันทึกหลังสอนแบบละเอียด (สำหรับหน้า Dashboard ประวัติ)
 */
function getDetailedLessonRecords(teacherId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  const results = [];
  
  // วนลูปจากล่างขึ้นบน เพื่อเอาข้อมูลล่าสุดขึ้นก่อน
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    if (!row[1]) continue;

    // เช็คว่าตรงกับครู เทอม และปีที่ค้นหาไหม
    if (String(row[17]).trim() === String(teacherId).trim() && 
        String(row[2]).trim() === String(term).trim() && 
        String(row[3]).trim() === String(year).trim()) {
      
      // ป้องกัน Error เวลาแปลง JSON
      let dpaArray = [];
      let skillsArray = [];
      try { dpaArray = JSON.parse(row[12]); } catch(e) {}
      try { skillsArray = JSON.parse(row[13]); } catch(e) {}

      results.push({
        timestamp: row[0],
        date: row[1],
        subjectCode: row[4],
        subjectName: row[5],
        className: row[6],
        period: row[7],
        topic: row[8],
        outcomes: row[9],
        problems: row[10],
        solutions: row[11],
        dpa: dpaArray,
        skills: skillsArray,
        studentResults: row[14],
        workUrl: row[15],
        imageUrl: row[16]
      });
    }
  }
  return results;
}

/**
 * 🗑️ ลบข้อมูลบันทึกหลังสอนแบบละเอียด
 */
function deleteDetailedLessonRecord(timestampStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };

  // 🌟 ใช้ getDisplayValues() เพื่อให้ข้อความตรงกับที่หน้าบ้านส่งมาเป๊ะๆ
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(timestampStr).trim()) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "🗑️ ลบข้อมูลบันทึกเรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลที่ต้องการลบ" };
}

/**
 * ✏️ แก้ไขอัปเดตข้อมูลบันทึกหลังสอนแบบละเอียด
 */
function updateDetailedLessonRecord(timestampStr, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };

  // 🌟 ใช้ getDisplayValues() เพื่อให้ข้อความตรงกับที่หน้าบ้านส่งมาเป๊ะๆ
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(timestampStr).trim()) {
      const row = i + 1;
      
      // อัปเดตข้อมูลข้อความ
      sheet.getRange(row, 9).setValue(record.topic);
      sheet.getRange(row, 10).setValue(record.outcomes);
      sheet.getRange(row, 11).setValue(record.problems);
      sheet.getRange(row, 12).setValue(record.solutions);
      sheet.getRange(row, 13).setValue(JSON.stringify(record.dpa));
      sheet.getRange(row, 14).setValue(JSON.stringify(record.skills));
      sheet.getRange(row, 15).setValue(record.studentResults);
      
      // ถ้ามีการอัปโหลดไฟล์ใหม่ ค่อยทับของเดิม
      if (record.workFileBase64) {
         const workUrl = uploadFileToDrive(record.workFileBase64, `Work_Updated_${new Date().getTime()}`);
         sheet.getRange(row, 16).setValue(workUrl);
      }
      if (record.imageFileBase64) {
         const imageUrl = uploadFileToDrive(record.imageFileBase64, `Atmosphere_Updated_${new Date().getTime()}`);
         sheet.getRange(row, 17).setValue(imageUrl);
      }
      
      return { status: "success", message: "✅ อัปเดตข้อมูลการสอนเรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลที่ต้องการแก้ไข" };
}

// ==========================================
// 📚 ระบบ ปพ.5: โครงสร้างรายวิชา (Subject Config) (อัปเกรดป้องกันเว้นวรรค)
// ==========================================

// ==========================================
// ปรับปรุงฟังก์ชันดึงโครงสร้างวิชา (รองรับการดึงข้อมูลจากปีเก่าอัตโนมัติ)
// ==========================================
function getSubjectConfig(subjectCode, className, term, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subject_Config");
  if(!sheet) return null;
  
  const data = sheet.getDataRange().getDisplayValues(); 
  
  const targetSubj = String(subjectCode).trim();
  const targetClass = String(className).trim();
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();
  
  let exactMatch = null;
  let historyMatch = null;
  
  // วนลูปจาก "ล่างขึ้นบน" (เพื่อให้เจอข้อมูลล่าสุดก่อนเสมอ)
  for(let i = data.length - 1; i >= 1; i--) {
    const rSubj = String(data[i][1]).trim();
    const rClass = String(data[i][2]).trim();
    const rTerm = String(data[i][3]).trim();
    const rYear = String(data[i][4]).trim();

    // ถ้ารหัสวิชาตรงกัน (ไม่สนปีการศึกษา) ให้เก็บไว้เป็น "แม่แบบสำรอง" 
    // เผื่อปีปัจจุบันยังไม่มีการตั้งค่า
    if (rSubj === targetSubj) {
      if (!historyMatch) {
        historyMatch = {
          ratio: data[i][5], 
          indicators: JSON.parse(data[i][6] || '[]')
        };
      }
      
      // แต่ถ้าเจอข้อมูลที่ "ตรงเป๊ะ" ทั้งวิชา ห้อง เทอม และปี ให้ยึดอันนี้เป็นหลักแล้วหยุดค้นหา
      if (rClass === targetClass && rTerm === targetTerm && rYear === targetYear) {
        exactMatch = {
          ratio: data[i][5], 
          indicators: JSON.parse(data[i][6] || '[]')
        };
        break; 
      }
    }
  }
  
  // ส่งคืนข้อมูลที่ตรงเป๊ะก่อน ถ้าไม่มีให้ส่งคืนแม่แบบสำรองจากปีเก่า ถ้าไม่มีเลยส่ง null
  return exactMatch || historyMatch || null; 
}

function saveSubjectConfig(configData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subject_Config");
  if(!sheet) return {status: 'error', message: 'ไม่พบ Database: Subject_Config'};
  
  const data = sheet.getDataRange().getDisplayValues();
  
  const targetSubj = String(configData.subjectCode).trim();
  const targetClass = String(configData.className).trim();
  const targetTerm = String(configData.term).trim();
  const targetYear = String(configData.year).trim();
  
  const subjectId = `${targetSubj}_${targetClass}_${targetTerm}_${targetYear}`;
  const ratioStr = `${configData.formative}:${configData.midterm}:${configData.final}`;
  
  const rowData = [
    subjectId, targetSubj, targetClass, targetTerm, targetYear,
    ratioStr, JSON.stringify(configData.indicators), configData.teacherId
  ];

  for(let i = 1; i < data.length; i++) {
    const rSubj = String(data[i][1]).trim();
    const rClass = String(data[i][2]).trim();
    const rTerm = String(data[i][3]).trim();
    const rYear = String(data[i][4]).trim();

    if(rSubj === targetSubj && rClass === targetClass && rTerm === targetTerm && rYear === targetYear) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([rowData]);
      return {status: 'success', message: 'อัปเดตโครงสร้างวิชาเรียบร้อยแล้ว!'};
    }
  }
  
  sheet.appendRow(rowData);
  return {status: 'success', message: 'บันทึกโครงสร้างวิชาใหม่เรียบร้อยแล้ว!'};
}

// ==========================================
// 13. ระบบ ปพ.5: All-in-One Score & Evaluation
// ==========================================

// ==========================================
// 📥 ระบบดึงข้อมูลคะแนน (อัปเกรด: ค้นหาคอลัมน์ remedial_status อัตโนมัติ)
// ==========================================
function getAllInOneScoreGridData(subjectCode, className, term, year) {
  let config = getSubjectConfig(subjectCode, className, term, year);
  if (!config) config = { ratio: "70:10:20", indicators: [{name: "คะแนนเก็บ 1", score: 70}] }; 

  const students = getStudentsByClass(className);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
  const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

  const existingScores = {};
  
  // 1. ดึงจาก Score_Database
  const sheetScore = ss.getSheetByName("Score_Database");
  if (sheetScore) {
    const scoreData = sheetScore.getDataRange().getDisplayValues();
    for(let i = 1; i < scoreData.length; i++) {
      if(normStr(scoreData[i][2]) === normStr(subjectCode) && normStr(scoreData[i][5]) === normStr(term) && normStr(scoreData[i][6]) === normStr(year)) {
        const stdKey = normID(scoreData[i][1]);
        const indKey = normStr(scoreData[i][3]);
        const val = String(scoreData[i][4]).trim();

        if (indKey === 'remark') {
            if (val === 'ร' || val === 'มส') existingScores[`${stdKey}_remark`] = val;
            else if ((val === '-' || val === '') && !existingScores[`${stdKey}_remark`]) existingScores[`${stdKey}_remark`] = '-';
        } else {
            existingScores[`${stdKey}_${indKey}`] = val;
        }
      }
    }
  }

  // 2. 🚨 ดึงจาก Grade_Summary (ท่าไม้ตาย: สแกนทั้งบรรทัด ไม่สนชื่อคอลัมน์!)
  const gradeSheet = ss.getSheetByName("Grade_Summary");
  if (gradeSheet && gradeSheet.getLastRow() > 0) {
    const gradeData = gradeSheet.getDataRange().getDisplayValues();
    
    for(let i = 1; i < gradeData.length; i++) {
        // เช็ครหัสวิชาแบบยืดหยุ่น ป้องกันเคสมีช่องว่างหรือตัวอักษรแปลกๆ
        const rowSub = normStr(gradeData[i][1]);
        const targetSub = normStr(subjectCode);
        
        if(rowSub === targetSub || rowSub.includes(targetSub)) {
            const stdKey = normID(gradeData[i][0]);
            let foundRemark = false;

            // 🌟 สแกนกวาดทุกเซลล์ในบรรทัดของเด็กคนนี้ (เริ่มจากคอลัมน์ที่ 3 เป็นต้นไป)
            for (let col = 2; col < gradeData[i].length; col++) {
                const cellVal = String(gradeData[i][col]).trim();
                if (cellVal === 'ร' || cellVal === 'มส') {
                    existingScores[`${stdKey}_remark`] = cellVal; // ล็อคค่าทันทีถ้าเจอ
                    foundRemark = true;
                    break; // หยุดสแกนบรรทัดนี้ เพราะเจอเป้าหมายแล้ว
                }
            }
            
            // ถ้าสแกนจนจบแล้วไม่เจอ ร/มส แต่มีข้อมูลช่องว่างหรือขีด ให้จำไว้ว่าไม่ติด
            if (!foundRemark && !existingScores[`${stdKey}_remark`]) {
                existingScores[`${stdKey}_remark`] = '-';
            }
        }
    }
  }

  const qualSheet = ss.getSheetByName("Qualitative_Assess");
  const qualData = qualSheet ? qualSheet.getDataRange().getDisplayValues() : [];
  const existingQuals = {};
  
  for(let i = 1; i < qualData.length; i++) {
    const row = qualData[i];
    if(normStr(row[1]) === normStr(subjectCode) && normStr(row[2]) === normStr(term) && normStr(row[3]) === normStr(year)) {
      // 🌟 สมองกล: เช็คว่าเป็นโครงสร้างเก่า (7 คอลัมน์) หรือโครงสร้างใหม่ (17 คอลัมน์)
      if (row.length >= 16) {
          existingQuals[normID(row[0])] = { 
              read1: row[4], read2: row[5], read3: row[6], read4: row[7], readTotal: row[8], read: row[9],
              char1: row[10], char2: row[11], char3: row[12], char4: row[13], charTotal: row[14], char: row[15],
              comp: row[16] || '3'
          };
      } else {
          existingQuals[normID(row[0])] = { read: row[4], char: row[5], comp: row[6] };
      }
    }
  }

  let attStats = {};
  try {
    const report = getSemesterReport(subjectCode, className, term, year);
    if(report && report.students) {
      report.students.forEach(s => { attStats[normID(s.id)] = parseFloat(s.percent); });
    }
  } catch(e) {}

  return { config: config, students: students, existingScores: existingScores, existingQuals: existingQuals, attStats: attStats };
}

// ==========================================
// 💾 ระบบบันทึกคะแนนทั้งหมด (Ultimate Fix - ทะลวงชีต 100%)
// ==========================================
function saveAllInOneWithConfig(payload) {
  const { subjectCode, className, teacherId, term, year, newConfig, scoreRecords, qualRecords, gradeRecords } = payload;
  
  if (typeof verifyTeacherPermission === 'function' && !verifyTeacherPermission(teacherId, subjectCode, className, term, year)) {
     return { status: 'error', message: '❌ ความปลอดภัย: คุณไม่มีสิทธิ์บันทึกคะแนนในรายวิชาและห้องนี้!' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
    const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

    // 🚨 ดักจับกรณีแผ่นงานหายหรือชื่อผิด
    const configSheet = ss.getSheetByName("Subject_Config");
    const sheetScore = ss.getSheetByName("Score_Database");
    const qualSheet = ss.getSheetByName("Qualitative_Assess");
    const gradeSheet = ss.getSheetByName("Grade_Summary");

    if (!configSheet || !sheetScore || !qualSheet || !gradeSheet) {
        return { status: 'error', message: '❌ ไม่พบชีตฐานข้อมูล ปพ.5 (อาจจะถูกลบหรือเปลี่ยนชื่อ)' };
    }

    const msMap = {};
    if (scoreRecords) {
       scoreRecords.forEach(r => {
          if (String(r.indicatorId).trim().toLowerCase() === 'remark' && r.score) {
             const v = String(r.score).trim();
             if(v === 'ร' || v === 'มส') msMap[normID(r.studentId)] = v;
          }
       });
    }
    if (gradeRecords) {
       gradeRecords.forEach(r => {
          let g = String(r.grade || '').trim();
          let rm = String(r.remark || '').trim();
          if(g === 'ร' || g === 'มส') msMap[normID(r.studentId)] = g;
          if(rm === 'ร' || rm === 'มส') msMap[normID(r.studentId)] = rm;
       });
    }

    Object.keys(msMap).forEach(sid => {
       let found = false;
       scoreRecords.forEach(r => {
          if (normID(r.studentId) === sid && String(r.indicatorId).trim().toLowerCase() === 'remark') {
              r.score = msMap[sid]; found = true;
          }
       });
       if (!found) scoreRecords.push({ studentId: sid, subjectCode: subjectCode, term: term, year: year, indicatorId: 'remark', score: msMap[sid] });
    });
    
    // 🌟 1. Config (จัดระเบียบตารางบังคับ 8 คอลัมน์เป๊ะๆ)
    if (newConfig) {
         if (configSheet.getLastRow() === 0) configSheet.appendRow(["subject_id", "subject_code", "class_name", "term", "year", "score_ratio", "indicators_json", "teacher_id"]);
         const configData = configSheet.getDataRange().getValues();
         let configUpdated = false;
         const ratioStr = `${newConfig.formative || 70}:${newConfig.midterm || 10}:${newConfig.final || 20}`;
         const subjectId = `${subjectCode}_${className}_${term}_${year}`;

         for (let i = 1; i < configData.length; i++) {
           if (String(configData[i][1]) === String(subjectCode) && String(configData[i][2]) === String(className) && String(configData[i][3]) === String(term) && String(configData[i][4]) === String(year)) {
               configData[i][5] = ratioStr; 
               configData[i][6] = JSON.stringify(newConfig.indicators || []);
               configUpdated = true; break;
           }
         }
         
         if (configUpdated) {
             const uniformData = configData.map(row => {
                 let r = row.slice(0, 8);
                 while(r.length < 8) r.push("");
                 return r;
             });
             configSheet.getRange(1, 1, uniformData.length, 8).setValues(uniformData);
         } else {
             configSheet.appendRow([subjectId, subjectCode, className, term, year, ratioStr, JSON.stringify(newConfig.indicators || []), teacherId]);
         }
    }

    let debugMsg = "";

    // 🌟 2. Score Database
    if (scoreRecords && scoreRecords.length > 0) {
        if (sheetScore.getLastRow() === 0) sheetScore.appendRow(["uid", "student_id", "subject_code", "indicator_id", "score", "term", "year"]);
        let scoreData = sheetScore.getDataRange().getValues(); 
        const scoreMap = {};
        scoreRecords.forEach(r => {
           if(r.score !== undefined && r.score !== null) { 
               const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.indicatorId)}_${normStr(r.term)}_${normStr(r.year)}`;
               scoreMap[uid] = r;
           }
        });

        let scoreUpdated = false;
        let cUpdate = 0;
        for(let i = 1; i < scoreData.length; i++) {
          const uid = `${normID(scoreData[i][1])}_${normStr(scoreData[i][2])}_${normStr(scoreData[i][3])}_${normStr(scoreData[i][5])}_${normStr(scoreData[i][6])}`;
          if(scoreMap[uid]) {
             if (String(scoreData[i][4]) !== String(scoreMap[uid].score)) {
                 scoreData[i][4] = scoreMap[uid].score; 
                 scoreUpdated = true;
                 cUpdate++;
             }
             scoreMap[uid].processed = true; 
          }
        }
        if(scoreUpdated) {
            const uniformScore = scoreData.map(row => { let r = row.slice(0, 7); while(r.length < 7) r.push(""); return r; });
            sheetScore.getRange(1, 1, uniformScore.length, 7).setValues(uniformScore);
        }

        const newScores = [];
        for (let uid in scoreMap) {
           if (!scoreMap[uid].processed) {
               const r = scoreMap[uid];
               if (r.score !== '') { // อย่าบันทึกช่องว่างเปล่าๆ ให้รกฐานข้อมูล
                   newScores.push([uid, "'" + r.studentId, r.subjectCode, r.indicatorId, r.score, r.term, r.year]);
               }
           }
        }
        if (newScores.length > 0) {
            sheetScore.getRange(sheetScore.getLastRow() + 1, 1, newScores.length, 7).setValues(newScores);
        }
        debugMsg += `(คะแนนใหม่: ${newScores.length}, อัปเดต: ${cUpdate})`;
    }

    // 🌟 3. Qualitative Assess (อัปเกรดเก็บคะแนนดิบ 17 คอลัมน์)
    if (qualRecords && qualRecords.length > 0) {
        if (qualSheet.getLastRow() === 0) {
            qualSheet.appendRow(["student_id", "subject_code", "term", "year", "read1", "read2", "read3", "read4", "readTotal", "read_grade", "char1", "char2", "char3", "char4", "charTotal", "char_grade", "comp"]);
        }
        
        // 🚨 บังคับขยายชีตอัตโนมัติ ถ้าคอลัมน์ไม่พอ 17 คอลัมน์
        if (qualSheet.getMaxColumns() < 17) {
            qualSheet.insertColumnsAfter(qualSheet.getMaxColumns(), 17 - qualSheet.getMaxColumns());
        }

        let qualData = qualSheet.getDataRange().getValues();
        const qualMap = {};
        qualRecords.forEach(r => { qualMap[`${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.term)}_${normStr(r.year)}`] = r; });

        let qualUpdated = false;
        for(let i = 1; i < qualData.length; i++) {
          const uid = `${normID(qualData[i][0])}_${normStr(qualData[i][1])}_${normStr(qualData[i][2])}_${normStr(qualData[i][3])}`;
          if(qualMap[uid]) {
             const q = qualMap[uid];
             qualData[i][4] = q.read1; qualData[i][5] = q.read2; qualData[i][6] = q.read3; qualData[i][7] = q.read4; 
             qualData[i][8] = q.readTotal; qualData[i][9] = q.read; 
             qualData[i][10] = q.char1; qualData[i][11] = q.char2; qualData[i][12] = q.char3; qualData[i][13] = q.char4; 
             qualData[i][14] = q.charTotal; qualData[i][15] = q.char; 
             qualData[i][16] = q.comp || '3';
             qualUpdated = true;
             qualMap[uid].processed = true;
          }
        }
        if(qualUpdated) {
            const uniformQual = qualData.map(row => { let r = row.slice(0, 17); while(r.length < 17) r.push(""); return r; });
            qualSheet.getRange(1, 1, uniformQual.length, 17).setValues(uniformQual);
        }
        
        const newQuals = [];
        for (let uid in qualMap) {
           if(!qualMap[uid].processed) {
              const r = qualMap[uid];
              newQuals.push([
                  "'" + r.studentId, r.subjectCode, r.term, r.year, 
                  r.read1, r.read2, r.read3, r.read4, r.readTotal, r.read,
                  r.char1, r.char2, r.char3, r.char4, r.charTotal, r.char,
                  r.comp || '3'
              ]);
           }
        }
        if(newQuals.length > 0) qualSheet.getRange(qualSheet.getLastRow() + 1, 1, newQuals.length, 17).setValues(newQuals);
    }

    // 🌟 4. Grade Summary
    if (gradeRecords && gradeRecords.length > 0) { 
        if (gradeSheet.getLastRow() === 0) gradeSheet.appendRow(["student_id", "subject_code", "total_score", "grade", "remedial_status", "attendance_percent"]);
        let gradeData = gradeSheet.getDataRange().getValues();
        const gradeMap = {};
        
        gradeRecords.forEach(r => {
            const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}`;
            let cleanRemark = String(r.remark || '').trim();
            if (cleanRemark === '') cleanRemark = '-';
            r.remark = cleanRemark; 
            gradeMap[uid] = r;
        });

        let gradeUpdated = false;
        for(let i = 1; i < gradeData.length; i++) {
            const row = gradeData[i];
            const uid = `${normID(row[0])}_${normStr(row[1])}`;
            if(gradeMap[uid]) {
                if(String(gradeData[i][2]) !== String(gradeMap[uid].totalScore) || 
                   String(gradeData[i][3]) !== String(gradeMap[uid].grade) || 
                   String(gradeData[i][4]) !== String(gradeMap[uid].remark)) {
                    
                    gradeData[i][2] = gradeMap[uid].totalScore;
                    gradeData[i][3] = gradeMap[uid].grade;
                    gradeData[i][4] = gradeMap[uid].remark; 
                    gradeUpdated = true;
                }
                gradeMap[uid].processed = true;
            }
        }
        
        if(gradeUpdated) {
            const uniformGrade = gradeData.map(row => { let r = row.slice(0, 6); while(r.length < 6) r.push(""); return r; });
            gradeSheet.getRange(1, 1, uniformGrade.length, 6).setValues(uniformGrade);
        }
        
        const newGrades = [];
        for (let uid in gradeMap) {
            if(!gradeMap[uid].processed) {
                const r = gradeMap[uid];
                newGrades.push(["'" + r.studentId, r.subjectCode, r.totalScore, r.grade, r.remark, "100"]);
            }
        }
        if(newGrades.length > 0) gradeSheet.getRange(gradeSheet.getLastRow() + 1, 1, newGrades.length, 6).setValues(newGrades);
    } 

    SpreadsheetApp.flush(); 
    return {status: 'success', message: `✅ บันทึกเสร็จสมบูรณ์! ${debugMsg}`};
    
  } catch(e) {
    return { status: 'error', message: e.message + " | บรรทัดที่: " + (e.lineNumber || 'ไม่ทราบ') };
  } finally {
    lock.releaseLock();
  }
}

// 🌟 ฟังก์ชันบันทึกขั้นสูง (ความเร็วแสง - Batch Array Update)
function saveAllInOneScores(payload) {
  // 🌟 ดึง className และ teacherId ออกมาเช็คด้วย
  const { subjectCode, className, teacherId, term, year, scoreRecords, qualRecords, gradeRecords } = payload;
  
  // 🚨 ด่านตรวจ: แกะโค้ดมาแก้คะแนนใช่ไหม? โดนเตะกลับ!
  if (!verifyTeacherPermission(teacherId, subjectCode, className, term, year)) {
     return { status: 'error', message: '❌ ความปลอดภัย: คุณไม่มีสิทธิ์บันทึกคะแนนในรายวิชาและห้องนี้!' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
  const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

  // ==============================
  // 1. บันทึกคะแนน (Score_Database)
  // ==============================
  const sheetScore = ss.getSheetByName("Score_Database");
  let scoreData = sheetScore.getDataRange().getValues(); 
  
  const scoreMap = {};
  scoreRecords.forEach(r => {
     const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.indicatorId)}_${normStr(r.term)}_${normStr(r.year)}`;
     scoreMap[uid] = r;
  });

  let scoreUpdated = false;
  for(let i = 1; i < scoreData.length; i++) {
    const row = scoreData[i];
    const uid = `${normID(row[1])}_${normStr(row[2])}_${normStr(row[3])}_${normStr(row[5])}_${normStr(row[6])}`;
    if(scoreMap[uid]) {
       if (String(scoreData[i][4]) !== String(scoreMap[uid].score)) {
           logScoreHistory(teacherId, scoreMap[uid].studentId, subjectCode, scoreMap[uid].indicatorId, scoreData[i][4], scoreMap[uid].score, term, year);
           scoreData[i][4] = scoreMap[uid].score; 
           scoreUpdated = true;
       }
       scoreMap[uid].processed = true; 
    }
  }
  if(scoreUpdated) sheetScore.getRange(1, 1, scoreData.length, scoreData[0].length).setValues(scoreData);

  const newScores = [];
  for (let uid in scoreMap) {
     if (!scoreMap[uid].processed) {
         const r = scoreMap[uid];
         newScores.push([uid, "'" + r.studentId, r.subjectCode, r.indicatorId, r.score, r.term, r.year]);
     }
  }
  if (newScores.length > 0) sheetScore.getRange(sheetScore.getLastRow() + 1, 1, newScores.length, newScores[0].length).setValues(newScores);

  // ==============================
  // 2. บันทึกคุณลักษณะ (Qualitative_Assess) (อัปเกรด 17 คอลัมน์)
  // ==============================
  const qualSheet = ss.getSheetByName("Qualitative_Assess");
  if (qualSheet.getMaxColumns() < 17) qualSheet.insertColumnsAfter(qualSheet.getMaxColumns(), 17 - qualSheet.getMaxColumns());
  let qualData = qualSheet.getDataRange().getValues();
  const qualMap = {};
  qualRecords.forEach(r => {
     const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.term)}_${normStr(r.year)}`;
     qualMap[uid] = r;
  });

  let qualUpdated = false;
  for(let i = 1; i < qualData.length; i++) {
    const row = qualData[i];
    const uid = `${normID(row[0])}_${normStr(row[1])}_${normStr(row[2])}_${normStr(row[3])}`;
    if(qualMap[uid]) {
       const q = qualMap[uid];
       qualData[i][4] = q.read1; qualData[i][5] = q.read2; qualData[i][6] = q.read3; qualData[i][7] = q.read4; 
       qualData[i][8] = q.readTotal; qualData[i][9] = q.read; 
       qualData[i][10] = q.char1; qualData[i][11] = q.char2; qualData[i][12] = q.char3; qualData[i][13] = q.char4; 
       qualData[i][14] = q.charTotal; qualData[i][15] = q.char; 
       qualData[i][16] = q.comp || '3';
       qualUpdated = true;
       qualMap[uid].processed = true;
    }
  }
  if(qualUpdated) {
      const uniformQual = qualData.map(row => { let r = row.slice(0, 17); while(r.length < 17) r.push(""); return r; });
      qualSheet.getRange(1, 1, uniformQual.length, 17).setValues(uniformQual);
  }
  
  const newQuals = [];
  for (let uid in qualMap) {
     if(!qualMap[uid].processed) {
        const r = qualMap[uid];
        newQuals.push([
            "'" + r.studentId, r.subjectCode, r.term, r.year, 
            r.read1, r.read2, r.read3, r.read4, r.readTotal, r.read,
            r.char1, r.char2, r.char3, r.char4, r.charTotal, r.char,
            r.comp || '3'
        ]);
     }
  }
  if(newQuals.length > 0) qualSheet.getRange(qualSheet.getLastRow() + 1, 1, newQuals.length, 17).setValues(newQuals);

  // ==============================
  // 3. บันทึกเกรด (Grade_Summary)
  // ==============================
  const gradeSheet = ss.getSheetByName("Grade_Summary");
  let gradeData = gradeSheet.getDataRange().getValues();
  const gradeMap = {};
  gradeRecords.forEach(r => {
     const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}`;
     gradeMap[uid] = r;
  });

  let gradeUpdated = false;
  for(let i = 1; i < gradeData.length; i++) {
    const row = gradeData[i];
    const uid = `${normID(row[0])}_${normStr(row[1])}`;
    if(gradeMap[uid]) {
       if(String(gradeData[i][2]) !== String(gradeMap[uid].totalScore) || String(gradeData[i][3]) !== String(gradeMap[uid].grade)) {
           gradeData[i][2] = gradeMap[uid].totalScore;
           gradeData[i][3] = gradeMap[uid].grade;
           gradeUpdated = true;
       }
       gradeMap[uid].processed = true;
    }
  }
  if(gradeUpdated) gradeSheet.getRange(1, 1, gradeData.length, gradeData[0].length).setValues(gradeData);
  
  const newGrades = [];
  for (let uid in gradeMap) {
     if(!gradeMap[uid].processed) {
        const r = gradeMap[uid];
        newGrades.push(["'" + r.studentId, r.subjectCode, r.totalScore, r.grade, "-", "100"]);
     }
  }
  if(newGrades.length > 0) gradeSheet.getRange(gradeSheet.getLastRow() + 1, 1, newGrades.length, newGrades[0].length).setValues(newGrades);

  SpreadsheetApp.flush(); 
  return {status: 'success', message: 'บันทึกคะแนน เกรด และคุณลักษณะเรียบร้อยแล้ว!'};
}

// ==========================================
// 🕵️‍♂️ ระบบประวัติการแก้ไขคะแนน (Audit Log)
// ==========================================
function logScoreHistory(teacherId, stdId, subCode, indId, oldScore, newScore, term, year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Score_History");
    if (!sheet) {
        sheet = ss.insertSheet("Score_History");
        sheet.appendRow(["Timestamp", "TeacherID", "StudentID", "SubjectCode", "IndicatorID", "OldScore", "NewScore", "Term", "Year"]);
    }
    // จดเฉพาะกรณีที่มีการเปลี่ยนค่าจริงๆ
    if (String(oldScore).trim() !== String(newScore).trim()) {
        const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        sheet.appendRow([timeStr, teacherId, stdId, subCode, indId, oldScore, newScore, term, year]);
    }
  } catch(e) {} 
}

function getScoreHistory(stdId, subCode, indId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Score_History");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  const history = [];
  
  // วนลูปดึงข้อมูลจากล่างขึ้นบน
  for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      // ป้องกัน Error ช่องว่างด้วย String().trim()
      if (String(row[2]).trim() === String(stdId).trim() && 
          String(row[3]).trim() === String(subCode).trim() && 
          String(row[4]).trim() === String(indId).trim() && 
          String(row[7]).trim() === String(term).trim() && 
          String(row[8]).trim() === String(year).trim()) {
          
          history.push({
              time: row[0], 
              old: row[5] === "" ? "-" : row[5], 
              new: row[6] === "" ? "-" : row[6]  
          });
          if(history.length >= 10) break; // โชว์ 10 รายการล่าสุด
      }
  }
  return history;
}

// ==========================================
// ☀️ 14. ระบบกิจกรรมยามเช้า (เข้าเขต / ทำเวร / เสาธง)
// ==========================================

// 1. ฟังก์ชันดึงข้อมูลการเช็คชื่อยามเช้า (เพื่อนำมาแสดงค่าเดิมถ้าเคยเช็คไปแล้ว)
function getMorningActivityData(dateStr, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Morning_Activity");
  if (!sheet) return {};

  const data = sheet.getDataRange().getDisplayValues();
  const targetSession = `${dateStr}_${className}`;
  const results = {};

  // 🌟 เริ่มหาจากแถวล่างสุด ถอยหลังขึ้นไป
  for (let i = data.length - 1; i >= 1; i--) {
    
    // ดึงวันที่จากคอลัมน์ B (Index 1) มาเช็ค
    const rowDateStr = String(data[i][1]).substring(0, 10);
    
    // 🚀 หยุดลูปทันทีถ้าเจอข้อมูลที่เก่ากว่าวันที่ค้นหา ประหยัดเวลาไปได้เยอะมาก
    if (rowDateStr < dateStr) break;

    // คอลัมน์ L (Index 11) คือ SessionID
    if (String(data[i][11]) === targetSession) {
      const stdId = String(data[i][5]); 
      
      // เก็บเฉพาะค่าแรกที่เจอ (ซึ่งก็คือค่าล่าสุดเพราะเราวนลูปจากล่างขึ้นบน)
      if (!results[stdId]) {
         results[stdId] = {
           area: data[i][7], 
           duty: data[i][8], 
           flag: data[i][9]  
         };
      }
    }
  }
  return results; 
}

// 2. ฟังก์ชันบันทึกข้อมูล (รองรับการบันทึกใหม่และการอัปเดตข้อมูลเดิม)
function saveMorningActivityBatch(payload) {
  const { date, term, year, className, teacherId, records } = payload;
  
  // 🚨 ด่านตรวจ: เช็คสิทธิ์ครูที่ปรึกษา (รหัส HR)
  if (!verifyTeacherPermission(teacherId, 'HR', className, term, year)) {
     return { status: "error", message: "❌ ความปลอดภัย: คุณไม่ใช่ครูที่ปรึกษาของห้องนี้!" };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Morning_Activity");
    
    if (!sheet) return { status: "error", message: "ไม่พบชีต Morning_Activity กรุณารัน setupDatabase ก่อนครับ" };

    const sessionID = `${date}_${className}`;
    const timestamp = new Date();
    
    const data = sheet.getDataRange().getValues();
    let rowMap = {};
    
    // หาแถวเดิมที่เคยบันทึกไว้
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][11]) === sessionID) {
        rowMap[String(data[i][5])] = i + 1;
      }
    }

    const newRows = [];

    records.forEach(r => {
      const stdId = String(r.studentId);
      if (rowMap[stdId]) {
        sheet.getRange(rowMap[stdId], 8, 1, 3).setValues([[r.area, r.duty, r.flag]]);
      } else {
        newRows.push([
          timestamp, date, term, year, className, stdId, r.studentName,
          r.area, r.duty, r.flag, teacherId, sessionID
        ]);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
    }
    
    SpreadsheetApp.flush(); 
    return { status: "success", message: "✅ บันทึกข้อมูลกิจกรรมโฮมรูมเรียบร้อยแล้ว!" };
    
  } catch (e) {
    return { status: "error", message: "ระบบกำลังมีผู้ใช้งานพร้อมกันจำนวนมาก กรุณากดบันทึกอีกครั้ง" };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ☀️ ฟังก์ชันดึงสรุปโฮมรูมรายวัน สำหรับ Dashboard
// ==========================================
function getTodayMorningSummary(teacherId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeSheet = ss.getSheetByName("Timetable_Database");
  const mornSheet = ss.getSheetByName("Morning_Activity");

  if (!timeSheet || !mornSheet) return null;

  const timeData = timeSheet.getDataRange().getDisplayValues();
  const mornData = mornSheet.getDataRange().getDisplayValues();

  const now = new Date();
  const todayStr = new Date(now.getTime() - (now.getTimezoneOffset() * 60000)).toISOString().split('T')[0];

  let hrClass = "";
  for (let i = 1; i < timeData.length; i++) {
    const tTeacherID = String(timeData[i][5]).trim().toLowerCase();
    const tCode = String(timeData[i][0]).toUpperCase();
    const tName = String(timeData[i][1]);
    
    if (tTeacherID === String(teacherId).trim().toLowerCase() &&
        (tCode === 'HR' || tName.includes('โฮมรูม')) &&
        String(timeData[i][8]).trim() === String(term).trim() &&
        String(timeData[i][9]).trim() === String(year).trim()) {
      hrClass = `${String(timeData[i][2]).trim()}/${String(timeData[i][3]).trim()}`;
      break;
    }
  }

  if (!hrClass) return { hasHR: false };

  const targetSession = `${todayStr}_${hrClass}`;
  const latestData = {};

  // 🌟 เริ่มหาจากแถวล่างสุด ถอยหลังขึ้นไป
  for (let i = mornData.length - 1; i >= 1; i--) {
    
    // ดึงวันที่จากคอลัมน์ B (Index 1) มาเช็ค
    const rowDateStr = String(mornData[i][1]).substring(0, 10);
    
    // 🚀 หยุดลูปทันทีถ้าเจอข้อมูลของเมื่อวาน
    if (rowDateStr < todayStr) break;

    if (String(mornData[i][11]) === targetSession) {
      const stdName = String(mornData[i][6]).trim(); 
      
      // เก็บเฉพาะข้อมูลล่าสุดเท่านั้น
      if (!latestData[stdName]) {
         latestData[stdName] = {
           area: String(mornData[i][7]).trim(),
           duty: String(mornData[i][8]).trim(),
           flag: String(mornData[i][9]).trim()
         };
      }
    }
  }

  const summary = {
    className: hrClass,
    absent: [], late: [], leave: [], notArea: [], notDuty: [],
    hasData: Object.keys(latestData).length > 0
  };

  for (const name in latestData) {
    const d = latestData[name];
    if (d.flag === 'ขาด') summary.absent.push(name);
    if (d.flag === 'สาย') summary.late.push(name);
    if (d.flag === 'ลา') summary.leave.push(name);
    if (d.area === 'ไม่เข้า') summary.notArea.push(name);
    if (d.duty === 'ไม่ทำ') summary.notDuty.push(name);
  }

  return { hasHR: true, data: summary };
}

// ==========================================
// 🚨 ดึงข้อมูล Dashboard กลุ่มเสี่ยง (0, ร, มส.) สำหรับครู (แก้บัคไม่ทราบชื่อ 100%)
// ==========================================
// ==========================================
// 🚨 ดึงข้อมูล Dashboard กลุ่มเสี่ยง (0, ร, มส.) สำหรับครู (Ultimate Fix: กรองวิชากิจกรรม & กรองวิชาที่ยังไม่กรอกคะแนน)
// ==========================================
function getTeacherRiskDashboard(teacherId, term, year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };

    const timeSheet = ss.getSheetByName("Timetable_Database");
    if (!timeSheet) return { status: 'error', message: 'ไม่พบฐานข้อมูลตารางสอน' };

    const timeData = timeSheet.getDataRange().getValues();
    const teacherSubjects = {}; 

    const searchTeacher = String(teacherId).trim().toLowerCase();
    for (let i = 1; i < timeData.length; i++) {
      if (String(timeData[i][5]).trim().toLowerCase() === searchTeacher &&
          String(timeData[i][8]).trim() === String(term).trim() &&
          String(timeData[i][9]).trim() === String(year).trim()) {
          let subCode = String(timeData[i][0]).trim();
          let subName = String(timeData[i][1]).trim();
          teacherSubjects[subCode] = subName; 
      }
    }

    if (Object.keys(teacherSubjects).length === 0) return { status: 'success', summary: { zero: 0, r: 0, ms: 0 }, details: [] };

    const userSheet = ss.getSheetByName("User_Database");
    const studentMap = {};
    if (userSheet) {
        const userData = userSheet.getDataRange().getDisplayValues();
        for(let i = 1; i < userData.length; i++) {
            if (String(userData[i][3]).toLowerCase() === 'student' || String(userData[i][3]) === 'นักเรียน') {
                let rawId = String(userData[i][0]).replace(/'/g, '').trim(); 
                studentMap[normID(rawId)] = {
                    displayId: rawId,
                    name: String(userData[i][2]).trim(),
                    cls: String(userData[i][4]).trim()
                };
            }
        }
    }

    let riskList = [];
    let count0 = 0, countR = 0, countMS = 0;
    let riskCheckMap = {}; 

    const gradeSheet = ss.getSheetByName("Grade_Summary");
    if (gradeSheet) {
        const gradeData = gradeSheet.getDataRange().getDisplayValues();
        
        // 🌟 ด่านที่ 2: วนลูปเช็ค "คะแนนสูงสุด" ของแต่ละรายวิชาก่อน
        const subjectMaxScore = {};
        for (let i = 1; i < gradeData.length; i++) {
            let subCode = String(gradeData[i][1]).trim();
            let totalScore = parseFloat(gradeData[i][2]) || 0;
            if (!subjectMaxScore[subCode]) subjectMaxScore[subCode] = 0;
            if (totalScore > subjectMaxScore[subCode]) subjectMaxScore[subCode] = totalScore;
        }

        for (let i = 1; i < gradeData.length; i++) {
           let safeId = normID(gradeData[i][0]); 
           let subCode = String(gradeData[i][1]).trim();
           let grade = String(gradeData[i][3]).trim();   // คอลัมน์ เกรด
           let remark = String(gradeData[i][4] || '').trim(); // คอลัมน์ หมายเหตุ (ร, มส)

           if (teacherSubjects[subCode]) {
               let riskType = null;
               
               // 🌟 ด่านที่ 1: เช็คว่าเป็นวิชากิจกรรมหรือไม่ (ขึ้นต้นด้วย ก. หรือ I.)
               let isActivitySubject = subCode.startsWith('ก') || subCode.startsWith('I') || subCode.startsWith('i');

               // กฎข้อ 1: ร และ มส ติดได้ทุกวิชา (รวมถึงวิชากิจกรรม)
               if (grade === 'ร' || remark === 'ร') riskType = 'ร';
               else if (grade === 'มส' || remark === 'มส') riskType = 'มส';
               
               // กฎข้อ 2: แจกเกรด 0 เฉพาะ "วิชาที่ไม่ใช่กิจกรรม" และ "ครูได้เริ่มกรอกคะแนนวิชานี้ไปบ้างแล้ว (MaxScore > 0)"
               else if (!isActivitySubject && subjectMaxScore[subCode] > 0) {
                   if (grade === '0' || grade === '0.0') riskType = '0';
               }

               if (riskType) {
                   let key = `${safeId}_${subCode}`;
                   if (!riskCheckMap[key]) {
                       riskCheckMap[key] = true;
                       if (riskType === '0') count0++;
                       else if (riskType === 'ร') countR++;
                       else if (riskType === 'มส') countMS++;

                       riskList.push({
                           stdId: studentMap[safeId] ? studentMap[safeId].displayId : safeId,
                           stdName: studentMap[safeId] ? studentMap[safeId].name : "ไม่ทราบชื่อ",
                           className: studentMap[safeId] ? studentMap[safeId].cls : "-",
                           subjectCode: subCode,
                           subjectName: teacherSubjects[subCode],
                           type: riskType
                       });
                   }
               }
           }
        }
    }

    return { status: 'success', summary: { zero: count0, r: countR, ms: countMS }, details: riskList.sort((a, b) => a.className.localeCompare(b.className)) };

  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// ==========================================
// ⚡ ระบบ Auto-Save เฉพาะกิจสำหรับ ร และ มส (ยิงตรงลงชีตทันที)
// ==========================================
function saveStudentRemarkDirectly(studentId, subjectCode, term, year, remarkVal) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
    const safeId = normID(studentId);
    
    // ถ้าครูเลือกช่องว่าง ให้บันทึกเป็นเครื่องหมาย -
    let finalRemark = (remarkVal === '') ? '-' : remarkVal;

    // 1. ทะลวงชีต Grade_Summary (คอลัมน์ E)
    const gradeSheet = ss.getSheetByName("Grade_Summary");
    if (gradeSheet) {
      if (gradeSheet.getMaxColumns() < 6) gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), 6 - gradeSheet.getMaxColumns());
      
      const data = gradeSheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (normID(data[i][0]) === safeId && String(data[i][1]).trim().toLowerCase() === String(subjectCode).trim().toLowerCase()) {
          // อัปเดตลงคอลัมน์ E (Index 4)
          gradeSheet.getRange(i + 1, 5).setValue(finalRemark); 
          found = true;
          break;
        }
      }
      // ถ้าหาไม่เจอ ให้สร้างบรรทัดใหม่
      if (!found) {
        gradeSheet.appendRow(["'" + studentId, subjectCode, 0, 0, finalRemark, 100]);
      }
    }

    // 2. ทะลวงชีต Score_Database ด้วย
    const scoreSheet = ss.getSheetByName("Score_Database");
    if (scoreSheet) {
      const sData = scoreSheet.getDataRange().getValues();
      let sFound = false;
      for (let i = 1; i < sData.length; i++) {
        if (normID(sData[i][1]) === safeId &&
            String(sData[i][2]).trim().toLowerCase() === String(subjectCode).trim().toLowerCase() &&
            String(sData[i][3]).trim().toLowerCase() === 'remark' &&
            String(sData[i][5]).trim() === String(term).trim() &&
            String(sData[i][6]).trim() === String(year).trim()) {
          
          scoreSheet.getRange(i + 1, 5).setValue(finalRemark);
          sFound = true;
          break;
        }
      }
      if (!sFound && finalRemark !== '-') {
         scoreSheet.appendRow([safeId + "_" + subjectCode + "_remark", "'" + studentId, subjectCode, "remark", finalRemark, term, year]);
      }
    }

    return { success: true, val: remarkVal };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// 📚 15. ระบบงานสารบรรณ (Sarabun System)
// ==========================================

/**
 * ฟังก์ชันขอเลขที่เอกสารใหม่ (เพิ่มฟิลด์เวลา & รองรับกลุ่ม)
 */
function requestSarabunNumber(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Sarabun_Database");

    // 🌟 โครงสร้างใหม่แบบ 16 คอลัมน์ (แทรก DocTime)
    if (!sheet) {
        sheet = ss.insertSheet("Sarabun_Database");
        const headers = [
           "Timestamp", "DocType", "DocNumber", "Subject", "TargetDate", "DocTime", 
           "ActionDate", "DocRefNo", "RefDocDate", "DocFrom", "DocTo", 
           "Assignee", "Requester", "Status", "FileURL", "Year"
        ];
        sheet.appendRow(headers);
        sheet.getRange("A1:P1").setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
        sheet.setFrozenRows(1);
    }

    const config = getSystemConfig();
    const currentYear = String(config.year).trim(); 
    const docType = String(payload.docType).trim(); 
    const amount = parseInt(payload.amount) || 1; 

    const data = sheet.getDataRange().getDisplayValues();
    let lastNumber = 0;

    // สแกนหาเลขล่าสุด
    for (let i = data.length - 1; i >= 1; i--) {
        const rowType = String(data[i][1]).trim();
        const rowYear = String(data[i][15]).trim(); // 🌟 Index เลื่อนไปที่ 15

        if (rowType === docType && rowYear === currentYear) {
            const docNumStr = String(data[i][2]); 
            const numPart = docNumStr.split('/')[0]; 
            lastNumber = parseInt(numPart) || 0;
            break; 
        }
    }

    const timestamp = new Date();
    let startNumber = lastNumber + 1;
    let endNumber = lastNumber + amount;
    let rowsToAppend = [];

    // วนลูปสร้างแถวข้อมูลตามจำนวน (Amount)
    for(let i = 0; i < amount; i++) {
        let currentNum = lastNumber + 1 + i;
        let newDocNumber = `${currentNum}/${currentYear}`;
        
        rowsToAppend.push([
          timestamp,                    // 0: Timestamp
          docType,                      // 1: DocType
          newDocNumber,                 // 2: DocNumber
          payload.subject || "-",       // 3: Subject
          payload.targetDate || "-",    // 4: TargetDate
          payload.docTime || "-",       // 5: DocTime (🌟 เวลาที่เพิ่มมา)
          payload.actionDate || "-",    // 6: ActionDate
          payload.docRefNo || "-",      // 7: DocRefNo
          payload.refDocDate || "-",    // 8: RefDocDate
          payload.docFrom || "-",       // 9: DocFrom
          payload.docTo || "-",         // 10: DocTo
          payload.assignee || "-",      // 11: Assignee
          payload.requester || "Unknown",// 12: Requester
          "ใช้งาน",                      // 13: Status
          "",                           // 14: FileURL
          currentYear                   // 15: Year
        ]);
    }

    if(rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 16).setValues(rowsToAppend);
    }

    SpreadsheetApp.flush(); 

    let resultDocNum = amount > 1 ? `${startNumber}/${currentYear} ถึง ${endNumber}/${currentYear}` : `${startNumber}/${currentYear}`;

    return { 
      status: "success", 
      message: `✅ สำเร็จ! ดำเนินการออกเลข ${amount} รายการ`, 
      docNumber: resultDocNum 
    };

  } catch (e) {
    return { status: "error", message: "คิวเต็ม กรุณากดขอเลขใหม่อีกครั้งครับ" };
  } finally {
    lock.releaseLock(); 
  }
}

/**
 * ฟังก์ชันดึงประวัติการขอเลข (16 คอลัมน์)
 */
function getSarabunHistory(requesterName, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sarabun_Database");
  if (!sheet) return [];

  const config = getSystemConfig();
  const data = sheet.getDataRange().getDisplayValues();
  const results = [];

  for (let i = data.length - 1; i >= 1; i--) {
     const row = data[i];
     const rowRequester = String(row[12]).trim(); 
     const rowYear = String(row[15]).trim();      
     
     if ((role.toUpperCase() === 'ADMIN' || rowRequester === requesterName) && rowYear === String(config.year)) {
        results.push({
           id: i + 1, 
           timestamp: row[0],
           docType: row[1],
           docNumber: row[2],
           subject: row[3],
           targetDate: row[4],
           docTime: row[5],
           actionDate: row[6],
           docRefNo: row[7],
           refDocDate: row[8],
           docFrom: row[9],
           docTo: row[10],
           assignee: row[11],
           requester: row[12],
           status: row[13],
           fileUrl: row[14]
        });
     }
  }
  return results;
}

/**
 * 💾 ฟังก์ชันอัปโหลดไฟล์ของงานสารบรรณ และบันทึกลงฐานข้อมูล
 */
function uploadSarabunFile(id, base64Data, filename, docNumber) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    // 1. อัปโหลดไฟล์ขึ้น Google Drive
    const safeDocNum = String(docNumber).replace(/\//g, '-'); 
    const safeFilename = `Sarabun_${safeDocNum}_${filename}`;
    const fileUrl = uploadFileToDrive(base64Data, safeFilename);
    
    if (fileUrl.startsWith("Error")) throw new Error(fileUrl);

    // 2. บันทึก URL ไฟล์ ลงในฐานข้อมูล
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sarabun_Database");
    if (!sheet) throw new Error("ไม่พบฐานข้อมูล");

    // 🌟 ท่าไม้ตาย: ใช้ getDisplayValue() บังคับอ่านค่าแบบตัวอักษรเพียวๆ
    const targetDocNum = String(docNumber).trim();
    const currentDocNum = String(sheet.getRange(parseInt(id), 3).getDisplayValue()).trim(); 

    if (currentDocNum === targetDocNum) {
        sheet.getRange(parseInt(id), 15).setValue(fileUrl); // คอลัมน์ 15 (O) คือ FileURL
    } else {
        // ถ้าคลาดเคลื่อน ให้สแกนหาใหม่ โดยใช้ getDisplayValues() กวาดอ่านแบบตัวอักษร
        const allData = sheet.getDataRange().getDisplayValues();
        let found = false;
        for(let i = 1; i < allData.length; i++) {
            if(String(allData[i][2]).trim() === targetDocNum) {
                sheet.getRange(i + 1, 15).setValue(fileUrl);
                found = true;
                break;
            }
        }
        if(!found) throw new Error(`ไม่พบเอกสารเลขที่ ${targetDocNum} ในระบบ`);
    }
    
    return { status: "success", message: "แนบไฟล์เสร็จสมบูรณ์" };

  } catch(e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}