/**
 * 🏫 PSSMS - Phuphrabat Smart School Management System
 * ระบบบริหารจัดการสถานศึกษา 4 ฝ่าย (Single Page Application)
 * พัฒนาโดย: ครูน๊อต ศิกษก เดินรีบรัมย์
 * Updated: 2026-02-19 | ปรับปรุงโครงสร้างฐานข้อมูล (Level, Room, Location)
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
  let config = { term: "1", year: "2568", termStart: "", termEnd: "" };
  if (!sheet) return config;
  
  const data = sheet.getDataRange().getValues();
  data.forEach(row => {
    if(row[0] === "Current_Term") config.term = String(row[1]);
    if(row[0] === "Current_Year") config.year = String(row[1]);
    if(row[0] === "Term_Start") config.termStart = String(row[1]);
    if(row[0] === "Term_End") config.termEnd = String(row[1]);
  });
  return config;
}

function saveSystemConfig(term, year, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("System_Settings") || ss.insertSheet("System_Settings");
  const settings = { "Current_Term": term, "Current_Year": year, "Term_Start": startDate, "Term_End": endDate };
  
  sheet.clear();
  sheet.appendRow(["Key", "Value"]);
  for (let key in settings) { sheet.appendRow([key, settings[key]]); }
  return { status: 'success', message: '✅ บันทึกระยะเวลาภาคเรียนเรียบร้อย' };
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
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance_Database");
    const ts = new Date();
    const rows = list.map(item => [
      ts, item.date, item.term, item.year, item.subjectCode, item.subjectName,
      item.className, item.period, item.studentId, item.studentName, item.status, 
      item.teacherId, `${item.date}|${item.subjectCode}|${item.className}|${item.period}`
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { status: "success", message: "✅ เช็คชื่อเรียบร้อย" };
  } catch (e) { return { status: "error", message: e.message }; }
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Academic_Records");
  const config = getSystemConfig();
  sheet.appendRow([
    new Date(), record.date, config.term, config.year, record.subjectCode, record.subjectName, 
    record.className, record.period, record.topic, record.totalPresent, record.totalAbsent, 
    record.totalLeave, record.teacherId, record.signature, 
    `${record.date}|${record.subjectCode}|${record.className}|${record.period}`
  ]);
  return { status: "success", message: "✅ บันทึกข้อมูลการสอนเรียบร้อยแล้ว" };
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

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (!row[1]) continue; 

    let rowDateStr = "";
    if (row[1] instanceof Date) {
      rowDateStr = Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd");
    } else {
      rowDateStr = String(row[1]).substring(0, 10);
    }

    const rowSub = String(row[4]).trim().replace(/\s/g, '');
    const rowClass = String(row[6]).trim().replace(/\s/g, '');

    if (rowDateStr === cleanTargetDate && rowSub === cleanSub && rowClass === cleanClass) {
      
      const rawID = String(row[8]).trim();
      const idNoZero = String(parseInt(rawID)); 
      
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
    // Update Header ใหม่
    { name: "Timetable_Database", headers: ["SubjectCode", "SubjectName", "Level", "Room", "Location", "TeacherID", "Day", "Period", "Term", "Year"] }
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

function getSubjectConfig(subjectCode, className, term, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subject_Config");
  if(!sheet) return null;
  
  const data = sheet.getDataRange().getDisplayValues(); 
  
  const targetSubj = String(subjectCode).trim();
  const targetClass = String(className).trim();
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();
  
  for(let i = 1; i < data.length; i++) {
    const rSubj = String(data[i][1]).trim();
    const rClass = String(data[i][2]).trim();
    const rTerm = String(data[i][3]).trim();
    const rYear = String(data[i][4]).trim();

    if(rSubj === targetSubj && rClass === targetClass && rTerm === targetTerm && rYear === targetYear) {
      return {
        ratio: data[i][5], 
        indicators: JSON.parse(data[i][6] || '[]')
      };
    }
  }
  return null; 
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

function getAllInOneScoreGridData(subjectCode, className, term, year) {
  let config = getSubjectConfig(subjectCode, className, term, year);
  if (!config) config = { ratio: "70:10:20", indicators: [{name: "คะแนนเก็บ 1", score: 70}] }; 

  const students = getStudentsByClass(className);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
  const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

  const sheetScore = ss.getSheetByName("Score_Database");
  const scoreData = sheetScore ? sheetScore.getDataRange().getDisplayValues() : [];
  const existingScores = {};
  
  for(let i = 1; i < scoreData.length; i++) {
    const row = scoreData[i];
    if(normStr(row[2]) === normStr(subjectCode) && normStr(row[5]) === normStr(term) && normStr(row[6]) === normStr(year)) {
      const stdKey = normID(row[1]);
      const indKey = normStr(row[3]);
      existingScores[`${stdKey}_${indKey}`] = row[4];
    }
  }

  const qualSheet = ss.getSheetByName("Qualitative_Assess");
  const qualData = qualSheet ? qualSheet.getDataRange().getDisplayValues() : [];
  const existingQuals = {};
  
  for(let i = 1; i < qualData.length; i++) {
    const row = qualData[i];
    if(normStr(row[1]) === normStr(subjectCode) && normStr(row[2]) === normStr(term) && normStr(row[3]) === normStr(year)) {
      existingQuals[normID(row[0])] = { read: row[4], char: row[5], comp: row[6] };
    }
  }

  let attStats = {};
  try {
    const report = getSemesterReport(subjectCode, className, term, year);
    if(report && report.students) {
      report.students.forEach(s => { 
         attStats[normID(s.id)] = parseFloat(s.percent); 
      });
    }
  } catch(e) { console.log("Auto-MS Error: " + e); }

  return { config: config, students: students, existingScores: existingScores, existingQuals: existingQuals, attStats: attStats };
}

function saveAllInOneWithConfig(payload) {
  saveSubjectConfig({
    subjectCode: payload.subjectCode, className: payload.className,
    term: payload.term, year: payload.year, teacherId: payload.teacherId,
    formative: payload.newConfig.formative, midterm: payload.newConfig.midterm,
    final: payload.newConfig.final, indicators: payload.newConfig.indicators
  });
  return saveAllInOneScores(payload);
}

// 🌟 ฟังก์ชันบันทึกขั้นสูง (ความเร็วแสง - Batch Array Update)
function saveAllInOneScores(payload) {
  const { subjectCode, term, year, scoreRecords, qualRecords, gradeRecords } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
  const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

  // ==============================
  // 1. บันทึกคะแนน (Score_Database)
  // ==============================
  const sheetScore = ss.getSheetByName("Score_Database");
  let scoreData = sheetScore.getDataRange().getValues(); // ดึงมาเป็น Array ทั้งตาราง
  
  const scoreMap = {};
  scoreRecords.forEach(r => {
     const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.indicatorId)}_${normStr(r.term)}_${normStr(r.year)}`;
     scoreMap[uid] = r;
  });

  let scoreUpdated = false;
  // ⚡ อัปเดตในความจำให้เสร็จก่อน
  for(let i = 1; i < scoreData.length; i++) {
    const row = scoreData[i];
    const uid = `${normID(row[1])}_${normStr(row[2])}_${normStr(row[3])}_${normStr(row[5])}_${normStr(row[6])}`;
    if(scoreMap[uid]) {
       if (String(scoreData[i][4]) !== String(scoreMap[uid].score)) {
           scoreData[i][4] = scoreMap[uid].score; // เขียนทับค่าใน Array
           scoreUpdated = true;
       }
       scoreMap[uid].processed = true; 
    }
  }
  // ⚡ สาดข้อมูลทั้งหมดกลับลง Sheet รวดเดียว
  if(scoreUpdated) sheetScore.getRange(1, 1, scoreData.length, scoreData[0].length).setValues(scoreData);

  // ⚡ แถวไหนเป็นข้อมูลใหม่ เอาไปต่อท้าย
  const newScores = [];
  for (let uid in scoreMap) {
     if (!scoreMap[uid].processed) {
         const r = scoreMap[uid];
         newScores.push([uid, "'" + r.studentId, r.subjectCode, r.indicatorId, r.score, r.term, r.year]);
     }
  }
  if (newScores.length > 0) sheetScore.getRange(sheetScore.getLastRow() + 1, 1, newScores.length, newScores[0].length).setValues(newScores);

  // ==============================
  // 2. บันทึกคุณลักษณะ (Qualitative_Assess)
  // ==============================
  const qualSheet = ss.getSheetByName("Qualitative_Assess");
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
       if (String(qualData[i][4]) !== String(qualMap[uid].read) || String(qualData[i][5]) !== String(qualMap[uid].char) || String(qualData[i][6]) !== String(qualMap[uid].comp)) {
           qualData[i][4] = qualMap[uid].read;
           qualData[i][5] = qualMap[uid].char;
           qualData[i][6] = qualMap[uid].comp;
           qualUpdated = true;
       }
       qualMap[uid].processed = true;
    }
  }
  if(qualUpdated) qualSheet.getRange(1, 1, qualData.length, qualData[0].length).setValues(qualData);
  
  const newQuals = [];
  for (let uid in qualMap) {
     if(!qualMap[uid].processed) {
        const r = qualMap[uid];
        newQuals.push(["'" + r.studentId, r.subjectCode, r.term, r.year, r.read, r.char, r.comp]);
     }
  }
  if(newQuals.length > 0) qualSheet.getRange(qualSheet.getLastRow() + 1, 1, newQuals.length, newQuals[0].length).setValues(newQuals);

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