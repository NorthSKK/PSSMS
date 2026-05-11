/**
 * 🏫 PSSMS - Phuphrabat Smart School Management System
 * ระบบบริหารจัดการสถานศึกษา 4 ฝ่าย (Single Page Application)
 * พัฒนาโดย: ครูน๊อต ศิกษก เดินรีบรัมย์
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

function getAvailableTerms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("System_Settings");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const terms = [];
  data.forEach(row => {
    if (row[0] === "TermData") {
      const parts = String(row[1]).split("_"); // "1_2568"
      if (parts.length === 2) {
        terms.push({ term: parts[0], year: parts[1], key: String(row[1]) });
      }
    }
  });
  // Sort: year desc, term desc (newest first)
  terms.sort((a, b) => {
    if (b.year !== a.year) return parseInt(b.year) - parseInt(a.year);
    return parseInt(b.term) - parseInt(a.term);
  });
  return terms;
}

// ==========================================
// 3. DASHBOARD & STATS
// ==========================================

var NOTION_TOKEN = 'ntn_K30250483172wMxPDJaiHUmHF5DRmU3aNj7y5RuglMk6iq'; 
var DATABASE_ID = '1b4a3504c04d48c182068d064c38d1e1'; 
var PROJECT_ID = '1920b44e-92fd-8013-828a-c06028c1c231';

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === "create") {
    const responseText = sendTaskToNotion(data.taskName);
    const result = JSON.parse(responseText);
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "id": result.id })).setMimeType(ContentService.MimeType.JSON);
  } else if (data.action === "update") {
    updateTaskStatus(data.pageId, data.isDone);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }
}

function sendTaskToNotion(taskName) {
  const url = 'https://api.notion.com/v1/pages';
  const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const payload = { "parent": { "database_id": DATABASE_ID }, "icon": { "type": "emoji", "emoji": "✏️" }, "properties": { "Name": { "title": [{ "text": { "content": taskName } }] }, "Date": { "date": { "start": today } }, "Status": { "status": { "name": "Not started" } }, "Projects": { "relation": [{ "id": PROJECT_ID }] } } };
  const options = { "method": "post", "headers": { "Authorization": "Bearer " + NOTION_TOKEN, "Notion-Version": "2022-06-28", "Content-Type": "application/json" }, "payload": JSON.stringify(payload), "muteHttpExceptions": true };
  const response = UrlFetchApp.fetch(url, options);
  return response.getContentText();
}

function updateTaskStatus(pageId, isDone) {
  const url = 'https://api.notion.com/v1/pages/' + pageId;
  const statusName = (isDone === true || isDone === undefined) ? "Done" : "Not started";
  const payload = { "properties": { "Status": { "status": { "name": statusName } }, "Archive": { "checkbox": (isDone === true || isDone === undefined) } } };
  const options = { "method": "patch", "headers": { "Authorization": "Bearer " + NOTION_TOKEN, "Notion-Version": "2022-06-28", "Content-Type": "application/json" }, "payload": JSON.stringify(payload), "muteHttpExceptions": true };
  UrlFetchApp.fetch(url, options);
}

function getTodoList(userId) {
  if (!userId) return "[]";
  try {
    const rawData = PropertiesService.getScriptProperties().getProperty('TODO_' + userId);
    return rawData ? rawData : "[]";
  } catch(e) { return "[]"; }
}

function saveTodoList(userId, todosJSON) {
  if (!userId || !todosJSON) return false;
  try { PropertiesService.getScriptProperties().setProperty('TODO_' + userId, todosJSON); return true; } 
  catch(e) { return false; }
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
  sheet.appendRow([form.username, form.password, form.fullname, form.role, form.dept, form.email, config.year, form.status || "ปกติ"]);
  return {status: 'success', message: 'เพิ่มผู้ใช้งานสำเร็จ'};
}

function editUser(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(form.username)){
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
      const key = `${tCode}_${tClassID}`; 

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

  for (let i = 1; i < attData.length; i++) {
    const row = attData[i];
    if (!row[1]) continue;
    
    const rowTerm = String(row[2]).trim();
    const rowYear = String(row[3]).trim();
    if(rowTerm !== targetTerm || rowYear !== targetYear) continue;

    const rowSub = normalize(row[4]);
    const rowClass = normalize(row[6]);
    const key = `${rowSub}_${rowClass}`;

    if (teacherClasses[key]) {
      const stdID = String(row[8]).trim();
      const stdName = row[9];
      const status = row[10];
      const sessionID = String(row[12]).trim() || (row[1] + "_" + row[7]);

      teacherClasses[key].sessions.add(sessionID);

      if (!teacherClasses[key].students[stdID]) {
        teacherClasses[key].students[stdID] = { name: stdName, records: {} };
      }
      teacherClasses[key].students[stdID].records[sessionID] = status;
    }
  }

  const weeksPerTerm = 20;
  let critical = []; 
  let ms = [];       
  let risk = [];     

  for (const key in teacherClasses) {
    const cls = teacherClasses[key];
    const currentTotalTaught = cls.sessions.size;
    
    if (currentTotalTaught === 0) continue; 

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

      const totalMissed = absent + leave;
      const percent = ((totalCoursePeriods - totalMissed) / totalCoursePeriods) * 100;

      if (percent <= 85) {
        const studentData = {
          id: stdID, name: student.name, subjectCode: cls.rawCode, subjectName: cls.rawName,
          className: cls.rawClassID, present, late, leave, absent, percent: percent.toFixed(2), taught: currentTotalTaught
        };
        if (percent < 60) critical.push(studentData);
        else if (percent < 80) ms.push(studentData);
        else risk.push(studentData);
      }
    }
  }

  critical.sort((a, b) => parseFloat(a.percent) - parseFloat(b.percent));
  ms.sort((a, b) => parseFloat(a.percent) - parseFloat(b.percent));
  risk.sort((a, b) => parseFloat(a.percent) - parseFloat(b.percent));

  return { critical, ms, risk };
}

function getTeacherTimetableByDate(teacherId, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  const config = getSystemConfig();
  const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  
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

  return data.slice(1).map(r => {
      const tTeacherID = String(r[5]).trim().toLowerCase();
      const tDay = String(r[6]).trim();
      const tTerm = String(r[8]).trim();
      const tYear = String(r[9]).trim();
      
      if (tTeacherID === searchTeacherId && tDay === today && tTerm === searchTerm && tYear === searchYear) {
         const tLevel = String(r[2]).trim();
         const tRoom = String(r[3]).trim();
         const tLoc = String(r[4]).trim();
         const tClassID = `${tLevel}/${tRoom}`; 
         return [r[0], r[1], tClassID, tRoom, tLoc, r[7], r[6]]; 
      }
      return null;
  }).filter(item => item !== null);
}

function getTeacherTimetableWithStatus(teacherId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSystemConfig();
  const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  const now = new Date();
  const today = days[now.getDay()];
  const todayStr = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd");

  const ttSheet = ss.getSheetByName("Timetable_Database");
  if (!ttSheet) return [];

  const ttData = ttSheet.getDataRange().getValues();
  const searchTeacherId = String(teacherId).trim().toLowerCase();
  const searchTerm = String(config.term).trim();
  const searchYear = String(config.year).trim();

  const items = ttData.slice(1).map(r => {
    if (String(r[5]).trim().toLowerCase() !== searchTeacherId) return null;
    if (String(r[6]).trim() !== today) return null;
    if (String(r[8]).trim() !== searchTerm || String(r[9]).trim() !== searchYear) return null;
    const tClassID = `${String(r[2]).trim()}/${String(r[3]).trim()}`;
    return [r[0], r[1], tClassID, r[3], r[4], r[7], r[6]];
  }).filter(Boolean);

  if (items.length === 0) return [];

  const fmtDate = (v) => v instanceof Date
    ? Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd")
    : String(v).substring(0, 10);

  // ตรวจ Attendance_Database (วิชาปกติ): Date[1], SubjectCode[4], Class[6], Period[7], TeacherID[11]
  const attSheet = ss.getSheetByName("Attendance_Database");
  const checkedSet = new Set();
  if (attSheet) {
    attSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (fmtDate(r[1]) === todayStr && String(r[11]).trim().toLowerCase() === searchTeacherId) {
        checkedSet.add(`${String(r[4]).trim()}_${String(r[6]).trim()}_${String(r[7]).trim()}`);
      }
    });
  }

  // ตรวจ Morning_Activity (HR): Date[1], Class[4], TeacherID[10]
  const mrSheet = ss.getSheetByName("Morning_Activity");
  const hrCheckedSet = new Set();
  if (mrSheet) {
    mrSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (fmtDate(r[1]) === todayStr && String(r[10]).trim().toLowerCase() === searchTeacherId) {
        hrCheckedSet.add(String(r[4]).trim());
      }
    });
  }

  return items.map(item => {
    const isHR = String(item[0]).toUpperCase() === 'HR' || String(item[1]).includes('โฮมรูม');
    const checked = isHR
      ? hrCheckedSet.has(item[2])
      : checkedSet.has(`${String(item[0]).trim()}_${item[2]}_${String(item[5]).trim()}`);
    return [...item, checked]; // item[7] = boolean
  });
}

function normalizeClassName(c) {
  // ทำให้รหัสห้องเทียบกันได้: ตัด space, ตัด "ม." ออก → "ม.1/1" และ "1/1" ตรงกัน
  return String(c || "").replace(/\s+/g, '').replace(/^ม\.?/i, '').toLowerCase();
}

function getStudentsByClass(className, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const config = getSystemConfig();
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  const targetClass = normalizeClassName(className);
  const targetYear = String(year || config.year).trim();

  // 1. ค้นหาในฐานข้อมูลปัจจุบันก่อน
  let filtered = data.slice(1).filter(r => {
    const rowRole = String(r[3]).trim().toLowerCase();
    const rowClass = normalizeClassName(r[4]);
    const rowYear = String(r[6]).trim();
    let rowStatus = "ปกติ";
    if (r.length > 7 && String(r[7]).trim() !== "") rowStatus = String(r[7]).trim();

    const isStudent = (rowRole === 'student' || rowRole === 'นักเรียน');
    const isClassMatch = (rowClass === targetClass);
    const isYearMatch = (rowYear === targetYear || rowYear === "");
    const isStatusNormal = (rowStatus === 'ปกติ');

    return isStudent && isClassMatch && isYearMatch && isStatusNormal;
  });

  // 2. fallback: ค้นใน User_Database โดยไม่กรอง year (รองรับช่วงเปลี่ยนปีการศึกษา)
  if (filtered.length === 0) {
      filtered = data.slice(1).filter(r => {
          const rowRole = String(r[3]).trim().toLowerCase();
          const rowClass = normalizeClassName(r[4]);
          let rowStatus = "ปกติ";
          if (r.length > 7 && String(r[7]).trim() !== "") rowStatus = String(r[7]).trim();
          return (rowRole === 'student' || rowRole === 'นักเรียน') && rowClass === targetClass && rowStatus === 'ปกติ';
      });
  }

  // 3. ท่าไม้ตาย: ขุดจาก User_History_Database (กรณีดูย้อนหลังปีที่นักเรียนจบ/ย้ายไปแล้ว)
  if (filtered.length === 0) {
      const histSheet = ss.getSheetByName("User_History_Database");
      if (histSheet && histSheet.getLastRow() > 1) {
          const histData = histSheet.getDataRange().getDisplayValues();
          filtered = histData.slice(1).filter(r => {
              const rowRole = String(r[3]).trim().toLowerCase();
              const rowClass = normalizeClassName(r[4]);
              const rowYear = String(r[6]).trim();
              let rowStatus = "ปกติ";
              if (r.length > 7 && String(r[7]).trim() !== "") rowStatus = String(r[7]).trim();

              const isStudent = (rowRole === 'student' || rowRole === 'นักเรียน');
              const isClassMatch = (rowClass === targetClass);
              const isYearMatch = (rowYear === targetYear);
              const isStatusValid = (rowStatus === 'ปกติ' || rowStatus === 'จบการศึกษา');

              return isStudent && isClassMatch && isYearMatch && isStatusValid;
          });
      }
  }

  return filtered;
}

// Debug: เรียกจาก frontend เพื่อดูว่านักเรียนใน sheet มี class/role/status อะไรบ้าง
function debugStudentsByClass(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  const config = getSystemConfig();
  if (!sheet) return { error: "no User_Database" };

  const data = sheet.getDataRange().getDisplayValues();
  const targetClass = normalizeClassName(className);

  const studentRows = data.slice(1).filter(r => {
    const rowRole = String(r[3]).trim().toLowerCase();
    return rowRole === 'student' || rowRole === 'นักเรียน';
  });

  const classes = {};
  studentRows.forEach(r => {
    const c = String(r[4] || "").trim();
    classes[c] = (classes[c] || 0) + 1;
  });

  const matched = studentRows.filter(r => normalizeClassName(r[4]) === targetClass);

  // เช็ค Timetable_Database ด้วย
  const tt = ss.getSheetByName("Timetable_Database");
  let timetableInfo = { error: "no Timetable_Database" };
  if (tt) {
    const ttData = tt.getDataRange().getDisplayValues();
    const yearTermCounts = {};
    ttData.slice(1).forEach(r => {
      const term = String(r[8] || "").trim();
      const year = String(r[9] || "").trim();
      const key = `${term}_${year}`;
      yearTermCounts[key] = (yearTermCounts[key] || 0) + 1;
    });
    const matchActive = ttData.slice(1).filter(r =>
      String(r[8]).trim() === String(config.term).trim() &&
      String(r[9]).trim() === String(config.year).trim()
    );
    timetableInfo = {
      totalRows: ttData.length - 1,
      byTermYear: yearTermCounts,
      matchActiveTermYear: matchActive.length
    };
  }

  // เช็ค status breakdown ของ students ในห้อง
  const statusBreakdown = {};
  matched.forEach(r => {
    const s = r.length > 7 ? String(r[7] || "(empty)").trim() : "(no col)";
    statusBreakdown[s] = (statusBreakdown[s] || 0) + 1;
  });

  // ดู year breakdown ของห้องนี้
  const yearBreakdown = {};
  matched.forEach(r => {
    const y = String(r[6] || "(empty)").trim();
    yearBreakdown[y] = (yearBreakdown[y] || 0) + 1;
  });

  return {
    activeYear: config.year,
    activeTerm: config.term,
    targetClass: className,
    normalizedTarget: targetClass,
    totalStudents: studentRows.length,
    classesInDB: classes,
    matchedCount: matched.length,
    matchedYearBreakdown: yearBreakdown,
    matchedStatusBreakdown: statusBreakdown,
    timetable: timetableInfo,
    sampleMatched: matched.slice(0, 3).map(r => ({
      id: r[0], name: r[2], role: r[3], class: r[4], year: r[6], status: r[7] || "(empty)"
    }))
  };
}

// Debug: ดูว่า Timetable_Database มี HR (โฮมรูม) entries สำหรับครูคนไหนใน active term/year
function debugHomeroom(teacherId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tt = ss.getSheetByName("Timetable_Database");
  const config = getSystemConfig();
  if (!tt) return { error: "no Timetable_Database" };

  const data = tt.getDataRange().getDisplayValues();
  const activeTerm = String(config.term).trim();
  const activeYear = String(config.year).trim();

  // HR rows ทั้งหมดใน Timetable
  const allHR = data.slice(1).filter(r => {
    const code = String(r[0] || "").toUpperCase().trim();
    const name = String(r[1] || "");
    return code === 'HR' || name.includes('โฮมรูม');
  }).map(r => ({
    subjectCode: r[0], subjectName: r[1], level: r[2], room: r[3], loc: r[4],
    teacherId: r[5], day: r[6], period: r[7], term: r[8], year: r[9]
  }));

  // HR ของ active term/year
  const activeHR = allHR.filter(r => String(r.term).trim() === activeTerm && String(r.year).trim() === activeYear);

  // ถ้าระบุ teacherId มา → กรองเพิ่ม
  let teacherHR = null;
  if (teacherId) {
    teacherHR = activeHR.filter(r => String(r.teacherId).trim().toLowerCase() === String(teacherId).trim().toLowerCase());
  }

  return {
    activeTerm, activeYear,
    queryTeacherId: teacherId || "(not specified)",
    allHRCount: allHR.length,
    activeTermYearHRCount: activeHR.length,
    activeTermYearHR: activeHR.slice(0, 20),
    teacherSpecificHR: teacherHR
  };
}

function debugTimetableResult(teacherId, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  const config = getSystemConfig();
  const days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

  let targetDateObj = dateStr ? new Date(dateStr) : new Date();
  const targetDayName = days[targetDateObj.getDay()];
  const searchTeacherId = String(teacherId).trim().toLowerCase();
  const searchTerm = String(config.term).trim();
  const searchYear = String(config.year).trim();

  if (!sheet) return { error: "no Timetable_Database" };
  const data = sheet.getDataRange().getValues();

  // ผลที่ function จริงจะ return
  const matched = data.slice(1).filter(r => {
    return String(r[5]).trim().toLowerCase() === searchTeacherId
      && String(r[6]).trim() === targetDayName
      && String(r[8]).trim() === searchTerm
      && String(r[9]).trim() === searchYear;
  }).map(r => ({
    subjectCode: r[0], subjectName: r[1], level: r[2], room: r[3],
    day: r[6], period: r[7], term: r[8], year: r[9]
  }));

  // rows ของ teacher นี้ทั้งหมด (ไม่กรอง day/term/year) เพื่อ compare
  const teacherAllRows = data.slice(1).filter(r =>
    String(r[5]).trim().toLowerCase() === searchTeacherId
  ).map(r => ({ subjectCode: r[0], day: r[6], period: r[7], term: r[8], year: r[9] }));

  return {
    input: { teacherId, dateStr },
    computed: { targetDayName, searchTerm, searchYear },
    matchedCount: matched.length,
    matched,
    teacherAllRowsCount: teacherAllRows.length,
    teacherAllRows: teacherAllRows.slice(0, 30)
  };
}

function updateAttendanceStatus(studentId, sessionID, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };

  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8]).trim() === String(studentId).trim() && String(data[i][12]).trim() === String(sessionID).trim()) {
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
      history.push({ date: row[1], period: row[7], status: row[10], sessionId: row[12] });
    }
  });
  return history.reverse(); 
}

function saveAttendanceBatch(list) {
  if (!list || list.length === 0) return { status: "error", message: "ไม่มีข้อมูล" };
  const first = list[0];

  if (!verifyTeacherPermission(first.teacherId, first.subjectCode, first.className, first.term, first.year)) {
     return { status: "error", message: "❌ ความปลอดภัย: คุณไม่มีสิทธิ์เช็คชื่อในวิชาและห้องนี้!" };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance_Database");
    const ts = new Date();
    const rows = list.map(item => [
      ts, item.date, item.term, item.year, item.subjectCode, item.subjectName,
      item.className, item.period, item.studentId, item.studentName, item.status, 
      item.teacherId, `${item.date}|${item.subjectCode}|${item.className}|${item.period}`
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush(); 
    return { status: "success", message: "✅ เช็คชื่อเรียบร้อย" };
  } catch (e) { 
    return { status: "error", message: "คิวบันทึกเต็ม กรุณากดบันทึกอีกครั้งครับ" }; 
  } finally {
    lock.releaseLock(); 
  }
}

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
  
  let periodsPerWeek = 0;
  for (let i = 1; i < timeData.length; i++) {
    const row = timeData[i];
    const tLevel = String(row[2]).trim();
    const tRoom = String(row[3]).trim();
    const tClassID = normalize(`${tLevel}/${tRoom}`); 

    if (normalize(row[0]) === cleanSub && tClassID === cleanClass && String(row[8]).trim() === targetTerm && String(row[9]).trim() === targetYear) {
      periodsPerWeek++;
    }
  }

  if (periodsPerWeek === 0) periodsPerWeek = 3;

  const weeksPerTerm = 20; 
  const totalCoursePeriods = periodsPerWeek * weeksPerTerm; 
  const maxAbsenceQuota = Math.floor(totalCoursePeriods * 0.2); 

  const studentDataMap = {}; 
  const studentInfo = {}; 
  const sessionDetails = {}; 

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
      
      if (!sessionDetails[sessionID]) {
          let d;
          if (row[1] instanceof Date) d = row[1];
          else d = new Date(row[1]);
          
          if (!isNaN(d.getTime())) {
              let months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
              sessionDetails[sessionID] = {
                  id: sessionID, rawDate: d.getTime(), month: months[d.getMonth()], date: d.getDate(), period: String(row[7]).trim()
              };
          } else {
              sessionDetails[sessionID] = { id: sessionID, rawDate: 0, month: "-", date: "-", period: String(row[7]).trim() };
          }
      }

      if (!studentInfo[stdID]) studentInfo[stdID] = stdName;
      if (!studentDataMap[stdID]) studentDataMap[stdID] = {};
      studentDataMap[stdID][sessionID] = status;
    }
  }
  
  const currentTotalTaught = Object.keys(sessionDetails).length;

  let sessionsList = Object.values(sessionDetails).sort((a, b) => a.rawDate - b.rawDate);
  let currentWeek = 1; let pCount = 0;
  sessionsList.forEach(s => {
      s.week = currentWeek;
      pCount++;
      if (pCount >= periodsPerWeek) { pCount = 0; currentWeek++; }
  });

  const reportData = Object.keys(studentDataMap).map(stdID => {
    let present = 0, late = 0, leave = 0, absent = 0;
    const records = studentDataMap[stdID];
    
    for (const sessKey in records) {
      const s = records[sessKey];
      if (s === 'มา') present++; else if (s === 'สาย') late++; else if (s === 'ลา') leave++; else if (s === 'ขาด') absent++;
    }

    const totalMissed = absent + leave; 
    let percent = ((totalCoursePeriods - totalMissed) / totalCoursePeriods) * 100;
    
    return { 
      id: stdID, name: studentInfo[stdID], present, late, leave, absent, 
      percent: percent.toFixed(2), currentTotalTaught, records: records 
    };
  });

  return {
    students: reportData.sort((a, b) => a.id.localeCompare(b.id)),
    meta: { periodsPerWeek, weeksPerTerm, totalCoursePeriods, maxAbsenceQuota, currentTotalTaught, sessionsList: sessionsList }
  };
}

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
          rawCode: row[0], rawName: row[1], rawClassID: `${String(row[2]).trim()}/${String(row[3]).trim()}`,
          periodsPerWeek: 0, sessions: new Set(), students: {}
        };
      }
      teacherClasses[key].periodsPerWeek++;
    }
  }

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
      teacherClasses[key].students[stdID].records[sessionID] = row[10]; 
    }
  }

  const weeksPerTerm = 20;
  let allStudents = [];

  for (const key in teacherClasses) {
    const cls = teacherClasses[key];
    const currentTotalTaught = cls.sessions.size;
    
    if (currentTotalTaught === 0) continue; 

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
        id: stdID, name: student.name, subjectCode: cls.rawCode, subjectName: cls.rawName,
        className: cls.rawClassID, present, late, leave, absent, percent: percent.toFixed(2),
        taught: currentTotalTaught, totalCoursePeriods: totalCoursePeriods, remainingQuota: remainingQuota
      });
    }
  }

  allStudents.sort((a, b) => {
    if (a.subjectCode !== b.subjectCode) return a.subjectCode.localeCompare(b.subjectCode);
    if (a.className !== b.className) return a.className.localeCompare(b.className);
    return a.id.localeCompare(b.id);
  });

  return allStudents;
}

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
    
    const tCode = row[0].trim();
    const tName = row[1].trim();
    const tLevel = String(row[2]).trim();
    const tRoom = String(row[3]).trim();
    const tLoc = String(row[4]).trim();
    
    const tClassID = `${tLevel}/${tRoom}`; 
    const tDisplay = `${tClassID} (${tLoc})`; 
    
    const tTeacherID = String(row[5]).trim().toLowerCase(); 
    const tTerm = String(row[8]).trim(); 
    const tYear = String(row[9]).trim(); 

    let isOwner = false;
    
    if (userRole && userRole.toUpperCase() === 'ADMIN') {
      isOwner = true; 
    } else {
      if (tTeacherID === searchUserId) isOwner = true;
      else if (tTeacherID === "teacher" + searchUserId) isOwner = true;
      else if (searchUserId === "teacher" + tTeacherID) isOwner = true;
      else if (tTeacherID.replace(/\D/g,'') !== "" && tTeacherID.replace(/\D/g,'') === searchUserId.replace(/\D/g,'')) isOwner = true;
    }

    if (isOwner && tTerm === searchTerm && tYear === searchYear) {
      const key = `${tCode}-${tClassID}`;
      if (!uniqueKeys.has(key)) {
        uniqueKeys.add(key);
        subjects.push([tCode, tName, tClassID, tDisplay]); 
      }
    }
  }
  
  return subjects;
}

function saveLessonRecord(record) {
  if (!verifyTeacherPermission(record.teacherId, record.subjectCode, record.className, record.term, record.year)) {
     return { status: "error", message: "❌ ความปลอดภัย: คุณไม่มีสิทธิ์บันทึกข้อมูลวิชานี้!" };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Academic_Records");
    const config = getSystemConfig();
    sheet.appendRow([
      new Date(), record.date, config.term, config.year, record.subjectCode, record.subjectName, 
      record.className, record.period, record.topic, record.totalPresent, record.totalAbsent, 
      record.totalLeave, record.teacherId, record.signature, 
      `${record.date}|${record.subjectCode}|${record.className}|${record.period}`
    ]);
    SpreadsheetApp.flush(); 
    return { status: "success", message: "✅ บันทึกข้อมูลการสอนเรียบร้อยแล้ว" };
  } catch (e) {
    return { status: "error", message: "คิวบันทึกเต็ม กรุณากดบันทึกอีกครั้งครับ" };
  } finally {
    lock.releaseLock(); 
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

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (!row[1]) continue; 

    let rowDateStr = "";
    if (row[1] instanceof Date) {
      rowDateStr = Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd");
    } else {
      rowDateStr = String(row[1]).substring(0, 10);
    }

    if (rowDateStr < cleanTargetDate) break;

    const rowSub = String(row[4]).trim().replace(/\s/g, '');
    const rowClass = String(row[6]).trim().replace(/\s/g, '');

    if (rowDateStr === cleanTargetDate && rowSub === cleanSub && rowClass === cleanClass) {
      const rawID = String(row[8]).trim();
      const idNoZero = String(parseInt(rawID, 10)); 
      
      if (!uniqueHistory[idNoZero]) {
        uniqueHistory[idNoZero] = {
          studentId: rawID, cleanId: idNoZero, status: row[10], period: row[7], sessionId: row[12], studentName: row[9] 
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
      try { dateKey = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch (e) { continue; }

      if (!sessionMap[dateKey]) {
        sessionMap[dateKey] = { date: dateKey, displayDate: Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "dd/MM/yyyy"), period: row[7], students: new Set() };
      }
      sessionMap[dateKey].students.add(String(row[8]).trim());
    }
  }

  return Object.values(sessionMap).map(s => ({ date: s.date, displayDate: s.displayDate, period: s.period, count: s.students.size })).sort((a, b) => b.date.localeCompare(a.date));
}

function updateAttendanceBatch(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };

  const data = sheet.getDataRange().getValues();
  const updateMap = {};
  list.forEach(item => { updateMap[String(item.studentId).trim()] = item.status; });

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

function getMassiveAttendanceGrid(subjectCode, className, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName("Attendance_Database");
  const students = getStudentsByClass(className);
  
  const cleanSub = String(subjectCode).replace(/[^a-zA-Z0-9ก-๙]/g, '');
  const cleanClass = String(className).replace(/[^a-zA-Z0-9ก-๙]/g, '');
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();
  
  const attData = attSheet ? attSheet.getDataRange().getDisplayValues() : [];
  
  const sessionsMap = {}; 
  const attendanceMap = {}; 
  
  for (let i = 1; i < attData.length; i++) {
      const row = attData[i];
      if (!row[1]) continue;
      
      const rSub = String(row[4]).replace(/[^a-zA-Z0-9ก-๙]/g, '');
      const rClass = String(row[6]).replace(/[^a-zA-Z0-9ก-๙]/g, '');
      const rTerm = String(row[2]).trim();
      const rYear = String(row[3]).trim();
      
      if (rSub === cleanSub && rClass === cleanClass && rTerm === targetTerm && rYear === targetYear) {
          const stdId = String(parseInt(String(row[8]).trim(), 10)); 
          const status = row[10];
          const period = String(row[7]).trim();
          
          let dateStr = String(row[1]).split(' ')[0].trim(); 
          if(row[12]) {
              const parts = String(row[12]).split('_');
              if(parts.length > 1) dateStr = parts[0];
          }
          
          const sessionKey = dateStr + "_" + period;
          if (!sessionsMap[sessionKey]) sessionsMap[sessionKey] = { date: dateStr, period: period, displayDate: dateStr };
          if (!attendanceMap[stdId]) attendanceMap[stdId] = {};
          attendanceMap[stdId][sessionKey] = { status: status, rowIdx: i + 1 };
      }
  }
  
  return { students: students, sessions: Object.values(sessionsMap).sort((a, b) => a.date.localeCompare(b.date) || parseInt(a.period) - parseInt(b.period)), attendance: attendanceMap };
}

function saveMassiveAttendanceGrid(subjectCode, subjectName, className, term, year, updates, newRecords, teacherId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance_Database");
  if (!sheet) return { status: "error", message: "ไม่พบชีต Attendance_Database" };
  
  if (updates && updates.length > 0) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 0) {
          const statusRange = sheet.getRange(1, 11, lastRow, 1);
          const statusValues = statusRange.getValues();
          updates.forEach(u => { if (u.rowIdx && u.rowIdx <= lastRow) statusValues[u.rowIdx - 1][0] = u.status; });
          statusRange.setValues(statusValues);
      }
  }
  
  if (newRecords && newRecords.length > 0) {
      const timestamp = new Date();
      const dataToAppend = newRecords.map(r => [ timestamp, r.date, term, year, subjectCode, subjectName, className, r.period, r.studentId, r.studentName, r.status, teacherId, `${r.date}_${r.period}` ]);
      sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
  }
  
  return { status: "success", message: "บันทึกข้อมูลตารางรวมเรียบร้อยแล้ว" };
}

// ==========================================
// 6. TIMETABLE SYSTEM (ระบบจัดการตารางสอน)
// ==========================================

function getFilteredTimetables(tid, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const targetTerm = String(term).trim();
  const targetYear = String(year).trim();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowTid  = String(row[5]).trim(); 
    const rowTerm = String(row[8]).trim();
    const rowYear = String(row[9]).trim();

    if ((rowTerm === targetTerm) && (rowYear === targetYear) && (tid === "" || rowTid === tid)) {
      results.push({ rowIndex: i + 1, data: row });
    }
  }
  return results;
}

// ==========================================
// 📅 นำเข้าข้อมูลตารางสอนผ่านไฟล์ CSV (Ultimate Fix: ล็อกคำว่า รหัสวิชา และเติม teacher ก่อนเลข)
// ==========================================
function importTimetableCSV(base64Data, clearOldData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Timetable_Database");
    if (!sheet) return { status: 'error', message: 'ไม่พบชีต Timetable_Database' };

    // 🌟 1. ดึงเทอมและปีการศึกษาปัจจุบันมาใช้อัตโนมัติ
    const config = getSystemConfig();
    const currentTerm = String(config.term).trim();
    const currentYear = String(config.year).trim();

    // 🌟 2. ดึงข้อมูลครูเพื่อทำ Mapping (ชื่อ -> Username/TeacherID)
    const userSheet = ss.getSheetByName("User_Database");
    const teacherMap = {};
    if (userSheet) {
        const userData = userSheet.getDataRange().getDisplayValues();
        for(let i = 1; i < userData.length; i++) {
            const username = String(userData[i][0]).trim();
            const fullName = String(userData[i][2]).trim();
            const role = String(userData[i][3]).trim().toUpperCase();
            
            if (role === 'TEACHER' || role === 'ADMIN') {
                // ตัดช่องว่างทิ้งและทำตัวเล็ก เพื่อให้จับคู่ได้แม่นยำ 100%
                const cleanName = fullName.replace(/\s+/g, '').toLowerCase();
                teacherMap[cleanName] = username;
                const firstName = fullName.split(' ')[0].toLowerCase();
                teacherMap[firstName] = username;
                teacherMap[username.toLowerCase()] = username;
            }
        }
    }

    const decoded = Utilities.base64Decode(base64Data);
    const csvText = Utilities.newBlob(decoded).getDataAsString('UTF-8');
    const csv = Utilities.parseCsv(csvText);

    if (csv.length < 2) return { status: 'error', message: 'ไฟล์ CSV ไม่มีข้อมูล' };

    // 🌟 3. ระบบจับคู่คอลัมน์อัตโนมัติ (Dynamic Column Mapping)
    const headers = csv[0].map(h => String(h).replace(/[\s\u200B-\u200D\uFEFF]/g, '').toLowerCase());
    
    // 🛡️ เปลี่ยนมาใช้ === เพื่อบังคับว่าต้องเป็นคำว่า "รหัสวิชา" เป๊ะๆ เท่านั้น ป้องกันการสับสนกับคำว่ารหัสครู
    let cCode = headers.findIndex(h => h === 'รหัสวิชา' || h === 'subjectcode');
    let cName = headers.findIndex(h => h === 'ชื่อวิชา' || h === 'subjectname');
    let cLevel = headers.findIndex(h => h === 'ระดับชั้น' || h === 'ระดับ' || h === 'ชั้น' || h === 'level');
    let cRoom = headers.findIndex(h => h === 'ห้อง' || h === 'ห้องเรียน' || h === 'room');
    let cLoc = headers.findIndex(h => h === 'สถานที่' || h === 'อาคาร' || h === 'ห้องเรียน' || h === 'location');
    let cTeacher = headers.findIndex(h => h === 'ครูผู้สอน' || h === 'รหัสครู' || h === 'teacherid' || h === 'teacher');
    let cDay = headers.findIndex(h => h === 'วัน' || h === 'วันสอน' || h === 'day');
    let cPeriod = headers.findIndex(h => h === 'คาบ' || h === 'คาบที่' || h === 'period');

    // 🛡️ สำรองตำแหน่งเดิมไว้ ถ้าหาคอลัมน์ใน CSV ไม่เจอ
    if (cCode === -1) cCode = 0;
    if (cName === -1) cName = 1;
    if (cLevel === -1) cLevel = 2;
    if (cRoom === -1) cRoom = 3;
    if (cLoc === -1) cLoc = 4;
    if (cTeacher === -1) cTeacher = 5;
    if (cDay === -1) cDay = 6;
    if (cPeriod === -1) cPeriod = 7;

    // 🌟 4. ระบบล้างข้อมูลอัจฉริยะ (ลบเฉพาะของเทอมปัจจุบัน)
    if (clearOldData && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const rowsToKeep = [data[0]]; 
      for (let i = 1; i < data.length; i++) {
         const rowTerm = String(data[i][8]).trim();
         const rowYear = String(data[i][9]).trim();
         if (rowTerm !== currentTerm || rowYear !== currentYear) rowsToKeep.push(data[i]);
      }
      sheet.clearContents();
      if (rowsToKeep.length > 0) sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    }

    let newRows = [];
    
    // 🌟 5. วนลูปอ่านและจัดเรียงข้อมูล
    for (let i = 1; i < csv.length; i++) {
      let subjectCode = String(csv[i][cCode] || "").trim();
      let subjectName = String(csv[i][cName] || "").trim();
      
      // 🌟 ท่าไม้ตาย: ถ้าใส่รหัสวิชาเป็น "-" ให้เปลี่ยนเป็น "กิจกรรม" ทันที
      if (subjectCode === '-') {
          subjectCode = 'กิจกรรม';
      }
      
      // ถ้าช่องรหัสวิชาว่างเปล่าจริงๆ ค่อยข้ามบรรทัดนี้ไป
      if (subjectCode === '') continue; 
      
      let level = String(csv[i][cLevel] || "").trim();
      let room = String(csv[i][cRoom] || "1").trim();
      let location = String(csv[i][cLoc] || "-").trim();
      
      // ดึงชื่อครู หรือ รหัสครู มาจาก CSV
      let rawTeacher = String(csv[i][cTeacher] || "").trim();
      let searchTeacher = rawTeacher.replace(/\s+/g, '').toLowerCase();
      
      // แปลงชื่อเป็นรหัส (ถ้าค้นเจอ)
      let teacherId = teacherMap[searchTeacher] ? teacherMap[searchTeacher] : rawTeacher;

      // 🌟 ท่าไม้ตาย: บังคับเพิ่มคำว่า "teacher" ก่อนตัวเลข
      let numPart = teacherId.replace(/\D/g, ''); // ดึงเฉพาะตัวเลขออกมา (เช่น 12345)
      if (numPart !== "") {
          teacherId = "teacher" + numPart; // นำมาประกอบร่างเป็น teacher12345
      }

      // จัดระเบียบวันในสัปดาห์
      let rawDay = String(csv[i][cDay] || "").trim();
      let day = rawDay;
      if (rawDay.includes('จ') || rawDay.toLowerCase() === 'monday') day = 'จันทร์';
      else if (rawDay.includes('อ') || rawDay.toLowerCase() === 'tuesday') day = 'อังคาร';
      else if (rawDay.includes('พฤ') || rawDay.toLowerCase() === 'thursday') day = 'พฤหัสบดี';
      else if (rawDay.includes('พ') || rawDay.toLowerCase() === 'wednesday') day = 'พุธ';
      else if (rawDay.includes('ศ') || rawDay.toLowerCase() === 'friday') day = 'ศุกร์';

      let period = String(csv[i][cPeriod] || "1").replace(/\D/g, ''); 

      // ✅ บังคับโครงสร้าง 10 คอลัมน์เป๊ะๆ เทลงชีต
      newRows.push([
          subjectCode, subjectName, level, room, location, 
          teacherId, day, period, currentTerm, currentYear
      ]);
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
    }
    
    SpreadsheetApp.flush();
    return { 
        status: 'success', 
        message: `✅ นำเข้าและจัดเรียงตารางสอนสำเร็จ ${newRows.length} คาบ\n(แปลงรหัสครูสำเร็จ และจับคู่ภาคเรียน ${currentTerm}/${currentYear} อัตโนมัติ)` 
    };

  } catch (e) {
    return { status: 'error', message: '❌ ข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateTimetableRow(idx, data) {
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
    s.getRange(1, 1, 1, sh.headers.length).setValues([sh.headers]).setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
    if(sh.name === "System_Settings" && s.getLastRow() === 1) { s.appendRow(["Current_Term", "1"]); s.appendRow(["Current_Year", "2568"]); }
  });
  
  const uSheet = ss.getSheetByName("User_Database");
  if (uSheet.getLastRow() === 1) uSheet.appendRow(["admin", "1234", "ครูน๊อต ศิกษก", "Admin", "บริหาร", "not@school.ac.th", "2568"]);
  return "✅ ฐานข้อมูลพร้อมใช้งาน!";
}

function migrateTimetableStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Timetable_Database");
  if (!sheet) return;

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return;

  const newRows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    let level = String(row[2]).trim(), location = "-", room = "1";
    const parts = level.split(/\s+/); 
    if (parts.length >= 2) { level = parts[0]; location = parts[1]; } 
    else { level = parts[0]; location = "-"; }

    newRows.push([ row[0], row[1], level, room, location, row[3], row[4], row[5], row[6], row[7] ]);
  }

  sheet.clearContents(); 
  sheet.appendRow(["SubjectCode", "SubjectName", "Level", "Room", "Location", "TeacherID", "Day", "Period", "Term", "Year"]);
  if (newRows.length > 0) sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);
  sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#fff2cc");
  sheet.setFrozenRows(1);
}

// ==========================================
// 🚀 ฟังก์ชันสร้างฐานข้อมูลสำหรับระบบ ปพ.5 แบบอัตโนมัติ (อัปเกรด Grade_Summary)
// ==========================================
function setupPorPor5Database() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. สร้าง Sheet: Subject_Config
  let sheetConfig = ss.getSheetByName("Subject_Config");
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet("Subject_Config");
    sheetConfig.appendRow(["subject_id", "subject_code", "class_name", "term", "year", "score_ratio", "indicators_json", "teacher_id"]);
    sheetConfig.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
    sheetConfig.setFrozenRows(1);
  }

  // 2. สร้าง Sheet: Score_Database
  let sheetScore = ss.getSheetByName("Score_Database");
  if (!sheetScore) {
    sheetScore = ss.insertSheet("Score_Database");
    sheetScore.appendRow(["uid", "student_id", "subject_code", "indicator_id", "score", "term", "year"]);
    sheetScore.getRange("A1:G1").setFontWeight("bold").setBackground("#fff2cc");
    sheetScore.setFrozenRows(1);
  }

  // 3. สร้าง Sheet: Qualitative_Assess
  let sheetQual = ss.getSheetByName("Qualitative_Assess");
  if (!sheetQual) {
    sheetQual = ss.insertSheet("Qualitative_Assess");
    sheetQual.appendRow(["student_id", "subject_code", "term", "year", "reading_writing", "char_json", "comp_json"]);
    sheetQual.getRange("A1:G1").setFontWeight("bold").setBackground("#c9daf8");
    sheetQual.setFrozenRows(1);
  }

  // 4. สร้าง Sheet: Grade_Summary (อัปเกรด 8 คอลัมน์ รองรับประวัติย้อนหลัง)
  let sheetGrade = ss.getSheetByName("Grade_Summary");
  if (!sheetGrade) {
    sheetGrade = ss.insertSheet("Grade_Summary");
    sheetGrade.appendRow(["student_id", "subject_code", "total_score", "grade", "remedial_status", "attendance_percent", "term", "year"]);
    sheetGrade.getRange("A1:H1").setFontWeight("bold").setBackground("#f4cccc");
    sheetGrade.setFrozenRows(1);
  } else {
    // 🌟 แอบอัปเกรดให้ถ้าชีตเก่ามีแค่ 6 คอลัมน์
    if (sheetGrade.getMaxColumns() < 8) {
      sheetGrade.insertColumnsAfter(sheetGrade.getMaxColumns(), 8 - sheetGrade.getMaxColumns());
      sheetGrade.getRange(1, 7, 1, 2).setValues([["term", "year"]]).setFontWeight("bold").setBackground("#f4cccc");
    }
  }

  return "✅ สร้างฐานข้อมูล ปพ.5 ทั้ง 4 แผ่นเรียบร้อยแล้วครับ!";
}

// ==========================================
// ระบบตั้งค่าการพิมพ์ ปพ.5 (ผู้ลงนาม & ครูที่ปรึกษา)
// ==========================================

function getPrintConfigData(term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Print_Config');
  if (!sheet) { sheet = ss.insertSheet('Print_Config'); sheet.appendRow(['term', 'year', 'sys_data_json', 'homeroom_data_json']); }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
     if (String(data[i][0]) === String(term) && String(data[i][1]) === String(year)) {
         return { status: 'success', sys: JSON.parse(data[i][2] || '{}'), hr: JSON.parse(data[i][3] || '[]') };
     }
  }
  return { status: 'success', sys: { school_name: 'โรงเรียนภูพระบาทวิทยา', principal_name: '', measure_head: '', academic_head: '' }, hr: [] };
}

function savePrintConfigData(payload) {
  const lock = LockService.getScriptLock();
  try {
     lock.waitLock(10000);
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     let sheet = ss.getSheetByName('Print_Config');
     if (!sheet) { sheet = ss.insertSheet('Print_Config'); sheet.appendRow(['term', 'year', 'sys_data_json', 'homeroom_data_json']); }
     const data = sheet.getDataRange().getValues();
     let found = false;
     for (let i = 1; i < data.length; i++) {
         if (String(data[i][0]) === String(payload.term) && String(data[i][1]) === String(payload.year)) {
             sheet.getRange(i + 1, 3).setValue(JSON.stringify(payload.sys));
             sheet.getRange(i + 1, 4).setValue(JSON.stringify(payload.hr));
             found = true; break;
         }
     }
     if (!found) sheet.appendRow([payload.term, payload.year, JSON.stringify(payload.sys), JSON.stringify(payload.hr)]);
     return { status: 'success', message: `✅ บันทึกตั้งค่า ปพ.5 ของภาคเรียนที่ ${payload.term}/${payload.year} เรียบร้อย!` };
  } catch(e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function getTeacherListForDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_Database");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const teachers = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toUpperCase() === 'TEACHER') teachers.push(String(data[i][2]).trim()); 
  }
  return teachers.sort((a, b) => a.localeCompare(b, 'th')); 
}

function generatePP5Template(payload) {
  const template = HtmlService.createTemplateFromFile('Template_PP5');
  template.data = payload;
  return template.evaluate().getContent();
}

// ==========================================
// 12. LESSON RECORD & FILE UPLOAD (บันทึกหลังสอนแบบละเอียด)
// ==========================================

function getOrCreateUploadFolder() {
  const folderName = "PSSMS_Uploads";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function uploadFileToDrive(base64Data, filename) {
  if (!base64Data || base64Data === "" || base64Data === "null") return ""; 
  try {
    const folder = getOrCreateUploadFolder();
    const splitBase = base64Data.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteCharacters, type, filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl(); 
  } catch (e) {
    console.error("Upload Error: " + e.message); return "Error: " + e.message; 
  }
}

function saveDetailedLessonRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) {
    sheet = ss.insertSheet("Detailed_Lesson_Records");
    sheet.appendRow(["Timestamp", "Date", "Term", "Year", "SubjectCode", "SubjectName", "Class", "Period", "Topic", "Outcomes", "Problems", "Solutions", "DPA_Indicators", "Skills_3R8C", "Student_Results", "WorkFileURL", "AtmosphereImageURL", "TeacherID", "SessionID"]);
    sheet.getRange("A1:S1").setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
  }

  let workUrl = "", imageUrl = "";
  const timeStampStr = new Date().getTime();
  if (record.workFileBase64) workUrl = uploadFileToDrive(record.workFileBase64, `Work_${record.subjectCode}_${timeStampStr}`);
  if (record.imageFileBase64) imageUrl = uploadFileToDrive(record.imageFileBase64, `Atmosphere_${record.subjectCode}_${timeStampStr}`);

  const config = getSystemConfig();
  const sessionID = `${record.date}|${record.subjectCode}|${record.className}|${record.period}`;

  sheet.appendRow([ new Date(), record.date, config.term, config.year, record.subjectCode, record.subjectName, record.className, record.period, record.topic, record.outcomes, record.problems, record.solutions, JSON.stringify(record.dpa), JSON.stringify(record.skills), record.studentResults, workUrl, imageUrl, record.teacherId, sessionID ]);
  return { status: "success", message: "✅ บันทึกข้อมูลการสอนแบบละเอียดเรียบร้อยแล้ว!" };
}

function getDetailedLessonRecords(teacherId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  const results = [];
  
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    if (!row[1]) continue;
    if (String(row[17]).trim() === String(teacherId).trim() && String(row[2]).trim() === String(term).trim() && String(row[3]).trim() === String(year).trim()) {
      let dpaArray = [], skillsArray = [];
      try { dpaArray = JSON.parse(row[12]); } catch(e) {}
      try { skillsArray = JSON.parse(row[13]); } catch(e) {}

      results.push({ timestamp: row[0], date: row[1], subjectCode: row[4], subjectName: row[5], className: row[6], period: row[7], topic: row[8], outcomes: row[9], problems: row[10], solutions: row[11], dpa: dpaArray, skills: skillsArray, studentResults: row[14], workUrl: row[15], imageUrl: row[16] });
    }
  }
  return results;
}

function deleteDetailedLessonRecord(timestampStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(timestampStr).trim()) {
      sheet.deleteRow(i + 1); return { status: "success", message: "🗑️ ลบข้อมูลบันทึกเรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลที่ต้องการลบ" };
}

function updateDetailedLessonRecord(timestampStr, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Detailed_Lesson_Records");
  if (!sheet) return { status: "error", message: "ไม่พบฐานข้อมูล" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(timestampStr).trim()) {
      const row = i + 1;
      sheet.getRange(row, 9).setValue(record.topic);
      sheet.getRange(row, 10).setValue(record.outcomes);
      sheet.getRange(row, 11).setValue(record.problems);
      sheet.getRange(row, 12).setValue(record.solutions);
      sheet.getRange(row, 13).setValue(JSON.stringify(record.dpa));
      sheet.getRange(row, 14).setValue(JSON.stringify(record.skills));
      sheet.getRange(row, 15).setValue(record.studentResults);
      if (record.workFileBase64) sheet.getRange(row, 16).setValue(uploadFileToDrive(record.workFileBase64, `Work_Updated_${new Date().getTime()}`));
      if (record.imageFileBase64) sheet.getRange(row, 17).setValue(uploadFileToDrive(record.imageFileBase64, `Atmosphere_Updated_${new Date().getTime()}`));
      return { status: "success", message: "✅ อัปเดตข้อมูลการสอนเรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลที่ต้องการแก้ไข" };
}

// ==========================================
// 📚 ระบบ ปพ.5: โครงสร้างรายวิชา (Subject Config)
// ==========================================

function getSubjectConfig(subjectCode, className, term, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subject_Config");
  if(!sheet) return null;
  const data = sheet.getDataRange().getValues(); 
  const targetSubj = String(subjectCode).trim(); const targetClass = String(className).trim();
  const targetTerm = String(term).trim(); const targetYear = String(year).trim();
  let exactMatch = null; let historyMatch = null;
  
  for(let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === targetSubj) {
      let parsedIndicators = [];
      try { parsedIndicators = typeof data[i][6] === 'string' ? JSON.parse(data[i][6] || '[]') : data[i][6]; } catch(e) { parsedIndicators = []; }

      let examInds = null;
      try { examInds = typeof data[i][8] === 'string' ? JSON.parse(data[i][8] || 'null') : data[i][8]; } catch(e) {}

      let rawRatio = data[i][5]; let safeRatio = "70:10:20";
      if (rawRatio instanceof Date) { safeRatio = `${rawRatio.getHours()}:${rawRatio.getMinutes()}:${rawRatio.getSeconds()}`; if (safeRatio === "22:10:20") safeRatio = "70:10:20"; } else if (String(rawRatio).includes(':')) { safeRatio = String(rawRatio).replace(/'/g, '').trim(); }

      if (!historyMatch) historyMatch = { ratio: safeRatio, indicators: parsedIndicators, examIndicators: examInds };
      if (String(data[i][2]).trim() === targetClass && String(data[i][3]).trim() === targetTerm && String(data[i][4]).trim() === targetYear) {
        exactMatch = { ratio: safeRatio, indicators: parsedIndicators, examIndicators: examInds }; break; 
      }
    }
  }
  return exactMatch || historyMatch || null; 
}

function saveSubjectConfig(configData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subject_Config");
  if(!sheet) return {status: 'error', message: 'ไม่พบ Database: Subject_Config'};
  const data = sheet.getDataRange().getValues();
  const targetSubj = String(configData.subjectCode).trim(); const targetClass = String(configData.className).trim();
  const targetTerm = String(configData.term).trim(); const targetYear = String(configData.year).trim();
  const subjectId = `${targetSubj}_${targetClass}_${targetTerm}_${targetYear}`;
  const ratioStr = `'${configData.formative}:${configData.midterm}:${configData.final}`;
  
  const rowData = [ subjectId, targetSubj, targetClass, targetTerm, targetYear, ratioStr, JSON.stringify(configData.indicators), configData.teacherId, JSON.stringify(configData.examIndicators || null) ];

  for(let i = 1; i < data.length; i++) {
    if(String(data[i][1]).trim() === targetSubj && String(data[i][2]).trim() === targetClass && String(data[i][3]).trim() === targetTerm && String(data[i][4]).trim() === targetYear) {
      sheet.getRange(i + 1, 1, 1, 9).setValues([rowData]); return {status: 'success', message: 'อัปเดตโครงสร้างวิชาเรียบร้อยแล้ว!'};
    }
  }
  sheet.appendRow(rowData); return {status: 'success', message: 'บันทึกโครงสร้างวิชาใหม่เรียบร้อยแล้ว!'};
}

// ==========================================
// 13. ระบบ ปพ.5: All-in-One Score & Evaluation
// ==========================================

// 🌟 อัปเกรด: กรองเกรดให้ตรงเทอมและปี + กู้คืนรายชื่อในอดีต (Historical Roster)
function getAllInOneScoreGridData(subjectCode, className, term, year) {
  let config = getSubjectConfig(subjectCode, className, term, year);
  if (!config) config = { ratio: "70:10:20", indicators: [{name: "คะแนนเก็บ 1", score: 70}] }; 

  let students = getStudentsByClass(className, year);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
  const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

  // เพิ่มนักเรียนจาก Grade_Summary ที่อาจไม่มีใน User_Database/History แล้ว
  const configSys = getSystemConfig();
  if (String(year).trim() !== String(configSys.year).trim()) {
      const gradeSheet = ss.getSheetByName("Grade_Summary");
      const userSheet = ss.getSheetByName("User_Database");
      if (gradeSheet && userSheet) {
          const gradeData = gradeSheet.getDataRange().getDisplayValues();
          const userData = userSheet.getDataRange().getDisplayValues();
          
          const userMap = {};
          for(let i=1; i<userData.length; i++) userMap[normID(userData[i][0])] = userData[i];
          
          const existingIds = new Set(students.map(s => normID(s[0])));
          
          for(let i=1; i<gradeData.length; i++) {
              const rowSub = normStr(gradeData[i][1]);
              const rowTerm = normStr(gradeData[i][6]);
              const rowYear = normStr(gradeData[i][7]);
              
              // ควานหาเด็กที่เคยมีเกรดวิชานี้ ในเทอม/ปีในอดีต
              if((rowSub === normStr(subjectCode) || rowSub.includes(normStr(subjectCode))) &&
                 rowTerm === normStr(term) && rowYear === normStr(year)) {
                  
                  const stdId = normID(gradeData[i][0]);
                  // ถ้าเจอแต่ไม่มีชื่อในชั้นเรียนปัจจุบัน ให้ดึงโปรไฟล์กลับมาโชว์!
                  if (!existingIds.has(stdId) && userMap[stdId]) {
                      students.push(userMap[stdId]);
                      existingIds.add(stdId);
                  }
              }
          }
      }
  }

  const existingScores = {};
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

  const gradeSheet = ss.getSheetByName("Grade_Summary");
  if (gradeSheet && gradeSheet.getLastRow() > 0) {
    const gradeData = gradeSheet.getDataRange().getDisplayValues();
    for(let i = 1; i < gradeData.length; i++) {
        const rowSub = normStr(gradeData[i][1]);
        const rowTerm = normStr(gradeData[i][6]);
        const rowYear = normStr(gradeData[i][7]);
        
        if((rowSub === normStr(subjectCode) || rowSub.includes(normStr(subjectCode))) &&
           (rowTerm === normStr(term) || rowTerm === '') && 
           (rowYear === normStr(year) || rowYear === '')) {
            
            const stdKey = normID(gradeData[i][0]);
            let foundRemark = false;
            for (let col = 2; col < gradeData[i].length; col++) {
                const cellVal = String(gradeData[i][col]).trim();
                if (cellVal === 'ร' || cellVal === 'มส') { existingScores[`${stdKey}_remark`] = cellVal; foundRemark = true; break; }
            }
            if (!foundRemark && !existingScores[`${stdKey}_remark`]) existingScores[`${stdKey}_remark`] = '-';
        }
    }
  }

  const qualSheet = ss.getSheetByName("Qualitative_Assess");
  const qualData = qualSheet ? qualSheet.getDataRange().getDisplayValues() : [];
  const existingQuals = {};
  for(let i = 1; i < qualData.length; i++) {
    const row = qualData[i];
    if(normStr(row[1]) === normStr(subjectCode) && normStr(row[2]) === normStr(term) && normStr(row[3]) === normStr(year)) {
      if (row.length >= 16) {
          existingQuals[normID(row[0])] = { 
              read1: row[4], read2: row[5], read3: row[6], read4: row[7], readTotal: row[8], read: row[9],
              char1: row[10], char2: row[11], char3: row[12], char4: row[13], charTotal: row[14], char: row[15],
              comp: row[16] || '3'
          };
      } else { existingQuals[normID(row[0])] = { read: row[4], char: row[5], comp: row[6] }; }
    }
  }

  let attStats = {}; let attDetails = {}; let attSessions = []; 
  try {
    const report = getSemesterReport(subjectCode, className, term, year);
    if(report && report.students) {
      attSessions = report.meta.sessionsList || []; 
      report.students.forEach(s => { attStats[normID(s.id)] = parseFloat(s.percent); attDetails[normID(s.id)] = s; });
    }
  } catch(e) {}

  return { config: config, students: students, existingScores: existingScores, existingQuals: existingQuals, attStats: attStats, attDetails: attDetails, attSessions: attSessions };
}

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

    const configSheet = ss.getSheetByName("Subject_Config");
    const sheetScore = ss.getSheetByName("Score_Database");
    const qualSheet = ss.getSheetByName("Qualitative_Assess");
    const gradeSheet = ss.getSheetByName("Grade_Summary");

    if (!configSheet || !sheetScore || !qualSheet || !gradeSheet) return { status: 'error', message: '❌ ไม่พบชีตฐานข้อมูล ปพ.5' };

    const msMap = {};
    if (scoreRecords) {
       scoreRecords.forEach(r => {
          if (String(r.indicatorId).trim().toLowerCase() === 'remark' && r.score) {
             const v = String(r.score).trim(); if(v === 'ร' || v === 'มส') msMap[normID(r.studentId)] = v;
          }
       });
    }
    if (gradeRecords) {
       gradeRecords.forEach(r => {
          let g = String(r.grade || '').trim(); let rm = String(r.remark || '').trim();
          if(g === 'ร' || g === 'มส') msMap[normID(r.studentId)] = g;
          if(rm === 'ร' || rm === 'มส') msMap[normID(r.studentId)] = rm;
       });
    }

    Object.keys(msMap).forEach(sid => {
       let found = false;
       scoreRecords.forEach(r => { if (normID(r.studentId) === sid && String(r.indicatorId).trim().toLowerCase() === 'remark') { r.score = msMap[sid]; found = true; } });
       if (!found) scoreRecords.push({ studentId: sid, subjectCode: subjectCode, term: term, year: year, indicatorId: 'remark', score: msMap[sid] });
    });
    
    if (newConfig) {
         if (configSheet.getLastRow() === 0) configSheet.appendRow(["subject_id", "subject_code", "class_name", "term", "year", "score_ratio", "indicators_json", "teacher_id"]);
         const configData = configSheet.getDataRange().getValues();
         let configUpdated = false;
         const ratioStr = `'${newConfig.formative || 70}:${newConfig.midterm || 10}:${newConfig.final || 20}`;
         const subjectId = `${subjectCode}_${className}_${term}_${year}`;
         const indicatorsJson = JSON.stringify(newConfig.indicators || []);

         for (let i = 1; i < configData.length; i++) {
           if (String(configData[i][1]).trim() === String(subjectCode).trim() && String(configData[i][2]).trim() === String(className).trim() && String(configData[i][3]).trim() === String(term).trim() && String(configData[i][4]).trim() === String(year).trim()) {
               configSheet.getRange(i + 1, 6).setValue(ratioStr); configSheet.getRange(i + 1, 7).setValue(indicatorsJson); configUpdated = true; break;
           }
         }
         if (!configUpdated) configSheet.appendRow([subjectId, subjectCode, className, term, year, ratioStr, indicatorsJson, teacherId]);
    }

    if (scoreRecords && scoreRecords.length > 0) {
        if (sheetScore.getLastRow() === 0) sheetScore.appendRow(["uid", "student_id", "subject_code", "indicator_id", "score", "term", "year"]);
        let scoreData = sheetScore.getDataRange().getValues(); 
        const scoreMap = {};
        scoreRecords.forEach(r => { if(r.score !== undefined && r.score !== null) scoreMap[`${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.indicatorId)}_${normStr(r.term)}_${normStr(r.year)}`] = r; });

        let scoreUpdated = false;
        for(let i = 1; i < scoreData.length; i++) {
          const uid = `${normID(scoreData[i][1])}_${normStr(scoreData[i][2])}_${normStr(scoreData[i][3])}_${normStr(scoreData[i][5])}_${normStr(scoreData[i][6])}`;
          if(scoreMap[uid]) {
             if (String(scoreData[i][4]) !== String(scoreMap[uid].score)) { scoreData[i][4] = scoreMap[uid].score; scoreUpdated = true; }
             scoreMap[uid].processed = true; 
          }
        }
        if(scoreUpdated) {
            const uniformScore = scoreData.map(row => { let r = row.slice(0, 7); while(r.length < 7) r.push(""); return r; });
            sheetScore.getRange(1, 1, uniformScore.length, 7).setValues(uniformScore);
        }

        const newScores = [];
        for (let uid in scoreMap) {
           if (!scoreMap[uid].processed && scoreMap[uid].score !== '') newScores.push([uid, "'" + scoreMap[uid].studentId, scoreMap[uid].subjectCode, scoreMap[uid].indicatorId, scoreMap[uid].score, scoreMap[uid].term, scoreMap[uid].year]);
        }
        if (newScores.length > 0) sheetScore.getRange(sheetScore.getLastRow() + 1, 1, newScores.length, 7).setValues(newScores);
    }

    if (qualRecords && qualRecords.length > 0) {
        if (qualSheet.getLastRow() === 0) qualSheet.appendRow(["student_id", "subject_code", "term", "year", "read1", "read2", "read3", "read4", "readTotal", "read_grade", "char1", "char2", "char3", "char4", "charTotal", "char_grade", "comp"]);
        if (qualSheet.getMaxColumns() < 17) qualSheet.insertColumnsAfter(qualSheet.getMaxColumns(), 17 - qualSheet.getMaxColumns());

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
             qualData[i][14] = q.charTotal; qualData[i][15] = q.char; qualData[i][16] = q.comp || '3';
             qualUpdated = true; qualMap[uid].processed = true;
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
              newQuals.push(["'" + r.studentId, r.subjectCode, r.term, r.year, r.read1, r.read2, r.read3, r.read4, r.readTotal, r.read, r.char1, r.char2, r.char3, r.char4, r.charTotal, r.char, r.comp || '3']);
           }
        }
        if(newQuals.length > 0) qualSheet.getRange(qualSheet.getLastRow() + 1, 1, newQuals.length, 17).setValues(newQuals);
    }

    // 🌟 4. Grade Summary (อัปเกรด Term/Year อัตโนมัติ)
    if (gradeRecords && gradeRecords.length > 0) { 
        if (gradeSheet.getLastRow() === 0) gradeSheet.appendRow(["student_id", "subject_code", "total_score", "grade", "remedial_status", "attendance_percent", "term", "year"]);
        if (gradeSheet.getMaxColumns() < 8) gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), 8 - gradeSheet.getMaxColumns());
        
        let gradeData = gradeSheet.getDataRange().getValues();
        const gradeMap = {};
        
        gradeRecords.forEach(r => {
            const uid = `${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(term)}_${normStr(year)}`;
            let cleanRemark = String(r.remark || '').trim(); if (cleanRemark === '') cleanRemark = '-';
            r.remark = cleanRemark; gradeMap[uid] = r;
        });

        let gradeUpdated = false;
        for(let i = 1; i < gradeData.length; i++) {
            const row = gradeData[i];
            const rTerm = normStr(row[6]);
            const rYear = normStr(row[7]);
            
            // ผูก ID ด้วยเทอม/ปี (ถ้าของเก่าไม่มีให้ถือว่าเป็นของเทอมนี้ไปเลย เพื่ออัปเกรด)
            const uid = `${normID(row[0])}_${normStr(row[1])}_${rTerm === '' ? normStr(term) : rTerm}_${rYear === '' ? normStr(year) : rYear}`;
            
            if(gradeMap[uid]) {
                if(String(gradeData[i][2]) !== String(gradeMap[uid].totalScore) || 
                   String(gradeData[i][3]) !== String(gradeMap[uid].grade) || 
                   String(gradeData[i][4]) !== String(gradeMap[uid].remark) ||
                   String(gradeData[i][6]) !== String(term) ||
                   String(gradeData[i][7]) !== String(year)) {
                    
                    gradeData[i][2] = gradeMap[uid].totalScore;
                    gradeData[i][3] = gradeMap[uid].grade;
                    gradeData[i][4] = gradeMap[uid].remark; 
                    gradeData[i][6] = term; // อัปเกรดเทอม
                    gradeData[i][7] = year; // อัปเกรดปี
                    gradeUpdated = true;
                }
                gradeMap[uid].processed = true;
            }
        }
        
        if(gradeUpdated) {
            const uniformGrade = gradeData.map(row => { let r = row.slice(0, 8); while(r.length < 8) r.push(""); return r; });
            gradeSheet.getRange(1, 1, uniformGrade.length, 8).setValues(uniformGrade);
        }
        
        const newGrades = [];
        for (let uid in gradeMap) {
            if(!gradeMap[uid].processed) {
                const r = gradeMap[uid];
                newGrades.push(["'" + r.studentId, r.subjectCode, r.totalScore, r.grade, r.remark, "100", term, year]);
            }
        }
        if(newGrades.length > 0) gradeSheet.getRange(gradeSheet.getLastRow() + 1, 1, newGrades.length, 8).setValues(newGrades);
    } 

    SpreadsheetApp.flush(); 
    return {status: 'success', message: `✅ บันทึกเสร็จสมบูรณ์!`};
  } catch(e) { return { status: 'error', message: e.message + " | บรรทัด: " + (e.lineNumber||'') }; } finally { lock.releaseLock(); }
}

function saveAllInOneScores(payload) {
  const { subjectCode, className, teacherId, term, year, scoreRecords, qualRecords, gradeRecords } = payload;
  if (!verifyTeacherPermission(teacherId, subjectCode, className, term, year)) return { status: 'error', message: '❌ ความปลอดภัย: คุณไม่มีสิทธิ์บันทึกคะแนน!' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
  const normStr = (str) => String(str).replace(/\s+/g, '').toLowerCase();

  const sheetScore = ss.getSheetByName("Score_Database");
  let scoreData = sheetScore.getDataRange().getValues(); 
  const scoreMap = {};
  scoreRecords.forEach(r => { scoreMap[`${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.indicatorId)}_${normStr(r.term)}_${normStr(r.year)}`] = r; });

  let scoreUpdated = false;
  for(let i = 1; i < scoreData.length; i++) {
    const row = scoreData[i];
    const uid = `${normID(row[1])}_${normStr(row[2])}_${normStr(row[3])}_${normStr(row[5])}_${normStr(row[6])}`;
    if(scoreMap[uid]) {
       if (String(scoreData[i][4]) !== String(scoreMap[uid].score)) { logScoreHistory(teacherId, scoreMap[uid].studentId, subjectCode, scoreMap[uid].indicatorId, scoreData[i][4], scoreMap[uid].score, term, year); scoreData[i][4] = scoreMap[uid].score; scoreUpdated = true; }
       scoreMap[uid].processed = true; 
    }
  }
  if(scoreUpdated) sheetScore.getRange(1, 1, scoreData.length, scoreData[0].length).setValues(scoreData);

  const newScores = [];
  for (let uid in scoreMap) { if (!scoreMap[uid].processed) newScores.push([uid, "'" + scoreMap[uid].studentId, scoreMap[uid].subjectCode, scoreMap[uid].indicatorId, scoreMap[uid].score, scoreMap[uid].term, scoreMap[uid].year]); }
  if (newScores.length > 0) sheetScore.getRange(sheetScore.getLastRow() + 1, 1, newScores.length, newScores[0].length).setValues(newScores);

  const qualSheet = ss.getSheetByName("Qualitative_Assess");
  if (qualSheet.getMaxColumns() < 17) qualSheet.insertColumnsAfter(qualSheet.getMaxColumns(), 17 - qualSheet.getMaxColumns());
  let qualData = qualSheet.getDataRange().getValues();
  const qualMap = {};
  qualRecords.forEach(r => { qualMap[`${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(r.term)}_${normStr(r.year)}`] = r; });

  let qualUpdated = false;
  for(let i = 1; i < qualData.length; i++) {
    const row = qualData[i];
    const uid = `${normID(row[0])}_${normStr(row[1])}_${normStr(row[2])}_${normStr(row[3])}`;
    if(qualMap[uid]) {
       const q = qualMap[uid];
       qualData[i][4] = q.read1; qualData[i][5] = q.read2; qualData[i][6] = q.read3; qualData[i][7] = q.read4; qualData[i][8] = q.readTotal; qualData[i][9] = q.read; 
       qualData[i][10] = q.char1; qualData[i][11] = q.char2; qualData[i][12] = q.char3; qualData[i][13] = q.char4; qualData[i][14] = q.charTotal; qualData[i][15] = q.char; qualData[i][16] = q.comp || '3';
       qualUpdated = true; qualMap[uid].processed = true;
    }
  }
  if(qualUpdated) { const uniformQual = qualData.map(row => { let r = row.slice(0, 17); while(r.length < 17) r.push(""); return r; }); qualSheet.getRange(1, 1, uniformQual.length, 17).setValues(uniformQual); }
  
  const newQuals = [];
  for (let uid in qualMap) { if(!qualMap[uid].processed) { const r = qualMap[uid]; newQuals.push(["'" + r.studentId, r.subjectCode, r.term, r.year, r.read1, r.read2, r.read3, r.read4, r.readTotal, r.read, r.char1, r.char2, r.char3, r.char4, r.charTotal, r.char, r.comp || '3']); } }
  if(newQuals.length > 0) qualSheet.getRange(qualSheet.getLastRow() + 1, 1, newQuals.length, 17).setValues(newQuals);

  // 🌟 3. บันทึกเกรด (Grade_Summary) - อัปเกรด Term/Year
  const gradeSheet = ss.getSheetByName("Grade_Summary");
  if (gradeSheet.getMaxColumns() < 8) gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), 8 - gradeSheet.getMaxColumns());
  let gradeData = gradeSheet.getDataRange().getValues();
  const gradeMap = {};
  gradeRecords.forEach(r => { gradeMap[`${normID(r.studentId)}_${normStr(r.subjectCode)}_${normStr(term)}_${normStr(year)}`] = r; });

  let gradeUpdated = false;
  for(let i = 1; i < gradeData.length; i++) {
    const row = gradeData[i];
    const rTerm = normStr(row[6]); const rYear = normStr(row[7]);
    const uid = `${normID(row[0])}_${normStr(row[1])}_${rTerm === '' ? normStr(term) : rTerm}_${rYear === '' ? normStr(year) : rYear}`;
    
    if(gradeMap[uid]) {
       if(String(gradeData[i][2]) !== String(gradeMap[uid].totalScore) || String(gradeData[i][3]) !== String(gradeMap[uid].grade) || String(gradeData[i][4]) !== String(gradeMap[uid].remark) || String(gradeData[i][6]) !== String(term) || String(gradeData[i][7]) !== String(year)) {
           gradeData[i][2] = gradeMap[uid].totalScore; gradeData[i][3] = gradeMap[uid].grade; gradeData[i][4] = gradeMap[uid].remark || "-"; gradeData[i][6] = term; gradeData[i][7] = year; gradeUpdated = true;
       }
       gradeMap[uid].processed = true;
    }
  }
  if(gradeUpdated) { const uniformGrade = gradeData.map(row => { let r = row.slice(0, 8); while(r.length < 8) r.push(""); return r; }); gradeSheet.getRange(1, 1, uniformGrade.length, 8).setValues(uniformGrade); }
  
  const newGrades = [];
  for (let uid in gradeMap) { if(!gradeMap[uid].processed) { const r = gradeMap[uid]; newGrades.push(["'" + r.studentId, r.subjectCode, r.totalScore, r.grade, r.remark || "-", "100", term, year]); } }
  if(newGrades.length > 0) gradeSheet.getRange(gradeSheet.getLastRow() + 1, 1, newGrades.length, 8).setValues(newGrades);

  SpreadsheetApp.flush(); 
  return {status: 'success', message: 'บันทึกคะแนน เกรด และคุณลักษณะเรียบร้อยแล้ว!'};
}

function logScoreHistory(teacherId, stdId, subCode, indId, oldScore, newScore, term, year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Score_History");
    if (!sheet) { sheet = ss.insertSheet("Score_History"); sheet.appendRow(["Timestamp", "TeacherID", "StudentID", "SubjectCode", "IndicatorID", "OldScore", "NewScore", "Term", "Year"]); }
    if (String(oldScore).trim() !== String(newScore).trim()) sheet.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"), teacherId, stdId, subCode, indId, oldScore, newScore, term, year]);
  } catch(e) {} 
}

function getScoreHistory(stdId, subCode, indId, term, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Score_History");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  const history = [];
  for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      if (String(row[2]).trim() === String(stdId).trim() && String(row[3]).trim() === String(subCode).trim() && String(row[4]).trim() === String(indId).trim() && String(row[7]).trim() === String(term).trim() && String(row[8]).trim() === String(year).trim()) {
          history.push({ time: row[0], old: row[5] === "" ? "-" : row[5], new: row[6] === "" ? "-" : row[6] });
          if(history.length >= 10) break;
      }
  }
  return history;
}

function getMorningActivityData(dateStr, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Morning_Activity");
  if (!sheet) return {};
  const data = sheet.getDataRange().getDisplayValues();
  const targetSession = `${dateStr}_${className}`;
  const results = {};
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDateStr = String(data[i][1]).substring(0, 10);
    if (rowDateStr < dateStr) break;
    if (String(data[i][11]) === targetSession) {
      const stdId = String(data[i][5]); 
      if (!results[stdId]) results[stdId] = { area: data[i][7], duty: data[i][8], flag: data[i][9] };
    }
  }
  return results; 
}

function saveMorningActivityBatch(payload) {
  const { date, term, year, className, teacherId, records } = payload;
  if (!verifyTeacherPermission(teacherId, 'HR', className, term, year)) return { status: "error", message: "❌ ความปลอดภัย: คุณไม่ใช่ครูที่ปรึกษาของห้องนี้!" };

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
    for (let i = 1; i < data.length; i++) { if (String(data[i][11]) === sessionID) rowMap[String(data[i][5])] = i + 1; }

    const newRows = [];
    records.forEach(r => {
      const stdId = String(r.studentId);
      if (rowMap[stdId]) sheet.getRange(rowMap[stdId], 8, 1, 3).setValues([[r.area, r.duty, r.flag]]);
      else newRows.push([timestamp, date, term, year, className, stdId, r.studentName, r.area, r.duty, r.flag, teacherId, sessionID]);
    });

    if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
    SpreadsheetApp.flush(); 
    return { status: "success", message: "✅ บันทึกข้อมูลกิจกรรมโฮมรูมเรียบร้อยแล้ว!" };
  } catch (e) { return { status: "error", message: "ระบบกำลังมีผู้ใช้งานพร้อมกันจำนวนมาก กรุณากดบันทึกอีกครั้ง" }; } finally { lock.releaseLock(); }
}

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
    
    if (tTeacherID === String(teacherId).trim().toLowerCase() && (tCode === 'HR' || tName.includes('โฮมรูม')) && String(timeData[i][8]).trim() === String(term).trim() && String(timeData[i][9]).trim() === String(year).trim()) {
      hrClass = `${String(timeData[i][2]).trim()}/${String(timeData[i][3]).trim()}`; break;
    }
  }

  if (!hrClass) return { hasHR: false };

  const targetSession = `${todayStr}_${hrClass}`;
  const latestData = {};

  for (let i = mornData.length - 1; i >= 1; i--) {
    const rowDateStr = String(mornData[i][1]).substring(0, 10);
    if (rowDateStr < todayStr) break;

    if (String(mornData[i][11]) === targetSession) {
      const stdName = String(mornData[i][6]).trim(); 
      if (!latestData[stdName]) latestData[stdName] = { area: String(mornData[i][7]).trim(), duty: String(mornData[i][8]).trim(), flag: String(mornData[i][9]).trim() };
    }
  }

  const summary = { className: hrClass, absent: [], late: [], leave: [], notArea: [], notDuty: [], hasData: Object.keys(latestData).length > 0 };
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
// 🚨 ดึงข้อมูล Dashboard กลุ่มเสี่ยง (0, ร, มส.) สำหรับครู (Ultimate Fix: ดึงชั้นเรียนตามปีที่แอดมินตั้งค่า)
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

    // 🌟 ระบบค้นหาระดับชั้นตาม "ปีการศึกษาที่แอดมินเลือก"
    const targetYear = String(year).trim();
    const studentMap = {};
    const configCur = getSystemConfig();
    const isHistoricalView = String(targetYear).trim() !== String(configCur.year).trim();

    const userSheet = ss.getSheetByName("User_Database");
    if (userSheet) {
        const userData = userSheet.getDataRange().getDisplayValues();
        // 1. User_Database ที่ year ตรง (สำหรับเทอมปัจจุบัน หรือเด็กที่ year ตรง)
        for(let i = 1; i < userData.length; i++) {
            if (String(userData[i][3]).toLowerCase() === 'student' || String(userData[i][3]) === 'นักเรียน') {
                let stdYear = String(userData[i][6]).trim();
                if (stdYear === targetYear) {
                    let rawId = String(userData[i][0]).replace(/'/g, '').trim();
                    studentMap[normID(rawId)] = { displayId: rawId, name: String(userData[i][2]).trim(), cls: String(userData[i][4]).trim(), source: 'user_db_match' };
                }
            }
        }
    }

    // 2. User_History_Database ที่ year ตรง (snapshot ก่อน admin promote) — ใช้ logic ของ admin
    const histSheet = ss.getSheetByName("User_History_Database");
    if (histSheet && histSheet.getLastRow() > 1) {
        const histData = histSheet.getDataRange().getDisplayValues();
        for(let i = 1; i < histData.length; i++) {
            if (String(histData[i][3]).toLowerCase() === 'student' || String(histData[i][3]) === 'นักเรียน') {
                let histYear = String(histData[i][6]).trim();
                let rawId = String(histData[i][0]).replace(/'/g, '').trim();
                let nId = normID(rawId);
                if (!studentMap[nId] && histYear === targetYear) {
                    studentMap[nId] = { displayId: rawId, name: String(histData[i][2]).trim(), cls: String(histData[i][4]).trim(), source: 'user_history' };
                }
            }
        }
    }

    // 3. fallback: ดึง User_Database (ห้องปัจจุบัน) เฉพาะกรณีเทอมปัจจุบัน หรือใช้แค่ name (cls จะ override จาก attClassMap)
    if (userSheet) {
        const userData = userSheet.getDataRange().getDisplayValues();
        for(let i = 1; i < userData.length; i++) {
            if (String(userData[i][3]).toLowerCase() === 'student' || String(userData[i][3]) === 'นักเรียน') {
                let rawId = String(userData[i][0]).replace(/'/g, '').trim();
                let nId = normID(rawId);
                if (!studentMap[nId]) {
                    studentMap[nId] = { displayId: rawId, name: String(userData[i][2]).trim(), cls: String(userData[i][4]).trim(), source: 'user_db_fallback' };
                }
            }
        }
    }

    // ห้องเรียนตามเทอม/ปีนั้นๆ จาก Attendance_Database (ไม่ใช่ห้องปัจจุบัน)
    const attClassMap = {};
    const attSheet = ss.getSheetByName("Attendance_Database");
    if (attSheet) {
        const attData = attSheet.getDataRange().getDisplayValues();
        const targetTerm = String(term).trim();
        for (let i = 1; i < attData.length; i++) {
            const rTerm = String(attData[i][2]).trim();
            const rYear = String(attData[i][3]).trim();
            if (rTerm === targetTerm && rYear === targetYear) {
                const stdId = normID(attData[i][8]);
                const cls = String(attData[i][6]).trim();
                if (stdId && cls && !attClassMap[stdId]) attClassMap[stdId] = cls;
            }
        }
    }

    let riskList = [];
    let count0 = 0, countR = 0, countMS = 0;
    let riskCheckMap = {};
    let debugMatchedRows = 0;
    let debugSampleRows = [];

    const gradeSheet = ss.getSheetByName("Grade_Summary");
    if (gradeSheet) {
        const gradeData = gradeSheet.getDataRange().getDisplayValues();

        const targetTerm = String(term).trim();
        const subjectMaxScore = {};
        for (let i = 1; i < gradeData.length; i++) {
            let subCode = String(gradeData[i][1]).trim();
            let totalScore = parseFloat(gradeData[i][2]) || 0;
            let rTerm = String(gradeData[i][6]).trim();
            let rYear = String(gradeData[i][7]).trim();

            if (rTerm === targetTerm && rYear === targetYear) {
                if (!subjectMaxScore[subCode]) subjectMaxScore[subCode] = 0;
                if (totalScore > subjectMaxScore[subCode]) subjectMaxScore[subCode] = totalScore;
            }
        }

        for (let i = gradeData.length - 1; i >= 1; i--) {
           let rTerm = String(gradeData[i][6]).trim();
           let rYear = String(gradeData[i][7]).trim();

           if (rTerm === targetTerm && rYear === targetYear) {
               debugMatchedRows++;
               if (debugSampleRows.length < 5) {
                 debugSampleRows.push({
                   stdId: gradeData[i][0], subject: gradeData[i][1], grade: gradeData[i][3],
                   remark: gradeData[i][4], term: gradeData[i][6], year: gradeData[i][7]
                 });
               }
               let safeId = normID(gradeData[i][0]);
               let subCode = String(gradeData[i][1]).trim();
               let grade = String(gradeData[i][3]).trim();
               let remark = String(gradeData[i][4] || '').trim();

               if (teacherSubjects[subCode]) {
                   let key = `${safeId}_${subCode}`;
                   
                   if (!riskCheckMap[key]) {
                       riskCheckMap[key] = true;

                       let riskType = null;
                       let isActivitySubject = subCode.startsWith('ก') || subCode.startsWith('I') || subCode.startsWith('i');

                       if (grade === 'ร' || remark === 'ร') riskType = 'ร';
                       else if (grade === 'มส' || remark === 'มส') riskType = 'มส';
                       else if (!isActivitySubject && subjectMaxScore[subCode] > 0) {
                           if (grade === '0' || grade === '0.0') riskType = '0';
                       }

                       if (riskType) {
                           if (riskType === '0') count0++;
                           else if (riskType === 'ร') countR++;
                           else if (riskType === 'มส') countMS++;

                           // ลำดับการเลือกห้อง: user_history (snapshot ปีนั้น) > user_db_match (ตรง year) > attClassMap (จากเช็คชื่อ) > fallback (ห้องปัจจุบัน)
                           let displayClass;
                           const sm = studentMap[safeId];
                           if (sm && (sm.source === 'user_history' || sm.source === 'user_db_match')) {
                               displayClass = sm.cls;
                           } else if (attClassMap[safeId]) {
                               displayClass = attClassMap[safeId];
                           } else {
                               displayClass = sm ? sm.cls : "ไม่ทราบชั้น";
                           }
                           riskList.push({
                               stdId: sm ? sm.displayId : safeId,
                               stdName: sm ? sm.name : "ไม่ทราบชื่อ",
                               className: displayClass,
                               subjectCode: subCode,
                               subjectName: teacherSubjects[subCode],
                               type: riskType
                           });
                       }
                   }
               }
           }
        }
    }

    // นับ source ของ studentMap
    let sourceCount = { user_db_match: 0, user_history: 0, user_db_fallback: 0 };
    Object.values(studentMap).forEach(s => { if (sourceCount[s.source] !== undefined) sourceCount[s.source]++; });

    return {
      status: 'success',
      summary: { zero: count0, r: countR, ms: countMS },
      details: riskList.sort((a, b) => a.className.localeCompare(b.className)),
      debug: {
        paramTerm: String(term).trim(),
        paramYear: String(year).trim(),
        isHistoricalView: isHistoricalView,
        teacherSubjectsCount: Object.keys(teacherSubjects).length,
        teacherSubjects: Object.keys(teacherSubjects),
        matchedGradeRows: debugMatchedRows,
        sampleMatchedRows: debugSampleRows,
        studentMapSize: Object.keys(studentMap).length,
        studentMapSources: sourceCount,
        attClassMapSize: Object.keys(attClassMap).length,
        histSheetRows: histSheet ? histSheet.getLastRow() - 1 : 0
      }
    };

  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function saveStudentRemarkDirectly(studentId, subjectCode, term, year, remarkVal) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const normID = (id) => { let clean = String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, ''); return clean || '0'; };
    const safeId = normID(studentId);
    let finalRemark = (remarkVal === '') ? '-' : remarkVal;

    const gradeSheet = ss.getSheetByName("Grade_Summary");
    if (gradeSheet) {
      if (gradeSheet.getMaxColumns() < 8) gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), 8 - gradeSheet.getMaxColumns());
      
      const data = gradeSheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        const rTerm = String(data[i][6]).trim();
        const rYear = String(data[i][7]).trim();
        
        if (normID(data[i][0]) === safeId && 
            String(data[i][1]).trim().toLowerCase() === String(subjectCode).trim().toLowerCase() &&
            (rTerm === String(term).trim() || rTerm === '') && 
            (rYear === String(year).trim() || rYear === '')) {
          
          gradeSheet.getRange(i + 1, 5).setValue(finalRemark); 
          gradeSheet.getRange(i + 1, 7).setValue(term); 
          gradeSheet.getRange(i + 1, 8).setValue(year); 
          found = true; break;
        }
      }
      if (!found) gradeSheet.appendRow(["'" + studentId, subjectCode, 0, 0, finalRemark, 100, term, year]);
    }

    const scoreSheet = ss.getSheetByName("Score_Database");
    if (scoreSheet) {
      const sData = scoreSheet.getDataRange().getValues();
      let sFound = false;
      for (let i = 1; i < sData.length; i++) {
        if (normID(sData[i][1]) === safeId && String(sData[i][2]).trim().toLowerCase() === String(subjectCode).trim().toLowerCase() && String(sData[i][3]).trim().toLowerCase() === 'remark' && String(sData[i][5]).trim() === String(term).trim() && String(sData[i][6]).trim() === String(year).trim()) {
          scoreSheet.getRange(i + 1, 5).setValue(finalRemark); sFound = true; break;
        }
      }
      if (!sFound && finalRemark !== '-') scoreSheet.appendRow([safeId + "_" + subjectCode + "_remark", "'" + studentId, subjectCode, "remark", finalRemark, term, year]);
    }

    return { success: true, val: remarkVal };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==========================================
// 📚 15. ระบบงานสารบรรณ (Sarabun System - อัปเกรดอิงตาม ปี พ.ศ. ปฏิทิน)
// ==========================================

function requestSarabunNumber(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Sarabun_Database");

    if (!sheet) {
        sheet = ss.insertSheet("Sarabun_Database");
        sheet.appendRow(["Timestamp", "DocType", "DocNumber", "Subject", "TargetDate", "DocTime", "ActionDate", "DocRefNo", "RefDocDate", "DocFrom", "DocTo", "Assignee", "Requester", "Status", "FileURL", "Year"]);
        sheet.getRange("A1:P1").setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
        sheet.setFrozenRows(1);
    }

    // 🌟 เปลี่ยนจากการใช้ "ปีการศึกษา" มาเป็น "ปี พ.ศ. ตามปฏิทินปัจจุบัน" ทันที
    const currentYearBE = new Date().getFullYear() + 543;
    const currentYearStr = String(currentYearBE); 
    
    const docType = String(payload.docType).trim(); 
    const amount = parseInt(payload.amount) || 1; 

    const data = sheet.getDataRange().getDisplayValues();
    let lastNumber = 0;
    
    // 🌟 ย้อนหาเลขล่าสุด "เฉพาะของปี พ.ศ. ปัจจุบัน"
    for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]).trim() === docType && String(data[i][15]).trim() === currentYearStr) {
            lastNumber = parseInt(String(data[i][2]).split('/')[0]) || 0; 
            break; 
        }
    }

    const timestamp = new Date();
    let startNumber = lastNumber + 1;
    let endNumber = lastNumber + amount;
    let rowsToAppend = [];

    for(let i = 0; i < amount; i++) {
        rowsToAppend.push([timestamp, docType, `${lastNumber + 1 + i}/${currentYearStr}`, payload.subject || "-", payload.targetDate || "-", payload.docTime || "-", payload.actionDate || "-", payload.docRefNo || "-", payload.refDocDate || "-", payload.docFrom || "-", payload.docTo || "-", payload.assignee || "-", payload.requester || "Unknown", "ใช้งาน", "", currentYearStr]);
    }

    if(rowsToAppend.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 16).setValues(rowsToAppend);
    SpreadsheetApp.flush(); 

    return { status: "success", message: `✅ สำเร็จ! ดำเนินการออกเลข ${amount} รายการ`, docNumber: amount > 1 ? `${startNumber}/${currentYearStr} ถึง ${endNumber}/${currentYearStr}` : `${startNumber}/${currentYearStr}` };
  } catch (e) { return { status: "error", message: "คิวเต็ม กรุณากดขอเลขใหม่อีกครั้งครับ" }; } finally { lock.releaseLock(); }
}

function getSarabunHistory(requesterName, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sarabun_Database");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  const results = [];

  // 🌟 ยกเลิกการกรองตาม "ปีการศึกษา" เพื่อให้เห็นประวัติย้อนหลังของปี พ.ศ. เก่าๆ ด้วย
  for (let i = data.length - 1; i >= 1; i--) {
     const row = data[i];
     // ถ้าเป็น Admin ให้เห็นทั้งหมด / ถ้าเป็นครูให้เห็นเฉพาะของตัวเอง
     if (role.toUpperCase() === 'ADMIN' || String(row[12]).trim() === requesterName) {
        results.push({ 
            id: i + 1, timestamp: row[0], docType: row[1], docNumber: row[2], 
            subject: row[3], targetDate: row[4], docTime: row[5], actionDate: row[6], 
            docRefNo: row[7], refDocDate: row[8], docFrom: row[9], docTo: row[10], 
            assignee: row[11], requester: row[12], status: row[13], fileUrl: row[14],
            year: row[15] // ส่งเลขปี พ.ศ. กลับไปเผื่อหน้าเว็บใช้ประโยชน์
        });
     }
  }
  return results;
}

function uploadSarabunFile(id, base64Data, filename, docNumber) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const safeDocNum = String(docNumber).replace(/\//g, '-'); 
    const fileUrl = uploadFileToDrive(base64Data, `Sarabun_${safeDocNum}_${filename}`);
    if (fileUrl.startsWith("Error")) throw new Error(fileUrl);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sarabun_Database");
    if (!sheet) throw new Error("ไม่พบฐานข้อมูล");

    const targetDocNum = String(docNumber).trim();
    const currentDocNum = String(sheet.getRange(parseInt(id), 3).getDisplayValue()).trim(); 

    if (currentDocNum === targetDocNum) { sheet.getRange(parseInt(id), 15).setValue(fileUrl); } 
    else {
        const allData = sheet.getDataRange().getDisplayValues();
        let found = false;
        for(let i = 1; i < allData.length; i++) {
            if(String(allData[i][2]).trim() === targetDocNum) { sheet.getRange(i + 1, 15).setValue(fileUrl); found = true; break; }
        }
        if(!found) throw new Error(`ไม่พบเอกสารเลขที่ ${targetDocNum} ในระบบ`);
    }
    return { status: "success", message: "แนบไฟล์เสร็จสมบูรณ์" };
  } catch(e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

// ==========================================
// 📚 คลังตัวชี้วัดและผลการเรียนรู้ (Curriculum Database)
// ==========================================

function setupCurriculumDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Curriculum_Database");
  if (!sheet) {
    sheet = ss.insertSheet("Curriculum_Database");
    sheet.appendRow(["SubjectCode", "SubjectType", "StandardCode", "Description", "EvalType"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
    sheet.setFrozenRows(1);
    return "✅ สร้างฐานข้อมูล Curriculum_Database เรียบร้อยแล้ว!";
  }
  return "ฐานข้อมูล Curriculum_Database มีอยู่แล้วครับ";
}

function getCurriculumData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Curriculum_Database");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({ subjectCode: row[0], subjectType: row[1], standardCode: row[2], description: row[3], evalType: row[4] }));
}

function importCurriculumCSV(base64Data, clearOldData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Curriculum_Database");
    if (!sheet) return { status: 'error', message: 'ไม่พบชีต Curriculum_Database กรุณากดปุ่มสร้างฐานข้อมูลก่อนครับ' };

    const decoded = Utilities.base64Decode(base64Data);
    const csvText = Utilities.newBlob(decoded).getDataAsString('UTF-8');
    const csv = Utilities.parseCsv(csvText);

    if (clearOldData && sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();

    let newRows = [];
    for (let i = 1; i < csv.length; i++) {
      if (!csv[i][0]) continue; 
      newRows.push([String(csv[i][0]).trim(), String(csv[i][1]).trim(), String(csv[i][2]).trim(), String(csv[i][3]).trim(), String(csv[i][4] || "-").trim()]);
    }

    if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
    return { status: 'success', message: `✅ นำเข้าข้อมูลตัวชี้วัด/ผลการเรียนรู้ สำเร็จ ${newRows.length} รายการ` };
  } catch (e) { return { status: 'error', message: '❌ ข้อผิดพลาด: ' + e.message }; } finally { lock.releaseLock(); }
}

function getCurriculumBySubject(subjectCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Curriculum_Database");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  const results = [];
  const cleanCode = String(subjectCode).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanCode) {
      results.push({ subjectCode: data[i][0], subjectType: data[i][1], standardCode: data[i][2], description: data[i][3], evalType: data[i][4] });
    }
  }
  return results;
}

// ==========================================
// 🚀 ระบบเลื่อนชั้นประจำปี (Annual Student Promotion)
// ==========================================
function promoteStudentsToNextYear() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // เผื่อเวลาประมวลผลเด็กทั้งโรงเรียน
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("User_Database");
    if (!userSheet) return { status: "error", message: "ไม่พบฐานข้อมูล User_Database" };

    const config = getSystemConfig();
    const currentYear = parseInt(config.year); // ปีใหม่ที่แอดมินเพิ่งตั้งค่า (เช่น 2569)
    if (!currentYear) return { status: "error", message: "อ่านค่าปีการศึกษาปัจจุบันไม่สำเร็จ" };

    const data = userSheet.getDataRange().getValues();
    let updateCount = 0;
    let graduateCount = 0;
    let eligibleCount = 0;

    const historyToAppend = []; // กล่องเก็บรายชื่อเพื่อนำไปเป็นประวัติ

    for (let i = 1; i < data.length; i++) {
      const role = String(data[i][3]).trim().toLowerCase();
      const status = String(data[i][7]).trim();
      const stdYear = parseInt(data[i][6]) || 0;
      
      // หารายชื่อเด็กที่สถานะปกติ และปีการศึกษา "น้อยกว่า" ปีปัจจุบันที่แอดมินเพิ่งเปลี่ยน
      if ((role === 'student' || role === 'นักเรียน') && status === 'ปกติ' && stdYear < currentYear) {
        eligibleCount++;
        // 🌟 1. ดึงข้อมูลก่อนที่จะถูกเปลี่ยนห้อง ไปเก็บลงกล่องประวัติ
        historyToAppend.push([...data[i]]);
      }
    }

    if (eligibleCount === 0) {
      return { 
        status: "error", 
        message: "⚠️ ไม่พบนักเรียนที่เข้าเงื่อนไขการเลื่อนชั้น!\n(คุณอาจจะยังไม่ได้เปลี่ยน 'ปีการศึกษา' ในหน้าตั้งค่าระบบ หรือนักเรียนถูกเลื่อนชั้นไปหมดแล้ว)" 
      };
    }

    // 🌟 2. สร้างชีต User_History_Database (ถ้ายังไม่มี) และบันทึกประวัติลงไป
    let histSheet = ss.getSheetByName("User_History_Database");
    if (!histSheet) {
        histSheet = ss.insertSheet("User_History_Database");
        histSheet.appendRow(["Username", "Password", "FullName", "Role", "Department", "Email", "Year", "Status"]);
        histSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#ea4335").setFontColor("white");
    }
    if (historyToAppend.length > 0) {
        histSheet.getRange(histSheet.getLastRow() + 1, 1, historyToAppend.length, historyToAppend[0].length).setValues(historyToAppend);
    }

    // 🌟 3. สร้างไฟล์ Backup ยกล็อตกันเหนียวให้อีกชั้นนึง
    const backupName = `User_Database_Backup_${currentYear - 1}`;
    if (!ss.getSheetByName(backupName)) {
      const backupSheet = userSheet.copyTo(ss);
      backupSheet.setName(backupName);
    }

    // 🌟 4. เริ่มกระบวนการเลื่อนชั้น (เปลี่ยนค่าในระบบ)
    for (let i = 1; i < data.length; i++) {
      const role = String(data[i][3]).trim().toLowerCase();
      let status = String(data[i][7]).trim();
      let stdYear = parseInt(data[i][6]) || 0;

      if ((role === 'student' || role === 'นักเรียน') && status === 'ปกติ' && stdYear < currentYear) {
        let cls = String(data[i][4]).trim();
        let match = cls.match(/ม\.(\d+)\/(\d+)/); // แยกตัวเลขชั้นกับห้อง

        if (match) {
          let level = parseInt(match[1]);
          let room = parseInt(match[2]);

          if (level === 3 || level === 6) {
            // ม.3 และ ม.6 สั่งให้จบการศึกษา
            data[i][7] = "จบการศึกษา";
            graduateCount++;
          } else {
            // ชั้นอื่นๆ เลื่อนขึ้น 1 ระดับ
            level++;
            data[i][4] = `ม.${level}/${room}`;
            updateCount++;
          }
        }
        // อัปเดตปีใหม่ให้เด็ก
        data[i][6] = currentYear; 
      }
    }

    // 🌟 5. เทข้อมูลที่เปลี่ยนแล้วกลับลงไปในชีต
    userSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();

    return {
      status: "success",
      message: `✅ เลื่อนชั้นเรียนสำเร็จทั้งหมด ${updateCount} คน\n🎓 จบการศึกษา (ม.3, ม.6) จำนวน ${graduateCount} คน\n\n🛡️ ระบบได้จัดเก็บประวัติห้องเรียนเดิมไว้ใน User_History_Database เรียบร้อยแล้ว (สามารถย้อนปีกลับไปดู/แก้ไขคะแนนปีเก่าได้ปกติ 100%)`
    };

  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 📅 16. ระบบปฏิทินปฏิบัติงานโรงเรียน (School Calendar)
// ==========================================

function setupCalendarDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Calendar_Database");
  if (!sheet) {
      sheet = ss.insertSheet("Calendar_Database");
      sheet.appendRow(["ID", "Title", "Start", "End", "Color", "Description", "CreatedBy", "Timestamp"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#e83e8c").setFontColor("white");
      sheet.setFrozenRows(1);
      return "✅ สร้างฐานข้อมูล Calendar_Database เรียบร้อยแล้ว!";
  }
  return "ฐานข้อมูล Calendar_Database มีอยู่แล้วครับ";
}

function getCalendarEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar_Database");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const events = [];
  for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      events.push({
          id: data[i][0],
          title: data[i][1],
          start: data[i][2],
          end: data[i][3] !== "" ? data[i][3] : null,
          backgroundColor: data[i][4],
          borderColor: data[i][4],
          description: data[i][5],
          createdBy: data[i][6]
      });
  }
  return events;
}

function saveCalendarEvent(payload) {
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("Calendar_Database");
      if (!sheet) throw new Error("ไม่พบชีตปฏิทิน กรุณากด Setup DB ก่อน");

      const data = sheet.getDataRange().getValues();
      const timestamp = new Date();

      // ถ้ามีการส่ง ID มา แปลว่าอัปเดตของเดิม
      if (payload.id && payload.id !== "") {
          for (let i = 1; i < data.length; i++) {
              if (String(data[i][0]) === payload.id) {
                  sheet.getRange(i + 1, 2, 1, 7).setValues([[
                      payload.title, payload.start, payload.end, payload.color, 
                      payload.description, payload.createdBy, timestamp
                  ]]);
                  return { status: "success", message: "อัปเดตกิจกรรมเรียบร้อย" };
              }
          }
      }

      // ถ้าไม่มี ID แปลว่าสร้างใหม่ สร้าง ID โดยใช้วันที่เวลาชนกัน
      const newId = "EVT" + timestamp.getTime();
      sheet.appendRow([newId, payload.title, payload.start, payload.end, payload.color, payload.description, payload.createdBy, timestamp]);
      return { status: "success", message: "เพิ่มกิจกรรมเรียบร้อย" };

  } catch(e) {
      return { status: "error", message: e.message };
  } finally {
      lock.releaseLock();
  }
}

function deleteCalendarEvent(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Calendar_Database");
  if (!sheet) return { status: "error", message: "ไม่พบชีตปฏิทิน" };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
          sheet.deleteRow(i + 1);
          return { status: "success", message: "ลบสำเร็จ" };
      }
  }
  return { status: "error", message: "ไม่พบกิจกรรมที่ต้องการลบ" };
}

function importCalendarCSV(base64Data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Calendar_Database");
    if (!sheet) throw new Error("ไม่พบชีตปฏิทิน กรุณากด Setup DB ก่อน");

    const decoded = Utilities.base64Decode(base64Data);
    const csv = Utilities.parseCsv(Utilities.newBlob(decoded).getDataAsString('UTF-8'));

    const thaiMonths = {
      'ม.ค.': 1, 'ก.พ.': 2, 'มี.ค.': 3, 'เม.ย.': 4,
      'พ.ค.': 5, 'มิ.ย.': 6, 'ก.ค.': 7, 'ส.ค.': 8,
      'ก.ย.': 9, 'ต.ค.': 10, 'พ.ย.': 11, 'ธ.ค.': 12
    };

    function parseThaiDate(str) {
      str = String(str || '').trim();
      if (!str) return null;
      if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return { start: str, end: null };

      for (const abbr in thaiMonths) {
        const idx = str.indexOf(abbr);
        if (idx === -1) continue;

        const month = thaiMonths[abbr];
        const afterMonth = str.substring(idx + abbr.length).trim();
        const yearMatch = afterMonth.match(/^(\d{2,4})/);
        if (!yearMatch) continue;

        const y = parseInt(yearMatch[1]);
        const year = (y < 100 ? 2500 + y : y) - 543;

        const beforeMonth = str.substring(0, idx).trim();
        const dayMatch = beforeMonth.match(/^(\d{1,2})(?:-(\d{1,2}))?/);
        if (!dayMatch) continue;

        const mm = String(month).padStart(2, '0');
        const startDay = String(dayMatch[1]).padStart(2, '0');
        const startStr = `${year}-${mm}-${startDay}`;

        if (dayMatch[2]) {
          const endDay = String(dayMatch[2]).padStart(2, '0');
          return { start: startStr, end: `${year}-${mm}-${endDay}` };
        }
        return { start: startStr, end: null };
      }
      return null;
    }

    function autoColor(text) {
      const t = String(text || '').toLowerCase();
      if (/หยุด|วิสาข|มาฆ|อาสาฬห|เข้าพรรษา|ออกพรรษา/.test(t)) return '#dc3545';
      if (/สอบ|ทดสอบ|วัดผล|ประเมิน|นิเทศ/.test(t)) return '#198754';
      if (/ประชุม|สัมมนา|อบรม/.test(t)) return '#ffc107';
      return '#0d6efd';
    }

    // ข้ามแถว header ถ้ามี
    let startRow = 0;
    if (csv.length > 0 && /วันที่|วันเดือน|date/i.test(String(csv[0][0] || ''))) startRow = 1;

    const rows = [];
    let skipped = 0;
    const now = new Date();

    for (let i = startRow; i < csv.length; i++) {
      const row = csv[i];
      const dateStr = String(row[0] || '').trim();
      const title = String(row[1] || '').trim();
      if (!dateStr || !title) continue;

      const description = String(row[2] || '').trim();
      let color = String(row[3] || '').trim();
      if (!color || !color.startsWith('#')) color = autoColor(title);

      const parsed = parseThaiDate(dateStr);
      if (!parsed) { skipped++; continue; }

      // FullCalendar ใช้ exclusive end date (+1 วัน)
      let endDate = '';
      if (parsed.end) {
        const e = new Date(parsed.end + 'T00:00:00');
        e.setDate(e.getDate() + 1);
        endDate = Utilities.formatDate(e, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      rows.push(['EVT' + now.getTime() + '_' + i, title, parsed.start, endDate, color, description, 'Import', now]);
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
    }

    return {
      status: 'success',
      imported: rows.length,
      skipped,
      message: `นำเข้าสำเร็จ ${rows.length} รายการ${skipped > 0 ? ` (ข้าม ${skipped} รายการ วันที่ไม่ถูกต้อง)` : ''}`
    };
  } catch(e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}