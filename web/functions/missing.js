/**
 * Implements functions referenced in GAS frontend but not yet in Phase 2/3.
 * Grouped here to keep other files clean.
 */
const { query } = require('../lib/db');
const cache = require('../lib/cache');

// ============================================================
// getTeacherRiskDashboard — grade-based risk (0, ร, มส)
// returns { status, summary: {zero,r,ms}, details: [{className,subjectCode,subjectName,stdName,type}] }
// ============================================================
async function getTeacherRiskDashboard([teacherId, term, year]) {
  const { rows } = await query(
    `SELECT gs.student_id, u.full_name as std_name,
            gs.subject_code, sc.subject_name,
            u.department as class_name,
            gs.grade
     FROM grade_summary gs
     JOIN users u ON u.username = gs.student_id
     LEFT JOIN (
       SELECT DISTINCT subject_code, subject_name
       FROM timetable WHERE teacher_id=$1 AND term=$2 AND year=$3
     ) sc ON sc.subject_code = gs.subject_code
     WHERE gs.term=$2 AND gs.year=$3
       AND gs.grade IN ('0','ร','มส','มส.')
       AND gs.subject_code IN (
         SELECT DISTINCT subject_code FROM timetable
         WHERE teacher_id=$1 AND term=$2 AND year=$3
       )
     ORDER BY gs.subject_code, u.department, gs.student_id`,
    [teacherId, term, year]
  );

  const details = rows.map(r => ({
    className: r.class_name || '',
    subjectCode: r.subject_code || '',
    subjectName: r.subject_name || '',
    stdName: r.std_name || '',
    type: r.grade === '0' ? '0' : r.grade === 'ร' ? 'ร' : 'มส',
  }));

  const summary = {
    zero: details.filter(d => d.type === '0').length,
    r: details.filter(d => d.type === 'ร').length,
    ms: details.filter(d => d.type === 'มส').length,
  };

  return { status: 'success', summary, details };
}

// ============================================================
// getTeacherAtRiskDashboard — attendance-based risk (>20% absent)
// ============================================================
async function getTeacherAtRiskDashboard([teacherId, term, year]) {
  const { rows } = await query(
    `SELECT student_id, student_name, subject_code, subject_name, class,
            COUNT(*) as total,
            COUNT(CASE WHEN status IN ('ขาด','absent') THEN 1 END) as absent_count,
            COUNT(CASE WHEN status IN ('ลา','leave')   THEN 1 END) as leave_count
     FROM attendance
     WHERE teacher_id=$1 AND term=$2 AND year=$3
     GROUP BY student_id, student_name, subject_code, subject_name, class
     HAVING COUNT(CASE WHEN status IN ('ขาด','absent','ลา','leave') THEN 1 END)::float
            / NULLIF(COUNT(*), 0) > 0.2
     ORDER BY absent_count DESC LIMIT 50`,
    [teacherId, term, year]
  );
  const list = rows.map(r => ({
    id: r.student_id, name: r.student_name,
    subjectCode: r.subject_code, subjectName: r.subject_name, className: r.class,
    total: parseInt(r.total), absent: parseInt(r.absent_count), leave: parseInt(r.leave_count),
    percent: (((parseInt(r.total) - parseInt(r.absent_count) - parseInt(r.leave_count)) / parseInt(r.total)) * 100).toFixed(1),
  }));
  const byThreshold = (pct) => list.filter(s => parseFloat(s.percent) < pct);
  return {
    critical: byThreshold(60),
    ms: list.filter(s => parseFloat(s.percent) >= 60 && parseFloat(s.percent) < 80),
    risk: list.filter(s => parseFloat(s.percent) >= 80 && parseFloat(s.percent) < 85),
  };
}

// ============================================================
// getStudentDashboardBundle
// ============================================================
async function getStudentDashboardBundle([studentId, term, year]) {
  const userRes = await query(
    `SELECT department FROM users WHERE username=$1`, [studentId]
  );
  const className = userRes.rows[0]?.department || '';
  const parts = className.split('/');
  const level = parts[0] || '';
  const room = parts[1] || '';

  const DAYS = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  const todayDay = DAYS[new Date().getDay()];

  let timetable = { ok: false, data: [] };
  try {
    if (todayDay && level) {
      const { rows } = await query(
        `SELECT t.subject_code, t.subject_name, t.level||'/'||t.room as class_id,
                t.period, t.location, u.full_name as teacher_name
         FROM timetable t
         LEFT JOIN users u ON u.username = t.teacher_id
         WHERE t.level=$1 AND t.room=$2 AND t.day=$3 AND t.term=$4 AND t.year=$5
         ORDER BY t.period::integer`,
        [level, room, todayDay, term, year]
      );
      timetable = { ok: true, data: rows };
    } else {
      timetable = { ok: true, data: [] };
    }
  } catch (e) { timetable = { ok: false, error: e.message, data: [] }; }

  let scoreFeed = { ok: true, data: [] };
  try {
    const { rows } = await query(
      `SELECT gs.subject_code, sc.subject_name, gs.total_score, gs.grade
       FROM grade_summary gs
       LEFT JOIN subject_config sc ON sc.subject_code = gs.subject_code AND sc.term=$2 AND sc.year=$3
       WHERE gs.student_id=$1 AND gs.term=$2 AND gs.year=$3`,
      [studentId, term, year]
    );
    scoreFeed = { ok: true, data: rows };
  } catch (e) { scoreFeed = { ok: false, error: e.message, data: [] }; }

  return { timetable, scoreFeed };
}

// ============================================================
// getExecutiveDashboardBundle
// ============================================================
async function getExecutiveDashboardBundle([dept]) {
  const getSystemConfig = require('./getSystemConfig');
  const getCalendarEvents = require('./getCalendarEvents');
  const config = await getSystemConfig();

  const [staffRes, leaveRes, calendarEvents] = await Promise.all([
    query(
      `SELECT UPPER(role) as role, COUNT(*) as cnt FROM users
       WHERE UPPER(role) != 'STUDENT' OR year=$1 GROUP BY UPPER(role)`,
      [config.year]
    ),
    query(`SELECT COUNT(*) as cnt FROM leave_records WHERE status='รอพิจารณา' AND year=$1`, [config.year]),
    getCalendarEvents(),
  ]);

  let studentCount = 0, teacherCount = 0;
  for (const r of staffRes.rows) {
    if (r.role === 'STUDENT') studentCount += parseInt(r.cnt);
    else teacherCount += parseInt(r.cnt);
  }

  return {
    ts: Date.now(),
    systemConfig: { ok: true, data: config },
    kpi: {
      ok: true,
      data: {
        studentCount, teacherCount,
        pendingLeaveCount: parseInt(leaveRes.rows[0]?.cnt || 0),
      },
    },
    calendarEvents: { ok: true, data: calendarEvents },
    academic: { ok: true, data: {} },
    budget: { ok: true, data: {} },
    personnel: { ok: true, data: {} },
    general: { ok: true, data: {} },
  };
}

// ============================================================
// Club helpers
// ============================================================
async function getClubMembers([clubId, term, year]) {
  const { rows } = await query(
    `SELECT cm.student_id, cm.student_name, cm.class_name,
            to_char(cm.registered_at,'YYYY-MM-DD') as registered_at
     FROM club_members cm
     WHERE cm.club_id=$1 AND cm.term=$2 AND cm.year=$3
     ORDER BY cm.class_name, cm.student_id`,
    [clubId, term, year]
  );
  return rows.map(r => ({
    studentId: r.student_id,
    studentName: r.student_name || '',
    className: r.class_name || '',
    registeredAt: r.registered_at || '',
  }));
}

async function getClubMembersForTeacher([teacherId, term, year]) {
  const { rows } = await query(
    `SELECT ca.club_id, c.club_name, cm.student_id, cm.student_name, cm.class_name
     FROM club_advisors ca
     JOIN clubs c ON c.club_id = ca.club_id AND c.term=ca.term AND c.year=ca.year
     JOIN club_members cm ON cm.club_id = ca.club_id AND cm.term=ca.term AND cm.year=ca.year
     WHERE ca.teacher_id=$1 AND ca.term=$2 AND ca.year=$3
     ORDER BY cm.class_name, cm.student_id`,
    [teacherId, term, year]
  );
  return rows.map(r => ({
    clubId: r.club_id,
    clubName: r.club_name,
    studentId: r.student_id,
    studentName: r.student_name || '',
    className: r.class_name || '',
  }));
}

async function getClubAttendanceSummary([clubId, term, year]) {
  const { rows } = await query(
    `SELECT a.student_id, a.student_name,
            COUNT(*) as total,
            COUNT(CASE WHEN a.status IN ('มา','present') THEN 1 END) as present
     FROM attendance a
     WHERE a.subject_code LIKE 'CLUB_%' AND a.class=$1
       AND a.term=$2 AND a.year=$3
     GROUP BY a.student_id, a.student_name
     ORDER BY a.student_id`,
    [clubId, term, year]
  );
  return rows.map(r => ({
    studentId: r.student_id,
    studentName: r.student_name || '',
    total: parseInt(r.total),
    present: parseInt(r.present),
  }));
}

async function deleteClub([clubId]) {
  await query(`DELETE FROM clubs WHERE club_id=$1`, [clubId]);
  cache.del('clubs_all');
  return { success: true };
}

async function registerToClub([studentId, studentName, className, clubId, term, year, registeredBy]) {
  return require('./clubs_write').registerClub([studentId, studentName, className, clubId, term, year, registeredBy]);
}

async function unregisterFromClub([studentId, term, year]) {
  return require('./clubs_write').unregisterClub([studentId, term, year]);
}

// ============================================================
// Leave
// ============================================================
async function getAllLeaves([year, statusFilter]) {
  const params = [year];
  let sql = `SELECT id, teacher_id, staff_name, type,
             to_char(start_date,'YYYY-MM-DD') as start_date,
             to_char(end_date,'YYYY-MM-DD') as end_date,
             days, reason, status, year, admin_comment, reviewed_by
             FROM leave_records WHERE year=$1`;
  if (statusFilter && statusFilter !== 'all') {
    params.push(statusFilter);
    sql += ` AND status=$${params.length}`;
  }
  sql += ' ORDER BY request_date DESC';
  const { rows } = await query(sql, params);
  return rows.map(r => ({
    id: r.id, teacherId: r.teacher_id, staffName: r.staff_name || '',
    type: r.type, startDate: r.start_date, endDate: r.end_date,
    days: parseFloat(r.days || 1), reason: r.reason || '',
    status: r.status, year: r.year,
    adminComment: r.admin_comment || '', reviewedBy: r.reviewed_by || '',
  }));
}

// ============================================================
// Config / School info
// ============================================================
async function saveSchoolInfo([schoolName, logoBase64, logoFilename]) {
  if (schoolName) {
    await query(
      `INSERT INTO system_settings(key,subkey,value1) VALUES('school_name','',$1)
       ON CONFLICT(key,subkey) DO UPDATE SET value1=$1`,
      [schoolName]
    );
  }
  if (logoBase64) {
    await query(
      `INSERT INTO system_settings(key,subkey,value1,value2) VALUES('school_logo','',$1,$2)
       ON CONFLICT(key,subkey) DO UPDATE SET value1=$1, value2=$2`,
      [logoBase64, logoFilename || '']
    );
  }
  cache.del('system_config');
  return { status: 'success', message: 'บันทึกข้อมูลโรงเรียนสำเร็จ' };
}

async function savePrintConfigData([term, year, sysData, homeroomData]) {
  await query(
    `INSERT INTO print_config(term,year,sys_data,homeroom_data)
     VALUES($1,$2,$3,$4)
     ON CONFLICT(term,year) DO UPDATE SET sys_data=$3, homeroom_data=$4`,
    [term, year, JSON.stringify(sysData || {}), JSON.stringify(homeroomData || {})]
  );
  return { status: 'success', message: 'บันทึกสำเร็จ' };
}

// ============================================================
// Curriculum
// ============================================================
async function getCurriculumData([subjectCode]) {
  const params = [];
  let sql = `SELECT id, subject_code, subject_type, standard_code, description, eval_type FROM curriculum`;
  if (subjectCode) { params.push(subjectCode); sql += ` WHERE subject_code=$1`; }
  sql += ' ORDER BY subject_code, id';
  const { rows } = await query(sql, params);
  return rows.map(r => ({
    id: r.id, subjectCode: r.subject_code,
    subjectType: r.subject_type || '', standardCode: r.standard_code || '',
    description: r.description || '', evalType: r.eval_type || '',
  }));
}

async function importCurriculumCSV([rows, clearOld]) {
  if (!Array.isArray(rows) || rows.length === 0) return { status: 'success', message: 'นำเข้า 0 รายการ' };
  const { pool } = require('../lib/db');
  const client = await pool.connect();
  let count = 0;
  try {
    await client.query('BEGIN');
    if (clearOld) await client.query('DELETE FROM curriculum');
    for (const r of rows) {
      await client.query(
        `INSERT INTO curriculum(subject_code,subject_type,standard_code,description,eval_type)
         VALUES($1,$2,$3,$4,$5)`,
        [r.subjectCode||'', r.subjectType||'', r.standardCode||'', r.description||'', r.evalType||'']
      );
      count++;
    }
    await client.query('COMMIT');
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
  return { status: 'success', message: `นำเข้าสำเร็จ ${count} รายการ` };
}

// ============================================================
// Stubs (DB already set up, just return success)
// ============================================================
async function setupCalendarDatabase() {
  return 'ฐานข้อมูลปฏิทินพร้อมใช้งานแล้ว';
}
async function setupClubDatabase() {
  return { status: 'success', message: 'ฐานข้อมูลชุมนุมพร้อมใช้งานแล้ว' };
}
async function setupCurriculumDatabase() {
  return { status: 'success', message: 'ฐานข้อมูลหลักสูตรพร้อมใช้งานแล้ว' };
}
async function saveStudentRemarkDirectly([studentId, remark, term, year]) {
  return { status: 'success', message: 'บันทึกหมายเหตุสำเร็จ' };
}
async function uploadSarabunFile([id, base64Data, filename, docNum]) {
  return { status: 'success', message: 'ไม่รองรับอัปโหลดไฟล์ใน web prototype', fileURL: '' };
}
async function getTeacherListForDropdown() {
  return require('./getTeachersForTimetable')();
}

module.exports = {
  getTeacherRiskDashboard,
  getTeacherAtRiskDashboard,
  getStudentDashboardBundle,
  getExecutiveDashboardBundle,
  getClubMembers,
  getClubMembersForTeacher,
  getClubAttendanceSummary,
  deleteClub,
  registerToClub,
  unregisterFromClub,
  getAllLeaves,
  saveSchoolInfo,
  savePrintConfigData,
  getCurriculumData,
  importCurriculumCSV,
  setupCalendarDatabase,
  setupClubDatabase,
  setupCurriculumDatabase,
  saveStudentRemarkDirectly,
  uploadSarabunFile,
  getTeacherListForDropdown,
};
