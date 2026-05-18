const { query } = require('../lib/db');
const getSystemConfig = require('./getSystemConfig');
const getCalendarEvents = require('./getCalendarEvents');
const { getTeacherTimetableWithStatus } = require('./timetable');

function section(fn) {
  return fn().then(data => ({ ok: true, data })).catch(e => ({ ok: false, error: e.message }));
}

async function getAtRiskStudents(teacherId, term, year) {
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
     ORDER BY absent_count DESC LIMIT 30`,
    [teacherId, term, year]
  );
  return rows.map(r => ({
    id: r.student_id, name: r.student_name,
    subjectCode: r.subject_code, subjectName: r.subject_name, className: r.class,
    total: parseInt(r.total),
    absent: parseInt(r.absent_count),
    leave: parseInt(r.leave_count),
    percent: (((parseInt(r.total) - parseInt(r.absent_count) - parseInt(r.leave_count)) / parseInt(r.total)) * 100).toFixed(1),
  }));
}

module.exports = async function getTeacherDashboardBundle([teacherId, term, year]) {
  const config = await getSystemConfig();
  const t = term || config.term;
  const y = year || config.year;

  const [timetable, calendarEvents, atRiskDashboard] = await Promise.all([
    section(() => getTeacherTimetableWithStatus([teacherId])),
    section(() => getCalendarEvents()),
    section(() => getAtRiskStudents(teacherId, t, y)),
  ]);

  return { ts: Date.now(), timetable, calendarEvents, atRiskDashboard };
};
