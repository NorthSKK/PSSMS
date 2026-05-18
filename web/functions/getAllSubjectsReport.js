const { query } = require('../lib/db');

const WEEKS_PER_TERM = 20;

module.exports = async function getAllSubjectsReport([teacherId, term, year]) {
  // Get periods-per-week per subject-class from timetable
  const ppwRes = await query(
    `SELECT subject_code, subject_name, level||'/'||room as class_id, COUNT(*) as ppw
     FROM timetable WHERE teacher_id=$1 AND term=$2 AND year=$3
     GROUP BY subject_code, subject_name, level, room`,
    [teacherId, term, year]
  );
  const ppwMap = {};
  const nameMap = {};
  for (const r of ppwRes.rows) {
    const k = `${r.subject_code}||${r.class_id}`;
    ppwMap[k] = parseInt(r.ppw) || 1;
    nameMap[k] = r.subject_name;
  }

  // Attendance aggregation
  const { rows } = await query(
    `SELECT student_id, student_name, subject_code, class as class_id,
            COUNT(DISTINCT session_id) as taught,
            COUNT(CASE WHEN status IN ('มา','present') THEN 1 END) as present,
            COUNT(CASE WHEN status IN ('สาย','late')   THEN 1 END) as late,
            COUNT(CASE WHEN status IN ('ลา','leave')   THEN 1 END) as leave,
            COUNT(CASE WHEN status IN ('ขาด','absent') THEN 1 END) as absent
     FROM attendance
     WHERE teacher_id=$1 AND term=$2 AND year=$3
     GROUP BY student_id, student_name, subject_code, class
     HAVING COUNT(*) > 0
     ORDER BY subject_code, class, student_id`,
    [teacherId, term, year]
  );

  return rows.map(r => {
    const k = `${r.subject_code}||${r.class_id}`;
    const ppw = ppwMap[k] || 3;
    const totalPeriods = ppw * WEEKS_PER_TERM;
    const absent = parseInt(r.absent), leave = parseInt(r.leave);
    const missed = absent + leave;
    return {
      id: r.student_id, name: r.student_name,
      subjectCode: r.subject_code, subjectName: nameMap[k] || '',
      className: r.class_id,
      present: parseInt(r.present), late: parseInt(r.late), leave, absent,
      percent: (((totalPeriods - missed) / totalPeriods) * 100).toFixed(2),
      taught: parseInt(r.taught),
      totalCoursePeriods: totalPeriods,
      remainingQuota: Math.floor(totalPeriods * 0.2) - missed,
    };
  });
};
