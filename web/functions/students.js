const { query } = require('../lib/db');

function normalizeClass(str) {
  return String(str || '').replace(/[^a-zA-Z0-9ก-๙]/g, '').toLowerCase();
}

async function getStudentsByClass([className, year]) {
  const norm = normalizeClass(className);
  const config = year
    ? { year }
    : await require('./getSystemConfig')();
  const y = year || config.year;

  const { rows } = await query(
    `SELECT username, password, full_name, role, department, email, year, status
     FROM users WHERE UPPER(role)='STUDENT' AND year=$1 AND status='ปกติ'
     ORDER BY username`,
    [y]
  );

  let matched = rows.filter(r => normalizeClass(r.department) === norm);
  if (matched.length === 0) {
    matched = rows.filter(r => {
      const d = String(r.department || '');
      return d === className || normalizeClass(d) === norm;
    });
  }
  return matched.map(r => [
    r.username, r.password, r.full_name, r.role,
    r.department || '', r.email || '', r.year || '', r.status || 'ปกติ',
  ]);
}

async function getStudentsByClub([clubId]) {
  const { rows } = await query(
    `SELECT u.username, u.password, u.full_name, u.role,
            u.department, u.email, u.year, u.status,
            cm.class_name
     FROM club_members cm
     JOIN users u ON u.username = cm.student_id
     WHERE cm.club_id = $1
     ORDER BY u.username`,
    [clubId]
  );
  return rows.map(r => [
    r.username, r.password, r.full_name, r.role,
    r.department || r.class_name || '', r.email || '', r.year || '', r.status || 'ปกติ',
  ]);
}

module.exports = { getStudentsByClass, getStudentsByClub };
