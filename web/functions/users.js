const { query } = require('../lib/db');
const cache = require('../lib/cache');

function invalidateUsers() {
  cache.del('all_users');
}

async function addUser([userData]) {
  const u = userData || {};
  await query(
    `INSERT INTO users(username, password, full_name, role, department, email, year, status)
     VALUES($1,$2,$3,$4,$5,$6,$7,$8)
     ON CONFLICT (username) DO UPDATE SET
       password=$2, full_name=$3, role=$4, department=$5, email=$6, year=$7, status=$8`,
    [
      String(u.username || '').trim(),
      String(u.password || '').trim(),
      String(u.fullName || u.full_name || '').trim(),
      String(u.role || 'Teacher'),
      String(u.department || u.dept || '').trim(),
      String(u.email || '').trim(),
      String(u.year || '').trim(),
      String(u.status || 'ปกติ'),
    ]
  );
  invalidateUsers();
  return { status: 'success', message: 'บันทึกสำเร็จ' };
}

async function editUser([username, userData]) {
  const u = userData || {};
  const sets = [];
  const params = [];
  const push = (col, val) => { params.push(val); sets.push(`${col}=$${params.length}`); };

  if (u.password   !== undefined) push('password',   String(u.password));
  if (u.fullName   !== undefined) push('full_name',  String(u.fullName));
  if (u.full_name  !== undefined) push('full_name',  String(u.full_name));
  if (u.role       !== undefined) push('role',       String(u.role));
  if (u.department !== undefined) push('department', String(u.department));
  if (u.dept       !== undefined) push('department', String(u.dept));
  if (u.email      !== undefined) push('email',      String(u.email));
  if (u.year       !== undefined) push('year',       String(u.year));
  if (u.status     !== undefined) push('status',     String(u.status));

  if (sets.length === 0) return { status: 'success', message: 'บันทึกสำเร็จ' };
  params.push(String(username).trim());
  await query(`UPDATE users SET ${sets.join(',')} WHERE username=$${params.length}`, params);
  invalidateUsers();
  return { status: 'success', message: 'บันทึกสำเร็จ' };
}

async function deleteUser([username]) {
  await query(`DELETE FROM users WHERE username=$1`, [String(username).trim()]);
  invalidateUsers();
  return { status: 'success', message: 'บันทึกสำเร็จ' };
}

async function importStudentCSV([rows, year]) {
  if (!Array.isArray(rows) || rows.length === 0) return { status: 'success', message: 'นำเข้า 0 รายการ', imported: 0 };
  const { pool } = require('../lib/db');
  const client = await pool.connect();
  let count = 0;
  try {
    await client.query('BEGIN');
    for (const r of rows) {
      if (!r.username) continue;
      await client.query(
        `INSERT INTO users(username, password, full_name, role, department, email, year, status)
         VALUES($1,$2,$3,'Student',$4,$5,$6,$7)
         ON CONFLICT (username) DO UPDATE SET
           full_name=$3, department=$4, email=$5, year=$6, status=$7`,
        [
          String(r.username).trim(),
          String(r.password || r.username).trim(),
          String(r.fullName || r.full_name || '').trim(),
          String(r.department || r.className || '').trim(),
          String(r.email || '').trim(),
          String(r.year || year || '').trim(),
          String(r.status || 'ปกติ'),
        ]
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
  invalidateUsers();
  return { status: 'success', message: `นำเข้าสำเร็จ ${count} รายการ`, imported: count };
}

async function importTeacherCSV([rows]) {
  if (!Array.isArray(rows) || rows.length === 0) return { status: 'success', message: 'นำเข้า 0 รายการ', imported: 0 };
  const { pool } = require('../lib/db');
  const client = await pool.connect();
  let count = 0;
  try {
    await client.query('BEGIN');
    for (const r of rows) {
      if (!r.username) continue;
      await client.query(
        `INSERT INTO users(username, password, full_name, role, department, email, status)
         VALUES($1,$2,$3,$4,$5,$6,$7)
         ON CONFLICT (username) DO UPDATE SET
           full_name=$3, role=$4, department=$5, email=$6, status=$7`,
        [
          String(r.username).trim(),
          String(r.password || r.username).trim(),
          String(r.fullName || r.full_name || '').trim(),
          String(r.role || 'Teacher'),
          String(r.department || r.dept || '').trim(),
          String(r.email || '').trim(),
          String(r.status || 'ปกติ'),
        ]
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
  invalidateUsers();
  return { status: 'success', message: `นำเข้าสำเร็จ ${count} รายการ`, imported: count };
}

async function getStudentSummaryStats([year]) {
  const { rows } = await query(
    `SELECT department, COUNT(*) as cnt
     FROM users WHERE UPPER(role)='STUDENT' AND year=$1 AND status='ปกติ'
     GROUP BY department ORDER BY department`,
    [year]
  );
  const byClass = Object.fromEntries(rows.map(r => [r.department, parseInt(r.cnt)]));
  const total = rows.reduce((s, r) => s + parseInt(r.cnt), 0);
  return { total, byClass };
}

module.exports = {
  addUser,
  editUser,
  deleteUser,
  importStudentCSV,
  importTeacherCSV,
  getStudentSummaryStats,
};
