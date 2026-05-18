const { query } = require('../lib/db');
const cache = require('../lib/cache');

function invalidateClubs(term, year) {
  cache.del(`clubs_${term}_${year}`);
}

async function createClub([clubData]) {
  const c = clubData || {};
  const clubId = c.clubId || `club_${Date.now()}`;
  await query(
    `INSERT INTO clubs(club_id,club_name,description,capacity,term,year,status)
     VALUES($1,$2,$3,$4,$5,$6,$7)
     ON CONFLICT(club_id) DO UPDATE SET
       club_name=$2,description=$3,capacity=$4,status=$7,updated_at=NOW()`,
    [clubId, c.clubName || '', c.description || '', c.capacity || 0,
     c.term, c.year, c.status || 'open']
  );

  if (Array.isArray(c.advisors) && c.advisors.length > 0) {
    await query(`DELETE FROM club_advisors WHERE club_id=$1 AND term=$2 AND year=$3`, [clubId, c.term, c.year]);
    for (const a of c.advisors) {
      await query(
        `INSERT INTO club_advisors(club_id,teacher_id,teacher_name,role,term,year)
         VALUES($1,$2,$3,$4,$5,$6)
         ON CONFLICT DO NOTHING`,
        [clubId, a.teacherId, a.teacherName || '', a.role || 'หัวหน้า', c.term, c.year]
      );
    }
  }

  invalidateClubs(c.term, c.year);
  return { status: 'success', message: 'สร้างชุมนุมสำเร็จ', clubId };
}

async function updateClub([clubId, updateData]) {
  const u = updateData || {};
  const sets = [];
  const params = [];
  const push = (col, val) => { params.push(val); sets.push(`${col}=$${params.length}`); };

  if (u.clubName    !== undefined) push('club_name',   u.clubName);
  if (u.description !== undefined) push('description', u.description);
  if (u.capacity    !== undefined) push('capacity',    u.capacity);
  if (u.status      !== undefined) push('status',      u.status);
  push('updated_at', 'NOW()');

  if (sets.length > 0) {
    params.push(clubId);
    await query(
      `UPDATE clubs SET ${sets.join(',')} WHERE club_id=$${params.length}`,
      params
    );
  }

  if (Array.isArray(u.advisors)) {
    await query(`DELETE FROM club_advisors WHERE club_id=$1 AND term=$2 AND year=$3`, [clubId, u.term, u.year]);
    for (const a of u.advisors) {
      await query(
        `INSERT INTO club_advisors(club_id,teacher_id,teacher_name,role,term,year)
         VALUES($1,$2,$3,$4,$5,$6) ON CONFLICT DO NOTHING`,
        [clubId, a.teacherId, a.teacherName || '', a.role || 'หัวหน้า', u.term, u.year]
      );
    }
  }

  if (u.term && u.year) invalidateClubs(u.term, u.year);
  return { status: 'success', message: 'อัปเดตชุมนุมสำเร็จ' };
}

async function registerClub([studentId, studentName, className, clubId, term, year, registeredBy]) {
  const existing = await query(
    `SELECT club_id FROM club_members WHERE student_id=$1 AND term=$2 AND year=$3`,
    [studentId, term, year]
  );
  if (existing.rows.length > 0) {
    throw new Error(`นักเรียน ${studentId} ลงทะเบียนชุมนุมอื่นแล้ว`);
  }
  const clubRes = await query(
    `SELECT capacity FROM clubs WHERE club_id=$1`, [clubId]
  );
  if (clubRes.rows.length === 0) throw new Error('ไม่พบชุมนุม');
  const cap = clubRes.rows[0].capacity;
  if (cap > 0) {
    const cnt = await query(
      `SELECT COUNT(*) as n FROM club_members WHERE club_id=$1 AND term=$2 AND year=$3`,
      [clubId, term, year]
    );
    if (parseInt(cnt.rows[0].n) >= cap) throw new Error('ชุมนุมเต็มแล้ว');
  }
  await query(
    `INSERT INTO club_members(club_id,student_id,student_name,class_name,term,year,registered_by)
     VALUES($1,$2,$3,$4,$5,$6,$7)`,
    [clubId, studentId, studentName || '', className || '', term, year, registeredBy || '']
  );
  invalidateClubs(term, year);
  return { status: 'success', message: 'ลงทะเบียนชุมนุมสำเร็จ' };
}

async function unregisterClub([studentId, term, year]) {
  await query(
    `DELETE FROM club_members WHERE student_id=$1 AND term=$2 AND year=$3`,
    [studentId, term, year]
  );
  invalidateClubs(term, year);
  return { status: 'success', message: 'ยกเลิกลงทะเบียนชุมนุมสำเร็จ' };
}

module.exports = { createClub, updateClub, registerClub, unregisterClub };
