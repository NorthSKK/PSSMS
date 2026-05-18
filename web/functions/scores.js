const { query } = require('../lib/db');

async function getSubjectConfig([subjectCode, className, term, year]) {
  const { rows } = await query(
    `SELECT subject_id, subject_code, class_name, term, year, score_ratio, indicators_json, teacher_id
     FROM subject_config WHERE subject_code=$1 AND class_name=$2 AND term=$3 AND year=$4`,
    [subjectCode, className, term, year]
  );
  if (rows.length === 0) return null;
  const r = rows[0];
  return {
    subjectId: r.subject_id,
    subjectCode: r.subject_code,
    className: r.class_name,
    term: r.term,
    year: r.year,
    scoreRatio: r.score_ratio || '',
    indicators: r.indicators_json || [],
    teacherId: r.teacher_id || '',
  };
}

async function saveSubjectConfig([configData]) {
  const c = configData || {};
  await query(
    `INSERT INTO subject_config(subject_id,subject_code,class_name,term,year,score_ratio,indicators_json,teacher_id)
     VALUES($1,$2,$3,$4,$5,$6,$7,$8)
     ON CONFLICT(subject_code,class_name,term,year) DO UPDATE SET
       subject_id=$1, score_ratio=$6, indicators_json=$7, teacher_id=$8`,
    [
      c.subjectId || `${c.subjectCode}_${c.className}_${c.term}_${c.year}`,
      c.subjectCode, c.className, c.term, c.year,
      c.scoreRatio || '',
      JSON.stringify(c.indicators || []),
      c.teacherId || '',
    ]
  );
  return { status: 'success', message: 'บันทึกสำเร็จ' };
}

async function getAllInOneScoreGridData([teacherId, subjectCode, className, term, year]) {
  const studentsRes = await require('./students').getStudentsByClass([className, null]);

  const scoresRes = await query(
    `SELECT student_id, indicator_id, score
     FROM score_database WHERE subject_code=$1 AND term=$2 AND year=$3`,
    [subjectCode, term, year]
  );
  const scoreMap = {};
  for (const r of scoresRes.rows) {
    if (!scoreMap[r.student_id]) scoreMap[r.student_id] = {};
    scoreMap[r.student_id][r.indicator_id] = r.score !== null ? parseFloat(r.score) : null;
  }

  const configRes = await query(
    `SELECT indicators_json, score_ratio FROM subject_config
     WHERE subject_code=$1 AND class_name=$2 AND term=$3 AND year=$4`,
    [subjectCode, className, term, year]
  );
  const config = configRes.rows[0] || {};

  const qualRes = await query(
    `SELECT student_id, reading_writing, char_json, comp_json
     FROM qualitative_assess WHERE subject_code=$1 AND term=$2 AND year=$3`,
    [subjectCode, term, year]
  );
  const qualMap = {};
  for (const r of qualRes.rows) {
    qualMap[r.student_id] = {
      readingWriting: r.reading_writing || '',
      charJson: r.char_json || {},
      compJson: r.comp_json || {},
    };
  }

  return {
    students: studentsRes,
    scoreMap,
    indicators: config.indicators_json || [],
    scoreRatio: config.score_ratio || '',
    qualMap,
  };
}

async function saveAllInOneScores([scoreRows, subjectCode, term, year, teacherId]) {
  if (!Array.isArray(scoreRows) || scoreRows.length === 0) return { status: 'success', message: 'บันทึกสำเร็จ' };
  const { pool } = require('../lib/db');
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    for (const row of scoreRows) {
      const { studentId, indicatorId, score } = row;
      const uid = `${studentId}_${subjectCode}_${indicatorId}_${term}_${year}`;
      if (score !== null && score !== undefined && score !== '') {
        await client.query(
          `INSERT INTO score_database(uid,student_id,subject_code,indicator_id,score,term,year)
           VALUES($1,$2,$3,$4,$5,$6,$7)
           ON CONFLICT(student_id,subject_code,indicator_id,term,year) DO UPDATE SET score=$5`,
          [uid, studentId, subjectCode, indicatorId, score, term, year]
        );
        await client.query(
          `INSERT INTO score_history(teacher_id,student_id,subject_code,indicator_id,new_score,term,year)
           VALUES($1,$2,$3,$4,$5,$6,$7)`,
          [teacherId || '', studentId, subjectCode, indicatorId, score, term, year]
        );
      }
    }
    await client.query('COMMIT');
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
  return { success: true, saved: scoreRows.length };
}

async function saveAllInOneWithConfig([configData, scoreRows, qualRows, teacherId]) {
  if (configData) await saveSubjectConfig([configData]);
  if (Array.isArray(scoreRows) && scoreRows.length > 0) {
    await saveAllInOneScores([scoreRows, configData?.subjectCode, configData?.term, configData?.year, teacherId]);
  }
  if (Array.isArray(qualRows) && qualRows.length > 0) {
    const { pool } = require('../lib/db');
    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      for (const r of qualRows) {
        await client.query(
          `INSERT INTO qualitative_assess(student_id,subject_code,term,year,reading_writing,char_json,comp_json)
           VALUES($1,$2,$3,$4,$5,$6,$7)
           ON CONFLICT(student_id,subject_code,term,year) DO UPDATE SET
             reading_writing=$5, char_json=$6, comp_json=$7`,
          [r.studentId, r.subjectCode, r.term, r.year,
           r.readingWriting || '', JSON.stringify(r.charJson || {}), JSON.stringify(r.compJson || {})]
        );
      }
      await client.query('COMMIT');
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    } finally {
      client.release();
    }
  }
  return { status: 'success', message: 'บันทึกสำเร็จ' };
}

module.exports = {
  getSubjectConfig,
  saveSubjectConfig,
  getAllInOneScoreGridData,
  saveAllInOneScores,
  saveAllInOneWithConfig,
};
