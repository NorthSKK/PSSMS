/**
 * PSSMS Clubs (Backend) — ระบบลงทะเบียนชุมนุม
 *
 * Schema: 3 sheets
 *   Club_Database   — ClubID | ClubName | Description | Capacity | Term | Year | Status | CreatedAt | UpdatedAt
 *   Club_Advisors   — ClubID | TeacherID | TeacherName | Role | Term | Year
 *   Club_Members    — ClubID | StudentID | StudentName | ClassName | Term | Year | RegisteredAt | RegisteredBy
 *
 * Permission model:
 *   ADMIN: ทุกอย่าง
 *   TEACHER: เห็นเฉพาะชุมนุมที่ตัวเองเป็น advisor + members
 *   STUDENT: self-register/unregister ของตัวเอง + ดูชุมนุมของตัวเอง
 *
 * Concurrency:
 *   registerToClub ใช้ LockService — atomic capacity check
 */

var CLUB_SHEET = 'Club_Database';
var CLUB_ADVISOR_SHEET = 'Club_Advisors';
var CLUB_MEMBER_SHEET = 'Club_Members';

var CLUB_HEADERS = ['ClubID', 'ClubName', 'Description', 'Capacity', 'Term', 'Year', 'Status', 'CreatedAt', 'UpdatedAt'];
var CLUB_ADVISOR_HEADERS = ['ClubID', 'TeacherID', 'TeacherName', 'Role', 'Term', 'Year'];
var CLUB_MEMBER_HEADERS = ['ClubID', 'StudentID', 'StudentName', 'ClassName', 'Term', 'Year', 'RegisteredAt', 'RegisteredBy'];

// ==========================================
// Setup
// ==========================================
function setupClubDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _ensureSheet(ss, CLUB_SHEET, CLUB_HEADERS);
  _ensureSheet(ss, CLUB_ADVISOR_SHEET, CLUB_ADVISOR_HEADERS);
  _ensureSheet(ss, CLUB_MEMBER_SHEET, CLUB_MEMBER_HEADERS);
  return { ok: true, message: '✅ สร้างฐานข้อมูลชุมนุมเรียบร้อย' };
}

function _ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4A86E8').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _normID(id) { return String(id || '').replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, '') || '0'; }

// ==========================================
// Admin CRUD
// ==========================================

function createClub(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = _ensureSheet(ss, CLUB_SHEET, CLUB_HEADERS);
    var now = new Date();
    var clubId = 'CLUB' + now.getTime();
    var capacity = parseInt(payload.capacity) || 0;
    sheet.appendRow([
      clubId,
      String(payload.clubName || '').trim(),
      String(payload.description || '').trim(),
      capacity,
      String(payload.term).trim(),
      String(payload.year).trim(),
      payload.status || 'open',
      now,
      now
    ]);

    // Save advisors
    if (Array.isArray(payload.advisors) && payload.advisors.length) {
      _writeAdvisors(clubId, payload.term, payload.year, payload.advisors);
    }

    invalidateCacheKeys(['clubs_' + payload.term + '_' + payload.year]);
    return { status: 'success', clubId: clubId, message: 'สร้างชุมนุมเรียบร้อย' };
  } catch (e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateClub(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CLUB_SHEET);
    if (!sheet) return { status: 'error', message: 'ไม่พบชีต Club_Database' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.clubId)) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[
          String(payload.clubName || data[i][1]).trim(),
          String(payload.description || '').trim(),
          parseInt(payload.capacity) || data[i][3],
          data[i][4], // term immutable
          data[i][5], // year immutable
          payload.status || data[i][6],
          data[i][7]  // createdAt
        ]]);
        sheet.getRange(i + 1, 9).setValue(new Date()); // updatedAt

        if (Array.isArray(payload.advisors)) {
          _clearAdvisors(payload.clubId);
          _writeAdvisors(payload.clubId, data[i][4], data[i][5], payload.advisors);
        }

        invalidateCacheKeys(['clubs_' + data[i][4] + '_' + data[i][5]]);
        return { status: 'success', message: 'อัปเดตชุมนุมเรียบร้อย' };
      }
    }
    return { status: 'error', message: 'ไม่พบชุมนุม' };
  } catch (e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteClub(clubId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clubSheet = ss.getSheetByName(CLUB_SHEET);
    if (!clubSheet) return { status: 'error', message: 'ไม่พบชีต Club_Database' };

    var data = clubSheet.getDataRange().getValues();
    var term, year;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(clubId)) {
        term = data[i][4]; year = data[i][5];
        clubSheet.deleteRow(i + 1);
      }
    }
    _clearAdvisors(clubId);
    _clearMembers(clubId);

    if (term && year) invalidateCacheKeys(['clubs_' + term + '_' + year]);
    return { status: 'success', message: 'ลบชุมนุมเรียบร้อย' };
  } catch (e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _writeAdvisors(clubId, term, year, advisors) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = _ensureSheet(ss, CLUB_ADVISOR_SHEET, CLUB_ADVISOR_HEADERS);
  var rows = advisors.map(function(a) {
    return [clubId, String(a.teacherId).trim(), String(a.teacherName || '').trim(), a.role || 'หัวหน้า', term, year];
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, CLUB_ADVISOR_HEADERS.length).setValues(rows);
}

function _clearAdvisors(clubId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLUB_ADVISOR_SHEET);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(clubId)) sheet.deleteRow(i + 1);
  }
}

function _clearMembers(clubId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLUB_MEMBER_SHEET);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(clubId)) sheet.deleteRow(i + 1);
  }
}

// ==========================================
// Read APIs
// ==========================================

/**
 * List clubs for a term/year, decorated with advisor list + member count.
 * Cached 300s; invalidated on any club write.
 */
function getClubList(term, year) {
  var key = 'clubs_' + term + '_' + year;
  return getCached(key, 300, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clubSheet = ss.getSheetByName(CLUB_SHEET);
    if (!clubSheet) return [];

    var clubData = clubSheet.getDataRange().getValues();
    var advisorSheet = ss.getSheetByName(CLUB_ADVISOR_SHEET);
    var memberSheet = ss.getSheetByName(CLUB_MEMBER_SHEET);

    var advisorData = advisorSheet ? advisorSheet.getDataRange().getValues() : [];
    var memberData = memberSheet ? memberSheet.getDataRange().getValues() : [];

    // Index advisors + member count by clubId (1 pass each — Phase 4 batch pattern)
    var advisorByClub = {};
    var memberCountByClub = {};
    for (var i = 1; i < advisorData.length; i++) {
      var aClubId = String(advisorData[i][0]);
      if (!advisorByClub[aClubId]) advisorByClub[aClubId] = [];
      advisorByClub[aClubId].push({
        teacherId: String(advisorData[i][1]),
        teacherName: String(advisorData[i][2]),
        role: String(advisorData[i][3] || 'หัวหน้า')
      });
    }
    for (var j = 1; j < memberData.length; j++) {
      var mClubId = String(memberData[j][0]);
      memberCountByClub[mClubId] = (memberCountByClub[mClubId] || 0) + 1;
    }

    var result = [];
    for (var k = 1; k < clubData.length; k++) {
      var row = clubData[k];
      if (String(row[4]).trim() !== String(term).trim() || String(row[5]).trim() !== String(year).trim()) continue;
      var clubId = String(row[0]);
      var capacity = parseInt(row[3]) || 0;
      var memberCount = memberCountByClub[clubId] || 0;
      result.push({
        clubId: clubId,
        clubName: row[1],
        description: row[2],
        capacity: capacity,
        term: String(row[4]),
        year: String(row[5]),
        status: row[6] || 'open',
        memberCount: memberCount,
        seatsLeft: Math.max(capacity - memberCount, 0),
        full: capacity > 0 && memberCount >= capacity,
        advisors: advisorByClub[clubId] || []
      });
    }
    return result.sort(function(a, b) { return a.clubName.localeCompare(b.clubName, 'th'); });
  });
}

/**
 * Members of a club (admin or assigned advisor only).
 */
function getClubMembers(clubId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLUB_MEMBER_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(clubId)) {
      out.push({
        studentId: String(data[i][1]),
        studentName: String(data[i][2]),
        className: String(data[i][3]),
        term: String(data[i][4]),
        year: String(data[i][5]),
        registeredAt: data[i][6],
        registeredBy: data[i][7]
      });
    }
  }
  return out.sort(function(a, b) { return a.className.localeCompare(b.className) || a.studentId.localeCompare(b.studentId); });
}

// ==========================================
// Student-side
// ==========================================

/**
 * Atomic register — checks capacity, prevents duplicate, prevents 2 clubs/term.
 * registeredBy: 'self' (student) or 'admin'
 */
function registerToClub(studentId, clubId, registeredBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clubSheet = ss.getSheetByName(CLUB_SHEET);
    var memberSheet = _ensureSheet(ss, CLUB_MEMBER_SHEET, CLUB_MEMBER_HEADERS);
    var userSheet = ss.getSheetByName('User_Database');

    if (!clubSheet) return { status: 'error', message: 'ไม่พบฐานข้อมูลชุมนุม' };

    // Find club
    var clubData = clubSheet.getDataRange().getValues();
    var club = null;
    for (var i = 1; i < clubData.length; i++) {
      if (String(clubData[i][0]) === String(clubId)) {
        club = {
          clubId: String(clubData[i][0]),
          clubName: clubData[i][1],
          capacity: parseInt(clubData[i][3]) || 0,
          term: String(clubData[i][4]),
          year: String(clubData[i][5]),
          status: clubData[i][6] || 'open'
        };
        break;
      }
    }
    if (!club) return { status: 'error', message: 'ไม่พบชุมนุมที่เลือก' };
    if (club.status !== 'open' && registeredBy !== 'admin') return { status: 'error', message: 'ชุมนุมปิดรับสมัครแล้ว' };

    // Find student profile
    var studentName = '', className = '';
    if (userSheet) {
      var u = userSheet.getDataRange().getDisplayValues();
      var sid = _normID(studentId);
      for (var j = 1; j < u.length; j++) {
        if (_normID(u[j][0]) === sid) {
          studentName = String(u[j][2]).trim();
          className = String(u[j][4]).trim();
          break;
        }
      }
    }
    if (!studentName) studentName = String(studentId);

    // Check existing membership in this term/year
    var memberData = memberSheet.getDataRange().getValues();
    var memberCount = 0;
    for (var k = 1; k < memberData.length; k++) {
      var rClubId = String(memberData[k][0]);
      var rStdId = _normID(memberData[k][1]);
      var rTerm = String(memberData[k][4]);
      var rYear = String(memberData[k][5]);
      if (rClubId === club.clubId) memberCount++;
      if (rStdId === _normID(studentId) && rTerm === club.term && rYear === club.year) {
        return { status: 'error', message: 'นักเรียนได้ลงทะเบียนชุมนุมอื่นในเทอมนี้แล้ว', existingClubId: rClubId };
      }
    }

    if (club.capacity > 0 && memberCount >= club.capacity && registeredBy !== 'admin') {
      return { status: 'error', message: 'ชุมนุมเต็มแล้ว (' + memberCount + '/' + club.capacity + ')' };
    }

    memberSheet.appendRow([
      club.clubId, "'" + String(studentId), studentName, className,
      club.term, club.year, new Date(), registeredBy || 'self'
    ]);

    invalidateCacheKeys(['clubs_' + club.term + '_' + club.year]);
    return { status: 'success', message: 'ลงทะเบียนชุมนุม "' + club.clubName + '" สำเร็จ', clubName: club.clubName };
  } catch (e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function unregisterFromClub(studentId, clubId, removedBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName(CLUB_MEMBER_SHEET);
    if (!memberSheet) return { status: 'error', message: 'ไม่พบชีตสมาชิก' };

    var data = memberSheet.getDataRange().getValues();
    var sid = _normID(studentId);
    var term, year;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(clubId) && _normID(data[i][1]) === sid) {
        term = data[i][4]; year = data[i][5];
        memberSheet.deleteRow(i + 1);
      }
    }
    if (term && year) invalidateCacheKeys(['clubs_' + term + '_' + year]);
    return { status: 'success', message: 'ยกเลิกการลงทะเบียนเรียบร้อย' };
  } catch (e) {
    return { status: 'error', message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ชุมนุมของนักเรียนคนนี้ ในเทอม/ปีนี้.
 */
function getMyClub(studentId, term, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(CLUB_MEMBER_SHEET);
  if (!memberSheet) return null;
  var data = memberSheet.getDataRange().getValues();
  var sid = _normID(studentId);
  var clubId = null;
  for (var i = 1; i < data.length; i++) {
    if (_normID(data[i][1]) === sid && String(data[i][4]) === String(term) && String(data[i][5]) === String(year)) {
      clubId = String(data[i][0]); break;
    }
  }
  if (!clubId) return null;
  // Look up club details from cached list
  var list = getClubList(term, year);
  for (var j = 0; j < list.length; j++) if (list[j].clubId === clubId) return list[j];
  return null;
}

// ==========================================
// Teacher-side
// ==========================================

/**
 * ชุมนุมที่ครูคนนี้เป็น advisor.
 */
function getMyClubs(teacherId, term, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var advisorSheet = ss.getSheetByName(CLUB_ADVISOR_SHEET);
  if (!advisorSheet) return [];
  var data = advisorSheet.getDataRange().getValues();
  var tid = String(teacherId).trim().toLowerCase();
  var myClubIds = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === tid &&
        String(data[i][4]) === String(term) && String(data[i][5]) === String(year)) {
      myClubIds[String(data[i][0])] = String(data[i][3] || 'หัวหน้า');
    }
  }
  var list = getClubList(term, year);
  return list.filter(function(c) { return myClubIds[c.clubId]; }).map(function(c) {
    c.myRole = myClubIds[c.clubId];
    return c;
  });
}

/**
 * Members of a club, verified that teacher is advisor (Admin bypass).
 */
function getClubMembersForTeacher(teacherId, clubId, term, year, userRole) {
  if (userRole && String(userRole).toUpperCase() === 'ADMIN') return getClubMembers(clubId);
  var mine = getMyClubs(teacherId, term, year);
  for (var i = 0; i < mine.length; i++) if (mine[i].clubId === clubId) return getClubMembers(clubId);
  return { error: 'ไม่มีสิทธิ์เข้าถึงสมาชิกชุมนุมนี้' };
}

// ==========================================
// Admin member management (override)
// ==========================================
function adminAddMember(clubId, studentId) {
  return registerToClub(studentId, clubId, 'admin');
}

function adminRemoveMember(clubId, studentId) {
  return unregisterFromClub(studentId, clubId, 'admin');
}

function adminBulkAssign(clubId, studentIds) {
  var results = { added: 0, skipped: 0, errors: [] };
  if (!Array.isArray(studentIds)) return results;
  for (var i = 0; i < studentIds.length; i++) {
    var res = registerToClub(studentIds[i], clubId, 'admin');
    if (res.status === 'success') results.added++;
    else { results.skipped++; results.errors.push({ studentId: studentIds[i], reason: res.message }); }
  }
  return results;
}

// ==========================================
// Dropdown helpers
// ==========================================
/**
 * Export all clubs + members + advisors for given term/year as a structured
 * payload the frontend can render into print/excel.
 */
function exportClubsForTerm(term, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clubs = getClubList(term, year);
  var memberSheet = ss.getSheetByName(CLUB_MEMBER_SHEET);
  var membersByClub = {};
  if (memberSheet) {
    var data = memberSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4]) !== String(term) || String(data[i][5]) !== String(year)) continue;
      var cid = String(data[i][0]);
      if (!membersByClub[cid]) membersByClub[cid] = [];
      membersByClub[cid].push({
        studentId: String(data[i][1]).replace(/'/g, '').trim(),
        studentName: String(data[i][2]),
        className: String(data[i][3]),
        registeredAt: data[i][6],
        registeredBy: data[i][7]
      });
    }
  }
  clubs.forEach(function(c) { c.members = membersByClub[c.clubId] || []; });
  return {
    term: String(term), year: String(year),
    clubs: clubs,
    generatedAt: new Date()
  };
}

function getTeacherListForClubDropdown() {
  return getCached('teacher_list_dropdown', 300, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('User_Database');
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    var out = [];
    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][3]).toLowerCase();
      if (role === 'teacher' || role === 'ครู' || role === 'admin') {
        out.push({ id: String(data[i][0]).replace(/'/g, '').trim(), name: String(data[i][2]).trim() });
      }
    }
    return out.sort(function(a, b) { return a.name.localeCompare(b.name, 'th'); });
  });
}
