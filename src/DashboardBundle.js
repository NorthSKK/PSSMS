/**
 * PSSMS Dashboard Bundle (Backend)
 * Batches multiple per-dashboard reads into a single google.script.run call.
 *
 * Why: each google.script.run call pays ~500ms-1s round-trip overhead.
 * 4 parallel calls = max(exec) + 4× overhead. One bundled call = sum(exec) + 1× overhead.
 * Net win because most sub-calls are cached (Phase 2) and execute in <50ms.
 *
 * Each section is wrapped in try/catch so a single failure doesn't tank the whole
 * dashboard — frontend can show partial data.
 */

function _bundleSection(name, fn) {
  try {
    var result = fn();
    return { ok: true, data: result };
  } catch (e) {
    if (typeof debugLog === 'function') debugLog('BUNDLE_ERR', name + ': ' + e);
    return { ok: false, error: String(e && e.message || e) };
  }
}

/**
 * Bundle for teacher dashboard.
 * Combines: timetable, risk grades, at-risk attendance, calendar events.
 * Skips todoList (per-user, lightweight, has its own caching).
 */
function getTeacherDashboardBundle(teacherId, term, year) {
  return withTiming('getTeacherDashboardBundle', function() {
    return {
      ts: Date.now(),
      timetable: _bundleSection('timetable', function() {
        return getTeacherTimetableWithStatus(teacherId);
      }),
      riskDashboard: _bundleSection('riskDashboard', function() {
        return getTeacherRiskDashboard(teacherId, term, year);
      }),
      atRiskDashboard: _bundleSection('atRiskDashboard', function() {
        return getTeacherAtRiskDashboard(teacherId, term, year);
      }),
      calendarEvents: _bundleSection('calendarEvents', function() {
        return getCalendarEvents();
      })
    };
  });
}

/**
 * Bundle for executive dashboard.
 * dept: "ผอ." | "วิชาการ" | "งบประมาณ" | "บุคคล" | "ทั่วไป"
 * All sections always fetched — frontend filters display by dept.
 */
function getExecutiveDashboardBundle(dept) {
  return withTiming('getExecutiveDashboardBundle', function() {
    var config = getSystemConfig();
    return {
      ts: Date.now(),
      dept: dept || 'ผอ.',
      kpi:        _bundleSection('kpi',        function() { return _execKPI(config); }),
      academic:   _bundleSection('academic',   function() { return _execAcademic(config); }),
      budget:     _bundleSection('budget',     function() { return _execBudget(config); }),
      personnel:  _bundleSection('personnel',  function() { return _execPersonnel(config); }),
      general:    _bundleSection('general',    function() { return _execGeneral(config); }),
      calendar:   _bundleSection('calendar',   function() { return getCalendarEvents(); })
    };
  });
}

function _execKPI(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('User_Database');
  var attSheet  = ss.getSheetByName('Attendance_Database');
  var budSheet  = ss.getSheetByName('Budgets');
  var year = String(config.year);
  var term = String(config.term);
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

  var studentCount = 0, teacherCount = 0;
  if (userSheet && userSheet.getLastRow() > 1) {
    var ud = userSheet.getDataRange().getValues().slice(1);
    ud.forEach(function(r) {
      var role = String(r[3]).toUpperCase();
      if (role === 'STUDENT' && String(r[6]) === year) studentCount++;
      if (role === 'TEACHER') teacherCount++;
    });
  }

  var todayPresent = 0, todayTotal = 0;
  if (attSheet && attSheet.getLastRow() > 1) {
    var att = attSheet.getDataRange().getValues().slice(1);
    var seen = {};
    att.forEach(function(r) {
      var d = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[1]);
      if (d !== today || String(r[2]) !== term || String(r[3]) !== year) return;
      var sid = String(r[8]).replace(/^'/, '').trim();
      if (seen[sid]) return;
      seen[sid] = true;
      todayTotal++;
      if (String(r[10]) === 'มา' || String(r[10]) === 'present') todayPresent++;
    });
  }
  var attPct = todayTotal > 0 ? Math.round(todayPresent / todayTotal * 100) : null;

  var budgetUsedPct = 0, budgetTotal = 0, budgetUsed = 0;
  if (budSheet && budSheet.getLastRow() > 1) {
    budSheet.getDataRange().getValues().slice(1).forEach(function(r) {
      if (String(r[6]) === year) { budgetTotal += Number(r[2]); budgetUsed += Number(r[3]); }
    });
    budgetUsedPct = budgetTotal > 0 ? Math.round(budgetUsed / budgetTotal * 100) : 0;
  }

  return { studentCount: studentCount, teacherCount: teacherCount, attPct: attPct, todayPresent: todayPresent, todayTotal: todayTotal, budgetUsedPct: budgetUsedPct, budgetTotal: budgetTotal, budgetUsed: budgetUsed };
}

function _execAcademic(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var term = String(config.term), year = String(config.year);
  var result = { trend: [], riskCount: 0, noAttendanceTeachers: [] };

  // Attendance trend: last 14 days
  var attSheet = ss.getSheetByName('Attendance_Database');
  if (attSheet && attSheet.getLastRow() > 1) {
    var now = new Date();
    var days = {};
    for (var d = 13; d >= 0; d--) {
      var dt = new Date(now.getTime() - d * 86400000);
      var ds = Utilities.formatDate(dt, 'Asia/Bangkok', 'yyyy-MM-dd');
      days[ds] = { date: ds, present: 0, absent: 0, leave: 0, total: 0 };
    }
    var att = attSheet.getDataRange().getValues().slice(1);
    att.forEach(function(r) {
      var ds = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[1]);
      if (!days[ds] || String(r[2]) !== term || String(r[3]) !== year) return;
      days[ds].total++;
      var s = String(r[10]);
      if (s === 'มา' || s === 'present') days[ds].present++;
      else if (s === 'ลา' || s === 'leave') days[ds].leave++;
      else days[ds].absent++;
    });
    result.trend = Object.values(days);
  }

  // Risk from _Computed_Cache
  var cacheSheet = ss.getSheetByName('_Computed_Cache');
  if (cacheSheet) {
    var cData = cacheSheet.getDataRange().getValues();
    for (var i = 1; i < cData.length; i++) {
      if (String(cData[i][0]) === 'atRisk' && String(cData[i][1]) === term + '_' + year) {
        try { var rd = JSON.parse(String(cData[i][2])); result.riskCount = (rd.critical||[]).length + (rd.ms||[]).length + (rd.risk||[]).length; } catch(e) {}
        break;
      }
    }
  }

  // Teachers not logged attendance last 3 days
  var arSheet = ss.getSheetByName('Academic_Records');
  if (arSheet && arSheet.getLastRow() > 1) {
    var arData = arSheet.getDataRange().getValues().slice(1);
    var lastLog = {};
    arData.forEach(function(r) {
      if (String(r[1]) !== term || String(r[2]) !== year) return;
      var tid = String(r[10]).trim();
      var ds = r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[0]);
      if (!lastLog[tid] || ds > lastLog[tid]) lastLog[tid] = ds;
    });
    var cutoff = Utilities.formatDate(new Date(Date.now() - 3*86400000), 'Asia/Bangkok', 'yyyy-MM-dd');
    var ud = ss.getSheetByName('User_Database');
    if (ud) {
      ud.getDataRange().getValues().slice(1).forEach(function(r) {
        if (String(r[3]).toUpperCase() !== 'TEACHER') return;
        var tid = String(r[0]).trim();
        if (!lastLog[tid] || lastLog[tid] < cutoff) result.noAttendanceTeachers.push(String(r[2]));
      });
    }
  }
  return result;
}

function _execBudget(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Budgets');
  if (!sheet || sheet.getLastRow() < 2) return { projects: [], total: 0, used: 0 };
  var year = String(config.year);
  var projects = [];
  sheet.getDataRange().getValues().slice(1).forEach(function(r) {
    if (String(r[6]) !== year) return;
    var budget = Number(r[2]), used = Number(r[3]);
    projects.push({ id: String(r[0]), name: String(r[1]), budget: budget, used: used, balance: Number(r[4]), status: String(r[5]), pct: budget > 0 ? Math.round(used/budget*100) : 0 });
  });
  projects.sort(function(a,b) { return b.pct - a.pct; });
  var total = projects.reduce(function(s,p) { return s+p.budget; }, 0);
  var used  = projects.reduce(function(s,p) { return s+p.used; }, 0);
  return { projects: projects, total: total, used: used };
}

function _execPersonnel(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = ss.getSheetByName('Leave_Records');
  var result = { thisMonthLeave: [], leaveByType: {}, staffCount: 0 };
  var year = String(config.year);

  var ud = ss.getSheetByName('User_Database');
  if (ud) result.staffCount = ud.getDataRange().getValues().slice(1).filter(function(r) { return String(r[3]).toUpperCase() === 'TEACHER'; }).length;

  if (leaveSheet && leaveSheet.getLastRow() > 1) {
    var now = new Date();
    var monthStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
    leaveSheet.getDataRange().getValues().slice(1).forEach(function(r) {
      if (String(r[6]) !== year) return;
      var start = r[2] instanceof Date ? Utilities.formatDate(r[2], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[2]);
      var type = String(r[1]);
      result.leaveByType[type] = (result.leaveByType[type] || 0) + 1;
      if (start.startsWith(monthStr)) result.thisMonthLeave.push({ name: String(r[0]), type: type, start: start, end: r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Bangkok', 'yyyy-MM-dd') : String(r[3]) });
    });
  }
  return result;
}

function _execGeneral(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sarabun_Database');
  if (!sheet || sheet.getLastRow() < 2) return { recent: [], pendingFile: 0 };
  var data = sheet.getDataRange().getDisplayValues().slice(1);
  var recent = [], pendingFile = 0;
  for (var i = data.length - 1; i >= 0 && recent.length < 5; i--) {
    var r = data[i];
    if (!String(r[2]).trim() && !String(r[1]).trim()) continue;
    if (!String(r[14]).trim()) pendingFile++;
    if (recent.length < 5) recent.push({ docNumber: String(r[2]), docType: String(r[1]), subject: String(r[3]), date: String(r[4]) });
  }
  // count all pending (not just top 5)
  data.forEach(function(r) { if (!String(r[2]).trim()) return; if (!String(r[14]).trim()) pendingFile = pendingFile; });
  return { recent: recent, pendingFile: pendingFile };
}

/**
 * Bundle for admin dashboard.
 * Combines: adminStats, studentSummary, calendarEvents, availableTerms.
 */
function getAdminDashboardBundle() {
  return withTiming('getAdminDashboardBundle', function() {
    return {
      ts: Date.now(),
      adminStats: _bundleSection('adminStats', function() { return getAdminStats(); }),
      studentSummary: _bundleSection('studentSummary', function() { return getStudentSummaryStats(); }),
      calendarEvents: _bundleSection('calendarEvents', function() { return getCalendarEvents(); }),
      availableTerms: _bundleSection('availableTerms', function() { return getAvailableTerms(); }),
      systemConfig: _bundleSection('systemConfig', function() { return getSystemConfig(); })
    };
  });
}
