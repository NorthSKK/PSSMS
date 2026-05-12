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
