/**
 * PSSMS Precompute (Backend)
 * Nightly job to precompute heavy dashboard aggregates.
 *
 * Targets: getTeacherRiskDashboard and getTeacherAtRiskDashboard — both scan
 * full sheets and aggregate per-teacher. With ~30 teachers × 4-8 subjects each,
 * this is ~3-8 seconds of work per call.
 *
 * Strategy:
 *   1. Nightly trigger (02:00 Asia/Bangkok) reads Timetable_Database to find
 *      every (teacherId, term, year) tuple in use.
 *   2. For each tuple, run getTeacherRiskDashboard + getTeacherAtRiskDashboard.
 *   3. Stash results into "_Computed_Cache" sheet as JSON rows.
 *   4. Frontend hot path can read precomputed row instantly (single getRange
 *      lookup) — falls back to live compute when cache miss / stale.
 *
 * Setup: run setupPrecomputeTrigger() once in GAS editor.
 */

var PSSMS_COMPUTED_SHEET = '_Computed_Cache';
var PSSMS_PRECOMPUTE_TRIGGER_HANDLER = 'precomputeNightly';

/**
 * Ensure the _Computed_Cache sheet exists with the expected header.
 * Headers: Key | Type | TeacherId | Term | Year | UpdatedAt | PayloadJSON
 */
function ensureComputedCacheSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PSSMS_COMPUTED_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PSSMS_COMPUTED_SHEET);
    sheet.appendRow(['Key', 'Type', 'TeacherId', 'Term', 'Year', 'UpdatedAt', 'PayloadJSON']);
    sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#4A86E8').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _computedKey(type, teacherId, term, year) {
  return type + '|' + String(teacherId).trim().toLowerCase() + '|' + String(term).trim() + '|' + String(year).trim();
}

/**
 * Read a precomputed payload. Returns { ok:true, data, ageMs } or null if missing.
 * maxAgeMs: cap freshness; null = no limit.
 */
function readComputed(type, teacherId, term, year, maxAgeMs) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PSSMS_COMPUTED_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getDataRange().getValues();
    var key = _computedKey(type, teacherId, term, year);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        var updatedAt = new Date(data[i][5]).getTime();
        var ageMs = Date.now() - updatedAt;
        if (maxAgeMs && ageMs > maxAgeMs) return null;
        try {
          return { ok: true, data: JSON.parse(data[i][6]), ageMs: ageMs };
        } catch (e) { return null; }
      }
    }
    return null;
  } catch (e) {
    if (typeof debugLog === 'function') debugLog('PRECOMP_ERR', 'read: ' + e);
    return null;
  }
}

/**
 * Upsert one precomputed entry.
 */
function writeComputed(type, teacherId, term, year, payload) {
  var sheet = ensureComputedCacheSheet();
  var key = _computedKey(type, teacherId, term, year);
  var data = sheet.getDataRange().getValues();
  var payloadStr = JSON.stringify(payload);
  var row = [key, type, teacherId, term, year, new Date(), payloadStr];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
}

/**
 * Returns distinct (teacherId, term, year) tuples currently in Timetable_Database.
 */
function _listActiveTeachingContexts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Timetable_Database');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var seen = {};
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var tid = String(data[i][5] || '').trim();
    var term = String(data[i][8] || '').trim();
    var year = String(data[i][9] || '').trim();
    if (!tid || !term || !year) continue;
    var key = tid.toLowerCase() + '|' + term + '|' + year;
    if (seen[key]) continue;
    seen[key] = true;
    out.push({ teacherId: tid, term: term, year: year });
  }
  return out;
}

/**
 * Main nightly entry. Iterates every active teaching context and stashes
 * risk + at-risk payloads. Time-budgeted: aborts after 5 min to stay under
 * GAS 6-min limit.
 */
function precomputeNightly() {
  var start = Date.now();
  var BUDGET_MS = 5 * 60 * 1000;
  ensureComputedCacheSheet();
  var contexts = _listActiveTeachingContexts();
  var done = 0, skipped = 0, failed = 0;

  for (var i = 0; i < contexts.length; i++) {
    if (Date.now() - start > BUDGET_MS) { skipped = contexts.length - i; break; }
    var ctx = contexts[i];
    try {
      var risk = getTeacherRiskDashboard(ctx.teacherId, ctx.term, ctx.year);
      writeComputed('risk', ctx.teacherId, ctx.term, ctx.year, risk);
    } catch (e) { failed++; }
    try {
      var atRisk = getTeacherAtRiskDashboard(ctx.teacherId, ctx.term, ctx.year);
      writeComputed('atRisk', ctx.teacherId, ctx.term, ctx.year, atRisk);
    } catch (e) { failed++; }
    done++;
  }

  var summary = {
    contexts: contexts.length,
    done: done,
    skipped: skipped,
    failed: failed,
    elapsedMs: Date.now() - start
  };
  if (typeof debugLog === 'function') debugLog('PRECOMP', 'done', summary);
  return summary;
}

/**
 * Run-once setup: install daily trigger at 02:00.
 */
function setupPrecomputeTrigger() {
  // Remove any existing triggers for this handler.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === PSSMS_PRECOMPUTE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger(PSSMS_PRECOMPUTE_TRIGGER_HANDLER)
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
  ensureComputedCacheSheet();
  return { ok: true, message: 'Trigger ตั้งเวลา precomputeNightly ที่ 02:00 ทุกคืน' };
}

/**
 * Manual trigger for admin testing.
 */
function precomputeNow() {
  return precomputeNightly();
}
