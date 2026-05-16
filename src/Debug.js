/**
 * PSSMS Debug Toolkit (Backend)
 * Performance instrumentation + structured logging.
 *
 * Usage:
 *   - Toggle: setDebugMode(true)  // หรือ false
 *   - Wrap:   return withTiming('myFunc', () => { ... });
 *   - Log:    debugLog('TAG', 'message', { extra: 1 });
 *   - Cache:  debugCache(key, isHit);
 *   - Sheets: debugSheets(sheetName, op, rowCount);
 *
 * All logs go to both Logger.log (clasp logs) and console.log (Stackdriver).
 * When DEBUG_MODE is false, withTiming still runs but skips logging overhead.
 */

var PSSMS_DEBUG_PROP_KEY = 'DEBUG_MODE';

function isDebugMode() {
  try {
    return PropertiesService.getScriptProperties().getProperty(PSSMS_DEBUG_PROP_KEY) === 'true';
  } catch (e) {
    return false;
  }
}

function setDebugMode(on) {
  PropertiesService.getScriptProperties()
    .setProperty(PSSMS_DEBUG_PROP_KEY, on ? 'true' : 'false');
  return { ok: true, debugMode: on === true };
}

function getDebugMode() {
  return { debugMode: isDebugMode() };
}

function debugLog(tag, msg, data) {
  if (!isDebugMode()) return;
  var line = '[' + tag + '] ' + msg;
  if (data !== undefined) {
    try { line += ' ' + JSON.stringify(data); } catch (e) { line += ' [unserializable]'; }
  }
  try { Logger.log(line); } catch (e) {}
  try { console.log(line); } catch (e) {}
}

function withTiming(name, fn) {
  if (!isDebugMode()) return fn();
  var start = Date.now();
  try {
    var result = fn();
    debugLog('PERF', name + ' ' + (Date.now() - start) + 'ms');
    return result;
  } catch (e) {
    debugLog('PERF_ERR', name + ' failed after ' + (Date.now() - start) + 'ms: ' + e);
    throw e;
  }
}

function debugCache(key, hit) {
  debugLog('CACHE', (hit ? 'HIT  ' : 'MISS ') + key);
}

function debugSheets(sheet, op, rows) {
  debugLog('SHEETS', op + ' ' + sheet + ' rows=' + rows);
}

/**
 * เรียกจาก frontend เพื่อ ping latency + ดู debug state.
 */
function debugPing() {
  return {
    ok: true,
    ts: Date.now(),
    debugMode: isDebugMode(),
    tz: Session.getScriptTimeZone()
  };
}

// One-click wrappers for GAS Editor "Run" button (no args needed)
function _enableDebug() { return setDebugMode(true); }
function _disableDebug() { return setDebugMode(false); }

/**
 * Export timetable rows for a given term/year grouped by TeacherID.
 * Run: debugExportTimetable('1','2569')
 * Output goes to Logger — view with: npx clasp logs
 */
function debugExportTimetable(term, year) {
  term = term || '1'; year = year || '2569';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Timetable_Database');
  if (!sheet) return { error: 'no Timetable_Database' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; // SubjectCode,SubjectName,Level,Room,Location,TeacherID,Day,Period,Term,Year

  // Collect rows for this term/year
  var byTeacher = {};
  var total = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var t = String(r[8]).trim(), y = String(r[9]).trim();
    if (t !== String(term) || y !== String(year)) continue;
    var tid = String(r[5]).trim();
    if (!byTeacher[tid]) byTeacher[tid] = [];
    byTeacher[tid].push({
      code: String(r[0]).trim(),
      name: String(r[1]).trim(),
      level: String(r[2]).trim(),
      room: String(r[3]).trim(),
      loc:  String(r[4]).trim(),
      day:  String(r[6]).trim(),
      period: String(r[7]).trim()
    });
    total++;
  }

  // Get teacher names from User_Database
  var uSheet = ss.getSheetByName('User_Database');
  var nameMap = {};
  if (uSheet) {
    var ud = uSheet.getDataRange().getValues();
    for (var j = 1; j < ud.length; j++) {
      nameMap[String(ud[j][0]).trim()] = String(ud[j][2]).trim();
    }
  }

  var DAY_ORDER = {จันทร์:1,อังคาร:2,พุธ:3,พฤหัสบดี:4,ศุกร์:5};
  var lines = ['=== Timetable Export: เทอม ' + term + '/' + year + ' | ' + total + ' rows | ' + Object.keys(byTeacher).length + ' teachers ===\n'];

  Object.keys(byTeacher).sort().forEach(function(tid) {
    var slots = byTeacher[tid];
    slots.sort(function(a,b){
      var dd = (DAY_ORDER[a.day]||9) - (DAY_ORDER[b.day]||9);
      return dd !== 0 ? dd : Number(a.period) - Number(b.period);
    });
    lines.push('--- ' + tid + ' (' + (nameMap[tid]||'?') + ') | ' + slots.length + ' คาบ ---');
    slots.forEach(function(s) {
      lines.push('  ' + s.day + ' คาบ' + s.period + ' | ' + s.code + ' | ' + s.name + ' | ' + s.level + '/' + s.room + (s.loc ? ' [' + s.loc + ']' : ''));
    });
    lines.push('');
  });

  var output = lines.join('\n');
  Logger.log(output);
  return { term: term, year: year, total: total, teachers: Object.keys(byTeacher).length, output: output };
}

function _debugExportTimetable1_2569() { return debugExportTimetable('1','2569'); }
