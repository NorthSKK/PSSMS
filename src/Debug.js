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
