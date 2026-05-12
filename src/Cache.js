/**
 * PSSMS Cache Helper (Backend)
 * Wraps Apps Script CacheService with logging + safe JSON handling.
 *
 * Usage:
 *   return getCached('users_all', 600, function() {
 *     return getAllUsersFromSheet();
 *   });
 *
 * Invalidate:
 *   invalidateCache('users_all');
 *   invalidateCacheByPrefix('students_');
 *   invalidateAllCache();
 *
 * Limits (Apps Script CacheService):
 *   - Value size: 100 KB per key
 *   - TTL max: 21600 sec (6 hours)
 *   - Keys per put: 1000 max
 *
 * Behavior:
 *   - If serialized value > 95 KB → skips cache (logs warning), still returns fetched data
 *   - Cache HIT/MISS logged via debugCache() (Phase -1 toolkit)
 *   - Errors during cache read fall back to fetcher (cache transparent)
 */

var PSSMS_CACHE_PREFIX = 'pssms:v1:';
var PSSMS_CACHE_MAX_BYTES = 95 * 1024; // 95 KB safety margin under 100 KB hard limit
var PSSMS_CACHE_DEFAULT_TTL = 600; // 10 min default

/**
 * Get a value from cache, or compute + store it if missing.
 * @param {string} key short logical key (prefix auto-added)
 * @param {number} ttlSec time-to-live in seconds
 * @param {function} fetcher zero-arg function that computes the value
 * @return {*} cached or freshly computed value
 */
function getCached(key, ttlSec, fetcher) {
  var fullKey = PSSMS_CACHE_PREFIX + key;
  var cache;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }

  if (cache) {
    try {
      var raw = cache.get(fullKey);
      if (raw !== null && raw !== undefined) {
        if (typeof debugCache === 'function') debugCache(key, true);
        return JSON.parse(raw);
      }
    } catch (e) {
      if (typeof debugLog === 'function') debugLog('CACHE_ERR', 'read ' + key + ': ' + e);
    }
  }

  if (typeof debugCache === 'function') debugCache(key, false);
  var value = fetcher();

  if (cache && value !== undefined) {
    try {
      var serialized = JSON.stringify(value);
      if (serialized.length <= PSSMS_CACHE_MAX_BYTES) {
        cache.put(fullKey, serialized, ttlSec || PSSMS_CACHE_DEFAULT_TTL);
      } else if (typeof debugLog === 'function') {
        debugLog('CACHE_SKIP', 'value too large for ' + key + ' (' + serialized.length + ' bytes)');
      }
    } catch (e) {
      if (typeof debugLog === 'function') debugLog('CACHE_ERR', 'write ' + key + ': ' + e);
    }
  }

  return value;
}

/**
 * Remove a single key from cache.
 */
function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(PSSMS_CACHE_PREFIX + key);
    if (typeof debugLog === 'function') debugLog('CACHE_INVAL', key);
  } catch (e) {}
}

/**
 * Remove multiple keys at once.
 */
function invalidateCacheKeys(keys) {
  if (!keys || !keys.length) return;
  try {
    var prefixed = keys.map(function(k) { return PSSMS_CACHE_PREFIX + k; });
    CacheService.getScriptCache().removeAll(prefixed);
    if (typeof debugLog === 'function') debugLog('CACHE_INVAL', 'batch:' + keys.length);
  } catch (e) {}
}

/**
 * Nuclear option — only when schema/data changed broadly.
 * CacheService has no listAllKeys API, so we track known keys.
 */
var PSSMS_TRACKED_KEYS = [
  'system_config',
  'all_users',
  'teacher_list_dropdown',
  'curriculum_all',
  'calendar_events',
  'available_terms'
];

function invalidateAllCache() {
  invalidateCacheKeys(PSSMS_TRACKED_KEYS);
  if (typeof debugLog === 'function') debugLog('CACHE_INVAL', 'ALL');
}

/**
 * Returns cache hit/miss stats since last reset.
 * For ad-hoc verification — call from debug action / admin tool.
 */
function getCacheStatus() {
  // CacheService has no native stats API; return last known config only
  return {
    prefix: PSSMS_CACHE_PREFIX,
    maxBytes: PSSMS_CACHE_MAX_BYTES,
    trackedKeys: PSSMS_TRACKED_KEYS
  };
}
