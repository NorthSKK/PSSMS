# PSSMS Performance Optimization Log

> Tracking ทุก phase ของการเพิ่ม performance ระบบ PSSMS
> เริ่ม: 2026-05-12

---

## Debug Toolkit (Phase -1) ✅

**Goal**: instrumentation infrastructure เพื่อวัดผลทุก phase ถัดไป.

### Backend (`src/Debug.js`)
- `isDebugMode()` / `setDebugMode(on)` — toggle ผ่าน `PropertiesService` (`DEBUG_MODE` key)
- `debugLog(tag, msg, data)` — log ไป `Logger.log` + `console.log` เมื่อ DEBUG เปิดเท่านั้น
- `withTiming(name, fn)` — wrap function วัด ms (no-op เมื่อปิด debug)
- `debugCache(key, hit)` — log cache HIT/MISS
- `debugSheets(sheet, op, rows)` — log Sheets I/O ops
- `debugPing()` — frontend ใช้ ping ดู latency + debug state

### Frontend (`src/Scripts_Core.html`)
- `window.PSSMS_DEBUG` — เปิดด้วย `?debug=1` ที่ URL หรือ `localStorage.pssms_debug=1`
- `window.pssmsDebugLog(tag, msg, data)` — log + render บน overlay panel
- `window.pssmsDebugStats` — `{ calls, totalMs, byFn }` สะสมตลอด session
- Overlay debug panel มุมขวาล่าง:
  - แสดง tag + message + duration ของทุก GAS call
  - มี `clear` + `off` buttons
  - Summary header: `calls:N | XXXms`
- Patched `safeRun()` Proxy — ทุก GAS call ที่ผ่าน `safeRun` ถูกวัดเวลาอัตโนมัติ
  - Log `[CALL] fnName start` → `[CALL] fnName XXms ok` หรือ `[ERR] fnName XXms fail`

### How to enable
```bash
# Backend
# ใน GAS editor → Run setDebugMode(true)
# หรือเรียกผ่าน admin tool ทีหลัง

# Frontend
# เปิด URL ลงท้ายด้วย ?debug=1
# หรือ localStorage.setItem('pssms_debug','1') แล้ว reload
```

### Status
- ✅ Files created: `src/Debug.js`, edited `src/Scripts_Core.html`
- ⏳ Pending: push + verify ใน production
- Commit: (รอ commit)

---

## Phase 0: Baseline Audit
- Status: pending

## Phase 1: Cleanup backup files
- Status: pending

## Phase 2: CacheService
- Status: pending

## Phase 3: PropertiesService term/year
- Status: pending

## Phase 4: Sheets I/O batch
- Status: pending

## Phase 5: Bundle dashboard calls
- Status: pending

## Phase 6: Precompute reports
- Status: pending

## Phase 7: Frontend optimizations
- Status: pending

## Phase 8: Finalize
- Status: pending
