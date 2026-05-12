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

## Phase 0: Baseline Audit ✅

**Goal**: instrument hot path functions วัด baseline ก่อนปรับปรุง.

### Audit findings
- `src/Code.js` มี 84 top-level functions, ~3171 บรรทัด
- ใช้ `getDataRange().getValues()` ทั่วทุก function (ดี — ไม่มี `getValue()` loops singular)
- ปัญหาที่พบ:
  - `getSystemConfig()` ถูกเรียกซ้ำในแทบทุก request → cache target
  - `getStudentsByClass()` อ่าน User_Database + อาจอ่าน User_History_Database → expensive, ใช้บ่อยมาก
  - `getAllInOneScoreGridData()` อ่าน 4 sheets (Subject_Config, User_Database, Grade_Summary, User_Database อีก) → heavy
  - `getTeacherRiskDashboard()` ~200 บรรทัด, อ่าน Timetable + Grade_Summary + User_History
  - `getMassiveAttendanceGrid()` scan Attendance_Database ทุก row → ปัญหาเมื่อ rows เยอะ

### Functions wrapped with withTiming
1. `getStudentsByClass` — Code.js:721
2. `getSystemConfig` — Code.js:177
3. `getAdminStats` — Code.js:346
4. `getTeacherSubjects` — Code.js:1190
5. `getMassiveAttendanceGrid` — Code.js:1367
6. `getAllInOneScoreGridData` — Code.js:1945
7. `getTeacherRiskDashboard` — Code.js:2450
8. `getCalendarEvents` — Code.js:3004

### How to read baseline
1. เปิด debug:
   - Backend: GAS Editor → run `setDebugMode(true)`
   - Frontend: URL `?debug=1`
2. ใช้งานปกติ (login → dashboard → score entry → ...)
3. ดูค่าใน:
   - Overlay debug panel ขวาล่าง — frontend latency
   - `clasp logs` — backend `[PERF]` lines
4. บันทึก timing สำคัญลงตาราง Baseline ด้านล่าง (รอ user เก็บข้อมูลจริง)

### Baseline timings (เก็บภายหลังการใช้งาน)

| Function | Avg ms | Sample size | Notes |
|---|---|---|---|
| getSystemConfig | TBD | — | จะ cache ใน Phase 2-3 |
| getStudentsByClass | TBD | — | target cache ใน Phase 2 |
| getAllInOneScoreGridData | TBD | — | batch read ใน Phase 4 |
| getTeacherRiskDashboard | TBD | — | precompute ใน Phase 6 |
| getMassiveAttendanceGrid | TBD | — | batch + pagination ใน Phase 4 |
| getTeacherSubjects | TBD | — | cache ใน Phase 2 |
| getCalendarEvents | TBD | — | cache 5 นาที ใน Phase 2 |
| getAdminStats | TBD | — | bundle ใน Phase 5 |

- Commit: (pending)

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
