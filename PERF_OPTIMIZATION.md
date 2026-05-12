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

## Phase 1: Cleanup backup files ✅

**Goal**: ลบไฟล์ backup/unused ลด `clasp push` size + GAS compile time.

### Files removed
| File | Size | Reason |
|---|---|---|
| `src/Scripts_Backup.html` | 170 KB | backup file (CLAUDE.md flagged) |
| `src/Scripts_Score_Backup.html` | 123 KB | backup file |
| `src/Scripts_backup_2.html` | 159 KB | backup file |
| `src/Code_Backup` | 139 KB | backup file |
| `src/Scripts.html` | 18 B | empty `<script></script>` ไม่ใช้ |
| `src/AdminDashboard.html` | 2 KB | unused (replaced by Page_Dashboard_Admin) |
| `src/StudentDashboard.html` | 1.5 KB | unused (replaced by Page_Dashboard_Student) |

**Total removed**: ~595 KB

### Before / After
| Metric | Before | After | Δ |
|---|---|---|---|
| `src/` size | 1.8 MB | 1.2 MB | **-33%** |
| File count | 38 | 31 | -7 |

### Verification
- `grep` Index.html + Scripts_Core.html ไม่มี reference ถึงไฟล์ที่ลบ
- Reference เก่าใน `Scripts_Backup.html` + `Scripts_backup_2.html` เป็น self-reference ภายในไฟล์ backup เอง — ปลอดภัยลบ
- `npx clasp push --force` — sync สำเร็จ
- Commit: f5b7c44 (Phase 0) → next commit

## Phase 2: CacheService ✅

**Goal**: ลด Sheets I/O ที่ซ้ำซ้อนด้วย ScriptCache (shared across all users).

### Files
- New: `src/Cache.js` — `getCached(key, ttl, fetcher)` + `invalidateCache(key)` + `invalidateCacheKeys([])`
- Edited: `src/Code.js`

### Functions cached
| Function | Key | TTL | Reason |
|---|---|---|---|
| `getSystemConfig` | `system_config` | 300s (5 min) | เรียกในทุก request, เปลี่ยนยาก |
| `getAllUsers` | `all_users` | 300s | dashboard + dropdown ใช้บ่อย |
| `getCurriculumData` | `curriculum_all` | 1800s (30 min) | เปลี่ยนเฉพาะ admin import |
| `getCalendarEvents` | `calendar_events` | 600s (10 min) | event-driven invalidation |
| `getAvailableTerms` | `available_terms` | 600s | derived จาก System_Settings |

### Invalidation hooks (write functions)
| Write function | Invalidates |
|---|---|
| `saveSystemConfig` | system_config, available_terms, all_users |
| `addUser` / `editUser` / `deleteUser` | all_users |
| `importStudentCSV` / `importTeacherCSV` | all_users (when news.length > 0) |
| `promoteStudentsToNextYear` | all_users |
| `saveCalendarEvent` (update + create) | calendar_events |
| `deleteCalendarEvent` | calendar_events |
| `importCalendarCSV` | calendar_events |
| `importCurriculumCSV` | curriculum_all |

### Safety
- Cache key prefix `pssms:v1:` — bumpable เมื่อ schema เปลี่ยน
- Value > 95 KB → skip cache (still returns data, logs warn)
- Cache read error → fall back to fetcher (transparent)
- TTL อย่างเดียวก็ค่อนข้างปลอดภัยอยู่แล้ว — ทุก write มี invalidate hook สำคัญด้วย

### Verification
- Open `?debug=1` → call dashboard 2 ครั้งติด → รอบ 2 ต้องเห็น `[CACHE] HIT system_config`
- Backend log (clasp logs) → `[CACHE] MISS ... HIT ...`
- After admin saves new term → next call ต้องเห็น MISS อีกครั้ง

### Expected impact
- `getSystemConfig` ~50-200ms → <5ms เมื่อ HIT
- `getAllUsers` 200-800ms → <10ms เมื่อ HIT (User_Database 1000+ rows)
- `getCurriculumData` ~150ms → <5ms
- Dashboard ที่เรียก 3-5 cached functions → ลดเวลาจาก ~1s → ~50ms (cold) → ~10ms (warm)

- Commit: (pending)

## Phase 3: PropertiesService term/year ✅

**Goal**: sub-millisecond accessor สำหรับ active term/year (ใช้บ่อยที่สุด).

### Helper added (`src/Cache.js`)
- `getActiveTermYear()` — อ่าน term + year จาก `ScriptProperties` ทันที, fallback ไป `getSystemConfig()` ครั้งแรก แล้ว prime properties
- `setActiveTermYear(term, year)` — เซต properties

### Sync hook
- `saveSystemConfig()` → เรียก `setActiveTermYear(term, year)` หลัง invalidateCache

### Caller migrated
- `getStudentsByClass` — เดิมเรียก `getSystemConfig()` แค่เพื่อเอา `.year` → ตอนนี้ใช้ `getActiveTermYear()`

### Why not migrate every caller?
Phase 2 cache ทำให้ `getSystemConfig` HIT ใช้แค่ ~5ms อยู่แล้ว — gain จาก Properties path ~3-4ms ต่อ call. Migrate เฉพาะ hot path ที่เรียกบ่อย (`getStudentsByClass`). ฟังก์ชันที่ต้องการ `termHistory`/`termStart`/`termEnd` ยังคงใช้ `getSystemConfig` (cached).

### Verification
- เปิด `?debug=1` → ดู log ในรอบแรก: `[PROPS] MISS term/year — fallback` → รอบสอง: `[PROPS] HIT term/year`
- Admin save term ใหม่ → `setActiveTermYear` log + ครู refresh ต้องเห็น term ใหม่ทันที (ไม่รอ cache 5 min)

### Expected impact
- `getStudentsByClass` ต่ำลง ~3-4ms ต่อ call (เรียก 5-10 ครั้ง/dashboard = ลด ~30ms)
- Active term refresh ทันที (Properties + Cache invalidate ทำงานพร้อมกัน)

- Commit: (pending)

## Phase 4: Sheets I/O batch ✅

**Goal**: ลบ duplicate Sheet reads ภายใน hot functions.

### Audit findings
- 86 `getDataRange()` calls ใน Code.js — ไม่มี `.getValue()` loop singular (good)
- ปัญหาหลัก: ฟังก์ชันเดียวอ่าน sheet เดียวกัน 2 ครั้ง

### Functions optimized

**1. `getAllInOneScoreGridData`** (Code.js:1966)
- เดิม: อ่าน `Grade_Summary` 2 รอบ (line 1982 + 2032) + `getSystemConfig` ในช่วง historical check
- ตอนนี้: อ่าน Grade_Summary ครั้งเดียว ที่ top → reuse ใน 2 logic blocks
- ใช้ `getActiveTermYear()` แทน `getSystemConfig()` สำหรับ historical check (faster)

**2. `getTeacherRiskDashboard`** (Code.js:2471)
- เดิม: อ่าน `User_Database` 2 รอบ (line 2504 + 2535) — แบบ `year-match` + `fallback`
- ตอนนี้: อ่าน User_Database ครั้งเดียว ที่ top → loop 2 ครั้ง over array เดียว
- ใช้ `getActiveTermYear()` แทน `getSystemConfig()`

### Before / After (theoretical, depends on User_Database size)
| Function | Before | After |
|---|---|---|
| getAllInOneScoreGridData | 2× Grade_Summary read (~100ms each) | 1× read (~100ms) |
| getTeacherRiskDashboard | 2× User_Database read (~80ms each) | 1× read (~80ms) |

Per call savings: ~100ms + ~80ms = **~180ms** for these 2 hot functions.

### debugSheets() logs added at batched reads
- `[SHEETS] read-once Grade_Summary rows=N`
- `[SHEETS] read-once User_Database rows=N`

### Verification
- เปิด `?debug=1` → call Score Entry → ดู log `[SHEETS]` ปรากฏครั้งเดียวต่อ sheet
- Verify output: scores + remarks + students ครบเหมือนเดิม

### Functions reviewed, no change needed
- `saveAttendanceBatch` — already uses `setValues` batch
- `importStudentCSV` / `importTeacherCSV` — already batched setValues
- `getStudentAttendanceHistory` — single read, ok
- `getSemesterReport` — already reads each sheet once

- Commit: (pending)

## Phase 5: Bundle dashboard calls ✅

**Goal**: ลด round-trip ไป-กลับ GAS โดยรวม dashboard calls หลายตัวเป็น call เดียว.

### Why bundle
- ทุก `google.script.run` call: ~500ms-1s overhead (cold start + network round-trip)
- 4 calls แยก = 4× overhead + max(exec)
- 1 bundle = 1× overhead + sum(exec) — ส่วนใหญ่ sub-calls hit cache → ~50ms รวม

### Files
- New: `src/DashboardBundle.js`
- Edited: `src/Scripts_Teacher.html`

### Backend bundle
**`getTeacherDashboardBundle(teacherId, term, year)`** — รวม 4 sections:
- `timetable` ← `getTeacherTimetableWithStatus`
- `riskDashboard` ← `getTeacherRiskDashboard`
- `atRiskDashboard` ← `getTeacherAtRiskDashboard`
- `calendarEvents` ← `getCalendarEvents`

แต่ละ section ห่อ try/catch → ถ้าตัวใดตัวหนึ่ง fail, ตัวอื่นยังกลับ data ได้

**`getAdminDashboardBundle()`** (ready for use) — รวม 5 sections:
- `adminStats`, `studentSummary`, `calendarEvents`, `availableTerms`, `systemConfig`

### Frontend integration (Teacher dashboard)
- `initTeacherDashboard()` เรียก `getTeacherDashboardBundle()` ครั้งเดียว
- Render cached pieces ทันทีจาก sessionStorage ก่อน bundle response กลับมา (fast paint)
- Bundle response → update + cache แต่ละ section
- ToDo list ยังโหลดแยก (per-user, lightweight)

### Fallback strategy
- `loadTeacherDashboardLegacy(user)` — เรียกฟังก์ชันแยกแบบเดิม
- ทำงานถ้า bundle endpoint fail / response null
- ฟังก์ชันแยกเดิม (`loadTeacherRiskDashboard`, `loadTeacherAtRiskDashboard`, `loadDashboardCalendarEvents`) ยังคงอยู่ — backward compat

### Render helpers extracted
- `renderTeacherRiskFromCache(res)` — DRY: bundle path + cache path ใช้ร่วมกัน
- `renderDashboardCalendarEvents(events)` — DRY: รับ events array ที่ extract แล้ว

### Expected impact
- Teacher dashboard: 4 GAS calls (~2-3s) → 1 GAS call (~600ms-1s) — **lazy paint จาก cache แทบ instant**
- Combined with Phase 2 cache: หลังครั้งแรก, sub-calls cache HIT → bundle ~150-300ms

### Verification
- เปิด `?debug=1` ที่ teacher dashboard
- Debug panel ต้องเห็น `[CALL] getTeacherDashboardBundle XXms ok` ครั้งเดียว
- เปิด dashboard อีกครั้ง (cache warm) — bundle response < 500ms
- หยุด network กลางคัน → fallback path ทำงาน

- Commit: (pending)

## Phase 6: Precompute reports ✅

**Goal**: time-trigger คำนวณ heavy dashboards ตอน 02:00 ทุกคืน → frontend ดึงจาก snapshot.

### Files
- New: `src/Precompute.js`
- Edited: `src/Code.js` (เพิ่ม readComputed check ใน 2 functions)

### Design
- **Sheet `_Computed_Cache`** auto-created — columns: `Key | Type | TeacherId | Term | Year | UpdatedAt | PayloadJSON`
- **Key format**: `<type>|<teacherIdLower>|<term>|<year>`
- **Types**: `risk` + `atRisk`

### Functions
- `ensureComputedCacheSheet()` — auto-create sheet + header
- `readComputed(type, teacherId, term, year, maxAgeMs)` — lookup, return null if missing/stale
- `writeComputed(type, teacherId, term, year, payload)` — upsert (overwrite if key exists)
- `_listActiveTeachingContexts()` — distinct (teacher, term, year) tuples from Timetable_Database
- `precomputeNightly()` — main job: loop contexts, compute risk + atRisk, save. Time-budgeted 5 min
- `precomputeNow()` — manual trigger สำหรับ admin/test
- `setupPrecomputeTrigger()` — install daily trigger at 02:00 (run once)

### Hot path integration
`getTeacherRiskDashboard` + `getTeacherAtRiskDashboard` ตอนนี้ check `readComputed()` ก่อน:
- HIT (cache อายุ ≤ 12 ชม.) → return ทันที (~10ms)
- MISS / stale → live compute เหมือนเดิม

### Setup
ใน GAS Editor:
1. รัน `setupPrecomputeTrigger()` ครั้งเดียว → ตั้ง trigger
2. รัน `precomputeNow()` ครั้งแรกเพื่อ prime _Computed_Cache
3. หลังจากนั้นทุกคืน 02:00 → auto refresh

### Expected impact
- Risk dashboard live compute ~3-5s → cache hit ~10ms (**~99% faster**)
- ครู login เช้ามา → dashboard load ใน <500ms (จากที่เคย 3-5s)
- Trade-off: data ตอนเช้า represents เมื่อคืน — เหมาะกับ dashboard summary (ไม่ใช่ realtime)

### Safety
- Trigger time-budgeted 5 นาที (under GAS 6-min limit)
- Live fallback ถ้า cache miss
- ครู save grade ใหม่ → cache 12h ยังคงอยู่ แต่ live compute เมื่อ TTL expires (Phase 2 cache invalidate ครอบ getCalendarEvents/users, ไม่ครอบ _Computed_Cache โดยตรง)

### Verification
- รัน `precomputeNow()` ใน GAS Editor → ดู return summary
- เช็ก `_Computed_Cache` sheet → มี rows
- เปิด dashboard ดู log `[PRECOMP] risk HIT age=Xms`

- Commit: (pending)

## Phase 7: Frontend optimizations
- Status: pending

## Phase 8: Finalize
- Status: pending
