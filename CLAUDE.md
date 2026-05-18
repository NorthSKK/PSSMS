# PSSMS — Phuphrabat Smart School Management System

ระบบบริหารจัดการสถานศึกษา 4 ฝ่าย สำหรับโรงเรียนภูพระบาทวิทยา  
พัฒนาโดย: ครูน๊อต ศิกษก เดินรีบรัมย์

---

## Platform & Deploy

- **Runtime**: Google Apps Script (GAS) V8, Timezone: Asia/Bangkok
- **Script ID**: `1fOZg9_N5LrsOMHPozhgf09SOW_x-2os7biXZ_-De4DIgsyvig2bDktrW`
- **Deploy as**: Web App — `executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS`
- **Tool**: [clasp](https://github.com/google/clasp) (Local → GAS)

### คำสั่ง Deploy

```bash
# push โค้ดขึ้น GAS
npx clasp push

# ดู log ใน GAS
npx clasp logs

# เปิด GAS Editor
npx clasp open
```

> **หมายเหตุ:** `clasp push` จะดูด **ทุกไฟล์ใน `src/`** รวมถึงไฟล์ backup (`Scripts_Backup.html`, `Scripts_Score_Backup.html`, `Scripts_backup_2.html`, `Code_Backup/`) — ระวังไฟล์เหล่านี้เพิ่ม execution time และ quota

---

## Architecture

```
Local (src/)  →  clasp push  →  Google Apps Script
                                      │
                        ┌─────────────┼──────────────┐
                   Code.js         Sheets DB      Drive
                 (Backend)     (ข้อมูลทั้งหมด)  (ไฟล์แนบ)
                        └─────────────┼──────────────┘
                                      │
                              Index.html (SPA)
                           (Bootstrap 5.3 + JS)
```

- **Backend**: `Code.js` (~3465 บรรทัด) — GAS server-side, ฟังก์ชันทั้งหมดเรียกผ่าน `google.script.run`
- **Database**: Google Sheets (`SpreadsheetApp.getActiveSpreadsheet()`)
- **Frontend**: SPA ใน `Index.html` — routing ด้วย `loadPage()` + `google.script.run.getPage()`
- **File Storage**: Google Drive (`DriveApp`)

---

## โครงสร้างไฟล์ (`src/`)

### Backend
| ไฟล์ | หน้าที่ |
|---|---|
| `Code.js` | Server-side ทั้งหมด |
| `Debug.js` | Debug toolkit (timing, cache logs, DEBUG_MODE toggle) |
| `Cache.js` | `getCached(key, ttl, fetcher)` + `getActiveTermYear()` fast-path |
| `DashboardBundle.js` | `getTeacherDashboardBundle()` / `getAdminDashboardBundle()` / `getExecutiveDashboardBundle(dept)` |
| `Precompute.js` | Nightly trigger → `_Computed_Cache` sheet (risk + atRisk) |
| `Clubs.js` | ระบบลงทะเบียนชุมนุม (CRUD + atomic register + permission) |
| `appsscript.json` | GAS manifest (timezone, webapp config) |

### Frontend — Shell & Shared
| ไฟล์ | หน้าที่ |
|---|---|
| `Index.html` | SPA shell: sidebar nav, navbar, `<div id="page-content">` |
| `Login.html` | หน้า login — มี saved-accounts autofill dropdown (จาก `pssms_saved_accounts`) |
| `Styles.html` | CSS ทั้งหมด (include ใน Index.html) |

### Frontend — Scripts (แยกตาม role/feature)
| ไฟล์ | หน้าที่ |
|---|---|
| `Scripts_Core.html` | Auth, routing (`loadPage`), UI core, `syncSystemTerm()` |
| `Scripts_Admin.html` | ฟังก์ชัน Admin |
| `Scripts_Teacher.html` | ฟังก์ชันครู (เช็คชื่อ, บันทึกการสอน) |
| `Scripts_Academic.html` | วิชาการ (ตารางสอน, กลุ่มเสี่ยง) |
| `Scripts_Score.html` | บันทึกคะแนน All-in-One |
| `Scripts_General.html` | สารบรรณ, ทั่วไป |
| `Scripts_Calendar.html` | ปฏิทิน FullCalendar 6 |

### Frontend — Pages
| ไฟล์ | ชื่อใช้ใน `loadPage()` | หน้าที่ |
|---|---|---|
| `Page_Dashboard_Admin.html.html` | `Page_Dashboard_Admin` | Dashboard Admin |
| `Page_Dashboard_Teacher.html.html` | `Page_Dashboard_Teacher` | Dashboard ครู — `#dashCalendarStrip` แสดง skeleton pills ขณะโหลด, แสดงกิจกรรม 14 วันข้างหน้า, ซ่อนถ้าไม่มี event |
| `Page_Dashboard_Student.html.html` | `Page_Dashboard_Student` | Dashboard นักเรียน |
| `Page_Admin_Users.html` | `Page_Admin_Users` | จัดการผู้ใช้ — 3 tabs: ทะเบียนนักเรียน / บุคลากร (ครู+แอดมิน) / ครูที่ปรึกษาประจำชั้น |
| `Page_Admin_Timetable.html` | `Page_Admin_Timetable` | จัดการตารางสอน |
| `Page_Admin_Settings.html` | `Page_Admin_Settings` | ตั้งค่าระบบ |
| `Page_Admin_Curriculum.html` | `Page_Admin_Curriculum` | หลักสูตร |
| `Page_Academic.html` | `Page_Academic` | เช็คชื่อ / กิจกรรมหน้าเสาธง |
| `Page_Academic_Report.html.html` | `Page_Academic_Report` | รายงานสถิติ (มส.) |
| `Page_Score_Entry.html` | `Page_Score_Entry` | บันทึกคะแนน (ปพ.5) |
| `Page_Subject_Config.html` | `Page_Subject_Config` | ตั้งค่าโครงสร้างวิชา |
| `Page_Grade_Summary.html` | `Page_Grade_Summary` | ปพ.5 สมุดบันทึกผลการเรียน |
| `Page_Calendar.html` | `Page_Calendar` | ปฏิทินปฏิบัติงาน |
| `Page_General.html` | `Page_General` | สารบรรณ |
| `Page_Budget.html` | `Page_Budget` | งบประมาณ |
| `Page_Personnel.html` | `Page_Personnel` | บุคลากร |
| `Page_Lesson_History.html.html` | `Page_Lesson_History` | แฟ้มบันทึกหลังสอน |
| `Page_Homeroom.html` | `Page_Homeroom` | ครูที่ปรึกษาประจำชั้น (standalone — ไม่มีลิงก์ใน sidebar แล้ว, เข้าผ่าน tab ใน Page_Admin_Users แทน) |
| `Page_Dashboard_Executive.html.html` | `Page_Dashboard_Executive` | Dashboard ผู้บริหาร — KPI strip, alerts, dept-scoped sections, calendar (ดู EXECUTIVE role) |
| `Template_PP5.html` | — | Template พิมพ์ ปพ.5 |

> ไฟล์ที่มีนามสกุล `.html.html` (เช่น `Page_Dashboard_Admin.html.html`) — GAS จะตัดนามสกุลออกชั้นหนึ่ง ชื่อที่ใช้ใน `loadPage()` จึงไม่มี `.html`

---

## Google Sheets — ชีตทั้งหมด

### ชีตที่สร้างโดย `setupDatabase()`
| Sheet | Headers หลัก | ใช้งาน |
|---|---|---|
| `User_Database` | [0]Username, [1]Password, [2]FullName, [3]Role, [4]Department, [5]Email, [6]Year | ผู้ใช้ทั้งหมด |
| `Attendance_Database` | Timestamp, Date, Term, Year, SubjectCode, SubjectName, Class, Period, StudentID, StudentName, Status, TeacherID, SessionID | บันทึกการเช็คชื่อ |
| `Academic_Records` | Date, Term, Year, SubjectCode, Class, Period, Topic, Present, Absent, Leave, TeacherID, SessionID | บันทึกการสอน (เนื้อหา+สถิติ) |
| `Budgets` | ProjectID, ProjectName, BudgetAmount, UsedAmount, Balance, Status, Year | งบประมาณ |
| `Leave_Records` | StaffName, Type, StartDate, EndDate, Reason, Status, Year | บันทึกการลา |
| `Maintenance` | ID, Location, Issue, Reporter, Status, Technician | บำรุงรักษา |
| `System_Settings` | Key, Value | ตั้งค่าระบบ |
| `Timetable_Database` | [0]SubjectCode, [1]SubjectName, [2]Level, [3]Room, [4]Location, [5]TeacherID, [6]Day, [7]Period, [8]Term, [9]Year | ตารางสอน |
| `Morning_Activity` | Date, Term, Year, Class, StudentID, StudentName, Area_Status, Duty_Status, Flag_Status, TeacherID, SessionID | กิจกรรมหน้าเสาธง |
| `Sarabun_Database` | Timestamp, DocType, DocNumber, Subject, Requester, TargetDate, Status, FileURL, Year | ทะเบียนสารบรรณ |

### ชีตที่สร้างโดย `setupPorPor5Database()`
| Sheet | ใช้งาน |
|---|---|
| `Subject_Config` | ตั้งค่าโครงสร้างวิชา (ตัวชี้วัด, น้ำหนักคะแนน) |
| `Score_Database` | คะแนนรายตัวชี้วัด |
| `Qualitative_Assess` | ประเมินคุณลักษณะอันพึงประสงค์ |
| `Grade_Summary` | สรุปผลการเรียนรายวิชา |
| `Print_Config` | ตั้งค่าหัวกระดาษสำหรับพิมพ์ ปพ.5 |

### ชีตอื่นที่สร้างอัตโนมัติเมื่อใช้งาน
| Sheet | ใช้งาน |
|---|---|
| `Detailed_Lesson_Records` | บันทึกการสอนแบบละเอียด |
| `Score_History` | Log ประวัติการแก้คะแนน |
| `Calendar_Database` | ปฏิทินกิจกรรม |
| `Curriculum_Database` | หลักสูตร / ตัวชี้วัด |
| `User_History_Database` | ประวัติการแก้ไขข้อมูลผู้ใช้ |
| `Club_Database` | ชุมนุม master list (เทอม/ปี) |
| `Club_Advisors` | ครูที่ปรึกษา (many-to-many) |
| `Club_Members` | สมาชิกชุมนุม (1 นักเรียน : 1 ชุมนุม : 1 เทอม) |
| `_Computed_Cache` | precomputed dashboards (nightly trigger) |

### System_Settings Format (แบบใหม่)
```
Row: ["Active", "Term", "1", "2568", ...]        ← เทอมปัจจุบัน
Row: ["TermData", "1_2568", startDate, endDate]  ← วันเริ่ม-สิ้นสุดเทอม
```

---

## Roles & Permissions

| Role | สิทธิ์ |
|---|---|
| `ADMIN` | เข้าถึงทุกส่วน, bypass `verifyTeacherPermission` |
| `TEACHER` | เช็คชื่อ, บันทึกคะแนน, ตารางสอน, ปพ.5 เฉพาะวิชาที่สอน |
| `STUDENT` | ดูข้อมูลตัวเอง |
| `EXECUTIVE` | read-only ภาพรวมโรงเรียน, route → `Page_Dashboard_Executive`, dept ใน `User_Database[row][4]` กำหนด layout (ผอ./วิชาการ/งบประมาณ/บุคคล/ทั่วไป) — KPI strip, alerts, dept-scoped sections, calendar strip, bundle via `getExecutiveDashboardBundle(dept)` |

> Role เก็บใน `User_Database[row][3]` — เปรียบเทียบด้วย `.toUpperCase()`

---

## Authentication & Session

```javascript
// หลัง login สำเร็จ เก็บใน localStorage เสมอ (ไม่มี sessionStorage สำหรับ session หลัก)
localStorage.getItem('pssms_user')       // session หลัก — อายุ 90 วัน
localStorage.getItem('pssms_user_savedAt') // timestamp ที่ save (ms) — สำหรับ expiry check
localStorage.getItem('pssms_creds')      // btoa(user:pass) — สำหรับ silent re-auth

// Structure ของ pssms_user:
{ id, name, role, dept, currentTerm, currentYear }

// keys อื่นใน localStorage:
'pssms_last_page'        // หน้าล่าสุด — restore อัตโนมัติเมื่อ reload
'pssms_theme'            // 'light' | 'dark'
'pssms_saved_accounts'   // JSON array [{u, p: btoa(pass), n, r}] — autofill dropdown
'pssms_debug'            // '1' = เปิด debug overlay

// cache ใน sessionStorage (ไม่ใช่ session หลัก):
`subjects_${userId}_${term}_${year}`   // list วิชาของครู
`atRisk_${userId}_${term}_${year}`     // at-risk cache สำหรับ dashboard
```

**หมายเหตุ**: ไม่มีปุ่ม "จดจำการเข้าสู่ระบบ" อีกต่อไป — login ทุกครั้งบันทึก session 90 วันอัตโนมัติ  
Migration path: `checkAutoLogin()` ย้าย session เก่าจาก sessionStorage → localStorage อัตโนมัติครั้งเดียว

`syncSystemTerm()` — เรียกหลัง login เพื่ออัปเดตเทอม/ปีแบบ silent (ไม่ reload หน้า)

---

## Frontend Routing

```javascript
// เปลี่ยนหน้าด้วย: (ชื่อฟังก์ชันจริงคือ loadPage ไม่ใช่ showPage)
loadPage('Page_Score_Entry')   // โหลด HTML จาก GAS แล้วใส่ใน #page-content

// เรียก Backend:
google.script.run
  .withSuccessHandler(callback)
  .withFailureHandler(errCallback)
  .functionName(args)
```

- ไม่มี URL routing — ทุกอย่างอยู่ใน URL เดียว (GAS Web App URL)
- `loadPage()` บันทึกชื่อหน้าลง `pssms_last_page` และ restore เมื่อ reload

---

## Sidebar Menu — CSS Classes

| Class | ใช้กับ | รูปแบบ |
|---|---|---|
| `nav-link-custom` | หัวเมนูหลัก (Dashboard) | full-width, ไม่มี border-left |
| `dept-btn` | หัว section 4 ฝ่าย + การจัดการระบบ Admin | `font-size: 0.85rem; font-weight: 600`, border-left accent, border-radius 7px |
| `nav-link-sub` | sub-item ใต้ dept-btn | indent `margin: 0 10px 0 22px`, font 0.8rem |
| `menu-divider` | หัวกลุ่ม (การจัดการระบบ, บริหาร 4 ฝ่าย) | uppercase, เส้นขีดล่าง |

> **dept-btn semantic**: ใช้ `<div class="dept-btn">` สำหรับหัว section ที่กดไม่ได้ (4 ฝ่าย), ใช้ `<a class="dept-btn">` สำหรับที่กดได้ — CSS: `div.dept-btn { pointer-events: none }`, `a.dept-btn:hover { background: var(--p-hover) }`  
> หัวฝ่ายทั้ง 4 (วิชาการ/งบประมาณ/บุคคล/ทั่วไป) เป็น `<div>` ทั้งหมด — navigation อยู่ที่ sub-item ใต้แต่ละฝ่าย  
> `nav-link-sub` ใช้เฉพาะ sub-item (เยื้อง indent 22px จากซ้าย)

### Auto-tab navigation pattern (Page_Admin_Users)
```javascript
// ก่อน loadPage — set flag เพื่อให้หน้าปลายทาง switch tab อัตโนมัติ
window._usersAutoTab = 'homeroom';  // ชื่อ tab id prefix (เช่น 'homeroom' → 'homeroom-tab')
loadPage('Page_Admin_Users');
// หน้าปลายทาง: ตรวจ window._usersAutoTab แล้ว bootstrap.Tab.getOrCreateInstance(el).show()
```

---

## ระบบตรวจสิทธิ์ครู (`verifyTeacherPermission`)

```javascript
verifyTeacherPermission(teacherId, subjectCode, className, term, year)
// คืน true/false
// Admin ผ่านทันที
// ครูทั่วไป: เช็คว่ามีใน Timetable_Database (ห้อง+วิชา+เทอม+ปีต้องตรง)
// className format: "ม.1/1" หรือ "1/1"
// subjectCode "hr" = โฮมรูม (อนุโลม)
// subjectCode "CLUB_<id>" = ชุมนุม → เช็ค Club_Advisors (teacherId+term+year) แทน Timetable_Database
```

### Club–Timetable Integration (2026-05-15)

ตารางสอนของครู/Admin แสดงชื่อชุมนุมจริงแทน "ชุมนุม" generic:
- `_getTeacherClubForTerm(teacherId, term, year)` — ค้น Club_Advisors + Club_Database → `{clubId, clubName}`
- `_applyClubOverride(row, club)` — ถ้า SubjectName มี "ชุมนุม" → แทนด้วยชื่อจริง, SubjectCode → `CLUB_<clubId>`
- ใช้ใน: `getTeacherTimetableByDate`, `getTeacherTimetable`, `getTeacherTimetableWithStatus`, `getFilteredTimetables`
- Frontend: `CLUB_xxx` ใน SubjectCode = signal ว่าเป็นคาบชุมนุม → แสดง badge "ชุมนุม" + ชื่อ, ดึงรายชื่อด้วย `getStudentsByClub(clubId)` แทน `getStudentsByClass`

---

## UI Libraries (CDN ใน Index.html)

| Library | Version | ใช้งาน |
|---|---|---|
| Bootstrap | 5.3 | Layout, components |
| Font Awesome | 6 | Icons |
| Chart.js | latest | กราฟ Dashboard |
| FullCalendar | 6 | ปฏิทิน |
| Flatpickr | latest | Date picker |
| Kanit (Google Fonts) | - | Font ภาษาไทย |

---

## Third-party Integrations

- **Notion API** — Todo list, credentials เป็น global var ใน `Code.js` บรรทัด 204-206
- **Google Drive** — อัปโหลดเอกสารสารบรรณ (`DriveApp`)
- **AMSS** (`https://amss.sesaud.go.th`) — scaffold อยู่ใน `Code.js` (`_amssLogin`, `testAmssConnection`, `syncAmssIncoming`) แต่ **parked** เพราะ Cloudflare บล็อก Google server IPs ทุก request ครูใช้ credential `41042010`/`41042010` เก็บใน PropertiesService หรือ fallback hardcode

> **ระวัง:** `NOTION_TOKEN`, `DATABASE_ID`, `PROJECT_ID` hardcode เป็น `var` ใน `Code.js` — อย่า push ขึ้น public repo

---

## ไฟล์ทดสอบ (`ไฟล์ทดสอบ/`)

| ไฟล์ | ใช้ทดสอบ |
|---|---|
| `ปฏิทิน_69-1.csv` | นำเข้าปฏิทินกิจกรรม ภาคเรียน 1/2569 |
| `ตัวชี้วัด.csv` / `.xlsx` | นำเข้าตัวชี้วัดหลักสูตร |

---

## Convention สำคัญ

- ปีการศึกษาเป็น **พ.ศ.** (2568, 2569) ไม่ใช่ ค.ศ.
- เทอม: `"1"` หรือ `"2"` (string ไม่ใช่ number)
- ID ผู้ใช้: เปรียบเทียบด้วย `String(x).trim()` เสมอ
- ทุกฟังก์ชัน GAS ต้องรองรับทั้ง Admin และ non-Admin
- เพิ่มฟีเจอร์ใหม่ → คำนึงถึง role และ term/year เสมอ
- ภาษาไทยทั้งหมด (UI + error message)
- Default admin จาก `setupDatabase()`: username `admin` / password `1234` — เปลี่ยนก่อน production
- **GAS serialization gotcha**: `google.script.run` คืน `null` ให้ `withSuccessHandler` ถ้า return value มี `Date` object → แก้ด้วย `Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm')` หรือ `String(d)` ทุกครั้งก่อน return
- `getSarabunHistory`: skip row ที่ทั้ง `docNumber` และ `docType` ว่าง (ป้องกัน empty rows จาก deleted sheet data)

---

## Design System (2026-05-18)

ดู `DESIGN.md` ที่ project root สำหรับ tokens + component patterns ทั้งหมด.

### Quick reference
- Wrapper: `<div class="container-fluid pssms-page pssms-dept-X">` เลือก dept: `academic` (ฟ้า) / `budget` (เขียว) / `personnel` (น้ำเงิน) / `general` (ส้ม)
- ทุก color reference ที่ต้องเปลี่ยนตามฝ่าย → `var(--p-accent)` (ห้าม hardcode hex)
- Card: `pssms-card` + `pssms-card-header`
- Button: `btn-accent` (primary, dept-colored) / `btn-soft` (secondary)
- Table: `pssms-table`
- Modal: add class `pssms-modal` + neutral white header
- Tabs: `<ul class="nav pssms-tabs mb-3 gap-2">` — underline active, muted inactive, hover bg; สลับ tab ด้วย `bootstrap.Tab.getOrCreateInstance(el).show()`; lazy init ด้วย `shown.bs.tab` event
- Skeleton: `<span class="pssms-skel rounded-pill" style="width:Xpx;height:26px;display:inline-block;">` สำหรับ pill-shaped skeleton (เช่น calendar badges)

---

## Performance & Debug (2026-05-12 optimization pass)

ดู `PERF_OPTIMIZATION.md` สำหรับรายละเอียดทั้งหมด.

### Cache layers
1. **ScriptCache** (`src/Cache.js` — `getCached(key, ttl, fetcher)`) — shared across users, 5-30 min TTL
2. **PropertiesService** (`getActiveTermYear()`) — sub-ms active term/year
3. **_Computed_Cache sheet** — nightly precomputed risk/atRisk dashboards (TTL 12h on read)

### Cache invalidation hooks
Every write path เรียก `invalidateCache(key)` หรือ `invalidateCacheKeys([])`:
- `saveSystemConfig` → system_config, available_terms, all_users + `setActiveTermYear`
- `addUser`/`editUser`/`deleteUser`/`importStudentCSV`/`importTeacherCSV`/`promoteStudentsToNextYear` → all_users
- `saveCalendarEvent`/`deleteCalendarEvent`/`importCalendarCSV` → calendar_events
- `importCurriculumCSV` → curriculum_all

### Debug mode
**Backend**: GAS Editor → `setDebugMode(true)` / `setDebugMode(false)` → `clasp logs --watch`
**Frontend**: เปิด URL ลงท้าย `?debug=1` → overlay panel มุมขวาล่าง (CALL/CACHE/PERF/PROPS logs)

### Precompute trigger setup (one-time)
ใน GAS Editor:
```javascript
setupPrecomputeTrigger();  // ตั้ง trigger ทุกคืน 02:00
precomputeNow();            // prime cache ครั้งแรก
```

### Bundled GAS calls
- `getTeacherDashboardBundle(teacherId, term, year)` — 4 sections (timetable + risk + atRisk + calendar)
- `getAdminDashboardBundle()` — 5 sections (stats + summary + calendar + terms + config)
- `getExecutiveDashboardBundle(dept)` — 6 sections (kpi + academic + budget + personnel + general + calendar)
- Frontend dashboard (`Scripts_Teacher.html` → `initTeacherDashboard()`) ใช้ bundle, fallback ฟังก์ชันแยกเมื่อ bundle fail
- Dashboard guard pattern: เช็ค `document.getElementById('<page-root-id>')` ก่อน inject DOM ทุกครั้ง — ป้องกัน race condition เมื่อ navigate ออกก่อน bundle callback กลับ

### Adding new cached function — pattern
```javascript
function getMyData() {
  return getCached('my_data_key', 300, function() {
    // ... expensive read
  });
}
// + invalidate hook ในทุก write path:
function saveMyData() {
  // ... write
  invalidateCache('my_data_key');
}
```

---

## Web Prototype (Node + Railway PostgreSQL)

Parallel codebase ที่ `web/` — Express + PostgreSQL replicating GAS behavior. Frontend ใช้ `gas-shim.js` แทน `google.script.run`: ทุก call `google.script.run.fn(a, b, c)` → `POST /api/gas/fn` body `{ args: [a, b, c] }`.

### Function signature convention
Backend handlers รับ args เป็น array destructured:
```javascript
async function fnName([arg1, arg2, arg3]) { ... }
```
**ต้องตรงกับ frontend call** — frontend ส่ง `args = [a, b, c]` ตามลำดับที่ `google.script.run.fn(a, b, c)` ระบุ. Signature mismatch = bug ที่หาย ๆ ยาก (เช่น args shift จน param หลังกลายเป็น undefined).

ตัวอย่างที่เคยเจอ:
- `editUser(form)` — รับ object เดียว, อย่าใส่ `[username, form]` แยก
- `getAllInOneScoreGridData(subjectCode, className, term, year)` — 4 args, ไม่มี teacherId
- `getSemesterReport(subjectCode, className, term, year)` — 4 args, ไม่ใช่ `(teacherId, term, year)`
- `createClub(payload)` / `updateClub(payload)` — 1 object, clubId อยู่ใน payload (สำหรับ update). createClub generate clubId ใหม่ใน backend เสมอ (format `CLUB${Date.now()}` ตรง GAS)

### Field naming
GAS form objects ใช้ลำดับ priority `fullname > fullName > full_name` และ `dept > department` — backend ต้องรองรับทั้งคู่. Helper pattern ใน `users.js`:
```javascript
const pickName = (u) => String(u.fullname || u.fullName || u.full_name || '').trim();
const pickDept = (u) => String(u.department || u.dept || '').trim();
```

### Schema differences vs GAS Sheets

| ตาราง | คอลัมน์ที่ต่างจาก GAS / ข้อควรระวัง |
|---|---|
| `score_database` | `score TEXT` (ไม่ใช่ numeric) เพราะ remark column เก็บ `'-'`,`'ร'`,`'มส'` ปนกับตัวเลข. PK = composite `(student_id, subject_code, indicator_id, term, year)`. ใช้ `ON CONFLICT(student_id,subject_code,indicator_id,term,year)` |
| `score_history` | `old_score`, `new_score` ทั้งคู่เป็น `TEXT` (เหตุผลเดียวกับ score) |
| `subject_config` | มี `exam_indicators_json JSONB` (ตรงกับ col[8] ของ GAS); PK = `(subject_code, class_name, term, year)` |
| `qualitative_assess` | มี sub-score columns: `char1-4, char_total, char_grade, read1-4, read_total, read_grade, comp`. PK = `(student_id, subject_code, term, year)` |
| `users` | ไม่มี year snapshot — promote update in-place. snapshot ปีเก่าเก็บใน `user_history` (ดู Historical roster) |
| `user_history` | `username, action, changed_by, old_data jsonb, new_data jsonb, timestamp` — audit log + snapshot ตอน promote |

### Score indicator_id convention
- `formative_0`, `formative_1`, ... = คะแนนเก็บแต่ละชิ้น (index ตรงกับ `subject_config.indicators_json[i]`)
- `midterm` = สอบกลาง
- `midterm_re` = ซ่อมกลาง
- `final` = สอบปลาย
- `remark` = `ร` / `มส` / `-`

### Attendance & report logic — shared module
`web/functions/attendanceReport.js` เป็น single source of truth สำหรับ:
- `getSemesterReport([subjectCode, className, term, year])` — หน้ารายงานสถิติเวลาเรียน
- `getAllSubjectsReport([teacherId, term, year])` — ทุกวิชาที่ครูสอน
- `getTeacherAtRiskDashboard([teacherId, term, year])` — Dashboard card

Formula (ตรง GAS):
```
periodsPerWeek = COUNT timetable rows (subject+level+room+term+year)
totalCoursePeriods = periodsPerWeek × 20  (fallback 3 ถ้าไม่เจอ)
percent = ((totalCoursePeriods − absent − leave) / totalCoursePeriods) × 100
buckets: <60 critical, 60-79 ms, 80-84 risk (≤85 เข้า list)
```

อย่าใช้ `COUNT(*)` ของ attendance เป็นตัวหาร — จะ inflate percent.

### Teacher dashboard bundle
`getTeacherDashboardBundle([teacherId, term, year])` parallel sections:
- `timetable` (today schedule)
- `calendarEvents` (14-day strip)
- `riskDashboard` — grade-based (0, ร, มส.) จาก `grade_summary`
- `atRiskDashboard` — attendance-based จาก `attendanceReport.getTeacherAtRiskDashboard`

แต่ละ section wrap ด้วย `section()` helper → `{ok:true,data}` หรือ `{ok:false,error}`.

### Historical roster fallback chain
`getStudentsByClass([className, year])` priority:
1. `users` table by `class+year+status='ปกติ'` (exact)
2a. **historical only** — `user_history WHERE action='promote' AND old_data->>year=$y AND old_data->>department=$class` (มี email/password ครบ)
2b. **historical only** — `DISTINCT attendance(student_id, student_name, class)` ที่ `year=$y AND class=$class` (no email/password)
3. **current year only** — `users` ignore year filter (last resort)

ปีถัดไป promote → `promoteStudentsToNextYear()` จะ INSERT user_history ทุก row ก่อน UPDATE → snapshot ครบทุกครั้ง.

### Write function return format
Frontend เช็ค `res.status === 'success'` ทุก write function. คืน:
```javascript
{ status: 'success', message: 'ข้อความภาษาไทย' }
```
ไม่ใช่ `{ success: true }` (GAS frontend ไม่ตรวจ).
