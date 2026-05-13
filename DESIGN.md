# PSSMS Design System

Minimal 70-20-10 palette + standard component patterns สำหรับทุก non-dashboard page.

## Color tokens (`src/Styles.html` :root)

### 70% Neutral
- `--p-surface` `#ffffff` — card, modal
- `--p-subtle` `#fafbfc` — page bg, table header
- `--p-border` `#eef0f3` — divider/border
- `--p-border-strong` `#dde1e6` — hovered border
- `--p-hover` `#f6f8fa` — row hover

### 20% Secondary
- `--p-text` `#1f2937` — heading + body
- `--p-text-muted` `#6c757d` — subtitle, hint
- `--p-chip-bg` `#f5f6f8` — passive badge
- `--p-chip-text` `#444`

### 10% Accent (global brand)
- `--p-accent` `#e83e8c` — primary CTA, brand (pink) — DEFAULT
- `--p-accent-hover` `#d1336e`
- `--p-accent-soft` `#fff4f7` — selected row highlight, today

### Department palette — scoped overrides
Apply `pssms-dept-<name>` class on page wrapper. The scope re-binds `--p-accent` so every `.btn-accent`, icon, and hover inside inherits automatically.

| Class | สี | Hex |
|---|---|---|
| `pssms-dept-academic` | ฟ้า | `#3498db` |
| `pssms-dept-budget` | เขียว | `#2ecc71` |
| `pssms-dept-personnel` | น้ำเงิน | `#2563eb` |
| `pssms-dept-general` | ส้ม | `#f97316` |

Each scope also overrides `--p-accent-hover` and `--p-accent-soft`.

**Pages → dept mapping** (current):
- Academic, Score_Entry, Subject_Config, Academic_Report, Grade_Summary, Lesson_History, Admin_Timetable, Admin_Curriculum, Admin_Clubs, Student_Clubs, Teacher_Clubs → `academic`
- Personnel, Admin_Users → `personnel`
- Budget → `budget`
- General, Calendar → `general`
- Admin_Settings, Login → no dept class (keeps global brand pink)

### Semantic states (sparingly)
- `--p-info` `#0d6efd`
- `--p-success` `#198754`
- `--p-warning` `#ffc107`
- `--p-danger` `#dc3545`

## Geometry
- `--p-radius` `10px` — card, modal, button
- `--p-radius-sm` `6px` — inputs, chips

## Rules
1. Page bg ใช้ neutral เท่านั้น
2. Accent ใช้เฉพาะ primary CTA + active state + active link
3. State colors ใช้เฉพาะตอนสื่อ meaning จริงๆ (success/error/warn/info)
4. ห้ามใช้ `shadow-sm` + `rounded-4` — ใช้ `pssms-card` แทน
5. ห้ามใช้ `btn-success/btn-warning/btn-danger` สำหรับ generic action — ใช้ `btn-accent` หรือ `btn-soft`
6. Heading: `fw-bold` + letter-spacing `-0.01em`
7. Subtitle: `text-muted small`

## Page shell

```html
<div class="container-fluid pssms-page pssms-dept-academic">
  <div class="pssms-page-header">
    <div>
      <h4><i class="fas fa-X me-2" style="color: var(--p-accent);"></i>ชื่อหน้า</h4>
      <p class="pssms-subtitle">คำอธิบายสั้น</p>
    </div>
    <div class="pssms-actions">
      <button class="btn btn-soft btn-sm"><i class="fas fa-sync me-1"></i>รีเฟรช</button>
      <button class="btn btn-accent btn-sm"><i class="fas fa-plus me-1"></i>เพิ่ม</button>
    </div>
  </div>

  <div class="pssms-card">
    <div class="pssms-card-header">หัวข้อ</div>
    <!-- content -->
  </div>
</div>
```

**สำคัญ**: ทุก color reference ที่ต้องเปลี่ยนตามฝ่าย ใช้ `var(--p-accent)` เท่านั้น — อย่า hardcode hex. ตัวอย่าง:
- ✅ `style="color: var(--p-accent);"`
- ✅ `<button class="btn btn-accent">`
- ❌ `style="color: #3498db;"` (จะไม่เปลี่ยนตามฝ่าย)
- ❌ `class="text-primary"` (Bootstrap blue, ไม่ใช่ token)

## Tables

```html
<div class="pssms-card p-0">
  <table class="pssms-table mb-0">
    <thead><tr><th>คอลัมน์</th></tr></thead>
    <tbody><tr><td>ข้อมูล</td></tr></tbody>
  </table>
</div>
```

## Modals

```html
<div class="modal fade pssms-modal" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header"><h5 class="modal-title">หัวข้อ</h5><button class="btn-close" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">...</div>
      <div class="modal-footer">
        <button class="btn btn-soft btn-sm" data-bs-dismiss="modal">ยกเลิก</button>
        <button class="btn btn-accent btn-sm">บันทึก</button>
      </div>
    </div>
  </div>
</div>
```

## Buttons reference

| Use | Class |
|---|---|
| Primary CTA | `btn btn-accent` |
| Secondary | `btn btn-soft` |
| Destructive | `btn btn-outline-danger` (use sparingly) |
| Success state confirm | `btn btn-success` (only on completion modals) |

## Migration checklist (per page)

- [ ] Wrapper: `<div class="container-fluid pssms-page pssms-dept-X">` (เลือก dept ตาม mapping)
- [ ] Header: title + subtitle + actions pattern
- [ ] Replace `card shadow-sm rounded-4` → `pssms-card`
- [ ] Replace primary action color buttons → `btn-accent`
- [ ] Replace `btn-outline-secondary shadow-sm` → `btn-soft`
- [ ] Tables → `pssms-table`
- [ ] Modal → add `pssms-modal` class + neutral header
- [ ] Drop `py-4 mb-4` → use `pssms-page` shell
- [ ] Drop `animate__animated` noise
- [ ] Hardcoded colors (#e83e8c, #3498db, ...) → `var(--p-accent)`

## หน้าใหม่ — สร้างยังไง

1. **เริ่มจาก template ใน "Page shell"** ข้างบน
2. **เลือก dept class** ตาม category (academic / budget / personnel / general)
3. **Layout**: `row g-3` + cols, ใน col ใส่ `pssms-card`
4. **Tables**: `<table class="pssms-table">` ใน `<div class="pssms-card p-0">`
5. **Forms**: `form-control` + `form-select` ปกติ (focus border ใช้ `--p-accent` อัตโนมัติ)
6. **Modals**: `<div class="modal fade pssms-modal">` + footer มี btn-soft (cancel) + btn-accent (save)
7. **Status colors**: ใช้ `text-success / text-danger / text-warning` ได้ตามปกติ — แต่อย่าใช้สี state เป็น primary CTA
