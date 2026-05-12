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

### 10% Accent
- `--p-accent` `#e83e8c` — primary CTA, brand
- `--p-accent-hover` `#d1336e`
- `--p-accent-soft` `#fff4f7` — selected row highlight, today

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
<div class="container-fluid pssms-page">
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

- [ ] Header: title + subtitle + actions pattern
- [ ] Replace `card shadow-sm rounded-4` → `pssms-card`
- [ ] Replace primary action color buttons → `btn-accent`
- [ ] Replace `btn-outline-secondary shadow-sm` → `btn-soft`
- [ ] Tables → `pssms-table`
- [ ] Modal → add `pssms-modal` class + neutral header
- [ ] Drop `py-4 mb-4` → use `pssms-page` shell
- [ ] Drop `animate__animated` noise
