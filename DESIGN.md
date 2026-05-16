# PSSMS Design System

60-30-10 palette + token-driven layout + frosted-glass sticky bars + dept scoping. Used by every non-dashboard page.

## Color tokens (`src/Styles.html` :root)

### 60% Dominant — page bg
- `--p-dom` `#f4fbfc` — page background

### 30% Secondary — card headers, subtle surfaces
- `--p-sec` `#cce8eb` — card header bg, progress track, skeleton stripe
- `--p-surface` `#ffffff` — card body
- `--p-subtle` `#f9fafb` — input bg, table row alt
- `--p-border` `#e5e7eb` — divider
- `--p-border-strong` `#dde1e6` — emphasized border
- `--p-hover` `#f3f4f6` — row hover

### 10% Accent — CTA + active state
- `--p-accent` `#076874` — primary action (default Admin teal)
- `--p-accent-hover` `#054f59`
- `--p-accent-soft` `#e0f4f6` — focus ring, selected highlight

### Text
- `--p-text` `#1f2937`
- `--p-text-muted` `#6b7280`

### Geometry
- `--p-radius` `14px` — card, modal
- `--p-radius-sm` `8px` — input, chip

### Dept palette — body-scoped overrides

`body.dept-<name>` (set by JS routing in `_setBodyDept()`) AND `.pssms-dept-<name>` (set on page wrapper) re-bind all three layers (`--p-dom`, `--p-sec`, `--p-accent`).

| Dept | Accent | dom | sec |
|---|---|---|---|
| Admin/default (no class) | `#076874` | `#f4fbfc` | `#cce8eb` |
| academic | `#0d95a0` | `#f0fbfd` | `#b8e2e7` |
| budget | `#4a8a54` | `#f2faf3` | `#c4e2c8` |
| personnel | `#c4501e` | `#fef6f3` | `#f7d0c4` |
| general | `#7a5e00` | `#fefdf0` | `#f0e5b0` |

**JS routing** (Scripts_Core.html `_setBodyDept(pageName)` called inside `setupPageContent()`):

```js
const PAGE_DEPT = {
  Page_Academic:'academic', Page_Score_Entry:'academic', Page_Subject_Config:'academic',
  Page_Academic_Report:'academic', Page_Lesson_History:'academic',
  Page_Admin_Timetable:'academic', Page_Grade_Summary:'academic', Page_Admin_Curriculum:'academic',
  Page_Budget:'budget',
  Page_Personnel:'personnel', Page_Leave_Admin:'personnel',
  Page_Leave_Request:'personnel', Page_Substitute_Admin:'personnel', Page_Admin_Users:'personnel',
  Page_General:'general'
};
```

Pages not in map → no body dept class → Admin neutral tokens.

### Semantic states (sparingly)
- `--p-info` `#0d6efd`
- `--p-success` `#198754`
- `--p-warning` `#ffc107`
- `--p-danger` `#dc3545`

## Typography hierarchy

Font: **IBM Plex Sans Thai** (300/400/500/600/700).

| Element | Size | Weight |
|---|---|---|
| Card title (`pssms-card-header`) | 1.1rem | 700 |
| Form label / dropdown-item | 0.85rem | 500 |
| Form input (`form-control`, `form-select`) | 0.82rem | 400 |
| Table cell primary (tbody td) | 0.88rem | 400 |
| Table cell secondary (.text-muted, small inside td) | 0.72rem | 400 |
| Column header (thead th) | 0.7rem | 500, opacity .75 |
| Sub-menu (`nav-link-sub`) | 0.8rem | 400 |

## Card patterns

### Standard card
```html
<div class="pssms-card">
  <div class="pssms-card-header">หัวข้อ</div>
  <!-- body -->
</div>
```

`pssms-card` has `padding: 1rem; overflow: hidden`. `pssms-card-header` uses `margin: -1rem -1rem .25rem -1rem` to bleed flush to card edges (no padding gap). Header bg = `var(--p-sec)`.

### Card with table (no body padding)
```html
<div class="pssms-card p-0">
  <div class="pssms-card-header">หัวข้อ</div>
  <table class="pssms-table">…</table>
</div>
```

The negative-margin on header still works because `overflow: hidden` on card clips overflow.

## Sticky bars (frosted glass)

### Sidebar — `#sidebar`
- `position: sticky; top: 0; height: 100vh`
- Re-defines neutral Admin tokens locally so dept-tinted pages don't bleed in
- bg: `rgba(255,255,255,0.85)` + `backdrop-filter: blur(10px)` (light) / `rgba(14,28,31,0.92)` (dark)
- Bottom fade via `mask-image: linear-gradient(to bottom, #000 calc(100% - 48px), transparent)` — hints scroll
- Bootstrap `.bg-dark` / `.text-white` on `<nav id="sidebar">` overridden via `#sidebar.bg-dark` specificity

### Top navbar — `#content > .navbar`
- `position: sticky; top: 0`
- bg: `rgba(255,255,255,0.85)` (light) / `rgba(14,28,31,0.92)` (dark) + backdrop blur

### Sticky thead — `#allInOneTable`, `.grid-scroll-container table`
- `position: sticky; top: 0; z-index: 10`
- frosted bg `rgba(255,255,255,0.85)` + blur
- works only in scrollable container

## Tables

```html
<div class="pssms-card p-0">
  <table class="pssms-table">
    <thead><tr><th>คอลัมน์</th></tr></thead>
    <tbody><tr><td>ข้อมูล <div class="text-muted small">รอง</div></td></tr></tbody>
  </table>
</div>
```

Styles also apply to Bootstrap `<table class="table">` via `#page-content .table` scope. Cell secondary (`.text-muted` / `small` inside `<td>`) auto-sizes to 0.72rem.

## Dark mode

`body.dark-mode` (toggle in Scripts_Core.html line ~1051). Re-binds all tokens to dark values.

### Critical gotcha — dept wrapper cascade

CSS variables cascade **by DOM tree**, not selector specificity. `.pssms-dept-academic` on a page wrapper re-binds `--p-sec` to a LIGHT value locally. Even with `body.dark-mode` setting dark `--p-sec`, the deeper wrapper wins.

Fix: explicit `body.dark-mode .pssms-dept-*` rules re-apply dark dept tokens. All four depts are covered in Styles.html.

### Bottom fade overlay
```html
<div class="pssms-fade-bottom"></div>
```
Fades to `var(--p-surface)` — white in light, `#1e2024` in dark. Used for scroll-list hints (e.g., To-Do list).

## Forms

`form-control` / `form-select` auto-styled:
- font-size 0.82rem
- padding 0.4rem 0.7rem
- focus border = `--p-accent`, ring = `--p-accent-soft`

Dropdowns (`dropdown-menu`, `dropdown-item`) at 0.85rem.

## Modals

```html
<div class="modal fade pssms-modal" tabindex="-1">…</div>
```
Adds neutral white header + token-aware body bg in dark mode.

## Buttons

| Use | Class |
|---|---|
| Primary CTA (dept-colored) | `btn btn-accent` |
| Secondary | `btn btn-soft` |
| Destructive | `btn btn-outline-danger` |

## Skeleton loaders

`pssms-skel` (shimmer animation using `--p-sec` / `--p-dom` gradient). Helpers in Scripts_Core.html:
- `_skelCards(n, h)` — n×h-px placeholder cards
- `_skelTbody(cols)` — single tbody row spanning N cols
- `_skelRows(n, cols)` — n rows × cols cells

**Do not** replace button loading states (`btn.innerHTML = spinner`) during save/submit — those are intentional action feedback.

## Rules

1. Page bg uses `--p-dom` only — never hardcode.
2. Card body uses `--p-surface`; card header uses `--p-sec`.
3. Every color that must shift per dept → `var(--p-accent)` — no hex.
4. No `shadow-sm` + `rounded-4` — use `pssms-card`.
5. No `btn-success/warning/danger` for generic actions — use `btn-accent` / `btn-soft`.
6. Heading: `fw-bold` + letter-spacing `-0.01em`.
7. State colors (`text-success/danger/warning`) only when conveying real meaning.

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

## Component class reference

| Class | Usage |
|---|---|
| `pssms-page` | Page wrapper (padding 1.5rem) |
| `pssms-card` | Card (border, radius, surface bg, overflow:hidden, padding 1rem) |
| `pssms-card-header` | Card header — bg `--p-sec`, font 1.1rem 700, bleeds flush to card edges |
| `pssms-table` | Minimal table |
| `pssms-chip` | Soft badge/chip |
| `pssms-icon-circle` | Circular icon container (accent-soft bg) |
| `pssms-kpi-value` | Large KPI number (accent color, 1.75rem) |
| `pssms-kpi-sub` | KPI subtitle (0.78rem muted) |
| `pssms-card-link` | Clickable card (hover border + glow) |
| `pssms-card-accent` | Tinted card (accent-soft bg) |
| `pssms-empty` | Empty state (centered, muted icon) |
| `pssms-status` + `pssms-status-{pending,approved,rejected,info}` | Status chip |
| `pssms-progress` / `pssms-progress-bar` | Progress (track sec, fill accent) |
| `pssms-badge` | Number badge (accent, 20px circle) |
| `pssms-pills` / `pssms-pill` (+ `.active`) | Nav pill set |
| `btn-accent` / `btn-soft` | Buttons |
| `pssms-modal` | Modal (neutral header, dark-aware body) |
| `pssms-skel` / `pssms-skel-circle` | Skeleton shimmer |
| `pssms-fade-bottom` | Bottom fade overlay (dark-aware) |
