# Maternity Leave Timeline Calculator

An accessible, client-side web prototype that replaces a legacy Excel (`.xlsm`) maternity leave worksheet used across the University of California system. Given a handful of simple inputs, it generates a complete leave timeline — Pregnancy Disability Leave (PDL), Family Medical Leave (FML), California Family Rights Act (CFRA) leave, the Lincoln Financial disability waiting period and income window, sick/vacation spend, PFCB, and Child Caring Leave (CCL) — and renders it as both a text timeline and a monthly calendar view.

The tool runs entirely in the browser. It makes no network calls, stores nothing, and has no dependencies.

## Live site

Once published via GitHub Pages, the calculator is available at:

**https://jdhori.github.io/MaternityLeaveCalculator/**

## Contents

| File | Purpose |
|---|---|
| `index.html` | Main entry point — semantic form + results layout |
| `app.js` | All calculation logic, calendar rendering, and form handling |
| `styles.css` | Full stylesheet (light + dark theme, print styles) |
| `theme-loader.js` | Tiny pre-paint script that applies stored theme to avoid flash of unstyled content |
| `leave_calculator.html` | Standalone single-file build with inline CSS/JS (for offline or email distribution) |

Open `index.html` in a browser to use. No build step, no install.

---

## From Excel to the web

The original tool was an `.xlsm` workbook that combined static reference tables, date math, and macro logic to compute a leave schedule. It worked, but it had real problems:

- **Accessibility.** Excel macros are largely unusable for screen reader users; the formula-driven cells were visually dense and hard to interpret.
- **Platform lock-in.** Required a desktop version of Excel with macros enabled — a blocker for Mac users, web-Excel users, and anyone with organizational macro restrictions.
- **Distribution.** Every update required redistributing a fresh `.xlsm` and trusting users to replace their local copy.
- **Auditability.** Formulas hidden in cells made it hard to see *why* a date moved when an input changed.

### What the conversion did

Each Excel mechanism was mapped to a direct web equivalent:

| Excel mechanism | Web equivalent |
|---|---|
| `WORKDAY.INTL(start, n, 1, holidays)` | `addWorkdays(start, n)` in `app.js` |
| `NETWORKDAYS.INTL(start, end, 1, holidays)` | `networkdays(start, end)` in `app.js` |
| Holiday reference table | `HOLIDAYS` Set in `app.js`, seeded from the UCOP systemwide calendar (2023–2027) + Winter Curtailment |
| Row 53 — claim file date (28 days before leave start) | `fileClaim = addDays(leaveStart, -28)` |
| Row 59 — Lincoln Financial income window | Computed from waiting-period end to PDL end |
| Row 62 — PDL duration (42 natural / 56 C-section, inclusive of birth day) | `pdlEnd = addDays(pdlAnchor, pdlDurationDays - 1)` |
| Row 65 — FMLA (84 days, calendar-year cap) | `fmlBegin`/`fmlEnd` + carry-over to new year |
| Row 71 — CFRA (84 days after PDL) | `cfraBegin`/`cfraEnd` |
| Row 74 — PFCB | User-supplied weeks, default start = day after PDL |
| L30 — End of PIE (31 days after birth) | `endPIE` |
| Sick/vacation hour → day conversion | `Math.floor(hours / effHrsPerDay)`, capped by schedule type |
| Cell-driven inputs | Real HTML `<form>` with labels, hints, and error messages |
| Macro UI | Submit button + live-region status announcements |
| Color-coded cells | Semantic category classes (`cat-pdl`, `cat-fml`, etc.) rendered both in a timeline list and in per-month calendars |

All math runs client-side in a single `calculate(input)` function in `app.js`. Nothing leaves the browser.

---

## Features

### Inputs the form accepts

- **Leave start date** (required) — the first day of leave. Variable employees should pick the actual first day off, not necessarily the day after the last shift.
- **Estimated due date** (required) — used as the PDL anchor when no actual birth date is provided.
- **Actual birth date** (optional) — once known, re-anchors PDL and computes End of PIE + first birthday.
- **Estimated return to work** (optional) — shown as a milestone on the timeline.
- **Delivery type** (required) — Natural (42-day PDL) or C-section (56-day PDL).
- **Employee schedule type**:
  - *Regular* — 40 hr/wk, 5 × 8, sick cap of 22 workdays.
  - *Variable* — user-specified `hoursPerDay` and `daysPerWeek` (3×12, 4×10, etc.), sick cap of 30 calendar days.
- **Sick hours** (required) and **vacation hours** (optional) — converted to working days using the selected schedule.
- **Fallback strategy** — if sick leave doesn't cover the waiting period, either use vacation first or go straight to leave-without-pay.
- **Disability waiting period** — 7 / 14 / 30 / 90 / 180 days (Lincoln Financial plan-selectable).
- **Job-protection eligibility** — four independent checkboxes for PDL, FMLA, CFRA, and FMLA/CFRA. PDL is checked by default. FMLA and CFRA each gate their own timeline rows; the combined "FMLA/CFRA" option turns both on (the typical UC case where the two run concurrently).
- **PFCB weeks** (0–8) and **PFCB begin date** (optional; defaults to day after PDL ends).
- **CCL weeks** (0–26) and **CCL anchor** — start after end of PDL, FMLA, or CFRA. Falls back automatically if the requested anchor isn't available in the scenario. Up to 14 weeks after FMLA/CFRA ends, or up to 26 weeks if not eligible for FMLA/CFRA — described in the in-form info button.

### What the calculator outputs

**Summary callout** — at-a-glance dates for PDL, FMLA, CFRA, Lincoln income, PFCB, and CCL.

**Leave events timeline** — a chronological, categorized list including:

- Leave start date
- Claim file suggestion (28 days before leave starts; LOA recommends 1–2 weeks before, plan ceiling is 30 days)
- Estimated due date and, if provided, actual birth date
- Sick leave span (with note if capped by schedule maximum)
- Vacation span *or* a note explaining why vacation isn't being used
- Disability waiting period
- Lincoln Financial disability income window (with note if PDL ends before benefits would start)
- Pregnancy Disability Leave (PDL) — when the employee is PDL-eligible
- Family & Medical Leave Act (FMLA) — including split display when it crosses a calendar year
- California Family Rights Act (CFRA)
- Pay for Family Care and Bonding (PFCB)
- Child Caring Leave (CCL) — annotated if a fallback anchor was used
- Return-to-work estimate
- End of PIE (baby-enrollment deadline, 31 days post-birth)
- Baby's first birthday (parental-bonding window end)

**Monthly calendar view** — per-month grids (Sun–Sat) covering the whole leave window, showing:

- Category shading per day (PDL, FML, CFRA, Lincoln, waiting period, PFCB, CCL, sick, vacation, leave-no-salary)
- Holiday stars for UC systemwide holidays and Winter Curtailment days
- Legend with color swatches

### Holiday data

`HOLIDAYS` in `app.js` includes every UCOP systemwide holiday from **2023 through 2026** and projected 2027 dates using standard UC observation rules. Winter Curtailment days commonly observed across UCOP, UCB, UCD, UCI, UCLA, UCM, UCR, UCSD, UCSF, UCSB, and UCSC are included as an inclusive union, because individual campuses vary. The in-app disclaimer directs users to verify specific dates with their local HR.

### Accessibility

Designed from the ground up to work for screen reader, keyboard, and low-vision users:

- Semantic structure — `<header>`, `<main>`, `<aside>`, `<fieldset>`/`<legend>`, labeled form controls, heading hierarchy.
- Skip-to-main-content link.
- **Error summary** with `role="alert"`, focused on validation failure, each item linking and focusing the offending field.
- Per-field `aria-describedby` wiring for both hints and error messages.
- `aria-live="polite"` regions for dynamic selects and a live status announcer for results.
- Full keyboard operation — no mouse-only interactions.
- Visible focus rings that meet contrast requirements in both themes.
- Accessible dark mode toggle with `aria-pressed` state and `prefers-color-scheme` default.
- Print stylesheet so the timeline + calendar can be printed cleanly.
- Never builds HTML from dynamic values — all user data flows through `textContent` (no injection sinks).

### Security posture

- Strict Content Security Policy via `<meta http-equiv="Content-Security-Policy">`. `default-src 'none'`; only same-origin scripts, styles, images, and fonts allowed. No inline scripts, no inline styles, no external requests, no embedding via iframe.
- `X-Content-Type-Options: nosniff`, `referrer: strict-origin-when-cross-origin`.
- `connect-src 'none'` — the page cannot make network requests. Everything is computed locally.
- No cookies. No tracking. No analytics.
- Local theme preference is the only persistent state (localStorage key `uc-leave-calc-theme`).

### Themes & responsiveness

- Light and dark themes; preference persists via `localStorage`, respects `prefers-color-scheme` on first visit.
- `theme-loader.js` is a blocking `<head>` script that sets `data-theme` before the first paint to avoid flash-of-unstyled-content.
- Responsive single-column → two-column layout.

### Controls

- **Calculate timeline** — submit and render.
- **Reset** — clear all fields and results.
- **Print** — browser print dialog with print-optimized layout.
- **Save as PDF** — triggers the browser's print-to-PDF pipeline and sets a descriptive default filename (`UC-Maternity-Leave-Timeline-YYYY-MM-DD.pdf`, stamped with the Last Day Worked). In the print dialog, pick **Save as PDF** as the destination. Reuses the same `@media print` styles as Print, so the PDF is clean (form hidden, accordions open, calendars in a 3-column grid, good page breaks). No external libraries — CSP stays strict.

---

## Running locally

```bash
# From the folder:
python3 -m http.server 8000
# then open http://localhost:8000/
```

Or just double-click `index.html` — it works from the `file://` scheme too. For fully offline / single-file distribution, use `leave_calculator.html`.

---

## Disclaimer

This is a **work-in-progress accessibility prototype**. It is not an application for, or a guarantee of, disability benefits through Lincoln Financial or any other carrier. Always confirm specific plan details, eligibility, and dates with your campus HR Leave Administration office. Verify holiday observances against your local campus calendar — especially Winter Curtailment, which varies by location.
