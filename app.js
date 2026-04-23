/* ============================================================
   UC systemwide holiday table
   -----------------------------------------------------------
   Baseline: UCOP systemwide holiday calendar (authoritative).
   Plus: Winter Curtailment days commonly observed across UC
   campuses. Individual campuses may add or remove days — this
   list takes an inclusive "union" approach so the calculator
   works for employees at any UC location (UCOP, UCB, UCD, UCI,
   UCLA, UCM, UCR, UCSD, UCSF, UCSB, UCSC). Verify against your
   local HR if a specific date matters.
   Source: ucop.edu/local-human-resources/resources/holiday-calendar.html
   ============================================================ */
const HOLIDAYS = new Set([
  /* 2023 */
  "2023-01-02", // New Year Holiday (observed; Jan 1 was Sunday)
  "2023-01-16", // Martin Luther King Jr. Day
  "2023-02-20", // Presidents' Day
  "2023-03-31", // Cesar Chavez Day
  "2023-05-29", // Memorial Day
  "2023-06-19", // Juneteenth
  "2023-07-04", // Independence Day
  "2023-09-04", // Labor Day
  "2023-11-10", // Veterans Day (observed)
  "2023-11-23", // Thanksgiving
  "2023-11-24", // Day after Thanksgiving
  "2023-12-25", // Christmas Day
  "2023-12-26", // Winter Holiday
  "2023-12-27", // Winter Curtailment
  "2023-12-28", // Winter Curtailment
  "2023-12-29", // Winter Curtailment

  /* 2024 */
  "2024-01-01", // New Year Holiday
  "2024-01-02", // Winter Curtailment (some campuses)
  "2024-01-15", // Martin Luther King Jr. Day
  "2024-02-19", // Presidents' Day
  "2024-03-29", // Cesar Chavez (observed; Mar 31 was Sunday)
  "2024-05-27", // Memorial Day
  "2024-06-19", // Juneteenth
  "2024-07-04", // Independence Day
  "2024-09-02", // Labor Day
  "2024-11-11", // Veterans Day
  "2024-11-28", // Thanksgiving
  "2024-11-29", // Day after Thanksgiving
  "2024-12-24", // Winter Holiday
  "2024-12-25", // Christmas Day
  "2024-12-26", // Winter Curtailment
  "2024-12-27", // Winter Curtailment
  "2024-12-30", // Winter Curtailment
  "2024-12-31", // New Year Holiday (eve)

  /* 2025 */
  "2025-01-01", // New Year Holiday
  "2025-01-02", // Winter Curtailment (some campuses)
  "2025-01-20", // Martin Luther King Jr. Day
  "2025-02-17", // Presidents' Day
  "2025-03-28", // Cesar Chavez (observed)
  "2025-05-26", // Memorial Day
  "2025-06-19", // Juneteenth
  "2025-07-04", // Independence Day
  "2025-09-01", // Labor Day
  "2025-11-11", // Veterans Day
  "2025-11-27", // Thanksgiving
  "2025-11-28", // Day after Thanksgiving
  "2025-12-22", // Winter Curtailment (UCSC, UCB, others)
  "2025-12-23", // Winter Curtailment (UCSC, UCB, others)
  "2025-12-24", // Winter Holiday
  "2025-12-25", // Christmas Day
  "2025-12-26", // Winter Curtailment
  "2025-12-29", // Winter Curtailment
  "2025-12-30", // Winter Curtailment
  "2025-12-31", // New Year Eve Holiday

  /* 2026 — UCOP authoritative */
  "2026-01-01", // New Year Holiday
  "2026-01-02", // Winter Curtailment extension (UCSC, others)
  "2026-01-19", // Martin Luther King Jr. Day
  "2026-02-16", // Presidents' Day
  "2026-03-27", // Cesar Chavez Holiday
  "2026-05-25", // Memorial Day
  "2026-06-19", // Juneteenth
  "2026-07-03", // Independence Day (observed; Jul 4 is Saturday)
  "2026-09-07", // Labor Day
  "2026-11-11", // Veterans Day
  "2026-11-26", // Thanksgiving
  "2026-11-27", // Day after Thanksgiving
  "2026-12-24", // Winter Holiday
  "2026-12-25", // Christmas Day
  "2026-12-28", // Winter Curtailment
  "2026-12-29", // Winter Curtailment
  "2026-12-30", // Winter Curtailment
  "2026-12-31", // New Year Holiday

  /* 2027 — projected using standard UC observation rules */
  "2027-01-01", // New Year Holiday
  "2027-01-18", // Martin Luther King Jr. Day
  "2027-02-15", // Presidents' Day
  "2027-03-31", // Cesar Chavez Day
  "2027-05-31", // Memorial Day
  "2027-06-18", // Juneteenth (observed; Jun 19 is Saturday)
  "2027-07-05", // Independence Day (observed; Jul 4 is Sunday)
  "2027-09-06", // Labor Day
  "2027-11-11", // Veterans Day
  "2027-11-25", // Thanksgiving
  "2027-11-26", // Day after Thanksgiving
  "2027-12-24"  // Christmas Day (observed; Dec 25 is Saturday)
]);

/* ============================================================
   Date utilities — ISO strings throughout to avoid TZ bugs
   ============================================================ */
const ISO = d => {
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,'0');
  const day = String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${day}`;
};
const parseISO = s => {
  if (!s) return null;
  const [y,m,d] = s.split('-').map(Number);
  return new Date(y, m-1, d);
};
const addDays = (d, n) => {
  const r = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  r.setDate(r.getDate() + n);
  return r;
};
const isWeekend = d => { const w = d.getDay(); return w === 0 || w === 6; };
const isHoliday = d => HOLIDAYS.has(ISO(d));
const isBusinessDay = d => !isWeekend(d) && !isHoliday(d);

/* Excel WORKDAY.INTL(start, n, 1, holidays): returns the date that is n
   working days after start, skipping weekends and holidays. n may be 0. */
const addWorkdays = (start, n) => {
  if (n <= 0) return new Date(start.getFullYear(), start.getMonth(), start.getDate());
  let d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  let count = 0;
  while (count < n) {
    d = addDays(d, 1);
    if (isBusinessDay(d)) count++;
  }
  return d;
};

/* Excel NETWORKDAYS.INTL(start, end, 1, holidays): inclusive count of
   business days between two dates. */
const networkdays = (start, end) => {
  if (!start || !end || end < start) return 0;
  let count = 0;
  let d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const stop = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  while (d <= stop) {
    if (isBusinessDay(d)) count++;
    d = addDays(d, 1);
  }
  return count;
};

const fmtLong = d => d ? d.toLocaleDateString(undefined, { weekday:'short', month:'short', day:'numeric', year:'numeric' }) : '';
const fmtShort = d => d ? d.toLocaleDateString(undefined, { month:'short', day:'numeric', year:'numeric' }) : '';
const fmtMonthYear = d => d.toLocaleDateString(undefined, { month:'long', year:'numeric' });

/* ============================================================
   Core calculator — ports the Excel formulas into plain JS
   ============================================================ */
function calculate(input) {
  const {
    lastDay, dueDate, actualBirth, returnDate,
    deliveryType, sickHours, vacHours,
    waitingPeriodDays, fmlEligible, pfcbWeeks, pfcbStart,
    cclWeeks, cclAnchor,
    scheduleType, hoursPerDay, daysPerWeek, fallbackStrategy
  } = input;

  /* --- Schedule-derived constants ---
     Regular employees (5×8): 8 hr/day, 5 days/week, sick cap 22 workdays.
     Variable employees (e.g. 3×12): user-specified hr/day and days/week,
     sick cap 30 workdays.
     waitingWorkdays = working days that fall within the 14-day calendar
     waiting period for this schedule — 10 for regular, ~6 for 3×12. */
  const isVariable = scheduleType === 'variable';
  const effHrsPerDay = isVariable && hoursPerDay > 0 ? hoursPerDay : 8;
  const effDaysPerWeek = isVariable && daysPerWeek > 0 ? daysPerWeek : 5;
  const maxSickCap = isVariable ? 30 : 22;
  const waitingWorkdays = Math.ceil(14 * effDaysPerWeek / 7);

  const sickDaysRaw = Math.floor((sickHours || 0) / effHrsPerDay);
  const vacDays     = Math.floor((vacHours  || 0) / effHrsPerDay);
  const sickDays    = Math.min(sickDaysRaw, maxSickCap);
  const sickCapped  = sickDaysRaw > maxSickCap;

  /* --- SICK leave span --- */
  let sickBegin = null, sickEnd = null;
  if (sickDays >= 1) {
    sickBegin = addWorkdays(lastDay, 1);
    sickEnd   = addWorkdays(lastDay, sickDays);
  }

  /* --- VACATION span ---
     Vacation is used only when all three apply:
       1. The user has vacation hours
       2. Sick leave doesn't cover the waiting period (sick < waitingWorkdays)
       3. The user elected "Use vacation" as the fallback strategy. */
  let vacBegin = null, vacEnd = null, vacNote = '';
  const sickCoversWaiting = sickDays >= waitingWorkdays;
  const useVacation = vacDays >= 1 && !sickCoversWaiting && fallbackStrategy === 'vacation';

  if (useVacation) {
    vacBegin = sickDays === 0
      ? addWorkdays(lastDay, 1)
      : addWorkdays(sickEnd, 1);
    vacEnd = addWorkdays(vacBegin, vacDays - 1);
  } else if (vacDays >= 1 && sickCoversWaiting) {
    vacNote = 'Using vacation is not necessary — sick leave covers the waiting period';
  } else if (vacDays >= 1 && fallbackStrategy === 'lns') {
    vacNote = 'Vacation available but not used (elected to go without pay)';
  }

  /* --- WAITING PERIOD ---
     Calendar-based 14-day (or plan-selected) window starting the day after
     last-day-worked. If sick covers or exceeds the waiting-period workdays,
     the effective "still-waiting-for-disability-pay" window extends until
     sick runs out (capped at maxSickCap). */
  const waitBegin = addDays(lastDay, 1);
  let waitEnd;
  if (sickDays <= waitingWorkdays) {
    waitEnd = addDays(waitBegin, (waitingPeriodDays || 14) - 1);
  } else {
    waitEnd = sickEnd;
  }

  /* --- Lincoln Financial CLAIM FILE date (row 53): 28 days before lastDay --- */
  const fileClaim = addDays(lastDay, -28);

  /* --- PDL (row 62): 42 natural / 56 C-section.
         Starts day after lastDay. Ends from (actualBirth || dueDate). --- */
  const pdlBegin = addDays(lastDay, 1);
  const pdlDurationDays = deliveryType === 'C-section' ? 56 : 42;
  const pdlAnchor = actualBirth || dueDate;
  const pdlEnd = addDays(pdlAnchor, pdlDurationDays);

  /* --- Lincoln Financial income (row 59):
         Begins day after waiting period ends, ends when PDL ends.
         If PDL ends before waiting period completes, income never pays. --- */
  let lincBegin = null, lincEnd = null, lincNote = '';
  if (pdlEnd < waitEnd) {
    lincNote = 'Disability ends before benefit pays';
  } else {
    lincBegin = addDays(waitEnd, 1);
    lincEnd   = pdlEnd;
  }

  /* --- FML (row 65): 84 days from FML begin (= pdlBegin).
         Capped at Dec 31 of start year if it would cross year boundary. --- */
  let fmlBegin = null, fmlEnd = null, fmlCapped = false;
  let fmlNewYearBegin = null, fmlNewYearEnd = null;
  if (fmlEligible) {
    fmlBegin = pdlBegin;
    const naive = addDays(fmlBegin, 83); // 84 days inclusive
    if (naive.getFullYear() > fmlBegin.getFullYear()) {
      fmlCapped = true;
      fmlEnd = new Date(fmlBegin.getFullYear(), 11, 31);
      fmlNewYearBegin = new Date(fmlBegin.getFullYear()+1, 0, 1);
      fmlNewYearEnd   = addDays(fmlNewYearBegin, 83);
    } else {
      fmlEnd = naive;
    }
  }

  /* --- CFRA (row 71): 84 days starting day after PDL ends.
         Only if FML eligible. --- */
  let cfraBegin = null, cfraEnd = null;
  if (fmlEligible) {
    cfraBegin = addDays(pdlEnd, 1);
    cfraEnd   = addDays(cfraBegin, 83);
  }

  /* --- PFCB (row 74) ---
     If the user entered a number of weeks but no start date, default the
     start to the day after PDL ends (the natural bonding window beginning). */
  let pfcbStartResolved = pfcbStart || null;
  let pfcbEnd = null;
  let pfcbStartInferred = false;
  if (pfcbWeeks > 0) {
    if (!pfcbStartResolved && pdlEnd) {
      pfcbStartResolved = addDays(pdlEnd, 1);
      pfcbStartInferred = true;
    }
    if (pfcbStartResolved) {
      pfcbEnd = addDays(pfcbStartResolved, pfcbWeeks * 7 - 1);
    }
  }

  /* --- CCL (Child Caring Leave) ---
     Up to 12 weeks of unpaid leave that begins after the chosen anchor stage
     ends. The anchor varies by campus — some start CCL after PDL ends, some
     after FML, some after CFRA. If the selected anchor isn't available in
     this scenario (e.g. user picked "after FML" but isn't FML-eligible), we
     fall back to the next available anchor in the list and note the change. */
  let cclBegin = null, cclEnd = null, cclAnchorUsed = null;
  const cclAnchorRequested = cclAnchor;
  if (cclWeeks > 0) {
    const available = {
      cfra: cfraEnd,
      fml:  fmlNewYearEnd || fmlEnd,
      pdl:  pdlEnd
    };
    const fallbackOrder = {
      cfra: ['cfra', 'fml', 'pdl'],
      fml:  ['fml', 'pdl'],
      pdl:  ['pdl']
    };
    const tryList = fallbackOrder[cclAnchor] || ['pdl'];
    for (const key of tryList) {
      if (available[key]) { cclAnchorUsed = key; break; }
    }
    const anchorDate = cclAnchorUsed ? available[cclAnchorUsed] : null;
    if (anchorDate) {
      cclBegin = addDays(anchorDate, 1);
      cclEnd   = addDays(cclBegin, cclWeeks * 7 - 1);
    }
  }

  /* --- End of PIE (L30): 31 days after actual birth --- */
  const endPIE = actualBirth ? addDays(actualBirth, 31) : null;
  /* --- Baby's first birthday: same calendar date, one year later.
     Using setFullYear(+1) rather than addDays(365/366) because the
     literal number of days between two yearly anniversaries is 365 in
     non-leap spans and 366 across a Feb 29, and setFullYear returns
     the true calendar anniversary in both cases. --- */
  let firstBday = null;
  if (actualBirth) {
    firstBday = new Date(
      actualBirth.getFullYear() + 1,
      actualBirth.getMonth(),
      actualBirth.getDate()
    );
  }

  return {
    sickDays, sickDaysRaw, sickCapped, maxSickCap, vacDays, vacNote, lincNote, fmlCapped,
    scheduleType, effHrsPerDay, effDaysPerWeek, waitingWorkdays, fallbackStrategy,
    sickBegin, sickEnd, vacBegin, vacEnd,
    waitBegin, waitEnd,
    fileClaim,
    pdlBegin, pdlEnd,
    lincBegin, lincEnd,
    fmlBegin, fmlEnd, fmlNewYearBegin, fmlNewYearEnd,
    cfraBegin, cfraEnd,
    pfcbStart: pfcbStartResolved, pfcbEnd, pfcbWeeks, pfcbStartInferred,
    cclBegin, cclEnd, cclWeeks, cclAnchorRequested, cclAnchorUsed,
    endPIE, firstBday,
    lastDay, dueDate, actualBirth, returnDate
  };
}

/* ============================================================
   Render helpers
   ============================================================ */
function durationDays(start, end) {
  if (!start || !end) return '';
  const ms = end - start;
  const days = Math.round(ms / (1000*60*60*24)) + 1;
  return days + ' day' + (days === 1 ? '' : 's');
}

/* Tiny DOM helpers used by renderTimeline and renderSummary. Keeping
   these generic means we never build HTML strings from dynamic values —
   which eliminates the class of bugs where a stray angle bracket in
   user data becomes an injection sink. textContent is inherently safe:
   whatever string you give it is rendered as literal text. */
function createEl(tag, className, text) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text != null && text !== '') el.textContent = String(text);
  return el;
}
function clearChildren(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

function renderTimeline(r) {
  const list = document.getElementById('timelineList');
  clearChildren(list);

  /* Push a row with begin/optional-end dates. */
  const push = (cls, label, meta, begin, end) => {
    if (!begin && !end) return;
    const datesText = begin && end && +begin !== +end
      ? fmtShort(begin) + ' → ' + fmtShort(end)
      : fmtShort(begin || end);
    const durText = (begin && end) ? durationDays(begin, end) : '';

    const li = createEl('li', 'cat-' + cls);
    li.appendChild(createEl('span', 'label', label));

    const datesSpan = createEl('span', 'dates', datesText);
    if (durText) datesSpan.appendChild(createEl('span', 'dur', durText));
    li.appendChild(datesSpan);

    if (meta) li.appendChild(createEl('span', 'meta', meta));
    list.appendChild(li);
  };

  /* Dateless note card — used for "vacation not necessary", "LNS gap", etc. */
  const pushNote = (cls, label, meta) => {
    const li = createEl('li', 'cat-' + cls);
    li.appendChild(createEl('span', 'label', label));
    li.appendChild(createEl('span', 'meta', meta));
    list.appendChild(li);
  };

  push('milestone', 'Last day worked', '', r.lastDay, null);
  if (r.fileClaim) push('linc', 'File Lincoln Financial claim', '28 days before last day worked', r.fileClaim, null);
  push('milestone', 'Estimated due date', '', r.dueDate, null);
  if (r.actualBirth) push('milestone', 'Actual birth date', '', r.actualBirth, null);

  if (r.sickBegin) {
    let sickMeta = r.sickDays + ' day' + (r.sickDays===1?'':'s') + ' used';
    if (r.sickCapped) {
      sickMeta += ' (capped at ' + r.maxSickCap + ' — you have ' + r.sickDaysRaw + ' total)';
    }
    push('sick', 'Sick leave', sickMeta, r.sickBegin, r.sickEnd);
  }
  if (r.vacBegin) push('vac', 'Vacation leave', r.vacDays + ' day' + (r.vacDays===1?'':'s') + ' used', r.vacBegin, r.vacEnd);
  else if (r.vacNote) pushNote('vac', 'Vacation leave', r.vacNote);

  if (r.sickDays < r.waitingWorkdays && r.fallbackStrategy === 'lns') {
    const gapDays = r.waitingWorkdays - r.sickDays;
    pushNote('wait', 'Leave without pay (waiting-period gap)',
      'Sick covers ' + r.sickDays + ' of ' + r.waitingWorkdays + ' waiting-period working days. ' +
      gapDays + ' working day' + (gapDays===1?'':'s') + ' will be unpaid.');
  }

  push('wait', 'Disability waiting period', '', r.waitBegin, r.waitEnd);

  if (r.lincBegin) push('linc', 'Lincoln Financial disability income', '', r.lincBegin, r.lincEnd);
  else if (r.lincNote) pushNote('linc', 'Lincoln Financial disability income', r.lincNote);

  push('pdl', 'Pregnancy Disability Leave (PDL)',
       r.pdlEnd && r.actualBirth ? 'Anchored to actual birth date' : 'Anchored to estimated due date',
       r.pdlBegin, r.pdlEnd);

  if (r.fmlBegin) {
    push('fml', 'Family Medical Leave (FML)',
         r.fmlCapped ? 'Capped at calendar year end — balance carries over' : '',
         r.fmlBegin, r.fmlEnd);
    if (r.fmlNewYearBegin) push('fml', 'FML — new calendar year', '', r.fmlNewYearBegin, r.fmlNewYearEnd);
  }

  if (r.cfraBegin) push('cfra', 'California Family Rights Act (CFRA)', '', r.cfraBegin, r.cfraEnd);

  if (r.pfcbStart) {
    const pfcbMeta = r.pfcbWeeks + ' week' + (r.pfcbWeeks === 1 ? '' : 's')
      + (r.pfcbStartInferred ? ' · starts day after PDL ends (default)' : '');
    push('pfcb', 'Pay for Family Care and Bonding (PFCB)', pfcbMeta, r.pfcbStart, r.pfcbEnd);
  }

  if (r.cclBegin) {
    const anchorLabel = { pdl: 'PDL', fml: 'FML', cfra: 'CFRA' }[r.cclAnchorUsed];
    const fellBack = r.cclAnchorUsed !== r.cclAnchorRequested;
    const requestedLabel = { pdl: 'PDL', fml: 'FML', cfra: 'CFRA' }[r.cclAnchorRequested];
    let cclMeta = r.cclWeeks + ' week' + (r.cclWeeks === 1 ? '' : 's')
      + ' · starts day after ' + anchorLabel + ' ends';
    if (fellBack) {
      cclMeta += ' (fallback — ' + requestedLabel + ' not available in this scenario)';
    }
    push('ccl', 'Child Caring Leave (CCL)', cclMeta, r.cclBegin, r.cclEnd);
  } else if (r.cclWeeks > 0) {
    pushNote('ccl', 'Child Caring Leave (CCL)',
      r.cclWeeks + ' week' + (r.cclWeeks === 1 ? '' : 's')
      + ' requested — no anchor date available. Enter at least a last day worked and due date.');
  }

  if (r.returnDate) push('milestone', 'Estimated return to work', '', r.returnDate, null);
  if (r.endPIE) push('milestone', 'End of PIE (enroll baby by)', '31 days after birth', r.endPIE, null);
  if (r.firstBday) push('milestone', "Baby's first birthday", 'End of parental bonding window', r.firstBday, null);
}

function renderSummary(r) {
  const container = document.getElementById('summaryCallout');
  clearChildren(container);

  let isFirst = true;
  const addLine = (label, body) => {
    if (!isFirst) container.appendChild(document.createTextNode(' '));
    isFirst = false;
    const strong = createEl('strong', null, label + ':');
    container.appendChild(strong);
    container.appendChild(document.createTextNode(' ' + body));
  };

  addLine('Pregnancy Disability Leave', fmtShort(r.pdlBegin) + ' → ' + fmtShort(r.pdlEnd) + '.');
  if (r.fmlBegin) {
    addLine('FML', fmtShort(r.fmlBegin) + ' → ' + fmtShort(r.fmlEnd)
      + (r.fmlCapped ? ' (calendar year cap)' : '') + '.');
  }
  if (r.cfraBegin) {
    addLine('CFRA', fmtShort(r.cfraBegin) + ' → ' + fmtShort(r.cfraEnd) + '.');
  }
  if (r.lincBegin) {
    addLine('Lincoln Financial income', fmtShort(r.lincBegin) + ' → ' + fmtShort(r.lincEnd) + '.');
  }
  if (r.pfcbStart) {
    addLine('PFCB', fmtShort(r.pfcbStart) + ' → ' + fmtShort(r.pfcbEnd)
      + ' (' + r.pfcbWeeks + ' week' + (r.pfcbWeeks === 1 ? '' : 's') + ').');
  }
  if (r.cclBegin) {
    const anchorLabel = { pdl: 'PDL', fml: 'FML', cfra: 'CFRA' }[r.cclAnchorUsed];
    addLine('CCL', fmtShort(r.cclBegin) + ' → ' + fmtShort(r.cclEnd)
      + ' (' + r.cclWeeks + ' week' + (r.cclWeeks === 1 ? '' : 's')
      + ', after ' + anchorLabel + ').');
  }
}

/* ============================================================
   Calendar rendering
   ============================================================ */
function buildEventIndex(r) {
  /* Map ISO date → array of {type, label}.
     Precedence (visual): category background (sick/vac/lns), then chips. */
  const idx = new Map();
  const add = (d, type, label) => {
    if (!d) return;
    const key = ISO(d);
    if (!idx.has(key)) idx.set(key, { chips: [], cats: new Set(), labels: [] });
    const bucket = idx.get(key);
    if (type === 'cat') bucket.cats.add(label);
    else bucket.chips.push({ type, label });
    bucket.labels.push(label);
  };

  /* categories — fill in the ranges */
  const fillRange = (begin, end, cat) => {
    if (!begin || !end) return;
    let d = new Date(begin);
    while (d <= end) {
      if (isBusinessDay(d)) add(new Date(d), 'cat', cat);
      d = addDays(d, 1);
    }
  };
  if (r.sickBegin) fillRange(r.sickBegin, r.sickEnd, 'sick');
  if (r.vacBegin)  fillRange(r.vacBegin,  r.vacEnd,  'vac');
  /* leave-no-salary = within waiting period but outside sick/vac */
  if (r.waitBegin) {
    let d = new Date(r.waitBegin);
    while (d <= r.waitEnd) {
      if (isBusinessDay(d)) {
        const key = ISO(d);
        const ex = idx.get(key);
        if (!ex || (!ex.cats.has('sick') && !ex.cats.has('vac'))) {
          add(new Date(d), 'cat', 'lns');
        }
      }
      d = addDays(d, 1);
    }
  }

  /* event chips — only at begin/end dates to keep calendar readable */
  const mark = (d, type, label) => add(d, type, label);
  mark(r.lastDay, 'milestone', 'Last day worked');
  mark(r.dueDate, 'milestone', 'Est due date');
  mark(r.actualBirth, 'milestone', 'Actual birth');
  mark(r.returnDate, 'milestone', 'Return to work');
  mark(r.endPIE, 'milestone', 'End of PIE');
  mark(r.firstBday, 'milestone', 'Baby 1st bday');
  mark(r.fileClaim, 'linc', 'File claim');

  mark(r.waitBegin, 'wait', 'Waiting begins');
  mark(r.waitEnd, 'wait', 'Waiting ends');

  mark(r.lincBegin, 'linc', 'Lincoln begins');
  mark(r.lincEnd, 'linc', 'Lincoln ends');

  mark(r.pdlBegin, 'pdl', 'PDL begins');
  mark(r.pdlEnd, 'pdl', 'PDL ends');

  mark(r.fmlBegin, 'fml', 'FML begins');
  mark(r.fmlEnd, 'fml', r.fmlCapped ? 'FML ends (CY)' : 'FML ends');
  mark(r.fmlNewYearBegin, 'fml', 'FML resumes');
  mark(r.fmlNewYearEnd, 'fml', 'FML ends');

  mark(r.cfraBegin, 'cfra', 'CFRA begins');
  mark(r.cfraEnd, 'cfra', 'CFRA ends');

  mark(r.pfcbStart, 'pfcb', 'PFCB begins');
  mark(r.pfcbEnd, 'pfcb', 'PFCB ends');

  mark(r.cclBegin, 'ccl', 'CCL begins');
  mark(r.cclEnd, 'ccl', 'CCL ends');

  return idx;
}

function renderCalendars(r) {
  const wrap = document.getElementById('calendars');
  clearChildren(wrap);

  /* Determine range: one month before lastDay → through latest event */
  const anchor = r.lastDay;
  const start  = new Date(anchor.getFullYear(), anchor.getMonth() - 1, 1);
  const candidates = [
    r.pdlEnd, r.fmlEnd, r.fmlNewYearEnd, r.cfraEnd, r.pfcbEnd, r.cclEnd,
    r.endPIE, r.firstBday, r.lincEnd, r.returnDate
  ].filter(Boolean);
  let maxDate = candidates.reduce((a,b) => (a > b ? a : b), anchor);
  const end = new Date(maxDate.getFullYear(), maxDate.getMonth() + 1, 0);

  const events = buildEventIndex(r);

  let cursor = new Date(start.getFullYear(), start.getMonth(), 1);
  const endMonth = new Date(end.getFullYear(), end.getMonth(), 1);

  while (cursor <= endMonth) {
    wrap.appendChild(buildMonth(cursor, events));
    cursor = new Date(cursor.getFullYear(), cursor.getMonth()+1, 1);
  }
}

function buildMonth(monthDate, events) {
  const container = document.createElement('section');
  container.className = 'cal';

  const table = document.createElement('table');
  table.setAttribute('role', 'table');
  const caption = document.createElement('caption');
  caption.textContent = fmtMonthYear(monthDate);
  table.appendChild(caption);

  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].forEach(d => {
    const th = document.createElement('th');
    th.scope = 'col'; th.textContent = d;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const year = monthDate.getFullYear();
  const month = monthDate.getMonth();
  const firstWeekday = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month+1, 0).getDate();

  let tr = document.createElement('tr');
  /* leading empties */
  for (let i = 0; i < firstWeekday; i++) {
    const td = document.createElement('td');
    td.className = 'empty';
    tr.appendChild(td);
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month, d);
    const key = ISO(date);
    const td = document.createElement('td');
    td.className = 'day';
    if (isWeekend(date)) td.classList.add('weekend');
    if (isHoliday(date)) td.classList.add('holiday');

    const ev = events.get(key);
    if (ev) {
      if (ev.cats.has('sick')) td.classList.add('c-sick');
      else if (ev.cats.has('vac')) td.classList.add('c-vac');
      else if (ev.cats.has('lns')) td.classList.add('c-lns');
    }

    const dnum = document.createElement('span');
    dnum.className = 'd';
    dnum.textContent = d;
    td.appendChild(dnum);

    if (ev) {
      /* category label (sick / vac) as small text */
      if (ev.cats.has('sick')) {
        const cat = document.createElement('span');
        cat.className = 'cat'; cat.textContent = 'SICK';
        td.appendChild(cat);
      } else if (ev.cats.has('vac')) {
        const cat = document.createElement('span');
        cat.className = 'cat'; cat.textContent = 'VAC';
        td.appendChild(cat);
      } else if (ev.cats.has('lns')) {
        const cat = document.createElement('span');
        cat.className = 'cat'; cat.textContent = 'LNS';
        cat.title = 'Leave no salary';
        td.appendChild(cat);
      }
      /* chips for discrete events */
      ev.chips.forEach(c => {
        const chip = document.createElement('span');
        chip.className = 'chip ' + c.type;
        chip.textContent = c.label;
        td.appendChild(chip);
      });

      /* accessible summary on the cell */
      td.setAttribute('aria-label',
        date.toLocaleDateString(undefined,{weekday:'long',month:'long',day:'numeric',year:'numeric'})
        + '. ' + ev.labels.join('. ') + '.');
    } else if (isHoliday(date)) {
      td.setAttribute('aria-label',
        date.toLocaleDateString(undefined,{weekday:'long',month:'long',day:'numeric',year:'numeric'})
        + '. Campus holiday.');
    }

    tr.appendChild(td);
    if ((firstWeekday + d) % 7 === 0) {
      tbody.appendChild(tr);
      tr = document.createElement('tr');
    }
  }
  /* trailing empties */
  if (tr.children.length > 0) {
    while (tr.children.length < 7) {
      const td = document.createElement('td');
      td.className = 'empty';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  container.appendChild(table);
  return container;
}

/* ============================================================
   Validation + wiring
   ============================================================ */
const showErr = (id, msg) => {
  const el = document.getElementById(id + '-err');
  const field = document.getElementById(id);
  if (msg) {
    el.textContent = msg; el.hidden = false;
    field.setAttribute('aria-invalid','true');
  } else {
    el.textContent = ''; el.hidden = true;
    field.removeAttribute('aria-invalid');
  }
};

function collect() {
  const v = id => document.getElementById(id).value;
  const n = id => { const x = parseFloat(v(id)); return Number.isFinite(x) ? x : 0; };
  const radioValue = name => {
    const el = document.querySelector('input[name="' + name + '"]:checked');
    return el ? el.value : null;
  };
  return {
    lastDay: parseISO(v('lastDay')),
    dueDate: parseISO(v('dueDate')),
    actualBirth: parseISO(v('actualBirth')),
    returnDate: parseISO(v('returnDate')),
    deliveryType: v('deliveryType'),
    sickHours: n('sickHours'),
    vacHours: n('vacHours'),
    waitingPeriodDays: parseInt(v('waitingPeriod'),10) || 14,
    fmlEligible: document.getElementById('fmlEligible').checked,
    pfcbWeeks: n('pfcbWeeks'),
    pfcbStart: parseISO(v('pfcbStart')),
    cclWeeks: n('cclWeeks'),
    cclAnchor: v('cclAnchor') || 'pdl',
    scheduleType: radioValue('scheduleType') || 'regular',
    hoursPerDay: n('hoursPerDay'),
    daysPerWeek: n('daysPerWeek'),
    fallbackStrategy: radioValue('fallbackStrategy') || 'vacation'
  };
}

/* Human-readable labels for the form fields we validate. Keeping these
   in one place means the error summary, inline errors, and any future
   export share the same wording. */
const FIELD_LABELS = {
  lastDay:      'Last day worked',
  dueDate:      'Estimated due date',
  deliveryType: 'Delivery type',
  sickHours:    'Total sick hours',
  hoursPerDay:  'Hours per day',
  daysPerWeek:  'Days per week'
};

const VALIDATED_FIELDS = Object.keys(FIELD_LABELS);

function validate(input) {
  /* Clear all previous inline errors first so resolved issues disappear. */
  VALIDATED_FIELDS.forEach(id => showErr(id, ''));

  /* Accumulator: each entry is { fieldId, message }. The order here
     determines the display order in the summary. */
  const errors = [];
  const addError = (fieldId, message) => {
    errors.push({ fieldId, message });
    showErr(fieldId, message);
  };

  if (!input.lastDay) {
    addError('lastDay', 'Last day worked is required.');
  }
  if (!input.dueDate) {
    addError('dueDate', 'Estimated due date is required.');
  }
  if (!input.deliveryType) {
    addError('deliveryType', 'Please select a delivery type.');
  }

  const sickField = document.getElementById('sickHours');
  if (sickField.value === '' || isNaN(parseFloat(sickField.value))) {
    addError('sickHours', 'Total sick hours is required (enter 0 if none).');
  } else if (input.sickHours < 0) {
    addError('sickHours', 'Sick hours cannot be negative.');
  }

  if (input.lastDay && input.dueDate && input.lastDay > input.dueDate) {
    addError('dueDate', 'Due date should be on or after last day worked.');
  }

  if (input.scheduleType === 'variable') {
    if (!(input.hoursPerDay > 0)) {
      addError('hoursPerDay', 'Enter hours per day (greater than 0).');
    }
    if (!(input.daysPerWeek > 0)) {
      addError('daysPerWeek', 'Enter days per week (1–7).');
    }
  }

  renderErrorSummary(errors);
  return errors.length === 0;
}

/* Render (or clear) the consolidated error summary at the top of the
   form. On show, the summary element gets programmatic focus so screen
   readers read its content — a belt-and-braces backup to role="alert",
   because focus-based announcement is the most reliable mechanism across
   AT combinations. Also routes a count message through the polite
   status channel for redundancy. */
function renderErrorSummary(errors) {
  const summary = document.getElementById('errorSummary');
  const list = document.getElementById('errorSummaryList');
  if (!summary || !list) return;

  clearChildren(list);

  if (errors.length === 0) {
    summary.hidden = true;
    return;
  }

  errors.forEach(({ fieldId, message }) => {
    const label = FIELD_LABELS[fieldId] || fieldId;
    const li = document.createElement('li');
    const a = document.createElement('a');
    a.href = '#' + fieldId;
    a.textContent = label + ': ' + message;
    /* Clicking a summary link should jump to AND focus the field. Native
       anchor behavior handles scroll for focusable inputs, but we call
       focus() explicitly for consistency across browsers and to ensure
       the field's accessible name + error are announced by screen
       readers on arrival. */
    a.addEventListener('click', function (e) {
      const target = document.getElementById(fieldId);
      if (target) {
        e.preventDefault();
        target.focus();
        if (typeof target.scrollIntoView === 'function') {
          target.scrollIntoView({ block: 'center', behavior: 'smooth' });
        }
      }
    });
    li.appendChild(a);
    list.appendChild(li);
  });

  summary.hidden = false;
  /* Defer focus to next tick so the element is fully rendered and the
     role="alert" has had a chance to register as a live region. */
  setTimeout(() => { summary.focus(); }, 0);

  const countMsg = errors.length === 1
    ? 'There is 1 problem with your information. Please review the error summary.'
    : 'There are ' + errors.length + ' problems with your information. Please review the error summary.';
  liveSay(countMsg);
}

const liveSay = msg => {
  document.getElementById('statusLive').textContent = msg;
};

function run(e) {
  if (e && typeof e.preventDefault === 'function') e.preventDefault();
  try {
    const input = collect();
    if (!validate(input)) {
      /* Validation failed. renderErrorSummary (called from validate)
         already populated the summary, announced the count through the
         polite live region, and moved focus to the summary itself so
         the user's screen reader reads the heading and error list. No
         further action needed here. */
      return false;
    }
    const r = calculate(input);
    renderSummary(r);
    renderTimeline(r);
    renderCalendars(r);
    document.getElementById('resultsEmpty').hidden = true;
    document.getElementById('resultsContent').hidden = false;
    liveSay(
      'Timeline updated. Pregnancy Disability Leave runs from '
      + fmtShort(r.pdlBegin) + ' to ' + fmtShort(r.pdlEnd) + '.'
    );
  } catch (err) {
    console.error('Calculation failed:', err);
    liveSay('Something went wrong while calculating. Check the browser console for details.');
    alert('Calculation error: ' + err.message);
  }
  return false;
}

function resetAll() {
  document.getElementById('calcForm').reset();
  document.querySelectorAll('[aria-invalid="true"]').forEach(el => el.removeAttribute('aria-invalid'));
  document.querySelectorAll('.err-msg').forEach(el => { el.textContent=''; el.hidden = true; });
  /* Also clear the consolidated error summary. */
  const summary = document.getElementById('errorSummary');
  if (summary) {
    summary.hidden = true;
    clearChildren(document.getElementById('errorSummaryList'));
  }
  document.getElementById('resultsEmpty').hidden = false;
  document.getElementById('resultsContent').hidden = true;
  liveSay('Form cleared.');
  document.getElementById('lastDay').focus();
}

/* ============================================================
   Theme management (dark / light mode)
   -----------------------------------------------------------
   Resolution order:
     1. Stored user preference in localStorage
     2. OS-level prefers-color-scheme
   An inline script in <head> applies the initial theme before
   CSS loads, so there is no flash. This module wires the toggle
   button and listens for OS-level preference changes (which only
   take effect when the user has not made an explicit choice).
   ============================================================ */
const THEME_KEY = 'uc-leave-calc-theme';

function readStoredTheme() {
  try {
    const v = localStorage.getItem(THEME_KEY);
    return (v === 'light' || v === 'dark') ? v : null;
  } catch (e) { return null; }
}
function writeStoredTheme(t) {
  try { localStorage.setItem(THEME_KEY, t); } catch (e) {}
}
function systemPrefersDark() {
  return window.matchMedia &&
         window.matchMedia('(prefers-color-scheme: dark)').matches;
}
function resolveTheme() {
  return readStoredTheme() || (systemPrefersDark() ? 'dark' : 'light');
}
function applyTheme(t) {
  document.documentElement.setAttribute('data-theme', t);
  const btn = document.getElementById('themeToggle');
  if (!btn) return;
  const other = t === 'dark' ? 'light' : 'dark';
  const otherLabel = other.charAt(0).toUpperCase() + other.slice(1) + ' mode';
  btn.setAttribute('aria-pressed', t === 'dark' ? 'true' : 'false');
  btn.setAttribute('aria-label', 'Switch to ' + other + ' mode');
  const label = btn.querySelector('.theme-toggle-label');
  if (label) label.textContent = otherLabel;
}
function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || resolveTheme();
  const next = current === 'dark' ? 'light' : 'dark';
  writeStoredTheme(next);
  applyTheme(next);
  liveSay(next === 'dark' ? 'Switched to dark mode.' : 'Switched to light mode.');
}

/* Show/hide the variable-schedule sub-fields based on which radio is
   selected. Runs on every change of the scheduleType radios. */
function wireScheduleToggle() {
  const radios = document.querySelectorAll('input[name="scheduleType"]');
  const sub = document.getElementById('variableFields');
  if (!radios.length || !sub) return;
  const update = () => {
    const selected = document.querySelector('input[name="scheduleType"]:checked');
    sub.hidden = !selected || selected.value !== 'variable';
  };
  radios.forEach(r => r.addEventListener('change', update));
  update();
}

/* Wire any element with class .info-btn and aria-controls pointing at a
   disclosure region. Click toggles aria-expanded and the target's hidden
   attribute. Escape key on the button collapses the note.
   On expand, the revealed content is announced through the polite live
   region in addition to the aria-live attribute on the note itself —
   this belt-and-braces approach covers screen readers that don't
   reliably announce hidden→visible transitions on live regions. */
function wireInfoButtons() {
  document.querySelectorAll('.info-btn[aria-controls]').forEach(btn => {
    const targetId = btn.getAttribute('aria-controls');
    const target = document.getElementById(targetId);
    if (!target) return;

    const announceContent = () => {
      const text = (target.textContent || '').replace(/\s+/g, ' ').trim();
      if (text) liveSay(text);
    };

    btn.addEventListener('click', function () {
      const expanded = btn.getAttribute('aria-expanded') === 'true';
      btn.setAttribute('aria-expanded', expanded ? 'false' : 'true');
      target.hidden = expanded;
      if (!expanded) announceContent();
    });
    btn.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && btn.getAttribute('aria-expanded') === 'true') {
        btn.setAttribute('aria-expanded', 'false');
        target.hidden = true;
      }
    });
  });
}

/* Announce dropdown (select) changes through the polite live region.
   aria-live on the <select> itself (set in the HTML) gives modern screen
   readers a hint, and this handler supplements it with a clear,
   human-readable confirmation of what was chosen. The announcement is
   "polite" so it never interrupts other screen-reader speech. */
function wireSelectAnnouncements() {
  const labels = {
    deliveryType:   'Delivery type',
    waitingPeriod:  'Waiting period',
    cclAnchor:      'CCL starts after'
  };
  Object.keys(labels).forEach(id => {
    const el = document.getElementById(id);
    if (!el || el.tagName !== 'SELECT') return;
    el.addEventListener('change', function () {
      const opt = el.options[el.selectedIndex];
      const valueText = opt ? opt.text : el.value;
      liveSay(labels[id] + ' set to: ' + valueText + '.');
    });
  });
}

/* Wire handlers once DOM is ready. Uses both a submit handler (the correct
   semantic) and a click handler on the button (belt-and-braces fallback in
   case some browser quirk or embedded preview swallows the submit event). */
function wire() {
  const form = document.getElementById('calcForm');
  const calcBtn = document.querySelector('button.btn-primary');
  const resetBtn = document.getElementById('resetBtn');
  const printBtn = document.getElementById('printBtn');
  const themeBtn = document.getElementById('themeToggle');

  if (form) form.addEventListener('submit', run);
  if (calcBtn) calcBtn.addEventListener('click', run);
  if (resetBtn) resetBtn.addEventListener('click', resetAll);
  if (printBtn) printBtn.addEventListener('click', function () { window.print(); });
  if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

  wireInfoButtons();
  wireScheduleToggle();
  wireSelectAnnouncements();

  /* Sync the toggle button's aria state with whatever the no-flash
     inline loader picked (light or dark). */
  applyTheme(resolveTheme());

  /* Listen for OS-level theme changes. Only apply when the user has
     not expressed an explicit preference (i.e. nothing in storage). */
  if (window.matchMedia) {
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    const handler = function (e) {
      if (!readStoredTheme()) applyTheme(e.matches ? 'dark' : 'light');
    };
    if (mq.addEventListener) mq.addEventListener('change', handler);
    else if (mq.addListener) mq.addListener(handler); /* legacy Safari */
  }

  /* Enable smooth theme transitions only after the first paint, so
     the initial page load does not animate from a default state. */
  requestAnimationFrame(function () {
    document.body.classList.add('theme-transitions');
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', wire);
} else {
  wire();
}
