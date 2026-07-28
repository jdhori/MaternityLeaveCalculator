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

/* PFCB end date: counts `n` calendar days forward from start (inclusive) but
   does NOT count holidays toward the total. Each holiday inside the window
   pushes the end out by one calendar day, so the employee gets the full 8
   weeks of PFCB without holidays being deducted. Weekends still count. */
const addPfcbDays = (start, n) => {
  let d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  let counted = 0;
  while (true) {
    if (!isHoliday(d)) counted++;
    if (counted >= n) return d;
    d = addDays(d, 1);
  }
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
    waitingPeriodDays, appliesLincoln,
    pdlEligible, fmlEligible, cfraEligible,
    pfcbWeeks, pfcbStart,
    cclWeeks, cclAnchor,
    scheduleType, hoursPerDay, daysPerWeek, fallbackStrategy,
    employeeType
  } = input;
  /* Postdocs follow different plan rules: a 7-day disability waiting period,
     short-term disability (STD) through The Standard instead of Lincoln
     Financial, PPFL instead of PFCB (8 weeks, no FMLA/CFRA eligibility
     required, per birth), and no Child Caring Leave. The PDL span is
     unchanged — 42 days natural / 56 days C-section already matches the
     postdoc 6–8 week guideline. */
  const isPostdoc = employeeType === 'postdoc';
  /* `lastDay` is the field id, but the field semantically represents the
     Leave Start Date — the first day of leave. We keep the variable name
     for diff stability while the meaning is documented here. All downstream
     date math assumes this is day 1 of leave, NOT the last day at work. */
  const leaveStart = lastDay;

  /* --- Schedule-derived constants ---
     Regular employees (5×8): 8 hr/day, 5 days/week, sick cap 22 workdays.
     Variable employees (e.g. 3×12): user-specified hr/day and days/week,
     sick cap 30 workdays.
     waitingWorkdays = working days that fall within the calendar waiting
     period for this schedule — e.g. 10 for regular staff (14-day wait),
     5 for regular postdocs (7-day wait), ~6 for 3×12 staff. */
  const isVariable = scheduleType === 'variable';
  const effHrsPerDay = isVariable && hoursPerDay > 0 ? hoursPerDay : 8;
  const effDaysPerWeek = isVariable && daysPerWeek > 0 ? daysPerWeek : 5;
  const maxSickCap = isVariable ? 30 : 22;
  const effWaitingDays = waitingPeriodDays || (isPostdoc ? 7 : 14);
  const waitingWorkdays = Math.ceil(effWaitingDays * effDaysPerWeek / 7);

  const sickDaysRaw = Math.floor((sickHours || 0) / effHrsPerDay);
  const vacDays     = Math.floor((vacHours  || 0) / effHrsPerDay);
  const sickDays    = Math.min(sickDaysRaw, maxSickCap);
  const sickCapped  = sickDaysRaw > maxSickCap;

  /* --- SICK leave span ---
     Starts on the leave-start date (assumed to be a workday) and runs for
     `sickDays` working days inclusive. addWorkdays(d, 0) returns d unchanged
     if d is itself a workday. */
  let sickBegin = null, sickEnd = null;
  if (sickDays >= 1) {
    sickBegin = addWorkdays(leaveStart, 0);
    sickEnd   = addWorkdays(sickBegin, sickDays - 1);
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
      ? addWorkdays(leaveStart, 0)
      : addWorkdays(sickEnd, 1);
    vacEnd = addWorkdays(vacBegin, vacDays - 1);
  } else if (vacDays >= 1 && sickCoversWaiting) {
    vacNote = 'Using vacation is not necessary — sick leave covers the waiting period';
  } else if (vacDays >= 1 && fallbackStrategy === 'lns') {
    vacNote = 'Vacation available but not used (elected to go without pay)';
  }

  /* --- WAITING PERIOD ---
     Only applies when the employee is applying for disability income through
     their carrier (Lincoln Financial for staff — 14 days; The Standard for
     postdocs — 7 days). Calendar window starting on the leave start date.
     If sick covers or exceeds the waiting-period workdays, the effective
     "still-waiting-for-disability-pay" window extends until sick runs out
     (capped at maxSickCap) — this is also what lets sick leave be used
     during disability leave while no disability benefits are being paid.
     When the employee is not applying for disability, there is no waiting
     period. `appliesLincoln` keeps its historical name but means "applying
     for carrier disability" in both modes. */
  let waitBegin = null, waitEnd = null;
  if (appliesLincoln) {
    waitBegin = leaveStart;
    if (sickDays <= waitingWorkdays) {
      waitEnd = addDays(waitBegin, effWaitingDays - 1);
    } else {
      waitEnd = sickEnd;
    }
  }

  /* --- Lincoln Financial CLAIM FILE date: 28 days before leave start.
         Health LOA recommends 1–2 weeks before; plan ceiling is 30 days.
         Only relevant when applying for Lincoln disability. --- */
  const fileClaim = appliesLincoln ? addDays(leaveStart, -28) : null;

  /* --- PDL (row 62): 42 natural / 56 C-section.
         Starts on leave-start date. Ends `pdlDurationDays` calendar days
         after birth (inclusive of the birth day), so we subtract 1 from
         the addDays offset. Only displayed when the employee marks PDL
         eligibility. --- */
  let pdlBegin = null, pdlEnd = null;
  const pdlDurationDays = deliveryType === 'C-section' ? 56 : 42;
  const pdlAnchor = actualBirth || dueDate;
  if (pdlEligible && pdlAnchor) {
    pdlBegin = leaveStart;
    pdlEnd   = addDays(pdlAnchor, pdlDurationDays - 1);
  }

  /* --- Lincoln Financial income (row 59):
         Begins day after waiting period ends, ends when PDL ends.
         If PDL ends before waiting period completes, income never pays. --- */
  let lincBegin = null, lincEnd = null, lincNote = '';
  if (!appliesLincoln || !pdlEnd) {
    /* Not applying for Lincoln disability, or no PDL window to anchor to. */
    lincNote = '';
  } else if (pdlEnd < waitEnd) {
    lincNote = 'Disability ends before benefit pays';
  } else {
    lincBegin = addDays(waitEnd, 1);
    lincEnd   = pdlEnd;
  }

  /* --- FMLA (row 65): 84 days from FMLA begin (= leave start).
         Capped at Dec 31 of start year if it would cross year boundary. --- */
  let fmlBegin = null, fmlEnd = null, fmlCapped = false;
  let fmlNewYearBegin = null, fmlNewYearEnd = null;
  if (fmlEligible) {
    fmlBegin = leaveStart;
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

  /* --- CFRA (row 71): 84 days starting day after PDL ends (if PDL is in
         play), else day after leave starts (CFRA-only without PDL). --- */
  let cfraBegin = null, cfraEnd = null;
  if (cfraEligible) {
    cfraBegin = pdlEnd ? addDays(pdlEnd, 1) : leaveStart;
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
      /* Holidays inside the PFCB window are not deducted from the 8 weeks —
         each one extends the end date by a day. */
      pfcbEnd = addPfcbDays(pfcbStartResolved, pfcbWeeks * 7);
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
  /* CCL is a staff benefit — postdocs never get a CCL block even if a
     value survived in the hidden field. */
  if (cclWeeks > 0 && !isPostdoc) {
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
    isPostdoc, fmlEligible,
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
    cclBegin, cclEnd, cclWeeks: isPostdoc ? 0 : cclWeeks,
    cclAnchorRequested, cclAnchorUsed,
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
  return days + ' calendar day' + (days === 1 ? '' : 's');
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

  /* Order requested by the service-channel management team: the primary
     leave blocks come first, in this sequence — Leave start date, Disability
     waiting period, Pregnancy Disability Leave, Family & Medical Leave, then
     CFRA. Supporting milestones and pay items follow. */
  push('milestone', 'Leave start date', '', r.lastDay, null);

  push('wait', 'Disability waiting period', '', r.waitBegin, r.waitEnd);

  if (r.pdlBegin) {
    push('pdl', 'Pregnancy Disability Leave (PDL)',
         r.actualBirth ? 'Anchored to actual birth date' : 'Anchored to estimated due date',
         r.pdlBegin, r.pdlEnd);
  }

  if (r.fmlBegin) {
    push('fml', 'Family & Medical Leave Act (FMLA)',
         r.fmlCapped ? 'Capped at calendar year end — balance carries over' : '',
         r.fmlBegin, r.fmlEnd);
    if (r.fmlNewYearBegin) push('fml', 'FMLA — new calendar year', '', r.fmlNewYearBegin, r.fmlNewYearEnd);
  }

  if (r.cfraBegin) push('cfra', 'California Family Rights Act (CFRA)', '', r.cfraBegin, r.cfraEnd);

  /* Carrier naming differs by employee type: staff use Lincoln Financial,
     postdocs apply for short-term disability (STD) through The Standard. */
  const carrierIncomeLabel = r.isPostdoc
    ? 'The Standard short-term disability (STD) income'
    : 'Lincoln Financial disability income';
  if (r.lincBegin) push('linc', carrierIncomeLabel, '', r.lincBegin, r.lincEnd);
  else if (r.lincNote) pushNote('linc', carrierIncomeLabel, r.lincNote);

  if (r.pfcbStart) {
    const bondingLabel = r.isPostdoc
      ? 'Postdoc Paid Family Leave (PPFL)'
      : 'Pay for Family Care and Bonding (PFCB)';
    let pfcbMeta = r.pfcbWeeks + ' week' + (r.pfcbWeeks === 1 ? '' : 's')
      + (r.pfcbStartInferred ? ' · starts day after PDL ends (default)' : '');
    if (r.isPostdoc) {
      pfcbMeta += ' · per birth — cannot be used again in the new year';
      pfcbMeta += r.fmlEligible
        ? ' · sick leave may not be used for pay during family leave (PPFL or vacation only)'
        : ' · personal leave paid via PPFL — no departmental approval needed';
    }
    push('pfcb', bondingLabel, pfcbMeta, r.pfcbStart, r.pfcbEnd);
  }

  if (r.fileClaim) push('linc',
    r.isPostdoc ? 'File STD claim with The Standard' : 'File Lincoln Financial claim',
    'May file up to 30 days before leave begins. Health LOA recommends 1–2 weeks before. Requires medical certification; the LOA team processes the leave.',
    r.fileClaim, null);

  if (r.sickBegin) {
    let sickMeta = r.sickDays + ' calendar day' + (r.sickDays===1?'':'s') + ' used';
    if (r.sickCapped) {
      sickMeta += ' (capped at ' + r.maxSickCap + ' — you have ' + r.sickDaysRaw + ' total)';
    }
    push('sick', 'Sick leave', sickMeta, r.sickBegin, r.sickEnd);
  }
  if (r.vacBegin) push('vac', 'Vacation leave', r.vacDays + ' day' + (r.vacDays===1?'':'s') + ' used', r.vacBegin, r.vacEnd);
  else if (r.vacNote) pushNote('vac', 'Vacation leave', r.vacNote);

  if (r.waitBegin && r.sickDays < r.waitingWorkdays && r.fallbackStrategy === 'lns') {
    const gapDays = r.waitingWorkdays - r.sickDays;
    pushNote('wait', 'Leave without pay (waiting-period gap)',
      'Sick covers ' + r.sickDays + ' of ' + r.waitingWorkdays + ' waiting-period working days. ' +
      gapDays + ' working day' + (gapDays===1?'':'s') + ' will be unpaid.');
  }

  push('milestone', 'Estimated due date', '', r.dueDate, null);
  if (r.actualBirth) push('milestone', 'Actual birth date', '', r.actualBirth, null);

  if (r.cclBegin) {
    const anchorLabel = { pdl: 'PDL', fml: 'FMLA', cfra: 'CFRA' }[r.cclAnchorUsed];
    const fellBack = r.cclAnchorUsed !== r.cclAnchorRequested;
    const requestedLabel = { pdl: 'PDL', fml: 'FMLA', cfra: 'CFRA' }[r.cclAnchorRequested];
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

  if (r.pdlBegin) {
    addLine('Pregnancy Disability Leave', fmtShort(r.pdlBegin) + ' → ' + fmtShort(r.pdlEnd) + '.');
  }
  if (r.fmlBegin) {
    addLine('FMLA', fmtShort(r.fmlBegin) + ' → ' + fmtShort(r.fmlEnd)
      + (r.fmlCapped ? ' (calendar year cap)' : '') + '.');
  }
  if (r.cfraBegin) {
    addLine('CFRA', fmtShort(r.cfraBegin) + ' → ' + fmtShort(r.cfraEnd) + '.');
  }
  if (r.lincBegin) {
    addLine(r.isPostdoc ? 'The Standard STD income' : 'Lincoln Financial income',
      fmtShort(r.lincBegin) + ' → ' + fmtShort(r.lincEnd) + '.');
  }
  if (r.pfcbStart) {
    addLine(r.isPostdoc ? 'PPFL' : 'PFCB', fmtShort(r.pfcbStart) + ' → ' + fmtShort(r.pfcbEnd)
      + ' (' + r.pfcbWeeks + ' week' + (r.pfcbWeeks === 1 ? '' : 's') + ').');
  }
  if (r.cclBegin) {
    const anchorLabel = { pdl: 'PDL', fml: 'FMLA', cfra: 'CFRA' }[r.cclAnchorUsed];
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
  mark(r.lastDay, 'milestone', 'Leave start date');
  mark(r.dueDate, 'milestone', 'Est due date');
  mark(r.actualBirth, 'milestone', 'Actual birth');
  mark(r.returnDate, 'milestone', 'Return to work');
  mark(r.endPIE, 'milestone', 'End of PIE');
  mark(r.firstBday, 'milestone', 'Baby 1st bday');
  mark(r.fileClaim, 'linc', 'File claim');

  mark(r.waitBegin, 'wait', 'Waiting begins');
  mark(r.waitEnd, 'wait', 'Waiting ends');

  const carrierChip = r.isPostdoc ? 'STD' : 'Lincoln';
  mark(r.lincBegin, 'linc', carrierChip + ' begins');
  mark(r.lincEnd, 'linc', carrierChip + ' ends');

  mark(r.pdlBegin, 'pdl', 'PDL begins');
  mark(r.pdlEnd, 'pdl', 'PDL ends');

  mark(r.fmlBegin, 'fml', 'FML begins');
  mark(r.fmlEnd, 'fml', r.fmlCapped ? 'FML ends (CY)' : 'FML ends');
  mark(r.fmlNewYearBegin, 'fml', 'FML resumes');
  mark(r.fmlNewYearEnd, 'fml', 'FML ends');

  mark(r.cfraBegin, 'cfra', 'CFRA begins');
  mark(r.cfraEnd, 'cfra', 'CFRA ends');

  const bondingChip = r.isPostdoc ? 'PPFL' : 'PFCB';
  mark(r.pfcbStart, 'pfcb', bondingChip + ' begins');
  mark(r.pfcbEnd, 'pfcb', bondingChip + ' ends');

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
  /* Eligibility:
     PDL is automatic for California pregnancy, so it is always shown. The
     single "FMLA/CFRA" checkbox is the common UC case where the employee
     qualifies for both federal FMLA and state CFRA, which run concurrently —
     checking it turns on both the FMLA and CFRA leave blocks. */
  const cb = id => {
    const el = document.getElementById(id);
    return el ? el.checked : false;
  };
  const eligFMLCFRA = cb('eligFMLCFRA');
  /* Carrier disability: when "no", the disability waiting period and carrier
     income are omitted from the calculation entirely. The radio keeps its
     historical `lincolnDisability` name; for postdocs it means STD through
     The Standard. */
  const appliesLincoln = (radioValue('lincolnDisability') || 'yes') === 'yes';
  const employeeType = radioValue('employeeType') || 'staff';
  const isPostdoc = employeeType === 'postdoc';

  return {
    employeeType,
    lastDay: parseISO(v('lastDay')),
    dueDate: parseISO(v('dueDate')),
    actualBirth: parseISO(v('actualBirth')),
    returnDate: parseISO(v('returnDate')),
    deliveryType: v('deliveryType'),
    sickHours: n('sickHours'),
    vacHours: n('vacHours'),
    waitingPeriodDays: parseInt(v('waitingPeriod'),10) || (isPostdoc ? 7 : 14),
    appliesLincoln,
    pdlEligible: true,
    fmlEligible: eligFMLCFRA,
    cfraEligible: eligFMLCFRA,
    pfcbWeeks: n('pfcbWeeks'),
    pfcbStart: parseISO(v('pfcbStart')),
    /* CCL is a staff benefit; the field is hidden in postdoc mode, so any
       leftover value is ignored. */
    cclWeeks: isPostdoc ? 0 : n('cclWeeks'),
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
  lastDay:      'Leave start date',
  dueDate:      'Estimated due date',
  deliveryType: 'Delivery type',
  sickHours:    'Total sick hours',
  hoursPerDay:  'Hours per day',
  daysPerWeek:  'Days per week',
  pfcbWeeks:    'PFCB weeks',
  cclWeeks:     'CCL weeks'
};

/* The pfcbWeeks field is labeled "PPFL weeks" in postdoc mode. Resolving the
   label at error-render time keeps FIELD_LABELS itself immutable. */
const fieldLabel = (fieldId) => {
  if (fieldId === 'pfcbWeeks' && document.body.classList.contains('mode-postdoc')) {
    return 'PPFL weeks';
  }
  return FIELD_LABELS[fieldId] || fieldId;
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
    addError('lastDay', 'Leave start date is required.');
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
    addError('dueDate', 'Due date should be on or after the leave start date.');
  }

  if (input.scheduleType === 'variable') {
    if (!(input.hoursPerDay > 0)) {
      addError('hoursPerDay', 'Enter hours per day (greater than 0).');
    }
    if (!(input.daysPerWeek > 0)) {
      addError('daysPerWeek', 'Enter days per week (1–7).');
    }
  }

  /* PFCB / CCL week counts: blank or 0 skips the block; otherwise it must be
     a whole number within the allowed range. CCL's upper bound depends on
     FMLA/CFRA eligibility (14 weeks if eligible, 26 weeks if not). */
  const validateWeeks = (id, max, message) => {
    const field = document.getElementById(id);
    if (!field) return;
    const raw = field.value.trim();
    if (raw === '') return; // blank = skip
    const num = parseFloat(raw);
    if (isNaN(num) || !Number.isInteger(num) || num < 0 || num > max) {
      addError(id, message);
    }
  };
  const isPostdoc = input.employeeType === 'postdoc';
  const bondingTerm = isPostdoc ? 'PPFL' : 'PFCB';
  validateWeeks('pfcbWeeks', 8,
    bondingTerm + ' weeks must be 0 or left blank to skip, or a whole number from 1 to 8.');
  /* CCL does not apply to postdocs — its fields are hidden, so skip
     validating a value the user can no longer see or fix. */
  if (!isPostdoc) {
    const cclMax = input.fmlEligible ? 14 : 26;
    const cclReason = input.fmlEligible
      ? 'Because you are eligible for FMLA/CFRA, CCL is up to 14 weeks.'
      : 'Because you are not eligible for FMLA/CFRA, CCL is up to 26 weeks.';
    validateWeeks('cclWeeks', cclMax,
      'CCL weeks must be 0 or left blank to skip, or a whole number from 1 to ' +
      cclMax + '. ' + cclReason);
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
    const label = fieldLabel(fieldId);
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
    /* Headline announcement: prefer PDL if shown, else FMLA, else just confirm. */
    let saidMsg;
    if (r.pdlBegin) {
      saidMsg = 'Timeline updated. Pregnancy Disability Leave runs from '
        + fmtShort(r.pdlBegin) + ' to ' + fmtShort(r.pdlEnd) + '.';
    } else if (r.fmlBegin) {
      saidMsg = 'Timeline updated. Family and Medical Leave Act leave runs from '
        + fmtShort(r.fmlBegin) + ' to ' + fmtShort(r.fmlEnd) + '.';
    } else {
      saidMsg = 'Timeline updated.';
    }
    liveSay(saidMsg);
  } catch (err) {
    console.error('Calculation failed:', err);
    liveSay('Something went wrong while calculating. Check the browser console for details.');
    alert('Calculation error: ' + err.message);
  }
  return false;
}

function resetAll() {
  document.getElementById('calcForm').reset();
  /* Reset defaults to "Yes" for Lincoln disability, so restore the
     waiting-period field's visibility to match. */
  const wpf = document.getElementById('waitingPeriodField');
  if (wpf) wpf.hidden = false;
  /* Employee type resets to Staff — re-sync the mode class, swapped
     labels, and waiting-period options. */
  if (refreshEmployeeTypeUI) refreshEmployeeTypeUI();
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

/* ============================================================
   Dwell clicking (hover-to-activate)
   -----------------------------------------------------------
   Assistive feature for people who steer a pointer (head
   pointer, eye gaze, joystick) but cannot press a physical
   button. When enabled, resting the pointer on an actionable
   control for the configured dwell time activates it: click
   targets (buttons, links, accordion summaries, radio/checkbox
   labels and inputs) receive a synthetic click; text-entry
   controls and selects receive focus instead, so the user can
   continue with their own input method. Moving the pointer off
   a control before the time elapses cancels the pending
   activation (pointer-cancellation semantics, SC 2.5.2), as
   does a real pointer press. After an activation the control
   will not re-arm until the pointer leaves it, so resting on a
   toggle does not flip it repeatedly. Both the on/off state and
   the dwell time persist in localStorage alongside the theme
   preference.
   ============================================================ */
const DWELL_KEY = 'uc-leave-calc-dwell';
const DWELL_TIME_KEY = 'uc-leave-calc-dwell-time';
const DWELL_DEFAULT_SECONDS = 3;
const DWELL_MIN_SECONDS = 1;
const DWELL_MAX_SECONDS = 60;
/* Elements activated with a synthetic click vs. focused. Radio and
   checkbox inputs are click targets (click() toggles them and fires
   change); text-entry inputs are focus targets. Selects get special
   handling: a synthetic click cannot open a native dropdown (browsers
   require a real user gesture), so dwelling on a select expands it
   into an inline listbox (via the size attribute) whose options are
   themselves dwell targets — dwelling an option chooses it and
   collapses the list. */
const DWELL_CLICK_SELECTOR =
  'button, a[href], summary, label.radio, label.check, ' +
  'input[type="radio"], input[type="checkbox"]';
const DWELL_FOCUS_SELECTOR = 'input, textarea';
const DWELL_SELECT_SELECTOR = 'select, option';

const clampDwellSeconds = n =>
  Math.min(DWELL_MAX_SECONDS, Math.max(DWELL_MIN_SECONDS, Math.round(n)));

function readStoredDwellOn() {
  try { return localStorage.getItem(DWELL_KEY) === 'on'; }
  catch (e) { return false; }
}
function writeStoredDwellOn(on) {
  try { localStorage.setItem(DWELL_KEY, on ? 'on' : 'off'); } catch (e) {}
}
function readStoredDwellSeconds() {
  try {
    const n = parseInt(localStorage.getItem(DWELL_TIME_KEY), 10);
    if (Number.isFinite(n)) return clampDwellSeconds(n);
  } catch (e) {}
  return DWELL_DEFAULT_SECONDS;
}
function writeStoredDwellSeconds(s) {
  try { localStorage.setItem(DWELL_TIME_KEY, String(s)); } catch (e) {}
}

function wireDwell() {
  const toggle = document.getElementById('dwellToggle');
  const field = document.getElementById('dwellTimeField');
  const timeInput = document.getElementById('dwellTime');
  if (!toggle || !field || !timeInput) return;
  const label = toggle.querySelector('.dwell-toggle-label');

  let enabled = readStoredDwellOn();
  let seconds = readStoredDwellSeconds();
  let armedEl = null;   /* element currently counting down */
  let timer = null;
  let firedEl = null;   /* element that just activated — blocked until pointerout */
  let openSelect = null; /* select currently expanded into an inline listbox */

  const collapseSelect = (sel) => {
    if (!sel) return;
    sel.classList.remove('dwell-open');
    sel.removeAttribute('size');
    if (openSelect === sel) openSelect = null;
  };

  const expandSelect = (sel) => {
    if (openSelect && openSelect !== sel) collapseSelect(openSelect);
    openSelect = sel;
    /* size > 1 renders the select as an inline listbox, so its options
       become visible, hoverable dwell targets. At least 2 rows (a
       1-option select still needs listbox rendering), at most 8. */
    sel.size = Math.min(Math.max(sel.options.length, 2), 8);
    sel.classList.add('dwell-open');
    sel.focus();
    liveSay('List expanded. Rest the pointer on an option to choose it, or move away to close.');
  };

  const chooseOption = (opt) => {
    const sel = opt.closest('select');
    if (!sel) return;
    const changed = sel.value !== opt.value;
    sel.value = opt.value;
    collapseSelect(sel);
    /* The pointer is now parked on the collapsed select — block it from
       re-arming (and re-expanding) until the pointer leaves. */
    firedEl = sel;
    if (changed) {
      /* Fire a bubbling change so the app's normal listeners run —
         including the select announcements, which tell the user what
         was chosen. */
      sel.dispatchEvent(new Event('change', { bubbles: true }));
    } else {
      liveSay('Kept ' + (opt.text || opt.value) + '.');
    }
  };

  const disarm = () => {
    if (timer) { clearTimeout(timer); timer = null; }
    if (armedEl) {
      armedEl.classList.remove('dwell-arming');
      armedEl.style.removeProperty('--dwell-ms');
      armedEl = null;
    }
  };

  const resolveTarget = (node) => {
    if (!node || node.nodeType !== 1) return null;
    let el = node.closest(DWELL_CLICK_SELECTOR + ', '
      + DWELL_FOCUS_SELECTOR + ', ' + DWELL_SELECT_SELECTOR);
    if (!el) return null;
    /* Inside an expanded select the options are the targets, not the
       select itself — but only while that select is actually open. */
    if (el.tagName === 'OPTION' && (!openSelect || !openSelect.contains(el))) {
      el = el.closest('select') || el;
    }
    /* Radio/checkbox inputs sit inside their styled label. Normalize to
       the wrapping label so the pointer drifting between the label's
       padding and the input's own box resolves to ONE logical target —
       otherwise the countdown restarts mid-hover and, after a fire, the
       inner input could arm again and toggle the control back. */
    if (el.matches('input[type="radio"], input[type="checkbox"]')) {
      const wrapper = el.closest('label.radio, label.check');
      if (wrapper) el = wrapper;
    }
    /* Never arm a disabled control. Labels have no .disabled of their
       own, so also check the control they wrap. */
    if (el.disabled || el.matches(':disabled')) return null;
    if (el.matches('label')) {
      const inner = el.querySelector('input');
      if (inner && inner.disabled) return null;
    }
    return el;
  };

  const activate = (el) => {
    if (el.tagName === 'SELECT') expandSelect(el);
    else if (el.tagName === 'OPTION') chooseOption(el);
    else if (el.matches(DWELL_CLICK_SELECTOR)) el.click();
    else el.focus();
  };

  const onPointerOver = (e) => {
    if (!enabled) return;
    const target = resolveTarget(e.target);
    if (target === armedEl || (target && target === firedEl)) return;
    disarm();
    firedEl = null;
    if (!target) return;
    armedEl = target;
    target.style.setProperty('--dwell-ms', (seconds * 1000) + 'ms');
    target.classList.add('dwell-arming');
    timer = setTimeout(() => {
      timer = null;
      const el = armedEl;
      disarm();
      firedEl = el;
      activate(el);
    }, seconds * 1000);
  };

  const onPointerOut = (e) => {
    if (!enabled) return;
    const related = e.relatedTarget;
    if (armedEl && armedEl.contains(e.target) &&
        (!related || !armedEl.contains(related))) {
      disarm();
    }
    if (firedEl && firedEl.contains(e.target) &&
        (!related || !firedEl.contains(related))) {
      firedEl = null;
    }
    /* Leaving an expanded select entirely closes it without choosing. */
    if (openSelect && openSelect.contains(e.target) &&
        (!related || !openSelect.contains(related))) {
      collapseSelect(openSelect);
    }
  };

  /* A real pointer press means the user can click — cancel any pending
     dwell so the intended press is the only activation. A press outside
     an expanded select also closes it (a press inside lets the native
     listbox selection happen; the change listener below collapses). */
  const onPointerDown = (e) => {
    disarm();
    firedEl = null;
    if (openSelect && e.target.nodeType === 1 && !openSelect.contains(e.target)) {
      collapseSelect(openSelect);
    }
  };

  /* If the user natively picks an option while the listbox is expanded
     (real click or keyboard), collapse it back to a dropdown. */
  const onChange = (e) => {
    if (openSelect && e.target === openSelect) collapseSelect(openSelect);
  };

  document.addEventListener('pointerover', onPointerOver, true);
  document.addEventListener('pointerout', onPointerOut, true);
  document.addEventListener('pointerdown', onPointerDown, true);
  document.addEventListener('change', onChange, true);

  const applyDwellUI = (announce) => {
    toggle.setAttribute('aria-pressed', enabled ? 'true' : 'false');
    if (label) label.textContent = enabled ? 'Dwell: on' : 'Dwell: off';
    field.hidden = !enabled;
    timeInput.value = String(seconds);
    /* On disable, clear ALL engine state. Note: when a dwell-fire lands
       on the Dwell toggle itself, the fire path sets firedEl just before
       click() runs this handler — clearing it here is correct because the
       pointer handlers ignore events entirely while disabled, so the
       re-arm guard has no job left to do. */
    if (!enabled) { disarm(); firedEl = null; collapseSelect(openSelect); }
    if (announce) {
      liveSay(enabled
        ? 'Dwell clicking on. Rest the pointer on a control for '
          + seconds + ' second' + (seconds === 1 ? '' : 's')
          + ' to activate it.'
        : 'Dwell clicking off.');
    }
  };

  toggle.addEventListener('click', function () {
    enabled = !enabled;
    writeStoredDwellOn(enabled);
    applyDwellUI(true);
  });

  timeInput.addEventListener('change', function () {
    const raw = parseFloat(timeInput.value);
    seconds = Number.isFinite(raw) ? clampDwellSeconds(raw) : DWELL_DEFAULT_SECONDS;
    timeInput.value = String(seconds);
    writeStoredDwellSeconds(seconds);
    liveSay('Dwell time set to ' + seconds
      + ' second' + (seconds === 1 ? '' : 's') + '.');
  });

  applyDwellUI(false);
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

/* Employee type (Staff vs Postdoc). Selecting Postdoc:
   - adds .mode-postdoc to <body>, which swaps every .staff-only /
     .postdoc-only text variant (carrier wording, PFCB→PPFL labels, hints)
     and hides the CCL section (#cclSection carries .staff-only)
   - swaps the waiting-period option: 14 days (Lincoln, staff) vs 7 days
     (The Standard STD, postdoc)
   The current mode is announced through the polite live region so screen
   reader users hear what changed. resetAll() calls refreshEmployeeTypeUI
   to re-sync the UI after the form resets back to Staff. */
let refreshEmployeeTypeUI = null;
function wireEmployeeTypeToggle() {
  const radios = document.querySelectorAll('input[name="employeeType"]');
  const wpSelect = document.getElementById('waitingPeriod');
  if (!radios.length) return;
  const setWaitingOptions = (isPostdoc) => {
    if (!wpSelect) return;
    /* Update the existing option in place — clearing then re-appending
       would leave the select momentarily empty, a transient state some
       assistive tech can pick up. The mode change itself is announced
       through the dedicated polite live region, so the select carries no
       aria-live of its own. */
    const opt = wpSelect.options[0] || wpSelect.appendChild(document.createElement('option'));
    opt.value = isPostdoc ? '7' : '14';
    opt.textContent = isPostdoc ? '7 days (postdoc)' : '14 days (standard)';
    opt.selected = true;
    wpSelect.value = opt.value;
  };
  const update = (announce) => {
    const selected = document.querySelector('input[name="employeeType"]:checked');
    const isPostdoc = !!selected && selected.value === 'postdoc';
    document.body.classList.toggle('mode-postdoc', isPostdoc);
    setWaitingOptions(isPostdoc);
    /* The icon-only info button's accessible name follows the visible
       group title (PFCB for staff, PPFL for postdocs). */
    const pfcbBtn = document.getElementById('pfcbInfoBtn');
    if (pfcbBtn) pfcbBtn.setAttribute('aria-label',
      isPostdoc ? 'About PPFL eligibility' : 'About PFCB eligibility');
    if (announce) {
      liveSay(isPostdoc
        ? 'Postdoc selected: 7-day waiting period, short-term disability through The Standard, PPFL instead of PFCB, no Child Caring Leave.'
        : 'Staff selected: 14-day waiting period, Lincoln Financial disability, PFCB and Child Caring Leave available.');
    }
  };
  radios.forEach(r => r.addEventListener('change', function () { update(true); }));
  refreshEmployeeTypeUI = function () { update(false); };
  update(false);
}

/* Show/hide the disability waiting-period field based on whether the employee
   is applying for disability leave through Lincoln Financial. "Yes" reveals
   the waiting-period selector; "No" hides it. */
function wireLincolnToggle() {
  const radios = document.querySelectorAll('input[name="lincolnDisability"]');
  const field = document.getElementById('waitingPeriodField');
  if (!radios.length || !field) return;
  const update = () => {
    const selected = document.querySelector('input[name="lincolnDisability"]:checked');
    field.hidden = !selected || selected.value !== 'yes';
  };
  radios.forEach(r => r.addEventListener('change', update));
  update();
}

/* Toggle the monthly calendars between a one-column and two-column layout.
   The chosen button reflects its state via aria-pressed; the calendar grid
   carries a cols-1 / cols-2 class that the stylesheet maps to a column count. */
function wireCalColumns() {
  const wrap = document.getElementById('calendars');
  const btn1 = document.getElementById('calCols1');
  const btn2 = document.getElementById('calCols2');
  if (!wrap || !btn1 || !btn2) return;
  const setCols = (n) => {
    wrap.classList.toggle('cols-1', n === 1);
    wrap.classList.toggle('cols-2', n === 2);
    btn1.setAttribute('aria-pressed', String(n === 1));
    btn2.setAttribute('aria-pressed', String(n === 2));
    liveSay(n === 1 ? 'Calendars shown in one column.' : 'Calendars shown in two columns.');
  };
  btn1.addEventListener('click', () => setCols(1));
  btn2.addEventListener('click', () => setCols(2));
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

/* Save-as-PDF handler.
   Uses the browser's native print pipeline with "Save as PDF" as a destination
   — this keeps us fully CSP-safe (no external libraries, no network) while
   reusing the @media print stylesheet that already hides the form, forces the
   light palette, opens accordions, and sets good page breaks.

   The one thing we add over plain Print is a descriptive document.title during
   the print job: browsers use document.title as the default PDF filename, so
   swapping it from the generic page title to something like
   "UC-Maternity-Leave-Timeline-2026-05-01.pdf" gives the user a meaningful
   filename by default. Title is restored on the 'afterprint' event. */
function saveAsPDF() {
  const lastDayInput = document.getElementById('lastDay');
  const stamp = (lastDayInput && lastDayInput.value)
    ? lastDayInput.value
    : new Date().toISOString().slice(0, 10);
  const desiredTitle = 'UC-Maternity-Leave-Timeline-' + stamp;
  const originalTitle = document.title;

  const restore = function () {
    document.title = originalTitle;
    window.removeEventListener('afterprint', restore);
  };
  window.addEventListener('afterprint', restore);

  document.title = desiredTitle;
  window.print();

  /* Fallback restore for browsers that don't fire afterprint reliably
     (older Safari). The print dialog is synchronous in most engines, so by the
     time we get here the user has already dismissed it and we can safely
     restore. */
  setTimeout(restore, 1000);
}

/* Wire handlers once DOM is ready. Uses both a submit handler (the correct
   semantic) and a click handler on the button (belt-and-braces fallback in
   case some browser quirk or embedded preview swallows the submit event). */
function wire() {
  const form = document.getElementById('calcForm');
  const calcBtn = document.querySelector('button.btn-primary');
  const resetBtn = document.getElementById('resetBtn');
  const printBtn = document.getElementById('printBtn');
  const pdfBtn = document.getElementById('pdfBtn');
  const themeBtn = document.getElementById('themeToggle');

  if (form) form.addEventListener('submit', run);
  if (calcBtn) calcBtn.addEventListener('click', run);
  if (resetBtn) resetBtn.addEventListener('click', resetAll);
  if (printBtn) printBtn.addEventListener('click', function () { window.print(); });
  if (pdfBtn) pdfBtn.addEventListener('click', saveAsPDF);
  if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

  wireInfoButtons();
  wireDwell();
  wireEmployeeTypeToggle();
  wireScheduleToggle();
  wireLincolnToggle();
  wireCalColumns();
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
