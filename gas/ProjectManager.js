/**
 * ProjectManager.gs — Calm "Focus Now" Dispatcher, Morning Brain Dump,
 * Universal 3-Slot Task Scheduler (10m Buffer), Project Breakdowns & Tagged Gmail Triage
 */

const PM_TZ = typeof TIMEZONE !== 'undefined' ? TIMEZONE : 'America/Bogota';
const SCHOOL_BRIDGE_URL = 'PASTE_COJOWA_WEB_APP_URL_HERE?token=YOUR_SECRET_TOKEN';
const COS_GMAIL_LABEL = 'ChiefOfStaff';

const DEFAULT_PROJECTS = [
  {
    projectId: 'PRJ_WAL',
    name: 'Workload Automation Lab / PD',
    tier: 'Tier 1B - Consulting/PD',
    phase: 'Active Sprint',
    deadline: 'Next César Sync',
    pacingRule: '15m weekday bites / 45m weekend blocks',
    notes: 'Consulting prototype and PD workshop curriculum for educators. Break work into 15-minute modular pieces: outlining single slides, testing one script helper, or drafting one facilitator prompt.',
    checklist: [
      { id: 'CHK_WAL_1', text: 'Define current phase deliverables in description', done: false, scheduledSlot: '' }
    ]
  },
  {
    projectId: 'PRJ_COJOWA',
    name: 'COJOWA Physics / Async Videos',
    tier: 'Tier 1A - Teaching',
    phase: 'Weekly Async Content',
    deadline: 'Fridays 3:30 PM',
    pacingRule: '15m Plan-period sprints',
    notes: '4 weekly async physics instructional videos and rubric grading. Break into 15m chunks: script 1 video outline, record 1 video take, or grade 4 lab submissions.',
    checklist: [
      { id: 'CHK_COJ_1', text: 'Outline video #1 key equations & example (15m)', done: false, scheduledSlot: '' }
    ]
  }
];

function ensureProjectsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Projects');
  const headers = [
    'Project_ID', 'Project_Name', 'Tier', 'Current_Phase',
    'Target_Deadline', 'Pacing_Rule', 'Context_Notes', 'Checklist_JSON', 'Updated_At', 'Status'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Projects');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');
    const rows = DEFAULT_PROJECTS.map(p => [
      p.projectId, p.name, p.tier, p.phase, p.deadline,
      p.pacingRule, p.notes, JSON.stringify(p.checklist), nowStr, 'Active'
    ]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  } else if (sheet.getLastColumn() < 10) {
    sheet.getRange(1, 10).setValue('Status').setFontWeight('bold');
  }
  return sheet;
}

function ensureRunningTodosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Running_Todos');
  const headers = [
    'Todo_ID', 'Text', 'Category', 'Est_Min', 'Is_Daily_Habit', 'Done', 'Completed_Date', 'Created_At', 'Scheduled_Slot'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Running_Todos');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');
    sheet.appendRow([
      'RTD_' + Utilities.getUuid().slice(0, 6).toUpperCase(),
      '15m Spanish practice (listening, reading, or vocab)',
      'Spanish',
      15,
      true,
      false,
      '',
      nowStr,
      ''
    ]);
  } else if (sheet.getLastColumn() < 9) {
    sheet.getRange(1, 9).setValue('Scheduled_Slot').setFontWeight('bold');
  }
  return sheet;
}

function readProjectCards_(pruneCompleted) {
  const sheet = ensureProjectsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const cards = [];

  values.forEach((r, idx) => {
    const pid = String(r[0] || '').trim();
    const status = String(r[9] || 'Active').trim();
    if (!pid || status === 'Archived') return;

    let checklist = [];
    try {
      checklist = JSON.parse(r[7] || '[]');
    } catch (e) {
      checklist = [];
    }
    if (!Array.isArray(checklist)) checklist = [];

    if (pruneCompleted) {
      const openOnly = checklist.filter(item => !item.done);
      if (openOnly.length !== checklist.length) {
        sheet.getRange(idx + 2, 8).setValue(JSON.stringify(openOnly));
      }
      checklist = openOnly;
    }

    cards.push({
      projectId: pid,
      name: String(r[1]),
      tier: String(r[2]),
      phase: String(r[3]),
      deadline: String(r[4]),
      pacingRule: String(r[5]),
      notes: String(r[6]),
      checklist: checklist,
      updatedAt: String(r[8] || ''),
      status: status
    });
  });

  return cards;
}

function getProjectCards() {
  return readProjectCards_(true);
}

function getRunningTodos() {
  const sheet = ensureRunningTodosSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const todayStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd');
  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const activeTodos = [];

  values.forEach((r, idx) => {
    const id = String(r[0] || '').trim();
    const text = String(r[1] || '').trim();
    if (!id || !text) return;

    const category = String(r[2] || 'Quick / Admin').trim();
    const estMin = Number(r[3]) || 15;
    const isDaily = Boolean(r[4]);
    let isDone = Boolean(r[5]);
    const completedDate = String(r[6] || '').trim();
    const scheduledSlot = String(r[8] || '').trim();

    if (isDaily && isDone && completedDate !== todayStr) {
      isDone = false;
      sheet.getRange(idx + 2, 6).setValue(false);
      sheet.getRange(idx + 2, 9).setValue('');
    }

    if (!isDone) {
      activeTodos.push({
        todoId: id,
        text: text,
        category: category,
        estMin: estMin,
        isDailyHabit: isDaily,
        scheduledSlot: scheduledSlot
      });
    }
  });

  return activeTodos;
}

function getInitialDashboardData() {
  return {
    cards: readProjectCards_(true),
    runningTodos: getRunningTodos()
  };
}

function addRunningTodo(text, category, estMin, isDailyHabit) {
  const cleanText = String(text || '').trim();
  if (!cleanText) return getRunningTodos();

  const sheet = ensureRunningTodosSheet_();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');
  const id = 'RTD_' + Utilities.getUuid().slice(0, 6).toUpperCase();
  const cleanCat = String(category || (/\bspanish\b/i.test(cleanText) ? 'Spanish' : 'Quick / Admin')).trim();
  const mins = Math.max(5, Math.min(90, Number(estMin) || 15));
  const daily = Boolean(isDailyHabit) || /\bspanish\b/i.test(cleanCat);

  sheet.appendRow([id, cleanText, cleanCat, mins, daily, false, '', nowStr, '']);
  return getRunningTodos();
}

function completeRunningTodo(todoId) {
  const sheet = ensureRunningTodosSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const todayStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(todoId)) {
      sheet.getRange(i + 2, 6).setValue(true);
      sheet.getRange(i + 2, 7).setValue(todayStr);
      sheet.getRange(i + 2, 9).setValue('');
      break;
    }
  }
  return getRunningTodos();
}

function deleteRunningTodo(todoId) {
  const sheet = ensureRunningTodosSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(todoId)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  return getRunningTodos();
}

/* ================= MORNING BRAIN DUMP ================= */

function processMorningBrainDump(rawDumpText) {
  const input = String(rawDumpText || '').trim();
  if (!input) {
    throw new Error('Write down what you need to get done today first.');
  }

  const cards = readProjectCards_(true);
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'EEEE, MMM d, yyyy h:mm a');
  const dayOfWeek = Number(Utilities.formatDate(new Date(), PM_TZ, 'u'));
  const isWeekend = (dayOfWeek === 6 || dayOfWeek === 7);

  const portfolioCtx = cards.length > 0
    ? cards.map(c => `PROJECT [${c.projectId}]: ${c.name} (${c.tier}) | Deadline: ${c.deadline}\n  Description/Breakdown Guide: ${c.notes || '(none)'}`).join('\n')
    : '(No active projects yet)';

  const scrubbed = typeof scrubPII === 'function' ? scrubPII(input) : input;

  const prompt = `Current Time: ${nowStr} (${isWeekend ? 'Weekend — 30-60m blocks available' : 'School Day — 15m micro-chunks preferred'}).
Joel just wrote his brain dump of tasks to tackle.
Parse every distinct task/thought and route each one cleanly:
1. If an item belongs to an existing ACTIVE PROJECT below, set "destination": "PROJECT" and provide the matching "project_id". If big, break it into 15-minute bite-sized steps using the project's Description/Breakdown Guide.
2. Otherwise (standalone tasks, quick 5-15m errands, teaching one-offs, Spanish practice, personal admin), set "destination": "RUNNING_TODO" and assign the apparent "category" ("Teaching", "Spanish", "Consulting/PD", "Personal", or "Quick / Admin"), realistic "est_min" (5, 15, 30, 45, or 60), and "is_daily_habit" (true only for repeatable daily habits like Spanish practice).

ACTIVE PROJECTS:
${portfolioCtx}

BRAIN DUMP:
${scrubbed}`;

  const schema = {
    type: 'OBJECT',
    properties: {
      summary: { type: 'STRING' },
      routed_items: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            destination: { type: 'STRING', enum: ['PROJECT', 'RUNNING_TODO'] },
            project_id: { type: 'STRING', nullable: true },
            text: { type: 'STRING' },
            category: { type: 'STRING', enum: ['Teaching', 'Spanish', 'Consulting/PD', 'Personal', 'Quick / Admin'] },
            est_min: { type: 'INTEGER' },
            is_daily_habit: { type: 'BOOLEAN' }
          },
          required: ['destination', 'text', 'category', 'est_min', 'is_daily_habit']
        }
      }
    },
    required: ['summary', 'routed_items']
  };

  const parsed = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'morning_brain_dump'
  );

  const restore = txt => (txt && typeof restorePII === 'function' ? restorePII(txt) : txt);
  const validProjectIds = new Set(cards.map(c => c.projectId));
  const projectMap = {};
  cards.forEach(c => { projectMap[c.projectId] = c.name; });

  const logLines = [];
  (parsed.routed_items || []).forEach(it => {
    const cleanText = restore(it.text).trim();
    if (!cleanText) return;

    if (it.destination === 'PROJECT' && it.project_id && validProjectIds.has(String(it.project_id))) {
      addProjectChecklistItem(String(it.project_id), cleanText);
      logLines.push(`→ [${projectMap[it.project_id]}] ${cleanText}`);
    } else {
      addRunningTodo(cleanText, it.category, it.est_min || 15, Boolean(it.is_daily_habit));
      logLines.push(`→ [${it.category} • ${it.est_min || 15}m] ${cleanText}`);
    }
  });

  return {
    summary: restore(parsed.summary),
    logLines: logLines,
    cards: readProjectCards_(true),
    runningTodos: getRunningTodos()
  };
}

/* ================= "WHAT DO I DO NOW?" DISPATCHER ================= */

function getVisibleBusyAndPlanEvents_(startDate, endDate) {
  const calendars = CalendarApp.getAllCalendars().filter(cal => {
    try {
      return cal.isMyPrimaryCalendar() || cal.isSelected();
    } catch (e) {
      return false;
    }
  });

  const busyBlocks = [];
  const planBlocks = [];

  calendars.forEach(cal => {
    let events = [];
    try {
      events = cal.getEvents(startDate, endDate);
    } catch (e) {
      return;
    }
    events.forEach(ev => {
      if (ev.isAllDayEvent()) return;
      const myStatus = typeof ev.getMyStatus === 'function' ? ev.getMyStatus() : null;
      if (myStatus === CalendarApp.GuestStatus.NO) return;

      const title = String(ev.getTitle() || 'Commitment');
      if (/\b(plan|prep)\b/i.test(title) && !/\[Focus\]/i.test(title)) {
        planBlocks.push({ title: title, start: ev.getStartTime(), end: ev.getEndTime() });
        return;
      }

      busyBlocks.push({ title: title, start: ev.getStartTime(), end: ev.getEndTime() });
    });
  });

  busyBlocks.sort((a, b) => a.start.getTime() - b.start.getTime());
  planBlocks.sort((a, b) => a.start.getTime() - b.start.getTime());
  return { busyBlocks, planBlocks };
}

function whatDoIDoNow() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, PM_TZ, 'yyyy-MM-dd');
  const tzOffset = Utilities.formatDate(now, PM_TZ, 'XXX');
  const dayCutoff = new Date(`${dateStr}T19:00:00${tzOffset}`);

  if (now >= dayCutoff) {
    throw new Error('It is past your 7:00 PM cutoff — protect your evening recharge and pick back up tomorrow!');
  }

  const dayOfWeek = Number(Utilities.formatDate(now, PM_TZ, 'u'));
  const isWeekend = (dayOfWeek === 6 || dayOfWeek === 7);
  const { busyBlocks } = getVisibleBusyAndPlanEvents_(now, dayCutoff);

  const BUFFER_MIN = 10;
  let windowStart = new Date(now.getTime());
  let windowEnd = dayCutoff;
  let nextCommitmentLabel = '7:00 PM cutoff';
  let statusBanner = '';

  const currentBusy = busyBlocks.find(b => b.start <= now && b.end > now);
  if (currentBusy) {
    windowStart = new Date(currentBusy.end.getTime() + BUFFER_MIN * 60000);
    const following = busyBlocks.find(b => b.start >= currentBusy.end);
    if (following) {
      windowEnd = new Date(following.start.getTime() - BUFFER_MIN * 60000);
      nextCommitmentLabel = `${following.title} at ${Utilities.formatDate(following.start, PM_TZ, 'h:mm a')}`;
    }
    statusBanner = `In "${currentBusy.title}" until ${Utilities.formatDate(currentBusy.end, PM_TZ, 'h:mm a')}. After a 10m brain break, your next window is ${Utilities.formatDate(windowStart, PM_TZ, 'h:mm a')} – ${Utilities.formatDate(windowEnd, PM_TZ, 'h:mm a')}.`;
  } else {
    const nextBusy = busyBlocks.find(b => b.start > now);
    if (nextBusy) {
      const rawMins = Math.floor((nextBusy.start.getTime() - now.getTime()) / 60000);
      if (rawMins < 15) {
        windowStart = new Date(nextBusy.end.getTime() + BUFFER_MIN * 60000);
        const afterNext = busyBlocks.find(b => b.start >= nextBusy.end);
        if (afterNext) {
          windowEnd = new Date(afterNext.start.getTime() - BUFFER_MIN * 60000);
          nextCommitmentLabel = `${afterNext.title} at ${Utilities.formatDate(afterNext.start, PM_TZ, 'h:mm a')}`;
        }
        statusBanner = `"${nextBusy.title}" starts in ${rawMins}m — take a 10m brain break first. Next window: ${Utilities.formatDate(windowStart, PM_TZ, 'h:mm a')} – ${Utilities.formatDate(windowEnd, PM_TZ, 'h:mm a')} (10m buffer before ${nextCommitmentLabel}).`;
      } else {
        windowEnd = new Date(nextBusy.start.getTime() - BUFFER_MIN * 60000);
        nextCommitmentLabel = `${nextBusy.title} at ${Utilities.formatDate(nextBusy.start, PM_TZ, 'h:mm a')}`;
        const usable = Math.max(5, Math.floor((windowEnd.getTime() - windowStart.getTime()) / 60000));
        statusBanner = `${usable} usable mins right now (${Utilities.formatDate(windowStart, PM_TZ, 'h:mm a')} – ${Utilities.formatDate(windowEnd, PM_TZ, 'h:mm a')}, with a 10m brain break before ${nextCommitmentLabel}).`;
      }
    } else {
      const usable = Math.max(5, Math.floor((dayCutoff.getTime() - now.getTime()) / 60000));
      statusBanner = `${usable} usable mins open until your 7:00 PM cutoff.`;
    }
  }

  const usableMinutes = Math.max(5, Math.floor((windowEnd.getTime() - windowStart.getTime()) / 60000));
  if (windowStart >= dayCutoff || usableMinutes < 5) {
    throw new Error('No usable windows left before your 7:00 PM cutoff today.');
  }

  const cards = readProjectCards_(true);
  const runningTodos = getRunningTodos();

  if (cards.length === 0 && runningTodos.length === 0) {
    throw new Error('Your lists are empty! Add a task above or create a project in the Projects tab.');
  }

  const projectsText = cards.map(c => {
    const items = (c.checklist || []).map(i => `  - [ITEM_ID: ${i.id}] ${i.text}`).join('\n') || '  (No specific items — use Project Description to propose a next step)';
    return `PROJECT [PROJ_ID: ${c.projectId}]: ${c.name} | Tier: ${c.tier} | Deadline: ${c.deadline}
  Description/Breakdown Guide: ${c.notes || '(none)'}
  Open Items:\n${items}`;
  }).join('\n\n') || '(None)';

  const runningText = runningTodos.map(t =>
    `- [TODO_ID: ${t.todoId}] (${t.category}, ~${t.estMin}m${t.isDailyHabit ? ', Daily Habit' : ''}) ${t.text}`
  ).join('\n') || '(None)';

  const modeGuidance = isWeekend
    ? `TODAY IS A WEEKEND: Prefer 30 to 60 minute deep-work chunks (as long as duration_min <= ${usableMinutes}), while still offering a 15m option if available.`
    : `TODAY IS A SCHOOL DAY: Prefer 15-minute micro-chunks (15m, or 30m max if window is wide) so Joel can chip away steadily between classes. Use a project's Description/Breakdown Guide to frame a concrete 15-minute bite of larger items.`;

  const prompt = `Joel just clicked "What Do I Do Now?".
TIME WINDOW AVAILABLE: ${usableMinutes} usable minutes (${Utilities.formatDate(windowStart, PM_TZ, 'h:mm a')} – ${Utilities.formatDate(windowEnd, PM_TZ, 'h:mm a')}, includes a 10m brain break before ${nextCommitmentLabel}).
${modeGuidance}

ACTIVE PROJECTS & DESCRIPTIONS:
${projectsText}

RUNNING TO-DO LIST:
${runningText}

INSTRUCTIONS:
Return 3 to 4 concrete options that strictly fit inside ${usableMinutes} minutes (every option's "duration_min" MUST be <= ${usableMinutes}).
1. Option #1 MUST be the highest-priority or closest-deadline task that fits this window.
2. Include variety across the remaining options so Joel can choose based on his energy (e.g., another high-leverage project bite, his repeatable 15m Spanish practice if listed and >= 15m available, or a quick 5–15m Running To-Do item).
3. Link each option cleanly to its source:
   - Project checklist item: "source_type": "PROJECT_ITEM", "parent_id": PROJ_ID, "item_id": ITEM_ID.
   - Sliced from Project without specific item ID: "source_type": "PROJECT_NEXT", "parent_id": PROJ_ID, "item_id": "".
   - Running To-Do list: "source_type": "RUNNING_TODO", "parent_id": "", "item_id": TODO_ID.`;

  const schema = {
    type: 'OBJECT',
    properties: {
      window_advice: { type: 'STRING' },
      options: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            rank: { type: 'INTEGER' },
            title: { type: 'STRING' },
            category_badge: { type: 'STRING' },
            tier: { type: 'STRING' },
            duration_min: { type: 'INTEGER' },
            priority_reason: { type: 'STRING' },
            two_min_start: { type: 'STRING' },
            stopping_rule: { type: 'STRING' },
            source_type: { type: 'STRING', enum: ['PROJECT_ITEM', 'PROJECT_NEXT', 'RUNNING_TODO'] },
            parent_id: { type: 'STRING' },
            item_id: { type: 'STRING' }
          },
          required: [
            'rank', 'title', 'category_badge', 'tier', 'duration_min',
            'priority_reason', 'two_min_start', 'stopping_rule', 'source_type', 'parent_id', 'item_id'
          ]
        }
      }
    },
    required: ['window_advice', 'options']
  };

  const aiRes = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'what_do_i_do_now'
  );

  const restore = txt => (txt && typeof restorePII === 'function' ? restorePII(txt) : txt);
  const start24 = Utilities.formatDate(windowStart, PM_TZ, 'HH:mm');

  const options = (aiRes.options || []).map(opt => ({
    rank: opt.rank,
    title: restore(opt.title),
    categoryBadge: restore(opt.category_badge),
    tier: opt.tier || 'Tier 1A - Teaching',
    durationMin: Math.max(5, Math.min(usableMinutes, Number(opt.duration_min) || 15)),
    priorityReason: restore(opt.priority_reason),
    twoMinStart: restore(opt.two_min_start),
    stoppingRule: restore(opt.stopping_rule),
    sourceType: opt.source_type,
    parentId: opt.parent_id || '',
    itemId: opt.item_id || ''
  }));

  return {
    statusBanner: statusBanner,
    windowAdvice: restore(aiRes.window_advice),
    dateStr: dateStr,
    startTime24: start24,
    usableMinutes: usableMinutes,
    isWeekend: isWeekend,
    options: options
  };
}

function completeRecommendedOption(sourceType, parentId, itemId) {
  if (sourceType === 'PROJECT_ITEM' && parentId && itemId) {
    toggleProjectChecklistItem(parentId, itemId, true);
  } else if (sourceType === 'RUNNING_TODO' && itemId) {
    completeRunningTodo(itemId);
  }
  return getInitialDashboardData();
}

/* ================= UNIVERSAL 3-SLOT TASK SCHEDULER ================= */

/**
 * Finds 3 open calendar slots across today + next 3 days with a 10-minute brain-break buffer.
 */
function findScheduleSlotsForTask(title, requestedMin, sourceType, parentId, itemId) {
  const now = new Date();
  const dayOfWeek = Number(Utilities.formatDate(now, PM_TZ, 'u'));
  const isWeekendNow = (dayOfWeek === 6 || dayOfWeek === 7);
  const durationMin = Math.max(10, Math.min(90, Number(requestedMin) || (isWeekendNow ? 45 : 15)));
  const BUFFER_MS = 10 * 60000;
  const STEP_MS = 15 * 60000;

  const horizonEnd = new Date(now.getTime() + 4 * 24 * 60 * 60000);
  const { busyBlocks, planBlocks } = getVisibleBusyAndPlanEvents_(now, horizonEnd);

  const candidates = [];

  for (let d = 0; d < 4; d++) {
    const dayObj = new Date(now.getTime() + d * 24 * 60 * 60000);
    const dateStr = Utilities.formatDate(dayObj, PM_TZ, 'yyyy-MM-dd');
    const tzOffset = Utilities.formatDate(dayObj, PM_TZ, 'XXX');
    const dayName = d === 0 ? 'Today' : (d === 1 ? 'Tomorrow' : Utilities.formatDate(dayObj, PM_TZ, 'EEE, MMM d'));

    let windowStart = new Date(`${dateStr}T07:30:00${tzOffset}`);
    const windowEnd = new Date(`${dateStr}T19:00:00${tzOffset}`);

    if (d === 0) {
      // Round current time up to the next 5-minute increment
      const roundedNow = new Date(Math.ceil(now.getTime() / (5 * 60000)) * (5 * 60000));
      if (roundedNow > windowStart) windowStart = roundedNow;
    }

    for (let t = windowStart.getTime(); t + durationMin * 60000 <= windowEnd.getTime(); t += STEP_MS) {
      const slotStart = t;
      const slotEnd = t + durationMin * 60000;

      const hasCollision = busyBlocks.some(b => {
        const paddedStart = b.start.getTime() - BUFFER_MS;
        const paddedEnd = b.end.getTime() + BUFFER_MS;
        return slotStart < paddedEnd && slotEnd > paddedStart;
      });

      if (!hasCollision) {
        const matchingPlan = planBlocks.find(p => slotStart >= p.start.getTime() && slotEnd <= p.end.getTime());
        const startDateObj = new Date(slotStart);
        const endDateObj = new Date(slotEnd);
        const start12 = Utilities.formatDate(startDateObj, PM_TZ, 'h:mm a');
        const end12 = Utilities.formatDate(endDateObj, PM_TZ, 'h:mm a');

        candidates.push({
          slotStartMs: slotStart,
          dateStr: dateStr,
          startTime24: Utilities.formatDate(startDateObj, PM_TZ, 'HH:mm'),
          durationMin: durationMin,
          isPlan: Boolean(matchingPlan),
          label: `${dayName} • ${start12} – ${end12}`,
          sublabel: matchingPlan ? `Inside ${matchingPlan.title} (10m buffer protected)` : 'Open window (10m buffer protected)',
          shortBadge: `${d === 0 ? 'Today' : Utilities.formatDate(startDateObj, PM_TZ, 'EEE')} ${start12}`
        });
      }
    }
  }

  if (candidates.length === 0) {
    throw new Error('No open slots with a 10-minute buffer found before 7:00 PM in the next 3 days.');
  }

  // Pick up to 3 well-spaced slots:
  // 1) Earliest Plan-period slot today (or earliest overall slot)
  // 2) Next distinct window at least 45 minutes later
  // 3) Next distinct window on a later block or tomorrow
  const chosen = [];
  const todayPlan = candidates.find(c => c.dateStr === candidates[0].dateStr && c.isPlan);
  chosen.push(todayPlan || candidates[0]);

  for (let i = 0; i < candidates.length && chosen.length < 3; i++) {
    const cand = candidates[i];
    const farEnough = chosen.every(ch => Math.abs(cand.slotStartMs - ch.slotStartMs) >= 45 * 60000);
    if (farEnough) {
      chosen.push(cand);
    }
  }

  return {
    title: String(title || '').trim(),
    sourceType: sourceType || 'RUNNING_TODO',
    parentId: parentId || '',
    itemId: itemId || '',
    slots: chosen
  };
}

/**
 * Books the user-approved calendar slot and stamps a scheduled badge onto the task item.
 */
function approveScheduledSlot(spec) {
  const title = String(spec.title || 'Focus Task').trim();
  const dateStr = String(spec.dateStr || '').trim();
  const startTime24 = String(spec.startTime24 || '').trim();
  const durationMin = Math.max(10, Number(spec.durationMin) || 15);
  const shortBadge = String(spec.shortBadge || `${dateStr} ${startTime24}`).trim();

  if (typeof commitScheduledBlocks === 'function') {
    commitScheduledBlocks([{
      isProtectedAnchor: false,
      writeToCalendar: true,
      existingTaskId: null,
      projectId: spec.parentId || '',
      projectTag: spec.sourceType === 'PROJECT_ITEM' ? 'Project' : 'Running To-Do',
      title: title,
      tier: 'Tier 1A - Teaching',
      dateStr: dateStr,
      startTimeStr: startTime24,
      durationMin: durationMin,
      twoMinStart: `Open ${title}`,
      stoppingRule: `Stop cleanly after ${durationMin}m for a 10m brain break`
    }]);
  } else {
    const tzOffset = Utilities.formatDate(new Date(), PM_TZ, 'XXX');
    const startDt = new Date(`${dateStr}T${startTime24}:00${tzOffset}`);
    const endDt = new Date(startDt.getTime() + durationMin * 60000);
    CalendarApp.getDefaultCalendar().createEvent(`[Focus] ${title}`, startDt, endDt);
  }

  // Stamp the scheduled time badge onto the underlying Running To-Do or Project item
  if (spec.sourceType === 'RUNNING_TODO' && spec.itemId) {
    const sheet = ensureRunningTodosSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(spec.itemId)) {
          sheet.getRange(i + 2, 9).setValue(shortBadge);
          break;
        }
      }
    }
  } else if (spec.sourceType === 'PROJECT_ITEM' && spec.parentId && spec.itemId) {
    const sheet = ensureProjectsSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(spec.parentId)) {
        let list = [];
        try { list = JSON.parse(data[i][7] || '[]'); } catch (e) {}
        list = list.map(it => (String(it.id) === String(spec.itemId) ? { ...it, scheduledSlot: shortBadge } : it));
        sheet.getRange(i + 1, 8).setValue(JSON.stringify(list));
        break;
      }
    }
  }

  return {
    message: `Scheduled "${title}" for ${shortBadge}`,
    dashboard: getInitialDashboardData()
  };
}

/* ================= PROJECT CARDS & BREAKDOWN ================= */

function breakDownProjectWithAI(projectId) {
  const cards = readProjectCards_(true);
  const project = cards.find(c => c.projectId === String(projectId));
  if (!project) throw new Error('Project card not found.');

  const currentItems = (project.checklist || []).map(i => `- ${i.text}`).join('\n') || '(No items yet)';
  const prompt = `Break down this project into concrete, bite-sized 15-minute steps using its Description & Breakdown Guide.
Keep existing small 15m items intact, split any large/vague items into sequential 15-minute micro-actions, and produce 4 to 8 clear, actionable checklist items.

PROJECT: ${project.name} (${project.tier})
DEADLINE: ${project.deadline}
DESCRIPTION & BREAKDOWN GUIDE:
${project.notes || '(No description provided — infer sensible 15m steps from the project title and existing tasks)'}

CURRENT LIST:
${currentItems}`;

  const schema = {
    type: 'OBJECT',
    properties: {
      steps: {
        type: 'ARRAY',
        items: { type: 'STRING' }
      }
    },
    required: ['steps']
  };

  const res = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'breakdown_project_15m'
  );

  const restore = txt => (txt && typeof restorePII === 'function' ? restorePII(txt) : txt);
  const newChecklist = (res.steps || [])
    .map(s => restore(s).trim())
    .filter(Boolean)
    .map(text => ({
      id: 'CHK_' + Utilities.getUuid().slice(0, 6).toUpperCase(),
      text: text,
      done: false,
      scheduledSlot: ''
    }));

  if (newChecklist.length > 0) {
    const sheet = ensureProjectsSheet_();
    const data = sheet.getDataRange().getValues();
    const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(projectId)) {
        sheet.getRange(i + 1, 8).setValue(JSON.stringify(newChecklist));
        sheet.getRange(i + 1, 9).setValue(nowStr);
        break;
      }
    }
  }

  return readProjectCards_(true);
}

function archiveAllActiveProjects() {
  const sheet = ensureProjectsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][9] || 'Active').trim() !== 'Archived') {
      data[i][8] = nowStr;
      data[i][9] = 'Archived';
    }
  }
  sheet.getRange(2, 1, data.length, 10).setValues(data);
  return [];
}

function archiveProjectCard(projectId) {
  const sheet = ensureProjectsSheet_();
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(projectId)) {
      sheet.getRange(i + 1, 9).setValue(nowStr);
      sheet.getRange(i + 1, 10).setValue('Archived');
      return readProjectCards_(true);
    }
  }
  return readProjectCards_(true);
}

function createProjectCard(projectSpec) {
  if (!projectSpec || !String(projectSpec.name || '').trim()) {
    throw new Error('Project title is required.');
  }

  const sheet = ensureProjectsSheet_();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');
  const projectId = 'PRJ_' + Utilities.getUuid().slice(0, 6).toUpperCase();

  const rawItems = Array.isArray(projectSpec.checklistItems) ? projectSpec.checklistItems : [];
  const checklist = rawItems
    .map(item => String(typeof item === 'string' ? item : (item && item.text) || '').trim())
    .filter(Boolean)
    .map(text => ({
      id: 'CHK_' + Utilities.getUuid().slice(0, 6).toUpperCase(),
      text: text,
      done: false,
      scheduledSlot: ''
    }));

  sheet.appendRow([
    projectId,
    String(projectSpec.name).trim(),
    String(projectSpec.tier || 'Tier 1B - Consulting/PD').trim(),
    String(projectSpec.phase || 'Active').trim(),
    String(projectSpec.deadline || 'Ongoing').trim(),
    String(projectSpec.pacingRule || '15m weekday / 45m weekend chunks').trim(),
    String(projectSpec.description || '').trim(),
    JSON.stringify(checklist),
    nowStr,
    'Active'
  ]);

  return {
    createdProjectId: projectId,
    cards: readProjectCards_(true)
  };
}

function createProjectFromDescription(rawDescription) {
  const input = String(rawDescription || '').trim();
  if (!input) {
    throw new Error('Describe the project and how it should be broken down first.');
  }

  const scrubbed = typeof scrubPII === 'function' ? scrubPII(input) : input;
  const prompt = `Joel wants to set up a Project Card with a clear Description & Breakdown Guide so future work can be chipped away in 15-minute increments.
Extract:
1. "name": Clear project title.
2. "tier": Best matching portfolio tier.
3. "phase": Current phase.
4. "deadline": Target deadline or "Ongoing".
5. "description": A clear 2-3 sentence Description & Breakdown Guide explaining what this project is, its deliverable, and how larger pieces should be split into 15-minute steps.
6. "checklist_items": 3 to 7 concrete, bite-sized 15-minute starter steps.

USER DESCRIPTION:
${scrubbed}`;

  const schema = {
    type: 'OBJECT',
    properties: {
      name: { type: 'STRING' },
      tier: {
        type: 'STRING',
        enum: ['Tier 1A - Teaching', 'Tier 1B - Consulting/PD', 'Tier 2 - AI Policy', 'Personal / Spanish & Ops', 'Tier 4 - Sandbox/Icebox']
      },
      phase: { type: 'STRING' },
      deadline: { type: 'STRING' },
      description: { type: 'STRING' },
      checklist_items: {
        type: 'ARRAY',
        items: { type: 'STRING' }
      }
    },
    required: ['name', 'tier', 'phase', 'deadline', 'description', 'checklist_items']
  };

  const parsed = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'create_project_card'
  );

  const restore = txt => (txt && typeof restorePII === 'function' ? restorePII(txt) : txt);
  return createProjectCard({
    name: restore(parsed.name),
    tier: parsed.tier,
    phase: restore(parsed.phase),
    deadline: restore(parsed.deadline),
    pacingRule: '15m weekday / 45m weekend chunks',
    description: restore(parsed.description),
    checklistItems: (parsed.checklist_items || []).map(i => restore(i))
  });
}

function toggleProjectChecklistItem(projectId, itemId, isDone) {
  const sheet = ensureProjectsSheet_();
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(projectId)) {
      let list = [];
      try { list = JSON.parse(data[i][7] || '[]'); } catch (e) {}
      list = isDone
        ? list.filter(item => String(item.id) !== String(itemId))
        : list.map(item => (String(item.id) === String(itemId) ? { ...item, done: false } : item));
      sheet.getRange(i + 1, 8).setValue(JSON.stringify(list));
      sheet.getRange(i + 1, 9).setValue(nowStr);
      return readProjectCards_(true);
    }
  }
  return readProjectCards_(true);
}

function addProjectChecklistItem(projectId, text) {
  const clean = String(text || '').trim();
  if (!clean) return readProjectCards_(true);

  const sheet = ensureProjectsSheet_();
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(projectId)) {
      let list = [];
      try { list = JSON.parse(data[i][7] || '[]'); } catch (e) {}
      list.push({
        id: 'CHK_' + Utilities.getUuid().slice(0, 6).toUpperCase(),
        text: clean,
        done: false,
        scheduledSlot: ''
      });
      sheet.getRange(i + 1, 8).setValue(JSON.stringify(list));
      sheet.getRange(i + 1, 9).setValue(nowStr);
      return readProjectCards_(true);
    }
  }
  return readProjectCards_(true);
}

function updateProjectMetadata(projectId, phase, deadline, pacingRule, notes, name) {
  const sheet = ensureProjectsSheet_();
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(projectId)) {
      if (name !== undefined && String(name).trim()) sheet.getRange(i + 1, 2).setValue(String(name).trim());
      if (phase !== undefined) sheet.getRange(i + 1, 4).setValue(String(phase).trim());
      if (deadline !== undefined) sheet.getRange(i + 1, 5).setValue(String(deadline).trim());
      if (pacingRule !== undefined) sheet.getRange(i + 1, 6).setValue(String(pacingRule).trim());
      if (notes !== undefined) sheet.getRange(i + 1, 7).setValue(String(notes).trim());
      sheet.getRange(i + 1, 9).setValue(nowStr);
      return readProjectCards_(true);
    }
  }
  return readProjectCards_(true);
}

/* ================= GMAIL TRIAGE WITH "ChiefOfStaff" PROCESSED LABEL ================= */

function getOrCreateCosLabel_() {
  return GmailApp.getUserLabelByName(COS_GMAIL_LABEL) || GmailApp.createLabel(COS_GMAIL_LABEL);
}

function extractThreadDigest_(thread, sourceLabel, userEmail) {
  const msgs = thread.getMessages();
  const recentMsgs = msgs.slice(-3);
  const latest = recentMsgs[recentMsgs.length - 1];
  const latestFrom = String(latest.getFrom() || '');
  const sentByUser = userEmail && latestFrom.toLowerCase().includes(userEmail.toLowerCase());

  const labels = (thread.getLabels() || []).map(l => l.getName()).join(', ');
  const flags = [
    thread.isUnread() ? 'UNREAD' : null,
    thread.hasStarredMessages() ? 'STARRED' : null,
    labels ? `LABELS: ${labels}` : null
  ].filter(Boolean).join(' | ');

  const threadBody = recentMsgs.map((m, idx) => {
    let plain = String(m.getPlainBody() || '');
    if (plain.trim().length < 25) {
      plain = String(m.getBody() || '')
        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
        .replace(/<[^>]+>/g, ' ');
    }
    plain = plain.replace(/\s+/g, ' ').trim().slice(0, idx === recentMsgs.length - 1 ? 900 : 450);
    const dStr = Utilities.formatDate(m.getDate(), PM_TZ, 'MMM d, h:mm a');
    return `  [Msg ${idx + 1}/${recentMsgs.length} • ${dStr} • From: ${m.getFrom()}]: ${plain}`;
  }).join('\n');

  return `THREAD_ID: ${thread.getId()}
SOURCE: [${sourceLabel}]
FLAGS: ${flags || 'INBOX'}
LAST_SENDER: ${sentByUser ? 'Joel (already replied last — only flag if a self-reminder/starred)' : latestFrom}
SUBJECT: ${thread.getFirstMessageSubject()}
CONVERSATION:
${threadBody}`;
}

function scanGmailForActionItems(customQuery) {
  const rawFilter = String(customQuery || '').trim();
  let threads = [];
  let queryUsed = rawFilter;
  const permalinkMap = {};

  try {
    if (rawFilter) {
      queryUsed = rawFilter;
      threads = GmailApp.search(rawFilter, 0, 30);
    } else {
      const primaryThreads = GmailApp.search(`in:inbox category:primary -label:${COS_GMAIL_LABEL}`, 0, 30);
      const importantThreads = GmailApp.search(
        `in:inbox (is:important OR is:starred) newer_than:14d -category:promotions -category:social -label:${COS_GMAIL_LABEL}`,
        0,
        15
      );

      queryUsed = `Last 30 Primary + Important/Starred (excluding ${COS_GMAIL_LABEL})`;
      const seenIds = new Set();
      [...primaryThreads, ...importantThreads].forEach(t => {
        const id = t.getId();
        if (!seenIds.has(id)) {
          seenIds.add(id);
          threads.push(t);
        }
      });
    }
  } catch (e) {
    throw new Error(
      'Gmail access not authorized yet. Confirm "https://www.googleapis.com/auth/gmail.modify" is in appsscript.json and re-authorize.'
    );
  }

  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail() || '';
  } catch (e) {}

  const emailDigest = (threads || []).map(t => {
    permalinkMap[t.getId()] = t.getPermalink();
    return extractThreadDigest_(t, 'Personal Gmail', userEmail);
  }).join('\n\n---\n\n');

  let schoolDigest = '';
  let schoolCount = 0;
  let bridgeStatusNote = '';

  if (
    typeof SCHOOL_BRIDGE_URL === 'string' &&
    SCHOOL_BRIDGE_URL.startsWith('http') &&
    !SCHOOL_BRIDGE_URL.includes('PASTE_COJOWA')
  ) {
    try {
      const fetchUrl = rawFilter
        ? `${SCHOOL_BRIDGE_URL}&q=${encodeURIComponent(rawFilter)}`
        : SCHOOL_BRIDGE_URL;
      const resp = UrlFetchApp.fetch(fetchUrl, { muteHttpExceptions: true });
      const rawBody = String(resp.getContentText() || '').trim();

      if (!rawBody.startsWith('{')) {
        bridgeStatusNote = 'COJOWA bridge returned a login page (school domain blocks external Web Apps)';
      } else {
        const parsedBridge = JSON.parse(rawBody);
        if (parsedBridge.error) {
          bridgeStatusNote = `COJOWA bridge error: ${parsedBridge.error}`;
        } else {
          const schoolThreads = parsedBridge.threads || [];
          schoolCount = schoolThreads.length;
          schoolDigest = schoolThreads.map(st => {
            permalinkMap[st.threadId] = st.permalink || ('https://mail.google.com/mail/#all/' + String(st.threadId).replace(/^SCHOOL_/, ''));
            const dStr = Utilities.formatDate(new Date(st.date), PM_TZ, 'MMM d, h:mm a');
            return `THREAD_ID: ${st.threadId}
SOURCE: [COJOWA School Email]
FLAGS: ${st.flags || 'INBOX'}
LAST_SENDER: ${st.from}
DATE: ${dStr}
SUBJECT: ${st.subject}
CONVERSATION:
  ${st.excerpt}`;
          }).join('\n\n---\n\n');
        }
      }
    } catch (e) {
      bridgeStatusNote = `COJOWA bridge unreachable: ${e.message}`;
    }
  }

  const scanSummary = `Scanned ${threads.length} untriaged personal + ${schoolCount} COJOWA threads`
    + (bridgeStatusNote ? ` (${bridgeStatusNote})` : '');

  const combinedDigest = [emailDigest, schoolDigest].filter(Boolean).join('\n\n---\n\n');
  if (!combinedDigest) {
    return {
      queryUsed: queryUsed,
      scanSummary: `${scanSummary} — 0 new action items.`,
      items: []
    };
  }

  const cards = readProjectCards_(true);
  const portfolioContext = cards.length > 0
    ? formatProjectsContextForPrompt_(cards)
    : '(No active project cards yet)';

  const scrubbed = typeof scrubPII === 'function' ? scrubPII(combinedDigest) : combinedDigest;

  const prompt = `Scan these untriaged email threads (from Joel's Personal Gmail and COJOWA School Email) and extract EVERY thread that requires a reply, decision, follow-up, deliverable, or action from him.

CAST A WIDE NET FOR REAL HUMAN & WORK ITEMS:
- Include direct questions, soft check-ins ("thoughts?", "when are you free?", "following up"), scheduling requests, student/parent/colleague inquiries, school admin deadlines, grad school / consulting correspondence, bills/payments due, and Google Docs/Drive comment notifications mentioning him.
- If a thread is STARRED or UNREAD from a real person/organization, treat it as likely actionable unless it is clearly a marketing newsletter or automated receipt.
- Look across ALL messages in CONVERSATION (not just the last line), because the actual ask is often in Msg 1 or Msg 2 when Msg 3 is a short bump.
- Ignore only true spam, promotional newsletters, and automated order/shipping receipts with nothing to do.

ACTIVE PROJECTS:
${portfolioContext}

EMAIL THREADS:
${scrubbed}

For each actionable email found:
- "thread_id": Exact THREAD_ID from above.
- "sender_short": Sender name/org, prefixed with "[COJOWA] " if SOURCE is [COJOWA School Email].
- "subject": Email subject.
- "action_title": Crisp 1-line action item starting with a verb.
- "deadline_hint": Mentioned deadline, "Starred", or "No hard deadline".
- "urgency": "HIGH", "MEDIUM", or "LOW".
- "matched_project_id": Best matching active project_id if it fits an active project, otherwise null.
- "suggested_category": Best category if added to Running To-Dos ("Teaching", "Consulting/PD", "Personal", or "Quick / Admin").`;

  const schema = {
    type: 'OBJECT',
    properties: {
      items: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            thread_id: { type: 'STRING' },
            sender_short: { type: 'STRING' },
            subject: { type: 'STRING' },
            action_title: { type: 'STRING' },
            deadline_hint: { type: 'STRING' },
            urgency: { type: 'STRING', enum: ['HIGH', 'MEDIUM', 'LOW'] },
            matched_project_id: { type: 'STRING', nullable: true },
            suggested_category: { type: 'STRING', enum: ['Teaching', 'Consulting/PD', 'Personal', 'Quick / Admin'] }
          },
          required: ['thread_id', 'sender_short', 'subject', 'action_title', 'deadline_hint', 'urgency', 'suggested_category']
        }
      }
    },
    required: ['items']
  };

  const parsed = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'scan_gmail_actions'
  );

  const restore = txt => (txt && typeof restorePII === 'function' ? restorePII(txt) : txt);
  const projectMap = {};
  cards.forEach(c => { projectMap[c.projectId] = c.name; });

  const items = (parsed.items || []).map(it => {
    const pid = it.matched_project_id && projectMap[it.matched_project_id] ? it.matched_project_id : '';
    return {
      threadId: it.thread_id,
      permalink: permalinkMap[it.thread_id] || '',
      senderShort: restore(it.sender_short),
      subject: restore(it.subject),
      actionTitle: restore(it.action_title),
      deadlineHint: restore(it.deadline_hint),
      urgency: it.urgency,
      matchedProjectId: pid,
      matchedProjectName: pid ? projectMap[pid] : '',
      suggestedCategory: it.suggested_category || 'Quick / Admin'
    };
  });

  return {
    queryUsed: queryUsed,
    scanSummary: `${scanSummary} — ${items.length} to review.`,
    items: items
  };
}

/**
 * Applies the "ChiefOfStaff" label (so it never resurfaces in Chief of Staff), archives, or trashes a thread.
 * action: 'tag' (keep in Gmail inbox, mark triaged), 'archive' (tag + archive), or 'trash' (move to Trash).
 */
function performGmailThreadAction(threadId, action) {
  const idStr = String(threadId || '').trim();
  if (!idStr) return { ok: false };

  if (idStr.startsWith('SCHOOL_')) {
    if (
      typeof SCHOOL_BRIDGE_URL === 'string' &&
      SCHOOL_BRIDGE_URL.startsWith('http') &&
      !SCHOOL_BRIDGE_URL.includes('PASTE_COJOWA')
    ) {
      const url = `${SCHOOL_BRIDGE_URL}&action=${encodeURIComponent(action)}&threadId=${encodeURIComponent(idStr)}`;
      UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    }
    return {
      ok: true,
      message: action === 'trash'
        ? 'Deleted in COJOWA Gmail.'
        : (action === 'archive' ? 'Archived in COJOWA Gmail.' : `Tagged "${COS_GMAIL_LABEL}" (stays in COJOWA inbox, won't reappear).`)
    };
  }

  const thread = GmailApp.getThreadById(idStr);
  if (thread) {
    if (action === 'trash') {
      thread.moveToTrash();
    } else {
      const label = getOrCreateCosLabel_();
      thread.addLabel(label);
      if (action === 'archive') {
        thread.markRead();
        thread.moveToArchive();
      }
    }
  }
  return {
    ok: true,
    message: action === 'trash'
      ? 'Moved to Trash in Gmail.'
      : (action === 'archive' ? 'Archived & tagged in Gmail.' : `Tagged "${COS_GMAIL_LABEL}" (kept in inbox, won't reappear).`)
  };
}

function bulkTagGmailThreads(threadIds) {
  const list = Array.isArray(threadIds) ? threadIds : [];
  let count = 0;
  list.forEach(id => {
    try {
      performGmailThreadAction(id, 'tag');
      count++;
    } catch (e) {}
  });
  return {
    ok: true,
    count: count,
    message: `Tagged ${count} email${count === 1 ? '' : 's'} with "${COS_GMAIL_LABEL}" — kept in Gmail inbox, won't reappear here.`
  };
}

/**
 * Adds a Gmail action item to a Project or Running To-Do and tags the email "ChiefOfStaff"
 * so it stays visible in Gmail while never re-appearing in the Chief of Staff scanner.
 */
function importGmailItem(actionType, projectId, itemText, category, threadId) {
  const cleanText = String(itemText || '').trim();
  if (!cleanText) throw new Error('Action item text is required.');

  if (threadId) {
    try {
      performGmailThreadAction(threadId, 'tag');
    } catch (e) {}
  }

  if (actionType === 'ADD_TO_PROJECT' && projectId) {
    const cards = addProjectChecklistItem(projectId, cleanText);
    const target = cards.find(c => c.projectId === String(projectId));
    return {
      message: `Added to "${target ? target.name : 'Project'}" & tagged "${COS_GMAIL_LABEL}"`,
      cards: cards,
      runningTodos: getRunningTodos()
    };
  }

  const todos = addRunningTodo(cleanText, category || 'Quick / Admin', 15, false);
  return {
    message: `Added to Running To-Dos & tagged "${COS_GMAIL_LABEL}"`,
    cards: readProjectCards_(true),
    runningTodos: todos
  };
}

function syncCompletedTaskToProjectCard_(taskTitle, projectIdHint) {
  if (!taskTitle) return;
  const cleanTarget = String(taskTitle)
    .replace(/^\[Focus\]\s*/i, '')
    .replace(/\s*\(\d+\/\d+\)\s*$/, '')
    .trim()
    .toLowerCase();

  const sheet = ensureProjectsSheet_();
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), PM_TZ, 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9] || 'Active').trim() === 'Archived') continue;
    const rowProjId = String(data[i][0]);
    if (projectIdHint && rowProjId !== String(projectIdHint)) continue;

    let list = [];
    try { list = JSON.parse(data[i][7] || '[]'); } catch (e) {}
    const filtered = list.filter(item => {
      const itemLow = String(item.text || '').toLowerCase();
      return !(itemLow === cleanTarget || itemLow.includes(cleanTarget) || cleanTarget.includes(itemLow));
    });

    if (filtered.length !== list.length) {
      sheet.getRange(i + 1, 8).setValue(JSON.stringify(filtered));
      sheet.getRange(i + 1, 9).setValue(nowStr);
    }
  }
}

function getCompletionAndScheduleContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  const doneTasks = [];
  const scheduledTasks = [];

  if (queueSheet && queueSheet.getLastRow() > 1) {
    const rows = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, 11).getValues();
    rows.forEach(r => {
      const title = String(r[2] || '').trim();
      const status = String(r[4] || '').trim();
      const slot = String(r[9] || '').trim();
      if (!title) return;
      if (status === 'Done') doneTasks.push(title);
      if (status === 'Scheduled') scheduledTasks.push(`${title} (${slot})`);
    });
  }

  return {
    doneTasks: doneTasks.slice(-25),
    scheduledTasks: scheduledTasks.slice(-20),
    calendarEvents: []
  };
}

function formatProjectsContextForPrompt_(cards) {
  return cards.map(c => {
    const openItems = c.checklist.filter(i => !i.done).map(i => `  [ ] (${i.id}) ${i.text}`).join('\n') || '  (No items yet)';
    return `PROJECT [${c.projectId}]: ${c.name} (${c.tier}) | Deadline: ${c.deadline}
- Description & Breakdown Guide: ${c.notes}
- Open To-Dos:\n${openItems}`;
  }).join('\n\n');
}