/**
 * Phase 3: Cross-Device Web App + Desktop Sidebar Orchestrator,
 * Structured Gemini Engine, Calendar Writer (7:00 PM Cutoff), & Daily Triggers.
 */

const TIMEZONE = 'America/Bogota';
const HARD_CUTOFF_HOUR = 19; // 7:00 PM

/**
 * Serves the cross-device Web App interface (Phone, Tablet, Desktop Browser).
 */
function doGet() {
return HtmlService.createHtmlOutputFromFile('ControlCenter')
  .setTitle('Chief of Staff')
  .setFaviconUrl('https://img.icons8.com/fluency/96/compass.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Adds desktop menu inside Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Chief of Staff')
    .addItem('Open Control Center Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Run Morning Brief Now', 'runMorningDigest')
    .addItem('Run Evening Shutdown Now', 'runEveningShutdown')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ControlCenter')
    .setTitle('Chief of Staff')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * One-click setup: creates Task_Queue, Parameters, and Audit_Log tabs
 * with frozen headers and your active sprint locks.
 */
function setupHubTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = {
    Task_Queue: [
      'Task_ID', 'Created_At', 'Title', 'Tier', 'Status',
      'TMT_Bottleneck', 'Two_Min_Start', 'If_Then_Cue',
      'Next_Step_Park', 'Scheduled_Block', 'Calendar_Event_ID'
    ],
    Parameters: ['Parameter_Key', 'Value', 'Notes'],
    Audit_Log: ['Timestamp', 'Mode', 'Model_ID', 'HTTP_Status', 'Est_Tokens', 'Verdict']
  };

  Object.entries(schemas).forEach(([tabName, headers]) => {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  const paramSheet = ss.getSheetByName('Parameters');
  if (paramSheet.getLastRow() < 2) {
    const defaults = [
      ['Tier_1A_Quota', '4 async instructional videos/week (plan periods + Sunday)', 'COJOWA physics/science async rollout'],
      ['Tier_1B_Sprint', 'Finalize Workload Automation Lab prototype + AI Student Profile Creator PD prep', 'PD cycle starts this month'],
      ['Tier_2_Research', 'School AI Policy (3 pillars + 5-color Zotero Better Notes)', 'Background PDF triage only'],
      ['Hard_Cutoff', '19:00 (7:00 PM) daily — zero work after cutoff', 'Protects evening recovery and time with Sarah'],
      ['Gatekeeper_Strictness', '8.5 / 10 — block mid-week tangents if Tier 1A/1B behind', 'Enforces Cost of Yes protocol']
    ];
    paramSheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }

  Logger.log('Hub tabs initialized successfully.');
}

/**
 * Evaluates a brain dump, stalled task, or tangent against Charter.gs
 * using structured JSON outputs and local PII scrubbing.
 */
function evaluateInput(rawInput, mode) {
  if (!rawInput || !rawInput.trim()) {
    throw new Error('Enter a task, idea, or blocker first.');
  }

  const scrubbedText = typeof scrubPII === 'function' ? scrubPII(rawInput.trim()) : rawInput.trim();
  const modelId = mode === 'devils_advocate' ? 'gemini-4-argon' : 'gemini-3.8-flash';

  const modeDirectives = {
    gatekeeper: 'Evaluate this incoming idea or task with 8.5/10 Gatekeeper strictness. Decide whether to APPROVE_NOW, DEFER_SUNDAY, or ICEBOX. If it is a tangent while Tier 1A (4 async videos) or Tier 1B (Workload Automation Lab prototype/PD) needs focus, push back directly and expose hidden assumptions.',
    unstuck: 'Diagnose why Joel is stalling on this task using Temporal Motivation Theory (Expectancy, Value, Impulsiveness, or Delay) with zero shame. Provide one concrete 2-minute entry action, an 85% good-enough stopping rule, and one If-Then implementation intention.',
    park: 'Close this cognitive loop so Joel can switch tasks cleanly without attention residue. Summarize exact stop-state, assign Tier, and define the exact next 2-minute physical action when reopened.',
    devils_advocate: 'Run a rigorous Devil\'s Advocate pre-mortem critique. List 3 concrete failure modes, hidden technical/time debt, and challenge the core premise directly.'
  };

  const userPrompt = `MODE: ${mode.toUpperCase()}\nDIRECTIVE: ${modeDirectives[mode] || modeDirectives.gatekeeper}\n\nUSER INPUT:\n${scrubbedText}`;

  const schema = {
    type: 'OBJECT',
    properties: {
      clean_title: { type: 'STRING', description: 'Concise 4-to-8 word task title.' },
      tier: {
        type: 'STRING',
        enum: ['Tier 1A - Teaching', 'Tier 1B - Consulting/PD', 'Tier 2 - AI Policy', 'Tier 3 - UIUC (Frozen)', 'Tier 4 - Sandbox/Icebox'],
        description: 'Portfolio tier classification.'
      },
      verdict: {
        type: 'STRING',
        enum: ['APPROVE_NOW', 'DEFER_SUNDAY', 'ICEBOX', 'UNSTUCK_READY', 'LOOP_PARKED'],
        description: 'Routing decision.'
      },
      coach_response: {
        type: 'STRING',
        description: 'Direct, opinionated Chief of Staff response (under 140 words). If pushing back, name the avoided Tier 1 task and ask what is cut from Sunday.'
      },
      hidden_assumptions: {
        type: 'STRING',
        nullable: true,
        description: 'Bullet list of 2-3 hidden assumptions or failure risks, or null if trivial clerical task.'
      },
      tmt_bottleneck: {
        type: 'STRING',
        enum: ['Low Expectancy', 'Low Value', 'High Impulsiveness', 'High Delay', 'None'],
        description: 'Primary Temporal Motivation Theory friction point.'
      },
      two_min_start: {
        type: 'STRING',
        description: 'Single concrete physical action taking 2 minutes or less to begin.'
      },
      if_then_cue: {
        type: 'STRING',
        nullable: true,
        description: 'Implementation intention formatted as "If [specific cue], then I will [concrete action]", or null.'
      },
      suggested_duration_min: {
        type: 'INTEGER',
        description: 'Recommended capped focus duration in minutes (15, 30, 45, 60, or 90).'
      }
    },
    required: ['clean_title', 'tier', 'verdict', 'coach_response', 'tmt_bottleneck', 'two_min_start', 'suggested_duration_min']
  };

  const aiResult = callGeminiStructured_(modelId, getSystemCharter(), userPrompt, schema, mode);

  const restore = (val) => (val && typeof restorePII === 'function' ? restorePII(val) : val);
  aiResult.clean_title = restore(aiResult.clean_title);
  aiResult.coach_response = restore(aiResult.coach_response);
  aiResult.hidden_assumptions = restore(aiResult.hidden_assumptions);
  aiResult.two_min_start = restore(aiResult.two_min_start);
  aiResult.if_then_cue = restore(aiResult.if_then_cue);

  const taskId = 'TSK_' + Utilities.getUuid().slice(0, 6).toUpperCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  const status = aiResult.verdict === 'ICEBOX' ? 'Icebox' : (aiResult.verdict === 'LOOP_PARKED' ? 'Parked' : 'Queued');

  queueSheet.appendRow([
    taskId,
    Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'),
    aiResult.clean_title,
    aiResult.tier,
    status,
    aiResult.tmt_bottleneck,
    aiResult.two_min_start,
    aiResult.if_then_cue || '',
    aiResult.verdict === 'LOOP_PARKED' ? aiResult.two_min_start : '',
    '',
    ''
  ]);

  return { taskId, ...aiResult };
}

/**
 * Returns active Queued, Parked, and Scheduled tasks as plain strings
 * so the mobile/desktop Web App can display and manage open loops.
 */
function getActiveQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Task_Queue');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getDataRange().getValues().slice(1);
  return rows
    .filter(r => ['Queued', 'Parked', 'Scheduled'].includes(String(r[4])))
    .reverse()
    .slice(0, 8)
    .map(r => ({
      taskId: String(r[0]),
      title: String(r[2]),
      tier: String(r[3]),
      status: String(r[4]),
      twoMinStart: String(r[6]),
      scheduledBlock: String(r[9] || '')
    }));
}

function markTaskDone(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Task_Queue');
  if (!sheet || sheet.getLastRow() < 2) return false;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(taskId)) {
      const taskTitle = String(data[i][2] || '').trim();
      const rawHint = String(data[i][8] || '').trim();
      const projectIdHint = rawHint.startsWith('PRJ_') ? rawHint : '';

      sheet.getRange(i + 1, 5).setValue('Done');

      if (typeof syncCompletedTaskToProjectCard_ === 'function') {
        syncCompletedTaskToProjectCard_(taskTitle, projectIdHint);
      }
      return true;
    }
  }
  return false;
}

function scheduleTaskToCalendar(taskId, dateStr, startTimeStr, durationMin) {
  if (!dateStr || !startTimeStr) {
    throw new Error('Select both a date and start time.');
  }

  const duration = Number(durationMin) || 45;
  const startIso = `${dateStr}T${startTimeStr}:00-05:00`;
  const startDate = new Date(startIso);
  const endDate = new Date(startDate.getTime() + duration * 60000);

  const startHour = Number(Utilities.formatDate(startDate, TIMEZONE, 'H'));
  const endHour = Number(Utilities.formatDate(endDate, TIMEZONE, 'H'));
  const endMin = Number(Utilities.formatDate(endDate, TIMEZONE, 'm'));

  if (startHour >= HARD_CUTOFF_HOUR || endHour > HARD_CUTOFF_HOUR || (endHour === HARD_CUTOFF_HOUR && endMin > 0)) {
    throw new Error('Blocked by Charter: Focus blocks cannot run at or past your 7:00 PM hard work cutoff.');
  }

  const calendars = typeof getIncludedCalendars_ === 'function'
    ? getIncludedCalendars_()
    : [{ id: 'primary' }];
  const freeRegex = typeof getFreePeriodRegex_ === 'function'
    ? getFreePeriodRegex_()
    : /\bplan\b/i;

  for (const cal of calendars) {
    try {
      const existing = Calendar.Events.list(cal.id, {
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true
      });
      const hit = (existing.items || []).find(ev => {
        if (!ev.start || !ev.start.dateTime) return false;
        if (ev.transparency === 'transparent') return false;
        if (freeRegex.test(ev.summary || '')) return false;
        return true;
      });
      if (hit) {
        const busyTitle = hit.summary || 'an existing event';
        throw new Error(`Conflict detected with "${busyTitle}". Pick an open window.`);
      }
    } catch (e) {
      if (e.message && e.message.startsWith('Conflict detected')) throw e;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  const data = queueSheet.getDataRange().getValues();
  let rowIndex = -1;
  let taskRow = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(taskId)) {
      rowIndex = i + 1;
      taskRow = data[i];
      break;
    }
  }

  if (!taskRow) throw new Error('Task not found in Task_Queue.');

  const [, , title, tier, , , twoMinStart, ifThenCue] = taskRow;
  const colorMap = {
    'Tier 1A - Teaching': '9',
    'Tier 1B - Consulting/PD': '10',
    'Tier 2 - AI Policy': '3',
    'Personal / Spanish & Ops': '6',
    'Tier 4 - Sandbox/Icebox': '5'
  };

  const eventResource = {
    summary: `[Focus] ${title}`,
    description: `Tier: ${tier}\n2-Min Entry Action: ${twoMinStart}\n${ifThenCue ? 'Trigger: ' + ifThenCue : ''}`,
    start: { dateTime: startDate.toISOString(), timeZone: TIMEZONE },
    end: { dateTime: endDate.toISOString(), timeZone: TIMEZONE },
    colorId: colorMap[tier] || '9',
    reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 5 }] }
  };

  const created = Calendar.Events.insert(eventResource, 'primary');
  const formattedSlot = Utilities.formatDate(startDate, TIMEZONE, 'EEE MMM d, h:mm a');

  queueSheet.getRange(rowIndex, 5).setValue('Scheduled');
  queueSheet.getRange(rowIndex, 10).setValue(formattedSlot);
  queueSheet.getRange(rowIndex, 11).setValue(created.id);

  return `Scheduled "${title}" for ${formattedSlot} (${duration} min).`;
}

/**
 * Morning Activation Digest (runs at 6:45 AM): sends one clean single-channel brief.
 */
function runMorningDigest() {
  const active = getActiveQueue();
  if (active.length === 0) return;

  const summaryList = active.map(t =>
    `- [${t.tier}] ${t.title} (${t.status}${t.scheduledBlock ? ' @ ' + t.scheduledBlock : ''}) -> 2-Min Start: ${t.twoMinStart}`
  ).join('\n');

  const prompt = `Generate a crisp, low-friction Morning Activation Brief (under 150 words) for Joel based on his active queue below. Highlight his top 3 priorities and remind him of his 7:00 PM hard cutoff.\n\nQUEUE:\n${summaryList}`;

  const schema = {
    type: 'OBJECT',
    properties: {
      subject: { type: 'STRING', description: 'Short email subject line.' },
      body: { type: 'STRING', description: 'Plain-text morning brief.' }
    },
    required: ['subject', 'body']
  };

  const brief = callGeminiStructured_('gemini-3.5-flash-lite', getSystemCharter(), prompt, schema, 'morning_brief');
  if (MailApp.getRemainingDailyQuota() > 0) {
    MailApp.sendEmail(Session.getActiveUser().getEmail(), brief.subject, brief.body);
  }
}

/**
 * Evening Shutdown Ritual (runs at 6:45 PM before 7:00 PM cutoff):
 * Closes open cognitive loops and sets tomorrow's Next_Step_Park.
 */
function runEveningShutdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  if (!queueSheet || queueSheet.getLastRow() < 2) return;

  const range = queueSheet.getDataRange();
  const values = range.getValues();
  let parkedCount = 0;

  for (let i = 1; i < values.length; i++) {
    const status = String(values[i][4]);
    if (status === 'Scheduled' || status === 'In_Progress') {
      values[i][4] = 'Parked';
      if (!values[i][8]) {
        values[i][8] = values[i][6];
      }
      parkedCount++;
    }
  }

  if (parkedCount > 0) {
    range.setValues(values);
  }
  Logger.log(`Evening shutdown complete. Closed ${parkedCount} open loop(s) before 7:00 PM cutoff.`);
}

/**
 * Registers the 6:45 AM Morning Brief and 6:45 PM Evening Shutdown triggers.
 */
function installDailyTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'runMorningDigest' || fn === 'runEveningShutdown') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runMorningDigest')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(45)
    .create();

  ScriptApp.newTrigger('runEveningShutdown')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .nearMinute(45)
    .create();

  Logger.log('Daily triggers installed (6:45 AM Morning Brief & 6:45 PM Shutdown).');
}

/**
 * Core Gemini REST wrapper with exponential backoff on HTTP 429 and Audit_Log tracking.
 */
function callGeminiStructured_(modelId, systemInstruction, userPrompt, responseSchema, modeLabel) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Missing GEMINI_API_KEY in Script Properties.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;
  const payload = {
    systemInstruction: { parts: [{ text: systemInstruction }] },
    contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
    generationConfig: {
      temperature: 0.25,
      responseMimeType: 'application/json',
      responseSchema: responseSchema
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response;
  let status = 0;
  for (let attempt = 0; attempt < 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    status = response.getResponseCode();
    if (status === 200) break;
    if (status === 429 || status >= 500) {
      Utilities.sleep(Math.pow(2, attempt) * 1500);
      continue;
    }
    break;
  }

  const rawBody = response.getContentText();
  const parsedBody = JSON.parse(rawBody);
  const estTokens = parsedBody.usageMetadata ? parsedBody.usageMetadata.totalTokenCount : 0;

  const auditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Log');
  let verdictLog = status === 200 ? 'OK' : 'ERROR';

  if (status !== 200) {
    if (auditSheet) {
      auditSheet.appendRow([
        Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
        modeLabel, modelId, status, estTokens, verdictLog
      ]);
    }
    throw new Error(`Gemini API Error (${status}): ${parsedBody.error ? parsedBody.error.message : rawBody}`);
  }

  const jsonText = parsedBody.candidates[0].content.parts[0].text;
  const structured = JSON.parse(jsonText);
  verdictLog = structured.verdict || 'OK';

  if (auditSheet) {
    auditSheet.appendRow([
      Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      modeLabel, modelId, status, estTokens, verdictLog
    ]);
  }

  return structured;
}

/**
 * Fetches Today + Next 2 Days of events across Joel's active calendars
 * (including COJOWA classes and "Period 2 (Plan)" windows) and cross-links
 * them against Task_Queue so tracked items are flagged automatically.
 */
function getCalendarAgenda(daysAhead) {
  const horizon = Math.min(Math.max(Number(daysAhead) || 3, 1), 7);
  const now = new Date();
  const todayStr = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  const startWindow = new Date(`${todayStr}T00:00:00-05:00`);
  const endWindow = new Date(startWindow.getTime() + horizon * 86400000);

  // Map existing Calendar_Event_IDs in Task_Queue to avoid duplicate pulls
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  const linkedByEventId = {};
  if (queueSheet && queueSheet.getLastRow() > 1) {
    const qRows = queueSheet.getDataRange().getValues().slice(1);
    qRows.forEach(r => {
      const evId = String(r[10] || '').trim();
      if (evId) {
        linkedByEventId[evId] = {
          taskId: String(r[0]),
          status: String(r[4])
        };
      }
    });
  }

  const calendars = typeof getIncludedCalendars_ === 'function'
    ? getIncludedCalendars_()
    : [{ id: 'primary', summary: 'Primary', primary: true }];

  const freeRegex = typeof getFreePeriodRegex_ === 'function'
    ? getFreePeriodRegex_()
    : /\bplan\b/i;

  const events = [];

  calendars.forEach(cal => {
    try {
      const res = Calendar.Events.list(cal.id, {
        timeMin: startWindow.toISOString(),
        timeMax: endWindow.toISOString(),
        singleEvents: true,
        orderBy: 'startTime'
      });

      (res.items || []).forEach(ev => {
        const title = String(ev.summary || '(Untitled Event)');
        const isAllDay = Boolean(ev.start && ev.start.date && !ev.start.dateTime);

        let startDate, endDate, dateStr, dayLabel, timeLabel, startTime24, durationMin;

        if (isAllDay) {
          dateStr = ev.start.date;
          startDate = new Date(`${dateStr}T00:00:00-05:00`);
          endDate = new Date(`${dateStr}T23:59:59-05:00`);
          dayLabel = Utilities.formatDate(startDate, TIMEZONE, 'EEE, MMM d');
          timeLabel = 'All Day';
          startTime24 = '09:00';
          durationMin = 45;
        } else {
          startDate = new Date(ev.start.dateTime);
          endDate = new Date(ev.end.dateTime);
          dateStr = Utilities.formatDate(startDate, TIMEZONE, 'yyyy-MM-dd');
          dayLabel = Utilities.formatDate(startDate, TIMEZONE, 'EEE, MMM d');
          const sFmt = Utilities.formatDate(startDate, TIMEZONE, 'h:mm a');
          const eFmt = Utilities.formatDate(endDate, TIMEZONE, 'h:mm a');
          timeLabel = `${sFmt} – ${eFmt}`;
          startTime24 = Utilities.formatDate(startDate, TIMEZONE, 'HH:mm');
          durationMin = Math.max(Math.round((endDate.getTime() - startDate.getTime()) / 60000), 15);
        }

        const isPlanPeriod = freeRegex.test(title);
        const isFocusBlock = title.startsWith('[Focus]') || title.startsWith('[Protected]');
        const linked = linkedByEventId[ev.id] || null;

        events.push({
          eventId: String(ev.id),
          title: title,
          calendarName: cal.primary ? 'Personal' : String(cal.summary || 'School'),
          dateStr: dateStr,
          dayLabel: dayLabel,
          timeLabel: timeLabel,
          startTime24: startTime24,
          durationMin: durationMin,
          isAllDay: isAllDay,
          isPast: !isAllDay && endDate.getTime() < now.getTime(),
          isPlanPeriod: isPlanPeriod,
          isFocusBlock: isFocusBlock,
          isTracked: Boolean(linked),
          linkedStatus: linked ? linked.status : '',
          sortMs: startDate.getTime()
        });
      });
    } catch (e) {
      // Skip inaccessible secondary calendars
    }
  });

  events.sort((a, b) => a.sortMs - b.sortMs);
  return events.slice(0, 30);
}

/**
 * Pulls an existing Google Calendar event (meeting, deadline, or prep task)
 * into Task_Queue so it appears in Active Queue & Scheduled Blocks.
 */
function pullCalendarEventToQueue(ev) {
  if (!ev || !ev.eventId) {
    throw new Error('Invalid calendar event.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  if (!queueSheet) throw new Error('Missing Task_Queue sheet.');

  // Prevent duplicate imports if this Calendar_Event_ID is already in Task_Queue
  if (queueSheet.getLastRow() > 1) {
    const data = queueSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][10] || '').trim() === String(ev.eventId).trim()) {
        if (String(data[i][4]) === 'Done') {
          queueSheet.getRange(i + 1, 5).setValue('Scheduled');
        }
        return `Already tracked in Active Queue (${data[i][0]}).`;
      }
    }
  }

  const cleanTitle = String(ev.title || 'Calendar Event')
    .replace(/^\[(Focus|Protected)\]\s*/i, '')
    .trim();

  // Lightweight tier classification based on Joel's portfolio domains
  const lower = cleanTitle.toLowerCase();
  let tier = 'Tier 1A - Teaching';
  if (/(workload|automation|césar|cesar|profile|add-on|pd\b|workshop|consulting)/i.test(lower)) {
    tier = 'Tier 1B - Consulting/PD';
  } else if (/(policy|zotero|ethics|paper|literature|vericite)/i.test(lower)) {
    tier = 'Tier 2 - AI Policy';
  } else if (/(dj|mix|ableton|music|garden|monticello|goose creek)/i.test(lower)) {
    tier = 'Tier 4 - Sandbox/Icebox';
  }

  const status = ev.isAllDay ? 'Queued' : 'Scheduled';
  const scheduledLabel = ev.isAllDay ? '' : `${ev.dayLabel}, ${ev.timeLabel.split('–')[0].trim()}`;
  const taskId = 'TSK_' + Utilities.getUuid().slice(0, 6).toUpperCase();

  queueSheet.appendRow([
    taskId,
    Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'),
    cleanTitle,
    tier,
    status,
    'None',
    'Open event notes/materials 2 minutes before start.',
    '',
    '',
    scheduledLabel,
    String(ev.eventId)
  ]);

  return `Pulled "${cleanTitle}" into Active Queue (${taskId}).`;
}