/**
 * Phase 5: Weekly Friday Calibration Audit, Failure-Mode Stress Tests,
 * and Google Workspace Studio Webhook Endpoint (doPost).
 */

const CAL_TIMEZONE = 'America/Bogota';

/**
 * Generates a secret token in Script Properties for authenticating
 * incoming Google Workspace Studio webhook requests.
 */
function generateWebhookSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('WEBHOOK_SECRET');
  if (!secret) {
    secret = 'cos_wh_' + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('WEBHOOK_SECRET', secret);
  }
  Logger.log(`WEBHOOK_SECRET ready: ${secret}`);
  return secret;
}

/**
 * Receives incoming JSON payloads from Google Workspace Studio Flows
 * (e.g., incoming emails from César, Google Form PD feedback, or Drive alerts)
 * and routes them through evaluateInput() with local PII scrubbing.
 */
function doPost(e) {
  try {
    const expectedSecret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
    if (!expectedSecret) {
      return jsonResponse_({ ok: false, error: 'Server WEBHOOK_SECRET not configured.' });
    }

    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const providedSecret = body.secret || (e && e.parameter && e.parameter.secret);

    if (providedSecret !== expectedSecret) {
      return jsonResponse_({ ok: false, error: 'Unauthorized: invalid webhook secret.' });
    }

    const rawText = [body.source ? `` : '', body.text || body.summary || '']
      .filter(Boolean)
      .join(' ')
      .trim();

    if (!rawText) {
      return jsonResponse_({ ok: false, error: 'Missing text payload.' });
    }

    const mode = ['gatekeeper', 'unstuck', 'park', 'devils_advocate'].includes(body.mode)
      ? body.mode
      : 'gatekeeper';

    const result = evaluateInput(rawText, mode);
    return jsonResponse_({ ok: true, result });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Runs every Friday at 2:30 PM (during Joel's lighter Friday admin window)
 * to audit Tier 1A/1B execution, parked loops, API token usage, and Gatekeeper calibration.
 */
function runWeeklyCalibrationAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  const auditSheet = ss.getSheetByName('Audit_Log');

  const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const queueRows = queueSheet && queueSheet.getLastRow() > 1
    ? queueSheet.getDataRange().getValues().slice(1)
    : [];

  const recentTasks = queueRows.filter(r => {
    const dt = new Date(r[1]);
    return !isNaN(dt.getTime()) && dt >= sevenDaysAgo;
  });

  const doneCount = recentTasks.filter(r => r[4] === 'Done').length;
  const parkedCount = recentTasks.filter(r => r[4] === 'Parked').length;
  const iceboxCount = recentTasks.filter(r => r[4] === 'Icebox').length;
  const tier1ACount = recentTasks.filter(r => String(r[3]).startsWith('Tier 1A')).length;
  const tier1BCount = recentTasks.filter(r => String(r[3]).startsWith('Tier 1B')).length;

  const auditRows = auditSheet && auditSheet.getLastRow() > 1
    ? auditSheet.getDataRange().getValues().slice(1)
    : [];

  const recentCalls = auditRows.filter(r => {
    const dt = new Date(r[0]);
    return !isNaN(dt.getTime()) && dt >= sevenDaysAgo;
  });

  const totalTokens = recentCalls.reduce((sum, r) => sum + (Number(r[4]) || 0), 0);
  const errorCalls = recentCalls.filter(r => Number(r[3]) !== 200).length;

  const openLoops = queueRows
    .filter(r => ['Queued', 'Parked', 'Scheduled'].includes(String(r[4])))
    .slice(-10)
    .map(r => `- [${r[3]}] ${r[2]} (${r[4]}, Bottleneck: ${r[5]})`)
    .join('\n') || 'None';

  const prompt = `Conduct Joel's Weekly Friday Calibration Audit (under 220 words).

WEEKLY METRICS (LAST 7 DAYS):
- Tasks Logged: ${recentTasks.length} (Done: ${doneCount}, Parked: ${parkedCount}, Icebox: ${iceboxCount})
- Tier 1A (COJOWA Async Videos — target 4/wk): ${tier1ACount} logged
- Tier 1B (Workload Automation Lab / PD): ${tier1BCount} logged
- Gemini API Calls: ${recentCalls.length} | Total Tokens: ${totalTokens} | Errors: ${errorCalls}

CURRENT OPEN LOOPS:
${openLoops}

INSTRUCTIONS:
1. Assess honestly whether Tier 1A (4 weekly videos) and Tier 1B (prototype/PD) stayed protected or got crowded out.
2. Identify the #1 recurring TMT bottleneck this week and recommend one concrete adjustment for next week.
3. Recommend whether Gatekeeper strictness (currently 8.5/10) should hold, tighten, or loosen.`;

  const schema = {
    type: 'OBJECT',
    properties: {
      subject: { type: 'STRING', description: 'Friday Calibration email subject line.' },
      audit_report: { type: 'STRING', description: 'Concise plain-text weekly calibration report.' },
      recommended_strictness: { type: 'STRING', description: 'Recommended Gatekeeper strictness setting for next week.' }
    },
    required: ['subject', 'audit_report', 'recommended_strictness']
  };

  const report = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'weekly_calibration'
  );

  if (MailApp.getRemainingDailyQuota() > 0) {
    MailApp.sendEmail(Session.getActiveUser().getEmail(), report.subject, report.audit_report);
  }

  Logger.log(`Weekly calibration complete. Recommended strictness: ${report.recommended_strictness}`);
}

/**
 * Installs the weekly Friday 2:30 PM Calibration Audit trigger.
 */
function installFridayCalibrationTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runWeeklyCalibrationAudit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runWeeklyCalibrationAudit')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(14)
    .nearMinute(30)
    .create();

  Logger.log('Friday 2:30 PM Weekly Calibration trigger installed.');
}

/**
 * Non-destructive 4-point failure-mode verification suite:
 * 1. API Key & Charter load check
 * 2. 7:00 PM hard work cutoff enforcement check
 * 3. Webhook unauthorized request rejection check
 * 4. PII regex escaping check
 */
function runFailureModeTests() {
  const results = [];

  // Test 1: API Key & Charter
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const charter = getSystemCharter();
  results.push(apiKey && charter.length > 500
    ? 'PASS [1/4]: GEMINI_API_KEY present and Charter loaded.'
    : 'FAIL [1/4]: Missing GEMINI_API_KEY or Charter.');

  // Test 2: 7:00 PM Hard Cutoff Guardrail
  try {
    scheduleTaskToCalendar('TSK_DUMMY', '2026-11-20', '19:15', 30);
    results.push('FAIL [2/4]: 7:15 PM block was NOT blocked.');
  } catch (e) {
    results.push(e.message.includes('7:00 PM')
      ? 'PASS [2/4]: 7:00 PM hard cutoff blocked post-cutoff scheduling.'
      : `FAIL [2/4]: Unexpected error: ${e.message}`);
  }

  // Test 3: Webhook Unauthorized Rejection
  generateWebhookSecret();
  const badWebhook = doPost({
    postData: { contents: JSON.stringify({ secret: 'wrong_secret', text: 'Test tangent' }) }
  });
  const parsedWebhook = JSON.parse(badWebhook.getContent());
  results.push(parsedWebhook.ok === false && parsedWebhook.error.includes('Unauthorized')
    ? 'PASS [3/4]: Unauthorized webhook rejected cleanly.'
    : 'FAIL [3/4]: Webhook did not reject invalid secret.');

  // Test 4: PII Regex Special-Character Safety
  const escaped = escapeRegex_('Student (Grade 10) + Test@school.edu');
  results.push(escaped.includes('\\(') && escaped.includes('\\+')
    ? 'PASS [4/4]: PII regex sanitizer escapes special characters safely.'
    : 'FAIL [4/4]: Regex escaping failed.');

  Logger.log('\n' + results.join('\n'));
}