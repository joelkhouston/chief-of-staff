/**
 * AutoScheduler.gs (v8 — Project-Card Aware, Zero-Repeat Suggestions & Default-Free Canvas)
 */

const SCHED_TZ = typeof TIMEZONE !== 'undefined' ? TIMEZONE : 'America/Bogota';
const SCHED_CUTOFF = typeof HARD_CUTOFF_HOUR !== 'undefined' ? HARD_CUTOFF_HOUR : 19;

/**
 * Proactive Open-Slot Suggester (Max 3 ranked choices, strictly excluding completed/scheduled work).
 */
function suggestForOpenSlot(dateStr, startTime24, durationMin, slotTitle) {
  let targetDate = dateStr;
  let targetStart = startTime24;
  let targetDur = Number(durationMin) || 60;
  let label = slotTitle || 'Open Window';

  if (!targetDate || !targetStart) {
    const upcoming = getAvailableCalendarSlots_(4);
    if (upcoming.length === 0) {
      throw new Error('No open focus windows found in the next 4 days.');
    }
    const first = upcoming[0];
    targetDate = first.date;
    targetStart = first.start_time;
    targetDur = Math.min(first.available_min, 60);
    label = `${first.day_name} ${first.energy_type}`;
  }

  const slotStartDate = new Date(`${targetDate}T${targetStart}:00-05:00`);
  const dayName = Utilities.formatDate(slotStartDate, SCHED_TZ, 'EEEE');
  const dow = Number(Utilities.formatDate(slotStartDate, SCHED_TZ, 'u')) % 7;
  const isHeavyTeachingDay = dow === 2 || dow === 4;
  const cappedDur = Math.min(Math.max(targetDur, 15), 60);

  const cards = typeof getProjectCards === 'function' ? getProjectCards() : [];
  const projectsSummary = typeof formatProjectsContextForPrompt_ === 'function'
    ? formatProjectsContextForPrompt_(cards)
    : 'No project cards found.';

  const history = typeof getCompletionAndScheduleContext_ === 'function'
    ? getCompletionAndScheduleContext_()
    : { doneTasks: [], scheduledTasks: [], calendarEvents: [] };

  const prompt = `Joel wants 3 ranked, high-leverage suggestions for his next open calendar window.

OPEN WINDOW DETAILS:
- Window: ${label}
- Day & Date: ${dayName} (${targetDate})
- Start Time: ${targetStart} | Available Duration: ${targetDur} minutes (max block ${cappedDur} min)
- Energy Constraint: ${isHeavyTeachingDay ? 'HEAVY TEACHING DAY (Tuesday/Thursday) — suggest ONLY low-cognitive-load tasks: 15–20 min Spanish homework sprint, video script outlining, or rubric batching. NO heavy coding.' : 'Standard/Deep energy day — prefer 60-minute chunks for Tier 1A/1B deep work if >= 60m is available, plus a 15–20 min Spanish homework micro-sprint option.'}

LIVE PROJECT CARDS & OPEN CHECKLIST ITEMS:
${projectsSummary}

STRICT EXCLUSION LIST — ALREADY DONE OR ALREADY SCHEDULED (NEVER SUGGEST THESE):
- Completed Tasks: ${history.doneTasks.join('; ') || 'None'}
- Already Scheduled: ${history.scheduledTasks.join('; ') || 'None'}
- Recent/Upcoming Calendar Events: ${history.calendarEvents.join('; ') || 'None'}

INSTRUCTIONS (SATISFICING CAP = MAX 3 RANKED OPTIONS):
1. Return exactly 3 ranked options drawn from OPEN (unchecked) checklist items across Joel's project cards.
2. NEVER pitch anything that overlaps with Completed Tasks, checked-off items, Already Scheduled blocks, or Calendar Events.
3. Prefer 60-minute chunks when >= 60m is free; cap Spanish homework strictly at 15 or 20 minutes.`;

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
            project_id: {
              type: 'STRING',
              enum: ['PRJ_WAL', 'PRJ_COJOWA', 'PRJ_AIPOLICY', 'PRJ_SPANISH', 'PRJ_SANDBOX']
            },
            project_tag: { type: 'STRING' },
            tier: {
              type: 'STRING',
              enum: ['Tier 1A - Teaching', 'Tier 1B - Consulting/PD', 'Tier 2 - AI Policy', 'Personal / Spanish & Ops', 'Tier 4 - Sandbox/Icebox']
            },
            title: { type: 'STRING' },
            duration_min: { type: 'INTEGER' },
            why_now: { type: 'STRING' },
            two_min_start: { type: 'STRING' },
            stopping_rule: { type: 'STRING' }
          },
          required: ['rank', 'project_id', 'project_tag', 'tier', 'title', 'duration_min', 'why_now', 'two_min_start', 'stopping_rule']
        }
      }
    },
    required: ['window_advice', 'options']
  };

  const res = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'open_slot_suggester'
  );

  return {
    dateStr: targetDate,
    dayName: dayName,
    startTime24: targetStart,
    windowLabel: `${dayName} ${targetDate} @ ${Utilities.formatDate(slotStartDate, SCHED_TZ, 'h:mm a')} (${label})`,
    advice: res.window_advice,
    options: (res.options || []).slice(0, 3).map(opt => ({
      ...opt,
      duration_min: Math.min(Math.max(Number(opt.duration_min) || 60, 15), targetDur)
    }))
  };
}

/**
 * Writes approved meetings, anchors & focus blocks to Google Calendar and links them to Project Cards.
 */
function commitScheduledBlocks(blocks) {
  if (!blocks || blocks.length === 0) {
    throw new Error('No blocks selected to write to Google Calendar.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('Task_Queue');
  const queueData = queueSheet ? queueSheet.getDataRange().getValues() : [];
  const calendars = getIncludedCalendars_();
  const freeRegex = getFreePeriodRegex_();

  const colorMap = {
    'Tier 1A - Teaching': '9',           // Blueberry
    'Tier 1B - Consulting/PD': '10',     // Basil Green
    'Tier 2 - AI Policy': '3',           // Grape
    'Personal / Spanish & Ops': '6',     // Tangerine
    'Tier 4 - Sandbox/Icebox': '5',      // Banana
    'Protected Anchor / Recovery': '2'   // Sage Green
  };

  let createdCount = 0;
  const skipped = [];

  blocks.forEach(b => {
    if (b.writeToCalendar === false) return;

    const startDate = new Date(`${b.dateStr}T${b.startTimeStr}:00-05:00`);
    const endDate = new Date(startDate.getTime() + Number(b.durationMin) * 60000);

    const endHour = Number(Utilities.formatDate(endDate, SCHED_TZ, 'H'));
    const endMin = Number(Utilities.formatDate(endDate, SCHED_TZ, 'm'));
    if (endHour > SCHED_CUTOFF || (endHour === SCHED_CUTOFF && endMin > 0)) {
      skipped.push(`${b.title} (crosses 7:00 PM cutoff)`);
      return;
    }

    let conflictTitle = null;
    for (const cal of calendars) {
      try {
        const res = Calendar.Events.list(cal.id, {
          timeMin: startDate.toISOString(),
          timeMax: endDate.toISOString(),
          singleEvents: true
        });
        const hit = (res.items || []).find(ev => {
          if (!ev.start || !ev.start.dateTime) return false;
          if (ev.transparency === 'transparent') return false;
          if (freeRegex.test(ev.summary || '')) return false;
          return true;
        });
        if (hit) {
          conflictTitle = hit.summary || 'existing event';
          break;
        }
      } catch (e) {}
    }

    if (conflictTitle) {
      skipped.push(`${b.title} (conflicts with "${conflictTitle}")`);
      return;
    }

    const isFixed = Boolean(b.isFixedEvent || b.isProtectedAnchor);
    const summaryTitle = isFixed ? b.title : `[Focus] ${b.title}`;
    const descLines = [
      b.projectTag ? `Project: ${b.projectTag}` : '',
      `Tier: ${b.tier}`,
      `2-Min Entry Action: ${b.twoMinStart}`,
      `85% Stopping Rule: ${b.stoppingRule}`
    ].filter(Boolean).join('\n');

    const eventResource = {
      summary: summaryTitle,
      description: descLines,
      start: { dateTime: startDate.toISOString(), timeZone: SCHED_TZ },
      end: { dateTime: endDate.toISOString(), timeZone: SCHED_TZ },
      colorId: colorMap[b.tier] || '9',
      reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 10 }] }
    };

    const createdEvent = Calendar.Events.insert(eventResource, 'primary');
    const formattedSlot = Utilities.formatDate(startDate, SCHED_TZ, 'EEE MMM d, h:mm a');

    if (!isFixed && queueSheet) {
      let updatedExisting = false;
      if (b.existingTaskId) {
        for (let i = 1; i < queueData.length; i++) {
          if (String(queueData[i][0]) === String(b.existingTaskId)) {
            queueSheet.getRange(i + 1, 5).setValue('Scheduled');
            queueSheet.getRange(i + 1, 9).setValue(b.projectId || '');
            queueSheet.getRange(i + 1, 10).setValue(formattedSlot);
            queueSheet.getRange(i + 1, 11).setValue(createdEvent.id);
            updatedExisting = true;
            break;
          }
        }
      }

      if (!updatedExisting) {
        const newTaskId = 'TSK_' + Utilities.getUuid().slice(0, 6).toUpperCase();
        queueSheet.appendRow([
          newTaskId,
          Utilities.formatDate(new Date(), SCHED_TZ, 'yyyy-MM-dd HH:mm'),
          b.title,
          b.tier,
          'Scheduled',
          'None',
          b.twoMinStart,
          `Stop when: ${b.stoppingRule}`,
          b.projectId || '',
          formattedSlot,
          createdEvent.id
        ]);
      }
    }

    createdCount++;
  });

  let summary = `Wrote ${createdCount} block(s) to Google Calendar.`;
  if (skipped.length > 0) {
    summary += ` Skipped ${skipped.length}: ${skipped.join('; ')}.`;
  }
  return summary;
}

function getIncludedCalendars_() {
  const listRes = Calendar.CalendarList.list();
  const allCals = listRes.items || [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paramSheet = ss.getSheetByName('Parameters');
  let customNames = null;

  if (paramSheet && paramSheet.getLastRow() > 1) {
    const rows = paramSheet.getRange(2, 1, paramSheet.getLastRow() - 1, 2).getValues();
    const found = rows.find(r => String(r[0]).trim() === 'Included_Calendars' && String(r[1]).trim());
    if (found) {
      customNames = String(found[1])
        .split(',')
        .map(s => s.trim().toLowerCase())
        .filter(Boolean);
    }
  }

  return allCals.filter(c => {
    if (c.primary) return true;
    const id = String(c.id || '');
    if (id.includes('#holiday@') || id.includes('#sports@')) return false;
    if (customNames) {
      return customNames.includes(String(c.summary || '').trim().toLowerCase());
    }
    return c.selected !== false;
  });
}

function getFreePeriodRegex_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paramSheet = ss.getSheetByName('Parameters');
  if (paramSheet && paramSheet.getLastRow() > 1) {
    const rows = paramSheet.getRange(2, 1, paramSheet.getLastRow() - 1, 2).getValues();
    const found = rows.find(r => String(r[0]).trim() === 'Free_Period_Keywords' && String(r[1]).trim());
    if (found) {
      const escaped = String(found[1])
        .split(',')
        .map(k => k.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
        .filter(Boolean);
      if (escaped.length > 0) {
        return new RegExp(`\\b(${escaped.join('|')})\\b`, 'i');
      }
    }
  }
  return /\bplan\b/i;
}

function subtractIntervalsFromSlots_(slots, blockedIntervals) {
  const result = [];
  let counter = 1;

  slots.forEach(slot => {
    let segments = [{ start: slot.start_ms, end: slot.end_ms }];

    blockedIntervals.forEach(b => {
      const next = [];
      segments.forEach(seg => {
        if (b.end <= seg.start || b.start >= seg.end) {
          next.push(seg);
        } else {
          if (b.start - seg.start >= 15 * 60000) {
            next.push({ start: seg.start, end: b.start });
          }
          if (seg.end - b.end >= 15 * 60000) {
            next.push({ start: b.end, end: seg.end });
          }
        }
      });
      segments = next;
    });

    segments.forEach(seg => {
      const availMin = Math.floor((seg.end - seg.start) / 60000);
      if (availMin >= 15) {
        result.push({
          slot_id: `CARVED_${counter++}`,
          parent_slot_id: slot.slot_id,
          is_school_period: Boolean(slot.is_school_period),
          date: slot.date,
          day_name: slot.day_name,
          start_time: Utilities.formatDate(new Date(seg.start), SCHED_TZ, 'HH:mm'),
          end_time: Utilities.formatDate(new Date(seg.end), SCHED_TZ, 'HH:mm'),
          start_ms: seg.start,
          end_ms: seg.end,
          available_min: availMin,
          energy_type: slot.energy_type
        });
      }
    });
  });

  return result;
}

function getAvailableCalendarSlots_(daysAhead) {
  const now = new Date();
  const todayMidStr = Utilities.formatDate(now, SCHED_TZ, 'yyyy-MM-dd');
  const anchorNoon = new Date(`${todayMidStr}T12:00:00-05:00`);

  const rangeStart = new Date(now.getTime());
  const rangeEnd = new Date(anchorNoon.getTime() + daysAhead * 86400000);

  const calendars = getIncludedCalendars_();
  const freeRegex = getFreePeriodRegex_();
  const busyIntervals = [];
  const planEvents = [];
  const minBufferNowMs = now.getTime() + 15 * 60000;

  calendars.forEach(cal => {
    try {
      const calRes = Calendar.Events.list(cal.id, {
        timeMin: rangeStart.toISOString(),
        timeMax: rangeEnd.toISOString(),
        singleEvents: true,
        orderBy: 'startTime'
      });

      (calRes.items || []).forEach(ev => {
        if (!ev.start || !ev.start.dateTime || !ev.end || !ev.end.dateTime) return;
        if (ev.eventType === 'workingLocation') return;

        const title = String(ev.summary || '');
        const evStart = new Date(ev.start.dateTime).getTime();
        const evEnd = new Date(ev.end.dateTime).getTime();
        const evDateStr = Utilities.formatDate(new Date(evStart), SCHED_TZ, 'yyyy-MM-dd');

        if (freeRegex.test(title)) {
          planEvents.push({
            date: evDateStr,
            start: evStart,
            end: evEnd,
            title: title
          });
          return;
        }

        if (ev.transparency !== 'transparent') {
          busyIntervals.push({ start: evStart, end: evEnd });
        }
      });
    } catch (e) {}
  });

  const rawSlots = [];
  let slotCounter = 1;

  for (let d = 0; d < daysAhead; d++) {
    const dayDate = new Date(anchorNoon.getTime() + d * 86400000);
    const dateStr = Utilities.formatDate(dayDate, SCHED_TZ, 'yyyy-MM-dd');
    const dayName = Utilities.formatDate(dayDate, SCHED_TZ, 'EEEE');
    const dow = Number(Utilities.formatDate(dayDate, SCHED_TZ, 'u')) % 7;
    const isWeekend = dow === 0 || dow === 6;

    const defaultStartStr = isWeekend ? '09:00' : '08:00';
    let winStart = new Date(`${dateStr}T${defaultStartStr}:00-05:00`).getTime();
    const winEnd = new Date(`${dateStr}T${String(SCHED_CUTOFF).padStart(2, '0')}:00:00-05:00`).getTime();

    planEvents
      .filter(p => p.date === dateStr)
      .forEach(p => {
        if (p.start < winStart) winStart = p.start;
      });

    if (winStart < minBufferNowMs) winStart = minBufferNowMs;
    if (winEnd - winStart < 15 * 60000) continue;

    const isLightDay = dow === 2 || dow === 4;
    const energyLabel = isWeekend
      ? 'Weekend Open Canvas (1h chunks or 15-20m micro-sprints)'
      : isLightDay
        ? 'Tuesday/Thursday Open Window (Light Grading / Admin / 15-20m Micro-Sprint)'
        : 'Weekday Open Window (Deep Work / Video Prep / Prototype)';

    rawSlots.push({
      slot_id: `SLOT_${slotCounter++}`,
      is_school_period: false,
      date: dateStr,
      day_name: dayName,
      start_time: Utilities.formatDate(new Date(winStart), SCHED_TZ, 'HH:mm'),
      end_time: Utilities.formatDate(new Date(winEnd), SCHED_TZ, 'HH:mm'),
      start_ms: winStart,
      end_ms: winEnd,
      available_min: Math.floor((winEnd - winStart) / 60000),
      energy_type: energyLabel
    });
  }

  const carved = subtractIntervalsFromSlots_(rawSlots, busyIntervals);
  return carved.map(seg => {
    const overlappingPlan = planEvents.find(p => p.start < seg.end_ms && p.end > seg.start_ms);
    if (overlappingPlan) {
      return {
        ...seg,
        is_school_period: true,
        energy_type: `School Plan Period (${overlappingPlan.title})`
      };
    }
    return seg;
  });
}