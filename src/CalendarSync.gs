/**
 * CalendarSync.gs — Stage 1 Safe Mode
 *
 * Detects events on Personal + Kids calendars, adds 15-min commute buffers,
 * copies to Automation calendar, marks source events as processed, and logs.
 *
 * Runs once daily at ~5:30 PM, every day including weekends.
 * Supports 9/80 schedule — detects OFF Fridays via "Off Friday" all-day
 * event on the Kids calendar. ON Fridays use 7 AM–3 PM hours.
 * Stage 1: NO work emails sent. Logs + Automation calendar only.
 *
 * GitHub: https://github.com/jcann/calendar-sync
 * Version: 1.5.0 — 9/80 schedule support, kids calendar enabled
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

const CONFIG = {
  // Calendar IDs
  PERSONAL_CALENDAR_ID:   "primary",
  KIDS_CALENDAR_ID:       "8763718fbb7b77eabf8128f92dd566f8ba12b4059f1e8668d7f2ef95d26f5ca9@group.calendar.google.com",
  AUTOMATION_CALENDAR_ID: "3e6bf5b370e1609f0bee08614941ff268aae0ebe087552a48eafb037f539c587@group.calendar.google.com",

  // Your personal Gmail address for log summary emails
  PERSONAL_EMAIL: "cannj65@gmail.com",

  // How far ahead to look for new events (days)
  LOOKAHEAD_DAYS: 14,

  // Commute buffer in minutes
  COMMUTE_BUFFER_MINUTES: 15,

  // Standard business hours (Mon–Thu and ON Fridays)
  BUSINESS_START_HOUR: 7,
  BUSINESS_END_HOUR:   17, // 5 PM Mon–Thu

  // ON Friday hours (9/80 schedule — shorter day)
  FRIDAY_END_HOUR: 15, // 3 PM on ON Fridays

  // Weekdays (Sun=0, Mon=1 … Sat=6)
  WORK_DAYS: [1, 2, 3, 4, 5],

  // OFF Friday detection — title of the all-day event on the Kids calendar
  OFF_FRIDAY_TITLE: "Off Friday",

  // Tag appended to processed source events — short and unobtrusive
  PROCESSED_TAG: "#cannudigit-cal-sync",

  // Stage control — flip these as you progress through stages
  STAGE: {
    COPY_TO_AUTOMATION: true,  // Stage 1: ON
    LOG_EMAIL:          true,  // Daily summary to personal Gmail — flip false to stop
    SEND_EMAIL:         true,  // Stage 4+: work Outlook email — OFF for now
    INCLUDE_KIDS:       true,  // Stage 3: ON
  },

  // Work email config (unused in Stage 1 — filled in for Stage 4)
  EMAIL: {
    TO:      "cannj65@gmail.com", // Stage 4: swap to work Outlook address
    SUBJECT: "Busy",
    BODY:    "Blocked - personal appointment",
  },
};

// ─── MAIN ENTRY POINT ────────────────────────────────────────────────────────

/**
 * runSync() — called by the daily 5:30 PM trigger, every day.
 * Safe to run manually at any time for testing.
 */
function runSync() {
  const now      = new Date();
  const logLines = [];

  const record = (msg) => {
    log(msg);
    logLines.push(msg);
  };

  record(`=== Sync run at ${formatTime(now)} ===`);

  const calendarIds = [CONFIG.PERSONAL_CALENDAR_ID];
  if (CONFIG.STAGE.INCLUDE_KIDS) {
    calendarIds.push(CONFIG.KIDS_CALENDAR_ID);
  }

  let totalProcessed = 0;
  let totalSkipped   = 0;
  let totalErrors    = 0;

  calendarIds.forEach(calId => {
    const calName = calId === CONFIG.PERSONAL_CALENDAR_ID ? "Personal" : "Kids";
    const events  = getUpcomingEvents(calId);

    record(`  [${calName}] Found ${events.length} event(s) in ${CONFIG.LOOKAHEAD_DAYS}-day window`);

    events.forEach(event => {
      try {
        // Skip the OFF Friday marker event itself
        if (isOffFridayMarker(event)) {
          return;
        }

        if (isAlreadyProcessed(event)) {
          record(`    SKIP: "${event.getTitle()}" — already processed`);
          totalSkipped++;
          return;
        }

        const overlapResult = overlapsWorkHours(event);

        if (!overlapResult.overlaps) {
          record(`    SKIP: "${event.getTitle()}" — ${overlapResult.reason}`);
          totalSkipped++;
          return;
        }

        const success = processEvent(event, calName, record);
        if (success) {
          totalProcessed++;
        } else {
          totalErrors++;
        }
      } catch (err) {
        record(`    ERROR processing "${event.getTitle()}": ${err.message}`);
        totalErrors++;
      }
    });
  });

  const summary = `=== Done. ${totalProcessed} processed, ${totalSkipped} skipped, ${totalErrors} error(s) ===`;
  record(summary);

  // Send summary email to personal Gmail if enabled
  if (CONFIG.STAGE.LOG_EMAIL) {
    sendLogEmail(now, totalProcessed, totalSkipped, totalErrors, logLines);
  }
}

// ─── EVENT PROCESSING ────────────────────────────────────────────────────────

/**
 * processEvent() — copies to Automation calendar, then marks as processed.
 * Returns true only if the copy succeeded. Marking happens AFTER copy.
 */
function processEvent(event, sourceCalName, record) {
  const title       = event.getTitle() || "(no title)";
  const start       = event.getStartTime();
  const end         = event.getEndTime();
  const bufferStart = addMinutes(start, -CONFIG.COMMUTE_BUFFER_MINUTES);
  const bufferEnd   = addMinutes(end,   +CONFIG.COMMUTE_BUFFER_MINUTES);

  record(`    PROCESSING: "${title}" [${formatTime(start)} – ${formatTime(end)}]`);
  record(`      Buffer window: ${formatTime(bufferStart)} – ${formatTime(bufferEnd)}`);

  let success = true;

  if (CONFIG.STAGE.COPY_TO_AUTOMATION) {
    success = copyToAutomationCalendar(title, bufferStart, bufferEnd, sourceCalName, record);
  }

  if (!success) {
    record(`      NOT marked as processed — copy failed. Will retry next run.`);
    return false;
  }

  if (CONFIG.STAGE.SEND_EMAIL) {
    sendOutlookBlock(bufferStart, bufferEnd, record);
  }

  markAsProcessed(event);
  record(`      Marked as processed.`);
  return true;
}

// ─── CALENDAR HELPERS ────────────────────────────────────────────────────────

function getUpcomingEvents(calendarId) {
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) {
    log(`  WARNING: Calendar not found: ${calendarId}`);
    return [];
  }
  const start = new Date();
  const end   = addDays(start, CONFIG.LOOKAHEAD_DAYS);
  return cal.getEvents(start, end);
}

/**
 * copyToAutomationCalendar() — returns true on success, false on failure.
 */
function copyToAutomationCalendar(title, start, end, sourceCalName, record) {
  const automationCal = CalendarApp.getCalendarById(CONFIG.AUTOMATION_CALENDAR_ID);
  if (!automationCal) {
    record(`      ERROR: Automation calendar not found. Check AUTOMATION_CALENDAR_ID.`);
    return false;
  }

  const safeTitle       = "Busy";
  const safeDescription = [
    "Auto-blocked (incl. commute buffer)",
    `Source: ${sourceCalName} calendar`,
    `Original: ${title}`,
    CONFIG.PROCESSED_TAG,
  ].join("\n");

  // Duplicate check before creating
  const existing  = automationCal.getEvents(start, end);
  const duplicate = existing.some(e =>
    e.getTitle() === safeTitle &&
    Math.abs(e.getStartTime() - start) < 60000
  );

  if (duplicate) {
    record(`      SKIP copy — duplicate already exists on Automation calendar`);
    return true; // not a failure — already there
  }

  automationCal.createEvent(safeTitle, start, end, {
    description: safeDescription,
  });

  record(`      Copied to Automation calendar: "${safeTitle}" [${formatTime(start)} – ${formatTime(end)}]`);
  return true;
}

function markAsProcessed(event) {
  const desc = event.getDescription() || "";
  if (!desc.includes(CONFIG.PROCESSED_TAG)) {
    event.setDescription(desc + "\n" + CONFIG.PROCESSED_TAG);
  }
}

function isAlreadyProcessed(event) {
  const desc = event.getDescription() || "";
  return desc.includes(CONFIG.PROCESSED_TAG);
}

/**
 * isOffFridayMarker() — returns true if this event IS the "Off Friday"
 * all-day marker itself, so we don't try to process or skip-log it.
 */
function isOffFridayMarker(event) {
  return event.isAllDayEvent() &&
         event.getTitle().trim() === CONFIG.OFF_FRIDAY_TITLE;
}

/**
 * isOffFriday() — checks the Kids calendar for an all-day "Off Friday"
 * event on the given date. Returns true if found.
 */
function isOffFriday(date) {
  const cal = CalendarApp.getCalendarById(CONFIG.KIDS_CALENDAR_ID);
  if (!cal) return false;

  // Check the full day of the given date
  const dayStart = new Date(date);
  dayStart.setHours(0, 0, 0, 0);
  const dayEnd = new Date(date);
  dayEnd.setHours(23, 59, 59, 999);

  const events = cal.getEvents(dayStart, dayEnd);
  return events.some(e =>
    e.isAllDayEvent() &&
    e.getTitle().trim() === CONFIG.OFF_FRIDAY_TITLE
  );
}

// ─── TIME HELPERS ─────────────────────────────────────────────────────────────

/**
 * overlapsWorkHours() — returns { overlaps: bool, reason: string }
 *
 * Rules:
 *   - Mon–Thu: 7 AM–5 PM
 *   - OFF Friday: skip entirely (treated like weekend)
 *   - ON Friday: 7 AM–3 PM
 *   - Sat–Sun: skip
 */
function overlapsWorkHours(event) {
  const start   = event.getStartTime();
  const end     = event.getEndTime();
  const dayOfWk = start.getDay(); // 0=Sun, 5=Fri, 6=Sat

  // Weekend
  if (!CONFIG.WORK_DAYS.includes(dayOfWk)) {
    return { overlaps: false, reason: "weekend event" };
  }

  // Friday handling
  if (dayOfWk === 5) {
    if (isOffFriday(start)) {
      return { overlaps: false, reason: "OFF Friday" };
    }
    // ON Friday — shorter day
    const startHour = start.getHours() + start.getMinutes() / 60;
    const endHour   = end.getHours()   + end.getMinutes()   / 60;
    if (startHour >= CONFIG.FRIDAY_END_HOUR || endHour <= CONFIG.BUSINESS_START_HOUR) {
      return { overlaps: false, reason: "outside ON Friday hours (7 AM–3 PM)" };
    }
    return { overlaps: true, reason: "" };
  }

  // Mon–Thu standard hours
  const startHour = start.getHours() + start.getMinutes() / 60;
  const endHour   = end.getHours()   + end.getMinutes()   / 60;
  if (startHour >= CONFIG.BUSINESS_END_HOUR || endHour <= CONFIG.BUSINESS_START_HOUR) {
    return { overlaps: false, reason: "outside business hours" };
  }

  return { overlaps: true, reason: "" };
}

function addMinutes(date, minutes) {
  return new Date(date.getTime() + minutes * 60 * 1000);
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "EEE MM/dd HH:mm");
}

// ─── LOGGING ─────────────────────────────────────────────────────────────────

function log(message) {
  console.log(message);
}

// ─── SUMMARY EMAIL ────────────────────────────────────────────────────────────

/**
 * sendLogEmail() — sends the full run log to your personal Gmail.
 * One email per day. Subject gives you the key stats at a glance.
 * Flip CONFIG.STAGE.LOG_EMAIL to false to stop receiving these.
 */
function sendLogEmail(runTime, processed, skipped, errors, logLines) {
  const dateStr = Utilities.formatDate(runTime, Session.getScriptTimeZone(), "EEE MM/dd");
  const status  = errors > 0 ? "⚠️ ERROR" : processed > 0 ? "✅ OK" : "— idle";
  const subject = `CalendarSync — ${dateStr} — ${processed} processed, ${skipped} skipped ${status}`;
  const body    = logLines.join("\n");

  GmailApp.sendEmail(CONFIG.PERSONAL_EMAIL, subject, body);
  log(`Summary email sent to ${CONFIG.PERSONAL_EMAIL}`);
}

// ─── WORK EMAIL (STAGE 4+) ───────────────────────────────────────────────────

/**
 * Sends a plain-text email your work Outlook can receive.
 * No sensitive details — generic subject + time window only.
 * NOT ACTIVE in Stage 1.
 */
function sendOutlookBlock(start, end, record) {
  const subject = CONFIG.EMAIL.SUBJECT;
  const body    = [
    CONFIG.EMAIL.BODY,
    "",
    `Start: ${formatTime(start)}`,
    `End:   ${formatTime(end)}`,
    "",
    "(Auto-generated. Do not reply.)",
  ].join("\n");

  GmailApp.sendEmail(CONFIG.EMAIL.TO, subject, body);
  record(`      Email sent to ${CONFIG.EMAIL.TO}: "${subject}"`);
}

// ─── TRIGGER SETUP ───────────────────────────────────────────────────────────

/**
 * createTrigger() — run ONCE manually to install the daily trigger.
 * Fires every day at ~5:30 PM including weekends.
 *
 * Confirm it appears in Apps Script → ⏰ Triggers after running.
 * Do NOT call this inside runSync() — it would create duplicate triggers.
 */
function createTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runSync")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("runSync")
    .timeBased()
    .atHour(17)
    .nearMinute(30)
    .everyDays(1)
    .create();

  console.log("Trigger created: runSync daily at ~5:30 PM.");
}

/**
 * deleteTriggers() — removes all project triggers. Useful during testing.
 */
function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  console.log("All triggers deleted.");
}

// ─── STAGE ADVANCEMENT CHECKLIST ─────────────────────────────────────────────
//
// Stage 1 (NOW):  COPY_TO_AUTOMATION=true, LOG_EMAIL=true,
//                 SEND_EMAIL=false, INCLUDE_KIDS=true
// Stage 2:        Observe a few days — confirm no duplicates — done
// Stage 3:        ✅ INCLUDE_KIDS=true — done
// Stage 4:        SEND_EMAIL=true, EMAIL.TO=personal Gmail — confirm email format
// Stage 5:        EMAIL.TO=work Outlook address — production
// Any time:       LOG_EMAIL=false to stop daily summary emails