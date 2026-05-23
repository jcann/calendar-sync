/**
 * CalendarSync.gs
 *
 * Detects events on Personal + Kids calendars, adds 15-min commute buffers,
 * copies to Automation calendar, marks source events as processed, and logs.
 *
 * Runs once daily at ~5:30 PM, every day including weekends.
 * Supports 9/80 schedule — detects OFF Fridays via "Off Friday" all-day
 * event on the Kids calendar. ON Fridays use 7 AM–3 PM hours.
 *
 * GitHub: https://github.com/jcann/calendar-sync
 * Version: 1.7.0 — ICS blob on per-event work email
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

const CONFIG = {
    // Calendar IDs
    PERSONAL_CALENDAR_ID: "primary",
    KIDS_CALENDAR_ID: 
        "8763718fbb7b77eabf8128f92dd566f8ba12b4059f1e8668d7f2ef95d26f5ca9@group.calendar.google.com",
    AUTOMATION_CALENDAR_ID:
        "3e6bf5b370e1609f0bee08614941ff268aae0ebe087552a48eafb037f539c587@group.calendar.google.com",

    // Your personal Gmail address for log summary emails
    PERSONAL_EMAIL: "cannj65@gmail.com",

    // How far ahead to look for new events (days)
    LOOKAHEAD_DAYS: 14,

    // Commute buffer in minutes
    COMMUTE_BUFFER_MINUTES: 15,

    // Standard business hours (Mon–Thu and ON Fridays)
    BUSINESS_START_HOUR: 7,
    BUSINESS_END_HOUR: 17, // 5 PM Mon–Thu

    // ON Friday hours (9/80 schedule — shorter day)
    FRIDAY_END_HOUR: 15, // 3 PM on ON Fridays

    // Weekdays (Sun=0, Mon=1 … Sat=6)
    WORK_DAYS: [1, 2, 3, 4, 5],

    // OFF Friday detection — title of the all-day event on the Kids calendar
    OFF_FRIDAY_TITLE: "Off Friday",

    // Tag appended to processed source events — short and unobtrusive
    PROCESSED_TAG: "#cannudigit-cal-sync",

    // Control — flip these for dev/prod processing
    CONTROL: {
        COPY_TO_AUTOMATION: true,
        LOG_EMAIL: true,
        SEND_EMAIL: true,
        INCLUDE_KIDS: true,
    },

    // Work email config
    EMAIL: {
        TO: "jonathan.cann@gd-ms.com",
    },
};

// ─── MAIN ENTRY POINT ────────────────────────────────────────────────────────

/**
 * runSync() — called by the daily 5:30 PM trigger, every day.
 * Safe to run manually at any time for testing.
 */
function runSync() {
    const now = new Date();
    const logLines = [];

    const record = (msg) => {
        log(msg);
        logLines.push(msg);
    };

    record(`=== Sync run at ${formatTime(now)} ===`);

    const calendarIds = [CONFIG.PERSONAL_CALENDAR_ID];
    if (CONFIG.CONTROL.INCLUDE_KIDS) {
        calendarIds.push(CONFIG.KIDS_CALENDAR_ID);
    }

    let totalProcessed = 0;
    let totalSkipped = 0;
    let totalErrors = 0;

    calendarIds.forEach((calId) => {
        const calName = calId === CONFIG.PERSONAL_CALENDAR_ID ? "Personal" : "Kids";
        const events = getUpcomingEvents(calId);

        record(`  [${calName}] Found ${events.length} event(s) in ${CONFIG.LOOKAHEAD_DAYS}-day window`);

        events.forEach((event) => {
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
    if (CONFIG.CONTROL.LOG_EMAIL) {
        sendLogEmail(now, totalProcessed, totalSkipped, totalErrors, logLines);
    }
}

// ─── EVENT PROCESSING ────────────────────────────────────────────────────────

/**
 * processEvent() — copies to Automation calendar, sends ICS email,
 * then marks as processed. Marking only happens after success.
 */
function processEvent(event, sourceCalName, record) {
    const title = event.getTitle() || "(no title)";
    const start = event.getStartTime();
    const end = event.getEndTime();
    const bufferStart = addMinutes(start, -CONFIG.COMMUTE_BUFFER_MINUTES);
    const bufferEnd = addMinutes(end, +CONFIG.COMMUTE_BUFFER_MINUTES);

    record(`    PROCESSING: "${title}" [${formatTime(start)} – ${formatTime(end)}]`);
    record(`      Buffer window: ${formatTime(bufferStart)} – ${formatTime(bufferEnd)}`);

    let success = true;

    if (CONFIG.CONTROL.COPY_TO_AUTOMATION) {
        success = copyToAutomationCalendar(title, bufferStart, bufferEnd, sourceCalName, record);
    }

    if (!success) {
        record(`      NOT marked as processed — copy failed. Will retry next run.`);
        return false;
    }

    if (CONFIG.CONTROL.SEND_EMAIL) {
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
    const end = addDays(start, CONFIG.LOOKAHEAD_DAYS);
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

    const safeTitle = "Busy";
    const safeDescription = [
        "Auto-blocked (incl. commute buffer)",
        `Source: ${sourceCalName} calendar`,
        `Original: ${title}`,
        CONFIG.PROCESSED_TAG,
    ].join("\n");

    // Duplicate check before creating
    const existing = automationCal.getEvents(start, end);
    const duplicate = existing.some((e) => e.getTitle() === safeTitle && Math.abs(e.getStartTime() - start) < 60000);

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
    return event.isAllDayEvent() && event.getTitle().trim() === CONFIG.OFF_FRIDAY_TITLE;
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
    return events.some((e) => e.isAllDayEvent() && e.getTitle().trim() === CONFIG.OFF_FRIDAY_TITLE);
}

// ─── ICS BUILDER ─────────────────────────────────────────────────────────────

/**
 * buildICS() — generates a valid ICS file string for a single busy block.
 *
 * - TRANSP:OPAQUE marks the time as busy in Outlook
 * - CLASS:PRIVATE keeps the event private
 * - No ORGANIZER, ATTENDEE, or METHOD fields — plain import, not an invite
 * - UID is derived from start time + title hash to stay consistent across runs
 */
function buildICS(start, end) {
    const uid = generateUID(start);
    const stampNow = formatICSDate(new Date());
    const dtStart = formatICSDate(start);
    const dtEnd = formatICSDate(end);

    return [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//CalendarSync//EN",
        "CALSCALE:GREGORIAN",
        "BEGIN:VEVENT",
        `UID:${uid}`,
        `DTSTAMP:${stampNow}`,
        `DTSTART:${dtStart}`,
        `DTEND:${dtEnd}`,
        "SUMMARY:Busy",
        "CLASS:PRIVATE",
        "TRANSP:OPAQUE",
        "END:VEVENT",
        "END:VCALENDAR",
    ].join("\r\n");
}

/**
 * formatICSDate() — formats a Date as ICS UTC timestamp: 20250327T144500Z
 */
function formatICSDate(date) {
    return Utilities.formatDate(date, "UTC", "yyyyMMdd'T'HHmmss'Z'");
}

/**
 * generateUID() — creates a unique stable ID for this event.
 * Same event on re-runs produces the same UID, so Outlook
 * won't create duplicates if the ICS is imported more than once.
 */
function generateUID(start) {
    const base = start.getTime().toString();
    return `${base}-cannudigit-cal-sync@google-apps-script`;
}

// ─── WORK EMAIL ──────────────────────────────────────────────────────────────

/**
 * sendOutlookBlock() — sends one email per event with an ICS blob.
 * Subject includes the date and time window for at-a-glance reading.
 * Body is intentionally minimal — no personal details.
 */
function sendOutlookBlock(start, end, record) {
    const dateStr = Utilities.formatDate(start, Session.getScriptTimeZone(), "EEE MM/dd");
    const subject = `Busy — ${dateStr} ${formatTimeShort(start)} – ${formatTimeShort(end)}`;

    const body = [
        "Calendar block — personal appointment.",
        "",
        `Date:  ${dateStr}`,
        `Start: ${formatTimeShort(start)}`,
        `End:   ${formatTimeShort(end)}`,
        "",
        "Open the attached .ics file to add this block to your Outlook calendar.",
        "(Auto-generated. Do not reply.)",
    ].join("\n");

    const icsContent = buildICS(start, end);
    const icsFilename = `busy-${Utilities.formatDate(start, Session.getScriptTimeZone(), "MM-dd-HHmm")}.ics`;

    const blob = Utilities.newBlob(icsContent).setName(icsFilename).setContentType("application/octet-stream");

    MailApp.sendEmail({
        to: CONFIG.EMAIL.TO,
        subject: subject,
        body: body,
        attachments: [blob],
        name: "CalendarSync",
    });

    record(`      Email + ICS sent to ${CONFIG.EMAIL.TO}: "${subject}"`);
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
    const start = event.getStartTime();
    const end = event.getEndTime();
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
        const endHour = end.getHours() + end.getMinutes() / 60;
        if (startHour >= CONFIG.FRIDAY_END_HOUR || endHour <= CONFIG.BUSINESS_START_HOUR) {
            return { overlaps: false, reason: "outside ON Friday hours (7 AM–3 PM)" };
        }
        return { overlaps: true, reason: "" };
    }

    // Mon–Thu standard hours
    const startHour = start.getHours() + start.getMinutes() / 60;
    const endHour = end.getHours() + end.getMinutes() / 60;
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

function formatTimeShort(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

// ─── LOGGING ─────────────────────────────────────────────────────────────────

function log(message) {
    console.log(message);
}

// ─── SUMMARY EMAIL ────────────────────────────────────────────────────────────

/**
 * sendLogEmail() — sends the full run log to your personal Gmail.
 * One email per day. Subject gives you the key stats at a glance.
 * Flip CONFIG.CONTROL.LOG_EMAIL to false to stop receiving these.
 */
function sendLogEmail(runTime, processed, skipped, errors, logLines) {
    const dateStr = Utilities.formatDate(runTime, Session.getScriptTimeZone(), "EEE MM/dd");
    const status = errors > 0 ? "⚠️ ERROR" : processed > 0 ? "✅ OK" : "— idle";
    const subject = `CalendarSync — ${dateStr} — ${processed} processed, ${skipped} skipped ${status}`;
    const body = logLines.join("\n");

    GmailApp.sendEmail(CONFIG.PERSONAL_EMAIL, subject, body);
    log(`Summary email sent to ${CONFIG.PERSONAL_EMAIL}`);
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
        .filter((t) => t.getHandlerFunction() === "runSync")
        .forEach((t) => ScriptApp.deleteTrigger(t));

    ScriptApp.newTrigger("runSync").timeBased().atHour(17).nearMinute(30).everyDays(1).create();

    console.log("Trigger created: runSync daily at ~5:30 PM.");
}

/**
 * deleteTriggers() — removes all project triggers. Useful during testing.
 */
function deleteTriggers() {
    ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
    console.log("All triggers deleted.");
}
