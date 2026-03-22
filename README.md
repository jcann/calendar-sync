# Google Calendar → Work Outlook Automation

Auto-blocks time on your work Outlook calendar when you add appointments to
Google Calendar. Adds 15-minute commute buffers. No sensitive details leak to
work. Built for DoD contractor environment (no direct Google sync, firewall
restrictions, Outlook BYOD blocked).

## Architecture

```
Personal / Kids Calendar
        │
        ▼ (every 15 min, Mon–Fri 7 AM–5 PM)
  Google Apps Script
        │
        ├── Copies "Busy" block (with buffers) → Automation Calendar
        └── (Stage 4+) Sends email → Work Outlook
```

## Current Stage: Stage 1 — Safe Mode

| Feature                        | Status |
|-------------------------------|--------|
| Detect events                 | ✅ ON  |
| Log to console                | ✅ ON  |
| Copy to Automation calendar   | ✅ ON  |
| Commute buffers (+/– 15 min)  | ✅ ON  |
| Duplicate prevention          | ✅ ON  |
| Kids calendar                 | ⏸ OFF  |
| Send emails                   | ⏸ OFF  |

## Setup

### 1. Create the Automation Calendar
- Open Google Calendar → `+` next to "Other calendars" → Create new calendar
- Name it `Automation`
- Go to its Settings → copy the **Calendar ID** (looks like `abc123@group.calendar.google.com`)

### 2. Get your Kids Calendar ID
- Open Google Calendar → Settings for the shared kids calendar
- Copy the **Calendar ID**

### 3. Create the Apps Script project
- Go to [script.google.com](https://script.google.com) → New project
- Rename it to `CalendarSync`
- Paste the contents of `CalendarSync.gs` into the editor

### 4. Fill in CONFIG
Edit the top of `CalendarSync.gs`:

```javascript
PERSONAL_CALENDAR_ID:   "primary",               // or your Gmail address
KIDS_CALENDAR_ID:       "paste-here@group...",   // from step 2
AUTOMATION_CALENDAR_ID: "paste-here@group...",   // from step 1
```

### 5. Run manually first (no trigger yet)
- In the Apps Script editor, select `runSync` from the function dropdown
- Click ▶ Run
- Accept OAuth permissions when prompted
- Check View → Logs to confirm events are found and processed

### 6. Create the time trigger
- In the editor, select `createTrigger` and click ▶ Run
- Go to the clock icon (Triggers) in the left sidebar to verify

## Advancing Stages

Edit the `STAGE` block in `CONFIG`:

```javascript
STAGE: {
  COPY_TO_AUTOMATION: true,   // Stage 1
  SEND_EMAIL:         false,  // Stage 4+
  INCLUDE_KIDS:       false,  // Stage 3+
},
```

| Stage | Config change                                  | Verify before proceeding                   |
|-------|------------------------------------------------|--------------------------------------------|
| 1     | (default)                                      | Check Automation calendar, check logs      |
| 2     | None — observe for 2–3 days                    | No duplicate events                        |
| 3     | `INCLUDE_KIDS: true`                           | Kids events copy correctly, no duplication |
| 4     | `SEND_EMAIL: true`, `EMAIL.TO` = test address  | Email arrives, no spam, correct format     |
| 5     | `EMAIL.TO` = work Outlook address              | Production                                 |

## GitHub Workflow

### First push
```bash
# Clone or init
git init calendar-sync && cd calendar-sync

# Copy script
cp /path/to/CalendarSync.gs .

# Commit
git add .
git commit -m "Stage 1: safe mode — copy + log, no emails"
git remote add origin https://github.com/YOUR_USERNAME/calendar-sync.git
git push -u origin main
```

### Syncing changes from Apps Script editor → GitHub

Apps Script doesn't push to GitHub automatically. Use this workflow:

1. Edit in Apps Script editor
2. Copy the full file content
3. Paste into your local `CalendarSync.gs`
4. Commit with a meaningful message:

```bash
git add CalendarSync.gs
git commit -m "Stage 2: verified no duplicates after 3-day observation"
git push
```

### Using clasp (optional — full local dev)
[clasp](https://github.com/google/clasp) lets you push/pull Apps Script from the CLI.

```bash
npm install -g @google/clasp
clasp login
clasp clone YOUR_SCRIPT_ID   # found in Apps Script → Project Settings
```

Then edit locally and push:
```bash
clasp push
```

## Security Notes

- Events copied to Automation calendar use generic title `"Busy"` — no personal details
- Original event title is stored in the Automation event **description only**, not visible in Outlook
- The `[AUTO-PROCESSED]` tag is appended to source event descriptions to prevent re-processing
- No Google credentials or calendar IDs should be committed to a public repo — consider making the repo private

## Files

| File             | Purpose                                   |
|-----------------|-------------------------------------------|
| `CalendarSync.gs`| Main Apps Script — all logic lives here  |
| `README.md`      | This file                                |

## Troubleshooting

**No events detected**
- Check `PERSONAL_CALENDAR_ID` is correct (use your full Gmail if `"primary"` doesn't work)
- Verify events are within the next 7 days and during business hours
- Run `runSync()` manually and check View → Logs

**Duplicate events on Automation calendar**
- The script checks for duplicates before creating; if you see duplicates, check that
  `AUTOMATION_CALENDAR_ID` is correct (wrong ID = writing to wrong calendar)

**PROCESSED tag not sticking**
- Some recurring events require `event.setDescription()` permissions — re-authorize
  via script.google.com → Services → re-run and accept OAuth

**Trigger not firing**
- Confirm trigger exists: Apps Script → ⏰ Triggers → `runSync` every 15 min
- Check Apps Script → ⚠️ Executions for errors
