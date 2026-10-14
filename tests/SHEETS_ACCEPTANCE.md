# Google Sheets acceptance test (about 1 hour, two people, two browsers)

**Status: NOT RUN YET.** The workbook has only been tested in Excel. `check_workbook.py` confirms that every function used exists in Google Sheets and that no conditional format points at another tab, but a real upload is the only proof. Do not call the desk production-ready until every row below passes.

## Setup

1. Upload `Weekday_Scheduling_Desk.xlsx` to Google Drive, then **File > Save as Google Sheets**.
2. **File > Settings > Calculation > Recalculation: "On change and every minute".** Without this, `NOW()` only updates when someone edits, so the lists will not move on their own as time passes.
3. Leave SETTINGS B3 (test clock) as shipped for checks 1-8. Clear it for check 9.

## Checks

| # | Check | How | Pass |
|---|---|---|---|
| 1 | No formula errors | Ctrl+F, "Also search within formulas" off: search `#REF`, `#NAME`, `#VALUE`, `#N/A`, `#ERROR` on every tab, including hidden DESK and LISTS (right-click the tab bar > unhide). | None found |
| 2 | Lists fill | RAM shows 25 rows. SHYAM shows "Your shift has not started" (test clock is 10:00). | As described |
| 3 | WhatsApp link | Click "Open WhatsApp" on RAM row 1, on a laptop and on a phone. | Chat opens with the message already typed |
| 4 | Calendar and Call links | Click Calendar 1. Click Call. | Calendar opens. If Call does nothing (Sheets may block `tel:` links), interns dial the number shown in column C; note it in START HERE |
| 5 | List updates after a log line | Copy RAM row 1's phone, add a LOG line: time (Ctrl+Shift+;), Ram, phone, NO_ANSWER. Time how long until row 1 changes. | The person leaves the list in under 5 seconds |
| 6 | Time entry | The line from check 5: column G shows the right date and time. | Correct |
| 7 | Dropdowns, IDs and checks | Outcome and Who dropdowns work. A LOG line with a wrong phone, or with a request ID that is not that person's, turns orange with a reason in column J. INPUT column K shows R001, R002, ... | All work |
| 8 | Two interns at once | Both people add a LOG line in the same minute, each typing on what they think is the next empty row. | Neither line is lost or overwritten |
| 9 | Live clock | Clear SETTINGS B3 and wait 2 minutes. | "Time now" on RAM moves; lists reflect the real time |
| 10 | Permissions | Share with interns as Editor, then **Data > Protect sheets and ranges**: lead-only for DASHBOARD, MESSAGES, PARTNERS, INPUT, SETTINGS, DESK, LISTS and LOG columns G-M. An intern account tries to edit DESK. | Blocked; LOG A-F and RAM D4 still editable |
| 11 | Speed with a full day | Paste 2,000 dummy LOG lines (copy the dry-run log from `tests/output/dry_run_day1.xlsx` several times). Repeat check 5. | Still under 10 seconds |

## If a check fails

- **Check 5 or 11 too slow:** reduce LOG to 1,000 rows and INPUT to 550 in `src/build_desk.py` (`LOGN`, `NROWS`), rebuild, re-upload.
- **Check 8 loses a line:** log through a Google Form instead of typing in LOG. A Form appends rows safely and stamps the time itself. Make the Form's questions Who, Phone, Outcome, Request ID, Notes, in that order, so its response sheet matches LOG columns A-F. Then rewrite LOG G-M as one `ARRAYFORMULA` each in row 2, so rows the Form adds are covered. Estimated effort: one hour.
- **Check 3 fails:** the message column (RAM column M) is the fallback. Open the chat and paste.

## Record

| Date | Tester | Checks passed | Failures and fixes |
|---|---|---|---|
| | | | |
