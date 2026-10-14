# Weekday Scheduling Desk

This is my take on the Weekday scheduling brief: around 500 interview requests land every evening, and two interns (Ram and Shyam) have to get 80% of them scheduled within 24 hours, using a process simple enough to follow without judgement calls.

If you only open two files, make it these:

- **`Weekday Scheduling Desk.xlsx`**: the sheet the interns work from. It ships with made-up demo data and a test clock set to 14 Oct, 10:00, so the lists are already filled when you open it.
- **`Weekday Scheduling Plan.docx`**: two pages. Page one is the summary: the problem, what I built, the finding I'm most confident in, what I haven't tested, and the next experiment. The rest is the evidence behind it.

## The short version

Two interns can't reach 80% by phone. Before the deadline they have time for about 210 tries, and there are about 430 people to reach. So even if every single call ended in a booking, interns would schedule at most about 49% of requests. My working estimate is about 45%. Getting to 80% needs candidates booking on their own, more people, or both, and whether more people would even help depends on numbers I can only get from two live days. The plan lays out what to measure and what to do with the answer.

## How the sheet is laid out

- **Interns:** START HERE, RAM, SHYAM and LOG. Each intern gets a ranked list with one instruction per row, a WhatsApp message that's already typed, and the partner's booking link. They log one line per conversation.
- **Lead:** DASHBOARD, MESSAGES (for piloting an evening WhatsApp), PARTNERS, INPUT and SETTINGS.
- **Hidden:** DESK (the engine) and LISTS (outcome codes and message text).

Every request gets an ID (R001, R002, ...) that follows it from the partner's sheet to the booking. A sent booking link never counts as a booking.

## Where it stands

It isn't ready for live use yet. Everything has been tested in Excel on made-up data, but two things still need real people:

1. a 30-minute run with two people who have never seen the file, and
2. the Google Sheets checklist in `tests/google-sheets-checklist.md`.

## What's in the folder

| Path | What it is |
|---|---|
| `build/workbook.py` | Builds the workbook. Everything in it is live formulas. |
| `build/plan.js` | Builds the plan document. |
| `analysis/scenarios.py` | The "what would 80% take" table and the cost sensitivity. Plain arithmetic, so it can be checked by hand. |
| `analysis/simulation.py` | A simulation of candidate behaviour. Here it only generates the demo data. |
| `tests/check_workbook.py` | Google Sheets compatibility, message wording, formula errors and WhatsApp link lengths. |
| `tests/request_tracking.py` | 29 checks that each request keeps its ID through repeated rounds, multiple requests per person, duplicates and partial answers. |
| `tests/dry_run.py` | A bot works one simulated day through the sheet. The pass/fail checks are what matter; the booking numbers it produces come from assumptions. |
| `tests/google-sheets-checklist.md` | The checklist for the first run in Google Sheets. |
| `tests/results/` | Results from the last test runs, including `demo-candidates.json` (how each made-up candidate behaves). |

## Rebuilding and testing

Needs Python with openpyxl, numpy and pywin32, Excel on Windows for the tests, and Node with the `docx` package for the plan.

```
python build/workbook.py
python tests/check_workbook.py
python tests/request_tracking.py
python tests/dry_run.py                  # with the evening WhatsApp
EVENING_MSG=0 python tests/dry_run.py    # without it
python analysis/scenarios.py
node build/plan.js
```

## Going live

1. In SETTINGS, clear the test clock (B3).
2. In INPUT, delete the demo rows and paste the partner sheet into columns A-I. Column K gives each request its ID; paste the partner's own IDs over it if they have them. After interns start logging, only add rows at the bottom and don't sort.
3. In PARTNERS, add each company's booking link. Only mark "We may book for candidates?" as YES once the partner has agreed in writing.
4. Work through `tests/google-sheets-checklist.md`. Turn on "recalculate every minute" in Sheets and protect the lead tabs.
5. Run the 30-minute test with two new people (plan, section 6).
6. Keep the evening WhatsApp switched off (SETTINGS B22 = NO) until candidates have opted in, there's a provider account and the templates are approved (plan, section 5).
