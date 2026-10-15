# Weekday Scheduling Desk

This is my take on the Weekday scheduling brief: around 500 interview requests land every evening, and two interns (Ram and Shyam) have to get 80% of them scheduled within 24 hours, using a process simple enough to follow without judgement calls.

If you only open two files, make it these:

- **`Weekday Scheduling Desk.xlsx`**: the sheet the interns work from. It ships with made-up demo data and a test clock set to 14 Oct, 10:00, so the lists are already filled when you open it.
- **`Weekday Scheduling Plan.pdf`**: the write-up, It starts with the finding (two interns top out near 49% even if every call books), works through the assignment's parts (clarifying questions with assumptions, the breakdown, options, workflow, the interns' to-do sheet, messages, edge cases and the no-show bonus, KPIs)

## The short version

Two interns can't reach 80% by phone. Before the deadline they have time for about 210 tries, and 426 people sit behind the 500 requests. Even if every single call ended in a booking, each settling the average 1.17 requests, interns would book at most about 49% of requests. My working estimate is about 45%. Every share is a share of the 500 requests. Getting to 80% needs candidates booking on their own, more people, or both, and whether more people would even help depends on numbers I can only get from two live days. The plan lays out what to measure and what to do with the answer.

## How the sheet is laid out

- **Interns:** START HERE, RAM, SHYAM and LOG. Each intern gets a ranked list with one instruction per row, a WhatsApp message that's already typed, and the partner's booking link. They log one line per conversation.
- **Lead:** DASHBOARD, MESSAGES (for piloting an evening WhatsApp), PARTNERS, INPUT and SETTINGS.
- **Hidden:** DESK (the engine) and LISTS (outcome codes and message text).

Every request gets an ID (R001, R002, ...) that follows it from the partner's sheet to the booking. A sent booking link never counts as a booking.

## What's in the folder

| Path | What it is |
|---|---|
| `analysis/scenarios.py` | The "what would 80% take" table, the capacity ceiling and its outer bounds, the fee scenario and the cost sensitivity. One assumption set, plain arithmetic, every share a share of the 500 requests. The DASHBOARD MODEL block does the same arithmetic in formulas. |
| `analysis/simulation.py` | A simulation of candidate behaviour. Here it only generates the demo data. |
| `tests/request_tracking.py` | 29 checks that each request keeps its ID through repeated rounds, multiple requests per person, duplicates and partial answers. |
| `tests/model_consistency.py` | Rebuilds the ceiling and its bounds from the raw demo batch, then checks the DASHBOARD model against `analysis/scenarios.py` on the same batch for seven scenarios, and reports any gap. |
| `tests/dry_run.py` | A bot works one simulated day through the sheet. The pass/fail checks are what matter; the booking numbers it produces come from assumptions. |
| `archive/` | The workbook as it was before the dashboard moved to the request-level model (15 Oct). |
| `tests/results/` | Results from the last test runs, including `demo-candidates.json` (how each made-up candidate behaves). |

## Going live

1. In SETTINGS, clear the test clock (B3). Until it's blank, a red TEST CLOCK warning shows on START HERE, RAM, SHYAM and DASHBOARD.
2. In INPUT, delete the demo rows and paste the partner sheet into columns A-I. Column K gives each request its ID; paste the partner's own IDs over it if they have them. After interns start logging, only add rows at the bottom and don't sort.
3. In PARTNERS, add each company's booking link. Only mark "We may book for candidates?" as YES once the partner has agreed in writing.
