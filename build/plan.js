// Builds Weekday Scheduling Plan.docx (A4). Run from the repo root: node build/plan.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  AlignmentType, LevelFormat, HeadingLevel, Footer, PageNumber,
} = require("docx");

const OUT = path.join(__dirname, "..", "Weekday Scheduling Plan.docx");
const FONT = "Arial";
const W = 10466; // A4 width 11906 - 2 x 720 margins
const NAVY = "1F3864", GREY = "595959", AMBER = "FFF2CC";

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 19, ...o });
const para = (children, o = {}) =>
  new Paragraph({ spacing: { after: 90, line: 260 }, ...o, children: Array.isArray(children) ? children : [run(children)] });
const h1 = (t) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1, keepNext: true, spacing: { before: 200, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 2 } },
    children: [new TextRun({ text: t, font: FONT, size: 23, bold: true, color: NAVY })],
  });
const bullet = (children) =>
  new Paragraph({
    numbering: { reference: "bul", level: 0 }, spacing: { after: 50, line: 255 },
    children: Array.isArray(children) ? children : [run(children)],
  });
const num = (children) =>
  new Paragraph({
    numbering: { reference: "num", level: 0 }, spacing: { after: 50, line: 255 },
    children: Array.isArray(children) ? children : [run(children)],
  });
const b = (t) => run(t, { bold: true });

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, width, o = {}) {
  const parts = Array.isArray(text) ? text : [run(text, { size: 17, bold: !!o.bold, color: o.color })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ spacing: { after: 0, line: 235 }, alignment: o.right ? AlignmentType.RIGHT : undefined, children: parts })],
  });
}
function table(widths, header, rows, rightCols = []) {
  const total = widths.reduce((a, c) => a + c, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: NAVY, color: "FFFFFF" })) }),
      ...rows.map((r, ri) => new TableRow({
        children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? "F2F2F2" : undefined, right: rightCols.includes(i) })),
      })),
    ],
  });
}
function box(lines) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, borders,
      shading: { fill: AMBER, type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 110, bottom: 110, left: 150, right: 150 },
      children: lines,
    })] })],
  });
}
const gap = () => new Paragraph({ spacing: { after: 60 }, children: [] });
const summ = (label, text) =>
  new Paragraph({ spacing: { after: 60, line: 250 }, children: [run(label + ".  ", { bold: true, color: NAVY }), run(text)] });

const doc = new Document({
  creator: "Weekday scheduling desk",
  title: "Weekday Scheduling Desk: plan",
  styles: { default: { document: { run: { font: FONT, size: 19 } } } },
  numbering: {
    config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 250 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 300 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Weekday scheduling desk  |  page ", font: FONT, size: 16, color: GREY }),
                 new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: [
      new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: "Weekday Scheduling Desk", font: FONT, size: 34, bold: true, color: NAVY })] }),
      new Paragraph({ spacing: { after: 120 }, children: [run("Summary, the evidence behind it, and what happens next", { color: GREY })] }),

      box([
        summ("Problem", "About 500 interview requests arrive each evening around 19:51. Weekday wants 80% scheduled within 24 hours by two interns, Ram and Shyam, following a process simple enough to use without judgement."),
        summ("What I built", "A spreadsheet that runs in Excel and should run in Google Sheets. Interns see four tabs: a ranked list per intern with one instruction per row, a pre-filled WhatsApp message and booking links, and a log of one line per conversation. Every request keeps its ID from the partner's sheet to the final booking, and a sent link never counts as a booking. The lead gets a dashboard that puts each assumption next to what the log measures, plus a tab for piloting an evening WhatsApp."),
        summ("Most defensible finding", "Two interns cannot reach 80% by phone. Before the deadline they can make about 210 tries for about 430 people, so even if every try booked, interns would schedule at most about 49% of requests [A]. Reaching 80% needs candidates to book at least 31% of requests on their own, more hands, or both. My planning estimate for two interns is about 45% [B]."),
        summ("Limitations", "No real intern, candidate or partner has used the sheet. It has not been opened in Google Sheets. Every behavioural rate is an assumption: booking per try, self-booking, and the evening message's effect. The simulated results come from those same assumptions, so they cannot confirm them."),
        summ("Next experiment", "First, a 30-minute test with two people new to the file, and the Google Sheets check. Then two live days with two interns, measuring three numbers: confirmed bookings by try number, minutes per try, and the share of candidates who book alone. Those three decide whether more interns can reach 80% at all, before anyone is hired."),
      ]),

      h1("1. How much to trust each number"),
      table([2350, 3700, 2500, 1916], ["Tier", "What it is", "Numbers in this document", "Trust"], [
        [[b("A  Analytical ceiling")], "Arithmetic from shift length and minutes per try. No assumption about candidates.", "~210 tries; interns book at most ~49%", "High; minutes per try is measured on day 1"],
        [[b("B  Planning estimate")], "The same arithmetic with assumed rates for booking per try and self-booking.", "~45% with two interns; the grid in section 3", "Low until day-1 data"],
        [[b("C  Simulated results")], "Candidate behaviour drawn at random from the same assumed rates, in the dry run.", "40% (no evening message), 49% (with)", "Low: cannot test the assumptions"],
        [[b("D  Synthetic dry-run results")], "Pass/fail checks of the sheet, run on fake data.", "All mechanics checks pass (section 6)", "High for the sheet; says nothing about candidates"],
        [[b("E  Untested intervention scenarios")], "What-ifs: the evening message, more interns, later tries holding up.", "Self-booking 25% to 40%; 5-6 interns", "Hypotheses until piloted"],
        [[b("F  Measured in real operations")], "Real interns, candidates, partners.", "None yet", "-"],
      ]),

      h1("2. The finding"),
      table([6400, 1700, 2366], ["Step (one evening batch, worked the next day)", "Value", "Tier"], [
        ["Unique requests / people behind them (demo batch)", "500 / ~430", "Brief; demo data"],
        ["Intern minutes before the 19:51 deadline (Ram 10:00-18:00, Shyam 12:00-19:51)", "951", "A"],
        [[b("Tries two interns can make, at 4.5 minutes per try")], [b("~210")], "A"],
        [[b("Most interns could book, even if every try booked (1.17 requests per person)")], [b("~49%")], "A"],
        ["With 40% of first tries confirmed (fewer later) and 25% of requests booked by candidates alone", [b("~45%")], "B"],
      ], [1]),
      para([run("\"Confirmed\" means booked by an intern, or booked by the candidate and logged after we sent the link. A sent link alone never counts. That stricter definition lowered the planning rate from 45% (\"said yes\") to 40%.", { size: 17, color: GREY })], { spacing: { before: 60, after: 40 } }),

      h1("3. What 80% would take, and the fallback"),
      para("The share scheduled within 24 hours, from the arithmetic in analysis/scenarios.py [B/E]. The dashboard runs the same calculation on live numbers."),
      table([1500, 1300, 1700, 1050, 1050, 1050, 1300, 1516],
        ["Candidates book alone", "Try 1 confirmed", "Later tries keep", "2 interns", "4 interns", "6 interns", "Best case", "Interns for 80%"], [
        ["25%", "40%", "100%", "45%", "65%", "82%", "82%", "6"],
        ["25%", "40%", "70%", "45%", "62%", "72%", "72%", [b("not reachable")]],
        ["25%", "40%", "40%", "45%", "58%", "63%", "63%", [b("not reachable")]],
        ["40%", "40%", "100%", "60%", "80%", "86%", "86%", "5"],
        ["40%", "40%", "70%", "60%", "74%", "78%", "78%", [b("not reachable")]],
        ["40%", "40%", "40%", "60%", "69%", "71%", "71%", [b("not reachable")]],
      ], [3, 4, 5, 6]),
      para([b("Reading it: "), run("80% needs second and third tries to work almost as well as first tries, plus 5-6 interns. If later tries lose even a third of their power, the best case is 72-78% however many interns are added. My own expectation is that later tries do lose power, so the likely outcome is a conversation with Weekday about the target, backed by measured numbers. That is an expectation, not a finding, and the next experiment tests it.")], { spacing: { before: 70 } }),
      para([b("Gates. "), run("Days 1-2: two interns live; the dashboard measures confirmed bookings by try number, minutes per try and self-booking. Days 3-5: evening-message pilot (section 5). Day 6: if the measured best case is 80% or more, staff to the dashboard's number. If it is below 80%, use the fallback.")]),
      para([b("Fallback. "), run("(a) Agree the KPI with Weekday on the measured best case: for example, 80% of requests with usable contacts that did not opt out, or 80% within 48 hours. (b) Upstream: re-confirm interest before a request is sent, and get slot guarantees from partners. (c) Keep any extra capacity that pays for itself (section 5). With no extra staff at all, the list still puts the likeliest bookings and the nearest deadlines first.")]),

      h1("4. What the interns do"),
      bullet([b("Their list: "), run("one row per person, best first. Each request on the row shows its ID (R001...), with one instruction, a WhatsApp link with the message typed in, and the partner's booking link.")]),
      bullet([b("One log line per conversation: "), run("time, name, phone, outcome. A blank request ID covers all of the person's open requests. A second line, naming the ID, is needed only when the answers differ by request.")]),
      bullet([b("No false promises, and consent first: "), run("no held-slot promises. Interns book in the partner's live calendar only where the partner agreed; otherwise they send the partner's link and check later.")]),

      h1("5. Supporting detail [E]: evening WhatsApp, costs, sensitivity"),
      para([b("Evening WhatsApp (untested). "), run("Built: a lead-only MESSAGES tab with template texts and a send list that stays empty until candidate opt-in is confirmed. Candidates are split SEND/HOLD for an A/B pilot. The provider's report is pasted back in: YES moves a person up an intern's list (it is not a booking), STOP closes them. Not built: any automatic link to the WhatsApp API. Before the first send it needs a provider account, Meta-approved templates and confirmed opt-in. Cost: Rs 100-750 a day, depending on how Meta classes the template. Pilot: 3 days, about 750 requests per group, with a bar of SEND at least 8 points above HOLD.")]),
      para([b("Cost and sensitivity. "), run("A try costs about Rs 9 of intern time. At Weekday's 15% fee, Rs 18 lakh CTC and a 3% hire rate, a scheduled interview is worth about Rs 8,100. So extra tries pay if at least 1 in 1,111 books, and still 1 in 244 at a 1% hire rate and Rs 12 lakh CTC. Adding interns is therefore a sound decision under any of these assumptions, as long as later tries book at all. What is fragile is whether staffing reaches 80% (section 3). \"Fees at risk\" from the gap ranges from Rs 1.2 to 5.9 lakh a day across hire rates of 1-5%, so it is indicative only.")]),

      h1("6. Gates before calling it production-ready, and test status [D]"),
      table([3000, 7466], ["Item", "Status"], [
        ["Request tracking", "29 of 29 scenario checks pass (tests/request_tracking.py). Covered: Round 1 and Round 2 at one company, three requests from one candidate, same company and round with two interviewers, a re-sent duplicate, a link that later converts, an explicit ID beating a blank line, wrong IDs flagged, NO_SLOTS by ID, and an ID used twice."],
        ["Simulated day (two runs)", "Every check passes. The requests the sheet shows as booked exactly match the requests the bot booked (198 and 244), with the right split between intern bookings and candidate bookings. No sent link was counted as a booking. Every log line checks OK, 0 formula errors, at most 3 tries per person, no contact outside 10:00-21:00, about 1.06 log lines per conversation."],
        ["Dashboard model vs simulated day [C]", "41% expected vs 40% delivered without the message; 50% vs 49% with it. Close, but both sides rest on the same assumed rates."],
        ["30-minute usability test", "Not run. Pass: 9 of 10 log lines right, at most 2 questions per person, under 6 minutes per try."],
        ["Google Sheets acceptance", "Not run. 11 checks in tests/google-sheets-checklist.md. Every function used exists in Sheets."],
        ["One controlled live day", "After the two tests above. This produces the first tier-F numbers."],
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("written", OUT, buf.length);
});
