// Builds Weekday_Scheduling_Desk_Plan.docx (A4). Run from the repo root: node src/build_plan.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  AlignmentType, LevelFormat, HeadingLevel, Footer, PageNumber,
} = require("docx");

const OUT = path.join(__dirname, "..", "Weekday_Scheduling_Desk_Plan.docx");
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
      new Paragraph({ spacing: { after: 130 }, children: [run("What two interns can deliver, what it takes to reach 80%, and what happens if we can't", { color: GREY })] }),

      box([
        new Paragraph({ spacing: { after: 70 }, children: [run("Two interns working by hand will schedule about 45% of requests within 24 hours. 80% is reachable only if three things hold at once.", { bold: true, size: 22 })] }),
        new Paragraph({ spacing: { after: 40, line: 255 }, children: [run("(1) About 40% of candidates book by themselves, which is the job of an evening WhatsApp message (today's planning value is 25%). (2) Second and third tries keep about 70% of a first try's success. (3) About 5 interns. All three are measured in the first week. If any of them fails, the dashboard shows it before the day-6 decision, and we switch to the fallback in section 2 instead of hiring blindly.")] }),
      ]),

      h1("1. Why about 45% today"),
      table([6400, 1700, 2366], ["Step (one evening batch, worked the next day)", "Value", "Source"], [
        ["Unique requests / people behind them", "500 / ~430", "Brief; demo data"],
        ["Intern minutes before the 19:51 deadline (Ram 10:00-18:00, Shyam 12:00-19:51)", "951", "Shifts"],
        [[b("Tries two interns can make")], [b("~210")], "951 / 4.5 min per try"],
        [[b("Most they could book, even if every try booked")], [b("~49%")], "No behavioural assumption"],
        ["Realistic: about half of first tries say yes; a quarter of requests booked by candidates alone", [b("~45%")], "ASSUMED; measured day 1"],
      ], [1]),
      para([run("The simulated test day delivered 42-48% (without and with the evening message). That only shows the sheet works: the test bot uses the same assumed rates, so it is not evidence about real candidates.", { size: 17, color: GREY })], { spacing: { before: 60, after: 40 } }),

      h1("2. How the system reaches 80%, and what if it can't"),
      para("A plain arithmetic model (analysis/scenarios.py, no simulation) gives the share scheduled within 24 hours. The dashboard runs the same calculation live with the measured rates."),
      table([1500, 1300, 1700, 1050, 1050, 1050, 1300, 1516],
        ["Candidates self-book", "Try 1: says yes", "Later tries keep", "2 interns", "4 interns", "6 interns", "Best case", "Interns for 80%"], [
        ["25%", "45%", "100%", "47%", "70%", "86%", "86%", "5"],
        ["25%", "45%", "70%", "47%", "66%", "76%", "76%", [b("not reachable")]],
        ["25%", "45%", "40%", "47%", "62%", "67%", "67%", [b("not reachable")]],
        ["40%", "45%", "70%", "62%", "78%", "81%", "81%", "5"],
        ["40%", "35%", "70%", "57%", "70%", "75%", "75%", [b("not reachable")]],
        ["40%", "45%", "40%", "62%", "73%", "74%", "74%", [b("not reachable")]],
      ], [3, 4, 5, 6]),
      para([b("Reading it: "), run("extra interns help only while later tries still book. If the third try works far worse than the first, the best case stalls in the 60s and 70s whatever the staffing. Then only more self-booking moves the number.")], { spacing: { before: 70 } }),
      para([b("Three gates, each with a decision:")]),
      num([b("Days 1-2, two interns live "), run("(after the tests in section 5). The DASHBOARD readout measures minutes per try, the yes-rate on tries 1, 2 and 3, and the self-booking share, then recomputes the best case and the interns needed.")]),
      num([b("Days 3-5, evening-message pilot "), run("(section 3). It measures how much the message raises self-booking: the first row of the table above.")]),
      num([b("Day 6, decide. "), run("If the best case is 80% or more, staff to the number the dashboard gives. If it is below 80%, use the fallback.")]),
      para([b("Fallback if 80% is out of reach: "), run("(a) Take the measured best case to Weekday and agree what the KPI should count, using the dashboard's own numbers. For example: 80% of requests with usable contacts that did not opt out, or 80% within 48 hours. (b) Upstream fixes: re-confirm interest before a request is sent, and get slot guarantees from partners. (c) Keep the extra capacity that still pays for itself (section 4); it raises the result even below 80%. If no extra interns can be hired at all, the list still orders work by who is most likely to book and who is closest to the deadline, so two interns get the most out of their 210 tries.")]),

      h1("3. The evening WhatsApp: a hypothesis to pilot, not a fact"),
      para([run("Its cost (Rs 100-750 a day) and its effect are unproven. Here is what exists, what doesn't, and how to test it. "), b("Built: "),
            run("a lead-only MESSAGES tab with the two template texts and a send list for the provider's bulk upload. The send list stays empty until opt-in is confirmed, and candidates are split SEND/HOLD by a phone digit. There is a paste-back area for the provider's report: a YES reply puts the person at the top of an intern's list, a STOP reply closes them, and a failed delivery tells the intern to call only. A SEND-vs-HOLD readout compares 24-hour completion. In the test, all 6 STOP repliers stayed off the lists, and the 23 YES repliers who reached a list appeared in its top 18 rows. "),
            b("Not built: "), run("an automatic link to the WhatsApp API. The lead exports and imports files by hand once each evening.")]),
      table([3000, 7466], ["Needed before the first send", "Detail"], [
        ["WhatsApp Business account", "Through a provider (BSP). A few days to set up, plus a monthly platform fee."],
        ["Approved templates", "Meta approves each template and decides its class. Utility is about Rs 0.12 a message, marketing about Rs 0.86, so 430 people x 2 messages costs Rs 100-750 a day."],
        ["Candidate opt-in", "Weekday confirms candidates agreed to WhatsApp messages. Until then SETTINGS B22 stays NO and the send list stays empty."],
        ["Opt-outs", "Every template says \"Reply STOP\"; the sheet closes anyone who does. The provider's own block list is kept too."],
        ["Delivery failures and replies", "Come back in the provider's report. Pasting it in is the integration with the interns' lists."],
        ["Quality rating", "Meta throttles senders whose messages get blocked or reported. Watch the rating during the pilot."],
      ]),
      para([b("Pilot design: "), run("3 working days with consented candidates only, with SEND and HOLD both called as normal. That is about 750 requests per group, enough to see an 8-point difference in 24-hour completion. Bar: SEND at least 8 points above HOLD. Below that, the message does not carry the path to 80%.")], { spacing: { before: 70 } }),

      h1("4. Which decisions survive wrong assumptions"),
      table([2700, 3300, 4466], ["Decision", "Depends on", "Holds if"], [
        ["Use the sheet and its ordering", "Nothing behavioural", "Always: it organises the same work better."],
        ["Add interns", "Value of an interview vs cost of a try (Rs 9)", "An extra try pays if at least 1 in 244 tries books, even at a 1% hire rate and Rs 12L CTC (an interview worth Rs 1,800). At 3% and Rs 18L it is 1 in 1,111. Planning value for try 3: about 1 in 5. Robust."],
        ["Reach 80% by staffing", "Later-try success, self-booking", "Fragile: see the table in section 2. Measured on days 1-2."],
        ["Run the evening message", "Its uplift", "Unknown until the pilot. At up to Rs 750 a day it costs less than one intern-day."],
        ["\"Fees at risk\" figure", "Hire rate, CTC, share of late requests never interviewed", "Indicative only: Rs 1.2 lakh a day at a 1% hire rate, Rs 3.5 lakh at 3%, Rs 5.9 lakh at 5%. Use it to rank priorities, not as a forecast."],
      ]),

      h1("5. Before calling it production-ready"),
      table([2600, 5300, 2566], ["Gate", "What passes", "Status"], [
        ["30-minute usability test", "Two people who have never seen the file read START HERE (5 min), then work their list (20 min) while the lead plays 10 scripted candidates. Pass: 9 of 10 log lines right, at most 2 questions each, under 6 minutes per try.", "Not run"],
        ["Google Sheets acceptance", "11 checks in tests/SHEETS_ACCEPTANCE.md: formulas, links, the list updating within 5 seconds of a log line, two interns logging at once, sheet protection, a live clock (Sheets must recalculate every minute), and speed with a full day.", "Not run"],
        ["One controlled live day", "Two interns, real requests. The DASHBOARD readout shows minutes per try, yes-rate by try number, self-booking share and 24-hour completion, each next to its planning value.", "After the two gates above"],
      ]),

      h1("6. Status of the evidence"),
      table([3000, 7466], ["Item", "Status"], [
        ["Workbook mechanics (Excel)", "A full simulated day, run with and without the evening message: 0 formula errors, every log line checks OK, the sheet's try count equals the test bot's, at most 3 tries per person, no contact outside 10:00-21:00, 1.02-1.07 log lines per conversation."],
        ["Dashboard capacity model", "On the simulated days it ran 5-6 points above what was delivered (47% vs 42% without the evening message, 53% vs 48% with it). Treat its \"expected\" figure as an upper bound."],
        ["Google Sheets", "Every function used exists in Sheets. Not yet opened in Sheets."],
        ["Behavioural rates, evening-message effect, hire rate", "Assumptions until week 1 and the pilot."],
        ["Real interns, partners, candidates", "Not tested. Section 5 comes first."],
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("written", OUT, buf.length);
});
