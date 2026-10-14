// Builds Weekday_Scheduling_Desk_Plan.docx (A4, two pages). Run from the repo root: node src/build_plan.js
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
      new Paragraph({ spacing: { after: 130 }, children: [run("What two interns can do with a 24-hour target, and what 80% would cost", { color: GREY })] }),

      box([
        new Paragraph({ spacing: { after: 70 }, children: [run("Two interns working by hand will schedule about 45% of requests within 24 hours, not 80%.", { bold: true, size: 23 })] }),
        new Paragraph({ spacing: { after: 50, line: 255 }, children: [run("This follows from arithmetic, not from a model. The gap is cheap to close up to a point: one booked interview costs about Rs 20 of intern time and is worth about Rs 8,000 to Weekday. So buy capacity, starting with an automated evening WhatsApp. Whether 80% is reachable at all depends on one number we can measure on day 1: how often the second and third tries still book.")] }),
      ]),

      h1("1. Why about 45%"),
      table([6400, 1700, 2366], ["Step (one typical evening batch, worked the next day)", "Value", "Source"], [
        ["Unique requests after removing duplicates", "500", "Brief"],
        ["People behind them (some have 2-3 requests)", "~430", "Demo data; measure"],
        ["Intern minutes before the 19:51 deadline (Ram 10:00-18:00, Shyam 12:00-19:51)", "951", "Shifts"],
        ["Minutes per try (call, then WhatsApp if no answer)", "4.5", "Time 20 tries on day 1"],
        [[b("Tries two interns can make")], [b("~210")], "951 / 4.5"],
        [[b("Most they could book, even if every single try booked")], [b("~49%")], "210 people x 1.17 requests"],
        ["Realistic: about half of tries book, and about a quarter of requests are booked by the candidate alone", [b("~45%")], "Both ASSUMED; measured on day 1"],
      ], [1]),
      para([run("The ceiling of ~49% needs no assumption about candidates. To reach 80% with two interns, candidates would have to book over 30% of requests themselves, even if every call succeeded. ", { size: 18 }),
            run("I do not quote simulation results as evidence here: the simulator and the dry-run bot use the same assumed rates, so they cannot confirm them.", { size: 18, color: GREY })], { spacing: { before: 70, after: 60 } }),

      h1("2. What closing the gap costs"),
      table([6400, 1700, 2366], ["Item", "Per day", "Basis"], [
        ["One intern (Rs 20,000 a month, 22 days)", "Rs 909", "Advertised Bengaluru stipends"],
        ["Cost of one try / of one interview booked by an intern", "Rs 9 / Rs 20", "~105 tries per intern-day"],
        ["Expected value of one scheduled interview to Weekday", "Rs 8,100", "15% fee x Rs 18L CTC x 3% hire rate"],
        ["Gap to 80% with two interns", "~175 requests", "80% - 45% of 500"],
        ["Fees at risk if a quarter of late requests are never interviewed", "~Rs 3.5 lakh", "ASSUMED; formula on DASHBOARD"],
        ["Close the gap with interns, if later tries keep booking", "+4 to 5 interns, ~Rs 4,000", "~550-730 tries needed"],
        ["Automated evening WhatsApp to everyone (2 messages each)", "~Rs 750", "Rs 0.86 per message + provider fee"],
      ], [1]),
      para([b("When is extra capacity worth it? "), run("While an extra intern books more than one interview a week, it pays for itself. At ~Rs 20 per booking against ~Rs 8,000 of value, that holds even if my hire-rate guess is ten times too high. Capacity stops paying when everyone has already had their three tries. After that, only reply rates move the number.")], { spacing: { before: 70 } }),
      para([b("Recommended order. "), run("(1) An automated evening WhatsApp at about 20:00, when candidates are free and interns are not working: about Rs 750 a day, no intern time. (2) Two more interns for week 1 (Rs 1,800 a day). (3) After day 2, read the DASHBOARD row \"Do later tries still work?\". If try 3 still books, add interns until the verdict says the target is reachable. If it doesn't, stop hiring and work on reply rates and the target.")]),

      h1("3. What Ram and Shyam do"),
      para("The workbook (it also loads in Google Sheets) gives interns four tabs: START HERE, RAM, SHYAM and LOG. START HERE is four steps plus a table of nine outcome codes. Everything else is the lead's."),
      bullet([b("Their list: "), run("one row per person, best first, with a column that says exactly what to do, plus Call, WhatsApp and calendar links. Rows close to the deadline turn yellow.")]),
      bullet([b("One log line per conversation: "), run("time, name, phone, outcome. A second line is needed only when the candidate gives different answers for two companies. On a simulated day this averaged 1.02 lines per conversation.")]),
      bullet([b("No promises the process can't keep: "), run("messages never mention a held slot. Interns book in the partner's live calendar during the call, so two candidates can't be given the same slot.")]),
      bullet([b("Booking for the candidate only with consent: "), run("each partner is marked YES or NO in PARTNERS, and blank counts as NO. For NO partners the message carries the partner's own booking link, and a follow-up checks whether the candidate booked.")]),
      bullet([b("WhatsApp messages are pre-filled: "), run("messages are short enough that the link fits Excel's 255-character limit even with a long company name and booking link (25 of 25 in the test).")]),

      h1("4. What is not known yet, and how week 1 finds out"),
      table([3300, 4300, 2866], ["Unknown", "How we measure it", "When"], [
        ["Minutes per try", "Time 20 real tries", "Day 1"],
        ["Bookings per try, by try number", "DASHBOARD reads it from LOG", "Day 1-2"],
        ["Share candidates book alone", "Scheduled by candidate / unique requests", "Day 2 morning"],
        ["Partner consent to book for candidates", "Lead asks the top 10 partners by volume; record YES in PARTNERS", "Before day 1"],
        ["Hire rate and real CTC", "Weekday's own funnel data into SETTINGS", "Week 1"],
        ["Google Sheets behaviour", "Upload the file. Check that the lists fill, the WhatsApp links open, and LOG accepts a line. About 1 hour.", "Before day 1"],
      ]),

      h1("5. Before go-live: the 30-minute test that matters most"),
      para("This is not done yet. It needs two people who have never seen the file, and a lead who plays the candidates from a card."),
      num([b("5 minutes: "), run("they read START HERE only. No coaching.")]),
      num([b("20 minutes: "), run("each works their list. The lead answers as 10 scripted candidates: no answer; yes; yes to one company and no to another; stop contacting me; wrong number; replies YES later on WhatsApp; partner calendar full; and three more yeses.")]),
      num([b("5 minutes: "), run("where did you hesitate? Record every question asked.")]),
      para([b("Pass: "), run("at least 9 of 10 log lines correct (LOG column J shows OK and the outcome is right), at most 2 questions per person, and under 6 minutes per try. Fail: rewrite the wording they stumbled on and run it again with two new people.")]),

      h1("6. Status of the evidence"),
      table([3000, 7466], ["Item", "Status"], [
        ["Workbook mechanics", "Tested on a full simulated day: 0 formula errors, every log line checks OK, the sheet's try count equals the test bot's, nobody tried more than 3 times, no contact outside 10:00-21:00."],
        ["Google Sheets", "Every function used exists in Sheets, and no conditional format points at another tab (Sheets rejects that). Not yet opened in Sheets."],
        ["Behavioural rates", "Assumptions. The dashboard switches to the measured bookings per try after 30 tries; the lead updates the other planning values in SETTINGS from day-1 data."],
        ["Real interns, partners, candidates", "Not tested. Section 5 is the first test."],
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("written", OUT, buf.length);
});
