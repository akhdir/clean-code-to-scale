const pptxgen = require("pptxgenjs");

// ---------- palette (shared with the course's web materials) ----------
const DARK = "0F3038";
const DARK2 = "16414B";
const DARK3 = "123943";
const DARK_WATERMARK = "1B4750";
const TEAL = "1F5D6B";
const TEAL_TINT = "DCEAEA";
const PLUM = "6B4A8F";
const PLUM_TINT = "EAE2F1";
const AMBER = "A6752B";
const AMBER_LIGHT = "E0A85C";
const AMBER_TINT = "F3E6CE";
const RUST = "B3452F";
const RUST_TINT = "F5E3DE";
const PAPER = "F4F6F5";
const WHITE = "FFFFFF";
const INK = "1A2226";
const INK_SOFT = "55636A";
const INK_FAINT = "8B979B";
const LINE = "D6DEDD";
const CODE_BG = "101A1D";
const CODE_INK = "DCEAEA";
const CODE_COMMENT = "6FAFAE";
const SLATE = "5C7A85";

const HEAD = "Cambria";
const BODY = "Calibri";
const CODE = "Courier New";
const FOOT = "Day 08 · Clean Code to Scale";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// slides are queued so footers can show "NN / TOTAL"
const queue = [];
const queueNotes = [];
let pendingNotes = null;
function N(en, ar) { pendingNotes = [en, ar]; }
function add(fn) { queue.push(fn); queueNotes.push(pendingNotes || [null, null]); pendingNotes = null; }
let TOTAL = 0;

const allSlides = [];
function newSlide(dark) {
  const s = pptx.addSlide();
  s.background = { color: dark ? DARK : PAPER };
  s._notes = [];
  s._applyNotes = s.addNotes.bind(s);
  s.addNotes = (t) => { s._notes.push(t); };
  allSlides.push(s);
  return s;
}

function footer(s, num, dark, label) {
  s.addText(label || FOOT, { x: 0.5, y: 7.12, w: 8, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 9, color: dark ? "6C8B90" : INK_FAINT });
  s.addText(String(num).padStart(2, "0") + " / " + TOTAL, { x: 11.83, y: 7.12, w: 1, h: 0.3, isTextBox: true, margin: 0, align: "right", fontFace: CODE, fontSize: 9, color: dark ? "6C8B90" : INK_FAINT });
}

function header(s, kicker, title, dark, o = {}) {
  s._title = title;
  s.addText(kicker.toUpperCase(), { x: 0.5, y: 0.42, w: 12.3, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 12, color: dark ? AMBER_LIGHT : TEAL, charSpacing: 1 });
  s.addText(title, { x: 0.5, y: 0.72, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: o.size || 30, color: dark ? WHITE : INK });
  if (o.sub) s.addText(o.sub, { x: 0.5, y: 1.32, w: 12.3, h: 0.34, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: o.subSize || 12, color: dark ? "6C8B90" : INK_SOFT });
}

function circle(s, x, y, d, fill, label, o = {}) {
  s.addShape(pptx.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  s.addText(label, { x, y, w: d, h: d, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: o.fontFace || CODE, bold: true, fontSize: o.fontSize || 13, color: o.color || WHITE });
}

function card(s, x, y, w, h, fill, o = {}) {
  s.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lineW || 1 } : { type: "none" },
    shadow: o.shadow === false ? undefined : { type: "outer", color: "1A2226", opacity: 0.12, blur: 8, offset: 2, angle: 90 },
  });
}

function codePanel(s, x, y, w, h, fill) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || CODE_BG }, line: { type: "none" } });
}

// "  # comment" and leading "# comment" render in comment colour; "#!" renders amber
function runsFrom(lines) {
  const runs = [];
  lines.forEach((ln, i) => {
    const brk = i < lines.length - 1;
    let code = ln, com = "";
    if (ln.trimStart().startsWith("#")) { code = ""; com = ln; }
    else { const m = ln.match(/\s{2,}#/); if (m) { code = ln.slice(0, m.index); com = ln.slice(m.index); } }
    const warn = com.includes("#!");
    if (warn) com = com.replace("#!", "#");
    const cc = warn ? AMBER_LIGHT : CODE_COMMENT;
    if (code !== "" && com !== "") {
      runs.push({ text: code, options: { color: CODE_INK } });
      runs.push({ text: com, options: { color: cc, breakLine: brk } });
    } else if (com !== "") {
      runs.push({ text: com, options: { color: cc, breakLine: brk } });
    } else {
      runs.push({ text: ln === "" ? "" : code, options: { color: CODE_INK, breakLine: brk } });
    }
  });
  return runs;
}

function codeBlock(s, x, y, w, h, label, labelColor, lines, fs) {
  s.addText(label.toUpperCase(), { x: x + 0.05, y, w: w - 0.1, h: 0.24, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 9.5, color: labelColor, charSpacing: 0.5 });
  const py = y + 0.26, ph = h - 0.26;
  codePanel(s, x, py, w, ph);
  s.addText(runsFrom(lines), { x: x + 0.18, y: py + 0.08, w: w - 0.36, h: ph - 0.16, isTextBox: true, margin: 0, fontFace: CODE, fontSize: fs || 9.5, lineSpacingMultiple: 1.05, valign: "middle" });
}

function arrow(s, x1, y1, x2, y2, color, o = {}) {
  const dx = x2 - x1, dy = y2 - y1;
  s.addShape(pptx.ShapeType.line, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(dx), h: Math.abs(dy),
    flipH: dx < 0, flipV: dy < 0,
    line: { color: color || INK_SOFT, width: o.w || 1.25, endArrowType: o.noHead ? undefined : "triangle", dashType: o.dash || "solid" },
  });
}

function box(s, x, y, w, h, fill, line, lw) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.07, fill: fill ? { color: fill } : { type: "none" }, line: { color: line || INK_SOFT, width: lw || 1.25 } });
}

function label(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x, y, w, h, isTextBox: true, margin: 0, align: o.align || "center", valign: o.valign || "middle", fontFace: o.font || CODE, bold: !!o.bold, italic: !!o.italic, fontSize: o.size || 10, color: o.color || INK, breakLine: false });
}

// generic table; first row is the header
function tbl(s, x, y, colW, rows, o = {}) {
  const dark = !!o.dark, fs = o.fs || 11;
  const body = rows.map((r, ri) => r.map((cell, ci) => {
    const head = ri === 0;
    return {
      text: cell,
      options: {
        bold: head || ci === 0,
        fontFace: BODY, fontSize: head ? fs - 0.5 : fs,
        color: head ? WHITE : dark ? (ci === 0 ? WHITE : "C6DADC") : (ci === 0 ? INK : INK_SOFT),
        fill: { color: head ? TEAL : dark ? (ri % 2 ? DARK2 : DARK3) : (ri % 2 ? WHITE : "EEF3F2") },
        valign: "middle",
      },
    };
  }));
  s.addTable(body, { x, y, w: colW.reduce((a, b) => a + b, 0), colW, margin: [0.05, 0.1, 0.05, 0.1], border: { type: "solid", pt: 0.75, color: dark ? "2B5560" : LINE }, rowH: o.rowH });
}


// bullet/number runs with explicit font settings (renderers fall back to a serif otherwise)
function bl(items, o = {}) {
  const fs = o.fs || 13, color = o.color || INK_SOFT, gap = o.gap == null ? 9 : o.gap;
  return items.map((t, i) => ({
    text: t,
    options: {
      fontFace: BODY, fontSize: fs, color, breakLine: i < items.length - 1, paraSpaceAfter: gap,
      bullet: o.number ? { type: "number", indent: 22 } : { code: "25AA", indent: 16 },
    },
  }));
}
function qBlock(s, x, y, w, h, title, items, fs) {
  const runs = [{ text: title, options: { fontFace: HEAD, fontSize: 15, bold: true, color: TEAL, breakLine: true, paraSpaceAfter: 6 } }]
    .concat(bl(items, { number: true, fs: fs || 12, color: INK, gap: 8 }));
  s.addText(runs, { x, y, w, h, isTextBox: true, margin: 0, valign: "top" });
}

function banner(s, text, x, y, w) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.32, rectRadius: 0.16, fill: { color: AMBER_TINT }, line: { color: AMBER, width: 1 } });
  s.addText(text.toUpperCase(), { x, y, w, h: 0.32, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: CODE, bold: true, fontSize: 10, color: AMBER, charSpacing: 1 });
}

function divider(part, time, title, blurb) {
  add((n) => {
    const s = newSlide(true);
    s.addText(String(part), { x: 0.6, y: 0.9, w: 5, h: 3.5, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 200, color: DARK_WATERMARK });
    s.addText(("Part " + part + "  ·  " + time).toUpperCase(), { x: 0.7, y: 3.6, w: 11, h: 0.4, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 14, color: AMBER_LIGHT, charSpacing: 1.5 });
    s.addText(title, { x: 0.7, y: 4.05, w: 11.8, h: 1.0, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 40, color: WHITE });
    s.addText(blurb, { x: 0.7, y: 5.1, w: 10.5, h: 1.0, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 16, color: "B9D3D6" });
    footer(s, n, true);
  });
}

// prompts: array of strings. code: optional {label, lines, fs}. notes: answer key text.
function checkIn(num, mins, title, intro, prompts, code, notes, kicker) {
  add((n) => {
    const s = newSlide(false);
    banner(s, "Check-in " + num + " · " + mins + " min", 0.5, 0.42, 2.6);
    s.addText(title, { x: 0.5, y: 0.85, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 30, color: INK });
    s.addText(intro, { x: 0.5, y: 1.55, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK_SOFT, valign: "top" });
    if (code) {
      codeBlock(s, 0.5, 2.3, 6.9, 4.6, code.label, TEAL, code.lines, code.fs);
      card(s, 7.65, 2.3, 5.2, 4.6, WHITE);
      s.addText(bl(prompts, { number: true, fs: 15, color: INK, gap: 14 }), { x: 7.9, y: 2.5, w: 4.75, h: 4.2, isTextBox: true, margin: 0, valign: "top" });
    } else {
      card(s, 0.5, 2.4, 12.35, 4.4, WHITE);
      s.addText(bl(prompts, { number: true, fs: 20, color: INK, gap: 18 }), { x: 0.9, y: 2.75, w: 11.4, h: 3.9, isTextBox: true, margin: 0, valign: "top" });
    }
    s.addText("Answers are in the speaker notes", { x: 0.5, y: 6.95, w: 6, h: 0.2, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 9, color: INK_FAINT });
    footer(s, n, false);
    s.addNotes("ANSWER KEY — Check-in " + num + ": " + title + "\n\n" + notes);
  });
}

// ------------------------------------------------------------ simple slide kinds
function cards3(kicker, title, sub, cols, notes, dark, foot) {
  add((n) => {
    const s = newSlide(!!dark);
    header(s, kicker, title, !!dark, { sub });
    const cw = 3.95, gx = 0.25, ox = 0.5, oy = 1.95, ch = foot ? 3.55 : 4.8;
    cols.forEach((c, i) => {
      const x = ox + i * (cw + gx);
      card(s, x, oy, cw, ch, dark ? DARK2 : WHITE);
      s.addShape(pptx.ShapeType.rect, { x, y: oy, w: cw, h: 0.09, fill: { color: c.color }, line: { type: "none" } });
      s.addText(c.title, { x: x + 0.3, y: oy + 0.3, w: cw - 0.6, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: dark ? WHITE : INK });
      s.addText(bl(c.points, { fs: 14, color: dark ? "C6DADC" : INK_SOFT, gap: 10 }), { x: x + 0.3, y: oy + 0.95, w: cw - 0.6, h: ch - 1.2, isTextBox: true, margin: 0, valign: "top" });
    });
    if (foot) callout(s, 0.5, oy + ch + 0.25, 12.35, 1.05, foot[0], foot[1], !!dark);
    footer(s, n, !!dark);
    if (notes) s.addNotes(notes);
  });
}

function callout(s, x, y, w, h, head, body, dark) {
  card(s, x, y, w, h, dark ? DARK2 : WHITE, { shadow: false, line: dark ? undefined : LINE });
  s.addText([
    { text: head + "  ", options: { bold: true, color: dark ? AMBER_LIGHT : TEAL, fontFace: BODY, fontSize: 13 } },
    { text: body, options: { color: dark ? "C6DADC" : INK_SOFT, fontFace: BODY, fontSize: 13 } },
  ], { x: x + 0.25, y, w: w - 0.5, h, isTextBox: true, margin: 0, valign: "middle" });
}

function infoCard(s, x, y, w, h, color, title, items, fs) {
  card(s, x, y, w, h, WHITE);
  s.addShape(pptx.ShapeType.rect, { x, y, w: 0.09, h, fill: { color }, line: { type: "none" } });
  const runs = [{ text: title, options: { bold: true, color, fontFace: HEAD, fontSize: 14, breakLine: true, paraSpaceAfter: 5 } }];
  const body = typeof items === "string" ? [{ text: items, options: { color: INK_SOFT, fontFace: BODY, fontSize: fs || 12 } }] : bl(items, { fs: fs || 12, color: INK_SOFT, gap: 3 });
  s.addText(runs.concat(body), { x: x + 0.28, y: y + 0.12, w: w - 0.45, h: h - 0.2, isTextBox: true, margin: 0, valign: "top" });
}

// four stage boxes in a row with arrows; items: [title, sub, fill, line]
function flowRow(s, y, items, o = {}) {
  const n = items.length, w = o.w || 2.5, h = o.h || 1.0, gap = (12.35 - n * w) / (n - 1);
  items.forEach((it, i) => {
    const x = 0.5 + i * (w + gap);
    box(s, x, y, w, h, it[2] || WHITE, it[3] || INK_SOFT, it[4] || 1.5);
    s.addText([{ text: it[0], options: { fontFace: HEAD, bold: true, fontSize: o.tfs || 13, color: INK, breakLine: !!it[1] } }].concat(it[1] ? [{ text: it[1], options: { fontFace: BODY, fontSize: o.sfs || 10.5, color: INK_SOFT } }] : []), { x: x + 0.1, y, w: w - 0.2, h, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    if (i < n - 1) arrow(s, x + w + 0.05, y + h / 2, x + w + gap - 0.05, y + h / 2, o.arrowColor || INK_SOFT, { w: 1.5 });
  });
}



// ============================================================ DECK

// ---- title
N(`TIME: 2 minutes.
SAY: Welcome to Day 8, the last session. On Day 7 we saw that once data is spread over many services, you can no longer join it. Today we put it back together without slowing the application. We follow one question from the production database to a dashboard: what was our revenue by customer segment last quarter?
ASK: Who has been asked for "a quick number" from production? What happened?`,
`الوقت: دقيقتان.
قل: مرحبًا بكم في اليوم الثامن، الجلسة الأخيرة. رأينا في اليوم السابع أنه عندما تتوزع البيانات على خدمات كثيرة لا يمكن عمل join عليها. اليوم نجمعها من جديد دون إبطاء التطبيق. نتابع سؤالًا واحدًا من قاعدة بيانات الإنتاج إلى لوحة المؤشرات: كم كان إيرادنا حسب شريحة العملاء في الربع الأخير؟
اسأل: من طُلب منه "رقم سريع" من قاعدة الإنتاج؟ ماذا حدث؟`);
add((n) => {
  const s = newSlide(true);
  s.addText("SELECT segment, SUM(amount)\nFROM fact_orders\nGROUP BY segment", { x: 6.6, y: 0.5, w: 6.4, h: 2.0, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 18, color: DARK_WATERMARK, lineSpacingMultiple: 1.3 });
  s.addText("CLEAN CODE TO SCALE   ·   DAY 08 OF 08", { x: 0.7, y: 2.1, w: 10, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Data Engineering", { x: 0.7, y: 2.55, w: 11.8, h: 0.9, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 44, color: WHITE });
  s.addText("Warehouses, lakes and pipelines", { x: 0.7, y: 3.4, w: 11.8, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 26, color: "B9D3D6" });
  s.addText("How data from the application becomes answers people can trust, without slowing the application down.", { x: 0.7, y: 4.25, w: 10, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 17, color: "B9D3D6", valign: "top" });
  [["2:30", "DURATION"], ["SQL + PYTHON", "CODE SAMPLES"], ["PAIRS", "PIPELINE DESIGN + QUIZ"]].forEach((st, i) => {
    const x = 0.7 + i * 3.1;
    s.addText(st[0], { x, y: 5.65, w: 2.9, h: 0.45, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 19, color: WHITE });
    s.addText(st[1], { x, y: 6.13, w: 2.9, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 9.5, color: "6C8B90", charSpacing: 1 });
  });
  footer(s, n, true, "Clean Code to Scale");
});

// ---- agenda
N(`TIME: 2 minutes.
SAY: Six blocks. An opening question. Part 1: two kinds of work, and why analytics should not run on production. Part 2: warehouse, lake and lakehouse, and how to model data for questions. Part 3: how data moves: change data capture, ETL and ELT, and pipelines you can run twice. Part 4: how to trust the data. Then an exercise in pairs and a short review.
NOTE: Three check-ins and a quiz are woven in. If time is short, run them as a show of hands.`,
`الوقت: دقيقتان.
قل: ست فقرات. سؤال افتتاحي. الجزء الأول: نوعان من العمل، ولماذا لا يجب أن تعمل التحليلات على قاعدة الإنتاج. الجزء الثاني: مستودع البيانات والبحيرة والـ lakehouse، وكيف نصمم البيانات للأسئلة. الجزء الثالث: كيف تنتقل البيانات: التقاط التغييرات (CDC) وETL وELT ومسارات يمكن تشغيلها مرتين. الجزء الرابع: كيف نثق بالبيانات. ثم تمرين في أزواج ومراجعة قصيرة.
ملاحظة: هناك ثلاث وقفات واختبار موزعة. إذا ضاق الوقت فنفذها برفع الأيدي.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Session Map", "Four parts, an exercise, and a question to start", false);
  const steps = [
    ["01", "0:00–0:10", "OPENING", INK_FAINT, "One Monday question."],
    ["02", "0:10–0:35", "PART 1", TEAL, "Two kinds of work: OLTP and OLAP."],
    ["03", "0:35–1:00", "PART 2", SLATE, "Warehouse, lake, lakehouse, and the star schema."],
    ["04", "1:00–1:30", "PART 3", PLUM, "Moving the data: CDC, ETL and ELT, pipelines."],
    ["05", "1:30–2:00", "PART 4", RUST, "Trusting the data: checks, contracts, privacy."],
    ["06", "2:00–2:30", "EXERCISE + REVIEW", AMBER, "From checkout to dashboard, then group critique."],
  ];
  const startX = 0.75, gap = 2.1, cy = 2.4, d = 1.0;
  s.addShape(pptx.ShapeType.line, { x: startX + d / 2, y: cy + d / 2, w: gap * 5, h: 0, line: { color: LINE, width: 1.5 } });
  steps.forEach((st, i) => {
    const x = startX + i * gap, cx = x + d / 2;
    circle(s, x, cy, d, st[3], st[0], { fontSize: 19 });
    s.addText(st[1], { x: cx - 1.0, y: cy + d + 0.15, w: 2.0, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: CODE, bold: true, fontSize: 12.5, color: INK });
    s.addText(st[2], { x: cx - 1.0, y: cy + d + 0.47, w: 2.0, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: CODE, fontSize: 9.5, color: st[3], charSpacing: 1 });
    s.addText(st[4], { x: cx - 1.0, y: cy + d + 0.9, w: 2.0, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "top", fontFace: BODY, fontSize: 11.5, color: INK_SOFT });
  });
  s.addText("Three short check-ins, a quiz and an exercise in pairs are woven through the parts.", { x: 0.5, y: 6.0, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  s.addText("The streaming block (windows, watermarks, state) adds about 25 minutes. Use it in full, or hand it out. Short on time? Run the check-ins as a show of hands.", { x: 0.5, y: 6.4, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  footer(s, n, false);
});

// ---- objectives
N(`TIME: 1 minute.
SAY: Four things to leave with. Explain why analytics should not run on the production database, and how row and column storage differ. Tell a warehouse, a lake and a lakehouse apart, and read a star schema. Move data safely: ETL and ELT, batch and streaming, change data capture, and pipelines you can run again. And keep data trustworthy: quality checks, contracts, duplicates and late data, and personal data.`,
`الوقت: دقيقة واحدة.
قل: أربعة أشياء تخرجون بها. شرح لماذا لا يجب أن تعمل التحليلات على قاعدة الإنتاج، وكيف يختلف التخزين بالصفوف عن التخزين بالأعمدة. التمييز بين المستودع والبحيرة والـ lakehouse، وقراءة مخطط النجمة (star schema). نقل البيانات بأمان: ETL وELT، والدفعات والبث، والتقاط التغييرات، ومسارات يمكن إعادة تشغيلها. والحفاظ على موثوقية البيانات: فحوص الجودة والعقود والتكرار والبيانات المتأخرة والبيانات الشخصية.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Objectives", "What today is for", false);
  [
    ["Know why analytics should not run on production", "And how row storage and column storage differ."],
    ["Tell a warehouse, a lake and a lakehouse apart", "And read a star schema: facts and dimensions."],
    ["Move data safely", "ETL vs ELT, batch vs streaming, change data capture, and pipelines you can run again."],
    ["Keep data trustworthy", "Quality checks, contracts, duplicates and late data, and personal data."],
  ].forEach((it, i) => {
    const y = 1.75 + i * 1.3;
    circle(s, 0.7, y, 0.56, TEAL, "✓", { fontFace: BODY, fontSize: 20 });
    s.addText(it[0], { x: 1.6, y: y - 0.06, w: 10.8, h: 0.42, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: INK });
    s.addText(it[1], { x: 1.6, y: y + 0.38, w: 10.8, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  });
  footer(s, n, false);
});

// ---- room
N(`TIME: 1 minute.
SAY: Why this is for you. Junior and mid engineers: most of the data you create ends up in somebody's report, and you learn what can go wrong on the way. Tech leads: "can we get this number?" is a weekly request, and you learn the options and what each costs production. QEs: a pipeline needs tests, and the worst bug is a number that is quietly wrong. AI engineers: search indexes, evaluation sets and features are all pipelines, and an agent asked for the numbers will query production unless told otherwise.`,
`الوقت: دقيقة واحدة.
قل: لماذا هذا مهم لكم. المهندسون المبتدئون والمتوسطون: معظم البيانات التي تنتجونها تنتهي في تقرير أحدهم، وتتعلمون ما قد يسوء في الطريق. القادة التقنيون: "هل نستطيع الحصول على هذا الرقم؟" طلب أسبوعي، وتعرفون الخيارات وكلفة كل منها على الإنتاج. مهندسو الجودة: المسار يحتاج اختبارات، وأسوأ خطأ هو رقم خاطئ بصمت. مهندسو الذكاء الاصطناعي: فهارس البحث ومجموعات التقييم والخصائص كلها مسارات بيانات، والوكيل الذي يُطلب منه الأرقام سيستعلم من الإنتاج ما لم يُقل له غير ذلك.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Room", "Why this session, for you specifically", false);
  const roles = [
    ["JR", TEAL, "Junior – Mid Engineers", "Most of the data you create ends up in somebody's report. You learn where it goes, and what can go wrong on the way."],
    ["SR", SLATE, "Senior Eng & Tech Leads", "\"Can we get this number?\" is a weekly request. You learn the options and what each one costs production."],
    ["QE", RUST, "QEs", "Pipelines need tests too, and the worst bug is a number that is quietly wrong. You learn the checks that catch it."],
    ["AI", PLUM, "AI Engineers", "Search indexes, evaluation sets and features are pipelines. An agent asked for \"the numbers\" will query production unless told otherwise."],
  ];
  const cw = 5.85, ch = 2.1, gx = 0.3, gy = 0.25, ox = 0.5, oy = 1.85;
  roles.forEach((r, i) => {
    const x = ox + (i % 2) * (cw + gx), y = oy + Math.floor(i / 2) * (ch + gy);
    card(s, x, y, cw, ch, WHITE);
    circle(s, x + 0.3, y + 0.3, 0.72, r[1], r[0], { fontSize: 14 });
    s.addText(r[2], { x: x + 1.25, y: y + 0.32, w: cw - 1.5, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 15, color: INK });
    s.addText(r[3], { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: INK_SOFT, valign: "top" });
  });
  footer(s, n, false);
});

// ---- opening
N(`TIME: about 4 minutes.
SAY: Monday morning. A product manager asks: what was our revenue by customer segment last quarter? It sounds simple. An engineer runs a big query on the orders database at 10:00, and checkout gets twice as slow for twenty minutes. The number is wrong anyway: refunds live in the Payments database and segments live in the Customers service, and nothing joins them. And next week the same person asks for it by month and for two years back, but nobody kept the old values, because the database only stores the current state.
ASK: Take two or three stories of a quick number that went wrong. We will come back to them.`,
`الوقت: حوالي 4 دقائق.
قل: صباح الاثنين. مدير منتج يسأل: كم كان إيرادنا حسب شريحة العملاء في الربع الأخير؟ يبدو بسيطًا. مهندس يشغّل استعلامًا كبيرًا على قاعدة الطلبات الساعة العاشرة، فيصبح الدفع أبطأ بمرتين لعشرين دقيقة. والرقم خاطئ على أي حال: المبالغ المستردة في قاعدة المدفوعات والشرائح في خدمة العملاء، ولا شيء يربطها. والأسبوع القادم يسأل الشخص نفسه عنه حسب الشهر وعن سنتين للوراء، لكن لا أحد احتفظ بالقيم القديمة، لأن قاعدة البيانات تخزن الحالة الحالية فقط.
اسأل: خذوا قصتين أو ثلاثًا عن رقم سريع كان خاطئًا. سنعود إليها.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Opening", "One Monday question", false, { sub: "\"What was our revenue by customer segment last quarter?\" It sounds simple. Here is what happens:" });
  infoCard(s, 0.5, 1.85, 4.0, 3.3, RUST, "It slows checkout", "An engineer runs a big query on the orders database at 10:00. Checkout gets twice as slow for twenty minutes.", 13.5);
  infoCard(s, 4.67, 1.85, 4.0, 3.3, RUST, "The number is wrong", "Refunds live in the Payments database. Segments live in the Customers service. Nothing joins them (Day 7).", 13.5);
  infoCard(s, 8.84, 1.85, 4.0, 3.3, RUST, "There is no history", "Next week: \"by month, for two years\". The database stores the current state, so the old values are gone.", 13.5);
  callout(s, 0.5, 5.5, 12.35, 1.1, "The answer is not a bigger query.", "It is a different place to ask: a copy of the data, built for questions, kept away from checkout.", false);
  footer(s, n, false);
});

N(`TIME: 1 minute.
SAY: Three questions for today. Where should questions like this run, so that checkout never notices? How does the data get there, and how do we know it arrived once and in full? And how do we trust the number?`,
`الوقت: دقيقة واحدة.
قل: ثلاثة أسئلة لليوم. أين يجب أن تعمل أسئلة كهذه بحيث لا يلاحظ الدفع شيئًا؟ كيف تصل البيانات إلى هناك، وكيف نعرف أنها وصلت مرة واحدة وكاملة؟ وكيف نثق بالرقم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Opening", "Three questions for today", false);
  [
    ["1", TEAL, "Where should questions like this run?", "So that checkout never notices."],
    ["2", PLUM, "How does the data get there?", "And how do we know it arrived once, and in full?"],
    ["3", RUST, "How do we trust the number?", "Quiet failures are worse than loud ones."],
  ].forEach((it, i) => {
    const y = 1.95 + i * 1.6;
    card(s, 0.5, y, 12.35, 1.4, WHITE);
    circle(s, 0.8, y + 0.3, 0.8, it[1], it[0], { fontSize: 24 });
    s.addText(it[2], { x: 1.9, y: y + 0.22, w: 10.7, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 20, color: INK });
    s.addText(it[3], { x: 1.9, y: y + 0.78, w: 10.7, h: 0.45, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK_SOFT, valign: "top" });
  });
  footer(s, n, false);
});

// ============================================================ PART 1
N(`TIME: 30 seconds.
SAY: Part 1. Two kinds of work.`,
`الوقت: 30 ثانية.
قل: الجزء الأول. نوعان من العمل.`);
divider(1, "0:10–0:35", "Two kinds of work", "Running the business and understanding the business are different jobs.");

N(`TIME: about 4 minutes.
SAY: Databases do two very different jobs. OLTP, online transaction processing, runs the business: take an order, update a stock count. OLAP, online analytical processing, helps you understand the business: revenue by segment, returns by product.
SAY: Read the rows. A typical OLTP query gets or changes a few rows by key, in milliseconds, with thousands at once, on row-by-row storage, keeping only the current state. A typical OLAP query scans millions of rows and adds them up, in seconds or minutes, a few at a time, on column-by-column storage, and keeps all the history.
ASK: Which of the two is the database you work with most?`,
`الوقت: حوالي 4 دقائق.
قل: قواعد البيانات تقوم بعملين مختلفين جدًا. OLTP، معالجة المعاملات الفورية، يدير العمل: استلام طلب، تحديث رصيد مخزون. وOLAP، المعالجة التحليلية، يساعد على فهم العمل: الإيراد حسب الشريحة، المرتجعات حسب المنتج.
قل: اقرأوا الصفوف. استعلام OLTP النموذجي يجلب أو يغيّر صفوفًا قليلة بالمفتاح، في أجزاء من الثانية، بآلاف في وقت واحد، على تخزين صفًا بصف، ويحتفظ بالحالة الحالية فقط. واستعلام OLAP النموذجي يمسح ملايين الصفوف ويجمعها، في ثوانٍ أو دقائق، بعدد قليل في كل مرة، على تخزين عمودًا بعمود، ويحتفظ بكل التاريخ.
اسأل: أي القاعدتين هي التي تعملون بها أكثر؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Two kinds of work", "OLTP and OLAP", false, { sub: "OLTP runs the business. OLAP helps you understand it." });
  tbl(s, 0.5, 1.85, [2.3, 5.0, 5.05], [
    ["", "OLTP (the application)", "OLAP (analytics)"],
    ["Purpose", "Run the business", "Understand the business"],
    ["A typical query", "Get, insert or update a few rows by key", "Scan millions of rows and add them up"],
    ["Speed needed", "Milliseconds, with thousands at once", "Seconds to minutes, with a few at once"],
    ["Storage shape", "Row by row", "Column by column"],
    ["History", "The current state. Old values are overwritten", "All the history is kept"],
    ["Used by", "Customers and services", "Analysts, dashboards, ML"],
    ["Examples", "Postgres, MySQL", "BigQuery, Snowflake, Redshift, ClickHouse"],
  ], { fs: 12.5, rowH: [0.42, 0.55, 0.65, 0.65, 0.55, 0.65, 0.55, 0.55] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Why not just run the report on production? Five reasons. It competes with checkout for CPU, memory and disk: the shared resource from Day 4. Long queries cause side effects: they slow cleanup in the database, and on a replica they can make it fall behind. The schema is shaped for transactions, split into many small tables, and the useful answer needs many joins. The history is gone: a customer who moved country makes all their old orders show the new country. And after Day 7 no single database has the answer.
SAY: A read replica is a start, not an answer. It protects checkout from the load, but it does not fix the schema, the missing history, or the data in other services.`,
`الوقت: حوالي 3 دقائق.
قل: لماذا لا نشغّل التقرير على الإنتاج؟ خمسة أسباب. ينافس الدفع على المعالج والذاكرة والقرص: المورد المشترك من اليوم الرابع. والاستعلامات الطويلة لها آثار جانبية: تبطئ عمليات التنظيف في القاعدة، وعلى النسخة قد تجعلها تتأخر. والمخطط مصمم للمعاملات، مقسّم إلى جداول صغيرة كثيرة، والإجابة المفيدة تحتاج joins كثيرة. والتاريخ ضاع: عميل انتقل إلى بلد آخر تجعل كل طلباته القديمة تظهر بالبلد الجديد. وبعد اليوم السابع لا توجد قاعدة واحدة فيها الإجابة.
قل: نسخة القراءة (read replica) بداية وليست جوابًا. تحمي الدفع من الحمل لكنها لا تصلح المخطط ولا التاريخ المفقود ولا البيانات في خدمات أخرى.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Two kinds of work", "Why not just run the report on production?", false);
  [
    [RUST, "It competes with checkout", "A big scan uses the CPU, memory and disk that customers need (Day 4)."],
    [RUST, "Long queries have side effects", "They slow cleanup in the database, and on a replica they can make it fall behind."],
    [RUST, "The schema is for transactions", "Many small tables. The useful answer needs many joins."],
    [RUST, "The history is gone", "A customer who moved country makes all their old orders show the new country."],
    [RUST, "No database has the answer", "After the split on Day 7 the data lives in many services."],
  ].forEach(([c, t, d], i) => {
    const y = 1.8 + i * 0.84;
    card(s, 0.5, y, 12.35, 0.72, WHITE);
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 0.09, h: 0.72, fill: { color: c }, line: { type: "none" } });
    s.addText([{ text: t + "   ", options: { bold: true, color: INK, fontFace: HEAD, fontSize: 14 } }, { text: d, options: { color: INK_SOFT, fontFace: BODY, fontSize: 12.5 } }], { x: 0.8, y, w: 11.9, h: 0.72, isTextBox: true, margin: 0, valign: "middle" });
  });
  callout(s, 0.5, 6.05, 12.35, 0.8, "A read replica is a start, not an answer.", "It protects checkout from the load, but it does not fix the schema, the missing history, or the data that lives in other services.", false);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Row storage and column storage. An order database stores each order as a row, so reading one whole order is fast. On the left, looking up order 2 touches one row. An analytical question usually needs one or two columns from millions of rows. On the right, adding up the amounts touches one column and skips the other three. A column store keeps each column together, so it reads only what the question needs, and it compresses well because similar values sit next to each other. Parquet is the common file format built this way.
ASK: If the table had 100 columns and the question needs 2, how much less does the column store read?`,
`الوقت: حوالي 3 دقائق.
قل: التخزين بالصفوف والتخزين بالأعمدة. قاعدة الطلبات تخزن كل طلب كصف، فقراءة طلب كامل سريعة. على اليسار، البحث عن الطلب 2 يلمس صفًا واحدًا. أما السؤال التحليلي فعادة يحتاج عمودًا أو اثنين من ملايين الصفوف. على اليمين، جمع المبالغ يلمس عمودًا واحدًا ويتجاوز الثلاثة الأخرى. مخزن الأعمدة يحتفظ بكل عمود معًا، فيقرأ ما يحتاجه السؤال فقط، ويضغط جيدًا لأن القيم المتشابهة متجاورة. وParquet هو صيغة الملفات الشائعة المبنية بهذه الطريقة.
اسأل: لو كان الجدول بمئة عمود والسؤال يحتاج عمودين، كم أقل يقرأ مخزن الأعمدة؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Two kinds of work", "Row storage and column storage", false, { sub: "The same four orders, stored two ways." });
  const cols = ["id", "customer", "country", "amount"];
  const data = [["1", "Ana", "EG", "40"], ["2", "Sam", "IE", "25"], ["3", "Lee", "US", "90"], ["4", "Mia", "EG", "15"]];
  const grid = (x0, hiRow, hiCol, color, tint) => {
    cols.forEach((c, j) => label(s, c, x0 + j * 1.3, 2.45, 1.3, 0.3, { size: 9.5, color: INK_FAINT }));
    data.forEach((r, i) => r.forEach((v, j) => {
      const hi = (hiRow === i) || (hiCol === j);
      s.addShape(pptx.ShapeType.rect, { x: x0 + j * 1.3, y: 2.8 + i * 0.55, w: 1.3, h: 0.55, fill: { color: hi ? tint : WHITE }, line: { color: hi ? color : LINE, width: hi ? 1.5 : 1 } });
      label(s, v, x0 + j * 1.3, 2.8 + i * 0.55, 1.3, 0.55, { size: 11 });
    }));
  };
  label(s, "ROW STORE: read one order", 0.7, 1.95, 5.2, 0.3, { size: 11, bold: true, color: TEAL });
  grid(0.7, 1, -1, TEAL, TEAL_TINT);
  label(s, "stored one row after another", 0.7, 5.15, 5.2, 0.3, { size: 10.5, color: INK_SOFT });
  label(s, "good for: get order 2", 0.7, 5.5, 5.2, 0.35, { size: 12, bold: true, color: TEAL });
  label(s, "COLUMN STORE: add up the amounts", 7.1, 1.95, 5.2, 0.3, { size: 11, bold: true, color: PLUM });
  grid(7.1, -1, 3, AMBER, AMBER_TINT);
  label(s, "stored one column after another", 7.1, 5.15, 5.2, 0.3, { size: 10.5, color: INK_SOFT });
  label(s, "good for: SUM(amount)", 7.1, 5.5, 5.2, 0.35, { size: 12, bold: true, color: AMBER });
  s.addText("A column store reads only the columns the question needs, and compresses well because similar values sit together. Parquet is the common file format built this way.", { x: 0.7, y: 6.1, w: 11.9, h: 0.7, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Give pairs two minutes. For each of the six queries they say OLTP or OLAP. Take a show of hands, then reveal.
SAY: A good rule: if a person is waiting for one answer about one thing, it is OLTP. If you are asking about many things at once, it is OLAP.`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: أعطِ الأزواج دقيقتين. لكل من الاستعلامات الستة يقولون OLTP أم OLAP. خذ رفع الأيدي ثم اكشف الإجابات.
قل: قاعدة جيدة: إذا كان شخص ينتظر جوابًا واحدًا عن شيء واحد فهو OLTP. وإذا كنتم تسألون عن أشياء كثيرة معًا فهو OLAP.`);
checkIn(1, 3, "OLTP or OLAP?", "For each query, is it the application's kind of work (OLTP) or analytics (OLAP)?", [
  "Show customer 42 their order history page.",
  "Total revenue by country last quarter.",
  "Record a new payment.",
  "Average basket size by week, for two years.",
  "Is seat 14C already taken?",
  "The ten products with the most returns this year.",
], null, "a. OLTP. A few rows found by key, and it must be fast.\nb. OLAP. A scan over a whole quarter, then a sum.\nc. OLTP. A single write that must be safe.\nd. OLAP. Two years of history, grouped by week.\ne. OLTP. One row, and the answer decides a sale.\nf. OLAP. A scan and a ranking over a whole year.\n\nA good rule: if a person is waiting for one answer about one thing, it is OLTP. If you are asking about many things at once, it is OLAP.");

// ============================================================ PART 2
N(`TIME: 30 seconds.
SAY: Part 2. Where should the analytical copy of the data live?`,
`الوقت: 30 ثانية.
قل: الجزء الثاني. أين يجب أن تعيش النسخة التحليلية من البيانات؟`);
divider(2, "0:35–1:00", "Warehouse, lake, lakehouse", "Three places for the analytical copy, the layers inside them, and how to shape data for questions.");

N(`TIME: about 4 minutes.
SAY: Three common answers, three points on a line from tidy and ready to query, to cheap and keep everything. A data warehouse holds cleaned, modelled tables. The schema is decided before loading, schema on write. It is good for dashboards, finance and SQL for everyone. Watch out: new messy sources need work first. BigQuery, Snowflake, Redshift.
SAY: A data lake holds files of any kind, JSON, CSV, logs, images, Parquet, in cheap object storage like S3. The schema is applied when you read, schema on read. It is good for raw history and machine learning. Watch out for the data swamp: nobody knows what is in it or which copy is right.
SAY: A lakehouse is lake files organised as tables by a table format: Delta Lake, Iceberg or Hudi. You get tables, updates and deletes on cheap storage, one copy for reports and ML. It is newer, with more moving parts.
ASK: Which of these does your company have today, if any?`,
`الوقت: حوالي 4 دقائق.
قل: ثلاث إجابات شائعة، ثلاث نقاط على خط يبدأ من "مرتب وجاهز للاستعلام" وينتهي بـ"رخيص واحتفظ بكل شيء". مستودع البيانات (data warehouse) يحتفظ بجداول نظيفة ومصممة. المخطط يُحدد قبل التحميل (schema on write). مناسب للوحات المؤشرات والمالية وSQL للجميع. انتبهوا: المصادر الجديدة الفوضوية تحتاج عملًا أولًا. BigQuery وSnowflake وRedshift.
قل: بحيرة البيانات (data lake) تحتفظ بملفات من أي نوع، JSON وCSV وسجلات وصور وParquet، في تخزين كائنات رخيص مثل S3. المخطط يُطبّق عند القراءة (schema on read). مناسبة للتاريخ الخام ولتعلم الآلة. انتبهوا لمستنقع البيانات: لا أحد يعرف ما فيه ولا أي نسخة هي الصحيحة.
قل: الـ lakehouse ملفات بحيرة منظمة كجداول بواسطة صيغة جداول: Delta Lake أو Iceberg أو Hudi. تحصلون على جداول وتحديثات وحذف على تخزين رخيص، ونسخة واحدة للتقارير وتعلم الآلة. وهو أحدث وفيه أجزاء متحركة أكثر.
اسأل: أي هذه الثلاثة عند شركتكم اليوم، إن وجد؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Where it lives", "Warehouse, lake and lakehouse", false, { sub: "Three points on a line, from \"tidy and ready to query\" to \"cheap, keep everything\"." });
  tbl(s, 0.5, 1.85, [1.9, 3.5, 3.5, 3.45], [
    ["", "Data warehouse", "Data lake", "Lakehouse"],
    ["What it holds", "Cleaned, modelled tables", "Files of any kind: JSON, CSV, logs, images, Parquet", "Lake files, organised as tables"],
    ["Schema", "Decided before loading (schema on write)", "Applied when you read (schema on read)", "A table format adds a schema to the files"],
    ["Storage", "Managed by the warehouse", "Cheap object storage (S3, GCS, Azure)", "Object storage, plus a table format (Delta Lake, Iceberg, Hudi)"],
    ["Good for", "Dashboards, finance, SQL for everyone", "Raw history, ML, data with no fixed shape", "One copy for reports and ML, with updates and deletes"],
    ["Watch out for", "New messy sources need work first", "A \"data swamp\": nobody knows what is in it", "Newer, with more moving parts"],
    ["Examples", "BigQuery, Snowflake, Redshift", "Files on S3 read by Spark or Athena", "Databricks, or Iceberg on S3"],
  ], { fs: 11.5, rowH: [0.42, 0.65, 0.8, 0.95, 0.8, 0.8, 0.65] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Whichever you pick, good teams do not load data straight into the final tables. They keep layers. Sources feed a raw layer, the data as received, append only, often called bronze. Then a cleaned layer: deduplicated, typed, fixed, called silver. Then a modelled layer of business tables, called gold. Then consumers: dashboards, ML, APIs.
SAY: The point is on the bottom line: found a bug in the cleaning step? Fix it, and re-run from raw. Nothing is lost. If you load straight into the final tables, you cannot do that.
ASK: Does your pipeline keep the raw data, or only the cleaned result?`,
`الوقت: حوالي 3 دقائق.
قل: أيًا كان اختياركم، الفرق الجيدة لا تحمّل البيانات مباشرة إلى الجداول النهائية. تحتفظ بطبقات. المصادر تغذي طبقة خام (raw)، البيانات كما وصلت، إضافة فقط، وتسمى غالبًا برونزية. ثم طبقة منظفة: بلا تكرار ومحددة الأنواع ومصححة، وتسمى فضية. ثم طبقة مصممة من جداول الأعمال، وتسمى ذهبية. ثم المستهلكون: لوحات المؤشرات وتعلم الآلة وواجهات البرمجة.
قل: النقطة المهمة في السطر السفلي: وجدتم خطأ في خطوة التنظيف؟ أصلحوه وأعيدوا التشغيل من الخام. لا شيء يضيع. أما إذا حمّلتم مباشرة إلى الجداول النهائية فلا تستطيعون ذلك.
اسأل: هل يحتفظ مسار بياناتكم بالبيانات الخام، أم بالنتيجة المنظفة فقط؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Layers", "Raw, cleaned, modelled", false, { sub: "Do not load straight into the final tables. Keep layers." });
  flowRow(s, 2.1, [
    ["Sources", "services, files, APIs", WHITE, INK_SOFT, 1.5],
    ["Raw (bronze)", "as received, append only", AMBER_TINT, AMBER, 1.5],
    ["Cleaned (silver)", "deduplicated, typed, fixed", TEAL_TINT, TEAL, 1.5],
    ["Modelled (gold)", "facts and dimensions", PLUM_TINT, PLUM, 2.5],
    ["Consumers", "dashboards, ML, APIs", WHITE, INK_SOFT, 1.5],
  ], { w: 2.1, h: 1.4, tfs: 13, sfs: 10.5 });
  card(s, 0.5, 4.2, 12.35, 1.3, WHITE, { shadow: false, line: LINE });
  s.addText("Found a bug in the cleaning step? Fix it, and re-run from raw. Nothing is lost.", { x: 0.5, y: 4.2, w: 12.35, h: 1.3, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: HEAD, bold: true, fontSize: 20, color: TEAL });
  s.addText("Bronze, silver and gold are a common naming. The idea matters more than the names.", { x: 0.5, y: 5.85, w: 12.3, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT });
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: The top layer is shaped for questions. The classic shape is a star. In the middle is a fact table: one row per event, with numbers you add up, like amount and quantity. Around it are dimension tables: the things you group and filter by, customer, product, date, region.
SAY: The fact table is long and narrow, a row for every order. The dimensions are short and wide, a row for every customer, product, day or region. Every analytical question joins the fact table to the dimensions it needs.`,
`الوقت: حوالي 4 دقائق.
قل: الطبقة العليا مصممة للأسئلة. الشكل الكلاسيكي نجمة. في الوسط جدول حقائق (fact table): صف لكل حدث، فيه أرقام تجمعونها مثل المبلغ والكمية. حوله جداول أبعاد (dimension tables): الأشياء التي تجمّعون وتصفّون حسبها: العميل والمنتج والتاريخ والمنطقة.
قل: جدول الحقائق طويل وضيق، صف لكل طلب. والأبعاد قصيرة وعريضة، صف لكل عميل أو منتج أو يوم أو منطقة. كل سؤال تحليلي يربط جدول الحقائق بالأبعاد التي يحتاجها.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Modelling", "The star schema", false, { sub: "One fact table in the middle. Dimension tables around it." });
  const cx = 4.9, cy = 3.35, cw = 3.5, ch = 1.9;
  const dims = [
    [0.5, 1.9, "dim_customer", "customer_key\nname, segment, country"],
    [9.35, 1.9, "dim_product", "product_key\nname, category"],
    [0.5, 4.8, "dim_date", "date_key\nday, month, quarter, year"],
    [9.35, 4.8, "dim_region", "region_key\ncountry, region"],
  ];
  dims.forEach(([x, y, t, d]) => {
    arrow(s, x < 5 ? x + 3.0 : x, y + 0.6, x < 5 ? cx : cx + cw, y < 3 ? cy + 0.4 : cy + ch - 0.4, INK_FAINT, { w: 1.5, noHead: true });
    box(s, x, y, 3.0, 1.2, TEAL_TINT, TEAL, 1.5);
    s.addText([{ text: t, options: { fontFace: CODE, bold: true, fontSize: 12, color: INK, breakLine: true, paraSpaceAfter: 3 } }, { text: d, options: { fontFace: CODE, fontSize: 9.5, color: INK_SOFT } }], { x: x + 0.1, y, w: 2.8, h: 1.2, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  });
  box(s, cx, cy, cw, ch, PLUM_TINT, PLUM, 2.5);
  s.addText([{ text: "fact_orders", options: { fontFace: CODE, bold: true, fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 4 } }, { text: "order_id", options: { fontFace: CODE, fontSize: 10, color: INK_SOFT, breakLine: true } }, { text: "customer_key, product_key,", options: { fontFace: CODE, fontSize: 10, color: INK_SOFT, breakLine: true } }, { text: "date_key, region_key", options: { fontFace: CODE, fontSize: 10, color: INK_SOFT, breakLine: true, paraSpaceAfter: 4 } }, { text: "quantity, amount", options: { fontFace: CODE, bold: true, fontSize: 11, color: PLUM } }], { x: cx + 0.1, y: cy, w: cw - 0.2, h: ch, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  s.addText("The fact table is long and narrow: a row for every order. The dimensions are short and wide: a row for every customer, product, day or region.", { x: 0.5, y: 6.35, w: 12.35, h: 0.55, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT });
  footer(s, n, false);
});

// ---- snowflake schema
N(`TIME: about 3 minutes.
SAY: A variation of the star: the snowflake schema. In a star, each dimension is one wide table. In a snowflake, the dimensions are normalised, split into smaller linked tables. The customer table no longer holds the country and region. It holds a geography key, and a geography table holds them. The product table holds a category key, and a category table holds the category name.
SAY: The gain: less repeated data, and a change is made in one place. Rename a category once, not on ten thousand product rows. The price: more joins, so queries are longer and a little slower, and it is harder for analysts to read. Most teams start with a star, because the simple shape is easier to use. Reach for a snowflake when a dimension is very large, or has a hierarchy that changes often.
NOTE: Not related to the Snowflake company. The name comes from the shape of the diagram.
ASK: Which would your analysts prefer: one extra join they never think about, or a wider table with repeated values?`,
`الوقت: حوالي 3 دقائق.
قل: نسخة مختلفة من النجمة: مخطط ندفة الثلج (snowflake schema). في النجمة كل بُعد جدول واحد عريض. وفي ندفة الثلج تُطبَّع الأبعاد، أي تُقسَّم إلى جداول أصغر مرتبطة. جدول العميل لم يعد يحمل البلد والمنطقة. صار يحمل مفتاح الجغرافيا، وجدول جغرافيا يحملهما. وجدول المنتج يحمل مفتاح الفئة، وجدول فئات يحمل اسم الفئة.
قل: المكسب: بيانات أقل تكرارًا، والتغيير يتم في مكان واحد. أعيدوا تسمية فئة مرة واحدة لا على عشرة آلاف صف منتج. الثمن: joins أكثر، فالاستعلامات أطول وأبطأ قليلًا، ويصعب على المحللين قراءتها. أغلب الفرق تبدأ بالنجمة لأن الشكل البسيط أسهل استخدامًا. استخدموا ندفة الثلج عندما يكون البُعد كبيرًا جدًا أو فيه تسلسل هرمي يتغير كثيرًا.
ملاحظة: لا علاقة له بشركة Snowflake. الاسم من شكل الرسم.
اسأل: أيهما يفضّل محللوكم: join إضافي لا يفكرون فيه، أم جدول أعرض بقيم مكررة؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Modelling", "The snowflake schema", false, { sub: "The same idea as the star, but each dimension is split into smaller, linked tables." });
  const tbox = (x, y, w, h, title, cols, f, c, lw) => {
    box(s, x, y, w, h, f, c, lw || 1.5);
    s.addText([{ text: title, options: { fontFace: CODE, bold: true, fontSize: 11, color: INK, breakLine: true, paraSpaceAfter: 3 } }, { text: cols, options: { fontFace: CODE, fontSize: 8.5, color: INK_SOFT } }], { x: x + 0.08, y, w: w - 0.16, h, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  };
  const Y = 2.0, H = 1.35;
  arrow(s, 2.6, Y + H / 2, 2.95, Y + H / 2, INK_FAINT, { w: 1.6, noHead: true });
  arrow(s, 5.05, Y + H / 2, 5.4, Y + H / 2, INK_FAINT, { w: 1.6, noHead: true });
  arrow(s, 7.9, Y + H / 2, 8.25, Y + H / 2, INK_FAINT, { w: 1.6, noHead: true });
  arrow(s, 10.35, Y + H / 2, 10.7, Y + H / 2, INK_FAINT, { w: 1.6, noHead: true });
  arrow(s, 6.65, Y + H, 6.65, 3.95, INK_FAINT, { w: 1.6, noHead: true });
  tbox(0.5, Y, 2.1, H, "dim_geography", "geography_key\ncountry, region", AMBER_TINT, AMBER);
  tbox(2.95, Y, 2.1, H, "dim_customer", "customer_key\nname, segment,\ngeography_key", TEAL_TINT, TEAL);
  tbox(5.4, Y, 2.5, H, "fact_orders", "order_id\ncustomer_key, product_key,\ndate_key\nquantity, amount", PLUM_TINT, PLUM, 2.5);
  tbox(8.25, Y, 2.1, H, "dim_product", "product_key\nname, category_key", TEAL_TINT, TEAL);
  tbox(10.7, Y, 2.15, H, "dim_category", "category_key\nname, department", AMBER_TINT, AMBER);
  tbox(5.4, 3.95, 2.5, 0.85, "dim_date", "date_key, day, month, quarter, year", TEAL_TINT, TEAL);
  label(s, "split out of dim_customer", 0.5, 3.45, 2.1, 0.3, { size: 8, color: AMBER, italic: true });
  label(s, "split out of dim_product", 10.7, 3.45, 2.15, 0.3, { size: 8, color: AMBER, italic: true });
  infoCard(s, 0.5, 5.05, 6.05, 1.2, TEAL, "Star", "Fewer joins, easier to read. Some values are repeated, such as the category name on every product row.", 11.5);
  infoCard(s, 6.8, 5.05, 6.05, 1.2, AMBER, "Snowflake", "Less repeated data, and a change in one place. More joins, so longer, slightly slower queries.", 11.5);
  s.addText("Most teams start with a star. Use a snowflake when a dimension is very large, or has a hierarchy that changes often. Not related to the Snowflake company: the name comes from the shape.", { x: 0.5, y: 6.4, w: 12.35, h: 0.55, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 12, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 2 minutes.
SAY: And now our Monday question is a simple join and a sum. Revenue by segment, last quarter: sum the amount from the fact table, join the customer dimension to get the segment, join the date dimension to filter the quarter, group by segment. One join for each thing you group or filter by.
ASK: How many lines of SQL would this question take on the normalised production schema?`,
`الوقت: حوالي دقيقتين.
قل: والآن سؤال الاثنين صار join بسيطًا وجمعًا. الإيراد حسب الشريحة في الربع الأخير: اجمعوا المبلغ من جدول الحقائق، واربطوا بُعد العميل لتحصلوا على الشريحة، واربطوا بُعد التاريخ لتصفّوا الربع، وجمّعوا حسب الشريحة. join واحد لكل شيء تجمّعون أو تصفّون به.
اسأل: كم سطر SQL كان سيحتاج هذا السؤال على مخطط الإنتاج المطبّع؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Modelling", "Our Monday question, as a query", false, { sub: "A join for each thing you group or filter by, and a sum." });
  codeBlock(s, 0.5, 1.85, 12.35, 3.2, "revenue by segment, last quarter", TEAL, [
    "SELECT c.segment, SUM(f.amount) AS revenue",
    "FROM fact_orders f",
    "JOIN dim_customer c ON c.customer_key = f.customer_key",
    "JOIN dim_date d     ON d.date_key     = f.date_key",
    "WHERE d.quarter = '2026-Q3'",
    "GROUP BY c.segment",
  ], 14);
  infoCard(s, 0.5, 5.3, 6.05, 1.5, TEAL, "Group by", "Dimensions: the things you group or filter by.", 13);
  infoCard(s, 6.8, 5.3, 6.05, 1.5, PLUM, "Add up", "Facts: the numbers you sum, like amount and quantity.", 13);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Remember the customer who moved country, whose old orders suddenly show the new one? The fix is to keep both versions of the customer row, each with the dates it was true. This is often called a type 2 slowly changing dimension.
SAY: Customer 42 was standard from February 2024 to May 2026, key 101. Then premium from June 2026, key 874. An order placed in April points at key 101, so it still counts as standard. An order in July points at 874. Revenue by segment now matches what was true at the time.`,
`الوقت: حوالي 3 دقائق.
قل: تذكروا العميل الذي انتقل إلى بلد آخر وصارت طلباته القديمة تظهر بالبلد الجديد؟ الحل أن نحتفظ بالنسختين من صف العميل، لكل منهما التواريخ التي كانت صحيحة فيها. يسمى هذا غالبًا بُعدًا متغيرًا ببطء من النوع الثاني (type 2 SCD).
قل: العميل 42 كان "عاديًا" من فبراير 2024 إلى مايو 2026، بالمفتاح 101. ثم "مميزًا" من يونيو 2026، بالمفتاح 874. طلب في أبريل يشير إلى المفتاح 101 فيُحسب "عاديًا". وطلب في يوليو يشير إلى 874. فصار الإيراد حسب الشريحة يطابق ما كان صحيحًا وقتها.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Modelling", "Keeping the history right", false, { sub: "Keep both versions of the row, each with the dates it was true." });
  tbl(s, 0.5, 1.85, [2.2, 2.2, 2.3, 2.8, 2.85], [
    ["customer_key", "customer_id", "segment", "valid_from", "valid_to"],
    ["101", "42", "standard", "2024-02-01", "2026-05-31"],
    ["874", "42", "premium", "2026-06-01", "(still true)"],
  ], { fs: 14, rowH: [0.5, 0.65, 0.65] });
  infoCard(s, 0.5, 4.0, 6.05, 2.0, TEAL, "An order in April", "Points at key 101. It still counts as standard, because that is what was true then.", 13.5);
  infoCard(s, 6.8, 4.0, 6.05, 2.0, PLUM, "An order in July", "Points at key 874. It counts as premium. Revenue by segment now matches what was true at the time.", 13.5);
  footer(s, n, false);
});

N(`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Pairs decide fact or dimension for five fields. Then reveal.
SAY: The test: would it make sense to sum it? Then it is a fact. Would you group by it or filter on it? Then it is a dimension.`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: تقرر الأزواج حقيقة أم بُعد لخمسة حقول. ثم اكشف الإجابات.
قل: الاختبار: هل يكون لجمعه معنى؟ إذن هو حقيقة. هل تجمّعون به أو تصفّون به؟ إذن هو بُعد.`);
checkIn(2, 3, "Fact or dimension?", "You are building the star schema for orders. For each field, does it belong in the fact table (a number you add up) or in a dimension (something you group or filter by)?", [
  "The order amount.",
  "The customer's segment.",
  "The quantity ordered.",
  "The product category.",
  "The calendar quarter.",
], null, "a. Fact. A number you add up.\nb. Dimension (customer). You group by it.\nc. Fact. A number you add up.\nd. Dimension (product). You group by it.\ne. Dimension (date). It is a label, not an amount.\n\nA good test: would it make sense to SUM it? Then it is a fact. Would you GROUP BY it or filter on it? Then it is a dimension.");

// ============================================================ PART 3
N(`TIME: 30 seconds.
SAY: Part 3. The data starts in the production database. How does it get to the warehouse without hurting the application, and without losing or repeating anything?`,
`الوقت: 30 ثانية.
قل: الجزء الثالث. البيانات تبدأ في قاعدة الإنتاج. كيف تصل إلى المستودع دون أن تؤذي التطبيق ودون أن نفقد شيئًا أو نكرره؟`);
divider(3, "1:00–1:30", "Moving the data", "Four ways to copy, change data capture, ETL and ELT, and pipelines you can run twice.");

// ---- anatomy of a pipeline
N(`TIME: about 3 minutes.
SAY: Before the details, the anatomy of a data pipeline. Four stages in a row. Ingest: bring the data in, from change data capture, files, APIs or events. Store: land it, raw, in the lake or warehouse. Transform: clean, join and model it. Serve: give it to dashboards, ML models and APIs, and sometimes back into the application, which is called reverse ETL.
SAY: And four things around every pipeline, which people forget: orchestration to run it, quality checks to trust it, lineage to know where a number came from, and access control for who may see what. A pipeline without these four is a script that happens to run.
ASK: Which of the four stages and four extras does your team have?`,
`الوقت: حوالي 3 دقائق.
قل: قبل التفاصيل، تشريح مسار البيانات. أربع مراحل في صف. الاستيعاب (ingest): إدخال البيانات من التقاط التغييرات أو الملفات أو واجهات البرمجة أو الأحداث. التخزين (store): إنزالها خامًا في البحيرة أو المستودع. التحويل (transform): تنظيفها وربطها وتصميمها. التقديم (serve): إعطاؤها للوحات المؤشرات ونماذج تعلم الآلة وواجهات البرمجة، وأحيانًا إعادتها إلى التطبيق، وهذا يسمى reverse ETL.
قل: وأربعة أشياء حول كل مسار ينساها الناس: التنسيق (orchestration) لتشغيله، وفحوص الجودة للوثوق به، وتتبع النسب (lineage) لمعرفة من أين جاء الرقم، والتحكم في الوصول لمن يرى ماذا. مسار بلا هذه الأربعة سكربت يصادف أنه يعمل.
اسأل: أي المراحل الأربع والإضافات الأربع عند فريقكم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Pipelines", "The anatomy of a data pipeline", false, { sub: "Four stages in a row, and four things around them that people forget." });
  flowRow(s, 1.95, [
    ["Ingest", "CDC, files, APIs, events", AMBER_TINT, AMBER, 1.5],
    ["Store", "raw, in the lake or warehouse", TEAL_TINT, TEAL, 1.5],
    ["Transform", "clean, join, model", TEAL_TINT, TEAL, 1.5],
    ["Serve", "dashboards, ML, APIs, reverse ETL", PLUM_TINT, PLUM, 2.5],
  ], { w: 2.6, h: 1.3, tfs: 15, sfs: 11 });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 3.65, w: 12.35, h: 3.1, rectRadius: 0.1, fill: { type: "none" }, line: { color: INK_FAINT, width: 1.25, dashType: "dash" } });
  label(s, "AROUND EVERY PIPELINE", 0.7, 3.72, 6, 0.3, { size: 10, bold: true, color: INK_SOFT, align: "left" });
  [
    [TEAL, "Orchestration", "Runs the steps in order, retries, backfills."],
    [PLUM, "Quality checks", "Freshness, volume, uniqueness. Stop a bad load."],
    [AMBER, "Lineage", "Where did this number come from, and what does a change break?"],
    [RUST, "Access control", "Who may see what, especially personal data."],
  ].forEach(([c, t, d], i) => infoCard(s, 0.7 + i * 3.0, 4.15, 2.85, 2.4, c, t, d, 12));
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: Four common ways. A nightly full copy: select star from every table, every night. It is heavy on production, it misses deleted rows, and there is nothing between two nights. Incremental by timestamp: copy rows where updated_at is newer than the last run. Lighter, but it misses deletes, and it misses rows whose updated_at was not kept up to date.
SAY: Change data capture, CDC: read the database's own change log and send every insert, update and delete, in order. You need a tool, and you must handle the first full copy and schema changes. And events from the app: the service publishes business events, the outbox from Day 7. Best meaning, but the service team must build and maintain them.
ASK: Which of these does your team use? What has gone wrong with it?`,
`الوقت: حوالي 4 دقائق.
قل: أربع طرق شائعة. نسخة كاملة كل ليلة: select * من كل جدول كل ليلة. ثقيلة على الإنتاج، وتفوّتها الصفوف المحذوفة، ولا شيء بين ليلتين. نسخ تزايدي حسب الوقت: انسخوا الصفوف التي updated_at فيها أحدث من آخر تشغيل. أخف، لكنه يفوّت الحذف، ويفوّت الصفوف التي لم يُحدَّث updated_at فيها.
قل: التقاط تغييرات البيانات (CDC): اقرأوا سجل التغييرات الخاص بالقاعدة وأرسلوا كل إدخال وتحديث وحذف بالترتيب. تحتاجون أداة، ويجب التعامل مع النسخة الكاملة الأولى وتغييرات المخطط. وأحداث من التطبيق: الخدمة تنشر أحداث أعمال، وهو الـ outbox من اليوم السابع. أفضل معنى، لكن فريق الخدمة يجب أن يبنيها ويصونها.
اسأل: أيها يستخدم فريقكم؟ وما الذي ساء فيه؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Moving the data", "Four ways to get the data out", false);
  tbl(s, 0.5, 1.85, [2.9, 4.8, 4.65], [
    ["Way", "How it works", "Problems"],
    ["Nightly full copy", "SELECT * from every table, every night.", "Heavy on production. Misses deleted rows. Nothing between two nights."],
    ["Incremental by timestamp", "Copy rows where updated_at is newer than the last run.", "Lighter, but misses deletes, and misses rows whose updated_at was not kept up to date."],
    ["Change data capture (CDC)", "Read the database's own change log. Send every insert, update and delete, in order.", "You need a tool, and must handle the first full copy and schema changes."],
    ["Events from the app", "The service publishes business events (the outbox from Day 7).", "Best meaning, but the service team must build and maintain them."],
  ], { fs: 12.5, rowH: [0.45, 0.95, 1.1, 1.1, 0.95] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: CDC in a picture. Every database already writes every change to a log, so it can recover after a crash; in Postgres this is the write-ahead log. A CDC tool such as Debezium reads that log, so the queries on the database are not touched: no extra queries on checkout. It publishes each change to Kafka, from Day 4, one topic per table. A job writes it to the raw layer. Then transformations build the modelled tables, and the dashboard reads them.
SAY: The result is near real time, with every insert, update and delete, in order.
ASK: What does CDC give you that a nightly copy never could?`,
`الوقت: حوالي 3 دقائق.
قل: التقاط التغييرات في صورة. كل قاعدة بيانات تكتب أصلًا كل تغيير في سجل لتتعافى بعد الانهيار؛ في Postgres هو سجل الكتابة المسبقة (WAL). أداة CDC مثل Debezium تقرأ ذلك السجل، فلا تُمَسّ الاستعلامات على القاعدة: لا استعلامات إضافية على الدفع. وتنشر كل تغيير إلى Kafka من اليوم الرابع، موضوع لكل جدول. ثم مهمة تكتبه في الطبقة الخام. ثم تبني التحويلات الجداول المصممة وتقرأها لوحة المؤشرات.
قل: النتيجة قريبة من الزمن الحقيقي، مع كل إدخال وتحديث وحذف بالترتيب.
اسأل: ماذا يعطيكم CDC مما لا تعطيه النسخة الليلية أبدًا؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Moving the data", "Change data capture, in a picture", false, { sub: "Every database already writes every change to a log. Read the log, not the tables." });
  flowRow(s, 2.1, [
    ["Postgres", "+ its change log", WHITE, INK_SOFT, 1.5],
    ["CDC tool", "reads the log", AMBER_TINT, AMBER, 1.5],
    ["Kafka", "one topic per table", PLUM_TINT, PLUM, 1.5],
    ["Raw layer", "lake or warehouse", TEAL_TINT, TEAL, 1.5],
    ["Modelled", "facts, dimensions", TEAL_TINT, TEAL, 2.5],
  ], { w: 2.1, h: 1.3, tfs: 13, sfs: 10.5 });
  label(s, "no extra queries on checkout", 0.5, 3.5, 2.1, 0.4, { size: 9.5, color: INK_FAINT });
  card(s, 0.5, 4.3, 12.35, 1.1, WHITE, { shadow: false, line: LINE });
  s.addText("Near real time, with every insert, update and delete, in order.", { x: 0.5, y: 4.3, w: 12.35, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: HEAD, bold: true, fontSize: 19, color: TEAL });
  s.addText("In Postgres the log is the write-ahead log. Debezium is a common CDC tool.", { x: 0.5, y: 5.75, w: 12.3, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT });
  footer(s, n, false);
});

N(`TIME: about 2 minutes.
SAY: This is what one change event looks like. The op field says what kind of change: c for create, u for update, d for delete. It carries the old row and the new row. And it says where in the log it came from: the LSN, the position, which is a ready-made event ID. Remember event IDs from Day 4: this is what you dedupe on.`,
`الوقت: حوالي دقيقتين.
قل: هكذا يبدو حدث تغيير واحد. حقل op يقول نوع التغيير: c للإنشاء، u للتحديث، d للحذف. ويحمل الصف القديم والصف الجديد. ويقول من أين جاء في السجل: الـ LSN، أي الموضع، وهو رقم حدث جاهز. تذكروا أرقام الأحداث من اليوم الرابع: هذا ما تزيلون التكرار به.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Moving the data", "One change event", false, { sub: "Old row, new row, what kind of change, and where in the log it came from." });
  codeBlock(s, 0.5, 1.85, 7.8, 4.3, "a change event (simplified)", TEAL, [
    "{",
    "  \"op\": \"u\",                                    # u = update",
    "  \"before\": { \"order_id\": 42, \"status\": \"pending\" },",
    "  \"after\":  { \"order_id\": 42, \"status\": \"paid\" },",
    "  \"source\": { \"table\": \"orders\", \"lsn\": 2318471 },   # position in the log",
    "  \"ts_ms\": 1790000000000",
    "}",
  ], 11.5);
  infoCard(s, 8.55, 1.85, 4.3, 2.0, TEAL, "What it carries", "op: c = create, u = update, d = delete. The old and the new row, so you can see exactly what changed.", 12);
  infoCard(s, 8.55, 4.0, 4.3, 2.15, PLUM, "A ready-made event ID", "The log position (lsn) is unique and ordered. Dedupe on it, as on Day 4.", 12.5);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: ETL or ELT: where does the cleaning happen? ETL, extract, transform, load, cleans the data on the way in, on a separate server, before it reaches the warehouse. ELT loads the raw data first and transforms it inside the warehouse, usually with SQL. A common tool for this is dbt.
SAY: Why ELT today? Storage and compute are cheap, and keeping the raw data means you can fix a bug in the logic and run it again. With ETL, if the transform was wrong, the raw data may be gone.`,
`الوقت: حوالي 3 دقائق.
قل: ETL أم ELT: أين يحدث التنظيف؟ ETL، استخراج ثم تحويل ثم تحميل، ينظف البيانات في الطريق، على خادم منفصل، قبل أن تصل إلى المستودع. أما ELT فيحمّل البيانات الخام أولًا ثم يحوّلها داخل المستودع، عادة بـ SQL. ومن الأدوات الشائعة لذلك dbt.
قل: لماذا ELT اليوم؟ التخزين والحساب رخيصان، والاحتفاظ بالبيانات الخام يعني أنكم تستطيعون إصلاح خطأ في المنطق وإعادة التشغيل. مع ETL، إذا كان التحويل خاطئًا فقد تكون البيانات الخام قد ضاعت.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Moving the data", "ETL or ELT: where does the cleaning happen?", false);
  label(s, "ETL: TRANSFORM FIRST", 0.5, 1.8, 5, 0.3, { size: 11, bold: true, color: AMBER, align: "left" });
  flowRow(s, 2.15, [["1  Extract", "", WHITE, INK_SOFT, 1.5], ["2  Transform", "on a separate server", AMBER_TINT, AMBER, 2.5], ["3  Load", "into the warehouse", WHITE, INK_SOFT, 1.5]], { w: 3.3, h: 1.1, tfs: 14, sfs: 11 });
  label(s, "ELT: LOAD FIRST", 0.5, 3.75, 5, 0.3, { size: 11, bold: true, color: TEAL, align: "left" });
  flowRow(s, 4.1, [["1  Extract", "", WHITE, INK_SOFT, 1.5], ["2  Load raw", "into the warehouse", WHITE, INK_SOFT, 1.5], ["3  Transform", "inside, with SQL (dbt is common)", TEAL_TINT, TEAL, 2.5]], { w: 3.3, h: 1.1, tfs: 14, sfs: 11 });
  s.addText("Storage and compute are cheap now, and keeping the raw data means you can fix a bug in the logic and run it again.", { x: 0.5, y: 5.65, w: 12.3, h: 0.8, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 13.5, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Batch or streaming. Batch processes a chunk of data on a schedule, every night or every hour. Delay: minutes to hours. Simpler to build, test and re-run. Use it for dashboards, reports and daily finance. Streaming processes each event as it arrives. Delay: seconds. Harder: ordering, late data, state and restarts. Use it for fraud checks, live inventory and alerts.
SAY: Do not stream what a nightly batch answers. Most dashboards need minutes or hours, not seconds. Ask the person who will read it how fresh it must be, and pick the simplest option that meets that.
ASK: How fresh does your most important dashboard really need to be?`,
`الوقت: حوالي 3 دقائق.
قل: دفعات أم بث. الدفعات تعالج مجموعة من البيانات حسب جدول، كل ليلة أو كل ساعة. التأخير: دقائق إلى ساعات. أبسط في البناء والاختبار وإعادة التشغيل. تُستخدم للوحات والتقارير والمالية اليومية. أما البث فيعالج كل حدث عند وصوله. التأخير: ثوانٍ. أصعب: الترتيب والبيانات المتأخرة والحالة وإعادة التشغيل. يُستخدم لفحص الاحتيال والمخزون الحي والتنبيهات.
قل: لا تبثّوا ما تجيب عنه دفعة ليلية. أغلب لوحات المؤشرات تحتاج دقائق أو ساعات لا ثوانٍ. اسألوا من سيقرأها كم يجب أن تكون حديثة، واختاروا أبسط خيار يحقق ذلك.
اسأل: ما مدى حداثة أهم لوحة مؤشرات عندكم فعلًا؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Moving the data", "Batch or streaming", false);
  tbl(s, 0.5, 1.85, [2.0, 5.1, 5.25], [
    ["", "Batch", "Streaming"],
    ["How", "Process a chunk of data on a schedule (every night, every hour)", "Process each event as it arrives"],
    ["Delay", "Minutes to hours", "Seconds"],
    ["Complexity", "Simpler to build, test and re-run", "Harder: ordering, late data, state, restarts"],
    ["Use when", "Dashboards, reports, daily finance", "Fraud checks, live inventory, alerts"],
  ], { fs: 13.5, rowH: [0.45, 0.8, 0.6, 0.8, 0.7] });
  callout(s, 0.5, 5.2, 12.35, 1.3, "Do not stream what a nightly batch answers.", "Most dashboards need minutes or hours, not seconds. Ask the person who will read it how fresh it must be, and pick the simplest option that meets that.", false);
  footer(s, n, false);
});

// ============================================================ STREAMS
N(`TIME: about 3 minutes.
SAY: Now the part people find hardest: processing a stream. A batch job has a file with an end. Count all of it, and you are done. A stream never ends. Events keep coming. So when do you stop counting, and give an answer?
SAY: The answer is a window: a slice of time. We count the orders in each five-minute slice. Streaming is mostly about windows, about time, and about what to do when events arrive out of order. The next slides take these one at a time.
ASK: Where would you need an answer "per minute" or "per session" in your system?`,
`الوقت: حوالي 3 دقائق.
قل: الآن الجزء الذي يجده الناس أصعب: معالجة التدفق (stream). مهمة الدفعات لها ملف له نهاية. عدّوه كله وانتهيتم. أما التدفق فلا ينتهي. الأحداث تستمر في الوصول. فمتى تتوقفون عن العد وتعطون جوابًا؟
قل: الجواب هو النافذة (window): شريحة من الزمن. نعدّ الطلبات في كل شريحة من خمس دقائق. معالجة التدفق في معظمها نوافذ، وزمن، وماذا نفعل عندما تصل الأحداث بترتيب مختلف. الشرائح التالية تأخذها واحدة واحدة.
اسأل: أين تحتاجون جوابًا "في الدقيقة" أو "في الجلسة" في نظامكم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "A stream never ends. So when do you count?", false, { sub: "Batch has a file with an end. A stream keeps coming." });
  card(s, 0.5, 1.85, 6.05, 3.2, WHITE, { shadow: false, line: LINE });
  label(s, "BATCH: a file with an end", 0.7, 1.95, 5.6, 0.3, { size: 10.5, bold: true, color: TEAL, align: "left" });
  for (let i = 0; i < 12; i++) s.addShape(pptx.ShapeType.rect, { x: 0.8 + i * 0.44, y: 2.6, w: 0.4, h: 0.5, fill: { color: TEAL_TINT }, line: { color: TEAL, width: 1 } });
  label(s, "count all of it, then you are done", 0.7, 3.3, 5.6, 0.4, { size: 11, color: INK_SOFT });
  s.addShape(pptx.ShapeType.line, { x: 6.1, y: 2.5, w: 0, h: 0.7, line: { color: TEAL, width: 3 } });
  card(s, 6.8, 1.85, 6.05, 3.2, WHITE, { shadow: false, line: LINE });
  label(s, "STREAM: events keep coming", 7.0, 1.95, 5.6, 0.3, { size: 10.5, bold: true, color: PLUM, align: "left" });
  for (let i = 0; i < 11; i++) s.addShape(pptx.ShapeType.ellipse, { x: 7.1 + i * 0.46, y: 2.7, w: 0.3, h: 0.3, fill: { color: PLUM_TINT }, line: { color: PLUM, width: 1 } });
  arrow(s, 12.2, 2.85, 12.7, 2.85, PLUM, { w: 2 });
  label(s, "no end. When do you give an answer?", 7.0, 3.3, 5.6, 0.4, { size: 11, color: INK_SOFT });
  callout(s, 0.5, 5.35, 12.35, 1.3, "The answer is a window:", "a slice of time. Count the orders in each five-minute slice. Streaming is mostly about windows, about time, and about events that arrive out of order.", false);
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: Three kinds of window. Tumbling: fixed size, no overlap. Every event belongs to exactly one window. Orders per five minutes, revenue per day. Sliding, also called hopping: fixed size, but a new window starts more often than it ends, so they overlap. Average response time over the last ten minutes, updated every five. One event can be in several windows. Session: no fixed size. A window groups events that are close together, and closes after a gap of silence. How long a visitor stayed on the site, ending after thirty minutes of inactivity.
ASK: In the picture, which window would you use to alert on "three failed payments within ten minutes"?`,
`الوقت: حوالي 4 دقائق.
قل: ثلاثة أنواع من النوافذ. المتتابعة (tumbling): حجم ثابت بلا تداخل. كل حدث ينتمي إلى نافذة واحدة بالضبط. الطلبات كل خمس دقائق، والإيراد كل يوم. المنزلقة (sliding أو hopping): حجم ثابت لكن تبدأ نافذة جديدة أكثر مما تنتهي، فتتداخل. متوسط زمن الاستجابة في آخر عشر دقائق، يُحدَّث كل خمس. الحدث الواحد قد يكون في عدة نوافذ. نافذة الجلسة (session): بلا حجم ثابت. تجمع الأحداث المتقاربة وتُغلق بعد فجوة صمت. كم بقي الزائر في الموقع، وتنتهي بعد ثلاثين دقيقة من عدم النشاط.
اسأل: في الصورة، أي نافذة تستخدمون للتنبيه على "ثلاث دفعات فاشلة خلال عشر دقائق"؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Three kinds of window", false, { sub: "A window is a slice of time. You count, sum or average inside each slice." });
  const X0 = 2.6, PPM = 0.66, tx = (t) => X0 + t * PPM;
  const evs = (y, list) => list.forEach((t) => s.addShape(pptx.ShapeType.ellipse, { x: tx(t) - 0.08, y: y - 0.08, w: 0.16, h: 0.16, fill: { color: INK }, line: { type: "none" } }));
  const rowLabel = (y, name, d, c) => { label(s, name, 0.5, y, 1.95, 0.35, { size: 11, bold: true, color: c, align: "left" }); label(s, d, 0.5, y + 0.35, 1.95, 0.5, { size: 8.5, color: INK_SOFT, align: "left", font: BODY }); };
  const ev1 = [1, 2.5, 4, 6, 7, 11, 12, 13.5];
  // tumbling
  rowLabel(1.95, "TUMBLING", "fixed size, no overlap", TEAL);
  s.addShape(pptx.ShapeType.line, { x: tx(0), y: 2.3, w: 15 * PPM, h: 0, line: { color: INK_FAINT, width: 1 } }); evs(2.3, ev1);
  [[0, 5, TEAL_TINT, TEAL], [5, 10, PLUM_TINT, PLUM], [10, 15, TEAL_TINT, TEAL]].forEach(([a, b, f, c]) => { s.addShape(pptx.ShapeType.rect, { x: tx(a) + 0.02, y: 2.55, w: (b - a) * PPM - 0.04, h: 0.45, fill: { color: f }, line: { color: c, width: 1.25 } }); label(s, "window", tx(a), 2.55, (b - a) * PPM, 0.45, { size: 9, color: c }); });
  // sliding
  rowLabel(3.5, "SLIDING", "size 10, a new one every 5", PLUM);
  s.addShape(pptx.ShapeType.line, { x: tx(0), y: 3.85, w: 15 * PPM, h: 0, line: { color: INK_FAINT, width: 1 } }); evs(3.85, ev1);
  s.addShape(pptx.ShapeType.rect, { x: tx(0) + 0.02, y: 4.1, w: 10 * PPM - 0.04, h: 0.4, fill: { color: TEAL_TINT }, line: { color: TEAL, width: 1.25 } }); label(s, "0 to 10", tx(0), 4.1, 10 * PPM, 0.4, { size: 9, color: TEAL });
  s.addShape(pptx.ShapeType.rect, { x: tx(5) + 0.02, y: 4.55, w: 10 * PPM - 0.04, h: 0.4, fill: { color: PLUM_TINT }, line: { color: PLUM, width: 1.25 } }); label(s, "5 to 15", tx(5), 4.55, 10 * PPM, 0.4, { size: 9, color: PLUM });
  // session
  rowLabel(5.2, "SESSION", "closes after a gap of silence", AMBER);
  const ev2 = [1, 2, 3.5, 10, 11, 12.5];
  s.addShape(pptx.ShapeType.line, { x: tx(0), y: 5.55, w: 15 * PPM, h: 0, line: { color: INK_FAINT, width: 1 } }); evs(5.55, ev2);
  [[0.7, 4.0], [9.7, 13.0]].forEach(([a, b]) => { s.addShape(pptx.ShapeType.rect, { x: tx(a), y: 5.8, w: (b - a) * PPM, h: 0.45, fill: { color: AMBER_TINT }, line: { color: AMBER, width: 1.25 } }); label(s, "session", tx(a), 5.8, (b - a) * PPM, 0.45, { size: 9, color: AMBER }); });
  label(s, "gap", tx(4.0), 5.8, 5.7 * PPM, 0.45, { size: 9, color: INK_FAINT, italic: true });
  [0, 5, 10, 15].forEach((t) => label(s, t + " min", tx(t) - 0.4, 6.4, 0.8, 0.25, { size: 8.5, color: INK_FAINT }));
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: In code, a window is one extra idea on top of the SQL you know. This query counts orders and sums revenue in tumbling five-minute windows: the same group by as in batch, but grouped by the window. The result is one row per window. Engines such as Flink, Spark Structured Streaming, Kafka Streams and ksqlDB all support this, and the exact syntax differs by engine.
SAY: Note that the last window is still open, so its numbers can still change.
ASK: What does the dashboard show for a window that is still open?`,
`الوقت: حوالي 3 دقائق.
قل: في الكود، النافذة فكرة إضافية واحدة فوق الـ SQL الذي تعرفونه. هذا الاستعلام يعد الطلبات ويجمع الإيراد في نوافذ متتابعة من خمس دقائق: نفس group by في الدفعات، لكن حسب النافذة. والنتيجة صف لكل نافذة. محركات مثل Flink وSpark Structured Streaming وKafka Streams وksqlDB كلها تدعم ذلك، والصياغة الدقيقة تختلف حسب المحرك.
قل: لاحظوا أن آخر نافذة ما زالت مفتوحة، فقد تتغير أرقامها.
اسأل: ماذا تعرض اللوحة لنافذة ما زالت مفتوحة؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "A window in a query", false, { sub: "The same GROUP BY as in batch, but grouped by the window." });
  codeBlock(s, 0.5, 1.85, 7.2, 2.9, "orders per five minutes (simplified)", TEAL, [
    "SELECT window_start,",
    "       COUNT(*)    AS orders,",
    "       SUM(amount) AS revenue",
    "FROM orders",
    "GROUP BY TUMBLE(order_time, INTERVAL '5' MINUTE)",
  ], 12.5);
  tbl(s, 7.95, 1.85, [1.9, 1.3, 1.7], [["window", "orders", "revenue"], ["12:00 to 12:05", "41", "1,830"], ["12:05 to 12:10", "57", "2,410"], ["12:10 to 12:15", "38", "1,505 (open)"]], { fs: 11.5, rowH: [0.45, 0.55, 0.55, 0.55] });
  infoCard(s, 0.5, 5.0, 6.05, 1.75, TEAL, "One row per window", "The result is a new row each time a window closes. The newest window is still open, so its numbers can still change.", 12);
  infoCard(s, 6.8, 5.0, 6.05, 1.75, PLUM, "Engines", "Flink, Spark Structured Streaming, Kafka Streams and ksqlDB all do this. The exact syntax differs by engine.", 12);
  footer(s, n, false);
});

N(`TIME: about 5 minutes.
SAY: Time is the tricky part. There are two clocks. Event time: when it happened, on the phone. Processing time: when it reached the pipeline. They differ, and events arrive out of order.
SAY: A watermark is the pipeline's guess: "I believe nothing older than this will arrive any more." Here the watermark is the newest event time minus two minutes. Walk the table. Event one at 12:01: the watermark is 11:59, the window 12:00 to 12:05 is open. Event three at 12:06 belongs to the next window, and moves the watermark to 12:04. Event four at 12:04 arrives late, after event three, but the window is still open, so it counts. Event five at 12:08 moves the watermark to 12:06, which passes 12:05, so the first window closes and its result is emitted. Event six at 12:02 arrives after that: too late for the window.
SAY: What to do with it? Drop it, send it to a side output to inspect, or fix the number later by re-running the day in batch. That is why we keep the raw data.
ASK: How long would you wait before closing a window, and who decides?`,
`الوقت: حوالي 5 دقائق.
قل: الزمن هو الجزء الصعب. هناك ساعتان. وقت الحدث (event time): متى حدث، على الهاتف. ووقت المعالجة (processing time): متى وصل إلى المسار. يختلفان، وتصل الأحداث بترتيب مختلف.
قل: العلامة المائية (watermark) هي تخمين المسار: "أعتقد أنه لن يصل شيء أقدم من هذا". هنا العلامة المائية هي أحدث وقت حدث ناقص دقيقتين. امشوا على الجدول. الحدث الأول 12:01: العلامة 11:59، والنافذة 12:00 إلى 12:05 مفتوحة. الحدث الثالث 12:06 ينتمي إلى النافذة التالية، ويحرك العلامة إلى 12:04. الحدث الرابع 12:04 يصل متأخرًا بعد الثالث، لكن النافذة ما زالت مفتوحة فيُحسب. الحدث الخامس 12:08 يحرك العلامة إلى 12:06 وهي تتجاوز 12:05، فتُغلق النافذة الأولى ويُصدر نتيجتها. الحدث السادس 12:02 يصل بعد ذلك: متأخر جدًا على النافذة.
قل: ماذا نفعل به؟ نرميه، أو نرسله إلى مخرج جانبي لنفحصه، أو نصحح الرقم لاحقًا بإعادة تشغيل اليوم بالدفعات. ولهذا نحتفظ بالبيانات الخام.
اسأل: كم تنتظرون قبل إغلاق نافذة، ومن يقرر؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Event time, watermarks and late events", false, { sub: "Watermark = the newest event time minus 2 minutes. Window under watch: 12:00 to 12:05." });
  tbl(s, 0.5, 1.8, [1.3, 1.9, 1.9, 7.25], [
    ["Arrival", "Event time", "Watermark", "What happens"],
    ["1", "12:01", "11:59", "The window 12:00 to 12:05 is open. Counted."],
    ["2", "12:03", "12:01", "Still open. Counted."],
    ["3", "12:06", "12:04", "Belongs to the next window. Moves the watermark."],
    ["4", "12:04 (late)", "12:04", "Arrived after event 3, but the window is still open. Counted."],
    ["5", "12:08", "12:06", "The watermark passes 12:05. The window closes and its result is emitted."],
    ["6", "12:02 (too late)", "12:06", "The window is closed. Drop it, send it to a side output, or fix it later in batch."],
  ], { fs: 12, rowH: [0.42, 0.5, 0.5, 0.5, 0.58, 0.58, 0.58] });
  infoCard(s, 0.5, 5.3, 4.0, 1.5, TEAL, "Watermark", "\"I believe nothing older than this will arrive.\"", 12);
  infoCard(s, 4.67, 5.3, 4.0, 1.5, AMBER, "The trade-off", "Wait longer: more accurate, slower. Wait less: faster, more late events lost.", 12);
  infoCard(s, 8.84, 5.3, 4.0, 1.5, PLUM, "Too late", "Fix it later by re-running the day from the raw data.", 12);
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: Counting per window needs memory: a count for every window and every key. This is state. If the job crashes, the state is lost, and the count restarts from zero. That is a wrong number.
SAY: The fix is checkpoints. Every minute or so, the engine saves a snapshot of the state and the Kafka offsets together, to a safe store. After a crash, it restarts from the last checkpoint and replays from the saved offsets. This is the offset idea from Day 4.
SAY: Replaying means some events are processed twice. So the effect on the outside is exactly once only if the output is safe to repeat: an idempotent write, or a transactional one. Exactly-once is at-least-once delivery plus a sink that handles repeats. Same rule as the whole course.
ASK: What does your stream job do with its state today? Is it saved anywhere?`,
`الوقت: حوالي 4 دقائق.
قل: العد لكل نافذة يحتاج ذاكرة: عدّاد لكل نافذة ولكل مفتاح. هذه هي الحالة (state). إذا انهارت المهمة ضاعت الحالة، ويبدأ العد من الصفر. وهذا رقم خاطئ.
قل: الحل هو نقاط التحقق (checkpoints). كل دقيقة تقريبًا يحفظ المحرك لقطة من الحالة ومواضع Kafka معًا في مخزن آمن. بعد الانهيار يعيد التشغيل من آخر نقطة تحقق ويعيد القراءة من المواضع المحفوظة. هذه فكرة الـ offset من اليوم الرابع.
قل: إعادة القراءة تعني أن بعض الأحداث تُعالج مرتين. فالأثر الخارجي مرة واحدة بالضبط (exactly-once) فقط إذا كان المخرج آمنًا عند التكرار: كتابة idempotent أو معاملاتية. أي أن exactly-once هو تسليم مرة على الأقل مع مخرج يتحمل التكرار. القاعدة نفسها في الدورة كلها.
اسأل: ماذا تفعل مهمة البث عندكم بحالتها اليوم؟ وهل تُحفظ في مكان؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "State and checkpoints", false, { sub: "Counting per window needs memory. A crash must not lose it." });
  box(s, 0.5, 2.0, 2.6, 1.5, PLUM_TINT, PLUM, 2); s.addText([{ text: "Kafka topic", options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true } }, { text: "events with offsets", options: { fontFace: BODY, fontSize: 11, color: INK_SOFT } }], { x: 0.55, y: 2.0, w: 2.5, h: 1.5, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 4.4, 1.85, 4.0, 1.8, TEAL_TINT, TEAL, 2.5); s.addText([{ text: "Stream job", options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 4 } }, { text: "state: a count for each window and key", options: { fontFace: BODY, fontSize: 11.5, color: TEAL } }], { x: 4.45, y: 1.85, w: 3.9, h: 1.8, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 9.7, 2.0, 3.1, 1.5, WHITE, INK_SOFT, 1.5); s.addText([{ text: "Sink", options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true } }, { text: "table or topic", options: { fontFace: BODY, fontSize: 11, color: INK_SOFT } }], { x: 9.75, y: 2.0, w: 3.0, h: 1.5, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 3.15, 2.75, 4.35, 2.75, INK_SOFT, { w: 1.6 }); arrow(s, 8.45, 2.75, 9.65, 2.75, INK_SOFT, { w: 1.6 });
  box(s, 4.4, 4.45, 4.0, 1.2, AMBER_TINT, AMBER, 2); s.addText([{ text: "Checkpoint store", options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true } }, { text: "state + offsets, saved together", options: { fontFace: BODY, fontSize: 11, color: AMBER } }], { x: 4.45, y: 4.45, w: 3.9, h: 1.2, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 6.4, 3.7, 6.4, 4.4, AMBER, { w: 1.8 }); label(s, "every minute or so", 6.55, 3.85, 2.2, 0.3, { size: 9, color: AMBER, align: "left" });
  s.addText("Crash? Restart from the last checkpoint, and replay from the saved offsets.", { x: 0.5, y: 4.3, w: 3.7, h: 1.2, isTextBox: true, margin: 0, fontFace: BODY, bold: true, fontSize: 13, color: RUST, valign: "top" });
  callout(s, 0.5, 5.95, 12.35, 0.85, "Exactly-once effect =", "at-least-once delivery plus a sink that is safe to repeat (an idempotent or transactional write).", false);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Joins in a stream. The common case is enrichment: an order event has a customer ID, and you want the segment. The trap is to call the Customers database for every event. That is the N plus 1 problem from Day 7, now at stream speed, and it will melt the database.
SAY: The fix is a stream-table join. Keep a local copy of the customers table inside the job, updated from the change data capture topic of the Customers database. Each event is joined against the local copy, with no remote call. For a small table you can also broadcast a copy to every worker. And a stream-stream join, for example matching a payment to its order, needs a window: match them if they arrive within thirty minutes of each other.
ASK: What would happen to the Customers database if every order called it?`,
`الوقت: حوالي 3 دقائق.
قل: الربط في التدفق. الحالة الشائعة هي الإثراء: حدث الطلب فيه رقم العميل وتريدون الشريحة. الفخ أن تستدعوا قاعدة العملاء لكل حدث. هذه مشكلة N+1 من اليوم السابع، الآن بسرعة التدفق، وستنهك القاعدة.
قل: الحل هو الربط بين تدفق وجدول. احتفظوا بنسخة محلية من جدول العملاء داخل المهمة، تُحدَّث من موضوع التقاط التغييرات (CDC) لقاعدة العملاء. ويُربط كل حدث بالنسخة المحلية دون استدعاء بعيد. ولجدول صغير يمكن أيضًا بث نسخة إلى كل عامل. أما الربط بين تدفقين، مثل مطابقة دفعة بطلبها، فيحتاج نافذة: يُطابَقان إذا وصلا خلال ثلاثين دقيقة من بعضهما.
اسأل: ماذا يحدث لقاعدة العملاء لو استدعاها كل طلب؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Joins in a stream", false, { sub: "An order event has a customer_id. You want the segment." });
  infoCard(s, 0.5, 1.85, 4.0, 2.6, RUST, "The trap", "Call the Customers database for every event. It is the N+1 problem from Day 7, at stream speed. It melts the database.", 13);
  infoCard(s, 4.67, 1.85, 4.0, 2.6, TEAL, "Stream-table join", "Keep a local copy of customers inside the job, updated from the CDC topic. Join each event to the copy. No remote call.", 13);
  infoCard(s, 8.84, 1.85, 4.0, 2.6, PLUM, "Stream-stream join", "Match a payment to its order. It needs a window: match them if they arrive within 30 minutes of each other.", 13);
  callout(s, 0.5, 4.8, 12.35, 1.1, "For a small table,", "you can also send a copy to every worker (a broadcast join). The rule is the same: never call out of the stream for each event.", false);
  footer(s, n, false);
});

// ---- Lambda
N(`TIME: about 4 minutes.
SAY: Lambda architecture, in a picture. Events come in from Kafka or files, and go down two paths at once. The top path is the batch layer. It recomputes the answer from all the raw data, say every night. It is accurate, but slow: its results are always hours old. The bottom path is the speed layer, a stream job that looks only at the newest events. It is fast, seconds, but approximate, because it has not seen everything and may not handle late events perfectly.
SAY: On the right is the serving layer. When someone asks a question, it merges both views: the batch view covers everything up to last night, the real-time view covers what has happened since. Add the two, and you get an answer that is both complete and fresh.
SAY: The price is on the slide. Two code paths, one in the batch engine and one in the stream engine, that must give the same answer. When they disagree, someone has to find out why.
ASK: Why is it hard to keep the two code paths in step?`,
`الوقت: حوالي 4 دقائق.
قل: معمارية لامدا في صورة. تأتي الأحداث من Kafka أو ملفات وتنزل في مسارين معًا. المسار العلوي هو طبقة الدفعات (batch layer). تعيد حساب الجواب من كل البيانات الخام، مثلًا كل ليلة. دقيقة لكنها بطيئة: نتائجها عمرها ساعات دائمًا. والمسار السفلي هو طبقة السرعة (speed layer)، مهمة بث تنظر إلى أحدث الأحداث فقط. سريعة، ثوانٍ، لكنها تقريبية، لأنها لم ترَ كل شيء وقد لا تعالج الأحداث المتأخرة بدقة.
قل: على اليمين طبقة التقديم (serving layer). عندما يسأل أحدهم تدمج العرضين: عرض الدفعات يغطي كل شيء حتى الليلة الماضية، وعرض الزمن الحقيقي يغطي ما حدث منذ ذلك الحين. اجمعوهما فتحصلون على جواب كامل وحديث معًا.
قل: الثمن على الشريحة. مساران للكود، واحد في محرك الدفعات وواحد في محرك البث، يجب أن يعطيا الجواب نفسه. وعندما يختلفان على أحدهم أن يعرف لماذا.
اسأل: لماذا يصعب إبقاء مساري الكود متطابقين؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Lambda architecture: two paths", false, { sub: "A batch path for accuracy and a speed path for freshness, merged when you query." });
  box(s, 0.5, 3.35, 1.9, 1.4, WHITE, INK_SOFT, 2); s.addText([{ text: "Events", options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true } }, { text: "Kafka or files", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 0.55, y: 3.35, w: 1.8, h: 1.4, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  // batch lane
  card(s, 3.0, 1.8, 6.55, 1.7, WHITE, { shadow: false, line: TEAL });
  label(s, "BATCH LAYER: accurate, slow", 3.2, 1.87, 6.1, 0.28, { size: 9.5, bold: true, color: TEAL, align: "left" });
  box(s, 3.2, 2.2, 3.0, 1.1, TEAL_TINT, TEAL, 1.5); s.addText([{ text: "Recompute everything", options: { fontFace: HEAD, bold: true, fontSize: 12.5, color: INK, breakLine: true } }, { text: "from all the raw data, nightly", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 3.25, y: 2.2, w: 2.9, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 6.65, 2.2, 2.7, 1.1, TEAL_TINT, TEAL, 1.5); s.addText([{ text: "Batch views", options: { fontFace: HEAD, bold: true, fontSize: 12.5, color: INK, breakLine: true } }, { text: "complete, hours old", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 6.7, y: 2.2, w: 2.6, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 6.22, 2.75, 6.63, 2.75, TEAL, { w: 1.6 });
  // speed lane
  card(s, 3.0, 4.2, 6.55, 1.7, WHITE, { shadow: false, line: AMBER });
  label(s, "SPEED LAYER: fast, approximate", 3.2, 4.27, 6.1, 0.28, { size: 9.5, bold: true, color: AMBER, align: "left" });
  box(s, 3.2, 4.6, 3.0, 1.1, AMBER_TINT, AMBER, 1.5); s.addText([{ text: "Stream job", options: { fontFace: HEAD, bold: true, fontSize: 12.5, color: INK, breakLine: true } }, { text: "only the newest events", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 3.25, y: 4.6, w: 2.9, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 6.65, 4.6, 2.7, 1.1, AMBER_TINT, AMBER, 1.5); s.addText([{ text: "Real-time views", options: { fontFace: HEAD, bold: true, fontSize: 12.5, color: INK, breakLine: true } }, { text: "since the last batch, seconds old", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 6.7, y: 4.6, w: 2.6, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 6.22, 5.15, 6.63, 5.15, AMBER, { w: 1.6 });
  // serving
  box(s, 10.2, 3.35, 2.65, 1.4, PLUM_TINT, PLUM, 2.5); s.addText([{ text: "Serving layer", options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 3 } }, { text: "merges both views when you query", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 10.25, y: 3.35, w: 2.55, h: 1.4, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 2.45, 3.7, 2.95, 2.8, INK_SOFT, { w: 1.6 }); arrow(s, 2.45, 4.4, 2.95, 5.3, INK_SOFT, { w: 1.6 });
  arrow(s, 9.4, 2.9, 10.15, 3.7, TEAL, { w: 1.6 }); arrow(s, 9.4, 5.1, 10.15, 4.4, AMBER, { w: 1.6 });
  callout(s, 0.5, 6.15, 12.35, 0.8, "The price:", "two code paths that must give the same answer. The batch view covers everything up to last night. The speed view covers what has happened since.", false);
  footer(s, n, false);
});

// ---- Kappa
N(`TIME: about 4 minutes.
SAY: Kappa architecture removes the second path. There is only one: a stream job. The key is the Kafka log on the left. It keeps the events for a long time, days, weeks or forever, as a record you can read again from any offset.
SAY: Normally, job version one reads the newest events and keeps a table up to date. Now you find a bug, or you want to change the logic. In Lambda you would change two code bases. In Kappa you start job version two, with the new logic, and it replays the log from offset zero into a new table. It runs alongside version one. When version two has caught up with the present, you switch queries to the new table, and stop version one.
SAY: The gain: one code base. The price: you need a log that keeps the data long enough, state you can rebuild, and enough capacity to replay quickly.
ASK: Which is easier for your team to run: two code paths, or a long-lived log and replays?`,
`الوقت: حوالي 4 دقائق.
قل: معمارية كابا تزيل المسار الثاني. يوجد مسار واحد فقط: مهمة بث. المفتاح هو سجل Kafka على اليسار. يحتفظ بالأحداث مدة طويلة، أيامًا أو أسابيع أو للأبد، كسجل يمكن قراءته من جديد من أي موضع (offset).
قل: عادة تقرأ المهمة بالإصدار الأول أحدث الأحداث وتبقي جدولًا محدّثًا. الآن وجدتم خطأً أو تريدون تغيير المنطق. في لامدا ستغيّرون قاعدتي كود. في كابا تشغّلون المهمة بالإصدار الثاني، بالمنطق الجديد، فتعيد قراءة السجل من الموضع صفر إلى جدول جديد. تعمل بجانب الإصدار الأول. وعندما يلحق الإصدار الثاني بالحاضر تحوّلون الاستعلامات إلى الجدول الجديد وتوقفون الإصدار الأول.
قل: المكسب: قاعدة كود واحدة. الثمن: تحتاجون سجلًا يحتفظ بالبيانات مدة كافية، وحالة يمكن إعادة بنائها، وطاقة كافية لإعادة القراءة بسرعة.
اسأل: أيهما أسهل لفريقكم: مساران للكود، أم سجل طويل العمر وإعادات قراءة؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Kappa architecture: one path, and replay", false, { sub: "One stream job. To change the logic, replay the log into a new table, then switch." });
  card(s, 0.5, 2.3, 5.2, 2.4, WHITE, { shadow: false, line: PLUM });
  label(s, "KAFKA LOG, KEPT LONG ENOUGH", 0.7, 2.38, 4.8, 0.28, { size: 9.5, bold: true, color: PLUM, align: "left" });
  for (let i = 0; i < 12; i++) { s.addShape(pptx.ShapeType.rect, { x: 0.7 + i * 0.4, y: 3.0, w: 0.36, h: 0.5, fill: { color: i < 8 ? PLUM_TINT : TEAL_TINT }, line: { color: i < 8 ? PLUM : TEAL, width: 1 } }); label(s, String(i), 0.7 + i * 0.4, 3.0, 0.36, 0.5, { size: 8 }); }
  label(s, "offset 0", 0.6, 3.6, 1.0, 0.3, { size: 8.5, color: INK_SOFT, align: "left" }); label(s, "newest", 4.5, 3.6, 1.0, 0.3, { size: 8.5, color: TEAL });
  label(s, "events are kept, not deleted when read", 0.7, 4.05, 4.8, 0.4, { size: 9.5, color: INK_SOFT, font: BODY });
  box(s, 6.5, 1.85, 3.0, 1.35, TEAL_TINT, TEAL, 2); s.addText([{ text: "Job v1 (running)", options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: "reads the newest events", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 6.55, y: 1.85, w: 2.9, h: 1.35, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.5, y: 4.0, w: 3.0, h: 1.35, rectRadius: 0.07, fill: { color: PLUM_TINT }, line: { color: PLUM, width: 2, dashType: "dash" } });
  s.addText([{ text: "Job v2 (new logic)", options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: "replays from offset 0", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 6.55, y: 4.0, w: 2.9, h: 1.35, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 10.2, 1.85, 2.65, 1.35, WHITE, TEAL, 1.5); s.addText([{ text: "Table v1", options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: "serves queries now", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 10.25, y: 1.85, w: 2.55, h: 1.35, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  s.addShape(pptx.ShapeType.roundRect, { x: 10.2, y: 4.0, w: 2.65, h: 1.35, rectRadius: 0.07, fill: { color: WHITE }, line: { color: PLUM, width: 1.5, dashType: "dash" } });
  s.addText([{ text: "Table v2", options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: "being rebuilt", options: { fontFace: BODY, fontSize: 10.5, color: INK_SOFT } }], { x: 10.25, y: 4.0, w: 2.55, h: 1.35, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 5.75, 3.2, 6.45, 2.6, TEAL, { w: 1.8 }); arrow(s, 5.75, 3.8, 6.45, 4.6, PLUM, { w: 1.8, dash: "dash" });
  arrow(s, 9.55, 2.5, 10.15, 2.5, TEAL, { w: 1.8 }); arrow(s, 9.55, 4.7, 10.15, 4.7, PLUM, { w: 1.8, dash: "dash" });
  s.addText("When v2 has caught up, switch queries to table v2 and stop v1.", { x: 6.5, y: 5.5, w: 6.35, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: BODY, bold: true, fontSize: 12.5, color: PLUM });
  callout(s, 0.5, 6.1, 12.35, 0.85, "The gain:", "one code base. The price: a log that keeps the data long enough, state you can rebuild, and enough capacity to replay quickly.", false);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: How to organise batch and streaming together. Lambda architecture: two paths. A batch path that is accurate but slow, and a speed path that is fast but approximate, merged when serving. The cost is two code bases that must agree. Kappa architecture: one streaming path only, and when you need to reprocess you replay the log from the start. One code base, but it needs a log that keeps the data long enough, and state you can rebuild.
SAY: And the simple option many teams should pick: batch every five minutes, called micro-batch. It is easy to build, test and re-run, and for most dashboards it is fresh enough. Pick the simplest option that meets the freshness the reader needs.
ASK: Does anyone run two code paths for the same number today? How do they stay in sync?`,
`الوقت: حوالي 3 دقائق.
قل: كيف ننظم الدفعات والتدفق معًا. معمارية لامدا (Lambda): مساران. مسار دفعات دقيق لكنه بطيء، ومسار سرعة سريع لكنه تقريبي، ويُدمجان عند التقديم. الكلفة قاعدتا كود يجب أن تتفقا. معمارية كابا (Kappa): مسار تدفق واحد فقط، وعندما تحتاجون إعادة المعالجة تعيدون قراءة السجل من البداية. قاعدة كود واحدة، لكنها تحتاج سجلًا يحتفظ بالبيانات مدة كافية وحالة يمكن إعادة بنائها.
قل: والخيار البسيط الذي ينبغي لفرق كثيرة اختياره: دفعة كل خمس دقائق، وتسمى micro-batch. سهلة البناء والاختبار وإعادة التشغيل، ولأغلب لوحات المؤشرات حداثتها كافية. اختاروا أبسط خيار يحقق الحداثة التي يحتاجها القارئ.
اسأل: هل يشغّل أحدكم مسارَي كود للرقم نفسه اليوم؟ وكيف يبقيان متفقين؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Lambda, Kappa, or just micro-batch?", false, { sub: "How to organise batch and streaming together." });
  tbl(s, 0.5, 1.85, [2.3, 5.0, 5.05], [
    ["", "How it works", "The price"],
    ["Lambda", "Two paths: a batch path (accurate, slow) and a speed path (fast, approximate), merged when serving.", "Two code bases that must give the same answer."],
    ["Kappa", "One streaming path. To reprocess, replay the log from the start.", "Needs a log that keeps the data long enough, and state you can rebuild."],
    ["Micro-batch", "Plain batch, every few minutes (for example every 5).", "Delay of minutes. But easy to build, test and re-run."],
  ], { fs: 13, rowH: [0.45, 1.0, 0.95, 0.9] });
  callout(s, 0.5, 5.4, 12.35, 1.3, "Pick the simplest option", "that meets the freshness the reader needs. Many dashboards are fine with a batch every five minutes.", false);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Five failures you will meet in a streaming pipeline, and most are old friends from Day 4. Backpressure: the consumer is slower than the producer, lag grows. Scale the partitions and consumers, alert on lag. Bad records: one message always fails and blocks the partition. Retry a few times, then dead-letter it. Skew: one key sends most of the traffic and one partition is a hot spot. Salt the key. Schema drift: a producer changes a field. Use a contract and a schema check. And reprocessing: you fixed a bug and need to redo the past. Replay from the log, into a sink that is safe to repeat.
ASK: Which of these five has bitten your team?`,
`الوقت: حوالي 3 دقائق.
قل: خمسة إخفاقات ستواجهونها في مسار التدفق، ومعظمها معارف قديمة من اليوم الرابع. الضغط العكسي (backpressure): المستهلك أبطأ من المنتِج فيكبر التأخر. وسّعوا الأقسام والمستهلكين ونبّهوا على التأخر. السجلات السيئة: رسالة تفشل دائمًا وتسد القسم. أعيدوا المحاولة عدة مرات ثم أرسلوها إلى طابور الرسائل الميتة. الانحياز (skew): مفتاح واحد يرسل معظم الحركة وقسم واحد يصير نقطة ساخنة. أضيفوا ملحًا للمفتاح. انجراف المخطط: منتِج يغيّر حقلًا. استخدموا عقدًا وفحص مخطط. وإعادة المعالجة: أصلحتم خطأ وتريدون إعادة الماضي. أعيدوا القراءة من السجل إلى مخرج آمن عند التكرار.
اسأل: أي هذه الخمسة أصاب فريقكم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Streams", "Five failures you will meet", false, { sub: "Most of them are old friends from Day 4." });
  [
    [RUST, "Backpressure", "The consumer is slower than the producer, and lag grows. Add partitions and consumers. Alert on lag."],
    [AMBER, "Bad records", "One message always fails and blocks the partition. Retry a few times, then dead-letter it."],
    [PLUM, "Skew", "One key sends most of the traffic, and one partition is a hot spot. Add a random part to the key."],
    [TEAL, "Schema drift", "A producer changes a field. Use a contract and a schema check at the boundary."],
    [SLATE, "Reprocessing", "You fixed a bug and need to redo the past. Replay from the log, into a sink that is safe to repeat."],
  ].forEach(([c, t, d], i) => {
    const y = 1.8 + i * 0.99;
    card(s, 0.5, y, 12.35, 0.88, WHITE);
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 0.09, h: 0.88, fill: { color: c }, line: { type: "none" } });
    s.addText([{ text: t + "   ", options: { bold: true, color: c, fontFace: HEAD, fontSize: 15 } }, { text: d, options: { color: INK_SOFT, fontFace: BODY, fontSize: 13 } }], { x: 0.8, y, w: 11.9, h: 0.88, isTextBox: true, margin: 0, valign: "middle" });
  });
  footer(s, n, false);
});

N(`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Give pairs two minutes. For each of the five needs, they pick tumbling, sliding or session. Then reveal.
SAY: The test: does the question have a fixed calendar slice (tumbling)? A rolling look back that updates often (sliding)? Or does it follow bursts of activity (session)?`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: أعطِ الأزواج دقيقتين. لكل من الاحتياجات الخمسة يختارون: متتابعة أم منزلقة أم جلسة. ثم اكشف الإجابات.
قل: الاختبار: هل للسؤال شريحة تقويمية ثابتة (متتابعة)؟ أم نظرة متدحرجة للخلف تتحدث كثيرًا (منزلقة)؟ أم يتبع دفعات النشاط (جلسة)؟`);
checkIn(3, 3, "Which window?", "For each need, pick a window: tumbling, sliding or session.", [
  "Orders per five minutes, shown on a dashboard.",
  "Average response time over the last 10 minutes, updated every minute.",
  "How long a visitor stays, ending after 30 minutes of silence.",
  "Alert when one card has 3 failed payments within 10 minutes.",
  "Revenue per calendar day.",
], null, "a. Tumbling. A fixed slice, no overlap.\nb. Sliding. A rolling look back, updated every minute.\nc. Session. It follows the visitor's burst of activity and closes after a gap.\nd. Sliding. The 10 minutes must roll, so that a burst across a boundary is not missed.\ne. Tumbling. One window per calendar day.\n\nThe test: a fixed calendar slice is tumbling. A rolling look back is sliding. Bursts of activity are session.");

N(`TIME: about 3 minutes.
SAY: Pipelines are many steps with dependencies: load, clean, model, then check. An orchestrator, Airflow, Dagster or Prefect, runs them in the right order, on a schedule, retries failed steps, and lets you backfill: re-run the past for a range of days after fixing a bug.
SAY: Backfills only work if each step is safe to run twice. That is the next slide.`,
`الوقت: حوالي 3 دقائق.
قل: المسارات خطوات كثيرة بينها اعتماديات: تحميل، تنظيف، تصميم، ثم فحص. المنسّق (orchestrator)، Airflow أو Dagster أو Prefect، يشغّلها بالترتيب الصحيح وحسب جدول، ويعيد الخطوات الفاشلة، ويتيح الـ backfill: إعادة تشغيل الماضي لمدى من الأيام بعد إصلاح خطأ.
قل: الـ backfill لا ينجح إلا إذا كانت كل خطوة آمنة عند تشغيلها مرتين. وهذه الشريحة التالية.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Running it", "An orchestrator runs the steps in order", false, { sub: "Airflow, Dagster and Prefect are common. They schedule, retry and backfill." });
  flowRow(s, 2.0, [["Load raw", "", AMBER_TINT, AMBER, 1.5], ["Clean", "", TEAL_TINT, TEAL, 1.5], ["Model", "", TEAL_TINT, TEAL, 1.5], ["Check", "stops the run if it fails", PLUM_TINT, PLUM, 2.5]], { w: 2.5, h: 1.1, tfs: 14, sfs: 10.5 });
  infoCard(s, 0.5, 3.6, 4.0, 2.4, TEAL, "Schedule", "Run every night, every hour, or when new data lands.", 13);
  infoCard(s, 4.67, 3.6, 4.0, 2.4, AMBER, "Retry", "If a step fails, try it again, a few times, before waking a person.", 13);
  infoCard(s, 8.84, 3.6, 4.0, 2.4, PLUM, "Backfill", "Fixed a bug? Re-run the past, for a range of days. Only works if each step is safe to run twice.", 13);
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: Two versions of the same load. Top: append. The load reads raw orders for the day and appends them to the fact table. Run it twice, a retry, a restart, a backfill, and every order is counted twice. Revenue doubles.
SAY: Bottom: replace the day. Read the raw orders, dedupe by event ID, so the same event twice becomes one, and overwrite the partition for that day. Replace the whole day, never append. Then a retry, a restart and a backfill all give the same result. This is the idempotent consumer from Day 4, in a new place.
ASK: Which of your pipelines append? What happens when they are retried?`,
`الوقت: حوالي 4 دقائق.
قل: نسختان من التحميل نفسه. الأعلى: إضافة (append). يقرأ التحميل طلبات اليوم الخام ويضيفها إلى جدول الحقائق. شغّلوه مرتين، إعادة محاولة أو إعادة تشغيل أو backfill، فيُحسب كل طلب مرتين. يتضاعف الإيراد.
قل: الأسفل: استبدال اليوم. اقرأوا الطلبات الخام، وأزيلوا التكرار برقم الحدث، فيصير الحدث نفسه مرتين حدثًا واحدًا، واستبدلوا قسم ذلك اليوم. استبدلوا اليوم كله ولا تضيفوا أبدًا. فتعطي إعادة المحاولة وإعادة التشغيل والـ backfill النتيجة نفسها. هذا هو المستهلك الآمن عند التكرار من اليوم الرابع، في مكان جديد.
اسأل: أي مساراتكم تضيف؟ وماذا يحدث عند إعادة محاولتها؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Running it again", "Replace the day, never append", false, { sub: "A pipeline must be safe to run twice. This is the idempotent consumer from Day 4." });
  codeBlock(s, 0.5, 1.8, 6.05, 3.4, "load_orders.py: append", RUST, [
    "def load_orders(day):",
    "    rows = read_raw(\"orders\", day)",
    "    append(\"fact_orders\", rows)",
    "",
    "# run it twice and every",
    "# order is counted twice",
  ], 12);
  codeBlock(s, 6.8, 1.8, 6.05, 3.4, "load_orders.py: replace the day", TEAL, [
    "def load_orders(day):",
    "    rows = read_raw(\"orders\", day)",
    "    rows = dedupe(rows, key=\"event_id\")",
    "    overwrite_partition(",
    "        \"fact_orders\", day, rows)",
    "# replace the whole day",
  ], 12);
  callout(s, 0.5, 5.5, 12.35, 1.2, "The rule:", "replace the day, never append to it. Then a retry, a restart and a backfill all give the same result.", false);
  footer(s, n, false);
});

// ============================================================ PART 4
N(`TIME: 30 seconds.
SAY: Part 4. When a service breaks, you get an error and an alert. When a pipeline breaks, you often get a number that is quietly wrong, and a decision made on it. Trust has to be built in.`,
`الوقت: 30 ثانية.
قل: الجزء الرابع. عندما تتعطل خدمة تحصلون على خطأ وتنبيه. وعندما يتعطل مسار بيانات تحصلون غالبًا على رقم خاطئ بصمت وقرار مبني عليه. الثقة يجب أن تُبنى من الداخل.`);
divider(4, "1:30–2:00", "Trusting the data", "Checks that stop a bad load, contracts, duplicates and late data, and personal data.");

N(`TIME: about 3 minutes.
SAY: Six checks that stop a bad load. Freshness: did the newest data arrive on time? It catches a stuck pipeline. Volume: is the row count close to normal? It catches a half-loaded day or a double load. Uniqueness: does any key appear twice? Duplicates. Completeness: are required fields filled? It catches a renamed field that arrives empty. Validity: are values in range? Reconciliation: does the total match the source? It catches lost rows.
ASK: Which of these six does your most important table have today?`,
`الوقت: حوالي 3 دقائق.
قل: ست فحوص توقف التحميل السيئ. الحداثة (freshness): هل وصلت أحدث البيانات في وقتها؟ تلتقط مسارًا متوقفًا. الحجم (volume): هل عدد الصفوف قريب من المعتاد؟ تلتقط يومًا محمّلًا جزئيًا أو تحميلًا مزدوجًا. التفرد (uniqueness): هل يظهر مفتاح مرتين؟ التكرار. الاكتمال (completeness): هل الحقول المطلوبة معبأة؟ تلتقط حقلًا أُعيدت تسميته فيصل فارغًا. الصلاحية (validity): هل القيم في النطاق؟ والمطابقة (reconciliation): هل المجموع يطابق المصدر؟ تلتقط الصفوف الضائعة.
اسأل: أي هذه الفحوص الستة موجود اليوم على أهم جداولكم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Trusting the data", "Checks that stop a bad load", false);
  tbl(s, 0.5, 1.85, [2.4, 5.0, 4.95], [
    ["Check", "What it asks", "It catches"],
    ["Freshness", "Did the newest data arrive on time?", "A stuck pipeline, and a dashboard showing yesterday"],
    ["Volume", "Is the row count close to normal?", "A half-loaded day, or a double load"],
    ["Uniqueness", "Does any key appear twice?", "Duplicates and double counting"],
    ["Completeness", "Are required fields filled?", "A renamed field that arrives empty"],
    ["Validity", "Are values in range (amount above zero, known status)?", "Bad inputs and wrong units"],
    ["Reconciliation", "Does the total match the source?", "Lost rows between the source and the warehouse"],
  ], { fs: 12.5, rowH: [0.42, 0.65, 0.65, 0.65, 0.65, 0.75, 0.75] });
  footer(s, n, false);
});

N(`TIME: about 2 minutes.
SAY: The checks are just queries. The pipeline stops if any of them returns a row. Uniqueness: any order ID that appears more than once. Volume: fewer than a thousand rows for the day is wrong. Freshness: nothing newer than thirty minutes. Run them after every load.
NOTE: In dbt these are built-in tests, in other tools they are called expectations. The idea is the same.`,
`الوقت: حوالي دقيقتين.
قل: الفحوص مجرد استعلامات. يتوقف المسار إذا أعاد أي منها صفًا. التفرد: أي رقم طلب يظهر أكثر من مرة. الحجم: أقل من ألف صف لليوم خطأ. الحداثة: لا شيء أحدث من ثلاثين دقيقة. شغّلوها بعد كل تحميل.
ملاحظة: في dbt هي اختبارات مدمجة، وفي أدوات أخرى تسمى توقعات (expectations). الفكرة نفسها.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Trusting the data", "The checks are just queries", false, { sub: "The pipeline stops if any of these returns a row. Run them after every load." });
  codeBlock(s, 0.5, 1.85, 12.35, 4.6, "checks.sql", TEAL, [
    "-- uniqueness",
    "SELECT order_id FROM fact_orders GROUP BY order_id HAVING COUNT(*) > 1;",
    "",
    "-- volume: fewer than 1,000 rows for the day is wrong",
    "SELECT 1 FROM fact_orders WHERE order_date = '2026-10-07' HAVING COUNT(*) < 1000;",
    "",
    "-- freshness: nothing newer than 30 minutes",
    "SELECT 1 FROM fact_orders HAVING MAX(loaded_at) < NOW() - INTERVAL '30 minutes';",
  ], 13);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: The Orders team renames amount to total. The service still works, and the dashboard column goes empty. This is the Day 7 contract problem on the data side. The fix is the same: the producer publishes a schema, a schema registry with Avro or Protobuf is a common way, and a change that would break a consumer fails the producer's build.
SAY: Which changes break the pipeline? Adding an optional field: no. Renaming a field: yes; add the new name, keep the old for a while, then remove it. Changing a type: yes. Removing a field: yes, if anyone still reads it. A new value for a fixed list, like a new status: maybe, check anything that sorts or filters on the list.
ASK: QEs: do you have a schema check at the boundary today?`,
`الوقت: حوالي 3 دقائق.
قل: فريق الطلبات يعيد تسمية amount إلى total. الخدمة ما زالت تعمل، وعمود لوحة المؤشرات يصير فارغًا. هذه مشكلة العقود من اليوم السابع على جانب البيانات. والحل نفسه: المنتِج ينشر مخططًا، وسجل المخططات (schema registry) مع Avro أو Protobuf طريقة شائعة، والتغيير الذي سيكسر مستهلكًا يفشّل بناء المنتِج.
قل: أي التغييرات تكسر المسار؟ إضافة حقل اختياري: لا. إعادة تسمية حقل: نعم؛ أضيفوا الاسم الجديد وأبقوا القديم مدة ثم احذفوه. تغيير النوع: نعم. حذف حقل: نعم إن كان أحد ما زال يقرؤه. قيمة جديدة في قائمة ثابتة مثل حالة جديدة: ربما، افحصوا كل ما يرتب أو يصفّي على القائمة.
اسأل: مهندسو الجودة: هل لديكم فحص مخطط عند الحد اليوم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Trusting the data", "Contracts: when the source changes", false, { sub: "The Orders team renames amount to total. The service works. The dashboard column goes empty." });
  tbl(s, 0.5, 1.85, [5.4, 6.95], [
    ["Change to a field", "Breaks the pipeline?"],
    ["Add an optional field", "No"],
    ["Rename a field", "Yes. Add the new name, keep the old for a while, then remove it."],
    ["Change a type (number to text)", "Yes"],
    ["Remove a field", "Yes, if anyone still reads it"],
    ["A new value for a fixed list (a new status)", "Maybe. Check anything that sorts or filters on the list."],
  ], { fs: 13.5, rowH: [0.45, 0.6, 0.8, 0.6, 0.6, 0.8] });
  callout(s, 0.5, 5.55, 12.35, 1.2, "The fix is the Day 7 contract:", "the producer publishes a schema (a schema registry with Avro or Protobuf is common), and a change that would break a consumer fails the producer's build.", false);
  footer(s, n, false);
});

// ---- lineage
N(`TIME: about 3 minutes.
SAY: Lineage answers: where did this number come from? Read the picture from right to left. The dashboard tile, revenue for the quarter, comes from fact orders, which comes from cleaned orders, which comes from raw orders, which comes from the CDC topic, which comes from the Postgres orders table.
SAY: Three uses. Find the source of a wrong number: follow the chain back until it breaks. Impact analysis: before you rename a field, see what depends on it, the contract problem again. And privacy: see where personal data goes, so a delete can follow it. Tools: OpenLineage, dbt docs, data catalogs. Also add an owner and a freshness promise to each table, so someone is responsible.
ASK: Pick a number in your dashboard. Could you draw this chain for it right now?`,
`الوقت: حوالي 3 دقائق.
قل: تتبع النسب (lineage) يجيب: من أين جاء هذا الرقم؟ اقرأوا الصورة من اليمين إلى اليسار. بطاقة اللوحة، إيراد الربع، تأتي من fact_orders، التي تأتي من الطلبات المنظفة، التي تأتي من الطلبات الخام، التي تأتي من موضوع CDC، الذي يأتي من جدول الطلبات في Postgres.
قل: ثلاثة استخدامات. إيجاد مصدر رقم خاطئ: اتبعوا السلسلة للخلف حتى تنكسر. تحليل الأثر: قبل إعادة تسمية حقل، انظروا ما يعتمد عليه، مشكلة العقود مرة أخرى. والخصوصية: رؤية أين تذهب البيانات الشخصية ليتبعها الحذف. الأدوات: OpenLineage ووثائق dbt وفهارس البيانات (data catalogs). وأضيفوا أيضًا مالكًا ووعدًا بالحداثة لكل جدول، ليكون أحدهم مسؤولًا.
اسأل: اختاروا رقمًا في لوحتكم. هل تستطيعون رسم هذه السلسلة له الآن؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Trusting the data", "Lineage: where did this number come from?", false, { sub: "Read the chain from right to left, until you find the source." });
  const w = 1.85, gap = (12.35 - 6 * w) / 5, X = (i) => 0.5 + (5 - i) * (w + gap);
  [
    ["Dashboard tile", "revenue, Q3", PLUM_TINT, PLUM],
    ["fact_orders", "modelled", PLUM_TINT, PLUM],
    ["clean_orders", "deduplicated", TEAL_TINT, TEAL],
    ["raw_orders", "append only", AMBER_TINT, AMBER],
    ["CDC topic", "orders", PLUM_TINT, PLUM],
    ["Postgres", "orders table", WHITE, INK_SOFT],
  ].forEach(([t, sub, f, c], i) => {
    box(s, X(i), 2.0, w, 1.2, f, c, i === 0 ? 2.5 : 1.5);
    s.addText([{ text: t, options: { fontFace: HEAD, bold: true, fontSize: 12, color: INK, breakLine: true } }, { text: sub, options: { fontFace: BODY, fontSize: 10, color: INK_SOFT } }], { x: X(i) + 0.05, y: 2.0, w: w - 0.1, h: 1.2, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    if (i < 5) arrow(s, X(i + 1) + w + 0.03, 2.6, X(i) - 0.03, 2.6, RUST, { w: 1.5 });
  });
  label(s, "follow it back", 0.5, 3.4, 12.35, 0.3, { size: 10, color: RUST, bold: true });
  infoCard(s, 0.5, 4.0, 4.0, 2.0, TEAL, "Find the source of a wrong number", "Follow the chain back until it breaks.", 13);
  infoCard(s, 4.67, 4.0, 4.0, 2.0, AMBER, "Impact analysis", "Before you rename a field, see what depends on it. The contract problem again.", 13);
  infoCard(s, 8.84, 4.0, 4.0, 2.0, RUST, "Privacy", "See where personal data goes, so a delete can follow it.", 13);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Duplicates and late data. At-least-once delivery from Day 4 means the same event can arrive twice, so every pipeline needs a key to dedupe on. Events also arrive late: a phone was offline, and its order reaches the pipeline two days after it was placed.
SAY: Two rules. Group by event time, when it happened, not processing time, when it arrived. Otherwise Tuesday's revenue changes depending on when the phone reconnects. And allow for late data: when a late event arrives, re-run the day it belongs to. This is why the pipeline must be safe to run again.`,
`الوقت: حوالي 3 دقائق.
قل: التكرار والبيانات المتأخرة. التسليم مرة على الأقل من اليوم الرابع يعني أن الحدث نفسه قد يصل مرتين، فكل مسار يحتاج مفتاحًا لإزالة التكرار. والأحداث قد تصل متأخرة: هاتف كان غير متصل، فيصل طلبه إلى المسار بعد يومين من وضعه.
قل: قاعدتان. جمّعوا حسب وقت الحدث (event time)، أي متى حدث، لا وقت المعالجة (processing time)، أي متى وصل. وإلا يتغير إيراد الثلاثاء حسب موعد اتصال الهاتف. وتحسّبوا للبيانات المتأخرة: عندما يصل حدث متأخر أعيدوا تشغيل اليوم الذي ينتمي إليه. ولهذا يجب أن يكون المسار آمنًا عند إعادة التشغيل.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Trusting the data", "Duplicates and late data", false, { sub: "A phone was offline. Its order reaches the pipeline two days after it was placed." });
  s.addShape(pptx.ShapeType.line, { x: 0.8, y: 3.2, w: 11.7, h: 0, line: { color: INK_FAINT, width: 2 } });
  [["Tue", 2.2], ["Wed", 6.0], ["Thu", 9.8]].forEach(([d, x]) => { s.addShape(pptx.ShapeType.line, { x, y: 3.1, w: 0, h: 0.2, line: { color: INK_FAINT, width: 2 } }); label(s, d, x - 0.5, 3.35, 1.0, 0.3, { size: 11, bold: true, color: INK_SOFT }); });
  s.addShape(pptx.ShapeType.ellipse, { x: 2.5, y: 3.0, w: 0.4, h: 0.4, fill: { color: TEAL }, line: { type: "none" } });
  label(s, "order placed", 1.3, 2.35, 2.8, 0.3, { size: 11, bold: true, color: TEAL }); label(s, "EVENT TIME", 1.3, 2.05, 2.8, 0.3, { size: 9, color: TEAL });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.3, y: 3.0, w: 0.4, h: 0.4, fill: { color: RUST }, line: { type: "none" } });
  label(s, "arrives in the pipeline", 8.9, 2.35, 3.2, 0.3, { size: 11, bold: true, color: RUST }); label(s, "PROCESSING TIME", 8.9, 2.05, 3.2, 0.3, { size: 9, color: RUST });
  arrow(s, 3.0, 3.75, 10.2, 3.75, RUST, { w: 1.6, dash: "dash" }); label(s, "2 days late", 5.0, 3.85, 3.2, 0.3, { size: 11, color: RUST });
  infoCard(s, 0.5, 4.5, 6.05, 2.25, TEAL, "Group by event time", "When it happened, not when it arrived. Otherwise Tuesday's revenue changes depending on when the phone reconnects.", 12.5);
  infoCard(s, 6.8, 4.5, 6.05, 2.25, PLUM, "Re-run the day it belongs to", "When a late event arrives, re-run Tuesday. That is why the pipeline must be safe to run again.", 12.5);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Personal data. Copying data into a lake also copies its personal information, names, emails, addresses, into a place with many more readers. A delete me request now has to reach every copy.
SAY: Four habits that help. Copy less: do not send fields nobody will ever ask about. Separate it: keep personal fields in one small, tightly controlled table, and use a customer ID everywhere else. Control access, by column and by role, and log who reads what. And make deletes travel: send delete events, CDC does this, and use a table format that supports row-level deletes, Iceberg or Delta. Keep a record of what was deleted and when.
NOTE: This is engineering guidance, not legal advice. Rules differ by country, and legal advice comes from your legal team.`,
`الوقت: حوالي 3 دقائق.
قل: البيانات الشخصية. نسخ البيانات إلى بحيرة ينسخ معها معلوماتها الشخصية، الأسماء والبريد والعناوين، إلى مكان يقرؤه ناس أكثر بكثير. وطلب "احذفوني" عليه الآن أن يصل إلى كل نسخة.
قل: أربع عادات تساعد. انسخوا أقل: لا ترسلوا حقولًا لن يسأل عنها أحد. افصلوها: احتفظوا بالحقول الشخصية في جدول واحد صغير محكوم الوصول، واستخدموا رقم العميل في كل مكان آخر. تحكموا في الوصول، حسب العمود وحسب الدور، وسجّلوا من يقرأ ماذا. واجعلوا الحذف يسافر: أرسلوا أحداث الحذف، وCDC يفعل ذلك، واستخدموا صيغة جداول تدعم الحذف على مستوى الصف، Iceberg أو Delta. واحتفظوا بسجل لما حُذف ومتى.
ملاحظة: هذه إرشادات هندسية وليست استشارة قانونية. القواعد تختلف حسب البلد، والاستشارة القانونية من فريقكم القانوني.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Trusting the data", "Personal data", false, { sub: "Copying data into a lake also copies its personal information to many more readers." });
  [
    [TEAL, "Copy less", "Do not send fields nobody will ever ask about."],
    [PLUM, "Separate it", "Keep personal fields in one small, tightly controlled table. Use a customer ID everywhere else."],
    [AMBER, "Control access", "By column and by role, and log who reads what."],
    [RUST, "Make deletes travel", "Send delete events (CDC does this). Use a table format with row-level deletes (Iceberg, Delta). Keep a record of what was deleted and when."],
  ].forEach(([c, t, d], i) => {
    const x = 0.5 + (i % 2) * 6.3, y = 1.95 + Math.floor(i / 2) * 2.3;
    infoCard(s, x, y, 6.05, 2.1, c, t, d, 13.5);
  });
  s.addText("Engineering guidance, not legal advice. Rules differ by country. Legal advice comes from your legal team.", { x: 0.5, y: 6.55, w: 12.35, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11.5, color: INK_FAINT });
  footer(s, n, false);
});

// quiz
N(`TIME: about 4 minutes (2 in pairs, 2 to compare).
FACILITATE: Give pairs two minutes. For each of the four incidents, they say what most likely went wrong and which check or habit would have caught it. Take a few answers each. Do not show the answers yet.
SAY: Notice that in every incident nothing crashed. These are silent failures.
NOTE: The answers are on the next slide.`,
`الوقت: حوالي 4 دقائق (دقيقتان في أزواج ودقيقتان للمقارنة).
سهّل النقاش: أعطِ الأزواج دقيقتين. لكل من الحوادث الأربع يقولون ما الذي ساء على الأرجح وأي فحص أو عادة كانت ستلتقطه. خذ بعض الإجابات لكل حادثة. لا تعرض الإجابات بعد.
قل: لاحظوا أنه في كل حادثة لم ينهَر شيء. هذه إخفاقات صامتة.
ملاحظة: الإجابات في الشريحة التالية.`);
add((n) => {
  const s = newSlide(false);
  banner(s, "Quiz · 4 min · in pairs", 0.5, 0.42, 3.0);
  s.addText("What went wrong?", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("For each incident, say what most likely went wrong, and which check or habit would have caught it.", { x: 0.5, y: 1.45, w: 12.3, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  [
    ["Monday: revenue is exactly double", "On Monday the revenue dashboard shows double the usual number. The pipeline failed on Sunday night and was retried."],
    ["3 pm: yesterday's numbers", "At 3 pm the dashboard still shows yesterday's numbers, and nobody was told."],
    ["Deleted, but still in a report", "A customer asked to be deleted last month, and still appears in a report."],
    ["The amount column is empty", "After a release, the \"amount\" column on the dashboard is empty."],
  ].forEach(([t, d], i) => {
    const x = 0.5 + (i % 2) * 6.3, y = 2.05 + Math.floor(i / 2) * 2.4;
    card(s, x, y, 6.05, 2.2, WHITE);
    s.addText([{ text: (i + 1) + " · " + t, options: { fontFace: HEAD, bold: true, fontSize: 15, color: INK, breakLine: true, paraSpaceAfter: 6 } }, { text: d, options: { fontFace: BODY, fontSize: 13, color: INK_SOFT } }], { x: x + 0.25, y: y + 0.15, w: 5.55, h: 1.9, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Quiz: what went wrong?\n\n1 The load appends, so the retry added the day twice. Replace the day instead of appending. A uniqueness or volume check would have stopped it.\n2 No freshness check. The pipeline was stuck or late, and nothing compared the newest data to the clock. Alert when the newest row is older than the limit.\n3 Deletes did not travel. A nightly copy misses deleted rows, and the lake kept the old files. Use CDC delete events and a table format with row-level deletes.\n4 A schema change with no contract. A field was renamed at the source. A completeness check (null count) would have caught it, and a schema check would have stopped the release.\n\nEvery one of these is a silent failure: nothing crashed. That is why the checks are part of the pipeline, not an extra.");
});

N(`TIME: about 3 minutes.
SAY: The answers. One: the load appends, so the retry added the day twice. Replace the day instead, and a uniqueness or volume check would have stopped it. Two: no freshness check; alert when the newest row is older than the limit. Three: deletes did not travel; a nightly copy misses deleted rows, so use CDC delete events and a table format with row-level deletes. Four: a schema change with no contract; a completeness check would have caught it, and a schema check would have stopped the release.
SAY: Every one is a silent failure. That is why the checks are part of the pipeline, not an extra.`,
`الوقت: حوالي 3 دقائق.
قل: الإجابات. واحد: التحميل يضيف، فأضافت إعادة المحاولة اليوم مرتين. استبدلوا اليوم بدل الإضافة، وفحص التفرد أو الحجم كان سيوقفه. اثنان: لا فحص حداثة؛ نبّهوا عندما يكون أحدث صف أقدم من الحد. ثلاثة: الحذف لم يسافر؛ النسخة الليلية تفوّت الصفوف المحذوفة، فاستخدموا أحداث حذف CDC وصيغة جداول بحذف على مستوى الصف. أربعة: تغيير مخطط دون عقد؛ فحص الاكتمال كان سيلتقطه وفحص المخطط كان سيوقف الإصدار.
قل: كل واحدة منها إخفاق صامت. لهذا الفحوص جزء من المسار وليست إضافة.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Quiz answers", "What went wrong?", false, { sub: "In every incident nothing crashed. These are silent failures." });
  tbl(s, 0.5, 1.9, [3.0, 4.6, 4.75], [
    ["Incident", "Most likely cause", "What would have caught it"],
    ["Revenue exactly double", "The load appends, so the retry added the day twice.", "Replace the day. A uniqueness or volume check."],
    ["Yesterday's numbers at 3 pm", "No freshness check. The pipeline was stuck or late.", "Alert when the newest row is older than the limit."],
    ["Deleted, but still in a report", "Deletes did not travel. A nightly copy misses deleted rows.", "CDC delete events, and a table format with row-level deletes."],
    ["The amount column is empty", "A field was renamed at the source, with no contract.", "A completeness check, and a schema check that stops the release."],
  ], { fs: 12.5, rowH: [0.42, 0.95, 0.95, 1.0, 1.0] });
  footer(s, n, false);
});

// ============================================================ HANDS-ON
N(`TIME: 30 seconds.
SAY: Now the exercise. We put the whole session to work on one design.`,
`الوقت: 30 ثانية.
قل: الآن التمرين. نضع الجلسة كلها في تصميم واحد.`);
add((n) => {
  const s = newSlide(true);
  s.addText("B", { x: 0.6, y: 0.9, w: 5, h: 3.5, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 200, color: DARK_WATERMARK });
  s.addText("HANDS-ON  ·  2:00–2:20", { x: 0.7, y: 3.6, w: 11, h: 0.4, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 14, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("From checkout to dashboard", { x: 0.7, y: 4.05, w: 11.8, h: 1.0, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 40, color: WHITE });
  s.addText("Design the path for the revenue dashboard, in pairs.", { x: 0.7, y: 5.1, w: 10.5, h: 1.0, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 16, color: "B9D3D6" });
  footer(s, n, true);
});

N(`TIME: 15 minutes (about 12 in pairs, 3 to compare).
SAY: Exercise B. The product manager wants a dashboard of revenue by customer segment. It must refresh every 15 minutes, keep three years of history, never slow down checkout, and remove a deleted customer's personal data within 30 days. The data starts in the Orders and Customers Postgres databases. Draw the path from the databases to the dashboard.
SAY: Five questions. How does the data leave production, and why not the other three ways? Where does the raw data land, and where is it modelled? How do you re-run a bad day safely? Which two checks stop a bad load? What is your decision about personal data, and how does a delete reach the warehouse?
WATCH FOR: A pair that says "run a nightly query on a replica". It protects checkout, but it fails the 15-minute refresh, the history, and the delete. Ask what each requirement rules out.
NOTE: A sample answer is on the next two slides. Do not show it until pairs have compared.`,
`الوقت: 15 دقيقة (حوالي 12 في أزواج و3 للمقارنة).
قل: التمرين ب. مدير المنتج يريد لوحة إيراد حسب شريحة العملاء. يجب أن تتحدث كل 15 دقيقة، وتحتفظ بثلاث سنوات من التاريخ، ولا تبطئ الدفع أبدًا، وتزيل البيانات الشخصية لعميل محذوف خلال 30 يومًا. البيانات تبدأ في قاعدتي Postgres للطلبات والعملاء. ارسموا المسار من القاعدتين إلى اللوحة.
قل: خمسة أسئلة. كيف تغادر البيانات الإنتاج، ولماذا لا الطرق الثلاث الأخرى؟ أين تهبط البيانات الخام وأين تُصمَّم؟ كيف تعيدون تشغيل يوم سيئ بأمان؟ أي فحصين يوقفان تحميلًا سيئًا؟ وما قراركم بشأن البيانات الشخصية، وكيف يصل الحذف إلى المستودع؟
انتبه: الزوج الذي يقول "استعلام ليلي على نسخة قراءة". يحمي الدفع لكنه يفشل في التحديث كل 15 دقيقة وفي التاريخ وفي الحذف. اسأل ماذا يستبعد كل شرط.
ملاحظة: الإجابة النموذجية في الشريحتين التاليتين. لا تعرضها قبل أن تقارن الأزواج.`);
add((n) => {
  const s = newSlide(false);
  banner(s, "Exercise B · 15 min · in pairs", 0.5, 0.42, 3.4);
  s.addText("Design the path", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("The product manager wants a dashboard of revenue by customer segment. The data starts in the Orders and Customers Postgres databases.", { x: 0.5, y: 1.5, w: 12.3, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  [["Refresh every 15 minutes", TEAL], ["3 years of history", PLUM], ["Never slow checkout", RUST], ["Personal data gone in 30 days", AMBER]].forEach(([t, c], i) => {
    const x = 0.5 + i * 3.15;
    box(s, x, 2.15, 3.0, 0.7, WHITE, c, 2); label(s, t, x, 2.15, 3.0, 0.7, { size: 11, bold: true, color: c });
  });
  card(s, 0.5, 3.15, 7.4, 3.65, WHITE);
  qBlock(s, 0.75, 3.3, 6.9, 3.4, "Your design", ["How does the data leave the production databases? Why not the other three ways?", "Where does the raw data land, and where is it modelled? Name the fact table and the dimensions.", "How do you re-run a bad day safely?", "Which two checks stop a bad load before the dashboard sees it?", "What is your decision about personal data, and how does a delete reach the warehouse?"], 12.5);
  card(s, 8.15, 3.15, 4.7, 3.65, WHITE);
  s.addText([{ text: "Done when", options: { fontFace: HEAD, fontSize: 15, bold: true, color: TEAL, breakLine: true, paraSpaceAfter: 6 } }].concat(bl(["Checkout never runs an extra query for the pipeline.", "Running any step twice gives the same result.", "You can say how fresh the dashboard is, and how you would know if it was late."], { fs: 12.5, color: INK, gap: 8 })), { x: 8.4, y: 3.3, w: 4.25, h: 3.4, isTextBox: true, margin: 0, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: One good answer, as a picture. The Orders and Customers databases feed a CDC tool that reads the Postgres change log, so checkout runs no extra query. It publishes to Kafka, one topic per table. The raw layer takes it append only, partitioned by day. The cleaned layer dedupes by event ID. The modelled layer has the fact table and dimensions. The dashboard reads it, refreshed every 15 minutes.
SAY: Two things under the line. After every load, checks stop the run: freshness, uniqueness, and row count against the source. And personal data sits in one small access-controlled table, so the lake and the facts carry only the customer ID.`,
`الوقت: حوالي 3 دقائق.
قل: إجابة جيدة واحدة كصورة. قاعدتا الطلبات والعملاء تغذيان أداة CDC تقرأ سجل تغييرات Postgres، فلا يشغّل الدفع أي استعلام إضافي. تنشر إلى Kafka، موضوع لكل جدول. الطبقة الخام تستقبلها إضافة فقط مقسّمة حسب اليوم. الطبقة المنظفة تزيل التكرار برقم الحدث. الطبقة المصممة فيها جدول الحقائق والأبعاد. واللوحة تقرأ منها وتتحدث كل 15 دقيقة.
قل: شيئان تحت الخط. بعد كل تحميل تُوقف الفحوص التشغيل: الحداثة والتفرد وعدد الصفوف مقارنة بالمصدر. والبيانات الشخصية في جدول صغير واحد محكوم الوصول، فالبحيرة والحقائق تحمل رقم العميل فقط.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Exercise B · A sample answer", "One good answer, as a picture", false, { sub: "There is no single right answer. Judge by the reasons." });
  const w = 1.85, gap = (12.35 - 6 * w) / 5, X = (i) => 0.5 + i * (w + gap);
  [
    ["Orders + Customers", "Postgres", WHITE, INK_SOFT],
    ["CDC tool", "reads the log", AMBER_TINT, AMBER],
    ["Kafka", "topic per table", PLUM_TINT, PLUM],
    ["Raw", "append, by day", AMBER_TINT, AMBER],
    ["Cleaned", "dedupe by event_id", TEAL_TINT, TEAL],
    ["Modelled", "fact + dimensions", PLUM_TINT, PLUM],
  ].forEach(([t, sub, f, c], i) => {
    box(s, X(i), 1.95, w, 1.2, f, c, i === 5 ? 2.5 : 1.5);
    s.addText([{ text: t, options: { fontFace: HEAD, bold: true, fontSize: 12, color: INK, breakLine: true } }, { text: sub, options: { fontFace: BODY, fontSize: 10, color: INK_SOFT } }], { x: X(i) + 0.05, y: 1.95, w: w - 0.1, h: 1.2, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    if (i < 5) arrow(s, X(i) + w + 0.03, 2.55, X(i + 1) - 0.03, 2.55, INK_SOFT, { w: 1.4 });
  });
  box(s, X(5), 3.85, w, 1.0, WHITE, INK_SOFT, 1.5); label(s, "Dashboard\nevery 15 min", X(5), 3.85, w, 1.0, { size: 10, bold: true, font: HEAD });
  arrow(s, X(5) + w / 2, 3.15, X(5) + w / 2, 3.83, INK_SOFT, { w: 1.4 });
  box(s, X(4), 3.85, w, 1.0, RUST_TINT, RUST, 1.5); s.addText([{ text: "PII table", options: { fontFace: HEAD, bold: true, fontSize: 11.5, color: RUST, breakLine: true } }, { text: "name, email, access-controlled", options: { fontFace: BODY, fontSize: 9.5, color: INK_SOFT } }], { x: X(4) + 0.05, y: 3.85, w: w - 0.1, h: 1.0, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, X(4) + w / 2, 3.15, X(4) + w / 2, 3.83, RUST, { w: 1.4, dash: "dash" });
  card(s, 0.5, 3.85, X(3) + w - 0.5, 1.0, WHITE, { shadow: false, line: LINE });
  s.addText("Checkout runs no extra query. The CDC tool reads the Postgres change log.", { x: 0.7, y: 3.85, w: X(3) + w - 0.9, h: 1.0, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT });
  card(s, 0.5, 5.2, 12.35, 1.5, WHITE, { shadow: false, line: PLUM });
  s.addText([{ text: "After every load, checks stop the run:  ", options: { bold: true, color: PLUM, fontFace: HEAD, fontSize: 14 } }, { text: "freshness (the newest row is under 30 minutes old), uniqueness of order_id, and the row count against the source (within about 1%). Every load replaces its day, so any step can be run again.", options: { color: INK_SOFT, fontFace: BODY, fontSize: 13 } }], { x: 0.8, y: 5.2, w: 11.8, h: 1.5, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: The five answers, in words. How data leaves: change data capture. A nightly copy fails the 15-minute refresh, misses deletes, and loads production. A timestamp copy misses deletes too, and trusts updated_at. Where it lands: raw, append only, partitioned by date; then cleaned and deduplicated by event ID; then modelled: a fact table for orders with dimensions for customer, date and product, with the customer dimension type 2 so the history stays right.
SAY: Re-running a bad day: every load replaces the partition for its day, or merges on the event ID. Two checks: freshness, and uniqueness of order ID. Row count against the source is a strong third. Personal data: the lake and the facts carry only the customer ID. Name and email live in one small, access-controlled table. A delete in the source arrives as a CDC delete event and removes the person there, and a table format with row-level deletes removes them from the history within the 30 days. The dashboard only ever sees the segment.`,
`الوقت: حوالي 3 دقائق.
قل: الإجابات الخمس بالكلمات. كيف تغادر البيانات: التقاط التغييرات (CDC). النسخة الليلية تفشل في التحديث كل 15 دقيقة وتفوّت الحذف وتحمّل الإنتاج. والنسخ بالوقت يفوّت الحذف أيضًا ويثق بـ updated_at. أين تهبط: خام، إضافة فقط، مقسّم حسب التاريخ؛ ثم منظف بلا تكرار برقم الحدث؛ ثم مصمم: جدول حقائق للطلبات وأبعاد للعميل والتاريخ والمنتج، وبُعد العميل من النوع الثاني ليبقى التاريخ صحيحًا.
قل: إعادة تشغيل يوم سيئ: كل تحميل يستبدل قسم يومه، أو يدمج برقم الحدث. فحصان: الحداثة، وتفرد رقم الطلب. وعدد الصفوف مقابل المصدر ثالث قوي. البيانات الشخصية: البحيرة والحقائق تحمل رقم العميل فقط. الاسم والبريد في جدول صغير واحد محكوم الوصول. والحذف في المصدر يصل كحدث حذف CDC فيزيل الشخص من هناك، وصيغة جداول بحذف على مستوى الصف تزيله من التاريخ خلال الثلاثين يومًا. واللوحة لا ترى سوى الشريحة.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Exercise B · A sample answer", "The five answers, in words", false);
  tbl(s, 0.5, 1.8, [2.6, 9.75], [
    ["Question", "A good answer"],
    ["How data leaves", "Change data capture. A nightly copy fails the 15-minute refresh, misses deletes, and loads production. A timestamp copy misses deletes too."],
    ["Where it lands", "Raw (append only, by date), then cleaned (deduplicated by event ID), then modelled: fact_orders with dim_customer (type 2), dim_date, dim_product."],
    ["Re-running a bad day", "Every load replaces the partition for its day, or merges on the event ID. Late events re-run their own day."],
    ["Two checks", "Freshness, and uniqueness of order_id. Row count against the source is a strong third."],
    ["Personal data", "The lake and facts carry only customer_id. Name and email sit in one access-controlled table. Deletes arrive as CDC events, and row-level deletes clean the history within 30 days."],
  ], { fs: 12, rowH: [0.42, 0.95, 0.95, 0.8, 0.7, 1.1] });
  footer(s, n, false);
});

// ---- review + closing
N(`TIME: 8 minutes.
SAY: Pick the question that matches who is in the room, and let people answer in their own words. Everyone: pick one number your team reports, and trace it back to its source, step by step. Where would you not trust it? Tech leads: someone asks for a quick number and proposes a query on production. What do you offer instead, today? QEs: write the three checks you would add to the most important table in your warehouse, and say how a failure reaches a person. AI engineers: if you asked an agent to get the numbers for the board, what four questions should it ask before it runs any query?`,
`الوقت: 8 دقائق.
قل: اختاروا السؤال الذي يناسب من في القاعة ودعوا الناس يجيبون بكلماتهم. الجميع: اختاروا رقمًا يُبلغ عنه فريقكم وتتبعوه إلى مصدره خطوة بخطوة. أين لن تثقوا به؟ القادة التقنيون: أحدهم يطلب رقمًا سريعًا ويقترح استعلامًا على الإنتاج. ماذا تعرضون بدلًا منه اليوم؟ مهندسو الجودة: اكتبوا الفحوص الثلاثة التي ستضيفونها لأهم جدول في مستودعكم، وقولوا كيف يصل الإخفاق إلى شخص. مهندسو الذكاء الاصطناعي: لو طلبتم من وكيل "احضر أرقام مجلس الإدارة"، ما الأسئلة الأربعة التي يجب أن يطرحها قبل أن يشغّل أي استعلام؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "2:20–2:30 · Review", "Group critique", false);
  const qs = [
    ["JR", TEAL, "Everyone", "Pick one number your team reports. Can you trace it back to its source, step by step? Where would you not trust it?"],
    ["SR", SLATE, "Tech Leads", "Someone asks for \"a quick number\" and proposes a query on production. What do you offer them instead, today, and what would it take to build?"],
    ["QE", RUST, "QEs", "Write the three checks you would add to the most important table in your warehouse. How would a failure reach a person?"],
    ["AI", PLUM, "AI Engineers", "If you asked an agent to \"get the numbers for the board\", what four questions should it ask before it runs any query?"],
  ];
  const cw = 5.85, ch = 2.05, gx = 0.3, gy = 0.3, ox = 0.5, oy = 1.85;
  qs.forEach((q, i) => {
    const x = ox + (i % 2) * (cw + gx), y = oy + Math.floor(i / 2) * (ch + gy);
    card(s, x, y, cw, ch, WHITE);
    circle(s, x + 0.3, y + 0.3, 0.55, q[1], q[0], { fontSize: 12 });
    s.addText(q[2], { x: x + 1.05, y: y + 0.34, w: cw - 1.3, h: 0.4, isTextBox: true, margin: 0, valign: "middle", fontFace: CODE, bold: true, fontSize: 12, color: q[1], charSpacing: 0.5 });
    s.addText(q[3], { x: x + 0.3, y: y + 1.0, w: cw - 0.6, h: 0.95, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: INK, valign: "top" });
  });
  footer(s, n, false);
});

N(`TIME: 3 minutes.
SAY: Before you go, pick one and do it this week. Trace one real report back to its source and write down every step it takes. Find a pipeline that appends, and change one step to replace the day. Add a freshness check to the dashboard people open most. List the personal fields that reach your warehouse, and where a delete would have to go.
SAY: And that is the course. Eight sessions. Readable code on Days 1 to 3. What happens at scale on Day 4. How to structure it on Day 7, and how to get trustworthy answers out of it today. Thank you.
ASK: One sentence each: the one thing you will do differently on Monday.`,
`الوقت: 3 دقائق.
قل: قبل أن تغادروا، اختاروا واحدًا ونفذوه هذا الأسبوع. تتبعوا تقريرًا حقيقيًا إلى مصدره واكتبوا كل خطوة يمر بها. ابحثوا عن مسار يضيف، وغيّروا خطوة فيه لتستبدل اليوم. أضيفوا فحص حداثة إلى اللوحة التي يفتحها الناس أكثر. اذكروا الحقول الشخصية التي تصل إلى مستودعكم، وأين يجب أن يصل الحذف.
قل: وهذه هي الدورة. ثماني جلسات. كود مقروء في الأيام من 1 إلى 3. ما يحدث عند التوسع في اليوم الرابع. كيف نبنيه في اليوم السابع، وكيف نستخرج منه إجابات موثوقة اليوم. شكرًا لكم.
اسأل: جملة واحدة من كل شخص: الشيء الوحيد الذي ستفعلونه بشكل مختلف يوم الاثنين.`);
add((n) => {
  const s = newSlide(true);
  s.addText("BEFORE YOU GO", { x: 0.7, y: 0.55, w: 8, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Pick one. Do it this week.", { x: 0.7, y: 0.9, w: 11, h: 0.8, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 34, color: WHITE });
  [
    "Trace one real report back to its source, and write down every step it takes",
    "Find a pipeline that appends, and change one step to replace the day",
    "Add a freshness check to the dashboard people open most",
    "List the personal fields that reach your warehouse, and where a delete would have to go",
  ].forEach((it, i) => {
    const y = 1.95 + i * 0.85;
    card(s, 0.7, y, 11.9, 0.68, DARK2, { shadow: false });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: y + 0.19, w: 0.3, h: 0.3, rectRadius: 0.04, fill: { type: "none" }, line: { color: TEAL, width: 1.5 } });
    s.addText(it, { x: 1.45, y, w: 10.9, h: 0.68, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 14, color: WHITE });
  });
  s.addShape(pptx.ShapeType.line, { x: 0.7, y: 5.6, w: 11.9, h: 0, line: { color: "234750", width: 1 } });
  s.addText("COURSE COMPLETE", { x: 0.7, y: 5.8, w: 6, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  s.addText("Eight sessions. Clean code, to scale.", { x: 0.7, y: 6.1, w: 11, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 22, color: WHITE });
  footer(s, n, true, "Clean Code to Scale");
});

// ============================================================ build
TOTAL = queue.length;
queue.forEach((fn, i) => fn(i + 1));
let missing = 0;
queueNotes.forEach((nn, i) => { if (!nn[0] || !nn[1]) { missing++; console.warn("slide " + (i + 1) + " has no notes: " + (allSlides[i]._title || "?")); } });
allSlides.forEach((sl, i) => {
  const [en, ar] = queueNotes[i] || [null, null];
  const parts = [];
  if (en) parts.push(en);
  if (ar) parts.push("──────── العربية (Arabic) ────────\n" + ar);
  if (sl._notes.length) parts.push("MORE (answers and extra points):\n" + sl._notes.join("\n\n"));
  if (parts.length) sl._applyNotes(parts.join("\n\n"));
});
pptx.writeFile({ fileName: "session-08.pptx" }).then(() => {
  try {
    require("child_process").execFileSync("python3", [__dirname + "/fix_rtl_notes.py", "session-08.pptx"], { stdio: "inherit" });
  } catch (e) { console.warn("RTL post-process skipped:", e.message); }
  console.log("written", TOTAL, "slides;", missing, "without notes");
});
