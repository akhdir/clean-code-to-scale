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
const FOOT = "Day 07 · Clean Code to Scale";

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
const FOOT_UNUSED = null;

// ---- title
N(`TIME: 2 minutes.
SAY: Welcome to Day 7. Day 4 showed what happens when many machines work together: calls get slow, fail, repeat and return old data. Today we ask the question that comes next: should the code be one program or many? We use the same order system as before.
ASK: Who works in a monolith, who in microservices, and who in something in between?`,
`الوقت: دقيقتان.
قل: مرحبًا بكم في اليوم السابع. أظهر اليوم الرابع ما يحدث عندما تعمل أجهزة كثيرة معًا: الاستدعاءات تبطؤ وتفشل وتتكرر وتعيد بيانات قديمة. اليوم نسأل السؤال التالي: هل يجب أن يكون الكود برنامجًا واحدًا أم برامج كثيرة؟ نستخدم نظام الطلبات نفسه.
اسأل: من يعمل في نظام واحد متكامل (monolith)، ومن في خدمات مصغرة (microservices)، ومن في شيء بينهما؟`);
add((n) => {
  const s = newSlide(true);
  s.addText("class PlaceOrder:\n    def __init__(self, inventory):\n        self.inventory = inventory\n\n    # a plain call,\n    # or a network call?", { x: 6.4, y: 0.5, w: 6.7, h: 2.6, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 16, color: DARK_WATERMARK, lineSpacingMultiple: 1.3 });
  s.addText("CLEAN CODE TO SCALE   ·   DAY 07 OF 08", { x: 0.7, y: 2.3, w: 10, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Microservices vs Monolith", { x: 0.7, y: 2.75, w: 11.8, h: 1.2, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 44, color: WHITE });
  s.addText("One program or many? What you gain, what you pay, where to draw the lines, how the pieces work together, and how to move safely in either direction.", { x: 0.7, y: 4.05, w: 10, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 17, color: "B9D3D6", valign: "top" });
  [["2:30", "DURATION"], ["PYTHON", "CODE SAMPLES"], ["PAIRS", "BOUNDARIES + 3 QUIZZES"]].forEach((st, i) => {
    const x = 0.7 + i * 3.1;
    s.addText(st[0], { x, y: 5.65, w: 2.9, h: 0.45, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 19, color: WHITE });
    s.addText(st[1], { x, y: 6.13, w: 2.9, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 9.5, color: "6C8B90", charSpacing: 1 });
  });
  footer(s, n, true, "Clean Code to Scale");
});

// ---- agenda
N(`TIME: 2 minutes.
SAY: Six blocks. An opening story. Part 1: what they are, what changes when a call becomes a network call, and three deep dives. Part 2: start with a modular monolith. Part 3: where to draw the lines, with an exercise in pairs. Part 4: how services work together: events, sagas, the outbox. Part 5: moving out safely, then a group critique.
NOTE: Three short check-ins and the quizzes are woven in. If time is short, run them as a quick show of hands.`,
`الوقت: دقيقتان.
قل: ست فقرات. قصة افتتاحية. الجزء الأول: ما هي هذه التصميمات، وماذا يتغير عندما يصبح الاستدعاء استدعاءً عبر الشبكة، مع ثلاث وقفات معمقة. الجزء الثاني: ابدأ بنظام واحد مقسّم إلى وحدات (modular monolith). الجزء الثالث: أين نرسم الحدود، مع تمرين في أزواج. الجزء الرابع: كيف تعمل الخدمات معًا: الأحداث والـ saga والـ outbox. الجزء الخامس: الانتقال بأمان، ثم نقاش جماعي.
ملاحظة: هناك ثلاث وقفات قصيرة واختبارات موزعة. إذا ضاق الوقت فنفذها برفع الأيدي.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Session Map", "Five parts, and a story to start", false);
  const steps = [
    ["01", "0:00–0:10", "OPENING", INK_FAINT, "A team of eight and fourteen services."],
    ["02", "0:10–0:35", "PART 1", TEAL, "What they are, what changes, three deep dives."],
    ["03", "0:35–1:00", "PART 2", SLATE, "Start with a modular monolith."],
    ["04", "1:00–1:30", "PART 3", PLUM, "Where to draw the lines. Exercise A."],
    ["05", "1:30–2:05", "PART 4", RUST, "Working together: events, sagas, the outbox."],
    ["06", "2:05–2:30", "PART 5 + REVIEW", AMBER, "Moving out safely, then group critique."],
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
  s.addText("Three short check-ins, an exercise in pairs and a quiz are woven through the parts.", { x: 0.5, y: 6.0, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  s.addText("Short on time? Run the check-ins as a quick show of hands, then make the strangler section a handout.", { x: 0.5, y: 6.4, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  footer(s, n, false);
});

// ---- objectives
N(`TIME: 1 minute.
SAY: Four things to leave with. Know the three designs and what each costs. Know what changes when a function call becomes a network call. Be able to draw boundaries around data and change, and spot a distributed monolith. And know how to keep data correct across services, and how to move out in small, safe steps.`,
`الوقت: دقيقة واحدة.
قل: أربعة أشياء تخرجون بها. معرفة التصميمات الثلاثة وكلفة كل منها. معرفة ما يتغير عندما يصبح استدعاء الدالة استدعاءً عبر الشبكة. القدرة على رسم الحدود حول البيانات والتغيير، وكشف النظام الموزع المتشابك (distributed monolith). ومعرفة كيف نحافظ على صحة البيانات بين الخدمات، وكيف ننتقل بخطوات صغيرة آمنة.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Objectives", "What today is for", false);
  [
    ["Know the three designs and what each one costs", "Monolith, modular monolith and microservices: what you gain and what you pay."],
    ["Know what changes when a call crosses the network", "Speed, failure, repeats, transactions, data and debugging. The Day 4 problems come back."],
    ["Draw boundaries around data and change", "And spot a distributed monolith before you build one."],
    ["Keep data correct, and move in small safe steps", "Events, sagas and the outbox. The strangler fig for moving out of a monolith."],
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
SAY: Why this is for you. Junior and mid engineers: you will work inside one of these designs from day one, and today you learn why it is split the way it is. Tech leads: "should we split this?" lands on your desk. You leave with a checklist. QEs: more services mean more contracts that can break. AI engineers: an agent asked to add a service will happily create one. Today gives you the questions to ask first.`,
`الوقت: دقيقة واحدة.
قل: لماذا هذا مهم لكم. المهندسون المبتدئون والمتوسطون: ستعملون داخل أحد هذه التصميمات من اليوم الأول، واليوم تفهمون لماذا قُسّم هكذا. القادة التقنيون: سؤال "هل نقسّم هذا؟" يصلكم أنتم، وتخرجون بقائمة تحقق. مهندسو الجودة: خدمات أكثر تعني عقودًا أكثر قد تنكسر. مهندسو الذكاء الاصطناعي: الوكيل الذي يُطلب منه إضافة خدمة سينشئها بسرور. اليوم تحصلون على الأسئلة التي تُطرح أولًا.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Room", "Why this session, for you specifically", false);
  const roles = [
    ["JR", TEAL, "Junior – Mid Engineers", "You will work inside one of these designs from your first day. Today you learn why the code is split the way it is, and how to ask for a better boundary."],
    ["SR", SLATE, "Senior Eng & Tech Leads", "\"Should we split this?\" lands on your desk. You leave with a checklist and the words to say yes or no, with reasons."],
    ["QE", RUST, "QEs", "More services mean more contracts that can break and more places a bug can hide. You learn what to test at each boundary."],
    ["AI", PLUM, "AI Engineers", "An agent asked to \"add a service\" will happily create one. You get the questions to ask first, and the rule that a service owns its data."],
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

// ---- opening story
N(`TIME: about 4 minutes.
SAY: A team of eight has a working monolith. It is a bit slow to release, so they split it into fourteen services. Six months later: releases take longer, because three services must go out together. One order screen needs nine calls, and one is always slow. A bug crosses three repositories, and nobody can run the whole system on a laptop.
SAY: This is not a story that says microservices are bad. They solve real problems for the right team at the right size. The question is when, and how much.
ASK: Take two or three stories about your worst deploy or worst bug. We will come back to them.`,
`الوقت: حوالي 4 دقائق.
قل: فريق من ثمانية لديه نظام واحد يعمل. إصداره بطيء قليلًا، فقسّموه إلى أربع عشرة خدمة. بعد ستة أشهر: الإصدارات صارت أطول لأن ثلاث خدمات يجب أن تُنشر معًا. شاشة طلب واحدة تحتاج تسعة استدعاءات وأحدها بطيء دائمًا. والخطأ الواحد يعبر ثلاثة مستودعات، ولا أحد يستطيع تشغيل النظام كله على حاسوب محمول.
قل: هذه ليست قصة تقول إن الخدمات المصغرة سيئة. هي تحل مشاكل حقيقية للفريق المناسب بالحجم المناسب. السؤال هو متى وكم.
اسأل: خذوا قصتين أو ثلاثًا عن أسوأ نشر أو أسوأ خطأ مررتم به. سنعود إليها.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Opening", "A team of eight and fourteen services", false, { sub: "They split a working monolith because releases were a bit slow. Six months later:" });
  card(s, 0.5, 1.85, 5.6, 4.9, WHITE, { shadow: false, line: LINE });
  label(s, "WHAT THEY DID", 0.7, 1.95, 5.2, 0.28, { size: 10, bold: true, color: TEAL, align: "left" });
  box(s, 0.8, 2.9, 1.8, 1.4, TEAL_TINT, TEAL, 2); label(s, "1 monolith", 0.8, 2.95, 1.8, 0.6, { size: 12, bold: true }); label(s, "8 engineers", 0.8, 3.5, 1.8, 0.4, { size: 9.5, color: INK_SOFT });
  arrow(s, 2.7, 3.6, 3.45, 3.6, INK_SOFT, { w: 1.6 });
  for (let k = 0; k < 14; k++) { const cx = 3.6 + (k % 4) * 0.62, cy = 2.5 + Math.floor(k / 4) * 0.62; box(s, cx, cy, 0.5, 0.45, PLUM_TINT, PLUM, 1.25); }
  label(s, "14 services", 3.6, 5.0, 2.4, 0.3, { size: 11, bold: true, color: PLUM });
  s.addText("Releases were a bit slow, so they split.", { x: 0.7, y: 5.6, w: 5.2, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  infoCard(s, 6.35, 1.85, 6.5, 1.45, RUST, "Releases got slower", "Three services must go out together, in a fixed order.", 12.5);
  infoCard(s, 6.35, 3.45, 6.5, 1.45, RUST, "One screen, nine calls", "Showing one order needs nine calls, and one of them is always slow.", 12.5);
  infoCard(s, 6.35, 5.05, 6.5, 1.7, RUST, "Bugs cross repositories", "A bug crosses three repositories, and nobody can run the whole system on a laptop.", 12.5);
  footer(s, n, false);
});

// ---- three questions
N(`TIME: 1 minute.
SAY: Three questions for today. What exactly do we gain, and what exactly do we pay? Where do we draw the lines, so that a split helps instead of hurts? And how do the pieces work together, and how do we move without breaking things?`,
`الوقت: دقيقة واحدة.
قل: ثلاثة أسئلة لليوم. ماذا نكسب بالضبط وماذا ندفع بالضبط؟ أين نرسم الحدود لتفيدنا القسمة بدل أن تضرنا؟ وكيف تعمل القطع معًا، وكيف ننتقل دون أن نكسر شيئًا؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Opening", "Three questions for today", false);
  [
    ["1", TEAL, "What do we gain, and what do we pay?", "A clear price list, so a split is a choice and not a fashion."],
    ["2", PLUM, "Where do we draw the lines?", "So that a split helps instead of hurts."],
    ["3", RUST, "How do the pieces work together, and how do we move?", "Events, sagas, the outbox, and a safe way out of a monolith."],
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
SAY: Part 1. What they are, what changes, and three deep dives into the things that surprise teams most: transactions, data and debugging.`,
`الوقت: 30 ثانية.
قل: الجزء الأول. ما هي هذه التصميمات، وماذا يتغير، وثلاث وقفات معمقة في أكثر ما يفاجئ الفرق: المعاملات والبيانات وتتبع الأخطاء.`);
divider(1, "0:10–0:35", "What they are, and what changes", "Three designs in plain words, the price of a network call, and three deep dives.");

N(`TIME: about 3 minutes.
SAY: Three words in plain language. Monolith: one program, deployed as one unit, usually one database. Modular monolith: still one deployable, but inside it the code is split into modules with clear walls, and each module owns its own tables. Microservices: many small programs, deployed on their own, talking over the network, and each owns its data. Nobody else reads it.
ASK: Which of these is closest to what you work in today?`,
`الوقت: حوالي 3 دقائق.
قل: ثلاث كلمات بلغة بسيطة. النظام الواحد المتكامل (monolith): برنامج واحد يُنشر كوحدة واحدة، وعادة قاعدة بيانات واحدة. النظام الواحد المقسّم إلى وحدات (modular monolith): ما زال قابلًا للنشر كوحدة واحدة، لكنه من الداخل مقسّم إلى وحدات بجدران واضحة، وكل وحدة تملك جداولها. الخدمات المصغرة (microservices): برامج صغيرة كثيرة تُنشر كلٌّ على حدة وتتحدث عبر الشبكة، وكل واحدة تملك بياناتها ولا يقرؤها غيرها.
اسأل: أيها أقرب إلى ما تعملون عليه اليوم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · The basics", "Three words, in plain language", false);
  tbl(s, 0.5, 1.9, [2.6, 6.2, 3.55], [
    ["Word", "What it means", "Data"],
    ["Monolith", "One program, deployed as one unit. Parts of it call each other as normal function calls.", "Usually one database"],
    ["Modular monolith", "Still one deployable, but inside it the code is split into modules with clear walls and public APIs.", "One database, but each module owns its own tables"],
    ["Microservices", "Many small programs, each deployed on its own, talking to each other over the network.", "Each service owns its data. Nobody else reads it"],
  ], { fs: 14, rowH: [0.5, 1.2, 1.3, 1.3] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Here it is as a picture. On the left, one deployable with four modules that call each other as functions, and one database. On the right, four separate services, each with its own database, and every arrow between them is a network call.
SAY: The difference is not the number of boxes. It is who owns the data, and what kind of call sits between the boxes.`,
`الوقت: حوالي 3 دقائق.
قل: هنا الفكرة كصورة. على اليسار، وحدة نشر واحدة فيها أربع وحدات تستدعي بعضها كدوال، وقاعدة بيانات واحدة. على اليمين أربع خدمات منفصلة، لكل منها قاعدة بياناتها، وكل سهم بينها هو استدعاء عبر الشبكة.
قل: الفرق ليس عدد الصناديق. الفرق هو من يملك البيانات، ونوع الاستدعاء بين الصناديق.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · The basics", "One deployable, or many", false);
  label(s, "MONOLITH: one deployable", 0.5, 1.8, 5.4, 0.3, { size: 11, bold: true, color: TEAL });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 2.15, w: 5.4, h: 3.1, rectRadius: 0.1, fill: { type: "none" }, line: { color: TEAL, width: 2 } });
  [["Orders", 0.8, 2.45], ["Inventory", 3.3, 2.45], ["Payments", 0.8, 3.75], ["Notifications", 3.3, 3.75]].forEach(([t, x, y]) => { box(s, x, y, 2.4, 0.95, TEAL_TINT, TEAL, 1.5); label(s, t, x, y, 2.4, 0.95, { size: 11.5, bold: true }); });
  label(s, "function calls, same process", 0.8, 4.75, 4.8, 0.3, { size: 9, color: INK_FAINT });
  arrow(s, 3.2, 5.25, 3.2, 5.6, INK_SOFT, { w: 1.4 });
  box(s, 2.1, 5.65, 2.2, 0.7, WHITE, INK_SOFT, 1.5); label(s, "one database", 2.1, 5.65, 2.2, 0.7, { size: 11 });
  label(s, "MICROSERVICES: many deployables", 6.5, 1.8, 6.35, 0.3, { size: 11, bold: true, color: PLUM });
  [["Orders", 6.7, 2.2], ["Inventory", 10.2, 2.2], ["Payments", 6.7, 4.4], ["Notifications", 10.2, 4.4]].forEach(([t, x, y]) => {
    box(s, x, y, 2.5, 0.95, PLUM_TINT, PLUM, 1.5); label(s, t, x, y, 2.5, 0.95, { size: 11.5, bold: true });
    box(s, x + 0.1, y + 1.15, 1.5, 0.5, WHITE, INK_SOFT, 1.1); label(s, t.toLowerCase() + " DB", x + 0.1, y + 1.15, 1.5, 0.5, { size: 8.5 });
    arrow(s, x + 0.85, y + 0.95, x + 0.85, y + 1.15, INK_SOFT, { w: 1, noHead: true });
  });
  arrow(s, 9.25, 2.68, 10.15, 2.68, RUST, { w: 1.8, dash: "dash" }); label(s, "network call", 9.1, 2.3, 1.2, 0.3, { size: 8.5, color: RUST, bold: true });
  arrow(s, 9.0, 3.2, 9.0, 4.35, RUST, { w: 1.8, dash: "dash" });
  arrow(s, 12.4, 3.2, 12.4, 4.35, RUST, { w: 1.8, dash: "dash" });
  s.addText("In the monolith the database is shared and calls are in-process. In microservices each service owns its data, and every arrow between services is a network call that can be slow, fail, or arrive twice.", { x: 6.5, y: 6.35, w: 6.35, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 10.5, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: What you may gain: independent deploys, scaling one part, isolated failure, team ownership, technology choice. What you pay for each: many deploys to manage, network calls, no shared transaction, harder to see what is happening, and more to know and run.
SAY: Every gain on the left has a price on the right. The question is whether you need the gain badly enough to pay.`,
`الوقت: حوالي 3 دقائق.
قل: ما قد تكسبه: نشر مستقل، وتوسيع جزء واحد، وعزل الفشل، وملكية الفريق، وحرية اختيار التقنية. وما تدفعه مقابل كل واحد: نشرات كثيرة لإدارتها، واستدعاءات عبر الشبكة، وعدم وجود معاملة مشتركة، وصعوبة رؤية ما يجري، وأشياء أكثر لتعرفها وتشغّلها.
قل: لكل مكسب على اليسار ثمن على اليمين. السؤال: هل تحتاجون المكسب بما يكفي لتدفعوا الثمن؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · The basics", "What you gain and what you pay", false);
  tbl(s, 0.5, 1.9, [6.1, 6.25], [
    ["You may gain…", "…and you pay"],
    ["Independent deploys. A team ships its service without waiting for the others.", "Many deploys to manage: pipelines, versions, and the order in which things go out."],
    ["Scale one part. Search gets 20 machines for the sale, the rest stays small.", "Network calls. Slower, and they fail in new ways (Day 4)."],
    ["Isolated failure. A broken notification service does not have to stop checkout.", "No shared transaction. Data across services is only eventually consistent."],
    ["Team ownership. One team, one service, clear responsibility.", "Harder to see. Logs, metrics and traces across many services, or nobody knows why it is slow."],
    ["Technology choice. The right database or language for each part.", "More to know and run. Many stacks, and a harder test setup."],
  ], { fs: 12.5, rowH: [0.45, 0.8, 0.8, 0.8, 0.8, 0.8] });
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: This table is the heart of the session. Every arrow between services is a call that used to be a line of code. Speed: nanoseconds become milliseconds. Failure: the whole program fails together, versus the other side being slow, down or unreachable. Repeats: a timeout and retry can run it twice. Transactions: one transaction versus none. Versions: compiled together versus deployed apart. Data: a join versus an API call. Debugging: one stack trace versus a trace across services.
NOTE: Each row is a problem from Day 4. The next three slides take transactions, data and debugging in depth.`,
`الوقت: حوالي 4 دقائق.
قل: هذا الجدول قلب الجلسة. كل سهم بين الخدمات هو استدعاء كان سطر كود. السرعة: النانو ثانية تصير ميلي ثانية. الفشل: البرنامج كله يفشل معًا، مقابل أن يكون الطرف الآخر بطيئًا أو متوقفًا أو غير قابل للوصول. التكرار: مهلة ثم إعادة محاولة قد تنفذه مرتين. المعاملات: معاملة واحدة مقابل لا شيء. الإصدارات: تُبنى معًا مقابل تُنشر منفصلة. البيانات: join مقابل استدعاء API. تتبع الأخطاء: أثر واحد للأخطاء مقابل أثر عبر الخدمات.
ملاحظة: كل صف مشكلة من اليوم الرابع. الشرائح الثلاث التالية تفصّل المعاملات والبيانات وتتبع الأخطاء.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · What changes", "A function call becomes a network call", false, { sub: "Every arrow between services is a call that used to be a line of code." });
  tbl(s, 0.5, 1.85, [1.8, 4.7, 5.85], [
    ["", "Function call (in a monolith)", "Network call (between services)"],
    ["Speed", "Nanoseconds", "Milliseconds, and sometimes much slower"],
    ["Failure", "Rare. The whole program fails together", "Often, and partly: the other side may be slow, down, or unreachable"],
    ["Repeats", "Called once", "A timeout and retry can run it twice, so it must be safe to repeat"],
    ["Transactions", "One transaction covers it all", "No shared transaction. Each service commits alone"],
    ["Versions", "Compiled together, always match", "Deployed apart, so old and new meet each other"],
    ["Data", "A join", "An API call, or a copy of the data"],
    ["Debugging", "One stack trace", "A trace across many services, if you built it"],
  ], { fs: 12.5, rowH: [0.42, 0.55, 0.65, 0.65, 0.55, 0.55, 0.55, 0.55] });
  footer(s, n, false);
});

// ---- deep dive: transactions
N(`TIME: about 3 minutes.
SAY: Deep dive one: transactions. In the monolith, checkout is one transaction: reserve stock, charge the card, create the order. If anything raises, everything rolls back.
SAY: Across services, stock, payments and orders each have their own database and commit alone. There is no BEGIN and COMMIT that covers all three. If a step fails, the steps before it have already happened. Charge fails after stock is reserved: stock held for nobody. Order not created after the card was charged: a customer who paid and has no order, the worst one. Reply lost after a successful charge: a retry that charges twice.
ASK: Which of these three would hurt your business most?`,
`الوقت: حوالي 3 دقائق.
قل: الوقفة المعمقة الأولى: المعاملات. في النظام الواحد، الدفع معاملة واحدة: احجز المخزون، اخصم من البطاقة، أنشئ الطلب. إذا فشل أي شيء يتراجع كل شيء.
قل: عبر الخدمات، المخزون والدفع والطلبات لكل منها قاعدة بياناتها وتثبّت (commit) وحدها. لا يوجد BEGIN وCOMMIT يغطي الثلاثة. إذا فشلت خطوة فالتي قبلها قد حدثت فعلًا. فشل الخصم بعد حجز المخزون: مخزون محجوز لأحد لا وجود له. عدم إنشاء الطلب بعد خصم البطاقة: عميل دفع وليس له طلب، وهذا أسوأها. ضياع الرد بعد خصم ناجح: إعادة محاولة تخصم مرتين.
اسأل: أي هذه الثلاثة يؤذي عملكم أكثر؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Deep dive", "Transactions in microservices", false, { sub: "In the monolith, checkout was one transaction. Across services, there is none." });
  codeBlock(s, 0.5, 1.8, 6.3, 2.6, "checkout in the monolith", TEAL, [
    "with db.transaction():              # all or nothing",
    "    stock.reserve(sku, qty)",
    "    payments.charge(card, total)",
    "    orders.create(cart, total)      # any error: all rolled back",
  ], 11);
  infoCard(s, 7.05, 1.8, 5.8, 2.6, RUST, "Across services", "Stock, Payments and Orders each have their own database and commit alone. There is no BEGIN and COMMIT that covers all three. If a step fails, the steps before it have already happened.", 13);
  tbl(s, 0.5, 4.65, [5.2, 7.15], [
    ["What fails", "What you are left with"],
    ["The charge fails after stock was reserved", "Stock held for nobody"],
    ["The order is not created after the card was charged", "A customer who paid and has no order. The worst one"],
    ["The charge succeeds but the reply is lost", "A retry that charges twice, unless the charge is safe to repeat"],
  ], { fs: 12, rowH: [0.4, 0.5, 0.55, 0.55] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: What can you do? First, avoid the need: keep things that must change together inside one service and one database. This is why Orders and Inventory stayed together in our worked example. Second, two-phase commit: a coordinator asks everyone to prepare, then to commit. It is all or nothing, but slow, it holds locks, and it blocks if the coordinator dies. It is rarely used between services. Third, a saga: local transactions, each with an undo step. We build one in Part 4. Fourth, the outbox plus idempotency, so nothing is lost and repeats are harmless.
SAY: Four small rules help. Put the hardest-to-undo step last: reserve stock first, charge the card last. Give every reservation a time limit. Send an idempotency key with every charge. And show pending honestly to the user.`,
`الوقت: حوالي 3 دقائق.
قل: ماذا نفعل؟ أولًا، تجنب الحاجة: أبقِ ما يجب أن يتغير معًا داخل خدمة واحدة وقاعدة بيانات واحدة. لهذا بقيت الطلبات والمخزون معًا في مثالنا. ثانيًا، الالتزام على مرحلتين (two-phase commit): منسّق يطلب من الجميع الاستعداد ثم الالتزام. كل شيء أو لا شيء، لكنه بطيء ويحتفظ بأقفال ويتوقف إذا مات المنسّق، ونادرًا ما يُستخدم بين الخدمات. ثالثًا، الـ saga: معاملات محلية لكل منها خطوة تراجع. نبني واحدة في الجزء الرابع. رابعًا، الـ outbox مع الأمان عند التكرار (idempotency)، فلا يضيع شيء والتكرار غير مؤذٍ.
قل: أربع قواعد صغيرة تساعد. ضع الخطوة الأصعب في التراجع أخيرًا: احجز المخزون أولًا واخصم من البطاقة أخيرًا. أعطِ كل حجز مهلة زمنية. أرسل مفتاح idempotency مع كل خصم. وأظهر حالة "قيد المعالجة" للمستخدم بصدق.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Deep dive", "Transactions: your options", false);
  tbl(s, 0.5, 1.75, [2.3, 4.4, 5.65], [
    ["Option", "How it works", "Trade-off"],
    ["Avoid the need", "Keep things that must change together inside one service and one database.", "The best answer when you can. It is why Orders and Inventory stayed together."],
    ["Two-phase commit", "A coordinator asks every service \"can you commit?\", then tells them all to commit.", "All or nothing, but slow, holds locks, and blocks if the coordinator dies. Rarely used between services."],
    ["Saga", "A series of local transactions, each with an undo step (Part 4).", "No locks across services, but the system is only partly done for a while, and you write the undo code."],
    ["Outbox + idempotency", "Save the event with the data, and make every step safe to repeat.", "Nothing is lost and repeats are harmless. Works together with a saga."],
  ], { fs: 12, rowH: [0.4, 0.75, 0.85, 0.85, 0.75] });
  card(s, 0.5, 5.5, 12.35, 1.35, WHITE, { shadow: false, line: LINE });
  s.addText([{ text: "Rules that help", options: { bold: true, color: TEAL, fontFace: HEAD, fontSize: 13, breakLine: true, paraSpaceAfter: 3 } }].concat(bl([
    "Put the step that is hardest to undo last: reserve stock first, charge the card last.",
    "Give every reservation a time limit, so a lost message cannot hold stock forever.",
    "Send an idempotency key with every charge. Show \"pending\" honestly to the user.",
  ], { fs: 11.5, color: INK_SOFT, gap: 2 })), { x: 0.75, y: 5.58, w: 11.9, h: 1.2, isTextBox: true, margin: 0, valign: "top" });
  footer(s, n, false);
});

// ---- deep dive: data
N(`TIME: about 3 minutes.
SAY: Deep dive two: data. In the monolith, showing a customer's orders with their names is one query, and the database does the join. Across services, Orders only has a customer ID. The names live in another service with its own database. There is no join.
SAY: The first thing people write is a loop: for each order, ask the Customers service for the name. This is the N plus 1 trap. Fifty orders at 20 milliseconds each, one after another, is about a second. One batch call with many IDs does the same job in about 25 milliseconds.
ASK: Has anyone met an N plus 1 problem, in a database or an API?`,
`الوقت: حوالي 3 دقائق.
قل: الوقفة المعمقة الثانية: البيانات. في النظام الواحد، عرض طلبات عميل مع أسمائهم استعلام واحد وقاعدة البيانات تتولى الـ join. عبر الخدمات، خدمة الطلبات لديها رقم العميل فقط. الأسماء في خدمة أخرى لها قاعدتها. لا يوجد join.
قل: أول ما يكتبه الناس حلقة تكرار: لكل طلب اسأل خدمة العملاء عن الاسم. هذا فخ N+1. خمسون طلبًا بعشرين ميلي ثانية لكل منها، واحدًا بعد الآخر، حوالي ثانية. استدعاء واحد بمجموعة أرقام ينجز العمل نفسه في حوالي 25 ميلي ثانية.
اسأل: هل صادف أحدكم مشكلة N+1 في قاعدة بيانات أو في API؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Deep dive", "Data: joining vs API calls", false, { sub: "In the monolith the database does the join. Across services, there is nothing to join." });
  codeBlock(s, 0.5, 1.8, 6.05, 1.55, "monolith: one query, one round trip", TEAL, [
    "SELECT o.id, o.total, c.name",
    "FROM orders o JOIN customers c ON c.id = o.customer_id",
    "WHERE o.status = 'pending'",
  ], 10.5);
  infoCard(s, 6.8, 1.8, 6.05, 1.55, RUST, "Across services", "Orders only has a customer_id. The names live in the Customers service, which has its own database. So the first thing people write is a loop.", 12);
  codeBlock(s, 0.5, 3.6, 6.05, 3.2, "the N+1 trap: 50 orders, 50 calls", RUST, [
    "for order in orders:",
    "    order.customer_name = (",
    "        customers.get(order.customer_id).name",
    "    )",
    "# 50 orders x 20 ms = about 1 second,",
    "# one after another",
  ], 11);
  codeBlock(s, 6.8, 3.6, 6.05, 3.2, "one call for all of them: about 25 ms", TEAL, [
    "names = customers.get_many(",
    "    {o.customer_id for o in orders}",
    ")   # one request, many IDs",
    "",
    "for order in orders:",
    "    order.customer_name = names[order.customer_id]",
  ], 11);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Your options. Call per row: simple, but slow, and you depend on the other service. Batch call: fast enough for most screens, still depends on the other service at read time. Compose in a gateway: one place calls several services and joins in memory. The screen gets one call, but that layer now holds join logic. Local copy from events: listen to CustomerRenamed and keep your own copy. Fast and independent, but a little old. Snapshot at write time: store the customer name on the order itself. Best for facts that should not change after the order.
SAY: What you lose for good: ad hoc joins across the whole company's data. The usual answer is to send events from every service into one place built for those questions: a data warehouse. That is the next session.`,
`الوقت: حوالي 3 دقائق.
قل: خياراتكم. استدعاء لكل صف: بسيط لكنه بطيء وتعتمدون على الخدمة الأخرى. استدعاء مجمّع (batch): سريع بما يكفي لمعظم الشاشات، ويعتمد على الخدمة الأخرى وقت القراءة. التجميع في بوابة: مكان واحد يستدعي عدة خدمات ويدمج في الذاكرة. الشاشة تستدعي مرة واحدة لكن تلك الطبقة صارت تحمل منطق الدمج. نسخة محلية من الأحداث: استمع لحدث CustomerRenamed واحتفظ بنسختك. سريعة ومستقلة لكنها قديمة قليلًا. لقطة وقت الكتابة: خزّن اسم العميل في الطلب نفسه. الأفضل للحقائق التي لا يجب أن تتغير بعد الطلب.
قل: ما تخسرونه نهائيًا: الـ join الحر على بيانات الشركة كلها. الجواب المعتاد إرسال الأحداث من كل خدمة إلى مكان واحد مصمم لهذه الأسئلة: مستودع بيانات (data warehouse). وهذا موضوع الجلسة القادمة.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Deep dive", "Data: your options", false);
  tbl(s, 0.5, 1.75, [2.4, 4.8, 5.15], [
    ["Option", "How it works", "Trade-off"],
    ["Call per row", "Ask the other service for each item.", "Simple, but slow (N+1), and you depend on that service being up."],
    ["Batch call", "Ask once for many IDs.", "Fast enough for most screens. Still depends on the other service at read time."],
    ["Compose in a gateway", "One place (an API gateway or backend for frontend) calls several services and joins in memory.", "The screen gets one call, but that layer now holds join logic."],
    ["Local copy from events", "Listen to CustomerRenamed and keep your own copy.", "Fast and independent, but slightly old. Handle missed or repeated events."],
    ["Snapshot at write time", "Store the customer name on the order itself.", "Best for facts that should not change. Not for live data."],
  ], { fs: 11.5, rowH: [0.4, 0.6, 0.6, 0.85, 0.7, 0.6] });
  callout(s, 0.5, 5.95, 12.35, 0.9, "What you lose for good:", "ad hoc joins across the whole company's data. The usual answer is to send events from every service into a data warehouse. That is the next session.", false);
  footer(s, n, false);
});

// ---- deep dive: debugging
N(`TIME: about 3 minutes.
SAY: Deep dive three: debugging. In the monolith, one request is one process: one stack trace, one log file, and a debugger you can attach. Across services, one request visits five services on five machines. Each has its own logs, and their clocks do not agree. The question "why was this order slow?" has no single place to look, unless you built one.
SAY: Four tools. A correlation ID: one ID created at the edge, passed on every call, and written in every log line. Structured logs: JSON with fields you can filter. Distributed tracing: a timeline of one request across services. Metrics per service: rate, errors and duration, which show which service to look at.`,
`الوقت: حوالي 3 دقائق.
قل: الوقفة المعمقة الثالثة: تتبع الأخطاء. في النظام الواحد، الطلب الواحد عملية واحدة: أثر خطأ واحد، وملف سجل واحد، ومصحّح يمكن ربطه. عبر الخدمات، الطلب الواحد يزور خمس خدمات على خمسة أجهزة. لكل منها سجلاتها، وساعاتها غير متفقة. سؤال "لماذا تأخر هذا الطلب؟" ليس له مكان واحد للنظر، إلا إذا بنيتموه.
قل: أربع أدوات. معرّف الارتباط (correlation ID): رقم واحد يُنشأ عند الحافة ويمرّر في كل استدعاء ويُكتب في كل سطر سجل. سجلات منظمة: JSON بحقول يمكن التصفية عليها. التتبع الموزع (distributed tracing): خط زمني لطلب واحد عبر الخدمات. ومقاييس لكل خدمة: المعدل والأخطاء والمدة، وهي تدلكم على الخدمة التي تنظرون إليها.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Deep dive", "Debugging across services", false);
  infoCard(s, 0.5, 1.75, 6.05, 1.35, TEAL, "In the monolith", "One request is one process: one stack trace, one log file, a debugger you can attach.", 12.5);
  infoCard(s, 6.8, 1.75, 6.05, 1.35, RUST, "Across services", "One request visits five services on five machines. Each has its own logs, and their clocks do not agree.", 12.5);
  tbl(s, 0.5, 3.35, [2.6, 9.75], [
    ["Tool", "What it gives you"],
    ["Correlation ID", "One ID created at the edge, passed on every call, and written in every log line. Search for it and you see the whole story."],
    ["Structured logs", "Logs as JSON with fields (request_id, order_id, service), so you can filter, not just read."],
    ["Distributed tracing", "A timeline of one request across services, with how long each step took (OpenTelemetry is the common standard)."],
    ["Metrics per service", "The rate of requests, the share that fail, and how long they take. Shows which service to look at."],
  ], { fs: 12, rowH: [0.4, 0.7, 0.6, 0.7, 0.7] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Here is a trace of one slow checkout. The bars are time, from zero to 500 milliseconds. The user waits in Orders, so the first guess is "Orders is slow". The trace shows Orders is only waiting. The time is spent in Payments, waiting for the card provider. Without the trace you would spend a day in the wrong service.
SAY: And the code to make this possible is small: reuse the caller's request ID or create one, put it in every log line, and pass it on in a header. Three lines per service.
ASK: Could you find out today, in under five minutes, where a slow request in your system spent its time?`,
`الوقت: حوالي 3 دقائق.
قل: هذا أثر (trace) لطلب دفع بطيء واحد. الأشرطة هي الزمن من صفر إلى 500 ميلي ثانية. المستخدم ينتظر في خدمة الطلبات، فالتخمين الأول "الطلبات بطيئة". يبيّن الأثر أن الطلبات تنتظر فقط. الوقت يُصرف في خدمة الدفع وهي تنتظر مزوّد البطاقات. بدون الأثر ستقضون يومًا في الخدمة الخطأ.
قل: والكود اللازم صغير: أعد استخدام رقم الطلب الذي جاء من المستدعي أو أنشئ رقمًا، واكتبه في كل سطر سجل، ومرّره في ترويسة. ثلاثة أسطر لكل خدمة.
اسأل: هل تستطيعون اليوم، في أقل من خمس دقائق، معرفة أين أمضى طلب بطيء في نظامكم وقته؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Deep dive", "A trace of one slow checkout", false, { sub: "The user waits in Orders. Where does the time really go?" });
  card(s, 0.5, 1.8, 12.35, 3.35, WHITE, { shadow: false, line: LINE });
  const X0 = 2.6, W = 9.8, sc = W / 500, tx = (t) => X0 + t * sc;
  [["Gateway", 0, 480, false], ["Orders", 10, 470, false], ["Inventory", 30, 70, false], ["Payments", 80, 460, true], ["Card provider", 100, 450, true]].forEach(([nm, a, b, bad], i) => {
    const y = 2.05 + i * 0.52;
    label(s, nm, 0.7, y, 1.8, 0.36, { size: 10.5, align: "left" });
    s.addShape(pptx.ShapeType.rect, { x: tx(a), y, w: (b - a) * sc, h: 0.36, fill: { color: bad ? RUST_TINT : TEAL_TINT }, line: { color: bad ? RUST : TEAL, width: bad ? 1.5 : 1 } });
    label(s, (b - a) + " ms", tx(a), y, (b - a) * sc, 0.36, { size: 9.5, bold: true, color: bad ? RUST : INK_SOFT, align: "right" });
  });
  s.addShape(pptx.ShapeType.line, { x: X0, y: 4.75, w: W, h: 0, line: { color: INK_FAINT, width: 1 } });
  [0, 100, 200, 300, 400, 500].forEach((t) => label(s, String(t) + (t === 500 ? " ms" : ""), tx(t) - 0.4, 4.78, 0.8, 0.25, { size: 8.5, color: INK_FAINT }));
  codeBlock(s, 0.5, 5.3, 12.35, 1.55, "pass the ID on, and log it: three lines per service", TEAL, [
    "request_id = request.headers.get(\"X-Request-ID\") or new_id()      # reuse the caller's ID if there is one",
    "log.info(\"placing order\", request_id=request_id, order_id=order.id)   # every log line carries it",
    "inventory.reserve(sku, qty, headers={\"X-Request-ID\": request_id})   # and every outgoing call",
  ], 10.5);
  footer(s, n, false);
});

// ---- spectrum
N(`TIME: about 2 minutes.
SAY: A spectrum, not a switch. Monolith, modular monolith, a few services, many microservices. More independence and more cost as you move right. A modular monolith is a good place to start. A few services is a good place to stay, for a long time. Most teams do not need to go all the way to the right.
ASK: Where is your system on this line, and where does your team think it is?`,
`الوقت: حوالي دقيقتين.
قل: طيف وليس مفتاحًا. نظام واحد، نظام واحد بوحدات، قليل من الخدمات، خدمات مصغرة كثيرة. استقلال أكبر وكلفة أكبر كلما تحركتم لليمين. النظام الواحد المقسّم إلى وحدات مكان جيد للبدء. وقليل من الخدمات مكان جيد للبقاء مدة طويلة. أغلب الفرق لا تحتاج الوصول إلى أقصى اليمين.
اسأل: أين يقع نظامكم على هذا الخط، وأين يظن فريقكم أنه يقع؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · The big picture", "A spectrum, not a switch", false);
  flowRow(s, 2.4, [
    ["Monolith", "", WHITE, INK_SOFT, 1.5],
    ["Modular monolith", "", TEAL_TINT, TEAL, 3],
    ["A few services", "", TEAL_TINT, TEAL, 3],
    ["Many microservices", "", WHITE, INK_SOFT, 1.5],
  ], { w: 2.6, h: 1.1, tfs: 15 });
  label(s, "a good place to start", 3.6, 3.65, 2.6, 0.3, { size: 10.5, bold: true, color: TEAL });
  label(s, "a good place to stay", 6.55, 3.65, 2.6, 0.3, { size: 10.5, bold: true, color: TEAL });
  arrow(s, 0.8, 5.2, 12.5, 5.2, INK_FAINT, { w: 2 });
  label(s, "more independence, and more cost", 3.5, 5.3, 6.3, 0.35, { size: 11, color: INK_SOFT });
  s.addText("Most teams do not need to go all the way to the right.", { x: 0.5, y: 6.1, w: 12.3, h: 0.5, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 14, color: INK_SOFT });
  footer(s, n, false);
});

N(`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Give pairs two minutes. For each situation they say better, worse or both. Take a show of hands, then reveal.
WATCH FOR: People who say every row is better. Ask what they pay for each.
NOTE: The answers are in the notes of this slide. Three of five are worse. A split trades simple code for independence, and you should make that trade only when you need the independence.`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: أعطِ الأزواج دقيقتين. لكل حالة يقولون أفضل أم أسوأ أم كلاهما. خذ رفع الأيدي ثم اكشف الإجابات.
انتبه: من يقول إن كل الصفوف أفضل. اسأله ماذا يدفع مقابل كل صف.
ملاحظة: ثلاثة من خمسة أسوأ. القسمة تستبدل كودًا بسيطًا باستقلال، ولا يجب أن تفعلوا ذلك إلا إذا احتجتم الاستقلال.`);
checkIn(1, 3, "Better or worse after a split?", "A shop splits its monolith into services along sensible lines. For each situation, is it better, worse, or both?", [
  "Deploying a fix to the checkout without touching anything else.",
  "Showing the customer's name next to each order.",
  "Finding out why one order failed.",
  "Adding more machines for just the search part, for a sale.",
  "Renaming a field that three other teams read.",
], null, "a. Better. Independent deploy, if the checkout can really stand alone.\nb. Worse. The join is gone. You make an API call, or keep a copy of the name that you must keep fresh.\nc. Worse. The failure may be in any of five services. You need a correlation ID and tracing, or you are guessing.\nd. Better. Scale only what is hot. One of the clearest reasons to split.\ne. Worse. It is now a contract between teams. Add the new field, keep the old one for a while, and move people over.\n\nThree of five are worse. A split trades simple code for independence, and you should make that trade only when you need the independence.");

// ============================================================ PART 2
N(`TIME: 30 seconds.
SAY: Part 2. Start with a modular monolith. A boundary costs very little inside one program, and you can still move it later.`,
`الوقت: 30 ثانية.
قل: الجزء الثاني. ابدأ بنظام واحد مقسّم إلى وحدات. الحد داخل برنامج واحد رخيص جدًا، ويمكنكم نقله لاحقًا.`);
divider(2, "0:35–1:00", "Start with a modular monolith", "Draw the lines inside one program first, and see if they hold.");

N(`TIME: about 3 minutes.
SAY: Here is the shortcut that makes any later split impossible. Orders needs to reserve stock, so it reads the inventory table directly, and updates it. It works today. But Inventory can never change its table, and Orders can never live in another program, because it depends on a table that is not its own.
ASK: Has anyone seen this in a real codebase? How did it get there?`,
`الوقت: حوالي 3 دقائق.
قل: هذا هو الاختصار الذي يجعل أي قسمة لاحقة مستحيلة. خدمة الطلبات تريد حجز المخزون، فتقرأ جدول المخزون مباشرة وتعدّله. يعمل اليوم. لكن المخزون لن يستطيع أبدًا تغيير جدوله، والطلبات لن تستطيع العيش في برنامج آخر، لأنها تعتمد على جدول ليس لها.
اسأل: هل رأى أحدكم هذا في كود حقيقي؟ كيف وصل إليه؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · The shortcut", "The shortcut that makes splitting impossible", false);
  codeBlock(s, 0.5, 1.8, 7.3, 4.2, "orders/service.py", RUST, [
    "def place_order(db, sku, qty):",
    "    row = db.query_one(",
    "        \"SELECT stock FROM inventory_items WHERE sku = %s\", sku",
    "    )                                  # Orders reads Inventory's table",
    "    if row.stock < qty:",
    "        raise OutOfStock(sku)",
    "    db.execute(",
    "        \"UPDATE inventory_items SET stock = stock - %s WHERE sku = %s\",",
    "        qty, sku,",
    "    )                                  # and changes it",
  ], 11.5);
  infoCard(s, 8.05, 1.8, 4.8, 4.2, RUST, "Why it hurts", ["It works today.", "Inventory can never change its table.", "Orders can never live in another program.", "Orders depends on a table that is not its own.", "Splitting later means rewriting every place like this."], 13);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: The fix is a wall and a front door. Inventory has one public file, api.py, with a small interface: reserve and release. Everything else in inventory is private. Orders takes an InventoryApi and calls reserve. Today it is a plain function call.
SAY: Tomorrow this can be an HTTP client with the same interface, and Orders does not change. But note: the failures change, because it is now a network call. Add a timeout, and an idempotency key on reserve so a retry does not reserve twice. That is Day 4 again.`,
`الوقت: حوالي 3 دقائق.
قل: الحل جدار وباب أمامي. للمخزون ملف عام واحد، api.py، فيه واجهة صغيرة: reserve وrelease. كل ما عدا ذلك داخل المخزون خاص. الطلبات تستقبل InventoryApi وتستدعي reserve. اليوم هو استدعاء دالة عادي.
قل: غدًا يمكن أن يصبح عميل HTTP بالواجهة نفسها دون أن تتغير الطلبات. لكن انتبهوا: الإخفاقات تتغير لأنه صار استدعاء عبر الشبكة. أضيفوا مهلة ومفتاح idempotency على reserve حتى لا تحجز إعادة المحاولة مرتين. هذا اليوم الرابع مرة أخرى.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · The front door", "A wall and a front door", false);
  codeBlock(s, 0.5, 1.8, 7.3, 4.2, "inventory/api.py + orders/service.py", TEAL, [
    "# inventory/api.py: the only file other modules may import",
    "class InventoryApi(Protocol):",
    "    def reserve(self, sku: str, qty: int) -> Reservation: ...",
    "    def release(self, reservation_id: str) -> None: ...",
    "",
    "# orders/service.py",
    "class PlaceOrder:",
    "    def __init__(self, inventory: InventoryApi):",
    "        self.inventory = inventory",
    "",
    "    def __call__(self, cart):",
    "        reservation = self.inventory.reserve(cart.sku, cart.qty)   # a plain call",
  ], 10.5);
  infoCard(s, 8.05, 1.8, 4.8, 2.6, TEAL, "Tomorrow", "This can be an HTTP client with the same interface, and Orders does not change.", 13);
  infoCard(s, 8.05, 4.55, 4.8, 1.45, AMBER, "But the failures change", "Add a timeout, and an idempotency key on reserve (Day 4).", 12.5);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: A wall that only exists in a document gets broken within a month. Check it automatically. Public API only: other modules may import inventory dot api and nothing else from inventory. Own your tables: each module owns its tables, ideally in its own schema, and only that module's database user can touch them. One direction: if Orders uses Inventory, Inventory must not use Orders. And a test that fails the build. In Python, import-linter can do this; in Java, ArchUnit.
SAY: The config on the right says: Orders may not import inventory models or inventory service. If someone does, the build fails.`,
`الوقت: حوالي 3 دقائق.
قل: الجدار الموجود في وثيقة فقط يُكسر خلال شهر. تحققوا منه آليًا. واجهة عامة فقط: للوحدات الأخرى أن تستورد inventory.api ولا شيء غيره من المخزون. امتلكوا جداولكم: كل وحدة تملك جداولها، ويفضّل في مخطط (schema) خاص بها، ولا يلمسها إلا مستخدم قاعدة بيانات تلك الوحدة. اتجاه واحد: إذا كانت الطلبات تستخدم المخزون فلا يجوز للمخزون استخدام الطلبات. واختبار يفشل البناء. في بايثون تفعل ذلك import-linter، وفي جافا ArchUnit.
قل: الإعداد على اليمين يقول: لا يجوز للطلبات استيراد inventory.models أو inventory.service. إذا فعل أحدهم يفشل البناء.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Make it real", "Make the wall real", false, { sub: "A wall that only exists in a document gets broken within a month. Check it automatically." });
  infoCard(s, 0.5, 1.8, 6.1, 4.9, TEAL, "Four rules", ["Public API only. Other modules may import inventory.api and nothing else from inventory.", "Own your tables. Each module owns its tables, ideally in its own schema, and only that module's database user can touch them.", "One direction. If Orders uses Inventory, Inventory must not use Orders. Use an event instead.", "A test that fails the build. import-linter in Python, ArchUnit in Java."], 13);
  codeBlock(s, 6.85, 1.8, 6.0, 4.9, ".importlinter: fails the build", TEAL, [
    "[importlinter:contract:orders-uses-inventory-api-only]",
    "name = Orders may only use the inventory public API",
    "type = forbidden",
    "source_modules = orders",
    "forbidden_modules =",
    "    inventory.models",
    "    inventory.service",
  ], 10.5);
  footer(s, n, false);
});

N(`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Pairs mark each of the four as fine or broken. Then reveal.
SAY: The test: could the other module move to another program without changing this code, apart from the transport? If not, the wall is broken.
WATCH FOR: Row d. Importing another module's model feels harmless, but it ties the two together.`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: تحدد الأزواج لكل من الأربعة هل هو سليم أم مكسور. ثم اكشف الإجابات.
قل: الاختبار: هل تستطيع الوحدة الأخرى الانتقال إلى برنامج آخر دون تغيير هذا الكود، عدا وسيلة النقل؟ إذا لا فالجدار مكسور.
انتبه: الصف د. استيراد نموذج وحدة أخرى يبدو غير مؤذٍ لكنه يربط الوحدتين معًا.`);
checkIn(2, 3, "Spot the broken wall", "These are four things that happen in a modular monolith. Which ones break a module wall?", [
  "The Orders module imports inventory.api and calls reserve().",
  "The Orders module runs SELECT ... FROM inventory_items.",
  "Inventory publishes a StockReserved event and Orders listens for it.",
  "The Payments module imports orders.models.Order and reads its fields.",
], null, "a. Fine. That is the front door.\nb. Broken. Orders reaches into a table it does not own. This is the shortcut.\nc. Fine. An event is a public message, not a private detail.\nd. Broken. Payments depends on Orders' private model. If Orders changes its fields, Payments breaks. Payments should receive what it needs (an order ID and an amount) through an API or an event.\n\nThe test: could the other module move to another program without changing this code, apart from the transport? If not, the wall is broken.");

// ============================================================ PART 3
N(`TIME: 30 seconds.
SAY: Part 3. Where to draw the lines. This is the hardest part and the one people rush.`,
`الوقت: 30 ثانية.
قل: الجزء الثالث. أين نرسم الحدود. هذا أصعب جزء وأكثر ما يستعجله الناس.`);
divider(3, "1:00–1:30", "Where to draw the lines", "Rules of thumb, good and bad reasons, the distributed monolith, and an exercise in pairs.");

N(`TIME: about 3 minutes.
SAY: Five rules of thumb. Own your data: a service is the only one that reads and writes its tables. If two services need the same table, they are probably one service. Change together, stay together: if a feature always needs edits in two places, they belong in one. Split by business capability: Orders, Payments, Search, not "the database service". Match the team: Conway's law says a system tends to copy the communication of the team that builds it. Use that on purpose. And split for a reason you can name.
ASK: Which of these does your system break most often?`,
`الوقت: حوالي 3 دقائق.
قل: خمس قواعد عملية. امتلك بياناتك: الخدمة هي الوحيدة التي تقرأ جداولها وتكتب فيها. إذا احتاجت خدمتان الجدول نفسه فهما على الأرجح خدمة واحدة. ما يتغير معًا يبقى معًا: إذا احتاجت ميزة دائمًا تعديلًا في مكانين فمكانهما واحد. اقسم حسب القدرة التجارية: الطلبات والدفع والبحث، وليس "خدمة قاعدة البيانات". طابق الفريق: قانون كونواي يقول إن النظام يميل إلى نسخ تواصل الفريق الذي يبنيه. استخدموا ذلك عن قصد. واقسموا لسبب تستطيعون تسميته.
اسأل: أي هذه القواعد يكسرها نظامكم أكثر؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · The rules", "Five rules of thumb", false);
  tbl(s, 0.5, 1.8, [3.3, 9.05], [
    ["Rule", "What it means"],
    ["Own your data", "A service is the only one that reads and writes its tables. If two services need the same table, they are probably one service."],
    ["Change together, stay together", "If a feature always needs edits in two places, they belong in one place."],
    ["Split by business capability", "Orders, Payments, Search. Not \"the database service\" or \"the UI service\"."],
    ["Match the team", "Conway's law: a system tends to copy the communication of the team that builds it. Use that on purpose: one team per service."],
    ["Split for a reason you can name", "A different scale, a different release speed, a different security need, or a team that is blocked."],
  ], { fs: 13, rowH: [0.45, 0.85, 0.75, 0.75, 0.9, 0.8] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Good reasons: teams keep blocking each other's releases; one part needs much more scale than the rest, like search on sale day; one part must be isolated for security or compliance, like card data; parts that release at very different speeds; a clear contract that rarely changes.
SAY: Bad reasons: it is modern; our code is a mess, and a mess split into services is a mess over a network; we might need to scale someday; one small team maintaining fourteen services; splitting before the boundary has been stable inside the monolith.
ASK: Which bad reason have you heard in a real meeting?`,
`الوقت: حوالي 3 دقائق.
قل: أسباب وجيهة: الفرق تعطّل إصدارات بعضها باستمرار؛ جزء يحتاج توسعًا أكبر بكثير من الباقي مثل البحث في يوم التخفيضات؛ جزء يجب عزله للأمان أو الامتثال مثل بيانات البطاقات؛ أجزاء تُصدر بسرعات مختلفة جدًا؛ عقد واضح نادرًا ما يتغير.
قل: أسباب سيئة: لأنه حديث؛ كودنا فوضى، والفوضى المقسّمة إلى خدمات فوضى عبر شبكة؛ قد نحتاج التوسع يومًا ما؛ فريق صغير واحد يصون أربع عشرة خدمة؛ القسمة قبل أن يثبت الحد داخل النظام الواحد.
اسأل: أي سبب سيئ سمعتموه في اجتماع حقيقي؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · The reasons", "Good reasons and bad reasons to split", false);
  card(s, 0.5, 1.8, 6.05, 5.0, WHITE);
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.8, w: 6.05, h: 0.09, fill: { color: TEAL }, line: { type: "none" } });
  s.addText([{ text: "Good reasons", options: { bold: true, color: TEAL, fontFace: HEAD, fontSize: 18, breakLine: true, paraSpaceAfter: 8 } }].concat(bl([
    "Teams keep blocking each other's releases.",
    "One part needs much more scale than the rest (search on sale day).",
    "One part must be isolated for security or compliance (card data).",
    "Parts that release at very different speeds.",
    "A clear contract that rarely changes.",
  ], { fs: 13.5, color: INK_SOFT, gap: 9 })), { x: 0.8, y: 2.1, w: 5.5, h: 4.5, isTextBox: true, margin: 0, valign: "top" });
  card(s, 6.8, 1.8, 6.05, 5.0, WHITE);
  s.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.8, w: 6.05, h: 0.09, fill: { color: RUST }, line: { type: "none" } });
  s.addText([{ text: "Bad reasons", options: { bold: true, color: RUST, fontFace: HEAD, fontSize: 18, breakLine: true, paraSpaceAfter: 8 } }].concat(bl([
    "\"It is modern\" or \"everyone does it\".",
    "\"Our code is a mess.\" A mess split into services is a mess over a network.",
    "\"We might need to scale someday.\"",
    "One small team of five maintaining fourteen services.",
    "Splitting before the boundary has been stable inside the monolith.",
  ], { fs: 13.5, color: INK_SOFT, gap: 9 })), { x: 7.1, y: 2.1, w: 5.5, h: 4.5, isTextBox: true, margin: 0, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: If you split and keep the coupling, you get a distributed monolith: all the cost of a network and none of the independence. The picture: four services in a chain, all sharing one database. The signs: services must deploy together in a fixed order; several services share one database; one change needs edits in four repositories; one request makes a long chain of calls; you cannot run or test one service without the others.
ASK: Count how many of these five signs your system has.`,
`الوقت: حوالي 3 دقائق.
قل: إذا قسمتم وأبقيتم الترابط، تحصلون على نظام موزع متشابك (distributed monolith): كلفة الشبكة كاملة بلا أي استقلال. الصورة: أربع خدمات في سلسلة تشترك في قاعدة بيانات واحدة. العلامات: الخدمات يجب أن تُنشر معًا بترتيب ثابت؛ عدة خدمات تشترك في قاعدة واحدة؛ التغيير الواحد يحتاج تعديلات في أربعة مستودعات؛ الطلب الواحد يمر بسلسلة طويلة من الاستدعاءات؛ لا تستطيعون تشغيل أو اختبار خدمة دون الباقي.
اسأل: عدّوا كم علامة من الخمس موجودة في نظامكم.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · The trap", "The distributed monolith: the worst of both", false, { sub: "All the cost of a network, and none of the independence." });
  card(s, 0.5, 1.85, 7.4, 4.95, WHITE, { shadow: false, line: LINE });
  ["A", "B", "C", "D"].forEach((t, i) => {
    const x = 0.8 + i * 1.75;
    box(s, x, 2.5, 1.35, 0.9, PLUM_TINT, PLUM, 1.5); label(s, "Service " + t, x, 2.5, 1.35, 0.9, { size: 10.5, bold: true });
    if (i < 3) arrow(s, x + 1.38, 2.95, x + 1.72, 2.95, RUST, { w: 1.6 });
    arrow(s, x + 0.675, 3.4, x + 0.675, 4.35, INK_SOFT, { w: 1.2 });
  });
  box(s, 1.6, 4.4, 5.0, 0.8, WHITE, INK_SOFT, 2); label(s, "one shared database", 1.6, 4.4, 5.0, 0.8, { size: 12, bold: true });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.65, y: 2.3, w: 6.9, h: 1.3, rectRadius: 0.1, fill: { type: "none" }, line: { color: RUST, width: 1.5, dashType: "dash" } });
  label(s, "deploy together, in order", 0.65, 5.45, 6.9, 0.3, { size: 10.5, bold: true, color: RUST });
  label(s, "a long chain of calls for one request", 0.65, 5.85, 6.9, 0.3, { size: 10.5, color: INK_SOFT });
  infoCard(s, 8.15, 1.85, 4.7, 4.95, RUST, "The signs", ["Services must be deployed together, in a fixed order.", "Several services share one database.", "One change needs edits in four repositories.", "One request makes a long chain of calls.", "You cannot run or test one service without the others."], 12.5);
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Our shop. Notifications: split now. It is asynchronous and a failure must never stop checkout. Payments: split now, a third party with card-data rules. Search: split now, a different data shape and scale, and it can be rebuilt from events. Catalog: maybe, read-heavy and cacheable, but ask whether you have a real problem first. Orders plus Inventory: not yet. They need one transaction and change together. Splitting them forces a saga before you have earned it. Loyalty points: later, easy but low value.
ASK: Does anyone disagree with Orders and Inventory staying together?`,
`الوقت: حوالي 3 دقائق.
قل: متجرنا. الإشعارات: نقسّمها الآن، فهي غير متزامنة وفشلها يجب ألا يوقف الدفع أبدًا. المدفوعات: الآن، طرف ثالث وقواعد بيانات البطاقات. البحث: الآن، شكل بيانات وحجم مختلفان ويمكن إعادة بنائه من الأحداث. الكتالوج: ربما، كثير القراءة ويقبل التخزين المؤقت، لكن اسألوا أولًا هل لديكم مشكلة حقيقية. الطلبات مع المخزون: ليس بعد. يحتاجان معاملة واحدة ويتغيران معًا. فصلهما يفرض saga قبل أن تستحقوها. نقاط الولاء: لاحقًا، سهلة لكن قيمتها قليلة.
اسأل: هل يعترض أحد على بقاء الطلبات والمخزون معًا؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Worked example", "Our shop: what would you split first?", false);
  tbl(s, 0.5, 1.8, [2.7, 1.6, 8.05], [
    ["Part", "Split now?", "Why"],
    ["Notifications", "Yes", "It is asynchronous. A failure here must never stop checkout. It reads events from a queue."],
    ["Payments", "Yes", "A third party, card-data rules, and its own scaling and audit needs."],
    ["Search", "Yes", "Different data shape and scale. It can be rebuilt from events, so a copy of the data is fine."],
    ["Catalog", "Maybe", "Read-heavy and cacheable. Easy to split, but ask whether you have a real problem first."],
    ["Orders + Inventory", "Not yet", "They need one transaction and change together. Splitting them forces a saga before you have earned it."],
    ["Loyalty points", "Later", "It only needs events from Orders, so it is easy to split, but low value."],
  ], { fs: 13, rowH: [0.45, 0.75, 0.7, 0.75, 0.75, 0.85, 0.7] });
  footer(s, n, false);
});

N(`TIME: 12 minutes (about 10 in pairs, 2 to compare).
SAY: Exercise A. A food-delivery monolith has these parts: Restaurants, Menus, Orders, Promotions, Payments, Couriers and live location, Notifications, Reviews. One database, a team of 20, and slow releases because everyone edits the same code. Draw the boundaries. At most four deployable units. Mark the data each owns. Name one thing you will not split yet. Draw one synchronous call and one event.
WATCH FOR: Pairs who make eight services, one per noun. Ask what they gain from each. And pairs who put Orders and Menus apart: ask how an order gets its prices in one transaction.
NOTE: A sample answer is on the next slide. Do not show it until pairs have compared.`,
`الوقت: 12 دقيقة (حوالي 10 في أزواج و2 للمقارنة).
قل: التمرين أ. نظام توصيل طعام بكتلة واحدة فيها: المطاعم، القوائم، الطلبات، العروض، الدفع، السائقون والموقع الحي، الإشعارات، التقييمات. قاعدة بيانات واحدة وفريق من 20 والإصدارات بطيئة لأن الجميع يعدّل الكود نفسه. ارسموا الحدود. أربع وحدات نشر على الأكثر. حددوا بيانات كل وحدة. سمّوا شيئًا لن تقسموه الآن. وارسموا استدعاءً متزامنًا واحدًا وحدثًا واحدًا.
انتبه: الأزواج التي تصنع ثماني خدمات، واحدة لكل اسم. اسألهم ماذا يكسبون من كل واحدة. والأزواج التي تفصل الطلبات عن القوائم: اسألهم كيف يحصل الطلب على أسعاره في معاملة واحدة.
ملاحظة: الإجابة النموذجية في الشريحة التالية. لا تعرضها قبل أن تقارن الأزواج.`);
add((n) => {
  const s = newSlide(false);
  banner(s, "Exercise A · 12 min · in pairs", 0.5, 0.42, 3.4);
  s.addText("Draw the boundaries", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("A food-delivery monolith has one database and a team of 20. Releases are slow because everyone edits the same code. These are its parts:", { x: 0.5, y: 1.5, w: 12.3, h: 0.55, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  ["Restaurants", "Menus", "Orders", "Promotions", "Payments", "Couriers + live location", "Notifications", "Reviews"].forEach((t, i) => {
    const x = 0.5 + (i % 4) * 3.1, y = 2.15 + Math.floor(i / 4) * 0.7;
    box(s, x, y, 2.9, 0.55, PLUM_TINT, PLUM, 1.25); label(s, t, x, y, 2.9, 0.55, { size: 11, bold: true });
  });
  card(s, 0.5, 3.75, 7.0, 3.05, WHITE);
  qBlock(s, 0.75, 3.9, 6.5, 2.8, "Your task", ["Group the parts into at most four deployable units. Give each a one-line reason.", "Mark which data each unit owns.", "Name one thing you will not split yet, and why.", "Draw one synchronous call and one event between your units."], 13);
  card(s, 7.75, 3.75, 5.1, 3.05, WHITE);
  s.addText([{ text: "Done when", options: { fontFace: HEAD, fontSize: 15, bold: true, color: TEAL, breakLine: true, paraSpaceAfter: 6 } }].concat(bl(["No two units share a table.", "Each unit could be owned by one team.", "You can say what failure you are protecting against, or what scale you are gaining."], { fs: 13, color: INK, gap: 8 })), { x: 8.0, y: 3.9, w: 4.65, h: 2.8, isTextBox: true, margin: 0, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: One good answer. Four units. Marketplace core: restaurants, menus, orders and promotions. They change together, and an order needs menu prices and promotions in one transaction. Payments: a third party, card-data and audit rules. Courier dispatch: couriers and live location, a very different scale because location updates are frequent, and a different team. Notifications: asynchronous, a failure must not stop an order. Reviews can stay in the core for now.
SAY: Not split yet: Orders from Menus. They share a transaction and change together. One synchronous call: Orders asks Payments to charge, because it needs the answer. One event: OrderPlaced, which Courier dispatch and Notifications both read.
NOTE: There is no single right answer. Judge by reasons.`,
`الوقت: حوالي 3 دقائق.
قل: إجابة جيدة واحدة. أربع وحدات. نواة السوق: المطاعم والقوائم والطلبات والعروض. تتغير معًا والطلب يحتاج أسعار القائمة والعروض في معاملة واحدة. الدفع: طرف ثالث وقواعد البطاقات والتدقيق. توزيع السائقين: السائقون والموقع الحي، حجم مختلف جدًا لأن تحديثات الموقع كثيرة، وفريق مختلف. الإشعارات: غير متزامنة وفشلها يجب ألا يوقف الطلب. التقييمات يمكن أن تبقى في النواة الآن.
قل: لم نقسم بعد: الطلبات عن القوائم. تشتركان في معاملة وتتغيران معًا. استدعاء متزامن واحد: الطلبات تطلب من الدفع الخصم لأنها تحتاج الجواب. وحدث واحد: OrderPlaced، تقرؤه خدمة توزيع السائقين والإشعارات.
ملاحظة: لا توجد إجابة صحيحة واحدة. احكموا بالأسباب.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Exercise A · A sample answer", "One good answer for the food-delivery app", false, { sub: "There is no single right answer. Judge by the reasons." });
  const units = [
    [0.5, 1.95, 5.6, 2.35, TEAL_TINT, TEAL, "Marketplace core", "Restaurants, Menus, Orders, Promotions", "Owns: restaurants, menus, orders, promotions. They change together and share one transaction."],
    [7.25, 1.95, 5.6, 2.35, PLUM_TINT, PLUM, "Payments", "A third party, card-data and audit rules", "Owns: payment records."],
    [0.5, 4.6, 5.6, 2.2, AMBER_TINT, AMBER, "Courier dispatch", "Couriers and live location", "Owns: courier and location data. Very different scale."],
    [7.25, 4.6, 5.6, 2.2, "E3EAEC", SLATE, "Notifications", "Reviews stay in the core for now", "Owns: delivery logs. A failure must not stop an order."],
  ];
  units.forEach(([x, y, w, h, f, c, t, sub, own]) => {
    box(s, x, y, w, h, f, c, 2);
    s.addText([{ text: t, options: { fontFace: HEAD, bold: true, fontSize: 16, color: c, breakLine: true, paraSpaceAfter: 2 } }, { text: sub, options: { fontFace: BODY, fontSize: 12, color: INK, breakLine: true, paraSpaceAfter: 5 } }, { text: own, options: { fontFace: BODY, italic: true, fontSize: 11, color: INK_SOFT } }], { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: h - 0.3, isTextBox: true, margin: 0, valign: "top" });
  });
  arrow(s, 6.15, 3.1, 7.2, 3.1, RUST, { w: 2 }); label(s, "sync call:\ncharge", 6.1, 3.2, 1.2, 0.5, { size: 8.5, bold: true, color: RUST });
  arrow(s, 3.3, 4.32, 3.3, 4.58, TEAL, { w: 2, dash: "dash" }); label(s, "event: OrderPlaced (also to Notifications)", 3.4, 4.3, 3.2, 0.3, { size: 9, bold: true, color: TEAL, align: "left" });
  arrow(s, 6.15, 3.9, 7.2, 4.75, TEAL, { w: 2, dash: "dash" }); 
  footer(s, n, false);
});

// ============================================================ PART 4
N(`TIME: 30 seconds.
SAY: Part 4. Once there are several services, four questions come up: how do they call each other, how do they change without breaking each other, where does shared data live, and how do you keep it correct without a shared transaction?`,
`الوقت: 30 ثانية.
قل: الجزء الرابع. عندما توجد عدة خدمات تظهر أربعة أسئلة: كيف تستدعي بعضها، وكيف تتغير دون أن تكسر بعضها، وأين تعيش البيانات المشتركة، وكيف نحافظ على صحتها دون معاملة مشتركة؟`);
divider(4, "1:30–2:05", "Working together across the network", "Ask-and-wait or tell-and-move-on, contracts, sagas, and never losing an event.");

N(`TIME: about 3 minutes.
SAY: Two styles. A synchronous call, REST or gRPC: ask and wait for the answer. Use it when the caller needs the answer now, like is this card valid. An event or message, Kafka or a queue: tell and move on. Use it when others only need to know it happened, like an order was placed.
SAY: Failure differs. If the other side is down, a synchronous caller is down or waits. A message waits in the queue until the other side is back. Watch out for long chains and slow calls on the left, and for duplicates and order on the right. You have the tools from Day 4: timeouts, retries, circuit breakers, idempotent consumers.`,
`الوقت: حوالي 3 دقائق.
قل: أسلوبان. استدعاء متزامن، REST أو gRPC: اسأل وانتظر الجواب. استخدمه عندما يحتاج المستدعي الجواب الآن، مثل هل هذه البطاقة صالحة. وحدث أو رسالة، Kafka أو طابور: أخبر وامضِ. استخدمه عندما يكفي أن يعرف الآخرون أن الأمر حدث، مثل أن طلبًا وُضع.
قل: الفشل يختلف. إذا كان الطرف الآخر متوقفًا فالمستدعي المتزامن متوقف أو ينتظر. أما الرسالة فتنتظر في الطابور حتى يعود. احذروا من السلاسل الطويلة والاستدعاءات البطيئة على اليسار، ومن التكرار والترتيب على اليمين. لديكم أدوات اليوم الرابع: المهلات وإعادة المحاولة وقاطع الدائرة والمستهلكات الآمنة عند التكرار.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Talking to each other", "Ask and wait, or tell and move on", false);
  tbl(s, 0.5, 1.85, [1.9, 5.2, 5.25], [
    ["", "Synchronous call (REST, gRPC)", "Event or message (Kafka, a queue)"],
    ["Style", "Ask and wait for the answer", "Tell and move on"],
    ["Use when", "The caller needs the answer now (is this card valid?)", "Others only need to know it happened (an order was placed)"],
    ["Failure", "If the other side is down, you are down or you wait", "The message waits in the queue until the other side is back"],
    ["Watch out for", "Long chains and slow calls (timeouts, retries, circuit breakers: Day 4)", "Duplicates and order (idempotent consumers: Day 4)"],
  ], { fs: 13.5, rowH: [0.5, 0.7, 0.95, 0.95, 1.1] });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Chains multiply failure. Four services in a chain, each up 99.9 percent of the time, are together up only 99.6 percent of the time: 0.999 four times. That sounds small. It is about 35 hours of downtime a year instead of about 9. Every extra synchronous hop makes it worse.
ASK: How many synchronous hops are in your longest request?`,
`الوقت: حوالي 3 دقائق.
قل: السلاسل تضاعف الفشل. أربع خدمات في سلسلة، كل منها تعمل 99.9% من الوقت، تعمل معًا 99.6% فقط: 0.999 أربع مرات. يبدو الفرق صغيرًا، لكنه حوالي 35 ساعة توقف في السنة بدل حوالي 9. كل قفزة متزامنة إضافية تزيد الأمر سوءًا.
اسأل: كم قفزة متزامنة في أطول طلب عندكم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Talking to each other", "Chains multiply failure", false, { sub: "Four services in a chain, each up 99.9% of the time." });
  flowRow(s, 1.95, [["Service A", "99.9%", TEAL_TINT, TEAL], ["Service B", "99.9%", TEAL_TINT, TEAL], ["Service C", "99.9%", TEAL_TINT, TEAL], ["Service D", "99.9%", TEAL_TINT, TEAL]], { w: 2.4, h: 1.1, tfs: 14, sfs: 13 });
  card(s, 0.5, 3.4, 12.35, 0.9, WHITE, { shadow: false, line: LINE });
  s.addText("0.999 × 0.999 × 0.999 × 0.999 = 99.6% together", { x: 0.5, y: 3.4, w: 12.35, h: 0.9, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: CODE, bold: true, fontSize: 18, color: RUST });
  label(s, "ONE SERVICE AT 99.9%", 0.7, 4.65, 3.5, 0.3, { size: 10, bold: true, color: TEAL, align: "left" });
  s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 5.0, w: 8.76 * 9 / 35, h: 0.5, fill: { color: TEAL_TINT }, line: { color: TEAL, width: 1 } });
  label(s, "about 9 hours down a year", 0.8 + 8.76 * 9 / 35, 5.0, 4.5, 0.5, { size: 12, align: "left" });
  label(s, "FOUR IN A CHAIN AT 99.6%", 0.7, 5.65, 4.5, 0.3, { size: 10, bold: true, color: RUST, align: "left" });
  s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 6.0, w: 8.76, h: 0.5, fill: { color: RUST_TINT }, line: { color: RUST, width: 1.5 } });
  label(s, "about 35 hours down a year", 9.6, 6.0, 3.2, 0.5, { size: 12, bold: true, color: RUST, align: "left" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: An API between two teams is a contract. Adding a field is safe. Renaming or removing one breaks the other side. So: add the new field, keep the old one for a while, move the consumers over, then remove the old one.
SAY: For QEs, a consumer-driven contract test is the tool. Pact is a common one. Each consumer states what it needs from the provider, and the provider's build fails if a change would break one of them. You find out before release, not in production.
ASK: QEs: do you have any contract test at a boundary today?`,
`الوقت: حوالي 3 دقائق.
قل: الـ API بين فريقين عقد. إضافة حقل آمنة. أما إعادة تسمية حقل أو حذفه فيكسر الطرف الآخر. لذلك: أضف الحقل الجديد، وأبقِ القديم مدة، وانقل المستهلكين، ثم احذف القديم.
قل: لمهندسي الجودة الأداة هي اختبار العقد الموجَّه بالمستهلك (consumer-driven contract test). Pact من الأدوات الشائعة. كل مستهلك يصرّح بما يحتاجه من المزوّد، ويفشل بناء المزوّد إذا كان التغيير سيكسر أحدهم. تعرفون قبل الإصدار لا في الإنتاج.
اسأل: مهندسو الجودة: هل لديكم أي اختبار عقد عند حد ما اليوم؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Contracts", "Changing an API without breaking the other side", false, { sub: "Adding a field is safe. Renaming or removing one breaks the other team." });
  flowRow(s, 1.95, [["1  Add the new field", "", TEAL_TINT, TEAL], ["2  Keep the old one", "for a while", WHITE, INK_SOFT], ["3  Move consumers over", "", WHITE, INK_SOFT], ["4  Remove the old one", "", TEAL_TINT, TEAL]], { w: 2.5, h: 1.2, tfs: 13.5, sfs: 11 });
  infoCard(s, 0.5, 3.8, 6.05, 2.95, TEAL, "Consumer-driven contract test", ["Each consumer states what it needs from the provider (Pact is a common tool).", "The provider's build fails if a change would break a consumer.", "You find out before release, not in production."], 12.5);
  infoCard(s, 6.8, 3.8, 6.05, 2.95, PLUM, "For QEs", "A contract test is cheap and fast, and runs without the other service. It is the test that matters most at each boundary. A full end-to-end test of fourteen services is slow and fragile.", 12.5);
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: In the monolith, checkout was one transaction. With three services there is no shared transaction. A saga is a series of local steps, each with an undo step that runs if a later step fails.
SAY: The happy path: Orders creates the order as pending, Inventory reserves stock, Payments charges the card, Orders marks it confirmed. If the payment fails: Inventory releases the stock, and Orders cancels the order.
SAY: Notice that no step is a database rollback. The undo for stock is a real new action. The undo for the order is a state, cancelled. For a short time the system is only partly done, and that is eventual consistency.
ASK: What if the undo step itself fails?`,
`الوقت: حوالي 4 دقائق.
قل: في النظام الواحد كان الدفع معاملة واحدة. مع ثلاث خدمات لا توجد معاملة مشتركة. الـ saga سلسلة خطوات محلية، لكل منها خطوة تراجع تعمل إذا فشلت خطوة لاحقة.
قل: المسار الناجح: الطلبات تنشئ الطلب بحالة قيد المعالجة، المخزون يحجز، الدفع يخصم من البطاقة، الطلبات تؤكد. إذا فشل الدفع: المخزون يحرر الحجز والطلبات تلغي الطلب.
قل: لاحظوا أن لا خطوة هنا تراجع في قاعدة البيانات. التراجع عن المخزون فعل جديد حقيقي. والتراجع عن الطلب حالة، ملغى. لوقت قصير يكون النظام منجزًا جزئيًا، وهذا هو الاتساق في النهاية.
اسأل: ماذا لو فشلت خطوة التراجع نفسها؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Transactions across services", "A checkout saga", false, { sub: "A series of local steps. If a later step fails, each earlier step runs its undo." });
  label(s, "THE HAPPY PATH", 0.5, 1.8, 4, 0.3, { size: 10.5, bold: true, color: TEAL, align: "left" });
  flowRow(s, 2.15, [["1  Orders", "create order: pending", TEAL_TINT, TEAL], ["2  Inventory", "reserve stock", TEAL_TINT, TEAL], ["3  Payments", "charge the card", TEAL_TINT, TEAL], ["4  Orders", "mark: confirmed", TEAL_TINT, TEAL]], { w: 2.5, h: 1.15, tfs: 13, sfs: 11 });
  label(s, "IF PAYMENT FAILS, UNDO", 0.5, 4.0, 5, 0.3, { size: 10.5, bold: true, color: RUST, align: "left" });
  const gap = (12.35 - 4 * 2.5) / 3;
  const bx = (i) => 0.5 + i * (2.5 + gap);
  [[2, "3  Payments", "charge fails"], [1, "undo 2  Inventory", "release stock"], [0, "undo 1  Orders", "cancel the order"]].forEach(([i, t, sub]) => {
    box(s, bx(i), 4.4, 2.5, 1.15, RUST_TINT, RUST, 1.5);
    s.addText([{ text: t, options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: sub, options: { fontFace: BODY, fontSize: 11, color: RUST } }], { x: bx(i) + 0.1, y: 4.4, w: 2.3, h: 1.15, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  });
  arrow(s, bx(2) + 1.25, 3.35, bx(2) + 1.25, 4.38, RUST, { w: 1.6, dash: "dash" });
  arrow(s, bx(2) - 0.05, 4.97, bx(1) + 2.55, 4.97, RUST, { w: 1.6 });
  arrow(s, bx(1) - 0.05, 4.97, bx(0) + 2.55, 4.97, RUST, { w: 1.6 });
  s.addText("No step is a database rollback. The undo for stock is a real new action (release it). The undo for the order is a state (cancelled). For a short time the system is only partly done. That is eventual consistency.", { x: 0.5, y: 5.85, w: 12.35, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: Two ways to run a saga. Choreography: each service listens for events and reacts. Orders publishes OrderPlaced, Inventory reacts, then Payments reacts. There is no central brain, but it is hard to see the whole flow. Orchestration: one coordinator calls each step and the undo steps, and records where the saga is. It is easy to follow and debug, but the coordinator is one more thing to build and run.
SAY: A simple rule: for two or three steps, choreography is fine. As the flow grows or has many failure paths, an orchestrator pays for itself.`,
`الوقت: حوالي 3 دقائق.
قل: طريقتان لتشغيل الـ saga. التناغم (choreography): كل خدمة تستمع للأحداث وتتفاعل. الطلبات تنشر OrderPlaced، فيتفاعل المخزون ثم الدفع. لا يوجد دماغ مركزي، لكن يصعب رؤية التدفق كله. التنسيق المركزي (orchestration): منسّق واحد يستدعي كل خطوة وخطوات التراجع ويسجل موقع الـ saga. سهل المتابعة والتصحيح، لكن المنسّق شيء آخر يجب بناؤه وتشغيله.
قل: قاعدة بسيطة: لخطوتين أو ثلاث يكفي التناغم. عندما يكبر التدفق أو تكثر مسارات الفشل، يستحق المنسّق كلفته.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Transactions across services", "Two ways to run a saga", false);
  card(s, 0.5, 1.85, 6.05, 4.85, WHITE);
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.85, w: 6.05, h: 0.09, fill: { color: PLUM }, line: { type: "none" } });
  s.addText("Choreography", { x: 0.8, y: 2.1, w: 5.5, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: PLUM });
  s.addText("Each service listens for events and reacts.", { x: 0.8, y: 2.55, w: 5.5, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK });
  [["Orders", 0.8], ["Inventory", 2.5], ["Payments", 4.2]].forEach(([t, x], i) => { box(s, x, 3.2, 1.5, 0.8, PLUM_TINT, PLUM, 1.5); label(s, t, x, 3.2, 1.5, 0.8, { size: 10.5, bold: true }); if (i < 2) arrow(s, x + 1.52, 3.6, x + 1.68, 3.6, RUST, { w: 1.5, dash: "dash" }); });
  label(s, "event, event", 0.8, 4.1, 5.2, 0.3, { size: 9, color: INK_FAINT });
  s.addText(bl(["No central brain.", "Hard to see the whole flow.", "Fine for two or three steps."], { fs: 13, color: INK_SOFT, gap: 5 }), { x: 0.8, y: 4.6, w: 5.5, h: 1.9, isTextBox: true, margin: 0, valign: "top" });
  card(s, 6.8, 1.85, 6.05, 4.85, WHITE);
  s.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.85, w: 6.05, h: 0.09, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("Orchestration", { x: 7.1, y: 2.1, w: 5.5, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: TEAL });
  s.addText("One coordinator runs the steps and the undo steps.", { x: 7.1, y: 2.55, w: 5.5, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK });
  box(s, 8.5, 3.15, 2.0, 0.65, AMBER_TINT, AMBER, 2); label(s, "Coordinator", 8.5, 3.15, 2.0, 0.65, { size: 10.5, bold: true });
  [["Orders", 7.1], ["Inventory", 8.8], ["Payments", 10.5]].forEach(([t, x]) => { box(s, x, 4.2, 1.5, 0.7, TEAL_TINT, TEAL, 1.5); label(s, t, x, 4.2, 1.5, 0.7, { size: 10, bold: true }); arrow(s, 9.5, 3.82, x + 0.75, 4.18, INK_SOFT, { w: 1.2 }); });
  s.addText(bl(["Easy to follow and debug.", "One more thing to build and run.", "Pays for itself as the flow grows."], { fs: 13, color: INK_SOFT, gap: 5 }), { x: 7.1, y: 5.15, w: 5.5, h: 1.5, isTextBox: true, margin: 0, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: A trap. Orders saves the order, then publishes OrderPlaced. If it crashes between the two, the order exists and nobody hears about it. On the right is the fix: the outbox. In one database transaction, write the order and an event row into an outbox table. A separate relay process publishes it later.
SAY: The relay may publish a row twice if it crashes before marking it sent. So consumers must be idempotent, the same event ID check from Day 4. The outbox gives you never lost. The idempotent consumer gives you safe when repeated. You need both.`,
`الوقت: حوالي 3 دقائق.
قل: فخ. الطلبات تحفظ الطلب ثم تنشر OrderPlaced. إذا انهارت بين الخطوتين يكون الطلب موجودًا ولا أحد يعرف. على اليمين الحل: الـ outbox. في معاملة قاعدة بيانات واحدة نكتب الطلب وصف حدث في جدول outbox. ثم عملية ترحيل منفصلة تنشره لاحقًا.
قل: قد ينشر المرحّل الصف مرتين إذا انهار قبل أن يعلّمه مرسلًا. لذلك يجب أن تكون المستهلكات آمنة عند التكرار، وهو فحص رقم الحدث نفسه من اليوم الرابع. الـ outbox يعطيكم "لا يضيع شيء". والمستهلك الآمن عند التكرار يعطيكم "التكرار غير مؤذٍ". تحتاجون الاثنين.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Never lose an event", "The outbox", false, { sub: "Saving the order and publishing the event are two steps. A crash can fall between them." });
  card(s, 0.5, 1.85, 6.05, 4.9, WHITE, { shadow: false, line: RUST });
  label(s, "WITHOUT AN OUTBOX", 0.7, 1.95, 5.6, 0.3, { size: 10.5, bold: true, color: RUST, align: "left" });
  [[2.5, "1  Save the order", "committed", TEAL_TINT, TEAL], [3.9, "2  Crash", "the server dies here", RUST_TINT, RUST], [5.3, "3  Publish OrderPlaced", "never happens", WHITE, INK_FAINT]].forEach(([y, t, sub, f, c], i) => {
    box(s, 0.9, y, 5.2, 0.95, f, c, 1.5);
    s.addText([{ text: t, options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: sub, options: { fontFace: BODY, fontSize: 11, color: c === INK_FAINT ? INK_SOFT : c } }], { x: 1.0, y, w: 5.0, h: 0.95, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    if (i < 2) arrow(s, 3.5, y + 0.95, 3.5, y + 1.4, INK_SOFT, { w: 1.4 });
  });
  card(s, 6.8, 1.85, 6.05, 4.9, WHITE, { shadow: false, line: TEAL });
  label(s, "WITH AN OUTBOX", 7.0, 1.95, 5.6, 0.3, { size: 10.5, bold: true, color: TEAL, align: "left" });
  [[2.5, "1  One transaction", "save the order + an outbox row", TEAL_TINT, TEAL], [3.9, "2  A relay publishes", "later, from the outbox", TEAL_TINT, TEAL], [5.3, "3  Consumers skip repeats", "by event ID (Day 4)", PLUM_TINT, PLUM]].forEach(([y, t, sub, f, c], i) => {
    box(s, 7.2, y, 5.2, 0.95, f, c, 1.5);
    s.addText([{ text: t, options: { fontFace: HEAD, bold: true, fontSize: 13, color: INK, breakLine: true } }, { text: sub, options: { fontFace: BODY, fontSize: 11, color: c } }], { x: 7.3, y, w: 5.0, h: 0.95, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    if (i < 2) arrow(s, 9.8, y + 0.95, 9.8, y + 1.4, INK_SOFT, { w: 1.4 });
  });
  footer(s, n, false);
});

N(`TIME: about 3 minutes.
SAY: The code. place_order opens one local transaction, inserts the order, and inserts an outbox row with the event, in the same commit. A separate relay reads unsent rows, publishes each to the broker with the row ID as the event ID, and marks it sent.
SAY: Look at the last line. If the relay crashes after publishing and before the update, the row is published again. That is fine, because the consumer checks the event ID.
ASK: What would happen if the relay marked the row sent first, then published?`,
`الوقت: حوالي 3 دقائق.
قل: الكود. place_order تفتح معاملة محلية واحدة، وتدخل الطلب، وتدخل صف outbox فيه الحدث، في الـ commit نفسه. ثم مرحّل منفصل يقرأ الصفوف غير المرسلة، وينشر كل صف إلى الوسيط برقم الصف كرقم للحدث، ويعلّمه مرسلًا.
قل: انظروا إلى السطر الأخير. إذا انهار المرحّل بعد النشر وقبل التحديث يُنشر الصف مرة أخرى. هذا مقبول لأن المستهلك يفحص رقم الحدث.
اسأل: ماذا لو علّم المرحّل الصف مرسلًا أولًا ثم نشر؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Never lose an event", "The outbox in code", false);
  codeBlock(s, 0.5, 1.75, 12.35, 4.35, "orders/outbox.py: save and publish safely", TEAL, [
    "def place_order(db, cart):",
    "    with db.transaction():                                    # one local transaction",
    "        order_id = db.insert(\"orders\", status=\"pending\", total=cart.total)",
    "        db.insert(\"outbox\", event=\"OrderPlaced\",",
    "                  payload={\"order_id\": order_id, \"items\": cart.items})    # same commit as the order",
    "    return order_id",
    "",
    "# a separate relay process",
    "def relay_outbox(db, broker):",
    "    for row in db.query(\"SELECT * FROM outbox WHERE sent_at IS NULL ORDER BY id LIMIT 100\"):",
    "        broker.publish(\"order-events\", key=row.payload[\"order_id\"], value=row.payload, event_id=row.id)",
    "        db.execute(\"UPDATE outbox SET sent_at = now() WHERE id = %s\", row.id)   # a crash before this line means one repeat",
  ], 10.5);
  callout(s, 0.5, 6.25, 12.35, 0.6, "This is Day 4 again.", "The outbox gives you never lost. The idempotent consumer gives you safe when repeated. You need both.", false);
  footer(s, n, false);
});

// quiz
N(`TIME: about 4 minutes (2 in pairs, 2 to compare).
FACILITATE: Give pairs two minutes. For each of the five situations they pick: modular monolith, split it out, or not yet. Take a show of hands for each. Do not give the answers yet.
SAY: For each one ask: is the reason about teams, scale or risk? Or is it about fashion?
NOTE: The answers are on the next slide.`,
`الوقت: حوالي 4 دقائق (دقيقتان في أزواج ودقيقتان للمقارنة).
سهّل النقاش: أعطِ الأزواج دقيقتين. لكل حالة من الحالات الخمس يختارون: نظام واحد بوحدات، أو يُفصل، أو ليس بعد. خذ رفع الأيدي لكل حالة. لا تعطِ الإجابات بعد.
قل: لكل حالة اسأل: هل السبب يتعلق بالفرق أو الحجم أو المخاطر؟ أم يتعلق بالموضة؟
ملاحظة: الإجابات في الشريحة التالية.`);
add((n) => {
  const s = newSlide(false);
  banner(s, "Quiz · 4 min · in pairs", 0.5, 0.42, 3.0);
  s.addText("Split it, or not?", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("For each situation, circle one: a modular monolith, split it out, or not yet.", { x: 0.5, y: 1.45, w: 12.3, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  [
    "A team of six is building one new product and must launch in three months.",
    "Card payments must be isolated from the rest of the system for compliance.",
    "Search needs 20 times more machines on sale days, but the rest of the system does not.",
    "Forty engineers in six teams keep blocking each other's releases on one codebase.",
    "\"Our code is a mess, so let's rewrite it as microservices.\"",
  ].forEach((q, i) => {
    const x = 0.5 + (i % 3) * 4.15, y = 2.0 + Math.floor(i / 3) * 2.4, w = 4.0;
    card(s, x, y, w, 2.2, WHITE);
    s.addText([{ text: String(i + 1), options: { fontFace: CODE, bold: true, fontSize: 14, color: TEAL, breakLine: true, paraSpaceAfter: 4 } }, { text: q, options: { fontFace: BODY, fontSize: 13, color: INK } }], { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: 1.35, isTextBox: true, margin: 0, valign: "top" });
    [["Modular monolith", TEAL, 1.55], ["Split it", PLUM, 0.95], ["Not yet", RUST, 0.85]].reduce((cx, [c, col, cw]) => {
      s.addShape(pptx.ShapeType.roundRect, { x: cx, y: y + 1.65, w: cw, h: 0.34, rectRadius: 0.17, fill: { type: "none" }, line: { color: col, width: 1.5 } });
      label(s, c, cx, y + 1.65, cw, 0.34, { size: 8.5, bold: true, color: col });
      return cx + cw + 0.1;
    }, x + 0.2);
  });
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Quiz: split it, or not?\n\n1 Modular monolith. Speed matters, the team is small, and you do not yet know where the boundaries are. Keep the walls clean so you can split later.\n2 Split it (Payments). A clear security and compliance reason, and a narrow contract.\n3 Split it (Search). A different scale is a good reason, and search can be rebuilt from events.\n4 Split it, along team lines. Independent deploys solve a real problem here. Start with the modules that cause the most blocking.\n5 Not yet. Clean it up and draw the walls inside the monolith first. Messy code in services is the same mess, now slow and spread across a network.\n\nThe reasons are about teams, scale and risk, never about fashion.");
});

N(`TIME: about 3 minutes.
SAY: The answers, with reasons. One: modular monolith, because speed matters and you do not yet know the boundaries. Two: split Payments, a clear compliance reason. Three: split Search, a different scale. Four: split along team lines, because independent deploys solve a real problem. Five: not yet. Clean it up inside the monolith first.
SAY: Notice that every reason is about teams, scale or risk. None is about fashion.`,
`الوقت: حوالي 3 دقائق.
قل: الإجابات مع الأسباب. واحد: نظام واحد بوحدات، لأن السرعة مهمة ولا تعرفون الحدود بعد. اثنان: افصلوا الدفع، سبب امتثال واضح. ثلاثة: افصلوا البحث، حجم مختلف. أربعة: افصلوا حسب الفرق، لأن النشر المستقل يحل مشكلة حقيقية. خمسة: ليس بعد. نظّفوه داخل النظام الواحد أولًا.
قل: لاحظوا أن كل سبب يتعلق بالفرق أو الحجم أو المخاطر، ولا واحد يتعلق بالموضة.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 4 · Quiz answers", "Split it, or not?", false, { sub: "Every reason is about teams, scale or risk. None is about fashion." });
  tbl(s, 0.5, 1.9, [4.7, 2.1, 5.55], [
    ["Situation", "Pick", "Why"],
    ["A team of six, one new product, launch in three months", "Modular monolith", "Speed matters and you do not yet know the boundaries. Keep the walls clean so you can split later."],
    ["Card payments must be isolated for compliance", "Split it", "A clear security and compliance reason, and a narrow contract."],
    ["Search needs 20 times more machines on sale days", "Split it", "A different scale is a good reason, and search can be rebuilt from events."],
    ["Forty engineers in six teams block each other's releases", "Split it", "Independent deploys solve a real problem. Split along team lines, starting where the blocking is worst."],
    ["\"Our code is a mess, let's rewrite as microservices\"", "Not yet", "Clean up and draw the walls inside the monolith first. Messy code in services is the same mess, now slow and over a network."],
  ], { fs: 12, rowH: [0.42, 0.85, 0.7, 0.7, 0.9, 1.0] });
  footer(s, n, false);
});

// ============================================================ PART 5
N(`TIME: 30 seconds.
SAY: Part 5. If you decide to split, do it slowly.`,
`الوقت: 30 ثانية.
قل: الجزء الخامس. إذا قررتم القسمة فافعلوا ذلك ببطء.`);
divider(5, "2:05–2:20", "Moving out safely, in either direction", "The strangler fig, six steps, and why going back is allowed.");

N(`TIME: about 3 minutes.
SAY: The strangler fig pattern, named by Martin Fowler after a vine that grows around a tree. You put a router in front of the system. At the start it sends everything to the old monolith. You build one new service for one capability. Then you send a small share of requests to it, for example 10 percent, or a few test users, and compare results. When it is trusted, you send all of it, and delete the old code.
ASK: Which capability in your system would be the first, safest one to strangle?`,
`الوقت: حوالي 3 دقائق.
قل: نمط التين الخانق (strangler fig)، سمّاه مارتن فاولر نسبة إلى نبتة تنمو حول شجرة. تضعون موجّهًا (router) أمام النظام. في البداية يرسل كل شيء إلى النظام القديم. تبنون خدمة جديدة واحدة لقدرة واحدة. ثم ترسلون إليها نسبة صغيرة من الطلبات، مثلًا 10%، أو بضعة مستخدمين تجريبيين، وتقارنون النتائج. عندما تثقون بها ترسلون كل الطلبات إليها وتحذفون الكود القديم.
اسأل: أي قدرة في نظامكم ستكون الأولى والأكثر أمانًا للاستبدال؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 5 · Moving out", "The strangler fig", false, { sub: "Replace the old system piece by piece, while it keeps running." });
  box(s, 0.5, 2.7, 1.9, 1.3, WHITE, INK_SOFT, 1.5); label(s, "Clients", 0.5, 2.7, 1.9, 1.3, { size: 13, bold: true });
  box(s, 3.4, 2.5, 2.6, 1.7, AMBER_TINT, AMBER, 2.5); s.addText([{ text: "Router", options: { fontFace: HEAD, bold: true, fontSize: 15, color: INK, breakLine: true } }, { text: "proxy or facade", options: { fontFace: BODY, fontSize: 11, color: INK_SOFT } }], { x: 3.4, y: 2.5, w: 2.6, h: 1.7, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 8.0, 1.85, 4.85, 1.5, WHITE, INK_SOFT, 1.5); s.addText([{ text: "Old monolith", options: { fontFace: HEAD, bold: true, fontSize: 15, color: INK, breakLine: true } }, { text: "still does everything else", options: { fontFace: BODY, fontSize: 11, color: INK_SOFT } }], { x: 8.0, y: 1.85, w: 4.85, h: 1.5, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  box(s, 8.0, 3.85, 4.85, 1.5, PLUM_TINT, PLUM, 2.5); s.addText([{ text: "New service", options: { fontFace: HEAD, bold: true, fontSize: 15, color: INK, breakLine: true } }, { text: "one capability, for example Notifications", options: { fontFace: BODY, fontSize: 11, color: INK_SOFT } }], { x: 8.0, y: 3.85, w: 4.85, h: 1.5, isTextBox: true, margin: 0, align: "center", valign: "middle" });
  arrow(s, 2.45, 3.35, 3.35, 3.35, INK_SOFT, { w: 1.8 });
  arrow(s, 6.05, 2.9, 7.95, 2.6, INK_SOFT, { w: 1.8 }); label(s, "90%", 6.6, 2.35, 0.9, 0.3, { size: 12, bold: true });
  arrow(s, 6.05, 3.8, 7.95, 4.5, PLUM, { w: 2.6 }); label(s, "10%", 6.6, 4.25, 0.9, 0.3, { size: 12, bold: true, color: PLUM });
  s.addText("The router decides, per request, where it goes. Start at 0%, then 10% (or a few test users), then more, and compare results. When the new service is trusted, send 100% and delete the old code.", { x: 0.5, y: 5.65, w: 12.35, h: 1.0, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 13, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

N(`TIME: about 4 minutes.
SAY: Six steps. One: put a router in front. Two: pick an edge first, like notifications, search or reporting. They depend on the rest, but little depends on them. Leave the core, Orders and Inventory, for last, or forever. Three: build the new service behind an anti-corruption layer so it does not copy the old messy model. Four: move the data in steps. First the new service reads through the old one's API, then keeps its own copy updated by events, then owns it. Avoid dual writes, where you write the same data in two places by hand. One of the two will drift. Five: compare, then shift traffic slowly, with a way to go back. Six: delete the old code. If you do not, you have two systems.
SAY: And going back is allowed. A video-quality team at Amazon Prime Video wrote in 2023 that moving one tool from many cloud pieces into a single process cut its infrastructure cost by more than 90 percent. The lesson is not that monoliths win. It is to measure, and be willing to move in either direction.
NOTE: Check this reference before you cite it.`,
`الوقت: حوالي 4 دقائق.
قل: ست خطوات. واحد: ضعوا موجّهًا أمام النظام. اثنان: اختاروا طرفًا أولًا، مثل الإشعارات أو البحث أو التقارير. تعتمد على الباقي، لكن قليل يعتمد عليها. اتركوا النواة، الطلبات والمخزون، للأخير أو للأبد. ثلاثة: ابنوا الخدمة الجديدة خلف طبقة حماية من الفساد (anti-corruption layer) حتى لا تنسخ النموذج القديم الفوضوي. أربعة: انقلوا البيانات على مراحل. تقرأ الخدمة الجديدة أولًا عبر واجهة القديمة، ثم تحتفظ بنسختها محدّثة بالأحداث، ثم تملكها. تجنبوا الكتابة المزدوجة (dual writes)، أي كتابة البيانات نفسها في مكانين يدويًا. أحدهما سيتباعد. خمسة: قارنوا ثم حوّلوا الطلبات ببطء مع إمكانية الرجوع. ستة: احذفوا الكود القديم. وإلا صار لديكم نظامان.
قل: والرجوع مسموح. كتب فريق جودة الفيديو في Amazon Prime Video عام 2023 أن نقل أداة واحدة من قطع سحابية كثيرة إلى عملية واحدة خفّض كلفة بنيتها التحتية بأكثر من 90%. الدرس ليس أن النظام الواحد يفوز. الدرس أن تقيسوا وأن تكونوا مستعدين للتحرك في الاتجاهين.
ملاحظة: تحققوا من هذا المرجع قبل أن تذكروه.`);
add((n) => {
  const s = newSlide(false);
  header(s, "Part 5 · Moving out", "Six steps, and going back is allowed", false);
  [
    ["1", "Put a router in front", "It can send each request to the old code or the new service."],
    ["2", "Pick an edge first", "Notifications, search, reporting. Leave the core for last, or forever."],
    ["3", "Use an anti-corruption layer", "So the new service does not copy the old, messy model."],
    ["4", "Move the data in steps", "Read through the old API, then keep a copy from events, then own it. Avoid dual writes."],
    ["5", "Compare, then shift slowly", "With a way to go back at any time."],
    ["6", "Delete the old code", "If you do not, you have two systems."],
  ].forEach(([k, t, d], i) => {
    const x = 0.5 + (i % 3) * 4.15, y = 1.8 + Math.floor(i / 3) * 1.95;
    card(s, x, y, 4.0, 1.8, WHITE);
    circle(s, x + 0.2, y + 0.2, 0.6, TEAL, k, { fontSize: 16 });
    s.addText(t, { x: x + 0.95, y: y + 0.2, w: 2.9, h: 0.6, isTextBox: true, margin: 0, valign: "middle", fontFace: HEAD, bold: true, fontSize: 13.5, color: INK });
    s.addText(d, { x: x + 0.2, y: y + 0.95, w: 3.65, h: 0.8, isTextBox: true, margin: 0, valign: "top", fontFace: BODY, fontSize: 11.5, color: INK_SOFT });
  });
  callout(s, 0.5, 5.8, 12.35, 1.0, "Going back is allowed.", "A video-quality team at Amazon Prime Video wrote in 2023 that moving one tool from many cloud pieces into one process cut its infrastructure cost by more than 90%. The lesson is not \"monoliths win\". It is to measure, and be willing to move in either direction.", false);
  footer(s, n, false);
});

// ---- review + closing
N(`TIME: 8 minutes.
SAY: Pick the question that matches who is in the room, and let people answer in their own words. Everyone: draw your own system as boxes and arrows. Is it a monolith, a modular monolith, or a distributed monolith, and what is the clearest sign? Tech leads: has your team split something for a reason from the bad list, and how would you raise it without an argument? QEs: pick one boundary, and say which contract test would catch a breaking change before release. AI engineers: what four questions should an agent ask before it writes a new service?`,
`الوقت: 8 دقائق.
قل: اختاروا السؤال الذي يناسب من في القاعة ودعوا الناس يجيبون بكلماتهم. الجميع: ارسموا نظامكم كصناديق وأسهم. هل هو نظام واحد أم بوحدات أم نظام موزع متشابك، وما أوضح علامة؟ القادة التقنيون: هل قسّم فريقكم شيئًا لسبب من قائمة الأسباب السيئة، وكيف تطرحون ذلك دون جدال؟ مهندسو الجودة: اختاروا حدًا واحدًا وقولوا أي اختبار عقد سيلتقط تغييرًا يكسر الطرف الآخر قبل الإصدار. مهندسو الذكاء الاصطناعي: ما الأسئلة الأربعة التي يجب أن يطرحها الوكيل قبل أن يكتب خدمة جديدة؟`);
add((n) => {
  const s = newSlide(false);
  header(s, "2:20–2:30 · Review", "Group critique", false);
  const qs = [
    ["JR", TEAL, "Everyone", "Draw your own system as boxes and arrows. Is it a monolith, a modular monolith, or a distributed monolith? What is the clearest sign?"],
    ["SR", SLATE, "Tech Leads", "Has your team split something for a reason from the \"bad reasons\" list? How would you raise it without starting an argument?"],
    ["QE", RUST, "QEs", "Pick one boundary in your system. What contract test would tell you, before release, that a change breaks the other side?"],
    ["AI", PLUM, "AI Engineers", "If you asked an agent to \"add a service for X\", what four questions should it ask before writing any code?"],
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

N(`TIME: 2 minutes.
SAY: Before you go, pick one and do it this week. Draw the service map of your system and mark every shared database. Find one place where one module reads another module's tables, and write down the front door it should use. Count the synchronous calls in your longest request chain. Check that every event your system publishes can be handled twice without harm.
SAY: Next session, Day 8: data engineering. Today we said that when data is spread over many services, you cannot join it any more. Next time we see how to put it back together for questions: warehouses, lakes, and the pipelines between them.`,
`الوقت: دقيقتان.
قل: قبل أن تغادروا، اختاروا واحدًا ونفذوه هذا الأسبوع. ارسموا خريطة الخدمات في نظامكم وعلّموا كل قاعدة بيانات مشتركة. ابحثوا عن مكان واحد تقرأ فيه وحدة جداول وحدة أخرى، واكتبوا الباب الأمامي الذي يجب أن تستخدمه. عدّوا الاستدعاءات المتزامنة في أطول سلسلة طلبات عندكم. وتأكدوا أن كل حدث ينشره نظامكم يمكن معالجته مرتين دون ضرر.
قل: الجلسة القادمة، اليوم الثامن: هندسة البيانات. قلنا اليوم إنه عندما تتوزع البيانات على خدمات كثيرة لا يمكن عمل join عليها. في المرة القادمة نرى كيف نجمعها من جديد للإجابة عن الأسئلة: مستودعات البيانات والبحيرات والمسارات بينها.`);
add((n) => {
  const s = newSlide(true);
  s.addText("BEFORE YOU GO", { x: 0.7, y: 0.55, w: 8, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Pick one. Do it this week.", { x: 0.7, y: 0.9, w: 11, h: 0.8, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 34, color: WHITE });
  [
    "Draw the service map of your system and mark every shared database",
    "Find one place where a module reads another module's tables, and write down the front door it should use",
    "Count the synchronous calls in your longest request chain",
    "Check that every event your system publishes can be handled twice without harm",
  ].forEach((it, i) => {
    const y = 1.95 + i * 0.85;
    card(s, 0.7, y, 11.9, 0.68, DARK2, { shadow: false });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: y + 0.19, w: 0.3, h: 0.3, rectRadius: 0.04, fill: { type: "none" }, line: { color: TEAL, width: 1.5 } });
    s.addText(it, { x: 1.45, y, w: 10.9, h: 0.68, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 14, color: WHITE });
  });
  s.addShape(pptx.ShapeType.line, { x: 0.7, y: 5.6, w: 11.9, h: 0, line: { color: "234750", width: 1 } });
  s.addText("NEXT: DAY 08", { x: 0.7, y: 5.8, w: 6, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  s.addText("Data Engineering: warehouses, lakes and pipelines", { x: 0.7, y: 6.1, w: 11, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 22, color: WHITE });
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
pptx.writeFile({ fileName: "session-07.pptx" }).then(() => {
  try {
    require("child_process").execFileSync("python3", [__dirname + "/fix_rtl_notes.py", "session-07.pptx"], { stdio: "inherit" });
  } catch (e) { console.warn("RTL post-process skipped:", e.message); }
  console.log("written", TOTAL, "slides;", missing, "without notes");
});
