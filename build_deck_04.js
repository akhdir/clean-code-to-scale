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
const FOOT = "Day 04 · Clean Code to Scale";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// slides are queued so footers can show "NN / TOTAL"
const queue = [];
function add(fn) { queue.push(fn); }
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

// ============================================================ reusable slide kinds
function titleSlide() {
  add((n) => {
    const s = newSlide(true);
    s.addText("def buy(product_id):\n    with db.lock(product_id):\n        stock -= 1", { x: 7.3, y: 0.5, w: 5.7, h: 2.6, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 18, color: DARK_WATERMARK, lineSpacingMultiple: 1.3 });
    s.addText("CLEAN CODE TO SCALE   ·   DAY 04 OF 10", { x: 0.7, y: 2.3, w: 10, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
    s.addText("Scalability Fundamentals", { x: 0.7, y: 2.75, w: 11.5, h: 1.4, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 44, color: WHITE });
    s.addText("What happens when many people use your system at once: the web servers, the background jobs behind them, and the database underneath. Three layers, one flash sale.", { x: 0.7, y: 4.05, w: 10, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 16, color: "B9D3D6" });
    [["3:55", "DURATION"], ["PYTHON", "CODE SAMPLES"], ["PAIRS", "PLAN + 3 EXERCISES"]].forEach((st, i) => {
      const x = 0.7 + i * 3.1;
      s.addText(st[0], { x, y: 5.65, w: 2.9, h: 0.45, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 19, color: WHITE });
      s.addText(st[1], { x, y: 6.13, w: 2.9, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 9.5, color: "6C8B90", charSpacing: 1 });
    });
    footer(s, n, true, "Clean Code to Scale");
  });
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

// ============================================================ DECK
titleSlide();

// ---- agenda
add((n) => {
  const s = newSlide(false);
  header(s, "Session Map", "Three layers, then two hands-on blocks", false);
  const steps = [
    ["01", "0:00–0:55", "PART 1", TEAL, "Scaling web and API requests: stateless servers, sessions, the hot-row lock."],
    ["02", "0:55–1:45", "PART 2", SLATE, "Scaling backend jobs: queues, partitions, Kafka."],
    ["03", "1:45–2:40", "PART 3", PLUM, "Database scalability: replicas, multi-master, partitioning, distributed systems and CAP, and how failures spread."],
    ["04", "2:40–3:40", "HANDS-ON", RUST, "A capacity plan, then three infrastructure diagrams to diagnose."],
    ["05", "3:40–3:55", "REVIEW", INK_FAINT, "Group critique and one commitment each."],
  ];
  const startX = 0.75, gap = 2.55, cy = 2.4, d = 1.05;
  s.addShape(pptx.ShapeType.line, { x: startX + d / 2, y: cy + d / 2, w: gap * 4, h: 0, line: { color: LINE, width: 1.5 } });
  steps.forEach((st, i) => {
    const x = startX + i * gap;
    circle(s, x, cy, d, st[3], st[0], { fontSize: 19 });
    s.addText(st[1], { x: x - 0.7, y: cy + d + 0.15, w: d + 1.4, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: CODE, bold: true, fontSize: 13, color: INK });
    s.addText(st[2], { x: x - 0.7, y: cy + d + 0.47, w: d + 1.4, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: CODE, fontSize: 10, color: st[3], charSpacing: 1 });
    s.addText(st[4], { x: x - 0.7, y: cy + d + 0.9, w: d + 1.4, h: 1.1, isTextBox: true, margin: 0, align: "center", valign: "top", fontFace: BODY, fontSize: 11.5, color: INK_SOFT });
  });
  s.addText("Ten short check-in exercises are woven through Parts 1 to 3 (about 25 minutes in all).", { x: 0.5, y: 5.95, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  s.addText("Short on time? Run the check-ins as a quick poll first, then shorten the exercises, then make the capacity plan take-home.", { x: 0.5, y: 6.35, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  footer(s, n, false);
});

// ---- objectives
add((n) => {
  const s = newSlide(false);
  header(s, "Objectives", "What today is for", false);
  const items = [
    ["Scale a web or API tier", "Keep servers stateless, choose where sessions live, and avoid a locked row that makes 1,000 requests wait in line."],
    ["Know when adding servers stops helping", "Amdahl's Law and the USL, and why a capacity plan keeps spare room."],
    ["Put slow or bursty work behind a queue", "Scale it with partitions, and know what can go wrong and how Kafka handles it."],
    ["Pick a way to scale a database", "Replicas, multi-master, or splitting the data, and what each one gives up under CAP."],
  ];
  items.forEach((it, i) => {
    const y = 1.75 + i * 1.3;
    circle(s, 0.7, y, 0.56, TEAL, "✓", { fontFace: BODY, fontSize: 20 });
    s.addText(it[0], { x: 1.6, y: y - 0.06, w: 10.8, h: 0.42, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: INK });
    s.addText(it[1], { x: 1.6, y: y + 0.38, w: 10.8, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  });
  footer(s, n, false);
});

// ---- room
add((n) => {
  const s = newSlide(false);
  header(s, "Room", "Why this session, for you specifically", false);
  const roles = [
    ["JR", TEAL, "Junior – Mid Engineers", "The first session about what happens under heavy load, not just whether the code is readable. It builds on skills you already have."],
    ["SR", SLATE, "Senior Eng & Tech Leads", "Amdahl's Law answers \"will twice the servers give us twice the speed?\" before anyone spends the money."],
    ["QE", RUST, "QEs", "A load test only means something once \"scaled\" has a number. Queueing theory tells you how busy to run the system when you test."],
    ["AI", PLUM, "AI Engineers", "An agent told to \"make this faster\" often adds a cache or a bigger machine without asking if the service can run on several machines. This is the question to ask first."],
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

// ============================================================ PART 1
divider(1, "0:00–0:55", "Scaling web and API requests", "Web servers are the easiest part to scale out, if they keep no user data, sessions are stored on purpose, and a thousand requests never wait in line for one row.");

cards3("Part 1 · Concept", "Three ways to scale, one requirement", "Adding servers is the easiest to repeat, but it only works if the servers are stateless.", [
  { color: TEAL, title: "Vertical — scale up", points: ["A bigger machine. No code changes, so it is the easiest option.", "There is a limit: the biggest machine you can buy.", "One machine is one point of failure."] },
  { color: PLUM, title: "Horizontal — scale out", points: ["More machines running the same code, behind a load balancer.", "Cheap and close to linear at first, but it needs stateless servers.", "Backup comes free: if one server dies, the others keep working."] },
  { color: AMBER, title: "Functional — split by responsibility", points: ["Break a big application into parts, so each part can scale on its own.", "Contains failures: one overloaded part does not take down the rest.", "Function calls become network calls. That is Day 7's whole topic."] },
]);

// stateless diagram
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Concept", "Stateless servers: any server can answer", false, { sub: "Every server is identical and remembers nothing. That is the one requirement that scaling out cannot skip." });
  box(s, 0.6, 3.45, 1.5, 0.95, WHITE, INK_SOFT); label(s, "Client", 0.6, 3.45, 1.5, 0.95, { size: 12 });
  box(s, 2.75, 3.4, 1.9, 1.05, AMBER_TINT, AMBER, 2); label(s, "Load balancer", 2.75, 3.4, 1.9, 1.05, { size: 12, bold: true });
  [2.1, 3.55, 5.0].forEach((y, i) => {
    box(s, 5.5, y, 2.4, 0.95, TEAL_TINT, TEAL, 1.5);
    label(s, "instance " + "ABC"[i], 5.5, y + 0.05, 2.4, 0.5, { size: 12, bold: true });
    label(s, "no local state", 5.5, y + 0.5, 2.4, 0.35, { size: 9.5, color: INK_SOFT });
    arrow(s, 4.65, 3.92, 5.45, y + 0.47, INK_SOFT);
  });
  arrow(s, 2.1, 3.92, 2.7, 3.92, INK_SOFT);
  box(s, 9.3, 2.45, 2.3, 0.95, WHITE, INK_SOFT); label(s, "Cache", 9.3, 2.45, 2.3, 0.95, { size: 12 });
  box(s, 9.3, 4.95, 2.3, 0.95, WHITE, INK_SOFT); label(s, "Database", 9.3, 4.95, 2.3, 0.95, { size: 12 });
  [2.57, 4.02, 5.47].forEach((y) => arrow(s, 7.9, y, 9.25, 2.93, INK_SOFT));
  arrow(s, 10.45, 3.4, 10.45, 4.9, INK_SOFT);
  label(s, "miss → fetch", 10.55, 4.05, 1.8, 0.3, { size: 9, color: INK_FAINT, align: "left" });
  label(s, "any instance will do", 4.5, 1.9, 2.4, 0.3, { size: 10, color: AMBER, bold: true });
  label(s, "STATELESS — NO INSTANCE KNOWS ABOUT ANOTHER", 4.5, 6.2, 5.5, 0.3, { size: 10, color: TEAL, bold: true });
  footer(s, n, false);
});

// stateless vs stateful table
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Concept", "Stateless vs. stateful", false, { sub: "A server does not remember a user between two requests. The system still has state. It just lives somewhere else." });
  tbl(s, 0.5, 2.0, [2.9, 4.7, 4.7], [
    ["", "Stateful instance", "Stateless instance"],
    ["Where a user's data lives", "In this server's memory, such as a cart or session object", "A shared store (Redis, the database) or inside the request itself"],
    ["Which instance can answer", "Only the server that holds the data", "Any of them"],
    ["An instance dies", "Every user on it loses their session and cart", "Nothing is lost. The load balancer sends traffic to other servers"],
    ["Autoscale or deploy", "Users get logged out, and servers cannot be removed freely", "Servers are disposable. Users do not notice scaling"],
  ], { fs: 13, rowH: [0.45, 0.85, 0.85, 0.85, 0.85] });
  footer(s, n, false);
});

checkIn(1, 2, "What breaks with three servers?", "This code works on one server. Now there are three servers behind a load balancer. For each marked line, what goes wrong?", [
  "(a) request_count: each server has its own counter. Is that a problem?",
  "(b) rate_limits: a user may upload 10 files a minute. What happens with three servers?",
  "(c) thumb_cache: each server has its own cache. Is that a problem?",
  "(d) UPLOAD_DIR: the file is saved on the server's own disk. What happens on the next request?",
], {
  label: "upload.py", fs: 11.5, lines: [
    "request_count = 0                        # (a)",
    "rate_limits = {}                         # (b) user -> uploads/min",
    "thumb_cache = {}                         # (c)",
    'UPLOAD_DIR = "/var/app/uploads"          # (d)',
    "",
    '@app.post("/upload")',
    "def upload(file, user_id):",
    "    global request_count",
    "    request_count += 1                  # (a)",
    "    if rate_limits.get(user_id, 0) >= 10:   # (b)",
    "        return error(429)",
    "    rate_limits[user_id] = rate_limits.get(user_id, 0) + 1",
    "",
    '    path = f"{UPLOAD_DIR}/{file.name}"',
    "    file.save(path)                     # (d) THIS machine's disk",
    "    thumb_cache[path] = make_thumbnail(file)   # (c)",
    '    return {"path": path}',
  ],
}, "(a) Only the number is wrong. Each server counts only its own requests. Send the count to a metrics tool.\n(b) Real bug. A user can upload up to 30 files a minute (10 on each server). Keep the counter in Redis, so all servers share it.\n(c) Mostly fine. It is only a cache, so the worst case is a few extra misses. Use a shared cache if every server must agree.\n(d) Real bug. The next request may go to another server, which does not have the file (404). Store files in shared storage like S3.");

// session handling table
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Sessions", "Three ways to remember a user", false, { sub: "Each choice is a tradeoff, not a \"best practice\"." });
  tbl(s, 0.5, 1.95, [2.0, 3.9, 2.9, 3.55], [
    ["Approach", "How it works", "Strength", "Tradeoff"],
    ["Sticky sessions", "The load balancer sends each user to the same server every time (cookie or IP address). Session data stays in that server's memory.", "No code changes. Works for an app that already keeps state in memory.", "The server is still the only place the session lives. See the next slide."],
    ["Shared session store", "Servers keep nothing. Every request looks up the session in Redis using an ID in a cookie.", "Any server can handle any request. Servers are disposable.", "An extra network call, and the store is now critical: it needs its own backup and failover."],
    ["Client-held token", "The session travels with the request: a signed token (e.g. JWT) checked without a lookup.", "Nothing is stored on the server.", "Hard to cancel before it expires. Size limits. Private data must be encrypted, not just signed."],
  ], { fs: 12, rowH: [0.45, 1.3, 1.3, 1.3] });
  footer(s, n, false);
});

// sticky sessions
add((n) => {
  const s = newSlide(true);
  header(s, "Part 1 · Sessions", "Sticky sessions: why they cause trouble", true);
  codeBlock(s, 0.5, 1.7, 5.4, 2.2, "nginx.conf", TEAL_TINT, [
    "upstream checkout {",
    "    ip_hash;     # same IP -> same instance",
    "    server app-a:8000;",
    "    server app-b:8000;",
    "    server app-c:8000;",
    "}",
  ], 12);
  const pts = [
    ["Uneven load.", "A few heavy users stuck to one server can overload it while the rest sit idle, and the balancer cannot move them."],
    ["When a server dies, its sessions die.", "Every user stuck to it is logged out or loses their cart mid-sale."],
    ["Deploys and autoscaling get harder.", "To remove a server you must wait for its sessions to end."],
    ["IP-based routing misleads.", "Thousands of people behind one company network share one IP, so they land on the same server."],
    ["It hides the bug.", "A stateful app looks fine behind sticky routing, until a server is replaced under load."],
  ];
  pts.forEach((p, i) => {
    const y = 1.7 + i * 1.02;
    circle(s, 6.3, y + 0.05, 0.4, TEAL, String(i + 1), { fontSize: 12 });
    s.addText([
      { text: p[0] + " ", options: { bold: true, color: WHITE, fontFace: BODY, fontSize: 13 } },
      { text: p[1], options: { color: "C6DADC", fontFace: BODY, fontSize: 12 } },
    ], { x: 6.9, y, w: 5.95, h: 0.95, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, true);
});

checkIn(2, 3, "Which session approach?", "Pick one: sticky sessions, a shared session store (like Redis), or a token the user carries. Give one reason.", [
  "An internal admin tool. 2 servers, 20 users. Nobody minds logging in again after a deploy.",
  "A shopping site. Servers go from 3 to 30 during sales. Carts must survive deploys.",
  "A public API for mobile apps, running in several regions.",
], null, "1. Sticky sessions. It is small, and a lost session just means logging in again.\n2. Shared session store. Servers come and go all the time, so the cart must live outside them.\n3. A token the user carries. No single place can look up sessions for every region. Keep tokens short-lived, because they are hard to cancel.");

// demo: flash sale cart
add((n) => {
  const s = newSlide(true);
  header(s, "Part 1 · Demo", "Flash sale: carts start disappearing", true, { sub: "800 req/s was fine. The sale hits 6,000 req/s and one server hits 100% CPU. The obvious fix: add eight more servers. It breaks the carts." });
  codeBlock(s, 0.5, 1.9, 6.1, 3.6, "checkout_service.py — stateful", AMBER_LIGHT, [
    "# carts lives in this process's memory only",
    "carts = {}",
    "",
    "def add_to_cart(session_id, item):",
    "    carts.setdefault(session_id, []).append(item)",
    "",
    "def get_cart(session_id):",
    "    return carts.get(session_id, [])",
  ], 13);
  codeBlock(s, 6.75, 1.9, 6.1, 3.6, "checkout_service.py — stateless", TEAL_TINT, [
    "def add_to_cart(session_id, item):",
    '    redis.rpush(f"cart:{session_id}", item)',
    "",
    "def get_cart(session_id):",
    '    return redis.lrange(f"cart:{session_id}", 0, -1)',
  ], 13);
  callout(s, 0.5, 5.7, 12.35, 1.15, "Why more servers made it worse:", "request 1 (add to cart) goes to server A; request 2 (view cart) goes to server C, which has never heard of this customer. The bug was always there: with one server, the same server answered both. Sticky sessions would have hidden it until a server was replaced mid-sale.", true);
  footer(s, n, true);
  s.addNotes("Talking points:\n- Same two functions, same call sites — only where the state lives changed. Statelessness is a persistence decision, not a rewrite.\n- Now any of the nine instances can answer either request, because Redis is the source of truth.\n- One more cheap lever: the product-price lookup was hitting the database on every cart operation. A cache in front of it cuts database load without adding an instance.\n- Sticky sessions would have hidden this bug: each customer keeps hitting the instance holding their cart, so it works — until that instance is replaced mid-sale and every pinned customer's cart vanishes at once.");
});

// hot row: before + throughput table
add((n) => {
  const s = newSlide(true);
  header(s, "Part 1 · The hot row", "1,000 requests, one row", true, { sub: "The lock exists for a good reason. But it turns many requests at once into one request at a time." });
  codeBlock(s, 0.5, 1.8, 6.6, 4.3, "buy.py — before", AMBER_LIGHT, [
    "def buy(product_id, user_id):",
    "    with db.transaction():",
    "        row = db.query_one(",
    '            "SELECT stock, price FROM products "',
    '            "WHERE id = %s FOR UPDATE", product_id)',
    "                    # row lock taken here...",
    "        if row.stock < 1:",
    "            raise SoldOut()",
    "",
    "        payment_api.charge(user_id, row.price)",
    "                    #! ...held through a ~200 ms network call",
    "",
    '        db.execute("UPDATE products SET stock = stock - 1 "',
    '                   "WHERE id = %s", product_id)',
    '        db.execute("INSERT INTO orders ...")',
  ], 11.5);
  s.addText("THROUGHPUT ON ONE HOT ROW", { x: 7.4, y: 1.8, w: 5.4, h: 0.26, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 9.5, color: TEAL_TINT, charSpacing: 0.5 });
  tbl(s, 7.4, 2.15, [2.4, 1.5, 1.5], [
    ["Lock held for", "Max / second", "1,000 drain in"],
    ["200 ms (payment inside)", "5", "200 s"],
    ["20 ms", "50", "20 s"],
    ["2 ms", "500", "2 s"],
    ["0.2 ms (atomic; illustrative)", "5,000", "0.2 s"],
  ], { dark: true, fs: 12, rowH: [0.45, 0.6, 0.6, 0.6, 0.7] });
  callout(s, 7.4, 5.2, 5.45, 1.5, "Speed = 1 / time the lock is held.", "Nine servers or ninety: the same five buyers per second get through. The rest are not idle. They are waiting, and waiting costs something.", true);
  footer(s, n, true);
});

// cascade timeline
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · The hot row", "How one slow row can take down the site", false);
  const rows = [
    ["0 s", "The sale opens", "1,000 buy requests arrive across 9 servers. One gets the lock and 999 wait. Each holds a thread and a database connection."],
    ["2 s", "The connection pool fills", "Say it holds 100. Now requests that have nothing to do with the sale, like browsing, login and the cart, cannot get a connection either."],
    ["5 s", "Timeouts trigger retries", "The old requests are still running on the server, waiting for the same lock. Every retry adds another waiting request instead of removing one."],
    ["10 s", "Health checks start failing", "They use the database too. The balancer removes \"unhealthy\" servers, so the others get more load and fail faster."],
    ["15 s", "The failure spreads", "Every service that calls checkout is stuck waiting for an answer that never comes. Day 7 studies this cascade, and here one hot row started it."],
  ];
  rows.forEach((r, i) => {
    const y = 1.65 + i * 1.04;
    card(s, 0.5, y, 12.35, 0.92, WHITE, { shadow: false, line: LINE });
    s.addText(r[0], { x: 0.7, y, w: 1.0, h: 0.92, isTextBox: true, margin: 0, valign: "middle", fontFace: CODE, bold: true, fontSize: 18, color: RUST });
    s.addText([
      { text: r[1] + "  ", options: { bold: true, color: INK, fontFace: HEAD, fontSize: 14 } },
      { text: r[2], options: { color: INK_SOFT, fontFace: BODY, fontSize: 12 } },
    ], { x: 1.85, y, w: 10.8, h: 0.92, isTextBox: true, margin: 0, valign: "middle" });
  });
  footer(s, n, false);
});

// hot row: fixed
add((n) => {
  const s = newSlide(true);
  header(s, "Part 1 · The hot row", "Six fixes, in the order to try them", true);
  codeBlock(s, 0.5, 1.6, 6.3, 5.25, "buy.py — after", TEAL_TINT, [
    "# at pool setup: SET lock_timeout = '300ms'",
    "#   wait briefly, then fail fast",
    "",
    "def buy(product_id, user_id):",
    "    reserved = db.execute(",
    '        "UPDATE products SET stock = stock - 1 "',
    '        "WHERE id = %s AND stock > 0", product_id)',
    "                  #! check + decrement in ONE statement",
    "    if reserved == 0:",
    "        raise SoldOut()",
    "",
    "    try:",
    "        payment_api.charge(user_id, price)",
    "                  #! slow call, no DB lock held",
    "    except PaymentFailed:",
    '        db.execute("UPDATE products SET stock = stock + 1 "',
    '                   "WHERE id = %s", product_id)',
    "        raise      #! give the unit back",
    "",
    '    db.execute("INSERT INTO orders ...")',
  ], 11);
  const moves = [
    ["Hold the lock for less time", "Slow calls happen outside the lock."],
    ["One-step conditional update", "Check and change in a single statement."],
    ["Limit the waiting", "lock_timeout, small pool per endpoint; fail fast."],
    ["Optimistic locking", "Good when conflicts are rare; with 1,000 buyers almost everyone retries."],
    ["Split the hot value", "Ten stock rows: waiting drops 10 times."],
    ["Keep it off the database", "Redis DECR, or put it on a queue (Part 2)."],
  ];
  moves.forEach((m, i) => {
    const y = 1.6 + i * 0.88;
    circle(s, 7.15, y + 0.08, 0.42, TEAL, String(i + 1), { fontSize: 12 });
    s.addText([
      { text: m[0], options: { bold: true, color: WHITE, fontFace: BODY, fontSize: 13.5, breakLine: true } },
      { text: m[1], options: { color: "A9C2C5", fontFace: BODY, fontSize: 11.5 } },
    ], { x: 7.75, y, w: 5.1, h: 0.8, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, true);
});

// ---- pessimistic vs optimistic locking (sequence diagrams)
const SEQ = { A: 1.6, DB: 4.5, B: 7.4 };
function seqBase(s, bottom) {
  card(s, 0.5, 1.75, 8.1, 5.2, WHITE, { shadow: false, line: LINE });
  [["Buyer A", SEQ.A, TEAL_TINT, TEAL], ["Database", SEQ.DB, "E3EAEC", SLATE], ["Buyer B", SEQ.B, PLUM_TINT, PLUM]].forEach(([t, x, f, c]) => {
    box(s, x - 0.75, 1.88, 1.5, 0.45, f, c, 1.5);
    label(s, t, x - 0.75, 1.88, 1.5, 0.45, { size: 11, bold: true });
    s.addShape(pptx.ShapeType.line, { x, y: 2.33, w: 0, h: bottom - 2.33, line: { color: INK_FAINT, width: 1, dashType: "dash" } });
  });
}
function seqMsg(s, from, to, y, text, o = {}) {
  const dir = to > from ? 1 : -1;
  arrow(s, from + dir * 0.06, y, to - dir * 0.06, y, o.color || INK_SOFT, { w: 1.4, dash: o.dash ? "dash" : "solid" });
  label(s, text, Math.min(from, to) + 0.05, y - 0.27, Math.abs(to - from) - 0.1, 0.25, { size: o.size || 9.5, color: o.textColor || INK, bold: !!o.bold });
}
function seqBar(s, x, y0, y1, fill, line) {
  s.addShape(pptx.ShapeType.rect, { x: x - 0.09, y: y0, w: 0.18, h: y1 - y0, fill: { color: fill }, line: { color: line, width: 1 } });
}
function sidePanel(s, items) {
  items.forEach((it, i) => {
    const y = 1.75 + i * 1.75;
    card(s, 8.85, y, 4.0, 1.6, WHITE);
    s.addShape(pptx.ShapeType.rect, { x: 8.85, y, w: 0.09, h: 1.6, fill: { color: it[0] }, line: { type: "none" } });
    s.addText([
      { text: it[1], options: { bold: true, color: INK, fontFace: HEAD, fontSize: 14, breakLine: true, paraSpaceAfter: 4 } },
      { text: it[2], options: { color: INK_SOFT, fontFace: BODY, fontSize: 12 } },
    ], { x: 9.15, y: y + 0.1, w: 3.55, h: 1.45, isTextBox: true, margin: 0, valign: "top" });
  });
}

add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Locking", "Pessimistic locking: lock first, then work", false, { sub: "Buyer B has to wait until Buyer A is completely done." });
  seqBase(s, 6.85);
  const y = [2.8, 3.4, 4.0, 4.6, 5.2, 5.8, 6.4];
  seqBar(s, SEQ.DB, y[0], y[4], RUST_TINT, RUST);
  seqBar(s, SEQ.B, y[2], y[5], AMBER_TINT, AMBER);
  seqMsg(s, SEQ.A, SEQ.DB, y[0], "SELECT … FOR UPDATE");
  seqMsg(s, SEQ.DB, SEQ.A, y[1], "stock = 5", { dash: true });
  seqMsg(s, SEQ.B, SEQ.DB, y[2], "SELECT … FOR UPDATE");
  seqMsg(s, SEQ.A, SEQ.DB, y[3], "UPDATE stock = 4");
  seqMsg(s, SEQ.A, SEQ.DB, y[4], "COMMIT → lock released", { bold: true, textColor: RUST });
  seqMsg(s, SEQ.DB, SEQ.B, y[5], "lock granted: stock = 4", { dash: true, bold: true, textColor: TEAL });
  seqMsg(s, SEQ.B, SEQ.DB, y[6], "UPDATE stock = 3 + COMMIT");
  label(s, "row locked", SEQ.DB + 0.15, y[2] + 0.1, 1.3, 0.25, { size: 9, color: RUST, bold: true, align: "left" });
  label(s, "B waits…", SEQ.B + 0.14, y[3], 0.95, 0.3, { size: 9, color: AMBER, bold: true, align: "left" });
  sidePanel(s, [
    [RUST, "A takes the lock", "The first SELECT … FOR UPDATE locks the row. Nobody else can change it, or even lock it, until A commits."],
    [AMBER, "B waits and does nothing", "B asked for the same lock, so it is stuck in line. The wait is the cost: 1,000 buyers means 999 waiting."],
    [TEAL, "Safe and simple", "When B gets the lock it reads the new value (4), so it can never overwrite A's change."],
  ]);
  footer(s, n, false);
  s.addNotes("Read the diagram top to bottom: time flows down. The red bar on the database is the row lock. The amber bar on Buyer B is the time B spends waiting.\n\nKey point: pessimistic locking prevents the conflict by making buyers take turns. It is safe, but the waiting is the price, and a lock held across a slow call (like a payment) makes the line very long.");
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Locking", "Optimistic locking: work first, check at the end", false, { sub: "Nobody waits. The buyer who loses finds out when it tries to save." });
  seqBase(s, 6.85);
  const y = []; for (let i = 0; i < 10; i++) y.push(2.75 + i * 0.43);
  seqMsg(s, SEQ.A, SEQ.DB, y[0], "SELECT stock, version");
  seqMsg(s, SEQ.DB, SEQ.A, y[1], "stock = 5, version = 7", { dash: true });
  seqMsg(s, SEQ.B, SEQ.DB, y[2], "SELECT stock, version");
  seqMsg(s, SEQ.DB, SEQ.B, y[3], "stock = 5, version = 7", { dash: true });
  seqMsg(s, SEQ.A, SEQ.DB, y[4], "UPDATE … WHERE version = 7");
  seqMsg(s, SEQ.DB, SEQ.A, y[5], "1 row updated (now version 8)", { dash: true, bold: true, textColor: TEAL });
  seqMsg(s, SEQ.B, SEQ.DB, y[6], "UPDATE … WHERE version = 7");
  seqMsg(s, SEQ.DB, SEQ.B, y[7], "0 rows updated: A got there first", { dash: true, bold: true, textColor: RUST });
  seqMsg(s, SEQ.B, SEQ.DB, y[8], "retry: read, then WHERE version = 8", { size: 9 });
  seqMsg(s, SEQ.DB, SEQ.B, y[9], "1 row updated (version 9)", { dash: true, bold: true, textColor: TEAL });
  label(s, "no lock held", SEQ.DB + 0.15, y[0] - 0.15, 1.3, 0.25, { size: 9, color: TEAL, bold: true, align: "left" });
  sidePanel(s, [
    [TEAL, "Both read at the same time", "No lock, no waiting. A and B both see stock 5 and version 7."],
    [RUST, "The loser finds out at save time", "A saves first. B's UPDATE says WHERE version = 7, but the row is now version 8, so 0 rows change."],
    [AMBER, "B reads again and retries", "The cost is repeated work. With 1,000 buyers on one row, most attempts are wasted."],
  ]);
  footer(s, n, false);
  s.addNotes("Same two buyers, same row. The difference: nobody holds a lock. The version number is the check. The database makes the check and the update in one step, so two buyers cannot both succeed with the same version.\n\nB is not blocked, but it did its work twice. When conflicts are rare that is a good trade. When 1,000 buyers fight over one row it is not: about 500,000 attempts for 1,000 sales in the worst case.");
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · Locking", "Pessimistic vs. optimistic: which one when?", false);
  tbl(s, 0.5, 1.6, [2.7, 4.8, 4.85], [
    ["", "Pessimistic (lock first)", "Optimistic (check at the end)"],
    ["The idea", "Take a lock so nobody else can change the row", "Don't lock. Check that the row didn't change when you save"],
    ["Does anyone wait?", "Yes. Everyone else waits for the lock", "No. The loser retries instead"],
    ["Conflicts are rare", "You pay for locking you didn't need", "Fast. Almost always works the first time"],
    ["Conflicts are common (hot row)", "A line forms, but each buyer succeeds once", "Many retries and a lot of wasted work"],
    ["Long gap between read and save (forms)", "Bad. The lock would be held for minutes", "Good. Nothing is held while the person types"],
    ["Typical use", "Stock, counters, moving money", "Editing profiles, documents, carts"],
  ], { fs: 13.5, rowH: [0.45, 0.62, 0.55, 0.62, 0.7, 0.7, 0.55] });
  callout(s, 0.5, 5.95, 12.35, 0.95, "Rule of thumb:", "if many people fight over the same row, make them take turns (lock or one-step update). If conflicts are rare and the work takes a while, check at the end.", false);
  footer(s, n, false);
  s.addNotes("The one-step conditional update from the fixes slide (UPDATE ... SET stock = stock - 1 WHERE stock > 0) sits between the two: its WHERE clause is a tiny optimistic check, but it needs no version column and no retry loop because the check and the change are one statement.\n\nHTTP's ETag / If-Match headers are optimistic locking for web APIs: a 412 Precondition Failed means someone saved first.");
});

checkIn(3, 3, "How fast can one row go?", "Many buyers want the same product row. The row lock decides how fast they can buy.", [
  "A buyer holds the lock for 40 ms. How many buyers per second can the row handle? How long do 1,000 buyers take?",
  "Now the lock is held for only 5 ms. Redo both numbers.",
  "We add 10 more app servers. What happens to the speed?",
  "A product has a like counter and a stock number. Many people update both at once. Which of the six fixes would you use for each?",
], null, "1. 1 second / 0.040 s = 25 buyers per second. 1,000 / 25 = 40 seconds.\n2. 1 / 0.005 = 200 per second. 1,000 / 200 = 5 seconds.\n3. Nothing changes. Only one buyer can hold the lock at a time, so more servers just make a longer line.\n4. Likes do not need to be exact: count them in Redis and save in batches, or spread them over several rows (fixes 5 and 6). Stock must be exact: use the one-step update (fix 2) and a short lock_timeout (fix 3).");

// does 2x machines give 2x work
add((n) => {
  const s = newSlide(false);
  header(s, "Part 1 · The limit", "Does 2× the machines give 2× the work?", false, { sub: "Speedup = how many times more work per second N machines do than one machine. Perfect scaling means speedup = N." });
  s.addText("Your checkout servers are close to the best case: every request is independent, so doubling the servers roughly doubles the requests per second. That holds until the servers depend on something they share, like the hot row you just saw. That breaks the straight line in two ways.", { x: 0.5, y: 1.85, w: 12.3, h: 1.0, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK_SOFT, valign: "top" });
  [
    { c: TEAL, t: "1. Work that can't be split", p: ["Some step must happen one at a time, in order, on one machine: a lock, a final summary write, a single main database.", "That is the serial part. Extra machines cannot speed it up, no matter how many you add."] },
    { c: PLUM, t: "2. Machines having to talk", p: ["Machines that share data must coordinate: clear each other's caches, agree on a lock, copy a write.", "This talking takes time, and it grows with every machine you add."] },
  ].forEach((c, i) => {
    const x = 0.5 + i * 6.3;
    card(s, x, 3.05, 6.05, 3.7, WHITE);
    s.addShape(pptx.ShapeType.rect, { x, y: 3.05, w: 6.05, h: 0.09, fill: { color: c.c }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.3, y: 3.3, w: 5.5, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: INK });
    s.addText(bl(c.p, { fs: 13.5, color: INK_SOFT, gap: 10 }), { x: x + 0.3, y: 3.95, w: 5.5, h: 2.7, isTextBox: true, margin: 0, valign: "top" });
  });
  s.addText("We put numbers on this in Part 2, using a nightly job.", { x: 0.5, y: 6.82, w: 12.3, h: 0.25, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT });
  footer(s, n, false);
});

// ============================================================ PART 2
divider(2, "0:55–1:45", "Scaling backend jobs with queues", "Not all work is a web request. Nightly jobs, emails and reports run in the background. First we see how much extra machines really help a job, then how a queue hands the work out.");

// loyalty job code
add((n) => {
  const s = newSlide(true);
  header(s, "Part 2 · Splitting a job", "A nightly job: one machine, then many", true, { sub: "A nightly job recalculates every customer's loyalty tier, then writes one summary table. On one machine it takes 100 minutes." });
  codeBlock(s, 0.5, 1.8, 6.1, 3.7, "loyalty_job.py — one machine", TEAL_TINT, [
    "def recalculate_loyalty_tiers(customers):",
    "    tiers = {}",
    "",
    "    # part 1: per-customer work, all independent",
    "    for customer in customers:",
    "        tiers[customer.id] = compute_tier(customer.orders)",
    "                          # ~95 min, all customers",
    "",
    "    # part 2: one table, one lock, one writer",
    "    with summary_table.lock():",
    "        summary_table.write_all(tiers)    # ~5 min",
    "        summary_table.write_totals(tiers)",
    "                          # needs every tier first",
  ], 10.5);
  codeBlock(s, 6.75, 1.8, 6.1, 3.7, "loyalty_job.py — N machines", TEAL_TINT, [
    "def tier_chunk(chunk):         # runs on every machine",
    "    return {c.id: compute_tier(c.orders) for c in chunk}",
    "",
    "def recalculate_loyalty_tiers(customers, machines):",
    "    chunks = split(customers, parts=len(machines))",
    "    partials = run_on_each(machines, tier_chunk, chunks)",
    "                          # ~95 / N min",
    "    tiers = merge(partials)",
    "",
    "    with summary_table.lock():     # still one writer",
    "        summary_table.write_all(tiers)    # still ~5 min",
    "        summary_table.write_totals(tiers)",
  ], 10.5);
  callout(s, 0.5, 5.7, 12.35, 1.15, "Total time = 95 / N + 5 minutes.", "4 machines: 28.75 min. 20 machines: 9.75 min. 1,000 machines: about 5.1 min. The 5 never changes; only the 95 gets smaller. That 5 out of 100 is what \"5% serial\" means.", true);
  footer(s, n, true);
});

// ideal vs 5% serial
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Splitting a job", "Why each extra machine helps less", false, { sub: "The 5% of the job that cannot be split stays the same size, so every new machine adds a little less than the one before." });
  tbl(s, 0.5, 1.95, [2.6, 3.0, 3.4, 3.35], [
    ["Instances (N)", "Ideal speedup", "With 5% serial work", "Share of what you paid for"],
    ["2", "2×", "1.9×", "95%"],
    ["4", "4×", "3.5×", "87%"],
    ["8", "8×", "5.9×", "74%"],
    ["16", "16×", "9.1×", "57%"],
    ["32", "32×", "12.6×", "39%"],
  ], { fs: 14, rowH: [0.5, 0.55, 0.55, 0.55, 0.55, 0.55] });
  s.addText("Speedup(N) = 1 / ( s + (1 − s) / N )     —  Amdahl's Law, with s the serial fraction", { x: 0.5, y: 5.55, w: 12.3, h: 0.4, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: TEAL });
  s.addText("Live web traffic hides this, because you cannot time \"the serial part\" of live requests. In checkout it is the lock on the hot row. A batch job lets us time every piece, so we use it as the example.", { x: 0.5, y: 6.05, w: 12.3, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

// amdahl vs usl chart
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Splitting a job", "More machines help less, then hurt", false);
  // native plot: x = instances 1..32, y = speedup 0..20
  const X0 = 1.3, Y0 = 6.15, PW = 6.9, PH = 4.2;
  const px = (k) => X0 + (k - 1) * PW / 31, py = (v) => Y0 - v * PH / 20;
  [0, 5, 10, 15, 20].forEach((v) => {
    s.addShape(pptx.ShapeType.line, { x: X0, y: py(v), w: PW, h: 0, line: { color: v === 0 ? INK_SOFT : LINE, width: v === 0 ? 1.25 : 0.75 } });
    label(s, String(v), X0 - 0.55, py(v) - 0.15, 0.45, 0.3, { size: 10, color: INK_FAINT, align: "right" });
  });
  s.addShape(pptx.ShapeType.line, { x: X0, y: py(20), w: 0, h: PH, line: { color: INK_SOFT, width: 1.25 } });
  [1, 8, 16, 24, 32].forEach((k) => label(s, String(k), px(k) - 0.3, Y0 + 0.08, 0.6, 0.28, { size: 10, color: INK_FAINT }));
  label(s, "instances (N)", X0 + PW / 2 - 1, Y0 + 0.38, 2, 0.3, { size: 10, color: INK_SOFT });
  label(s, "speedup", 0.35, py(20) - 0.4, 1.2, 0.3, { size: 10, color: INK_SOFT, align: "left" });
  s.addShape(pptx.ShapeType.line, { x: X0, y: py(20), w: PW, h: 0, line: { color: TEAL, width: 1, dashType: "dash" } });
  label(s, "asymptote: 1/s = 20", X0 + PW - 2.2, py(20) - 0.3, 2.2, 0.28, { size: 9.5, color: TEAL, align: "right" });
  const amdahl = (k) => 1 / (0.05 + 0.95 / k);
  const usl = (k) => k / (1 + 0.02 * (k - 1) + 0.0015 * k * (k - 1));
  const plot = (f, color) => {
    for (let k = 1; k < 32; k++) arrow(s, px(k), py(f(k)), px(k + 1), py(f(k + 1)), color, { w: 3, noHead: true });
  };
  plot(amdahl, TEAL); plot(usl, PLUM);
  s.addShape(pptx.ShapeType.ellipse, { x: px(24) - 0.09, y: py(usl(24)) - 0.09, w: 0.18, h: 0.18, fill: { color: PLUM }, line: { type: "none" } });
  label(s, "USL peaks near N = 24, then declines", px(13), py(6.5), px(32) - px(13), 0.3, { size: 10.5, color: PLUM, bold: true, align: "right" });
  label(s, "Amdahl (5% serial)", px(20), py(amdahl(20)) - 0.55, 2.4, 0.3, { size: 10, color: TEAL, bold: true, align: "left" });
  card(s, 8.8, 1.7, 4.05, 5.0, WHITE);
  s.addText([
    { text: "Amdahl's Law", options: { bold: true, color: TEAL, fontFace: HEAD, fontSize: 15, breakLine: true } },
    { text: "Speedup is limited by the part that cannot be split up. With 5% serial work, nothing ever beats 20 times.", options: { color: INK_SOFT, fontFace: BODY, fontSize: 12, breakLine: true, paraSpaceAfter: 12 } },
    { text: "Universal Scalability Law", options: { bold: true, color: PLUM, fontFace: HEAD, fontSize: 15, breakLine: true } },
    { text: "Adds the cost of machines talking to each other (clearing caches, fighting over locks, copying data). Past a point that cost grows faster than the useful work, so it peaks near N = 24 and then goes down.", options: { color: INK_SOFT, fontFace: BODY, fontSize: 12, breakLine: true, paraSpaceAfter: 12 } },
    { text: "Adding machines can make the system slower. Both curves are real. Real systems behave like the USL.", options: { color: INK, bold: true, fontFace: BODY, fontSize: 12 } },
  ], { x: 9.05, y: 1.85, w: 3.6, h: 4.7, isTextBox: true, margin: 0, valign: "top" });
  footer(s, n, false);
  s.addNotes("Spend your time on the main idea (the cost of machines coordinating can make more machines actively worse), not on deriving formulas. Nobody needs to compute kappa and sigma by hand. The USL curve here uses sigma = 0.02 and kappa = 0.0015, which peaks near N = 24 at roughly 10.5x.");
});

checkIn(4, 2, "Will 50 machines be 50 times faster?", "A nightly report takes 200 minutes on one machine. 10 of those minutes are a final step that cannot be split. The other 190 minutes can be shared across machines. So: time = 10 + 190 / machines.", [
  "What percent of the job cannot be split?",
  "How long does it take on 10 machines? On 50?",
  "Your boss wants it 25 times faster (8 minutes). Is that possible?",
], null, "1. 10 / 200 = 5%.\n2. 10 machines: 10 + 190/10 = 29 minutes (about 7 times faster). 50 machines: 10 + 190/50 = 13.8 minutes (about 14.5 times faster). Five times more machines gave only about twice the speedup.\n3. No. The step that cannot be split already takes 10 minutes, which is more than 8. Even with endless machines the best is 10 minutes (20 times faster). To go faster, make the 10-minute step smaller.");


add((n) => {
  const s = newSlide(true);
  header(s, "Part 2 · Queues", "Why put a queue in the middle", true, { sub: "In the job above, run_on_each() hid how chunks reach machines. A queue does that job, and lets a request finish before slow work is done." });
  codeBlock(s, 0.5, 1.8, 6.1, 4.0, "checkout.py — before", AMBER_LIGHT, [
    "def place_order(order):",
    "    save_order(order)",
    "    send_receipt_email(order)       # ~300 ms, third party",
    "    update_loyalty_points(order)    # ~150 ms",
    "    notify_warehouse(order)         # ~400 ms",
    '    return "ok"',
    "        #! ~850 ms later — and a failure in any",
    "        #! step fails the order",
  ], 11.5);
  codeBlock(s, 6.75, 1.8, 6.1, 4.0, "checkout.py — after", TEAL_TINT, [
    "def place_order(order):",
    "    save_order(order)",
    '    queue.publish("order-placed", {"order_id": order.id})',
    "                                    # ~2 ms",
    '    return "ok"',
    "",
    "# on any number of worker machines",
    "def loyalty_worker():",
    '    for message in queue.consume("order-placed"):',
    '        update_loyalty_points(message["order_id"])',
  ], 11.5);
  s.addText("Day 3's Observer pattern was this same idea inside one program. A queue is Observer across programs, with durability and some new problems.", { x: 0.5, y: 6.1, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 13, color: "A9C2C5", valign: "top" });
  footer(s, n, true);
});

cards3("Part 2 · Queues", "What a queue buys — and what it doesn't", "A queue absorbs spikes, not constant overload.", [
  { color: TEAL, title: "Separates timing", points: ["The request returns when the work is accepted, not finished.", "A slow email provider no longer slows checkout."] },
  { color: PLUM, title: "Absorbs spikes", points: ["A burst of orders becomes a backlog that workers clear at their own pace.", "Workers run at a steady load instead of the traffic peak."] },
  { color: AMBER, title: "Fails and scales on its own", points: ["A crashed worker does not fail the order; its message goes to another worker.", "Add workers without touching the web servers."] },
], "If 6,000 messages arrive per second and workers finish 800, the backlog grows by 5,200 every second, over three million messages in ten minutes. The queue did not remove the waiting; it only moved it somewhere safer.", false, ["Spikes, not constant overload.", "6,000 in per second and 800 out: the backlog grows 5,200/s, over three million messages in ten minutes. You still need to add workers, drop load at the source, or accept a growing delay."]);

checkIn(5, 2, "How big does the backlog get?", "A promo lasts 10 minutes and sends 1,200 messages per second to a queue. You have 3 workers. Each worker handles 300 messages per second. After the promo, 400 messages per second arrive.", [
  "How many messages are waiting when the promo ends?",
  "How long until the queue is empty?",
  "How many workers would keep up during the promo?",
], null, "1. Workers handle 3 x 300 = 900 per second, but 1,200 arrive. That is 300 extra per second for 600 seconds = 180,000 messages waiting.\n2. After the promo, 900 are handled and 400 arrive, so the queue shrinks by 500 per second. 180,000 / 500 = 360 seconds, about 6 minutes.\n3. 1,200 / 300 = 4 workers (5 if you leave spare room, as in the capacity plan).\n\nA queue helps with short spikes. If the overload never ends, the backlog grows forever.");

add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Scaling the queue", "Scaling a queue: more workers, then partitions", false);
  [
    { c: TEAL, t: "1. More consumers on one queue", p: ["Each message goes to exactly one worker (competing consumers), so speed grows with the number of workers. This is the \"any server will do\" idea from Part 1, applied to jobs.", "It stops being enough: a single queue still runs on one machine.", "And once order matters, several workers pulling from one line cannot keep the order."] },
    { c: PLUM, t: "2. Split the queue: partitions", p: ["Instead of one line there are P separate lines (partitions), possibly on different machines.", "The sender picks the partition by hashing a key, so every message with the same key goes to the same partition, in order.", "Each partition is read by one consumer at a time. Speed grows with P, and order is kept, but only inside a partition."] },
  ].forEach((c, i) => {
    const x = 0.5 + i * 6.3;
    card(s, x, 1.65, 6.05, 5.1, WHITE);
    s.addShape(pptx.ShapeType.rect, { x, y: 1.65, w: 6.05, h: 0.09, fill: { color: c.c }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.3, y: 1.9, w: 5.5, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: INK });
    s.addText(bl(c.p, { fs: 14, color: INK_SOFT, gap: 12 }), { x: x + 0.3, y: 2.6, w: 5.5, h: 4.0, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, false);
});

// Kafka diagram
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Kafka", "Kafka: a queue split into partitions", false);
  box(s, 0.5, 3.35, 1.7, 1.1, WHITE, INK_SOFT); label(s, "Producer", 0.5, 3.45, 1.7, 0.45, { size: 12, bold: true }); label(s, "key = customer_id", 0.5, 3.9, 1.7, 0.35, { size: 9, color: INK_FAINT });
  label(s, "hash(key) picks one", 0.5, 3.0, 1.7, 0.3, { size: 9, color: AMBER, bold: true });
  label(s, "TOPIC: order-events", 3.9, 1.7, 4.2, 0.3, { size: 11, bold: true, color: TEAL });
  const lanesY = [2.1, 3.5, 4.9];
  lanesY.forEach((y, i) => {
    box(s, 3.2, y, 5.7, 1.0, WHITE, INK_SOFT);
    label(s, "P" + i, 3.3, y + 0.3, 0.5, 0.4, { size: 13, bold: true, align: "left" });
    const cells = i === 2 ? 4 : 6;
    for (let c = 0; c < cells; c++) {
      const lag = (i < 2 && c >= (i === 0 ? 4 : 5));
      s.addShape(pptx.ShapeType.roundRect, { x: 3.85 + c * 0.82, y: y + 0.27, w: 0.7, h: 0.46, rectRadius: 0.05, fill: { color: lag ? AMBER_TINT : TEAL_TINT }, line: { color: lag ? AMBER : TEAL, width: 1 } });
      label(s, String(c), 3.85 + c * 0.82, y + 0.27, 0.7, 0.46, { size: 11 });
    }
    arrow(s, 2.2, 3.9, 3.15, y + 0.5, INK_SOFT);
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 9.5, y: 1.95, w: 3.35, h: 4.15, rectRadius: 0.1, fill: { type: "none" }, line: { color: PLUM, width: 1.25, dashType: "dash" } });
  label(s, "GROUP: loyalty-workers", 9.5, 1.65, 3.35, 0.3, { size: 10, color: PLUM, bold: true });
  box(s, 9.8, 2.35, 2.75, 1.25, PLUM_TINT, PLUM, 1.5); label(s, "consumer 1", 9.8, 2.5, 2.75, 0.45, { size: 12, bold: true }); label(s, "reads P0 and P1", 9.8, 2.95, 2.75, 0.4, { size: 10, color: INK_SOFT });
  box(s, 9.8, 4.45, 2.75, 1.25, PLUM_TINT, PLUM, 1.5); label(s, "consumer 2", 9.8, 4.6, 2.75, 0.45, { size: 12, bold: true }); label(s, "reads P2", 9.8, 5.05, 2.75, 0.4, { size: 10, color: INK_SOFT });
  arrow(s, 8.9, 2.6, 9.75, 2.85, INK_SOFT);
  arrow(s, 8.9, 4.0, 9.75, 3.15, INK_SOFT);
  arrow(s, 8.9, 5.4, 9.75, 5.1, INK_SOFT);
  s.addShape(pptx.ShapeType.roundRect, { x: 3.2, y: 6.3, w: 0.28, h: 0.28, rectRadius: 0.04, fill: { color: TEAL_TINT }, line: { color: TEAL, width: 1 } });
  label(s, "processed (offset committed)", 3.55, 6.28, 2.8, 0.32, { size: 9.5, color: INK_FAINT, align: "left" });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.4, y: 6.3, w: 0.28, h: 0.28, rectRadius: 0.04, fill: { color: AMBER_TINT }, line: { color: AMBER, width: 1 } });
  label(s, "waiting — the consumer lag", 6.75, 6.28, 2.8, 0.32, { size: 9.5, color: INK_FAINT, align: "left" });
  s.addText("A third consumer would get one partition of its own; a fourth would sit idle. The number of partitions is the limit on parallel work.", { x: 0.5, y: 6.62, w: 12.3, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12, color: INK_SOFT });
  footer(s, n, false);
});

cards3("Part 2 · Kafka", "Kafka in three ideas", "Messages are kept in a log, and partitions control order and speed.", [
  { color: TEAL, title: "A log, not a queue", points: ["Messages are added to the end and kept (for a set time or size), not deleted when read.", "A consumer just remembers its position (offset). It can go back and replay.", "Several consumer groups can read the same topic independently."] },
  { color: PLUM, title: "Order and parallelism", points: ["Key, then hash, then partition: the same key always goes to the same partition, in order.", "Order is guaranteed only inside one partition, never across partitions.", "Number of partitions = the most consumers that can work in parallel in a group."] },
  { color: AMBER, title: "Staying up", points: ["Each partition is copied to other servers. One copy leads and the rest follow.", "acks=all waits until the in-sync copies have the message before confirming the write.", "A group coordinator hands out partitions again when a consumer joins or dies (a rebalance)."] },
]);

// Kafka offsets
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Kafka", "Offset: a consumer's place in the log", false, { sub: "Every message in a partition has a number. Each consumer group saves one number per partition, like a bookmark." });
  box(s, 0.5, 1.95, 12.35, 1.15, WHITE, INK_SOFT, 1.3);
  label(s, "P0", 0.6, 2.2, 0.55, 0.65, { size: 13, bold: true, align: "left" });
  for (let c = 0; c < 10; c++) {
    const done = c <= 5, x = 1.3 + c * 1.1;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.2, w: 1.0, h: 0.65, rectRadius: 0.05, fill: { color: done ? TEAL_TINT : AMBER_TINT }, line: { color: done ? TEAL : AMBER, width: 1 } });
    label(s, String(c), x, 2.2, 1.0, 0.65, { size: 14, bold: true });
  }
  label(s, "offset numbers: 0, 1, 2, … never change, never reused", 1.3, 1.62, 6.0, 0.3, { size: 10, color: INK_FAINT, align: "left" });
  const sx = 1.3 + 5 * 1.1 + 0.5, nx = 1.3 + 9 * 1.1 + 0.5;
  arrow(s, sx, 3.95, sx, 3.15, TEAL, { w: 2 });
  label(s, "saved offset = 5", sx - 1.6, 4.0, 3.2, 0.3, { size: 11.5, bold: true, color: TEAL });
  label(s, "the bookmark: done up to here", sx - 1.6, 4.3, 3.2, 0.3, { size: 9.5, color: INK_SOFT });
  arrow(s, nx, 3.95, nx, 3.15, AMBER, { w: 2 });
  label(s, "newest offset = 9", nx - 1.4, 4.0, 2.8, 0.3, { size: 11.5, bold: true, color: AMBER });
  label(s, "the last message written", nx - 1.4, 4.3, 2.8, 0.3, { size: 9.5, color: INK_SOFT });
  label(s, "lag = 9 − 5 = 4 messages waiting", 7.9, 3.3, 3.0, 0.3, { size: 11, bold: true, color: RUST });
  s.addShape(pptx.ShapeType.line, { x: sx + 0.15, y: 3.65, w: nx - sx - 0.3, h: 0, line: { color: RUST, width: 1.5, dashType: "dash" } });
  const opts = [
    [0.5, TEAL, TEAL_TINT, "WORK FIRST, THEN SAVE  (what we do)", "Read message 6, do the work, then save offset 6.", "Crash in between: after the restart, message 6 comes again. A message may repeat, but none is lost."],
    [6.8, RUST, RUST_TINT, "SAVE FIRST, THEN WORK  (avoid)", "Read message 6, save offset 6, then do the work.", "Crash in between: after the restart, Kafka continues from 7. Nothing repeats, but message 6 is lost."],
  ];
  opts.forEach((o) => {
    box(s, o[0], 4.85, 6.05, 1.95, o[2], o[1], 1.5);
    label(s, o[3], o[0] + 0.2, 4.95, 5.65, 0.35, { size: 11, bold: true, color: o[1], align: "left" });
    s.addText(o[4], { x: o[0] + 0.2, y: 5.35, w: 5.65, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK, valign: "top" });
    s.addText(o[5], { x: o[0] + 0.2, y: 5.85, w: 5.65, h: 0.85, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK, valign: "top" });
  });
  footer(s, n, false);
});

add((n) => {
  const s = newSlide(true);
  header(s, "Part 2 · Kafka", "Handling the same message twice", true, { sub: "Do the work first, then save your position. If the consumer crashes in between, the message arrives again instead of being lost." });
  codeBlock(s, 0.5, 1.75, 12.35, 5.1, "loyalty_worker.py", TEAL_TINT, [
    "# producer side: the key decides the partition",
    "producer.produce(",
    '    "order-events",',
    "    key=str(order.customer_id),     # same customer -> same partition -> in order",
    "    value=json.dumps(event),",
    ")                                   # producer configured with acks=all",
    "",
    "# consumer side",
    "consumer = Consumer({",
    '    "group.id": "loyalty-workers",',
    '    "enable.auto.commit": False,    # we decide when a message counts as done',
    "})",
    'consumer.subscribe(["order-events"])',
    "",
    "while True:",
    "    msg = consumer.poll(1.0)",
    "    if msg is None:",
    "        continue",
    "    event = json.loads(msg.value())",
    '    if not already_processed(event["event_id"]):   # redelivery is normal; be idempotent',
    "        update_loyalty_points(event)",
    '        mark_processed(event["event_id"])',
    "    consumer.commit(msg)            # commit AFTER the work",
  ], 11);
  footer(s, n, true);
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · What goes wrong", "What can go wrong with queues (1 of 2)", false, { sub: "The queue sits between two programs that can each fail on their own. That causes all of these." });
  tbl(s, 0.5, 1.95, [2.1, 2.95, 3.2, 4.1], [
    ["Issue", "Why it happens", "General handling", "In Kafka"],
    ["Out-of-order processing", "Partitions are separate lines.", "Give everything that must stay in order the same key.", "Order holds inside a partition. Pick the key (customer_id) so related events share one."],
    ["Duplicate delivery", "A consumer crashes after doing the work but before confirming it.", "Make the consumer idempotent: running it twice has the same effect as once (skip events already seen, by their ID).", "Save the offset after processing = delivered at least once. \"Exactly once\" means at least once plus an idempotent consumer, or Kafka transactions."],
    ["Poison message (always fails)", "One message always fails. Retrying forever blocks everything behind it.", "Retry a limited number of times, then move it to a dead-letter queue.", "No built-in dead-letter queue: publish the failure to a .dlq topic and move past it."],
    ["Backlog (consumer lag)", "Producers send faster than consumers can process.", "Add consumers, alert on lag, slow down the producer.", "Lag = newest offset minus saved offset. Add consumers, but never more than the number of partitions."],
  ], { fs: 11, rowH: [0.45, 1.0, 1.25, 1.0, 1.15] });
  footer(s, n, false);
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · What goes wrong", "What can go wrong with queues (2 of 2)", false);
  tbl(s, 0.5, 1.6, [2.1, 2.95, 3.2, 4.1], [
    ["Issue", "Why it happens", "General handling", "In Kafka"],
    ["Hot partition", "Uneven keys: one huge customer sends most of the traffic.", "A better key, or add a random part to the key for that customer.", "One partition (and one consumer) stays overloaded while the rest sit idle. It is the Part 1 hot row again."],
    ["Rebalance pauses", "A consumer joins or leaves, partitions are reassigned, and processing pauses briefly.", "Keep the work per message short. Avoid consumers that keep restarting.", "Cooperative rebalancing and static membership shorten the pause."],
    ["Can't add partitions freely", "Which partition a key goes to depends on the number of partitions: the same trap as hash % N.", "Create more partitions than you need at the start.", "Adding partitions later changes where keys go, which breaks order for existing keys."],
  ], { fs: 12, rowH: [0.45, 1.15, 1.15, 1.15] });
  callout(s, 0.5, 5.65, 12.35, 1.2, "Kafka is not the only option.", "SQS or RabbitMQ delete a message once it is acknowledged and keep no replayable history. Choose the log (Kafka) when you need replay or several independent readers.", false);
  footer(s, n, false);
  s.addNotes("Talking points:\n- The batch job from Part 1 becomes: publish one message per chunk, let a consumer group drain them. The serial summary step is now the fan-in; it still can't be split, so the ceiling hasn't moved.\n- Duplicates aren't a Kafka bug to be fixed; they're the price of not losing messages. The only real question is whether the consumer can tolerate seeing one twice.\n- 'More consumers' and 'more partitions' are different knobs. Consumers beyond the partition count do nothing.");
});

// ---- seven queue problems, one slide each (problem and fix side by side)
const PXs = [0.5, 6.8], PYs = 2.05, PWs = 6.05, PHs = 3.95;
function issuePanel(s, i, text, color) {
  label(s, text.toUpperCase(), PXs[i] + 0.05, PYs - 0.3, PWs, 0.26, { size: 10, bold: true, color, align: "left" });
  card(s, PXs[i], PYs, PWs, PHs, WHITE, { shadow: false, line: color, lineW: 1.5 });
}
function chip(s, x, y, t, o = {}) {
  const w = o.w || 1.3, h = o.h || 0.4;
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06, fill: o.dashed ? { type: "none" } : { color: o.fill || TEAL_TINT }, line: { color: o.line || TEAL, width: 1, dashType: o.dashed ? "dash" : "solid" } });
  label(s, t, x, y, w, h, { size: o.size || 9.5, color: o.color || INK, bold: !!o.bold });
}
function laneBox(s, x, y, w, h, name, o = {}) {
  box(s, x, y, w, h, o.fill || WHITE, o.line || INK_SOFT, o.lw || 1.1);
  if (name) label(s, name, x + 0.08, y + (o.top ? 0.02 : 0), 0.6, o.top ? 0.3 : h, { size: 10.5, bold: true, align: "left" });
}
function flowSteps(s, ox, oy, steps) {
  const kinds = { neutral: [WHITE, INK_SOFT, INK], ok: [TEAL_TINT, TEAL, INK], bad: [RUST_TINT, RUST, RUST], crash: [AMBER_TINT, AMBER, AMBER] };
  steps.forEach((st, i) => {
    const y = oy + 0.25 + i * 0.74, k = kinds[st[1]];
    box(s, ox + 0.35, y, 5.35, 0.52, k[0], k[1], st[1] === "neutral" ? 1.1 : 1.8);
    label(s, st[0], ox + 0.45, y, 5.15, 0.52, { size: 10.5, color: k[2], bold: st[1] !== "neutral", font: BODY });
    if (i < steps.length - 1) arrow(s, ox + 3.02, y + 0.52, ox + 3.02, y + 0.72, INK_SOFT, { w: 1.1 });
  });
}
function issueSlide(num, title, sub, drawL, drawR, caption, labels) {
  add((n) => {
    const s = newSlide(false);
    header(s, "Part 2 · What can go wrong (" + num + " of 7)", title, false, { sub });
    issuePanel(s, 0, (labels && labels[0]) || "The problem", RUST);
    issuePanel(s, 1, (labels && labels[1]) || "The fix", TEAL);
    drawL(s, PXs[0], PYs);
    drawR(s, PXs[1], PYs);
    callout(s, 0.5, 6.15, 12.35, 0.8, "How it's handled:", caption, false);
    footer(s, n, false);
  });
}

// 1 out-of-order
issueSlide(1, "Out-of-order processing", "Customer 8842 places an order, pays, then cancels. Do the events arrive in that order?",
  (s, ox, oy) => {
    box(s, ox + 0.25, oy + 0.95, 1.25, 0.85, WHITE, INK_SOFT); label(s, "Producer", ox + 0.25, oy + 1.0, 1.25, 0.4, { size: 10.5, bold: true }); label(s, "no key", ox + 0.25, oy + 1.38, 1.25, 0.3, { size: 9, color: INK_FAINT });
    [["P0", "3 cancelled"], ["P1", "1 placed"], ["P2", "2 paid"]].forEach((l, k) => {
      const y = oy + 0.25 + k * 0.85;
      laneBox(s, ox + 2.2, y, 3.4, 0.62, l[0]); chip(s, ox + 3.2, y + 0.11, l[1], { w: 1.5, h: 0.4, fill: AMBER_TINT, line: AMBER });
      arrow(s, ox + 1.5, oy + 1.38, ox + 2.17, y + 0.31, INK_SOFT, { w: 1 });
    });
    label(s, "Processed in this order:", ox + 0.25, oy + 2.85, 5.4, 0.28, { size: 10, color: INK_SOFT, align: "left" });
    [["3 cancelled", 0.25], ["1 placed", 2.1], ["2 paid", 3.95]].forEach((c, k) => { chip(s, ox + c[1], oy + 3.2, c[0], { w: 1.55, h: 0.4, fill: RUST_TINT, line: RUST, color: RUST, bold: true }); if (k < 2) arrow(s, ox + c[1] + 1.57, oy + 3.4, ox + c[1] + 1.85, oy + 3.4, RUST, { w: 1.2 }); });
  },
  (s, ox, oy) => {
    box(s, ox + 0.25, oy + 0.95, 1.25, 0.85, WHITE, INK_SOFT); label(s, "Producer", ox + 0.25, oy + 1.0, 1.25, 0.4, { size: 10.5, bold: true }); label(s, "key = customer_id", ox + 0.1, oy + 1.38, 1.55, 0.3, { size: 8.5, color: TEAL, bold: true });
    ["P0", "P1", "P2"].forEach((nme, k) => {
      const y = oy + 0.25 + k * 0.85;
      laneBox(s, ox + 2.2, y, 3.4, 0.62, nme, { line: k === 1 ? TEAL : INK_FAINT, lw: k === 1 ? 1.8 : 1, fill: k === 1 ? TEAL_TINT : WHITE });
      if (k === 1) [["1 placed", 0.85], ["2 paid", 2.0], ["3 cancelled", 2.95]].slice(0, 3).forEach((c, j) => chip(s, ox + 2.8 + j * 0.85, y + 0.13, ["1 placed", "2 paid", "3 cancelled"][j], { w: 0.82, h: 0.36, size: 7.5, fill: WHITE }));
    });
    arrow(s, ox + 1.5, oy + 1.38, ox + 2.17, oy + 1.1 + 0.15, TEAL, { w: 1.6 });
    label(s, "Processed in this order:", ox + 0.25, oy + 2.85, 5.4, 0.28, { size: 10, color: INK_SOFT, align: "left" });
    [["1 placed", 0.25], ["2 paid", 2.1], ["3 cancelled", 3.95]].forEach((c, k) => { chip(s, ox + c[1], oy + 3.2, c[0], { w: 1.55, h: 0.4, fill: TEAL_TINT, line: TEAL, color: TEAL, bold: true }); if (k < 2) arrow(s, ox + c[1] + 1.57, oy + 3.4, ox + c[1] + 1.85, oy + 3.4, TEAL, { w: 1.2 }); });
  },
  "give events that must stay in order the same key (here the customer ID). In Kafka, order is kept inside one partition, never across partitions.");

// 2 duplicate delivery
issueSlide(2, "Duplicate delivery", "A consumer can crash after doing the work but before saving its position.",
  (s, ox, oy) => flowSteps(s, ox, oy, [
    ["Read message M7", "neutral"], ["Add 50 points to the customer", "neutral"], ["Crash before saving the position", "crash"],
    ["Restart: Kafka sends M7 again", "neutral"], ["Add 50 points again: the customer now has 100", "bad"],
  ]),
  (s, ox, oy) => flowSteps(s, ox, oy, [
    ["Read message M7", "neutral"], ["Already handled M7? No", "neutral"], ["Add 50 points and record M7 (one transaction)", "ok"],
    ["Crash, restart: Kafka sends M7 again", "crash"], ["Already handled M7? Yes, so skip it", "ok"],
  ]),
  "make the consumer safe to repeat. Keep a table of handled event IDs and check it in the same transaction as the update. Saving the offset after the work means duplicates are possible, but messages are never lost.");

// duplicate delivery: the fix in detail (event ID table, Redis variant)
add((n) => {
  const s = newSlide(false);
  header(s, "Part 2 · Fixing duplicates", "Skip repeats using the event ID", false, { sub: "Give every event an ID when it is created. The ID is the same on every delivery, so the consumer can tell a repeat." });
  // option 1: database
  label(s, "OPTION 1 · THE DATABASE (SAFEST)", 0.5, 1.72, 7.3, 0.26, { size: 10, bold: true, color: TEAL, align: "left" });
  codeBlock(s, 0.5, 2.0, 7.3, 1.65, "handle M7 (+50 points)", TEAL, [
    "BEGIN;",
    "INSERT INTO processed_events (event_id) VALUES ('M7');",
    "UPDATE loyalty SET points = points + 50 WHERE customer_id = 8842;",
    "COMMIT;",
  ], 10.5);
  label(s, "processed_events · after delivery 1", 0.5, 3.8, 3.4, 0.26, { size: 9, bold: true, align: "left" });
  tbl(s, 0.5, 4.08, [1.3, 1.8], [["event_id (key)", "handled_at"], ["M6", "12:00:01"], ["M7", "12:00:03"]], { fs: 11 });
  label(s, "loyalty · after delivery 1", 4.0, 3.8, 3.8, 0.26, { size: 9, bold: true, align: "left" });
  tbl(s, 4.0, 4.08, [1.6, 1.2], [["customer_id", "points"], ["8842", "50"]], { fs: 11 });
  box(s, 0.5, 5.4, 7.3, 1.45, RUST_TINT, RUST, 1.5);
  s.addText([
    { text: "Delivery 2 of M7 (after the crash)", options: { fontFace: CODE, bold: true, fontSize: 11, color: RUST, breakLine: true, paraSpaceAfter: 6 } },
    { text: "The INSERT fails: M7 is already in the table, and event_id is the key. The whole transaction rolls back, so the points stay at 50. The consumer skips M7 and moves on.", options: { fontFace: BODY, fontSize: 13, color: INK } },
  ], { x: 0.7, y: 5.4, w: 6.9, h: 1.45, isTextBox: true, margin: 0, valign: "middle" });
  // option 2: redis
  label(s, "OPTION 2 · REDIS (FAST)", 8.1, 1.72, 4.75, 0.26, { size: 10, bold: true, color: PLUM, align: "left" });
  codeBlock(s, 8.1, 2.0, 4.75, 1.65, "check before the work", PLUM, [
    "SET processed:M7 1 NX EX 86400",
    "",
    "OK      first time: do the work",
    "(nil)   seen before: skip it",
  ], 10.5);
  label(s, "Redis keys", 8.1, 3.8, 4.75, 0.26, { size: 10, bold: true, align: "left" });
  tbl(s, 8.1, 4.08, [2.0, 0.8, 1.95], [["key", "value", "deleted after"], ["processed:M6", "1", "23 h 59 min"], ["processed:M7", "1", "23 h 58 min"]], { fs: 11 });
  box(s, 8.1, 5.4, 4.75, 1.45, AMBER_TINT, AMBER, 1.5);
  s.addText([
    { text: "Watch out", options: { fontFace: CODE, bold: true, fontSize: 11, color: AMBER, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Redis and the database are two systems. A crash after the SET, before the update, means M7 is marked but never done. Use Redis for harmless work. For points or money, use the database.", options: { fontFace: BODY, fontSize: 12, color: INK } },
  ], { x: 8.3, y: 5.4, w: 4.35, h: 1.45, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, n, false);
});

// 3 poison message
issueSlide(3, "Poison message", "One message always fails. What happens to everything behind it?",
  (s, ox, oy) => {
    laneBox(s, ox + 0.3, oy + 0.55, 5.45, 0.85, "");
    [["M1", TEAL_TINT, TEAL], ["M2", TEAL_TINT, TEAL], ["BAD", RUST_TINT, RUST], ["M4", AMBER_TINT, AMBER], ["M5", AMBER_TINT, AMBER]].forEach((c, k) => chip(s, ox + 0.45 + k * 1.03, oy + 0.77, c[0], { w: 0.88, h: 0.42, fill: c[1], line: c[2], color: c[2] === RUST ? RUST : INK, bold: true }));
    label(s, "done", ox + 0.45, oy + 1.45, 2.0, 0.28, { size: 9, color: TEAL, align: "left" });
    label(s, "blocked behind it", ox + 3.55, oy + 1.45, 2.2, 0.28, { size: 9, color: AMBER, align: "left" });
    box(s, ox + 2.35, oy + 2.1, 1.5, 0.7, WHITE, INK_SOFT); label(s, "Consumer", ox + 2.35, oy + 2.1, 1.5, 0.7, { size: 10.5, bold: true });
    arrow(s, ox + 3.1, oy + 2.1, ox + 3.1, oy + 1.22, RUST, { w: 1.6 });
    label(s, "retry, retry, retry… forever", ox + 0.4, oy + 3.0, 5.3, 0.4, { size: 11, color: RUST, bold: true });
  },
  (s, ox, oy) => {
    laneBox(s, ox + 0.3, oy + 0.35, 5.45, 0.85, "");
    [["M1", TEAL_TINT, TEAL, false], ["M2", TEAL_TINT, TEAL, false], ["BAD", WHITE, INK_FAINT, true], ["M4", TEAL_TINT, TEAL, false], ["M5", TEAL_TINT, TEAL, false]].forEach((c, k) => chip(s, ox + 0.45 + k * 1.03, oy + 0.57, c[0], { w: 0.88, h: 0.42, fill: c[1], line: c[2], dashed: c[3], color: c[3] ? INK_FAINT : INK, bold: true }));
    label(s, "the consumer carries on", ox + 0.4, oy + 1.25, 3.0, 0.28, { size: 9.5, color: TEAL, align: "left" });
    laneBox(s, ox + 0.3, oy + 2.35, 5.45, 0.85, "");
    label(s, "Dead-letter queue", ox + 0.45, oy + 3.25, 3.0, 0.28, { size: 9.5, color: INK_SOFT, bold: true, align: "left" });
    chip(s, ox + 2.0, oy + 2.57, "BAD", { w: 0.88, h: 0.42, fill: RUST_TINT, line: RUST, color: RUST, bold: true });
    arrow(s, ox + 2.44, oy + 1.0, ox + 2.44, oy + 2.33, RUST, { w: 1.4 });
    label(s, "after 3 tries", ox + 2.6, oy + 1.5, 1.8, 0.3, { size: 9.5, color: RUST, bold: true, align: "left" });
    label(s, "someone looks at it later", ox + 3.0, oy + 2.6, 2.7, 0.3, { size: 9, color: INK_SOFT, align: "left" });
  },
  "retry a limited number of times, then move the message to a dead-letter queue and carry on. Kafka has no built-in dead-letter queue: the consumer publishes the failed message to a .dlq topic.");

// 4 backlog
issueSlide(4, "Backlog (consumer lag)", "Messages arrive faster than consumers can process them.",
  (s, ox, oy) => {
    box(s, ox + 0.2, oy + 1.2, 1.15, 0.85, WHITE, INK_SOFT); label(s, "Producer", ox + 0.2, oy + 1.2, 1.15, 0.85, { size: 10.5, bold: true });
    arrow(s, ox + 1.35, oy + 1.62, ox + 1.95, oy + 1.62, RUST, { w: 3 }); label(s, "1,200 / s", ox + 0.1, oy + 2.15, 1.5, 0.3, { size: 9.5, color: RUST, bold: true });
    box(s, ox + 2.0, oy + 0.35, 1.7, 2.75, WHITE, INK_SOFT, 1.5); label(s, "Queue", ox + 2.0, oy + 0.38, 1.7, 0.3, { size: 10, bold: true });
    for (let k = 0; k < 6; k++) s.addShape(pptx.ShapeType.rect, { x: ox + 2.15, y: oy + 0.75 + k * 0.38, w: 1.4, h: 0.3, fill: { color: k < 2 ? AMBER_TINT : RUST_TINT }, line: { color: k < 2 ? AMBER : RUST, width: 0.75 } });
    arrow(s, ox + 3.7, oy + 1.62, ox + 4.3, oy + 1.62, INK_SOFT, { w: 1 }); label(s, "300 / s", ox + 3.7, oy + 1.72, 0.65, 0.3, { size: 9, bold: true });
    box(s, ox + 4.35, oy + 1.2, 1.4, 0.85, WHITE, INK_SOFT); label(s, "1 consumer", ox + 4.35, oy + 1.2, 1.4, 0.85, { size: 10.5, bold: true });
    label(s, "Backlog grows by 900 every second", ox + 0.3, oy + 3.3, 5.5, 0.35, { size: 12, color: RUST, bold: true });
  },
  (s, ox, oy) => {
    box(s, ox + 0.2, oy + 1.2, 1.15, 0.85, WHITE, INK_SOFT); label(s, "Producer", ox + 0.2, oy + 1.2, 1.15, 0.85, { size: 10.5, bold: true });
    arrow(s, ox + 1.35, oy + 1.62, ox + 1.95, oy + 1.62, INK_SOFT, { w: 2 }); label(s, "1,200 / s", ox + 0.1, oy + 2.15, 1.5, 0.3, { size: 9.5, bold: true });
    box(s, ox + 2.0, oy + 0.35, 1.5, 2.75, WHITE, INK_SOFT, 1.5); label(s, "4 partitions", ox + 2.0, oy + 0.38, 1.5, 0.3, { size: 10, bold: true });
    for (let k = 0; k < 4; k++) { const y = oy + 0.8 + k * 0.57; s.addShape(pptx.ShapeType.rect, { x: ox + 2.12, y, w: 1.26, h: 0.4, fill: { color: TEAL_TINT }, line: { color: TEAL, width: 0.75 } }); arrow(s, ox + 3.5, y + 0.2, ox + 4.15, y + 0.2, INK_SOFT, { w: 1 }); box(s, ox + 4.2, y, 1.6, 0.4, PLUM_TINT, PLUM, 1.1); label(s, "consumer " + (k + 1), ox + 4.2, y, 1.6, 0.4, { size: 9.5 }); }
    label(s, "4 x 300 = 1,200 / s", ox + 3.0, oy + 3.1, 2.8, 0.3, { size: 10, color: TEAL, bold: true });
    label(s, "Alert when lag grows. Never more consumers than partitions.", ox + 0.3, oy + 3.4, 5.5, 0.3, { size: 9.5, color: INK_SOFT });
  },
  "add consumers (but never more than the number of partitions), alert when lag grows, and slow down or drop load at the producer if the overload never ends. Lag = newest offset minus saved offset.");

// 5 hot partition
issueSlide(5, "Hot partition", "One huge customer sends most of the traffic.",
  (s, ox, oy) => {
    label(s, "key = merchant_id    (merchant 7 sends 68%)", ox + 0.25, oy + 0.2, 5.5, 0.3, { size: 10, color: RUST, bold: true, align: "left" });
    [["P0", 2.9, RUST_TINT, RUST, "consumer 1: overloaded", RUST], ["P1", 0.68, TEAL_TINT, TEAL, "consumer 2: idle", INK_FAINT], ["P2", 0.68, TEAL_TINT, TEAL, "consumer 3: idle", INK_FAINT]].forEach((r, k) => {
      const y = oy + 0.85 + k * 0.95;
      label(s, r[0], ox + 0.25, y, 0.5, 0.55, { size: 11, bold: true, align: "left" });
      s.addShape(pptx.ShapeType.rect, { x: ox + 0.8, y, w: r[1], h: 0.55, fill: { color: r[2] }, line: { color: r[3], width: 1.2 } });
      label(s, r[4], ox + 0.8 + r[1] + 0.12, y, 5.85 - (0.8 + r[1]) - 0.1, 0.55, { size: 9, color: r[5], bold: r[5] === RUST, align: "left" });
    });
    label(s, "Adding consumers will not help: one partition, one consumer.", ox + 0.25, oy + 3.4, 5.6, 0.3, { size: 10, color: RUST, bold: true, align: "left" });
  },
  (s, ox, oy) => {
    label(s, "key = merchant_id + random 0 to 2 (big merchant only)", ox + 0.25, oy + 0.2, 5.6, 0.3, { size: 10, color: TEAL, bold: true, align: "left" });
    ["P0", "P1", "P2"].forEach((nme, k) => {
      const y = oy + 0.85 + k * 0.95;
      label(s, nme, ox + 0.25, y, 0.5, 0.55, { size: 11, bold: true, align: "left" });
      s.addShape(pptx.ShapeType.rect, { x: ox + 0.8, y, w: 1.4, h: 0.55, fill: { color: TEAL_TINT }, line: { color: TEAL, width: 1.2 } });
      label(s, "consumer " + (k + 1) + ": busy", ox + 2.35, y, 3.4, 0.55, { size: 9, color: TEAL, bold: true, align: "left" });
    });
    label(s, "Load is spread evenly. The cost: order across that merchant's events is lost, so check that nothing depends on it.", ox + 0.25, oy + 3.3, 5.6, 0.5, { size: 9.5, color: INK_SOFT, align: "left", font: BODY });
  },
  "choose a better key, or add a random part to the key for the big customer. The cost is that order across that customer's events is no longer guaranteed. It is the Part 1 hot row again, in a queue.");

// 6 rebalance pauses
function timelineRows(s, ox, oy, rows, joinAt) {
  const X0 = ox + 0.95, W = 4.85;
  label(s, "time →", ox + 4.9, oy + 3.45, 0.9, 0.25, { size: 9, color: INK_FAINT, align: "right" });
  rows.forEach((r, k) => {
    const y = oy + 0.7 + k * 0.85;
    label(s, r.name, ox + 0.1, y, 0.85, 0.5, { size: 8.5, bold: true, align: "left" });
    r.blocks.forEach((b) => {
      s.addShape(pptx.ShapeType.rect, { x: X0 + b[0] * W, y, w: (b[1] - b[0]) * W, h: 0.5, fill: b[2] === "idle" ? { type: "none" } : { color: b[2] === "work" ? TEAL_TINT : b[2] === "pause" ? RUST_TINT : AMBER_TINT }, line: { color: b[2] === "work" ? TEAL : b[2] === "pause" ? RUST : b[2] === "short" ? AMBER : INK_FAINT, width: 1, dashType: b[2] === "idle" ? "dash" : "solid" } });
      label(s, b[3] || "", X0 + b[0] * W, y, (b[1] - b[0]) * W, 0.5, { size: 8.5, color: b[2] === "pause" ? RUST : b[2] === "short" ? AMBER : INK, bold: b[2] !== "work" });
    });
  });
  s.addShape(pptx.ShapeType.line, { x: X0 + joinAt * W, y: oy + 0.45, w: 0, h: 2.7, line: { color: PLUM, width: 1.2, dashType: "dash" } });
  label(s, "consumer 3 joins", X0 + joinAt * W - 0.9, oy + 0.15, 1.8, 0.28, { size: 9.5, color: PLUM, bold: true });
}
issueSlide(6, "Rebalance pauses", "A consumer joins or leaves, and the group hands out the partitions again.",
  (s, ox, oy) => {
    timelineRows(s, ox, oy, [
      { name: "consumer 1", blocks: [[0, 0.4, "work", "working"], [0.4, 0.65, "pause", "stop"], [0.65, 1, "work", "working"]] },
      { name: "consumer 2", blocks: [[0, 0.4, "work", "working"], [0.4, 0.65, "pause", "stop"], [0.65, 1, "work", "working"]] },
      { name: "consumer 3", blocks: [[0, 0.4, "idle", ""], [0.4, 0.65, "pause", "stop"], [0.65, 1, "work", "working"]] },
    ], 0.4);
    label(s, "Everyone stops while partitions are reassigned.", ox + 0.25, oy + 3.4, 4.6, 0.3, { size: 10, color: RUST, bold: true, align: "left" });
  },
  (s, ox, oy) => {
    timelineRows(s, ox, oy, [
      { name: "consumer 1", blocks: [[0, 1, "work", "keeps working"]] },
      { name: "consumer 2", blocks: [[0, 0.4, "work", "working"], [0.4, 0.52, "short", "short"], [0.52, 1, "work", "working"]] },
      { name: "consumer 3", blocks: [[0, 0.4, "idle", ""], [0.4, 0.52, "short", "short"], [0.52, 1, "work", "working"]] },
    ], 0.4);
    label(s, "Only the partitions that move pause.", ox + 0.25, oy + 3.4, 4.6, 0.3, { size: 10, color: TEAL, bold: true, align: "left" });
  },
  "use cooperative rebalancing, so only the partitions that move pause, and static membership, so a quick restart does not trigger a rebalance. Keep the work per message short so pauses stay short.");

// 7 cannot add partitions freely
issueSlide(7, "Can't add partitions freely", "Which partition a key goes to depends on how many partitions there are.",
  (s, ox, oy) => {
    label(s, "3 partitions: customer 8842 goes to hash % 3 = P2", ox + 0.25, oy + 0.2, 5.6, 0.28, { size: 9.5, bold: true, align: "left" });
    ["P0", "P1", "P2"].forEach((nme, k) => laneBox(s, ox + 0.3 + k * 1.2, oy + 0.55, 1.1, 0.85, nme, { top: true }));
    chip(s, ox + 0.3 + 2 * 1.2 + 0.05, oy + 1.0, "old events", { w: 1.0, h: 0.28, size: 8, fill: TEAL_TINT });
    label(s, "Add a 4th partition: customer 8842 now goes to hash % 4 = P0", ox + 0.25, oy + 1.7, 5.7, 0.28, { size: 9.5, bold: true, align: "left" });
    ["P0", "P1", "P2", "P3"].forEach((nme, k) => laneBox(s, ox + 0.3 + k * 1.2, oy + 2.05, 1.1, 0.85, nme, { top: true }));
    chip(s, ox + 0.35, oy + 2.5, "new events", { w: 1.0, h: 0.28, size: 8, fill: AMBER_TINT, line: AMBER });
    chip(s, ox + 0.3 + 2 * 1.2 + 0.05, oy + 2.5, "old events", { w: 1.0, h: 0.28, size: 8, fill: TEAL_TINT });
    label(s, "New events can overtake old ones still waiting in P2. Order is broken for this customer.", ox + 0.25, oy + 3.1, 5.6, 0.6, { size: 10.5, color: RUST, bold: true, align: "left", font: BODY });
  },
  (s, ox, oy) => {
    const grp = (rowY, groups, text) => {
      label(s, text, ox + 0.25, rowY - 0.38, 5.7, 0.28, { size: 9.5, bold: true, align: "left" });
      const per = 12 / groups, bw = 0.42;
      for (let k = 0; k < 12; k++) { const gi = Math.floor(k / per); s.addShape(pptx.ShapeType.rect, { x: ox + 0.3 + k * (bw + 0.03), y: rowY, w: bw, h: 0.5, fill: { color: [TEAL_TINT, PLUM_TINT, AMBER_TINT, "E3EAEC", RUST_TINT, "E8F0E0"][gi] }, line: { color: INK_SOFT, width: 0.75 } }); label(s, String(k), ox + 0.3 + k * (bw + 0.03), rowY, bw, 0.5, { size: 8 }); }
      for (let gi = 0; gi < groups; gi++) label(s, "C" + (gi + 1), ox + 0.3 + gi * per * (bw + 0.03), rowY + 0.55, per * (bw + 0.03), 0.25, { size: 8.5, bold: true, color: PLUM });
    };
    grp(oy + 0.75, 3, "Start with 12 partitions and 3 consumers");
    grp(oy + 2.0, 6, "Later: 6 consumers, same 12 partitions");
    label(s, "Keys never change partition, so order is kept.", ox + 0.25, oy + 3.25, 5.6, 0.3, { size: 10.5, color: TEAL, bold: true, align: "left" });
  },
  "create more partitions than you need at the start, and add consumers as you grow. If you must add partitions later, plan the change (for example a new topic and a cutover) because keys will move.");


checkIn(6, 3, "The same message arrives twice", "Sometimes a consumer crashes after doing the work but before saving its position. Then Kafka sends the same message again. This code gives loyalty points.", [
  "What happens when the message arrives twice?",
  "How can you stop the points being added twice?",
  "What kind of operation would not have this problem?",
], {
  label: "loyalty_worker.py — what if it runs twice?", fs: 13, lines: [
    "def handle(event):",
    "    db.execute(",
    '        "UPDATE loyalty SET points = points + %s "',
    '        "WHERE customer_id = %s",',
    '        event["points"], event["customer_id"],',
    "    )",
  ],
}, "1. The customer gets the points twice.\n2. Keep a table of event IDs you have already handled. Check and update in the same transaction:\n\ndef handle(event):\n    with db.transaction():\n        first_time = db.execute(\n            \"INSERT INTO processed_events (event_id) VALUES (%s) ON CONFLICT DO NOTHING\",\n            event[\"event_id\"])   # 1 row if new, 0 if already seen\n        if first_time:\n            db.execute(\"UPDATE loyalty SET points = points + %s WHERE customer_id = %s\", event[\"points\"], event[\"customer_id\"])\n\nThe event ID must come from the sender, so the same message always has the same ID.\n3. An operation that gives the same result when repeated: SET points = 500 instead of points = points + 50.");

// ============================================================ PART 3
divider(3, "1:45–2:40", "Scaling the database", "Servers and workers are copies of code, so they multiply easily. The database holds the one thing that cannot be copied freely: the data. Copy it, let every copy take writes, or split it, and CAP says what each one costs.");

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Step 1", "Primary and replicas (master–slave)", false, { sub: "One primary takes every write and sends changes to replicas. Reads can go to any of them. It cannot scale writes: there is still only one writer." });
  tbl(s, 0.5, 2.05, [2.3, 2.7, 3.5, 3.85], [
    ["Replication mode", "Write confirmed when…", "Strength", "Tradeoff"],
    ["Asynchronous", "the primary has it", "Fast writes. A slow replica never blocks anyone.", "Replicas fall behind. If the primary crashes, the last few unsent writes can be lost."],
    ["Synchronous", "the replica has it too", "No confirmed write is lost on failover.", "Every write waits for the replica, and a down replica can stall writes."],
    ["Semi-synchronous", "at least one replica has it", "A middle path used by many real systems.", "Slower than async, and limited by the slowest chosen replica."],
  ], { fs: 12, rowH: [0.45, 0.95, 0.95, 0.95] });
  callout(s, 0.5, 5.6, 12.35, 1.25, "Failover has its own danger.", "If the old primary is only unreachable, not dead, and a replica takes over anyway, two servers both think they are the writer. This is called split brain. Real systems block the old primary from accepting writes.", false);
  footer(s, n, false);
});

add((n) => {
  const s = newSlide(true);
  header(s, "Part 3 · Step 1", "Problem: a replica can be out of date", true, { sub: "After adding two read replicas, a customer writes in: \"your site changed my address back.\"" });
  codeBlock(s, 0.5, 1.8, 6.1, 3.0, "orders.py — before", AMBER_LIGHT, [
    "def get_order_confirmation(order_id):",
    "    return db.read_replica().get_order(order_id)",
    "        #! every read hits a random replica",
  ], 12.5);
  codeBlock(s, 6.75, 1.8, 6.1, 3.0, "orders.py — after", TEAL_TINT, [
    "def get_order_confirmation(order_id, session):",
    "    if session.wrote_recently(order_id):",
    "        return db.primary().get_order(order_id)",
    "    return db.read_replica().get_order(order_id)",
  ], 12);
  callout(s, 0.5, 5.05, 12.35, 1.8, "What happened:", "the customer's address update went to the primary. A split second later the confirmation page read it back from a replica that had not received it yet. Replication delay, normally a few tens of milliseconds, was enough. The fix: only reads that follow the user's own recent write go to the primary. Sending every read there would remove the reason for having replicas.", true);
  footer(s, n, true);
});

checkIn(7, 2, "Replica or primary?", "A replica can be a few seconds behind the primary. For each read, is it OK to use a replica, or must it use the primary?", [
  "The product catalog page.",
  "The order page, right after the customer pays.",
  "The admin revenue dashboard.",
  "A stock check just before charging a card.",
  "Another user's public profile.",
], null, "1. Replica. Slightly old data hurts nobody.\n2. Primary. The customer just paid and expects to see their order.\n3. Replica. A few seconds old is fine.\n4. Primary. Old stock data can make you sell an item twice. (Even better: reserve the stock with a one-step update.)\n5. Replica. Nobody can tell the data is a few seconds old.\n\nThe test each time: would the reader notice, or be hurt by, old data?");

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Step 2", "Multi-master: every server accepts writes", false, { sub: "Writes get faster, and if one region goes down, writes still work elsewhere. The price comes when two servers change the same row before hearing about each other." });
  const rows = [
    ["T+0 ms", "Cairo server", "Customer 8842 changes their address to \"12 Nile St.\" The write is confirmed locally."],
    ["T+80 ms", "Dublin server", "The same customer, on their phone, changes it to \"4 Quay Rd.\" Dublin has not received Cairo's change, so it confirms this one too."],
    ["T+200 ms", "Copying catches up", "Each server receives the other's write. Both histories are valid, but they disagree, and neither server did anything wrong."],
  ];
  rows.forEach((r, i) => {
    const y = 2.05 + i * 1.15;
    card(s, 0.5, y, 12.35, 1.0, WHITE, { shadow: false, line: LINE });
    s.addText(r[0], { x: 0.7, y, w: 1.4, h: 1.0, isTextBox: true, margin: 0, valign: "middle", fontFace: CODE, bold: true, fontSize: 14, color: RUST });
    s.addText([
      { text: r[1] + "  ", options: { bold: true, color: INK, fontFace: HEAD, fontSize: 14 } },
      { text: r[2], options: { color: INK_SOFT, fontFace: BODY, fontSize: 12.5 } },
    ], { x: 2.2, y, w: 10.5, h: 1.0, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addText("Someone must decide what the row should say, and every option has a cost.", { x: 0.5, y: 5.65, w: 12.3, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 16, color: TEAL });
  footer(s, n, false);
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Step 2", "How to settle write conflicts", false);
  tbl(s, 0.5, 1.55, [2.3, 5.0, 5.05], [
    ["Strategy", "How it works", "Tradeoff"],
    ["Last write wins", "Keep the version with the later timestamp.", "Simple, but it silently throws away one customer's write. Clocks on different machines disagree."],
    ["Merge", "Version vectors detect the conflict; some types (sets, counters) merge automatically (CRDTs).", "Works when a merge makes sense; most business fields, like an address, cannot be merged."],
    ["Home region per row", "Each row has one owner server; others forward writes to it.", "No conflicts, but remote writes are slower and fail if the owner is unreachable."],
    ["Avoid by design", "Separate ID ranges or UUIDs, so two servers never create the same key.", "Only prevents duplicate keys, not two edits to one row."],
  ], { fs: 12, rowH: [0.45, 0.85, 1.0, 0.85, 0.85] });
  callout(s, 0.5, 5.75, 12.35, 1.1, "Some values cannot be merged by a rule.", "stock = stock - 1 on two masters: both read \"1 left\", both subtract, both confirm. The last item is sold twice, and no timestamp rule fixes it. Contested values need a single owner or a type built to merge.", false);
  footer(s, n, false);
});

checkIn(8, 3, "Two servers edit the same customer", "Two servers accept writes. At the same moment, both change this customer. Start: address = \"1 Old Rd\", marketing_opt_in = true, points = 100.", [
  "Cairo: sets address to \"12 Nile St\", turns marketing_opt_in off, adds 50 points.",
  "Dublin: sets address to \"4 Quay Rd\", adds 30 points.",
  "If the system just keeps the newest whole row, what is lost?",
  "What would you do for each field?",
], null, "Keeping the newest whole row throws away the other server's changes. If Dublin wins, the customer's opt-out is lost (a legal problem) and 50 points disappear.\n- points: keep a list of changes and add them up: 100 + 50 + 30 = 180.\n- marketing_opt_in: use a clear rule, such as \"an opt-out always wins\".\n- address: there are two real values. Ask the customer, or send all address changes to the customer's home region so they cannot conflict.");

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Step 3", "Partitioning (sharding): split the data", false, { sub: "Replicas copy all the data; multi-master copies the conflicts too. To scale writes and storage, give each server part of the data, and give each row exactly one home." });
  s.addText("The obvious way to pick a shard: shard = hash(key) % N. N is an input, so changing N changes almost every key's shard:", { x: 0.5, y: 2.0, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK_SOFT, valign: "top" });
  tbl(s, 0.5, 2.75, [2.4, 3.0, 3.0, 2.2], [
    ["hash(key)", "shard @ N=4", "shard @ N=5", "moved?"],
    ["10", "2", "0", "yes"], ["11", "3", "1", "yes"], ["12", "0", "2", "yes"], ["13", "1", "3", "yes"],
    ["14", "2", "4", "yes"], ["15", "3", "0", "yes"], ["16", "0", "1", "yes"], ["17", "1", "2", "yes"],
  ], { fs: 11, rowH: [0.38, 0.36, 0.36, 0.36, 0.36, 0.36, 0.36, 0.36, 0.36] });
  card(s, 11.3, 2.75, 1.55, 3.3, TEAL_TINT, { shadow: false });
  s.addText("Every key moved.", { x: 11.4, y: 2.9, w: 1.35, h: 3.0, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: HEAD, bold: true, fontSize: 15, color: TEAL });
  s.addText("At real scale, adding capacity means copying nearly all the data before the system works properly again. That is the opposite of what \"add a shard\" should give you.", { x: 0.5, y: 6.0, w: 12.3, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
  s.addNotes("Partitioning means splitting data into pieces; sharding is when the pieces live on separate database servers. Splitting by columns or by feature is vertical partitioning; here we mean splitting rows.");
});

// consistent hashing ring
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Step 3", "Consistent hashing: adding a shard moves few keys", false);
  const cx = 4.3, cy = 4.2, r = 2.1;
  s.addShape(pptx.ShapeType.ellipse, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { type: "none" }, line: { color: INK_SOFT, width: 1.5 } });
  const pos = (hour) => { const a = hour * Math.PI / 6; return [cx + r * Math.sin(a), cy - r * Math.cos(a)]; };
  const shard = (hour, name, color, big) => {
    const [x, y] = pos(hour); const d = big ? 0.46 : 0.38;
    s.addShape(pptx.ShapeType.ellipse, { x: x - d / 2, y: y - d / 2, w: d, h: d, fill: { color }, line: { type: "none" } });
    const out = 0.45, a = hour * Math.PI / 6, lx = cx + (r + out + 0.35) * Math.sin(a), ly = cy - (r + out + 0.2) * Math.cos(a);
    label(s, name, lx - 0.9, ly - 0.17, 1.8, 0.34, { size: 11, bold: !!big, color: big ? AMBER : INK });
  };
  shard(2, "shard A", TEAL); shard(5, "shard B", TEAL); shard(8, "shard C", TEAL); shard(11, "shard D", TEAL); shard(9, "shard E (new)", AMBER, true);
  const key = (hour, name, color) => {
    const [x, y] = pos(hour);
    s.addShape(pptx.ShapeType.ellipse, { x: x - 0.1, y: y - 0.1, w: 0.2, h: 0.2, fill: { color }, line: { type: "none" } });
    const a = hour * Math.PI / 6, lx = cx + (r - 0.5) * Math.sin(a), ly = cy - (r - 0.5) * Math.cos(a);
    label(s, name, lx - 0.7, ly - 0.15, 1.4, 0.3, { size: 10, color: color === AMBER ? AMBER : INK_FAINT, bold: color === AMBER });
  };
  key(3, "K1", TEAL); key(6.5, "K2", TEAL); key(0.5, "K3", TEAL); key(8.5, "K4 → moved", AMBER);
  const pts = [
    ["The rule", "Hash the shard names onto a ring. Hash a key onto the same ring. Walk clockwise to the first shard. No modulo, and no shard count anywhere."],
    ["Adding shard E", "E lands between C and D. K1, K2 and K3 do not move; only K4, which used to walk all the way to D, now stops at E."],
    ["The payoff", "Adding a shard moves only about 1/N of the keys. Plain hash % N would have reshuffled almost all of them."],
  ];
  pts.forEach((p, i) => {
    const y = 1.75 + i * 1.65;
    card(s, 8.0, y, 4.85, 1.5, WHITE, { shadow: false, line: LINE });
    s.addText([
      { text: p[0], options: { bold: true, color: TEAL, fontFace: HEAD, fontSize: 14, breakLine: true } },
      { text: p[1], options: { color: INK_SOFT, fontFace: BODY, fontSize: 11.5 } },
    ], { x: 8.2, y, w: 4.5, h: 1.5, isTextBox: true, margin: 0, valign: "middle" });
  });
  footer(s, n, false);
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Step 3", "Four ways to decide which shard a row lives on", false, { sub: "None is strictly better. It depends on whether queries need ranges, how uneven the keys are, and how much complexity your team can handle." });
  tbl(s, 0.5, 1.95, [2.0, 2.9, 3.4, 4.05], [
    ["Strategy", "Where customer 8842 lives", "Strength", "Tradeoff"],
    ["Range-based", "IDs 8000–8999 → Shard C", "Range scans stay on one shard.", "Increasing IDs: the newest shard gets all the new writes."],
    ["Hash (mod N)", "hash(8842) % 4 → Shard B", "Even load, no hotspot.", "Adding a 5th shard reshuffles about 80% of the keys."],
    ["Consistent hashing", "Placed on the ring, owned by the nearest shard clockwise", "Even load; adding a shard moves ~1/N of the keys.", "No range scans; needs virtual nodes per shard, or load stays uneven."],
    ["Directory-based", "Looked up in a shard-map service → Shard D", "Total flexibility: you can move any single key.", "The directory becomes a new bottleneck and single point of failure."],
  ], { fs: 12, rowH: [0.45, 0.95, 0.95, 1.1, 0.95] });
  footer(s, n, false);
});

checkIn(9, 3, "Pick the shard key", "You split a big table across several servers (shards). Pick the column that decides which shard a row goes to, and say what could go wrong.", [
  "messages in a chat app. Most common query: the latest 50 messages in one conversation.",
  "orders in a shop. Most common queries: \"my orders\" (by customer) and \"orders to ship\" (by seller).",
  "sensor readings from many devices. Most common query: one device, last hour.",
], null, "1. conversation_id. All messages in a conversation sit on one shard, so the query asks only one server. User ID or message ID would spread a conversation over all shards.\n2. customer_id, for \"my orders\". The seller query must now ask every shard. A common fix: keep a second copy of the data grouped by seller.\n3. device_id, then split by time inside each shard. Do not use time as the shard key: the newest data would all go to one shard.");

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Putting it together", "Using partitions and replicas together", false, { sub: "These steps work together; they do not compete." });
  [
    { c: TEAL, t: "What you build", p: ["Partition the data so writes and storage spread across shards.", "Make each shard a primary with replicas, so reads scale and failover works.", "Send recent writers to the primary so they read their own writes."] },
    { c: RUST, t: "What partitioning costs", p: ["Queries and transactions that cross shards get slower and harder.", "An uneven key can still make one shard hot.", "Changing the shard layout later is a real project, not a config change.", "Analytics and cross-shard reports usually move to a separate data warehouse."] },
  ].forEach((c, i) => {
    const x = 0.5 + i * 6.3;
    card(s, x, 1.95, 6.05, 4.7, WHITE);
    s.addShape(pptx.ShapeType.rect, { x, y: 1.95, w: 6.05, h: 0.09, fill: { color: c.c }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.3, y: 2.2, w: 5.5, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 18, color: INK });
    s.addText(bl(c.p, { fs: 14, color: INK_SOFT, gap: 12 }), { x: x + 0.3, y: 2.9, w: 5.5, h: 3.6, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, false);
});

// ============================================================ DISTRIBUTED SYSTEMS AND CAP THEOREM (new section)
function infoCard(s, x, y, w, h, color, title, items, fs) {
  card(s, x, y, w, h, WHITE);
  s.addShape(pptx.ShapeType.rect, { x, y, w: 0.09, h, fill: { color }, line: { type: "none" } });
  const runs = [{ text: title, options: { bold: true, color, fontFace: HEAD, fontSize: 14, breakLine: true, paraSpaceAfter: 5 } }];
  const body = typeof items === "string" ? [{ text: items, options: { color: INK_SOFT, fontFace: BODY, fontSize: fs || 12 } }] : bl(items, { fs: fs || 12, color: INK_SOFT, gap: 3 });
  s.addText(runs.concat(body), { x: x + 0.28, y: y + 0.12, w: w - 0.45, h: h - 0.2, isTextBox: true, margin: 0, valign: "top" });
}

// section opener
add((n) => {
  const s = newSlide(true);
  s.addText("PART 3 · NEW SECTION", { x: 0.7, y: 0.55, w: 8, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Distributed systems and the CAP theorem", { x: 0.7, y: 0.9, w: 11.9, h: 0.9, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 36, color: WHITE });
  [
    ["What a distributed system is", "Many computers, one system, talking over a network"],
    ["Network partitions and split brain", "When the network splits, both sides can think they are in charge"],
    ["The CAP theorem", "Consistency, availability, partition tolerance"],
    ["Why you can only have two, and how to choose", "One lost message, two honest answers. AP or CP?"],
    ["Eventual consistency", "Copies that disagree for a while, then agree"],
  ].forEach((it, i) => {
    const y = 2.2 + i * 0.92;
    card(s, 0.7, y, 11.9, 0.74, DARK2, { shadow: false });
    circle(s, 0.95, y + 0.12, 0.5, TEAL, String(i + 1), { fontSize: 13 });
    s.addText(it[0], { x: 1.7, y, w: 5.4, h: 0.74, isTextBox: true, margin: 0, valign: "middle", fontFace: HEAD, bold: true, fontSize: 16, color: WHITE });
    s.addText(it[1], { x: 7.2, y, w: 5.3, h: 0.74, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13, color: "C6DADC" });
  });
  footer(s, n, true);
});

// what is a distributed system
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "What is a distributed system?", false, { sub: "Many computers that work together by sending messages to each other, and look like one system to the user." });
  card(s, 0.5, 1.85, 7.4, 5.0, WHITE, { shadow: false, line: LINE });
  box(s, 0.7, 3.75, 1.3, 1.0, WHITE, INK_SOFT); label(s, "Customer", 0.7, 3.75, 1.3, 1.0, { size: 10.5, bold: true });
  s.addShape(pptx.ShapeType.roundRect, { x: 2.35, y: 2.25, w: 5.35, h: 4.4, rectRadius: 0.1, fill: { type: "none" }, line: { color: TEAL, width: 1.5, dashType: "dash" } });
  label(s, "ONE SYSTEM TO THE CUSTOMER", 2.35, 1.95, 5.35, 0.25, { size: 9, bold: true, color: TEAL });
  box(s, 2.55, 3.8, 1.4, 0.9, AMBER_TINT, AMBER, 1.5); label(s, "Load balancer", 2.55, 3.8, 1.4, 0.9, { size: 10, bold: true });
  [2.5, 3.8, 5.1].forEach((y, i) => { box(s, 4.35, y, 1.4, 0.9, TEAL_TINT, TEAL, 1.25); label(s, "web-" + (i + 1), 4.35, y, 1.4, 0.9, { size: 10, bold: true }); arrow(s, 3.95, 4.25, 4.3, y + 0.45, INK_SOFT, { w: 1 }); });
  [[2.4, "Redis cache", AMBER_TINT, AMBER], [3.45, "Postgres primary", "E3EAEC", SLATE], [4.5, "Postgres replica", "E3EAEC", SLATE], [5.55, "Kafka", PLUM_TINT, PLUM]].forEach(([y, t, f, c]) => { box(s, 6.1, y, 1.45, 0.8, f, c, 1.25); label(s, t, 6.1, y, 1.45, 0.8, { size: 9, bold: true }); });
  [[2.95, 2.8], [4.1, 3.85], [4.4, 4.9], [5.55, 5.95]].forEach(([ya, yb]) => arrow(s, 5.75, ya, 6.05, yb, INK_SOFT, { w: 1 }));
  arrow(s, 2.0, 4.25, 2.5, 4.25, INK_SOFT, { w: 1.2 });
  label(s, "every arrow is a message over the network", 0.6, 5.0, 1.65, 0.8, { size: 8.5, color: INK_FAINT });
  infoCard(s, 8.15, 1.85, 4.7, 2.35, TEAL, "Why we build them", ["Scale: one machine is not enough.", "Stay up: if one machine dies, the others carry on.", "Be close to users: data in Europe and in Asia."], 12);
  infoCard(s, 8.15, 4.4, 4.7, 2.45, RUST, "What it costs", ["Messages get lost or arrive late.", "Machines fail, each on its own.", "No shared clock, and no node sees the whole picture."], 12);
  footer(s, n, false);
});

// a call can fail in three ways
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "A call can fail in three ways, and you can't tell which", false, { sub: "Service A calls Service B and waits. Nothing comes back." });
  [["1 · The request was lost", 0], ["2 · B crashed before it replied", 1], ["3 · B did the work, but the reply was lost", 2]].forEach(([t, k]) => {
    const y = 1.85 + k * 1.7;
    card(s, 0.5, y, 8.2, 1.55, WHITE, { shadow: false, line: LINE });
    label(s, t, 0.7, y + 0.08, 7.8, 0.3, { size: 10.5, bold: true, align: "left" });
    box(s, 0.8, y + 0.5, 1.3, 0.8, TEAL_TINT, TEAL, 1.25); label(s, "A", 0.8, y + 0.5, 1.3, 0.8, { size: 12, bold: true });
    box(s, 6.9, y + 0.5, 1.5, 0.8, k === 1 ? RUST_TINT : PLUM_TINT, k === 1 ? RUST : PLUM, 1.25); label(s, k === 1 ? "B (crashed)" : "B", 6.9, y + 0.5, 1.5, 0.8, { size: 11, bold: true, color: k === 1 ? RUST : INK });
    if (k === 0) {
      arrow(s, 2.1, y + 0.9, 4.6, y + 0.9, RUST, { w: 1.6 }); label(s, "✗", 4.65, y + 0.7, 0.4, 0.4, { size: 16, bold: true, color: RUST });
      label(s, "request", 2.2, y + 0.6, 2.2, 0.25, { size: 9, color: INK_SOFT }); label(s, "B never saw it", 5.1, y + 0.75, 1.7, 0.3, { size: 9, color: RUST });
    } else if (k === 1) {
      arrow(s, 2.1, y + 0.9, 6.85, y + 0.9, INK_SOFT, { w: 1.6 }); label(s, "request", 2.2, y + 0.6, 4.4, 0.25, { size: 9, color: INK_SOFT });
    } else {
      arrow(s, 2.1, y + 0.75, 6.85, y + 0.75, INK_SOFT, { w: 1.6 }); label(s, "request: B does the work", 2.2, y + 0.47, 4.4, 0.25, { size: 9, color: INK_SOFT });
      arrow(s, 6.85, y + 1.1, 4.3, y + 1.1, RUST, { w: 1.6, dash: "dash" }); label(s, "✗", 3.85, y + 0.9, 0.4, 0.4, { size: 16, bold: true, color: RUST });
      label(s, "reply", 4.5, y + 1.12, 2.2, 0.25, { size: 9, color: RUST });
    }
  });
  infoCard(s, 8.95, 1.85, 3.9, 1.5, RUST, "What A sees", "In all three cases: silence, then a timeout. It looks exactly the same.", 12);
  infoCard(s, 8.95, 3.5, 3.9, 2.35, AMBER, "So what can A do?", ["Retry: the work may happen twice (Part 2: make it safe to repeat).", "Give up: the work may already be done.", "Ask again: that answer can be lost too."], 11.5);
  infoCard(s, 8.95, 6.0, 3.9, 0.85, TEAL, "Everything next grows from this.", "", 11);
  footer(s, n, false);
});

// network partition
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Network partition: every node is up, but they can't talk", false, { sub: "The network splits the nodes into groups. Each group is still running, but it cannot reach the other." });
  card(s, 0.5, 1.85, 7.4, 5.0, WHITE, { shadow: false, line: LINE });
  [[0.85, "Group A", TEAL, TEAL_TINT, ["N1", "N2", "N3"]], [5.05, "Group B", PLUM, PLUM_TINT, ["N4", "N5", "N6"]]].forEach(([x, name, c, f, nodes]) => {
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.35, w: 2.0, h: 3.7, rectRadius: 0.1, fill: { type: "none" }, line: { color: c, width: 1.5, dashType: "dash" } });
    label(s, name, x, 2.0, 2.0, 0.3, { size: 10.5, bold: true, color: c });
    nodes.forEach((nm, i) => { box(s, x + 0.25, 2.65 + i * 1.1, 1.5, 0.8, f, c, 1.5); label(s, nm, x + 0.25, 2.65 + i * 1.1, 1.5, 0.8, { size: 12, bold: true }); });
    for (let i = 0; i < 2; i++) arrow(s, x + 1.0, 3.45 + i * 1.1, x + 1.0, 3.75 + i * 1.1, INK_SOFT, { w: 1.2, noHead: true });
  });
  [3.05, 4.15, 5.25].forEach((y) => { arrow(s, 2.65, y, 5.25, y, RUST, { w: 1.5, dash: "dash", noHead: true }); label(s, "✗", 3.7, y - 0.22, 0.5, 0.44, { size: 18, bold: true, color: RUST }); });
  label(s, "the network is cut here", 2.9, 5.7, 2.2, 0.3, { size: 9.5, bold: true, color: RUST });
  label(s, "Each node still works. Each side just cannot see the other.", 0.7, 6.25, 7.0, 0.4, { size: 10, color: INK_SOFT, font: BODY });
  infoCard(s, 8.15, 1.85, 4.7, 1.35, TEAL, "What it is", "The network splits into groups that cannot talk to each other. No machine has to crash.", 12);
  infoCard(s, 8.15, 3.35, 4.7, 2.1, AMBER, "What causes it", ["A switch or router fails.", "A cable or a cloud-zone link is cut.", "A firewall or config change.", "A busy link drops messages.", "A long pause makes a node look dead."], 11.5);
  infoCard(s, 8.15, 5.6, 4.7, 1.25, RUST, "Why it hurts", "Each group thinks the other group died, and keeps going on its own.", 12);
  footer(s, n, false);
});

// split brain problem
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Split brain: two leaders, one set of data", false, { sub: "A bank account has 100. A leader and a replica keep it in sync. Then the network is cut." });
  const px = [0.5, 4.7, 8.9];
  ["1 · Normal", "2 · The network is cut", "3 · The network heals"].forEach((t, i) => {
    card(s, px[i], 1.85, 3.95, 4.15, WHITE, { shadow: false, line: i === 1 ? RUST : LINE });
    label(s, t, px[i] + 0.15, 1.95, 3.65, 0.3, { size: 11, bold: true, align: "left", color: i === 1 ? RUST : INK });
  });
  // panel 1
  box(s, 0.75, 3.0, 1.45, 0.9, TEAL_TINT, TEAL, 1.5); label(s, "Leader", 0.75, 3.0, 1.45, 0.9, { size: 11, bold: true });
  box(s, 2.75, 3.0, 1.45, 0.9, PLUM_TINT, PLUM, 1.5); label(s, "Replica", 2.75, 3.0, 1.45, 0.9, { size: 11, bold: true });
  arrow(s, 2.2, 3.45, 2.72, 3.45, INK_SOFT, { w: 1.4 }); label(s, "copies", 2.1, 3.05, 0.8, 0.25, { size: 8.5, color: INK_FAINT });
  label(s, "balance 100", 0.75, 3.95, 1.45, 0.3, { size: 9.5 }); label(s, "balance 100", 2.75, 3.95, 1.45, 0.3, { size: 9.5 });
  s.addText("The replica copies every change from the leader. Only the leader accepts writes.", { x: 0.75, y: 4.6, w: 3.45, h: 1.1, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: INK_SOFT, valign: "top" });
  // panel 2
  box(s, 4.95, 2.45, 1.45, 0.45, WHITE, INK_SOFT); label(s, "Client 1", 4.95, 2.45, 1.45, 0.45, { size: 9 });
  box(s, 6.95, 2.45, 1.45, 0.45, WHITE, INK_SOFT); label(s, "Client 2", 6.95, 2.45, 1.45, 0.45, { size: 9 });
  arrow(s, 5.675, 2.9, 5.675, 3.35, INK_SOFT, { w: 1.2 }); arrow(s, 7.675, 2.9, 7.675, 3.35, INK_SOFT, { w: 1.2 });
  label(s, "withdraw 80", 5.8, 2.98, 1.0, 0.3, { size: 8.5, color: RUST, align: "left" }); label(s, "withdraw 80", 7.8, 2.98, 0.85, 0.3, { size: 8.5, color: RUST, align: "left" });
  box(s, 4.95, 3.4, 1.45, 0.9, RUST_TINT, RUST, 2); label(s, "LEADER", 4.95, 3.4, 1.45, 0.9, { size: 11, bold: true, color: RUST });
  box(s, 6.95, 3.4, 1.45, 0.9, RUST_TINT, RUST, 2); label(s, "LEADER", 6.95, 3.4, 1.45, 0.9, { size: 11, bold: true, color: RUST });
  arrow(s, 6.4, 3.85, 6.95, 3.85, RUST, { w: 1.4, dash: "dash", noHead: true }); label(s, "✗", 6.4, 3.55, 0.55, 0.4, { size: 16, bold: true, color: RUST });
  label(s, "balance 20", 4.95, 4.35, 1.45, 0.3, { size: 9.5, bold: true, color: RUST }); label(s, "balance 20", 6.95, 4.35, 1.45, 0.3, { size: 9.5, bold: true, color: RUST });
  label(s, "thinks the other is down", 4.9, 4.65, 1.55, 0.25, { size: 7.5, italic: true, color: RUST }); label(s, "thinks the other is down", 6.9, 4.65, 1.55, 0.25, { size: 7.5, italic: true, color: RUST });
  s.addText("Each side thinks the other is down. The replica hears nothing from the leader, so it promotes itself. The old leader hears nothing from the replica, so it carries on. Two leaders accept writes.", { x: 4.95, y: 4.95, w: 3.45, h: 1.0, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11, color: INK_SOFT, valign: "top" });
  // panel 3
  box(s, 9.15, 3.0, 1.45, 0.9, RUST_TINT, RUST, 1.5); label(s, "balance 20", 9.15, 3.0, 1.45, 0.9, { size: 10.5, bold: true });
  box(s, 11.15, 3.0, 1.45, 0.9, RUST_TINT, RUST, 1.5); label(s, "balance 20", 11.15, 3.0, 1.45, 0.9, { size: 10.5, bold: true });
  label(s, "?", 10.6, 3.15, 0.55, 0.6, { size: 20, bold: true, color: RUST });
  s.addText("160 was paid out of 100. The two copies have the same number but a different history. One withdrawal may have to be undone.", { x: 9.15, y: 4.2, w: 3.45, h: 1.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: INK_SOFT, valign: "top" });
  callout(s, 0.5, 6.2, 12.35, 0.65, "Split brain:", "each side believes the other is down, so both act as the leader, both accept writes, and the data drifts apart.", false);
  footer(s, n, false);
});

// quorum
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Stopping split brain: the majority rules", false, { sub: "Five nodes. The network splits them 3 and 2." });
  card(s, 0.5, 1.85, 7.4, 5.0, WHITE, { shadow: false, line: LINE });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.85, y: 2.3, w: 2.1, h: 3.75, rectRadius: 0.1, fill: { type: "none" }, line: { color: TEAL, width: 1.75 } });
  s.addShape(pptx.ShapeType.roundRect, { x: 5.1, y: 2.3, w: 2.1, h: 3.75, rectRadius: 0.1, fill: { type: "none" }, line: { color: RUST, width: 1.75, dashType: "dash" } });
  [2.55, 3.7, 4.85].forEach((y, i) => { box(s, 1.15, y, 1.5, 0.8, TEAL_TINT, TEAL, 1.5); label(s, "N" + (i + 1), 1.15, y, 1.5, 0.8, { size: 12, bold: true }); });
  [3.1, 4.25].forEach((y, i) => { box(s, 5.4, y, 1.5, 0.8, RUST_TINT, RUST, 1.5); label(s, "N" + (i + 4), 5.4, y, 1.5, 0.8, { size: 12, bold: true }); });
  [3.5, 4.5].forEach((y) => { arrow(s, 2.65, y, 5.35, y, RUST, { w: 1.4, dash: "dash", noHead: true }); label(s, "✗", 3.75, y - 0.22, 0.5, 0.44, { size: 16, bold: true, color: RUST }); });
  label(s, "3 of 5: a majority", 0.75, 6.12, 2.4, 0.3, { size: 10, bold: true, color: TEAL });
  label(s, "elects a leader, keeps working", 0.7, 6.4, 2.6, 0.3, { size: 9, color: INK_SOFT });
  label(s, "2 of 5: a minority", 4.95, 6.12, 2.4, 0.3, { size: 10, bold: true, color: RUST });
  label(s, "cannot elect a leader, refuses writes", 4.7, 6.4, 3.0, 0.3, { size: 9, color: INK_SOFT });
  infoCard(s, 8.15, 1.85, 4.7, 3.35, TEAL, "The rules", ["A leader needs votes from more than half of all nodes (3 of 5).", "Two groups can never both have a majority, so there is only ever one leader.", "Use an odd number of nodes: 3 or 5.", "Fencing: each leader gets a number. Storage rejects writes with an older number."], 11.5);
  infoCard(s, 8.15, 5.4, 4.7, 1.45, RUST, "What it costs", "The minority side stops answering. We gave up availability to keep the data correct. That is the CAP choice.", 12);
  footer(s, n, false);
});

// a follower (non-leader) goes down
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "When a follower goes down: what happens?", false, { sub: "A follower is any node that is not the leader. Five nodes, N1 is the leader, and N4 crashes." });
  card(s, 0.5, 1.85, 7.4, 5.0, WHITE, { shadow: false, line: LINE });
  const rows = [
    ["1 · A follower crashes", { 1: ["LEADER", TEAL_TINT, TEAL, 2.5], 2: ["ok", WHITE, INK_SOFT, 1.25], 3: ["ok", WHITE, INK_SOFT, 1.25], 4: ["down", RUST_TINT, RUST, 2], 5: ["ok", WHITE, INK_SOFT, 1.25] }, "The leader stops getting replies from N4. It keeps going. There is no election, because the leader is fine."],
    ["2 · Writes still work", { 1: ["LEADER", TEAL_TINT, TEAL, 2.5], 2: ["acks", TEAL_TINT, TEAL, 1.5], 3: ["acks", TEAL_TINT, TEAL, 1.5], 4: ["down", RUST_TINT, RUST, 2], 5: ["acks", TEAL_TINT, TEAL, 1.5] }, "A write needs a majority: 3 of 5. The leader plus N2, N3 and N5 is 4, so it is confirmed. The user notices nothing."],
    ["3 · N4 comes back and catches up", { 1: ["LEADER", TEAL_TINT, TEAL, 2.5], 2: ["ok", WHITE, INK_SOFT, 1.25], 3: ["ok", WHITE, INK_SOFT, 1.25], 4: ["catching up", AMBER_TINT, AMBER, 2], 5: ["ok", WHITE, INK_SOFT, 1.25] }, "N4 restarts as a follower. The leader sends it the writes it missed (or a snapshot, if it was gone a long time). Then it is a full member again."],
  ];
  rows.forEach(([title, nodes, cap], r) => {
    const y = 1.95 + r * 1.62;
    label(s, title, 0.7, y, 7.0, 0.28, { size: 11, bold: true, align: "left" });
    [1, 2, 3, 4, 5].forEach((k, j) => {
      const [t, f, c, lw] = nodes[k];
      const bx = 0.75 + j * 1.4;
      box(s, bx, y + 0.35, 1.15, 0.62, f, c, lw);
      label(s, "N" + k, bx, y + 0.36, 1.15, 0.33, { size: 11, bold: true });
      label(s, t, bx, y + 0.67, 1.15, 0.26, { size: 7.5, bold: true, color: c === INK_SOFT ? INK_FAINT : c });
    });
    s.addText(cap, { x: 0.7, y: y + 1.07, w: 7.0, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11, color: INK_SOFT, valign: "top" });
  });
  label(s, "How many nodes can be down?", 8.15, 1.85, 4.7, 0.3, { size: 10.5, bold: true, color: TEAL, align: "left" });
  tbl(s, 8.15, 2.2, [1.4, 1.5, 1.8], [["Cluster", "Majority", "Can lose"], ["3 nodes", "2", "1"], ["5 nodes", "3", "2"], ["7 nodes", "4", "3"]], { fs: 11.5, rowH: [0.38, 0.38, 0.38, 0.38] });
  infoCard(s, 8.15, 4.05, 4.7, 1.45, TEAL, "What the user sees", "Nothing. Reads and writes keep working. No election, no pause.", 12);
  infoCard(s, 8.15, 5.65, 4.7, 1.2, RUST, "The limit", "Lose the majority (3 of 5) and writes stop. That is the CAP choice again.", 12);
  footer(s, n, false);
});

// leader election: the steps
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "When the leader goes down: how a new one is chosen", false, { sub: "Five nodes. A follower waits for heartbeats from the leader. When they stop, it is time for an election." });
  const stages = [
    ["1 · Normal: the leader sends heartbeats", { 1: ["LEADER", TEAL_TINT, TEAL, 2.5], 2: ["follower", WHITE, INK_SOFT, 1.25], 3: ["follower", WHITE, INK_SOFT, 1.25], 4: ["follower", WHITE, INK_SOFT, 1.25], 5: ["follower", WHITE, INK_SOFT, 1.25] }, "A short message every ~50 ms tells the followers: I am alive, I am still the leader."],
    ["2 · The leader crashes", { 1: ["down", RUST_TINT, RUST, 2], 2: ["waiting", AMBER_TINT, AMBER, 1.5], 3: ["waiting", AMBER_TINT, AMBER, 1.5], 4: ["waiting", AMBER_TINT, AMBER, 1.5], 5: ["waiting", AMBER_TINT, AMBER, 1.5] }, "No heartbeat. Each follower waits a different, random time (say 150 to 300 ms) before it decides the leader is gone."],
    ["3 · One timer runs out first", { 1: ["down", RUST_TINT, RUST, 2], 2: ["voter", WHITE, INK_SOFT, 1.25], 3: ["CANDIDATE", PLUM_TINT, PLUM, 2.5], 4: ["voter", WHITE, INK_SOFT, 1.25], 5: ["voter", WHITE, INK_SOFT, 1.25] }, "N3 adds 1 to the term (now term 2), votes for itself, and asks everyone: vote for me? Each node votes once per term."],
    ["4 · A majority says yes", { 1: ["down", RUST_TINT, RUST, 2], 2: ["voted yes", TEAL_TINT, TEAL, 1.5], 3: ["NEW LEADER", TEAL_TINT, TEAL, 2.5], 4: ["voted yes", TEAL_TINT, TEAL, 1.5], 5: ["voted yes", TEAL_TINT, TEAL, 1.5] }, "3 votes out of 5 is a majority (N3 + any two). N3 is the leader of term 2 and starts sending heartbeats."],
  ];
  stages.forEach(([title, nodes, cap], i) => {
    const x = i % 2 === 0 ? 0.5 : 6.8, y = i < 2 ? 1.85 : 4.3;
    card(s, x, y, 6.05, 2.3, WHITE, { shadow: false, line: LINE });
    label(s, title, x + 0.2, y + 0.1, 5.65, 0.3, { size: 11, bold: true, align: "left" });
    [1, 2, 3, 4, 5].forEach((k, j) => {
      const [t, f, c, lw] = nodes[k];
      const bx = x + 0.2 + j * 1.14;
      box(s, bx, y + 0.55, 1.0, 0.75, f, c, lw);
      label(s, "N" + k, bx, y + 0.57, 1.0, 0.35, { size: 11, bold: true });
      label(s, t, bx, y + 0.92, 1.0, 0.3, { size: 7.5, bold: true, color: c === INK_SOFT ? INK_FAINT : c });
    });
    s.addText(cap, { x: x + 0.2, y: y + 1.45, w: 5.65, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: INK_SOFT, valign: "top" });
  });
  footer(s, n, false);
});

// leader election: the rules
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "The election rules that keep it safe", false, { sub: "These are the ideas behind Raft, used by etcd, Consul and Kafka's KRaft. ZooKeeper uses a similar protocol." });
  tbl(s, 0.5, 1.85, [3.5, 4.9, 3.95], [
    ["What could go wrong", "The rule", "Why it works"],
    ["Two nodes ask for votes at the same time", "Wait a random time before asking. If nobody wins, try again with a new term.", "A tie is rare and fixes itself."],
    ["A node votes twice", "One vote per term.", "Two leaders can never both win a majority."],
    ["A node that missed some writes becomes leader", "Vote only for a candidate whose data is at least as new as your own.", "The new leader has every write that was confirmed."],
    ["The old leader was only slow, and comes back", "Every message has a term number. A leader that sees a higher term steps down.", "The old leader cannot keep accepting writes."],
    ["Only 2 of the 5 nodes can reach each other", "No one can get 3 votes, so there is no election result.", "Safe, but unavailable. That is the CAP choice."],
  ], { fs: 11.5, rowH: [0.4, 0.85, 0.7, 0.85, 0.85, 0.85] });
  callout(s, 0.5, 6.2, 12.35, 0.7, "Remember:", "an election needs a majority, and a majority exists on only one side of a split. That is how it prevents two leaders.", false);
  footer(s, n, false);
});

// CAP properties
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "The CAP theorem: three properties", false, { sub: "Eric Brewer proposed it in 2000. Gilbert and Lynch proved it in 2002." });
  [
    [TEAL, TEAL_TINT, "C", "Consistency", "Every read sees the latest write.", "You transfer 50. Your phone and your partner's phone both show the new balance straight away."],
    [PLUM, PLUM_TINT, "A", "Availability", "Every request to a working node gets an answer, not an error.", "The ATM always answers, even if the answer is a little old."],
    [AMBER, AMBER_TINT, "P", "Partition tolerance", "The system keeps working when messages between nodes are lost.", "Cairo and Dublin cannot reach each other, and both keep running."],
  ].forEach(([c, f, ch, name, def, ex], i) => {
    const x = 0.5 + i * 4.15, y = 1.85, w = 4.0, h = 4.1;
    card(s, x, y, w, h, WHITE);
    s.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.09, fill: { color: c }, line: { type: "none" } });
    circle(s, x + 0.3, y + 0.35, 0.8, c, ch, { fontSize: 22 });
    s.addText(name, { x: x + 1.3, y: y + 0.35, w: w - 1.5, h: 0.8, isTextBox: true, margin: 0, valign: "middle", fontFace: HEAD, bold: true, fontSize: 17, color: c });
    s.addText(def, { x: x + 0.3, y: y + 1.35, w: w - 0.6, h: 1.0, isTextBox: true, margin: 0, valign: "top", fontFace: BODY, fontSize: 14, color: INK });
    box(s, x + 0.3, y + 2.45, w - 0.6, 1.4, f, c, 1);
    s.addText([{ text: "EXAMPLE", options: { fontFace: CODE, bold: true, fontSize: 9, color: c, breakLine: true, paraSpaceAfter: 3 } }, { text: ex, options: { fontFace: BODY, fontSize: 12, color: INK } }], { x: x + 0.45, y: y + 2.5, w: w - 0.9, h: 1.3, isTextBox: true, margin: 0, valign: "middle" });
  });
  callout(s, 0.5, 6.15, 12.35, 0.75, "The theorem:", "when a network partition happens, a system cannot be both consistent and available. CAP's \"consistency\" is not the C in ACID.", false);
  footer(s, n, false);
});

// ---- CAP as messages: one lost replication message (sequence diagram)
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Why you can't have all three: one lost message", false, { sub: "Two nodes keep a copy of x. The link between them is cut. A client asks Node 2 for x. Whatever Node 2 does, it gives something up." });
  card(s, 0.5, 1.75, 8.1, 5.2, WHITE, { shadow: false, line: LINE });
  const X = { C: 1.6, N1: 4.5, N2: 7.4 };
  [["Client", X.C, TEAL_TINT, TEAL], ["Node 1", X.N1, "E3EAEC", SLATE], ["Node 2", X.N2, PLUM_TINT, PLUM]].forEach(([t, x, f, c]) => {
    box(s, x - 0.75, 1.88, 1.5, 0.45, f, c, 1.5);
    label(s, t, x - 0.75, 1.88, 1.5, 0.45, { size: 11, bold: true });
    s.addShape(pptx.ShapeType.line, { x, y: 2.33, w: 0, h: 6.85 - 2.33, line: { color: INK_FAINT, width: 1, dashType: "dash" } });
  });
  seqMsg(s, X.C, X.N1, 2.95, "1  write x = 5");
  seqMsg(s, X.N1, X.C, 3.45, "saved", { dash: true });
  arrow(s, X.N1 + 0.06, 4.15, 6.3, 4.15, RUST, { w: 1.6 });
  label(s, "2  replicate x = 5", X.N1 + 0.05, 3.88, 1.9, 0.25, { size: 9.5, color: RUST, bold: true });
  label(s, "✗", 6.35, 3.98, 0.4, 0.35, { size: 16, bold: true, color: RUST });
  label(s, "lost: network cut", X.N1 + 0.05, 4.2, 1.9, 0.25, { size: 8.5, color: RUST });
  seqMsg(s, X.C, X.N2, 4.95, "3  read x");
  seqMsg(s, X.N2, X.C, 5.6, "A  x = 4   (old value, but an answer)", { dash: true, color: AMBER, textColor: AMBER, bold: true });
  seqMsg(s, X.N2, X.C, 6.35, "B  error: try again later", { dash: true, color: RUST, textColor: RUST, bold: true });
  sidePanel(s, [
    [SLATE, "What happened", "Node 1 saved x = 5, but the message to Node 2 was lost. Node 2 still has x = 4 and cannot know a newer value exists."],
    [AMBER, "Choice A: answer (availability)", "Node 2 replies x = 4. The user gets an answer, but it is old."],
    [RUST, "Choice B: refuse (consistency)", "Node 2 returns an error until the link is back. Nothing wrong is ever shown, but this request fails."],
  ]);
  footer(s, n, false);
});

// ---- pick two (triangle + examples)
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Why only two of three: CP, AP or CA", false, { sub: "The network will fail sometimes, so P is not a choice. The real choice is C or A." });
  card(s, 0.5, 1.85, 6.1, 5.0, WHITE, { shadow: false, line: LINE });
  const C = [3.55, 2.65], A = [1.45, 5.0], P = [5.65, 5.0];
  arrow(s, C[0], C[1], P[0], P[1], TEAL, { w: 3, noHead: true });
  arrow(s, A[0], A[1], P[0], P[1], PLUM, { w: 3, noHead: true });
  arrow(s, C[0], C[1], A[0], A[1], INK_FAINT, { w: 1.5, noHead: true, dash: "dash" });
  [[C, "C", TEAL_TINT, TEAL], [A, "A", PLUM_TINT, PLUM], [P, "P", AMBER_TINT, AMBER]].forEach(([pt, t, f, c]) => {
    s.addShape(pptx.ShapeType.ellipse, { x: pt[0] - 0.45, y: pt[1] - 0.45, w: 0.9, h: 0.9, fill: { color: f }, line: { color: c, width: 2 } });
    label(s, t, pt[0] - 0.45, pt[1] - 0.45, 0.9, 0.9, { size: 20, bold: true, color: c });
  });
  label(s, "Consistency", 4.1, 2.4, 2.0, 0.28, { size: 10, color: INK_SOFT, align: "left" });
  label(s, "Availability", 0.7, 5.55, 1.5, 0.28, { size: 10, color: INK_SOFT });
  label(s, "Partition tolerance", 4.5, 5.55, 2.0, 0.28, { size: 10, color: INK_SOFT });
  label(s, "CP", 4.85, 3.7, 0.6, 0.3, { size: 14, bold: true, color: TEAL });
  label(s, "AP", 3.2, 4.65, 0.6, 0.3, { size: 14, bold: true, color: PLUM });
  label(s, "CA", 1.55, 3.6, 0.6, 0.3, { size: 12, bold: true, color: INK_FAINT });
  label(s, "needs a network that never fails", 0.65, 3.95, 1.5, 0.5, { size: 8, color: INK_FAINT });
  label(s, "We always need P, so pick one of the two thick edges.", 0.7, 6.15, 5.7, 0.5, { size: 10.5, bold: true, color: INK_SOFT, font: BODY });
  [
    [TEAL, "CP: refuse when unsure", "You give up availability. Some requests fail during a partition.", "Bank transfer, seat 14C, etcd, ZooKeeper, a database where all reads go to the primary."],
    [PLUM, "AP: answer, fix later", "You give up consistency. Answers may be old until the copies catch up.", "Like count, shopping cart, DNS, Cassandra, DynamoDB (default reads)."],
    [INK_FAINT, "CA: only on one machine", "There is no network to split, so no partition to survive. It is not a distributed system.", "One Postgres server with no replicas. Add a second machine and you must pick CP or AP."],
  ].forEach(([c, t, a, ex], i) => {
    const y = 1.85 + i * 1.7;
    card(s, 6.8, y, 6.05, 1.55, WHITE);
    s.addShape(pptx.ShapeType.rect, { x: 6.8, y, w: 0.09, h: 1.55, fill: { color: c }, line: { type: "none" } });
    s.addText([
      { text: t, options: { bold: true, color: c === INK_FAINT ? INK_SOFT : c, fontFace: HEAD, fontSize: 14, breakLine: true, paraSpaceAfter: 3 } },
      { text: a, options: { color: INK, fontFace: BODY, fontSize: 11.5, breakLine: true, paraSpaceAfter: 3 } },
      { text: "Examples: " + ex, options: { color: INK_SOFT, fontFace: BODY, fontSize: 11 } },
    ], { x: 7.08, y: y + 0.08, w: 5.65, h: 1.4, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, false);
});

// AP or CP: decision flow
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "How to choose: AP or CP?", false, { sub: "Ask what a wrong answer costs, and what a refusal costs. The business already has an opinion." });
  [
    [1.95, "1 · Is an old or wrong answer dangerous?", "money, safety, double booking, a unique name"],
    [3.55, "2 · If two sides disagree, can you fix it later?", "merge the changes, last write wins, or a person reviews"],
    [5.15, "3 · Does refusing cost more than a slightly old answer?", "lost sales, blank pages, angry users"],
  ].forEach(([y, q, sub]) => {
    box(s, 0.5, y, 5.7, 1.2, WHITE, INK_SOFT, 1.5);
    s.addText([{ text: q, options: { fontFace: HEAD, bold: true, fontSize: 13.5, color: INK, breakLine: true, paraSpaceAfter: 3 } }, { text: sub, options: { fontFace: BODY, fontSize: 11.5, color: INK_SOFT } }], { x: 0.7, y, w: 5.3, h: 1.2, isTextBox: true, margin: 0, valign: "middle" });
  });
  arrow(s, 3.35, 3.15, 3.35, 3.55, INK_SOFT, { w: 1.4 }); label(s, "no", 3.45, 3.17, 0.5, 0.3, { size: 10, bold: true, align: "left" });
  arrow(s, 3.35, 4.75, 3.35, 5.15, INK_SOFT, { w: 1.4 }); label(s, "yes", 3.45, 4.77, 0.5, 0.3, { size: 10, bold: true, align: "left" });
  arrow(s, 6.2, 2.55, 7.25, 2.55, TEAL, { w: 1.8 }); label(s, "yes", 6.3, 2.25, 0.7, 0.28, { size: 10, bold: true, color: TEAL });
  arrow(s, 6.2, 4.15, 7.25, 3.5, TEAL, { w: 1.8 }); label(s, "no", 6.3, 4.25, 0.7, 0.28, { size: 10, bold: true, color: TEAL });
  arrow(s, 6.2, 5.75, 7.25, 5.75, PLUM, { w: 1.8 }); label(s, "yes", 6.3, 5.45, 0.7, 0.28, { size: 10, bold: true, color: PLUM });
  label(s, "no to all three? Either works. Pick the simpler one.", 0.5, 6.45, 5.7, 0.35, { size: 10, color: INK_SOFT, font: BODY });
  infoCard(s, 7.4, 1.85, 5.45, 2.2, TEAL, "Choose CP", ["Refuse or wait when unsure.", "One leader, a quorum, or consensus (etcd, ZooKeeper).", "Accept errors during a partition."], 12);
  infoCard(s, 7.4, 4.3, 5.45, 2.55, PLUM, "Choose AP", ["Keep answering with what you have.", "Replicas that catch up, merge rules, safe retries (idempotent writes).", "Accept old data for a while.", "Plan for conflicts: they will happen."], 12);
  footer(s, n, false);
});

// quiz: AP or CP?
add((n) => {
  const s = newSlide(false);
  banner(s, "Quiz · 3 min · in pairs", 0.5, 0.42, 3.0);
  s.addText("AP or CP? Pick one for each feature", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("The network between two data centers is cut. For each feature, should the system refuse and wait (CP), or answer with what it has and fix it later (AP)? Circle one.", { x: 0.5, y: 1.45, w: 12.3, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  [
    ["Bank withdrawal", "Pay out the money, or refuse until sure?"],
    ["The last seat for a concert", "Sell it, or ask the buyer to try again?"],
    ["A unique username", "Accept the name, or wait?"],
    ["Leader election and locks", "Elect a leader on this side, or wait?"],
    ["Shopping cart", "Accept the new item, or show an error?"],
    ["Like count", "Show the count you have, or an error?"],
    ["Social feed", "Show the posts you have, or an empty feed?"],
    ["Product catalog page", "Show a price that may be 30 s old, or nothing?"],
    ["Chat messages", "Send now and sort the order out later, or refuse?"],
  ].forEach(([t, q], i) => {
    const x = 0.5 + (i % 3) * 4.15, y = 2.1 + Math.floor(i / 3) * 1.6;
    card(s, x, y, 4.0, 1.45, WHITE);
    s.addText([
      { text: (i + 1) + " · " + t, options: { fontFace: HEAD, bold: true, fontSize: 13.5, color: INK, breakLine: true, paraSpaceAfter: 3 } },
      { text: q, options: { fontFace: BODY, fontSize: 11, color: INK_SOFT } },
    ], { x: x + 0.2, y: y + 0.1, w: 3.65, h: 0.85, isTextBox: true, margin: 0, valign: "top" });
    ["CP", "AP"].forEach((c, k) => {
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.2 + k * 1.0, y: y + 1.0, w: 0.85, h: 0.3, rectRadius: 0.15, fill: { type: "none" }, line: { color: k === 0 ? TEAL : PLUM, width: 1.5 } });
      label(s, c, x + 0.2 + k * 1.0, y + 1.0, 0.85, 0.3, { size: 10.5, bold: true, color: k === 0 ? TEAL : PLUM });
    });
  });
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Quiz: AP or CP?\n\nCP: 1 bank withdrawal (money paid out against an unknown balance), 2 the last seat (selling it twice is worse than a retry), 3 a unique username (two accounts with one name), 4 leader election and locks (two leaders is split brain).\n\nAP: 5 shopping cart (a missing item can be merged later, and a blocked cart loses a sale), 6 like count (off by a few is harmless), 7 social feed (a late post beats an empty feed), 8 product catalog (a price 30 seconds old is fine, but the price is read again from the CP side at checkout), 9 chat messages (send now, order per conversation).\n\nHardest to call: 5 and 8. Ask: what does it cost if we are wrong, and what does it cost if we refuse?");
});

// AP or CP by feature
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Quiz answers: choosing by feature", false, { sub: "One product uses both. Browsing the catalog is AP. Paying is CP." });
  tbl(s, 0.5, 1.85, [3.1, 3.6, 3.35, 2.3], [
    ["Feature", "If the answer is wrong", "If it is refused", "Pick"],
    ["Bank withdrawal", "Money paid out that is not there", "The customer retries later", "CP"],
    ["Last seat, last unit in stock", "Sold twice, then a refund", "\"Try again\" message", "CP"],
    ["Unique username", "Two accounts with one name", "The user retries", "CP"],
    ["Leader election, locks, config", "Two leaders (split brain)", "A short pause", "CP"],
    ["Shopping cart", "A missing item, merged later", "The customer cannot shop", "AP"],
    ["Like and view counts", "Off by a few", "A blank counter", "AP"],
    ["Social feed, search results", "A post shows up late", "An empty feed", "AP"],
    ["Product catalog page", "A price that is 30 seconds old", "Nobody can browse", "AP"],
    ["Chat messages", "Order differs for a moment", "Cannot send", "AP, ordered per chat"],
  ], { fs: 11, rowH: [0.4, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5] });
  s.addText("The price at checkout is read again from the CP side, so the catalog can be old and the payment cannot.", { x: 0.5, y: 6.55, w: 12.35, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12, color: INK_SOFT });
  footer(s, n, false);
});

// tunable consistency (overlap)
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Overlap: how to always read the latest value", false, { sub: "A value is stored on 3 copies. Write to 2 of them. Read from 2 of them. Any 2 of 3 must share at least one copy." });
  card(s, 0.5, 1.85, 7.4, 5.0, WHITE, { shadow: false, line: LINE });
  const xs = [0.95, 3.35, 5.75];
  [1, 2, 3].forEach((r, i) => {
    box(s, xs[i], 3.35, 1.7, 1.0, r === 2 ? AMBER_TINT : (r === 1 ? TEAL_TINT : PLUM_TINT), r === 2 ? AMBER : (r === 1 ? TEAL : PLUM), r === 2 ? 3 : 1.5);
    label(s, "Copy " + r, xs[i], 3.35, 1.7, 1.0, { size: 13, bold: true });
  });
  s.addShape(pptx.ShapeType.line, { x: xs[0], y: 3.1, w: xs[1] + 1.7 - xs[0], h: 0, line: { color: TEAL, width: 3 } });
  label(s, "the write goes to these 2", xs[0], 2.55, xs[1] + 1.7 - xs[0], 0.35, { size: 11, bold: true, color: TEAL });
  s.addShape(pptx.ShapeType.line, { x: xs[1], y: 4.6, w: xs[2] + 1.7 - xs[1], h: 0, line: { color: PLUM, width: 3 } });
  label(s, "the read asks these 2", xs[1], 4.75, xs[2] + 1.7 - xs[1], 0.35, { size: 11, bold: true, color: PLUM });
  s.addText("Copy 2 is in both sets, so the read always finds the newest value.", { x: 0.8, y: 5.55, w: 6.7, h: 0.9, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: BODY, bold: true, fontSize: 14, color: AMBER });
  infoCard(s, 8.15, 1.85, 4.7, 1.55, TEAL, "Write 2, read 2", "Reads see the latest value, and the system still works if one copy is down.", 12);
  infoCard(s, 8.15, 3.55, 4.7, 1.55, AMBER, "Write 1, read 1", "Fast and always available. But a read can land on a copy that has not heard the new value yet.", 12);
  infoCard(s, 8.15, 5.25, 4.7, 1.6, PLUM, "Write all 3, read 1", "Reads are cheap. But if one copy is down, writes are refused.", 12);
  footer(s, n, false);
  s.addNotes("The rule: copies written + copies read must be more than the number of copies (2 + 2 > 3). Cassandra and DynamoDB let you choose this per request. If the room finds it too deep, skip it.");
});

// quiz: what does each database design do when the network splits?
add((n) => {
  const s = newSlide(false);
  banner(s, "Quiz · 3 min · in pairs", 0.5, 0.42, 3.0);
  s.addText("What happens when the network splits?", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("For each database design from this part, circle what happens. Then say whether it leans toward consistency (C) or availability (A).", { x: 0.5, y: 1.45, w: 12.3, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  [
    ["Single primary, all reads from the primary", "Some servers are cut off from the primary."],
    ["Primary + async replicas serving reads", "Some replicas are cut off from the primary."],
    ["Multi-master", "Two data centers can no longer see each other."],
    ["Partitioned: each shard is a primary + replicas", "The network cuts off some of the shards."],
  ].forEach(([t, q], i) => {
    const x = 0.5 + (i % 2) * 6.3, y = 2.15 + Math.floor(i / 2) * 2.35;
    card(s, x, y, 6.05, 2.15, WHITE);
    s.addText([
      { text: (i + 1) + " · " + t, options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 4 } },
      { text: q, options: { fontFace: BODY, fontSize: 12, color: INK_SOFT } },
    ], { x: x + 0.25, y: y + 0.15, w: 5.6, h: 1.0, isTextBox: true, margin: 0, valign: "top" });
    [["Refuses", TEAL], ["Old data", AMBER], ["Both sides write", PLUM], ["Part still works", SLATE]].forEach(([c, col], k) => {
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.25 + k * 1.4, y: y + 1.55, w: 1.3, h: 0.36, rectRadius: 0.18, fill: { type: "none" }, line: { color: col, width: 1.5 } });
      label(s, c, x + 0.25 + k * 1.4, y + 1.55, 1.3, 0.36, { size: 8.5, bold: true, color: col });
    });
  });
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Quiz: what happens when the network splits?\n\n1 Single primary, all reads from the primary: Refuses. The cut-off side cannot reach the primary, so it refuses reads and writes. Leans toward consistency.\n\n2 Primary + async replicas serving reads: Old data. The replicas keep answering, maybe with old data, but writes still need the primary. Leans toward availability for reads and consistency for writes.\n\n3 Multi-master: Both sides write. Both sides keep accepting writes, then drift apart and are merged later, with the conflicts from earlier. Leans toward availability.\n\n4 Partitioned (each shard a primary + replicas): Part still works. Shards on the far side are unreachable, and the rest keep working. Consistency per key, with limited damage.");
});

add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Quiz answers: what each design gives up when the network splits", false, { sub: "Partitions are not optional, so the real choice is: during one, refuse some requests, or answer with data that may be old or different on each side." });
  s.addText("The same choice, applied to the database designs from this part.", { x: 0.5, y: 1.9, w: 12.3, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT, valign: "top" });
  tbl(s, 0.5, 2.75, [3.9, 5.5, 2.95], [
    ["Design", "When the network splits…", "Leans toward"],
    ["Single primary; all reads from the primary", "The cut-off side cannot reach the primary, so it refuses reads and writes.", "Consistency"],
    ["Primary + async replicas serving reads", "Replicas keep answering, maybe with old data; writes still need the primary.", "Availability for reads, consistency for writes"],
    ["Multi-master", "Both sides keep accepting writes, then drift apart and are merged later.", "Availability, plus the conflicts"],
    ["Partitioned (each shard primary + replicas)", "Shards on the far side are unreachable; the rest keep working.", "Consistency per key, limited damage"],
  ], { fs: 12, rowH: [0.45, 0.8, 0.8, 0.8, 0.8] });
  s.addText("CAP's \"consistency\" is not ACID's C. And it says nothing about the 99.9% of the time with no partition, when the tradeoff is speed (latency) versus consistency (PACELC). Day 7 returns to that.", { x: 0.5, y: 6.25, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
  s.addNotes("Talking points:\n- Each step gains something by giving up something: replicas give up freshness, multi-master gives up agreement, partitioning gives up simple queries. CAP is the same list, for the moment the network breaks.\n- 'Which should we use?' is the wrong first question. Ask what your service does if a user sees old data, and what it does if it refuses a write. The business already has an opinion.\n- No design here makes the problem disappear; each one only moves it.");
});

// ---- CAP in Kafka
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "The same choice inside a message queue", false, { sub: "Kafka makes the same trade with different words. Here each partition has 3 copies: a leader and two followers." });
  const scen = (y, title, cut1, cut2, caption, color) => {
    card(s, 0.5, y, 5.8, 2.35, WHITE, { shadow: false, line: LINE });
    label(s, title, 0.7, y + 0.1, 5.4, 0.3, { size: 11, bold: true, align: "left" });
    [[0.75, "follower 1", TEAL_TINT, TEAL], [2.85, "leader", TEAL_TINT, TEAL], [4.95, "follower 2", TEAL_TINT, TEAL]].forEach(([x, t, f, c], i) => {
      box(s, x, y + 0.6, 1.3, 0.7, i === 1 ? "BFD9DC" : f, c, i === 1 ? 2 : 1.25);
      label(s, t, x, y + 0.6, 1.3, 0.7, { size: 10, bold: true });
    });
    [[2.05, 2.85, cut1], [4.15, 4.95, cut2]].forEach(([xa, xb, cut], i) => {
      const ya = y + 0.95;
      arrow(s, i === 0 ? xb : xa, ya, i === 0 ? xa : xb, ya, cut ? RUST : TEAL, { w: 1.6, dash: cut ? "dash" : "solid", noHead: true });
      if (cut) label(s, "✗", xa + 0.2, ya - 0.28, 0.4, 0.3, { size: 14, bold: true, color: RUST });
    });
    label(s, caption, 0.7, y + 1.5, 5.4, 0.7, { size: 10.5, bold: true, color, font: BODY });
  };
  scen(1.85, "One follower cut off", false, true, "Leader + follower 1 = 2 copies. The write needs 2, so it is accepted.", TEAL);
  scen(4.4, "Both followers cut off", true, true, "Leader alone = 1 copy. The write needs 2, so the producer gets an error.", RUST);
  tbl(s, 6.6, 1.85, [2.45, 3.8], [
    ["Setting", "When copies are cut off…"],
    ["acks=1", "The leader confirms at once. If it dies before copying, the message is lost. Leans toward availability."],
    ["acks=all + min.insync.replicas=2", "A write needs 2 copies. With fewer reachable, the producer gets an error. Leans toward consistency."],
    ["Unclean leader election: on", "A copy that is behind may become leader. The topic stays up, but some messages are lost. Availability."],
    ["Unclean leader election: off (default)", "The partition waits until an up-to-date copy returns. Consistency."],
  ], { fs: 11, rowH: [0.4, 0.95, 1.05, 1.05, 0.95] });
  s.addText("Same trade as the database slide: refuse the producer, or accept the write and risk losing it.", { x: 6.6, y: 6.15, w: 6.25, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

// eventual consistency 1
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "What eventual consistency means", false, { sub: "Copies may disagree right now. If no new writes arrive, they end up the same. It does not say when." });
  card(s, 0.5, 1.85, 12.35, 2.95, WHITE, { shadow: false, line: LINE });
  const X0 = 2.0, W = 10.6, T0 = -40, T1 = 300, sc = W / (T1 - T0), tx = (t) => X0 + (t - T0) * sc;
  label(s, "write x = 5 on Node 1", tx(0) - 1.1, 2.0, 2.2, 0.25, { size: 9, bold: true, color: TEAL });
  [["Node 1", [[-40, 0, "amber", "x = 4"], [0, 300, "teal", "x = 5"]]], ["Node 2", [[-40, 80, "amber", "x = 4"], [80, 300, "teal", "x = 5"]]], ["Node 3", [[-40, 200, "amber", "x = 4  (old)"], [200, 300, "teal", "x = 5"]]]].forEach(([nm, segs], r) => {
    const y = 2.4 + r * 0.6;
    label(s, nm, 0.7, y, 1.2, 0.45, { size: 10.5, bold: true, align: "left" });
    segs.forEach(([a, b, col, tt]) => {
      s.addShape(pptx.ShapeType.rect, { x: tx(a), y, w: (b - a) * sc, h: 0.45, fill: { color: col === "teal" ? TEAL_TINT : AMBER_TINT }, line: { color: col === "teal" ? TEAL : AMBER, width: 1 } });
      label(s, tt, tx(a), y, (b - a) * sc, 0.45, { size: 9.5, bold: true, color: col === "teal" ? TEAL : AMBER });
    });
  });
  s.addShape(pptx.ShapeType.line, { x: tx(0), y: 2.3, w: 0, h: 1.95, line: { color: TEAL, width: 1.2, dashType: "dash" } });
  s.addShape(pptx.ShapeType.line, { x: tx(50), y: 2.3, w: 0, h: 1.95, line: { color: RUST, width: 1.2, dashType: "dash" } });
  s.addShape(pptx.ShapeType.line, { x: tx(300) - 0.02, y: 2.3, w: 0, h: 1.95, line: { color: TEAL, width: 1.2, dashType: "dash" } });
  label(s, "time →", 0.7, 4.3, 1.2, 0.25, { size: 8.5, color: INK_FAINT, align: "left" });
  label(s, "a read at 50 ms: Node 2 still says 4", tx(50) - 0.3, 4.3, 3.6, 0.3, { size: 9.5, bold: true, color: RUST, align: "left" });
  label(s, "a read at 300 ms: all say 5", 8.7, 4.3, 3.9, 0.3, { size: 9.5, bold: true, color: TEAL, align: "right" });
  [
    [PLUM, "Like count", "You see 1,204 likes. A friend sees 1,206. A few seconds later both see 1,207."],
    [AMBER, "DNS", "You change a website's address. Some people still reach the old one for minutes. Then everyone agrees."],
    [TEAL, "Shopping cart", "You add a book on your phone and a pen on your laptop. Each shows a different cart for a moment. Then they merge."],
  ].forEach(([c, t, d], i) => infoCard(s, 0.5 + i * 4.15, 5.0, 4.0, 1.85, c, t, d, 12));
  footer(s, n, false);
});

// eventual consistency 2
add((n) => {
  const s = newSlide(false);
  header(s, "Distributed systems", "Living with eventual consistency", false, { sub: "Alice changes her display name, then opens her profile." });
  card(s, 0.5, 1.75, 8.1, 5.2, WHITE, { shadow: false, line: LINE });
  const X = { A: 1.6, P: 4.5, R: 7.4 };
  [["Alice", X.A, TEAL_TINT, TEAL], ["Primary", X.P, "E3EAEC", SLATE], ["Replica", X.R, PLUM_TINT, PLUM]].forEach(([t, x, f, c]) => {
    box(s, x - 0.75, 1.88, 1.5, 0.45, f, c, 1.5); label(s, t, x - 0.75, 1.88, 1.5, 0.45, { size: 11, bold: true });
    s.addShape(pptx.ShapeType.line, { x, y: 2.33, w: 0, h: 6.85 - 2.33, line: { color: INK_FAINT, width: 1, dashType: "dash" } });
  });
  seqMsg(s, X.A, X.P, 2.95, "1  set name = Sam");
  seqMsg(s, X.P, X.A, 3.45, "saved", { dash: true });
  seqMsg(s, X.A, X.R, 4.15, "2  open profile (read from a replica)");
  seqMsg(s, X.R, X.A, 4.75, "name = Alex   (old)", { dash: true, color: AMBER, textColor: AMBER, bold: true });
  seqMsg(s, X.P, X.R, 5.6, "3  replicate: arrives 2 s later", { dash: true, color: TEAL, textColor: TEAL });
  label(s, "Alice thinks her change was lost.", 1.0, 6.05, 6.9, 0.4, { size: 11, bold: true, color: RUST });
  sidePanel(s, [
    [RUST, "What went wrong", "Alice saved on the primary, then read from a replica that had not caught up. She sees the old name."],
    [TEAL, "Fix 1: read your own writes", "For a few seconds after a write, send that user's reads to the primary, or remember which version they wrote."],
    [PLUM, "Fix 2: show it right away", "The app shows the new name at once and confirms later. Other people may see the old name for a moment, and that is fine."],
  ]);
  footer(s, n, false);
});

checkIn(10, 2, "Refuse, or answer with old data?", "The network between two data centers is cut. For each feature, should the system refuse the request, or answer with data that may be out of date?", [
  "A bank withdrawal.",
  "The like count on a post.",
  "Adding an item to a shopping cart.",
  "Reserving seat 14C for a concert.",
], null, "1. Refuse. Paying out against an unknown balance costs real money.\n2. Answer with old data. A count a few seconds old hurts nobody, and refusing page views would be worse.\n3. Accept the write. A cart that keeps working and merges later is better than an error.\n4. Refuse. Selling the same seat twice is worse than asking one customer to try again.\n\nThe right answer depends on what it costs to be wrong. That is a business question, not a technical one.");

// ============================================================ HOW FAILURE SPREADS
add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Failure", "How a failure spreads", false, { sub: "A database gets slow. Nothing is down yet. Follow it up the chain." });
  label(s, "the slowness travels back toward the user", 1.7, 1.95, 9.5, 0.28, { size: 10.5, bold: true, color: RUST });
  arrow(s, 11.3, 2.35, 1.7, 2.35, RUST, { w: 2 });
  [["Web tier", "users refresh, clients retry", TEAL_TINT, TEAL, 0.5], ["Orders API", "threads wait, calls pile up", AMBER_TINT, AMBER, 3.7], ["Inventory API", "all 50 threads waiting", RUST_TINT, RUST, 6.9], ["Database", "slow, not down: 2 s per query", RUST_TINT, RUST, 10.1]].forEach(([t, sub, f, c, x], i) => {
    box(s, x, 2.7, 2.3, 1.1, f, c, i === 3 ? 2.5 : 1.5);
    label(s, t, x, 2.78, 2.3, 0.45, { size: 12, bold: true });
    label(s, sub, x + 0.1, 3.22, 2.1, 0.5, { size: 9, color: INK_SOFT });
    if (i < 3) arrow(s, x + 2.3, 3.25, x + 3.2, 3.25, INK_SOFT, { w: 1.4 });
  });
  const cards = [
    [AMBER, "Waiting", "A slow service holds its callers' threads and connections. Slow is worse than down: a down service fails fast, a slow one ties everyone up."],
    [RUST, "Retries", "Each retry is more load on the service that is already struggling, so it gets slower, so more calls time out, so there are more retries."],
    [PLUM, "Shared resources", "One connection pool, CPU or database serves many callers. One busy caller can starve all the others."],
    [SLATE, "Shared fate", "Same deploy, same zone, same config. One bad change breaks many services at once."],
  ];
  cards.forEach((c, i) => {
    const x = 0.5 + i * 3.15, y = 4.3, w = 2.95, h = 2.3;
    card(s, x, y, w, h, WHITE);
    s.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.09, fill: { color: c[0] }, line: { type: "none" } });
    s.addText([
      { text: c[1], options: { bold: true, color: c[0], fontFace: HEAD, fontSize: 15, breakLine: true, paraSpaceAfter: 6 } },
      { text: c[2], options: { color: INK_SOFT, fontFace: BODY, fontSize: 12 } },
    ], { x: x + 0.2, y: y + 0.25, w: w - 0.4, h: h - 0.35, isTextBox: true, margin: 0, valign: "top" });
  });
  footer(s, n, false);
  s.addNotes("Four ways a failure spreads: waiting (threads held), retries (extra load), shared resources (one pool or database for many callers), shared fate (same deploy, zone or config). Most real incidents use at least two at once.");
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Failure", "Dissecting a cascading failure: the timeline", false, { sub: "One server out of four crashes at 12:00. Twenty-five minutes later the whole shop is down." });
  tbl(s, 0.5, 1.85, [1.0, 5.6, 5.75], [
    ["Time", "What happened", "Why it spread"],
    ["12:00", "One of four inventory servers crashes (a memory leak).", "The trigger is small. The other three now carry 33% more traffic each."],
    ["12:02", "The three left run at 95% CPU. Replies slow from 80 ms to 900 ms.", "Slow, not down. Callers keep waiting instead of failing."],
    ["12:03", "The Orders API gives up after 1 second and retries, up to 3 times.", "Amplifier: traffic to inventory is now about 3 times higher."],
    ["12:05", "All Orders API threads are stuck waiting. It stops answering health checks, so the load balancer removes it.", "The failure moves up a layer. The remaining Orders servers get even more traffic."],
    ["12:10", "The team restarts the crashed server and adds new ones.", "New servers start with empty caches and meet the full retry backlog. They fall over too."],
    ["12:25", "The shop recovers only after traffic is cut at the load balancer and retries are turned off.", "While the load stayed high, the system could not heal itself."],
  ], { fs: 11.5, rowH: [0.4, 0.62, 0.68, 0.68, 0.85, 0.8, 0.8] });
  footer(s, n, false);
  s.addNotes("Ask the room first: at which minute did this stop being 'one server crashed' and become 'the shop is down'? (12:03, when retries tripled the load.) Then walk the rows.");
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Failure", "Dissecting it: trigger, amplifiers, recovery trap", false, { sub: "The crash was not the real problem. What the system did next was." });
  const cols = [
    [SLATE, "TRIGGER", "One server out of four crashes.", ["A normal event. The system should survive it.", "It removed only 25% of capacity."]],
    [AMBER, "AMPLIFIERS", "Things that made it worse.", ["Retries multiply the load.", "Waiting threads fill up.", "Health checks remove slow servers, so capacity shrinks.", "New servers start with empty caches."]],
    [RUST, "RECOVERY TRAP", "Why it did not heal.", ["The retry backlog hits every new server.", "The overload keeps itself going, even after the trigger is gone.", "This is called a metastable failure."]],
  ];
  cols.forEach((c, i) => {
    const x = 0.5 + i * 4.15, y = 1.85, w = 4.0, h = 2.65;
    card(s, x, y, w, h, WHITE);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.09, h, fill: { color: c[0] }, line: { type: "none" } });
    s.addText([
      { text: c[1], options: { fontFace: CODE, bold: true, fontSize: 11, color: c[0], charSpacing: 1, breakLine: true, paraSpaceAfter: 2 } },
      { text: c[2], options: { fontFace: HEAD, bold: true, fontSize: 14, color: INK, breakLine: true, paraSpaceAfter: 6 } },
    ].concat(bl(c[3], { fs: 12, color: INK_SOFT, gap: 4 })), { x: x + 0.3, y: y + 0.15, w: w - 0.5, h: h - 0.25, isTextBox: true, margin: 0, valign: "top" });
  });
  label(s, "RETRY MATH: ONE CLICK, THREE LAYERS, 3 TRIES AT EACH LAYER", 0.5, 4.75, 12.3, 0.28, { size: 10, bold: true, color: INK_SOFT, align: "left" });
  [["User click", "1 call", TEAL_TINT, TEAL], ["Orders API", "3 calls", AMBER_TINT, AMBER], ["Inventory API", "9 calls", RUST_TINT, RUST], ["Database", "27 calls", RUST_TINT, RUST]].forEach(([t, c, f, l], i) => {
    const x = 0.5 + i * 3.15;
    box(s, x, 5.15, 2.3, 0.95, f, l, i === 3 ? 2.5 : 1.5);
    label(s, t, x, 5.2, 2.3, 0.4, { size: 11, bold: true });
    label(s, c, x, 5.6, 2.3, 0.4, { size: 14, bold: true, color: i === 0 ? INK : RUST });
    if (i < 3) arrow(s, x + 2.3, 5.62, x + 3.1, 5.62, INK_SOFT, { w: 1.4 });
  });
  s.addText("If every layer retries 3 times, one click can become 27 database calls. Retry at one layer only.", { x: 0.5, y: 6.3, w: 12.3, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
});

add((n) => {
  const s = newSlide(false);
  header(s, "Part 3 · Failure", "Breaking the chain: one defence for each step", false, { sub: "Every row here would have stopped, or shrunk, a step of the incident we just dissected." });
  tbl(s, 0.5, 1.85, [3.0, 3.3, 6.05], [
    ["Step in the incident", "Defence", "How it works"],
    ["Callers wait for a slow service", "Short timeouts", "Give the whole request a time budget (for example 1 s). Pass what is left to the next call."],
    ["Retries triple the load", "Retry budget, backoff, jitter", "Retry only if retries are under about 10% of calls. Wait longer each time, with a random delay."],
    ["A slow service keeps getting called", "Circuit breaker", "After a number of failures, stop calling for 30 s and return a fallback instead."],
    ["Overload makes everything slow", "Load shedding", "Reject the extra requests early (HTTP 429 or 503) so the rest finish. Keep checkout, drop browsing first."],
    ["One busy caller starves the rest", "Bulkheads", "Separate thread and connection pools for each dependency, so one cannot use them all."],
    ["New servers fall over when they start", "Slow start", "Send a trickle of traffic to new servers and ramp up while their caches warm."],
    ["The system cannot recover by itself", "A kill switch", "A flag to turn off retries and optional features, so people can cut the load fast."],
  ], { fs: 11, rowH: [0.4, 0.62, 0.62, 0.62, 0.7, 0.62, 0.62, 0.62] });
  footer(s, n, false);
  s.addNotes("Ask: if you could add only one of these before the next Black Friday, which one, and why? Most teams say timeouts, because the rest need them to work.");
});

// ============================================================ HANDS-ON
add((n) => {
  const s = newSlide(false);
  header(s, "2:40–3:10 · Hands-on A", "Capacity-plan a real service", false, { sub: "Work through the example, then do the same math for a service you own. Pairs, 30 minutes." });
  tbl(s, 0.5, 1.95, [3.3, 6.0, 3.05], [
    ["Line item", "Worked example (checkout)", "Yours"],
    ["Current peak traffic", "800 req/s", "…"],
    ["Expected traffic (6 mo)", "2,400 req/s (3× growth)", "…"],
    ["Capacity per instance", "750 req/s at 70% CPU", "…"],
    ["Instances needed", "2,400 ÷ 750 = 3.2 → 4", "…"],
    ["Spare capacity (for failover)", "+1 for failover = 5 total", "…"],
    ["Actually stateless?", "Yes — after the Redis fix", "…"],
    ["Hottest shared resource", "One product row: ~5 buys/s with a payment call inside the lock; ~5,000/s after the one-step update", "…"],
    ["Work that can leave the request path", "Receipt email, loyalty update, warehouse notify → a queue (Part 2)", "…"],
  ], { fs: 11.5, rowH: [0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.62, 0.55] });
  s.addText("Why 70%, not 95%? Waiting time grows roughly like ρ/(1−ρ), where ρ is how busy the system is: 2.3 at 70% busy, 9 at 90%, 19 at 95%. The spare room is what keeps the queue from blowing up.", { x: 0.5, y: 6.1, w: 12.3, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  footer(s, n, false);
  s.addNotes("Questions for pairs:\n- What is your service's real bottleneck: CPU, memory, or something it depends on?\n- If it is not stateless, what is the smallest piece of state to move out first?\n- Which of the three layers (web servers, background jobs, database) hits its limit first, and what is the cheapest way to raise it?\n- Is your database split, replicated, or a single server?\n\nDone when: you have numbers for every row; your server count includes spare capacity; you know whether the service is stateless; you named the shared resource that limits it first and one piece of work a queue could take off the request path.\n\nExample restated: 800 req/s today, 2,400 expected, 750 req/s per server at 70% CPU: 2400/750 = 3.2 -> 4, plus 1 for failover = 5. None of it matters until the cart state moves to Redis.");
});

// Exercise 1 diagram
add((n) => {
  const s = newSlide(false);
  banner(s, "Exercise 1 · 10 min", 0.5, 0.42, 2.6);
  s.addText("The Black Friday storefront", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("The sale opened an hour ago. Pages that took 400 ms now take over six seconds. The on-call engineer suggests \"add more app servers\". Most traffic is catalog pages, and every page view reads the catalog from the database.", { x: 0.5, y: 1.5, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  box(s, 0.4, 3.35, 1.3, 0.9, WHITE, INK_SOFT); label(s, "Clients", 0.4, 3.4, 1.3, 0.45, { size: 11 }); label(s, "6,000 req/s", 0.4, 3.82, 1.3, 0.3, { size: 8.5, color: INK_FAINT });
  box(s, 2.0, 3.35, 1.5, 0.9, AMBER_TINT, AMBER, 1.5); label(s, "Load balancer", 2.0, 3.4, 1.5, 0.45, { size: 10 }); label(s, "ip_hash (sticky)", 2.0, 3.82, 1.5, 0.3, { size: 8.5, color: INK_FAINT });
  [["app-1", "31%", 2.35], ["app-2", "28%", 3.35], ["app-3", "33%", 4.35]].forEach((a) => {
    box(s, 3.95, a[2], 1.9, 0.82, TEAL_TINT, TEAL, 1.25);
    label(s, a[0] + "  ·  CPU " + a[1], 3.95, a[2] + 0.05, 1.9, 0.4, { size: 9.5, bold: true });
    label(s, "sessions in memory", 3.95, a[2] + 0.42, 1.9, 0.3, { size: 8.5, color: INK_FAINT });
    arrow(s, 3.5, 3.8, 3.92, a[2] + 0.41, INK_SOFT, { w: 1 });
  });
  arrow(s, 1.7, 3.8, 1.97, 3.8, INK_SOFT, { w: 1 });
  box(s, 6.65, 2.95, 2.1, 1.75, RUST_TINT, RUST, 2.5);
  label(s, "Postgres primary", 6.65, 3.02, 2.1, 0.35, { size: 10 });
  label(s, "CPU 98%", 6.65, 3.38, 2.1, 0.35, { size: 12, bold: true, color: RUST });
  label(s, "connections 100/100", 6.65, 3.75, 2.1, 0.3, { size: 9 });
  label(s, "95% reads", 6.65, 4.05, 2.1, 0.3, { size: 9 });
  box(s, 6.65, 5.0, 2.1, 0.85, WHITE, INK_SOFT); label(s, "Email API", 6.65, 5.05, 2.1, 0.4, { size: 10 }); label(s, "called inline, ~300 ms", 6.65, 5.42, 2.1, 0.3, { size: 8.5, color: INK_FAINT });
  [2.76, 3.76, 4.54].forEach((y, i) => arrow(s, 5.85, y, 6.6, [3.3, 3.82, 4.3][i], INK_SOFT, { w: 1 }));
  arrow(s, 5.85, 5.0, 6.6, 5.35, INK_SOFT, { w: 1 });
  label(s, "p99 latency", 0.5, 5.2, 1.6, 0.25, { size: 9, color: INK_FAINT, align: "left" }); label(s, "6.2 s", 0.5, 5.45, 1.6, 0.4, { size: 16, bold: true, color: RUST, align: "left" });
  label(s, "app CPU", 2.2, 5.2, 1.6, 0.25, { size: 9, color: INK_FAINT, align: "left" }); label(s, "~30%", 2.2, 5.45, 1.6, 0.4, { size: 16, bold: true, align: "left" });
  card(s, 9.2, 2.45, 3.65, 4.4, WHITE);
  qBlock(s, 9.4, 2.6, 3.3, 4.1, "Questions", ["Where is the bottleneck, and which numbers prove it?", "If the team adds five more app servers, what happens to latency, and why?", "List your changes, in order. Which gives the biggest relief for the least work?", "Two other things here are not causing harm yet but will. Which?"], 12);
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Exercise 1\n\nBottleneck: the database. App CPU is near 30% while the database is at 98% CPU with 100/100 connections. The app servers are mostly waiting for a connection, and that is where the 6.2 s p99 comes from. Adding app servers makes it worse: each new server opens more connections against a pool that is already full.\n\nFixes, in order:\n1. Cache the product catalog (Redis, or a CDN for page parts). 95% of the load is reads of data that rarely changes.\n2. Add a connection pooler, then read replicas for reads the cache does not cover, sending a user's reads to the primary shortly after they write.\n3. Move the email call to a queue. It is not the bottleneck today, but it puts a 300 ms outside service inside every checkout.\n4. Store sessions in Redis and drop ip_hash.\n\nHidden problems: sticky sessions with in-memory sessions (a server replaced mid-sale logs its users out; heavy users can overload one server) and the inline email call (a slow provider slows or breaks checkout).");
});

// Exercise 2 diagram
add((n) => {
  const s = newSlide(false);
  banner(s, "Exercise 2 · 10 min", 0.5, 0.42, 2.6);
  s.addText("The event pipeline that won't catch up", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("Order events flow through Kafka to a service that writes to a database. Lag has grown for two days, so the team tripled the consumers from 2 to 6. The lag graph did not move. The topic is keyed by merchant_id, and one big marketplace merchant is now most of the traffic.", { x: 0.5, y: 1.5, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  box(s, 0.4, 3.3, 1.6, 1.15, WHITE, INK_SOFT);
  label(s, "order-service", 0.4, 3.35, 1.6, 0.35, { size: 9.5 }); label(s, "key = merchant_id", 0.4, 3.68, 1.6, 0.3, { size: 8.5, color: INK_FAINT }); label(s, "merchant 7: 68%", 0.4, 3.98, 1.6, 0.3, { size: 8.5, color: RUST, bold: true });
  const lanes = [[2.45, "P0", "lag 480,000", true], [3.4, "P1", "lag 2,100", false], [4.35, "P2", "lag 900", false]];
  lanes.forEach((l) => {
    box(s, 2.7, l[0], 2.5, 0.75, l[3] ? RUST_TINT : WHITE, l[3] ? RUST : INK_SOFT, l[3] ? 2 : 1.1);
    label(s, l[1], 2.8, l[0], 0.5, 0.75, { size: 11, bold: true, align: "left" });
    label(s, l[2], 3.3, l[0], 1.8, 0.75, { size: 10.5, bold: l[3], color: l[3] ? RUST : INK, align: "left" });
    arrow(s, 2.0, 3.87, 2.65, l[0] + 0.37, INK_SOFT, { w: 1 });
  });
  label(s, "TOPIC: 3 partitions", 2.7, 2.15, 2.5, 0.25, { size: 9, color: TEAL, bold: true });
  s.addShape(pptx.ShapeType.roundRect, { x: 5.8, y: 2.1, w: 1.75, h: 4.1, rectRadius: 0.08, fill: { type: "none" }, line: { color: PLUM, width: 1.1, dashType: "dash" } });
  const cons = [["c1", "CPU 100%", true], ["c2", "CPU 4%", false], ["c3", "CPU 3%", false], ["c4", "idle", null], ["c5", "idle", null], ["c6", "idle", null]];
  cons.forEach((c, i) => {
    const y = 2.2 + i * 0.66;
    s.addShape(pptx.ShapeType.roundRect, { x: 5.95, y, w: 1.45, h: 0.55, rectRadius: 0.05, fill: c[2] === true ? { color: RUST_TINT } : c[2] === false ? { color: PLUM_TINT } : { type: "none" }, line: { color: c[2] === true ? RUST : c[2] === false ? PLUM : INK_FAINT, width: 1, dashType: c[2] === null ? "dash" : "solid" } });
    label(s, c[0] + " · " + c[1], 5.95, y, 1.45, 0.55, { size: 8.5, color: c[2] === null ? INK_FAINT : INK, bold: c[2] === true });
  });
  arrow(s, 5.2, 2.82, 5.9, 2.47, INK_SOFT, { w: 1 }); arrow(s, 5.2, 3.77, 5.9, 3.13, INK_SOFT, { w: 1 }); arrow(s, 5.2, 4.72, 5.9, 3.79, INK_SOFT, { w: 1 });
  box(s, 7.95, 3.45, 1.3, 0.95, WHITE, INK_SOFT); label(s, "Postgres", 7.95, 3.5, 1.3, 0.4, { size: 10 }); label(s, "one writer", 7.95, 3.9, 1.3, 0.3, { size: 8.5, color: INK_FAINT });
  arrow(s, 7.4, 2.5, 7.9, 3.7, INK_SOFT, { w: 1 }); arrow(s, 7.4, 3.15, 7.9, 3.85, INK_SOFT, { w: 1 }); arrow(s, 7.4, 3.8, 7.9, 4.0, INK_SOFT, { w: 1 });
  label(s, "consumers scaled 2 → 6:", 0.5, 5.3, 4.5, 0.3, { size: 9.5, color: INK_FAINT, align: "left" }); label(s, "total lag unchanged", 0.5, 5.6, 4.5, 0.4, { size: 16, bold: true, color: RUST, align: "left" });
  card(s, 9.7, 2.45, 3.15, 4.4, WHITE);
  qBlock(s, 9.9, 2.6, 2.8, 4.1, "Questions", ["Consumers went 2 → 6 and lag did not improve. Give two separate reasons.", "Why is one partition far behind the other two?", "Propose two fixes and what each costs.", "Once the lag clears, what gives out next?"], 11.5);
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Exercise 2\n\nTwo reasons the extra consumers did nothing: (1) parallel work in a consumer group is limited by the number of partitions: three partitions means three active consumers, so consumers 4 to 6 are idle; (2) the load is not spread across those three: merchant 7 is 68% of traffic and the key is merchant_id, so almost everything goes to P0, and one consumer is overloaded while the other two sit at 3-4%. More consumers cannot split one partition.\n\nFixes and costs:\n- Change the key (order_id, or merchant_id plus a small random part for the big merchant): spreads load right away, but order across one merchant's events is no longer guaranteed, so check whether anything depends on it.\n- Add partitions (say 12): raises the limit, but changes where existing keys go, which can break order during the switch, and does nothing for a hot key by itself.\n- Make the consumer faster (batch the inserts): no ordering risk, but it only helps by a constant amount.\n\nNext to give out: the single Postgres writer. Batching, idempotent upserts, or partitioning the table come next. It is the Part 1 lock problem one layer down.");
});

// Exercise 3 diagram
add((n) => {
  const s = newSlide(false);
  banner(s, "Exercise 3 · 10 min", 0.5, 0.42, 2.6);
  s.addText("More replicas, same outage", { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
  s.addText("One primary, three read replicas. Writes time out at peak. Adding the third replica barely changed anything. Support also gets tickets: \"my order confirmation shows my old address\". The orders table is 2.1 TB and growing.", { x: 0.5, y: 1.5, w: 12.3, h: 0.7, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT, valign: "top" });
  box(s, 0.4, 3.35, 1.5, 1.2, WHITE, INK_SOFT); label(s, "App tier", 0.4, 3.4, 1.5, 0.4, { size: 10.5 }); label(s, "writes 4,200/s", 0.4, 3.8, 1.5, 0.3, { size: 8.5, color: INK_FAINT }); label(s, "reads 11,000/s", 0.4, 4.08, 1.5, 0.3, { size: 8.5, color: INK_FAINT });
  box(s, 3.2, 2.95, 2.1, 2.0, RUST_TINT, RUST, 2.5);
  label(s, "Primary", 3.2, 3.1, 2.1, 0.35, { size: 11 }); label(s, "CPU 97%", 3.2, 3.5, 2.1, 0.4, { size: 13, bold: true, color: RUST }); label(s, "all writes land here", 3.2, 3.95, 2.1, 0.3, { size: 9 }); label(s, "orders table: 2.1 TB", 3.2, 4.25, 2.1, 0.3, { size: 9 });
  arrow(s, 1.9, 3.7, 3.15, 3.7, RUST, { w: 1.75 }); label(s, "all writes", 1.95, 3.35, 1.2, 0.25, { size: 8.5, color: RUST });
  [[2.2, "1"], [3.45, "2"], [4.7, "3"]].forEach((r) => {
    box(s, 6.8, r[0], 1.9, 1.0, TEAL_TINT, TEAL, 1.25);
    label(s, "replica " + r[1], 6.8, r[0] + 0.05, 1.9, 0.35, { size: 10 }); label(s, "CPU 18%", 6.8, r[0] + 0.38, 1.9, 0.28, { size: 9 }); label(s, "lag p99 4 s", 6.8, r[0] + 0.64, 1.9, 0.28, { size: 8.5, color: INK_FAINT });
    arrow(s, 5.3, 3.95, 6.75, r[0] + 0.5, TEAL, { w: 1, dash: "dash" });
  });
  label(s, "async replication", 5.2, 3.0, 1.6, 0.25, { size: 8.5, color: TEAL });
  s.addShape(pptx.ShapeType.line, { x: 1.15, y: 4.55, w: 0, h: 1.6, line: { color: INK_SOFT, width: 1.1 } });
  s.addShape(pptx.ShapeType.line, { x: 1.15, y: 6.15, w: 8.15, h: 0, line: { color: INK_SOFT, width: 1.1 } });
  s.addShape(pptx.ShapeType.line, { x: 9.3, y: 2.7, w: 0, h: 3.45, line: { color: INK_SOFT, width: 1.1 } });
  [2.7, 3.95, 5.2].forEach((y) => arrow(s, 9.3, y, 8.75, y, INK_SOFT, { w: 1.1 }));
  label(s, "reads, spread across all three replicas", 2.5, 5.85, 5.5, 0.28, { size: 9, color: INK_FAINT });
  card(s, 9.75, 2.45, 3.1, 4.4, WHITE);
  qBlock(s, 9.95, 2.6, 2.75, 4.1, "Questions", ["Which node is the bottleneck? Why did a third replica barely help?", "Explain the \"old address\" tickets from the diagram alone.", "Suggest a way to scale writes. Which shard key, and which would you reject?", "What gets harder afterward? Name two things."], 11.5);
  footer(s, n, false);
  s.addNotes("ANSWER KEY — Exercise 3\n\nBottleneck: the primary. It is at 97% CPU taking every write (4,200/s) while the replicas sit under 20%. Replicas only serve reads, and each also replays the same writes, so a third replica adds nothing for write capacity. Replicas scale reads, never writes.\n\nThe 'old address' tickets are replication lag: replication is asynchronous with a p99 lag of 4 s. A customer saves an address (a write to the primary), and then the confirmation page reads from a replica that has not received it. Fix: for a short time after a user writes, read from the primary (or use a session-consistent read).\n\nScaling writes: partition (shard) the orders table. Shard key: customer_id, because the main queries are 'this customer's orders', which then touch exactly one shard; consistent hashing keeps resharding to a small fraction of keys. Each shard is its own primary with replicas, so the replica servers are reused. Reject: a date or auto-increment id as a range key (the newest shard takes every new write), and order_id hashed (spreads writes but scatters each customer's orders across all shards). Multi-master is not the answer: two nodes accepting writes to the same order is the conflict problem from Part 3.\n\nWhat gets harder: queries and transactions that cross shards (reports, joins across customers), a very large customer making one shard hot, and changing the shard layout later. Analytics and cross-customer reports usually move to a data warehouse.");
});

// ============================================================ SYSTEM DESIGN QUESTIONS
function designSlide(num, title, scenario, stats, start, questions, answer) {
  add((n) => {
    const s = newSlide(false);
    banner(s, "System design " + num + " · 10 min", 0.5, 0.42, 2.9);
    s.addText(title, { x: 0.5, y: 0.85, w: 12.3, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 28, color: INK });
    s.addText(scenario, { x: 0.5, y: 1.5, w: 12.3, h: 0.75, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT, valign: "top" });
    // numbers
    stats.forEach((st, i) => {
      const x = 0.5 + i * 2.15;
      card(s, x, 2.45, 1.95, 1.05, WHITE);
      label(s, st[0], x, 2.52, 1.95, 0.5, { size: 17, bold: true, color: st[2] || INK });
      label(s, st[1], x + 0.08, 3.0, 1.79, 0.45, { size: 8.5, color: INK_FAINT });
    });
    // starting sketch
    label(s, "START HERE", 0.5, 3.7, 3, 0.26, { size: 9.5, bold: true, color: TEAL, align: "left" });
    start.forEach((b, i) => {
      const x = 0.5 + i * 2.95;
      box(s, x, 4.0, 2.2, 0.95, WHITE, INK_SOFT, 1.25);
      label(s, b[0], x, 4.05, 2.2, 0.45, { size: 10.5, bold: true });
      label(s, b[1], x, 4.5, 2.2, 0.35, { size: 8.5, color: INK_FAINT });
      if (i < start.length - 1) arrow(s, x + 2.2, 4.47, x + 2.9, 4.47, INK_SOFT, { w: 1.25 });
    });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 5.3, w: 8.4, h: 1.55, rectRadius: 0.1, fill: { type: "none" }, line: { color: INK_FAINT, width: 1.25, dashType: "dash" } });
    s.addText("Sketch your design here. What do you add, and where? Write the numbers next to each box.", { x: 0.7, y: 5.3, w: 8.0, h: 1.55, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: BODY, italic: true, fontSize: 13, color: INK_FAINT });
    // questions
    card(s, 9.2, 2.45, 3.65, 4.4, WHITE);
    qBlock(s, 9.4, 2.6, 3.3, 4.1, "Questions", questions, 11.5);
    footer(s, n, false);
    s.addNotes("ANSWER KEY — System design " + num + "\n\n" + answer);
  });
}

designSlide(1, "The ticket drop",
  "Tickets for a stadium concert go on sale at 10:00. In the first minute, 200,000 fans open the page for 5,000 seats. Two fans must never get the same seat, and the site must stay up. Payment takes about 5 seconds.",
  [["200,000", "fans in the first minute", RUST], ["5,000", "seats, each sold once", null], ["~5 s", "payment time per buyer", null], ["0", "double-sold seats allowed", RUST]],
  [["Fans", "browser or app"], ["App servers", "stateless"], ["Database", "seats table"]],
  ["Which resource is the hottest, and what breaks first?", "How do you stop two fans buying seat 14C without holding a lock during the 5 s payment?", "What do the other 195,000 fans see, and where does that waiting happen?", "A payment confirmation arrives twice. What stops a double charge?"],
  "Hot spot: the seat rows (and any 'seats left' counter). Everyone wants the same few rows, so a lock held during payment would cap sales at about 1 buyer per 5 s per seat.\n\nHold, don't lock. Reserve the seat with one conditional UPDATE: set status = held, held_by = this fan, held_until = now + 5 min, WHERE seat_id = ? AND (status = free OR held_until < now). One row changed means the fan has the seat; zero rows means someone else does. Pay outside any lock. On success, a second conditional UPDATE (WHERE held_by = this fan) marks it sold. If the hold expires, the seat becomes free again.\n\nWaiting: put a waiting room in front. A queue hands out tokens and admits a few hundred fans per minute to the buying pages. Static pages and the seat map come from a CDN or cache (a few seconds old is fine). Only the buy step needs the database. Keep app servers stateless (token or shared session store) so you can add servers during the rush.\n\nDouble charge: send an idempotency key (the order ID) to the payment provider and record handled payment event IDs in the same transaction as the 'sold' update (the Part 2 duplicate fix).\n\nExtra credit: read replicas serve the seat map (lag is acceptable there), but the buy step must hit the primary.");

designSlide(2, "The viral link counter",
  "A link shortener shows \"clicked N times\" on each link. One link goes viral: 50,000 clicks per second on the same link. Counts may be a few seconds late, but no click may be lost. Most traffic is redirects, which only read.",
  [["50,000/s", "clicks on one link", RUST], ["100 : 1", "redirects per link created", null], ["5 s", "late counts are fine", null], ["0", "lost clicks allowed", RUST]],
  [["Visitors", "click a short link"], ["App servers", "redirect + count"], ["Database", "UPDATE clicks = clicks + 1"]],
  ["Why does one UPDATE per click fall over at 50,000/s?", "How can you count without 50,000 writes per second to one row?", "Where should the redirects come from so the database barely sees them?", "A counting worker crashes mid-batch. How do you avoid losing or double-counting?"],
  "Why it fails: this is the Part 1 hot row. Every click takes the same row lock, so clicks are handled one after another. If each update holds the lock for 1 ms, the ceiling is about 1,000 per second, and the waiting requests pile up and take the redirects down with them.\n\nFix 1, batch: the app puts a click event on a queue (key = link ID, with a random part for the viral link so it does not become a hot partition). A consumer adds up clicks in memory and writes one 'clicks + 4,812' update per link per second. 50,000 writes per second become about 1.\n\nFix 2, split the counter: keep, say, 10 rows per link and add to a random one; the page sums them. Fewer collisions on each row.\n\nRedirects: a short link's target never changes, so serve it from a cache or CDN with a long TTL. The database is not involved. Reads can also use replicas.\n\nCrash: save the offset after the batch is written (at-least-once). To avoid double-counting, write the batch with a batch ID in the same transaction as the update, so a repeated batch is skipped. For a click counter, a tiny error may be acceptable; for billing it would not be.");

designSlide(3, "Ten million push notifications",
  "After a cup final, the app must send a push notification to 10 million users within 10 minutes. The push provider accepts 20,000 messages per second and sometimes returns errors. A user must not get the same notification twice.",
  [["10 M", "users to notify", null], ["10 min", "time allowed", null], ["20,000/s", "provider rate limit", RUST], ["~1%", "provider errors", RUST]],
  [["Event", "match ends"], ["Notification service", "builds the list"], ["Push provider", "20,000/s limit"]],
  ["What rate do you need, and can one server loop over 10 M users?", "What goes on the queue, and how many partitions and consumers do you need if one consumer sends 500/s?", "One user's device token fails every time. What happens to the messages behind it?", "A consumer crashes and its messages come back. How do you avoid a second push?"],
  "Rate: 10,000,000 users in 600 s is about 16,700 per second. That is close to the provider limit of 20,000, so there is little headroom and the limit must be enforced.\n\nDesign: the service splits the audience into batches (for example 1,000 users each) and puts one message per batch on the queue. Order does not matter, so use a random key. Consumers send the batch to the provider. At 500/s per consumer you need at least 34 consumers, so create about 48 partitions and run up to 48 consumers, leaving room to grow and to retry.\n\nRate limit: the sum of all consumers must stay under 20,000/s. Use a shared token bucket (for example in Redis) or fix the consumer count, so scaling up cannot get you blocked by the provider.\n\nFailures: retry provider errors with backoff. A permanently bad token: after 3 tries move it to a dead-letter queue and carry on, so it does not block the partition. Alert on consumer lag so you know if the 10 minutes will be missed.\n\nDuplicates: Redis SET processed:(notification_id, user_id) NX with a TTL of a day. A push is harmless work, so Redis is fine here (Part 2 caveat: it is not atomic with the send, so a rare duplicate or a rare miss is possible).");

designSlide(4, "The cart that works in three regions",
  "A shop has customers in Europe, Asia and the US. Each region must keep working if the link to the others is cut. Carts belong to one customer, who may use a phone and a laptop at the same time. Today there is one database, in Europe.",
  [["3", "regions", null], ["150 ms", "between continents", RUST], ["2,000/s", "cart edits", null], ["1", "region may be offline", RUST]],
  [["Customers", "Europe, Asia, US"], ["App servers", "in every region"], ["One database", "in Europe"]],
  ["What is wrong with one primary in Europe for Asian customers?", "If every region accepts writes, what conflict can one customer cause?", "Pick a conflict strategy for the cart (last write wins, merge, or home region) and say what it costs.", "Would you use the same design for checkout? When the link is cut, which do you pick for the cart, and which for checkout?"],
  "Problem with one primary: every Asian write travels about 150 ms each way before it is confirmed, and if Europe is unreachable, Asia cannot write at all.\n\nConflict: the customer adds a book on the phone in Asia and a pen on the laptop in Europe while the link is down. Each region accepted its write, and later each sees the other's cart version.\n\nStrategies: last write wins keeps one cart and silently drops the other item, which is unacceptable here. Merge fits a cart: a cart is a set of items, so combine both sets (add wins; removals need a marker so a removed item does not come back). Home region per customer is simpler (one owner, no conflicts), but a customer who travels pays the distance, and writes fail if the owner is unreachable.\n\nCheckout is different. Stock and payments need one owner for the contested value: two regions could both sell the last unit. So checkout leans to consistency: refuse or wait during a cut, with the order service owned by one region (or one primary per product).\n\nCAP: cart = availability (keep accepting, merge later). Checkout = consistency (refuse rather than oversell or double-charge). The same system can make a different choice for each feature.");

designSlide(5, "Money between shards",
  "A wallet app has 80 million users, split across 16 database shards by user ID. A transfer moves money from user A to user B, who is usually on a different shard. No money may be created or lost, even if a server crashes halfway.",
  [["80 M", "users", null], ["16", "shards (by user ID)", null], ["3,000/s", "transfers", null], ["94%", "of transfers cross shards", RUST]],
  [["App servers", "stateless"], ["Shard router", "user ID → shard"], ["Shards 1 to 16", "each: primary + replicas"]],
  ["Why can one transaction not cover A's shard and B's shard?", "Walk through debit A, credit B. What if the server crashes between the two?", "How does a transfer ID make retries safe?", "One merchant receives 10% of all transfers. What gets hot, and what do you do?"],
  "Why it is hard: a database transaction only covers one shard. A cross-shard commit protocol (two-phase commit) exists, but it is slow and one stuck shard blocks the other.\n\nSafe pattern: use a transfer ID and two local transactions. Step 1, on A's shard, in ONE transaction: debit A and insert a row in an outbox table: 'credit B, transfer T1'. Step 2: a worker reads the outbox and publishes to a queue keyed by B's user ID. Step 3, on B's shard, in ONE transaction: insert T1 into processed_transfers (T1 is the key) and credit B. If anything crashes, the message comes back (at-least-once) and the processed_transfers key makes the second credit a no-op. Money is never created (the credit happens once) and never lost (the outbox row survives the crash).\n\nIf B's account is closed, send a compensating credit back to A. Do not leave the money in limbo.\n\nHot merchant: it is the hot row and hot shard from Parts 1 and 3. Credit them in batches (one +N update per second), or split the merchant into several sub-accounts and sum them.\n\nHonest alternative: 3,000 transfers per second can fit on one strong primary with replicas. Do not shard until you must: sharding makes this exact problem appear. Shard when writes or storage outgrow one primary.");

// ============================================================ REVIEW + CLOSING
add((n) => {
  const s = newSlide(false);
  header(s, "3:40–3:55 · Review", "Group critique", false);
  const qs = [
    ["JR", TEAL, "Everyone", "What surprised you about your own numbers? Did you already know your service's real capacity, or did you have to guess?"],
    ["SR", SLATE, "Tech Leads", "Has anyone suggested \"just add more servers\" for a problem Amdahl's Law or the USL says that will not fix? How would you make that case in the room?"],
    ["QE", RUST, "QEs", "What load test would really check the plan someone just wrote, and does that test exist today?"],
    ["AI", PLUM, "AI Engineers", "If you asked an agent to \"make this service handle more load\", what is the first question it should ask before touching any code?"],
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

add((n) => {
  const s = newSlide(true);
  s.addText("BEFORE YOU GO", { x: 0.7, y: 0.55, w: 8, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Pick one. Do it this week.", { x: 0.7, y: 0.9, w: 11, h: 0.8, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 34, color: WHITE });
  [
    "Save your capacity plan where your team can see it, before the next traffic spike and not after",
    "Find the one piece of state that would block adding servers, and decide where it should live instead",
    "Find the hottest row or key in your database, and write down how long its lock is held",
    "Find one slow call inside a request that could be a queued message instead",
    "Run, or schedule, the load test that would really check today's numbers",
  ].forEach((it, i) => {
    const y = 1.95 + i * 0.78;
    card(s, 0.7, y, 11.9, 0.6, DARK2, { shadow: false });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: y + 0.15, w: 0.3, h: 0.3, rectRadius: 0.04, fill: { type: "none" }, line: { color: TEAL, width: 1.5 } });
    s.addText(it, { x: 1.45, y, w: 10.9, h: 0.6, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 14, color: WHITE });
  });
  s.addShape(pptx.ShapeType.line, { x: 0.7, y: 6.0, w: 11.9, h: 0, line: { color: "234750", width: 1 } });
  s.addText("NEXT — DAY 05", { x: 0.7, y: 6.15, w: 6, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  s.addText("Refactoring Technique & the Smell Catalog", { x: 0.7, y: 6.42, w: 10, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 22, color: WHITE });
  footer(s, n, true, "Clean Code to Scale");
});


// ============================================================ PRESENTER NOTES
const PRESENTER = [
`TIME: about 2 minutes.
SAY: Welcome. Today is about what happens when many people use your system at the same time. We follow one order system through a flash sale, one layer at a time: the web servers, the background jobs behind them, and the database underneath. Along the way there are short check-in exercises, a capacity-planning exercise, and three diagrams where you find the bottleneck.
ASK: Has anyone seen a system slow down or fall over under load? Take two or three quick stories and refer back to them later.`,
`TIME: about 2 minutes.
SAY: Five blocks. Part 1 is web servers, Part 2 is background jobs and queues, Part 3 is the database. Then a capacity plan and three diagram exercises in pairs, and a short group discussion. The ten check-ins take two to three minutes each and sit right after the idea they practise.
WATCH FOR: The full session is 3:55, so plan breaks (for example after Part 2 and before the hands-on). If you are short on time: run the check-ins as a quick poll first, then shorten the exercises, then make the capacity plan take-home.`,
`TIME: about 1 minute.
SAY: By the end you should be able to do four things: scale a web tier, know when more servers stop helping, put slow work behind a queue, and pick a way to scale a database. You will not become an expert in one session. The goal is to ask the right questions in design reviews.`,
`TIME: about 2 minutes.
SAY: Read the role that matches the room. Each role gets something different: juniors get a new lens, tech leads get a way to challenge "just add servers", QEs get numbers to test against, AI engineers get the question to ask an agent.
ASK: Which of these four lenses is closest to your job?`,
`TIME: 30 seconds.
SAY: Web servers are the easiest part of a system to scale out, if three things are true: they keep no user data, you chose on purpose where sessions live, and a thousand requests never wait in line for one database row. We take those three in order.`,
`TIME: about 3 minutes.
SAY: There are three ways to scale. A bigger machine is simplest, but it has a limit and is one point of failure. More machines is cheap and gives you backup for free, but only works if servers are stateless. Splitting into parts lets each part scale on its own, but turns function calls into network calls (that is Day 7).
ASK: Which of these has your team used most? Which did it regret?
NEXT: Everything hinges on "stateless". Let's see what that means.`,
`TIME: about 3 minutes.
SAY: Trace one request: client, load balancer, any of the three servers, then the cache or database. No server remembers anything. If server A dies, nothing is lost. If you add server D, nothing changes for users.
WATCH FOR: People think stateless means "no state anywhere". Correct that: the state lives in the cache or database, not in the web servers.`,
`TIME: about 2 minutes.
SAY: Walk the four rows: where the data lives, which server can answer, what happens when a server dies, and what happens on autoscale or deploy. The last row is the practical pain: stateful servers make deploys and scaling hurt.
NEXT: A quick check-in to see if you can spot state hiding in code.`,
`TIME: 2 minutes (pairs talk for about a minute, then reveal).
SAY: Read the code once. There are three servers now. For each marked line, what goes wrong?
FACILITATE: Take (a) to (d) one at a time. The point: some state breaks correctness (the rate limit, the local disk), some only breaks a number (the counter), and some is fine (a cache).`,
`TIME: about 3 minutes.
SAY: Login state and carts must live somewhere. Three common choices. Sticky sessions: the load balancer sends a user to the same server. Shared session store: servers look the session up in Redis. Token: the session travels with the request. Each is a tradeoff; there is no single best answer.
ASK: Which one does your system use today?`,
`TIME: about 3 minutes.
SAY: Sticky sessions are popular because they are easy: two lines of config. But look at the five problems: uneven load, lost sessions when a server dies, harder deploys, IP-based routing surprises, and the biggest one: it hides the fact that your app is stateful.
NEXT: Let's check which approach fits three different systems.`,
`TIME: 3 minutes.
FACILITATE: Give pairs one minute. Then take a vote on each case before revealing. The goal is that every choice comes with one reason, for example "tiny scale, so sticky is fine".`,
`TIME: about 5 minutes. Tell this like a story.
SAY: A normal day is 800 requests per second. The flash sale brings 6,000 and one server hits 100% CPU. The obvious fix is eight more servers. It breaks: customers see empty carts. Left code: the cart lives in this server's memory. Request 1 goes to server A, request 2 goes to server C, and C has never heard of this customer.
SAY: The fix is on the right: the cart lives in Redis. Same two functions; only where the state lives changed.
WATCH FOR: Someone will say "use sticky sessions". Good answer: it would hide the bug until a server is replaced mid-sale.`,
`TIME: about 4 minutes.
SAY: A different problem. Everyone wants the last items of one product. The code locks the row, then calls the payment API (about 200 ms) while still holding the lock. So only five buyers per second get through. Read the table: speed on a hot row is 1 divided by the time the lock is held. Nine servers or ninety does not change it.
ASK: What happens to the other 995 requests? (They wait, and waiting costs threads and connections.)`,
`TIME: about 3 minutes.
SAY: Walk the five steps. Waiting requests hold threads and database connections. The connection pool fills, so even unrelated pages fail. Timeouts cause retries, which add more waiting. Health checks fail, so the load balancer removes servers and the rest get worse. Finally every service that calls checkout gets stuck too. Day 7 studies this cascade in detail.
WATCH FOR: People blame the database. The real cause is holding a lock across a slow call.`,
`TIME: about 4 minutes.
SAY: Left: the fixed code. Reserve the stock with one atomic update, then call payment outside any lock, and give the unit back if payment fails. Right: six fixes in the order to try them. 1 hold the lock for less time. 2 check and change in one statement. 3 limit the waiting with a timeout. 4 optimistic locking. 5 split the hot value. 6 keep it off the database.
NEXT: Fix 4 deserves its own look, because people mix up pessimistic and optimistic locking.`,
`TIME: about 3 minutes.
SAY: Read the diagram top to bottom; time flows down. A locks the row. B asks for the same lock and waits (the amber bar). A updates and commits, the lock is released, and B gets the lock and sees the new stock. Safe, but B did nothing while waiting.
ASK: What happens to B if A takes 200 ms because of a slow payment call?`,
`TIME: about 3 minutes.
SAY: Same buyers, same row, no lock. Both read stock 5, version 7. A saves first with WHERE version = 7 and the row becomes version 8. B's save says WHERE version = 7, so 0 rows change. B finds out at save time, reads again and retries. Nobody waited, but B did its work twice.
ASK: With 1,000 buyers on one row, how many retries happen? (Many. Most of the work is wasted.)`,
`TIME: about 2 minutes.
SAY: Rule of thumb: if many people fight over the same row, make them take turns (a lock or the one-step update). If conflicts are rare and there is a long gap between read and save, like editing a form, check at the end.
NOTE: The one-step update from fix 2 sits between the two: its WHERE clause is a tiny optimistic check with no retry loop.`,
`TIME: 3 minutes.
FACILITATE: Give pairs a minute with a calculator. Question 3 is the key one: adding servers changes nothing for a hot row. Question 4: likes can be approximate, stock cannot.`,
`TIME: about 2 minutes.
SAY: Everything so far is really one question: if I add machines, how much more work gets done? The ratio is called speedup. Web servers are close to perfect because requests are independent. It breaks in two ways: some work cannot be split (like the hot row), and machines have to talk to each other. We put numbers on this in Part 2 using a nightly job.`,
`TIME: 30 seconds.
SAY: Not all work is a web request. Nightly jobs, emails and reports run in the background. First we see how much extra machines help a job, then how a queue hands the work out.`,
`TIME: about 4 minutes.
SAY: A nightly job recalculates every customer's loyalty tier, then writes one summary table. On one machine it takes 100 minutes: 95 minutes of per-customer work, which can be shared, and 5 minutes writing the summary, which cannot (one lock, one writer). On the right we split the first part across N machines. Total time is 95 divided by N, plus 5.
ASK: With 1,000 machines, how long does it take? (About 5.1 minutes. The 5 never changes.)`,
`TIME: about 3 minutes.
SAY: The same idea as a table. If 5% of the job cannot be split, 2 machines give 1.9 times, 8 give 5.9 times, and 32 give 12.6 times. At 32 machines you get only 39% of what you paid for. Each extra machine helps a little less. The formula is Amdahl's Law.
WATCH FOR: Do not derive the formula. The point is the shape.`,
`TIME: about 4 minutes.
SAY: Two curves. The teal one is Amdahl's Law: it climbs and flattens towards 20 times, which is 1 divided by the 5% serial part. The purple one is the Universal Scalability Law, which also counts the cost of machines talking to each other. It peaks near 24 machines and then goes down. Past that point, adding machines makes the system slower.
ASK: Has anyone seen a system get slower when more instances were added?
WATCH FOR: This is the most theoretical slide. Keep it about the idea (coordination costs grow), not the maths.`,
`TIME: 2 minutes.
FACILITATE: Pairs work out the three answers; the numbers are small enough for mental maths.`,
`TIME: about 4 minutes.
SAY: Back to the nightly job: who hands each chunk to a machine? Here the checkout does everything inside the request: save, email, loyalty, warehouse. The customer waits 850 ms and any failure fails the order. With a queue, the request saves the order, publishes one message, and returns in about 2 ms. Workers do the rest.
NOTE: Day 3's Observer pattern was this idea inside one program. A queue is Observer across programs.`,
`TIME: about 3 minutes.
SAY: Three benefits: it separates timing, absorbs spikes, and fails and scales on its own. But read the callout: a queue absorbs spikes, not constant overload. If 6,000 messages arrive per second and workers finish 800, the backlog grows by 5,200 every second. The queue only moves the waiting somewhere safer.`,
`TIME: 2 minutes.
FACILITATE: Pairs do the arithmetic. Common mistake: forgetting that the backlog keeps draining after the promo ends.`,
`TIME: about 3 minutes.
SAY: Step 1: add more workers on the same queue. Each message goes to one worker. That scales until the single queue becomes the limit, and until order matters. Step 2: split the queue into partitions. The sender hashes a key, so messages with the same key always go to the same partition, in order.
ASK: What would you use as the key for loyalty points? (customer_id: all events for one customer stay in order.)`,
`TIME: about 5 minutes.
SAY: A topic is a set of logs. Each message gets an offset, its position. The producer hashes the key to choose the partition. In a consumer group, each partition is read by exactly one consumer. The amber cells are messages waiting, which is the lag. A third consumer would get one partition; a fourth would sit idle. The number of partitions is the limit on parallel work.
WATCH FOR: Confusing "more consumers" with "more partitions". Consumers beyond the partition count do nothing.`,
`TIME: about 3 minutes.
SAY: Three ideas. A log, not a queue: messages are kept, and a consumer remembers its position, so you can replay. Order and parallelism: same key, same partition, in order, but only inside a partition. Surviving failures: each partition is copied, and a coordinator reassigns partitions when a consumer dies. That is enough Kafka for this session.`,
`TIME: about 3 minutes.
SAY: Before the problems, one idea we need: the offset. Every message in a partition gets a number: 0, 1, 2, and so on. The number never changes.
SAY: Each consumer group saves one number per partition. That is its bookmark: "I have finished everything up to here". Kafka stores it for you. After a crash or restart, the consumer continues from the bookmark.
SAY: Lag is how far behind the consumer is: newest offset minus saved offset. Here 9 minus 5 is 4 messages waiting.
SAY: When do we save the bookmark? Work first, then save: a crash means the message comes again, but nothing is lost. Save first, then work: a crash means the message is skipped and lost. We choose to allow repeats. That choice is where problem two, duplicates, comes from.
ASK: Which is worse for a loyalty-points job: giving points twice, or never giving them? What about for a payment?`,
`TIME: about 4 minutes.
SAY: Read the code top to bottom. The consumer does the work first, and only then saves its position (the commit). If it crashes in between, the message arrives again instead of being lost. So duplicates are normal, and the consumer checks whether it has already handled this event ID. Delivered at least once plus a safe consumer is how you get "exactly once".
ASK: What happens if we save the position before doing the work? (A crash would lose the message.)`,
`TIME: about 4 minutes.
SAY: Four problems. Order: only inside a partition, so choose the key carefully. Duplicates: make the consumer safe to repeat. Bad messages that always fail: retry a few times, then move them to a dead-letter queue. Backlog: add consumers, but never more than the number of partitions.
NOTE: Use the "In Kafka" column to ground each one.`,
`TIME: about 3 minutes.
SAY: Three more. Hot partition: one huge customer sends most of the traffic, so one partition and one consumer are overloaded while the rest idle. This is the Part 1 hot row again. Rebalance pauses happen when a consumer joins or leaves. And you cannot add partitions freely, because the key-to-partition mapping changes. Kafka is not the only option: SQS and RabbitMQ delete messages once acknowledged.`,
`TIME: about 2 minutes.
SAY: Problem one: order. A customer places an order, pays, then cancels. If those three events go to different partitions, different consumers process them at different speeds, and "cancelled" can be handled before "placed".
SAY: The fix is the key. Use the customer ID as the key, so all of that customer's events land in the same partition. Inside one partition, Kafka keeps the order.
ASK: Which events in your system must stay in order? Which ones do not care?`,
`TIME: about 2 minutes.
SAY: Problem two: duplicates. The consumer does the work, then crashes before it saves its position. After the restart, Kafka sends the same message again, and the customer gets 100 points instead of 50.
SAY: We cannot stop duplicates, because saving the position after the work is what guarantees we never lose a message. So we make the consumer safe to repeat: keep a table of handled event IDs and check it in the same transaction as the update.
NOTE: This is the same idea as the code on the earlier "Handling the same message twice" slide.`,
`TIME: about 3 minutes.
SAY: How do we skip a repeat? The producer gives every event an ID when it creates it, for example M7. The ID travels with the event, so it is the same on every delivery.
SAY: Option one, the database. In one transaction, insert the event ID into a table where the ID is the key, and do the update. First delivery: both succeed. Second delivery: the insert fails because M7 is already there, the transaction rolls back, and the points stay at 50. Because both steps are in one transaction, we can never have "marked as done" without "done".
SAY: Option two, Redis. SET with NX means "only set if it does not exist". If Redis says OK, this is the first time, so do the work. If it says nil, we have seen it, so skip. EX 86400 deletes the key after a day, so the table does not grow forever. The catch: Redis and the database are two systems. If we crash between the SET and the update, M7 is marked but never done. So use Redis for harmless work, and the database for points or money.
ASK: How long should we remember event IDs? (Longer than the longest time a message can be redelivered.)`,
`TIME: about 2 minutes.
SAY: Problem three: one bad message. It always fails, so the consumer retries it forever, and everything behind it in that partition waits.
SAY: The fix is a limit. Retry a few times, then move the message to a dead-letter queue and carry on. Someone looks at the dead-letter queue later. Kafka has no built-in dead-letter queue, so the consumer publishes the message to a topic such as orders.dlq.
ASK: Who would look at the dead-letter queue in your team, and how would they find out it has messages?`,
`TIME: about 2 minutes.
SAY: Problem four: backlog. Messages arrive at 1,200 per second, one consumer handles 300, so the queue grows by 900 every second. The word for this is consumer lag: the newest offset minus the offset the consumer has saved.
SAY: The fix is more consumers, up to the number of partitions. With four partitions, four consumers handle 1,200 per second. Also alert when lag grows, so you hear about it before customers do.
NOTE: A fifth consumer on four partitions would sit idle. One partition is read by one consumer at a time.`,
`TIME: about 2 minutes.
SAY: Problem five: a hot partition. We chose the merchant as the key, but one merchant sends most of the traffic. That partition and its one consumer are overloaded while the others sit idle. Adding consumers does not help.
SAY: Fix: choose a better key, or add a random part to the key for the big merchant so its events spread over several partitions. The cost: order across that merchant's events is lost, so check that nothing depends on it.
SAY: This is the hot row from Part 1, now in a queue.`,
`TIME: about 2 minutes.
SAY: Problem six: rebalance pauses. When a consumer joins or leaves, the group hands out the partitions again. In the old way, everyone stops while that happens. Deploys and autoscaling trigger this often.
SAY: The fix is to make rebalancing smaller. Cooperative rebalancing means only the partitions that move pause. Static membership means a quick restart does not count as leaving. And keep the work per message short, so a pause is short too.`,
`TIME: about 2 minutes.
SAY: Problem seven: adding partitions later. The partition for a key comes from hash % number of partitions. Change the number and keys move, so new events can overtake old ones still waiting in the old partition. It is the same trap as mod hashing from Part 3.
SAY: The fix is to plan ahead. Create more partitions than you need today, for example 12, and add consumers as you grow. With 12 partitions you can go from 3 to 6 consumers with no key moving.
ASK: How many partitions would you start with for this system, and why?`,
`TIME: 3 minutes.
FACILITATE: Key idea: record the handled event IDs in the same transaction as the update. The event ID must come from the sender, not the consumer.`,
`TIME: 30 seconds.
SAY: Servers and workers are copies of code, so they multiply easily. The database holds the one thing that cannot be copied freely: the data. Three steps: copy it, let every copy accept writes, or split it. CAP tells us what each step gives up.`,
`TIME: about 4 minutes.
SAY: One primary takes every write and sends changes to replicas. Reads can go to any of them. Great for read-heavy systems, and you get failover. It cannot scale writes. Walk the table: async is fast but can lose the last writes if the primary crashes; sync loses nothing but every write waits; semi-sync is a middle path. Then the callout: split brain is when two servers both think they are the primary.`,
`TIME: about 4 minutes.
SAY: A real problem. A customer changes their address (a write to the primary). A split second later the confirmation page reads it back from a replica that has not received it yet. The customer thinks the site undid their change. The fix: only reads that follow the user's own recent write go to the primary.
ASK: Why not send every read to the primary? (That removes the reason for having replicas.)`,
`TIME: 2 minutes.
FACILITATE: The test each time: would this reader notice, or be hurt by, old data?`,
`TIME: about 4 minutes.
SAY: If one writer is the bottleneck, let several servers accept writes, usually one per region. Writes get faster and a region outage does not stop writes. The price is in the timeline: Cairo and Dublin both change the same address within 80 ms, and both confirm. When replication catches up, both histories are valid but they disagree. Someone must decide.`,
`TIME: about 4 minutes.
SAY: Four ways to settle it. Last write wins is simple but silently throws away a write, and clocks disagree. Merge works for sets and counters. Home region per row avoids conflicts but makes remote writes slower. Avoiding duplicate keys only helps with IDs. The callout: stock = stock - 1 on two masters sells the last item twice, and no timestamp rule fixes it.
WATCH FOR: People hear "multi-master" and think it is a free upgrade. It is not.`,
`TIME: 3 minutes.
FACILITATE: Push for a different rule per field. The legal point (opt-out) usually gets a reaction.`,
`TIME: about 4 minutes.
SAY: Replicas copy all the data; multi-master copies the conflicts too. To scale writes and storage, give each server only part of the data and each row exactly one home. The obvious way to pick a shard is hash(key) % N. The table shows the catch: N is an input, so going from 4 to 5 shards moves every key in this sample. At real scale you would copy nearly all the data.`,
`TIME: about 5 minutes. Draw on the board if possible.
SAY: Picture a clock with thousands of positions. Hash each shard name onto the ring. Hash a key onto the same ring and walk clockwise to the first shard. Add shard E between C and D: only K4 moves, from D to E. K1, K2 and K3 stay. Adding a shard moves about 1 over N of the keys, not all of them.
WATCH FOR: Do not go into how to pick a hash function unless asked.`,
`TIME: about 3 minutes.
SAY: Four ways to decide where a row lives: range, hash mod N, consistent hashing, and a directory. None is strictly better. Ask: do my queries need ranges, how uneven are my keys, and how much complexity can the team handle?`,
`TIME: 3 minutes.
FACILITATE: Key idea: choose the key from the most common query.`,
`TIME: about 3 minutes.
SAY: These steps work together. A real database is partitioned, and each shard is a primary with replicas. The cost: queries and transactions across shards get harder, an uneven key can still make one shard hot, and changing the layout later is a real project.`,
`TIME: 1 minute.
SAY: New section. So far each step gave us something and cost us something. Now we ask why, from the beginning. Five pieces: what a distributed system is, what a network partition and split brain are, the CAP theorem, how to choose between the two answers it leaves us, and eventual consistency.
NOTE: Later slides on AP versus CP and on eventual consistency are the practical part. If time is short, keep those and shorten the first four.`,
`TIME: about 3 minutes.
SAY: A distributed system is many computers that work together by sending messages to each other, and look like one system to the user. Your shop already is one: a load balancer, web servers, a cache, a primary and a replica, Kafka.
SAY: We build them for three reasons. Scale: one machine is not enough. Staying up: if one machine dies the others carry on. Being close to users. The cost is on the right: messages get lost or arrive late, machines fail on their own, and no node sees the whole picture.
ASK: Count the machines your last request touched. Which of them could fail while the others keep running?`,
`TIME: about 3 minutes.
SAY: The key problem. A calls B and waits. Nothing comes back. Three things could have happened. The request was lost. B crashed before replying. Or B did the work and the reply was lost.
SAY: From where A sits, all three look the same: silence, then a timeout. A cannot tell them apart. If it retries, the work may happen twice. If it gives up, the work may already be done. This is why Part 2 insisted on safe retries, and it is the root of everything in this section.`,
`TIME: about 3 minutes.
SAY: A network partition is when the network splits the nodes into groups that cannot talk to each other. No machine has to crash. Group A is fine, group B is fine, they just cannot reach each other.
SAY: Causes are boring: a switch fails, a link between cloud zones is cut, a firewall change, a busy link that drops messages, a long pause that makes a node look dead. Large systems see these regularly.
SAY: The dangerous part: each group thinks the other group died, and carries on alone.
ASK: Has anyone seen an outage where everything was "up" but nothing worked?`,
`TIME: about 4 minutes.
SAY: Here is what that thinking does. A bank account has 100. A leader accepts writes and a replica copies them. The network is cut. Each side thinks the other is down. The replica hears nothing from the leader, so it decides the leader is dead and promotes itself. The old leader hears nothing from the replica, so it thinks the replica is dead and carries on.
SAY: Now two nodes both believe they are the leader. Client 1 withdraws 80 on one side. Client 2 withdraws 80 on the other. Both succeed. When the network heals, both copies say balance 20, but 160 was paid out of 100. This is split brain.
WATCH FOR: People who say "just detect it faster". Detection cannot tell a dead node from a cut link. That is the previous slide again.`,
`TIME: about 4 minutes.
SAY: The fix is a majority. Five nodes, split 3 and 2. A leader needs votes from more than half of all nodes: 3 of 5. The group of three can elect one. The group of two cannot, so it refuses writes. Two groups can never both have a majority, so there is only one leader.
SAY: Use an odd number of nodes, 3 or 5, so a split cannot be even. Fencing adds a number to each leader; storage rejects writes carrying an older number, so a stale leader cannot do harm.
SAY: Look at the cost. The minority side stopped answering. We gave up availability to keep the data correct. That is exactly the CAP choice, and the next slides name it.`,
`TIME: about 3 minutes.
SAY: Now the easy case. A follower is any node that is not the leader. N4 crashes. The leader stops getting replies from it, and that is all. There is no election, because the leader is fine.
SAY: Writes still work. A write needs a majority, 3 of 5. The leader plus N2, N3 and N5 is four, so the write is confirmed. The user notices nothing. When N4 comes back it starts as a follower, the leader sends it the writes it missed, or a snapshot if it was gone a long time, and it is a full member again.
SAY: The table shows how many nodes you can lose: one of three, two of five, three of seven. Lose the majority and writes stop. That is the CAP choice again.
ASK: With five nodes, how many follower failures can you survive? What about during a deploy, when you restart nodes one by one?`,
`TIME: about 4 minutes.
SAY: When the leader goes down, the others must choose a new one. This is the idea in Raft, used by etcd and Consul. Step 1: the leader sends a small heartbeat every 50 milliseconds or so. Step 2: the leader crashes and the heartbeats stop. Each follower waits a different random time before it decides the leader is gone.
SAY: Step 3: one timer runs out first, say N3. It adds one to the term number, votes for itself, and asks the others to vote for it. Each node votes once per term. Step 4: N3 gets two more votes. That is 3 out of 5, a majority. N3 is now the leader of term 2 and starts sending heartbeats.
SAY: Notice the quorum here. N3 needs more than half of all five nodes, not just the ones it can see. That is exactly what stops two leaders.
ASK: What happens if the network splits 3 and 2 right now? Who can elect a leader?`,
`TIME: about 4 minutes.
SAY: Five small rules make the election safe. A tie: wait a random time before asking, and if nobody wins, try again with a new term. Voting twice: one vote per term, so two leaders can never both reach a majority. A node that missed writes: vote only for a candidate whose data is at least as new as yours, so the new leader has every confirmed write.
SAY: An old leader that was only slow and comes back: every message carries the term number, and a leader that sees a higher term steps down. Last row: if only 2 of 5 can reach each other, nobody can get 3 votes, so there is no new leader. That is safe, but the system is unavailable. That is the CAP choice again.
NOTE: Use odd numbers of nodes, 3 or 5, so a split can never be even.`,
`TIME: about 3 minutes.
SAY: Now the three words. Consistency: every read sees the latest write, for example after a transfer both phones show the new balance. Availability: every request to a working node gets an answer, not an error. Partition tolerance: the system keeps working when messages between nodes are lost.
SAY: The theorem, proposed by Eric Brewer in 2000 and proved by Gilbert and Lynch in 2002: when a partition happens, a system cannot be both consistent and available.
WATCH FOR: Mixing this C up with the C in ACID. They are different words that happen to share a letter.`,
`TIME: about 4 minutes.
SAY: CAP is easier to see as messages. Client writes x = 5 to Node 1. Node 1 saves it and tells the client. Node 1 then sends a message to Node 2: "x is now 5". That message is lost, because the network between the nodes is cut.
SAY: Now the client reads x from Node 2. Node 2 has no way to know that a newer value exists. It has two honest choices. A: answer with x = 4. The user gets an answer, but it is old. B: say "error, try again later". Nothing wrong is shown, but this request fails. There is no third choice that is both always right and always answering.
ASK: In your own system, which one does the business want for each feature?`,
`TIME: about 4 minutes.
SAY: Why only two of three? Because P is not optional: networks fail whether we plan for it or not. A system that gives up P is a system that assumes a network that never fails, which means one machine. That is CA, and it is not distributed.
SAY: So the real choice is the two thick edges. CP: refuse when unsure, and give up availability. Bank transfer, seat 14C, etcd. AP: answer, fix later, and give up consistency. Like count, shopping cart, DNS, Cassandra.
ASK: Add a second machine to a single database. Which of the two did you just choose?`,
`TIME: about 5 minutes.
SAY: How to choose. Three questions, in order. One: is an old or wrong answer dangerous? Money, safety, double booking, a unique name. If yes, choose CP. Two: if the two sides disagree, can you fix it later by merging, last write wins, or a person reviewing? If no, CP. Three: does refusing cost more than a slightly old answer? Lost sales, blank pages. If yes, AP.
SAY: If the answer to all three is no, either works. Pick the simpler one.
NOTE: This is a business question as much as a technical one. Ask what the product owner would say.
ASK: Take one feature of your product and walk it through the three questions.`,
`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Give pairs two minutes to circle CP or AP for each of the nine features. Then take a show of hands per feature. Do not give the answers yet.
SAY: For each one ask two things: what does it cost if the answer is wrong, and what does it cost if we refuse?
WATCH FOR: The shopping cart and the catalog page. They split the room. Cart: a blocked cart loses a sale, a missing item can be merged. Catalog: a price 30 seconds old is fine, because the price is read again at checkout.
NOTE: The answers are on the next slide.`,
`TIME: about 4 minutes.
SAY: The same questions applied to features. Bank withdrawal, last seat, unique username, leader election: wrong answers are expensive and refusals are cheap, so CP. Cart, like counts, feed, catalog, chat: refusals are expensive and a slightly old answer is cheap, so AP.
SAY: One product uses both. The catalog page can be 30 seconds old, but the price is read again from the CP side at checkout. Do not ask "is our system AP or CP". Ask it per feature.`,
`TIME: about 3 minutes.
SAY: A simple trick to always read the latest value. The value is stored on three copies. Write to two of the three. Read from two of the three. Any two out of three must share at least one copy, here copy 2. So a read always finds the newest value.
SAY: Change the numbers and you change the trade. Write one, read one: fast and always available, but a read can miss the new value. Write all three, read one: reads are cheap, but if one copy is down, writes are refused. Cassandra and DynamoDB let you choose this per request.
NOTE: If the room finds this too deep, skip it. It is a bonus.`,
`TIME: about 3 minutes (2 in pairs, 1 to compare).
FACILITATE: Give pairs two minutes. For each of the four designs, they circle what happens when the network splits, and say whether it leans toward consistency or availability. Take a show of hands for each. Do not give the answers yet.
SAY: Use the earlier slides: a single primary, replicas, multi-master, shards. Ask what happens to the side that cannot reach the primary.
WATCH FOR: Design 2 is mixed. Reads lean toward availability, writes toward consistency. Accept that answer.
NOTE: The answers are on the next slide.`,
`TIME: about 3 minutes.
SAY: The same choice applied to the database designs from this part. Single primary with all reads from the primary: consistency. Primary plus async replicas serving reads: availability for reads, consistency for writes. Multi-master: availability, plus the conflicts from earlier. Partitioned shards: consistency per key and limited damage.
NOTE: CAP says nothing about the 99.9% of the time with no partition, when the trade is speed against consistency (PACELC). Day 7 returns to that.`,
`TIME: about 4 minutes.
SAY: A message queue has the same choice. Each Kafka partition has a leader and followers. With acks equals 1, the leader confirms right away. It is fast and available, but if the leader dies before copying, the message is lost.
SAY: With acks equals all and min in-sync replicas equals 2, a write needs two copies. If one follower is cut off, the leader and the other follower still make two, so writes go on. If both are cut off, the producer gets an error. That is the consistent choice: no loss, but sometimes a refusal.
SAY: Unclean leader election is the same trade again. On: a copy that is behind can become leader, so the topic stays up and some messages are lost. Off, the default: the partition waits.
NOTE: The words change, the trade is the same as the database slide.`,
`TIME: about 4 minutes.
SAY: Eventual consistency is what AP systems promise. We write x equals 5 on Node 1. Node 2 hears about it after 80 milliseconds, Node 3 after 200. A read at 50 ms on Node 2 gives the old value. A read at 300 ms gives 5 everywhere.
SAY: The promise: if no new writes arrive, all copies end up with the same value. It does not say when. Examples: a like count, you see 1,204 and a friend sees 1,206; DNS, a changed address takes minutes to reach everyone; a shopping cart that merges later.
WATCH FOR: "Eventually consistent means sometimes wrong." It means temporarily different, and it is the right choice when refusing costs more than being a little old.`,
`TIME: about 3 minutes.
SAY: What does it feel like to a user? Alice changes her name, then opens her profile. The app reads from a replica that has not caught up, so she sees the old name and thinks the save failed. Two seconds later it is right.
SAY: Two common fixes. Read your own writes: for a few seconds after a write, send that user's reads to the primary, or remember the version they wrote. And show the change right away in the app and confirm later. Other people may see the old name for a moment, and that is fine.
ASK: Where in your product could a user think a save failed?`,
`TIME: 2 minutes.
FACILITATE: The right answer depends on the cost of being wrong, which is a business question.`,
`TIME: about 4 minutes.
SAY: Now the same failure, but in a chain of services. A database becomes slow, not dead. The inventory service waits on it. The orders service waits on inventory. The web tier waits on orders. The slowness travels back toward the user, against the direction of the requests.
SAY: Four ways it spreads. Waiting: slow is worse than down, because down fails fast and slow holds threads. Retries: each retry is more load on a struggling service. Shared resources: one pool or database serves many callers. Shared fate: the same deploy, zone or config.
ASK: In your system, what is the slowest dependency, and who is waiting on it right now?`,
`TIME: about 6 minutes.
SAY: Read the timeline slowly. Ask first: at which minute did this stop being "a server crashed" and become "the shop is down"? The best answer is 12:03, when retries tripled the load.
SAY: Walk the rows. 12:00 is the trigger and it is small. 12:02 is slow, not down. 12:03 is the amplifier. 12:05 is the failure moving up a layer, because the health check removes a server that is only slow. 12:10 is the trap: restarting makes it worse because new servers meet the full backlog. 12:25 is the real fix: cut the load.
WATCH FOR: People who say "we should have restarted faster". Restarting was part of the problem.`,
`TIME: about 4 minutes.
SAY: Split the incident in three. The trigger was normal: one of four servers, 25 percent of capacity. We should survive that. The amplifiers are what we built: retries, waiting threads, health checks that remove slow servers, autoscaling into cold caches. The recovery trap is that the overload keeps itself going after the trigger is gone. Teams call this a metastable failure.
SAY: The retry math at the bottom is the easiest to remember. Three tries at each of three layers: one click becomes 27 calls to the database. Retry at one layer only.
ASK: Does anyone know how many retries happen between your user and your database?`,
`TIME: about 4 minutes.
SAY: One defence for each step. Timeouts stop the waiting, and a time budget for the whole request means each call knows how long it has left. A retry budget with backoff and jitter stops the storm. A circuit breaker stops calling something that is failing. Load shedding rejects extra requests early so the others finish. Bulkheads stop one caller from using every thread. Slow start protects new servers. A kill switch lets people cut load quickly.
ASK: If you could add only one before the next big sale, which one? Most teams say timeouts, because the others need them to work.`,
`TIME: 30 minutes, in pairs.
SAY: Work through the checkout example first: 800 req/s now, 2,400 expected, 750 req/s per server at 70% CPU, so 4 servers plus 1 spare. Then pairs fill the Yours column for a service they own, including the two newer rows: the hottest shared resource and the work that could leave the request path. Explain why 70% and not 95%: waiting time grows fast as a system gets busier.
FACILITATE: Walk the room. Some pairs will lack real numbers; accept rough guesses and steer them to the method.`,
`TIME: 10 minutes (about 6 in pairs, 4 to discuss).
SAY: Read the situation. Ask pairs to find the bottleneck using only the numbers on the diagram before they suggest fixes.
WATCH FOR: Teams jumping to "add more app servers". The app CPU is about 30% while the database is at 98% with all connections used.`,
`TIME: 10 minutes.
SAY: The team tripled the consumers and lag did not move. Ask for two reasons.
WATCH FOR: People forget that the partition count limits consumers, and that the key sends most traffic to one partition.`,
`TIME: 10 minutes.
SAY: Three replicas and writes still time out. Ask which node is the bottleneck, and why the "old address" tickets appear.
WATCH FOR: Suggesting multi-master. It brings back the conflict problem from Part 3.`,
`TIME: 10 minutes (about 6 in pairs, 4 to discuss).
SAY: A ticket drop is a flash sale with a deadline: 200,000 fans, 5,000 seats. Ask pairs to sketch before they talk about fixes.
WATCH FOR: Teams who lock the seat while payment runs. Ask how many seats that sells per minute. Also watch for people who forget the waiting room: the answer is not only "more servers".`,
`TIME: 10 minutes.
SAY: This is the hot row from Part 1 plus the queue from Part 2. Ask why one UPDATE per click cannot work at 50,000 per second.
WATCH FOR: "Just use Redis INCR". That is a fair answer for a count that may be a little off, so ask what happens when Redis restarts, and whether the hot key itself is still a limit.`,
`TIME: 10 minutes.
SAY: Start with the arithmetic: 10 million in 10 minutes is about 16,700 per second, close to the provider limit. Then ask pairs to design the queue.
WATCH FOR: One consumer looping over all users, and nobody mentioning the provider limit. Also check that someone covers bad messages (dead-letter queue) and duplicates.`,
`TIME: 10 minutes.
SAY: This brings back Part 3: latency, multi-master conflicts and CAP. Ask what happens when the same customer edits the cart in two regions during a cut.
WATCH FOR: "Last write wins" chosen without asking what it loses. Push on checkout: the same answer does not work there.`,
`TIME: 10 minutes.
SAY: Sharding makes simple things hard. Ask how to move money when A and B are on different shards, and what happens if the server crashes in between.
WATCH FOR: People who reach for one big transaction across shards. Lead them to a transfer ID, an outbox and an idempotent credit, and make sure someone asks whether you need to shard at all.`,
`TIME: 15 minutes.
SAY: Pick the question that matches who is in the room. Let people answer in their own words. Close by asking each person to commit to one action.`,
`TIME: 2 minutes.
SAY: Ask each person to choose one item and say it aloud. Next session is Day 5: Refactoring Technique and the Smell Catalog.`,
];

const ARABIC = [
`الوقت: حوالي دقيقتين.
قل: مرحبًا بكم. موضوع اليوم هو ما يحدث عندما يستخدم عدد كبير من الناس نظامك في نفس الوقت. سنتابع نظام طلبات واحدًا خلال حملة تخفيضات مفاجئة (flash sale)، طبقة بعد طبقة: خوادم الويب، ثم المهام الخلفية، ثم قاعدة البيانات تحتها. وخلال الجلسة هناك تمارين قصيرة (check-in)، وتمرين لتخطيط السعة، وثلاثة رسومات تبحثون فيها عن عنق الزجاجة (bottleneck).
اسأل: هل رأى أحدكم نظامًا يبطؤ أو ينهار تحت الضغط؟ خذ قصتين أو ثلاثًا بسرعة، وارجع إليها لاحقًا.`,
`الوقت: حوالي دقيقتين.
قل: خمس فقرات. الجزء الأول عن خوادم الويب، والثاني عن المهام الخلفية والطوابير، والثالث عن قاعدة البيانات. ثم خطة السعة وثلاثة تمارين رسومات بالأزواج، ثم نقاش جماعي قصير. التمارين العشرة القصيرة (check-ins) تستغرق دقيقتين إلى ثلاث لكل منها، وتأتي بعد الفكرة التي تتدرب عليها مباشرة.
انتبه: مدة الجلسة كاملة 3:55، لذا خطط لفترات راحة (مثلًا بعد الجزء الثاني وقبل التمارين العملية). إذا ضاق الوقت: نفّذ التمارين القصيرة كاستطلاع سريع أولًا، ثم اختصر التمارين الكبيرة، ثم اجعل خطة السعة واجبًا منزليًا.`,
`الوقت: حوالي دقيقة.
قل: في النهاية يجب أن تكونوا قادرين على أربعة أشياء: توسيع طبقة الويب، ومعرفة متى تتوقف الخوادم الإضافية عن الفائدة، ووضع العمل البطيء خلف طابور (queue)، واختيار طريقة لتوسيع قاعدة البيانات. لن تصبحوا خبراء في جلسة واحدة. الهدف أن تطرحوا الأسئلة الصحيحة في مراجعات التصميم.`,
`الوقت: حوالي دقيقتين.
قل: اقرأ الدور الذي يناسب الحاضرين. كل دور يستفيد بشكل مختلف: المبتدئون يحصلون على زاوية جديدة للنظر، والقادة التقنيون يحصلون على طريقة لمناقشة عبارة "فقط أضف خوادم"، ومهندسو الجودة يحصلون على أرقام للاختبار، ومهندسو الذكاء الاصطناعي يحصلون على السؤال الذي يطرحونه على الوكيل (agent).
اسأل: أي من هذه الزوايا الأربع أقرب إلى عملك؟`,
`الوقت: 30 ثانية.
قل: خوادم الويب هي أسهل جزء في النظام للتوسع الأفقي، بشرط ثلاثة أمور: ألا تحتفظ بأي بيانات عن المستخدمين، وأن تختاروا عمدًا أين تُخزَّن الجلسات، وألا ينتظر ألف طلب في طابور على صف واحد من قاعدة البيانات. سنأخذها بالترتيب.`,
`الوقت: حوالي 3 دقائق.
قل: هناك ثلاث طرق للتوسع. الجهاز الأكبر هو الأبسط لكن له حدًا وهو نقطة فشل واحدة. زيادة عدد الأجهزة رخيصة وتعطيك نسخًا احتياطية مجانًا، لكنها تعمل فقط إذا كانت الخوادم عديمة الحالة (stateless). وتقسيم النظام إلى أجزاء يسمح لكل جزء أن يتوسع بمفرده، لكنه يحوّل استدعاءات الدوال إلى استدعاءات شبكة (وهذا موضوع اليوم السابع).
اسأل: أيها استخدمه فريقكم أكثر؟ وأيها ندمتم عليه؟
التالي: كل شيء يعتمد على "عديم الحالة". لنرَ ما معناه.`,
`الوقت: حوالي 3 دقائق.
قل: تتبّع طلبًا واحدًا: العميل، ثم موزّع الحمل (load balancer)، ثم أي خادم من الثلاثة، ثم الذاكرة المؤقتة (cache) أو قاعدة البيانات. لا يتذكر أي خادم شيئًا. إذا تعطل الخادم A لا يضيع شيء. وإذا أضفت الخادم D لا يتغير شيء على المستخدمين.
انتبه: يظن البعض أن "عديم الحالة" يعني "لا توجد حالة في أي مكان". صحّح ذلك: الحالة موجودة في الـ cache أو قاعدة البيانات، وليست في خوادم الويب.`,
`الوقت: حوالي دقيقتين.
قل: مرّ على الصفوف الأربعة: أين تعيش بيانات المستخدم، وأي خادم يستطيع الإجابة، وماذا يحدث عند تعطل خادم، وماذا يحدث عند التوسع التلقائي أو النشر (deploy). الصف الأخير هو الألم العملي: الخوادم ذات الحالة تجعل النشر والتوسع مؤلمين.
التالي: تمرين سريع لنرى هل تستطيعون اكتشاف الحالة المخفية في الشيفرة.`,
`الوقت: دقيقتان (يتناقش الأزواج دقيقة تقريبًا ثم تكشف الإجابة).
قل: اقرأ الشيفرة مرة واحدة. لدينا الآن ثلاثة خوادم. لكل سطر معلّم: ما الذي يسوء؟
سهّل النقاش: خذ (a) إلى (d) واحدًا واحدًا. الفكرة: بعض الحالات تكسر الصحة (حد المعدل، القرص المحلي)، وبعضها يكسر رقمًا فقط (العداد)، وبعضها لا مشكلة فيه (الـ cache).`,
`الوقت: حوالي 3 دقائق.
قل: حالة تسجيل الدخول وسلات التسوق يجب أن تعيش في مكان ما. هناك ثلاثة خيارات شائعة. الجلسات اللاصقة (sticky): يرسل موزّع الحمل المستخدم إلى نفس الخادم دائمًا. مخزن جلسات مشترك: تبحث الخوادم عن الجلسة في Redis. الرمز (token): تسافر الجلسة مع الطلب. لكل خيار مقايضة، ولا يوجد خيار أفضل دائمًا.
اسأل: أيها يستخدمه نظامكم اليوم؟`,
`الوقت: حوالي 3 دقائق.
قل: الجلسات اللاصقة شائعة لأنها سهلة: سطران من الإعداد. لكن انظروا إلى المشاكل الخمس: حمل غير متوازن، وفقدان الجلسات عند تعطل الخادم، ونشر أصعب، ومفاجآت التوجيه حسب عنوان IP، والأهم: أنها تُخفي حقيقة أن تطبيقكم يحتفظ بالحالة.
التالي: لنرَ أي نهج يناسب ثلاثة أنظمة مختلفة.`,
`الوقت: 3 دقائق.
سهّل النقاش: أعطِ الأزواج دقيقة واحدة. ثم صوّتوا على كل حالة قبل كشف الإجابة. الهدف أن يأتي كل اختيار مع سبب واحد، مثل "الحجم صغير جدًا، فاللاصقة مناسبة".`,
`الوقت: حوالي 5 دقائق. احكها كقصة.
قل: اليوم العادي 800 طلب في الثانية. حملة التخفيضات تجلب 6,000 ويصل خادم واحد إلى 100% من المعالج. الحل الواضح هو ثمانية خوادم إضافية. لكنه يفشل: العملاء يرون سلالًا فارغة. الشيفرة اليسرى: السلة تعيش في ذاكرة هذا الخادم. الطلب الأول يذهب إلى الخادم A، والطلب الثاني إلى الخادم C، وC لم يسمع بهذا العميل.
قل: الحل على اليمين: السلة تعيش في Redis. الدالتان نفسهما؛ تغيّر فقط مكان الحالة.
انتبه: سيقول أحدهم "استخدموا الجلسات اللاصقة". إجابة جيدة: ستُخفي الخطأ إلى أن يُستبدل خادم في منتصف الحملة.`,
`الوقت: حوالي 4 دقائق.
قل: مشكلة مختلفة. الجميع يريدون آخر قطع منتج واحد. الشيفرة تقفل الصف ثم تستدعي واجهة الدفع (حوالي 200 مللي ثانية) وهي ما زالت تمسك القفل. لذلك يمر خمسة مشترين فقط في الثانية. اقرأ الجدول: السرعة على صف ساخن (hot row) = 1 مقسومًا على زمن بقاء القفل. تسعة خوادم أو تسعون لا تغيّر ذلك.
اسأل: ماذا يحدث للطلبات الـ 995 الأخرى؟ (تنتظر، والانتظار يكلّف خيوطًا واتصالات.)`,
`الوقت: حوالي 3 دقائق.
قل: مرّ على الخطوات الخمس. الطلبات المنتظرة تمسك خيوطًا واتصالات بقاعدة البيانات. يمتلئ مجمع الاتصالات (connection pool) فتفشل حتى الصفحات غير المرتبطة. المهلات (timeouts) تسبب إعادة محاولات تزيد الانتظار. تفشل فحوصات الصحة فيزيل موزّع الحمل خوادم، وتزداد حالة الباقي سوءًا. وأخيرًا تعلق كل خدمة تستدعي الدفع. اليوم السابع يدرس هذا التسلسل بالتفصيل.
انتبه: يلوم الناس قاعدة البيانات. السبب الحقيقي هو إمساك قفل أثناء استدعاء بطيء.`,
`الوقت: حوالي 4 دقائق.
قل: اليسار: الشيفرة المصلحة. احجز المخزون بتحديث ذري واحد، ثم استدعِ الدفع خارج أي قفل، وأعد الوحدة إذا فشل الدفع. اليمين: ستة حلول بترتيب التجربة. 1 أمسك القفل وقتًا أقل. 2 افحص وغيّر في جملة واحدة. 3 حدّ الانتظار بمهلة. 4 القفل المتفائل. 5 قسّم القيمة الساخنة. 6 أبعدها عن قاعدة البيانات.
التالي: الحل الرابع يستحق وقفة، لأن الناس تخلط بين القفل المتشائم والمتفائل.`,
`الوقت: حوالي 3 دقائق.
قل: اقرأ الرسم من الأعلى إلى الأسفل؛ الزمن يمضي للأسفل. A يقفل الصف. B يطلب القفل نفسه وينتظر (الشريط البرتقالي). A يحدّث ثم يثبّت (commit) فيُحرَّر القفل، وB يحصل على القفل ويرى المخزون الجديد. آمن، لكن B لم يفعل شيئًا أثناء الانتظار.
اسأل: ماذا يحدث لـ B إذا استغرق A مدة 200 مللي ثانية بسبب استدعاء دفع بطيء؟`,
`الوقت: حوالي 3 دقائق.
قل: المشتريان نفسهما والصف نفسه، بلا قفل. كلاهما يقرأ المخزون 5 والإصدار 7. A يحفظ أولًا بشرط WHERE version = 7 فيصبح الصف إصدار 8. حفظ B يقول WHERE version = 7 فلا يتغير أي صف. يكتشف B ذلك وقت الحفظ، ويقرأ من جديد ويعيد المحاولة. لم ينتظر أحد، لكن B كرر عمله مرتين.
اسأل: مع 1,000 مشترٍ على صف واحد، كم محاولة إعادة تحدث؟ (كثيرة. معظم العمل يضيع.)`,
`الوقت: حوالي دقيقتين.
قل: قاعدة عامة: إذا كان كثيرون يتنافسون على الصف نفسه، اجعلهم يتناوبون (قفل أو التحديث بخطوة واحدة). وإذا كانت التعارضات نادرة وهناك فجوة طويلة بين القراءة والحفظ، مثل تحرير نموذج، فتحقق عند النهاية.
ملاحظة: التحديث بخطوة واحدة من الحل 2 يقع بين الاثنين: شرط WHERE فيه فحص متفائل صغير بلا حلقة إعادة محاولة.`,
`الوقت: 3 دقائق.
سهّل النقاش: أعطِ الأزواج دقيقة مع آلة حاسبة. السؤال 3 هو الأهم: إضافة خوادم لا تغيّر شيئًا في الصف الساخن. السؤال 4: الإعجابات يمكن أن تكون تقريبية، أما المخزون فلا.`,
`الوقت: حوالي دقيقتين.
قل: كل ما سبق هو في الحقيقة سؤال واحد: إذا أضفت أجهزة، كم يزداد العمل المنجز؟ تُسمى النسبة التسريع (speedup). خوادم الويب قريبة من المثالي لأن الطلبات مستقلة. ويفشل ذلك بطريقتين: بعض العمل لا يمكن تقسيمه (مثل الصف الساخن)، والأجهزة تحتاج أن تتحدث مع بعضها. سنضع أرقامًا لذلك في الجزء الثاني باستخدام مهمة ليلية.`,
`الوقت: 30 ثانية.
قل: ليس كل عمل طلب ويب. المهام الليلية والبريد والتقارير تعمل في الخلفية. أولًا نرى كم تفيد الأجهزة الإضافية مهمة ما، ثم كيف يوزّع الطابور العمل.`,
`الوقت: حوالي 4 دقائق.
قل: مهمة ليلية تعيد حساب مستوى ولاء كل عميل ثم تكتب جدول ملخص واحدًا. على جهاز واحد تستغرق 100 دقيقة: 95 دقيقة عمل لكل عميل يمكن توزيعه، و5 دقائق لكتابة الملخص لا يمكن توزيعها (قفل واحد وكاتب واحد). على اليمين نقسّم الجزء الأول على N جهازًا. الزمن الكلي = 95 مقسومًا على N، زائد 5.
اسأل: مع 1,000 جهاز، كم يستغرق؟ (حوالي 5.1 دقيقة. الرقم 5 لا يتغير أبدًا.)`,
`الوقت: حوالي 3 دقائق.
قل: الفكرة نفسها في جدول. إذا كان 5% من المهمة لا يمكن تقسيمه، فجهازان يعطيان 1.9 مرة، و8 أجهزة تعطي 5.9 مرة، و32 جهازًا تعطي 12.6 مرة. عند 32 جهازًا لا تحصل إلا على 39% مما دفعته. كل جهاز إضافي يفيد أقل قليلًا. الصيغة اسمها قانون أمدال (Amdahl's Law).
انتبه: لا تشتق الصيغة. المهم الشكل العام.`,
`الوقت: حوالي 4 دقائق.
قل: منحنيان. الأخضر المزرق هو قانون أمدال: يرتفع ثم يستوي نحو 20 مرة، أي 1 مقسومًا على الجزء المتسلسل 5%. والبنفسجي هو قانون التوسع الشامل (USL)، وهو يحسب أيضًا تكلفة تحدّث الأجهزة مع بعضها. يبلغ ذروته قرب 24 جهازًا ثم ينخفض. بعد تلك النقطة، إضافة أجهزة تجعل النظام أبطأ.
اسأل: هل رأى أحدكم نظامًا يصبح أبطأ عند إضافة نسخ؟
انتبه: هذه أكثر شريحة نظرية. ركّز على الفكرة (تكلفة التنسيق تنمو) لا على الرياضيات.`,
`الوقت: دقيقتان.
سهّل النقاش: يحل الأزواج الأسئلة الثلاثة؛ الأرقام صغيرة بما يكفي للحساب الذهني.`,
`الوقت: حوالي 4 دقائق.
قل: نعود للمهمة الليلية: من يسلّم كل جزء إلى جهاز؟ هنا خدمة الدفع تفعل كل شيء داخل الطلب: حفظ، بريد، ولاء، مستودع. العميل ينتظر 850 مللي ثانية وأي فشل يُفشل الطلب. مع الطابور، يحفظ الطلب وينشر رسالة واحدة ويعود في حوالي 2 مللي ثانية. والعمال (workers) يقومون بالباقي.
ملاحظة: نمط Observer في اليوم الثالث كان هذه الفكرة داخل برنامج واحد. الطابور هو Observer عبر البرامج.`,
`الوقت: حوالي 3 دقائق.
قل: ثلاث فوائد: يفصل التوقيت، ويمتص الذروات، ويفشل ويتوسع بمفرده. لكن اقرأ التنبيه: الطابور يمتص الذروات لا الحمل الزائد الدائم. إذا وصلت 6,000 رسالة في الثانية وأنجز العمال 800، فيزيد التراكم 5,200 كل ثانية. الطابور فقط ينقل الانتظار إلى مكان أكثر أمانًا.`,
`الوقت: دقيقتان.
سهّل النقاش: يقوم الأزواج بالحساب. خطأ شائع: نسيان أن التراكم يستمر في الانخفاض بعد انتهاء الحملة.`,
`الوقت: حوالي 3 دقائق.
قل: الخطوة 1: أضف عمالًا على الطابور نفسه. كل رسالة تذهب إلى عامل واحد. يكفي ذلك حتى يصبح الطابور الواحد هو الحد، أو حتى يصبح الترتيب مهمًا. الخطوة 2: قسّم الطابور إلى أقسام (partitions). يحسب المرسل hash للمفتاح، فتذهب الرسائل ذات المفتاح نفسه إلى القسم نفسه بالترتيب.
اسأل: ماذا تستخدمون مفتاحًا لنقاط الولاء؟ (customer_id: تبقى كل أحداث العميل الواحد بالترتيب.)`,
`الوقت: حوالي 5 دقائق.
قل: الـ topic مجموعة من السجلات (logs). كل رسالة لها offset أي موضعها. ينتج المنتج (producer) hash للمفتاح ليختار القسم. في مجموعة المستهلكين (consumer group) يقرأ كل قسم مستهلك واحد فقط. الخانات البرتقالية رسائل تنتظر، وهذا هو التأخر (lag). المستهلك الثالث سيحصل على قسم، والرابع سيبقى بلا عمل. عدد الأقسام هو حد العمل المتوازي.
انتبه: الخلط بين "مزيد من المستهلكين" و"مزيد من الأقسام". المستهلكون الزائدون عن عدد الأقسام لا يفعلون شيئًا.`,
`الوقت: حوالي 3 دقائق.
قل: ثلاث أفكار. سجل لا طابور: الرسائل تُحفظ، والمستهلك يتذكر موضعه فقط، فيمكنك إعادة التشغيل (replay). الترتيب والتوازي: المفتاح نفسه، القسم نفسه، بالترتيب، لكن داخل القسم فقط. النجاة من الأعطال: كل قسم له نسخ، والمنسّق يعيد توزيع الأقسام عند موت مستهلك. هذا يكفي من Kafka لهذه الجلسة.`,
`الوقت: حوالي 3 دقائق.
قل: قبل المشاكل نحتاج فكرة واحدة: الـ offset. كل رسالة في القسم لها رقم: 0، 1، 2 وهكذا. والرقم لا يتغير أبدًا.
قل: كل مجموعة مستهلكين تحفظ رقمًا واحدًا لكل قسم. هذا هو الفاصل (bookmark) الخاص بها: "أنهيت كل شيء حتى هنا". يخزّنه Kafka لك. وبعد انهيار أو إعادة تشغيل يكمل المستهلك من الفاصل.
قل: التأخر (lag) هو مقدار تأخر المستهلك: أحدث offset ناقص الـ offset المحفوظ. هنا 9 ناقص 5 يساوي 4 رسائل تنتظر.
قل: متى نحفظ الفاصل؟ العمل أولًا ثم الحفظ: الانهيار يعني أن الرسالة تأتي مرة أخرى، لكن لا شيء يضيع. الحفظ أولًا ثم العمل: الانهيار يعني أن الرسالة تُتخطى وتضيع. نحن نقبل التكرار. ومن هذا الاختيار تأتي المشكلة الثانية: التكرار.
اسأل: أيهما أسوأ في مهمة نقاط الولاء: إعطاء النقاط مرتين أم عدم إعطائها أبدًا؟ وماذا عن الدفع؟`,
`الوقت: حوالي 4 دقائق.
قل: اقرأ الشيفرة من الأعلى. يؤدي المستهلك العمل أولًا ثم يحفظ موضعه (commit). إذا تعطل بينهما تصل الرسالة مرة أخرى بدل أن تضيع. إذًا التكرار طبيعي، ويتحقق المستهلك مما إذا عالج معرّف الحدث من قبل. "مرة واحدة على الأقل" مع مستهلك آمن هو الطريق إلى "مرة واحدة بالضبط".
اسأل: ماذا لو حفظنا الموضع قبل أداء العمل؟ (سيُفقد الحدث عند أي عطل.)`,
`الوقت: حوالي 4 دقائق.
قل: أربع مشاكل. الترتيب: داخل القسم فقط، فاختر المفتاح بعناية. التكرار: اجعل المستهلك آمنًا عند التكرار. الرسائل السيئة التي تفشل دائمًا: أعد المحاولة عدة مرات ثم انقلها إلى طابور الرسائل الميتة (dead-letter queue). التراكم: أضف مستهلكين، لكن لا تزد عن عدد الأقسام.
ملاحظة: استخدم عمود "في Kafka" لتثبيت كل نقطة.`,
`الوقت: حوالي 3 دقائق.
قل: ثلاث أخرى. القسم الساخن: عميل ضخم يرسل معظم الحركة فيمتلئ قسم واحد ومستهلك واحد بينما يبقى الباقي خاملًا. هذا هو الصف الساخن من الجزء الأول مجددًا. توقفات إعادة التوزيع (rebalance) عند دخول مستهلك أو خروجه. ولا يمكنك إضافة أقسام بحرية لأن ربط المفتاح بالقسم يتغير. Kafka ليس الخيار الوحيد: SQS وRabbitMQ يحذفان الرسالة بعد تأكيد استلامها.`,
`الوقت: حوالي دقيقتين.
قل: المشكلة الأولى: الترتيب. يضع العميل طلبًا، ثم يدفع، ثم يلغي. إذا ذهبت هذه الأحداث الثلاثة إلى أقسام مختلفة فستعالجها مستهلكات مختلفة بسرعات مختلفة، وقد يُعالَج "الإلغاء" قبل "الطلب".
قل: الحل هو المفتاح. استخدم رقم العميل مفتاحًا، فتذهب كل أحداثه إلى القسم نفسه. وداخل القسم الواحد يحافظ Kafka على الترتيب.
اسأل: أي الأحداث في نظامكم يجب أن تبقى مرتبة؟ وأيها لا يهتم بالترتيب؟`,
`الوقت: حوالي دقيقتين.
قل: المشكلة الثانية: التكرار. ينفذ المستهلك العمل ثم ينهار قبل حفظ موضعه. بعد إعادة التشغيل يرسل Kafka الرسالة نفسها مرة أخرى، فيحصل العميل على 100 نقطة بدل 50.
قل: لا نستطيع منع التكرار، لأن حفظ الموضع بعد العمل هو ما يضمن ألا نفقد أي رسالة. لذلك نجعل المستهلك آمنًا عند التكرار: نحتفظ بجدول لمعرّفات الأحداث المعالجة ونفحصه في المعاملة نفسها التي نحدّث فيها.
ملاحظة: هذه هي الفكرة نفسها في الشريحة السابقة "معالجة الرسالة نفسها مرتين".`,
`الوقت: حوالي 3 دقائق.
قل: كيف نتخطى الرسالة المكررة؟ يعطي المنتج كل حدث رقم تعريف (ID) عند إنشائه، مثل M7. الرقم يسافر مع الحدث، فهو نفسه في كل مرة تُسلَّم فيها الرسالة.
قل: الخيار الأول: قاعدة البيانات. في معاملة واحدة ندخل رقم الحدث في جدول يكون الرقم فيه هو المفتاح، ثم ننفذ التحديث. التسليم الأول: ينجح الاثنان. التسليم الثاني: يفشل الإدخال لأن M7 موجود، فتتراجع المعاملة وتبقى النقاط 50. ولأن الخطوتين في معاملة واحدة، لا يمكن أن يكون الحدث "معلَّمًا كمنتهٍ" دون أن يكون قد "انتهى" فعلًا.
قل: الخيار الثاني: Redis. الأمر SET مع NX يعني "ضع القيمة فقط إذا لم تكن موجودة". إذا رد Redis بـ OK فهذه أول مرة، فننفذ العمل. وإذا رد بـ nil فقد رأيناها، فنتخطى. وEX 86400 يحذف المفتاح بعد يوم، فلا يكبر الجدول إلى الأبد. المشكلة: Redis وقاعدة البيانات نظامان منفصلان. إذا انهرنا بين SET والتحديث يبقى M7 معلَّمًا ولم يُنفَّذ. لذلك نستخدم Redis للأعمال غير الحساسة، وقاعدة البيانات للنقاط أو المال.
اسأل: كم من الوقت يجب أن نتذكر أرقام الأحداث؟ (أطول من أطول مدة يمكن أن تُعاد فيها الرسالة.)`,
`الوقت: حوالي دقيقتين.
قل: المشكلة الثالثة: رسالة سيئة واحدة. تفشل دائمًا، فيعيد المستهلك محاولتها إلى ما لا نهاية، وكل ما خلفها في القسم ينتظر.
قل: الحل هو وضع حد. أعد المحاولة بضع مرات، ثم انقل الرسالة إلى طابور الرسائل الميتة (dead-letter queue) وتابع العمل. يراجعها شخص ما لاحقًا. لا يملك Kafka طابور رسائل ميتة جاهزًا، فينشر المستهلك الرسالة إلى موضوع مثل orders.dlq.
اسأل: من في فريقكم سيراجع طابور الرسائل الميتة، وكيف سيعرف أن فيه رسائل؟`,
`الوقت: حوالي دقيقتين.
قل: المشكلة الرابعة: التراكم. تصل الرسائل بمعدل 1200 في الثانية، والمستهلك الواحد يعالج 300، فيكبر الطابور 900 كل ثانية. اسم هذا هو تأخر المستهلك (consumer lag): أحدث موضع ناقص آخر موضع حفظه المستهلك.
قل: الحل هو مزيد من المستهلكين، حتى عدد الأقسام. مع أربعة أقسام، أربعة مستهلكين يعالجون 1200 في الثانية. ونضع تنبيهًا عندما يكبر التأخر لنعرف قبل العملاء.
ملاحظة: مستهلك خامس على أربعة أقسام سيبقى بلا عمل، فالقسم الواحد يقرأه مستهلك واحد في كل مرة.`,
`الوقت: حوالي دقيقتين.
قل: المشكلة الخامسة: القسم الساخن. اخترنا التاجر مفتاحًا، لكن تاجرًا واحدًا يرسل معظم الحركة. يمتلئ ذلك القسم ومستهلكه الوحيد بينما يبقى غيره بلا عمل. وإضافة مستهلكين لا تفيد.
قل: الحل: اختر مفتاحًا أفضل، أو أضف جزءًا عشوائيًا إلى مفتاح التاجر الكبير فتتوزع أحداثه على عدة أقسام. الثمن: يضيع الترتيب بين أحداث هذا التاجر، فتأكد أن لا شيء يعتمد عليه.
قل: هذا هو الصف الساخن من الجزء الأول، لكن هذه المرة داخل طابور.`,
`الوقت: حوالي دقيقتين.
قل: المشكلة السادسة: توقفات إعادة التوزيع (rebalance). عندما ينضم مستهلك أو يغادر، تعيد المجموعة توزيع الأقسام. في الطريقة القديمة يتوقف الجميع أثناء ذلك. والنشر والتوسع التلقائي يسببان هذا كثيرًا.
قل: الحل هو جعل إعادة التوزيع أصغر. إعادة التوزيع التعاونية تعني أن الأقسام التي تنتقل فقط هي التي تتوقف. والعضوية الثابتة تعني أن إعادة تشغيل سريعة لا تُحسب مغادرة. ونُبقي العمل لكل رسالة قصيرًا فيكون التوقف قصيرًا.`,
`الوقت: حوالي دقيقتين.
قل: المشكلة السابعة: إضافة أقسام لاحقًا. القسم الخاص بالمفتاح يُحسب من hash % عدد الأقسام. غيّر العدد فتنتقل المفاتيح، وقد تسبق الأحداث الجديدة القديمة التي لا تزال تنتظر في القسم القديم. إنه الفخ نفسه في التجزئة بالباقي من الجزء الثالث.
قل: الحل هو التخطيط المسبق. أنشئ أقسامًا أكثر مما تحتاجه اليوم، مثلًا 12، وأضف المستهلكين مع النمو. مع 12 قسمًا تنتقل من 3 إلى 6 مستهلكين دون أن ينتقل أي مفتاح.
اسأل: كم قسمًا ستبدأون به لهذا النظام، ولماذا؟`,
`الوقت: 3 دقائق.
سهّل النقاش: الفكرة الأساسية: سجّل معرّفات الأحداث المعالجة في المعاملة (transaction) نفسها مع التحديث. ويجب أن يأتي معرّف الحدث من المرسل لا من المستهلك.`,
`الوقت: 30 ثانية.
قل: الخوادم والعمال نسخ من الشيفرة فتتكاثر بسهولة. قاعدة البيانات تحمل الشيء الوحيد الذي لا يُنسخ بحرية: البيانات. ثلاث خطوات: انسخها، أو اجعل كل نسخة تقبل الكتابة، أو قسّمها. ونظرية CAP تخبرنا بما تتخلى عنه كل خطوة.`,
`الوقت: حوالي 4 دقائق.
قل: خادم أساسي (primary) واحد يقبل كل الكتابات ويرسل التغييرات إلى النسخ (replicas). القراءات يمكن أن تذهب لأي منها. ممتاز للأنظمة كثيرة القراءة، وتحصل على تحويل عند الفشل (failover). لكنه لا يوسّع الكتابة. مرّ على الجدول: غير المتزامن سريع لكنه قد يفقد آخر الكتابات إذا انهار الأساسي؛ والمتزامن لا يفقد شيئًا لكن كل كتابة تنتظر؛ وشبه المتزامن حل وسط. ثم التنبيه: انقسام الدماغ (split brain) هو أن يظن خادمان أن كلًا منهما هو الأساسي.`,
`الوقت: حوالي 4 دقائق.
قل: مشكلة حقيقية. عميل يغيّر عنوانه (كتابة إلى الأساسي). بعد لحظة تقرأ صفحة التأكيد العنوان من نسخة لم تستلمه بعد. يظن العميل أن الموقع تراجع عن تغييره. الحل: القراءات التي تلي كتابة المستخدم نفسه مباشرة فقط تذهب إلى الأساسي.
اسأل: لماذا لا نرسل كل القراءات إلى الأساسي؟ (لأن ذلك يلغي سبب وجود النسخ.)`,
`الوقت: دقيقتان.
سهّل النقاش: الاختبار في كل مرة: هل سيلاحظ القارئ البيانات القديمة أو يتضرر منها؟`,
`الوقت: حوالي 4 دقائق.
قل: إذا كان الكاتب الواحد هو عنق الزجاجة، فاجعل عدة خوادم تقبل الكتابة، عادة واحدًا لكل منطقة. الكتابة تصبح أسرع، وتوقف منطقة لا يوقف الكتابة. الثمن في الجدول الزمني: القاهرة ودبلن تغيّران العنوان نفسه خلال 80 مللي ثانية، وكلاهما يؤكد. عندما يلحق النسخ، يكون التاريخان صحيحين لكنهما مختلفان. يجب أن يقرر أحدهم.`,
`الوقت: حوالي 4 دقائق.
قل: أربع طرق للحسم. "آخر كتابة تفوز" بسيطة لكنها تُسقط كتابة بصمت، والساعات تختلف. الدمج (merge) يصلح للمجموعات والعدادات. المنطقة الأم لكل صف تتجنب التعارض لكن تجعل الكتابة البعيدة أبطأ. تجنب تكرار المفاتيح يفيد المعرّفات فقط. والتنبيه: stock = stock - 1 على خادمين رئيسيين يبيع آخر قطعة مرتين، ولا تصلحه أي قاعدة للتوقيت.
انتبه: يسمع الناس "multi-master" ويظنونه ترقية مجانية. ليس كذلك.`,
`الوقت: 3 دقائق.
سهّل النقاش: ادفعهم لاختيار قاعدة مختلفة لكل حقل. نقطة الاعتراض على الإعلانات (opt-out) القانونية تثير تفاعلًا عادة.`,
`الوقت: حوالي 4 دقائق.
قل: النسخ تنسخ كل البيانات، والكتابة المتعددة تنسخ التعارضات أيضًا. لتوسيع الكتابة والتخزين، أعطِ كل خادم جزءًا من البيانات، ولكل صف بيتًا واحدًا. الطريقة الواضحة لاختيار الشريحة: hash(key) % N. الجدول يظهر المشكلة: N مُدخل في المعادلة، فالانتقال من 4 إلى 5 شرائح يحرك كل مفتاح في هذه العينة. وعلى نطاق حقيقي ستنسخ تقريبًا كل البيانات.`,
`الوقت: حوالي 5 دقائق. ارسم على السبورة إن أمكن.
قل: تخيّل ساعة بآلاف المواضع. اجعل hash لأسماء الشرائح يضعها على الحلقة. واجعل hash لمفتاح يضعه على الحلقة نفسها، ثم امشِ مع عقارب الساعة إلى أول شريحة. أضف الشريحة E بين C وD: يتحرك المفتاح K4 فقط، من D إلى E. وتبقى K1 وK2 وK3. إضافة شريحة تحرك نحو 1/N من المفاتيح لا كلها.
انتبه: لا تدخل في كيفية اختيار دالة الـ hash إلا إذا سُئلت.`,
`الوقت: حوالي 3 دقائق.
قل: أربع طرق لتحديد مكان الصف: النطاق (range)، وhash mod N، والتجزئة المتسقة (consistent hashing)، والدليل (directory). لا واحدة أفضل مطلقًا. اسأل: هل تحتاج استعلاماتي نطاقات؟ وما مدى عدم توازن مفاتيحي؟ وكم تعقيدًا يتحمل الفريق؟`,
`الوقت: 3 دقائق.
سهّل النقاش: الفكرة الأساسية: اختر المفتاح انطلاقًا من الاستعلام الأكثر شيوعًا.`,
`الوقت: حوالي 3 دقائق.
قل: هذه الخطوات تعمل معًا. قاعدة البيانات الحقيقية مقسمة، وكل شريحة هي أساسي مع نسخ. التكلفة: الاستعلامات والمعاملات عبر الشرائح تصبح أصعب، والمفتاح غير المتوازن قد يجعل شريحة ساخنة، وتغيير التوزيع لاحقًا مشروع حقيقي.`,
`الوقت: دقيقة واحدة.
قل: قسم جديد. حتى الآن كانت كل خطوة تعطينا شيئًا وتكلفنا شيئًا. الآن نسأل لماذا، من البداية. خمسة أجزاء: ما هو النظام الموزع، وما هو انقسام الشبكة (network partition) والدماغ المنقسم (split brain)، ونظرية CAP، وكيف نختار بين الإجابتين اللتين تتركهما لنا، والاتساق في النهاية.
ملاحظة: الشرائح اللاحقة عن AP مقابل CP وعن الاتساق في النهاية هي الجزء العملي. إذا ضاق الوقت فاحتفظ بها واختصر الأربع الأولى.`,
`الوقت: حوالي 3 دقائق.
قل: النظام الموزع هو عدة حواسيب تعمل معًا بإرسال رسائل بعضها إلى بعض، وتبدو للمستخدم نظامًا واحدًا. متجركم هو نظام كذلك: موزّع أحمال، وخوادم ويب، وذاكرة مؤقتة، وقاعدة بيانات رئيسية ونسخة، وKafka.
قل: نبنيها لثلاثة أسباب. التوسع: جهاز واحد لا يكفي. الاستمرار: إذا مات جهاز يكمل الباقون. والقرب من المستخدمين. والثمن على اليمين: الرسائل تضيع أو تتأخر، والأجهزة تفشل كل على حدة، ولا تعرف أي عقدة الصورة كاملة.
اسأل: عدّوا الأجهزة التي مر بها آخر طلب لكم. أي منها يمكن أن يفشل بينما يكمل الباقون؟`,
`الوقت: حوالي 3 دقائق.
قل: هذه هي المشكلة المفتاحية. A يستدعي B وينتظر. لا يأتي شيء. ثلاثة احتمالات. الطلب ضاع. أو B انهار قبل أن يرد. أو B أنجز العمل وضاع الرد.
قل: من مكان A تبدو الثلاثة متطابقة: صمت ثم انتهاء مهلة. لا يستطيع A التمييز بينها. إذا أعاد المحاولة فقد يُنفَّذ العمل مرتين. وإذا استسلم فقد يكون العمل قد تم. لهذا أصرّ الجزء الثاني على إعادة المحاولة الآمنة، وهذا أصل كل شيء في هذا القسم.`,
`الوقت: حوالي 3 دقائق.
قل: انقسام الشبكة (network partition) هو أن تقسم الشبكة العقد إلى مجموعات لا تستطيع الحديث بعضها مع بعض. لا يلزم أن ينهار أي جهاز. المجموعة A سليمة، والمجموعة B سليمة، لكنهما لا تصلان إلى بعضهما.
قل: الأسباب عادية: مفتاح شبكة يتعطل، رابط بين مناطق السحابة ينقطع، تغيير في جدار الحماية، رابط مزدحم يُسقط الرسائل، توقف طويل يجعل العقدة تبدو ميتة. الأنظمة الكبيرة تراها بانتظام.
قل: الجزء الخطير: كل مجموعة تظن أن الأخرى ماتت، فتكمل وحدها.
اسأل: هل رأى أحدكم عطلًا كان فيه كل شيء "يعمل" لكن لا شيء ينجح؟`,
`الوقت: حوالي 4 دقائق.
قل: هذا ما يفعله ذلك التفكير. حساب بنكي فيه 100. قائد يقبل الكتابة ونسخة تنسخ عنه. تنقطع الشبكة. كل جهة تظن أن الأخرى متوقفة. النسخة لا تسمع شيئًا من القائد فتقرر أنه مات وترقّي نفسها قائدًا. والقائد القديم لا يسمع شيئًا من النسخة فيظن أنها ماتت ويكمل عمله.
قل: الآن عقدتان تعتقد كل منهما أنها القائد. العميل 1 يسحب 80 في جهة، والعميل 2 يسحب 80 في الجهة الأخرى. ينجح الاثنان. عندما تعود الشبكة تقول النسختان إن الرصيد 20، لكن دُفع 160 من أصل 100. هذا هو الدماغ المنقسم (split brain).
انتبه: من يقول "نكتشف أسرع". الاكتشاف لا يميز بين عقدة ميتة ورابط مقطوع. وهذه هي الشريحة السابقة نفسها.`,
`الوقت: حوالي 4 دقائق.
قل: الحل هو الأغلبية. خمس عقد انقسمت 3 و2. يحتاج القائد إلى أصوات أكثر من نصف كل العقد: 3 من 5. مجموعة الثلاثة تستطيع انتخاب قائد. ومجموعة الاثنين لا تستطيع، فترفض الكتابة. لا يمكن لمجموعتين أن تملكا أغلبية معًا، فلا يوجد إلا قائد واحد.
قل: استخدم عددًا فرديًا من العقد، 3 أو 5، حتى لا يكون الانقسام متساويًا. والحجز (fencing) يعطي كل قائد رقمًا؛ يرفض التخزين الكتابات ذات الرقم الأقدم، فلا يستطيع قائد قديم أن يفسد شيئًا.
قل: انظروا إلى الثمن. الجهة الأقلية توقفت عن الإجابة. تخلّينا عن التوافر (availability) لنحافظ على صحة البيانات. هذا هو اختيار CAP نفسه، والشرائح التالية تسميه.`,
`الوقت: حوالي 3 دقائق.
قل: الآن الحالة السهلة. المتابع (follower) هو أي عقدة ليست القائد. تنهار N4. يتوقف القائد عن تلقي ردود منها، وهذا كل شيء. لا يوجد انتخاب، لأن القائد سليم.
قل: الكتابة تستمر. تحتاج الكتابة إلى أغلبية، 3 من 5. القائد مع N2 وN3 وN5 أربعة، فتُؤكَّد الكتابة. لا يلاحظ المستخدم شيئًا. وعندما تعود N4 تبدأ متابعًا، فيرسل لها القائد الكتابات التي فاتتها، أو لقطة (snapshot) إذا غابت طويلًا، وتصبح عضوًا كاملًا من جديد.
قل: الجدول يبيّن كم عقدة يمكنكم خسارتها: واحدة من ثلاث، واثنتان من خمس، وثلاث من سبع. إذا خسرتم الأغلبية تتوقف الكتابة. وهذا اختيار CAP مرة أخرى.
اسأل: مع خمس عقد، كم فشل متابع تتحملون؟ وماذا عن وقت النشر حين نعيد تشغيل العقد واحدة بعد أخرى؟`,
`الوقت: حوالي 4 دقائق.
قل: عندما يسقط القائد، على الباقين اختيار قائد جديد. هذه فكرة Raft المستخدمة في etcd وConsul. الخطوة 1: القائد يرسل نبضة صغيرة كل 50 ميلي ثانية تقريبًا. الخطوة 2: ينهار القائد وتتوقف النبضات. كل متابع ينتظر وقتًا عشوائيًا مختلفًا قبل أن يقرر أن القائد رحل.
قل: الخطوة 3: ينتهي مؤقت أحدهم أولًا، لنقل N3. يزيد رقم الدورة (term) واحدًا، ويصوّت لنفسه، ويطلب من الآخرين التصويت له. كل عقدة تصوّت مرة واحدة في كل دورة. الخطوة 4: يحصل N3 على صوتين آخرين. هذه 3 من 5، أي أغلبية. أصبح N3 قائد الدورة 2 ويبدأ إرسال النبضات.
قل: لاحظوا النصاب (quorum) هنا. يحتاج N3 إلى أكثر من نصف العقد الخمس كلها، لا إلى من يراهم فقط. وهذا بالضبط ما يمنع وجود قائدين.
اسأل: ماذا يحدث لو انقسمت الشبكة الآن 3 و2؟ من يستطيع انتخاب قائد؟`,
`الوقت: حوالي 4 دقائق.
قل: خمس قواعد صغيرة تجعل الانتخاب آمنًا. التعادل: انتظر وقتًا عشوائيًا قبل الطلب، وإذا لم يفز أحد حاول مجددًا بدورة جديدة. التصويت مرتين: صوت واحد لكل دورة، فلا يمكن لقائدين أن يصلا معًا إلى أغلبية. عقدة فاتتها كتابات: لا تصوّت إلا لمرشح بياناته على الأقل بحداثة بياناتك، فيملك القائد الجديد كل ما أُكّد من كتابات.
قل: قائد قديم كان بطيئًا فقط ثم عاد: كل رسالة تحمل رقم الدورة، والقائد الذي يرى دورة أعلى يتنحى. الصف الأخير: إذا كانت عقدتان فقط من خمس تصلان إلى بعضهما فلا أحد يستطيع جمع 3 أصوات، فلا يوجد قائد جديد. هذا آمن لكن النظام غير متاح. هذا اختيار CAP نفسه مرة أخرى.
ملاحظة: استخدم عددًا فرديًا من العقد، 3 أو 5، حتى لا يكون الانقسام متساويًا.`,
`الوقت: حوالي 3 دقائق.
قل: الآن الكلمات الثلاث. الاتساق (Consistency): كل قراءة ترى آخر كتابة، مثلًا بعد التحويل يظهر الرصيد الجديد على الهاتفين. التوافر (Availability): كل طلب إلى عقدة تعمل يحصل على إجابة وليس خطأ. تحمّل الانقسام (Partition tolerance): يستمر النظام في العمل عندما تضيع الرسائل بين العقد.
قل: النظرية اقترحها إريك بروير عام 2000 وأثبتها جيلبرت وليتش عام 2002: عندما يحدث انقسام لا يستطيع النظام أن يكون متسقًا ومتوافرًا معًا.
انتبه: خلط هذا الاتساق مع الحرف C في ACID. هما كلمتان مختلفتان اشتركتا في حرف.`,
`الوقت: حوالي 4 دقائق.
قل: يسهل فهم CAP إذا رأيناه كرسائل. يكتب العميل x = 5 على العقدة 1. تحفظه العقدة 1 وتخبر العميل. ثم ترسل رسالة إلى العقدة 2: "x أصبحت 5". تضيع هذه الرسالة لأن الشبكة بين العقدتين مقطوعة.
قل: الآن يقرأ العميل x من العقدة 2. لا تستطيع العقدة 2 أن تعرف أن هناك قيمة أحدث. أمامها خياران صادقان. أ: تجيب x = 4. يحصل المستخدم على إجابة لكنها قديمة. ب: تقول "خطأ، حاول لاحقًا". لا يظهر شيء خاطئ، لكن هذا الطلب يفشل. لا يوجد خيار ثالث يكون صحيحًا دائمًا ويجيب دائمًا.
اسأل: في نظامكم، أي الخيارين تريده الجهة المسؤولة عن العمل لكل ميزة؟`,
`الوقت: حوالي 4 دقائق.
قل: لماذا اثنان فقط من ثلاثة؟ لأن P ليس اختياريًا: الشبكات تفشل سواء خططنا لذلك أم لا. والنظام الذي يتخلى عن P يفترض شبكة لا تفشل أبدًا، أي جهازًا واحدًا. هذا هو CA وهو ليس موزعًا.
قل: فالاختيار الحقيقي هو الضلعان السميكان. CP: ارفض عندما لا تكون متأكدًا، وتتخلى عن التوافر. تحويل بنكي، المقعد 14C، etcd. AP: أجب وأصلح لاحقًا، وتتخلى عن الاتساق. عدّاد الإعجابات، سلة التسوق، DNS، Cassandra.
اسأل: أضفتم جهازًا ثانيًا إلى قاعدة بيانات مفردة. أيّ الاثنين اخترتم للتو؟`,
`الوقت: حوالي 5 دقائق.
قل: كيف نختار. ثلاثة أسئلة بالترتيب. الأول: هل الإجابة القديمة أو الخاطئة خطيرة؟ مال، سلامة، حجز مزدوج، اسم فريد. إذا نعم فاختر CP. الثاني: إذا اختلف الطرفان، هل يمكن إصلاح ذلك لاحقًا بالدمج أو بأن تفوز آخر كتابة أو بمراجعة شخص؟ إذا لا فـ CP. الثالث: هل الرفض أغلى من إجابة قديمة قليلًا؟ مبيعات ضائعة، صفحات فارغة. إذا نعم فـ AP.
قل: إذا كانت الإجابة على الثلاثة لا، فكلاهما يصلح. اختر الأبسط.
ملاحظة: هذا سؤال عمل بقدر ما هو سؤال تقني. اسأل ماذا سيقول مالك المنتج.
اسأل: خذوا ميزة واحدة من منتجكم ومرّوها على الأسئلة الثلاثة.`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: أعطِ كل زوج دقيقتين ليضع دائرة حول CP أو AP لكل ميزة من الميزات التسع. ثم خذ رفع الأيدي لكل ميزة. لا تعطِ الإجابات بعد.
قل: لكل ميزة اسأل شيئين: كم تكلفنا الإجابة الخاطئة، وكم يكلفنا الرفض؟
انتبه: سلة التسوق وصفحة الكتالوج. تقسمان الحضور. السلة: السلة المحجوبة تخسر بيعًا، والعنصر الناقص يمكن دمجه. الكتالوج: سعر عمره 30 ثانية مقبول، لأن السعر يُقرأ مرة أخرى عند الدفع.
ملاحظة: الإجابات في الشريحة التالية.`,
`الوقت: حوالي 4 دقائق.
قل: الأسئلة نفسها مطبقة على ميزات. سحب بنكي، آخر مقعد، اسم مستخدم فريد، انتخاب قائد: الإجابة الخاطئة غالية والرفض رخيص، فـ CP. السلة، عدّاد الإعجابات، الخلاصة، الكتالوج، المحادثة: الرفض غالٍ والإجابة القديمة قليلًا رخيصة، فـ AP.
قل: المنتج الواحد يستخدم الاثنين. صفحة الكتالوج يمكن أن تكون قديمة 30 ثانية، لكن السعر يُقرأ من جانب CP عند الدفع. لا تسألوا "هل نظامنا AP أم CP". اسألوا لكل ميزة.`,
`الوقت: حوالي 3 دقائق.
قل: حيلة بسيطة لنقرأ دائمًا آخر قيمة. القيمة مخزنة على ثلاث نسخ. اكتب إلى اثنتين من الثلاث. واقرأ من اثنتين من الثلاث. أي اثنتين من ثلاث لا بد أن تشتركا في نسخة واحدة على الأقل، هنا النسخة 2. فتجد القراءة دائمًا أحدث قيمة.
قل: غيّر الأرقام تتغير المقايضة. اكتب إلى واحدة واقرأ من واحدة: سريع ومتاح دائمًا، لكن قد تفوّت القراءة القيمة الجديدة. اكتب إلى الثلاث واقرأ من واحدة: القراءة رخيصة، لكن إذا سقطت نسخة واحدة تُرفض الكتابة. Cassandra وDynamoDB تتيحان اختيار هذا لكل طلب.
ملاحظة: إذا وجد الحضور هذا عميقًا فتجاوزه. إنه إضافة.`,
`الوقت: حوالي 3 دقائق (دقيقتان في أزواج ودقيقة للمقارنة).
سهّل النقاش: أعطِ كل زوج دقيقتين. لكل تصميم من التصميمات الأربعة يضعون دائرة حول ما يحدث عند انقسام الشبكة، ويقولون هل يميل إلى الاتساق أم إلى التوافر. خذ رفع الأيدي لكل تصميم. لا تعطِ الإجابات بعد.
قل: استعينوا بالشرائح السابقة: القاعدة الرئيسية الواحدة، والنسخ، وتعدد الرؤوس، والقطع المجزأة. اسألوا ماذا يحدث للجهة التي لا تصل إلى القاعدة الرئيسية.
انتبه: التصميم الثاني مختلط. القراءة تميل إلى التوافر والكتابة إلى الاتساق. اقبل هذه الإجابة.
ملاحظة: الإجابات في الشريحة التالية.`,
`الوقت: حوالي 3 دقائق.
قل: الاختيار نفسه مطبقًا على تصميمات قواعد البيانات في هذا الجزء. قاعدة رئيسية واحدة وكل القراءات منها: اتساق. رئيسية مع نسخ غير متزامنة تخدم القراءة: توافر للقراءة واتساق للكتابة. تعدد الرؤوس (multi-master): توافر مع التعارضات التي رأيناها. القطع المجزأة: اتساق لكل مفتاح وضرر محدود.
ملاحظة: CAP لا تقول شيئًا عن 99.9% من الوقت الذي لا يوجد فيه انقسام، حيث تكون المقايضة بين السرعة والاتساق (PACELC). اليوم السابع يعود إلى هذا.`,
`الوقت: حوالي 4 دقائق.
قل: طابور الرسائل أمامه الاختيار نفسه. لكل قسم في Kafka قائد (leader) ومتابعون (followers). مع acks=1 يؤكد القائد فورًا. هذا سريع ومتاح، لكن إذا مات القائد قبل النسخ تضيع الرسالة.
قل: مع acks=all وmin.insync.replicas=2 تحتاج الكتابة إلى نسختين. إذا انقطع متابع واحد يبقى القائد والمتابع الآخر، فتستمر الكتابة. وإذا انقطع الاثنان يحصل المنتج على خطأ. هذا هو الخيار المتسق: لا خسارة، لكن أحيانًا رفض.
قل: انتخاب القائد غير النظيف (unclean leader election) هو الصفقة نفسها. إذا فُعّل: يمكن لنسخة متأخرة أن تصبح قائدًا، فيبقى الموضوع متاحًا وتضيع بعض الرسائل. وإذا عُطّل (الافتراضي): ينتظر القسم.
ملاحظة: الكلمات تتغير والصفقة هي نفسها في شريحة قاعدة البيانات.`,
`الوقت: حوالي 4 دقائق.
قل: الاتساق في النهاية (eventual consistency) هو ما تعد به أنظمة AP. نكتب x = 5 على العقدة 1. تعلم العقدة 2 بعد 80 ميلي ثانية، والعقدة 3 بعد 200. قراءة عند 50 ميلي ثانية من العقدة 2 تعطي القيمة القديمة. وقراءة عند 300 ميلي ثانية تعطي 5 في كل مكان.
قل: الوعد: إذا لم تصل كتابات جديدة، تنتهي كل النسخ بالقيمة نفسها. لكنه لا يقول متى. أمثلة: عدّاد إعجابات تراه 1,204 ويراه صديقك 1,206؛ DNS يحتاج تغيير العنوان دقائق ليصل إلى الجميع؛ وسلة تسوق تُدمج لاحقًا.
انتبه: "الاتساق في النهاية يعني أحيانًا خطأ". بل يعني اختلافًا مؤقتًا، وهو الخيار الصحيح عندما يكون الرفض أغلى من أن تكون القيمة قديمة قليلًا.`,
`الوقت: حوالي 3 دقائق.
قل: كيف يبدو الأمر للمستخدم؟ تغيّر أليس اسمها ثم تفتح ملفها. يقرأ التطبيق من نسخة لم تلحق بعد، فترى الاسم القديم وتظن أن الحفظ فشل. بعد ثانيتين يصبح صحيحًا.
قل: حلان شائعان. اقرأ ما كتبته أنت (read your own writes): لبضع ثوانٍ بعد الكتابة، وجّه قراءات هذا المستخدم إلى القاعدة الرئيسية، أو تذكّر النسخة التي كتبها. وأظهر التغيير في التطبيق فورًا وأكّده لاحقًا. قد يرى الآخرون الاسم القديم للحظة، وهذا مقبول.
اسأل: أين في منتجكم قد يظن المستخدم أن الحفظ فشل؟`,
`الوقت: دقيقتان.
سهّل النقاش: الإجابة الصحيحة تعتمد على كلفة الخطأ، وهذا سؤال أعمال.`,
`الوقت: حوالي 4 دقائق.
قل: الآن الفشل نفسه لكن في سلسلة خدمات. قاعدة البيانات تصبح بطيئة، لا ميتة. خدمة المخزون تنتظرها. وخدمة الطلبات تنتظر المخزون. وطبقة الويب تنتظر الطلبات. البطء يسير إلى الخلف نحو المستخدم، عكس اتجاه الطلبات.
قل: أربع طرق للانتشار. الانتظار: البطيء أسوأ من المعطّل، لأن المعطّل يفشل سريعًا والبطيء يحجز الخيوط. إعادة المحاولة: كل محاولة حمل إضافي على خدمة متعبة. الموارد المشتركة: مجمع اتصالات أو قاعدة بيانات واحدة لكثير من المستدعين. المصير المشترك: النشر نفسه أو المنطقة نفسها أو الإعداد نفسه.
اسأل: في نظامكم، ما أبطأ خدمة تعتمدون عليها، ومن ينتظرها الآن؟`,
`الوقت: حوالي 6 دقائق.
قل: اقرأ الجدول الزمني ببطء. اسأل أولًا: في أي دقيقة توقف الأمر عن كونه "خادم انهار" وأصبح "المتجر متوقف"؟ أفضل إجابة 12:03، عندما ضاعفت إعادة المحاولة الحمل ثلاث مرات.
قل: امشِ على الصفوف. الساعة 12:00 هي المُحفّز وهو صغير. 12:02 بطء لا توقف. 12:03 هو المضخِّم. 12:05 الفشل ينتقل طبقة إلى الأعلى لأن فحص الصحة يزيل خادمًا بطيئًا فقط. 12:10 هو الفخ: إعادة التشغيل تزيد الأمر سوءًا لأن الخوادم الجديدة تواجه كل الطلبات المتراكمة. 12:25 هو الحل الحقيقي: خفض الحمل.
انتبه: من يقول "كان علينا إعادة التشغيل أسرع". إعادة التشغيل كانت جزءًا من المشكلة.`,
`الوقت: حوالي 4 دقائق.
قل: قسّم الحادثة إلى ثلاثة. المُحفّز كان عاديًا: خادم من أربعة، أي 25% من السعة. يجب أن ننجو من هذا. المضخِّمات هي ما بنيناه نحن: إعادة المحاولة، والخيوط المنتظرة، وفحوص الصحة التي تزيل الخوادم البطيئة، والتوسع التلقائي إلى ذاكرة مؤقتة فارغة. وفخ التعافي هو أن الحمل الزائد يغذي نفسه بعد زوال المُحفّز. يسمي الفرق هذا الفشل شبه المستقر (metastable failure).
قل: حساب إعادة المحاولة في الأسفل هو الأسهل للتذكر. ثلاث محاولات في كل طبقة من ثلاث طبقات: نقرة واحدة تصبح 27 استدعاءً لقاعدة البيانات. أعد المحاولة في طبقة واحدة فقط.
اسأل: هل يعرف أحدكم كم مرة تُعاد المحاولة بين المستخدم وقاعدة البيانات عندكم؟`,
`الوقت: حوالي 4 دقائق.
قل: دفاع واحد لكل خطوة. المهلات (timeouts) توقف الانتظار، وميزانية زمنية للطلب كله تجعل كل استدعاء يعرف كم بقي له. ميزانية إعادة المحاولة مع التراجع (backoff) والعشوائية (jitter) توقف العاصفة. قاطع الدائرة (circuit breaker) يتوقف عن استدعاء ما يفشل. التخلص من الحمل (load shedding) يرفض الطلبات الزائدة مبكرًا لتنتهي بقية الطلبات. الحواجز (bulkheads) تمنع مستدعيًا واحدًا من استخدام كل الخيوط. البدء البطيء (slow start) يحمي الخوادم الجديدة. ومفتاح الإيقاف (kill switch) يتيح خفض الحمل بسرعة.
اسأل: لو أضفتم واحدًا فقط قبل التخفيضات الكبيرة القادمة، أيها؟ أغلب الفرق تقول المهلات، لأن البقية تحتاجها لتعمل.`,
`الوقت: 30 دقيقة، بالأزواج.
قل: اعمل أولًا على مثال الدفع: 800 طلب/ثانية الآن، و2,400 متوقعة، و750 طلب/ثانية لكل خادم عند 70% معالج، أي 4 خوادم زائد واحد احتياطي. ثم يملأ الأزواج عمود "Yours" لخدمة يملكونها، بما في ذلك الصفان الأحدث: أسخن مورد مشترك، والعمل الذي يمكن إخراجه من مسار الطلب. اشرح لماذا 70% لا 95%: زمن الانتظار ينمو بسرعة كلما ازداد انشغال النظام.
سهّل النقاش: تجوّل في القاعة. بعض الأزواج لن تملك أرقامًا حقيقية؛ اقبل التقديرات التقريبية ووجّههم نحو الطريقة.`,
`الوقت: 10 دقائق (حوالي 6 بالأزواج و4 للنقاش).
قل: اقرأ الموقف. اطلب من الأزواج إيجاد عنق الزجاجة باستخدام الأرقام الموجودة على الرسم فقط قبل اقتراح الحلول.
انتبه: الفرق التي تقفز إلى "أضف خوادم تطبيق". معالج التطبيق حوالي 30% بينما قاعدة البيانات عند 98% وكل اتصالاتها مستخدمة.`,
`الوقت: 10 دقائق.
قل: ضاعف الفريق المستهلكين ثلاث مرات ولم يتحرك التأخر. اطلب سببين.
انتبه: ينسى الناس أن عدد الأقسام يحدّ المستهلكين، وأن المفتاح يرسل معظم الحركة إلى قسم واحد.`,
`الوقت: 10 دقائق.
قل: ثلاث نسخ والكتابة ما زالت تنتهي مهلتها. اسأل أي عقدة هي عنق الزجاجة، ولماذا تظهر تذاكر "العنوان القديم".
انتبه: اقتراح multi-master. يعيد مشكلة التعارض من الجزء الثالث.`,
`الوقت: 10 دقائق (حوالي 6 في أزواج و4 للنقاش).
قل: بيع التذاكر مثل بيع التخفيضات لكن بموعد نهائي: 200,000 مشجع و5,000 مقعد. اطلب من كل زوج أن يرسم تصميمه قبل الحديث عن الحلول.
انتبه: فرق تقفل المقعد أثناء الدفع. اسألهم كم مقعدًا يُباع في الدقيقة بهذه الطريقة. وانتبه لمن ينسى غرفة الانتظار: الحل ليس "المزيد من الخوادم" فقط.`,
`الوقت: 10 دقائق.
قل: هذا هو الصف الساخن من الجزء الأول مع الطابور من الجزء الثاني. اسأل لماذا لا يصلح تحديث واحد لكل نقرة عند 50,000 في الثانية.
انتبه: "استخدم Redis INCR فقط". إجابة معقولة لعدّاد يقبل خطأً بسيطًا، فاسأل ماذا يحدث عند إعادة تشغيل Redis، وهل المفتاح الساخن نفسه ما زال حدًا.`,
`الوقت: 10 دقائق.
قل: ابدأ بالحساب: 10 ملايين في 10 دقائق تساوي حوالي 16,700 في الثانية، وهذا قريب من حد المزوّد. ثم اطلب من الأزواج تصميم الطابور.
انتبه: مستهلك واحد يمر على كل المستخدمين، ولا أحد يذكر حد المزوّد. وتأكد أن أحدًا يتناول الرسائل السيئة (طابور الرسائل الميتة) والتكرار.`,
`الوقت: 10 دقائق.
قل: هذا يعيد الجزء الثالث: زمن الاستجابة وتعارضات تعدد الرؤوس (multi-master) ونظرية CAP. اسأل ماذا يحدث عندما يعدّل العميل نفسه السلة في منطقتين أثناء انقطاع الاتصال.
انتبه: اختيار "آخر كتابة تفوز" دون سؤال عمّا يضيع. وادفعهم إلى الدفع (checkout): الإجابة نفسها لا تصلح هناك.`,
`الوقت: 10 دقائق.
قل: التجزئة (sharding) تجعل الأشياء البسيطة صعبة. اسأل كيف ننقل المال عندما يكون A وB على قطعتين مختلفتين، وماذا يحدث إذا انهار الخادم بينهما.
انتبه: من يلجأ إلى معاملة واحدة كبيرة عبر القطع. وجّههم إلى رقم تحويل (transfer ID) وصندوق صادر (outbox) وإضافة رصيد آمنة عند التكرار، وتأكد أن أحدًا يسأل: هل نحتاج التجزئة أصلًا؟`,
`الوقت: 15 دقيقة.
قل: اختر السؤال الذي يناسب الحاضرين. دع الناس يجيبون بكلماتهم. اختم بأن يلتزم كل شخص بإجراء واحد.`,
`الوقت: دقيقتان.
قل: اطلب من كل شخص أن يختار بندًا واحدًا ويقوله بصوت عالٍ. الجلسة القادمة هي اليوم الخامس: تقنية إعادة الهيكلة وفهرس الروائح (smell catalog).`,
];

// ============================================================ build
TOTAL = queue.length;
queue.forEach((fn, i) => fn(i + 1));
if (PRESENTER.length !== allSlides.length || ARABIC.length !== allSlides.length) console.warn("notes count mismatch: English " + PRESENTER.length + ", Arabic " + ARABIC.length + ", slides " + allSlides.length);
allSlides.forEach((sl, i) => {
  const parts = [];
  if (PRESENTER[i]) parts.push(PRESENTER[i]);
  if (ARABIC[i]) parts.push("──────── العربية (Arabic) ────────\n" + ARABIC[i]);
  if (sl._notes.length) parts.push("MORE (answers and extra points):\n" + sl._notes.join("\n\n"));
  if (parts.length) sl._applyNotes(parts.join("\n\n"));
});
pptx.writeFile({ fileName: "session-04.pptx" }).then(() => {
  // mark Arabic note paragraphs right-to-left (pptxgenjs cannot set this)
  try {
    require("child_process").execFileSync("python3", [__dirname + "/fix_rtl_notes.py", "session-04.pptx"], { stdio: "inherit" });
  } catch (e) { console.warn("RTL post-process skipped:", e.message); }
  console.log("written", TOTAL, "slides");
});
