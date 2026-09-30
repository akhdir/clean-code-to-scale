const pptxgen = require("pptxgenjs");

// ---------- palette (shared with the course's web materials) ----------
const DARK = "0F3038";        // dominant dark bg (title / code / closing)
const DARK2 = "16414B";       // raised panel on dark bg
const DARK_WATERMARK = "1B4750";
const TEAL = "1F5D6B";        // secondary / core accent
const TEAL_TINT = "DCEAEA";
const PLUM = "6B4A8F";        // accent 2
const PLUM_TINT = "EAE2F1";
const AMBER = "A6752B";       // deep gold, for light backgrounds
const AMBER_LIGHT = "E0A85C"; // lighter amber, for the dark code panel
const AMBER_TINT = "F3E6CE";
const RUST = "B3452F";        // semantic "bad"
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
const GRID_LINE = "3A4A4E";

const HEAD = "Cambria";
const BODY = "Calibri";
const CODE = "Courier New";

const TOTAL_SLIDES = 33;

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

function bg(slide, color) {
  slide.background = { color };
}

function footer(slide, label, num, dark) {
  slide.addText(label, {
    x: 0.5, y: 7.12, w: 8, h: 0.3, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 9, color: dark ? "6C8B90" : INK_FAINT,
  });
  slide.addText(String(num).padStart(2, "0") + " / " + TOTAL_SLIDES, {
    x: 11.83, y: 7.12, w: 1, h: 0.3, isTextBox: true, margin: 0, align: "right",
    fontFace: CODE, fontSize: 9, color: dark ? "6C8B90" : INK_FAINT,
  });
}

function header(slide, kicker, title, dark, opts = {}) {
  slide.addText(kicker.toUpperCase(), {
    x: 0.5, y: 0.42, w: 12.3, h: 0.3, isTextBox: true, margin: 0,
    fontFace: CODE, fontSize: 12, color: dark ? AMBER_LIGHT : TEAL, charSpacing: 1,
  });
  slide.addText(title, {
    x: 0.5, y: 0.72, w: opts.titleW || 12.3, h: opts.titleH || 0.7, isTextBox: true, margin: 0,
    fontFace: HEAD, bold: true, fontSize: opts.titleSize || 30, color: dark ? WHITE : INK,
  });
}

function circle(slide, x, y, d, fill, label, opts = {}) {
  slide.addShape(pptx.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  slide.addText(label, {
    x, y, w: d, h: d, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: opts.fontFace || CODE, bold: true, fontSize: opts.fontSize || 13,
    color: opts.color || WHITE,
  });
}

function card(slide, x, y, w, h, fill, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" },
    shadow: opts.shadow === false ? undefined : { type: "outer", color: "1A2226", opacity: 0.12, blur: 8, offset: 2, angle: 90 },
  });
}

function codePanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: CODE_BG }, line: { type: "none" } });
}

// Renders one code block (array of {t, c} lines) into a dark panel, with a
// small colored divider label above it ("THE ISSUE" / "THE FIX" / "WITHOUT IT" / "WITH X").
function labeledCode(slide, x, y, w, h, labelText, labelColor, lines, fontSize) {
  slide.addText(labelText.toUpperCase(), {
    x: x + 0.05, y, w: w - 0.1, h: 0.24, isTextBox: true, margin: 0,
    fontFace: CODE, bold: true, fontSize: 9.5, color: labelColor, charSpacing: 0.5,
  });
  const panelY = y + 0.26;
  const panelH = h - 0.26;
  codePanel(slide, x, panelY, w, panelH);
  const runs = lines.map((l, i) => ({ text: l.t, options: { color: l.c, breakLine: i < lines.length - 1 } }));
  slide.addText(runs, {
    x: x + 0.18, y: panelY + 0.08, w: w - 0.36, h: panelH - 0.16, isTextBox: true, margin: 0,
    fontFace: CODE, fontSize: fontSize || 9.5, lineSpacingMultiple: 1.05, valign: "middle",
  });
}

// ============================================================ SLIDE 1 — TITLE
{
  const s = pptx.addSlide();
  bg(s, DARK);
  s.addText("class OrderProcessor:\n    def __init__(self):\n        self.db = Postgres()", {
    x: 7.3, y: 0.5, w: 5.7, h: 2.6, isTextBox: true, margin: 0,
    fontFace: CODE, fontSize: 18, color: DARK_WATERMARK, lineSpacingMultiple: 1.3,
  });
  s.addText("CLEAN CODE TO SCALE   ·   DAY 03 OF 10", {
    x: 0.7, y: 2.3, w: 10, h: 0.35, isTextBox: true, margin: 0,
    fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5,
  });
  s.addText("OOP and Design Patterns", {
    x: 0.7, y: 2.75, w: 11.5, h: 1.4, isTextBox: true, margin: 0,
    fontFace: HEAD, bold: true, fontSize: 44, color: WHITE,
  });
  s.addText("Five principles, each an answer to a failure you've already lived through — plus the four design patterns that turn out to just be those principles, applied.", {
    x: 0.7, y: 4.05, w: 10, h: 0.9, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 16, color: "B9D3D6",
  });
  const stats = [["3:20", "DURATION"], ["PYTHON", "CODE SAMPLES"], ["PAIRS", "2 KATAS"]];
  stats.forEach((st, i) => {
    const x = 0.7 + i * 3.1;
    s.addText(st[0], { x, y: 5.65, w: 2.9, h: 0.45, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 19, color: WHITE });
    s.addText(st[1], { x, y: 6.13, w: 2.9, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 9.5, color: "6C8B90", charSpacing: 1 });
  });
  footer(s, "Clean Code to Scale", 1, true);
}

// ============================================================ SLIDE 2 — AGENDA
{
  const s = pptx.addSlide();
  bg(s, PAPER);
  header(s, "Session Map", "How the next three hours and twenty run", false);
  const steps = [
    ["01", "0:00–1:30", "CONCEPT", TEAL, "Five principles made concrete, four patterns you already know, four quick-rep katas."],
    ["02", "1:30–2:10", "LIVE DEMO", "5C7A85", "One god class, three violations, refactored on screen."],
    ["03", "2:10–3:05", "HANDS-ON", RUST, "Two core katas, in order — the second depends on the first."],
    ["04", "3:05–3:20", "REVIEW", INK_FAINT, "When SOLID is the right call, and when it's too much."],
  ];
  const startX = 1.1, gap = 3.0, cy = 2.55, d = 1.15;
  s.addShape(pptx.ShapeType.line, {
    x: startX + d / 2, y: cy + d / 2, w: gap * 3, h: 0,
    line: { color: LINE, width: 1.5, dashType: "solid" },
  });
  steps.forEach((st, i) => {
    const x = startX + i * gap;
    circle(s, x, cy, d, st[3], st[0], { fontSize: 20 });
    s.addText(st[1], { x: x - 0.55, y: cy + d + 0.18, w: d + 1.1, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: CODE, bold: true, fontSize: 13, color: INK });
    s.addText(st[2], { x: x - 0.55, y: cy + d + 0.5, w: d + 1.1, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: CODE, fontSize: 10, color: st[3], charSpacing: 1 });
    s.addText(st[4], { x: x - 0.8, y: cy + d + 0.95, w: d + 1.6, h: 0.95, isTextBox: true, margin: 0, align: "center", fontFace: BODY, fontSize: 11.5, color: INK_SOFT });
  });
  s.addText("New to OOP? An optional, not-on-the-clock primer on classes and inheritance opens the session — skip it if that's already second nature.", {
    x: 0.5, y: 6.15, w: 12.3, h: 0.35, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT,
  });
  s.addText("Locked to a 2:00 slot? Cut the four quick-rep katas first (~35 min), then the O/L/I fixes and patterns grid as a handout (~25 min) — RepositoryDecorators is the last kata to cut.", {
    x: 0.5, y: 6.55, w: 12.3, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: BODY, italic: true, fontSize: 11, color: INK_FAINT,
  });
  footer(s, "Day 03 · Clean Code to Scale", 2, false);
}

// ============================================================ SLIDE 3 — OBJECTIVES
{
  const s = pptx.addSlide();
  bg(s, PAPER);
  header(s, "Objectives", "What today is for", false);
  const items = [
    ["Spot the violation, not just the diagram", "Recognize an SRP or OCP violation from a diff, without needing five boxes and arrows."],
    ["Invert a dependency for real", "Make a concrete dependency swappable — for production, or for a test double."],
    ["Recognize the patterns you already know", "Strategy, Factory Method, Decorator, and Observer are SOLID applied, not new theory to memorize."],
    ["Know when it's too much", "Know when applying SOLID is over-engineering for the size of the problem in front of you."],
  ];
  items.forEach((it, i) => {
    const y = 1.75 + i * 1.2;
    circle(s, 0.7, y, 0.56, TEAL, "✓", { fontFace: BODY, fontSize: 20 });
    s.addText(it[0], { x: 1.6, y: y - 0.06, w: 10.5, h: 0.42, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 17, color: INK });
    s.addText(it[1], { x: 1.6, y: y + 0.37, w: 10.5, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT });
  });
  footer(s, "Day 03 · Clean Code to Scale", 3, false);
}

// ============================================================ SLIDE 4 — ROOM
{
  const s = pptx.addSlide();
  bg(s, PAPER);
  header(s, "Room", "Why this session, for you specifically", false);
  const roles = [
    ["JR", TEAL, "Junior – Mid Engineers", "SOLID sounds academic until you're the one afraid to touch a class because everything might break. That fear is exactly what these five principles target."],
    ["SR", "5C7A85", "Senior Eng & Tech Leads", "The real skill isn't reciting five letters — it's calling \"this is over-engineered\" in review as confidently as you call \"this needs a class.\""],
    ["QE", RUST, "QEs", "Every dependency this session inverts is a seam your tests can mock. Watch \"needs a real database to test\" become \"pass in a fake.\""],
    ["AI", PLUM, "AI Engineers", "An agent asked to \"add a feature\" will often bolt it onto an existing class rather than extend it. OCP is the difference between those two moves."],
  ];
  const cw = 5.85, ch = 2.1, gx = 0.3, gy = 0.25, ox = 0.5, oy = 1.85;
  roles.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = ox + col * (cw + gx), y = oy + row * (ch + gy);
    card(s, x, y, cw, ch, WHITE);
    circle(s, x + 0.3, y + 0.3, 0.72, r[1], r[0], { fontSize: 14 });
    s.addText(r[2], { x: x + 1.25, y: y + 0.32, w: cw - 1.5, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 15, color: INK });
    s.addText(r[3], { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: INK_SOFT });
  });
  footer(s, "Day 03 · Clean Code to Scale", 4, false);
}

// ============================================================ SLIDE 5 — OOP PRIMER
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Optional · Not in the clock", "New to object-oriented code? Start here.", true, { titleSize: 26 });
  s.addText("Everything today assumes you know what a class and a subclass are — skip ahead to the five principles if that's already second nature.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const classCode = [
    { t: "class Employee:", c: CODE_INK },
    { t: "    def __init__(self, name, rate):", c: CODE_INK },
    { t: "        self.name = name", c: CODE_INK },
    { t: "        self.hourly_rate = rate", c: CODE_INK },
    { t: "        self.vacation_days = 15      # attribute", c: CODE_COMMENT },
    { t: "", c: CODE_INK },
    { t: "    def calculate_pay(self, hours):  # method", c: CODE_COMMENT },
    { t: "        return hours * self.hourly_rate", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def describe(self):              # method", c: CODE_COMMENT },
    { t: '        return f"{self.name}: {self.vacation_days} vacation days/year"', c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: 'employee = Employee("Sam", 25)   # an object', c: CODE_INK },
    { t: "print(employee.calculate_pay(40))  # prints: 1000", c: CODE_INK },
  ];
  const inheritCode = [
    { t: "class Manager(Employee):            # inherits from Employee", c: CODE_COMMENT },
    { t: "    def calculate_pay(self, hours): # overrides this one method", c: CODE_COMMENT },
    { t: "        return hours * self.hourly_rate * 1.15", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    # no __init__, no describe() — inherited untouched", c: CODE_COMMENT },
    { t: "", c: CODE_INK },
    { t: 'manager = Manager("Jordan", 25)', c: CODE_INK },
    { t: "print(manager.calculate_pay(40))", c: CODE_INK },
    { t: "# prints: 1150.0 — Manager's own method ran", c: CODE_COMMENT },
    { t: "", c: CODE_INK },
    { t: "print(manager.describe())", c: CODE_INK },
    { t: "# prints: Jordan: 15 vacation days/year", c: CODE_COMMENT },
    { t: "# describe() and vacation_days came from Employee,", c: CODE_COMMENT },
    { t: "# untouched — that's the real value of inheritance", c: CODE_COMMENT },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 3.95, "What is a class?", TEAL_TINT, classCode, 9.7);
  labeledCode(s, 6.75, 1.75, 6.1, 3.95, "What is inheritance?", TEAL_TINT, inheritCode, 9.3);

  card(s, 0.5, 5.85, 12.35, 0.95, DARK2, { shadow: false });
  s.addText([
    { text: "The takeaway:  ", options: { fontFace: BODY, bold: true, fontSize: 12.5, color: AMBER_LIGHT, breakLine: false } },
    { text: "inheritance's value isn't just permission to override — it's free reuse of the methods and attributes a subclass never has to touch.", options: { fontFace: BODY, fontSize: 12.5, color: "C6DADC" } },
  ], { x: 0.75, y: 5.85, w: 11.85, h: 0.95, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, "Day 03 · Clean Code to Scale", 5, true);
}

// ============================================================ SLIDE 6 — FIVE PRINCIPLES
{
  const s = pptx.addSlide();
  bg(s, PAPER);
  header(s, "0:00–1:30 · Concept", "Five letters, five failures", false);
  s.addText("Each principle is a named answer to a way codebases actually break. Today's demo and kata live mostly in the last one.", { x: 0.5, y: 1.45, w: 12, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK_SOFT });
  const principles = [
    ["S", TEAL, "Single Responsibility — one reason to change", "An InvoiceGenerator that also formats and sends emails has two reasons to change: pricing rules, and email wording."],
    ["O", PLUM, "Open/Closed — open for extension, closed for modification", "A new discount type shouldn't mean editing an if/elif chain — it should mean adding one new class."],
    ["L", AMBER, "Liskov Substitution — a subclass shouldn't surprise the caller", "A GiftCardPayment that raises on .refund() breaks any caller that trusts every Payment supports it."],
    ["I", RUST, "Interface Segregation — don't force unused methods", "A RobotWorker forced to implement .eat_lunch() and .sleep() only because every Worker does."],
    ["D", TEAL, "Dependency Inversion — depend on abstractions, not concretions", "An OrderService that builds a PostgresRepository inside itself, instead of receiving a repository interface."],
  ];
  const oy = 2.0, rh = 0.92;
  principles.forEach((p, i) => {
    const y = oy + i * rh;
    if (i === 4) {
      card(s, 0.4, y - 0.06, 12.5, rh - 0.06, TEAL_TINT, { shadow: false });
    }
    circle(s, 0.65, y, 0.5, p[1], p[0], { fontSize: 17 });
    s.addText(p[2], { x: 1.4, y: y - 0.08, w: 11.3, h: 0.34, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 13.5, color: INK });
    s.addText(p[3], { x: 1.4, y: y + 0.26, w: 11.3, h: 0.5, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 10, color: INK_SOFT });
  });
  footer(s, "Day 03 · Clean Code to Scale", 6, false);
}

// ============================================================ SLIDE 7 — OCP ISSUE + FIX
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Open/Closed, made concrete", true, { titleSize: 27 });
  s.addText("An order-status handler with one elif per status. Every new status means editing a function that already works.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const issue = [
    { t: "def handle_status_change(order):", c: CODE_INK },
    { t: '    if order.status == "paid":', c: CODE_INK },
    { t: "        send_confirmation_email(order)", c: CODE_INK },
    { t: '    elif order.status == "shipped":', c: CODE_INK },
    { t: "        send_tracking_email(order)", c: CODE_INK },
    { t: '    elif order.status == "delivered":', c: CODE_INK },
    { t: "        request_review(order)", c: CODE_INK },
    { t: '    elif order.status == "cancelled":', c: CODE_INK },
    { t: "        send_cancellation_email(order)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: '# adding "refunded" means editing this function', c: AMBER_LIGHT },
    { t: "# and re-testing every status that already worked", c: CODE_COMMENT },
  ];
  const fix = [
    { t: "STATUS_HANDLERS = {", c: CODE_INK },
    { t: '    "paid": send_confirmation_email,', c: CODE_INK },
    { t: '    "shipped": send_tracking_email,', c: CODE_INK },
    { t: '    "delivered": request_review,', c: CODE_INK },
    { t: '    "cancelled": send_cancellation_email,', c: CODE_INK },
    { t: "}", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "def handle_status_change(order):", c: CODE_INK },
    { t: "    handler = STATUS_HANDLERS.get(order.status)", c: CODE_INK },
    { t: "    if handler:", c: CODE_INK },
    { t: "        handler(order)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: '# adding "refunded" is one new dict entry —', c: CODE_COMMENT },
    { t: "# handle_status_change never changes again", c: CODE_COMMENT },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 5.05, "The issue", AMBER_LIGHT, issue, 10.5);
  labeledCode(s, 6.75, 1.75, 6.1, 5.05, "The fix — a lookup table, not a chain", TEAL_TINT, fix, 9.7);
  footer(s, "Day 03 · Clean Code to Scale", 7, true);
}

// ============================================================ SLIDE 8 — LSP ISSUE + FIX
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Liskov Substitution, made concrete", true, { titleSize: 26 });
  s.addText("Every Payment is supposed to support .refund(). A gift card can't be — so code that trusts the contract crashes eventually.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const issue = [
    { t: "class Payment:", c: CODE_INK },
    { t: "    def refund(self, amount): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class CreditCardPayment(Payment):", c: CODE_INK },
    { t: "    def refund(self, amount):", c: CODE_INK },
    { t: "        gateway.refund(self.card_token, amount)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class GiftCardPayment(Payment):", c: CODE_INK },
    { t: "    def refund(self, amount):", c: CODE_INK },
    { t: '        raise NotImplementedError("not refundable")', c: AMBER_LIGHT },
    { t: "", c: CODE_INK },
    { t: "for payment in customer.payments:", c: CODE_INK },
    { t: "    payment.refund(order.total)  # boom, eventually", c: CODE_COMMENT },
  ];
  const fix = [
    { t: "class Payment:", c: CODE_INK },
    { t: "    ...  # whatever every payment shares", c: CODE_COMMENT },
    { t: "", c: CODE_INK },
    { t: "class Refundable:", c: CODE_INK },
    { t: "    def refund(self, amount): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class CreditCardPayment(Payment, Refundable):", c: CODE_INK },
    { t: "    def refund(self, amount):", c: CODE_INK },
    { t: "        gateway.refund(self.card_token, amount)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class GiftCardPayment(Payment):", c: CODE_INK },
    { t: "    pass  # no refund() — never a real capability", c: CODE_COMMENT },
    { t: "", c: CODE_INK },
    { t: "refundable = [p for p in customer.payments", c: CODE_INK },
    { t: "              if isinstance(p, Refundable)]", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 5.05, "The issue", AMBER_LIGHT, issue, 10);
  labeledCode(s, 6.75, 1.75, 6.1, 5.05, "The fix — refund becomes its own capability", TEAL_TINT, fix, 9.3);
  footer(s, "Day 03 · Clean Code to Scale", 8, true);
}

// ============================================================ SLIDE 9 — ISP ISSUE + FIX
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Interface Segregation, made concrete", true, { titleSize: 26 });
  s.addText("The classic version of this smell: a Worker interface with a lunch break and a sleep schedule — two things a robot doesn't have.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const issue = [
    { t: "class Worker:", c: CODE_INK },
    { t: "    def work(self): raise NotImplementedError", c: CODE_INK },
    { t: "    def eat_lunch(self): raise NotImplementedError", c: CODE_INK },
    { t: "    def sleep(self): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class HumanWorker(Worker):", c: CODE_INK },
    { t: "    def work(self): assemble_parts()", c: CODE_INK },
    { t: "    def eat_lunch(self): take_break()", c: CODE_INK },
    { t: "    def sleep(self): clock_out()", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class RobotWorker(Worker):", c: CODE_INK },
    { t: "    def work(self): assemble_parts()", c: CODE_INK },
    { t: "    def eat_lunch(self):", c: CODE_INK },
    { t: '        raise NotImplementedError("robots don\'t eat")', c: AMBER_LIGHT },
    { t: "    def sleep(self):", c: CODE_INK },
    { t: '        raise NotImplementedError("robots don\'t sleep")', c: AMBER_LIGHT },
  ];
  const fix = [
    { t: "class Workable:", c: CODE_INK },
    { t: "    def work(self): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class Schedulable:", c: CODE_INK },
    { t: "    def eat_lunch(self): raise NotImplementedError", c: CODE_INK },
    { t: "    def sleep(self): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class HumanWorker(Workable, Schedulable):", c: CODE_INK },
    { t: "    def work(self): assemble_parts()", c: CODE_INK },
    { t: "    def eat_lunch(self): take_break()", c: CODE_INK },
    { t: "    def sleep(self): clock_out()", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class RobotWorker(Workable):", c: CODE_INK },
    { t: "    def work(self): assemble_parts()", c: CODE_INK },
    { t: "    # no eat_lunch or sleep — never had a use for either", c: CODE_COMMENT },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 5.05, "The issue", AMBER_LIGHT, issue, 10);
  labeledCode(s, 6.75, 1.75, 6.1, 5.05, "The fix — split the fat interface in two", TEAL_TINT, fix, 9.7);
  footer(s, "Day 03 · Clean Code to Scale", 9, true);
}

// ============================================================ SLIDE 10 — DEPENDENCY INVERSION ISSUE + FIX
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Dependency Inversion, made concrete", true, { titleSize: 26 });
  s.addText("A service that builds its own dependency is welded to it. Receive it instead, and the swap — for production or a test double — costs nothing.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const issue = [
    { t: "class OrderService:", c: CODE_INK },
    { t: "    def __init__(self):", c: CODE_INK },
    { t: "        self.repository = PostgresRepository()", c: AMBER_LIGHT },
    { t: "", c: CODE_INK },
    { t: "    def place_order(self, order):", c: CODE_INK },
    { t: "        total = calculate_total(order)", c: CODE_INK },
    { t: "        self.repository.save(order.id, total)", c: CODE_INK },
    { t: "        return total", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# swapping providers, or testing without a real", c: CODE_COMMENT },
    { t: "# database, means editing OrderService itself", c: CODE_COMMENT },
  ];
  const fix = [
    { t: "class OrderService:", c: CODE_INK },
    { t: "    def __init__(self, repository):", c: CODE_INK },
    { t: "        self.repository = repository", c: TEAL_TINT },
    { t: "", c: CODE_INK },
    { t: "    def place_order(self, order):", c: CODE_INK },
    { t: "        total = calculate_total(order)", c: CODE_INK },
    { t: "        self.repository.save(order.id, total)", c: CODE_INK },
    { t: "        return total", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# production", c: CODE_COMMENT },
    { t: "service = OrderService(PostgresRepository())", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# test — no real database", c: CODE_COMMENT },
    { t: "service = OrderService(FakeRepository())", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 5.05, "The issue", AMBER_LIGHT, issue, 10.5);
  labeledCode(s, 6.75, 1.75, 6.1, 5.05, "The fix — receive the dependency, don't build it", TEAL_TINT, fix, 10);
  footer(s, "Day 03 · Clean Code to Scale", 10, true);
}

// ============================================================ SLIDE 11 — PATTERN: STRATEGY
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Pattern: Strategy", true, { titleSize: 28 });
  s.addText("OCP · DIP  —  you already built this today, as DiscountStrategy. Same code, in miniature.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const before = [
    { t: "def calculate_discount(total, discount_type):", c: CODE_INK },
    { t: '    if discount_type == "loyalty":', c: CODE_INK },
    { t: "        return total * 0.9", c: CODE_INK },
    { t: '    elif discount_type == "seasonal":', c: CODE_INK },
    { t: "        return total * 0.85", c: CODE_INK },
    { t: "    return total", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# adding \"student\" means editing this function", c: CODE_COMMENT },
  ];
  const after = [
    { t: "class DiscountStrategy:", c: CODE_INK },
    { t: "    def apply(self, total): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class SeasonalDiscount(DiscountStrategy):", c: CODE_INK },
    { t: "    def apply(self, total): return total * 0.85", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# adding \"student\" means adding a class", c: CODE_COMMENT },
    { t: "total = discount.apply(total)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 4.0, "Without it", AMBER_LIGHT, before, 11.5);
  labeledCode(s, 6.75, 1.75, 6.1, 4.0, "With Strategy", TEAL_TINT, after, 11.5);

  card(s, 0.5, 5.9, 12.35, 0.9, DARK2, { shadow: false });
  s.addText([
    { text: "What improved:  ", options: { fontFace: BODY, bold: true, fontSize: 12.5, color: AMBER_LIGHT, breakLine: false } },
    { text: "the function that varies is gone. Nothing left to re-test when a discount type is added — there's only something to add.", options: { fontFace: BODY, fontSize: 12.5, color: "C6DADC" } },
  ], { x: 0.75, y: 5.9, w: 11.85, h: 0.9, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, "Day 03 · Clean Code to Scale", 11, true);
}

// ============================================================ SLIDE 12 — PATTERN: FACTORY METHOD
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Pattern: Factory Method", true, { titleSize: 28 });
  s.addText("DIP · OCP  —  centralizes object construction in one place, instead of an if/elif at every call site that needs one.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const before = [
    { t: "# repeated at every call site that needs a notifier", c: CODE_COMMENT },
    { t: 'if customer.channel == "email":', c: CODE_INK },
    { t: "    notifier = EmailNotifier(smtp_host=\"smtp.x.com\")", c: CODE_INK },
    { t: 'elif customer.channel == "sms":', c: CODE_INK },
    { t: '    notifier = SmsNotifier(api_key="xxxx")', c: CODE_INK },
    { t: "notifier.send(message)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# ...copy-pasted again, three files over", c: CODE_COMMENT },
  ];
  const after = [
    { t: "class NotifierFactory:", c: CODE_INK },
    { t: "    def create(self, channel):", c: CODE_INK },
    { t: '        if channel == "email": return EmailNotifier()', c: CODE_INK },
    { t: '        if channel == "sms": return SmsNotifier()', c: CODE_INK },
    { t: "        raise ValueError(channel)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# every call site asks, none of them decide", c: CODE_COMMENT },
    { t: "notifier = NotifierFactory().create(customer.channel)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 4.0, "Without it", AMBER_LIGHT, before, 11.5);
  labeledCode(s, 6.75, 1.75, 6.1, 4.0, "With Factory Method", TEAL_TINT, after, 11.5);

  card(s, 0.5, 5.9, 12.35, 0.9, DARK2, { shadow: false });
  s.addText([
    { text: "What improved:  ", options: { fontFace: BODY, bold: true, fontSize: 12.5, color: AMBER_LIGHT, breakLine: false } },
    { text: "construction logic lived in five places before — a new channel meant finding and editing all five. Now it exists once.", options: { fontFace: BODY, fontSize: 12.5, color: "C6DADC" } },
  ], { x: 0.75, y: 5.9, w: 11.85, h: 0.9, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, "Day 03 · Clean Code to Scale", 12, true);
}

// ============================================================ SLIDE 13 — PATTERN: DECORATOR
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Pattern: Decorator", true, { titleSize: 28 });
  s.addText("OCP · SRP  —  adds behavior by wrapping an object, not editing it. Today's new kata builds one on a repository.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const before = [
    { t: "class EmailNotifier:", c: CODE_INK },
    { t: "    def send(self, to, message):", c: CODE_INK },
    { t: "        for attempt in range(3):", c: CODE_INK },
    { t: "            try:", c: CODE_INK },
    { t: "                smtp.send(to, message)", c: CODE_INK },
    { t: "                return", c: CODE_INK },
    { t: "            except SmtpError:", c: CODE_INK },
    { t: "                if attempt == 2: raise", c: CODE_INK },
    { t: "                time.sleep(0.5)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# every notifier class repeats this same loop", c: CODE_COMMENT },
  ];
  const after = [
    { t: "class RetryingNotifier:", c: CODE_INK },
    { t: "    def __init__(self, notifier, attempts=3):", c: CODE_INK },
    { t: "        self.notifier = notifier", c: CODE_INK },
    { t: "        self.attempts = attempts", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def send(self, to, message):", c: CODE_INK },
    { t: "        for attempt in range(self.attempts):", c: CODE_INK },
    { t: "            try:", c: CODE_INK },
    { t: "                return self.notifier.send(to, message)", c: CODE_INK },
    { t: "            except SmtpError:", c: CODE_INK },
    { t: "                if attempt == self.attempts - 1: raise", c: CODE_INK },
    { t: "                time.sleep(0.5)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "notifier = RetryingNotifier(EmailNotifier())", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.75, 6.1, 4.0, "Without it", AMBER_LIGHT, before, 10);
  labeledCode(s, 6.75, 1.75, 6.1, 4.0, "With Decorator", TEAL_TINT, after, 9.3);

  card(s, 0.5, 5.9, 12.35, 0.9, DARK2, { shadow: false });
  s.addText([
    { text: "What improved:  ", options: { fontFace: BODY, bold: true, fontSize: 12.5, color: AMBER_LIGHT, breakLine: false } },
    { text: "the retry loop was copy-pasted into every notifier before. Now it wraps any of them once, and EmailNotifier just sends email.", options: { fontFace: BODY, fontSize: 12.5, color: "C6DADC" } },
  ], { x: 0.75, y: 5.9, w: 11.85, h: 0.9, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, "Day 03 · Clean Code to Scale", 13, true);
}

// ============================================================ SLIDE 14 — PATTERN: OBSERVER (full implementation)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "0:00–1:30 · Concept", "Pattern: Observer", true, { titleSize: 28 });
  s.addText("SRP · DIP  —  the subject doesn't know or care what reacts to it. Full implementation below, not just the usage.", { x: 0.5, y: 1.3, w: 12, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: "6C8B90" });

  const before = [
    { t: "class OrderService:", c: CODE_INK },
    { t: "    def place_order(self, order):", c: CODE_INK },
    { t: "        save_order(order)", c: CODE_INK },
    { t: "        send_confirmation_email(order)", c: CODE_INK },
    { t: "        update_analytics(order)", c: CODE_INK },
    { t: "        notify_warehouse(order)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: '# "text the courier" means editing this', c: CODE_COMMENT },
    { t: "# method, and OrderService now knows", c: CODE_COMMENT },
    { t: "# about four other teams", c: CODE_COMMENT },
  ];
  const after = [
    { t: "class OrderEvents:", c: CODE_INK },
    { t: "    def __init__(self):", c: CODE_INK },
    { t: "        self._subscribers = []", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def on_placed(self, callback):", c: CODE_INK },
    { t: "        self._subscribers.append(callback)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def emit(self, order):", c: CODE_INK },
    { t: "        for callback in self._subscribers:", c: CODE_INK },
    { t: "            callback(order)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# wiring: each team subscribes itself, on startup", c: CODE_COMMENT },
    { t: "order_events = OrderEvents()", c: CODE_INK },
    { t: "order_events.on_placed(send_confirmation_email)", c: CODE_INK },
    { t: "order_events.on_placed(update_analytics)", c: CODE_INK },
    { t: "order_events.on_placed(notify_warehouse)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# the subject: doesn't import or know about any of them", c: CODE_COMMENT },
    { t: "class OrderService:", c: CODE_INK },
    { t: "    def __init__(self, events):", c: CODE_INK },
    { t: "        self.events = events", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def place_order(self, order):", c: CODE_INK },
    { t: "        save_order(order)", c: CODE_INK },
    { t: "        self.events.emit(order)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.75, 4.3, 4.15, "Without it", AMBER_LIGHT, before, 10.5);
  labeledCode(s, 5.0, 1.75, 7.85, 4.15, "With Observer — the whole implementation", TEAL_TINT, after, 9.3);

  card(s, 0.5, 6.05, 12.35, 0.85, DARK2, { shadow: false });
  s.addText([
    { text: "What improved:  ", options: { fontFace: BODY, bold: true, fontSize: 11.5, color: AMBER_LIGHT, breakLine: false } },
    { text: "place_order() used to grow a line for every team that cared about a new order. Now it emits one event — subscribing is each team's job, not OrderService's. events arrives through the constructor, not a global — DIP again.", options: { fontFace: BODY, fontSize: 11.5, color: "C6DADC" } },
  ], { x: 0.75, y: 6.05, w: 11.85, h: 0.85, isTextBox: true, margin: 0, valign: "middle" });
  footer(s, "Day 03 · Clean Code to Scale", 14, true);
}

// ============================================================ Helper for quick-rep kata slides
function quickRepKata(slideNum, kicker, title, task, starterLabel, starterLines, doneWhen, mirrorNote, fontSize) {
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, kicker, title, true, { titleSize: 27 });
  s.addText(task, { x: 0.5, y: 1.3, w: 12, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11.5, color: "6C8B90" });

  codePanel(s, 0.6, 1.85, 12.1, 2.3);
  const runs = starterLines.map((l, i) => ({ text: l, options: { breakLine: i < starterLines.length - 1 } }));
  s.addText(runs, { x: 0.9, y: 1.97, w: 11.5, h: 2.05, isTextBox: true, margin: 0, fontFace: CODE, fontSize: fontSize || 13, color: CODE_INK, lineSpacingMultiple: 1.0, valign: "middle" });

  s.addShape(pptx.ShapeType.line, { x: 0.6, y: 4.35, w: 12.1, h: 0, line: { color: "234750", width: 1 } });
  s.addText([
    { text: "Done when:  ", options: { fontFace: BODY, bold: true, fontSize: 12.5, color: AMBER_LIGHT, breakLine: false } },
    { text: doneWhen, options: { fontFace: BODY, italic: true, fontSize: 12.5, color: AMBER_LIGHT } },
  ], { x: 0.6, y: 4.48, w: 12.1, h: 0.4, isTextBox: true, margin: 0 });

  card(s, 0.6, 5.15, 12.1, 1.15, DARK2, { shadow: false });
  s.addText([
    { text: "The move:  ", options: { fontFace: BODY, bold: true, fontSize: 12.5, color: TEAL_TINT, breakLine: false } },
    { text: mirrorNote, options: { fontFace: BODY, fontSize: 12.5, color: "C6DADC" } },
  ], { x: 0.85, y: 5.15, w: 11.6, h: 1.15, isTextBox: true, margin: 0, valign: "middle" });

  footer(s, "Day 03 · Clean Code to Scale", slideNum, true);
}

// ============================================================ SLIDE 15 — QUICK REP: STRATEGY (kata)
quickRepKata(
  15, "Quick Rep · Strategy · 8 min", "ShippingCostCalculator — swap the elif chain",
  "Replace the elif chain with a ShippingStrategy hierarchy, the same shape as DiscountStrategy.",
  "shipping_cost.py",
  [
    'def calculate_shipping(weight_kg, carrier):',
    '    if carrier == "ups":',
    '        return weight_kg * 4.5',
    '    elif carrier == "fedex":',
    '        return weight_kg * 4.8',
    '    elif carrier == "usps":',
    '        return weight_kg * 3.9',
    '    raise ValueError(carrier)',
  ],
  "adding \"dhl\" means adding a class — calculate_shipping's replacement never changes again",
  "Same move as Strategy above: one class per carrier, a common .cost() method, and the elif chain disappears for good.",
  13
);

// ============================================================ SLIDE 16 — QUICK REP: STRATEGY (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Quick Rep · Strategy · 8 min", "ShippingCostCalculator — original vs. solution", true, { titleSize: 25 });

  const original = [
    { t: "def calculate_shipping(weight_kg, carrier):", c: CODE_INK },
    { t: '    if carrier == "ups":', c: CODE_INK },
    { t: "        return weight_kg * 4.5", c: CODE_INK },
    { t: '    elif carrier == "fedex":', c: CODE_INK },
    { t: "        return weight_kg * 4.8", c: CODE_INK },
    { t: '    elif carrier == "usps":', c: CODE_INK },
    { t: "        return weight_kg * 3.9", c: CODE_INK },
    { t: "    raise ValueError(carrier)", c: CODE_INK },
  ];
  const solution = [
    { t: "class ShippingStrategy:", c: CODE_INK },
    { t: "    def cost(self, weight_kg): raise NotImplementedError", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class UpsShipping(ShippingStrategy):", c: CODE_INK },
    { t: "    def cost(self, weight_kg): return weight_kg * 4.5", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class FedexShipping(ShippingStrategy):", c: CODE_INK },
    { t: "    def cost(self, weight_kg): return weight_kg * 4.8", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class UspsShipping(ShippingStrategy):", c: CODE_INK },
    { t: "    def cost(self, weight_kg): return weight_kg * 3.9", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: '# adding "dhl" is one new class — no existing code changes', c: CODE_COMMENT },
    { t: "cost = strategy.cost(weight_kg)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 6.1, 5.3, "The original", AMBER_LIGHT, original, 12);
  labeledCode(s, 6.75, 1.5, 6.1, 5.3, "The solution", TEAL_TINT, solution, 10.5);
  footer(s, "Day 03 · Clean Code to Scale", 16, true);
}

// ============================================================ SLIDE 17 — QUICK REP: FACTORY METHOD (kata)
quickRepKata(
  17, "Quick Rep · Factory Method · 7 min", "Database connections — extract a ConnectionFactory",
  "This exact if/elif also lives in two startup scripts. Extract a ConnectionFactory so it exists once.",
  "db_connection.py",
  [
    'def get_database_connection(env):',
    '    if env == "production":',
    '        return PostgresConnection(host="prod.db.internal")',
    '    elif env == "staging":',
    '        return PostgresConnection(host="staging.db.internal")',
    '    elif env == "local":',
    '        return SqliteConnection(path="local.db")',
    '    raise ValueError(env)',
  ],
  "every call site asks the factory for a connection — none of them decide the class",
  "Same move as Factory Method above: the if/elif doesn't vanish, it just moves into one class that every caller asks instead of repeats.",
  12.5
);

// ============================================================ SLIDE 18 — QUICK REP: FACTORY METHOD (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Quick Rep · Factory Method · 7 min", "Database connections — original vs. solution", true, { titleSize: 24 });

  const original = [
    { t: "def get_database_connection(env):", c: CODE_INK },
    { t: '    if env == "production":', c: CODE_INK },
    { t: '        return PostgresConnection(host="prod.db.internal")', c: CODE_INK },
    { t: '    elif env == "staging":', c: CODE_INK },
    { t: '        return PostgresConnection(host="staging.db.internal")', c: CODE_INK },
    { t: '    elif env == "local":', c: CODE_INK },
    { t: '        return SqliteConnection(path="local.db")', c: CODE_INK },
    { t: "    raise ValueError(env)", c: CODE_INK },
  ];
  const solution = [
    { t: "class ConnectionFactory:", c: CODE_INK },
    { t: "    def create(self, env):", c: CODE_INK },
    { t: '        if env == "production":', c: CODE_INK },
    { t: '            return PostgresConnection(host="prod.db.internal")', c: CODE_INK },
    { t: '        if env == "staging":', c: CODE_INK },
    { t: '            return PostgresConnection(host="staging.db.internal")', c: CODE_INK },
    { t: '        if env == "local":', c: CODE_INK },
    { t: '            return SqliteConnection(path="local.db")', c: CODE_INK },
    { t: "        raise ValueError(env)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# every script asks the factory — including the two other", c: CODE_COMMENT },
    { t: "# startup scripts that used to carry this same if/elif", c: CODE_COMMENT },
    { t: "connection = ConnectionFactory().create(env)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 6.1, 5.3, "The original", AMBER_LIGHT, original, 10.5);
  labeledCode(s, 6.75, 1.5, 6.1, 5.3, "The solution", TEAL_TINT, solution, 9.3);
  footer(s, "Day 03 · Clean Code to Scale", 18, true);
}

// ============================================================ SLIDE 19 — QUICK REP: DECORATOR (kata)
quickRepKata(
  19, "Quick Rep · Decorator · 8 min", "CachingPriceLookup — cache without touching PriceLookup",
  "get_price() hits the network every call, even for the same SKU twice in a row. Write a decorator that caches, without touching PriceLookup.",
  "price_lookup.py",
  [
    'class PriceLookup:',
    '    def get_price(self, sku):',
    '        return pricing_api.fetch(sku)  # slow, every time',
  ],
  "a second call for the same SKU never reaches pricing_api",
  "Same move as Decorator above: wrap the object instead of editing it. PriceLookup never learns it's being cached.",
  15
);

// ============================================================ SLIDE 20 — QUICK REP: DECORATOR (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Quick Rep · Decorator · 8 min", "CachingPriceLookup — original vs. solution", true, { titleSize: 25 });

  const original = [
    { t: "class PriceLookup:", c: CODE_INK },
    { t: "    def get_price(self, sku):", c: CODE_INK },
    { t: "        # slow, every time", c: CODE_COMMENT },
    { t: "        return pricing_api.fetch(sku)", c: CODE_INK },
  ];
  const solution = [
    { t: "class CachingPriceLookup:", c: CODE_INK },
    { t: "    def __init__(self, price_lookup):", c: CODE_INK },
    { t: "        self.price_lookup = price_lookup", c: CODE_INK },
    { t: "        self._cache = {}", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def get_price(self, sku):", c: CODE_INK },
    { t: "        if sku not in self._cache:", c: CODE_INK },
    { t: "            self._cache[sku] = self.price_lookup.get_price(sku)", c: CODE_INK },
    { t: "        return self._cache[sku]", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# a second call for the same SKU never reaches pricing_api", c: CODE_COMMENT },
    { t: "prices = CachingPriceLookup(PriceLookup())", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 6.1, 5.3, "The original", AMBER_LIGHT, original, 13);
  labeledCode(s, 6.75, 1.5, 6.1, 5.3, "The solution", TEAL_TINT, solution, 10.5);
  footer(s, "Day 03 · Clean Code to Scale", 20, true);
}

// ============================================================ SLIDE 21 — QUICK REP: OBSERVER (kata)
quickRepKata(
  21, "Quick Rep · Observer · 8 min", "InventoryService — give restock() an event",
  "Give restock() the same OrderEvents-shaped treatment you just read, with an on_restocked event.",
  "inventory_service.py",
  [
    'class InventoryService:',
    '    def restock(self, item, qty):',
    '        item.stock += qty',
    '        send_restock_email(item)       # marketing',
    '        update_reorder_forecast(item)  # ops',
  ],
  "restock() only updates state and emits — subscribing is each team's job",
  "Same move as Observer above: a subscriber list, an on_X to join it, an emit to walk it. Structurally identical to OrderEvents.",
  14
);

// ============================================================ SLIDE 22 — QUICK REP: OBSERVER (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Quick Rep · Observer · 8 min", "InventoryService — original vs. solution", true, { titleSize: 25 });

  const original = [
    { t: "class InventoryService:", c: CODE_INK },
    { t: "    def restock(self, item, qty):", c: CODE_INK },
    { t: "        item.stock += qty", c: CODE_INK },
    { t: "        # marketing", c: CODE_COMMENT },
    { t: "        send_restock_email(item)", c: CODE_INK },
    { t: "        # ops", c: CODE_COMMENT },
    { t: "        update_reorder_forecast(item)", c: CODE_INK },
  ];
  const solution = [
    { t: "class InventoryEvents:", c: CODE_INK },
    { t: "    def __init__(self):", c: CODE_INK },
    { t: "        self._subscribers = []", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def on_restocked(self, callback):", c: CODE_INK },
    { t: "        self._subscribers.append(callback)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def emit(self, item):", c: CODE_INK },
    { t: "        for callback in self._subscribers:", c: CODE_INK },
    { t: "            callback(item)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# wiring: each team subscribes itself, on startup", c: CODE_COMMENT },
    { t: "inventory_events = InventoryEvents()", c: CODE_INK },
    { t: "inventory_events.on_restocked(send_restock_email)", c: CODE_INK },
    { t: "inventory_events.on_restocked(update_reorder_forecast)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# the subject: doesn't import or know about either team", c: CODE_COMMENT },
    { t: "class InventoryService:", c: CODE_INK },
    { t: "    def __init__(self, events):", c: CODE_INK },
    { t: "        self.events = events", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def restock(self, item, qty):", c: CODE_INK },
    { t: "        item.stock += qty", c: CODE_INK },
    { t: "        self.events.emit(item)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 4.3, 5.3, "The original", AMBER_LIGHT, original, 11);
  labeledCode(s, 5.0, 1.5, 7.85, 5.3, "The solution", TEAL_TINT, solution, 10);
  footer(s, "Day 03 · Clean Code to Scale", 22, true);
}

// ============================================================ SLIDE 23 — DEMO: BEFORE
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "1:30–2:10 · Live Demo", "order_processor.py — before", true, { titleSize: 27 });
  s.addText("SRP, OCP, and DIP all violated at once — find the three before revealing the talking points", { x: 0.5, y: 1.3, w: 10, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11.5, color: "6C8B90" });

  codePanel(s, 0.5, 1.8, 7.3, 3.5);
  const before = [
    'class OrderProcessor:',
    '    def __init__(self):',
    '        self.db = PostgresDatabase()',
    '        self.email = SmtpEmailService()',
    '',
    '    def process(self, order):',
    '        if not order.items:',
    '            raise ValueError("empty order")',
    '        total = sum(i.price * i.qty for i in order.items)',
    '        if order.discount_type == "loyalty":',
    '            total *= 0.9',
    '        elif order.discount_type == "seasonal":',
    '            total *= 0.85',
    '        elif order.discount_type == "employee":',
    '            total *= 0.7',
    '        self.db.save(order.id, total)',
    '        self.email.send(order.customer_email, f"${total:.2f}")',
    '        return total',
  ].join("\n");
  s.addText(before, { x: 0.75, y: 1.95, w: 6.8, h: 3.2, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 10.5, color: CODE_INK, lineSpacingMultiple: 1.0 });

  card(s, 8.05, 1.8, 4.8, 5.15, DARK2, { shadow: false });
  s.addText("TALKING POINTS — BEFORE", { x: 8.35, y: 2.05, w: 4.2, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  const points = [
    "Three reasons to change: pricing math, discount rules, email wording. SRP says one.",
    "The if/elif discount chain means every new type means editing this method — OCP violated.",
    "PostgresDatabase() and SmtpEmailService() are built inside __init__. Swap either and you edit this class — DIP, inverted the wrong way.",
  ];
  points.forEach((p, i) => {
    const y = 2.55 + i * 1.35;
    circle(s, 8.35, y, 0.38, TEAL, String(i + 1), { color: WHITE, fontSize: 13 });
    s.addText(p, { x: 8.88, y: y - 0.1, w: 3.85, h: 1.2, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: "C6DADC" });
  });
  footer(s, "Day 03 · Clean Code to Scale", 23, true);
}

// ============================================================ SLIDE 24 — DEMO: AFTER
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "1:30–2:10 · Live Demo", "order_processor.py — after", true, { titleSize: 27 });
  s.addText("One job each — new discounts add a class, DiscountStrategy classes omitted below for space", { x: 0.5, y: 1.3, w: 10.5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11.5, color: "6C8B90" });

  codePanel(s, 0.5, 1.8, 7.3, 3.9);
  const afterLines = [
    { t: "# repository and notifier are received, not built (DIP)", c: CODE_COMMENT },
    { t: "class OrderProcessor:", c: CODE_INK },
    { t: "    def __init__(self, repository, notifier):", c: CODE_INK },
    { t: "        self.repository = repository", c: CODE_INK },
    { t: "        self.notifier = notifier", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def process(self, order, discount: DiscountStrategy):", c: CODE_INK },
    { t: "        if not order.items:", c: CODE_INK },
    { t: '            raise ValueError("empty order")', c: CODE_INK },
    { t: "        total = sum(i.price * i.qty for i in order.items)", c: CODE_INK },
    { t: "        total = discount.apply(total)", c: CODE_INK },
    { t: "        self.repository.save(order.id, total)", c: CODE_INK },
    { t: "        self.notifier.notify(order.customer_email, total)", c: CODE_INK },
    { t: "        return total", c: CODE_INK },
  ];
  const runs = afterLines.map((l, i) => ({ text: l.t, options: { color: l.c, breakLine: i < afterLines.length - 1 } }));
  s.addText(runs, { x: 0.75, y: 1.95, w: 6.8, h: 3.6, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 10.5, lineSpacingMultiple: 1.05 });

  card(s, 8.05, 1.8, 4.8, 5.15, DARK2, { shadow: false });
  s.addText("TALKING POINTS — AFTER", { x: 8.35, y: 2.05, w: 4.2, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  const points = [
    "OrderProcessor now only orchestrates. Pricing, persistence, and notification each live elsewhere.",
    "A new discount type is a new class, not a new elif — process() never changes for a pricing change. That's OCP.",
    "repository and notifier are passed in — swap a fake in for a test without touching OrderProcessor. That's DIP.",
    "Any DiscountStrategy works anywhere one's expected, for free. That's LSP.",
  ];
  points.forEach((p, i) => {
    const y = 2.55 + i * 1.05;
    circle(s, 8.35, y, 0.36, TEAL, "✓", { color: WHITE, fontFace: BODY, fontSize: 14 });
    s.addText(p, { x: 8.85, y: y - 0.08, w: 3.9, h: 0.95, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11, color: "C6DADC" });
  });
  footer(s, "Day 03 · Clean Code to Scale", 24, true);
}

// ============================================================ SLIDE 25 — HANDS-ON OVERVIEW
{
  const s = pptx.addSlide();
  bg(s, PAPER);
  header(s, "2:10–3:05 · Hands-on", "Invert the dependency, then wrap it", false);
  s.addText("Two core katas, in order — the second only works because of what the first one built. Work in pairs.", { x: 0.5, y: 1.42, w: 11.5, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK_SOFT });
  const katas = [
    ["A", TEAL, TEAL_TINT, "InventoryAlert", "35 min · Core", "Invert two constructed dependencies — a notifier and a clock — into constructor parameters."],
    ["B", PLUM, PLUM_TINT, "RepositoryDecorators", "20 min · Core", "Wrap a repository with logging and retry decorators, using the exact seam DIP just built."],
    ["C", RUST, RUST_TINT, "SubscriptionManager", "~15 min · Stretch", "Split a four-responsibility renew() method — for pairs who finish both core katas early."],
  ];
  const cw = 3.95, gx = 0.2, ox = 0.5, oy = 2.1, ch = 4.4;
  katas.forEach((k, i) => {
    const x = ox + i * (cw + gx);
    card(s, x, oy, cw, ch, WHITE);
    circle(s, x + 0.35, oy + 0.35, 0.7, k[1], "Kata " + k[0], { fontSize: 12 });
    s.addText(k[4], { x: x + 0.35, y: oy + 1.25, w: cw - 0.7, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: k[1], charSpacing: 0.5 });
    s.addText(k[3], { x: x + 0.35, y: oy + 1.6, w: cw - 0.7, h: 0.7, isTextBox: true, margin: 0, fontFace: CODE, bold: true, fontSize: 17, color: INK });
    s.addText(k[5], { x: x + 0.35, y: oy + 2.4, w: cw - 0.7, h: 1.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: INK_SOFT });
  });
  footer(s, "Day 03 · Clean Code to Scale", 25, false);
}

// ============================================================ SLIDE 26 — KATA A: INVENTORYALERT
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Kata A · Core · 35 min", "InventoryAlert — invert two dependencies", true, { titleSize: 25 });
  s.addText("Change InventoryAlert to receive its notifier and clock, instead of building them itself", { x: 0.5, y: 1.3, w: 11, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11.5, color: "6C8B90" });

  codePanel(s, 0.6, 1.6, 12.1, 2.2);
  const inventoryAlert = [
    'class InventoryAlert:',
    '    def __init__(self, threshold):',
    '        self.threshold = threshold',
    '        self.notifier = SlackNotifier(webhook_url="https://hooks.slack.com/T000/B000/xxxx")',
    '        self.clock = SystemClock()',
    '',
    '    def check(self, item):',
    '        if item.stock < self.threshold:',
    '            timestamp = self.clock.now()',
    '            self.notifier.send(f"[{timestamp}] Low stock: {item.name} ({item.stock} left)")',
  ].join("\n");
  s.addText(inventoryAlert, { x: 0.9, y: 1.72, w: 11.5, h: 1.95, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11.5, color: CODE_INK, lineSpacingMultiple: 0.95, valign: "middle" });

  s.addText("GUIDING QUESTIONS", { x: 0.6, y: 4.15, w: 11.5, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  const qs = [
    "What would you have to do today to unit test check() without pinging Slack?",
    "What does InventoryAlert actually need from a notifier — the smallest interface — versus what SlackNotifier offers?",
    "If tomorrow you needed email instead, how many lines would you touch today — and after your fix?",
  ];
  qs.forEach((q, i) => {
    const y = 4.55 + i * 0.62;
    circle(s, 0.9, y + 0.02, 0.3, AMBER_LIGHT, "?", { color: DARK, fontFace: BODY, fontSize: 13 });
    s.addText(q, { x: 1.35, y: y - 0.1, w: 11.05, h: 0.55, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: "A9C2C5" });
  });
  s.addShape(pptx.ShapeType.line, { x: 0.9, y: 6.6, w: 11.5, h: 0, line: { color: "234750", width: 1 } });
  s.addText("Done when: __init__ takes notifier and clock as parameters, not constructs them.", { x: 0.9, y: 6.68, w: 11.5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11, color: AMBER_LIGHT });
  footer(s, "Day 03 · Clean Code to Scale", 26, true);
}

// ============================================================ SLIDE 27 — KATA A: INVENTORYALERT (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Kata A · Core · 35 min", "InventoryAlert — original vs. solution", true, { titleSize: 25 });

  const original = [
    { t: "class InventoryAlert:", c: CODE_INK },
    { t: "    def __init__(self, threshold):", c: CODE_INK },
    { t: "        self.threshold = threshold", c: CODE_INK },
    { t: '        self.notifier = SlackNotifier(webhook_url="...")', c: CODE_INK },
    { t: "        self.clock = SystemClock()", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def check(self, item):", c: CODE_INK },
    { t: "        if item.stock < self.threshold:", c: CODE_INK },
    { t: "            timestamp = self.clock.now()", c: CODE_INK },
    { t: "            self.notifier.send(", c: CODE_INK },
    { t: '                f"[{timestamp}] Low stock: {item.name}")', c: CODE_INK },
  ];
  const solution = [
    { t: "class InventoryAlert:", c: CODE_INK },
    { t: "    def __init__(self, threshold, notifier, clock):", c: CODE_INK },
    { t: "        self.threshold = threshold", c: CODE_INK },
    { t: "        self.notifier = notifier", c: CODE_INK },
    { t: "        self.clock = clock", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def check(self, item):", c: CODE_INK },
    { t: "        if item.stock < self.threshold:", c: CODE_INK },
    { t: "            timestamp = self.clock.now()", c: CODE_INK },
    { t: "            self.notifier.send(", c: CODE_INK },
    { t: '                f"[{timestamp}] Low stock: {item.name}")', c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# production", c: CODE_COMMENT },
    { t: 'alert = InventoryAlert(10, SlackNotifier(webhook_url="..."), SystemClock())', c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# test — no network, no real clock", c: CODE_COMMENT },
    { t: 'alert = InventoryAlert(10, FakeNotifier(), FixedClock("2026-01-01"))', c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 6.1, 5.3, "The original", AMBER_LIGHT, original, 10.5);
  labeledCode(s, 6.75, 1.5, 6.1, 5.3, "The solution", TEAL_TINT, solution, 9);
  footer(s, "Day 03 · Clean Code to Scale", 27, true);
}

// ============================================================ SLIDE 28 — KATA B: REPOSITORYDECORATORS
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Kata B · Core · 20 min", "RepositoryDecorators — wrap it, don't edit it", true, { titleSize: 24 });

  codePanel(s, 0.6, 1.55, 12.1, 2.75);
  const repo = [
    'class PostgresRepository:',
    '    def save(self, order_id, total):',
    '        start = time.time()',
    '        for attempt in range(3):',
    '            try:',
    '                db.execute(',
    '                    "INSERT INTO orders (id, total) VALUES (%s, %s)",',
    '                    order_id, total,',
    '                )',
    '                log.info(f"saved order {order_id} in {time.time() - start:.2f}s")',
    '                return',
    '            except ConnectionError:',
    '                if attempt == 2:',
    '                    raise',
    '                time.sleep(0.5)',
  ].join("\n");
  s.addText(repo, { x: 0.9, y: 1.67, w: 11.5, h: 2.5, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 10.5, color: CODE_INK, lineSpacingMultiple: 0.9, valign: "middle" });

  card(s, 0.6, 4.45, 12.1, 0.6, DARK2, { shadow: false });
  s.addText([
    { text: "Why this only works because of the last kata:  ", options: { fontFace: BODY, bold: true, fontSize: 11, color: AMBER_LIGHT, breakLine: false } },
    { text: "OrderProcessor.process() already receives repository instead of constructing it — Decorator needs exactly that seam.", options: { fontFace: BODY, fontSize: 11, color: "C6DADC" } },
  ], { x: 0.85, y: 4.45, w: 11.8, h: 0.6, isTextBox: true, margin: 0, valign: "middle" });

  const qs = [
    "What does save() have to do with logging or retries, conceptually? Name its one real job.",
    "How many lines of OrderProcessor change if it receives a decorated repository instead of a plain one?",
    "Which decorator would you want to skip in a test — and can you, without touching PostgresRepository?",
  ];
  qs.forEach((q, i) => {
    const y = 5.2 + i * 0.48;
    circle(s, 0.9, y + 0.02, 0.26, AMBER_LIGHT, "?", { color: DARK, fontFace: BODY, fontSize: 11 });
    s.addText(q, { x: 1.3, y: y - 0.08, w: 11.15, h: 0.42, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11, color: "A9C2C5" });
  });
  s.addText("Done when: save() only persists, logging and retry are separate wrapper classes, and they compose in either order.", { x: 0.9, y: 6.7, w: 11.5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 10.5, color: AMBER_LIGHT });
  footer(s, "Day 03 · Clean Code to Scale", 28, true);
}

// ============================================================ SLIDE 29 — KATA B: REPOSITORYDECORATORS (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Kata B · Core · 20 min", "RepositoryDecorators — original vs. solution", true, { titleSize: 23 });

  const original = [
    { t: "class PostgresRepository:", c: CODE_INK },
    { t: "    def save(self, order_id, total):", c: CODE_INK },
    { t: "        start = time.time()", c: CODE_INK },
    { t: "        for attempt in range(3):", c: CODE_INK },
    { t: "            try:", c: CODE_INK },
    { t: "                db.execute(SQL, order_id, total)", c: CODE_INK },
    { t: '                log.info("saved order")', c: CODE_INK },
    { t: "                return", c: CODE_INK },
    { t: "            except ConnectionError:", c: CODE_INK },
    { t: "                if attempt == 2: raise", c: CODE_INK },
    { t: "                time.sleep(0.5)", c: CODE_INK },
  ];
  const solution = [
    { t: "class PostgresRepository:", c: CODE_INK },
    { t: "    def save(self, order_id, total):", c: CODE_INK },
    { t: "        db.execute(SQL, order_id, total)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class LoggingRepository:", c: CODE_INK },
    { t: "    def __init__(self, repository):", c: CODE_INK },
    { t: "        self.repository = repository", c: CODE_INK },
    { t: "    def save(self, order_id, total):", c: CODE_INK },
    { t: "        start = time.time()", c: CODE_INK },
    { t: "        self.repository.save(order_id, total)", c: CODE_INK },
    { t: '        log.info(f"saved in {time.time()-start:.2f}s")', c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class RetryingRepository:", c: CODE_INK },
    { t: "    def __init__(self, repository, attempts=3):", c: CODE_INK },
    { t: "        self.repository = repository", c: CODE_INK },
    { t: "        self.attempts = attempts", c: CODE_INK },
    { t: "    def save(self, order_id, total):", c: CODE_INK },
    { t: "        for attempt in range(self.attempts):", c: CODE_INK },
    { t: "            try:", c: CODE_INK },
    { t: "                return self.repository.save(order_id, total)", c: CODE_INK },
    { t: "            except ConnectionError:", c: CODE_INK },
    { t: "                if attempt == self.attempts - 1: raise", c: CODE_INK },
    { t: "                time.sleep(0.5)", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# compose in either order — neither class knows about the other", c: CODE_COMMENT },
    { t: "repository = LoggingRepository(RetryingRepository(PostgresRepository()))", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 4.3, 5.3, "The original", AMBER_LIGHT, original, 9.5);
  labeledCode(s, 5.0, 1.5, 7.85, 5.3, "The solution", TEAL_TINT, solution, 8.3);
  footer(s, "Day 03 · Clean Code to Scale", 29, true);
}

// ============================================================ SLIDE 30 — KATA C: SUBSCRIPTIONMANAGER
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Kata C · Stretch · ~15 min", "SubscriptionManager — reduce the responsibilities", true, { titleSize: 23 });
  s.addText("Finished both core katas early? renew() validates, prices, charges, updates state, and emails — split it.", { x: 0.5, y: 1.3, w: 11.5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 11.5, color: "6C8B90" });

  codePanel(s, 0.6, 1.7, 12.1, 3.1);
  const sub = [
    'class SubscriptionManager:',
    '    def __init__(self, payment_gateway, email_service):',
    '        self.payment_gateway = payment_gateway',
    '        self.email_service = email_service',
    '',
    '    def renew(self, subscription, plan):',
    '        if plan not in ("monthly", "annual"):',
    '            raise ValueError("unknown plan")',
    '        price = 29.00 if plan == "monthly" else 290.00',
    '        self.payment_gateway.charge(subscription.customer_id, price)',
    '        subscription.plan = plan',
    '        subscription.renewed_at = datetime.now()',
    '        self.email_service.send(',
    '            subscription.customer_email,',
    '            f"Your {plan} subscription has been renewed for ${price:.2f}"',
    '        )',
  ].join("\n");
  s.addText(sub, { x: 0.9, y: 1.85, w: 11.5, h: 2.8, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11.5, color: CODE_INK, lineSpacingMultiple: 0.95 });

  const qs = [
    "How many reasons does renew() have to change today? Name each one.",
    "Which extracted piece would you want unit tested directly, without a real payment gateway?",
    "If a third plan type is added tomorrow, how many places need to change?",
  ];
  qs.forEach((q, i) => {
    const y = 5.15 + i * 0.55;
    circle(s, 0.9, y + 0.02, 0.28, AMBER_LIGHT, "?", { color: DARK, fontFace: BODY, fontSize: 12 });
    s.addText(q, { x: 1.32, y: y - 0.1, w: 11.1, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: "A9C2C5" });
  });
  s.addText("Done when: renew() reads as a short ordered list of calls, and each extracted piece has exactly one reason to change.", { x: 0.9, y: 6.95, w: 11.5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, italic: true, fontSize: 10.5, color: AMBER_LIGHT });
  footer(s, "Day 03 · Clean Code to Scale", 30, true);
}

// ============================================================ SLIDE 31 — KATA C: SUBSCRIPTIONMANAGER (solution)
{
  const s = pptx.addSlide();
  bg(s, DARK);
  header(s, "Kata C · Stretch · ~15 min", "SubscriptionManager — original vs. solution", true, { titleSize: 22 });

  const original = [
    { t: "class SubscriptionManager:", c: CODE_INK },
    { t: "    def __init__(self, payment_gateway, email_service):", c: CODE_INK },
    { t: "        self.payment_gateway = payment_gateway", c: CODE_INK },
    { t: "        self.email_service = email_service", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def renew(self, subscription, plan):", c: CODE_INK },
    { t: '        if plan not in ("monthly", "annual"):', c: CODE_INK },
    { t: '            raise ValueError("unknown plan")', c: CODE_INK },
    { t: '        price = 29.00 if plan == "monthly" else 290.00', c: CODE_INK },
    { t: "        self.payment_gateway.charge(", c: CODE_INK },
    { t: "            subscription.customer_id, price)", c: CODE_INK },
    { t: "        subscription.plan = plan", c: CODE_INK },
    { t: "        subscription.renewed_at = datetime.now()", c: CODE_INK },
    { t: "        self.email_service.send(", c: CODE_INK },
    { t: "            subscription.customer_email,", c: CODE_INK },
    { t: '            f"Renewed: {plan} plan, ${price:.2f}")', c: CODE_INK },
  ];
  const solution = [
    { t: "class SubscriptionValidator:", c: CODE_INK },
    { t: '    VALID_PLANS = ("monthly", "annual")', c: CODE_INK },
    { t: "    def validate(self, plan):", c: CODE_INK },
    { t: "        if plan not in self.VALID_PLANS:", c: CODE_INK },
    { t: '            raise ValueError("unknown plan")', c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "class SubscriptionPricing:", c: CODE_INK },
    { t: '    PRICES = {"monthly": 29.00, "annual": 290.00}', c: CODE_INK },
    { t: "    def price_for(self, plan):", c: CODE_INK },
    { t: "        return self.PRICES[plan]", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "# subscription.mark_renewed() moves the mutation onto the entity", c: CODE_COMMENT },
    { t: "class SubscriptionManager:", c: CODE_INK },
    { t: "    def __init__(self, payment_gateway, email_service,", c: CODE_INK },
    { t: "                 validator, pricing):", c: CODE_INK },
    { t: "        self.payment_gateway = payment_gateway", c: CODE_INK },
    { t: "        self.email_service = email_service", c: CODE_INK },
    { t: "        self.validator = validator", c: CODE_INK },
    { t: "        self.pricing = pricing", c: CODE_INK },
    { t: "", c: CODE_INK },
    { t: "    def renew(self, subscription, plan):", c: CODE_INK },
    { t: "        self.validator.validate(plan)", c: CODE_INK },
    { t: "        price = self.pricing.price_for(plan)", c: CODE_INK },
    { t: "        self.payment_gateway.charge(", c: CODE_INK },
    { t: "            subscription.customer_id, price)", c: CODE_INK },
    { t: "        subscription.mark_renewed(plan)", c: CODE_INK },
    { t: "        self.email_service.send_renewal_notice(", c: CODE_INK },
    { t: "            subscription, plan, price)", c: CODE_INK },
  ];
  labeledCode(s, 0.5, 1.5, 5.5, 5.3, "The original", AMBER_LIGHT, original, 9.5);
  labeledCode(s, 6.15, 1.5, 6.7, 5.3, "The solution", TEAL_TINT, solution, 8.3);
  footer(s, "Day 03 · Clean Code to Scale", 31, true);
}

// ============================================================ SLIDE 32 — REVIEW / DEBRIEF
{
  const s = pptx.addSlide();
  bg(s, PAPER);
  header(s, "3:05–3:20 · Review", "When is SOLID too much?", false);
  const qs = [
    ["JR", TEAL, "Everyone", "Where have you seen SOLID applied so eagerly that a 10-line script became five files? What was actually gained?"],
    ["SR", "5C7A85", "Tech Leads", "What's your team's real signal for \"this needs an interface\" — is it written down, or does it live in one person's head?"],
    ["QE", RUST, "QEs", "Does an interface you can't easily fake in a test actually satisfy DIP, or does it just look like it does?"],
    ["AI", PLUM, "AI Engineers", "An agent told to \"make this more SOLID\" will often add interfaces everywhere. When is that right, and when is it over-engineering?"],
  ];
  const cw = 5.85, ch = 2.05, gx = 0.3, gy = 0.3, ox = 0.5, oy = 1.85;
  qs.forEach((q, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = ox + col * (cw + gx), y = oy + row * (ch + gy);
    card(s, x, y, cw, ch, WHITE);
    circle(s, x + 0.3, y + 0.3, 0.55, q[1], q[0], { fontSize: 12 });
    s.addText(q[2], { x: x + 1.05, y: y + 0.34, w: cw - 1.3, h: 0.4, isTextBox: true, margin: 0, valign: "middle", fontFace: CODE, bold: true, fontSize: 12, color: q[1], charSpacing: 0.5 });
    s.addText(q[3], { x: x + 0.3, y: y + 1.0, w: cw - 0.6, h: 0.95, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: INK });
  });
  footer(s, "Day 03 · Clean Code to Scale", 32, false);
}

// ============================================================ SLIDE 33 — CLOSING
{
  const s = pptx.addSlide();
  bg(s, DARK);
  s.addText("BEFORE YOU GO", { x: 0.7, y: 0.55, w: 8, h: 0.35, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 13, color: AMBER_LIGHT, charSpacing: 1.5 });
  s.addText("Pick one. Do it this week.", { x: 0.7, y: 0.9, w: 11, h: 0.8, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 34, color: WHITE });

  const items = [
    "Find one class with more than one reason to change, and name both reasons",
    "Find one place a concrete dependency is constructed inside a class instead of passed in",
    "Delete one interface with exactly one implementation that will likely only ever have one",
    "Bring one real \"is this over-engineered?\" example to raise with your team this week",
  ];
  items.forEach((it, i) => {
    const y = 2.05 + i * 0.8;
    card(s, 0.7, y, 11.9, 0.62, DARK2, { shadow: false });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: y + 0.16, w: 0.3, h: 0.3, rectRadius: 0.04, fill: { type: "none" }, line: { color: TEAL, width: 1.5 } });
    s.addText(it, { x: 1.45, y, w: 10.9, h: 0.62, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 14.5, color: WHITE });
  });

  s.addShape(pptx.ShapeType.line, { x: 0.7, y: 5.7, w: 11.9, h: 0, line: { color: "234750", width: 1 } });
  s.addText("NEXT — DAY 04", { x: 0.7, y: 5.95, w: 6, h: 0.3, isTextBox: true, margin: 0, fontFace: CODE, fontSize: 11, color: AMBER_LIGHT, charSpacing: 1 });
  s.addText("Refactoring Technique & the Smell Catalog", { x: 0.7, y: 6.25, w: 10, h: 0.5, isTextBox: true, margin: 0, fontFace: HEAD, bold: true, fontSize: 22, color: WHITE });
  footer(s, "Clean Code to Scale", 33, true);
}

pptx.writeFile({ fileName: "session-03.pptx" }).then(() => {
  console.log("written");
});
