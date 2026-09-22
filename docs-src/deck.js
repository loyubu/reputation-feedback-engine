const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";              // 13.3 x 7.5 — set BEFORE adding slides
pres.author = "Capstone group";
pres.title  = "Reputation & Feedback Intelligence Engine";

// ---- palette: the routing logic IS the palette -------------------------------
const INK = "14171A", GRAPH = "23272E", BODY = "2B3038", MUTED = "5E6874";
const SURF = "F5F6F8", WHITE = "FFFFFF", RULE = "DCE0E6";
const RA = "1C7A4D", RB = "A96410", RC = "B02A21", RD = "2F5D8C";
const DIM = "9BA5B2", PALE = "C6CDD6";
const HEAD = "Cambria", SANS = "Calibri";

const ROUTES = [
  { c: RA, k: "ROUTE A", n: "Positive" },
  { c: RB, k: "ROUTE B", n: "Negative · mild · first time" },
  { c: RC, k: "ROUTE C", n: "Negative · severe or repeat" },
  { c: RD, k: "ROUTE D", n: "Cannot score confidently" },
];

// ---- helpers (fresh option objects every call — pptxgenjs mutates them) ------
const title = (s, t, dark) => s.addText(t, {
  x: 0.55, y: 0.38, w: 12.2, h: 0.72, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 30, bold: true, color: dark ? WHITE : INK,
});
const kicker = (s, t, dark) => s.addText(t, {
  x: 0.57, y: 1.12, w: 12.2, h: 0.32, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12.5, color: dark ? DIM : MUTED, charSpacing: 0.6,
});
// the repeating motif: four route chips
const chips = (s, y) => ROUTES.forEach((r, i) =>
  s.addShape(pres.ShapeType.rect, {
    x: 0.57 + i * 0.42, y, w: 0.3, h: 0.3, fill: { color: r.c },
  }));
const card = (s, o) => {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.04,
    fill: { color: o.fill || WHITE },
    line: { color: o.stroke || RULE, width: o.lw || 1.25 },
  });
};
const arrow = (s, o) => s.addShape(pres.ShapeType.line, {
  x: o.x, y: o.y, w: o.w || 0, h: o.h || 0,
  line: { color: o.c || GRAPH, width: o.w2 || 1.75, endArrowType: "triangle",
          dashType: o.dash || "solid" },
});

// =============================================================== 1. TITLE ====
let s = pres.addSlide();
s.background = { color: GRAPH };
s.addText("CAPSTONE  ·  DESIGN REVIEW  ·  NOT YET APPROVED", {
  x: 0.57, y: 1.5, w: 11, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12, color: DIM, charSpacing: 2,
});
s.addText("Reputation & Feedback\nIntelligence Engine", {
  x: 0.55, y: 2.0, w: 11.5, h: 1.9, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 46, bold: true, color: WHITE, lineSpacing: 52,
});
s.addText("How a completed visit becomes scored feedback, routed four ways, with a draft reply waiting and a manager alerted when it matters.", {
  x: 0.57, y: 4.1, w: 9.2, h: 0.8, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 15, color: PALE, lineSpacing: 23,
});
chips(s, 5.25);
s.addText("Four routes · one dashboard · nothing published or sent without a human", {
  x: 0.57, y: 5.72, w: 10, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12, color: DIM,
});
s.addText("22 September 2026     ·     Built on n8n     ·     Group work     ·     Source: Capstone Project Information.txt", {
  x: 0.57, y: 6.6, w: 12, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 11.5, color: MUTED,
});
s.addNotes("Design review, not an approved build. Four questions still go to the group — they are on the last slide.");

// ============================================================ 2. THE PROBLEM ==
s = pres.addSlide();
title(s, "Two problems, one loop");
kicker(s, "What the business actually said");
const probs = [
  { h: "Bad reviews arrive before we do", b: "Follow-up after a visit depends on whoever is at the front desk that day. The first anyone hears of a problem is a public review — by which point it is already public." },
  { h: "Every complaint treated the same", b: "A mild grumble about a delay and a furious complaint about a botched job go into the same pile. Somebody reads them eventually and decides what to do." },
];
probs.forEach((p, i) => {
  const x = 0.55 + i * 6.2;
  card(s, { x, y: 1.75, w: 5.85, h: 2.15, fill: SURF, stroke: RULE });
  s.addText(p.h, { x: x + 0.32, y: 1.98, w: 5.2, h: 0.4, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 17, bold: true, color: INK });
  s.addText(p.b, { x: x + 0.32, y: 2.48, w: 5.25, h: 1.3, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 13, color: BODY, lineSpacing: 20 });
});
card(s, { x: 0.55, y: 4.3, w: 12.05, h: 1.85, fill: GRAPH, stroke: GRAPH });
s.addText("The standing rule", { x: 0.95, y: 4.58, w: 6, h: 0.35, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 17, bold: true, color: WHITE });
s.addText("Nothing is ever published publicly by the system, and nothing is ever sent to a customer by the system. A human does both. The system's job is to notice, score, route and prepare — never to act on the customer's behalf.", {
  x: 0.95, y: 5.04, w: 11.2, h: 0.95, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 13.5, color: PALE, lineSpacing: 21 });
s.addText("This is the constraint the whole design is built around.", {
  x: 0.55, y: 6.45, w: 8, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12, italic: true, color: MUTED });
s.addNotes("The brief is emphatic on this: every review goes through a human first, and no reply is auto-sent.");

// ========================================================= 3. EASY TO MISS ====
s = pres.addSlide();
title(s, "Three things in the brief that are easy to miss");
kicker(s, "Each one changes what gets built");
const miss = [
  { n: "01", h: "Four routes, not three", b: "The flow lists three. Point 8 adds a fourth: anything the system cannot confidently score is flagged for a human rather than guessed at.", c: RD },
  { n: "02", h: "Two numbers, not one", b: "“How positive or negative it is, and how severe.” Different axes. Collapsing them is exactly the flattening the brief objects to.", c: RB },
  { n: "03", h: "History is an override", b: "“Regardless of how mild this message reads.” History does not nudge the score — it bypasses the severity threshold entirely.", c: RC },
];
miss.forEach((m, i) => {
  const x = 0.55 + i * 4.12;
  card(s, { x, y: 1.8, w: 3.85, h: 3.5, stroke: RULE });
  s.addText(m.n, { x: x + 0.3, y: 2.02, w: 1.5, h: 0.65, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 34, bold: true, color: m.c });
  s.addText(m.h, { x: x + 0.3, y: 2.82, w: 3.3, h: 0.72, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 17, bold: true, color: INK, lineSpacing: 22 });
  s.addText(m.b, { x: x + 0.3, y: 3.62, w: 3.3, h: 1.5, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 12.5, color: BODY, lineSpacing: 19 });
});
s.addText("Get the third one wrong and the repeat-customer demo case stops looking different from the ordinary negative case.", {
  x: 0.55, y: 5.6, w: 12, h: 0.4, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 13, color: MUTED });
s.addNotes("The fourth route and the override are the two most commonly missed requirements in this brief.");

// =========================================================== 4. WORKFLOW 1 ====
s = pres.addSlide();
title(s, "Workflow 1 — a completed visit asks for feedback");
kicker(s, "Deliberately simple. The brief asks only for a reliable trigger.");
const w1 = [
  { k: "TRIGGER", h: "Visit Complete", b: "Webhook standing in for\nthe shop's job system" },
  { k: "STEP 2", h: "Upsert Customer", b: "Match on email →\nstable identity" },
  { k: "STEP 3", h: "Create Visit", b: "visit_id · location_id\ncompleted_at" },
  { k: "STEP 4", h: "Send Request", b: "Link built from the\nConfig base URL" },
  { k: "STEP 5", h: "Stamp sent_at", b: "Dashboard shows\n“sent, awaiting reply”" },
];
w1.forEach((n, i) => {
  const x = 0.55 + i * 2.51;
  card(s, { x, y: 2.15, w: 2.2, h: 1.95, fill: i === 0 ? GRAPH : WHITE,
            stroke: i === 0 ? GRAPH : RULE, lw: i === 1 ? 2 : 1.25 });
  s.addText(n.k, { x: x + 0.2, y: 2.32, w: 1.8, h: 0.25, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 9.5, bold: true, color: i === 0 ? DIM : MUTED, charSpacing: 1 });
  s.addText(n.h, { x: x + 0.2, y: 2.63, w: 1.9, h: 0.35, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 14.5, bold: true, color: i === 0 ? WHITE : INK });
  s.addText(n.b, { x: x + 0.2, y: 3.08, w: 1.9, h: 0.85, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 11, color: i === 0 ? PALE : BODY, lineSpacing: 16 });
  if (i < 4) arrow(s, { x: x + 2.24, y: 3.12, w: 0.23, h: 0 });
});
card(s, { x: 2.45, y: 4.62, w: 8.2, h: 1.25, fill: SURF, stroke: RB, lw: 2 });
s.addText("Why the customer record comes before the visit record", {
  x: 2.75, y: 4.8, w: 7.7, h: 0.32, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 15, bold: true, color: INK });
s.addText("The repeat-customer check later only works if the same person is recognisable across two visits. Get this wrong and Route C never fires, however correct the routing logic is.", {
  x: 2.75, y: 5.16, w: 7.7, h: 0.6, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12.5, color: BODY, lineSpacing: 19 });
s.addNotes("Identity is established on the way in. It is the hinge the whole repeat-customer rule hangs on.");

// ====================================================== 5. WORKFLOW 2 SPINE ===
s = pres.addSlide();
title(s, "Workflow 2 — the reply gets scored");
kicker(s, "Every threshold lives in one Config node, so the rules are readable at a glance");
const w2 = [
  { h: "Feedback Reply", b: "n8n Form · free text\nvisit_id in the URL", dark: true },
  { h: "Load Context", b: "visit + customer + location\ndedupe if already scored" },
  { h: "Score the reply", b: "polarity −1…+1\nseverity 0…5 · confidence", model: true },
  { h: "Confidence gate", b: "is this score\ntrustworthy?" },
  { h: "History Check", b: "prior negative for this\ncustomer? a DB query" },
  { h: "Decide", b: "one Code node\noutputs: route", dark: true },
];
w2.forEach((n, i) => {
  const x = 0.55 + i * 2.09;
  card(s, { x, y: 2.1, w: 1.82, h: 1.85,
    fill: n.dark ? GRAPH : WHITE, stroke: n.dark ? GRAPH : (n.model ? INK : RULE),
    lw: n.model ? 2.5 : 1.25 });
  if (n.model) s.addText("MODEL CALL 1", { x: x + 0.16, y: 2.26, w: 1.6, h: 0.22,
    isTextBox: true, margin: 0, fontFace: SANS, fontSize: 8.5, bold: true, color: MUTED, charSpacing: 0.8 });
  s.addText(n.h, { x: x + 0.16, y: n.model ? 2.54 : 2.34, w: 1.6, h: 0.5, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 13.5, bold: true, color: n.dark ? WHITE : INK, lineSpacing: 16 });
  s.addText(n.b, { x: x + 0.16, y: n.model ? 3.06 : 2.94, w: 1.62, h: 0.8, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 10.5, color: n.dark ? PALE : BODY, lineSpacing: 15 });
  if (i < 5) arrow(s, { x: x + 1.86, y: 3.0, w: 0.19, h: 0 });
});
// bypass
arrow(s, { x: 5.64, y: 3.95, w: 0, h: 1.0, c: RD, dash: "dash", w2: 1.6 });
arrow(s, { x: 7.73, y: 3.95, w: 0, h: 1.0, c: RD, dash: "dash", w2: 1.6 });
card(s, { x: 2.6, y: 4.98, w: 6.3, h: 1.15, fill: WHITE, stroke: RD, lw: 2 });
s.addText("The bypass — Route D", { x: 2.9, y: 5.14, w: 5.7, h: 0.3, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 14.5, bold: true, color: RD });
s.addText("Model failed, malformed output, below the confidence floor, or not really feedback → straight to a human. Nothing guessed, nothing routed.", {
  x: 2.9, y: 5.48, w: 5.8, h: 0.55, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 11.5, color: BODY, lineSpacing: 17 });
s.addText("One Code node decides the route; a single Switch acts on it — the whole decision in one readable place.", {
  x: 0.55, y: 6.45, w: 12, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12, italic: true, color: MUTED });
s.addNotes("A Basic LLM Chain with a structured output parser — not an Agent. One classification call; an Agent would add nondeterminism to a demo that must behave identically three times.");

// ========================================================= 6. FOUR ROUTES =====
s = pres.addSlide();
title(s, "Four routes, four different outcomes");
kicker(s, "Routing depends on the score and the customer's history — never on a keyword");
const steps = [
  [["Ready-to-post queue", "stored for a human to open,\nread and copy out later"]],
  [["Private queue", "the team handles it in\ntheir own time"], ["Draft response", "model call 2 · status pending"]],
  [["Private queue", "never pointed toward a\npublic review"], ["Draft response", "generated in parallel"], ["Alert a manager — now", "open alert logged"]],
  [["Review queue", "flagged for a human —\nnothing guessed"]],
];
ROUTES.forEach((r, i) => {
  const x = 0.42 + i * 3.16;
  card(s, { x, y: 1.85, w: 2.95, h: 0.92, fill: r.c, stroke: r.c });
  s.addText(r.k, { x: x + 0.22, y: 1.99, w: 2.5, h: 0.24, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 9.5, bold: true, color: WHITE, charSpacing: 1 });
  s.addText(r.n, { x: x + 0.22, y: 2.26, w: 2.6, h: 0.42, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 13.5, bold: true, color: WHITE, lineSpacing: 16 });
  steps[i].forEach((st, j) => {
    const y = 3.02 + j * 1.22;
    card(s, { x, y, w: 2.95, h: 1.0, fill: st[0].startsWith("Alert") ? r.c : WHITE,
              stroke: r.c, lw: 1.5 });
    const fg = st[0].startsWith("Alert");
    s.addText(st[0], { x: x + 0.22, y: y + 0.14, w: 2.6, h: 0.3, isTextBox: true, margin: 0,
      fontFace: HEAD, fontSize: 12.5, bold: true, color: fg ? WHITE : INK });
    s.addText(st[1], { x: x + 0.22, y: y + 0.47, w: 2.6, h: 0.5, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 10.5, color: fg ? PALE : BODY, lineSpacing: 15 });
    arrow(s, { x: x + 1.47, y: y - 0.22, w: 0, h: 0.18, c: r.c, w2: 1.6 });
  });
});
s.addText("Route A is never published automatically. Routes B and C are never sent automatically. Both wait for a human.", {
  x: 0.42, y: 6.6, w: 12.2, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12, color: MUTED });
s.addNotes("Show all four populated on the dashboard during the demo — that is what proves the paths behave differently.");

// ======================================================== 7. DECISION LOGIC ===
s = pres.addSlide();
title(s, "The decision logic — one Code node");
kicker(s, "Order matters. First match wins.");
s.addTable([
  [ { text: "#",         options: { bold: true, color: INK, fill: { color: SURF }, fontSize: 11 } },
    { text: "CONDITION", options: { bold: true, color: INK, fill: { color: SURF }, fontSize: 11 } },
    { text: "ROUTE",     options: { bold: true, color: INK, fill: { color: SURF }, fontSize: 11 } },
    { text: "WHY IT SITS HERE", options: { bold: true, color: INK, fill: { color: SURF }, fontSize: 11 } } ],
  ["1", "Model failed, or confidence < 0.6, or not really feedback",
       { text: "D — review",        options: { bold: true, color: RD } },
       "Checked first, so a broken score can never reach the rules below"],
  ["2", "polarity < 0  AND  customer has a prior negative",
       { text: "C — escalate",      options: { bold: true, color: RC } },
       "Above the severity test deliberately — this is the override"],
  ["3", "polarity < 0  AND  severity ≥ 4",
       { text: "C — escalate",      options: { bold: true, color: RC } },
       "Severity alone is enough, no history required"],
  ["4", "polarity < 0",
       { text: "B — private queue", options: { bold: true, color: RB } },
       "First time, lower severity. Queued; a draft is still generated"],
  ["5", "polarity > 0",
       { text: "A — ready to post", options: { bold: true, color: RA } },
       "Stored for a human to read and copy out"],
  ["6", "Anything left — neutral",
       { text: "D — review",        options: { bold: true, color: RD } },
       "Neutral is not positive, and Route A feeds something a human may publish"],
], {
  x: 0.55, y: 1.8, w: 12.2, colW: [0.5, 4.2, 1.9, 5.6],
  fontFace: SANS, fontSize: 11.5, color: BODY, valign: "middle",
  border: { type: "solid", color: RULE, pt: 0.75 }, rowH: 0.52, margin: 6,
});
s.addText("Rule 2 sits above rule 3 on purpose. Flip them and a mild repeat complaint gets severity-tested first, lands in the ordinary queue, and the third demo case stops behaving differently.", {
  x: 0.55, y: 5.85, w: 12.2, h: 0.6, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 13, color: INK, lineSpacing: 20 });
s.addNotes("This is the slide to linger on. It is the part a reviewer will ask about.");

// ====================================================== 8. CHALLENGE POINTS ===
s = pres.addSlide();
title(s, "Two calls worth challenging now");
kicker(s, "Both are judgement calls, not requirements — easier to change before we build");
const ch = [
  { c: RC, h: "Rule 2 above rule 3", b: "The brief says escalate a repeat complainer “regardless of how mild this particular message reads”. Taken literally, history has to be tested before severity.",
    q: "Agreed, or should severity win first?" },
  { c: RD, h: "Neutral goes to review", b: "“The car's ready then?” is not a glowing review, and Route A feeds something a human may publish. So neutral is sent for a human look rather than into the positive queue.",
    q: "Too cautious, or right?" },
];
ch.forEach((m, i) => {
  const x = 0.55 + i * 6.2;
  card(s, { x, y: 1.85, w: 5.85, h: 3.5, stroke: m.c, lw: 2 });
  s.addText(m.h, { x: x + 0.35, y: 2.12, w: 5.1, h: 0.4, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 19, bold: true, color: m.c });
  s.addText(m.b, { x: x + 0.35, y: 2.68, w: 5.15, h: 1.5, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 13, color: BODY, lineSpacing: 20 });
  s.addText(m.q, { x: x + 0.35, y: 4.55, w: 5.15, h: 0.4, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 14, bold: true, italic: true, color: INK });
});
s.addText("Everything else in the design is a direct reading of the brief. These two are ours.", {
  x: 0.55, y: 5.7, w: 12, h: 0.35, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 13, color: MUTED });
s.addNotes("Raise these explicitly with the group. If they disagree, the change is one line in the Code node — but only before it is built.");

// ================================================================= 9. COST ====
s = pres.addSlide();
title(s, "Cost is not what should decide this");
kicker(s, "Prices checked 22 September 2026");
const stats = [
  { n: "~2¢", l: "per piece of feedback\n(two model calls on a negative)", c: RA },
  { n: "$5", l: "for the entire build —\nalso the minimum top-up", c: RA },
  { n: "£0", l: "n8n on Oracle, domain,\ndashboard hosting", c: RA },
  { n: "14", l: "days of n8n Cloud trial —\nthe real constraint", c: RC },
];
stats.forEach((st, i) => {
  const x = 0.55 + i * 3.08;
  card(s, { x, y: 1.85, w: 2.85, h: 2.1, fill: SURF, stroke: i === 3 ? RC : RULE, lw: i === 3 ? 2 : 1.25 });
  s.addText(st.n, { x: x + 0.25, y: 2.05, w: 2.4, h: 0.8, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 40, bold: true, color: st.c });
  s.addText(st.l, { x: x + 0.25, y: 2.95, w: 2.45, h: 0.85, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 11.5, color: BODY, lineSpacing: 16 });
});
card(s, { x: 0.55, y: 4.25, w: 5.85, h: 1.95, fill: WHITE, stroke: RC, lw: 2 });
s.addText("The deadline, not the bill", { x: 0.9, y: 4.45, w: 5.2, h: 0.32, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 16, bold: true, color: RC });
s.addText("Decide on day one, not day thirteen: finish inside 14 days and pay nothing; overrun and pay €24 for one month; or migrate to Oracle early and absorb that work mid-project.", {
  x: 0.9, y: 4.83, w: 5.3, h: 1.2, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12.5, color: BODY, lineSpacing: 19 });
card(s, { x: 6.75, y: 4.25, w: 5.85, h: 1.95, fill: WHITE, stroke: RULE });
s.addText("Two gotchas that outrank the money", { x: 7.1, y: 4.45, w: 5.2, h: 0.32, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 16, bold: true, color: INK });
s.addText([
  { text: "Supabase free projects pause after a week idle", options: { bullet: true, breakLine: true } },
  { text: "Airtable's free plan allows exactly one interface", options: { bullet: true } },
], { x: 7.1, y: 4.85, w: 5.3, h: 1.1, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12.5, color: BODY, paraSpaceAfter: 8 });
s.addText("About $5 of Claude credit covers development, testing and the live demo. Every other component on the chosen path is free.", {
  x: 0.55, y: 6.5, w: 12.2, h: 0.35, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12.5, italic: true, color: MUTED });
s.addNotes("About $5 of Claude credit covers development, testing and the live demo. Everything else is free on the chosen path.");

// ============================================================ 10. QUESTIONS ===
s = pres.addSlide();
s.background = { color: GRAPH };
s.addText("Four questions for the group", {
  x: 0.55, y: 0.75, w: 12, h: 0.8, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 34, bold: true, color: WHITE });
s.addText("Everything else is stable enough to build against. These four are the blockers.", {
  x: 0.57, y: 1.62, w: 11, h: 0.35, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 14, color: DIM });
const qs = [
  ["Datastore + dashboard", "Airtable + Interface Designer  ·  Supabase + custom dashboard  ·  Supabase + hosted page  ·  Sheets + Looker Studio"],
  ["Request and reply channel", "n8n Form link (recommended)  ·  email out and reply in  ·  Telegram bot  ·  simulated webhook"],
  ["Manager alert channel", "Telegram  ·  email  ·  Slack  ·  both"],
  ["“Ready to post” content", "The customer's raw words, or a lightly polished review drafted from them"],
];
qs.forEach((q, i) => {
  const y = 2.32 + i * 0.94;
  s.addText(String(i + 1), { x: 0.6, y, w: 0.5, h: 0.5, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 22, bold: true, color: ROUTES[i].c });
  s.addText(q[0], { x: 1.2, y: y - 0.02, w: 4.1, h: 0.35, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 16, bold: true, color: WHITE });
  s.addText(q[1], { x: 5.4, y: y - 0.06, w: 7.3, h: 0.6, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 12, color: PALE, lineSpacing: 17 });
});
chips(s, 6.5);
s.addText("Nothing gets built until these are settled.", {
  x: 2.5, y: 6.52, w: 8, h: 0.3, isTextBox: true, margin: 0,
  fontFace: SANS, fontSize: 12.5, italic: true, color: DIM });
s.addNotes("Take these four away, come back with answers, and the build order follows immediately.");

pres.writeFile({ fileName: "../Feedback-Engine-Design.pptx" })
    .then(f => console.log("written:", f));
