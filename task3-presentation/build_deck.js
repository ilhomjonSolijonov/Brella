// Builds Task3_BMC_Customer_Side.pptx (Entrepreneurship, Week 3 – Business Model Canvas I).
// Run: NODE_PATH=<node_modules with pptxgenjs, react, react-dom, react-icons, sharp> APPLY_THEME=<path to apply_theme.js> node build_deck.js [output.pptx]
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2] || "Task3_BMC_Customer_Side.pptx";

const THEME = {
  name: "WashGo",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A2B2E", lt1: "FFFFFF", dk2: "0B4F5C", lt2: "EAF5F6",
    accent1: "0F8B8D", accent2: "F2A541", accent3: "D9534F", accent4: "2E9E5B",
    accent5: "6B7C80", accent6: "B9E1E3", hlink: "0F8B8D", folHlink: "0B4F5C",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333" x 7.5"
pres.title = "WashGo – Business Model Canvas I (Customer side)";
pres.subject = "Entrepreneurship 2026–2027, Task 3";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- icons ----------
async function icon(Comp, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + hex, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 7.6, h: 1.4,
      fontSize: 60, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.55, w: 7.6, h: 1.2,
      fontSize: 24, color: C.accent6, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "SUMMARY_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.2, h: 0.85,
      fontSize: 32, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  slideNumber: { x: 12.3, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  objects: [
    { text: { text: "WashGo  ·  Task 3: Business Model Canvas I (Customer side)",
      options: { x: 0.6, y: 6.95, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.2, h: 0.85,
      fontSize: 32, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });

function card(slide, x, y, w, h, fill = C.background2, name, withShadow = false) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" },
    ...(withShadow ? { shadow: shadow() } : {}), objectName: name,
  });
}

function iconCircle(slide, data, x, y, d, fill, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + " circle" });
  const p = d * 0.24;
  slide.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name + " icon" });
}

function chip(slide, text, x = 11.0, y = 0.6, w = 1.75) {
  slide.addText(text, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.2, x, y, w, h: 0.42,
    fill: { color: C.background2 }, line: { type: "none" }, fontSize: 12, bold: true,
    color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Point chip",
  });
}

function txt(slide, text, opts) {
  slide.addText(text, { isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 14, ...opts });
}

(async () => {
  const W = HEX.lt1;
  const I = {
    shirt: await icon(fa.FaShirt, W), shirtBig: await icon(fa.FaShirt, HEX.accent2, 512),
    telegram: await icon(fa.FaTelegram, W), truck: await icon(fa.FaTruck, W),
    soap: await icon(fa.FaJugDetergent, W), clock: await icon(fa.FaClock, W),
    grad: await icon(fa.FaUserGraduate, W), flat: await icon(fa.FaHouseUser, W),
    brief: await icon(fa.FaBriefcase, W), hotel: await icon(fa.FaHotel, W),
    gradTeal: await icon(fa.FaUserGraduate, HEX.accent1),
    bullhorn: await icon(fa.FaBullhorn, W), search: await icon(fa.FaMagnifyingGlass, W),
    card: await icon(fa.FaCreditCard, W), headset: await icon(fa.FaHeadset, W),
    xmark: await icon(fa.FaXmark, W), check: await icon(fa.FaCheck, W),
    hands: await icon(fa.FaHandshake, W), robot: await icon(fa.FaArrowsRotate, W),
    plus: await icon(fa.FaPlus, W), flask: await icon(fa.FaFlask, W),
    q: await icon(fa.FaCircleQuestion, HEX.accent2, 512),
    arrow: await icon(fa.FaArrowRight, HEX.accent5),
  };

  // ======================= Slide 1: title =======================
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  s.addText("WashGo", { placeholder: "title" });
  s.addText("Laundry pick-up for students in Tashkent dormitories", { placeholder: "body" });
  txt(s, "Task 3 · Business Model Canvas I: the customer side", { x: 0.8, y: 5.0, w: 7.6, h: 0.4, fontSize: 16, color: C.background1, bold: true });
  txt(s, "Group [number]  ·  [Member names]  ·  Entrepreneurship 2026–2027", { x: 0.8, y: 5.45, w: 7.6, h: 0.4, fontSize: 14, color: C.accent6 });
  s.addShape(pres.shapes.OVAL, { x: 8.9, y: 1.45, w: 3.9, h: 3.9, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Hero circle" });
  s.addImage({ data: I.shirtBig, x: 9.85, y: 2.4, w: 2.0, h: 2.0, objectName: "Hero shirt icon" });
  s.addNotes(
    "Opening (about 15 seconds). Introduce the group and the idea name. " +
    "Before presenting: replace 'Group [number]' and '[Member names]' with your details."
  );

  // ======================= Slide 2: business idea =======================
  pres.addSection({ title: "Customer side of the canvas" });
  const SEC = "Customer side of the canvas";
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("The idea: we do the laundry for dorm students", { placeholder: "title" });
  chip(s, "Point 1 of 7");

  card(s, 0.6, 1.5, 6.6, 1.55, C.background2, "Idea card");
  txt(s, [
    { text: "Students order in our Telegram bot. ", options: { bold: true, color: C.text2 } },
    { text: "We collect the bag at the dorm entrance, a partner laundry washes, dries and folds it, and we bring it back within 48 hours." },
  ], { x: 0.85, y: 1.7, w: 6.1, h: 1.2, fontSize: 18, valign: "middle" });

  const steps = [
    [I.telegram, "Order", "in the Telegram bot"],
    [I.truck, "Pick-up", "Tue & Fri evenings at the dorm"],
    [I.soap, "Wash & fold", "at a partner laundry"],
    [I.clock, "Return", "within 48 hours"],
  ];
  steps.forEach(([ic, h, d], i) => {
    const x = 0.6 + i * 1.72;
    iconCircle(s, ic, x + 0.37, 3.75, 0.8, i % 2 ? C.text2 : C.accent1, "Step " + (i + 1));
    txt(s, h, { x, y: 4.7, w: 1.55, h: 0.35, fontSize: 15, bold: true, align: "center", color: C.text2 });
    txt(s, d, { x, y: 5.08, w: 1.55, h: 0.7, fontSize: 13, align: "center", color: C.accent5 });
    if (i < 3) s.addImage({ data: I.arrow, x: x + 1.5, y: 4.02, w: 0.26, h: 0.26, objectName: "Step arrow " + (i + 1) });
  });

  card(s, 7.6, 1.5, 5.15, 5.2, C.background1, "Changes card", true);
  txt(s, "What changed after the Task 2 feedback", { x: 7.9, y: 1.7, w: 4.6, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
  const changes = [
    ["Scope", "All of Tashkent", "3 dormitories near our campus"],
    ["Washing", "Buying our own machines", "Partner laundry, no equipment cost"],
    ["Price", "Per kilogram", "One price per bag (up to 6 kg)"],
    ["Delivery", "On demand", "Two fixed pick-up slots a week"],
  ];
  changes.forEach(([k, before, after], i) => {
    const y = 2.3 + i * 1.08;
    txt(s, k.toUpperCase(), { x: 7.9, y, w: 4.6, h: 0.28, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1 });
    txt(s, [
      { text: before, options: { color: C.accent5, strike: "sngStrike" } },
      { text: "  →  ", options: { color: C.accent5 } },
      { text: after, options: { bold: true, color: C.text1 } },
    ], { x: 7.9, y: y + 0.3, w: 4.6, h: 0.6, fontSize: 14 });
  });
  s.addNotes(
    "Keep this under 1 minute. " +
    "Idea: students in dormitories have no easy way to wash and dry clothes. WashGo collects the bag at the dorm, a partner laundry washes, dries and folds it, and we return it within 48 hours. " +
    "Changes after Task 2 feedback: we narrowed the scope to three dormitories, switched from buying machines to a partner laundry, moved from per-kg to per-bag pricing because interviewees found per-kg confusing, and replaced on-demand delivery with two fixed slots to keep courier costs down. " +
    "REPLACE: make these four changes match the feedback your group actually received in Task 2."
  );

  // ======================= Slide 3: customer segments =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("Customer segments: dorm students first", { placeholder: "title" });
  chip(s, "Point 2 of 7");

  txt(s, [
    { text: "Segment type: segmented. ", options: { bold: true, color: C.text2 } },
    { text: "Related groups that all need clean clothes, but differ in price sensitivity, channels and add-ons." },
  ], { x: 0.6, y: 1.45, w: 12.1, h: 0.45, fontSize: 16 });

  const segs = [
    [I.grad, "Students in dormitories", "Shared or broken machines, no drying space, tight budget. Reached through dorm Telegram groups.", true],
    [I.flat, "Students in shared flats", "Have a machine but little time. Reached through university channels.", false],
    [I.brief, "Young professionals", "Less price-sensitive, want ironing. Reached through Instagram.", false],
    [I.hotel, "Dorms and hostels (B2B)", "Bulk bedding laundry. Personal sales. Later, not now.", false],
  ];
  const sw = 2.9, sg = 0.17;
  segs.forEach(([ic, h, d, first], i) => {
    const x = 0.6 + i * (sw + sg), y = 2.15;
    card(s, x, y, sw, 2.85, first ? C.accent1 : C.background2, "Segment " + (i + 1), first);
    iconCircle(s, ic, x + 0.25, y + 0.3, 0.7, first ? C.text2 : C.accent1, "Segment " + (i + 1));
    if (first) {
      s.addText("SERVE FIRST", { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, x: x + 1.2, y: y + 0.45,
        w: 1.45, h: 0.38, fill: { color: C.accent2 }, line: { type: "none" }, fontSize: 11, bold: true,
        color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Serve first tag" });
    }
    txt(s, h, { x: x + 0.25, y: y + 1.15, w: sw - 0.5, h: 0.4, fontSize: 16, bold: true, color: first ? C.background1 : C.text2 });
    txt(s, d, { x: x + 0.25, y: y + 1.6, w: sw - 0.5, h: 1.15, fontSize: 13, color: first ? C.background1 : C.text1 });
  });

  txt(s, "Why dorm students first?", { x: 0.6, y: 5.25, w: 6, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
  const why = [
    ["Dense", "hundreds of customers in one building, so one trip collects many bags"],
    ["Strongest pain", "the problem came up most often in our interviews"],
    ["Easy to reach", "we study next to these dorms and are in their chats"],
  ];
  why.forEach(([b, t], i) => {
    txt(s, [{ text: b + ": ", options: { bold: true, color: C.accent1 } }, { text: t }],
      { x: 0.6 + i * 4.1, y: 5.7, w: 3.85, h: 0.9, fontSize: 14 });
  });
  s.addNotes(
    "About 1.5 minutes. Our customers are people who need clean clothes but lack time, machines or space. " +
    "Segment type is segmented: the groups are related, but differ in how much they will pay, where we reach them and which extras they want, so each would need a slightly different offer. " +
    "We serve dorm students first: they are concentrated in a few buildings, which makes pick-up cheap; their pain was strongest in our interviews; and we can reach them ourselves. " +
    "B2B (dorm administrations and hostels) is a later option, not part of the first launch."
  );

  // ======================= Slide 4: value proposition =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("Value proposition", { placeholder: "title" });
  chip(s, "Point 3 of 7");

  card(s, 0.6, 1.45, 12.15, 1.95, C.text2, "Value proposition card");
  const B = (t) => ({ text: t, options: { bold: true, color: C.accent2 } });
  const N = (t) => ({ text: t });
  txt(s, [
    N("For "), B("students in Tashkent dormitories"), N(" who "), B("lose hours every week queuing for shared machines and drying clothes in their rooms"),
    N(", WashGo "), B("collects, washes, dries and folds their laundry and returns it within 48 hours"),
    N(". Unlike "), B("the dorm laundry room or a city laundry"), N(", we "), B("come to the dorm door at a student price, booked in Telegram in under a minute"), N("."),
  ], { x: 0.95, y: 1.6, w: 11.45, h: 1.65, fontSize: 19, color: C.background1, valign: "middle", lineSpacingMultiple: 1.05 });

  txt(s, "Value elements: Convenience  ·  Getting the job done  ·  Price", { x: 0.6, y: 3.55, w: 12, h: 0.35, fontSize: 13, italic: true, color: C.accent5 });

  const col = (x, title, ic, fill, items, nm) => {
    txt(s, title, { x, y: 4.05, w: 5.8, h: 0.4, fontSize: 18, bold: true, color: fill });
    items.forEach((t, i) => {
      const y = 4.6 + i * 0.68;
      iconCircle(s, ic, x, y, 0.46, fill, nm + " " + (i + 1));
      txt(s, t, { x: x + 0.65, y: y - 0.02, w: 5.1, h: 0.55, fontSize: 15, valign: "middle" });
    });
  };
  col(0.6, "Pains we relieve", I.xmark, C.accent3, [
    "Queues for shared machines that are often broken",
    "Nowhere to dry clothes, so rooms stay damp",
    "Hand-washing takes 2–3 hours every week",
  ], "Pain");
  col(6.95, "Gains we create", I.check, C.accent4, [
    "More free time for study, work and rest",
    "Clean, folded clothes on a fixed schedule",
    "A known price per bag, paid by Click or Payme",
  ], "Gain");
  s.addNotes(
    "About 1.5 minutes. Read the sentence once, slowly. It follows the Week 3 template: For [customer] who [problem], our [product] [benefit]. Unlike [alternative], we [difference]. " +
    "Pains come from the interviews: queues and broken machines, no drying space, time spent hand-washing. " +
    "Gains: time back, a predictable schedule and a fixed price. The main value elements are convenience, getting the job done and price."
  );

  // ======================= Slide 5: channels + relationships =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("Channels and customer relationships", { placeholder: "title" });
  chip(s, "Points 4–5 of 7");

  txt(s, "Channels: how students find us, buy and receive the service", { x: 0.6, y: 1.4, w: 12, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
  const ch = [
    [I.bullhorn, "Awareness", "Floor ambassadors, posters, dorm Telegram groups", "Owned"],
    [I.search, "Evaluation", "First bag at half price; prices shown in the bot", "Owned"],
    [I.card, "Purchase", "Telegram bot; pay with Click or Payme", "Owned + partner"],
    [I.truck, "Delivery", "Our courier at the dorm; partner laundry washes", "Owned + partner"],
    [I.headset, "After-sales", "Support chat inside the bot", "Owned"],
  ];
  const cw = 2.33, cg = 0.12;
  ch.forEach(([ic, h, d, tag], i) => {
    const x = 0.6 + i * (cw + cg), y = 1.9;
    card(s, x, y, cw, 2.0, C.background2, "Channel " + h);
    iconCircle(s, ic, x + 0.2, y + 0.2, 0.55, C.accent1, "Channel " + h);
    txt(s, h, { x: x + 0.85, y: y + 0.27, w: cw - 0.95, h: 0.4, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    txt(s, d, { x: x + 0.2, y: y + 0.88, w: cw - 0.4, h: 0.7, fontSize: 13 });
    txt(s, tag, { x: x + 0.2, y: y + 1.6, w: cw - 0.4, h: 0.3, fontSize: 11, bold: true, color: tag === "Owned" ? C.accent1 : C.accent5 });
  });

  txt(s, "Customer relationships: personal assistance to win the first 100, then self-service and automation", { x: 0.6, y: 4.15, w: 12.1, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
  const rel = [
    [I.hands, "Attract (acquisition)", "Personal assistance", "An ambassador on each floor, a launch-day stand in the lobby, and a referral code: both friends get 10,000 UZS off."],
    [I.robot, "Keep (retention)", "Automated services", "Bot reminders before each pick-up, saved preferences, and a monthly plan that renews itself."],
    [I.plus, "Grow (boosting sales)", "Self-service + co-creation", "Add-ons in the bot (ironing, bedding); ratings after each order shape the service."],
  ];
  const rw = 3.95, rg = 0.15;
  rel.forEach(([ic, h, type, d], i) => {
    const x = 0.6 + i * (rw + rg), y = 4.6;
    card(s, x, y, rw, 2.15, C.background1, "Relationship " + (i + 1), true);
    iconCircle(s, ic, x + 0.2, y + 0.2, 0.55, C.text2, "Relationship " + (i + 1));
    txt(s, h, { x: x + 0.9, y: y + 0.17, w: rw - 1.05, h: 0.32, fontSize: 15, bold: true, color: C.text2 });
    txt(s, type, { x: x + 0.9, y: y + 0.48, w: rw - 1.05, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
    txt(s, d, { x: x + 0.2, y: y + 0.95, w: rw - 0.4, h: 1.1, fontSize: 13 });
  });
  s.addNotes(
    "About 2 minutes for both blocks. " +
    "Channels: mostly owned and direct. Students already live in dorm Telegram groups, so the bot is where they hear about us, check prices, order and pay. Partners handle payments (Click, Payme) and the washing itself. We chose owned channels because they cost little to start and keep us close to the customer. " +
    "Relationships: to win the first customers we use personal assistance (ambassadors, a launch-day stand, referrals). To keep them we automate: reminders, saved preferences and a monthly plan. To grow, we sell add-ons and use ratings after each order (co-creation)."
  );

  // ======================= Slide 6: revenue streams =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("Revenue streams: pay per bag or monthly", { placeholder: "title" });
  chip(s, "Point 6 of 7");

  const wfh = [
    ["Who pays", "The student, sometimes with money from parents"],
    ["For what", "Wash, dry and fold of one bag (up to 6 kg), plus add-ons"],
    ["How", "Click or Payme in the bot, or cash at pick-up"],
  ];
  wfh.forEach(([k, v], i) => {
    const y = 1.5 + i * 0.97;
    card(s, 0.6, y, 3.85, 0.88, C.background2, k + " card");
    txt(s, k.toUpperCase(), { x: 0.82, y: y + 0.1, w: 3.3, h: 0.25, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1 });
    txt(s, v, { x: 0.82, y: y + 0.36, w: 3.5, h: 0.45, fontSize: 13 });
  });

  const tile = (x, y, label, price, sub, fill, fg, nm) => {
    card(s, x, y, 3.85, 1.55, fill, nm, true);
    txt(s, label, { x: x + 0.25, y: y + 0.15, w: 3.4, h: 0.3, fontSize: 12, bold: true, color: fg });
    txt(s, price, { x: x + 0.25, y: y + 0.45, w: 3.4, h: 0.55, fontSize: 28, bold: true, color: fg });
    txt(s, sub, { x: x + 0.25, y: y + 1.05, w: 3.4, h: 0.35, fontSize: 13, color: fg });
  };
  tile(0.6, 4.5, "TRANSACTION · USAGE FEE", "45,000 UZS", "Single bag, pay per order", C.accent1, C.background1, "Single bag tile");
  tile(4.6, 4.5, "RECURRING · SUBSCRIPTION", "160,000 UZS", "Monthly plan: 4 bags (40,000 each)", C.text2, C.background1, "Monthly plan tile");
  txt(s, [{ text: "Add-ons: ", options: { bold: true } }, { text: "ironing 3,000 UZS per item · bedding set 25,000 UZS" }],
    { x: 0.6, y: 6.25, w: 7.85, h: 0.35, fontSize: 13, color: C.accent5 });

  card(s, 4.6, 1.5, 3.85, 2.79, C.background2, "Mix card");
  txt(s, "Why both?", { x: 4.85, y: 1.65, w: 3.4, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
  txt(s, [
    { text: "Single bags lower the barrier to a first try.", options: { bullet: true, breakLine: true } },
    { text: "The monthly plan trades margin for predictable, recurring revenue and fewer lost customers.", options: { bullet: true } },
  ], { x: 4.85, y: 2.1, w: 3.4, h: 2.1, fontSize: 14, paraSpaceAfter: 8 });

  // unit economics table
  txt(s, "Unit economics per bag (from Task 2)", { x: 8.75, y: 1.5, w: 4.0, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
  const hdr = (t, al = "right") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: al } });
  const cell = (t, o = {}) => ({ text: t, options: { align: "right", ...o } });
  const rows = [
    [hdr("UZS", "left"), hdr("Single"), hdr("Plan")],
    [cell("Price", { align: "left" }), cell("45,000"), cell("40,000")],
    [cell("Partner laundry", { align: "left" }), cell("−22,000"), cell("−22,000")],
    [cell("Courier & transport", { align: "left" }), cell("−6,000"), cell("−6,000")],
    [cell("Bag & payment fee", { align: "left" }), cell("−2,000"), cell("−2,000")],
    [cell("Contribution", { align: "left", bold: true, color: C.accent4 }), cell("15,000", { bold: true, color: C.accent4 }), cell("10,000", { bold: true, color: C.accent4 })],
    [cell("Margin", { align: "left", bold: true, color: C.accent4 }), cell("33%", { bold: true, color: C.accent4 }), cell("25%", { bold: true, color: C.accent4 })],
  ];
  s.addTable(rows, {
    x: 8.75, y: 1.95, w: 4.0, colW: [1.9, 1.05, 1.05], rowH: 0.4, fontSize: 13, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent6 }, fill: { color: C.background1 }, margin: [0, 0.08, 0, 0.08], valign: "middle",
    objectName: "Unit economics table",
  });
  txt(s, "Courier cost assumes about 15 bags per pick-up trip. Fixed costs come next week (cost structure).",
    { x: 8.75, y: 4.95, w: 4.0, h: 0.75, fontSize: 12, italic: true, color: C.accent5 });
  s.addNotes(
    "About 1.5 minutes. The student pays, usually from their own allowance, sometimes with parents' money. They pay for wash, dry and fold of one bag. " +
    "Two streams: a transaction usage fee (45,000 UZS per bag) and a recurring monthly plan (160,000 UZS for 4 bags). " +
    "Link to Task 2: each single bag leaves 15,000 UZS contribution (33%), each plan bag 10,000 UZS (25%). Courier cost only works if we batch about 15 bags per trip, which is why we start with dense dormitories. " +
    "REPLACE: put your real Task 2 unit economics into this table if they differ."
  );

  // ======================= Slide 7: least sure =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("What we are least sure about", { placeholder: "title" });
  chip(s, "Point 7 of 7");

  txt(s, "Block", { x: 0.6, y: 1.45, w: 2.4, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  txt(s, "Based on", { x: 3.05, y: 1.45, w: 4.0, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  txt(s, "Confidence", { x: 6.2, y: 1.45, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  const ev = [
    ["Customer segments", "Interviews: 6 of 8 dorm students named laundry a weekly problem", 4],
    ["Value proposition", "Interviews: queues and drying space came up most", 4],
    ["Channels", "Mixed: all 8 use dorm Telegram chats; ambassadors are assumed", 3],
    ["Relationships", "Assumption: referrals and reminders are not tested", 2],
    ["Revenue streams", "Assumption: nobody has paid us yet", 1],
  ];
  ev.forEach(([b, src, conf], i) => {
    const y = 1.85 + i * 0.95, least = conf === 1;
    card(s, 0.6, y, 7.3, 0.8, least ? C.accent2 : C.background2, "Evidence " + b);
    txt(s, b, { x: 0.8, y, w: 2.2, h: 0.8, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    txt(s, src, { x: 3.05, y, w: 3.0, h: 0.8, fontSize: 12, valign: "middle" });
    for (let d = 0; d < 4; d++) {
      s.addShape(pres.shapes.OVAL, { x: 6.3 + d * 0.36, y: y + 0.29, w: 0.22, h: 0.22,
        fill: { color: d < conf ? C.text2 : C.background1 }, line: { color: C.text2, width: 1 }, objectName: `Confidence ${b} ${d + 1}` });
    }
  });

  card(s, 8.3, 1.45, 4.45, 5.25, C.text2, "Least sure card");
  s.addImage({ data: I.q, x: 8.6, y: 1.7, w: 0.65, h: 0.65, objectName: "Question icon" });
  txt(s, "Least sure: revenue streams", { x: 9.4, y: 1.72, w: 3.2, h: 0.6, fontSize: 17, bold: true, color: C.accent2, valign: "middle" });
  txt(s, "Will dorm students pay 45,000 UZS a bag, and will anyone take the monthly plan?",
    { x: 8.6, y: 2.55, w: 3.95, h: 0.95, fontSize: 15, bold: true, color: C.background1 });
  txt(s, "Interviews showed the pain, not willingness to pay: saying yes is not paying.",
    { x: 8.6, y: 3.55, w: 3.95, h: 0.75, fontSize: 13, color: C.accent6 });
  iconCircle(s, I.flask, 8.6, 4.45, 0.5, C.accent1, "Test");
  txt(s, "How we will test it", { x: 9.25, y: 4.47, w: 3.3, h: 0.45, fontSize: 15, bold: true, color: C.background1, valign: "middle" });
  txt(s, "Pre-sell in one dorm for 2 weeks at full price. Success: at least 20 paid bags and 5 monthly plans.",
    { x: 8.6, y: 5.1, w: 3.95, h: 1.4, fontSize: 13, color: C.background1 });
  s.addNotes(
    "About 1.5 minutes. Segments and the value proposition rest on our Task 2 interviews. Channels are half evidence, half assumption. Relationships and revenue are assumptions. " +
    "We are least sure about revenue: interviewees described the pain, but none has paid. Following The Mom Test, we will ask for a commitment: a two-week pre-sale in one dorm at full price. " +
    "REPLACE: change '6 of 8' and 'all 8' to the real numbers from your Task 2 interviews."
  );

  // ======================= Slide 8: summary =======================
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "SUMMARY_DARK", sectionTitle: "Close" });
  s.addText("Our customer side on one page", { placeholder: "title" });
  const blk = (x, y, w, h, title, body, nm, hi = false, ic = null) => {
    card(s, x, y, w, h, hi ? C.accent1 : C.background1, nm);
    if (ic) iconCircle(s, ic, x + w - 0.85, y + h - 0.85, 0.6, hi ? C.text2 : C.accent1, nm);
    txt(s, title.toUpperCase(), { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: hi ? C.background1 : C.accent1, charSpacing: 1 });
    txt(s, body, { x: x + 0.2, y: y + 0.5, w: w - 0.4, h: h - 0.6, fontSize: 16, color: hi ? C.background1 : C.text1 });
  };
  blk(0.6, 1.4, 4.0, 3.6, "Value proposition", "Laundry collected at the dorm, washed, dried and folded, and back within 48 hours. Booked in Telegram in under a minute, at a student price.", "VP block", true, I.shirt);
  blk(4.75, 1.4, 3.95, 1.72, "Customer relationships", "Personal assistance to win the first 100; automated reminders and a monthly plan to keep them.", "CR block");
  blk(4.75, 3.28, 3.95, 1.72, "Channels", "Owned: Telegram bot, floor ambassadors, our courier. Partners: laundry, Click and Payme.", "CH block");
  blk(8.85, 1.4, 3.9, 3.6, "Customer segments", "Segmented market. First: students in 3 dormitories near our campus. Later: shared flats, young professionals, then B2B.", "CS block", false, I.grad);
  blk(0.6, 5.15, 12.15, 1.15, "Revenue streams", "45,000 UZS per bag (transaction) or 160,000 UZS a month for 4 bags (subscription); margin 33% / 25%.", "RS block");
  txt(s, "Questions?", { x: 9.6, y: 6.5, w: 3.15, h: 0.6, fontSize: 24, bold: true, color: C.accent2, align: "right", valign: "middle" });
  s.addNotes(
    "About 30 seconds. Summarise the five customer-side blocks on the canvas and open for questions. " +
    "Every group member should be ready to answer questions on any block, especially the revenue test and why dorm students come first."
  );

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", path.resolve(OUT));
})();
