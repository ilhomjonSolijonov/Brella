// Builds the RoomMatch Task 3 deck (Entrepreneurship, Week 3 – Business Model Canvas I, customer side).
// Run: NODE_PATH=<node_modules with pptxgenjs, react, react-dom, react-icons, sharp> APPLY_THEME=<path to apply_theme.js> node roommatch_deck.js [output.pptx]
// Images (logo, illustrations, background blobs) come from the group's Task 2 deck and live in ./assets.
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2] || "Task3_RoomMatch_BMC_Customer_Side.pptx";
const ASSETS = path.join(__dirname, "assets");

// Palette taken from the Task 2 deck: grey-blue background, deep blue, violet and aqua accents.
const THEME = {
  name: "RoomMatch",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "2B2B2B", lt1: "FFFFFF", dk2: "0F2557", lt2: "D0D7DD",
    accent1: "004AAD", accent2: "8C52FF", accent3: "D64545", accent4: "23884F",
    accent5: "56616B", accent6: "5CE1E6", hlink: "004AAD", folHlink: "5E17EB",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333" x 7.5"
pres.title = "RoomMatch – Business Model Canvas I (Customer side)";
pres.subject = "Entrepreneurship 2026–2027, Task 3";
pres.author = "Gulsevar Abdurasulova, Muqaddas Karamadinova, Ilhomjon Solijonov";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const img = (f) => "image/png;base64," + fs.readFileSync(path.join(ASSETS, f)).toString("base64");
async function svgPng(svg, width) {
  return "image/png;base64," + (await sharp(Buffer.from(svg)).resize({ width }).png().toBuffer()).toString("base64");
}
async function icon(Comp, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: String(size) }));
  return svgPng(svg, size);
}

(async () => {
  const LOGO = await svgPng(fs.readFileSync(path.join(ASSETS, "logo.svg")), 480);
  const LOGO_W = await svgPng(fs.readFileSync(path.join(ASSETS, "logo.svg"), "utf8").replace(/#030403/g, "#FFFFFF"), 480);
  const BLOB1 = img("blob1.png"), BLOB2 = img("blob2.png");

  // ---------- layouts ----------
  const blobs = [
    { image: { x: 8.6, y: -0.55, w: 5.4, h: 2.07, data: BLOB2, transparency: 35 } },
    { image: { x: -1.1, y: 5.3, w: 2.6, h: 2.92, data: BLOB1, transparency: 35 } },
  ];

  pres.defineSlideMaster({
    title: "TITLE",
    background: { color: C.background2 },
    objects: [
      ...blobs,
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.1, w: 6.6, h: 1.2,
        fontSize: 60, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.4, w: 6.6, h: 0.6,
        fontSize: 24, color: C.accent1, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background2 },
    margin: [0.5, 0.6, 0.6, 0.6],
    slideNumber: { x: 12.25, y: 6.98, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
    objects: [
      ...blobs,
      { image: { x: 0.6, y: 0.38, w: 0.66, h: 0.5, data: LOGO } },
      { text: { text: "RoomMatch", options: { x: 1.38, y: 0.33, w: 3.5, h: 0.32, fontSize: 17, bold: true, color: C.text1, margin: 0, fontFace: "Arial" } } },
      { text: { text: "Find the right people to share a home", options: { x: 1.38, y: 0.64, w: 3.5, h: 0.26, fontSize: 11, color: C.text1, margin: 0 } } },
      { text: { text: "Task 3 · Business Model Canvas I: customer side", options: { x: 0.6, y: 6.98, w: 6, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: 5.2, y: 0.33, w: 7.55, h: 0.6,
        fontSize: 26, bold: true, color: C.text1, valign: "middle", align: "right", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "SUMMARY_DARK",
    background: { color: C.text2 },
    objects: [
      { image: { x: 0.6, y: 0.42, w: 0.66, h: 0.5, data: LOGO_W } },
      { placeholder: { options: { name: "title", type: "title", x: 1.45, y: 0.35, w: 9.5, h: 0.65,
        fontSize: 30, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
  });

  // ---------- helpers ----------
  const shadow = () => ({ type: "outer", color: "1B2A3A", opacity: 0.12, blur: 8, offset: 2, angle: 90 });
  const card = (s, x, y, w, h, fill = C.background1, name = "Card", sh = true) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.14, fill: { color: fill },
      line: { type: "none" }, ...(sh ? { shadow: shadow() } : {}), objectName: name });
  const circle = (s, data, x, y, d, fill, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + " circle" });
    const p = d * 0.25;
    s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name + " icon" });
  };
  const txt = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 14, ...o });
  const lead = (s, text) => txt(s, text, { x: 0.6, y: 1.22, w: 12.15, h: 0.5, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
  const pill = (s, text, x, y, w, fill, color, name) => s.addText(text, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, x, y, w, h: 0.34, fill: { color: fill }, line: { type: "none" },
    fontSize: 11, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: name });

  const W = HEX.lt1;
  const I = {};
  for (const [k, comp] of Object.entries({
    user: fa.FaUser, match: fa.FaPeopleArrows, group: fa.FaUserGroup, house: fa.FaHouse,
    grad: fa.FaUserGraduate, brief: fa.FaBriefcase, building: fa.FaBuilding, family: fa.FaPeopleRoof,
    bullhorn: fa.FaBullhorn, search: fa.FaMagnifyingGlass, card: fa.FaCreditCard, telegram: fa.FaTelegram,
    headset: fa.FaHeadset, hands: fa.FaHandshake, robot: fa.FaRobot, users: fa.FaUsers, shield: fa.FaShieldHalved,
    xmark: fa.FaXmark, check: fa.FaCheck, flask: fa.FaFlask,
  })) I[k] = await icon(comp, W);
  I.q = await icon(fa.FaCircleQuestion, HEX.accent2, 512);
  I.arrow = await icon(fa.FaArrowRight, HEX.accent5);

  // ======================= 1. Title =======================
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Opening" });
  s.addImage({ data: LOGO, x: 0.8, y: 1.05, w: 1.2, h: 0.91, objectName: "Logo" });
  s.addText("RoomMatch", { placeholder: "title" });
  s.addText("Find the right people to share a home", { placeholder: "body" });
  txt(s, "Task 3 · Business Model Canvas I: the customer side", { x: 0.8, y: 4.35, w: 6.6, h: 0.4, fontSize: 17, bold: true });
  txt(s, [
    { text: "u040203 · Gulsevar Abdurasulova", options: { breakLine: true } },
    { text: "u040198 · Muqaddas Karamadinova", options: { breakLine: true } },
    { text: "u040190 · Ilhomjon Solijonov" },
  ], { x: 0.8, y: 4.9, w: 6.6, h: 1.1, fontSize: 14, color: C.accent5, paraSpaceAfter: 4 });
  card(s, 7.75, 1.3, 4.95, 3.65, C.background1, "Illustration frame");
  s.addImage({ data: img("students.png"), x: 7.95, y: 1.71, w: 4.55, h: 2.83, objectName: "Students illustration" });
  s.addNotes("Opening, about 15 seconds. Introduce the three of us and the idea name: RoomMatch, find the right people to share a home.");

  // ======================= 2. Business idea =======================
  const SEC = "Customer side of the canvas";
  pres.addSection({ title: SEC });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("1 · BUSINESS IDEA", { placeholder: "title" });
  lead(s, "Listings answer “Where can I live?” We answer “Who can I live with?” first");

  card(s, 0.6, 1.95, 6.95, 4.75, C.background1, "Idea card");
  txt(s, "A roommate-matching platform in a Telegram bot. Students find compatible roommates, form a group of 3–5 and rent an apartment together.",
    { x: 0.9, y: 2.15, w: 6.35, h: 1.0, fontSize: 16 });
  const steps = [
    [I.user, "Profile", "budget, area, lifestyle"],
    [I.match, "Match", "AI match score with similar people"],
    [I.group, "Group", "3–5 compatible roommates"],
    [I.house, "Home", "apartments that fit the group"],
  ];
  steps.forEach(([ic, h, d], i) => {
    const x = 0.85 + i * 1.66;
    circle(s, ic, x + 0.36, 3.35, 0.78, i % 2 ? C.accent2 : C.accent1, "Step " + h);
    txt(s, `0${i + 1} ${h}`, { x, y: 4.25, w: 1.5, h: 0.35, fontSize: 15, bold: true, align: "center", color: C.text2 });
    txt(s, d, { x, y: 4.62, w: 1.5, h: 0.75, fontSize: 13, align: "center", color: C.accent5 });
    if (i < 3) s.addImage({ data: I.arrow, x: x + 1.47, y: 3.61, w: 0.24, h: 0.24, objectName: "Arrow " + (i + 1) });
  });
  txt(s, "Not an apartment listing site: we match people, then help the group find a home.",
    { x: 0.9, y: 5.6, w: 6.35, h: 0.8, fontSize: 14, italic: true, color: C.accent1 });

  card(s, 7.85, 1.95, 4.9, 4.75, C.background1, "Changes card");
  txt(s, "What changed after Task 2", { x: 8.1, y: 2.12, w: 4.4, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
  const changes = [
    ["Customers", "Students + professionals + landlords + everyone", "University students looking for 3–5 roommates"],
    ["Landlords", "Part of the launch", "Added later, once we have ready groups"],
    ["Positioning", "Another housing channel", "Value beyond listings: compatibility matching"],
    ["Free entry", "Pay to use", "Free profile and match preview, pay to build a group"],
  ];
  changes.forEach(([k, before, after], i) => {
    const y = 2.65 + i * 1.0;
    txt(s, k.toUpperCase(), { x: 8.1, y, w: 4.4, h: 0.26, fontSize: 11, bold: true, color: C.accent2, charSpacing: 1 });
    txt(s, [
      { text: before, options: { color: C.accent5, strike: "sngStrike", breakLine: true } },
      { text: "→ " + after, options: { bold: true } },
    ], { x: 8.1, y: y + 0.28, w: 4.4, h: 0.66, fontSize: 13 });
  });
  s.addNotes(
    "Keep this under 1 minute. RoomMatch is a roommate-matching platform. Housing channels answer 'where can I live?'; we answer 'who can I live with?' and then 'where can we live together?'. " +
    "Four steps in a Telegram bot: profile, match, group of 3–5, home. " +
    "After Task 2 we narrowed the customer to university students looking for 3–5 roommates, moved landlords to a later stage because 'will landlords participate?' is untested, positioned ourselves beyond listings because substitutes are the strongest Porter force, " +
    "and made the profile and match preview free because existing Telegram channels are free. " +
    "If the lecturer gave you specific Task 2 feedback, mention it here."
  );

  // ======================= 3. Customer segments =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("2 · CUSTOMER SEGMENTS", { placeholder: "title" });
  lead(s, "Type: multi-sided platform. We serve university students first");

  txt(s, [
    { text: "Two interdependent sides: ", options: { bold: true } },
    { text: "groups need apartments, and landlords need reliable groups. Without students there is no group to offer, so we build that side first." },
  ], { x: 0.6, y: 1.8, w: 12.15, h: 0.55, fontSize: 15 });

  const segs = [
    [I.grad, "University students", "Can't afford to rent alone and don't want crowded housing. Looking for 3–5 roommates.", "SERVE FIRST", C.accent1, true],
    [I.brief, "Young professionals", "Relocating to the city for work; same need, higher budget.", "SECONDARY", C.accent5, false],
    [I.building, "Landlords (lessors)", "Want reliable groups of tenants for 2–4-room apartments.", "OTHER SIDE · LATER", C.accent2, false],
    [I.family, "Families", "Not a target: they usually rent a whole home on their own.", "NOT NOW", C.accent5, false],
  ];
  const sw = 2.92, sg = 0.157;
  segs.forEach(([ic, h, d, tag, tagFill, first], i) => {
    const x = 0.6 + i * (sw + sg), y = 2.55;
    card(s, x, y, sw, 2.75, first ? C.accent1 : C.background1, "Segment " + h);
    circle(s, ic, x + 0.25, y + 0.25, 0.7, first ? C.text2 : tagFill, "Segment " + h);
    pill(s, tag, x + 1.05, y + 0.43, sw - 1.3, first ? C.accent6 : C.background2, first ? C.text2 : C.text1, "Tag " + h);
    txt(s, h, { x: x + 0.25, y: y + 1.12, w: sw - 0.5, h: 0.4, fontSize: 16, bold: true, color: first ? C.background1 : C.text2 });
    txt(s, d, { x: x + 0.25, y: y + 1.55, w: sw - 0.5, h: 1.1, fontSize: 13, color: first ? C.background1 : C.text1 });
  });

  txt(s, "Why students first?", { x: 0.6, y: 5.5, w: 6, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
  const why = [
    ["Strongest pain", "4/5 interviewees called price a major constraint"],
    ["Same place", "students cluster around universities and dorms"],
    ["Our network", "we are students: cheap to reach and to test"],
  ];
  why.forEach(([b, t], i) =>
    txt(s, [{ text: b + ": ", options: { bold: true, color: C.accent1 } }, { text: t }],
      { x: 0.6 + i * 4.1, y: 5.9, w: 3.9, h: 0.8, fontSize: 14 }));
  s.addNotes(
    "About 1.5 minutes. Segment type: multi-sided platform. One side is people looking for roommates, the other is landlords with apartments. The sides depend on each other, which is why the canvas calls it multi-sided. " +
    "We serve university students first: they felt the pain most in our interviews (4 of 5 said price is a major constraint), they are concentrated around campuses, and as students we can reach them cheaply. " +
    "Young professionals are a secondary segment. Landlords are the second side we add later. Families are not a target because they rent a whole home on their own. " +
    "People who prefer to live alone (one of our Task 2 challenges) are also not our customer."
  );

  // ======================= 4. Value proposition =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("3 · VALUE PROPOSITION", { placeholder: "title" });

  card(s, 0.6, 1.25, 12.15, 1.95, C.text2, "Value proposition sentence", false);
  const Bv = (t) => ({ text: t, options: { bold: true, color: C.accent6 } });
  const Nv = (t) => ({ text: t });
  txt(s, [
    Nv("For "), Bv("university students"), Nv(" who "), Bv("can't afford to rent alone but struggle to find enough compatible roommates"),
    Nv(", RoomMatch "), Bv("matches them by budget, area and lifestyle and helps them form a group of 3–5 to rent together"),
    Nv(". Unlike "), Bv("Telegram housing channels and asking friends"), Nv(", we "), Bv("match people by compatibility, not just list apartments"), Nv("."),
  ], { x: 0.95, y: 1.4, w: 11.45, h: 1.65, fontSize: 19, color: C.background1, valign: "middle" });
  txt(s, "Value elements: Cost reduction  ·  Convenience  ·  Risk reduction  ·  Customisation",
    { x: 0.6, y: 3.32, w: 12.15, h: 0.32, fontSize: 13, italic: true, color: C.accent5 });

  const col = (x, title, ic, fill, items, nm) => {
    card(s, x, 3.8, 5.95, 2.95, C.background1, nm + " card");
    txt(s, title, { x: x + 0.25, y: 3.95, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: fill });
    items.forEach(([t, ev], i) => {
      const y = 4.5 + i * 0.55;
      circle(s, ic, x + 0.25, y, 0.4, fill, nm + " " + (i + 1));
      txt(s, t, { x: x + 0.8, y: y - 0.03, w: ev ? 4.15 : 4.9, h: 0.46, fontSize: 14, valign: "middle" });
      if (ev) pill(s, ev, x + 5.0, y + 0.03, 0.7, C.background2, C.text1, "Evidence " + ev);
    });
  };
  col(0.6, "Pains we relieve", I.xmark, C.accent3, [
    ["Renting alone is too expensive", "4/5"],
    ["Can't find enough compatible people", "3/5"],
    ["Forced to compromise on location or comfort", "3/5"],
    ["Telegram channels are fragmented and random", null],
  ], "Pain");
  col(6.8, "Gains we create", I.check, C.accent4, [
    ["Lower rent, split between 3–5 people", null],
    ["Roommates with similar habits and budget", null],
    ["One place from profile to apartment", null],
    ["Verified profiles: less risk with strangers", null],
  ], "Gain");
  s.addNotes(
    "About 1.5 minutes. Read the sentence once, slowly. It follows the Week 3 template: For [customer] who [need], our [product] [benefit]. Unlike [alternative], we [difference]. " +
    "The pains come from our five Task 2 interviews; the numbers show how many of the five mentioned each one. Fragmented Telegram channels came up as support for our hypothesis. " +
    "The gains are what we create: cheaper rent, compatible roommates, one place for the whole process, and verified profiles to reduce the risk of living with strangers."
  );

  // ======================= 5. Channels + relationships =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("4–5 · CHANNELS & RELATIONSHIPS", { placeholder: "title" });
  lead(s, "Channels: we go where students already look, then keep everything in our bot");

  const ch = [
    [I.bullhorn, "Awareness", "University and housing Telegram groups, Instagram, campus posters", "Owned + partner"],
    [I.search, "Evaluation", "Free profile and match preview in the bot", "Owned"],
    [I.card, "Purchase", "Subscription inside the bot, paid with Click or Payme", "Owned + partner"],
    [I.telegram, "Delivery", "Matches, group chat and apartment shortlist in the bot", "Owned"],
    [I.headset, "After-sales", "In-bot support and feedback after move-in", "Owned"],
  ];
  const cw = 2.33, cg = 0.125;
  ch.forEach(([ic, h, d, tag], i) => {
    const x = 0.6 + i * (cw + cg), y = 1.85;
    card(s, x, y, cw, 2.05, C.background1, "Channel " + h);
    circle(s, ic, x + 0.2, y + 0.2, 0.55, C.accent1, "Channel " + h);
    txt(s, h, { x: x + 0.85, y: y + 0.27, w: cw - 0.95, h: 0.4, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    txt(s, d, { x: x + 0.2, y: y + 0.88, w: cw - 0.35, h: 0.8, fontSize: 13 });
    txt(s, tag, { x: x + 0.2, y: y + 1.68, w: cw - 0.4, h: 0.28, fontSize: 11, bold: true, color: tag === "Owned" ? C.accent1 : C.accent2 });
  });

  txt(s, "Relationships: personal help for the first groups, then automated self-service and community",
    { x: 0.6, y: 4.08, w: 12.15, h: 0.4, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
  const rel = [
    [I.hands, "Attract (acquisition)", "Personal assistance", "We personally help the first groups form, and run roommate meetups on campus."],
    [I.robot, "Keep (retention)", "Automated services", "AI match score and alerts on new matches; verified profiles and moderation build trust."],
    [I.users, "Grow (boosting sales)", "Communities + co-creation", "Group chats, reviews of roommates and flats, referrals, and a return visit when the lease ends."],
  ];
  const rw = 3.95, rg = 0.15;
  rel.forEach(([ic, h, type, d], i) => {
    const x = 0.6 + i * (rw + rg), y = 4.6;
    card(s, x, y, rw, 2.1, C.background1, "Relationship " + h);
    circle(s, ic, x + 0.2, y + 0.2, 0.55, C.accent2, "Relationship " + h);
    txt(s, h, { x: x + 0.9, y: y + 0.17, w: rw - 1.05, h: 0.32, fontSize: 15, bold: true, color: C.text2 });
    txt(s, type, { x: x + 0.9, y: y + 0.48, w: rw - 1.05, h: 0.3, fontSize: 12, bold: true, color: C.accent2 });
    txt(s, d, { x: x + 0.2, y: y + 0.95, w: rw - 0.4, h: 1.05, fontSize: 13 });
  });
  s.addNotes(
    "About 2 minutes for both blocks. " +
    "Channels: 4 of 5 interviewees already search through Telegram channels or personal networks, so we advertise there, on Instagram and on campus. Students try the free profile and match preview, pay inside the bot with Click or Payme, and receive matches, a group chat and an apartment shortlist in the same bot. " +
    "Mostly owned, direct channels: the Telegram bot from our Task 2 costs ($100 once). Partners: university Telegram groups and student unions for reach, and payment providers. " +
    "Relationships: to attract the first customers we use personal assistance, doing the first matches almost by hand. To keep them we use automated services, with match scores, alerts and verification for trust. To grow, communities and co-creation: group chats, reviews and referrals. " +
    "A customer may only need us once a year, so referrals and returning at lease end matter."
  );

  // ======================= 6. Revenue streams =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("6 · REVENUE STREAMS", { placeholder: "title" });
  lead(s, "Recurring revenue: subscriptions on both sides of the platform");

  const payer = (y, ic, fill, who, what, price, unit, tag, nm) => {
    card(s, 0.6, y, 7.3, 2.15, C.background1, nm);
    circle(s, ic, 0.85, y + 0.25, 0.65, fill, nm);
    txt(s, who, { x: 1.7, y: y + 0.22, w: 3.6, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
    pill(s, tag, 1.7, y + 0.62, 2.0, C.background2, C.text1, nm + " tag");
    txt(s, what, { x: 0.85, y: y + 1.12, w: 4.45, h: 0.95, fontSize: 13 });
    txt(s, price, { x: 5.35, y: y + 0.3, w: 2.35, h: 0.75, fontSize: 36, bold: true, color: fill, align: "right" });
    txt(s, unit, { x: 5.35, y: y + 1.05, w: 2.35, h: 0.9, fontSize: 12, color: C.accent5, align: "right" });
  };
  payer(1.9, I.grad, C.accent1, "Students (lessees)", "Profile and match preview are free. They pay to message matches, build a group and see apartments. Paid with Click or Payme in the bot.",
    "$10", "per month, lessee subscription", "FIRST · SUBSCRIPTION", "Students payer");
  payer(4.2, I.building, C.accent2, "Landlords (lessors)", "Pay to list apartments and receive requests from ready groups. Premium adds top placement.",
    "$15–20", "per month: ordinary $15, premium $20", "LATER · SUBSCRIPTION", "Landlords payer");

  card(s, 8.2, 1.9, 4.55, 4.45, C.background1, "Unit economics card");
  txt(s, "Link to Task 2 unit economics", { x: 8.45, y: 2.05, w: 4.1, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
  const hdr = (t, al = "right") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: al } });
  const cell = (t, o = {}) => ({ text: t, options: { align: "right", ...o } });
  s.addTable([
    [hdr("Monthly costs", "left"), hdr("Months 1–3"), hdr("Month 4+")],
    [cell("Bot support", { align: "left" }), cell("$20"), cell("$20")],
    [cell("Server", { align: "left" }), cell("$10"), cell("$10")],
    [cell("Marketing", { align: "left" }), cell("$200"), cell("–")],
    [cell("Total", { align: "left", bold: true }), cell("$230", { bold: true }), cell("$30", { bold: true })],
    [cell("Students at $10 to cover it", { align: "left", bold: true, color: C.accent1 }), cell("23", { bold: true, color: C.accent1 }), cell("3", { bold: true, color: C.accent1 })],
  ], {
    x: 8.45, y: 2.5, w: 4.05, colW: [1.95, 1.1, 1.0], rowH: 0.42, fontSize: 12, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.lt2 }, fill: { color: C.background1 }, margin: [0, 0.06, 0, 0.06], valign: "middle",
    objectName: "Break-even table",
  });
  txt(s, "Plus $100 once for the bot = 10 student-months.", { x: 8.45, y: 5.15, w: 4.05, h: 0.5, fontSize: 12, color: C.accent5 });
  txt(s, "Risk: students may subscribe only while searching (1–2 months).", { x: 8.45, y: 5.65, w: 4.05, h: 0.55, fontSize: 12, bold: true, color: C.accent3 });
  s.addNotes(
    "About 1.5 minutes. Who pays: students first, landlords later. For what: students get the profile and match preview free and pay to message matches, build a group and see apartments; landlords pay to list apartments and reach ready groups. " +
    "How: monthly subscriptions paid in the bot with Click or Payme, so this is recurring revenue. " +
    "Link to Task 2: our costs were $100 once for the bot, $20 a month bot support, $10 a month server and $200 a month marketing for the first three months. So we need 23 paying students a month at launch and only 3 a month after month 3 to cover running costs. " +
    "The risk: students may pay only while they search, so revenue can behave like one-off transactions. " +
    "Check: the Task 2 table showed no currency. We assumed US dollars; change the $ signs if you meant something else."
  );

  // ======================= 7. Least sure =======================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SEC });
  s.addText("7 · WHAT WE ARE LEAST SURE ABOUT", { placeholder: "title" });
  lead(s, "Evidence from 5 interviews (3 students, 1 graduate, 1 young professional, 5–7 October)");

  txt(s, "Block", { x: 0.85, y: 1.85, w: 2.2, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  txt(s, "Based on", { x: 3.05, y: 1.85, w: 3.2, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  txt(s, "Confidence", { x: 6.35, y: 1.85, w: 1.5, h: 0.3, fontSize: 12, bold: true, color: C.accent5 });
  const ev = [
    ["Customer segments", "Interviews: students struggle most; only 5 people so far", 3],
    ["Value proposition", "Interviews: 4/5 price, 3/5 too few compatible roommates", 3],
    ["Channels", "Interviews: 4/5 already use Telegram or friends", 3],
    ["Relationships", "Assumption: will students trust strangers online?", 2],
    ["Revenue streams", "Assumption: no one has paid; Telegram channels are free", 1],
  ];
  ev.forEach(([b, src, conf], i) => {
    const y = 2.2 + i * 0.9, least = conf === 1;
    card(s, 0.6, y, 7.35, 0.78, least ? C.accent2 : C.background1, "Evidence " + b, false);
    const fg = least ? C.background1 : C.text1;
    txt(s, b, { x: 0.85, y, w: 2.15, h: 0.78, fontSize: 15, bold: true, color: least ? C.background1 : C.text2, valign: "middle" });
    txt(s, src, { x: 3.05, y, w: 3.15, h: 0.78, fontSize: 12, valign: "middle", color: fg });
    for (let d = 0; d < 4; d++) {
      s.addShape(pres.shapes.OVAL, { x: 6.4 + d * 0.36, y: y + 0.28, w: 0.22, h: 0.22,
        fill: { color: d < conf ? (least ? C.background1 : C.accent1) : (least ? C.accent2 : C.background1) },
        line: { color: least ? HEX.lt1 : HEX.accent1, width: 1 }, objectName: `Confidence ${b} ${d + 1}` });
    }
  });

  card(s, 8.25, 1.85, 4.5, 4.85, C.text2, "Least sure card", false);
  s.addImage({ data: I.q, x: 8.5, y: 2.08, w: 0.6, h: 0.6, objectName: "Question icon" });
  txt(s, "Least sure: revenue streams", { x: 9.25, y: 2.08, w: 3.35, h: 0.6, fontSize: 17, bold: true, color: C.accent6, valign: "middle" });
  txt(s, "Will students pay $10 a month when Telegram housing channels are free?",
    { x: 8.5, y: 2.9, w: 4.0, h: 0.95, fontSize: 15, bold: true, color: C.background1 });
  txt(s, "Interviews showed the pain, not willingness to pay. None of the 5 has ever paid for roommate matching.",
    { x: 8.5, y: 3.85, w: 4.0, h: 0.85, fontSize: 13, color: C.background2 });
  circle(s, I.flask, 8.5, 4.85, 0.5, C.accent2, "Test");
  txt(s, "How we will test it", { x: 9.15, y: 4.87, w: 3.3, h: 0.45, fontSize: 15, bold: true, color: C.background1, valign: "middle" });
  txt(s, "Run matching in the bot for 2 weeks at our university. Success: at least 10 of 30 active students pay to build a group.",
    { x: 8.5, y: 5.45, w: 4.0, h: 1.1, fontSize: 13, color: C.background1 });
  s.addNotes(
    "About 1.5 minutes. Segments, value proposition and channels are based on our five Task 2 interviews from 5 to 7 October: 4 of 5 used Telegram channels or personal networks, 4 of 5 said price is a major constraint, 3 of 5 struggled to find enough compatible roommates, and 3 of 5 compromised on location or comfort. " +
    "Relationships rest on an assumption: that students will trust strangers they meet online. " +
    "We are least sure about revenue streams: no one has paid for matching yet and Telegram channels are free. Following The Mom Test, we will ask for a real commitment: two weeks of matching in the bot, with success meaning at least 10 of 30 active students pay. " +
    "Second-least sure: whether landlords will participate."
  );

  // ======================= 8. Summary =======================
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "SUMMARY_DARK", sectionTitle: "Close" });
  s.addText("Our customer side on one page", { placeholder: "title" });
  const blk = (x, y, w, h, title, body, nm, hi = false, ic = null) => {
    card(s, x, y, w, h, hi ? C.accent1 : C.background1, nm, false);
    if (ic) circle(s, ic, x + w - 0.85, y + h - 0.85, 0.6, hi ? C.text2 : C.accent2, nm);
    txt(s, title.toUpperCase(), { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: hi ? C.accent6 : C.accent2, charSpacing: 1 });
    txt(s, body, { x: x + 0.2, y: y + 0.5, w: w - 0.4, h: h - 0.6, fontSize: 16, color: hi ? C.background1 : C.text1 });
  };
  blk(0.6, 1.35, 4.0, 3.65, "Value proposition", "Compatible roommates matched by budget, area and lifestyle, and a group of 3–5 ready to rent together. Not just listings.", "VP block", true, I.match);
  blk(4.75, 1.35, 3.95, 1.75, "Customer relationships", "Personal help for the first groups; automated matching, verification and community after.", "CR block");
  blk(4.75, 3.25, 3.95, 1.75, "Channels", "Telegram groups, Instagram and campus to reach; our Telegram bot to sell and deliver.", "CH block");
  blk(8.85, 1.35, 3.9, 3.65, "Customer segments", "Multi-sided platform. First: university students looking for 3–5 roommates. Later: young professionals and landlords.", "CS block", false, I.grad);
  blk(0.6, 5.15, 12.15, 1.15, "Revenue streams", "Subscriptions: students $10 a month; landlords $15–20 a month later. 23 paying students a month cover launch costs.", "RS block");
  txt(s, "Questions?", { x: 9.6, y: 6.5, w: 3.15, h: 0.6, fontSize: 24, bold: true, color: C.accent6, align: "right", valign: "middle" });
  s.addNotes("About 30 seconds. Summarise the five customer-side blocks and open for questions. All three of us should be ready to answer on any block, especially the revenue test and why students come first.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", path.resolve(OUT));
})();
