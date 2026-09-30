// Sabio proactive strategy deck renderer.
//
// Takes a JSON spec (see weeks/*/specs) and renders a 16:9 .pptx in Sabio
// house style: Sabio Brand Guideline palette (grey theme 60%, Sabio yellow
// #FFC63E primary, blue #3E77FF secondary), Poppins type, the official Sabio
// logo, and the page flow of Ben's 2026 decks (APEX, Kinesso, Walmart Sparky):
// cover → agenda → what's happening at the client → the opening → Sabio at a
// glance → audiences → activation → measurement → proof → next moves → thanks
// → sources.
//
// Usage: node lib/sabio-deck.js path/to/spec.json path/to/out.pptx

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");

// Sabio Brand Guideline (Branding Materials/Sabio_Styleguide) + Sabio Holdings palette.
const C = {
  bg: "1A1A1A", // sabio dark grey
  card: "262626",
  card2: "2F2F2F",
  line: "3A3A3A",
  yellow: "FFC63E", // Primary 500, sabio yellow
  yellowSoft: "FFE7B5", // Primary 200
  blue: "3E77FF", // Secondary 500
  white: "FFFFFF",
  text: "E3E3E3", // Gray 50
  muted: "A7A7A7", // Gray 200
  dim: "7E7E7E", // Gray 300
  ink: "1A1A1A",
};
// Poppins is the Sabio body face and App Science®'s heading face. Field Gothic
// XWide (Sabio headings) is a licensed font most viewers won't have, so
// headings use Poppins ExtraBold to render the same on every machine.
const F = {
  head: "Poppins ExtraBold",
  semi: "Poppins SemiBold",
  body: "Poppins",
};
const W = 10;
const H = 5.625;
const MX = 0.5;
const CW = W - 2 * MX;
const ASSETS = path.join(__dirname, "..", "assets");
const LOGO = {
  yellow: path.join(ASSETS, "sabio-logo-yellow.png"),
  white: path.join(ASSETS, "sabio-logo-white-sm.png"),
  grey: path.join(ASSETS, "sabio-logo-grey.png"),
  ratio: 5074 / 1218,
};

function spaced(s) {
  return s
    .toUpperCase()
    .split(" ")
    .map((w) => w.split("").join(" "))
    .join("   ");
}

function t(slide, str, o) {
  slide.addText(str, { fontFace: F.body, color: C.text, isTextBox: true, margin: 0, valign: "top", ...o });
}

function box(slide, x, y, w, h, fill = C.card) {
  slide.addShape("rect", { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } });
}

function logo(slide, which, x, y, h) {
  slide.addImage({ path: LOGO[which], x, y, h, w: h * LOGO.ratio });
}

// Standard content page: dark background, title, optional kicker, source line,
// page number and a small logo bottom right.
function page(pres, s, n, deck) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  const size = s.title.length > 46 ? 18 : s.title.length > 36 ? 20 : 23;
  t(slide, s.title.toUpperCase(), { x: MX, y: 0.38, w: CW, h: 0.5, fontFace: F.head, fontSize: size, color: C.white, valign: "middle" });
  if (s.kicker) {
    t(slide, s.kicker.toUpperCase(), { x: MX, y: 0.9, w: CW, h: 0.26, fontFace: F.semi, fontSize: 8.5, color: C.yellow, charSpacing: 1 });
  }
  if (s.source) {
    t(slide, s.source, { x: MX, y: H - 0.66, w: CW - 1.3, h: 0.26, fontSize: 6.5, color: C.dim, valign: "bottom" });
  }
  t(slide, String(n), { x: MX, y: H - 0.34, w: 0.4, h: 0.2, fontFace: F.semi, fontSize: 8, color: C.dim });
  t(slide, `Sabio × ${deck.client}  ·  ${deck.footer || "Proactive point of view"}  ·  Confidential`, {
    x: MX + 0.45, y: H - 0.34, w: 6, h: 0.2, fontSize: 7, color: C.dim,
  });
  logo(slide, "white", W - MX - 0.72, H - 0.37, 0.17);
  if (s.notes) slide.addNotes(s.notes);
  return slide;
}

const R = {};

R.cover = (pres, s, n, deck) => {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  logo(slide, "yellow", MX, 0.45, 0.34);
  t(slide, spaced(`Sabio × ${deck.client}`), { x: MX, y: 1.15, w: 6.2, h: 0.3, fontFace: F.semi, fontSize: 10.5, color: C.yellow });
  t(slide, s.title.toUpperCase(), { x: MX, y: 1.5, w: 6.2, h: 1.55, fontFace: F.head, fontSize: 28, color: C.white, valign: "middle", lineSpacingMultiple: 0.95 });
  t(slide, s.subtitle, { x: MX, y: 3.1, w: 6.0, h: 0.75, fontSize: 11.5, color: C.text });
  t(slide, s.meta, { x: MX, y: 4.0, w: 6.2, h: 0.25, fontFace: F.semi, fontSize: 9, color: C.white });
  if (s.names) t(slide, s.names, { x: MX, y: 4.3, w: 6.2, h: 0.5, fontSize: 8.5, color: C.muted });
  if (s.stat) {
    // Right-hand stat panel: the one number that makes this timely.
    box(slide, 6.95, 0.45, W - MX - 6.95, 4.72, C.card);
    t(slide, s.stat.label.toUpperCase(), { x: 7.2, y: 0.8, w: 2.1, h: 0.5, fontFace: F.semi, fontSize: 8.5, color: C.yellow, charSpacing: 1 });
    const vs = s.stat.value.length > 4 ? 40 : s.stat.value.length > 2 ? 50 : 64;
    t(slide, s.stat.value, { x: 7.2, y: 1.35, w: 2.15, h: 1.2, fontFace: F.head, fontSize: vs, color: C.yellow, valign: "middle", fit: "none" });
    t(slide, s.stat.unit.toUpperCase(), { x: 7.2, y: 2.6, w: 2.1, h: 0.5, fontFace: F.head, fontSize: 12, color: C.white });
    t(slide, s.stat.body, { x: 7.2, y: 3.25, w: 2.05, h: 1.75, fontSize: 9, color: C.text });
  }
  t(slide, "Sabio Holdings · TSXV: SBIO | OTCQB: SABOF    Creator TV® | App Science® | Sabio", {
    x: MX, y: H - 0.34, w: 6.4, h: 0.2, fontSize: 7, color: C.dim,
  });
};

R.agenda = (pres, s, n, deck) => {
  const slide = page(pres, { title: "Agenda", source: s.source }, n, deck);
  const top = 1.2;
  const rowH = 0.52;
  s.items.forEach((it, i) => {
    const y = top + i * rowH;
    t(slide, String(i + 1).padStart(2, "0"), { x: MX, y, w: 0.6, h: rowH - 0.06, fontFace: F.head, fontSize: 18, color: C.yellow, valign: "middle" });
    t(slide, it, { x: 1.2, y, w: 7.5, h: rowH - 0.06, fontFace: F.semi, fontSize: 13, color: C.white, valign: "middle" });
    if (i < s.items.length - 1) slide.addShape("line", { x: 1.2, y: y + rowH - 0.03, w: CW - 0.7, h: 0, line: { color: C.line, width: 0.75 } });
  });
  if (s.note) t(slide, s.note, { x: 1.2, y: top + s.items.length * rowH + 0.12, w: 7.6, h: 0.4, fontSize: 9.5, italic: true, color: C.muted });
};

// What's happening at the client: dated news cards with "what it means".
R.news = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const k = s.items.length;
  const gap = 0.16;
  const w = (CW - gap * (k - 1)) / k;
  const y = 1.3;
  const h = 3.55;
  s.items.forEach((it, i) => {
    const x = MX + i * (w + gap);
    box(slide, x, y, w, h);
    t(slide, it.date.toUpperCase(), { x: x + 0.16, y: y + 0.16, w: w - 0.32, h: 0.22, fontFace: F.semi, fontSize: 8, color: C.yellow, charSpacing: 1 });
    t(slide, it.head, { x: x + 0.16, y: y + 0.42, w: w - 0.32, h: 0.72, fontFace: F.semi, fontSize: 10.5, color: C.white });
    t(slide, it.body, { x: x + 0.16, y: y + 1.17, w: w - 0.32, h: 1.15, fontSize: 8, color: C.muted });
    box(slide, x + 0.12, y + 2.38, w - 0.24, 1.05, C.card2);
    t(slide, "WHAT IT MEANS", { x: x + 0.22, y: y + 2.46, w: w - 0.44, h: 0.18, fontFace: F.semi, fontSize: 7, color: C.yellow, charSpacing: 1 });
    t(slide, it.means, { x: x + 0.22, y: y + 2.66, w: w - 0.44, h: 0.74, fontSize: 8, color: C.white });
  });
};

// The opening: one big number, the argument, and three proof points.
R.opening = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  box(slide, MX, 1.3, 3.6, 3.55, C.card);
  t(slide, s.stat, { x: MX + 0.25, y: 1.5, w: 3.2, h: 1.3, fontFace: F.head, fontSize: 54, color: C.yellow, valign: "middle" });
  t(slide, s.statLabel, { x: MX + 0.25, y: 2.85, w: 3.1, h: 1.0, fontFace: F.semi, fontSize: 11, color: C.white });
  if (s.statSource) t(slide, s.statSource, { x: MX + 0.25, y: 4.35, w: 3.1, h: 0.4, fontSize: 7, color: C.dim, valign: "bottom" });
  t(slide, s.headline, { x: 4.4, y: 1.3, w: 5.1, h: 0.9, fontFace: F.head, fontSize: 16, color: C.white });
  s.points.forEach((p, i) => {
    const y = 2.3 + i * 0.85;
    t(slide, p.head.toUpperCase(), { x: 4.4, y, w: 5.1, h: 0.22, fontFace: F.semi, fontSize: 8.5, color: C.yellow, charSpacing: 1 });
    t(slide, p.body, { x: 4.4, y: y + 0.25, w: 5.1, h: 0.55, fontSize: 9.5, color: C.text });
  });
};

// Sabio at a glance: five stats, three pillars, one takeaway band.
R.glance = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const stats = [
    ["2014", "Founded in LA by mobile & TV veterans"],
    ["280MM", "Mobile devices observed by App Science®"],
    ["115MM", "CTV devices across 80MM validated households"],
    ["70%", "Of US streaming homes reached"],
    ["SBIO", "TSXV listed · NMSDC-certified MBE"],
  ];
  const sw = (CW - 0.12 * 4) / 5;
  stats.forEach(([v, l], i) => {
    const x = MX + i * (sw + 0.12);
    box(slide, x, 1.28, sw, 1.12);
    t(slide, v, { x: x + 0.12, y: 1.36, w: sw - 0.24, h: 0.5, fontFace: F.head, fontSize: 22, color: C.yellow, valign: "middle" });
    t(slide, l, { x: x + 0.12, y: 1.88, w: sw - 0.24, h: 0.48, fontSize: 7.5, color: C.muted });
  });
  const pillars = [
    ["Our own data", "App Science® reads real on-device app ownership and use, opened in the last 30 days. First-party signal, not licensed third-party segments or panel inference."],
    ["Our own tech", "Sabio DSP, SSP and ad server built for CTV. The same stack powers always-on PMP Deal IDs and managed IO, so we can commit supply, not resell it."],
    ["Our own content", "Creator TV®, our owned-and-operated streaming channel, puts creator-led spots and integrations next to the audience buy."],
  ];
  const pw = (CW - 0.16 * 2) / 3;
  pillars.forEach(([h, b], i) => {
    const x = MX + i * (pw + 0.16);
    t(slide, h.toUpperCase(), { x, y: 2.6, w: pw, h: 0.24, fontFace: F.semi, fontSize: 9, color: C.yellow, charSpacing: 1 });
    t(slide, b, { x, y: 2.88, w: pw, h: 0.95, fontSize: 8.5, color: C.text });
  });
  box(slide, MX, 3.95, CW, 0.62, C.yellow);
  t(slide, [
    { text: `WHAT IT MEANS FOR ${deck.client.toUpperCase()}:  `, options: { fontFace: F.head, color: C.ink } },
    { text: s.takeaway, options: { color: C.ink } },
  ], { x: MX + 0.2, y: 3.95, w: CW - 0.4, h: 0.62, fontSize: 9.5, valign: "middle" });
};

// Three core audiences: tag / name / desc / app ecosystem / role / reach.
R.audiences = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const gap = 0.16;
  const w = (CW - gap * 2) / 3;
  const y = 1.3;
  const h = 3.55;
  s.cards.forEach((c, i) => {
    const x = MX + i * (w + gap);
    box(slide, x, y, w, h);
    t(slide, c.tag.toUpperCase(), { x: x + 0.18, y: y + 0.16, w: w - 0.36, h: 0.2, fontFace: F.semi, fontSize: 7.5, color: C.yellow, charSpacing: 1 });
    t(slide, c.name.toUpperCase(), { x: x + 0.18, y: y + 0.4, w: w - 0.36, h: 0.45, fontFace: F.head, fontSize: 12, color: C.white, valign: "middle" });
    t(slide, c.desc, { x: x + 0.18, y: y + 0.9, w: w - 0.36, h: 0.6, fontSize: 8.5, color: C.text });
    t(slide, "APP ECOSYSTEM", { x: x + 0.18, y: y + 1.55, w: w - 0.36, h: 0.18, fontFace: F.semi, fontSize: 7, color: C.dim, charSpacing: 1 });
    t(slide, c.apps, { x: x + 0.18, y: y + 1.75, w: w - 0.36, h: 0.55, fontFace: F.semi, fontSize: 8.5, color: C.white });
    t(slide, c.goal, { x: x + 0.18, y: y + 2.35, w: w - 0.36, h: 0.4, fontSize: 8.5, italic: true, color: C.yellowSoft });
    box(slide, x + 0.18, y + 2.82, w - 0.36, 0.56, C.card2);
    t(slide, "UNIQUE MONTHLY REACH  [CONFIRM]", { x: x + 0.28, y: y + 2.87, w: w - 0.56, h: 0.18, fontFace: F.semi, fontSize: 6.5, color: C.yellow, charSpacing: 1 });
    t(slide, c.reach, { x: x + 0.28, y: y + 3.07, w: w - 0.56, h: 0.26, fontFace: F.semi, fontSize: 8, color: C.white });
  });
};

// Deduplicated reach build as a native column chart plus the headline total.
R.reach = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  box(slide, MX, 1.3, 5.6, 3.55);
  slide.addChart("bar", [{ name: "Net-new CTV households (MM)", labels: s.bars.map((b) => b.label), values: s.bars.map((b) => b.value) }], {
    x: MX + 0.1, y: 1.4, w: 5.4, h: 3.35, barDir: "col", barGapWidthPct: 70,
    chartColors: [C.yellow], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.white, dataLabelFontSize: 11,
    dataLabelFontFace: F.semi, dataLabelFormatCode: '"+"0"MM"', catAxisLabelColor: C.muted, catAxisLabelFontSize: 8.5,
    catAxisLabelFontFace: F.body, catAxisLineShow: false, valAxisHidden: true, valGridLine: { style: "none" },
    catGridLine: { style: "none" }, showLegend: false, showTitle: true, title: "Net-new CTV households by audience (MM)",
    titleFontSize: 9, titleColor: C.muted, titleFontFace: F.semi,
  });
  t(slide, s.total, { x: 6.4, y: 1.3, w: 3.1, h: 0.95, fontFace: F.head, fontSize: 46, color: C.yellow, valign: "middle" });
  t(slide, s.totalLabel, { x: 6.4, y: 2.3, w: 3.1, h: 0.5, fontFace: F.semi, fontSize: 10.5, color: C.white });
  t(slide, s.body, { x: 6.4, y: 2.95, w: 3.1, h: 1.9, fontSize: 9, color: C.text });
};

// Activation: channel cards left, what-you-get panel right.
R.activation = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const cw = 2.6;
  s.channels.forEach((c, i) => {
    const x = MX + (i % 2) * (cw + 0.14);
    const y = 1.3 + Math.floor(i / 2) * 1.8;
    box(slide, x, y, cw, 1.68);
    t(slide, c.head.toUpperCase(), { x: x + 0.18, y: y + 0.18, w: cw - 0.36, h: 0.25, fontFace: F.head, fontSize: 11, color: C.yellow });
    t(slide, c.body, { x: x + 0.18, y: y + 0.52, w: cw - 0.36, h: 1.05, fontSize: 8.5, color: C.text });
  });
  const rx = MX + 2 * cw + 0.3;
  const rw = W - MX - rx;
  box(slide, rx, 1.3, rw, 3.48, C.yellow);
  t(slide, s.getsHead.toUpperCase(), { x: rx + 0.22, y: 1.5, w: rw - 0.44, h: 0.3, fontFace: F.head, fontSize: 12, color: C.ink });
  t(slide, s.gets.map((g, i) => ({ text: g, options: { bullet: { indent: 12 }, breakLine: i < s.gets.length - 1 } })), {
    x: rx + 0.22, y: 1.92, w: rw - 0.44, h: 2.7, fontSize: 9.5, color: C.ink, paraSpaceAfter: 7,
  });
};

// Three-moment flow with chevrons.
R.flow = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const k = s.steps.length;
  const gap = 0.4;
  const w = (CW - gap * (k - 1)) / k;
  const y = 1.3;
  const h = 2.75;
  s.steps.forEach((st, i) => {
    const x = MX + i * (w + gap);
    box(slide, x, y, w, h);
    t(slide, st.tag.toUpperCase(), { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.22, fontFace: F.semi, fontSize: 8.5, color: C.yellow, charSpacing: 1 });
    t(slide, st.head, { x: x + 0.2, y: y + 0.5, w: w - 0.4, h: 0.65, fontFace: F.head, fontSize: 13, color: C.white });
    t(slide, st.body, { x: x + 0.2, y: y + 1.25, w: w - 0.4, h: h - 1.4, fontSize: 9, color: C.text });
    if (i < k - 1) {
      t(slide, "›", { x: x + w, y: y + h / 2 - 0.35, w: gap, h: 0.7, fontFace: F.head, fontSize: 30, color: C.yellow, align: "center", valign: "middle" });
    }
  });
  if (s.benchmark) t(slide, s.benchmark, { x: MX, y: 4.25, w: CW, h: 0.25, fontFace: F.semi, fontSize: 8.5, color: C.muted });
};

// Two measurement paths, four steps each.
R.paths = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  s.paths.forEach((p, i) => {
    const y = 1.28 + i * 1.78;
    t(slide, p.head.toUpperCase(), { x: MX, y, w: CW, h: 0.24, fontFace: F.semi, fontSize: 9, color: C.white, charSpacing: 1 });
    const sw = (CW - 0.12 * 3) / 4;
    p.steps.forEach((st, j) => {
      const x = MX + j * (sw + 0.12);
      const last = j === 3;
      box(slide, x, y + 0.32, sw, 1.28, last ? C.yellow : C.card);
      t(slide, `0${j + 1}`, { x: x + 0.14, y: y + 0.42, w: 0.6, h: 0.3, fontFace: F.head, fontSize: 13, color: last ? C.ink : C.yellow });
      t(slide, st, { x: x + 0.14, y: y + 0.76, w: sw - 0.28, h: 0.8, fontSize: 8.5, color: last ? C.ink : C.text });
    });
  });
};

// Proof: stat rows left, "why now" panel right.
R.proof = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  s.stats.forEach((st, i) => {
    const y = 1.3 + i * 1.2;
    box(slide, MX, y, 5.5, 1.08);
    t(slide, st.value, { x: MX + 0.2, y, w: 1.9, h: 1.08, fontFace: F.head, fontSize: 28, color: C.yellow, valign: "middle" });
    t(slide, st.label, { x: MX + 2.15, y: y + 0.17, w: 3.2, h: 0.3, fontFace: F.semi, fontSize: 10, color: C.white });
    t(slide, st.context, { x: MX + 2.15, y: y + 0.5, w: 3.2, h: 0.5, fontSize: 8, color: C.muted });
  });
  const rx = MX + 5.7;
  box(slide, rx, 1.3, W - MX - rx, 3.48, C.card2);
  t(slide, s.sideHead.toUpperCase(), { x: rx + 0.22, y: 1.5, w: W - MX - rx - 0.44, h: 0.3, fontFace: F.head, fontSize: 12, color: C.yellow });
  t(slide, s.side, { x: rx + 0.22, y: 1.92, w: W - MX - rx - 0.44, h: 2.75, fontSize: 9.5, color: C.white, paraSpaceAfter: 6 });
};

// 2x3 grid of guardrails / considerations.
R.grid = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const cols = 3;
  const gap = 0.14;
  const w = (CW - gap * (cols - 1)) / cols;
  const rows = Math.ceil(s.cells.length / cols);
  const h = (3.5 - gap * (rows - 1)) / rows;
  s.cells.forEach((c, i) => {
    const x = MX + (i % cols) * (w + gap);
    const y = 1.3 + Math.floor(i / cols) * (h + gap);
    box(slide, x, y, w, h);
    t(slide, c.head.toUpperCase(), { x: x + 0.18, y: y + 0.16, w: w - 0.36, h: 0.22, fontFace: F.semi, fontSize: 8.5, color: C.yellow, charSpacing: 1 });
    t(slide, c.body, { x: x + 0.18, y: y + 0.44, w: w - 0.36, h: h - 0.55, fontSize: 8.5, color: C.text });
  });
};

// Where we go next: three numbered moves with owner and timing, as in the Kinesso deck.
R.moves = (pres, s, n, deck) => {
  const slide = page(pres, s, n, deck);
  const gap = 0.16;
  const w = (CW - gap * 2) / 3;
  const y = 1.3;
  const h = 3.0;
  s.moves.forEach((m, i) => {
    const x = MX + i * (w + gap);
    box(slide, x, y, w, h);
    t(slide, String(i + 1), { x: x + 0.18, y: y + 0.12, w: 0.6, h: 0.6, fontFace: F.head, fontSize: 30, color: C.yellow, valign: "middle" });
    t(slide, m.head.toUpperCase(), { x: x + 0.18, y: y + 0.78, w: w - 0.36, h: 0.45, fontFace: F.head, fontSize: 11.5, color: C.white });
    t(slide, m.body, { x: x + 0.18, y: y + 1.27, w: w - 0.36, h: 1.15, fontSize: 8.5, color: C.text });
    box(slide, x + 0.18, y + 2.45, w - 0.36, 0.42, C.card2);
    t(slide, m.owner, { x: x + 0.28, y: y + 2.45, w: w - 0.56, h: 0.42, fontFace: F.semi, fontSize: 7.5, color: C.yellowSoft, valign: "middle" });
  });
  if (s.closer) t(slide, s.closer, { x: MX, y: 4.48, w: CW, h: 0.32, fontFace: F.semi, fontSize: 9.5, color: C.white });
};

R.thanks = (pres, s, n, deck) => {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  logo(slide, "yellow", MX, 0.5, 0.42);
  t(slide, "THANK YOU!", { x: MX, y: 1.3, w: 9, h: 0.9, fontFace: F.head, fontSize: 44, color: C.white, valign: "middle" });
  if (s.line) t(slide, s.line, { x: MX, y: 2.2, w: 8.4, h: 0.5, fontSize: 11.5, color: C.text });
  s.names.forEach((p, i) => {
    const x = MX + i * 3.05;
    box(slide, x, 3.05, 2.9, 1.2);
    t(slide, p.name, { x: x + 0.2, y: 3.2, w: 2.5, h: 0.3, fontFace: F.head, fontSize: 12, color: C.white });
    t(slide, p.role, { x: x + 0.2, y: 3.5, w: 2.5, h: 0.25, fontFace: F.semi, fontSize: 8.5, color: C.yellow });
    t(slide, p.contact, { x: x + 0.2, y: 3.78, w: 2.5, h: 0.4, fontSize: 8.5, color: C.text });
  });
  t(slide, "Sabio Holdings · TSXV: SBIO | OTCQB: SABOF    Creator TV® | App Science® | Sabio", {
    x: MX, y: H - 0.34, w: 9, h: 0.2, fontSize: 7, color: C.dim,
  });
};

// Appendix: numbered sources for every market claim in the deck.
R.sources = (pres, s, n, deck) => {
  const slide = page(pres, { title: "Sources", kicker: `Market claims in this deck, as reported · Retrieved ${s.retrieved}` }, n, deck);
  const half = Math.ceil(s.items.length / 2);
  [s.items.slice(0, half), s.items.slice(half)].forEach((col, ci) => {
    t(slide, col.map((it, i) => ({
      text: `${ci * half + i + 1}.  ${it}`,
      options: { breakLine: i < col.length - 1 },
    })), { x: MX + ci * 4.6, y: 1.3, w: 4.4, h: 3.5, fontSize: 8.5, color: C.text, paraSpaceAfter: 8 });
  });
};

function build(spec, out) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = spec.author || "Ben Bring, Sabio";
  pres.company = "Sabio Holdings";
  pres.title = spec.fileTitle || `Sabio x ${spec.client}`;
  spec.slides.forEach((s, i) => {
    const fn = R[s.type];
    if (!fn) throw new Error(`Unknown slide type "${s.type}" on slide ${i + 1}`);
    fn(pres, s, i + 1, spec);
  });
  // pptxgenjs stores every part uncompressed (~5x larger) whatever the
  // compression flag says, so re-zip with DEFLATE before writing.
  return pres
    .write({ outputType: "nodebuffer" })
    .then((buf) => JSZip.loadAsync(buf))
    .then((zip) => zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 9 } }))
    .then((buf) => {
      fs.writeFileSync(out, buf);
      return out;
    });
}

if (require.main === module) {
  const [specPath, out] = process.argv.slice(2);
  if (!specPath || !out) {
    console.error("usage: node lib/sabio-deck.js spec.json out.pptx");
    process.exit(1);
  }
  const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
  build(spec, path.resolve(out)).then((f) => console.log(`wrote ${f}`));
}

module.exports = { build };
