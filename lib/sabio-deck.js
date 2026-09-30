// Sabio proactive strategy deck renderer.
// Takes a JSON spec (see specs in weeks/*/specs) and renders a 16:9 .pptx that
// follows the structure of the Sabio x Walmart Sparky deck: short answer,
// Audience, Activation & Scale, Measurement & Privacy, Plan & Next Steps.
//
// Usage: node lib/sabio-deck.js path/to/spec.json path/to/out.pptx

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");

// Sabio palette (sabio-rfp brand-style.md). Navy dominates, orange is the accent.
const C = {
  navy: "1F2A44",
  blue: "2E5C8A",
  ice: "DDEBF7",
  orange: "E8703A",
  grey: "6B7280",
  panel: "F4F6F9",
  white: "FFFFFF",
  ink: "1F2937",
};
const FONT = "Arial";
const W = 10;
const H = 5.625;
const MX = 0.5; // side margin

function spaced(s) {
  // "SABIO" -> "S A B I O", keeping word gaps wide, as on the Walmart cover.
  return s
    .toUpperCase()
    .split(" ")
    .map((w) => w.split("").join(" "))
    .join("   ");
}

function text(slide, str, opts) {
  slide.addText(str, { fontFace: FONT, isTextBox: true, margin: 0, ...opts });
}

function kicker(slide, str, y = 0.95) {
  text(slide, str.toUpperCase(), {
    x: MX, y, w: W - 2 * MX, h: 0.3, fontSize: 9.5, color: C.orange, bold: true, charSpacing: 1,
  });
}

function title(slide, str) {
  // Long titles step down a size so they stay on one line above the kicker.
  const size = str.length > 44 ? 17 : str.length > 36 ? 19 : 22;
  text(slide, str.toUpperCase(), {
    x: MX, y: 0.4, w: W - 2 * MX, h: 0.55, fontSize: size, bold: true, color: C.navy, valign: "top",
  });
}

function footer(slide, n, footnote) {
  text(slide, String(n), { x: MX, y: H - 0.38, w: 0.4, h: 0.25, fontSize: 9, color: C.grey });
  if (footnote) {
    text(slide, footnote, {
      x: MX + 0.45, y: H - 0.42, w: W - 2 * MX - 0.45, h: 0.32, fontSize: 7.5, color: C.grey, valign: "middle",
    });
  }
}

function card(slide, x, y, w, h, fill = C.panel) {
  slide.addShape("rect", { x, y, w, h, fill: { color: fill }, line: { color: fill } });
}

function header(slide, s, n) {
  slide.background = { color: C.white };
  title(slide, s.title);
  if (s.kicker) kicker(slide, s.kicker);
  footer(slide, n, s.footnote);
}

const R = {};

R.cover = (pres, s, n, deck) => {
  const slide = pres.addSlide();
  slide.background = { color: C.navy };
  text(slide, spaced(`Sabio × ${deck.client}`), {
    x: MX, y: 0.55, w: 9, h: 0.35, fontSize: 12, color: C.orange, bold: true,
  });
  text(slide, s.title.toUpperCase(), {
    x: MX, y: 1.1, w: 8.6, h: 1.5, fontSize: 36, bold: true, color: C.white, valign: "top",
  });
  text(slide, s.subtitle, {
    x: MX, y: 2.75, w: 8.2, h: 0.8, fontSize: 14, color: C.ice, valign: "top",
  });
  text(slide, s.meta, { x: MX, y: 3.85, w: 9, h: 0.3, fontSize: 10.5, color: C.white, bold: true });
  if (s.names) {
    text(slide, s.names, { x: MX, y: 4.25, w: 9, h: 0.5, fontSize: 9.5, color: C.ice, valign: "top" });
  }
  text(slide, "CREATOR TV®  |  APP SCIENCE®  |  SABIO", {
    x: MX, y: H - 0.45, w: 9, h: 0.25, fontSize: 8, color: C.ice, charSpacing: 1,
  });
};

R.contents = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, { title: "Contents", footnote: s.footnote }, n);
  const top = 1.15;
  const rowH = (H - top - 0.55) / s.items.length;
  s.items.forEach((it, i) => {
    const y = top + i * rowH;
    text(slide, it.num, { x: MX, y, w: 0.7, h: rowH - 0.1, fontSize: 22, bold: true, color: C.orange, valign: "middle" });
    text(slide, it.title.toUpperCase(), { x: 1.3, y: y + 0.05, w: 6.6, h: 0.3, fontSize: 13, bold: true, color: C.navy });
    text(slide, it.desc, { x: 1.3, y: y + 0.35, w: 6.6, h: rowH - 0.45, fontSize: 10, color: C.grey, valign: "top" });
    text(slide, it.pages, { x: 8.1, y, w: 1.4, h: rowH - 0.1, fontSize: 10, color: C.blue, align: "right", valign: "middle" });
  });
};

R.table = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const hdr = s.headers.map((h) => ({
    text: h.toUpperCase(),
    options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: 9.5 },
  }));
  const rows = s.rows.map((r, i) =>
    r.map((c, j) => ({
      text: c,
      options: {
        fontSize: j === 1 ? 9.5 : 10,
        bold: j === 0,
        color: j === 0 ? C.navy : C.ink,
        fill: { color: i % 2 ? C.white : C.panel },
        align: j === 2 ? "center" : "left",
      },
    }))
  );
  slide.addTable([hdr, ...rows], {
    x: MX, y: 1.35, w: W - 2 * MX, colW: [2.6, 5.4, 1.0], fontFace: FONT, valign: "middle",
    border: { type: "solid", pt: 0.5, color: "E5E7EB" }, margin: 0.07,
    rowH: [0.32, ...s.rows.map(() => (H - 1.35 - 0.32 - 0.65) / s.rows.length)],
  });
};

R.divider = (pres, s, n) => {
  const slide = pres.addSlide();
  slide.background = { color: C.navy };
  text(slide, spaced(`Section ${s.num}`), { x: MX, y: 1.3, w: 9, h: 0.35, fontSize: 12, bold: true, color: C.orange });
  text(slide, s.title.toUpperCase(), { x: MX, y: 1.8, w: 9, h: 0.9, fontSize: 38, bold: true, color: C.white });
  text(slide, s.desc, { x: MX, y: 2.8, w: 8, h: 0.8, fontSize: 15, color: C.ice, valign: "top" });
  text(slide, s.pages, { x: MX, y: 4.3, w: 9, h: 0.4, fontSize: 10, color: C.ice });
};

// Up to four signal cards, label / head / body.
R.cards = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const k = s.cards.length;
  const gap = 0.2;
  const w = (W - 2 * MX - gap * (k - 1)) / k;
  const y = 1.4;
  const h = H - y - 0.65;
  s.cards.forEach((c, i) => {
    const x = MX + i * (w + gap);
    card(slide, x, y, w, h);
    text(slide, c.label.toUpperCase(), { x: x + 0.18, y: y + 0.18, w: w - 0.36, h: 0.3, fontSize: 11, bold: true, color: C.orange });
    text(slide, c.head, { x: x + 0.18, y: y + 0.55, w: w - 0.36, h: 0.8, fontSize: 11, bold: true, color: C.navy, valign: "top" });
    text(slide, c.body, { x: x + 0.18, y: y + 1.45, w: w - 0.36, h: h - 1.6, fontSize: 10, color: C.ink, valign: "top" });
  });
};

// Four-step ladder with two headline stats underneath.
R.ladder = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const k = s.steps.length;
  const gap = 0.12;
  const w = (W - 2 * MX - gap * (k - 1)) / k;
  const y = 1.4;
  const h = 1.75;
  s.steps.forEach((st, i) => {
    const x = MX + i * (w + gap);
    const hot = i >= k - 2;
    card(slide, x, y, w, h, hot ? C.navy : C.panel);
    text(slide, `${st.num}  ${st.label}`.toUpperCase(), {
      x: x + 0.15, y: y + 0.15, w: w - 0.3, h: 0.25, fontSize: 9, bold: true, color: C.orange,
    });
    text(slide, st.head, {
      x: x + 0.15, y: y + 0.45, w: w - 0.3, h: 0.75, fontSize: 12, bold: true, color: hot ? C.white : C.navy, valign: "top",
    });
    text(slide, st.sub, {
      x: x + 0.15, y: y + 1.2, w: w - 0.3, h: 0.45, fontSize: 9, color: hot ? C.ice : C.grey, valign: "top",
    });
  });
  const sy = y + h + 0.3;
  s.stats.forEach((st, i) => {
    const x = MX + i * 4.6;
    text(slide, st.value, { x, y: sy, w: 1.9, h: 0.75, fontSize: 36, bold: true, color: C.orange, valign: "middle" });
    text(slide, st.label, { x: x + 1.95, y: sy, w: 2.5, h: 0.75, fontSize: 11, color: C.navy, valign: "middle" });
  });
};

// Three columns of named apps with illustrative device counts.
R.lists = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const k = s.columns.length;
  const gap = 0.25;
  const w = (W - 2 * MX - gap * (k - 1)) / k;
  const y = 1.4;
  const h = H - y - 0.65;
  s.columns.forEach((col, i) => {
    const x = MX + i * (w + gap);
    card(slide, x, y, w, h);
    text(slide, col.head.toUpperCase(), { x: x + 0.18, y: y + 0.15, w: w - 0.36, h: 0.3, fontSize: 11, bold: true, color: C.navy });
    text(slide, col.sub, { x: x + 0.18, y: y + 0.45, w: w - 0.36, h: 0.45, fontSize: 9, color: C.grey, valign: "top" });
    col.items.forEach(([name, count], j) => {
      const iy = y + 1.0 + j * 0.34;
      text(slide, name, { x: x + 0.18, y: iy, w: w - 1.1, h: 0.3, fontSize: 10.5, color: C.ink });
      if (count) {
        text(slide, count, { x: x + w - 0.95, y: iy, w: 0.77, h: 0.3, fontSize: 10.5, bold: true, color: C.blue, align: "right" });
      }
    });
  });
};

// Three core audiences: tag / name / description / app ecosystem / goal / reach.
R.audiences = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const gap = 0.2;
  const w = (W - 2 * MX - gap * 2) / 3;
  const y = 1.35;
  const h = H - y - 0.62;
  s.cards.forEach((c, i) => {
    const x = MX + i * (w + gap);
    card(slide, x, y, w, h);
    slide.addShape("rect", { x, y, w, h: 0.66, fill: { color: C.navy }, line: { color: C.navy } });
    text(slide, c.tag.toUpperCase(), { x: x + 0.15, y: y + 0.08, w: w - 0.3, h: 0.22, fontSize: 9, bold: true, color: C.orange });
    text(slide, c.name.toUpperCase(), { x: x + 0.15, y: y + 0.3, w: w - 0.3, h: 0.32, fontSize: c.name.length > 22 ? 10 : 11.5, bold: true, color: C.white });
    text(slide, c.desc, { x: x + 0.15, y: y + 0.72, w: w - 0.3, h: 0.5, fontSize: 9.5, color: C.ink, valign: "top" });
    text(slide, "APP ECOSYSTEM", { x: x + 0.15, y: y + 1.27, w: w - 0.3, h: 0.2, fontSize: 8, bold: true, color: C.orange });
    text(slide, c.apps, { x: x + 0.15, y: y + 1.47, w: w - 0.3, h: 0.55, fontSize: 9.5, color: C.navy, bold: true, valign: "top" });
    text(slide, c.goal, { x: x + 0.15, y: y + 2.05, w: w - 0.3, h: 0.35, fontSize: 9.5, italic: true, color: C.blue, valign: "top" });
    text(slide, "MONTHLY REACH  [CONFIRM]", { x: x + 0.15, y: y + 2.45, w: w - 0.3, h: 0.2, fontSize: 8, bold: true, color: C.orange });
    text(slide, c.reach, { x: x + 0.15, y: y + 2.65, w: w - 0.3, h: 0.3, fontSize: 10, bold: true, color: C.navy });
  });
};

// Deduplicated reach build as a native column chart plus a big total.
R.reach = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  slide.addChart("bar", [{ name: "Net-new CTV households (MM)", labels: s.bars.map((b) => b.label), values: s.bars.map((b) => b.value) }], {
    x: MX, y: 1.3, w: 5.4, h: 3.5, barDir: "col",
    chartColors: [C.blue], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.navy, dataLabelFontSize: 11,
    dataLabelFormatCode: '"+"0"MM"', catAxisLabelColor: C.grey, catAxisLabelFontSize: 9, valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "Net-new CTV households by audience (MM)", titleFontSize: 10, titleColor: C.navy,
  });
  text(slide, s.total, { x: 6.2, y: 1.4, w: 3.3, h: 0.9, fontSize: 44, bold: true, color: C.orange });
  text(slide, s.totalLabel, { x: 6.2, y: 2.3, w: 3.3, h: 0.5, fontSize: 11, bold: true, color: C.navy, valign: "top" });
  text(slide, s.body, { x: 6.2, y: 2.9, w: 3.3, h: 1.8, fontSize: 10, color: C.ink, valign: "top" });
};

// Two-column split: channel cards left, "what you get" list right.
R.activation = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const y = 1.4;
  const cw = 2.55;
  s.channels.forEach((c, i) => {
    const x = MX + (i % 2) * (cw + 0.15);
    const cy = y + Math.floor(i / 2) * 1.6;
    card(slide, x, cy, cw, 1.45);
    text(slide, c.head.toUpperCase(), { x: x + 0.15, y: cy + 0.15, w: cw - 0.3, h: 0.25, fontSize: 11, bold: true, color: C.orange });
    text(slide, c.body, { x: x + 0.15, y: cy + 0.5, w: cw - 0.3, h: 0.85, fontSize: 9.5, color: C.ink, valign: "top" });
  });
  const rx = MX + 2 * cw + 0.45;
  const rw = W - MX - rx;
  card(slide, rx, y, rw, 3.05, C.navy);
  text(slide, s.getsHead.toUpperCase(), { x: rx + 0.2, y: y + 0.2, w: rw - 0.4, h: 0.3, fontSize: 11, bold: true, color: C.orange });
  text(
    slide,
    s.gets.map((g, i) => ({ text: g, options: { bullet: true, breakLine: i < s.gets.length - 1 } })),
    { x: rx + 0.2, y: y + 0.6, w: rw - 0.4, h: 2.3, fontSize: 10.5, color: C.white, valign: "top", paraSpaceAfter: 6 }
  );
};

// Three-moment flow: tag / head / body with chevrons between.
R.flow = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const k = s.steps.length;
  const gap = 0.45;
  const w = (W - 2 * MX - gap * (k - 1)) / k;
  const y = 1.45;
  const h = 2.9;
  s.steps.forEach((st, i) => {
    const x = MX + i * (w + gap);
    card(slide, x, y, w, h);
    text(slide, st.tag.toUpperCase(), { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.25, fontSize: 10, bold: true, color: C.orange });
    text(slide, st.head, { x: x + 0.2, y: y + 0.5, w: w - 0.4, h: 0.6, fontSize: 13, bold: true, color: C.navy, valign: "top" });
    text(slide, st.body, { x: x + 0.2, y: y + 1.15, w: w - 0.4, h: h - 1.3, fontSize: 10, color: C.ink, valign: "top" });
    if (i < k - 1) {
      text(slide, "›", { x: x + w, y: y + h / 2 - 0.3, w: gap, h: 0.6, fontSize: 28, bold: true, color: C.orange, align: "center", valign: "middle" });
    }
  });
};

// Two measurement paths, each four numbered steps.
R.paths = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  s.paths.forEach((p, i) => {
    const y = 1.35 + i * 1.72;
    text(slide, p.head.toUpperCase(), { x: MX, y, w: 9, h: 0.28, fontSize: 11, bold: true, color: C.navy });
    const sw = (W - 2 * MX - 0.15 * 3) / 4;
    p.steps.forEach((st, j) => {
      const x = MX + j * (sw + 0.15);
      card(slide, x, y + 0.35, sw, 1.2, j === 3 ? C.navy : C.panel);
      text(slide, `0${j + 1}`, { x: x + 0.12, y: y + 0.45, w: 0.5, h: 0.3, fontSize: 12, bold: true, color: C.orange });
      text(slide, st, { x: x + 0.12, y: y + 0.78, w: sw - 0.24, h: 0.72, fontSize: 9.5, color: j === 3 ? C.white : C.ink, valign: "top" });
    });
  });
};

// Proof: big stat callouts on the left, context on the right.
R.proof = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  s.stats.forEach((st, i) => {
    const y = 1.35 + i * 1.05;
    card(slide, MX, y, 5.3, 0.92);
    text(slide, st.value, { x: MX + 0.15, y, w: 1.75, h: 0.92, fontSize: 28, bold: true, color: C.orange, valign: "middle" });
    text(slide, st.label, { x: MX + 1.95, y: y + 0.1, w: 3.25, h: 0.35, fontSize: 10.5, bold: true, color: C.navy });
    text(slide, st.context, { x: MX + 1.95, y: y + 0.43, w: 3.25, h: 0.45, fontSize: 9, color: C.grey, valign: "top" });
  });
  card(slide, 6.05, 1.35, W - MX - 6.05, 3.05, C.navy);
  text(slide, s.sideHead.toUpperCase(), { x: 6.25, y: 1.55, w: 3.0, h: 0.3, fontSize: 11, bold: true, color: C.orange });
  text(slide, s.side, { x: 6.25, y: 1.95, w: 3.0, h: 2.3, fontSize: 10.5, color: C.white, valign: "top" });
};

// 2x3 grid of considerations.
R.grid = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const cols = 3;
  const gap = 0.18;
  const w = (W - 2 * MX - gap * (cols - 1)) / cols;
  const rows = Math.ceil(s.cells.length / cols);
  const y0 = 1.35;
  const h = (H - y0 - 0.6 - gap * (rows - 1)) / rows;
  s.cells.forEach((c, i) => {
    const x = MX + (i % cols) * (w + gap);
    const y = y0 + Math.floor(i / cols) * (h + gap);
    card(slide, x, y, w, h);
    text(slide, c.head.toUpperCase(), { x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.25, fontSize: 10, bold: true, color: C.orange });
    text(slide, c.body, { x: x + 0.15, y: y + 0.42, w: w - 0.3, h: h - 0.5, fontSize: 9.5, color: C.ink, valign: "top" });
  });
};

// Test plan: three phases across the top, holdout split and asks below.
R.plan = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const widths = [2.3, 4.0, 2.3];
  let x = MX;
  s.phases.forEach((p, i) => {
    const w = widths[i];
    card(slide, x, 1.35, w, 1.55, i === 1 ? C.navy : C.panel);
    text(slide, p.tag.toUpperCase(), { x: x + 0.15, y: 1.47, w: w - 0.3, h: 0.25, fontSize: 9.5, bold: true, color: C.orange });
    text(slide, p.head.toUpperCase(), { x: x + 0.15, y: 1.72, w: w - 0.3, h: 0.25, fontSize: 11, bold: true, color: i === 1 ? C.white : C.navy });
    text(slide, p.body, { x: x + 0.15, y: 2.02, w: w - 0.3, h: 0.85, fontSize: 9, color: i === 1 ? C.ice : C.ink, valign: "top" });
    x += w + 0.2;
  });
  // Holdout split as a native doughnut chart.
  slide.addChart("doughnut", [{ name: "Split", labels: ["Exposed", "Holdout"], values: [s.split[0], s.split[1]] }], {
    x: MX, y: 3.05, w: 1.9, h: 1.9, chartColors: [C.blue, C.orange], holeSize: 60, showLegend: false,
    showValue: false, showPercent: true, dataLabelColor: C.white, dataLabelFontSize: 9,
  });
  text(slide, s.splitCaption, { x: 2.45, y: 3.25, w: 2.3, h: 1.4, fontSize: 9.5, color: C.ink, valign: "top" });
  card(slide, 5.0, 3.1, W - MX - 5.0, 1.8);
  text(slide, "WHAT WE NEED TO START", { x: 5.15, y: 3.2, w: 4.2, h: 0.25, fontSize: 9.5, bold: true, color: C.orange });
  text(
    slide,
    s.asks.map((a, i) => [
      { text: `${a.head}  `, options: { bold: true, color: C.navy } },
      { text: a.body, options: { color: C.ink, breakLine: i < s.asks.length - 1 } },
    ]).flat(),
    { x: 5.15, y: 3.5, w: 4.2, h: 1.35, fontSize: 9.5, valign: "top", paraSpaceAfter: 4 }
  );
};

R.nextsteps = (pres, s, n) => {
  const slide = pres.addSlide();
  header(slide, s, n);
  const gap = 0.2;
  const h = (H - 1.2 - 0.6 - gap * (s.items.length - 1)) / s.items.length;
  s.items.forEach((it, i) => {
    const y = 1.2 + i * (h + gap);
    card(slide, MX, y, W - 2 * MX, h);
    text(slide, it.tag.toUpperCase(), { x: MX + 0.25, y, w: 1.6, h, fontSize: 13, bold: true, color: C.orange, valign: "middle" });
    text(slide, it.head, { x: 2.3, y: y + h / 2 - 0.42, w: 7.0, h: 0.36, fontSize: 14, bold: true, color: C.navy });
    text(slide, it.body, { x: 2.3, y: y + h / 2 - 0.02, w: 7.0, h: 0.5, fontSize: 11, color: C.ink, valign: "top" });
  });
};

R.thanks = (pres, s, n) => {
  const slide = pres.addSlide();
  slide.background = { color: C.navy };
  text(slide, "THANK YOU", { x: MX, y: 0.9, w: 9, h: 0.9, fontSize: 40, bold: true, color: C.white });
  if (s.line) text(slide, s.line, { x: MX, y: 1.85, w: 8.5, h: 0.5, fontSize: 13, color: C.ice, valign: "top" });
  s.names.forEach((p, i) => {
    const x = MX + i * 3.0;
    text(slide, p.name, { x, y: 2.9, w: 2.9, h: 0.3, fontSize: 13, bold: true, color: C.white });
    text(slide, p.role, { x, y: 3.2, w: 2.9, h: 0.25, fontSize: 10, color: C.orange });
    text(slide, p.contact, { x, y: 3.45, w: 2.9, h: 0.45, fontSize: 9.5, color: C.ice, valign: "top" });
  });
  text(slide, "Sabio Holdings · TSXV: SBIO | OTCQB: SABOF    CREATOR TV®  |  APP SCIENCE®  |  SABIO", {
    x: MX, y: H - 0.45, w: 9, h: 0.25, fontSize: 8, color: C.ice,
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
    if (s.notes) pres.slides[pres.slides.length - 1].addNotes(s.notes);
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
