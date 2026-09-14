#!/usr/bin/env node
/**
 * Sparks pitch deck generator.
 *
 *   node src/build_deck.js prospects/<slug>.yml [-o decks/<name>.pptx]
 *
 * Reads brand/config.yml, src/platforms.yml, src/case_studies.yml and one prospect
 * file; writes a platform-by-platform .pptx (Slidebean Airbnb structure adapted to an
 * agency sales pitch). The focus platform (prospect.focus, default weedmaps) gets a
 * full section; the other platforms get one preview slide each, in the order
 * Weedmaps → Leafly → Yelp → Google.
 */
const fs = require("fs");
const path = require("path");
const yaml = require("js-yaml");
const pptxgen = require("pptxgenjs");
const { icon } = require("./icons");

const ROOT = path.resolve(__dirname, "..");
const args = process.argv.slice(2);
if (!args[0]) {
  console.error("usage: node src/build_deck.js prospects/<slug>.yml [-o out.pptx]");
  process.exit(1);
}
const prospectPath = path.resolve(args[0]);
const outIdx = args.indexOf("-o");

const brand = yaml.load(fs.readFileSync(path.join(ROOT, "brand/config.yml"), "utf8"));
const caseStudies = yaml.load(fs.readFileSync(path.join(ROOT, "src/case_studies.yml"), "utf8"));
const PLAT = yaml.load(fs.readFileSync(path.join(ROOT, "src/platforms.yml"), "utf8"));
const P = yaml.load(fs.readFileSync(prospectPath, "utf8"));

const slug = path.basename(prospectPath, ".yml").replace(/^_/, "");
const outPath =
  outIdx > -1 ? path.resolve(args[outIdx + 1]) : path.join(ROOT, "decks", `Sparks-x-${slug}.pptx`);

// ───────────────────────────── config ─────────────────────────────
const C = brand.palette;
const FONT_H = "Arial";
const FONT_B = "Calibri";
const W = 13.333;
const H = 7.5;
const M = 0.6; // margin

const pros = P.prospect || {};
const prospectName = pros.name || "Your Dispensary";
const place = [pros.city, pros.state].filter(Boolean).join(", ");
const today =
  pros.date ||
  new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
const typeWord = pros.type_label || { dispensary: "dispensary", delivery: "delivery service", brand: "brand" }[pros.type] || "dispensary";

const SV = P.services || {};
const ORDER = PLAT.order || ["weedmaps", "leafly", "yelp", "google"];
const focusKey = ORDER.includes(pros.focus) ? pros.focus : "weedmaps";
const focus = PLAT[focusKey];
const others = ORDER.filter((k) => k !== focusKey);
const isIncluded = (k) => !!(SV[k] && SV[k].include);
const inv = P.investment || {};
const addonPrice = inv.addon_price || "$495/mo";
// Bundle mode: investment.addon = { price, detail } prices all non-focus platforms together.
const bundle = inv.addon && inv.addon.price ? inv.addon : null;
const total = inv.total && inv.total.price ? inv.total : null;
const trial = typeof inv.trial === "string" && inv.trial.trim() ? inv.trial.trim() : null;
const listWords = (a) => (a.length <= 1 ? a.join("") : a.slice(0, -1).join(", ") + (a.length > 2 ? "," : "") + " and " + a[a.length - 1]);
const includedNames = ORDER.filter(isIncluded).map((k) => PLAT[k].name);
const otherNames = () => ORDER.filter((k) => k !== focusKey).map((k) => PLAT[k].name);
const bundleLabel = () => (bundle && bundle.label) || `${otherNames().join(" + ")} bundle`;
const tileAddonLabel = bundle ? `ADD-ON BUNDLE ${bundle.price}` : `ADD-ON ${addonPrice}`;
const scopeList = listWords(includedNames.length ? includedNames : [focus.name]);
const oneLiner = `We manage your ${scopeList} listings and ads so your ${typeWord} gets found, clicked, and ordered from.`;

// audit rows grouped by platform key (row.platform starts with the platform name)
const auditRows = Array.isArray(P.audit) ? P.audit.filter((r) => r && (r.finding || r.impact)) : [];
const auditFor = (k) => auditRows.filter((r) => String(r.platform || "").toLowerCase().startsWith(PLAT[k].name.toLowerCase()));

const LOGO = {};
for (const [k, rel] of Object.entries(brand.logo || {})) {
  const f = path.join(ROOT, rel);
  if (fs.existsSync(f)) LOGO[k] = "image/png;base64," + fs.readFileSync(f).toString("base64");
}
const AR = { full: 923 / 696, wordmark: 908 / 607, icon: 227 / 298 };
const logoW = (key, w) => ({ data: LOGO[key], w, h: w / AR[key.split("_")[0]] });

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = brand.company;
pres.company = brand.company;
pres.title = `${brand.company} × ${prospectName}`;

let slideNo = 0;

// ───────────────────────────── helpers ─────────────────────────────
function base(dark = false) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.primary : "FFFFFF" };
  slideNo += 1;
  if (dark) {
    for (const y of [0.32, H - 0.62]) {
      s.addShape(pres.ShapeType.line, { x: M, y, w: W - 2 * M, h: 0, line: { color: "FFFFFF", width: 1 } });
    }
  }
  s.addText(brand.website, {
    x: M, y: H - 0.45, w: 6, h: 0.25, fontFace: FONT_B, fontSize: 10,
    color: dark ? C.accent : C.muted, isTextBox: true, margin: 0, valign: "middle",
  });
  const iconKey = dark ? "icon_white" : "icon_orange";
  if (LOGO[iconKey]) {
    const iw = 0.26;
    s.addImage({ ...logoW(iconKey, iw), x: W - M - iw, y: H - 0.55 });
  }
  s.addText(String(slideNo), {
    x: W - M - 1.4, y: H - 0.45, w: 1.0, h: 0.25, fontFace: FONT_B, fontSize: 10, align: "right",
    color: dark ? C.accent : C.muted, isTextBox: true, margin: 0, valign: "middle",
  });
  return s;
}

// Eyebrow (section label) + title + subtitle. Titles shrink to stay on one line where
// possible; a title that still wraps pushes the subtitle down. Returns the y where
// content can start.
function title(s, text, opts = {}) {
  const dark = !!opts.dark;
  const maxW = W - 2 * M;
  let y = 0.55;
  if (opts.eyebrow) {
    s.addText(opts.eyebrow.toUpperCase(), {
      x: M, y, w: maxW, h: 0.3, fontFace: FONT_B, fontSize: 11, bold: true, charSpacing: 4,
      color: dark ? C.accent : C.primary, isTextBox: true, margin: 0, valign: "top",
    });
    y += 0.36;
  }
  let size = 34;
  const widthAt = (pt) => text.length * pt * 0.6 / 72;
  while (size > 24 && widthAt(size) > maxW) size -= 2;
  const lines = Math.max(1, Math.ceil(widthAt(size) / maxW));
  const th = lines * size * 1.25 / 72 + 0.1;
  s.addText(text, {
    x: M, y, w: maxW, h: th, fontFace: FONT_H, fontSize: size, bold: true,
    color: dark ? "FFFFFF" : C.primary, isTextBox: true, margin: 0, valign: "top",
  });
  y += th + 0.1;
  if (opts.sub) {
    const subLines = Math.max(1, Math.ceil(opts.sub.length * 15 * 0.5 / 72 / maxW));
    const sh = subLines * 0.28 + 0.1;
    s.addText(opts.sub, {
      x: M, y, w: maxW, h: sh, fontFace: FONT_B, fontSize: 15,
      color: dark ? C.soft : C.muted, isTextBox: true, margin: 0, valign: "top",
    });
    y += sh + 0.05;
  }
  return Math.max(y + 0.25, 2.2);
}

function card(s, x, y, w, h, opts = {}) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: opts.fill || C.light },
    line: { color: opts.line || opts.fill || C.light, width: 0.5 },
  });
}

async function iconCircle(s, name, x, y, d = 0.6, bg = C.secondary, fg = "FFFFFF") {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg, width: 0 } });
  const pad = d * 0.27;
  s.addImage({ data: await icon(name, fg), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

function text(s, str, o) {
  s.addText(str, { fontFace: FONT_B, fontSize: 14, color: C.ink, isTextBox: true, margin: 0, valign: "top", ...o });
}

function bullets(s, items, o) {
  const arr = items.map((t, i) => ({
    text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 },
  }));
  s.addText(arr, { fontFace: FONT_B, fontSize: 13, color: C.ink, isTextBox: true, valign: "top", ...o });
}

function pill(s, label, x, y, w, opts = {}) {
  const h = 0.34;
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.17, fill: { color: opts.fill || C.primary }, line: { color: opts.fill || C.primary, width: 0 } });
  text(s, label, { x, y, w, h, fontSize: 9.5, bold: true, color: opts.color || "FFFFFF", align: "center", valign: "middle", charSpacing: 1 });
}

// ───────────────────────────── slides ─────────────────────────────
async function build() {
  // 1. COVER
  {
    const s = base(true);
    if (LOGO.icon_white) s.addImage({ ...logoW("icon_white", 3.0), x: W - M - 3.0, y: 1.0, transparency: 70 });
    if (LOGO.wordmark_white) s.addImage({ ...logoW("wordmark_white", 2.1), x: M, y: 0.6 });
    const coverTitle = `${brand.company} × ${prospectName}`;
    const coverW = 8.6;
    let cs = 48;
    const cwAt = (pt) => coverTitle.length * pt * 0.6 / 72;
    while (cs > 30 && Math.ceil(cwAt(cs) / coverW) > 2) cs -= 2;
    const coverLines = Math.max(1, Math.ceil(cwAt(cs) / coverW));
    const coverH = coverLines * cs * 1.2 / 72 + 0.15;
    const coverY = Math.max(2.35, 3.2 - coverH / 2);
    s.addText(coverTitle, {
      x: M, y: coverY, w: coverW, h: coverH, fontFace: FONT_H, fontSize: cs, bold: true, color: "FFFFFF",
      isTextBox: true, margin: 0, valign: "top",
    });
    const afterTitle = coverY + coverH + 0.2;
    text(s, `${focus.service} proposal${trial ? "  ·  Free 30-day trial" : ""}`, { x: M, y: afterTitle, w: 8.2, h: 0.4, fontSize: 18, bold: true, color: "FFFFFF" });
    text(s, oneLiner, { x: M, y: afterTitle + 0.45, w: 7.6, h: 1.0, fontSize: 16, color: C.soft });
    text(s, [place, today, pros.contact_name ? `Prepared for ${pros.contact_name}${pros.contact_title ? ", " + pros.contact_title : ""}` : ""].filter(Boolean).join("   ·   "), {
      x: M, y: Math.min(afterTitle + 1.55, H - 1.1), w: 9, h: 0.4, fontSize: 12, color: C.accent,
    });
    s.addNotes(`Open with the one-liner. ${prospectName} is a ${typeWord} in ${place || "their market"}. This deck is ${focus.name}-first: the full pitch is ${focus.name} management; the other platforms are ${bundle ? `an add-on bundle at ${bundle.price} total` : `add-ons at ${addonPrice} each`}.`);
  }

  // 2. PROBLEM
  {
    const s = base();
    const y0 = title(s, "Cannabis can't market like everyone else", { sub: "Ads are restricted, so the customer journey lives on pay-to-play platforms — and most listings are left on autopilot." });
    const pains = [
      ["FaBan", "Restricted advertising", "Google Ads, Meta and most networks block cannabis. Discovery happens on Weedmaps, Leafly, Yelp and Google Maps."],
      ["FaClock", "Listings decay", "Menus drift from POS, deals expire, photos go stale, reviews sit unanswered. Every gap costs orders."],
      ["FaChartLine", "Spend isn't managed", "Budgets are set once and never optimized. Platform reps grow platform revenue, not your ROAS."],
      ["FaUserSlash", "Nobody owns it", "GMs and budtenders juggle dashboards on the side with no attribution and no accountability."],
    ];
    const cw = (W - 2 * M - 3 * 0.3) / 4;
    for (let i = 0; i < pains.length; i++) {
      const x = M + i * (cw + 0.3);
      card(s, x, y0, cw, 3.9);
      await iconCircle(s, pains[i][0], x + 0.3, y0 + 0.3, 0.65);
      text(s, pains[i][1], { x: x + 0.3, y: y0 + 1.15, w: cw - 0.6, h: 0.6, fontSize: 17, bold: true, color: C.primary });
      text(s, pains[i][2], { x: x + 0.3, y: y0 + 1.8, w: cw - 0.6, h: 2.0, fontSize: 13, color: C.ink });
    }
    s.addNotes("Airbnb's problem slide: plain-language pains. Tie each one to something you saw on the prospect's listings (the audit slide comes in the Weedmaps section).");
  }

  // 3. SOLUTION
  {
    const s = base(true);
    const y0 = title(s, `We run your ${focus.name} listings and ads — and we're accountable for the orders.`, { dark: true, eyebrow: "The solution" });
    text(s, "Full-service listing management: spend, listings, menus, reviews, deals, creative and data — handled by ex-Weedmaps operators. Start with the platform that matters most, add the rest when you're ready.", {
      x: M, y: y0 + 0.1, w: 7.8, h: 1.3, fontSize: 15, color: C.soft,
    });
    const bw = 1.75, gap = 0.25, x0 = W - M - (4 * bw + 3 * gap);
    for (let i = 0; i < ORDER.length; i++) {
      const k = ORDER[i], pl = PLAT[k];
      const isFocus = k === focusKey, inc = isIncluded(k);
      const on = isFocus || inc;
      const x = x0 + i * (bw + gap);
      card(s, x, 4.95, bw, 1.7, { fill: on ? "FFFFFF" : C.deep });
      s.addImage({ data: await icon(pl.icon, on ? C.primary : C.dim), x: x + bw / 2 - 0.3, y: 5.15, w: 0.6, h: 0.6 });
      text(s, pl.name, { x, y: 5.85, w: bw, h: 0.3, fontSize: 14, bold: true, color: on ? C.primary : C.dim, align: "center" });
      text(s, isFocus ? "THIS PROPOSAL" : inc ? "INCLUDED" : tileAddonLabel, { x, y: 6.15, w: bw, h: 0.25, fontSize: 8, bold: true, color: on ? C.ink : C.dim, align: "center", charSpacing: 1 });
    }
    s.addNotes(bundle
      ? `One sentence the prospect can repeat back. The tiles preview the deck's order: ${focus.name} in full, then the add-on bundle (${bundleLabel()} for ${bundle.price} total${total ? `; whole package ${total.price}` : ""}).`
      : "One sentence the prospect can repeat back. The tiles preview the deck's order: the focus platform in full, then each add-on.");
  }

  // ── FOCUS PLATFORM SECTION ──
  const secEyebrow = (label) => `${focus.service}  ·  ${label}`;

  // 4. WHY THIS PLATFORM
  {
    const s = base();
    const y0 = title(s, `Where ${prospectName}'s customers are on ${focus.name}`, { eyebrow: secEyebrow("Why it matters"), sub: focus.tagline });
    const items = focus.why_it_matters || [];
    const rh = (H - 0.8 - y0 - 0.15 * (items.length - 1)) / Math.max(items.length, 1);
    for (let i = 0; i < items.length; i++) {
      const y = y0 + i * (rh + 0.15);
      card(s, M, y, 7.9, rh);
      await iconCircle(s, items[i].icon || focus.icon, M + 0.2, y + (rh - 0.6) / 2, 0.6);
      text(s, items[i].title, { x: M + 1.0, y: y + 0.15, w: 6.7, h: 0.35, fontSize: 15, bold: true, color: C.primary });
      text(s, items[i].body, { x: M + 1.0, y: y + 0.5, w: 6.7, h: rh - 0.6, fontSize: 11.5, color: C.ink });
    }
    const mx = M + 8.2, mw = W - M - mx, mh = H - 0.8 - y0;
    card(s, mx, y0, mw, mh, { fill: C.primary });
    text(s, "YOUR MARKET", { x: mx + 0.35, y: y0 + 0.25, w: mw - 0.7, h: 0.35, fontSize: 11, color: C.accent, charSpacing: 4, bold: true });
    const comp = P.market && P.market.competitors_in_zone;
    const locs = Number(pros.locations) || 0;
    const bigNum = comp ? String(comp) : locs > 1 ? String(locs) : "—";
    const bigLabel = comp
      ? `competing ${pros.type === "brand" ? "brands" : "dispensaries"} in your ${focus.name} service zone`
      : locs > 1 ? `${focus.name} listings to manage across ${pros.state || "your market"}` : `competitors in your ${focus.name} service zone`;
    s.addText(bigNum, { x: mx + 0.35, y: y0 + 0.65, w: mw - 0.7, h: 1.1, fontFace: FONT_H, fontSize: 60, bold: true, color: C.accent, isTextBox: true, margin: 0 });
    text(s, bigLabel, { x: mx + 0.35, y: y0 + 1.8, w: mw - 0.7, h: 0.7, fontSize: 13, color: "FFFFFF" });
    text(s, (P.market && P.market.notes) || "", { x: mx + 0.35, y: y0 + 2.6, w: mw - 0.7, h: mh - 2.8, fontSize: 12, color: C.soft });
    s.addNotes(`Market validation for ${focus.name}. Ask which platform they think drives the most orders — most owners are guessing, which is the opening for the audit slide.`);
  }

  // 5. WHAT WE FOUND (audit for focus platform; falls back to all rows)
  {
    const s = base();
    const y0 = title(s, `What we found on ${prospectName}'s ${focus.name} listings`, { eyebrow: secEyebrow("Audit"), sub: "Every row is revenue currently left on the table." });
    let rows = auditFor(focusKey);
    if (!rows.length) rows = auditRows;
    const hasAudit = rows.length > 0;
    const audit = hasAudit ? rows : [{ platform: focus.name, finding: "Audit pending", impact: "Run the audit before presenting" }];
    const head = { fill: { color: C.primary }, color: "FFFFFF", bold: true, fontFace: FONT_B, fontSize: 13, valign: "middle" };
    const cell = { fontFace: FONT_B, fontSize: 12, color: C.ink, valign: "middle" };
    const tableRows = [
      [{ text: "Listing", options: head }, { text: "What we found", options: head }, { text: "What it's costing you", options: head }],
      ...audit.map((r, i) => {
        const fill = { color: i % 2 ? "FFFFFF" : C.light };
        const label = String(r.platform || "").replace(new RegExp(`^${focus.name}\\s*[·:-]?\\s*`, "i"), "") || focus.name;
        return [
          { text: label, options: { ...cell, bold: true, color: C.primary, fill } },
          { text: r.finding || "", options: { ...cell, fill } },
          { text: r.impact || "", options: { ...cell, fill } },
        ];
      }),
    ];
    const rowH = Math.min(0.6, (H - 0.8 - y0 - 0.2) / (audit.length + 1));
    s.addTable(tableRows, {
      x: M, y: y0, w: W - 2 * M, colW: [2.3, 5.1, 4.73], rowH,
      border: { type: "solid", color: "E6DED7", pt: 0.75 }, margin: [0.06, 0.15, 0.06, 0.15],
    });
    const yAfter = y0 + rowH * (audit.length + 1) + 0.35;
    const h = H - 0.8 - yAfter;
    if (hasAudit && h >= 0.9) {
      card(s, M, yAfter, W - 2 * M, Math.min(h, 1.2), { fill: C.primary });
      text(s, `${audit.length} gap${audit.length === 1 ? "" : "s"} on ${focus.name}`, { x: M + 0.35, y: yAfter + 0.2, w: 5, h: 0.45, fontSize: 18, bold: true, color: C.accent });
      text(s, "Each one maps to a line on the next slide. All of them are fixed inside the first 30 days of the launch plan.", { x: M + 0.35, y: yAfter + 0.62, w: W - 2 * M - 0.7, h: 0.5, fontSize: 12.5, color: "FFFFFF" });
    }
    s.addNotes("Read it like a doctor reading a chart — calm, specific, no blame. Rows marked 'to confirm' could not be verified from outside the account; ask on the call.");
  }

  // 6. WHAT WE'LL MANAGE
  {
    const s = base();
    const cfg = SV[focusKey] || {};
    const y0 = title(s, `What Sparks will manage on ${focus.name} for ${prospectName}`, { eyebrow: secEyebrow("Scope"), sub: cfg.notes ? `Scoped for this proposal: ${cfg.notes}` : undefined });
    const items = focus.what_we_manage || [];
    const cols = 2, gap = 0.3;
    const cw = (W - 2 * M - gap) / cols;
    const rowsN = Math.ceil(items.length / cols);
    const rh = Math.min(0.85, (H - 0.8 - y0 - 0.15 * (rowsN - 1)) / rowsN);
    for (let i = 0; i < items.length; i++) {
      const r = Math.floor(i / cols), c = i % cols;
      const x = M + c * (cw + gap), y = y0 + r * (rh + 0.15);
      card(s, x, y, cw, rh);
      s.addShape(pres.ShapeType.ellipse, { x: x + 0.25, y: y + (rh - 0.42) / 2, w: 0.42, h: 0.42, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });
      s.addText(String(i + 1), { x: x + 0.25, y: y + (rh - 0.42) / 2, w: 0.42, h: 0.42, fontFace: FONT_H, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
      text(s, items[i], { x: x + 0.85, y: y + 0.1, w: cw - 1.1, h: rh - 0.2, fontSize: 13, color: C.ink, valign: "middle" });
    }
    s.addNotes("The seven Sparks pillars (spend, listing, menu, reviews, deals, creative, data) applied to this platform. Tie each line back to an audit finding.");
  }

  // 7. RESULTS
  {
    const s = base(true);
    title(s, "The results Sparks clients see", { dark: true, eyebrow: secEyebrow("Results"), sub: "Averages across Sparks-managed accounts." });
    const proof = [...(brand.proof || [])];
    if (focus.proof && !proof.some((p) => p.stat === focus.proof.stat)) proof.unshift(focus.proof);
    const shown = proof.slice(0, 4);
    const n = shown.length || 1, gap = 0.3;
    const cw = (W - 2 * M - (n - 1) * gap) / n;
    shown.forEach((p, i) => {
      const x = M + i * (cw + gap);
      card(s, x, 2.55, cw, 3.2, { fill: C.deep });
      s.addText(p.stat, { x: x + 0.3, y: 2.8, w: cw - 0.4, h: 1.3, fontFace: FONT_H, fontSize: 48, bold: true, color: C.accent, isTextBox: true, margin: 0, valign: "top" });
      text(s, p.label, { x: x + 0.3, y: 4.2, w: cw - 0.6, h: 1.4, fontSize: 14, color: "FFFFFF" });
    });
    text(s, "Significant conversion gains typically show within 60–90 days as optimizations compound.", { x: M, y: 6.05, w: W - 2 * M, h: 0.5, fontSize: 13, color: C.soft, italic: true });
    s.addNotes("Only cite numbers published on sparkscann.com. ROAS is measured from on-platform conversions plus POS-integrated attribution in the Sparks dashboard.");
  }

  // 8. CASE STUDIES
  {
    const s = base();
    const y0 = title(s, `${focus.name} case studies`, { eyebrow: secEyebrow("Proof"), sub: "What full ownership of a listing looks like in practice." });
    const keys = Array.isArray(P.case_studies) && P.case_studies.length ? P.case_studies : focus.case_studies || Object.keys(caseStudies);
    const cs = keys.map((k) => caseStudies[k]).filter(Boolean).slice(0, 2);
    const gap = 0.35, cw = (W - 2 * M - gap * (cs.length - 1)) / Math.max(cs.length, 1);
    const h = H - 0.8 - y0;
    cs.forEach((c, i) => {
      const x = M + i * (cw + gap), y = y0;
      card(s, x, y, cw, h);
      text(s, c.name, { x: x + 0.35, y: y + 0.3, w: cw - 0.7, h: 0.45, fontSize: 20, bold: true, color: C.primary });
      text(s, `${c.type} · ${c.location}`, { x: x + 0.35, y: y + 0.75, w: cw - 0.7, h: 0.35, fontSize: 12, color: C.muted });
      text(s, [{ text: "Situation  ", options: { bold: true, color: C.primary } }, { text: c.situation }], { x: x + 0.35, y: y + 1.2, w: cw - 0.7, h: 0.8, fontSize: 12 });
      text(s, [{ text: "What we did  ", options: { bold: true, color: C.primary } }, { text: c.actions }], { x: x + 0.35, y: y + 2.0, w: cw - 0.7, h: 0.95, fontSize: 12 });
      const res = c.results || [];
      const sw = (cw - 0.7) / Math.max(res.length, 1);
      res.forEach((r, j) => {
        const sx = x + 0.35 + j * sw;
        s.addText(r.stat, { x: sx, y: y + h - 1.45, w: sw - 0.1, h: 0.7, fontFace: FONT_H, fontSize: res.length > 2 ? 24 : 30, bold: true, color: C.primary, isTextBox: true, margin: 0 });
        text(s, r.label, { x: sx, y: y + h - 0.75, w: sw - 0.15, h: 0.65, fontSize: 11, color: C.ink });
      });
    });
    s.addNotes("Purple Lotus is the 'decade-old account we doubled' story; Nuna Harvest is the 'new market, fast growth' story. Pick the one closest to the prospect.");
  }

  // 9. INVESTMENT (focus platform)
  {
    const s = base(true);
    title(s, `${focus.service} investment`, { dark: true, eyebrow: secEyebrow("Investment"), sub: inv.term || "" });
    const opts = Array.isArray(inv.options) ? inv.options.filter((o) => o && o.name) : [];
    const addonsPending = others.filter((k) => !isIncluded(k));
    if (bundle && addonsPending.length) opts.push({ name: "Add-on bundle", price: bundle.price, includes: `${bundleLabel()}. ${bundle.detail || ""}`.trim() });
    const gap = 0.3;
    if (opts.length) {
      const n = Math.min(opts.length, 3), cw = (W - 2 * M - gap * (n - 1)) / n;
      opts.slice(0, 3).forEach((o, i) => {
        const x = M + i * (cw + gap);
        card(s, x, 2.55, cw, 2.9, { fill: "FFFFFF" });
        text(s, o.name, { x: x + 0.35, y: 2.8, w: cw - 0.7, h: 0.4, fontSize: 16, bold: true, color: C.primary });
        const price = o.price || "Custom";
        const pf = price.length <= 9 ? 34 : price.length <= 14 ? 28 : 22;
        s.addText(price, { x: x + 0.35, y: 3.2, w: cw - 0.7, h: 0.9, fontFace: FONT_H, fontSize: pf, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
        text(s, o.includes || "", { x: x + 0.35, y: 4.2, w: cw - 0.7, h: 1.15, fontSize: 12, color: C.ink });
      });
    } else {
      card(s, M, 2.55, 6, 2.9, { fill: "FFFFFF" });
      text(s, focus.service, { x: M + 0.35, y: 2.8, w: 5.3, h: 0.4, fontSize: 16, bold: true, color: C.primary });
      s.addText(inv.monthly_fee || "Custom proposal", { x: M + 0.35, y: 3.2, w: 5.3, h: 0.95, fontFace: FONT_H, fontSize: 34, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    }
    if (trial) {
      text(s, trial, { x: M, y: 5.6, w: 6.6, h: 0.4, fontSize: 13, bold: true, color: C.accent });
      text(s, brand.guarantee, { x: M, y: 6.0, w: 6.5, h: 0.3, fontSize: 12, bold: true, color: "FFFFFF" });
      text(s, inv.ad_spend_note || "", { x: M, y: 6.3, w: 6.5, h: 0.4, fontSize: 10.5, color: C.soft, italic: true });
    } else {
      text(s, brand.guarantee, { x: M, y: 5.7, w: 6.5, h: 0.4, fontSize: 15, bold: true, color: C.accent });
      text(s, inv.ad_spend_note || "", { x: M, y: 6.1, w: 6.5, h: 0.5, fontSize: 11, color: C.soft, italic: true });
    }
    const addons = addonsPending.map((k) => PLAT[k].name);
    if (addons.length) {
      const teaser = bundle
        ? (total ? `Total package with ${listWords(addons)}: ${total.price}${total.detail ? ` — ${total.detail}` : ""}` : `Add ${listWords(addons)} management together for ${bundle.price}.`)
        : `Add ${listWords(addons)} management for ${addonPrice} each — previews on the next ${addons.length === 1 ? "slide" : "slides"}.`;
      text(s, teaser, { x: M + 6.8, y: 5.7, w: W - 2 * M - 6.8, h: 0.9, fontSize: 13, bold: true, color: "FFFFFF", align: "right" });
    }
    s.addNotes(`${trial ? trial + " " : ""}Pricing for ${focus.service}. ${brand.guarantee} Ad spend is billed by the platform; the fee covers management. ${bundle ? `The add-on bundle (${bundleLabel()}) is ${bundle.price} total${bundle.detail ? `: ${bundle.detail}` : ""}.` : `Each add-on is ${addonPrice}.`}${total ? ` Whole package: ${total.price}${total.detail ? ` ${total.detail}` : ""}.` : ""}`);
  }

  // ── ADD-ON / SECONDARY PLATFORM PREVIEWS ──
  for (const k of others) {
    const pl = PLAT[k];
    const inc = isIncluded(k);
    const cfg = SV[k] || {};
    const s = base();
    const eyebrow = inc ? `Included  ·  ${pl.name}` : bundle ? `Add-on bundle  ·  ${pl.name}  ·  ${bundle.price} for ${otherNames().join(" + ")}` : `Add-on  ·  ${pl.name}  ·  ${addonPrice}`;
    const y0 = title(s, pl.service, { eyebrow, sub: pl.tagline });
    const items = (pl.why_it_matters || []).slice(0, 3);
    const leftW = 7.4;
    const seen = auditFor(k);
    const rowsAvail = H - 0.8 - y0 - (seen.length ? 1.05 : 0);
    const rh = (rowsAvail - 0.12 * (items.length - 1)) / Math.max(items.length, 1);
    for (let i = 0; i < items.length; i++) {
      const y = y0 + i * (rh + 0.12);
      card(s, M, y, leftW, rh);
      await iconCircle(s, items[i].icon || pl.icon, M + 0.2, y + (rh - 0.52) / 2, 0.52);
      text(s, items[i].title, { x: M + 0.9, y: y + 0.12, w: leftW - 1.1, h: 0.32, fontSize: 14, bold: true, color: C.primary });
      text(s, items[i].body, { x: M + 0.9, y: y + 0.44, w: leftW - 1.1, h: rh - 0.5, fontSize: 11, color: C.ink });
    }
    if (seen.length) {
      const y = y0 + rowsAvail + 0.1;
      card(s, M, y, leftW, 0.95, { fill: C.highlight });
      text(s, `What we saw on ${prospectName}'s ${pl.name}`, { x: M + 0.3, y: y + 0.1, w: leftW - 0.6, h: 0.3, fontSize: 12, bold: true, color: C.primary });
      text(s, seen.map((r) => r.finding).join("  ·  "), { x: M + 0.3, y: y + 0.4, w: leftW - 0.6, h: 0.5, fontSize: 10.5, color: C.ink });
    }
    const rx = M + leftW + 0.3, rw = W - M - rx, rhgt = H - 0.8 - y0;
    card(s, rx, y0, rw, rhgt, { fill: C.primary });
    text(s, "WHAT'S INCLUDED", { x: rx + 0.35, y: y0 + 0.25, w: rw - 0.7, h: 0.3, fontSize: 11, bold: true, charSpacing: 4, color: C.accent });
    s.addText((pl.what_we_manage || []).map((t, i, a) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < a.length - 1, paraSpaceAfter: 3 } })),
      { x: rx + 0.2, y: y0 + 0.62, w: rw - 0.5, h: rhgt - 2.25, fontFace: FONT_B, fontSize: 10.5, color: "FFFFFF", isTextBox: true, valign: "top" });
    if (pl.proof) {
      s.addShape(pres.ShapeType.line, { x: rx + 0.35, y: y0 + rhgt - 1.5, w: rw - 0.7, h: 0, line: { color: C.dim, width: 0.75 } });
      s.addText(pl.proof.stat, { x: rx + 0.35, y: y0 + rhgt - 1.38, w: 2.0, h: 0.55, fontFace: FONT_H, fontSize: 24, bold: true, color: C.accent, isTextBox: true, margin: 0, valign: "middle" });
      text(s, pl.proof.label, { x: rx + 2.4, y: y0 + rhgt - 1.38, w: rw - 2.75, h: 0.55, fontSize: 10.5, color: C.soft, valign: "middle" });
    }
    pill(s, inc ? "INCLUDED IN THIS PROPOSAL" : bundle ? `IN THE ${bundle.price} ADD-ON BUNDLE` : `ADD-ON  ·  +${addonPrice}`, rx + 0.35, y0 + rhgt - 0.65, rw - 0.7, { fill: "FFFFFF", color: C.primary });
    s.addNotes(inc
      ? `${pl.service} is included in this proposal. Scope notes: ${cfg.notes || "see scope of services"}.`
      : bundle
        ? `${pl.service} preview. Part of the add-on bundle: ${bundleLabel()} for ${bundle.price} total${bundle.detail ? ` (${bundle.detail})` : ""}. Same seven pillars applied to ${pl.name}; can be switched on at any point.`
        : `${pl.service} preview. Offered as an add-on at ${addonPrice}. Same seven pillars applied to ${pl.name}; can be switched on at any point in the engagement.`);
  }

  // WHY SPARKS
  {
    const s = base(true);
    title(s, `Why ${brand.company}`, { dark: true, sub: "Competitive advantages you can verify." });
    const adv = [
      ["FaUserTie", "Built by ex-Weedmaps operators", "Four years inside Weedmaps working with the largest dispensaries. We know how the auction, ranking and menu systems actually work."],
      ["FaLayerGroup", "All four platforms + 50 directories", "One team, one strategy, one report across Weedmaps, Leafly, Yelp, Google and every citation that feeds local SEO."],
      ["FaPlug", "POS-integrated attribution", "Treez, Dutchie and other POS data joined with ad spend and platform metrics — you see which platform, campaign and deal drove revenue."],
      ["FaAward", "Recognized", "Top Cannabis Listing Management Agency 2026 (Cannabis Business Insights). Yelp Ads Certified Partner. 100+ businesses managed."],
    ];
    const gap = 0.3, cw = (W - 2 * M - gap) / 2, ch = 1.95;
    for (let i = 0; i < adv.length; i++) {
      const x = M + (i % 2) * (cw + gap), y = 2.35 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch, { fill: C.deep });
      await iconCircle(s, adv[i][0], x + 0.3, y + 0.3, 0.6, C.accent, C.primary);
      text(s, adv[i][1], { x: x + 1.1, y: y + 0.3, w: cw - 1.4, h: 0.6, fontSize: 16, bold: true, color: "FFFFFF" });
      text(s, adv[i][2], { x: x + 1.1, y: y + 0.9, w: cw - 1.4, h: 1.0, fontSize: 12, color: C.soft });
    }
    s.addNotes("Four reasons to pick us over doing it in-house, trusting platform reps, or hiring a general agency. The founder story is the credibility anchor; the dashboard is the differentiator.");
  }

  // TEAM
  {
    const s = base();
    const y0 = title(s, "Your account team", { sub: "The people who will run your listings." });
    const team = (brand.team || []).filter((t) => t && t.name);
    const n = Math.max(team.length, 1), gap = 0.3;
    const cw = Math.min(4.0, (W - 2 * M - gap * (n - 1)) / n);
    const h = H - 0.8 - y0;
    team.forEach((t, i) => {
      const x = M + i * (cw + gap), y = y0;
      card(s, x, y, cw, h);
      s.addShape(pres.ShapeType.ellipse, { x: x + 0.35, y: y + 0.35, w: 1.1, h: 1.1, fill: { color: C.primary }, line: { color: C.primary, width: 0 } });
      const initials = t.name.split(/\s+/).map((w) => w[0]).join("").slice(0, 2).toUpperCase();
      s.addText(initials, { x: x + 0.35, y: y + 0.35, w: 1.1, h: 1.1, fontFace: FONT_H, fontSize: 28, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
      text(s, t.name, { x: x + 0.35, y: y + 1.65, w: cw - 0.7, h: 0.45, fontSize: 18, bold: true, color: C.primary });
      text(s, t.title || "", { x: x + 0.35, y: y + 2.1, w: cw - 0.7, h: 0.35, fontSize: 12, color: C.primary, bold: true });
      text(s, t.bio || "", { x: x + 0.35, y: y + 2.55, w: cw - 0.7, h: h - 2.7, fontSize: 11.5, color: C.ink });
    });
    const px = M + n * (cw + gap) + 0.2;
    if (px < W - M - 3) {
      const pw = W - M - px;
      card(s, px, y0, pw, h, { fill: C.primary });
      text(s, "HOW WE WORK WITH YOU", { x: px + 0.35, y: y0 + 0.3, w: pw - 0.7, h: 0.35, fontSize: 11, color: C.accent, charSpacing: 4, bold: true });
      bullets(s, [
        "Dedicated account lead and weekly performance report",
        "Menu & listing specialists working daily in your platforms",
        "In-house creative for ads, deals and listing assets",
        "Data team maintaining your POS-integrated dashboard",
        "Direct line to leadership — no ticket queues",
      ], { x: px + 0.2, y: y0 + 0.75, w: pw - 0.5, h: h - 1.0, fontSize: 13, color: "FFFFFF" });
    }
    s.addNotes("Add team members in brand/config.yml. Headshots can replace the initials circles in PowerPoint.");
  }

  // NEXT STEPS
  {
    const s = base(true);
    title(s, "Next steps", { dark: true, sub: `A 90-day ${focus.name} launch plan for ${prospectName}.` });
    const plan = Array.isArray(P.plan) && P.plan.length ? P.plan : [
      { period: "Days 1–30", items: "Full audit, strategy, listing + menu clean-up, campaigns live" },
      { period: "Days 31–60", items: "Bid & deal optimization, review program, dashboard attribution" },
      { period: "Days 61–90", items: "Scale winners, quarterly plan, add-on platforms switched on" },
    ];
    const leftW = 7.3;
    plan.slice(0, 3).forEach((p, i) => {
      const y = 2.35 + i * 1.05;
      s.addShape(pres.ShapeType.ellipse, { x: M, y: y + 0.1, w: 0.55, h: 0.55, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
      s.addText(String(i + 1), { x: M, y: y + 0.1, w: 0.55, h: 0.55, fontFace: FONT_H, fontSize: 16, bold: true, color: C.primary, align: "center", valign: "middle", isTextBox: true, margin: 0 });
      text(s, p.period + (trial && i === 0 ? "  ·  Free trial" : ""), { x: M + 0.8, y, w: leftW - 0.8, h: 0.4, fontSize: 15, bold: true, color: "FFFFFF" });
      text(s, p.items, { x: M + 0.8, y: y + 0.4, w: leftW - 0.8, h: 0.6, fontSize: 12, color: C.soft });
    });
    if (trial) {
      text(s, trial, { x: M, y: 5.45, w: leftW, h: 0.35, fontSize: 13, bold: true, color: C.accent });
      text(s, brand.guarantee, { x: M, y: 5.8, w: leftW, h: 0.3, fontSize: 12, bold: true, color: "FFFFFF" });
    } else {
      text(s, brand.guarantee, { x: M, y: 5.6, w: leftW, h: 0.4, fontSize: 14, bold: true, color: C.accent });
    }
    if (LOGO.wordmark_white) s.addImage({ ...logoW("wordmark_white", 1.0), x: M, y: 6.15 });
    const px = M + leftW + 0.4, pw = W - M - px;
    card(s, px, 2.35, pw, 4.2, { fill: C.deep });
    text(s, "INVESTMENT SUMMARY", { x: px + 0.35, y: 2.6, w: pw - 0.7, h: 0.3, fontSize: 11, color: C.accent, charSpacing: 4, bold: true });
    const lines = [];
    const opts = Array.isArray(inv.options) ? inv.options.filter((o) => o && o.name) : [];
    if (opts.length) opts.forEach((o) => lines.push([o.name.toLowerCase().startsWith(focus.name.toLowerCase()) ? o.name : `${focus.name} · ${o.name}`, o.price || "Custom"]));
    else lines.push([focus.service, inv.monthly_fee || "Custom"]);
    const pendingKeys = others.filter((k) => !isIncluded(k));
    for (const k of others.filter(isIncluded)) lines.push([PLAT[k].service.replace(/ \(.*\)$/, ""), "Included"]);
    if (pendingKeys.length) {
      if (bundle) lines.push([pendingKeys.map((k) => PLAT[k].name).join(" + "), `+${bundle.price}`]);
      else for (const k of pendingKeys) lines.push([PLAT[k].service.replace(/ \(.*\)$/, ""), `+${addonPrice}`]);
    }
    if (total) lines.push([total.label || "Total package", total.price, true]);
    const lh = Math.min(0.42, 2.3 / lines.length);
    lines.forEach((l, i) => {
      const y = 3.0 + i * lh;
      if (l[2]) s.addShape(pres.ShapeType.line, { x: px + 0.35, y: y + 0.02, w: pw - 0.7, h: 0, line: { color: C.dim, width: 0.75 } });
      text(s, l[0], { x: px + 0.35, y, w: pw - 2.1, h: lh, fontSize: 12, bold: !!l[2], color: "FFFFFF", valign: "middle" });
      text(s, l[1], { x: px + pw - 1.95, y, w: 1.6, h: lh, fontSize: l[2] ? 14 : 12, bold: true, color: C.accent, align: "right", valign: "middle" });
    });
    if (total && total.detail) text(s, total.detail, { x: px + 0.35, y: 3.0 + lines.length * lh + 0.02, w: pw - 0.7, h: 0.45, fontSize: 10, italic: true, color: C.soft });
    const ct = brand.contact || {};
    const contactLine = [ct.name && `${ct.name}${ct.title ? ", " + ct.title : ""}`, ct.email, ct.phone, ct.booking_link].filter(Boolean).join("  ·  ");
    text(s, "Ready when you are.", { x: px + 0.35, y: 5.55, w: pw - 0.7, h: 0.4, fontSize: 16, bold: true, color: "FFFFFF" });
    text(s, contactLine, { x: px + 0.35, y: 5.95, w: pw - 0.7, h: 0.5, fontSize: 11, color: C.accent });
    s.addNotes(`${trial ? trial + " " : ""}Close: restate the guarantee, propose a start date, and ask for ${focus.name} admin access to begin the audit. ${bundle ? `The add-on bundle (${bundleLabel()}) is ${bundle.price} total${total ? `, so the whole package is ${total.price}` : ""}${bundle.detail ? ` — ${bundle.detail}` : ""}.` : `Add-ons are ${addonPrice} each.`} Can be switched on at any time.`);
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await pres.writeFile({ fileName: outPath });
  console.log(`Wrote ${path.relative(ROOT, outPath)} (${slideNo} slides)`);
}

build().catch((e) => {
  console.error(e);
  process.exit(1);
});
