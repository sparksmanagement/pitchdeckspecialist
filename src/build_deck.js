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
const tileAddonLabel = bundle ? "ADD-ON BUNDLE" : "ADD-ON";
const scopeList = listWords(includedNames.length ? includedNames : [focus.name]);
const oneLiner = `We manage your ${scopeList} listings and ads so your ${typeWord} gets found, clicked, and ordered from.`;

// audit rows grouped by platform key (row.platform starts with the platform name)
const auditRows = Array.isArray(P.audit) ? P.audit.filter((r) => r && (r.finding || r.impact)) : [];
const auditFor = (k) => auditRows.filter((r) => String(r.platform || "").toLowerCase().startsWith(PLAT[k].name.toLowerCase()));

// true when a callout declares it covers this audit row (callout.covers matches row.platform)
const c_in = (row, callouts) => callouts.some((c) => c.covers && String(row.platform || "").toLowerCase().includes(String(c.covers).toLowerCase()));

const itemTitle = (it) => (typeof it === "string" ? it : it.title || "");
const itemBody = (it) => (typeof it === "string" ? "" : it.body || "");

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

// Free-trial starburst. (x, y) is the top-left of the burst's bounding box.
function trialBurst(s, x, y, d = 1.7) {
  s.addShape(pres.ShapeType.star16, { x, y, w: d, h: d, fill: { color: "FFFFFF" }, line: { color: C.primary, width: 1.5 }, rotate: 8 });
  s.addText([
    { text: "FREE", options: { fontFace: "Arial Black", fontSize: 17, bold: true, breakLine: true } },
    { text: "30 DAYS", options: { fontFace: "Arial Black", fontSize: 14.5, bold: true, breakLine: true } },
    { text: "no invoice until day 31", options: { fontFace: FONT_B, fontSize: 6, bold: true } },
  ], { x: x + 0.15, y: y + 0.3, w: d - 0.3, h: d - 0.6, color: C.primary, align: "center", valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 0.95 });
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
      ["FaBan", "Restricted advertising", "Google Ads, Meta and most networks block cannabis. Discovery happens on Weedmaps, Leafly, Google Maps, Apple Maps and Yelp."],
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
    text(s, "Full-service listing management across the marketplaces, maps and directories your customers use: spend, listings, menus, reviews, deals, creative and data — handled by ex-Weedmaps operators. Start with the platform that matters most, add the rest when you're ready.", {
      x: M, y: y0 + 0.1, w: 7.8, h: 1.3, fontSize: 15, color: C.soft,
    });
    const dirOn = isIncluded("local_seo") || isIncluded("google");
    const tileFor = (k) => { const pl = PLAT[k], isFocus = k === focusKey, inc = isIncluded(k); return { icon: pl.icon, name: pl.name, on: isFocus || inc, label: isFocus ? "THIS PROPOSAL" : inc ? "INCLUDED" : tileAddonLabel }; };
    const groups = [
      ["Marketplaces", [tileFor("weedmaps"), tileFor("leafly")]],
      ["Maps", [tileFor("google"), { icon: "FaApple", name: "Apple Maps", on: dirOn, label: dirOn ? "INCLUDED" : tileAddonLabel }]],
      ["Directories", [tileFor("yelp"), { icon: "FaSitemap", name: "+20 directories", on: dirOn, label: dirOn ? "INCLUDED" : tileAddonLabel, sub: true }]],
    ];
    const bw = 1.62, gapIn = 0.18, gapOut = 0.45;
    const totalW = groups.reduce((a, g) => a + g[1].length * bw + (g[1].length - 1) * gapIn, 0) + (groups.length - 1) * gapOut;
    let x = W - M - totalW;
    for (const [gname, tiles] of groups) {
      const gw = tiles.length * bw + (tiles.length - 1) * gapIn;
      text(s, gname.toUpperCase(), { x, y: 4.6, w: gw, h: 0.25, fontSize: 9, bold: true, charSpacing: 3, color: C.accent, align: "center" });
      for (const t of tiles) {
        card(s, x, 4.95, bw, 1.7, { fill: t.on ? "FFFFFF" : C.deep });
        s.addImage({ data: await icon(t.icon, t.on ? C.primary : C.dim), x: x + bw / 2 - 0.3, y: 5.15, w: 0.6, h: 0.6 });
        text(s, t.name, { x, y: 5.85, w: bw, h: 0.3, fontSize: t.sub ? 11.5 : 13, bold: true, color: t.on ? C.primary : C.dim, align: "center" });
        text(s, t.label, { x, y: 6.15, w: bw, h: 0.25, fontSize: 7.5, bold: true, color: t.on ? C.ink : C.dim, align: "center", charSpacing: 1 });
        x += bw + gapIn;
      }
      x += gapOut - gapIn;
    }
    s.addNotes(bundle
      ? `One sentence the prospect can repeat back. The tiles preview the deck's order: ${focus.name} in full, then the add-on bundle (${bundleLabel()} for ${bundle.price} total${total ? `; whole package ${total.price}` : ""}).`
      : "One sentence the prospect can repeat back. The tiles preview the deck's order: the focus platform in full, then each add-on.");
  }

  // ── FOCUS PLATFORM SECTION ──
  const secEyebrow = (label) => `${focus.service}  ·  ${label}`;

  // 4. WHY SPARKS
  {
    const s = base();
    const y0 = title(s, `Why ${brand.company}`, { eyebrow: "Why Sparks", sub: "The team behind 100+ dispensary listings — recognized, partnered, and built by a Weedmaps executive." });
    const gap = 0.3;
    // row 1: three stat cards
    const stats = [
      ["100+", "dispensaries, deliveries and brands managed — the largest Weedmaps and directory management platform in cannabis"],
      ["#1", "Top Cannabis Listing Management Agency 2026 — Cannabis Business Insights"],
      ["10+ yrs", "working inside and alongside Weedmaps and Leafly. Our founder ran accounts at Weedmaps as an executive before starting Sparks."],
    ];
    const cw = (W - 2 * M - 2 * gap) / 3, r1h = 1.75;
    stats.forEach(([big, label], i) => {
      const x = M + i * (cw + gap);
      card(s, x, y0, cw, r1h);
      s.addText(big, { x: x + 0.3, y: y0 + 0.2, w: cw - 0.6, h: 0.75, fontFace: FONT_H, fontSize: 36, bold: true, color: C.primary, isTextBox: true, margin: 0, valign: "middle" });
      text(s, label, { x: x + 0.3, y: y0 + 0.95, w: cw - 0.6, h: r1h - 1.05, fontSize: 11, color: C.ink });
    });
    // row 2: official partners (left) + two advantage cards (right)
    const y1 = y0 + r1h + gap, r2h = H - 0.8 - y1;
    const pw = cw * 1.5 + gap * 0.5;
    card(s, M, y1, pw, r2h, { fill: C.primary });
    text(s, "OFFICIAL PARTNERS", { x: M + 0.35, y: y1 + 0.25, w: pw - 0.7, h: 0.3, fontSize: 11, bold: true, charSpacing: 4, color: C.accent });
    text(s, "Official partner programs with the directories and maps that decide local search.", { x: M + 0.35, y: y1 + 0.6, w: pw - 0.7, h: 0.5, fontSize: 12, color: C.soft });
    const partners = [["FaYelp", "Yelp"], ["FaGoogle", "Google"], ["FaApple", "Apple Maps"]];
    const tw = (pw - 0.7 - 0.2 * 2) / 3, ty = y1 + 1.1, th = r2h - 1.4;
    for (let i = 0; i < partners.length; i++) {
      const x = M + 0.35 + i * (tw + 0.2), y = ty;
      card(s, x, y, tw, th, { fill: "FFFFFF" });
      s.addImage({ data: await icon(partners[i][0], C.primary), x: x + tw / 2 - 0.25, y: y + 0.12, w: 0.5, h: 0.5 });
      text(s, partners[i][1], { x, y: y + 0.64, w: tw, h: 0.26, fontSize: 12.5, bold: true, color: C.primary, align: "center" });
      text(s, "OFFICIAL PARTNER", { x, y: y + 0.9, w: tw, h: 0.2, fontSize: 7, bold: true, charSpacing: 1, color: C.ink, align: "center" });
    }
    const rx = M + pw + gap, rw = W - M - rx, ah = (r2h - gap) / 2;
    const adv = [
      ["FaTrophy", "Best results of any Weedmaps advertiser", "Sparks clients out-convert every other advertiser on the platform: 88% more orders in 90 days, $17 back per ad dollar."],
      ["FaPlug", "POS-integrated attribution", "Treez, Dutchie and other POS data joined with spend and platform metrics — you see which platform, campaign and deal drove revenue."],
    ];
    for (let i = 0; i < adv.length; i++) {
      const y = y1 + i * (ah + gap);
      card(s, rx, y, rw, ah);
      await iconCircle(s, adv[i][0], rx + 0.3, y + (ah - 0.55) / 2, 0.55);
      text(s, adv[i][1], { x: rx + 1.05, y: y + 0.15, w: rw - 1.3, h: 0.3, fontSize: 13.5, bold: true, color: C.primary });
      text(s, adv[i][2], { x: rx + 1.05, y: y + 0.45, w: rw - 1.3, h: ah - 0.55, fontSize: 11, color: C.ink });
    }
    s.addNotes("Credibility before the audit. Say 'official partner' only for Yelp, Google and Apple Maps. For Weedmaps and Leafly, say we've worked with them for 10+ years and the founder was a Weedmaps executive — not 'partner'.");
  }

  // 5. WHAT WE FOUND — iPhone mockup of the worst listing + callouts (audit_mock), else table
  {
    const s = base();
    const mock = P.audit_mock && typeof P.audit_mock === "object" ? P.audit_mock : null;
    let rows = auditFor(focusKey);
    if (!rows.length) rows = auditRows;
    const hasAudit = rows.length > 0;
    const y0 = title(s, `What we found on ${prospectName}'s ${focus.name} listings`, { eyebrow: secEyebrow("Audit"), sub: mock ? `${mock.listing || "Their listing"} as a shopper sees it — and what it's costing.` : "Every row is revenue currently left on the table." });
    const FLAG = "C8322B"; // callout red
    if (mock) {
      // ── phone ──
      const ph = H - 0.75 - y0, pw = 2.3, px = M, py = y0;
      const shot = mock.screenshot && fs.existsSync(path.join(ROOT, mock.screenshot)) ? path.join(ROOT, mock.screenshot) : null;
      s.addShape(pres.ShapeType.roundRect, { x: px, y: py, w: pw, h: ph, rectRadius: 0.32, fill: { color: "1C1C1E" }, line: { color: "3A3A3C", width: 1 } });
      const sx = px + 0.09, sy = py + 0.09, sw = pw - 0.18, sh = ph - 0.18;
      s.addShape(pres.ShapeType.roundRect, { x: sx, y: sy, w: sw, h: sh, rectRadius: 0.25, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
      s.addShape(pres.ShapeType.roundRect, { x: px + pw / 2 - 0.35, y: sy + 0.08, w: 0.7, h: 0.16, rectRadius: 0.08, fill: { color: "1C1C1E" }, line: { color: "1C1C1E", width: 0 } });
      const sections = {}; // section key -> y (for markers)
      const img = { x: sx + 0.12, y: sy + 0.36, w: sw - 0.24, h: sh - 0.5 };
      if (shot) {
        s.addImage({ path: shot, x: img.x, y: img.y, w: img.w, h: img.h, sizing: { type: "cover", w: img.w, h: img.h } });
      } else {
        const ix = sx + 0.14, iw = sw - 0.28;
        let y = sy + 0.32;
        text(s, "9:41", { x: ix, y, w: 0.6, h: 0.18, fontSize: 7, bold: true, color: "1C1C1E" });
        text(s, "●●● 100%", { x: sx + sw - 0.9, y, w: 0.76, h: 0.18, fontSize: 6, color: "1C1C1E", align: "right" });
        y += 0.24;
        // hero
        s.addShape(pres.ShapeType.rect, { x: sx + 0.02, y, w: sw - 0.04, h: 0.9, fill: { color: "E9E4DF" }, line: { color: "E9E4DF", width: 0 } });
        s.addImage({ data: await icon("FaImage", "B8AFA7"), x: sx + sw / 2 - 0.2, y: y + 0.25, w: 0.4, h: 0.4 });
        y += 1.0;
        sections.name = y + 0.15;
        text(s, mock.listing || prospectName, { x: ix, y, w: iw - 0.42, h: 0.42, fontSize: 10, bold: true, color: "1C1C1E" });
        y += 0.42;
        text(s, mock.subtitle || "", { x: ix, y, w: iw, h: 0.2, fontSize: 6.5, color: "6E6E73" });
        y += 0.24;
        sections.rating = y + 0.1;
        text(s, [{ text: "★ ", options: { color: "F5A623", bold: true } }, { text: `${mock.rating || "—"}  `, options: { bold: true } }, { text: mock.reviews || "", options: { color: "6E6E73" } }], { x: ix, y, w: iw, h: 0.22, fontSize: 8, color: "1C1C1E" });
        y += 0.28;
        sections.badges = y + 0.12;
        let bx = ix;
        for (const b of (mock.badges || []).slice(0, 3)) {
          const bwid = Math.min(0.62, 0.12 + b.length * 0.05);
          s.addShape(pres.ShapeType.roundRect, { x: bx, y, w: bwid, h: 0.22, rectRadius: 0.11, fill: { color: "F2EFEC" }, line: { color: "E2DDD8", width: 0.5 } });
          text(s, b, { x: bx, y, w: bwid, h: 0.22, fontSize: 6, color: "1C1C1E", align: "center", valign: "middle" });
          bx += bwid + 0.06;
        }
        y += 0.36;
        const rowsDef = [["deals", "Deals", mock.deals], ["menu", "Menu", mock.menu], ["reviews", "Reviews", mock.reviews_line]];
        for (const [key, label, val] of rowsDef) {
          s.addShape(pres.ShapeType.line, { x: ix, y: y - 0.04, w: iw, h: 0, line: { color: "ECE7E2", width: 0.5 } });
          sections[key] = y + 0.2;
          text(s, label, { x: ix, y, w: iw, h: 0.2, fontSize: 7.5, bold: true, color: "1C1C1E" });
          text(s, val || "—", { x: ix, y: y + 0.19, w: iw, h: 0.3, fontSize: 6.5, color: "6E6E73" });
          y += 0.5;
        }
      }
      // ── callouts ──
      const callouts = (Array.isArray(mock.callouts) ? mock.callouts : []).slice(0, 5);
      const cx = px + pw + 0.35, cwid = 4.4, ch = (ph - 0.1 * (callouts.length - 1)) / Math.max(callouts.length, 1);
      callouts.forEach((c, i) => {
        const y = y0 + i * (ch + 0.1);
        card(s, cx, y, cwid, ch);
        s.addShape(pres.ShapeType.ellipse, { x: cx + 0.2, y: y + 0.18, w: 0.36, h: 0.36, fill: { color: FLAG }, line: { color: FLAG, width: 0 } });
        s.addText(String(i + 1), { x: cx + 0.2, y: y + 0.18, w: 0.36, h: 0.36, fontFace: FONT_H, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
        text(s, c.title || "", { x: cx + 0.7, y: y + 0.1, w: cwid - 0.85, h: 0.28, fontSize: 11.5, bold: true, color: C.primary, valign: "middle" });
        text(s, c.text || "", { x: cx + 0.7, y: y + 0.38, w: cwid - 0.85, h: ch - 0.42, fontSize: 9, color: C.ink });
        // marker on the phone
        let my = null, mx = sx + sw - 0.4;
        if (shot && Array.isArray(c.at)) { mx = img.x + c.at[0] * img.w - 0.15; my = img.y + c.at[1] * img.h - 0.15; }
        else if (c.section && sections[c.section] != null) my = sections[c.section] - 0.15;
        if (my != null) {
          s.addShape(pres.ShapeType.ellipse, { x: mx, y: my, w: 0.3, h: 0.3, fill: { color: FLAG }, line: { color: "FFFFFF", width: 1.5 } });
          s.addText(String(i + 1), { x: mx, y: my, w: 0.3, h: 0.3, fontFace: FONT_H, fontSize: 9, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
        }
      });
      // ── fleet-wide findings ──
      const fx = cx + cwid + 0.3, fw = W - M - fx;
      const fleet = rows.filter((r) => !c_in(r, callouts)).slice(0, 5);
      card(s, fx, y0, fw, ph, { fill: C.primary });
      text(s, (mock.fleet_title || "Across all listings").toUpperCase(), { x: fx + 0.3, y: y0 + 0.22, w: fw - 0.6, h: 0.3, fontSize: 10, bold: true, charSpacing: 3, color: C.accent });
      const fh = (ph - 0.75) / Math.max(fleet.length, 1);
      fleet.forEach((r, i) => {
        const y = y0 + 0.6 + i * fh;
        const label = String(r.platform || "").replace(new RegExp(`^${focus.name}\\s*[·:-]?\\s*`, "i"), "") || focus.name;
        text(s, label, { x: fx + 0.3, y, w: fw - 0.6, h: 0.22, fontSize: 10, bold: true, color: "FFFFFF" });
        text(s, r.finding || "", { x: fx + 0.3, y: y + 0.22, w: fw - 0.6, h: fh - 0.26, fontSize: 9, color: C.soft });
      });
      s.addNotes(`Walk the phone left to right: this is ${mock.listing || "their listing"} exactly as a shopper sees it. Each numbered flag is a fix in the first 30 days. Right column: the same pattern across the other listings. Items marked 'to confirm' need account access to verify.`);
    } else {
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
      const body = itemBody(items[i]);
      if (body) {
        text(s, itemTitle(items[i]), { x: x + 0.85, y: y + 0.1, w: cw - 1.1, h: 0.3, fontSize: 13.5, bold: true, color: C.primary });
        text(s, body, { x: x + 0.85, y: y + 0.4, w: cw - 1.1, h: rh - 0.45, fontSize: 11.5, color: C.ink });
      } else {
        text(s, itemTitle(items[i]), { x: x + 0.85, y: y + 0.1, w: cw - 1.1, h: rh - 0.2, fontSize: 13, color: C.ink, valign: "middle" });
      }
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
    if (trial) trialBurst(s, W - M - 1.75, 0.5, 1.75);
    let opts = Array.isArray(inv.options) ? inv.options.filter((o) => o && o.name) : [];
    const addonsPending = others.filter((k) => !isIncluded(k));
    if (!opts.length) {
      opts = [{ name: focus.service, price: inv.monthly_fee || "Custom proposal", includes: "" }];
      if (bundle && addonsPending.length) opts.push({ name: "Listing bundle", price: bundle.price, includes: `${bundleLabel()}. ${bundle.detail || ""}`.trim() });
      if (total) opts.push({ name: total.label || "Total package", price: total.price, includes: total.detail || "", recommended: true });
    }
    const gap = 0.3, n = Math.min(opts.length, 3), cw = (W - 2 * M - gap * (n - 1)) / n;
    const cy = 2.5, ch = 3.15;
    opts.slice(0, 3).forEach((o, i) => {
      const x = M + i * (cw + gap);
      const rec = !!o.recommended;
      card(s, x, cy, cw, ch, { fill: rec ? C.highlight : "FFFFFF", line: rec ? "FFFFFF" : undefined });
      if (rec) {
        s.addShape(pres.ShapeType.roundRect, { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { type: "none" }, line: { color: "FFFFFF", width: 2 } });
        pill(s, "RECOMMENDED", x + cw - 1.65, cy + 0.28, 1.35);
      }
      text(s, o.name, { x: x + 0.35, y: cy + 0.25, w: cw - (rec ? 2.1 : 0.7), h: 0.4, fontSize: 16, bold: true, color: C.primary, valign: "middle" });
      const price = o.price || "Custom";
      const pf = price.length <= 9 ? 36 : price.length <= 14 ? 28 : 22;
      s.addText(price, { x: x + 0.35, y: cy + 0.7, w: cw - 0.7, h: 0.8, fontFace: FONT_H, fontSize: pf, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
      if (o.alt) text(s, o.alt, { x: x + 0.35, y: cy + 1.5, w: cw - 0.7, h: 0.55, fontSize: 11.5, italic: true, color: C.primary });
      text(s, o.includes || "", { x: x + 0.35, y: cy + (o.alt ? 2.1 : 1.6), w: cw - 0.7, h: ch - (o.alt ? 2.25 : 1.75), fontSize: 12, color: C.ink });
    });
    const by = cy + ch + 0.12;
    if (trial) {
      text(s, trial, { x: M, y: by, w: 7.5, h: 0.35, fontSize: 13, bold: true, color: C.accent });
      text(s, brand.guarantee, { x: M, y: by + 0.36, w: 7.5, h: 0.3, fontSize: 12, bold: true, color: "FFFFFF" });
      text(s, inv.ad_spend_note || "", { x: M, y: by + 0.66, w: 7.5, h: 0.35, fontSize: 10.5, color: C.soft, italic: true });
    } else {
      text(s, brand.guarantee, { x: M, y: by, w: 7.5, h: 0.4, fontSize: 14, bold: true, color: C.accent });
      text(s, inv.ad_spend_note || "", { x: M, y: by + 0.42, w: 7.5, h: 0.4, fontSize: 10.5, color: C.soft, italic: true });
    }
    const showsTotal = total && opts.some((o) => o.price === total.price);
    const addons = addonsPending.map((k) => PLAT[k].name);
    if (addons.length && !showsTotal) {
      const teaser = bundle
        ? (total ? `Total package with ${listWords(addons)}: ${total.price}${total.detail ? ` — ${total.detail}` : ""}` : `Add ${listWords(addons)} management together for ${bundle.price}.`)
        : `Add ${listWords(addons)} management for ${addonPrice} each — previews on the next ${addons.length === 1 ? "slide" : "slides"}.`;
      text(s, teaser, { x: M + 7.8, y: by, w: W - 2 * M - 7.8, h: 0.9, fontSize: 12.5, bold: true, color: "FFFFFF", align: "right" });
    } else if (addons.length) {
      text(s, `${listWords(addons)} previews follow. Any platform can be switched on at any point in the engagement.`, { x: M + 7.8, y: by, w: W - 2 * M - 7.8, h: 0.9, fontSize: 11.5, color: C.soft, align: "right" });
    }
    s.addNotes(`${trial ? trial + " " : ""}Pricing for ${focus.service}. ${brand.guarantee} Ad spend is billed by the platform; the fee covers management. ${bundle ? `The listing bundle (${bundleLabel()}) is ${bundle.price} total${bundle.detail ? `: ${bundle.detail}` : ""}.` : `Each add-on is ${addonPrice}.`}${total ? ` Recommended: the whole package at ${total.price}${total.detail ? ` ${total.detail}` : ""}.` : ""}`);
  }

  // ── ADD-ON / SECONDARY PLATFORM PREVIEWS ──
  for (const k of others) {
    const pl = PLAT[k];
    const inc = isIncluded(k);
    const cfg = SV[k] || {};
    const s = base();
    const eyebrow = inc ? `Included  ·  ${pl.name}` : bundle ? `Add-on bundle  ·  ${pl.name}  ·  ${bundle.price} for ${bundle.label || otherNames().join(" + ")}` : `Add-on  ·  ${pl.name}  ·  ${addonPrice}`;
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
    {
      const runs = [];
      const wm = pl.what_we_manage || [];
      wm.forEach((it, i) => {
        const last = i === wm.length - 1;
        const body = itemBody(it);
        runs.push({ text: itemTitle(it), options: { bold: true, bullet: { indent: 12 }, breakLine: !body && !last, paraSpaceAfter: 2 } });
        if (body) runs.push({ text: ` — ${body}`, options: { color: C.soft, breakLine: !last } });
      });
      s.addText(runs, { x: rx + 0.2, y: y0 + 0.6, w: rw - 0.5, h: rhgt - 2.0, fontFace: FONT_B, fontSize: 10, color: "FFFFFF", isTextBox: true, valign: "top" });
    }
    if (pl.proof) {
      s.addShape(pres.ShapeType.line, { x: rx + 0.35, y: y0 + rhgt - 1.32, w: rw - 0.7, h: 0, line: { color: C.dim, width: 0.75 } });
      s.addText(pl.proof.stat, { x: rx + 0.35, y: y0 + rhgt - 1.25, w: 2.0, h: 0.5, fontFace: FONT_H, fontSize: 22, bold: true, color: C.accent, isTextBox: true, margin: 0, valign: "middle" });
      text(s, pl.proof.label, { x: rx + 2.4, y: y0 + rhgt - 1.25, w: rw - 2.75, h: 0.5, fontSize: 10, color: C.soft, valign: "middle" });
    }
    pill(s, inc ? "INCLUDED IN THIS PROPOSAL" : bundle ? `IN THE ${bundle.price} ADD-ON BUNDLE` : `ADD-ON  ·  +${addonPrice}`, rx + 0.35, y0 + rhgt - 0.65, rw - 0.7, { fill: "FFFFFF", color: C.primary });
    s.addNotes(inc
      ? `${pl.service} is included in this proposal. Scope notes: ${cfg.notes || "see scope of services"}.`
      : bundle
        ? `${pl.service} preview. Part of the add-on bundle: ${bundleLabel()} for ${bundle.price} total${bundle.detail ? ` (${bundle.detail})` : ""}. Same seven pillars applied to ${pl.name}; can be switched on at any point.`
        : `${pl.service} preview. Offered as an add-on at ${addonPrice}. Same seven pillars applied to ${pl.name}; can be switched on at any point in the engagement.`);
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
    text(s, "INVESTMENT SUMMARY", { x: px + 0.35, y: 2.6, w: pw - 1.5, h: 0.3, fontSize: 10.5, color: C.accent, charSpacing: 3, bold: true });
    if (trial) trialBurst(s, px + pw - 1.7 * 0.7, 2.35 - 1.7 * 0.62, 1.7);
    const lines = [];
    const opts = Array.isArray(inv.options) ? inv.options.filter((o) => o && o.name) : [];
    const pendingKeys = others.filter((k) => !isIncluded(k));
    if (opts.length) opts.forEach((o) => lines.push([o.name + (o.optional ? "  (optional)" : ""), o.price || "Custom", !!o.recommended]));
    else lines.push([focus.service, inv.monthly_fee || "Custom"]);
    for (const k of others.filter(isIncluded)) lines.push([PLAT[k].service.replace(/ \(.*\)$/, ""), "Included"]);
    if (pendingKeys.length && !(bundle && opts.some((o) => o.price === bundle.price))) {
      if (bundle) lines.push([pendingKeys.map((k) => PLAT[k].name).join(" + "), `+${bundle.price}`]);
      else for (const k of pendingKeys) lines.push([PLAT[k].service.replace(/ \(.*\)$/, ""), `+${addonPrice}`]);
    }
    if (total && !opts.some((o) => o.price === total.price)) lines.push([total.label || "Total package", total.price, true]);
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
