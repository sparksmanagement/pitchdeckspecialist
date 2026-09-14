#!/usr/bin/env node
/**
 * Sparks pitch deck generator.
 *
 *   node src/build_deck.js prospects/<slug>.yml [-o decks/<name>.pptx]
 *
 * Reads brand/config.yml, brand/sparks.md (for reference only), src/case_studies.yml
 * and one prospect file; writes a 14-slide .pptx modeled on the Slidebean Airbnb deck.
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
const P = yaml.load(fs.readFileSync(prospectPath, "utf8"));

const slug = path.basename(prospectPath, ".yml").replace(/^_/, "");
const outPath =
  outIdx > -1 ? path.resolve(args[outIdx + 1]) : path.join(ROOT, "decks", `Sparks-x-${slug}.pptx`);

// ───────────────────────────── helpers ─────────────────────────────
const C = brand.palette;
const FONT_H = "Cambria";
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
const typeWord = { dispensary: "dispensary", delivery: "delivery service", brand: "brand" }[pros.type] || "dispensary";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = brand.company;
pres.company = brand.company;
pres.title = `${brand.company} × ${prospectName}`;

let slideNo = 0;

function base(dark = false) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.primary : "FFFFFF" };
  slideNo += 1;
  // footer
  s.addText(`${brand.company}  ·  ${brand.website}`, {
    x: M, y: H - 0.45, w: 6, h: 0.3, fontFace: FONT_B, fontSize: 10,
    color: dark ? C.accent : C.muted, isTextBox: true, margin: 0,
  });
  s.addText(String(slideNo), {
    x: W - M - 1, y: H - 0.45, w: 1, h: 0.3, fontFace: FONT_B, fontSize: 10, align: "right",
    color: dark ? C.accent : C.muted, isTextBox: true, margin: 0,
  });
  return s;
}

// Titles shrink to stay on one line where possible; a title that still wraps
// pushes the subtitle down so the two never collide.
function title(s, text, opts = {}) {
  const dark = !!opts.dark;
  const maxW = W - 2 * M;
  let size = 36;
  const widthAt = (pt) => text.length * pt * 0.66 / 72; // conservative avg glyph width (covers Cambria + LibreOffice fallback serif)
  while (size > 24 && widthAt(size) > maxW) size -= 2;
  const lines = Math.max(1, Math.ceil(widthAt(size) / maxW));
  const th = lines * size * 1.25 / 72 + 0.1;
  s.addText(text, {
    x: M, y: 0.55, w: maxW, h: th, fontFace: FONT_H, fontSize: size, bold: true,
    color: dark ? "FFFFFF" : C.primary, isTextBox: true, margin: 0, valign: "top",
  });
  if (opts.sub) {
    s.addText(opts.sub, {
      x: M, y: 0.55 + th + 0.12, w: maxW, h: 0.5, fontFace: FONT_B, fontSize: 16,
      color: dark ? C.accent : C.muted, isTextBox: true, margin: 0, valign: "top",
    });
  }
}

function card(s, x, y, w, h, opts = {}) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: opts.fill || C.light },
    line: { color: opts.line || C.light, width: 0.5 },
    shadow: opts.shadow ? { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.12 } : undefined,
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
    text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 6 },
  }));
  s.addText(arr, { fontFace: FONT_B, fontSize: 14, color: C.ink, isTextBox: true, valign: "top", ...o });
}

// ───────────────────────────── slides ─────────────────────────────
async function build() {
  // 1. COVER
  {
    const s = base(true);
    s.addShape(pres.ShapeType.ellipse, { x: 9.2, y: -1.6, w: 6, h: 6, fill: { color: C.secondary }, line: { color: C.secondary, width: 0 } });
    s.addShape(pres.ShapeType.ellipse, { x: 11.2, y: 2.6, w: 3.2, h: 3.2, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
    text(s, brand.company.toUpperCase(), { x: M, y: 0.7, w: 6, h: 0.4, fontSize: 14, color: C.accent, charSpacing: 6, bold: true });
    const coverTitle = `${brand.company} × ${prospectName}`;
    const coverW = 8.6;
    let cs = 48;
    const cwAt = (pt) => coverTitle.length * pt * 0.66 / 72;
    while (cs > 30 && Math.ceil(cwAt(cs) / coverW) > 2) cs -= 2;
    const coverLines = Math.max(1, Math.ceil(cwAt(cs) / coverW));
    const coverH = coverLines * cs * 1.2 / 72 + 0.15;
    const coverY = Math.max(1.3, 2.9 - coverH / 2);
    s.addText(coverTitle, {
      x: M, y: coverY, w: coverW, h: coverH, fontFace: FONT_H, fontSize: cs, bold: true, color: "FFFFFF",
      isTextBox: true, margin: 0, valign: "top",
    });
    const afterTitle = coverY + coverH + 0.2;
    text(s, brand.one_liner, { x: M, y: afterTitle, w: 7.6, h: 1.1, fontSize: 18, color: "E6EFE9" });
    text(s, [place, today, pros.contact_name ? `Prepared for ${pros.contact_name}${pros.contact_title ? ", " + pros.contact_title : ""}` : ""].filter(Boolean).join("   ·   "), {
      x: M, y: Math.min(afterTitle + 1.35, H - 1.1), w: 9, h: 0.4, fontSize: 12, color: C.accent,
    });
    s.addNotes(`Open with the one-liner. ${prospectName} is a ${typeWord} in ${place || "their market"}. Goal of this deck: show we've already looked at their listings, we know the platforms from the inside, and we can commit to results in 30 days.`);
  }

  // 2. PROBLEM
  {
    const s = base();
    title(s, "Cannabis can't market like everyone else", { sub: "Ads are restricted, so the customer journey lives on four pay-to-play platforms — and most listings are left on autopilot." });
    const pains = [
      ["FaBan", "Restricted advertising", "Google Ads, Meta and most networks block cannabis. Discovery happens on Weedmaps, Leafly, Yelp and Google Maps."],
      ["FaClock", "Listings decay", "Menus drift from POS, deals expire, photos go stale, reviews sit unanswered. Every gap costs orders."],
      ["FaChartLine", "Spend isn't managed", "Budgets are set once and never optimized. Platform reps grow platform revenue, not your ROAS."],
      ["FaUserSlash", "Nobody owns it", "GMs and budtenders juggle four dashboards on the side with no attribution and no accountability."],
    ];
    const cw = (W - 2 * M - 3 * 0.3) / 4;
    for (let i = 0; i < pains.length; i++) {
      const x = M + i * (cw + 0.3);
      card(s, x, 2.25, cw, 3.9);
      await iconCircle(s, pains[i][0], x + 0.3, 2.55, 0.65);
      text(s, pains[i][1], { x: x + 0.3, y: 3.4, w: cw - 0.6, h: 0.6, fontSize: 17, bold: true, color: C.primary });
      text(s, pains[i][2], { x: x + 0.3, y: 4.05, w: cw - 0.6, h: 2.2, fontSize: 13, color: C.ink });
    }
    s.addNotes("Airbnb's problem slide: three plain-language pains. Ours: four. Tie each one to something you saw on the prospect's listings (the audit slide comes next).");
  }

  // 3. SOLUTION
  {
    const s = base(true);
    text(s, "THE SOLUTION", { x: M, y: 0.6, w: 6, h: 0.4, fontSize: 13, color: C.accent, charSpacing: 6, bold: true });
    s.addText("We run your listings and ads on every platform your customers use — and we're accountable for the orders.", {
      x: M, y: 1.15, w: 8.2, h: 2.5, fontFace: FONT_H, fontSize: 32, bold: true, color: "FFFFFF", isTextBox: true, margin: 0, valign: "top",
    });
    text(s, `Full-service listing management for ${typeWord === "brand" ? "cannabis brands" : "dispensaries and deliveries"}: spend, listings, menus, reviews, deals, creative and data — handled by ex-Weedmaps operators.`, {
      x: M, y: 3.85, w: 7.8, h: 1.2, fontSize: 16, color: "D9E5DD",
    });
    const platforms = [["FaMapMarkedAlt", "Weedmaps"], ["FaLeaf", "Leafly"], ["FaYelp", "Yelp"], ["FaGoogle", "Google"]];
    const bw = 1.75, gap = 0.25, x0 = W - M - (4 * bw + 3 * gap);
    for (let i = 0; i < platforms.length; i++) {
      const x = x0 + i * (bw + gap);
      card(s, x, 5.1, bw, 1.55, { fill: C.secondary, line: C.secondary });
      s.addImage({ data: await icon(platforms[i][0], "FFFFFF"), x: x + bw / 2 - 0.3, y: 5.3, w: 0.6, h: 0.6 });
      text(s, platforms[i][1], { x, y: 6.0, w: bw, h: 0.5, fontSize: 14, bold: true, color: "FFFFFF", align: "center" });
    }
    s.addNotes("Like Airbnb's 'Book rooms with locals, rather than hotels' — one sentence the prospect can repeat back. Don't move on until they nod.");
  }

  // 4. WHERE YOUR CUSTOMERS ARE (market validation)
  {
    const s = base();
    title(s, `Where ${place || "your"} customers find a ${typeWord}`, { sub: "Market validation: the four channels that drive cannabis discovery and ordering." });
    const rows = [
      ["Weedmaps", "The largest cannabis marketplace. Auction-based placement in your service zone; menu, deals and reviews drive the order.", "FaMapMarkedAlt"],
      ["Leafly", "Strain-led discovery and online ordering; second marketplace most customers check.", "FaLeaf"],
      ["Yelp", "Where reviews decide. Yelp Ads reach shoppers comparing local options; Sparks is a Yelp Ads Certified Partner.", "FaYelp"],
      ["Google", "\"Dispensary near me\" — the Maps 3-pack, Business Profile, posts, reviews and compliant Ads.", "FaGoogle"],
    ];
    const rh = 1.0;
    for (let i = 0; i < rows.length; i++) {
      const y = 2.2 + i * (rh + 0.15);
      card(s, M, y, 7.9, rh);
      await iconCircle(s, rows[i][2], M + 0.2, y + 0.2, 0.6);
      text(s, rows[i][0], { x: M + 1.0, y: y + 0.15, w: 2, h: 0.35, fontSize: 15, bold: true, color: C.primary });
      text(s, rows[i][1], { x: M + 1.0, y: y + 0.48, w: 6.7, h: 0.5, fontSize: 11.5, color: C.ink });
    }
    // market panel
    const mx = M + 8.2, mw = W - M - mx;
    card(s, mx, 2.2, mw, 4.45, { fill: C.primary, line: C.primary });
    text(s, "YOUR MARKET", { x: mx + 0.35, y: 2.45, w: mw - 0.7, h: 0.35, fontSize: 11, color: C.accent, charSpacing: 4, bold: true });
    const comp = P.market && P.market.competitors_in_zone;
    s.addText(comp ? String(comp) : "—", {
      x: mx + 0.35, y: 2.85, w: mw - 0.7, h: 1.1, fontFace: FONT_H, fontSize: 60, bold: true, color: C.accent, isTextBox: true, margin: 0,
    });
    text(s, comp ? `competing ${typeWord === "brand" ? "brands" : "dispensaries"} in your service zone` : "competitors in your service zone (fill in market.competitors_in_zone)", {
      x: mx + 0.35, y: 4.0, w: mw - 0.7, h: 0.7, fontSize: 13, color: "FFFFFF",
    });
    text(s, (P.market && P.market.notes) || "", { x: mx + 0.35, y: 4.8, w: mw - 0.7, h: 1.7, fontSize: 12, color: "D9E5DD" });
    s.addNotes("Validate the channels before the pitch. Ask which platform they think drives the most orders — most owners are guessing, which is the opening for the audit slide.");
  }

  // 5. YOUR OPPORTUNITY (audit)
  {
    const s = base();
    title(s, `What we found on ${prospectName}'s listings`, { sub: "Every row is revenue currently left on the table." });
    const hasAudit = Array.isArray(P.audit) && P.audit.length > 0;
    const audit = hasAudit ? P.audit : [{ platform: "—", finding: "Audit pending", impact: "Run the audit before presenting" }];
    const head = { fill: { color: C.primary }, color: "FFFFFF", bold: true, fontFace: FONT_B, fontSize: 13, valign: "middle" };
    const cell = { fontFace: FONT_B, fontSize: 12.5, color: C.ink, valign: "middle" };
    const tableRows = [
      [{ text: "Platform", options: head }, { text: "What we found", options: head }, { text: "What it's costing you", options: head }],
      ...audit.map((r, i) => {
        const fill = { color: i % 2 ? "FFFFFF" : C.light };
        return [
          { text: r.platform || "", options: { ...cell, bold: true, color: C.secondary, fill } },
          { text: r.finding || "", options: { ...cell, fill } },
          { text: r.impact || "", options: { ...cell, fill } },
        ];
      }),
    ];
    s.addTable(tableRows, {
      x: M, y: 2.2, w: W - 2 * M, colW: [2.0, 5.3, 4.83], rowH: 0.6,
      border: { type: "solid", color: "DDE5DF", pt: 0.75 },
      margin: [0.08, 0.15, 0.08, 0.15],
    });
    {
      const y = 2.2 + 0.6 * (audit.length + 1) + 0.45;
      const h = H - 0.6 - y;
      if (hasAudit && h >= 0.9) {
        const nPlat = new Set(audit.map((r) => r.platform)).size;
        card(s, M, y, W - 2 * M, Math.min(h, 1.3), { fill: C.primary, line: C.primary });
        text(s, `${audit.length} gap${audit.length === 1 ? "" : "s"} across ${nPlat} platform${nPlat === 1 ? "" : "s"}`, { x: M + 0.35, y: y + 0.2, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: C.accent });
        text(s, "Each one maps to a service in your scope (slide 7). All of them are fixed inside the first 30 days of the launch plan.", { x: M + 0.35, y: y + 0.62, w: W - 2 * M - 0.7, h: 0.55, fontSize: 12.5, color: "FFFFFF" });
      }
    }
    s.addNotes("This is the Airbnb 'market size' slide turned into their market. Read it like a doctor reading a chart — calm, specific, no blame. Each finding maps to a service on the scope slide.");
  }

  // 6. HOW IT WORKS (seven pillars)
  {
    const s = base();
    title(s, "How it works: seven pillars on every platform", { sub: "The same operating system runs on Weedmaps, Leafly, Yelp and Google." });
    const pillars = [
      ["FaDollarSign", "Spend", "Bid, budget and negotiate placements; optimize for ROAS, not impressions."],
      ["FaStore", "Listing", "Hours, address, photos, attributes and Q&A kept accurate and conversion-ready."],
      ["FaClipboardList", "Menu", "Products verified, pricing and stock synced to POS, best-sellers featured."],
      ["FaStar", "Reviews", "Same-day responses; removal requests for guideline-breaking reviews."],
      ["FaTags", "Deals", "Daily and monthly promos, BOGOs, events and Google Posts on every platform."],
      ["FaPaintBrush", "Creative", "In-house graphics for every ad format and listing asset."],
      ["FaChartBar", "Data", "One dashboard with POS-integrated attribution across every channel."],
    ];
    const cols = 4, gw = 0.25;
    const cw = (W - 2 * M - (cols - 1) * gw) / cols;
    const ch = 2.05;
    for (let i = 0; i < pillars.length; i++) {
      const r = Math.floor(i / cols), c = i % cols;
      const x = M + c * (cw + gw), y = 2.2 + r * (ch + 0.25);
      card(s, x, y, cw, ch);
      await iconCircle(s, pillars[i][0], x + 0.25, y + 0.25, 0.55);
      text(s, pillars[i][1], { x: x + 0.95, y: y + 0.3, w: cw - 1.15, h: 0.45, fontSize: 16, bold: true, color: C.primary });
      text(s, pillars[i][2], { x: x + 0.25, y: y + 0.95, w: cw - 0.5, h: 1.05, fontSize: 12, color: C.ink });
    }
    // 8th cell = guarantee callout
    {
      const x = M + 3 * (cw + gw), y = 2.2 + ch + 0.25;
      card(s, x, y, cw, ch, { fill: C.accent, line: C.accent });
      text(s, "Guarantee", { x: x + 0.25, y: y + 0.3, w: cw - 0.5, h: 0.45, fontSize: 16, bold: true, color: C.primary });
      text(s, brand.guarantee, { x: x + 0.25, y: y + 0.85, w: cw - 0.5, h: 1.1, fontSize: 14, bold: true, color: C.primary });
    }
    s.addNotes("Product slide. Emphasize that these are operating routines with weekly cadence, not a one-time setup. Menu + Data are the two nobody else offers with POS integration.");
  }

  // 7. YOUR SCOPE OF SERVICES
  {
    const s = base();
    title(s, `Your scope of services`, { sub: `What ${brand.company} will run for ${prospectName}.` });
    const sv = P.services || {};
    const g = sv.google || {};
    const googleLabel = "Google" + (g.gbp || g.ads ? ` (${[g.gbp && "GBP", g.ads && "Ads"].filter(Boolean).join(" + ")})` : " (GBP, Ads)");
    const rows = [
      ["FaMapMarkedAlt", "Weedmaps Management", sv.weedmaps],
      ["FaLeaf", "Leafly Management", sv.leafly],
      ["FaYelp", "Yelp Management", sv.yelp],
      ["FaGoogle", `${googleLabel} Management`, g],
      ["FaMapMarkerAlt", "Local SEO", sv.local_seo],
    ];
    const others = Array.isArray(sv.other) ? sv.other.filter((o) => o && o.name) : [];
    for (const o of others) rows.push(["FaPlus", o.name, { include: true, notes: o.notes || "" }]);
    const rh = 0.66, gap = 0.1;
    const avail = H - 0.55 - 2.15;
    const rhUse = Math.min(rh, (avail - (rows.length - 1) * gap) / rows.length);
    for (let i = 0; i < rows.length; i++) {
      const [ic, name, cfg] = rows[i];
      const inc = !!(cfg && cfg.include);
      const y = 2.15 + i * (rhUse + gap);
      card(s, M, y, W - 2 * M, rhUse, { fill: inc ? C.light : "FFFFFF", line: inc ? C.light : "E3E8E4" });
      await iconCircle(s, ic, M + 0.15, y + (rhUse - 0.44) / 2, 0.44, inc ? C.secondary : "C9D2CC");
      text(s, name, { x: M + 0.75, y: y + 0.08, w: 3.5, h: rhUse - 0.16, fontSize: 14, bold: true, color: inc ? C.primary : C.muted, valign: "middle" });
      text(s, inc ? (cfg.notes || "Included") : "Not in this proposal — available as an add-on", {
        x: M + 4.35, y: y + 0.08, w: W - 2 * M - 5.95, h: rhUse - 0.16, fontSize: 11, color: inc ? C.ink : C.muted, valign: "middle",
      });
      // status pill
      const pw = 1.25, px = W - M - pw - 0.15;
      s.addShape(pres.ShapeType.roundRect, { x: px, y: y + (rhUse - 0.34) / 2, w: pw, h: 0.34, rectRadius: 0.17, fill: { color: inc ? C.accent : "EEF1EF" }, line: { color: inc ? C.accent : "EEF1EF", width: 0 } });
      text(s, inc ? "INCLUDED" : "NOT INCLUDED", { x: px, y: y + (rhUse - 0.34) / 2, w: pw, h: 0.34, fontSize: 9.5, bold: true, color: inc ? C.primary : C.muted, align: "center", valign: "middle", charSpacing: 1 });
    }
    s.addNotes("The commercial heart of the deck. Every one of the six standard services is shown, included or not, so the prospect sees the whole menu. Walk the included rows and tie each to an audit finding.");
  }

  // 8. RESULTS
  {
    const s = base(true);
    title(s, "The results our clients see", { dark: true, sub: "Averages across Sparks-managed accounts." });
    const proof = brand.proof || [];
    const n = proof.length || 1, gap = 0.3;
    const cw = (W - 2 * M - (n - 1) * gap) / n;
    proof.forEach((p, i) => {
      const x = M + i * (cw + gap);
      card(s, x, 2.4, cw, 3.4, { fill: "163826", line: "163826" });
      s.addText(p.stat, { x: x + 0.3, y: 2.65, w: cw - 0.4, h: 1.4, fontFace: FONT_H, fontSize: 50, bold: true, color: C.accent, isTextBox: true, margin: 0, valign: "top" });
      text(s, p.label, { x: x + 0.3, y: 4.15, w: cw - 0.6, h: 1.4, fontSize: 14, color: "FFFFFF" });
    });
    text(s, "Significant conversion gains typically show within 60–90 days as optimizations compound across every channel.", { x: M, y: 6.1, w: W - 2 * M, h: 0.5, fontSize: 13, color: "D9E5DD", italic: true });
    s.addNotes("Traction slide. Only cite numbers published on sparkscann.com. If asked how ROAS is measured: on-platform conversions plus POS-integrated attribution in the Sparks dashboard.");
  }

  // 9. CASE STUDIES
  {
    const s = base();
    title(s, "Case studies", { sub: "What full ownership of listings looks like in practice." });
    const keys = Array.isArray(P.case_studies) && P.case_studies.length ? P.case_studies : Object.keys(caseStudies);
    const cs = keys.map((k) => caseStudies[k]).filter(Boolean).slice(0, 2);
    const gap = 0.35, cw = (W - 2 * M - gap * (cs.length - 1)) / Math.max(cs.length, 1);
    cs.forEach((c, i) => {
      const x = M + i * (cw + gap), y = 2.2, h = 4.5;
      card(s, x, y, cw, h);
      text(s, c.name, { x: x + 0.35, y: y + 0.3, w: cw - 0.7, h: 0.45, fontSize: 20, bold: true, color: C.primary });
      text(s, `${c.type} · ${c.location}`, { x: x + 0.35, y: y + 0.75, w: cw - 0.7, h: 0.35, fontSize: 12, color: C.muted });
      text(s, [{ text: "Situation  ", options: { bold: true, color: C.secondary } }, { text: c.situation }], { x: x + 0.35, y: y + 1.2, w: cw - 0.7, h: 0.8, fontSize: 12 });
      text(s, [{ text: "What we did  ", options: { bold: true, color: C.secondary } }, { text: c.actions }], { x: x + 0.35, y: y + 2.0, w: cw - 0.7, h: 0.95, fontSize: 12 });
      const res = c.results || [];
      const sw = (cw - 0.7) / Math.max(res.length, 1);
      res.forEach((r, j) => {
        const sx = x + 0.35 + j * sw;
        s.addText(r.stat, { x: sx, y: y + 3.0, w: sw - 0.1, h: 0.75, fontFace: FONT_H, fontSize: res.length > 2 ? 24 : 30, bold: true, color: C.secondary, isTextBox: true, margin: 0 });
        text(s, r.label, { x: sx, y: y + 3.72, w: sw - 0.15, h: 0.7, fontSize: 11, color: C.ink });
      });
    });
    s.addNotes("Pick the case study closest to the prospect's situation (multi-location, delivery, new market). Purple Lotus is the 'we've done this for a decade-old account' story; Nuna Harvest is the 'new market, fast growth' story.");
  }

  // 10. YOUR OPTIONS (competition)
  {
    const s = base();
    title(s, "Your options", { sub: "How the alternatives compare." });
    const head = { fill: { color: C.primary }, color: "FFFFFF", bold: true, fontFace: FONT_B, fontSize: 12.5, valign: "middle", align: "center" };
    const cell = { fontFace: FONT_B, fontSize: 12, color: C.ink, valign: "middle", align: "center" };
    const first = { ...cell, bold: true, align: "left", color: C.primary };
    const yes = { ...cell, color: C.secondary, bold: true };
    const no = { ...cell, color: "B0453A" };
    const part = { ...cell, color: C.muted };
    const R = (label, a, b, c, d) => [
      { text: label, options: first },
      { text: a[0], options: a[1] }, { text: b[0], options: b[1] }, { text: c[0], options: c[1] }, { text: d[0], options: { ...d[1], fill: { color: "E9F7DC" } } },
    ];
    const Y = ["Yes", yes], N = ["No", no], PT = (t) => [t, part];
    const rows = [
      [{ text: "", options: head }, { text: "In-house", options: head }, { text: "Platform reps", options: head }, { text: "General agency", options: head }, { text: brand.company, options: { ...head, fill: { color: C.secondary } } }],
      R("Covers Weedmaps, Leafly, Yelp & Google", PT("Partially"), N, PT("Rarely"), Y),
      R("Marketplace auction & bidding expertise", N, PT("One platform"), N, Y),
      R("Incentive aligned to your ROAS", Y, N, PT("Varies"), Y),
      R("Cannabis compliance know-how", PT("Varies"), Y, N, Y),
      R("POS-integrated attribution dashboard", N, N, N, Y),
      R("Weekly menu, deal & review cadence", PT("If staffed"), N, N, Y),
      R("Results guarantee", N, N, N, Y),
    ];
    s.addTable(rows, { x: M, y: 2.2, w: W - 2 * M, colW: [3.9, 2.05, 2.05, 2.05, 2.08], rowH: 0.5, border: { type: "solid", color: "DDE5DF", pt: 0.75 }, margin: [0.05, 0.1, 0.05, 0.1] });
    s.addNotes("Competition slide. Be fair to the in-house option — it works if they have a full-time person who has bid on Weedmaps before. Most don't.");
  }

  // 11. WHY SPARKS (competitive advantages)
  {
    const s = base(true);
    title(s, `Why ${brand.company}`, { dark: true, sub: "Competitive advantages you can verify." });
    const adv = [
      ["FaUserTie", "Built by ex-Weedmaps operators", "Four years inside Weedmaps working with the largest dispensaries. We know how the auction, ranking and menu systems actually work."],
      ["FaLayerGroup", "All four platforms + 50 directories", "One team, one strategy, one report across Weedmaps, Leafly, Yelp, Google and every citation that feeds local SEO."],
      ["FaPlug", "POS-integrated attribution", "Treez, Dutchie and other POS data joined with ad spend and platform metrics — you see which platform, campaign and deal drove revenue."],
      ["FaShieldAlt", brand.guarantee, "We've done this for 100+ businesses. If the first 30 days don't show results, you don't pay."],
    ];
    const gap = 0.3, cw = (W - 2 * M - gap) / 2, ch = 1.95;
    for (let i = 0; i < adv.length; i++) {
      const x = M + (i % 2) * (cw + gap), y = 2.25 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch, { fill: "163826", line: "163826" });
      await iconCircle(s, adv[i][0], x + 0.3, y + 0.3, 0.6, C.accent, C.primary);
      text(s, adv[i][1], { x: x + 1.1, y: y + 0.3, w: cw - 1.4, h: 0.6, fontSize: 16, bold: true, color: "FFFFFF" });
      text(s, adv[i][2], { x: x + 1.1, y: y + 0.9, w: cw - 1.4, h: 1.0, fontSize: 12, color: "D9E5DD" });
    }
    s.addNotes("These are the four reasons to pick us over the options on the previous slide. The founder story is the credibility anchor; the dashboard is the differentiator no competitor matches.");
  }

  // 12. TEAM
  {
    const s = base();
    title(s, "Team", { sub: "The people who will run your account." });
    const team = (brand.team || []).filter((t) => t && t.name);
    const n = Math.max(team.length, 1), gap = 0.3;
    const cw = Math.min(4.0, (W - 2 * M - gap * (n - 1)) / n);
    team.forEach((t, i) => {
      const x = M + i * (cw + gap), y = 2.3;
      card(s, x, y, cw, 4.2);
      s.addShape(pres.ShapeType.ellipse, { x: x + 0.35, y: y + 0.35, w: 1.1, h: 1.1, fill: { color: C.secondary }, line: { color: C.secondary, width: 0 } });
      const initials = t.name.split(/\s+/).map((w) => w[0]).join("").slice(0, 2).toUpperCase();
      s.addText(initials, { x: x + 0.35, y: y + 0.35, w: 1.1, h: 1.1, fontFace: FONT_H, fontSize: 28, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
      text(s, t.name, { x: x + 0.35, y: y + 1.65, w: cw - 0.7, h: 0.45, fontSize: 18, bold: true, color: C.primary });
      text(s, t.title || "", { x: x + 0.35, y: y + 2.1, w: cw - 0.7, h: 0.35, fontSize: 12, color: C.secondary, bold: true });
      text(s, t.bio || "", { x: x + 0.35, y: y + 2.55, w: cw - 0.7, h: 1.5, fontSize: 11.5, color: C.ink });
    });
    // right-side: what a client team looks like
    const px = M + n * (cw + gap) + 0.2;
    if (px < W - M - 3) {
      const pw = W - M - px;
      card(s, px, 2.3, pw, 4.2, { fill: C.primary, line: C.primary });
      text(s, "YOUR ACCOUNT TEAM", { x: px + 0.35, y: 2.6, w: pw - 0.7, h: 0.35, fontSize: 11, color: C.accent, charSpacing: 4, bold: true });
      bullets(s, [
        "Dedicated account lead and weekly performance report",
        "Menu & listing specialists working daily in your platforms",
        "In-house creative for ads, deals and listing assets",
        "Data team maintaining your POS-integrated dashboard",
        "Direct line to leadership — no ticket queues",
      ], { x: px + 0.2, y: 3.05, w: pw - 0.5, h: 3.3, fontSize: 13, color: "FFFFFF" });
    }
    s.addNotes("Add team members in brand/config.yml. Headshots can be dropped in later by replacing the initials circles in PowerPoint.");
  }

  // 13. RECOGNITION
  {
    const s = base();
    title(s, "Recognition", { sub: "Press and partner credentials." });
    const items = [
      ["FaAward", "Top Cannabis Listing Management Agency 2026", "Cannabis Business Insights Magazine, based on an editorial interview with CEO Kevin Sparks."],
      ["FaYelp", "Yelp Ads Certified Partner", "One of the few cannabis partners recognized by Yelp, with direct access to Yelp Ads and review tools."],
      ["FaMicrophone", "Respect My Region podcast partner", "Three-episode series on optimizing Weedmaps, Leafly and Yelp for revenue."],
      ["FaMapMarkedAlt", "Weedmaps alumni-founded", "Founder spent four years inside Weedmaps before starting Sparks in 2022."],
    ];
    const gap = 0.3, cw = (W - 2 * M - gap) / 2, ch = 1.9;
    for (let i = 0; i < items.length; i++) {
      const x = M + (i % 2) * (cw + gap), y = 2.25 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch);
      await iconCircle(s, items[i][0], x + 0.3, y + 0.3, 0.6);
      text(s, items[i][1], { x: x + 1.1, y: y + 0.3, w: cw - 1.4, h: 0.6, fontSize: 16, bold: true, color: C.primary });
      text(s, items[i][2], { x: x + 1.1, y: y + 0.9, w: cw - 1.4, h: 0.9, fontSize: 12, color: C.ink });
    }
    s.addNotes("Press slide. Add logos here later if desired — the layout leaves room inside each card.");
  }

  // 14. INVESTMENT & NEXT STEPS
  {
    const s = base(true);
    title(s, "Investment & next steps", { dark: true, sub: `A 90-day launch plan for ${prospectName}.` });
    const inv = P.investment || {};
    const opts = Array.isArray(inv.options) ? inv.options.filter((o) => o && o.name) : [];
    const leftW = 7.3;
    // plan timeline (left)
    const plan = Array.isArray(P.plan) && P.plan.length ? P.plan : [
      { period: "Days 1–30", items: "Full audit, strategy, listing + menu clean-up, campaigns live" },
      { period: "Days 31–60", items: "Bid & deal optimization, review program, dashboard attribution" },
      { period: "Days 61–90", items: "Scale winners, Maps 3-pack push, quarterly plan" },
    ];
    plan.slice(0, 3).forEach((p, i) => {
      const y = 2.25 + i * 1.1;
      s.addShape(pres.ShapeType.ellipse, { x: M, y: y + 0.12, w: 0.55, h: 0.55, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
      s.addText(String(i + 1), { x: M, y: y + 0.12, w: 0.55, h: 0.55, fontFace: FONT_H, fontSize: 16, bold: true, color: C.primary, align: "center", valign: "middle", isTextBox: true, margin: 0 });
      text(s, p.period, { x: M + 0.8, y: y, w: leftW - 0.8, h: 0.4, fontSize: 15, bold: true, color: "FFFFFF" });
      text(s, p.items, { x: M + 0.8, y: y + 0.4, w: leftW - 0.8, h: 0.6, fontSize: 12, color: "D9E5DD" });
    });
    text(s, brand.guarantee, { x: M, y: 5.6, w: leftW, h: 0.4, fontSize: 14, bold: true, color: C.accent });
    text(s, inv.ad_spend_note || "", { x: M, y: 6.0, w: leftW, h: 0.5, fontSize: 11, color: "D9E5DD", italic: true });
    // pricing (right)
    const px = M + leftW + 0.4, pw = W - M - px;
    card(s, px, 2.25, pw, 4.3, { fill: "163826", line: "163826" });
    text(s, "INVESTMENT", { x: px + 0.35, y: 2.5, w: pw - 0.7, h: 0.35, fontSize: 11, color: C.accent, charSpacing: 4, bold: true });
    if (opts.length) {
      const rh = Math.min(0.8, 2.45 / opts.length);
      opts.slice(0, 4).forEach((o, i) => {
        const y = 2.95 + i * rh;
        text(s, o.name, { x: px + 0.35, y, w: pw - 2.4, h: 0.36, fontSize: 14, bold: true, color: "FFFFFF", valign: "middle" });
        text(s, o.price || "Custom", { x: px + pw - 2.05, y, w: 1.7, h: 0.36, fontSize: 14, bold: true, color: C.accent, align: "right", valign: "middle" });
        text(s, o.includes || "", { x: px + 0.35, y: y + 0.36, w: pw - 0.7, h: rh - 0.4, fontSize: 11, color: "D9E5DD" });
      });
    } else {
      s.addText(inv.monthly_fee || "Custom proposal", { x: px + 0.35, y: 2.95, w: pw - 0.7, h: 1.0, fontFace: FONT_H, fontSize: 40, bold: true, color: C.accent, isTextBox: true, margin: 0 });
      text(s, inv.term || "", { x: px + 0.35, y: 3.95, w: pw - 0.7, h: 0.5, fontSize: 13, color: "FFFFFF" });
    }
    const ct = brand.contact || {};
    const contactLine = [ct.name && `${ct.name}${ct.title ? ", " + ct.title : ""}`, ct.email, ct.phone, ct.booking_link].filter(Boolean).join("  ·  ");
    text(s, "Ready when you are.", { x: px + 0.35, y: 5.45, w: pw - 0.7, h: 0.4, fontSize: 16, bold: true, color: "FFFFFF" });
    text(s, contactLine, { x: px + 0.35, y: 5.9, w: pw - 0.7, h: 0.55, fontSize: 11, color: C.accent });
    s.addNotes(`Close: restate the 30-day guarantee, propose the start date, and ask for platform access (Weedmaps, Leafly, Yelp, GBP admin) to begin the audit. Term: ${inv.term || "see proposal"}.`);
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await pres.writeFile({ fileName: outPath });
  console.log(`Wrote ${path.relative(ROOT, outPath)} (${slideNo} slides)`);
}

build().catch((e) => {
  console.error(e);
  process.exit(1);
});
