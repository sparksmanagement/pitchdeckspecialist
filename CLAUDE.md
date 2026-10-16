# Pitch Deck Specialist — instructions for Claude

This repo produces sales pitch decks for **Sparks** (sparkscann.com), a cannabis
listing-management agency that manages Weedmaps, Leafly, Yelp and Google for
dispensaries, deliveries and brands. The user will name a prospective business;
your job is to turn that into a finished `.pptx`.

## Workflow when the user gives you a business

1. **Research the prospect.** Weedmaps and sparkscann.com serve bot-challenge pages
   to curl and the WebFetch tool; load them with the pre-installed headless Chromium
   via Playwright instead (default profile, no UA spoofing, `proxy: { server:
   process.env.HTTPS_PROXY }`; if you get `ERR_CERT_AUTHORITY_INVALID`, register the
   proxy CA once: `apt-get install -y libnss3-tools && certutil -d sql:$HOME/.pki/nssdb
   -A -t "C,," -n ccr-agent-proxy -i /root/.ccr/agent-proxy-ca.crt`). Screenshot each
   listing at a 430×932 viewport, save `document.body.innerText`, and crop the worst
   listing's screenshot into `prospects/assets/<slug>-<platform>.png` for
   `audit_mock.screenshot`. Also use web search for the business name + city: type
   (dispensary / delivery / brand), locations, website, and anything visible about
   their Weedmaps, Leafly, Yelp and Google presence (rating, review count, deals,
   whether GBP is claimed, competitor density). Do not invent facts; if you
   cannot verify a listing detail, write it as a question for the user in the
   audit notes rather than a finding. For the Google rating, load
   `https://www.google.com/maps/search/<business name city>?hl=en` in headless
   Chromium — Maps works from the cloud environment and shows the live star
   rating (Google Search returns a bot challenge and Yelp rejects datacenter
   traffic, so those two can't be read directly). Yelp ratings and the Google
   review count therefore come from third-party directories via web search and
   are often stale — label them "per <source>" and ask the user for a screenshot
   of the live Google Business Profile; a visible "Own this business?" link on
   it means the profile is unclaimed, which is a finding in itself.
2. **Create the prospect file.** Copy `prospects/_template.yml` to
   `prospects/<slug>.yml` and fill every section. Set `focus` (default
   `weedmaps`) — that platform gets the full pitch; the others become one-slide
   add-on previews priced as a bundle via `investment.addon`. Sparks standard
   pricing: Single Platform Management $495/mo for any one platform (the focus
   platform, not necessarily Weedmaps), scaling to 10% of that platform's ad spend
   above $4,950/mo; the Listing Bundle adds every other platform for +$495/mo;
   `investment.total` is the $990/mo Total Package. Keep the card named "Single
   Platform Management" and say which platform it is in `includes`. `addon_price` is the per-platform fallback. The à-la-carte add-ons under the pricing cards (Review Generation $100/mo per location + messaging fees; Product Image Management $10/hr) come from `brand/config.yml` → `extras` and are never added to the package total; override per prospect with `investment.extras`. `investment.trial` (free
   30-day trial, no invoice until day 31) shows on the cover, investment and
   next-steps slides when set. The
   `services` block must always address all six: `weedmaps`, `leafly`, `yelp`,
   `google` (with `gbp` and `ads` flags), `local_seo`, and `other` (a list; `[]`
   if none). `include: true` on a non-focus platform marks its preview slide
   INCLUDED instead of ADD-ON. Prefix each `audit` row's `platform` with the
   platform name (e.g. "Weedmaps · Sacramento") so rows land on the right slide.
   Add `audit_mock` (see the template) to show the worst listing as an iPhone
   mockup with numbered callouts instead of the table; a real screenshot can be
   dropped in via `audit_mock.screenshot`. Never pitch Yelp Ads (not available
   to cannabis right now); directories are "20+", not 50. Say "official partner"
   only for Yelp, Google and Apple Maps; for Weedmaps/Leafly say "10+ years
   working inside and alongside". Founder timeline, always in this order: 3 years
   at Purple Lotus, then 4 years at Weedmaps as an executive, then 4+ years running
   Sparks ("over 10 years in cannabis"). Never name
   Google Ads or Meta as platforms that block cannabis — say "most mainstream ad
   networks and social platforms restrict cannabis"; Sparks plans to sell those
   channels later.
3. **Build:** `npm run deck -- prospects/<slug>.yml` → `decks/Sparks-x-<slug>.pptx`,
   and uploads it to Google Drive when a credential is configured (step 6).
   Add `--short` for the 4-slide version (cover → scope → focus-platform investment
   → next steps) at `decks/Sparks-x-<slug>-short.pptx`; most prospects don't read
   the full deck, so build both unless the user says otherwise. `npm run publish --
   prospects/<slug>.yml --short` uploads the short one.
4. **QA:** validate and render every slide, and look at the images.
   ```bash
   python3 <pptx-skill>/scripts/office/validate.py decks/Sparks-x-<slug>.pptx
   python3 <pptx-skill>/scripts/office/soffice.py --headless --convert-to pdf --outdir /tmp/qa decks/Sparks-x-<slug>.pptx
   ```
   (`libreoffice-impress` may need `apt-get install`; `pymupdf` renders the PDF
   pages to images.) Fix overflow or collisions in `src/build_deck.js`, never by
   hand-editing the pptx.
5. **Deliver:** commit the prospect file and the deck, push, and tell the user
   what you assumed (pricing, audit findings you could not verify).
6. **File it in Google Drive.** Every client gets a folder named after
   `prospect.name` inside the "Pitch Decks" folder
   (`brand/config.yml` → `drive.pitch_decks_folder`,
   https://drive.google.com/drive/folders/1bcAP65vMoyceh0UHXUFprdSmiOtIfTXf), and every
   deck for that client goes in it. `npm run deck` does this automatically when
   `DRIVE_UPLOAD_URL` + `DRIVE_UPLOAD_KEY` are set (the Apps Script uploader in
   `drive/AppsScript.gs`; setup in `drive/README.md`); `npm run publish --
   prospects/<slug>.yml` uploads an already-built deck. Put the Drive link in the
   summary to the user. If no credential is configured, say so, create the client
   folder with the Google Drive connector (`create_file` with the folder mime type
   under that parent), and attach the .pptx for the user to drop in — the connector
   cannot carry a multi-megabyte file as base64.

## Where things live

- `brand/sparks.md` — the Sparks knowledge base: services, seven pillars, proof
  points, case studies, competitive angle. Only cite numbers that are in here.
- `brand/config.yml` — contact, team, palette, proof stats, guarantee.
- `brand/deck-structure.md` — the 14-slide structure (Slidebean Airbnb deck
  adapted to an agency sales pitch) and what feeds each slide.
- `src/build_deck.js` — pptxgenjs generator. `src/platforms.yml` — per-platform
  pitch copy (why it matters, what we manage, proof). `src/case_studies.yml` —
  case study data. `src/icons.js` — react-icons → PNG.
- `brand/assets/` — Sparks logo variants (orange/white/black full logo, wordmark,
  sparkles icon). Palette and logo paths are in `brand/config.yml`.
- `prospects/` — one YAML per business. `decks/` — output.

## Rules

- Decks are split by platform in the order Weedmaps → Leafly → Yelp → Google.
  The focus platform gets ~80% of the deck; every other platform still gets its
  preview slide (included or add-on) so all services are always presented.
- Keep the deck at 14 slides; add content by editing the prospect file or
  `src/platforms.yml`, not by bolting on slides, unless the user asks.
- Brand: Sparks orange `DD5F13`, white, charcoal. Use the logo assets in
  `brand/assets/`; never recolor the logo outside orange/white/black.
- Proof points and case studies come only from `brand/sparks.md`; if the
  website changes, update that file first.
- Fonts stay Arial (headings) and Calibri (body) so QA renders are trustworthy.
- Never put a model name in commits, files, or the deck.
