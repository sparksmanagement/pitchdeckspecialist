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
   audit notes rather than a finding.
2. **Create the prospect file.** Copy `prospects/_template.yml` to
   `prospects/<slug>.yml` and fill every section. Set `focus` (default
   `weedmaps`) — that platform gets the full pitch; the others become one-slide
   add-on previews priced as a bundle via `investment.addon` (Sparks standard:
   Leafly + Yelp + Google for $495/mo total, scaling to 10% of ad spend above
   $4,950/mo on any platform; `investment.total` shows the package price,
   $1,490/mo with Weedmaps at $995/mo). `addon_price` is the per-platform fallback. `investment.trial` (free
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
   to cannabis right now); directories are "20+", not 50.
3. **Build:** `npm run build -- prospects/<slug>.yml` → `decks/Sparks-x-<slug>.pptx`.
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
- Keep the deck at 15 slides; add content by editing the prospect file or
  `src/platforms.yml`, not by bolting on slides, unless the user asks.
- Brand: Sparks orange `DD5F13`, white, charcoal. Use the logo assets in
  `brand/assets/`; never recolor the logo outside orange/white/black.
- Proof points and case studies come only from `brand/sparks.md`; if the
  website changes, update that file first.
- Fonts stay Arial (headings) and Calibri (body) so QA renders are trustworthy.
- Never put a model name in commits, files, or the deck.
