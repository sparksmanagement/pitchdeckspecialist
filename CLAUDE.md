# Pitch Deck Specialist — instructions for Claude

This repo produces sales pitch decks for **Sparks** (sparkscann.com), a cannabis
listing-management agency that manages Weedmaps, Leafly, Yelp and Google for
dispensaries, deliveries and brands. The user will name a prospective business;
your job is to turn that into a finished `.pptx`.

## Workflow when the user gives you a business

1. **Research the prospect.** Use web search for the business name + city: type
   (dispensary / delivery / brand), locations, website, and anything visible about
   their Weedmaps, Leafly, Yelp and Google presence (rating, review count, deals,
   whether GBP is claimed, competitor density). Do not invent facts; if you
   cannot verify a listing detail, write it as a question for the user in the
   audit notes rather than a finding.
2. **Create the prospect file.** Copy `prospects/_template.yml` to
   `prospects/<slug>.yml` and fill every section. The `services` block must
   always address all six: `weedmaps`, `leafly`, `yelp`, `google` (with `gbp`
   and `ads` flags), `local_seo`, and `other` (a list; `[]` if none). Set
   `include: false` for anything out of scope so it still shows as
   "not included" on the scope slide. Write per-platform `notes` that are
   specific to this prospect's audit findings.
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
- `src/build_deck.js` — pptxgenjs generator. `src/case_studies.yml` — case
  study data. `src/icons.js` — react-icons → PNG.
- `prospects/` — one YAML per business. `decks/` — output.

## Rules

- Every deck covers all six services (Weedmaps, Leafly, Yelp, Google GBP/Ads,
  Local SEO, Other) on the scope slide, included or not.
- Keep the deck at 14 slides; add content by editing the prospect file, not by
  bolting on slides, unless the user asks.
- Proof points and case studies come only from `brand/sparks.md`; if the
  website changes, update that file first.
- Fonts stay Cambria (headings) and Calibri (body) so QA renders are trustworthy.
- Never put a model name in commits, files, or the deck.
