# Sparks Pitch Deck Specialist

Generates professional, personalized pitch decks (`.pptx`) that Sparks uses to win
dispensaries, delivery services and cannabis brands as listing-management clients.
The structure follows the Slidebean Airbnb pitch deck (14 slides) adapted to an
agency sales pitch.

## Quick start

```bash
npm install
cp prospects/_template.yml prospects/my-dispensary.yml   # fill it in
npm run build -- prospects/my-dispensary.yml             # → decks/Sparks-x-my-dispensary.pptx
npm run build:all                                        # rebuild every prospect
```

An example is included: `prospects/example-green-valley.yml` →
`decks/Sparks-x-example-green-valley.pptx`.

## What a deck contains

Decks are split by platform: the focus platform (default Weedmaps) gets the full
pitch, then Leafly, Yelp and Google follow as one-slide previews, each marked
INCLUDED or ADD-ON at the price in `investment.addon_price`.

| # | Slide | Personalized by |
|---|---|---|
| 1 | Cover — Sparks × Prospect, "<Focus> Management proposal" | name, city, contact, date |
| 2 | The Problem | — |
| 3 | The Solution — platform tiles | `services`, `focus`, `addon_price` |
| 4 | Focus · Why it matters + your market | `market`, `locations` |
| 5 | Focus · What we found (audit) | `audit` rows for the focus platform |
| 6 | Focus · What Sparks will manage | `services.<focus>.notes` |
| 7 | Focus · Results | `brand/config.yml` proof |
| 8 | Focus · Case studies | `case_studies` |
| 9 | Focus · Investment | `investment.options` / `monthly_fee` |
| 10–12 | Add-on previews: Leafly, Yelp, Google | `services.<platform>.include`, audit rows |
| 13 | Why Sparks | — |
| 14 | Your account team | `brand/config.yml` team |
| 15 | Next steps + investment summary | `plan`, `investment`, contact |

Every slide has speaker notes. Branding (Sparks orange, wordmark, sparkles icon)
comes from `brand/config.yml` and `brand/assets/`.

## Files

- `prospects/_template.yml` — the intake form. Copy it per business. The
  `services` block is required and always lists all six service lines with
  `include: true/false` and prospect-specific notes; `other` is a free-form list.
- `brand/config.yml` — contact details, team, palette, proof points, guarantee.
- `brand/sparks.md` — everything we know about Sparks from sparkscann.com
  (services, pillars, stats, case studies, recognition). Source of truth for copy.
- `brand/deck-structure.md` — slide-by-slide rationale.
- `src/build_deck.js` — the generator (pptxgenjs). `src/platforms.yml` holds the
  per-platform pitch copy; `src/case_studies.yml` and `src/icons.js` support it.
- `brand/assets/` — logo variants used on the slides.
- `CLAUDE.md` — the workflow Claude follows when you name a new business.

## Customizing

- **Brand colors / logo / contact / team:** edit `brand/config.yml` (logo files in `brand/assets/`).
- **Platform pitch copy or add-on price:** `src/platforms.yml` and `investment.addon_price` in the prospect file.
- **New case study:** add a key to `src/case_studies.yml`, reference it in a
  prospect's `case_studies` list.
- **Pricing tiers:** fill `investment.options` in the prospect file, or leave the
  list empty and set `investment.monthly_fee` for a single price.
- **Headshots / logos:** open the `.pptx` in PowerPoint and replace the initials
  circles on the Team slide; the layout leaves room for logos on Recognition.
