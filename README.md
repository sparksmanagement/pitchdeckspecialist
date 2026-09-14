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

| # | Slide | Personalized by |
|---|---|---|
| 1 | Cover — Sparks × Prospect | prospect name, city, contact, date |
| 2 | The Problem | — |
| 3 | The Solution | prospect type |
| 4 | Where your customers are | city, competitor count, market notes |
| 5 | What we found on your listings | `audit` table |
| 6 | How it works — seven pillars | — |
| 7 | **Your scope of services** — Weedmaps, Leafly, Yelp, Google (GBP/Ads), Local SEO, Other | `services` block |
| 8 | Results we deliver | `brand/config.yml` proof |
| 9 | Case studies | `case_studies` keys |
| 10 | Your options (competition) | — |
| 11 | Why Sparks | — |
| 12 | Team | `brand/config.yml` team |
| 13 | Recognition | — |
| 14 | Investment & next steps | `investment`, `plan`, contact |

Every slide has speaker notes.

## Files

- `prospects/_template.yml` — the intake form. Copy it per business. The
  `services` block is required and always lists all six service lines with
  `include: true/false` and prospect-specific notes; `other` is a free-form list.
- `brand/config.yml` — contact details, team, palette, proof points, guarantee.
- `brand/sparks.md` — everything we know about Sparks from sparkscann.com
  (services, pillars, stats, case studies, recognition). Source of truth for copy.
- `brand/deck-structure.md` — slide-by-slide rationale.
- `src/build_deck.js` — the generator (pptxgenjs). `src/case_studies.yml`,
  `src/icons.js` support it.
- `CLAUDE.md` — the workflow Claude follows when you name a new business.

## Customizing

- **Brand colors / contact / team:** edit `brand/config.yml`.
- **New case study:** add a key to `src/case_studies.yml`, reference it in a
  prospect's `case_studies` list.
- **Pricing tiers:** fill `investment.options` in the prospect file, or leave the
  list empty and set `investment.monthly_fee` for a single price.
- **Headshots / logos:** open the `.pptx` in PowerPoint and replace the initials
  circles on the Team slide; the layout leaves room for logos on Recognition.
