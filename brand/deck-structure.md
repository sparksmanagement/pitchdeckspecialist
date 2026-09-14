# Deck Structure — Sparks Prospect Pitch

Modeled on the Slidebean Airbnb pitch deck (Cover, Problem, Solution, Market
Validation, Market Size, Product, Business Model, Traction, Competition, Team, Press,
Financials) and adapted from "raise money" to "win a dispensary as a client". The deck
is **split by platform**: the focus platform (default Weedmaps) gets the full pitch;
Leafly, Yelp and Google follow as one-slide previews, in that order.

Rules that carry over from the Airbnb deck:
- The viewer knows exactly what we do by slide 3, in plain words.
- One idea per slide. Big headline, one supporting visual, minimal body copy.
- Numbers are the argument. Every claim gets a stat or a case study.
- 15 slides, presentable in under 15 minutes.

| # | Slide | Airbnb equivalent | Content source |
|---|---|---|---|
| 1 | Cover — Sparks × Prospect, "<Focus> Management proposal" | Cover | prospect + brand |
| 2 | The Problem | Problem | brand |
| 3 | The Solution — platform tiles: THIS PROPOSAL / INCLUDED / ADD-ON BUNDLE $495 | Solution | services + investment.addon_price |
| 4 | **Focus · Why it matters** — 3 reasons + "your market" panel | Market validation | platforms.yml + prospect.market |
| 5 | **Focus · Audit** — what we found on their listings | Market size | prospect.audit (rows whose platform starts with the focus name) |
| 6 | **Focus · Scope** — what Sparks will manage (8 numbered lines) | Product | platforms.yml + services.<focus>.notes |
| 7 | **Focus · Results** — proof stats | Traction | brand.proof + platform proof |
| 8 | **Focus · Case studies** | Testimonials | case_studies.yml |
| 9 | **Focus · Investment** — pricing tiers, guarantee, add-on teaser | Business model | prospect.investment |
| 10 | Add-on preview · Leafly | — | platforms.yml (+ audit rows for Leafly) |
| 11 | Add-on preview · Yelp | — | platforms.yml |
| 12 | Add-on preview · Google (GBP, Ads, Local SEO) | — | platforms.yml |
| 13 | Why Sparks | Competitive advantages + Press | brand |
| 14 | Your account team | Team | brand/config.yml |
| 15 | Next steps — 90-day plan, investment summary, contact | Financials / ask | prospect.plan + investment |

If a non-focus platform is `include: true` in the prospect file, its preview slide is
labeled INCLUDED instead of ADD-ON and the summary shows "Included". Speaker notes are
generated for every slide.

## Branding

Sparks orange `DD5F13` on white, charcoal body text, thin white rules on orange
slides, the SPARKS wordmark on the cover and closing slide, and the sparkles icon in
every footer — matching the Sparks 2025 proposal deck. Logo variants live in
`brand/assets/`; palette and logo paths in `brand/config.yml`.
