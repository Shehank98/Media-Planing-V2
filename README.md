# TV Media Planner

A browser-only TV media planning dashboard. Upload a TV ratings export (Excel or CSV), explore when and where people watch, drill into any channel or program, build a channel split and program basket with net reach, then test and export scenarios.

## Privacy model
- Files are parsed **in the browser** (SheetJS) and saved in that browser's IndexedDB only.
- User A's data is never visible to user B. B sees an empty dashboard until they upload their own file.
- Planner AI sends only a small **aggregated summary** of the filtered numbers to the server, which passes it to Gemini. The raw file is never sent.
- "Remove my data from this browser" (Data panel) clears it.

## Features
| Tab | What it does |
|---|---|
| Overview | Insight cards, KPIs, channel performance, top programs, category mix, daily trend, dayparts |
| Explore | Day/channel/category × hour heatmap (TVR, Reach %, Share %, Airings), day of week, TVR by hour, dayparts by channel, **Slot finder** ("everything at 10 PM on Hiru, Derana, Sirasa") |
| Drill down | Configurable tree (e.g. Channel ▸ Category ▸ Program ▸ airings) |
| Planner | Tier pyramid with Reach / Balanced / Frequency presets, spot-by-spot optimiser, spot caps per program per week, daypart min/max, weekly flighting (even, burst, pulse; roadblock or stagger), net reach and reach 3+, lock or exclude programs, fixed channel shares, health check, memo, reach build curve |
| Schedule | Booking sheet per channel like an agency schedule: programme rows grouped by creative, day / from / to / duration, TVR, GRP (TVR × spots), NGRP (GRP × duration ÷ 30), spots, rate card rate, negotiated rate, All Exposure Value (rate card × spots), Media Value (negotiated × spots), Investment 100%, SR Value (default 85%), one column per campaign date, SSCL and VAT on investment and SR value, asset split (paid value and ratio per creative) and Com Only CPRP (media value ÷ GRP) / NCPRP (media value ÷ NGRP). Editable 30-sec rate card per programme and discount per channel. **Export Excel** (cover, one sheet per channel, spot list) with live formulas |
| Scenarios | Save, load, compare (what changes), quick budget what-ifs, **client deck (PPTX)**, spot plan CSV, print/PDF, copy memo |
| Duplication | Editable channel duplication matrix, gross vs net reach |
| Data | Sortable, searchable, paged table; CSV export; data quality checks |

Every chart bar, heatmap cell, table row and insight card opens a **detail drawer** (KPIs, trend, breakdown, airings) where you can drill further, filter the dashboard to it, or lock or exclude the program from the plan.

### Client deck (PowerPoint)
"Export deck (PPTX)" (Planner or Scenarios tab) asks for client, campaign and "prepared by", then builds a client-ready deck in the browser (pptxgenjs, loaded on demand from jsDelivr). It is written for non-technical readers:
- Cover, agenda, executive summary with the decision needed, and a plain-English glossary (TVR, reach, frequency, reach 3+, spot, duplication, daypart, tier) using examples from the plan.
- Sections with divider slides: the audience (hours, days, channels), our recommendation (approach, channel split with a reason per channel, tier pyramid, programme selection, anchor programmes), reach and schedule (overlap diagram, how many channels, weekly booking calendar), budget options and risks (risk / why it matters / what we will do), next steps (dated timeline, decisions needed).
- Every content slide has an action title, a "What this means" box and "How to read this" captions under charts; every slide has speaker notes.
- Optional appendix: method, assumptions, full programme list.
- Charts and tables are native and editable. With `GEMINI_API_KEY` set, Gemini can write the commentary; otherwise it is written from the plan numbers.

### Creatives, campaign period and rates
- **Creatives** (Planner settings): brand / version, duration (5–60 sec) and rotation share. Spot cost = 30-sec rate × duration ÷ 30; spots are rotated between creatives by share.
- **Campaign period**: start and end date. Spots are placed only on dates inside the period on which the programme airs, at most N per programme per day and the weekly cap.
- **Rates**: the 30-sec rate card is estimated as max(minimum rate, CPRP × TVR) until you type the real value in the Schedule tab; negotiated rate = rate card × (1 − channel discount). The optimiser uses negotiated costs.

### Required columns
`Channel, Date, Start, Program, TVR`. Optional: `Day, End, Duration, Category, TVR Share %, Reach %`. Times like `20:09`, `8:09 PM` and Excel time cells are all handled.

### Models used
- **Channels** are short-listed on 0.6 × avg Reach % + 0.4 × avg TVR (or picked manually). The same score sets each channel's budget share (or type your own fixed shares). Spots are bought within each channel's budget, then topped up until the next spot no longer fits.
- **Programme basket**: on each channel, the best N programmes of every tier that has budget (N = "Best programmes per channel per tier"; best = average TVR × steadiness factor), plus locked programmes and any you add with "+ Add programme" in the basket. Every basket programme gets a first spot, then spots are added within the channel budget until the next spot no longer fits. Removing a programme lets the next best one take its place.
- **Creatives**: each creative's % is its share of the **budget**. Spots are rotated so the paid value per creative ends close to that %; a longer creative gets fewer spots for the same money.
- **Plan settings** are edited as a draft and only used when you click **Apply** (or **Discard** to go back).
- **Day placement** (Flighting, default "Best day, split overlaps"): each spot goes on the programme's best-rated air day; if a same-hour programme with audience overlap at or above the limit (default 0.50, from the Duplication factors) is already on that day, the next best day is used. Roadblock and stagger are still available.
- **Duplication by time belt** (Planner): per hour of the plan, reach added up vs different people, % duplicated, and whether each high-overlap pair was kept on different days.
- **Tiers:** eligible programs (avg TVR at or above "Ignore programs below TVR", default 0.5) are ranked by average TVR. Tier 1 = at or above the 75th percentile, Tier 3 = below the 25th, Tier 2 in between (both editable). Presets: Reach 65/35/0, Balanced 45/35/20, Frequency 30/40/30.
- **Optimiser:** buys one spot at a time where it adds the most new net reach per rupee (weighted by steadiness = mean ÷ std. dev.). Order: locked programs, fixed channel shares and daypart minimums, each tier up to its budget, then leftover. Caps: spots per program per week, programs per channel per tier, daypart maximums.
- **Cost per spot** = max(minimum spot rate, CPRP × TVR × length/30).
- **Reach:** each repeat spot of a program adds a set % (default 40%) of the previous spot's new reach. Programs on one channel are combined with an intra-channel overlap (0.8); channels with the duplication matrix: `A + B − overlap`, overlap = `A·B/100 + d·(min(A,B) − A·B/100)`.
- **Reach 3+** assumes a zero-truncated Poisson frequency spread at the plan's average frequency.
- **Flighting:** spots spread over the weeks by pacing weights (even, burst = descending, pulse = alternate weeks), never above the weekly cap, on each program's air days. Roadblock puts rival same-hour spots on the same nights; stagger spreads them.

## Run locally
```bash
npm start            # http://localhost:3000
```
No dependencies. Node 18+.

## Deploy on Railway
1. New Project → Deploy from GitHub repo → pick this repo. Railway detects Node and runs `npm start`.
2. Optional: Variables → `GEMINI_API_KEY` = your Google AI Studio key (and `GEMINI_MODEL`, default `gemini-2.5-flash`).
3. Settings → Networking → Generate Domain.

Without `GEMINI_API_KEY`, Planner AI still works with a built-in offline planner (budget cuts, best slots, programs at an hour, Tamil option, channel-count what-ifs).

## Structure
```
server.js            static server + /api/ai Gemini proxy (rate limited)
public/index.html    layout
public/css/app.css   styles (light and dark)
public/js/engine.js  parsing, aggregation, duplication, tiered optimiser, flighting
public/js/app.js     UI, drill-down, planner, scenarios
public/js/ai.js      AI context builder, Gemini call, offline planner
public/js/deck.js    PowerPoint deck export with rationale
public/js/demo.js    synthetic demo dataset (fictional)
public/js/store.js   IndexedDB / localStorage helpers
```
