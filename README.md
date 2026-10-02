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
| Scenarios | Save, load, compare (what changes), quick budget what-ifs, **client deck (PPTX)**, spot plan CSV, print/PDF, copy memo |
| Duplication | Editable channel duplication matrix, gross vs net reach |
| Data | Sortable, searchable, paged table; CSV export; data quality checks |

Every chart bar, heatmap cell, table row and insight card opens a **detail drawer** (KPIs, trend, breakdown, airings) where you can drill further, filter the dashboard to it, or lock or exclude the program from the plan.

### Client deck (PowerPoint)
"Export deck (PPTX)" on the Planner or Scenarios tab builds a 14-slide deck in the browser (pptxgenjs, loaded on demand from jsDelivr): title, executive summary, brief and approach, when the audience watches, channel strength, recommended split with a reason per channel, program basket, tier strategy, anchor programs, net reach after duplication, weekly flighting, budget scenarios, risks and next steps, methodology. Charts and tables are native and editable; every slide has speaker notes. With `GEMINI_API_KEY` set, Gemini drafts the headline, summary, per-channel reasons, risks and next steps; otherwise they are written from the plan numbers.

### Required columns
`Channel, Date, Start, Program, TVR`. Optional: `Day, End, Duration, Category, TVR Share %, Reach %`. Times like `20:09`, `8:09 PM` and Excel time cells are all handled.

### Models used
- **Channels** are short-listed on 0.6 × avg Reach % + 0.4 × avg TVR (or picked manually).
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
