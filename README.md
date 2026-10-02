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
| Planner | Brief, budget, cost per rating point, spot length, target reach, auto or manual channels, budget cut, manual split override, lock or exclude programs, health check, planner memo, reach build curve, competing slots |
| Scenarios | Save, load, compare (what changes), quick budget what-ifs, export spot plan CSV, print/PDF, copy memo |
| Duplication | Editable channel duplication matrix, gross vs net reach |
| Data | Sortable, searchable, paged table; CSV export; data quality checks |

Every chart bar, heatmap cell, table row and insight card opens a **detail drawer** (KPIs, trend, breakdown, airings) where you can drill further, filter the dashboard to it, or lock or exclude the program from the plan.

### Required columns
`Channel, Date, Start, Program, TVR`. Optional: `Day, End, Duration, Category, TVR Share %, Reach %`. Times like `20:09`, `8:09 PM` and Excel time cells are all handled.

### Models used
- Channel weight = 0.6 × avg Reach % + 0.4 × avg TVR (normalised across selected channels), with optional manual override.
- Program score = 0.6 × relative mean TVR + 0.4 × steadiness (mean ÷ std dev, capped at 5).
- Net reach: `A + B − overlap`, overlap = `A·B/100 + d·(min(A,B) − A·B/100)`, where `d` comes from the duplication matrix (defaults: Hiru–Derana 0.70, Hiru–Shakthi 0.20, …).
- Spots ≈ budget ÷ (CPRP × TVR × spot length/30).

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
public/js/engine.js  parsing, aggregation, duplication, plan allocation
public/js/app.js     UI, drill-down, planner, scenarios
public/js/ai.js      AI context builder, Gemini call, offline planner
public/js/demo.js    synthetic demo dataset (fictional)
public/js/store.js   IndexedDB / localStorage helpers
```
