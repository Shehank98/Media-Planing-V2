import {
  DAYS, DS, DIMS, DPS, daypart, esc, nf, ni, hl, band, chName, pn, fmtDate, dayDiff, lkr,
  parseRows, grp, avgT, avgR, avgS, ciOf, steadiness, summarize, mkGetD, dkey, netReach, uni,
  channelStats, planCalc, scen, competing, healthChecks, insights, toCSV, STRATS, buildSchedule, campaignDays, creativeMix, avgLen
} from './engine.js';
import { exportScheduleXlsx } from './xlsx.js';
import { makeDemoRows } from './demo.js';
import { idbGet, idbSet, idbDel, lsGet, lsSet } from './store.js';
import { buildContext, askRemote, askLocal, splitAction } from './ai.js';
import { buildDeck } from './deck.js';

const $ = s => document.querySelector(s);
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

/* ---------- state ---------- */
const nextMonday = () => { const d = new Date(); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); return d.toISOString().slice(0, 10); };
const DEFAULT_P = {
  budget: 10000000, nCh: 4, nProg: 2, cut: 0, cprp: 25000, spotLen: 30, minRate: 15000, target: 0, chMode: 'auto', chPick: [], split: {}, lock: [], excl: [],
  strategy: 'balanced', tiers: [45, 35, 20], tp1: 75, tp3: 25, minTvrPlan: .5, capWk: 3, dpMin: [0, 0, 50, 0], dpMax: [10, 20, 100, 15],
  weeks: 4, start: nextMonday(), end: '', perDay: 1, pacing: 'even', same: 'roadblock', repQ: 40,
  creatives: [{ name: 'Creative A', dur: 30, share: 100 }], rates: {}, disc: {}, sscl: 2.5, vat: 18
};
const S = {
  ROWS: [], F: [], CH: [], CATS: [], SPAN: ['', ''], HRS: [0, 23], meta: null,
  FS: null, TAB: 'overview', DT: 'upload',
  P: Object.assign({}, DEFAULT_P, lsGet('plan', {})),
  DUP: lsGet('dup', {}), DINTRA: lsGet('dintra', .8),
  scenarios: lsGet('scenarios', []), cmp: null,
  plan: null, base: null, cur: null, health: [],
  EX: { mode: 'channel', metric: 'tvr' }, SF: { h: 20, ch: null },
  DR: { levels: ['ch', 'cat', 'p'], open: new Set() },
  DSORT: { k: 'tvr', d: -1 }, DPAGE: 0, DQ: '',
  X: [], XB: null,
  fv: 'weeks', sched: null, SCH: { ch: null }, SD: { h: null, sdDays: 'all', sdScope: 'plan' }, BK: { view: 'channel', tier: 'all', open: new Set(), all: false },
  chat: [], aiOn: false, aiModel: null, busy: false, editSplit: false
};
const getD = (a, b) => mkGetD(S.DUP)(a, b);
// Older saved plans: derive the campaign end date and creatives.
if (!S.P.end) S.P.end = campaignDays(S.P).end;
if (!Array.isArray(S.P.creatives) || !S.P.creatives.length) S.P.creatives = [{ name: 'Creative A', dur: S.P.spotLen || 30, share: 100 }];
['rates', 'disc'].forEach(k => { if (!S.P[k] || typeof S.P[k] !== 'object') S.P[k] = {}; });
const mixText = P => { const m = creativeMix(P); return m.map(c => `${c.name} ${c.dur}s${m.length > 1 ? ' (' + nf(c.w * 100, 0) + '%)' : ''}`).join(', '); };
S.getD = getD;
const charts = {};
const savePlan = () => lsSet('plan', S.P);

/* ---------- helpers ---------- */
const cIdx = ch => (Math.max(0, S.CH.indexOf(ch)) % 9) + 1;
const cc = ch => css('--c' + cIdx(ch));
const cv = ch => 'var(--c' + cIdx(ch) + ')';
const dot = ch => `<i class="dot" style="background:${cv(ch)}"></i>`;
const xa = f => `data-x="${esc(JSON.stringify(f))}"`;
const panel = (t, s, body, extra = '', pb = '') => `<div class="panel"><div class="ph"><h3>${t}</h3>${s ? `<span class="s">${s}</span>` : ''}${extra ? `<span class="push">${extra}</span>` : ''}</div><div class="pb ${pb}">${body}</div></div>`;
const kp = (l, v, s, cls = '') => `<div class="kpi ${cls}"><span>${l}</span><b>${v}</b><small title="${esc(String(s).replace(/<[^>]+>/g, ''))}">${s}</small></div>`;
const seg = (k, obj, opts) => `<div class="seg" data-seg="${k}">${opts.map(o => `<button data-v="${o[0]}" class="${obj[k] === o[0] ? 'on' : ''}">${o[1]}</button>`).join('')}</div>`;
const stdTag = (ci, n) => { const s = steadiness(ci, n); return `<span class="std ${s.cls}">${s.label}${n >= 2 ? ' · ' + nf(ci, 1) : ''}</span>`; };
const svgI = d => `<svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;
const ICON = { ok: svgI('<path d="M5 12l5 5L19 7"/>'), wa: svgI('<path d="M12 6v8M12 18.5v.5"/>'), in: svgI('<path d="M12 11v7M12 6.5v.5"/>') };
const roleTag = r => `<span class="tag ${r}">${{ anchor: 'Anchor', support: 'Support', review: 'Review' }[r]}</span>`;
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('on'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('on'), 2200); }
function download(name, text, type = 'text/csv;charset=utf-8') {
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['﻿' + text], { type })); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

/* ---------- data lifecycle ---------- */
function setData(rows, meta, persist = true) {
  S.ROWS = rows; S.meta = meta;
  const t = new Map(); rows.forEach(r => t.set(r.ch, (t.get(r.ch) || 0) + r.tvr));
  S.CH = [...t.keys()].sort((a, b) => t.get(b) - t.get(a));
  S.CATS = [...new Set(rows.map(r => r.cat))].sort();
  let lo = '9999', hi = '0000', h0 = 23, h1 = 0;
  rows.forEach(r => { if (r.date < lo) lo = r.date; if (r.date > hi) hi = r.date; if (r.h < h0) h0 = r.h; if (r.h > h1) h1 = r.h; });
  S.SPAN = [lo, hi]; S.HRS = [h0, h1];
  // Drop plan references to channels or programs not in this dataset.
  S.P.chPick = S.P.chPick.filter(c => S.CH.includes(c));
  const keys = new Set(rows.map(r => r.ch + '||' + r.p));
  S.P.lock = S.P.lock.filter(k => keys.has(k)); S.P.excl = S.P.excl.filter(k => keys.has(k));
  Object.keys(S.P.split).forEach(c => { if (!S.CH.includes(c)) delete S.P.split[c]; });
  S.SF.ch = null; S.DR.open.clear(); S.chat = [];
  resetFilters(false);
  if (persist) idbSet('dataset', { rows, meta });
  applyFilters();
  renderAICtx();
}
function emptyFS() { return { from: S.SPAN[0], to: S.SPAN[1], ch: new Set(S.CH), cat: new Set(S.CATS), day: new Set(DAYS), h0: S.HRS[0], h1: S.HRS[1], minTvr: 0, q: '', p: '' }; }
function resetFilters(apply = true) { S.FS = emptyFS(); if (apply) { applyFilters(); if ($('#drawer').classList.contains('on')) renderDrawer(); } }
function filterRows(f) {
  const q = (f.q || '').trim().toLowerCase();
  return S.ROWS.filter(r => r.date >= f.from && r.date <= f.to && f.ch.has(r.ch) && f.cat.has(r.cat) && f.day.has(r.day) && r.h >= f.h0 && r.h <= f.h1 && r.tvr >= f.minTvr && (!q || r.p.toLowerCase().includes(q)) && (!f.p || r.p === f.p));
}
function applyFilters() {
  if (!S.ROWS.length) { chrome(); renderActive(); return; }
  S.F = filterRows(S.FS);
  S.DPAGE = 0;
  recalc(); chrome(); renderActive(); renderAICtx();
}
function recalc() {
  S.plan = S.F.length ? planCalc(S.F, S.P) : null;
  if (!S.plan || !S.plan.items.length) { S.base = S.cur = null; S.health = []; return; }
  S.base = scen(S.plan, 1, S.P, getD, S.DINTRA);
  S.cur = S.P.cut > 0 ? scen(S.plan, 1 - S.P.cut / 100, S.P, getD, S.DINTRA) : S.base;
  if (S.P.cut > 0) { const ks = new Set(S.cur.kept.map(x => x.key)); S.cur.dropped = S.base.kept.filter(x => !ks.has(x.key)); }
  if (!S.cur.kept.length) { S.base = S.cur = null; S.health = []; S.sched = null; return; }
  S.sched = buildSchedule(S.cur, S.P);
  S.health = healthChecks(S.F, S.plan, S.cur, S.P, getD);
}
function simulate(patch, fsPatch) {
  const P2 = Object.assign({}, S.P, patch, { frozen: null });
  const rows = fsPatch ? filterRows(Object.assign({}, S.FS, fsPatch)) : S.F;
  const pl = planCalc(rows, P2);
  if (!pl || !pl.items.length) return null;
  return scen(pl, 1 - P2.cut / 100, P2, getD, S.DINTRA);
}
function activeFilters() {
  const f = S.FS, out = [];
  if (!f) return out;
  if (f.from !== S.SPAN[0] || f.to !== S.SPAN[1]) out.push(['date', 'Period', fmtDate(f.from) + ' – ' + fmtDate(f.to)]);
  if (f.ch.size !== S.CH.length) out.push(['ch', 'Channels', f.ch.size <= 3 ? [...f.ch].map(chName).join(', ') : f.ch.size + ' of ' + S.CH.length]);
  if (f.cat.size !== S.CATS.length) out.push(['cat', 'Categories', f.cat.size <= 2 ? [...f.cat].map(pn).join(', ') : f.cat.size + ' of ' + S.CATS.length]);
  if (f.day.size !== 7) out.push(['day', 'Days', f.day.size <= 3 ? [...f.day].map(d => d.slice(0, 3)).join(', ') : DAYS.filter(d => f.day.has(d)).map(d => d.slice(0, 2)).join(' ')]);
  if (f.h0 !== S.HRS[0] || f.h1 !== S.HRS[1]) out.push(['h', 'Time', f.h0 === f.h1 ? band(f.h0) : hl(f.h0) + ' – ' + hl((f.h1 + 1) % 24)]);
  if (f.minTvr > 0) out.push(['min', 'Min TVR', nf(f.minTvr, 1)]);
  if (f.q.trim()) out.push(['q', 'Search', '"' + f.q + '"']);
  if (f.p) out.push(['p', 'Program', pn(f.p)]);
  return out;
}
function clearFilter(k) {
  const f = S.FS, e = emptyFS();
  if (k === 'date') { f.from = e.from; f.to = e.to; } else if (k === 'h') { f.h0 = e.h0; f.h1 = e.h1; } else if (k === 'min') f.minTvr = 0;
  else if (k === 'all') { S.FS = e; } else f[k] = e[k];
  applyFilters(); if ($('#drawer').classList.contains('on')) renderDrawer();
}
function chrome() {
  const has = S.ROWS.length > 0;
  $('#periodTxt').textContent = has ? fmtDate(S.FS.from) + ' to ' + fmtDate(S.FS.to) : 'No data loaded';
  $('#periodSub').textContent = has ? dayDiff(S.FS.from, S.FS.to) + ' days · ' + ni(S.F.length) + ' airings' : '';
  const af = activeFilters();
  const b = $('#badge'); b.textContent = af.length; b.classList.toggle('show', af.length > 0);
  $('#dcount').textContent = has ? ni(S.F.length) + ' of ' + ni(S.ROWS.length) + ' airings' : '';
  $('#fbar').innerHTML = has ? (af.length ? `<span class="lbl">Filtered by</span>${af.map(a => `<span class="fchip">${a[1]}: <b>${esc(a[2])}</b><button data-clear="${a[0]}" aria-label="Remove ${a[1]} filter">×</button></span>`).join('')}<button class="link" data-clear="all">Clear all</button>` : '') +
    (S.meta && S.meta.demo ? `<span class="fchip" style="border-style:dashed">Demo data, not real ratings <button data-open="upload" aria-label="Upload your data" style="width:auto;border-radius:9px;padding:0 6px">Upload yours</button></span>` : '') : '';
}

/* ---------- charts ---------- */
function mkChart(id, cfg) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
  const el = document.getElementById(id);
  if (!el || typeof Chart === 'undefined') return;
  Chart.defaults.color = css('--muted'); Chart.defaults.borderColor = css('--line');
  Chart.defaults.font.family = '"IBM Plex Sans",system-ui,sans-serif'; Chart.defaults.font.size = 11.5;
  cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: false }, cfg.options || {});
  if (cfg.options.onClick) cfg.options.onHover = (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; };
  charts[id] = new Chart(el, cfg);
}
function killCharts(prefix) { Object.keys(charts).forEach(k => { if (!prefix || k.startsWith(prefix)) { charts[k].destroy(); delete charts[k]; } }); }
const emptyHTML = `<div class="empty"><b>No airings match these filters</b><p>Widen the date period, channels or time band.</p><button class="btn pri" data-clear="all">Clear all filters</button> <button class="btn" data-open="filters">Open filters</button></div>`;

/* ---------- welcome ---------- */
function welcomeHTML() {
  return `<div class="welcome">
    <div><h1>Plan TV smarter, straight from your ratings file</h1><p class="lead">Upload your program ratings export (Excel or CSV). Explore when and where people watch, drill into any channel or program, then build a channel split and program basket with net reach.</p></div>
    <div class="drop" id="wdrop"><b>Drop your Excel or CSV file here</b><p>Needed columns: Channel, Date, Start, Program, TVR. Also used: Day, End, Duration, Category, TVR Share %, Reach %.</p>
      <button class="btn pri" data-pick>Choose file</button> <button class="btn" data-demo>Try with demo data</button></div>
    <div class="priv"><svg class="pico" viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg><div><b>Your data stays in your browser.</b> The file is read on this device and saved only in this browser. Colleagues opening the same link see an empty dashboard until they upload their own file. Planner AI receives only a short summary of the filtered numbers, never the file.</div></div>
    <div class="feat">
      <div><b>1 · Explore</b><span>Day × hour heatmaps, best days, slot finder: "what airs at 10 PM on Hiru, Derana and Sirasa?"</span></div>
      <div><b>2 · Drill down</b><span>Click any bar, cell or row. Channel ▸ category ▸ program ▸ every airing.</span></div>
      <div><b>3 · Plan and test</b><span>Channel split, program basket, net reach after duplication, budget cut scenarios, export.</span></div>
    </div></div>`;
}

/* ---------- Overview ---------- */
function vOverview(el) {
  const F = S.F, sm = summarize(F);
  const chs = [...grp(F, r => r.ch).values()].sort((a, b) => avgT(b) - avgT(a));
  let pg = [...grp(F, r => r.ch + '||' + r.p).values()];
  let q = pg.filter(g => g.n >= 2); if (q.length < 10) q = pg;
  q.sort((a, b) => avgT(b) - avgT(a)); q = q.slice(0, 15);
  const cats = [...grp(F, r => r.cat).values()].sort((a, b) => b.sT - a.sT);
  const topCat = cats.slice(0, 7), rest = cats.slice(7).reduce((a, g) => a + g.sT, 0);
  const dps = [...grp(F, r => daypart(r.h)).values()].sort((a, b) => DPS.indexOf(a.k) - DPS.indexOf(b.k));
  const totG = sm.sT || 1;
  const ins = insights(F);
  el.innerHTML = `
  <div class="vhead"><div><h2>Market overview</h2><p>Click any bar, slice, row or card to drill down.</p></div></div>
  ${ins.length ? `<div class="ins">${ins.map(i => `<button class="insc ${i.cls} ${i.drill ? 'click' : ''}" ${i.drill ? xa(i.drill) : ''}><div class="big">${esc(i.big)}</div><div class="t">${esc(i.t)}</div></button>`).join('')}</div>` : ''}
  <div class="kpis">
    ${kp('Airings', ni(sm.n), ni(S.ROWS.length) + ' loaded in total')}
    ${kp('Programs', ni(sm.progs), sm.chs + ' channels')}
    ${kp('Average TVR', nf(sm.avgT, 2), 'per airing')}
    ${kp('Peak TVR', nf(sm.pk.tvr, 1), esc(pn(sm.pk.p)) + ', ' + esc(chName(sm.pk.ch)))}
    ${kp('Average reach', nf(sm.avgR, 2) + '%', 'per airing')}
    ${kp('Days covered', dayDiff(S.FS.from, S.FS.to), fmtDate(S.FS.from) + ' – ' + fmtDate(S.FS.to))}
  </div>
  <div class="grid g-ov">
    ${panel('Channel performance', 'average per airing · click a bar', '<div class="chart lg"><canvas id="c-ch"></canvas></div>')}
    ${panel('Top programs', 'by average TVR, min 2 airings · click a row', `<div class="tw" style="max-height:340px;overflow:auto"><table><thead><tr><th>Program</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>Airings</th><th>Steadiness</th></tr></thead><tbody>${q.map(g => { const [c, p] = g.k.split('||'); return `<tr class="click" ${xa({ ch: c, p })}><td><div class="pn" title="${esc(pn(p))}">${esc(pn(p))}</div><span class="sub">${dot(c)}${esc(chName(c))} · ${hl(g.mxr.h)}</span></td><td>${nf(avgT(g), 2)}</td><td>${nf(g.mx, 1)}</td><td>${nf(avgR(g), 1)}</td><td>${g.n}</td><td>${stdTag(ciOf(g), g.n)}</td></tr>`; }).join('')}</tbody></table></div>`)}
    ${panel('Category mix', 'share of all viewing (ratings) · click a slice', '<div class="chart lg"><canvas id="c-cat"></canvas></div>')}
  </div>
  <div class="grid g-wide">
    ${panel('Daily trend', 'average reach and rating per day · click a day', '<div class="chart"><canvas id="c-trend"></canvas></div>')}
    ${panel('Dayparts', 'where the viewing happens · click a row', `<table><thead><tr><th>Daypart</th><th>Airings</th><th>Avg TVR</th><th>Reach %</th><th>Share of viewing</th></tr></thead><tbody>${dps.map(g => `<tr class="click" ${xa({ dp: g.k })}><td>${esc(g.k)}</td><td>${ni(g.n)}</td><td>${nf(avgT(g), 2)}</td><td>${nf(avgR(g), 1)}</td><td><div class="mini-bar"><div class="tr"><div class="fl" style="width:${(g.sT / totG * 100).toFixed(0)}%"></div></div>${nf(g.sT / totG * 100, 0)}%</div></td></tr>`).join('')}</tbody></table>`)}
  </div>`;
  mkChart('c-ch', {
    type: 'bar', data: {
      labels: chs.map(g => chName(g.k)), datasets: [
        { label: 'Avg TVR', data: chs.map(g => +avgT(g).toFixed(2)), backgroundColor: css('--accent'), borderRadius: 3 },
        { label: 'Avg reach %', data: chs.map(g => +avgR(g).toFixed(2)), backgroundColor: css('--heat'), borderRadius: 3 }]
    },
    options: { indexAxis: 'y', onClick: (e, els) => els.length && openDetail({ ch: chs[els[0].index].k }), plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } } }, scales: { x: { grid: { color: css('--line') } }, y: { grid: { display: false } } } }
  });
  const labs = topCat.map(g => pn(g.k)), vals = topCat.map(g => g.sT);
  if (rest > 0) { labs.push('Other'); vals.push(rest); }
  mkChart('c-cat', {
    type: 'doughnut', data: { labels: labs, datasets: [{ data: vals, backgroundColor: [1, 2, 3, 4, 5, 6, 7, 9].map(i => css('--c' + i)), borderColor: css('--panel'), borderWidth: 2 }] },
    options: { cutout: '58%', onClick: (e, els) => { if (els.length && els[0].index < topCat.length) openDetail({ cat: topCat[els[0].index].k }); }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } }, tooltip: { callbacks: { label: c => ' ' + c.label + ': ' + nf(c.raw / totG * 100, 1) + '% of viewing' } } } }
  });
  const dates = [...grp(F, r => r.date).values()].sort((a, b) => a.k < b.k ? -1 : 1);
  mkChart('c-trend', {
    data: {
      labels: dates.map(g => fmtDate(g.k).slice(0, 6) + ' ' + g.mxr.day.slice(0, 2)), datasets: [
        { type: 'bar', label: 'Avg reach %', data: dates.map(g => +avgR(g).toFixed(2)), backgroundColor: dates.map(g => DAYS.indexOf(g.mxr.day) >= 5 ? css('--heat') : css('--accent')), borderRadius: 2, yAxisID: 'y' },
        { type: 'line', label: 'Avg TVR', data: dates.map(g => +avgT(g).toFixed(2)), borderColor: css('--ink'), backgroundColor: css('--ink'), pointRadius: 2, borderWidth: 1.5, tension: .25, yAxisID: 'y1' }]
    },
    options: { onClick: (e, els) => els.length && openDetail({ date: dates[els[0].index].k }), plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } } }, scales: { x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: true } }, y: { beginAtZero: true, title: { display: true, text: 'Avg reach %' } }, y1: { position: 'right', beginAtZero: true, grid: { display: false }, title: { display: true, text: 'Avg TVR' } } } }
  });
}

/* ---------- Explore ---------- */
function vExplore(el) {
  const F = S.F, M = S.EX;
  const metric = { tvr: r => r.tvr, rp: r => r.rp, sh: r => r.sh, n: () => 1 }[M.metric];
  const dimKey = { channel: 'ch', day: 'day', cat: 'cat' }[M.mode];
  let h0 = 24, h1 = -1; F.forEach(r => { if (r.h < h0) h0 = r.h; if (r.h > h1) h1 = r.h; });
  const cells = new Map(); let dims = new Set();
  F.forEach(r => { const d = DIMS[dimKey].key(r); dims.add(d); const k = d + '|' + r.h; let c = cells.get(k); if (!c) { c = [0, 0]; cells.set(k, c); } c[0] += metric(r); c[1]++; });
  if (dimKey === 'ch') dims = S.CH.filter(c => dims.has(c));
  else if (dimKey === 'day') dims = DAYS.filter(d => dims.has(d));
  else { const g = grp(F, r => r.cat); dims = [...dims].sort((a, b) => g.get(b).sT - g.get(a).sT).slice(0, 14); }
  const val = c => M.metric === 'n' ? c[1] : c[0] / c[1];
  let mx = 0, best = null; cells.forEach((c, k) => { const v = val(c); if (v > mx && (M.metric === 'n' || c[1] >= 2 || F.length < 200)) { mx = v; best = k; } });
  const hrs = []; for (let h = h0; h <= h1; h++) hrs.push(h);
  const mLab = { tvr: 'average TVR', rp: 'average reach %', sh: 'average TVR share %', n: 'number of airings' }[M.metric];
  let g = `<div class="tw"><div class="heat" style="grid-template-columns:130px repeat(${hrs.length},minmax(34px,1fr))"><div></div>${hrs.map(h => `<button class="hh hr" style="justify-content:center" ${xa({ h })} title="${band(h)}">${hl(h).replace(' ', '')}</button>`).join('')}`;
  dims.forEach(d => {
    const lab = dimKey === 'ch' ? dot(d) + esc(chName(d)) : esc(DIMS[dimKey].name(d));
    g += `<button class="hr" ${xa({ [dimKey]: d })} title="${esc(DIMS[dimKey].name(d))}">${lab}</button>`;
    hrs.forEach(h => {
      const c = cells.get(d + '|' + h);
      if (!c) { g += `<div class="hc" style="background:var(--soft)"></div>`; return; }
      const v = val(c), p = mx > 0 ? 8 + 84 * Math.pow(v / mx, .8) : 8;
      g += `<button class="hc ${best === d + '|' + h ? 'best' : ''}" ${xa({ [dimKey]: d, h })} title="${esc(DIMS[dimKey].name(d))}, ${band(h)}: ${nf(v, 2)} (${c[1]} airings). Click for programs." style="background:color-mix(in srgb,var(--heat) ${p.toFixed(0)}%,var(--panel));color:${p > 52 ? '#1b1206' : 'var(--ink)'}">${M.metric === 'n' ? v : v >= 10 ? nf(v, 0) : nf(v, 1)}</button>`;
    });
  });
  g += '</div></div>';
  const [bd, bh] = best ? best.split('|') : [];
  const days = [...grp(F, r => r.day).values()].sort((a, b) => DAYS.indexOf(a.k) - DAYS.indexOf(b.k));
  const bestDays = [...days].sort((a, b) => avgR(b) - avgR(a)).slice(0, 3);
  const hg = [...grp(F, r => r.h).values()].sort((a, b) => a.k - b.k);

  // Slot finder: all programs in a time band across chosen channels.
  const present = S.CH.filter(c => S.FS.ch.has(c));
  if (!S.SF.ch) S.SF.ch = new Set(present.slice(0, 3));
  const sfRows = F.filter(r => r.h === S.SF.h && S.SF.ch.has(r.ch));
  const sf = [...grp(sfRows, r => r.ch + '||' + r.p).values()].sort((a, b) => avgT(b) - avgT(a));
  const hopts = []; for (let h = 0; h < 24; h++) hopts.push(h);

  el.innerHTML = `
  <div class="vhead"><div><h2>Explore the market</h2><p>When and where your audience watches. Click a cell to see the programs in that slot.</p></div></div>
  ${panel('When do people watch?', `${mLab} by ${{ channel: 'channel', day: 'day', cat: 'category' }[M.mode]} and start hour${best ? ` · outlined = best slot (${esc(DIMS[dimKey].name(bd))}, ${hl(+bh)})` : ''}`, g + `<div class="legend">Lower ${[8, 25, 45, 65, 92].map(p => `<span class="sw" style="background:color-mix(in srgb,var(--heat) ${p}%,var(--panel))"></span>`).join('')} Higher</div>`,
    `${seg('mode', M, [['channel', 'By channel'], ['day', 'By day'], ['cat', 'By category']])}${seg('metric', M, [['tvr', 'TVR'], ['rp', 'Reach %'], ['sh', 'Share %'], ['n', 'Airings']])}`)}
  <div class="grid g3">
    ${panel('Day of week', 'bars: mean reach % · line: peak TVR · click a day', '<div class="chart"><canvas id="c-day"></canvas></div>' + `<div class="chips" style="margin-top:8px">${bestDays.map((d, i) => `<button class="tag" ${xa({ day: d.k })}>#${i + 1} ${d.k.slice(0, 3)} · ${nf(avgR(d), 2)}% reach</button>`).join('')}</div>`)}
    ${panel('Average TVR by hour', 'all selected channels · click an hour', '<div class="chart"><canvas id="c-hr"></canvas></div>')}
    ${panel('Dayparts by channel', 'share of each channel\'s viewing by daypart', '<div class="chart"><canvas id="c-dpc"></canvas></div>')}
  </div>
  ${panel('Slot finder', `every program starting ${band(S.SF.h)} on the chosen channels`, `
    <div class="inl" style="margin-bottom:10px"><label class="inl"><span class="muted">Hour</span><select id="sf-h">${hopts.map(h => `<option value="${h}" ${h === S.SF.h ? 'selected' : ''}>${band(h)}</option>`).join('')}</select></label>
    <div class="chips">${present.map(c => `<button class="chip ${S.SF.ch.has(c) ? 'on' : ''}" data-sfch="${esc(c)}">${esc(chName(c))}</button>`).join('')}</div></div>
    ${sf.length ? `<div class="tw"><table><thead><tr><th>Program</th><th class="l">Channel</th><th class="l">Category</th><th class="l">Days</th><th>Airings</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>Steadiness</th></tr></thead><tbody>${sf.map(g => { const [c, p] = g.k.split('||'); return `<tr class="click" ${xa({ ch: c, p })}><td>${esc(pn(p))}</td><td class="l">${dot(c)}${esc(chName(c))}</td><td class="l">${esc(pn(g.mxr.cat))}</td><td class="l">${DAYS.filter(d => g.dc[d]).map(d => d.slice(0, 2)).join(' ')}</td><td>${g.n}</td><td>${nf(avgT(g), 2)}</td><td>${nf(g.mx, 1)}</td><td>${nf(avgR(g), 1)}</td><td>${stdTag(ciOf(g), g.n)}</td></tr>`; }).join('')}</tbody></table></div>` : '<p class="muted">Nothing airs in this hour on the chosen channels.</p>'}`)}`;

  mkChart('c-day', {
    data: {
      labels: days.map(d => d.k.slice(0, 3)), datasets: [
        { type: 'bar', label: 'Mean reach %', data: days.map(d => +avgR(d).toFixed(2)), backgroundColor: days.map(d => bestDays.includes(d) ? css('--heat') : css('--accent')), borderRadius: 3 },
        { type: 'line', label: 'Peak TVR', data: days.map(d => +d.mx.toFixed(2)), borderColor: css('--ink'), backgroundColor: css('--ink'), pointRadius: 3, borderWidth: 1.5, tension: .25 }]
    },
    options: { onClick: (e, els) => els.length && openDetail({ day: days[els[0].index].k }), plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } } }, scales: { x: { grid: { display: false } } } }
  });
  mkChart('c-hr', {
    type: 'bar', data: { labels: hg.map(g => hl(g.k)), datasets: [{ label: 'Avg TVR', data: hg.map(g => +avgT(g).toFixed(2)), backgroundColor: hg.map(g => g.k >= 18 && g.k < 22 ? css('--accent') : css('--c9')), borderRadius: 3 }] },
    options: { onClick: (e, els) => els.length && openDetail({ h: hg[els[0].index].k }), plugins: { legend: { display: false } }, scales: { x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: true } } } }
  });
  const chs = S.CH.filter(c => F.some(r => r.ch === c));
  const byC = grp(F, r => r.ch);
  const dpc = grp(F, r => r.ch + '|' + daypart(r.h));
  mkChart('c-dpc', {
    type: 'bar', data: {
      labels: chs.map(chName), datasets: DPS.map((d, i) => ({ label: d, data: chs.map(c => { const g = dpc.get(c + '|' + d); return g ? +(g.sT / byC.get(c).sT * 100).toFixed(1) : 0; }), backgroundColor: css('--c' + [9, 6, 1, 3][i]) }))
    },
    options: { indexAxis: 'y', onClick: (e, els) => els.length && openDetail({ ch: chs[els[0].index], dp: DPS[els[0].datasetIndex] }), plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10.5 } } }, tooltip: { callbacks: { label: c => ' ' + c.dataset.label + ': ' + c.raw + '%' } } }, scales: { x: { stacked: true, max: 100 }, y: { stacked: true, grid: { display: false } } } }
  });
}

/* ---------- Drill down ---------- */
function vDrill(el) {
  const L = S.DR.levels, F = S.F, tot = F.reduce((a, r) => a + r.tvr, 0) || 1;
  const opts = (i, v) => `<select data-lvl="${i}">${i > 0 ? `<option value="">(none)</option>` : ''}${Object.keys(DIMS).map(k => `<option value="${k}" ${v === k ? 'selected' : ''} ${L.slice(0, i).includes(k) ? 'disabled' : ''}>${DIMS[k].label}</option>`).join('')}</select>`;
  let html = '';
  const render = (rows, depth, path, parentT, fixed) => {
    const dim = L[depth];
    if (!dim) { // leaf: individual airings
      const ar = [...rows].sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : a.s < b.s ? -1 : 1);
      ar.slice(0, 40).forEach(r => { html += `<tr class="air l${Math.min(depth, 3)}"><td>${r.date} ${r.day.slice(0, 3)} · ${r.s}${r.e ? '–' + r.e : ''} · ${esc(chName(r.ch))} · ${esc(pn(r.p))}</td><td>1</td><td>${nf(r.tvr, 2)}</td><td></td><td>${nf(r.rp, 2)}</td><td></td><td></td></tr>`; });
      if (ar.length > 40) html += `<tr class="air l${Math.min(depth, 3)}"><td colspan="7">… ${ar.length - 40} more airings. Open the detail view to see all.</td></tr>`;
      return;
    }
    const D = DIMS[dim];
    let gs = [...grp(rows, D.key).values()];
    gs.sort(D.order ? (a, b) => D.order(a.k, b.k) : (a, b) => b.sT - a.sT);
    const lim = depth === 0 ? 200 : 60;
    gs.slice(0, lim).forEach(g => {
      const p = path + '¦' + g.k, open = S.DR.open.has(p), f = Object.assign({}, fixed, { [dim]: g.k });
      const share = g.sT / (parentT || 1) * 100;
      html += `<tr class="l${Math.min(depth, 3)} ${depth === 0 ? 'l0' : ''}"><td><button class="tg" data-tg="${esc(p)}" aria-label="${open ? 'Collapse' : 'Expand'}">${open ? '▼' : '▶'}</button>${dim === 'ch' ? dot(g.k) : ''}<button class="nm" data-tg="${esc(p)}">${esc(D.name(g.k))}</button></td><td>${ni(g.n)}</td><td>${nf(avgT(g), 2)}</td><td>${nf(g.mx, 1)}</td><td>${nf(avgR(g), 2)}</td><td><div class="mini-bar"><div class="tr"><div class="fl" style="width:${Math.min(100, share).toFixed(0)}%;background:${dim === 'ch' ? cv(g.k) : 'var(--accent)'}"></div></div>${nf(share, 1)}%</div></td><td><button class="ib" ${xa(f)} title="Open detail">Detail ›</button></td></tr>`;
      if (open) render(g.k === undefined ? [] : rows.filter(r => D.key(r) === g.k), depth + 1, p, g.sT, f);
    });
    if (gs.length > lim) html += `<tr class="air"><td colspan="7">… ${gs.length - lim} more</td></tr>`;
  };
  render(F, 0, '', tot, {});
  el.innerHTML = `
  <div class="vhead"><div><h2>Drill down</h2><p>Expand any row to break it down further. The last level opens individual airings. Share = this row's share of the parent row's viewing (ratings).</p></div>
    <span class="push"><button class="btn sm" id="dr-exp">Expand level 1</button><button class="btn sm" id="dr-col">Collapse all</button></span></div>
  ${panel('Breakdown', '', `<div class="inl" style="margin-bottom:10px"><span class="muted">Level 1</span>${opts(0, L[0])}<span class="muted">▸ Level 2</span>${opts(1, L[1] || '')}<span class="muted">▸ Level 3</span>${opts(2, L[2] || '')}<span class="muted">▸ airings</span></div>
    <div class="tw"><table class="tree"><thead><tr><th>${L.map(k => DIMS[k].label).join(' ▸ ')}</th><th>Airings</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>Share of parent</th><th></th></tr></thead><tbody>${html}</tbody></table></div>`)}`;
}

/* ---------- Planner ---------- */
const TIER_NAMES = { 1: 'Peak impact', 2: 'Efficiency anchors', 3: 'Frequency builders' };
const TIER_COL = { 1: 'var(--c1)', 2: 'var(--c2)', 3: 'var(--c3)' };
const tierTag = t => `<span class="tag tier t${t}">T${t}</span>`;
function vPlan(el) {
  killCharts('c-');
  const P = S.P, f = S.FS;
  const chOn = S.CH.filter(c => f.ch.has(c));
  const num = (k, o = {}) => `<input type="number" data-p="${k}" value="${P[k]}" ${o.min != null ? `min="${o.min}"` : ''} ${o.max != null ? `max="${o.max}"` : ''} step="${o.step || 1}">`;
  const strat = STRATS[P.strategy] ? P.strategy : '';
  el.innerHTML = `
  <div class="vhead"><div><h2>Build your plan</h2><p>Spots are bought one at a time where they add the most new reach per rupee, within your tier split, caps and daypart limits.</p></div>
    <span class="push"><button class="btn" id="pl-save">Save as scenario</button><button class="btn" id="pl-csv">Export spot plan</button><button class="btn" data-deck>Export deck (PPTX)</button><button class="btn pri" data-ai="Write a planner memo for the client explaining this plan">Ask AI to explain</button></span></div>
  ${S.P.frozen ? `<div class="frozen"><div><b>Schedule is fixed</b> because rates were edited in the Schedule tab. Changing any setting below re-optimises the plan.</div><button class="btn pri" data-reopt>Re-optimise now</button></div>` : ''}
  <div class="planbar" id="pl-bar"></div>
  <div class="grid g-plan">
    <div class="col">
      <div class="panel"><div class="ph"><h3>Your brief</h3><span class="push"><button class="link" data-open="filters">Edit</button></span></div><div class="pb">
        <table class="brief"><tbody>
          <tr><td class="l muted">Channels</td><td>${f.ch.size === S.CH.length ? 'All ' + S.CH.length : f.ch.size + ' of ' + S.CH.length}</td></tr>
          <tr><td class="l muted">Categories</td><td>${f.cat.size === S.CATS.length ? 'All' : f.cat.size <= 2 ? [...f.cat].map(pn).join(', ') : f.cat.size + ' selected'}</td></tr>
          <tr><td class="l muted">Days</td><td>${f.day.size === 7 ? 'All week' : DAYS.filter(d => f.day.has(d)).map(d => d.slice(0, 3)).join(', ')}</td></tr>
          <tr><td class="l muted">Time band</td><td>${hl(f.h0)} – ${hl((f.h1 + 1) % 24)}</td></tr>
          <tr><td class="l muted">Ratings period</td><td>${fmtDate(f.from)} – ${fmtDate(f.to)}</td></tr>
        </tbody></table>
        <div class="presets"><button class="btn sm" data-preset="prime">Prime 6–10 PM</button><button class="btn sm" data-preset="allday">All day</button><button class="btn sm" data-preset="wd">Weekdays</button><button class="btn sm" data-preset="we">Weekend</button></div>
      </div></div>
      <div class="panel settings"><div class="ph"><h3>Plan settings</h3><span class="push"><button class="link" id="p-reset">Reset</button></span></div><div class="pb">
        <details open><summary>Budget and cost</summary>
          <label class="fld"><span>Budget (LKR)</span>${num('budget', { min: 0, step: 500000 })}</label>
          <div class="two"><label class="fld"><span>Cost per rating point (30 sec)</span>${num('cprp', { min: 1000, step: 1000 })}</label>
          <label class="fld"><span>Minimum 30-sec rate</span>${num('minRate', { min: 0, step: 1000 })}</label></div>
          <p class="hint">Estimated 30-sec rate = the higher of the two × TVR. Type real rate cards and channel discounts in the <button class="link" data-tab-go="schedule">Schedule</button> tab.</p>
          <label class="fld"><span>Budget change <b id="l-cut">${P.cut ? '−' + P.cut + '%' : 'none'}</b></span><input type="range" id="p-cut" min="0" max="50" step="5" value="${P.cut}"></label>
          <label class="fld"><span>Target net reach <b id="l-tgt">${P.target ? P.target + '%' : 'none'}</b></span><input type="range" id="p-tgt" min="0" max="90" step="1" value="${P.target}"></label>
        </details>
        <details open><summary>Creatives</summary>
          <div class="crs">${P.creatives.map((c, i) => `<div class="cr-row"><input type="text" data-cr="${i}" data-f="name" value="${esc(c.name)}" placeholder="Brand / creative" aria-label="Creative name"><select data-cr="${i}" data-f="dur" aria-label="Duration">${[5, 10, 15, 20, 25, 30, 45, 60].map(d => `<option value="${d}" ${+c.dur === d ? 'selected' : ''}>${d}s</option>`).join('')}</select><input type="number" min="0" max="100" data-cr="${i}" data-f="share" value="${c.share}" aria-label="Rotation share %"><span class="muted">%</span>${P.creatives.length > 1 ? `<button class="ico x" data-cr-del="${i}" title="Remove creative" aria-label="Remove creative">${ICO_X}</button>` : '<span></span>'}</div>`).join('')}</div>
          <div class="inl" style="justify-content:space-between;margin-top:6px"><button class="btn sm" data-cr-add>+ Add creative</button><span class="muted" id="l-crsum">Shares ${P.creatives.reduce((a, c) => a + (+c.share || 0), 0)}% · avg ${nf(avgLen(P), 1)} sec</span></div>
          <p class="hint">Cost of a spot = 30-sec rate × duration ÷ 30. Spots are rotated between creatives by share.</p>
        </details>
        <details open><summary>Strategy and tiers</summary>
          <div class="fld">${seg('strategy', { strategy: strat }, Object.entries(STRATS).map(([k, v]) => [k, v.label]))}${strat ? '' : ' <span class="tag">Custom</span>'}</div>
          <div class="fld"><span>Tier budget split (%) <b id="l-tsum">${P.tiers.reduce((a, b) => a + (+b || 0), 0)}%</b></span>
            <div class="three">${[0, 1, 2].map(i => `<label><span class="sub">Tier ${i + 1}</span><input type="number" min="0" max="100" data-pa="tiers" data-i="${i}" value="${P.tiers[i]}"></label>`).join('')}</div></div>
          <div class="two"><label class="fld"><span>Tier 1 from percentile</span>${num('tp1', { min: 50, max: 99 })}</label><label class="fld"><span>Tier 3 below percentile</span>${num('tp3', { min: 1, max: 50 })}</label></div>
          <label class="fld"><span>Ignore programs below TVR</span>${num('minTvrPlan', { min: 0, step: .1 })}</label>
          <p class="hint" id="l-thr"></p>
        </details>
        <details><summary>Channels and programs</summary>
          <div class="fld">${seg('chMode', P, [['auto', 'Auto (best N)'], ['manual', 'Pick channels']])}</div>
          ${P.chMode === 'auto' ? `<label class="fld"><span>Channels in plan <b id="l-nch">${P.nCh}</b></span><input type="range" id="p-nch" min="1" max="${Math.max(1, chOn.length)}" value="${Math.min(P.nCh, Math.max(1, chOn.length))}"></label>`
      : `<div class="fld"><div class="chips">${chOn.map(c => `<button class="chip ${P.chPick.includes(c) ? 'on' : ''}" data-pick-ch="${esc(c)}">${esc(chName(c))}</button>`).join('')}</div></div>`}
          <label class="fld"><span>Programs per channel per tier <b id="l-np">${P.nProg}</b></span><input type="range" id="p-np" min="1" max="5" value="${P.nProg}"></label>
        </details>
        <details><summary>Caps and dayparts</summary>
          <label class="fld"><span>Max spots per program per week <b id="l-cap">${P.capWk}</b></span><input type="range" id="p-cap" min="1" max="10" value="${P.capWk}"></label>
          <table class="dpt"><thead><tr><th>Daypart</th><th>Min %</th><th>Max %</th></tr></thead><tbody>${DPS.map((d, i) => `<tr><td>${esc(d)}</td><td><input type="number" min="0" max="100" data-pa="dpMin" data-i="${i}" value="${P.dpMin[i]}"></td><td><input type="number" min="0" max="100" data-pa="dpMax" data-i="${i}" value="${P.dpMax[i]}"></td></tr>`).join('')}</tbody></table>
          <p class="hint">Share of spend by the program's usual start hour. Limits for dayparts outside the brief are ignored.</p>
        </details>
        <details><summary>Flighting</summary>
          <div class="two"><label class="fld"><span>Campaign start</span><input type="date" data-p="start" value="${P.start}"></label><label class="fld"><span>Campaign end</span><input type="date" data-p="end" value="${P.end}" min="${P.start}"></label></div>
          <p class="hint" style="margin-top:-6px">${campaignDays(P).days.length} days · ${campaignDays(P).W} weeks</p>
          <label class="fld"><span>Max spots per programme per day</span><select data-p="perDay">${[1, 2, 3].map(n => `<option value="${n}" ${+P.perDay === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
          <div class="fld"><span>Pacing</span>${seg('pacing', P, [['even', 'Even (drip)'], ['burst', 'Burst'], ['pulse', 'Pulse']])}</div>
          <div class="fld"><span>Same-hour spots on rival channels</span>${seg('same', P, [['roadblock', 'Roadblock'], ['stagger', 'Stagger']])}</div>
          <p class="hint">${P.same === 'roadblock' ? 'Roadblock: rival channels at the same hour on the same night. One viewer cannot watch both, so this reaches more different people.' : 'Stagger: rival same-hour spots on different nights. The same viewers see the ad more often (frequency).'}</p>
        </details>
      </div></div>
    </div>
    <div class="col">
      <div class="grid g2">
        <div class="panel"><div class="ph"><h3>Estimated reach</h3></div><div class="pb" id="pl-net"></div></div>
        <div class="panel"><div class="ph"><h3>Plan health check</h3></div><div class="pb"><ul class="health" id="pl-health"></ul></div></div>
      </div>
      <div class="panel"><div class="ph"><h3>Tier pyramid</h3><span class="s">target vs actual share of spend</span></div><div class="pb" id="pl-tier"></div></div>
      <div class="panel"><div class="ph"><h3>Channel split</h3><span class="s">budget per channel, filled with spots</span><span class="push"><button class="btn sm" id="pl-adj">${S.editSplit ? 'Done' : 'Fix channel shares'}</button></span></div><div class="pb" id="pl-split"></div></div>
      <div class="panel"><div class="ph"><h3>Programme basket</h3><span class="s" id="bk-sum"></span></div><div class="pb" id="pl-basket"></div></div>
      <div class="panel"><div class="ph"><h3>Weekly flighting</h3><span class="s" id="fl-sub"></span><span class="push">${seg('fv', S, [['weeks', 'By week'], ['days', 'Day plan']])}<button class="btn sm" id="fl-csv">Export schedule</button></span></div><div class="pb" id="pl-flight"></div></div>
    </div>
    <div class="col">
      <div class="panel"><div class="ph"><h3>Planner memo</h3><span class="s">auto-written from the numbers</span><span class="push"><button class="link" id="pl-copy">Copy</button></span></div><div class="pb memo" id="pl-memo"></div></div>
      <div class="panel"><div class="ph"><h3>Reach build</h3><span class="s">net reach as channels are added</span></div><div class="pb"><div class="chart sm"><canvas id="c-build"></canvas></div></div></div>
      <div class="panel"><div class="ph"><h3>Competing slots</h3><span class="s">${P.same === 'roadblock' ? 'roadblocked on the same nights' : 'staggered across nights'}</span></div><div class="pb" id="pl-comp" style="max-height:300px;overflow:auto"></div></div>
    </div>
  </div>
  <div class="nextbar"><span>Happy with the basket? Save it and stress-test it against budget changes.</span><button class="btn pri" data-tab-go="scen">Next: Scenarios →</button></div>`;
  updatePlan();
}
function updatePlan() {
  if (!$('#pl-net')) return;
  const A = S.cur, base = S.base, P = S.P;
  if (S.plan && S.plan.thr) $('#l-thr').textContent = `In this brief: Tier 1 averages ≥ ${nf(S.plan.thr.t1, 2)} TVR, Tier 3 < ${nf(S.plan.thr.t3, 2)} TVR. ${S.plan.items.length} candidate programs.`;
  if (!A || !A.kept.length) {
    $('#pl-bar').innerHTML = ''; $('#pl-net').innerHTML = '<p class="muted">No spots could be bought. Add channels, lower the minimum TVR, or un-exclude programs.</p>';
    ['#pl-health', '#pl-split', '#pl-basket', '#pl-memo', '#pl-comp', '#pl-tier', '#pl-flight'].forEach(s => $(s).innerHTML = ''); return;
  }
  const warn = S.health.filter(h => h.t === 'wa').length;
  const m = (l, v, cls = '', t = '') => `<div class="m ${cls}" title="${t}"><small>${l}</small><b>${v}</b></div>`;
  $('#pl-bar').innerHTML = m('Est. net reach', nf(A.net, 1) + '%', 'hl', 'Reached at least once') + m('Reach 3+', nf(A.r3, 1) + '%', '', 'Reached at least 3 times') + m('Budget used', lkr(A.spent)) +
    m('Channels', A.chs.length) + m('Programs', A.kept.length) + m('Spots', ni(A.spots)) + m('Avg frequency', nf(A.freq, 1) + 'x') + m('Health', warn ? warn + (warn > 1 ? ' warnings' : ' warning') : 'OK');
  const full = P.cut > 0 ? ` <span class="muted" style="font-size:12px">(full budget ${nf(base.net, 1)}%)</span>` : '';
  $('#pl-net').innerHTML = `<div class="bigline"><b>${nf(A.net, 1)}%</b><span class="muted">reached at least once${full}</span></div>
    <div class="flow"><div><b>${nf(A.gross, 1)}%</b><small>Gross reach</small></div><i>−</i><div><b>${nf(A.gross - A.net, 1)} pts</b><small>Same viewers</small></div><i>=</i><div><b style="color:var(--accent)">${nf(A.net, 1)}%</b><small>Net reach</small></div></div>
    <div class="flow" style="grid-template-columns:1fr 1fr 1fr"><div><b>${nf(A.r3, 1)}%</b><small>Reach 3+</small></div><div><b>${nf(A.freq, 1)}x</b><small>Avg frequency</small></div><div><b>${nf(A.eff, 2)}</b><small>Net pts per LKR 1M</small></div></div>
    <p class="hint">Reach 3+ = reached at least three times, the usual effective-frequency goal. Duplication and repeat-spot reach are planning assumptions (Duplication tab).${P.target ? ` Target ${P.target}%: ${A.net >= P.target ? '<b style="color:var(--good)">met</b>' : '<b style="color:var(--bad)">not met</b>'}.` : ''}</p>`;
  $('#pl-health').innerHTML = S.health.map(h => `<li><span class="ic ${h.t}">${ICON[h.t]}</span><span>${h.m}</span></li>`).join('');
  // Tier pyramid
  const thr = A.thr || { t1: 0, t3: 0 };
  const band_ = t => t === 1 ? `avg TVR ≥ ${nf(thr.t1, 2)}` : t === 3 ? `avg TVR < ${nf(thr.t3, 2)}` : `${nf(thr.t3, 2)} – ${nf(thr.t1, 2)}`;
  $('#pl-tier').innerHTML = `<div class="tiers">${A.tiers.map(t => `<div class="trow"><div><b>Tier ${t.t}</b> · ${TIER_NAMES[t.t]}<span class="sub">${band_(t.t)} · ${t.n} of ${t.avail} programs bought</span></div>
      <div class="tbars"><div class="tb"><i style="width:${Math.min(100, t.target)}%;background:color-mix(in srgb,${TIER_COL[t.t]} 30%,var(--panel))"></i><span>target ${nf(t.target, 0)}%</span></div><div class="tb"><i style="width:${Math.min(100, t.actual)}%;background:${TIER_COL[t.t]}"></i><span><b>actual ${nf(t.actual, 0)}%</b></span></div></div></div>`).join('')}</div>
    <div class="tw" style="margin-top:10px"><table><thead><tr><th>Daypart</th><th>Limit</th><th>Actual</th><th></th></tr></thead><tbody>${A.dps.filter(d => d.present).map(d => { const bad = d.actual > d.max + .5 || d.actual + .5 < d.min; return `<tr><td>${esc(d.d)}</td><td>${nf(d.min, 0)}–${nf(d.max, 0)}%</td><td><b>${nf(d.actual, 0)}%</b></td><td style="width:40%"><div class="mini-bar" style="justify-content:flex-start"><div class="tr" style="width:100%"><div class="fl" style="width:${Math.min(100, d.actual)}%;background:${bad ? 'var(--bad)' : 'var(--accent)'}"></div></div></div></td></tr>`; }).join('')}</tbody></table></div>`;
  // Split
  $('#pl-split').innerHTML = `<div class="stack">${A.chs.map(c => `<div style="width:${(c.allocW * 100).toFixed(2)}%;background:${cv(c.ch)}" title="${esc(chName(c.ch))} ${nf(c.allocW * 100, 1)}%">${c.allocW > .08 ? esc(chName(c.ch).replace(' TV', '')) + ' ' + nf(c.allocW * 100, 0) + '%' : ''}</div>`).join('')}</div>
   <div class="tw"><table><thead><tr><th>Channel</th><th>Share</th><th>Budget</th><th>Booked</th><th>Left</th><th>Programs</th><th>Spots</th><th>Reach</th><th class="l">Why this weight</th></tr></thead><tbody>${A.chs.map(c => {
    const sel = S.plan.sel.find(x => x.ch === c.ch);
    const rev = c.progs.some(x => x.role === 'review') && c.progs.length === 1;
    const why = sel && sel.manual ? 'Fixed share (yours)' : `Channel score ${nf((sel?.score || 0) * 100, 1)}: 0.6 reach share + 0.4 TVR share`;
    const left = A.fixed ? '' : c.gapWhy === 'cap' ? 'every programme is full' : c.gapWhy === 'daypart' ? 'daypart limit reached' : c.gapWhy === 'budget' ? 'total budget used' : c.gapWhy === 'spot' ? 'less than one spot' : '';
    return `<tr class="${rev || c.gapWhy === 'cap' ? 'warn' : ''}"><td><button class="nm ib" style="padding:0;color:var(--ink)" ${xa({ ch: c.ch })}>${dot(c.ch)}${esc(chName(c.ch))}</button></td><td>${S.editSplit ? `<input class="split-in" type="number" min="0" max="100" step="1" data-split="${esc(c.ch)}" value="${(c.allocW * 100).toFixed(0)}">` : `<b>${nf(c.allocW * 100, 1)}%</b>`}</td><td>${lkr(c.alloc).replace('LKR ', '')}</td><td><b>${lkr(c.bud).replace('LKR ', '')}</b></td><td>${lkr(c.gap).replace('LKR ', '')}${left ? `<span class="sub">${left}</span>` : ''}</td><td>${c.progs.length}</td><td>${ni(c.spots)}</td><td>${nf(c.R, 1)}%</td><td class="l">${esc(why)}</td></tr>`;
  }).join('')}</tbody></table></div><p class="hint">Each channel gets its share of the budget, and spots are added until the next spot would no longer fit. Programmes already in the basket get more spots first; if money is still left, the channel's next best programmes are added (marked Top-up in the basket).${S.editSplit ? ' Type a share to fix a channel; the other channels share the rest by score. <button class="link" id="pl-split-reset">Clear fixed shares</button>' : ''}</p>`;
  // Basket
  renderBasket();
  renderFlight();
  $('#pl-memo').innerHTML = memo(A, base);
  const cp = competing(S.F, A.chs.map(c => c.ch), 8);
  $('#pl-comp').innerHTML = cp.length ? cp.map(o => `<div class="flag"><b>${band(o.h)}</b>${o.l.slice(0, 4).map(x => `<span class="sub">${dot(x.ch)}${esc(chName(x.ch))}: ${esc(pn(x.p))}, TVR ${nf(x.m, 1)}</span>`).join('')}</div>`).join('') : '<p class="muted">No hour has two plan channels airing programs.</p>';
  const n = Math.min(S.plan.cs.length, 9), pts = [];
  for (let k = 1; k <= n; k++) { const r = simulate({ chMode: 'auto', nCh: k, split: {} }); pts.push(r ? +r.net.toFixed(1) : 0); }
  mkChart('c-build', {
    type: 'line', data: { labels: pts.map((_, i) => (i + 1) + ' ch'), datasets: [{ label: 'Net reach %', data: pts, borderColor: css('--accent'), backgroundColor: css('--accent'), pointRadius: pts.map((_, i) => i + 1 === A.chs.length ? 6 : 3), pointBackgroundColor: pts.map((_, i) => i + 1 === A.chs.length ? css('--heat') : css('--accent')), tension: .3 }] },
    options: { plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.raw}% net reach with top ${c.dataIndex + 1} channel${c.dataIndex ? 's' : ''}` } } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true } } }
  });
}
// Program basket: compact, grouped by channel (or a flat list by spend), with tier filters.
const ICO_LOCK = '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>';
const ICO_X = '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>';
function renderBasket() {
  const el = $('#pl-basket'), A = S.cur, B = S.BK;
  if (!el || !A) return;
  const all = A.kept, mxk = Math.max(...all.map(x => x.k), .01);
  const watch = all.filter(x => x.role === 'review');
  const pass = x => B.tier === 'all' || (B.tier === 'watch' ? x.role === 'review' : x.tier === +B.tier);
  const items = all.filter(pass);
  const cnt = t => all.filter(x => x.tier === t).length;
  const chip = (v, l, n) => `<button class="chip ${B.tier === String(v) ? 'on' : ''}" data-bkt="${v}">${l} <b>${n}</b></button>`;
  const row = x => {
    const sd = steadiness(x.ci, x.n);
    const tags = (x.role === 'anchor' ? '<span class="tag anchor">Anchor</span>' : '') + (x.role === 'review' ? `<span class="tag review" title="Ratings swing week to week (steadiness ${nf(x.ci, 1)})">Volatile</span>` : '') + (x.locked ? '<span class="tag">Locked</span>' : '') + (x.topup ? '<span class="tag" title="Added to use the rest of this channel\'s budget">Top-up</span>' : '') + (x.atCap ? `<span class="tag" title="At the ${S.P.capWk}/week spot cap">At cap</span>` : '');
    return `<div class="bk-row">
      <div class="bk-name"><button class="nm" ${xa({ ch: x.ch, p: x.p })} title="${esc(pn(x.p))}: open detail">${esc(pn(x.p))}</button>${tags}
        <span class="sub">${B.view === 'list' ? dot(x.ch) + esc(chName(x.ch)) + ' · ' : ''}${hl(x.hour)} · ${esc(pn(x.cat))} · ${sd.label.toLowerCase()}</span></div>
      <div>${tierTag(x.tier)}</div>
      <div class="r"><b>${nf(x.mean, 1)}</b><span class="sub">TVR</span></div>
      <div class="r"><b>${nf(x.rp, 0)}%</b><span class="sub">reach</span></div>
      <div class="r"><b>${x.spots}</b><span class="sub">spots</span></div>
      <div class="bk-share"><div class="tr"><div class="fl" style="width:${(x.k / mxk * 100).toFixed(0)}%"></div></div><span><b>${nf(x.k * 100, 1)}%</b> · ${lkr(x.bud).replace('LKR ', '')}</span></div>
      <div class="bk-act"><button class="ico ${x.locked ? 'on' : ''}" data-lock="${esc(x.key)}" title="${x.locked ? 'Unlock' : 'Lock: always buy this programme'}" aria-label="${x.locked ? 'Unlock' : 'Lock'} ${esc(pn(x.p))}">${ICO_LOCK}</button><button class="ico x" data-excl="${esc(x.key)}" title="Remove from plan" aria-label="Remove ${esc(pn(x.p))}">${ICO_X}</button></div>
    </div>`;
  };
  let body = '';
  if (!items.length) body = '<p class="muted" style="padding:10px 0">No programmes in this group.</p>';
  else if (B.view === 'list') {
    const lim = B.all ? items.length : Math.min(10, items.length);
    body = items.slice(0, lim).map(row).join('') + (items.length > lim ? `<button class="bk-more" data-bkall>Show all ${items.length} programmes</button>` : items.length > 10 ? '<button class="bk-more" data-bkall>Show top 10 only</button>' : '');
  } else {
    body = A.chs.map(c => {
      const its = items.filter(x => x.ch === c.ch); if (!its.length) return '';
      const open = B.open.has(c.ch), lim = open ? its.length : Math.min(2, its.length);
      return `<div class="bk-grp"><div class="bk-gh">${dot(c.ch)}<b>${esc(chName(c.ch))}</b><span class="muted">${nf(c.w * 100, 0)}% of spend · ${lkr(c.bud).replace('LKR ', '')} · ${c.spots} spots · ${c.progs.length} programme${c.progs.length > 1 ? 's' : ''}</span></div>
        ${its.slice(0, lim).map(row).join('')}
        ${its.length > 2 ? `<button class="bk-more" data-bkopen="${esc(c.ch)}">${open ? 'Show less' : `+ ${its.length - lim} more on ${esc(chName(c.ch))}`}</button>` : ''}</div>`;
    }).join('');
  }
  if ($('#bk-sum')) $('#bk-sum').textContent = `${all.length} programmes · ${ni(A.spots)} spots · click a name for detail`;
  const dropped = A.dropped || [], ex = S.P.excl;
  el.innerHTML = `
    <div class="bk-top">
      <div class="chips">${chip('all', 'All', all.length)}${chip(1, 'Tier 1', cnt(1))}${chip(2, 'Tier 2', cnt(2))}${chip(3, 'Tier 3', cnt(3))}${watch.length ? chip('watch', 'Volatile', watch.length) : ''}</div>
      ${seg('view', B, [['channel', 'By channel'], ['list', 'By spend']])}
    </div>
    <div class="bk">${body}</div>
    ${dropped.length ? `<p class="hint">Dropped by the ${S.P.cut}% budget cut: ${dropped.map(x => esc(pn(x.p))).join(', ')}.</p>` : ''}
    ${ex.length ? `<p class="hint">Removed by you: ${ex.map(k => `<span class="tag">${esc(pn(k.split('||')[1]))} <button class="link" data-unexcl="${esc(k)}">restore</button></span>`).join(' ')}</p>` : ''}`;
}
function renderFlight() {
  const A = S.cur, Sc = S.sched, P = S.P, el = $('#pl-flight');
  if (!el || !Sc) return;
  const pl = { even: 'even (drip)', burst: 'burst, front-loaded', pulse: 'pulse, on/off weeks' }[P.pacing];
  $('#fl-sub').textContent = `${fmtDate(Sc.start)} – ${fmtDate(Sc.end)} (${A.W} weeks) · ${pl} · max ${P.capWk}/week, ${P.perDay}/day per programme`;
  const wkDate = w => { const d = new Date(new Date(Sc.start + 'T12:00:00').getTime() + w * 7 * 864e5); return fmtDate(d.toISOString().slice(0, 10)).slice(0, 6); };
  if (S.fv === 'days') {
    const rows = Sc.rows;
    el.innerHTML = `<div class="tw" style="max-height:420px;overflow:auto"><table><thead><tr><th>Date</th><th class="l">Day</th><th class="l">Time band</th><th class="l">Channel</th><th class="l">Program</th><th>Tier</th><th class="l">Creative</th><th>Spots</th><th>Cost</th></tr></thead><tbody>${rows.slice(0, 400).map(r => `<tr><td>${fmtDate(r.date)}</td><td class="l">${r.day.slice(0, 3)}</td><td class="l">${band(r.hour)}</td><td class="l">${dot(r.ch)}${esc(chName(r.ch))}</td><td class="l">${esc(pn(r.p))}</td><td>${tierTag(r.tier)}</td><td class="l">${esc(Sc.mix[r.cr].name)} · ${r.dur}s</td><td>${r.spots}</td><td>${lkr(r.cost).replace('LKR ', '')}</td></tr>`).join('')}</tbody></table></div>
      <p class="hint">${Sc.clashes} night${Sc.clashes === 1 ? ' has' : 's have'} rival channels in the same hour (${P.same === 'roadblock' ? 'roadblock, by design' : 'kept low by staggering'}). Days follow each program's air days in the data.</p>`;
    return;
  }
  el.innerHTML = `<div class="chart sm"><canvas id="c-flight"></canvas></div>
    <div class="tw" style="max-height:360px;overflow:auto;margin-top:10px"><table><thead><tr><th>Program</th><th>Tier</th>${Sc.weeks.map((_, i) => `<th>W${i + 1}<span class="sub">${wkDate(i)}</span></th>`).join('')}<th>Total</th></tr></thead><tbody>
    ${Sc.grid.slice(0, S.BK.flAll ? Sc.grid.length : 8).map(g => `<tr><td><div class="pn">${esc(pn(g.x.p))}</div><span class="sub">${dot(g.x.ch)}${esc(chName(g.x.ch))} · ${hl(g.x.hour)}</span></td><td>${tierTag(g.x.tier)}</td>${g.alloc.map(n => `<td style="${n ? '' : 'color:var(--muted)'}">${n || '·'}</td>`).join('')}<td><b>${g.x.spots}</b></td></tr>`).join('')}
    ${Sc.grid.length > 8 ? `<tr><td colspan="${Sc.weeks.length + 3}"><button class="link" data-flall>${S.BK.flAll ? 'Show top 8 only' : `Show all ${Sc.grid.length} programmes`}</button></td></tr>` : ''}
    <tr><td><b>Spots</b></td><td></td>${Sc.weeks.map(w => `<td><b>${w.spots}</b></td>`).join('')}<td><b>${A.spots}</b></td></tr>
    <tr><td><b>Budget</b></td><td></td>${Sc.weeks.map(w => `<td>${lkr(w.bud).replace('LKR ', '')}</td>`).join('')}<td><b>${lkr(A.spent).replace('LKR ', '')}</b></td></tr></tbody></table></div>`;
  const chs = A.chs.map(c => c.ch);
  mkChart('c-flight', {
    type: 'bar', data: { labels: Sc.weeks.map((_, i) => 'W' + (i + 1) + ' ' + wkDate(i)), datasets: chs.map(ch => ({ label: chName(ch), data: Sc.weeks.map(w => w.byCh[ch] || 0), backgroundColor: cc(ch), borderRadius: 2 })) },
    options: { plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.raw} spots` } } }, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, title: { display: true, text: 'Spots' } } } }
  });
}
function scheduleCSV() {
  const r = S.sched.rows;
  return toCSV(['Week', 'Date', 'Day', 'Time band', 'Channel', 'Program', 'Tier', 'Creative', 'Duration (sec)', 'Spots', 'Cost LKR'],
    r.map(x => [x.week, x.date, x.day, band(x.hour), chName(x.ch), pn(x.p), 'Tier ' + x.tier, S.sched.mix[x.cr].name, x.dur, x.spots, Math.round(x.cost)]));
}
function memo(A, base) {
  const L = [], P = S.P;
  L.push(`<p><b>Recommended split:</b> ${A.chs.map(c => `${esc(chName(c.ch))} ${nf(c.w * 100, 0)}%`).join(', ')} of ${lkr(A.spent)}.</p>`);
  const sl = STRATS[P.strategy] ? STRATS[P.strategy].label : 'Custom';
  L.push(`<p><b>Strategy:</b> ${sl} pyramid. ${A.tiers.map(t => `Tier ${t.t} ${nf(t.actual, 0)}% (target ${nf(t.target, 0)}%)`).join(', ')}. Tier 1 programs average ${nf(A.thr.t1, 1)}+ TVR.</p>`);
  const anc = A.kept.filter(x => x.role === 'anchor').slice(0, 3);
  if (anc.length) L.push(`<p><b>Anchors:</b> ${anc.map(x => `${esc(pn(x.p))} on ${esc(chName(x.ch))} (TVR ${nf(x.mean, 1)}, ${x.spots} spots)`).join('; ')}.</p>`);
  if (A.chs.length >= 2) {
    const a = A.chs[0], b = A.chs[1], d = getD(a.ch, b.ch);
    L.push(`<p><b>Duplication:</b> ${esc(chName(a.ch))} and ${esc(chName(b.ch))} share about ${nf(d * 100, 0)}% of viewers in this model; ${nf(A.loss * 100, 0)}% of gross reach is overlap.</p>`);
  }
  L.push(`<p><b>Delivery:</b> ${ni(A.spots)} spots (${esc(mixText(P))}), net reach ${nf(A.net, 1)}%, reach 3+ ${nf(A.r3, 1)}%, average frequency ${nf(A.freq, 1)}x.</p>`);
  if (S.sched) L.push(`<p><b>Flighting:</b> ${A.W} weeks from ${fmtDate(S.sched.start)}, ${P.pacing} pacing (${S.sched.weeks.map(w => w.spots).join(' / ')} spots per week), max ${P.capWk} spots per program per week. Rival same-hour spots are ${P.same === 'roadblock' ? 'roadblocked on the same nights for reach' : 'staggered across nights for frequency'}.</p>`);
  if (P.cut > 0) L.push(`<p><b>Budget cut of ${P.cut}%:</b> net reach ${nf(base.net, 1)}% → ${nf(A.net, 1)}%, reach 3+ ${nf(base.r3, 1)}% → ${nf(A.r3, 1)}%. The cut removes repeat spots first, so frequency falls faster than reach.</p>`);
  const rv = A.kept.filter(x => x.role === 'review');
  if (rv.length) L.push(`<p><b>Watch:</b> ${rv.map(x => esc(pn(x.p))).join(', ')} ${rv.length > 1 ? 'have' : 'has'} volatile ratings.</p>`);
  return L.join('');
}
const memoText = () => $('#pl-memo') ? $('#pl-memo').innerText : '';

/* ---------- Schedule (booking sheet per channel) ---------- */
const MONS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function scheduleModel() {
  const A = S.cur, Sc = S.sched, P = S.P;
  const sscl = (+P.sscl || 0) / 100, vat = (+P.vat || 0) / 100;
  const chans = A.chs.map(c => {
    const progs = A.kept.filter(x => x.ch === c.ch).sort((a, b) => (a.from || '') < (b.from || '') ? -1 : 1);
    const sections = Sc.mix.map((m, i) => ({ m, i, rows: progs.map(x => {
      const rs = Sc.rows.filter(r => r.key === x.key && r.cr === i);
      const byDate = {}; rs.forEach(r => byDate[r.date] = (byDate[r.date] || 0) + r.spots);
      const spots = rs.reduce((a, r) => a + r.spots, 0), rc = x.rate30 * m.dur / 30, ng = x.net30 * m.dur / 30;
      return { x, byDate, spots, rc, ng, trc: rc * spots, tng: ng * spots };
    }).filter(r => r.spots > 0) })).filter(sec => sec.rows.length);
    const all = sections.flatMap(sec => sec.rows);
    const net = all.reduce((a, r) => a + r.tng, 0), rcT = all.reduce((a, r) => a + r.trc, 0), spots = all.reduce((a, r) => a + r.spots, 0);
    const tS = net * sscl, tV = (net + tS) * vat;
    return { ch: c.ch, sections, net, rcT, spots, sscl: tS, vat: tV, total: net + tS + tV, disc: +(P.disc[c.ch] || 0) };
  });
  const T = k => chans.reduce((a, c) => a + c[k], 0);
  return { chans, days: Sc.days, mix: Sc.mix, sscl, vat, tot: { net: T('net'), rcT: T('rcT'), spots: T('spots'), sscl: T('sscl'), vat: T('vat'), total: T('total') } };
}
function vSchedule(el) {
  const A = S.cur, P = S.P;
  if (!A || !S.sched) { el.innerHTML = `<div class="empty"><b>No plan to schedule yet</b><p>Build a plan in the Planner first.</p><button class="btn pri" data-tab-go="plan">Go to Planner</button></div>`; return; }
  const M = scheduleModel(), meta = lsGet('deckMeta', {});
  if (!S.SCH.ch || !M.chans.find(c => c.ch === S.SCH.ch)) S.SCH.ch = M.chans[0].ch;
  const C = M.chans.find(c => c.ch === S.SCH.ch), days = M.days;
  const months = []; days.forEach(d => { const k = d.iso.slice(0, 7); const l = months[months.length - 1]; if (l && l.k === k) l.n++; else months.push({ k, n: 1, label: MONS[+k.slice(5) - 1] + ' ' + k.slice(0, 4) }); });
  const wk = d => d.day === 'Saturday' || d.day === 'Sunday';
  const L = v => ni(Math.round(v));
  const head = `<tr class="h1"><th class="sx" colspan="12"></th>${months.map(m => `<th class="mth" colspan="${m.n}">${m.label}</th>`).join('')}</tr>
    <tr><th class="sx l">Programme</th><th class="l">Day</th><th>From</th><th>To</th><th>Dur</th><th>Spots</th><th class="l">Brand</th><th title="Rate card for 30 seconds. Estimated values in italics: type the real rate.">Rate card 30s</th><th>Rate card</th><th>Negotiated</th><th>Total rate card</th><th>Total negotiated</th>${days.map(d => `<th class="dc ${wk(d) ? 'we' : ''}">${d.day[0]}<br>${+d.iso.slice(8)}</th>`).join('')}</tr>`;
  const body = C.sections.map(sec => `<tr class="secr"><td class="sx l" colspan="12">${esc(sec.m.name)} · ${sec.m.dur} sec</td>${days.map(d => `<td class="${wk(d) ? 'we' : ''}"></td>`).join('')}</tr>` +
    sec.rows.map(r => `<tr><td class="sx l"><button class="nm" ${xa({ ch: r.x.ch, p: r.x.p })}>${esc(pn(r.x.p))}</button></td><td class="l">${r.x.pattern}</td><td>${r.x.from}</td><td>${r.x.to}</td><td>${sec.m.dur}</td><td><b>${r.spots}</b></td><td class="l">${esc(sec.m.name)}</td>
      <td><input class="rate ${r.x.rateSet ? '' : 'est'}" data-rate="${esc(r.x.key)}" value="${Math.round(r.x.rate30)}" title="${r.x.rateSet ? 'Your rate card value' : 'Estimated: CPRP × TVR (min. rate). Type the real 30-sec rate.'}"></td>
      <td>${L(r.rc)}</td><td>${L(r.ng)}</td><td>${L(r.trc)}</td><td><b>${L(r.tng)}</b></td>${days.map(d => { const v = r.byDate[d.iso]; return `<td class="dc ${wk(d) ? 'we' : ''} ${v ? 'on' : ''}">${v || ''}</td>`; }).join('')}</tr>`).join('')).join('');
  const dayTot = days.map(d => C.sections.reduce((a, sec) => a + sec.rows.reduce((b, r) => b + (r.byDate[d.iso] || 0), 0), 0));
  const foot = `<tr class="tot"><td class="sx l" colspan="5">${esc(chName(C.ch))} total</td><td>${C.spots}</td><td></td><td></td><td></td><td></td><td>${L(C.rcT)}</td><td>${L(C.net)}</td>${dayTot.map((v, i) => `<td class="dc ${wk(days[i]) ? 'we' : ''}">${v || ''}</td>`).join('')}</tr>`;
  const sm = (l, v, cls = '') => `<div class="m ${cls}"><small>${l}</small><b>${v}</b></div>`;
  el.innerHTML = `
  <div class="vhead"><div><h2>Booking schedule</h2><p>Day-by-day spots per channel, by creative, ready to send to channels. Same layout as the Excel export.</p></div>
    <span class="push"><button class="btn pri" id="sch-xlsx">Export Excel schedule</button></span></div>
  ${A.fixed ? `<div class="frozen ${A.over > 0 ? 'over' : ''}"><div><b>Schedule kept as planned while you enter rates.</b> Same programmes, spots and days; only the costs changed. Budget used <b>LKR ${L(A.spent)}</b> of LKR ${L(A.B)}${A.over > 0 ? ` · <b>over by LKR ${L(A.over)}</b>` : ` · LKR ${L(A.unspent)} left`}.</div><button class="btn pri" data-reopt>Re-optimise with these rates</button></div>` : ''}
  <div class="planbar">${sm('Campaign period', fmtDate(S.sched.start) + ' – ' + fmtDate(S.sched.end), 'hl')}${sm('Spots', ni(M.tot.spots))}${sm('Rate card value', 'LKR ' + L(M.tot.rcT))}${sm('Negotiated (investment)', 'LKR ' + L(M.tot.net))}${sm('Total with taxes', 'LKR ' + L(M.tot.total))}${sm('Creatives', esc(mixText(P)))}</div>
  <div class="grid g2">
    ${panel('Schedule details', 'shown on the Excel cover and channel sheets', `<div class="two">
      <label class="fld"><span>Client</span><input type="text" data-meta="client" value="${esc(meta.client || '')}"></label>
      <label class="fld"><span>Brand</span><input type="text" data-meta="brand" value="${esc(meta.brand || '')}"></label>
      <label class="fld"><span>Campaign</span><input type="text" data-meta="campaign" value="${esc(meta.campaign || '')}"></label>
      <label class="fld"><span>Primary TG</span><input type="text" data-meta="tg" value="${esc(meta.tg || '')}" placeholder="e.g. 16-50 Male & Female SEC All"></label>
      <label class="fld"><span>SSCL %</span>${`<input type="number" data-p="sscl" min="0" max="20" step="0.5" value="${P.sscl}">`}</label>
      <label class="fld"><span>VAT %</span>${`<input type="number" data-p="vat" min="0" max="30" step="0.5" value="${P.vat}">`}</label></div>
      <p class="hint">Change the campaign dates, creatives or caps in <button class="link" data-tab-go="plan">Planner settings</button>; the schedule is rebuilt automatically.</p>`)}
    ${panel('By channel', 'negotiated rate = rate card × (1 − discount)', `<div class="tw"><table><thead><tr><th>Channel</th><th>Discount %</th><th>Spots</th><th>Rate card</th><th>Negotiated</th><th>With taxes</th></tr></thead><tbody>
      ${M.chans.map(c => `<tr><td>${dot(c.ch)}${esc(chName(c.ch))}</td><td><input class="rate" type="number" min="0" max="95" step="1" data-disc="${esc(c.ch)}" value="${c.disc}" style="width:64px"></td><td>${c.spots}</td><td>${L(c.rcT)}</td><td><b>${L(c.net)}</b></td><td>${L(c.total)}</td></tr>`).join('')}
      <tr class="tot"><td>Total</td><td></td><td>${M.tot.spots}</td><td>${L(M.tot.rcT)}</td><td><b>${L(M.tot.net)}</b></td><td>${L(M.tot.total)}</td></tr></tbody></table></div>
      <p class="hint">Changing a rate or discount changes spot costs, so the plan is re-optimised.</p>`)}
  </div>
  <div class="panel"><div class="ph"><h3>Channel sheet</h3><span class="push"><div class="chips">${M.chans.map(c => `<button class="chip ${c.ch === C.ch ? 'on' : ''}" data-sch-ch="${esc(c.ch)}">${esc(chName(c.ch))} · ${c.spots}</button>`).join('')}</div></span></div>
    <div class="pb"><div class="tw sched-wrap"><table class="sched"><thead>${head}</thead><tbody>${body}</tbody><tfoot>${foot}</tfoot></table></div>
    <div class="taxes"><span>Total negotiated <b>LKR ${L(C.net)}</b></span><span>SSCL ${P.sscl}% <b>LKR ${L(C.sscl)}</b></span><span>VAT ${P.vat}% <b>LKR ${L(C.vat)}</b></span><span>Total with taxes <b>LKR ${L(C.total)}</b></span></div>
    <p class="hint">Rate card 30s in <i>italics</i> is estimated from the rating; type the channel's rate card to replace it. Shaded columns are weekends. Click a programme for its ratings detail.</p></div></div>`;
}
async function runXlsx(btn) {
  if (!S.cur || !S.sched) { toast('Build a plan first'); return; }
  btn.disabled = true; const t = btn.textContent; btn.textContent = 'Building…';
  try { const name = await exportScheduleXlsx(scheduleModel(), { meta: lsGet('deckMeta', {}), P: S.P, start: S.sched.start, end: S.sched.end, net: S.cur.net, r3: S.cur.r3 }); toast('Downloaded ' + name); }
  catch (e) { toast('Excel export failed: ' + e.message); }
  finally { btn.disabled = false; btn.textContent = t; }
}

/* ---------- Scenarios ---------- */
function snapshot(name, A, P, FS) {
  return {
    id: Date.now().toString(36), name, created: new Date().toISOString(),
    P: JSON.parse(JSON.stringify(P)),
    FS: { from: FS.from, to: FS.to, ch: [...FS.ch], cat: [...FS.cat], day: [...FS.day], h0: FS.h0, h1: FS.h1, minTvr: FS.minTvr, q: FS.q, p: FS.p },
    r: { B: A.B, spent: A.spent, net: A.net, r3: A.r3, gross: A.gross, loss: A.loss, spots: A.spots, freq: A.freq, tiers: A.tiers.map(t => ({ t: t.t, target: t.target, actual: t.actual })), chs: A.chs.map(c => ({ ch: c.ch, w: c.w, bud: c.bud })), items: A.kept.map(x => ({ ch: x.ch, p: x.p, k: x.k, bud: x.bud, spots: x.spots, mean: x.mean, rp: x.rp, ci: x.ci, role: x.role, tier: x.tier, hour: x.hour, cat: x.cat, n: x.n, days: x.days })) }
  };
}
function saveScenario(name, A, P = S.P, FS = S.FS) {
  if (!A) { toast('Nothing to save: the plan is empty'); return; }
  S.scenarios.push(snapshot(name || ('Scenario ' + String.fromCharCode(65 + S.scenarios.length % 26)), A, P, FS));
  lsSet('scenarios', S.scenarios); toast('Scenario saved');
}
function loadScenario(sc) {
  S.P = Object.assign({}, DEFAULT_P, JSON.parse(JSON.stringify(sc.P))); savePlan();
  const f = sc.FS;
  S.FS = { from: f.from, to: f.to, ch: new Set(f.ch.filter(c => S.CH.includes(c))), cat: new Set(f.cat.filter(c => S.CATS.includes(c))), day: new Set(f.day), h0: f.h0, h1: f.h1, minTvr: f.minTvr, q: f.q || '', p: f.p || '' };
  if (!S.FS.ch.size) S.FS.ch = new Set(S.CH); if (!S.FS.cat.size) S.FS.cat = new Set(S.CATS);
  applyFilters(); setTab('plan'); toast('Loaded "' + sc.name + '"');
}
const curAsScen = () => S.cur ? snapshot('Current plan', S.cur, S.P, S.FS) : null;
function vScen(el) {
  const cur = curAsScen();
  const list = S.scenarios;
  const card = (sc, isCur) => {
    const r = sc.r, d = cur && !isCur ? (r.net / cur.r.net - 1) * 100 : 0;
    return `<div class="scc ${isCur ? 'cur' : ''}"><span><span class="tag ${isCur ? 'support' : ''}">${esc(sc.name)}</span>${sc.P.cut ? ` <span class="tag review">−${sc.P.cut}% budget</span>` : ''}</span>
      <div class="big">${nf(r.net, 1)}%</div><div class="muted">Est. net reach · reach 3+ ${r.r3 != null ? nf(r.r3, 1) + '%' : '—'} · ${lkr(r.B)}</div>
      ${!isCur && cur ? `<div class="delta ${d >= 0 ? 'up' : 'dn'}">${d >= 0 ? '+' : ''}${nf(d, 1)}% reach vs current · ${nf((r.B / cur.r.B - 1) * 100, 0)}% budget</div>` : '<div class="delta muted">live, follows your filters and settings</div>'}
      <div class="mini">${r.chs.map(c => `<div style="width:${(c.w * 100).toFixed(1)}%;background:${cv(c.ch)}"></div>`).join('')}</div>
      <div class="foot">${r.chs.map(c => esc(chName(c.ch).replace(' TV', '')) + ' ' + nf(c.w * 100, 0)).join(' · ')}</div>
      <div class="foot">${r.items.length} programs${r.spots ? ' · ' + ni(r.spots) + ' spots' : ''}</div>
      <div class="acts">${isCur ? `<button class="btn sm pri" id="sc-save">Save current</button>` : `<button class="btn sm" data-sc-load="${sc.id}">Load</button><button class="btn sm ${S.cmp === sc.id ? 'pri' : ''}" data-sc-cmp="${sc.id}">Compare</button><button class="btn sm" data-sc-csv="${sc.id}">CSV</button><button class="btn sm danger" data-sc-del="${sc.id}">Delete</button>`}</div></div>`;
  };
  const quick = [10, 25, 40].map(c => ({ c, r: S.cur ? simulate({ cut: c }) : null }));
  const sel = list.find(s => s.id === S.cmp);
  el.innerHTML = `
  <div class="vhead"><div><h2>Test scenarios and export</h2><p>Save versions of the plan, compare them side by side, then export the one you choose. Scenarios are kept in this browser only.</p></div></div>
  <div class="sc">${cur ? card(cur, true) : ''}${list.map(s => card(s, false)).join('')}
    <div class="scc new"><div style="font-size:28px;color:var(--accent);line-height:1">+</div><b style="color:var(--ink)">New scenario</b><div class="foot">Change the plan, then save it, or ask Planner AI for a what-if</div><button class="btn sm" data-ai="Cut budget 25%">Ask AI: cut 25%</button></div></div>
  ${S.cur ? panel('Quick budget what-ifs', 're-optimised at each budget; cuts remove repeat spots first, so reach 3+ falls faster than net reach', `<div class="tw"><table><thead><tr><th>Budget</th><th>Net reach</th><th>Change</th><th>Reach 3+</th><th>Spots</th><th>Split</th><th></th></tr></thead><tbody>
      <tr><td>${lkr(S.base.B)} (full)</td><td>${nf(S.base.net, 1)}%</td><td>—</td><td>${nf(S.base.r3, 1)}%</td><td>${S.base.spots}</td><td class="l">${S.base.chs.map(c => esc(chName(c.ch).replace(' TV', '')) + ' ' + nf(c.w * 100, 0)).join(' · ')}</td><td></td></tr>
      ${quick.map(q => q.r ? `<tr><td>−${q.c}% (${lkr(q.r.B)})</td><td>${nf(q.r.net, 1)}%</td><td style="color:var(--bad)">${nf((q.r.net / S.base.net - 1) * 100, 1)}%</td><td>${nf(q.r.r3, 1)}%</td><td>${q.r.spots}</td><td class="l">${q.r.chs.map(c => esc(chName(c.ch).replace(' TV', '')) + ' ' + nf(c.w * 100, 0)).join(' · ')}</td><td><button class="btn sm" data-quick="${q.c}">Save</button></td></tr>` : '').join('')}
    </tbody></table></div>`) : ''}
  ${sel && cur ? compareHTML(cur, sel) : ''}
  ${panel('Export', 'choose an output for the current plan', `<div class="ex">
    <button data-deck><b>Client deck</b><small>PowerPoint · client-ready, plain-English commentary and speaker notes</small></button>
    <button id="ex-csv"><b>Spot plan</b><small>CSV · channel × program × slot × budget × spots</small></button>
    <button id="ex-print"><b>Print / PDF</b><small>Planner page, print-ready</small></button>
    <button id="ex-copy"><b>Copy memo</b><small>Plain text for email or deck</small></button>
    <button id="ex-data"><b>Filtered data</b><small>CSV of the airings behind this plan</small></button></div>`)}`;
}
function compareHTML(a, b) {
  const ra = a.r, rb = b.r;
  const row = (l, x, y, f) => `<tr><td>${l}</td><td>${f(x)}</td><td>${f(y)}</td><td>${typeof x === 'number' ? `<span style="color:${y - x >= 0 ? 'var(--good)' : 'var(--bad)'}">${y - x >= 0 ? '+' : ''}${f(y - x)}</span>` : ''}</td></tr>`;
  const ch = [];
  const ma = new Map(ra.chs.map(c => [c.ch, c.w])), mb = new Map(rb.chs.map(c => [c.ch, c.w]));
  new Set([...ma.keys(), ...mb.keys()]).forEach(c => {
    const x = ma.get(c) || 0, y = mb.get(c) || 0;
    if (!x) ch.push(`<li><span class="ic ok">+</span><span>Adds ${esc(chName(c))} at ${nf(y * 100, 0)}%</span></li>`);
    else if (!y) ch.push(`<li><span class="ic wa">−</span><span>Removes ${esc(chName(c))} (was ${nf(x * 100, 0)}%)</span></li>`);
    else if (Math.abs(y - x) >= .02) ch.push(`<li><span class="ic ${y > x ? 'ok' : 'wa'}">${y > x ? '↑' : '↓'}</span><span>${esc(chName(c))} ${nf(x * 100, 0)}% → ${nf(y * 100, 0)}%</span></li>`);
  });
  const pa = new Set(ra.items.map(x => x.ch + '||' + x.p)), pb = new Set(rb.items.map(x => x.ch + '||' + x.p));
  const add = [...pb].filter(k => !pa.has(k)), rem = [...pa].filter(k => !pb.has(k));
  if (add.length) ch.push(`<li><span class="ic ok">+</span><span>Programs added: ${add.map(k => esc(pn(k.split('||')[1]))).join(', ')}</span></li>`);
  if (rem.length) ch.push(`<li><span class="ic wa">−</span><span>Programs dropped: ${rem.map(k => esc(pn(k.split('||')[1]))).join(', ')}</span></li>`);
  if (!ch.length) ch.push(`<li><span class="ic in">i</span><span>Same channels and programs; only the numbers differ.</span></li>`);
  return `<div class="grid g2">${panel(`Compare: current vs "${esc(b.name)}"`, '', `<table><thead><tr><th>Metric</th><th>Current</th><th>${esc(b.name)}</th><th>Difference</th></tr></thead><tbody>
    ${row('Net reach %', ra.net, rb.net, v => nf(v, 1))}${ra.r3 != null && rb.r3 != null ? row('Reach 3+ %', ra.r3, rb.r3, v => nf(v, 1)) : ''}${row('Gross reach %', ra.gross, rb.gross, v => nf(v, 1))}${row('Budget (LKR M)', ra.B / 1e6, rb.B / 1e6, v => nf(v, 2))}
    ${row('Programs', ra.items.length, rb.items.length, v => ni(v))}${row('Channels', ra.chs.length, rb.chs.length, v => ni(v))}${ra.spots || rb.spots ? row('Est. spots', ra.spots, rb.spots, v => ni(v)) : ''}
    ${row('Net reach per LKR 1M', ra.net / (ra.B / 1e6 || 1), rb.net / (rb.B / 1e6 || 1), v => nf(v, 2))}</tbody></table>`)}
    ${panel(`What changes in "${esc(b.name)}"`, 'compared with the current plan', `<ul class="health">${ch.join('')}</ul>`)}</div>`;
}
function planCSV(sc) {
  const r = sc.r;
  const rows = r.items.map(x => [chName(x.ch), pn(x.p), x.tier ? 'Tier ' + x.tier : '', pn(x.cat), band(x.hour), (x.days || []).map(d => d.slice(0, 3)).join(' '), x.n, x.mean.toFixed(2), x.rp.toFixed(2), x.ci.toFixed(1), x.role, (x.k * 100).toFixed(1), Math.round(x.bud), x.spots]);
  const head = ['Channel', 'Program', 'Tier', 'Category', 'Usual slot', 'Days aired', 'Airings in data', 'Avg TVR', 'Avg reach %', 'Steadiness', 'Role', 'Share of plan %', 'Budget LKR', 'Est. spots'];
  const top = [['Plan', sc.name], ['Budget LKR', Math.round(r.B)], ['Est. net reach %', r.net.toFixed(1)], ['Reach 3+ %', r.r3 != null ? r.r3.toFixed(1) : ''], ['Gross reach %', r.gross.toFixed(1)], ['Split', r.chs.map(c => chName(c.ch) + ' ' + (c.w * 100).toFixed(0) + '%').join('; ')], []];
  return toCSV(['TV Media Planner', 'export ' + new Date().toISOString().slice(0, 10)], top) + '\n' + toCSV(head, rows);
}

/* ---------- Duplication ---------- */
// Union of audiences where d(a, b) gives the overlap factor between two items.
function unionOf(list, dfn) {
  let U = 0; const inc = [];
  [...list].sort((a, b) => b.R - a.R).forEach(c => {
    if (!inc.length) U = c.R;
    else { const d = inc.reduce((s, x) => s + dfn(x, c), 0) / inc.length; U = uni(U, c.R, d); }
    inc.push(c);
  });
  return U;
}
const explain = (txt) => `<div class="explain"><span class="ic in">${ICON.in}</span><div>${txt}</div></div>`;
const meaning = (txt) => `<div class="meaning"><b>What this means</b><p>${txt}</p></div>`;
const howRead = (txt) => `<p class="howread"><b>How to read this:</b> ${txt}</p>`;

function slotData() {
  const SD = S.SD, A = S.cur;
  const planKeys = A ? new Set(A.kept.map(x => x.key)) : new Set();
  const planCh = A ? new Set(A.chs.map(c => c.ch)) : new Set(S.CH);
  const dayOk = r => SD.sdDays === 'all' || (SD.sdDays === 'wd' ? DAYS.indexOf(r.day) < 5 : DAYS.indexOf(r.day) >= 5);
  const scopeOk = r => SD.sdScope === 'plan' ? planKeys.has(r.ch + '||' + r.p) : SD.sdScope === 'planch' ? planCh.has(r.ch) : true;
  const base = S.F.filter(r => dayOk(r) && scopeOk(r));
  const hours = [...grp(base, r => r.h).values()].sort((a, b) => a.k - b.k);
  if (SD.h == null || !hours.find(g => g.k === SD.h)) {
    const best = [...hours].sort((a, b) => new Set(base.filter(r => r.h === b.k).map(r => r.ch + r.p)).size - new Set(base.filter(r => r.h === a.k).map(r => r.ch + r.p)).size || avgT(b) - avgT(a))[0];
    SD.h = best ? best.k : null;
  }
  const rows = base.filter(r => r.h === SD.h);
  const items = [...grp(rows, r => r.ch + '||' + r.p).values()].map(g => { const [ch, p] = g.k.split('||'); return { id: g.k, ch, p, R: avgR(g) || avgT(g) * 1.4, tvr: avgT(g), n: g.n, days: DAYS.filter(d => g.dc[d]) }; })
    .sort((a, b) => b.R - a.R).slice(0, 10);
  const dfn = (a, b) => a.ch === b.ch ? S.DINTRA : getD(a.ch, b.ch);
  const net = unionOf(items, dfn), gross = items.reduce((s, x) => s + x.R, 0);
  items.forEach(x => { x.uniq = Math.max(0, net - unionOf(items.filter(y => y !== x), dfn)); x.shared = Math.max(0, x.R - x.uniq); });
  const pairs = [];
  items.forEach((a, i) => items.forEach((b, j) => { if (j > i) { const d = dfn(a, b); pairs.push({ a, b, d, ov: a.R + b.R - uni(a.R, b.R, d) }); } }));
  return { hours, items, net, gross, pairs, dfn };
}

function vDup(el) {
  const present = new Set(S.F.map(r => r.ch)), chs = S.CH.filter(c => present.has(c)), n = chs.length;
  const A = S.cur;
  // --- worked example from the two biggest plan channels
  let ex = '';
  if (A && A.chs.length >= 2) {
    const a = A.chs[0], b = A.chs[1], d = getD(a.ch, b.ch), un = uni(a.R, b.R, d), both = a.R + b.R - un;
    ex = `Over this campaign, <b>${esc(chName(a.ch))}</b> reaches ${nf(a.R, 1)}% and <b>${esc(chName(b.ch))}</b> reaches ${nf(b.R, 1)}%. Added together that is ${nf(a.R + b.R, 1)}%, but about <b>${nf(both, 1)}%</b> watch both, so the real number of different people is <b>${nf(un, 1)}%</b>.`;
  }
  // --- matrix
  let g = `<div class="tw"><div class="mx" style="grid-template-columns:110px repeat(${n},minmax(50px,1fr))"><div></div>${chs.map(c => `<div class="mh" title="${esc(chName(c))}">${esc(chName(c).replace(' TV', ''))}</div>`).join('')}`;
  chs.forEach(a => {
    g += `<div class="hr">${dot(a)}${esc(chName(a))}</div>`;
    chs.forEach(b => {
      if (a === b) { g += `<input disabled value="—" aria-label="same channel">`; return; }
      const d = getD(a, b), edited = S.DUP[dkey(a, b)] != null;
      g += `<input type="number" step="0.05" min="0" max="1" data-a="${esc(a)}" data-b="${esc(b)}" value="${d.toFixed(2)}" title="${esc(chName(a))} and ${esc(chName(b))}: ${nf(d * 100, 0)} on a 0–100 overlap scale" style="background:color-mix(in srgb,var(--accent) ${(d * 70).toFixed(0)}%,var(--panel));color:${d > .6 ? 'var(--accent-ink)' : 'var(--ink)'};${edited ? 'font-weight:700;border-color:var(--ink)' : ''}" aria-label="Duplication ${esc(chName(a))} with ${esc(chName(b))}">`;
    });
  });
  g += '</div></div>';
  // --- gross vs net
  const st = channelStats(S.F).map(c => ({ ch: c.ch, R: c.reach || c.tvr * 1.4 }));
  const gs = [], ns = []; let gr = 0;
  st.forEach((c, i) => { gr += c.R; gs.push(+gr.toFixed(1)); ns.push(+netReach(st.slice(0, i + 1), getD).toFixed(1)); });
  const gains = ns.map((v, i) => i ? v - ns[i - 1] : v);
  let knee = 1; gains.forEach((v, i) => { if (i && v >= 1) knee = i + 1; });
  const gnMeaning = st.length > 1 ? `Watching one average programme on each of the top ${st.length} channels adds up to ${nf(gs[gs.length - 1], 1)}%, but only about <b>${nf(ns[ns.length - 1], 1)}% are different people</b>. The gap (${nf(gs[gs.length - 1] - ns[ns.length - 1], 1)} points) is people counted twice. After ${knee} channel${knee > 1 ? 's' : ''}, each extra channel adds less than 1 point of new people.` : 'Only one channel is in the filter.';
  // --- slot
  const SD = slotData();
  const hopts = SD.hours.map(h => `<option value="${h.k}" ${h.k === S.SD.h ? 'selected' : ''}>${band(h.k)}</option>`).join('');
  const it = SD.items;
  let slotHTML;
  if (!it.length) slotHTML = '<p class="muted">No programmes in this slot for the chosen days and scope. Try “All programmes” or another hour.</p>';
  else {
    const dupPct = SD.gross > 0 ? (SD.gross - SD.net) / SD.gross * 100 : 0;
    const exc = [...it].sort((a, b) => b.uniq / b.R - a.uniq / a.R);
    const top = [...SD.pairs].sort((a, b) => b.ov - a.ov)[0];
    const sn = s => esc(pn(s.length > 22 ? s.slice(0, 21) + '…' : s));
    const mxOv = Math.max(...SD.pairs.map(p => p.ov / Math.min(p.a.R, p.b.R)), .01);
    const mtx = `<div class="tw"><table class="pdm"><thead><tr><th></th>${it.map((x, i) => `<th title="${esc(pn(x.p))} (${esc(chName(x.ch))})">${i + 1}</th>`).join('')}</tr></thead><tbody>${it.map((a, i) => `<tr><td class="l"><b>${i + 1}</b> ${sn(a.p)}<span class="sub">${dot(a.ch)}${esc(chName(a.ch))}</span></td>${it.map((b, j) => {
      if (i === j) return `<td class="diag" title="${esc(pn(a.p))} reaches ${nf(a.R, 1)}%">${nf(a.R, 1)}</td>`;
      const pr = SD.pairs.find(p => (p.a === a && p.b === b) || (p.a === b && p.b === a)), sh = pr.ov / Math.min(a.R, b.R);
      return `<td style="background:color-mix(in srgb,var(--heat) ${(8 + 80 * sh / mxOv).toFixed(0)}%,var(--panel))" title="About ${nf(pr.ov, 1)}% of viewers watch both ${esc(pn(a.p))} and ${esc(pn(b.p))} (${nf(sh * 100, 0)}% of the smaller audience)${a.ch === b.ch ? '. Same channel.' : ''}">${nf(pr.ov, 1)}</td>`;
    }).join('')}</tr>`).join('')}</tbody></table></div>`;
    const sameCh = it.some((a, i) => it.some((b, j) => j > i && a.ch === b.ch));
    slotHTML = `
      <div class="kpis k4">
        ${kp('Programmes in this slot', it.length, `${band(S.SD.h)} · ${{ all: 'all days', wd: 'weekdays', we: 'weekends' }[S.SD.sdDays]}`)}
        ${kp('Reach added up', nf(SD.gross, 1) + '%', 'each programme counted separately')}
        ${kp('Different people', nf(SD.net, 1) + '%', 'each person counted once', 'hl')}
        ${kp('Duplicated', nf(dupPct, 0) + '%', 'of the added-up reach is the same people')}
      </div>
      <div class="grid g2" style="margin-top:12px">
        <div>
          <h4 class="h4">How much of each programme's audience is unique?</h4>
          <div class="chart" style="height:${Math.max(220, it.length * 34 + 70)}px"><canvas id="c-sdup"></canvas></div>
          ${howRead('each bar is a programme\'s reach. The <b>teal part</b> is viewers <i>only</i> this programme brings in. The <b>grey part</b> is viewers who also watch another programme in this list. Click a bar for the programme detail.')}
        </div>
        <div>
          <h4 class="h4">Who watches both? (programme × programme)</h4>
          ${mtx}
          ${howRead(`numbers are the estimated % of all TV viewers who watch <b>both</b> programmes (on different nights). The grey diagonal is each programme's own reach. Darker orange = a bigger share of the smaller programme's audience also watches the other.${sameCh ? ' Programmes on the same channel use the within-channel overlap.' : ''}`)}
        </div>
      </div>
      ${meaning(`If you buy all ${it.length} programmes at ${band(S.SD.h)}, they add up to ${nf(SD.gross, 1)}% but reach about <b>${nf(SD.net, 1)}% different people</b>. <b>${esc(pn(exc[0].p))}</b> (${esc(chName(exc[0].ch))}) brings the most viewers no one else here reaches (${nf(exc[0].uniq / exc[0].R * 100, 0)}% of its audience), so it is the best value for new reach. <b>${esc(pn(exc[exc.length - 1].p))}</b> mostly repeats viewers the others already reach (only ${nf(exc[exc.length - 1].uniq / exc[exc.length - 1].R * 100, 0)}% unique). ${top ? `The biggest overlap is between <b>${esc(pn(top.a.p))}</b> and <b>${esc(pn(top.b.p))}</b> (about ${nf(top.ov, 1)}% watch both).` : ''} These programmes air at the same hour, so on any one night a viewer can only watch one of them: the overlap is people who switch between them on different nights. Placing spots on all of them on the same night (a roadblock) reaches the most different people.`)}`;
  }
  el.innerHTML = `
  <div class="vhead"><div><h2>Duplication: counting each viewer once</h2><p>Many people watch more than one channel or programme. This page shows how much the audiences overlap, so reach is not counted twice.</p></div>
    <span class="push"><button class="btn sm" id="dup-reset">Reset overlap factors</button></span></div>
  ${explain(`<b>The idea in one line:</b> if 30% of people watch channel A and 25% watch channel B, you do not reach 55%. Some people watch both, and they should be counted only once. ${ex}`)}

  <h3 class="sec"><span>1</span> Between channels</h3>
  <div class="grid g2">
    ${panel('Channel overlap factors', 'how strongly two channels share viewers', g +
      `<div class="legend" style="margin-top:10px">Low overlap ${[8, 25, 45, 60, 70].map(p => `<span class="sw" style="background:color-mix(in srgb,var(--accent) ${p}%,var(--panel))"></span>`).join('')} High overlap</div>` +
      howRead('each cell is a factor from 0 to 1 for a pair of channels. <b>0</b> = the two audiences overlap only by chance. <b>1</b> = everyone in the smaller audience also watches the other channel. Big Sinhala channels (Hiru, Derana) share many viewers (about 0.70); Sinhala and Tamil channels share few (about 0.15–0.20). These are planning assumptions; type a new value if you have measured data. Bold cells are your edits.') +
      `<div class="two" style="margin-top:12px"><label class="fld"><span>Overlap between programmes on the same channel</span><input type="number" id="d-intra" step="0.05" min="0" max="1" value="${S.DINTRA}"></label>
       <label class="fld"><span>New viewers from each repeat spot (% of the previous spot)</span><input type="number" id="d-rep" step="5" min="0" max="90" value="${S.P.repQ}"></label></div>`)}
    ${panel('Added-up reach vs different people', 'channels added in order of strength', '<div class="chart"><canvas id="c-net"></canvas></div>' +
      howRead('the <b>grey</b> bars add each channel\'s average reach on top of the previous ones, counting people twice. The <b>teal</b> bars count each person once. The growing gap between them is duplication.') + meaning(gnMeaning))}
  </div>

  <h3 class="sec"><span>2</span> Between programmes in the same time slot</h3>
  ${panel('Programme duplication in a slot', 'pick an hour to see how much its programmes share viewers',
    `<div class="inl" style="margin-bottom:12px"><label class="inl"><span class="muted">Slot</span><select id="sd-h">${hopts}</select></label>
      ${seg('sdDays', S.SD, [['all', 'All days'], ['wd', 'Weekdays'], ['we', 'Weekend']])}
      ${seg('sdScope', S.SD, [['plan', 'Plan programmes'], ['planch', 'All on plan channels'], ['all', 'All channels']])}</div>` + slotHTML +
    `<p class="hint">Estimates: the ratings file has no viewer-level data, so programme overlap is modelled from the channel overlap factors above (and the within-channel factor for two programmes on the same channel). Top 10 programmes by reach are shown.</p>`)}

  <h3 class="sec"><span>3</span> Competing programmes at the same hour</h3>
  ${panel('Head-to-head slots', 'the strongest programme per channel at each busy hour', '<div id="x2" style="max-height:340px;overflow:auto"></div>' +
    howRead('each block is an hour where several of the top channels air programmes at the same time, sorted by combined rating. Click the hour to drill in.') +
    meaning('At the same hour, viewers must choose one channel, so these programmes split the audience rather than duplicate it. Buying several of them on the same night reaches more different people (roadblock); spreading them over different nights repeats the message to the same people (stagger). Choose this in the Planner under Flighting.'))}`;

  mkChart('c-net', {
    type: 'bar', data: { labels: st.map((c, i) => (i + 1) + '. ' + chName(c.ch)), datasets: [{ label: 'Added up (counts people twice)', data: gs, backgroundColor: css('--line'), borderRadius: 3 }, { label: 'Different people', data: ns, backgroundColor: css('--accent'), borderRadius: 3 }] },
    options: { plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.raw}%`, afterBody: items => { const i = items[0].dataIndex; return `Counted twice: ${nf(gs[i] - ns[i], 1)} pts`; } } } }, scales: { x: { grid: { display: false }, ticks: { maxRotation: 35 } }, y: { beginAtZero: true, title: { display: true, text: '% of TV viewers' } } } }
  });
  if (it.length) mkChart('c-sdup', {
    type: 'bar', data: {
      labels: it.map((x, i) => `${i + 1}. ${pn(x.p).slice(0, 24)} (${chName(x.ch).replace(' TV', '')})`), datasets: [
        { label: 'Only this programme reaches them', data: it.map(x => +x.uniq.toFixed(2)), backgroundColor: css('--accent'), borderRadius: 2 },
        { label: 'Also watch another programme here', data: it.map(x => +x.shared.toFixed(2)), backgroundColor: css('--line'), borderRadius: 2 }]
    },
    options: { indexAxis: 'y', onClick: (e, els) => els.length && openDetail({ ch: it[els[0].index].ch, p: it[els[0].index].p }), plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } }, tooltip: { callbacks: { title: items => { const x = it[items[0].dataIndex]; return `${pn(x.p)} · ${chName(x.ch)}`; }, label: c => ` ${c.dataset.label}: ${nf(c.raw, 1)}%`, footer: items => { const x = it[items[0].dataIndex]; return `Reach ${nf(x.R, 1)}% · ${nf(x.uniq / x.R * 100, 0)}% unique · TVR ${nf(x.tvr, 1)} · ${x.n} airings`; } } } }, scales: { x: { stacked: true, beginAtZero: true, title: { display: true, text: '% of TV viewers' } }, y: { stacked: true, grid: { display: false } } } }
  });
  const cp = competing(S.F, st.slice(0, 5).map(c => c.ch), 10);
  $('#x2').innerHTML = cp.length ? `<table><tbody>${cp.map(o => `<tr><td class="l" style="vertical-align:top"><button class="ib" ${xa({ h: o.h })}>${band(o.h)}</button></td><td class="l" style="white-space:normal">${o.l.slice(0, 4).map(x => `<span class="sub" style="color:var(--ink)">${dot(x.ch)}${esc(chName(x.ch))}: ${esc(pn(x.p))} <span class="muted">TVR ${nf(x.m, 1)}</span></span>`).join('')}</td></tr>`).join('')}</tbody></table>` : '<p class="muted">No overlapping hours found.</p>';
}

/* ---------- Data table ---------- */
const DCOLS = [['date', 'Date', 'l'], ['day', 'Day', 'l'], ['s', 'Start', 'l'], ['e', 'End', 'l'], ['ch', 'Channel', 'l'], ['p', 'Program', 'l'], ['cat', 'Category', 'l'], ['dur', 'Min'], ['tvr', 'TVR'], ['sh', 'Share %'], ['rp', 'Reach %']];
function vData(el) {
  const k = S.DSORT.k, d = S.DSORT.d, q = S.DQ.trim().toLowerCase();
  let rows = q ? S.F.filter(r => (r.p + ' ' + r.ch + ' ' + r.cat + ' ' + r.date).toLowerCase().includes(q)) : S.F;
  rows = [...rows].sort((a, b) => { const x = a[k], y = b[k]; return (typeof x === 'number' ? x - y : String(x).localeCompare(y)) * d; });
  const PS = 100, pages = Math.max(1, Math.ceil(rows.length / PS)); S.DPAGE = Math.min(S.DPAGE, pages - 1);
  const pg = rows.slice(S.DPAGE * PS, S.DPAGE * PS + PS);
  const odd = S.F.filter(r => r.sh > 100).length, noR = S.F.filter(r => !r.rp).length;
  el.innerHTML = `
  <div class="vhead"><div><h2>Data</h2><p>${ni(S.F.length)} filtered airings of ${ni(S.ROWS.length)} loaded${S.meta ? ' · ' + esc(S.meta.name) : ''}. Click a row for its program detail.</p></div>
    <span class="push"><input type="search" id="dq" placeholder="Search program, channel, category, date" value="${esc(S.DQ)}" style="width:280px"><button class="btn" id="ex-data2">Export CSV</button></span></div>
  ${odd || noR ? `<div class="panel"><div class="pb" style="padding-top:12px"><ul class="health">${odd ? `<li><span class="ic wa">!</span><span>${odd} rows have TVR share above 100%. Check the source file.</span></li>` : ''}${noR ? `<li><span class="ic in">i</span><span>${noR} rows have no Reach %. Reach-based numbers use TVR × 1.4 as a stand-in for those.</span></li>` : ''}</ul></div></div>` : ''}
  ${panel('Program airings', `${ni(rows.length)} rows · page ${S.DPAGE + 1} of ${pages}`, `<div class="tw" style="max-height:62vh;overflow:auto"><table><thead><tr>${DCOLS.map(c => `<th class="sort ${c[2] || ''}" data-sort="${c[0]}">${c[1]}${k === c[0] ? (d < 0 ? ' ↓' : ' ↑') : ''}</th>`).join('')}</tr></thead><tbody>${pg.map(r => `<tr class="click" ${xa({ ch: r.ch, p: r.p })}><td>${r.date}</td><td class="l">${r.day.slice(0, 3)}</td><td class="l">${r.s}</td><td class="l">${r.e}</td><td class="l">${dot(r.ch)}${esc(chName(r.ch))}</td><td class="l">${esc(pn(r.p))}</td><td class="l">${esc(pn(r.cat))}</td><td>${ni(r.dur)}</td><td>${nf(r.tvr, 2)}</td><td>${nf(r.sh, 1)}</td><td>${nf(r.rp, 2)}</td></tr>`).join('')}</tbody></table></div>
    <div class="inl" style="justify-content:flex-end;margin-top:8px"><button class="btn sm" data-pg="-1" ${S.DPAGE ? '' : 'disabled'}>‹ Prev</button><span class="muted">Page ${S.DPAGE + 1} / ${pages}</span><button class="btn sm" data-pg="1" ${S.DPAGE < pages - 1 ? '' : 'disabled'}>Next ›</button></div>`)}`;
}
function dataCSV(rows) {
  return toCSV(['Channel', 'Date', 'Day', 'Start', 'End', 'Program', 'Duration', 'Category', 'TVR', 'TVR Share %', 'Reach %'], rows.map(r => [r.ch, r.date, r.day, r.s, r.e, r.p, r.dur, r.cat, r.tvr, r.sh, r.rp]));
}

/* ---------- Detail drawer (drill-down from anywhere) ---------- */
const matchF = f => r => Object.keys(f).every(k => k === 'dp' ? daypart(r.h) === f[k] : DIMS[k].key(r) === f[k]);
function fTitle(f) {
  const order = ['p', 'ch', 'cat', 'day', 'dp', 'h', 'date'];
  return order.filter(k => f[k] != null).map(k => DIMS[k].name(f[k])).join(' · ') || 'All airings';
}
function openDetail(f, push = false) {
  if (!push) S.X = [];
  S.X.push(f);
  if (!S.XB || push === false) S.XB = null;
  openPane('detail'); renderDetail();
}
function renderDetail() {
  killCharts('x-');
  const f = S.X[S.X.length - 1], rows = S.F.filter(matchF(f));
  $('#crumbs').innerHTML = S.X.map((x, i) => i === S.X.length - 1 ? `<b>${esc(fTitle(x))}</b>` : `<button data-crumb="${i}">${esc(fTitle(x))}</button><span class="sep">›</span>`).join('');
  const body = $('#xbody');
  if (!rows.length) { body.innerHTML = '<p class="muted">No airings for this selection under the current filters.</p>'; return; }
  const sm = summarize(rows), g = grp(rows, () => 1).get(1), ci = ciOf(g), stdy = steadiness(ci, g.n);
  const isProg = f.p != null;
  const progKey = isProg ? (f.ch || rows[0].ch) + '||' + f.p : null;
  const dims = ['p', 'ch', 'cat', 'day', 'h', 'date'].filter(k => f[k] == null && new Set(rows.map(DIMS[k].key)).size > 1);
  if (!S.XB || !dims.includes(S.XB)) S.XB = isProg ? (dims.includes('day') ? 'day' : dims[0]) : dims[0];
  const by = S.XB;
  const totT = sm.sT || 1;
  let bt = '';
  if (by) {
    const D = DIMS[by];
    const gs = [...grp(rows, D.key).values()].sort(D.order ? (a, b) => D.order(a.k, b.k) : (a, b) => b.sT - a.sT);
    bt = `<div class="tw" style="max-height:340px;overflow:auto"><table><thead><tr><th>${D.label}</th><th>Airings</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>Share %</th><th>Steadiness</th><th>Share of viewing</th></tr></thead><tbody>${gs.slice(0, 150).map(x => `<tr class="click" ${xa(Object.assign({}, f, { [by]: x.k }))} data-push="1"><td>${by === 'ch' ? dot(x.k) : ''}${esc(D.name(x.k))}${by === 'p' ? `<span class="sub">${esc(chName(x.mxr.ch))} · ${hl(modeH(x))}</span>` : ''}</td><td>${x.n}</td><td>${nf(avgT(x), 2)}</td><td>${nf(x.mx, 1)}</td><td>${nf(avgR(x), 1)}</td><td>${nf(avgS(x), 1)}</td><td>${stdTag(ciOf(x), x.n)}</td><td><div class="mini-bar"><div class="tr"><div class="fl" style="width:${(x.sT / totT * 100).toFixed(0)}%"></div></div>${nf(x.sT / totT * 100, 1)}%</div></td></tr>`).join('')}</tbody></table></div>`;
  }
  const airs = [...rows].sort((a, b) => b.date < a.date ? -1 : b.date > a.date ? 1 : a.s < b.s ? -1 : 1);
  const inPlan = S.cur && progKey && S.cur.kept.find(x => x.key === progKey);
  body.innerHTML = `
    <div class="inl" style="margin-bottom:12px">
      <button class="btn sm pri" data-apply-x>Filter dashboard to this</button>
      ${progKey ? `<button class="btn sm" data-lock="${esc(progKey)}">${S.P.lock.includes(progKey) ? 'Unlock from plan' : 'Lock in plan'}</button><button class="btn sm" data-excl="${esc(progKey)}">${S.P.excl.includes(progKey) ? 'Restore to plan' : 'Exclude from plan'}</button>` : ''}
      ${inPlan ? `<span class="tag anchor">In plan · ${nf(inPlan.k * 100, 1)}% · ${roleTag(inPlan.role)}</span>` : ''}
      <button class="btn sm" data-ai="${esc('Tell me about ' + fTitle(f) + ': is it worth buying, and when?')}">Ask AI about this</button>
    </div>
    <div class="kpis" style="grid-template-columns:repeat(3,minmax(0,1fr))">
      ${kp('Airings', ni(sm.n), sm.progs + ' program' + (sm.progs > 1 ? 's' : '') + ' · ' + sm.chs + ' channel' + (sm.chs > 1 ? 's' : ''))}
      ${kp('Average TVR', nf(sm.avgT, 2), 'per airing')}
      ${kp('Peak TVR', nf(sm.pk.tvr, 1), esc(sm.pk.date + ' ' + sm.pk.s + (isProg ? '' : ' · ' + pn(sm.pk.p))))}
      ${kp('Average reach', nf(sm.avgR, 2) + '%', 'per airing')}
      ${kp('Steadiness', g.n >= 2 ? nf(ci, 1) : '—', stdy.label)}
      ${kp('Usual slot', hl(modeH(g)), 'most common start hour')}
    </div>
    <div class="grid g2" style="margin-top:12px">
      ${panel('TVR over time', 'average per day', '<div class="chart sm"><canvas id="x-trend"></canvas></div>')}
      ${panel('By start hour', 'average TVR', '<div class="chart sm"><canvas id="x-hour"></canvas></div>')}
    </div>
    ${by ? `<div style="margin-top:12px">${panel('Break down by', 'click a row to drill deeper', bt, `<div class="seg" data-xb>${dims.map(k => `<button data-v="${k}" class="${k === by ? 'on' : ''}">${DIMS[k].label}</button>`).join('')}</div>`)}</div>` : ''}
    <div style="margin-top:12px">${panel('Airings', `${ni(airs.length)} total${airs.length > 200 ? ', latest 200 shown' : ''}`, `<div class="tw" style="max-height:360px;overflow:auto"><table><thead><tr><th>Date</th><th class="l">Day</th><th class="l">Time</th><th class="l">Channel</th><th class="l">Program</th><th>TVR</th><th>Share %</th><th>Reach %</th></tr></thead><tbody>${airs.slice(0, 200).map(r => `<tr><td>${r.date}</td><td class="l">${r.day.slice(0, 3)}</td><td class="l">${r.s}${r.e ? '–' + r.e : ''}</td><td class="l">${dot(r.ch)}${esc(chName(r.ch))}</td><td class="l">${esc(pn(r.p))}</td><td>${nf(r.tvr, 2)}</td><td>${nf(r.sh, 1)}</td><td>${nf(r.rp, 2)}</td></tr>`).join('')}</tbody></table></div>`)}</div>`;
  const dates = [...grp(rows, r => r.date).values()].sort((a, b) => a.k < b.k ? -1 : 1);
  mkChart('x-trend', { type: 'line', data: { labels: dates.map(d => d.k.slice(5) + ' ' + d.mxr.day.slice(0, 2)), datasets: [{ label: 'Avg TVR', data: dates.map(d => +avgT(d).toFixed(2)), borderColor: css('--accent'), backgroundColor: css('--accent'), pointRadius: 2.5, tension: .25 }] }, options: { plugins: { legend: { display: false } }, scales: { x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: true } }, y: { beginAtZero: true } } } });
  const hs = [...grp(rows, r => r.h).values()].sort((a, b) => a.k - b.k);
  mkChart('x-hour', { type: 'bar', data: { labels: hs.map(h => hl(h.k)), datasets: [{ label: 'Avg TVR', data: hs.map(h => +avgT(h).toFixed(2)), backgroundColor: css('--heat'), borderRadius: 3 }] }, options: { plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } } } } });
}
const modeH = g => +Object.keys(g.hc).sort((a, b) => g.hc[b] - g.hc[a])[0];
function applyDetailFilter() {
  const f = S.X[S.X.length - 1], F = S.FS;
  if (f.ch != null) F.ch = new Set([f.ch]);
  if (f.cat != null) F.cat = new Set([f.cat]);
  if (f.day != null) F.day = new Set([f.day]);
  if (f.h != null) { F.h0 = F.h1 = +f.h; }
  if (f.dp != null) { const r = { [DPS[0]]: [5, 11], [DPS[1]]: [12, 17], [DPS[2]]: [18, 21], [DPS[3]]: [22, 23] }[f.dp]; F.h0 = r[0]; F.h1 = r[1]; }
  if (f.date != null) { F.from = F.to = f.date; }
  if (f.p != null) F.p = f.p;
  closePanes(); applyFilters(); toast('Dashboard filtered to ' + fTitle(f));
}

/* ---------- render control ---------- */
const VIEWS = { overview: vOverview, explore: vExplore, drill: vDrill, plan: vPlan, schedule: vSchedule, scen: vScen, dup: vDup, data: vData };
function refreshPlanViews() { if (S.TAB === 'plan') updatePlan(); else if (S.TAB === 'schedule') vSchedule($('#v-schedule')); }
function renderActive() {
  killCharts('c-');
  document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
  if (!S.ROWS.length) { const w = $('#v-welcome'); w.classList.add('on'); w.innerHTML = welcomeHTML(); bindDrop($('#wdrop')); return; }
  const el = $('#v-' + S.TAB); el.classList.add('on');
  if (!S.F.length && S.TAB !== 'scen') { el.innerHTML = emptyHTML; return; }
  VIEWS[S.TAB](el);
}
function setTab(t) {
  S.TAB = t;
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  renderActive(); window.scrollTo({ top: 0 });
}

/* ---------- drawers ---------- */
function openPane(id) { ['drawer', 'detail', 'ai'].forEach(x => $('#' + x).classList.toggle('on', x === id)); $('#scrim').classList.add('on'); if (id === 'ai') setTimeout(() => $('#aiQ').focus(), 50); }
function closePanes() { ['drawer', 'detail', 'ai'].forEach(x => $('#' + x).classList.remove('on')); $('#scrim').classList.remove('on'); killCharts('x-'); }
function openDrawer(tab) { S.DT = tab || S.DT; openPane('drawer'); renderDrawer(); }
function renderDrawer() {
  document.querySelectorAll('#dtabs button').forEach(b => b.classList.toggle('on', b.dataset.dt === S.DT));
  $('#resetF').style.display = S.DT === 'filters' && S.ROWS.length ? '' : 'none';
  const b = $('#dbody'), m = S.meta;
  if (S.DT === 'upload' || !S.ROWS.length) {
    b.innerHTML = `<div class="drop" id="drop"><b>Drop an Excel or CSV file here</b><p>Needed columns: Channel, Date, Start, Program, TVR.<br>Also used: Day, End, Duration, Category, TVR Share %, Reach %.</p><button class="btn pri" data-pick>Choose file</button></div>
      <div class="status ${m && m.err ? 'err' : ''}" role="status" aria-live="polite">${!m ? '<b>No data loaded yet</b>' : m.err ? esc(m.err) : `<b>${esc(m.name)}</b>${m.demo ? ' <span class="tag">demo</span>' : ''}<dl><dt>Rows loaded</dt><dd>${ni(S.ROWS.length)}${m.skip ? ' (' + ni(m.skip) + ' skipped)' : ''}</dd><dt>Date period</dt><dd>${fmtDate(S.SPAN[0])} to ${fmtDate(S.SPAN[1])}</dd><dt>Channels</dt><dd>${S.CH.length}</dd><dt>Categories</dt><dd>${S.CATS.length}</dd>${m.at ? `<dt>Loaded</dt><dd>${new Date(m.at).toLocaleString()}</dd>` : ''}</dl>`}</div>
      <p style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap"><button class="btn" data-demo>Load demo data</button>${S.ROWS.length ? '<button class="btn danger" id="clearData">Remove my data from this browser</button>' : ''}</p>
      <p class="hint">Files are read in your browser and saved only in this browser (IndexedDB). Nothing is uploaded to the server, and other users cannot see your data. Clearing your browser's site data also removes it.</p>`;
    bindDrop($('#drop'));
    return;
  }
  const hopts = []; for (let h = 0; h < 24; h++) hopts.push(h);
  const chip = (set, v, label, attr) => `<button class="chip ${set.has(v) ? 'on' : ''}" data-${attr}="${esc(v)}">${esc(label)}</button>`;
  const F = S.FS;
  b.innerHTML = `
   <div class="grp"><h4>Date period</h4><div class="two"><label><span class="sub">From</span><input type="date" id="f-from" min="${S.SPAN[0]}" max="${S.SPAN[1]}" value="${F.from}"></label><label><span class="sub">To</span><input type="date" id="f-to" min="${S.SPAN[0]}" max="${S.SPAN[1]}" value="${F.to}"></label></div></div>
   <div class="grp"><h4>Channels<span class="push"><button class="link" data-all="ch">All</button><button class="link" data-none="ch">None</button></span></h4><div class="chips">${S.CH.map(c => chip(F.ch, c, chName(c), 'ch')).join('')}</div></div>
   <div class="grp"><h4>Days<span class="push"><button class="link" data-days="wd">Weekdays</button><button class="link" data-days="we">Weekend</button><button class="link" data-all="day">All</button></span></h4><div class="chips">${DAYS.map((d, i) => chip(F.day, d, DS[i], 'day')).join('')}</div></div>
   <div class="grp"><h4>Time band (start hour)<span class="push"><button class="link" data-preset="prime">Prime</button><button class="link" data-preset="day">Daytime</button><button class="link" data-preset="allday">All day</button></span></h4><div class="two"><label><span class="sub">From</span><select id="f-h0">${hopts.map(h => `<option value="${h}" ${h === F.h0 ? 'selected' : ''}>${hl(h)}</option>`).join('')}</select></label><label><span class="sub">To (inclusive)</span><select id="f-h1">${hopts.map(h => `<option value="${h}" ${h === F.h1 ? 'selected' : ''}>${hl(h)}</option>`).join('')}</select></label></div></div>
   <div class="grp"><h4>Categories<span class="push"><button class="link" data-all="cat">All</button><button class="link" data-none="cat">None</button></span></h4><input type="search" id="f-catq" placeholder="Search categories"><div class="catlist" id="catlist"></div></div>
   <div class="grp"><h4>Program and rating</h4><div class="two"><label><span class="sub">Program name contains</span><input type="text" id="f-q" value="${esc(F.q)}"></label><label><span class="sub">Minimum TVR</span><input type="number" id="f-min" min="0" step="0.5" value="${F.minTvr}"></label></div>${F.p ? `<p class="hint">Exact program: <b>${esc(pn(F.p))}</b> <button class="link" data-clear="p">remove</button></p>` : ''}</div>`;
  fillCats('');
}
function fillCats(q) {
  const l = $('#catlist'); if (!l) return;
  const qq = q.toLowerCase();
  l.innerHTML = S.CATS.filter(c => c.toLowerCase().includes(qq)).map(c => `<label><input type="checkbox" data-cat="${esc(c)}" ${S.FS.cat.has(c) ? 'checked' : ''}>${esc(pn(c))}</label>`).join('') || '<span class="sub">No match</span>';
}
function preset(k) {
  const F = S.FS;
  if (k === 'prime') { F.h0 = Math.max(18, S.HRS[0]); F.h1 = Math.min(21, S.HRS[1]); }
  if (k === 'day') { F.h0 = Math.max(6, S.HRS[0]); F.h1 = Math.min(17, S.HRS[1]); }
  if (k === 'allday') { F.h0 = S.HRS[0]; F.h1 = S.HRS[1]; }
  if (k === 'wd') F.day = new Set(DAYS.slice(0, 5));
  if (k === 'we') F.day = new Set(DAYS.slice(5));
  applyFilters(); if ($('#drawer').classList.contains('on')) renderDrawer();
}

/* ---------- upload ---------- */
let fileInput;
function pickFile() {
  if (!fileInput) { fileInput = document.createElement('input'); fileInput.type = 'file'; fileInput.accept = '.xlsx,.xls,.csv'; fileInput.onchange = () => fileInput.files[0] && handleFile(fileInput.files[0]); }
  fileInput.value = ''; fileInput.click();
}
function bindDrop(drop) {
  if (!drop) return;
  ['dragover', 'dragenter'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', ev => { const f = ev.dataTransfer.files[0]; if (f) handleFile(f); });
}
async function handleFile(file) {
  toast('Reading ' + file.name + '…');
  try {
    if (typeof XLSX === 'undefined') throw new Error('The spreadsheet reader did not load. Check your connection and try again.');
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    // Use the first sheet that has a Channel/Program header.
    let json = [];
    for (const n of wb.SheetNames) { json = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '', raw: true }); if (json.length && Object.keys(json[0]).some(k => /channel/i.test(k))) break; }
    const { rows, skip } = parseRows(json);
    setData(rows, { name: file.name, skip, at: Date.now(), demo: false });
    toast(ni(rows.length) + ' airings loaded');
    if ($('#drawer').classList.contains('on')) renderDrawer();
  } catch (err) {
    S.meta = Object.assign({}, S.meta, { err: 'Could not load the file. ' + err.message });
    openDrawer('upload'); renderDrawer();
    S.meta = S.ROWS.length ? Object.assign({}, S.meta, { err: null }) : null;
  }
}
function loadDemo() { setData(makeDemoRows(), { name: 'Demo data (synthetic, Sep 2026)', skip: 0, at: Date.now(), demo: true }); toast('Demo data loaded'); if ($('#drawer').classList.contains('on')) renderDrawer(); }

/* ---------- Planner AI ---------- */
function renderAICtx() {
  if (!S.ROWS.length) { $('#aiCtx').textContent = 'Load data first. The AI answers from your filtered numbers.'; return; }
  const f = S.FS;
  $('#aiCtx').textContent = `Using: ${f.ch.size === S.CH.length ? 'all' : f.ch.size} channels · ${f.day.size === 7 ? 'all days' : f.day.size + ' days'} · ${hl(f.h0)}–${hl((f.h1 + 1) % 24)} · ${fmtDate(f.from)} to ${fmtDate(f.to)} · ${ni(S.F.length)} airings${S.cur ? ' · plan net reach ' + nf(S.cur.net, 1) + '%' : ''}`;
}
function mdLite(t) {
  const lines = esc(t).split('\n'); let html = '', inList = false;
  for (let l of lines) {
    l = l.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/(^|\s)\*(\S.+?)\*(?=\s|$)/g, '$1<i>$2</i>');
    const m = l.match(/^\s*(?:[-*•]|\d+[.)])\s+(.*)/);
    if (m) { if (!inList) { html += '<ul>'; inList = true; } html += `<li>${m[1]}</li>`; continue; }
    if (inList) { html += '</ul>'; inList = false; }
    if (/^#{1,4}\s/.test(l)) html += `<p><b>${l.replace(/^#+\s/, '')}</b></p>`;
    else if (l.trim()) html += `<p>${l}</p>`;
  }
  return html + (inList ? '</ul>' : '');
}
function renderChat() {
  const log = $('#aiLog');
  log.innerHTML = S.chat.map((m, i) => m.role === 'user' ? `<div class="q">${esc(m.text)}</div>` :
    `<div class="a ${m.err ? 'err' : ''}">${m.html}${m.action ? `<div class="acts"><button class="btn sm pri" data-act="${i}">Apply to plan</button><button class="btn sm" data-act-save="${i}">Save as scenario</button></div>` : ''}<span class="src">${esc(m.src || '')}</span></div>`).join('') +
    (S.busy ? '<div class="a typing">Thinking…</div>' : '');
  log.lastElementChild && log.lastElementChild.scrollIntoView({ block: 'end' });
}
async function ask(q) {
  q = q.trim(); if (!q || S.busy) return;
  if (!S.ROWS.length) { toast('Load data first'); return; }
  openPane('ai');
  S.chat.push({ role: 'user', text: q }); S.busy = true; renderChat();
  let msg;
  if (S.aiOn) {
    try {
      const hist = S.chat.slice(-9, -1).map(m => ({ role: m.role, text: m.text || '' }));
      const txt = await askRemote(q, buildContext(S), hist);
      const { text, action } = splitAction(txt);
      msg = { role: 'ai', text, html: mdLite(text), action: cleanAction(action), src: 'Gemini · ' + (S.aiModel || '') };
    } catch (e) {
      if (e.status === 501) { S.aiOn = false; setAIMode(); }
      const loc = askLocal(q, S, simulate);
      msg = { role: 'ai', text: loc.html.replace(/<[^>]+>/g, ' '), html: loc.html, action: cleanAction(loc.action), src: 'Offline planner (AI unavailable: ' + e.message + ')' };
    }
  } else {
    const loc = askLocal(q, S, simulate);
    msg = { role: 'ai', text: loc.html.replace(/<[^>]+>/g, ' '), html: loc.html, action: cleanAction(loc.action), src: 'Offline planner' };
  }
  S.busy = false; S.chat.push(msg); renderChat();
}
// Only allow known, valid keys from AI actions.
function cleanAction(a) {
  if (!a || typeof a !== 'object') return null;
  const o = {};
  if (a.cut != null && isFinite(a.cut)) o.cut = Math.max(0, Math.min(50, Math.round(+a.cut / 5) * 5));
  if (a.nCh != null && isFinite(a.nCh)) o.nCh = Math.max(1, Math.min(S.CH.length, Math.round(+a.nCh)));
  if (a.nProg != null && isFinite(a.nProg)) o.nProg = Math.max(1, Math.min(8, Math.round(+a.nProg)));
  if (a.budget != null && isFinite(a.budget) && +a.budget > 0) o.budget = Math.round(+a.budget);
  if (Array.isArray(a.channels)) { const c = a.channels.map(x => String(x).toUpperCase()).filter(x => S.CH.includes(x)); if (c.length) o.channels = c; }
  if (a.h0 != null && isFinite(a.h0)) o.h0 = Math.max(0, Math.min(23, Math.round(+a.h0)));
  if (a.h1 != null && isFinite(a.h1)) o.h1 = Math.max(0, Math.min(23, Math.round(+a.h1)));
  if (Array.isArray(a.days)) { const d = a.days.map(x => DAYS.find(D => D.toLowerCase().startsWith(String(x).toLowerCase().slice(0, 3)))).filter(Boolean); if (d.length) o.days = d; }
  return Object.keys(o).length ? o : null;
}
function actionParts(a) {
  const p = {}, fs = {};
  ['cut', 'nCh', 'nProg', 'budget'].forEach(k => { if (a[k] != null) p[k] = a[k]; });
  if (a.nCh != null) { p.chMode = 'auto'; p.split = {}; }
  if (a.channels) { p.chMode = 'manual'; p.chPick = a.channels; p.split = {}; fs.ch = new Set([...S.FS.ch, ...a.channels]); }
  if (a.h0 != null) fs.h0 = a.h0; if (a.h1 != null) fs.h1 = a.h1;
  if (fs.h0 != null && fs.h1 == null) fs.h1 = Math.max(fs.h0, S.FS.h1);
  if (fs.h1 != null && fs.h0 == null) fs.h0 = Math.min(fs.h1, S.FS.h0);
  if (a.days) fs.day = new Set(a.days);
  return { p, fs };
}
function applyAction(a) {
  const { p, fs } = actionParts(a);
  Object.assign(S.P, p); Object.assign(S.FS, fs); savePlan();
  applyFilters(); closePanes(); setTab('plan'); toast('Applied to plan');
}
function saveActionScenario(a, label) {
  const { p, fs } = actionParts(a);
  const P2 = Object.assign({}, S.P, p), FS2 = Object.assign({}, S.FS, fs);
  const pl = planCalc(filterRows(FS2), P2);
  if (!pl || !pl.items.length) { toast('That scenario has no programs'); return; }
  saveScenario(label, scen(pl, 1 - P2.cut / 100, P2, getD, S.DINTRA), P2, FS2);
}
function setAIMode() { const t = $('#aiMode'); t.textContent = S.aiOn ? 'Gemini' : 'Offline planner'; t.className = 'tag ' + (S.aiOn ? 'anchor' : ''); t.title = S.aiOn ? 'Model: ' + S.aiModel : 'Server has no GEMINI_API_KEY. Rule-based answers from your data.'; }

function exportDeck() {
  if (!S.cur) { toast('Nothing to export: the plan is empty'); return; }
  const m = lsGet('deckMeta', {});
  $('#dk-client').value = m.client || ''; $('#dk-campaign').value = m.campaign || ''; $('#dk-by').value = m.by || '';
  $('#dk-app').checked = m.appendix !== false; $('#dk-ai').checked = m.useAI !== false;
  $('#dk-ai-row').style.display = S.aiOn ? '' : 'none';
  $('#deckModal').hidden = false; setTimeout(() => $('#dk-client').focus(), 30);
}
async function runDeck() {
  const o = { client: $('#dk-client').value, campaign: $('#dk-campaign').value, by: $('#dk-by').value, appendix: $('#dk-app').checked, useAI: S.aiOn && $('#dk-ai').checked };
  lsSet('deckMeta', o);
  const go = $('#dk-go'); go.disabled = true; go.textContent = o.useAI ? 'Writing commentary…' : 'Building…';
  try { const r = await buildDeck(S, simulate, o); $('#deckModal').hidden = true; toast('Deck downloaded' + (r.ai ? ' (AI commentary)' : '')); }
  catch (e) { toast('Deck export failed: ' + e.message); }
  finally { go.disabled = false; go.textContent = 'Create deck'; }
}

const PLAN_EDIT = '#v-plan .settings, #v-plan [data-preset], [data-lock], [data-excl], [data-unexcl], #v-plan [data-split], #pl-split-reset, [data-cr-add], [data-cr-del], [data-act]';
function unfreezeIf(t) { if (S.P.frozen && t && t.closest && t.closest(PLAN_EDIT)) { S.P.frozen = null; savePlan(); document.querySelectorAll('#v-plan .frozen').forEach(e => e.remove()); } }
['click', 'change', 'input'].forEach(ev => document.addEventListener(ev, e => unfreezeIf(e.target), true));

/* ---------- events ---------- */
$('#tabs').addEventListener('click', e => { const b = e.target.closest('button'); if (b && S.ROWS.length) setTab(b.dataset.tab); });
$('#filterBtn').onclick = () => openDrawer(S.ROWS.length ? 'filters' : 'upload');
$('#periodBtn').onclick = () => openDrawer(S.ROWS.length ? 'filters' : 'upload');
$('#dataBtn').onclick = () => openDrawer('upload');
$('#aiBtn').onclick = () => { openPane('ai'); renderChat(); };
$('#closeD').onclick = closePanes; $('#doneD').onclick = closePanes; $('#scrim').onclick = closePanes;
$('#closeX').onclick = closePanes; $('#closeAI').onclick = closePanes;
$('#resetF').onclick = () => resetFilters();
$('#themeBtn').onclick = () => {
  const dark = document.documentElement.dataset.theme ? document.documentElement.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
  document.documentElement.dataset.theme = dark ? 'light' : 'dark'; lsSet('theme', document.documentElement.dataset.theme);
};
$('#dtabs').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { S.DT = b.dataset.dt; renderDrawer(); } });
$('#deckForm').addEventListener('submit', e => { e.preventDefault(); runDeck(); });
$('#aiForm').addEventListener('submit', e => { e.preventDefault(); const q = $('#aiQ').value; $('#aiQ').value = ''; ask(q); });
$('#aiSugg').addEventListener('click', e => { const b = e.target.closest('button'); if (b) ask(b.textContent); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') { closePanes(); $('#deckModal').hidden = true; } });

document.addEventListener('click', e => {
  const t = e.target;
  let el;
  if ((el = t.closest('[data-pick]'))) { pickFile(); return; }
  if (t.closest('[data-cr-add]')) { S.P.creatives.push({ name: 'Creative ' + String.fromCharCode(65 + S.P.creatives.length), dur: 10, share: 0 }); savePlan(); vPlan($('#v-plan')); return; }
  if ((el = t.closest('[data-cr-del]'))) { S.P.creatives.splice(+el.dataset.crDel, 1); S.P.spotLen = Math.round(avgLen(S.P)); savePlan(); recalc(); vPlan($('#v-plan')); return; }
  if ((el = t.closest('[data-sch-ch]'))) { S.SCH.ch = el.dataset.schCh; vSchedule($('#v-schedule')); return; }
  if (t.closest('#sch-xlsx')) { runXlsx(t.closest('#sch-xlsx')); return; }
  if (t.closest('[data-reopt]')) { S.P.frozen = null; savePlan(); recalc(); refreshPlanViews(); toast('Re-optimised with the new rates'); return; }
  if ((el = t.closest('[data-bkt]'))) { S.BK.tier = el.dataset.bkt; renderBasket(); return; }
  if ((el = t.closest('[data-bkopen]'))) { const c = el.dataset.bkopen; S.BK.open.has(c) ? S.BK.open.delete(c) : S.BK.open.add(c); renderBasket(); return; }
  if (t.closest('[data-flall]')) { S.BK.flAll = !S.BK.flAll; renderFlight(); return; }
  if (t.closest('[data-bkall]')) { S.BK.all = !S.BK.all; renderBasket(); return; }
  if ((el = t.closest('[data-deck]'))) { exportDeck(); return; }
  if (t.closest('#dk-cancel') || t.id === 'deckModal') { $('#deckModal').hidden = true; return; }
  if ((el = t.closest('[data-demo]'))) { loadDemo(); return; }
  if (t.closest('#clearData')) {
    if (!confirm('Remove the loaded data, scenarios and plan settings from this browser?')) return;
    idbDel('dataset'); S.scenarios = []; lsSet('scenarios', []); S.P = Object.assign({}, DEFAULT_P); savePlan();
    S.ROWS = []; S.F = []; S.meta = null; S.chat = []; closePanes(); chrome(); renderActive(); toast('Data removed from this browser'); return;
  }
  if ((el = t.closest('[data-open]'))) { openDrawer(el.dataset.open); return; }
  if ((el = t.closest('[data-clear]'))) { clearFilter(el.dataset.clear); return; }
  if ((el = t.closest('[data-tab-go]'))) { setTab(el.dataset.tabGo); return; }
  if ((el = t.closest('[data-ai]'))) { ask(el.dataset.ai); return; }
  if ((el = t.closest('[data-crumb]'))) { S.X = S.X.slice(0, +el.dataset.crumb + 1); S.XB = null; renderDetail(); return; }
  if ((el = t.closest('[data-xb] button'))) { S.XB = el.dataset.v; renderDetail(); return; }
  if (t.closest('[data-apply-x]')) { applyDetailFilter(); return; }
  if ((el = t.closest('[data-lock]'))) {
    const k = el.dataset.lock, L = S.P.lock;
    L.includes(k) ? L.splice(L.indexOf(k), 1) : (L.push(k), S.P.excl = S.P.excl.filter(x => x !== k));
    savePlan(); recalc(); if (S.TAB === 'plan') updatePlan(); if ($('#detail').classList.contains('on')) renderDetail();
    toast(L.includes(k) ? 'Locked in plan' : 'Unlocked'); return;
  }
  if ((el = t.closest('[data-excl]'))) {
    const k = el.dataset.excl, X = S.P.excl;
    X.includes(k) ? X.splice(X.indexOf(k), 1) : (X.push(k), S.P.lock = S.P.lock.filter(x => x !== k));
    savePlan(); recalc(); if (S.TAB === 'plan') updatePlan(); if ($('#detail').classList.contains('on')) renderDetail();
    toast(X.includes(k) ? 'Removed from plan' : 'Restored to plan'); return;
  }
  if ((el = t.closest('[data-unexcl]'))) { S.P.excl = S.P.excl.filter(x => x !== el.dataset.unexcl); savePlan(); recalc(); updatePlan(); return; }
  if ((el = t.closest('[data-x]'))) {
    const f = JSON.parse(el.dataset.x);
    if (el.closest('#detail')) { S.X.push(f); S.XB = null; renderDetail(); $('#xbody').scrollTop = 0; } else openDetail(f);
    return;
  }
  if ((el = t.closest('[data-tg]'))) { const p = el.dataset.tg; S.DR.open.has(p) ? S.DR.open.delete(p) : S.DR.open.add(p); vDrill($('#v-drill')); return; }
  if (t.closest('#dr-col')) { S.DR.open.clear(); vDrill($('#v-drill')); return; }
  if (t.closest('#dr-exp')) { const D = DIMS[S.DR.levels[0]]; new Set(S.F.map(D.key)).forEach(k => S.DR.open.add('¦' + k)); vDrill($('#v-drill')); return; }
  if ((el = t.closest('[data-seg] button'))) {
    const k = el.parentElement.dataset.seg;
    if (['strategy', 'pacing', 'same'].includes(k)) {
      S.P[k] = el.dataset.v; if (k === 'strategy') S.P.tiers = [...STRATS[el.dataset.v].t];
      savePlan(); recalc(); vPlan($('#v-plan')); return;
    }
    if (k === 'sdDays' || k === 'sdScope') { S.SD[k] = el.dataset.v; killCharts('c-'); vDup($('#v-dup')); return; }
    if (k === 'view') { S.BK.view = el.dataset.v; S.BK.all = false; renderBasket(); return; }
    if (k === 'fv') { S.fv = el.dataset.v; el.parentElement.querySelectorAll('button').forEach(b => b.classList.toggle('on', b === el)); renderFlight(); return; }
    if (k === 'chMode') { S.P.chMode = el.dataset.v; S.P.split = {}; if (S.P.chMode === 'manual' && !S.P.chPick.length && S.cur) S.P.chPick = S.cur.chs.map(c => c.ch); savePlan(); recalc(); vPlan($('#v-plan')); return; }
    S.EX[k] = el.dataset.v; killCharts('c-'); vExplore($('#v-explore')); return;
  }
  if ((el = t.closest('[data-sfch]'))) { const c = el.dataset.sfch; S.SF.ch.has(c) ? S.SF.ch.delete(c) : S.SF.ch.add(c); killCharts('c-'); vExplore($('#v-explore')); return; }
  if ((el = t.closest('[data-pick-ch]'))) {
    const c = el.dataset.pickCh, a = S.P.chPick; a.includes(c) ? a.splice(a.indexOf(c), 1) : a.push(c); delete S.P.split[c];
    savePlan(); recalc(); el.classList.toggle('on'); updatePlan(); return;
  }
  if ((el = t.closest('[data-preset]'))) { preset(el.dataset.preset); return; }
  if ((el = t.closest('[data-days]'))) { preset(el.dataset.days); return; }
  if (t.closest('#pl-adj')) { S.editSplit = !S.editSplit; t.closest('#pl-adj').textContent = S.editSplit ? 'Done' : 'Fix channel shares'; updatePlan(); return; }
  if (t.closest('#fl-csv')) { if (S.sched) download('weekly-schedule-' + S.sched.start + '.csv', scheduleCSV()); return; }
  if (t.closest('#pl-split-reset')) { S.P.split = {}; savePlan(); recalc(); updatePlan(); return; }
  if (t.closest('#p-reset')) { const keep = S.P.budget; S.P = Object.assign({}, DEFAULT_P, { budget: keep }); savePlan(); recalc(); vPlan($('#v-plan')); return; }
  if (t.closest('#pl-save') || t.closest('#sc-save')) { const n = prompt('Name this scenario', 'Scenario ' + String.fromCharCode(65 + S.scenarios.length % 26)); if (n !== null) { saveScenario(n, S.cur); if (S.TAB === 'scen') vScen($('#v-scen')); } return; }
  if (t.closest('#pl-csv') || t.closest('#ex-csv')) { const c = curAsScen(); if (c) download('spot-plan-' + new Date().toISOString().slice(0, 10) + '.csv', planCSV(c)); return; }
  if (t.closest('#pl-copy') || t.closest('#ex-copy')) {
    const txt = S.TAB === 'plan' ? memoText() : (() => { const d = document.createElement('div'); d.innerHTML = memo(S.cur, S.base); return d.innerText; })();
    navigator.clipboard.writeText(txt).then(() => toast('Memo copied'), () => toast('Copy failed')); return;
  }
  if (t.closest('#ex-print')) { setTab('plan'); setTimeout(() => window.print(), 300); return; }
  if (t.closest('#ex-data') || t.closest('#ex-data2')) { download('tv-data-filtered.csv', dataCSV(S.F)); return; }
  if ((el = t.closest('[data-quick]'))) { const c = +el.dataset.quick; saveScenario('Budget −' + c + '%', simulate({ cut: c }), Object.assign({}, S.P, { cut: c })); vScen($('#v-scen')); return; }
  if ((el = t.closest('[data-sc-load]'))) { loadScenario(S.scenarios.find(s => s.id === el.dataset.scLoad)); return; }
  if ((el = t.closest('[data-sc-cmp]'))) { S.cmp = S.cmp === el.dataset.scCmp ? null : el.dataset.scCmp; vScen($('#v-scen')); return; }
  if ((el = t.closest('[data-sc-csv]'))) { const s = S.scenarios.find(x => x.id === el.dataset.scCsv); download('plan-' + s.name.replace(/\W+/g, '-') + '.csv', planCSV(s)); return; }
  if ((el = t.closest('[data-sc-del]'))) { S.scenarios = S.scenarios.filter(s => s.id !== el.dataset.scDel); lsSet('scenarios', S.scenarios); vScen($('#v-scen')); return; }
  if ((el = t.closest('[data-act]'))) { applyAction(S.chat[+el.dataset.act].action); return; }
  if ((el = t.closest('[data-act-save]'))) { const m = S.chat[+el.dataset.actSave]; const q = S.chat[+el.dataset.actSave - 1]; saveActionScenario(m.action, 'AI: ' + (q ? q.text.slice(0, 30) : 'what-if')); return; }
  if ((el = t.closest('th[data-sort]'))) { const k = el.dataset.sort; S.DSORT = S.DSORT.k === k ? { k, d: -S.DSORT.d } : { k, d: -1 }; vData($('#v-data')); return; }
  if ((el = t.closest('[data-pg]'))) { S.DPAGE += +el.dataset.pg; vData($('#v-data')); return; }
  if (t.closest('#dup-reset')) { S.DUP = {}; lsSet('dup', S.DUP); recalc(); killCharts('c-'); vDup($('#v-dup')); return; }
  if (t.closest('#dbody')) {
    const ch = t.closest('.chip');
    if (ch) {
      const a = ch.dataset, set = a.ch !== undefined ? S.FS.ch : S.FS.day, v = a.ch !== undefined ? a.ch : a.day;
      set.has(v) ? set.delete(v) : set.add(v); ch.classList.toggle('on'); applyFilters(); return;
    }
    const al = t.closest('[data-all],[data-none]');
    if (al) { const key = al.dataset.all || al.dataset.none, all = al.dataset.all !== undefined; S.FS[key] = new Set(all ? (key === 'ch' ? S.CH : key === 'day' ? DAYS : S.CATS) : []); applyFilters(); renderDrawer(); return; }
  }
});
document.addEventListener('change', e => {
  const t = e.target, F = S.FS;
  if (t.matches('#catlist input')) { const v = t.dataset.cat; t.checked ? F.cat.add(v) : F.cat.delete(v); applyFilters(); return; }
  if (t.id === 'f-from') { F.from = t.value || S.SPAN[0]; if (F.from > F.to) F.to = F.from; applyFilters(); renderDrawer(); return; }
  if (t.id === 'f-to') { F.to = t.value || S.SPAN[1]; if (F.to < F.from) F.from = F.to; applyFilters(); renderDrawer(); return; }
  if (t.id === 'f-h0') { F.h0 = +t.value; if (F.h0 > F.h1) F.h1 = F.h0; applyFilters(); renderDrawer(); return; }
  if (t.id === 'f-h1') { F.h1 = +t.value; if (F.h1 < F.h0) F.h0 = F.h1; applyFilters(); renderDrawer(); return; }
  if (t.id === 'f-min') { F.minTvr = Math.max(0, +t.value || 0); applyFilters(); return; }
  if (t.id === 'sd-h') { S.SD.h = +t.value; killCharts('c-'); vDup($('#v-dup')); return; }
  if (t.id === 'sf-h') { S.SF.h = +t.value; killCharts('c-'); vExplore($('#v-explore')); return; }
  if (t.dataset.lvl !== undefined) {
    const i = +t.dataset.lvl, L = S.DR.levels.slice(0, i); if (t.value) L.push(t.value);
    if (i < S.DR.levels.length - 1 && t.value) S.DR.levels.slice(i + 1).forEach(k => { if (!L.includes(k)) L.push(k); });
    S.DR.levels = L.slice(0, 3); S.DR.open.clear(); vDrill($('#v-drill')); return;
  }
  if (t.dataset.split !== undefined) { S.P.split[t.dataset.split] = Math.max(0, Math.min(100, +t.value || 0)); savePlan(); recalc(); updatePlan(); return; }
  if (t.dataset.p !== undefined && t.type !== 'number') {
    S.P[t.dataset.p] = t.tagName === 'SELECT' ? +t.value : t.value;
    if (S.P.end && S.P.end < S.P.start) S.P.end = S.P.start;
    savePlan(); recalc(); if (t.type === 'date' && S.TAB === 'plan') vPlan($('#v-plan')); else refreshPlanViews(); return;
  }
  if (t.dataset.cr !== undefined) {
    const c = S.P.creatives[+t.dataset.cr], f = t.dataset.f; c[f] = f === 'name' ? t.value.trim() || 'Creative' : Math.max(0, +t.value || 0);
    S.P.spotLen = Math.round(avgLen(S.P)); savePlan(); recalc(); vPlan($('#v-plan')); return;
  }
  if ((t.dataset.rate !== undefined || t.dataset.disc !== undefined) && !S.P.frozen && S.cur) S.P.frozen = Object.fromEntries(S.cur.kept.map(x => [x.key, x.spots]));
  if (t.dataset.rate !== undefined) { const v = +String(t.value).replace(/[^\d.]/g, ''); if (v > 0) S.P.rates[t.dataset.rate] = v; else delete S.P.rates[t.dataset.rate]; savePlan(); recalc(); refreshPlanViews(); return; }
  if (t.dataset.disc !== undefined) { S.P.disc[t.dataset.disc] = Math.min(95, Math.max(0, +t.value || 0)); savePlan(); recalc(); refreshPlanViews(); return; }
  if (t.dataset.meta !== undefined) { const m = lsGet('deckMeta', {}); m[t.dataset.meta] = t.value; lsSet('deckMeta', m); return; }
  if (t.id === 'd-rep') { S.P.repQ = Math.min(90, Math.max(0, +t.value || 0)); savePlan(); recalc(); return; }
  if (t.dataset && t.dataset.a !== undefined && t.closest('.mx')) {
    S.DUP[dkey(t.dataset.a, t.dataset.b)] = Math.min(1, Math.max(0, +t.value || 0)); lsSet('dup', S.DUP); recalc(); killCharts('c-'); vDup($('#v-dup')); return;
  }
  if (t.id === 'd-intra') { S.DINTRA = Math.min(1, Math.max(0, +t.value || 0)); lsSet('dintra', S.DINTRA); recalc(); killCharts('c-'); vDup($('#v-dup')); return; }
});
let qTimer, pTimer;
document.addEventListener('input', e => {
  const t = e.target, P = S.P;
  if (t.id === 'f-q') { S.FS.q = t.value; clearTimeout(qTimer); qTimer = setTimeout(applyFilters, 250); return; }
  if (t.id === 'f-catq') { fillCats(t.value); return; }
  if (t.id === 'dq') { S.DQ = t.value; S.DPAGE = 0; clearTimeout(qTimer); qTimer = setTimeout(() => { vData($('#v-data')); const i = $('#dq'); i.focus(); i.setSelectionRange(i.value.length, i.value.length); }, 250); return; }
  const upd = () => { savePlan(); clearTimeout(pTimer); pTimer = setTimeout(() => { recalc(); refreshPlanViews(); }, 150); };
  if (t.dataset.p !== undefined && t.type === 'number') { const v = +t.value; if (t.value !== '' && isFinite(v)) { P[t.dataset.p] = v; upd(); } return; }
  if (t.dataset.pa !== undefined) {
    const v = Math.max(0, Math.min(100, +t.value || 0)); P[t.dataset.pa][+t.dataset.i] = v;
    if (t.dataset.pa === 'tiers') { $('#l-tsum').textContent = P.tiers.reduce((a, b) => a + (+b || 0), 0) + '%'; const m = Object.keys(STRATS).find(k => STRATS[k].t.every((x, i) => x === P.tiers[i])); P.strategy = m || 'custom'; }
    upd(); return;
  }
  if (t.id === 'p-cap') { P.capWk = +t.value; $('#l-cap').textContent = P.capWk; upd(); return; }
  if (t.id === 'p-bud') { P.budget = Math.max(0, +t.value || 0); upd(); return; }
  if (t.id === 'p-cprp') { P.cprp = Math.max(0, +t.value || 0); upd(); return; }
  if (t.id === 'p-tgt') { P.target = +t.value; t.previousElementSibling.querySelector('b').textContent = P.target ? P.target + '%' : 'none'; upd(); return; }
  if (t.id === 'p-nch') { P.nCh = +t.value; P.split = {}; $('#l-nch').textContent = P.nCh; upd(); return; }
  if (t.id === 'p-np') { P.nProg = +t.value; $('#l-np').textContent = P.nProg; upd(); return; }
  if (t.id === 'p-cut') { P.cut = +t.value; $('#l-cut').textContent = P.cut ? '−' + P.cut + '%' : 'none'; upd(); return; }
});
let rT; window.addEventListener('resize', () => { clearTimeout(rT); rT = setTimeout(() => Object.values(charts).forEach(c => c.resize()), 120); });
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => renderActive());
new MutationObserver(() => { renderActive(); if ($('#detail').classList.contains('on')) renderDetail(); }).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

/* ---------- boot ---------- */
(async function boot() {
  const th = lsGet('theme', null); if (th) document.documentElement.dataset.theme = th;
  fetch('/api/config').then(r => r.json()).then(c => { S.aiOn = !!c.ai; S.aiModel = c.model; }).catch(() => { S.aiOn = false; }).finally(setAIMode);
  const saved = await idbGet('dataset');
  if (saved && saved.rows && saved.rows.length) setData(saved.rows, saved.meta, false);
  else { chrome(); renderActive(); renderAICtx(); }
})();
