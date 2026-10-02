import {
  DAYS, DS, DIMS, DPS, daypart, esc, nf, ni, hl, band, chName, pn, fmtDate, dayDiff, lkr,
  parseRows, grp, avgT, avgR, avgS, ciOf, steadiness, summarize, mkGetD, dkey, netReach, uni,
  channelStats, planCalc, scen, competing, healthChecks, insights, toCSV, STRATS, buildSchedule
} from './engine.js';
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
  weeks: 4, start: nextMonday(), pacing: 'even', same: 'roadblock', repQ: 40
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
  fv: 'weeks', sched: null,
  chat: [], aiOn: false, aiModel: null, busy: false, editSplit: false
};
const getD = (a, b) => mkGetD(S.DUP)(a, b);
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
  const P2 = Object.assign({}, S.P, patch);
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
    ${kp('Total GRPs', ni(sm.sT), 'sum of TVR')}
  </div>
  <div class="grid g-ov">
    ${panel('Channel performance', 'average per airing · click a bar', '<div class="chart lg"><canvas id="c-ch"></canvas></div>')}
    ${panel('Top programs', 'by average TVR, min 2 airings · click a row', `<div class="tw" style="max-height:340px;overflow:auto"><table><thead><tr><th>Program</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>Airings</th><th>Steadiness</th></tr></thead><tbody>${q.map(g => { const [c, p] = g.k.split('||'); return `<tr class="click" ${xa({ ch: c, p })}><td><div class="pn" title="${esc(pn(p))}">${esc(pn(p))}</div><span class="sub">${dot(c)}${esc(chName(c))} · ${hl(g.mxr.h)}</span></td><td>${nf(avgT(g), 2)}</td><td>${nf(g.mx, 1)}</td><td>${nf(avgR(g), 1)}</td><td>${g.n}</td><td>${stdTag(ciOf(g), g.n)}</td></tr>`; }).join('')}</tbody></table></div>`)}
    ${panel('Category mix', 'share of total GRPs · click a slice', '<div class="chart lg"><canvas id="c-cat"></canvas></div>')}
  </div>
  <div class="grid g-wide">
    ${panel('Daily trend', 'GRPs and average TVR per day · click a day', '<div class="chart"><canvas id="c-trend"></canvas></div>')}
    ${panel('Dayparts', 'where the GRPs come from · click a row', `<table><thead><tr><th>Daypart</th><th>Airings</th><th>Avg TVR</th><th>Reach %</th><th>GRP share</th></tr></thead><tbody>${dps.map(g => `<tr class="click" ${xa({ dp: g.k })}><td>${esc(g.k)}</td><td>${ni(g.n)}</td><td>${nf(avgT(g), 2)}</td><td>${nf(avgR(g), 1)}</td><td><div class="mini-bar"><div class="tr"><div class="fl" style="width:${(g.sT / totG * 100).toFixed(0)}%"></div></div>${nf(g.sT / totG * 100, 0)}%</div></td></tr>`).join('')}</tbody></table>`)}
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
    options: { cutout: '58%', onClick: (e, els) => { if (els.length && els[0].index < topCat.length) openDetail({ cat: topCat[els[0].index].k }); }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } }, tooltip: { callbacks: { label: c => ' ' + c.label + ': ' + ni(c.raw) + ' GRPs (' + nf(c.raw / totG * 100, 1) + '%)' } } } }
  });
  const dates = [...grp(F, r => r.date).values()].sort((a, b) => a.k < b.k ? -1 : 1);
  mkChart('c-trend', {
    data: {
      labels: dates.map(g => fmtDate(g.k).slice(0, 6) + ' ' + g.mxr.day.slice(0, 2)), datasets: [
        { type: 'bar', label: 'GRPs', data: dates.map(g => Math.round(g.sT)), backgroundColor: dates.map(g => DAYS.indexOf(g.mxr.day) >= 5 ? css('--heat') : css('--accent')), borderRadius: 2, yAxisID: 'y' },
        { type: 'line', label: 'Avg TVR', data: dates.map(g => +avgT(g).toFixed(2)), borderColor: css('--ink'), backgroundColor: css('--ink'), pointRadius: 2, borderWidth: 1.5, tension: .25, yAxisID: 'y1' }]
    },
    options: { onClick: (e, els) => els.length && openDetail({ date: dates[els[0].index].k }), plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } } }, scales: { x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: true } }, y: { title: { display: true, text: 'GRPs' } }, y1: { position: 'right', grid: { display: false }, title: { display: true, text: 'Avg TVR' } } } }
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
    ${panel('Dayparts by channel', 'GRP share of each channel', '<div class="chart"><canvas id="c-dpc"></canvas></div>')}
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
      ar.slice(0, 40).forEach(r => { html += `<tr class="air l${Math.min(depth, 3)}"><td>${r.date} ${r.day.slice(0, 3)} · ${r.s}${r.e ? '–' + r.e : ''} · ${esc(chName(r.ch))} · ${esc(pn(r.p))}</td><td>1</td><td>${nf(r.tvr, 2)}</td><td></td><td>${nf(r.rp, 2)}</td><td>${nf(r.tvr, 1)}</td><td></td><td></td></tr>`; });
      if (ar.length > 40) html += `<tr class="air l${Math.min(depth, 3)}"><td colspan="8">… ${ar.length - 40} more airings. Open the detail view to see all.</td></tr>`;
      return;
    }
    const D = DIMS[dim];
    let gs = [...grp(rows, D.key).values()];
    gs.sort(D.order ? (a, b) => D.order(a.k, b.k) : (a, b) => b.sT - a.sT);
    const lim = depth === 0 ? 200 : 60;
    gs.slice(0, lim).forEach(g => {
      const p = path + '¦' + g.k, open = S.DR.open.has(p), f = Object.assign({}, fixed, { [dim]: g.k });
      const share = g.sT / (parentT || 1) * 100;
      html += `<tr class="l${Math.min(depth, 3)} ${depth === 0 ? 'l0' : ''}"><td><button class="tg" data-tg="${esc(p)}" aria-label="${open ? 'Collapse' : 'Expand'}">${open ? '▼' : '▶'}</button>${dim === 'ch' ? dot(g.k) : ''}<button class="nm" data-tg="${esc(p)}">${esc(D.name(g.k))}</button></td><td>${ni(g.n)}</td><td>${nf(avgT(g), 2)}</td><td>${nf(g.mx, 1)}</td><td>${nf(avgR(g), 2)}</td><td>${ni(g.sT)}</td><td><div class="mini-bar"><div class="tr"><div class="fl" style="width:${Math.min(100, share).toFixed(0)}%;background:${dim === 'ch' ? cv(g.k) : 'var(--accent)'}"></div></div>${nf(share, 1)}%</div></td><td><button class="ib" ${xa(f)} title="Open detail">Detail ›</button></td></tr>`;
      if (open) render(g.k === undefined ? [] : rows.filter(r => D.key(r) === g.k), depth + 1, p, g.sT, f);
    });
    if (gs.length > lim) html += `<tr class="air"><td colspan="8">… ${gs.length - lim} more</td></tr>`;
  };
  render(F, 0, '', tot, {});
  el.innerHTML = `
  <div class="vhead"><div><h2>Drill down</h2><p>Expand any row to break it down further. The last level opens individual airings. Share is of the parent row's GRPs.</p></div>
    <span class="push"><button class="btn sm" id="dr-exp">Expand level 1</button><button class="btn sm" id="dr-col">Collapse all</button></span></div>
  ${panel('Breakdown', '', `<div class="inl" style="margin-bottom:10px"><span class="muted">Level 1</span>${opts(0, L[0])}<span class="muted">▸ Level 2</span>${opts(1, L[1] || '')}<span class="muted">▸ Level 3</span>${opts(2, L[2] || '')}<span class="muted">▸ airings</span></div>
    <div class="tw"><table class="tree"><thead><tr><th>${L.map(k => DIMS[k].label).join(' ▸ ')}</th><th>Airings</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>GRPs</th><th>Share of parent</th><th></th></tr></thead><tbody>${html}</tbody></table></div>`)}`;
}

/* ---------- Planner ---------- */
const TIER_NAMES = { 1: 'Peak impact', 2: 'Efficiency anchors', 3: 'Frequency builders' };
const TIER_COL = { 1: 'var(--c1)', 2: 'var(--c2)', 3: 'var(--c3)' };
const tierTag = t => `<span class="tag tier t${t}">T${t}</span>`;
function vPlan(el) {
  const P = S.P, f = S.FS;
  const chOn = S.CH.filter(c => f.ch.has(c));
  const num = (k, o = {}) => `<input type="number" data-p="${k}" value="${P[k]}" ${o.min != null ? `min="${o.min}"` : ''} ${o.max != null ? `max="${o.max}"` : ''} step="${o.step || 1}">`;
  const strat = STRATS[P.strategy] ? P.strategy : '';
  el.innerHTML = `
  <div class="vhead"><div><h2>Build your plan</h2><p>Spots are bought one at a time where they add the most new reach per rupee, within your tier split, caps and daypart limits.</p></div>
    <span class="push"><button class="btn" id="pl-save">Save as scenario</button><button class="btn" id="pl-csv">Export spot plan</button><button class="btn" data-deck>Export deck (PPTX)</button><button class="btn pri" data-ai="Write a planner memo for the client explaining this plan">Ask AI to explain</button></span></div>
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
          <div class="two"><label class="fld"><span>Cost per rating point</span>${num('cprp', { min: 1000, step: 1000 })}</label>
          <label class="fld"><span>Spot length</span><select data-p="spotLen">${[10, 15, 20, 30, 45, 60].map(s => `<option ${s === P.spotLen ? 'selected' : ''} value="${s}">${s} sec</option>`).join('')}</select></label></div>
          <label class="fld"><span>Minimum spot rate (LKR)</span>${num('minRate', { min: 0, step: 1000 })}</label>
          <label class="fld"><span>Budget change <b id="l-cut">${P.cut ? '−' + P.cut + '%' : 'none'}</b></span><input type="range" id="p-cut" min="0" max="50" step="5" value="${P.cut}"></label>
          <label class="fld"><span>Target net reach <b id="l-tgt">${P.target ? P.target + '%' : 'none'}</b></span><input type="range" id="p-tgt" min="0" max="90" step="1" value="${P.target}"></label>
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
          <div class="two"><label class="fld"><span>Flight start</span><input type="date" data-p="start" value="${P.start}"></label><label class="fld"><span>Weeks</span>${num('weeks', { min: 1, max: 13 })}</label></div>
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
      <div class="panel"><div class="ph"><h3>Channel split</h3><span class="s">result of the spot-by-spot buy</span><span class="push"><button class="btn sm" id="pl-adj">${S.editSplit ? 'Done' : 'Fix channel shares'}</button></span></div><div class="pb" id="pl-split"></div></div>
      <div class="panel"><div class="ph"><h3>Program basket</h3><span class="s">Lock forces a buy, Remove takes a program out</span></div><div class="pb" id="pl-basket"></div></div>
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
    m('Channels', A.chs.length) + m('Programs', A.kept.length) + m('Spots', ni(A.spots)) + m('GRPs', ni(A.grps)) + m('Avg frequency', nf(A.freq, 1) + 'x') + m('Health', warn ? warn + (warn > 1 ? ' warnings' : ' warning') : 'OK');
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
  $('#pl-split').innerHTML = `<div class="stack">${A.chs.map(c => `<div style="width:${(c.w * 100).toFixed(2)}%;background:${cv(c.ch)}" title="${esc(chName(c.ch))} ${nf(c.w * 100, 1)}%">${c.w > .08 ? esc(chName(c.ch).replace(' TV', '')) + ' ' + nf(c.w * 100, 0) + '%' : ''}</div>`).join('')}</div>
   <div class="tw"><table><thead><tr><th>Channel</th><th>Share</th><th>Budget</th><th>Avg TVR</th><th>Avg reach</th><th>Programs</th><th>Spots</th><th>Reach</th><th class="l">Why this weight</th></tr></thead><tbody>${A.chs.map(c => {
    const st = S.plan.cs.find(x => x.ch === c.ch), sel = S.plan.sel.find(x => x.ch === c.ch);
    const rev = c.progs.some(x => x.role === 'review') && c.progs.length === 1;
    const t1 = c.progs.filter(x => x.tier === 1).length;
    const why = sel && sel.manual ? 'Fixed share' : c === A.chs[0] ? 'Most new reach per rupee' : c.progs.length === 1 ? (rev ? 'Risk: all spend on one volatile program' : 'Single program: ' + pn(c.progs[0].p)) : t1 ? `${t1} Tier 1 program${t1 > 1 ? 's' : ''}, adds reach beyond ${chName(A.chs[0].ch)}` : 'Cheaper frequency and incremental reach';
    return `<tr class="${rev ? 'warn' : ''}"><td><button class="nm ib" style="padding:0;color:var(--ink)" ${xa({ ch: c.ch })}>${dot(c.ch)}${esc(chName(c.ch))}</button></td><td>${S.editSplit ? `<input class="split-in" type="number" min="0" max="100" step="1" data-split="${esc(c.ch)}" value="${(c.w * 100).toFixed(0)}">` : `<b>${nf(c.w * 100, 1)}%</b>`}</td><td>${lkr(c.bud).replace('LKR ', '')}</td><td>${nf(st.tvr, 2)}</td><td>${nf(st.reach, 2)}%</td><td>${c.progs.length}</td><td>${ni(c.spots)}</td><td>${nf(c.R, 1)}%</td><td class="l">${esc(why)}</td></tr>`;
  }).join('')}</tbody></table></div>${S.editSplit ? `<p class="hint">Type a share to fix a channel; the optimiser fills the rest. <button class="link" id="pl-split-reset">Clear fixed shares</button></p>` : ''}`;
  // Basket
  const mxk = Math.max(...A.kept.map(x => x.k), .01);
  const row = (x, drop) => `<tr class="${drop ? 'drop' : x.role === 'review' ? 'warn' : ''}"><td><button class="nm ib" style="padding:0;color:var(--ink);font-weight:600;text-align:left" ${xa({ ch: x.ch, p: x.p })}><div class="pn" title="${esc(pn(x.p))}">${esc(pn(x.p))}</div></button><span class="sub">${dot(x.ch)}${esc(chName(x.ch))} · ${esc(pn(x.cat))}</span></td><td>${tierTag(x.tier)}</td><td>${hl(x.hour)}</td><td>${nf(x.mean, 2)}</td><td>${nf(x.rp, 1)}%</td><td class="l">${stdTag(x.ci, x.n)}</td><td>${drop ? '<span class="tag">dropped by cut</span>' : `<div class="mini-bar"><div class="tr"><div class="fl" style="width:${(x.k / mxk * 100).toFixed(0)}%"></div></div><b>${nf(x.k * 100, 1)}%</b></div>`}</td><td>${drop ? '' : lkr(x.bud).replace('LKR ', '')}</td><td>${drop ? '' : ni(x.spots) + (x.atCap ? ' <span class="tag" title="At the weekly spot cap">cap</span>' : '')}</td><td class="l">${roleTag(x.role)}</td><td><button class="ib ${x.locked ? 'on' : ''}" data-lock="${esc(x.key)}" title="${x.locked ? 'Unlock' : 'Always buy this program'}">${x.locked ? 'Locked' : 'Lock'}</button><button class="ib x" data-excl="${esc(x.key)}" title="Remove from plan">Remove</button></td></tr>`;
  const ex = S.P.excl;
  $('#pl-basket').innerHTML = `<div class="tw"><table><thead><tr><th>Program</th><th>Tier</th><th>Slot</th><th>Avg TVR</th><th>Reach</th><th class="l">Steadiness</th><th>Share of spend</th><th>LKR</th><th>Spots</th><th class="l">Role</th><th></th></tr></thead><tbody>${A.kept.map(x => row(x, false)).join('')}${(A.dropped || []).map(x => row(x, true)).join('')}</tbody></table></div>
    ${ex.length ? `<p class="hint">Removed: ${ex.map(k => `<span class="tag">${esc(pn(k.split('||')[1]))} <button class="link" data-unexcl="${esc(k)}">restore</button></span>`).join(' ')}</p>` : ''}
    <p class="hint">Steadiness = mean TVR ÷ standard deviation (7+ very steady, under 3 volatile). Cost per spot = max(minimum rate, CPRP × TVR × length/30).</p>`;
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
function renderFlight() {
  const A = S.cur, Sc = S.sched, P = S.P, el = $('#pl-flight');
  if (!el || !Sc) return;
  const pl = { even: 'even (drip)', burst: 'burst, front-loaded', pulse: 'pulse, on/off weeks' }[P.pacing];
  $('#fl-sub').textContent = `${A.W} weeks from ${fmtDate(Sc.start)} · ${pl} · max ${P.capWk}/week per program`;
  const wkDate = w => { const d = new Date(new Date(Sc.start + 'T12:00:00').getTime() + w * 7 * 864e5); return fmtDate(d.toISOString().slice(0, 10)).slice(0, 6); };
  if (S.fv === 'days') {
    const rows = Sc.rows;
    el.innerHTML = `<div class="tw" style="max-height:420px;overflow:auto"><table><thead><tr><th>Date</th><th class="l">Day</th><th class="l">Time band</th><th class="l">Channel</th><th class="l">Program</th><th>Tier</th><th>Spots</th><th>Cost</th><th>GRPs</th></tr></thead><tbody>${rows.slice(0, 400).map(r => `<tr><td>${fmtDate(r.date)}</td><td class="l">${r.day.slice(0, 3)}</td><td class="l">${band(r.hour)}</td><td class="l">${dot(r.ch)}${esc(chName(r.ch))}</td><td class="l">${esc(pn(r.p))}</td><td>${tierTag(r.tier)}</td><td>${r.spots}</td><td>${lkr(r.cost).replace('LKR ', '')}</td><td>${nf(r.grps, 1)}</td></tr>`).join('')}</tbody></table></div>
      <p class="hint">${Sc.clashes} night${Sc.clashes === 1 ? ' has' : 's have'} rival channels in the same hour (${P.same === 'roadblock' ? 'roadblock, by design' : 'kept low by staggering'}). Days follow each program's air days in the data.</p>`;
    return;
  }
  el.innerHTML = `<div class="chart sm"><canvas id="c-flight"></canvas></div>
    <div class="tw" style="max-height:360px;overflow:auto;margin-top:10px"><table><thead><tr><th>Program</th><th>Tier</th>${Sc.weeks.map((_, i) => `<th>W${i + 1}<span class="sub">${wkDate(i)}</span></th>`).join('')}<th>Total</th></tr></thead><tbody>
    ${Sc.grid.map(g => `<tr><td><div class="pn">${esc(pn(g.x.p))}</div><span class="sub">${dot(g.x.ch)}${esc(chName(g.x.ch))} · ${hl(g.x.hour)}</span></td><td>${tierTag(g.x.tier)}</td>${g.alloc.map(n => `<td style="${n ? '' : 'color:var(--muted)'}">${n || '·'}</td>`).join('')}<td><b>${g.x.spots}</b></td></tr>`).join('')}
    <tr><td><b>Spots</b></td><td></td>${Sc.weeks.map(w => `<td><b>${w.spots}</b></td>`).join('')}<td><b>${A.spots}</b></td></tr>
    <tr><td><b>Budget</b></td><td></td>${Sc.weeks.map(w => `<td>${lkr(w.bud).replace('LKR ', '')}</td>`).join('')}<td><b>${lkr(A.spent).replace('LKR ', '')}</b></td></tr>
    <tr><td><b>GRPs</b></td><td></td>${Sc.weeks.map(w => `<td>${ni(w.grps)}</td>`).join('')}<td><b>${ni(A.grps)}</b></td></tr></tbody></table></div>`;
  const chs = A.chs.map(c => c.ch);
  mkChart('c-flight', {
    type: 'bar', data: { labels: Sc.weeks.map((_, i) => 'W' + (i + 1) + ' ' + wkDate(i)), datasets: chs.map(ch => ({ label: chName(ch), data: Sc.weeks.map(w => w.byCh[ch] || 0), backgroundColor: cc(ch), borderRadius: 2 })) },
    options: { plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.raw} spots` } } }, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, title: { display: true, text: 'Spots' } } } }
  });
}
function scheduleCSV() {
  const r = S.sched.rows;
  return toCSV(['Week', 'Date', 'Day', 'Time band', 'Channel', 'Program', 'Tier', 'Spots', 'Spot length (sec)', 'Cost LKR', 'Est. GRPs'],
    r.map(x => [x.week, x.date, x.day, band(x.hour), chName(x.ch), pn(x.p), 'Tier ' + x.tier, x.spots, S.P.spotLen, Math.round(x.cost), x.grps.toFixed(1)]));
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
  L.push(`<p><b>Delivery:</b> ${ni(A.spots)} spots of ${P.spotLen} sec, ${ni(A.grps)} GRPs, net reach ${nf(A.net, 1)}%, reach 3+ ${nf(A.r3, 1)}%, average frequency ${nf(A.freq, 1)}x.</p>`);
  if (S.sched) L.push(`<p><b>Flighting:</b> ${A.W} weeks from ${fmtDate(S.sched.start)}, ${P.pacing} pacing (${S.sched.weeks.map(w => w.spots).join(' / ')} spots per week), max ${P.capWk} spots per program per week. Rival same-hour spots are ${P.same === 'roadblock' ? 'roadblocked on the same nights for reach' : 'staggered across nights for frequency'}.</p>`);
  if (P.cut > 0) L.push(`<p><b>Budget cut of ${P.cut}%:</b> net reach ${nf(base.net, 1)}% → ${nf(A.net, 1)}%, reach 3+ ${nf(base.r3, 1)}% → ${nf(A.r3, 1)}%. The cut removes repeat spots first, so frequency falls faster than reach.</p>`);
  const rv = A.kept.filter(x => x.role === 'review');
  if (rv.length) L.push(`<p><b>Watch:</b> ${rv.map(x => esc(pn(x.p))).join(', ')} ${rv.length > 1 ? 'have' : 'has'} volatile ratings.</p>`);
  return L.join('');
}
const memoText = () => $('#pl-memo') ? $('#pl-memo').innerText : '';

/* ---------- Scenarios ---------- */
function snapshot(name, A, P, FS) {
  return {
    id: Date.now().toString(36), name, created: new Date().toISOString(),
    P: JSON.parse(JSON.stringify(P)),
    FS: { from: FS.from, to: FS.to, ch: [...FS.ch], cat: [...FS.cat], day: [...FS.day], h0: FS.h0, h1: FS.h1, minTvr: FS.minTvr, q: FS.q, p: FS.p },
    r: { B: A.B, spent: A.spent, net: A.net, r3: A.r3, gross: A.gross, loss: A.loss, spots: A.spots, grps: A.grps, freq: A.freq, tiers: A.tiers.map(t => ({ t: t.t, target: t.target, actual: t.actual })), chs: A.chs.map(c => ({ ch: c.ch, w: c.w, bud: c.bud })), items: A.kept.map(x => ({ ch: x.ch, p: x.p, k: x.k, bud: x.bud, spots: x.spots, mean: x.mean, rp: x.rp, ci: x.ci, role: x.role, tier: x.tier, hour: x.hour, cat: x.cat, n: x.n, days: x.days })) }
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
      <div class="foot">${r.items.length} programs${r.spots ? ' · ' + ni(r.spots) + ' spots · ' + ni(r.grps) + ' GRPs' : ''}</div>
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
    ${row('Programs', ra.items.length, rb.items.length, v => ni(v))}${row('Channels', ra.chs.length, rb.chs.length, v => ni(v))}${ra.spots || rb.spots ? row('Est. spots', ra.spots, rb.spots, v => ni(v)) + row('Est. GRPs', ra.grps, rb.grps, v => ni(v)) : ''}
    ${row('Net reach per LKR 1M', ra.net / (ra.B / 1e6 || 1), rb.net / (rb.B / 1e6 || 1), v => nf(v, 2))}</tbody></table>`)}
    ${panel(`What changes in "${esc(b.name)}"`, 'compared with the current plan', `<ul class="health">${ch.join('')}</ul>`)}</div>`;
}
function planCSV(sc) {
  const r = sc.r;
  const rows = r.items.map(x => [chName(x.ch), pn(x.p), x.tier ? 'Tier ' + x.tier : '', pn(x.cat), band(x.hour), (x.days || []).map(d => d.slice(0, 3)).join(' '), x.n, x.mean.toFixed(2), x.rp.toFixed(2), x.ci.toFixed(1), x.role, (x.k * 100).toFixed(1), Math.round(x.bud), x.spots, Math.round(x.spots * x.mean)]);
  const head = ['Channel', 'Program', 'Tier', 'Category', 'Usual slot', 'Days aired', 'Airings in data', 'Avg TVR', 'Avg reach %', 'Steadiness', 'Role', 'Share of plan %', 'Budget LKR', 'Est. spots', 'Est. GRPs'];
  const top = [['Plan', sc.name], ['Budget LKR', Math.round(r.B)], ['Est. net reach %', r.net.toFixed(1)], ['Reach 3+ %', r.r3 != null ? r.r3.toFixed(1) : ''], ['Gross reach %', r.gross.toFixed(1)], ['Split', r.chs.map(c => chName(c.ch) + ' ' + (c.w * 100).toFixed(0) + '%').join('; ')], []];
  return toCSV(['TV Media Planner', 'export ' + new Date().toISOString().slice(0, 10)], top) + '\n' + toCSV(head, rows);
}

/* ---------- Duplication ---------- */
function vDup(el) {
  const present = new Set(S.F.map(r => r.ch)), chs = S.CH.filter(c => present.has(c)), n = chs.length;
  let g = `<div class="tw"><div class="mx" style="grid-template-columns:110px repeat(${n},minmax(50px,1fr))"><div></div>${chs.map(c => `<div class="mh" title="${esc(chName(c))}">${esc(chName(c).replace(' TV', ''))}</div>`).join('')}`;
  chs.forEach(a => {
    g += `<div class="hr">${dot(a)}${esc(chName(a))}</div>`;
    chs.forEach(b => {
      if (a === b) { g += `<input disabled value="1.00" aria-label="same channel">`; return; }
      const d = getD(a, b), edited = S.DUP[dkey(a, b)] != null;
      g += `<input type="number" step="0.05" min="0" max="1" data-a="${esc(a)}" data-b="${esc(b)}" value="${d.toFixed(2)}" style="background:color-mix(in srgb,var(--accent) ${(d * 70).toFixed(0)}%,var(--panel));${edited ? 'font-weight:700;border-color:var(--ink)' : ''}" aria-label="Duplication ${esc(chName(a))} with ${esc(chName(b))}">`;
    });
  });
  g += '</div></div>';
  el.innerHTML = `
  <div class="vhead"><div><h2>Duplication and net reach</h2><p>10% reach on channel A + 10% on channel B is not 20%: some viewers watch both. Edit the overlap assumptions if you have measured data.</p></div>
    <span class="push"><button class="btn sm" id="dup-reset">Reset to defaults</button></span></div>
  <div class="grid g2">
    <div class="col">
      ${panel('Duplication matrix', '0 = independent audiences, 1 = the smaller audience sits fully inside the larger', g + `<label class="fld" style="margin:12px 0 0;display:flex;align-items:center;gap:12px"><span style="margin:0;white-space:nowrap">Overlap between programs on the same channel</span><input type="number" id="d-intra" step="0.05" min="0" max="1" value="${S.DINTRA}" style="width:80px"></label>
        <label class="fld" style="margin:12px 0 0;display:flex;align-items:center;gap:12px"><span style="margin:0;white-space:nowrap">New reach from each repeat spot (% of the previous spot's new reach)</span><input type="number" id="d-rep" step="5" min="0" max="90" value="${S.P.repQ}" style="width:80px"></label>
        <p class="hint">Formula: Net(A ∪ B) = A + B − overlap, where overlap = A×B/100 + d × (min(A,B) − A×B/100).</p>`)}
    </div>
    <div class="col">
      ${panel('Gross vs net reach', 'channels added in order of strength · average reach per airing', '<div class="chart"><canvas id="c-net"></canvas></div>')}
      ${panel('Competing slots', 'top channels, hours with the biggest head-to-head', '<div id="x2" style="max-height:320px;overflow:auto"></div>')}
    </div>
  </div>`;
  const st = channelStats(S.F).map(c => ({ ch: c.ch, R: c.reach || c.tvr * 1.4 }));
  const gs = [], ns = []; let gr = 0;
  st.forEach((c, i) => { gr += c.R; gs.push(+gr.toFixed(2)); ns.push(+netReach(st.slice(0, i + 1), getD).toFixed(2)); });
  mkChart('c-net', {
    type: 'bar', data: { labels: st.map((c, i) => (i + 1) + '. ' + chName(c.ch)), datasets: [{ label: 'Gross reach %', data: gs, backgroundColor: css('--line'), borderRadius: 3 }, { label: 'Net reach %', data: ns, backgroundColor: css('--accent'), borderRadius: 3 }] },
    options: { plugins: { legend: { position: 'bottom', labels: { boxWidth: 10 } } }, scales: { x: { grid: { display: false }, ticks: { maxRotation: 35 } } } }
  });
  const cp = competing(S.F, st.slice(0, 5).map(c => c.ch), 10);
  $('#x2').innerHTML = cp.length ? `<table><tbody>${cp.map(o => `<tr><td class="l" style="vertical-align:top"><button class="ib" ${xa({ h: o.h })}>${hl(o.h)}</button></td><td class="l" style="white-space:normal">${o.l.slice(0, 4).map(x => `<span class="sub" style="color:var(--ink)">${dot(x.ch)}${esc(chName(x.ch))}: ${esc(pn(x.p))} <span class="muted">TVR ${nf(x.m, 1)}</span></span>`).join('')}</td></tr>`).join('')}</tbody></table>` : '<p class="muted">No overlapping hours found.</p>';
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
    bt = `<div class="tw" style="max-height:340px;overflow:auto"><table><thead><tr><th>${D.label}</th><th>Airings</th><th>Avg TVR</th><th>Peak</th><th>Reach %</th><th>Share %</th><th>Steadiness</th><th>GRP share</th></tr></thead><tbody>${gs.slice(0, 150).map(x => `<tr class="click" ${xa(Object.assign({}, f, { [by]: x.k }))} data-push="1"><td>${by === 'ch' ? dot(x.k) : ''}${esc(D.name(x.k))}${by === 'p' ? `<span class="sub">${esc(chName(x.mxr.ch))} · ${hl(modeH(x))}</span>` : ''}</td><td>${x.n}</td><td>${nf(avgT(x), 2)}</td><td>${nf(x.mx, 1)}</td><td>${nf(avgR(x), 1)}</td><td>${nf(avgS(x), 1)}</td><td>${stdTag(ciOf(x), x.n)}</td><td><div class="mini-bar"><div class="tr"><div class="fl" style="width:${(x.sT / totT * 100).toFixed(0)}%"></div></div>${nf(x.sT / totT * 100, 1)}%</div></td></tr>`).join('')}</tbody></table></div>`;
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
      ${kp('Average TVR', nf(sm.avgT, 2), 'GRPs ' + ni(sm.sT))}
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
const VIEWS = { overview: vOverview, explore: vExplore, drill: vDrill, plan: vPlan, scen: vScen, dup: vDup, data: vData };
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
  if (t.id === 'sf-h') { S.SF.h = +t.value; killCharts('c-'); vExplore($('#v-explore')); return; }
  if (t.dataset.lvl !== undefined) {
    const i = +t.dataset.lvl, L = S.DR.levels.slice(0, i); if (t.value) L.push(t.value);
    if (i < S.DR.levels.length - 1 && t.value) S.DR.levels.slice(i + 1).forEach(k => { if (!L.includes(k)) L.push(k); });
    S.DR.levels = L.slice(0, 3); S.DR.open.clear(); vDrill($('#v-drill')); return;
  }
  if (t.dataset.split !== undefined) { S.P.split[t.dataset.split] = Math.max(0, Math.min(100, +t.value || 0)); savePlan(); recalc(); updatePlan(); return; }
  if (t.dataset.p !== undefined && t.type !== 'number') { S.P[t.dataset.p] = t.tagName === 'SELECT' ? +t.value : t.value; savePlan(); recalc(); updatePlan(); return; }
  if (t.id === 'd-rep') { S.P.repQ = Math.min(90, Math.max(0, +t.value || 0)); savePlan(); recalc(); return; }
  if (t.dataset && t.dataset.a !== undefined && t.closest('.mx')) {
    S.DUP[dkey(t.dataset.a, t.dataset.b)] = Math.min(1, Math.max(0, +t.value || 0)); lsSet('dup', S.DUP); recalc(); killCharts('c-'); vDup($('#v-dup')); return;
  }
  if (t.id === 'd-intra') { S.DINTRA = Math.min(1, Math.max(0, +t.value || 0)); lsSet('dintra', S.DINTRA); recalc(); return; }
});
let qTimer, pTimer;
document.addEventListener('input', e => {
  const t = e.target, P = S.P;
  if (t.id === 'f-q') { S.FS.q = t.value; clearTimeout(qTimer); qTimer = setTimeout(applyFilters, 250); return; }
  if (t.id === 'f-catq') { fillCats(t.value); return; }
  if (t.id === 'dq') { S.DQ = t.value; S.DPAGE = 0; clearTimeout(qTimer); qTimer = setTimeout(() => { vData($('#v-data')); const i = $('#dq'); i.focus(); i.setSelectionRange(i.value.length, i.value.length); }, 250); return; }
  const upd = () => { savePlan(); clearTimeout(pTimer); pTimer = setTimeout(() => { recalc(); updatePlan(); }, 150); };
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
