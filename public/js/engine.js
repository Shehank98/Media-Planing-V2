// Analytics engine: parsing, aggregation, duplication model and plan allocation.
// Pure functions, no DOM access.

export const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
export const DS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/* ---------- formatting ---------- */
export const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const nf = (v, d = 1) => Number(v || 0).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
export const ni = v => Number(v || 0).toLocaleString('en-US', { maximumFractionDigits: 0 });
export const hl = h => (h % 12 || 12) + (h < 12 ? ' AM' : ' PM');
export const band = h => hl(h) + ' – ' + hl((h + 1) % 24);
export const chName = s => s.split(' ').map(w => (w.length <= 3 && /^[A-Z0-9]+$/.test(w)) ? w : w.charAt(0) + w.slice(1).toLowerCase()).join(' ');
export const pn = s => s.toLowerCase().replace(/(^|[\s\-(/])([a-z])/g, (m, a, b) => a + b.toUpperCase()).replace(/\b(Tv|Itn|Pm|Am|Fm|Uk|Usa|Tv1)\b/g, x => x.toUpperCase());
export const fmtDate = iso => { const [y, m, d] = iso.split('-'); return `${d} ${MON[+m - 1]} ${y}`; };
export const dayDiff = (a, b) => Math.round((new Date(b) - new Date(a)) / 864e5) + 1;
export const lkr = v => v >= 1e6 ? 'LKR ' + nf(v / 1e6, 2) + 'M' : v >= 1e3 ? 'LKR ' + nf(v / 1e3, 0) + 'K' : 'LKR ' + ni(v);
export const pct = v => nf(v, 1) + '%';

/* ---------- dimensions (used by drill-down and detail views) ---------- */
export const DIMS = {
  ch: { label: 'Channel', plural: 'Channels', key: r => r.ch, name: chName },
  cat: { label: 'Category', plural: 'Categories', key: r => r.cat, name: pn },
  p: { label: 'Program', plural: 'Programs', key: r => r.p, name: pn },
  day: { label: 'Day', plural: 'Days', key: r => r.day, name: d => d, order: (a, b) => DAYS.indexOf(a) - DAYS.indexOf(b) },
  h: { label: 'Time band', plural: 'Time bands', key: r => r.h, name: h => band(+h), order: (a, b) => a - b },
  date: { label: 'Date', plural: 'Dates', key: r => r.date, name: fmtDate, order: (a, b) => a < b ? -1 : 1 },
  dp: { label: 'Daypart', plural: 'Dayparts', key: r => daypart(r.h), name: d => d, order: (a, b) => DPS.indexOf(a) - DPS.indexOf(b) }
};
export const DPS = ['Morning (5–12)', 'Daytime (12–18)', 'Prime (18–22)', 'Late (22–5)'];
export function daypart(h) { return h >= 5 && h < 12 ? DPS[0] : h >= 12 && h < 18 ? DPS[1] : h >= 18 && h < 22 ? DPS[2] : DPS[3]; }

/* ---------- parsing uploaded files ---------- */
export function mk(ch, date, day, s, e, p, dur, cat, tvr, tot, sh, reach, rp) {
  const h = parseInt(s.slice(0, 2), 10) % 24;
  return { ch, date, day, s, e, h, p, dur, cat, tvr, tot, sh, reach, rp };
}
const normKey = k => String(k).toLowerCase().replace(/[\s_]+/g, '');
const toNum = v => { const n = parseFloat(String(v).replace(/[,%]/g, '')); return isNaN(n) ? 0 : n; };
const pad = n => String(n).padStart(2, '0');
function toISO(v) {
  if (v instanceof Date && !isNaN(v)) { const d = new Date(v.getTime() + 12 * 36e5); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  if (typeof v === 'number' && v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 864e5) + 12 * 36e5); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/))) return m[3] + '-' + pad(m[2]) + '-' + pad(m[1]);
  const t = Date.parse(s); if (!isNaN(t)) { const d = new Date(t); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  return null;
}
// "20:09" -> "20:09", "8:09 PM" -> "20:09", Excel fraction 0.84 -> "20:09"
function toTime(v) {
  if (v === '' || v == null) return null;
  if (v instanceof Date && !isNaN(v)) { const d = new Date(v.getTime() + 30000); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
  if (typeof v === 'number') { const t = Math.round((v % 1) * 1440); return pad(Math.floor(t / 60) % 24) + ':' + pad(t % 60); }
  const m = String(v).match(/(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(AM|PM)?/i); if (!m) return null;
  let h = +m[1]; if (m[3]) { const pm = /pm/i.test(m[3]); if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
  return pad(h % 24) + ':' + m[2];
}

// json = array of row objects from SheetJS. Returns {rows, skip} or throws.
export function parseRows(json) {
  if (!json.length) throw new Error('The first sheet is empty.');
  const map = {}; Object.keys(json[0]).forEach(k => map[normKey(k)] = k);
  const col = (...n) => { for (const x of n) if (map[x] !== undefined) return map[x]; return null; };
  const c = {
    ch: col('channel', 'station'), date: col('date', 'broadcastdate'), day: col('day', 'dayname'), s: col('start', 'starttime'), e: col('end', 'endtime'),
    p: col('program', 'programme', 'programname'), dur: col('duration', 'dur', 'durationmins'), cat: col('category', 'genre'), tvr: col('tvr', 'rating'),
    tot: col('totaltvr'), sh: col('tvrshare%', 'tvrshare', 'share%', 'share'), reach: col('reach', "reach000", "reach('000)"), rp: col('reach%', 'reachpct', 'reachpercent')
  };
  const miss = ['ch', 'date', 's', 'p', 'tvr'].filter(k => !c[k]);
  if (miss.length) {
    const nm = { ch: 'Channel', date: 'Date', s: 'Start', p: 'Program', tvr: 'TVR' };
    throw new Error('Missing required column(s): ' + miss.map(k => nm[k]).join(', ') + '. Found: ' + Object.keys(json[0]).join(', '));
  }
  const rows = []; let skip = 0;
  for (const r of json) {
    const date = toISO(r[c.date]), s = toTime(r[c.s]), ch = String(r[c.ch]).trim().toUpperCase(), p = String(r[c.p]).trim();
    if (!date || !s || !ch || !p) { skip++; continue; }
    let day = c.day ? String(r[c.day]).trim() : '';
    day = DAYS.find(d => d.toLowerCase() === day.toLowerCase() || d.slice(0, 3).toLowerCase() === day.toLowerCase()) || DAYS[(new Date(date + 'T12:00:00').getDay() + 6) % 7];
    const e = c.e ? (toTime(r[c.e]) || '') : '';
    let dur = c.dur ? toNum(r[c.dur]) : 0;
    if (!dur && e) { const [a, b] = s.split(':').map(Number), [x, y] = e.split(':').map(Number); dur = ((x * 60 + y) - (a * 60 + b) + 1440) % 1440; }
    rows.push(mk(ch, date, day, s, e, p, dur, c.cat ? String(r[c.cat]).trim().toUpperCase() || 'OTHER' : 'OTHER',
      toNum(r[c.tvr]), c.tot ? toNum(r[c.tot]) : 0, c.sh ? toNum(r[c.sh]) : 0, c.reach ? toNum(r[c.reach]) : 0, c.rp ? toNum(r[c.rp]) : 0));
  }
  if (!rows.length) throw new Error('No usable rows found. Check the Date and Start formats.');
  return { rows, skip };
}

/* ---------- aggregation ---------- */
export function grp(rows, kf) {
  const m = new Map();
  for (const r of rows) {
    const k = kf(r); let g = m.get(k);
    if (!g) { g = { k, n: 0, sT: 0, sQ: 0, sR: 0, sS: 0, sD: 0, mx: 0, mxr: null, hc: {}, dc: {} }; m.set(k, g); }
    g.n++; g.sT += r.tvr; g.sQ += r.tvr * r.tvr; g.sR += r.rp; g.sS += r.sh; g.sD += r.dur;
    if (g.mxr === null || r.tvr > g.mx) { g.mx = r.tvr; g.mxr = r; }
    g.hc[r.h] = (g.hc[r.h] || 0) + 1; g.dc[r.day] = (g.dc[r.day] || 0) + 1;
  }
  return m;
}
export const avgT = g => g.n ? g.sT / g.n : 0;
export const avgR = g => g.n ? g.sR / g.n : 0;
export const avgS = g => g.n ? g.sS / g.n : 0;
export const modeHour = hc => +Object.keys(hc).sort((a, b) => hc[b] - hc[a])[0];
export const sdOf = g => { const m = g.sT / g.n; return Math.sqrt(Math.max(0, g.sQ / g.n - m * m)); };
// Consistency index = mean TVR / standard deviation (higher = steadier).
export const ciOf = g => { const m = g.sT / g.n, s = sdOf(g); return g.n < 2 ? 0 : s > .005 ? m / s : (m > 0 ? 10 : 0); };
export function steadiness(ci, n) {
  if (n < 2) return { cls: 's', label: 'One airing' };
  if (ci >= 7) return { cls: 'vs', label: 'Very steady' };
  if (ci >= 5) return { cls: 's', label: 'Steady' };
  if (ci >= 3) return { cls: 'm', label: 'Mixed' };
  return { cls: 'v', label: 'Volatile' };
}
export function summarize(rows) {
  let sT = 0, sR = 0, pk = rows[0];
  rows.forEach(r => { sT += r.tvr; sR += r.rp; if (r.tvr > pk.tvr) pk = r; });
  return { n: rows.length, sT, avgT: rows.length ? sT / rows.length : 0, avgR: rows.length ? sR / rows.length : 0, pk, progs: new Set(rows.map(r => r.ch + '|' + r.p)).size, chs: new Set(rows.map(r => r.ch)).size };
}

/* ---------- duplication model ---------- */
// Default co-viewing duplication between two channels (0 = independent, 1 = full overlap).
export function defD(a, b) {
  const has = n => a === n || b === n;
  const big = ['HIRU TV', 'DERANA TV', 'SIRASA TV'];
  if (has('HIRU TV') && has('DERANA TV')) return .70;
  if (big.includes(a) && big.includes(b)) return .55;
  if (has('VASANTHAM TV')) return has('SHAKTHI TV') ? .35 : .15;
  if (has('TV1')) return .25;
  if (has('SHAKTHI TV')) return (has('HIRU TV') || has('DERANA TV')) ? .20 : .25;
  if (has('ITN')) return .35;
  if (has('SWARNAVAHINI') || has('SIYATHA TV')) return (big.includes(a) || big.includes(b)) ? .50 : .40;
  return .40;
}
export const dkey = (a, b) => [a, b].sort().join('|');
export const mkGetD = DUP => (a, b) => a === b ? 1 : (DUP[dkey(a, b)] ?? defD(a, b));
// Union of two reach figures (in %) given a duplication coefficient d.
// d=0 -> random overlap (a*b/100); d=1 -> the smaller audience sits fully inside the larger.
export function uni(a, b, d) { const ind = a * b / 100; const ov = ind + d * (Math.min(a, b) - ind); return a + b - ov; }
export function netReach(list, getD) {
  let U = 0; const inc = [];
  for (const c of list) {
    if (!inc.length) U = c.R;
    else { const d = inc.reduce((s, x) => s + getD(x, c.ch), 0) / inc.length; U = uni(U, c.R, d); }
    inc.push(c.ch);
  }
  return U;
}

/* ---------- plan allocation ---------- */
// Strategy presets: tier budget split (Tier 1 / 2 / 3, %) and how much the optimiser
// values new reach (alpha) versus extra rating points (1 - alpha).
export const STRATS = {
  reach: { label: 'Reach', t: [65, 35, 0], a: 1 },
  balanced: { label: 'Balanced', t: [45, 35, 20], a: .85 },
  frequency: { label: 'Frequency', t: [30, 40, 30], a: .6 }
};
export function pctl(sorted, p) {
  if (!sorted.length) return 0;
  const i = (sorted.length - 1) * Math.min(100, Math.max(0, p)) / 100, lo = Math.floor(i), hi = Math.ceil(i);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (i - lo);
}
// Week weights for pacing: even (drip), burst (front-loaded), pulse (on/off weeks).
export function pacingWeights(pacing, W) {
  return Array.from({ length: W }, (_, i) => pacing === 'burst' ? W - i : pacing === 'pulse' ? (i % 2 === 0 ? 1 : 0) : 1);
}
export function channelStats(rows) {
  const cs = [...grp(rows, r => r.ch).values()].map(g => ({ ch: g.k, n: g.n, reach: avgR(g), tvr: avgT(g), share: avgS(g), grp: g.sT }));
  // Use reach where present; fall back to TVR if the file has no Reach % column.
  const hasReach = cs.some(c => c.reach > 0);
  const sr = cs.reduce((a, c) => a + (hasReach ? c.reach : c.tvr), 0) || 1, st = cs.reduce((a, c) => a + c.tvr, 0) || 1;
  cs.forEach(c => c.score = .6 * (hasReach ? c.reach : c.tvr) / sr + .4 * c.tvr / st);
  cs.sort((a, b) => b.score - a.score);
  return cs;
}

// Candidate programs on the chosen channels, each placed in a tier.
// Programs below the minimum TVR (minTvrPlan) are not candidates unless locked.
// Tier thresholds are percentiles of the candidates' average TVR:
// Tier 1 = at or above the tp1 percentile, Tier 3 = below the tp3 percentile, Tier 2 in between.
export function planCalc(rows, P) {
  const cs = channelStats(rows);
  if (!cs.length) return null;
  const lock = new Set(P.lock || []), excl = new Set(P.excl || []);
  let sel;
  if (P.chMode === 'manual' && P.chPick && P.chPick.length) sel = cs.filter(c => P.chPick.includes(c.ch));
  else sel = cs.slice(0, Math.min(P.nCh, cs.length));
  // Channels holding a locked program are always in the plan.
  lock.forEach(k => { const ch = k.split('||')[0]; if (!sel.find(c => c.ch === ch)) { const c = cs.find(x => x.ch === ch); if (c) sel.push(c); } });
  const ov = P.split || {};
  sel.forEach(c => c.manual = ov[c.ch] != null && ov[c.ch] !== '');
  if (!sel.length) return { cs, sel: [], items: [], thr: { t1: 0, t3: 0 } };
  const items = [], floor = Math.max(0, P.minTvrPlan ?? .5);
  for (const c of sel) {
    const pg = [...grp(rows.filter(r => r.ch === c.ch), r => r.p).values()].map(g => ({
      ch: c.ch, p: g.k, key: c.ch + '||' + g.k, n: g.n, mean: avgT(g), max: g.mx, rp: avgR(g) || avgT(g) * 1.4, ci: ciOf(g),
      hour: modeHour(g.hc), days: DAYS.filter(d => g.dc[d]), cat: g.mxr.cat, dur: g.sD / g.n, locked: lock.has(c.ch + '||' + g.k)
    }));
    let cand = pg.filter(x => x.n >= 2); if (cand.length < 3) cand = pg;
    cand = cand.filter(x => !excl.has(x.key) && x.mean > 0 && x.mean >= floor);
    pg.forEach(x => { if (x.locked && !cand.includes(x)) cand.push(x); });
    items.push(...cand);
  }
  const tv = items.map(x => x.mean).sort((a, b) => a - b);
  const thr = { t1: pctl(tv, P.tp1 ?? 75), t3: pctl(tv, P.tp3 ?? 25) };
  {
    items.forEach(x => {
      x.tier = x.mean >= thr.t1 ? 1 : x.mean < thr.t3 ? 3 : 2;
      x.dp = daypart(x.hour);
      x.sf = x.n < 2 ? .7 : .6 + .4 * Math.min(x.ci, 5) / 5; // steadiness factor
      x.role = x.ci < 3 && x.n >= 2 ? 'review' : (x.tier === 1 && x.ci >= 5) ? 'anchor' : 'support';
    });
  }
  return { cs, sel, items, thr };
}

// Spot-by-spot allocation for a budget factor f (1 = full budget).
// Each step buys the spot with the best gain per rupee, where gain = alpha x new net reach
// + (1 - alpha) x rating, scaled by steadiness. Order: locked programs (1 spot per active
// week), then minimum shares (manual channel shares, daypart minimums), then each tier up to
// its budget, then any leftover across tiers. Caps: spots per program per week, programs per
// channel per tier, daypart maximums, manual channel shares and the budget itself.
export function scen(plan, f, P, getD, dIntra) {
  const B = P.budget * f, W = Math.max(1, Math.round(P.weeks || 4));
  const wts = pacingWeights(P.pacing, W), act = wts.filter(w => w > 0).length;
  const capTot = Math.max(1, Math.round(P.capWk || 3)) * act;
  const q = Math.min(.95, Math.max(0, (P.repQ ?? 40) / 100));
  const len = (P.spotLen || 30) / 30, cprp = P.cprp > 0 ? P.cprp : 25000, floor = Math.max(0, P.minRate || 0);
  const alpha = (STRATS[P.strategy] || STRATS.balanced).a;
  const tiers = (P.tiers && P.tiers.length === 3 ? P.tiers : STRATS.balanced.t).map(v => Math.max(0, +v || 0));
  const items = plan.items.map(x => ({ ...x, cost: Math.max(floor, cprp * x.mean * len), cnt: 0 }));
  const byCh = new Map(); items.forEach(x => { if (!byCh.has(x.ch)) byCh.set(x.ch, []); byCh.get(x.ch).push(x); });
  // Reach of one program after k spots: each repeat spot adds q x the previous spot's new reach.
  const progR = (x, k) => k <= 0 ? 0 : x.rp * (1 - Math.pow(q, k)) / (1 - q);
  const chR = (ch, sx, sk) => byCh.get(ch).map(x => progR(x, x === sx ? sk : x.cnt)).filter(v => v > 0).sort((a, b) => b - a).reduce((u, r, i) => i === 0 ? r : uni(u, r, dIntra), 0);
  const chRv = new Map([...byCh.keys()].map(c => [c, 0]));
  const net = (och, oR) => netReach([...chRv.entries()].map(([ch, R]) => ({ ch, R: ch === och ? oR : R })).filter(c => c.R > 0).sort((a, b) => b.R - a.R), getD);

  // Daypart limits apply only to dayparts that exist among the candidates.
  const present = new Set(items.map(x => x.dp));
  const dpMin = {}, dpMax = {}; let relaxed = false;
  DPS.forEach((d, i) => { dpMin[d] = present.has(d) ? (P.dpMin?.[i] || 0) / 100 : 0; dpMax[d] = present.has(d) ? (P.dpMax?.[i] ?? 100) / 100 : 0; });
  const maxSum = [...present].reduce((a, d) => a + dpMax[d], 0);
  if (maxSum < 1 && maxSum > 0) { relaxed = true; present.forEach(d => dpMax[d] = dpMax[d] / maxSum); }
  const ov = P.split || {}, chFix = {};
  plan.sel.forEach(c => { if (c.manual) chFix[c.ch] = Math.max(0, Math.min(100, +ov[c.ch])) / 100; });

  let curNet = 0, spent = 0;
  const tierSp = [0, 0, 0, 0], dpSp = {}, chSp = {}, chN = {};
  DPS.forEach(d => dpSp[d] = 0); byCh.forEach((_, c) => { chSp[c] = 0; });
  const E = 1e-6;
  const ok = x => spent + x.cost <= B + E && x.cnt < capTot && (x.cnt > 0 || (chN[x.ch + x.tier] || 0) < P.nProg || x.locked) &&
    dpSp[x.dp] + x.cost <= dpMax[x.dp] * B + E && (chFix[x.ch] == null || chSp[x.ch] + x.cost <= chFix[x.ch] * B + E);
  const gain = x => { const R = chR(x.ch, x, x.cnt + 1); return (alpha * (net(x.ch, R) - curNet) + (1 - alpha) * x.mean) * x.sf / x.cost; };
  const take = x => { if (!x.cnt) chN[x.ch + x.tier] = (chN[x.ch + x.tier] || 0) + 1; x.cnt++; spent += x.cost; tierSp[x.tier] += x.cost; dpSp[x.dp] += x.cost; chSp[x.ch] += x.cost; chRv.set(x.ch, chR(x.ch)); curNet = net(); };
  const fill = (pool, extra) => {
    for (;;) {
      let best = null, bg = -Infinity;
      for (const x of pool) { if (!ok(x) || (extra && !extra(x))) continue; const g = gain(x); if (g > bg) { bg = g; best = x; } }
      if (!best) break; take(best);
    }
  };
  items.filter(x => x.locked).forEach(x => { while (x.cnt < act && ok(x)) take(x); });
  Object.keys(chFix).forEach(ch => fill(byCh.get(ch) || [], x => chSp[ch] + x.cost <= chFix[ch] * B + E));
  DPS.forEach(d => { if (dpMin[d] > 0) fill(items.filter(x => x.dp === d), () => dpSp[d] < dpMin[d] * B); });
  [1, 2, 3].forEach(t => fill(items.filter(x => x.tier === t), x => tierSp[t] + x.cost <= tiers[t - 1] / 100 * B + E));
  fill(items);

  const tot = spent || 1;
  const kept = items.filter(x => x.cnt > 0).map(x => ({ ...x, spots: x.cnt, bud: x.cnt * x.cost, grps: x.cnt * x.mean, R: progR(x, x.cnt), k: x.cnt * x.cost / tot, atCap: x.cnt >= capTot }))
    .sort((a, b) => b.bud - a.bud);
  const m = new Map();
  kept.forEach(x => { let c = m.get(x.ch); if (!c) { c = { ch: x.ch, w: 0, bud: 0, progs: [], spots: 0, grps: 0, R: chRv.get(x.ch) }; m.set(x.ch, c); } c.w += x.k; c.bud += x.bud; c.spots += x.spots; c.grps += x.grps; c.progs.push(x); });
  const chs = [...m.values()].sort((a, b) => b.w - a.w);
  const gross = chs.reduce((a, c) => a + c.R, 0), grps = kept.reduce((a, x) => a + x.grps, 0);
  return {
    B, f, spent, unspent: Math.max(0, B - spent), kept, dropped: [], chs, gross, net: curNet,
    loss: gross > 0 ? (gross - curNet) / gross : 0, eff: B > 0 ? curNet / (B / 1e6) : 0, grps, spots: kept.reduce((a, x) => a + x.spots, 0),
    freq: curNet > 0 && grps > 0 ? grps / curNet : 0, r3: reachAtLeast(curNet, curNet > 0 ? grps / curNet : 0, 3), capTot, act, W, thr: plan.thr, relaxed,
    tiers: [1, 2, 3].map(t => ({ t, target: tiers[t - 1], actual: tierSp[t] / tot * 100, n: kept.filter(x => x.tier === t).length, avail: items.filter(x => x.tier === t).length })),
    dps: DPS.map(d => ({ d, present: present.has(d), min: dpMin[d] * 100, max: dpMax[d] * 100, actual: dpSp[d] / tot * 100 }))
  };
}

// Effective reach: % reached at least n times, assuming the frequency among reached
// viewers follows a zero-truncated Poisson with the plan's average frequency.
export function reachAtLeast(net, freq, n = 3) {
  if (!(net > 0) || !(freq > 1)) return n <= 1 ? net : 0;
  let lo = 1e-6, hi = 60;
  for (let i = 0; i < 60; i++) { const m = (lo + hi) / 2; if (m / (1 - Math.exp(-m)) < freq) lo = m; else hi = m; }
  const l = (lo + hi) / 2; let cdf = 0, t = Math.exp(-l);
  for (let k = 0; k < n; k++) { cdf += t; t *= l / (k + 1); }
  return net * (1 - cdf) / (1 - Math.exp(-l));
}

// Weekly flighting: spreads each program's spots over the flight using the pacing weights
// (never more than the weekly cap), then places them on the program's air days.
// Roadblock puts competing same-hour spots on the same nights (different viewers, more reach);
// stagger puts them on different nights (same viewers more often, more frequency).
export function buildSchedule(A, P) {
  const W = A.W, wts = pacingWeights(P.pacing, W), cap = Math.max(1, Math.round(P.capWk || 3));
  const start = new Date((P.start || new Date().toISOString().slice(0, 10)) + 'T12:00:00');
  const sDow = (start.getDay() + 6) % 7;
  const occ = new Map(), rows = [], grid = [];
  const weeks = Array.from({ length: W }, () => ({ spots: 0, bud: 0, grps: 0, byCh: {} }));
  A.kept.forEach((x, pi) => {
    // Largest-remainder split by weight (ties rotate by program so even pacing stays even),
    // then spill anything above the weekly cap.
    const tw = wts.reduce((a, b) => a + b, 0) || 1;
    const raw = wts.map(w => x.spots * w / tw), alloc = raw.map(Math.floor);
    let left = x.spots - alloc.reduce((a, b) => a + b, 0);
    raw.map((v, i) => [v - Math.floor(v), i]).filter(([, i]) => wts[i] > 0).sort((a, b) => (b[0] - a[0]) || ((a[1] - pi % W + W) % W) - ((b[1] - pi % W + W) % W)).forEach(([, i]) => { if (left > 0) { alloc[i]++; left--; } });
    let spill = 0; alloc.forEach((n, i) => { if (n > cap) { spill += n - cap; alloc[i] = cap; } });
    for (let i = 0; spill > 0 && i < W; i++) if (wts[i] > 0) { const room = cap - alloc[i]; const t = Math.min(room, spill); alloc[i] += t; spill -= t; }
    grid.push({ x, alloc });
    const days = x.days.length ? x.days : DAYS;
    alloc.forEach((n, w) => {
      if (!n) return;
      const used = {};
      for (let i = 0; i < n; i++) {
        const sc = d => { const s = occ.get(w + '|' + d + '|' + x.hour); const o = s ? [...s].filter(c => c !== x.ch).length : 0; return P.same === 'stagger' ? -o : o; };
        const d = [...days].sort((a, b) => (used[a] || 0) - (used[b] || 0) || sc(b) - sc(a) || DAYS.indexOf(a) - DAYS.indexOf(b))[0];
        used[d] = (used[d] || 0) + 1;
        const k = w + '|' + d + '|' + x.hour; if (!occ.has(k)) occ.set(k, new Set()); occ.get(k).add(x.ch);
      }
      Object.keys(used).sort((a, b) => DAYS.indexOf(a) - DAYS.indexOf(b)).forEach(d => {
        const dt = new Date(start.getTime() + (w * 7 + (DAYS.indexOf(d) - sDow + 7) % 7) * 864e5);
        rows.push({ week: w + 1, date: dt.toISOString().slice(0, 10), day: d, ch: x.ch, p: x.p, hour: x.hour, tier: x.tier, spots: used[d], cost: used[d] * x.cost, grps: used[d] * x.mean });
      });
      const wk = weeks[w]; wk.spots += n; wk.bud += n * x.cost; wk.grps += n * x.mean; wk.byCh[x.ch] = (wk.byCh[x.ch] || 0) + n;
    });
  });
  rows.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : a.hour - b.hour);
  // Same-hour clashes between different channels on the same night.
  const clashes = [...occ.entries()].filter(([, s]) => s.size > 1).length;
  return { rows, grid, weeks, wts, clashes, start: start.toISOString().slice(0, 10) };
}

// Hours where two or more plan channels air programs head to head.
export function competing(rows, chList, lim = 8) {
  const set = new Set(chList);
  const g = grp(rows.filter(r => set.has(r.ch)), r => r.h + '|' + r.ch + '|' + r.p);
  const byH = new Map();
  g.forEach(x => {
    const [h, ch, p] = x.k.split('|'); const hv = +h; let o = byH.get(hv); if (!o) { o = new Map(); byH.set(hv, o); }
    const cur = o.get(ch), m = x.sT / x.n; if (!cur || m > cur.m) o.set(ch, { ch, p, m, n: x.n });
  });
  const out = [];
  byH.forEach((o, h) => { const l = [...o.values()].sort((a, b) => b.m - a.m); if (l.length >= 2) out.push({ h, l, tot: l.reduce((a, x) => a + x.m, 0) }); });
  return out.sort((a, b) => b.tot - a.tot).slice(0, lim);
}

/* ---------- health checks ---------- */
export function healthChecks(rows, plan, A, P, getD) {
  const L = [];
  if (!plan || !plan.items.length || !A) return L;
  const it = A.kept;
  const steady = it.filter(x => x.ci >= 5 || x.n < 2).length;
  L.push({ t: steady >= it.length * .7 ? 'ok' : 'wa', m: `${steady} of ${it.length} programs deliver steady ratings week to week` });
  const anchors = it.filter(x => x.role === 'anchor');
  L.push({ t: anchors.length ? 'ok' : 'wa', m: anchors.length ? `${anchors.length} anchor program${anchors.length > 1 ? 's' : ''} carry the plan (Tier 1 and steady)` : 'No anchor programs: nothing in the basket is both Tier 1 and steady' });
  if (A.unspent > A.B * .02) L.push({ t: 'wa', m: `<b>${lkr(A.unspent)}</b> (${nf(A.unspent / A.B * 100, 0)}%) is unspent because caps are reached. Raise the spots-per-week cap, allow more programs per channel or widen daypart limits.` });
  A.tiers.forEach(t => { if (t.target > 0 && Math.abs(t.actual - t.target) > 10) L.push({ t: 'in', m: `Tier ${t.t} gets ${nf(t.actual, 0)}% against a ${nf(t.target, 0)}% target${t.avail ? '' : ' (no Tier ' + t.t + ' programs in this brief)'}; leftover budget rolled to other tiers.` }); });
  A.dps.forEach(d => { if (d.present && d.min > 0 && d.actual + .5 < d.min) L.push({ t: 'wa', m: `${d.d} gets ${nf(d.actual, 0)}%, below its ${nf(d.min, 0)}% minimum. Not enough spots available there under the caps.` }); });
  if (A.relaxed) L.push({ t: 'in', m: 'Daypart maximums were scaled up because the brief only covers some dayparts.' });
  const capped = it.filter(x => x.atCap).length;
  if (capped) L.push({ t: 'in', m: `${capped} program${capped > 1 ? 's are' : ' is'} at the ${P.capWk} spots/week cap, so budget moved to the next best slots.` });
  const keptKeys = new Set(it.map(x => x.key));
  const missed = [...plan.items].sort((a, b) => b.mean - a.mean).slice(0, 3).filter(x => !keptKeys.has(x.key));
  if (missed.length) L.push({ t: 'in', m: `Top-rated ${missed.map(x => `<b>${pn(x.p)}</b> (TVR ${nf(x.mean, 1)})`).join(', ')} not bought: other slots add more new reach per rupee. Lock ${missed.length > 1 ? 'them' : 'it'} in the basket to force a buy.` });
  const vol = it.filter(x => x.role === 'review').sort((a, b) => b.k - a.k)[0];
  if (vol && vol.k >= .08) L.push({ t: 'wa', m: `<b>${pn(vol.p)}</b> (${chName(vol.ch)}) takes ${nf(vol.k * 100, 0)}% of the plan but its ratings are volatile (consistency ${nf(vol.ci, 1)}). Lock a steadier backup or exclude it.` });
  if (A.loss > .35) L.push({ t: 'wa', m: `${nf(A.loss * 100, 0)}% of gross reach is lost to duplication. Channels in this plan share a lot of viewers.` });
  else L.push({ t: 'ok', m: `Duplication loss is ${nf(A.loss * 100, 0)}%, an acceptable overlap` });
  if (P.target > 0) {
    if (A.net >= P.target) L.push({ t: 'ok', m: `Target reach of ${nf(P.target, 0)}% is met (${nf(A.net, 1)}%)` });
    else L.push({ t: 'wa', m: `Net reach ${nf(A.net, 1)}% is below the ${nf(P.target, 0)}% target. Add budget or channels, or switch to the Reach strategy.` });
  }
  const odd = rows.filter(r => r.sh > 100).length;
  if (odd) L.push({ t: 'wa', m: `${odd} row${odd > 1 ? 's show' : ' shows'} TVR share above 100%. Check the source file.` });
  return L;
}

/* ---------- market insights (Overview) ---------- */
export function insights(rows) {
  const out = [];
  if (!rows.length) return out;
  const hrs = [...grp(rows, r => r.h).values()].filter(g => g.n >= 3).sort((a, b) => avgT(b) - avgT(a));
  if (hrs.length) {
    const a = hrs[0], b = hrs[1];
    out.push({ big: hl(a.k), t: `is the peak hour. Average TVR ${nf(avgT(a), 2)} and reach ${nf(avgR(a), 1)}%${b ? `, ${nf((avgT(a) / (avgT(b) || 1) - 1) * 100, 0)}% above ${hl(b.k)}` : ''}.`, cls: '', drill: { h: a.k } });
  }
  const wd = rows.filter(r => DAYS.indexOf(r.day) < 5), we = rows.filter(r => DAYS.indexOf(r.day) >= 5);
  if (wd.length && we.length) {
    const a = wd.reduce((s, r) => s + r.tvr, 0) / wd.length, b = we.reduce((s, r) => s + r.tvr, 0) / we.length;
    const d = (b - a) / (a || 1) * 100;
    out.push({ big: (d > 0 ? '+' : '−') + nf(Math.abs(d), 0) + '%', t: `${d < 0 ? 'lower' : 'higher'} average TVR on weekends (${nf(b, 2)}) than on weekdays (${nf(a, 2)}).`, cls: 'w', drill: null });
  }
  const cats = [...grp(rows, r => r.cat).values()].sort((a, b) => b.sT - a.sT);
  const tot = cats.reduce((s, g) => s + g.sT, 0) || 1;
  if (cats.length) out.push({ big: nf(cats[0].sT / tot * 100, 0) + '%', t: `of all viewing (GRPs) comes from ${pn(cats[0].k)}, the biggest category.`, cls: 'g', drill: { cat: cats[0].k } });
  const day = rows.filter(r => r.h < 17 && r.h >= 6).sort((a, b) => b.tvr - a.tvr)[0];
  if (day && day.tvr >= 2) out.push({ big: nf(day.tvr, 1), t: `TVR daytime sleeper hit: ${pn(day.p)} on ${chName(day.ch)} at ${day.s}.`, cls: 'p', drill: { ch: day.ch, p: day.p } });
  return out;
}

/* ---------- CSV helper ---------- */
export function toCSV(head, rows) {
  const q = v => { const s = String(v ?? ''); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  return [head.map(q).join(','), ...rows.map(r => r.map(q).join(','))].join('\n');
}
