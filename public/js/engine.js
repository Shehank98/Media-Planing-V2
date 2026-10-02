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
// P: {budget, nCh, nProg, cut, cprp, chMode, chPick[], split{}, lock[], excl[], target}
export function channelStats(rows) {
  const cs = [...grp(rows, r => r.ch).values()].map(g => ({ ch: g.k, n: g.n, reach: avgR(g), tvr: avgT(g), share: avgS(g), grp: g.sT }));
  // Use reach where present; fall back to TVR if the file has no Reach % column.
  const hasReach = cs.some(c => c.reach > 0);
  const sr = cs.reduce((a, c) => a + (hasReach ? c.reach : c.tvr), 0) || 1, st = cs.reduce((a, c) => a + c.tvr, 0) || 1;
  cs.forEach(c => c.score = .6 * (hasReach ? c.reach : c.tvr) / sr + .4 * c.tvr / st);
  cs.sort((a, b) => b.score - a.score);
  return cs;
}

export function planCalc(rows, P) {
  const cs = channelStats(rows);
  if (!cs.length) return null;
  const lock = new Set(P.lock || []), excl = new Set(P.excl || []);
  let sel;
  if (P.chMode === 'manual' && P.chPick && P.chPick.length) sel = cs.filter(c => P.chPick.includes(c.ch));
  else sel = cs.slice(0, Math.min(P.nCh, cs.length));
  // Channels holding a locked program are always in the plan.
  lock.forEach(k => { const ch = k.split('||')[0]; if (!sel.find(c => c.ch === ch)) { const c = cs.find(x => x.ch === ch); if (c) sel.push(c); } });
  if (!sel.length) return { cs, sel: [], items: [] };

  // Channel weights: score-based, then manual overrides (in %) if given.
  const ts = sel.reduce((a, c) => a + c.score, 0) || 1;
  sel.forEach(c => { c.auto = c.score / ts; c.w = c.auto; c.manual = false; });
  const ov = P.split || {};
  const fixed = sel.filter(c => ov[c.ch] != null && ov[c.ch] !== '');
  if (fixed.length) {
    const fixedSum = Math.min(100, fixed.reduce((a, c) => a + Math.max(0, +ov[c.ch]), 0)) / 100;
    const rest = sel.filter(c => !fixed.includes(c)), restAuto = rest.reduce((a, c) => a + c.auto, 0) || 1;
    fixed.forEach(c => { c.w = Math.max(0, +ov[c.ch]) / 100; c.manual = true; });
    rest.forEach(c => c.w = (1 - fixedSum) * c.auto / restAuto);
    const tot = sel.reduce((a, c) => a + c.w, 0) || 1; sel.forEach(c => c.w /= tot);
  }

  const items = [];
  for (const c of sel) {
    const pg = [...grp(rows.filter(r => r.ch === c.ch), r => r.p).values()].map(g => ({
      ch: c.ch, p: g.k, key: c.ch + '||' + g.k, n: g.n, mean: avgT(g), max: g.mx, rp: avgR(g), ci: ciOf(g),
      hour: modeHour(g.hc), days: Object.keys(g.dc), cat: g.mxr.cat, dur: g.sD / g.n
    })).filter(x => !excl.has(x.key));
    let cand = pg.filter(x => x.n >= 2); if (cand.length < P.nProg) cand = pg;
    const mm = Math.max(...cand.map(x => x.mean), .0001);
    pg.forEach(x => x.score = .6 * x.mean / mm + .4 * Math.min(x.ci, 5) / 5);
    cand.sort((a, b) => b.score - a.score);
    const locked = pg.filter(x => lock.has(x.key));
    const top = [...locked, ...cand.filter(x => !lock.has(x.key)).slice(0, Math.max(0, P.nProg - locked.length))];
    const ss = top.reduce((a, x) => a + x.score, 0);
    top.forEach(x => { x.locked = lock.has(x.key); x.share = ss > 0 ? x.score / ss : 1 / top.length; x.alloc = c.w * x.share; items.push(x); });
    c.nProg = top.length;
  }
  items.sort((a, b) => b.alloc - a.alloc);
  // Roles: Anchor = high rating and steady, Review = volatile, else Support.
  const means = items.map(x => x.mean).sort((a, b) => b - a);
  const cut = means[Math.max(0, Math.floor(means.length * .4) - 1)] || 0;
  items.forEach(x => x.role = x.ci < 3 && x.n >= 2 ? 'review' : (x.mean >= cut && x.ci >= 5) ? 'anchor' : 'support');
  return { cs, sel, items };
}

// Apply a budget factor f (1 = full budget). Cuts drop the lowest scoring slots first; locked slots are protected.
export function scen(plan, f, P, getD, dIntra) {
  const B = P.budget * f;
  let kept = plan.items, dropped = [];
  if (f < 1) {
    const keepN = Math.max(1, Math.ceil(plan.items.length * f));
    const order = [...plan.items].sort((a, b) => (b.locked - a.locked) || (b.alloc - a.alloc));
    const keepSet = new Set(order.slice(0, keepN));
    kept = plan.items.filter(x => keepSet.has(x)); dropped = plan.items.filter(x => !keepSet.has(x));
  }
  const tot = kept.reduce((a, x) => a + x.alloc, 0) || 1;
  kept = kept.map(x => {
    const k = x.alloc / tot, bud = B * k;
    const spotCost = P.cprp > 0 ? P.cprp * Math.max(x.mean, .05) * (P.spotLen || 30) / 30 : 0;
    const spots = spotCost > 0 ? Math.floor(bud / spotCost) : 0;
    return { ...x, k, bud, spots, grps: spots * x.mean };
  });
  const m = new Map();
  kept.forEach(x => { let c = m.get(x.ch); if (!c) { c = { ch: x.ch, w: 0, bud: 0, progs: [], spots: 0, grps: 0 }; m.set(x.ch, c); } c.w += x.k; c.bud += x.bud; c.spots += x.spots; c.grps += x.grps; c.progs.push(x); });
  const chs = [...m.values()].sort((a, b) => b.w - a.w);
  chs.forEach(c => {
    const rs = c.progs.map(x => x.rp || x.mean * 1.4).sort((a, b) => b - a);
    c.R = rs.reduce((u, r, i) => i === 0 ? r : uni(u, r, dIntra), 0);
  });
  const gross = chs.reduce((a, c) => a + c.R, 0), net = netReach(chs, getD);
  const grps = kept.reduce((a, x) => a + x.grps, 0);
  return { B, f, kept, dropped, chs, gross, net, loss: gross > 0 ? (gross - net) / gross : 0, eff: B > 0 ? net / (B / 1e6) : 0, grps, spots: kept.reduce((a, x) => a + x.spots, 0), freq: net > 0 && grps > 0 ? grps / net : 0 };
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
  if (!plan || !plan.items.length) return L;
  const it = A.kept;
  const steady = it.filter(x => x.ci >= 5 || x.n < 2).length;
  L.push({ t: steady === it.length ? 'ok' : steady >= it.length * .7 ? 'ok' : 'wa', m: `${steady} of ${it.length} programs deliver steady ratings week to week` });
  const anchors = it.filter(x => x.role === 'anchor');
  L.push({ t: anchors.length ? 'ok' : 'wa', m: anchors.length ? `${anchors.length} anchor program${anchors.length > 1 ? 's' : ''} carry the plan (high rating and steady)` : 'No anchor programs: nothing in the basket is both high rating and steady' });
  const vol = it.filter(x => x.role === 'review').sort((a, b) => b.k - a.k)[0];
  if (vol && vol.k >= .08) L.push({ t: 'wa', m: `<b>${pn(vol.p)}</b> (${chName(vol.ch)}) takes ${nf(vol.k * 100, 0)}% of the plan but its ratings are volatile (consistency ${nf(vol.ci, 1)}). Consider moving 5–7 pts to a steadier anchor.` });
  const solo = A.chs.find(c => c.progs.length === 1 && c.w >= .15);
  if (solo) L.push({ t: 'wa', m: `All of ${chName(solo.ch)}'s ${nf(solo.w * 100, 0)}% rides on one program. Add a second program to spread the risk.` });
  if (A.loss > .35) L.push({ t: 'wa', m: `${nf(A.loss * 100, 0)}% of gross reach is lost to duplication. Channels in this plan share a lot of viewers.` });
  else L.push({ t: 'ok', m: `Duplication loss is ${nf(A.loss * 100, 0)}%, an acceptable overlap` });
  const days = new Set(rows.map(r => r.day));
  const wd = rows.filter(r => DAYS.indexOf(r.day) < 5), we = rows.filter(r => DAYS.indexOf(r.day) >= 5);
  if (wd.length && we.length) {
    const a = wd.reduce((s, r) => s + r.tvr, 0) / wd.length, b = we.reduce((s, r) => s + r.tvr, 0) / we.length;
    const d = (b - a) / (a || 1) * 100;
    L.push({ t: 'in', m: `Weekend airings average ${nf(Math.abs(d), 0)}% ${d < 0 ? 'lower' : 'higher'} TVR than weekdays in the current filter` });
  } else if (days.size) L.push({ t: 'in', m: `Filter covers ${[...days].length} day${days.size > 1 ? 's' : ''} only` });
  if (P.target > 0) {
    if (A.net >= P.target) L.push({ t: 'ok', m: `Target reach of ${nf(P.target, 0)}% is met (${nf(A.net, 1)}%)` });
    else L.push({ t: 'wa', m: `Net reach ${nf(A.net, 1)}% is below the ${nf(P.target, 0)}% target. Add channels or programs, or widen the time band.` });
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
