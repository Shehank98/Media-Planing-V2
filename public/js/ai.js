// Planner AI: builds a compact aggregate summary for Gemini (via /api/ai),
// and answers locally with rule-based logic when the server has no API key.
import { grp, avgT, avgR, ciOf, hl, nf, pn, chName, DAYS, summarize } from './engine.js';

export function buildContext(S) {
  const F = S.F;
  const sm = summarize(F);
  const r2 = v => Math.round(v * 100) / 100;
  const chans = [...grp(F, r => r.ch).values()].sort((a, b) => avgT(b) - avgT(a)).map(g => ({ channel: g.k, airings: g.n, avgTVR: r2(avgT(g)), avgReachPct: r2(avgR(g)), GRPs: Math.round(g.sT) }));
  const hours = [...grp(F, r => r.h).values()].sort((a, b) => a.k - b.k).map(g => ({ hour: hl(g.k), h24: g.k, airings: g.n, avgTVR: r2(avgT(g)), avgReachPct: r2(avgR(g)) }));
  const days = [...grp(F, r => r.day).values()].sort((a, b) => DAYS.indexOf(a.k) - DAYS.indexOf(b.k)).map(g => ({ day: g.k, avgTVR: r2(avgT(g)), avgReachPct: r2(avgR(g)), peakTVR: r2(g.mx) }));
  const cats = [...grp(F, r => r.cat).values()].sort((a, b) => b.sT - a.sT).slice(0, 10).map(g => ({ category: g.k, GRPs: Math.round(g.sT), avgTVR: r2(avgT(g)) }));
  const progs = [...grp(F, r => r.ch + '||' + r.p).values()].filter(g => g.n >= 2).sort((a, b) => avgT(b) - avgT(a)).slice(0, 25)
    .map(g => { const [ch, p] = g.k.split('||'); return { program: p, channel: ch, airings: g.n, avgTVR: r2(avgT(g)), peakTVR: r2(g.mx), avgReachPct: r2(avgR(g)), consistency: r2(ciOf(g)), category: g.mxr.cat, usualHour: hl(g.mxr.h) }; });
  const daytime = F.filter(r => r.h >= 6 && r.h < 17).sort((a, b) => b.tvr - a.tvr).slice(0, 5).map(r => ({ program: r.p, channel: r.ch, date: r.date, start: r.s, TVR: r.tvr }));
  const A = S.cur;
  const plan = A ? {
    budgetLKR: Math.round(A.B), budgetCutPct: S.P.cut, channelsInPlan: S.P.nCh, programsPerChannel: S.P.nProg,
    netReachPct: r2(A.net), reach3plusPct: r2(A.r3), grossReachPct: r2(A.gross), duplicationLossPct: r2(A.loss * 100), budgetUsedLKR: Math.round(A.spent),
    strategy: S.P.strategy, tiers: A.tiers.map(t => ({ tier: t.t, targetPct: t.target, actualPct: r2(t.actual), programsBought: t.n })), tier1MinTVR: r2(A.thr.t1), tier3BelowTVR: r2(A.thr.t3),
    maxSpotsPerProgramPerWeek: S.P.capWk, dayparts: A.dps.filter(d => d.present).map(d => ({ daypart: d.d, minPct: d.min, maxPct: r2(d.max), actualPct: r2(d.actual) })),
    flighting: S.sched ? { weeks: A.W, start: S.sched.start, pacing: S.P.pacing, sameHourRivals: S.P.same, spotsPerWeek: S.sched.weeks.map(w => w.spots) } : null,
    estSpots: A.spots, estGRPs: Math.round(A.grps), costPerRatingPointLKR: S.P.cprp,
    split: A.chs.map(c => ({ channel: c.ch, sharePct: r2(c.w * 100), budgetLKR: Math.round(c.bud) })),
    basket: A.kept.map(x => ({ program: x.p, channel: x.ch, tier: x.tier, spots: x.spots, role: x.role, sharePct: r2(x.k * 100), avgTVR: r2(x.mean), consistency: r2(x.ci), usualHour: hl(x.hour) })),
    droppedByCut: A.dropped.map(x => x.p),
    healthChecks: (S.health || []).map(h => h.m.replace(/<[^>]+>/g, ''))
  } : null;
  return {
    filters: { from: S.FS.from, to: S.FS.to, channels: [...S.FS.ch], days: [...S.FS.day], hours: hl(S.FS.h0) + ' to ' + hl(S.FS.h1), categoriesSelected: S.FS.cat.size + ' of ' + S.CATS.length },
    totals: { airings: sm.n, programs: sm.progs, avgTVR: r2(sm.avgT), avgReachPct: r2(sm.avgR), totalGRPs: Math.round(sm.sT) },
    allChannels: S.CH, channels: chans, hours, days, topCategories: cats, topPrograms: progs, daytimeSleeperHits: daytime, currentPlan: plan
  };
}

export async function askRemote(question, context, history) {
  const r = await fetch('/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ question, context, history }) });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) { const e = new Error(j.error || 'AI error ' + r.status); e.status = r.status; throw e; }
  return j.text;
}

// Pulls a trailing "ACTION: {...}" line out of an answer.
export function splitAction(text) {
  const m = text.match(/ACTION:\s*(\{[\s\S]*?\})\s*$/m);
  if (!m) return { text, action: null };
  let action = null; try { action = JSON.parse(m[1]); } catch (e) { action = null; }
  return { text: text.replace(m[0], '').trim(), action };
}

/* ---------- offline planner (no API key) ---------- */
export function askLocal(q, S, simulate) {
  const t = q.toLowerCase(), F = S.F, A = S.cur;
  const li = a => '<ul>' + a.map(x => `<li>${x}</li>`).join('') + '</ul>';
  if (!F.length) return { html: 'No data matches the current filters. Widen the filters first.' };
  if (!A) return { html: 'There is no plan to work with yet. On the Planner tab, pick at least one channel or restore excluded programs.' };

  // Budget change
  let m = t.match(/(cut|reduce|lower|drop|decrease|less)[^\d]*(\d{1,2})\s*%?/) || t.match(/(\d{1,2})\s*%[^.]*?(cut|less|lower|reduc)/);
  if (m) {
    const pct = Math.min(50, +(m[2] && /\d/.test(m[2]) ? m[2] : m[1]));
    const B = simulate({ cut: pct });
    return {
      html: `<b>Short answer:</b> net reach moves from ${nf(S.base.net, 1)}% (full budget) to about ${nf(B.net, 1)}% for a ${pct}% budget cut, and reach 3+ from ${nf(S.base.r3, 1)}% to ${nf(B.r3, 1)}%.` +
        li([`Spots fall from ${S.base.spots} to ${B.spots}; the optimiser removes repeat spots first, so frequency drops more than reach`,
          B.dropped.length ? `Drop: ${B.dropped.slice(0, 4).map(x => pn(x.p)).join(', ')}` : 'No slots dropped',
          `New split: ${B.chs.map(c => chName(c.ch) + ' ' + nf(c.w * 100, 0) + '%').join(', ')}`,
          'Locked programs and the tier split are kept']),
      action: { cut: pct }
    };
  }
  // Programs at an hour: "10 pm", "8pm"
  m = t.match(/\b(\d{1,2})\s*(am|pm)\b/);
  if (m) {
    let h = +m[1] % 12; if (m[2] === 'pm') h += 12;
    const at = [...grp(F.filter(r => r.h === h), r => r.ch + '||' + r.p).values()].sort((a, b) => avgT(b) - avgT(a)).slice(0, 8);
    if (!at.length) return { html: `Nothing airs at ${hl(h)} in the current filter.` };
    return { html: `<b>Top programs starting at ${hl(h)}</b> (current filter):` + li(at.map(g => { const [c, p] = g.k.split('||'); return `${pn(p)} on ${chName(c)}: TVR ${nf(avgT(g), 2)}, reach ${nf(avgR(g), 1)}% (${g.n} airings)`; })), action: { h0: h, h1: h } };
  }
  if (/best.*(day|days)|which day/.test(t)) {
    const d = [...grp(F, r => r.day).values()].sort((a, b) => avgR(b) - avgR(a));
    return { html: `<b>Best days by average reach:</b>` + li(d.slice(0, 3).map(g => `${g.k}: ${nf(avgR(g), 2)}% reach, average TVR ${nf(avgT(g), 2)}, peak ${nf(g.mx, 1)}`)) + `Weakest: ${d[d.length - 1].k}.`, action: { days: d.slice(0, 3).map(g => g.k) } };
  }
  if (/best.*(slot|hour|time)|when/.test(t)) {
    const we = /weekend|saturday|sunday/.test(t);
    const rows = we ? F.filter(r => DAYS.indexOf(r.day) >= 5) : F;
    const cells = [...grp(rows, r => r.day + '|' + r.h).values()].filter(g => g.n >= 2).sort((a, b) => avgT(b) - avgT(a)).slice(0, 5);
    if (!cells.length) return { html: 'Not enough data for that question in the current filter.' };
    return { html: `<b>Best ${we ? 'weekend ' : ''}slots by average TVR:</b>` + li(cells.map(g => { const [d, h] = g.k.split('|'); return `${d} ${hl(+h)}: TVR ${nf(avgT(g), 2)}, reach ${nf(avgR(g), 1)}%`; })) };
  }
  if (/tamil/.test(t)) {
    const tam = F.filter(r => /TAMIL/.test(r.cat) || /SHAKTHI|VASANTHAM/.test(r.ch));
    if (!tam.length) return { html: 'No Tamil programming in the current filter. Open Filters and include Tamil categories or channels.' };
    const top = [...grp(tam, r => r.ch + '||' + r.p).values()].sort((a, b) => avgT(b) - avgT(a)).slice(0, 5);
    const chs = [...new Set(top.map(g => g.k.split('||')[0]))];
    return { html: `<b>Tamil option:</b> strongest Tamil-audience programs:` + li(top.map(g => { const [c, p] = g.k.split('||'); return `${pn(p)} on ${chName(c)}: TVR ${nf(avgT(g), 2)} at ${hl(g.mxr.h)}`; })) + 'Tamil channels overlap little with Sinhala channels, so they add almost fully to net reach.', action: { channels: [...new Set([...A.chs.map(c => c.ch), ...chs])] } };
  }
  m = S.CH.find(c => t.includes(c.toLowerCase().replace(' tv', '')) || t.includes(c.toLowerCase()));
  if (m && /why|weight|share|%/.test(t)) {
    const c = A.chs.find(x => x.ch === m), st = S.plan.cs.find(x => x.ch === m);
    if (!st) return { html: `${chName(m)} has no airings in the current filter.` };
    if (!c) return { html: `${chName(m)} is not in the current plan. Its score ranks #${S.plan.cs.indexOf(st) + 1} of ${S.plan.cs.length} (average reach ${nf(st.reach, 2)}%, TVR ${nf(st.tvr, 2)}).` };
    return { html: `<b>${chName(m)} gets ${nf(c.w * 100, 1)}%</b> because:` + li([`Weight = 60% average reach + 40% average TVR, rescaled across the plan channels`, `Average reach per airing ${nf(st.reach, 2)}%, average TVR ${nf(st.tvr, 2)}`, `It carries ${c.progs.length} program${c.progs.length > 1 ? 's' : ''}: ${c.progs.map(x => pn(x.p)).join(', ')}`, c.progs.some(x => x.role === 'review') ? 'One of its programs is volatile, so consider capping it' : 'Its programs rate steadily']) };
  }
  if (/dup|overlap|net reach/.test(t)) {
    return { html: `<b>Net reach ${nf(A.net, 1)}% vs gross ${nf(A.gross, 1)}%</b>: ${nf(A.loss * 100, 0)}% of gross reach is the same viewers seen on more than one channel.` + li(['Biggest overlap: Hiru and Derana (assumed 0.70)', 'Tamil channels overlap little with Sinhala channels', 'Edit the assumptions on the Duplication tab if you have measured data']) };
  }
  if (/sleeper|opportun|anomal|hidden|daytime/.test(t)) {
    const d = F.filter(r => r.h >= 6 && r.h < 17).sort((a, b) => b.tvr - a.tvr).slice(0, 5);
    return { html: `<b>Daytime opportunities</b> (high TVR outside prime time):` + li(d.map(r => `${pn(r.p)} on ${chName(r.ch)}, ${r.day.slice(0, 3)} ${r.date} at ${r.s}: TVR ${nf(r.tvr, 1)}`)) };
  }
  m = t.match(/(\d)\s*channels?/);
  if (m) { const n = Math.max(1, Math.min(S.CH.length, +m[1])); const B = simulate({ nCh: n }); return { html: `With <b>${n} channels</b> net reach would be about ${nf(B.net, 1)}% (now ${nf(A.net, 1)}%).` + li(B.chs.map(c => `${chName(c.ch)} ${nf(c.w * 100, 0)}%`)), action: { nCh: n } }; }
  // Default: plan summary
  return {
    html: `<b>Current plan:</b> ${A.chs.map(c => chName(c.ch) + ' ' + nf(c.w * 100, 0) + '%').join(', ')}. Estimated net reach ${nf(A.net, 1)}%.` +
      li([`Anchors: ${A.kept.filter(x => x.role === 'anchor').map(x => pn(x.p)).slice(0, 3).join(', ') || 'none'}`, `Try asking: "cut budget 25%", "best weekend slot", "programs at 10 PM", "why Sirasa", "add a Tamil option", "use 3 channels"`]) +
      `<p class="muted">Offline planner answers. Add a GEMINI_API_KEY on the server for full AI answers.</p>`
  };
}
