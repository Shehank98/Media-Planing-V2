// Client-ready PowerPoint deck, built in the browser with pptxgenjs (loaded on demand).
// Written for mixed audiences: every content slide has an action title, a plain-English
// "What this means" box, "How to read" captions under charts, and speaker notes.
import { DAYS, grp, avgT, avgR, steadiness, hl, band, chName, pn, nf, ni, lkr, fmtDate, uni, STRATS, DPS, daypart, creativeMix } from './engine.js';
import { buildContext, askRemote } from './ai.js';

const LIB = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const C = {
  navy: '1B2A41', teal: '0E6E78', tealLt: 'E3F1F2', tealMid: '7FB8BE', amber: 'D9822B', amberLt: 'FBEBDD',
  ink: '1F2933', muted: '5F6B7A', line: 'D9DEE5', bg: 'F4F6F9', white: 'FFFFFF', good: '2E7D4F', bad: 'B4412F'
};
const PAL = ['0E6E78', 'D9822B', '6B4FBB', 'C9475B', '3E8E4A', '2F6FD0', 'A8883A', '8A5A44', '5F7A8C'];
const FONT = 'Calibri';
const W = 13.333, M = 0.6, CW = W - 2 * M;

function loadLib() {
  if (window.PptxGenJS) return Promise.resolve();
  return new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = LIB; s.onload = res;
    s.onerror = () => rej(new Error('Could not load the PowerPoint library. Check your internet connection.'));
    document.head.appendChild(s);
  });
}
const strip = h => String(h || '').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
const short = (s, n) => { s = String(s || ''); return s.length > n ? s.slice(0, n - 1).trimEnd() + '…' : s; };
const addDays = (iso, n) => { const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
const dShort = iso => fmtDate(iso).slice(0, 6);
const mix = (a, b, t) => { const p = h => [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16)); const x = p(a), y = p(b); return x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('').toUpperCase(); };
const pctTxt = v => nf(v, 1) + '%';
const cn = ch => chName(ch);

/* ---------- facts used across slides ---------- */
function facts(S) {
  const F = S.F, A = S.cur;
  // Market context uses every channel and hour in the ratings period, not just the brief.
  const MK = S.ROWS.filter(r => r.date >= S.FS.from && r.date <= S.FS.to);
  const hrs = [...grp(MK, r => r.h).values()].filter(g => g.n >= 3 && g.k >= 6).sort((a, b) => a.k - b.k);
  const peak = [...hrs].sort((a, b) => avgT(b) - avgT(a))[0];
  const days = [...grp(F, r => r.day).values()].sort((a, b) => DAYS.indexOf(a.k) - DAYS.indexOf(b.k));
  const daysByR = [...days].sort((a, b) => avgR(b) - avgR(a));
  const wd = F.filter(r => DAYS.indexOf(r.day) < 5), we = F.filter(r => DAYS.indexOf(r.day) >= 5);
  const wdT = wd.length ? wd.reduce((s, r) => s + r.tvr, 0) / wd.length : 0, weT = we.length ? we.reduce((s, r) => s + r.tvr, 0) / we.length : 0;
  const totG = MK.reduce((s, r) => s + r.tvr, 0) || 1;
  const primeShare = MK.filter(r => daypart(r.h) === DPS[2]).reduce((s, r) => s + r.tvr, 0) / totG * 100;
  const topProg = [...grp(F, r => r.ch + '||' + r.p).values()].filter(g => g.n >= 2).sort((a, b) => avgT(b) - avgT(a))[0];
  const endDate = addDays(S.sched.start, A.W * 7 - 1);
  return { hrs, peak, days, daysByR, wdT, weT, primeShare, topProg, endDate, cs: S.plan.cs };
}

/* ---------- plain-English commentary (rules), optionally replaced by Gemini ---------- */
function ruleText(S, A, X) {
  const P = S.P, top = A.chs[0];
  const anc = A.kept.filter(x => x.role === 'anchor');
  const why = {};
  A.chs.forEach(c => {
    const st = X.cs.find(x => x.ch === c.ch) || { reach: 0, tvr: 0 };
    const best = [...c.progs].sort((a, b) => b.mean - a.mean)[0];
    why[c.ch] = c === top ? `Largest audience for the money. Programmes such as ${pn(best.p)} reach about ${nf(best.rp, 0)}% of viewers each airing.`
      : st.tvr < 2 ? 'Smaller audiences, but cheaper spots and viewers the bigger channels miss, so it adds new people at low cost.'
        : `Strong programmes (${pn(best.p)}, rating ${nf(best.mean, 1)}) that add viewers ${cn(top.ch)} does not reach.`;
  });
  const sl = (STRATS[P.strategy] || { label: 'Custom' }).label;
  return {
    headline: `A ${A.chs.length}-channel TV plan that reaches about ${nf(A.net, 0)}% of viewers, with ${nf(A.r3, 0)}% seeing the ad three or more times`,
    summary: [
      `Spend ${lkr(A.spent)} on ${A.kept.length} programmes across ${A.chs.map(c => cn(c.ch)).join(', ')} over ${A.W} weeks.`,
      anc.length ? `The plan is anchored on proven, steady shows such as ${anc.slice(0, 3).map(x => pn(x.p)).join(', ')}.` : 'The plan favours shows whose audiences are steady from week to week.',
      'Money is balanced between big shows (to reach many people) and cheaper slots (to repeat the message), instead of only the most expensive programmes.'
    ],
    channelWhy: why,
    soWhat: {
      market: X.peak ? `Across all channels, viewing builds through the evening and peaks at ${hl(X.peak.k)}. Prime time (6–10 PM) carries ${nf(X.primeShare, 0)}% of all TV viewing, so that is where most of the budget goes.` : 'Viewing is spread across the day.',
      days: X.wdT && X.weT ? `Weekday programmes rate ${nf(Math.abs((X.weT / X.wdT - 1) * 100), 0)}% ${X.weT < X.wdT ? 'higher' : 'lower'} than weekend ones on average, mainly because the popular teledramas run Monday to Friday. ${X.daysByR[0].k}${X.daysByR[1] ? ' and ' + X.daysByR[1].k : ''} reach the most people.` : `${X.daysByR[0].k} reaches the most people in this brief.`,
      channels: `${cn(X.cs[0].ch)} and ${cn((X.cs[1] || X.cs[0]).ch)} reach the most viewers per programme. Smaller channels still matter: they are cheaper and reach some people the leaders miss.`,
      split: `${cn(top.ch)} leads with ${nf(top.w * 100, 0)}% of spend. The rest is spread so that each channel brings new viewers, rather than paying twice for the same people.`,
      tiers: `${sl} mix: about ${nf(A.tiers[0].actual, 0)}% goes to the biggest shows for fast reach, ${nf(A.tiers[1].actual, 0)}% to steady mid-sized shows for value, and ${nf(A.tiers[2].actual, 0)}% to cheaper slots that remind people again.`,
      basket: `${A.kept.length} programmes across ${A.chs.length} channels. No single programme takes more than ${nf(Math.max(...A.kept.map(x => x.k)) * 100, 0)}% of the budget, which spreads the risk if one show loses viewers.`,
      reach: `About ${nf(A.net, 0)} in every 100 TV viewers will see the ad at least once, and about ${nf(A.r3, 0)} in 100 will see it three or more times, which is usually what it takes for a message to be remembered.`,
      curve: '',
      flight: `${ni(A.spots)} spots over ${A.W} weeks, ${P.pacing === 'burst' ? 'heavier at the start to build awareness quickly, then tapering' : P.pacing === 'pulse' ? 'in on and off weeks to stretch the budget over a longer period' : 'evenly each week for a steady presence'}. No programme runs more than ${P.capWk} times a week.`,
      budget: ''
    }
  };
}
async function aiText(S) {
  if (!S.aiOn) return null;
  const q = `Write client presentation text for the current plan for a NON-TECHNICAL audience (marketing managers). Plain English, short sentences, explain any TV term you use. Reply with JSON only, no markdown fences:
{"headline":"one-sentence recommendation","summary":["3 short points on why this plan"],"channelWhy":{"<exact channel name>":"one plain sentence why it gets its share"},
"soWhat":{"market":"","days":"","channels":"","split":"","tiers":"","basket":"","reach":"","curve":"","flight":"","budget":""},
"risks":[{"risk":"","why":"","action":""}],"nextSteps":["3-4 actions the agency will take"]}
Each soWhat value: 1-2 plain sentences on what that part of the plan means for the client. Use only numbers from the summary. Channel names must match currentPlan.split exactly.`;
  try {
    const txt = await Promise.race([askRemote(q, buildContext(S), []), new Promise((_, r) => setTimeout(() => r(new Error('timeout')), 40000))]);
    const j = JSON.parse(txt.replace(/^```(?:json)?|```$/gm, '').trim().replace(/ACTION:.*$/m, ''));
    return j && j.headline ? j : null;
  } catch (e) { return null; }
}
function riskRows(S, A) {
  const out = [], P = S.P, top = A.chs[0];
  if (A.unspent > A.B * .02) out.push({ risk: `${lkr(A.unspent)} cannot be placed`, why: 'The weekly limits leave no more good slots to buy.', action: 'Allow more spots per programme per week, or add programmes.' });
  const vol = A.kept.filter(x => x.role === 'review');
  if (vol.length) out.push({ risk: `Uneven ratings: ${short(vol.slice(0, 2).map(x => pn(x.p)).join(', '), 50)}`, why: 'Their audience size swings from week to week, so delivery is less certain.', action: 'Keep their share small, review ratings weekly and switch to a backup show if they drop.' });
  if (top.w > .35) out.push({ risk: `${cn(top.ch)} carries ${nf(top.w * 100, 0)}% of spend`, why: 'A ratings dip on one channel would affect a large part of the plan.', action: `Hold a backup slot on ${cn((A.chs[1] || top).ch)} that can be used quickly.` });
  if (A.loss > .35) out.push({ risk: 'Overlapping audiences', why: `${nf(A.loss * 100, 0)}% of the combined channel audiences are the same people.`, action: 'Keep the recommended mix; avoid adding more spend to the biggest channels.' });
  if (P.target > 0 && A.net < P.target) out.push({ risk: `Reach target of ${P.target}% not met`, why: `The plan reaches about ${nf(A.net, 0)}%.`, action: 'Increase the budget or add a channel with a different audience.' });
  out.push({ risk: 'Programme changes', why: 'Dramas end, new seasons start and special events move audiences.', action: 'Re-check the latest ratings before each booking week.' });
  out.push({ risk: 'Rates and availability', why: 'The plan uses an estimated cost per rating point, not confirmed channel prices.', action: 'Confirm rate cards and slot availability before booking.' });
  return out.slice(0, 5);
}

/* ---------- deck ---------- */
export async function buildDeck(S, simulate, opts = {}) {
  await loadLib();
  const A = S.cur, P = S.P, F = S.F, FS = S.FS;
  if (!A || !S.sched) throw new Error('The plan is empty. Add channels or programmes first.');
  const X = facts(S);
  const rule = ruleText(S, A, X);
  const ai = opts.useAI === false ? null : await aiText(S);
  const T = Object.assign({}, rule, ai || {});
  T.summary = Array.isArray(T.summary) && T.summary.length ? T.summary : rule.summary;
  T.channelWhy = Object.assign({}, rule.channelWhy, (ai && ai.channelWhy) || {});
  T.soWhat = Object.assign({}, rule.soWhat, (ai && ai.soWhat) || {});
  Object.keys(rule.soWhat).forEach(k => { if (!T.soWhat[k]) T.soWhat[k] = rule.soWhat[k]; });
  const risks = (ai && Array.isArray(ai.risks) && ai.risks.length && ai.risks[0].risk ? ai.risks : riskRows(S, A)).slice(0, 5);
  const client = (opts.client || '').trim(), campaign = (opts.campaign || '').trim() || 'TV Media Plan', by = (opts.by || '').trim();
  const src = ai ? 'Commentary drafted by Planner AI (Gemini) from the plan numbers; please review before sending.' : 'Commentary written from the plan numbers.';

  const pres = new window.PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = campaign; pres.author = by || 'TV Media Planner'; pres.subject = 'TV media plan';
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  const footer = (client ? client + '  ·  ' : '') + campaign + '  ·  Confidential';

  pres.defineSlideMaster({ title: 'COVER', background: { color: C.navy } });
  pres.defineSlideMaster({ title: 'DIVIDER', background: { color: C.navy } });
  pres.defineSlideMaster({
    title: 'CONTENT', background: { color: C.white },
    objects: [
      { text: { text: footer, options: { x: M, y: 7.05, w: 9, h: 0.3, fontSize: 9, color: C.muted, fontFace: FONT, margin: 0 } } },
      { placeholder: { options: { name: 'title', type: 'title', x: M, y: 0.68, w: CW, h: 0.95, fontSize: 24, bold: true, color: C.navy, fontFace: FONT, valign: 'top', align: 'left', margin: 0 }, text: '' } }
    ],
    slideNumber: { x: W - M - 0.6, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: C.muted, align: 'right', fontFace: FONT }
  });

  /* ---------- helpers ---------- */
  let section = '';
  const T_ = (s, t, o) => s.addText(t, Object.assign({ fontFace: FONT, color: C.ink, margin: 0, isTextBox: true, valign: 'top' }, o));
  const content = (kicker, title) => {
    const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: section });
    T_(s, kicker.toUpperCase(), { x: M, y: 0.38, w: CW, h: 0.28, fontSize: 11, bold: true, color: C.teal, charSpacing: 2 });
    s.addText(title, { placeholder: 'title' });
    return s;
  };
  const box = (s, x, y, w, h, fill = C.bg) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.06 });
  const soWhat = (s, x, y, w, h, text) => {
    box(s, x, y, w, h, C.tealLt);
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.18, w: 0.32, h: 0.32, fill: { color: C.teal }, line: { color: C.teal } });
    T_(s, 'i', { x: x + 0.2, y: y + 0.18, w: 0.32, h: 0.32, fontSize: 13, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Cambria' });
    T_(s, 'What this means', { x: x + 0.65, y: y + 0.18, w: w - 0.85, h: 0.32, fontSize: 13, bold: true, color: C.teal, valign: 'middle' });
    T_(s, short(strip(text), 420), { x: x + 0.25, y: y + 0.6, w: w - 0.5, h: h - 0.72, fontSize: 12.5, lineSpacingMultiple: 1.05 });
  };
  const howTo = (s, x, y, w, text) => T_(s, 'How to read this: ' + text, { x, y, w, h: 0.45, fontSize: 10, italic: true, color: C.muted });
  const kpi = (s, x, y, w, h, big, label, sub, color = C.teal) => {
    box(s, x, y, w, h);
    T_(s, big, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.72, fontSize: 32, bold: true, color });
    T_(s, label, { x: x + 0.25, y: y + 0.88, w: w - 0.5, h: 0.3, fontSize: 13, bold: true });
    if (sub) T_(s, sub, { x: x + 0.25, y: y + 1.18, w: w - 0.5, h: Math.max(0.25, h - 1.25), fontSize: 10.5, color: C.muted });
  };
  const bullets = (s, items, o) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
    Object.assign({ fontSize: 13, color: C.ink, fontFace: FONT, paraSpaceAfter: 6, valign: 'top', margin: 0.02, isTextBox: true }, o));
  const numbered = (s, items, x, y, w, rowH) => items.forEach((it, i) => {
    const yy = y + i * rowH;
    s.addShape(pres.shapes.OVAL, { x, y: yy, w: 0.42, h: 0.42, fill: { color: C.teal }, line: { color: C.teal } });
    T_(s, String(i + 1), { x, y: yy, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: C.white, align: 'center', valign: 'middle' });
    T_(s, it[0], { x: x + 0.6, y: yy + (it[1] ? -0.02 : 0.02), w: w - 0.6, h: it[1] ? 0.32 : rowH - 0.1, fontSize: it[1] ? 14 : 13, bold: !!it[1], color: it[1] ? C.navy : C.ink });
    if (it[1]) T_(s, it[1], { x: x + 0.6, y: yy + 0.32, w: w - 0.6, h: rowH - 0.38, fontSize: 11.5, color: C.muted });
  });
  const chartBase = { catAxisLabelColor: C.muted, valAxisLabelColor: C.muted, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10, valGridLine: { color: C.line, size: 0.5 }, catGridLine: { style: 'none' }, dataLabelFontFace: FONT, dataLabelFontSize: 9, dataLabelColor: C.ink, titleFontFace: FONT, titleColor: C.ink, titleFontSize: 12, legendFontFace: FONT, legendFontSize: 10, catAxisLineColor: C.line };
  const th = t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: 10.5, valign: 'middle' } });
  const tblOpt = o => Object.assign({ fontFace: FONT, fontSize: 10.5, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [3, 5, 3, 5] }, o);
  const chIdx = ch => Math.max(0, S.CH.indexOf(ch)) % PAL.length;
  const motif = (s, x, y, sc = 1) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: 2.6 * sc, h: 2.6 * sc, fill: { color: C.teal, transparency: 35 }, line: { color: C.teal, transparency: 100 } });
    s.addShape(pres.shapes.OVAL, { x: x + 1.5 * sc, y: y + 0.4 * sc, w: 2.6 * sc, h: 2.6 * sc, fill: { color: C.amber, transparency: 45 }, line: { color: C.amber, transparency: 100 } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.75 * sc, y: y + 1.6 * sc, w: 2.6 * sc, h: 2.6 * sc, fill: { color: C.tealMid, transparency: 55 }, line: { color: C.tealMid, transparency: 100 } });
  };
  const divider = (num, title, question, notes) => {
    section = title; pres.addSection({ title });
    const s = pres.addSlide({ masterName: 'DIVIDER', sectionTitle: title });
    T_(s, num, { x: M, y: 2.0, w: 3, h: 1.1, fontSize: 60, bold: true, color: C.amber });
    T_(s, title, { x: M, y: 3.1, w: 8, h: 0.9, fontSize: 36, bold: true, color: C.white });
    T_(s, question, { x: M, y: 4.05, w: 8, h: 0.8, fontSize: 18, color: 'C9D6E8' });
    motif(s, 9.0, 1.9, 0.9);
    s.addNotes(notes || question);
  };

  /* ===== COVER ===== */
  section = 'Introduction'; pres.addSection({ title: section });
  let s = pres.addSlide({ masterName: 'COVER', sectionTitle: section });
  motif(s, 8.7, 1.2, 1.2);
  T_(s, 'TV MEDIA PLAN', { x: M, y: 1.2, w: 7, h: 0.35, fontSize: 13, bold: true, color: C.tealMid, charSpacing: 3 });
  T_(s, short(campaign, 60), { x: M, y: 1.65, w: 7.8, h: 1.5, fontSize: 40, bold: true, color: C.white, valign: 'bottom' });
  T_(s, short(strip(T.headline), 170), { x: M, y: 3.3, w: 7.6, h: 1.1, fontSize: 17, color: 'C9D6E8' });
  const cov = [client ? ['Prepared for', client] : null, ['Campaign period', `${fmtDate(S.sched.start)} – ${fmtDate(X.endDate)} (${A.W} weeks)`], ['Budget', lkr(A.B)], by ? ['Prepared by', by] : null, ['Date', new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' })]].filter(Boolean);
  cov.forEach((r, i) => {
    T_(s, r[0].toUpperCase(), { x: M, y: 4.85 + i * 0.42, w: 2.1, h: 0.3, fontSize: 10, bold: true, color: C.tealMid, charSpacing: 1 });
    T_(s, r[1], { x: M + 2.1, y: 4.83 + i * 0.42, w: 5.6, h: 0.32, fontSize: 13, color: C.white });
  });
  s.addNotes(`Welcome. This deck recommends how to spend the TV budget for ${campaign}. It explains where the audience is, where the money goes and why, how many people will see the ad, the weekly schedule, budget options, risks and next steps. ${src}`);

  /* ===== AGENDA ===== */
  s = content('Agenda', 'What we will cover');
  [['01', 'The audience', 'When people watch, which days and which channels'],
    ['02', 'Our recommendation', 'How the budget is split across channels, tiers and programmes'],
    ['03', 'Reach and schedule', 'How many people will see the ad, how often, and the week-by-week plan'],
    ['04', 'Budget options and risks', 'What happens if the budget changes, and what could go wrong'],
    ['05', 'Next steps', 'Decisions needed and the booking timeline'],
    ['A', 'Appendix', 'Method, assumptions and the full programme list']].forEach((a, i) => {
    const col = i % 2, row = Math.floor(i / 2), x = M + col * (CW / 2 + 0.15), y = 1.95 + row * 1.6, w = CW / 2 - 0.15;
    box(s, x, y, w, 1.35);
    T_(s, a[0], { x: x + 0.3, y: y + 0.25, w: 0.9, h: 0.8, fontSize: 30, bold: true, color: C.amber });
    T_(s, a[1], { x: x + 1.25, y: y + 0.25, w: w - 1.5, h: 0.4, fontSize: 17, bold: true, color: C.navy });
    T_(s, a[2], { x: x + 1.25, y: y + 0.68, w: w - 1.5, h: 0.6, fontSize: 12, color: C.muted });
  });
  T_(s, 'The executive summary and a short glossary of TV terms follow this page.', { x: M, y: 6.6, w: CW, h: 0.3, fontSize: 11, italic: true, color: C.muted });
  s.addNotes('A quick map of the presentation. Technical details are in the appendix for anyone who wants them.');

  /* ===== EXECUTIVE SUMMARY ===== */
  s = content('Executive summary', short(strip(T.headline), 120));
  const kw = (7.6 - 0.4) / 3;
  kpi(s, M, 1.85, kw, 1.75, nf(A.net, 0) + '%', 'People reached', 'see the ad at least once', C.teal);
  kpi(s, M + kw + 0.2, 1.85, kw, 1.75, nf(A.r3, 0) + '%', 'Reached 3+ times', 'enough repetition to be remembered', C.amber);
  kpi(s, M + 2 * (kw + 0.2), 1.85, kw, 1.75, lkr(A.spent).replace('LKR ', ''), 'Budget (LKR)', `${ni(A.spots)} spots over ${A.W} weeks`, C.navy);
  T_(s, 'Why this plan works', { x: M, y: 3.85, w: 7.6, h: 0.35, fontSize: 15, bold: true, color: C.navy });
  numbered(s, T.summary.slice(0, 3).map(t => [short(strip(t), 150), '']), M, 4.3, 7.6, 0.85);
  box(s, 8.55, 1.85, W - M - 8.55, 3.1);
  T_(s, 'Recommended channel split', { x: 8.8, y: 2.0, w: 3.8, h: 0.35, fontSize: 13, bold: true, color: C.navy });
  A.chs.slice(0, 6).forEach((c, i) => {
    const y = 2.5 + i * 0.4, bw = (W - M - 8.8 - 1.95) * c.w / A.chs[0].w;
    T_(s, short(cn(c.ch), 15), { x: 8.8, y, w: 1.4, h: 0.3, fontSize: 11, valign: 'middle' });
    s.addShape(pres.shapes.RECTANGLE, { x: 10.25, y: y + 0.06, w: Math.max(0.05, bw), h: 0.2, fill: { color: PAL[chIdx(c.ch)] }, line: { color: PAL[chIdx(c.ch)] } });
    T_(s, nf(c.w * 100, 0) + '%', { x: 10.3 + bw, y, w: 0.6, h: 0.3, fontSize: 11, bold: true, valign: 'middle' });
  });
  const today = new Date().toISOString().slice(0, 10);
  const approve = addDays(S.sched.start, -14), bookBy = addDays(S.sched.start, -7);
  const when = d => d < today ? 'as soon as possible' : fmtDate(d);
  const urgent = approve < today;
  box(s, 8.55, 5.1, W - M - 8.55, 1.75, C.amberLt);
  T_(s, 'Decision needed', { x: 8.8, y: 5.25, w: 3.8, h: 0.32, fontSize: 13, bold: true, color: C.amber });
  T_(s, urgent ? `Approve the channel split and budget as soon as possible. The campaign starts on ${fmtDate(S.sched.start)}, so booking time is short.` : `Approve the channel split and budget by ${fmtDate(approve)}, so the main programmes can be booked by ${fmtDate(bookBy)}.`, { x: 8.8, y: 5.6, w: W - M - 9.05, h: 1.15, fontSize: 12 });
  s.addNotes(`Headline: ${strip(T.headline)}. Three numbers to remember: about ${nf(A.net, 0)}% of TV viewers will see the ad at least once; ${nf(A.r3, 0)}% will see it three or more times; and the plan uses ${lkr(A.spent)}. ${T.summary.map(strip).join(' ')} We need approval ${urgent ? 'as soon as possible' : 'by ' + fmtDate(approve)} to secure the best slots.`);

  /* ===== GLOSSARY ===== */
  const tp = X.topProg, tpn = tp ? pn(tp.k.split('||')[1]) : 'a top drama';
  s = content('Key terms', 'The TV planning words used in this deck, in plain English');
  [['TVR (rating)', 'The share of TV viewers watching a programme at an average moment.', tp ? `${tpn} has a TVR of ${nf(avgT(tp), 1)}: about ${nf(avgT(tp), 0)} in 100 viewers are watching.` : ''],
    ['Reach', 'The share of people who see the ad at least once.', `This plan: about ${nf(A.net, 0)} in 100 viewers.`],
    ['Frequency', 'How many times, on average, a reached person sees the ad.', `This plan: about ${nf(A.freq, 1)} times.`],
    ['Reach 3+', 'The share who see the ad at least three times. Repetition helps people remember.', `This plan: ${nf(A.r3, 0)} in 100 viewers.`],
    ['Spot', 'One placement of the ad in a programme\'s ad break.', `This plan: ${ni(A.spots)} spots over ${A.W} weeks.`],
    ['Duplication', 'People who watch more than one of our channels. We count them only once.', `${nf(A.loss * 100, 0)}% of the combined channel audiences overlap.`],
    ['Daypart', 'A part of the day: morning, daytime, prime time (6–10 PM) or late night.', `Prime time holds ${nf(X.primeShare, 0)}% of all TV viewing.`],
    ['Tier', 'A group of programmes by audience size: big (1), mid-sized (2), smaller (3).', `Tier 1 shows have a rating of ${nf(A.thr.t1, 1)} or more.`]].forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2), x = M + col * (CW / 2 + 0.15), y = 1.85 + row * 1.25, w = CW / 2 - 0.15;
    box(s, x, y, w, 1.1);
    T_(s, t[0], { x: x + 0.25, y: y + 0.14, w: 2.0, h: 0.35, fontSize: 14, bold: true, color: C.teal });
    T_(s, t[1], { x: x + 2.35, y: y + 0.12, w: w - 2.55, h: 0.52, fontSize: 11.5 });
    T_(s, t[2], { x: x + 2.35, y: y + 0.66, w: w - 2.55, h: 0.38, fontSize: 10.5, italic: true, color: C.muted });
  });
  s.addNotes('Before the numbers, here are the eight terms we use. The examples come from this plan, so they connect directly to the slides that follow.');

  /* ===== 01 AUDIENCE ===== */
  divider('01', 'The audience', 'When are people watching, on which days, and on which channels?');

  s = content('When people watch', X.peak ? `Viewing peaks at ${hl(X.peak.k)}; prime time carries ${nf(X.primeShare, 0)}% of all TV viewing` : 'When people watch TV');
  s.addChart(pres.charts.BAR, [{ name: 'Average rating (TVR)', labels: X.hrs.map(g => hl(g.k)), values: X.hrs.map(g => +avgT(g).toFixed(2)) }],
    Object.assign({}, chartBase, { x: M, y: 1.8, w: 7.9, h: 4.5, barDir: 'col', chartColors: X.hrs.map(g => g.k >= 18 && g.k < 22 ? C.teal : 'A9C9CC'), showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', showLegend: false, showTitle: true, title: 'Average rating by programme start time, all channels', valAxisMinVal: 0 }));
  howTo(s, M, 6.4, 7.9, 'taller bars mean more people watching at that time. Dark bars are prime time (6–10 PM).');
  soWhat(s, 8.8, 1.8, W - M - 8.8, 2.6, T.soWhat.market);
  box(s, 8.8, 4.6, W - M - 8.8, 1.75);
  T_(s, 'Top three hours', { x: 9.05, y: 4.72, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: C.navy });
  bullets(s, [...X.hrs].sort((a, b) => avgT(b) - avgT(a)).slice(0, 3).map(g => `${band(g.k)}: average rating ${nf(avgT(g), 1)}`), { x: 9.05, y: 5.08, w: W - M - 9.3, h: 1.2, fontSize: 11.5 });
  s.addNotes(`This chart shows how many people watch at each hour. ${strip(T.soWhat.market)}`);

  s = content('Which days work best', `${X.daysByR[0].k}${X.daysByR[1] ? ' and ' + X.daysByR[1].k : ''} reach the most people`);
  s.addChart(pres.charts.BAR, [{ name: 'Average reach %', labels: X.days.map(d => d.k.slice(0, 3)), values: X.days.map(d => +avgR(d).toFixed(2)) }],
    Object.assign({}, chartBase, { x: M, y: 1.8, w: 7.9, h: 4.5, barDir: 'col', chartColors: X.days.map(d => X.daysByR.slice(0, 2).includes(d) ? C.amber : C.teal), showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', showLegend: false, showTitle: true, title: 'Average reach per programme airing, by day (%)', valAxisMinVal: 0 }));
  howTo(s, M, 6.4, 7.9, 'each bar is the share of viewers an average programme reaches on that day. Orange bars are the two best days.');
  soWhat(s, 8.8, 1.8, W - M - 8.8, 2.6, T.soWhat.days);
  if (X.wdT && X.weT) {
    const hw = (W - M - 8.8 - 0.2) / 2;
    kpi(s, 8.8, 4.6, hw, 1.75, nf(X.wdT, 1), 'Weekdays', 'average rating', C.teal);
    kpi(s, 8.8 + hw + 0.2, 4.6, hw, 1.75, nf(X.weT, 1), 'Weekends', 'average rating', C.amber);
  }
  s.addNotes(`Day-of-week view. ${strip(T.soWhat.days)}`);

  const cs = X.cs.slice(0, 9), inPlan = new Set(A.chs.map(c => c.ch));
  s = content('The channel landscape', `${cn(cs[0].ch)} and ${cn((cs[1] || cs[0]).ch)} reach the most viewers; smaller channels add people the leaders miss`);
  s.addChart(pres.charts.BAR, [
    { name: 'Average reach %', labels: cs.map(c => cn(c.ch)), values: cs.map(c => +c.reach.toFixed(1)) },
    { name: 'Average rating (TVR)', labels: cs.map(c => cn(c.ch)), values: cs.map(c => +c.tvr.toFixed(1)) }
  ], Object.assign({}, chartBase, { x: M, y: 1.8, w: 7.9, h: 4.5, barDir: 'bar', barGrouping: 'clustered', chartColors: [C.amber, C.teal], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', showLegend: true, legendPos: 'b', catAxisOrientation: 'maxMin', showTitle: true, title: 'Average audience per programme airing', valAxisMinVal: 0 }));
  howTo(s, M, 6.4, 7.9, 'longer bars mean a bigger audience. Reach counts different people; rating (TVR) is the average audience at any moment.');
  soWhat(s, 8.8, 1.8, W - M - 8.8, 2.3, T.soWhat.channels);
  s.addTable([[th('Channel'), th('In our plan')], ...cs.map(c => [cn(c.ch), { text: inPlan.has(c.ch) ? 'Yes' : '—', options: { bold: inPlan.has(c.ch), color: inPlan.has(c.ch) ? C.good : C.muted } }])],
    tblOpt({ x: 8.8, y: 4.3, w: W - M - 8.8, colW: [2.4, W - M - 8.8 - 2.4], fontSize: 10, rowH: 0.22 }));
  s.addNotes(`Channel landscape. ${strip(T.soWhat.channels)}`);

  /* ===== 02 RECOMMENDATION ===== */
  divider('02', 'Our recommendation', 'Where should the money go, and why?');

  s = content('How we built the plan', 'Four steps turn the ratings data into a plan that buys the most viewers for the money');
  numbered(s, [
    ['Study the audience', `We analysed ${ni(F.length)} programme airings from ${fmtDate(FS.from)} to ${fmtDate(FS.to)} to see who watches what, and when.`],
    ['Group programmes into tiers', 'Big shows (Tier 1) build reach fast; steady mid-sized shows (Tier 2) give value; cheaper slots (Tier 3) repeat the message.'],
    ['Buy one spot at a time', `Each next spot goes where it adds the most new viewers per rupee, within limits such as ${P.capWk} spots per programme per week.`],
    ['Count each viewer once', 'People who watch several channels are counted only once, so the reach figure is not inflated.']
  ], M, 1.9, 6.6, 1.2);
  const brief = [
    ['Channels considered', FS.ch.size === S.CH.length ? `All ${S.CH.length}` : [...FS.ch].map(cn).join(', ')],
    ['Programme types', FS.cat.size === S.CATS.length ? 'All categories' : short([...FS.cat].map(pn).join(', '), 70)],
    ['Days', FS.day.size === 7 ? 'All week' : DAYS.filter(d => FS.day.has(d)).map(d => d.slice(0, 3)).join(', ')],
    ['Time band', `${hl(FS.h0)} – ${hl((FS.h1 + 1) % 24)}`],
    ['Budget', lkr(A.B)],
    ['Campaign', `${A.W} weeks from ${fmtDate(S.sched.start)}`],
    ['Creatives', creativeMix(P).map(c => `${c.name} ${c.dur}s`).join(', ')],
    ['Strategy', (STRATS[P.strategy] || { label: 'Custom' }).label + ' tier mix']
  ];
  T_(s, 'The brief', { x: 7.7, y: 1.9, w: 5, h: 0.35, fontSize: 15, bold: true, color: C.navy });
  s.addTable(brief.map(r => [{ text: r[0], options: { color: C.muted } }, { text: r[1], options: { bold: true } }]), tblOpt({ x: 7.7, y: 2.35, w: W - M - 7.7, colW: [1.9, W - M - 7.7 - 1.9], rowH: 0.42, fontSize: 11 }));
  s.addNotes('This is how the recommendation was produced, step by step. The brief on the right sets the boundaries we planned within.');

  s = content('Channel split', `Where the budget goes: ${A.chs.map(c => cn(c.ch).replace(' TV', '') + ' ' + nf(c.w * 100, 0) + '%').join(', ')}`);
  s.addChart(pres.charts.DOUGHNUT, [{ name: 'Share', labels: A.chs.map(c => cn(c.ch)), values: A.chs.map(c => +(c.w * 100).toFixed(1)) }],
    { x: M, y: 1.75, w: 4.2, h: 4.0, holeSize: 55, chartColors: A.chs.map(c => PAL[chIdx(c.ch)]), showPercent: true, showValue: false, dataLabelColor: C.white, dataLabelFontSize: 11, dataLabelFontFace: FONT, showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 10 });
  s.addTable([[th('Channel'), th('Share'), th('Budget (LKR)'), th('Spots'), th('Why this channel')],
    ...A.chs.map(c => [{ text: cn(c.ch), options: { bold: true } }, nf(c.w * 100, 0) + '%', lkr(c.bud).replace('LKR ', ''), String(c.spots), short(strip(T.channelWhy[c.ch] || ''), 140)])],
  tblOpt({ x: 5.0, y: 1.8, w: W - M - 5.0, colW: [1.45, 0.7, 1.15, 0.65, W - M - 5.0 - 3.95], fontSize: 10.5 }));
  soWhat(s, M, 5.85, CW, 1.1, T.soWhat.split);
  s.addNotes(A.chs.map(c => `${cn(c.ch)} gets ${nf(c.w * 100, 0)}%: ${strip(T.channelWhy[c.ch] || '')}`).join(' ') + ' ' + strip(T.soWhat.split));

  const TN = { 1: 'Peak impact', 2: 'Steady value', 3: 'Reminders' };
  const TD = { 1: 'The biggest shows. They reach many people fast.', 2: 'Reliable mid-sized shows at a better price.', 3: 'Cheaper slots that repeat the message.' };
  s = content('Tier strategy', `${(STRATS[P.strategy] || { label: 'Custom' }).label} mix: big shows for reach, steady shows for value, cheaper slots for repetition`);
  const pyW = [2.6, 4.0, 5.4], pyCol = [C.teal, '4A9CA4', 'A9C9CC'];
  [0, 1, 2].forEach(i => {
    const w = pyW[i], x = M + (5.6 - w) / 2, y = 1.95 + i * 1.3, t = A.tiers[i], txt = i === 2 ? C.navy : C.white;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.1, fill: { color: pyCol[i] }, line: { color: pyCol[i] }, rectRadius: 0.06 });
    T_(s, `Tier ${i + 1}`, { x, y: y + 0.12, w, h: 0.35, fontSize: 13, bold: true, color: txt, align: 'center' });
    T_(s, `${nf(t.actual, 0)}% of budget`, { x, y: y + 0.48, w, h: 0.45, fontSize: 18, bold: true, color: txt, align: 'center' });
  });
  A.tiers.forEach((t, i) => {
    const y = 1.95 + i * 1.3, x = 6.5;
    const progs = A.kept.filter(x2 => x2.tier === t.t).slice(0, 3).map(x2 => pn(x2.p)).join(', ');
    const rng = t.t === 1 ? `Rating ${nf(A.thr.t1, 1)} or more` : t.t === 3 ? `Rating under ${nf(A.thr.t3, 1)}` : `Rating ${nf(A.thr.t3, 1)}–${nf(A.thr.t1, 1)}`;
    T_(s, `Tier ${t.t} · ${TN[t.t]}`, { x, y, w: W - M - x, h: 0.32, fontSize: 14, bold: true, color: C.navy });
    T_(s, TD[t.t], { x, y: y + 0.33, w: W - M - x, h: 0.35, fontSize: 11.5 });
    T_(s, `${rng} · ${t.n} programme${t.n === 1 ? '' : 's'}${progs ? ': ' + short(progs, 70) : ''}`, { x, y: y + 0.68, w: W - M - x, h: 0.4, fontSize: 10.5, italic: true, color: C.muted });
  });
  soWhat(s, M, 5.8, CW, 1.15, T.soWhat.tiers);
  s.addNotes(`Think of the plan as a pyramid. ${TD[1]} ${TD[2]} ${TD[3]} ${strip(T.soWhat.tiers)} Budget that cannot be used in one tier moves to the others.`);

  const items = A.kept.slice(0, 12);
  s = content('Programme selection', `${A.kept.length} programmes selected; the top ${items.length} take ${nf(items.reduce((a, x) => a + x.k, 0) * 100, 0)}% of the budget`);
  s.addTable([[th('Programme'), th('Channel'), th('Usual time'), th('Rating (TVR)'), th('Reach per airing'), th('Ratings stability'), th('Tier'), th('Spots'), th('Budget share')],
    ...items.map(x => [{ text: short(pn(x.p), 34), options: { bold: true } }, cn(x.ch), hl(x.hour), nf(x.mean, 1), nf(x.rp, 0) + '%',
      { text: steadiness(x.ci, x.n).label, options: { color: x.ci < 3 && x.n >= 2 ? C.bad : x.ci >= 5 ? C.good : C.ink } },
      { text: `Tier ${x.tier}${x.role === 'anchor' ? ' · Anchor' : x.role === 'review' ? ' · Watch' : ''}`, options: { bold: x.role !== 'support' } }, String(x.spots), nf(x.k * 100, 1) + '%'])],
  tblOpt({ x: M, y: 1.8, w: CW, colW: [3.0, 1.55, 1.05, 1.05, 1.1, 1.45, 1.4, 0.75, 0.78], fontSize: 10.5, rowH: 0.3 }));
  howTo(s, M, 6.15, CW, '"Ratings stability" shows whether a show\'s audience is similar every week (Very steady, Steady) or swings (Mixed, Volatile). "Anchor" = a big, steady show the plan relies on; "Watch" = keep an eye on its ratings.');
  if (A.kept.length > items.length) T_(s, `The full list of ${A.kept.length} programmes is in the appendix.`, { x: M, y: 6.62, w: 8, h: 0.3, fontSize: 10, color: C.muted });
  s.addNotes(`${strip(T.soWhat.basket)} Programmes are chosen for audience size, week-to-week stability and cost. Shows marked Watch have uneven ratings, so we keep their share small.`);

  const anc = A.kept.filter(x => x.role === 'anchor').slice(0, 3);
  const show = anc.length ? anc : A.kept.slice(0, 3);
  s = content('Anchor programmes', 'The shows that carry the plan: large audiences that turn up every week');
  const cw = (CW - 0.6) / 3;
  show.forEach((x, i) => {
    const x0 = M + i * (cw + 0.3), hw = (cw - 0.5) / 2;
    box(s, x0, 1.85, cw, 3.55);
    T_(s, short(pn(x.p), 38), { x: x0 + 0.25, y: 2.0, w: cw - 0.5, h: 0.7, fontSize: 17, bold: true, color: C.navy });
    T_(s, `${cn(x.ch)} · usually ${hl(x.hour)} · ${(x.days || []).map(d => d.slice(0, 3)).join(', ')}`, { x: x0 + 0.25, y: 2.72, w: cw - 0.5, h: 0.5, fontSize: 10.5, color: C.muted });
    T_(s, nf(x.mean, 1), { x: x0 + 0.25, y: 3.2, w: hw, h: 0.7, fontSize: 30, bold: true, color: C.teal });
    T_(s, 'rating (TVR)', { x: x0 + 0.25, y: 3.88, w: hw, h: 0.3, fontSize: 10.5, color: C.muted });
    T_(s, nf(x.rp, 0) + '%', { x: x0 + 0.25 + hw, y: 3.2, w: hw, h: 0.7, fontSize: 30, bold: true, color: C.amber });
    T_(s, 'reach per airing', { x: x0 + 0.25 + hw, y: 3.88, w: hw, h: 0.3, fontSize: 10.5, color: C.muted });
    T_(s, `${steadiness(x.ci, x.n).label} ratings over ${x.n} airings. ${x.spots} spots, ${nf(x.k * 100, 0)}% of the budget.`, { x: x0 + 0.25, y: 4.35, w: cw - 0.5, h: 0.9, fontSize: 11.5 });
  });
  soWhat(s, M, 5.6, CW, 1.3, `These shows reach ${show.map(x => nf(x.rp, 0) + '%').join(', ')} of viewers each time they air, and their audiences are stable, so delivery is predictable. Together they take ${nf(show.reduce((a, x) => a + x.k, 0) * 100, 0)}% of the budget.`);
  s.addNotes('Anchor programmes are the backbone. They combine large audiences with stable ratings, so they are the safest place for a large share of the money.');

  /* ===== 03 REACH AND SCHEDULE ===== */
  divider('03', 'Reach and schedule', 'How many people will see the ad, how often, and when?');

  s = content('Reach', `About ${nf(A.net, 0)} in 100 TV viewers will see the ad; ${nf(A.r3, 0)} in 100 will see it three or more times`);
  if (A.chs.length >= 2) {
    const a = A.chs[0], b = A.chs[1], d = S.getD(a.ch, b.ch), both = a.R + b.R - uni(a.R, b.R, d);
    const mxR = Math.max(a.R, b.R); let ra = 1.0 + Math.sqrt(a.R / mxR) * 0.75, rb = 1.0 + Math.sqrt(b.R / mxR) * 0.75;
    const fit = Math.min(1, 6.3 / (ra * 2 + rb * 2 - 1.3)); ra *= fit; rb *= fit;
    const cy = 3.9, ax = M + 0.1, bx = ax + ra * 2 - 1.3;
    s.addShape(pres.shapes.OVAL, { x: ax, y: cy - ra, w: ra * 2, h: ra * 2, fill: { color: PAL[chIdx(a.ch)], transparency: 45 }, line: { color: PAL[chIdx(a.ch)], width: 1.5 } });
    s.addShape(pres.shapes.OVAL, { x: bx, y: cy - rb, w: rb * 2, h: rb * 2, fill: { color: PAL[chIdx(b.ch)], transparency: 45 }, line: { color: PAL[chIdx(b.ch)], width: 1.5 } });
    T_(s, [{ text: cn(a.ch), options: { breakLine: true } }, { text: pctTxt(a.R) }], { x: ax + 0.15, y: cy - 0.4, w: ra * 2 - 1.55, h: 0.8, fontSize: 12, bold: true, color: C.navy, align: 'center', valign: 'middle' });
    T_(s, [{ text: 'Watch both', options: { breakLine: true } }, { text: pctTxt(both) }], { x: bx, y: cy - 0.4, w: 1.3, h: 0.8, fontSize: 11, bold: true, color: C.navy, align: 'center', valign: 'middle' });
    T_(s, [{ text: cn(b.ch), options: { breakLine: true } }, { text: pctTxt(b.R) }], { x: bx + 1.4, y: cy - 0.4, w: rb * 2 - 1.5, h: 0.8, fontSize: 12, bold: true, color: C.navy, align: 'center', valign: 'middle' });
    T_(s, `Share of viewers each channel reaches with this plan. People who watch both ${cn(a.ch)} and ${cn(b.ch)} sit in the middle and are counted once. Across all ${A.chs.length} channels, ${nf(A.gross - A.net, 0)} points of overlap are removed.`, { x: M, y: 6.15, w: 6.5, h: 0.7, fontSize: 11, italic: true, color: C.muted });
    T_(s, 'Example: the two largest channels in the plan', { x: M, y: 1.8, w: 6.5, h: 0.3, fontSize: 12, bold: true, color: C.navy });
  }
  const fx = 7.4, fw = (W - M - fx - 0.4) / 3;
  [[pctTxt(A.gross), 'All channel audiences added up', C.ink], [nf(A.gross - A.net, 1) + ' pts', 'Same people, removed', C.amber], [pctTxt(A.net), 'Different people reached', C.teal]].forEach((v, i) => {
    box(s, fx + i * (fw + 0.2), 1.85, fw, 1.3);
    T_(s, v[0], { x: fx + i * (fw + 0.2) + 0.1, y: 1.97, w: fw - 0.2, h: 0.6, fontSize: 21, bold: true, color: v[2], align: 'center' });
    T_(s, v[1], { x: fx + i * (fw + 0.2) + 0.1, y: 2.6, w: fw - 0.2, h: 0.5, fontSize: 10.5, color: C.muted, align: 'center' });
  });
  const hk = (W - M - fx - 0.2) / 2;
  kpi(s, fx, 3.35, hk, 1.55, pctTxt(A.r3), 'Reached 3+ times', 'enough repetition to remember', C.amber);
  kpi(s, fx + hk + 0.2, 3.35, hk, 1.55, nf(A.freq, 1) + 'x', 'Average frequency', 'times each person sees it', C.navy);
  soWhat(s, fx, 5.1, W - M - fx, 1.8, T.soWhat.reach);
  s.addNotes(`Reach is the number of different people who see the ad. Because many people watch more than one channel, we remove the overlap so each person is counted once. ${strip(T.soWhat.reach)}`);

  const nMax = Math.min(X.cs.length, 9), pts = [];
  for (let k = 1; k <= nMax; k++) { const r = simulate({ chMode: 'auto', nCh: k, split: {} }); pts.push(r ? +r.net.toFixed(1) : 0); }
  const gains = pts.map((p, i) => i ? p - pts[i - 1] : p);
  let knee = 1; gains.forEach((g, i) => { if (i && g >= 3) knee = i + 1; });
  const nowN = A.chs.length, atNow = pts[nowN - 1] ?? A.net, top = pts[pts.length - 1];
  const curveTitle = nowN >= knee ? `Reach levels off after ${knee} channels; the plan's ${nowN} channels capture ${nf(atNow, 0)}% of a possible ${nf(top, 0)}%` : `More channels would still add reach: ${nowN} channels reach ${nf(atNow, 0)}%, up to ${nf(top, 0)}% with ${knee}`;
  if (!T.soWhat.curve) T.soWhat.curve = nowN >= knee
    ? `Up to ${knee} channels, each one brings in many new viewers. After that, extra channels add less than 3 points each, because their viewers already watch the channels in the plan. Staying at ${nowN} channels keeps the money working harder.`
    : `Adding channels up to ${knee} would still bring in noticeably more people (about ${nf(pts[knee - 1] - atNow, 0)} points). Consider this if reaching more people matters more than repetition.`;
  s = content('How many channels?', curveTitle);
  s.addChart(pres.charts.LINE, [{ name: 'People reached %', labels: pts.map((_, i) => `${i + 1} channel${i ? 's' : ''}`), values: pts }],
    Object.assign({}, chartBase, { x: M, y: 1.8, w: 7.9, h: 4.5, chartColors: [C.teal], lineSize: 3, lineDataSymbolSize: 8, showValue: true, dataLabelPosition: 't', dataLabelFormatCode: '0', showLegend: false, showTitle: true, title: 'People reached (%) with the same budget, as channels are added', valAxisMinVal: 0 }));
  howTo(s, M, 6.4, 7.9, 'each point shows how many people the same budget reaches with that many channels. Where the line goes flat, extra channels add few new people.');
  soWhat(s, 8.8, 1.8, W - M - 8.8, 2.6, T.soWhat.curve);
  box(s, 8.8, 4.6, W - M - 8.8, 1.75);
  T_(s, 'New people added by each channel', { x: 9.05, y: 4.72, w: 3.8, h: 0.3, fontSize: 12, bold: true, color: C.navy });
  bullets(s, pts.slice(1, 6).map((p, i) => `Channel ${i + 2}: +${nf(p - pts[i], 1)} points`), { x: 9.05, y: 5.08, w: W - M - 9.3, h: 1.2, fontSize: 11.5 });
  s.addNotes(`This chart answers: should we use more channels? ${strip(T.soWhat.curve)}`);

  const Sc = S.sched;
  s = content('Weekly schedule', `${ni(A.spots)} spots over ${A.W} weeks from ${fmtDate(Sc.start)}, ${P.pacing === 'burst' ? 'front-loaded to build awareness fast' : P.pacing === 'pulse' ? 'in on and off weeks' : 'spread evenly'}`);
  const mxs = Math.max(1, ...A.chs.flatMap(c => Sc.weeks.map(w => w.byCh[c.ch] || 0)));
  const calW = 8.3, wc = (calW - 1.8 - 0.8) / A.W, calFont = A.W > 8 ? 9 : 11;
  s.addTable([
    [th('Channel'), ...Sc.weeks.map((_, i) => ({ text: `Week ${i + 1}\n${dShort(addDays(Sc.start, i * 7))}`, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: calFont - 1, align: 'center' } })), th('Total')],
    ...A.chs.map(c => [{ text: cn(c.ch), options: { bold: true } }, ...Sc.weeks.map(w => { const v = w.byCh[c.ch] || 0; return { text: v ? String(v) : '–', options: { align: 'center', fill: { color: v ? mix('FFFFFF', C.teal, .15 + .7 * v / mxs) : 'FFFFFF' }, color: v / mxs > .55 ? C.white : C.ink, bold: true } }; }), { text: String(c.spots), options: { align: 'center', bold: true } }]),
    [{ text: 'All channels', options: { bold: true, fill: { color: C.bg } } }, ...Sc.weeks.map(w => ({ text: String(w.spots), options: { align: 'center', bold: true, fill: { color: C.bg } } })), { text: String(A.spots), options: { align: 'center', bold: true, fill: { color: C.bg } } }],
    [{ text: 'Budget (LKR)', options: { color: C.muted } }, ...Sc.weeks.map(w => ({ text: lkr(w.bud).replace('LKR ', ''), options: { align: 'center', color: C.muted } })), { text: lkr(A.spent).replace('LKR ', ''), options: { align: 'center', color: C.muted } }]
  ], tblOpt({ x: M, y: 1.8, w: calW, colW: [1.8, ...Sc.weeks.map(() => wc), 0.8], fontSize: calFont, rowH: 0.42 }));
  howTo(s, M, Math.min(6.3, 1.8 + 0.45 * (A.chs.length + 3) + 0.2), calW, 'each number is the count of spots booked that week. Darker cells mean more spots.');
  box(s, 9.2, 1.8, W - M - 9.2, 3.1);
  T_(s, 'Booking rules', { x: 9.45, y: 1.95, w: 3.3, h: 0.32, fontSize: 13, bold: true, color: C.navy });
  bullets(s, [
    `At most ${P.capWk} spots per programme per week, so the same viewers are not over-exposed.`,
    P.same === 'bestday' ? 'Each spot runs on the programme\'s best-rated day. When two programmes at the same hour share many viewers, they run on different days so the same people are not hit twice in one evening.' : P.same === 'roadblock' ? 'Rival channels at the same hour run on the same nights. One person cannot watch both, so this reaches more people.' : 'Rival same-hour spots run on different nights, so the same people see the ad more often.',
    'Spots run on the days each programme actually airs.'
  ], { x: 9.45, y: 2.35, w: W - M - 9.7, h: 2.5, fontSize: 11.5 });
  soWhat(s, 9.2, 5.05, W - M - 9.2, 1.85, T.soWhat.flight);
  s.addNotes(`The weekly plan. ${strip(T.soWhat.flight)} The day-by-day booking list is available as a spreadsheet from the planning tool.`);

  /* ===== 04 BUDGET AND RISKS ===== */
  divider('04', 'Budget options and risks', 'What happens if the budget changes, and what could go wrong?');

  const levels = [0, 10, 25, 40].map(c => ({ c, r: simulate({ cut: c }) })).filter(x => x.r);
  const full = levels[0].r, l25 = levels.find(l => l.c === 25);
  const budgetSo = T.soWhat.budget || (l25 ? `Cutting 25% lowers the people reached only from ${nf(full.net, 0)}% to ${nf(l25.r.net, 0)}%, but those seeing it three or more times fall from ${nf(full.r3, 0)}% to ${nf(l25.r.r3, 0)}%. A cut mainly costs repetition, which affects how well the message is remembered.` : '');
  s = content('Budget options', l25 ? `A 25% cut keeps most of the reach but reduces repetition: 3+ reach falls from ${nf(full.r3, 0)}% to ${nf(l25.r.r3, 0)}%` : 'Budget options');
  s.addChart(pres.charts.BAR, [
    { name: 'Reached at least once', labels: levels.map(l => l.c ? `−${l.c}%` : 'Full budget'), values: levels.map(l => +l.r.net.toFixed(1)) },
    { name: 'Reached 3+ times', labels: levels.map(l => l.c ? `−${l.c}%` : 'Full budget'), values: levels.map(l => +l.r.r3.toFixed(1)) }
  ], Object.assign({}, chartBase, { x: M, y: 1.8, w: 6.6, h: 4.5, barDir: 'col', barGrouping: 'clustered', chartColors: [C.teal, C.amber], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0', showLegend: true, legendPos: 'b', showTitle: true, title: 'People reached (%) at each budget level', valAxisMinVal: 0 }));
  howTo(s, M, 6.4, 6.6, 'compare the two bars at each budget. Teal = seen at least once; orange = seen three or more times.');
  s.addTable([[th('Option'), th('Budget (LKR)'), th('Reached'), th('Reached 3+'), th('Spots')],
    ...levels.map(l => [{ text: l.c ? `Cut ${l.c}%` : 'Recommended', options: { bold: !l.c } }, lkr(l.r.B).replace('LKR ', ''), nf(l.r.net, 0) + '%', nf(l.r.r3, 0) + '%', String(l.r.spots)]),
    ...(S.scenarios || []).slice(0, 3).map(sc => [short(sc.name, 22), lkr(sc.r.B).replace('LKR ', ''), nf(sc.r.net, 0) + '%', sc.r.r3 != null ? nf(sc.r.r3, 0) + '%' : '—', String(sc.r.spots)])],
  tblOpt({ x: 7.5, y: 1.8, w: W - M - 7.5, colW: [1.45, 1.1, 0.85, 0.95, W - M - 7.5 - 4.35], fontSize: 11, rowH: 0.36 }));
  soWhat(s, 7.5, 4.6, W - M - 7.5, 2.3, budgetSo);
  s.addNotes(`Budget options. ${strip(budgetSo)} If a cut is necessary, a smaller cut of about 10% protects most of the repetition.`);

  s = content('Risks and how we manage them', 'The main things that could affect delivery, and what we will do about each');
  s.addTable([[th('Risk'), th('Why it matters'), th('What we will do')],
    ...risks.map((r, i) => { const f = { color: i % 2 ? C.white : C.bg }; return [{ text: short(strip(r.risk), 70), options: { bold: true, fill: f } }, { text: short(strip(r.why), 130), options: { fill: f } }, { text: short(strip(r.action), 130), options: { fill: f, color: C.teal } }]; })],
  tblOpt({ x: M, y: 1.85, w: CW, colW: [3.4, 4.35, CW - 7.75], fontSize: 12, rowH: 0.75 }));
  s.addNotes('These are the risks we see and our plan for each. None of them change the recommendation, but they shape how we monitor the campaign.');

  /* ===== 05 NEXT STEPS ===== */
  divider('05', 'Next steps', 'What we need from you, and when');
  const endW = addDays(Sc.start, A.W * 7 + 6);
  const tl = [
    ['Approve plan', urgent ? 'This week' : fmtDate(approve), 'Sign off the split and budget'],
    ['Confirm rates', urgent ? 'Within 2 days' : fmtDate(addDays(Sc.start, -10)), 'Check rate cards and slot availability'],
    ['Book slots', urgent ? 'Before go-live' : fmtDate(bookBy), 'Secure anchor programmes first'],
    ['Go live', fmtDate(Sc.start), `${A.W}-week campaign starts`],
    ['Weekly check', 'Every Monday', 'Compare ratings and adjust weak slots'],
    ['Review', fmtDate(endW), 'Results and learnings']
  ];
  s = content('Next steps', urgent ? `Approval this week keeps the ${fmtDate(Sc.start)} start date on track` : `Approve by ${fmtDate(approve)} to book the main programmes by ${fmtDate(bookBy)}`);
  const sw = CW / tl.length;
  tl.forEach((t, i) => {
    const x = M + i * sw;
    s.addShape(pres.shapes.CHEVRON, { x, y: 2.0, w: sw + 0.05, h: 0.9, fill: { color: i === 0 ? C.amber : mix(C.teal, C.navy, i / (tl.length - 1)) }, line: { color: C.white, width: 1 } });
    T_(s, t[0], { x: x + 0.38, y: 2.0, w: sw - 0.6, h: 0.9, fontSize: 11.5, bold: true, color: C.white, align: 'center', valign: 'middle' });
    T_(s, t[1], { x: x + 0.1, y: 3.1, w: sw - 0.2, h: 0.35, fontSize: 12, bold: true, color: C.navy, align: 'center' });
    T_(s, t[2], { x: x + 0.1, y: 3.45, w: sw - 0.2, h: 0.75, fontSize: 10.5, color: C.muted, align: 'center' });
  });
  box(s, M, 4.6, CW / 2 - 0.15, 2.3, C.amberLt);
  T_(s, 'Decisions we need from you', { x: M + 0.25, y: 4.75, w: 5, h: 0.32, fontSize: 14, bold: true, color: C.amber });
  bullets(s, [`Approve the ${A.chs.length}-channel split and the ${lkr(A.spent)} budget`, `Confirm the campaign period (${fmtDate(Sc.start)} – ${fmtDate(Sc.end)}) and the creatives`, 'Tell us any programmes that must be included or avoided'], { x: M + 0.25, y: 5.15, w: CW / 2 - 0.65, h: 1.65, fontSize: 12 });
  box(s, M + CW / 2 + 0.15, 4.6, CW / 2 - 0.15, 2.3);
  T_(s, 'What we will do', { x: M + CW / 2 + 0.4, y: 4.75, w: 5, h: 0.32, fontSize: 14, bold: true, color: C.teal });
  const ours = ai && Array.isArray(ai.nextSteps) && ai.nextSteps.length >= 3 ? ai.nextSteps.slice(0, 4).map(t => short(strip(t), 110)) : ['Confirm rates and availability with each channel', 'Book anchor programmes first, then the rest', 'Send a weekly delivery update with any changes', 'Share a results report after the campaign'];
  bullets(s, ours, { x: M + CW / 2 + 0.4, y: 5.15, w: CW / 2 - 0.65, h: 1.65, fontSize: 12 });
  s.addNotes(`To secure the best slots, we need approval ${when(approve)}. After that we confirm rates, book anchor programmes ${urgent ? 'immediately' : 'by ' + fmtDate(bookBy)}, and go live on ${fmtDate(Sc.start)}.`);

  /* ===== APPENDIX ===== */
  if (opts.appendix !== false) {
    divider('A', 'Appendix', 'Method, assumptions and the full programme list', 'Reference material for anyone who wants the detail behind the numbers.');
    s = content('Appendix · Method', 'How the numbers are calculated');
    s.addTable([[th('Step'), th('What we did'), th('Why')],
      ['Data', `${ni(S.ROWS.length)} programme airings loaded; ${ni(F.length)} match the brief (${fmtDate(FS.from)} – ${fmtDate(FS.to)}).`, 'Recent, real ratings for the channels and times in the brief.'],
      ['Channel shortlist', P.chMode === 'manual' ? 'Channels chosen by the planner.' : `Top ${P.nCh} channels by 60% average reach + 40% average rating.`, 'Focus on channels that deliver audiences.'],
      ['Tiers', `Programmes averaging under ${nf(P.minTvrPlan ?? .5, 1)} rating excluded. Top ${100 - P.tp1}% by rating = Tier 1, bottom ${P.tp3}% = Tier 3. Budget split ${P.tiers.join('/')}%.`, 'Balances fast reach (big shows) with repetition (cheaper slots).'],
      ['Spot buying', `One spot at a time, wherever it adds the most new viewers per rupee, weighted by ratings stability. Max ${P.capWk} per programme per week; daypart limits applied.`, 'Avoids putting all the money into a few expensive shows.'],
      ['Cost', `Cost per spot = the higher of LKR ${ni(P.minRate || 0)} or LKR ${ni(P.cprp)} × rating × duration ÷ 30, less the channel discount.`, 'An estimate until channel rate cards are confirmed.'],
      ['Reach', 'Each person counted once across channels using overlap factors; repeat spots add fewer new viewers each time.', 'Gives a realistic, not inflated, reach figure.'],
      ['Reach 3+', 'Estimated from reach and average frequency using a standard statistical spread (Poisson).', 'A common measure of effective repetition.']],
    tblOpt({ x: M, y: 1.8, w: CW, colW: [1.8, 6.0, CW - 7.8], fontSize: 11 }));
    s.addNotes(src);

    s = content('Appendix · Assumptions', 'Planning assumptions that can be changed in the tool');
    const ab = A.chs.length >= 2 ? S.getD(A.chs[0].ch, A.chs[1].ch) : null;
    s.addTable([[th('Assumption'), th('Value used'), th('What it means')],
      ['Cost per rating point', `LKR ${ni(P.cprp)} (30 sec)`, 'Price of one rating point; replace with confirmed rates.'],
      ['Minimum spot rate', `LKR ${ni(P.minRate || 0)}`, 'Even small shows cost at least this much per spot.'],
      ['Channel overlap', ab != null ? `${cn(A.chs[0].ch)}–${cn(A.chs[1].ch)}: ${nf(ab * 100, 0)}%` : '—', 'Share of the smaller audience that also watches the other channel.'],
      ['Overlap within a channel', `${nf(S.DINTRA * 100, 0)}%`, 'Viewers of one programme who also watch others on the same channel.'],
      ['New viewers from repeat spots', `${P.repQ ?? 40}% of the previous spot`, 'Each repeat of a programme reaches fewer new people.'],
      ['Spots per programme per week', String(P.capWk), 'Limit to avoid over-exposing the same viewers.'],
      ['Daypart limits', A.dps.filter(d => d.present).map(d => `${d.d.split(' ')[0]} ${nf(d.min, 0)}–${nf(d.max, 0)}%`).join(', '), 'Share of spend allowed in each part of the day.']],
    tblOpt({ x: M, y: 1.8, w: CW, colW: [3.0, 3.8, CW - 6.8], fontSize: 11 }));
    s.addNotes('These assumptions are visible and editable in the planning tool. Measured overlap data or confirmed rate cards would replace them.');

    const per = 14;
    for (let i = 0; i < A.kept.length; i += per) {
      const chunk = A.kept.slice(i, i + per);
      s = content('Appendix · Full programme list', `Programmes ${i + 1}–${i + chunk.length} of ${A.kept.length}`);
      s.addTable([[th('Programme'), th('Channel'), th('Category'), th('Time'), th('Days'), th('TVR'), th('Reach'), th('Tier'), th('Spots'), th('Budget (LKR)')],
        ...chunk.map(x => [short(pn(x.p), 34), cn(x.ch), short(pn(x.cat), 22), hl(x.hour), (x.days || []).map(d => d.slice(0, 2)).join(' '), nf(x.mean, 1), nf(x.rp, 0) + '%', 'T' + x.tier, String(x.spots), lkr(x.bud).replace('LKR ', '')])],
      tblOpt({ x: M, y: 1.8, w: CW, colW: [2.55, 1.3, 1.75, 0.85, 1.95, 0.55, 0.65, 0.5, 0.65, CW - 10.75], fontSize: 10, rowH: 0.33 }));
    }
  }

  const name = `${campaign.replace(/[^\w-]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '')}-${new Date().toISOString().slice(0, 10)}.pptx`;
  await pres.writeFile({ fileName: name });
  return { name, ai: !!ai };
}
