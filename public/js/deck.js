// Builds a client-ready PowerPoint deck (with rationale) entirely in the browser.
// pptxgenjs is loaded on demand; native charts and tables stay editable in PowerPoint.
import { DAYS, grp, avgT, avgR, ciOf, steadiness, hl, band, chName, pn, nf, ni, lkr, fmtDate, uni, insights } from './engine.js';
import { buildContext, askRemote } from './ai.js';

const LIB = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const COL = { teal: '0E6E78', tealDk: '0A4F57', amber: 'DE8A2E', ink: '15212C', muted: '5B6A79', soft: 'F2F5F7', line: 'DBE1E7', white: 'FFFFFF', good: '2E7D4F', bad: 'B4412F', info: '3B55B5' };
const PAL = ['0E6E78', 'DE8A2E', '7A5AC8', 'C9475B', '3E8E4A', '2F6FD0', 'A8883A', '8A5A44', '5F7A8C'];
const FONT = 'Calibri';
const W = 13.333, M = 0.6;

function loadLib() {
  if (window.PptxGenJS) return Promise.resolve();
  return new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = LIB; s.onload = res;
    s.onerror = () => rej(new Error('Could not load the PowerPoint library. Check your internet connection.'));
    document.head.appendChild(s);
  });
}
const strip = h => String(h || '').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
const short = (s, n) => s.length > n ? s.slice(0, n - 1) + '…' : s;

/* ---------- rationale (Gemini if available, else written from the numbers) ---------- */
function ruleRationale(S, A) {
  const top = A.chs[0], anc = A.kept.filter(x => x.role === 'anchor');
  const why = {};
  A.chs.forEach(c => {
    const st = S.plan.cs.find(x => x.ch === c.ch);
    const a = c.progs.find(x => x.role === 'anchor');
    why[c.ch] = c === top ? `Highest blended reach and rating (avg reach ${nf(st.reach, 1)}%, TVR ${nf(st.tvr, 2)}), so it leads the plan.`
      : a ? `Carries anchor ${pn(a.p)} (TVR ${nf(a.mean, 1)}, steady). Adds reach beyond ${chName(top.ch)}.`
        : c.progs.length === 1 ? `Single program ${pn(c.progs[0].p)} adds incremental audience at a lower weight.`
          : `Adds incremental reach; ${c.progs.length} programs spread the risk.`;
  });
  const dupNote = A.chs.length > 1 ? `${nf(A.loss * 100, 0)}% of gross reach overlaps across channels, so weights favour channels that add new viewers rather than repeat them.` : '';
  return {
    headline: `${A.chs.map(c => chName(c.ch).replace(' TV', '') + ' ' + nf(c.w * 100, 0)).join(' / ')} split reaches an estimated ${nf(A.net, 1)}% of the TV audience`,
    summary: [
      `Budget of ${lkr(A.B)} across ${A.chs.length} channels and ${A.kept.length} programs.`,
      anc.length ? `Built around ${anc.length} anchor program${anc.length > 1 ? 's' : ''}: ${anc.slice(0, 3).map(x => pn(x.p)).join(', ')}, which combine high ratings with week-to-week consistency.` : 'No program is both top rated and steady; the basket leans on consistent mid-rated shows.',
      dupNote,
      S.P.cprp > 0 ? `Expected delivery: about ${ni(A.spots)} spots, ${ni(A.grps)} GRPs, average frequency ${nf(A.freq, 1)}x.` : ''
    ].filter(Boolean),
    channelWhy: why,
    risks: (() => {
      const r = S.health.filter(h => h.t === 'wa').map(h => strip(h.m));
      if (top.w > .4) r.push(`${chName(top.ch)} holds ${nf(top.w * 100, 0)}% of the budget. A rating drop there hits the whole plan; keep a backup slot on the second channel.`);
      const vol = A.kept.filter(x => x.role === 'review');
      if (vol.length && !r.some(t => /volatile/.test(t))) r.push(`${vol.map(x => pn(x.p)).slice(0, 3).join(', ')} rate${vol.length > 1 ? '' : 's'} unevenly week to week. Cap the share and review weekly.`);
      r.push(`Net reach depends on assumed channel overlap (e.g. ${nf(S.getD('HIRU TV', 'DERANA TV') * 100, 0)}% Hiru–Derana). Measured duplication may move the estimate.`);
      r.push(`Ratings cover ${fmtDate(S.FS.from)} – ${fmtDate(S.FS.to)}. Programme changes, finales or seasonal events can shift delivery.`);
      return r;
    })(),
    nextSteps: ['Confirm rate card and cost per rating point with each channel', 'Book anchor slots first; they carry most of the reach', 'Review volatile programs weekly and switch to backups if ratings drop', 'Re-run the plan on the latest ratings file before each flight']
  };
}
async function aiRationale(S, A, base) {
  if (!S.aiOn) return null;
  const q = `Write the strategic rationale for a client presentation of the current plan. Reply with JSON only, no markdown fences, in this shape:
{"headline":"one sentence recommendation","summary":["3-4 short sentences on why this plan"],"channelWhy":{"<exact channel name>":"one sentence on why this weight"},"risks":["2-4 risks with mitigation"],"nextSteps":["3-4 actions"]}
Use only numbers from the summary. Channel names must match currentPlan.split exactly.`;
  try {
    const txt = await Promise.race([askRemote(q, buildContext(S), []), new Promise((_, r) => setTimeout(() => r(new Error('timeout')), 30000))]);
    const j = JSON.parse(txt.replace(/^```(?:json)?|```$/gm, '').trim().replace(/ACTION:.*$/m, ''));
    if (!j.headline || !Array.isArray(j.summary)) return null;
    return j;
  } catch (e) { return null; }
}

/* ---------- deck ---------- */
export async function buildDeck(S, simulate, opts = {}) {
  await loadLib();
  const A = S.cur, base = S.base, P = S.P, F = S.F, FS = S.FS;
  if (!A) throw new Error('The plan is empty. Add channels or programs first.');
  const rule = ruleRationale(S, A);
  const ai = opts.useAI === false ? null : await aiRationale(S, A, base);
  const R = Object.assign({}, rule, ai || {}, { channelWhy: Object.assign({}, rule.channelWhy, (ai && ai.channelWhy) || {}) });
  const src = ai ? 'Rationale drafted by Planner AI (Gemini) from the plan numbers; review before sending.' : 'Rationale written from the plan numbers.';

  const pres = new window.PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = 'TV Media Plan'; pres.author = 'TV Media Planner'; pres.subject = 'TV media plan and rationale';
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  const period = `${fmtDate(FS.from)} – ${fmtDate(FS.to)}`;

  pres.defineSlideMaster({
    title: 'TITLE', background: { color: COL.tealDk },
    placeholders: []
  });
  pres.defineSlideMaster({
    title: 'CONTENT', background: { color: COL.white },
    objects: [
      { text: { text: 'TV Media Plan · ' + period, options: { x: M, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: COL.muted, fontFace: FONT, margin: 0 } } },
      { placeholder: { options: { name: 'title', type: 'title', x: M, y: 0.35, w: W - 2 * M, h: 0.75, fontSize: 28, bold: true, color: COL.ink, fontFace: FONT, valign: 'top', align: 'left', margin: 0 }, text: '' } }
    ],
    slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 9, color: COL.muted, align: 'right', fontFace: FONT }
  });
  const add = (title, kicker) => {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    s.addText(title, { placeholder: 'title' });
    if (kicker) s.addText(kicker, { x: M, y: 1.08, w: W - 2 * M, h: 0.4, fontSize: 14, color: COL.muted, fontFace: FONT, margin: 0, isTextBox: true });
    return s;
  };
  const card = (s, x, y, w, h, fill = COL.soft) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 });
  const stat = (s, x, y, w, big, label, color = COL.teal) => {
    card(s, x, y, w, 1.45);
    s.addText(big, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.75, fontSize: 34, bold: true, color, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(label, { x: x + 0.25, y: y + 0.9, w: w - 0.5, h: 0.4, fontSize: 12, color: COL.muted, fontFace: FONT, margin: 0, isTextBox: true });
  };
  const bullets = (s, items, o) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
    Object.assign({ fontSize: 14, color: COL.ink, fontFace: FONT, paraSpaceAfter: 8, valign: 'top', margin: 0.05, isTextBox: true }, o));
  const chartBase = { catAxisLabelColor: COL.muted, valAxisLabelColor: COL.muted, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10, valGridLine: { color: COL.line, size: 0.5 }, catGridLine: { style: 'none' }, dataLabelFontFace: FONT, dataLabelFontSize: 9, dataLabelColor: COL.ink, titleFontFace: FONT, titleColor: COL.ink, titleFontSize: 13, legendFontFace: FONT, legendFontSize: 10 };
  const th = t => ({ text: t, options: { bold: true, color: COL.muted, fill: { color: COL.soft }, fontSize: 10 } });
  const chIdx = ch => Math.max(0, S.CH.indexOf(ch)) % PAL.length;

  /* 1. Title */
  let s = pres.addSlide({ masterName: 'TITLE' });
  s.addText('TV Media Plan', { x: M, y: 1.7, w: 11, h: 1.0, fontSize: 44, bold: true, color: COL.white, fontFace: FONT, margin: 0, isTextBox: true });
  s.addText(short(R.headline, 160), { x: M, y: 2.8, w: 11.5, h: 1.1, fontSize: 20, color: 'CFE8EA', fontFace: FONT, margin: 0, valign: 'top', isTextBox: true });
  s.addText([
    { text: `Period analysed: ${period}`, options: { breakLine: true } },
    { text: `Budget: ${lkr(A.B)}${P.cut ? ` (after ${P.cut}% cut from ${lkr(base.B)})` : ''}`, options: { breakLine: true } },
    { text: `Prepared ${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}` }
  ], { x: M, y: 4.6, w: 9, h: 1.2, fontSize: 14, color: COL.white, fontFace: FONT, margin: 0, isTextBox: true, paraSpaceAfter: 4 });
  s.addNotes(`Opening: ${R.headline}. ${src}`);

  /* 2. Executive summary */
  s = add('Executive summary', short(R.headline, 140));
  const sw = (W - 2 * M - 0.9) / 4;
  stat(s, M, 1.65, sw, nf(A.net, 1) + '%', 'Estimated net reach');
  stat(s, M + (sw + 0.3), 1.65, sw, lkr(A.B).replace('LKR ', ''), 'Budget (LKR)', COL.ink);
  stat(s, M + 2 * (sw + 0.3), 1.65, sw, `${A.chs.length} / ${A.kept.length}`, 'Channels / programs', COL.ink);
  stat(s, M + 3 * (sw + 0.3), 1.65, sw, P.cprp > 0 ? ni(A.grps) : nf(A.gross, 1) + '%', P.cprp > 0 ? `Est. GRPs (${ni(A.spots)} spots)` : 'Gross reach', COL.amber);
  s.addText('Why this plan', { x: M, y: 3.4, w: 6, h: 0.4, fontSize: 18, bold: true, color: COL.ink, fontFace: FONT, margin: 0, isTextBox: true });
  bullets(s, R.summary.slice(0, 4), { x: M, y: 3.85, w: 7.4, h: 2.9 });
  card(s, 8.4, 3.4, W - M - 8.4, 3.35);
  s.addText('Recommended split', { x: 8.65, y: 3.55, w: 3.8, h: 0.35, fontSize: 14, bold: true, color: COL.ink, fontFace: FONT, margin: 0, isTextBox: true });
  A.chs.slice(0, 6).forEach((c, i) => {
    const y = 4.05 + i * 0.42, bw = (W - M - 8.65 - 1.9) * c.w / A.chs[0].w;
    s.addText(short(chName(c.ch), 16), { x: 8.65, y, w: 1.5, h: 0.32, fontSize: 11, color: COL.ink, fontFace: FONT, margin: 0, valign: 'middle', isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: 10.2, y: y + 0.06, w: Math.max(0.05, bw), h: 0.2, fill: { color: PAL[chIdx(c.ch)] }, line: { color: PAL[chIdx(c.ch)] } });
    s.addText(nf(c.w * 100, 0) + '%', { x: 10.25 + bw, y, w: 0.6, h: 0.32, fontSize: 11, bold: true, color: COL.ink, fontFace: FONT, margin: 0, valign: 'middle', isTextBox: true });
  });
  s.addNotes(`Executive summary. ${R.summary.join(' ')} ${src}`);

  /* 3. Brief and approach */
  s = add('The brief and how we built the plan');
  const ff = FS;
  const briefRows = [
    ['Channels considered', ff.ch.size === S.CH.length ? `All ${S.CH.length}` : [...ff.ch].map(chName).join(', ')],
    ['Categories', ff.cat.size === S.CATS.length ? 'All categories' : short([...ff.cat].map(pn).join(', '), 90)],
    ['Days', ff.day.size === 7 ? 'All week' : DAYS.filter(d => ff.day.has(d)).map(d => d.slice(0, 3)).join(', ')],
    ['Time band', `${hl(ff.h0)} – ${hl((ff.h1 + 1) % 24)}`],
    ['Ratings period', period],
    ['Airings analysed', ni(F.length)],
    ['Budget', lkr(A.B)],
    ['Cost per rating point', P.cprp > 0 ? 'LKR ' + ni(P.cprp) + ` (${P.spotLen} sec)` : 'Not set']
  ];
  s.addTable(briefRows.map(r => [{ text: r[0], options: { color: COL.muted } }, { text: r[1], options: { bold: true } }]),
    { x: M, y: 1.4, w: 5.6, colW: [2.1, 3.5], fontSize: 12, fontFace: FONT, color: COL.ink, border: { type: 'solid', pt: 0.5, color: COL.line }, rowH: 0.48, valign: 'middle' });
  const steps = [
    ['Explore the market', 'Rank hours, days and channels by average TVR and reach across every airing in the brief.'],
    ['Weight the channels', 'Channel weight = 60% average reach + 40% average TVR, rescaled across the chosen channels.'],
    ['Pick the programs', 'Score = 60% rating + 40% steadiness (mean TVR ÷ std. dev.), so one-off spikes do not drive spend.'],
    ['Remove duplication', 'Net reach removes viewers who watch more than one channel, using a channel overlap matrix.']
  ];
  steps.forEach((st, i) => {
    const y = 1.4 + i * 1.3, x = 6.8;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: COL.teal }, line: { color: COL.teal } });
    s.addText(String(i + 1), { x, y: y + 0.05, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: COL.white, align: 'center', valign: 'middle', fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(st[0], { x: x + 0.8, y, w: 5.1, h: 0.4, fontSize: 15, bold: true, color: COL.ink, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(st[1], { x: x + 0.8, y: y + 0.4, w: 5.1, h: 0.75, fontSize: 12, color: COL.muted, fontFace: FONT, margin: 0, valign: 'top', isTextBox: true });
  });
  s.addNotes('The plan is built from the client brief filters. Every number comes from the uploaded ratings file for the period shown.');

  /* 4. Market context */
  const hrs = [...grp(F, r => r.h).values()].sort((a, b) => a.k - b.k);
  const ins = insights(F);
  s = add('When the audience watches', ins[0] ? `${ins[0].big} ${ins[0].t}` : '');
  s.addChart(pres.charts.BAR, [{ name: 'Avg TVR', labels: hrs.map(g => hl(g.k)), values: hrs.map(g => +avgT(g).toFixed(2)) }],
    Object.assign({}, chartBase, { x: M, y: 1.55, w: 7.8, h: 5.2, barDir: 'col', chartColors: [COL.teal], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', showLegend: false, showTitle: true, title: 'Average TVR by start hour' }));
  const days = [...grp(F, r => r.day).values()].sort((a, b) => avgR(b) - avgR(a));
  card(s, 8.7, 1.55, W - M - 8.7, 5.2);
  s.addText('What it means for the plan', { x: 8.95, y: 1.7, w: 3.6, h: 0.4, fontSize: 15, bold: true, color: COL.ink, fontFace: FONT, margin: 0, isTextBox: true });
  bullets(s, [
    ...ins.slice(0, 3).map(i => `${i.big} ${i.t}`),
    days.length > 1 ? `Best days by reach: ${days.slice(0, 3).map(d => d.k.slice(0, 3)).join(', ')}.` : ''
  ].filter(Boolean), { x: 8.95, y: 2.2, w: W - M - 9.2, h: 4.4, fontSize: 12 });
  s.addNotes('Ratings concentrate in prime time. The plan buys where the audience is, and the peak hour drives the anchor selection.');

  /* 5. Channel performance */
  const cs = S.plan.cs.slice(0, 9);
  s = add('Channel strength', `${chName(cs[0].ch)} leads on reach per airing; weights follow average reach (60%) and TVR (40%).`);
  s.addChart(pres.charts.BAR, [
    { name: 'Avg reach %', labels: cs.map(c => chName(c.ch)), values: cs.map(c => +c.reach.toFixed(2)) },
    { name: 'Avg TVR', labels: cs.map(c => chName(c.ch)), values: cs.map(c => +c.tvr.toFixed(2)) }
  ], Object.assign({}, chartBase, { x: M, y: 1.55, w: 7.8, h: 5.2, barDir: 'bar', barGrouping: 'clustered', chartColors: [COL.amber, COL.teal], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', showLegend: true, legendPos: 'b', catAxisOrientation: 'maxMin' }));
  const inPlan = new Set(A.chs.map(c => c.ch));
  s.addTable([[th('Channel'), th('Score'), th('In plan')], ...cs.map(c => [chName(c.ch), nf(c.score * 100, 1), inPlan.has(c.ch) ? 'Yes' : '—'])],
    { x: 8.7, y: 1.55, w: W - M - 8.7, colW: [1.9, 1.0, W - M - 8.7 - 2.9], fontSize: 11, fontFace: FONT, color: COL.ink, border: { type: 'solid', pt: 0.5, color: COL.line }, rowH: 0.36 });
  s.addNotes(`Channel score = 0.6 x share of average reach + 0.4 x share of average TVR. Top channels selected: ${A.chs.map(c => chName(c.ch)).join(', ')}.`);

  /* 6. Channel split with rationale */
  s = add('Recommended channel split', `${A.chs.map(c => chName(c.ch).replace(' TV', '') + ' ' + nf(c.w * 100, 0) + '%').join(' · ')} of ${lkr(A.B)}`);
  s.addChart(pres.charts.DOUGHNUT, [{ name: 'Share', labels: A.chs.map(c => chName(c.ch)), values: A.chs.map(c => +(c.w * 100).toFixed(1)) }],
    { x: M, y: 1.55, w: 4.3, h: 5.2, holeSize: 55, chartColors: A.chs.map(c => PAL[chIdx(c.ch)]), showPercent: true, showValue: false, dataLabelColor: COL.white, dataLabelFontSize: 11, dataLabelFontFace: FONT, showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 10 });
  s.addTable([[th('Channel'), th('Share'), th('Budget'), th('Shows'), th('Why this weight')],
    ...A.chs.map(c => [{ text: chName(c.ch), options: { bold: true } }, nf(c.w * 100, 1) + '%', lkr(c.bud).replace('LKR ', ''), String(c.progs.length), short(R.channelWhy[c.ch] || '', 150)])],
  { x: 5.15, y: 1.55, w: W - M - 5.15, colW: [1.45, 0.75, 0.9, 0.85, W - M - 5.15 - 3.95], fontSize: 10.5, fontFace: FONT, color: COL.ink, border: { type: 'solid', pt: 0.5, color: COL.line }, valign: 'middle' });
  s.addNotes(A.chs.map(c => `${chName(c.ch)} ${nf(c.w * 100, 1)}%: ${R.channelWhy[c.ch] || ''}`).join(' '));

  /* 7. Program basket */
  s = add('Program basket', 'Ranked by share of the plan. Steadiness = mean TVR ÷ standard deviation (7+ very steady, under 3 volatile).');
  const items = A.kept.slice(0, 13);
  const roleTxt = { anchor: 'Anchor', support: 'Support', review: 'Review' };
  s.addTable([[th('Program'), th('Channel'), th('Slot'), th('Avg TVR'), th('Reach %'), th('Steadiness'), th('Role'), th('Share'), th('LKR'), ...(P.cprp > 0 ? [th('Spots')] : [])],
    ...items.map(x => [{ text: short(pn(x.p), 34), options: { bold: true } }, chName(x.ch), hl(x.hour), nf(x.mean, 2), nf(x.rp, 1), steadiness(x.ci, x.n).label + (x.n >= 2 ? ' ' + nf(x.ci, 1) : ''),
      { text: roleTxt[x.role], options: { color: x.role === 'anchor' ? COL.good : x.role === 'review' ? COL.bad : COL.info, bold: true } }, nf(x.k * 100, 1) + '%', lkr(x.bud).replace('LKR ', ''), ...(P.cprp > 0 ? [ni(x.spots)] : [])])],
  { x: M, y: 1.55, w: W - 2 * M, colW: P.cprp > 0 ? [3.0, 1.5, 0.8, 0.85, 0.85, 1.45, 0.9, 0.85, 1.1, 0.83] : [3.3, 1.6, 0.9, 0.95, 0.95, 1.55, 1.0, 0.95, 0.93], fontSize: 10, fontFace: FONT, color: COL.ink, border: { type: 'solid', pt: 0.5, color: COL.line }, rowH: 0.36, valign: 'middle' });
  if (A.kept.length > items.length) s.addText(`+ ${A.kept.length - items.length} more programs in the spot plan export.`, { x: M, y: 6.6, w: 8, h: 0.3, fontSize: 10, color: COL.muted, fontFace: FONT, margin: 0, isTextBox: true });
  s.addNotes('Programs are picked within each channel by a blend of rating (60%) and steadiness (40%). Anchors are top-rated and steady; Review flags volatile programs.');

  /* 8. Why these programs */
  const anc = A.kept.filter(x => x.role === 'anchor').slice(0, 3);
  const showcase = anc.length ? anc : A.kept.slice(0, 3);
  s = add('Why these anchor programs', 'High ratings that repeat week after week carry the plan; volatile shows are capped.');
  const cw = (W - 2 * M - 0.6) / 3;
  showcase.forEach((x, i) => {
    const x0 = M + i * (cw + 0.3);
    card(s, x0, 1.6, cw, 3.4);
    s.addText(short(pn(x.p), 40), { x: x0 + 0.25, y: 1.75, w: cw - 0.5, h: 0.75, fontSize: 17, bold: true, color: COL.ink, fontFace: FONT, margin: 0, valign: 'top', isTextBox: true });
    s.addText(`${chName(x.ch)} · usually ${hl(x.hour)} · ${(x.days || []).map(d => d.slice(0, 2)).join(' ')}`, { x: x0 + 0.25, y: 2.5, w: cw - 0.5, h: 0.35, fontSize: 11, color: COL.muted, fontFace: FONT, margin: 0, isTextBox: true });
    const hw = (cw - 0.5) / 2;
    s.addText(nf(x.mean, 1), { x: x0 + 0.25, y: 2.95, w: hw, h: 0.8, fontSize: 32, bold: true, color: COL.teal, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText('avg TVR', { x: x0 + 0.25, y: 3.7, w: hw, h: 0.3, fontSize: 11, color: COL.muted, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(nf(x.rp, 1) + '%', { x: x0 + 0.25 + hw, y: 2.95, w: hw, h: 0.8, fontSize: 32, bold: true, color: COL.amber, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText('avg reach', { x: x0 + 0.25 + hw, y: 3.7, w: hw, h: 0.3, fontSize: 11, color: COL.muted, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(`${steadiness(x.ci, x.n).label} (${nf(x.ci, 1)}) over ${x.n} airings · ${nf(x.k * 100, 1)}% of plan`, { x: x0 + 0.25, y: 4.2, w: cw - 0.5, h: 0.6, fontSize: 11, color: COL.ink, fontFace: FONT, margin: 0, valign: 'top', isTextBox: true });
  });
  const rv = A.kept.filter(x => x.role === 'review');
  bullets(s, [
    `Anchors hold ${nf(anc.reduce((a, x) => a + x.k, 0) * 100, 0)}% of the budget, so most spend goes where ratings are both high and predictable.`,
    rv.length ? `Volatile programs (${rv.slice(0, 3).map(x => pn(x.p)).join(', ')}) are kept small and should be watched weekly.` : 'No program in the basket has volatile ratings.',
    'Support programs add reach in adjacent slots without relying on a single show.'
  ], { x: M, y: 5.25, w: W - 2 * M, h: 1.5, fontSize: 13 });
  s.addNotes('Anchor programs are chosen for both rating and consistency. Volatile programs carry less weight to protect delivery.');

  /* 9. Net reach and duplication */
  s = add('Net reach after duplication', `${nf(A.gross, 1)}% gross reach becomes ${nf(A.net, 1)}% net once shared viewers are counted once.`);
  const fw = 2.2;
  [[nf(A.gross, 1) + '%', 'Gross reach', COL.ink], [nf(A.gross - A.net, 1) + ' pts', 'Same viewers', COL.amber], [nf(A.net, 1) + '%', 'Net reach', COL.teal]].forEach((v, i) => {
    stat(s, M, 1.6 + i * 1.7, fw + 1.2, v[0], v[1], v[2]);
  });
  const nMax = Math.min(S.plan.cs.length, 9), pts = [];
  for (let k = 1; k <= nMax; k++) { const r = simulate({ chMode: 'auto', nCh: k, split: {} }); pts.push(r ? +r.net.toFixed(1) : 0); }
  s.addChart(pres.charts.LINE, [{ name: 'Net reach %', labels: pts.map((_, i) => `${i + 1} ch`), values: pts }],
    Object.assign({}, chartBase, { x: 4.4, y: 1.55, w: 5.0, h: 5.2, chartColors: [COL.teal], lineSize: 2.5, lineDataSymbolSize: 7, showValue: true, dataLabelPosition: 't', dataLabelFormatCode: '0.0', showLegend: false, showTitle: true, title: 'Net reach as channels are added', valAxisMinVal: 0 }));
  card(s, 9.7, 1.55, W - M - 9.7, 5.2);
  s.addText('How it works', { x: 9.9, y: 1.7, w: 2.7, h: 0.4, fontSize: 15, bold: true, color: COL.ink, fontFace: FONT, margin: 0, isTextBox: true });
  const dupTxt = [];
  if (A.chs.length > 1) { const a = A.chs[0], b = A.chs[1], d = S.getD(a.ch, b.ch); dupTxt.push(`${chName(a.ch)} and ${chName(b.ch)} share about ${nf(d * 100, 0)}% of viewers, so ${chName(b.ch)} adds ${nf(uni(a.R, b.R, d) - a.R, 1)} pts, not ${nf(b.R, 1)}.`); }
  dupTxt.push('The curve flattens as channels are added: each new channel brings fewer new viewers.');
  dupTxt.push('Overlap factors are planning assumptions; replace them with measured duplication when available.');
  bullets(s, dupTxt, { x: 9.9, y: 2.2, w: W - M - 10.1, h: 4.4, fontSize: 11.5 });
  s.addNotes('Net reach formula: A + B minus overlap, where overlap = A x B / 100 + d x (min(A, B) - A x B / 100) and d is the channel duplication factor.');

  /* 10. Budget scenarios */
  const levels = [0, 10, 25, 40].map(c => ({ c, r: c ? simulate({ cut: c }) : simulate({ cut: 0 }) })).filter(x => x.r);
  s = add('Budget scenarios', 'Cuts drop the weakest slots first, so reach falls more slowly than budget.');
  s.addChart(pres.charts.BAR, [{ name: 'Net reach %', labels: levels.map(l => l.c ? `−${l.c}%` : 'Full'), values: levels.map(l => +l.r.net.toFixed(1)) }],
    Object.assign({}, chartBase, { x: M, y: 1.55, w: 5.4, h: 5.2, barDir: 'col', chartColors: [COL.teal], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', showLegend: false, showTitle: true, title: 'Net reach % by budget level', valAxisMinVal: 0 }));
  const full = levels[0].r;
  const scRows = levels.map(l => [l.c ? `Budget −${l.c}%` : 'Full budget', lkr(l.r.B).replace('LKR ', ''), nf(l.r.net, 1) + '%', l.c ? nf((l.r.net / full.net - 1) * 100, 1) + '%' : '—', String(l.r.kept.length)]);
  (S.scenarios || []).slice(0, 4).forEach(sc => scRows.push([short(sc.name, 24), lkr(sc.r.B).replace('LKR ', ''), nf(sc.r.net, 1) + '%', nf((sc.r.net / full.net - 1) * 100, 1) + '%', String(sc.r.items.length)]));
  s.addTable([[th('Scenario'), th('Budget'), th('Net reach'), th('vs full'), th('Programs')], ...scRows],
    { x: 6.3, y: 1.55, w: W - M - 6.3, colW: [2.2, 1.15, 1.1, 1.0, W - M - 6.3 - 5.45], fontSize: 11, fontFace: FONT, color: COL.ink, border: { type: 'solid', pt: 0.5, color: COL.line }, rowH: 0.38 });
  const c25 = levels.find(l => l.c === 25);
  if (c25) s.addText(`A 25% cut keeps ${c25.r.kept.length} of ${full.kept.length} slots and loses ${nf(Math.abs((c25.r.net / full.net - 1) * 100), 0)}% of net reach.`, { x: 6.3, y: 6.2, w: W - M - 6.3, h: 0.5, fontSize: 12, color: COL.ink, fontFace: FONT, margin: 0, isTextBox: true, italic: true });
  s.addNotes('Budget scenarios re-run the same allocation with fewer slots; locked anchors are protected.');

  /* 11. Risks and next steps */
  s = add('Risks and next steps');
  card(s, M, 1.4, 5.9, 5.35);
  s.addText('Risks to watch', { x: M + 0.25, y: 1.55, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: COL.bad, fontFace: FONT, margin: 0, isTextBox: true });
  bullets(s, (R.risks && R.risks.length ? R.risks : ['No major risks flagged by the plan health check.']).slice(0, 5).map(t => short(strip(t), 190)), { x: M + 0.25, y: 2.05, w: 5.4, h: 4.5, fontSize: 14 });
  card(s, M + 6.2, 1.4, W - 2 * M - 6.2, 5.35);
  s.addText('Next steps', { x: M + 6.45, y: 1.55, w: 5, h: 0.4, fontSize: 16, bold: true, color: COL.teal, fontFace: FONT, margin: 0, isTextBox: true });
  bullets(s, (R.nextSteps || []).slice(0, 5).map(t => short(strip(t), 160)), { x: M + 6.45, y: 2.05, w: W - 2 * M - 6.7, h: 4.5, fontSize: 14 });
  s.addNotes('Risks come from the plan health check. Next steps turn the plan into bookings.');

  /* 12. Methodology */
  s = add('Methodology and assumptions');
  bullets(s, [
    `Source: uploaded ratings file, ${ni(S.ROWS.length)} airings; ${ni(F.length)} match the brief.`,
    'Channel weight = 0.6 × average reach % + 0.4 × average TVR, normalised across plan channels' + (Object.keys(P.split || {}).length ? '; some shares set manually.' : '.'),
    'Program score = 0.6 × relative mean TVR + 0.4 × steadiness (mean ÷ standard deviation, capped at 5).',
    `Net reach uses a channel duplication matrix (e.g. Hiru–Derana ${nf(S.getD('HIRU TV', 'DERANA TV'), 2)}) and ${nf(S.DINTRA, 2)} overlap between programs on the same channel.`,
    P.cprp > 0 ? `Spots ≈ budget ÷ (LKR ${ni(P.cprp)} × TVR × ${P.spotLen}/30). GRPs = spots × average TVR.` : 'Spot estimates are off (no cost per rating point set).',
    'Past ratings do not guarantee future delivery; re-run with the latest file before booking.',
    src
  ], { x: M, y: 1.4, w: W - 2 * M, h: 5.4, fontSize: 14 });

  const name = `TV-Media-Plan-${new Date().toISOString().slice(0, 10)}.pptx`;
  await pres.writeFile({ fileName: name });
  return { name, ai: !!ai };
}
