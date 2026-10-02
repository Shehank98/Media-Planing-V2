// Excel booking schedule (one sheet per channel + cover + spot list), built in the browser
// with ExcelJS. Layout follows the agency schedule format: programme rows grouped by creative,
// one column per campaign date, live formulas for TVR/GRP/NGRP, spots, rates, values, taxes,
// the asset (creative) split and CPRP / NCPRP.
import { chName, pn, fmtDate, hl } from './engine.js';

const LIB = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
const NAVY = 'FF1B2A41', TEAL = 'FF0E6E78', SOFT = 'FFF2F5F8', SEC = 'FFE3F1F2', WE = 'FFE9ECF1', ON = 'FFCFE7E9', LINE = 'FFD0D7E2';
const FONT = 'Calibri';
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const NUM = '#,##0;(#,##0);"-"';

function loadLib() {
  if (window.ExcelJS) return Promise.resolve();
  return new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = LIB; s.onload = res;
    s.onerror = () => rej(new Error('Could not load the Excel library. Check your internet connection.'));
    document.head.appendChild(s);
  });
}
// Column letter for a 1-based index.
const col = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const sheetName = (s, used) => { let n = s.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31).trim() || 'Sheet'; let k = n, i = 2; while (used.has(k)) k = n.slice(0, 28) + ' ' + i++; used.add(k); return k; };
const thin = { style: 'thin', color: { argb: LINE } };
const box = { top: thin, left: thin, bottom: thin, right: thin };

export async function exportScheduleXlsx(M, o) {
  await loadLib();
  const { meta: meta0 = {}, P, start, end } = o;
  const meta = Object.assign({}, meta0, meta0.scenario ? { campaign: [meta0.campaign, 'Scenario: ' + meta0.scenario].filter(Boolean).join(' · ') } : {});
  const wb = new window.ExcelJS.Workbook();
  wb.creator = meta.by || 'TV Media Planner'; wb.created = new Date();
  const used = new Set(['Cover', 'Spot list']);
  const cover = wb.addWorksheet('Cover', { views: [{ showGridLines: false }], pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1, paperSize: 9 } });
  const days = M.days, FIRST = 18; // column R = first date, as in the agency format
  const lastCol = FIRST + days.length - 1;
  const period = `${fmtDate(start)} – ${fmtDate(end)}`;
  const chanRefs = [];

  for (const C of M.chans) {
    const name = sheetName(chName(C.ch), used);
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', xSplit: 2, ySplit: 8, showGridLines: false }], pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 } });
    ws.properties.defaultRowHeight = 16;
    const widths = { A: 20, B: 34, C: 12, D: 8, E: 8, F: 5, G: 5, H: 7, I: 8, J: 8, K: 9, L: 13, M: 13, N: 15, O: 15, P: 15, Q: 15 };
    Object.entries(widths).forEach(([k, w]) => ws.getColumn(k).width = w);
    for (let c = FIRST; c <= lastCol; c++) ws.getColumn(c).width = 4.2;
    const f = (cell, v, s = {}) => { const x = ws.getCell(cell); x.value = v; x.font = { name: FONT, size: s.size || 10, bold: !!s.bold, color: s.color ? { argb: s.color } : undefined, italic: !!s.italic }; if (s.fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.fill } }; if (s.fmt) x.numFmt = s.fmt; x.alignment = { vertical: 'middle', horizontal: s.h || 'left', wrapText: !!s.wrap }; if (s.border) x.border = box; return x; };
    // Header block
    [['Client', meta.client || ''], ['Brand', meta.brand || ''], ['Campaign', meta.campaign || ''], ['Negotiated discount', (C.disc || 0) / 100], ['Date', new Date()]].forEach((r, i) => {
      f('A' + (i + 1), r[0], { bold: true }); f('B' + (i + 1), r[1], { bold: i === 0, fmt: i === 3 ? '0%' : i === 4 ? 'dd-mmm-yyyy' : undefined });
    });
    f('D1', 'Channel', { bold: true }); f('E1', chName(C.ch), { bold: true, color: TEAL, size: 12 });
    f('D2', 'Period', { bold: true }); f('E2', period);
    f('D3', 'SR value %', { bold: true }); f('E3', M.srP, { fmt: '0%', color: TEAL, bold: true });
    // Month row (6), headers (7-8)
    let c0 = FIRST;
    days.forEach((d, i) => {
      const last = i === days.length - 1 || days[i + 1].iso.slice(0, 7) !== d.iso.slice(0, 7);
      if (last) {
        const c1 = FIRST + i; if (c1 > c0) ws.mergeCells(6, c0, 6, c1);
        const cell = ws.getCell(6, c0); cell.value = new Date(d.iso.slice(0, 7) + '-01T12:00:00'); cell.numFmt = 'mmm-yy';
        cell.font = { name: FONT, bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TEAL } }; cell.alignment = { horizontal: 'center' };
        c0 = c1 + 1;
      }
    });
    const srL = `SR Value - ${Math.round(M.srP * 100)}%`;
    const hdr = [['A', 'Channel'], ['B', 'Programme Name'], ['C', 'Day'], ['D', 'Time'], ['F', 'Dur'], ['H', 'TVR'], ['I', 'GRP'], ['J', 'NGRP'], ['K', 'No of Spots'], ['L', 'Rate Card Rate'], ['M', 'Negotiated Rate'], ['N', 'All Exposure Value'], ['O', 'Media Value'], ['P', 'Investment - 100%'], ['Q', srL]];
    hdr.forEach(([k, v]) => f(k + '7', v, { bold: true, color: 'FFFFFFFF', fill: NAVY, h: 'center', wrap: true, border: true }));
    ['A', 'B', 'C', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'].forEach(k => ws.mergeCells(`${k}7:${k}8`));
    ws.mergeCells('D7:E7'); ws.mergeCells('F7:G8');
    f('D8', 'From', { bold: true, color: 'FFFFFFFF', fill: NAVY, h: 'center', border: true }); f('E8', 'To', { bold: true, color: 'FFFFFFFF', fill: NAVY, h: 'center', border: true });
    days.forEach((d, i) => {
      const we = d.day === 'Saturday' || d.day === 'Sunday';
      const a = ws.getCell(7, FIRST + i), b = ws.getCell(8, FIRST + i);
      a.value = d.day[0]; b.value = +d.iso.slice(8);
      [a, b].forEach(x => { x.font = { name: FONT, bold: true, size: 9, color: { argb: 'FFFFFFFF' } }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: we ? TEAL : NAVY } }; x.alignment = { horizontal: 'center' }; x.border = box; });
    });
    ws.getRow(7).height = 30;
    // Body
    let r = 9; const dataRows = [];
    const secRng = [];
    for (const sec of C.sections) {
      f('B' + r, `${sec.m.name} - ${sec.m.dur} Sec`, { bold: true, color: TEAL });
      for (let c = 1; c <= lastCol; c++) { const x = ws.getCell(r, c); x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SEC } }; x.border = box; }
      r++;
      for (const row of sec.rows) {
        const x = row.x, isFirst = !dataRows.length;
        f('A' + r, isFirst ? chName(C.ch) : '', { bold: true, border: true });
        f('B' + r, pn(x.p), { border: true });
        f('C' + r, x.pattern, { h: 'center', border: true });
        f('D' + r, x.from || hl(x.hour), { h: 'center', border: true }); f('E' + r, x.to || '', { h: 'center', border: true });
        f('F' + r, sec.m.dur, { h: 'center', border: true }); f('G' + r, 'Sec', { h: 'center', border: true });
        f('H' + r, +row.tvr.toFixed(2), { fmt: '0.00', h: 'center', border: true });
        f('I' + r, { formula: `H${r}*K${r}`, result: row.grp }, { fmt: '0.00', h: 'center', border: true });
        f('J' + r, { formula: `(I${r}*F${r})/30`, result: row.ngrp }, { fmt: '0.00', h: 'center', border: true });
        f('K' + r, { formula: `SUM(${col(FIRST)}${r}:${col(lastCol)}${r})`, result: row.spots }, { h: 'center', bold: true, border: true });
        f('L' + r, { formula: `${Math.round(x.rate30)}/30*F${r}`, result: row.rc }, { fmt: NUM, h: 'right', border: true, italic: !x.rateSet });
        f('M' + r, { formula: `L${r}*(1-$B$4)`, result: row.ng }, { fmt: NUM, h: 'right', border: true });
        f('N' + r, { formula: `L${r}*K${r}`, result: row.trc }, { fmt: NUM, h: 'right', border: true });
        f('O' + r, { formula: `M${r}*K${r}`, result: row.tng }, { fmt: NUM, h: 'right', border: true });
        f('P' + r, { formula: `O${r}`, result: row.tng }, { fmt: NUM, h: 'right', bold: true, border: true });
        f('Q' + r, { formula: `P${r}*$E$3`, result: row.sr }, { fmt: NUM, h: 'right', border: true });
        days.forEach((d, i) => {
          const v = row.byDate[d.iso] || null, cell = ws.getCell(r, FIRST + i), we = d.day === 'Saturday' || d.day === 'Sunday';
          cell.value = v; cell.alignment = { horizontal: 'center' }; cell.border = box; cell.font = { name: FONT, size: 10, bold: !!v };
          if (v || we) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: v ? ON : WE } };
        });
        dataRows.push(r); r++;
      }
      secRng.push({ sec, a: r - sec.rows.length, b: r - 1 });
      r++; // blank row between creatives
    }
    // Totals and taxes
    const top = 9, bot = r - 1, T = r;
    ws.mergeCells(`A${T}:B${T}`); f('A' + T, `${chName(C.ch)} TOTAL`, { bold: true, fill: SOFT, border: true });
    for (let c = 3; c <= lastCol; c++) { const x = ws.getCell(T, c); x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SOFT } }; x.border = box; x.font = { name: FONT, bold: true }; }
    const sum = k => ({ formula: `SUM(${k}${top}:${k}${bot})` });
    const totV = { I: C.grp, J: C.ngrp, K: C.spots, N: C.rcT, O: C.net, P: C.net, Q: C.sr };
    Object.entries(totV).forEach(([k, v]) => { const x = ws.getCell(k + T); x.value = { ...sum(k), result: v }; x.numFmt = 'IJ'.includes(k) ? '0.00' : k === 'K' ? '0' : NUM; x.alignment = { horizontal: 'IJK'.includes(k) ? 'center' : 'right' }; });
    for (let c = FIRST; c <= lastCol; c++) { const x = ws.getCell(T, c); x.value = { formula: `SUM(${col(c)}${top}:${col(c)}${bot})` }; x.alignment = { horizontal: 'center' }; x.numFmt = '0;-0;""'; }
    const sp = (+P.sscl || 0) / 100, vp = (+P.vat || 0) / 100;
    const taxes = [['SSCL ' + P.sscl + '%', k => `${k}${T}*${sp}`, [C.sscl, C.srSscl]], ['VAT ' + P.vat + '%', k => `(${k}${T}+${k}${T + 1})*${vp}`, [C.vat, C.srVat]], ['Total with Taxes', k => `SUM(${k}${T}:${k}${T + 2})`, [C.total, C.srTotal]]];
    taxes.forEach(([l, fm, v], i) => {
      const rr = T + 1 + i, dark = i === 2; ws.mergeCells(`A${rr}:B${rr}`);
      f('A' + rr, l, { bold: true, fill: dark ? NAVY : undefined, color: dark ? 'FFFFFFFF' : undefined, border: true });
      ['P', 'Q'].forEach((k, j) => f(k + rr, { formula: fm(k), result: v[j] }, { fmt: NUM, h: 'right', bold: true, fill: dark ? NAVY : undefined, color: dark ? 'FFFFFFFF' : undefined, border: true }));
    });
    f('B' + (T + 4), 'Note: Any changes to the government taxes will change the final value of the invoice.', { italic: true, color: 'FF5F6B7A', size: 9 });
    // Asset table: paid value per creative and its share, then CPRP / NCPRP (commercial only).
    const A0 = T + 6, PEACH = 'FFF8CBAD', AF = 'FFD6DCE5';
    ws.mergeCells(`A${A0}:B${A0}`);
    f('A' + A0, 'Asset', { bold: true, fill: 'FFD9D9D9', h: 'center', border: true }); f('C' + A0, 'Ratio', { bold: true, fill: 'FFD9D9D9', h: 'center', border: true }); f('D' + A0, 'Paid Value', { bold: true, fill: 'FFD9D9D9', h: 'center', border: true });
    ws.mergeCells(`D${A0}:E${A0}`);
    const aTot = A0 + secRng.length + 1;
    secRng.forEach(({ sec, a, b }, i) => {
      const rr = A0 + 1 + i;
      f('B' + rr, `${sec.m.name} - ${sec.m.dur} Sec`, { bold: true, h: 'center', border: true, fill: AF });
      f('C' + rr, { formula: `IFERROR(D${rr}/$D$${aTot},0)`, result: sec.ratio }, { fmt: '0%', h: 'center', border: true, fill: AF });
      ws.mergeCells(`D${rr}:E${rr}`);
      f('D' + rr, { formula: `SUM(O${a}:O${b})`, result: sec.paid }, { fmt: NUM, h: 'right', border: true, fill: AF });
    });
    if (secRng.length) { if (secRng.length > 1) ws.mergeCells(`A${A0 + 1}:A${A0 + secRng.length}`); f('A' + (A0 + 1), 'Commercial', { bold: true, size: 12, h: 'center', fill: PEACH, border: true }); }
    f('A' + aTot, 'Total', { bold: true, border: true, fill: SOFT }); f('B' + aTot, '', { border: true, fill: SOFT });
    f('C' + aTot, { formula: `SUM(C${A0 + 1}:C${aTot - 1})`, result: 1 }, { fmt: '0%', h: 'center', bold: true, border: true, fill: SOFT });
    ws.mergeCells(`D${aTot}:E${aTot}`);
    f('D' + aTot, { formula: `SUM(D${A0 + 1}:D${aTot - 1})`, result: C.net }, { fmt: NUM, h: 'right', bold: true, border: true, fill: SOFT });
    const GREEN = 'FFA9D08E', c0r = aTot + 2;
    [['Com Only CPRP', `IFERROR(O${T}/I${T},0)`, C.cprp], ['Com Only NCPRP', `IFERROR(O${T}/J${T},0)`, C.ncprp]].forEach(([l, fm, v], i) => {
      const rr = c0r + i; f('A' + rr, l, { bold: true, fill: GREEN, border: true });
      ws.mergeCells(`B${rr}:D${rr}`); f('B' + rr, { formula: fm, result: v }, { fmt: '#,##0.00', h: 'right', bold: true, fill: GREEN, border: true });
    });
    // Estimated reach on this channel (planning model values).
    const R0 = c0r + 3, RH = ['Reach 1+', 'Reach 2+', 'Reach 3+', 'Reach 4+', 'Reach 5+'];
    ws.mergeCells(`A${R0}:B${R0}`); f('A' + R0, `Estimated reach on ${chName(C.ch)} (% of target audience)`, { bold: true, color: 'FFFFFFFF', fill: TEAL, border: true });
    RH.forEach((h, i) => { f('A' + (R0 + 1 + i), h, { bold: i === 0 || i === 2, border: true }); f('B' + (R0 + 1 + i), C.reach.at[i] / 100, { fmt: '0.0%', h: 'right', border: true, bold: i === 0 || i === 2 }); });
    f('A' + (R0 + 6), 'Average frequency', { border: true }); f('B' + (R0 + 6), C.reach.freq, { fmt: '0.0"x"', h: 'right', border: true });
    f('A' + (R0 + 7), 'Channel alone, before duplication with other channels. Planning estimate (NBD model fitted to reach and GRP).', { italic: true, color: 'FF5F6B7A', size: 9 });
    f('A' + (c0r + 12), 'GRP = TVR × spots. NGRP = GRP × duration ÷ 30. CPRP = Media Value ÷ GRP; NCPRP = Media Value ÷ NGRP. Rate card rates in italics are estimates (CPRP × TVR); replace them with the channel rate card. Negotiated rate = rate card × (1 − discount in B4).', { italic: true, color: 'FF5F6B7A', size: 9 });
    ws.pageSetup.printArea = `A1:${col(lastCol)}${c0r + 12}`;
    chanRefs.push({ name, ch: C.ch, T, C });
  }

  // Cover (matches the agency cover: client, campaign, period, TG, budget, date) + channel summary
  cover.getColumn('B').width = 22; cover.getColumn('C').width = 4; cover.getColumn('D').width = 34;
  ['E', 'F', 'G', 'H'].forEach(k => cover.getColumn(k).width = 20);
  const cf = (a, v, s = {}) => { const x = cover.getCell(a); x.value = v; x.font = { name: FONT, size: s.size || 12, bold: !!s.bold, color: s.color ? { argb: s.color } : undefined }; if (s.fmt) x.numFmt = s.fmt; if (s.fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.fill } }; x.alignment = { vertical: 'middle', horizontal: s.h || 'left' }; if (s.border) x.border = box; return x; };
  cover.mergeCells('B2:H3'); cf('B2', 'TV MEDIA SCHEDULE', { bold: true, size: 20, color: 'FFFFFFFF', fill: NAVY });
  const netRef = chanRefs.map(c => `'${c.name}'!P${c.T}`).join('+') || '0';
  const totRef = chanRefs.map(c => `'${c.name}'!P${c.T + 3}`).join('+') || '0';
  const rows = [['Client :', meta.client || ''], ['Brand :', meta.brand || ''], ['Campaign :', meta.campaign || ''], ['Period :', period], ['Primary TG :', meta.tg || ''],
    ['Budget 100% :', { formula: netRef, result: M.tot.net }], ['Total with taxes :', { formula: totRef, result: M.tot.total }], ['Creatives :', M.mix.map(m => `${m.name} ${m.dur}s (${Math.round(m.w * 100)}%)`).join(', ')], ['Date :', new Date()]];
  const goals = [o.gT ? `Target GRP ${Math.round(o.gT)} (scheduled ${M.tot.grp.toFixed(1)})` : '', o.minR ? `Minimum net reach ${o.minR}% (estimated ${M.tot.reach.at[0].toFixed(1)}%)` : ''].filter(Boolean).join(' · ');
  if (goals) { cf('B14', 'Goals :', { bold: true }); cf('D14', goals); }
  rows.forEach(([l, v], i) => { cf('B' + (5 + i), l, { bold: true }); const x = cf('D' + (5 + i), v, { fmt: i === 5 || i === 6 ? '"LKR "#,##0' : i === 8 ? 'dd-mmm-yyyy' : undefined }); if (i === 5) cf('E' + (5 + i), '(without tax)', { size: 10, color: 'FF5F6B7A' }); });
  const h0 = 16;
  ['I', 'J', 'K'].forEach(k => cover.getColumn(k).width = 20);
  cover.mergeCells(`B${h0}:C${h0}`);
  const heads = [['B', 'Channel'], ['D', 'Sheet'], ['E', 'Spots'], ['F', 'GRP'], ['G', 'NGRP'], ['H', 'All Exposure Value'], ['I', 'Media Value'], ['J', `SR Value - ${Math.round(M.srP * 100)}%`], ['K', 'Total with taxes']];
  heads.forEach(([k, t], i) => cf(k + h0, t, { bold: true, size: 11, color: 'FFFFFFFF', fill: NAVY, border: true, h: i ? 'center' : 'left' }));
  const refs = [['E', 'K', 'spots', '0'], ['F', 'I', 'grp', '0.00'], ['G', 'J', 'ngrp', '0.00'], ['H', 'N', 'rcT', NUM], ['I', 'O', 'net', NUM], ['J', 'Q', 'sr', NUM]];
  chanRefs.forEach((c, i) => {
    const rr = h0 + 1 + i;
    cover.mergeCells(`B${rr}:C${rr}`); cf('B' + rr, chName(c.ch), { size: 11, border: true }); cf('D' + rr, c.name, { size: 11, border: true });
    refs.forEach(([k, src, v, fm]) => cf(k + rr, { formula: `'${c.name}'!${src}${c.T}`, result: c.C[v] }, { size: 11, fmt: fm, border: true, h: 'right' }));
    cf('K' + rr, { formula: `'${c.name}'!P${c.T + 3}`, result: c.C.total }, { size: 11, fmt: NUM, border: true, h: 'right' });
  });
  const tr = h0 + 1 + chanRefs.length;
  cover.mergeCells(`B${tr}:C${tr}`); cf('B' + tr, 'TOTAL', { bold: true, size: 11, fill: SOFT, border: true }); cf('D' + tr, '', { fill: SOFT, border: true });
  const tv = { E: M.tot.spots, F: M.tot.grp, G: M.tot.ngrp, H: M.tot.rcT, I: M.tot.net, J: M.tot.sr, K: M.tot.total };
  Object.entries(tv).forEach(([k, v]) => cf(k + tr, { formula: `SUM(${k}${h0 + 1}:${k}${tr - 1})`, result: v }, { bold: true, size: 11, fmt: 'FG'.includes(k) ? '0.00' : k === 'E' ? '0' : NUM, fill: SOFT, border: true, h: 'right' }));
  // Campaign asset split and CPRP across all channels.
  const a0 = tr + 3, aT = a0 + M.assets.length + 1;
  cover.mergeCells(`B${a0}:D${a0}`); cf('B' + a0, 'Asset', { bold: true, size: 11, fill: 'FFD9D9D9', border: true, h: 'center' });
  cf('E' + a0, 'Ratio', { bold: true, size: 11, fill: 'FFD9D9D9', border: true, h: 'center' }); cf('F' + a0, 'Paid Value', { bold: true, size: 11, fill: 'FFD9D9D9', border: true, h: 'center' });
  M.assets.forEach((a, i) => {
    const rr = a0 + 1 + i, parts = chanRefs.map(c => { const sec = c.C.sections.find(x => x.i === a.i); return sec ? `'${c.name}'!D${c.T + 6 + 1 + c.C.sections.indexOf(sec)}` : null; }).filter(Boolean);
    cf('C' + rr, `${a.m.name} - ${a.m.dur} Sec`, { size: 11, bold: true, border: true, fill: 'FFD6DCE5' }); cover.mergeCells(`C${rr}:D${rr}`);
    cf('E' + rr, { formula: `IFERROR(F${rr}/$F$${aT},0)`, result: a.ratio }, { size: 11, fmt: '0%', border: true, h: 'center', fill: 'FFD6DCE5' });
    cf('F' + rr, { formula: parts.length ? parts.join('+') : '0', result: a.paid }, { size: 11, fmt: NUM, border: true, h: 'right', fill: 'FFD6DCE5' });
  });
  if (M.assets.length) { if (M.assets.length > 1) cover.mergeCells(`B${a0 + 1}:B${a0 + M.assets.length}`); cf('B' + (a0 + 1), 'Commercial', { bold: true, size: 12, fill: 'FFF8CBAD', border: true, h: 'center' }); }
  cover.mergeCells(`B${aT}:D${aT}`); cf('B' + aT, 'Total', { bold: true, size: 11, fill: SOFT, border: true });
  cf('E' + aT, { formula: `SUM(E${a0 + 1}:E${aT - 1})`, result: 1 }, { bold: true, size: 11, fmt: '0%', fill: SOFT, border: true, h: 'center' });
  cf('F' + aT, { formula: `SUM(F${a0 + 1}:F${aT - 1})`, result: M.tot.net }, { bold: true, size: 11, fmt: NUM, fill: SOFT, border: true, h: 'right' });
  [['Com Only CPRP', `IFERROR(I${tr}/F${tr},0)`, M.tot.cprp], ['Com Only NCPRP', `IFERROR(I${tr}/G${tr},0)`, M.tot.ncprp]].forEach(([l, fm, v], i) => {
    const rr = aT + 2 + i; cover.mergeCells(`B${rr}:D${rr}`); cf('B' + rr, l, { bold: true, size: 11, fill: 'FFA9D08E', border: true });
    cf('E' + rr, { formula: fm, result: v }, { bold: true, size: 11, fmt: '#,##0.00', fill: 'FFA9D08E', border: true, h: 'right' }); cover.mergeCells(`E${rr}:F${rr}`);
  });
  // Estimated reach: campaign (each person once) and each channel on its own.
  const r0 = aT + 6, RH = ['Reach 1+', 'Reach 2+', 'Reach 3+', 'Reach 4+', 'Reach 5+', 'Avg frequency', 'GRP'];
  cover.mergeCells(`B${r0 - 1}:H${r0 - 1}`); cf('B' + (r0 - 1), 'ESTIMATED REACH (% of target audience)', { bold: true, size: 12, color: 'FFFFFFFF', fill: TEAL });
  cover.mergeCells(`B${r0}:C${r0}`); cf('B' + r0, '', { fill: NAVY, border: true });
  ['D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k, i) => cf(k + r0, RH[i], { bold: true, size: 11, color: 'FFFFFFFF', fill: NAVY, border: true, h: 'center' }));
  const rrow = (rr, lbl, x, bold) => {
    cover.mergeCells(`B${rr}:C${rr}`); cf('B' + rr, lbl, { size: 11, bold, border: true, fill: bold ? SOFT : undefined });
    x.at.forEach((v, i) => cf(String.fromCharCode(68 + i) + rr, v / 100, { size: 11, fmt: '0.0%', border: true, h: 'center', bold: bold || i === 2, fill: bold ? SOFT : undefined }));
    cf('I' + rr, x.freq, { size: 11, fmt: '0.0"x"', border: true, h: 'center', fill: bold ? SOFT : undefined, bold });
    cf('J' + rr, x.grp, { size: 11, fmt: '0.00', border: true, h: 'center', fill: bold ? SOFT : undefined, bold });
  };
  rrow(r0 + 1, 'Campaign total (net)', M.tot.reach, true);
  M.chans.forEach((c, i) => rrow(r0 + 2 + i, chName(c.ch), c.reach, false));
  const rn = r0 + 2 + M.chans.length;
  cf('B' + (rn + 1), 'Reach 1+ = % who see the ad at least once; 3+ = at least three times. Channel rows are each channel alone; they add to more than the total because viewers watch several channels.', { size: 10, color: 'FF5F6B7A' });
  cf('B' + (rn + 2), 'Planning estimates from the ratings data (programme reach, repeat-spot and duplication factors, negative binomial (NBD) spread of exposures fitted to reach and GRP). Average frequency = GRP ÷ reach.', { size: 10, color: 'FF5F6B7A' });

  // Spot list: one row per date x programme x creative, for booking and checking.
  const sl = wb.addWorksheet('Spot list', { views: [{ state: 'frozen', ySplit: 1 }] });
  sl.columns = [{ header: 'Date', key: 'date', width: 12 }, { header: 'Day', key: 'day', width: 11 }, { header: 'Channel', key: 'ch', width: 15 }, { header: 'Programme', key: 'p', width: 34 }, { header: 'From', key: 'from', width: 8 }, { header: 'To', key: 'to', width: 8 },
    { header: 'Creative', key: 'cr', width: 22 }, { header: 'Dur (sec)', key: 'dur', width: 9 }, { header: 'Spots', key: 'spots', width: 7 }, { header: 'TVR', key: 'tvr', width: 7 }, { header: 'GRP', key: 'grp', width: 7 }, { header: 'NGRP', key: 'ngrp', width: 7 }, { header: 'Rate card', key: 'rc', width: 13 }, { header: 'Negotiated', key: 'ng', width: 13 }];
  sl.getRow(1).eachCell(c => { c.font = { name: FONT, bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
  const list = [];
  M.chans.forEach(C => C.sections.forEach(sec => sec.rows.forEach(row => Object.entries(row.byDate).forEach(([iso, n]) => {
    const d = days.find(x => x.iso === iso);
    list.push({ iso, date: new Date(iso + 'T12:00:00'), day: d ? d.day : '', ch: chName(C.ch), p: pn(row.x.p), from: row.x.from, to: row.x.to, cr: sec.m.name, dur: sec.m.dur, spots: n, tvr: +row.tvr.toFixed(2), grp: +(row.tvr * n).toFixed(2), ngrp: +(row.tvr * n * sec.m.dur / 30).toFixed(2), rc: Math.round(row.rc * n), ng: Math.round(row.ng * n) });
  }))));
  list.sort((a, b) => a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : a.ch.localeCompare(b.ch) || String(a.from).localeCompare(String(b.from)));
  list.forEach(({ iso, ...v }) => sl.addRow(v));
  sl.getColumn('date').numFmt = 'dd-mmm-yyyy'; sl.getColumn('rc').numFmt = NUM; sl.getColumn('ng').numFmt = NUM;
  sl.autoFilter = { from: 'A1', to: 'N1' };

  const buf = await wb.xlsx.writeBuffer();
  const safe = s => String(s || '').replace(/[^\w\- ]+/g, '').trim();
  const name = `${safe(meta.client) || 'TV'} - Schedule${meta0.scenario ? ' - ' + safe(meta0.scenario) : ''} - ${start} to ${end}.xlsx`.replace(/\s+/g, ' ');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  return name;
}
