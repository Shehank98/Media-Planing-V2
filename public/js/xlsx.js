// Excel booking schedule (one sheet per channel + cover + spot list), built in the browser
// with ExcelJS. Layout follows the agency schedule format: programme rows grouped by creative,
// one column per campaign date, live formulas for spots, rates, totals and taxes.
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
  const { meta = {}, P, start, end } = o;
  const wb = new window.ExcelJS.Workbook();
  wb.creator = meta.by || 'TV Media Planner'; wb.created = new Date();
  const used = new Set(['Cover', 'Spot list']);
  const cover = wb.addWorksheet('Cover', { views: [{ showGridLines: false }], pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1, paperSize: 9 } });
  const days = M.days, FIRST = 15; // column O = first date, as in the agency format
  const lastCol = FIRST + days.length - 1;
  const period = `${fmtDate(start)} – ${fmtDate(end)}`;
  const chanRefs = [];

  for (const C of M.chans) {
    const name = sheetName(chName(C.ch), used);
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', xSplit: 2, ySplit: 8, showGridLines: false }], pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 } });
    ws.properties.defaultRowHeight = 16;
    const widths = { A: 20, B: 34, C: 14, D: 8, E: 8, F: 5, G: 5, H: 9, I: 22, J: 13, K: 13, L: 15, M: 15, N: 16 };
    Object.entries(widths).forEach(([k, w]) => ws.getColumn(k).width = w);
    for (let c = FIRST; c <= lastCol; c++) ws.getColumn(c).width = 4.2;
    const f = (cell, v, s = {}) => { const x = ws.getCell(cell); x.value = v; x.font = { name: FONT, size: s.size || 10, bold: !!s.bold, color: s.color ? { argb: s.color } : undefined, italic: !!s.italic }; if (s.fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.fill } }; if (s.fmt) x.numFmt = s.fmt; x.alignment = { vertical: 'middle', horizontal: s.h || 'left', wrapText: !!s.wrap }; if (s.border) x.border = box; return x; };
    // Header block
    [['Client', meta.client || ''], ['Brand', meta.brand || ''], ['Campaign', meta.campaign || ''], ['Negotiated discount', (C.disc || 0) / 100], ['Date', new Date()]].forEach((r, i) => {
      f('A' + (i + 1), r[0], { bold: true }); f('B' + (i + 1), r[1], { bold: i === 0, fmt: i === 3 ? '0%' : i === 4 ? 'dd-mmm-yyyy' : undefined });
    });
    f('D1', 'Channel', { bold: true }); f('E1', chName(C.ch), { bold: true, color: TEAL, size: 12 });
    f('D2', 'Period', { bold: true }); f('E2', period);
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
    const hdr = [['A', 'Channel'], ['B', 'Programme Name'], ['C', 'Day'], ['D', 'Time'], ['F', 'Dur'], ['H', 'No of Spots'], ['I', 'Brand'], ['J', 'Rate Card Rate'], ['K', 'Negotiated Rate'], ['L', 'Total Rate Card Rate'], ['M', 'Total Negotiated Rate'], ['N', 'Investment']];
    hdr.forEach(([k, v]) => f(k + '7', v, { bold: true, color: 'FFFFFFFF', fill: NAVY, h: 'center', wrap: true, border: true }));
    ['A', 'B', 'C', 'H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach(k => ws.mergeCells(`${k}7:${k}8`));
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
    for (const sec of C.sections) {
      f('B' + r, `${sec.m.name} · ${sec.m.dur} sec`, { bold: true, color: TEAL });
      for (let c = 1; c <= lastCol; c++) { const x = ws.getCell(r, c); x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SEC } }; x.border = box; }
      r++;
      for (const row of sec.rows) {
        const x = row.x, isFirst = !dataRows.length;
        f('A' + r, isFirst ? chName(C.ch) : '', { bold: true, border: true });
        f('B' + r, pn(x.p), { border: true });
        f('C' + r, x.pattern, { h: 'center', border: true });
        f('D' + r, x.from || hl(x.hour), { h: 'center', border: true }); f('E' + r, x.to || '', { h: 'center', border: true });
        f('F' + r, sec.m.dur, { h: 'center', border: true }); f('G' + r, 'Sec', { h: 'center', border: true });
        f('H' + r, { formula: `SUM(${col(FIRST)}${r}:${col(lastCol)}${r})`, result: row.spots }, { h: 'center', bold: true, border: true });
        f('I' + r, sec.m.name, { border: true });
        f('J' + r, { formula: `${Math.round(x.rate30)}/30*F${r}`, result: row.rc }, { fmt: NUM, h: 'right', border: true, italic: !x.rateSet });
        f('K' + r, { formula: `J${r}*(1-$B$4)`, result: row.ng }, { fmt: NUM, h: 'right', border: true });
        f('L' + r, { formula: `J${r}*H${r}`, result: row.trc }, { fmt: NUM, h: 'right', border: true });
        f('M' + r, { formula: `K${r}*H${r}`, result: row.tng }, { fmt: NUM, h: 'right', border: true });
        f('N' + r, { formula: `M${r}`, result: row.tng }, { fmt: NUM, h: 'right', bold: true, border: true });
        days.forEach((d, i) => {
          const v = row.byDate[d.iso] || null, cell = ws.getCell(r, FIRST + i), we = d.day === 'Saturday' || d.day === 'Sunday';
          cell.value = v; cell.alignment = { horizontal: 'center' }; cell.border = box; cell.font = { name: FONT, size: 10, bold: !!v };
          if (v || we) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: v ? ON : WE } };
        });
        dataRows.push(r); r++;
      }
      r++; // blank row between creatives
    }
    // Totals and taxes
    const top = 9, bot = r - 1, T = r;
    ws.mergeCells(`A${T}:B${T}`); f('A' + T, `${chName(C.ch)} TOTAL`, { bold: true, fill: SOFT, border: true });
    for (let c = 3; c <= lastCol; c++) { const x = ws.getCell(T, c); x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SOFT } }; x.border = box; x.font = { name: FONT, bold: true }; }
    const sum = k => ({ formula: `SUM(${k}${top}:${k}${bot})` });
    ws.getCell('H' + T).value = { ...sum('H'), result: C.spots }; ws.getCell('H' + T).alignment = { horizontal: 'center' };
    ['L', 'M', 'N'].forEach(k => { const x = ws.getCell(k + T); x.value = { ...sum(k), result: k === 'L' ? C.rcT : C.net }; x.numFmt = NUM; });
    for (let c = FIRST; c <= lastCol; c++) { const x = ws.getCell(T, c); x.value = { formula: `SUM(${col(c)}${top}:${col(c)}${bot})` }; x.alignment = { horizontal: 'center' }; x.numFmt = '0;-0;""'; }
    const taxes = [['SSCL ' + P.sscl + '%', `N${T}*${(+P.sscl || 0) / 100}`, C.sscl], ['VAT ' + P.vat + '%', `SUM(N${T}:N${T + 1})*${(+P.vat || 0) / 100}`, C.vat], ['Total with taxes', `SUM(N${T}:N${T + 2})`, C.total]];
    taxes.forEach(([l, fm, v], i) => {
      const rr = T + 1 + i; ws.mergeCells(`A${rr}:B${rr}`);
      f('A' + rr, l, { bold: true, fill: i === 2 ? NAVY : undefined, color: i === 2 ? 'FFFFFFFF' : undefined, border: true });
      f('N' + rr, { formula: fm, result: v }, { fmt: NUM, h: 'right', bold: true, fill: i === 2 ? NAVY : undefined, color: i === 2 ? 'FFFFFFFF' : undefined, border: true });
    });
    f('B' + (T + 5), 'Rate card rates in italics are estimates from the ratings (CPRP × TVR); replace them with the channel rate card. Negotiated rate = rate card × (1 − discount in B4).', { italic: true, color: 'FF5F6B7A', size: 9 });
    ws.pageSetup.printArea = `A1:${col(lastCol)}${T + 5}`;
    chanRefs.push({ name, ch: C.ch, T, spots: C.spots, rcT: C.rcT, net: C.net, sscl: C.sscl, vat: C.vat, total: C.total });
  }

  // Cover (matches the agency cover: client, campaign, period, TG, budget, date) + channel summary
  cover.getColumn('B').width = 22; cover.getColumn('C').width = 4; cover.getColumn('D').width = 34;
  ['E', 'F', 'G', 'H'].forEach(k => cover.getColumn(k).width = 20);
  const cf = (a, v, s = {}) => { const x = cover.getCell(a); x.value = v; x.font = { name: FONT, size: s.size || 12, bold: !!s.bold, color: s.color ? { argb: s.color } : undefined }; if (s.fmt) x.numFmt = s.fmt; if (s.fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.fill } }; x.alignment = { vertical: 'middle', horizontal: s.h || 'left' }; if (s.border) x.border = box; return x; };
  cover.mergeCells('B2:H3'); cf('B2', 'TV MEDIA SCHEDULE', { bold: true, size: 20, color: 'FFFFFFFF', fill: NAVY });
  const netRef = chanRefs.map(c => `'${c.name}'!N${c.T}`).join('+') || '0';
  const totRef = chanRefs.map(c => `'${c.name}'!N${c.T + 3}`).join('+') || '0';
  const rows = [['Client :', meta.client || ''], ['Brand :', meta.brand || ''], ['Campaign :', meta.campaign || ''], ['Period :', period], ['Primary TG :', meta.tg || ''],
    ['Budget 100% :', { formula: netRef, result: M.tot.net }], ['Total with taxes :', { formula: totRef, result: M.tot.total }], ['Creatives :', M.mix.map(m => `${m.name} ${m.dur}s (${Math.round(m.w * 100)}%)`).join(', ')], ['Date :', new Date()]];
  rows.forEach(([l, v], i) => { cf('B' + (5 + i), l, { bold: true }); const x = cf('D' + (5 + i), v, { fmt: i === 5 || i === 6 ? '"LKR "#,##0' : i === 8 ? 'dd-mmm-yyyy' : undefined }); if (i === 5) cf('E' + (5 + i), '(without tax)', { size: 10, color: 'FF5F6B7A' }); });
  const h0 = 16;
  cover.mergeCells(`B${h0}:C${h0}`);
  [['B', 'Channel'], ['D', 'Sheet'], ['E', 'Spots'], ['F', 'Rate card value'], ['G', 'Negotiated'], ['H', 'Total with taxes']].forEach(([k, t], i) => cf(k + h0, t, { bold: true, size: 11, color: 'FFFFFFFF', fill: NAVY, border: true, h: i ? 'center' : 'left' }));
  chanRefs.forEach((c, i) => {
    const rr = h0 + 1 + i;
    cover.mergeCells(`B${rr}:C${rr}`); cf('B' + rr, chName(c.ch), { size: 11, border: true }); cf('D' + rr, c.name, { size: 11, border: true });
    cf('E' + rr, { formula: `'${c.name}'!H${c.T}`, result: c.spots }, { size: 11, border: true, h: 'center' });
    cf('F' + rr, { formula: `'${c.name}'!L${c.T}`, result: c.rcT }, { size: 11, fmt: NUM, border: true, h: 'right' });
    cf('G' + rr, { formula: `'${c.name}'!N${c.T}`, result: c.net }, { size: 11, fmt: NUM, border: true, h: 'right' });
    cf('H' + rr, { formula: `'${c.name}'!N${c.T + 3}`, result: c.total }, { size: 11, fmt: NUM, border: true, h: 'right' });
  });
  const tr = h0 + 1 + chanRefs.length;
  cover.mergeCells(`B${tr}:C${tr}`); cf('B' + tr, 'TOTAL', { bold: true, size: 11, fill: SOFT, border: true }); cf('D' + tr, '', { fill: SOFT, border: true });
  ['E', 'F', 'G', 'H'].forEach((k, i) => cf(k + tr, { formula: `SUM(${k}${h0 + 1}:${k}${tr - 1})`, result: [M.tot.spots, M.tot.rcT, M.tot.net, M.tot.total][i] }, { bold: true, size: 11, fmt: i ? NUM : undefined, fill: SOFT, border: true, h: i ? 'right' : 'center' }));
  cf('B' + (tr + 2), `Estimated net reach ${o.net.toFixed(1)}% · reached 3+ times ${o.r3.toFixed(1)}% (planning estimates from the ratings data).`, { size: 10, color: 'FF5F6B7A' });

  // Spot list: one row per date x programme x creative, for booking and checking.
  const sl = wb.addWorksheet('Spot list', { views: [{ state: 'frozen', ySplit: 1 }] });
  sl.columns = [{ header: 'Date', key: 'date', width: 12 }, { header: 'Day', key: 'day', width: 11 }, { header: 'Channel', key: 'ch', width: 15 }, { header: 'Programme', key: 'p', width: 34 }, { header: 'From', key: 'from', width: 8 }, { header: 'To', key: 'to', width: 8 },
    { header: 'Creative', key: 'cr', width: 22 }, { header: 'Dur (sec)', key: 'dur', width: 9 }, { header: 'Spots', key: 'spots', width: 7 }, { header: 'Rate card', key: 'rc', width: 13 }, { header: 'Negotiated', key: 'ng', width: 13 }];
  sl.getRow(1).eachCell(c => { c.font = { name: FONT, bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
  const list = [];
  M.chans.forEach(C => C.sections.forEach(sec => sec.rows.forEach(row => Object.entries(row.byDate).forEach(([iso, n]) => {
    const d = days.find(x => x.iso === iso);
    list.push({ iso, date: new Date(iso + 'T12:00:00'), day: d ? d.day : '', ch: chName(C.ch), p: pn(row.x.p), from: row.x.from, to: row.x.to, cr: sec.m.name, dur: sec.m.dur, spots: n, rc: Math.round(row.rc * n), ng: Math.round(row.ng * n) });
  }))));
  list.sort((a, b) => a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : a.ch.localeCompare(b.ch) || String(a.from).localeCompare(String(b.from)));
  list.forEach(({ iso, ...v }) => sl.addRow(v));
  sl.getColumn('date').numFmt = 'dd-mmm-yyyy'; sl.getColumn('rc').numFmt = NUM; sl.getColumn('ng').numFmt = NUM;
  sl.autoFilter = { from: 'A1', to: 'K1' };

  const buf = await wb.xlsx.writeBuffer();
  const safe = s => String(s || '').replace(/[^\w\- ]+/g, '').trim();
  const name = `${safe(meta.client) || 'TV'} - Schedule - ${start} to ${end}.xlsx`.replace(/\s+/g, ' ');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  return name;
}
