// Synthetic demo dataset (fictional programs, invented numbers).
// Lets a first-time visitor explore every screen before uploading real data.
import { DAYS } from './engine.js';

function rng(seed) {
  let s = seed >>> 0;
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
}

// [channel, strength, schedule]; schedule rows: [start, minutes, program, category, baseTVR, days]
// days: 'wd' weekdays, 'we' weekend, 'all'
const GRID = [
  ['HIRU TV', 1.0, [
    ['06:30', 90, 'Hiru Morning Show', 'BREAKFAST SHOW', 0.9, 'all'],
    ['12:00', 30, 'Hiru Midday News', 'NEWS', 1.8, 'all'],
    ['14:00', 30, 'Sanda Eliya', 'TELEDRAMAS - SINHALA', 1.6, 'wd'],
    ['18:55', 35, 'Hiru Prime News', 'NEWS', 7.8, 'all'],
    ['19:30', 30, 'Paata Kurullo', 'TELEDRAMAS - SINHALA', 11.0, 'wd'],
    ['20:00', 30, 'Akurata Yana Welawe', 'TELEDRAMAS - SINHALA', 9.6, 'wd'],
    ['20:30', 30, 'Mal Pipena Kale', 'TELEDRAMAS - SINHALA', 6.4, 'wd'],
    ['21:00', 30, 'Hiru Star', 'REALITY PROGRAMMES', 5.2, 'we'],
    ['21:30', 60, 'Hiru Night Cinema', 'MOVIES - SINHALA', 2.4, 'we'],
    ['22:00', 30, 'Hiru Late News', 'NEWS', 1.1, 'wd']
  ]],
  ['DERANA TV', 0.9, [
    ['06:00', 120, 'Aruna', 'BREAKFAST SHOW', 1.0, 'all'],
    ['13:00', 30, 'Derana Lunch News', 'NEWS', 1.5, 'all'],
    ['15:00', 30, 'Kids Corner', 'CHILDRENS PROGRAMME - SINHALA', 1.1, 'all'],
    ['19:00', 30, 'Ada Derana Prime Time News', 'NEWS', 7.4, 'all'],
    ['19:30', 30, 'Iskole', 'TELEDRAMAS - SINHALA', 8.2, 'wd'],
    ['20:00', 30, 'Sangeethe', 'TELEDRAMAS - SINHALA', 7.2, 'wd'],
    ['20:30', 30, 'Deweni Inima', 'TELEDRAMAS - SINHALA', 6.8, 'wd'],
    ['20:00', 90, 'Derana Dream Star', 'REALITY PROGRAMMES', 6.1, 'we'],
    ['21:30', 30, 'Ada Derana 360', 'DISCUSSION', 2.2, 'wd'],
    ['22:00', 60, 'Derana Late Movie', 'MOVIES - SINHALA', 1.2, 'all']
  ]],
  ['SIRASA TV', 0.6, [
    ['06:30', 90, 'Sirasa Pathikada', 'BREAKFAST SHOW', 0.7, 'wd'],
    ['12:30', 30, 'Sirasa Lunch News', 'NEWS', 0.9, 'all'],
    ['18:30', 30, 'Sirasa News 1st', 'NEWS', 3.4, 'all'],
    ['19:30', 30, 'Prema Dadayama', 'TELEDRAMAS - SINHALA', 4.2, 'wd'],
    ['20:00', 30, 'Sirasa Superstar', 'REALITY PROGRAMMES', 3.6, 'we'],
    ['20:30', 30, 'Hadawatha', 'TELEDRAMAS - SINHALA', 2.5, 'wd'],
    ['21:00', 60, 'Sirasa Cricket Live', 'SPORTS', 2.8, 'we'],
    ['22:00', 30, 'Pethikada Late', 'DISCUSSION', 0.6, 'wd']
  ]],
  ['SWARNAVAHINI', 0.55, [
    ['07:00', 60, 'Swarna Udasana', 'BREAKFAST SHOW', 0.5, 'all'],
    ['19:00', 30, 'Swarna News', 'NEWS', 2.2, 'all'],
    ['19:30', 30, 'Hiri Poda Wassa', 'TELEDRAMAS - SINHALA', 7.9, 'wd'],
    ['20:00', 30, 'Kolamba Kathawa', 'TELEDRAMAS - SINHALA', 3.1, 'wd'],
    ['20:30', 60, 'Swarna Weekend Comedy', 'COMEDY - SINHALA', 2.6, 'we'],
    ['21:00', 30, 'Swarna Sirasa', 'MAGAZINE PROGRAMME', 1.4, 'wd']
  ]],
  ['ITN', 0.4, [
    ['07:00', 60, 'ITN Morning', 'BREAKFAST SHOW', 0.4, 'all'],
    ['19:00', 30, 'ITN Prime News', 'NEWS', 1.4, 'all'],
    ['19:30', 30, 'Ran Kirilli', 'TELEDRAMAS - SINHALA', 1.9, 'wd'],
    ['20:00', 30, 'Isuru Bawana', 'TELEDRAMAS - SINHALA', 1.5, 'wd'],
    ['21:00', 60, 'ITN Film Night', 'MOVIES - SINHALA', 1.0, 'we']
  ]],
  ['SHAKTHI TV', 0.35, [
    ['06:30', 60, 'Shakthi Kaalai', 'BREAKFAST SHOW', 0.4, 'all'],
    ['13:00', 60, 'Puthu Vasantham', 'TELEDRAMAS - TAMIL', 3.8, 'wd'],
    ['18:30', 30, 'Shakthi News First', 'NEWS', 1.3, 'all'],
    ['19:30', 30, 'Anbe Sivam', 'TELEDRAMAS - TAMIL', 1.6, 'wd'],
    ['20:00', 60, 'Junior Super Star', 'REALITY PROGRAMMES', 1.5, 'we']
  ]],
  ['SIYATHA TV', 0.25, [
    ['19:00', 30, 'Siyatha News', 'NEWS', 0.7, 'all'],
    ['19:30', 30, 'Sanda Sahu', 'TELEDRAMAS - SINHALA', 1.1, 'wd'],
    ['20:30', 60, 'Siyatha Cinema', 'MOVIES - SINHALA', 0.6, 'we']
  ]],
  ['TV1', 0.12, [
    ['19:00', 30, 'TV1 Evening Report', 'NEWS', 0.2, 'all'],
    ['20:00', 60, 'Hollywood Hits', 'MOVIES - ENGLISH', 0.3, 'all'],
    ['21:00', 30, 'Business Today', 'BUSINESS PROGRAMME', 0.1, 'wd']
  ]],
  ['VASANTHAM TV', 0.1, [
    ['18:30', 30, 'Vasantham Seithigal', 'NEWS', 0.2, 'all'],
    ['19:30', 30, 'Uravugal', 'TELEDRAMAS - TAMIL', 0.3, 'wd'],
    ['20:30', 60, 'Tamil Cinema', 'MOVIES - TAMIL', 0.2, 'we']
  ]]
];

// Hour multiplier: prime time peaks at 8 PM.
const HOURW = { 6: .8, 7: .85, 12: .9, 13: 1, 14: .9, 15: .9, 18: 1, 19: 1.05, 20: 1.1, 21: .9, 22: .7 };

export function makeDemoRows() {
  const R = rng(20260901);
  const out = [];
  const start = new Date('2026-09-01T12:00:00');
  for (let i = 0; i < 30; i++) {
    const d = new Date(start.getTime() + i * 864e5);
    const iso = d.toISOString().slice(0, 10);
    const day = DAYS[(d.getDay() + 6) % 7];
    const we = day === 'Saturday' || day === 'Sunday';
    const dayW = { Monday: 1.05, Tuesday: .97, Wednesday: 1.02, Thursday: 1.06, Friday: .98, Saturday: .9, Sunday: .95 }[day];
    for (const [ch, , sched] of GRID) {
      for (const [s, mins, p, cat, base, days] of sched) {
        if (days === 'wd' && we) continue;
        if (days === 'we' && !we) continue;
        const h = +s.slice(0, 2);
        // Volatile shows swing more week to week.
        const vol = /Prema|Cricket|Hadawatha|Puthu/.test(p) ? .45 : .14;
        const noise = 1 + (R() - .5) * 2 * vol;
        const tvr = Math.max(0, +(base * dayW * (HOURW[h] || 1) * noise).toFixed(2));
        const reach = +(Math.min(60, tvr * (1.35 + R() * .25) + (mins >= 60 ? tvr * .3 : 0))).toFixed(2);
        const [sh, sm] = s.split(':').map(Number);
        const endM = sh * 60 + sm + mins;
        const e = String(Math.floor(endM / 60) % 24).padStart(2, '0') + ':' + String(endM % 60).padStart(2, '0');
        out.push({ ch, date: iso, day, s, e, h, p, dur: mins, cat, tvr, tot: tvr, sh: 0, reach: 0, rp: reach });
      }
    }
  }
  // TVR share % = program TVR / all channels' TVR in the same date + hour.
  const tot = new Map();
  out.forEach(r => { const k = r.date + '|' + r.h; tot.set(k, (tot.get(k) || 0) + r.tvr); });
  out.forEach(r => { const t = tot.get(r.date + '|' + r.h) || 1; r.sh = +(r.tvr / t * 100).toFixed(1); });
  return out;
}
