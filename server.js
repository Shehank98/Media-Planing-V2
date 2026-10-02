// Zero-dependency Node server for Railway.
// - Serves the static dashboard from /public
// - POST /api/ai proxies questions to Gemini (only if GEMINI_API_KEY is set)
// Uploaded datasets never reach this server: files are parsed in the browser.
// The AI endpoint only receives a small aggregated summary the browser builds.
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = path.join(__dirname, 'public');
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-2.5-flash';
const MAX_BODY = 256 * 1024;

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.csv': 'text/csv; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8'
};

const SYSTEM_PROMPT = `You are "Planner AI", a senior TV media planner for the Sri Lankan market acting as strategic director.
You receive a JSON summary of the user's TV ratings data (already filtered in their browser) and their current media plan.
Rules:
- Base every number you quote on the JSON summary. Never invent programs, channels or figures. If the data cannot answer, say so.
- Be practical and short: lead with a one-line answer, then 2-5 bullet points. Use plain English, agency tone.
- Terms: TVR = average rating, Reach % = unique audience reached, net reach removes duplicated viewers. Never mention or calculate GRPs.
- When the user asks for a what-if (budget change, different channels, hours, days, channel count, programs per channel),
  end your reply with ONE line exactly in this form so the app can apply it:
  ACTION: {"cut":25,"nCh":4,"nProg":3,"budget":10000000,"channels":["HIRU TV"],"h0":19,"h1":22,"days":["Monday"]}
  Include only the keys that change. "cut" is a budget reduction % from 0 to 50. Channel names must match the summary exactly.
- Do not include an ACTION line when no plan change is suggested.`;

// Very small in-memory rate limiter: 30 AI calls per minute per IP.
const hits = new Map();
function limited(ip) {
  const now = Date.now();
  const arr = (hits.get(ip) || []).filter(t => now - t < 60000);
  arr.push(now);
  hits.set(ip, arr);
  return arr.length > 30;
}

function send(res, code, body, type = 'application/json; charset=utf-8', extra = {}) {
  res.writeHead(code, Object.assign({
    'Content-Type': type,
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'same-origin'
  }, extra));
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > MAX_BODY) { reject(new Error('Request too large')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

async function handleAI(req, res) {
  if (!GEMINI_API_KEY) {
    return send(res, 501, JSON.stringify({ error: 'AI is not configured. Set GEMINI_API_KEY on the server.' }));
  }
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();
  if (limited(ip)) return send(res, 429, JSON.stringify({ error: 'Too many AI requests. Wait a minute and try again.' }));

  let payload;
  try { payload = JSON.parse(await readBody(req)); }
  catch (e) { return send(res, 400, JSON.stringify({ error: 'Invalid request body.' })); }

  const question = String(payload.question || '').slice(0, 2000);
  const context = payload.context || {};
  const history = Array.isArray(payload.history) ? payload.history.slice(-8) : [];
  if (!question.trim()) return send(res, 400, JSON.stringify({ error: 'Empty question.' }));

  const contents = [];
  contents.push({ role: 'user', parts: [{ text: 'DATA SUMMARY (JSON):\n' + JSON.stringify(context).slice(0, 120000) }] });
  contents.push({ role: 'model', parts: [{ text: 'Understood. I will answer only from this summary.' }] });
  for (const h of history) {
    if (!h || !h.text) continue;
    contents.push({ role: h.role === 'ai' ? 'model' : 'user', parts: [{ text: String(h.text).slice(0, 4000) }] });
  }
  contents.push({ role: 'user', parts: [{ text: question }] });

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent`;
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': GEMINI_API_KEY },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] },
        contents,
        generationConfig: { temperature: 0.4, maxOutputTokens: 1500 }
      })
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) {
      const msg = (data.error && data.error.message) || ('Gemini returned ' + r.status);
      return send(res, 502, JSON.stringify({ error: msg }));
    }
    const parts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
    const text = parts.map(p => p.text || '').join('').trim() || 'No answer was returned.';
    return send(res, 200, JSON.stringify({ text, model: GEMINI_MODEL }));
  } catch (e) {
    return send(res, 502, JSON.stringify({ error: 'Could not reach Gemini: ' + e.message }));
  }
}

function serveStatic(req, res) {
  let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (p === '/') p = '/index.html';
  const file = path.normalize(path.join(PUBLIC_DIR, p));
  if (!file.startsWith(PUBLIC_DIR)) return send(res, 403, 'Forbidden', 'text/plain');
  fs.stat(file, (err, st) => {
    if (err || !st.isFile()) {
      // Single page app: unknown paths fall back to index.html
      return fs.readFile(path.join(PUBLIC_DIR, 'index.html'), (e2, buf) =>
        e2 ? send(res, 404, 'Not found', 'text/plain') : send(res, 200, buf, TYPES['.html'], { 'Cache-Control': 'no-cache' }));
    }
    const type = TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream';
    const cache = /^(text\/html|text\/javascript|text\/css)/.test(type) ? 'no-cache' : 'public, max-age=300';
    fs.readFile(file, (e, buf) => e ? send(res, 500, 'Error', 'text/plain') : send(res, 200, buf, type, { 'Cache-Control': cache }));
  });
}

const server = http.createServer((req, res) => {
  const url = req.url.split('?')[0];
  if (url === '/api/health') return send(res, 200, JSON.stringify({ ok: true }));
  if (url === '/api/config') return send(res, 200, JSON.stringify({ ai: !!GEMINI_API_KEY, model: GEMINI_API_KEY ? GEMINI_MODEL : null }));
  if (url === '/api/ai' && req.method === 'POST') return handleAI(req, res);
  if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'Method not allowed', 'text/plain');
  return serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`TV Media Planner running on port ${PORT} (AI ${GEMINI_API_KEY ? 'on: ' + GEMINI_MODEL : 'off'})`);
});
