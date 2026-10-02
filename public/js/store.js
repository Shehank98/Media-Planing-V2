// Per-browser persistence. The dataset lives in IndexedDB on this device only;
// each visitor has their own copy and nothing is sent to the server.
const DB = 'tv-planner', STORE = 'kv';

function open() {
  return new Promise((res, rej) => {
    if (!('indexedDB' in window)) return rej(new Error('IndexedDB unavailable'));
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function tx(mode, fn) {
  const db = await open();
  return new Promise((res, rej) => {
    const t = db.transaction(STORE, mode), s = t.objectStore(STORE);
    const req = fn(s);
    t.oncomplete = () => { res(req && req.result); db.close(); };
    t.onerror = () => { rej(t.error); db.close(); };
  });
}
export const idbGet = k => tx('readonly', s => s.get(k)).catch(() => null);
export const idbSet = (k, v) => tx('readwrite', s => s.put(v, k)).catch(() => null);
export const idbDel = k => tx('readwrite', s => s.delete(k)).catch(() => null);

// Small settings (plan, scenarios, theme) go to localStorage.
export function lsGet(k, d) { try { const v = localStorage.getItem('tvp:' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } }
export function lsSet(k, v) { try { localStorage.setItem('tvp:' + k, JSON.stringify(v)); } catch (e) { /* storage blocked */ } }
export function lsDel(k) { try { localStorage.removeItem('tvp:' + k); } catch (e) { /* storage blocked */ } }
