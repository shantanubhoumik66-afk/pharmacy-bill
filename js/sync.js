/* Live sync between a shop's devices through a free Supabase project (REST / PostgREST).
   - Every synced row carries a stable id (_uid) and a changed-at time (_up).
   - On a local change the row is pushed; every few seconds new rows are pulled.
   - Newest change wins (last-write-wins by _up). Deletes travel as tombstones.
   - Medicines combine by name + batch in the stock view, so the same item from two
     devices shows as one line — no duplicates in stock. */
'use strict';

const TABLE = 'records';
const POLL_MS = 7000;

const SyncState = { state: 'off', msg: '', last: 0, pending: 0, pulling: false };
let cfg = null, timer = null, queue = new Map(), flushT = null, listeners = [];
const onStatus = fn => { listeners.push(fn); };
const emit = () => listeners.forEach(f => { try { f(SyncState); } catch (e) {} });
function setState(state, msg) { SyncState.state = state; if (msg != null) SyncState.msg = msg; emit(); }

function config() {
  const url = (DB.getSetting('syncUrl', '') || '').trim().replace(/\/+$/, '');
  const key = (DB.getSetting('syncKey', '') || '').trim();
  const shop = (DB.getSetting('syncShop', '') || '').trim();
  return url && key && shop ? { url, key, shop } : null;
}
const enabled = () => !!config();
const headers = () => ({ apikey: cfg.key, Authorization: 'Bearer ' + cfg.key, 'Content-Type': 'application/json' });
const rest = path => cfg.url + '/rest/v1/' + path;
/* fetch that gives up after `ms` instead of hanging forever (slow mobile networks) */
async function fetchT(url, opts, ms) {
  const ac = new AbortController(), t = setTimeout(() => ac.abort(), ms || 25000);
  try { return await fetch(url, { ...opts, signal: ac.signal }); }
  catch (e) { throw (e && e.name === 'AbortError') ? new Error('The connection timed out. Check the internet (turn off Data Saver / Lite mode) and try again.') : e; }
  finally { clearTimeout(t); }
}

/* a local row -> the cloud record (strip the device-local autoincrement id) */
function toRecord(store, r) {
  const data = { ...r }; delete data.id;
  return { shop: cfg.shop, store, uid: r._uid, up: r._up, del: 0, data };
}
function tombRecord(t) { return { shop: cfg.shop, store: t.store, uid: t.uid, up: t._up, del: 1, data: null }; }

async function pushRecords(recs) {
  if (!recs.length) return;
  const res = await fetchT(rest(TABLE), { method: 'POST', headers: { ...headers(), Prefer: 'resolution=merge-duplicates,return=minimal' }, body: JSON.stringify(recs) }, 30000);
  if (!res.ok) throw new Error('push ' + res.status + ' ' + (await res.text()).slice(0, 120));
}

/* pull rows changed after ts; apply them; return {max, seen:Map(store|uid -> up)} */
async function pullSince(ts, collectSeen) {
  const PAGE = 1000;
  let max = ts, count = 0, from = 0; const seen = collectSeen ? new Map() : null;
  for (;;) {
    const res = await fetchT(rest(TABLE) + '?shop=eq.' + encodeURIComponent(cfg.shop) + '&up=gt.' + ts + '&order=up.asc', { headers: { ...headers(), Range: from + '-' + (from + PAGE - 1), 'Range-Unit': 'items' } }, 30000);
    if (!res.ok && res.status !== 206) throw new Error('pull ' + res.status + ' ' + (await res.text()).slice(0, 120));
    const rows = await res.json();
    for (const row of rows) {
      max = Math.max(max, Number(row.up) || 0);
      if (seen) seen.set(row.store + '|' + row.uid, Number(row.up) || 0);
      if (!DB.SYNC_STORES.includes(row.store)) continue;
      const rec = row.del ? { _uid: row.uid, _up: Number(row.up), _del: 1 } : { ...(row.data || {}), _uid: row.uid, _up: Number(row.up) };
      await DB.applyRemote(row.store, rec);
    }
    count += rows.length;
    if (rows.length < PAGE) break;
    from += PAGE;
  }
  return { max, count, seen };
}

async function allLocalRecords() {
  const out = [];
  for (const store of DB.SYNC_STORES) for (const r of DB.D[store]) if (r._uid) out.push(toRecord(store, r));
  const tombs = await DB.tx(['synctomb'], 'readonly', async t => DB.reqP(t.objectStore('synctomb').getAll()));
  for (const tmb of tombs) out.push(tombRecord(tmb));
  return out;
}

/* first connect for this device: pull everything, then push anything the server is missing or older on */
async function reconcile() {
  const lastUp = Number(DB.getSetting('syncLastUp_' + cfg.shop, 0)) || 0;
  const full = lastUp === 0;
  const { max, seen } = await pullSince(lastUp, full);
  let newMax = max;
  if (full) {
    const mine = await allLocalRecords();
    const toPush = mine.filter(r => (seen.get(r.store + '|' + r.uid) || 0) < r.up);
    for (let i = 0; i < toPush.length; i += 500) await pushRecords(toPush.slice(i, i + 500));
    toPush.forEach(r => { newMax = Math.max(newMax, r.up); });
  }
  await DB.setSetting('syncLastUp_' + cfg.shop, newMax);
}

/* local changes are queued and flushed together */
function queueChange(store, rows) {
  if (!enabled()) return;
  for (const r of rows) queue.set(store + '|' + (r._uid || r.uid), store === 'synctomb' ? tombRecord(r) : toRecord(store, r));
  SyncState.pending = queue.size; emit();
  clearTimeout(flushT); flushT = setTimeout(flush, 800);
}
async function flush() {
  if (!cfg || !queue.size || !navigator.onLine) return;
  const recs = [...queue.values()]; queue.clear();
  try {
    await pushRecords(recs);
    const mx = recs.reduce((a, r) => Math.max(a, r.up), Number(DB.getSetting('syncLastUp_' + cfg.shop, 0)));
    await DB.setSetting('syncLastUp_' + cfg.shop, mx);
    SyncState.pending = 0; SyncState.last = Date.now(); setState('ok');
  } catch (e) {
    recs.forEach(r => queue.set(r.store + '|' + r.uid, r));   // keep for retry
    SyncState.pending = queue.size; setState(navigator.onLine ? 'error' : 'offline', e.message);
  }
}

async function poll() {
  if (!cfg || SyncState.pulling || !navigator.onLine || document.hidden) return;
  SyncState.pulling = true;
  try {
    const lastUp = Number(DB.getSetting('syncLastUp_' + cfg.shop, 0)) || 0;
    const { max, count } = await pullSince(lastUp, false);
    if (max > lastUp) await DB.setSetting('syncLastUp_' + cfg.shop, max);
    SyncState.last = Date.now(); setState('ok');
    if (count && typeof onPulled === 'function') onPulled();
  } catch (e) { setState(navigator.onLine ? 'error' : 'offline', e.message); }
  finally { SyncState.pulling = false; }
}
let onPulled = null;

async function start(afterPull) {
  stop();
  cfg = config();
  if (!cfg) { setState('off'); return; }
  onPulled = afterPull;
  DB.setLocalChange(queueChange);
  setState('connecting');
  try { await reconcile(); SyncState.last = Date.now(); setState('ok'); if (afterPull) afterPull(); }
  catch (e) { setState(navigator.onLine ? 'error' : 'offline', e.message); }
  timer = setInterval(poll, POLL_MS);
  await flush();
}
function stop() { if (timer) clearInterval(timer); timer = null; }

addEventListener('online', () => { if (cfg) { flush(); poll(); } });
addEventListener('focus', () => { if (cfg) poll(); });
document.addEventListener('visibilitychange', () => { if (cfg && !document.hidden) poll(); });

/* check the project before saving the settings */
async function test(url, key, shop) {
  url = url.trim().replace(/\/+$/, ''); key = key.trim();
  const res = await fetchT(url + '/rest/v1/' + TABLE + '?select=shop&limit=1', { headers: { apikey: key, Authorization: 'Bearer ' + key } }, 15000);
  if (res.status === 404) throw new Error('Table “records” not found. Run the setup SQL in Supabase first.');
  if (res.status === 401 || res.status === 403) throw new Error('Key not accepted. Check the anon key, and that the table allows anon access.');
  if (!res.ok) throw new Error('Could not reach the project (' + res.status + '). Check the Project URL.');
  return true;
}

window.Sync = { start, stop, test, enabled, config, state: SyncState, onStatus, poll, flush };
