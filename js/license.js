/* License + app lock.
   - The app is activated with a signed code that only the owner (you) can make with the private key.
     The app holds only the PUBLIC key, so a shop cannot forge or extend its own licence.
   - A code carries the shop name and an expiry date. After expiry the app runs READ-ONLY
     (view and print still work; saving new data is blocked) until a renewal code is entered.
   - A daily password locks the screen so no one but the shop can open the app on that device.
   NOTE: this controls access and the renewal period; it does not hide the source code of a web app. */
'use strict';

/* Public verify key (the matching private key is only in your licence generator, never shipped). */
const LICENSE_PUBKEY = 'MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE3zh4+AziY7WOZQRA5ZfLUORoyxGpzPatxoKCEO6yvlKuL9zzfK3sdSm8CDaCksuZ8+Zp3hMbnrK3uWaD6hE6sA==';
const GRACE_DAYS = 0;        // extra days after expiry before read-only (0 = lock down on the day)

const b64uToBytes = s => { s = s.replace(/-/g, '+').replace(/_/g, '/'); while (s.length % 4) s += '='; return Uint8Array.from(atob(s), c => c.charCodeAt(0)); };
const bytesToB64u = b => btoa(String.fromCharCode(...new Uint8Array(b))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const td = new TextDecoder(), te = new TextEncoder();
let _vk;
async function verifyKey() { return _vk || (_vk = await crypto.subtle.importKey('spki', b64uToBytes(LICENSE_PUBKEY.replace(/\+/g, '-').replace(/\//g, '_')), { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify'])); }

/* code = base64url(payloadJSON) . base64url(signature) */
async function parseCode(code) {
  const parts = String(code || '').trim().replace(/\s+/g, '').split('.');
  if (parts.length !== 2) throw new Error('This licence code is not complete. Copy the whole code.');
  let payload;
  try { payload = JSON.parse(td.decode(b64uToBytes(parts[0]))); } catch (e) { throw new Error('This licence code is not valid.'); }
  const ok = await crypto.subtle.verify({ name: 'ECDSA', hash: 'SHA-256' }, await verifyKey(), b64uToBytes(parts[1]), te.encode(JSON.stringify(payload)));
  if (!ok) throw new Error('This licence code is not genuine (signature does not match). Ask for a correct code.');
  if (!payload.exp || !/^\d{4}-\d{2}-\d{2}$/.test(payload.exp)) throw new Error('This licence code has no valid expiry date.');
  return payload;   // {shop, exp, iss, id, plan?}
}

const LIC = {
  info: null,        // {shop, exp, iss, id, activatedAt}
  readOnly: false,
  today() { return DB.todayISO(); },
  daysLeft() { if (!this.info) return 0; return Math.ceil((new Date(this.info.exp + 'T23:59:59') - new Date(this.today() + 'T00:00:00')) / 86400000); },
  expired() { return this.daysLeft() < -GRACE_DAYS; },
  /* clock-tamper guard: remember the latest date seen; if the device clock jumps back, don't trust it */
  tampered() { const seen = DB.getSetting('lic_seen', ''); return seen && this.today() < seen; },
};

async function activate(code, { renew } = {}) {
  const p = await parseCode(code);
  if (renew && LIC.info && DB.normKey(p.shop) !== DB.normKey(LIC.info.shop))
    throw new Error('This code is for “' + p.shop + '”, not for this shop (' + LIC.info.shop + ').');
  if (LIC.expired() === false && LIC.info && !renew && DB.normKey(p.shop) !== DB.normKey(LIC.info.shop) && DB.D.purchases.length)
    throw new Error('This device already belongs to “' + LIC.info.shop + '”.');
  const info = { shop: p.shop, exp: p.exp, iss: p.iss || '', id: p.id || '', plan: p.plan || '', activatedAt: new Date().toISOString(), code };
  await DB.setSetting('license', info);
  LIC.info = info;
  if (p.shop && !DB.getSetting('shopName')) await DB.setSetting('shopName', p.shop);
  return info;
}

/* ---------- password (daily lock) ---------- */
async function pbkdf2(pw, saltB64) {
  const salt = saltB64 ? b64uToBytes(saltB64) : crypto.getRandomValues(new Uint8Array(16));
  const key = await crypto.subtle.importKey('raw', te.encode(pw), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt, iterations: 120000, hash: 'SHA-256' }, key, 256);
  return { salt: bytesToB64u(salt), hash: bytesToB64u(bits) };
}
async function setPassword(pw) { await DB.setSetting('lockHash', pw ? await pbkdf2(pw) : null); }
async function checkPassword(pw) { const h = DB.getSetting('lockHash', null); if (!h) return true; const r = await pbkdf2(pw, h.salt); return r.hash === h.hash; }
const hasPassword = () => !!DB.getSetting('lockHash', null);

/* ---------- full-screen gates ---------- */
function gateHTML(inner) {
  return `<div id="gate" style="position:fixed;inset:0;z-index:200;display:grid;place-items:center;padding:20px;background:radial-gradient(120% 90% at 50% 0%,var(--pri-soft),var(--bg))">
    <div class="card" style="max-width:420px;width:100%;box-shadow:0 20px 60px #0003">
      <div style="text-align:center;margin-bottom:14px"><div style="width:52px;height:52px;border-radius:14px;background:var(--pri);color:var(--pri-ink);display:grid;place-items:center;font-size:26px;margin:0 auto 8px">℞</div>
      <b style="font-size:17px">${esc(DB.getSetting('shopName', '') || 'Pharmacy Bill')}</b></div>${inner}</div></div>`;
}
function showGate(inner) { document.querySelector('.app').style.display = 'none'; $('fab').style.display = 'none'; const o = document.createElement('div'); o.innerHTML = gateHTML(inner); document.body.appendChild(o.firstElementChild); }
function closeGate() { const g = $('gate'); if (g) g.remove(); document.querySelector('.app').style.display = ''; }

function activationGate() {
  return new Promise(res => {
    showGate(`<h2 style="text-align:center;margin:0 0 4px">Activate this app</h2><p class="muted small" style="text-align:center;margin-bottom:14px">Enter the licence code given by your provider.</p>
      <label class="l">Licence code</label><textarea class="f" id="gCode" rows="3" placeholder="Paste the code here" autocomplete="off"></textarea>
      <div id="gMsg" class="small" style="margin:8px 0;min-height:18px"></div>
      <button class="btn pri" id="gGo" style="width:100%">Activate</button>
      <p class="muted small" style="margin-top:12px;text-align:center">No code yet? Contact your provider to get one.</p>`);
    $('gGo').onclick = async () => {
      $('gMsg').innerHTML = '<span class="spin"></span> Checking…';
      try { const i = await activate($('gCode').value); $('gMsg').innerHTML = '<span class="pos">Activated for ' + esc(i.shop) + '.</span>'; setTimeout(() => { closeGate(); res(); }, 500); }
      catch (e) { $('gMsg').innerHTML = '<span class="neg">' + esc(e.message) + '</span>'; }
    };
    $('gCode').onkeydown = e => { if (e.key === 'Enter' && e.ctrlKey) $('gGo').click(); };
  });
}
function lockGate() {
  return new Promise(res => {
    showGate(`<h2 style="text-align:center;margin:0 0 14px">Enter password</h2>
      <input class="f" type="password" id="gPw" placeholder="App password" autocomplete="current-password">
      <div id="gMsg" class="small neg" style="margin:8px 0;min-height:18px"></div>
      <button class="btn pri" id="gGo" style="width:100%">Unlock</button>`);
    const go2 = async () => { if (await checkPassword($('gPw').value)) { closeGate(); res(); } else { $('gMsg').textContent = 'Wrong password.'; $('gPw').select(); } };
    $('gGo').onclick = go2; $('gPw').onkeydown = e => { if (e.key === 'Enter') go2(); };
    $('gPw').focus();
  });
}

/* read-only banner + guard used by every save action */
function licenseBanner() {
  const d = LIC.daysLeft();
  if (LIC.readOnly) return `<div class="note bad" style="margin:0 0 14px">🔒 Licence expired on ${esc(DB.showDate(LIC.info.exp))} — the app is <b>read-only</b>. You can view and print, but not save new bills, sales or stock changes. <a href="#" onclick="renewDialog();return false">Enter renewal code</a>.</div>`;
  if (LIC.info && d <= 15) return `<div class="note warn" style="margin:0 0 14px">⏳ Licence ends in <b>${d} day${d === 1 ? '' : 's'}</b> (${esc(DB.showDate(LIC.info.exp))}). <a href="#" onclick="renewDialog();return false">Renew now</a>.</div>`;
  return '';
}
function guardWrite() {
  if (!LIC.readOnly) return true;
  toast('Licence expired — the app is read-only until renewed'); renewDialog(); return false;
}
function renewDialog() {
  modal(`<h3>Renew licence <button class="btn sm x" onclick="closeModal()">✕</button></h3>
    <p class="small muted">Shop: <b>${esc(LIC.info ? LIC.info.shop : '')}</b>${LIC.info ? ' · current expiry ' + esc(DB.showDate(LIC.info.exp)) : ''}. Paste the renewal code from your provider.</p>
    <label class="l">Renewal code</label><textarea class="f" id="rnCode" rows="3" autocomplete="off"></textarea>
    <div id="rnMsg" class="small" style="margin:8px 0;min-height:18px"></div>
    <div class="row-flex" style="justify-content:flex-end"><button class="btn" onclick="closeModal()">Cancel</button><button class="btn pri" id="rnGo">Apply</button></div>`);
  $('rnGo').onclick = async () => {
    $('rnMsg').innerHTML = '<span class="spin"></span> Checking…';
    try {
      const i = await activate($('rnCode').value, { renew: true });
      $('rnMsg').innerHTML = '<span class="pos">Renewed until ' + esc(DB.showDate(i.exp)) + '.</span>';
      setTimeout(() => { closeModal(); location.reload(); }, 800);
    } catch (e) { $('rnMsg').innerHTML = '<span class="neg">' + esc(e.message) + '</span>'; }
  };
}

/* called by app start before render() */
async function licenseStart() {
  LIC.info = DB.getSetting('license', null);
  if (!LIC.info) await activationGate();          // first run: must activate
  LIC.info = DB.getSetting('license', null);
  // clock-tamper or past expiry -> read-only
  LIC.readOnly = LIC.expired() || LIC.tampered();
  const seen = DB.getSetting('lic_seen', '');
  if (!seen || LIC.today() > seen) await DB.setSetting('lic_seen', LIC.today());
  if (hasPassword()) await lockGate();
}

window.LIC = LIC;
window.License = { activate, parseCode, setPassword, checkPassword, hasPassword, licenseStart, licenseBanner, guardWrite, renewDialog, LICENSE_PUBKEY };
window.renewDialog = renewDialog;
