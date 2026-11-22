/* Sales: Sales Invoice, Sales (one row per invoice), Sale Items (product wise), print, GST (GSTR-1 summaries).
   Same rules as the workbook macros (Sales Invoice, InvoicePrint and GST sheets). */
'use strict';

/* ---------- rules copied from the workbook ---------- */
const effPacks = (strip, loose, pack) => N(strip) + (N(pack) > 0 ? N(loose) / N(pack) : 0);
const lineQty = l => effPacks(l.strip, N(l.loose) + N(l.free), l.pack);   // what leaves the stock (free pieces too)
const billQty = l => effPacks(l.strip, l.loose, l.pack);                  // what is charged (free pieces are not)
/* Net = (strip + loose/pack) x MRP x (1 - Disc%/100), rounded to the paisa. MRP already includes GST. */
const saleNet = l => DB.r2x(billQty(l) * N(l.mrp) * (1 - N(l.disc) / 100));
const gstIn = (net, g) => DB.r2x(N(net) * N(g) / (100 + N(g)));             // GST inside an MRP amount
function saleTotals(lines) {
  let total = 0, net = 0, gst = 0;
  for (const l of lines) {
    const q = billQty(l); if (!(q > 0)) continue;
    const n = saleNet(l);
    total += DB.r2x(q * N(l.mrp)); net += n; gst += gstIn(n, l.gst);
  }
  return { total: DB.r2x(total), gst: DB.r2x(gst), disc: DB.r2x(total - net), net: DB.r2x(net) };
}
const invKey = s => String(s || '').toUpperCase().trim();
function nextInvoiceNo() {
  let mx = 0;
  for (const x of [...DB.D.saleItems, ...DB.D.sales]) { const s = invKey(x.invoiceNo); if (s.startsWith('INV-') && /^\d+$/.test(s.slice(4))) mx = Math.max(mx, +s.slice(4)); }
  return 'INV-' + String(mx + 1).padStart(4, '0');
}
/* Packs in stock for a medicine + batch. The invoice being modified gives its own lines back first. */
function available(item, batch, editInv) {
  const k = DB.normKey(item) + '|' + DB.batchKey(batch);
  const s = DB.stock().rows.find(r => r.key === k);
  let a = s ? s.rem : 0;
  if (editInv) for (const x of DB.D.saleItems) if (invKey(x.invoiceNo) === invKey(editInv) && DB.normKey(x.item) + '|' + DB.batchKey(x.batch) === k) a += N(x.qtyPacks);
  return a;
}
/* Pack size used last time this medicine was sold (saves typing) */
function lastPack(item) {
  const k = DB.normKey(item);
  for (let i = DB.D.saleItems.length - 1; i >= 0; i--) { const x = DB.D.saleItems[i]; if (DB.normKey(x.item) === k && N(x.pack) > 0) return N(x.pack); }
  const m = String(item).match(/\b(\d{1,3})\s*'?S\b|\b1\s*X\s*(\d{1,3})\b|\b(\d{1,3})\s*TAB/i);
  return m ? +(m[1] || m[2] || m[3]) : '';
}

/* ---------- numbers in words (Indian system), as on the printed bill ---------- */
function rupeesInWords(amt) {
  amt = DB.r2x(amt);
  let r = Math.floor(amt), p = Math.round((amt - r) * 100);
  if (p >= 100) { r++; p = 0; }
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const u100 = n => n < 20 ? ones[n] : tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
  const words = n => {
    let s = '';
    if (n >= 1e7) { s += words(Math.floor(n / 1e7)) + ' Crore '; n %= 1e7; }
    if (n >= 1e5) { s += u100(Math.floor(n / 1e5)) + ' Lakh '; n %= 1e5; }
    if (n >= 1e3) { s += u100(Math.floor(n / 1e3)) + ' Thousand '; n %= 1e3; }
    if (n >= 100) { s += u100(Math.floor(n / 100)) + ' Hundred '; n %= 100; }
    if (n > 0) s += u100(n);
    return s.trim();
  };
  return (r ? 'Rupees ' + words(r) : 'Rupees Zero') + (p > 0 ? ' and ' + words(p) + ' Paise' : '') + ' Only.';
}

/* ================= SALES INVOICE ================= */
let SI = null;
const blankSale = () => ({ item: '', mfg: '', hsn: '', batch: '', exp: '', mrp: '', pack: '', strip: '', loose: '', free: '', gst: '', disc: '' });
function newSale() { return { patient: '', address: '', doctor: '', drReg: '', invoiceNo: nextInvoiceNo(), date: DB.todayISO(), lines: [blankSale()] }; }
function saveSaleDraft() { clearTimeout(saveSaleDraft._h); saveSaleDraft._h = setTimeout(() => DB.setSetting('saleDraft', SI), 600); }

PAGES.sale = q => {
  const edit = q.get('edit');
  if (edit) {
    const h = DB.D.sales.find(s => invKey(s.invoiceNo) === invKey(edit));
    const ls = DB.D.saleItems.filter(x => invKey(x.invoiceNo) === invKey(edit));
    if (!h && !ls.length) return go('sales');
    SI = { editInv: edit, patient: h ? h.patient : (ls[0] || {}).patient || '', address: h ? h.address : '', doctor: h ? h.doctor : '', drReg: h ? h.drReg : '',
      invoiceNo: edit, date: (h && h.date) || (ls[0] || {}).date || DB.todayISO(),
      lines: ls.map(x => ({ item: x.item, mfg: x.mfg, hsn: x.hsn, batch: x.batch, exp: x.exp, mrp: x.mrp, pack: x.pack || '', strip: x.strip || (!x.pack && !x.loose ? x.qtyPacks : ''),
        loose: x.loose || '', free: x.freePcs || '', gst: x.gst, disc: x.disc || '' })) };
    if (!SI.lines.length) SI.lines.push(blankSale());
  } else if (!SI || SI.editInv) {
    SI = DB.getSetting('saleDraft', null);
    if (!SI || SI.editInv) SI = newSale();
    if (!SI.lines.some(l => l.item)) SI.invoiceNo = nextInvoiceNo();
  }
  drawSale();
};

const SCOLS = [['item', 'Product Name', 'wide', 0, 230], ['mfg', 'MFG', 'half', 0, 80], ['hsn', 'HSN Code', 'half', 0, 90], ['batch', 'Batch No', 'half', 0, 100], ['exp', 'Exp', '', 0, 56],
  ['mrp', 'MRP', '', 1, 74], ['pack', 'Pack', '', 1, 52], ['strip', 'Strip Qty', '', 1, 62], ['loose', 'Loose (pcs)', '', 1, 70], ['free', 'Free (pcs)', '', 1, 64], ['gst', 'GST %', '', 1, 52], ['disc', 'Disc %', '', 1, 56]];
function drawSale() {
  setTop(SI.editInv ? 'Modify ' + SI.editInv : 'Sales Invoice', `<button class="btn" id="siClear">Clear</button><button class="btn pri" id="siSave">💾 <span class="lbl">Save</span></button><button class="btn pri" id="siSavePrint">🖨 <span class="lbl">Save &amp; print</span></button>`);
  $('page').innerHTML = `${SI.editInv ? `<div class="note warn">Modifying ${esc(SI.editInv)}. Saving replaces the old invoice; stock is corrected automatically.</div>` : ''}
    <div class="card"><div class="fields">
      <div><label class="l">Patient Name *</label><input class="f" id="siP" value="${esc(SI.patient)}" list="patList" autocomplete="off"><datalist id="patList">${[...new Set(DB.D.sales.map(s => s.patient).filter(Boolean))].sort().map(n => `<option value="${esc(n)}">`).join('')}</datalist></div>
      <div><label class="l">Address</label><input class="f" id="siA" value="${esc(SI.address)}"></div>
      <div><label class="l">Doctor Name</label><input class="f" id="siD" value="${esc(SI.doctor)}" list="docList" autocomplete="off"><datalist id="docList">${[...new Set(DB.D.sales.map(s => s.doctor).filter(Boolean))].sort().map(n => `<option value="${esc(n)}">`).join('')}</datalist></div>
      <div><label class="l">Dr. Register No</label><input class="f" id="siR" value="${esc(SI.drReg)}"></div>
      <div><label class="l">Invoice No</label><input class="f" id="siN" value="${esc(SI.invoiceNo)}" ${SI.editInv ? 'readonly' : ''}></div>
      <div><label class="l">Invoice Date</label><input class="f" type="date" id="siT" value="${esc(SI.date)}"></div>
    </div></div>
    <div class="card" style="padding:10px">
      <div class="tw lines-w"><table class="t lines"><thead><tr><th>Sr</th>${SCOLS.map(c => `<th class="${c[3] ? 'n' : ''}" style="min-width:${c[4]}px">${c[1]}</th>`).join('')}<th class="n" style="min-width:90px">Net Amount</th><th></th></tr></thead><tbody id="siBody"></tbody></table></div>
      <div class="row-flex" style="margin-top:10px"><button class="btn sm" id="siAdd">+ Add product</button><span class="muted small">Type 2+ letters of a medicine and pick the batch. Pack, then Strip Qty and/or Loose pieces.</span></div>
      <div id="siTot" style="display:flex;justify-content:flex-end;margin-top:12px"></div>
    </div>`;
  drawSaleLines();
  const bind = (id, k) => $(id).oninput = e => { SI[k] = e.target.value; saveSaleDraft(); };
  bind('siP', 'patient'); bind('siA', 'address'); bind('siD', 'doctor'); bind('siR', 'drReg'); bind('siN', 'invoiceNo');
  $('siP').onchange = e => { // a returning patient: fill address and doctor from the last invoice
    const last = DB.D.sales.filter(s => DB.normKey(s.patient) === DB.normKey(e.target.value)).sort((a, b) => (b.date || '').localeCompare(a.date || ''))[0];
    if (last) { if (!SI.address) SI.address = $('siA').value = last.address || ''; if (!SI.doctor) SI.doctor = $('siD').value = last.doctor || ''; if (!SI.drReg) SI.drReg = $('siR').value = last.drReg || ''; }
    saveSaleDraft();
  };
  $('siT').onchange = e => { SI.date = e.target.value; saveSaleDraft(); };
  $('siAdd').onclick = () => { SI.lines.push(blankSale()); drawSaleLines(); const r = $('siBody').querySelectorAll('tr.ln'); r[r.length - 1].querySelector('input').focus(); };
  $('siSave').onclick = () => saveSale(false);
  $('siSavePrint').onclick = () => saveSale(true);
  $('siClear').onclick = async () => { if (await ask('Clear invoice', 'Clear this invoice? Nothing is saved.', 'Clear', true)) { SI = newSale(); DB.setSetting('saleDraft', null); go('sale'); drawSale(); } };
}
function drawSaleLines() {
  $('siBody').innerHTML = SI.lines.map((l, i) => `<tr class="ln" data-i="${i}"><td class="sr">${i + 1}</td>${SCOLS.map(c =>
    `<td class="${c[2]}" data-l="${c[1]}"><input data-k="${c[0]}" class="${c[3] ? 'num' : ''}" ${c[3] ? 'inputmode="decimal"' : ''} value="${esc(l[c[0]])}" autocomplete="off" ${c[0] !== 'item' && c[0] !== 'pack' && c[0] !== 'strip' && c[0] !== 'loose' && c[0] !== 'free' && c[0] !== 'disc' ? 'tabindex="-1"' : ''} ${c[0] === 'item' ? 'placeholder="Type medicine…"' : ''}></td>`).join('')}
    <td class="netc half" data-l="Net Amount">${l.item ? money(saleNet(l)) : ''}</td><td data-l=""><button class="btn sm" data-del="${i}" title="Remove">✕</button></td></tr>`).join('');
  const body = $('siBody');
  body.querySelectorAll('input').forEach(inp => {
    const tr = inp.closest('tr'), i = +tr.dataset.i, k = inp.dataset.k;
    inp.oninput = () => { SI.lines[i][k] = k === 'item' || k === 'batch' ? inp.value.toUpperCase() : inp.value; refreshSaleRow(i, inp); saveSaleDraft(); };
    inp.onchange = () => { if (k === 'item') pickBatch(i, inp.value); else if (['pack', 'strip', 'loose', 'free'].includes(k)) checkQty(i, k); else if (k === 'disc') { const d = N(inp.value); if (d < 0 || d > 100) { toast('Discount must be between 0 and 100 (percent)'); SI.lines[i].disc = inp.value = ''; refreshSaleRow(i); } } };
    inp.onkeydown = e => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      if (k === 'item') return pickBatch(i, inp.value);
      const order = ['pack', 'strip', 'loose', 'free', 'disc'], at = order.indexOf(k);
      if (at >= 0 && at < order.length - 1) return focusSale(i, order[at + 1]);
      if (i === SI.lines.length - 1) { SI.lines.push(blankSale()); drawSaleLines(); }
      focusSale(i + 1, 'item');
    };
  });
  body.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { SI.lines.splice(+b.dataset.del, 1); if (!SI.lines.length) SI.lines.push(blankSale()); drawSaleLines(); saveSaleDraft(); });
  drawSaleTotals();
}
function focusSale(i, k) { const el = $('siBody').querySelector(`tr[data-i="${i}"] input[data-k="${k}"]`); if (el) { el.focus(); el.select && el.select(); } }
function refreshSaleRow(i, typing) {
  const tr = $('siBody').querySelector(`tr.ln[data-i="${i}"]`); if (!tr) return;
  const l = SI.lines[i];
  tr.querySelectorAll('input[data-k]').forEach(inp => { const v = String(l[inp.dataset.k] == null ? '' : l[inp.dataset.k]); if (inp !== typing && inp.value !== v) inp.value = v; });
  tr.querySelector('.netc').textContent = l.item && l.batch ? money(saleNet(l)) : '';
  drawSaleTotals();
}
function drawSaleTotals() {
  const t = saleTotals(SI.lines.filter(l => l.item));
  $('siTot').innerHTML = `<table class="t" style="width:auto;min-width:260px"><tbody>
    <tr><td>Total Amount</td><td class="n">${money(t.total)}</td></tr><tr><td>GST Amount</td><td class="n">${money(t.gst)}</td></tr>
    <tr><td>Discount Amount</td><td class="n">${money(t.disc)}</td></tr><tr><td><b>NET AMOUNT</b></td><td class="n"><b style="font-size:17px">${money(t.net)}</b></td></tr></tbody></table>`;
}
/* type medicine -> list of batches with stock (soonest expiry first) -> row filled */
function pickBatch(i, txt) {
  txt = String(txt || '').trim();
  const l = SI.lines[i];
  if (!txt) { SI.lines[i] = blankSale(); refreshSaleRow(i); return; }
  if (l.batch && DB.normKey(l.item) === DB.normKey(txt)) return focusSale(i, 'pack');
  if (txt.length < 2) { toast('Type at least 2 letters of the product name'); return; }
  const used = k => SI.lines.reduce((a, x, j) => j !== i && x.item && DB.normKey(x.item) + '|' + DB.batchKey(x.batch) === k ? a + lineQty(x) : a, 0);
  const opts = DB.stock().rows.filter(r => DB.matchWords(r.name, txt)).map(r => ({ ...r, av: available(r.name, r.batch === '-' ? '' : r.batch, SI.editInv) - used(r.key) }))
    .filter(r => r.av > 0.0001).sort((a, b) => a.name.localeCompare(b.name) || (a.exp || 9e9) - (b.exp || 9e9));
  if (!opts.length) { toast('No product with available stock found for “' + txt + '”'); return; }
  if (opts.length > 400) { toast(opts.length + ' batches match. Type more letters.'); return; }
  const ym = new Date().getFullYear() * 100 + new Date().getMonth() + 1;
  modal(`<h3>Select medicine <button class="btn sm x" onclick="closeModal()">✕</button></h3>
    <input class="f" id="pkQ" value="${esc(txt)}" autocomplete="off" style="margin-bottom:10px">
    <div class="tw" style="max-height:60vh"><table class="t"><thead><tr><th>Medicine</th><th>Batch</th><th>Exp</th><th class="n">MRP</th><th class="n">Stock</th></tr></thead><tbody id="pkB"></tbody></table></div>
    <p class="muted small">↑ ↓ to move, Enter to select. Soonest expiry is listed first for each medicine.</p>`, true);
  let sel = 0, shown = opts;
  const draw = () => {
    $('pkB').innerHTML = shown.slice(0, 300).map((r, k) => `<tr class="click" data-k="${k}" style="${k === sel ? 'outline:2px solid var(--pri);outline-offset:-2px' : ''}"><td>${esc(r.name)}</td><td>${esc(r.batch)}</td>
      <td>${r.exp ? `<span class="pill ${r.exp < ym ? 'bad' : r.exp <= ym + 3 ? 'warn' : ''}">${DB.expShow(r.exp)}</span>` : ''}</td><td class="n">${r.mrp ? r.mrp.toFixed(2) : ''}</td><td class="n"><b>${qty0(r.av)}</b></td></tr>`).join('') || '<tr><td colspan="5" class="empty">No match</td></tr>';
    $('pkB').querySelectorAll('tr[data-k]').forEach(tr => tr.onclick = () => choose(shown[+tr.dataset.k]));
  };
  const choose = r => {
    if (!r) return;
    if (r.exp && r.exp < ym) toast('Warning: this batch expired in ' + DB.expShow(r.exp), 4000);
    closeModal();
    Object.assign(SI.lines[i], { item: r.name, mfg: r.mfg || '', hsn: r.hsn || '', batch: r.batch === '-' ? '' : r.batch, exp: r.exp ? DB.expShow(r.exp) : '', mrp: r.mrp || '', gst: r.gst, pack: lastPack(r.name), strip: '', loose: '', free: '', disc: '' });
    refreshSaleRow(i); saveSaleDraft(); focusSale(i, 'pack');
  };
  draw();
  const qi = $('pkQ'); qi.focus(); qi.select();
  qi.oninput = () => { shown = opts.filter(r => DB.matchWords(r.name + ' ' + r.batch, qi.value)); sel = 0; draw(); };
  qi.onkeydown = e => {
    if (e.key === 'ArrowDown') { sel = Math.min(sel + 1, shown.length - 1); draw(); e.preventDefault(); }
    else if (e.key === 'ArrowUp') { sel = Math.max(sel - 1, 0); draw(); e.preventDefault(); }
    else if (e.key === 'Enter') { e.preventDefault(); choose(shown[sel]); }
  };
}
function checkQty(i, k) {
  const l = SI.lines[i];
  for (const f of ['pack', 'strip', 'loose', 'free']) if (N(l[f]) < 0 || (l[f] !== '' && isNaN(+l[f]))) { toast('Quantity must be a positive number'); l[f] = ''; refreshSaleRow(i); return focusSale(i, f); }
  if (!l.item) { toast('Select the batch first: type the product name and press Enter'); return; }
  if ((N(l.loose) > 0 || N(l.free) > 0) && !(N(l.pack) > 0)) { toast('Enter the Pack size first (for example 15) to sell loose or free pieces'); return focusSale(i, 'pack'); }
  const k2 = DB.normKey(l.item) + '|' + DB.batchKey(l.batch);
  const others = SI.lines.reduce((a, x, j) => j !== i && x.item && DB.normKey(x.item) + '|' + DB.batchKey(x.batch) === k2 ? a + lineQty(x) : a, 0);
  const av = available(l.item, l.batch, SI.editInv) - others;
  if (lineQty(l) > av + 0.0001) {
    toast('Only ' + qty0(av) + ' pack(s) in stock' + (N(l.pack) > 0 ? ' (about ' + Math.round(av * N(l.pack)) + ' pieces)' : '') + ' for ' + l.item + ' (batch ' + (l.batch || '-') + ')', 5000);
    l[k] = ''; refreshSaleRow(i); return focusSale(i, k);
  }
  refreshSaleRow(i);
}
async function saveSale(print) {
  if (!License.guardWrite()) return;
  const pn = String(SI.patient || '').trim();
  if (!pn) { toast('Enter the patient name first'); return $('siP').focus(); }
  const ls = SI.lines.filter(l => String(l.item).trim());
  if (!ls.length) return toast('There is nothing to save. Add at least one product.');
  const need = new Map();
  for (let i = 0; i < SI.lines.length; i++) {
    const l = SI.lines[i]; if (!String(l.item).trim()) continue;
    if (!(lineQty(l) > 0)) { toast('Line ' + (i + 1) + ' (' + l.item + '): enter the quantity first'); return focusSale(i, 'strip'); }
    const k = DB.normKey(l.item) + '|' + DB.batchKey(l.batch);
    need.set(k, (need.get(k) || { q: 0, l })); need.get(k).q += lineQty(l);
  }
  for (const [, x] of need) {
    const av = available(x.l.item, x.l.batch, SI.editInv);
    if (x.q > av + 0.0001) return toast('Not enough stock for ' + x.l.item + ' (batch ' + (x.l.batch || '-') + '). Available ' + qty0(av) + ', invoice needs ' + qty0(x.q) + '.', 6000);
  }
  let inv = SI.editInv || String(SI.invoiceNo || '').trim();
  if (!SI.editInv && (!inv || DB.D.sales.some(s => invKey(s.invoiceNo) === invKey(inv)) || DB.D.saleItems.some(s => invKey(s.invoiceNo) === invKey(inv)))) inv = nextInvoiceNo();
  if (SI.editInv) {
    await DB.delMany('saleItems', DB.D.saleItems.filter(x => invKey(x.invoiceNo) === invKey(inv)).map(x => x.id));
  }
  const date = SI.date || DB.todayISO();
  const rows = ls.map(l => ({
    date, item: DB.normKey(l.item), qtyPacks: lineQty(l), batch: String(l.batch || '').toUpperCase().trim() || '-', invoiceNo: inv, mrp: N(l.mrp), disc: N(l.disc), net: saleNet(l),
    pack: N(l.pack), strip: N(l.strip), loose: N(l.loose), exp: l.exp || '', gst: N(l.gst), patient: pn, mfg: l.mfg || '', hsn: String(l.hsn || ''), freePcs: N(l.free),
  }));
  await DB.putMany('saleItems', rows);
  const t = saleTotals(ls);
  await DB.putMany('sales', [{ patient: pn, invoiceNo: inv, date, total: t.net, address: SI.address || '', doctor: SI.doctor || '', drReg: SI.drReg || '' }]);
  const wasEdit = !!SI.editInv;
  SI = newSale(); await DB.setSetting('saleDraft', null);
  toast('Invoice ' + inv + (wasEdit ? ' modified' : ' saved') + ' for ' + pn + ' (' + rows.length + ' line(s)). Stock updated.', 4000);
  if (print) printInvoice(inv);
  if (wasEdit) go('sales'); else { go('sale'); drawSale(); }
}

/* ================= PRINT (CASH MEMO / PRESCRIPTION REGISTER) ================= */
function invoiceHTML(inv) {
  const h = DB.D.sales.find(s => invKey(s.invoiceNo) === invKey(inv)) || {};
  const ls = DB.D.saleItems.filter(x => invKey(x.invoiceNo) === invKey(inv));
  const S = k => esc(DB.getSetting(k, '') || '');
  let tTotal = 0, tNet = 0, tGst = 0;
  const rows = ls.map((x, i) => {
    let st = N(x.strip); const pk = N(x.pack), lo = N(x.loose), fr = N(x.freePcs);
    if (!pk && !st && !lo) st = N(x.qtyPacks);            // old records
    const gross = DB.r2x((pk > 0 ? st + lo / pk : st) * N(x.mrp));
    tTotal += gross; tNet += N(x.net); tGst += gstIn(x.net, x.gst);
    const loose = [lo > 0 ? lo + ' piece' : '', fr > 0 ? '+' + fr + ' free' : ''].filter(Boolean).join(' ');
    return `<tr><td>${i + 1}</td><td class="l">${esc(x.item)}</td><td>${esc(x.mfg)}</td><td>${pk || ''}</td><td>${esc(x.batch)}</td><td>${esc(x.exp)}</td><td>${esc(x.hsn)}</td><td>${st || ''}</td><td>${loose}</td>
      <td class="r">${N(x.mrp).toFixed(2)}</td><td>${N(x.disc) ? N(x.disc) + '%' : ''}</td><td>${N(x.gst) / 2}</td><td>${N(x.gst) / 2}</td><td class="r">${gross.toFixed(2)}</td></tr>`;
  });
  for (let i = ls.length; i < 8; i++) rows.push('<tr class="blank">' + '<td></td>'.repeat(14) + '</tr>');
  const d = h.date || (ls[0] || {}).date || '';
  const dt = d ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(2, 4) : '';
  const terms = (DB.getSetting('terms', DEFAULT_TERMS) || '').split('\n').filter(Boolean);
  const shop = S('shopName') || 'Pharmacy';
  return `<div class="memo"><div class="c top">CASH MEMO/PRESCRIPTION REGISTER</div>
  <div class="box"><div class="hd"><div class="shop"><b class="big">${shop}</b><div>${S('shopAddress')}</div><div>${S('shopPhone') ? 'Phone: ' + S('shopPhone') : ''}</div><div>${S('shopEmail') ? 'EMAIL: ' + S('shopEmail') : ''}</div></div>
    <div class="pt"><div><span>Patient :</span><b>${esc(h.patient || (ls[0] || {}).patient)}</b></div><div><span>PT.Address</span>${esc(h.address)}</div><div><span>Phone No.</span></div><div><span>ReferBy :</span>${esc(h.doctor)}</div><div><span>Dr. Reg. No. :</span>${esc(h.drReg)}</div></div></div>
  <div class="hd2"><div><div>GSTIN: ${S('shopGstin')}</div><div>Licence No. : ${S('shopLic1')}</div><div>Licence No. : ${S('shopLic2')}</div></div><div class="c chem">CHEMISTS<br>&amp;<br>DRUGIST</div>
    <div class="inv"><div><b>INVOICE NO :</b> <b>${esc(inv)}</b></div><div><b>Printed By :</b> <b>${S('printedBy') || 'ADMIN'}</b></div><div><b>TIME :</b> ${new Date().toTimeString().slice(0, 5)}</div></div><div class="dt">Date : ${dt}<br>Mobile</div></div>
  <table class="it"><thead><tr><th>SN.</th><th>PRODUCT NAME</th><th>MFG</th><th>PACK</th><th>BATCH</th><th>EXPIRY</th><th>HSN</th><th>QTY</th><th>LOOSE QTY</th><th>MRP</th><th>SHD</th><th>CGST</th><th>SGST</th><th>AMOUNT</th></tr></thead><tbody>${rows.join('')}</tbody></table>
  <div class="ft"><div class="terms"><b>Terms &amp; Conditions</b>${terms.map(t => `<div>${esc(t)}</div>`).join('')}<div class="for">For ${shop}</div></div>
    <table class="sum"><tr><td>TOTAL</td><td>${DB.r2x(tTotal).toFixed(2)}</td></tr><tr><td>DISC</td><td>${DB.r2x(tTotal - tNet).toFixed(2)}</td></tr><tr><td>CGST</td><td>${DB.r2x(tGst / 2).toFixed(2)}</td></tr>
      <tr><td>SGST</td><td>${DB.r2x(tGst / 2).toFixed(2)}</td></tr><tr><td>OTH. ADJ</td><td>0.00</td></tr><tr><td><b>NET AMT</b></td><td><b>${DB.r2x(tNet).toFixed(2)}</b></td></tr></table></div>
  <div class="words"><span>${rupeesInWords(DB.r2x(tNet))}</span><span>${S('signText') || 'For Sig of Qualified Person'}</span></div></div></div>`;
}
const DEFAULT_TERMS = '* WEDNESDAY CLOSED*\nTake medicine as per Doctor\'s advice\nPlease check your medicine before leaving the shop\nPlease bring our Bill while return of our medicine';
const PRINT_CSS = `@page{size:A4;margin:10mm}body{font:11px Arial,sans-serif;color:#000;margin:0}.memo{max-width:190mm;margin:0 auto}.c{text-align:center}.top{font-size:12px;margin-bottom:3px}
.box{border:1px solid #000}.hd{display:flex;border-bottom:1px solid #000}.shop{flex:1;padding:6px 8px;line-height:1.5}.big{font-size:15px}.pt{flex:1;padding:6px 8px;line-height:1.6}.pt span{display:inline-block;width:92px}
.hd2{display:flex;border-bottom:1px solid #000;align-items:center}.hd2>div{padding:5px 8px;line-height:1.6}.hd2>div:first-child{flex:1.2}.chem{flex:.8;font-weight:bold}.inv{flex:1.2;border-left:1px solid #000}.dt{flex:.8}
.it{width:100%;border-collapse:collapse}.it th,.it td{border:1px solid #000;padding:3px 4px;text-align:center;font-size:10.5px}.it th{font-size:10px}.it td.l{text-align:left}.it td.r{text-align:right}.it tr.blank td{height:16px;border-top:0;border-bottom:0}
.it{border-left:0;border-right:0}.it th:first-child,.it td:first-child{border-left:0}.it th:last-child,.it td:last-child{border-right:0}
.ft{display:flex;border-top:1px solid #000}.terms{flex:1;padding:6px 8px;line-height:1.6;position:relative}.for{position:absolute;right:12px;bottom:6px}.sum{border-collapse:collapse;min-width:190px}.sum td{padding:2px 8px;text-align:right}.sum td:first-child{width:90px}
.words{display:flex;justify-content:space-between;border-top:1px solid #000;padding:6px 8px}`;
function printInvoice(inv) {
  const w = window.open('', '_blank');
  if (!w) { toast('Allow pop-ups to print the invoice'); return; }
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(inv)}</title><style>${PRINT_CSS}</style></head><body>${invoiceHTML(inv)}<script>setTimeout(()=>print(),300)<\/script></body></html>`);
  w.document.close();
}
function viewInvoice(inv) {
  const h = DB.D.sales.find(s => invKey(s.invoiceNo) === invKey(inv)) || {};
  modal(`<h3>${esc(inv)} · ${esc(h.patient || '')} <button class="btn sm x" onclick="closeModal()">✕</button></h3>
    <div style="background:#fff;color:#000;padding:10px;border-radius:10px;overflow:auto"><style>${PRINT_CSS.replace(/@page[^}]*}/, '').replace(/body\{[^}]*}/, '')}</style>${invoiceHTML(inv)}</div>
    <div class="row-flex" style="margin-top:14px"><button class="btn bad" id="ivDel">Delete</button><span class="grow"></span><button class="btn" id="ivMod">✎ Modify</button><button class="btn pri" id="ivPr">🖨 Print</button></div>`, true);
  $('ivPr').onclick = () => printInvoice(inv);
  $('ivMod').onclick = () => { closeModal(); go('sale', { edit: inv }); };
  $('ivDel').onclick = async () => {
    if (!License.guardWrite()) return;
    const ls = DB.D.saleItems.filter(x => invKey(x.invoiceNo) === invKey(inv));
    if (!await ask('Delete invoice', 'Delete invoice ' + inv + ' of ' + (h.patient || '') + '?\nStock of its products will be restored. This cannot be undone.', 'Delete', true)) return;
    await DB.delMany('saleItems', ls.map(x => x.id));
    await DB.delMany('sales', DB.D.sales.filter(s => invKey(s.invoiceNo) === invKey(inv)).map(s => s.invoiceNo));
    closeModal(); toast('Invoice ' + inv + ' deleted (' + ls.length + ' product line(s) removed)'); render();
  };
}

/* ================= SALES (one row per invoice) ================= */
function periodCard(from, to, s, page, label) {
  return `<div class="card"><div class="fields"><div><label class="l">From date</label><input class="f" type="date" id="pF" value="${esc(from)}"></div><div><label class="l">To date</label><input class="f" type="date" id="pT" value="${esc(to)}"></div>
    ${label ? `<div><label class="l">${label}</label><input class="f" id="pS" value="${esc(s)}"></div>` : ''}</div>
    <div class="chips" style="margin-top:10px"><button class="chip" data-p="t">Today</button><button class="chip" data-p="0">This month</button><button class="chip" data-p="1">Last month</button><button class="chip" data-p="a">All</button></div></div>`;
}
function wirePeriod(page, extra) {
  const upd = () => go(page, Object.fromEntries(Object.entries({ from: $('pF').value, to: $('pT').value, s: $('pS') ? $('pS').value : '', ...(extra || {}) }).filter(x => x[1])));
  $('pF').onchange = upd; $('pT').onchange = upd; if ($('pS')) $('pS').onchange = upd;
  document.querySelectorAll('[data-p]').forEach(c => c.onclick = () => {
    const d = new Date(), y = d.getFullYear(), m = d.getMonth(), v = c.dataset.p, iso = DB.toISO;
    if (v === 'a') return go(page, extra);
    if (v === 't') return go(page, { from: iso(d), to: iso(d), ...(extra || {}) });
    go(page, { from: iso(new Date(y, m - +v, 1)), to: iso(+v ? new Date(y, m, 0) : d), ...(extra || {}) });
  });
}
PAGES.sales = q => {
  const from = q.get('from') || '', to = q.get('to') || '', s = q.get('s') || '';
  setTop('Sales', `<button class="btn pri" onclick="go('sale')">＋ <span class="lbl">New invoice</span></button>`);
  const cnt = {}; DB.D.saleItems.forEach(x => cnt[invKey(x.invoiceNo)] = (cnt[invKey(x.invoiceNo)] || 0) + 1);
  const rows = DB.D.sales.filter(x => (!from || (x.date || '') >= from) && (!to || (x.date || '') <= to) && DB.matchWords(x.patient + ' ' + x.invoiceNo + ' ' + (x.doctor || ''), s))
    .sort((a, b) => (b.date || '').localeCompare(a.date || '') || invKey(b.invoiceNo).localeCompare(invKey(a.invoiceNo)));
  const tot = rows.reduce((a, x) => a + N(x.total), 0);
  $('page').innerHTML = periodCard(from, to, s, 'sales', 'Patient, invoice or doctor') +
    `<div class="note">${rows.length} invoice(s) · total ${money(tot)}. Tap an invoice to view, print, modify or delete it.</div>` +
    table('sl', [{ h: 'Patient Name', f: x => esc(x.patient) }, { h: 'Invoice No', f: x => esc(x.invoiceNo) }, { h: 'Invoice Date', f: x => esc(DB.showDate(x.date)) }, { h: 'Items', n: 1, f: x => cnt[invKey(x.invoiceNo)] || 0 },
      { h: 'Total Invoice Amount', n: 1, f: x => money(x.total) }, { h: 'Patient Address', f: x => esc(x.address) }, { h: 'Doctor Name', f: x => esc(x.doctor) }, { h: 'Dr. Register No', f: x => esc(x.drReg) }],
      rows, { click: 1, foot: ['TOTAL', rows.length + ' invoices', '', '', money(tot), '', '', ''], empty: 'No sales in this period.' });
  wirePeriod('sales');
  wireTable('sl', rows, x => viewInvoice(x.invoiceNo), () => PAGES.sales(route().q));
};

/* ================= SALE ITEMS (product wise) ================= */
PAGES.saleitems = q => {
  const from = q.get('from') || '', to = q.get('to') || '', s = q.get('s') || '';
  setTop('Sale Items');
  const rows = DB.D.saleItems.filter(x => (!from || (x.date || '') >= from) && (!to || (x.date || '') <= to) && DB.matchWords(x.item + ' ' + x.batch + ' ' + x.invoiceNo + ' ' + x.patient, s))
    .sort((a, b) => (b.date || '').localeCompare(a.date || '') || (b.id || 0) - (a.id || 0));
  const sum = k => rows.reduce((a, x) => a + N(x[k]), 0);
  $('page').innerHTML = periodCard(from, to, s, 'saleitems', 'Medicine, batch, invoice or patient') +
    table('si', [{ h: 'Date', f: x => esc(DB.showDate(x.date)) }, { h: 'Medicine Name', f: x => esc(x.item) }, { h: 'Qty Sold (packs)', n: 1, f: x => qty(DB.r2(x.qtyPacks)) }, { h: 'Batch No', f: x => esc(x.batch) },
      { h: 'Invoice No', f: x => esc(x.invoiceNo) }, { h: 'MRP', n: 1, f: x => N(x.mrp).toFixed(2) }, { h: 'Disc %', n: 1, f: x => qty(x.disc) }, { h: 'Net Amount', n: 1, f: x => money(x.net) }, { h: 'Pack', n: 1, f: x => qty(x.pack) },
      { h: 'Strip', n: 1, f: x => qty(x.strip) }, { h: 'Loose', n: 1, f: x => qty(x.loose) }, { h: 'Free (pcs)', n: 1, f: x => qty(x.freePcs) }, { h: 'Exp', f: x => esc(x.exp) }, { h: 'GST %', n: 1, f: x => qty(x.gst) },
      { h: 'Patient Name', f: x => esc(x.patient) }, { h: 'MFG', f: x => esc(x.mfg) }, { h: 'HSN', f: x => esc(x.hsn) }],
      rows, { click: 1, foot: ['TOTAL', rows.length + ' lines', qty0(sum('qtyPacks')), '', '', '', '', money(sum('net')), '', '', '', '', '', '', '', '', ''], empty: 'No sale lines in this period.' });
  wirePeriod('saleitems');
  wireTable('si', rows, x => viewInvoice(x.invoiceNo), () => PAGES.saleitems(route().q));
};

/* ================= GST (GSTR-1 summaries) =================
   Net Amount in sales INCLUDES GST: taxable = net / (1 + GST%/100), tax = net - taxable (half CGST, half SGST). */
PAGES.gst = q => {
  const d = new Date(), y = d.getFullYear(), m = d.getMonth();
  const from = q.get('from') || DB.toISO(new Date(y, m - 1, 1)), to = q.get('to') || DB.toISO(new Date(y, m, 0)), tab = q.get('t') || 'b2cs';
  setTop('GST', `<button class="btn" id="gXls">⬇ <span class="lbl">Excel</span></button>`);
  const items = DB.D.saleItems.filter(x => x.item && (x.date || '') >= from && (x.date || '') <= to);
  const pos = DB.getSetting('placeOfSupply', '19-West Bengal') || '19-West Bengal';
  const R = n => DB.r2x(n);
  let title = '', cols = [], rows = [], foot = null, msg = '', sheet = [];
  if (tab === 'b2cs') {
    const rates = new Map(); items.forEach(x => { const g = N(x.gst); if (g > 0) rates.set(g, (rates.get(g) || 0) + N(x.net) / (1 + g / 100)); });
    rows = [...rates.entries()].sort((a, b) => a[0] - b[0]).map(([g, v]) => ({ type: 'OE', pos, rate: g, val: R(v) }));
    title = 'Summary For B2CS(7)';
    cols = [{ h: 'Type', f: r => r.type }, { h: 'Place of Supply', f: r => esc(r.pos) }, { h: 'Applicable %Tax', f: () => '' }, { h: 'Rate', n: 1, f: r => r.rate }, { h: 'Taxable Value', n: 1, f: r => r.val.toFixed(2) }, { h: 'Cess Amount', n: 1, f: () => '0.00' }, { h: 'E-Commerce GSTIN', f: () => '' }];
    foot = ['TOTAL', '', '', '', R(rows.reduce((a, r) => a + r.val, 0)).toFixed(2), '0.00', ''];
    sheet = [cols.map(c => c.h), ...rows.map(r => [r.type, r.pos, '', r.rate, r.val, 0, ''])];
  } else if (tab === 'nil') {
    const zero = R(items.filter(x => N(x.gst) === 0).reduce((a, x) => a + N(x.net), 0));
    const asEx = DB.getSetting('zeroGstAs', 'NIL') === 'EXEMPT', nil = asEx ? 0 : zero, ex = asEx ? zero : 0;
    title = 'Summary For Nil rated, exempted and non GST outward supplies (8)';
    rows = [['Inter-State supplies to registered persons', 0, 0, 0], ['Intra-State supplies to registered persons', 0, 0, 0], ['Inter-State supplies to unregistered persons', 0, 0, 0], ['Intra-State supplies to unregistered persons', nil, ex, 0]];
    cols = [{ h: 'Description', f: r => r[0] }, { h: 'Nil Rated Supplies', n: 1, f: r => r[1].toFixed(2) }, { h: 'Exempted (other than nil rated/non GST supply)', n: 1, f: r => r[2].toFixed(2) }, { h: 'Non-GST supplies', n: 1, f: r => r[3].toFixed(2) }];
    foot = ['TOTAL', nil.toFixed(2), ex.toFixed(2), '0.00'];
    sheet = [cols.map(c => c.h), ...rows];
  } else if (tab === 'hsn') {
    const g = new Map(); let noHsn = 0;
    items.forEach(x => {
      let h = String(x.hsn || '').trim(); if (!h) { h = '(no HSN)'; noHsn++; }
      const k = h + '|' + N(x.gst); const a = g.get(k) || { hsn: h, rate: N(x.gst), qty: 0, val: 0, tx: 0 };
      a.qty += N(x.qtyPacks); a.val += N(x.net); a.tx += N(x.net) / (1 + N(x.gst) / 100); g.set(k, a);
    });
    rows = [...g.values()].sort((a, b) => a.hsn.localeCompare(b.hsn) || a.rate - b.rate).map(a => { const val = R(a.val), tx = R(a.tx), half = R((val - tx) / 2); return { ...a, val, tx, c: half, s: half }; });
    title = 'Summary For HSN(12)';
    cols = [{ h: 'HSN', f: r => esc(r.hsn) }, { h: 'Description', f: r => esc(r.hsn) }, { h: 'UQC', f: () => 'OTH-OTHERS' }, { h: 'Total Quantity', n: 1, f: r => qty0(r.qty) }, { h: 'Total Value', n: 1, f: r => r.val.toFixed(2) }, { h: 'Rate', n: 1, f: r => r.rate },
      { h: 'Taxable Value', n: 1, f: r => r.tx.toFixed(2) }, { h: 'Integrated Tax Amount', n: 1, f: () => '0.00' }, { h: 'Central Tax Amount', n: 1, f: r => r.c.toFixed(2) }, { h: 'State/UT Tax Amount', n: 1, f: r => r.s.toFixed(2) }, { h: 'Cess Amount', n: 1, f: () => '0.00' }];
    const t = k => R(rows.reduce((a, r) => a + r[k], 0)).toFixed(2);
    foot = [new Set(rows.map(r => r.hsn)).size + ' HSN', '', '', '', t('val'), '', t('tx'), '0.00', t('c'), t('s'), '0.00'];
    if (noHsn) msg = noHsn + ' sold line(s) have no HSN code, shown as “(no HSN)”. Add HSN in Bill Entry for the next bills.';
    sheet = [cols.map(c => c.h), ...rows.map(r => [r.hsn, r.hsn, 'OTH-OTHERS', DB.r2(r.qty), r.val, r.rate, r.tx, 0, r.c, r.s, 0])];
  } else {
    const ser = new Map(), seen = new Set();
    DB.D.sales.filter(x => (x.date || '') >= from && (x.date || '') <= to).forEach(x => {
      const inv = String(x.invoiceNo || '').trim(), K = invKey(inv); if (!inv || seen.has(K)) return; seen.add(K);
      const m2 = inv.match(/^(.*?)(\d*)$/), pre = m2[1].toUpperCase(), num = m2[2] ? +m2[2] : 0;
      const a = ser.get(pre) || { min: num, max: num, n: 0, minT: inv, maxT: inv };
      if (num < a.min) { a.min = num; a.minT = inv; } if (num > a.max) { a.max = num; a.maxT = inv; } a.n++; ser.set(pre, a);
    });
    rows = [...ser.entries()].sort().map(([, a]) => { const tot = a.max - a.min + 1; return { from: a.minT, to: a.maxT, tot, canc: Math.max(0, tot - a.n) }; });
    title = 'Summary of documents issued during the tax period (13)';
    cols = [{ h: 'Nature of Document', f: () => 'Invoices for outward supply' }, { h: 'Sr.No. From', f: r => esc(r.from) }, { h: 'Sr.No. To', f: r => esc(r.to) }, { h: 'Total Number', n: 1, f: r => r.tot }, { h: 'Cancelled', n: 1, f: r => r.canc }];
    const cn = rows.reduce((a, r) => a + r.canc, 0);
    foot = ['TOTAL', '', '', rows.reduce((a, r) => a + r.tot, 0), cn];
    if (cn) msg = cn + ' invoice number(s) are missing in the series (deleted), counted as Cancelled.';
    sheet = [cols.map(c => c.h), ...rows.map(r => ['Invoices for outward supply', r.from, r.to, r.tot, r.canc])];
  }
  const tabs = [['b2cs', '1. B2CS (7)'], ['nil', '2. Nil / Exempted (8)'], ['hsn', '3. HSN (12)'], ['docs', '4. Documents (13)']];
  $('page').innerHTML = `<div class="card"><div class="fields"><div><label class="l">FROM date</label><input class="f" type="date" id="pF" value="${esc(from)}"></div><div><label class="l">TO date</label><input class="f" type="date" id="pT" value="${esc(to)}"></div></div>
    <div class="chips" style="margin-top:10px"><button class="chip" data-p="0">This month</button><button class="chip" data-p="1">Last month</button></div></div>
    <div class="chips" style="margin-bottom:12px">${tabs.map(([k, t]) => `<button class="chip ${tab === k ? 'on' : ''}" data-t="${k}">${t}</button>`).join('')}</div>
    <div class="note">${esc(title)} · ${DB.showDate(from)} to ${DB.showDate(to)} · ${items.length} sold line(s)${tab === 'docs' ? '' : ''}</div>
    ${msg ? `<div class="note warn">${esc(msg)}</div>` : ''}
    ${table('gs', cols, rows, { foot, empty: tab === 'docs' ? 'No invoices in this period.' : 'No sales found in this period.' })}
    <p class="muted small">Sales prices include GST: taxable value = net ÷ (1 + GST%), tax split half CGST, half SGST. Place of supply: ${esc(pos)} (change in Settings). Check the figures before filing.</p>`;
  const keep = { t: tab };
  $('pF').onchange = $('pT').onchange = () => go('gst', { from: $('pF').value, to: $('pT').value, t: tab });
  document.querySelectorAll('[data-p]').forEach(c => c.onclick = () => { const v = +c.dataset.p; go('gst', { from: DB.toISO(new Date(y, m - v, 1)), to: DB.toISO(v ? new Date(y, m, 0) : d), ...keep }); });
  document.querySelectorAll('[data-t]').forEach(c => c.onclick = () => go('gst', { from, to, t: c.dataset.t }));
  $('gXls').onclick = async () => {
    await Importer.loadXLSX();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[title], ['From ' + DB.showDate(from) + ' to ' + DB.showDate(to)], [], ...sheet]), tab.toUpperCase());
    XLSX.writeFile(wb, 'GST_' + tab + '_' + from + '_' + to + '.xlsx');
  };
};

window.Sales = { saleNet, saleTotals, lineQty, billQty, nextInvoiceNo, rupeesInWords, printInvoice, viewInvoice, DEFAULT_TERMS };
