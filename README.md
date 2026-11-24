# Pharmacy Bill

A modern web app version of the **PHARMACY_BILL.xlsm** workbook. It works on Android and desktop (install it from Chrome or Edge with "Install app" / "Add to Home screen") and reads distributor bills automatically with Google Gemini (free tier).

## Screens (same as the Excel tabs)
| App screen | Excel tab |
|---|---|
| Scan bills (Google Drive folder, upload, camera) | new: automatic bill reading |
| Bill Entry | BILL ENTRY |
| Distributor ledger | one sheet per distributor (THE CAPITAL, S.K. PHARMA, ...) |
| Purchase Details | Purchase Details |
| Party Details | Party Details |
| Stocks (with Adjust stock, Undo, Adjust log) | Stocks, StockAdjust, StockAdjust Undone |
| Medicine Search | Medicine Search |
| Expiry Search (with Save as PDF / Print) | Expiry Search |
| Sales Invoice (batch picker, Pack / Strip / Loose / Free, stock check, Save & print) | Sales Invoice, InvoicePrint |
| Sales (view, print, modify, delete an invoice) | Sales |
| Sale Items (product-wise lines) | SaleItems |
| GST: B2CS (7), Nil rated (8), HSN (12), Documents (13), export to Excel | GST |
| Dashboard | new: sales today, stock value, monthly purchases, expiring stock |

## Calculations (copied from the workbook macros)
- **Net Amount** = Quantity × Rate × (1 − Discount% / 100) × (1 + (CGST% + SGST%) / 100). Free quantity is not charged.
- **Remaining stock** per medicine + batch = Purchased + Free + Add − Reduce − Sold.
- **Stock value** = Remaining × cost per pack, where cost per pack = money paid ÷ all packs received (free packs make it cheaper).
- **Sale line** = (Strip + Loose ÷ Pack) × MRP × (1 − Disc%), rounded to the paisa; free pieces leave stock but are not charged. MRP includes GST, so GST = amount × GST% ÷ (100 + GST%).
- **GST reports**: taxable value = net ÷ (1 + GST%), tax split half CGST, half SGST; place of supply 19-West Bengal (Settings).
- Rounding follows Excel VBA `Round` (halves to even). Checked against the real workbook: all 2,092 batches and all quantity totals match. 3 batches differ by 1 paisa, because Excel keeps extra hidden digits when a value lands exactly on half a paisa.

## How automatic bill reading works
1. The bill photo or PDF (from the Drive folder, an upload or the camera) is sent to Gemini with a fixed list of fields: distributor, GSTIN, invoice no, date, and for every line item, Mfg, HSN, batch, expiry, qty, free, MRP, rate, disc %, CGST %, SGST % and amount.
2. The answer is put into **your** Bill Entry format:
   - The distributor is matched by GSTIN, then by name.
   - Medicine names are matched to the names already in stock ("DOLO-650 TAB" → "DOLO 650 TAB").
   - Empty Mfg, HSN or GST is filled from older bills.
   - Expiry is converted to M/YY.
3. Every line is checked:
   - Your Net Amount formula is compared with the amount printed on the bill.
   - The sum is compared with the bill total.
   - Missing batch, expiry or MRP is highlighted, and so is a rate above MRP.
4. You check it next to the bill image and press **Save bill**. Nothing goes into stock before that.

## Setup
- **Gemini key (free):** Settings → follow the link to aistudio.google.com/app/apikey → paste the key → Test.
- **Google Drive folder:** Settings → step-by-step guide (create a free Google Cloud OAuth Client ID once, then paste the folder link).
- **Your data:** Settings → *Import Excel workbook* → choose PHARMACY_BILL.xlsm. The file is read on the device and never uploaded.

## Licence & app lock
- The app must be **activated** with a signed licence code. The code carries the shop name and an expiry date and is signed with a private key only the provider holds, so a shop cannot forge or extend it. The app holds only the public key.
- After expiry the app becomes **read-only**: viewing and printing still work, but new bills, sales and stock changes are blocked until a **renewal code** is entered (Settings → Renew, or the banner).
- An optional **app password** locks the screen each time the app opens on that device. If forgotten, the provider makes a **password-reset code** (valid a few days, for that one shop) which clears it from the "Forgot password?" link.
- Codes are made with `license-generator.html` (kept privately by the provider, never in this repo). It holds the private signing key — keep a safe backup and never share it.
- Note: a web app's source cannot be hidden; the licence controls *access and the renewal period*, not code secrecy.

## Importing Excel
Importing an Excel workbook now **adds to** the existing data (it does not replace it). Lines already present are skipped by content, so re-importing an overlapping file adds only the new lines. The **same file** (matched by its exact contents) is **skipped** automatically. Settings → Data lists imported files and can "Allow re-import" if you need to load one again.

## Live sync (phone & PC show the same data)
Optional. Connect every device of one shop to a free **Supabase** project and they stay in sync automatically.
- Setup is in **Settings → Live sync**: create a free project, run the shown SQL once, paste the Project URL, anon key and a Shop ID (one per shop, same on all devices).
- Each synced row carries a stable id (`_uid`) and changed-at time (`_up`); changes are pushed on save and pulled every few seconds. **Newest change wins.** Deletes travel as tombstones.
- Medicines combine by **name + batch** in the stock view, so the same item from two devices is one stock line.
- It still works offline and catches up when the connection returns. A sidebar dot shows the status.
- The Shop ID is the key to the data — keep it private. (The open RLS policy means anyone with URL + anon key + Shop ID can read that shop's data; use a long random Shop ID.)

## Where the data is kept
In this browser on each device (IndexedDB), and it works offline. Use **Settings → Download backup** for a file copy, and *Restore backup* to move data to a device without sync.

## Files
`index.html`, `css/app.css`, `js/db.js` (data + calculations), `js/import.js` (Excel import/export), `js/ai.js` (Gemini + Google Drive), `js/app.js` (screens), `sw.js` (offline), `vendor/xlsx.full.min.js` (SheetJS, Apache-2.0).
No build step: host the folder on any static web host (for example GitHub Pages) with HTTPS.
