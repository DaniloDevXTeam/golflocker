function __debugActiveSheetId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/** ================== CONFIG ================== **/
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

const TABS = ['Clubs','Sets','Tassen',"Trolley's",'Overig','Diensten']; // voorraad-tabs

const COL  = {
  sku:        1,
  desc:       2,
  buy:        4,
  party:      5,
  expected:   6,
  sale:       7,
  saleDate:   8,
  expMargin:  9,
  backupExp: 12
};

const CACHE_MIN     = 120;   // winkelmandje 2 uur per gebruiker
const INDEX_TTL_MIN = 30;    // SKU-index 30 min cache

// Waar 80mm-bon PDF's worden opgeslagen
const TICKET80_FOLDER_ID = '1uI70DxaLW_RaYxpC2FIcEdhtIVWG6nO1';

// Branding gedeeld door front-end, PDF en 80mm
const BRAND = {
  name:       'Golf Locker',
  line1:      'Dorpstraat 38',
  line2:      '3981 EB, Bunnik',
  phone:      '06 383 907 22',
  email:      'info@golflocker.nl',
  vat:        'NL861782495B01',
  kvk:        '80742300',
  extra:      'Btw inbegrepen',
  logoUrl:    'https://shop.golf-locker.nl/wp-content/uploads/2024/04/Golf-Locker-Logo-1-scaled.png',
  webshopUrl: 'https://shop.golf-locker.nl'
};

// Log-sheets
const LOG = {
  headSheet: 'Sales',
  lineSheet: 'Sales_Lines'
};


/** ================== HELPER: 80mm TICKET HTML ================== **/

/**
 * Bouwt de HTML voor de 80mm-bon.
 * Wordt gebruikt door:
 *  - _serveTicket_ (live herprint via ?file=ticket&no=..)
 *  - apiBookAndReceipt (voor PDF-archief in Drive)
 */
function _build80mmTicketHtml_(opts) {
  const {
    receiptNo,
    payMethod,
    customerEmail,
    total,
    dateString,
    items
  } = opts;

  const enc = encodeURIComponent;
  const fmt = n => Utilities.formatString(
    "€ %s",
    Number(n || 0).toFixed(2).replace('.', ',')
  );

  const qrUrl = `https://quickchart.io/qr?text=${enc(BRAND.webshopUrl || '')}&size=180&margin=1&format=png`;
  const code128Url = `https://bwipjs-api.metafloor.com/?bcid=code128&text=${enc(receiptNo || '')}&scale=3&height=12&includetext&textxalign=center`;

  // let op: </script> escapen als <\/script> in template literal
  const autoPrintScript = `
    <script>
      window.addEventListener('load', function(){
        setTimeout(function(){
          try { window.print(); } catch(e){}
          setTimeout(function(){ try { window.close(); } catch(e){} }, 800);
        }, 300);
      });
    <\\/script>
  `;

  const rowsHtml = (items || []).map(it => {
    const skuShort = String(it.sku || '').slice(0, 5).padEnd(5, ' ');
    const descSafe = String(it.desc || '').replace(/</g, '&lt;');
    const qty      = Number(it.qty || 1);
    const price    = Number(it.price || 0);
    const subt     = it.subt != null ? Number(it.subt) : (price * qty);
    return `
      <tr>
        <td class="sku mono">${skuShort}</td>
        <td class="qty">${qty}</td>
        <td class="desc">${descSafe}</td>
        <td class="price">${fmt(price)}</td>
        <td class="subt mono">${fmt(subt)}</td>
      </tr>`;
  }).join('');

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Bon ${receiptNo || ''}</title>
  <style>
    @page { size: 80mm auto; margin: 4mm 4mm; }
    * { box-sizing:border-box; }
    body {
      font:11px/1.25 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
      color:#000;
      margin:0;
    }
    .ticket { width:72mm; max-width:72mm; margin:0 auto; }
    .center { text-align:center; }
    .right { text-align:right; }
    .muted { color:#555; }
    .logo { display:block; width:54mm; margin:0 auto 8px; filter:grayscale(100%); }
    h1 { font-size:14px; margin:0 0 6px; text-align:center; }
    .small { font-size:10px; }
    .mono { font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; }
    .row { display:flex; justify-content:space-between; }
    hr { border:0; border-top:1px dashed #000; margin:6px 0; }
    table { width:100%; border-collapse:collapse; }
    th,td { padding:2px 0; vertical-align:top; }
    th { text-align:left; font-weight:700; }
    .sku   { width:12mm; font-size:9px; white-space:pre; }
    .qty   { width:7mm; text-align:left; padding-left:2px; }
    .desc  { width:27mm; padding-right:2mm; }
    .price { width:11mm; text-align:right; }
    .subt  { width:13mm; text-align:right; }
    .total { font-size:13px; font-weight:700; }
    .qr { width:36mm; margin:6px auto 0; }
    .barcode { width:60mm; margin:6px auto 0; display:block; }
    @media print { .noprint { display:none!important; } }
  </style>
</head>
<body>
  <div class="ticket">
    ${BRAND.logoUrl ? `<img class="logo" src="${BRAND.logoUrl}" alt="logo">` : ''}
    <h1>${BRAND.name || ''}</h1>
    <div class="center small muted">${BRAND.line1 || ''} • ${BRAND.line2 || ''}</div>
    <div class="center small muted">${BRAND.phone || ''} • ${BRAND.email || ''}</div>
    <div class="center small muted">BTW: ${BRAND.vat || '-'} • KvK: ${BRAND.kvk || '-'}</div>

    <hr>
    <div class="row small">
      <div>Betaalwijze: ${payMethod || '-'}</div>
      <div class="right">${dateString || ''}</div>
    </div>
    ${customerEmail ? `<div class="small">Klant: ${customerEmail}</div>` : ''}
    <hr>

    <table>
      <thead>
        <tr>
          <th class="sku">SKU</th>
          <th class="qty">Aant.</th>
          <th class="desc">Artikel</th>
          <th class="price">Prijs</th>
          <th class="subt">Subt.</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>

    <hr>
    <table>
      <tr><td class="right total" colspan="5">Totaal: ${fmt(total)}</td></tr>
    </table>

    ${BRAND.extra ? `<div class="small muted" style="margin-top:6px">${BRAND.extra}</div>` : ''}

    <div class="center">
      <img class="qr" src="${qrUrl}" alt="qr">
    </div>
    <div class="center small muted" style="margin-top:4px">
      ${BRAND.webshopUrl || ''}
    </div>

    <img class="barcode" src="${code128Url}" alt="Bon ${receiptNo || ''}">
    <div class="center small muted" style="margin-top:6px">Bedankt voor uw aankoop!<br>Ruilen binnen 30 dagen met deze kassabon.</div>

    <div class="noprint center" style="margin-top:8px">
      <button onclick="window.print()">Print</button>
    </div>
  </div>

  ${autoPrintScript}
</body>
</html>`;
}


/** ================== WEB ENTRY (PWA + TICKET) ================== **/

function doGet(e) {
  const p = (e && e.parameter) || {};
  const f = p.file || '';

  if (f === 'manifest') return _serveManifest_();
  if (f === 'sw')       return _serveServiceWorker_();
  if (f === 'ticket')   return _serveTicket_(e);

  // 🔥 NIEUWE ROUTE → retourbon
  if (f === 'returnticket') {
    const no = p.no || '';
    return HtmlService.createHtmlOutput(buildReturnTicket80mm(no))
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default → laad POS UI
  return HtmlService.createHtmlOutputFromFile('app')
    .setTitle('Golf Locker POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function _serveManifest_() {
  const manifest = {
    name: "Golf Locker POS",
    short_name: "GL POS",
    description: "Snel afrekenen, bon en e-mail — ook als PWA.",
    start_url: "./",
    scope: "./",
    display: "standalone",
    background_color: "#ffffff",
    theme_color: "#297900",
    icons: [
      {
        src: BRAND.logoUrl,
        sizes: "192x192",
        type: "image/png",
        purpose: "any"
      },
      {
        src: BRAND.logoUrl,
        sizes: "512x512",
        type: "image/png",
        purpose: "any"
      }
    ]
  };

  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

function _serveServiceWorker_() {
  const sw = [
    "const NAME = 'gl-pos-v3';",
    "const STATIC = [",
    "  self.registration.scope,",
    "  self.registration.scope + '?file=manifest',",
    "  'https://unpkg.com/@zxing/library@0.20.0',",
    `  '${BRAND.logoUrl}'`,
    "];",
    "",
    "self.addEventListener('install', e => {",
    "  e.waitUntil(",
    "    caches.open(NAME)",
    "      .then(c => c.addAll(STATIC))",
    "      .then(() => self.skipWaiting())",
    "  );",
    "});",
    "",
    "self.addEventListener('activate', e => {",
    "  e.waitUntil(",
    "    caches.keys().then(keys =>",
    "      Promise.all(keys.map(k => k === NAME ? null : caches.delete(k)))",
    "    )",
    "  );",
    "});",
    "",
    "self.addEventListener('fetch', e => {",
    "  const req = e.request;",
    "  const accept = req.headers.get('accept') || '';",
    "  const isHTML = accept.includes('text/html');",
    "  if (isHTML) {",
    "    e.respondWith(",
    "      fetch(req).then(res => {",
    "        const copy = res.clone();",
    "        caches.open(NAME).then(c => c.put(req, copy));",
    "        return res;",
    "      }).catch(() => caches.match(req))",
    "    );",
    "  } else {",
    "    e.respondWith(",
    "      caches.match(req).then(hit => hit ||",
    "        fetch(req).then(res => {",
    "          const copy = res.clone();",
    "          caches.open(NAME).then(c => c.put(req, copy));",
    "          return res;",
    "        })",
    "      )",
    "    );",
    "  }",
    "});"
  ].join('\n');

  return ContentService.createTextOutput(sw)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * HTML endpoint voor 80mm-ticket (?file=ticket&no=...)
 * Leest uit Sales + Sales_Lines en gebruikt dezelfde layout-helper.
 */
function _serveTicket_(e) {
  const no = (e && e.parameter && e.parameter.no) || '';
  if (!no) {
    return HtmlService.createHtmlOutput('Bonnummer ontbreekt');
  }

  const ss    = SpreadsheetApp.getActive();
  const head  = ss.getSheetByName(LOG.headSheet);
  const lines = ss.getSheetByName(LOG.lineSheet);

  if (!head || !lines) {
    return HtmlService.createHtmlOutput('Log-sheets niet gevonden');
  }

  // Kopregel zoeken
  const lastHead = head.getLastRow();
  const headVals = lastHead > 1
    ? head.getRange(2, 1, lastHead - 1, head.getLastColumn()).getValues()
    : [];

  const hRow = headVals.find(r => String(r[0]) === String(no));
  if (!hRow) {
    return HtmlService.createHtmlOutput('Bon niet gevonden: ' + no);
  }

  const receiptNo  = String(hRow[0]);
  const datum      = hRow[1];
  const payMethod  = String(hRow[2] || '');
  const total      = Number(hRow[3] || 0);
  const custEmail  = String(hRow[4] || '');

  const tz = Session.getScriptTimeZone() || 'Europe/Amsterdam';
  const dt = datum instanceof Date
    ? Utilities.formatDate(datum, tz, 'dd-MM-yyyy HH:mm')
    : String(datum || '');

  // Regels ophalen
  const lastLine = lines.getLastRow();
  const lineVals = lastLine > 1
    ? lines.getRange(2, 1, lastLine - 1, 7).getValues()
    : [];

  const items = lineVals
    .filter(r => String(r[0]) === receiptNo)
    .map(r => ({
      sku:   String(r[1] || ''),
      desc:  String(r[2] || ''),
      price: Number(r[3] || 0),
      qty:   Number(r[4] || 1),
      subt:  Number(r[5] || 0)
    }));

  const html = _build80mmTicketHtml_({
    receiptNo,
    payMethod,
    customerEmail: custEmail,
    total,
    dateString: dt,
    items
  });

  return HtmlService.createHtmlOutput(html)
    .setTitle('Ticket ' + receiptNo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/** ================== KLEINE PING ================== **/

function apiPing() {
  return 'ok';
}


/** ================== CART STORAGE ================== **/

function _cartKey_() {
  const email = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'anon';
  return email + '::' + SpreadsheetApp.getActive().getId();
}

function getCart() {
  const cache = CacheService.getUserCache();
  const raw = cache.get(_cartKey_());
  return raw ? JSON.parse(raw) : [];
}

function saveCart(cart) {
  CacheService.getUserCache().put(_cartKey_(), JSON.stringify(cart), CACHE_MIN);
  return cart;
}

function clearCart() {
  CacheService.getUserCache().remove(_cartKey_());
  return [];
}

function apiGetCart()       { return getCart(); }
function apiClearCart()     { clearCart(); return true; }


/** ================== NIEUWE SUPER-SAFE SKU INDEX ================== **/

/**
 * Maak per sheet een aparte index.
 * Cache key per sheet, nooit te groot.
 */
function _indexKeyForSheet_(sheetName) {
  const ssId = SpreadsheetApp.getActive().getId();
  return `SKU_INDEX_V2::${ssId}::${sheetName}`;
}

/** Ophalen index voor een sheet */
function _getIndexForSheet_(sheetName) {
  const raw = CacheService.getScriptCache().get(_indexKeyForSheet_(sheetName));
  return raw ? JSON.parse(raw) : null;
}

/** Opslaan index voor een sheet (nooit te groot) */
function _saveIndexForSheet_(sheetName, data) {
  const key = _indexKeyForSheet_(sheetName);
  CacheService.getScriptCache().put(key, JSON.stringify(data), INDEX_TTL_MIN * 60);
}

/** Invalideer ALLE indexen voor ALLE voorraadtabs */
function invalidateSkuIndex_() {
  const ssId = SpreadsheetApp.getActive().getId();
  TABS.forEach(name => {
    CacheService.getScriptCache().remove(`SKU_INDEX_V2::${ssId}::${name}`);
  });
}

/**
 * Bouw index per sheet:
 * {
 *   "1513": [ { row: 12, free:true }, { row:18, free:false } ]
 * }
 */
function _buildIndexForSheet_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return null;

  const last = sh.getLastRow();
  if (last <= 1) {
    _saveIndexForSheet_(sheetName, {});
    return {};
  }

  const data = sh.getRange(2, 1, last - 1, 7).getValues();
  const index = {};

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const sku = String(r[0] || '').trim();
    if (!sku) continue;

    const saleVal = r[COL.sale - 1];
    const free = (saleVal === '' || saleVal === null);

    if (!index[sku]) index[sku] = [];
    index[sku].push({
      row: i + 2,
      free
    });
  }

  _saveIndexForSheet_(sheetName, index);
  return index;
}

/** Haal index voor sheet op, bouw indien nodig */
function _ensureIndexForSheet_(sheetName) {
  let idx = _getIndexForSheet_(sheetName);
  if (!idx) idx = _buildIndexForSheet_(sheetName);
  return idx;
}

/**
 * NIEUWE snelle findBySku:
 * ✓ zoekt per sheet
 * ✓ first-free
 * ✓ nooit te groot
 */
function findBySku(sku, opts) {
  opts = opts || {};
  const key = String(sku).trim();
  const ss = SpreadsheetApp.getActive();

  for (let t = 0; t < TABS.length; t++) {
    const name = TABS[t];
    const index = _ensureIndexForSheet_(name);
    const list = index[key];
    if (!list || !list.length) continue;

    // free-first
    const freeOne = list.find(x => x.free);
    const chosen = freeOne || list[0];

    // detaildata ophalen
    const sh = ss.getSheetByName(name);
    const row = chosen.row;
    const rowVals = sh.getRange(row, 1, 1, 7).getValues()[0];

    return {
      item: {
        sheetName: name,
        row,
        desc: rowVals[1],
        expected: Number(rowVals[5] || 0),
        party: rowVals[4]
      },
      positions: list
    };
  }

  return null; // Niet gevonden
}

/** Warm de hele index op */
function apiWarmIndex() {
  TABS.forEach(name => _buildIndexForSheet_(name));
  return true;
}


/** ================== CART API ================== **/

function apiAddFast(sku) {
  const rawSku = String(sku).trim();


// GENERATOR-SKU LOGICA
// ===========================
 if (GENERATOR_SKUS && GENERATOR_SKUS[rawSku]) {
    const generated = createGeneratedItem(rawSku);

    const cart = getCart();
    const line = {
      sku:       generated.sku,
      sheetName: generated.sheetName,
      row:       generated.row,
      desc:      generated.description,
      price:     generated.expected,  // jij past deze aan in de POS
      qty:       1,
      party:     "",
      edited:    false
    };

    cart.push(line);
    saveCart(cart);

    const total = cart.reduce((s, it) =>
      s + (Number(it.price) || 0) * (Number(it.qty) || 1), 0
    );

    return { line, total };
  }
// ===========================


  const res = findBySku(sku);
  if (!res) throw new Error('SKU niet gevonden: ' + sku);
  const { item } = res;

  const cart = getCart();
  const i = cart.findIndex(x => String(x.sku) === String(sku));
  let line;

  if (i >= 0) {
    cart[i].qty += 1;
    line = cart[i];
  } else {
    line = {
      sku:   String(sku),
      sheetName: item.sheetName,
      row:       item.row,
      desc:      item.desc,
      price:     item.expected,
      qty:       1,
      party:     item.party,
      edited:    false
    };
    cart.push(line);
  }

  saveCart(cart);
  const total = cart.reduce((s, it) =>
    s + (Number(it.price) || 0) * (Number(it.qty) || 1), 0);

  return { line, total };
}

function apiRemoveFast(sku) {
  const cart = getCart();
  const i = cart.findIndex(x => String(x.sku) === String(sku));
  if (i >= 0) cart.splice(i, 1);
  saveCart(cart);

  const total = cart.reduce((s, it) =>
    s + (Number(it.price) || 0) * (Number(it.qty) || 1), 0);

  return { ok: true, total };
}

function apiSetQtyFast(sku, qty) {
  const cart = getCart();
  const it = cart.find(x => String(x.sku) === String(sku));
  if (!it) throw new Error('Niet in mandje');

  it.qty = Math.max(1, Number(qty) || 1);
  saveCart(cart);

  const total = cart.reduce((s, i) =>
    s + (Number(i.price) || 0) * (Number(i.qty) || 1), 0);

  return { ok: true, total, line: it };
}

function apiSetPriceFast(sku, price) {
  const cart = getCart();
  const it = cart.find(x => String(x.sku) === String(sku));
  if (!it) throw new Error('Niet in mandje');

  it.price  = Math.max(0, Number(String(price).replace(',', '.')) || 0);
  it.edited = true;
  saveCart(cart);

  const total = cart.reduce((s, i) =>
    s + (Number(i.price) || 0) * (Number(i.qty) || 1), 0);

  return { ok: true, total, line: it };
}


/** ================== REFRESH VANUIT SHEET ================== **/

function apiRefreshFromSheet() {
  const ss   = SpreadsheetApp.getActive();
  const cart = getCart();
  if (!cart.length) return { cart, total: 0, changed: 0 };

  let changed = 0;
  const bySheet = {};

  cart.forEach(it => {
    if (!bySheet[it.sheetName]) bySheet[it.sheetName] = [];
    bySheet[it.sheetName].push(it);
  });

  Object.keys(bySheet).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const items = bySheet[name];
    items.forEach(it => {
      if (it.edited) return; // handmatige prijs respecteren

      const desc     = sh.getRange(it.row, COL.desc).getValue();
      const expected = Number(sh.getRange(it.row, COL.expected).getValue()) || 0;

      if (it.desc !== desc || it.price !== expected) {
        it.desc  = desc;
        it.price = expected;
        changed++;
      }
    });
  });

  saveCart(cart);

  const total = cart.reduce((s, i) =>
    s + (Number(i.price) || 0) * (Number(i.qty) || 1), 0);

  return { cart, total, changed };
}


/** ================== MINI DATABASE (Sales / Sales_Lines) ================== **/

function ensureLogSheets_() {
  const ss = SpreadsheetApp.getActive();

  // HEAD-sheet
  let head = ss.getSheetByName(LOG.headSheet);
  if (!head) {
    head = ss.insertSheet(LOG.headSheet);
    head.getRange(1, 1, 1, 8).setValues([[
      'receipt_no', 'datum', 'betaalwijze', 'totaal',
      'klant_email', 'pdfUrl', 'mail_status', 'bon80Url'
    ]]);
    head.setFrozenRows(1);
  } else {
    const lastCol    = head.getLastColumn();
    const headerVals = head.getRange(1, 1, 1, lastCol).getValues()[0];

    // zorg dat pdfUrl, mail_status, bon80Url bestaan (zonder data te slopen)
    const needed = ['pdfUrl', 'mail_status', 'bon80Url'];
    needed.forEach(name => {
      if (!headerVals.includes(name)) {
        const newCol = head.getLastColumn() + 1;
        head.getRange(1, newCol).setValue(name);
      }
    });
  }

  // LINES-sheet
  let lines = ss.getSheetByName(LOG.lineSheet);
  if (!lines) {
    lines = ss.insertSheet(LOG.lineSheet);
    lines.getRange(1, 1, 1, 7).setValues([[
      'receipt_no','sku','omschrijving','prijs','aantal','subtotaal','partij'
    ]]);
    lines.setFrozenRows(1);
  }
}

function nextReceiptNo_() {
  const props = PropertiesService.getDocumentProperties();
  const d  = new Date();
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  const key  = `RCPT_SEQ_${yyyy}${mm}${dd}`;
  const cur  = Number(props.getProperty(key) || '0') + 1;
  props.setProperty(key, String(cur));
  const seq  = String(cur).padStart(3, '0');
  return `GL-${yyyy}${mm}${dd}-${seq}`;
}


/** ================== BON-STYLING SHEET 'Bon' ================== **/

function styleBon_(bon, receiptNo, payMethod, customerEmail, total, now) {
  try {
    bon.clearFormats();
    bon.setHiddenGridlines(true);
    
    // Safe column widths
    const cols = bon.getMaxColumns();
    for (let c = 1; c <= Math.min(cols, 5); c++) {
      try { bon.setColumnWidth(c, 60); } catch(e){}
    }

    // Safe row heights
    const rows = bon.getMaxRows();
    for (let r = 1; r <= Math.min(rows, 10); r++) {
      try { bon.setRowHeight(r, 20); } catch(e){}
    }

    // Basic layout only – NO risky formatting
    bon.getRange("A1").setValue("Golf Locker Bon");
    bon.getRange("A2").setValue("Bonnummer: " + receiptNo);
    bon.getRange("A3").setValue("Datum: " + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm'));
    bon.getRange("A4").setValue("Betaalwijze: " + payMethod);
    bon.getRange("A5").setValue("Klant: " + (customerEmail || "-"));
    
    // Table header
    bon.getRange("A7:E7").setValues([["SKU","Omschrijving","Prijs","Aantal","Subtotaal"]]);
    bon.getRange("A7:E7").setFontWeight("bold");

  } catch(err) {
    Logger.log("⚠️ styleBon_ fout: " + err);
  }
}


/** ================== BOOK + RECEIPT (batch, locked) ================== **/

function apiBookAndReceipt(payMethod, customerEmail) {
  const lock = LockService.getDocumentLock();
  lock.tryLock(5000);

  try {
    ensureLogSheets_();

    const ss   = SpreadsheetApp.getActive();
    const cart = getCart();
    if (!cart.length) throw new Error('Mandje is leeg');

    const now   = new Date();
    const tz    = Session.getScriptTimeZone() || 'Europe/Amsterdam';
    const email = (customerEmail || '').trim();
    const receiptNo = nextReceiptNo_();

    let total = 0;

    // groepeer per sheet
    const bySheet = {};
    cart.forEach(it => {
      if (!bySheet[it.sheetName]) bySheet[it.sheetName] = [];
      bySheet[it.sheetName].push(it);
    });

    // schrijf sales in voorraad-tabbladen
    Object.keys(bySheet).forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error('Tab niet gevonden: ' + name);

      const items = bySheet[name];
      const last  = sh.getLastRow();
      if (last <= 1) throw new Error('Geen data in ' + name);

      const skuCol  = sh.getRange(2, COL.sku,  last - 1, 1).getValues().map(r => String(r[0]).trim());
      const saleCol = sh.getRange(2, COL.sale, last - 1, 1).getValues().map(r => r[0]);

      items.forEach(it => {
        let need = it.qty;
        const rowsToWrite = [];

        for (let i = 0; i < skuCol.length && need > 0; i++) {
          if (skuCol[i] === String(it.sku).trim() && (saleCol[i] === '' || saleCol[i] === null)) {
            rowsToWrite.push(i + 2);
            need--;
          }
        }

        if (need > 0) {
          throw new Error(`Niet genoeg vrije regels voor SKU ${it.sku} in ${name} (ontbreken: ${need})`);
        }

        rowsToWrite.forEach(r => {
          sh.getRange(r, COL.sale).setValue(it.price);
          sh.getRange(r, COL.saleDate).setValue(now).setNumberFormat('dd-mm-yyyy');
          sh.getRange(r, COL.expected).setValue(0);
          sh.getRange(r, COL.expMargin).setFormulaR1C1('=IF(ISBLANK(RC6);0;RC6-RC4)');
          total += Number(it.price) || 0;
          saleCol[r - 2] = it.price;
        });
      });
    });

    // Bon-sheet opbouwen
    const bon = ss.getSheetByName('Bon') || ss.insertSheet('Bon');
    bon.clear();

    bon.getRange('A9:E9').setValues([[
      'SKU','Omschrijving','Prijs','Aantal','Subtotaal'
    ]]).setFontWeight('bold');

    const rows = cart.map(it => [
      it.sku,
      it.desc || '',
      Number(it.price) || 0,
      Number(it.qty) || 1,
      (Number(it.price) || 0) * (Number(it.qty) || 1)
    ]);

    if (rows.length) {
      bon.getRange(10, 1, rows.length, 5).setValues(rows);
    }

    const totRow = 10 + rows.length;
    bon.getRange(totRow, 4)
      .setValue('Totaal:')
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    bon.getRange(totRow, 5)
      .setValue(total)
      .setNumberFormat('€ #,##0.00')
      .setFontWeight('bold');

    styleBon_(bon, receiptNo, payMethod, email, total, now);
    SpreadsheetApp.flush();

    // PDF van Bon-sheet
    const pdfUrl = Utilities.formatString(
      'https://docs.google.com/spreadsheets/d/%s/export?format=pdf&gid=%s&portrait=true&size=A5&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false',
      ss.getId(),
      bon.getSheetId()
    );

    // e-mail met PDF (optioneel)
    let mailStatus = 'no email';
    const looksLikeEmail = /\S+@\S+\.\S+/.test(email);

    if (email) {
      if (looksLikeEmail) {
        try {
          const token = ScriptApp.getOAuthToken();
          const pdfBlob = UrlFetchApp.fetch(pdfUrl, {
            headers: { Authorization: 'Bearer ' + token }
          }).getBlob().setName(`GolfLocker-Bon-${receiptNo}.pdf`);

          const subject = `Bon ${receiptNo} – Golf Locker`;
          const body =
            `Bedankt voor je aankoop bij Golf Locker!\n\n` +
            `Bijgevoegd vind je de bon (${receiptNo}).\n` +
            `Totaal: € ${Number(total).toFixed(2)}\n` +
            `Betaalwijze: ${payMethod || '-'}\n\n` +
            `Met sportieve groet,\nGolf Locker`;

          MailApp.sendEmail({
            to: email,
            subject,
            body,
            name: 'Golf Locker',
            attachments: [pdfBlob]
          });

          mailStatus = 'sent';
        } catch (err) {
          mailStatus = 'error: ' + (err && err.message ? err.message : String(err));
        }
      } else {
        mailStatus = 'invalid email';
      }
    }

    // URL voor live 80mm-ticket (?file=ticket&no=...)
    const baseUrl   = ScriptApp.getService().getUrl();
    const ticketUrl = baseUrl + '?file=ticket&no=' + encodeURIComponent(receiptNo);

    // Loggen in Sales + Sales_Lines
    const head = ss.getSheetByName(LOG.headSheet);
    head.appendRow([
      receiptNo,
      now,
      String(payMethod || ''),
      total,
      String(email || ''),
      pdfUrl,
      mailStatus,
      ''          // placeholder bon80Url (kolom H)
    ]);

    const headRowIndex = head.getLastRow(); // net toegevoegde regel

    const linesSheet = ss.getSheetByName(LOG.lineSheet);
    const lineRows = cart.map(it => [
      receiptNo,
      it.sku,
      it.desc || '',
      Number(it.price) || 0,
      Number(it.qty) || 1,
      (Number(it.price) || 0) * (Number(it.qty) || 1),
      it.party || ''
    ]);
    if (lineRows.length) {
      linesSheet
        .getRange(linesSheet.getLastRow() + 1, 1, lineRows.length, 7)
        .setValues(lineRows);
    }

    // 80mm-PDF genereren en opslaan in Drive
    try {
      const dateString = Utilities.formatDate(now, tz, 'dd-MM-yyyy HH:mm');
      const itemsFor80 = cart.map(it => ({
        sku:   it.sku,
        desc:  it.desc || '',
        price: Number(it.price) || 0,
        qty:   Number(it.qty) || 1,
        subt:  (Number(it.price) || 0) * (Number(it.qty) || 1)
      }));

      const ticketHtml = _build80mmTicketHtml_({
        receiptNo,
        payMethod,
        customerEmail: email,
        total,
        dateString,
        items: itemsFor80
      });

      const htmlBlob = Utilities.newBlob(ticketHtml, 'text/html', `Bon80-${receiptNo}.html`);
      const pdf80    = htmlBlob.getAs('application/pdf').setName(`Bon80-${receiptNo}.pdf`);

      const folder   = DriveApp.getFolderById(TICKET80_FOLDER_ID);
      const file80   = folder.createFile(pdf80);
      const bon80Url = file80.getUrl();

      // schrijf URL in kolom H (bon80Url)
      head.getRange(headRowIndex, 8).setValue(bon80Url);
    } catch (err80) {
      // alleen loggen, geen hard error
      Logger.log('Fout bij 80mm-bon PDF: ' + err80);
    }

    clearCart();
    invalidateSkuIndex_();

    // front-end gebruikt ticketUrl voor directe (live) 80mm-print
    return { pdfUrl, total, receiptNo, ticketUrl };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/** ================== BONNEN-OVERZICHT API ================== **/


/**
 * Data voor 80mm-herprint vanuit Bonnen-tab.
 * Front-end kan hiermee via buildReceipt80mmHtml opnieuw een venster openen.
 */
function apiGetReceiptForPrint(receiptNo) {
  ensureLogSheets_();

  const ss    = SpreadsheetApp.getActive();
  const head  = ss.getSheetByName(LOG.headSheet);
  const lines = ss.getSheetByName(LOG.lineSheet);
  if (!head || !lines) {
    throw new Error('Sales sheets ontbreken');
  }

  const lastHead = head.getLastRow();
  const hVals = lastHead > 1
    ? head.getRange(2, 1, lastHead - 1, Math.max(8, head.getLastColumn())).getValues()
    : [];

  const h = hVals.find(r => String(r[0]) === String(receiptNo));
  if (!h) {
    throw new Error('Bon niet gevonden');
  }

  const headObj = {
    receipt_no: String(h[0]),
    date:       h[1],
    pay:        String(h[2] || ''),
    total:      Number(h[3] || 0),
    email:      String(h[4] || ''),
    pdfUrl:     String(h[5] || ''),
    mail:       String(h[6] || ''),
    bon80Url:   String(h[7] || '')
  };

  const lastLine = lines.getLastRow();
  const lVals = lastLine > 1
    ? lines.getRange(2, 1, lastLine - 1, 7).getValues()
    : [];

  const items = lVals
    .filter(r => String(r[0]) === String(receiptNo))
    .map(r => ({
      sku:      String(r[1]),
      desc:     String(r[2] || ''),
      price:    Number(r[3] || 0),
      qty:      Number(r[4] || 1),
      subtotal: Number(r[5] || 0),
      party:    String(r[6] || '')
    }));

  return { head: headObj, items };
}


function apiListReceipts(limit, query) {
  try {
    // Defaults als er niks wordt meegegeven
    limit = Number(limit) || 200;
    query = (query || '').toString().trim();
    Logger.log('apiListReceipts START — limit=%s, query="%s"', limit, query);

    ensureLogSheets_();

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(LOG.headSheet); // 'Sales'
    if (!sh) {
      Logger.log('apiListReceipts: ❌ Sales-sheet niet gevonden');
      return [];
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    Logger.log('apiListReceipts: lastRow=%s, lastCol=%s', lastRow, lastCol);

    if (lastRow <= 1) {
      Logger.log('apiListReceipts: geen data onder de header (alleen rij 1)');
      return [];
    }

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log('apiListReceipts: header=%s', JSON.stringify(header));

    const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Header → index mapping
    const idx = {};
    header.forEach((h, i) => { idx[h] = i; });
    Logger.log('apiListReceipts: idx=%s', JSON.stringify(idx));

    // Als één van de kolommen niet bestaat, meteen loggen
    const required = [
      'receipt_no',
      'datum',
      'betaalwijze',
      'totaal',
      'klant_email',
      'pdfUrl',
      'mail_status',
      'bon80Url'
    ];
    required.forEach(k => {
      if (!(k in idx)) {
        Logger.log('apiListReceipts: ⚠️ kolom "%s" niet gevonden in header', k);
      }
    });

    // Rijen mappen naar objecten
    let rows = vals.map(r => {
      return {
        receipt_no: r[idx['receipt_no']] || '',
        date:       r[idx['datum']] || '',
        pay:        r[idx['betaalwijze']] || '',
        total:      Number(r[idx['totaal']] || 0),
        email:      r[idx['klant_email']] || '',
        pdfUrl:     r[idx['pdfUrl']] || '',
        mail:       r[idx['mail_status']] || '',
        bon80Url:   r[idx['bon80Url']] || ''
      };
    });

    Logger.log('apiListReceipts: ruwe mapped rows (eerste 3)=%s',
      JSON.stringify(rows.slice(0, 3))
    );

    // Filter op zoekterm (bonnummer of e-mail)
    const q = query.toLowerCase();
    if (q) {
      rows = rows.filter(r =>
        (r.receipt_no && String(r.receipt_no).toLowerCase().includes(q)) ||
        (r.email      && String(r.email).toLowerCase().includes(q))
      );
      Logger.log('apiListReceipts: na filter, rows=%s', rows.length);
    }

    // Nieuwste bovenaan op datum
    rows.sort((a, b) => {
      const da = new Date(a.date || 0).getTime();
      const db = new Date(b.date || 0).getTime();
      return db - da;
    });

    if (limit > 0 && rows.length > limit) {
      rows = rows.slice(0, limit);
    }

    Logger.log('apiListReceipts: RETURN rows.length=%s', rows.length);

    // Hard serialiseerbaar maken (voor de zekerheid)
    const safeRows = JSON.parse(JSON.stringify(rows));
    return safeRows;

  } catch (e) {
    Logger.log('apiListReceipts ERROR: ' + (e && e.message ? e.message : String(e)));
    throw e; // zodat withFailureHandler wordt getriggerd in de front-end
  }
}

function debugListReceipts() {
  const ss   = SpreadsheetApp.getActive();
  const sh   = ss.getSheetByName('Sales');
  if (!sh) {
    Logger.log('❌ Sheet "Sales" bestaat niet.');
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  Logger.log('👉 LastRow: ' + lastRow + ', LastCol: ' + lastCol);

  const header = sh.getRange(1,1,1,lastCol).getValues()[0];
  Logger.log('👉 Header: ' + JSON.stringify(header));

  if (lastRow <= 1) {
    Logger.log('❌ Geen data onder de headers');
    return;
  }

  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  Logger.log('👉 Eerste 3 rijen:');
  Logger.log(JSON.stringify(rows.slice(0,3)));

  // Probeer mapping zoals apiListReceipts het doet
  const idx = {};
  header.forEach((h, i) => idx[h] = i);
  Logger.log('👉 Index mapping: ' + JSON.stringify(idx));

  const mapped = rows.map(r => ({
    receipt_no: r[idx['receipt_no']],
    datum:      r[idx['datum']],
    betaalwijze:r[idx['betaalwijze']],
    totaal:     r[idx['totaal']],
    klant_email:r[idx['klant_email']],
    pdfUrl:     r[idx['pdfUrl']],
    mail_status:r[idx['mail_status']],
    bon80Url:   r[idx['bon80Url']]
  }));

  Logger.log('👉 Eerste 3 mapped: ' + JSON.stringify(mapped.slice(0,3)));
}
