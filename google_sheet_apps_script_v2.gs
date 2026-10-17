/*
QUOTATION & INVOICE MAKER — GOOGLE APPS SCRIPT v4
---------------------------------------------------
Google Sheets as the single data source for:
  - Company Details (+ per-company numbering formats)
  - Invoices
  - Quotations
  - Item Catalog (with saved rates)
  - Document Counters (auto-numbering high-water marks)

Setup:
1) Create a Google Sheet.
2) Extensions > Apps Script open karein.
3) Is complete code ko paste karein.
4) Run `setupAdmin` ONCE from the editor (select it in the function dropdown > Run).
   Open View > Logs / Execution log: the generated admin password is printed
   there exactly once. Run it again any time to rotate the password (this also
   signs out every active admin session).
5) Deploy > New deployment > Web app
   Execute as: Me
   Who has access: Anyone
6) /exec URL ko index.html ke Settings (aur admin.html) me paste karein.

Optional: restrict the admin to certain companies by setting the Script
Property ADMIN_COMPANIES (Project Settings > Script properties) to a
pipe-separated list, e.g.  "Acme Traders|Acme Infra". Default "*" = all.

Admin credentials are NEVER stored in the repo / frontend. Only a salted,
iterated HMAC-SHA256 hash lives in Script Properties.

Sheets automatically create hongi:
  - CompanyDetails, Invoices, Quotations, ItemCatalog, DocCounters
*/

const SS = SpreadsheetApp.getActiveSpreadsheet();

const ADMIN_USERNAME = 'admin@email.com';
const SESSION_TTL_SEC = 6 * 60 * 60;          // CacheService max is 21600
const MAX_FAILED_LOGINS = 5;
const LOCKOUT_SEC = 15 * 60;
const HASH_ROUNDS = 2000;

const DEFAULT_FORMATS = {
  invoice: 'INV-{YYYY}-{SEQ:3}',
  quotation: 'QT-{YYYY}-{SEQ:3}'
};
const PAYMENT_STATUSES = ['Pending', 'Partial', 'Paid'];

// Document sheet columns (0-based). Positions 0..17 are the legacy layout and
// MUST NOT move, so existing records stay valid without a migration.
const COL = {
  docNo: 0, date: 1, company: 2, phone: 3, email: 4, address: 5, gst: 6,
  client: 7, clientMobile: 8, validity: 9, tax: 10, profit: 11, mode: 12,
  simpleMode: 13, notes: 14, total: 15, items: 16, updatedAt: 17,
  internalNote: 18, paymentStatus: 19
};
const DOC_HEADERS = [
  'DocNo','Date','Company','Phone','Email','Address','GSTIN','Client','ClientMobile',
  'Validity','Tax','DefaultProfit','Mode','SimpleMode','Notes','Total','ItemsJSON',
  'UpdatedAt','InternalNote','PaymentStatus'
];
const DOC_COLS = DOC_HEADERS.length;
const COMPANY_HEADERS = ['CompanyName','Phone','Email','Address','GSTIN','Currency','Template','UpdatedAt','InvoiceFormat','QuotationFormat'];
const COUNTER_HEADERS = ['DocType','Year','LastNumber','Pool'];

// ─── Setup ───

function setup_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('setup_v4')) return;

  ensureSheet_('CompanyDetails', COMPANY_HEADERS, false);
  ensureSheet_('Invoices', DOC_HEADERS, true);
  ensureSheet_('Quotations', DOC_HEADERS, true);
  ensureSheet_('ItemCatalog', ['ItemName','CostRate','DefaultProfit','LastRate','UpdatedAt'], false);
  ensureSheet_('DocCounters', COUNTER_HEADERS, false);

  cache.put('setup_v4', '1', 600);
}

// Creates the sheet if missing; otherwise only rewrites the header row (data untouched).
function ensureSheet_(name, headers, isDoc) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.appendRow(headers);
  } else {
    // Older versions labelled column 16 "InternalNote" although it held Total.
    // Rewriting the labels fixes that without touching any data row.
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (isDoc) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); // keep DocNo as text
  return sh;
}

function doGet(e) {
  setup_();
  return json_({ok:true,message:'Quotation & Invoice Maker API v4 is running'});
}

function doPost(e) {
  try {
    setup_();
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const sess = getSession_(data.token);
    switch(data.action) {
      case 'ping':
        return json_({ok:true,message:'Connected',admin:!!sess});
      case 'getItems':
        return json_({ok:true,items:getItems_()});
      case 'saveDocument':
        return json_(saveDocument_(data, sess));
      case 'fetchDocuments':
        return json_(fetchDocuments_(data.docNo||'', data.name||'', data.mobile||'', data.docType||'', sess));
      case 'getCompany':
        return json_(getCompany_(data.name));
      case 'getCompanies':
        return json_(getCompanies_());
      case 'saveCompany':
        return json_(saveCompany_(data.company));
      case 'getClients':
        return json_(getClients_());
      case 'getNextNumber':
        return json_(getNextNumber_(data.docType||'quotation', data.company||''));

      // ── Admin ──
      case 'adminLogin':
        return json_(adminLogin_(data.username, data.password));
      case 'adminLogout':
        return json_(adminLogout_(data.token));
      case 'adminMe':
        return json_(sess ? {ok:true, username:sess.u, companies:sess.c} : {ok:false, error:'Not signed in', authRequired:true});
      case 'adminList':
        return json_(requireAdmin_(sess) || adminList_(data, sess));
      case 'adminGet':
        return json_(requireAdmin_(sess) || adminGet_(data, sess));
      case 'adminUpdate':
        return json_(requireAdmin_(sess) || adminUpdate_(data, sess));
      default:
        return json_({ok:false,error:'Unknown action: '+data.action});
    }
  } catch(err) {
    return json_({ok:false,error:String(err && err.message || err)});
  }
}

// ═══════════════════════════════════════════════════════════════════
//  Admin authentication
// ═══════════════════════════════════════════════════════════════════

// Run manually from the Apps Script editor. Prints the new password once.
function setupAdmin() {
  const password = generatePassword_(22);
  const salt = randomHex_(16);
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    ADMIN_USER: ADMIN_USERNAME,
    ADMIN_SALT: salt,
    ADMIN_HASH: hashPassword_(password, salt),
    ADMIN_EPOCH: String(new Date().getTime())   // invalidates all earlier sessions
  });
  if (!props.getProperty('ADMIN_COMPANIES')) props.setProperty('ADMIN_COMPANIES', '*');
  CacheService.getScriptCache().remove('login_fail');
  Logger.log('=== ADMIN CREDENTIALS (shown once — store them in a password manager) ===');
  Logger.log('Username: ' + ADMIN_USERNAME);
  Logger.log('Password: ' + password);
  return password;
}

function adminLogin_(username, password) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fail') || 0);
  if (fails >= MAX_FAILED_LOGINS) {
    return {ok:false, error:'Too many failed attempts. Try again in 15 minutes.'};
  }
  const salt = props.getProperty('ADMIN_SALT');
  const storedHash = props.getProperty('ADMIN_HASH');
  if (!salt || !storedHash) {
    return {ok:false, error:'Admin is not set up yet. Run setupAdmin() in the Apps Script editor.'};
  }
  const userOk = safeEqual_(String(username || '').trim().toLowerCase(), String(props.getProperty('ADMIN_USER') || '').toLowerCase());
  const passOk = safeEqual_(hashPassword_(String(password || ''), salt), storedHash);
  if (!(userOk && passOk)) {
    cache.put('login_fail', String(fails + 1), LOCKOUT_SEC);
    return {ok:false, error:'Invalid username or password.'};
  }
  cache.remove('login_fail');

  const token = randomHex_(32);
  const companies = parseCompanyScope_(props.getProperty('ADMIN_COMPANIES'));
  const sess = {u: props.getProperty('ADMIN_USER'), c: companies, e: props.getProperty('ADMIN_EPOCH') || ''};
  cache.put('sess_' + sha256Hex_(token), JSON.stringify(sess), SESSION_TTL_SEC);
  return {ok:true, token:token, username:sess.u, companies:companies, expiresIn:SESSION_TTL_SEC};
}

function adminLogout_(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + sha256Hex_(String(token)));
  return {ok:true};
}

function getSession_(token) {
  if (!token || typeof token !== 'string' || token.length < 32) return null;
  const raw = CacheService.getScriptCache().get('sess_' + sha256Hex_(token));
  if (!raw) return null;
  let sess;
  try { sess = JSON.parse(raw); } catch (e) { return null; }
  const epoch = PropertiesService.getScriptProperties().getProperty('ADMIN_EPOCH') || '';
  if (sess.e !== epoch) return null;
  return sess;
}

function requireAdmin_(sess) {
  return sess ? null : {ok:false, error:'Not signed in', authRequired:true};
}

function parseCompanyScope_(v) {
  const s = String(v || '*').trim();
  if (!s || s === '*') return '*';
  return s.split('|').map(x => x.trim()).filter(Boolean);
}

function companyAllowed_(sess, company) {
  if (!sess || sess.c === '*') return true;
  const c = normName_(company);
  return sess.c.some(x => normName_(x) === c);
}

function hashPassword_(password, salt) {
  const key = Utilities.newBlob(password).getBytes();
  let h = Utilities.computeHmacSha256Signature(Utilities.newBlob(salt).getBytes(), key);
  for (let i = 0; i < HASH_ROUNDS; i++) h = Utilities.computeHmacSha256Signature(h, key);
  return bytesToHex_(h);
}

function sha256Hex_(s) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s));
}

function bytesToHex_(bytes) {
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

// UUID v4 values come from a CSPRNG; skip the fixed version/variant nibbles.
function randomHex_(len) {
  let out = '';
  while (out.length < len) {
    const h = Utilities.getUuid().replace(/-/g, '');
    out += h.slice(0, 12) + h.slice(13, 16) + h.slice(17);
  }
  return out.slice(0, len);
}

function generatePassword_(len) {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';  // 64 chars, no modulo bias
  const hex = randomHex_(len * 2);
  let pwd = '';
  for (let i = 0; i < len; i++) pwd += alphabet[parseInt(hex.substr(i * 2, 2), 16) & 63];
  return pwd;
}

function safeEqual_(a, b) {
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

// ─── Company Details ───

function companyFromRow_(r) {
  return {
    name:String(r[0]||''), phone:String(r[1]||''), email:String(r[2]||''),
    address:String(r[3]||''), gst:String(r[4]||''), currency:String(r[5]||'₹'),
    template:String(r[6]||'corporate'),
    invoiceFormat:String(r[8]||'') || DEFAULT_FORMATS.invoice,
    quotationFormat:String(r[9]||'') || DEFAULT_FORMATS.quotation
  };
}

function readCompanies_() {
  const sh = SS.getSheetByName('CompanyDetails');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COMPANY_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(companyFromRow_);
}

function normName_(s) { return String(s || '').trim().toLowerCase(); }

// With `name`: that company. Without: the first one (legacy behaviour).
function getCompany_(name) {
  const list = readCompanies_();
  if (!list.length) return {ok:true,company:null};
  if (name) {
    const found = list.filter(c => normName_(c.name) === normName_(name))[0];
    return {ok:true,company:found || null};
  }
  return {ok:true,company:list[0]};
}

function getCompanies_() {
  return {ok:true, companies: readCompanies_()};
}

function getClients_() {
  const clientsMap = {};
  ['Invoices', 'Quotations'].forEach(sheetName => {
    const sh = SS.getSheetByName(sheetName);
    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, DOC_COLS).getValues();
      rows.forEach(r => {
        const clientName = String(r[COL.client] || '').trim();
        const clientMobile = String(r[COL.clientMobile] || '').trim();
        if (clientName) {
          const key = clientName.toLowerCase();
          if (!clientsMap[key]) {
            clientsMap[key] = {
              name: clientName,
              mobile: clientMobile,
              lastDate: fmtDate_(r[COL.date]),
              company: String(r[COL.company] || ''),
              totalDocs: 1
            };
          } else {
            clientsMap[key].totalDocs += 1;
            if (clientMobile && !clientsMap[key].mobile) clientsMap[key].mobile = clientMobile;
          }
        }
      });
    }
  });
  return { ok: true, clients: Object.values(clientsMap) };
}

// Upsert by company name (a new name adds a company). Numbering formats are only changed when supplied.
function saveCompany_(c) {
  if (!c || !String(c.name || '').trim()) return {ok:false,error:'Company name is required'};
  const sh = SS.getSheetByName('CompanyDetails');
  const name = String(c.name).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let rowNo = -1, existing = null;
    if (sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COMPANY_HEADERS.length).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (normName_(rows[i][0]) === normName_(name)) { rowNo = i + 2; existing = rows[i]; break; }
      }
    }

    let invFmt = existing ? String(existing[8] || '') : '';
    let qtFmt = existing ? String(existing[9] || '') : '';
    if (c.invoiceFormat != null && String(c.invoiceFormat).trim() !== '') {
      const err = validateFormat_(c.invoiceFormat); if (err) return {ok:false,error:'Invoice format: ' + err};
      invFmt = String(c.invoiceFormat).trim();
    }
    if (c.quotationFormat != null && String(c.quotationFormat).trim() !== '') {
      const err = validateFormat_(c.quotationFormat); if (err) return {ok:false,error:'Quotation format: ' + err};
      qtFmt = String(c.quotationFormat).trim();
    }

    const row = [name,c.phone||'',c.email||'',c.address||'',c.gst||'',c.currency||'₹',c.template||'corporate',new Date(),invFmt,qtFmt];
    if (rowNo > 0) sh.getRange(rowNo,1,1,row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return {ok:true,message:'Company details saved'};
}

// ─── Document Save (Invoice or Quotation) ───

function docSheetName_(type) { return type === 'invoice' ? 'Invoices' : 'Quotations'; }

function findDocRow_(sh, docNo) {
  if (sh.getLastRow() < 2) return -1;
  const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const want = String(docNo).trim();
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() === want) return i + 2;
  }
  return -1;
}

function saveDocument_(d, sess) {
  if (!d) return {ok:false,error:'No document data'};
  const type = (d.docType || 'quotation').toLowerCase() === 'invoice' ? 'invoice' : 'quotation';
  const sh = SS.getSheetByName(docSheetName_(type));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const items = d.items || [];
    const subtotal = items.reduce((s,x) => s + Number(x.amount || 0), 0);
    const total = subtotal * (1 + Number(d.tax || 0) / 100);

    let docNo = String(d.docNo || '').trim();
    let rowNo = docNo ? findDocRow_(sh, docNo) : -1;
    let renumbered = false;

    if (rowNo > 0) {
      const existing = sh.getRange(rowNo, 1, 1, DOC_COLS).getValues()[0];
      if (d.isNew) {
        // Client asked for a *new* document but the number was taken meanwhile
        // (another session saved first). Never overwrite — take the next free one.
        docNo = allocateNumber_(type, d.company);
        rowNo = -1;
        renumbered = true;
      } else if (normName_(existing[COL.company]) !== normName_(d.company) && String(existing[COL.company]).trim() !== '') {
        return {ok:false, error:'Document number ' + docNo + ' already belongs to another company. Duplicate numbers are not allowed.'};
      }
    } else if (!docNo) {
      docNo = allocateNumber_(type, d.company);
    }

    const prev = rowNo > 0 ? sh.getRange(rowNo, 1, 1, DOC_COLS).getValues()[0] : null;
    // Internal note & payment status are admin-only: anonymous saves keep what is stored.
    let internalNote = prev ? String(prev[COL.internalNote] || '') : '';
    let paymentStatus = prev ? String(prev[COL.paymentStatus] || '') : '';
    if (sess && companyAllowed_(sess, d.company)) {
      if (d.internalNote != null) internalNote = String(d.internalNote);
      if (d.paymentStatus && PAYMENT_STATUSES.indexOf(d.paymentStatus) >= 0) paymentStatus = d.paymentStatus;
    }
    if (!paymentStatus) paymentStatus = 'Pending';

    const row = [
      docNo, d.date, d.company, d.phone, d.email, d.address, d.gst,
      d.client, d.clientMobile, d.validity, d.tax, d.profit, d.mode,
      d.simpleMode, d.notes, total, JSON.stringify(items), new Date(),
      internalNote, paymentStatus
    ];

    if (rowNo > 0) sh.getRange(rowNo, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    bumpCounter_(type, d.company, docNo);
    updateItemCatalog_(items);

    return {ok:true, message:(type === 'invoice' ? 'Invoice' : 'Quotation')+' saved', total:total,
            docNo:docNo, renumbered:renumbered};
  } finally {
    lock.releaseLock();
  }
}

// ─── Item Catalog ───

function updateItemCatalog_(items) {
  const sh = SS.getSheetByName('ItemCatalog');
  items.forEach(x => {
    const name = String(x.name||'').trim();
    if (!name) return;
    let rowNo = -1;
    for (let r = 1; r < sh.getLastRow(); r++) {
      if (String(sh.getRange(r+1,1).getValue()).trim().toLowerCase() === name.toLowerCase()) {
        rowNo = r + 1;
        break;
      }
    }
    // Save rate as LastRate (column 4) along with cost and profit
    const itemRow = [name, Number(x.c||0), Number(x.p||0), Number(x.rate||0), new Date()];
    if (rowNo > 0) sh.getRange(rowNo, 1, 1, 5).setValues([itemRow]);
    else sh.appendRow(itemRow);
  });
}

function getItems_() {
  const sh = SS.getSheetByName('ItemCatalog');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]),
      cost: Number(r[1]||0),
      defaultProfit: Number(r[2]||0),
      lastRate: Number(r[3]||0)
    }));
}

// ═══════════════════════════════════════════════════════════════════
//  Company-wise numbering
//
//  Each company has an invoice format and a quotation format, e.g.
//    INV-{YYYY}-{SEQ:3}   ->  INV-2026-001   (the legacy default)
//    ACME/{YY}{MM}/{SEQ:4} -> ACME/2610/0007
//  Tokens: {YYYY} {YY} {MM} {SEQ} {SEQ:n} (n = zero-pad width 1-8).
//  Documents whose numbers match the same rendered pattern share one
//  sequence, so two companies can never be handed the same number even if
//  they use identical formats. Companies with different prefixes number
//  independently. Existing document numbers are never altered.
// ═══════════════════════════════════════════════════════════════════

const TOKEN_RE = /\{(YYYY|YY|MM|SEQ(?::\d+)?)\}/;

function validateFormat_(fmt) {
  fmt = String(fmt || '').trim();
  if (!fmt) return 'format is empty';
  if (fmt.length > 60) return 'format is too long (max 60 characters)';
  const parts = fmt.split(new RegExp(TOKEN_RE.source, 'g'));
  let seqCount = 0;
  for (let i = 0; i < parts.length; i++) {
    if (i % 2 === 1) {
      if (parts[i].indexOf('SEQ') === 0) {
        seqCount++;
        const m = parts[i].match(/:(\d+)/);
        if (m && (Number(m[1]) < 1 || Number(m[1]) > 8)) return '{SEQ:n} width must be 1 to 8';
      }
    } else if (!/^[A-Za-z0-9\-_\/. #]*$/.test(parts[i])) {
      return 'only letters, digits and - _ / . # space are allowed besides {YYYY} {YY} {MM} {SEQ:n}';
    }
  }
  if (seqCount !== 1) return 'format must contain exactly one {SEQ} or {SEQ:n}';
  return '';
}

function formatFor_(type, company) {
  const c = company ? readCompanies_().filter(x => normName_(x.name) === normName_(company))[0] : null;
  if (!c) return DEFAULT_FORMATS[type];
  return (type === 'invoice' ? c.invoiceFormat : c.quotationFormat) || DEFAULT_FORMATS[type];
}

function periodTokens_(date) {
  const y = date.getFullYear();
  return {YYYY: String(y), YY: String(y).slice(-2), MM: ('0' + (date.getMonth() + 1)).slice(-2)};
}

function renderFormat_(fmt, date, seq) {
  const p = periodTokens_(date);
  return fmt.replace(new RegExp(TOKEN_RE.source, 'g'), (m, t) => {
    if (t === 'YYYY' || t === 'YY' || t === 'MM') return p[t];
    const w = t.indexOf(':') > 0 ? Number(t.split(':')[1]) : 0;
    return typeof seq === 'number' ? String(seq).padStart(w, '0') : String(seq);   // marker/placeholder stays unpadded
  });
}

// Regex that captures the sequence in numbers of the *current period* of `fmt`.
function seqRegex_(fmt, date) {
  const marker = '\u0001';
  const rendered = renderFormat_(fmt, date, marker);
  const escaped = rendered.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&').replace(marker, '(\\d+)');
  return new RegExp('^' + escaped + '$');
}

function poolKey_(fmt, date) { return renderFormat_(fmt, date, '#'); }

// Counter rows: DocType | Year | LastNumber | Pool. Legacy rows have an empty Pool
// and belong to the default format's pool.
function counterLookup_(type, fmt, date) {
  const sh = SS.getSheetByName('DocCounters');
  const values = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const isDefault = fmt === DEFAULT_FORMATS[type];
  const key = poolKey_(fmt, date);
  const year = fmt.indexOf('{YYYY}') >= 0 || fmt.indexOf('{YY}') >= 0 || fmt.indexOf('{MM}') >= 0 ? date.getFullYear() : 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (String(v[0]).toLowerCase() !== type) continue;
    const pool = String(v[3] || '');
    if (isDefault ? (pool === '' && Number(v[1]) === year) : pool === key) {
      return {sh:sh, rowNo:i + 2, last:Number(v[2] || 0), year:year, key:key, isDefault:isDefault};
    }
  }
  return {sh:sh, rowNo:-1, last:0, year:year, key:key, isDefault:isDefault};
}

function maxSeqInSheet_(type, fmt, date) {
  const sh = SS.getSheetByName(docSheetName_(type));
  if (sh.getLastRow() < 2) return {max:0, used:{}};
  const re = seqRegex_(fmt, date);
  const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const used = {};
  let max = 0;
  col.forEach(r => {
    const s = String(r[0]).trim();
    used[s] = true;
    const m = s.match(re);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return {max:max, used:used};
}

// Next number without consuming it.
function peekNumber_(type, company) {
  const date = new Date();
  const fmt = formatFor_(type, company);
  const ctr = counterLookup_(type, fmt, date);
  const scan = maxSeqInSheet_(type, fmt, date);
  let seq = Math.max(ctr.last, scan.max) + 1;
  while (scan.used[renderFormat_(fmt, date, seq)]) seq++;
  return {docNo:renderFormat_(fmt, date, seq), seq:seq, format:fmt};
}

function getNextNumber_(docType, company) {
  const type = String(docType || 'quotation').toLowerCase() === 'invoice' ? 'invoice' : 'quotation';
  const n = peekNumber_(type, company);
  return {ok:true, docNo:n.docNo, nextNumber:n.seq, format:n.format};
}

// Must be called under the script lock; the caller then writes the row.
function allocateNumber_(type, company) {
  const n = peekNumber_(type, company);
  return n.docNo;
}

// Record the high-water mark when a saved number belongs to the company's current pool.
function bumpCounter_(type, company, docNo) {
  const date = new Date();
  const fmt = formatFor_(type, company);
  const m = String(docNo).trim().match(seqRegex_(fmt, date));
  if (!m) return;
  const seq = Number(m[1]);
  const ctr = counterLookup_(type, fmt, date);
  if (ctr.rowNo > 0) {
    if (seq > ctr.last) ctr.sh.getRange(ctr.rowNo, 3).setValue(seq);
  } else {
    ctr.sh.appendRow([type, ctr.year, seq, ctr.isDefault ? '' : ctr.key]);
  }
}

// ─── Fetch Documents (public: never exposes admin-only fields) ───

function fetchDocuments_(docNo, name, mobile, docType, sess) {
  const results = [];
  if (!docType || docType === 'invoice') {
    results.push(...searchSheet_('Invoices', 'invoice', docNo, name, mobile, sess));
  }
  if (!docType || docType === 'quotation') {
    results.push(...searchSheet_('Quotations', 'quotation', docNo, name, mobile, sess));
  }
  return {ok:true, rows:results, admin:!!sess};
}

function parseItems_(raw) {
  try { return JSON.parse(raw || '[]'); } catch(e) { return []; }
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}

// Full record. Admin-only fields are included only when `admin` is true.
function docFromRow_(r, type, admin) {
  const doc = {
    docType: type,
    docNo: String(r[COL.docNo]), date: fmtDate_(r[COL.date]), company: r[COL.company], phone: r[COL.phone],
    email: r[COL.email], address: r[COL.address], gst: r[COL.gst],
    client: r[COL.client], clientMobile: r[COL.clientMobile], validity: r[COL.validity],
    tax: r[COL.tax], profit: r[COL.profit], mode: r[COL.mode], simpleMode: r[COL.simpleMode],
    notes: r[COL.notes], total: r[COL.total], items: parseItems_(r[COL.items])
  };
  if (admin) {
    doc.internalNote = String(r[COL.internalNote] || '');
    doc.paymentStatus = String(r[COL.paymentStatus] || '') || 'Pending';
    doc.updatedAt = r[COL.updatedAt] instanceof Date ? r[COL.updatedAt].toISOString() : String(r[COL.updatedAt] || '');
  }
  return doc;
}

function searchSheet_(sheetName, docType, docNo, name, mobile, sess) {
  const sh = SS.getSheetByName(sheetName);
  if (sh.getLastRow() < 2) return [];

  const rows = sh.getRange(2, 1, sh.getLastRow()-1, DOC_COLS).getValues();
  const out = [];

  rows.forEach(r => {
    const dNo = String(r[COL.docNo]||'').trim();
    const mob = String(r[COL.clientMobile]||'').replace(/\D/g,'');

    const noMatch = docNo && dNo.toLowerCase().includes(String(docNo).toLowerCase());
    const nameMatch = name && String(r[COL.client]||'').toLowerCase().includes(String(name).toLowerCase());
    const mobMatch = mobile && mob.includes(String(mobile).replace(/\D/g,''));

    if ((docNo && noMatch) || (name && nameMatch) || (mobile && mobMatch)) {
      out.push(docFromRow_(r, docType, !!(sess && companyAllowed_(sess, r[COL.company]))));
    }
  });

  return out;
}

// ═══════════════════════════════════════════════════════════════════
//  Admin dashboard API (all require a valid session)
// ═══════════════════════════════════════════════════════════════════

function readAllDocs_(sess) {
  const docs = [];
  [['Invoices','invoice'],['Quotations','quotation']].forEach(p => {
    const sh = SS.getSheetByName(p[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, DOC_COLS).getValues();
    rows.forEach((r, i) => {
      if (!String(r[COL.docNo]).trim()) return;
      if (!companyAllowed_(sess, r[COL.company])) return;
      docs.push({type:p[1], rowNo:i + 2, r:r});
    });
  });
  return docs;
}

function listRow_(x) {
  const r = x.r;
  return {
    docType: x.type,
    docNo: String(r[COL.docNo]), date: fmtDate_(r[COL.date]), company: String(r[COL.company] || ''),
    client: String(r[COL.client] || ''), clientMobile: String(r[COL.clientMobile] || ''),
    total: Number(r[COL.total] || 0),
    paymentStatus: String(r[COL.paymentStatus] || '') || 'Pending',
    internalNote: String(r[COL.internalNote] || ''),
    updatedAt: r[COL.updatedAt] instanceof Date ? r[COL.updatedAt].toISOString() : String(r[COL.updatedAt] || '')
  };
}

function adminList_(q, sess) {
  const all = readAllDocs_(sess).map(listRow_);
  const companies = Array.from(new Set(all.map(d => d.company).filter(Boolean))).sort();

  const type = String(q.docType || '');
  const company = String(q.company || '');
  const status = String(q.status || '');
  const from = String(q.from || ''), to = String(q.to || '');
  const term = String(q.q || '').trim().toLowerCase();

  const filtered = all.filter(d => {
    if (type && d.docType !== type) return false;
    if (company && normName_(d.company) !== normName_(company)) return false;
    if (status && d.paymentStatus !== status) return false;
    if (from && d.date < from) return false;
    if (to && d.date > to) return false;
    if (term) {
      const hay = [d.docNo, d.company, d.client, d.clientMobile, d.internalNote].join(' ').toLowerCase();
      if (hay.indexOf(term) < 0) return false;
    }
    return true;
  });

  const summary = {
    invoiceCount:0, quotationCount:0, invoiceTotal:0, quotationTotal:0,
    paidTotal:0, outstandingTotal:0, paidCount:0, outstandingCount:0
  };
  filtered.forEach(d => {
    if (d.docType === 'invoice') {
      summary.invoiceCount++; summary.invoiceTotal += d.total;
      if (d.paymentStatus === 'Paid') { summary.paidCount++; summary.paidTotal += d.total; }
      else { summary.outstandingCount++; summary.outstandingTotal += d.total; }
    } else {
      summary.quotationCount++; summary.quotationTotal += d.total;
    }
  });

  const sortKey = ['date','total','docNo','client','company','paymentStatus'].indexOf(q.sort) >= 0 ? q.sort : 'date';
  const dir = q.dir === 'asc' ? 1 : -1;
  filtered.sort((a, b) => {
    const x = a[sortKey], y = b[sortKey];
    const c = (typeof x === 'number') ? x - y : String(x).localeCompare(String(y), undefined, {numeric:true});
    return (c || String(a.docNo).localeCompare(String(b.docNo), undefined, {numeric:true})) * dir;
  });

  const pageSize = Math.min(100, Math.max(5, Number(q.pageSize) || 25));
  const pages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const page = Math.min(pages, Math.max(1, Number(q.page) || 1));
  return {
    ok:true, rows:filtered.slice((page - 1) * pageSize, page * pageSize),
    total:filtered.length, page:page, pages:pages, pageSize:pageSize,
    summary:summary, companies:companies, statuses:PAYMENT_STATUSES
  };
}

function adminGet_(q, sess) {
  const type = String(q.docType) === 'invoice' ? 'invoice' : 'quotation';
  const sh = SS.getSheetByName(docSheetName_(type));
  const rowNo = findDocRow_(sh, q.docNo || '');
  if (rowNo < 0) return {ok:false, error:'Document not found'};
  const r = sh.getRange(rowNo, 1, 1, DOC_COLS).getValues()[0];
  if (!companyAllowed_(sess, r[COL.company])) return {ok:false, error:'Access denied for this company'};
  return {ok:true, doc:docFromRow_(r, type, true)};
}

// Admin may change only payment status and internal note from the dashboard.
function adminUpdate_(q, sess) {
  const type = String(q.docType) === 'invoice' ? 'invoice' : 'quotation';
  const sh = SS.getSheetByName(docSheetName_(type));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rowNo = findDocRow_(sh, q.docNo || '');
    if (rowNo < 0) return {ok:false, error:'Document not found'};
    const company = sh.getRange(rowNo, COL.company + 1).getValue();
    if (!companyAllowed_(sess, company)) return {ok:false, error:'Access denied for this company'};
    if (q.paymentStatus != null) {
      if (PAYMENT_STATUSES.indexOf(q.paymentStatus) < 0) return {ok:false, error:'Invalid payment status'};
      sh.getRange(rowNo, COL.paymentStatus + 1).setValue(q.paymentStatus);
    }
    if (q.internalNote != null) sh.getRange(rowNo, COL.internalNote + 1).setValue(String(q.internalNote));
    sh.getRange(rowNo, COL.updatedAt + 1).setValue(new Date());
  } finally {
    lock.releaseLock();
  }
  return {ok:true};
}

// ─── Utility ───

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
