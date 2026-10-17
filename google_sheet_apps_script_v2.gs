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

const API_VERSION = 6;   // bump when the frontend needs new backend actions
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
  internalNote: 18, paymentStatus: 19, comparisonId: 20, currency: 21,
  discount: 22, discountType: 23
};
const DOC_HEADERS = [
  'DocNo','Date','Company','Phone','Email','Address','GSTIN','Client','ClientMobile',
  'Validity','Tax','DefaultProfit','Mode','SimpleMode','Notes','Total','ItemsJSON',
  'UpdatedAt','InternalNote','PaymentStatus','ComparisonId','Currency','Discount','DiscountType'
];
const DOC_COLS = DOC_HEADERS.length;
const COMPANY_HEADERS = ['CompanyName','Phone','Email','Address','GSTIN','Currency','Template','UpdatedAt','InvoiceFormat','QuotationFormat'];
const COUNTER_HEADERS = ['DocType','Year','LastNumber','Pool'];

// ─── Money rules (mirror of qm-calc.js — whole numbers, half-up, rounded at every level) ───

function rnd_(x) {
  const n = Number(x);
  if (!isFinite(n)) return 0;
  return n < 0 ? -Math.round(-n + 1e-9) : Math.round(n + 1e-9);
}

function lineAmount_(item) {
  const a = item.amount;
  return (a !== undefined && a !== null && a !== '') ? rnd_(a) : rnd_((Number(item.q) || 0) * rnd_(item.rate));
}

function totals_(items, taxPct, discount, discountType) {
  const subtotal = (items || []).reduce((s, x) => s + lineAmount_(x), 0);
  let disc = discountType === 'amount' ? rnd_(discount) : rnd_(subtotal * (Number(discount) || 0) / 100);
  disc = Math.min(Math.max(disc, 0), subtotal);
  const taxable = subtotal - disc;
  const tax = rnd_(taxable * (Number(taxPct) || 0) / 100);
  return {subtotal:subtotal, discount:disc, taxable:taxable, tax:tax, total:taxable + tax};
}

// Items with whole-number rate/amount, as stored in the sheet.
function normalizeItems_(items) {
  return (items || []).map(x => {
    const q = Number(x.q) || 0, rate = rnd_(x.rate);
    const amount = lineAmount_(Object.assign({}, x, {q:q, rate:rate}));
    // profit = what the customer pays minus cost (whole numbers); internal only, never printed
    return Object.assign({}, x, {q:q, rate:rate, amount:amount, profit:amount - rnd_(q * (Number(x.c) || 0))});
  });
}

// Internal view = internal note, payment status and the profit breakdown in the staff form.
// Default: visible to everyone using the app (it is the internal form; it is never printed or exported).
// Strict: set Script Property INTERNAL_NOTES_ADMIN_ONLY = 1 to require an admin session for them.
function internalStrict_() {
  return PropertiesService.getScriptProperties().getProperty('INTERNAL_NOTES_ADMIN_ONLY') === '1';
}
function internalAllowed_(sess, company) {
  return internalStrict_() ? !!(sess && companyAllowed_(sess, company)) : true;
}

// ─── Setup ───

function setup_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('setup_v6')) return;

  ensureSheet_('CompanyDetails', COMPANY_HEADERS, false);
  ensureSheet_('Invoices', DOC_HEADERS, true);
  ensureSheet_('Quotations', DOC_HEADERS, true);
  ensureSheet_('ItemCatalog', ['ItemName','CostRate','DefaultProfit','LastRate','UpdatedAt'], false);
  ensureSheet_('DocCounters', COUNTER_HEADERS, false);

  cache.put('setup_v6', '1', 600);
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
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (name === 'CompanyDetails' && sh.getMaxColumns() > headers.length && sh.getRange(1, headers.length + 1).getValue() === 'CommissionPercent') sh.getRange(1, headers.length + 1).clearContent();   // feature removed
  }
  if (isDoc) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); // keep DocNo as text
  return sh;
}

function doGet(e) {
  setup_();
  return json_({ok:true,message:'Quotation & Invoice Maker API v' + API_VERSION + ' is running',version:API_VERSION});
}

function doPost(e) {
  try {
    setup_();
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const sess = getSession_(data.token);
    switch(data.action) {
      case 'ping':
        return json_({ok:true,message:'Connected (backend v' + API_VERSION + ')',version:API_VERSION,admin:!!sess,notesAdminOnly:internalStrict_()});
      case 'getItems':
        return json_({ok:true,items:getItems_()});
      case 'saveDocument':
        return json_(saveDocument_(data, sess));
      case 'saveComparison':
        return json_(saveComparison_(data, sess));
      case 'getComparison':
        return json_(getComparison_(data, sess));
      case 'fetchDocuments':
        return json_(fetchDocuments_(data.docNo||'', data.name||'', data.mobile||'', data.docType||'', sess));
      case 'getCompany':
        return json_(getCompany_(data.name, sess));
      case 'getCompanies':
        return json_(getCompanies_(sess));
      case 'saveCompany':
        return json_(saveCompany_(data.company, sess));
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
      case 'adminCompanies':
        return json_(requireAdmin_(sess) || adminCompanies_(sess));
      case 'adminSaveCompany':
        return json_(requireAdmin_(sess) || adminSaveCompany_(data, sess));
      case 'adminSaveDocument':
        return json_(requireAdmin_(sess) || adminSaveDocument_(data, sess));
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
function getCompany_(name, sess) {
  const list = readCompanies_();
  if (!list.length) return {ok:true,company:null};
  if (name) {
    const found = list.filter(c => normName_(c.name) === normName_(name))[0];
    return {ok:true,company:found || null};
  }
  return {ok:true,company:list[0]};
}

function getCompanies_(sess) {
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
function saveCompany_(c, sess) {
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
    // Numbering formats are admin-only: anonymous saves keep the stored formats.
    if (sess && c.invoiceFormat != null && String(c.invoiceFormat).trim() !== '') {
      const err = validateFormat_(c.invoiceFormat); if (err) return {ok:false,error:'Invoice format: ' + err};
      invFmt = String(c.invoiceFormat).trim();
    }
    if (sess && c.quotationFormat != null && String(c.quotationFormat).trim() !== '') {
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
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return saveDocumentCore_(d, sess, {});
  } finally {
    lock.releaseLock();
  }
}

// Caller must hold the script lock.
function saveDocumentCore_(d, sess, opts) {
  const type = (d.docType || 'quotation').toLowerCase() === 'invoice' ? 'invoice' : 'quotation';
  const sh = SS.getSheetByName(docSheetName_(type));
  const items = normalizeItems_(d.items);

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
  if (internalAllowed_(sess, d.company)) {
    if (d.internalNote != null) internalNote = String(d.internalNote);
    if (d.paymentStatus && PAYMENT_STATUSES.indexOf(d.paymentStatus) >= 0) paymentStatus = d.paymentStatus;
  }
  if (!paymentStatus) paymentStatus = 'Pending';
  const comparisonId = String(d.comparisonId || (prev ? prev[COL.comparisonId] : '') || '');
  const currency = String(d.currency || (prev ? prev[COL.currency] : '') || '');
  const discount = d.discount != null && d.discount !== '' ? Math.max(0, Number(d.discount) || 0) : (prev ? Number(prev[COL.discount] || 0) : 0);
  const discountType = (d.discountType === 'amount' || d.discountType === 'percent') ? d.discountType : (prev && prev[COL.discountType] ? String(prev[COL.discountType]) : 'percent');
  const total = totals_(items, d.tax, discount, discountType).total;

  const row = [
    docNo, d.date, d.company, d.phone, d.email, d.address, d.gst,
    d.client, d.clientMobile, d.validity, d.tax, d.profit, d.mode,
    d.simpleMode, d.notes, total, JSON.stringify(items), new Date(),
    internalNote, paymentStatus, comparisonId, currency, discount, discountType
  ];

  if (rowNo > 0) sh.getRange(rowNo, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  bumpCounter_(type, d.company, docNo);
  if (!opts.skipCatalog) updateItemCatalog_(items);

  return {ok:true, message:(type === 'invoice' ? 'Invoice' : 'Quotation')+' saved', total:total,
          docNo:docNo, renumbered:renumbered, docType:type};
}

// ─── 3-company comparison (one product list, one quotation per company) ───

function newComparisonId_() {
  const d = new Date();
  return 'CMP-' + d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2) + '-' + randomHex_(6).toUpperCase();
}

// Saves every quotation of a comparison under ONE lock, validating first so a
// failure never leaves a half-saved set. Company 1 (quotes[0]) is the base: only
// its rates update the item catalog, so variations never overwrite saved rates.
function saveComparison_(d, sess) {
  const quotes = (d && d.quotes) || [];
  if (quotes.length < 2 || quotes.length > 3) return {ok:false, error:'A comparison needs 2 or 3 companies.'};
  const seen = {};
  for (let i = 0; i < quotes.length; i++) {
    const n = normName_(quotes[i].company);
    if (!n) return {ok:false, error:'Select a company for every column.'};
    if (seen[n]) return {ok:false, error:'Each column must use a different company.'};
    seen[n] = true;
  }
  const sig = it => JSON.stringify((it || []).map(x => [String(x.name || '').trim().toLowerCase(), Number(x.q || 0)]));
  const base = sig(quotes[0].items);
  for (let i = 1; i < quotes.length; i++) {
    if (sig(quotes[i].items) !== base) return {ok:false, error:'All companies must quote the same product list and quantities.'};
  }
  if (!(quotes[0].items || []).length) return {ok:false, error:'Add at least one product.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SS.getSheetByName('Quotations');
    // Pre-flight: a given, non-new number must belong to the same company.
    for (let i = 0; i < quotes.length; i++) {
      const q = quotes[i];
      if (q.docNo && !q.isNew) {
        const rn = findDocRow_(sh, q.docNo);
        if (rn > 0 && normName_(sh.getRange(rn, COL.company + 1).getValue()) !== normName_(q.company)) {
          return {ok:false, error:'Document number ' + q.docNo + ' already belongs to another company.'};
        }
      }
    }
    const cmpId = String(d.comparisonId || '') || newComparisonId_();
    const results = quotes.map((q, i) => {
      const items = (q.items || []).map(x => { const rate = rnd_(x.rate), qn = Number(x.q) || 0; return Object.assign({}, x, {q: qn, rate: rate, amount: rnd_(qn * rate)}); });
      const r = saveDocumentCore_(Object.assign({}, d.common || {}, q, {docType:'quotation', comparisonId:cmpId, items:items}),
                                  sess, {skipCatalog: i > 0});
      r.company = q.company;
      return r;
    });
    return {ok:true, comparisonId:cmpId, results:results};
  } finally {
    lock.releaseLock();
  }
}

// Load a saved comparison by id or by any of its quotation numbers.
function getComparison_(q, sess) {
  const sh = SS.getSheetByName('Quotations');
  if (sh.getLastRow() < 2) return {ok:false, error:'Comparison not found'};
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, DOC_COLS).getValues();
  let id = String(q.comparisonId || '').trim();
  if (!id && q.docNo) {
    const hit = rows.filter(r => String(r[COL.docNo]).trim() === String(q.docNo).trim())[0];
    id = hit ? String(hit[COL.comparisonId] || '') : '';
  }
  if (!id) return {ok:false, error:'This quotation is not part of a comparison.'};
  const quotes = rows.filter(r => String(r[COL.comparisonId]) === id)
    .map(r => docFromRow_(r, 'quotation', internalAllowed_(sess, r[COL.company])));
  if (!quotes.length) return {ok:false, error:'Comparison not found'};
  return {ok:true, comparisonId:id, quotes:quotes};
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
    notes: r[COL.notes], total: rnd_(r[COL.total]), items: parseItems_(r[COL.items]),
    discount: Number(r[COL.discount] || 0), discountType: String(r[COL.discountType] || '') || 'percent',
    comparisonId: String(r[COL.comparisonId] || ''), currency: String(r[COL.currency] || '')
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
      out.push(docFromRow_(r, docType, internalAllowed_(sess, r[COL.company])));
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
    total: rnd_(r[COL.total]), comparisonId: String(r[COL.comparisonId] || ''), currency: String(r[COL.currency] || '') || '₹',
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
    paidTotal:0, outstandingTotal:0, paidCount:0, outstandingCount:0, currencies:[]
  };
  summary.currencies = Array.from(new Set(filtered.map(d => d.currency)));
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
  const doc = docFromRow_(r, type, true);
  let siblings = [];
  if (doc.comparisonId) {
    siblings = readAllDocs_(sess).filter(x => x.type === 'quotation' && String(x.r[COL.comparisonId]) === doc.comparisonId).map(listRow_);
  }
  return {ok:true, doc:doc, siblings:siblings};
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

// Full edit of an existing document by an admin: items, quantities, rates, discount, tax,
// client/validity/notes and the admin-only fields. Number, company and comparison link stay fixed.
// Totals are always recomputed here, so the sheet is the single source of truth.
function adminSaveDocument_(q, sess) {
  const type = String(q.docType) === 'invoice' ? 'invoice' : 'quotation';
  const f = q.fields || {};
  const items = f.items;
  if (!Array.isArray(items) || !items.length) return {ok:false, error:'A document needs at least one item.'};
  for (let i = 0; i < items.length; i++) {
    if (!String(items[i].name || '').trim()) return {ok:false, error:'Item ' + (i + 1) + ' has no name.'};
    if (!(Number(items[i].q) >= 0) || !(Number(items[i].rate) >= 0)) return {ok:false, error:'Item ' + (i + 1) + ': quantity and rate must be 0 or more.'};
  }
  if (!(Number(f.tax || 0) >= 0 && Number(f.tax || 0) <= 100)) return {ok:false, error:'Tax % must be between 0 and 100.'};
  if (f.discountType != null && f.discountType !== 'amount' && f.discountType !== 'percent') return {ok:false, error:'Invalid discount type.'};
  if (!(Number(f.discount || 0) >= 0) || (f.discountType === 'percent' && Number(f.discount || 0) > 100)) return {ok:false, error:'Invalid discount.'};
  if (f.paymentStatus != null && PAYMENT_STATUSES.indexOf(f.paymentStatus) < 0) return {ok:false, error:'Invalid payment status'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SS.getSheetByName(docSheetName_(type));
    const rowNo = findDocRow_(sh, q.docNo || '');
    if (rowNo < 0) return {ok:false, error:'Document not found'};
    const prev = sh.getRange(rowNo, 1, 1, DOC_COLS).getValues()[0];
    if (!companyAllowed_(sess, prev[COL.company])) return {ok:false, error:'Access denied for this company'};
    const pick = (k, cur) => (f[k] != null ? f[k] : cur);
    const merged = {
      docType: type, docNo: String(prev[COL.docNo]), isNew: false,
      date: pick('date', fmtDate_(prev[COL.date])), company: prev[COL.company],
      phone: prev[COL.phone], email: prev[COL.email], address: prev[COL.address], gst: prev[COL.gst],
      client: pick('client', prev[COL.client]), clientMobile: pick('clientMobile', prev[COL.clientMobile]),
      validity: pick('validity', prev[COL.validity]), tax: pick('tax', prev[COL.tax]),
      profit: prev[COL.profit], mode: prev[COL.mode], simpleMode: prev[COL.simpleMode],
      notes: pick('notes', prev[COL.notes]), items: items, discount: f.discount, discountType: f.discountType,
      internalNote: pick('internalNote', null), paymentStatus: f.paymentStatus
    };
    const r = saveDocumentCore_(merged, sess, {skipCatalog: true});
    return r.ok ? {ok:true, docNo:r.docNo, total:r.total, message:'Saved'} : r;
  } finally {
    lock.releaseLock();
  }
}

// ── Company management (admin) ──

const TEMPLATE_IDS = ['corporate','govt','modern','navy','classic','accent','sidebar','gradient','executive','project','vintage','startup','twotone'];

function validateCompany_(c) {
  if (!c) return 'No company data';
  if (!String(c.name || '').trim()) return 'Company name is required';
  if (String(c.name).trim().length > 80) return 'Company name is too long';
  if (!String(c.phone || '').trim()) return 'Phone number is required';
  if (!/^[0-9+()\-\s]{6,20}$/.test(String(c.phone).trim())) return 'Phone number looks invalid';
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(c.email || '').trim())) return 'A valid email is required';
  if (!String(c.address || '').trim()) return 'Address is required';
  if (TEMPLATE_IDS.indexOf(String(c.template || 'corporate')) < 0) return 'Unknown template';
  if (!String(c.currency || '₹').trim() || String(c.currency || '₹').length > 5) return 'Currency symbol is invalid';
  return '';
}

function adminCompanies_(sess) {
  const docs = readAllDocs_(sess);
  const list = readCompanies_().filter(c => companyAllowed_(sess, c.name)).map(c => {
    const mine = docs.filter(x => normName_(x.r[COL.company]) === normName_(c.name));
    return Object.assign({}, c, {
      invoiceCount: mine.filter(x => x.type === 'invoice').length,
      quotationCount: mine.filter(x => x.type === 'quotation').length,
      nextInvoiceNo: peekNumber_('invoice', c.name).docNo,
      nextQuotationNo: peekNumber_('quotation', c.name).docNo
    });
  });
  return {ok:true, companies:list, templates:TEMPLATE_IDS};
}

// mode 'create' refuses an existing name (no accidental overwrite); 'update' requires it.
function adminSaveCompany_(q, sess) {
  const c = q.company;
  const err = validateCompany_(c);
  if (err) return {ok:false, error:err};
  if (!companyAllowed_(sess, c.name)) return {ok:false, error:'Access denied for this company'};
  const exists = readCompanies_().some(x => normName_(x.name) === normName_(c.name));
  if (q.mode === 'update' && !exists) return {ok:false, error:'Company not found'};
  if (q.mode !== 'update' && exists) return {ok:false, error:'A company with this name already exists.'};
  const res = saveCompany_(c, sess);
  return res.ok ? {ok:true, message: exists ? 'Company updated' : 'Company added'} : res;
}

// ─── Utility ───

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
