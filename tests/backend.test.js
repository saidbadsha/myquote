// Run: node tests/backend.test.js
const assert = require('assert');
const { makeEnv } = require('./mock-env');
const QMCalc = require('../qm-calc.js');

const year = new Date().getFullYear();
const doc = (o = {}) => Object.assign({
  docType: 'invoice', docNo: '', date: '2026-10-17', company: 'Acme', client: 'Bob', clientMobile: '9999999999',
  tax: 18, notes: 'public note', internalNote: '', items: [{name: 'x', amount: 1000}], isNew: true
}, o);

let passed = 0;
const test = (name, fn) => { try { fn(); passed++; console.log('  ✓', name); } catch (e) { console.error('  ✗', name, '\n   ', e.message); process.exitCode = 1; } };

// ── Numbering ──
console.log('Numbering');
{
  const e = makeEnv();
  const adminTok = (() => { const pw = e.setupAdmin(); return e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token; })();
  test('peek does not consume numbers', () => {
    const a = e.call({action: 'getNextNumber', docType: 'invoice', company: 'Acme'});
    const b = e.call({action: 'getNextNumber', docType: 'invoice', company: 'Acme'});
    assert.strictEqual(a.docNo, `INV-${year}-001`);
    assert.strictEqual(b.docNo, a.docNo);
  });
  test('default format preserved (legacy INV-YYYY-NNN / QT-YYYY-NNN)', () => {
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'quotation'}).docNo, `QT-${year}-001`);
  });
  test('save advances sequence, no gaps', () => {
    const s = e.call({action: 'saveDocument', ...doc({docNo: `INV-${year}-001`})});
    assert.ok(s.ok); assert.strictEqual(s.docNo, `INV-${year}-001`);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'invoice'}).docNo, `INV-${year}-002`);
  });
  test('stale "new" number is re-assigned, never overwrites', () => {
    const s = e.call({action: 'saveDocument', ...doc({docNo: `INV-${year}-001`, client: 'Someone else'})});
    assert.ok(s.ok && s.renumbered); assert.strictEqual(s.docNo, `INV-${year}-002`);
    const rows = e.sheets.Invoices.rows.slice(1);
    assert.strictEqual(rows.length, 2);
    assert.strictEqual(rows[0][7], 'Bob');
  });
  test('editing an existing doc (isNew=false) updates in place', () => {
    const s = e.call({action: 'saveDocument', ...doc({docNo: `INV-${year}-001`, isNew: false, client: 'Bobby'})});
    assert.ok(s.ok && !s.renumbered);
    assert.strictEqual(e.sheets.Invoices.rows.length, 3);
    assert.strictEqual(e.sheets.Invoices.rows[1][7], 'Bobby');
  });
  test('editing a doc number owned by another company is rejected', () => {
    const s = e.call({action: 'saveDocument', ...doc({docNo: `INV-${year}-001`, isNew: false, company: 'Other Co'})});
    assert.ok(!s.ok && /another company/.test(s.error));
  });
  test('company-wise formats are independent and validated', () => {
    assert.ok(e.call({action: 'saveCompany', token: adminTok, company: {name: 'Acme', invoiceFormat: 'ACM/{YY}{MM}/{SEQ:4}', quotationFormat: 'AQ-{SEQ}'}}).ok);
    assert.ok(e.call({action: 'saveCompany', token: adminTok, company: {name: 'Beta', invoiceFormat: 'BET-{YYYY}-{SEQ:3}'}}).ok);
    const yy = String(year).slice(-2), mm = ('0' + (new Date().getMonth() + 1)).slice(-2);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'invoice', company: 'Acme'}).docNo, `ACM/${yy}${mm}/0001`);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'invoice', company: 'Beta'}).docNo, `BET-${year}-001`);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'quotation', company: 'Acme'}).docNo, 'AQ-1');
    const bad = ['no-seq', 'A-{SEQ}-{SEQ}', 'A-{SEQ:9}', 'A<{SEQ}>', '$$'].map(f => e.call({action: 'saveCompany', token: adminTok, company: {name: 'Acme', invoiceFormat: f}}));
    bad.forEach(r => assert.ok(!r.ok, 'should reject'));
    assert.strictEqual(e.call({action: 'getCompanies'}).companies.length, 2);
  });
  test('anonymous saveCompany cannot change numbering formats', () => {
    e.call({action: 'saveCompany', company: {name: 'Acme', invoiceFormat: 'HACK-{SEQ}'}});
    assert.strictEqual(e.call({action: 'getCompany', name: 'Acme'}).company.invoiceFormat, 'ACM/{YY}{MM}/{SEQ:4}');
  });
  test('saving with custom format uses & advances that company sequence', () => {
    const a = e.call({action: 'saveDocument', ...doc({company: 'Beta', docNo: ''})});
    assert.strictEqual(a.docNo, `BET-${year}-001`);
    const b = e.call({action: 'saveDocument', ...doc({company: 'Beta', docNo: ''})});
    assert.strictEqual(b.docNo, `BET-${year}-002`);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'invoice', company: 'Beta'}).docNo, `BET-${year}-003`);
  });
  test('two companies sharing a format share one sequence (no duplicates)', () => {
    const e2 = makeEnv();
    e2.call({action: 'saveCompany', company: {name: 'A'}}); e2.call({action: 'saveCompany', company: {name: 'B'}});
    const nums = ['A', 'B', 'A', 'B'].map(c => e2.call({action: 'saveDocument', ...doc({company: c})}).docNo);
    assert.strictEqual(new Set(nums).size, 4, nums.join());
  });
  test('legacy data + legacy counter are honoured (burned numbers not reused)', () => {
    const e3 = makeEnv((sheets, mk) => {
      sheets.Invoices = mk('Invoices');
      sheets.Invoices.appendRow(['DocNo']);
      const legacy = n => [`INV-${year}-00${n}`, '2026-01-01', 'Old Co', '', '', '', '', 'C', '', '', 0, 0, '', '', '', 100, '[]', new Date()];
      sheets.Invoices.appendRow(legacy(1)); sheets.Invoices.appendRow(legacy(2));
      sheets.DocCounters = mk('DocCounters');
      sheets.DocCounters.appendRow(['DocType', 'Year', 'LastNumber']);
      sheets.DocCounters.appendRow(['invoice', year, 5]);          // 3,4,5 were handed out but never saved
    });
    assert.strictEqual(e3.call({action: 'getNextNumber', docType: 'invoice'}).docNo, `INV-${year}-006`);
    assert.strictEqual(e3.sheets.Invoices.rows.length, 3, 'existing rows untouched');
    // legacy rows still load (18-col layout, no internal note column)
    const f = e3.call({action: 'fetchDocuments', docNo: 'INV-'});
    assert.strictEqual(f.rows.length, 2); assert.strictEqual(f.rows[0].date, '2026-01-01');
  });
}

// ── Internal note privacy + admin auth ──
console.log('Admin auth & privacy (strict mode: INTERNAL_NOTES_ADMIN_ONLY=1)');
{
  const e = makeEnv();
  const pwd = e.setupAdmin();
  e.props.INTERNAL_NOTES_ADMIN_ONLY = '1';
  test('password is strong, random, and not stored in plain text', () => {
    assert.ok(pwd.length >= 20 && /^[A-Za-z0-9_-]+$/.test(pwd));
    assert.ok(!JSON.stringify(e.props).includes(pwd));
    assert.notStrictEqual(pwd, makeEnv().setupAdmin());
  });
  test('wrong password / username rejected', () => {
    assert.ok(!e.call({action: 'adminLogin', username: 'admin@email.com', password: 'nope'}).ok);
    assert.ok(!e.call({action: 'adminLogin', username: 'x@email.com', password: pwd}).ok);
  });
  let token;
  test('correct login yields a session token', () => {
    const r = e.call({action: 'adminLogin', username: 'Admin@Email.com', password: pwd});
    assert.ok(r.ok && r.token.length >= 32); token = r.token;
    assert.ok(e.call({action: 'adminMe', token}).ok);
  });
  test('admin endpoints reject no/forged token', () => {
    ['adminList', 'adminGet', 'adminUpdate'].forEach(a => {
      assert.ok(!e.call({action: a}).ok); assert.ok(!e.call({action: a, token: 'a'.repeat(48)}).ok);
    });
  });
  test('internal note saved only for admin; anonymous save keeps it', () => {
    const s = e.call({action: 'saveDocument', token, ...doc({internalNote: 'supplier cost 400', paymentStatus: 'Paid'})});
    assert.ok(s.ok);
    const anon = e.call({action: 'saveDocument', ...doc({docNo: s.docNo, isNew: false, internalNote: 'HACK', paymentStatus: 'Pending', client: 'Edited'})});
    assert.ok(anon.ok);
    const row = e.sheets.Invoices.rows[1];
    assert.strictEqual(row[18], 'supplier cost 400'); assert.strictEqual(row[19], 'Paid'); assert.strictEqual(row[7], 'Edited');
  });
  test('public fetch never returns internal note / status; admin fetch does', () => {
    const pub = e.call({action: 'fetchDocuments', docNo: 'INV-'});
    assert.ok(pub.rows.length === 1 && !('internalNote' in pub.rows[0]) && !('paymentStatus' in pub.rows[0]));
    assert.ok(!JSON.stringify(pub).includes('supplier cost'));
    const adm = e.call({action: 'fetchDocuments', docNo: 'INV-', token});
    assert.strictEqual(adm.rows[0].internalNote, 'supplier cost 400');
    const stale = e.call({action: 'fetchDocuments', docNo: 'INV-', token: 'b'.repeat(48)});
    assert.ok(!('internalNote' in stale.rows[0]));
  });
  test('brute force lockout after 5 failures', () => {
    const e2 = makeEnv(); const p = e2.setupAdmin();
    for (let i = 0; i < 5; i++) e2.call({action: 'adminLogin', username: 'admin@email.com', password: 'bad' + i});
    const r = e2.call({action: 'adminLogin', username: 'admin@email.com', password: p});
    assert.ok(!r.ok && /Too many/.test(r.error));
  });
  test('logout invalidates the token; re-running setupAdmin invalidates all sessions', () => {
    const e2 = makeEnv(); const p = e2.setupAdmin();
    const t = e2.call({action: 'adminLogin', username: 'admin@email.com', password: p}).token;
    e2.call({action: 'adminLogout', token: t});
    assert.ok(!e2.call({action: 'adminMe', token: t}).ok);
    const t2 = e2.call({action: 'adminLogin', username: 'admin@email.com', password: p}).token;
    const p2 = e2.setupAdmin();
    assert.ok(!e2.call({action: 'adminMe', token: t2}).ok);
    assert.ok(!e2.call({action: 'adminLogin', username: 'admin@email.com', password: p}).ok);
    assert.ok(e2.call({action: 'adminLogin', username: 'admin@email.com', password: p2}).ok);
  });
  test('login before setupAdmin gives a clear error', () => {
    const r = makeEnv().call({action: 'adminLogin', username: 'admin@email.com', password: 'x'});
    assert.ok(!r.ok && /setup/i.test(r.error));
  });

  console.log('Admin dashboard API');
  // dataset
  const mk = (company, type, status, amt, client, note) => e.call({action: 'saveDocument', token, ...doc({docType: type, company, paymentStatus: status, client, internalNote: note || '', items: [{amount: amt}], tax: 0})});
  mk('Acme', 'invoice', 'Paid', 500, 'Zed'); mk('Acme', 'invoice', 'Pending', 300, 'Yan', 'chase payment');
  mk('Beta', 'invoice', 'Partial', 200, 'Xi'); mk('Beta', 'quotation', 'Pending', 900, 'Wu');
  test('list totals/summary match the sheet data', () => {
    const r = e.call({action: 'adminList', token});
    const invRows = e.sheets.Invoices.rows.slice(1), qtRows = e.sheets.Quotations.rows.slice(1);
    assert.strictEqual(r.total, invRows.length + qtRows.length);
    const sumInv = invRows.reduce((s, x) => s + Number(x[15]), 0);
    assert.strictEqual(r.summary.invoiceTotal, sumInv);
    assert.strictEqual(r.summary.quotationTotal, 900);
    assert.strictEqual(r.summary.paidTotal + r.summary.outstandingTotal, sumInv);
    assert.strictEqual(r.summary.paidTotal, 1180 + 500);   // 1000 + 18% tax, plus the 500 doc
  });
}

console.log('Company management');
{
  const e = makeEnv(); const pw = e.setupAdmin();
  const token = e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token;
  const co = (o = {}) => Object.assign({name: 'Zeta Corp', phone: '+91 98765 43210', email: 'a@zeta.com', address: '1 Main St', gst: '', currency: '$', template: 'navy', invoiceFormat: 'ZT-{SEQ:4}', quotationFormat: 'ZQ-{YY}-{SEQ:3}'}, o);
  test('add company requires admin', () => assert.ok(!e.call({action: 'adminSaveCompany', company: co()}).ok));
  test('validation: required fields, email, phone, template', () => {
    for (const bad of [{name: ''}, {phone: ''}, {phone: 'abc'}, {email: 'nope'}, {address: ' '}, {template: 'zzz'}, {invoiceFormat: 'x'}])
      assert.ok(!e.call({action: 'adminSaveCompany', token, company: co(bad)}).ok, JSON.stringify(bad));
    assert.strictEqual(e.call({action: 'getCompanies'}).companies.length, 0, 'nothing saved by failed attempts');
  });
  test('add company stores everything; GST optional; duplicate name refused', () => {
    assert.ok(e.call({action: 'adminSaveCompany', token, company: co()}).ok);
    const c = e.call({action: 'getCompany', name: 'zeta corp'}).company;
    assert.deepStrictEqual([c.phone, c.email, c.currency, c.template, c.invoiceFormat, c.gst], ['+91 98765 43210', 'a@zeta.com', '$', 'navy', 'ZT-{SEQ:4}', '']);
    const dup = e.call({action: 'adminSaveCompany', token, company: co()});
    assert.ok(!dup.ok && /already exists/.test(dup.error));
    assert.strictEqual(e.call({action: 'getCompanies'}).companies.length, 1);
  });
  test('update changes details but keeps one row; adminCompanies lists next numbers', () => {
    assert.ok(e.call({action: 'adminSaveCompany', token, mode: 'update', company: co({phone: '12345678', template: 'govt'})}).ok);
    assert.strictEqual(e.call({action: 'getCompanies'}).companies.length, 1);
    const l = e.call({action: 'adminCompanies', token});
    assert.strictEqual(l.companies[0].template, 'govt'); assert.strictEqual(l.companies[0].nextInvoiceNo, 'ZT-0001');
    assert.ok(!e.call({action: 'adminCompanies'}).ok);
  });
  test('company scope: restricted admin only sees/edits its companies', () => {
    e.props.ADMIN_COMPANIES = 'Zeta Corp';
    const t2 = e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token;
    assert.ok(!e.call({action: 'adminSaveCompany', token: t2, company: co({name: 'Other'})}).ok);
    assert.strictEqual(e.call({action: 'adminCompanies', token: t2}).companies.length, 1);
    e.props.ADMIN_COMPANIES = '*';
  });
}

console.log('3-company comparison');
{
  const e = makeEnv(); const pw = e.setupAdmin();
  e.props.INTERNAL_NOTES_ADMIN_ONLY = '1';
  const token = e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token;
  ['A Co', 'B Co', 'C Co'].forEach((n, i) => e.call({action: 'adminSaveCompany', token, company: {name: n, phone: '123456', email: 'x@y.com', address: 'addr', currency: ['₹', '$', '€'][i], template: 'corporate', quotationFormat: n[0] + 'Q-{SEQ:3}'}}));
  const mk = (company, rates, extra = {}) => ({company, phone: '123456', isNew: true, currency: company === 'B Co' ? '$' : '₹',
    items: [{name: 'DVR', q: 2, rate: rates[0], c: 1000}, {name: 'Cam', q: 4, rate: rates[1], c: 500}], ...extra});
  const common = {client: 'Acme Ltd', date: '2026-10-17', tax: 18, notes: 'pub', internalNote: 'margin 12%'};
  let cmp;
  test('saves 3 quotations in one go, numbered per company, totals computed server-side', () => {
    cmp = e.call({action: 'saveComparison', token, common, quotes: [mk('A Co', [100, 50]), mk('B Co', [110.5, 60]), mk('C Co', [150, 75])]});
    assert.ok(cmp.ok, cmp.error);
    assert.deepStrictEqual(cmp.results.map(r => r.docNo), ['AQ-001', 'BQ-001', 'CQ-001']);
    // whole numbers: B's 110.5 rate rounds to 111 -> 222+240=462, tax 83 -> 545
    assert.deepStrictEqual(cmp.results.map(r => r.total), [472, 545, 708]);
    assert.strictEqual(e.sheets.Quotations.rows.length, 4);
    assert.ok(e.sheets.Quotations.rows.slice(1).every(r => r[20] === cmp.comparisonId && r[18] === 'margin 12%'));
  });
  test('item catalog records only Company 1 (base) rates', () => {
    const cat = e.sheets.ItemCatalog.rows.slice(1);
    assert.deepStrictEqual(cat.map(r => [r[0], r[3]]), [['DVR', 100], ['Cam', 50]]);
  });
  test('re-saving the same comparison updates in place — no duplicates, same numbers', () => {
    const again = e.call({action: 'saveComparison', token, comparisonId: cmp.comparisonId, common, quotes: cmp.results.map(r => mk(r.company, [100, 50], {docNo: r.docNo, isNew: false}))});
    assert.ok(again.ok); assert.deepStrictEqual(again.results.map(r => r.docNo), ['AQ-001', 'BQ-001', 'CQ-001']);
    assert.strictEqual(e.sheets.Quotations.rows.length, 4);
  });
  test('validation: same company twice / different product lists / <2 quotes rejected, nothing written', () => {
    const n = e.sheets.Quotations.rows.length;
    assert.ok(!e.call({action: 'saveComparison', token, common, quotes: [mk('A Co', [1, 1]), mk('a co', [2, 2])]}).ok);
    const odd = mk('B Co', [1, 1]); odd.items[1].q = 9;
    assert.ok(!e.call({action: 'saveComparison', token, common, quotes: [mk('A Co', [1, 1]), odd]}).ok);
    assert.ok(!e.call({action: 'saveComparison', token, common, quotes: [mk('A Co', [1, 1])]}).ok);
    assert.strictEqual(e.sheets.Quotations.rows.length, n);
  });
  test('getComparison: public view hides internal note, admin sees it; lookup by any doc no', () => {
    const pub = e.call({action: 'getComparison', docNo: 'BQ-001'});
    assert.ok(pub.ok && pub.quotes.length === 3 && pub.quotes.every(q => !('internalNote' in q)));
    const adm = e.call({action: 'getComparison', comparisonId: cmp.comparisonId, token});
    assert.ok(adm.quotes.every(q => q.internalNote === 'margin 12%'));
    assert.ok(!e.call({action: 'getComparison', docNo: 'nope'}).ok);
  });
  test('admin dashboard: comparison siblings + per-doc currency, totals match sheet', () => {
    const g = e.call({action: 'adminGet', token, docType: 'quotation', docNo: 'BQ-001'});
    assert.strictEqual(g.siblings.length, 3); assert.strictEqual(g.doc.currency, '$');
    const l = e.call({action: 'adminList', token});
    assert.strictEqual(l.summary.quotationTotal, e.sheets.Quotations.rows.slice(1).reduce((s, r) => s + Number(r[15]), 0));
    assert.deepStrictEqual(l.summary.currencies.sort(), ['$', '₹']);
    assert.ok(l.rows.every(r => r.comparisonId === cmp.comparisonId));
  });
}

console.log('Internal view (default: visible to staff form, still never public-printed)');
{
  const e = makeEnv();
  test('default: note + status saved and fetched without admin login; ping reports mode', () => {
    assert.strictEqual(e.call({action: 'ping'}).notesAdminOnly, false);
    const s = e.call({action: 'saveDocument', docType: 'invoice', company: 'Staff Co', client: 'c', items: [{name: 'a', q: 2, c: 500, rate: 600}], tax: 0, internalNote: 'visible to staff', paymentStatus: 'Paid', isNew: true});
    assert.ok(s.ok);
    const f = e.call({action: 'fetchDocuments', docNo: s.docNo}).rows[0];
    assert.strictEqual(f.internalNote, 'visible to staff'); assert.strictEqual(f.paymentStatus, 'Paid');
    assert.strictEqual(f.items[0].profit, 200);
    // update keeps/changes the note
    e.call({action: 'saveDocument', docType: 'invoice', docNo: s.docNo, isNew: false, company: 'Staff Co', client: 'c', items: [{name: 'a', q: 2, c: 500, rate: 600}], tax: 0, internalNote: 'changed'});
    assert.strictEqual(e.call({action: 'fetchDocuments', docNo: s.docNo}).rows[0].internalNote, 'changed');
  });
  test('strict switch turns admin-only back on (note hidden + not writable without login)', () => {
    e.props.INTERNAL_NOTES_ADMIN_ONLY = '1';
    assert.strictEqual(e.call({action: 'ping'}).notesAdminOnly, true);
    const no = e.sheets.Invoices.rows[1][0];
    assert.ok(!('internalNote' in e.call({action: 'fetchDocuments', docNo: no}).rows[0]));
    e.call({action: 'saveDocument', docType: 'invoice', docNo: no, isNew: false, company: 'Staff Co', client: 'c', items: [{name: 'a', q: 1, c: 1, rate: 2}], tax: 0, internalNote: 'HACK'});
    assert.strictEqual(e.sheets.Invoices.rows[1][18], 'changed');
    e.props.INTERNAL_NOTES_ADMIN_ONLY = '0';
  });
}

console.log('Whole-number money rules');
{
  const e = makeEnv(); const pw = e.setupAdmin();
  const token = e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token;
  e.call({action: 'adminSaveCompany', token, company: {name: 'Comm Co', phone: '123456', email: 'a@b.com', address: 'x', currency: '₹', template: 'corporate'}});
  test('rounding rule: half-up, symmetric, no float surprises', () => {
    assert.deepStrictEqual([0.5, 1.5, 2.5, 1.005, 2.675, -0.5, -2.5, 1e-10].map(QMCalc.rnd), [1, 2, 3, 1, 3, -1, -3, 0]);
  });
  test('backend totals == qm-calc.js on 500 random documents (incl. discounts)', () => {
    let seed = 12345; const rand = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
    for (let n = 0; n < 500; n++) {
      const items = Array.from({length: 1 + Math.floor(rand() * 5)}, (_, i) => ({name: 'it' + i, q: Math.floor(rand() * 9) + 1, rate: Math.round(rand() * 500000) / 100}));
      const tax = [0, 5, 12, 18, 28, 7.5][Math.floor(rand() * 6)], discType = rand() < 0.5 ? 'percent' : 'amount';
      const discount = discType === 'percent' ? Math.round(rand() * 2500) / 100 : Math.round(rand() * 100000) / 100;
      const r = e.call({action: 'saveDocument', token, docType: 'invoice', company: 'Comm Co', client: 'c', items, tax, discount, discountType: discType, isNew: true, date: '2026-10-17'});
      const exp = QMCalc.totals(items.map(x => ({...x, rate: QMCalc.rnd(x.rate)})), tax, discount, discType);
      assert.strictEqual(r.total, exp.total, JSON.stringify({items, tax, discount, discType}));
      assert.ok(Number.isInteger(r.total));
    }
  });
  test('stored items & totals are whole numbers; discount never exceeds subtotal', () => {
    const rows = e.sheets.Invoices.rows.slice(1);
    assert.ok(rows.every(r => Number.isInteger(Number(r[15])) && JSON.parse(r[16]).every(i => Number.isInteger(i.rate) && Number.isInteger(i.amount))));
    const r = e.call({action: 'saveDocument', token, docType: 'invoice', company: 'Comm Co', client: 'c', items: [{name: 'a', q: 1, rate: 100}], tax: 18, discount: 500, discountType: 'amount', isNew: true});
    assert.strictEqual(r.total, 0);
  });
  test('worked example: 3 x 33.33 + 10.4 line, 10% discount, 18% tax', () => {
    const r = e.call({action: 'saveDocument', token, docType: 'quotation', company: 'Comm Co', client: 'c', tax: 18, discount: 10, discountType: 'percent', isNew: true,
      items: [{name: 'a', q: 3, rate: 33.33}, {name: 'b', q: 1, rate: 10.4}]});
    // rates -> 33, 10; amounts 99 + 10 = 109; discount 11 (10.9); taxable 98; tax 18 (17.64); total 116
    assert.strictEqual(r.total, 116);
  });
  test('profit rule: unit = rate - cost, line = amount - qty x cost, rate from cost + profit%', () => {
    assert.strictEqual(QMCalc.rateFromCost(500, 20), 600);               // the spec example
    const p = QMCalc.profit([{name: 'a', q: 2, c: 500, rate: 600}, {name: 'b', q: 3, c: 99.5, rate: 120}]);
    assert.deepStrictEqual(p.rows.map(r => [r.unit, r.line]), [[100, 200], [20, 61]]);   // b: 120-100=20/unit; 360-299(298.5) = 61
    assert.strictEqual(p.total, 261);
    // backend stores the same figures
    const r = e.call({action: 'saveDocument', token, docType: 'invoice', company: 'Comm Co', client: 'c', items: [{name: 'a', q: 2, c: 500, rate: 600}], tax: 0, isNew: true});
    const row = e.sheets.Invoices.rows[e.sheets.Invoices.rows.length - 1];
    assert.strictEqual(r.total, 1200); assert.strictEqual(JSON.parse(row[16])[0].profit, 200);
  });
  test('commission feature is gone: no commission in company or document APIs', () => {
    const all = JSON.stringify([e.call({action: 'getCompanies', token}), e.call({action: 'adminCompanies', token}), e.call({action: 'fetchDocuments', docNo: 'INV'})]);
    assert.ok(!/ommission/.test(all));
  });
}

console.log('Admin edit of existing documents');
{
  const e = makeEnv(); const pw = e.setupAdmin();
  const token = e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token;
  const base = e.call({action: 'saveDocument', token, docType: 'invoice', company: 'Edit Co', phone: '111', email: 'e@e.com', address: 'Addr', gst: 'G1', client: 'Orig', clientMobile: '999', validity: '15 Days', tax: 18, notes: 'n', internalNote: 'secret', items: [{name: 'A', q: 2, c: 50, rate: 100}, {name: 'B', q: 1, c: 0, rate: 400}], isNew: true, date: '2026-10-17', currency: '$'});
  const no = base.docNo;
  const edit = f => e.call({action: 'adminSaveDocument', token, docType: 'invoice', docNo: no, fields: f});
  test('requires admin; unknown doc rejected', () => {
    assert.ok(!e.call({action: 'adminSaveDocument', docType: 'invoice', docNo: no, fields: {items: [{name: 'x', q: 1, rate: 1}]}}).ok);
    assert.ok(!e.call({action: 'adminSaveDocument', token, docType: 'invoice', docNo: 'NOPE', fields: {items: [{name: 'x', q: 1, rate: 1}]}}).ok);
  });
  test('edit items/qty/rate/discount/tax/client → total recomputed & persisted, identity fields untouched', () => {
    const r = edit({client: 'Changed', clientMobile: '888', tax: 12, discount: 50, discountType: 'amount', notes: 'new note', validity: '30 Days',
      items: [{name: 'A', q: 3, rate: 150.4}, {name: 'B', q: 1, rate: 400}, {name: 'C', q: 2, rate: 25}]});
    assert.ok(r.ok, r.error);
    // 3*150=450, +400, +50 = 900; discount 50 -> 850; tax 12% = 102 -> 952
    assert.strictEqual(r.total, 952);
    assert.strictEqual(e.sheets.Invoices.rows.length, 2, 'edited in place, no duplicate row');
    const row = e.sheets.Invoices.rows[1];
    assert.deepStrictEqual([row[0], row[2], row[7], row[8], row[9], row[10], row[14], Number(row[15]), row[18], row[21], row[22], row[23]],
      [no, 'Edit Co', 'Changed', '888', '30 Days', 12, 'new note', 952, 'secret', '$', 50, 'amount']);
    assert.strictEqual(JSON.parse(row[16]).length, 3);
    const f = e.call({action: 'fetchDocuments', docNo: no, token}).rows[0];
    assert.strictEqual(f.total, 952); assert.strictEqual(f.discount, 50); assert.strictEqual(f.discountType, 'amount'); assert.strictEqual(f.internalNote, 'secret');
    const l = e.call({action: 'adminList', token}); assert.strictEqual(l.summary.invoiceTotal, 952);
  });
  test('omitted fields are preserved (discount/type stay as stored)', () => {
    assert.ok(edit({items: [{name: 'A', q: 10, rate: 100}]}).ok);   // no discount given: keeps 50 flat
    const f = e.call({action: 'adminGet', token, docType: 'invoice', docNo: no}).doc;
    assert.strictEqual(f.discount, 50); assert.strictEqual(f.discountType, 'amount'); assert.strictEqual(f.total, 1064);  // 1000-50=950, +12% = 114 -> 1064
  });
  test('item catalog is not altered by admin edits', () => {
    assert.ok(e.sheets.ItemCatalog.rows.slice(1).every(r => r[0] !== 'C'));
  });
  test('internal note + payment status can be edited in the same call', () => {
    assert.ok(edit({internalNote: 'updated secret', paymentStatus: 'Paid', discount: 0, items: [{name: 'A', q: 1, rate: 10}]}).ok);
    const f = e.call({action: 'adminGet', token, docType: 'invoice', docNo: no}).doc;
    assert.strictEqual(f.internalNote, 'updated secret'); assert.strictEqual(f.paymentStatus, 'Paid'); assert.strictEqual(f.total, 11);   // 10 + 12% tax (kept from the previous edit) = 11.2 -> 11
  });
  test('validation: no items, blank name, negative, bad tax/discount/status', () => {
    for (const bad of [{items: []}, {items: [{name: ' ', q: 1, rate: 1}]}, {items: [{name: 'x', q: -1, rate: 1}]}, {items: [{name: 'x', q: 1, rate: -5}]},
      {tax: 120, items: [{name: 'x', q: 1, rate: 1}]}, {discount: 150, discountType: 'percent', items: [{name: 'x', q: 1, rate: 1}]}, {paymentStatus: 'Weird', items: [{name: 'x', q: 1, rate: 1}]}])
      assert.ok(!edit(bad).ok, JSON.stringify(bad));
  });
  test('company-scoped admin cannot edit other companies’ documents', () => {
    e.props.ADMIN_COMPANIES = 'Someone Else';
    const t2 = e.call({action: 'adminLogin', username: 'admin@email.com', password: pw}).token;
    assert.ok(!e.call({action: 'adminSaveDocument', token: t2, docType: 'invoice', docNo: no, fields: {items: [{name: 'x', q: 1, rate: 1}]}}).ok);
    e.props.ADMIN_COMPANIES = '*';
  });
}
console.log(`\n${passed} tests passed`);
