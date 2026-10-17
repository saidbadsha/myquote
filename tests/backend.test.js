// Run: node tests/backend.test.js
const assert = require('assert');
const { makeEnv } = require('./mock-env');

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
console.log('Admin auth & privacy');
{
  const e = makeEnv();
  const pwd = e.setupAdmin();
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
    assert.deepStrictEqual(cmp.results.map(r => r.total), [(200 + 200) * 1.18, (221 + 240) * 1.18, (300 + 300) * 1.18].map(x => Math.round(x * 100) / 100));
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
console.log(`\n${passed} tests passed`);
