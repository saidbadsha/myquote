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
    assert.ok(e.call({action: 'saveCompany', company: {name: 'Acme', invoiceFormat: 'ACM/{YY}{MM}/{SEQ:4}', quotationFormat: 'AQ-{SEQ}'}}).ok);
    assert.ok(e.call({action: 'saveCompany', company: {name: 'Beta', invoiceFormat: 'BET-{YYYY}-{SEQ:3}'}}).ok);
    const yy = String(year).slice(-2), mm = ('0' + (new Date().getMonth() + 1)).slice(-2);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'invoice', company: 'Acme'}).docNo, `ACM/${yy}${mm}/0001`);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'invoice', company: 'Beta'}).docNo, `BET-${year}-001`);
    assert.strictEqual(e.call({action: 'getNextNumber', docType: 'quotation', company: 'Acme'}).docNo, 'AQ-1');
    const bad = ['no-seq', 'A-{SEQ}-{SEQ}', 'A-{SEQ:9}', 'A<{SEQ}>', '$$'].map(f => e.call({action: 'saveCompany', company: {name: 'Acme', invoiceFormat: f}}));
    bad.forEach(r => assert.ok(!r.ok, 'should reject'));
    assert.strictEqual(e.call({action: 'getCompanies'}).companies.length, 2);
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
console.log(`\n${passed} tests passed`);
