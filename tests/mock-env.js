// In-memory Google Apps Script environment used to run google_sheet_apps_script_v2.gs under Node.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

function makeEnv(seed) {
  const sheets = {};
  const mkSheet = name => {
    const rows = []; // array of arrays
    const sh = {
      name, rows,
      getLastRow: () => rows.length,
      getMaxRows: () => Math.max(rows.length, 1000),
      getMaxColumns: () => 26,
      insertColumnsAfter() {},
      appendRow(r) { rows.push(r.map(conv)); },
      getRange(r, c, nr = 1, nc = 1) {
        return {
          getValues: () => Array.from({length: nr}, (_, i) => Array.from({length: nc}, (_, j) => (rows[r - 1 + i] || [])[c - 1 + j] ?? '')),
          getValue: () => (rows[r - 1] || [])[c - 1] ?? '',
          setValues(v) { v.forEach((row, i) => row.forEach((x, j) => { (rows[r - 1 + i] = rows[r - 1 + i] || [])[c - 1 + j] = conv(x); })); },
          setValue(x) { (rows[r - 1] = rows[r - 1] || [])[c - 1] = conv(x); },
          setNumberFormat() {}
        };
      }
    };
    return sh;
  };
  // Sheets turns "2026-10-17" into a Date; mimic that so fmtDate_ is exercised.
  const conv = x => (typeof x === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(x)) ? new Date(x + 'T00:00:00') : x;

  const cache = {}, props = {};
  const ctx = {
    console, Logger: {log() {}}, Date, JSON, Math, Number, String, Array, Set, Object, RegExp,
    SpreadsheetApp: {getActiveSpreadsheet: () => ({
      getSheetByName: n => sheets[n] || null,
      insertSheet: n => (sheets[n] = mkSheet(n))
    })},
    CacheService: {getScriptCache: () => ({
      get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; }
    })},
    PropertiesService: {getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = v; },
      setProperties: o => Object.assign(props, o)
    })},
    LockService: {getScriptLock: () => ({waitLock() {}, releaseLock() {}})},
    ContentService: {createTextOutput: s => ({s, setMimeType() { return this; }}), MimeType: {JSON: 'json'}},
    Session: {getScriptTimeZone: () => 'UTC'},
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      newBlob: s => ({getBytes: () => Array.from(Buffer.from(s)).map(b => (b << 24) >> 24)}),
      computeHmacSha256Signature: (v, k) => {
        const b = x => Buffer.from(typeof x === 'string' ? Buffer.from(x) : x.map(n => n & 255));
        return Array.from(crypto.createHmac('sha256', b(k)).update(b(v)).digest()).map(n => (n << 24) >> 24);
      },
      computeDigest: (_, s) => Array.from(crypto.createHash('sha256').update(s).digest()).map(n => (n << 24) >> 24),
      DigestAlgorithm: {SHA_256: 'sha256'},
      formatDate: (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
    }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'google_sheet_apps_script_v2.gs'), 'utf8')
    + '\n;this.__api = {doPost, setupAdmin, sheets: () => null};', ctx);
  if (seed) seed(sheets, mkSheet);
  const call = (body) => JSON.parse(ctx.__api.doPost({postData: {contents: JSON.stringify(body)}}).s);
  return {ctx, sheets, call, cache, props, setupAdmin: ctx.__api.setupAdmin};
}

module.exports = { makeEnv };
