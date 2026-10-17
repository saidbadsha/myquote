/*
 Shared money rules for index.html, question_maker.html and admin.html.
 The Apps Script backend mirrors these functions (tests/backend.test.js checks both agree).

 Rule: every amount is a WHOLE NUMBER, rounded half-up (symmetric for negatives), and each
 level is rounded before it is used by the next, so what is displayed always adds up:
   line amount = round(qty x rate)         (or the amount typed directly, rounded)
   subtotal    = sum of line amounts
   discount    = round(subtotal x d / 100)  or flat round(d), never more than the subtotal
   tax         = round((subtotal - discount) x tax% / 100)
   total       = subtotal - discount + tax
   profit/unit = rate - cost,  line profit = amount - round(qty x cost)   (internal only, never printed)
   markup rate = round(cost + cost x profit% / 100)
*/
(function (root) {
  function rnd(x) {
    const n = Number(x);
    if (!isFinite(n)) return 0;
    return n < 0 ? -Math.round(-n + 1e-9) : Math.round(n + 1e-9);
  }
  function lineAmount(item) {
    const a = item.amount;
    return (a !== undefined && a !== null && a !== '') ? rnd(a) : rnd((Number(item.q) || 0) * rnd(item.rate));
  }
  function totals(items, taxPct, discount, discountType) {
    const subtotal = (items || []).reduce((s, x) => s + lineAmount(x), 0);
    let disc = discountType === 'amount' ? rnd(discount) : rnd(subtotal * (Number(discount) || 0) / 100);
    disc = Math.min(Math.max(disc, 0), subtotal);
    const taxable = subtotal - disc;
    const tax = rnd(taxable * (Number(taxPct) || 0) / 100);
    return { subtotal: subtotal, discount: disc, taxable: taxable, tax: tax, total: taxable + tax };
  }
  // Rate from cost and profit %: cost + (cost x profit% / 100), whole number.
  function rateFromCost(cost, profitPct) { const c = Number(cost) || 0; return rnd(c + c * (Number(profitPct) || 0) / 100); }
  // Per-item and total profit for the internal view.
  function profit(items) {
    const rows = (items || []).map(function (x) {
      const q = Number(x.q) || 0, c = Number(x.c) || 0;
      const rate = rnd(x.rate), amount = lineAmount(x);
      return { name: x.name, q: q, cost: rnd(c), rate: rate, unit: rate - rnd(c), line: amount - rnd(q * c), amount: amount };
    });
    return { rows: rows, total: rows.reduce(function (s, r) { return s + r.line; }, 0) };
  }
  function fmt(n, cur) { return (cur || '₹') + ' ' + rnd(n).toLocaleString('en-IN', { maximumFractionDigits: 0 }); }

  const api = { rnd: rnd, lineAmount: lineAmount, totals: totals, rateFromCost: rateFromCost, profit: profit, fmt: fmt };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.QMCalc = api;
})(typeof window !== 'undefined' ? window : globalThis);
