/*!
 * ITR Generator — Tax engine for BIR Form 1701Q and 1702Q
 * Quarterly returns are cumulative: each quarter's tax due is computed on the
 * taxable income to date, then prior-quarter payments and credits are deducted.
 *
 * 1701Q (individuals)  : graduated rates (itemized or 40% OSD) or 8% income tax option
 * 1702Q (corporations) : RCIT (25% / 20% / custom) vs MCIT (2% of gross income), whichever is higher
 */
(function (root) {
  'use strict';

  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const pos = (n) => Math.max(0, Number(n) || 0);

  // [upper bound, base tax, rate on excess over previous bound]
  const BRACKETS_2018 = [
    [250000, 0, 0], [400000, 0, 0.20], [800000, 30000, 0.25],
    [2000000, 130000, 0.30], [8000000, 490000, 0.32], [Infinity, 2410000, 0.35]
  ];
  // TRAIN Law schedule effective January 1, 2023 onwards
  const BRACKETS_2023 = [
    [250000, 0, 0], [400000, 0, 0.15], [800000, 22500, 0.20],
    [2000000, 102500, 0.25], [8000000, 402500, 0.30], [Infinity, 2202500, 0.35]
  ];

  function graduatedTax(income, year) {
    const table = Number(year) >= 2023 ? BRACKETS_2023 : BRACKETS_2018;
    const x = Number(income) || 0;
    if (x <= 0) return 0;
    let lower = 0;
    for (const [upper, base, rate] of table) {
      if (x <= upper) return r2(base + (x - lower) * rate);
      lower = upper;
    }
    return 0;
  }

  /** Representation & entertainment is deductible only up to 0.5% of net sales (goods) or 1% of net revenue (services). RR 10-2002 */
  function entertainmentCap(accounts, revenue, sellerOf) {
    const rate = sellerOf === 'Services' ? 0.01 : 0.005;
    const cap = r2(pos(revenue) * rate);
    const re = accounts.filter(a => a.category === 'operating_expense' && /represent|entertain/i.test(a.account));
    const reTotal = r2(re.reduce((s, a) => s + (Number(a.amount) || 0), 0));
    const disallowed = r2(Math.max(0, reTotal - cap));
    return { cap, rate, reTotal, disallowed };
  }

  function credits(i) {
    const p = i.prior || {};
    const items = [
      { key: 'priorYearExcess', label: "Prior Year's Excess Credits", amount: r2(p.priorYearExcess) },
      { key: 'taxPaymentsPrev', label: 'Tax Payment/s for the Previous Quarter/s', amount: r2(p.taxPaymentsPrev) },
      { key: 'cwtPrev', label: 'Creditable Tax Withheld for the Previous Quarter/s', amount: r2(p.cwtPrev) },
      { key: 'cwtThisQ', label: 'Creditable Tax Withheld per BIR Form 2307 for this Quarter', amount: r2(i.cwtThisQ) },
      { key: 'otherCredits', label: 'Other Tax Credits/Payments', amount: r2(i.otherCredits) }
    ];
    const total = r2(items.reduce((s, x) => s + x.amount, 0));
    return { items, total };
  }

  function compute1701Q(i) {
    const lines = [], warnings = [];
    const year = Number(i.year), q = Number(i.quarter);
    const rev = r2(i.revenue), nonop = r2(i.otherIncome);
    const p = i.prior || {};
    let taxDue = 0, taxableToDate = 0, schedule;

    if (i.method === 'eight') {
      schedule = 'Schedule II — 8% Income Tax Rate';
      if (i.vatStatus === 'VAT') warnings.push('The 8% option is only available to non-VAT taxpayers whose gross sales/receipts do not exceed ₱3,000,000.');
      const totalQ = r2(rev + nonop);
      const prev = r2(p.cumulativeIncomePrev);
      const cum = r2(totalQ + prev);
      const reduction = i.mixedIncome ? 0 : 250000;
      taxableToDate = r2(pos(cum - reduction));
      taxDue = r2(taxableToDate * 0.08);
      if (cum > 3000000) warnings.push('Cumulative gross sales/receipts now exceed ₱3,000,000. The 8% option no longer applies; switch to graduated rates.');
      lines.push(
        { label: 'Sales/Revenues/Receipts/Fees', amount: rev },
        { label: 'Add: Non-Operating Income', amount: nonop, editable: 'otherIncome' },
        { label: 'Total Income for the Quarter', amount: totalQ, type: 'subtotal' },
        { label: 'Add: Total Taxable Income/(Loss) Previous Quarter/s', amount: prev, editable: 'cumulativeIncomePrev' },
        { label: 'Cumulative Taxable Income/(Loss) as of this Quarter', amount: cum, type: 'subtotal' },
        { label: i.mixedIncome ? 'Less: Allowable reduction (not applicable — mixed income earner)' : 'Less: Allowable reduction from gross sales/receipts', amount: reduction },
        { label: 'Taxable Income/(Loss) To Date', amount: taxableToDate, type: 'subtotal' },
        { label: 'Tax Due (8%)', amount: taxDue, type: 'total' }
      );
    } else {
      const itemized = i.method !== 'osd';
      schedule = 'Schedule I — Graduated Income Tax Rates (' + (itemized ? 'Itemized Deductions' : 'Optional Standard Deduction') + ')';
      const cos = itemized ? r2(i.costOfSales) : 0;
      const gross = r2(rev - cos);
      const ded = itemized ? r2(i.opex) : r2(rev * 0.40);
      const netQ = r2(gross - ded);
      const prev = r2(p.taxableIncomePrev);
      taxableToDate = r2(netQ + nonop + prev);
      taxDue = graduatedTax(taxableToDate, year);
      lines.push(
        { label: 'Sales/Revenues/Receipts/Fees (net of returns, allowances and discounts)', amount: rev },
        { label: 'Less: Cost of Sales/Services', amount: cos, note: itemized ? '' : 'Not deducted under OSD' },
        { label: 'Gross Income/(Loss) from Operation', amount: gross, type: 'subtotal' },
        itemized
          ? { label: 'Less: Total Allowable Itemized Deductions', amount: ded }
          : { label: 'Less: Optional Standard Deduction (40% of Sales/Receipts)', amount: ded },
        { label: 'Net Income/(Loss) this Quarter', amount: netQ, type: 'subtotal' },
        { label: 'Add: Taxable Income/(Loss) Previous Quarter/s', amount: prev, editable: 'taxableIncomePrev' },
        { label: 'Non-Operating Income', amount: nonop, editable: 'otherIncome' },
        { label: 'Total Taxable Income/(Loss) To Date', amount: taxableToDate, type: 'subtotal' },
        { label: 'Tax Due (graduated rates, ' + (year >= 2023 ? '2023 onwards' : '2018–2022') + ' table)', amount: taxDue, type: 'total' }
      );
    }

    const c = credits(i);
    return finish('1701Q', schedule, lines, c, taxDue, taxableToDate, warnings, q);
  }

  function compute1702Q(i) {
    const lines = [], warnings = [];
    const q = Number(i.quarter);
    const rate = Number(i.rcitRate) || 0.25;
    const mcitRate = Number(i.mcitRate) || 0.02;
    const p = i.prior || {};
    const itemized = i.method !== 'osd';

    const rev = r2(i.revenue);
    const cos = r2(i.costOfSales);
    const grossOps = r2(rev - cos);
    const other = r2(i.otherIncome);
    const totalGross = r2(grossOps + other);
    const ded = itemized ? r2(i.opex) : r2(pos(totalGross) * 0.40);
    const taxableQ = r2(totalGross - ded);
    const prev = r2(p.taxableIncomePrev);
    const toDate = r2(taxableQ + prev);
    const rcit = r2(pos(toDate) * rate);

    const mcitBase = r2(totalGross + r2(p.grossIncomePrev));
    const mcit = i.mcitApplies ? r2(pos(mcitBase) * mcitRate) : 0;
    const taxDue = r2(Math.max(rcit, mcit));
    if (i.mcitApplies && mcit > rcit) warnings.push('MCIT is higher than the regular income tax, so MCIT is the tax due this quarter.');
    if (rate === 0.20) warnings.push('The 20% rate applies only if net taxable income is ₱5,000,000 or less AND total assets (excluding land) are ₱100,000,000 or less.');

    lines.push(
      { label: 'Sales/Receipts/Revenues/Fees', amount: rev },
      { label: 'Less: Cost of Sales/Services', amount: cos },
      { label: 'Gross Income from Operation', amount: grossOps, type: 'subtotal' },
      { label: 'Add: Other Taxable Income Not Subjected to Final Tax', amount: other, editable: 'otherIncome' },
      { label: 'Total Gross Income', amount: totalGross, type: 'subtotal' },
      itemized
        ? { label: 'Less: Deductions (Itemized)', amount: ded }
        : { label: 'Less: Optional Standard Deduction (40% of Gross Income)', amount: ded },
      { label: 'Taxable Income this Quarter', amount: taxableQ, type: 'subtotal' },
      { label: 'Add: Taxable Income Previous Quarter/s', amount: prev, editable: 'taxableIncomePrev' },
      { label: 'Total Taxable Income To Date', amount: toDate, type: 'subtotal' },
      { label: 'Applicable Income Tax Rate', amount: rate * 100, type: 'rate' },
      { label: 'Income Tax Due other than MCIT', amount: rcit },
      { label: 'Gross Income Previous Quarter/s (for MCIT)', amount: r2(p.grossIncomePrev), editable: 'grossIncomePrev' },
      { label: 'Cumulative Gross Income (MCIT base)', amount: mcitBase },
      { label: 'Minimum Corporate Income Tax (' + (mcitRate * 100) + '%)' + (i.mcitApplies ? '' : ' — not yet applicable'), amount: mcit },
      { label: 'Tax Due (higher of RCIT and MCIT)', amount: taxDue, type: 'total' }
    );

    const c = credits(i);
    return finish('1702Q', 'Schedule I — Regular/Normal Rate' + (itemized ? ' (Itemized Deductions)' : ' (OSD)'), lines, c, taxDue, toDate, warnings, q);
  }

  function finish(form, schedule, lines, c, taxDue, taxableToDate, warnings, q) {
    const payable = r2(taxDue - c.total);
    if (payable < 0) warnings.push('Tax credits exceed the tax due — this quarter shows an overpayment of ₱' + Math.abs(payable).toLocaleString('en-PH', { minimumFractionDigits: 2 }) + '.');
    if (q === 1 && (c.items[1].amount || c.items[2].amount)) warnings.push('Q1 returns normally have no previous-quarter payments or withholding.');
    return {
      form, schedule, lines,
      credits: c.items, totalCredits: c.total,
      taxDue: r2(taxDue), taxableToDate: r2(taxableToDate),
      taxPayable: payable, warnings
    };
  }

  function compute(input) {
    return input.form === '1702Q' ? compute1702Q(input) : compute1701Q(input);
  }

  const api = { compute, graduatedTax, entertainmentCap, r2 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ITRTax = api;
})(typeof window !== 'undefined' ? window : globalThis);
