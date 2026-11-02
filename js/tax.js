/*!
 * ITR Generator — Tax engine for BIR Form 1701Q and 1702Q
 * Quarterly returns are cumulative: each quarter's tax due is computed on the
 * taxable income to date, then prior-quarter payments and credits are deducted.
 *
 * 1701Q (individuals)  : graduated rates (itemized or 40% OSD) or 8% income tax option
 * 1702Q (corporations) : RCIT 25% / 20% vs MCIT (whichever is higher), MCIT only, IAET 10%, Non-resident 25%
 * Credits/labels follow the Sta. Rita Tax Calc: Tax Still Due = max(0, Tax Due − credits)
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
      { key: 'taxPaymentsPrev', label: 'Tax Paid in Prior Quarters', amount: r2(p.taxPaymentsPrev) },
      { key: 'cwtPrev', label: 'Creditable Withholding Tax — Previous Quarter/s', amount: r2(p.cwtPrev) },
      { key: 'cwtThisQ', label: 'Creditable Withholding Tax — This Quarter (BIR 2307)', amount: r2(i.cwtThisQ), note: i.cwtNote || '' }
    ];
    if (i.form === '1702Q') items.push({ key: 'excessMcit', label: 'Excess MCIT Credit', amount: r2(i.excessMcit) });
    items.push({ key: 'otherCredits', label: 'Other Tax Credits', amount: r2(i.otherCredits) });
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
        { label: 'Gross Income / Gross Receipts', amount: rev },
        { label: 'Add: Other Taxable Income', amount: nonop, editable: 'otherIncome' },
        { label: 'Total Income for the Quarter', amount: totalQ, type: 'subtotal' },
        { label: 'Add: Cumulative Income — Previous Quarter/s', amount: prev, editable: 'cumulativeIncomePrev' },
        { label: 'Cumulative Income as of this Quarter', amount: cum, type: 'subtotal' },
        { label: i.mixedIncome ? 'Less: ₱250,000 reduction (not applicable — mixed income earner)' : 'Less: ₱250,000 reduction', amount: reduction },
        { label: 'Net Taxable Income', amount: taxableToDate, type: 'subtotal' },
        { label: 'Income Tax Due (8%)', amount: taxDue, type: 'total' }
      );
    } else {
      // Deduction methods (Tax Calc logic: gross income = gross receipts − COGS; deduction = OSD or itemized)
      //  itemized : gross income − itemized deductions
      //  osd      : 40% of gross sales/receipts, COGS not deducted (BIR rule for individuals, RR 16-2008)
      //  osd_gi   : 40% of gross income after COGS (as computed in the office Tax Calc)
      const m = i.method || 'itemized';
      const label = { itemized: 'Itemized', osd: 'OSD 40% of gross sales', osd_gi: 'OSD 40% of gross income' }[m] || 'Itemized';
      schedule = 'Graduated Income Tax Rates — ' + label;
      const cos = m === 'osd' ? 0 : r2(i.costOfSales);
      const gross = r2(rev - cos);
      const ded = m === 'itemized' ? r2(i.opex) : m === 'osd' ? r2(rev * 0.40) : r2(pos(gross) * 0.40);
      const netQ = r2(gross - ded);
      const prev = r2(p.taxableIncomePrev);
      taxableToDate = r2(pos(netQ + nonop + prev));
      taxDue = graduatedTax(taxableToDate, year);
      lines.push(
        { label: 'Gross Income / Gross Receipts', amount: rev },
        { label: 'Less: Cost of Sales / Services (COGS)', amount: cos, note: m === 'osd' ? 'Not deducted when OSD is based on gross sales' : '' },
        { label: 'Gross Income (net)', amount: gross, type: 'subtotal' },
        m === 'itemized' ? { label: 'Less: Total Allowable Itemized Deductions (Operating Expenses)', amount: ded }
          : { label: 'Less: OSD (40% of ' + (m === 'osd' ? 'gross sales/receipts' : 'gross income') + ')', amount: ded },
        { label: 'Net Income this Quarter', amount: netQ, type: 'subtotal' },
        { label: 'Add: Other Taxable Income', amount: nonop, editable: 'otherIncome' },
        { label: 'Add: Taxable Income — Previous Quarter/s', amount: prev, editable: 'taxableIncomePrev' },
        { label: 'Net Taxable Income (to date)', amount: taxableToDate, type: 'subtotal' },
        { label: 'Income Tax Due (graduated, ' + (year >= 2023 ? '2023 onwards' : '2018–2022') + ' table)', amount: taxDue, type: 'total' }
      );
    }
    return finish('1701Q', schedule, lines, credits(i), taxDue, taxableToDate, warnings, q);
  }

  const TAX_TYPES = {
    rcit25: { label: 'RCIT 25%', rate: 0.25, mcit: true },
    rcit20: { label: 'RCIT 20%', rate: 0.20, mcit: true },
    mcit: { label: 'MCIT only', rate: 0, mcit: true, only: true },
    iaet10: { label: 'IAET 10%', rate: 0.10, mcit: false },
    nonres25: { label: 'RCIT — Non-Resident 25%', rate: 0.25, mcit: false }
  };

  function compute1702Q(i) {
    const lines = [], warnings = [];
    const q = Number(i.quarter);
    const tt = TAX_TYPES[i.taxType] || TAX_TYPES.rcit25;
    const mcitRate = Number(i.mcitRate) || 0.02;
    const p = i.prior || {};
    const itemized = i.method !== 'osd';

    const rev = r2(i.revenue);
    const cos = r2(i.costOfSales);
    const grossProfit = r2(rev - cos);
    const other = r2(i.otherIncome);
    const totalGross = r2(grossProfit + other);
    const opex = itemized ? r2(i.opex) : r2(pos(totalGross) * 0.40);
    const otherDed = r2(i.otherDeductions);
    const taxableQ = r2(totalGross - opex - otherDed);
    const prev = r2(p.taxableIncomePrev);
    const toDate = r2(pos(taxableQ + prev));
    const regular = r2(toDate * tt.rate);

    // MCIT base: gross income (NIRC Sec. 27(E)) by default, or gross sales as in the office Tax Calc
    const useSales = i.mcitBase === 'sales';
    const baseThisQ = useSales ? rev : totalGross;
    const mcitBase = r2(baseThisQ + r2(p.grossIncomePrev));
    const mcitOn = tt.mcit && (tt.only || !i.mcitExempt);
    const mcit = mcitOn ? r2(pos(mcitBase) * mcitRate) : 0;
    let taxDue, taxLabel;
    if (tt.only) { taxDue = mcit; taxLabel = 'MCIT ' + (mcitRate * 100) + '%'; }
    else if (mcitOn) { taxDue = Math.max(regular, mcit); taxLabel = regular >= mcit ? tt.label : 'MCIT ' + (mcitRate * 100) + '% (higher)'; }
    else { taxDue = regular; taxLabel = tt.label; }
    if (mcitOn && !tt.only && mcit > regular) warnings.push('MCIT is higher than the regular income tax, so MCIT is the tax due this quarter.');
    if (i.taxType === 'rcit20') warnings.push('The 20% rate applies only if net taxable income is ₱5,000,000 or less AND total assets (excluding land) are ₱100,000,000 or less.');

    lines.push(
      { label: 'Gross Sales / Receipts', amount: rev },
      { label: 'Less: Cost of Goods Sold / Services', amount: cos },
      { label: 'Gross Profit', amount: grossProfit, type: 'subtotal' },
      itemized ? { label: 'Less: Total Allowable Itemized Deductions (Operating Expenses)', amount: opex } : { label: 'Less: OSD (40% of gross income)', amount: opex },
      { label: 'Add: Other Income', amount: other, editable: 'otherIncome' },
      { label: 'Less: Other Deductions', amount: otherDed, editable: 'otherDeductions' },
      { label: 'Net Taxable Income this Quarter', amount: taxableQ, type: 'subtotal' },
      { label: 'Add: Taxable Income — Previous Quarter/s', amount: prev, editable: 'taxableIncomePrev' },
      { label: 'Net Taxable Income (to date)', amount: toDate, type: 'subtotal' }
    );
    if (!tt.only) lines.push({ label: tt.label + ' on Net Taxable Income', amount: regular });
    if (tt.mcit) {
      lines.push(
        { label: (useSales ? 'Gross Sales' : 'Gross Income') + ' — Previous Quarter/s (for MCIT)', amount: r2(p.grossIncomePrev), editable: 'grossIncomePrev' },
        { label: 'MCIT Base (cumulative ' + (useSales ? 'gross sales' : 'gross income') + ')', amount: mcitBase },
        { label: 'MCIT ' + (mcitRate * 100) + '%' + (mcitOn ? '' : ' — not applicable (first 3 years)'), amount: mcit }
      );
    }
    lines.push({ label: 'Tax Due (' + taxLabel + ')', amount: r2(taxDue), type: 'total' });

    return finish('1702Q', 'Regular Income Tax — ' + (itemized ? 'Itemized' : 'OSD') + ' · ' + tt.label, lines, credits(i), taxDue, toDate, warnings, q);
  }

  function finish(form, schedule, lines, c, taxDue, taxableToDate, warnings, q) {
    // Tax Calc logic: Tax Still Due = max(0, Tax Due − credits); any excess is carried over
    const net = r2(taxDue - c.total);
    const stillDue = r2(Math.max(0, net));
    const excess = r2(Math.max(0, -net));
    if (excess > 0) warnings.push('Credits exceed the tax due by ₱' + excess.toLocaleString('en-PH', { minimumFractionDigits: 2 }) + '. Tax Still Due is ₱0.00 and the excess carries over to the next quarter.');
    if (q === 1 && (c.items[1].amount || c.items[2].amount)) warnings.push('Q1 returns normally have no previous-quarter payments or withholding.');
    return {
      form, schedule, lines,
      credits: c.items, totalCredits: c.total,
      taxDue: r2(taxDue), taxableToDate: r2(taxableToDate),
      taxPayable: stillDue, excessCredits: excess, warnings
    };
  }

  function compute(input) {
    return input.form === '1702Q' ? compute1702Q(input) : compute1701Q(input);
  }

  /** Auto-compute CWT: 1% of revenue for goods, 2% for services (RR 11-2018 / RR 31-2019) */
  function autoCWT(revenue, kind) { return r2(pos(revenue) * (kind === 'Services' ? 0.02 : 0.01)); }

  const api = { compute, graduatedTax, entertainmentCap, autoCWT, TAX_TYPES, r2 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ITRTax = api;
})(typeof window !== 'undefined' ? window : globalThis);
