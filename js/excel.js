/* ITR Generator — Excel export (ExcelJS, loaded on demand) */
(function () {
  'use strict';
  const NAVY = 'FF0B2340', GOLD = 'FFF5B301', LIGHT = 'FFF2F6FC', SEC = 'FFE8EEF7', WHITE = 'FFFFFFFF';
  const NUM = '#,##0.00;(#,##0.00);"-"';
  const CAT = { cost_of_sales: 'Cost of Sales', operating_expense: 'Operating Expense', non_deductible: 'Non-deductible' };
  const QTR = { 1: 'January – March', 2: 'April – June', 3: 'July – September' };

  function fill(c, argb) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }; }
  function border(c, style, color) { c.border = Object.assign({}, c.border, { top: { style, color: { argb: color || NAVY } } }); }

  function buildComputation(wb, M) {
    const ws = wb.addWorksheet('ITR Computation', { views: [{ showGridLines: false }], pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = [{ width: 66 }, { width: 22 }];
    const T = M.tax;

    const title = ws.addRow(['BIR Form ' + M.form + ' — Quarterly Income Tax Return Computation']);
    ws.mergeCells(title.number, 1, title.number, 2);
    title.height = 28;
    title.getCell(1).font = { bold: true, size: 15, color: { argb: WHITE } };
    title.getCell(1).alignment = { vertical: 'middle', indent: 1 };
    fill(title.getCell(1), NAVY);
    const sub = ws.addRow(['Sta. Rita Accounting Office · ITR Generator · figures extracted by Muse Spark 1.3 and computed by formula']);
    ws.mergeCells(sub.number, 1, sub.number, 2);
    sub.getCell(1).font = { italic: true, size: 9, color: { argb: 'FF64748B' } };
    ws.addRow([]);

    const info = [
      ['Taxpayer / Registered name', M.client.name],
      ['TIN', M.client.tin],
      ['RDO code', M.client.rdo || ''],
      ['Taxpayer type / VAT status', M.client.type + ' · ' + M.client.vat],
      ['Line of business', M.client.lob || ''],
      ['Taxable period', 'Q' + M.quarter + ' ' + M.year + ' (' + QTR[M.quarter] + ')'],
      ['Tax method', M.methodLabel],
      ['Expenses source', M.R.expenseSource === 'auto' ? 'Auto-generated at ' + M.R.auto.percent + '% of revenue' : 'Uploaded documents'],
      ['Prepared by', (M.user && M.user.name) || ''],
      ['Generated on', new Date().toLocaleString('en-PH')]
    ];
    info.forEach(([k, v]) => {
      const r = ws.addRow([k, v]);
      r.getCell(1).font = { bold: true, color: { argb: 'FF334155' } };
      r.getCell(2).alignment = { horizontal: 'right' };
    });
    ws.addRow([]);

    const section = (label) => {
      const r = ws.addRow([label, '']);
      [1, 2].forEach((i) => { fill(r.getCell(i), NAVY); r.getCell(i).font = { bold: true, color: { argb: WHITE } }; });
      r.height = 20;
    };
    const head = (label) => {
      const r = ws.addRow([label, '']);
      [1, 2].forEach((i) => { fill(r.getCell(i), SEC); r.getCell(i).font = { bold: true, size: 10, color: { argb: 'FF1D4A7E' } }; });
    };
    const line = (label, amt, kind) => {
      const r = ws.addRow([label, typeof amt === 'number' ? amt : amt]);
      const a = r.getCell(2);
      if (typeof amt === 'number') a.numFmt = NUM;
      a.alignment = { horizontal: 'right' };
      if (kind === 'ind') r.getCell(1).alignment = { indent: 2 };
      if (kind === 'sub') { r.font = { bold: true }; }
      if (kind === 'tot') { r.font = { bold: true }; [1, 2].forEach((i) => { fill(r.getCell(i), LIGHT); border(r.getCell(i), 'thin'); }); }
      if (kind === 'grand') {
        [1, 2].forEach((i) => { fill(r.getCell(i), NAVY); r.getCell(i).font = { bold: true, size: 12, color: { argb: i === 2 ? GOLD : WHITE } }; });
        r.height = 22;
      }
      if (kind === 'muted') r.font = { color: { argb: 'FF64748B' }, italic: true };
      return r;
    };

    section('INCOME STATEMENT');
    head('Revenue');
    M.revenueLines.forEach((l) => line(l.label, l.amount, 'ind'));
    line('Net sales / revenue', M.revenue, 'sub');
    head('Cost of sales / services');
    const cos = M.accounts.filter((a) => a.category === 'cost_of_sales');
    if (cos.length) cos.forEach((a) => line(a.account, a.amount, 'ind')); else line('None', 0, 'ind');
    line('Total cost of sales', M.cos, 'tot');
    line('Gross profit', M.grossProfit, 'sub');
    head('Operating expenses');
    const op = M.accounts.filter((a) => a.category === 'operating_expense');
    if (op.length) op.forEach((a) => line(a.account, a.amount, 'ind')); else line('None', 0, 'ind');
    line('Total operating expenses', M.opex, 'tot');
    if (M.ent.disallowed > 0 && M.method === 'itemized') {
      line('Less: Representation & entertainment over ' + (M.ent.rate * 100) + '% limit', -M.ent.disallowed, 'muted');
      line('Allowable operating expenses', M.deductibleOpex, 'sub');
    }
    const nd = M.accounts.filter((a) => a.category === 'non_deductible');
    if (nd.length) { head('Non-deductible (excluded from tax)'); nd.forEach((a) => line(a.account, a.amount, 'muted')); }
    line('Net income / (loss) per books', M.netIncome, 'grand');
    ws.addRow([]);

    section('TAX COMPUTATION — BIR FORM ' + M.form);
    head(T.schedule);
    T.lines.forEach((l) => {
      if (l.type === 'rate') line(l.label, l.amount / 100, 'ind').getCell(2).numFmt = '0%';
      else line(l.label, l.amount, l.type === 'total' ? 'tot' : l.type === 'subtotal' ? 'sub' : 'ind');
    });
    head('Less: Tax credits / payments');
    T.credits.forEach((c) => line(c.label, c.amount, 'ind'));
    line('Total tax credits / payments', T.totalCredits, 'tot');
    line(T.taxPayable < 0 ? 'OVERPAYMENT' : 'TAX PAYABLE', T.taxPayable, 'grand');
    ws.addRow([]);
    const note = ws.addRow(['Computed under the TRAIN Law and CREATE Act rates. Check all figures against the source documents before filing in eBIRForms / eFPS.']);
    ws.mergeCells(note.number, 1, note.number, 2);
    note.getCell(1).font = { italic: true, size: 9, color: { argb: 'FF64748B' } };
    note.getCell(1).alignment = { wrapText: true };
    note.height = 26;
  }

  function table(wb, name, headers, rows, widths, numCols, totals) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = headers.map((h, i) => ({ header: h, width: widths[i] || 16 }));
    const hr = ws.getRow(1);
    hr.eachCell((c) => { fill(c, NAVY); c.font = { bold: true, color: { argb: WHITE } }; c.alignment = { vertical: 'middle' }; });
    hr.height = 20;
    rows.forEach((r) => ws.addRow(r));
    numCols.forEach((ci) => ws.getColumn(ci).numFmt = NUM);
    if (totals && rows.length) {
      const t = ws.addRow(totals);
      t.font = { bold: true };
      t.eachCell((c) => { fill(c, LIGHT); border(c, 'thin'); });
      numCols.forEach((ci) => t.getCell(ci).numFmt = NUM);
    }
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    return ws;
  }

  async function download(M) {
    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'ITR Generator';
    wb.created = new Date();
    buildComputation(wb, M);

    const sumRows = M.accounts.map((a) => [a.account, CAT[a.category], a.amount, a.items.length, a.edited ? 'Edited (extracted ' + a.extracted.toFixed(2) + ')' : '']);
    const total = M.accounts.reduce((s, a) => s + a.amount, 0);
    table(wb, 'Expense Summary', ['Account', 'Category', 'Amount', 'Items', 'Remarks'], sumRows, [44, 20, 18, 8, 32], [3], ['TOTAL', '', Math.round(total * 100) / 100, M.R.items.length, '']);

    const det = M.R.items.map((it) => [it.file, it.date, it.vendor, it.description, it.account, CAT[it.category], it.amount, it.vat_amount || 0]);
    table(wb, 'Expense Details', ['Source file', 'Date', 'Vendor / payee', 'Description', 'Account', 'Category', 'Amount', 'VAT'], det, [28, 12, 26, 40, 32, 18, 16, 12], [7, 8],
      ['TOTAL', '', '', '', '', '', Math.round(M.R.items.reduce((s, i) => s + i.amount, 0) * 100) / 100, Math.round(M.R.items.reduce((s, i) => s + (i.vat_amount || 0), 0) * 100) / 100]);

    const docRows = M.R.docs.map((d) => {
      const t = { sales_vat: '2550Q', sales_nonvat: '2551Q', prior_itr: 'Previous ' + M.form, expenses: 'Expenses' }[d.type];
      const amt = d.type === 'expenses' ? d.data.document_total : d.type === 'prior_itr' ? d.data.tax_due : d.data.revenue;
      const per = d.data.quarter ? 'Q' + d.data.quarter + ' ' + (d.data.year || '') : (d.data.period_text || '');
      return [t, d.file, per, amt, d.data.confidence || '', d.data.notes || '', d.url || ''];
    });
    table(wb, 'Source Documents', ['Type', 'File', 'Period', 'Key amount', 'Confidence', 'Notes', 'Archived copy'], docRows, [16, 34, 12, 18, 12, 50, 40], [4]);

    const notes = (M.R.warnings || []).concat(M.tax.warnings || []).map((w) => [w.replace(/^info:/, '')]);
    if (notes.length) table(wb, 'Review Notes', ['Items to review before filing'], notes, [120], []);

    const buf = await wb.xlsx.writeBuffer();
    const safe = (s) => String(s).replace(/[^\w\-]+/g, '_').replace(/_+/g, '_').slice(0, 40);
    const name = M.form + '_' + M.year + 'Q' + M.quarter + '_' + safe(M.client.name) + '.xlsx';
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    return name;
  }

  window.ITRExcel = { download };
})();
