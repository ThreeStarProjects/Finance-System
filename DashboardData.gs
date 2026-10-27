// ============================================================
// 3 Star Projects — Unified Finance System
// PHASE 4 : DASHBOARD + P&L  (backend compute endpoints)
// File: Dashboard.gs   (depends on Masters, Transactions, Invoices)
//
// All endpoints return a sanitized envelope:
//   {ok:true, data:{...}, error:null}  |  {ok:false, data:null, error:'msg'}
// so the client (Phase 5/6) never receives null from google.script.run.
//
// Money model consolidated here:
//   • Sales / revenue  ← Invoices (taxable value, GST excluded from revenue)
//   • Expenses         ← Transactions  class = Expense
//   • Drawings         ← Transactions  class = Drawing   (cash only, NOT P&L)
//   • Cash in          ← Payments
//   • Receivables/AR   ← Invoices outstanding + ageing
//   • GST              ← output (invoices) vs input (expense GST)
//   • Personal         ← Personal_Transactions  (separate, never in business P&L)
//
// filters = {fy, month, project}   fy defaults to current FY; month optional.
// ============================================================

// ── Envelope + JSON-safe (so google.script.run never gets null) ─
function _envOk(data){ return {ok:true, data:_jsonSafe(data), error:null}; }
function _envErr(e){ return {ok:false, data:null, error:String(e && e.message || e)}; }
function _jsonSafe(v){
  if (v === null || v === undefined) return null;
  var t = typeof v;
  if (t === 'number') return isFinite(v) ? v : null;
  if (t === 'string' || t === 'boolean') return v;
  if (v instanceof Date){ return isNaN(v.getTime()) ? null : Utilities.formatDate(v,'GMT+5:30','yyyy-MM-dd'); }
  if (Array.isArray(v)){ var a=[]; for (var i=0;i<v.length;i++) a.push(_jsonSafe(v[i])); return a; }
  if (t === 'object'){ var o={}; for (var k in v){ if (Object.prototype.hasOwnProperty.call(v,k)) o[k]=_jsonSafe(v[k]); } return o; }
  return String(v);
}

// ── Period helpers ──────────────────────────────────────────
function _activeFY(filters){
  filters = filters || {};
  if (!filters.fy || filters.fy === 'current') return _currentFY();
  return _s(filters.fy);
}
function _inMonth(filters){ return filters && filters.month && filters.month !== 'all' ? _s(filters.month) : null; }

// All FYs that have a Transactions tab (for the FY picker)
function _availableFYs(){
  var out = [];
  _getSS().getSheets().forEach(function(sh){
    var n = sh.getName();
    if (n.indexOf(T.TXN_PREFIX) === 0) out.push(n.substring(T.TXN_PREFIX.length));
  });
  if (out.indexOf(_currentFY()) === -1) out.push(_currentFY());
  return out.sort();
}
function _monthInvoice(inv){ try { return _monthOf(inv.invoiceDate); } catch(e){ return ''; } }

// Live Category → Group map, rebuilt per call so renames in M_Category take
// effect immediately on the next dashboard load. Includes employee-only
// categories so an employee-tagged row resolves too. One sheet read instead
// of per-row _groupForCategory lookups.
function _buildCatToGroupMapLive(){
  var map = {};
  try {
    (getCategories() || []).forEach(function(c){
      if (c && c.category) map[String(c.category).trim()] = String(c.group || '').trim();
    });
  } catch(e){}
  try {
    (getEmployeeCategories() || []).forEach(function(c){
      if (c && c.category && !map[c.category]) map[String(c.category).trim()] = String(c.group || '').trim();
    });
  } catch(e){}
  return map;
}

// ============================================================
// KPIs — both sides, single pass each over txns + invoices
// ============================================================
// ============================================================
// findOwnerFundsCompanyRows() — lists every Fund In/Out (Owner) transfer row
// whose On_Behalf_Of column reads "Company" (or blank). These pollute the
// Owner Funds breakdown by masquerading as an owner. Returns row locations
// and a click-through hint so the user can fix them fast.
//
// Callable from the Tools tab button ("Find Owner-Funds 'Company' rows").
// ============================================================
function findOwnerFundsCompanyRows(){
  try {
    var offenders = [];
    var sheets = _allTxnSheets();
    sheets.forEach(function(sheet){
      if (!sheet || sheet.getLastRow() < 2) return;
      var w = Math.min(sheet.getLastColumn(), TC_TOTAL);
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
      for (var i=0;i<rows.length;i++){
        var r = rows[i];
        if (_s(r[TC.CLASS]) !== 'Transfer') continue;
        var cat = _s(r[TC.CATEGORY]);
        if (cat !== 'Fund In (Owner)' && cat !== 'Fund Out (Owner)') continue;
        var onBehalf = _s(r[TC.ON_BEHALF_OF]);
        // Anything that isn't a real owner name is suspect: blank, "Company",
        // "Owner" (generic default), etc.
        var suspect = !onBehalf || /^company$/i.test(onBehalf) || /^owner$/i.test(onBehalf);
        if (!suspect) continue;
        offenders.push({
          sheet:       sheet.getName(),
          row:         i + 2,
          id:          _s(r[TC.ID]),
          date:        _safeDateStr(r[TC.TX_DATE]),
          category:    cat,
          amount:      _n(r[TC.AMOUNT]),
          onBehalfOf:  onBehalf || '(blank)',
          submittedBy: _s(r[TC.SUBMITTED_BY]),
          source:      _s(r[TC.SOURCE]),
          description: _s(r[TC.DESCRIPTION]) || _s(r[TC.NOTES])
        });
      }
    });
    // Sort newest first for easy fixing
    offenders.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
    var total = offenders.reduce(function(s,o){ return s + (o.category === 'Fund In (Owner)' ? o.amount : -o.amount); }, 0);
    return _envOk({
      count:    offenders.length,
      netImpact: _round(total),  // Fund In minus Fund Out
      rows:     offenders,
      howToFix: 'Open each sheet+row above and set On_Behalf_Of to the actual owner name (Reeta, Girish, etc). After fixing, refresh the dashboard — the "Company" line disappears from Owner Funds.'
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// deepTally(fy) — diagnostic for reconciling dashboard KPIs against raw sheet
// data. Reads every source the dashboard uses and returns a nested breakdown
// so mismatches (case, whitespace, blank FY, hidden filters, mystery rows)
// pop out.
//
// Callable from:
//   • Apps Script Editor → Select `deepTally` → Run → check Logger.log
//   • Frontend Tools tab (button wires to this fn and pretty-prints).
//
// Returns an object with per-source breakdowns and a top-level "reconciliation"
// section showing every KPI's formula and value so you can eyeball what's off.
// ============================================================
function deepTally(fy){
  try {
    var ss = _getSS();
    // Strip ALL whitespace so a typo like "2025 -26" or "2026 - 27" still
    // resolves to the correct FY. The FY stored in sheets is always
    // whitespace-free (e.g. "2026-27").
    fy = _s(fy).replace(/\s+/g, '') || _currentFY();
    var out = { fy: fy, at: Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss') };

    // ── Transactions_YYYY-YY ─────────────────────────────────
    var txSheet = ss.getSheetByName(_txnSheetName(fy));
    var tx = { sheetName: _txnSheetName(fy), rows: 0, byClass: {}, byFY: {}, excludedDeleted: 0, blankFY: 0, blankClass: 0, oddClassValues: {} };
    if (txSheet && txSheet.getLastRow() > 1){
      var tw = Math.min(txSheet.getLastColumn(), TC_TOTAL);
      var trows = txSheet.getRange(2,1,txSheet.getLastRow()-1,tw).getValues();
      tx.rows = trows.length;
      trows.forEach(function(r){
        var id = _s(r[TC.ID]); if (!id) return;
        var rawClass = _s(r[TC.CLASS]);
        var rawFy    = _s(r[TC.FY]);
        var deleted  = _s(r[TC.DELETED]) === 'Yes';
        var amt      = _n(r[TC.AMOUNT]);
        var key      = rawClass || '(blank)';
        if (!tx.byClass[key]) tx.byClass[key] = {countAll:0, amountAll:0, countActive:0, amountActive:0, countDeleted:0, amountDeleted:0};
        tx.byClass[key].countAll++;
        tx.byClass[key].amountAll += amt;
        if (deleted){
          tx.byClass[key].countDeleted++;
          tx.byClass[key].amountDeleted += amt;
          tx.excludedDeleted += amt;
        } else {
          tx.byClass[key].countActive++;
          tx.byClass[key].amountActive += amt;
        }
        if (!tx.byFY[rawFy || '(blank)']) tx.byFY[rawFy || '(blank)'] = 0;
        tx.byFY[rawFy || '(blank)']++;
        if (!rawFy) tx.blankFY += amt;
        if (!rawClass) tx.blankClass += amt;
        // Odd class values (case/whitespace variants of "Expense")
        if (rawClass && rawClass !== 'Expense' && rawClass.toLowerCase().indexOf('expens') !== -1){
          if (!tx.oddClassValues[rawClass]) tx.oddClassValues[rawClass] = 0;
          tx.oddClassValues[rawClass] += amt;
        }
      });
      // Round everything for readability
      Object.keys(tx.byClass).forEach(function(k){
        var b = tx.byClass[k];
        b.amountAll = _round(b.amountAll);
        b.amountActive = _round(b.amountActive);
        b.amountDeleted = _round(b.amountDeleted);
      });
      tx.excludedDeleted = _round(tx.excludedDeleted);
      tx.blankFY = _round(tx.blankFY);
      tx.blankClass = _round(tx.blankClass);
    }
    out.transactions = tx;

    // ── Bills ─────────────────────────────────────────────────
    var bs = ss.getSheetByName(AP.BILLS);
    var b = { rows: 0, byFY: {}, byStatus: {}, totals: {gross:0, taxable:0, paid:0, outstanding:0}, excludedCancelled: 0 };
    if (bs && bs.getLastRow() > 1){
      var bw = Math.min(bs.getLastColumn(), BC_TOTAL);
      var brows = bs.getRange(2,1,bs.getLastRow()-1,bw).getValues();
      b.rows = brows.length;
      brows.forEach(function(r){
        if (!_s(r[BC.ID])) return;
        var rawFy = _s(r[BC.FY]) || '(blank)';
        var status = _s(r[BC.STATUS]) || '(blank)';
        if (!b.byFY[rawFy]) b.byFY[rawFy] = {count:0, gross:0, paid:0, outstanding:0};
        b.byFY[rawFy].count++;
        b.byFY[rawFy].gross += _n(r[BC.AMOUNT]);
        b.byFY[rawFy].paid += _n(r[BC.PAID]);
        b.byFY[rawFy].outstanding += _n(r[BC.OUTSTANDING]);
        if (!b.byStatus[status]) b.byStatus[status] = {count:0, gross:0};
        b.byStatus[status].count++;
        b.byStatus[status].gross += _n(r[BC.AMOUNT]);
        if (status === BILL_STATUS.CANCELLED){
          b.excludedCancelled += _n(r[BC.AMOUNT]);
        } else if (rawFy === fy){
          b.totals.gross += _n(r[BC.AMOUNT]);
          b.totals.taxable += _n(r[BC.TAXABLE]) || _n(r[BC.AMOUNT]);
          b.totals.paid += _n(r[BC.PAID]);
          b.totals.outstanding += _n(r[BC.OUTSTANDING]);
        }
      });
      Object.keys(b.byFY).forEach(function(k){
        b.byFY[k].gross = _round(b.byFY[k].gross);
        b.byFY[k].paid = _round(b.byFY[k].paid);
        b.byFY[k].outstanding = _round(b.byFY[k].outstanding);
      });
      Object.keys(b.byStatus).forEach(function(k){ b.byStatus[k].gross = _round(b.byStatus[k].gross); });
      b.totals.gross = _round(b.totals.gross);
      b.totals.taxable = _round(b.totals.taxable);
      b.totals.paid = _round(b.totals.paid);
      b.totals.outstanding = _round(b.totals.outstanding);
      b.excludedCancelled = _round(b.excludedCancelled);
    }
    out.bills = b;

    // ── Bill_Payments ─────────────────────────────────────────
    var bp = ss.getSheetByName(AP.PAYMENTS);
    var bpOut = { rows: 0, byFY: {}, activeTotalFY: 0, linkedTotalFY: 0, unlinkedTotalFY: 0, cancelledFY: 0 };
    if (bp && bp.getLastRow() > 1){
      var bpw = Math.min(bp.getLastColumn(), BPMC_TOTAL);
      var bprows = bp.getRange(2,1,bp.getLastRow()-1,bpw).getValues();
      bpOut.rows = bprows.length;
      bprows.forEach(function(r){
        if (!_s(r[BPMC.ID])) return;
        var rawFy = _s(r[BPMC.FY]) || '(blank)';
        var cancelled = bpw > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON]);
        var applied = _s(r[BPMC.APPLIED]);
        var amt = _n(r[BPMC.AMOUNT]);
        if (!bpOut.byFY[rawFy]) bpOut.byFY[rawFy] = {count:0, amount:0, active:0, linked:0, unlinked:0};
        bpOut.byFY[rawFy].count++;
        bpOut.byFY[rawFy].amount += amt;
        if (!cancelled){
          bpOut.byFY[rawFy].active += amt;
          if (applied) bpOut.byFY[rawFy].linked += amt;
          else         bpOut.byFY[rawFy].unlinked += amt;
        }
        if (rawFy === fy){
          if (cancelled){ bpOut.cancelledFY += amt; }
          else {
            bpOut.activeTotalFY += amt;
            if (applied) bpOut.linkedTotalFY += amt;
            else         bpOut.unlinkedTotalFY += amt;
          }
        }
      });
      Object.keys(bpOut.byFY).forEach(function(k){
        var x = bpOut.byFY[k];
        x.amount = _round(x.amount); x.active = _round(x.active);
        x.linked = _round(x.linked); x.unlinked = _round(x.unlinked);
      });
      bpOut.activeTotalFY = _round(bpOut.activeTotalFY);
      bpOut.linkedTotalFY = _round(bpOut.linkedTotalFY);
      bpOut.unlinkedTotalFY = _round(bpOut.unlinkedTotalFY);
      bpOut.cancelledFY = _round(bpOut.cancelledFY);
    }
    out.billPayments = bpOut;

    // ── Reconciliation vs dashKpis ────────────────────────────
    var kpi = dashKpis({fy: fy});
    var d = kpi && kpi.data ? kpi.data : null;
    if (d){
      var directFromTx = (tx.byClass['Expense'] ? tx.byClass['Expense'].amountActive : 0);
      out.reconciliation = {
        dashboardValues: {
          totalExpense:   d.totalExpenses,
          billsBooked:    d.billsBooked,
          billsPaid:      d.billsPaid,
          billsUnpaid:    d.billsUnpaid,
          directExpense:  d.directExpense,
          unBilledVendorPay: d.unBilledVendorPay,
          cashOut:        d.cashOut,
          cashIn:         d.cashIn,
          totalSales:     d.totalSales,
          outstandingAR:  d.outstandingAR
        },
        expected: {
          '(Total Expense) should equal (Bills Booked + Direct + Un-billed vendor)':
            _round(d.billsBooked + d.directExpense + d.unBilledVendorPay),
          '(Bills Booked) should equal (Bills Paid + Bills Remaining)':
            _round(d.billsPaid + d.billsUnpaid),
          '(Direct Expense on dashboard) should equal (Transactions Expense active for this FY)':
            _round(directFromTx)
        },
        gaps: {
          totalExpense_vs_sumOfParts:
            _round(d.totalExpenses - (d.billsBooked + d.directExpense + d.unBilledVendorPay)),
          billsBooked_vs_paidPlusRemaining:
            _round(d.billsBooked - (d.billsPaid + d.billsUnpaid)),
          directExpense_vs_txnClassExpense:
            _round(d.directExpense - directFromTx)
        }
      };
    }

    Logger.log(JSON.stringify(out, null, 2));
    return _envOk(out);
  } catch(e){
    Logger.log('deepTally error: ' + e.message + '\n' + e.stack);
    return _envErr(e);
  }
}

function dashKpis(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters), month = _inMonth(filters), proj = filters.project ? _s(filters.project) : null;
    var txns = getAllTransactions({fy:fy});
    var invs = getInvoices({fy:fy});

    var expense=0, directExpense=0, drawings=0, gstInput=0, pendAmt=0, pendCount=0, advAmt=0, advCount=0, cashOut=0;
    var advByEmp = {};
    txns.forEach(function(r){
      if (month && r.month !== month) return;
      if (proj && r.project !== proj) return;
      if (r.txnClass === 'Expense'){
        expense += r.amount;
        directExpense += r.amount;   // non-bill expenses from the Transactions sheet
        gstInput += (r.cgst + r.sgst + r.igst);
        var isEmpPending = (r.entryType === 'Employee' && r.settleStatus === 'Pending');
        if (isEmpPending){ pendAmt += r.amount; pendCount++; }
        // Critical #4 — pending employee expenses are the EMPLOYEE's cash, not
        // the company's, until the owner settles. Skip them from cashOut.
        if (!isEmpPending) cashOut += r.amount;
      } else if (r.txnClass === 'Drawing'){
        drawings += r.amount; cashOut += r.amount;
      }
      // advances (recovery pending), all-time within FY. Tier 4.2: use
      // (Amount − Recovered_Amount) so partial recoveries are reflected.
      if (r.recovery === 'Pending' && r.onBehalfOf){
        var _out = r.amount - (r.recoveredAmount || 0);
        if (_out > 0) advByEmp[r.onBehalfOf] = (advByEmp[r.onBehalfOf]||0) + _out;
      }
    });
    for (var e in advByEmp){ if (advByEmp[e] > 0){ advAmt += advByEmp[e]; advCount++; } }

    // Bill contribution to Total Expense — GROSS (incl GST) so the KPIs tally:
    //   Total Expense (Billed) = billsBooked + directExpense + un-billed vendor payments
    //   Bills Paid + Bills Remaining = billsBooked
    // (Previously used TAXABLE which created a GST-sized gap between Total
    // Expense and Bills Booked. Statutory P&L still uses taxable — see dashPnl.)
    var billsBooked = 0, billsUnpaidTotal = 0;
    var bs2 = _getSS().getSheetByName(AP.BILLS);
    if (bs2 && bs2.getLastRow() > 1){
      var bW2 = Math.min(bs2.getLastColumn(), BC_TOTAL);
      bs2.getRange(2,1,bs2.getLastRow()-1,bW2).getValues().forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (_s(r[BC.FY]) !== fy) return;
        if (month && _monthOf(r[BC.BILL_DATE]) !== month) return;
        if (proj && bW2 > BC.PROJECT && _s(r[BC.PROJECT]) !== proj) return;
        var billGross = _n(r[BC.AMOUNT]);
        expense          += billGross;
        billsBooked      += billGross;              // gross incl GST — matches "amount you owed"
        billsUnpaidTotal += _n(r[BC.OUTSTANDING]);  // still due
      });
    }
    // Bills paid = sum of vendor payments in this FY (cash out on bills).
    // Uses BPMC.AMOUNT (gross applied against vendor account). Includes both
    // bill-linked and unlinked advance payments — all real cash out.
    var billsPaidCash = 0;
    var bpKpi = _getSS().getSheetByName(AP.PAYMENTS);
    if (bpKpi && bpKpi.getLastRow() > 1){
      var bpKW = Math.min(bpKpi.getLastColumn(), BPMC_TOTAL);
      bpKpi.getRange(2,1,bpKpi.getLastRow()-1,bpKW).getValues().forEach(function(r){
        if (bpKW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        if (_s(r[BPMC.FY]) !== fy) return;
        if (month && _monthOf(r[BPMC.DATE]) !== month) return;
        billsPaidCash += _n(r[BPMC.AMOUNT]);
      });
    }
    // Un-billed vendor payments (Applied_Bills blank) — real cash purchases
    // with no bill on file. Contribute to Total Expense AND cash out.
    var unBilledVendorPay = 0;
    var bp2 = _getSS().getSheetByName(AP.PAYMENTS);
    if (bp2 && bp2.getLastRow() > 1){
      var bpW2 = Math.min(bp2.getLastColumn(), BPMC_TOTAL);
      bp2.getRange(2,1,bp2.getLastRow()-1,bpW2).getValues().forEach(function(r){
        if (bpW2 > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        if (_s(r[BPMC.APPLIED])) return;
        if (_s(r[BPMC.FY]) !== fy) return;
        if (month && _monthOf(r[BPMC.DATE]) !== month) return;
        if (proj && bpW2 > BPMC.PROJECT && _s(r[BPMC.PROJECT]) !== proj) return;
        var amt = _n(r[BPMC.AMOUNT]);
        expense += amt;
        unBilledVendorPay += amt;
        cashOut += amt;   // real cash out — actual bank debit
      });
    }

    var sales=0, gstOutput=0, outstanding=0, overdue=0, openCount=0, cashIn=0;
    var today = new Date(); today.setHours(0,0,0,0);
    // Build a project map ONCE if we're filtering by project (used for both
    // invoices and returns — returns cite invoice_no, so we resolve via this).
    var invProjOf = null;
    if (proj){
      invProjOf = {};
      invs.forEach(function(v){ if (v.no) invProjOf[v.no] = v.project || ''; });
    }
    invs.forEach(function(v){
      if (v.status === INV_STATUS.CANCELLED) return;
      if (proj && v.project !== proj) return;
      var inPeriod = !month || _monthInvoice(v) === month;
      if (inPeriod){
        // Gross (incl GST) so Total Sales tallies with Cash In and Outstanding
        // AR, both of which are gross. Statutory P&L below uses taxable — this
        // is only for the operator KPI.
        sales += _n(v.amount);
        gstOutput += (v.cgst + v.sgst + v.igst);
      }
      // outstanding/overdue are live balances (not month-bound)
      if (v.outstanding > 0){
        outstanding += v.outstanding; openCount++;
        var due = _asDate(v.dueDate); due.setHours(0,0,0,0);
        if (due < today) overdue += v.outstanding;
      }
    });
    // Critical #5 — returns reduce net sales for the period. Respect project filter.
    var rSheet2 = _getSS().getSheetByName(AR.RETURNS);
    if (rSheet2 && rSheet2.getLastRow() > 1){
      var rW2 = Math.min(rSheet2.getLastColumn(), RC_TOTAL);
      rSheet2.getRange(2,1,rSheet2.getLastRow()-1,rW2).getValues().forEach(function(r){
        if (_s(r[RC.FY]) !== fy) return;
        if (month && _monthOf(r[RC.DATE]) !== month) return;
        if (proj && invProjOf && invProjOf[_s(r[RC.INV_NO])] !== proj) return;
        sales -= _n(r[RC.AMOUNT]);
      });
    }

    // cash in = payments in period. Read only the columns the sheet actually
    // has (pre-Tier-2.1 sheets are 15 wide). Skip cancelled payments once the
    // audit column exists (Tier 2.1).
    var pSheet = _getSS().getSheetByName(AR.PAYMENTS);
    if (pSheet && pSheet.getLastRow()>1){
      var pW = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
      pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues().forEach(function(r){
        if (_s(r[PMC.FY]) !== fy) return;
        if (month && _monthOf(r[PMC.DATE]) !== month) return;
        if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
        cashIn += _n(r[PMC.AMOUNT]);
      });
    }

    return _envOk({
      fy:fy, month: month || 'all',
      totalSales:_round(sales), totalExpenses:_round(expense), netPnl:_round(sales - expense),
      cashIn:_round(cashIn), cashOut:_round(cashOut),
      outstandingAR:_round(outstanding), overdueAR:_round(overdue), openInvoices:openCount,
      pendingSettleAmt:_round(pendAmt), pendingSettleCount:pendCount,
      advancesOutstanding:_round(advAmt), advanceCount:advCount,
      drawings:_round(drawings),
      gstOutput:_round(gstOutput), gstInput:_round(gstInput), gstNet:_round(gstOutput - gstInput),
      // Bill-specific breakdown for the split "Booked vs Paid" KPI cards.
      billsBooked:_round(billsBooked), billsPaid:_round(billsPaidCash), billsUnpaid:_round(billsUnpaidTotal),
      // Total Expense breakdown so operators see where the number comes from:
      // Total Expense (Billed) = billsBooked + directExpense + unBilledVendorPay
      directExpense:_round(directExpense), unBilledVendorPay:_round(unBilledVendorPay)
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 3. P&L — expenses by group + sales + net  (replaces #2's P&L MIS)
//    Drawings + Transfers excluded. Every expense has a group (forced at
//    category creation) → nothing leaks → totals reconcile.
// ============================================================
function dashPnl(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters), month = _inMonth(filters), proj = filters.project ? _s(filters.project) : null;
    var txns = getAllTransactions({fy:fy});
    var invs = getInvoices({fy:fy});

    // Live-resolve: Category → Group from the current M_Category sheet (ignores
    // the stale Group value stored on each row at entry time). Groups appear
    // dynamically — any group used in M_Category today is a bucket today.
    var catToGroup = _buildCatToGroupMapLive();
    var groupMap = {};  // group name → amount, built on the fly from the live map
    var ungrouped = 0, totalExpense = 0;
    function _bump(cat, amt){
      var g = catToGroup[_s(cat).trim()];
      if (g){ groupMap[g] = (groupMap[g] || 0) + amt; }
      else   { ungrouped += amt; }
      totalExpense += amt;
    }

    txns.forEach(function(r){
      if (r.txnClass !== 'Expense') return;          // excludes Drawing / Transfer / Income
      // Advances to employees are receivables, not expenses — exclude them
      // from P&L (recovery=Pending still open, Recovered already offset by the
      // recovery transaction). The remaining 'N/A' / blank values are real
      // expenses that stay.
      if (r.recovery === 'Pending' || r.recovery === 'Recovered') return;
      if (month && r.month !== month) return;
      if (proj && r.project !== proj) return;
      _bump(r.category, r.amount);
    });

    // Critical #1 — Bills contribute to P&L (accrual side). Bill Category → Group,
    // live-resolved the same way as txns above (replaces the per-row
    // _groupForCategory sheet read).
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    if (billsSheet && billsSheet.getLastRow() > 1){
      var bW = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
      var bRows = billsSheet.getRange(2,1,billsSheet.getLastRow()-1,bW).getValues();
      bRows.forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (_s(r[BC.FY]) !== fy) return;
        if (month && _monthOf(r[BC.BILL_DATE]) !== month) return;
        if (proj && bW > BC.PROJECT && _s(r[BC.PROJECT]) !== proj) return;
        var amt = _n(r[BC.TAXABLE]) || _n(r[BC.AMOUNT]);
        _bump(r[BC.CATEGORY], amt);
      });
    }

    // Critical #3 — un-billed vendor payments (Applied_Bills empty) hit P&L
    // in a synthetic bucket so the money isn't invisible. Owner should
    // reclassify by creating the matching bill.
    var unbilledPay = 0;
    var bpSheet = _getSS().getSheetByName(AP.PAYMENTS);
    if (bpSheet && bpSheet.getLastRow() > 1){
      var pW = Math.min(bpSheet.getLastColumn(), BPMC_TOTAL);
      var pRows = bpSheet.getRange(2,1,bpSheet.getLastRow()-1,pW).getValues();
      pRows.forEach(function(r){
        if (pW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        if (_s(r[BPMC.APPLIED])) return;                 // has bill link → covered by Bills scan
        if (_s(r[BPMC.FY]) !== fy) return;
        if (month && _monthOf(r[BPMC.DATE]) !== month) return;
        if (proj && pW > BPMC.PROJECT && _s(r[BPMC.PROJECT]) !== proj) return;
        unbilledPay += _n(r[BPMC.AMOUNT]);
      });
      if (unbilledPay > 0) totalExpense += unbilledPay;
    }

    // Build invoice→project map for returns filter (only if project filter is active)
    var invProjOfPnl = null;
    if (proj){
      invProjOfPnl = {};
      invs.forEach(function(v){ if (v.no) invProjOfPnl[v.no] = v.project || ''; });
    }
    // Critical #5 — Returns reduce revenue. Respect project filter.
    var returnsInPeriod = 0;
    var rSheet = _getSS().getSheetByName(AR.RETURNS);
    if (rSheet && rSheet.getLastRow() > 1){
      var rW = Math.min(rSheet.getLastColumn(), RC_TOTAL);
      rSheet.getRange(2,1,rSheet.getLastRow()-1,rW).getValues().forEach(function(r){
        if (_s(r[RC.FY]) !== fy) return;
        if (month && _monthOf(r[RC.DATE]) !== month) return;
        if (proj && invProjOfPnl && invProjOfPnl[_s(r[RC.INV_NO])] !== proj) return;
        returnsInPeriod += _n(r[RC.AMOUNT]);
      });
    }

    var totalSales = 0;
    invs.forEach(function(v){
      if (v.status === INV_STATUS.CANCELLED) return;
      if (proj && v.project !== proj) return;
      if (month && _monthInvoice(v) !== month) return;
      totalSales += (v.taxable > 0 ? v.taxable : v.amount);
    });
    totalSales -= returnsInPeriod;

    // Groups come from whatever appears in the LIVE map now, not a hardcoded list
    // — renames and newly-added groups show up on the next refresh.
    var groups = Object.keys(groupMap).map(function(g){
      return {group:g, amount:_round(groupMap[g]),
              pct: totalExpense > 0 ? _round(groupMap[g]/totalExpense*100) : 0};
    }).filter(function(o){ return o.amount !== 0; })
      .sort(function(a,b){ return b.amount - a.amount; });

    if (ungrouped > 0) groups.push({group:'(Ungrouped)', amount:_round(ungrouped),
                                    pct: totalExpense>0?_round(ungrouped/totalExpense*100):0});
    if (unbilledPay > 0) groups.push({group:'(Un-billed vendor payments)', amount:_round(unbilledPay),
                                    pct: totalExpense>0?_round(unbilledPay/totalExpense*100):0});

    return _envOk({
      fy:fy, month: month || 'all',
      groups: groups,
      totalExpenses:_round(totalExpense),
      totalSales:_round(totalSales),
      netProfit:_round(totalSales - totalExpense)
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// AR AGEING — receivables by bucket + per-customer outstanding
// ============================================================
function dashArAgeing(filters){
  try {
    filters = filters || {};
    // FY-scope the ageing so it matches whatever FY the dashboard is on.
    // Pass fy='all' or blank to see receivables across every FY.
    var invs = getInvoices({fy: filters.fy || 'all'});
    var today = new Date(); today.setHours(0,0,0,0);
    var buckets = {current:0, b30:0, b60:0, b90:0, b90plus:0};
    var byCust = {};
    invs.forEach(function(v){
      if (v.status === INV_STATUS.CANCELLED || v.outstanding <= 0) return;
      var due = _asDate(v.dueDate); due.setHours(0,0,0,0);
      var days = Math.floor((today - due) / 86400000);
      if (days <= 0) buckets.current += v.outstanding;
      else if (days <= 30) buckets.b30 += v.outstanding;
      else if (days <= 60) buckets.b60 += v.outstanding;
      else if (days <= 90) buckets.b90 += v.outstanding;
      else buckets.b90plus += v.outstanding;
      var c = byCust[v.customerName] || (byCust[v.customerName] = {name:v.customerName, outstanding:0, invoices:0});
      c.outstanding += v.outstanding; c.invoices++;
    });
    var customers = Object.keys(byCust).map(function(k){ var c=byCust[k];
      return {name:c.name, outstanding:_round(c.outstanding), invoices:c.invoices}; })
      .sort(function(a,b){return b.outstanding-a.outstanding;});
    return _envOk({
      buckets:{ current:_round(buckets.current), d1_30:_round(buckets.b30), d31_60:_round(buckets.b60),
                d61_90:_round(buckets.b90), d90plus:_round(buckets.b90plus) },
      customers:customers,
      totalOutstanding:_round(customers.reduce(function(s,c){return s+c.outstanding;},0))
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// PERSONAL — separate personal-spend summary (never in business P&L)
// ============================================================
function dashPersonal(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters), month = _inMonth(filters);
    var rows = getPersonalTransactions({fy:fy});
    var byCat = {}, total = 0, fromCompany = 0;
    rows.forEach(function(r){
      if (month && r.month !== month) return;
      // Tier 5.2 mirror rows (Fund In / Fund Out / Internal) share the sheet
      // now — only Expense rows should count toward personal-spend totals.
      // Legacy rows (pre-Tier-5.1) have txnClass filled via the read adapter.
      if (r.txnClass && r.txnClass !== 'Expense') return;
      byCat[r.category] = (byCat[r.category]||0) + r.amount;
      total += r.amount;
      if (r.linkedRef) fromCompany += r.amount;   // paid from a company account (auto-drawing)
    });
    var categories = Object.keys(byCat).map(function(k){ return {category:k, amount:_round(byCat[k])}; })
      .sort(function(a,b){return b.amount-a.amount;});
    return _envOk({ fy:fy, month: month||'all', total:_round(total),
                    fromCompanyAccounts:_round(fromCompany), categories:categories });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 8b. RICHER DASHBOARD AGGREGATORS
//     All read getAllTransactions and respect {fy}. Transfers are
//     excluded from spend/revenue (they're not P&L).
// ============================================================

// Month-on-month sales vs expenses for the FY (Apr→Mar order).
//   Returns { months:[{month, sales, expenses, net}] }
function dashMonthly(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    // FY month order: Apr..Mar
    var startY = parseInt(_s(fy).split('-')[0],10);
    var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var seq = [3,4,5,6,7,8,9,10,11,0,1,2];
    var order = seq.map(function(m){ var y=(m>=3)?startY:startY+1; return MN[m]+'-'+String(y).slice(2); });
    var map = {};
    order.forEach(function(m){ map[m] = {month:m, sales:0, expenses:0, net:0}; });

    // Sales = Invoices this FY, bucketed by invoice date month. (Direct
    // 'Income' txns on the Transactions sheet are added too for the rare
    // cash-income case that bypasses invoicing.)
    var invSheet = _getSS().getSheetByName(AR.INVOICES);
    if (invSheet && invSheet.getLastRow() > 1){
      var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
      invSheet.getRange(2,1,invSheet.getLastRow()-1,iW).getValues().forEach(function(r){
        if (_s(r[IC.STATUS]) === INV_STATUS.CANCELLED) return;
        if (_s(r[IC.FY]) !== fy) return;
        var m = _monthOf(r[IC.INV_DATE]);
        if (map[m]) map[m].sales += _n(r[IC.AMOUNT]) || 0;
      });
    }
    // Returns reduce sales for the month they were booked in.
    var rSheet = _getSS().getSheetByName(AR.RETURNS);
    if (rSheet && rSheet.getLastRow() > 1){
      var rW = Math.min(rSheet.getLastColumn(), RC_TOTAL);
      rSheet.getRange(2,1,rSheet.getLastRow()-1,rW).getValues().forEach(function(r){
        if (_s(r[RC.FY]) !== fy) return;
        var m = _monthOf(r[RC.DATE]);
        if (map[m]) map[m].sales -= _n(r[RC.AMOUNT]) || 0;
      });
    }

    // Expenses = Bills (gross, bucketed by bill date month) + Direct Expense
    // txns NOT linked to a bill. No double-count with bills.
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    if (billsSheet && billsSheet.getLastRow() > 1){
      var bW = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
      billsSheet.getRange(2,1,billsSheet.getLastRow()-1,bW).getValues().forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (_s(r[BC.FY]) !== fy) return;
        var m = _monthOf(r[BC.BILL_DATE]);
        if (map[m]) map[m].expenses += _n(r[BC.AMOUNT]) || 0;
      });
    }
    getAllTransactions({fy:fy}).forEach(function(r){
      if (!map[r.month]) return;
      if (r.txnClass === 'Income') map[r.month].sales += r.amount || 0;
      else if (r.txnClass === 'Expense'){
        if (r.linkedRef && _s(r.linkedRef).indexOf('BILL-') === 0) return; // skip bill-linked
        map[r.month].expenses += r.amount || 0;
      }
    });

    var months = order.map(function(m){
      var o = map[m]; o.net = _round(o.sales - o.expenses);
      o.sales = _round(o.sales); o.expenses = _round(o.expenses);
      return o;
    });
    return _envOk({ fy:fy, months:months });
  } catch(e){ return _envErr(e); }
}

// Category-wise spend (expenses only). Walks the SAME three sources as dashPnl
// so drill-down totals reconcile with each group's total in P&L:
//   1. Transactions.Expense (direct expenses)      — uses TAXABLE-equivalent AMOUNT
//   2. Bills (vendor accrual)                      — uses TAXABLE || AMOUNT (ex-GST)
//   3. Bill_Payments with Applied_Bills empty      — un-billed vendor cash, lands
//                                                    in a synthetic '(Un-billed vendor payments)'
//                                                    category just like dashPnl.
//   Returns { categories:[{category, amount, group}] } sorted desc.
function dashByCategory(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters), month = _inMonth(filters);
    var catToGroup = _buildCatToGroupMapLive();  // live resolve
    var map = {};   // {category: {amount, group}}
    function _bump(cat, amt, group){
      var k = cat || '(uncategorised)';
      if (!map[k]) map[k] = {amount:0, group: group || catToGroup[k] || '(no group)'};
      map[k].amount += amt;
    }

    // 1) Direct expense transactions (same filter as dashPnl)
    getAllTransactions({fy:fy}).forEach(function(r){
      if (r.txnClass !== 'Expense') return;
      if (r.recovery === 'Pending' || r.recovery === 'Recovered') return;
      if (month && r.month !== month) return;
      _bump(r.category, r.amount);
    });

    // 2) Bills (accrual) — Category + taxable-first amount, same as dashPnl.
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    if (billsSheet && billsSheet.getLastRow() > 1){
      var bW = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
      billsSheet.getRange(2,1,billsSheet.getLastRow()-1,bW).getValues().forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (_s(r[BC.FY]) !== fy) return;
        if (month && _monthOf(r[BC.BILL_DATE]) !== month) return;
        var amt = _n(r[BC.TAXABLE]) || _n(r[BC.AMOUNT]) || 0;
        _bump(_s(r[BC.CATEGORY]), amt);
      });
    }

    // 3) Un-billed vendor payments — one synthetic bucket, matches dashPnl.
    var bpSheet = _getSS().getSheetByName(AP.PAYMENTS);
    var unbilled = 0;
    if (bpSheet && bpSheet.getLastRow() > 1){
      var pW = Math.min(bpSheet.getLastColumn(), BPMC_TOTAL);
      bpSheet.getRange(2,1,bpSheet.getLastRow()-1,pW).getValues().forEach(function(r){
        if (pW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        if (_s(r[BPMC.APPLIED])) return;              // has bill link → already counted in 2
        if (_s(r[BPMC.FY]) !== fy) return;
        if (month && _monthOf(r[BPMC.DATE]) !== month) return;
        unbilled += _n(r[BPMC.AMOUNT]) || 0;
      });
    }
    if (unbilled > 0) _bump('(Un-billed vendor payments)', unbilled, '(Un-billed vendor payments)');

    var categories = Object.keys(map).map(function(k){
      return {category:k, amount:_round(map[k].amount), group:map[k].group};
    }).sort(function(a,b){ return b.amount-a.amount; });
    return _envOk({ fy:fy, categories:categories });
  } catch(e){ return _envErr(e); }
}

// Project-wise spend + revenue.
//   Returns { projects:[{project, spend, revenue, net}] } sorted by spend desc
function dashByProject(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var ss = _getSS();
    var bucket = function(k){
      k = k || '(no project)';
      if (!map[k]) map[k] = {project:k, spend:0, revenue:0};
      return map[k];
    };
    var map = {};

    // Transactions (expense + income)
    getAllTransactions({fy:fy}).forEach(function(r){
      var b = bucket(r.project);
      if (r.txnClass === 'Expense') b.spend += r.amount;
      else if (r.txnClass === 'Income') b.revenue += r.amount;
    });

    // Bills — spend by bill's project
    var billsSheet = ss.getSheetByName(AP.BILLS);
    if (billsSheet && billsSheet.getLastRow() > 1){
      var bW = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
      billsSheet.getRange(2,1,billsSheet.getLastRow()-1,bW).getValues().forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (_s(r[BC.FY]) !== fy) return;
        var proj = bW > BC.PROJECT ? _s(r[BC.PROJECT]) : '';
        bucket(proj).spend += _n(r[BC.TAXABLE]) || _n(r[BC.AMOUNT]);
      });
    }

    // Un-billed bill payments (no Applied_Bills) — real spend
    var bpSheet = ss.getSheetByName(AP.PAYMENTS);
    if (bpSheet && bpSheet.getLastRow() > 1){
      var bpW = Math.min(bpSheet.getLastColumn(), BPMC_TOTAL);
      bpSheet.getRange(2,1,bpSheet.getLastRow()-1,bpW).getValues().forEach(function(r){
        if (bpW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        if (_s(r[BPMC.APPLIED])) return;
        if (_s(r[BPMC.FY]) !== fy) return;
        var proj = bpW > BPMC.PROJECT ? _s(r[BPMC.PROJECT]) : '';
        bucket(proj).spend += _n(r[BPMC.AMOUNT]);
      });
    }

    // Invoices — revenue by invoice's project
    var invSheet = ss.getSheetByName(AR.INVOICES);
    if (invSheet && invSheet.getLastRow() > 1){
      var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
      invSheet.getRange(2,1,invSheet.getLastRow()-1,iW).getValues().forEach(function(r){
        if (_s(r[IC.STATUS]) === INV_STATUS.CANCELLED) return;
        if (_s(r[IC.FY]) !== fy) return;
        var proj = iW > IC.PROJECT ? _s(r[IC.PROJECT]) : '';
        var val = _n(r[IC.TAXABLE]) || _n(r[IC.AMOUNT]);
        bucket(proj).revenue += val;
      });
    }

    // Returns — reduce revenue on the linked invoice's project
    var retSheet = ss.getSheetByName(AR.RETURNS);
    if (retSheet && retSheet.getLastRow() > 1 && invSheet){
      // Build Invoice_No → project map once
      var iW2 = Math.min(invSheet.getLastColumn(), IC_TOTAL);
      var invProjByNo = {};
      invSheet.getRange(2,1,invSheet.getLastRow()-1,iW2).getValues().forEach(function(r){
        var no = _s(r[IC.NO]);
        if (no) invProjByNo[no] = iW2 > IC.PROJECT ? _s(r[IC.PROJECT]) : '';
      });
      var rW = Math.min(retSheet.getLastColumn(), RC_TOTAL);
      retSheet.getRange(2,1,retSheet.getLastRow()-1,rW).getValues().forEach(function(r){
        if (_s(r[RC.FY]) !== fy) return;
        var proj = invProjByNo[_s(r[RC.INV_NO])] || '';
        bucket(proj).revenue -= _n(r[RC.AMOUNT]);
      });
    }

    var projects = Object.keys(map).map(function(k){
      var o = map[k]; o.net = _round(o.revenue - o.spend);
      o.spend = _round(o.spend); o.revenue = _round(o.revenue);
      return o;
    }).filter(function(o){ return o.spend !== 0 || o.revenue !== 0; })
      .sort(function(a,b){ return b.spend - a.spend; });
    return _envOk({ fy:fy, projects:projects });
  } catch(e){ return _envErr(e); }
}

// Employee-wise expense (who submitted / on whose behalf).
//   Returns { employees:[{employee, amount, count}] } sorted desc
function dashByEmployee(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var all = getAllTransactions({fy:fy});
    var map = {};
    all.forEach(function(r){
      if (r.txnClass !== 'Expense') return;
      // Prefer "on behalf of" if set, else who submitted
      var k = r.onBehalfOf || r.submittedBy || '(unknown)';
      if (k === 'Owner' || k === 'Company') return;   // skip owner/company-level rows
      if (!map[k]) map[k] = {employee:k, total:0, settled:0, pending:0, count:0};
      map[k].total += r.amount; map[k].count += 1;
      if (r.settleStatus === 'Settled')      map[k].settled += r.amount;
      else if (r.settleStatus === 'Pending') map[k].pending += r.amount;
    });
    var employees = Object.keys(map).map(function(k){
      return {employee:k, total:_round(map[k].total), settled:_round(map[k].settled),
              pending:_round(map[k].pending), count:map[k].count};
    }).sort(function(a,b){ return b.total - a.total; });
    return _envOk({ fy:fy, employees:employees });
  } catch(e){ return _envErr(e); }
}

// Top vendors by spend (expenses only).
//   Returns { vendors:[{party, amount, count}] } top N
// Top vendors by BILLED AMOUNT this FY — reads the Bills sheet (vendor ledger),
// NOT the Transactions sheet. This surfaces real vendors who issue bills, not
// one-off party-text spend from expense forms. Groups by Vendor_ID when present
// (dedupes name variants), falling back to Vendor_Name.
function dashTopVendors(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var limit = filters.limit || 10;
    var sheet = _getSS().getSheetByName(AP.BILLS);
    if (!sheet || sheet.getLastRow() < 2) return _envOk({ fy:fy, vendors:[] });
    var w = Math.min(sheet.getLastColumn(), BC_TOTAL);
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var map = {};
    data.forEach(function(r){
      if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
      if (_s(r[BC.FY]) !== fy) return;
      var vid  = _s(r[BC.VEND_ID]);
      var name = _s(r[BC.VEND_NAME]);
      var key  = vid || name;         // group by ID when we have one, else by name
      if (!key) return;
      if (!map[key]) map[key] = {vendorId:vid, party:name || vid, amount:0, count:0, outstanding:0};
      map[key].amount      += _n(r[BC.AMOUNT]) || 0;
      map[key].outstanding += _n(r[BC.OUTSTANDING]) || 0;
      map[key].count       += 1;
      // Prefer a non-empty name if the first row had the ID with a blank name
      if (!map[key].party && name) map[key].party = name;
    });
    var vendors = Object.keys(map).map(function(k){
      var v = map[k];
      return {vendorId:v.vendorId, party:v.party, amount:_round(v.amount),
              outstanding:_round(v.outstanding), count:v.count};
    }).sort(function(a,b){ return b.amount - a.amount; }).slice(0, limit);
    return _envOk({ fy:fy, vendors:vendors });
  } catch(e){ return _envErr(e); }
}

// Top customers by INVOICED AMOUNT this FY (net of returns) — reads the
// Invoices sheet. Groups by Customer_ID when present. Returns are subtracted
// per-customer so a customer with a big return doesn't rank artificially high.
function dashTopCustomers(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var limit = filters.limit || 10;
    var invSheet = _getSS().getSheetByName(AR.INVOICES);
    if (!invSheet || invSheet.getLastRow() < 2) return _envOk({ fy:fy, customers:[] });
    var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
    var iData = invSheet.getRange(2, 1, invSheet.getLastRow()-1, iW).getValues();
    var map = {};        // key → {customerId, party, amount, outstanding, count}
    var invCustOf = {};  // invoice_no → customer_key — used by the returns walk
    iData.forEach(function(r){
      if (_s(r[IC.STATUS]) === INV_STATUS.CANCELLED) return;
      if (_s(r[IC.FY]) !== fy) return;
      var cid  = _s(r[IC.CUST_ID]);
      var name = _s(r[IC.CUST_NAME]);
      var key  = cid || name;
      if (!key) return;
      if (!map[key]) map[key] = {customerId:cid, party:name || cid, amount:0, outstanding:0, count:0};
      map[key].amount      += _n(r[IC.AMOUNT]) || 0;
      map[key].outstanding += _n(r[IC.OUTSTANDING]) || 0;
      map[key].count       += 1;
      if (!map[key].party && name) map[key].party = name;
      invCustOf[_s(r[IC.NO])] = key;
    });
    // Returns reduce invoiced per customer (via the original invoice's customer)
    var rSheet = _getSS().getSheetByName(AR.RETURNS);
    if (rSheet && rSheet.getLastRow() > 1){
      var rW = Math.min(rSheet.getLastColumn(), RC_TOTAL);
      rSheet.getRange(2, 1, rSheet.getLastRow()-1, rW).getValues().forEach(function(r){
        if (_s(r[RC.FY]) !== fy) return;
        var key = invCustOf[_s(r[RC.INV_NO])];
        if (key && map[key]) map[key].amount -= _n(r[RC.AMOUNT]) || 0;
      });
    }
    var customers = Object.keys(map).map(function(k){
      var c = map[k];
      return {customerId:c.customerId, party:c.party, amount:_round(c.amount),
              outstanding:_round(c.outstanding), count:c.count};
    }).sort(function(a,b){ return b.amount - a.amount; }).slice(0, limit);
    return _envOk({ fy:fy, customers:customers });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// Employee-wise expense summary — totals + settled vs pending split per
// employee. Feeds the dashboard "Employee Expenses" card with a "Settle all"
// per-row action. Walks the Transactions sheet (Employee entry type only).
//   Returns [{employee, total, settled, pending, pendingCount, settledCount}]
// sorted pending desc so the owner sees who's owed the most first.
// ============================================================
function dashEmployeeExpensesSummary(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();
    var all = getAllTransactions({fy:fy}) || [];
    var map = {};
    all.forEach(function(r){
      if (r.txnClass !== 'Expense') return;
      if (r.entryType !== 'Employee') return;
      var k = r.submittedBy || r.onBehalfOf || '(unknown)';
      if (!map[k]) map[k] = {employee:k, total:0, settled:0, pending:0,
                              pendingCount:0, settledCount:0};
      var amt = r.amount || 0;
      map[k].total += amt;
      if (r.settleStatus === 'Settled'){
        map[k].settled += amt;
        map[k].settledCount += 1;
      } else if (r.settleStatus === 'Pending'){
        map[k].pending += amt;
        map[k].pendingCount += 1;
      }
    });
    var rows = Object.keys(map).map(function(k){
      var e = map[k];
      return {employee:e.employee, total:_round(e.total),
              settled:_round(e.settled), pending:_round(e.pending),
              pendingCount:e.pendingCount, settledCount:e.settledCount};
    }).sort(function(a,b){
      // Pending first, then total as tiebreaker
      if (b.pending !== a.pending) return b.pending - a.pending;
      return b.total - a.total;
    });
    var totals = rows.reduce(function(t,r){
      t.total += r.total; t.settled += r.settled; t.pending += r.pending;
      t.pendingCount += r.pendingCount; t.settledCount += r.settledCount;
      return t;
    }, {total:0, settled:0, pending:0, pendingCount:0, settledCount:0});
    return _envOk({fy:fy, rows:rows,
                   totals:{total:_round(totals.total), settled:_round(totals.settled),
                           pending:_round(totals.pending),
                           pendingCount:totals.pendingCount, settledCount:totals.settledCount}});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// Drill-through: payments this FY grouped by customer.
// Feeds the Payment Received tile's click-through modal.
// Returns [{customerId, party, amount, count}] sorted amount desc.
// ============================================================
function drillPaymentsByCustomer(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();
    var pSheet = _getSS().getSheetByName(AR.PAYMENTS);
    if (!pSheet || pSheet.getLastRow() < 2) return _envOk({fy:fy, rows:[]});
    var w = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
    var data = pSheet.getRange(2,1,pSheet.getLastRow()-1,w).getValues();
    var map = {};
    data.forEach(function(r){
      if (w > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
      if (_s(r[PMC.FY]) !== fy) return;
      var cid  = _s(r[PMC.CUST_ID]);
      var name = _s(r[PMC.CUST_NAME]);
      var key  = cid || name;
      if (!key) return;
      if (!map[key]) map[key] = {customerId:cid, party:name || cid, amount:0, count:0};
      map[key].amount += _n(r[PMC.AMOUNT]) || 0;
      map[key].count  += 1;
      if (!map[key].party && name) map[key].party = name;
    });
    var rows = Object.keys(map).map(function(k){
      var c = map[k];
      return {customerId:c.customerId, party:c.party, amount:_round(c.amount), count:c.count};
    }).sort(function(a,b){ return b.amount - a.amount; });
    return _envOk({fy:fy, rows:rows});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// PROJECT OVERVIEW — unified per-project view
//
// For a given project name, returns:
//   - summary: revenue, expenses, gross margin, totals
//   - expenses: all transactions where Project = this project (employee + owner)
//   - bills: all vendor bills tagged with this project
//   - invoices: all customer invoices tagged with this project + status
//   - payments: customer payments that applied to those invoices
//
// Note: customer payments aren't directly project-tagged; we link them via
// invoice number → project. Same for vendor bill payments.
// ============================================================
function getProjectOverview(projectName){
  try {
    var project = _s(projectName);
    if (!project) return _envOk({ project:'', summary:{}, expenses:[], bills:[], invoices:[], payments:[] });

    // 1. Expenses + Income transactions — walk Transactions across all FYs
    var expenses = [];
    var incomeTxns = [];   // Owner Income rows tagged to this project (rare — most income is via invoice)
    var txnSheets = (typeof _allTxnSheets === 'function') ? _allTxnSheets() : [];
    txnSheets.forEach(function(sheet){
      if (!sheet || sheet.getLastRow() < 2) return;
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,TC_TOTAL).getValues();
      rows.forEach(function(r){
        if (_s(r[TC.DELETED]) === 'Yes') return;
        if (_s(r[TC.PROJECT]) !== project) return;
        var klass = _s(r[TC.CLASS]);
        if (klass !== 'Expense' && klass !== 'Income') return;
        var rec = {
          id:          _s(r[TC.ID]),
          date:        Utilities.formatDate(_asDate(r[TC.TX_DATE]),'GMT+5:30','yyyy-MM-dd'),
          category:    _s(r[TC.CATEGORY]),
          group:       _s(r[TC.GROUP]),
          description: _s(r[TC.DESCRIPTION]),
          party:       _s(r[TC.PARTY_NAME]),
          submittedBy: _s(r[TC.SUBMITTED_BY]),
          amount:      _n(r[TC.AMOUNT]),
          settled:     _s(r[TC.SETTLE_STATUS]) === 'Settled'
        };
        if (klass === 'Expense') expenses.push(rec);
        else                     incomeTxns.push(rec);
      });
    });
    expenses.sort(function(a,b){ return b.date.localeCompare(a.date); });

    // 2. Bills — walk Bills sheet
    var bills = [];
    var matchedBillRefs = {};   // Bill_ID + Bill_No both indexed — Applied_Bills stores either.
    var bSheet = _getSS().getSheetByName(AP.BILLS);
    if (bSheet && bSheet.getLastRow() >= 2){
      var bW = Math.min(bSheet.getLastColumn(), BC_TOTAL);
      var bRows = bSheet.getRange(2,1,bSheet.getLastRow()-1,bW).getValues();
      bRows.forEach(function(r){
        if (bW <= BC.PROJECT || _s(r[BC.PROJECT]) !== project) return;
        if (_s(r[BC.STATUS]) === 'Cancelled') return;
        var bid = _s(r[BC.ID]);
        var bno = _s(r[BC.BILL_NO]);
        // Guard against empty keys: matchedBillRefs[""] plus applied.indexOf("")
        // was making every payment "match" every project. Only index non-empty refs.
        if (bid) matchedBillRefs[bid] = true;
        if (bno) matchedBillRefs[bno] = true;
        // Taxable = pre-GST value on the bill (the true economic cost for a
        // GST-registered business, since GST paid on inputs is claimable ITC).
        // If Taxable is 0/blank (older bills, non-GST cash purchases), fall
        // back to Total so this figure never under-reports spend.
        var billTaxable = _n(r[BC.TAXABLE]);
        if (billTaxable <= 0) billTaxable = _n(r[BC.AMOUNT]);
        bills.push({
          id:          bid,
          vendor:      _s(r[BC.VEND_NAME]),
          billNo:      bno,
          billDate:    Utilities.formatDate(_asDate(r[BC.BILL_DATE]),'GMT+5:30','yyyy-MM-dd'),
          amount:      _n(r[BC.AMOUNT]),
          taxable:     billTaxable,
          paid:        _n(r[BC.PAID]),
          outstanding: _n(r[BC.OUTSTANDING]),
          status:      _s(r[BC.STATUS])
        });
      });
    }
    bills.sort(function(a,b){ return b.billDate.localeCompare(a.billDate); });

    // 3. Invoices — walk Invoices sheet
    var invoices = [];
    var iSheet = _getSS().getSheetByName(AR.INVOICES);
    var matchedInvoiceNos = {};   // for payment lookup
    if (iSheet && iSheet.getLastRow() >= 2){
      var iW = Math.min(iSheet.getLastColumn(), IC_TOTAL);
      var iRows = iSheet.getRange(2,1,iSheet.getLastRow()-1,iW).getValues();
      iRows.forEach(function(r){
        if (iW <= IC.PROJECT || _s(r[IC.PROJECT]) !== project) return;
        if (_s(r[IC.STATUS]) === 'Cancelled') return;
        var no = _s(r[IC.NO]);
        // Guard: only index non-empty invoice numbers (same reason as bills above).
        if (no) matchedInvoiceNos[no] = true;
        invoices.push({
          no:          no,
          customer:    _s(r[IC.CUST_NAME]),
          invDate:     Utilities.formatDate(_asDate(r[IC.INV_DATE]),'GMT+5:30','yyyy-MM-dd'),
          amount:      _n(r[IC.AMOUNT]),
          paid:        _n(r[IC.PAID]),
          outstanding: _n(r[IC.OUTSTANDING]),
          status:      _s(r[IC.STATUS])
        });
      });
    }
    invoices.sort(function(a,b){ return b.invDate.localeCompare(a.invDate); });

    // 4. Payments — union of (a) payments whose Applied_Invoices reference an
    //    invoice on this project, and (b) payments directly tagged with the
    //    project (new Project column on Payments — for receipts without an invoice).
    var payments = [];
    var seenPay = {};
    var pSheet = _getSS().getSheetByName(AR.PAYMENTS);
    if (pSheet && pSheet.getLastRow() >= 2){
      var pW = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
      var pRows = pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues();
      pRows.forEach(function(r){
        if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
        var applied = _s(r[PMC.APPLIED]);
        // A payment counts against this project when EITHER it applied to one
        // of this project's invoices, OR (only when it has no applied invoices)
        // it was direct-tagged to this project. Invoices always win over the
        // tag — a payment applied to invoices is attributed by those invoices,
        // never by its own project field. Closes the double-count case where
        // the operator's tag differed from the applied invoice's project.
        //
        // Exact-match via parsed tokens. Substring matching (the old code) caused
        // two false-positive bugs: (a) short Invoice_Nos matched anywhere in
        // another payment's applied string, and (b) an empty key + `.indexOf("")`
        // returning 0 leaked every payment into every project.
        var linked = false;
        if (applied){
          var _pTokens = _parseAppliedString(applied);
          for (var _pi = 0; _pi < _pTokens.length; _pi++){
            var _rawNo = String(_pTokens[_pi].ref || '').split(/\s/)[0];
            if (_rawNo && matchedInvoiceNos[_rawNo]){ linked = true; break; }
          }
        }
        var direct = !applied && pW > PMC.PROJECT && _s(r[PMC.PROJECT]) === project;
        if (!linked && !direct) return;
        var id = _s(r[PMC.ID]);
        if (seenPay[id]) return; seenPay[id] = true;
        // Attribution:
        //   linked → per-invoice portion (parse Applied_Invoices, sum matching refs)
        //   direct (no applied invoices) → full payment amount
        var attributed;
        if (linked){
          attributed = 0;
          var chunks = _parseAppliedString(applied);
          for (var k=0;k<chunks.length;k++){
            var _cRef = String(chunks[k].ref || '').split(/\s/)[0];
            if (_cRef && matchedInvoiceNos[_cRef]) attributed += chunks[k].amount;
          }
        } else {
          attributed = _n(r[PMC.AMOUNT]);
        }
        payments.push({
          id:         id,
          customer:   _s(r[PMC.CUST_NAME]),
          date:       Utilities.formatDate(_asDate(r[PMC.DATE]),'GMT+5:30','yyyy-MM-dd'),
          amount:     _n(r[PMC.AMOUNT]),
          attributed: _round(attributed),
          mode:       _s(r[PMC.MODE]),
          applied:    applied || '(direct-tagged)'
        });
      });
    }
    payments.sort(function(a,b){ return b.date.localeCompare(a.date); });

    // 4b. Vendor bill-payments — union of (a) direct project-tagged and
    //     (b) applied to a bill of this project via Applied_Bills. Historical
    //     rows recorded before the Project column existed still surface via (b).
    var billPayments = [];
    var seenBP = {};
    var bpSheet = _getSS().getSheetByName(AP.PAYMENTS);
    if (bpSheet && bpSheet.getLastRow() >= 2){
      var bpW = Math.min(bpSheet.getLastColumn(), BPMC_TOTAL);
      var bpRows = bpSheet.getRange(2,1,bpSheet.getLastRow()-1,bpW).getValues();
      bpRows.forEach(function(r){
        if (bpW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        var applied = _s(r[BPMC.APPLIED]);
        // A payment counts against this project when EITHER it applied to one
        // of this project's bills, OR (only when it has no applied bills)
        // it was direct-tagged to this project. Bills always win over the tag —
        // a payment applied to bills is attributed by those bills, never by
        // its own project field. This closes the double-count case where the
        // operator's tag differed from the applied bill's project.
        //
        // Exact-match via parsed tokens. Substring matching (the old code) caused
        // two false-positive bugs: (a) short Bill_Nos like "1" or "5" matched
        // anywhere in another payment's applied string, and (b) an empty
        // Bill_No made matchedBillRefs[""] a key and `.indexOf("")` returns 0
        // so every payment in the system leaked into every project.
        var linked = false;
        if (applied){
          var _bpTokens = _parseAppliedString(applied);
          for (var _bpi = 0; _bpi < _bpTokens.length; _bpi++){
            var _rawId = String(_bpTokens[_bpi].ref || '').split(/\s/)[0];   // strip " (INV-234) [cancelled]" enrichment
            if (_rawId && matchedBillRefs[_rawId]){ linked = true; break; }
          }
        }
        var direct = !applied && bpW > BPMC.PROJECT && _s(r[BPMC.PROJECT]) === project;
        if (!linked && !direct) return;
        var id = _s(r[BPMC.ID]);
        if (seenBP[id]) return; seenBP[id] = true;
        // Attribution:
        //   linked → per-bill portion (parse Applied_Bills, sum matching refs)
        //   direct (no applied bills) → full payment amount
        var attributedBP;
        if (linked){
          attributedBP = 0;
          var bpChunks = _parseAppliedString(applied);
          for (var m=0;m<bpChunks.length;m++){
            var _bcRef = String(bpChunks[m].ref || '').split(/\s/)[0];
            if (_bcRef && matchedBillRefs[_bcRef]) attributedBP += bpChunks[m].amount;
          }
        } else {
          attributedBP = _n(r[BPMC.AMOUNT]);
        }
        billPayments.push({
          id:         id,
          vendor:     _s(r[BPMC.VEND_NAME]),
          date:       Utilities.formatDate(_asDate(r[BPMC.DATE]),'GMT+5:30','yyyy-MM-dd'),
          amount:     _n(r[BPMC.AMOUNT]),
          attributed: _round(attributedBP),
          mode:       _s(r[BPMC.MODE]),
          applied:    applied || '(direct-tagged)'
        });
      });
      billPayments.sort(function(a,b){ return b.date.localeCompare(a.date); });
    }

    // 5. Summary
    var invoiceRevenue = invoices.reduce(function(s,i){ return s + i.amount; }, 0);
    var directIncome   = incomeTxns.reduce(function(s,i){ return s + i.amount; }, 0);
    var revenue        = invoiceRevenue + directIncome;
    var billed         = bills.reduce(function(s,b){ return s + b.amount; }, 0);
    var billedTaxable  = bills.reduce(function(s,b){ return s + (b.taxable || b.amount); }, 0);
    var directExpense  = expenses.reduce(function(s,e){ return s + e.amount; }, 0);
    // Un-linked (direct-tagged) vendor payments per vendor.
    // A direct-tagged payment may be either (a) a settlement for an existing
    // unpaid bill of the same vendor the operator forgot to link — already
    // counted inside `billed`, or (b) a real cash purchase with no bill —
    // genuine new cost. Distinguish per vendor: excess of direct payment over
    // that vendor's open (outstanding) bill balance is new cost. Anything up
    // to open balance is treated as settlement (skip to avoid double count).
    var openByVendor = {};
    bills.forEach(function(b){
      openByVendor[b.vendor] = (openByVendor[b.vendor] || 0) + _n(b.outstanding);
    });
    var directByVendor = {};
    var directVendorPay = 0;   // total for display in detail line
    billPayments.forEach(function(p){
      if (p.applied !== '(direct-tagged)') return;
      directByVendor[p.vendor] = (directByVendor[p.vendor] || 0) + p.amount;
      directVendorPay += p.amount;
    });
    var directVendorPayExtra = 0;
    Object.keys(directByVendor).forEach(function(v){
      directVendorPayExtra += Math.max(0, directByVendor[v] - (openByVendor[v] || 0));
    });
    var totalCost = billed + directExpense + directVendorPayExtra;
    // Received = sum of amounts actually attributed to THIS project (not the
    // full payment amount when only part of it landed on a project invoice).
    var received  = payments.reduce(function(s,p){ return s + (p.attributed || p.amount); }, 0);
    var receivable = invoices.reduce(function(s,i){ return s + i.outstanding; }, 0);
    var payable    = bills.reduce(function(s,b){ return s + b.outstanding; }, 0);

    return _envOk({
      project: project,
      summary: {
        revenue:        _round(revenue),
        billed:         _round(billed),
        billedTaxable:  _round(billedTaxable),   // pre-GST value of project bills — real cost after ITC
        directExpense:  _round(directExpense),
        totalCost:      _round(totalCost),
        totalCostExGst: _round(billedTaxable + directExpense + directVendorPayExtra), // same shape as totalCost but with pre-GST bill values
        directVendorPayExtra: _round(directVendorPayExtra),   // portion of direct-tagged payments beyond open bills — real extra cost
        grossMargin:    _round(revenue - totalCost),
        received:       _round(received),
        receivable:     _round(receivable),
        payable:        _round(payable),
        expenseCount:      expenses.length,
        incomeCount:       incomeTxns.length,
        billCount:         bills.length,
        invoiceCount:      invoices.length,
        paymentCount:      payments.length,
        billPaymentCount:  billPayments.length,
        directIncome:      _round(directIncome),
        invoiceRevenue:    _round(invoiceRevenue),
        directVendorPay:   _round(directVendorPay)
      },
      expenses:     expenses,
      incomeTxns:   incomeTxns,
      bills:        bills,
      invoices:     invoices,
      payments:     payments,
      billPayments: billPayments
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// Dashboard analytics shipped with Batch A — consumed by the
// Dashboard rebuild in Batch B. Nothing calls them yet.
// ============================================================

// Total vendor bills ISSUED this FY (gross, Inc-GST). Not "paid".
function getBilledAmountFy(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();   // 'all' not meaningful here — fall back to current
    var sheet = _getSS().getSheetByName(AP.BILLS);
    if (!sheet || sheet.getLastRow() < 2) return _envOk({total:0, count:0, fy:fy});
    var w = Math.min(sheet.getLastColumn(), BC_TOTAL);
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var total = 0, cnt = 0;
    data.forEach(function(r){
      if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
      if (_s(r[BC.FY]) !== fy) return;
      total += _n(r[BC.AMOUNT]) || 0;
      cnt++;
    });
    return _envOk({total:_round(total), count:cnt, fy:fy});
  } catch(e){ return _envErr(e); }
}

// From-account × To-account grid of fund moves for this FY.
function getFundTransferMatrix(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();
    var sheet = _getSS().getSheetByName('Fund_Transfers');
    if (!sheet || sheet.getLastRow() < 2) return _envOk({matrix:{}, accounts:[], fy:fy});
    var w = Math.min(sheet.getLastColumn(), (typeof FT_TOTAL !== 'undefined' ? FT_TOTAL : 19));
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var matrix = {}, accSet = {};
    data.forEach(function(r){
      if (_s(r[FT.DELETED]) === 'Yes') return;
      if (_s(r[FT.FY]) !== fy) return;
      var from = _s(r[FT.FROM_ACCOUNT]);
      var to   = _s(r[FT.TO_ACCOUNT]);
      var amt  = _n(r[FT.AMOUNT]) || 0;
      if (!from || !to || amt <= 0) return;
      accSet[from] = true; accSet[to] = true;
      if (!matrix[from]) matrix[from] = {};
      matrix[from][to] = (matrix[from][to] || 0) + amt;
    });
    var accounts = Object.keys(accSet).sort();
    Object.keys(matrix).forEach(function(f){
      Object.keys(matrix[f]).forEach(function(t){ matrix[f][t] = _round(matrix[f][t]); });
    });
    return _envOk({matrix:matrix, accounts:accounts, fy:fy});
  } catch(e){ return _envErr(e); }
}

// Payment Received split by invoice FY — resolves the "received this FY but
// invoice was from prior FY" case so the operator can see the breakdown.
//   {total, currentFyInvoices, priorFyInvoices, unapplied, fy}
// currentFyInvoices  = received this FY against invoices from this FY
// priorFyInvoices    = received this FY against invoices from any prior FY
// unapplied          = received this FY not applied to any invoice (advance)
function getPaymentReceivedBreakdown(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();
    // Build invoice_no → invoice_fy lookup
    var invMap = {};
    var invSheet = _getSS().getSheetByName(AR.INVOICES);
    if (invSheet && invSheet.getLastRow() > 1){
      var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
      invSheet.getRange(2,1,invSheet.getLastRow()-1,iW).getValues().forEach(function(r){
        var no = _s(r[IC.NO]); if (no) invMap[no] = _s(r[IC.FY]);
      });
    }
    var pSheet = _getSS().getSheetByName(AR.PAYMENTS);
    if (!pSheet || pSheet.getLastRow() < 2){
      return _envOk({total:0, currentFyInvoices:0, priorFyInvoices:0, unapplied:0, fy:fy});
    }
    var pW = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
    var pData = pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues();
    var total = 0, currentFyInvs = 0, priorFyInvs = 0, unapplied = 0;
    pData.forEach(function(r){
      if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
      if (_s(r[PMC.FY]) !== fy) return;
      var amt = _n(r[PMC.AMOUNT]) || 0;
      total += amt;
      var applied = _parseAppliedString(_s(r[PMC.APPLIED]));
      if (!applied.length){ unapplied += amt; return; }
      // Sum applied-amounts and classify each by invoice FY
      var appliedSum = 0;
      applied.forEach(function(a){
        appliedSum += a.amount;
        var invFy = invMap[a.ref];
        if (!invFy)              unapplied   += a.amount;   // unknown invoice ref — treat as unapplied
        else if (invFy === fy)   currentFyInvs += a.amount;
        else                     priorFyInvs   += a.amount;
      });
      // Any residual between payment amount and applied-sum lands as unapplied
      // (operator credited part as advance). Rounding slack ignored under ₹1.
      var residual = amt - appliedSum;
      if (residual > 1) unapplied += residual;
    });
    // Reverse direction: payments received in PRIOR FY applied to CURRENT FY
    // invoices. These paid down current-FY invoice outstanding, but their cash
    // sits in last FY's books — so they don't appear in `total` above, and the
    // AR reconciliation is blind to them unless we surface them. Scan ALL
    // non-cancelled payments where payment FY != current and at least one
    // applied invoice has FY == current.
    var priorFyAgainstCurFy = 0;
    pData.forEach(function(r){
      if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
      if (_s(r[PMC.FY]) === fy) return;  // only prior-FY payments
      var applied = _parseAppliedString(_s(r[PMC.APPLIED]));
      applied.forEach(function(a){
        if (invMap[a.ref] === fy) priorFyAgainstCurFy += a.amount;
      });
    });

    return _envOk({
      total: _round(total),
      currentFyInvoices: _round(currentFyInvs),
      priorFyInvoices:   _round(priorFyInvs),
      unapplied:         _round(unapplied),
      // Payments landed in a prior FY but applied to a CURRENT FY invoice.
      // Reduces current-FY invoice outstanding without showing in `total`.
      priorFyAgainstCurFy: _round(priorFyAgainstCurFy),
      fy: fy
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// AR Reconciliation — explains the Customer Owes You tile math.
//   Invoiced = Received (this FY) + Received (prior FY against these) + Outstanding
// Also exposes Unapplied as a parallel bucket (customer advances for FUTURE
// invoices; they don't reduce current invoice outstanding).
// ============================================================
function getARReconciliation(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();

    // Invoiced (net of returns) — matches dashKpis.totalSales when no month/project filter.
    var invSheet = _getSS().getSheetByName(AR.INVOICES);
    var invoiced = 0, outstanding = 0, openCount = 0;
    var invMap = {};
    if (invSheet && invSheet.getLastRow() > 1){
      var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
      invSheet.getRange(2,1,invSheet.getLastRow()-1,iW).getValues().forEach(function(r){
        var no = _s(r[IC.NO]); if (no) invMap[no] = _s(r[IC.FY]);
        if (_s(r[IC.STATUS]) === INV_STATUS.CANCELLED) return;
        if (_s(r[IC.FY]) !== fy) return;
        invoiced += _n(r[IC.AMOUNT]) || 0;
        var out = _n(r[IC.OUTSTANDING]) || 0;
        if (out > 0){ outstanding += out; openCount++; }
      });
    }
    var rSheet = _getSS().getSheetByName(AR.RETURNS);
    if (rSheet && rSheet.getLastRow() > 1){
      var rW = Math.min(rSheet.getLastColumn(), RC_TOTAL);
      rSheet.getRange(2,1,rSheet.getLastRow()-1,rW).getValues().forEach(function(r){
        if (_s(r[RC.FY]) !== fy) return;
        invoiced -= _n(r[RC.AMOUNT]) || 0;
      });
    }

    // Payment classification — one walk, both directions.
    var pSheet = _getSS().getSheetByName(AR.PAYMENTS);
    var receivedCurFy = 0, receivedPriorFy = 0, receivedAgainstPrior = 0, unapplied = 0;
    if (pSheet && pSheet.getLastRow() > 1){
      var pW = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
      pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues().forEach(function(r){
        if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
        var payFy = _s(r[PMC.FY]);
        var amt   = _n(r[PMC.AMOUNT]) || 0;
        var applied = _parseAppliedString(_s(r[PMC.APPLIED]));

        if (payFy === fy){
          // Payment landed in current FY
          if (!applied.length){ unapplied += amt; return; }
          var sum = 0;
          applied.forEach(function(a){
            sum += a.amount;
            var invFy = invMap[a.ref];
            if (invFy === fy)        receivedCurFy        += a.amount;
            else if (invFy)          receivedAgainstPrior += a.amount;
            else                     unapplied            += a.amount;  // orphan ref
          });
          if (amt - sum > 1) unapplied += (amt - sum);
        } else if (payFy){
          // Payment landed in a PRIOR FY — only interested in applications to CURRENT FY invoices
          applied.forEach(function(a){
            if (invMap[a.ref] === fy) receivedPriorFy += a.amount;
          });
        }
      });
    }

    return _envOk({
      fy: fy,
      invoiced:              _round(invoiced),
      receivedCurFy:         _round(receivedCurFy),         // in this FY, against these invoices
      receivedPriorFy:       _round(receivedPriorFy),       // in prior FY, against these invoices (hidden chunk)
      receivedAgainstPrior:  _round(receivedAgainstPrior),  // in this FY, against OLD invoices
      outstanding:           _round(outstanding),
      openInvoices:          openCount,
      unapplied:             _round(unapplied),             // customer advances for future invoices
      // reconciliation check: invoiced ≈ receivedCurFy + receivedPriorFy + outstanding
      checkDiff:             _round(invoiced - receivedCurFy - receivedPriorFy - outstanding)
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// AP Reconciliation — explains the Bills Remaining tile math.
//   Billed (this FY) − Outstanding on these bills = Paid against them
// Also surfaces (a) cash this FY that went to prior-FY bills (inflates Bills Paid),
// and (b) prior-FY bills still open today (not in the tile but drag total AP).
// ============================================================
function getAPReconciliation(fy){
  try {
    fy = _resolveFy(fy, true) || _currentFY();

    // Bills walk — one pass, three buckets.
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    var billed = 0, outstandingCurFy = 0, outstandingPriorFy = 0;
    var priorUnpaidCount = 0, curBillCount = 0;
    var billFyByIdOrNo = {};  // key by Bill_ID AND by Bill_No — Applied_Bills may cite either
    if (billsSheet && billsSheet.getLastRow() > 1){
      var bW = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
      billsSheet.getRange(2,1,billsSheet.getLastRow()-1,bW).getValues().forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        var bFy  = _s(r[BC.FY]);
        var bId  = _s(r[BC.ID]);
        var bNo  = _s(r[BC.BILL_NO]);
        var out  = _n(r[BC.OUTSTANDING]) || 0;
        if (bId) billFyByIdOrNo[bId] = bFy;
        if (bNo) billFyByIdOrNo[bNo] = bFy;
        if (bFy === fy){
          billed += _n(r[BC.AMOUNT]) || 0;
          curBillCount++;
          if (out > 0) outstandingCurFy += out;
        } else if (out > 0){
          outstandingPriorFy += out;
          priorUnpaidCount++;
        }
      });
    }

    // Bill-payments walk — classify this FY's cash by the bill's FY.
    var bpSheet = _getSS().getSheetByName(AP.PAYMENTS);
    var totalPaidCash = 0, paidToCurFyBills = 0, paidToPriorFyBills = 0, paidUnapplied = 0;
    if (bpSheet && bpSheet.getLastRow() > 1){
      var bpW = Math.min(bpSheet.getLastColumn(), BPMC_TOTAL);
      bpSheet.getRange(2,1,bpSheet.getLastRow()-1,bpW).getValues().forEach(function(r){
        if (bpW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
        if (_s(r[BPMC.FY]) !== fy) return;
        var amt = _n(r[BPMC.AMOUNT]) || 0;
        totalPaidCash += amt;
        var applied = _parseAppliedString(_s(r[BPMC.APPLIED]));
        if (!applied.length){ paidUnapplied += amt; return; }
        var sum = 0;
        applied.forEach(function(a){
          sum += a.amount;
          var bFy = billFyByIdOrNo[a.ref];
          if (bFy === fy)   paidToCurFyBills   += a.amount;
          else if (bFy)     paidToPriorFyBills += a.amount;
          else              paidUnapplied      += a.amount;
        });
        if (amt - sum > 1) paidUnapplied += (amt - sum);
      });
    }

    return _envOk({
      fy: fy,
      billed:              _round(billed),
      billCount:           curBillCount,
      outstandingCurFy:    _round(outstandingCurFy),    // = Bills Remaining tile
      outstandingPriorFy:  _round(outstandingPriorFy),  // prior-FY bills still open
      priorUnpaidCount:    priorUnpaidCount,
      totalPaidCash:       _round(totalPaidCash),       // = Bills Paid in dashKpis
      paidToCurFyBills:    _round(paidToCurFyBills),
      paidToPriorFyBills:  _round(paidToPriorFyBills),  // hidden chunk inside Bills Paid
      paidUnapplied:       _round(paidUnapplied)
    });
  } catch(e){ return _envErr(e); }
}