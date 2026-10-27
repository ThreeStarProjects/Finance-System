// ============================================================
// 3 Star Projects — Unified Finance System
// MODULE: Monthly Close  (P&L + Cash Flow + Balance Sheet, per month)
// File: MonthlyClose.gs
//
// One row per FY + month in M_MonthlyClose. Auto-fills derivable
// figures from your transactions / invoices / bills; the rest are
// manual entries you save.
//
// Hybrid behaviour on reopen:
//   - auto-filled fields are refreshed from the latest transactions
//   - manual fields stay at whatever was last saved
// ============================================================

var MC = {
  SHEET:        'M_MonthlyClose',
  VIEW:         'MonthlyClose_View'
};

// Flat schema — column order = sheet order. Don't reorder; append only.
var MCC = {
  FY:0, MONTH:1, MONTH_SEQ:2, STATUS:3,
  // Sales & P&L (manual values you might override, auto = suggestions)
  SALES:4, OPENING_STOCK:5, PURCHASES:6, CLOSING_STOCK:7, COGS:8, GP:9, OPEX:10,
  EBITDA:11, DEPRECIATION:12, INTEREST:13, TAX:14, PAT:15,
  // Cash Flow
  OPENING_CASH:16, ADV_RECEIVED:17, NEW_SALES_CASH:18, FROM_DEBTORS:19,
  BANK_LOAN_IN:20, CAPITAL_IN:21, BIZ_OPS_OUT:22, LOAN_REPAY:23, PERSONAL_CAPEX:24,
  NET_CASH_FLOW:25, CLOSING_CASH:26,
  // Balance Sheet
  AR:27, CASH_BANK:28, CLOSING_INV:29, LOANS_GIVEN:30, TAX_CREDIT:31, PREPAID:32,
  AP:33, LOANS_TAKEN:34, SECURITY_DEPOSIT:35, TAX_PAYABLE:36,
  NWC:37, CC_LIMIT:38, CC_USED:39,
  // Meta
  LAST_SAVED:40, LAST_SAVED_BY:41, NOTES:42
};
var MCC_TOTAL = 43;

var MC_HEADERS = [
  'FY','Month','Month_Seq','Status',
  // P&L block (* = auto-suggested by default)
  'Sales*','Opening_Stock','Purchases','Closing_Stock','COGS_calc','Gross_Profit_calc','OPEX*',
  'EBITDA_calc','Depreciation','Interest','Tax','PAT_calc',
  // Cash Flow block
  'Opening_Cash','Advance_Received','New_Sales_Cash','From_Debtors*',
  'Bank_Loan_In','Capital_In','Business_Ops_Out*','Loan_Repay','Personal_CapEx',
  'Net_Cash_Flow_calc','Closing_Cash_calc',
  // Balance Sheet block
  'AR*','Cash_Bank','Closing_Inventory','Loans_Given','Tax_Credit','Prepaid',
  'AP*','Loans_Taken','Security_Deposits','Tax_Payable',
  'Net_Working_Capital_calc','CC_Limit','CC_Used',
  // Meta
  'Last_Saved','Last_Saved_By','Notes'
];
var MC_WIDTHS = (function(){
  var w = []; for (var i=0;i<MC_HEADERS.length;i++) w.push(120); w[1]=110; w[3]=110; w[42]=300;
  return w;
})();

// Months in FY order (Apr..Mar)
var _FY_MONTH_SEQ_NAMES = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];

// ============================================================
// Setup — idempotent. Wired into setupMasters in Masters.gs.
// ============================================================
function setupMonthlyClose(){
  var ss = _getSS();
  var sheet = _buildSheet(ss, MC.SHEET, MC_HEADERS, MC_WIDTHS);
  // Status validator
  _validateColSoft(sheet, MCC.STATUS+1, ['Draft','Saved','Locked']);
  // Build (or rebuild) the pivoted view sheet
  setupMonthlyCloseView(ss);
  return { sheet: sheet.getName(), view: MC.VIEW };
}

// ============================================================
// FY OVERVIEW — list all 12 months for the FY with status + headline figures.
//   Returns { ok, data:{ fy, months:[{month, monthSeq, status, sales, gp, gpPct, ebitda, pat, cashFlow, workingCapital}] } }
// ============================================================
function getFyOverview(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var sheet = _getSS().getSheetByName(MC.SHEET);
    if (!sheet){ setupMonthlyClose(); sheet = _getSS().getSheetByName(MC.SHEET); }
    var saved = {};
    if (_hasData(sheet)){
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,MCC_TOTAL).getValues();
      for (var i=0;i<rows.length;i++){
        if (_s(rows[i][MCC.FY]) === fy){
          saved[_s(rows[i][MCC.MONTH])] = rows[i];
        }
      }
    }
    var startY = parseInt(_s(fy).split('-')[0],10);
    var nowMonth = _currentMonth();
    var months = _FY_MONTH_SEQ_NAMES.map(function(mn, i){
      var y = (i < 9) ? startY : startY+1;
      var mStr = mn + '-' + String(y).slice(2);
      var row = saved[mStr];
      var status = !row ? 'No data' : (_s(row[MCC.STATUS]) || 'Saved');
      var isCurrent = (mStr === nowMonth);
      if (status === 'No data' && isCurrent) status = 'Current';
      var sales    = row ? _n(row[MCC.SALES])  : 0;
      var gp       = row ? _n(row[MCC.GP])     : 0;
      var ebitda   = row ? _n(row[MCC.EBITDA]) : 0;
      var pat      = row ? _n(row[MCC.PAT])    : 0;
      var cashFlow = row ? _n(row[MCC.NET_CASH_FLOW]) : 0;
      var nwc      = row ? _n(row[MCC.NWC])    : 0;
      var gpPct    = sales > 0 ? _round1(gp/sales*100) : 0;
      return {
        month: mStr, monthSeq: i+1, status: status, isCurrent: isCurrent, hasData: !!row,
        sales: _round(sales), grossProfit: _round(gp), gpPct: gpPct,
        ebitda: _round(ebitda), pat: _round(pat),
        cashFlow: _round(cashFlow), workingCapital: _round(nwc)
      };
    });
    // Summary totals across whole FY
    var totals = months.reduce(function(t, m){
      t.sales += m.sales; t.grossProfit += m.grossProfit; t.ebitda += m.ebitda; t.pat += m.pat;
      return t;
    }, {sales:0, grossProfit:0, ebitda:0, pat:0});
    Object.keys(totals).forEach(function(k){ totals[k] = _round(totals[k]); });
    return _envOk({ fy: fy, months: months, totals: totals });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// LOAD A MONTH — returns saved values merged with fresh auto-fills.
//   data = {fy, month}    month = 'Jun-26' style
//   Returns { ok, data:{
//     fy, month, status, saved:bool,
//     auto: {sales, fromDebtors, opex, businessOpsOut, ar, ap, cashBank},
//     manual: { … all manual fields from saved row, or zeros … },
//     calc:   { cogs, gp, ebitda, pat, netCashFlow, closingCash, nwc }
//   }}
// ============================================================
function loadMonthlyClose(input){
  try {
    input = input || {};
    var fy = _s(input.fy) || _currentFY();
    var month = _s(input.month) || _currentMonth();

    // Locate saved row, if any
    var sheet = _getSS().getSheetByName(MC.SHEET);
    if (!sheet){ setupMonthlyClose(); sheet = _getSS().getSheetByName(MC.SHEET); }
    var saved = null, savedRowIdx = -1;
    if (_hasData(sheet)){
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,MCC_TOTAL).getValues();
      for (var i=0;i<rows.length;i++){
        if (_s(rows[i][MCC.FY]) === fy && _s(rows[i][MCC.MONTH]) === month){
          saved = rows[i]; savedRowIdx = i+2; break;
        }
      }
    }

    // Auto-fill from transactions/invoices/bills
    var auto = _autoFillForMonth(fy, month);

    // Manual fields (use saved values if present, else zeros, except auto fields where saved overrides auto)
    var manual = {
      openingStock:     saved ? _n(saved[MCC.OPENING_STOCK])   : 0,
      purchases:        saved ? _n(saved[MCC.PURCHASES])       : 0,
      closingStock:     saved ? _n(saved[MCC.CLOSING_STOCK])   : 0,
      depreciation:     saved ? _n(saved[MCC.DEPRECIATION])    : 0,
      interest:         saved ? _n(saved[MCC.INTEREST])        : 0,
      tax:              saved ? _n(saved[MCC.TAX])             : 0,
      openingCash:      saved ? _n(saved[MCC.OPENING_CASH])    : 0,
      advReceived:      saved ? _n(saved[MCC.ADV_RECEIVED])    : 0,
      newSalesCash:     saved ? _n(saved[MCC.NEW_SALES_CASH])  : 0,
      bankLoanIn:       saved ? _n(saved[MCC.BANK_LOAN_IN])    : 0,
      capitalIn:        saved ? _n(saved[MCC.CAPITAL_IN])      : 0,
      loanRepay:        saved ? _n(saved[MCC.LOAN_REPAY])      : 0,
      personalCapex:    saved ? _n(saved[MCC.PERSONAL_CAPEX])  : 0,
      cashBank:         saved ? _n(saved[MCC.CASH_BANK])       : auto.cashBank,
      closingInventory: saved ? _n(saved[MCC.CLOSING_INV])     : 0,
      loansGiven:       saved ? _n(saved[MCC.LOANS_GIVEN])     : 0,
      taxCredit:        saved ? _n(saved[MCC.TAX_CREDIT])      : 0,
      prepaid:          saved ? _n(saved[MCC.PREPAID])         : 0,
      loansTaken:       saved ? _n(saved[MCC.LOANS_TAKEN])     : 0,
      securityDeposit:  saved ? _n(saved[MCC.SECURITY_DEPOSIT]): 0,
      taxPayable:       saved ? _n(saved[MCC.TAX_PAYABLE])     : 0,
      ccLimit:          saved ? _n(saved[MCC.CC_LIMIT])        : 0,
      ccUsed:           saved ? _n(saved[MCC.CC_USED])         : 0,
      notes:            saved ? _s(saved[MCC.NOTES])           : ''
    };

    // Compute live figures from auto + manual
    var calc = _computeMonthlyCalc(auto, manual);
    return _envOk({
      fy: fy, month: month,
      status: saved ? (_s(saved[MCC.STATUS]) || 'Saved') : 'Draft',
      saved: !!saved,
      lastSaved: saved ? saved[MCC.LAST_SAVED] : '',
      auto: auto, manual: manual, calc: calc
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// SAVE A MONTH — writes/updates row.
//   data = { fy, month, manual:{...}, status }
//   Auto-filled values are re-derived at save time and persisted so the
//   FY overview totals are correct without recomputing every read.
// ============================================================
function saveMonthlyClose(input){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    input = input || {};
    var fy = _s(input.fy) || _currentFY();
    var month = _s(input.month) || _currentMonth();
    var manual = input.manual || {};
    var status = _s(input.status) || 'Saved';

    var auto = _autoFillForMonth(fy, month);
    var calc = _computeMonthlyCalc(auto, manual);
    var monthSeq = _FY_MONTH_SEQ_NAMES.indexOf(_s(month).split('-')[0]) + 1;

    var sheet = _getSS().getSheetByName(MC.SHEET);
    if (!sheet){ setupMonthlyClose(); sheet = _getSS().getSheetByName(MC.SHEET); }

    var rowIdx = -1;
    if (_hasData(sheet)){
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,MCC_TOTAL).getValues();
      for (var i=0;i<rows.length;i++){
        if (_s(rows[i][MCC.FY]) === fy && _s(rows[i][MCC.MONTH]) === month){
          rowIdx = i+2; break;
        }
      }
    }

    var row = new Array(MCC_TOTAL).fill('');
    row[MCC.FY]               = fy;
    row[MCC.MONTH]            = month;
    row[MCC.MONTH_SEQ]        = monthSeq;
    row[MCC.STATUS]           = status;
    row[MCC.SALES]            = _round(auto.sales);
    row[MCC.OPENING_STOCK]    = _round(_n(manual.openingStock));
    row[MCC.PURCHASES]        = _round(_n(manual.purchases));
    row[MCC.CLOSING_STOCK]    = _round(_n(manual.closingStock));
    row[MCC.COGS]             = _round(calc.cogs);
    row[MCC.GP]               = _round(calc.gp);
    row[MCC.OPEX]             = _round(auto.opex);
    row[MCC.EBITDA]           = _round(calc.ebitda);
    row[MCC.DEPRECIATION]     = _round(_n(manual.depreciation));
    row[MCC.INTEREST]         = _round(_n(manual.interest));
    row[MCC.TAX]              = _round(_n(manual.tax));
    row[MCC.PAT]              = _round(calc.pat);
    row[MCC.OPENING_CASH]     = _round(_n(manual.openingCash));
    row[MCC.ADV_RECEIVED]     = _round(_n(manual.advReceived));
    row[MCC.NEW_SALES_CASH]   = _round(_n(manual.newSalesCash));
    row[MCC.FROM_DEBTORS]     = _round(auto.fromDebtors);
    row[MCC.BANK_LOAN_IN]     = _round(_n(manual.bankLoanIn));
    row[MCC.CAPITAL_IN]       = _round(_n(manual.capitalIn));
    row[MCC.BIZ_OPS_OUT]      = _round(auto.businessOpsOut);
    row[MCC.LOAN_REPAY]       = _round(_n(manual.loanRepay));
    row[MCC.PERSONAL_CAPEX]   = _round(_n(manual.personalCapex));
    row[MCC.NET_CASH_FLOW]    = _round(calc.netCashFlow);
    row[MCC.CLOSING_CASH]     = _round(calc.closingCash);
    row[MCC.AR]               = _round(auto.ar);
    row[MCC.CASH_BANK]        = _round(_n(manual.cashBank));
    row[MCC.CLOSING_INV]      = _round(_n(manual.closingInventory));
    row[MCC.LOANS_GIVEN]      = _round(_n(manual.loansGiven));
    row[MCC.TAX_CREDIT]       = _round(_n(manual.taxCredit));
    row[MCC.PREPAID]          = _round(_n(manual.prepaid));
    row[MCC.AP]               = _round(auto.ap);
    row[MCC.LOANS_TAKEN]      = _round(_n(manual.loansTaken));
    row[MCC.SECURITY_DEPOSIT] = _round(_n(manual.securityDeposit));
    row[MCC.TAX_PAYABLE]      = _round(_n(manual.taxPayable));
    row[MCC.NWC]              = _round(calc.nwc);
    row[MCC.CC_LIMIT]         = _round(_n(manual.ccLimit));
    row[MCC.CC_USED]          = _round(_n(manual.ccUsed));
    row[MCC.LAST_SAVED]       = new Date();
    row[MCC.LAST_SAVED_BY]    = (function(){ try { return Session.getActiveUser().getEmail() || ''; } catch(e){ return ''; } })();
    row[MCC.NOTES]            = _s(manual.notes);

    if (rowIdx > 0){
      sheet.getRange(rowIdx, 1, 1, MCC_TOTAL).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return {success:true, message:'Saved ' + fy + ' / ' + month, fy:fy, month:month, calc:calc};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// QUARTERLY / YEARLY ROLLUPS  (computed from saved monthly rows)
//   getQuarterlyRollup({fy})  -> { quarters:[{q, label, sales, gp, ...}] }
//   getYearlyRollup({fyList}) -> { years:[{fy, sales, gp, ...}] }
// ============================================================
function getQuarterlyRollup(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var sheet = _getSS().getSheetByName(MC.SHEET);
    var byMonth = {};
    if (sheet && _hasData(sheet)){
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,MCC_TOTAL).getValues();
      rows.forEach(function(r){
        if (_s(r[MCC.FY]) === fy) byMonth[_s(r[MCC.MONTH])] = r;
      });
    }
    var startY = parseInt(_s(fy).split('-')[0],10);
    var quarters = [];
    for (var q=0; q<4; q++){
      var qMonths = _FY_MONTH_SEQ_NAMES.slice(q*3, q*3+3);
      var qLabel = qMonths[0] + '–' + qMonths[2] + ' (Q' + (q+1) + ')';
      var acc = {sales:0, gp:0, opex:0, ebitda:0, pat:0, cashFlow:0};
      qMonths.forEach(function(mn, idx){
        var globalIdx = q*3 + idx;
        var y = (globalIdx < 9) ? startY : startY+1;
        var mStr = mn + '-' + String(y).slice(2);
        var r = byMonth[mStr];
        if (!r) return;
        acc.sales    += _n(r[MCC.SALES]);
        acc.gp       += _n(r[MCC.GP]);
        acc.opex     += _n(r[MCC.OPEX]);
        acc.ebitda   += _n(r[MCC.EBITDA]);
        acc.pat      += _n(r[MCC.PAT]);
        acc.cashFlow += _n(r[MCC.NET_CASH_FLOW]);
      });
      Object.keys(acc).forEach(function(k){ acc[k] = _round(acc[k]); });
      acc.q = q+1; acc.label = qLabel;
      acc.gpPct = acc.sales > 0 ? _round1(acc.gp/acc.sales*100) : 0;
      quarters.push(acc);
    }
    return _envOk({ fy:fy, quarters:quarters });
  } catch(e){ return _envErr(e); }
}

function getYearlyRollup(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(MC.SHEET);
    var byFy = {};
    if (sheet && _hasData(sheet)){
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,MCC_TOTAL).getValues();
      rows.forEach(function(r){
        var fy = _s(r[MCC.FY]);
        if (!fy) return;
        if (!byFy[fy]) byFy[fy] = {fy:fy, sales:0, gp:0, opex:0, ebitda:0, pat:0, cashFlow:0, monthsEntered:0};
        byFy[fy].sales    += _n(r[MCC.SALES]);
        byFy[fy].gp       += _n(r[MCC.GP]);
        byFy[fy].opex     += _n(r[MCC.OPEX]);
        byFy[fy].ebitda   += _n(r[MCC.EBITDA]);
        byFy[fy].pat      += _n(r[MCC.PAT]);
        byFy[fy].cashFlow += _n(r[MCC.NET_CASH_FLOW]);
        byFy[fy].monthsEntered++;
      });
    }
    var years = Object.keys(byFy).sort().map(function(fy){
      var y = byFy[fy];
      Object.keys(y).forEach(function(k){ if (typeof y[k] === 'number' && k!=='monthsEntered') y[k] = _round(y[k]); });
      y.gpPct = y.sales > 0 ? _round1(y.gp/y.sales*100) : 0;
      return y;
    });
    return _envOk({ years: years });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// AUTO-FILL  — what we can derive from the existing data layer
//   Returns {sales, opex, fromDebtors, businessOpsOut, ar, ap, cashBank}
// ============================================================
function _autoFillForMonth(fy, month){
  var auto = {sales:0, opex:0, fromDebtors:0, businessOpsOut:0, ar:0, ap:0, cashBank:0};
  try {
    var all = getAllTransactions({fy:fy});
    all.forEach(function(r){
      if (_s(r.month) !== month) return;
      if (r.txnClass === 'Income')  auto.sales += r.amount;
      if (r.txnClass === 'Expense') {
        auto.opex += r.amount;
        // "Business operations out" — settled (cash) expenses, excluding personal/capex
        if (r.settleStatus === 'Settled') auto.businessOpsOut += r.amount;
      }
    });
  } catch(e){ Logger.log('autoFill txn: '+e.message); }

  // From debtors — sum of payments received in this month against invoices
  try {
    var ss = _getSS();
    var paySheet = ss.getSheetByName('Payments');
    if (paySheet && _hasData(paySheet) && typeof PMC !== 'undefined'){
      var rows = paySheet.getRange(2,1,paySheet.getLastRow()-1,PMC_TOTAL).getValues();
      rows.forEach(function(r){
        var d = _asDate(r[PMC.DATE]);
        if (_monthOf(d) === month) auto.fromDebtors += _n(r[PMC.AMOUNT]);
      });
    }
  } catch(e){ Logger.log('autoFill debtors: '+e.message); }

  // AR snapshot — total outstanding across all invoices (point-in-time)
  try {
    var arRes = dashArAgeing({fy:fy});
    if (arRes && arRes.ok && arRes.data && arRes.data.buckets){
      var b = arRes.data.buckets;
      auto.ar = _n(b.current) + _n(b.d1_30) + _n(b.d31_60) + _n(b.d61_90) + _n(b.d90plus);
    }
  } catch(e){ Logger.log('autoFill ar: '+e.message); }

  // AP snapshot — total outstanding from Bills module
  try {
    if (typeof dashApAgeing === 'function'){
      var apRes = dashApAgeing();
      if (apRes && apRes.ok && apRes.data){
        auto.ap = _n(apRes.data.totalOutstanding);
      }
    }
  } catch(e){ Logger.log('autoFill ap: '+e.message); }

  return {
    sales: _round(auto.sales),
    opex: _round(auto.opex),
    fromDebtors: _round(auto.fromDebtors),
    businessOpsOut: _round(auto.businessOpsOut),
    ar: _round(auto.ar),
    ap: _round(auto.ap),
    cashBank: _round(auto.cashBank)
  };
}

// ============================================================
// LIVE CALCULATIONS  (P&L formulas + cash flow + working capital)
// ============================================================
function _computeMonthlyCalc(auto, manual){
  var sales = _n(auto.sales);
  var cogs = _n(manual.openingStock) + _n(manual.purchases) - _n(manual.closingStock);
  var gp   = sales - cogs;
  var opex = _n(auto.opex);
  var ebitda = gp - opex;
  var pat = ebitda - _n(manual.depreciation) - _n(manual.interest) - _n(manual.tax);

  var inflow  = _n(manual.advReceived) + _n(manual.newSalesCash) + _n(auto.fromDebtors)
              + _n(manual.bankLoanIn) + _n(manual.capitalIn);
  var outflow = _n(auto.businessOpsOut) + _n(manual.loanRepay) + _n(manual.personalCapex);
  var netCashFlow = inflow - outflow;
  var closingCash = _n(manual.openingCash) + netCashFlow;

  var tca = _n(auto.ar) + _n(manual.cashBank) + _n(manual.closingInventory)
          + _n(manual.loansGiven) + _n(manual.taxCredit) + _n(manual.prepaid);
  var tcl = _n(auto.ap) + _n(manual.loansTaken) + _n(manual.securityDeposit) + _n(manual.taxPayable);
  var nwc = tca - tcl;

  return {
    cogs: _round(cogs), gp: _round(gp), ebitda: _round(ebitda), pat: _round(pat),
    gpPct:     sales > 0 ? _round1(gp/sales*100) : 0,
    ebitdaPct: sales > 0 ? _round1(ebitda/sales*100) : 0,
    patPct:    sales > 0 ? _round1(pat/sales*100) : 0,
    inflow: _round(inflow), outflow: _round(outflow),
    netCashFlow: _round(netCashFlow), closingCash: _round(closingCash),
    tca: _round(tca), tcl: _round(tcl), nwc: _round(nwc)
  };
}

// _round override that accepts decimals param (existing _round uses 2 by default — re-export here)
// Note: the global _round in Invoices.gs takes one arg. We add a thin helper for percentages.
function _round1(n){ try { return parseFloat((parseFloat(n)||0).toFixed(1)); } catch(e){ return 0; } }

// ============================================================
// MONTHLY CLOSE VIEW — pivoted presentation sheet.
//   Row = line item.  Column = month.  Values are formulas
//   that read from M_MonthlyClose so the view always reflects
//   the latest saved figures.
//
//   Cell B1 = FY picker (data validation against M_MonthlyClose FY column).
//   Default = current FY.  Change B1 to switch the FY shown.
//
//   Idempotent: rebuilt on every setup. No data lives here.
// ============================================================
function setupMonthlyCloseView(ss){
  ss = ss || _getSS();
  var sheet = ss.getSheetByName(MC.VIEW);
  if (!sheet) sheet = ss.insertSheet(MC.VIEW);
  sheet.clear();
  try { sheet.clearConditionalFormatRules(); } catch(e){}
  // Clear any inherited / leftover data validations across the whole sheet.
  // (B1 validation gets re-applied below.)
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); } catch(e){}

  // ── FY picker (row 1) ──
  sheet.getRange('A1').setValue('FY');
  sheet.getRange('B1').setValue(_currentFY());
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setFontSize(12);
  sheet.setRowHeight(1, 32);
  // FY validation: a hard-coded list around the current FY so the picker
  // always has meaningful options even before any month has been saved.
  try {
    var nowY = new Date().getFullYear();
    var startMonth = new Date().getMonth();
    var curStart = (startMonth >= 3) ? nowY : nowY - 1;
    var fyOptions = [];
    for (var y = curStart - 2; y <= curStart + 1; y++){
      fyOptions.push(y + '-' + String(y+1).substring(2));
    }
    var fyRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(fyOptions, true)
      .setAllowInvalid(true).build();
    sheet.getRange('B1').setDataValidation(fyRule);
  } catch(e){}

  // ── Month headers (row 3) ──
  sheet.getRange('A3').setValue('Line Item').setFontWeight('bold');
  for (var i=0; i<12; i++){
    var col = i + 2;
    // Build month label: Apr–Dec use FY start year, Jan–Mar use FY start year + 1.
    // TEXT() forces the year to a string so RIGHT() works correctly across all columns.
    var formula = '=IFERROR("' + _FY_MONTH_SEQ_NAMES[i] +
      '-"&RIGHT(TEXT(IF(' + i + '<9,VALUE(LEFT($B$1,4)),VALUE(LEFT($B$1,4))+1),"0"),2),"")';
    sheet.getRange(3, col).setFormula(formula);
  }
  sheet.getRange(3, 1, 1, 13).setFontWeight('bold').setBackground('#E6E6E6')
    .setHorizontalAlignment('center');
  sheet.getRange('A3').setHorizontalAlignment('left');

  // ── Line items definition ──
  // Each row: {section: 'HEADER'} for a section divider,
  //         | {label, col(1-indexed), strong:true if highlighted, calc:true if italic}
  //         | {label, pct:{num, den}} for percentage rows
  var rows = [
    {section:'SALES & PROFIT (P&L)'},
    {label:'Sales / Revenue',           col: MCC.SALES + 1},
    {label:'Opening Stock',             col: MCC.OPENING_STOCK + 1},
    {label:'Purchases',                 col: MCC.PURCHASES + 1},
    {label:'Closing Stock',             col: MCC.CLOSING_STOCK + 1},
    {label:'COGS',                      col: MCC.COGS + 1, calc:true},
    {label:'Gross Profit',              col: MCC.GP + 1, calc:true, strong:true},
    {label:'GP %',                      pct: {num: MCC.GP+1, den: MCC.SALES+1}},
    {label:'Operating Expenses (OPEX)', col: MCC.OPEX + 1},
    {label:'EBITDA',                    col: MCC.EBITDA + 1, calc:true, strong:true},
    {label:'EBITDA %',                  pct: {num: MCC.EBITDA+1, den: MCC.SALES+1}},
    {label:'Depreciation',              col: MCC.DEPRECIATION + 1},
    {label:'Interest',                  col: MCC.INTEREST + 1},
    {label:'Tax',                       col: MCC.TAX + 1},
    {label:'PAT (Profit After Tax)',    col: MCC.PAT + 1, calc:true, strong:true},
    {label:'PAT %',                     pct: {num: MCC.PAT+1, den: MCC.SALES+1}},
    {section:'CASH FLOW'},
    {label:'Opening Cash',              col: MCC.OPENING_CASH + 1},
    {label:'Advance Received',          col: MCC.ADV_RECEIVED + 1},
    {label:'New Sales Cash',            col: MCC.NEW_SALES_CASH + 1},
    {label:'From Debtors',              col: MCC.FROM_DEBTORS + 1},
    {label:'Bank / CC OD Loan In',      col: MCC.BANK_LOAN_IN + 1},
    {label:'Capital In',                col: MCC.CAPITAL_IN + 1},
    {label:'Business Ops Out',          col: MCC.BIZ_OPS_OUT + 1},
    {label:'Loan Repay',                col: MCC.LOAN_REPAY + 1},
    {label:'Personal / CapEx',          col: MCC.PERSONAL_CAPEX + 1},
    {label:'Net Cash Flow',             col: MCC.NET_CASH_FLOW + 1, calc:true, strong:true},
    {label:'Closing Cash',              col: MCC.CLOSING_CASH + 1, calc:true, strong:true},
    {section:'BALANCE SHEET'},
    {label:'Accounts Receivable',       col: MCC.AR + 1},
    {label:'Cash & Bank Balance',       col: MCC.CASH_BANK + 1},
    {label:'Closing Inventory',         col: MCC.CLOSING_INV + 1},
    {label:'Loans Given',               col: MCC.LOANS_GIVEN + 1},
    {label:'Advance Tax / Tax Credit',  col: MCC.TAX_CREDIT + 1},
    {label:'Prepaid / Advances',        col: MCC.PREPAID + 1},
    {label:'Accounts Payable',          col: MCC.AP + 1},
    {label:'Short-term Loans Taken',    col: MCC.LOANS_TAKEN + 1},
    {label:'Security Deposit Taken',    col: MCC.SECURITY_DEPOSIT + 1},
    {label:'Tax Payable',               col: MCC.TAX_PAYABLE + 1},
    {label:'Net Working Capital',       col: MCC.NWC + 1, calc:true, strong:true},
    {label:'CC Limit',                  col: MCC.CC_LIMIT + 1},
    {label:'CC Used',                   col: MCC.CC_USED + 1}
  ];

  // ── Build rows ──
  var startRow = 4;
  var sectionRows = [], strongRows = [], calcRows = [];
  var lookupRange = _viewLookupRange();

  for (var r=0; r<rows.length; r++){
    var rowDef = rows[r];
    var actualRow = startRow + r;
    if (rowDef.section){
      sheet.getRange(actualRow, 1).setValue(rowDef.section);
      sectionRows.push(actualRow);
      continue;
    }
    sheet.getRange(actualRow, 1).setValue(rowDef.label);
    if (rowDef.strong) strongRows.push(actualRow);
    if (rowDef.calc)   calcRows.push(actualRow);

    for (var m=0; m<12; m++){
      var colNum = m + 2;
      var monthCell = '$' + _colA1(colNum) + '$3';
      var formula;
      if (rowDef.pct){
        // Percentage: numerator / denominator, with safe divide
        var numFormula = 'IFERROR(IFNA(FILTER(' + MC.SHEET + '!' + _colA1(rowDef.pct.num) + ':' + _colA1(rowDef.pct.num) +
                  ',' + MC.SHEET + '!A:A=$B$1,' + MC.SHEET + '!B:B=' + monthCell + '),0),0)';
        var denFormula = 'IFERROR(IFNA(FILTER(' + MC.SHEET + '!' + _colA1(rowDef.pct.den) + ':' + _colA1(rowDef.pct.den) +
                  ',' + MC.SHEET + '!A:A=$B$1,' + MC.SHEET + '!B:B=' + monthCell + '),0),0)';
        formula = '=IFERROR(IF(' + denFormula + '=0,"",ROUND(' + numFormula + '/' + denFormula + '*100,1)),"")';
      } else {
        // Look up value where M_MonthlyClose A=fy AND B=month
        formula = '=IFERROR(IFNA(FILTER(' + MC.SHEET + '!' + _colA1(rowDef.col) + ':' + _colA1(rowDef.col) +
                  ',' + MC.SHEET + '!A:A=$B$1,' + MC.SHEET + '!B:B=' + monthCell + '),""),"")';
      }
      sheet.getRange(actualRow, colNum).setFormula(formula);
    }
  }

  // ── Formatting ──
  var lastRow = startRow + rows.length - 1;
  // Number format for value cells
  sheet.getRange(startRow, 2, rows.length, 12).setNumberFormat('#,##0;(#,##0);"-"');
  // Section headers
  sectionRows.forEach(function(rr){
    sheet.getRange(rr, 1, 1, 13).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
    sheet.setRowHeight(rr, 26);
  });
  // Strong (highlighted total) rows
  strongRows.forEach(function(rr){
    sheet.getRange(rr, 1, 1, 13).setFontWeight('bold').setBackground('#FFF1E6');
  });
  // Italic (calculated) row labels
  calcRows.forEach(function(rr){ sheet.getRange(rr, 1).setFontStyle('italic'); });

  // Freeze first col + first 3 rows
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(3);

  // Column widths
  sheet.setColumnWidth(1, 240);
  for (var w=2; w<=13; w++) sheet.setColumnWidth(w, 110);

  // Help footer
  var helpRow = lastRow + 2;
  sheet.getRange(helpRow, 1).setValue(
    'Read-only view — change cell B1 to switch FY. Edits go in the app or directly in the M_MonthlyClose sheet.'
  );
  sheet.getRange(helpRow, 1).setFontColor('#6B7A89').setFontStyle('italic');

  return sheet;
}

// ── Helpers used by the view builder ────────────────────────
function _colA1(n){
  // 1 -> A, 2 -> B, ..., 27 -> AA, ...
  var s = '';
  while (n > 0){
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function _viewLookupRange(){
  return MC.SHEET + '!A:' + _colA1(MCC_TOTAL);
}