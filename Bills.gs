// ============================================================
// 3 Star Projects — Unified Finance System
// MODULE: Bills / Payables  (vendor-side mirror of Invoices/Payments)
// File: Bills.gs
//
// Two new sheets:
//   Bills           — vendor bills (the liability)
//   Bill_Payments   — payments you've made against those bills
//
// Reuses M_Party (Vendor type), payment modes, sources of funds.
// Auto-applies a payment to the vendor's oldest outstanding bills first (FIFO).
// Provides bill-list, single-bill, vendor-ledger, AP ageing for the dashboard.
// ============================================================

var AP = {
  BILLS:       'Bills',
  PAYMENTS:    'Bill_Payments',
  ALLOCATIONS: 'Bill_Allocations'   // Option-B project allocation per bill line
};

var BILL_STATUS = { UNPAID:'Unpaid', PARTIAL:'Partial', PAID:'Paid', CANCELLED:'Cancelled' };

// ── Column maps ─────────────────────────────────────────────
// COUNT_IN_PROJECT_COST (Batch A): when 'Yes', this bill's gross (Inc-GST)
// amount is added to its tagged project's Total Cost. Blank/'No' = excluded,
// so historical bills never double-count with Project_Costs entries.
var BC = {
  ID:0, VEND_ID:1, VEND_NAME:2, GSTIN:3, BILL_NO:4, BILL_DATE:5, DUE_DATE:6, FY:7,
  TAXABLE:8, GST_RATE:9, CGST:10, SGST:11, IGST:12, AMOUNT:13,
  OUTSTANDING:14, PAID:15, CATEGORY:16, PROOF:17, STATUS:18, NOTES:19, CREATED:20,
  PROJECT:21,
  COUNT_IN_PROJECT_COST:22,
  // Appended Oct 2026 — kept at the end so existing column indices don't shift.
  DESCRIPTION:23,   // short per-bill description ("transportation", "cement for site A"…)
  GROUP:24          // P&L group — auto-filled from Category via onEdit, same rule as Transactions
};
var BC_TOTAL = 25;
var BILL_HEADERS = ['Bill_ID','Vendor_ID','Vendor_Name','GSTIN','Bill_No','Bill_Date','Due_Date','FY',
  'Taxable','GST_Rate','CGST','SGST','IGST','Total_Amount','Outstanding','Paid','Category','Proof_Link',
  'Status','Notes','Created_Date','Project','Count_In_Project_Cost','Description','Group'];
var BILL_WIDTHS = [110,110,220,150,140,110,110,90,110,80,90,90,90,120,120,100,180,180,90,260,160,180,150,260,160];

var BPMC = {
  ID:0, VEND_ID:1, VEND_NAME:2, DATE:3, FY:4, AMOUNT:5, MODE:6, REF:7,
  SOURCE:8, NOTES:9, APPLIED:10, CREATED:11, TDS_AMT:12, BANK_AMT:13,
  // Tier 2.1 audit columns
  CANCELLED_ON:14, CANCELLED_BY:15, CANCELLATION_REASON:16,
  LAST_MODIFIED_ON:17, LAST_MODIFIED_BY:18,
  // Added: Project (for payments made without a bill — project can't inherit)
  PROJECT:19
};
var BPMC_TOTAL = 20;
var BPM_HEADERS = ['Payment_ID','Vendor_ID','Vendor_Name','Payment_Date','FY','Amount','Mode','Reference',
  'Source_of_Funds','Notes','Applied_Bills','Created_Date','TDS_Amount','Bank_Amount',
  'Cancelled_On','Cancelled_By','Cancellation_Reason','Last_Modified_On','Last_Modified_By',
  'Project'];
var BPM_WIDTHS = [110,110,220,110,90,120,140,160,180,260,300,160,110,120, 160,200,300,160,200, 180];

// ── Bill_Allocations (Option B) ─────────────────────────────────
// Carves a single bill across projects. One row per {bill, project} line.
// Sum of a bill's allocations should equal the bill's AMOUNT (gross incl GST);
// `setBillAllocations` enforces ≤ bill AMOUNT at write time. When a bill has
// one or more non-deleted allocations, project-cost rollups use them instead
// of the Bill.Project column (which stays as the primary tag for all other
// uses — filters, listings, etc.).
var BAC = {
  ID:0, BILL_ID:1, BILL_NO:2, PROJECT:3, AMOUNT:4, NOTES:5,
  CREATED_DATE:6, CREATED_BY:7, DELETED:8
};
var BAC_TOTAL = 9;
var BA_HEADERS = ['Allocation_ID','Bill_ID','Bill_No','Project','Amount','Notes',
                  'Created_Date','Created_By','Deleted'];
var BA_WIDTHS  = [120, 110, 140, 220, 120, 260, 160, 160, 90];

// ============================================================
// _writeRowVerified(sheet, row, keyColIdx0Based, keyValue)
// Reliable-write helper. appendRow + flush is asynchronous in Apps Script
// and has been observed to return without persisting under load, or persist
// so late the user's other browser tab still shows a stale sheet. This
// helper writes to an explicit range via setValues (synchronous), flushes,
// then reads back the key column of the just-written row to confirm the
// value landed. Returns {ok:true} on success, {ok:false, message:...} if
// verification fails after one retry.
// ============================================================
function _writeRowVerified(sheet, row, keyColIdx0Based, keyValue){
  function attempt(){
    var targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    var seen = _s(sheet.getRange(targetRow, keyColIdx0Based + 1).getValue());
    return seen === _s(keyValue);
  }
  if (attempt()) return {ok:true};
  // One retry — clears any transient GAS write hiccup.
  Utilities.sleep(300);
  if (attempt()) return {ok:true};
  return {ok:false, message:'Sheet write did not persist for key ' + keyValue + '. Please refresh the app (Ctrl+Shift+R) and try once more; if it recurs, the sheet may be protected or GAS is rate-limiting.'};
}

// ============================================================
// Setup — idempotent. Safe to run repeatedly via setupMasters.
// ============================================================
function setupBillsModule(){
  var ss = _getSS();
  var bills = _buildSheet(ss, AP.BILLS, BILL_HEADERS, BILL_WIDTHS);
  _validateColSoft(bills, BC.STATUS+1, ['Unpaid','Partial','Paid','Cancelled']);
  _validateColSoft(bills, BC.GST_RATE+1, ['0','5','12','18','28']);
  // Schema upgrade: if Bills existed before new columns were added,
  // append the missing headers. Historical rows stay blank.
  if (bills.getLastColumn() < BC_TOTAL){
    var bStart = bills.getLastColumn() + 1;
    var bNewHeaders = BILL_HEADERS.slice(bStart - 1);
    bills.getRange(1, bStart, 1, bNewHeaders.length).setValues([bNewHeaders]).setFontWeight('bold');
  }
  // Yes/No soft-validation on Count_In_Project_Cost (Batch A).
  // Done AFTER the schema upgrade so the column always exists.
  _validateColSoft(bills, BC.COUNT_IN_PROJECT_COST+1, ['Yes','No']);
  // Category — live dropdown from M_Category!A2:A1000, same rule as Transactions.
  // Group — no dropdown; auto-filled by onEdit when Category changes.
  // _validateColRange / _clearColValidation live in Transactions.gs (shared globals).
  try {
    if (typeof _validateColRange === 'function')
      _validateColRange(bills, BC.CATEGORY+1, "'M_Category'!A2:A1000");
    if (typeof _clearColValidation === 'function')
      _clearColValidation(bills, BC.GROUP+1);
  } catch(e){ Logger.log('setupBillsModule category/group validators: '+e.message); }
  var pays  = _buildSheet(ss, AP.PAYMENTS, BPM_HEADERS, BPM_WIDTHS);
  // Schema upgrade: if Bill_Payments existed before TDS columns were added,
  // append the new headers and back-fill (TDS=0, Bank=Amount) so historical
  // rows stay consistent.
  if (pays.getLastColumn() < BPMC_TOTAL){
    var startCol = pays.getLastColumn() + 1;
    var newHeaders = BPM_HEADERS.slice(startCol - 1);
    pays.getRange(1, startCol, 1, newHeaders.length).setValues([newHeaders]).setFontWeight('bold');
    var lr = pays.getLastRow();
    if (lr >= 2){
      var amounts = pays.getRange(2, BPMC.AMOUNT+1, lr-1, 1).getValues();
      var backfill = amounts.map(function(r){ return [0, _n(r[0])]; });
      pays.getRange(2, BPMC.TDS_AMT+1, lr-1, 2).setValues(backfill);
    }
  }
  // Bill_Allocations — idempotent create. Dropdown on Project column is applied
  // live by applyProjectValidationsEverywhere (Masters) so new projects reach
  // this sheet automatically.
  var alloc = _buildSheet(ss, AP.ALLOCATIONS, BA_HEADERS, BA_WIDTHS);
  _validateColSoft(alloc, BAC.DELETED+1, ['Yes','No']);
  return { bills: bills.getName(), payments: pays.getName(), allocations: alloc.getName() };
}

// ============================================================
// Next IDs
// ============================================================
function _nextBillId(){
  var sheet = _getSS().getSheetByName(AP.BILLS);
  if (!sheet || !_hasData(sheet)) return 'BILL-0001';
  var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  var max = 0;
  for (var i=0;i<ids.length;i++){
    var v = _s(ids[i][0]);
    if (v.indexOf('BILL-') === 0){
      var n = parseInt(v.substring(5),10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return 'BILL-' + String(max+1).padStart(4,'0');
}
function _nextBillPaymentId(){
  var sheet = _getSS().getSheetByName(AP.PAYMENTS);
  if (!sheet || !_hasData(sheet)) return 'BPAY-0001';
  var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  var max = 0;
  for (var i=0;i<ids.length;i++){
    var v = _s(ids[i][0]);
    if (v.indexOf('BPAY-') === 0){
      var n = parseInt(v.substring(5),10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return 'BPAY-' + String(max+1).padStart(4,'0');
}

// ============================================================
// CREATE BILL
//   data = {vendorId, billNo, billDate, dueDate, taxable, gstRate, cgst, sgst, igst,
//           amount, category, gst (bool), proofLink, notes}
// ============================================================
function createBill(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    if (!_s(data.vendorId)) return {success:false, message:'Vendor is required.'};
    if (!_s(data.billNo))   return {success:false, message:'Bill number is required.'};
    var amount = _n(data.amount);
    if (amount <= 0) return {success:false, message:'Amount must be greater than zero.'};

    // Look up vendor for name + GSTIN snapshot
    var vendor = _findPartyById(data.vendorId);
    if (!vendor) return {success:false, message:'Vendor not found.'};

    var ss = _getSS(), sheet = ss.getSheetByName(AP.BILLS);
    if (!sheet){ setupBillsModule(); sheet = ss.getSheetByName(AP.BILLS); }

    // Duplicate guard: same vendor + bill_no
    if (_hasData(sheet)){
      var existing = sheet.getRange(2,1,sheet.getLastRow()-1,BC_TOTAL).getValues();
      for (var i=0;i<existing.length;i++){
        if (_s(existing[i][BC.VEND_ID]) === _s(data.vendorId)
            && _s(existing[i][BC.BILL_NO]).toLowerCase() === _s(data.billNo).toLowerCase()
            && _s(existing[i][BC.STATUS]) !== BILL_STATUS.CANCELLED){
          return {success:false, message:'A bill with this number already exists for this vendor.'};
        }
      }
    }

    var billDate = data.billDate ? _asDate(data.billDate) : new Date();
    var dueDate  = data.dueDate  ? _asDate(data.dueDate)  : billDate;
    var id = _nextBillId();
    var now = new Date();

    var row = new Array(BC_TOTAL).fill('');
    row[BC.ID]          = id;
    row[BC.VEND_ID]     = vendor.id;
    row[BC.VEND_NAME]   = vendor.name;
    row[BC.GSTIN]       = vendor.gstin || '';
    row[BC.BILL_NO]     = _s(data.billNo);
    row[BC.BILL_DATE]   = billDate;
    row[BC.DUE_DATE]    = dueDate;
    row[BC.FY]          = _fyFromDate(billDate);
    row[BC.TAXABLE]     = _n(data.taxable) || amount;
    row[BC.GST_RATE]    = data.gst ? _n(data.gstRate) : 0;
    row[BC.CGST]        = data.gst ? _n(data.cgst) : 0;
    row[BC.SGST]        = data.gst ? _n(data.sgst) : 0;
    row[BC.IGST]        = data.gst ? _n(data.igst) : 0;
    row[BC.AMOUNT]      = amount;
    row[BC.OUTSTANDING] = amount;
    row[BC.PAID]        = 0;
    row[BC.CATEGORY]    = _s(data.category);
    row[BC.PROOF]       = _s(data.proofLink);
    row[BC.STATUS]      = BILL_STATUS.UNPAID;
    row[BC.NOTES]       = _s(data.notes);
    row[BC.CREATED]     = now;
    row[BC.PROJECT]     = _s(data.project);
    row[BC.DESCRIPTION] = _s(data.description);
    // Stamp Group from Category at write time (same rule onEdit uses). Falls
    // back to blank if the category isn't in M_Category; the live-resolve
    // dashboard pass fills the gap at render time anyway.
    try { row[BC.GROUP] = (typeof _groupForCategory === 'function') ? _s(_groupForCategory(data.category)) : ''; }
    catch(e){ row[BC.GROUP] = ''; }
    // Reliable write with verification — appendRow+flush was still leaving
    // rows invisible for some users. _writeRowVerified uses explicit setValues
    // and reads back the ID to confirm; retries once on mismatch.
    var writeRes = _writeRowVerified(sheet, row, BC.ID, id);
    if (!writeRes.ok) return {success:false, message:writeRes.message};
    var billRowIdx = sheet.getLastRow();

    // Order matters:
    //   1) paidNow first — it's a NEW payment the operator is entering right now,
    //      needs the bill fully open (targetBill=billNo). Running auto-consume
    //      before this could close the bill and cause paidNow's recordBillPayment
    //      to fail with "Bill not found among open bills".
    //   2) Then auto-consume any leftover with existing vendor advances.
    var paidNow = _n(data.paidNow);
    var paidNowInfo = null;
    if (paidNow > 0){
      if (paidNow > amount){
        // Refuse over-payment on creation — an easy typo trap.
        return {success:false, message:'"Amount already paid" (₹' + paidNow + ') cannot exceed the bill total (₹' + amount + ').'};
      }
      paidNowInfo = recordBillPayment({
        vendorId:    vendor.id,
        amount:      paidNow,
        paymentDate: data.paidNowDate || data.billDate,
        paymentMode: _s(data.paidNowMode),
        reference:   _s(data.paidNowRef),
        source:      _s(data.paidNowSource),
        tdsAmount:   _n(data.paidNowTds),
        notes:       'First payment against ' + _s(data.billNo),
        targetBill:  _s(data.billNo)
      });
      if (paidNowInfo && !paidNowInfo.success){
        // Bill IS created — just report the payment failure so the owner can retry it manually.
        return {success:true, billId:id, paymentError: paidNowInfo.message,
                message:'Bill ' + id + ' recorded, but the first payment failed: ' + paidNowInfo.message};
      }
    }

    // Auto-consume any open advances against whatever outstanding remains.
    // Closes the gap where an advance would sit forever until the operator
    // manually re-linked it in the Vendor Ledger bulk-fix panel.
    var autoInfo = _autoConsumeAdvancesForNewBill(id, vendor.id, billRowIdx);

    var msg = 'Bill ' + id + ' recorded.';
    if (paidNowInfo && paidNowInfo.success){
      msg += ' Paid ₹' + paidNow + ' now (' + paidNowInfo.paymentId + ').';
    }
    if (autoInfo && autoInfo.consumed > 0){
      msg += ' Auto-applied ₹' + autoInfo.consumed + ' from ' + autoInfo.paymentsCount +
             ' vendor advance' + (autoInfo.paymentsCount === 1 ? '' : 's') + '.';
    }
    return {success:true, billId:id, firstPaymentId:(paidNowInfo && paidNowInfo.paymentId) || '', message:msg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// RECORD BILL PAYMENT (FIFO auto-apply to oldest unpaid bills)
//   data = {vendorId, paymentDate, amount, paymentMode, reference, source, notes}
// ============================================================
// Returns open bills for a vendor, oldest first. Used by the AP payment
// form's optional "Apply to specific bill" dropdown.
//   [{no, date, total, outstanding}]
function getVendorOpenBills(vendorId){
  try {
    var vid = _s(vendorId);
    if (!vid) return [];
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    if (!billsSheet || billsSheet.getLastRow() < 2) return [];
    var rows = billsSheet.getRange(2,1,billsSheet.getLastRow()-1,BC_TOTAL).getValues();
    var out = [];
    for (var i=0;i<rows.length;i++){
      var r = rows[i];
      if (_s(r[BC.VEND_ID]) !== vid) continue;
      var status = _s(r[BC.STATUS]);
      if (status === BILL_STATUS.PAID || status === BILL_STATUS.CANCELLED) continue;
      var outstanding = _n(r[BC.OUTSTANDING]);
      if (outstanding <= 0) continue;
      out.push({
        no: _s(r[BC.BILL_NO]),
        date: Utilities.formatDate(_asDate(r[BC.BILL_DATE]),'GMT+5:30','yyyy-MM-dd'),
        total: _n(r[BC.TOTAL]),
        outstanding: outstanding,
        status: status,
        project: _s(r[BC.PROJECT])
      });
    }
    out.sort(function(a,b){ return a.date.localeCompare(b.date); });
    return out;
  } catch(e){ Logger.log('getVendorOpenBills: '+e.message); return []; }
}

function recordBillPayment(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    if (!_s(data.vendorId)) return {success:false, message:'Vendor is required.'};
    var amount = _n(data.amount);
    if (amount <= 0) return {success:false, message:'Amount must be greater than zero.'};

    var vendor = _findPartyById(data.vendorId);
    if (!vendor) return {success:false, message:'Vendor not found.'};

    var ss = _getSS();
    var billsSheet = ss.getSheetByName(AP.BILLS);
    var payments  = ss.getSheetByName(AP.PAYMENTS);
    if (!payments){ setupBillsModule(); payments = ss.getSheetByName(AP.PAYMENTS); }

    // Three ways to apply a payment, in priority order:
    //   1. data.billAllocations = [{billNo, amount}] — operator-chosen split
    //      across one or more open bills of THIS vendor. Exact amounts.
    //   2. data.targetBill = "BILL-xxxx" — single bill first, FIFO for leftover.
    //   3. default — FIFO across all open bills (oldest first).
    var applied = [];
    var remaining = amount;
    if (billsSheet && _hasData(billsSheet)){
      var n = billsSheet.getLastRow()-1;
      var all = billsSheet.getRange(2,1,n,BC_TOTAL).getValues();
      var rows = [];
      for (var i=0;i<all.length;i++){
        if (_s(all[i][BC.VEND_ID]) !== _s(data.vendorId)) continue;
        if (_s(all[i][BC.STATUS]) === BILL_STATUS.CANCELLED) continue;
        if (_n(all[i][BC.OUTSTANDING]) <= 0) continue;
        rows.push({rowIdx:i+2, data:all[i]});
      }
      // Sort by bill_date ascending (oldest first) — base order for both FIFO
      // and single-target.
      rows.sort(function(a,b){
        var da = _asDate(a.data[BC.BILL_DATE]).getTime();
        var db = _asDate(b.data[BC.BILL_DATE]).getTime();
        return da - db;
      });

      // Build a quick lookup for the explicit-split path.
      var byBillNo = {}, byBillId = {};
      rows.forEach(function(r){
        byBillNo[_s(r.data[BC.BILL_NO])] = r;
        byBillId[_s(r.data[BC.ID])]      = r;
      });

      // ── Mode 1: explicit multi-bill split ──
      if (Array.isArray(data.billAllocations) && data.billAllocations.length){
        var sumAlloc = 0;
        data.billAllocations.forEach(function(a){ sumAlloc += _n(a && a.amount) || 0; });
        if (sumAlloc > amount + 0.01){
          return {success:false, message:'Split total (₹'+_round(sumAlloc)+') exceeds payment amount (₹'+_round(amount)+').'};
        }
        for (var ai=0; ai<data.billAllocations.length; ai++){
          var a = data.billAllocations[ai] || {};
          var ref = _s(a.billNo || a.billId);
          var ask = _n(a.amount) || 0;
          if (!ref || ask <= 0) continue;
          var row = byBillNo[ref] || byBillId[ref];
          if (!row){
            return {success:false, message:'Bill ' + ref + ' is not an open bill for this vendor (closed, cancelled, or belongs to someone else).'};
          }
          var out = _n(row.data[BC.OUTSTANDING]);
          if (ask > out + 0.01){
            return {success:false, message:'Allocation ₹'+_round(ask)+' exceeds outstanding ₹'+_round(out)+' on bill '+_s(row.data[BC.BILL_NO])+'.'};
          }
          var newOut = _round(out - ask);
          var newPaid = _round(_n(row.data[BC.PAID]) + ask);
          var newStatus = newOut <= 0.01 ? BILL_STATUS.PAID : BILL_STATUS.PARTIAL;
          billsSheet.getRange(row.rowIdx, BC.OUTSTANDING+1).setValue(newOut);
          billsSheet.getRange(row.rowIdx, BC.PAID+1).setValue(newPaid);
          billsSheet.getRange(row.rowIdx, BC.STATUS+1).setValue(newStatus);
          applied.push(_s(row.data[BC.ID]) + ':' + _round(ask));
          remaining = _round(remaining - ask);
          // Mutate the in-memory row so a later allocation to the SAME bill in
          // the same payload sees the fresh outstanding (operator unlikely but
          // keeps us honest).
          row.data[BC.OUTSTANDING] = newOut;
          row.data[BC.PAID]        = newPaid;
          row.data[BC.STATUS]      = newStatus;
        }
      } else {
        // ── Mode 2 (single target) / Mode 3 (FIFO) ──
        var targetBill = _s(data.targetBill);
        if (targetBill){
          var targetIdx = -1;
          for (var t=0; t<rows.length; t++){ if (_s(rows[t].data[BC.BILL_NO]) === targetBill){ targetIdx = t; break; } }
          if (targetIdx > 0){
            var picked = rows.splice(targetIdx, 1)[0];
            rows.unshift(picked);
          } else if (targetIdx === -1){
            return {success:false, message:'Bill ' + targetBill + ' not found among this vendor\'s open bills.'};
          }
        }
        for (var k=0; k<rows.length && remaining > 0; k++){
          var r = rows[k], outstanding = _n(r.data[BC.OUTSTANDING]);
          var apply = Math.min(outstanding, remaining);
          var newOut2 = _round(outstanding - apply);
          var newPaid2 = _round(_n(r.data[BC.PAID]) + apply);
          var newStatus2 = newOut2 <= 0.01 ? BILL_STATUS.PAID : BILL_STATUS.PARTIAL;
          billsSheet.getRange(r.rowIdx, BC.OUTSTANDING+1).setValue(newOut2);
          billsSheet.getRange(r.rowIdx, BC.PAID+1).setValue(newPaid2);
          billsSheet.getRange(r.rowIdx, BC.STATUS+1).setValue(newStatus2);
          applied.push(_s(r.data[BC.ID]) + ':' + _round(apply));
          remaining = _round(remaining - apply);
        }
      }
    }

    // TDS handling: `amount` is the GROSS amount applied against bills
    // (what the vendor's account is credited by). `tdsAmount` is the
    // deduction you made and will deposit to government on the vendor's
    // behalf. `bankAmount` is what actually left your bank account.
    var tdsAmount = _n(data.tdsAmount);
    if (tdsAmount < 0) tdsAmount = 0;
    if (tdsAmount > amount) return {success:false, message:'TDS amount cannot exceed payment amount.'};
    var bankAmount = _round(amount - tdsAmount);

    // Write payment row (any unapplied amount is recorded as on-account/advance)
    var payDate = data.paymentDate ? _asDate(data.paymentDate) : new Date();
    var id = _nextBillPaymentId();
    var prow = new Array(BPMC_TOTAL).fill('');
    prow[BPMC.ID]        = id;
    prow[BPMC.VEND_ID]   = vendor.id;
    prow[BPMC.VEND_NAME] = vendor.name;
    prow[BPMC.DATE]      = payDate;
    prow[BPMC.FY]        = _fyFromDate(payDate);
    prow[BPMC.AMOUNT]    = amount;
    prow[BPMC.MODE]      = _s(data.paymentMode);
    prow[BPMC.REF]       = _s(data.reference);
    prow[BPMC.SOURCE]    = _s(data.source);
    prow[BPMC.NOTES]     = _s(data.notes);
    prow[BPMC.APPLIED]   = applied.join(' · ');
    prow[BPMC.CREATED]   = new Date();
    prow[BPMC.TDS_AMT]   = tdsAmount;
    prow[BPMC.BANK_AMT]  = bankAmount;
    // Project on payment row: derived ONLY from the single applied bill.
    // Multi-bill splits and advances leave it blank — per-project attribution
    // for those comes from the Applied_Bills string, not this field.
    // The form no longer exposes a Project dropdown; ignoring data.project
    // closes the double-count loophole where operator's tag differed from
    // the applied bill's project.
    var project = '';
    if (applied.length === 1){
      var onlyBillId = applied[0].split(':')[0];
      project = _getBillProject(onlyBillId);
    }
    prow[BPMC.PROJECT]   = project;
    var payWriteRes = _writeRowVerified(payments, prow, BPMC.ID, id);
    if (!payWriteRes.ok) return {success:false, message:payWriteRes.message};

    var msg = 'Payment ' + id + ' recorded';
    if (applied.length) msg += ' — applied to ' + applied.length + ' bill(s)';
    if (remaining > 0)  msg += ' · ' + _round(remaining) + ' on account (no open bill)';
    if (tdsAmount > 0)  msg += ' · TDS ₹' + tdsAmount + ' deducted, bank debited ₹' + bankAmount;
    return {success:true, paymentId:id, applied:applied, advance:_round(remaining),
            tdsAmount:tdsAmount, bankAmount:bankAmount, message:msg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// LIST BILLS (with filters)
//   filters = {fy, vendorId, status, search}
//   Returns [{id, vendorId, vendor, billNo, billDate, dueDate, amount,
//             outstanding, paid, status, category, notes, daysOverdue}]
// ============================================================
function listBills(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(AP.BILLS);
    if (!sheet || !_hasData(sheet)) return [];
    var n = sheet.getLastRow()-1;
    // Read only the columns the sheet actually has. Older Bills sheets are
    // narrower than BC_TOTAL (missing the Project column) — asking for
    // BC_TOTAL there throws "Range beyond bounds" and the frontend sees "0 bills".
    var w = Math.min(sheet.getLastColumn(), BC_TOTAL);
    var all = sheet.getRange(2,1,n,w).getValues();
    var today = new Date();
    // Translate 'current' → current FY the same way salary code does.
    var fy = _s(filters.fy);
    if (fy === 'current') fy = _currentFY();
    if (fy === 'all')     fy = '';
    var vid = _s(filters.vendorId);
    var st = _s(filters.status);
    var q = _s(filters.search).toLowerCase();
    var rows = [];
    for (var i=0;i<all.length;i++){
      var r = all[i];
      if (!_s(r[BC.ID])) continue;
      if (fy && _s(r[BC.FY]) !== fy) continue;
      if (vid && _s(r[BC.VEND_ID]) !== vid) continue;
      if (st && _s(r[BC.STATUS]) !== st) continue;
      if (q){
        var descHay = w > BC.DESCRIPTION ? _s(r[BC.DESCRIPTION]) : '';
        var hay = (_s(r[BC.VEND_NAME])+' '+_s(r[BC.BILL_NO])+' '+_s(r[BC.NOTES])+' '+descHay+' '+_s(r[BC.ID])).toLowerCase();
        if (hay.indexOf(q) === -1) continue;
      }
      var dueDate = _asDate(r[BC.DUE_DATE]);
      var outstanding = _n(r[BC.OUTSTANDING]);
      var daysOverdue = 0;
      if (outstanding > 0 && dueDate < today){
        daysOverdue = Math.floor((today.getTime() - dueDate.getTime()) / 86400000);
      }
      rows.push({
        id:           _s(r[BC.ID]),
        vendorId:     _s(r[BC.VEND_ID]),
        vendor:       _s(r[BC.VEND_NAME]),
        billNo:       _s(r[BC.BILL_NO]),
        // Serialise dates to yyyy-MM-dd — raw Date objects sometimes fail
        // google.script.run's JSON pass and return null for the entire response.
        billDate:     _safeDateStr(r[BC.BILL_DATE]),
        dueDate:      _safeDateStr(r[BC.DUE_DATE]),
        fy:           _s(r[BC.FY]),
        amount:       _n(r[BC.AMOUNT]),
        outstanding:  outstanding,
        paid:         _n(r[BC.PAID]),
        category:     _s(r[BC.CATEGORY]),
        status:       _s(r[BC.STATUS]),
        notes:        _s(r[BC.NOTES]),
        project:      w > BC.PROJECT ? _s(r[BC.PROJECT]) : '',
        description:  w > BC.DESCRIPTION ? _s(r[BC.DESCRIPTION]) : '',
        group:        w > BC.GROUP ? _s(r[BC.GROUP]) : '',
        daysOverdue:  daysOverdue
      });
    }
    return rows.sort(function(a,b){
      return (b.billDate || '').localeCompare(a.billDate || '');
    });
  } catch(e){ Logger.log('listBills: '+e.message); return []; }
}

// ============================================================
// VENDOR LEDGER  (mirror of getCustomerLedger)
//   Returns {entries:[{date, type, ref, debit, credit, balance}], balance}
//   Convention: bill = credit (we owe), payment = debit (we paid down)
//   Positive balance = we still owe vendor.
// ============================================================
// Auto-consume open vendor advances (unlinked payments with amount > applied)
// against a freshly-booked bill. FIFO oldest advance first. Updates each
// payment's Applied_Bills string and the bill's Paid/Outstanding/Status.
// Also re-derives payment.PROJECT: if exactly one bill ends up applied, tag
// with that bill's project; if more than one, clear (multi-bill = ambiguous).
// Returns {consumed, paymentsCount}.
function _autoConsumeAdvancesForNewBill(billId, vendorId, billRowIdx){
  try {
    var ss = _getSS();
    var billsSheet = ss.getSheetByName(AP.BILLS);
    var payments   = ss.getSheetByName(AP.PAYMENTS);
    if (!payments || !_hasData(payments) || !billsSheet) return {consumed:0, paymentsCount:0};

    var pW = Math.min(payments.getLastColumn(), BPMC_TOTAL);
    var pRows = payments.getRange(2, 1, payments.getLastRow()-1, pW).getValues();

    // Find open advances for this vendor
    var advances = [];
    for (var i=0; i<pRows.length; i++){
      var r = pRows[i];
      if (_s(r[BPMC.VEND_ID]) !== _s(vendorId)) continue;
      if (pW > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) continue;
      var applied = _s(r[BPMC.APPLIED]);
      // Sum already-applied amounts by parsing the string
      var sumApplied = 0;
      if (applied && applied !== '(advance)'){
        applied.split('·').forEach(function(chunk){
          var m = chunk.trim().match(/:\s*([-\d.]+)\s*$/);
          if (m) sumApplied += parseFloat(m[1]) || 0;
        });
      }
      var unapplied = _n(r[BPMC.AMOUNT]) - sumApplied;
      if (unapplied > 0.01){
        advances.push({
          rowIdx:   i + 2,
          applied:  applied,
          unapplied: unapplied,
          date:     _asDate(r[BPMC.DATE]).getTime()
        });
      }
    }
    if (!advances.length) return {consumed:0, paymentsCount:0};
    advances.sort(function(a,b){ return a.date - b.date; });  // FIFO oldest first

    // Read new bill's current outstanding/paid
    var billOutstanding = _n(billsSheet.getRange(billRowIdx, BC.OUTSTANDING+1).getValue());
    var billPaid        = _n(billsSheet.getRange(billRowIdx, BC.PAID+1).getValue());
    var newBillProject  = billsSheet.getLastColumn() > BC.PROJECT
      ? _s(billsSheet.getRange(billRowIdx, BC.PROJECT+1).getValue()) : '';

    var consumed = 0, count = 0;
    for (var k=0; k<advances.length && billOutstanding > 0.01; k++){
      var adv = advances[k];
      var apply = Math.min(adv.unapplied, billOutstanding);
      if (apply <= 0.01) continue;
      var newAppliedStr = adv.applied && adv.applied !== '(advance)'
        ? adv.applied + ' · ' + billId + ':' + _round(apply)
        : billId + ':' + _round(apply);
      payments.getRange(adv.rowIdx, BPMC.APPLIED+1).setValue(newAppliedStr);
      // Backfill project on the payment row: single-bill applied → derive; multi → blank
      if (pW > BPMC.PROJECT){
        var isSingleBill = newAppliedStr.indexOf(' · ') === -1;
        payments.getRange(adv.rowIdx, BPMC.PROJECT+1).setValue(isSingleBill ? newBillProject : '');
      }
      billPaid += apply;
      billOutstanding -= apply;
      consumed += apply;
      count++;
    }

    if (consumed > 0){
      billsSheet.getRange(billRowIdx, BC.PAID+1).setValue(_round(billPaid));
      billsSheet.getRange(billRowIdx, BC.OUTSTANDING+1).setValue(_round(billOutstanding));
      billsSheet.getRange(billRowIdx, BC.STATUS+1).setValue(
        billOutstanding <= 0.01 ? BILL_STATUS.PAID : BILL_STATUS.PARTIAL
      );
      SpreadsheetApp.flush();
    }
    return {consumed:_round(consumed), paymentsCount:count};
  } catch(e){
    Logger.log('_autoConsumeAdvancesForNewBill: ' + e.message);
    return {consumed:0, paymentsCount:0};
  }
}

// Pre-check for a duplicate bill on the New Bill form. Called live as the
// operator types vendor + bill number. Returns the first matching non-
// cancelled bill so the UI can warn (not block). Save still validates.
function checkDuplicateBill(vendorId, billNo){
  try {
    var vid = _s(vendorId), bn = _s(billNo);
    if (!vid || !bn) return {exists:false};
    var sheet = _getSS().getSheetByName(AP.BILLS);
    if (!sheet || !_hasData(sheet)) return {exists:false};
    var w = Math.min(sheet.getLastColumn(), BC_TOTAL);
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
    var bnLc = bn.toLowerCase();
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][BC.VEND_ID]) !== vid) continue;
      if (_s(rows[i][BC.BILL_NO]).toLowerCase() !== bnLc) continue;
      if (_s(rows[i][BC.STATUS]) === BILL_STATUS.CANCELLED) continue;
      return {
        exists: true,
        billId: _s(rows[i][BC.ID]),
        billDate: _safeDateStr(rows[i][BC.BILL_DATE]),
        amount: _n(rows[i][BC.AMOUNT]),
        status: _s(rows[i][BC.STATUS]),
        project: w > BC.PROJECT ? _s(rows[i][BC.PROJECT]) : ''
      };
    }
    return {exists:false};
  } catch(e){ return {exists:false, error:String(e && e.message || e)}; }
}

// Enrich an Applied_Bills string with each bill's vendor bill/PI number and
// cancellation status. Lookup may be either the simple {id → vendorRef} map
// (Payments Made list) or the enriched {id → {ref, status}} map (Ledger).
// Input : "BILL-0005:1000 · BILL-0002:640"
// Output: "BILL-0005 (INV-234):1000 · BILL-0002 (PI-987) [cancelled]:640"
// Refs not in the lookup pass through unchanged (bill hard-deleted or belongs
// to a different vendor).
function _enrichAppliedWithVendorBillNo(applied, lookup){
  applied = _s(applied);
  if (!applied || applied === '(advance)') return applied;
  return applied.split('·').map(function(chunk){
    var t = chunk.trim(); if (!t) return t;
    var m = t.match(/^(.+?):\s*([-\d.]+)\s*$/);
    if (!m) return t;
    var id = m[1].trim(), amt = m[2];
    var meta = lookup[id];
    if (!meta) return id + ':' + amt;
    // Support both string ("vRef") and object ({ref, status}) shapes.
    var vRef = (typeof meta === 'string') ? meta : meta.ref;
    var status = (typeof meta === 'object') ? meta.status : '';
    var base = vRef && vRef !== id ? (id + ' (' + vRef + ')') : id;
    if (status === BILL_STATUS.CANCELLED) base += ' [cancelled]';
    return base + ':' + amt;
  }).join(' · ');
}

function getVendorLedger(vendorId){
  try {
    if (!_s(vendorId)) return {ok:true, entries:[], balance:0, error:null};
    var entries = [];
    // Lookup: BILL_ID → {ref: vendor's bill/PI number, status: bill status}.
    // Used to enrich the payment's Applied_Bills string in the ledger so
    // operators see the vendor's ref AND whether the linked bill has since
    // been cancelled.
    var billLookup = {};
    var bs = _getSS().getSheetByName(AP.BILLS);
    if (bs && _hasData(bs)){
      var b = bs.getRange(2,1,bs.getLastRow()-1,BC_TOTAL).getValues();
      for (var i=0;i<b.length;i++){
        if (_s(b[i][BC.VEND_ID]) !== _s(vendorId)) continue;
        billLookup[_s(b[i][BC.ID])] = {
          ref:    _s(b[i][BC.BILL_NO]),
          status: _s(b[i][BC.STATUS])
        };
        if (_s(b[i][BC.STATUS]) === BILL_STATUS.CANCELLED) continue;
        entries.push({
          date: _safeDateStr(b[i][BC.BILL_DATE]),
          type: 'Bill',
          ref: _s(b[i][BC.ID]),                // our internal (BILL-0005)
          vendorRef: _s(b[i][BC.BILL_NO]),     // vendor's bill/PI number
          credit: _n(b[i][BC.AMOUNT]),
          debit: 0,
          project: _s(b[i][BC.PROJECT]),
          applied: '',
          unapplied: 0,
          sortKey: _asDate(b[i][BC.BILL_DATE]).getTime()
        });
      }
    }
    var ps = _getSS().getSheetByName(AP.PAYMENTS);
    if (ps && _hasData(ps)){
      var pW = Math.min(ps.getLastColumn(), BPMC_TOTAL);
      var p = ps.getRange(2,1,ps.getLastRow()-1,pW).getValues();
      for (var j=0;j<p.length;j++){
        if (_s(p[j][BPMC.VEND_ID]) !== _s(vendorId)) continue;
        // Skip soft-cancelled payments if the column exists (Tier 2.1 schema).
        if (typeof BPMC.CANCELLED_ON !== 'undefined' && pW > BPMC.CANCELLED_ON
            && _s(p[j][BPMC.CANCELLED_ON])) continue;
        var rawApplied = _s(p[j][BPMC.APPLIED]);
        // Compute unapplied: payment amount minus sum of everything already
        // applied. Frontend uses this to label "kept as advance" or "pure
        // advance" states distinctly.
        var sumApplied = 0;
        if (rawApplied && rawApplied !== '(advance)'){
          rawApplied.split('·').forEach(function(chunk){
            var mm = chunk.trim().match(/:\s*([-\d.]+)\s*$/);
            if (mm) sumApplied += parseFloat(mm[1]) || 0;
          });
        }
        var payAmt = _n(p[j][BPMC.AMOUNT]);
        entries.push({
          date: _safeDateStr(p[j][BPMC.DATE]),
          type: 'Payment',
          ref: _s(p[j][BPMC.ID]),
          credit: 0,
          debit: payAmt,
          project: pW > BPMC.PROJECT ? _s(p[j][BPMC.PROJECT]) : '',
          applied: _enrichAppliedWithVendorBillNo(rawApplied, billLookup),
          unapplied: _round(payAmt - sumApplied),
          sortKey: _asDate(p[j][BPMC.DATE]).getTime()
        });
      }
    }
    // Credit notes (vendor returns) — reduce vendor balance.
    // Lives in BillReturns.gs; defensive guard in case the file is missing.
    try {
      if (typeof getBillReturnLedgerEntries === 'function'){
        var __br = getBillReturnLedgerEntries(vendorId);
        for (var __k=0; __k<__br.length; __k++) entries.push(__br[__k]);
      }
    } catch(__e){ Logger.log('getVendorLedger returns inject: '+__e.message); }
    entries.sort(function(a,b){ return a.sortKey - b.sortKey; });
    var bal = 0;
    entries.forEach(function(e){ bal += e.credit - e.debit; e.balance = _round(bal); delete e.sortKey; });
    return {ok:true, entries:entries, balance:_round(bal), error:null};
  } catch(e){
    Logger.log('getVendorLedger: '+e.message);
    return {ok:false, entries:[], balance:0, error:String(e && e.message || e)};
  }
}

// ============================================================
// AP AGEING  (mirror of dashArAgeing for the dashboard)
//   Returns {ok, data:{buckets:{current, d1_30, d31_60, d61_90, d90plus}}}
// ============================================================
function dashApAgeing(filters){
  try {
    filters = filters || {};
    // FY filter: 'current' → live FY. 'all' or blank → every FY (legacy view).
    var wantFy = _s(filters.fy);
    if (wantFy === 'current') wantFy = _currentFY();
    if (wantFy === 'all')     wantFy = '';
    var sheet = _getSS().getSheetByName(AP.BILLS);
    var buckets = {current:0, b30:0, b60:0, b90:0, b90plus:0};
    if (sheet && _hasData(sheet)){
      var today = new Date();
      var w = Math.min(sheet.getLastColumn(), BC_TOTAL);
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
      rows.forEach(function(r){
        var out = _n(r[BC.OUTSTANDING]);
        if (out <= 0) return;
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (wantFy && _s(r[BC.FY]) !== wantFy) return;
        var due = _asDate(r[BC.DUE_DATE]);
        var days = Math.floor((today.getTime() - due.getTime()) / 86400000);
        if (days <= 0)        buckets.current  += out;
        else if (days <= 30)  buckets.b30      += out;
        else if (days <= 60)  buckets.b60      += out;
        else if (days <= 90)  buckets.b90      += out;
        else                  buckets.b90plus  += out;
      });
    }
    return _envOk({
      buckets:{ current:_round(buckets.current), d1_30:_round(buckets.b30),
                d31_60:_round(buckets.b60), d61_90:_round(buckets.b90), d90plus:_round(buckets.b90plus) },
      totalOutstanding: _round(buckets.current+buckets.b30+buckets.b60+buckets.b90+buckets.b90plus)
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// PENDING BILLS — dashboard quick-action list
//   Returns bills where Status is Unpaid or Partial.
//   Returns { ok, data: { rows: [...], total } }
// ============================================================
function listPendingBills(filters){
  try {
    filters = filters || {};
    var limit = filters.limit || 50;
    var sheet = _getSS().getSheetByName(AP.BILLS);
    if (!sheet || !_hasData(sheet)) return _envOk({ rows:[], total:0, count:0 });
    var all = sheet.getRange(2,1,sheet.getLastRow()-1,BC_TOTAL).getValues();
    var today = new Date();
    var rows = [];
    var total = 0;
    all.forEach(function(r){
      var status = _s(r[BC.STATUS]);
      if (status !== 'Unpaid' && status !== 'Partial') return;
      var outstanding = _n(r[BC.OUTSTANDING]);
      if (outstanding <= 0) return;
      var dueDate = _asDate(r[BC.DUE_DATE]);
      var daysOverdue = 0;
      if (dueDate && dueDate < today){
        daysOverdue = Math.floor((today.getTime() - dueDate.getTime()) / 86400000);
      }
      rows.push({
        id:           _s(r[BC.ID]),
        vendor:       _s(r[BC.VEND_NAME]),
        billNo:       _s(r[BC.BILL_NO]),
        billDate:     r[BC.BILL_DATE],
        dueDate:      r[BC.DUE_DATE],
        amount:       _n(r[BC.AMOUNT]),
        outstanding:  outstanding,
        status:       status,
        daysOverdue:  daysOverdue
      });
      total += outstanding;
    });
    // Sort overdue first (most urgent), then by outstanding amount desc
    rows.sort(function(a,b){
      if (a.daysOverdue !== b.daysOverdue) return b.daysOverdue - a.daysOverdue;
      return b.outstanding - a.outstanding;
    });
    var sliced = rows.slice(0, limit);
    return _envOk({ rows: sliced, total: _round(total), count: rows.length });
  } catch(e){ return _envErr(e); }
}

// QUICK PAY BILL — single-click full payment from the dashboard.
//   Creates a Bill_Payments row for the outstanding amount and marks the bill Paid.
//   Returns {success, message, paymentId}
function quickPayBill(billId, paymentMode){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    if (!_s(billId)) return {success:false, message:'Bill ID required.'};
    var ss = _getSS();
    var bills = ss.getSheetByName(AP.BILLS);
    if (!bills || !_hasData(bills)) return {success:false, message:'Bills sheet not found.'};
    var rowIdx = -1;
    var rowData = null;
    var data = bills.getRange(2,1,bills.getLastRow()-1,BC_TOTAL).getValues();
    for (var i=0;i<data.length;i++){
      if (_s(data[i][BC.ID]) === _s(billId)){
        rowIdx = i + 2;
        rowData = data[i];
        break;
      }
    }
    if (rowIdx < 0) return {success:false, message:'Bill not found.'};
    var outstanding = _n(rowData[BC.OUTSTANDING]);
    if (outstanding <= 0) return {success:false, message:'Nothing outstanding on this bill.'};

    // Apply full payment
    var newPaid = _round(_n(rowData[BC.PAID]) + outstanding);
    bills.getRange(rowIdx, BC.OUTSTANDING+1).setValue(0);
    bills.getRange(rowIdx, BC.PAID+1).setValue(newPaid);
    bills.getRange(rowIdx, BC.STATUS+1).setValue(BILL_STATUS.PAID);

    // Record the payment row
    var payments = ss.getSheetByName(AP.PAYMENTS);
    if (!payments){ setupBillsModule(); payments = ss.getSheetByName(AP.PAYMENTS); }
    var payId = _nextBillPaymentId();
    var prow = new Array(BPMC_TOTAL).fill('');
    prow[BPMC.ID]        = payId;
    prow[BPMC.VEND_ID]   = _s(rowData[BC.VEND_ID]);
    prow[BPMC.VEND_NAME] = _s(rowData[BC.VEND_NAME]);
    prow[BPMC.DATE]      = new Date();
    prow[BPMC.FY]        = _fyFromDate(new Date());
    prow[BPMC.AMOUNT]    = outstanding;
    prow[BPMC.MODE]      = _s(paymentMode) || 'Cash';
    prow[BPMC.REF]       = '';
    prow[BPMC.SOURCE]    = '';
    prow[BPMC.NOTES]     = 'Quick-paid from dashboard';
    prow[BPMC.APPLIED]   = _s(billId) + ':' + outstanding;
    prow[BPMC.CREATED]   = new Date();
    payments.appendRow(prow);
    SpreadsheetApp.flush();

    return {success:true, message:'Bill ' + billId + ' paid in full', paymentId: payId};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// CANCEL BILL — marks Status = 'Cancelled' (soft delete equivalent).
function cancelBill(billId){
  try {
    if (!_s(billId)) return {success:false, message:'Bill ID required.'};
    var sheet = _getSS().getSheetByName(AP.BILLS);
    if (!sheet || !_hasData(sheet)) return {success:false, message:'Bills sheet not found.'};
    var data = sheet.getRange(2,1,sheet.getLastRow()-1,BC_TOTAL).getValues();
    for (var i=0;i<data.length;i++){
      if (_s(data[i][BC.ID]) === _s(billId)){
        sheet.getRange(i+2, BC.STATUS+1).setValue(BILL_STATUS.CANCELLED);
        sheet.getRange(i+2, BC.OUTSTANDING+1).setValue(0);
        return {success:true, message:'Bill ' + billId + ' cancelled'};
      }
    }
    return {success:false, message:'Bill not found.'};
  } catch(err){ return {success:false, message:err.message}; }
}

// ============================================================
// Helpers
// ============================================================
function _findPartyById(id){
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(M.PARTY);
    if (!sheet || !_hasData(sheet)) return null;
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][0]) === _s(id)){
        return { id:_s(rows[i][0]), type:_s(rows[i][1]), name:_s(rows[i][2]),
                 gstin:_s(rows[i][3]), phone:_s(rows[i][5]) };
      }
    }
  } catch(e){}
  return null;
}
// ============================================================
// upgradePaymentSchemas — Tier 2.1 migration
// Widens Bill_Payments (14→19) and Payments (15→20) with the 5 audit cols:
//   Cancelled_On, Cancelled_By, Cancellation_Reason,
//   Last_Modified_On, Last_Modified_By
// Idempotent — skips a sheet that's already wide enough.
// User must run this once from the Apps Script editor after deploying Tier 2.
// ============================================================
function upgradePaymentSchemas(){
  var ss = _getSS();
  var report = { bills: 'no change', invoices: 'no change' };

  // Bill_Payments
  var bp = ss.getSheetByName(AP.PAYMENTS);
  if (bp){
    if (bp.getLastColumn() < BPMC_TOTAL){
      var start = bp.getLastColumn() + 1;
      var newH = BPM_HEADERS.slice(start - 1);
      bp.getRange(1, start, 1, newH.length).setValues([newH]).setFontWeight('bold');
      newH.forEach(function(_, i){ bp.setColumnWidth(start + i, BPM_WIDTHS[start - 1 + i] || 160); });
      report.bills = 'widened ' + (BPMC_TOTAL - (start - 1)) + ' col(s)';
    }
  } else {
    report.bills = 'sheet missing';
  }

  // Payments (customer AR)
  var pm = ss.getSheetByName(AR.PAYMENTS);
  if (pm){
    if (pm.getLastColumn() < PMC_TOTAL){
      var start2 = pm.getLastColumn() + 1;
      var newH2 = PAYMENT_HEADERS.slice(start2 - 1);
      pm.getRange(1, start2, 1, newH2.length).setValues([newH2]).setFontWeight('bold');
      newH2.forEach(function(_, i){ pm.setColumnWidth(start2 + i, PAYMENT_WIDTHS[start2 - 1 + i] || 160); });
      report.invoices = 'widened ' + (PMC_TOTAL - (start2 - 1)) + ' col(s)';
    }
  } else {
    report.invoices = 'sheet missing';
  }

  var msg = 'Payment audit schema upgrade:\n• Bill_Payments — ' + report.bills +
            '\n• Payments — ' + report.invoices +
            '\n\nOld rows leave the new columns empty (no writes to existing data).';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return { success:true, report:report };
}

// ============================================================
// PAYMENT EDIT / DELETE / LIST — Tier 2.2
// ============================================================

// Applied_Bills format: "BILL-042:5000 · BILL-043:3000"
// Payments.Applied_Invoices format: "3S/2526/045:5000; 3S/2526/046:3000"
// Both delimit chunks by ' · ' / '; ' / ','. Parser is tolerant.
function _parseAppliedString(s){
  s = _s(s);
  if (!s || s === '(advance)') return [];
  // Split on any of · ; ,   then trim + parse "REF:AMOUNT"
  return s.split(/[·;,]/).map(function(chunk){
    var t = chunk.trim(); if (!t) return null;
    var m = t.match(/^(.+?):\s*([-\d.]+)\s*$/);
    if (!m) return null;
    return { ref: m[1].trim(), amount: parseFloat(m[2]) || 0 };
  }).filter(Boolean);
}

// Quick project lookup for a bill — used by the payment edit modal.
function _getBillProject(billId){
  try {
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    if (!billsSheet || billsSheet.getLastRow() < 2) return '';
    var w = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
    var rows = billsSheet.getRange(2,1,billsSheet.getLastRow()-1,w).getValues();
    for (var i=0;i<rows.length;i++){
      // Match by Bill_ID or by Bill_No — Applied_Bills stores the Bill_No form.
      if (_s(rows[i][BC.ID]) === _s(billId) || _s(rows[i][BC.BILL_NO]) === _s(billId)){
        return w > BC.PROJECT ? _s(rows[i][BC.PROJECT]) : '';
      }
    }
    return '';
  } catch(e){ return ''; }
}

// updateBillPayment(paymentId, patches)
//   Editable: paymentDate, paymentMode, reference, source, notes, tdsAmount.
//   Amount / applied bills / vendor NOT editable (see plan — would require
//   FIFO reversal). tdsAmount edit auto-updates bankAmount = amount − tds.
function updateBillPayment(paymentId, patches){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    paymentId = _s(paymentId);
    if (!paymentId) return {success:false, message:'Payment ID required.'};
    patches = patches || {};
    var sheet = _getSS().getSheetByName(AP.PAYMENTS);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'Bill_Payments empty.'};
    var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    var rowNum = -1;
    for (var i=0;i<ids.length;i++){ if (_s(ids[i][0]) === paymentId){ rowNum = i+2; break; } }
    if (rowNum < 0) return {success:false, message:'Payment '+paymentId+' not found.'};

    // Refuse to edit a cancelled payment
    var cancelled = _s(sheet.getRange(rowNum, BPMC.CANCELLED_ON+1).getValue());
    if (cancelled) return {success:false, message:'Cannot edit a cancelled payment. Delete then re-record.'};

    // Whitelist writes
    if (patches.hasOwnProperty('paymentDate') && patches.paymentDate)
      sheet.getRange(rowNum, BPMC.DATE+1).setValue(_asDate(patches.paymentDate));
    if (patches.hasOwnProperty('paymentMode'))
      sheet.getRange(rowNum, BPMC.MODE+1).setValue(_s(patches.paymentMode));
    if (patches.hasOwnProperty('reference'))
      sheet.getRange(rowNum, BPMC.REF+1).setValue(_s(patches.reference));
    if (patches.hasOwnProperty('source'))
      sheet.getRange(rowNum, BPMC.SOURCE+1).setValue(_s(patches.source));
    if (patches.hasOwnProperty('notes'))
      sheet.getRange(rowNum, BPMC.NOTES+1).setValue(_s(patches.notes));
    if (patches.hasOwnProperty('project') && sheet.getLastColumn() > BPMC.PROJECT)
      sheet.getRange(rowNum, BPMC.PROJECT+1).setValue(_s(patches.project));
    if (patches.hasOwnProperty('tdsAmount')){
      var newTds = _n(patches.tdsAmount);
      if (newTds < 0) newTds = 0;
      var amount = _n(sheet.getRange(rowNum, BPMC.AMOUNT+1).getValue());
      if (newTds > amount) return {success:false, message:'TDS cannot exceed payment amount.'};
      sheet.getRange(rowNum, BPMC.TDS_AMT+1).setValue(newTds);
      sheet.getRange(rowNum, BPMC.BANK_AMT+1).setValue(_round(amount - newTds));
    }

    // Optional: unlink this payment from every bill it's currently applied to.
    // Reverses the effect on each linked bill (Paid −, Outstanding +, Status
    // recomputed), clears Applied_Bills. Payment row stays alive with the same
    // ID — use this before re-linking to correct bills.
    var linkMsg = '';
    if (patches.unlink === true){
      var billsSheetU = _getSS().getSheetByName(AP.BILLS);
      var rowU = sheet.getRange(rowNum, 1, 1, Math.min(sheet.getLastColumn(), BPMC_TOTAL)).getValues()[0];
      var restoredU = _reverseAppliedBills(rowU, billsSheetU);
      sheet.getRange(rowNum, BPMC.APPLIED+1).setValue('');
      linkMsg = restoredU.length
        ? ' Unlinked; restored ' + restoredU.join(', ') + '.'
        : ' Nothing was linked.';
    }

    // Optional: retro-link this payment to a bill that didn't exist at record time.
    // Applies min(unallocated payment, bill outstanding) to the bill and appends
    // "BILL-NO:amt" to Applied_Bills. Only same-vendor bills allowed.
    if (patches.hasOwnProperty('linkBill') && _s(patches.linkBill)){
      var linkBillNo = _s(patches.linkBill);
      var payVendor  = _s(sheet.getRange(rowNum, BPMC.VEND_ID+1).getValue());
      var payAmount  = _n(sheet.getRange(rowNum, BPMC.AMOUNT+1).getValue());
      var appliedStr = _s(sheet.getRange(rowNum, BPMC.APPLIED+1).getValue());
      var alreadyAllocated = _parseAppliedString(appliedStr).reduce(function(s,c){ return s + c.amount; }, 0);
      var unallocated = _round(payAmount - alreadyAllocated);
      if (unallocated <= 0.001){
        return {success:false, message:'Payment is already fully allocated to bills — nothing left to link.'};
      }
      // Find the bill; must belong to the same vendor and be open.
      var billsSheet = _getSS().getSheetByName(AP.BILLS);
      if (!billsSheet) return {success:false, message:'Bills sheet not found.'};
      var bW = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
      var bRows = billsSheet.getRange(2,1,billsSheet.getLastRow()-1,bW).getValues();
      // Bill numbers are vendor-scoped, not globally unique — two vendors can
      // each have a bill labelled "1". Silently skip rows for a different
      // vendor or cancelled ones and keep looking; only error if no matching
      // bill exists for THIS vendor.
      var billRow = -1, billOutstanding = 0, billPaid = 0, billAmount = 0;
      for (var i=0;i<bRows.length;i++){
        if (_s(bRows[i][BC.BILL_NO]) !== linkBillNo && _s(bRows[i][BC.ID]) !== linkBillNo) continue;
        if (_s(bRows[i][BC.VEND_ID]) !== payVendor) continue;
        if (_s(bRows[i][BC.STATUS]) === BILL_STATUS.CANCELLED) continue;
        billRow         = i + 2;
        billAmount      = _n(bRows[i][BC.AMOUNT]);
        billPaid        = _n(bRows[i][BC.PAID]);
        billOutstanding = _n(bRows[i][BC.OUTSTANDING]);
        break;
      }
      if (billRow < 0) return {success:false, message:'Bill ' + linkBillNo + ' not found for this vendor (or is cancelled).'};
      if (billOutstanding <= 0.001) return {success:false, message:'Bill ' + linkBillNo + ' has no outstanding.'};

      var apply = _round(Math.min(unallocated, billOutstanding));
      var newPaid  = _round(billPaid + apply);
      var newOut   = _round(billOutstanding - apply);
      var newStat  = (newOut <= 0.001) ? BILL_STATUS.PAID : BILL_STATUS.PARTIAL;
      billsSheet.getRange(billRow, BC.PAID+1).setValue(newPaid);
      billsSheet.getRange(billRow, BC.OUTSTANDING+1).setValue(newOut);
      billsSheet.getRange(billRow, BC.STATUS+1).setValue(newStat);

      var newChunk = linkBillNo + ':' + apply;
      var newApplied = appliedStr ? (appliedStr + ' · ' + newChunk) : newChunk;
      sheet.getRange(rowNum, BPMC.APPLIED+1).setValue(newApplied);
      linkMsg = ' Linked ₹' + apply + ' to ' + linkBillNo + '.';
    }

    // Audit stamp
    sheet.getRange(rowNum, BPMC.LAST_MODIFIED_ON+1)
      .setValue(Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss'));
    sheet.getRange(rowNum, BPMC.LAST_MODIFIED_BY+1).setValue(_currentUser());

    return {success:true, paymentId:paymentId, message:'Payment updated.' + linkMsg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Reverse the effect of a payment row on every bill it was applied to.
// Subtracts from Paid, adds to Outstanding, recomputes Status. Returns an
// array of human-readable "BILL-XXX (+₹amt)" strings for the caller's message.
// Callers: deleteBillPayment (full cancel), updateBillPayment with unlink:true.
function _reverseAppliedBills(row, billsSheet){
  var restored = [];
  var applied = _parseAppliedString(row[BPMC.APPLIED]);
  if (!applied.length || !billsSheet || billsSheet.getLastRow() < 2) return restored;
  var bw = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
  var brows = billsSheet.getRange(2, 1, billsSheet.getLastRow()-1, bw).getValues();
  applied.forEach(function(chunk){
    for (var j=0;j<brows.length;j++){
      if (_s(brows[j][BC.BILL_NO]) === chunk.ref || _s(brows[j][BC.ID]) === chunk.ref){
        var brn = j + 2;
        var curPaid = _n(brows[j][BC.PAID]);
        var curOut  = _n(brows[j][BC.OUTSTANDING]);
        var newPaid = _round(curPaid - chunk.amount);
        var newOut  = _round(curOut  + chunk.amount);
        if (newPaid < 0) newPaid = 0;
        var newStatus = (newPaid <= 0.001) ? BILL_STATUS.UNPAID
                      : (newOut  <= 0.001) ? BILL_STATUS.PAID
                      :                       BILL_STATUS.PARTIAL;
        billsSheet.getRange(brn, BC.PAID+1).setValue(newPaid);
        billsSheet.getRange(brn, BC.OUTSTANDING+1).setValue(newOut);
        billsSheet.getRange(brn, BC.STATUS+1).setValue(newStatus);
        restored.push(chunk.ref + ' (+₹' + chunk.amount + ')');
        break;
      }
    }
  });
  return restored;
}

// createSplitBills(baseData, splits) — creates N bill rows from one physical
// vendor bill, distributing Amount / Taxable / CGST / SGST / IGST across
// projects. Each split becomes an independent Bill row with the same Bill_No
// suffixed " (A)", " (B)", … so it's still traceable to the paper bill.
// splits = [{project, amount}] — amounts must sum (± ₹1) to baseData.amount.
// No outer lock: delegates to createBill per split (each takes its own lock).
function createSplitBills(baseData, splits){
  splits = splits || [];
  if (!splits.length) return {success:false, message:'No splits provided.'};
  // Safety net: split bills' whole point is per-project tagging. If the Bills
  // sheet is narrower than BC_TOTAL, appendRow silently truncates the PROJECT
  // column and every split loses its project tag. setupBillsModule widens the
  // sheet idempotently (no-op if already at full width).
  try { setupBillsModule(); } catch(e){}
  var total = _n(baseData.amount) || _n(baseData.invoiceAmount);
  var taxable = _n(baseData.taxable) || total;
  var sumSplits = splits.reduce(function(s, x){ return s + _n(x && x.amount); }, 0);
  if (Math.abs(sumSplits - total) > 1) {
    return {success:false, message:'Splits (₹' + _round(sumSplits) + ') must sum to bill total (₹' + _round(total) + '). Off by ₹' + _round(Math.abs(sumSplits - total)) + '.'};
  }
  for (var i=0;i<splits.length;i++){
    if (!_s(splits[i].project)) return {success:false, message:'Split #' + (i+1) + ' has no project.'};
    if (_n(splits[i].amount) <= 0) return {success:false, message:'Split #' + (i+1) + ' has non-positive amount.'};
  }
  // Letter suffix: A, B, ..., Z, AA, AB, ...
  function letter(n){
    var s = '';
    do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
    return s;
  }
  var results = [], okIds = [], failures = [];
  splits.forEach(function(sp, idx){
    var ratio = _n(sp.amount) / total;
    var childData = {};
    for (var k in baseData) if (baseData.hasOwnProperty(k)) childData[k] = baseData[k];
    childData.amount        = _n(sp.amount);
    childData.invoiceAmount = _n(sp.amount);
    childData.taxable       = _round(taxable * ratio);
    childData.cgst          = _round(_n(baseData.cgst) * ratio);
    childData.sgst          = _round(_n(baseData.sgst) * ratio);
    childData.igst          = _round(_n(baseData.igst) * ratio);
    childData.project       = _s(sp.project);
    childData.billNo        = _s(baseData.billNo) + ' (' + letter(idx) + ')';
    childData.notes         = _s(baseData.notes) + (baseData.notes ? ' · ' : '') + 'Split ' + letter(idx) + ' of ' + splits.length + ' from ' + _s(baseData.billNo);
    childData.paidNow       = 0;   // never split the "amount already paid" — record it manually against a specific split if needed
    var res = createBill(childData);
    results.push({project: sp.project, amount: sp.amount, billNo: childData.billNo, success: !!(res && res.success), billId: res && res.billId, message: res && res.message});
    if (res && res.success) okIds.push(res.billId); else failures.push(childData.billNo + ': ' + (res && res.message));
  });
  return {
    success: failures.length === 0,
    createdCount: okIds.length,
    failedCount: failures.length,
    results: results,
    message: okIds.length + ' split bill(s) created' + (failures.length ? '; ' + failures.length + ' failed: ' + failures.join('; ') : '.')
  };
}

// splitExistingBill(billId, splits) — retro-split an already-recorded bill.
// Refuses if the bill has any payment applied to it (operator must unlink
// via Vendor Ledger → Bulk fix first). On success, soft-cancels the original
// and creates N new split bills that carry the same vendor / date / GST math.
function splitExistingBill(billId, splits){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    billId = _s(billId);
    if (!billId) return {success:false, message:'Bill ID required.'};
    // Same safety net as createSplitBills — widen Bills sheet if it's narrower
    // than the current schema before we start writing.
    try { setupBillsModule(); } catch(e){}
    var ss = _getSS(), bSheet = ss.getSheetByName(AP.BILLS);
    if (!bSheet || bSheet.getLastRow() < 2) return {success:false, message:'Bills sheet empty.'};
    var w = Math.min(bSheet.getLastColumn(), BC_TOTAL);
    var rows = bSheet.getRange(2, 1, bSheet.getLastRow()-1, w).getValues();
    var idx = -1, r = null;
    for (var i=0;i<rows.length;i++){ if (_s(rows[i][BC.ID]) === billId){ idx = i; r = rows[i]; break; } }
    if (idx < 0) return {success:false, message:'Bill ' + billId + ' not found.'};
    if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return {success:false, message:'Bill is already cancelled.'};
    // Refuse if any active payment has this bill in its Applied_Bills.
    var pSheet = ss.getSheetByName(AP.PAYMENTS);
    if (pSheet && pSheet.getLastRow() > 1){
      var pW = Math.min(pSheet.getLastColumn(), BPMC_TOTAL);
      var pRows = pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues();
      var billNo = _s(r[BC.BILL_NO]);
      for (var j=0;j<pRows.length;j++){
        if (pW > BPMC.CANCELLED_ON && _s(pRows[j][BPMC.CANCELLED_ON])) continue;
        var applied = _parseAppliedString(_s(pRows[j][BPMC.APPLIED]));
        for (var k=0;k<applied.length;k++){
          if (applied[k].ref === billId || applied[k].ref === billNo){
            return {success:false, message:'Bill ' + (billNo || billId) + ' has payment(s) linked to it. Go to Vendor Ledger → Bulk fix → Unlink first, then retry the split.'};
          }
        }
      }
    }
    // Build baseData from the row.
    var baseData = {
      vendorId: _s(r[BC.VEND_ID]),
      billNo:   _s(r[BC.BILL_NO]),
      billDate: r[BC.BILL_DATE],
      dueDate:  r[BC.DUE_DATE],
      amount:   _n(r[BC.AMOUNT]),
      taxable:  _n(r[BC.TAXABLE]),
      gst:      _n(r[BC.CGST]) + _n(r[BC.SGST]) + _n(r[BC.IGST]) > 0,
      gstRate:  _n(r[BC.GST_RATE]),
      cgst:     _n(r[BC.CGST]),
      sgst:     _n(r[BC.SGST]),
      igst:     _n(r[BC.IGST]),
      category: w > BC.CATEGORY ? _s(r[BC.CATEGORY]) : '',
      proofLink: w > BC.PROOF ? _s(r[BC.PROOF]) : '',
      notes:    w > BC.NOTES ? _s(r[BC.NOTES]) : ''
    };
    // Soft-cancel the original so the ledger doesn't double-count.
    bSheet.getRange(idx + 2, BC.STATUS + 1).setValue(BILL_STATUS.CANCELLED);
    bSheet.getRange(idx + 2, BC.NOTES + 1).setValue(_s(baseData.notes) + (baseData.notes ? ' · ' : '') + 'CANCELLED — split into ' + splits.length + ' project bills on ' + Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd'));
    // Release before delegating so createSplitBills can acquire per-child.
    try { lock.releaseLock(); } catch(e){}
    var res = createSplitBills(baseData, splits);
    if (!res.success){
      // Rollback: restore the original status.
      try {
        var l2 = LockService.getScriptLock(); l2.waitLock(5000);
        bSheet.getRange(idx + 2, BC.STATUS + 1).setValue(_s(r[BC.STATUS]) || BILL_STATUS.UNPAID);
        l2.releaseLock();
      } catch(e){}
      return {success:false, message:'Split failed, original restored. Detail: ' + res.message};
    }
    return {success:true, cancelledBillId: billId, createdCount: res.createdCount, results: res.results, message: 'Original bill ' + (baseData.billNo || billId) + ' cancelled; ' + res.createdCount + ' split bills created.'};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// unlinkAllVendorPayments(vendorId) — vendor-scoped bulk unlink.
// For every active (not cancelled) payment belonging to this vendor that
// currently has any Applied_Bills, reverses the effect on each linked bill
// and clears Applied_Bills. Payment rows stay alive with the same IDs.
// Use this to "start fresh" before re-linking payments to the correct bills.
function unlinkAllVendorPayments(vendorId){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    vendorId = _s(vendorId);
    if (!vendorId) return {success:false, message:'Vendor required.'};
    var ss = _getSS();
    var sheet = ss.getSheetByName(AP.PAYMENTS);
    var billsSheet = ss.getSheetByName(AP.BILLS);
    if (!sheet || sheet.getLastRow() < 2) return {success:true, unlinked:0, restored:0, message:'No payments.'};
    var w = Math.min(sheet.getLastColumn(), BPMC_TOTAL);
    var all = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var stamp = Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
    var user = _currentUser();
    var unlinked = 0, restoredCount = 0;
    for (var i=0;i<all.length;i++){
      var r = all[i];
      if (_s(r[BPMC.VEND_ID]) !== vendorId) continue;
      if (w > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) continue;
      if (!_s(r[BPMC.APPLIED])) continue;
      var restored = _reverseAppliedBills(r, billsSheet);
      sheet.getRange(i+2, BPMC.APPLIED+1).setValue('');
      sheet.getRange(i+2, BPMC.LAST_MODIFIED_ON+1).setValue(stamp);
      sheet.getRange(i+2, BPMC.LAST_MODIFIED_BY+1).setValue(user);
      unlinked++;
      restoredCount += restored.length;
    }
    return {
      success:true, unlinked: unlinked, restored: restoredCount,
      message: 'Unlinked ' + unlinked + ' payment(s); restored balance due on ' + restoredCount + ' bill link(s).'
    };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// linkVendorPayments(assignments) — batch re-link. Each assignment =
// {paymentId, billNo}. Delegates row-by-row to updateBillPayment so all
// existing guardrails (same-vendor check, cancelled check, allocation
// cap = min(unallocated, bill outstanding)) still apply. Skips assignments
// with a blank billNo so the UI can pass rows the user chose to leave alone.
function linkVendorPayments(assignments){
  assignments = assignments || [];
  if (!assignments.length) return {success:false, message:'No assignments provided.'};
  var results = [], okCount = 0, failCount = 0, skipCount = 0;
  assignments.forEach(function(a){
    var payId = _s(a && a.paymentId), billNo = _s(a && a.billNo);
    if (!payId){ skipCount++; return; }
    if (!billNo){ skipCount++; return; }
    var res = updateBillPayment(payId, {linkBill: billNo});
    var ok = !!(res && res.success);
    results.push({paymentId: payId, billNo: billNo, success: ok, message: res && res.message});
    if (ok) okCount++; else failCount++;
  });
  return {
    success: failCount === 0,
    ok: okCount, failed: failCount, skipped: skipCount,
    results: results,
    message: okCount + ' linked, ' + failCount + ' failed' + (skipCount ? ', ' + skipCount + ' skipped' : '') + '.'
  };
}

// deleteBillPayment(paymentId, reason) — soft-delete + FIFO reversal
function deleteBillPayment(paymentId, reason){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    paymentId = _s(paymentId);
    if (!paymentId) return {success:false, message:'Payment ID required.'};
    var ss = _getSS();
    var sheet = ss.getSheetByName(AP.PAYMENTS);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'Bill_Payments empty.'};

    var w = Math.min(sheet.getLastColumn(), BPMC_TOTAL);
    var all = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var rowNum = -1, row = null;
    for (var i=0;i<all.length;i++){ if (_s(all[i][BPMC.ID]) === paymentId){ rowNum = i+2; row = all[i]; break; } }
    if (rowNum < 0) return {success:false, message:'Payment '+paymentId+' not found.'};
    if (w > BPMC.CANCELLED_ON && _s(row[BPMC.CANCELLED_ON])){
      return {success:false, message:'Payment already cancelled.'};
    }

    // Reverse the effect on every linked bill via shared helper.
    var restoredBills = _reverseAppliedBills(row, ss.getSheetByName(AP.BILLS));

    // Soft-delete stamp
    var stamp = Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(rowNum, BPMC.CANCELLED_ON+1).setValue(stamp);
    sheet.getRange(rowNum, BPMC.CANCELLED_BY+1).setValue(_currentUser());
    sheet.getRange(rowNum, BPMC.CANCELLATION_REASON+1).setValue(_s(reason));

    return {
      success: true,
      paymentId: paymentId,
      reversed: restoredBills.length,
      message: 'Payment ' + paymentId + ' cancelled' +
               (restoredBills.length ? '; restored ' + restoredBills.join(', ') : '')
    };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// listBillPayments({vendorId?, fy?, search?, includeCancelled?, limit?})
function listBillPayments(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(AP.PAYMENTS);
    if (!sheet || sheet.getLastRow() < 2) return _envOk({rows:[]});
    var w = Math.min(sheet.getLastColumn(), BPMC_TOTAL);
    var vals = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    // Global BILL_ID → {ref, status} lookup, one pass. Used to enrich the
    // Applied_Bills string (BILL-0005:1000 → BILL-0005 (INV-234) [cancelled]:1000)
    // so operators see the vendor's ref AND whether the linked bill has since
    // been cancelled — everywhere this list feeds (Payments Made, bulk-fix).
    var billLookup = {};
    var bs = _getSS().getSheetByName(AP.BILLS);
    if (bs && _hasData(bs)){
      var bw = Math.min(bs.getLastColumn(), BC_TOTAL);
      var brows = bs.getRange(2,1,bs.getLastRow()-1,bw).getValues();
      for (var bi=0;bi<brows.length;bi++){
        var bid = _s(brows[bi][BC.ID]); if (!bid) continue;
        billLookup[bid] = { ref: _s(brows[bi][BC.BILL_NO]), status: _s(brows[bi][BC.STATUS]) };
      }
    }
    var wantVendor  = _s(filters.vendorId);
    var wantFy      = _s(filters.fy);
    var q           = _s(filters.search).toLowerCase();
    var incCancel   = !!filters.includeCancelled;
    var limit       = _n(filters.limit) || 300;
    var out = [];
    for (var i=0;i<vals.length;i++){
      var r = vals[i];
      var id = _s(r[BPMC.ID]); if (!id) continue;
      var cancelledOn = w > BPMC.CANCELLED_ON ? _s(r[BPMC.CANCELLED_ON]) : '';
      if (cancelledOn && !incCancel) continue;
      if (wantVendor && _s(r[BPMC.VEND_ID]) !== wantVendor) continue;
      if (wantFy && _s(r[BPMC.FY]) !== wantFy) continue;
      var blob = (id + ' ' + _s(r[BPMC.VEND_NAME]) + ' ' + _s(r[BPMC.REF]) + ' ' +
                  _s(r[BPMC.APPLIED]) + ' ' + _s(r[BPMC.NOTES])).toLowerCase();
      if (q && blob.indexOf(q) === -1) continue;
      out.push({
        id: id,
        vendorId: _s(r[BPMC.VEND_ID]),
        vendor: _s(r[BPMC.VEND_NAME]),
        date: _safeDateStr(r[BPMC.DATE]),
        fy: _s(r[BPMC.FY]),
        amount: _n(r[BPMC.AMOUNT]),
        mode: _s(r[BPMC.MODE]),
        reference: _s(r[BPMC.REF]),
        source: _s(r[BPMC.SOURCE]),
        applied: _enrichAppliedWithVendorBillNo(_s(r[BPMC.APPLIED]), billLookup),
        tdsAmount: _n(r[BPMC.TDS_AMT]),
        bankAmount: _n(r[BPMC.BANK_AMT]),
        notes: _s(r[BPMC.NOTES]),
        cancelled: !!cancelledOn,
        cancelledOn: cancelledOn,
        cancelledBy: w > BPMC.CANCELLED_BY ? _s(r[BPMC.CANCELLED_BY]) : '',
        cancellationReason: w > BPMC.CANCELLATION_REASON ? _s(r[BPMC.CANCELLATION_REASON]) : '',
        lastModifiedOn: w > BPMC.LAST_MODIFIED_ON ? _s(r[BPMC.LAST_MODIFIED_ON]) : '',
        lastModifiedBy: w > BPMC.LAST_MODIFIED_BY ? _s(r[BPMC.LAST_MODIFIED_BY]) : '',
        project:        w > BPMC.PROJECT ? _s(r[BPMC.PROJECT]) : ''
      });
    }
    // Newest first
    out.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
    return _envOk({ rows: out.slice(0, limit), count: out.length });
  } catch(e){ return _envErr(e); }
}
// ============================================================
// Projects — vendor-bill tick ("count in project cost")
// (Batch A — consumed by the Projects tab)
// ============================================================

// 'Yes' / 'No' flag resolver. Safe when the row predates the column
// (returns false), and treats blank as No.
function _billCountsForProject(row){
  if (row.length <= BC.COUNT_IN_PROJECT_COST) return false;
  return _s(row[BC.COUNT_IN_PROJECT_COST]) === 'Yes';
}

// Flip the tick on a single bill. Called from the Projects tab.
function toggleBillInProjectCost(billId, countFlag){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var ss = _getSS();
    var sheet = ss.getSheetByName(AP.BILLS);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'Bills sheet empty.'};
    var ids = sheet.getRange(2, BC.ID+1, sheet.getLastRow()-1, 1).getValues();
    for (var i=0; i<ids.length; i++){
      if (_s(ids[i][0]) === _s(billId)){
        var value = (countFlag === true || countFlag === 'Yes') ? 'Yes' : 'No';
        sheet.getRange(i+2, BC.COUNT_IN_PROJECT_COST+1).setValue(value);
        SpreadsheetApp.flush();
        return {success:true, billId:billId, countFlag:value};
      }
    }
    return {success:false, message:'Bill '+billId+' not found.'};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Vendor bills tagged to this project (across all FYs) + their tick state.
// Rendered row-wise in the Projects tab with a checkbox per row.
function getProjectDetailBills(project){
  try {
    if (!project) return _envOk({bills:[]});
    var ss = _getSS();
    var sheet = ss.getSheetByName(AP.BILLS);
    if (!sheet || sheet.getLastRow() < 2) return _envOk({bills:[]});
    var w = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var out = [];
    for (var i=0; i<data.length; i++){
      var r = data[i];
      if (_s(r[BC.PROJECT]) !== project) continue;
      if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) continue;
      out.push({
        id:         _s(r[BC.ID]),
        billNo:     _s(r[BC.BILL_NO]),
        vendor:     _s(r[BC.VEND_NAME]),
        billDate:   _safeDateStr(r[BC.BILL_DATE]),
        dueDate:    _safeDateStr(r[BC.DUE_DATE]),
        amount:     _n(r[BC.AMOUNT]) || 0,
        paid:       _n(r[BC.PAID]) || 0,
        outstanding:_n(r[BC.OUTSTANDING]) || 0,
        status:     _s(r[BC.STATUS]),
        counted:    _billCountsForProject(r),
        category:   _s(r[BC.CATEGORY])
      });
    }
    out.sort(function(a,b){ return (b.billDate||'').localeCompare(a.billDate||''); });
    return _envOk({bills:out, count:out.length});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// BILL ALLOCATIONS (Option B) — carve a single bill across projects
//
// Model: zero, one, or many allocation rows per bill. A bill with zero
// non-deleted allocations behaves exactly as before (project cost rolls up
// via Bill.Project). A bill with one or more allocations overrides that:
// each allocation contributes its own amount to its own project. The sum of
// a bill's allocations must stay ≤ bill AMOUNT (gross). Any shortfall is
// implicitly unallocated — it still counts in expense totals, just not
// against any project.
//
// Public API:
//   getBillAllocations(billId)                 → envelope with array of rows
//   setBillAllocations(billId, allocations[])  → atomic replace, validated
//   deleteBillAllocation(allocationId)         → soft-delete one line
//   _getAllBillAllocationsMap()                → {billId: [{project, amount}]}
//                                                for dashboard rollups
// ============================================================
function _nextBillAllocationId(){
  var sheet = _getSS().getSheetByName(AP.ALLOCATIONS);
  if (!sheet || !_hasData(sheet)) return 'BA-0001';
  var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  var max = 0;
  for (var i=0;i<ids.length;i++){
    var v = _s(ids[i][0]);
    if (v.indexOf('BA-') === 0){
      var n = parseInt(v.substring(3),10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return 'BA-' + String(max+1).padStart(4,'0');
}

// Resolve a bill by its ID; returns {rowIdx, data, found}.
function _findBillRow(billId){
  var sheet = _getSS().getSheetByName(AP.BILLS);
  if (!sheet || sheet.getLastRow() < 2) return {found:false};
  var w = Math.min(sheet.getLastColumn(), BC_TOTAL);
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
  for (var i=0; i<data.length; i++){
    if (_s(data[i][BC.ID]) === _s(billId)){
      return {found:true, rowIdx:i+2, data:data[i], sheet:sheet};
    }
  }
  return {found:false};
}

// Return all non-deleted allocations for one bill, oldest first.
function getBillAllocations(billId){
  try {
    billId = _s(billId);
    if (!billId) return _envOk({billId:'', rows:[], total:0});
    var sheet = _getSS().getSheetByName(AP.ALLOCATIONS);
    if (!sheet || sheet.getLastRow() < 2) return _envOk({billId:billId, rows:[], total:0});
    var w = Math.min(sheet.getLastColumn(), BAC_TOTAL);
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var rows = [], total = 0;
    for (var i=0; i<data.length; i++){
      var r = data[i];
      if (_s(r[BAC.BILL_ID]) !== billId) continue;
      if (_s(r[BAC.DELETED]) === 'Yes') continue;
      rows.push({
        id:       _s(r[BAC.ID]),
        billId:   _s(r[BAC.BILL_ID]),
        billNo:   _s(r[BAC.BILL_NO]),
        project:  _s(r[BAC.PROJECT]),
        amount:   _n(r[BAC.AMOUNT]) || 0,
        notes:    _s(r[BAC.NOTES]),
        createdDate: _safeDateStr(r[BAC.CREATED_DATE])
      });
      total += rows[rows.length-1].amount;
    }
    return _envOk({billId:billId, rows:rows, total:_round(total), count:rows.length});
  } catch(e){ return _envErr(e); }
}

// Atomic replace: validates then wipes existing non-deleted rows for the
// bill, writes the new set. Sum must be ≤ bill AMOUNT. Each project must
// exist in M_Project (soft — allowed if master is missing, logs instead).
// allocations = [{project, amount, notes}]  — ignores entries with amount<=0.
function setBillAllocations(billId, allocations){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    billId = _s(billId);
    if (!billId) return {success:false, message:'Bill ID is required.'};
    allocations = Array.isArray(allocations) ? allocations : [];
    // Normalize + drop blanks
    var clean = [];
    for (var i=0; i<allocations.length; i++){
      var a = allocations[i] || {};
      var proj = _s(a.project).trim();
      var amt  = _n(a.amount) || 0;
      if (!proj || amt <= 0) continue;
      clean.push({project:proj, amount:_round(amt), notes:_s(a.notes)});
    }

    // Validate bill exists + sum ≤ bill AMOUNT
    var bill = _findBillRow(billId);
    if (!bill.found) return {success:false, message:'Bill ' + billId + ' not found.'};
    var billAmt = _n(bill.data[BC.AMOUNT]) || 0;
    var sum = 0;
    for (var j=0; j<clean.length; j++) sum += clean[j].amount;
    if (sum > billAmt + 0.01){
      return {success:false, message:'Allocation total ₹' + _round(sum) +
              ' exceeds bill amount ₹' + _round(billAmt) + '.'};
    }

    // Validate project names (soft — warn in msg, don't fail)
    var validProjects = {};
    try {
      (typeof getProjectNames === 'function' ? (getProjectNames() || []) : []).forEach(function(n){
        validProjects[String(n).trim()] = true;
      });
    } catch(e){}
    var unknown = [];
    clean.forEach(function(a){ if (Object.keys(validProjects).length && !validProjects[a.project]) unknown.push(a.project); });

    // Prepare sheet (idempotent).
    var ss = _getSS();
    var sheet = ss.getSheetByName(AP.ALLOCATIONS);
    if (!sheet){ setupBillsModule(); sheet = ss.getSheetByName(AP.ALLOCATIONS); }
    if (!sheet) return {success:false, message:'Bill_Allocations sheet could not be created.'};

    // Soft-delete existing non-deleted rows for this bill (one bulk write)
    if (sheet.getLastRow() > 1){
      var w = Math.min(sheet.getLastColumn(), BAC_TOTAL);
      var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
      var updates = [];  // [{rowIdx, deletedTo:'Yes'}]
      for (var k=0; k<data.length; k++){
        if (_s(data[k][BAC.BILL_ID]) === billId && _s(data[k][BAC.DELETED]) !== 'Yes'){
          sheet.getRange(k+2, BAC.DELETED+1).setValue('Yes');
        }
      }
    }

    // Append the fresh set
    var now = new Date();
    var who = Session.getActiveUser().getEmail() || '';
    var billNo = _s(bill.data[BC.BILL_NO]);
    var created = [];
    for (var m=0; m<clean.length; m++){
      var id = _nextBillAllocationId();
      var row = new Array(BAC_TOTAL).fill('');
      row[BAC.ID]           = id;
      row[BAC.BILL_ID]      = billId;
      row[BAC.BILL_NO]      = billNo;
      row[BAC.PROJECT]      = clean[m].project;
      row[BAC.AMOUNT]       = clean[m].amount;
      row[BAC.NOTES]        = clean[m].notes;
      row[BAC.CREATED_DATE] = now;
      row[BAC.CREATED_BY]   = who;
      row[BAC.DELETED]      = 'No';
      var wr = _writeRowVerified(sheet, row, BAC.ID, id);
      if (!wr.ok) return {success:false, message:wr.message};
      created.push(id);
    }

    var msg = created.length
      ? 'Saved ' + created.length + ' allocation(s) totalling ₹' + _round(sum) + ' for ' + billId
      : 'All allocations cleared for ' + billId;
    if (unknown.length) msg += ' · Unknown project(s): ' + unknown.join(', ');
    return {success:true, billId:billId, allocationIds:created, total:_round(sum),
            unmatchedProjects:unknown, message:msg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Soft-delete one allocation line (operator may prefer this to setBillAllocations
// when touching a single row).
function deleteBillAllocation(allocationId){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    allocationId = _s(allocationId);
    if (!allocationId) return {success:false, message:'Allocation ID is required.'};
    var sheet = _getSS().getSheetByName(AP.ALLOCATIONS);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'No allocations exist.'};
    var w = Math.min(sheet.getLastColumn(), BAC_TOTAL);
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    for (var i=0; i<data.length; i++){
      if (_s(data[i][BAC.ID]) === allocationId){
        if (_s(data[i][BAC.DELETED]) === 'Yes') return {success:true, message:'Already deleted.'};
        sheet.getRange(i+2, BAC.DELETED+1).setValue('Yes');
        return {success:true, allocationId:allocationId, message:'Allocation ' + allocationId + ' removed.'};
      }
    }
    return {success:false, message:'Allocation ' + allocationId + ' not found.'};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// One-pass read of all non-deleted allocations into {billId: [{project, amount}]}.
// Dashboard rollups call this once per request, then look up each bill.
function _getAllBillAllocationsMap(){
  var out = {};
  var sheet = _getSS().getSheetByName(AP.ALLOCATIONS);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var w = Math.min(sheet.getLastColumn(), BAC_TOTAL);
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
  for (var i=0; i<data.length; i++){
    var r = data[i];
    if (_s(r[BAC.DELETED]) === 'Yes') continue;
    var bid = _s(r[BAC.BILL_ID]);
    var proj = _s(r[BAC.PROJECT]);
    var amt  = _n(r[BAC.AMOUNT]) || 0;
    if (!bid || !proj || amt <= 0) continue;
    if (!out[bid]) out[bid] = [];
    out[bid].push({project:proj, amount:amt});
  }
  return out;
}