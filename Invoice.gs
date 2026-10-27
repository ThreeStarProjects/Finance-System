// ============================================================
// 3 Star Projects — Unified Finance System
// PHASE 3 : INVOICES + ACCOUNTS RECEIVABLE
// File: Invoices.gs   (depends on Masters.gs, Transactions.gs)
//
// Sheets:
//   Orders    — optional order head (invoice can link to one, or skip it)
//   Invoices  — sales invoices. Real 3S/FY/nnn numbering. GST split.
//               Status column (no more cancellation typed into the name).
//   Payments  — customer receipts. FIFO-applied to oldest open invoices;
//               leftover becomes an advance (credit) auto-applied to future invoices.
//   Returns   — material returns (credit against an invoice).
//
// Customer ledger is COMPUTED on demand (getCustomerLedger) from these tabs,
// sorted by date — never a stored running balance, so back-dated entries can't
// corrupt it. Balance > 0 = customer owes us; < 0 = advance/credit with us.
//
// Entry point (optional): run setupAR() to pre-create the four tabs.
// ============================================================

var AR = {
  ORDERS:   'Orders',
  INVOICES: 'Invoices',
  PAYMENTS: 'Payments',
  RETURNS:  'Returns'
};

var INV_STATUS = { UNPAID:'Unpaid', PARTIAL:'Partial', PAID:'Paid', CANCELLED:'Cancelled' };
var ORD_STATUS = { OPEN:'Open', PARTIAL:'Partial', CLOSED:'Closed', CANCELLED:'Cancelled' };

// ── Column maps ─────────────────────────────────────────────
var OC = { ID:0, CUST_ID:1, CUST_NAME:2, AMOUNT:3, DESCRIPTION:4, DELIVERY:5,
           INVOICED:6, PENDING:7, STATUS:8, CREATED:9 };
var OC_TOTAL = 10;
var ORDER_HEADERS = ['Order_ID','Customer_ID','Customer_Name','Order_Amount','Description',
  'Expected_Delivery','Invoiced_Amount','Pending_Amount','Status','Created_Date'];
var ORDER_WIDTHS = [120,110,240,120,260,140,130,130,90,160];

var IC = { NO:0, ORDER_ID:1, CUST_ID:2, CUST_NAME:3, GSTIN:4, INV_DATE:5, DUE_DATE:6, FY:7,
           TAXABLE:8, GST_RATE:9, CGST:10, SGST:11, IGST:12, AMOUNT:13,
           OUTSTANDING:14, PAID:15, TRANSPORT:16, LR:17, PROOF:18, STATUS:19, NOTES:20, CREATED:21,
           PROJECT:22,
           DISC_TYPE:23, DISC_VALUE:24, DISC_APPLIED:25, DISC_AMOUNT:26 };
var IC_TOTAL = 27;
var INVOICE_HEADERS = ['Invoice_No','Order_ID','Customer_ID','Customer_Name','Customer_GSTIN',
  'Invoice_Date','Due_Date','FY','Taxable_Value','GST_Rate','CGST','SGST','IGST','Invoice_Amount',
  'Outstanding','Paid','Transport_Name','LR_Number','Proof_Link','Status','Notes','Created_Date',
  'Project',
  'Discount_Type','Discount_Value','Discount_Applied','Discount_Amount'];
var INVOICE_WIDTHS = [150,120,110,240,150,120,120,80,120,80,90,90,90,130,120,110,150,120,200,100,200,160,180,
                     100,100,110,120];

var PMC = { ID:0, CUST_ID:1, CUST_NAME:2, DATE:3, FY:4, AMOUNT:5, MODE:6, REF:7,
            APPLIED:8, UNAPPLIED:9, NOTES:10, CREATED:11, TDS_AMT:12, BANK_AMT:13, RCV_ACCT:14,
            // Tier 2.1 audit columns
            CANCELLED_ON:15, CANCELLED_BY:16, CANCELLATION_REASON:17,
            LAST_MODIFIED_ON:18, LAST_MODIFIED_BY:19,
            PROJECT:20 };
var PMC_TOTAL = 21;
var PAYMENT_HEADERS = ['Payment_ID','Customer_ID','Customer_Name','Payment_Date','FY','Amount',
  'Payment_Mode','Reference_No','Applied_Invoices','Unapplied_Advance','Notes','Created_Date',
  'TDS_Amount','Bank_Amount','Received_In_Account',
  'Cancelled_On','Cancelled_By','Cancellation_Reason','Last_Modified_On','Last_Modified_By',
  'Project'];
var PAYMENT_WIDTHS = [130,110,240,120,80,120,130,160,320,140,220,160,110,120,160, 160,200,300,160,200, 180];

var RC = { ID:0, INV_NO:1, CUST_ID:2, CUST_NAME:3, DATE:4, FY:5, AMOUNT:6, REASON:7, NOTES:8, CREATED:9 };
var RC_TOTAL = 10;
var RETURN_HEADERS = ['Return_ID','Invoice_No','Customer_ID','Customer_Name','Return_Date','FY',
  'Return_Amount','Reason','Notes','Created_Date'];
var RETURN_WIDTHS = [120,150,110,240,120,80,120,260,220,160];

// ============================================================
// PROVISIONING
// ============================================================
// ============================================================
// backfillPaymentCustomerNames — one-shot repair for existing Payment rows
// where Customer_Name column is blank. Reads Customer_ID and looks up the
// current name in M_Party. Safe to run anytime; idempotent.
// ============================================================
function backfillPaymentCustomerNames(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(AR.PAYMENTS);
  if (!sheet || sheet.getLastRow() < 2){
    var m = 'Payments sheet is empty — nothing to repair.';
    try { SpreadsheetApp.getUi().alert(m); } catch(e){ Logger.log(m); }
    return {success:true, updated:0, message:m};
  }
  var lr = sheet.getLastRow();
  var ids   = sheet.getRange(2, PMC.CUST_ID+1,   lr-1, 1).getValues();
  var names = sheet.getRange(2, PMC.CUST_NAME+1, lr-1, 1).getValues();
  var updated = 0;
  for (var i=0; i<names.length; i++){
    if (_s(names[i][0])) continue;    // already has a name
    var id = _s(ids[i][0]);
    if (!id) continue;
    var nm = _nameForParty(id);
    if (nm){ names[i][0] = nm; updated++; }
  }
  if (updated > 0){
    sheet.getRange(2, PMC.CUST_NAME+1, names.length, 1).setValues(names);
  }
  var msg = 'Backfill complete: filled Customer_Name on ' + updated + ' payment row(s).';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, updated:updated, message:msg};
}

function setupAR(){
  var ss = _getSS();
  var o = _buildSheet(ss, AR.ORDERS,   ORDER_HEADERS,   ORDER_WIDTHS);
  _validateCol(o, OC.STATUS+1, [ORD_STATUS.OPEN,ORD_STATUS.PARTIAL,ORD_STATUS.CLOSED,ORD_STATUS.CANCELLED]);
  var i = _buildSheet(ss, AR.INVOICES, INVOICE_HEADERS, INVOICE_WIDTHS);
  i.getRange(2, IC.FY+1, 5000, 1).setNumberFormat('@STRING@');
  _validateCol(i, IC.STATUS+1, [INV_STATUS.UNPAID,INV_STATUS.PARTIAL,INV_STATUS.PAID,INV_STATUS.CANCELLED]);
  _validateCol(i, IC.DISC_TYPE+1,    ['','Percent','Amount']);
  _validateCol(i, IC.DISC_APPLIED+1, ['','Pre-GST','Post-GST']);
  // Schema upgrade: append any new headers when the sheet pre-dates them.
  // Discount cols stay blank on historical rows (no discount was captured).
  if (i.getLastColumn() < IC_TOTAL){
    var iStart = i.getLastColumn() + 1;
    var iNewHeaders = INVOICE_HEADERS.slice(iStart - 1);
    i.getRange(1, iStart, 1, iNewHeaders.length).setValues([iNewHeaders]).setFontWeight('bold');
  }
  _buildSheet(ss, AR.PAYMENTS, PAYMENT_HEADERS, PAYMENT_WIDTHS);
  // Schema upgrade: append any new headers and back-fill historical rows.
  // TDS_Amount → 0 (no TDS captured historically).
  // Bank_Amount → equal to Amount.
  // Received_In_Account → blank.
  var pSheet = ss.getSheetByName(AR.PAYMENTS);
  if (pSheet && pSheet.getLastColumn() < PMC_TOTAL){
    var startCol = pSheet.getLastColumn() + 1;
    var newHeaders = PAYMENT_HEADERS.slice(startCol - 1);
    pSheet.getRange(1, startCol, 1, newHeaders.length).setValues([newHeaders]).setFontWeight('bold');
    var lr = pSheet.getLastRow();
    if (lr >= 2){
      var amounts = pSheet.getRange(2, PMC.AMOUNT+1, lr-1, 1).getValues();
      // TDS_Amount + Bank_Amount back-fill (only if they're in the new range)
      if (startCol <= PMC.TDS_AMT+1){
        var backfill = amounts.map(function(r){ return [0, _n(r[0])]; });
        pSheet.getRange(2, PMC.TDS_AMT+1, lr-1, 2).setValues(backfill);
      }
      // Received_In_Account stays blank — no back-fill needed
    }
  }
  _buildSheet(ss, AR.RETURNS,  RETURN_HEADERS,  RETURN_WIDTHS);
  var msg = 'AR ready:\n• '+AR.ORDERS+'\n• '+AR.INVOICES+'\n• '+AR.PAYMENTS+'\n• '+AR.RETURNS;
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, message:msg};
}

// ── Sequential ID generators (lock-guarded by the caller) ───
function _nextSeqId(sheet, prefix){
  var max = 0;
  if (sheet.getLastRow() > 1){
    var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0;i<ids.length;i++){
      var v = _s(ids[i][0]);
      if (v.indexOf(prefix) === 0){ var n = parseInt(v.substring(prefix.length),10); if(!isNaN(n)&&n>max)max=n; }
    }
  }
  return prefix + String(max+1).padStart(4,'0');
}

// Suggest the next real invoice number in a FY series: 3S/<FY>/<nnn>
function suggestNextInvoiceNo(fy){
  fy = _s(fy) || _currentFY();
  var ss = _getSS(), sheet = ss.getSheetByName(AR.INVOICES);
  var max = 0, prefix = '3S/' + fy + '/';
  if (sheet && sheet.getLastRow() > 1){
    var nos = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0;i<nos.length;i++){
      var v = _s(nos[i][0]);
      if (v.indexOf(prefix) === 0){ var n = parseInt(v.substring(prefix.length),10); if(!isNaN(n)&&n>max)max=n; }
    }
  }
  return prefix + String(max+1).padStart(3,'0');
}

// Party GSTIN lookup
function _gstinForParty(customerId){
  try {
    var sheet = _getSS().getSheetByName(M.PARTY);
    if (!sheet || sheet.getLastRow()<2) return '';
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
    for (var i=0;i<rows.length;i++){ if (_s(rows[i][0]) === _s(customerId)) return _s(rows[i][3]); }
  } catch(e){}
  return '';
}

// Party name lookup (so we don't have to ask the frontend to send it)
function _nameForParty(customerId){
  try {
    var sheet = _getSS().getSheetByName(M.PARTY);
    if (!sheet || sheet.getLastRow()<2) return '';
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
    for (var i=0;i<rows.length;i++){ if (_s(rows[i][0]) === _s(customerId)) return _s(rows[i][2]); }
  } catch(e){}
  return '';
}

// ============================================================
// ORDERS  (optional)
//   data = {customerId, customerName, orderAmount, description, deliveryDate}
// ============================================================
// ============================================================
// INVOICES
//   data = {invoiceNo, orderId(optional), customerId, customerName,
//     invoiceDate, dueDate, taxableValue, gstRate, cgst, sgst, igst,
//     invoiceAmount(if no GST), transportName, lrNumber, notes}
//   Total = taxable + CGST + SGST + IGST  (if taxable given), else invoiceAmount.
//   Auto-applies any customer advance (oldest unapplied payment first).
// ============================================================
function createInvoice(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(AR.INVOICES);
    if (!sheet) { setupAR(); sheet = ss.getSheetByName(AR.INVOICES); }

    var invNo = _s(data.invoiceNo);
    if (!invNo) return {success:false, message:'Invoice number required.'};
    if (!_s(data.customerId)) return {success:false, message:'Customer required.'};

    // Unique invoice number
    if (sheet.getLastRow() > 1){
      var nos = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
      for (var i=0;i<nos.length;i++){ if (_s(nos[i][0]) === invNo) return {success:false, message:'Invoice number already exists.'}; }
    }

    var taxable = _n(data.taxable);
    var cgst = _n(data.cgst), sgst = _n(data.sgst), igst = _n(data.igst);
    // Frontend sends `data.amount` (already computed: taxable+gst if GST, else raw amount).
    // Keep backward compat with the older `invoiceAmount` field name.
    var amountIn = _n(data.amount);
    if (!amountIn) amountIn = _n(data.invoiceAmount);

    // ── Discount handling ──
    //   discountType    = 'Percent' | 'Amount' | ''
    //   discountValue   = the raw % or ₹ figure entered by the user
    //   discountApplied = 'Pre-GST' | 'Post-GST' | ''
    //   discountAmount  = the FINAL ₹ discount actually deducted (for reports)
    //
    // Pre-GST: subtract discount from taxable → recompute CGST/SGST/IGST at
    //   the SAME rate → total = discounted_taxable + new_tax. GST liability
    //   goes down (legally cleaner, standard Indian practice).
    // Post-GST: keep GST as-is → total = (taxable + tax) − discount.
    //   The GST you owe government doesn't change (post-GST = "cash discount").
    var discountType    = _s(data.discountType);
    var discountValue   = _n(data.discountValue);
    var discountApplied = _s(data.discountApplied);
    var discountAmount  = 0;
    var total = 0;

    if (discountType && discountValue > 0 && discountApplied){
      if (discountApplied === 'Pre-GST'){
        // Compute discount on taxable
        var discBase = taxable;
        if (discountType === 'Percent') discountAmount = _round(discBase * discountValue / 100);
        else                            discountAmount = _round(discountValue);
        if (discountAmount > discBase) discountAmount = discBase;
        var newTaxable = _round(taxable - discountAmount);
        // Recompute tax at same effective rate (preserve inter/intra split)
        var origTax = cgst + sgst + igst;
        var effectiveRate = taxable > 0 ? (origTax / taxable) : 0;
        var newTax = _round(newTaxable * effectiveRate);
        if (igst > 0){
          // Inter-state — full new tax to IGST
          igst = newTax; cgst = 0; sgst = 0;
        } else {
          // Intra-state — split 50/50
          cgst = _round(newTax / 2); sgst = _round(newTax / 2);
          igst = 0;
        }
        taxable = newTaxable;
        total = _round(newTaxable + cgst + sgst + igst);
      } else if (discountApplied === 'Post-GST'){
        // Compute discount on subtotal (taxable + tax)
        var subtotal = taxable > 0 ? _round(taxable + cgst + sgst + igst) : amountIn;
        if (discountType === 'Percent') discountAmount = _round(subtotal * discountValue / 100);
        else                            discountAmount = _round(discountValue);
        if (discountAmount > subtotal) discountAmount = subtotal;
        total = _round(subtotal - discountAmount);
      }
    } else {
      // No discount — original math
      total = taxable > 0 ? _round(taxable + cgst + sgst + igst) : amountIn;
    }
    if (total <= 0) return {success:false, message:'Invoice amount must be > 0.'};

    var invDate = data.invoiceDate ? _asDate(data.invoiceDate) : new Date();
    var dueDate = data.dueDate ? _asDate(data.dueDate) : invDate;
    var fy = _fyFromDate(invDate);
    var gstin = _gstinForParty(data.customerId);

    var outstanding = total, paid = 0, status = INV_STATUS.UNPAID;

    var invRow = [ invNo, _s(data.orderId), _s(data.customerId), (_s(data.customerName) || _nameForParty(data.customerId)), gstin,
      invDate, dueDate, fy, taxable, _s(data.gstRate), cgst, sgst, igst, total,
      outstanding, paid, _s(data.transport || data.transportName), _s(data.lrNumber), '', status, _s(data.notes),
      Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss'),
      _s(data.project),
      discountType, discountValue || '', discountApplied, discountAmount || '' ];
    // Reliable write with verification (see _writeRowVerified in Bills.gs)
    var invWriteRes = _writeRowVerified(sheet, invRow, IC.NO, invNo);
    if (!invWriteRes.ok) return {success:false, message:invWriteRes.message};
    var invRowNum = sheet.getLastRow();

    // Auto-apply existing customer advance (unapplied payments, oldest first)
    var applied = _applyAdvanceToInvoice(data.customerId, invNo, total);
    if (applied > 0){
      paid = applied; outstanding = _round(total - applied);
      status = outstanding <= 0.01 ? INV_STATUS.PAID : INV_STATUS.PARTIAL;
      sheet.getRange(invRowNum, IC.OUTSTANDING+1, 1, 2).setValues([[outstanding, paid]]);
      sheet.getRange(invRowNum, IC.STATUS+1).setValue(status);
    }

    // Optional first payment ("amount received" on the New Invoice form) —
    // mirrors the createBill "amount already paid" flow. Records a customer
    // payment targeted at this invoice so Outstanding / Status reflect the
    // deposit in one action.
    var receivedNow = _n(data.receivedNow);
    var recvNowInfo = null;
    if (receivedNow > 0){
      if (receivedNow > total){
        return {success:false, message:'"Amount received" (₹' + receivedNow + ') cannot exceed the invoice total (₹' + total + ').'};
      }
      if (outstanding > 0.01){
        recvNowInfo = recordPayment({
          customerId:        _s(data.customerId),
          amount:            receivedNow,
          paymentDate:       data.receivedNowDate || data.invoiceDate,
          mode:              _s(data.receivedNowMode),
          referenceNo:       _s(data.receivedNowRef),
          receivedInAccount: _s(data.receivedNowAcct),
          tdsAmount:         _n(data.receivedNowTds),
          notes:             'First payment against ' + invNo,
          targetInvoice:     invNo
        });
        if (recvNowInfo && !recvNowInfo.success){
          // Invoice IS created — report payment failure so the owner can retry manually.
          return {success:true, invoiceNo:invNo, total:total, paidFromAdvance:applied,
                  paymentError: recvNowInfo.message,
                  message:'Invoice ' + invNo + ' created, but the first payment failed: ' + recvNowInfo.message};
        }
      }
    }

    var msg = 'Invoice ' + invNo + ' created.';
    if (applied > 0) msg += ' Applied ₹' + applied + ' from advance.';
    if (recvNowInfo && recvNowInfo.success) msg += ' Received ₹' + receivedNow + ' now (' + recvNowInfo.paymentId + ').';

    return {success:true, invoiceNo:invNo, total:total, paidFromAdvance:applied,
            firstPaymentId:(recvNowInfo && recvNowInfo.paymentId) || '',
            status:status, message:msg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Consume customer advance (Unapplied on payment rows, oldest first) for a new invoice.
// Returns amount applied.
function _applyAdvanceToInvoice(customerId, invoiceNo, invoiceTotal){
  var sheet = _getSS().getSheetByName(AR.PAYMENTS);
  if (!sheet || sheet.getLastRow()<2) return 0;
  var w = Math.min(sheet.getLastColumn(), PMC_TOTAL);
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
  // oldest payment first
  var idx = [];
  for (var i=0;i<data.length;i++){
    if (_s(data[i][PMC.CUST_ID]) === _s(customerId) && _n(data[i][PMC.UNAPPLIED]) > 0) idx.push(i);
  }
  idx.sort(function(a,b){ return _asDate(data[a][PMC.DATE]) - _asDate(data[b][PMC.DATE]); });

  var need = invoiceTotal, applied = 0;
  for (var k=0; k<idx.length && need > 0.01; k++){
    var r = idx[k], rowNum = r+2;
    var avail = _n(data[r][PMC.UNAPPLIED]);
    var use = Math.min(avail, need);
    var newUnapplied = _round(avail - use);
    var note = _s(data[r][PMC.APPLIED]);
    note += (note ? '; ' : '') + invoiceNo + ':' + _round(use) + '(adv)';
    sheet.getRange(rowNum, PMC.APPLIED+1).setValue(note);
    sheet.getRange(rowNum, PMC.UNAPPLIED+1).setValue(newUnapplied);
    applied = _round(applied + use); need = _round(need - use);
  }
  return applied;
}

// ============================================================
// PAYMENTS  (FIFO apply to oldest open invoices; leftover = advance)
//   data = {customerId, customerName, paymentDate, amount, mode, referenceNo, notes}
// ============================================================
// Returns open invoices for a customer, oldest first. Used by the AR
// payment form's optional "Apply to specific invoice" dropdown.
//   [{no, date, total, outstanding}]
function getCustomerOpenInvoices(customerId){
  try {
    var custId = _s(customerId);
    if (!custId) return [];
    var iSheet = _getSS().getSheetByName(AR.INVOICES);
    if (!iSheet || iSheet.getLastRow() < 2) return [];
    var iW = Math.min(iSheet.getLastColumn(), IC_TOTAL);
    var rows = iSheet.getRange(2,1,iSheet.getLastRow()-1, iW).getValues();
    var out = [];
    for (var i=0;i<rows.length;i++){
      var r = rows[i];
      if (_s(r[IC.CUST_ID]) !== custId) continue;
      var status = _s(r[IC.STATUS]);
      if (status === INV_STATUS.PAID || status === INV_STATUS.CANCELLED) continue;
      var outstanding = _n(r[IC.OUTSTANDING]);
      if (outstanding <= 0) continue;
      out.push({
        no: _s(r[IC.NO]),
        date: Utilities.formatDate(_asDate(r[IC.INV_DATE]),'GMT+5:30','yyyy-MM-dd'),
        total: _n(r[IC.TOTAL]),
        outstanding: outstanding,
        status: status,
        project: iW > IC.PROJECT ? _s(r[IC.PROJECT]) : ''
      });
    }
    out.sort(function(a,b){ return a.date.localeCompare(b.date); });
    return out;
  } catch(e){ Logger.log('getCustomerOpenInvoices: '+e.message); return []; }
}

function recordPayment(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var ss = _getSS();
    var pSheet = ss.getSheetByName(AR.PAYMENTS), iSheet = ss.getSheetByName(AR.INVOICES);
    if (!pSheet || !iSheet) { setupAR(); pSheet = ss.getSheetByName(AR.PAYMENTS); iSheet = ss.getSheetByName(AR.INVOICES); }

    var custId = _s(data.customerId);
    if (!custId) return {success:false, message:'Customer required.'};
    var amount = _n(data.amount);
    if (amount <= 0) return {success:false, message:'Payment amount must be > 0.'};
    var payDate = data.paymentDate ? _asDate(data.paymentDate) : new Date();

    // Open invoices for this customer, oldest invoice date first (FIFO)
    var open = [];
    if (iSheet.getLastRow() > 1){
      var inv = iSheet.getRange(2,1,iSheet.getLastRow()-1, Math.min(iSheet.getLastColumn(), IC_TOTAL)).getValues();
      for (var i=0;i<inv.length;i++){
        if (_s(inv[i][IC.CUST_ID]) === custId &&
            _s(inv[i][IC.STATUS]) !== INV_STATUS.PAID &&
            _s(inv[i][IC.STATUS]) !== INV_STATUS.CANCELLED &&
            _n(inv[i][IC.OUTSTANDING]) > 0){
          open.push({rowNum:i+2, no:_s(inv[i][IC.NO]), outstanding:_n(inv[i][IC.OUTSTANDING]),
                     paid:_n(inv[i][IC.PAID]), date:_asDate(inv[i][IC.INV_DATE])});
        }
      }
    }
    open.sort(function(a,b){ return a.date - b.date; });

    // Per-payment override: when the caller passes `targetInvoice` (an
    // invoice number like "3S/2526/045"), reorder the queue so that
    // invoice gets paid first. Any leftover after settling that invoice
    // falls back to FIFO on the remaining open invoices.
    var targetInvoice = _s(data.targetInvoice);
    if (targetInvoice){
      var targetIdx = -1;
      for (var t=0; t<open.length; t++){ if (open[t].no === targetInvoice){ targetIdx = t; break; } }
      if (targetIdx > 0){
        var picked = open.splice(targetIdx, 1)[0];
        open.unshift(picked);
      } else if (targetIdx === -1){
        return {success:false, message:'Invoice ' + targetInvoice + ' not found among this customer\'s open invoices.'};
      }
    }

    var remaining = amount, appliedNotes = [];
    for (var k=0; k<open.length && remaining > 0.01; k++){
      var o = open[k];
      var pay = Math.min(o.outstanding, remaining);
      var newOut = _round(o.outstanding - pay), newPaid = _round(o.paid + pay);
      var newStatus = newOut <= 0.01 ? INV_STATUS.PAID : INV_STATUS.PARTIAL;
      iSheet.getRange(o.rowNum, IC.OUTSTANDING+1, 1, 2).setValues([[newOut, newPaid]]);
      iSheet.getRange(o.rowNum, IC.STATUS+1).setValue(newStatus);
      appliedNotes.push(o.no + ':' + _round(pay));
      remaining = _round(remaining - pay);
    }

    var unapplied = _round(remaining);   // becomes advance/credit
    var payId = _nextSeqId(pSheet, 'PAY-');

    // TDS handling: `amount` (passed in by the caller) is the GROSS amount
    // settled against invoices. `tdsAmount` is the deduction made by the
    // customer. `bankAmount` is what actually entered the bank.
    var tdsAmount = _n(data.tdsAmount);
    if (tdsAmount < 0) tdsAmount = 0;
    if (tdsAmount > amount) return {success:false, message:'TDS amount cannot exceed payment amount.'};
    var bankAmount = _round(amount - tdsAmount);

    // If frontend didn't send the display name, look it up from M_Party.
    // Same fallback pattern used by createInvoice above.
    var custName = _s(data.customerName) || _nameForParty(custId);

    var payRow = [ payId, custId, custName, payDate, _fyFromDate(payDate), amount,
      _s(data.mode), _s(data.referenceNo), appliedNotes.join('; ') || (unapplied>0?'(advance)':''),
      unapplied, _s(data.notes), Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss'),
      tdsAmount, bankAmount, _s(data.receivedInAccount) ];
    var payWriteRes = _writeRowVerified(pSheet, payRow, PMC.ID, payId);
    if (!payWriteRes.ok) return {success:false, message:payWriteRes.message};

    // Project on payment row: derived ONLY from the single applied invoice.
    // Multi-invoice splits and advances leave it blank — per-project attribution
    // for those comes from the Applied_Invoices string, not this field.
    // The form no longer exposes a Project dropdown; ignoring data.project
    // closes the double-count loophole where operator's tag differed from
    // the applied invoice's project.
    var project = '';
    if (appliedNotes.length === 1){
      var onlyInvNo = appliedNotes[0].split(':')[0];
      project = _getInvoiceProject(onlyInvNo);
    }
    if (project && pSheet.getLastColumn() > PMC.PROJECT){
      pSheet.getRange(pSheet.getLastRow(), PMC.PROJECT + 1).setValue(project);
    }

    return {success:true, paymentId:payId, applied:appliedNotes, advance:unapplied,
            tdsAmount:tdsAmount, bankAmount:bankAmount, project:project};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// RETURNS  (credit against an invoice — reduces its outstanding)
//   data = {invoiceNo, returnDate, returnAmount, reason, notes}
// ============================================================
function recordReturn(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var ss = _getSS();
    var rSheet = ss.getSheetByName(AR.RETURNS), iSheet = ss.getSheetByName(AR.INVOICES);
    if (!rSheet || !iSheet) { setupAR(); rSheet = ss.getSheetByName(AR.RETURNS); iSheet = ss.getSheetByName(AR.INVOICES); }

    var invNo = _s(data.invoiceNo);
    var amt = _n(data.returnAmount);
    if (!invNo) return {success:false, message:'Invoice number required.'};
    if (amt <= 0) return {success:false, message:'Return amount must be > 0.'};

    // Find invoice
    var inv = iSheet.getRange(2,1,iSheet.getLastRow()-1, Math.min(iSheet.getLastColumn(), IC_TOTAL)).getValues();
    var found = -1;
    for (var i=0;i<inv.length;i++){ if (_s(inv[i][IC.NO]) === invNo){ found = i; break; } }
    if (found === -1) return {success:false, message:'Invoice not found.'};

    var rowNum = found+2;
    var custId = _s(inv[found][IC.CUST_ID]), custName = _s(inv[found][IC.CUST_NAME]);
    var outstanding = _n(inv[found][IC.OUTSTANDING]);
    if (amt > _n(inv[found][IC.AMOUNT])) return {success:false, message:'Return exceeds invoice amount.'};

    var newOut = _round(Math.max(0, outstanding - amt));
    iSheet.getRange(rowNum, IC.OUTSTANDING+1).setValue(newOut);
    if (newOut <= 0.01) iSheet.getRange(rowNum, IC.STATUS+1).setValue(INV_STATUS.PAID);

    var retDate = data.returnDate ? _asDate(data.returnDate) : new Date();
    var retId = _nextSeqId(rSheet, 'RET-');
    rSheet.appendRow([ retId, invNo, custId, custName, retDate, _fyFromDate(retDate), amt,
      _s(data.reason), _s(data.notes), Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss') ]);
    SpreadsheetApp.flush();

    return {success:true, returnId:retId};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// COMPUTED CUSTOMER LEDGER  (date-sorted, never stored)
//   Invoice → debit (owes us). Payment → credit. Return → credit.
//   Running balance computed in date order. Balance>0 = receivable.
// ============================================================
function getCustomerLedger(customerId){
  try {
    var ss = _getSS();
    var entries = [];

    var iSheet = ss.getSheetByName(AR.INVOICES);
    if (iSheet && iSheet.getLastRow()>1){
      var iW = Math.min(iSheet.getLastColumn(), IC_TOTAL);
      iSheet.getRange(2,1,iSheet.getLastRow()-1,iW).getValues().forEach(function(r){
        if (_s(r[IC.CUST_ID]) === _s(customerId) && _s(r[IC.STATUS]) !== INV_STATUS.CANCELLED){
          // Note any discount applied — used by the ledger UI as a sub-line.
          // The invoice AMOUNT already reflects the discounted total, so we
          // do NOT push a separate ledger entry (avoids double-counting);
          // the UI shows the discount as context under the invoice row.
          var discNote = '';
          if (iW > IC.DISC_AMOUNT){
            var dAmt = _n(r[IC.DISC_AMOUNT]);
            if (dAmt > 0){
              var dType = _s(r[IC.DISC_TYPE]);
              var dVal  = _n(r[IC.DISC_VALUE]);
              var dApp  = _s(r[IC.DISC_APPLIED]);
              var label = (dType === 'Percent') ? (dVal + '%') : ('₹' + dVal);
              discNote = 'Discount: ' + label + ' ' + dApp + ' (₹' + _round(dAmt) + ')';
            }
          }
          entries.push({date:_asDate(r[IC.INV_DATE]), type:'Invoice', ref:_s(r[IC.NO]),
                        debit:_n(r[IC.AMOUNT]), credit:0,
                        project: iW > IC.PROJECT ? _s(r[IC.PROJECT]) : '',
                        applied: '',
                        discountNote: discNote});
        }
      });
    }
    var pSheet = ss.getSheetByName(AR.PAYMENTS);
    if (pSheet && pSheet.getLastRow()>1){
      var pW = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
      pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues().forEach(function(r){
        if (_s(r[PMC.CUST_ID]) !== _s(customerId)) return;
        // Skip soft-cancelled payments once the audit column exists (Tier 2.1).
        if (typeof PMC.CANCELLED_ON !== 'undefined' && pW > PMC.CANCELLED_ON
            && _s(r[PMC.CANCELLED_ON])) return;
        entries.push({date:_asDate(r[PMC.DATE]), type:'Payment', ref:_s(r[PMC.ID]),
                      debit:0, credit:_n(r[PMC.AMOUNT]),
                      project: pW > PMC.PROJECT ? _s(r[PMC.PROJECT]) : '',
                      applied: _s(r[PMC.APPLIED]),
                      unapplied: pW > PMC.UNAPPLIED ? _n(r[PMC.UNAPPLIED]) : 0});
      });
    }
    var rSheet = ss.getSheetByName(AR.RETURNS);
    if (rSheet && rSheet.getLastRow()>1){
      rSheet.getRange(2,1,rSheet.getLastRow()-1,RC_TOTAL).getValues().forEach(function(r){
        if (_s(r[RC.CUST_ID]) === _s(customerId))
          entries.push({date:_asDate(r[RC.DATE]), type:'Return', ref:_s(r[RC.ID]),
                        debit:0, credit:_n(r[RC.AMOUNT])});
      });
    }

    entries.sort(function(a,b){ return a.date - b.date; });
    var bal = 0;
    entries.forEach(function(e){
      bal = _round(bal + e.debit - e.credit);
      e.balance = bal;
      // Serialize the date to a string BEFORE returning — GAS serialisation
      // of native Date objects across the wire has been fragile in this app,
      // and a single non-serializable field turns the whole response into null.
      try { e.date = Utilities.formatDate(e.date, 'GMT+5:30', 'yyyy-MM-dd'); }
      catch(err){ e.date = ''; }
    });

    var payload = {
      ok: true,
      entries: entries,
      balance: bal,
      totalDebit:  _round(entries.reduce(function(s,e){return s+e.debit;},0)),
      totalCredit: _round(entries.reduce(function(s,e){return s+e.credit;},0)),
      error: null
    };
    // Defensive JSON wrap: any leftover undefined/NaN gets nulled out.
    return JSON.parse(JSON.stringify(payload, function(k,v){
      if (v === undefined) return null;
      if (typeof v === 'number' && !isFinite(v)) return null;
      return v;
    }));
  } catch(e){
    Logger.log('getCustomerLedger: ' + e.message);
    return {ok:false, entries:[], balance:0, totalDebit:0, totalCredit:0,
            error: String(e && e.message || e)};
  }
}

// ── Read helpers (filters: {fy, status, customerId}) ────────
function getInvoices(filters){
  filters = filters || {};
  var sheet = _getSS().getSheetByName(AR.INVOICES);
  if (!sheet || sheet.getLastRow()<2) return [];
  var w = Math.min(sheet.getLastColumn(), IC_TOTAL);
  return sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues()
    .map(function(r){ return {
      invoiceNo:_s(r[IC.NO]), orderId:_s(r[IC.ORDER_ID]), customerId:_s(r[IC.CUST_ID]),
      customerName:_s(r[IC.CUST_NAME]), gstin:_s(r[IC.GSTIN]), invoiceDate:r[IC.INV_DATE],
      dueDate:r[IC.DUE_DATE], fy:_s(r[IC.FY]), taxable:_n(r[IC.TAXABLE]), gstRate:_s(r[IC.GST_RATE]),
      cgst:_n(r[IC.CGST]), sgst:_n(r[IC.SGST]), igst:_n(r[IC.IGST]), amount:_n(r[IC.AMOUNT]),
      outstanding:_n(r[IC.OUTSTANDING]), paid:_n(r[IC.PAID]), status:_s(r[IC.STATUS]),
      discountType:   w > IC.DISC_TYPE    ? _s(r[IC.DISC_TYPE])    : '',
      discountValue:  w > IC.DISC_VALUE   ? _n(r[IC.DISC_VALUE])   : 0,
      discountApplied:w > IC.DISC_APPLIED ? _s(r[IC.DISC_APPLIED]) : '',
      discountAmount: w > IC.DISC_AMOUNT  ? _n(r[IC.DISC_AMOUNT])  : 0
    }; })
    .filter(function(o){
      if (!o.invoiceNo) return false;
      if (filters.status && o.status !== _s(filters.status)) return false;
      if (filters.customerId && o.customerId !== _s(filters.customerId)) return false;
      if (filters.fy && filters.fy!=='all' && o.fy !== (filters.fy==='current'?_currentFY():_s(filters.fy))) return false;
      return true;
    });
}
function _round(n){ try { return parseFloat((parseFloat(n)||0).toFixed(2)); } catch(e){ return 0; } }
// ============================================================
// PAYMENT EDIT / DELETE / LIST — customer side (Tier 2.2)
// ============================================================

function _getInvoiceProject(invoiceNo){
  try {
    var sheet = _getSS().getSheetByName(AR.INVOICES);
    if (!sheet || sheet.getLastRow() < 2) return '';
    var w = Math.min(sheet.getLastColumn(), IC_TOTAL);
    var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][IC.NO]) === _s(invoiceNo)){
        return w > IC.PROJECT ? _s(rows[i][IC.PROJECT]) : '';
      }
    }
    return '';
  } catch(e){ return ''; }
}

function updateCustomerPayment(payId, patches){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    payId = _s(payId);
    if (!payId) return {success:false, message:'Payment ID required.'};
    patches = patches || {};
    var sheet = _getSS().getSheetByName(AR.PAYMENTS);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'Payments sheet empty.'};
    var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    var rowNum = -1;
    for (var i=0;i<ids.length;i++){ if (_s(ids[i][0]) === payId){ rowNum = i+2; break; } }
    if (rowNum < 0) return {success:false, message:'Payment '+payId+' not found.'};

    if (sheet.getLastColumn() > PMC.CANCELLED_ON){
      if (_s(sheet.getRange(rowNum, PMC.CANCELLED_ON+1).getValue())){
        return {success:false, message:'Cannot edit a cancelled payment.'};
      }
    }

    if (patches.hasOwnProperty('paymentDate') && patches.paymentDate)
      sheet.getRange(rowNum, PMC.DATE+1).setValue(_asDate(patches.paymentDate));
    if (patches.hasOwnProperty('mode'))
      sheet.getRange(rowNum, PMC.MODE+1).setValue(_s(patches.mode));
    if (patches.hasOwnProperty('referenceNo'))
      sheet.getRange(rowNum, PMC.REF+1).setValue(_s(patches.referenceNo));
    if (patches.hasOwnProperty('notes'))
      sheet.getRange(rowNum, PMC.NOTES+1).setValue(_s(patches.notes));
    if (patches.hasOwnProperty('project') && sheet.getLastColumn() > PMC.PROJECT)
      sheet.getRange(rowNum, PMC.PROJECT+1).setValue(_s(patches.project));
    if (patches.hasOwnProperty('receivedInAccount'))
      sheet.getRange(rowNum, PMC.RCV_ACCT+1).setValue(_s(patches.receivedInAccount));
    if (patches.hasOwnProperty('tdsAmount')){
      var newTds = _n(patches.tdsAmount);
      if (newTds < 0) newTds = 0;
      var amount = _n(sheet.getRange(rowNum, PMC.AMOUNT+1).getValue());
      if (newTds > amount) return {success:false, message:'TDS cannot exceed payment amount.'};
      sheet.getRange(rowNum, PMC.TDS_AMT+1).setValue(newTds);
      sheet.getRange(rowNum, PMC.BANK_AMT+1).setValue(_round(amount - newTds));
    }

    if (sheet.getLastColumn() > PMC.LAST_MODIFIED_ON){
      sheet.getRange(rowNum, PMC.LAST_MODIFIED_ON+1)
        .setValue(Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss'));
      sheet.getRange(rowNum, PMC.LAST_MODIFIED_BY+1).setValue(_currentUser());
    }
    return {success:true, paymentId:payId, message:'Payment updated.'};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

function deleteCustomerPayment(payId, reason){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    payId = _s(payId);
    if (!payId) return {success:false, message:'Payment ID required.'};
    var ss = _getSS();
    var sheet = ss.getSheetByName(AR.PAYMENTS);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'Payments sheet empty.'};

    var w = Math.min(sheet.getLastColumn(), PMC_TOTAL);
    var all = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var rowNum = -1, row = null;
    for (var i=0;i<all.length;i++){ if (_s(all[i][PMC.ID]) === payId){ rowNum = i+2; row = all[i]; break; } }
    if (rowNum < 0) return {success:false, message:'Payment '+payId+' not found.'};
    if (w > PMC.CANCELLED_ON && _s(row[PMC.CANCELLED_ON])){
      return {success:false, message:'Payment already cancelled.'};
    }

    // Reverse invoice application. Applied_Invoices format: "3S/2526/045:5000; 3S/2526/046:3000"
    var applied = _parseAppliedString(row[PMC.APPLIED]);
    var iSheet = ss.getSheetByName(AR.INVOICES);
    var restored = [];
    if (applied.length && iSheet && iSheet.getLastRow() > 1){
      var iw = Math.min(iSheet.getLastColumn(), IC_TOTAL);
      var irows = iSheet.getRange(2, 1, iSheet.getLastRow()-1, iw).getValues();
      applied.forEach(function(chunk){
        for (var j=0;j<irows.length;j++){
          if (_s(irows[j][IC.NO]) === chunk.ref){
            var irn = j + 2;
            var curPaid = _n(irows[j][IC.PAID]);
            var curOut  = _n(irows[j][IC.OUTSTANDING]);
            var newPaid = _round(curPaid - chunk.amount);
            var newOut  = _round(curOut  + chunk.amount);
            if (newPaid < 0) newPaid = 0;
            var newStatus = (newPaid <= 0.001) ? INV_STATUS.UNPAID
                          : (newOut  <= 0.001) ? INV_STATUS.PAID
                          :                       INV_STATUS.PARTIAL;
            iSheet.getRange(irn, IC.OUTSTANDING+1, 1, 2).setValues([[newOut, newPaid]]);
            iSheet.getRange(irn, IC.STATUS+1).setValue(newStatus);
            restored.push(chunk.ref + ' (+₹' + chunk.amount + ')');
            break;
          }
        }
      });
    }

    // Zero out any unapplied advance so it can't be consumed later
    if (w > PMC.UNAPPLIED){
      sheet.getRange(rowNum, PMC.UNAPPLIED+1).setValue(0);
    }

    var stamp = Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(rowNum, PMC.CANCELLED_ON+1).setValue(stamp);
    sheet.getRange(rowNum, PMC.CANCELLED_BY+1).setValue(_currentUser());
    sheet.getRange(rowNum, PMC.CANCELLATION_REASON+1).setValue(_s(reason));

    return {
      success: true,
      paymentId: payId,
      reversed: restored.length,
      message: 'Payment ' + payId + ' cancelled' +
               (restored.length ? '; restored ' + restored.join(', ') : '')
    };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

function listCustomerPayments(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(AR.PAYMENTS);
    if (!sheet || sheet.getLastRow() < 2) return _envOk({rows:[]});
    var w = Math.min(sheet.getLastColumn(), PMC_TOTAL);
    var vals = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var wantCust = _s(filters.customerId);
    var wantFy   = _s(filters.fy);
    var q        = _s(filters.search).toLowerCase();
    var incCancel= !!filters.includeCancelled;
    var limit    = _n(filters.limit) || 300;
    var out = [];
    for (var i=0;i<vals.length;i++){
      var r = vals[i];
      var id = _s(r[PMC.ID]); if (!id) continue;
      var cancelledOn = w > PMC.CANCELLED_ON ? _s(r[PMC.CANCELLED_ON]) : '';
      if (cancelledOn && !incCancel) continue;
      if (wantCust && _s(r[PMC.CUST_ID]) !== wantCust) continue;
      if (wantFy && _s(r[PMC.FY]) !== wantFy) continue;
      var blob = (id + ' ' + _s(r[PMC.CUST_NAME]) + ' ' + _s(r[PMC.REF]) + ' ' +
                  _s(r[PMC.APPLIED]) + ' ' + _s(r[PMC.NOTES])).toLowerCase();
      if (q && blob.indexOf(q) === -1) continue;
      out.push({
        id: id,
        customerId: _s(r[PMC.CUST_ID]),
        customer: _s(r[PMC.CUST_NAME]),
        date: _safeDateStr(r[PMC.DATE]),
        fy: _s(r[PMC.FY]),
        amount: _n(r[PMC.AMOUNT]),
        mode: _s(r[PMC.MODE]),
        reference: _s(r[PMC.REF]),
        applied: _s(r[PMC.APPLIED]),
        unapplied: _n(r[PMC.UNAPPLIED]),
        tdsAmount: _n(r[PMC.TDS_AMT]),
        bankAmount: _n(r[PMC.BANK_AMT]),
        receivedInAccount: w > PMC.RCV_ACCT ? _s(r[PMC.RCV_ACCT]) : '',
        notes: _s(r[PMC.NOTES]),
        cancelled: !!cancelledOn,
        cancelledOn: cancelledOn,
        cancelledBy: w > PMC.CANCELLED_BY ? _s(r[PMC.CANCELLED_BY]) : '',
        cancellationReason: w > PMC.CANCELLATION_REASON ? _s(r[PMC.CANCELLATION_REASON]) : '',
        lastModifiedOn: w > PMC.LAST_MODIFIED_ON ? _s(r[PMC.LAST_MODIFIED_ON]) : '',
        lastModifiedBy: w > PMC.LAST_MODIFIED_BY ? _s(r[PMC.LAST_MODIFIED_BY]) : '',
        project:        w > PMC.PROJECT ? _s(r[PMC.PROJECT]) : ''
      });
    }
    out.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
    return _envOk({ rows: out.slice(0, limit), count: out.length });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// HARD DELETE — invoice + related customer payment allocations
//   Removes the Invoices row entirely, reverses any payment
//   allocation that pointed at this invoice (adds back to their
//   Unapplied balance so the money survives as a credit).
//   Called by hardDeleteInvoice below and by cascade delete of a
//   customer via hardDeletePartyCascade.
// ============================================================
function _reverseAppliedInvoicesForInvoice(invoiceNo){
  var restored = [];
  var ss = _getSS();
  var pSheet = ss.getSheetByName(AR.PAYMENTS);
  if (!pSheet || pSheet.getLastRow() < 2) return restored;
  var w = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
  var rows = pSheet.getRange(2, 1, pSheet.getLastRow()-1, w).getValues();
  for (var i=0;i<rows.length;i++){
    var applied = _s(rows[i][PMC.APPLIED]);
    if (!applied) continue;
    var chunks = _parseAppliedString(applied);
    var kept = [], releasedAmt = 0;
    chunks.forEach(function(c){
      if (c.ref === invoiceNo){ releasedAmt += c.amount; }
      else { kept.push(c.ref + ':' + c.amount); }
    });
    if (releasedAmt <= 0) continue;
    var curUnapplied = _n(rows[i][PMC.UNAPPLIED]);
    var newUnapplied = _round(curUnapplied + releasedAmt);
    pSheet.getRange(i+2, PMC.APPLIED+1).setValue(kept.join('; '));
    pSheet.getRange(i+2, PMC.UNAPPLIED+1).setValue(newUnapplied);
    restored.push(_s(rows[i][PMC.ID]) + ' (+₹' + releasedAmt + ' returned to advance)');
  }
  return restored;
}

// hardDeleteInvoice(invoiceNo, reason) — removes the invoice row entirely.
// Reverses any customer payment allocations that pointed at it.
function hardDeleteInvoice(invoiceNo, reason){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    invoiceNo = _s(invoiceNo);
    if (!invoiceNo) return {success:false, message:'Invoice number required.'};
    var iSheet = _getSS().getSheetByName(AR.INVOICES);
    if (!iSheet || iSheet.getLastRow() < 2) return {success:false, message:'Invoices sheet empty.'};
    var w = Math.min(iSheet.getLastColumn(), IC_TOTAL);
    var rows = iSheet.getRange(2, 1, iSheet.getLastRow()-1, w).getValues();
    var rowIdx = -1;
    for (var i=0;i<rows.length;i++){ if (_s(rows[i][IC.NO]) === invoiceNo){ rowIdx = i; break; } }
    if (rowIdx < 0) return {success:false, message:'Invoice ' + invoiceNo + ' not found.'};
    // First reverse any payment allocations pointing at this invoice.
    var restored = _reverseAppliedInvoicesForInvoice(invoiceNo);
    // Then delete the row.
    iSheet.deleteRow(rowIdx + 2);
    return {
      success: true,
      invoiceNo: invoiceNo,
      paymentsReversed: restored.length,
      message: 'Invoice ' + invoiceNo + ' deleted' + (restored.length ? '; ' + restored.length + ' payment allocation(s) reversed' : '') + (reason ? ' (reason: ' + reason + ')' : '') + '.'
    };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// listInvoices({customerId?, fy?, status?, search?}) — mirror of listBills.
function listInvoices(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(AR.INVOICES);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var w = Math.min(sheet.getLastColumn(), IC_TOTAL);
    var all = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    var today = new Date();
    var fy = _s(filters.fy);
    if (fy === 'current') fy = _currentFY();
    if (fy === 'all') fy = '';
    var cid = _s(filters.customerId);
    var st  = _s(filters.status);
    var q   = _s(filters.search).toLowerCase();
    var out = [];
    for (var i=0;i<all.length;i++){
      var r = all[i];
      if (!_s(r[IC.NO])) continue;
      if (fy && _s(r[IC.FY]) !== fy) continue;
      if (cid && _s(r[IC.CUST_ID]) !== cid) continue;
      if (st && _s(r[IC.STATUS]) !== st) continue;
      if (q){
        var hay = (_s(r[IC.CUST_NAME]) + ' ' + _s(r[IC.NO]) + ' ' + _s(r[IC.NOTES])).toLowerCase();
        if (hay.indexOf(q) === -1) continue;
      }
      var dueDate = _asDate(r[IC.DUE_DATE]);
      var outstanding = _n(r[IC.OUTSTANDING]);
      var daysOverdue = 0;
      if (outstanding > 0 && dueDate < today){
        daysOverdue = Math.floor((today.getTime() - dueDate.getTime()) / 86400000);
      }
      out.push({
        no:           _s(r[IC.NO]),
        customerId:   _s(r[IC.CUST_ID]),
        customer:     _s(r[IC.CUST_NAME]),
        invDate:      _safeDateStr(r[IC.INV_DATE]),
        dueDate:      _safeDateStr(r[IC.DUE_DATE]),
        fy:           _s(r[IC.FY]),
        amount:       _n(r[IC.AMOUNT]),
        outstanding:  outstanding,
        paid:         _n(r[IC.PAID]),
        status:       _s(r[IC.STATUS]),
        project:      w > IC.PROJECT ? _s(r[IC.PROJECT]) : '',
        notes:        _s(r[IC.NOTES]),
        daysOverdue:  daysOverdue
      });
    }
    return out.sort(function(a,b){ return (b.invDate||'').localeCompare(a.invDate||''); });
  } catch(e){ Logger.log('listInvoices: '+e.message); return []; }
}