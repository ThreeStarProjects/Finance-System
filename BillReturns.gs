// ============================================================
// 3 Star Projects — Unified Finance System
// MODULE: Bill Returns / Credit Notes  (AP mirror of AR Returns)
// File: BillReturns.gs   (depends on Bills.gs, Masters.gs, Transactions.gs)
//
// Vendor-side credit note: when we return material to a vendor and the
// vendor issues a credit note instead of cancelling the bill, record it
// here. The credit note REDUCES the bill's Outstanding directly — same
// behaviour as AR Return reducing invoice Outstanding (Invoice.gs:recordReturn).
//
// Sheet: Bill_Returns (created lazily on first use — no setup step needed).
//
// Hooks into getVendorLedger via getBillReturnLedgerEntries() so the
// credit note shows up as a debit (reduces vendor balance). If the bill
// was already fully paid, Outstanding stays 0 but the credit note still
// appears in the ledger as a vendor advance (negative vendor balance),
// matching how AR Return behaves against a paid invoice.
// ============================================================

var BR = { SHEET: 'Bill_Returns' };
var BRC = { ID:0, BILL_ID:1, VEND_ID:2, VEND_NAME:3, RET_DATE:4, FY:5, AMOUNT:6, REASON:7, NOTES:8, CREATED:9 };
var BRC_TOTAL = 10;
var BR_HEADERS = ['Return_ID','Bill_ID','Vendor_ID','Vendor_Name','Return_Date','FY',
                  'Return_Amount','Reason','Notes','Created_Date'];
var BR_WIDTHS  = [120, 110, 110, 230, 120, 80, 130, 260, 240, 160];

function _getBillReturnsSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(BR.SHEET);
  if (sheet) return sheet;
  return _buildSheet(ss, BR.SHEET, BR_HEADERS, BR_WIDTHS);
}

// Resolve a user-entered bill reference (our Bill_ID OR vendor's Bill_No)
// to its row in Bills. Returns {row, rowNum, sheet} or null.
function _findBillByIdOrVendorNo(ref){
  var r = _s(ref); if (!r) return null;
  var ss = _getSS(), b = ss.getSheetByName(AP.BILLS);
  if (!b || b.getLastRow() < 2) return null;
  var rows = b.getRange(2,1,b.getLastRow()-1,BC_TOTAL).getValues();
  for (var i=0;i<rows.length;i++){
    if (_s(rows[i][BC.ID]) === r || _s(rows[i][BC.BILL_NO]) === r){
      return {row: rows[i], rowNum: i+2, sheet: b};
    }
  }
  return null;
}

// ============================================================
// recordBillReturn — credit note against a vendor bill
//   data = {billRef, returnDate, returnAmount, reason, notes}
//     billRef = our Bill_ID (BILL-####) OR vendor's bill number
// ============================================================
function recordBillReturn(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    data = data || {};
    var amt = _n(data.returnAmount);
    if (amt <= 0) return {success:false, message:'Return amount must be > 0.'};

    var hit = _findBillByIdOrVendorNo(data.billRef);
    if (!hit) return {success:false, message:'Bill not found (tried Bill_ID and vendor bill number).'};
    var row = hit.row;
    if (_s(row[BC.STATUS]) === BILL_STATUS.CANCELLED)
      return {success:false, message:'Bill is cancelled — credit note not applicable.'};

    var billAmount = _n(row[BC.AMOUNT]);
    if (amt > billAmount) return {success:false, message:'Return exceeds bill amount.'};

    var outstanding = _n(row[BC.OUTSTANDING]);
    var newOut = _round(Math.max(0, outstanding - amt));
    hit.sheet.getRange(hit.rowNum, BC.OUTSTANDING+1).setValue(newOut);
    if (newOut <= 0.01) hit.sheet.getRange(hit.rowNum, BC.STATUS+1).setValue(BILL_STATUS.PAID);

    var retDate = data.returnDate ? _asDate(data.returnDate) : new Date();
    var rSheet = _getBillReturnsSheet();
    var retId = _nextSeqId(rSheet, 'BRET-');   // reuses helper from Invoice.gs
    rSheet.appendRow([
      retId, _s(row[BC.ID]), _s(row[BC.VEND_ID]), _s(row[BC.VEND_NAME]),
      retDate, _fyFromDate(retDate), amt,
      _s(data.reason), _s(data.notes),
      Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss')
    ]);
    SpreadsheetApp.flush();
    return {success:true, returnId:retId, billId:_s(row[BC.ID]), newOutstanding:newOut};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Ledger entries for getVendorLedger. Credit notes are DEBITS (reduce
// vendor's claim on us), same sign side as a payment.
function getBillReturnLedgerEntries(vendorId){
  try {
    var vid = _s(vendorId); if (!vid) return [];
    var ss = _getSS(), sheet = ss.getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return [];
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,BRC_TOTAL).getValues();
    var out = [];
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][BRC.VEND_ID]) !== vid) continue;
      out.push({
        date: _safeDateStr(rows[i][BRC.RET_DATE]),
        type: 'Credit Note',
        ref: _s(rows[i][BRC.ID]),
        vendorRef: _s(rows[i][BRC.BILL_ID]),
        credit: 0,
        debit: _n(rows[i][BRC.AMOUNT]),
        project: '',
        applied: 'against ' + _s(rows[i][BRC.BILL_ID]),
        unapplied: 0,
        sortKey: _asDate(rows[i][BRC.RET_DATE]).getTime()
      });
    }
    return out;
  } catch(e){ Logger.log('getBillReturnLedgerEntries: '+e.message); return []; }
}

// ============================================================
// listBillReturns — paged-style read for the "Credit Notes" list (future).
//   filters = {vendorId, fy}
// ============================================================
function listBillReturns(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return {ok:true, rows:[], error:null};
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,BRC_TOTAL).getValues();
    var vid = _s(filters.vendorId), fy = _s(filters.fy);
    var out = rows.filter(function(r){
      if (vid && _s(r[BRC.VEND_ID]) !== vid) return false;
      if (fy && fy !== 'all' && _s(r[BRC.FY]) !== fy) return false;
      return !!_s(r[BRC.ID]);
    }).map(function(r){
      return {
        id: _s(r[BRC.ID]),
        billId: _s(r[BRC.BILL_ID]),
        vendorId: _s(r[BRC.VEND_ID]),
        vendorName: _s(r[BRC.VEND_NAME]),
        date: _safeDateStr(r[BRC.RET_DATE]),
        fy: _s(r[BRC.FY]),
        amount: _n(r[BRC.AMOUNT]),
        reason: _s(r[BRC.REASON]),
        notes: _s(r[BRC.NOTES])
      };
    });
    out.sort(function(a,b){ return (b.date || '').localeCompare(a.date || ''); });
    return {ok:true, rows:out, error:null};
  } catch(e){ return {ok:false, rows:[], error:String(e && e.message || e)}; }
}

// ============================================================
// Smoke check — run from the Apps Script editor if credit notes
// start misbehaving. Verifies schema wiring and ledger shape.
// Writes nothing to the sheets.
// ============================================================
function _billReturnModuleCheck(){
  var msgs = [];
  // 1. Headers / column map agree.
  if (BR_HEADERS.length !== BRC_TOTAL) msgs.push('BR_HEADERS.length ('+BR_HEADERS.length+') ≠ BRC_TOTAL ('+BRC_TOTAL+')');
  if (BR_WIDTHS.length  !== BRC_TOTAL) msgs.push('BR_WIDTHS.length ('+BR_WIDTHS.length+') ≠ BRC_TOTAL ('+BRC_TOTAL+')');
  // 2. Lookup returns null for bogus ref — proves the function at least loads
  //    and the Bills sheet is reachable.
  if (_findBillByIdOrVendorNo('__nope__') !== null) msgs.push('_findBillByIdOrVendorNo should return null for bogus ref');
  // 3. Ledger reader returns [] for bogus vendor.
  var le = getBillReturnLedgerEntries('__nope__');
  if (!Array.isArray(le) || le.length !== 0) msgs.push('getBillReturnLedgerEntries should return [] for bogus vendor');
  // 4. The ledger-injection hook in Bills.gs:getVendorLedger must be able to
  //    see this file's symbol. If getVendorLedger is itself missing we're not
  //    in the Finance project — just log.
  if (typeof getVendorLedger !== 'function') msgs.push('getVendorLedger not loaded — wrong project?');
  var out = msgs.length ? 'FAIL:\n• '+msgs.join('\n• ') : 'PASS — BillReturns module wiring looks good.';
  Logger.log(out);
  try { SpreadsheetApp.getUi().alert(out); } catch(e){}
  return {ok: msgs.length === 0, messages: msgs};
}
