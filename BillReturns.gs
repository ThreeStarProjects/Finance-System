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
// Sheet: Bill_Returns (created lazily; schema auto-upgraded on first use).
//
// Schema v2 (Oct 2026) — adds Vendor_CN_No, Deleted. Old 10-col sheets are
// upgraded in place by _getBillReturnsSheet() before any write.
// ============================================================

var BR = { SHEET: 'Bill_Returns' };
var BRC = {
  ID:0, BILL_ID:1, VEND_ID:2, VEND_NAME:3, RET_DATE:4, FY:5, AMOUNT:6,
  REASON:7, NOTES:8, CREATED:9, CN_NO:10, DELETED:11
};
var BRC_TOTAL = 12;
var BR_HEADERS = [
  'Return_ID','Bill_ID','Vendor_ID','Vendor_Name','Return_Date','FY',
  'Return_Amount','Reason','Notes','Created_Date',
  'Vendor_CN_No','Deleted'
];
var BR_WIDTHS  = [120, 110, 110, 230, 120, 80, 130, 260, 240, 160, 150, 80];

// ============================================================
// Sheet accessor — creates on first use; grows schema on each access
// if the live sheet is narrower than BRC_TOTAL (idempotent).
// ============================================================
function _getBillReturnsSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(BR.SHEET);
  if (!sheet) return _buildSheet(ss, BR.SHEET, BR_HEADERS, BR_WIDTHS);
  // Schema upgrade — append missing headers if the sheet is older.
  if (sheet.getLastColumn() < BRC_TOTAL){
    var start = sheet.getLastColumn() + 1;
    var newHdrs = BR_HEADERS.slice(start - 1);
    // Make sure the sheet is wide enough to hold them.
    var maxCols = sheet.getMaxColumns();
    if (maxCols < BRC_TOTAL) sheet.insertColumnsAfter(maxCols, BRC_TOTAL - maxCols);
    sheet.getRange(1, start, 1, newHdrs.length).setValues([newHdrs])
         .setBackground(BRAND.headerBg).setFontColor(BRAND.headerFg)
         .setFontWeight('bold').setFontSize(11);
  }
  return sheet;
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
// recordBillReturn — single-bill credit note (locked wrapper)
//   data = {billRef, returnDate, returnAmount, reason, notes, cnNo}
// ============================================================
function recordBillReturn(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try { return _recordBillReturnCore(data); }
  catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Lock-free core — called by both the single wrapper and the batch wrapper.
// The batch holds ONE outer lock around many of these to avoid nested waitLock.
function _recordBillReturnCore(data){
  data = data || {};
  var amt = _n(data.returnAmount);
  if (amt <= 0) return {success:false, message:'Return amount must be > 0.'};

  var hit = _findBillByIdOrVendorNo(data.billRef);
  if (!hit) return {success:false, message:'Bill not found: '+_s(data.billRef)};
  var row = hit.row;
  if (_s(row[BC.STATUS]) === BILL_STATUS.CANCELLED)
    return {success:false, message:'Bill '+_s(row[BC.ID])+' is cancelled — credit note not applicable.'};

  var billAmount = _n(row[BC.AMOUNT]);
  if (amt > billAmount) return {success:false, message:'Credit note on '+_s(row[BC.ID])+' (₹'+amt+') exceeds bill amount (₹'+billAmount+').'};

  var outstanding = _n(row[BC.OUTSTANDING]);
  var newOut = _round(Math.max(0, outstanding - amt));
  hit.sheet.getRange(hit.rowNum, BC.OUTSTANDING+1).setValue(newOut);
  if (newOut <= 0.01) hit.sheet.getRange(hit.rowNum, BC.STATUS+1).setValue(BILL_STATUS.PAID);

  var retDate = data.returnDate ? _asDate(data.returnDate) : new Date();
  var rSheet = _getBillReturnsSheet();
  var retId = _nextSeqId(rSheet, 'BRET-');
  var newRow = new Array(BRC_TOTAL).fill('');
  newRow[BRC.ID]        = retId;
  newRow[BRC.BILL_ID]   = _s(row[BC.ID]);
  newRow[BRC.VEND_ID]   = _s(row[BC.VEND_ID]);
  newRow[BRC.VEND_NAME] = _s(row[BC.VEND_NAME]);
  newRow[BRC.RET_DATE]  = retDate;
  newRow[BRC.FY]        = _fyFromDate(retDate);
  newRow[BRC.AMOUNT]    = amt;
  newRow[BRC.REASON]    = _s(data.reason);
  newRow[BRC.NOTES]     = _s(data.notes);
  newRow[BRC.CREATED]   = Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss');
  newRow[BRC.CN_NO]     = _s(data.cnNo);
  newRow[BRC.DELETED]   = '';
  rSheet.appendRow(newRow);
  SpreadsheetApp.flush();
  return {success:true, returnId:retId, billId:_s(row[BC.ID]), newOutstanding:newOut};
}

// ============================================================
// recordBillReturnBatch — one credit note split across N bills
//   payload = {returnDate, reason, notes, cnNo,
//              lines: [{billRef, amount}, …]}
// One lock around all writes; per-line results reported.
// ============================================================
function recordBillReturnBatch(payload){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    payload = payload || {};
    var lines = payload.lines || [];
    if (!lines.length) return {success:false, message:'No bills selected.'};

    var results = [], okCount = 0, failCount = 0, returnIds = [];
    for (var i=0;i<lines.length;i++){
      var ln = lines[i] || {};
      var res = _recordBillReturnCore({
        billRef:      ln.billRef,
        returnAmount: _n(ln.amount),
        returnDate:   payload.returnDate,
        reason:       payload.reason,
        notes:        payload.notes,
        cnNo:         payload.cnNo
      });
      results.push({billRef: _s(ln.billRef), ok: !!res.success,
                    returnId: res.returnId || '', message: res.message || ''});
      if (res.success){ okCount++; returnIds.push(res.returnId); }
      else failCount++;
    }
    return {
      success: failCount === 0,
      okCount: okCount, failCount: failCount,
      returnIds: returnIds, results: results,
      message: (failCount === 0)
        ? ('Recorded '+okCount+' credit note'+(okCount===1?'':'s')+'.')
        : (okCount+' recorded, '+failCount+' failed — see results.')
    };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// cancelBillReturn — reverse a credit note
//   Restores the bill's Outstanding by the credit note amount, flips
//   bill status back to Unpaid/Partial if needed, soft-deletes the BRET row.
// ============================================================
function cancelBillReturn(returnId, reason){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var id = _s(returnId);
    if (!id) return {success:false, message:'Return ID required.'};
    var rSheet = _getBillReturnsSheet();
    if (!_hasData(rSheet)) return {success:false, message:'No credit notes on file.'};
    var rows = rSheet.getRange(2,1,rSheet.getLastRow()-1,BRC_TOTAL).getValues();
    var hit = -1;
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][BRC.ID]) === id){ hit = i; break; }
    }
    if (hit < 0) return {success:false, message:'Credit note '+id+' not found.'};
    if (_s(rows[hit][BRC.DELETED]) === 'Yes')
      return {success:false, message:'Credit note '+id+' already cancelled.'};

    var billId = _s(rows[hit][BRC.BILL_ID]);
    var amt = _n(rows[hit][BRC.AMOUNT]);

    // Restore the bill's Outstanding (if the bill still exists & isn't cancelled).
    var bHit = _findBillByIdOrVendorNo(billId);
    if (bHit && _s(bHit.row[BC.STATUS]) !== BILL_STATUS.CANCELLED){
      var curOut = _n(bHit.row[BC.OUTSTANDING]);
      var billAmt = _n(bHit.row[BC.AMOUNT]);
      var paid = _n(bHit.row[BC.PAID]);
      // Cap at what the bill could legitimately owe (amount - paid).
      var restored = _round(Math.min(curOut + amt, Math.max(0, billAmt - paid)));
      bHit.sheet.getRange(bHit.rowNum, BC.OUTSTANDING+1).setValue(restored);
      // Status: Paid → Partial if restored > 0 && paid > 0; Unpaid if paid == 0.
      var newStatus = (restored <= 0.01) ? BILL_STATUS.PAID
                    : (paid > 0)         ? BILL_STATUS.PARTIAL
                                         : BILL_STATUS.UNPAID;
      bHit.sheet.getRange(bHit.rowNum, BC.STATUS+1).setValue(newStatus);
    }

    // Soft-delete the credit note row.
    var rowNum = hit + 2;
    rSheet.getRange(rowNum, BRC.DELETED+1).setValue('Yes');
    // Stamp reason into Notes if provided (append, don't clobber).
    if (_s(reason)){
      var existing = _s(rows[hit][BRC.NOTES]);
      rSheet.getRange(rowNum, BRC.NOTES+1).setValue(
        (existing ? existing + ' | ' : '') + 'CANCELLED: ' + _s(reason)
      );
    }
    SpreadsheetApp.flush();
    return {success:true, returnId:id, billId:billId};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// updateBillReturn — edit metadata (CN No, Reason, Notes) only.
// Amount changes are intentionally refused — cancel + re-create instead,
// so the bill-outstanding reconciliation stays simple.
// ============================================================
function updateBillReturn(returnId, patches){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var id = _s(returnId); if (!id) return {success:false, message:'Return ID required.'};
    var rSheet = _getBillReturnsSheet();
    if (!_hasData(rSheet)) return {success:false, message:'No credit notes on file.'};
    var n = rSheet.getLastRow() - 1;
    var ids = rSheet.getRange(2, BRC.ID+1, n, 1).getValues();
    var hit = -1;
    for (var i=0;i<n;i++){ if (_s(ids[i][0]) === id){ hit = i; break; } }
    if (hit < 0) return {success:false, message:'Credit note '+id+' not found.'};
    var rowNum = hit + 2;
    patches = patches || {};
    if (patches.hasOwnProperty('cnNo'))   rSheet.getRange(rowNum, BRC.CN_NO+1).setValue(_s(patches.cnNo));
    if (patches.hasOwnProperty('reason')) rSheet.getRange(rowNum, BRC.REASON+1).setValue(_s(patches.reason));
    if (patches.hasOwnProperty('notes'))  rSheet.getRange(rowNum, BRC.NOTES+1).setValue(_s(patches.notes));
    SpreadsheetApp.flush();
    return {success:true, returnId:id};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// getBillReturnLedgerEntries — injected into getVendorLedger (Bills.gs).
// Credit notes are DEBITS (reduce vendor's claim on us). Skips soft-deleted.
// ============================================================
function getBillReturnLedgerEntries(vendorId){
  try {
    var vid = _s(vendorId); if (!vid) return [];
    var ss = _getSS(), sheet = ss.getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return [];
    var w = Math.min(sheet.getLastColumn(), BRC_TOTAL);
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
    var out = [];
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][BRC.VEND_ID]) !== vid) continue;
      if (w > BRC.DELETED && _s(rows[i][BRC.DELETED]) === 'Yes') continue;
      var cn = w > BRC.CN_NO ? _s(rows[i][BRC.CN_NO]) : '';
      out.push({
        date: _safeDateStr(rows[i][BRC.RET_DATE]),
        type: 'Credit Note',
        ref: _s(rows[i][BRC.ID]) + (cn ? ' · CN:' + cn : ''),
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
// listBillReturns — table-friendly list for the Credit Notes tab.
//   filters = {vendorId, fy, includeCancelled}
// ============================================================
function listBillReturns(filters){
  try {
    filters = filters || {};
    var sheet = _getSS().getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return {ok:true, rows:[], error:null};
    var w = Math.min(sheet.getLastColumn(), BRC_TOTAL);
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
    var vid = _s(filters.vendorId), fy = _s(filters.fy);
    var showCancelled = !!filters.includeCancelled;
    var out = rows.filter(function(r){
      if (!_s(r[BRC.ID])) return false;
      if (vid && _s(r[BRC.VEND_ID]) !== vid) return false;
      if (fy && fy !== 'all' && _s(r[BRC.FY]) !== fy) return false;
      if (!showCancelled && w > BRC.DELETED && _s(r[BRC.DELETED]) === 'Yes') return false;
      return true;
    }).map(function(r){
      return {
        id:        _s(r[BRC.ID]),
        billId:    _s(r[BRC.BILL_ID]),
        vendorId:  _s(r[BRC.VEND_ID]),
        vendorName:_s(r[BRC.VEND_NAME]),
        date:      _safeDateStr(r[BRC.RET_DATE]),
        fy:        _s(r[BRC.FY]),
        amount:    _n(r[BRC.AMOUNT]),
        reason:    _s(r[BRC.REASON]),
        notes:     _s(r[BRC.NOTES]),
        cnNo:      w > BRC.CN_NO    ? _s(r[BRC.CN_NO])    : '',
        cancelled: w > BRC.DELETED && _s(r[BRC.DELETED]) === 'Yes'
      };
    });
    out.sort(function(a,b){ return (b.date || '').localeCompare(a.date || ''); });
    return {ok:true, rows:out, error:null};
  } catch(e){ return {ok:false, rows:[], error:String(e && e.message || e)}; }
}

// ============================================================
// Dashboard aggregators — all skip soft-deleted rows.
// ============================================================

// Total credit notes received in a given FY.
function getCreditNotesTotalFy(fy){
  try {
    var sheet = _getSS().getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return 0;
    var w = Math.min(sheet.getLastColumn(), BRC_TOTAL);
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
    var tgt = _s(fy); var total = 0;
    for (var i=0;i<rows.length;i++){
      if (!_s(rows[i][BRC.ID])) continue;
      if (w > BRC.DELETED && _s(rows[i][BRC.DELETED]) === 'Yes') continue;
      if (tgt && tgt !== 'all' && _s(rows[i][BRC.FY]) !== tgt) continue;
      total += _n(rows[i][BRC.AMOUNT]);
    }
    return _round(total);
  } catch(e){ Logger.log('getCreditNotesTotalFy: '+e.message); return 0; }
}

// Per-vendor credit note totals for a given FY.
// Returns { 'VEND-0001': 1500, … }
function getCreditNotesByVendorFy(fy){
  try {
    var map = {};
    var sheet = _getSS().getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return map;
    var w = Math.min(sheet.getLastColumn(), BRC_TOTAL);
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
    var tgt = _s(fy);
    for (var i=0;i<rows.length;i++){
      if (!_s(rows[i][BRC.ID])) continue;
      if (w > BRC.DELETED && _s(rows[i][BRC.DELETED]) === 'Yes') continue;
      if (tgt && tgt !== 'all' && _s(rows[i][BRC.FY]) !== tgt) continue;
      var vid = _s(rows[i][BRC.VEND_ID]); if (!vid) continue;
      map[vid] = (map[vid] || 0) + _n(rows[i][BRC.AMOUNT]);
    }
    return map;
  } catch(e){ Logger.log('getCreditNotesByVendorFy: '+e.message); return {}; }
}

// Per-bill credit note totals, FY-scoped (optional). Used by project
// cost rollups to subtract credit notes from the bills they target.
function getCreditNotesByBillId(fy){
  try {
    var map = {};
    var sheet = _getSS().getSheetByName(BR.SHEET);
    if (!sheet || !_hasData(sheet)) return map;
    var w = Math.min(sheet.getLastColumn(), BRC_TOTAL);
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
    var tgt = _s(fy);
    for (var i=0;i<rows.length;i++){
      if (!_s(rows[i][BRC.ID])) continue;
      if (w > BRC.DELETED && _s(rows[i][BRC.DELETED]) === 'Yes') continue;
      if (tgt && tgt !== 'all' && _s(rows[i][BRC.FY]) !== tgt) continue;
      var bid = _s(rows[i][BRC.BILL_ID]); if (!bid) continue;
      map[bid] = (map[bid] || 0) + _n(rows[i][BRC.AMOUNT]);
    }
    return map;
  } catch(e){ Logger.log('getCreditNotesByBillId: '+e.message); return {}; }
}

// ============================================================
// cleanupBillReturnsExtraCols — one-shot cleanup for sheets that
// already got widened to 17 cols by the earlier schema. Deletes columns
// 13 onward (1-based) which held the dropped Taxable/GST_Rate/CGST/SGST/
// IGST. Idempotent — safe to run any number of times.
// ============================================================
function cleanupBillReturnsExtraCols(){
  var sheet = _getSS().getSheetByName(BR.SHEET);
  if (!sheet) return {ok:true, message:'No Bill_Returns sheet yet — nothing to clean.'};
  var now = sheet.getLastColumn();
  if (now <= BRC_TOTAL) return {ok:true, message:'Already '+now+' cols — no cleanup needed.'};
  var extra = now - BRC_TOTAL;
  sheet.deleteColumns(BRC_TOTAL + 1, extra);
  var msg = 'Deleted '+extra+' extra column'+(extra===1?'':'s')+' from Bill_Returns.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){}
  return {ok:true, message:msg};
}

// ============================================================
// Smoke check — run from the Apps Script editor if credit notes
// start misbehaving. Verifies schema wiring and ledger shape.
// Writes nothing to the sheets.
// ============================================================
function _billReturnModuleCheck(){
  var msgs = [];
  if (BR_HEADERS.length !== BRC_TOTAL) msgs.push('BR_HEADERS.length ('+BR_HEADERS.length+') ≠ BRC_TOTAL ('+BRC_TOTAL+')');
  if (BR_WIDTHS.length  !== BRC_TOTAL) msgs.push('BR_WIDTHS.length ('+BR_WIDTHS.length+') ≠ BRC_TOTAL ('+BRC_TOTAL+')');
  if (_findBillByIdOrVendorNo('__nope__') !== null) msgs.push('_findBillByIdOrVendorNo should return null for bogus ref');
  var le = getBillReturnLedgerEntries('__nope__');
  if (!Array.isArray(le) || le.length !== 0) msgs.push('getBillReturnLedgerEntries should return [] for bogus vendor');
  if (typeof getVendorLedger !== 'function') msgs.push('getVendorLedger not loaded — wrong project?');
  // Schema upgrade probe — fetches the sheet, which runs the auto-upgrade path.
  try {
    var s = _getBillReturnsSheet();
    if (s.getLastColumn() < BRC_TOTAL) msgs.push('Bill_Returns is still '+s.getLastColumn()+' cols — schema upgrade did not apply');
  } catch(e){ msgs.push('_getBillReturnsSheet threw: '+e.message); }
  var out = msgs.length ? 'FAIL:\n• '+msgs.join('\n• ') : 'PASS — BillReturns module wiring + schema OK.';
  Logger.log(out);
  try { SpreadsheetApp.getUi().alert(out); } catch(e){}
  return {ok: msgs.length === 0, messages: msgs};
}