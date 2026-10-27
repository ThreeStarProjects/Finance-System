// ============================================================
// 3 Star Projects — Unified Finance System
// MODULE: Salary Advances — single source of truth
// File: SalaryAdvance.gs   (depends on Masters, Transactions, Salary)
//
// ONE sheet: `Advances`
// Every advance-related event — employee request, owner-approved disbursement,
// owner-direct disbursement, salary-deduction recovery, cash/bank recovery —
// writes a single row to this sheet. Nothing about advances is stored anywhere
// else: not in Transactions, not in Salary, not in SalaryAdvanceRequest
// (the old approval sheet is DELETED by the one-shot migration).
//
// Row shape:
//   Advance_ID       REQ-####  (Request) | ADV-####  (Given) | REC-####  (Recovery)
//   Date             the transaction date
//   Employee         must exist in M_Employee (soft-validated)
//   Type             'Request' | 'Given' | 'Recovery'
//   Status           Pending | Approved | Rejected | Cancelled   (for Request)
//                    Open | Closed                                (for Given)
//                    Posted                                        (for Recovery)
//   Amount           always positive; semantics vary by Type
//   Source_Account   bank/cash account the money came from (Given) or
//                    'Salary Deduction' / destination account (Recovery)
//   Applied_To       Recovery only — comma-separated Given IDs this recovery closed
//   Linked_Ref       Given only — Request_ID that led to this disbursement, or blank
//                    (owner-direct)
//   Requested_By     who submitted the request (employee entry) or 'Owner' (direct)
//   Approved_By      stamped when a Request is Approved or Rejected
//   Approved_Date    same
//   Rejection_Reason filled when Status = Rejected
//   Payment_Mode     for Given / Recovery
//   Notes
//   Created_Date
//   Created_By
//   Deleted          'Yes' = soft-delete (ignored by all readers)
//
// Setup entry: run `setupAdvancesModule()` once from the editor.
// Migration: run `migrateAdvancesToNewSheet()` ONCE to wipe the old scattered
// advance data and start fresh with the new sheet. HARD-DELETE — not reversible.
// ============================================================

// ── Sheet name
var ADV = { SHEET: 'Advances', LEGACY_REQUEST_SHEET: 'SalaryAdvanceRequest' };

// ── Column map (0-based)
var AC = {
  ID:0, DATE:1, EMPLOYEE:2, TYPE:3, STATUS:4, AMOUNT:5,
  SOURCE_ACCOUNT:6, APPLIED_TO:7, LINKED_REF:8,
  REQUESTED_BY:9, APPROVED_BY:10, APPROVED_DATE:11, REJECTION_REASON:12,
  PAYMENT_MODE:13, NOTES:14, CREATED_DATE:15, CREATED_BY:16, DELETED:17
};
var AC_TOTAL = 18;

var ADV_HEADERS = [
  'Advance_ID','Date','Employee','Type','Status','Amount',
  'Source_Account','Applied_To','Linked_Ref',
  'Requested_By','Approved_By','Approved_Date','Rejection_Reason',
  'Payment_Mode','Notes','Created_Date','Created_By','Deleted'
];
var ADV_WIDTHS  = [150, 110, 200, 100, 110, 120, 180, 220, 150, 160, 160, 160, 300, 140, 300, 160, 180, 90];

var ADV_TYPE = { REQUEST: 'Request', GIVEN: 'Given', RECOVERY: 'Recovery' };

// Status vocab, scoped per Type:
//   Request : Pending | Approved | Rejected | Cancelled
//   Given   : Open    | Closed
//   Recovery: Posted
var ADV_STATUS = {
  PENDING: 'Pending', APPROVED: 'Approved', REJECTED: 'Rejected', CANCELLED: 'Cancelled',
  OPEN: 'Open', CLOSED: 'Closed',
  POSTED: 'Posted'
};

var ADV_SALARY_DEDUCTION = 'Salary Deduction';

// ============================================================
// SETUP — idempotent. Called from setupMasters.
// ============================================================
function setupAdvancesModule(){
  var ss = _getSS();
  var sheet = _buildSheet(ss, ADV.SHEET, ADV_HEADERS, ADV_WIDTHS);
  try {
    _validateColSoft(sheet, AC.TYPE+1,
      [ADV_TYPE.REQUEST, ADV_TYPE.GIVEN, ADV_TYPE.RECOVERY]);
    _validateColSoft(sheet, AC.STATUS+1, [
      ADV_STATUS.PENDING, ADV_STATUS.APPROVED, ADV_STATUS.REJECTED, ADV_STATUS.CANCELLED,
      ADV_STATUS.OPEN, ADV_STATUS.CLOSED,
      ADV_STATUS.POSTED
    ]);
    _validateColSoft(sheet, AC.DELETED+1, ['Yes','No']);
  } catch(e){ Logger.log('setupAdvancesModule validators: '+e.message); }
  return { sheet: sheet.getName() };
}

// Backward-compat alias — setupMasters still calls setupSalaryAdvance.
function setupSalaryAdvance(){
  var r = setupAdvancesModule();
  var msg = 'Advances module ready — sheet: ' + r.sheet +
            '\n\nIf you have NOT yet run migrateAdvancesToNewSheet(), do it now ' +
            'from the editor. It HARD-DELETES all advance-related data in ' +
            'Transactions_<FY>, zeros ADV_DED on Salary rows, and removes the ' +
            'legacy SalaryAdvanceRequest sheet. Not reversible.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, message:msg};
}

// ============================================================
// INTERNAL HELPERS
// ============================================================
function _advSheet(){
  var s = _getSS().getSheetByName(ADV.SHEET);
  if (!s) { setupAdvancesModule(); s = _getSS().getSheetByName(ADV.SHEET); }
  return s;
}
function _advReadAll(){
  var s = _advSheet();
  if (!s || s.getLastRow() < 2) return {sheet:s, rows:[], raw:[]};
  var w = Math.min(s.getLastColumn(), AC_TOTAL);
  var raw = s.getRange(2, 1, s.getLastRow()-1, w).getValues();
  var rows = raw.map(function(r, i){ return _advSerialize(r, i+2); });
  return {sheet:s, rows:rows, raw:raw};
}
function _advSerialize(r, rowNum){
  return {
    rowNum:          rowNum || 0,
    id:              _s(r[AC.ID]),
    date:            _safeDateStr(r[AC.DATE]),
    dateMs:          (function(v){ if(!v) return 0; var d=(v instanceof Date)?v:new Date(v); var t=d.getTime(); return isNaN(t)?0:t; })(r[AC.DATE]),
    employee:        _s(r[AC.EMPLOYEE]),
    type:            _s(r[AC.TYPE]),
    status:          _s(r[AC.STATUS]),
    amount:          _n(r[AC.AMOUNT]) || 0,
    sourceAccount:   _s(r[AC.SOURCE_ACCOUNT]),
    appliedTo:       _s(r[AC.APPLIED_TO]),
    linkedRef:       _s(r[AC.LINKED_REF]),
    requestedBy:     _s(r[AC.REQUESTED_BY]),
    approvedBy:      _s(r[AC.APPROVED_BY]),
    approvedDate:    _safeDateStr(r[AC.APPROVED_DATE]),
    rejectionReason: _s(r[AC.REJECTION_REASON]),
    paymentMode:     _s(r[AC.PAYMENT_MODE]),
    notes:           _s(r[AC.NOTES]),
    createdDate:     _safeDateStr(r[AC.CREATED_DATE]),
    createdBy:       _s(r[AC.CREATED_BY]),
    deleted:         _s(r[AC.DELETED]) === 'Yes'
  };
}
function _advWriteRow(sheet, rowObj){
  var row = new Array(AC_TOTAL).fill('');
  row[AC.ID]               = rowObj.id;
  row[AC.DATE]             = rowObj.date || new Date();
  row[AC.EMPLOYEE]         = rowObj.employee || '';
  row[AC.TYPE]             = rowObj.type || '';
  row[AC.STATUS]           = rowObj.status || '';
  row[AC.AMOUNT]           = rowObj.amount || 0;
  row[AC.SOURCE_ACCOUNT]   = rowObj.sourceAccount || '';
  row[AC.APPLIED_TO]       = rowObj.appliedTo || '';
  row[AC.LINKED_REF]       = rowObj.linkedRef || '';
  row[AC.REQUESTED_BY]     = rowObj.requestedBy || '';
  row[AC.APPROVED_BY]      = rowObj.approvedBy || '';
  row[AC.APPROVED_DATE]    = rowObj.approvedDate || '';
  row[AC.REJECTION_REASON] = rowObj.rejectionReason || '';
  row[AC.PAYMENT_MODE]     = rowObj.paymentMode || '';
  row[AC.NOTES]            = rowObj.notes || '';
  row[AC.CREATED_DATE]     = new Date();
  row[AC.CREATED_BY]       = _advWho();
  row[AC.DELETED]          = 'No';
  var targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();
  return targetRow;
}
function _advWho(){
  try { return Session.getActiveUser().getEmail() || ''; } catch(e){ return ''; }
}
function _advNextId(type){
  // Prefix + yyyymmddHHmmss timestamp + 3-digit random for uniqueness per lock window.
  var pfx = (type === ADV_TYPE.REQUEST) ? 'REQ-'
          : (type === ADV_TYPE.GIVEN)   ? 'ADV-'
          : (type === ADV_TYPE.RECOVERY)? 'REC-' : 'ADV-';
  var ts = Utilities.formatDate(new Date(),'GMT+5:30','yyyyMMddHHmmss');
  return pfx + ts + '-' + Math.floor(Math.random()*900 + 100);
}
function _advFindById(id){
  var all = _advReadAll().rows;
  for (var i=0; i<all.length; i++) if (all[i].id === id) return all[i];
  return null;
}
function _advResolveEmployee(name){
  // Soft-validate: look up in M_Employee list if available, else accept as given.
  var clean = _s(name);
  if (!clean) return {ok:false, message:'Employee is required'};
  try {
    var emps = (typeof getEmployeeNames === 'function') ? (getEmployeeNames() || []) : [];
    if (!emps.length) return {ok:true, name:clean};
    var norm = clean.toLowerCase();
    for (var i=0; i<emps.length; i++){
      if (String(emps[i]).toLowerCase() === norm) return {ok:true, name:emps[i]};
    }
    return {ok:false, message:'Employee "' + clean + '" is not in M_Employee master.'};
  } catch(e){ return {ok:true, name:clean}; }
}

// ============================================================
// PUBLIC API — Requests
// ============================================================
// Employee submits an advance request (from ExpenseForm.html).
function advRequest(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    var empRes = _advResolveEmployee(data.employee);
    if (!empRes.ok) return {success:false, message:empRes.message};
    var amt = _n(data.amount);
    if (amt <= 0) return {success:false, message:'Amount must be greater than zero'};

    // Dedupe guard: same employee + same amount + still Pending → refuse.
    var all = _advReadAll().rows;
    var dup = all.some(function(r){
      return !r.deleted && r.type === ADV_TYPE.REQUEST && r.status === ADV_STATUS.PENDING
          && r.employee === empRes.name && Math.abs(r.amount - amt) < 0.5;
    });
    if (dup){
      return {success:false, message:'A pending request for ' + empRes.name +
                                    ' with amount ₹' + _round(amt) + ' already exists.'};
    }

    var id = _advNextId(ADV_TYPE.REQUEST);
    var row = {
      id: id, date: data.date ? _asDate(data.date) : new Date(),
      employee: empRes.name, type: ADV_TYPE.REQUEST, status: ADV_STATUS.PENDING,
      amount: amt,
      requestedBy: _s(data.requestedBy) || empRes.name,
      notes: _s(data.reason) || _s(data.notes)
    };
    _advWriteRow(_advSheet(), row);
    return {success:true, requestId:id, message:'Request ' + id + ' submitted for ' + empRes.name + ' — ₹' + _round(amt)};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Backward-compat alias (ExpenseForm.html calls this name).
function submitAdvanceRequest(data){ return advRequest(data); }

// Owner approves a request → writes a Given row linked back to the Request.
function advApproveRequest(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    var reqId = _s(data.requestId);
    if (!reqId) return {success:false, message:'Request ID is required'};
    var req = _advFindById(reqId);
    if (!req) return {success:false, message:'Request ' + reqId + ' not found'};
    if (req.type !== ADV_TYPE.REQUEST) return {success:false, message:'Not a request row.'};
    if (req.status !== ADV_STATUS.PENDING){
      return {success:false, message:'Request is ' + req.status + ' — only Pending can be approved.'};
    }

    var now = new Date();
    var sheet = _advSheet();

    // 1. Update the Request row → Approved
    sheet.getRange(req.rowNum, AC.STATUS+1).setValue(ADV_STATUS.APPROVED);
    sheet.getRange(req.rowNum, AC.APPROVED_BY+1).setValue(_s(data.approvedBy) || _advWho());
    sheet.getRange(req.rowNum, AC.APPROVED_DATE+1).setValue(now);

    // 2. Write the Given row
    var givenId = _advNextId(ADV_TYPE.GIVEN);
    var disburseDate = data.date || data.txDate;  // accept either field name
    _advWriteRow(sheet, {
      id: givenId,
      date: disburseDate ? _asDate(disburseDate) : now,
      employee: req.employee, type: ADV_TYPE.GIVEN, status: ADV_STATUS.OPEN,
      amount: req.amount,
      sourceAccount: _s(data.source) || '',
      paymentMode:   _s(data.paymentMode) || 'Cash',
      linkedRef: req.id,
      requestedBy: req.requestedBy,
      approvedBy:  _s(data.approvedBy) || _advWho(),
      approvedDate: now,
      notes: _s(data.notes) || ('Approved ' + reqId)
    });
    return {success:true, requestId:req.id, givenId:givenId,
            message:'Approved ' + req.id + ' — disbursed as ' + givenId};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}
function approveAdvanceRequest(data){ return advApproveRequest(data); }

// Reject a Request.
function advRejectRequest(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    var reqId = _s(data.requestId);
    if (!reqId) return {success:false, message:'Request ID is required'};
    var req = _advFindById(reqId);
    if (!req) return {success:false, message:'Request ' + reqId + ' not found'};
    if (req.type !== ADV_TYPE.REQUEST) return {success:false, message:'Not a request row.'};
    if (req.status !== ADV_STATUS.PENDING){
      return {success:false, message:'Request is ' + req.status + ' — only Pending can be rejected.'};
    }
    var sheet = _advSheet();
    sheet.getRange(req.rowNum, AC.STATUS+1).setValue(ADV_STATUS.REJECTED);
    sheet.getRange(req.rowNum, AC.APPROVED_BY+1).setValue(_s(data.approvedBy) || _advWho());
    sheet.getRange(req.rowNum, AC.APPROVED_DATE+1).setValue(new Date());
    sheet.getRange(req.rowNum, AC.REJECTION_REASON+1).setValue(_s(data.reason) || '');
    return {success:true, requestId:req.id, message:'Rejected ' + req.id};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}
function rejectAdvanceRequest(data){ return advRejectRequest(data); }

// Employee or owner cancels a Pending request.
function advCancelRequest(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    var reqId = _s(data.requestId);
    if (!reqId) return {success:false, message:'Request ID is required'};
    var req = _advFindById(reqId);
    if (!req) return {success:false, message:'Request ' + reqId + ' not found'};
    if (req.type !== ADV_TYPE.REQUEST) return {success:false, message:'Not a request row.'};
    if (req.status !== ADV_STATUS.PENDING){
      return {success:false, message:'Request is ' + req.status + ' — only Pending can be cancelled.'};
    }
    var sheet = _advSheet();
    sheet.getRange(req.rowNum, AC.STATUS+1).setValue(ADV_STATUS.CANCELLED);
    sheet.getRange(req.rowNum, AC.APPROVED_DATE+1).setValue(new Date());
    return {success:true, requestId:req.id, message:'Cancelled ' + req.id};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}
function cancelAdvanceRequest(data){ return advCancelRequest(data); }

// ============================================================
// PUBLIC API — Given (owner-direct, no prior request)
// ============================================================
function advRecordGiven(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    var empRes = _advResolveEmployee(data.employee);
    if (!empRes.ok) return {success:false, message:empRes.message};
    var amt = _n(data.amount);
    if (amt <= 0) return {success:false, message:'Amount must be greater than zero'};
    // Not requiring a source account at the backend — the frontend can enforce
    // it. Keeps direct scripting paths flexible.

    var now = new Date();
    var id = _advNextId(ADV_TYPE.GIVEN);
    var row = {
      id: id, date: data.date ? _asDate(data.date) : now,
      employee: empRes.name, type: ADV_TYPE.GIVEN, status: ADV_STATUS.OPEN,
      amount: amt,
      sourceAccount: _s(data.source),
      paymentMode:   _s(data.paymentMode) || 'Cash',
      requestedBy: 'Owner',
      notes: _s(data.notes) || ('Direct advance to ' + empRes.name)
    };
    _advWriteRow(_advSheet(), row);
    // Also seed any already-recovered amount if the operator passed it on the
    // same row (replicates the old recordAdvanceGiven behaviour so old UI
    // keeps working). Treats it as a first recovery event.
    var alreadyRecovered = _n(data.alreadyRecovered);
    if (alreadyRecovered > 0){
      advRecordRecovery({
        employee: empRes.name, amount: alreadyRecovered,
        date: row.date, source: ADV_SALARY_DEDUCTION,
        notes: 'Already recovered at advance creation'
      });
    }
    return {success:true, id:id,
            message:'Advance ' + id + ' recorded — ' + empRes.name + ', ₹' + _round(amt)};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}
function recordAdvanceGiven(data){ return advRecordGiven(data); }

// ============================================================
// PUBLIC API — Recovery (FIFO, salary deduction or outside)
// ============================================================
function advRecordRecovery(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    data = data || {};
    var empRes = _advResolveEmployee(data.employee);
    if (!empRes.ok) return {success:false, message:empRes.message};
    var amt = _n(data.amount);
    if (amt <= 0) return {success:false, message:'Amount must be greater than zero'};

    // Find open Given rows for this employee, oldest first.
    var all = _advReadAll().rows;
    var open = all.filter(function(r){
      return !r.deleted && r.type === ADV_TYPE.GIVEN && r.status === ADV_STATUS.OPEN
          && r.employee === empRes.name;
    }).sort(function(a,b){ return a.dateMs - b.dateMs; });

    // Compute each Given's current outstanding from already-posted Recoveries
    // that referenced it.
    var postedByGiven = {};
    all.forEach(function(r){
      if (r.deleted || r.type !== ADV_TYPE.RECOVERY) return;
      var refs = r.appliedTo ? r.appliedTo.split(',').map(function(x){ return x.trim(); }).filter(Boolean) : [];
      if (!refs.length) return;
      // Allocate the recovery amount evenly across its linked Givens is wrong — the
      // Applied_To was written as "ID:amt" by this function. Parse that if present.
      var tokens = r.appliedTo.split(',').map(function(x){ return x.trim(); }).filter(Boolean);
      tokens.forEach(function(t){
        var m = t.match(/^(.+?):([\d.]+)$/);
        if (m){
          postedByGiven[m[1]] = (postedByGiven[m[1]] || 0) + (parseFloat(m[2]) || 0);
        } else {
          // Legacy comma list without amounts — attribute whole recovery amount to the first ref.
          postedByGiven[t] = (postedByGiven[t] || 0) + r.amount;
        }
      });
    });

    var totalOpen = 0;
    open.forEach(function(g){
      g._outstanding = Math.max(0, g.amount - (postedByGiven[g.id] || 0));
      totalOpen += g._outstanding;
    });
    if (totalOpen <= 0){
      return {success:false, message:'No outstanding advances for ' + empRes.name};
    }
    if (amt > totalOpen + 0.5){
      return {success:false,
              message:'Recovery of ₹' + _round(amt) + ' exceeds outstanding ₹' +
                      _round(totalOpen) + ' for ' + empRes.name};
    }

    // FIFO-apply
    var remaining = amt, appliedParts = [], closedIds = [];
    var sheet = _advSheet();
    open.forEach(function(g){
      if (remaining <= 0.5) return;
      var use = Math.min(g._outstanding, remaining);
      if (use <= 0) return;
      appliedParts.push(g.id + ':' + _round(use));
      remaining -= use;
      if (use >= g._outstanding - 0.5){
        sheet.getRange(g.rowNum, AC.STATUS+1).setValue(ADV_STATUS.CLOSED);
        closedIds.push(g.id);
      }
    });

    // Write the Recovery row
    var now = new Date();
    var recId = _advNextId(ADV_TYPE.RECOVERY);
    _advWriteRow(sheet, {
      id: recId, date: data.date ? _asDate(data.date) : now,
      employee: empRes.name, type: ADV_TYPE.RECOVERY, status: ADV_STATUS.POSTED,
      amount: _round(amt),
      sourceAccount: _s(data.source) || ADV_SALARY_DEDUCTION,
      paymentMode:   _s(data.paymentMode) || '',
      appliedTo: appliedParts.join(','),
      notes: _s(data.notes)
    });

    var msg = 'Recovery ' + recId + ' recorded — ' + empRes.name + ', ₹' + _round(amt) +
              ' applied across ' + appliedParts.length + ' advance(s)' +
              (closedIds.length ? ' (' + closedIds.length + ' closed)' : '');
    return {success:true, id:recId, applied:appliedParts, closed:closedIds, message:msg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}
function recordAdvanceRecovery(data){ return advRecordRecovery(data); }

// ============================================================
// READERS — used by Salary.gs, SalaryDashboard.gs, Home_FINAL_FIXED.html
// ============================================================

// Per-employee outstanding list — used by Salary Calculator panel + Record
// Recovery modal. Shape matches the old getEmployeeOutstandingAdvances so
// the frontend works unchanged.
function getEmployeeOutstandingAdvances(empName){
  try {
    var name = _s(empName);
    if (!name) return _envOk({ employee:'', total:0, count:0, rows:[],
                                totalGiven:0, totalRecovered:0, recoveryEvents:[] });

    var all = _advReadAll().rows.filter(function(r){ return !r.deleted; });
    var mine = all.filter(function(r){ return r.employee === name; });

    // Index of recoveries by the Given they applied to.
    var postedByGiven = {};
    var recoveryEvents = [];
    mine.forEach(function(r){
      if (r.type !== ADV_TYPE.RECOVERY) return;
      recoveryEvents.push({id:r.id, date:r.date, amount:r.amount,
                           source:r.sourceAccount, note:r.notes});
      r.appliedTo.split(',').map(function(x){ return x.trim(); }).filter(Boolean)
        .forEach(function(t){
          var m = t.match(/^(.+?):([\d.]+)$/);
          if (m) postedByGiven[m[1]] = (postedByGiven[m[1]] || 0) + (parseFloat(m[2]) || 0);
          else   postedByGiven[t]    = (postedByGiven[t]    || 0) + r.amount;
        });
    });

    var totalGiven = 0, totalRecovered = 0, outRows = [];
    mine.forEach(function(r){
      if (r.type !== ADV_TYPE.GIVEN) return;
      var rec = postedByGiven[r.id] || 0;
      totalGiven += r.amount;
      totalRecovered += rec;
      var outstanding = Math.max(0, r.amount - rec);
      if (r.status === ADV_STATUS.OPEN && outstanding > 0.01){
        outRows.push({
          id: r.id, date: r.date, amount: _round(r.amount),
          recovered: _round(rec), outstanding: _round(outstanding)
        });
      }
    });
    outRows.sort(function(a,b){ return (a.date||'').localeCompare(b.date||''); });
    recoveryEvents.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
    var total = outRows.reduce(function(s,r){ return s + r.outstanding; }, 0);

    return _envOk({
      employee:       name,
      total:          _round(total),
      count:          outRows.length,
      rows:           outRows,
      totalGiven:     _round(totalGiven),
      totalRecovered: _round(totalRecovered),
      recoveryEvents: recoveryEvents
    });
  } catch(e){ return _envErr(e); }
}

// Dashboard advances KPIs. One pass, returns totals the dashboard needs.
function advGetSummary(fy){
  try {
    fy = _s(fy);
    // Normalize UI tokens: 'current' → actual current FY string. Blank | 'all'
    // means every FY (no filter). Anything else is a literal FY like '2026-27'.
    if (fy === 'current') fy = (typeof _currentFY === 'function') ? _currentFY() : '';
    var rows = _advReadAll().rows.filter(function(r){ return !r.deleted; });
    if (fy && fy !== 'all'){
      rows = rows.filter(function(r){
        if (!r.dateMs) return false;
        var d = new Date(r.dateMs);
        var yrFy = _fyFromDate(d);
        return yrFy === fy;
      });
    }
    var totalGiven = 0, totalRecovered = 0;
    var openByEmp = {}, giveById = {};
    var postedByGiven = {};
    rows.forEach(function(r){
      if (r.type === ADV_TYPE.GIVEN){
        totalGiven += r.amount;
        giveById[r.id] = r;
      } else if (r.type === ADV_TYPE.RECOVERY){
        totalRecovered += r.amount;
        r.appliedTo.split(',').map(function(x){ return x.trim(); }).filter(Boolean)
          .forEach(function(t){
            var m = t.match(/^(.+?):([\d.]+)$/);
            if (m) postedByGiven[m[1]] = (postedByGiven[m[1]] || 0) + (parseFloat(m[2]) || 0);
            else   postedByGiven[t]    = (postedByGiven[t]    || 0) + r.amount;
          });
      }
    });
    var outstanding = 0;
    Object.keys(giveById).forEach(function(id){
      var g = giveById[id];
      var o = Math.max(0, g.amount - (postedByGiven[id] || 0));
      if (g.status === ADV_STATUS.OPEN && o > 0){
        outstanding += o;
        openByEmp[g.employee] = true;
      }
    });
    return _envOk({
      fy: fy || 'all',
      totalGiven:     _round(totalGiven),
      totalRecovered: _round(totalRecovered),
      outstanding:    _round(outstanding),
      employeeCount:  Object.keys(openByEmp).length
    });
  } catch(e){ return _envErr(e); }
}

// Advances ledger — dated history for an employee (or all). Used by Salary
// tab's "Advances Ledger" view. Shape matches old getAdvancesLedger.
function getAdvancesLedger(employeeName){
  try {
    var filter = _s(employeeName);
    var all = _advReadAll().rows.filter(function(r){ return !r.deleted; });
    if (filter) all = all.filter(function(r){ return r.employee === filter; });

    // Summary mode: no filter → per-employee totals
    if (!filter){
      var byEmp = {};
      var postedByGiven = {};
      all.forEach(function(r){
        if (r.type !== ADV_TYPE.RECOVERY) return;
        r.appliedTo.split(',').map(function(x){ return x.trim(); }).filter(Boolean)
          .forEach(function(t){
            var m = t.match(/^(.+?):([\d.]+)$/);
            if (m) postedByGiven[m[1]] = (postedByGiven[m[1]] || 0) + (parseFloat(m[2]) || 0);
            else   postedByGiven[t]    = (postedByGiven[t]    || 0) + r.amount;
          });
      });
      all.forEach(function(r){
        if (r.type !== ADV_TYPE.GIVEN) return;
        var emp = r.employee;
        var rec = postedByGiven[r.id] || 0;
        if (!byEmp[emp]) byEmp[emp] = {employee:emp, totalGiven:0, totalRecovered:0, outstanding:0, givenCount:0};
        byEmp[emp].totalGiven += r.amount;
        byEmp[emp].totalRecovered += rec;
        byEmp[emp].outstanding += Math.max(0, r.amount - rec);
        byEmp[emp].givenCount++;
      });
      var employees = Object.keys(byEmp).map(function(e){
        var x = byEmp[e];
        return {employee:x.employee, totalGiven:_round(x.totalGiven),
                totalRecovered:_round(x.totalRecovered),
                outstanding:_round(x.outstanding), givenCount:x.givenCount};
      }).sort(function(a,b){ return b.outstanding - a.outstanding; });
      var totGiven = employees.reduce(function(s,e){return s+e.totalGiven;},0);
      var totRec   = employees.reduce(function(s,e){return s+e.totalRecovered;},0);
      var totOut   = employees.reduce(function(s,e){return s+e.outstanding;},0);
      var empsWithBal = employees.filter(function(e){ return e.outstanding > 0; }).length;
      return _envOk({
        mode:'summary',
        totalGiven: _round(totGiven),
        totalRecovered: _round(totRec),
        outstanding: _round(totOut),
        totalOutstanding: _round(totOut),  // alias for frontend backward compat
        employeeCount: empsWithBal,
        employees: employees
      });
    }

    // Detail mode: single employee → chronological ledger
    var given = all.filter(function(r){ return r.type === ADV_TYPE.GIVEN; });
    var recs  = all.filter(function(r){ return r.type === ADV_TYPE.RECOVERY; });
    var postedByG = {};
    recs.forEach(function(r){
      r.appliedTo.split(',').map(function(x){ return x.trim(); }).filter(Boolean)
        .forEach(function(t){
          var m = t.match(/^(.+?):([\d.]+)$/);
          if (m) postedByG[m[1]] = (postedByG[m[1]] || 0) + (parseFloat(m[2]) || 0);
        });
    });
    var events = [];
    given.forEach(function(g){
      events.push({date:g.date, type:'Given', id:g.id, amount:g.amount,
                   status:g.status, note:g.notes, _delta: g.amount});
    });
    recs.forEach(function(r){
      // Frontend labels this type 'Recovered' — map 'Recovery' → 'Recovered' here.
      events.push({date:r.date, type:'Recovered', id:r.id, amount:r.amount,
                   source:r.sourceAccount, note:r.notes, _delta: -r.amount});
    });
    events.sort(function(a,b){ return (a.date||'').localeCompare(b.date||''); });
    var running = 0;
    events.forEach(function(e){ running += e._delta; e.balance = _round(running); delete e._delta; });
    var tg = given.reduce(function(s,g){ return s+g.amount; }, 0);
    var tr = recs.reduce(function(s,r){  return s+r.amount; }, 0);
    var out = Math.max(0, tg - tr);
    return _envOk({
      mode:'detail',
      employee: filter,
      summary: { totalGiven:_round(tg), totalRecovered:_round(tr), outstanding:_round(out) },
      events: events
    });
  } catch(e){ return _envErr(e); }
}

// Pending-requests queue for the dashboard. Shape matches old
// listPendingAdvanceRequests so the frontend works unchanged.
function listPendingAdvanceRequests(filters){
  try {
    filters = filters || {};
    var limit = filters.limit || 100;
    var rows = _advReadAll().rows.filter(function(r){
      return !r.deleted && r.type === ADV_TYPE.REQUEST && r.status === ADV_STATUS.PENDING;
    }).sort(function(a,b){ return b.dateMs - a.dateMs; });
    var sliced = rows.slice(0, limit).map(function(r){
      return {id:r.id, requestDate:r.date, employee:r.employee,
              amount:r.amount, reason:r.notes, status:r.status};
    });
    var total = rows.reduce(function(s,r){ return s + r.amount; }, 0);
    return _envOk({rows:sliced, total:_round(total), count:rows.length});
  } catch(e){ return _envErr(e); }
}

// List requests for an employee (or all) — used by employee form "My recent requests".
function listAdvanceRequests(filters){
  try {
    filters = filters || {};
    var emp = _s(filters.employee);
    var rows = _advReadAll().rows.filter(function(r){
      if (r.deleted) return false;
      if (r.type !== ADV_TYPE.REQUEST) return false;
      if (emp && r.employee !== emp) return false;
      return true;
    }).sort(function(a,b){ return b.dateMs - a.dateMs; });
    var out = rows.map(function(r){
      return {id:r.id, requestDate:r.date, employee:r.employee, amount:r.amount,
              reason:r.notes, status:r.status, approvedBy:r.approvedBy,
              approvedDate:r.approvedDate, rejectionReason:r.rejectionReason};
    });
    // Return the standard envelope — the employee ExpenseForm reads res.data.rows.
    return _envOk({rows: out, count: out.length});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// MIGRATION — one-shot wipe. Run ONCE from the editor.
// ============================================================
function migrateAdvancesToNewSheet(){
  var ss = _getSS();
  var report = { txnRowsDeleted: 0, salaryRowsZeroed: 0, legacyRequestsCopied: 0, legacyRequestsDropped: 0 };

  // 1. Build the new Advances sheet (idempotent).
  setupAdvancesModule();

  // 2. Hard-delete all advance-related rows from every Transactions_<FY> sheet.
  //    An "advance row" = any row that would have been read as such before:
  //      - RECOVERY in {Pending, Recovered}, OR
  //      - CATEGORY = 'Advance Recovered from Employee', OR
  //      - CATEGORY in {Salary Advance, Advance to Employee}
  //    Hard delete walks bottom-up so row indices stay valid.
  var txnPrefix = (typeof T !== 'undefined' && T.TXN_PREFIX) ? T.TXN_PREFIX : 'Transactions_';
  ss.getSheets().forEach(function(sheet){
    var name = sheet.getName();
    if (name.indexOf(txnPrefix) !== 0) return;
    if (sheet.getLastRow() < 2) return;
    var w = Math.min(sheet.getLastColumn(), TC_TOTAL);
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
    for (var i = data.length - 1; i >= 0; i--){
      var r = data[i];
      var rec = _s(r[TC.RECOVERY]);
      var cat = _s(r[TC.CATEGORY]);
      var isAdvance = (rec === 'Pending' || rec === 'Recovered'
                    || cat === 'Advance Recovered from Employee'
                    || cat === 'Salary Advance'
                    || cat === 'Advance to Employee');
      if (!isAdvance) continue;
      sheet.deleteRow(i + 2);
      report.txnRowsDeleted++;
    }
  });

  // 3. Zero out ADV_DED (and any recovered column) on every historical Salary
  //    record. The user explicitly asked for retroactive zeroing — payroll
  //    net salary math for those rows will shift accordingly.
  var salSheet = ss.getSheetByName((typeof P !== 'undefined' ? P.SALARY : 'Salary'));
  if (salSheet && salSheet.getLastRow() >= 2){
    var sw = Math.min(salSheet.getLastColumn(), (typeof SC_TOTAL !== 'undefined' ? SC_TOTAL : salSheet.getLastColumn()));
    var sData = salSheet.getRange(2, 1, salSheet.getLastRow()-1, sw).getValues();
    for (var si = 0; si < sData.length; si++){
      var ded = _n(sData[si][SC.ADV_DED]);
      if (ded <= 0) continue;
      salSheet.getRange(si+2, SC.ADV_DED+1).setValue(0);
      // Re-add the deduction to Net so the row stays internally consistent.
      // (The Net column constant in SC is SC.NET — not SC.NET_SALARY.)
      if (typeof SC.NET === 'number'){
        var net = _n(sData[si][SC.NET]);
        salSheet.getRange(si+2, SC.NET+1).setValue(_round(net + ded));
      }
      report.salaryRowsZeroed++;
    }
  }

  // 4. Delete the legacy SalaryAdvanceRequest sheet entirely.
  var legacy = ss.getSheetByName(ADV.LEGACY_REQUEST_SHEET);
  if (legacy){
    try { ss.deleteSheet(legacy); report.legacyRequestsDropped++; }
    catch(e){ Logger.log('could not delete legacy sheet: '+e.message); }
  }

  SpreadsheetApp.flush();
  var msg = 'Advances migration complete.\n\n' +
            'Transactions rows hard-deleted: ' + report.txnRowsDeleted + '\n' +
            'Salary rows ADV_DED zeroed:    ' + report.salaryRowsZeroed + '\n' +
            'Legacy SalaryAdvanceRequest:    ' + (report.legacyRequestsDropped ? 'deleted' : 'absent') + '\n\n' +
            'The new ' + ADV.SHEET + ' sheet is now the sole source of truth. ' +
            'All advance-related UI (employee form, owner form, dashboard queue, ' +
            'salary calc, record-recovery modal) writes and reads here.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, report:report, message:msg};
}

// ============================================================
// SHIM for Salary.gs — salary save calls this when ADV_DED > 0.
// Replaces the old _markAdvancesRecovered that touched Transactions.
// ============================================================
function recordSalaryDeductionRecovery(empName, amount, monthLabel, salaryId){
  if (!empName) return {applied:0, remaining:0};
  var amt = _n(amount);
  if (amt <= 0) return {applied:0, remaining:0};
  var res = advRecordRecovery({
    employee: empName, amount: amt,
    date: new Date(),
    source: ADV_SALARY_DEDUCTION,
    notes: 'Salary deduction · ' + _s(monthLabel) + (salaryId ? ' · ' + salaryId : '')
  });
  if (!res.success){
    // Not fatal — salary already saved. Just leave the recovery unlogged
    // (operator can post it later via Record Recovery) and surface to caller.
    return {applied:0, remaining:amt, error:res.message};
  }
  return {applied:amt, remaining:0, recoveryId:res.id};
}

// ============================================================
// Dashboard advance tile helper — returns the three numbers the KPI needs.
// Keeps shape {total, count, employeeCount} that dashboard-side code expects.
// ============================================================
function getAdvancesGivenFy(fy){
  try {
    var sum = advGetSummary(fy);
    if (!sum.ok) return sum;
    var d = sum.data;
    // "Advances given this FY" tile wants Given gross — not outstanding.
    // We also surface outstanding + recovered so callers who want them can read.
    return _envOk({
      fy: d.fy,
      total: d.totalGiven,
      recovered: d.totalRecovered,
      outstanding: d.outstanding,
      count: 0,            // row count not meaningful now; leave 0
      employeeCount: d.employeeCount
    });
  } catch(e){ return _envErr(e); }
}