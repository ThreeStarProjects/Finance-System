// ============================================================
// 3 Star Projects — Unified Finance System
// PARTY DIRECTORIES  (computed 360 rollups)
// File: Directory.gs   (depends on Masters, Transactions, Invoices)
//
// Builds two read-only summary sheets, one row per party:
//   Customer_Directory — orders, invoices, paid, returns, OUTSTANDING,
//                        advance held, last activity, status
//   Vendor_Directory   — # transactions, total spent, last activity, status
//
// Fully COMPUTED each run from Orders/Invoices/Payments/Returns (customers)
// and the Transactions ledger (vendors) — never written back live, so it
// can't drift or corrupt. Run buildDirectories() anytime, or install the
// nightly trigger so it refreshes on its own.
// ============================================================

var DIR = { CUSTOMERS:'Customer_Directory', VENDORS:'Vendor_Directory' };

var CUSTOMER_DIR_HEADERS = ['Customer_ID','Customer_Name','GSTIN','Phone','Email',
  'Orders','Total_Ordered','Invoices','Total_Invoiced','Total_Paid','Returns',
  'Outstanding','Advance_Held','Last_Activity','Status'];
var CUSTOMER_DIR_WIDTHS = [110,230,150,130,220,80,130,80,140,130,120,130,130,130,90];

var VENDOR_DIR_HEADERS = ['Vendor_ID','Vendor_Name','GSTIN','Phone','Email',
  'Transactions','Total_Spent','Last_Activity','Status'];
var VENDOR_DIR_WIDTHS = [110,230,150,130,220,110,140,130,90];

// M_Party column indices (append order in addParty)
var MP = { ID:0, TYPE:1, NAME:2, GSTIN:3, EMAIL:4, PHONE:5, ADDRESS:6, OPENING:7, STATUS:8 };

// ── Entry point ─────────────────────────────────────────────
function buildDirectories(){
  var c = _buildCustomerDirectory();
  var v = _buildVendorDirectory();
  var msg = 'Directories rebuilt:\n• '+DIR.CUSTOMERS+' ('+c+' customers)\n• '+DIR.VENDORS+' ('+v+' vendors)';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, customers:c, vendors:v};
}

function _partyRows(type){
  var sheet = _getSS().getSheetByName(M.PARTY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues()
    .filter(function(r){ return _s(r[MP.TYPE]) === type && _s(r[MP.NAME]); })
    .map(function(r){ return { id:_s(r[MP.ID]), name:_s(r[MP.NAME]), gstin:_s(r[MP.GSTIN]),
      email:_s(r[MP.EMAIL]), phone:_s(r[MP.PHONE]), status:_s(r[MP.STATUS])||'Active' }; });
}
function _maxDate(a,b){ if(!a) return b; if(!b) return a; return (a>b)?a:b; }

// ── Customers ───────────────────────────────────────────────
function _buildCustomerDirectory(){
  var ss = _getSS();
  var agg = {};   // id → rollup
  function slot(id){ return agg[id] || (agg[id] = {orders:0,ordered:0,invoices:0,invoiced:0,
    paid:0,returns:0,outstanding:0,advance:0,last:null,name:''}); }

  var o = ss.getSheetByName(AR.ORDERS);
  if (o && o.getLastRow()>1){
    var oW = Math.min(o.getLastColumn(), OC_TOTAL);
    o.getRange(2,1,o.getLastRow()-1,oW).getValues().forEach(function(r){
      if(!_s(r[OC.CUST_ID])) return; var s=slot(_s(r[OC.CUST_ID])); s.name=s.name||_s(r[OC.CUST_NAME]);
      s.orders++; s.ordered+=_n(r[OC.AMOUNT]); s.last=_maxDate(s.last,_asDate(r[OC.CREATED]));
    });
  }

  var inv = ss.getSheetByName(AR.INVOICES);
  if (inv && inv.getLastRow()>1){
    var iW = Math.min(inv.getLastColumn(), IC_TOTAL);
    inv.getRange(2,1,inv.getLastRow()-1,iW).getValues().forEach(function(r){
      if(!_s(r[IC.CUST_ID])) return; var s=slot(_s(r[IC.CUST_ID])); s.name=s.name||_s(r[IC.CUST_NAME]);
      if(_s(r[IC.STATUS])!==INV_STATUS.CANCELLED){ s.invoices++; s.invoiced+=_n(r[IC.AMOUNT]); s.outstanding+=_n(r[IC.OUTSTANDING]); }
      s.last=_maxDate(s.last,_asDate(r[IC.INV_DATE]));
    });
  }

  var p = ss.getSheetByName(AR.PAYMENTS);
  if (p && p.getLastRow()>1){
    // Payments sheet has grown over time (added TDS + received-in-account
    // columns). Old sheets have fewer physical columns than PMC_TOTAL, so
    // request only what's actually there.
    var pW = Math.min(p.getLastColumn(), PMC_TOTAL);
    p.getRange(2,1,p.getLastRow()-1,pW).getValues().forEach(function(r){
      if(!_s(r[PMC.CUST_ID])) return;
      // Tier 2.1 — skip cancelled payments so the directory rollup doesn't
      // credit money we've since refunded.
      if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
      var s=slot(_s(r[PMC.CUST_ID])); s.name=s.name||_s(r[PMC.CUST_NAME]);
      s.paid+=_n(r[PMC.AMOUNT]); s.advance+=_n(r[PMC.UNAPPLIED]); s.last=_maxDate(s.last,_asDate(r[PMC.DATE]));
    });
  }

  var rt = ss.getSheetByName(AR.RETURNS);
  if (rt && rt.getLastRow()>1){
    var rW = Math.min(rt.getLastColumn(), RC_TOTAL);
    rt.getRange(2,1,rt.getLastRow()-1,rW).getValues().forEach(function(r){
      if(!_s(r[RC.CUST_ID])) return; var s=slot(_s(r[RC.CUST_ID])); s.name=s.name||_s(r[RC.CUST_NAME]);
      s.returns+=_n(r[RC.AMOUNT]); s.last=_maxDate(s.last,_asDate(r[RC.DATE]));
    });
  }

  // master customers first, then any data-only ids
  var masters = _partyRows('Customer'), seen = {}, rows = [];
  masters.forEach(function(m){ seen[m.id]=true; var a=agg[m.id]||{};
    rows.push(_custRow(m.id, m.name, m.gstin, m.phone, m.email, m.status, a)); });
  Object.keys(agg).forEach(function(id){ if(seen[id]) return; var a=agg[id];
    rows.push(_custRow(id, a.name, '', '', '', 'Active', a)); });

  // sort by outstanding desc
  rows.sort(function(x,y){ return (y[11]||0)-(x[11]||0); });
  _writeDirectory(DIR.CUSTOMERS, CUSTOMER_DIR_HEADERS, CUSTOMER_DIR_WIDTHS, rows);
  return rows.length;
}
function _custRow(id,name,gstin,phone,email,status,a){
  a=a||{};
  return [ id, name, gstin, phone, email,
    a.orders||0, _round(a.ordered||0), a.invoices||0, _round(a.invoiced||0),
    _round(a.paid||0), _round(a.returns||0), _round(a.outstanding||0), _round(a.advance||0),
    a.last||'', status||'Active' ];
}

// ── Vendors ─────────────────────────────────────────────────
function _buildVendorDirectory(){
  // Build smart-match lookup from M_Party (Vendors only)
  var vendors = _partyRows('Vendor');
  var byNorm = {};
  vendors.forEach(function(v){ byNorm[_normName(v.name)] = v.id; });
  function fallbackId(name){
    var n = _normName(name); if (!n) return '';
    if (byNorm[n]) return byNorm[n];
    for (var i=0;i<vendors.length;i++){ if (_normMatch(vendors[i].name, name)) return vendors[i].id; }
    return '';
  }

  var agg = {};
  function slot(id){ return agg[id] || (agg[id] = {txns:0,spent:0,last:null,name:''}); }

  // ── Aggregate from Transactions (vendor expenses) ──
  _allTxnSheets().forEach(function(sheet){
    if (sheet.getLastRow()<2) return;
    var w = Math.min(sheet.getLastColumn(), TC_TOTAL);
    sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues().forEach(function(r){
      if (_s(r[TC.DELETED])==='Yes') return;
      if (_s(r[TC.CLASS])!=='Expense') return;
      var id = _s(r[TC.PARTY_ID]);
      var nm = _s(r[TC.PARTY_NAME]);
      // Defensive fallback: row has a name but no Party_ID → try to match
      if (!id && nm) id = fallbackId(nm);
      if (!id) return;
      var s = slot(id); s.name = s.name || nm;
      s.txns++; s.spent += _n(r[TC.AMOUNT]); s.last = _maxDate(s.last, _asDate(r[TC.TX_DATE]));
    });
  });

  // ── Aggregate from Bills (vendor liabilities) ──
  var billsSheet = _getSS().getSheetByName(AP.BILLS);
  if (billsSheet && billsSheet.getLastRow() > 1){
    var w = Math.min(billsSheet.getLastColumn(), BC_TOTAL);
    billsSheet.getRange(2,1,billsSheet.getLastRow()-1,w).getValues().forEach(function(r){
      var id = _s(r[BC.VEND_ID]);
      var nm = _s(r[BC.VEND_NAME]);
      if (!id && nm) id = fallbackId(nm);
      if (!id) return;
      var s = slot(id); s.name = s.name || nm;
      s.spent += _n(r[BC.AMOUNT]);
      s.last = _maxDate(s.last, _asDate(r[BC.BILL_DATE]));
    });
  }

  // ── Aggregate from Bill_Payments (payment dates) ──
  var billPaymentsSheet = _getSS().getSheetByName(AP.PAYMENTS);
  if (billPaymentsSheet && billPaymentsSheet.getLastRow() > 1){
    var w = Math.min(billPaymentsSheet.getLastColumn(), BPMC_TOTAL);
    billPaymentsSheet.getRange(2,1,billPaymentsSheet.getLastRow()-1,w).getValues().forEach(function(r){
      var id = _s(r[BPMC.VEND_ID]);
      if (!id) return;
      // Tier 2.1 — skip cancelled bill payments so "Last Activity" doesn't
      // reflect a payment we've since deleted.
      if (w > BPMC.CANCELLED_ON && _s(r[BPMC.CANCELLED_ON])) return;
      var s = slot(id);
      s.last = _maxDate(s.last, _asDate(r[BPMC.DATE]));
    });
  }

  var seen={}, rows=[];
  vendors.forEach(function(m){ seen[m.id]=true; var a=agg[m.id]||{};
    rows.push([ m.id,m.name,m.gstin,m.phone,m.email, a.txns||0,_round(a.spent||0), a.last||'', m.status||'Active' ]); });
  Object.keys(agg).forEach(function(id){ if(seen[id])return; var a=agg[id];
    rows.push([ id,a.name,'','','', a.txns,_round(a.spent), a.last||'', 'Active' ]); });

  rows.sort(function(x,y){ return (y[6]||0)-(x[6]||0); });
  _writeDirectory(DIR.VENDORS, VENDOR_DIR_HEADERS, VENDOR_DIR_WIDTHS, rows);
  return rows.length;
}

// ── Write helper: (re)build the sheet, clear old data, write fresh ──
function _writeDirectory(name, headers, widths, rows){
  var ss = _getSS();
  var sheet = _buildSheet(ss, name, headers, widths);   // styled via _beautifySheet
  if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).clearContent();
  if (rows.length) sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
  // move directory tabs to the front for visibility
  try { ss.setActiveSheet(sheet); ss.moveActiveSheet(1); } catch(e){}
  return sheet;
}

// ── Directory auto-refresh trigger (hourly) ─────────────────
function installHourlyDirectoryTrigger(){
  removeDirectoryTrigger();
  ScriptApp.newTrigger('buildDirectories').timeBased().everyHours(1).create();
  var msg = 'Hourly directory refresh installed (runs every hour).';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
}

function removeDirectoryTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'buildDirectories') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// dedupeParties()  —  M_PARTY MAINTENANCE
// ============================================================
// Groups M_Party rows by smart-normalised name + type. For each group
// with >1 row: keeps the row with the lowest-numbered ID, retargets
// every transaction/AR row pointing at the duplicate IDs to point at
// the keeper, and deletes the duplicate M_Party rows.
//
// Run any time you suspect duplicates have crept in. Safe to re-run.
// ============================================================
function dedupeParties(){
  var ss = _getSS();
  var party = ss.getSheetByName(M.PARTY);
  if (!party || party.getLastRow() < 2){
    var m = 'M_Party is empty.';
    try { SpreadsheetApp.getUi().alert(m); } catch(e){ Logger.log(m); }
    return {merged:0};
  }
  var last = party.getLastRow(), nCols = party.getLastColumn();
  var rows  = party.getRange(2, 1, last - 1, nCols).getValues();

  // Group by (type + normalised name)
  var groups = {};
  rows.forEach(function(r, i){
    var id = _s(r[0]), type = _s(r[1]), nm = _s(r[2]);
    if (!id || !nm) return;
    var key = type + '|' + _normName(nm);
    if (!groups[key]) groups[key] = [];
    groups[key].push({ id:id, name:nm, rowIdx:i, num:_numOf(id) });
  });

  var remap = {}, dropRows = [], mergeReport = [];
  Object.keys(groups).forEach(function(key){
    var g = groups[key];
    if (g.length < 2) return;
    g.sort(function(a,b){ return a.num - b.num; });
    var keeper = g[0], dups = g.slice(1);
    dups.forEach(function(d){
      remap[d.id] = keeper.id;
      dropRows.push(d.rowIdx + 2);
    });
    mergeReport.push(keeper.name+' ('+keeper.id+') ← '+dups.map(function(d){return d.id;}).join(', '));
  });

  // Retarget all sheets that carry Party_ID references
  var retargetResult = _retargetPartyIds(remap, ss);
  var totalRetargeted = retargetResult.total;
  var sheetRetarget = retargetResult.perSheet;

  // Delete duplicate M_Party rows (descending so indices stay valid)
  dropRows.sort(function(a,b){ return b - a; }).forEach(function(rn){ party.deleteRow(rn); });

  // Build report
  var msg = 'Deduped '+Object.keys(remap).length+' duplicate party row(s).';
  if (mergeReport.length){
    msg += '\n\nMerges:\n• '+mergeReport.slice(0,25).join('\n• ');
    if (mergeReport.length > 25) msg += '\n• …and '+(mergeReport.length-25)+' more';
  }
  if (sheetRetarget.length){
    msg += '\n\nTransaction rows retargeted: '+totalRetargeted+'\n• '+sheetRetarget.join('\n• ');
  }
  if (Object.keys(remap).length){
    try { buildDirectories(); msg += '\n\n✅ Directories refreshed.'; }
    catch(e){ msg += '\n⚠ buildDirectories failed: '+e.message; }
  } else {
    msg = 'No duplicates found in M_Party. All clean.';
  }
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return { merged: Object.keys(remap).length, retargeted: totalRetargeted };
}

// Shared retarget helper: rewrites every Party_ID cell whose value is a key
// of `remap` to the mapped target ID. Also refreshes the Vendor_Name /
// Customer_Name / Party_Name in the same row so the display column stays in
// sync. Returns {total, perSheet}.
function _retargetPartyIds(remap, ss){
  ss = ss || _getSS();
  // Build id → target-name lookup once from M_Party.
  var party = ss.getSheetByName(M.PARTY);
  var idToName = {};
  if (party && party.getLastRow() > 1){
    var prows = party.getRange(2,1,party.getLastRow()-1,3).getValues();
    prows.forEach(function(r){ idToName[_s(r[0])] = _s(r[2]); });
  }
  var targets = _allTxnSheets();
  ['Invoices','Payments','Returns','Orders','Bills','Bill_Payments'].forEach(function(n){
    var sh = ss.getSheetByName(n); if (sh) targets.push(sh);
  });
  var total = 0, perSheet = [];
  targets.forEach(function(sh){
    if (!sh || sh.getLastRow() < 2) return;
    var info = _findPartyCols(sh);
    if (!info) return;
    var n = sh.getLastRow() - 1;
    var ids   = sh.getRange(2, info.idCol,   n, 1).getValues();
    var names = sh.getRange(2, info.nameCol, n, 1).getValues();
    var changed = 0;
    for (var r = 0; r < ids.length; r++){
      var id = _s(ids[r][0]);
      if (id && remap[id]){
        ids[r][0]   = remap[id];
        var newNm   = idToName[remap[id]];
        if (newNm) names[r][0] = newNm;
        changed++;
      }
    }
    if (changed){
      sh.getRange(2, info.idCol,   n, 1).setValues(ids);
      sh.getRange(2, info.nameCol, n, 1).setValues(names);
      perSheet.push(sh.getName()+': '+changed);
      total += changed;
    }
  });
  return {total: total, perSheet: perSheet};
}

// mergePartyPair(sourceId, targetId) — manual merge, no name matching.
// Moves every Party_ID reference from sourceId to targetId, then removes
// the source row from M_Party. Guardrails: both IDs must exist, same type
// (both Vendor or both Customer), and cannot be the same.
function mergePartyPair(sourceId, targetId){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    sourceId = _s(sourceId); targetId = _s(targetId);
    if (!sourceId || !targetId) return {success:false, message:'Source and target party required.'};
    if (sourceId === targetId)  return {success:false, message:'Source and target must be different.'};
    var ss = _getSS(), party = ss.getSheetByName(M.PARTY);
    if (!party || party.getLastRow() < 2) return {success:false, message:'M_Party is empty.'};
    var rows = party.getRange(2,1,party.getLastRow()-1, party.getLastColumn()).getValues();
    var srcRow = -1, tgtRow = -1, srcType = '', tgtType = '', srcName = '', tgtName = '';
    for (var i=0;i<rows.length;i++){
      var id = _s(rows[i][0]);
      if (id === sourceId){ srcRow = i+2; srcType = _s(rows[i][1]); srcName = _s(rows[i][2]); }
      if (id === targetId){ tgtRow = i+2; tgtType = _s(rows[i][1]); tgtName = _s(rows[i][2]); }
    }
    if (srcRow < 0) return {success:false, message:'Source party ' + sourceId + ' not found.'};
    if (tgtRow < 0) return {success:false, message:'Target party ' + targetId + ' not found.'};
    if (srcType !== tgtType) return {success:false, message:'Cannot merge across types (' + srcType + ' → ' + tgtType + ').'};

    var remap = {}; remap[sourceId] = targetId;
    var res = _retargetPartyIds(remap, ss);
    party.deleteRow(srcRow);
    try { buildDirectories(); } catch(e){}
    return {
      success: true,
      merged: sourceId + ' (' + srcName + ') → ' + targetId + ' (' + tgtName + ')',
      retargeted: res.total,
      perSheet: res.perSheet,
      message: 'Merged ' + srcName + ' into ' + tgtName + '. ' + res.total + ' row(s) retargeted across ' + res.perSheet.length + ' sheet(s).'
    };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Helper: trailing numeric suffix of an ID (e.g. 'VEND-0023' → 23)
function _numOf(id){
  var m = String(id).match(/(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : 999999;
}

// Helper: locate Party_ID + Party_Name columns on any sheet.
function _findPartyCols(sheet){
  var name = sheet.getName();
  if (name.indexOf('Transactions_') === 0){
    return { idCol: TC.PARTY_ID + 1, nameCol: TC.PARTY_NAME + 1 };
  }
  var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idCol = 0, nameCol = 0;
  for (var i = 0; i < hdr.length; i++){
    var h = _s(hdr[i]);
    if (!idCol   && (h === 'Customer_ID'   || h === 'Vendor_ID'   || h === 'Party_ID'))   idCol   = i + 1;
    if (!nameCol && (h === 'Customer_Name' || h === 'Vendor_Name' || h === 'Party_Name')) nameCol = i + 1;
  }
  return (idCol && nameCol) ? { idCol: idCol, nameCol: nameCol } : null;
}
// hardDeletePartyCascade(partyId, reason)
// Deletes a Customer or Vendor and every row that references them:
//   • Bills + Bill_Payments (for vendors)
//   • Invoices + Payments (for customers) — payment allocations are just gone
//     since the whole invoice is gone; nothing to reverse.
//   • Transactions_<FY> rows tagged with this party.
// Removes the M_Party row last. Also rebuilds directories.
// IRREVERSIBLE — every soft-delete audit column is bypassed. Only call this
// when the user has explicitly typed a CONFIRM string on the frontend.
function hardDeletePartyCascade(partyId, reason){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    partyId = _s(partyId);
    if (!partyId) return {success:false, message:'Party ID required.'};
    var ss = _getSS();
    var party = ss.getSheetByName(M.PARTY);
    if (!party || party.getLastRow() < 2) return {success:false, message:'M_Party is empty.'};
    var pRows = party.getRange(2,1,party.getLastRow()-1, party.getLastColumn()).getValues();
    var partyRow = -1, partyType = '', partyName = '';
    for (var i=0;i<pRows.length;i++){
      if (_s(pRows[i][0]) === partyId){ partyRow = i+2; partyType = _s(pRows[i][1]); partyName = _s(pRows[i][2]); break; }
    }
    if (partyRow < 0) return {success:false, message:'Party ' + partyId + ' not found.'};

    var deleted = { bills:0, billPayments:0, invoices:0, customerPayments:0, transactions:0 };

    // Iterate target sheets and delete rows matching this party ID.
    // Delete in descending row order so indices remain valid.
    var targets = _allTxnSheets();
    ['Invoices','Payments','Bills','Bill_Payments'].forEach(function(n){
      var sh = ss.getSheetByName(n); if (sh) targets.push(sh);
    });
    targets.forEach(function(sh){
      if (!sh || sh.getLastRow() < 2) return;
      var info = _findPartyCols(sh);
      if (!info) return;
      var n = sh.getLastRow() - 1;
      var ids = sh.getRange(2, info.idCol, n, 1).getValues();
      var toDelete = [];
      for (var r = 0; r < ids.length; r++){
        if (_s(ids[r][0]) === partyId) toDelete.push(r + 2);
      }
      // Delete high→low so row indices don't shift underneath us.
      toDelete.sort(function(a,b){ return b - a; }).forEach(function(rn){ sh.deleteRow(rn); });
      // Attribute counts by sheet name for the report.
      var nm = sh.getName();
      if (nm === 'Bills')          deleted.bills            += toDelete.length;
      else if (nm === 'Bill_Payments') deleted.billPayments += toDelete.length;
      else if (nm === 'Invoices')  deleted.invoices         += toDelete.length;
      else if (nm === 'Payments')  deleted.customerPayments += toDelete.length;
      else if (nm.indexOf('Transactions_') === 0) deleted.transactions += toDelete.length;
    });

    // Finally, remove the M_Party row itself.
    party.deleteRow(partyRow);
    try { buildDirectories(); } catch(e){}

    var msg = 'Deleted ' + partyType + ' ' + partyName + ' (' + partyId + '). Rows removed: ' +
              deleted.bills + ' bills, ' + deleted.billPayments + ' bill payments, ' +
              deleted.invoices + ' invoices, ' + deleted.customerPayments + ' customer payments, ' +
              deleted.transactions + ' transactions.' + (reason ? ' Reason: ' + reason : '');
    return { success:true, partyId:partyId, partyType:partyType, partyName:partyName, deleted:deleted, message:msg };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// hardDeleteProjectUntag(projectName) — sets Project='' on every Bills /
// Invoices / Payments / Bill_Payments / Transactions_<FY> row currently
// tagged with this project, then removes the project from M_Project.
// Non-destructive to financial data — only untags. Re-runs applyProjectDropdowns.
function hardDeleteProjectUntag(projectName){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    projectName = _s(projectName);
    if (!projectName) return {success:false, message:'Project name required.'};
    var ss = _getSS();
    var untagged = { bills:0, invoices:0, payments:0, billPayments:0, transactions:0 };

    function untagSheet(sh, projectCol0Based, key){
      if (!sh || sh.getLastRow() < 2) return;
      if (sh.getLastColumn() <= projectCol0Based) return;
      var n = sh.getLastRow() - 1;
      var col = projectCol0Based + 1;
      var vals = sh.getRange(2, col, n, 1).getValues();
      var changed = 0;
      for (var r = 0; r < vals.length; r++){
        if (_s(vals[r][0]) === projectName){ vals[r][0] = ''; changed++; }
      }
      if (changed){
        sh.getRange(2, col, n, 1).setValues(vals);
        untagged[key] += changed;
      }
    }

    // Column indices: Bills.PROJECT=21, Invoices.PROJECT=22, Payments.PROJECT (PMC.PROJECT), Bill_Payments (BPMC.PROJECT)
    untagSheet(ss.getSheetByName('Bills'),         21, 'bills');
    untagSheet(ss.getSheetByName('Invoices'),      22, 'invoices');
    untagSheet(ss.getSheetByName('Payments'),      (typeof PMC !== 'undefined' ? PMC.PROJECT : 20), 'payments');
    untagSheet(ss.getSheetByName('Bill_Payments'), (typeof BPMC !== 'undefined' ? BPMC.PROJECT : 19), 'billPayments');
    _allTxnSheets().forEach(function(sh){
      untagSheet(sh, TC.PROJECT, 'transactions');
    });

    // Remove from M_Project via saveProjectMaster.
    var projs = getProjectMasterData();
    var kept = projs.filter(function(p){ return _s(p.name) !== projectName; });
    if (kept.length === projs.length){
      // Not found in M_Project — untag may still have run.
      return {success:true, untagged:untagged, message:'Untagged ' + JSON.stringify(untagged) + '. Project was not in M_Project (already removed).'};
    }
    var saveRes = saveProjectMaster(kept);
    if (!saveRes.success) return {success:false, message:'Untag ok but M_Project save failed: ' + saveRes.message};
    return { success:true, projectName:projectName, untagged:untagged, message:'Project "' + projectName + '" deleted. Untagged: ' + untagged.bills + ' bills, ' + untagged.invoices + ' invoices, ' + untagged.payments + ' customer payments, ' + untagged.billPayments + ' vendor payments, ' + untagged.transactions + ' transactions.' };
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}