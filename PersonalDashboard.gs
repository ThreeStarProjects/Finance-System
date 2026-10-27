// ============================================================
// 3 Star Projects — Unified Finance System
// MODULE: Personal Dashboard  (Tier 5.3)
// File: PersonalDashboard.gs   (depends on Transactions.gs — needs Tier 5.1 schema)
//
// Reads Personal_Transactions only. All endpoints return the standard
// {ok, data, error} envelope so Home.html's callServer wrapper works.
//
// KPI rule: Fund In / Fund Out totals count ONLY Class='Fund In' / 'Fund Out'.
// Internal (Personal↔Personal) is visible in the ledger but excluded from
// those totals so they don't double-count owner-to-owner moves.
// ============================================================

function _pAllRows(filters){
  filters = filters || {};
  var rows = (typeof getPersonalTransactions === 'function')
    ? getPersonalTransactions({ fy: filters.fy, includeDeleted: false })
    : [];
  if (filters.owner) rows = rows.filter(function(r){ return r.owner === filters.owner; });
  if (filters.month && filters.month !== 'all')
    rows = rows.filter(function(r){ return r.month === filters.month; });
  return rows;
}

// ── META — dropdown values for the Personal tab
function personalAvailableOwners(){
  try {
    var set = {};
    var rows = (typeof getPersonalTransactions === 'function')
      ? getPersonalTransactions({}) : [];
    rows.forEach(function(r){ if (r.owner) set[r.owner] = true; });
    // Also seed from M_Account owners so a freshly-migrated sheet with no
    // transfers yet still surfaces the owners the user has accounts for.
    try {
      (getAccounts() || []).forEach(function(a){
        if (a.side === 'Personal' && a.owner) set[a.owner] = true;
      });
    } catch(e){}
    var owners = Object.keys(set).sort();
    return _envOk({ owners: owners });
  } catch(e){ return _envErr(e); }
}

// ── KPIs
function personalKpis(filters){
  try {
    var rows = _pAllRows(filters);
    var totalSpent = 0, fundIn = 0, fundOut = 0, internal = 0;
    var catTotals = {};
    var txnCount = 0;
    rows.forEach(function(r){
      txnCount++;
      var amt = _n(r.amount);
      var klass = r.txnClass;
      if (klass === 'Expense'){
        totalSpent += amt;
        var c = r.category || '(uncategorised)';
        catTotals[c] = (catTotals[c] || 0) + amt;
      } else if (klass === 'Fund In'){
        fundIn += amt;
      } else if (klass === 'Fund Out'){
        fundOut += amt;
      } else if (klass === 'Internal'){
        internal += amt;
      }
    });
    var topCat = '', topCatAmt = 0;
    Object.keys(catTotals).forEach(function(c){
      if (catTotals[c] > topCatAmt){ topCat = c; topCatAmt = catTotals[c]; }
    });
    return _envOk({
      fy:     filters && filters.fy || 'all',
      owner:  filters && filters.owner || 'all',
      totalSpent: _round(totalSpent),
      fundIn:     _round(fundIn),
      fundOut:    _round(fundOut),
      internal:   _round(internal),
      netFlow:    _round(fundIn - fundOut),
      topCategory:      topCat,
      topCategoryAmt:   _round(topCatAmt),
      txnCount:         txnCount
    });
  } catch(e){ return _envErr(e); }
}

// ── P&L by category (bar + table)
function personalPnl(filters){
  try {
    var rows = _pAllRows(filters).filter(function(r){ return r.txnClass === 'Expense'; });
    var byCat = {};
    var total = 0;
    rows.forEach(function(r){
      var c = r.category || '(uncategorised)';
      byCat[c] = (byCat[c] || 0) + _n(r.amount);
      total += _n(r.amount);
    });
    var out = Object.keys(byCat).map(function(c){
      return { category: c, amount: _round(byCat[c]),
               pct: total > 0 ? _round(byCat[c]/total*100) : 0 };
    }).sort(function(a,b){ return b.amount - a.amount; });
    return _envOk({ groups: out, total: _round(total) });
  } catch(e){ return _envErr(e); }
}

// ── Month-on-month (Apr→Mar for the FY, chronological)
function personalMonthly(filters){
  try {
    var fy = (filters && filters.fy) || _currentFY();
    var rows = _pAllRows({ fy: fy, owner: filters && filters.owner });
    // Expense per month
    var byMonth = {};
    rows.forEach(function(r){
      if (r.txnClass !== 'Expense') return;
      var m = r.month; if (!m) return;
      byMonth[m] = (byMonth[m] || 0) + _n(r.amount);
    });
    // Build Apr→Mar sequence for the given FY
    var startY = parseInt(String(fy).split('-')[0], 10);
    var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var seq = [3,4,5,6,7,8,9,10,11,0,1,2];
    var out = seq.map(function(mIdx){
      var y = (mIdx >= 3) ? startY : startY + 1;
      var label = MN[mIdx] + '-' + String(y).slice(2);
      return { month: label, amount: _round(byMonth[label] || 0) };
    });
    return _envOk({ months: out });
  } catch(e){ return _envErr(e); }
}

// ── Category donut (alias-ish — same shape as personalPnl but capped to top N)
function personalByCategory(filters){
  try {
    var pnl = personalPnl(filters);
    if (!pnl.ok) return pnl;
    var groups = pnl.data.groups.slice(0, 10);
    return _envOk({ groups: groups, total: pnl.data.total });
  } catch(e){ return _envErr(e); }
}

// ── Top merchants
// Uses Party_Name (Tier 5.1 field). Falls back to the first token of Description
// so pre-migration rows without a party still bucket somewhere.
function personalTopParties(filters, limit){
  try {
    var rows = _pAllRows(filters).filter(function(r){ return r.txnClass === 'Expense'; });
    var byParty = {};
    rows.forEach(function(r){
      var p = (r.party || '').trim();
      if (!p){
        // Fallback: use first 3 words of description as a rough merchant tag
        var d = (r.description || '').trim();
        p = d ? d.split(/\s+/).slice(0, 3).join(' ') : '(unknown)';
      }
      byParty[p] = (byParty[p] || 0) + _n(r.amount);
    });
    var out = Object.keys(byParty).map(function(p){
      return { party: p, amount: _round(byParty[p]) };
    }).sort(function(a,b){ return b.amount - a.amount; })
      .slice(0, _n(limit) || 10);
    return _envOk({ rows: out });
  } catch(e){ return _envErr(e); }
}

// ── Personal ledger table — paged
function personalTransactionsPaged(filters){
  try {
    filters = filters || {};
    var rows = _pAllRows(filters);
    // Text search + class filter
    var q = _s(filters.search).toLowerCase();
    var cls = _s(filters.class);
    if (q){
      rows = rows.filter(function(r){
        return ((r.description || '') + ' ' + (r.category || '') + ' ' + (r.party || '') + ' ' + (r.source || ''))
          .toLowerCase().indexOf(q) !== -1;
      });
    }
    if (cls) rows = rows.filter(function(r){ return r.txnClass === cls; });
    // Newest first
    rows.sort(function(a,b){
      return (_safeDateStr(b.txDate) || '').localeCompare(_safeDateStr(a.txDate) || '');
    });
    var page  = _n(filters.page) || 1;
    var size  = _n(filters.pageSize) || 50;
    var total = rows.length;
    var start = (page - 1) * size;
    var slice = rows.slice(start, start + size).map(function(r){
      return {
        id: r.id,
        date: _safeDateStr(r.txDate),
        month: r.month, fy: r.fy,
        owner: r.owner, txnClass: r.txnClass, direction: r.direction,
        category: r.category, party: r.party, description: r.description,
        amount: r.amount, source: r.source, counterAccount: r.counterAccount,
        paymentMode: r.paymentMode, notes: r.notes, linkedRef: r.linkedRef
      };
    });
    return _envOk({ rows: slice, total: total, page: page, pageSize: size });
  } catch(e){ return _envErr(e); }
}