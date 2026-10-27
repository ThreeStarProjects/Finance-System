// ============================================================
// 3 Star Projects — Unified Finance System
// PHASE 1 : MASTERS
// File: Masters.gs
//
// Creates + seeds the four backbone master sheets:
//   M_Category           — business expense categories → P&L group (+ GST-ready cols)
//   M_Category_Personal  — personal-spend categories (separate ledger)
//   M_Party              — customers + vendors, auto IDs CUST-#### / VEND-####
//   M_Config             — typed lists: payment modes, source-of-funds, statuses
//
// Entry point:  run  setupMasters()  once from the Apps Script editor.
// Safe to re-run: headers/seed are idempotent, existing data is NOT wiped.
// ============================================================

// ── Sheet names (one place to rename) ───────────────────────
var M = {
  CATEGORY:          'M_Category',
  CATEGORY_PERSONAL: 'M_Category_Personal',
  PARTY:             'M_Party',
  CONFIG:            'M_Config',
  ACCOUNT:           'M_Account'
};

// ── Brand palette (matches the 3SP web app design tokens) ───
var BRAND = {
  headerBg:   '#0F172A',   // dark slate header row
  headerFg:   '#FFFFFF',
  primary:    '#1F4E79',
  accent:     '#2C7DA0'
};

// The 14 canonical P&L groups (drives the forced-group picker).
// Order = the order they appear on the P&L MIS report.
var PNL_GROUPS = [
  'Direct Production',
  'Logistics & Transport',
  'Salary & Wages',
  'Admin & Office',
  'Utilities',
  'Sales & Marketing',
  'Professional Services',
  'Bank, Finance & Loans',
  'Taxes & Compliance',
  'Income'
];

// ============================================================
// SPREADSHEET BINDING  (self-healing — bound or standalone)
// ============================================================
function _getSS(){
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e){}
  if (ss) return ss;
  var stored = null;
  try { stored = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); } catch(e){}
  if (stored){
    try { return SpreadsheetApp.openById(stored); }
    catch(e){ try { PropertiesService.getScriptProperties().deleteProperty('SPREADSHEET_ID'); } catch(_){} }
  }
  var newSs = SpreadsheetApp.create('3 Star Projects — Unified Finance');
  try { PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', newSs.getId()); } catch(e){}
  Logger.log('Created new spreadsheet: ' + newSs.getUrl());
  return newSs;
}

// ── Safe coercions (proven pattern, never throw) ────────────
function _s(v){
  try {
    if (v === null || v === undefined) return '';
    if (v instanceof Date){ return isNaN(v.getTime()) ? '' : Utilities.formatDate(v,'GMT+5:30','yyyy-MM-dd'); }
    return String(v).trim();
  } catch(e){ return ''; }
}
function _n(v){
  try { if (v === null || v === undefined || v === '') return 0; var n = parseFloat(v); return isNaN(n) ? 0 : n; }
  catch(e){ return 0; }
}

// ── Sheet builder: create if missing, set + style headers ───
function _buildSheet(ss, name, headers, widths){
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  // Grow the sheet's physical column count BEFORE writing headers. An
  // existing sheet created when BILL_HEADERS was narrower (23 cols) sits at
  // getMaxColumns() === 23, and the later setValues(1,1,1,25) silently
  // truncates to the existing width — new headers never appear.
  var maxCols = sheet.getMaxColumns();
  if (maxCols < headers.length){
    sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
  }
  // Header row (only (re)write if first cell empty or header changed length)
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(1,1,1,headers.length)
       .setBackground(BRAND.headerBg).setFontColor(BRAND.headerFg)
       .setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);
  if (widths) widths.forEach(function(w,i){ sheet.setColumnWidth(i+1, w); });
  try { if (typeof _beautifySheet === 'function') _beautifySheet(sheet); } catch(e){}
  return sheet;
}

// Apply a dropdown to a column from row 2 down.
function _validateCol(sheet, col, values){
  sheet.getRange(2, col, 2000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build()
  );
}

// Has this sheet got any data rows yet?
function _hasData(sheet){ return sheet.getLastRow() > 1; }

// ─────────────────────────────────────────────────────────────
// _normName(s) — smart party-name normaliser
// Used everywhere parties are matched: addParty (dup guard),
// findOrCreateParty (lookup), processBulkPaste (Party_ID backfill),
// dedupeParties (merge groups), Vendor_Directory (fallback match).
//
// What it normalises:
//   • lowercase
//   • whitespace collapse
//   • strips legal-form suffixes: pvt, pvt. ltd, private limited, llp, ltd, co.
//   • strips trailing punctuation (. , - / & ! ?)
//   • strips truncation markers like "PRIV" or "FURNITUR" when paired
//     with a longer canonical match (handled via prefix logic in callers)
// So these all normalise to the same string:
//   "JALAN BROTHERS PRIVATE"  →  "jalan brothers"
//   "Jalan Brothers Pvt Ltd"  →  "jalan brothers"
//   "jalan brothers"          →  "jalan brothers"
// ─────────────────────────────────────────────────────────────
function _normName(s){
  if (s === null || s === undefined) return '';
  var t = String(s).toLowerCase();
  // strip accents/diacritics (defensive)
  try { t = t.normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(e){}
  // remove legal-form suffixes (order matters: longest first)
  var SUFFIXES = [
    'private limited', 'private ltd', 'pvt limited', 'pvt ltd', 'pvt. ltd',
    'pvt. ltd.', 'p ltd', 'pvt', 'llp', 'l.l.p', 'limited', 'ltd',
    'and company', 'and co', '& co', '& sons', 'and sons',
    'industries', 'enterprise', 'enterprises', 'traders', 'trading',
    'corporation', 'corp', 'inc', 'co.'
  ];
  // also strip trailing "the " at start
  t = t.replace(/^the\s+/, '');
  // remove all punctuation except letters/digits/spaces
  t = t.replace(/[^a-z0-9\s]/g, ' ');
  // collapse whitespace
  t = t.replace(/\s+/g, ' ').trim();
  // strip suffixes (must be at end, word-bounded)
  var changed = true;
  while (changed){
    changed = false;
    for (var i = 0; i < SUFFIXES.length; i++){
      var suf = SUFFIXES[i];
      if (t === suf) { t = ''; changed = true; break; }
      if (t.length > suf.length + 1 && t.slice(-suf.length - 1) === ' ' + suf){
        t = t.slice(0, -suf.length - 1).trim();
        changed = true;
        break;
      }
    }
  }
  // Collapse runs of 2+ single letters into one token, so "a m wood plaza"
  // matches "am wood plaza". Catches "AM"/"A M", "SBI"/"S B I", "IBM"/"I B M"
  // typos. Iterates because a single pass only merges the first pair.
  // Single lone letter followed by a real word (e.g. "a wood") is untouched —
  // the [a-z]\b in the second group requires the next token to also be one letter.
  var prev;
  do {
    prev = t;
    t = t.replace(/\b([a-z])\s+([a-z]\b)/, '$1$2');
  } while (prev !== t);
  return t;
}

// _normMatch(a, b) — fuzzy-equal with prefix-collapse fallback.
// Returns true if normalised a===b, OR if one is a prefix of the other
// AND the shorter is at least 6 chars (avoids "MS" matching everything).
// This catches truncations like "Impression Furnitur" vs "Impression Furniture".
function _normMatch(a, b){
  var na = _normName(a), nb = _normName(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  if (na.length >= 6 && nb.length >= 6){
    if (nb.indexOf(na) === 0) return true;   // na is a prefix of nb
    if (na.indexOf(nb) === 0) return true;   // nb is a prefix of na
  }
  return false;
}

// ============================================================
// ENTRY POINT
// ============================================================
function setupMasters(){
  var ss = _getSS();
  setupCategoryMaster(ss);
  setupPersonalCategoryMaster(ss);
  setupPartyMaster(ss);
  setupConfigMaster(ss);
  setupAccountMaster(ss);
  var billsInfo = '';
  try {
    var r = setupBillsModule();
    // Hard repair: ensure the Bills sheet carries Description (col 24) + Group
    // (col 25) regardless of what Bills.gs's BILL_HEADERS declares. Grows the
    // sheet, writes the two header cells explicitly, then verifies by reading
    // back. This bypasses _buildSheet's silent-truncate failure mode.
    _repairBillsSchemaForDescGroup();
    var bSheet = ss.getSheetByName('Bills');
    var cols   = bSheet ? bSheet.getLastColumn() : 0;
    var h24 = bSheet ? _s(bSheet.getRange(1, 24).getValue()) : '';
    var h25 = bSheet ? _s(bSheet.getRange(1, 25).getValue()) : '';
    billsInfo = '\n• ' + ((r && r.bills) || 'Bills') +
                ' (' + cols + ' cols; col24=' + (h24 || '(blank)') + ', col25=' + (h25 || '(blank)') + ')' +
                '\n• ' + ((r && r.payments) || 'Bill_Payments') +
                '\n• ' + ((r && r.allocations) || 'Bill_Allocations');
  } catch(e){
    Logger.log('setupBillsModule / repair: '+e.message);
    billsInfo = '\n• Bills setup FAILED: ' + e.message;
  }
  try { setupMonthlyClose(); } catch(e){ Logger.log('setupMonthlyClose: '+e.message); }
  try { if (typeof setupSalaryAdvance === 'function') setupSalaryAdvance(); } catch(e){ Logger.log('setupSalaryAdvance: '+e.message); }
  try { ensurePriorFyTransactionSheets(['2024-25', '2025-26']); } catch(e){ Logger.log('ensurePriorFyTransactionSheets: '+e.message); }

  var msg = 'Masters ready:\n• ' + M.CATEGORY + '\n• ' + M.CATEGORY_PERSONAL +
            '\n• ' + M.PARTY + '\n• ' + M.CONFIG + '\n• ' + M.ACCOUNT +
            billsInfo +
            '\n• M_MonthlyClose' +
            '\n• Transactions_2024-25, Transactions_2025-26' +
            '\n\nSeeded categories, personal categories, payment modes, source-of-funds, accounts.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return { success: true, message: msg };
}

// ============================================================
// 1. M_Category  (business expense categories)
//    Cols: Category | Group | Direction | GST_Applicable | GST_Rate | HSN_SAC | Active
//    Direction = OUT (expense). Income categories added in a later phase.
//    GST cols are READY but blank — owner fills rates/HSN when needed.
// ============================================================
// Canonical category taxonomy. Format: [name, group, direction, employeeVisible]
//   - direction: 'OUT' for expense/drawing, 'IN' for income
//   - employeeVisible: true means the category appears in the field-staff
//     expense form. Owner form always shows everything.
//   - Every group has an "Other (Specify)" row to absorb the long tail;
//     when chosen, Notes is required on the form.
// ============================================================
var CATEGORY_SEED = [
  // ── Direct Production ─────────────────────────────────────
  ['Raw Material Purchase',                  'Direct Production',     'OUT', true],
  ['Parts / Components Purchase',            'Direct Production',     'OUT', true],
  ['Consumables (fasteners, packaging)',     'Direct Production',     'OUT', true],
  ['Vendor Parts (Wheels, Seats, etc.)',     'Direct Production',     'OUT', false],
  ['Packing Material',                       'Direct Production',     'OUT', true],
  ['Tooling & Drill Bits',                   'Direct Production',     'OUT', true],
  ['Job Work Charges',                       'Direct Production',     'OUT', false],
  ['Heavy Labour / Loading-Unloading',       'Direct Production',     'OUT', true],
  ['Daily Labour Charges',                   'Direct Production',     'OUT', true],
  ['Die / Mould Development & Repair',       'Direct Production',     'OUT', false],
  ['Machinery Purchase',                     'Direct Production',     'OUT', false],
  ['Machinery Repair & Maintenance',         'Direct Production',     'OUT', true],
  ['Plant Repair & Maintenance',             'Direct Production',     'OUT', true],
  ['Raw Material Testing',                   'Direct Production',     'OUT', false],
  ['Other Production (Specify)',             'Direct Production',     'OUT', true],

  // ── Logistics & Transport ─────────────────────────────────
  ['Freight & Forwarding (outbound)',        'Logistics & Transport', 'OUT', true],
  ['Courier Charges',                        'Logistics & Transport', 'OUT', true],
  ['Inbound Transport (material in)',        'Logistics & Transport', 'OUT', true],
  ['Petrol / Diesel (Vehicle Fuel)',         'Logistics & Transport', 'OUT', true],
  ['Vehicle Running & Maintenance',          'Logistics & Transport', 'OUT', true],
  ['Vehicle Purchase / Loan / Rent',         'Logistics & Transport', 'OUT', false],
  ['Toll, Parking, RTO',                     'Logistics & Transport', 'OUT', true],
  ['Employee Travel (Bus/Train/Auto)',       'Logistics & Transport', 'OUT', true],
  ['Other Transport (Specify)',              'Logistics & Transport', 'OUT', true],

  // ── Salary & Wages ────────────────────────────────────────
  ['Employee Monthly Salary',                'Salary & Wages',        'OUT', false],
  ["Director's Salary",                      'Salary & Wages',        'OUT', false],
  ['Bonus',                                  'Salary & Wages',        'OUT', false],
  ['Salary Advance',                         'Salary & Wages',        'OUT', false],
  ['Leave Encashment',                       'Salary & Wages',        'OUT', false],
  ['PF (Provident Fund)',                    'Salary & Wages',        'OUT', false],
  ['ESIC',                                   'Salary & Wages',        'OUT', false],
  ['Gratuity',                               'Salary & Wages',        'OUT', false],
  ['Employee Welfare',                       'Salary & Wages',        'OUT', true],
  ['Employee Medical',                       'Salary & Wages',        'OUT', true],
  ['Employee Insurance',                     'Salary & Wages',        'OUT', false],
  ['Other Salary (Specify)',                 'Salary & Wages',        'OUT', false],

  // ── Admin & Office ────────────────────────────────────────
  ['Rent Paid (office/factory)',             'Admin & Office',        'OUT', false],
  ['Office Supplies / Misc',                 'Admin & Office',        'OUT', true],
  ['Stationery & Printing',                  'Admin & Office',        'OUT', true],
  ['Housekeeping (supplies & services)',     'Admin & Office',        'OUT', true],
  ['Tea, Coffee, Kitchen Supplies',          'Admin & Office',        'OUT', true],
  ['Office Food / Lunch / Dinner',           'Admin & Office',        'OUT', true],
  ['Mobile / Telephone Bills',               'Admin & Office',        'OUT', false],
  ['Internet / Wi-Fi',                       'Admin & Office',        'OUT', false],
  ['IT Purchase / Repair / Maintenance',     'Admin & Office',        'OUT', true],
  ['Software Subscriptions',                 'Admin & Office',        'OUT', false],
  ['Security Services',                      'Admin & Office',        'OUT', false],
  ['Office Repair & Maintenance',            'Admin & Office',        'OUT', true],
  ['Property Insurance',                     'Admin & Office',        'OUT', false],
  ['Factory / Office Insurance',             'Admin & Office',        'OUT', false],
  ['Other Admin (Specify)',                  'Admin & Office',        'OUT', true],

  // ── Utilities ─────────────────────────────────────────────
  ['Electricity Bill',                       'Utilities',             'OUT', false],
  ['Water Supply (Bill)',                    'Utilities',             'OUT', false],
  ['Water Tanker',                           'Utilities',             'OUT', true],
  ['PNG / Diesel (Industrial Use)',          'Utilities',             'OUT', false],
  ['CO2 / O2 / Argon Cylinders',             'Utilities',             'OUT', true],
  ['Property Tax',                           'Utilities',             'OUT', false],
  ['Other Utilities (Specify)',              'Utilities',             'OUT', false],

  // ── Sales & Marketing ─────────────────────────────────────
  ['Marketing / Advertising',                'Sales & Marketing',     'OUT', false],
  ['Expo / Exhibition Costs',                'Sales & Marketing',     'OUT', false],
  ['Client Entertainment',                   'Sales & Marketing',     'OUT', true],
  ['Site Visit Expenses',                    'Sales & Marketing',     'OUT', true],
  ['Sales Tour Travel (multi-day)',          'Sales & Marketing',     'OUT', false],
  ['Commission Paid',                        'Sales & Marketing',     'OUT', false],
  ['Website / Domain / SEO',                 'Sales & Marketing',     'OUT', false],
  ['Other Sales (Specify)',                  'Sales & Marketing',     'OUT', true],

  // ── Professional Services ─────────────────────────────────
  ['CA / Accountant Fees',                   'Professional Services', 'OUT', false],
  ['Legal Fees',                             'Professional Services', 'OUT', false],
  ['Consultancy Charges (general)',          'Professional Services', 'OUT', false],
  ['R&D Expense',                            'Professional Services', 'OUT', false],
  ['Lab / Material Testing',                 'Professional Services', 'OUT', false],
  ['Other Professional (Specify)',           'Professional Services', 'OUT', false],

  // ── Bank, Finance & Loans ─────────────────────────────────
  ['Bank Charges',                           'Bank, Finance & Loans', 'OUT', false],
  ['Bank Interest Paid',                     'Bank, Finance & Loans', 'OUT', false],
  ['OD / CC Utilization Cost',               'Bank, Finance & Loans', 'OUT', false],
  ['Loan Repayment — Principal',             'Bank, Finance & Loans', 'OUT', false],
  ['Loan Repayment — Interest',              'Bank, Finance & Loans', 'OUT', false],
  ['Advance Given to Vendor',                'Bank, Finance & Loans', 'OUT', false],
  ['Other Finance (Specify)',                'Bank, Finance & Loans', 'OUT', false],

  // ── Taxes & Compliance ────────────────────────────────────
  ['GST Paid',                               'Taxes & Compliance',    'OUT', false],
  ['TDS Paid',                               'Taxes & Compliance',    'OUT', false],
  ['Professional Tax',                       'Taxes & Compliance',    'OUT', false],
  ['Income Tax / Advance Tax',               'Taxes & Compliance',    'OUT', false],
  ['License / Renewal Fees',                 'Taxes & Compliance',    'OUT', false],
  ['Late Fee / Penalty (Statutory)',         'Taxes & Compliance',    'OUT', false],
  ['Other Tax (Specify)',                    'Taxes & Compliance',    'OUT', false],

  // ── Income ────────────────────────────────────────────────
  ['Sales — Furniture',                      'Income',                'IN',  false],
  ['Sales — Interior Execution / Project',   'Income',                'IN',  false],
  ['Service Income',                         'Income',                'IN',  false],
  ['Discount Received',                      'Income',                'IN',  false],
  ['Interest Income',                        'Income',                'IN',  false],
  ['Other Income',                           'Income',                'IN',  false]
];

function setupCategoryMaster(ss){
  ss = ss || _getSS();   // self-heal when called directly from the Apps Script editor
  var sheet = _buildSheet(ss, M.CATEGORY,
    ['Category','Group','Direction','GST_Applicable','GST_Rate','HSN_SAC','Active','Employee_Visible'],
    [320, 180, 90, 120, 90, 120, 70, 130]);

  // Dropdowns
  _validateCol(sheet, 2, PNL_GROUPS);                 // Group
  _validateCol(sheet, 3, ['OUT','IN']);               // Direction
  _validateCol(sheet, 4, ['Yes','No']);               // GST_Applicable
  _validateCol(sheet, 7, ['Yes','No']);               // Active
  _validateCol(sheet, 8, ['Yes','No']);               // Employee_Visible

  // Seed only if empty (don't clobber edits on re-run)
  if (!_hasData(sheet)){
    var rows = CATEGORY_SEED.map(function(p){
      // p = [name, group, direction, employeeVisible]
      return [p[0], p[1], p[2] || 'OUT', 'No', '', '', 'Yes', (p[3] ? 'Yes' : 'No')];
    });
    sheet.getRange(2,1,rows.length,8).setValues(rows);
  } else {
    // Schema upgrade path: if the sheet already exists from an older
    // version (7 columns instead of 8), add the new Employee_Visible
    // column and back-fill default 'No' so behaviour stays conservative.
    if (sheet.getLastColumn() < 8){
      sheet.getRange(1, 8).setValue('Employee_Visible').setFontWeight('bold');
      sheet.setColumnWidth(8, 130);
      _validateCol(sheet, 8, ['Yes','No']);
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2){
        var defaults = new Array(lastRow - 1).fill(['No']);
        sheet.getRange(2, 8, lastRow - 1, 1).setValues(defaults);
      }
    }
  }
  return sheet;
}

// ============================================================
// 2. M_Category_Personal  (separate personal ledger categories)
//    Cols: Category | Active
//    Seed is a starter set — owner refines via the form's "Other".
// ============================================================
var PERSONAL_CATEGORY_SEED = [
  'Rent / Home Loan','School / Education Fees','Groceries','Medical / Health',
  'Home Utilities (Electricity/Gas/Water)','Personal Travel','EMI / Loan',
  'Insurance / LIC','Personal Fuel','Dining Out','Shopping','Entertainment',
  'Domestic Help','Family Support','Donations','Investments / Savings'
];

function setupPersonalCategoryMaster(ss){
  ss = ss || _getSS();   // self-heal when called directly from the Apps Script editor
  var sheet = _buildSheet(ss, M.CATEGORY_PERSONAL, ['Category','Active'], [320, 70]);
  _validateCol(sheet, 2, ['Yes','No']);
  if (!_hasData(sheet)){
    var rows = PERSONAL_CATEGORY_SEED.map(function(c){ return [c, 'Yes']; });
    sheet.getRange(2,1,rows.length,2).setValues(rows);
  }
  return sheet;
}

// ============================================================
// Employee-visible categories: a filtered view of M_Category where
// Employee_Visible=Yes. (The legacy M_Category_Employee sheet is no longer
// the source of truth — single master wins.)
// ============================================================
function getEmployeeCategories(){
  try {
    return getCategories().filter(function(c){ return c.employeeVisible === true; });
  } catch(e){ Logger.log('getEmployeeCategories: '+e.message); return []; }
}

// ============================================================
// 3. M_Party  (customers + vendors, auto sequential IDs)
//    Cols: Party_ID | Party_Type | Name | GSTIN | Email | Phone | Address |
//          Opening_Balance | Status | Created_Date
//    IDs: CUST-0001 / VEND-0001  — generated server-side, never reused.
// ============================================================
function setupPartyMaster(ss){
  ss = ss || _getSS();   // self-heal when called directly from the Apps Script editor
  var sheet = _buildSheet(ss, M.PARTY,
    ['Party_ID','Party_Type','Name','GSTIN','Email','Phone','Address',
     'Opening_Balance','Status','Created_Date'],
    [110, 100, 240, 150, 220, 130, 320, 130, 90, 160]);
  _validateCol(sheet, 2, ['Customer','Vendor']);
  _validateCol(sheet, 9, ['Active','Inactive']);
  return sheet;   // no seed — parties added via addParty()
}

// Generate the next sequential ID for a party type. Lock-guarded so two
// simultaneous adds can't grab the same number.
function _nextPartyId(sheet, partyType){
  var prefix = (partyType === 'Customer') ? 'CUST-' : 'VEND-';
  var max = 0;
  if (_hasData(sheet)){
    var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0;i<ids.length;i++){
      var id = _s(ids[i][0]);
      if (id.indexOf(prefix) === 0){
        var num = parseInt(id.substring(prefix.length), 10);
        if (!isNaN(num) && num > max) max = num;
      }
    }
  }
  return prefix + String(max + 1).padStart(4, '0');
}

// Add a customer or vendor. Returns {success, id} or {success:false, message}.
// data = {partyType, name, gstin, email, phone, address, openingBalance}
function addParty(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    var type = (_s(data.partyType) === 'Customer') ? 'Customer' : 'Vendor';
    var name = _s(data.name);
    if (!name) return {success:false, message:'Name is required.'};

    var ss = _getSS(), sheet = ss.getSheetByName(M.PARTY);

    // Smart duplicate guard: normalised name match within the same type
    // (catches "JALAN BROTHERS" vs "Jalan Brothers Pvt Ltd" etc.)
    if (_hasData(sheet)){
      var existing = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues(); // id, type, name
      for (var i=0;i<existing.length;i++){
        if (_s(existing[i][1]) === type && _normMatch(existing[i][2], name)){
          return {success:false, message:type + ' similar to "' + _s(existing[i][2]) + '" already exists ('+_s(existing[i][0])+').'};
        }
      }
    }

    var id = _nextPartyId(sheet, type);
    sheet.appendRow([
      id, type, name, _s(data.gstin), _s(data.email), _s(data.phone),
      _s(data.address), _n(data.openingBalance), 'Active',
      Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss')
    ]);
    return {success:true, id:id, name:name, partyType:type};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Read helpers for forms/dropdowns
function getParties(partyType){
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(M.PARTY);
    if (!sheet || !_hasData(sheet)) return [];
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues()
      .filter(function(r){ return _s(r[0]) && _s(r[8]) !== 'Inactive'
              && (!partyType || _s(r[1]) === partyType); })
      .map(function(r){ return {id:_s(r[0]), type:_s(r[1]), name:_s(r[2]), gstin:_s(r[3])}; });
    // Alphabetical by name, case-insensitive — makes long dropdowns scannable.
    rows.sort(function(a,b){
      var an = (a.name || '').toLowerCase();
      var bn = (b.name || '').toLowerCase();
      if (an < bn) return -1;
      if (an > bn) return 1;
      return 0;
    });
    return rows;
  } catch(e){ Logger.log('getParties: '+e.message); return []; }
}
function getCustomers(){ return getParties('Customer'); }
function getVendors(){   return getParties('Vendor'); }

// Find existing parties whose name looks similar to the query.
// Used by the autocomplete on the "Add new customer/vendor" form so users
// can spot duplicates as they type (rather than after Save).
//   query     — partial name being typed
//   partyType — 'Customer' | 'Vendor' | '' (all)
//   limit     — max results to return (default 8)
// Returns [{id, name, gstin, score}] sorted best-match first.
function findSimilarParties(query, partyType, limit){
  try {
    var q = _s(query);
    if (!q || q.length < 2) return [];                  // don't fire on single char
    limit = limit || 8;
    var nQ = _normName(q);
    if (!nQ) return [];

    var ss = _getSS(), sheet = ss.getSheetByName(M.PARTY);
    if (!sheet || !_hasData(sheet)) return [];
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();

    var matches = [];
    for (var i=0;i<rows.length;i++){
      var r = rows[i];
      var name = _s(r[2]); if (!name) continue;
      if (_s(r[8]) === 'Inactive') continue;
      if (partyType && _s(r[1]) !== partyType) continue;

      var nName = _normName(name);
      if (!nName) continue;

      // Score: 100 = exact match, 80 = prefix match either way, 60 = contains,
      //        40 = first 3+ chars match, 0 = no match
      var score = 0;
      if (nName === nQ)                                  score = 100;
      else if (nName.indexOf(nQ) === 0)                  score = 90;   // query is prefix of name
      else if (nQ.indexOf(nName) === 0 && nName.length >= 3) score = 80;
      else if (nName.indexOf(nQ) !== -1)                 score = 60;   // query found inside name
      else if (nQ.length >= 3 && nName.indexOf(nQ.substring(0,3)) === 0) score = 40;

      if (score > 0){
        matches.push({
          id: _s(r[0]),
          name: name,
          type: _s(r[1]),
          gstin: _s(r[3]),
          score: score
        });
      }
    }
    matches.sort(function(a,b){ return b.score - a.score; });
    return matches.slice(0, limit);
  } catch(e){ Logger.log('findSimilarParties: '+e.message); return []; }
}

// ============================================================
// 3b. M_Account  (owner + company bank accounts for fund transfers)
//     Cols: Account_ID | Display_Name | Owner | Bank | Side | Status
//       Side = 'Personal' (Girish/Reeta) or 'Company'
//     Used by the Fund Transfer flow + Source_of_Funds dropdowns.
// ============================================================
var ACC = { ID:0, NAME:1, OWNER:2, BANK:3, SIDE:4, STATUS:5 };

function setupAccountMaster(ss){
  ss = ss || _getSS();   // self-heal when called directly from the Apps Script editor
  var sheet = _buildSheet(ss, M.ACCOUNT,
    ['Account_ID','Display_Name','Owner','Bank','Side','Status'],
    [110, 240, 110, 120, 110, 90]);
  _validateCol(sheet, ACC.SIDE+1,   ['Personal','Company','Float']);
  _validateCol(sheet, ACC.STATUS+1, ['Active','Inactive']);

  // Seed the 7 accounts only if the sheet is empty (won't clobber edits).
  if (!_hasData(sheet)){
    var seed = [
      ['ACC-01','Girish – Axis – 91052',   'Girish',  'Axis',  'Personal','Active'],
      ['ACC-02','Girish – Kotak – 10999',  'Girish',  'Kotak', 'Personal','Active'],
      ['ACC-03','Girish – BOB – 20718',    'Girish',  'BOB',   'Personal','Active'],
      ['ACC-04','Reeta – Axis – 07873',    'Reeta',   'Axis',  'Personal','Active'],
      ['ACC-05','Reeta – – 24140',         'Reeta',   '',      'Personal','Active'],
      ['ACC-06','Company – Axis – 23253',  'Company', 'Axis',  'Company', 'Active'],
      ['ACC-07','Company – Kotak – 04352', 'Company', 'Kotak', 'Company', 'Active']
    ];
    sheet.getRange(2,1,seed.length,seed[0].length).setValues(seed);
  }
  return sheet;
}

// Returns [{id, name, owner, bank, side}] for active accounts.
function getAccounts(){
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(M.ACCOUNT);
    if (!sheet || !_hasData(sheet)) return [];
    return sheet.getRange(2,1,sheet.getLastRow()-1,6).getValues()
      .filter(function(r){ return _s(r[ACC.ID]) && _s(r[ACC.STATUS]) !== 'Inactive'; })
      .map(function(r){ return {
        id:_s(r[ACC.ID]), name:_s(r[ACC.NAME]), owner:_s(r[ACC.OWNER]),
        bank:_s(r[ACC.BANK]), side:_s(r[ACC.SIDE]) }; });
  } catch(e){ Logger.log('getAccounts: '+e.message); return []; }
}

// Look up one account by ID (or by display name as a fallback).
function getAccountByIdOrName(key){
  key = _s(key); if (!key) return null;
  var list = getAccounts();
  for (var i=0;i<list.length;i++){ if (list[i].id === key || list[i].name === key) return list[i]; }
  return null;
}

// ============================================================
// SITE FLOAT ACCOUNTS
//   A "float" is company cash held by a supervisor for site spending.
//   It's modeled as a regular M_Account with Side = 'Float'.
//
//   Lifecycle:
//     1. Owner transfers money from a company bank to a float account →
//        Fund Transfer, classified as Internal Transfer (no P&L impact).
//     2. Supervisor reports expenses paid from float → expenses hit P&L,
//        source = float account, settle_status = Settled (no pending row).
//     3. Supervisor returns unspent cash → Fund Transfer back to a company
//        bank account.
//   Float balance at any time = sum(transfers IN) − sum(transfers OUT) −
//                               sum(expenses with source = this float).
// ============================================================

// Create/refresh float accounts. Walks the employee master and ensures
// every employee with Holds_Float=Yes has a corresponding float account.
// Account IDs use the FL- prefix to keep them distinct from ACC- ones.
//   Returns {success, created:[names], total:N, skipped:[names]}.
function setupFloatAccounts(){
  var ss = _getSS();
  var accSheet = ss.getSheetByName(M.ACCOUNT);
  if (!accSheet) return {success:false, message:'M_Account not found. Run setupAccountMaster first.'};

  // Clear and re-apply Side validation across the WHOLE column. Old per-cell
  // validation rules with only ['Personal','Company'] would otherwise silently
  // reject any appended row whose Side = 'Float'.
  try {
    var maxRows = accSheet.getMaxRows();
    accSheet.getRange(2, ACC.SIDE+1, maxRows-1, 1).clearDataValidations();
    accSheet.getRange(2, ACC.SIDE+1, maxRows-1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Personal','Company','Float'], true).setAllowInvalid(false).build()
    );
  } catch(e){ Logger.log('setupFloatAccounts: clear Side validation: '+e.message); }

  var emps = (typeof getEmployeeMasterData === 'function') ? getEmployeeMasterData() : [];
  // ACCEPT ANY holder (active OR inactive) — if you ticked Holds_Float you want a float account.
  var floatHolders = emps.filter(function(e){ return e.holdsFloat; });
  if (!floatHolders.length){
    var m0 = 'No employees with Holds_Float=Yes found.\n\nOpen M_Employee, tick "Yes" in the Holds_Float column for each supervisor who carries site cash, then run this again.';
    try { SpreadsheetApp.getUi().alert(m0); } catch(e){}
    return {success:true, created:[], skipped:[], message:m0};
  }

  // Existing float accounts to skip duplicates
  var existing = getAccounts().filter(function(a){ return a.side === 'Float'; });
  var existingByName = {};
  existing.forEach(function(a){ existingByName[a.name.toLowerCase()] = a; });

  // Next FL-## sequence
  var nextSeq = 1;
  existing.forEach(function(a){
    var m = /^FL-(\d+)$/.exec(a.id);
    if (m) nextSeq = Math.max(nextSeq, parseInt(m[1],10) + 1);
  });

  var created = [], skipped = [], errors = [];
  floatHolders.forEach(function(emp){
    var displayName = 'Site Float — ' + emp.name;
    if (existingByName[displayName.toLowerCase()]){
      skipped.push(emp.name);
      return;
    }
    var id = 'FL-' + ('00' + nextSeq).slice(-2);
    try {
      accSheet.appendRow([id, displayName, emp.name, '—', 'Float', 'Active']);
      nextSeq++;
      created.push(emp.name);
    } catch(e){
      errors.push(emp.name + ' → ' + e.message);
    }
  });

  var msg = 'Float accounts setup result:\n• Created: ' + (created.length ? created.join(', ') : 'none') +
            (skipped.length ? '\n• Already existed (skipped): ' + skipped.join(', ') : '') +
            (errors.length ? '\n\nErrors:\n• ' + errors.join('\n• ') : '') +
            (created.length ? '\n\nNext: use Fund Transfer to move opening cash from a company bank account into each float.' : '');
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, created:created, skipped:skipped, errors:errors, message:msg};
}

// Returns just the float accounts (active). Used by the expense form's
// "Paid From" dropdown and the dashboard balance card.
function getFloatAccounts(){
  return getAccounts().filter(function(a){ return a.side === 'Float'; });
}

// Compute current balance of every float account.
// Balance = transfers in − transfers out − expenses with source = float.
// Returns [{id, name, owner, balance, lastActivity}] sorted by name.
// ============================================================
// diagnoseFloatSetup — prints state of the float subsystem
//
// Run from the Apps Script editor when float UI seems wrong. Reports:
//   • Employees with Holds_Float = Yes
//   • Float accounts found in M_Account
//   • Whether the names match (case-insensitive)
//   • The list as it would appear to the expense form's dropdown
// ============================================================
function diagnoseFloatSetup(){
  var lines = ['=== FLOAT SETUP DIAGNOSTIC ===\n'];

  // 1. Employees marked Holds_Float
  var emps = (typeof getEmployeeMasterData === 'function') ? getEmployeeMasterData() : [];
  var holders = emps.filter(function(e){ return e.holdsFloat; });
  lines.push('Employees with Holds_Float=Yes (active or inactive):');
  if (holders.length === 0) lines.push('  (none — open M_Employee, set Holds_Float=Yes for supervisors)');
  else holders.forEach(function(e){ lines.push('  • ' + e.name + (e.active ? '' : ' (inactive)')); });
  lines.push('');

  // 2. Float accounts in M_Account
  var accts = getAccounts();
  var floats = accts.filter(function(a){ return a.side === 'Float'; });
  lines.push('Float accounts in M_Account (Side="Float"):');
  if (floats.length === 0) lines.push('  (none — run setupFloatAccounts after marking Holds_Float=Yes)');
  else floats.forEach(function(a){
    lines.push('  • ' + a.id + ' | name="' + a.name + '" | owner="' + a.owner + '" | bank="' + a.bank + '"');
  });
  lines.push('');

  // 3. Matching test — for each holder, does a float account exist whose owner matches?
  lines.push('Match check (employee name vs float account owner):');
  if (!holders.length || !floats.length){
    lines.push('  (skipped — need at least one of each)');
  } else {
    holders.forEach(function(e){
      var n = (e.name || '').toLowerCase().trim();
      var hit = floats.filter(function(a){ return (a.owner || '').toLowerCase().trim() === n; });
      if (hit.length) lines.push('  ✓ ' + e.name + ' → ' + hit[0].id);
      else lines.push('  ✗ ' + e.name + ' → NO MATCH (owner field differs)');
    });
  }
  lines.push('');

  // 4. Frontend-visible accounts (what CFG.accounts will contain)
  lines.push('What the web app sees (getAccounts result, Float-side only):');
  if (!floats.length) lines.push('  (none)');
  else floats.forEach(function(a){
    lines.push('  side="' + a.side + '"  owner="' + a.owner + '"  → dropdown label: "From ' + a.name + '"');
  });
  lines.push('');

  var msg = lines.join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return msg;
}

function getFloatBalances(){
  try {
    var floats = getFloatAccounts();
    if (!floats.length) return _envOk({rows:[], total:0});

    // Index by both ID and display name (we cross-check both because the
    // Fund Transfer / expense Source field may store either).
    var byKey = {};
    floats.forEach(function(a){
      byKey[a.id] = {id:a.id, name:a.name, owner:a.owner, balance:0, lastActivity:''};
      byKey[a.name] = byKey[a.id];   // alias
    });

    // Walk transactions across all FYs; touch only what we need
    var all = (typeof getAllTransactions === 'function') ? getAllTransactions({fy:'all'}) : [];
    all.forEach(function(r){
      var src = _s(r.source);
      var note = _s(r.notes);
      // Find any float account ID/name appearing in source
      // Quick path: exact match against keys
      var hit = byKey[src];
      // Also consider transfers — the destination is embedded in notes/category for Transfer rows
      if (!hit && r.txnClass === 'Transfer'){
        // Transfer rows have category like "Fund In (Owner)" / "Internal Transfer" / etc.
        // We store both legs of the transfer; the 'source' is the FROM account.
        // For float accounts, OUT-direction txns reduce balance, IN-direction add.
        // Match by string presence of any known float account in source field.
        for (var k in byKey){
          if (byKey.hasOwnProperty(k) && src === k){ hit = byKey[k]; break; }
        }
      }
      if (!hit) return;

      var entry = hit;
      var amt = _n(r.amount);
      // Sign: for Transfer + IN direction → balance increases (money coming into the float).
      //       For Transfer + OUT direction → balance decreases (returning cash).
      //       For Expense with source=float → balance decreases.
      if (r.txnClass === 'Transfer'){
        if (_s(r.direction) === 'IN') entry.balance += amt;
        else entry.balance -= amt;
      } else if (r.txnClass === 'Expense'){
        entry.balance -= amt;
      }
      var d = _s(r.txDate);
      if (d && d > entry.lastActivity) entry.lastActivity = d;
    });

    // Dedupe (we aliased keys); return one entry per actual account
    var seen = {}, rows = [];
    floats.forEach(function(a){
      if (seen[a.id]) return;
      seen[a.id] = true;
      var e = byKey[a.id];
      rows.push({id:e.id, name:e.name, owner:e.owner, balance:_round(e.balance), lastActivity:e.lastActivity});
    });
    rows.sort(function(a,b){ return a.name.localeCompare(b.name); });

    var total = rows.reduce(function(s,r){ return s + r.balance; }, 0);
    return _envOk({rows:rows, total:_round(total)});
  } catch(e){ return _envErr(e); }
}


// ============================================================
// 4. M_Config  (typed lists — one tab, filter by List_Type)
//    Cols: List_Type | Value | Attribute | Active
//    List_Type values used:
//      PaymentMode    — Value = mode name
//      SourceOfFunds  — Value = account name, Attribute = Personal|Company
//      PartyStatus    — Active|Inactive
// ============================================================
var CONFIG_SEED = [
  // Payment modes — unified superset reconciling all three old systems
  ['PaymentMode','Cash',''],
  ['PaymentMode','Bank Transfer',''],
  ['PaymentMode','UPI / PhonePe',''],
  ['PaymentMode','Cheque',''],
  ['PaymentMode','Credit Card',''],
  ['PaymentMode','ECS / Mandate / Auto',''],
  ['PaymentMode','A/C Adjustment',''],
  ['PaymentMode','Account Transfer',''],
  ['PaymentMode','Advance',''],
  ['PaymentMode','To Be Paid',''],
  // Source of funds — Attribute = which side owns the account (drives auto-drawing)
  ['SourceOfFunds','Three Star Projects Pvt Ltd Account','Company'],
  ['SourceOfFunds','3 Star Projects Account (old)','Company'],
  ['SourceOfFunds','GPay (Company)','Company'],
  ['SourceOfFunds','GPay (old company)','Company'],
  ['SourceOfFunds','Cash (Company)','Company'],
  ['SourceOfFunds','Personal Account','Personal'],
  ['SourceOfFunds','GPay (Personal)','Personal'],
  ['SourceOfFunds','Cash (Personal)','Personal']
];

function setupConfigMaster(ss){
  ss = ss || _getSS();   // self-heal when called directly from the Apps Script editor
  var sheet = _buildSheet(ss, M.CONFIG,
    ['List_Type','Value','Attribute','Active'], [150, 320, 120, 70]);
  _validateCol(sheet, 4, ['Yes','No']);
  if (!_hasData(sheet)){
    var rows = CONFIG_SEED.map(function(r){ return [r[0], r[1], r[2], 'Yes']; });
    sheet.getRange(2,1,rows.length,4).setValues(rows);
  }
  return sheet;
}

// Read a config list by type. attributeFilter optional (e.g. 'Company').
function getConfigList(listType, attributeFilter){
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(M.CONFIG);
    if (!sheet || !_hasData(sheet)) return [];
    return sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues()
      .filter(function(r){ return _s(r[0]) === listType && _s(r[3]) !== 'No'
              && (!attributeFilter || _s(r[2]) === attributeFilter); })
      .map(function(r){ return {value:_s(r[1]), attribute:_s(r[2])}; });
  } catch(e){ Logger.log('getConfigList: '+e.message); return []; }
}
function getPaymentModes(){   return getConfigList('PaymentMode').map(function(o){return o.value;}); }
function getSourcesOfFunds(){ return getConfigList('SourceOfFunds'); }   // keeps Personal/Company tag

// Given a source-of-funds name, is it a Company account? (drives auto-drawing)
function isCompanyAccount(sourceName){
  var list = getConfigList('SourceOfFunds');
  for (var i=0;i<list.length;i++){ if (list[i].value === _s(sourceName)) return list[i].attribute === 'Company'; }
  return false;
}

// ============================================================
// CATEGORY ADD  ("Other" → new category, GROUP REQUIRED)
//    Called from entry forms when the owner picks "Other".
//    Refuses to save without a valid P&L group → P&L can never leak.
// ============================================================
function addCategory(data){
  // data = {category, group, direction, gstApplicable, gstRate, hsnSac}
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    var name  = _s(data.category);
    var group = _s(data.group);
    if (!name)  return {success:false, message:'Category name is required.'};
    if (!group) return {success:false, message:'A P&L group is required for every category.'};
    if (PNL_GROUPS.indexOf(group) === -1)
      return {success:false, message:'"' + group + '" is not a valid P&L group.'};

    var ss = _getSS(), sheet = ss.getSheetByName(M.CATEGORY);
    var lower = name.toLowerCase();
    if (_hasData(sheet)){
      var existing = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
      for (var i=0;i<existing.length;i++){
        if (_s(existing[i][0]).toLowerCase() === lower)
          return {success:true, added:false, message:'Category already exists.'};
      }
    }
    var dir = (_s(data.direction) === 'IN') ? 'IN' : 'OUT';
    sheet.appendRow([ name, group, dir, _s(data.gstApplicable) || 'No',
                      _s(data.gstRate), _s(data.hsnSac), 'Yes' ]);
    return {success:true, added:true, category:name, group:group};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Update an existing category, matched by original name.
//   data = {originalCategory, category, group, direction, active}
//   M_Category cols: Name | Group | Direction | GST_Applicable | GST_Rate | HSN_SAC | Active
function updateCategory(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy, try again.'}; }
  try {
    var orig = _s(data.originalCategory || data.category);
    if (!orig) return {success:false, message:'Original category name required.'};
    var newName = _s(data.category) || orig;
    var group   = _s(data.group);
    if (!newName) return {success:false, message:'Category name is required.'};
    if (group && PNL_GROUPS.indexOf(group) === -1)
      return {success:false, message:'"' + group + '" is not a valid P&L group.'};

    var ss = _getSS(), sheet = ss.getSheetByName(M.CATEGORY);
    if (!_hasData(sheet)) return {success:false, message:'No categories found.'};
    var n = sheet.getLastRow()-1;
    var vals = sheet.getRange(2,1,n,7).getValues();

    // Guard against rename collision
    if (newName.toLowerCase() !== orig.toLowerCase()){
      for (var k=0;k<vals.length;k++){
        if (_s(vals[k][0]).toLowerCase() === newName.toLowerCase())
          return {success:false, message:'Another category already has that name.'};
      }
    }
    for (var i=0;i<vals.length;i++){
      if (_s(vals[i][0]).toLowerCase() === orig.toLowerCase()){
        var rowNum = i + 2;
        if (data.hasOwnProperty('category'))  sheet.getRange(rowNum,1).setValue(newName);
        if (data.hasOwnProperty('group') && group) sheet.getRange(rowNum,2).setValue(group);
        if (data.hasOwnProperty('direction')) sheet.getRange(rowNum,3).setValue(_s(data.direction)==='IN'?'IN':'OUT');
        if (data.hasOwnProperty('active'))    sheet.getRange(rowNum,7).setValue(data.active?'Yes':'No');
        return {success:true, message:'Category updated.', category:newName};
      }
    }
    return {success:false, message:'Category not found.'};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// applyCategoryDropdowns — apply data-validation dropdowns to the Category
// column in every Transactions_<FY> sheet (and Personal_Transactions).
//
// Reads the live category list from M_Category. Safe to run repeatedly —
// just refreshes the rule. Run this once after taxonomy changes or whenever
// you want the in-sheet Category cells to show a picker.
// ============================================================
function applyCategoryDropdowns(){
  var ss = _getSS();
  // Build live category list from M_Category (active rows only)
  var catSheet = ss.getSheetByName(M.CATEGORY);
  if (!catSheet || !_hasData(catSheet)){
    var m0 = 'M_Category is empty — run setupCategoryMaster first.';
    try { SpreadsheetApp.getUi().alert(m0); } catch(e){ Logger.log(m0); }
    return {success:false, message:m0};
  }
  var catRows = catSheet.getRange(2, 1, catSheet.getLastRow()-1, 7).getValues();
  var allCats = catRows.filter(function(r){ return _s(r[0]) && _s(r[6]) !== 'No'; })
                       .map(function(r){ return _s(r[0]); });
  // Personal categories (for the Personal_Transactions sheet)
  var pcatSheet = ss.getSheetByName(M.CATEGORY_PERSONAL);
  var personalCats = [];
  if (pcatSheet && _hasData(pcatSheet)){
    personalCats = pcatSheet.getRange(2, 1, pcatSheet.getLastRow()-1, 1).getValues()
                            .map(function(r){ return _s(r[0]); })
                            .filter(function(s){ return s; });
  }

  if (!allCats.length){
    var m1 = 'No active categories found in M_Category.';
    try { SpreadsheetApp.getUi().alert(m1); } catch(e){ Logger.log(m1); }
    return {success:false, message:m1};
  }

  var report = { sheetsProcessed: 0, personalApplied: false };

  // Apply to every Transactions_<FY> sheet
  var txnSheets = _allTxnSheets();
  txnSheets.forEach(function(sheet){
    if (!sheet) return;
    try {
      var maxRows = sheet.getMaxRows();
      var rng = sheet.getRange(2, TC.CATEGORY+1, maxRows-1, 1);
      rng.clearDataValidations();
      rng.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(allCats, true)
          .setAllowInvalid(true)   // allow legacy/typed values without breaking
          .setHelpText('Pick a category from the master list.')
          .build()
      );
      report.sheetsProcessed++;
    } catch(e){ Logger.log('applyCategoryDropdowns ' + sheet.getName() + ': ' + e.message); }
  });

  // Personal_Transactions — schema-aware: after Tier 5.1 the Category column
  // moved from index 5 → 8. Refuse to touch the sheet if the widths disagree
  // with PC_TOTAL — running with the wrong column index would validate the
  // wrong data. Owner should run upgradePersonalTransactionsSchema first.
  var pSheet = ss.getSheetByName(T.PERSONAL);
  if (pSheet && personalCats.length){
    if (pSheet.getLastColumn() < PC_TOTAL){
      Logger.log('applyCategoryDropdowns: Personal_Transactions is pre-Tier-5.1 (' +
                 pSheet.getLastColumn() + ' cols). Skipping — run upgradePersonalTransactionsSchema first.');
    } else {
      try {
        var pMaxRows = pSheet.getMaxRows();
        var pRng = pSheet.getRange(2, PC.CATEGORY+1, pMaxRows-1, 1);
        pRng.clearDataValidations();
        pRng.setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(personalCats, true)
            .setAllowInvalid(true)
            .setHelpText('Pick a personal category from the master list.')
            .build()
        );
        report.personalApplied = true;
      } catch(e){ Logger.log('applyCategoryDropdowns Personal: ' + e.message); }
    }
  }

  var msg = 'Category dropdowns applied:\n' +
            '\n• Transaction sheets updated: ' + report.sheetsProcessed +
            '\n• Personal_Transactions updated: ' + (report.personalApplied ? 'yes' : 'no') +
            '\n• Categories in dropdown: ' + allCats.length +
            (personalCats.length ? ' (personal: ' + personalCats.length + ')' : '') +
            '\n\nThe Category column in every transaction sheet now shows a picker.\nWhen you add new categories to M_Category, re-run this to refresh the list.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, report:report};
}

// ============================================================
// repairTransactionGroups — safety-net group fixer
//
// After migrateCategories runs, some transaction rows may still have
// stale group names (especially if the category itself wasn't in the
// rename map). This function rebuilds the GROUP column on every
// Transactions_<FY> row purely from the CATEGORY value, using the live
// M_Category master as the source of truth.
//
// Idempotent — safe to run repeatedly. Run if you see red-bordered
// cells in the Group column of any transaction sheet.
// ============================================================
function repairTransactionGroups(){
  var ss = _getSS();
  // Build category → group lookup from the live M_Category sheet.
  // We read from the sheet (not CATEGORY_SEED) so any custom categories
  // the user added are included.
  var lookup = {};
  var catSheet = ss.getSheetByName(M.CATEGORY);
  if (catSheet && _hasData(catSheet)){
    var rows = catSheet.getRange(2, 1, catSheet.getLastRow()-1, 2).getValues();
    rows.forEach(function(r){
      var name = _s(r[0]), grp = _s(r[1]);
      if (name && grp) lookup[name] = grp;
    });
  }
  if (!Object.keys(lookup).length){
    var m = 'M_Category is empty — run setupCategoryMaster first.';
    try { SpreadsheetApp.getUi().alert(m); } catch(e){ Logger.log(m); }
    return {success:false, message:m};
  }

  var report = { sheetsProcessed:0, rowsRepaired:0, unknownCategories:[] };
  var txnSheets = _allTxnSheets();
  txnSheets.forEach(function(sheet){
    if (!sheet || sheet.getLastRow() < 2) return;
    report.sheetsProcessed++;
    var lastRow = sheet.getLastRow();

    // Clear stale validations so writes won't be rejected
    try {
      var maxRows = sheet.getMaxRows();
      sheet.getRange(2, TC.GROUP+1, maxRows-1, 1).clearDataValidations();
    } catch(e){ Logger.log('clearDataValidations '+sheet.getName()+': '+e.message); }

    // Re-derive Group from Category for every row
    var groupCol = sheet.getRange(2, TC.GROUP+1, lastRow-1, 1).getValues();
    var catCol   = sheet.getRange(2, TC.CATEGORY+1, lastRow-1, 1).getValues();
    var changed = false;
    for (var i=0; i<catCol.length; i++){
      var cat = _s(catCol[i][0]);
      if (!cat) continue;
      var expectedGroup = lookup[cat];
      if (expectedGroup && expectedGroup !== _s(groupCol[i][0])){
        groupCol[i][0] = expectedGroup;
        changed = true;
        report.rowsRepaired++;
      } else if (!expectedGroup){
        if (report.unknownCategories.indexOf(cat) < 0) report.unknownCategories.push(cat);
      }
    }
    if (changed) sheet.getRange(2, TC.GROUP+1, groupCol.length, 1).setValues(groupCol);

    // Re-apply the validation rule with the NEW PNL_GROUPS list
    try {
      _validateCol(sheet, TC.GROUP+1, PNL_GROUPS);
    } catch(e){ Logger.log('re-validate '+sheet.getName()+': '+e.message); }
  });

  var summary = 'Group repair complete:\n' +
                '\n• Sheets processed: ' + report.sheetsProcessed +
                '\n• Rows repaired: ' + report.rowsRepaired +
                (report.unknownCategories.length
                  ? '\n\nCategories present in transactions but not in M_Category (group left unchanged):\n• ' + report.unknownCategories.join('\n• ')
                  : '\n\nNo unknown categories — every transaction maps to a category in M_Category.');
  try { SpreadsheetApp.getUi().alert(summary); } catch(e){ Logger.log(summary); }
  return {success:true, report:report};
}


//
// Run this AFTER deploying the new CATEGORY_SEED and Employee_Visible
// column. It does three things:
//   1. Renames old categories to new ones in M_Category itself (so the
//      master picks up the cleaner naming).
//   2. Renames old categories AND updates Group in every Transactions_<FY>
//      and Personal_Transactions row.
//   3. Adds any brand-new categories from CATEGORY_SEED that aren't yet
//      in M_Category (preserving any custom rows the user added).
//   4. Reports unmapped categories so the user can fix them manually.
// Safe to run multiple times — idempotent.
// ============================================================
function migrateCategories(){
  // ── Mapping: old name → new name. Only categories that were RENAMED
  //    are listed here. Categories whose name is unchanged stay as-is.
  //    Group is taken from CATEGORY_SEED automatically — no need to map
  //    groups separately.
  var RENAMES = {
    // Bank, Interest, Loan group → Bank, Finance & Loans
    'Loan / Advances / Rental':         'Advance Given to Vendor',
    'OD Utilization':                   'OD / CC Utilization Cost',

    // Salary & Wages group renames
    'Employer Monthly Salary':          'Employee Monthly Salary',
    'Labour Charges':                   'Daily Labour Charges',
    'PF':                               'PF (Provident Fund)',
    // LIC / Policy / Savings has no clear company home — flag for review
    // (we map it to "Other Salary (Specify)" so it's grouped sensibly).
    'LIC / Policy / Savings':           'Other Salary (Specify)',

    // Utilities renames
    'Electricity Expense':              'Electricity Bill',
    'PNG / Diesel':                     'PNG / Diesel (Industrial Use)',

    // Finance & Compliance → Taxes & Compliance + others
    'Commission':                       'Commission Paid',
    'GST':                              'GST Paid',
    'TDS':                              'TDS Paid',

    // Production → Direct Production renames
    'Consumables (e.g., fasteners, packaging material)':  'Consumables (fasteners, packaging)',
    'Die, Mold Development & Repair':   'Die / Mould Development & Repair',
    'Heavy Labour Charges / Loading & Unloading':         'Heavy Labour / Loading-Unloading',
    'Packing Material Purchase (Poly, Box, Stickers etc.)':'Packing Material',
    'Parts / Product Purchase':         'Parts / Components Purchase',
    'Tooling Purchase & Repair (CNC Inserts, Drills, Bits etc.)':'Tooling & Drill Bits',
    'Vendor Parts (Wheels, Seats etc.)':'Vendor Parts (Wheels, Seats, etc.)',

    // Logistics renames
    'Freight & Forwarding':             'Freight & Forwarding (outbound)',

    // Admin renames
    'Food / Lunch / Dinner Expense':    'Office Food / Lunch / Dinner',
    'Medical Expense':                  'Employee Medical',
    'Tea & Snacks (All Types of Kitchen Purchases)':      'Tea, Coffee, Kitchen Supplies',
    'Housekeeping Expense & Purchases': 'Housekeeping (supplies & services)',
    'Office':                           'Office Supplies / Misc',
    'Rent Paid':                        'Rent Paid (office/factory)',
    'Stationery':                       'Stationery & Printing',

    // Marketing → Sales & Marketing
    'Marketing / Expo / Advertisement / Showcasing': 'Marketing / Advertising',

    // Transport & Fuel → Logistics & Transport
    'Petrol':                           'Petrol / Diesel (Vehicle Fuel)',
    'Vehicle Running / Maintenance':    'Vehicle Running & Maintenance',

    // Professional Services rename
    'Professional / Consultancy Charges':'Consultancy Charges (general)',

    // Sales / Travel → Sales & Marketing
    'Tour / Travel Expense':            'Sales Tour Travel (multi-day)',

    // ── Legacy EMPLOYEE-category renames (from the old M_Category_Employee
    //    list, which is no longer the source of truth). These map old
    //    employee-form values → the new unified category they're closest to.
    'Travel':                           'Employee Travel (Bus/Train/Auto)',
    'Food':                             'Office Food / Lunch / Dinner',
    'Factory Related':                  'Other Production (Specify)',
    'Logistics':                        'Freight & Forwarding (outbound)',
    'Office Related':                   'Office Supplies / Misc',
    'Payment / Settlement':             'Other Salary (Specify)',
    'Equipment / Machinery':            'Machinery Repair & Maintenance',
    'Stationery & Printing':            'Stationery & Printing',
    'Repair & Maintenance':             'Office Repair & Maintenance',
    'Miscellaneous':                    'Other Admin (Specify)',
    'Utilities':                        'Other Utilities (Specify)'
  };

  var ss = _getSS();
  var report = { categoriesRenamed:0, txRowsUpdated:0, personalRowsUpdated:0, newCatsAdded:0, unmapped:[] };

  // Build a lookup of NEW categories by name → group from the seed
  var seedLookup = {};
  CATEGORY_SEED.forEach(function(p){
    seedLookup[p[0]] = {group: p[1], direction: p[2], employeeVisible: p[3]};
  });

  // ── Step 1: update M_Category sheet ──
  var catSheet = ss.getSheetByName(M.CATEGORY);
  if (catSheet && _hasData(catSheet)){
    var lastCol = Math.max(catSheet.getLastColumn(), 7);
    var width = Math.min(lastCol, 8);

    // CRITICAL: clear old per-cell data validations on Group/Direction/
    // Employee_Visible columns. Otherwise, writing the new group names
    // (e.g. "Direct Production") fails because cells still have the OLD
    // group list as their dropdown rule.
    try {
      var maxRows = catSheet.getMaxRows();
      catSheet.getRange(2, 2, maxRows-1, 1).clearDataValidations();   // Group col
      catSheet.getRange(2, 3, maxRows-1, 1).clearDataValidations();   // Direction col
      if (width >= 8) catSheet.getRange(2, 8, maxRows-1, 1).clearDataValidations(); // Employee_Visible col
    } catch(e){ Logger.log('clearDataValidations: '+e.message); }

    var rows = catSheet.getRange(2, 1, catSheet.getLastRow()-1, width).getValues();
    var existingNames = {};
    rows.forEach(function(r){ existingNames[_s(r[0])] = true; });

    // Rewrite existing rows (rename + update group)
    var changed = false;
    for (var i=0;i<rows.length;i++){
      var oldName = _s(rows[i][0]);
      var newName = RENAMES[oldName] || oldName;
      if (newName !== oldName){
        rows[i][0] = newName;
        changed = true;
        report.categoriesRenamed++;
      }
      // If this category exists in seed, refresh group/direction
      if (seedLookup[newName]){
        rows[i][1] = seedLookup[newName].group;
        rows[i][2] = seedLookup[newName].direction;
        if (width >= 8){ rows[i][7] = seedLookup[newName].employeeVisible ? 'Yes' : 'No'; }
        changed = true;
      }
    }
    if (changed) catSheet.getRange(2, 1, rows.length, width).setValues(rows);

    // Re-apply the NEW validation rules so future manual edits are still constrained
    try {
      _validateCol(catSheet, 2, PNL_GROUPS);
      _validateCol(catSheet, 3, ['OUT','IN']);
      if (width >= 8) _validateCol(catSheet, 8, ['Yes','No']);
    } catch(e){ Logger.log('re-validate M_Category: '+e.message); }

    // ── Add brand-new categories from seed that aren't in the sheet ──
    // Refresh existingNames after renames
    var existsNow = {};
    catSheet.getRange(2, 1, catSheet.getLastRow()-1, 1).getValues().forEach(function(r){
      existsNow[_s(r[0])] = true;
    });
    var toAdd = [];
    CATEGORY_SEED.forEach(function(p){
      if (!existsNow[p[0]]){
        toAdd.push([p[0], p[1], p[2], 'No', '', '', 'Yes', p[3] ? 'Yes' : 'No']);
      }
    });
    if (toAdd.length){
      var startRow = catSheet.getLastRow() + 1;
      catSheet.getRange(startRow, 1, toAdd.length, 8).setValues(toAdd);
      report.newCatsAdded = toAdd.length;
    }
  }

  // ── Step 2: walk every transaction sheet, rename categories + update Group ──
  var txnSheets = _allTxnSheets();
  txnSheets.forEach(function(sheet){
    if (!sheet || sheet.getLastRow() < 2) return;
    var lastRow = sheet.getLastRow();

    // Clear validations on Category + Group columns so writes of renamed
    // values aren't rejected by stale per-cell dropdown rules.
    try {
      var maxRows = sheet.getMaxRows();
      sheet.getRange(2, TC.CATEGORY+1, maxRows-1, 1).clearDataValidations();
      sheet.getRange(2, TC.GROUP+1,    maxRows-1, 1).clearDataValidations();
    } catch(e){ Logger.log('clearDataValidations '+sheet.getName()+': '+e.message); }

    var data = sheet.getRange(2, 1, lastRow-1, TC_TOTAL).getValues();
    var changed = false;
    for (var i=0;i<data.length;i++){
      var cat = _s(data[i][TC.CATEGORY]);
      if (!cat) continue;
      var newCat = RENAMES[cat] || cat;
      // Look up the new group
      var newGroup = seedLookup[newCat] ? seedLookup[newCat].group : _s(data[i][TC.GROUP]);
      if (newCat !== cat || newGroup !== _s(data[i][TC.GROUP])){
        data[i][TC.CATEGORY] = newCat;
        data[i][TC.GROUP] = newGroup;
        changed = true;
        report.txRowsUpdated++;
        if (!seedLookup[newCat] && report.unmapped.indexOf(newCat) < 0){
          report.unmapped.push(newCat);
        }
      }
    }
    if (changed) sheet.getRange(2, 1, data.length, TC_TOTAL).setValues(data);

    // Re-apply validation with the NEW PNL_GROUPS list. (Category column gets
    // wider validation: any non-empty string — we don't pin to the master here
    // since the app handles category selection.)
    try {
      _validateCol(sheet, TC.GROUP+1, PNL_GROUPS);
    } catch(e){ Logger.log('re-validate '+sheet.getName()+': '+e.message); }
  });

  // ── Step 3: Personal_Transactions ──
  // Schema-aware: the Category column moved from index 5 → 8 in Tier 5.1.
  // Detect the actual layout by the sheet's current width so a pre-migration
  // sheet doesn't get its category writes redirected to the wrong column.
  var pSheet = ss.getSheetByName(T.PERSONAL);
  if (pSheet && pSheet.getLastRow() >= 2){
    var pWidth   = pSheet.getLastColumn();
    var pCatIdx  = (pWidth >= PC_TOTAL) ? PC.CATEGORY : 5;   // 5 = old layout
    var pReadW   = Math.max(pWidth, pCatIdx + 1);
    try {
      pSheet.getRange(2, pCatIdx+1, pSheet.getMaxRows()-1, 1).clearDataValidations();
    } catch(e){ Logger.log('clearDataValidations Personal: '+e.message); }
    var pData = pSheet.getRange(2, 1, pSheet.getLastRow()-1, pReadW).getValues();
    var pChanged = false;
    for (var j=0;j<pData.length;j++){
      var pcat = _s(pData[j][pCatIdx]);
      if (!pcat) continue;
      var npcat = RENAMES[pcat] || pcat;
      if (npcat !== pcat){
        pData[j][pCatIdx] = npcat;
        pChanged = true;
        report.personalRowsUpdated++;
      }
    }
    if (pChanged) pSheet.getRange(2, 1, pData.length, pReadW).setValues(pData);
  }

  // ── Refresh dropdowns/validation on transaction sheets ──
  // (Group + Category validations may reference the master)
  // Skipped — categories are typed-in by the app, not via sheet dropdown,
  // and the sheet validation is broad. No action needed here.

  // ── Report ──
  var summary = 'Category migration complete:\n' +
                '\n• Categories renamed in M_Category: ' + report.categoriesRenamed +
                '\n• New categories added from seed: ' + report.newCatsAdded +
                '\n• Transaction rows updated: ' + report.txRowsUpdated +
                '\n• Personal rows updated: ' + report.personalRowsUpdated +
                (report.unmapped.length ? '\n\nUnmapped category names found in transactions (please review):\n• ' + report.unmapped.join('\n• ') : '') +
                '\n\nM_Category_Employee sheet is no longer the source of truth;\nemployee-visible categories now come from M_Category itself\n(filtered by Employee_Visible = Yes).';
  try { SpreadsheetApp.getUi().alert(summary); } catch(e){ Logger.log(summary); }
  return {success:true, report:report};
}

// Read categories for dropdowns. Returns [{category, group, direction, employeeVisible}], active only.
// M_Category is the SOLE source of truth. A row is excluded unless it has a
// non-empty Name, Active != 'No', and Direction explicitly set to 'OUT' or
// 'IN'. Rows with blank Direction are rejected — they would otherwise leak
// into both the Company Expense and Income dropdowns.
function getCategories(){
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(M.CATEGORY);
    if (!sheet || !_hasData(sheet)) return [];
    var lastCol = Math.max(sheet.getLastColumn(), 7);
    var width = Math.min(lastCol, 8);
    var out = sheet.getRange(2,1,sheet.getLastRow()-1,width).getValues()
      .filter(function(r){
        var name = _s(r[0]);
        var dir  = _s(r[2]).toUpperCase();
        var act  = _s(r[6]);
        return name && (dir === 'OUT' || dir === 'IN') && act !== 'No';
      })
      .map(function(r){
        return {
          category: _s(r[0]),
          group:    _s(r[1]),
          direction:_s(r[2]).toUpperCase(),
          employeeVisible: (width >= 8) ? (_s(r[7]) === 'Yes') : false
        };
      });
    // Alphabetical (case-insensitive) so long dropdowns are easy to scan
    out.sort(function(a,b){
      var x = (a.category||'').toLowerCase(), y = (b.category||'').toLowerCase();
      return x < y ? -1 : x > y ? 1 : 0;
    });
    return out;
  } catch(e){ Logger.log('getCategories: '+e.message); return []; }
}
function getPersonalCategories(){
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(M.CATEGORY_PERSONAL);
    if (!sheet || !_hasData(sheet)) return [];
    return sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues()
      .filter(function(r){ return _s(r[0]) && _s(r[1]) !== 'No'; })
      .map(function(r){ return _s(r[0]); });
  } catch(e){ Logger.log('getPersonalCategories: '+e.message); return []; }
}
function getPnlGroups(){ return PNL_GROUPS.slice(); }

// ── Resolve a party by name within a type, creating it if new ──
// Returns {id, name, created}. Used so forms can accept a typed-in vendor
// name (pick from list OR add new) and still link a real party ID.
function findOrCreateParty(partyType, name){
  var type = (_s(partyType) === 'Customer') ? 'Customer' : 'Vendor';
  var nm = _s(name);
  if (!nm) return {id:'', name:'', created:false};
  var ss = _getSS(), sheet = ss.getSheetByName(M.PARTY);
  if (sheet && _hasData(sheet)){
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues(); // id, type, name
    // Pass 1: exact normalised match
    for (var i=0;i<rows.length;i++){
      if (_s(rows[i][1]) === type && _normName(rows[i][2]) === _normName(nm)){
        return {id:_s(rows[i][0]), name:_s(rows[i][2]), created:false};
      }
    }
    // Pass 2: prefix-based fuzzy match (truncation tolerant)
    for (var j=0;j<rows.length;j++){
      if (_s(rows[j][1]) === type && _normMatch(rows[j][2], nm)){
        return {id:_s(rows[j][0]), name:_s(rows[j][2]), created:false};
      }
    }
  }
  var r = addParty({partyType:type, name:nm});
  return r && r.success ? {id:r.id, name:r.name, created:true} : {id:'', name:nm, created:false};
}
// ============================================================
// applyProjectDropdowns — data-validation dropdown on every "Project" column
// across Transactions_<FY>, Bills, Invoices, Payments, Bill_Payments.
// Source: names from M_Project (active). Soft dropdown (allowInvalid=true) —
// pastes and legacy values won't be rejected. Run any time after adding new
// projects or new FY tabs.
// ============================================================
function applyProjectDropdowns(){
  var names = (typeof getProjectNames === 'function') ? (getProjectNames() || []) : [];
  if (!names.length){
    var m = 'No active projects in M_Project — add projects first, then re-run.';
    try { SpreadsheetApp.getUi().alert(m); } catch(e){ Logger.log(m); }
    return {success:false, message:m};
  }
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true).setAllowInvalid(true)
    .setHelpText('Pick from M_Project. Add new projects in M_Project then re-run applyProjectDropdowns.')
    .build();

  var ss = _getSS();
  var applied = [];

  // (sheetName, projectColIndex0Based, expectedTotalCols)
  var targets = [
    { sheet: ss.getSheetByName(AP.BILLS),       col: BC.PROJECT,   total: BC_TOTAL,   label: 'Bills' },
    { sheet: ss.getSheetByName(AR.INVOICES),    col: IC.PROJECT,   total: IC_TOTAL,   label: 'Invoices' },
    { sheet: ss.getSheetByName(AR.PAYMENTS),    col: PMC.PROJECT,  total: PMC_TOTAL,  label: 'Payments' },
    { sheet: ss.getSheetByName(AP.PAYMENTS),    col: BPMC.PROJECT, total: BPMC_TOTAL, label: 'Bill_Payments' },
    { sheet: ss.getSheetByName(AP.ALLOCATIONS), col: BAC.PROJECT,  total: BAC_TOTAL,  label: 'Bill_Allocations' }
  ];
  _allTxnSheets().forEach(function(s){
    targets.push({ sheet: s, col: TC.PROJECT, total: TC_TOTAL, label: s.getName() });
  });

  targets.forEach(function(t){
    if (!t.sheet){ return; }
    if (t.sheet.getLastColumn() <= t.col){
      Logger.log('applyProjectDropdowns: ' + t.label + ' missing Project column — run upgradePaymentSchemas / setupBillsModule first.');
      return;
    }
    var maxRows = t.sheet.getMaxRows();
    var rng = t.sheet.getRange(2, t.col + 1, maxRows - 1, 1);
    rng.clearDataValidations();
    rng.setDataValidation(rule);
    applied.push(t.label);
  });

  var msg = 'Project dropdowns applied to:\n• ' + applied.join('\n• ') +
            '\n\nProjects available: ' + names.length +
            '\nRe-run this after adding new projects or new FY sheets.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, sheets:applied, projectCount:names.length};
}
// ============================================================
// Standalone Bills-schema repair — grows the Bills sheet to 25 cols and
// writes 'Description' at col 24 and 'Group' at col 25 directly, without
// going through _buildSheet. Idempotent: running it on a 25-wide sheet is
// a no-op. Callable from the Apps Script editor if the user wants to run
// it on its own.
// ============================================================
function _repairBillsSchemaForDescGroup(){
  var ss = _getSS();
  var sheet = ss.getSheetByName('Bills');
  if (!sheet) return {success:false, message:'Bills sheet not found. Run setupBillsModule first.'};

  // 1. Ensure physical column count ≥ 25.
  var maxCols = sheet.getMaxColumns();
  if (maxCols < 25) sheet.insertColumnsAfter(maxCols, 25 - maxCols);

  // 2. Write the two header cells explicitly, one setValue at a time so a
  //    silent-truncate failure on a bulk range doesn't affect us.
  var h24 = String(sheet.getRange(1, 24).getValue() || '').trim();
  var h25 = String(sheet.getRange(1, 25).getValue() || '').trim();
  if (h24 !== 'Description') sheet.getRange(1, 24).setValue('Description');
  if (h25 !== 'Group')       sheet.getRange(1, 25).setValue('Group');

  // 3. Style them like the rest of the header row (uses BRAND if available).
  try {
    if (typeof BRAND !== 'undefined' && BRAND.headerBg){
      sheet.getRange(1, 24, 1, 2)
           .setBackground(BRAND.headerBg).setFontColor(BRAND.headerFg)
           .setFontWeight('bold').setFontSize(11);
    }
    sheet.setColumnWidth(24, 260);
    sheet.setColumnWidth(25, 160);
  } catch(e){ /* styling isn't load-bearing */ }

  SpreadsheetApp.flush();
  return {
    success: true,
    maxCols: sheet.getMaxColumns(),
    lastCol: sheet.getLastColumn(),
    col24:   String(sheet.getRange(1, 24).getValue() || ''),
    col25:   String(sheet.getRange(1, 25).getValue() || '')
  };
}

// ============================================================
// restoreGroupsFromSheet — one-shot repair for the M_Category Group
// dropdown. Reads whatever Group names you've actually used in M_Category
// right now, unions them with the built-in PNL_GROUPS list, and reapplies
// the Group column validator with that union. Also sets allowInvalid=true
// so a historical typo you haven't fixed yet doesn't block edits elsewhere
// in the row. Dashboard P&L is live-resolve, so any group present here is
// a real bucket on the dashboard too.
//
// Run from the Apps Script editor. Idempotent.
// ============================================================
function restoreGroupsFromSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(M.CATEGORY);
  if (!sheet || sheet.getLastRow() < 2){
    var m0 = 'M_Category is empty — nothing to restore.';
    try { SpreadsheetApp.getUi().alert(m0); } catch(e){ Logger.log(m0); }
    return {success:false, message:m0};
  }

  // Union: built-in PNL_GROUPS + every distinct non-blank group already in use.
  var seen = {}, order = [];
  function _add(g){
    g = String(g || '').trim();
    if (!g) return;
    if (!seen[g]){ seen[g] = true; order.push(g); }
  }
  PNL_GROUPS.forEach(_add);
  sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues()
       .forEach(function(r){ _add(r[0]); });

  // Reapply the Group column validator with the merged list + soft-invalid
  // so pre-existing data doesn't get flagged red.
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(order, true).setAllowInvalid(true)
    .setHelpText('Groups come from the built-in list plus any you\'ve used in M_Category. '
               + 'The dashboard P&L bucket uses whatever is in this cell.')
    .build();
  sheet.getRange(2, 2, Math.max(sheet.getMaxRows()-1, 1), 1).setDataValidation(rule);

  var added = order.filter(function(g){ return PNL_GROUPS.indexOf(g) === -1; });
  var msg = 'Group dropdown restored.\n'
          + 'Built-in groups: ' + PNL_GROUPS.length + '\n'
          + 'Your added groups: ' + added.length + (added.length ? ' — ' + added.join(', ') : '')
          + '\nTotal in dropdown now: ' + order.length
          + '\n\nThe dropdown now allows any value (soft validator) so earlier entries aren\'t blocked.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, total:order.length, added:added, builtIn:PNL_GROUPS.length};
}