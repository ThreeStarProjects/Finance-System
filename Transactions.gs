// ============================================================
// 3 Star Projects — Unified Finance System
// PHASE 2 : TRANSACTION ENGINE
// File: Transactions.gs   (depends on Masters.gs)
//
// Builds the core money ledger:
//   Transactions_<FY>     — one tab per financial year (Apr–Mar).
//                           Income + expense in one table. IN/OUT direction.
//                           Location is its OWN column. GST-ready columns.
//   Personal_Transactions — separate personal-spend ledger.
//
// Auto-drawing rule: a PERSONAL expense paid from a COMPANY account writes
//   (1) the personal row, AND (2) a paired "Owner Drawing" row in the business
//   books so company cash stays accurate. The two are linked by Linked_Ref.
//   Drawings are class=Drawing → excluded from P&L, included in cash.
//
// Entry point (optional): run setupTransactionInfra() to pre-create the
//   Personal_Transactions tab + the current-FY Transactions tab. FY tabs are
//   otherwise auto-created on first write to that year.
// ============================================================

var T = {
  PERSONAL: 'Personal_Transactions',
  TXN_PREFIX: 'Transactions_'      // + FY, e.g. Transactions_2025-26
};

var DRAWING_CATEGORY = 'Owner Drawing';   // reserved business-books class=Drawing

// ── Transactions column map (0-based) ───────────────────────
var TC = {
  ID:0, ENTRY_TYPE:1, CLASS:2, DIRECTION:3, SUB_DATE:4, TX_DATE:5, MONTH:6, FY:7,
  CATEGORY:8, GROUP:9, SUBMITTED_BY:10, ON_BEHALF_OF:11, PROJECT:12,
  PARTY_ID:13, PARTY_NAME:14, DESCRIPTION:15, LOCATION:16, NOTES:17,
  AMOUNT:18, TAXABLE:19, GST_RATE:20, CGST:21, SGST:22, IGST:23, HSN_SAC:24,
  PAY_MODE:25, SOURCE:26, PROOF:27,
  SETTLE_STATUS:28, SETTLE_DATE:29, SETTLE_AMT:30, SETTLE_MODE:31,
  RECOVERY:32, RECOVERY_REF:33, LINKED_REF:34, DELETED:35,
  // Tier 4.2 — partial advance recovery ledger
  RECOVERED_AMT:36
};
var TC_TOTAL = 37;

var TXN_HEADERS = [
  'Txn_ID','Entry_Type','Class','Direction','Submission_Date','Txn_Date','Month','FY',
  'Category','Group','Submitted_By','On_Behalf_Of','Project',
  'Party_ID','Party_Name','Description','Location','Notes',
  'Amount','Taxable_Value','GST_Rate','CGST','SGST','IGST','HSN_SAC',
  'Payment_Mode','Source_of_Funds','Proof_Link',
  'Settle_Status','Settle_Date','Settle_Amount','Settle_Mode',
  'Recovery_Status','Recovery_Ref','Linked_Ref','Deleted',
  'Recovered_Amount'
];
var TXN_WIDTHS = [150,90,90,80,150,120,75,80,260,160,140,140,180,110,200,240,150,200,
                  110,120,80,90,90,90,110,130,200,200,120,120,120,110,120,180,150,70,
                  130];

// ── Personal_Transactions column map ────────────────────────
// Tier 5.1 extended the schema from 13 → 18 cols to hold owner-side transfers
// and per-owner attribution. Column order documented below.
var PC = {
  ID:0, SUB_DATE:1, TX_DATE:2, MONTH:3, FY:4,
  OWNER:5, CLASS:6, DIRECTION:7,
  CATEGORY:8, PARTY_NAME:9, DESCRIPTION:10,
  AMOUNT:11, SOURCE:12, COUNTER_ACCOUNT:13, PAY_MODE:14,
  NOTES:15, LINKED_REF:16, DELETED:17
};
var PC_TOTAL = 18;
var PERSONAL_HEADERS = [
  'Ptxn_ID','Submission_Date','Txn_Date','Month','FY',
  'Owner','Class','Direction',
  'Category','Party_Name','Description',
  'Amount','Source_of_Funds','Counter_Account','Payment_Mode',
  'Notes','Linked_Ref','Deleted'
];
var PERSONAL_WIDTHS = [150,150,120,75,80, 110,110,80, 240,220,240, 110,200,200,130, 220,150,70];

// Personal class labels — kept in one place so mirror/read code stays consistent.
var PCLASS = { EXPENSE:'Expense', FUND_IN:'Fund In', FUND_OUT:'Fund Out', INTERNAL:'Internal' };

// ── Project_Costs column map (Batch 14: + GST for Simple/Material) ─────
// Entry Type: 'Simple' | 'Material' | 'ManHours'
// Subcategory from M_Category (c.category) → Category auto-filled (c.group)
// Amount column stores the final INC-GST total; Amount_Ex_GST is the base.
var PCC = {
  ID:0, ENTRY_DATE:1, PROJECT:2, ENTRY_TYPE:3,
  SUBCATEGORY:4, CATEGORY:5,
  DESCRIPTION:6, AMOUNT:7,
  // Material-only fields (populated only when Entry_Type='Material')
  MATERIAL_NAME:8, QUANTITY:9, UNIT:10, RATE_PER_UNIT:11,
  NOTES:12, STATUS:13, SUBMITTED_BY:14, SUBMISSION_DATE:15, DELETED:16,
  // Batch 14: GST tracking (Simple + Material only, optional)
  GST_RATE:17, AMOUNT_EX_GST:18
};
var PCC_TOTAL = 19;
var PROJECT_COSTS_HEADERS = [
  'Cost_ID','Entry_Date','Project','Entry_Type',
  'Subcategory','Category',
  'Description','Amount',
  'Material_Name','Quantity','Unit','Rate_Per_Unit',
  'Notes','Status','Submitted_By','Submission_Date','Deleted',
  'GST_Rate','Amount_Ex_GST'
];
var PROJECT_COSTS_WIDTHS = [150,120,180,100,
                             200,180,
                             260,130,
                             200,110,80,130,
                             220,100,140,150,70,
                             90,130];

// ============================================================
// DATE / FY HELPERS  (Indian FY: Apr 1 – Mar 31)
// ============================================================
var _MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function _asDate(v){
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? new Date() : d;
}
// "2025-26" for any date in that FY
function _fyFromDate(date){
  var d = _asDate(date), y = d.getFullYear(), m = d.getMonth(); // 0=Jan
  var start = (m >= 3) ? y : y - 1;                              // Apr=3
  return start + '-' + String(start + 1).substring(2);
}
function _monthOf(date){
  var d = _asDate(date);
  return _MONTHS[d.getMonth()] + '-' + String(d.getFullYear()).substring(2);
}
function _currentMonth(){ return _monthOf(new Date()); }
function _currentFY(){ return _fyFromDate(new Date()); }
function _txnSheetName(fy){ return T.TXN_PREFIX + fy; }

// ============================================================
// SHEET PROVISIONING
// ============================================================
// Get (or create + format) the Transactions tab for the FY of `date`.
function _getTxnSheet(date){ return _getTxnSheetByFy(_fyFromDate(date)); }

// Get (or create + fully provision) the Transactions tab for an explicit FY string.
function _getTxnSheetByFy(fy){
  var ss = _getSS();
  var name = _txnSheetName(fy);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // Create + format a fresh FY tab
  sheet = _buildSheet(ss, name, TXN_HEADERS, TXN_WIDTHS);
  sheet.getRange(2, TC.MONTH+1, 5000, 1).setNumberFormat('@STRING@'); // keep MMM-yy as text
  _validateCol(sheet, TC.ENTRY_TYPE+1,   ['Employee','Owner']);
  _validateCol(sheet, TC.CLASS+1,        ['Expense','Income','Drawing','Transfer']);
  _validateCol(sheet, TC.DIRECTION+1,    ['IN','OUT']);
  _validateCol(sheet, TC.SETTLE_STATUS+1,['Pending','Settled','N/A']);
  _validateCol(sheet, TC.RECOVERY+1,     ['Pending','Recovered','N/A']);
  _validateCol(sheet, TC.DELETED+1,      ['Yes','No']);
  _applyTxnDropdownsLive(sheet);   // live Category from M_Category; Group has no dropdown (auto-filled by onEdit)
  _applyTxnFormatting(sheet);
  return sheet;
}

// Create a transaction tab for a past/future FY (for bulk-pasting historical data).
// Run from the editor — it asks for the financial year.
function createTransactionTab(){
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e){ Logger.log('Run this from the spreadsheet.'); return; }
  var res = ui.prompt('Create a transaction tab',
    'Enter the financial year to create (format: 2025-26):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var fy = res.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(fy)){ ui.alert('Please use the format 2025-26 (four digits, dash, two digits).'); return; }
  var sheet = _getTxnSheetByFy(fy);
  ui.alert('Ready: '+sheet.getName()+'\n\nPaste your '+fy+' data here, fill the Date and Category columns, then run processBulkPaste to fill Groups, Party_IDs, and refresh directories.');
}

// Soft dropdowns (allow other values too — so bulk-pasting old data isn't blocked,
// but you still get a picker when editing Category / Group / etc. in the sheet).
var EXPENSE_LOCATIONS = ['On-site','Factory (Dombivli)','Office (Thane)','Vendor','Others'];
var GST_RATES         = ['0','5','12','18','28'];
var FY_CHOICES        = ['2023-24','2024-25','2025-26','2026-27','2027-28','2028-29','2029-30'];

function _validateColSoft(sheet, col, values){
  if (!sheet || !values || !values.length) return;
  var n = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, col, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build()
  );
}
function _validateColDateSoft(sheet, col){
  if (!sheet) return;
  var n = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, col, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
  );
}
function _validateColNumberSoft(sheet, col){
  if (!sheet) return;
  var n = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, col, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(true).build()
  );
}

function _applyTxnDropdowns(sheet){
  if (!sheet || typeof sheet.getRange !== 'function') return;
  // Collect master lists (each in its own try so one failure doesn't kill the rest)
  var cats = [];
  try { (getCategories()||[]).forEach(function(c){ if(c && c.category && cats.indexOf(c.category)===-1) cats.push(c.category); }); } catch(e){}
  try { (getEmployeeCategories()||[]).forEach(function(c){ if(c && c.category && cats.indexOf(c.category)===-1) cats.push(c.category); }); } catch(e){}
  var pmodes=[], sources=[], employees=[], projects=[];
  try { pmodes    = getPaymentModes()||[]; } catch(e){}
  try { sources   = (getSourcesOfFunds()||[]).map(function(s){ return s && s.value; }).filter(Boolean); } catch(e){}
  try { employees = (getEmployeeNames()||[]).slice(); } catch(e){}
  try { projects  = (getProjectNames()||[]).slice(); } catch(e){}
  // Always include 'Owner' in Submitted_By / On_Behalf_Of so direct-pasted owner rows pick cleanly
  var submitters = ['Owner'].concat(employees);
  var onBehalf   = ['Owner','Company'].concat(employees);

  // List dropdowns
  var listJobs = [
    [TC.CATEGORY+1,     cats],
    [TC.GROUP+1,        PNL_GROUPS],
    [TC.SUBMITTED_BY+1, submitters],
    [TC.ON_BEHALF_OF+1, onBehalf],
    [TC.PROJECT+1,      projects],
    [TC.LOCATION+1,     EXPENSE_LOCATIONS],
    [TC.GST_RATE+1,     GST_RATES],
    [TC.PAY_MODE+1,     pmodes],
    [TC.SOURCE+1,       sources],
    [TC.SETTLE_MODE+1,  pmodes],
    [TC.FY+1,           FY_CHOICES]
  ];
  listJobs.forEach(function(p){ try { _validateColSoft(sheet, p[0], p[1]); } catch(e){ Logger.log('list-dropdown col '+p[0]+': '+e.message); } });

  // Date validators — prevents the text-date bug that broke FY formulas
  [TC.TX_DATE+1, TC.SUB_DATE+1, TC.SETTLE_DATE+1].forEach(function(c){
    try { _validateColDateSoft(sheet, c); } catch(e){ Logger.log('date-validator col '+c+': '+e.message); }
  });

  // Number validators — prevents accidental text in numeric columns
  [TC.AMOUNT+1, TC.TAXABLE+1, TC.CGST+1, TC.SGST+1, TC.IGST+1, TC.SETTLE_AMT+1].forEach(function(c){
    try { _validateColNumberSoft(sheet, c); } catch(e){ Logger.log('number-validator col '+c+': '+e.message); }
  });
}

function _applyTxnFormatting(sheet){
  var rng = sheet.getRange(2,1,5000,TC_TOTAL);
  var rb  = SpreadsheetApp.newConditionalFormatRule;
  sheet.setConditionalFormatRules([
    // Deleted → grey italic (wins)
    rb().whenFormulaSatisfied('=$AJ2="Yes"').setBackground('#F1F5F9').setFontColor('#94A3B8').setItalic(true).setRanges([rng]).build(),
    // Income → light mint
    rb().whenFormulaSatisfied('=$C2="Income"').setBackground('#D1FAE5').setRanges([rng]).build(),
    // Drawing → light amber
    rb().whenFormulaSatisfied('=$C2="Drawing"').setBackground('#FEF3C7').setRanges([rng]).build(),
    // Employee expense pending → light orange
    rb().whenFormulaSatisfied('=AND($B2="Employee",$AC2="Pending")').setBackground('#FFEDD5').setRanges([rng]).build(),
    // Owner expense → light blue
    rb().whenFormulaSatisfied('=$B2="Owner"').setBackground('#E5EEF5').setRanges([rng]).build()
  ]);
}

function _getPersonalSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(T.PERSONAL);
  if (sheet) return sheet;
  sheet = _buildSheet(ss, T.PERSONAL, PERSONAL_HEADERS, PERSONAL_WIDTHS);
  sheet.getRange(2, PC.MONTH+1, 5000, 1).setNumberFormat('@STRING@');
  _validateCol(sheet, PC.DELETED+1, ['Yes','No']);
  return sheet;
}

// ============================================================
// PROJECT COSTS SHEET  (Batch 10 — category-based with Simple/Material entry)
// Entry Type: Simple or Material; Subcategory from M_Category auto-fills Category
// ============================================================
function _getProjectCostsSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName('Project_Costs');
  if (sheet) return sheet;
  sheet = _buildSheet(ss, 'Project_Costs', PROJECT_COSTS_HEADERS, PROJECT_COSTS_WIDTHS);
  // Date format
  sheet.getRange(2, PCC.ENTRY_DATE+1, 5000, 1).setNumberFormat('yyyy-MM-dd');
  sheet.getRange(2, PCC.SUBMISSION_DATE+1, 5000, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  // Validation
  _validateCol(sheet, PCC.ENTRY_TYPE+1, ['Simple','Material','ManHours']);
  _validateCol(sheet, PCC.STATUS+1, ['Pending','Approved','Rejected']);
  _validateCol(sheet, PCC.DELETED+1, ['Yes','No']);
  // Conditional formatting: Pending entries → light yellow
  var rng = sheet.getRange(2,1,5000,PCC_TOTAL);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$N2="Pending"')
      .setBackground('#FFFACD')
      .setRanges([rng])
      .build()
  ]);
  return sheet;
}

// Reset Project_Costs sheet (deletes old sheet and recreates with new structure)
// Run this once after deploying the Batch 10 restructure
function resetProjectCostsSheet(){
  var ss = _getSS();
  var old = ss.getSheetByName('Project_Costs');
  if (old) ss.deleteSheet(old);
  _getProjectCostsSheet();
  return {success:true, message:'Project_Costs sheet rebuilt with new structure (' + PCC_TOTAL + ' columns)'};
}

// Batch 14: Non-destructive upgrade — add GST_Rate + Amount_Ex_GST columns to
// an existing Project_Costs sheet without touching data. Run once from the
// Apps Script editor after deploying.
function upgradeProjectCostsSchema(){
  var ss = _getSS();
  var sheet = ss.getSheetByName('Project_Costs');
  if (!sheet) return {success:false, message:'Project_Costs sheet not found. Run setupTransactionInfra() first.'};
  var lastCol = sheet.getLastColumn();
  if (lastCol >= PCC_TOTAL) return {success:true, message:'Already up to date ('+lastCol+' columns).'};
  // Write headers for the two new columns
  sheet.getRange(1, PCC.GST_RATE+1).setValue('GST_Rate').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#0F172A');
  sheet.getRange(1, PCC.AMOUNT_EX_GST+1).setValue('Amount_Ex_GST').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#0F172A');
  sheet.setColumnWidth(PCC.GST_RATE+1, 90);
  sheet.setColumnWidth(PCC.AMOUNT_EX_GST+1, 130);
  // Backfill Amount_Ex_GST = Amount for existing rows (no GST was tracked before)
  if (sheet.getLastRow() > 1){
    var n = sheet.getLastRow() - 1;
    var amts = sheet.getRange(2, PCC.AMOUNT+1, n, 1).getValues();
    sheet.getRange(2, PCC.AMOUNT_EX_GST+1, n, 1).setValues(amts);
  }
  SpreadsheetApp.flush();
  return {success:true, message:'Added GST_Rate + Amount_Ex_GST columns. Existing rows backfilled (Amount_Ex_GST = Amount).'};
}

// URL of the backing Spreadsheet (for the "Open Sheet" button)
function getSheetUrl(){
  try { return _getSS().getUrl(); } catch(e){ return ''; }
}

function _nextCostId(){
  var sheet = _getProjectCostsSheet();
  if (!sheet || !_hasData(sheet)) return 'COST-' + Utilities.formatDate(new Date(),'GMT+5:30','yyyyMMdd') + '-0001';
  var ids = sheet.getRange(2, PCC.ID+1, sheet.getLastRow()-1, 1).getValues();
  var today = Utilities.formatDate(new Date(),'GMT+5:30','yyyyMMdd');
  var max = 0;
  for (var i=0; i<ids.length; i++){
    var v = _s(ids[i][0]);
    if (v.indexOf('COST-' + today) === 0){
      var n = parseInt(v.substring(18), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return 'COST-' + today + '-' + String(max + 1).padStart(4, '0');
}

// ============================================================
// SUBMIT PROJECT COSTS  (Batch 7 — bulk entry from form)
// ============================================================
function submitProjectCosts(entries){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    if (!entries || !entries.length) return {success:false, message:'No entries provided.'};
    var sheet = _getProjectCostsSheet();
    var now = new Date();
    var subDate = Utilities.formatDate(now, 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
    var submitted = [];

    // Build category→group map once for auto-fill
    var catGroupMap = {};
    try { getCategories().forEach(function(c){ if(c.category) catGroupMap[c.category] = c.group || ''; }); } catch(e){}

    entries.forEach(function(entry){
      var costId = _nextCostId();
      var row = new Array(PCC_TOTAL).fill('');
      var data = entry.data || {};
      var entryType = _s(entry.entry_type || 'Simple');
      var subcat = _s(data.subcategory || entry.subcategory || '');
      var category = _s(data.category || entry.category || catGroupMap[subcat] || '');

      row[PCC.ID] = costId;
      row[PCC.ENTRY_DATE] = _s(entry.entry_date || data.date || Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd'));
      row[PCC.PROJECT] = _s(entry.project);
      row[PCC.ENTRY_TYPE] = entryType;
      row[PCC.SUBCATEGORY] = subcat;
      row[PCC.CATEGORY] = category;
      row[PCC.DESCRIPTION] = _s(data.description || '');
      row[PCC.NOTES] = _s(data.notes || '');
      row[PCC.STATUS] = 'Pending';
      row[PCC.SUBMITTED_BY] = _s(entry.submitted_by || 'Finance');
      row[PCC.SUBMISSION_DATE] = subDate;
      row[PCC.DELETED] = 'No';

      // GST (optional, Simple + Material only). gstRate=0 → no GST applied.
      var gstRate = _n(data.gst_rate) || 0;
      var gstMult = 1 + (gstRate / 100);

      if (entryType === 'Material'){
        var qty = _n(data.quantity) || 0;
        var rate = _n(data.rate) || 0;
        var baseAmt = qty * rate;
        row[PCC.MATERIAL_NAME] = _s(data.materialName || '');
        row[PCC.QUANTITY] = qty;
        row[PCC.UNIT] = _s(data.unit || '');
        row[PCC.RATE_PER_UNIT] = rate;
        row[PCC.AMOUNT_EX_GST] = baseAmt;
        row[PCC.GST_RATE] = gstRate > 0 ? gstRate : '';
        row[PCC.AMOUNT] = baseAmt * gstMult;  // final Inc-GST
      } else if (entryType === 'ManHours'){
        // Labor cost = hours × (daily_rate / 8), or hours × (gross_monthly / 240) when no daily rate.
        // Employee name is normalized against M_Employee master (typo guard).
        var empName = _s(data.employee || data.materialName || '');
        if (typeof _resolveEmployeeName === 'function'){
          var _r = _resolveEmployeeName(empName);
          if (!_r.ok) throw new Error('Row for ' + empName + ': ' + _r.message);
          empName = _r.name;
        }
        var hours = _n(data.hours || data.quantity) || 0;
        if (hours <= 0) throw new Error('Hours must be > 0 for ' + empName);
        // Pull current rate from master so it reflects any salary change.
        var emp = (typeof _findEmployee === 'function') ? _findEmployee(empName) : null;
        var dailyRate = emp ? _n(emp.dailyRate) : 0;
        var hourlyRate = dailyRate > 0
          ? (dailyRate / 8)
          : (emp ? (_n(emp.grossSalary) / 240) : 0);
        row[PCC.MATERIAL_NAME] = empName;
        row[PCC.QUANTITY] = hours;
        row[PCC.UNIT] = 'Hours';
        row[PCC.RATE_PER_UNIT] = hourlyRate;
        var mhTotal = hours * hourlyRate;
        row[PCC.AMOUNT_EX_GST] = mhTotal;  // own-labour: no GST
        row[PCC.GST_RATE] = '';
        row[PCC.AMOUNT] = mhTotal;
        if (!row[PCC.DESCRIPTION]){
          row[PCC.DESCRIPTION] = empName + ' · ' + hours + 'h @ ₹' + hourlyRate.toFixed(2) + '/h';
        }
      } else {
        // Simple entry: data.amount is the BASE (ex-GST)
        var simpleBase = _n(data.amount) || 0;
        row[PCC.AMOUNT_EX_GST] = simpleBase;
        row[PCC.GST_RATE] = gstRate > 0 ? gstRate : '';
        row[PCC.AMOUNT] = simpleBase * gstMult;
      }

      sheet.appendRow(row);
      submitted.push(costId);
    });

    SpreadsheetApp.flush();
    return {success:true, count:submitted.length, ids:submitted};
  } catch(err){
    return {success:false, message:err.message};
  }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// QUERY FUNCTIONS — Project Cost Reporting (Batch 8)
// ============================================================
// Combine data from Transactions_<FY> and Project_Costs sheets
// for project-wise cost analysis and dashboard.

function getProjectCostsRows(project, approvedOnly){
  // Returns: {success:true, rows:[{id, date, entryType, subcategory, category, description, amount, materialName, qty, unit, rate, notes, status}], total}
  if(!project) return {success:false, error: 'Project required', rows:[], total:0};
  if(approvedOnly === undefined) approvedOnly = false;  // default: show all including Pending
  var sheet = _getProjectCostsSheet();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  var total = 0;

  for(var i=1; i<data.length; i++){
    var r = data[i];
    if(_s(r[PCC.DELETED]) === 'Yes') continue;
    if(_s(r[PCC.PROJECT]) !== project) continue;
    if(approvedOnly && _s(r[PCC.STATUS]) !== 'Approved') continue;
    var amt = _n(r[PCC.AMOUNT]) || 0;
    var hasGst = (r.length > PCC.GST_RATE) ? _n(r[PCC.GST_RATE]) : 0;
    rows.push({
      id: _s(r[PCC.ID]),
      date: _safeDateStr(r[PCC.ENTRY_DATE]),
      entryType: _s(r[PCC.ENTRY_TYPE]),
      subcategory: _s(r[PCC.SUBCATEGORY]),
      category: _s(r[PCC.CATEGORY]),
      description: _s(r[PCC.DESCRIPTION]),
      amount: amt,
      amountExGst: (r.length > PCC.AMOUNT_EX_GST) ? (_n(r[PCC.AMOUNT_EX_GST]) || 0) : amt,
      gstRate: hasGst,
      materialName: _s(r[PCC.MATERIAL_NAME]),
      qty: _n(r[PCC.QUANTITY]) || 0,
      unit: _s(r[PCC.UNIT]),
      rate: _n(r[PCC.RATE_PER_UNIT]) || 0,
      notes: _s(r[PCC.NOTES]),
      status: _s(r[PCC.STATUS])
    });
    total += amt;
  }
  // Sort newest first
  rows.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
  return {success:true, rows:rows, total:total, count:rows.length};
}

// Back-compat shim: old UI called getProjectCostsSummary — now returns a lightweight total
function getProjectCostsSummary(project, approvedOnly){
  var res = getProjectCostsRows(project, approvedOnly);
  return {success:res.success, total:res.total, count:res.count};
}

// Returns categories list with their groups, for dropdown population
function getCategoriesForProjectCosts(){
  try {
    var cats = getCategories() || [];
    return {success:true, categories: cats.filter(function(c){ return c && c.category; }).map(function(c){
      return {subcategory: c.category, category: c.group || ''};
    })};
  } catch(e){
    return {success:false, error:e.message, categories:[]};
  }
}

function getProjectTransactionsSummary(project){
  // Returns transactions (invoices/bills/expenses) for this project
  // {success:true, total, byType: {Invoice: {...}, Bill: {...}, Expense: {...}}}
  if(!project) return {success:false, error: 'Project required'};
  var sheet = _getTxnSheet(new Date());
  var data = sheet.getDataRange().getValues();
  var summary = {success:true, total: 0, byType: {}};

  for(var i=1; i<data.length; i++){
    var row = data[i];
    if(_s(row[TC.DELETED])==='Yes' || _s(row[TC.PROJECT])!==project) continue;
    var type = _s(row[TC.ENTRY_TYPE]); // Invoice, Bill, Expense, etc.
    var amt = _n(row[TC.AMOUNT])||0;
    summary.total += amt;
    if(!summary.byType[type]) summary.byType[type] = {count:0, amount:0};
    summary.byType[type].count++;
    summary.byType[type].amount += amt;
  }
  return summary;
}

function getProjectTotalCosts(project, approvedOnly){
  // Returns: {success:true, projectCosts, transactionCosts, combined, breakdown}
  if(!project) return {success:false, error: 'Project required'};
  if(approvedOnly === undefined) approvedOnly = true;
  var pc = getProjectCostsSummary(project, approvedOnly);
  var tx = getProjectTransactionsSummary(project);
  return {
    success: true,
    projectCosts: pc.total,
    transactionCosts: tx.total,
    combined: pc.total + tx.total,
    projectCostsByType: pc.byType,
    transactionsByType: tx.byType
  };
}

// Optional pre-create
function setupTransactionInfra(){
  _getPersonalSheet();
  _getTxnSheet(new Date());
  _getProjectCostsSheet();
  // Also ensure the prior-FY transaction tabs exist (empty), so old data
  // can be pasted in without first creating a transaction in that FY.
  var priorFys = ensurePriorFyTransactionSheets(['2024-25', '2025-26']);
  var msg = 'Transaction engine ready:\n• ' + _txnSheetName(_currentFY()) + '\n• ' + T.PERSONAL + '\n• Project_Costs';
  if (priorFys.length) msg += '\nPrior FY tabs ensured:\n• ' + priorFys.join('\n• ');
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, message:msg};
}

// Ensure a transaction sheet exists for each FY in the list.
// Uses the same _getTxnSheetByFy builder so headers/formatting match the
// current-FY tab exactly. Returns the list of sheet names ensured.
//   Safe to call repeatedly — does nothing if the tab already exists.
function ensurePriorFyTransactionSheets(fyList){
  if (!fyList || !fyList.length) return [];
  var out = [];
  for (var i=0; i<fyList.length; i++){
    var fy = _s(fyList[i]);
    if (!fy) continue;
    // _getTxnSheetByFy creates the sheet if missing, returns existing if present.
    var sheet = _getTxnSheetByFy(fy);
    if (sheet) out.push(sheet.getName());
  }
  return out;
}

// ── ID generator: TXN-yyyymmdd-#### (unique within FY tab) ──
function _nextTxnId(sheet, date){
  var stamp = Utilities.formatDate(_asDate(date),'GMT+5:30','yyyyMMdd');
  var n = sheet.getLastRow();   // row count ≈ running counter; ID stays unique per tab
  return 'TXN-' + stamp + '-' + (n + 1);
}
function _nextPersonalId(sheet, date){
  var stamp = Utilities.formatDate(_asDate(date),'GMT+5:30','yyyyMMdd');
  return 'PTX-' + stamp + '-' + (sheet.getLastRow() + 1);
}

// Look up a category's group (so P&L group is stored on the row, not recomputed).
function _groupForCategory(category){
  var c = _s(category);
  var cats = getCategories();
  for (var i=0;i<cats.length;i++){ if (cats[i].category === c) return cats[i].group; }
  var emp = getEmployeeCategories();
  for (var j=0;j<emp.length;j++){ if (emp[j].category === c) return emp[j].group; }
  return '';
}

// Build a blank 36-wide row, then fill named cells.
function _blankTxnRow(){ var r = []; for (var i=0;i<TC_TOTAL;i++) r.push(''); return r; }

// ============================================================
// SUBMIT — EMPLOYEE EXPENSE
//   formData = {employeeName, expenses:[{date, amount, category, project,
//     party, partyId, description, location, notes, paymentMode, paidBy,
//     taxable, gstRate, cgst, sgst, igst, hsnSac, proofData}]}
//   paidBy 'Company' → settled immediately; 'Employee' → pending reimbursement.
// ============================================================
function submitEmployeeExpense(formData){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){
    _logSubmission(formData.employeeName, formData.expenses, {success:false, message:'Busy — try again.'});
    return {success:false, message:'Busy — try again.'};
  }
  try {
    var now = new Date();
    var subDate = Utilities.formatDate(now,'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var saved = [];

    // ── Float-aware batch reconciliation ──────────────────────────
    // formData.paidFromAccount (optional) = a float account ID (FL-xx).
    // When set, expenses are paid from that float until the float runs out;
    // the excess becomes Pending Settlement (employee paid personally).
    var paidFromId = _s(formData.paidFromAccount);
    var floatBalanceRemaining = 0;
    var floatName = '';
    if (paidFromId){
      var floatAcc = (typeof getAccountByIdOrName === 'function') ? getAccountByIdOrName(paidFromId) : null;
      if (!floatAcc || floatAcc.side !== 'Float'){
        return {success:false, message:'Selected float account not found or not a float account.'};
      }
      floatName = floatAcc.name;
      // Get its current balance
      var bal = (typeof getFloatBalances === 'function') ? getFloatBalances() : null;
      if (bal && bal.ok && bal.data && bal.data.rows){
        for (var b=0; b<bal.data.rows.length; b++){
          if (bal.data.rows[b].id === floatAcc.id){
            floatBalanceRemaining = bal.data.rows[b].balance;
            break;
          }
        }
      }
    }

    (formData.expenses || []).forEach(function(exp){
      var txDate = exp.date ? _asDate(exp.date) : now;
      var sheet  = _getTxnSheet(txDate);                 // FY-correct tab
      var id     = _nextTxnId(sheet, now);
      var isCompany = (_s(exp.paidBy) === 'Company');

      var proof = '';
      if (exp.proofData && exp.proofData.data){
        try { proof = _saveProof(exp.proofData, exp.project, txDate, id); } catch(e){}
      }

      var amt = _n(exp.amount);

      // Decide settle status for this expense:
      //   - Company-paid (existing flow):     Settled
      //   - Paid from float AND float covers it: Settled, source = float account
      //   - Paid from float BUT float empty:  Pending (excess goes to employee)
      //   - No float specified (default):     Pending (or Settled if Company)
      var settleStatus, settleDate, settleAmt, settleMode, source;
      var rowPaidFromFloat = false;

      if (paidFromId && floatBalanceRemaining >= amt){
        // Float covers this expense in full
        settleStatus = 'Settled';
        settleDate   = subDate;
        settleAmt    = amt;
        settleMode   = _s(exp.paymentMode) || 'Cash';
        source       = paidFromId;       // record float account ID in Source
        floatBalanceRemaining -= amt;
        rowPaidFromFloat = true;
      } else if (isCompany){
        settleStatus = 'Settled';
        settleDate   = subDate;
        settleAmt    = amt;
        settleMode   = _s(exp.paymentMode) || 'Cash';
        source       = _s(exp.source);
      } else {
        // Pending — employee paid personally, owed money back
        settleStatus = 'Pending';
        settleDate   = '';
        settleAmt    = '';
        settleMode   = '';
        source       = _s(exp.source);
      }

      var row = _blankTxnRow();
      row[TC.ID]           = id;
      row[TC.ENTRY_TYPE]   = 'Employee';
      row[TC.CLASS]        = 'Expense';
      row[TC.DIRECTION]    = 'OUT';
      row[TC.SUB_DATE]     = subDate;
      row[TC.TX_DATE]      = exp.date || subDate;
      row[TC.MONTH]        = _monthOf(txDate);
      row[TC.FY]           = _fyFromDate(txDate);
      row[TC.CATEGORY]     = _s(exp.category);
      row[TC.GROUP]        = _groupForCategory(exp.category);
      row[TC.SUBMITTED_BY] = _s(formData.employeeName);
      row[TC.PROJECT]      = _s(exp.project);
      var _pid = _s(exp.partyId), _pnm = _s(exp.party);
      if (!_pid && _pnm){ var _pr = findOrCreateParty('Vendor', _pnm); _pid = _pr.id; _pnm = _pr.name; }
      row[TC.PARTY_ID]     = _pid;
      row[TC.PARTY_NAME]   = _pnm;
      row[TC.DESCRIPTION]  = _s(exp.description);
      row[TC.LOCATION]     = _s(exp.location);            // OWN COLUMN
      // Tag float-paid expenses in notes so the audit trail is obvious
      row[TC.NOTES]        = rowPaidFromFloat
                              ? ('[Paid from ' + floatName + '] ' + _s(exp.notes)).trim()
                              : _s(exp.notes);
      row[TC.AMOUNT]       = amt;
      row[TC.TAXABLE]      = exp.taxable !== undefined && exp.taxable !== '' ? _n(exp.taxable) : '';
      row[TC.GST_RATE]     = _s(exp.gstRate);
      row[TC.CGST]         = exp.cgst !== undefined && exp.cgst !== '' ? _n(exp.cgst) : '';
      row[TC.SGST]         = exp.sgst !== undefined && exp.sgst !== '' ? _n(exp.sgst) : '';
      row[TC.IGST]         = exp.igst !== undefined && exp.igst !== '' ? _n(exp.igst) : '';
      row[TC.HSN_SAC]      = _s(exp.hsnSac);
      row[TC.PAY_MODE]     = _s(exp.paymentMode) || 'Cash';
      row[TC.SOURCE]       = source;
      row[TC.PROOF]        = proof;
      row[TC.SETTLE_STATUS]= settleStatus;
      row[TC.SETTLE_DATE]  = settleDate;
      row[TC.SETTLE_AMT]   = settleAmt;
      row[TC.SETTLE_MODE]  = settleMode;
      row[TC.RECOVERY]     = 'N/A';
      row[TC.DELETED]      = 'No';

      sheet.appendRow(row);
      SpreadsheetApp.flush();
      saved.push(id);
    });

    // Compose a friendlier message if float was used
    var msg = 'Recorded ' + saved.length + ' expense(s).';
    if (paidFromId){
      msg += ' Remaining ' + floatName + ' balance: ₹' + _round(floatBalanceRemaining) + '.';
      if (floatBalanceRemaining < 0){
        // The last expense overran the float — that row was tagged Pending
        msg += ' Some entries exceeded the float and were recorded as Pending Settlement.';
      }
    }

    // Tier 8.3 — brief confirmation email to the submitting employee.
    // Best-effort: mail failure never blocks the primary write.
    try {
      var confirmEmail = (typeof getEmployeeEmail === 'function') ? getEmployeeEmail(formData.employeeName) : null;
      if (confirmEmail){
        var statusLine;
        if (paidFromId){
          statusLine = (floatBalanceRemaining < 0)
            ? 'Some entries exceeded ' + floatName + ' — the excess is Pending Settlement.'
            : 'Settled from ' + floatName + '.';
        } else {
          var anyPending = false;
          (formData.expenses || []).forEach(function(x){ if (_s(x.paidBy) !== 'Company') anyPending = true; });
          statusLine = anyPending ? 'Reimbursement status: Pending owner approval.' : 'Settled by company account.';
        }
        _sendExpenseConfirmationEmail(confirmEmail, _s(formData.employeeName), formData.expenses, statusLine);
      }
    } catch(mailErr){ Logger.log('expense confirm email: ' + mailErr.message); }

    // Log successful submission
    _logSubmission(_s(formData.employeeName), formData.expenses, {success:true, ids:saved, count:saved.length});

    return {success:true, ids:saved, count:saved.length, message:msg, floatBalance:floatBalanceRemaining};
  } catch(err){
    _logSubmission(_s(formData.employeeName), formData.expenses, {success:false, message:err.message});
    return {success:false, message:err.message};
  }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// SUBMIT — OWNER (COMPANY) TRANSACTION  (income OR expense)
//   formData = {direction:'IN'|'OUT', category, customCategory, customGroup,
//     txDate, amount, paymentMode, source, description, onBehalfOf, project,
//     party, partyId, location, notes, taxable, gstRate, cgst, sgst, igst, hsnSac,
//     isAdvance(bool)}
// ============================================================
function submitOwnerTransaction(formData){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var now = new Date();
    var subDate = Utilities.formatDate(now,'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var txDate  = formData.txDate ? _asDate(formData.txDate) : now;
    var sheet   = _getTxnSheet(txDate);
    var id      = _nextTxnId(sheet, now);

    // Resolve category (handle "Other" → addCategory with forced group)
    var category = _s(formData.category);
    if (category === 'Other'){
      var add = addCategory({category:formData.customCategory, group:formData.customGroup,
                             direction:formData.direction === 'IN' ? 'IN' : 'OUT'});
      if (!add.success) return add;          // bubbles "group required" etc.
      category = add.category || _s(formData.customCategory);
    }
    var dir   = (_s(formData.direction) === 'IN') ? 'IN' : 'OUT';
    var isAdvance = !!formData.isAdvance;
    var klass;
    if (isAdvance) {
      klass = 'Transfer';
    } else {
      klass = (dir === 'IN') ? 'Income' : 'Expense';
    }

    // When flagging as advance, normalize onBehalfOf against M_Employee to
    // prevent typo'd names from creating orphaned advances invisible to the
    // salary calculator / dashboard. Non-advance rows keep onBehalfOf as-is
    // (owner names are tracked in a different master).
    var _onBehalfOf = _s(formData.onBehalfOf);
    if (isAdvance && _onBehalfOf && typeof _resolveEmployeeName === 'function'){
      var _res = _resolveEmployeeName(_onBehalfOf);
      if (!_res.ok) return {success:false, message:_res.message};
      _onBehalfOf = _res.name;
    }

    var row = _blankTxnRow();
    row[TC.ID]           = id;
    row[TC.ENTRY_TYPE]   = isAdvance ? 'Employee' : 'Owner';
    row[TC.CLASS]        = klass;
    row[TC.DIRECTION]    = dir;
    row[TC.SUB_DATE]     = subDate;
    row[TC.TX_DATE]      = formData.txDate || subDate;
    row[TC.MONTH]        = _monthOf(txDate);
    row[TC.FY]           = _fyFromDate(txDate);
    row[TC.CATEGORY]     = category;
    row[TC.GROUP]        = _groupForCategory(category);
    row[TC.SUBMITTED_BY] = 'Owner';
    row[TC.ON_BEHALF_OF] = _onBehalfOf;
    row[TC.PROJECT]      = _s(formData.project);
    var _pid = _s(formData.partyId), _pnm = _s(formData.party);
    if (!_pid && _pnm){ var _pt = (_s(formData.direction)==='IN')?'Customer':'Vendor'; var _pr = findOrCreateParty(_pt, _pnm); _pid = _pr.id; _pnm = _pr.name; }
    row[TC.PARTY_ID]     = _pid;
    row[TC.PARTY_NAME]   = _pnm;
    row[TC.DESCRIPTION]  = _s(formData.description);
    row[TC.LOCATION]     = _s(formData.location);
    row[TC.NOTES]        = _s(formData.notes);
    row[TC.AMOUNT]       = _n(formData.amount);
    row[TC.TAXABLE]      = formData.taxable !== undefined && formData.taxable !== '' ? _n(formData.taxable) : '';
    row[TC.GST_RATE]     = _s(formData.gstRate);
    row[TC.CGST]         = formData.cgst !== undefined && formData.cgst !== '' ? _n(formData.cgst) : '';
    row[TC.SGST]         = formData.sgst !== undefined && formData.sgst !== '' ? _n(formData.sgst) : '';
    row[TC.IGST]         = formData.igst !== undefined && formData.igst !== '' ? _n(formData.igst) : '';
    row[TC.HSN_SAC]      = _s(formData.hsnSac);
    row[TC.PAY_MODE]     = _s(formData.paymentMode);
    row[TC.SOURCE]       = _s(formData.source);
    row[TC.SETTLE_STATUS]= 'N/A';
    row[TC.RECOVERY]     = isAdvance ? 'Pending' : 'N/A';
    row[TC.DELETED]      = 'No';

    sheet.appendRow(row);
    SpreadsheetApp.flush();
    return {success:true, txId:id};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// FUND TRANSFERS SHEET  (Tier 6 — separate Transfer-class activity)
// Columns: ID, SUB_DATE, TX_DATE, MONTH, FY, ENTRY_TYPE, CLASS, DIRECTION,
//          CATEGORY, ON_BEHALF_OF, AMOUNT, FROM_ACCOUNT, TO_ACCOUNT, SOURCE,
//          DESCRIPTION, NOTES, PROOF, LINKED_REF, DELETED
// ============================================================
var FT = {
  ID:0, SUB_DATE:1, TX_DATE:2, MONTH:3, FY:4, ENTRY_TYPE:5, CLASS:6, DIRECTION:7,
  CATEGORY:8, ON_BEHALF_OF:9, AMOUNT:10, FROM_ACCOUNT:11, TO_ACCOUNT:12, SOURCE:13,
  DESCRIPTION:14, NOTES:15, PROOF:16, LINKED_REF:17, DELETED:18
};
var FT_TOTAL = 19;
var FT_HEADERS = ['Fund_Transfer_ID','Submission_Date','Txn_Date','Month','FY','Entry_Type','Class','Direction',
                  'Category','On_Behalf_Of','Amount','From_Account','To_Account','Source_of_Funds',
                  'Description','Notes','Proof_Link','Linked_Ref','Deleted'];
var FT_WIDTHS = [150,180,120,75,80, 90,90,80, 180,140,130,150,150,150, 240,240,200,150,70];

function _getFundTransfersSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName('Fund_Transfers');
  if (sheet) return sheet;
  sheet = _buildSheet(ss, 'Fund_Transfers', FT_HEADERS, FT_WIDTHS);
  _validateCol(sheet, FT.ENTRY_TYPE+1, ['Owner','Transfer']);
  _validateCol(sheet, FT.CLASS+1, ['Transfer']);
  _validateCol(sheet, FT.DIRECTION+1, ['IN','OUT','']);
  _validateCol(sheet, FT.DELETED+1, ['Yes','No']);
  return sheet;
}

function _nextFundTransferId(){
  var sheet = _getFundTransfersSheet();
  if (!sheet || !_hasData(sheet)) return 'FND-0001';
  var ids = sheet.getRange(2, FT.ID+1, sheet.getLastRow()-1, 1).getValues();
  var max = 0;
  for (var i=0; i<ids.length; i++){
    var v = _s(ids[i][0]);
    if (v.indexOf('FND-') === 0){
      var n = parseInt(v.substring(4), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return 'FND-' + String(max + 1).padStart(4, '0');
}

// ============================================================
// RECORD FUND TRANSFER  (owner moving money between accounts)
//   data = {txDate, fromAccount, toAccount, amount, notes}
//     fromAccount / toAccount = Account_ID (or display name)
//   Classified automatically by the two accounts' sides:
//     Personal → Company  = 'Fund In (Owner)'      (company gets owner money)
//     Company  → Personal = 'Fund Out (Owner)'     (owner takes from company)
//     same side           = 'Internal Transfer'    (no accounting impact)
//   Always written as Class = Transfer → excluded from P&L / expense / income.
//   One row per transfer. Source_of_Funds = from, Party_Name = to.
// ============================================================
function recordFundTransfer(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    data = data || {};
    var amount = _n(data.amount);
    if (amount <= 0) return {success:false, message:'Amount must be greater than zero.'};

    var from = getAccountByIdOrName(data.fromAccount);
    var to   = getAccountByIdOrName(data.toAccount);
    if (!from) return {success:false, message:'From-account not found.'};
    if (!to)   return {success:false, message:'To-account not found.'};
    if (from.id === to.id) return {success:false, message:'From and To accounts must be different.'};

    // Classify by sides
    var category;
    if (from.side === 'Personal' && to.side === 'Company')      category = 'Fund In (Owner)';
    else if (from.side === 'Company' && to.side === 'Personal') category = 'Fund Out (Owner)';
    else                                                        category = 'Internal Transfer';

    var now = new Date();
    var subDate = Utilities.formatDate(now,'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var txDate  = data.txDate ? _asDate(data.txDate) : now;
    var sheet   = _getFundTransfersSheet();
    var id      = _nextFundTransferId();

    // Direction: from the company's perspective. Fund In = IN, Fund Out = OUT, internal = none.
    var dir = (category === 'Fund In (Owner)') ? 'IN' : (category === 'Fund Out (Owner)') ? 'OUT' : '';

    var row = new Array(FT_TOTAL).fill('');
    row[FT.ID]              = id;
    row[FT.SUB_DATE]        = subDate;
    row[FT.TX_DATE]         = data.txDate || subDate;
    row[FT.MONTH]           = _monthOf(txDate);
    row[FT.FY]              = _fyFromDate(txDate);
    row[FT.ENTRY_TYPE]      = 'Owner';
    row[FT.CLASS]           = 'Transfer';
    row[FT.DIRECTION]       = dir;
    row[FT.CATEGORY]        = category;
    row[FT.ON_BEHALF_OF]    = from.owner;          // whose money is moving
    row[FT.AMOUNT]          = amount;
    row[FT.FROM_ACCOUNT]    = from.name;           // source account
    row[FT.TO_ACCOUNT]      = to.name;             // destination account
    row[FT.SOURCE]          = from.name;
    row[FT.DESCRIPTION]     = from.name + ' → ' + to.name;
    row[FT.NOTES]           = _s(data.notes);
    row[FT.DELETED]         = 'No';

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    // Tier 5.2 — mirror the transfer into Personal_Transactions when a
    // Personal account is involved. Company↔Company / float involvement is
    // NOT mirrored (nothing personal to track).
    try {
      _mirrorTransferToPersonal({
        txId: id, from: from, to: to, amount: amount,
        txDate: txDate, subDate: subDate, notes: _s(data.notes)
      });
    } catch(mirrErr){ Logger.log('mirror fund transfer: '+mirrErr.message); }

    return {success:true, txId:id, category:category,
            message:'Recorded: ' + category + ' — ' + from.name + ' → ' + to.name};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Internal: write mirror rows into Personal_Transactions. See Tier 5.2.
//   Company→Personal → 1 row for destination owner, Class='Fund In', Direction='IN'
//   Personal→Company → 1 row for source owner, Class='Fund Out', Direction='OUT'
//   Personal→Personal → 2 rows: at source (OUT + Internal), at destination (IN + Internal)
//   Float involvement / Company↔Company → nothing written.
// linkedRef on each mirror = the company Transactions row's txId.
function _mirrorTransferToPersonal(args){
  if (!args) return;
  var from = args.from, to = args.to;
  if (!from || !to) return;
  if (from.side === 'Float' || to.side === 'Float') return;
  var pSheet = _getPersonalSheet();
  if (pSheet.getLastColumn() < PC_TOTAL){
    Logger.log('_mirrorTransferToPersonal: Personal_Transactions is pre-Tier-5.1 (width ' +
               pSheet.getLastColumn() + '). Skipping mirror. Run upgradePersonalTransactionsSchema.');
    return;
  }
  var toWrite = [];
  if (from.side === 'Company' && to.side === 'Personal'){
    toWrite.push({ owner: to.owner,   klass: PCLASS.FUND_IN,  dir: 'IN',
                   source: from.name, counter: to.name });
  } else if (from.side === 'Personal' && to.side === 'Company'){
    toWrite.push({ owner: from.owner, klass: PCLASS.FUND_OUT, dir: 'OUT',
                   source: from.name, counter: to.name });
  } else if (from.side === 'Personal' && to.side === 'Personal'){
    toWrite.push({ owner: from.owner, klass: PCLASS.INTERNAL, dir: 'OUT',
                   source: from.name, counter: to.name });
    toWrite.push({ owner: to.owner,   klass: PCLASS.INTERNAL, dir: 'IN',
                   source: to.name,   counter: from.name });
  } else {
    return;
  }
  toWrite.forEach(function(spec){
    // _nextPersonalId reads sheet.getLastRow() each call, and appendRow below
    // bumps that count — so consecutive iterations get distinct IDs naturally.
    // No suffix needed (a "-1"/"-2" tail would break the standard PTX-{date}-{seq} format).
    var pId = _nextPersonalId(pSheet, new Date());
    var prow = new Array(PC_TOTAL).fill('');
    prow[PC.ID]              = pId;
    prow[PC.SUB_DATE]        = args.subDate;
    prow[PC.TX_DATE]         = args.txDate ? Utilities.formatDate(_asDate(args.txDate),'GMT+5:30','yyyy-MM-dd') : args.subDate;
    prow[PC.MONTH]           = _monthOf(args.txDate || new Date());
    prow[PC.FY]              = _fyFromDate(args.txDate || new Date());
    prow[PC.OWNER]           = spec.owner || 'Owner';
    prow[PC.CLASS]           = spec.klass;
    prow[PC.DIRECTION]       = spec.dir;
    prow[PC.CATEGORY]        = '';
    prow[PC.PARTY_NAME]      = spec.counter;
    prow[PC.DESCRIPTION]     = spec.source + ' → ' + spec.counter;
    prow[PC.AMOUNT]          = args.amount;
    prow[PC.SOURCE]          = spec.source;
    prow[PC.COUNTER_ACCOUNT] = spec.counter;
    prow[PC.PAY_MODE]        = '';
    prow[PC.NOTES]           = args.notes || '';
    prow[PC.LINKED_REF]      = args.txId;
    prow[PC.DELETED]         = 'No';
    pSheet.appendRow(prow);
    SpreadsheetApp.flush();
  });
}

// Owner funds position — net of all Fund In / Fund Out across FYs (or one FY).
//   Returns { net, totalIn, totalOut, byOwner:[{owner, net}] }
// ============================================================
// PENDING SETTLEMENTS — dashboard quick-action list
//   Returns expense transactions where Settle_Status = 'Pending'.
//   Used by the "Pending Settlements" card on the Dashboard so
//   the owner can settle / edit / delete in one click.
//   Returns { ok, data: { rows: [...], total } }
// ============================================================
function listPendingSettlements(filters){
  try {
    filters = filters || {};
    var fy = _activeFY(filters);
    var limit = filters.limit || 50;
    var all = getAllTransactions({fy: fy});
    var rows = [];
    var total = 0;
    all.forEach(function(r){
      if (r.txnClass !== 'Expense') return;
      if (r.settleStatus !== 'Pending') return;
      rows.push({
        id: r.id,
        employee: r.submittedBy || r.onBehalfOf || '(unknown)',
        description: r.description || r.notes || r.category,
        date: r.txDate,
        amount: r.amount,
        category: r.category,
        party: r.party
      });
      total += r.amount;
    });
    // Sort by transaction date desc (newest first); amount desc as tiebreaker
    rows.sort(function(a,b){
      var da = _asDate(a.date), db = _asDate(b.date);
      var td = db - da;
      if (td) return td;
      return b.amount - a.amount;
    });
    var sliced = rows.slice(0, limit);
    return _envOk({ rows: sliced, total: _round(total), count: rows.length });
  } catch(e){ return _envErr(e); }
}

// QUICK SETTLE  — single-click settlement from the dashboard.
//   Sets Settle_Status = 'Settled', Settle_Date = today, Settle_Amount = Amount,
//   Settle_Mode defaults to 'Cash' (or whatever was specified).
//   Returns {success, message}
function quickSettleTransaction(id, paymentMode){
  try {
    if (!_s(id)) return {success:false, message:'Transaction ID required.'};
    var loc = _findTxnRow(_s(id));
    if (!loc) return {success:false, message:'Transaction not found.'};
    var sheet = loc.sheet, rowNum = loc.rowNum;
    var amount = _n(sheet.getRange(rowNum, TC.AMOUNT+1).getValue());
    var today = new Date();
    sheet.getRange(rowNum, TC.SETTLE_STATUS+1).setValue('Settled');
    sheet.getRange(rowNum, TC.SETTLE_DATE+1).setValue(today);
    sheet.getRange(rowNum, TC.SETTLE_AMT+1).setValue(amount);
    sheet.getRange(rowNum, TC.SETTLE_MODE+1).setValue(_s(paymentMode) || 'Cash');
    return {success:true, message:'Settled ' + id};
  } catch(err){ return {success:false, message:err.message}; }
}

// Bulk settle every pending expense for one employee across all FYs.
// Used by the dashboard's "Settle all" action per employee row.
//   employeeName = string (must match TC.SUBMITTED_BY)
//   paymentMode  = optional, defaults to 'Cash'
// Returns {success, count, totalSettled, message}
function bulkSettleEmployeePending(employeeName, paymentMode){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var name = _s(employeeName);
    if (!name) return {success:false, message:'Employee name required.'};
    var mode = _s(paymentMode) || 'Cash';
    var today = new Date();
    var count = 0, totalSettled = 0;
    var txnSheets = _allTxnSheets();
    txnSheets.forEach(function(sheet){
      if (!sheet || sheet.getLastRow() < 2) return;
      var lastRow = sheet.getLastRow();
      var data = sheet.getRange(2, 1, lastRow-1, TC_TOTAL).getValues();
      // Batch collect the rows to update, then write in one range if
      // possible (Sheets API is slow row-by-row).
      var updates = [];
      for (var i=0; i<data.length; i++){
        if (_s(data[i][TC.DELETED]) === 'Yes') continue;
        if (_s(data[i][TC.CLASS]) !== 'Expense') continue;
        if (_s(data[i][TC.SUBMITTED_BY]) !== name) continue;
        if (_s(data[i][TC.SETTLE_STATUS]) === 'Settled') continue;
        var amount = _n(data[i][TC.AMOUNT]);
        if (amount <= 0) continue;
        updates.push({row: i+2, amount: amount});
        count++;
        totalSettled += amount;
      }
      // Write updates — one row at a time is simpler here because the four
      // target columns (SETTLE_STATUS/DATE/AMT/MODE) aren't contiguous.
      updates.forEach(function(u){
        sheet.getRange(u.row, TC.SETTLE_STATUS+1).setValue('Settled');
        sheet.getRange(u.row, TC.SETTLE_DATE+1).setValue(today);
        sheet.getRange(u.row, TC.SETTLE_AMT+1).setValue(u.amount);
        sheet.getRange(u.row, TC.SETTLE_MODE+1).setValue(mode);
      });
    });
    var msg = count
      ? 'Settled ' + count + ' pending expense(s) for ' + name + ' — total ₹' + _round(totalSettled)
      : 'No pending expenses found for ' + name;
    return {success:true, count:count, totalSettled:_round(totalSettled), message:msg};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// Owner funds position — net of all Fund In / Fund Out across FYs (or one FY).
//   Returns { ok, data: { net, totalIn, totalOut, byOwner:[{owner, net}] } }
function getOwnerFundsPosition(filters){
  filters = filters || {};
  try {
    var sheet = _getFundTransfersSheet();
    if (!_hasData(sheet)) return _envOk({net:0, totalIn:0, totalOut:0, byOwner:[]});

    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, FT_TOTAL).getValues();
    var totalIn = 0, totalOut = 0, owners = {};

    data.forEach(function(r){
      var fy = r[FT.FY];
      if (filters.fy && filters.fy !== 'all' && fy !== filters.fy) return;
      if (_s(r[FT.DELETED]) === 'Yes') return;

      var dir = _s(r[FT.DIRECTION]);
      var amt = parseFloat(r[FT.AMOUNT])||0;
      var onBehalfOf = _s(r[FT.ON_BEHALF_OF]) || 'Owner';

      if (dir === 'IN'){
        totalIn += amt;
        owners[onBehalfOf] = (owners[onBehalfOf]||0) + amt;
      } else if (dir === 'OUT'){
        totalOut += amt;
        owners[onBehalfOf] = (owners[onBehalfOf]||0) - amt;
      }
    });

    var byOwner = Object.keys(owners).map(function(k){ return {owner:k, net:_round(owners[k])}; })
                   .filter(function(o){ return o.net !== 0; })
                   .sort(function(a,b){ return b.net - a.net; });
    return _envOk({
      net: _round(totalIn - totalOut),
      totalIn: _round(totalIn),
      totalOut: _round(totalOut),
      byOwner: byOwner
    });
  } catch(e){ return _envErr(e); }
}

// Migrate existing Transfer rows from Transactions to Fund_Transfers (one-time task).
//   Reads all CLASS='Transfer' rows across all FY sheets, copies to Fund_Transfers, marks originals deleted.
//   Returns {moved: count}
function migrateTransfersToNewSheet(){
  try {
    var ss = _getSS();
    var ftSheet = _getFundTransfersSheet();
    var moved = 0;

    // Iterate all Transactions_<FY> sheets
    ss.getSheets().forEach(function(sheet){
      var name = sheet.getName();
      if (!/^Transactions_/.test(name)) return;

      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      var rows = [];
      for (var i=1; i<data.length; i++){
        var r = data[i];
        if (_s(r[TC.CLASS]) !== 'Transfer') continue;
        rows.push({sheetName: name, rowNum: i+1, data: r});
      }

      // Copy each Transfer row to Fund_Transfers
      rows.forEach(function(item){
        var r = item.data;
        var ftRow = new Array(FT_TOTAL).fill('');
        ftRow[FT.ID]           = _s(r[TC.ID]);
        ftRow[FT.SUB_DATE]     = r[TC.SUB_DATE];
        ftRow[FT.TX_DATE]      = r[TC.TX_DATE];
        ftRow[FT.MONTH]        = r[TC.MONTH];
        ftRow[FT.FY]           = r[TC.FY];
        ftRow[FT.ENTRY_TYPE]   = _s(r[TC.ENTRY_TYPE]);
        ftRow[FT.CLASS]        = 'Transfer';
        ftRow[FT.DIRECTION]    = _s(r[TC.DIRECTION]);
        ftRow[FT.CATEGORY]     = _s(r[TC.CATEGORY]);
        ftRow[FT.ON_BEHALF_OF] = _s(r[TC.ON_BEHALF_OF]);
        ftRow[FT.AMOUNT]       = parseFloat(r[TC.AMOUNT])||0;
        ftRow[FT.FROM_ACCOUNT] = _s(r[TC.SOURCE]);
        ftRow[FT.TO_ACCOUNT]   = _s(r[TC.PARTY_NAME]);
        ftRow[FT.SOURCE]       = _s(r[TC.SOURCE]);
        ftRow[FT.DESCRIPTION]  = _s(r[TC.DESCRIPTION]);
        ftRow[FT.NOTES]        = _s(r[TC.NOTES]);
        ftRow[FT.DELETED]      = 'No';
        ftSheet.appendRow(ftRow);
        moved++;

        // Mark original as deleted
        sheet.getRange(item.rowNum, TC.DELETED+1).setValue('Yes');
      });
    });

    SpreadsheetApp.flush();
    return {moved: moved};
  } catch(e){ return {error: e.message}; }
}

// ============================================================
// FUND TRANSFER ANALYTICS  (Dashboard queries on Fund_Transfers)
// ============================================================

// Fund transfer summary — aggregated by owner and direction.
//   formData = {txDate, category, amount, source, paymentMode, notes}
//   If `source` is a COMPANY account → also write a paired Drawing row in the
//   business books (same FY tab) so company cash stays correct.
// ============================================================
function submitPersonalTransaction(formData){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    var now = new Date();
    var subDate = Utilities.formatDate(now,'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var txDate  = formData.txDate ? _asDate(formData.txDate) : now;
    var pSheet  = _getPersonalSheet();

    // Tier 5.1 guard: refuse to write new-schema (18-col) data into a
    // pre-migration sheet — the columns wouldn't line up under the old headers.
    if (pSheet.getLastColumn() > 0 && pSheet.getLastColumn() < PC_TOTAL){
      return {success:false, message:'Personal_Transactions is on the old 13-col schema. Run upgradePersonalTransactionsSchema from the Apps Script editor first, then retry.'};
    }

    var pId     = _nextPersonalId(pSheet, now);
    var source  = _s(formData.source);
    var fromCompany = isCompanyAccount(source);
    var drawingId = '';

    // 1) Paired drawing in business books (only if paid from a company account)
    if (fromCompany){
      var txnSheet = _getTxnSheet(txDate);
      drawingId = _nextTxnId(txnSheet, now);
      var drow = _blankTxnRow();
      drow[TC.ID]           = drawingId;
      drow[TC.ENTRY_TYPE]   = 'Owner';
      drow[TC.CLASS]        = 'Drawing';            // excluded from P&L, kept in cash
      drow[TC.DIRECTION]    = 'OUT';
      drow[TC.SUB_DATE]     = subDate;
      drow[TC.TX_DATE]      = formData.txDate || subDate;
      drow[TC.MONTH]        = _monthOf(txDate);
      drow[TC.FY]           = _fyFromDate(txDate);
      drow[TC.CATEGORY]     = DRAWING_CATEGORY;
      drow[TC.GROUP]        = '';                   // not a P&L group on purpose
      drow[TC.SUBMITTED_BY] = 'Owner';
      drow[TC.DESCRIPTION]  = 'Personal expense (' + _s(formData.category) + ') from company account';
      drow[TC.NOTES]        = _s(formData.notes);
      drow[TC.AMOUNT]       = _n(formData.amount);
      drow[TC.PAY_MODE]     = _s(formData.paymentMode);
      drow[TC.SOURCE]       = source;
      drow[TC.SETTLE_STATUS]= 'N/A';
      drow[TC.RECOVERY]     = 'N/A';
      drow[TC.LINKED_REF]   = pId;                  // ← links to personal row
      drow[TC.DELETED]      = 'No';
      txnSheet.appendRow(drow);
      SpreadsheetApp.flush();
    }

    // 2) Personal row
    // Owner attribution: look up the source account and use its owner. Falls
    // back to formData.owner (frontend override) or 'Owner' when unknown.
    var srcAcc = (typeof getAccountByIdOrName === 'function') ? getAccountByIdOrName(source) : null;
    var ownerName = _s(formData.owner) || (srcAcc && srcAcc.owner) || 'Owner';

    var prow = []; for (var i=0;i<PC_TOTAL;i++) prow.push('');
    prow[PC.ID]              = pId;
    prow[PC.SUB_DATE]        = subDate;
    prow[PC.TX_DATE]         = formData.txDate || subDate;
    prow[PC.MONTH]           = _monthOf(txDate);
    prow[PC.FY]              = _fyFromDate(txDate);
    prow[PC.OWNER]           = ownerName;
    prow[PC.CLASS]           = PCLASS.EXPENSE;
    prow[PC.DIRECTION]       = 'OUT';
    prow[PC.CATEGORY]        = _s(formData.category);
    prow[PC.PARTY_NAME]      = _s(formData.party);
    prow[PC.DESCRIPTION]     = _s(formData.description);
    prow[PC.AMOUNT]          = _n(formData.amount);
    prow[PC.SOURCE]          = source;
    prow[PC.COUNTER_ACCOUNT] = '';   // n/a for a plain expense
    prow[PC.PAY_MODE]        = _s(formData.paymentMode);
    prow[PC.NOTES]       = _s(formData.notes);
    prow[PC.LINKED_REF]  = drawingId;               // ← links to drawing (if any)
    prow[PC.DELETED]     = 'No';
    pSheet.appendRow(prow);
    SpreadsheetApp.flush();

    return {success:true, pId:pId, drawingId:drawingId, autoDrawing:fromCompany};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ============================================================
// READ — business transactions across FY tabs
//   filters: {fy, month, entryType, class, direction, employee, project,
//             includeDeleted}
//   Omit fy → reads ALL FY tabs. fy:'current' → active FY only.
// ============================================================
function _allTxnSheets(){
  var ss = _getSS(), out = [];
  ss.getSheets().forEach(function(sh){
    if (sh.getName().indexOf(T.TXN_PREFIX) === 0) out.push(sh);
  });
  return out;
}

function getAllTransactions(filters){
  filters = filters || {};
  try {
    var sheets;
    if (filters.fy && filters.fy !== 'all'){
      var want = (filters.fy === 'current') ? _currentFY() : _s(filters.fy);
      var ss = _getSS(), one = ss.getSheetByName(_txnSheetName(want));
      sheets = one ? [one] : [];
    } else {
      sheets = _allTxnSheets();
    }
    var includeDeleted = !!filters.includeDeleted;
    var result = [];

    sheets.forEach(function(sheet){
      if (!sheet || sheet.getLastRow() < 2) return;
      // Only read the columns that actually exist — pre-Tier-4.2 sheets are 36 wide.
      var w = Math.min(sheet.getLastColumn(), TC_TOTAL);
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues();
      for (var i=0;i<rows.length;i++){
        try {
          var r = rows[i];
          var id = _s(r[TC.ID]); if (!id) continue;
          var deleted = _s(r[TC.DELETED]) === 'Yes';
          if (deleted && !includeDeleted) continue;

          var rec = {
            id:           id,
            entryType:    _s(r[TC.ENTRY_TYPE]),
            txnClass:     _s(r[TC.CLASS]),
            direction:    _s(r[TC.DIRECTION]),
            txDate:       r[TC.TX_DATE],
            month:        _s(r[TC.MONTH]),
            fy:           _s(r[TC.FY]),
            category:     _s(r[TC.CATEGORY]),
            group:        _s(r[TC.GROUP]),
            submittedBy:  _s(r[TC.SUBMITTED_BY]),
            onBehalfOf:   _s(r[TC.ON_BEHALF_OF]),
            project:      _s(r[TC.PROJECT]),
            partyId:      _s(r[TC.PARTY_ID]),
            party:        _s(r[TC.PARTY_NAME]),
            description:  _s(r[TC.DESCRIPTION]),
            location:     _s(r[TC.LOCATION]),
            notes:        _s(r[TC.NOTES]),
            amount:       _n(r[TC.AMOUNT]),
            taxable:      _n(r[TC.TAXABLE]),
            gstRate:      _s(r[TC.GST_RATE]),
            cgst:         _n(r[TC.CGST]),
            sgst:         _n(r[TC.SGST]),
            igst:         _n(r[TC.IGST]),
            hsnSac:       _s(r[TC.HSN_SAC]),
            paymentMode:  _s(r[TC.PAY_MODE]),
            source:       _s(r[TC.SOURCE]),
            proofLink:    _s(r[TC.PROOF]),
            settleStatus: _s(r[TC.SETTLE_STATUS]),
            settleAmount: _n(r[TC.SETTLE_AMT]),
            settleMode:   _s(r[TC.SETTLE_MODE]),
            recovery:     _s(r[TC.RECOVERY]),
            recoveryRef:  _s(r[TC.RECOVERY_REF]),
            linkedRef:    _s(r[TC.LINKED_REF]),
            deleted:      deleted,
            // Tier 4.2 — blank on rows written before the column existed;
            // callers should treat missing/blank as: fully-recovered rows = Amount,
            // pending rows = 0. See getEmployeeOutstandingAdvances for the resolver.
            recoveredAmount: (r.length > TC.RECOVERED_AMT) ? _n(r[TC.RECOVERED_AMT]) : 0
          };

          if (filters.month     && rec.month       !== _s(filters.month))     continue;
          if (filters.entryType && rec.entryType    !== _s(filters.entryType)) continue;
          if (filters.class     && rec.txnClass      !== _s(filters.class))    continue;
          if (filters.direction && rec.direction     !== _s(filters.direction))continue;
          if (filters.employee  && rec.submittedBy   !== _s(filters.employee)) continue;
          if (filters.project   && rec.project        !== _s(filters.project)) continue;

          result.push(rec);
        } catch(rowErr){ Logger.log('getAllTransactions row: '+rowErr.message); }
      }
    });
    return result;
  } catch(e){ Logger.log('getAllTransactions fatal: '+e.message); return []; }
}

// READ — personal ledger. Reads only the columns that exist (pre-Tier-5.1
// sheets are 13 wide; new sheets are 18 wide).
function getPersonalTransactions(filters){
  filters = filters || {};
  try {
    var sheet = _getSS().getSheetByName(T.PERSONAL);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var w = Math.min(sheet.getLastColumn(), PC_TOTAL);
    var wide = (w >= PC_TOTAL);
    var includeDeleted = !!filters.includeDeleted;
    return sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues()
      .map(function(r){
        // Pre-migration rows use the OLD column layout — remap on the fly.
        if (!wide){
          // Old order: ID, SUB_DATE, TX_DATE, MONTH, FY, CATEGORY, DESCRIPTION,
          //            AMOUNT, SOURCE, PAY_MODE, NOTES, LINKED_REF, DELETED
          return {
            id:_s(r[0]), txDate:r[2], month:_s(r[3]), fy:_s(r[4]),
            owner:'', txnClass:PCLASS.EXPENSE, direction:'OUT',
            category:_s(r[5]), party:'', description:_s(r[6]),
            amount:_n(r[7]), source:_s(r[8]), counterAccount:'',
            paymentMode:_s(r[9]), notes:_s(r[10]),
            linkedRef:_s(r[11]), deleted:_s(r[12])==='Yes'
          };
        }
        return {
          id:_s(r[PC.ID]), txDate:r[PC.TX_DATE], month:_s(r[PC.MONTH]), fy:_s(r[PC.FY]),
          owner:_s(r[PC.OWNER]), txnClass:_s(r[PC.CLASS]), direction:_s(r[PC.DIRECTION]),
          category:_s(r[PC.CATEGORY]), party:_s(r[PC.PARTY_NAME]), description:_s(r[PC.DESCRIPTION]),
          amount:_n(r[PC.AMOUNT]), source:_s(r[PC.SOURCE]), counterAccount:_s(r[PC.COUNTER_ACCOUNT]),
          paymentMode:_s(r[PC.PAY_MODE]), notes:_s(r[PC.NOTES]),
          linkedRef:_s(r[PC.LINKED_REF]), deleted:_s(r[PC.DELETED])==='Yes'
        };
      })
      .filter(function(o){
        if (!o.id) return false;
        if (o.deleted && !includeDeleted) return false;
        if (filters.month && o.month !== _s(filters.month)) return false;
        if (filters.fy && filters.fy !== 'all' && o.fy !== (filters.fy==='current'?_currentFY():_s(filters.fy))) return false;
        return true;
      });
  } catch(e){ Logger.log('getPersonalTransactions: '+e.message); return []; }
}

// ============================================================
// SOFT DELETE  (handles the linked drawing/personal pair)
// ============================================================
function softDeleteTransaction(txId){
  try {
    var found = _findTxnRow(txId);
    if (!found) return {success:false, message:'Not found'};
    if (_s(found.values[TC.DELETED]) === 'Yes') return {success:true, message:'Already deleted'};
    found.sheet.getRange(found.rowNum, TC.DELETED+1).setValue('Yes');

    // If this row is half of a personal/drawing pair, soft-delete the partner too.
    var linked = _s(found.values[TC.LINKED_REF]);
    if (linked) _softDeletePersonalById(linked);
    return {success:true};
  } catch(e){ return {success:false, message:e.message}; }
}

function _softDeletePersonalById(pId){
  try {
    var sheet = _getSS().getSheetByName(T.PERSONAL);
    if (!sheet || sheet.getLastRow() < 2) return;
    // Schema-aware: Deleted column moved 12 → 17 in Tier 5.1. Detect the
    // actual layout by width so a pre-migration sheet isn't marked in the
    // wrong column.
    var deletedIdx = (sheet.getLastColumn() >= PC_TOTAL) ? PC.DELETED : 12;
    var ids = sheet.getRange(2, PC.ID+1, sheet.getLastRow()-1, 1).getValues();
    for (var i=0;i<ids.length;i++){
      if (_s(ids[i][0]) === _s(pId)){ sheet.getRange(i+2, deletedIdx+1).setValue('Yes'); return; }
    }
  } catch(e){ Logger.log('_softDeletePersonalById: '+e.message); }
}

// Locate a txn by ID across FY tabs → {sheet, rowNum, values}
function _findTxnRow(txId){
  var sheets = _allTxnSheets();
  for (var s=0;s<sheets.length;s++){
    var sheet = sheets[s];
    if (sheet.getLastRow() < 2) continue;
    var data = sheet.getRange(2,1,sheet.getLastRow()-1,TC_TOTAL).getValues();
    for (var i=0;i<data.length;i++){
      if (_s(data[i][TC.ID]) === _s(txId)) return {sheet:sheet, rowNum:i+2, values:data[i]};
    }
  }
  return null;
}

// ============================================================
// PROOF UPLOAD  (Drive) — base64 from the form
// ============================================================
function _saveProof(fileData, projectName, date, txId){
  try {
    var root = _driveFolder('3SP — Expense Proofs');
    var subName = _s(projectName) || Utilities.formatDate(_asDate(date),'GMT+5:30','MMM-yyyy');
    var sub  = _driveFolder(subName, root);
    var ext  = (_s(fileData.name) || 'file').split('.').pop().toLowerCase();
    var dt   = Utilities.formatDate(_asDate(date),'GMT+5:30','yyyy-MM-dd');
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType,
                                 dt + '_' + txId + '.' + ext);
    var file = sub.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e){ Logger.log('_saveProof: '+e.message); return ''; }
}
function _driveFolder(name, parent){
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

// ============================================================
// BULK-PASTE WORKFLOW
// ============================================================
// After pasting historical rows into Transactions_<FY> sheets, run
// processBulkPaste() — it sweeps every transaction tab and:
//   1. Fills Group from Category (using owner + employee category masters)
//   2. Fills Party_ID from Party_Name (smart-matched against M_Party)
//   3. Re-applies dropdown validators on Category/Group/Mode/Source/Location
//   4. Reports unknown categories, unmatched parties, placeholder sources
//   5. Auto-runs buildDirectories so rollups refresh in one go
// Safe to re-run — only fills blanks, never overwrites existing values.
// ============================================================

// ============================================================
// repairTransactionIds — rebuild all IDs in chronological order
//   Use after bulk-pasting historical data when IDs are blank,
//   duplicated, or out of sync. Generates fresh IDs in the
//   canonical format TXN-yyyymmdd-#### where #### resets per day.
//   Also rewrites LINKED_REF entries so paired drawing rows stay
//   matched. Safe to run multiple times — produces deterministic
//   IDs from the data.
// ============================================================
function repairTransactionIds(){
  var ss = _getSS();
  var sheets = _allTxnSheets();
  if (!sheets.length){
    var msg = 'No Transactions_<FY> tabs found.';
    try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
    return {success:false, message:msg};
  }

  var report = [];
  var totalChanged = 0;
  var totalRows = 0;
  // First pass — gather every row from every sheet, with sheet+row reference
  // so we can sort globally by date and assign fresh IDs across all FYs.
  var idMap = {};  // oldId -> newId  (for LINKED_REF rewriting)

  sheets.forEach(function(sheet){
    if (sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2,1,sheet.getLastRow()-1,TC_TOTAL).getValues();
    totalRows += data.length;

    // Pair each row with its original index so we can keep original
    // tiebreaker order within a single date.
    var indexed = data.map(function(row, idx){ return {row:row, idx:idx}; });

    // Sort by date asc, then by original row order
    indexed.sort(function(a,b){
      var da = _asDate(a.row[TC.TX_DATE]);
      var db = _asDate(b.row[TC.TX_DATE]);
      var ta = da ? da.getTime() : 0;
      var tb = db ? db.getTime() : 0;
      if (ta !== tb) return ta - tb;
      return a.idx - b.idx;
    });

    // Walk sorted rows, generating sequential IDs per date
    var dateCounter = {};
    var changedHere = 0;
    indexed.forEach(function(entry){
      var row = entry.row;
      var d = _asDate(row[TC.TX_DATE]);
      var stamp = d ? Utilities.formatDate(d, 'GMT+5:30', 'yyyyMMdd') : 'unknown';
      dateCounter[stamp] = (dateCounter[stamp] || 0) + 1;
      var newId = 'TXN-' + stamp + '-' + ('000' + dateCounter[stamp]).slice(-4);
      var oldId = _s(row[TC.ID]);
      if (oldId && oldId !== newId) idMap[oldId] = newId;
      if (oldId !== newId){
        row[TC.ID] = newId;
        changedHere++;
      }
    });

    // Write back in ORIGINAL row order (so visible row positions don't shuffle)
    // — restore indexed back to original order before writing
    var byOriginalIdx = new Array(data.length);
    indexed.forEach(function(e){ byOriginalIdx[e.idx] = e.row; });
    sheet.getRange(2,1,data.length,TC_TOTAL).setValues(byOriginalIdx);
    totalChanged += changedHere;
    report.push('• ' + sheet.getName() + ': ' + data.length + ' rows, ' + changedHere + ' IDs renumbered');
  });

  // Second pass — rewrite LINKED_REF entries that point to renumbered IDs
  var linkedFixed = 0;
  sheets.forEach(function(sheet){
    if (sheet.getLastRow() < 2) return;
    var col = sheet.getRange(2, TC.LINKED_REF+1, sheet.getLastRow()-1, 1).getValues();
    var changed = false;
    for (var i=0; i<col.length; i++){
      var v = _s(col[i][0]);
      if (v && idMap[v]){
        col[i][0] = idMap[v];
        changed = true;
        linkedFixed++;
      }
    }
    if (changed) sheet.getRange(2, TC.LINKED_REF+1, col.length, 1).setValues(col);
  });

  // Also fix LINKED_REF in Personal_Transactions
  var pSheet = ss.getSheetByName(T.PERSONAL);
  if (pSheet && pSheet.getLastRow() >= 2){
    var pCol = pSheet.getRange(2, PC.LINKED_REF+1, pSheet.getLastRow()-1, 1).getValues();
    var pChanged = false;
    for (var j=0; j<pCol.length; j++){
      var pv = _s(pCol[j][0]);
      if (pv && idMap[pv]){ pCol[j][0] = idMap[pv]; pChanged = true; linkedFixed++; }
    }
    if (pChanged) pSheet.getRange(2, PC.LINKED_REF+1, pCol.length, 1).setValues(pCol);
  }

  var summary = 'Transaction ID repair complete:\n' +
                '\nTotal rows scanned: ' + totalRows +
                '\nIDs renumbered: ' + totalChanged +
                '\nLinked references fixed: ' + linkedFixed +
                '\n\nBy sheet:\n' + report.join('\n');
  try { SpreadsheetApp.getUi().alert(summary); } catch(e){ Logger.log(summary); }
  return {success:true, totalRows:totalRows, idsChanged:totalChanged, linkedFixed:linkedFixed, report:report};
}

function processBulkPaste(){
  var ss = _getSS();
  var sheets = _allTxnSheets();
  if (!sheets.length){
    var m0 = 'No Transactions_<FY> tabs found. Run setupTransactionInfra first.';
    try { SpreadsheetApp.getUi().alert(m0); } catch(e){ Logger.log(m0); }
    return {success:false, message:m0};
  }

  var catGroupMap = {};
  getCategories().forEach(function(c){ if(c.category) catGroupMap[c.category] = c.group; });
  getEmployeeCategories().forEach(function(c){ if(c.category && !catGroupMap[c.category]) catGroupMap[c.category] = c.group; });

  var partyRows = [];
  var party = ss.getSheetByName(M.PARTY);
  if (party && _hasData(party)){
    party.getRange(2,1,party.getLastRow()-1,3).getValues().forEach(function(r){
      var id = _s(r[0]), type = _s(r[1]), nm = _s(r[2]);
      if (id && nm) partyRows.push({id:id, type:type, name:nm, norm:_normName(nm)});
    });
  }
  function smartLookup(name, isIncome){
    if (!name) return null;
    var target = _normName(name);
    if (!target) return null;
    var prefer = isIncome ? 'Customer' : 'Vendor';
    for (var i=0;i<partyRows.length;i++){
      var p = partyRows[i];
      if (p.type === prefer && p.norm === target) return p;
    }
    for (var j=0;j<partyRows.length;j++){
      if (partyRows[j].norm === target) return partyRows[j];
    }
    for (var k=0;k<partyRows.length;k++){
      var q = partyRows[k];
      if (q.type === prefer && _normMatch(q.name, name)) return q;
    }
    for (var l=0;l<partyRows.length;l++){
      if (_normMatch(partyRows[l].name, name)) return partyRows[l];
    }
    return null;
  }

  var totalGroups=0, totalIds=0, totalDropdowns=0;
  var unknownCats={}, unmatchedNames={}, placeholderSrc=0;
  var perSheet = [];

  sheets.forEach(function(sh){
    var last = sh.getLastRow();
    if (last < 2) return;
    var name = sh.getName();
    try { _applyTxnDropdowns(sh); totalDropdowns++; } catch(e){}

    var n = last - 1;
    var cats = sh.getRange(2, TC.CATEGORY+1,   n, 1).getValues();
    var grps = sh.getRange(2, TC.GROUP+1,      n, 1).getValues();
    var ids  = sh.getRange(2, TC.PARTY_ID+1,   n, 1).getValues();
    var nms  = sh.getRange(2, TC.PARTY_NAME+1, n, 1).getValues();
    var srcs = sh.getRange(2, TC.SOURCE+1,     n, 1).getValues();
    var cls  = sh.getRange(2, TC.CLASS+1,      n, 1).getValues();

    var gFill=0, iFill=0;
    for (var i=0;i<n;i++){
      var cat = _s(cats[i][0]);
      if (cat){
        if (catGroupMap.hasOwnProperty(cat)){
          if (_s(grps[i][0]) !== catGroupMap[cat]){ grps[i][0] = catGroupMap[cat]; gFill++; }
        } else {
          unknownCats[cat] = (unknownCats[cat]||0)+1;
        }
      }
      if (!_s(ids[i][0]) && _s(nms[i][0])){
        var isIncome = (_s(cls[i][0]) === 'Income');
        var hit = smartLookup(_s(nms[i][0]), isIncome);
        if (hit){ ids[i][0] = hit.id; iFill++; }
        else { unmatchedNames[_s(nms[i][0])] = (unmatchedNames[_s(nms[i][0])]||0)+1; }
      }
      if (/<<.*BANK_ACCOUNT.*>>/i.test(_s(srcs[i][0]))) placeholderSrc++;
    }

    if (gFill) sh.getRange(2, TC.GROUP+1,    n, 1).setValues(grps);
    if (iFill) sh.getRange(2, TC.PARTY_ID+1, n, 1).setValues(ids);

    if (gFill || iFill){
      perSheet.push(name + ': '+gFill+' groups, '+iFill+' party IDs');
      totalGroups += gFill; totalIds += iFill;
    }
  });

  var dirNote = '';
  try { buildDirectories(); dirNote = '\n✅ Directories refreshed.'; }
  catch(e){ dirNote = '\n⚠ buildDirectories failed: '+e.message; }

  var lines = ['Bulk-paste processing complete.\n'];
  lines.push('✅ '+totalGroups+' Group cells filled');
  lines.push('✅ '+totalIds+' Party_ID cells filled');
  lines.push('✅ Dropdowns applied on '+totalDropdowns+' tab'+(totalDropdowns===1?'':'s'));
  if (perSheet.length){ lines.push(''); lines.push('Per tab:'); perSheet.forEach(function(s){ lines.push('  • '+s); }); }
  var ucList = Object.keys(unknownCats);
  if (ucList.length){
    lines.push(''); lines.push('⚠ Categories not found in masters ('+ucList.length+'):');
    ucList.slice(0,10).forEach(function(c){ lines.push('  • '+c+' ('+unknownCats[c]+' row'+(unknownCats[c]>1?'s':'')+')'); });
    if (ucList.length>10) lines.push('  • …and '+(ucList.length-10)+' more');
  }
  var unList = Object.keys(unmatchedNames);
  if (unList.length){
    lines.push(''); lines.push('⚠ Parties not matched in M_Party ('+unList.length+'):');
    unList.slice(0,15).forEach(function(p){ lines.push('  • '+p+' ('+unmatchedNames[p]+' row'+(unmatchedNames[p]>1?'s':'')+')'); });
    if (unList.length>15) lines.push('  • …and '+(unList.length-15)+' more');
    lines.push('  → Add to M_Party, then re-run processBulkPaste.');
  }
  if (placeholderSrc){
    lines.push('');
    lines.push('⚠ '+placeholderSrc+' row(s) still have placeholder source.');
    lines.push('  → Ctrl+H, find <<COMPANY_BANK_ACCOUNT>>, replace with your real source.');
  }
  lines.push(dirNote);

  var msg = lines.join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {
    success: true,
    groupsFilled: totalGroups,
    idsFilled: totalIds,
    unknownCategories: ucList,
    unmatchedParties: unList,
    placeholderSourceRows: placeholderSrc
  };
}

// ============================================================
// TRANSACTIONS TAB — paged/filtered query + edit + bulk-edit
// Backs the Transactions tab in the unified Home UI.
// ============================================================

// getTransactionsPaged({fy, month, category, party, txnClass, status, search, sortBy, sortDir, page, pageSize})
//   Returns { success, rows, total, page, pageSize, totals:{amount, count} }
function getTransactionsPaged(filters){
  // VERSION-TAG: TXN-PAGED-V5  (if you see this string echoed in the response,
  // the latest code is deployed; if not, paste this file again into the editor)
  // Defensive: always return SOMETHING. Multiple try-layers ensure even
  // an unexpected failure inside Apps Script's serializer doesn't cause
  // a silent null return.
  try {
    var result = _getTransactionsPagedInner(filters || {});
    // Belt-and-braces: stringify+parse to strip out any weird types
    // (undefined, functions, NaN, Infinity) that could confuse the
    // google.script.run serializer and cause it to deliver null.
    var clean = JSON.parse(JSON.stringify(result, function(key, val){
      if (val === undefined) return null;
      if (typeof val === 'number' && !isFinite(val)) return null;
      return val;
    }));
    clean._version = 'TXN-PAGED-V5';
    return clean;
  } catch(outerErr){
    return { success: false, _version:'TXN-PAGED-V5', message: 'getTransactionsPaged crashed: ' + (outerErr && outerErr.message), stack: String(outerErr && outerErr.stack || outerErr) };
  }
}

// Tiny diagnostic that returns immediately with no data work.
// If THIS returns null to the client, the bug is environmental (not in our code).
// If this returns the object but getTransactionsPaged returns null, the bug
// is in the data path of getTransactionsPaged.
function pingTransactions(){
  return {
    success: true,
    _version: 'PING-V1',
    message: 'pong',
    timestamp: new Date().toISOString()
  };
}

function _getTransactionsPagedInner(filters){
  try {
    var all = getAllTransactions({fy: filters.fy || 'current', includeDeleted: !!filters.includeDeleted});

    // Apply post-fetch filters
    var month   = _s(filters.month);
    var cat     = _s(filters.category);
    var party   = _s(filters.party).toLowerCase();
    var partyId = _s(filters.partyId);
    var cls     = _s(filters.txnClass);
    var status  = _s(filters.status);
    var grp     = _s(filters.group);
    var q       = _s(filters.search).toLowerCase();

    var filtered = all.filter(function(r){
      if (month && r.month !== month) return false;
      if (cat && r.category !== cat) return false;
      if (grp && r.group !== grp) return false;
      if (cls && r.txnClass !== cls) return false;
      if (status && r.settleStatus !== status) return false;
      if (partyId && r.partyId !== partyId) return false;
      if (party && r.party.toLowerCase().indexOf(party) === -1) return false;
      if (q){
        var hay = (r.party + ' ' + r.description + ' ' + r.notes + ' ' + r.category + ' ' + r.id).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });

    // Sort
    var sortBy = _s(filters.sortBy) || 'txDate';
    var dir = (_s(filters.sortDir) === 'asc') ? 1 : -1;
    filtered.sort(function(a,b){
      var va = a[sortBy], vb = b[sortBy];
      if (va instanceof Date) va = va.getTime();
      if (vb instanceof Date) vb = vb.getTime();
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va == null) return 1; if (vb == null) return -1;
      if (va < vb) return -1*dir; if (va > vb) return 1*dir;
      return 0;
    });

    // Totals across the filtered set (before paging)
    var totalAmount = 0, totalCount = filtered.length;
    filtered.forEach(function(r){
      if (r.txnClass === 'Expense') totalAmount -= r.amount;
      else if (r.txnClass === 'Income') totalAmount += r.amount;
    });

    // Page
    var page = Math.max(1, parseInt(filters.page) || 1);
    var pageSize = Math.min(500, Math.max(10, parseInt(filters.pageSize) || 50));
    var startIdx = (page - 1) * pageSize;
    var paged = filtered.slice(startIdx, startIdx + pageSize);

    return {
      success: true,
      rows: paged.map(_serializeTxnRow),
      total: totalCount,
      page: page,
      pageSize: pageSize,
      totals: { count: totalCount, netAmount: _round(totalAmount) }
    };
  } catch(err){
    return { success: false, message: err.message };
  }
}

// Apps Script's google.script.run can choke on raw Date objects in
// arrays — especially when some rows have valid Dates and others have
// empty strings in the same field. Convert all Date-like fields to
// ISO date strings before returning to the client.
function _serializeTxnRow(r){
  if (!r) return r;
  return {
    id: r.id, entryType: r.entryType, txnClass: r.txnClass, direction: r.direction,
    txDate: _safeDateStr(r.txDate), subDate: _safeDateStr(r.subDate),
    month: r.month, fy: r.fy, category: r.category, group: r.group,
    submittedBy: r.submittedBy, onBehalfOf: r.onBehalfOf, project: r.project,
    partyId: r.partyId, party: r.party, description: r.description,
    location: r.location, notes: r.notes,
    amount: r.amount, taxable: r.taxable, gstRate: r.gstRate,
    cgst: r.cgst, sgst: r.sgst, igst: r.igst, hsnSac: r.hsnSac,
    paymentMode: r.paymentMode, source: r.source, proofLink: r.proofLink,
    settleStatus: r.settleStatus, settleDate: _safeDateStr(r.settleDate),
    settleAmount: r.settleAmount, settleMode: r.settleMode,
    recovery: r.recovery, recoveryRef: r.recoveryRef, linkedRef: r.linkedRef
  };
}

function _safeDateStr(v){
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date){
    if (isNaN(v.getTime())) return '';
    try { return Utilities.formatDate(v, 'GMT+5:30', 'yyyy-MM-dd'); } catch(e){ return ''; }
  }
  return String(v);
}

// updateTransaction(id, patches)
//   Lets the Transactions tab edit a row in place. Allowed fields:
//   category, group, project, party (resolves to partyId), partyId,
//   description, location, notes, paymentMode, source, settleStatus,
//   settleAmount, settleMode, recovery, recoveryRef
//   Re-derives Group when category changes.
function updateTransaction(id, patches){
  var lock = LockService.getDocumentLock();
  try { lock.waitLock(15000); } catch(e){ return {success:false, message:'Busy — try again.'}; }
  try {
    id = _s(id);
    if (!id) return {success:false, message:'Transaction ID required.'};
    var loc = _findTxnRow(id);
    if (!loc) return {success:false, message:'Transaction not found.'};
    var sheet = loc.sheet, rowNum = loc.rowNum;
    patches = patches || {};

    // Whitelist of editable columns: field -> TC col index
    var allow = {
      category:     TC.CATEGORY,
      group:        TC.GROUP,
      project:      TC.PROJECT,
      partyId:      TC.PARTY_ID,
      party:        TC.PARTY_NAME,
      description:  TC.DESCRIPTION,
      location:     TC.LOCATION,
      notes:        TC.NOTES,
      paymentMode:  TC.PAY_MODE,
      source:       TC.SOURCE,
      settleStatus: TC.SETTLE_STATUS,
      settleAmount: TC.SETTLE_AMT,
      settleMode:   TC.SETTLE_MODE,
      recovery:     TC.RECOVERY,
      recoveryRef:  TC.RECOVERY_REF
    };

    // If a Party is being changed by name (typed in), resolve to ID
    if (patches.party && !patches.partyId){
      var pr = findOrCreateParty('Vendor', patches.party);
      if (pr.id){ patches.partyId = pr.id; patches.party = pr.name; }
    }
    // If category changes, re-derive the Group automatically (unless explicitly patched)
    if (patches.category && !patches.hasOwnProperty('group')){
      var grp = _groupForCategory(patches.category);
      if (grp) patches.group = grp;
    }

    var changedFields = [];
    Object.keys(patches).forEach(function(k){
      if (!allow.hasOwnProperty(k)) return;
      var col = allow[k] + 1;
      var val = patches[k];
      // Coerce numeric fields
      if (k === 'settleAmount') val = _n(val);
      sheet.getRange(rowNum, col).setValue(val);
      changedFields.push(k);
    });

    if (!changedFields.length) return {success:false, message:'Nothing to update.'};
    return {success:true, id:id, changed:changedFields, sheet:sheet.getName(), row:rowNum};
  } catch(err){
    return {success:false, message:err.message};
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// bulkUpdateTransactions(ids[], patches)
//   Same patch shape as updateTransaction, applied to each ID in turn.
//   Useful for batch-recategorising or marking settled in one go.
function bulkUpdateTransactions(ids, patches){
  ids = Array.isArray(ids) ? ids : [];
  patches = patches || {};
  if (!ids.length) return {success:false, message:'No IDs provided.'};
  if (!Object.keys(patches).length) return {success:false, message:'No patches provided.'};

  var ok = 0, fail = 0, errors = [];
  ids.forEach(function(id){
    var r = updateTransaction(id, patches);
    if (r && r.success) ok++;
    else { fail++; errors.push(id + ': ' + (r && r.message ? r.message : 'unknown error')); }
  });
  return {
    success: ok > 0,
    updated: ok,
    failed: fail,
    errors: errors.slice(0, 25),
    message: 'Bulk update: ' + ok + ' succeeded, ' + fail + ' failed.'
  };
}
// ============================================================
// _sendExpenseConfirmationEmail — Tier 8.3
// Brief HTML confirmation to the submitting employee after
// submitEmployeeExpense succeeds. Not itemised — intentionally short.
// ============================================================
function _sendExpenseConfirmationEmail(email, name, expenses, statusLine){
  try {
    if (!email) return;
    var C = { dark:'#0F172A', primary:'#1F4E79', accent:'#2C7DA0', muted:'#5C6675', border:'#D1D1D1' };
    var expList = expenses || [];
    var total = 0;
    // Escape HTML: replace &<>"
    var esc = function(s){ var t=String(s==null?'':s); return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); };
    var rows = expList.map(function(e){
      total += _n(e.amount);
      return '<tr><td style="padding:8px 10px;border-bottom:1px solid #eef0f3;font-size:12px">'+
        Utilities.formatDate(_asDate(e.date),'GMT+5:30','dd MMM') +'</td>'+
        '<td style="padding:8px 10px;border-bottom:1px solid #eef0f3;font-size:12px">'+esc(_s(e.category))+'</td>'+
        '<td style="padding:8px 10px;border-bottom:1px solid #eef0f3;font-size:12px">'+esc(_s(e.description))+'</td>'+
        '<td style="padding:8px 10px;border-bottom:1px solid #eef0f3;font-size:12px;text-align:right">₹ '+_fmtN(_n(e.amount))+'</td></tr>';
    }).join('');
    var totalStr = _fmtN(total);
    var html = ''+
      '<div style="max-width:580px;margin:20px auto;font-family:Segoe UI,Arial,sans-serif;border-radius:12px;overflow:hidden;border:1px solid '+C.border+'">'+
        '<div style="background:'+C.dark+';padding:18px 22px">'+
          '<div style="color:#fff;font-size:16px;font-weight:700">Expenses recorded</div>'+
          '<div style="color:#7AB8D5;font-size:12px;margin-top:4px">Three Star Projects · Finance System — auto-confirmation</div>'+
        '</div>'+
        '<div style="background:#fff;padding:20px 22px">'+
          '<p style="margin:0 0 14px;font-size:14px">Hi <strong>'+ (name || 'there') +'</strong>,</p>'+
          '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:14px">'+
            '<thead><tr style="background:#f4f5f7"><th style="padding:8px 10px;text-align:left;font-weight:600;color:#475569">Date</th>'+
            '<th style="padding:8px 10px;text-align:left;font-weight:600;color:#475569">Category</th>'+
            '<th style="padding:8px 10px;text-align:left;font-weight:600;color:#475569">Description</th>'+
            '<th style="padding:8px 10px;text-align:right;font-weight:600;color:#475569">Amount</th></tr></thead>'+
            '<tbody>'+ rows +'</tbody>'+
          '</table>'+
          '<div style="padding:10px;background:#f4f5f7;border-radius:6px;text-align:right;font-weight:600;font-size:13px">'+
            'Total: <strong>₹'+ totalStr +'</strong>'+
          '</div>'+
          '<p style="margin:14px 0 0;font-size:13px;color:'+C.muted+'">'+ (statusLine || '') +'</p>'+
        '</div>'+
        '<div style="background:'+C.dark+';padding:10px;text-align:center">'+
          '<span style="color:#64748B;font-size:11px">Automated — do not reply.</span>'+
        '</div>'+
      '</div>';
    MailApp.sendEmail({
      to: email,
      subject: '[3SP] '+ expList.length +' expense(s) recorded — ₹' + totalStr,
      htmlBody: html
    });
  } catch(e){ Logger.log('_sendExpenseConfirmationEmail: '+e.message); }
}

// ============================================================
// SUBMISSION LOGGING  (Tier 8.4 — audit trail for submissions)
// ============================================================
function _getSubmissionLogSheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName('Submission_Log');
  if (sheet) return sheet;
  sheet = _buildSheet(ss, 'Submission_Log',
    ['Timestamp','Employee','Expense_Count','Total_Amount','Outcome','Error_Message'],
    [180, 160, 130, 130, 120, 300]);
  return sheet;
}

function _logSubmission(employeeName, expenses, result){
  try {
    var sheet = _getSubmissionLogSheet();
    var now = Utilities.formatDate(new Date(),'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var expList = expenses || [];
    var count = expList.length;
    var total = 0;
    expList.forEach(function(e){ total += _n(e.amount); });
    var row = [
      now,
      _s(employeeName),
      count,
      _round(total),
      (result && result.success) ? 'Success' : 'Failed',
      (result && result.message) ? _s(result.message) : ''
    ];
    sheet.appendRow(row);
    SpreadsheetApp.flush();
  } catch(e){ Logger.log('_logSubmission: '+e.message); }
}

// ============================================================
// upgradeTransactionSheets — Tier 4.2 migration
// Widens every Transactions_<FY> sheet by 1 column (Recovered_Amount).
// Back-fills:
//   • Recovered_Amount = Amount on rows where Recovery_Status = 'Recovered'
//   • Recovered_Amount = 0 everywhere else
// Idempotent — skips a sheet that's already wide enough.
// User must run this once from the editor after deploying Tier 4.2.
// ============================================================
function upgradeTransactionSheets(){
  var sheets = _allTxnSheets();
  var report = { widened:[], backfilled:[], skipped:[] };
  sheets.forEach(function(sheet){
    var name = sheet.getName();
    if (sheet.getLastColumn() >= TC_TOTAL){
      report.skipped.push(name + ' (already width ' + sheet.getLastColumn() + ')');
      return;
    }
    // Widen: append the header for the new column
    sheet.getRange(1, TC.RECOVERED_AMT+1)
      .setValue('Recovered_Amount').setFontWeight('bold');
    sheet.setColumnWidth(TC.RECOVERED_AMT+1, 130);
    report.widened.push(name);

    if (sheet.getLastRow() < 2) return;
    var lr = sheet.getLastRow();
    // Read Recovery_Status + Amount; write Recovered_Amount in one batch
    var recoveryCol = sheet.getRange(2, TC.RECOVERY+1, lr-1, 1).getValues();
    var amountCol   = sheet.getRange(2, TC.AMOUNT+1,   lr-1, 1).getValues();
    var out = new Array(lr-1);
    var backfilledN = 0;
    for (var i=0;i<recoveryCol.length;i++){
      if (_s(recoveryCol[i][0]) === 'Recovered'){
        out[i] = [_n(amountCol[i][0])];
        backfilledN++;
      } else {
        out[i] = [0];
      }
    }
    sheet.getRange(2, TC.RECOVERED_AMT+1, out.length, 1).setValues(out);
    if (backfilledN > 0) report.backfilled.push(name + ' (' + backfilledN + ' recovered rows)');
  });

  var msg = 'Transaction schema upgrade (Recovered_Amount):\n' +
            '• Widened: ' + (report.widened.length ? report.widened.join(', ') : 'none') + '\n' +
            '• Back-filled: ' + (report.backfilled.length ? report.backfilled.join(', ') : 'none') + '\n' +
            '• Skipped (already migrated): ' + (report.skipped.length ? report.skipped.join(', ') : 'none') +
            '\n\nSalary calculator will now track partial advance recovery correctly.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return { success:true, report:report };
}

// ============================================================
// upgradePersonalTransactionsSchema — Tier 5.1 migration
// Personal_Transactions grew from 13 → 18 columns. Column ORDER changed
// (owner/class/direction/party/counter-account inserted in the middle),
// so existing rows must be REWRITTEN to the new layout, not just widened.
//
// Handles three input shapes:
//   • 13 cols — pre-migration.
//   • 15 cols — an in-between layout used briefly by an earlier draft.
//   • 18 cols — already migrated (idempotent no-op).
//
// Owner is derived from the row's source-of-funds via M_Account (owner
// column). Class defaults to 'Expense'. Direction defaults to 'OUT'.
// Party_Name defaults to blank.
// ============================================================
function upgradePersonalTransactionsSchema(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(T.PERSONAL);
  if (!sheet){
    // Fresh install — just create with the new schema.
    _getPersonalSheet();
    var m0 = 'Personal_Transactions did not exist — created with the new 18-col schema.';
    try { SpreadsheetApp.getUi().alert(m0); } catch(e){ Logger.log(m0); }
    return { success:true, message:m0 };
  }

  var startCols = sheet.getLastColumn();
  var startRows = sheet.getLastRow();

  if (startCols >= PC_TOTAL){
    var m1 = 'Personal_Transactions already has ' + startCols + ' columns — no migration needed.';
    try { SpreadsheetApp.getUi().alert(m1); } catch(e){ Logger.log(m1); }
    return { success:true, message:m1 };
  }

  // Snapshot existing rows before we rewrite the sheet.
  var existing = startRows > 1
    ? sheet.getRange(2, 1, startRows - 1, startCols).getValues()
    : [];

  // Owner lookup by source-of-funds (name OR id).
  var accountsByKey = {};
  try {
    (getAccounts() || []).forEach(function(a){
      if (a.id)   accountsByKey[a.id]   = a;
      if (a.name) accountsByKey[a.name] = a;
    });
  } catch(e){ Logger.log('upgradePersonalTransactionsSchema accounts: '+e.message); }

  // Reshape: build new rows in the new column order.
  var newRows = existing.map(function(r){
    // Old-layout index mapping:
    //   13-col: [ID, SUB, TX, MONTH, FY, CATEGORY, DESCRIPTION, AMOUNT, SOURCE, PAY_MODE, NOTES, LINKED_REF, DELETED]
    //   15-col: same as 13-col + 2 unknown trailing cols we ignore.
    var id       = _s(r[0]);
    var subDate  = r[1];
    var txDate   = r[2];
    var monthLbl = _s(r[3]);
    var fyLbl    = _s(r[4]);
    var cat      = _s(r[5]);
    var desc     = _s(r[6]);
    var amt      = _n(r[7]);
    var source   = _s(r[8]);
    var payMode  = _s(r[9]);
    var notes    = _s(r[10]);
    var linkRef  = _s(r[11]);
    var deleted  = _s(r[12]);

    var acc = accountsByKey[source];
    var owner = acc ? acc.owner : (source ? source : 'Owner');

    var out = new Array(PC_TOTAL).fill('');
    out[PC.ID]              = id;
    out[PC.SUB_DATE]        = subDate;
    out[PC.TX_DATE]         = txDate;
    out[PC.MONTH]           = monthLbl;
    out[PC.FY]              = fyLbl;
    out[PC.OWNER]           = owner;
    out[PC.CLASS]           = PCLASS.EXPENSE;
    out[PC.DIRECTION]       = 'OUT';
    out[PC.CATEGORY]        = cat;
    out[PC.PARTY_NAME]      = '';
    out[PC.DESCRIPTION]     = desc;
    out[PC.AMOUNT]          = amt;
    out[PC.SOURCE]          = source;
    out[PC.COUNTER_ACCOUNT] = '';
    out[PC.PAY_MODE]        = payMode;
    out[PC.NOTES]           = notes;
    out[PC.LINKED_REF]      = linkRef;
    out[PC.DELETED]         = deleted || 'No';
    return out;
  });

  // Rewrite header + all rows in place.
  sheet.clearContents();
  sheet.getRange(1, 1, 1, PERSONAL_HEADERS.length).setValues([PERSONAL_HEADERS]);
  PERSONAL_WIDTHS.forEach(function(w, i){ sheet.setColumnWidth(i+1, w); });
  if (newRows.length){
    sheet.getRange(2, 1, newRows.length, PC_TOTAL).setValues(newRows);
  }
  sheet.setFrozenRows(1);
  _validateCol(sheet, PC.DELETED+1, ['Yes','No']);
  try { if (typeof _beautifySheet === 'function') _beautifySheet(sheet); } catch(e){}

  var msg = 'Personal_Transactions migrated: ' + startCols + ' → ' + PC_TOTAL + ' cols. Rewrote ' + newRows.length + ' row(s).';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return { success:true, message:msg, rewrote:newRows.length };
}

// ============================================================
// submitOwnerBulk — Tier 6 batch entry for owner transactions
// Input: { shared: {direction, txDate, onBehalfOf, paymentMode, source, location},
//          rows: [{category, party, amount, description, project, notes}, …] }
// Delegates each non-empty row to submitOwnerTransaction, so all downstream
// side-effects (party creation, category resolution, advance flagging) stay
// consistent. Empty rows (no amount and no category) are silently skipped.
// GAS LockService is re-entrant per execution, so we don't wrap here.
// Returns { savedCount, failedCount, saved:[{index,txId}], failed:[{index,message}] }.
// ============================================================
function submitOwnerBulk(payload){
  payload = payload || {};
  var shared = payload.shared || {};
  var rows   = payload.rows || [];
  if (!rows.length) return {success:false, message:'No rows to save.'};

  var saved = [], failed = [];
  rows.forEach(function(r, i){
    if (!r) return;
    var amt = _n(r.amount);
    var cat = _s(r.category);
    if (amt <= 0 && !cat) return;   // silently skip fully-empty rows
    if (amt <= 0){ failed.push({index:i, message:'amount required'}); return; }
    if (!cat)  { failed.push({index:i, message:'category required'}); return; }

    var formData = {
      direction:   shared.direction || 'OUT',
      txDate:      shared.txDate,
      onBehalfOf:  shared.onBehalfOf || 'Company',
      paymentMode: shared.paymentMode,
      source:      shared.source,
      location:    shared.location,
      category:    cat,
      party:       _s(r.party),
      amount:      amt,
      description: _s(r.description),
      project:     _s(r.project),
      notes:       _s(r.notes)
    };
    try {
      var res = submitOwnerTransaction(formData);
      if (res && res.success) saved.push({index:i, txId: res.txId});
      else failed.push({index:i, message:(res && res.message) || 'unknown'});
    } catch(e){
      failed.push({index:i, message:e.message});
    }
  });

  return {
    success: saved.length > 0,
    savedCount:  saved.length,
    failedCount: failed.length,
    saved: saved,
    failed: failed,
    message: 'Saved ' + saved.length + ' · Failed ' + failed.length
  };
}
// ============================================================
// Projects — per-project rollup + detailed overview (Batch A)
// Consumed by the Projects tab KPIs. Dashboard tiles (Batch B)
// also read getProjectRollup for the project leaderboard.
// ============================================================

// Resolve an incoming fy argument to a concrete FY string, or '' for all-FYs.
// Called by every function that filters by FY. `emptyMeansCurrent` controls
// what '' / undefined does — Dashboard tiles default to the CURRENT FY,
// the Projects tab's "All FY" picker expects to see everything.
//   'current'       → current FY
//   'all'           → ''
//   '' / undefined  → current FY (if emptyMeansCurrent) else ''
//   anything else   → as-is
function _resolveFy(fy, emptyMeansCurrent){
  var s = _s(fy);
  if (s === 'current') return _currentFY();
  if (s === 'all')     return '';
  if (s === '')        return emptyMeansCurrent ? _currentFY() : '';
  return s;
}

// Per-project rollup. `fy` may be a specific FY, 'current', 'all', or ''.
// '' / 'all' aggregates across every FY. Projects with zero activity on both
// sides are dropped. Spent = Direct expense txns + Project_Costs (Inc-GST)
// + ticked Vendor Bills. Budget = Invoiced (locked rule).
function getProjectRollup(fy){
  try {
    fy = _resolveFy(fy, false);   // '' means all FYs
    var byProj = {};
    function _b(p){
      if (!byProj[p]) byProj[p] = {project:p, invoiced:0, received:0,
        spentDirect:0, spentProjectCosts:0, spentBills:0};
      return byProj[p];
    }

    // 1) Invoiced + Received per project (read Invoices sheet directly —
    //    getInvoices() doesn't surface the Project column).
    var invSheet = _getSS().getSheetByName(AR.INVOICES);
    if (invSheet && invSheet.getLastRow() > 1){
      var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
      var iData = invSheet.getRange(2, 1, invSheet.getLastRow()-1, iW).getValues();
      iData.forEach(function(r){
        if (iW <= IC.PROJECT) return;
        var proj = _s(r[IC.PROJECT]);
        if (!proj) return;
        if (_s(r[IC.STATUS]) === INV_STATUS.CANCELLED) return;
        if (fy && _s(r[IC.FY]) !== fy) return;
        var amt = _n(r[IC.AMOUNT]) || 0;
        var out = _n(r[IC.OUTSTANDING]) || 0;
        _b(proj).invoiced += amt;
        _b(proj).received += (amt - out);
      });
    }

    // 2) Direct Expense transactions tagged to project — skip bill-linked rows
    //    so a BILL-… payment doesn't double-count with its own bill.
    var txns = getAllTransactions(fy ? {fy:fy} : {}) || [];
    txns.forEach(function(r){
      if (!r.project) return;
      if (r.txnClass !== 'Expense') return;
      if (r.linkedRef && _s(r.linkedRef).indexOf('BILL-') === 0) return;
      _b(r.project).spentDirect += _n(r.amount) || 0;
    });

    // 3) Project_Costs rows (Amount col = final Inc-GST total).
    var pcSheet = _getProjectCostsSheet();
    if (pcSheet && pcSheet.getLastRow() > 1){
      var pcW = Math.min(pcSheet.getLastColumn(), PCC_TOTAL);
      var pcData = pcSheet.getRange(2, 1, pcSheet.getLastRow()-1, pcW).getValues();
      pcData.forEach(function(r){
        if (_s(r[PCC.DELETED]) === 'Yes') return;
        var proj = _s(r[PCC.PROJECT]);
        if (!proj) return;
        if (fy && _fyFromDate(r[PCC.ENTRY_DATE]) !== fy) return;
        _b(proj).spentProjectCosts += _n(r[PCC.AMOUNT]) || 0;
      });
    }

    // 4) Vendor Bills ticked as "count in project cost" — gross (Inc-GST).
    // Allocation-aware (Option B): if the bill has non-deleted allocations,
    // each allocation's amount is credited to its own project. Bills with no
    // allocations fall back to Bill.Project (legacy single-project behavior).
    var allocMap = (typeof _getAllBillAllocationsMap === 'function') ? _getAllBillAllocationsMap() : {};
    var billsSheet = _getSS().getSheetByName(AP.BILLS);
    if (billsSheet && billsSheet.getLastRow() > 1){
      var bW = billsSheet.getLastColumn();
      var bData = billsSheet.getRange(2, 1, billsSheet.getLastRow()-1, bW).getValues();
      bData.forEach(function(r){
        if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
        if (fy && _s(r[BC.FY]) !== fy) return;
        if (!_billCountsForProject(r)) return;
        var bid    = _s(r[BC.ID]);
        var allocs = allocMap[bid];
        if (allocs && allocs.length){
          allocs.forEach(function(a){ _b(a.project).spentBills += a.amount || 0; });
        } else {
          var proj = _s(r[BC.PROJECT]);
          if (!proj) return;
          _b(proj).spentBills += _n(r[BC.AMOUNT]) || 0;
        }
      });
    }

    var rows = [];
    Object.keys(byProj).forEach(function(p){
      var b = byProj[p];
      var spent = b.spentDirect + b.spentProjectCosts + b.spentBills;
      if ((b.invoiced || 0) === 0 && spent === 0) return;
      var budget = b.invoiced;
      var margin = b.invoiced - spent;
      var marginPct = b.invoiced > 0 ? (margin / b.invoiced * 100) : null;
      rows.push({
        project: p,
        invoiced: _round(b.invoiced),
        received: _round(b.received),
        spent: _round(spent),
        spentDirect: _round(b.spentDirect),
        spentProjectCosts: _round(b.spentProjectCosts),
        spentBills: _round(b.spentBills),
        budget: _round(budget),
        variance: _round(budget - spent),
        margin: _round(margin),
        marginPct: marginPct === null ? null : _round(marginPct)
      });
    });
    rows.sort(function(a,b){ return b.invoiced - a.invoiced; });
    return {ok:true, data:{rows:rows, fy:fy || 'all', count:rows.length}};
  } catch(e){ return {ok:false, error:e.message}; }
}

// One call for all KPI tiles on the Projects tab, FY-filterable.
// fy=''|'all' means all FYs, 'current' means the current FY.
function getProjectOverviewV2(project, fy){
  try {
    if (!project) return _envOk({});
    fy = _resolveFy(fy, false);   // '' means all FYs

    var rollupRes = getProjectRollup(fy || 'all');
    var row = null;
    if (rollupRes.ok && rollupRes.data && rollupRes.data.rows){
      for (var i=0; i<rollupRes.data.rows.length; i++){
        if (rollupRes.data.rows[i].project === project){ row = rollupRes.data.rows[i]; break; }
      }
    }
    row = row || {project:project, invoiced:0, received:0, spent:0,
      spentDirect:0, spentProjectCosts:0, spentBills:0,
      budget:0, variance:0, margin:0, marginPct:null};

    var costExGst = _computeProjectCostExGst(project, fy);
    var marginExGst = row.invoiced - costExGst;
    var marginExGstPct = row.invoiced > 0 ? _round(marginExGst / row.invoiced * 100) : null;

    // Cash-basis receipts for this project — only computed when FY is specific
    // (so we can split "received this FY against prior-FY invoices"). All-FY
    // view has no distinction, so defaults to 0 / 0.
    var cash = fy ? _computeProjectReceiptsCashBasis(project, fy) : {all:0, priorFyInvs:0};

    return _envOk({
      project: project, fy: fy || 'all',
      invoiced: row.invoiced,
      received: row.received,
      balanceDue: _round(row.invoiced - row.received),
      spent: row.spent,
      spentDirect: row.spentDirect,
      spentProjectCosts: row.spentProjectCosts,
      spentBills: row.spentBills,
      spentExGst: costExGst,
      budget: row.budget,
      variance: row.variance,
      margin: row.margin,
      marginPct: row.marginPct,
      marginExGst: _round(marginExGst),
      marginExGstPct: marginExGstPct,
      // Cash-basis receipts (fy-only, else 0)
      receivedCashThisFy:      cash.all,
      receivedCashPriorFyInvs: cash.priorFyInvs
    });
  } catch(e){ return _envErr(e); }
}

// Cash-basis receipts for one project in one FY — walks the Payments sheet
// and resolves each applied invoice to its project + FY.
//   {all: Σ received this FY for any of this project's invoices,
//    priorFyInvs: Σ received this FY for this project's PRIOR-FY invoices}
function _computeProjectReceiptsCashBasis(project, fy){
  var invMap = {};
  var invSheet = _getSS().getSheetByName(AR.INVOICES);
  if (invSheet && invSheet.getLastRow() > 1){
    var iW = Math.min(invSheet.getLastColumn(), IC_TOTAL);
    invSheet.getRange(2,1,invSheet.getLastRow()-1,iW).getValues().forEach(function(r){
      var no = _s(r[IC.NO]); if (!no) return;
      invMap[no] = {fy: _s(r[IC.FY]), project: (iW > IC.PROJECT ? _s(r[IC.PROJECT]) : '')};
    });
  }
  var pSheet = _getSS().getSheetByName(AR.PAYMENTS);
  if (!pSheet || pSheet.getLastRow() < 2) return {all:0, priorFyInvs:0};
  var pW = Math.min(pSheet.getLastColumn(), PMC_TOTAL);
  var pData = pSheet.getRange(2,1,pSheet.getLastRow()-1,pW).getValues();
  var all = 0, priorFyInvs = 0;
  pData.forEach(function(r){
    if (pW > PMC.CANCELLED_ON && _s(r[PMC.CANCELLED_ON])) return;
    if (_s(r[PMC.FY]) !== fy) return;
    var applied = _parseAppliedString(_s(r[PMC.APPLIED]));
    applied.forEach(function(a){
      var inv = invMap[a.ref];
      if (!inv || inv.project !== project) return;
      all += a.amount;
      if (inv.fy !== fy) priorFyInvs += a.amount;
    });
  });
  return {all: _round(all), priorFyInvs: _round(priorFyInvs)};
}

// Ex-GST cost = Taxable where known, Amount as fallback. Direct txns
// usually have no GST split, so they're used as-is.
function _computeProjectCostExGst(project, fy){
  fy = _resolveFy(fy, false);   // '' means all FYs
  var total = 0;
  var pcSheet = _getProjectCostsSheet();
  if (pcSheet && pcSheet.getLastRow() > 1){
    var pcW = Math.min(pcSheet.getLastColumn(), PCC_TOTAL);
    var pcData = pcSheet.getRange(2, 1, pcSheet.getLastRow()-1, pcW).getValues();
    pcData.forEach(function(r){
      if (_s(r[PCC.DELETED]) === 'Yes') return;
      if (_s(r[PCC.PROJECT]) !== project) return;
      if (fy && _fyFromDate(r[PCC.ENTRY_DATE]) !== fy) return;
      var base = (pcW > PCC.AMOUNT_EX_GST) ? _n(r[PCC.AMOUNT_EX_GST]) : _n(r[PCC.AMOUNT]);
      total += base || 0;
    });
  }
  var txns = getAllTransactions(fy ? {fy:fy} : {}) || [];
  txns.forEach(function(r){
    if (r.project !== project) return;
    if (r.txnClass !== 'Expense') return;
    if (r.linkedRef && _s(r.linkedRef).indexOf('BILL-') === 0) return;
    total += _n(r.amount) || 0;
  });
  // Allocation-aware (Option B). Scales a bill's TAXABLE/AMOUNT base by each
  // allocation's share of the bill's gross AMOUNT, so ex-GST figures stay
  // proportional. Bills with no allocations keep legacy Bill.Project matching.
  var allocMapEx = (typeof _getAllBillAllocationsMap === 'function') ? _getAllBillAllocationsMap() : {};
  var billsSheet = _getSS().getSheetByName(AP.BILLS);
  if (billsSheet && billsSheet.getLastRow() > 1){
    var bW = billsSheet.getLastColumn();
    var bData = billsSheet.getRange(2, 1, billsSheet.getLastRow()-1, bW).getValues();
    bData.forEach(function(r){
      if (_s(r[BC.STATUS]) === BILL_STATUS.CANCELLED) return;
      if (fy && _s(r[BC.FY]) !== fy) return;
      if (!_billCountsForProject(r)) return;
      var bid     = _s(r[BC.ID]);
      var baseAmt = _n(r[BC.TAXABLE]) || _n(r[BC.AMOUNT]) || 0;  // ex-GST preferred
      var gross   = _n(r[BC.AMOUNT]) || 0;
      var allocs  = allocMapEx[bid];
      if (allocs && allocs.length){
        // Scale base proportionally to each allocation's share of gross.
        allocs.forEach(function(a){
          if (a.project !== project) return;
          var share = gross > 0 ? (a.amount / gross) : 0;
          total += baseAmt * share;
        });
      } else {
        if (_s(r[BC.PROJECT]) !== project) return;
        total += baseAmt;
      }
    });
  }
  return _round(total);
}

// ============================================================
// SHEET-EDIT ERGONOMICS (Oct 2026)
// Three jobs:
//  1. Make Category dropdowns LIVE — point at M_Category!A2:A so new
//     categories added there appear instantly, no re-run needed.
//  2. Remove Group dropdown. Fill Group from Category via onEdit trigger.
//  3. Fix Personal_Transactions dropdowns (Owner / Class / Direction /
//     Category / Source / Counter_Account / Payment_Mode / Deleted).
// ============================================================

// Range-based dropdown — the dropdown auto-updates when the source range
// is edited. Use for Category since M_Category is the live source of truth.
function _validateColRange(sheet, col, rangeString){
  if (!sheet || !rangeString) return;
  try {
    var ss = sheet.getParent();
    var src = ss.getRange(rangeString);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(src, true)
      .setAllowInvalid(true)   // soft — don't block paste of unknown values
      .build();
    sheet.getRange(2, col, Math.max(sheet.getMaxRows()-1, 1), 1).setDataValidation(rule);
  } catch(e){ Logger.log('_validateColRange col '+col+': '+e.message); }
}

// Strip any existing data validation from a column (for Group — we want no
// dropdown there, the value is filled by the onEdit trigger).
function _clearColValidation(sheet, col){
  try {
    sheet.getRange(2, col, Math.max(sheet.getMaxRows()-1, 1), 1).setDataValidation(null);
  } catch(e){ Logger.log('_clearColValidation col '+col+': '+e.message); }
}

// Build Category → Group map once per onEdit call.
// Reads M_Category directly so it tracks whatever's live in the sheet.
function _buildCatToGroupMap(){
  var m = {};
  try {
    var ss = _getSS();
    var sh = ss.getSheetByName('M_Category');
    if (!sh || sh.getLastRow() < 2) return m;
    var data = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();  // A=Category, B=Group
    data.forEach(function(r){
      var cat = _s(r[0]), grp = _s(r[1]);
      if (cat && grp) m[cat] = grp;
    });
  } catch(e){ Logger.log('_buildCatToGroupMap: '+e.message); }
  return m;
}

// Installable onEdit trigger — fires when ANY cell in the spreadsheet is edited.
// Watches the Category column on every Transactions_<FY> sheet. When
// Category changes, writes the matching Group into the Group column.
//
// Handles ALL edit shapes:
//   - single cell typed             (e.value populated)
//   - single cell picked from dropdown
//   - multi-cell paste              (e.value undefined, walk the range)
//   - drag-fill / fill-down         (same, multi-row)
//   - cleared cell                  (clears Group too)
// Matching is case-insensitive and whitespace-tolerant so a stray space
// in M_Category or in the pasted value doesn't silently skip the row.
function onEdit(e){
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var name  = sheet.getName();

    // Route by sheet: Transactions_<FY> uses TC.CATEGORY / TC.GROUP; Bills
    // uses BC.CATEGORY / BC.GROUP. Everything else is a no-op.
    var catCol, grpCol;
    if (name.indexOf(T.TXN_PREFIX) === 0){
      catCol = TC.CATEGORY+1; grpCol = TC.GROUP+1;
    } else if (name === (typeof AP !== 'undefined' ? AP.BILLS : 'Bills')){
      // BC lives in Bills.gs; shared globals in GAS.
      catCol = (typeof BC !== 'undefined' ? BC.CATEGORY : 16) + 1;
      grpCol = (typeof BC !== 'undefined' ? BC.GROUP    : 24) + 1;
    } else {
      return;
    }

    _onEditCategoryToGroup(e, sheet, catCol, grpCol);
  } catch(err){
    try { Logger.log('onEdit: '+err.message); } catch(_){}
  }
}

// Shared autofill: when the Category column of a row is edited, write the
// matching Group from M_Category into the row's Group column. Handles multi-
// cell edits (paste/fill), case/whitespace variants, and clearing. One batched
// read + one batched write per event.
function _onEditCategoryToGroup(e, sheet, catCol, grpCol){
  var firstCol = e.range.getColumn();
  var lastCol  = firstCol + e.range.getNumColumns() - 1;
  if (catCol < firstCol || catCol > lastCol) return;
  var catColOffset = catCol - firstCol;

  var firstRow = e.range.getRow();
  var numRows  = e.range.getNumRows();
  if (firstRow < 2 && numRows === 1) return;   // header-only edit
  var startRow = Math.max(firstRow, 2);
  var startOffset = startRow - firstRow;
  var effectiveRows = numRows - startOffset;
  if (effectiveRows < 1) return;

  // Read just the Category column slice from the edited range (reliable for
  // paste/fill where e.value is undefined).
  var values = e.range.offset(startOffset, catColOffset, effectiveRows, 1).getValues();

  var map = _buildCatToGroupMap();
  var norm = {};
  Object.keys(map).forEach(function(k){ norm[_normCat(k)] = map[k]; });

  var out = values.map(function(row){
    var cat = _s(row[0]);
    if (!cat) return [''];
    var grp = map[cat] || norm[_normCat(cat)];
    return [grp || ''];
  });
  sheet.getRange(startRow, grpCol, effectiveRows, 1).setValues(out);
}

// Normalize a category string for lookup: trim, collapse inner whitespace,
// lower-case. Lets "Admin & Office " match "Admin & Office" and
// "ADMIN & OFFICE" if someone pasted it that way.
function _normCat(s){
  return _s(s).toLowerCase().replace(/\s+/g, ' ').trim();
}

// Owner names derived from accounts master. Used by Personal_Transactions
// Owner column validator. Falls back to a sensible default if accounts
// aren't configured yet.
function _getOwnerNames(){
  var owners = {};
  try {
    (getAccounts() || []).forEach(function(a){
      var o = _s(a.owner);
      if (o) owners[o] = true;
    });
  } catch(e){ Logger.log('_getOwnerNames: '+e.message); }
  var list = Object.keys(owners).sort();
  if (!list.length) list = ['Reeta', 'Girish', 'Company'];
  return list;
}

// Account names (for Source_of_Funds / Counter_Account dropdowns).
function _getAccountNames(){
  var names = [];
  try {
    (getAccounts() || []).forEach(function(a){
      var n = _s(a.name);
      if (n && names.indexOf(n) === -1) names.push(n);
    });
  } catch(e){ Logger.log('_getAccountNames: '+e.message); }
  return names;
}

// Apply the correct dropdowns to the Personal_Transactions sheet.
// Overwrites any prior validation on every column we touch — fixes the
// "Yes/No on Source_of_Funds" and "personal categories on Owner" bugs.
function _applyPersonalDropdowns(sheet){
  if (!sheet) sheet = _getSS().getSheetByName(T.PERSONAL);
  if (!sheet) return;
  var owners   = _getOwnerNames();
  var accounts = _getAccountNames();
  var pmodes   = [];
  try { pmodes = getPaymentModes() || []; } catch(e){}

  // Hard-list validators (soft = allowInvalid)
  var jobs = [
    [PC.OWNER+1,           owners],
    [PC.CLASS+1,           ['Expense','Fund In','Fund Out','Internal']],
    [PC.DIRECTION+1,       ['IN','OUT']],
    [PC.SOURCE+1,          accounts],
    [PC.COUNTER_ACCOUNT+1, accounts],
    [PC.PAY_MODE+1,        pmodes],
    [PC.DELETED+1,         ['Yes','No']]
  ];
  jobs.forEach(function(j){
    try { _validateColSoft(sheet, j[0], j[1]); }
    catch(e){ Logger.log('personal dropdown col '+j[0]+': '+e.message); }
  });

  // Category — range-based (live from M_Category_Personal). Falls back to
  // M_Category if the personal sheet doesn't exist.
  var ss = _getSS();
  var personalCatSheet = ss.getSheetByName('M_Category_Personal');
  var catRange = personalCatSheet
    ? "'M_Category_Personal'!A2:A200"
    : "'M_Category'!A2:A1000";
  _validateColRange(sheet, PC.CATEGORY+1, catRange);
}

// Rewrite _applyTxnDropdowns to use live-range Category and no Group dropdown.
// Kept the old function name — just redirect to the new behavior.
function _applyTxnDropdownsLive(sheet){
  if (!sheet || typeof sheet.getRange !== 'function') return;
  // Category — live from M_Category!A2:A
  _validateColRange(sheet, TC.CATEGORY+1, "'M_Category'!A2:A1000");
  // Group — no dropdown; filled by onEdit from Category
  _clearColValidation(sheet, TC.GROUP+1);

  // Everything else — unchanged (hard-list dropdowns from masters)
  var pmodes=[], sources=[], employees=[], projects=[];
  try { pmodes    = getPaymentModes()||[]; } catch(e){}
  try { sources   = (getSourcesOfFunds()||[]).map(function(s){ return s && s.value; }).filter(Boolean); } catch(e){}
  try { employees = (getEmployeeNames()||[]).slice(); } catch(e){}
  try { projects  = (getProjectNames()||[]).slice(); } catch(e){}
  var submitters = ['Owner'].concat(employees);
  var onBehalf   = ['Owner','Company'].concat(employees);

  var listJobs = [
    [TC.SUBMITTED_BY+1, submitters],
    [TC.ON_BEHALF_OF+1, onBehalf],
    [TC.PROJECT+1,      projects],
    [TC.LOCATION+1,     EXPENSE_LOCATIONS],
    [TC.GST_RATE+1,     GST_RATES],
    [TC.PAY_MODE+1,     pmodes],
    [TC.SOURCE+1,       sources],
    [TC.SETTLE_MODE+1,  pmodes],
    [TC.FY+1,           FY_CHOICES]
  ];
  listJobs.forEach(function(p){ try { _validateColSoft(sheet, p[0], p[1]); } catch(e){ Logger.log('list-dropdown col '+p[0]+': '+e.message); } });
  [TC.TX_DATE+1, TC.SUB_DATE+1, TC.SETTLE_DATE+1].forEach(function(c){
    try { _validateColDateSoft(sheet, c); } catch(e){ Logger.log('date-validator col '+c+': '+e.message); }
  });
  [TC.AMOUNT+1, TC.TAXABLE+1, TC.CGST+1, TC.SGST+1, TC.IGST+1, TC.SETTLE_AMT+1].forEach(function(c){
    try { _validateColNumberSoft(sheet, c); } catch(e){ Logger.log('number-validator col '+c+': '+e.message); }
  });
}

// One-shot repair you run from the Apps Script editor whenever dropdowns
// look off — refreshes every Transactions_<FY> tab + Personal_Transactions.
function fixAllDropdowns(){
  var ss = _getSS();
  var touched = [];

  // Transactions_<FY> sheets
  ss.getSheets().forEach(function(sh){
    if (sh.getName().indexOf(T.TXN_PREFIX) === 0){
      try { _applyTxnDropdownsLive(sh); touched.push(sh.getName()); }
      catch(e){ Logger.log('fixAllDropdowns '+sh.getName()+': '+e.message); }
    }
  });

  // Personal_Transactions
  var psh = ss.getSheetByName(T.PERSONAL);
  if (psh){
    try { _applyPersonalDropdowns(psh); touched.push(T.PERSONAL); }
    catch(e){ Logger.log('fixAllDropdowns '+T.PERSONAL+': '+e.message); }
  }

  // Bills — live Category dropdown + clear Group (onEdit fills Group).
  var bsh = ss.getSheetByName(typeof AP !== 'undefined' ? AP.BILLS : 'Bills');
  if (bsh){
    try {
      _validateColRange(bsh, (typeof BC !== 'undefined' ? BC.CATEGORY : 16) + 1, "'M_Category'!A2:A1000");
      _clearColValidation(bsh, (typeof BC !== 'undefined' ? BC.GROUP : 24) + 1);
      touched.push('Bills');
    } catch(e){ Logger.log('fixAllDropdowns Bills: '+e.message); }
  }

  var msg = 'Dropdowns refreshed on: ' + touched.join(', ') +
    '\n\nCategory column now points at M_Category live — new categories added there appear automatically.' +
    '\nGroup column dropdown removed — Group auto-fills from Category via onEdit trigger.' +
    '\nPersonal_Transactions Owner / Source / Payment_Mode / Category all have correct validators now.' +
    '\nBills Category is now live from M_Category; Group auto-fills via onEdit.';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, touched:touched};
}

// Bulk backfill — walks every Transactions_<FY> sheet, reads Category on
// each row, writes the matching Group from M_Category. Use for historical
// rows that pre-date the onEdit trigger, or whenever you want to re-sync
// Group with Category after bulk M_Category changes.
//
// Rows with blank Category are left alone. Rows whose Category isn't in
// M_Category get their Group cleared AND are reported so you can fix the
// Category (or add the missing category to M_Category).
function backfillGroupsFromCategory(){
  var map = _buildCatToGroupMap();
  var norm = {};
  Object.keys(map).forEach(function(k){ norm[_normCat(k)] = map[k]; });

  var ss = _getSS();
  var report = [], totalWritten = 0, totalUnknown = 0;

  function _walk(sheet, catCol, grpCol){
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    // Expand sheet width if the Group column doesn't exist yet (fresh schema).
    if (sheet.getLastColumn() < grpCol){
      Logger.log('backfillGroupsFromCategory: ' + sheet.getName() + ' is missing Group column — run setupBillsModule / setupMasters first.');
      return;
    }
    var cats = sheet.getRange(2, catCol, lastRow-1, 1).getValues();
    var grps = sheet.getRange(2, grpCol, lastRow-1, 1).getValues();

    var writes = 0, unknowns = [];
    var out = cats.map(function(r, i){
      var cat = _s(r[0]);
      var curGrp = _s(grps[i][0]);
      if (!cat) return [curGrp];
      var newGrp = map[cat] || norm[_normCat(cat)] || '';
      if (!newGrp){
        unknowns.push({row: i+2, category: cat});
        if (curGrp !== '') writes++;
        return [''];
      }
      if (newGrp !== curGrp) writes++;
      return [newGrp];
    });

    if (writes > 0) sheet.getRange(2, grpCol, out.length, 1).setValues(out);
    totalWritten += writes;
    totalUnknown += unknowns.length;
    report.push({sheet: sheet.getName(), rows: cats.length, written: writes,
                 unknownCount: unknowns.length, unknowns: unknowns.slice(0, 10)});
  }

  // Transactions_<FY>
  ss.getSheets().forEach(function(sheet){
    if (sheet.getName().indexOf(T.TXN_PREFIX) !== 0) return;
    _walk(sheet, TC.CATEGORY+1, TC.GROUP+1);
  });
  // Bills — same rule
  var bsh = ss.getSheetByName(typeof AP !== 'undefined' ? AP.BILLS : 'Bills');
  if (bsh){
    _walk(bsh,
          (typeof BC !== 'undefined' ? BC.CATEGORY : 16) + 1,
          (typeof BC !== 'undefined' ? BC.GROUP    : 24) + 1);
  }

  // Build the summary message
  var lines = ['Group backfill complete.',
    'Total rows updated: ' + totalWritten,
    'Unknown categories found: ' + totalUnknown,
    ''];
  report.forEach(function(r){
    lines.push(r.sheet + ' — scanned ' + r.rows + ', updated ' + r.written +
               (r.unknownCount ? ', unknown ' + r.unknownCount : ''));
    if (r.unknownCount){
      r.unknowns.forEach(function(u){
        lines.push('  row ' + u.row + ': "' + u.category + '"');
      });
      if (r.unknownCount > 10) lines.push('  …and ' + (r.unknownCount - 10) + ' more');
    }
  });
  var msg = lines.join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, written:totalWritten, unknown:totalUnknown, report:report};
}

// One-time: installs the onEdit trigger so Category → Group auto-fill works
// on direct sheet edits. Safe to re-run (checks for existing trigger first).
function installOnEditTrigger(){
  var triggers = ScriptApp.getProjectTriggers();
  var existing = triggers.filter(function(t){
    return t.getHandlerFunction() === 'onEdit'
        && t.getEventType() === ScriptApp.EventType.ON_EDIT;
  });
  if (existing.length){
    var m0 = 'onEdit trigger already installed (' + existing.length + '). Nothing to do.';
    try { SpreadsheetApp.getUi().alert(m0); } catch(e){ Logger.log(m0); }
    return {success:true, message:m0};
  }
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(_getSS())
    .onEdit()
    .create();
  var m1 = 'onEdit trigger installed. Category edits in Transactions_<FY> sheets will now auto-fill the Group column.';
  try { SpreadsheetApp.getUi().alert(m1); } catch(e){ Logger.log(m1); }
  return {success:true, message:m1};
}