// ============================================================
// 3 Star Projects — M_Category Consolidation (Batch 13)
//
// Three one-shot functions. Run them in order from the Apps Script editor:
//
//   1. rebuildMCategorySheet()         — wipes M_Category rows and repopulates
//                                        with the final 86-category list.
//   2. migrateOldCategoriesInTxns()    — scans every Transactions_<FY> row and
//                                        rewrites the Category column from the
//                                        OLD→NEW mapping. Orphans logged.
//   3. deleteMCategoryEmployeeSheet()  — removes the legacy sheet after you
//                                        confirm no residual code references it.
//
// Idempotent: safe to re-run (will not duplicate rows).
// ============================================================

// ── NEW master list (final structure) ─────────────────────────
// [Category, Group, Direction, GST_Applicable, GST_Rate, HSN_SAC, Active, Employee_Visible]
var NEW_CATEGORIES = [
  // Office & Administration
  ['Office Rent','Office & Administration','OUT','No','','','Yes','No'],
  ['Office Electricity','Office & Administration','OUT','No','','','Yes','No'],
  ['Office Water Supply','Office & Administration','OUT','No','','','Yes','No'],
  ['Office Internet / Wi-Fi','Office & Administration','OUT','No','','','Yes','No'],
  ['Office Mobile / Telephone','Office & Administration','OUT','No','','','Yes','No'],
  ['Housekeeping','Office & Administration','OUT','No','','','Yes','Yes'],
  ['Printing & Stationery','Office & Administration','OUT','No','','','Yes','Yes'],
  ['Security Services','Office & Administration','OUT','No','','','Yes','No'],
  ['Office Repair & Maintenance','Office & Administration','OUT','No','','','Yes','Yes'],
  // Factory & Utilities
  ['Factory Rent','Factory & Utilities','OUT','No','','','Yes','No'],
  ['Factory Electricity','Factory & Utilities','OUT','No','','','Yes','No'],
  ['Water Tanker','Factory & Utilities','OUT','No','','','Yes','Yes'],
  ['PNG / Industrial Diesel','Factory & Utilities','OUT','No','','','Yes','No'],
  ['Factory Repair & Maintenance','Factory & Utilities','OUT','No','','','Yes','Yes'],
  // Direct Production
  ['Raw Material Purchase','Direct Production','OUT','No','','','Yes','Yes'],
  ['Packing Material','Direct Production','OUT','No','','','Yes','Yes'],
  ['Consumables','Direct Production','OUT','No','','','Yes','Yes'],
  ['Machinery Purchase','Direct Production','OUT','No','','','Yes','No'],
  ['Tools & Equipment','Direct Production','OUT','No','','','Yes','Yes'],
  ['Machinery Repair & Maintenance','Direct Production','OUT','No','','','Yes','Yes'],
  // Labour & Contractor
  ['Contractor Payment','Labour & Contractor','OUT','No','','','Yes','Yes'],
  ['Sub-contractor Payment','Labour & Contractor','OUT','No','','','Yes','Yes'],
  ['Daily Labour Charges','Labour & Contractor','OUT','No','','','Yes','Yes'],
  ['Heavy Labour / Loading-Unloading','Labour & Contractor','OUT','No','','','Yes','Yes'],
  // Employee & Staff
  ['Employee Monthly Salary','Employee & Staff','OUT','No','','','Yes','No'],
  ["Director's Salary",'Employee & Staff','OUT','No','','','Yes','No'],
  ['Bonus','Employee & Staff','OUT','No','','','Yes','No'],
  ['Salary Advance','Employee & Staff','OUT','No','','','Yes','No'],
  ['Food / Lunch / Dinner','Employee & Staff','OUT','No','','','Yes','Yes'],
  ['Tea, Coffee, Kitchen Supplies','Employee & Staff','OUT','No','','','Yes','Yes'],
  ['Employee Medical','Employee & Staff','OUT','No','','','Yes','No'],
  ['Employee Welfare','Employee & Staff','OUT','No','','','Yes','No'],
  ['Employee Transportation','Employee & Staff','OUT','No','','','Yes','Yes'],
  // Statutory & Payroll
  ['PF (Provident Fund)','Statutory & Payroll','OUT','No','','','Yes','No'],
  ['ESIC','Statutory & Payroll','OUT','No','','','Yes','No'],
  ['Gratuity','Statutory & Payroll','OUT','No','','','Yes','No'],
  ['Leave Encashment','Statutory & Payroll','OUT','No','','','Yes','No'],
  ['Professional Tax','Statutory & Payroll','OUT','No','','','Yes','No'],
  // Professional & Legal
  ['CA / Accountant Fees','Professional & Legal','OUT','No','','','Yes','No'],
  ['Legal Fees','Professional & Legal','OUT','No','','','Yes','No'],
  ['Consultancy Charges','Professional & Legal','OUT','No','','','Yes','No'],
  ['Lab / Material Testing','Professional & Legal','OUT','No','','','Yes','No'],
  // Taxes & Compliance
  ['GST Paid','Taxes & Compliance','OUT','No','','','Yes','No'],
  ['TDS Paid','Taxes & Compliance','OUT','No','','','Yes','No'],
  ['Income Tax / Advance Tax','Taxes & Compliance','OUT','No','','','Yes','No'],
  ['License / Renewal Fees','Taxes & Compliance','OUT','No','','','Yes','No'],
  ['Property Tax','Taxes & Compliance','OUT','No','','','Yes','No'],
  ['Late Fee / Penalty','Taxes & Compliance','OUT','No','','','Yes','No'],
  // Banking & Finance
  ['Bank Charges','Banking & Finance','OUT','No','','','Yes','No'],
  ['Bank Interest Paid','Banking & Finance','OUT','No','','','Yes','No'],
  ['Loan Repayment — Interest','Banking & Finance','OUT','No','','','Yes','No'],
  ['Loan Repayment — Principal','Banking & Finance','OUT','No','','','Yes','No'],
  ['OD / CC Utilization Cost','Banking & Finance','OUT','No','','','Yes','No'],
  ['Advance Given to Vendor','Banking & Finance','OUT','No','','','Yes','No'],
  // Logistics & Transport
  ['Freight & Forwarding (outbound)','Logistics & Transport','OUT','No','','','Yes','Yes'],
  ['Inbound Transport (material in)','Logistics & Transport','OUT','No','','','Yes','Yes'],
  ['Petrol / Diesel / CNG','Logistics & Transport','OUT','No','','','Yes','Yes'],
  ['Vehicle Running & Maintenance','Logistics & Transport','OUT','No','','','Yes','Yes'],
  ['Vehicle Purchase / Loan / Rent','Logistics & Transport','OUT','No','','','Yes','No'],
  ['Courier Charges','Logistics & Transport','OUT','No','','','Yes','Yes'],
  ['Toll, Parking, RTO','Logistics & Transport','OUT','No','','','Yes','Yes'],
  // IT & Software
  ['Computer / Laptop','IT & Software','OUT','No','','','Yes','No'],
  ['IT Repair & Maintenance','IT & Software','OUT','No','','','Yes','Yes'],
  ['Software Subscriptions','IT & Software','OUT','No','','','Yes','No'],
  // Insurance
  ['Vehicle Insurance','Insurance','OUT','No','','','Yes','No'],
  ['Employee Insurance','Insurance','OUT','No','','','Yes','No'],
  ['Property / Factory Insurance','Insurance','OUT','No','','','Yes','No'],
  // Project / Site Expenses
  ['Site Transport','Project / Site Expenses','OUT','No','','','Yes','Yes'],
  ['Site Food','Project / Site Expenses','OUT','No','','','Yes','Yes'],
  ['Site Stay','Project / Site Expenses','OUT','No','','','Yes','Yes'],
  ['Site Safety','Project / Site Expenses','OUT','No','','','Yes','Yes'],
  ['Equipment / Shuttering Rental','Project / Site Expenses','OUT','No','','','Yes','Yes'],
  ['Site Permissions / Municipal Fees','Project / Site Expenses','OUT','No','','','Yes','No'],
  ['Architect / Design Fees','Project / Site Expenses','OUT','No','','','Yes','No'],
  ['Site Visit Expenses','Project / Site Expenses','OUT','No','','','Yes','Yes'],
  // Sales & Marketing
  ['Marketing / Advertising','Sales & Marketing','OUT','No','','','Yes','No'],
  ['Website / Domain / SEO','Sales & Marketing','OUT','No','','','Yes','No'],
  ['Commission Paid','Sales & Marketing','OUT','No','','','Yes','No'],
  ['Client Entertainment','Sales & Marketing','OUT','No','','','Yes','No'],
  ['Expo / Exhibition Costs','Sales & Marketing','OUT','No','','','Yes','No'],
  ['Sales Tour Travel','Sales & Marketing','OUT','No','','','Yes','No'],
  // Income
  ['Service Income','Income','IN','No','','','Yes','No'],
  ['Sales — Interior Execution / Project','Income','IN','No','','','Yes','No'],
  ['Sales — Furniture','Income','IN','No','','','Yes','No'],
  ['Discount Received','Income','IN','No','','','Yes','No'],
  ['Interest Income','Income','IN','No','','','Yes','No'],
  ['Other Income','Income','IN','No','','','Yes','No']
];

// ── OLD → NEW mapping for historical transactions ────────────
// Keys are the OLD category names (exact spelling in current M_Category).
// Values are the NEW category names.
// Entries mapped to '__ORPHAN__' or missing from this map stay unchanged and
// are reported — you fix them manually, usually because they could legitimately
// split in two ways (e.g. "Electricity Bill" → Office or Factory?).
var CATEGORY_MIGRATION_MAP = {
  // Admin & Office
  'Factory / Office Insurance': 'Property / Factory Insurance',
  'Property Insurance': 'Property / Factory Insurance',
  'Housekeeping (supplies & services)': 'Housekeeping',
  'Internet / Wi-Fi': 'Office Internet / Wi-Fi',
  'IT Purchase / Repair / Maintenance': 'IT Repair & Maintenance',
  'Mobile / Telephone Bills': 'Office Mobile / Telephone',
  'Repair & Maintenance': 'Office Repair & Maintenance',
  'Supplies / Misc': 'Printing & Stationery',
  'Rent Paid (office/factory)': '__ORPHAN__',  // can't auto-split office vs factory
  'Security Services': 'Security Services',
  'Software Subscriptions': 'Software Subscriptions',
  'Stationery & Printing': 'Printing & Stationery',
  'Other Admin (Specify)': '__ORPHAN__',

  // Bank, Finance & Loans
  'Advance Given to Vendor': 'Advance Given to Vendor',
  'Bank Charges': 'Bank Charges',
  'Bank Interest Paid': 'Bank Interest Paid',
  'Loan Repayment — Interest': 'Loan Repayment — Interest',
  'Loan Repayment — Principal': 'Loan Repayment — Principal',
  'OD / CC Utilization Cost': 'OD / CC Utilization Cost',
  'Other Finance (Specify)': '__ORPHAN__',

  // Direct Production — including dropped manufacturing categories
  'Consumables (fasteners, packaging)': 'Consumables',
  'Daily Labour Charges': 'Daily Labour Charges',
  'Die / Mould Development & Repair': '__ORPHAN__',   // DROPPED
  'Heavy Labour / Loading-Unloading': 'Heavy Labour / Loading-Unloading',
  'Job Work Charges': '__ORPHAN__',                   // DROPPED
  'Machinery Purchase': 'Machinery Purchase',
  'Machinery Repair & Maintenance': 'Machinery Repair & Maintenance',
  'Other Production (Specify)': '__ORPHAN__',
  'Packing Material': 'Packing Material',
  'Parts / Components Purchase': 'Raw Material Purchase',
  'Plant Repair & Maintenance': 'Factory Repair & Maintenance',
  'Raw Material Purchase': 'Raw Material Purchase',
  'Raw Material Testing': 'Lab / Material Testing',
  'Tooling & Drill Bits': 'Tools & Equipment',
  'Vendor Parts (Wheels, Seats, etc.)': '__ORPHAN__', // DROPPED
  'Vendor Payment': '__ORPHAN__',                     // too generic — pick manually
  'CO2 / O2 / Argon Cylinders': '__ORPHAN__',         // DROPPED
  'Contractor Payment': 'Contractor Payment',

  // Income
  'Discount Received': 'Discount Received',
  'Interest Income': 'Interest Income',
  'Other Income': 'Other Income',
  'Sales — Furniture': 'Sales — Furniture',
  'Sales — Interior Execution / Project': 'Sales — Interior Execution / Project',
  'Service Income': 'Service Income',

  // Logistics & Transport
  'Courier Charges': 'Courier Charges',
  'Employee Travel (Bus/Train/Auto)': 'Employee Transportation',
  'Freight & Forwarding (outbound)': 'Freight & Forwarding (outbound)',
  'Inbound Transport (material in)': 'Inbound Transport (material in)',
  'Other Transport (Specify)': '__ORPHAN__',
  'Petrol / Diesel / CNG': 'Petrol / Diesel / CNG',
  'Toll, Parking, RTO': 'Toll, Parking, RTO',
  'Vehicle Purchase / Loan / Rent': 'Vehicle Purchase / Loan / Rent',
  'Vehicle Running & Maintenance': 'Vehicle Running & Maintenance',

  // Professional Services
  'CA / Accountant Fees': 'CA / Accountant Fees',
  'Consultancy Charges (general)': 'Consultancy Charges',
  'Lab / Material Testing': 'Lab / Material Testing',
  'Legal Fees': 'Legal Fees',
  'Other Professional (Specify)': '__ORPHAN__',
  'R&D Expense': 'Consultancy Charges',

  // Salary & Wages
  'Bonus': 'Bonus',
  "Director's Salary": "Director's Salary",
  'Employee Insurance': 'Employee Insurance',
  'Employee Medical': 'Employee Medical',
  'Employee Monthly Salary': 'Employee Monthly Salary',
  'Employee Welfare': 'Employee Welfare',
  'ESIC': 'ESIC',
  'Gratuity': 'Gratuity',
  'Leave Encashment': 'Leave Encashment',
  'Other Salary (Specify)': '__ORPHAN__',
  'PF (Provident Fund)': 'PF (Provident Fund)',
  'Salary Advance': 'Salary Advance',

  // Sales & Marketing
  'Site Expenses': 'Site Visit Expenses',
  'Client Entertainment': 'Client Entertainment',
  'Commission Paid': 'Commission Paid',
  'Expo / Exhibition Costs': 'Expo / Exhibition Costs',
  'Marketing / Advertising': 'Marketing / Advertising',
  'Other Sales (Specify)': '__ORPHAN__',
  'Sales Tour Travel (multi-day)': 'Sales Tour Travel',
  'Site Visit Expenses': 'Site Visit Expenses',
  'Website / Domain / SEO': 'Website / Domain / SEO',

  // Taxes & Compliance
  'GST Paid': 'GST Paid',
  'Income Tax / Advance Tax': 'Income Tax / Advance Tax',
  'Late Fee / Penalty (Statutory)': 'Late Fee / Penalty',
  'License / Renewal Fees': 'License / Renewal Fees',
  'Other Tax (Specify)': '__ORPHAN__',
  'Professional Tax': 'Professional Tax',
  'TDS Paid': 'TDS Paid',

  // Utilities (split into Office/Factory)
  'Electricity Bill': '__ORPHAN__',             // can't auto — office or factory?
  'Electricity Factory': 'Factory Electricity',
  'Electricity Office': 'Office Electricity',
  'Other Utilities (Specify)': '__ORPHAN__',
  'PNG / Diesel (Industrial Use)': 'PNG / Industrial Diesel',
  'Property Tax': 'Property Tax',
  'Water Supply (Bill)': 'Office Water Supply',  // best-guess default
  'Water Tanker': 'Water Tanker',
  'Fund Transfer': '__ORPHAN__'                 // system entry, keep as-is
};

// ============================================================
// 1. REBUILD M_CATEGORY SHEET
// ============================================================
function rebuildMCategorySheet(){
  var ss = _getSS();
  var sheet = ss.getSheetByName(M.CATEGORY);
  if (!sheet) throw new Error('M_Category sheet not found. Run setupMasters() first.');
  // Clear existing data AND validations on the full column range (not just
  // populated rows). The Group column has a dropdown restricted to the OLD
  // group names — writing new group names fails validation unless we clear
  // it first. Covers all 8 columns from row 2 down to maxRows.
  var maxRows = sheet.getMaxRows();
  var fullRange = sheet.getRange(2, 1, maxRows - 1, 8);
  fullRange.clearDataValidations();
  if (sheet.getLastRow() > 1){
    sheet.getRange(2, 1, sheet.getLastRow()-1, Math.max(sheet.getLastColumn(), 1)).clearContent();
  }
  // Write new rows
  sheet.getRange(2, 1, NEW_CATEGORIES.length, 8).setValues(NEW_CATEGORIES);
  // Re-apply validations with the NEW allowed values
  _validateCol(sheet, 3, ['IN','OUT']);
  _validateCol(sheet, 4, ['Yes','No']);
  _validateCol(sheet, 7, ['Yes','No']);
  _validateCol(sheet, 8, ['Yes','No']);
  // Group column: validate against the distinct new group names so future
  // manual edits stay consistent.
  var allGroups = {};
  NEW_CATEGORIES.forEach(function(r){ allGroups[r[1]] = true; });
  _validateCol(sheet, 2, Object.keys(allGroups));
  SpreadsheetApp.flush();
  var msg = 'M_Category rebuilt: ' + NEW_CATEGORIES.length + ' categories across ' +
            _distinctGroupsCount() + ' groups.';
  Logger.log(msg);
  return {success:true, count:NEW_CATEGORIES.length, message:msg};
}

function _distinctGroupsCount(){
  var g = {};
  NEW_CATEGORIES.forEach(function(r){ g[r[1]] = true; });
  return Object.keys(g).length;
}

// ============================================================
// 2. MIGRATE OLD CATEGORY NAMES IN TRANSACTIONS
// ============================================================
function migrateOldCategoriesInTxns(){
  var sheets = _allTxnSheets();
  var mapping = CATEGORY_MIGRATION_MAP;
  var newCatSet = {};
  NEW_CATEGORIES.forEach(function(r){ newCatSet[r[0]] = r[1]; });  // name → group

  var updated = 0;
  var orphanedByCategory = {};
  var alreadyNew = 0;

  sheets.forEach(function(sheet){
    if (!sheet || sheet.getLastRow() < 2) return;
    var lastRow = sheet.getLastRow();
    var catCol = TC.CATEGORY + 1;
    var grpCol = TC.GROUP + 1;
    var cats = sheet.getRange(2, catCol, lastRow-1, 1).getValues();
    var newCats = [];
    var newGrps = [];
    var changed = false;

    for (var i=0; i<cats.length; i++){
      var cur = _s(cats[i][0]);
      if (!cur){ newCats.push(['']); newGrps.push([null]); continue; }

      // Already matches a new category name — update group too in case it moved
      if (newCatSet[cur]){
        alreadyNew++;
        newCats.push([cur]);
        newGrps.push([newCatSet[cur]]);
        continue;
      }

      var target = mapping[cur];
      if (target === '__ORPHAN__' || !target){
        orphanedByCategory[cur] = (orphanedByCategory[cur] || 0) + 1;
        newCats.push([cur]);
        newGrps.push([null]);  // leave group alone
        continue;
      }

      // Map exists → rewrite
      newCats.push([target]);
      newGrps.push([newCatSet[target] || null]);
      changed = true;
      updated++;
    }

    if (changed){
      sheet.getRange(2, catCol, cats.length, 1).setValues(newCats);
      // Also update Group col for any row where new group is non-null
      var curGrps = sheet.getRange(2, grpCol, cats.length, 1).getValues();
      var finalGrps = [];
      for (var j=0; j<cats.length; j++){
        finalGrps.push([newGrps[j][0] !== null ? newGrps[j][0] : _s(curGrps[j][0])]);
      }
      sheet.getRange(2, grpCol, cats.length, 1).setValues(finalGrps);
    }
  });

  SpreadsheetApp.flush();
  var report = 'Migration complete:\n' +
               '  • Rewrote ' + updated + ' row(s) to new category names\n' +
               '  • ' + alreadyNew + ' row(s) already matched new names (updated group only)\n' +
               '  • ' + Object.keys(orphanedByCategory).length + ' distinct orphan category name(s) left alone';
  Logger.log(report);
  if (Object.keys(orphanedByCategory).length){
    Logger.log('ORPHANS — fix these manually in the sheet:');
    Object.keys(orphanedByCategory).forEach(function(k){
      Logger.log('  "' + k + '" → ' + orphanedByCategory[k] + ' row(s)');
    });
  }
  return {success:true, updated:updated, alreadyNew:alreadyNew, orphans:orphanedByCategory};
}

// ============================================================
// 3. DELETE LEGACY M_CATEGORY_EMPLOYEE SHEET
// ============================================================
function deleteMCategoryEmployeeSheet(){
  var ss = _getSS();
  var sh = ss.getSheetByName(M.CATEGORY_EMPLOYEE);
  if (!sh) return {success:true, message:'M_Category_Employee sheet already gone.'};
  ss.deleteSheet(sh);
  return {success:true, message:'M_Category_Employee sheet deleted. Employee visibility now driven by M_Category.Employee_Visible column.'};
}