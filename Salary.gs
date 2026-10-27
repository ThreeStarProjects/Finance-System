// ============================================================
// 3 Star Projects — Unified Finance System
// PHASE 5 : PAYROLL  (+ Employee & Project masters)
// File: Salary.gs   (depends on Masters, Transactions)
//
// Sheets:
//   M_Employee — Name | Email | Phone | Designation | Gross_Monthly_Salary | Active
//   M_Project  — Code | Name | Active
//   Salary     — one row per (employee, month) salary computation
//
// Salary save also AUTO-RECOVERS that employee's outstanding advances:
//   advances are Owner transactions (On_Behalf_Of = employee, Recovery = Pending)
//   across the FY tabs → flipped to Recovered + stamped with the salary record ID.
//
// Entry point (optional): run setupPayroll() to pre-create the three tabs.
// ============================================================

var P = { EMPLOYEE:'M_Employee', PROJECT:'M_Project', SALARY:'Salary' };

var EMC = { NAME:0, EMAIL:1, PHONE:2, DESIG:3, GROSS:4, ACTIVE:5, HOLDS_FLOAT:6, DAILY_RATE:7 };
var EMC_TOTAL = 8;
var EMPLOYEE_HEADERS = ['Name','Email','Phone','Designation','Gross_Monthly_Salary','Active','Holds_Float','Daily_Rate'];
var EMPLOYEE_WIDTHS = [220,240,140,180,180,80,110,110];

var PRC = { CODE:0, NAME:1, ACTIVE:2 };
var PRC_TOTAL = 3;
var PROJECT_HEADERS = ['Project_Code','Project_Name','Active'];
var PROJECT_WIDTHS = [130,320,80];

var SC = { ID:0, EMP:1, MONTH:2, FY:3, GROSS:4, DAILY:5, MONTH_DAYS:6, PRESENT:7, ABSENT:8,
           OT_HOURS:9, OT_EARN:10, ADV_DED:11, PENALTY:12, PEN_OT:13, PEN_SAL:14, ABS_DED:15,
           HALF_DAYS:16, HALF_DED:17, PENALTY_HOURS:18, PENALTY_AMOUNT:19, NET:20, CALC_ON:21, MODE:22 };
var SC_TOTAL = 23;
var SALARY_HEADERS = ['Salary_Record_ID','Employee','Month','FY','Gross','Daily_Rate','Month_Days',
  'Days_Present','Days_Absent','OT_Hours','OT_Earnings','Advance_Deducted','Penalty','Penalty_OT',
  'Penalty_Salary','Absent_Deduction','Half_Days','Half_Day_Deduction','Penalty_Hours','Penalty_Amount','Net_Salary','Calculated_On','Mode'];
var SALARY_WIDTHS = [190,150,80,80,120,110,100,110,100,90,110,130,100,110,120,130,90,140,100,100,120,160,80];

// ============================================================
// PROVISIONING
// ============================================================
function setupPayroll(){
  var ss = _getSS();
  // Capture the schema state BEFORE _buildSheet / _validateCol calls — these
  // can extend the column count of an existing narrow sheet by applying
  // validation to a previously-empty column, masking the upgrade signal.
  var existing = ss.getSheetByName(P.EMPLOYEE);
  var needsHoldsFloatUpgrade = false;
  var needsDailyRateUpgrade  = false;
  if (existing){
    var headers = existing.getRange(1, 1, 1, Math.max(existing.getLastColumn(), 1)).getValues()[0];
    needsHoldsFloatUpgrade = headers.indexOf('Holds_Float') === -1;
    needsDailyRateUpgrade  = headers.indexOf('Daily_Rate')  === -1;
  }

  var emp = _buildSheet(ss, P.EMPLOYEE, EMPLOYEE_HEADERS, EMPLOYEE_WIDTHS);
  _validateCol(emp, EMC.ACTIVE+1, ['Yes','No']);

  // Apply the Holds_Float upgrade if the header was missing from the live sheet
  if (needsHoldsFloatUpgrade){
    emp.getRange(1, EMC.HOLDS_FLOAT+1).setValue('Holds_Float').setFontWeight('bold');
    emp.setColumnWidth(EMC.HOLDS_FLOAT+1, 110);
    var lr = emp.getLastRow();
    if (lr >= 2){
      var defaults = new Array(lr-1).fill(['No']);
      emp.getRange(2, EMC.HOLDS_FLOAT+1, lr-1, 1).setValues(defaults);
    }
  }
  // Validate the Holds_Float column (after upgrade so the header sits next to it)
  _validateCol(emp, EMC.HOLDS_FLOAT+1, ['Yes','No']);

  // Apply the Daily_Rate upgrade if the header was missing.
  // Historical rows stay blank (means: use monthly-gross calculation as before).
  if (needsDailyRateUpgrade){
    emp.getRange(1, EMC.DAILY_RATE+1).setValue('Daily_Rate').setFontWeight('bold');
    emp.setColumnWidth(EMC.DAILY_RATE+1, 110);
    // No back-fill — blank means "not a daily-rate employee".
  }

  var proj = _buildSheet(ss, P.PROJECT, PROJECT_HEADERS, PROJECT_WIDTHS);
  _validateCol(proj, PRC.ACTIVE+1, ['Yes','No']);
  var sal = _buildSheet(ss, P.SALARY, SALARY_HEADERS, SALARY_WIDTHS);
  sal.getRange(2, SC.MONTH+1, 5000, 1).setNumberFormat('@STRING@');
  var msg = 'Payroll ready:\n• '+P.EMPLOYEE+'\n• '+P.PROJECT+'\n• '+P.SALARY +
            (needsHoldsFloatUpgrade ? '\n\nAdded new column: Holds_Float (back-filled "No" on existing rows).' : '');
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
  return {success:true, message:msg, upgraded:needsHoldsFloatUpgrade};
}

// ============================================================
// EMPLOYEE MASTER
// ============================================================
function getEmployeeNames(){
  try {
    var sheet = _getSS().getSheetByName(P.EMPLOYEE);
    if (!sheet || sheet.getLastRow()<2) return [];
    var w = Math.min(sheet.getLastColumn(), EMC_TOTAL);
    var names = sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues()
      .filter(function(r){ return _s(r[EMC.NAME]) && _s(r[EMC.ACTIVE]) !== 'No'; })
      .map(function(r){ return _s(r[EMC.NAME]); });
    names.sort(function(a,b){
      var x = a.toLowerCase(), y = b.toLowerCase();
      return x < y ? -1 : x > y ? 1 : 0;
    });
    return names;
  } catch(e){ Logger.log('getEmployeeNames: '+e.message); return []; }
}
function getEmployeeMasterData(){
  try {
    var sheet = _getSS().getSheetByName(P.EMPLOYEE);
    if (!sheet || sheet.getLastRow()<2) return [];
    var w = Math.min(sheet.getLastColumn(), EMC_TOTAL);
    return sheet.getRange(2,1,sheet.getLastRow()-1,w).getValues()
      .filter(function(r){ return _s(r[EMC.NAME]); })
      .map(function(r){ return {
        name:_s(r[EMC.NAME]), email:_s(r[EMC.EMAIL]), phone:_s(r[EMC.PHONE]),
        designation:_s(r[EMC.DESIG]), grossSalary:_n(r[EMC.GROSS]),
        active:_s(r[EMC.ACTIVE])!=='No',
        holdsFloat: (w > EMC.HOLDS_FLOAT) ? (_s(r[EMC.HOLDS_FLOAT]) === 'Yes') : false,
        dailyRate:  (w > EMC.DAILY_RATE)  ? _n(r[EMC.DAILY_RATE]) : 0
      }; });
  } catch(e){ return []; }
}
function getEmployeeSalaryMap(){
  var map = {};
  getEmployeeMasterData().forEach(function(e){ map[e.name] = e.grossSalary; });
  return map;
}
function getEmployeeDailyRateMap(){
  var map = {};
  getEmployeeMasterData().forEach(function(e){ map[e.name] = e.dailyRate || 0; });
  return map;
}
function getEmployeeEmail(name){
  try {
    var emps = getEmployeeMasterData();
    var n = _s(name);
    for (var i=0; i<emps.length; i++){
      if (emps[i].name === n) return emps[i].email || null;
    }
    return null;
  } catch(e){ return null; }
}

// Look up a single employee by name from M_Employee. Returns the row object
// (name, grossSalary, dailyRate, holdsFloat, active…) or null.
function _findEmployee(name){
  var list = getEmployeeMasterData();
  var n = _s(name);
  for (var i=0; i<list.length; i++){ if (list[i].name === n) return list[i]; }
  return null;
}
function saveEmployeeMaster(employees){
  try {
    var sheet = _getSS().getSheetByName(P.EMPLOYEE);
    if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,EMC_TOTAL).clearContent();
    if (employees && employees.length)
      sheet.getRange(2,1,employees.length,EMC_TOTAL).setValues(employees.map(function(e){
        return [_s(e.name),_s(e.email),_s(e.phone),_s(e.designation),_n(e.grossSalary),
                e.active===false?'No':'Yes',
                e.holdsFloat===true?'Yes':'No',
                _n(e.dailyRate)];
      }));
    return {success:true, count:(employees||[]).length};
  } catch(e){ return {success:false, message:e.message}; }
}
function addEmployee(data){
  var emps = getEmployeeMasterData();
  if (!_s(data.name)) return {success:false, message:'Name required.'};
  var lower = _s(data.name).toLowerCase();
  for (var i=0;i<emps.length;i++){ if (emps[i].name.toLowerCase() === lower) return {success:false, message:'Employee already exists.'}; }
  emps.push({name:_s(data.name), email:_s(data.email), phone:_s(data.phone),
             designation:_s(data.designation), grossSalary:_n(data.grossSalary),
             dailyRate:_n(data.dailyRate), active:true});
  return saveEmployeeMaster(emps);
}

// ============================================================
// PROJECT MASTER
// ============================================================
function getProjectNames(){
  try {
    var sheet = _getSS().getSheetByName(P.PROJECT);
    if (!sheet || sheet.getLastRow()<2) return [];
    return sheet.getRange(2,1,sheet.getLastRow()-1,PRC_TOTAL).getValues()
      .filter(function(r){ return _s(r[PRC.NAME]) && _s(r[PRC.ACTIVE]) !== 'No'; })
      .map(function(r){ var c=_s(r[PRC.CODE]), n=_s(r[PRC.NAME]); return c ? c+' — '+n : n; });
  } catch(e){ return []; }
}
function getProjectMasterData(){
  try {
    var sheet = _getSS().getSheetByName(P.PROJECT);
    if (!sheet || sheet.getLastRow()<2) return [];
    return sheet.getRange(2,1,sheet.getLastRow()-1,PRC_TOTAL).getValues()
      .filter(function(r){ return _s(r[PRC.NAME]); })
      .map(function(r){ return {code:_s(r[PRC.CODE]), name:_s(r[PRC.NAME]), active:_s(r[PRC.ACTIVE])!=='No'}; });
  } catch(e){ return []; }
}
function saveProjectMaster(projects){
  try {
    var sheet = _getSS().getSheetByName(P.PROJECT);
    if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,PRC_TOTAL).clearContent();
    if (projects && projects.length)
      sheet.getRange(2,1,projects.length,PRC_TOTAL).setValues(projects.map(function(p){
        return [_s(p.code),_s(p.name),p.active===false?'No':'Yes'];
      }));
    // Re-apply Project-column data-validation dropdowns on every target sheet
    // so newly added / renamed / deactivated projects are picked up without
    // the user needing to re-run applyProjectDropdowns from the Apps Script editor.
    // Best-effort — failure here does not fail the master save.
    try { if (typeof applyProjectDropdowns === 'function') applyProjectDropdowns(); } catch(e){}
    return {success:true, count:(projects||[]).length};
  } catch(e){ return {success:false, message:e.message}; }
}
function addProject(data){
  var projs = getProjectMasterData();
  var name = _s(data.name);
  if (!name) return {success:false, message:'Project name required.'};
  // Duplicate guard (case-insensitive)
  var key = name.toLowerCase();
  for (var i=0;i<projs.length;i++){
    if ((projs[i].name || '').toLowerCase() === key){
      return {success:false, message:'A project with this name already exists.'};
    }
  }
  projs.push({code:_s(data.code), name:name, active:true});
  return saveProjectMaster(projs);
}

// Update an existing employee, matched by original name.
//   data = {originalName, name, email, phone, designation, grossSalary, active}
function updateEmployee(data){
  var orig = _s(data.originalName || data.name);
  if (!orig) return {success:false, message:'Original employee name required.'};
  var emps = getEmployeeMasterData();
  var found = false;
  // If renaming, guard against collision with a different existing employee.
  var newName = _s(data.name) || orig;
  if (newName.toLowerCase() !== orig.toLowerCase()){
    for (var j=0;j<emps.length;j++){
      if (emps[j].name.toLowerCase() === newName.toLowerCase())
        return {success:false, message:'Another employee already has that name.'};
    }
  }
  for (var i=0;i<emps.length;i++){
    if (emps[i].name.toLowerCase() === orig.toLowerCase()){
      emps[i] = {
        name: newName,
        email: data.hasOwnProperty('email') ? _s(data.email) : emps[i].email,
        phone: data.hasOwnProperty('phone') ? _s(data.phone) : emps[i].phone,
        designation: data.hasOwnProperty('designation') ? _s(data.designation) : emps[i].designation,
        grossSalary: data.hasOwnProperty('grossSalary') ? _n(data.grossSalary) : emps[i].grossSalary,
        active: data.hasOwnProperty('active') ? !!data.active : emps[i].active,
        holdsFloat: data.hasOwnProperty('holdsFloat') ? !!data.holdsFloat : emps[i].holdsFloat,
        dailyRate: data.hasOwnProperty('dailyRate') ? _n(data.dailyRate) : (emps[i].dailyRate || 0)
      };
      found = true;
      break;
    }
  }
  if (!found) return {success:false, message:'Employee not found.'};
  var res = saveEmployeeMaster(emps);
  if (res && res.success){ res.message = 'Employee updated.'; res.name = newName; }
  return res;
}

// Update an existing project, matched by original code+name (code preferred).
//   data = {originalCode, originalName, code, name, active}
function updateProject(data){
  var oc = _s(data.originalCode), on = _s(data.originalName);
  if (!oc && !on) return {success:false, message:'Original project reference required.'};
  var projs = getProjectMasterData();
  var found = false;
  for (var i=0;i<projs.length;i++){
    var match = oc ? (projs[i].code === oc && (!on || projs[i].name === on))
                   : (projs[i].name.toLowerCase() === on.toLowerCase());
    if (match){
      projs[i] = {
        code: data.hasOwnProperty('code') ? _s(data.code) : projs[i].code,
        name: data.hasOwnProperty('name') ? (_s(data.name) || projs[i].name) : projs[i].name,
        active: data.hasOwnProperty('active') ? !!data.active : projs[i].active
      };
      found = true;
      break;
    }
  }
  if (!found) return {success:false, message:'Project not found.'};
  var res = saveProjectMaster(projs);
  if (res && res.success) res.message = 'Project updated.';
  return res;
}

// Advances all live in the Advances sheet (SalaryAdvance.gs). Readers go
// through getEmployeeOutstandingAdvances; writes from a payroll ADV_DED go
// through recordSalaryDeductionRecovery, called directly from saveSalaryRecord.

// ============================================================
// SALARY CALCULATION + SAVE
//   data = {employeeName, month, grossSalary, monthDays, daysPresent, daysAbsent,
//     halfDays, otHours, otDivisor, advanceDeducted, penaltyFromOT, penaltyFromSalary}
// ============================================================
function saveSalaryRecord(data){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Another save in progress — try again.'}; }
  try {
    var ss = _getSS(), sheet = ss.getSheetByName(P.SALARY) || (setupPayroll() && ss.getSheetByName(P.SALARY));
    var now = new Date(), calcOn = Utilities.formatDate(now,'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var emp = _s(data.employeeName), month = _s(data.month);
    if (!emp)   return {success:false, message:'Employee required.'};
    if (!month) return {success:false, message:'Month required.'};

    var gross = _n(data.grossSalary);
    // Auto-populate monthDays if not provided (calculate from month label)
    var monthDays = parseInt(data.monthDays,10) || _daysInMonth(month);
    if (monthDays <= 0) monthDays = 30;
    
    var present = _n(data.daysPresent), absent = _n(data.daysAbsent), halfDays = _n(data.halfDays);
    var otHours = _n(data.otHours), otDiv = _n(data.otDivisor) || 9;
    var advDed = _n(data.advanceDeducted);
    var penOT = _n(data.penaltyFromOT), penSal = _n(data.penaltyFromSalary);
    
    // NEW: Penalty from hours — convert to amount using daily_rate
    var penaltyHours = _n(data.penaltyHours) || 0;
    var penaltyAmount = _n(data.penaltyAmount) || 0;  // manual override if provided
    
    // Detect calculation mode by looking at the employee's Daily_Rate on M_Employee.
    //   > 0 → DAILY mode: net = daily × (present + 0.5×halfDays) + OT − adv − penalty.
    //   = 0 → MONTHLY mode (existing behavior): net = gross − (absent/half deductions) + OT − adv − penalty.
    // Kept in one function so downstream reports don't fork.
    var empData = _findEmployee(emp);
    var empDailyRate = empData ? _n(empData.dailyRate) : 0;
    var calcMode = (empDailyRate > 0) ? 'Daily' : 'Monthly';
    
    // Determine if this is an import or calculation
    var recordMode = _s(data.mode) === 'Imported' ? 'Imported' : 'Calculated';

    var dailyRate, otEarn, halfDed, absDed, net;
    if (calcMode === 'Daily'){
      dailyRate = empDailyRate;
      // OT only if user explicitly entered hours (default is 0 = no OT for daily workers)
      otEarn = (otHours > 0 && otDiv > 0) ? (dailyRate/otDiv)*otHours : 0;
      // Half days paid at 0.5 daily rate
      var halfPay = dailyRate * halfDays * 0.5;
      var presentPay = dailyRate * present;
      // No absent deduction — daily workers don't have a monthly baseline
      absDed = 0;
      halfDed = 0;
      // Calculate penalty_amount from penalty_hours if not manually provided
      if (penaltyHours > 0 && penaltyAmount === 0) penaltyAmount = (dailyRate / 8) * penaltyHours;
      var totalPenalty = penOT + penSal + penaltyAmount;
      net = presentPay + halfPay + otEarn - advDed - totalPenalty;
      // Populate gross so the Salary row stays consistent with monthly-mode reports.
      // Gross = base pay before deductions (present + half day earnings).
      gross = _round(presentPay + halfPay);
    } else {
      dailyRate = monthDays > 0 ? gross/monthDays : 0;
      otEarn = otDiv > 0 ? (dailyRate/otDiv)*otHours : 0;
      halfDed = dailyRate*halfDays*0.5;
      absDed  = dailyRate*absent;
      // Calculate penalty_amount from penalty_hours if not manually provided
      if (penaltyHours > 0 && penaltyAmount === 0) penaltyAmount = (dailyRate / 8) * penaltyHours;
      var totalPenalty = penOT + penSal + penaltyAmount;
      net = gross - absDed - halfDed + otEarn - advDed - totalPenalty;
    }

    var fy = _fyFromMonthLabel(month);
    var code = emp.replace(/\s+/g,'').substring(0,6).toUpperCase();
    var recId = 'SAL-' + Utilities.formatDate(now,'GMT+5:30','yyyyMMdd') + '-' + code;

    // Overwrite any existing record for this (employee, month)
    if (sheet.getLastRow() > 1){
      var ex = sheet.getRange(2,1,sheet.getLastRow()-1,SC_TOTAL).getValues();
      for (var i=ex.length-1;i>=0;i--){
        if (_s(ex[i][SC.EMP]) === emp && _s(ex[i][SC.MONTH]) === month) sheet.deleteRow(i+2);
      }
    }

    var row = []; for (var c=0;c<SC_TOTAL;c++) row.push('');
    row[SC.ID]=recId; row[SC.EMP]=emp; row[SC.MONTH]=month; row[SC.FY]=fy; row[SC.GROSS]=gross;
    row[SC.DAILY]=_round(dailyRate); row[SC.MONTH_DAYS]=monthDays; row[SC.PRESENT]=present;
    row[SC.ABSENT]=absent; row[SC.OT_HOURS]=otHours; row[SC.OT_EARN]=_round(otEarn);
    row[SC.ADV_DED]=advDed; row[SC.PENALTY]=_round(penOT + penSal); row[SC.PEN_OT]=penOT; row[SC.PEN_SAL]=penSal;
    row[SC.ABS_DED]=_round(absDed); row[SC.HALF_DAYS]=halfDays; row[SC.HALF_DED]=_round(halfDed);
    row[SC.PENALTY_HOURS]=_round(penaltyHours); row[SC.PENALTY_AMOUNT]=_round(penaltyAmount);
    row[SC.NET]=_round(net); row[SC.CALC_ON]=calcOn; row[SC.MODE]=recordMode;
    sheet.appendRow(row);
    SpreadsheetApp.flush();

    // Recover advances FIFO against this employee's open rows in the Advances sheet.
    if (advDed > 0 && recordMode !== 'Imported' && typeof recordSalaryDeductionRecovery === 'function'){
      recordSalaryDeductionRecovery(emp, advDed, month, recId);
    }

    // Salary-slip email intentionally removed (Tier 8.1). Owners now handle
    // slip distribution outside the system; keeping this path removed avoids
    // duplicate / stale slips going out.
    return {success:true, recordId:recId, netSalary:_round(net)};
  } catch(err){ return {success:false, message:err.message}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

function getSalaryRecords(filters){
  filters = filters || {};
  var sheet = _getSS().getSheetByName(P.SALARY);
  if (!sheet || sheet.getLastRow()<2) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,SC_TOTAL).getValues()
    .filter(function(r){
      if (!_s(r[SC.ID])) return false;
      if (filters.month && _s(r[SC.MONTH]) !== _s(filters.month)) return false;
      if (filters.employee && _s(r[SC.EMP]) !== _s(filters.employee)) return false;
      if (filters.fy && _s(r[SC.FY]) !== _s(filters.fy)) return false;
      return true;
    })
    .map(function(r){ return {recordId:_s(r[SC.ID]), employee:_s(r[SC.EMP]), month:_s(r[SC.MONTH]),
      fy:_s(r[SC.FY]), gross:_n(r[SC.GROSS]), net:_n(r[SC.NET]), calculatedOn:_s(r[SC.CALC_ON])}; });
}

function checkExistingSalaryRecord(empName, month){
  try {
    var sheet = _getSS().getSheetByName(P.SALARY);
    if (!sheet || sheet.getLastRow()<2) return {exists:false};
    var data = sheet.getRange(2,1,sheet.getLastRow()-1,SC_TOTAL).getValues();
    for (var i=0;i<data.length;i++){
      if (_s(data[i][SC.EMP]) === _s(empName) && _s(data[i][SC.MONTH]) === _s(month))
        return {exists:true, recordId:_s(data[i][SC.ID])};
    }
    return {exists:false};
  } catch(e){ return {exists:false}; }
}

// FY label from a "MMM-yy" month string
function _fyFromMonthLabel(month){
  try {
    var mm = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
    var p = _s(month).split('-'); var m = mm[p[0]], y = 2000 + parseInt(p[1],10);
    if (m === undefined || isNaN(y)) return _currentFY();
    return _fyFromDate(new Date(y, m, 1));
  } catch(e){ return _currentFY(); }
}

// ── Calculate days in month from "MMM-yy" label (e.g., "Jan-24") ────
function _daysInMonth(monthLabel){
  try {
    var mm = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
    var p = _s(monthLabel).split('-');
    var m = mm[p[0]], y = 2000 + parseInt(p[1],10);
    if (m === undefined || isNaN(y)) return 30;  // safe default
    var d = new Date(y, m+1, 0);  // day 0 of next month = last day of current month
    return d.getDate();
  } catch(e){ return 30; }
}

// ── Indian-format number (kept — used by expense confirmation email and other formatters) ─
function _fmtN(n){
  var v = parseFloat(n)||0, f = v.toFixed(2), p = f.split('.'), int = p[0], dec = p[1];
  var neg = int[0]==='-'; if (neg) int = int.substring(1);
  if (int.length>3){ var l3=int.substring(int.length-3), rest=int.substring(0,int.length-3), g=[];
    while(rest.length>2){ g.unshift(rest.substring(rest.length-2)); rest=rest.substring(0,rest.length-2); }
    if(rest.length) g.unshift(rest); int=g.join(',')+','+l3; }
  return (neg?'-':'')+int+'.'+dec;
}
// _sendSalarySlip removed (Tier 8.1). Slip emails are no longer sent by the system.