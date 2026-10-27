// ============================================================
// SALARY DASHBOARD — Analytics, KPIs, Trends, Bulk Import
// ============================================================
// All backend functions for salary analytics, dashboard KPIs,
// charts, and bulk import functionality.
//
// Depends on: Salary.gs (SC constants, P.SALARY sheet ref)
// ============================================================

// ------------------------------------------------------------
// Utility: Get all salary records with optional filters
// ------------------------------------------------------------
function _getSalaryRecords(filters){
  filters = filters || {};
  // Translate the frontend's period tokens once, at the source, so every
  // caller (KPIs, breakdown, top earners, deductions, records list) uses the
  // same resolved FY. 'current' → actual FY string; 'all' or blank → no FY filter.
  var wantFy = _s(filters.fy);
  if (wantFy === 'current') wantFy = _currentFY();
  if (wantFy === 'all')     wantFy = '';

  var sheet = _getSS().getSheetByName(P.SALARY);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var w = Math.min(sheet.getLastColumn(), SC_TOTAL);
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, w).getValues();
  var records = [];

  data.forEach(function(r){
    var rec = {
      id: _s(r[SC.ID]),
      employee: _s(r[SC.EMP]),
      month: _s(r[SC.MONTH]),
      fy: _s(r[SC.FY]),
      gross: _n(r[SC.GROSS]),
      dailyRate: _n(r[SC.DAILY]),
      monthDays: _n(r[SC.MONTH_DAYS]),
      present: _n(r[SC.PRESENT]),
      absent: _n(r[SC.ABSENT]),
      otHours: _n(r[SC.OT_HOURS]),
      otEarnings: _n(r[SC.OT_EARN]),
      advanceDeducted: _n(r[SC.ADV_DED]),
      penalty: _n(r[SC.PENALTY]),
      penaltyOT: _n(r[SC.PEN_OT]),
      penaltySalary: _n(r[SC.PEN_SAL]),
      absentDeduction: _n(r[SC.ABS_DED]),
      halfDays: _n(r[SC.HALF_DAYS]),
      halfDayDeduction: _n(r[SC.HALF_DED]),
      penaltyHours: _n(r[SC.PENALTY_HOURS]),
      penaltyAmount: _n(r[SC.PENALTY_AMOUNT]),
      netSalary: _n(r[SC.NET]),
      calculatedOn: _s(r[SC.CALC_ON]),
      mode: _s(r[SC.MODE]) || 'Calculated'
    };

    // Apply filters — fy uses the translated wantFy from above.
    if (filters.month && rec.month !== filters.month) return;
    if (wantFy && rec.fy !== wantFy) return;
    if (filters.employee && rec.employee !== filters.employee) return;
    if (filters.mode && rec.mode !== filters.mode) return;

    // Month range filter (e.g., last 3 months)
    if (filters.monthsBack){
      var monthDate = _monthLabelToDate(rec.month);
      var today = new Date();
      var cutoff = new Date(today.getFullYear(), today.getMonth() - filters.monthsBack, 1);
      if (monthDate < cutoff) return;
    }

    records.push(rec);
  });

  return records;
}

// Helper: Convert "MMM-yy" to Date
function _monthLabelToDate(label){
  try {
    var mm = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
    var p = _s(label).split('-');
    var m = mm[p[0]], y = 2000 + parseInt(p[1], 10);
    if (m === undefined || isNaN(y)) return new Date(2000,0,1);
    return new Date(y, m, 1);
  } catch(e){ return new Date(2000,0,1); }
}

// Helper: Get current month label "Jan-25"
function _currentMonthLabel(){
  var d = new Date();
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return months[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2);
}

// ============================================================
// 1. SALARY KPIs (for Main Dashboard + Salary Tab Cards)
// ============================================================
function dashSalaryKpis(filters){
  try {
    filters = filters || {};
    // Honor client filters. Default fy → currentFY only when the client sent
    // no period selector at all (month/monthsBack). fy='all' means every FY.
    var fy    = _s(filters.fy) === 'all' ? '' : _s(filters.fy);
    var month = _s(filters.month) === 'all' ? '' : _s(filters.month);
    if (!fy && !month && !filters.monthsBack) fy = _currentFY();

    var queryFilters = {};
    if (fy)                    queryFilters.fy         = fy;
    if (month)                 queryFilters.month      = month;
    if (filters.monthsBack)    queryFilters.monthsBack = _n(filters.monthsBack);
    if (_s(filters.employee))  queryFilters.employee   = _s(filters.employee);
    if (_s(filters.mode))      queryFilters.mode       = _s(filters.mode);

    var records = _getSalaryRecords(queryFilters);

    // Calculate totals
    var totalSalaryPaid = 0;
    var totalGross = 0;
    var totalDeductions = 0;
    var totalOT = 0;
    var totalPenalties = 0;
    var totalAdvancesRecovered = 0;
    var employees = {};
    var calculatedCount = 0;
    var importedCount = 0;

    records.forEach(function(r){
      totalSalaryPaid += r.netSalary;
      totalGross += r.gross;
      totalDeductions += (r.absentDeduction + r.halfDayDeduction + r.penalty + r.penaltyAmount + r.advanceDeducted);
      totalOT += r.otEarnings;
      totalPenalties += (r.penalty + r.penaltyAmount);
      totalAdvancesRecovered += r.advanceDeducted;
      employees[r.employee] = true;
      if (r.mode === 'Imported') importedCount++;
      else calculatedCount++;
    });

    var employeeCount = Object.keys(employees).length;

    // Get outstanding advances (from Transactions - recovery pending).
    // Tier 4.2: use (Amount − Recovered_Amount) so partial recoveries reflect.
    // FY-scoped: when the caller passes a specific fy, only that FY's advances
    // are counted (matches dashboard-FY behaviour of every other KPI). Callers
    // wanting all-FY totals should pass fy=''.
    // Advances now live in the dedicated `Advances` sheet (see SalaryAdvance.gs).
    // Totals come from advGetSummary — no more parallel walks over Transactions,
    // no more double-counting against ADV_DED. Overrides the Salary-sheet
    // ADV_DED total above so recoveries sum from ONE source only (the Advances
    // sheet itself), avoiding the historic drift between the two.
    var advancesOutstanding = 0, advanceEmployeeCount = 0;
    try {
      var advRes = (typeof advGetSummary === 'function') ? advGetSummary(fy || 'all') : null;
      if (advRes && advRes.ok && advRes.data){
        advancesOutstanding    = _n(advRes.data.outstanding) || 0;
        totalAdvancesRecovered = _n(advRes.data.totalRecovered) || 0;
        advanceEmployeeCount   = _n(advRes.data.employeeCount) || 0;
      }
    } catch(e){ Logger.log('dashSalaryKpis advGetSummary: '+e.message); }
    var advanceEmployees = {};  // kept for the return shape; count comes from advGetSummary now
    // advanceEmployeeCount already set from advGetSummary above

    return _envOk({
      fy: fy || 'all',
      month: month || 'all',
      monthsBack: filters.monthsBack || null,
      totalSalaryPaid: _round(totalSalaryPaid),
      totalGross: _round(totalGross),
      totalDeductions: _round(totalDeductions),
      totalOT: _round(totalOT),
      totalPenalties: _round(totalPenalties),
      totalAdvancesRecovered: _round(totalAdvancesRecovered),
      employeeCount: employeeCount,
      recordsCount: records.length,
      calculatedCount: calculatedCount,
      importedCount: importedCount,
      advancesOutstanding: _round(advancesOutstanding),
      advanceEmployeeCount: advanceEmployeeCount
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 2. SALARY TREND (last N months for line chart)
// ============================================================
function dashSalaryTrend(months){
  try {
    months = months || 12;
    var records = _getSalaryRecords({});

    // Group by month
    var byMonth = {};
    records.forEach(function(r){
      if (!byMonth[r.month]) byMonth[r.month] = {salary:0, count:0, gross:0};
      byMonth[r.month].salary += r.netSalary;
      byMonth[r.month].gross += r.gross;
      byMonth[r.month].count++;
    });

    // Generate last N months labels
    var trend = [];
    var today = new Date();
    var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    for (var i = months - 1; i >= 0; i--){
      var d = new Date(today.getFullYear(), today.getMonth() - i, 1);
      var label = monthNames[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2);
      var data = byMonth[label] || {salary:0, count:0, gross:0};
      trend.push({
        month: label,
        totalSalary: _round(data.salary),
        totalGross: _round(data.gross),
        employeeCount: data.count
      });
    }

    return _envOk({trend: trend});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 3. SALARY BREAKDOWN (for pie chart)
// ============================================================
function dashSalaryBreakdown(filters){
  try {
    var records = _getSalaryRecords(filters || {});

    var basic = 0, ot = 0, deductions = 0, penalties = 0, advances = 0;

    records.forEach(function(r){
      basic += (r.gross - r.absentDeduction - r.halfDayDeduction);
      ot += r.otEarnings;
      deductions += (r.absentDeduction + r.halfDayDeduction);
      penalties += (r.penalty + r.penaltyAmount);
      advances += r.advanceDeducted;
    });

    return _envOk({
      basic: _round(basic),
      overtime: _round(ot),
      deductions: _round(deductions),
      penalties: _round(penalties),
      advances: _round(advances),
      total: _round(basic + ot)
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 4. TOP EARNERS (for bar chart)
// ============================================================
function dashTopEarners(filters, count){
  try {
    count = count || 5;
    var records = _getSalaryRecords(filters || {});

    // Group by employee
    var byEmp = {};
    records.forEach(function(r){
      if (!byEmp[r.employee]) byEmp[r.employee] = {name: r.employee, netSalary: 0, gross: 0, count: 0};
      byEmp[r.employee].netSalary += r.netSalary;
      byEmp[r.employee].gross += r.gross;
      byEmp[r.employee].count++;
    });

    var earners = [];
    for (var emp in byEmp){ earners.push(byEmp[emp]); }
    earners.sort(function(a,b){ return b.netSalary - a.netSalary; });

    return _envOk({earners: earners.slice(0, count).map(function(e){
      return {
        employee: e.name,
        netSalary: _round(e.netSalary),
        totalGross: _round(e.gross),
        monthsCount: e.count
      };
    })});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 5. DEDUCTION BREAKDOWN (for pie chart)
// ============================================================
function dashDeductionBreakdown(filters){
  try {
    var records = _getSalaryRecords(filters || {});

    var absent = 0, halfDay = 0, advances = 0, penaltyOT = 0, penaltySalary = 0, penaltyHours = 0;

    records.forEach(function(r){
      absent += r.absentDeduction;
      halfDay += r.halfDayDeduction;
      advances += r.advanceDeducted;
      penaltyOT += r.penaltyOT;
      penaltySalary += r.penaltySalary;
      penaltyHours += r.penaltyAmount;
    });

    return _envOk({
      absentDeduction: _round(absent),
      halfDayDeduction: _round(halfDay),
      advancesDeducted: _round(advances),
      penaltyFromOT: _round(penaltyOT),
      penaltyFromSalary: _round(penaltySalary),
      penaltyFromHours: _round(penaltyHours),
      total: _round(absent + halfDay + advances + penaltyOT + penaltySalary + penaltyHours)
    });
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 6. SALARY RECORDS (for filterable table)
// ============================================================
function getSalaryRecordsFiltered(filters){
  try {
    filters = filters || {};
    var records = _getSalaryRecords(filters);

    // Search filter (partial name match)
    if (filters.search){
      var q = _s(filters.search).toLowerCase();
      records = records.filter(function(r){
        return _s(r.employee).toLowerCase().indexOf(q) >= 0
            || _s(r.month).toLowerCase().indexOf(q) >= 0
            || _s(r.id).toLowerCase().indexOf(q) >= 0;
      });
    }

    // Sort by month desc, then employee asc
    records.sort(function(a,b){
      var da = _monthLabelToDate(a.month).getTime();
      var db = _monthLabelToDate(b.month).getTime();
      if (db !== da) return db - da;
      return a.employee.localeCompare(b.employee);
    });

    // Format for display
    var formatted = records.map(function(r){
      var totalDeductions = r.absentDeduction + r.halfDayDeduction + r.penalty + r.penaltyAmount + r.advanceDeducted;
      return {
        id: r.id,
        employee: r.employee,
        month: r.month,
        fy: r.fy,
        gross: r.gross,
        deductions: _round(totalDeductions),
        ot: r.otEarnings,
        net: r.netSalary,
        mode: r.mode,
        calculatedOn: r.calculatedOn
      };
    });

    return _envOk({records: formatted, total: formatted.length});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 7. BULK IMPORT — Multiple salary records
// ============================================================
function importMultipleSalaryRecords(recordsArray){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e){ return {success:false, message:'Another import in progress — try again.'}; }

  try {
    if (!recordsArray || !recordsArray.length){
      return {success:false, message:'No records to import.'};
    }

    var ss = _getSS();
    var sheet = ss.getSheetByName(P.SALARY) || (setupPayroll() && ss.getSheetByName(P.SALARY));

    var now = new Date();
    var calcOn = Utilities.formatDate(now,'GMT+5:30','yyyy-MM-dd HH:mm:ss');
    var imported = 0;
    var errors = [];
    var successRecords = [];

    // Pre-load existing records to detect duplicates
    var existing = {};
    if (sheet.getLastRow() > 1){
      var ex = sheet.getRange(2, 1, sheet.getLastRow()-1, SC_TOTAL).getValues();
      ex.forEach(function(r, i){
        var key = _s(r[SC.EMP]) + '|' + _s(r[SC.MONTH]);
        existing[key] = i + 2; // row number
      });
    }

    recordsArray.forEach(function(data, idx){
      try {
        var emp = _s(data.employeeName);
        var month = _s(data.month);

        if (!emp){ errors.push('Row ' + (idx+1) + ': Employee required'); return; }
        if (!month){ errors.push('Row ' + (idx+1) + ': Month required'); return; }

        var gross = _n(data.grossSalary);
        var monthDays = parseInt(data.monthDays, 10) || _daysInMonth(month);
        if (monthDays <= 0) monthDays = 30;

        var present = _n(data.daysPresent);
        var absent = _n(data.daysAbsent);
        var halfDays = _n(data.halfDays);
        var otHours = _n(data.otHours);
        var otDiv = _n(data.otDivisor) || 9;
        var advDed = _n(data.advanceDeducted);
        var penOT = _n(data.penaltyFromOT);
        var penSal = _n(data.penaltyFromSalary);
        var penaltyHours = _n(data.penaltyHours) || 0;
        var penaltyAmount = _n(data.penaltyAmount) || 0;

        // Calculate (using monthly mode for imports by default)
        var empData = _findEmployee(emp);
        var empDailyRate = empData ? _n(empData.dailyRate) : 0;
        var isDailyMode = empDailyRate > 0;

        var dailyRate, otEarn, halfDed, absDed, net;

        if (isDailyMode){
          dailyRate = empDailyRate;
          otEarn = (otHours > 0 && otDiv > 0) ? (dailyRate/otDiv)*otHours : 0;
          absDed = 0;
          halfDed = 0;
          if (penaltyHours > 0 && penaltyAmount === 0) penaltyAmount = (dailyRate / 8) * penaltyHours;
          var totalPenalty = penOT + penSal + penaltyAmount;
          var presentPay = dailyRate * present;
          var halfPay = dailyRate * halfDays * 0.5;
          net = presentPay + halfPay + otEarn - advDed - totalPenalty;
          gross = _round(presentPay + halfPay);
        } else {
          dailyRate = monthDays > 0 ? gross/monthDays : 0;
          otEarn = otDiv > 0 ? (dailyRate/otDiv)*otHours : 0;
          halfDed = dailyRate * halfDays * 0.5;
          absDed = dailyRate * absent;
          if (penaltyHours > 0 && penaltyAmount === 0) penaltyAmount = (dailyRate / 8) * penaltyHours;
          var totalPenalty = penOT + penSal + penaltyAmount;
          net = gross - absDed - halfDed + otEarn - advDed - totalPenalty;
        }

        var fy = _fyFromMonthLabel(month);
        var code = emp.replace(/\s+/g,'').substring(0,6).toUpperCase();
        var recId = 'SAL-IMP-' + Utilities.formatDate(now,'GMT+5:30','yyyyMMddHHmmss') + '-' + code + '-' + idx;

        // Remove existing record for same employee+month if exists
        var dupKey = emp + '|' + month;
        if (existing[dupKey]){
          sheet.deleteRow(existing[dupKey]);
          // Adjust existing row numbers after deletion
          for (var k in existing){
            if (existing[k] > existing[dupKey]) existing[k]--;
          }
          delete existing[dupKey];
        }

        // Build row
        var row = [];
        for (var c=0; c<SC_TOTAL; c++) row.push('');
        row[SC.ID] = recId;
        row[SC.EMP] = emp;
        row[SC.MONTH] = month;
        row[SC.FY] = fy;
        row[SC.GROSS] = gross;
        row[SC.DAILY] = _round(dailyRate);
        row[SC.MONTH_DAYS] = monthDays;
        row[SC.PRESENT] = present;
        row[SC.ABSENT] = absent;
        row[SC.OT_HOURS] = otHours;
        row[SC.OT_EARN] = _round(otEarn);
        row[SC.ADV_DED] = advDed;
        row[SC.PENALTY] = _round(penOT + penSal);
        row[SC.PEN_OT] = penOT;
        row[SC.PEN_SAL] = penSal;
        row[SC.ABS_DED] = _round(absDed);
        row[SC.HALF_DAYS] = halfDays;
        row[SC.HALF_DED] = _round(halfDed);
        row[SC.PENALTY_HOURS] = _round(penaltyHours);
        row[SC.PENALTY_AMOUNT] = _round(penaltyAmount);
        row[SC.NET] = _round(net);
        row[SC.CALC_ON] = calcOn;
        row[SC.MODE] = 'Imported';

        sheet.appendRow(row);
        imported++;
        successRecords.push({employee: emp, month: month, net: _round(net)});

      } catch(recordErr){
        errors.push('Row ' + (idx+1) + ': ' + recordErr.message);
      }
    });

    SpreadsheetApp.flush();

    return {
      success: imported > 0,
      imported: imported,
      total: recordsArray.length,
      errors: errors,
      records: successRecords,
      message: 'Imported ' + imported + ' of ' + recordsArray.length + ' records' +
               (errors.length ? ' (' + errors.length + ' errors)' : '')
    };

  } catch(err){
    return {success:false, message:err.message};
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// ============================================================
// 8. Get available months for filter dropdown
// ============================================================
function getAvailableSalaryMonths(){
  try {
    var records = _getSalaryRecords({});
    var monthsSet = {};
    records.forEach(function(r){
      if (r.month) monthsSet[r.month] = _monthLabelToDate(r.month).getTime();
    });

    var months = Object.keys(monthsSet).sort(function(a,b){
      return monthsSet[b] - monthsSet[a]; // most recent first
    });

    return _envOk({months: months});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 9. Get employees for import dropdown
// ============================================================
function getEmployeesForImport(){
  try {
    var names = getEmployeeNames();
    var master = getEmployeeMasterData();
    var employees = master.map(function(e){
      return {
        name: e.name,
        grossSalary: e.grossSalary || 0,
        dailyRate: e.dailyRate || 0,
        email: e.email || ''
      };
    });
    return _envOk({employees: employees, names: names});
  } catch(e){ return _envErr(e); }
}

// ============================================================
// 10. Delete/Cancel salary record
// ============================================================
function deleteSalaryRecord(recordId){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ return {success:false, message:'Locked - try again'}; }

  try {
    var sheet = _getSS().getSheetByName(P.SALARY);
    if (!sheet || sheet.getLastRow() < 2) return {success:false, message:'No records found'};

    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, SC_TOTAL).getValues();
    for (var i = data.length - 1; i >= 0; i--){
      if (_s(data[i][SC.ID]) === recordId){
        sheet.deleteRow(i + 2);
        return {success: true, message: 'Record deleted'};
      }
    }
    return {success: false, message: 'Record not found'};

  } catch(err){
    return {success: false, message: err.message};
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// ============================================================
// Advances module moved out — all advance CRUD + readers now live in
// SalaryAdvance.gs backed by the single `Advances` sheet. The dashboard
// endpoints below (dashSalaryKpis) pull their advance numbers via
// advGetSummary(fy) from that module.
// ============================================================