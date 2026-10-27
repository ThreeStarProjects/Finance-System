// ============================================================
// 3 Star Projects — Unified Finance System
// PHASE 6 : WEB APP ROUTER
// File: Router.gs   (depends on all earlier files)
//
// Web-app pages (?page=...):
//   home       → Home (single-page app: Dashboard, Transactions,
//                Sales & AR, Salary, Masters, Tools — all in tabs)
//   expense    → ExpenseForm        (employee expense entry — separate)
//   owner      → OwnerTransaction   (owner quick entry — separate)
// ============================================================

function doGet(e){
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  var map = {
    'home':    ['Home',             'Finance'],
    'expense': ['ExpenseForm',      'Submit Expense'],
    'owner':   ['OwnerTransaction', 'New Transaction']
  };
  if (!map[page]) page = 'home';
  return HtmlService.createTemplateFromFile(map[page][0])
    .evaluate()
    .setTitle(map[page][1] + ' — Three Star Projects')
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Inline a shared partial:  <?!= include('_styles') ?>
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// The deployed web-app URL — used by the client to build absolute nav links
// that work from inside the sandboxed iframe (relative links break there).
function getAppUrl(){
  try { return ScriptApp.getService().getUrl() || ''; } catch(e){ return ''; }
}

// The underlying spreadsheet URL — for the shell's "Open Sheet" link.
function getSheetUrl(){
  try { return _getSS().getUrl() || ''; } catch(e){ return ''; }
}

// ── One round-trip config for every form ────────────────────
// Pulls all the lists a form might need so a page loads with one call.
function getBootstrap(){
  return {
    categories:        getCategories(),                  // [{category, group, direction}]
    employeeCategories:getEmployeeCategories(),          // [{category, group}] — staff short list
    personalCategories:getPersonalCategories(),          // [string]
    pnlGroups:         getPnlGroups(),                   // [string]
    customers:         getCustomers(),                   // [{id,name,gstin}]
    vendors:           getVendors(),                     // [{id,name,gstin}]
    employees:         getEmployeeNames(),               // [string]
    employeeSalaryMap: getEmployeeSalaryMap(),           // {name: gross}
    employeeDailyRateMap: getEmployeeDailyRateMap(),     // {name: dailyRate} — > 0 means daily-mode
    projects:          getProjectNames(),                // [string]
    paymentModes:      getPaymentModes(),                // [string]
    sourcesOfFunds:    getSourcesOfFunds(),              // [{value, attribute}]
    accounts:          getAccounts(),                    // [{id,name,owner,bank,side}]
    currentMonth:      _currentMonth(),
    currentFY:         _currentFY()
  };
}