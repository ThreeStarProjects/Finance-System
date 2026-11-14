/**
 * _Interlink_Masters.gs
 * -----------------------------------------------------------------------------
 * Thin wrapper around the Masters library (added as `Masters` dependency in
 * this script's Libraries). Caches reads in CacheService for 5 min so repeated
 * dropdown reads don't thrash the library.
 *
 * Paste this file verbatim into Finance / IMS / PM. It's self-contained —
 * no touches to any existing production code.
 *
 * One-time setup per system (do once after pasting this file):
 *   1. Script editor → Libraries → Add library
 *      Script ID: 1zWL2xvoZIZ1HAk-RrngKIeA4Ps4fdByYW9_hsMEm8lMPA2P8rOgTzSxC
 *      Version:   1   (or latest)
 *      Identifier: Masters
 *   2. Run IM_installInterlinkMenu() once from the editor — adds the
 *      🔗 Interlink menu to the sheet (via installable onOpen trigger, so
 *      the system's existing onOpen stays untouched).
 */

var IM_VERSION    = '0.4.0';
var IM_CACHE_TTL  = 300;   // 5 min
var IM_CACHE_KEY  = 'IM_v' + IM_VERSION + '_';

// -----------------------------------------------------------------------------
// Reads (cached)
// -----------------------------------------------------------------------------
function IM_getProjects()  { return IM_cache_('projects',  function () { return Masters.getProjects(); }); }
function IM_getCustomers() { return IM_cache_('customers', function () { return Masters.getCustomers(); }); }
/**
 * Employees are read by every payroll / form call, so keep a 6-hour fallback copy: if Masters cannot be read right now, the last
 * good list is used (payroll never stops because of a short outage). Throws only if there is no copy at all.
 */
function IM_getEmployees() {
  var cache = CacheService.getScriptCache();
  try {
    return IM_cache_('employees', function () {
      var fresh = Masters.getEmployees();
      try { var str = JSON.stringify(fresh); if (str.length < 95000) cache.put(IM_CACHE_KEY + 'employees_stale', str, 21600); } catch (_) {}
      return fresh;
    });
  } catch (e) {
    var old = cache.get(IM_CACHE_KEY + 'employees_stale');
    if (old) { try { return JSON.parse(old); } catch (_) {} }
    throw e;
  }
}
function IM_getSites()     { return IM_cache_('sites',     function () { return Masters.getSites(); }); }

/**
 * Parties with role flags. Call with a flag object to narrow:
 *   IM_getParties({ isSupplier:  true })  → suppliers only
 *   IM_getParties({ isJobWorker: true })  → job workers only
 *   IM_getParties({})                     → all parties
 */
function IM_getParties(filter) {
  var key = 'parties_' + JSON.stringify(filter || {});
  return IM_cache_(key, function () { return Masters.getParties(filter || {}); });
}

/** Freeform name → Masters record. For migration-time resolution. */
function IM_resolveName(entity, text) {
  return Masters.resolveFreeformName(entity, text);
}

// -----------------------------------------------------------------------------
// Dropdown-friendly label lists — "{ID} — {Name}"
// -----------------------------------------------------------------------------
function IM_projectLabels()  { return IM_asLabels_(IM_getProjects(),  'Project_ID',  'Name'); }
function IM_customerLabels() { return IM_asLabels_(IM_getCustomers(), 'Customer_ID', 'Name'); }
function IM_partyLabels(f)   { return IM_asLabels_(IM_getParties(f),  'Party_ID',    'Name'); }
function IM_employeeLabels() { return IM_asLabels_(IM_getEmployees(), 'Emp_ID',      'Name'); }
function IM_siteLabels()     { return IM_asLabels_(IM_getSites(),     'Site_ID',     'Name'); }

function IM_asLabels_(rows, idCol, nameCol) {
  return (rows || [])
    .filter(function (r) { return String(r['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (r) { return r[idCol] + ' — ' + r[nameCol]; });
}

// Just the IDs (for lookups where you don't want the label)
function IM_projectIds()  { return IM_getProjects().map(function (r)  { return r.Project_ID;  }); }
function IM_customerIds() { return IM_getCustomers().map(function (r) { return r.Customer_ID; }); }

// -----------------------------------------------------------------------------
// Cache plumbing
// -----------------------------------------------------------------------------
function IM_cache_(key, factory) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(IM_CACHE_KEY + key);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) {}
  }
  var fresh = factory();
  try {
    var str = JSON.stringify(fresh);
    if (str.length < 95000) cache.put(IM_CACHE_KEY + key, str, IM_CACHE_TTL);
  } catch (_) {}
  return fresh;
}

function IM_clearCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll([
    IM_CACHE_KEY + 'projects',
    IM_CACHE_KEY + 'customers',
    IM_CACHE_KEY + 'employees',
    IM_CACHE_KEY + 'sites',
    IM_CACHE_KEY + 'parties_{}',
    IM_CACHE_KEY + 'parties_{"isSupplier":true}',
    IM_CACHE_KEY + 'parties_{"isJobWorker":true}',
    IM_CACHE_KEY + 'parties_{"isTransporter":true}',
    IM_CACHE_KEY + 'parties_{"isServiceProv":true}'
  ]);
  SpreadsheetApp.getActive().toast('Masters cache cleared', '🔗 Interlink', 3);
}

// -----------------------------------------------------------------------------
// Menu — attached via installable onOpen trigger so the system's existing
// simple onOpen stays untouched. Run IM_installInterlinkMenu() ONCE.
// -----------------------------------------------------------------------------
function IM_installInterlinkMenu() {
  var ss = SpreadsheetApp.getActive();
  // De-dup: remove any existing IM_onOpen_ trigger first
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_onOpen_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_onOpen_').forSpreadsheet(ss).onOpen().create();
  // Also run once now so the menu appears without reopening
  IM_onOpen_();
  SpreadsheetApp.getUi().alert(
    'Interlink menu installed',
    'A 🔗 Interlink menu will appear every time this sheet is opened. ' +
    'The system\'s existing menu is unchanged.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function IM_onOpen_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔗 Interlink')
    .addItem('Add PM projects to Finance project list', 'IM_syncFinanceProjectsNow')
    .addItem('IMS material costs: import now', 'IM_importImsCostsNow')
    .addItem('IMS material costs: set subcategory…', 'IM_setCostSubcategory')
    .addItem('➕ Add employee',                'IM_addEmployeeDialog')
    .addItem('Refresh employees from Masters', 'IM_syncEmployeesNow')
    .addSeparator()
    .addItem('❓ Start here',                  'IM_showStartHere')
    .addItem('About',                          'IM_about_')
    .addToUi();
  try { Masters.registerSystemSheet('finance', SpreadsheetApp.getActive().getId()); } catch (e) {}  // lets PM's project delete find this sheet
  try { IM_retireEmployeeSheet_({ copyUp: false }); } catch (e) {}  // the old M_Employee sheet is removed as soon as everything in it is safely in Masters
  try { _getManHourSheet(); } catch (e) {}  // the Man_Hour_Costs sheet exists from the first open (project delete looks for it)
  try { IM_ensureImportTrigger_(); } catch (e) {}  // self-heal the 10-minute IMS cost import
  try { IM_syncFinanceProjects_(); } catch (e) {}  // every PM project is a Finance project (invoices / bills / customers can be tied to it)
}

function IM_about_() {
  var pv = '?';
  try { pv = Masters.getVersion(); } catch (e) {}
  SpreadsheetApp.getUi().alert(
    '🔗 Interlink v' + IM_VERSION,
    'This system reads master data (Projects, Customers, Parties, Employees, ' +
    'Sites) from the Masters library.\n\n' +
    'Masters version: ' + pv + '\n' +
    'Cache TTL: ' + IM_CACHE_TTL + ' sec',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// Phase 7 (v2) — IMS material costs → Project_Costs: automatic, grouped
// Finance PULLS (a time trigger every 10 min, on open, and 🔗 Interlink → import now); nothing
// else writes into this workbook. Material a project drew from IMS stock (Material Requisition lines)
// becomes a 'Material' cost.
//
//  * PM / Masters is the only truth for project names: a requisition counts when its Project BEGINS
//    with a PRJ-### id (the label picked from the IMS dropdown). The Finance project is the Masters
//    label "PRJ-001 — Name"; no pairing with IMS or Finance names is needed. The project is added to
//    M_Project (Code PRJ-001, Name) when Finance has no project with that code yet.
//  * Many partial issues do NOT create many rows: one Project_Costs row per project + item, which
//    grows (qty and amount) as more is issued. Note key "IMS:PRJ-001:ITEMCODE#1".
//  * Every requisition line is priced once and frozen in the hidden sheet IMS_Cost_Log (qty x the IMS
//    'Avg Rate (Rs.)' of that moment, GST 0 since GST was paid on the purchase bill), so a later rate
//    change never rewrites old cost and nothing is ever counted twice.
//  * A group row is only updated while its Status is Pending. Once approved / rejected (or soft
//    deleted) it is never changed; later issues of that item go to a new row "#2".
//  * Lines with no Avg Rate yet, or for a PRJ id Masters does not know, wait and are retried.
//  * Lines already booked by the old per-line import (key "IMS:REQ-0012:SR5") are skipped.
// Vendor bills tagged "Count in project cost" must not cover stock that is also issued via
// requisitions, or the same material is counted twice.
// One-time: run IM_installImportTrigger() from the editor (creates the 10-minute trigger).
// =============================================================================
var IM_IMS_SHEET_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';
var IM_LOG_SHEET    = 'IMS_Cost_Log';
var IM_LOG_HEAD     = ['Line', 'Date', 'Project', 'Item Code', 'Item', 'Qty', 'Unit', 'Rate', 'Amount', 'Group', 'Status'];
var IML = { LINE: 0, DATE: 1, PRJ: 2, CODE: 3, ITEM: 4, QTY: 5, UNIT: 6, RATE: 7, AMOUNT: 8, GROUP: 9, STATUS: 10 };
var IM_IMPORT_FN    = 'IM_importImsCostsTrigger';
var IM_IMPORT_MINUTES = 10;

/** Keys ("IMS:REQ-0012:SR5") booked line-by-line by the old import. Soft-deleted rows don't count. */
function IM_legacyKeys_() {
  var out = {};
  var sh = _getProjectCostsSheet();
  if (!sh || sh.getLastRow() < 2) return out;
  var width = Math.min(sh.getLastColumn(), PCC_TOTAL);
  sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(function (r) {
    var note = _s(r[PCC.NOTES]);
    if (note.indexOf('IMS:REQ-') === 0 && _s(r[PCC.DELETED]) !== 'Yes') out[note] = true;
  });
  return out;
}

/** IMS Item Master → { code: { name, rate } }, rate = 'Avg Rate (Rs.)'. */
function IM_avgRates_(items) {
  var out = {};
  if (!items || items.getLastRow() < 2) return out;
  var head = items.getRange(1, 1, 1, items.getLastColumn()).getValues()[0];
  var cRate = head.indexOf('Avg Rate (Rs.)'), cDesc = head.indexOf('Description'), cName = head.indexOf('Item Name');
  if (cRate < 0) return out;
  items.getRange(2, 1, items.getLastRow() - 1, items.getLastColumn()).getValues().forEach(function (r) {
    if (r[0]) out[r[0]] = { name: r[cDesc] || r[cName] || '', rate: Number(r[cRate]) || 0 };
  });
  return out;
}

function IM_dateStr_(d) {
  return d instanceof Date ? Utilities.formatDate(d, 'GMT+5:30', 'yyyy-MM-dd') : String(d || '');
}

function IM_round_(n, p) { var k = Math.pow(10, p); return Math.round(n * k) / k; }

function IM_logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(IM_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IM_LOG_SHEET);
    sh.getRange(1, 1, 1, IM_LOG_HEAD.length).setValues([IM_LOG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

/** Existing group rows in Project_Costs: { groupKey → { row, status, deleted } } and { "PRJ|CODE" → highest #n }. */
function IM_groupRows_() {
  var byKey = {}, top = {};
  var sh = _getProjectCostsSheet();
  if (!sh || sh.getLastRow() < 2) return { byKey: byKey, top: top };
  var width = Math.min(sh.getLastColumn(), PCC_TOTAL);
  sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(function (r, i) {
    var m = /^IMS:(PRJ-\d+):(.+)#(\d+)$/.exec(_s(r[PCC.NOTES]));
    if (!m) return;
    byKey[m[0]] = { row: i + 2, status: _s(r[PCC.STATUS]), deleted: _s(r[PCC.DELETED]) === 'Yes' };
    var k = m[1] + '|' + m[2], n = parseInt(m[3], 10);
    if (!top[k] || n > top[k].n) top[k] = { n: n, key: m[0] };
  });
  return { byKey: byKey, top: top };
}

/**
 * Every active Masters (PM) project becomes a Finance project, so invoices, bills, payments and customers can
 * be tied to it: M_Project (Code | Name | Active) gets Code PRJ-###, Name = the Masters name, which makes the
 * Finance label identical to the PM label "PRJ-001 — Name". Finance's own projects are never touched:
 *   - a code Finance does not have yet → a new row (inactive Masters projects are not added);
 *   - same code and same name → only deactivated if Masters has finished the project (never re-activated);
 *   - same code but another name → left alone, reported as a warning.
 * When rows were added the sheet dropdowns are refreshed (applyProjectDropdowns); the web forms read the list live.
 * Returns { added, deactivated, warnings[] }.
 */
function IM_syncFinanceProjects_() {
  var out = { added: 0, deactivated: 0, warnings: [] };
  var sh = _getSS().getSheetByName(P.PROJECT);
  if (!sh) return out;
  var rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, PRC_TOTAL).getValues() : [];
  var byCode = {};
  rows.forEach(function (r, i) { var c = _s(r[PRC.CODE]).toUpperCase(); if (c) byCode[c] = { i: i, row: r }; });
  var add = [];
  (Masters.getProjects() || []).forEach(function (p) {
    var id = String(p.Project_ID).toUpperCase(), name = String(p.Name), active = String(p['Active?'] || 'Yes').toLowerCase() !== 'no';
    var ex = byCode[id];
    if (!ex) {
      if (!active) return;
      var row = new Array(PRC_TOTAL).fill('');
      row[PRC.CODE] = id; row[PRC.NAME] = name; row[PRC.ACTIVE] = 'Yes';
      add.push(row);
    } else if (_s(ex.row[PRC.NAME]) !== name) {
      out.warnings.push('Finance already uses code ' + id + ' for "' + _s(ex.row[PRC.NAME]) + '" (Masters: "' + name + '"): left as it is.');
    } else if (!active && _s(ex.row[PRC.ACTIVE]) !== 'No') {
      sh.getRange(ex.i + 2, PRC.ACTIVE + 1).setValue('No');
      out.deactivated++;
    }
  });
  if (add.length) {
    sh.getRange(sh.getLastRow() + 1, 1, add.length, PRC_TOTAL).setValues(add);
    out.added = add.length;
  }
  if (out.added || out.deactivated) { try { applyProjectDropdowns(); } catch (e) { Logger.log('project dropdowns: ' + e.message); } }
  return out;
}

/** Menu: add the PM projects to Finance's project list now. */
function IM_syncFinanceProjectsNow() {
  var r = IM_syncFinanceProjects_();
  SpreadsheetApp.getActive().toast(
    (r.added ? r.added + ' project(s) added to the Finance project list.' : 'Finance project list already has every PM project.') +
    (r.deactivated ? ' ' + r.deactivated + ' deactivated.' : '') + (r.warnings.length ? ' ' + r.warnings.join(' ') : ''), '🔗 Interlink', 8);
}

/** The import. Safe to run any time, as often as you like: it only ever adds what is new. */
function IM_importImsCosts() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { busy: true };
  try { return IM_importImsCostsLocked_(); } finally { lock.releaseLock(); }
}

function IM_importImsCostsLocked_() {
  var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID);
  var req = ims.getSheetByName('Material Requisition Register'), items = ims.getSheetByName('Item Master');
  if (!req || !items) throw new Error('IMS sheets not found (Material Requisition Register / Item Master).');
  var res = { newLines: 0, created: 0, updated: 0, closed: 0, waitingRate: 0, waitingProject: 0, warnings: [],
              imsLines: 0, prjLines: 0, ignored: {} };  // diagnostics: what IMS holds, and which project texts were not PRJ-###

  var names = {};  // PRJ-### → Masters name
  (Masters.getProjects() || []).forEach(function (p) { names[String(p.Project_ID).toUpperCase()] = String(p.Name); });
  var rates = IM_avgRates_(items), legacy = IM_legacyKeys_();

  // 1. the PRJ requisition lines IMS has right now
  var lines = [], present = {};
  if (req.getLastRow() >= 2) req.getRange(2, 1, req.getLastRow() - 1, 8).getValues().forEach(function (r) {
    var reqNo = r[1], code = r[4], qty = Number(r[6]), proj = String(r[3] || '').trim(), m = /^PRJ-\d+/i.exec(proj);
    if (!reqNo || !code || !(qty > 0)) return;
    res.imsLines++;
    if (!m) { var nm = proj || '(blank)'; res.ignored[nm] = (res.ignored[nm] || 0) + 1; return; }
    res.prjLines++;
    var key = reqNo + ':SR' + r[0];
    present[key] = true;
    lines.push({ key: key, date: IM_dateStr_(r[2]), prj: m[0].toUpperCase(), code: code, qty: qty, unit: r[7],
                 item: (rates[code] && rates[code].name) || r[5] || code });
  });

  // 2. price + group the lines not seen before
  var log = IM_logSheet_();
  var logRows = log.getLastRow() >= 2 ? log.getRange(2, 1, log.getLastRow() - 1, IM_LOG_HEAD.length).getValues() : [];
  var seen = {}; logRows.forEach(function (r) { seen[r[IML.LINE]] = true; });
  var groups = IM_groupRows_();
  var open = {};  // "PRJ|CODE" → { key, n } the group new lines go to
  Object.keys(groups.top).forEach(function (k) {
    var g = groups.byKey[groups.top[k].key];
    if (g && !g.deleted && g.status === 'Pending') open[k] = groups.top[k];
  });
  var newRows = [];
  lines.forEach(function (l) {
    if (seen[l.key]) return;
    if (legacy['IMS:' + l.key]) { newRows.push([l.key, l.date, l.prj, l.code, l.item, l.qty, l.unit, '', '', '', 'Legacy']); return; }
    if (!names[l.prj]) { res.waitingProject++; return; }
    var rate = rates[l.code] ? rates[l.code].rate : 0;
    if (!(rate > 0)) { res.waitingRate++; return; }
    var k = l.prj + '|' + l.code;
    if (!open[k]) {
      var n = groups.top[k] ? groups.top[k].n + 1 : 1;
      open[k] = { n: n, key: 'IMS:' + l.prj + ':' + l.code + '#' + n };
      groups.top[k] = open[k];
    }
    newRows.push([l.key, l.date, l.prj, l.code, l.item, l.qty, l.unit, rate, IM_round_(l.qty * rate, 4), open[k].key, 'Imported']);   // 4 decimals per line, the group total is rounded to paise
  });
  if (newRows.length) {
    log.getRange(log.getLastRow() + 1, 1, newRows.length, IM_LOG_HEAD.length).setValues(newRows);
    res.newLines = newRows.filter(function (r) { return r[IML.STATUS] === 'Imported'; }).length;
  }

  // 3. every group row = the sum of its still-present logged lines
  var all = logRows.concat(newRows), sums = {};
  all.forEach(function (r) {
    if (r[IML.STATUS] !== 'Imported') return;
    var g = sums[r[IML.GROUP]] = sums[r[IML.GROUP]] || { key: r[IML.GROUP], prj: r[IML.PRJ], code: r[IML.CODE], item: r[IML.ITEM], unit: r[IML.UNIT], qty: 0, amount: 0, date: '', lines: 0 };
    if (!present[r[IML.LINE]]) return;                       // removed from IMS since: no longer counts
    g.qty += Number(r[IML.QTY]); g.amount += Number(r[IML.AMOUNT]); g.lines++;
    if (String(r[IML.DATE]) > g.date) g.date = String(r[IML.DATE]);
  });
  var pc = _getProjectCostsSheet(), entries = [];
  var sub = PropertiesService.getScriptProperties().getProperty('IM_LAST_SUBCAT') || '';
  Object.keys(sums).forEach(function (key) {
    var g = sums[key], ex = groups.byKey[key];
    g.qty = IM_round_(g.qty, 3); g.amount = IM_round_(g.amount, 2);
    if (ex && (ex.deleted || ex.status !== 'Pending')) return;        // decided by a person: never touched
    var desc = 'Issued from IMS stock · ' + g.lines + ' requisition line(s)';
    if (ex) {
      if (g.lines === 0) {                                              // every line was removed from IMS
        pc.getRange(ex.row, PCC.DELETED + 1).setValue('Yes'); res.closed++; return;
      }
      var cur = pc.getRange(ex.row, 1, 1, PCC_TOTAL).getValues()[0];
      if (Math.abs(Number(cur[PCC.QUANTITY]) - g.qty) < 0.0005 && Math.abs(Number(cur[PCC.AMOUNT]) - g.amount) < 0.005) return;  // unchanged
      pc.getRange(ex.row, PCC.ENTRY_DATE + 1).setValue(g.date);
      pc.getRange(ex.row, PCC.DESCRIPTION + 1).setValue(desc);
      pc.getRange(ex.row, PCC.QUANTITY + 1).setValue(g.qty);
      pc.getRange(ex.row, PCC.RATE_PER_UNIT + 1).setValue(g.qty ? g.amount / g.qty : 0);
      pc.getRange(ex.row, PCC.AMOUNT_EX_GST + 1).setValue(g.amount);
      pc.getRange(ex.row, PCC.AMOUNT + 1).setValue(g.amount);
      res.updated++;
    } else if (g.lines > 0) {
      entries.push({ entry_type: 'Material', project: g.prj + ' — ' + names[g.prj], entry_date: g.date, submitted_by: 'IMS import',
        data: { materialName: String(g.item), quantity: g.qty, unit: String(g.unit), rate: g.amount / g.qty, gst_rate: 0,
                subcategory: sub, description: desc, notes: g.key } });
    }
  });
  if (entries.length) {
    var out = submitProjectCosts(entries);
    if (!out || !out.success) throw new Error('Finance did not accept the cost rows: ' + ((out && out.message) || 'error') + ' (nothing is lost: they are retried on the next run)');
    res.created = out.count;
  }
  return res;
}

/** Trigger / menu entry points + heartbeat (log sheet M1: if it is stale, the trigger is not running). */
function IM_importImsCostsTrigger() { IM_runImport_(true); }
function IM_importImsCostsNow() { IM_runImport_(false); }

function IM_runImport_(quiet) {
  var res, proj = { added: 0, deactivated: 0, warnings: [] };
  try { IM_retireEmployeeSheet_({ copyUp: true }); } catch (e) { Logger.log('Employee move to Masters failed: ' + e.message); }   // the 10-minute job also finishes moving pay terms into Masters
  try { proj = IM_syncFinanceProjects_(); }                    // first: the PM projects must be Finance projects even if IMS is unreachable
  catch (e) { proj.warnings.push('Project list not synced: ' + e.message); }
  try { res = IM_importImsCosts(); }
  catch (e) {
    if (!quiet) SpreadsheetApp.getUi().alert('IMS material costs', 'Import failed: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    else Logger.log('IMS cost import failed: ' + e.message);
    return null;
  }
  res.warnings = (res.warnings || []).concat(proj.warnings);
  res.projectsAdded = proj.added;
  if (res.busy) { if (!quiet) SpreadsheetApp.getActive().toast('Another import is running; try again in a minute.', 'IMS material costs', 6); return res; }
  var n = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === IM_IMPORT_FN; }).length;
  try {
    IM_logSheet_().getRange('M1').setValue('Last import ' + Utilities.formatDate(new Date(), 'GMT+5:30', 'dd/MM HH:mm') +
      ' · +' + res.newLines + ' line(s), ' + res.created + ' row(s) created, ' + res.updated + ' updated · triggers: ' + n);
  } catch (e) {}
  if (!quiet) {
    var ign = Object.keys(res.ignored).sort(function (a, b) { return res.ignored[b] - res.ignored[a]; });
    var msg = 'IMS has ' + res.imsLines + ' requisition line(s); ' + res.prjLines + ' begin with PRJ-###.\n' +
      (ign.length ? 'Not imported (Project does not begin with PRJ-###): ' + ign.slice(0, 5).map(function (n) { return n + ' ×' + res.ignored[n]; }).join(', ') + (ign.length > 5 ? ', …' : '') + '\n' : '') +
      (proj.added ? '\n' + proj.added + ' PM project(s) added to the Finance project list.' : '') +
      '\n' + res.newLines + ' new requisition line(s) imported.\n' + res.created + ' cost row(s) created, ' + res.updated + ' updated' +
      (res.closed ? ', ' + res.closed + ' closed (lines removed in IMS)' : '') + '.' +
      (res.waitingRate ? '\n' + res.waitingRate + ' line(s) wait for an Avg Rate in IMS Item Master.' : '') +
      (res.waitingProject ? '\n' + res.waitingProject + ' line(s) wait: PRJ id not found in Masters.' : '') +
      (res.warnings.length ? '\n\nCheck: ' + res.warnings.join('\n') : '') +
      '\n\nFinance also does this by itself every 10 minutes.';
    SpreadsheetApp.getUi().alert('IMS material costs', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  return res;
}

/** Create the 10-minute trigger (once, from the editor; it also authorises the script). */
function IM_installImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === IM_IMPORT_FN) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(IM_IMPORT_FN).timeBased().everyMinutes(IM_IMPORT_MINUTES).create();
  var res = IM_runImport_(true);
  var msg = 'IMS cost import trigger created (every ' + IM_IMPORT_MINUTES + ' minutes). First run: ' +
            (res && !res.busy ? res.newLines + ' line(s), ' + res.created + ' row(s) created.' : 'will run at the next tick.');
  try { SpreadsheetApp.getUi().alert('IMS material costs', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(msg); }
}

/** Recreate the trigger if it has gone missing. Called when Finance opens. */
function IM_ensureImportTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === IM_IMPORT_FN; });
  if (!has) ScriptApp.newTrigger(IM_IMPORT_FN).timeBased().everyMinutes(IM_IMPORT_MINUTES).create();
}

// -----------------------------------------------------------------------------
// Tidy-ups after the move to the automatic import.
// -----------------------------------------------------------------------------

/**
 * Menu: choose the Finance category (subcategory) the imported material costs are filed under.
 * Also fills it on the imported rows that are still Pending and have none.
 */
function IM_setCostSubcategory() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  var cats = [];
  try { cats = getCategories() || []; } catch (e) {}
  var names = cats.map(function (c) { return c.category; }).filter(String);
  var cur = props.getProperty('IM_LAST_SUBCAT') || '(none)';
  var r = ui.prompt('IMS material costs: subcategory',
    'Current: ' + cur + '\n\nType the Finance category for imported material costs, exactly as in M_Category.\n' +
    'Some of them: ' + names.slice(0, 15).join(', ') + (names.length > 15 ? ', …' : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var text = String(r.getResponseText() || '').trim().toLowerCase();
  var hit = cats.filter(function (c) { return String(c.category).toLowerCase() === text; })[0];
  if (!hit) { ui.alert('Subcategory', '"' + r.getResponseText() + '" is not in M_Category. Nothing changed.', ui.ButtonSet.OK); return; }
  props.setProperty('IM_LAST_SUBCAT', hit.category);

  var filled = 0, pc = _getProjectCostsSheet();
  if (pc && pc.getLastRow() >= 2) {
    pc.getRange(2, 1, pc.getLastRow() - 1, Math.min(pc.getLastColumn(), PCC_TOTAL)).getValues().forEach(function (row, i) {
      if (/^IMS:PRJ-/.test(_s(row[PCC.NOTES])) && _s(row[PCC.STATUS]) === 'Pending' && _s(row[PCC.DELETED]) !== 'Yes' && !_s(row[PCC.SUBCATEGORY])) {
        pc.getRange(i + 2, PCC.SUBCATEGORY + 1, 1, 2).setValues([[hit.category, hit.group || '']]);   // Subcategory, Category (group)
        filled++;
      }
    });
  }
  ui.alert('Subcategory', 'Imported material costs are now filed under "' + hit.category + '".' +
    (filled ? '\n' + filled + ' existing Pending row(s) without a category were updated.' : ''), ui.ButtonSet.OK);
}

/** Menu: ❓ Start here. Creates the tab the first time, then just opens it. */
function IM_showStartHere() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('Start here');
  if (!sh) {
    sh = ss.insertSheet('Start here', 0);
    var rows = [
      ['Start here: Finance', 'title'],
      ['How material costs from IMS reach Finance', 'h'],
      ["1.  Every PM project appears in Finance's project list within a few minutes, named like PRJ-001 — Project name. Pick it on invoices, bills and payments to tie them to the project.", ''],
      ['2.  Material issued by the IMS store to a project (a Material Requisition whose project starts with PRJ-…) is imported into Project_Costs by itself every 10 minutes. You can also use 🔗 Interlink → IMS material costs: import now.', ''],
      ['3.  There is one cost row per project and item, and it grows as more is issued. The row stays Pending until you approve it. Once it is approved or rejected it never changes; later issues of that item start a second row.', ''],
      ['4.  Costs are quantity × the IMS Avg Rate at the time of issue (no GST: GST was paid on the purchase bill). If IMS has no Avg Rate for an item, those lines wait until it has one.', ''],
      ['5.  The category of these rows is set in 🔗 Interlink → IMS material costs: set subcategory…', ''],
      ['Take care', 'h'],
      ['•  A vendor bill ticked "Count in project cost" must not cover stock that is also issued from IMS, or the same material is counted twice.', ''],
      ['If something looks wrong', 'h'],
      ['•  Open this spreadsheet: it restarts the 10-minute import if it has stopped.', ''],
      ['•  In PM, 🔗 Interlink → 📋 Needs attention lists requisition lines that Finance cannot cost, and 🩺 Health check shows whether this import is running.', '']
    ];
    sh.getRange(1, 1, rows.length, 1).setValues(rows.map(function (r) { return [r[0]]; })).setWrap(true).setVerticalAlignment('top');
    sh.setColumnWidth(1, 900);
    rows.forEach(function (r, i) {
      if (r[1] === 'title') sh.getRange(i + 1, 1).setFontSize(16).setFontWeight('bold');
      if (r[1] === 'h') sh.getRange(i + 1, 1).setFontWeight('bold').setBackground('#DDEBF7');
    });
  }
  ss.setActiveSheet(sh);
}

// -----------------------------------------------------------------------------
// ➕ Add employee. The person is created in Masters (the one official list) and this system's own list is updated at
// once. Masters editors only (every Masters write is). Masters refuses an exact duplicate and warns on a similar name.
// -----------------------------------------------------------------------------
function IM_addEmployeeDialog() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(IM_addEmployeeHtml_()).setWidth(430).setHeight(500), '➕ Add employee');
}
function IM_addEmployeeHtml_() { return "<div style=\"font:14px Arial,sans-serif;padding:14px;color:#222\">\n<p style=\"margin:0 0 12px;color:#555\">The person is saved in <b>Masters</b>, the official employee list. The other systems pick the name up from there.</p>\n<label>Full name *</label><input id=\"n\" maxlength=\"80\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Role (optional)</label><input id=\"r\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Department (optional)</label><input id=\"d\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Phone (optional)</label><input id=\"p\" maxlength=\"30\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<div id=\"m\" style=\"min-height:42px;margin:6px 0;font-size:13px\"></div>\n<button id=\"go\" onclick=\"send(false)\" style=\"padding:8px 16px;background:#1F4E79;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add employee</button>\n<button id=\"any\" onclick=\"send(true)\" style=\"display:none;padding:8px 16px;background:#b45309;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add anyway (different person)</button>\n<button onclick=\"google.script.host.close()\" style=\"padding:8px 16px;margin-left:6px;border:1px solid #bbb;background:#fff;border-radius:4px;cursor:pointer\">Close</button>\n</div>\n<script>\nfunction $(i){return document.getElementById(i);}\nfunction say(t,c){var m=$('m');m.textContent=t;m.style.color=c||'#222';}\nfunction busy(b){$('go').disabled=b;$('any').disabled=b;}\nfunction send(confirm){\n  var name=$('n').value.replace(/\\s+/g,' ').trim();\n  if(!name){say('Please type the name.','#b91c1c');return;}\n  busy(true);say('Saving\u2026');\n  google.script.run.withSuccessHandler(done).withFailureHandler(function(e){busy(false);say('Could not save: '+(e&&e.message||e),'#b91c1c');})\n    .IM_addEmployee({name:name,role:$('r').value,department:$('d').value,phone:$('p').value,confirmSimilar:confirm===true});\n}\nfunction done(r){\n  busy(false);$('any').style.display='none';\n  if(r&&r.ok){say('Added '+r.name+' ('+r.id+'). '+(r.note||''),'#166534');$('go').style.display='none';$('n').disabled=$('r').disabled=$('d').disabled=$('p').disabled=true;}\n  else if(r&&r.duplicate){say(r.duplicate.name+' is already in Masters ('+r.duplicate.id+', '+r.duplicate.status+'). Nothing was added.','#b45309');}\n  else if(r&&r.similar){say('Similar name(s) already in Masters: '+r.similar.map(function(s){return s.name+' ('+s.id+', '+s.status+')';}).join('; ')+'. If this is the same person, close this window. If it is a different person, click Add anyway.','#b45309');$('any').style.display='inline-block';}\n  else{say((r&&r.error)||'Something went wrong.','#b91c1c');}\n}\n</script>"; }
/** Called by the dialog. Returns { ok, id, name, note } | { ok:false, duplicate | similar | error }. Never throws. */
function IM_addEmployee(d) {
  try {
    var r = Masters.addEmployeeChecked({ Name: d.name, Role: d.role, Department: d.department, Phone: d.phone, confirmSimilar: d.confirmSimilar === true });
    if (r.ok) { try { IM_clearCache(); } catch (e) {} r.note = IM_addEmployeePost_(r, d); }
    return r;
  } catch (e) {
    var m = String((e && e.message) || e);
    return { ok: false, error: /Editors only|Admin only|permission|access/i.test(m)
      ? 'Only Masters editors can add employees. Ask the administrator to add you as an editor, or ask them to add this person.' : m };
  }
}
/** Finance: the person is in Masters, which every Finance screen reads: nothing else to copy. Pay terms are edited in the Masters tab. */
function IM_addEmployeePost_(r, d) {
  return 'Available in every Finance screen now. Next: enter the pay terms (Finance → Masters tab, Edit): payroll, advances and man-hours cannot use this person until a daily or monthly salary is there.';
}

// -----------------------------------------------------------------------------
// Employees: Masters' Employees sheet is the ONLY source of employee names and pay terms. Every Finance screen reads it
// through getEmployeeMasterData() (Salary.gs; Masters list cached 5 min, plus a 6-hour fallback copy if Masters is briefly
// unreachable). The old local sheet M_Employee is retired by IM_retireEmployeeSheet_ below, but ONLY once everything in it is
// safely in Masters; until then it still fills blanks, so nothing disappears during the move.
// Names compared by IM_normName_ (case, punctuation, spaces only, never fuzzy).
// -----------------------------------------------------------------------------
function IM_normName_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

/** Plain-language text for a failed save to Masters. */
function IM_employeeSaveError_(e) {
  var m = String((e && e.message) || e);
  return /Editors only|Admin only|permission|access/i.test(m)
    ? 'Only Masters editors can add or change employees. Ask the administrator to add you as an editor.'
    : 'Could not save to Masters: ' + m;
}

/**
 * Save one employee's details to Masters (used by the Finance Masters tab). orig = the name as it is now.
 * e = { name, email, phone, designation, department, grossSalary, dailyRate, holdsFloat, active }.
 * Returns { ok } or { ok:false, message }. Nothing is changed in Masters if it refuses.
 */
function IM_pushEmployeeToMasters_(orig, e) {
  try {
    try { IM_clearCache(); } catch (x) {}
    var list = IM_getEmployees() || [], ko = IM_normName_(orig), kn = IM_normName_(e.name), mine = null, clash = null;
    list.forEach(function (m) { if (!mine && IM_normName_(m.Name) === ko) mine = m; });
    list.forEach(function (m) { if (!clash && IM_normName_(m.Name) === kn && (!mine || m.Emp_ID !== mine.Emp_ID)) clash = m; });
    if (clash) return { ok: false, message: 'Another employee in Masters already has that name (' + clash.Emp_ID + ').' };
    var today = Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd');
    var fields = { Name: e.name, Role: e.designation, Phone: e.phone, Email: e.email, Department: e.department,
      Gross_Monthly_Salary: e.grossSalary, Daily_Rate: e.dailyRate, 'Holds_Float?': e.holdsFloat ? 'Yes' : 'No', Status: e.active ? 'Active' : 'Left' };
    fields.Left_Date = e.active ? '' : ((mine && mine.Left_Date) || today);
    if (mine) Masters.updateEmployee(mine.Emp_ID, fields);
    else {
      // a Finance-only person: adopt into Masters, then apply the rest
      var r = Masters.addEmployeeChecked({ Name: e.name, Role: e.designation, Phone: e.phone, Email: e.email, Department: e.department,
        Gross_Monthly_Salary: e.grossSalary || '', Daily_Rate: e.dailyRate || '', 'Holds_Float?': e.holdsFloat ? 'Yes' : 'No', confirmSimilar: true });
      if (!r.ok) return { ok: false, message: 'Employee already exists in Masters' + (r.duplicate ? ' (' + r.duplicate.id + ').' : '.') };
      if (!e.active) Masters.updateEmployee(r.id, { Status: 'Left', Left_Date: today });
    }
    try { IM_clearCache(); } catch (x) {}
    return { ok: true };
  } catch (x) { return { ok: false, message: IM_employeeSaveError_(x) }; }
}

var IM_EMP_COPYUP_MAX = 15;   // Masters writes per run (each is about a second); the rest follows on the next run

/**
 * Retire the old local employee sheet M_Employee.
 *   1. (copyUp) whatever Finance has that Masters lacks (designation, phone, email, department, salary, daily rate) is copied UP
 *      into Masters, only into blank fields, 15 people per run. Needs a Masters editor; otherwise it is simply skipped.
 *   2. Verify: every ACTIVE person in M_Employee is in Masters, and none of their details is still missing in Masters.
 *   3. Only if that is true is the sheet DELETED. Otherwise it stays (and still fills blanks for the readers) and the result says why.
 * Returns { state: 'gone' | 'removed' | 'blocked', n, notInMasters:[], pending, copiedUp, copyFailed, moreToCopy, busy }.
 */
function IM_retireEmployeeSheet_(opts) {
  opts = opts || {};
  var ss = _getSS(), sh = ss.getSheetByName(P.EMPLOYEE);
  if (!sh) return { state: 'gone' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { state: 'blocked', busy: true, notInMasters: [], pending: 0, copiedUp: 0, copyFailed: 0, moreToCopy: 0 };
  try {
    var res = { state: 'blocked', n: 0, notInMasters: [], pending: 0, copiedUp: 0, copyFailed: 0, moreToCopy: 0 };
    if (opts.copyUp) { try { Masters.ensureEmployeeLayout(); } catch (x) {} }
    var last = sh.getLastRow(), cols = Math.min(Math.max(sh.getLastColumn(), 1), EMC_TOTAL);
    var local = last >= 2 ? sh.getRange(2, 1, last - 1, cols).getValues().map(function (r) { while (r.length < EMC_TOTAL) r.push(''); return r; }) : [];
    local = local.filter(function (r) { return String(r[EMC.NAME]).trim() !== ''; });
    res.n = local.length;
    var txt = function (v) { return String(v == null ? '' : v).trim(); };
    var num = function (v) { var x = Number(v); return isNaN(x) ? 0 : x; };
    var load = function () { var by = {}; (IM_getEmployees() || []).forEach(function (e) { var k = IM_normName_(e.Name); if (k && !by[k]) by[k] = e; }); return by; };
    var by = load();
    // what Finance has that Masters lacks, per person
    var gaps = function (r, e) {
      var up = {};
      if (txt(e.Role) === '' && txt(r[EMC.DESIG]) !== '') up.Role = txt(r[EMC.DESIG]);
      if (txt(e.Phone) === '' && txt(r[EMC.PHONE]) !== '') up.Phone = txt(r[EMC.PHONE]);
      if (txt(e.Email) === '' && txt(r[EMC.EMAIL]) !== '') up.Email = txt(r[EMC.EMAIL]);
      if (txt(e.Department) === '' && txt(r[EMC.DEPT]) !== '') up.Department = txt(r[EMC.DEPT]);
      if (txt(e.Gross_Monthly_Salary) === '' && num(r[EMC.GROSS]) > 0) up.Gross_Monthly_Salary = num(r[EMC.GROSS]);
      if (txt(e.Daily_Rate) === '' && num(r[EMC.DAILY_RATE]) > 0) up.Daily_Rate = num(r[EMC.DAILY_RATE]);
      return up;
    };
    if (opts.copyUp) {
      var batch = [];
      local.forEach(function (r) {
        var e = by[IM_normName_(r[EMC.NAME])]; if (!e) return;
        var up = gaps(r, e);
        if (Object.keys(up).length) { if (batch.length < IM_EMP_COPYUP_MAX) batch.push(Object.assign({ Name: e.Name }, up)); else res.moreToCopy++; }
      });
      if (batch.length) {
        try { var f = Masters.fillEmployeeBlanks(batch); res.copiedUp = f.filled; }
        catch (x) { res.copyFailed = batch.length; }   // not a Masters editor: nothing lost, the sheet just stays
        try { IM_clearCache(); } catch (x) {}
        by = load();
      }
    }
    local.forEach(function (r) {
      var e = by[IM_normName_(r[EMC.NAME])], active = txt(r[EMC.ACTIVE]) !== 'No';
      if (!e) { if (active) res.notInMasters.push(txt(r[EMC.NAME])); return; }
      if (Object.keys(gaps(r, e)).length) res.pending++;
    });
    if (!res.notInMasters.length && !res.pending && !res.moreToCopy) {
      ss.deleteSheet(sh);
      res.state = 'removed';
    }
    return res;
  } finally { lock.releaseLock(); }
}

/** Menu: 🔗 Refresh employees from Masters. Clears the cached list, finishes the move of pay terms into Masters, retires the old sheet. */
function IM_syncEmployeesNow() {
  var ss = SpreadsheetApp.getActive();
  try {
    try { IM_clearCache(); } catch (e) {}
    var count = (IM_getEmployees() || []).filter(function (e) { return String(e.Status) !== 'Left' && String(e.Name || '').trim() !== ''; }).length;
    var r = IM_retireEmployeeSheet_({ copyUp: true }), msg;
    if (r.state === 'gone') msg = 'Employees come only from Masters: ' + count + ' active employee(s) are available in every Finance screen.';
    else if (r.state === 'removed') msg = 'Done: all ' + r.n + ' people and their pay terms are in Masters' + (r.copiedUp ? ' (pay details of ' + r.copiedUp + ' copied up just now)' : '') + '. The old M_Employee sheet was removed. ' + count + ' active employee(s) now come only from Masters.';
    else {
      var why = [];
      if (r.busy) why.push('another update is running, try again in a minute');
      if (r.notInMasters.length) why.push(r.notInMasters.length + ' active person(s) are not in Masters (' + r.notInMasters.slice(0, 5).join(', ') + (r.notInMasters.length > 5 ? ', …' : '') + '): add them with ➕ Add employee or make them inactive');
      if (r.copyFailed) why.push('only Masters editors can copy pay details into Masters (' + r.copyFailed + ' waiting)');
      else if (r.pending || r.moreToCopy) why.push((r.pending + r.moreToCopy) + ' person(s) still have pay details to copy' + (r.moreToCopy ? ' (the rest on the next run)' : ''));
      msg = 'The old M_Employee sheet is kept for now: ' + why.join('; ') + '.';
    }
    ss.toast(msg, '🔗 Interlink', 15);
  } catch (e) { ss.toast('Could not read Masters: ' + e.message, '🔗 Interlink', 10); }
}
